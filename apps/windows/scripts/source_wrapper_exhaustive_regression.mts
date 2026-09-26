import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import {
  canonicalSourceWrapperCommand,
  sourceWrapperKind,
  sourceWrapperKinds,
} from "../src/editor/sourceWrapperRegistry.ts";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const commands = [...sourceWrapperKinds.keys()];
const offset = process.pid % 900;
const previewPort = 17600 + offset;
const debugPort = 22600 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const profile = createBrowserProfilePath("visualtex-source-wrapper-exhaustive");
const browserPath = resolveChromiumExecutable();
const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url: string, timeoutMs = 15_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return response;
    } catch {
      // Local preview/browser is still starting.
    }
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

class CdpClient {
  private nextId = 1;
  private pending = new Map<
    number,
    { resolve: (value: any) => void; reject: (reason: Error) => void }
  >();
  private socket!: WebSocket;

  constructor(private readonly url: string) {}

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise<void>((resolve, reject) => {
      this.socket.addEventListener("open", () => resolve(), { once: true });
      this.socket.addEventListener("error", () => reject(new Error("CDP open failed")), {
        once: true,
      });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(String(event.data));
      if (!message.id) return;
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }

  send(method: string, params: Record<string, unknown> = {}) {
    const id = this.nextId++;
    return new Promise<any>((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }

  close() {
    this.socket?.close();
  }
}

const preview = spawn(
  process.execPath,
  [
    "node_modules/vite/bin/vite.js",
    "preview",
    "--host",
    "127.0.0.1",
    "--port",
    String(previewPort),
    "--strictPort",
  ],
  { cwd: process.cwd(), stdio: "ignore" },
);
let browser: ReturnType<typeof spawn> | undefined;
let client: CdpClient | undefined;

try {
  await waitFor(baseUrl);
  browser = spawn(
    browserPath,
    [
      "--headless=new",
      "--disable-gpu",
      "--no-first-run",
      "--no-default-browser-check",
      `--remote-debugging-port=${debugPort}`,
      `--user-data-dir=${profile}`,
      "--window-size=1400,1000",
      baseUrl,
    ],
    { stdio: "ignore" },
  );

  await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
  const targets = await (
    await fetch(`http://127.0.0.1:${debugPort}/json/list`)
  ).json() as Array<{ type: string; url: string; webSocketDebuggerUrl: string }>;
  const page =
    targets.find((target) => target.type === "page" && target.url.startsWith(baseUrl)) ??
    targets.find((target) => target.type === "page");
  assert.ok(page, "No VisualTeX browser target found");

  client = new CdpClient(page.webSocketDebuggerUrl);
  await client.connect();
  await client.send("Runtime.enable");
  await client.send("Page.enable");
  await client.send("Page.navigate", { url: baseUrl });
  await sleep(1800);

  const evaluate = async (expression: string) => {
    const result = await client!.send("Runtime.evaluate", {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (result.exceptionDetails) {
      throw new Error(
        result.exceptionDetails.exception?.description ??
          result.exceptionDetails.text ??
          "Runtime.evaluate failed",
      );
    }
    return result.result.value;
  };

  await evaluate(`(() => {
    localStorage.setItem("visualtex.onboarding.v3.completed", "true");
    localStorage.setItem("visualtex.onboarding.windows.desktop.v1.1.0.completed", "true");
    const key = "visualtex-editor";
    const persisted = JSON.parse(localStorage.getItem(key) || "{}");
    const line = { id: crypto.randomUUID(), latex: "" };
    persisted.state = {
      ...(persisted.state || {}),
      lines: [line],
      activeLineId: line.id,
      suggestionCount: 10,
    };
    localStorage.setItem(key, JSON.stringify(persisted));
  })()`);
  await client.send("Page.reload", { ignoreCache: true });
  await sleep(1000);

  const mounted = await evaluate(`new Promise((resolve) => {
    const started = performance.now();
    const poll = () => {
      const field = document.querySelector("math-field.visual-mathfield, math-field");
      if (field?.shadowRoot) return resolve(true);
      if (performance.now() - started > 10000) return resolve(false);
      setTimeout(poll, 40);
    };
    poll();
  })`);
  assert.equal(mounted, true, "Mathfield did not mount");

  const dispatchKey = async (key: string, code: string, text = key) => {
    const vk =
      key === "\\"
        ? 220
        : key === " "
          ? 32
          : key === "Enter"
            ? 13
            : key.toUpperCase().charCodeAt(0);
    const common = {
      key,
      code,
      windowsVirtualKeyCode: vk,
      nativeVirtualKeyCode: vk,
    };
    await client!.send("Input.dispatchKeyEvent", {
      type: "keyDown",
      ...common,
      text,
      unmodifiedText: text,
    });
    await client!.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
    await sleep(12);
  };

  const typeCommand = async (command: string) => {
    await dispatchKey("\\", "Backslash", "\\");
    for (const character of command.slice(1)) {
      await dispatchKey(
        character,
        `Key${character.toUpperCase()}`,
        character,
      );
    }
  };

  const results: Array<Record<string, unknown>> = [];

  for (const command of commands) {
    const canonical = canonicalSourceWrapperCommand(command);
    const expectedKind = sourceWrapperKind(command);
    assert.ok(expectedKind, `${command} lost source-wrapper classification`);

    const ready = await evaluate(`(() => {
      const field = document.querySelector("math-field.visual-mathfield, math-field");
      if (!field?.shadowRoot) return false;
      field.setValue("", {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      field.dataset.visualtexAutoExitWrapperCommand = "false";
      field.dataset.visualtexAutoExitAccent = "false";
      field.dispatchEvent(new InputEvent("input", {
        bubbles: true,
        composed: true,
        inputType: "insertText",
      }));
      field.focus();
      field.position = field.lastOffset;
      field.shadowRoot.querySelector('[part="keyboard-sink"]')?.focus({
        preventScroll: true,
      });
      return field.isConnected && field.hasFocus();
    })()`);
    assert.equal(ready, true, `${command}: field did not focus`);

    await typeCommand(command);
    await sleep(45);
    const beforeSpace = await evaluate(`(() => {
      const field = document.querySelector("math-field.visual-mathfield, math-field");
      return {
        value: field?.value ?? "",
        mode: field?.mode ?? "",
        sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
        completion: field?._mathfield?.visualTexCompletion
          ? {
              command: field._mathfield.visualTexCompletion.command ?? "",
              suggestions: field._mathfield.visualTexCompletion.suggestions ?? [],
              suggestionIndex: field._mathfield.suggestionIndex ?? -1,
              records: (field._mathfield.visualTexCompletion.suggestions ?? []).map((latex) => {
                const record = field._mathfield.visualTexCompletionRecords?.get(latex);
                return {
                  latex,
                  command: record?.command ?? "",
                  id: record?.id ?? "",
                };
              }),
            }
          : null,
        placeholderCount:
          field?.shadowRoot?.querySelectorAll(
            ".visualtex-structural-placeholder, .ML__placeholder"
          ).length ?? -1,
        raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "")
          .join(""),
      };
    })()`);
    await dispatchKey(" ", "Space", " ");
    await sleep(70);

    const shell = await evaluate(`(() => {
      const field = document.querySelector("math-field.visual-mathfield, math-field");
      return {
        value: field?.value ?? "",
        mode: field?.mode ?? "",
        sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
        sourceKind: field?.dataset.visualtexSourceWrapperKind ?? "",
        sourceContent: field?.dataset.visualtexSourceWrapperContent ?? "",
        placeholderCount:
          field?.shadowRoot?.querySelectorAll(
            ".visualtex-structural-placeholder, .ML__placeholder"
          ).length ?? -1,
        raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
          .filter((node) => !node.classList.contains("ML__suggestion"))
          .map((node) => node.textContent ?? "")
          .join(""),
      };
    })()`);

    assert.equal(
      shell.mode,
      "latex",
      `${command}: did not enter source-first mode: ${JSON.stringify({ beforeSpace, shell })}`,
    );
    assert.equal(
      shell.sourceCommand,
      canonical,
      `${command}: wrong source-wrapper command`,
    );
    assert.equal(
      shell.sourceKind,
      expectedKind,
      `${command}: wrong source-wrapper kind`,
    );
    assert.equal(shell.sourceContent, "", `${command}: non-empty initial source`);
    assert.equal(
      shell.placeholderCount,
      0,
      `${command}: leaked placeholder on acceptance`,
    );
    assert.equal(
      shell.raw,
      `${canonical}{}`,
      `${command}: wrong raw source shell`,
    );

    const content = command === "\\c" ? "c" : "A";
    await dispatchKey(content, `Key${content.toUpperCase()}`, content);
    await sleep(45);

    const typed = await evaluate(`(() => {
      const field = document.querySelector("math-field.visual-mathfield, math-field");
      return {
        value: field?.value ?? "",
        mode: field?.mode ?? "",
        sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
        sourceContent: field?.dataset.visualtexSourceWrapperContent ?? "",
        placeholderCount:
          field?.shadowRoot?.querySelectorAll(
            ".visualtex-structural-placeholder, .ML__placeholder"
          ).length ?? -1,
      };
    })()`);
    assert.equal(typed.mode, "latex", `${command}: exited before confirmation`);
    assert.equal(
      typed.sourceCommand,
      canonical,
      `${command}: source wrapper disappeared before Enter`,
    );
    assert.equal(
      typed.sourceContent,
      content,
      `${command}: typed source was not retained`,
    );
    assert.equal(
      typed.placeholderCount,
      0,
      `${command}: placeholder appeared while typing source`,
    );

    await dispatchKey("Enter", "Enter", "");
    await sleep(70);
    const committed = await evaluate(`(() => {
      const field = document.querySelector("math-field.visual-mathfield, math-field");
      return {
        value: field?.value ?? "",
        mode: field?.mode ?? "",
        sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
        placeholderCount:
          field?.shadowRoot?.querySelectorAll(
            ".visualtex-structural-placeholder, .ML__placeholder"
          ).length ?? -1,
        error: Boolean(field?.shadowRoot?.querySelector(".ML__error")),
      };
    })()`);

    assert.equal(committed.mode, "math", `${command}: Enter did not commit`);
    assert.equal(
      committed.sourceCommand,
      "",
      `${command}: source-wrapper state survived Enter`,
    );
    assert.equal(
      committed.placeholderCount,
      0,
      `${command}: committed formula contains placeholder`,
    );
    assert.equal(committed.error, false, `${command}: committed with MathLive error`);
    results.push({
      command,
      canonical,
      kind: expectedKind,
      committed: committed.value,
    });
  }

  console.log(
    JSON.stringify(
      {
        passed: results.length,
        variants: results.filter((item) => item.kind === "variant").length,
        accents: results.filter((item) => item.kind === "accent").length,
        commands: results.map((item) => item.command),
      },
      null,
      2,
    ),
  );
  console.log("VisualTeX exhaustive source-wrapper regression passed.");
} finally {
  client?.close();
  browser?.kill("SIGTERM");
  preview.kill("SIGTERM");
  await sleep(200);
  await rm(profile, {
    recursive: true,
    force: true,
    maxRetries: 4,
    retryDelay: 150,
  }).catch((error: any) => {
    if (process.platform !== "win32" || error?.code !== "EBUSY") throw error;
  });
}
