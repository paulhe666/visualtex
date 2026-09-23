import { spawn } from "node:child_process";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";
import { commandRegistry } from "../src/autocomplete/commandRegistry.ts";

const offset = process.pid % 700;
const previewPort = 7100 + offset;
const debugPort = 12700 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = createBrowserProfilePath("visualtex-command-runtime-audit");
const chromePath = resolveChromiumExecutable();
const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url: string, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {}
    await sleep(60);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

class CdpClient {
  socket?: WebSocket;
  nextId = 1;
  pending = new Map<
    number,
    { resolve: (value: any) => void; reject: (error: Error) => void }
  >();

  constructor(readonly url: string) {}

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise<void>((resolve, reject) => {
      this.socket?.addEventListener("open", () => resolve(), { once: true });
      this.socket?.addEventListener("error", () => reject(new Error("CDP socket failed")), {
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
      this.socket?.send(JSON.stringify({ id, method, params }));
    });
  }

  close() {
    this.socket?.close();
  }
}

const literalDelimiters = new Set([
  "\\langle",
  "\\rangle",
  "\\lfloor",
  "\\rfloor",
  "\\lceil",
  "\\rceil",
  "\\lvert",
  "\\rvert",
  "\\lVert",
  "\\rVert",
  "\\vert",
  "\\Vert",
]);

const sourceWrapperCommands = new Set([
  "\\mathrm",
  "\\mathbf",
  "\\mathit",
  "\\mathsf",
  "\\mathtt",
  "\\mathnormal",
  "\\boldsymbol",
  "\\mathbb",
  "\\mathcal",
  "\\mathscr",
  "\\mathfrak",
  "\\mathbfit",
  "\\bm",
  "\\symup",
  "\\symit",
  "\\symbf",
  "\\symbfup",
  "\\symbfit",
  "\\symbb",
  "\\symcal",
  "\\symbfcal",
  "\\symscr",
  "\\symbfscr",
  "\\symfrak",
  "\\symbffrak",
  "\\symsfup",
  "\\symsfit",
  "\\symbfsfup",
  "\\symbfsfit",
  "\\symtt",
  "\\acute",
  "\\grave",
  "\\dot",
  "\\ddot",
  "\\dddot",
  "\\ddddot",
  "\\tilde",
  "\\bar",
  "\\breve",
  "\\check",
  "\\hat",
  "\\vec",
  "\\widehat",
  "\\widetilde",
  "\\overline",
  "\\overrightarrow",
  "\\overleftarrow",
  "\\overleftrightarrow",
  "\\mathring",
]);

const sourceWrapperCanonicalAliases = new Map([
  ["\\bf", "\\mathbf"],
  ["\\bfseries", "\\mathbf"],
  ["\\it", "\\mathit"],
  ["\\rm", "\\mathrm"],
  ["\\sf", "\\mathsf"],
  ["\\tt", "\\mathtt"],
  ["\\cal", "\\mathcal"],
  ["\\Bbb", "\\mathbb"],
  ["\\frak", "\\mathfrak"],
  ["\\bold", "\\mathbfit"],
  ["\\boldmath", "\\mathbfit"],
  ["\\pmb", "\\mathbfit"],
  ["\\simbfit", "\\symbfit"],
]);

const bareOperatorCommands = new Set([
  "\\int",
  "\\iint",
  "\\iiint",
  "\\oint",
  "\\oiint",
  "\\oiiint",
  "\\sum",
  "\\prod",
  "\\lim",
]);

const uniqueCommands = [
  ...new Set(
    commandRegistry
      .map((record) => record.command)
      .filter((command) => /^\\[A-Za-z]+$/.test(command)),
  ),
].sort((a, b) => a.localeCompare(b));

const startIndex = Math.max(0, Number.parseInt(process.argv[2] ?? "0", 10) || 0);
const requestedCount = Math.max(
  1,
  Number.parseInt(process.argv[3] ?? String(uniqueCommands.length), 10) ||
    uniqueCommands.length,
);
const commandsUnderTest = uniqueCommands.slice(
  startIndex,
  startIndex + requestedCount,
);

const parameterizedCommands = new Set(
  commandRegistry
    .filter((record) => record.insertTemplate.includes("\\placeholder{}"))
    .map((record) => record.command),
);

async function main() {
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
  let chrome: ReturnType<typeof spawn> | undefined;
  let client: CdpClient | undefined;

  try {
    await waitFor(baseUrl);
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
        "--window-size=1400,1000",
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = (await fetch(
      `http://127.0.0.1:${debugPort}/json/list`,
    ).then((response) => response.json())) as Array<{
      type: string;
      webSocketDebuggerUrl?: string;
    }>;
    const page = targets.find(
      (target) => target.type === "page" && target.webSocketDebuggerUrl,
    );
    if (!page?.webSocketDebuggerUrl) throw new Error("No browser page target");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");

    const evaluate = async (expression: string) => {
      const result = await client!.send("Runtime.evaluate", {
        expression,
        returnByValue: true,
        awaitPromise: true,
      });
      if (result.exceptionDetails) {
        throw new Error(
          result.exceptionDetails.exception?.description ??
            result.exceptionDetails.text ??
            "Runtime evaluation failed",
        );
      }
      return result.result?.value;
    };

    const waitForEvaluation = async (
      expression: string,
      description: string,
      timeoutMs = 10000,
    ) => {
      const started = Date.now();
      let lastValue: unknown = null;
      while (Date.now() - started < timeoutMs) {
        lastValue = await evaluate(expression);
        if ((lastValue as any)?.ready) return lastValue;
        await sleep(40);
      }
      throw new Error(
        `Timed out waiting for ${description}: ${JSON.stringify(lastValue)}`,
      );
    };

    const key = async (value: string, code: string, virtualKeyCode: number) => {
      const common = {
        key: value,
        code,
        windowsVirtualKeyCode: virtualKeyCode,
        nativeVirtualKeyCode: virtualKeyCode,
      };
      await client!.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        ...(value.length === 1 ? { text: value, unmodifiedText: value } : {}),
      });
      await client!.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(8);
    };

    const typeText = async (text: string) => {
      for (const character of text) {
        const code =
          character === "\\"
            ? "Backslash"
            : /[A-Za-z]/.test(character)
              ? `Key${character.toUpperCase()}`
              : "Unidentified";
        const virtualKeyCode =
          character === "\\"
            ? 220
            : /[A-Za-z]/.test(character)
              ? character.toUpperCase().charCodeAt(0)
              : character.charCodeAt(0);
        await key(character, code, virtualKeyCode);
      }
    };

    await waitForEvaluation(
      `(() => ({ ready: Boolean(document.querySelector("math-field")?.isConnected) }))()`,
      "mounted math field",
    );

    const failures: any[] = [];
    const stats = {
      totalCommands: uniqueCommands.length,
      batchStart: startIndex,
      commands: commandsUnderTest.length,
      parameterized: 0,
      sourceWrappers: 0,
      literalDelimiters: 0,
      placeholderTemplates: 0,
      bareOperators: 0,
    };

    for (const command of commandsUnderTest) {
      const rawName = command.slice(1);
      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        if (!field) return false;
        field.focus();
        field.executeCommand("selectAll");
        field.executeCommand("deleteBackward");
        field.position = field.lastOffset;
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        return true;
      })()`);
      await waitForEvaluation(
        `(() => {
          const field = document.querySelector("math-field");
          return { ready: Boolean(field?.isConnected && field.value === ""), value: field?.value ?? "" };
        })()`,
        `empty before ${command}`,
      );

      await typeText(command);
      await sleep(15);
      const before = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const mf = field?._mathfield;
        const completion = mf?.visualTexCompletion;
        const suggestion = completion?.suggestions?.[mf?.suggestionIndex ?? 0] ?? "";
        const record = suggestion
          ? mf?.visualTexCompletionRecords?.get(suggestion)
          : null;
        return {
          raw: [...(field?.shadowRoot?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          suggestion,
          record: record ? {
            id: record.id ?? "",
            command: record.command ?? "",
            latex: record.latex ?? "",
            score: record.score ?? null,
          } : null,
          suggestions: completion?.suggestions ?? [],
        };
      })()`);

      await key(" ", "Space", 32);
      await sleep(35);

      const after = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const root = field?.shadowRoot;
        const placeholderNodes = [...(root?.querySelectorAll(".visualtex-structural-placeholder, .ML__placeholder") ?? [])];
        const rect = (node) => {
          const bounds = node.getBoundingClientRect();
          return { width: bounds.width, height: bounds.height, left: bounds.left, top: bounds.top };
        };
        return {
          value: field?.value ?? "",
          mode: field?.mode ?? "",
          selection: field?.selection ?? null,
          selectionIsPlaceholder: field?.selectionIsPlaceholder ?? false,
          selectedPlaceholderCount:
            root?.querySelectorAll(".visualtex-structural-placeholder.ML__selected, .ML__placeholder.ML__selected").length ?? 0,
          placeholderCaretCount:
            root?.querySelectorAll(".visualtex-structural-placeholder-caret").length ?? 0,
          sourceCommand: field?.dataset.visualtexSourceWrapperCommand ?? "",
          sourceKind: field?.dataset.visualtexSourceWrapperKind ?? "",
          placeholderCount: placeholderNodes.length,
          placeholderRects: placeholderNodes.map(rect),
          raw: [...(root?.querySelectorAll(".ML__raw-latex") ?? [])]
            .filter((node) => !node.classList.contains("ML__suggestion"))
            .map((node) => node.textContent ?? "")
            .join(""),
          contentRect: (() => {
            const node = root?.querySelector('[part="content"]');
            if (!node) return null;
            const bounds = node.getBoundingClientRect();
            return { width: bounds.width, height: bounds.height };
          })(),
        };
      })()`);

      const selectedLatex = String(before?.suggestion ?? "");
      const selectedPlaceholderCount =
        selectedLatex.match(/\\placeholder\{\}/g)?.length ?? 0;
      const acceptedCommand = String(before?.record?.command ?? command);
      const canonicalSourceCommand =
        sourceWrapperCanonicalAliases.get(acceptedCommand) ?? acceptedCommand;
      const isSourceWrapper = sourceWrapperCommands.has(canonicalSourceCommand);
      const isLiteralDelimiter = literalDelimiters.has(command);
      const isParameterized =
        parameterizedCommands.has(command) &&
        !bareOperatorCommands.has(command) &&
        !literalDelimiters.has(command);

      if (isParameterized) stats.parameterized += 1;
      if (isSourceWrapper) stats.sourceWrappers += 1;
      if (isLiteralDelimiter) stats.literalDelimiters += 1;
      if (selectedPlaceholderCount > 0) stats.placeholderTemplates += 1;
      if (bareOperatorCommands.has(command)) stats.bareOperators += 1;

      const reasons: string[] = [];
      if (!after.contentRect || after.contentRect.width <= 0 || after.contentRect.height <= 0)
        reasons.push("rendered content has invalid geometry");
      if (
        after.placeholderRects.some(
          (rect: any) =>
            !Number.isFinite(rect.width) ||
            !Number.isFinite(rect.height) ||
            rect.width <= 0 ||
            rect.height <= 0,
        )
      )
        reasons.push("placeholder has invalid geometry");

      if (isSourceWrapper) {
        if (after.value !== canonicalSourceCommand + "{}")
          reasons.push(`source wrapper value=${after.value}`);
        if (after.mode !== "latex") reasons.push(`source wrapper mode=${after.mode}`);
        if (after.sourceCommand !== canonicalSourceCommand)
          reasons.push(`source wrapper command=${after.sourceCommand}`);
        if (after.placeholderCount !== 0)
          reasons.push(`source wrapper placeholderCount=${after.placeholderCount}`);
      } else {
        if (after.mode !== "math") reasons.push(`math command mode=${after.mode}`);
        if (after.sourceCommand)
          reasons.push(`unexpected source wrapper=${after.sourceCommand}`);
      }

      if (isLiteralDelimiter && before.suggestion !== command)
        reasons.push(`literal delimiter suggestion=${before.suggestion}`);

      if (
        isParameterized &&
        selectedLatex === command
      )
        reasons.push("parameterized command completed as bare command");

      if (
        !isSourceWrapper &&
        selectedPlaceholderCount > 0 &&
        after.placeholderCount !== selectedPlaceholderCount
      )
        reasons.push(
          `placeholder count ${after.placeholderCount} != selected template ${selectedPlaceholderCount}`,
        );

      if (
        !isSourceWrapper &&
        selectedPlaceholderCount > 0 &&
        (after.selectedPlaceholderCount !== 1 || after.placeholderCaretCount !== 1)
      )
        reasons.push(
          `template selection invalid selected=${after.selectedPlaceholderCount} caret=${after.placeholderCaretCount}`,
        );

      if (reasons.length) {
        failures.push({ command, reasons, before, after });
      }
    }

    console.log(
      JSON.stringify(
        {
          ...stats,
          failures: failures.slice(0, 50),
          failureCount: failures.length,
        },
        null,
        2,
      ),
    );
    if (failures.length) {
      throw new Error(
        `Command runtime audit failed for ${failures.length}/${commandsUnderTest.length} commands in batch ${startIndex}`,
      );
    }
    console.log("Command runtime audit passed");
  } finally {
    client?.close();
    chrome?.kill();
    preview.kill();
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
