import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";
import {
  browserTestProfilePath,
  resolveBrowserTestChromePath,
} from "./browser_test_runtime.mjs";

const offset = process.pid % 1000;
const previewPort = 9800 + offset;
const debugPort = 14800 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}/editor`;
const profile = browserTestProfilePath("visualtex-enter-transition");
const chromePath = resolveBrowserTestChromePath();
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while the local process starts.
    }
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

class CdpClient {
  constructor(url) {
    this.url = url;
    this.nextId = 1;
    this.pending = new Map();
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (!message.id) return;
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }

  send(method, params = {}) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }

  close() {
    this.socket?.close();
  }
}

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
  let chrome;
  let client;

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
        `--user-data-dir=${profile}`,
        "--window-size=1400,1000",
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page = targets.find(
      (target) => target.type === "page" && target.url.startsWith(baseUrl),
    );
    if (!page) throw new Error("No VisualTeX target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.navigate", { url: baseUrl });
    await sleep(550);

    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", {
        expression,
        awaitPromise: true,
        returnByValue: true,
      });
      if (result.exceptionDetails) {
        throw new Error(
          result.exceptionDetails.exception?.description ||
            result.exceptionDetails.text ||
            "Runtime.evaluate failed",
        );
      }
      return result.result.value;
    };

    const reload = async () => {
      await client.send("Page.reload", { ignoreCache: true });
      await sleep(520);
      await evaluate(`new Promise((resolve) => {
        const poll = () => document.querySelector("math-field")
          ? resolve(true)
          : setTimeout(poll, 25);
        poll();
      })`);
    };

    await evaluate(`(() => {
      localStorage.setItem("visualtex.onboarding.v3.completed", "true");
      localStorage.setItem("visualtex.office.macos.first-run.v1.completed", "true");
      localStorage.setItem("visualtex.onboarding.macos.desktop.v1.2.0.completed", "true");
      localStorage.setItem("visualtex.office.macos.native-first-run.v1.2.0.completed", "true");
    })()`);
    await reload();

    const loadSingleLine = async (latex, zoom = 1) => {
      await evaluate(`(() => {
        const key = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(key) || "{}");
        persisted.state = {
          ...(persisted.state || {}),
          lines: [{ id: "enter-probe-line", latex: ${JSON.stringify(latex)} }],
          activeLineId: "enter-probe-line",
          sourceOpen: false,
          zoom: ${JSON.stringify(zoom)},
        };
        localStorage.setItem(key, JSON.stringify(persisted));
      })()`);
      await reload();
    };

    const pressEnterWithProbe = async (prefix, selectionEndPrefix = null) => {
      const setup = await evaluate(`(() => {
        const field = document.querySelector("math-field");
        const compact = (value) => value
          .replace(/\\s+/g, "")
          .replace(/\\{([A-Za-z0-9])\\}/g, "$1");
        const findOffset = (target) => {
          for (let offset = 0; offset <= field.lastOffset; offset += 1) {
            if (compact(field.getValue(0, offset, "latex")) === compact(target)) {
              return offset;
            }
          }
          return -1;
        };
        const start = findOffset(${JSON.stringify(prefix)});
        const end = ${selectionEndPrefix === null
          ? "start"
          : `findOffset(${JSON.stringify(selectionEndPrefix)})`};
        if (start < 0 || end < 0) {
          return { error: "offset-not-found", start, end, value: field.value };
        }
        field.focus();
        field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus({ preventScroll: true });
        field.selection = {
          ranges: [[Math.min(start, end), Math.max(start, end)]],
          direction: start === end ? "none" : "forward",
        };
        if (start === end) field.position = end;

        const originalFields = Array.from(document.querySelectorAll("math-field"));
        const originalRows = Array.from(document.querySelectorAll(".formula-line"));
        const samples = [];
        const focusEvents = [];
        let createdField = null;
        let createdRow = null;
        const deepestActiveElement = () => {
          let active = document.activeElement;
          while (active?.shadowRoot?.activeElement) {
            active = active.shadowRoot.activeElement;
          }
          return active;
        };
        const lineIdForNode = (node) => {
          if (!(node instanceof Element)) return "";
          return node.closest("math-field")?.closest(".formula-line")?.dataset.lineId ||
            node.closest(".formula-line")?.dataset.lineId || "";
        };
        const capture = (label, keyboardEvent = null) => {
          const surface = document.querySelector(".multi-line-editor");
          const fields = Array.from(document.querySelectorAll("math-field"));
          const rows = Array.from(document.querySelectorAll(".formula-line"));
          const deepActive = deepestActiveElement();
          if (fields.length > originalFields.length && !createdField) {
            createdField = fields[fields.length - 1] ?? null;
            createdRow = rows[rows.length - 1] ?? null;
          }
          const deepActiveField = fields.find(
            (item) => item === deepActive || Boolean(deepActive && item.shadowRoot?.contains(deepActive)),
          );
          samples.push({
            label,
            time: performance.now(),
            activeLineId: surface?.dataset.activeLineId || "",
            activeRows: rows
              .filter((row) => row.classList.contains("is-active"))
              .map((row) => row.dataset.lineId || ""),
            fieldCount: fields.length,
            originalFieldStable: originalFields.map(
              (original) => original.isConnected && fields.includes(original),
            ),
            originalRowStable: originalRows.map(
              (original) => original.isConnected && rows.includes(original),
            ),
            values: fields.map((item) => item.value),
            positions: fields.map((item) => item.position),
            selections: fields.map((item) => item.selection),
            fieldHeights: fields.map((item) => item.getBoundingClientRect().height),
            rowHeights: rows.map((item) => item.getBoundingClientRect().height),
            fieldInlineHeights: fields.map((item) => item.style.height),
            fieldComputedMinHeights: fields.map((item) => getComputedStyle(item).minHeight),
            rowHeightVariables: rows.map((item) => item.style.getPropertyValue("--formula-row-height")),
            geometryGroups: fields.map((item) => {
              const content = item.shadowRoot?.querySelector('[part="content"]');
              const heightFor = (selector) => {
                const rects = Array.from(content?.querySelectorAll(selector) ?? [])
                  .filter((node) => {
                    const rect = node.getBoundingClientRect();
                    const style = getComputedStyle(node);
                    return rect.height > 0 &&
                      style.display !== "none" &&
                      style.visibility !== "hidden" &&
                      style.opacity !== "0";
                  })
                  .map((node) => node.getBoundingClientRect());
                return rects.length
                  ? Math.max(...rects.map((rect) => rect.bottom)) -
                    Math.min(...rects.map((rect) => rect.top))
                  : 0;
              };
              return {
                content: content?.getBoundingClientRect().height ?? 0,
                atoms: heightFor('[data-atom-id]'),
                structures: heightFor('.ML__mfrac, .ML__sqrt, .ML__op-group, .ML__vlist'),
                bases: heightFor('.ML__base'),
              };
            }),
            createdFieldStable: !createdField || (createdField.isConnected && fields.includes(createdField)),
            createdRowStable: !createdRow || (createdRow.isConnected && rows.includes(createdRow)),
            backgrounds: rows.map((item) => getComputedStyle(item).backgroundColor),
            transitionDurations: rows.map((item) => getComputedStyle(item).transitionDuration),
            documentActiveTag: document.activeElement?.tagName || "",
            deepActiveTag: deepActive?.tagName || "",
            deepActiveLineId:
              deepActiveField?.closest(".formula-line")?.dataset.lineId ||
              lineIdForNode(deepActive),
            keyTarget: keyboardEvent?.target?.tagName || "",
            keyPath: keyboardEvent
              ? keyboardEvent.composedPath().map((node) => node?.tagName || node?.nodeName || "")
              : [],
          });
        };
        const onFocusIn = (event) => {
          focusEvents.push({
            time: performance.now(),
            tag: event.target?.tagName || "",
            lineId: lineIdForNode(event.target),
          });
        };
        document.addEventListener("focusin", onFocusIn, true);
        window.__visualTexEnterProbePromise = new Promise((resolve) => {
          field.addEventListener("keydown", (event) => {
            if (event.key === "Enter") capture("field-after", event);
          }, { capture: true, once: true });
          window.addEventListener("keydown", (event) => {
            if (event.key !== "Enter") return;
            capture("keydown", event);
            queueMicrotask(() => capture("microtask"));
            requestAnimationFrame(() => {
              capture("raf1");
              requestAnimationFrame(() => {
                capture("raf2");
                requestAnimationFrame(() => capture("raf3"));
              });
            });
            setTimeout(() => capture("40ms"), 40);
            setTimeout(() => capture("90ms"), 90);
            setTimeout(() => {
              capture("180ms");
              document.removeEventListener("focusin", onFocusIn, true);
              resolve({ samples, focusEvents });
            }, 180);
          }, { capture: true, once: true });
        });
        return { start, end, value: field.value };
      })()`);
      assert.equal(setup.error, undefined, JSON.stringify(setup));

      await evaluate(`(() => {
        const field = document.querySelector("math-field");
        field.dispatchEvent(new KeyboardEvent("keydown", {
          key: "Enter",
          code: "Enter",
          bubbles: true,
          composed: true,
          cancelable: true,
        }));
        field.dispatchEvent(new KeyboardEvent("keyup", {
          key: "Enter",
          code: "Enter",
          bubbles: true,
          composed: true,
          cancelable: true,
        }));
      })()`);
      return evaluate(`window.__visualTexEnterProbePromise`);
    };

    const assertStableTransition = (label, probe) => {
      assert.ok(probe.samples.length >= 7, `${label}: missing frame samples`);
      const postMutation = probe.samples.filter((sample) => sample.label !== "keydown");
      for (const sample of postMutation) {
        assert.equal(
          sample.fieldCount,
          2,
          `${label}: field count changed at ${sample.label}: ${JSON.stringify({ sample, keydown: probe.samples[0] })}`,
        );
        assert.ok(
          sample.originalFieldStable.every(Boolean),
          `${label}: original Mathfield was remounted at ${sample.label}: ${JSON.stringify(sample)}`,
        );
        assert.ok(
          sample.originalRowStable.every(Boolean),
          `${label}: original row was remounted at ${sample.label}: ${JSON.stringify(sample)}`,
        );
        assert.ok(
          sample.fieldHeights.every((height) => height > 0),
          `${label}: zero field height at ${sample.label}: ${JSON.stringify(sample)}`,
        );
        assert.ok(
          sample.rowHeights.every((height) => height > 0),
          `${label}: zero row height at ${sample.label}: ${JSON.stringify(sample)}`,
        );
        assert.equal(
          sample.activeRows.length,
          1,
          `${label}: active row count at ${sample.label}: ${JSON.stringify(sample)}`,
        );
      }

      for (const sample of postMutation) {
        assert.equal(
          sample.createdFieldStable,
          true,
          `${label}: new Mathfield was remounted at ${sample.label}: ${JSON.stringify(sample)}`,
        );
        assert.equal(
          sample.createdRowStable,
          true,
          `${label}: new row was remounted at ${sample.label}: ${JSON.stringify(sample)}`,
        );
        assert.ok(
          sample.fieldInlineHeights[1],
          `${label}: new Mathfield had no synchronous inline height at ${sample.label}: ${JSON.stringify(sample)}`,
        );
        assert.ok(
          sample.rowHeightVariables[1],
          `${label}: new row had no synchronous height variable at ${sample.label}: ${JSON.stringify(sample)}`,
        );
      }

      const microtask = probe.samples.find((sample) => sample.label === "microtask");
      const final = probe.samples.find((sample) => sample.label === "180ms");
      assert.ok(microtask, `${label}: microtask sample missing`);
      assert.ok(final, `${label}: final sample missing`);
      assert.equal(final.activeRows[0], final.activeLineId, `${label}: active row mismatch`);
      assert.equal(final.deepActiveLineId, final.activeLineId, `${label}: focus is not in active line`);

      const settledValues = JSON.stringify(final.values);
      for (const sample of postMutation) {
        assert.equal(
          JSON.stringify(sample.values),
          settledValues,
          `${label}: formula values changed after synchronous Enter at ${sample.label}`,
        );
        assert.equal(
          sample.positions[1],
          0,
          `${label}: new line caret moved after Enter at ${sample.label}: ${JSON.stringify(sample)}`,
        );
      }

      const newRowHeights = postMutation.map((sample) => ({
        label: sample.label,
        height: sample.rowHeights[1],
        fieldHeight: sample.fieldHeights[1],
        variable: sample.rowHeightVariables[1],
        inlineHeight: sample.fieldInlineHeights[1],
        geometry: sample.geometryGroups[1],
      }));
      for (let index = 1; index < newRowHeights.length; index += 1) {
        const previous = newRowHeights[index - 1];
        const current = newRowHeights[index];
        assert.ok(
          current.height >= previous.height - 1,
          `${label}: new row visibly shrank from ${previous.label} to ${current.label}: ${JSON.stringify(newRowHeights)}`,
        );
        assert.ok(
          current.fieldHeight >= previous.fieldHeight - 1,
          `${label}: Mathfield visibly shrank from ${previous.label} to ${current.label}: ${JSON.stringify(newRowHeights)}`,
        );
      }
      assert.ok(
        microtask.rowHeights[1] <= final.rowHeights[1] + 1,
        `${label}: first post-Enter row was higher than its settled height: ${JSON.stringify(newRowHeights)}`,
      );
      assert.ok(
        microtask.fieldHeights[1] <= final.fieldHeights[1] + 1,
        `${label}: first post-Enter Mathfield was higher than its settled height: ${JSON.stringify(newRowHeights)}`,
      );
      const visibleHeights = newRowHeights.filter((sample) =>
        ["raf1", "raf2", "raf3", "40ms", "90ms", "180ms"].includes(
          sample.label,
        ),
      );
      const visibleRowValues = visibleHeights.map((sample) => sample.height);
      const visibleFieldValues = visibleHeights.map(
        (sample) => sample.fieldHeight,
      );
      assert.ok(
        Math.max(...visibleRowValues) - Math.min(...visibleRowValues) <= 1,
        `${label}: visible row height changed after Enter: ${JSON.stringify(newRowHeights)}`,
      );
      assert.ok(
        Math.max(...visibleFieldValues) - Math.min(...visibleFieldValues) <= 1,
        `${label}: visible Mathfield height changed after Enter: ${JSON.stringify(newRowHeights)}`,
      );

      const targetFocusEvents = probe.focusEvents.filter(
        (event) => event.lineId === final.activeLineId,
      );
      assert.ok(
        targetFocusEvents.length <= 2,
        `${label}: target line was repeatedly refocused: ${JSON.stringify(probe.focusEvents)}`,
      );
    };

    const cases = [
      {
        label: "empty line",
        latex: "",
        prefix: "",
        verify(values) {
          assert.deepEqual(values, ["", ""]);
        },
      },
      {
        label: "simple formula",
        latex: "abcdef",
        prefix: "abc",
        verify(values) {
          assert.deepEqual(values, ["abc", "def"]);
        },
      },
      {
        label: "fraction",
        latex: "p+\\frac{a}{b}+q",
        prefix: "p+",
        verify(values) {
          assert.equal(values[0], "p+");
          assert.match(values[1], /\\frac/);
          assert.match(values[1], /\+q$/);
        },
      },
      {
        label: "integral",
        latex: "u+\\int_{0}^{1}f(x)\\,\\mathrm{d}x",
        prefix: "u+",
        verify(values) {
          assert.equal(values[0], "u+");
          assert.match(values[1], /\\int/);
          assert.match(values[1], /\\mathrm\{d\}x/);
        },
      },
      {
        label: "scripts",
        latex: "x_{i}^{2}+y",
        prefix: "x_{i}^{2}",
        verify(values) {
          assert.match(values[0], /^x_/);
          assert.equal(values[1], "+y");
        },
      },
      {
        label: "selection",
        latex: "abcdef",
        prefix: "a",
        selectionEndPrefix: "abc",
        verify(values) {
          assert.deepEqual(values, ["a", "def"]);
        },
      },
      {
        label: "fraction end creates empty line",
        latex: "p+\\frac{a}{b}",
        prefix: "p+\\frac{a}{b}",
        verify(values) {
          assert.match(values[0], /\\frac/);
          assert.equal(values[1], "");
        },
      },
      {
        label: "integral end creates empty line",
        latex: "\\int_{0}^{1}\\frac{f(x)}{g(x)}\\,\\mathrm{d}x",
        prefix: "\\int_{0}^{1}\\frac{f(x)}{g(x)}\\,\\mathrm{d}x",
        verify(values) {
          assert.match(values[0], /\\int/);
          assert.equal(values[1], "");
        },
      },
      {
        label: "nested radical fraction",
        latex: "a+\\sqrt{\\frac{x_i^2}{1+x}}+b",
        prefix: "a+",
        verify(values) {
          assert.equal(values[0], "a+");
          assert.match(values[1], /\\sqrt/);
          assert.match(values[1], /\\frac/);
        },
      },
      {
        label: "complex selection",
        latex: "a+\\frac{x}{y}+b",
        prefix: "a+",
        selectionEndPrefix: "a+\\frac{x}{y}",
        verify(values) {
          assert.equal(values[0], "a+");
          assert.equal(values[1], "+b");
        },
      },
      {
        label: "small zoom simple",
        latex: "abcdef",
        prefix: "abc",
        zoom: 0.65,
        verify(values) {
          assert.deepEqual(values, ["abc", "def"]);
        },
      },
      {
        label: "large zoom fraction",
        latex: "a+\\frac{x}{y}",
        prefix: "a+",
        zoom: 1.35,
        verify(values) {
          assert.equal(values[0], "a+");
          assert.match(values[1], /\\frac/);
        },
      },
    ];

    for (const testCase of cases) {
      await loadSingleLine(testCase.latex, testCase.zoom ?? 1);
      const probe = await pressEnterWithProbe(
        testCase.prefix,
        testCase.selectionEndPrefix ?? null,
      );
      assertStableTransition(testCase.label, probe);
      const final = probe.samples.find((sample) => sample.label === "180ms");
      testCase.verify(final.values);
    }

    console.log("Enter transition regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(250);
    for (let attempt = 0; attempt < 4; attempt += 1) {
      try {
        await rm(profile, { recursive: true, force: true });
        break;
      } catch (error) {
        if (attempt === 3) throw error;
        await sleep(120);
      }
    }
  }
}

await main();
