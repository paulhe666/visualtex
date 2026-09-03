import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import process from "node:process";

const portOffset = process.pid % 800;
const previewPort = 7600 + portOffset;
const debugPort = 12600 + portOffset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const chromeProfile = `/tmp/visualtex-align-cases-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(url, timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // Retry while local processes start.
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
        `--user-data-dir=${chromeProfile}`,
        "--window-size=1400,1000",
        baseUrl,
      ],
      { stdio: "ignore" },
    );
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
    const page = targets.find(
      (target) => target.type === "page" && target.url.startsWith(baseUrl),
    );
    if (!page) throw new Error("No VisualTeX Chrome page target found");
    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await sleep(500);

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

    const typeKey = async (key, code, text = key, modifiers = 0) => {
      const specialVirtualKey =
        key === "Enter" ? 13 : key === "Tab" ? 9 : key === "Escape" ? 27 : 0;
      const virtualKey =
        specialVirtualKey ||
        (key.length === 1 ? key.toUpperCase().charCodeAt(0) : 0);
      const common = {
        key,
        code,
        modifiers,
        windowsVirtualKeyCode: virtualKey,
        nativeVirtualKeyCode: virtualKey,
      };
      await client.send("Input.dispatchKeyEvent", {
        type: "keyDown",
        ...common,
        text,
        unmodifiedText: text,
      });
      await client.send("Input.dispatchKeyEvent", { type: "keyUp", ...common });
      await sleep(30);
    };

    await evaluate(`new Promise((resolve) => {
      const done = () => document.querySelector("math-field") ? resolve(true) : setTimeout(done, 30);
      done();
    })`);

    const casesLatex = String.raw`\begin{cases}x & x>0 \\ 0 & x\le 0\end{cases}`;
    const casesProbe = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.setValue(${JSON.stringify(casesLatex)}, {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
      });
      const model = field._mathfield?.model;
      const positions = [];
      for (let position = 0; position <= field.lastOffset; position += 1) {
        field.position = position;
        positions.push({
          position,
          environment: model?.parentEnvironment?.environmentName ?? null,
          depth: field.getElementInfo(position)?.depth ?? null,
          latex: field.getElementInfo(position)?.latex ?? null,
        });
      }
      return { value: field.value, positions };
    })()`);
    assert.match(casesProbe.value, /\\begin\{cases\}/);
    assert.ok(
      casesProbe.positions.some((entry) => entry.environment === "cases"),
      `MathLive did not expose the active cases environment: ${JSON.stringify(casesProbe.positions)}`,
    );

    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.setValue("", { insertionMode: "replaceAll" });
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus();
      return true;
    })()`);
    await typeKey("\\", "Backslash", "\\");
    await typeKey("b", "KeyB", "b");
    await sleep(120);
    const shortCasesSuggestion = await evaluate(`(() => {
      const panel = document.querySelector('#visualtex-native-input-suggestion-popover');
      return {
        visible: panel?.classList.contains('is-visible') ?? false,
        commands: Array.from(panel?.querySelectorAll('li[data-command]') ?? [])
          .map((item) => item.dataset.command ?? ''),
      };
    })()`);
    assert.equal(
      shortCasesSuggestion.visible,
      true,
      "\\b did not show the native-style suggestion popover",
    );
    assert.ok(
      shortCasesSuggestion.commands.includes("\\begin{cases}"),
      `\\b did not include the cases environment: ${JSON.stringify(shortCasesSuggestion.commands)}`,
    );

    await client.send("Page.reload", { ignoreCache: true });
    await sleep(500);
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.setValue("", { insertionMode: "replaceAll" });
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus();
      return true;
    })()`);
    await typeKey("\\", "Backslash", "\\");
    for (const letter of "begin") {
      await typeKey(letter, `Key${letter.toUpperCase()}`);
    }
    await typeKey("{", "BracketLeft", "{", 8);
    for (const letter of "cases") {
      await typeKey(letter, `Key${letter.toUpperCase()}`);
    }
    await typeKey("}", "BracketRight", "}", 8);
    await sleep(120);
    const casesSuggestion = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      const panel = document.querySelector('#visualtex-native-input-suggestion-popover');
      return {
        value: field.value,
        raw: Array.from(field.shadowRoot?.querySelectorAll('.ML__raw-latex') ?? [])
          .map((node) => node.textContent ?? "")
          .join(""),
        visible: panel?.classList.contains('is-visible') ?? false,
        command: panel?.querySelector('li.ML__popover__current')?.dataset.command ?? '',
        commands: Array.from(panel?.querySelectorAll('li[data-command]') ?? []).map((item) => ({
          command: item.dataset.command ?? '',
          current: item.classList.contains('ML__popover__current'),
        })),
      };
    })()`);
    assert.equal(casesSuggestion.raw.replace(/\s+/g, ""), "\\begin{cases}");
    assert.equal(casesSuggestion.visible, true, "cases did not show the native-style suggestion popover");
    assert.equal(
      casesSuggestion.command,
      "\\begin{cases}",
      JSON.stringify(casesSuggestion),
    );
    assert.doesNotMatch(
      casesSuggestion.value,
      /\\begin\{cases\}/,
      "cases committed before the user accepted the native suggestion",
    );

    await typeKey(" ", "Space", " ");
    await sleep(80);
    const typedCases = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return {
        value: field.value,
        raw: Array.from(field.shadowRoot?.querySelectorAll('.ML__raw-latex') ?? [])
          .map((node) => node.textContent ?? "")
          .join(""),
      };
    })()`);
    assert.match(
      typedCases.value,
      /\\begin\{cases\}\\placeholder\{\} & \\placeholder\{\}\\end\{cases\}/,
      "Space did not accept the cases suggestion as a real editable environment",
    );
    assert.equal(typedCases.raw, "", "accepted cases input remained in raw-LaTeX mode");

    await typeKey("\\", "Backslash", "\\");
    for (const letter of "frac") {
      await typeKey(letter, `Key${letter.toUpperCase()}`);
    }
    await typeKey("Enter", "Enter", "\r");
    const candidateCommit = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return { value: field.value, lines: (field.value.match(/\\\\\\\\/g) ?? []).length };
    })()`);
    assert.match(
      candidateCommit.value,
      /\\frac/,
      "Enter did not preserve candidate-command priority inside cases",
    );
    assert.equal(
      candidateCommit.lines,
      0,
      "confirming a candidate unexpectedly added a cases row",
    );

    await typeKey("Enter", "Enter", "\r");
    const afterCasesEnter = await evaluate(`(() => {
      const field = document.querySelector("math-field");
      return { value: field.value, lines: (field.value.match(/\\\\\\\\/g) ?? []).length };
    })()`);
    assert.equal(
      afterCasesEnter.lines,
      1,
      `Enter inside cases did not add exactly one case row: ${afterCasesEnter.value}`,
    );

    const lineCountBeforeOuterEnter = await evaluate(
      `document.querySelectorAll(".formula-line").length`,
    );
    await evaluate(`(() => {
      const field = document.querySelector("math-field");
      field.position = field.lastOffset;
      field.selection = { ranges: [[field.lastOffset, field.lastOffset]], direction: "none" };
      field.focus();
      field.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus();
      return true;
    })()`);
    await typeKey("Enter", "Enter", "\r");
    const lineCountAfterOuterEnter = await evaluate(
      `document.querySelectorAll(".formula-line").length`,
    );
    assert.equal(
      lineCountAfterOuterEnter,
      lineCountBeforeOuterEnter + 1,
      "Enter outside the outermost cases did not create a new VisualTeX formula line",
    );

    await evaluate(`document.querySelector('.code-format-primary')?.click()`);
    await sleep(80);
    const selectedAlignFormat = await evaluate(`(() => {
      const item = document.querySelector('[data-format="align-star"]');
      item?.click();
      return Boolean(item);
    })()`);
    assert.equal(selectedAlignFormat, true, "align-star format menu item was unavailable");
    await sleep(80);
    await evaluate(`(() => {
      const fields = Array.from(document.querySelectorAll("math-field"));
      const field = fields.at(-1);
      field?.focus();
      field?.shadowRoot?.querySelector('[part="keyboard-sink"]')?.focus();
      return true;
    })()`);
    for (const character of "a") await typeKey(character, "KeyA");
    await typeKey("&", "Digit7", "&", 8);
    await client.send("Input.insertText", { text: "=" });
    await sleep(40);
    await typeKey("b", "KeyB");
    await typeKey("Enter", "Enter", "\r");
    for (const character of "longvariable") {
      await typeKey(character, `Key${character.toUpperCase()}`);
    }
    await typeKey("&", "Digit7", "&", 8);
    await client.send("Input.insertText", { text: "=" });
    await sleep(40);
    await typeKey("d", "KeyD");
    await sleep(160);

    const alignProbe = await evaluate(`(() => {
      const fields = Array.from(document.querySelectorAll(".formula-line math-field"))
        .filter((field) => field.shadowRoot?.querySelector('.visualtex-align-marker'));
      const visualAnchor = (field) => {
        for (let position = 0; position <= field.lastOffset; position += 1) {
          const info = field.getElementInfo(position);
          if (info?.latex?.trim() === "=" && info.bounds && info.depth === 0) {
            return info.bounds.left;
          }
        }
        return null;
      };
      return {
        values: fields.map((field) => field.value),
        anchors: fields.map(visualAnchor),
      };
    })()`);
    assert.equal(alignProbe.values.length, 2, "align probe did not create two marked formula rows");
    for (const value of alignProbe.values) {
      assert.match(
        value,
        /\\class\{visualtex-align-marker\}\{\\kern0pt\}/,
        "typed & was not preserved as an explicit VisualTeX alignment point",
      );
    }
    assert.ok(
      alignProbe.anchors.every((value) => typeof value === "number"),
      `visible alignment atom was missing: ${JSON.stringify(alignProbe)}`,
    );
    assert.ok(
      Math.abs(alignProbe.anchors[0] - alignProbe.anchors[1]) <= 1.5,
      `explicit align anchors were not visually aligned: ${JSON.stringify(alignProbe.anchors)}`,
    );

    const currentAlignmentLines = [
      String.raw`\begin{matrix}\end{matrix}a\class{visualtex-align-marker}{\kern0pt}=b`,
      String.raw`\class{visualtex-align-marker}{\kern0pt}=wer`,
    ];
    await evaluate(`(() => {
      const storageKey = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        editorLayout: "classic",
        sourceOpen: true,
        latexCodeFormat: "align",
        formulaAlignment: "left",
        zoom: 0.45,
        lines: [
          { id: "align-current-1", latex: ${JSON.stringify(currentAlignmentLines[0])}, mode: "display" },
          { id: "align-current-2", latex: ${JSON.stringify(currentAlignmentLines[1])}, mode: "display" },
        ],
        activeLineId: "align-current-2",
      };
      localStorage.setItem(storageKey, JSON.stringify(persisted));
      localStorage.setItem("visualtex-desktop-editor-toolbar-open", "true");
      location.reload();
    })()`);
    await evaluate(`new Promise((resolve) => {
      const done = () => document.querySelectorAll(".formula-line math-field").length === 2
        ? resolve(true)
        : setTimeout(done, 30);
      done();
    })`);
    await sleep(320);

    const currentAlignProbe = await evaluate(`(() => {
      const fields = [...document.querySelectorAll(".formula-line math-field")];
      const entries = fields.map((field) => {
        const marker = field.shadowRoot?.querySelector(".visualtex-align-marker");
        const markerRect = marker?.getBoundingClientRect() ?? null;
        const leafEquals = [...(field.shadowRoot?.querySelectorAll("*") ?? [])]
          .filter((node) => node.children.length === 0 && node.textContent?.trim() === "=")
          .map((node) => {
            const rect = node.getBoundingClientRect();
            return {
              className: node.className,
              left: rect.left,
              right: rect.right,
              center: (rect.left + rect.right) / 2,
              width: rect.width,
            };
          })
          .filter((entry) => entry.width > 0);
        const equalInfos = [];
        for (let position = 0; position <= field.lastOffset; position += 1) {
          const info = field.getElementInfo(position);
          if (info?.latex?.trim() !== "=" || !info.bounds) continue;
          equalInfos.push({
            position,
            latex: info.latex,
            left: info.bounds.left,
            right: info.bounds.right,
            center: (info.bounds.left + info.bounds.right) / 2,
            width: info.bounds.width,
            depth: info.depth,
          });
        }
        const fieldRect = field.getBoundingClientRect();
        return {
          value: field.value,
          fieldLeft: fieldRect.left,
          fieldWidth: fieldRect.width,
          marginLeft: getComputedStyle(field).marginLeft,
          markerLeft: markerRect?.left ?? null,
          markerWidth: markerRect?.width ?? null,
          leafEquals,
          equalInfos,
        };
      });
      const finite = (values) => values.filter(Number.isFinite);
      const spread = (values) => {
        const filtered = finite(values);
        return filtered.length ? Math.max(...filtered) - Math.min(...filtered) : null;
      };
      return {
        entries,
        markerSpread: spread(entries.map((entry) => entry.markerLeft)),
        equalLeafLeftSpread: spread(entries.map((entry) => entry.leafEquals[0]?.left)),
        equalLeafCenterSpread: spread(entries.map((entry) => entry.leafEquals[0]?.center)),
        equalInfoLeftSpread: spread(entries.map((entry) => entry.equalInfos[0]?.left)),
        equalInfoCenterSpread: spread(entries.map((entry) => entry.equalInfos[0]?.center)),
      };
    })()`);
    assert.equal(currentAlignProbe.entries.length, 2, JSON.stringify(currentAlignProbe));
    assert.ok(
      typeof currentAlignProbe.equalLeafLeftSpread === "number" &&
        currentAlignProbe.equalLeafLeftSpread <= 0.75,
      `visible equal glyphs were not aligned: ${JSON.stringify(currentAlignProbe)}`,
    );
    assert.ok(
      typeof currentAlignProbe.equalInfoLeftSpread === "number" &&
        currentAlignProbe.equalInfoLeftSpread <= 0.75,
      `MathLive equal atom bounds were not aligned: ${JSON.stringify(currentAlignProbe)}`,
    );

    const markerLatex = String.raw`\class{visualtex-align-marker}{\kern0pt}`;
    const relationCases = [
      { name: "equals", token: "=", glyph: "=" },
      { name: "less-equal", token: String.raw`\le`, glyph: "≤" },
      { name: "approx", token: String.raw`\approx`, glyph: "≈" },
      { name: "arrow", token: String.raw`\to`, glyph: "→" },
      { name: "plus", token: "+", glyph: "+" },
    ];
    const relationLines = relationCases.flatMap((entry, index) => [
      {
        id: `align-relation-${index}-left`,
        latex: `a${markerLatex}${entry.token}b`,
        mode: "display",
      },
      {
        id: `align-relation-${index}-start`,
        latex: `${markerLatex}${entry.token}c`,
        mode: "display",
      },
    ]);
    await evaluate(`(() => {
      const storageKey = "visualtex-editor";
      const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
      persisted.state = {
        ...(persisted.state || {}),
        editorLayout: "classic",
        sourceOpen: true,
        latexCodeFormat: "align",
        formulaAlignment: "left",
        zoom: 0.45,
        lines: ${JSON.stringify(relationLines)},
        activeLineId: ${JSON.stringify(relationLines.at(-1)?.id)},
      };
      localStorage.setItem(storageKey, JSON.stringify(persisted));
      localStorage.setItem("visualtex-desktop-editor-toolbar-open", "true");
      location.reload();
    })()`);
    await evaluate(`new Promise((resolve) => {
      const done = () => document.querySelectorAll(".formula-line math-field").length === ${relationLines.length}
        ? resolve(true)
        : setTimeout(done, 30);
      done();
    })`);
    await sleep(320);

    const relationProbe = await evaluate(`(() => {
      const relationCases = ${JSON.stringify(relationCases)};
      const fields = [...document.querySelectorAll(".formula-line math-field")];
      return relationCases.map((entry, index) => {
        const pair = fields.slice(index * 2, index * 2 + 2);
        const atomLefts = pair.map((field) => {
          const candidates = [];
          for (let position = 0; position <= field.lastOffset; position += 1) {
            const info = field.getElementInfo(position);
            if (!info?.bounds || info.depth !== 0 || info.bounds.width <= 0) continue;
            candidates.push({
              position,
              latex: info.latex?.trim() ?? "",
              left: info.bounds.left,
              width: info.bounds.width,
            });
          }
          const exact = candidates.find((candidate) => candidate.latex === entry.token);
          if (exact) return exact.left;
          const markerIndex = field.value.indexOf("visualtex-align-marker");
          if (markerIndex < 0) return null;
          const relationCandidate = candidates.find((candidate) =>
            candidate.latex === "=" ||
            candidate.latex === "+" ||
            candidate.latex.includes(entry.token.replace(/^\\\\/, ""))
          );
          return relationCandidate?.left ?? null;
        });
        const finite = atomLefts.filter(Number.isFinite);
        return {
          name: entry.name,
          token: entry.token,
          atomLefts,
          spread: finite.length === 2 ? Math.abs(finite[0] - finite[1]) : null,
        };
      });
    })()`);
    for (const probe of relationProbe) {
      assert.ok(
        typeof probe.spread === "number" && probe.spread <= 0.75,
        `visible alignment token was not aligned: ${JSON.stringify(probe)}`,
      );
    }

    const alignmentModeResults = [];
    for (const alignment of ["center", "right"]) {
      await evaluate(`(() => {
        const storageKey = "visualtex-editor";
        const persisted = JSON.parse(localStorage.getItem(storageKey) || "{}");
        persisted.state = {
          ...(persisted.state || {}),
          editorLayout: "classic",
          sourceOpen: true,
          latexCodeFormat: "align",
          formulaAlignment: ${JSON.stringify(alignment)},
          zoom: 0.45,
          lines: [
            { id: "align-mode-1", latex: ${JSON.stringify(currentAlignmentLines[0])}, mode: "display" },
            { id: "align-mode-2", latex: ${JSON.stringify(currentAlignmentLines[1])}, mode: "display" },
          ],
          activeLineId: "align-mode-2",
        };
        localStorage.setItem(storageKey, JSON.stringify(persisted));
        localStorage.setItem("visualtex-desktop-editor-toolbar-open", "true");
        location.reload();
      })()`);
      await evaluate(`new Promise((resolve) => {
        const done = () => document.querySelectorAll(".formula-line math-field").length === 2
          ? resolve(true)
          : setTimeout(done, 30);
        done();
      })`);
      await sleep(260);
      const modeProbe = await evaluate(`(() => {
        const fields = [...document.querySelectorAll(".formula-line math-field")];
        const lefts = fields.map((field) => {
          for (let position = 0; position <= field.lastOffset; position += 1) {
            const info = field.getElementInfo(position);
            if (info?.latex?.trim() === "=" && info.bounds && info.depth === 0) {
              return info.bounds.left;
            }
          }
          return null;
        });
        const finite = lefts.filter(Number.isFinite);
        return {
          alignment: ${JSON.stringify(alignment)},
          lefts,
          spread: finite.length === 2 ? Math.abs(finite[0] - finite[1]) : null,
        };
      })()`);
      assert.ok(
        typeof modeProbe.spread === "number" && modeProbe.spread <= 0.75,
        `visible equal glyphs were not aligned in ${alignment} mode: ${JSON.stringify(modeProbe)}`,
      );
      alignmentModeResults.push(modeProbe);
    }

    console.log(
      "Align/cases browser regression passed",
      JSON.stringify({
        currentEqualSpread: currentAlignProbe.equalLeafLeftSpread,
        relationProbe,
        alignmentModeResults,
      }),
    );
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(180);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }
}

await main();
