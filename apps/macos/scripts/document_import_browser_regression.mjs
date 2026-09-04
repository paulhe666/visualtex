import { spawn } from "node:child_process";
import { rm, writeFile } from "node:fs/promises";
import process from "node:process";
import { boundaryValueDocumentSource } from "./fixtures/boundary_value_document_source.mjs";
import { longPhysicsDocumentSource } from "./fixtures/long_physics_document_source.mjs";

const offset = process.pid % 1000;
const previewPort = 18400 + offset;
const debugPort = 23600 + offset;
const sessionId = "12345678-1234-4234-9234-123456789abc";
const longPhysicsRegression = process.argv.includes("--long-physics");
const boundaryValueRegression = process.argv.includes("--boundary-value");
const literalFallbackRegression = process.argv.includes("--literal-fallback");
const theoremStructureRegression = process.argv.includes("--theorem-structure");
const edgeStructureRegression = process.argv.includes("--edge-structures");
const artifactOutputArgument = process.argv.find((argument) =>
  argument.startsWith("--artifact-output="),
);
const artifactOutputPath = artifactOutputArgument?.slice(
  "--artifact-output=".length,
) ?? "";
if (
  [
    longPhysicsRegression,
    boundaryValueRegression,
    literalFallbackRegression,
    theoremStructureRegression,
    edgeStructureRegression,
  ].filter(Boolean).length > 1
) {
  throw new Error("Choose only one document import fixture regression");
}
const fixtureRegression = longPhysicsRegression || boundaryValueRegression;
const customSettingsRegression =
  !fixtureRegression &&
  !literalFallbackRegression &&
  !theoremStructureRegression &&
  !edgeStructureRegression;
const baseUrl = `http://127.0.0.1:${previewPort}/?view=office-document-import&sessionId=${sessionId}&transport=tauri`;
const chromeProfile = `/tmp/visualtex-document-import-${process.pid}`;
const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function assertWordCompatibleSvg(value, formulaId) {
  const svg = Buffer.from(value, "base64").toString("utf8");
  if (!/(?:fill|stroke)=["']#000000["']/i.test(svg)) {
    throw new Error(`Formula ${formulaId} SVG has no explicit black formula paint`);
  }
  if (
    /currentColor|var\(|(?:fill|stroke|color)\s*[:=]\s*["']?(?:inherit|white|#fff(?:fff)?)/i.test(
      svg,
    )
  ) {
    throw new Error(`Formula ${formulaId} SVG retains a deferred or white paint`);
  }
}

function assertRenderedPngPreview(value, formulaId) {
  if (typeof value !== "string" || !value) {
    throw new Error(`Formula ${formulaId} is missing its PNG compatibility preview`);
  }
  const bytes = Buffer.from(value, "base64");
  if (
    bytes.length < 24 ||
    !bytes.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))
  ) {
    throw new Error(`Formula ${formulaId} has an invalid PNG compatibility preview`);
  }
  const width = bytes.readUInt32BE(16);
  const height = bytes.readUInt32BE(20);
  if (width <= 1 || height <= 1 || bytes.length <= 70) {
    throw new Error(
      `Formula ${formulaId} used a transparent placeholder PNG (${width}x${height}, ${bytes.length} bytes)`,
    );
  }
}

async function waitFor(url, timeoutMs = 15_000) {
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
    this.events = [];
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (!message.id) {
        if (
          message.method === "Runtime.exceptionThrown" ||
          message.method === "Runtime.consoleAPICalled"
        ) {
          this.events.push(message);
        }
        return;
      }
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
    await waitFor(`http://127.0.0.1:${previewPort}`);
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
        "about:blank",
      ],
      { stdio: "ignore" },
    );

    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const page = targets.find((target) => target.type === "page");
    if (!page) throw new Error("No Chrome page target found");

    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");
    await client.send("Page.addScriptToEvaluateOnNewDocument", {
      source: `(() => {
        let callbackId = 1;
        const callbacks = new Map();
        window.__VISUALTEX_DOCUMENT_IMPORT_CALLS__ = [];
        Object.defineProperty(HTMLCanvasElement.prototype, "toBlob", {
          configurable: true,
          value(callback) {
            callback(null);
          },
        });
        window.__TAURI_INTERNALS__ = {
          metadata: {
            currentWindow: { label: "office-native-document-test" },
            currentWebview: { label: "office-native-document-test" },
          },
          transformCallback(callback, once = false) {
            const id = callbackId++;
            callbacks.set(id, { callback, once });
            return id;
          },
          unregisterCallback(id) {
            callbacks.delete(id);
          },
          async invoke(command, args) {
            window.__VISUALTEX_DOCUMENT_IMPORT_CALLS__.push({ command, args });
            if (command === "get_macos_offline_document_import_request") {
              return {
                protocolVersion: 1,
                sessionId: ${JSON.stringify(sessionId)},
                operation: "documentImport",
                host: "word",
                sourceDocumentId: "visualtex-word-test-document",
                bookmarkName: "VT_D_12345678123442349234",
                defaultFontSizePt: 12,
              };
            }
            if (command === "focus_macos_offline_document_import_target") {
              window.__VISUALTEX_DOCUMENT_IMPORT_TARGET_FOCUSED__ = true;
              return null;
            }
            if (command === "restore_macos_offline_document_import_window") return null;
            if (command === "get_macos_offline_document_import_progress") {
              return { current: 0, total: 0, stage: "preparing" };
            }
            if (command === "commit_macos_offline_document_import") {
              window.__VISUALTEX_DOCUMENT_IMPORT_COMMIT__ = args;
              return null;
            }
            if (command === "cancel_macos_offline_document_import") return null;
            if (command === "close_macos_offline_office_editor_window") {
              window.__VISUALTEX_DOCUMENT_IMPORT_CLOSED__ = true;
              return null;
            }
            if (command === "plugin:event|listen" || command === "plugin:event|unlisten") {
              return 1;
            }
            throw new Error("Unexpected fake Tauri command: " + command);
          },
        };
      })();`,
    });
    await client.send("Page.navigate", { url: baseUrl });

    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", {
        expression,
        awaitPromise: true,
        returnByValue: true,
      });
      if (result.exceptionDetails) {
        throw new Error(
          result.exceptionDetails.exception?.description ??
            result.exceptionDetails.text ??
            "Browser evaluation failed",
        );
      }
      return result.result.value;
    };

    const started = Date.now();
    while (Date.now() - started < 15_000) {
      const ready = await evaluate(
        `Boolean(document.querySelector(".doc-import-shell"))`,
      );
      if (ready) break;
      await sleep(80);
    }
    if (!(await evaluate(`Boolean(document.querySelector(".doc-import-shell"))`))) {
      const failure = await evaluate(`(() => ({
        text: document.body.innerText,
        html: document.getElementById("root")?.innerHTML?.slice(0, 2000) ?? "",
        calls: window.__VISUALTEX_DOCUMENT_IMPORT_CALLS__ ?? [],
      }))()`);
      const events = client.events.map((event) => ({
        method: event.method,
        exception:
          event.params?.exceptionDetails?.exception?.description ??
          event.params?.exceptionDetails?.text ??
          null,
        console:
          event.params?.args?.map((arg) => arg.value ?? arg.description ?? "") ??
          null,
      }));
      throw new Error(
        `Document importer did not mount: ${JSON.stringify({ failure, events })}`,
      );
    }

    const source = String.raw`正文中的行内公式 $p=mv$ 保持基线对齐。

\begin{equation}
E=mc^2
\end{equation}

\begin{align}
a &= b + c \\
d &= e
\end{align}

\begin{align*}
x &= y \\
y &= z
\end{align*}

结尾文字。`;
    const literalFallbackSource = String.raw`\documentclass{article}
\usepackage{custompkg}
\newcommand{\customsymbol}[1]{\mathbf{#1}}
\begin{document}
标准公式 $x=1$。
\[
\customsymbol{q}
\]
\end{document}`;
    const theoremStructureSource = String.raw`\newtheorem{thm}{定理}
\newtheorem{lem}[thm]{引理}
\begin{thm}[谱定理]
正文公式 $Av=\lambda v$。
\[
A=\begin{pmatrix}1&0\\0&2\end{pmatrix}
\]
\end{thm}
\begin{lem}
共享计数器正文。
\end{lem}
\begin{proof}[充分性]
由 $x=1$ 立即得到。\qedhere
\end{proof}`;
    const edgeStructureSource = String.raw`这就是色散媒质中频域形式的本构关系。

% ==================== % 6. 磁色散媒质中的本构关系 %
====================

对于良导体低频近似，若 \(\varepsilon_r(\omega)\) 的本征极化部分可以忽略，则

\begin{equation}
\begin{aligned}
f^{*}(\mathbf{x})
&=
\frac{1}{p(\mathbf{x})}
\int t\,p(\mathbf{x},t)\,\mathrm{d}t  \\
&=
\int t\,p(t\mid\mathbf{x})\,\mathrm{d}t
=
\mathbb{E}_{t}[t\mid\mathbf{x}]
\end{aligned}
\end{equation}`;
    const regressionSource = edgeStructureRegression
      ? edgeStructureSource
      : theoremStructureRegression
        ? theoremStructureSource
        : literalFallbackRegression
        ? literalFallbackSource
        : boundaryValueRegression
          ? boundaryValueDocumentSource
          : longPhysicsRegression
            ? longPhysicsDocumentSource
            : source;
    await evaluate(`(() => {
      const textarea = document.querySelector(".source-pane textarea");
      if (!textarea) throw new Error("Missing document import source textarea");
      const setter = Object.getOwnPropertyDescriptor(
        HTMLTextAreaElement.prototype,
        "value",
      ).set;
      setter.call(textarea, ${JSON.stringify(regressionSource)});
      textarea.dispatchEvent(new Event("input", { bubbles: true }));
    })()`);

    const expectedFormulaCount = edgeStructureRegression
      ? 2
      : theoremStructureRegression
        ? 3
        : literalFallbackRegression
        ? 2
        : boundaryValueRegression
          ? 22
          : longPhysicsRegression
            ? 16
            : 4;
    const parseStarted = Date.now();
    while (Date.now() - parseStarted < 10_000) {
      if ((await evaluate(`document.querySelectorAll(".document-import-formula-card").length`)) === expectedFormulaCount) {
        break;
      }
      await sleep(80);
    }

    const parsed = await evaluate(`(() => {
      const cards = [...document.querySelectorAll(".document-import-formula-card")];
      return {
        count: cards.length,
        modes: cards.map((card) => card.querySelector("select")?.value),
        numbered: cards.map((card) => card.querySelector('input[type="checkbox"]')?.checked ?? false),
        summary: document.querySelector(".document-import-summary")?.innerText ?? "",
      };
    })()`);
    if (edgeStructureRegression) {
      if (
        parsed.count !== 2 ||
        parsed.modes.join(",") !== "inline,block" ||
        parsed.numbered.join(",") !== "false,true"
      ) {
        throw new Error(`Unexpected edge structure blocks: ${JSON.stringify(parsed)}`);
      }
    } else if (theoremStructureRegression) {
      if (
        parsed.count !== 3 ||
        parsed.modes.join(",") !== "inline,block,inline" ||
        parsed.numbered.some(Boolean)
      ) {
        throw new Error(`Unexpected theorem structure blocks: ${JSON.stringify(parsed)}`);
      }
    } else if (literalFallbackRegression) {
      if (
        parsed.count !== 2 ||
        parsed.modes.join(",") !== "inline,block" ||
        parsed.numbered.some(Boolean)
      ) {
        throw new Error(`Unexpected literal fallback blocks: ${JSON.stringify(parsed)}`);
      }
    } else if (fixtureRegression) {
      if (
        parsed.count !== expectedFormulaCount ||
        parsed.modes.filter((mode) => mode === "inline").length !== 4 ||
        parsed.numbered.some(Boolean)
      ) {
        throw new Error(`Unexpected fixture formula blocks: ${JSON.stringify(parsed)}`);
      }
    } else if (
      parsed.count !== 4 ||
      parsed.modes.join(",") !== "inline,block,block,block" ||
      parsed.numbered.join(",") !== "false,true,true,false"
    ) {
      throw new Error(`Unexpected parsed formula blocks: ${JSON.stringify(parsed)}`);
    }

    const primaryButtonAppearance = await evaluate(`(() => {
      const button = document.querySelector(".doc-import-primary");
      if (!button) throw new Error("Missing primary document import button");
      const probe = document.createElement("span");
      probe.style.color = "var(--accent-primary, var(--accent))";
      document.body.appendChild(probe);
      const appearance = {
        disabled: button.disabled,
        backgroundColor: getComputedStyle(button).backgroundColor,
        accentColor: getComputedStyle(probe).color,
      };
      probe.remove();
      return appearance;
    })()`);
    if (
      primaryButtonAppearance.disabled ||
      primaryButtonAppearance.backgroundColor !== primaryButtonAppearance.accentColor
    ) {
      throw new Error(
        `Document import primary button lost its visible accent background: ${JSON.stringify(primaryButtonAppearance)}`,
      );
    }

    if (customSettingsRegression) {
      await evaluate(`(() => {
        const globalToggle = document.querySelector(
          ".doc-import-global-number-toggle input[type='checkbox']",
        );
        if (!globalToggle) throw new Error("Missing global display-formula numbering toggle");
        if (globalToggle.checked) {
          throw new Error("Mixed formula numbering must not present as globally enabled");
        }
        globalToggle.click();
      })()`);
      await sleep(100);
      const globalNumberingState = await evaluate(`(() => ({
        globalChecked: document.querySelector(
          ".doc-import-global-number-toggle input[type='checkbox']",
        )?.checked ?? false,
        displayNumbered: [...document.querySelectorAll(
          ".document-import-formula-card.is-block input[type='checkbox']",
        )].map((input) => input.checked),
        inlineCheckboxCount: document.querySelectorAll(
          ".document-import-formula-card.is-inline input[type='checkbox']",
        ).length,
      }))()`);
      if (
        !globalNumberingState.globalChecked ||
        globalNumberingState.displayNumbered.length !== 3 ||
        globalNumberingState.displayNumbered.some((numbered) => !numbered) ||
        globalNumberingState.inlineCheckboxCount !== 0
      ) {
        throw new Error(
          `Global display-formula numbering did not select every display formula: ${JSON.stringify(globalNumberingState)}`,
        );
      }
    }

    await evaluate(`(() => {
      if (!${JSON.stringify(longPhysicsRegression)}) {
        const outputSelect = document.querySelector(
          ".doc-import-options label:nth-child(2) select",
        );
        if (!outputSelect) throw new Error("Missing formula output select");
        const selectSetter = Object.getOwnPropertyDescriptor(
          HTMLSelectElement.prototype,
          "value",
        ).set;
        selectSetter.call(outputSelect, "image");
        outputSelect.dispatchEvent(new Event("input", { bubbles: true }));
        outputSelect.dispatchEvent(new Event("change", { bubbles: true }));
      }
      if (${JSON.stringify(customSettingsRegression)}) {
        const cards = [...document.querySelectorAll(".document-import-formula-card")];
        const sizes = [10.5, 18, 14, 16];
        cards.forEach((card, index) => {
          const input = card.querySelector('input[type="number"]');
          const setter = Object.getOwnPropertyDescriptor(
            HTMLInputElement.prototype,
            "value",
          ).set;
          setter.call(input, String(sizes[index]));
          input.dispatchEvent(new Event("input", { bubbles: true }));
          input.dispatchEvent(new Event("change", { bubbles: true }));
        });
      }
      const insertButton = [...document.querySelectorAll("button")].find(
        (button) =>
          button.textContent?.includes("插入 Word") ||
          button.textContent?.includes("导入到 Word"),
      );
      if (!insertButton) throw new Error("Missing insert button");
      insertButton.click();
    })()`);

    const commitStarted = Date.now();
    let commit;
    while (Date.now() - commitStarted < 30_000) {
      commit = await evaluate(`window.__VISUALTEX_DOCUMENT_IMPORT_COMMIT__ ?? null`);
      if (commit) break;
      const error = await evaluate(
        `document.querySelector(".doc-import-messages .error")?.innerText ?? ""`,
      );
      if (error) throw new Error(`Document import UI reported: ${error}`);
      await sleep(100);
    }
    if (!commit) throw new Error("Document importer did not submit its Tauri commit");
    const targetFocused = await evaluate(
      `window.__VISUALTEX_DOCUMENT_IMPORT_TARGET_FOCUSED__ === true`,
    );
    if (!targetFocused) {
      throw new Error("Document importer did not return focus to Word before commit");
    }

    const input = commit.input;
    const formulas = input?.items?.filter((item) => item.kind === "formula") ?? [];
    const texts = input?.items?.filter((item) => item.kind === "text") ?? [];
    const expectedOutputKind = longPhysicsRegression ? "omml" : "image";
    const expectedCommittedFormulaCount = expectedFormulaCount;
    if (
      input?.outputKind !== expectedOutputKind ||
      formulas.length !== expectedCommittedFormulaCount ||
      texts.length < 2
    ) {
      throw new Error(`Unexpected document import commit: ${JSON.stringify(commit)}`);
    }
    if (customSettingsRegression && (
      formulas[0].displayMode !== "inline" ||
      formulas[0].numbered !== false ||
      formulas[0].fontSizePt !== 10.5 ||
      formulas[1].displayMode !== "block" ||
      formulas[1].numbered !== true ||
      formulas[1].fontSizePt !== 18 ||
      formulas[2].displayMode !== "block" ||
      formulas[2].numbered !== true ||
      formulas[2].fontSizePt !== 14 ||
      formulas[3].displayMode !== "block" ||
      formulas[3].numbered !== true ||
      formulas[3].fontSizePt !== 16
    )) {
      throw new Error(`Independent formula settings were lost: ${JSON.stringify(formulas)}`);
    }
    for (const formula of formulas) {
      const validCommon =
        formula.formulaId &&
        formula.metadata &&
        formula.latex === formula.metadata.latex &&
        Array.isArray(formula.metadata.lines) &&
        formula.metadata.lines.length > 0 &&
        formula.ommlBase64 &&
        formula.ommlDocxBase64;
      const validOutput = longPhysicsRegression
        ? !formula.svgBase64 && !formula.pngBase64
        : formula.svgBase64 &&
          formula.pngBase64 &&
          formula.width > 0 &&
          formula.height > 0;
      if (!validCommon || !validOutput) {
        throw new Error(
          `Document formula regression payload is invalid: ${JSON.stringify(formula)}`,
        );
      }
      if (!longPhysicsRegression) {
        assertWordCompatibleSvg(formula.svgBase64, formula.formulaId);
        assertRenderedPngPreview(formula.pngBase64, formula.formulaId);
      }
    }
    if (new Set(formulas.map((formula) => formula.formulaId)).size !== formulas.length) {
      throw new Error("Imported formulas did not receive independent identities");
    }
    if (edgeStructureRegression) {
      const allText = texts.map((item) => item.text).join("\n");
      if (
        allText.includes("%") ||
        allText.includes("====================") ||
        allText.includes("磁色散媒质中的本构关系")
      ) {
        throw new Error(
          `LaTeX comments leaked into the commit: ${JSON.stringify(texts)}`,
        );
      }
      const beforeInline = texts.find(
        (item) => item.text === "对于良导体低频近似，若",
      );
      const afterInline = texts.find(
        (item) => item.text === "的本征极化部分可以忽略，则",
      );
      const inlineFormula = formulas.find((formula) =>
        formula.latex.includes(String.raw`\varepsilon_r(\omega)`),
      );
      const alignedFormula = formulas.find((formula) =>
        formula.latex.includes(String.raw`\begin{aligned}`),
      );
      if (
        !beforeInline ||
        !afterInline ||
        !inlineFormula ||
        beforeInline.paragraphId !== inlineFormula.paragraphId ||
        afterInline.paragraphId !== inlineFormula.paragraphId ||
        beforeInline.paragraphStart !== true ||
        inlineFormula.paragraphStart !== false ||
        inlineFormula.paragraphEnd !== false ||
        afterInline.paragraphEnd !== true
      ) {
        throw new Error(
          `Inline formula CJK boundaries retained spacing or lost paragraph identity: ${JSON.stringify({ texts, formulas })}`,
        );
      }
      if (
        !alignedFormula ||
        alignedFormula.displayMode !== "block" ||
        alignedFormula.numbered !== true ||
        alignedFormula.metadata?.codeFormat !== "equation" ||
        alignedFormula.metadata?.lines?.length !== 1 ||
        !alignedFormula.latex.startsWith("\\begin{equation}\n") ||
        !alignedFormula.latex.includes("\\begin{aligned}") ||
        alignedFormula.width < 240 ||
        alignedFormula.width > 360 ||
        alignedFormula.height > 130
      ) {
        throw new Error(
          `Nested aligned equation was renormalized or rendered vertically: ${JSON.stringify(alignedFormula)}`,
        );
      }
    }
    if (theoremStructureRegression) {
      const findText = (value, style, listKind = "none") =>
        texts.find(
          (item) =>
            item.text.includes(value) &&
            item.paragraphStyle === style &&
            item.listKind === listKind,
        );
      for (const [value, style] of [
        [String.raw`\newtheorem{thm}{定理}`, "code"],
        [String.raw`\newtheorem{lem}[thm]{引理}`, "code"],
        ["定理 1（谱定理）", "heading4"],
        ["引理 2", "heading4"],
        ["证明（充分性）", "heading4"],
        ["正文公式", "quote"],
        ["共享计数器正文", "quote"],
        ["由", "normal"],
        ["□", "normal"],
      ]) {
        if (!findText(value, style)) {
          throw new Error(
            `Theorem structure commit lost ${style} text ${value}: ${JSON.stringify(texts)}`,
          );
        }
      }
      const theoremInline = formulas.find((formula) =>
        formula.latex.includes(String.raw`Av=\lambda v`),
      );
      const theoremDisplay = formulas.find((formula) =>
        formula.latex.includes(String.raw`\begin{pmatrix}`),
      );
      const proofInline = formulas.find((formula) => formula.latex.includes("x=1"));
      if (
        theoremInline?.paragraphStyle !== "quote" ||
        theoremInline?.displayMode !== "inline" ||
        theoremDisplay?.displayMode !== "block" ||
        theoremDisplay?.paragraphId ||
        proofInline?.paragraphStyle !== "normal" ||
        proofInline?.displayMode !== "inline"
      ) {
        throw new Error(
          `Theorem formula paragraph metadata is invalid: ${JSON.stringify(formulas)}`,
        );
      }
    }
    if (literalFallbackRegression) {
      const literalText = texts.map((item) => item.text).join("\n");
      for (const expected of [
        String.raw`\documentclass{article}`,
        String.raw`\usepackage{custompkg}`,
        String.raw`\newcommand{\customsymbol}[1]{\mathbf{#1}}`,
      ]) {
        if (!literalText.includes(expected)) {
          throw new Error(
            `Literal fallback commit lost unsupported source ${expected}: ${JSON.stringify(texts)}`,
          );
        }
      }
      if (
        formulas[0]?.latex !== "x=1" ||
        formulas[1]?.latex !== "\\mathbf{q}"
      ) {
        throw new Error(
          `Supported formulas beside literal fallback were not converted: ${JSON.stringify(formulas)}`,
        );
      }
    }
    if (boundaryValueRegression) {
      const singleEquationFormulas = formulas.filter((formula) =>
        ["equation", "equation-star"].includes(formula.metadata?.codeFormat),
      );
      if (singleEquationFormulas.length < 8) {
        throw new Error(
          `Boundary-value regression did not preserve its single-equation environments: ${JSON.stringify(singleEquationFormulas)}`,
        );
      }
      for (const formula of singleEquationFormulas) {
        const environment =
          formula.metadata.codeFormat === "equation" ? "equation" : "equation*";
        const opening = `\\begin{${environment}}`;
        const closing = `\\end{${environment}}`;
        if (
          formula.metadata.lines.length !== 1 ||
          formula.latex.split(opening).length !== 2 ||
          formula.latex.split(closing).length !== 2 ||
          formula.latex.includes("\\begin{aligned}") ||
          formula.latex.includes("&")
        ) {
          throw new Error(
            `Single-equation source newlines were converted into alignment rows: ${JSON.stringify(formula)}`,
          );
        }
      }
    }
    const multilineExpectations = customSettingsRegression ? [
      {
        formula: formulas[2],
        codeFormat: "align",
        lines: ["a = b + c", "d = e"],
        environment: "align",
      },
      {
        formula: formulas[3],
        codeFormat: "align-star",
        lines: ["x = y", "y = z"],
        environment: "align*",
      },
    ] : [];
    for (const expectation of multilineExpectations) {
      const metadataLines = expectation.formula.metadata.lines.map(
        (line) => line.latex,
      );
      if (
        expectation.formula.metadata.codeFormat !== expectation.codeFormat ||
        JSON.stringify(metadataLines) !== JSON.stringify(expectation.lines) ||
        expectation.formula.latex !== expectation.formula.metadata.latex ||
        !expectation.formula.latex.startsWith(
          `\\begin{${expectation.environment}}\n`,
        ) ||
        !expectation.formula.latex.endsWith(
          `\n\\end{${expectation.environment}}`,
        )
      ) {
        throw new Error(
          `Multiline formula canonical metadata is inconsistent: ${JSON.stringify(expectation.formula)}`,
        );
      }
    }

    if (artifactOutputPath) {
      await writeFile(
        artifactOutputPath,
        JSON.stringify(
          {
            schema: "visualtex-word-browser-artifacts-v1",
            outputKind: input.outputKind,
            formulas,
            texts,
          },
          null,
          2,
        ),
        { mode: 0o600 },
      );
    }

    const closed = await evaluate(
      `window.__VISUALTEX_DOCUMENT_IMPORT_CLOSED__ === true`,
    );
    if (!closed) throw new Error("Document importer did not request window close");

    console.log(JSON.stringify({ parsed, outputKind: input.outputKind, formulas: formulas.map((formula) => ({
      formulaId: formula.formulaId,
      displayMode: formula.displayMode,
      numbered: formula.numbered,
      fontSizePt: formula.fontSizePt,
      hasOmml: Boolean(formula.ommlBase64),
      hasSvg: Boolean(formula.svgBase64),
      hasPng: Boolean(formula.pngBase64),
    })) }, null, 2));
    console.log("Document import browser regression passed");
  } finally {
    client?.close();
    chrome?.kill("SIGTERM");
    preview.kill("SIGTERM");
    await sleep(500);
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }
}

await main();
