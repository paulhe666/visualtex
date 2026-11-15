// the LaTeX source panel behaves like a code editor (the source is the universal format: the display row sits in
// its $$ block, one indent step deeper)
import assert from "node:assert/strict";

export const names = ["source-editor-ux"];

const sourceUxLatex = [
  String.raw`\begin{align}`,
  String.raw`\frac{\alpha+1}{\beta}&=\int_0^1 x^2\,\mathrm{d}x \\`,
  String.raw`\begin{matrix}`,
  String.raw`a&b \\`,
  String.raw`c&d`,
  String.raw`\end{matrix}&\approx\gamma \\`,
  String.raw`\det A+\left(B\cap C\right)&\to\hbar`,
  String.raw`\end{align}`,
].join("\n");

// state the page starts with (merged into the editor's persisted state)
export const storage = () => ({
  editorLayout: "classic",
  sourceOpen: true,
  zoom: 0.75,
  classicDockHeight: 280,
  lines: [{ id: "source-editor-ux-line", latex: sourceUxLatex, mode: "display" }],
  activeLineId: "source-editor-ux-line",
});
export const localStorageItems = () => ({
  "visualtex-desktop-editor-toolbar-open": "true",
  "visualtex-desktop-editor-source-open": "true",
});

export async function run({ client, evaluate, key, sleep, waitForEvaluation }) {
  await waitForEvaluation(`(() => ({
        ready:
          document.querySelector(".workspace")?.dataset.editorLayout === "classic" &&
          Boolean(document.querySelector(".classic-source-pane-slot .cm-content")) &&
          document.querySelectorAll(".classic-source-pane-slot .cm-line").length >= 10,
      }))()`, "VS Code-like source editor fixture");
  await sleep(500);
  await evaluate(`(() => {
        const laterButton = [...document.querySelectorAll('.office-first-run-backdrop button')]
          .find((button) => /Later|稍后处理/.test(button.textContent || ''));
        if (laterButton instanceof HTMLElement) laterButton.click();
        const content = document.querySelector(".classic-source-pane-slot .cm-content");
        content?.focus();
        return true;
      })()`);
  await sleep(120);

  const sourceUxState = await evaluate(`(() => {
        const pane = document.querySelector(".classic-source-pane-slot");
        const lines = [...(pane?.querySelectorAll(".cm-line") ?? [])].map((line) => {
          const text = line.textContent ?? "";
          return {
            text,
            leadingSpaces: text.match(/^ */)?.[0].length ?? 0,
          };
        });
        const semantic = [...(pane?.querySelectorAll('[class*="cm-vt-command-"]') ?? [])]
          .map((node) => ({
            text: node.textContent ?? "",
            classes: [...node.classList].filter((name) => name.startsWith("cm-vt-command-")),
            color: getComputedStyle(node).color,
          }));
        const firstByClass = {};
        for (const entry of semantic) {
          for (const className of entry.classes) {
            firstByClass[className] ??= entry;
          }
        }
        const guides = [...(pane?.querySelectorAll(".cm-vt-indent-guide") ?? [])];
        const foldGutter = pane?.querySelector(".cm-foldGutter");
        const foldGlyphs = [...(foldGutter?.querySelectorAll(".cm-gutterElement") ?? [])]
          .map((node) => node.textContent ?? "")
          .filter(Boolean);
        return {
          lines,
          firstByClass,
          semanticColorCount: new Set(semantic.map((entry) => entry.color)).size,
          semanticCount: semantic.length,
          guideCount: guides.length,
          guideBackground: guides[0] ? getComputedStyle(guides[0]).backgroundImage : "none",
          foldGutterPresent: Boolean(foldGutter),
          foldGlyphs,
          lineNumberCount: pane?.querySelectorAll(".cm-lineNumbers .cm-gutterElement").length ?? 0,
          activeLineCount: pane?.querySelectorAll(".cm-activeLine").length ?? 0,
          activeGutterCount: pane?.querySelectorAll(".cm-activeLineGutter").length ?? 0,
          focused: pane?.querySelector(".cm-editor")?.classList.contains("cm-focused") ?? false,
          sourceText: pane?.querySelector(".cm-content")?.textContent ?? "",
        };
      })()`);

  const expectedLeading = [0, 2, 4, 4, 6, 6, 4, 4, 2, 0];
  assert.deepEqual(
    sourceUxState.lines.slice(0, 10).map((line) => line.leadingSpaces),
    expectedLeading,
    JSON.stringify(sourceUxState.lines),
  );
  for (const category of [
    "structure",
    "calculus",
    "matrix",
    "greek",
    "relation",
    "set",
    "arrow",
    "physics",
  ]) {
    assert.ok(
      sourceUxState.firstByClass[`cm-vt-command-${category}`],
      `missing semantic source color for ${category}: ${JSON.stringify(sourceUxState.firstByClass)}`,
    );
  }
  assert.ok(
    sourceUxState.semanticColorCount >= 7,
    `source semantic colors are not sufficiently distinct: ${JSON.stringify(sourceUxState.firstByClass)}`,
  );
  assert.ok(sourceUxState.semanticCount >= 12, JSON.stringify(sourceUxState));
  assert.ok(sourceUxState.guideCount >= 5, JSON.stringify(sourceUxState));
  assert.notEqual(sourceUxState.guideBackground, "none", JSON.stringify(sourceUxState));
  assert.equal(sourceUxState.foldGutterPresent, true, JSON.stringify(sourceUxState));
  assert.ok(sourceUxState.foldGlyphs.length >= 2, JSON.stringify(sourceUxState));
  assert.ok(sourceUxState.lineNumberCount >= 8, JSON.stringify(sourceUxState));
  assert.equal(sourceUxState.activeLineCount, 1, JSON.stringify(sourceUxState));
  assert.ok(sourceUxState.activeGutterCount >= 1, JSON.stringify(sourceUxState));
  assert.equal(sourceUxState.focused, true, JSON.stringify(sourceUxState));

  const clickSourceLine = async (text) => {
    const point = await evaluate(`(() => {
          const line = [...document.querySelectorAll(".classic-source-pane-slot .cm-line")]
            .find((candidate) => (candidate.textContent || "").trim() === ${JSON.stringify(text)});
          const rect = line?.getBoundingClientRect();
          return rect ? { x: rect.right - 8, y: rect.top + rect.height / 2 } : null;
        })()`);
    assert.ok(point, `source line not found: ${text}`);
    await client.send("Input.dispatchMouseEvent", {
      type: "mousePressed",
      x: point.x,
      y: point.y,
      button: "left",
      buttons: 1,
      clickCount: 1,
    });
    await client.send("Input.dispatchMouseEvent", {
      type: "mouseReleased",
      x: point.x,
      y: point.y,
      button: "left",
      buttons: 0,
      clickCount: 1,
    });
    await sleep(70);
  };

  const matrixBeginLine = String.raw`\begin{matrix}`;
  await clickSourceLine(matrixBeginLine);
  await key("End", "End", 35);
  await key("Enter", "Enter", 13);
  await sleep(120);
  const enterIndentState = await evaluate(`(() => {
        const lines = [...document.querySelectorAll(".classic-source-pane-slot .cm-line")]
          .map((line) => line.textContent ?? "");
        const beginIndex = lines.findIndex((line) => line.trim() === ${JSON.stringify(matrixBeginLine)});
        const inserted = beginIndex >= 0 ? lines[beginIndex + 1] ?? "" : "";
        return {
          lineCount: lines.length,
          inserted,
          leadingSpaces: inserted.match(/^ */)?.[0].length ?? 0,
        };
      })()`);
  assert.equal(enterIndentState.leadingSpaces, 6, JSON.stringify(enterIndentState));
  // undo is Mod-z: Command on a Mac, Control elsewhere
  const undoModifiers = (await evaluate("/Mac/.test(navigator.platform)")) ? 4 : 2;
  await key("z", "KeyZ", 90, undoModifiers);
  await sleep(120);
  const afterUndoLineCount = await evaluate(`document.querySelectorAll(".classic-source-pane-slot .cm-line").length`);
  assert.equal(afterUndoLineCount, enterIndentState.lineCount - 1, "undo removes the line Enter added");

  await clickSourceLine("c&d");
  await key("Home", "Home", 36);
  await key("Tab", "Tab", 9);
  await sleep(90);
  const tabIndentState = await evaluate(`(() => {
        const line = [...document.querySelectorAll(".classic-source-pane-slot .cm-line")]
          .find((candidate) => (candidate.textContent || "").trim() === "c&d");
        const text = line?.textContent ?? "";
        return { text, leadingSpaces: text.match(/^ */)?.[0].length ?? 0 };
      })()`);
  assert.equal(tabIndentState.leadingSpaces, 8, JSON.stringify(tabIndentState));
  await key("Tab", "Tab", 9, 8);
  await sleep(90);
  const shiftTabIndentState = await evaluate(`(() => {
        const line = [...document.querySelectorAll(".classic-source-pane-slot .cm-line")]
          .find((candidate) => (candidate.textContent || "").trim() === "c&d");
        const text = line?.textContent ?? "";
        return { text, leadingSpaces: text.match(/^ */)?.[0].length ?? 0 };
      })()`);
  assert.equal(shiftTabIndentState.leadingSpaces, 6, JSON.stringify(shiftTabIndentState));

  console.log(
    "VS Code-like LaTeX source editor browser regression passed",
    JSON.stringify({
      leading: sourceUxState.lines.slice(0, 10).map((line) => line.leadingSpaces),
      colors: Object.fromEntries(
        Object.entries(sourceUxState.firstByClass).map(([key, value]) => [key, value.color]),
      ),
      guideCount: sourceUxState.guideCount,
      foldGlyphs: sourceUxState.foldGlyphs,
    }),
  );
  return;
}
