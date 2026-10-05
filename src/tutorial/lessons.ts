// Tutorial lessons. Each lesson has a recording of the real editor
// (public/tutorial/<id>.mp4, made by scripts/tutorial/record_tutorial.mjs)
// and a practice editor (/editor?tutorial=<id>) that starts from `start`.
//
// Text markup: [Ctrl+Enter] renders as key caps, `x^2` as code.

import type { FormulaLineMode, LatexFormatProfile } from "../types/formula";

export type TutorialLang = "zh" | "en";

export interface TutorialSnapshot {
  lines: { latex: string; mode: FormulaLineMode }[];
  profile: LatexFormatProfile;
  copiedText: string;
}

interface Text {
  zh: string;
  en: string;
}

export interface TutorialTask {
  label: Text;
  done: (snapshot: TutorialSnapshot) => boolean;
}

export interface TutorialLesson {
  id: string;
  title: Text;
  summary: Text;
  steps: Text[];
  tasks: TutorialTask[];
  start: { latex: string; mode: FormulaLineMode }[];
}

const filled = (snapshot: TutorialSnapshot) =>
  snapshot.lines.filter((line) => line.latex.trim() !== "");

const isMultiline = (latex: string) => /^\\begin\{(?:aligned|gathered)\}/.test(latex.trim());

// Fragments of the built-in formula tiles (src/toolbar/FormulaToolbar.tsx).
const TILE_FRAGMENTS = [
  "4ac", "i\\pi", "a^2+b^2=c^2", "\\binom", "e^{-x^2}", "f^{(n)}", "mc^2", "\\hbar", "\\varepsilon_0", "\\lambda I",
];

export const TUTORIAL_LESSONS: TutorialLesson[] = [
  {
    id: "typing",
    title: { zh: "输入公式", en: "Type a formula" },
    summary: {
      zh: "点一下公式行，直接打字。",
      en: "Click a formula row and start typing.",
    },
    steps: [
      { zh: "输入 `x^2`：`^` 后面的内容变成上标。", en: "Type `x^2`: whatever follows `^` becomes a superscript." },
      { zh: "按 [→] 离开上标，回到主线。", en: "Press [→] to leave the superscript." },
      { zh: "输入 `\\frac`，按 [Enter] 选中候选，插入分式。", en: "Type `\\frac` and press [Enter] to pick the suggestion." },
      { zh: "填好分子后按 [Tab] 跳到分母。", en: "Fill in the numerator, then press [Tab] to reach the denominator." },
      { zh: "`/` 只是斜杠，不会变成分式。", en: "`/` stays a plain slash; it does not make a fraction." },
    ],
    tasks: [
      { label: { zh: "写一个上标", en: "Write a superscript" }, done: (s) => s.lines.some((line) => line.latex.includes("^")) },
      { label: { zh: "用 \\frac 插入分式", en: "Insert a fraction with \\frac" }, done: (s) => s.lines.some((line) => line.latex.includes("\\frac")) },
    ],
    start: [{ latex: "", mode: "display" }],
  },
  {
    id: "lines",
    title: { zh: "新建行：行内与行间", en: "New rows: inline and display" },
    summary: {
      zh: "行首的 `$` 表示行内公式，`$$` 表示行间公式。",
      en: "The mark at the start of a row: `$` is an inline formula, `$$` a display formula.",
    },
    steps: [
      { zh: "[Enter]：从光标处断开，新建一行同类型的公式。", en: "[Enter]: split at the cursor into a new row of the same type." },
      { zh: "[Ctrl+Enter]：新建一行行内公式。", en: "[Ctrl+Enter]: new inline row." },
      { zh: "[Alt+Enter]：新建一行行间公式。", en: "[Alt+Enter]: new display row." },
      { zh: "点击行首的 `$` 或 `$$`，切换这一行的类型。", en: "Click the `$` or `$$` mark to switch the row's type." },
      { zh: "源码和复制结果里，行内公式写成 `$…$`，行间公式写成 `$$…$$`。", en: "In the source and when copying, inline rows become `$…$` and display rows `$$…$$`." },
    ],
    tasks: [
      { label: { zh: "新建一行行内公式", en: "Add an inline row" }, done: (s) => filled(s).some((line) => line.mode === "inline") },
      { label: { zh: "再新建一行行间公式", en: "Add another display row" }, done: (s) => filled(s).filter((line) => line.mode === "display").length >= 2 },
    ],
    start: [{ latex: "E=mc^2", mode: "display" }],
  },
  {
    id: "multiline",
    title: { zh: "多行公式与对齐", en: "Multi-line formulas and alignment" },
    summary: {
      zh: "一条公式写成多行，用 [Shift+Enter]。",
      en: "To break one formula over several lines, use [Shift+Enter].",
    },
    steps: [
      { zh: "在公式里按 [Shift+Enter]：在同一条公式里换行，它变成多行公式。", en: "Press [Shift+Enter] inside a formula to start a new line within it." },
      { zh: "多行公式总是行间公式。默认格式是 `gather`：各行不对齐。", en: "Multi-line formulas are always display formulas. The default, `gather`, does not align the lines." },
      { zh: "要对齐等号：点顶栏的 LaTeX 格式按钮 `</>`，「多行」选 `align`。", en: "To align the equals signs, open the LaTeX format button `</>` in the top bar and choose `align` under Multi-line." },
      { zh: "在每一行要对齐的位置输入 `&`。", en: "Type `&` where each line should line up." },
    ],
    tasks: [
      { label: { zh: "用 Shift+Enter 写出两行", en: "Write two lines with Shift+Enter" }, done: (s) => s.lines.some((line) => isMultiline(line.latex)) },
      { label: { zh: "把多行格式改成 align", en: "Switch Multi-line to align" }, done: (s) => s.profile.multilineEnvironment === "align" },
      {
        label: { zh: "用 & 让两行对齐", en: "Align two lines with &" },
        done: (s) => s.lines.some((line) => isMultiline(line.latex) && (line.latex.match(/visualtex-align-marker|&/g)?.length ?? 0) >= 2),
      },
    ],
    start: [{ latex: "", mode: "display" }],
  },
  {
    id: "tools",
    title: { zh: "公式工具和磁贴", en: "Formula tools and tiles" },
    summary: {
      zh: "不记得命令时，直接点。",
      en: "When you don't remember a command, click it.",
    },
    steps: [
      { zh: "底部「公式工具」按类别排好了结构和符号，点一下插入到光标处。", en: "The Formula tools panel at the bottom groups structures and symbols; click one to insert it at the cursor." },
      { zh: "右侧的磁贴是整条公式，点一下插入。窗口较窄时，用右上角的按钮打开磁贴栏。", en: "Tiles on the right are whole formulas; click to insert. In a narrow window, open the tile bar with the button at the top right." },
      { zh: "右键工具或磁贴：设为常用，或设置快捷键。", en: "Right-click a tool or tile to pin it to Common or give it a shortcut." },
      { zh: "磁贴栏「自定义」→「保存到…」，把当前公式存成自己的磁贴。", en: "Tiles → Custom → Save to… keeps the current formula as your own tile." },
    ],
    tasks: [
      { label: { zh: "从公式工具插入一个根号", en: "Insert a square root from the tools" }, done: (s) => s.lines.some((line) => line.latex.includes("\\sqrt")) },
      { label: { zh: "插入一个磁贴", en: "Insert a tile" }, done: (s) => s.lines.some((line) => TILE_FRAGMENTS.some((fragment) => line.latex.includes(fragment))) },
    ],
    start: [{ latex: "", mode: "display" }],
  },
  {
    id: "copy",
    title: { zh: "复制 LaTeX", en: "Copy the LaTeX" },
    summary: {
      zh: "写好的公式随时复制成 LaTeX，或者导出。",
      en: "Copy your formulas as LaTeX at any time, or export them.",
    },
    steps: [
      { zh: "底部面板右上角的复制按钮：复制整篇 LaTeX。", en: "The copy button at the top right of the bottom panel copies the whole document as LaTeX." },
      { zh: "「LaTeX 源码」页可以直接看和改源码，两边同步。", en: "The LaTeX source tab shows the source; edits there sync both ways." },
      { zh: "LaTeX 格式按钮 `</>` 决定写法：行内 `$…$` 或 `\\(…\\)`，行间 `$$…$$`、`\\[…\\]` 或 `equation`，以及是否编号。", en: "The LaTeX format button `</>` picks the wrappers: inline `$…$` or `\\(…\\)`, display `$$…$$`, `\\[…\\]` or `equation`, and numbering." },
      { zh: "在公式上右键可以复制 PNG；顶栏的导出按钮导出 Markdown、SVG、PNG。", en: "Right-click a formula to copy it as PNG; the Export button in the top bar saves Markdown, SVG or PNG." },
    ],
    tasks: [
      { label: { zh: "复制一次 LaTeX", en: "Copy the LaTeX once" }, done: (s) => s.copiedText.trim() !== "" },
      {
        label: { zh: "把行间公式改成 \\[…\\]", en: "Switch display formulas to \\[…\\]" },
        done: (s) => s.profile.displayWrapper === "bracket",
      },
    ],
    start: [
      { latex: "a^2+b^2=c^2", mode: "inline" },
      { latex: "\\int_0^1x^2\\,\\mathrm{d}x=\\frac{1}{3}", mode: "display" },
    ],
  },
];

export function findTutorialLesson(id: string | null) {
  return TUTORIAL_LESSONS.find((lesson) => lesson.id === id) ?? null;
}
