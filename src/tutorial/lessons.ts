// Tutorial lessons. Each lesson has a recording of the real editor
// (public/tutorial/<id>.<lang>.mp4), optional screenshots of the real editor
// (public/tutorial/<id>-<figure>.<lang>.png), both made by
// scripts/tutorial/record_tutorial.mjs, and a practice editor
// (/editor?tutorial=<id>) that starts from `start`.
//
// Text markup: [Ctrl+Enter] renders as key caps, `x^2` as code.

import type { FormulaLineMode, LatexFormatProfile } from "../types/formula";

export type TutorialLang = "zh" | "en";

export interface TutorialSnapshot {
  lines: { latex: string; mode: FormulaLineMode }[];
  profile: LatexFormatProfile;
  copiedText: string;
  /** Ids of formula hotkey bindings; built-in ones start with "default:". */
  hotkeyIds: string[];
  customTiles: { latex: string; sectionId: string }[];
  customSectionCount: number;
}

interface Text {
  zh: string;
  en: string;
}

export interface TutorialTask {
  label: Text;
  done: (snapshot: TutorialSnapshot) => boolean;
}

export interface TutorialFigure {
  id: string;
  caption: Text;
}

export interface TutorialLesson {
  id: string;
  title: Text;
  summary: Text;
  steps: Text[];
  figures?: TutorialFigure[];
  tasks: TutorialTask[];
  start: { latex: string; mode: FormulaLineMode }[];
}

const filled = (snapshot: TutorialSnapshot) =>
  snapshot.lines.filter((line) => line.latex.trim() !== "");

const isMultiline = (latex: string) => /^\\begin\{(?:aligned|gathered)\}/.test(latex.trim());

// Aligned rows: the first row has something left of its '&' (a real column).
const isAlignedWithAmpersand = (latex: string) => /^\\begin\{aligned\}\s*[^&\s]/.test(latex.trim()) && latex.includes("&");

// Fragments of the built-in formula tiles (src/toolbar/FormulaToolbar.tsx).
const TILE_FRAGMENTS = [
  "4ac", "i\\pi", "a^2+b^2=c^2", "\\binom", "e^{-x^2}", "f^{(n)}", "mc^2", "\\hbar", "\\varepsilon_0", "\\lambda I",
];

const GAUSS_LAW = String.raw`\oint_{\partial V}\mathbf{E}\cdot\mathrm{d}\mathbf{A}=\frac{Q}{\varepsilon_0}`;

export const TUTORIAL_LESSONS: TutorialLesson[] = [
  {
    id: "typing",
    title: { zh: "输入公式", en: "Type a formula" },
    summary: {
      zh: "点一下公式行，直接打字。",
      en: "Click a formula row and start typing.",
    },
    steps: [
      { zh: "输入 `x^2`：`^` 后面的内容变成上标，`_` 后面的变成下标。", en: "Type `x^2`: what follows `^` becomes a superscript, `_` a subscript." },
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
    title: { zh: "多行公式", en: "Multi-line formulas" },
    summary: {
      zh: "一条公式写成多行，用 [Shift+Enter]。",
      en: "To break one formula over several lines, use [Shift+Enter].",
    },
    steps: [
      { zh: "在公式里按 [Shift+Enter]：在同一条公式里换行，它变成多行公式。", en: "Press [Shift+Enter] inside a formula to start a new line within it." },
      { zh: "和 [Enter] 不同：[Enter] 新建的是另一条公式。", en: "Unlike [Enter], which starts a separate formula." },
      { zh: "多行公式总是行间公式。默认格式是 `gather`，各行不对齐；要对齐等号，看下一课。", en: "Multi-line formulas are always display formulas. The default, `gather`, does not align the lines; the next lesson shows how." },
    ],
    tasks: [
      { label: { zh: "用 Shift+Enter 写出两行", en: "Write two lines with Shift+Enter" }, done: (s) => s.lines.some((line) => isMultiline(line.latex)) },
    ],
    start: [{ latex: "", mode: "display" }],
  },
  {
    id: "align",
    title: { zh: "对齐多行公式", en: "Align a multi-line formula" },
    summary: {
      zh: "先把多行格式改成 `align`，再用 `&` 标出每行对齐的位置。",
      en: "Switch Multi-line to `align`, then mark where each line lines up with `&`.",
    },
    steps: [
      { zh: "点顶栏的 LaTeX 格式按钮 `</>`，「多行」选 `align`。", en: "Open the LaTeX format button `</>` in the top bar and choose `align` under Multi-line." },
      { zh: "在要对齐的符号前面输入 `&`，例如 `y&=x^2+2x+1`。", en: "Type `&` right before the symbol to line up, e.g. `y&=x^2+2x+1`." },
      { zh: "按 [Shift+Enter] 换行，下一行也从 `&` 开始：`&=(x+1)^2`。各行的 `&` 上下对齐。", en: "Press [Shift+Enter] and start the next line with `&` too: `&=(x+1)^2`. The `&`s line up." },
      { zh: "还是 `gather` 时，`&` 会变成普通的 & 字符。", en: "While Multi-line is still `gather`, `&` is typed as a plain & character." },
      { zh: "源码里是 `align*`；在格式菜单里打开「编号」就变成带编号的 `align`。", en: "The source uses `align*`; turn on Number in the format menu for a numbered `align`." },
    ],
    figures: [
      { id: "menu", caption: { zh: "LaTeX 格式菜单：「多行」选 align", en: "LaTeX format menu: Multi-line set to align" } },
    ],
    tasks: [
      { label: { zh: "把多行格式改成 align", en: "Switch Multi-line to align" }, done: (s) => s.profile.multilineEnvironment === "align" },
      { label: { zh: "用 & 让两行对齐", en: "Align two lines with &" }, done: (s) => s.lines.some((line) => isAlignedWithAmpersand(line.latex)) },
    ],
    start: [
      { latex: String.raw`\begin{gathered}y=x^2+2x+1\\ =\left(x+1\right)^2\end{gathered}`, mode: "display" },
      { latex: "", mode: "display" },
    ],
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
      { zh: "上面一排是类别：结构、微积分、矩阵、希腊字母……", en: "The row above switches categories: structures, calculus, matrices, Greek…" },
      { zh: "右侧的磁贴是整条公式，点一下插入。窗口较窄时，用右上角的按钮打开磁贴栏。", en: "Tiles on the right are whole formulas; click to insert. In a narrow window, open the tile bar with the button at the top right." },
    ],
    tasks: [
      { label: { zh: "从公式工具插入一个根号", en: "Insert a square root from the tools" }, done: (s) => s.lines.some((line) => line.latex.includes("\\sqrt")) },
      { label: { zh: "插入一个磁贴", en: "Insert a tile" }, done: (s) => s.lines.some((line) => TILE_FRAGMENTS.some((fragment) => line.latex.includes(fragment))) },
    ],
    start: [{ latex: "", mode: "display" }],
  },
  {
    id: "hotkeys",
    title: { zh: "设置快捷键", en: "Set a shortcut" },
    summary: {
      zh: "常用的工具和磁贴，可以绑定自己的快捷键。",
      en: "Give the tools and tiles you use most their own shortcut.",
    },
    steps: [
      { zh: "右键一个公式工具或磁贴，选「设置快捷键…」。", en: "Right-click a formula tool or tile and choose Set hotkey…." },
      { zh: "按下组合键，要带 [Ctrl] 或 [Alt]（Mac 上也可以用 ⌘）。显示「可以使用」后，点「绑定快捷键」。", en: "Press a combination with [Ctrl] or [Alt] (or ⌘ on a Mac). When it says Available, click Assign hotkey." },
      { zh: "在公式里按这个组合键，就会插入它。", en: "Press it in a formula to insert the item." },
      { zh: "全部快捷键在「设置 → 管理公式快捷键」里，可以修改或删除。", en: "Settings → Manage formula hotkeys lists them all; change or delete them there." },
    ],
    figures: [
      { id: "manager", caption: { zh: "设置 → 管理公式快捷键", en: "Settings → Manage formula hotkeys" } },
    ],
    tasks: [
      { label: { zh: "给一个工具设置快捷键", en: "Give a tool a shortcut" }, done: (s) => s.hotkeyIds.some((id) => !id.startsWith("default:")) },
    ],
    start: [{ latex: "", mode: "display" }],
  },
  {
    id: "custom-tiles",
    title: { zh: "自定义磁贴", en: "Your own tiles" },
    summary: {
      zh: "把常用的公式存成磁贴，以后点一下就插入。",
      en: "Save formulas you reuse as tiles and insert them with one click.",
    },
    steps: [
      { zh: "右侧磁贴栏切到「自定义」。", en: "Switch the tile bar to Custom." },
      { zh: "点「分区」新建一个分区并起名，按 [Enter] 确认。也可以直接用「未命名分区」。", en: "Click Section to add a section, name it and press [Enter]. Or keep the default section." },
      { zh: "点一下要保存的公式行，再点「保存到「分区名」」。", en: "Click the formula row to keep, then click Save to “section”." },
      { zh: "点磁贴插入这条公式。右键磁贴可以设快捷键、换颜色、移到别的分区或删除。", en: "Click the tile to insert it. Right-click it to set a shortcut, change its colour, move it or delete it." },
    ],
    figures: [
      { id: "menu", caption: { zh: "右键自定义磁贴", en: "Right-clicking a custom tile" } },
    ],
    tasks: [
      { label: { zh: "保存一个自定义磁贴", en: "Save a custom tile" }, done: (s) => s.customTiles.length > 0 },
      {
        label: { zh: "点磁贴再插入一次", en: "Insert it again from the tile" },
        done: (s) => s.customTiles.some((tile) => s.lines.filter((line) => line.latex.includes(tile.latex)).length >= 2),
      },
    ],
    start: [
      { latex: GAUSS_LAW, mode: "display" },
      { latex: "", mode: "display" },
    ],
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
    figures: [
      { id: "menu", caption: { zh: "LaTeX 格式菜单（默认设置）", en: "LaTeX format menu (defaults)" } },
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
