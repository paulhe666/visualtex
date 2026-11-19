// Document mode's tutorial cards: one task each; a card with an animation shows tutorial/<zh|en>/<id>.webp (the page
// only, recorded from the app in document mode). Text between backticks is drawn as a key.

export interface TutorialCard {
  id: string;
  /** [zh, en] */
  title: [string, string];
  text: [string, string];
  animation: boolean;
  /** a file path shown in place of an animation */
  path?: [string, string];
}

export const TUTORIAL_CARDS: TutorialCard[] = [
  {
    id: "formulas",
    title: ["输入公式", "Formulas"],
    text: [
      "输入 `$$` 开始行内公式，再输入 `$$` 结束；行间公式用 `$$$$`。在公式里输入 LaTeX 命令会出现候选，`Tab` 进入下一个位置。",
      "Type `$$` to start an inline formula and `$$` again to end it; use `$$$$` for a display formula. LaTeX commands bring up candidates as you type; `Tab` moves to the next slot.",
    ],
    animation: true,
  },
  {
    id: "markdown",
    title: ["Markdown 排版", "Markdown"],
    text: [
      "行首输入 `#` 和空格是标题，`-` 和空格是列表；`**粗体**`、`*斜体*` 照常输入。标记只在光标所在的行显示。",
      "At the start of a line, `#` and a space make a heading and `-` and a space a list; type `**bold**` and `*italic*` as usual. The marks show only on the line with the caret.",
    ],
    animation: true,
  },
  {
    id: "insert-menu",
    title: ["插入菜单", "The / menu"],
    text: [
      "在行首或空格后输入 `/`，可插入表格、图片、代码块和公式；继续输入名称可筛选，`↩` 插入。表格里按 `Tab` 到下一格。",
      "Type `/` at the start of a line or after a space to insert a table, picture, code block or formula; keep typing to filter and press `↩`. In a table, `Tab` moves to the next cell.",
    ],
    animation: true,
  },
  {
    id: "numbers",
    title: ["公式编号与引用", "Numbers and references"],
    text: [
      "在 `/` 菜单中选“编号公式”插入带编号的行间公式，选“引用公式”插入对它的引用。编号变化时，引用随之更新。",
      "In the `/` menu, Numbered formula inserts a display formula with a number, and Reference to equation refers to it. References follow when the numbers change.",
    ],
    animation: true,
  },
  {
    id: "files",
    title: ["文档即文件", "Documents are files"],
    text: [
      "每篇文档都是“文稿/VisualTeX”中的一个 .md 文件，边写边自动保存；其他 .md 文件也能用 VisualTeX 打开。主菜单可导出 PDF。",
      "Each document is a .md file in Documents/VisualTeX, saved as you type, and other .md files open in VisualTeX too. Export to PDF from the main menu.",
    ],
    animation: false,
    path: ["文稿/VisualTeX/实验记录.md", "Documents/VisualTeX/Lab notes.md"],
  },
];
