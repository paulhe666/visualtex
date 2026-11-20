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
      "在行首或空格后输入 `/`，可插入表格、图片、代码块和公式；继续输入名称可筛选，`Enter` 插入。表格里按 `Tab` 到下一格。",
      "Type `/` at the start of a line or after a space to insert a table, picture, code block or formula; keep typing to filter and press `Enter`. In a table, `Tab` moves to the next cell.",
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
    title: ["文档保存在浏览器中", "Documents stay in this browser"],
    text: [
      "网页版的文档边写边保存在当前浏览器中，可在文档库中导出 Markdown，主菜单导出 PDF（浏览器的“另存为 PDF”）。清除网站数据会删除未导出的文档；桌面版把每篇文档保存为 .md 文件。",
      "In the web edition documents are saved in this browser as you type; export Markdown from the library and PDF from the main menu (the browser's Save as PDF). Clearing the site's data deletes documents not exported; the desktop app keeps each document as a .md file.",
    ],
    animation: false,
  },
];
