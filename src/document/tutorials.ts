// Document mode's tutorials: one short screen recording per task, made from the app itself by the tutorial
// recorder (Chinese and English, each platform its own keys); the videos are public/tutorials/<zh|en>/<id>.mp4.
// Generated from the recorder's clip list; edit the clips there and regenerate.

export interface Tutorial {
  id: string;
  /** [zh, en] */
  title: [string, string];
  summary: [string, string];
}

export interface TutorialGroup {
  title: [string, string];
  tutorials: Tutorial[];
}

export const TUTORIAL_GROUPS: TutorialGroup[] = [
  {
    title: ["公式", "Formulas"],
    tutorials: [
      {
        id: "inline-formula",
        title: ["行内公式", "Inline formulas"],
        summary: [
          "输入 $ 在文字中插入公式，像在公式编辑器里一样输入，按 Tab 回到文字；点一下公式即可再次编辑。",
          "Type $ to put a formula in the text, write it as in the formula editor and press Tab to go back to the text; click a formula to edit it again.",
        ],
      },
      {
        id: "display-formula",
        title: ["行间公式与编号", "Display formulas and numbers"],
        summary: [
          "\\[ 开始一个行间公式；\\begin{equation} 等环境自动编号；⇧↩ 换行，↩ 结束。",
          "\\[ starts a display formula; \\begin{equation} and friends are numbered; ⇧↩ adds a line, ↩ ends it.",
        ],
      },
      {
        id: "references",
        title: ["公式引用", "Referring to formulas"],
        summary: [
          "输入 \\eqref（或 \\ref）从列表里选一个编号公式；插入新公式后，引用的编号自动跟着变。",
          "Type \\eqref (or \\ref) and pick a numbered formula; when formulas are added, the references follow.",
        ],
      },
      {
        id: "command-menu",
        title: ["反斜杠命令", "Backslash commands"],
        summary: [
          "在正文里输入 \\ 和命令名，菜单列出章节、列表、加粗、公式环境等；↑↓ 选择，↩ 或数字键执行。",
          "Type \\ and a command name in the text: the menu offers sections, lists, bold, formula environments…; ↑↓ to choose, ↩ or a digit to run it.",
        ],
      },
      {
        id: "formula-tools",
        title: ["公式工具面板", "Formula tools"],
        summary: [
          "底部的公式工具按类别列出结构和符号，点一下就插入到正在编辑的公式里；右侧磁贴是常用公式。",
          "The formula tools at the bottom list structures and symbols by category: a click inserts into the formula being edited; the tiles on the right hold whole formulas.",
        ],
      },
    ],
  },
  {
    title: ["排版", "Writing and layout"],
    tutorials: [
      {
        id: "text-format",
        title: ["文字格式与段落", "Text formatting and paragraphs"],
        summary: [
          "选中文字后用工具栏或快捷键加粗、倾斜、着色；章节、列表、对齐都在工具栏上。",
          "Select text and use the toolbar or keys for bold, italic and colour; sections, lists and alignment are in the toolbar too.",
        ],
      },
      {
        id: "tables",
        title: ["表格", "Tables"],
        summary: [
          "工具栏“表格”选行列数；Tab 跳到下一格，最后一格按 Tab 新增一行；单元格里也能写公式。",
          "Pick the size from Table in the toolbar; Tab goes to the next cell and adds a row from the last one; cells take formulas too.",
        ],
      },
      {
        id: "images",
        title: ["插入图片", "Pictures"],
        summary: [
          "粘贴或拖入图片；拖动两侧的手柄调整大小；选中图片后用浮动工具条设置对齐或文字环绕。",
          "Paste or drop a picture; drag its side handles to resize; with the picture selected, the floating bar sets alignment or text wrap.",
        ],
      },
      {
        id: "image-board",
        title: ["图片自由排版", "Arranging pictures freely"],
        summary: [
          "一行里放多张图片后，可以像 PPT 一样拖动排列：自动吸附对齐线；Shift 多选后用工具条对齐、等距、排成网格。",
          "With several pictures in a line, drag them as in PowerPoint: they snap to guides; Shift-click several and align, distribute or tidy them from the bar.",
        ],
      },
      {
        id: "code-blocks",
        title: ["代码块与流程图", "Code blocks and diagrams"],
        summary: [
          "工具栏“代码块”或 ``` 插入代码块，按语言高亮；语言选 mermaid 时直接画出流程图。",
          "Code block in the toolbar (or ```) adds highlighted code; with the mermaid language it draws the diagram.",
        ],
      },
      {
        id: "footnotes-citations",
        title: ["脚注与参考文献", "Footnotes and references"],
        summary: [
          "⌘⌥F 插入脚注，脚注排在页面底部；“引用”按钮新建或选择参考文献，文末自动生成参考文献列表。",
          "⌘⌥F adds a footnote at the foot of the page; Citation adds or picks a reference, and the list at the end follows.",
        ],
      },
    ],
  },
  {
    title: ["Markdown", "Markdown"],
    tutorials: [
      {
        id: "markdown",
        title: ["Markdown 输入", "Markdown input"],
        summary: [
          "切换到 Markdown 输入后，输入方式与 Typora 相同：**加粗**、- 列表、# 标题（回车）、$$ 公式、| 表格 | 都会转换；/ 打开命令菜单。",
          "In Markdown input you type as in Typora: **bold**, - lists, # headings (on ↩), $$ formulas and | tables | turn into the real thing; / opens the menu.",
        ],
      },
      {
        id: "md-blocks",
        title: ["Markdown：标题、列表与引用", "Markdown: headings, lists and quotes"],
        summary: [
          "行首输入 -、1.、> 加空格，得到无序列表、有序列表和引用；# 标题 回车或离开这一行变成标题；- [ ] 是待办事项，--- 回车是分隔线。",
          "At a line start, -, 1. and > with a space make bullet and numbered lists and quotes; # Title becomes a heading on ↩ or when you leave the line; - [ ] is a to-do item, --- and ↩ a rule.",
        ],
      },
      {
        id: "md-inline",
        title: ["Markdown：粗体、斜体与链接", "Markdown: bold, italic and links"],
        summary: [
          "**粗体**、*斜体*、~~删除线~~、==高亮==、`代码` 输完结束符号即转换，光标回到其中时显示标记，可直接修改；[文字](网址) 变成链接；\\* 输入字面字符。",
          "**bold**, *italic*, ~~strike~~, ==highlight== and `code` convert when you type the closing mark and show their marks again when the caret is back in them; [text](url) becomes a link; \\* types the character itself.",
        ],
      },
      {
        id: "md-math",
        title: ["Markdown：公式", "Markdown: formulas"],
        summary: [
          "$ 后紧跟公式内容开始行内公式（$5 这样的金额保持为文字）；$$ 开始行间公式；斜杠菜单里有带编号的公式和公式引用。",
          "$ followed by the formula starts an inline formula ($5 stays a price); $$ starts a display formula; the / menu has numbered formulas and references.",
        ],
      },
      {
        id: "md-tables-code",
        title: ["Markdown：表格与代码", "Markdown: tables and code"],
        summary: [
          "一行 | a | b | 回车变成表格，表格里回车到下一行；```语言 回车开始代码块；⌘T 直接插入表格。",
          "A line | a | b | and ↩ becomes a table, ↩ in it goes down; ```language and ↩ starts code; ⌘T inserts a table.",
        ],
      },
      {
        id: "md-slash",
        title: ["Markdown：斜杠菜单", "Markdown: the / menu"],
        summary: [
          "/ 打开菜单，可用英文、中文或拼音首字母查找：标题、表格、公式、脚注、图片、分隔线……",
          "/ opens the menu; find entries in English, Chinese or pinyin initials: headings, tables, formulas, footnotes, pictures, rules…",
        ],
      },
      {
        id: "md-keys",
        title: ["Markdown：Typora 快捷键", "Markdown: Typora's keys"],
        summary: [
          "Markdown 输入支持 Typora 的快捷键：⌘1–6 标题，⌘0 正文，⌘B/I/U，⌘K 链接，⌘\\ 清除格式，⌥⌘Q 引用，⌥⌘O / ⌥⌘U 列表，⌥⌘C 代码块，⌥⌘T 表格，⌥⌘B 行间公式。",
          "Markdown input has Typora's keys: ⌘1–6 headings, ⌘0 body text, ⌘B/I/U, ⌘K link, ⌘\\ clear format, ⌥⌘Q quote, ⌥⌘O / ⌥⌘U lists, ⌥⌘C code block, ⌥⌘T table, ⌥⌘B display formula.",
        ],
      },
      {
        id: "md-source",
        title: ["Markdown：源码与文件", "Markdown: source and files"],
        summary: [
          "源码面板把“输出”切换到 Markdown，实时查看生成的 Markdown；打开 .md 文件后，没有改动的部分按原样写回；也可以导出或复制为 Markdown。",
          "Switch the source panel's Output to Markdown to see the Markdown live; a .md file you open is written back as it was where you did not change it; export or copy as Markdown.",
        ],
      },
    ],
  },
  {
    title: ["页面与文件", "Pages and files"],
    tutorials: [
      {
        id: "page-setup",
        title: ["页面、字体与分页", "Page, fonts and page breaks"],
        summary: [
          "“页面设置”选择纸张、页边距和行距；中文、西文字体和字号分别设置；\\newpage 另起一页。",
          "Page setup picks the paper, margins and line spacing; Chinese and Latin fonts and the size are set apart; \\newpage starts a new page.",
        ],
      },
      {
        id: "paste-import",
        title: ["粘贴 LaTeX 与 Markdown", "Pasting LaTeX and Markdown"],
        summary: [
          "粘贴 LaTeX 源码或 AI 回答里的 Markdown，会直接变成标题、公式、列表和表格；文档库的“导入”可打开 .tex、.md 文件。",
          "Paste LaTeX source or Markdown from an AI answer and it becomes headings, formulas, lists and tables; Import in the library opens .tex and .md files.",
        ],
      },
      {
        id: "export",
        title: ["源码与导出", "Source and export"],
        summary: [
          "底部“源码”实时显示文档的 LaTeX；主菜单可导出 PDF、LaTeX（.tex）、Markdown，或复制为 LaTeX / Markdown。",
          "Source at the bottom shows the document's LaTeX live; the main menu exports PDF, LaTeX (.tex) and Markdown, or copies LaTeX / Markdown.",
        ],
      },
      {
        id: "library",
        title: ["文档库", "The document library"],
        summary: [
          "所有文档自动保存在左侧文档库，标题取自第一行；可以建文件夹、移动、重命名、导出，并按标题和正文搜索。",
          "Documents save themselves into the library on the left, named after their first line; make folders, move, rename and export them, and search titles and text.",
        ],
      },
    ],
  },
];
