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
          "选中文字后用工具栏或快捷键加粗、倾斜、着色；标题、列表、对齐都在工具栏上。",
          "Select text and use the toolbar or keys for bold, italic and colour; headings, lists and alignment are in the toolbar too.",
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
          "正文就是 Markdown 源码，和 Obsidian 一样：光标不在时按格式显示，光标移到标记上时显示源码；$$ 行内公式、$$$$ 行间公式；/ 打开命令菜单。",
          "The text is its Markdown source, as in Obsidian: it shows formatted where the caret is not and shows its marks where the caret is; $$ inline and $$$$ display formulas; / opens the menu.",
        ],
      },
      {
        id: "md-blocks",
        title: ["Markdown：标题、列表与引用", "Markdown: headings, lists and quotes"],
        summary: [
          "行首输入 -、1.、> 加空格，得到无序列表、有序列表和引用；行首 # 加空格就是标题，## 是二级标题，删掉 # 变回正文（# 标题 {-} 无编号）；- [ ] 是待办事项，⌘L 切换待办；--- 回车是分隔线。",
          "At a line start, -, 1. and > with a space make bullet and numbered lists and quotes; # and a space at a line start make a heading, ## a level-2 one, and deleting the # makes it text again (# Title {-}: unnumbered); - [ ] is a to-do item, ⌘L toggles it; --- and ↩ a rule.",
        ],
      },
      {
        id: "md-inline",
        title: ["Markdown：粗体、斜体与链接", "Markdown: bold, italic and links"],
        summary: [
          "**粗体**、*斜体*、~~删除线~~、==高亮==、`代码`、[文字](网址)：标记可以先成对输入再在中间打字，也可以随时修改（** 改成 * 就是斜体）；光标碰到时显示标记；选中文字再输入 * 或 ` 会包住它；\\* 是字面字符。",
          "**bold**, *italic*, ~~strike~~, ==highlight==, `code`, [text](url): type the marks in pairs and the text between, or edit them any time (** to * makes italic); the caret at a span shows its marks; * or ` typed over selected text goes around it; \\* is the character itself.",
        ],
      },
      {
        id: "md-math",
        title: ["Markdown：公式", "Markdown: formulas"],
        summary: [
          "$$ 开始行内公式，$$$$ 开始行间公式；单个 $ 就是美元符号（$5 保持为文字）；斜杠菜单里有带编号的公式和公式引用。",
          "$$ starts an inline formula and $$$$ a display formula; a single $ is a dollar sign ($5 stays text); the / menu has numbered formulas and references.",
        ],
      },
      {
        id: "md-tables-code",
        title: ["Markdown：表格与代码", "Markdown: tables and code"],
        summary: [
          "| a | b | 回车，再输入 | --- | --- | 回车就是表格，之后每输入一行 | … | 回车加一行；```语言 回车开始代码块，最后一行输入 ``` 回车结束；⌘T 直接插入表格。",
          "| a | b |, ↩, then | --- | --- | and ↩ make a table, and each | … | line with ↩ after it adds a row; ```language and ↩ starts code, ``` and ↩ on its last line ends it; ⌘T inserts a table.",
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
        title: ["Markdown：快捷键", "Markdown: keys"],
        summary: [
          "⌘1–6 标题，⌘0 正文；⌘B/I/U 在单词里加上或去掉标记，在粗体末尾再按 ⌘B 跳出；⌘K 链接，⌘\\ 清除格式，⌘L 待办；⌥⌘Q 引用，⌥⌘O / ⌥⌘U 列表，⌥⌘C 代码块，⌥⌘T 表格，⌥⌘B 行间公式。",
          "⌘1–6 headings, ⌘0 body text; ⌘B/I/U add or remove the marks around a word, and ⌘B at a bold span's end steps out of it; ⌘K link, ⌘\\ clear format, ⌘L to-do; ⌥⌘Q quote, ⌥⌘O / ⌥⌘U lists, ⌥⌘C code block, ⌥⌘T table, ⌥⌘B display formula.",
        ],
      },
      {
        id: "md-source",
        title: ["Markdown：源码与文件", "Markdown: source and files"],
        summary: [
          "源码面板实时显示生成的 Markdown；打开 .md 文件后，没有改动的部分按原样写回；也可以导出或复制为 Markdown。",
          "The source panel shows the Markdown live; a .md file you open is written back as it was where you did not change it; export or copy as Markdown.",
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
          "“页面设置”选择纸张、页边距和行距；中文、西文字体和字号分别设置。",
          "Page setup picks the paper, margins and line spacing; Chinese and Latin fonts and the size are set apart.",
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
