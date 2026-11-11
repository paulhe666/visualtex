// Commands typed in the text (DESIGN.md §6.1, §7.3). In LaTeX input "\" starts a command right in the line; in
// Markdown input "/" does (at the start of a line or after a space), and the menu offers Markdown's blocks and
// marks, written as Markdown writes them, found by name or by the Markdown itself ("/#" is a heading). The characters stay ordinary text while a menu under them offers the
// commands that match. Enter, Tab, Space or 1-9 run the highlighted (numbered) one, a closing brace completes
// \textbf{…} or \begin{itemize}, \( and \[ open formulas, Escape leaves the text as it is. The command is read from the
// text, not from the key, so a "\" or "/" typed through an input method starts one just the same.
import type { MarkType, Node } from "prosemirror-model";
import { requestImage } from "./images.ts";
import { Plugin, PluginKey, TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { setBlockType, toggleMark, wrapIn } from "prosemirror-commands";
import { sinkListItem, splitListItem, wrapInList } from "prosemirror-schema-list";
import { insertDivider, insertPageCommand, insertTable, setAlign } from "./blocks.ts";
import { PAGE_NUMBER_STYLES, PAGE_STYLES } from "./pageCommands.ts";
import { insertDisplayLatex, insertInlineFormula, insertInlineLatex } from "./commands.ts";
import { insertCodeBlock } from "./codeBlocks.ts";
import { requestLink } from "./links.ts";
import { openFormulaId } from "./formulas.ts";
import { MARKDOWN_HIGHLIGHT } from "./markdown.ts";
import type { InputMode, MarkdownRuleSwitches } from "./markdownInput.ts";
import { FONT_SIZES } from "./schema.ts";
import { emojiOf, searchEmoji } from "./emoji.ts";
import { insertAbstract, nextAuthor } from "./titleBlock.ts";
import { addReferenceEntry, bibliographyEntries, insertCitation, insertFootnote, requestCitation } from "./notes.ts";
import { inSourceCode } from "./sourceMarkdown.ts";

/** The command being typed. */
export interface CommandToken {
  /** the position of its "\" (or "/") */
  from: number;
  trigger: "\\" | "/" | "、" | "／" | ":";
  /** LaTeX commands or Markdown's (the input mode when the command was started) */
  mode: InputMode;
  /** what follows the trigger: a name, "name{argument" or "begin{environment" */
  query: string;
  /** the highlighted menu entry */
  index: number;
  /** the highlight was moved (↑/↓): Space takes it even before a name is typed */
  navigated?: boolean;
  /** Escape: the text stays as typed and the menu is gone */
  dismissed: boolean;
}

/** What an entry's picture in the menu shows (the app draws it). */
export type CommandPreview =
  | { kind: "math"; latex: string }
  | { kind: "display"; latex: string; numbered: boolean }
  | { kind: "text"; text: string; style?: "bold" | "italic" | "underline" | "mono" | "strike"; scale?: number }
  | { kind: "section"; level: 1 | 2 | 3; unnumbered?: boolean }
  | { kind: "title" }
  | { kind: "list"; ordered: boolean }
  | { kind: "align"; align: "left" | "center" | "right" }
  | { kind: "block"; block: "table" | "rule" | "quote" | "break" | "code" | "image" }
  | { kind: "ref"; number: number | null };

/** How an entry was chosen: from the menu, by Space after its full name, or by the closing brace. */
export interface CommandRun {
  via: "menu" | "space" | "complete";
  argument?: string;
}

export interface CommandEntry {
  id: string;
  /** shown in the menu, as LaTeX (or Markdown) writes it */
  label: string;
  /** the names that call it (\name, /name) */
  names: string[];
  /** \begin{environment} calls it too */
  environment?: string;
  /** \name{text}: the text is its argument */
  argument?: boolean;
  /** shown when only "\" is typed */
  featured?: boolean;
  description: { en: string; zh: string };
  preview: CommandPreview;
  run(view: EditorView, how: CommandRun): boolean;
}

/** A formula command of the app's formula editor (\alpha, \frac, ...): it goes into a formula. */
export interface MathCommandSuggestion {
  id: string;
  /** "\frac" */
  command: string;
  /** "\frac{\placeholder{}}{\placeholder{}}" */
  template: string;
  previewLatex: string;
  labelZh: string;
  labelEn: string;
}

/** The app's formula commands: searched like the formula editor's own suggestions, inserted by its editor. */
export interface MathCommandProvider {
  search(query: string, limit: number): MathCommandSuggestion[];
  /** Put the command into the formula that was just opened at the caret. */
  insertIntoOpenFormula(view: EditorView, command: MathCommandSuggestion): void;
}

type CommandMeta = { end: true } | { dismiss: true } | { index: number; keyboard?: boolean };

export const commandMenuKey = new PluginKey<CommandToken | null>("visualtex-document-commands");

export function commandToken(state: EditorState): CommandToken | null {
  return commandMenuKey.getState(state) ?? null;
}

const NAME = /^[A-Za-z]*\*?$/;
const ENVIRONMENT = /^begin\{([A-Za-z]*\*?)$/;
const ARGUMENT = /^([A-Za-z]+\*?)\{([^{}\\]*)$/;
/** characters LaTeX escapes with a backslash: "\$" is a dollar sign, not a formula */
const ESCAPED = "$%&#_{}";

/** a "/" command's name in Markdown input: Latin letters, Chinese (the words of the menu's descriptions), or the
 *  Markdown it writes ("#", "##", "-", ">", "```", "---", "|", "**") */
const SLASH_NAME = /^(?:[A-Za-z]*|[\u3400-\u9fff\uf900-\ufaff]+|[#*\-+>`|$=~[\]^<!\\]{1,6})$/;

/** an emoji's short code being typed after ":" (Typora): lower-case letters, digits, "_", "+", "-" */
const EMOJI_NAME = /^[a-z0-9_+-]*$/;

function validQuery(query: string, trigger: CommandToken["trigger"], mode: InputMode): boolean {
  if (trigger === ":") return mode === "markdown" && EMOJI_NAME.test(query);
  if (trigger !== "\\") return mode === "markdown" && SLASH_NAME.test(query);
  return NAME.test(query) || ENVIRONMENT.test(query) || ARGUMENT.test(query);
}

// ---------------------------------------------------------------- commands --

/** The caret starts a line of its own: text before it on the line is split off. */
function ownLine(view: EditorView) {
  const { $from } = view.state.selection;
  if ($from.parent.isTextblock && $from.parentOffset > 0) view.dispatch(view.state.tr.split($from.pos).scrollIntoView());
}

/** The title is complete: the caret goes on in body text on the next line (with the text that followed, if any). */
function bodyTextAfter(view: EditorView) {
  const { $from } = view.state.selection;
  if (!$from.parent.isTextblock) return;
  view.dispatch(view.state.tr.split($from.pos, 1, [{ type: view.state.schema.nodes.paragraph }]).scrollIntoView());
}

/** \section and its kin; `unnumbered`: the starred form. Typed on a section of the same level without an argument,
 *  the other form switches that section's number off (\section*) or on again (\section) in place. */
function section(level: 1 | 2 | 3, unnumbered = false): CommandEntry["run"] {
  return (view, how) => {
    const here = view.state.selection.$from;
    if (how.argument === undefined && here.parent.type.name === "heading" && here.parent.attrs.level === level && here.parent.attrs.unnumbered !== unnumbered) {
      view.dispatch(view.state.tr.setNodeAttribute(here.before(), "unnumbered", unnumbered).scrollIntoView());
      return true;
    }
    ownLine(view);
    const { $from } = view.state.selection;
    const block = $from.parent;
    if (!setBlockType(view.state.schema.nodes.heading, { level, unnumbered, align: block.attrs.align ?? null, indent: 0 })(view.state, view.dispatch)) {
      return false;
    }
    if (how.argument !== undefined) {
      if (how.argument) view.dispatch(view.state.tr.insertText(how.argument).scrollIntoView());
      bodyTextAfter(view);
    }
    return true;
  };
}

/** \title{…}: the line is the document's title (centred, unnumbered); after the argument, body text. */
const documentTitle: CommandEntry["run"] = (view, how) => {
  ownLine(view);
  if (!setBlockType(view.state.schema.nodes.title)(view.state, view.dispatch)) return false;
  if (how.argument !== undefined) {
    if (how.argument) view.dispatch(view.state.tr.insertText(how.argument).scrollIntoView());
    bodyTextAfter(view);
  }
  return true;
};

/** \author{…} / \date{…}: the line is an author / the date of the title block; after the argument, body text. */
function titleBlockLine(kind: "author" | "date"): CommandEntry["run"] {
  return (view, how) => {
    ownLine(view);
    if (!setBlockType(view.state.schema.nodes[kind])(view.state, view.dispatch)) return false;
    if (how.argument !== undefined) {
      if (how.argument) view.dispatch(view.state.tr.insertText(how.argument).scrollIntoView());
      bodyTextAfter(view);
    }
    return true;
  };
}

/** Markdown's "{-}": a section's number off or on again; on body text, an unnumbered section. */
const unnumberedSection: CommandEntry["run"] = (view, how) => {
  const { $from } = view.state.selection;
  if ($from.parent.type.name === "heading") {
    view.dispatch(view.state.tr.setNodeAttribute($from.before(), "unnumbered", !$from.parent.attrs.unnumbered).scrollIntoView());
    return true;
  }
  return section(1, true)(view, how);
};

function textStyle(mark: (view: EditorView) => MarkType, attrs: Record<string, unknown> | null = null): CommandEntry["run"] {
  return (view, how) => {
    const type = mark(view);
    if (how.argument === undefined) return toggleMark(type, attrs)(view.state, view.dispatch);
    if (!how.argument) return true;
    const { from } = view.state.selection;
    const tr = view.state.tr.insertText(how.argument, from);
    tr.addMark(from, from + how.argument.length, type.create(attrs));
    // The text after the argument is plain again.
    tr.setStoredMarks(tr.doc.resolve(from + how.argument.length).marks().filter((m) => m.type !== type));
    view.dispatch(tr.scrollIntoView());
    return true;
  };
}

function fontSize(size: string): CommandEntry["run"] {
  return (view) => {
    const type = view.state.schema.marks.fontSize;
    const tr = view.state.tr;
    if (size === "normalsize") tr.removeStoredMark(type);
    else tr.addStoredMark(type.create({ size }));
    view.dispatch(tr);
    return true;
  };
}

function listDepth(state: EditorState): { depth: number; node: Node } | null {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const name = $from.node(d).type.name;
    if (name === "bullet_list" || name === "ordered_list") return { depth: d, node: $from.node(d) };
  }
  return null;
}

/** \begin{itemize} / \begin{enumerate}: a list from this line; inside a list, a list nested in it (as in LaTeX). */
function list(kind: "bullet_list" | "ordered_list"): CommandEntry["run"] {
  return (view) => {
    const { schema } = view.state;
    const item = schema.nodes.list_item;
    if (!listDepth(view.state)) {
      ownLine(view);
      return wrapInList(schema.nodes[kind])(view.state, view.dispatch);
    }
    const { $from } = view.state.selection;
    if ($from.parent.content.size > 0) splitListItem(item)(view.state, view.dispatch);
    if (!sinkListItem(item)(view.state, view.dispatch)) return true;
    const nested = listDepth(view.state);
    if (nested && nested.node.type.name !== kind) {
      const pos = view.state.selection.$from.before(nested.depth);
      view.dispatch(view.state.tr.setNodeMarkup(pos, schema.nodes[kind], { ...nested.node.attrs }));
    }
    return true;
  };
}

/** \item: a new item of the list (the text after the caret goes with it); outside a list, a bulleted list. */
const listItem: CommandEntry["run"] = (view, how) => {
  if (!listDepth(view.state)) return list("bullet_list")(view, how);
  const { $from } = view.state.selection;
  if ($from.parentOffset === 0) return true;
  return splitListItem(view.state.schema.nodes.list_item)(view.state, view.dispatch);
};

const quote: CommandEntry["run"] = (view) => {
  ownLine(view);
  return wrapIn(view.state.schema.nodes.blockquote)(view.state, view.dispatch);
};

const lineBreak: CommandEntry["run"] = (view) => {
  view.dispatch(view.state.tr.replaceSelectionWith(view.state.schema.nodes.hard_break.create()).scrollIntoView());
  return true;
};

/** \footnote{text}: a footnote with that text; chosen from the menu, the footnote's editor opens. */
const footnote: CommandEntry["run"] = (view, how) => insertFootnote(how.argument?.trim() ?? "")(view.state, view.dispatch, view);

/** \cite{key}: a citation of the entries with those keys (or numbers); else the citation picker. */
const cite: CommandEntry["run"] = (view, how) => {
  const keys = (how.argument ?? "").split(",").map((key) => key.trim()).filter(Boolean);
  if (keys.length) {
    const entries = bibliographyEntries(view.state.doc);
    const ids = keys
      .map((key) => entries.find((entry) => entry.key === key || String(entry.number) === key)?.id)
      .filter((id): id is string => !!id);
    if (ids.length === keys.length) return insertCitation(ids)(view.state, view.dispatch, view);
  }
  return requestCitation(view.state, view.dispatch, view);
};

/** \pagenumbering{roman} and its kin: the typed argument when it is one of the command's values, else its default. */
const pageCommandEntry =
  (command: string, values: readonly string[], fallback: string): CommandEntry["run"] =>
  (view, how) => {
    const typed = how.argument?.trim() ?? "";
    const value = values.includes(typed) ? typed : fallback;
    return insertPageCommand(command, value)(view.state, view.dispatch, view);
  };

const asRun =
  (command: Command): CommandEntry["run"] =>
  (view) =>
    command(view.state, view.dispatch, view);

// The caret in the first row's left cell, as an aligned row typed in formula mode: an & moves what is before it
// into the right-aligned cell (the kernel's alignment point).
const ALIGNED = "\\begin{aligned}&\\placeholder{}\\end{aligned}";
const GATHERED = "\\begin{gathered}\\placeholder{}\\end{gathered}";

const t = (en: string, zh: string) => ({ en, zh });

/** The document's commands, in the order the menu offers them for a bare "\". */
export const DOCUMENT_COMMANDS: CommandEntry[] = [
  { id: "section", label: "\\section{…}", names: ["section"], argument: true, featured: true, description: t("Section", "一级标题"), preview: { kind: "section", level: 1 }, run: section(1) },
  { id: "subsection", label: "\\subsection{…}", names: ["subsection"], argument: true, featured: true, description: t("Subsection", "二级标题"), preview: { kind: "section", level: 2 }, run: section(2) },
  { id: "subsubsection", label: "\\subsubsection{…}", names: ["subsubsection"], argument: true, description: t("Subsubsection", "三级标题"), preview: { kind: "section", level: 3 }, run: section(3) },
  // the starred forms after the three levels: \sec + ↓ / 3 still pick \subsection / \subsubsection
  { id: "section*", label: "\\section*{…}", names: ["section*"], argument: true, description: t("Section, unnumbered", "一级标题（无编号）"), preview: { kind: "section", level: 1, unnumbered: true }, run: section(1, true) },
  { id: "subsection*", label: "\\subsection*{…}", names: ["subsection*"], argument: true, description: t("Subsection, unnumbered", "二级标题（无编号）"), preview: { kind: "section", level: 2, unnumbered: true }, run: section(2, true) },
  { id: "subsubsection*", label: "\\subsubsection*{…}", names: ["subsubsection*"], argument: true, description: t("Subsubsection, unnumbered", "三级标题（无编号）"), preview: { kind: "section", level: 3, unnumbered: true }, run: section(3, true) },
  { id: "title", label: "\\title{…}", names: ["title", "maketitle"], argument: true, featured: true, description: t("Title", "文档标题"), preview: { kind: "title" }, run: documentTitle },
  { id: "author", label: "\\author{…}", names: ["author", "zuozhe", "zz"], argument: true, description: t("Author", "作者"), preview: { kind: "text", text: "Author" }, run: titleBlockLine("author") },
  { id: "and", label: "\\and", names: ["and"], description: t("Next author", "下一位作者"), preview: { kind: "text", text: "A  B" }, run: asRun(nextAuthor) },
  { id: "date", label: "\\date{…}", names: ["date", "riqi", "rq"], argument: true, description: t("Date", "日期"), preview: { kind: "text", text: "2026" }, run: titleBlockLine("date") },
  { id: "abstract", label: "\\begin{abstract}", names: ["abstract", "zhaiyao", "zy"], environment: "abstract", description: t("Abstract", "摘要"), preview: { kind: "block", block: "quote" }, run: asRun(insertAbstract) },
  { id: "itemize", label: "\\begin{itemize}", names: ["itemize"], environment: "itemize", featured: true, description: t("Bulleted list", "无序列表"), preview: { kind: "list", ordered: false }, run: list("bullet_list") },
  { id: "enumerate", label: "\\begin{enumerate}", names: ["enumerate"], environment: "enumerate", featured: true, description: t("Numbered list", "有序列表"), preview: { kind: "list", ordered: true }, run: list("ordered_list") },
  { id: "item", label: "\\item", names: ["item"], description: t("List item", "列表项"), preview: { kind: "list", ordered: false }, run: listItem },
  { id: "inline-math", label: "\\( … \\)", names: [], environment: "math", featured: true, description: t("Inline formula", "行内公式"), preview: { kind: "math", latex: "x^2" }, run: asRun(insertInlineFormula) },
  { id: "display-math", label: "\\[ … \\]", names: ["displaymath"], environment: "displaymath", featured: true, description: t("Display formula", "行间公式"), preview: { kind: "display", latex: "E=mc^2", numbered: false }, run: asRun(insertDisplayLatex("")) },
  { id: "equation", label: "\\begin{equation}", names: ["equation"], environment: "equation", featured: true, description: t("Numbered equation", "编号公式"), preview: { kind: "display", latex: "E=mc^2", numbered: true }, run: asRun(insertDisplayLatex("", true)) },
  { id: "equation*", label: "\\begin{equation*}", names: ["equation*"], environment: "equation*", description: t("Unnumbered equation", "无编号公式"), preview: { kind: "display", latex: "E=mc^2", numbered: false }, run: asRun(insertDisplayLatex("")) },
  { id: "align", label: "\\begin{align}", names: ["align"], environment: "align", featured: true, description: t("Aligned lines (numbered)", "多行对齐（编号）"), preview: { kind: "display", latex: "\\begin{aligned}a&=b\\\\&=c\\end{aligned}", numbered: true }, run: asRun(insertDisplayLatex(ALIGNED, true)) },
  { id: "align*", label: "\\begin{align*}", names: ["align*"], environment: "align*", description: t("Aligned lines", "多行对齐"), preview: { kind: "display", latex: "\\begin{aligned}a&=b\\\\&=c\\end{aligned}", numbered: false }, run: asRun(insertDisplayLatex(ALIGNED)) },
  { id: "gather", label: "\\begin{gather}", names: ["gather"], environment: "gather", description: t("Centred lines (numbered)", "多行居中（编号）"), preview: { kind: "display", latex: "\\begin{gathered}a=b\\\\c=d\\end{gathered}", numbered: true }, run: asRun(insertDisplayLatex(GATHERED, true)) },
  { id: "gather*", label: "\\begin{gather*}", names: ["gather*"], environment: "gather*", description: t("Centred lines", "多行居中"), preview: { kind: "display", latex: "\\begin{gathered}a=b\\\\c=d\\end{gathered}", numbered: false }, run: asRun(insertDisplayLatex(GATHERED)) },
  { id: "textbf", label: "\\textbf{…}", names: ["textbf"], argument: true, featured: true, description: t("Bold", "粗体"), preview: { kind: "text", text: "Bold", style: "bold" }, run: textStyle((v) => v.state.schema.marks.strong) },
  { id: "textit", label: "\\textit{…}", names: ["textit"], argument: true, featured: true, description: t("Italic", "斜体"), preview: { kind: "text", text: "Italic", style: "italic" }, run: textStyle((v) => v.state.schema.marks.em) },
  { id: "emph", label: "\\emph{…}", names: ["emph"], argument: true, description: t("Emphasis", "强调"), preview: { kind: "text", text: "Emph", style: "italic" }, run: textStyle((v) => v.state.schema.marks.em) },
  { id: "underline", label: "\\underline{…}", names: ["underline", "uline"], argument: true, description: t("Underline", "下划线"), preview: { kind: "text", text: "Under", style: "underline" }, run: textStyle((v) => v.state.schema.marks.underline) },
  { id: "texttt", label: "\\texttt{…}", names: ["texttt"], argument: true, description: t("Typewriter", "等宽字体"), preview: { kind: "text", text: "mono", style: "mono" }, run: textStyle((v) => v.state.schema.marks.code) },
  { id: "sout", label: "\\sout{…}", names: ["sout"], argument: true, description: t("Strikethrough", "删除线"), preview: { kind: "text", text: "Strike", style: "strike" }, run: textStyle((v) => v.state.schema.marks.strike) },
  { id: "lstlisting", label: "\\begin{lstlisting}", names: ["lstlisting", "code", "minted"], environment: "lstlisting", featured: true, description: t("Code block (syntax highlighting)", "代码块（语法高亮）"), preview: { kind: "block", block: "code" }, run: asRun(insertCodeBlock("")) },
  { id: "verbatim", label: "\\begin{verbatim}", names: ["verbatim"], environment: "verbatim", description: t("Code block, plain text", "代码块（纯文本）"), preview: { kind: "block", block: "code" }, run: asRun(insertCodeBlock("")) },
  { id: "tabular", label: "\\begin{tabular}", names: ["tabular", "table"], environment: "tabular", featured: true, description: t("Table (3 × 3)", "表格（3 × 3）"), preview: { kind: "block", block: "table" }, run: asRun(insertTable(3, 3)) },
  { id: "center", label: "\\begin{center}", names: ["center", "centering"], environment: "center", description: t("Centred lines", "居中"), preview: { kind: "align", align: "center" }, run: asRun(setAlign("center")) },
  { id: "flushleft", label: "\\begin{flushleft}", names: ["flushleft", "raggedright"], environment: "flushleft", description: t("Left-aligned lines", "左对齐"), preview: { kind: "align", align: "left" }, run: asRun(setAlign("left")) },
  { id: "flushright", label: "\\begin{flushright}", names: ["flushright", "raggedleft"], environment: "flushright", description: t("Right-aligned lines", "右对齐"), preview: { kind: "align", align: "right" }, run: asRun(setAlign("right")) },
  { id: "quote", label: "\\begin{quote}", names: ["quote", "quotation"], environment: "quote", description: t("Quotation", "引用"), preview: { kind: "block", block: "quote" }, run: quote },
  { id: "href", label: "\\href{…}{…}", names: ["href", "url", "link"], description: t("Link (Ctrl+K)", "链接（Ctrl+K）"), preview: { kind: "text", text: "link", style: "underline" }, run: asRun(requestLink) },
  { id: "includegraphics", label: "\\includegraphics{…}", names: ["includegraphics", "image", "figure", "picture", "graphics", "tupian", "tp", "charutupian", "crtp"], featured: true, description: t("Image", "图片"), preview: { kind: "block", block: "image" }, run: (view) => requestImage(view) },
  { id: "newpage", label: "\\newpage", names: ["newpage", "pagebreak", "fenye", "fy", "huanye"], featured: true, description: t("Page break", "分页"), preview: { kind: "block", block: "rule" }, run: asRun(insertPageCommand("newpage")) },
  { id: "clearpage", label: "\\clearpage", names: ["clearpage", "fenye"], description: t("Page break (after floats)", "分页（浮动体之后）"), preview: { kind: "block", block: "rule" }, run: asRun(insertPageCommand("clearpage")) },
  { id: "pagenumbering", label: "\\pagenumbering{…}", names: ["pagenumbering", "yema", "ym", "yemageshi"], argument: true, description: t("Page numbers: arabic, roman, Roman, alph, Alph, gobble", "页码格式：arabic、roman、Roman、alph、Alph、gobble"), preview: { kind: "text", text: "i ii iii" }, run: pageCommandEntry("pagenumbering", PAGE_NUMBER_STYLES, "arabic") },
  { id: "pagestyle", label: "\\pagestyle{…}", names: ["pagestyle", "yemianfengge"], argument: true, description: t("Page style: plain, empty, headings, myheadings", "页面样式：plain、empty、headings、myheadings"), preview: { kind: "text", text: "plain" }, run: pageCommandEntry("pagestyle", PAGE_STYLES, "plain") },
  { id: "thispagestyle", label: "\\thispagestyle{…}", names: ["thispagestyle"], argument: true, description: t("This page's style: plain, empty, headings, myheadings", "本页样式：plain、empty、headings、myheadings"), preview: { kind: "text", text: "empty" }, run: pageCommandEntry("thispagestyle", PAGE_STYLES, "empty") },
  { id: "hrule", label: "\\hrule", names: ["hrule", "hrulefill"], featured: true, description: t("Horizontal line", "水平线"), preview: { kind: "block", block: "rule" }, run: asRun(insertDivider) },
  { id: "footnote", label: "\\footnote{…}", names: ["footnote", "jiaozhu", "jz"], argument: true, description: t("Footnote (Ctrl+Alt+F)", "脚注（Ctrl+Alt+F）"), preview: { kind: "text", text: "a¹" }, run: footnote },
  { id: "cite", label: "\\cite{…}", names: ["cite", "citep", "citet", "yinyong", "wenxian", "ckwx"], argument: true, description: t("Citation", "引用文献"), preview: { kind: "text", text: "[1]" }, run: cite },
  { id: "thebibliography", label: "\\begin{thebibliography}", names: ["thebibliography", "bibliography", "references", "cankaowenxian", "ckwx"], environment: "thebibliography", description: t("References", "参考文献"), preview: { kind: "text", text: "[1] …" }, run: asRun(addReferenceEntry) },
  { id: "newline", label: "\\newline", names: ["newline", "linebreak"], description: t("Line break", "换行"), preview: { kind: "block", block: "break" }, run: lineBreak },
  ...FONT_SIZES.map(
    ({ name, ratio }): CommandEntry => ({
      id: `size-${name}`,
      label: `\\${name}`,
      names: [name],
      description: t(`Text size ${name}`, `字号 ${name}`),
      preview: { kind: "text", text: "Aa", scale: ratio },
      run: fontSize(name),
    }),
  ),
];

/** Markdown input's "/" menu: Markdown's blocks and marks, as Markdown writes them (found by English names and by
 * the Chinese words, in characters or pinyin). */
export const MARKDOWN_COMMANDS: CommandEntry[] = [
  { id: "md-title", label: '<h1 align="center">', names: ["title", "biaoti", "bt", "wdbt"], description: t("Title", "文档标题"), preview: { kind: "title" }, run: documentTitle },
  { id: "md-author", label: '<p align="center" class="author">', names: ["author", "zuozhe", "zz"], description: t("Author", "作者"), preview: { kind: "text", text: "Author" }, run: titleBlockLine("author") },
  { id: "md-date", label: '<p align="center" class="date">', names: ["date", "riqi", "rq"], description: t("Date", "日期"), preview: { kind: "text", text: "2026" }, run: titleBlockLine("date") },
  { id: "md-abstract", label: '<div class="abstract">', names: ["abstract", "zhaiyao", "zy"], description: t("Abstract", "摘要"), preview: { kind: "block", block: "quote" }, run: asRun(insertAbstract) },
  { id: "md-h1", label: "# …", names: ["h1", "heading", "biaoti", "bt", "yjbt"], featured: true, description: t("Heading 1", "一级标题"), preview: { kind: "section", level: 1 }, run: section(1) },
  { id: "md-h2", label: "## …", names: ["h2", "heading", "subtitle", "biaoti", "bt", "ejbt"], featured: true, description: t("Heading 2", "二级标题"), preview: { kind: "section", level: 2 }, run: section(2) },
  { id: "md-h3", label: "### …", names: ["h3", "heading", "biaoti", "bt", "sjbt"], featured: true, description: t("Heading 3", "三级标题"), preview: { kind: "section", level: 3 }, run: section(3) },
  { id: "md-unnumbered", label: "# … {-}", names: ["unnumbered", "nonumber", "heading", "wubianhao", "wbh"], description: t("Heading, unnumbered", "标题（无编号）"), preview: { kind: "section", level: 1, unnumbered: true }, run: unnumberedSection },
  { id: "md-bullet", label: "- …", names: ["ul", "bullet", "list", "liebiao", "wuxu", "lb", "wxlb"], featured: true, description: t("Bulleted list", "无序列表"), preview: { kind: "list", ordered: false }, run: list("bullet_list") },
  { id: "md-ordered", label: "1. …", names: ["ol", "numbered", "ordered", "list", "liebiao", "youxu", "lb", "yxlb"], featured: true, description: t("Numbered list", "有序列表"), preview: { kind: "list", ordered: true }, run: list("ordered_list") },
  { id: "md-math-block", label: "$$ … $$", names: ["math", "formula", "equation", "display", "gongshi", "gs", "hjgs", "sxgs"], featured: true, description: t("Display formula", "行间公式"), preview: { kind: "display", latex: "E=mc^2", numbered: false }, run: asRun(insertDisplayLatex("")) },
  { id: "md-equation", label: "$$ … \\tag{n} $$", names: ["equation", "numbered", "tag", "math", "formula", "gongshi", "bianhao", "gs", "bh", "bhgs"], featured: true, description: t("Numbered formula", "编号公式"), preview: { kind: "display", latex: "E=mc^2", numbered: true }, run: asRun(insertDisplayLatex("", true)) },
  { id: "md-inline-math", label: "$ … $", names: ["inline", "math", "formula", "hangnei", "gongshi", "gs", "hngs"], featured: true, description: t("Inline formula", "行内公式"), preview: { kind: "math", latex: "x^2" }, run: asRun(insertInlineFormula) },
  { id: "md-code", label: "``` …", names: ["code", "codeblock", "fence", "daima", "dm", "dmk"], featured: true, description: t("Code block (syntax highlighting)", "代码块（语法高亮）"), preview: { kind: "block", block: "code" }, run: asRun(insertCodeBlock("")) },
  { id: "md-table", label: "| … | … |", names: ["table", "biaoge", "bg"], featured: true, description: t("Table (3 × 3)", "表格（3 × 3）"), preview: { kind: "block", block: "table" }, run: asRun(insertTable(3, 3)) },
  { id: "md-quote", label: "> …", names: ["quote", "blockquote", "yinyong", "yy"], featured: true, description: t("Quotation", "引用"), preview: { kind: "block", block: "quote" }, run: quote },
  { id: "md-footnote", label: "[^1]", names: ["footnote", "note", "jiaozhu", "zhushi", "jz"], description: t("Footnote (Ctrl+Alt+F)", "脚注（Ctrl+Alt+F）"), preview: { kind: "text", text: "a¹" }, run: footnote },
  { id: "md-cite", label: "[1]", names: ["cite", "citation", "reference", "wenxian", "cankao", "ckwx", "yinwen"], description: t("Citation", "引用文献"), preview: { kind: "text", text: "[1]" }, run: cite },
  { id: "md-references", label: "[1] …", names: ["references", "bibliography", "cankaowenxian", "wenxian", "ckwx"], description: t("Reference entry", "参考文献条目"), preview: { kind: "text", text: "[1] …" }, run: asRun(addReferenceEntry) },
  { id: "md-image", label: "![…](…)", names: ["image", "img", "picture", "figure", "tupian", "tp", "charutupian", "crtp"], featured: true, description: t("Image", "图片"), preview: { kind: "block", block: "image" }, run: (view) => requestImage(view) },
  { id: "md-rule", label: "---", names: ["hr", "divider", "rule", "line", "fengexian", "fgx"], featured: true, description: t("Horizontal line", "分割线"), preview: { kind: "block", block: "rule" }, run: asRun(insertDivider) },
  { id: "md-bold", label: "**…**", names: ["bold", "strong", "cuti", "jiacu", "ct", "jc"], description: t("Bold", "粗体"), preview: { kind: "text", text: "Bold", style: "bold" }, run: textStyle((v) => v.state.schema.marks.strong) },
  { id: "md-italic", label: "*…*", names: ["italic", "em", "emphasis", "xieti", "xt"], description: t("Italic", "斜体"), preview: { kind: "text", text: "Italic", style: "italic" }, run: textStyle((v) => v.state.schema.marks.em) },
  { id: "md-strike", label: "~~…~~", names: ["strike", "strikethrough", "del", "shanchuxian", "scx"], description: t("Strikethrough", "删除线"), preview: { kind: "text", text: "Strike", style: "strike" }, run: textStyle((v) => v.state.schema.marks.strike) },
  { id: "md-highlight", label: "==…==", names: ["highlight", "mark", "gaoliang", "gl"], description: t("Highlight", "高亮"), preview: { kind: "text", text: "Mark" }, run: textStyle((v) => v.state.schema.marks.highlight, { color: MARKDOWN_HIGHLIGHT }) },
  { id: "md-inline-code", label: "`…`", names: ["code", "inline", "mono", "daima", "dm", "hndm"], description: t("Inline code", "行内代码"), preview: { kind: "text", text: "code", style: "mono" }, run: textStyle((v) => v.state.schema.marks.code) },
  { id: "md-link", label: "[…](…)", names: ["link", "url", "href", "lianjie", "lj"], description: t("Link (Ctrl+K)", "链接（Ctrl+K）"), preview: { kind: "text", text: "link", style: "underline" }, run: asRun(requestLink) },
  { id: "md-underline", label: "<u>…</u>", names: ["underline", "xiahuaxian", "xhx"], description: t("Underline", "下划线"), preview: { kind: "text", text: "Under", style: "underline" }, run: textStyle((v) => v.state.schema.marks.underline) },
  { id: "md-break", label: "\\", names: ["br", "break", "newline", "huanhang", "hh"], description: t("Line break (Shift+Enter)", "换行（Shift+Enter）"), preview: { kind: "block", block: "break" }, run: lineBreak },
];

/** \eqref: one entry per numbered display formula (in Markdown input, written as the reference reads: "(n)"). */
function equationReferences(doc: Node, markdown = false): CommandEntry[] {
  const entries: CommandEntry[] = [];
  let n = 0;
  doc.descendants((node) => {
    if (node.type.name === "math_display" && node.attrs.numbered && node.attrs.id) {
      const number = ++n;
      const target = node.attrs.id as string;
      entries.push({
        id: `eqref-${target}`,
        label: markdown ? `(${number})` : `\\eqref{${number}}`,
        names: markdown ? ["ref", "eqref", "yinyong"] : ["eqref", "ref"],
        description: t(`Reference to equation (${number})`, `引用公式 (${number})`),
        preview: { kind: "math", latex: node.attrs.latex || "\\square" },
        run: (view) => {
          const tr = view.state.tr.replaceSelectionWith(view.state.schema.nodes.eqref.create({ target }), false);
          view.dispatch(tr.scrollIntoView());
          return true;
        },
      });
    }
    return true;
  });
  return entries;
}

function mathEntry(math: MathCommandProvider, command: MathCommandSuggestion): CommandEntry {
  return {
    id: `math-${command.id}`,
    label: command.command,
    names: [command.command.replace(/^\\/, "")],
    description: { en: command.labelEn, zh: command.labelZh },
    preview: { kind: "math", latex: command.previewLatex },
    run: (view) => {
      // A symbol goes into the text as a formula of its own and the text goes on; a command with slots to fill
      // opens a formula with it, the caret in its first slot.
      if (!command.template.includes("\\placeholder")) return insertInlineLatex(command.template)(view.state, view.dispatch);
      if (!insertInlineFormula(view.state, view.dispatch, view)) return false;
      math.insertIntoOpenFormula(view, command);
      return true;
    },
  };
}

const MATH_LIMIT = 12;

/** What an entry writes, without its placeholders: "# …" -> "#", "**…**" -> "**", "| … | … |" -> "|". */
function markdownSyntax(label: string): string {
  return label.split("…")[0].trim();
}

/** Markdown's "/" entries for `query`: typed Markdown finds what writes it ("#" a heading, "##" the next level);
 *  otherwise the names it starts first, then the ones whose name or description has it. */
function matchMarkdown(query: string, doc: Node): CommandEntry[] {
  if (!query) return MARKDOWN_COMMANDS.filter((e) => e.featured);
  if (!/^[A-Za-z㐀-鿿豈-﫿]/.test(query)) {
    const exact = MARKDOWN_COMMANDS.filter((e) => markdownSyntax(e.label) === query);
    const longer = MARKDOWN_COMMANDS.filter((e) => !exact.includes(e) && markdownSyntax(e.label).startsWith(query));
    return [...exact, ...longer];
  }
  const all = [...MARKDOWN_COMMANDS, ...equationReferences(doc, true)];
  const lower = query.toLowerCase();
  const prefix = all.filter((e) => e.names.some((n) => n.startsWith(lower)));
  const described = all.filter(
    (e) =>
      !prefix.includes(e) &&
      (e.names.some((n) => n.includes(lower)) || e.description.zh.includes(query) || e.description.en.toLowerCase().includes(lower)),
  );
  return [...prefix, ...described];
}

/** The menu's entries for a command being typed, best first. */
/** The emoji whose short codes start with (or hold) what follows ":" (Typora's emoji list). */
function emojiEntries(query: string): CommandEntry[] {
  return searchEmoji(query).map(({ name, emoji }) => ({
    id: `emoji-${name}`,
    label: `:${name}:`,
    names: [name],
    description: { en: "", zh: "" },
    preview: { kind: "text", text: emoji, scale: 1.3 },
    run(view) {
      view.dispatch(view.state.tr.insertText(emoji).scrollIntoView());
      return true;
    },
  }));
}

export function matchCommands(token: CommandToken, doc: Node, math: MathCommandProvider | null): CommandEntry[] {
  const query = token.query;
  if (token.trigger === ":") return emojiEntries(query);
  if (token.mode === "markdown") return matchMarkdown(query, doc);
  const environment = ENVIRONMENT.exec(query);
  if (environment) {
    const name = environment[1];
    const found = DOCUMENT_COMMANDS.filter((e) => e.environment?.startsWith(name));
    return [...found.filter((e) => e.environment === name), ...found.filter((e) => e.environment !== name)];
  }
  const argument = ARGUMENT.exec(query);
  if (argument) return DOCUMENT_COMMANDS.filter((e) => e.argument && e.names.includes(argument[1]));
  if (!NAME.test(query)) return [];
  const refs = equationReferences(doc);
  if (!query) return DOCUMENT_COMMANDS.filter((e) => e.featured);
  const lower = query.toLowerCase();
  const all = [...DOCUMENT_COMMANDS, ...refs];
  const exact = all.filter((e) => e.names.includes(query));
  // \sec also finds \subsection and \subsubsection.
  const prefix = all.filter((e) => !exact.includes(e) && e.names.some((n) => n.toLowerCase().replace(/^(sub)+/, "").startsWith(lower) || n.toLowerCase().startsWith(lower)));
  // "\beg" offers the environments as \begin{…}.
  const begin = lower.length >= 2 && "begin".startsWith(lower) ? DOCUMENT_COMMANDS.filter((e) => e.environment && e.label.startsWith("\\begin")) : [];
  const contains = all.filter(
    (e) => !exact.includes(e) && !prefix.includes(e) && !begin.includes(e) && e.names.some((n) => n.toLowerCase().includes(lower)),
  );
  const own = new Set([...exact, ...prefix, ...begin].flatMap((e) => e.names));
  const maths =
    token.trigger === "\\" && math
      ? math
          .search(query, MATH_LIMIT)
          .filter((c) => !own.has(c.command.replace(/^\\/, "")))
          .map((c) => mathEntry(math, c))
      : [];
  // An exact formula command comes before document commands that merely start with the query.
  const exactMath = maths.filter((e) => e.names.includes(query));
  const otherMath = maths.filter((e) => !exactMath.includes(e));
  return [...exact, ...exactMath, ...prefix, ...begin, ...otherMath, ...contains];
}

/** The entries the menu shows now (none when no command is being typed or the menu was dismissed). */
export function commandMenuEntries(state: EditorState): CommandEntry[] {
  const token = commandToken(state);
  if (!token || token.dismissed) return [];
  const plugin = commandMenuKey.get(state) as (Plugin<CommandToken | null> & { spec: { math?: MathCommandProvider | null; only?: ((entry: CommandEntry) => boolean) | null } }) | undefined;
  const entries = matchCommands(token, state.doc, plugin?.spec.math ?? null);
  const only = plugin?.spec.only;
  return only ? entries.filter(only) : entries;
}

// ------------------------------------------------------------- the plugin --

/** Run an entry in place of the typed command. */
export function runCommandEntry(view: EditorView, entry: CommandEntry, how: CommandRun = { via: "menu" }): boolean {
  const token = commandToken(view.state);
  if (token) {
    const argument = how.argument ?? ARGUMENT.exec(token.query)?.[2];
    const end = view.state.selection.head;
    view.dispatch(view.state.tr.delete(token.from, end).setMeta(commandMenuKey, { end: true } satisfies CommandMeta));
    how = { ...how, argument };
  }
  const done = entry.run(view, how);
  view.focus();
  return done;
}

/** Escape: the typed command stays text and the menu goes. */
export function dismissCommandMenu(view: EditorView) {
  if (commandToken(view.state)) view.dispatch(view.state.tr.setMeta(commandMenuKey, { dismiss: true } satisfies CommandMeta));
}

/** The highlighted entry; `keyboard`: moved with ↑/↓ (then Space takes it), not just under the pointer. */
export function highlightCommandEntry(view: EditorView, index: number, keyboard = false) {
  if (commandToken(view.state)) view.dispatch(view.state.tr.setMeta(commandMenuKey, { index, keyboard } satisfies CommandMeta));
}

/** The menu button: a "\" at the caret as if typed (in Markdown input a "/", after a space when it follows text). */
export function startCommand(view: EditorView) {
  view.focus();
  const { $from, empty } = view.state.selection;
  if (!empty || !$from.parent.isTextblock || openFormulaId(view.state)) return;
  const plugin = commandMenuKey.get(view.state) as (Plugin<CommandToken | null> & { spec: { inputMode?: () => InputMode } }) | undefined;
  if (plugin?.spec.inputMode?.() !== "markdown") {
    view.dispatch(view.state.tr.insertText("\\").scrollIntoView());
    return;
  }
  const before = $from.parent.textBetween(Math.max(0, $from.parentOffset - 1), $from.parentOffset, undefined, "￼");
  view.dispatch(view.state.tr.insertText(before && !/\s/.test(before) ? " /" : "/").scrollIntoView());
}

/** Whether `tr` itself put a character at `pos` (typed it, not undo or a paste). */
function insertedBy(tr: Transaction, pos: number): boolean {
  let found = false;
  tr.mapping.maps.forEach((map, i) => {
    map.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      if (found || newEnd <= newStart) return;
      const rest = tr.mapping.slice(i + 1);
      if (rest.map(newStart, -1) <= pos && rest.map(newEnd, 1) >= pos + 1) found = true;
    });
  });
  return found;
}

/** Chinese text and punctuation: a "/" right after them starts a Markdown command too (Chinese has no spaces). */
const CJK = /[\u3000-\u303f\u3400-\u9fff\uf900-\ufaff\uff00-\uffef]/;

/** A "\" (or "/" at a word start) this transaction typed right before the caret starts a command. */
function detect(tr: Transaction, state: EditorState, mode: InputMode, chinese: boolean, emoji = true): CommandToken | null {
  if (!tr.docChanged || tr.getMeta("history$") || tr.getMeta("uiEvent") === "paste" || tr.getMeta("uiEvent") === "drop") return null;
  const sel = state.selection;
  if (!(sel instanceof TextSelection) || !sel.empty) return null;
  const $head = sel.$head;
  const parent = $head.parent;
  if (!parent.isTextblock || parent.type.spec.code || $head.parentOffset < 1 || openFormulaId(state)) return null;
  const at = (offset: number) => (offset < 0 ? "" : parent.textBetween(offset, offset + 1, undefined, "￼"));
  const trigger = at($head.parentOffset - 1);
  // In Markdown input a Chinese input method's "／", and "、" (its "\" or "/") at the start of a line, are "/" too.
  // In LaTeX input only "\" starts a command.
  const slash = mode === "markdown" && (trigger === "/" || (chinese && (trigger === "／" || (trigger === "、" && $head.parentOffset === 1))));
  // ":" in Markdown input starts an emoji's short code (Typora), at a word's start (not in the source editor: Obsidian)
  const colon = mode === "markdown" && emoji && trigger === ":";
  if (trigger !== "\\" && !slash && !colon) return null;
  if (!insertedBy(tr, $head.pos - 1)) return null;
  // Code is literal text.
  if (state.doc.resolve($head.pos - 1).marks().some((m) => m.type.name === "code") || inSourceCode($head)) return null;
  const before = at($head.parentOffset - 2);
  // In Markdown input only "/" starts a command ("\" escapes the character after it).
  if (mode === "markdown" && trigger === "\\") return null;
  if (slash && before && !/\s/.test(before) && !(mode === "markdown" && CJK.test(before))) return null;
  if (colon && before && !/[\s(\[{"'“‘]/.test(before) && !CJK.test(before)) return null;
  if (trigger === "\\" && before === "\\") return null;
  return { from: $head.pos - 1, trigger: trigger as CommandToken["trigger"], mode, query: "", index: 0, dismissed: false };
}

/** The command after `tr`: it follows the typing; it ends when the caret leaves it or the text stops being one. */
function follow(token: CommandToken, tr: Transaction, state: EditorState): CommandToken | null {
  const mapped = tr.mapping.mapResult(token.from, 1);
  if (mapped.deleted) return null;
  const from = mapped.pos;
  const sel = state.selection;
  if (!(sel instanceof TextSelection) || !sel.empty || sel.head <= from) return null;
  const $from = state.doc.resolve(from);
  if (!$from.sameParent(sel.$head)) return null;
  const text = $from.parent.textBetween($from.parentOffset, sel.$head.parentOffset, undefined, "￼");
  if (text[0] !== token.trigger) return null;
  // The caret went back into the name: no longer typing it.
  const next = sel.$head.parent.textBetween(sel.$head.parentOffset, Math.min(sel.$head.parentOffset + 1, sel.$head.parent.content.size), undefined, "￼");
  if (/^[A-Za-z*{}]$/.test(next)) return null;
  const query = text.slice(1);
  if (!validQuery(query, token.trigger, token.mode)) return null;
  return query === token.query ? { ...token, from, query } : { ...token, from, query, index: 0, navigated: false };
}

export function commandMenuPlugin(
  math: MathCommandProvider | null = null,
  inputMode: () => InputMode = () => "latex",
  rules: () => MarkdownRuleSwitches = () => ({ blocks: true, marks: true, chinese: true }),
  /** the entries offered (a footnote's editor: formulas only); default all */
  only: ((entry: CommandEntry) => boolean) | null = null,
): Plugin<CommandToken | null> {
  return new Plugin<CommandToken | null>({
    key: commandMenuKey,
    math,
    inputMode,
    only,
    state: {
      init: () => null,
      apply(tr, value, _old, state) {
        const meta = tr.getMeta(commandMenuKey) as CommandMeta | undefined;
        if (meta && "end" in meta) return null;
        let next = value ? follow(value, tr, state) : null;
        if (!next) next = detect(tr, state, inputMode(), rules().chinese, rules().source !== true);
        if (next && meta && "dismiss" in meta) next = { ...next, dismissed: true };
        if (next && meta && "index" in meta) next = { ...next, index: meta.index, navigated: next.navigated || !!meta.keyboard };
        return next;
      },
    },
    props: {
      handleKeyDown(view, event) {
        const token = commandToken(view.state);
        if (!token || token.dismissed || event.isComposing || event.keyCode === 229) return false;
        if (event.ctrlKey || event.metaKey || event.altKey) return false;
        if (event.key === "Escape") {
          dismissCommandMenu(view);
          return true;
        }
        const entries = commandMenuEntries(view.state);
        if (!entries.length) return false;
        const index = Math.min(token.index, entries.length - 1);
        if (event.key === "ArrowDown" || event.key === "ArrowUp") {
          const step = event.key === "ArrowDown" ? 1 : -1;
          highlightCommandEntry(view, (index + step + entries.length) % entries.length, true);
          return true;
        }
        if ((event.key === "Enter" || event.key === "Tab") && !event.shiftKey) {
          runCommandEntry(view, entries[index], { via: "menu" });
          return true;
        }
        // 1-9 pick from the numbered page, as in the formula editor's suggestions (not in an argument: \textbf{12}; not
        // in an emoji's name: :100:).
        if (/^[1-9]$/.test(event.key) && token.trigger !== ":" && (ENVIRONMENT.test(token.query) || !ARGUMENT.test(token.query))) {
          const pick = index - (index % 9) + Number(event.key) - 1;
          if (pick < entries.length) {
            runCommandEntry(view, entries[pick], { via: "menu" });
            return true;
          }
        }
        return false;
      },
      handleTextInput(view, from, to, text) {
        const token = commandToken(view.state);
        if (!token || from !== to || from !== view.state.selection.head || text.length !== 1) return false;
        const query = token.query;
        const done = (entry: CommandEntry | undefined, how: CommandRun) => (entry && (!only || only(entry)) ? runCommandEntry(view, entry, how) : false);
        // After Escape the typed text is just text.
        if (token.dismissed) return false;
        // an emoji's short code: the closing ":" makes the emoji (":smile:" -> 😄); a space ends it as text
        if (token.trigger === ":") {
          const emoji = text === ":" ? emojiOf(query) : undefined;
          if (!emoji) return false;
          view.dispatch(view.state.tr.insertText(emoji, token.from, from).setMeta(commandMenuKey, { end: true } satisfies CommandMeta).scrollIntoView());
          return true;
        }
        if (token.trigger === "\\" && query === "") {
          if (text === "(") return done(DOCUMENT_COMMANDS.find((e) => e.id === "inline-math"), { via: "complete" });
          if (text === "[") return done(DOCUMENT_COMMANDS.find((e) => e.id === "display-math"), { via: "complete" });
          if (text === "\\") return done(DOCUMENT_COMMANDS.find((e) => e.id === "newline"), { via: "complete" });
          if (ESCAPED.includes(text)) {
            // \$ \% \& \# \_ \{ \}: the character itself.
            view.dispatch(view.state.tr.insertText(text, token.from, from).setMeta(commandMenuKey, { end: true } satisfies CommandMeta));
            return true;
          }
        }
        if (text === " ") {
          // Space picks the highlighted entry, as in the formula editor's suggestions — once a name is typed or the
          // highlight was moved with ↑/↓ (a lone "\" or "/" and Space stay text: "km / h") and not inside an
          // argument (\textbf{two words}).
          // In Markdown input a "/" alone at the start of a line takes Space too (a line does not begin with "/ ").
          const lineStart = token.mode === "markdown" && view.state.doc.resolve(token.from).parentOffset === 0;
          if ((!query && !token.navigated && !lineStart) || (ARGUMENT.test(query) && !ENVIRONMENT.test(query))) return false;
          const entries = commandMenuEntries(view.state);
          if (!entries.length) return false;
          return done(entries[Math.min(token.index, entries.length - 1)], { via: "space" });
        }
        if (text === "}") {
          const environment = ENVIRONMENT.exec(query);
          if (environment) return done(DOCUMENT_COMMANDS.find((e) => e.environment === environment[1]), { via: "complete" });
          const argument = ARGUMENT.exec(query);
          if (argument) {
            const entry = DOCUMENT_COMMANDS.find((e) => e.argument && e.names.includes(argument[1]));
            return done(entry, { via: "complete", argument: argument[2] });
          }
        }
        return false;
      },
    },
  });
}
