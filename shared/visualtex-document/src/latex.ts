// Document -> LaTeX (DESIGN.md §5). Every node type has one explicit
// mapping; only the packages the document uses are loaded.
import type { Mark, Node, Schema, Slice } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";
import { splitInlineMath } from "./mathText.ts";
import { bibliographyEntries } from "./notes.ts";
import { chosenSettings, documentSettings, ownEquationTag, settingsPreamble } from "./settings.ts";
import { IMAGE_GAP, imageFiles, isImage, isImageLine } from "./images.ts";
import { pageCommandLatex } from "./pageCommands.ts";
import { boardHeight, boardItems, canonicalRow } from "./imageBoards.ts";
import { figureParts, figureWrap, isLongTable, tableFigureParts } from "./figures.ts";
import { columnWidth, tableColumns } from "./tableColumns.ts";
import { sliceWithoutMarkdownSyntax, withoutMarkdownSyntax } from "./markdownLive.ts";
import { formulaForPackages, formulaPackageNeeds } from "./formulaPackages.ts";
import { CJK_TEXT } from "./titleBlock.ts";

export interface LatexExportOptions {
  /** The formula's LaTeX made portable (the app's own conventions -> plain amsmath). */
  formulaLatex?: (latex: string, display: boolean) => string;
  /** Packages a formula needs beyond amsmath/amssymb. */
  formulaPackages?: (latex: string) => string[];
  /** A whole document (preamble + \begin{document}) or only the body. */
  standalone?: boolean;
  /** Where an image's file is, relative to the .tex (default: figures/fig-n.ext in the order of appearance). */
  imagePath?: (key: string) => string | null;
  title?: string;
}

const CJK = CJK_TEXT;

/** Code block languages listings knows (by its own names). */
const LISTINGS_LANGUAGES: Record<string, string> = {
  python: "Python",
  c: "C",
  cpp: "C++",
  csharp: "[Sharp]C",
  java: "Java",
  matlab: "Matlab",
  r: "R",
  fortran: "Fortran",
  haskell: "Haskell",
  scala: "Scala",
  ruby: "Ruby",
  php: "PHP",
  lua: "Lua",
  bash: "bash",
  sql: "SQL",
  xml: "XML",
  html: "HTML",
  latex: "TeX",
};

/** \vtxwrapclear: wrapfig's remaining narrowed lines (its counter, + 1) as vertical space, and the wrapping ended. */
const WRAP_CLEAR = [
  "\\makeatletter",
  "\\newcommand\\vtxwrapclear{\\par\\ifnum\\c@WF@wrappedlines>\\@ne",
  "  \\vspace*{\\dimexpr\\baselineskip*\\numexpr\\c@WF@wrappedlines-1\\relax\\relax}\\WF@finale\\global\\c@WF@wrappedlines\\z@\\fi}",
  "\\makeatother",
].join("\n");

/** A width fraction as LaTeX writes it: 0.45, 0.333, 1. */
function widthOf(width: number): string {
  return String(Math.round(width * 1000) / 1000);
}

/** Text for LaTeX's text mode: the ten special characters escaped. */
export function escapeLatexText(text: string): string {
  return text.replace(/[\\{}#$%&_~^]/g, (ch) => {
    switch (ch) {
      case "\\":
        return "\\textbackslash{}";
      case "~":
        return "\\textasciitilde{}";
      case "^":
        return "\\textasciicircum{}";
      default:
        return "\\" + ch;
    }
  });
}

const MARK_OPEN: Record<string, string> = {
  strong: "\\textbf{",
  em: "\\emph{",
  underline: "\\uline{",
  strike: "\\sout{",
  code: "\\texttt{",
  superscript: "\\textsuperscript{",
  subscript: "\\textsubscript{",
  kbd: "\\fbox{\\footnotesize\\ttfamily ",
};

/** A CSS colour as an xcolor argument: [HTML]{RRGGBB} for #rgb/#rrggbb, {name} otherwise. */
function xcolor(color: string): string {
  const hex = color.trim().replace(/^#/, "");
  if (/^[0-9a-f]{3}$/i.test(hex)) return `[HTML]{${hex.split("").map((c) => c + c).join("").toUpperCase()}}`;
  if (/^[0-9a-f]{6}$/i.test(hex)) return `[HTML]{${hex.toUpperCase()}}`;
  const rgb = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/i.exec(color);
  if (rgb) return `[RGB]{${rgb[1]},${rgb[2]},${rgb[3]}}`;
  return `{${color.replace(/[^a-zA-Z!0-9.]/g, "")}}`;
}

/** What opens a mark's group (its closing is always "}"). */
function markOpen(mark: Mark, packages: Set<string>): string | null {
  const name = mark.type.name;
  if (MARK_OPEN[name]) {
    if (name === "underline" || name === "strike") packages.add("ulem");
    return MARK_OPEN[name];
  }
  if (name === "textColor") {
    packages.add("xcolor");
    return `\\textcolor${xcolor(mark.attrs.color)}{`;
  }
  if (name === "highlight") {
    packages.add("xcolor");
    return `\\colorbox${xcolor(mark.attrs.color)}{`;
  }
  if (name === "fontSize") return mark.attrs.size === "normalsize" ? "{" : `{\\${mark.attrs.size} `;
  if (name === "link") {
    packages.add("hyperref");
    const href = String(mark.attrs.href).replace(/[\\{}%#]/g, (c) => `\\${c}`);
    return `\\href{${href}}{`;
  }
  return null;
}

const PACKAGE_HINTS: Array<[RegExp, string]> = [
  [/\\bm\b/, "bm"],
  [/\\(?:b|c|x)?cancel(?:to)?\b/, "cancel"],
  [/\\mathscr\b/, "mathrsfs"],
  [/\\(?:color|textcolor|colorbox)\b/, "xcolor"],
  [/\\(?:mathbbm)\b/, "bbm"],
  [/\\(?:ce|pu)\b/, "mhchem"],
  [/\\(?:overset|underset)\b/, "amsmath"],
  // tables as the formula editor draws them: array.sty, then arydshln (after
  // it) for dashed rules, booktabs' rules, multirow
  [/\\begin\{(?:array|tabular)\}/, "array"],
  [/\\(?:hdashline|cdashline)\b|\\begin\{array\}\{[^{}]*(?:\{[^{}]*\}[^{}]*)*[:;]/, "arydshln"],
  [/\\(?:toprule|midrule|bottomrule|cmidrule|addlinespace|specialrule)\b/, "booktabs"],
  [/\\multirow\b/, "multirow"],
  // the layout engine sets these as mathtools (loaded first) and extarrows define them
  [/\\(?:x(?:mapsto|hook(?:left|right)arrow|leftrightarrow|L(?:eft|eftright)?arrow|Rightarrow|(?:left|right)harpoon(?:up|down)|rightleftharpoons|leftrightharpoons|long(?:left|right)arrow)|(?:Colon|colon|dbl)(?:eqq?|approx|sim)|[Ee]qq?colon|dblcolon|vcentcolon|math[lrc]lap|clap)\b/, "mathtools"],
  [/\\(?:xlongequal|xLong(?:left|right|leftright)arrow|xlongleftrightarrow)\b/, "extarrows"],
  // VisualTeX: the packages of commands the layout engine sets as they do
  [/\\(?:over|under)bracket\b/, "mathtools"],
  [/\\(?:llparenthesis|rrparenthesis|llbracket|rrbracket|llfloor|rrfloor|llceil|rrceil|varcurlyvee|varcurlywedge|sslash|bigsqcap|Mapsto|mapsfrom)\b/, "stmaryrd"],
  [/\\[uU]p(?:alpha|beta|gamma|delta|epsilon|varepsilon|zeta|eta|theta|vartheta|iota|kappa|lambda|mu|nu|xi|pi|varpi|rho|varrho|sigma|varsigma|tau|upsilon|phi|varphi|chi|psi|omega)\b/, "upgreek"],
  [/\\(?:tensor|indices)\b/, "tensor"],
];

export function defaultFormulaPackages(latex: string): string[] {
  return [
    ...PACKAGE_HINTS.filter(([pattern]) => pattern.test(latex)).map(([, name]) => name),
    // physics, braket, siunitx (\qty is both physics' and siunitx's)
    ...formulaPackageNeeds(latex),
  ];
}

/** A formula's place in the body until the document's packages are known (formulaForPackages). */
const FORMULA_MARK = /\u0000(\d+)\u0000/g;

class LatexWriter {
  readonly packages = new Set<string>(["amsmath", "amssymb"]);
  /** the formulas, marked in the body by their index (FORMULA_MARK) */
  readonly formulas: string[] = [];
  hasCjk = false;
  private listDepth = 0;
  private titled = false;
  /** writing the title block: footnotes are \thanks */
  private thanks = false;
  /** a wrapfigure's text may still be wrapping: a block that is not a body paragraph starts below the image */
  private wrapping = false;

  constructor(
    private options: LatexExportOptions,
    private numbers: Map<string, number>,
    /** the reference list's \bibitem keys by entry id */
    private citeKeys: Map<string, string> = new Map(),
  ) {}

  /** A footnote's text: its $…$ formulas as math, the rest as text. */
  private noteText(text: string): string {
    return splitInlineMath(text)
      .map((part) => {
        if (part.kind === "text") return this.text(part.text);
        const latex = this.formula(part.latex, false);
        return latex ? `\\(${latex}\\)` : "";
      })
      .join("");
  }

  private formula(latex: string, display: boolean): string {
    const portable = (this.options.formulaLatex ?? ((l) => l))(latex, display).trim();
    for (const p of (this.options.formulaPackages ?? defaultFormulaPackages)(portable)) this.packages.add(p);
    if (CJK.test(portable)) this.hasCjk = true;
    if (!portable) return "";
    this.formulas.push(portable);
    return `\u0000${this.formulas.length - 1}\u0000`;
  }

  /** The body with its formulas written for the packages the document loads. */
  withFormulas(body: string): string {
    return body.replace(FORMULA_MARK, (_, index: string) => formulaForPackages(this.formulas[Number(index)], this.packages));
  }

  text(text: string): string {
    if (CJK.test(text)) this.hasCjk = true;
    return escapeLatexText(text);
  }

  /** Inline content with marks closed and opened as runs. */
  inline(node: Node): string {
    let out = "";
    const stack: Mark[] = [];
    const closeTo = (keep: number) => {
      while (stack.length > keep) {
        stack.pop();
        out += "}";
      }
    };
    node.forEach((child) => {
      const marks = child.isText ? child.marks : [];
      let common = 0;
      while (common < stack.length && common < marks.length && stack[common].eq(marks[common])) common++;
      closeTo(common);
      for (let i = common; i < marks.length; i++) {
        const open = markOpen(marks[i], this.packages);
        if (!open) continue;
        out += open;
        stack.push(marks[i]);
      }
      // A tab character (Mathcha's Tab in text) is a quad.
      if (child.isText) out += this.text(child.text ?? "").replace(/\t/g, "\\quad{}");
      else if (child.type.name === "math_inline") {
        const latex = this.formula(child.attrs.latex, false);
        if (latex) out += `\\(${latex}\\)`;
      } else if (child.type.name === "hard_break") out += "\\\\\n";
      else if (isImage(child)) {
        // a wrapped image is its paragraph's wrapfigure (blockBody), not part of the line
        if (!child.attrs.wrap) out += this.image(child);
      }
      // in the title block LaTeX's \thanks (\maketitle's footnote)
      else if (child.type.name === "footnote") out += `\\${this.thanks ? "thanks" : "footnote"}{${this.noteText(child.attrs.text).trim()}}`;
      else if (child.type.name === "html_inline") {
        // a picture on the web that was not fetched: a link to it
        const html = String(child.attrs.html ?? "");
        const src = /^!\[[^\]]*\]\(<?([^)\s>]+)/.exec(html)?.[1] ?? /\bsrc\s*=\s*["']?([^"'\s>]+)/i.exec(html)?.[1];
        if (src && /^https?:/i.test(src)) {
          this.packages.add("hyperref");
          out += `\\url{${src.replace(/[\\{}%#]/g, (c) => `\\${c}`)}}`;
        }
      }
      else if (child.type.name === "cite") {
        const keys = (child.attrs.targets as string[]).map((id) => this.citeKeys.get(id)).filter((key): key is string => !!key);
        out += keys.length ? `\\cite{${keys.join(",")}}` : "[?]";
      } else if (child.type.name === "eqref") {
        // hyperref makes the reference a link to its equation in the PDF.
        if (child.attrs.target) this.packages.add("hyperref");
        out += child.attrs.target ? `\\eqref{${child.attrs.target.startsWith("eq:") ? child.attrs.target : `eq:${child.attrs.target}`}}` : "(??)";
      }
    });
    closeTo(0);
    // A line break may not start a paragraph or end one.
    return out.replace(/^(\\\\\n)+/, "").replace(/(\\\\\n)+$/, "");
  }

  /** \includegraphics of an image at its width (a fraction of \linewidth). */
  image(node: Node, width = `${widthOf(node.attrs.width)}\\linewidth`): string {
    this.packages.add("graphicx");
    const path = (this.options.imagePath ?? (() => null))(node.attrs.src) ?? node.attrs.src;
    return `\\includegraphics[width=${width}]{${path}}`;
  }

  /** An image line: its images bottom-aligned on one line, 0.02\linewidth apart, aligned as the paragraph. */
  private imageLine(node: Node): string {
    const images: string[] = [];
    node.forEach((child) => {
      if (isImage(child)) images.push(this.image(child));
    });
    const line = images.join(`\\hspace{${widthOf(IMAGE_GAP)}\\linewidth}%\n`);
    const indent = node.attrs.indent ? `\\hspace*{${node.attrs.indent * 2}em}` : "";
    if (node.attrs.align === "center") return `{\\centering\n${line}\\par}\n`;
    if (node.attrs.align === "right") return `{\\raggedleft\n${line}\\par}\n`;
    return `\\noindent${indent}%\n${line}\\par\n`;
  }

  /**
   * An image board: a plain row as an image line; otherwise a picture with \unitlength = \linewidth and one \put
   * per image at its bottom-left corner (kernel LaTeX; the same places as the editor's).
   */
  private imageBoard(node: Node): string {
    const row = canonicalRow(node);
    if (row) {
      const line = row.items.map((item) => this.image(item.node)).join(`\\hspace{${widthOf(IMAGE_GAP)}\\linewidth}%\n`);
      if (row.align === "center") return `{\\centering\n${line}\\par}\n`;
      if (row.align === "right") return `{\\raggedleft\n${line}\\par}\n`;
      return `\\noindent%\n${line}\\par\n`;
    }
    const f = (value: number) => (Math.round(value * 1000) / 1000).toFixed(3);
    // the height rounded up: no image above the picture's box
    const height = Math.ceil(boardHeight(node) * 1000 - 1e-6) / 1000;
    const puts = boardItems(node).map((item) => `\\put(${f(item.x)},${f(Math.max(0, height - item.top - item.height))}){${this.image(item.node)}}`);
    return `{\\noindent\\setlength{\\unitlength}{\\linewidth}%\n\\begin{picture}(1,${f(height)})\n${puts.join("\n")}\n\\end{picture}\\par}\n`;
  }

  /** A body paragraph that starts with a wrapped image: wrapfig's environment before it (DESIGN.md §11.1). */
  private wrapFigure(image: Node, caption = ""): string {
    // a wrap still going on ends first (wrapfig would collide), as the editor starts this paragraph below it
    const clear = this.wrapping ? "\\vtxwrapclear\n" : "";
    this.packages.add("graphicx");
    this.packages.add("wrapfig");
    this.wrapping = true;
    return `${clear}\\begin{wrapfigure}{${image.attrs.wrap === "left" ? "l" : "r"}}{${widthOf(image.attrs.width)}\\linewidth}\n\\centering\n${this.image(image, "\\linewidth")}\n${caption}\\end{wrapfigure}\n`;
  }

  block(node: Node): string {
    // wrapfig wraps body paragraphs only: before anything else the rest of the image's height is skipped
    // (\vtxwrapclear, defined in the preamble), as the editor starts such a block below the image
    const clear = this.wrapping && (node.type.name !== "paragraph" || isImageLine(node));
    if (clear) this.wrapping = false;
    const text = this.blockBody(node);
    return clear ? `\\vtxwrapclear\n${text}` : text;
  }

  private blockBody(node: Node): string {
    switch (node.type.name) {
      case "paragraph": {
        if (isImageLine(node)) return this.imageLine(node);
        const first = node.firstChild;
        const wrap = first && isImage(first) && first.attrs.wrap ? this.wrapFigure(first) : "";
        const body = this.inline(node).trim();
        if (!body) return "";
        const indented = node.attrs.indent ? `\\noindent\\hspace*{${node.attrs.indent * 2}em}` + body : body;
        if (node.attrs.align === "center") return `${wrap}\\begin{center}\n${indented}\n\\end{center}\n`;
        if (node.attrs.align === "right") return `${wrap}\\begin{flushright}\n${indented}\n\\end{flushright}\n`;
        return wrap + indented + "\n";
      }
      case "title":
      case "author":
      case "date": {
        // outside the title block (titleBlock): a centred line in the title's or the authors' size
        const text = this.inline(node).trim();
        if (!text) return "";
        return `\\begin{center}\n{\\${node.type.name === "title" ? "LARGE" : "large"} ${text}\\par}\n\\end{center}\n`;
      }
      case "abstract":
        return `\\begin{abstract}\n${this.blocks(node)}\\end{abstract}\n`;
      case "heading": {
        const command = ["section", "subsection", "subsubsection", "paragraph", "subparagraph", "subparagraph"][Math.min(6, Math.max(1, node.attrs.level)) - 1];
        return `\\${command}${node.attrs.unnumbered ? "*" : ""}{${this.inline(node).trim()}}\n`;
      }
      case "blockquote": {
        // an alert: its kind first, bold
        const alert = node.attrs.alert ? `\\textbf{${String(node.attrs.alert).replace(/^./, (c) => c.toUpperCase())}.}\n` : "";
        return `\\begin{quote}\n${alert}${this.blocks(node)}\\end{quote}\n`;
      }
      case "image_board":
        return this.imageBoard(node);
      case "toc":
        return "\\tableofcontents\n";
      case "html_block": {
        // HTML has no LaTeX: kept as a comment
        const html = String(node.attrs.html ?? "").trim();
        return html ? `${html.split("\n").map((line) => `% ${line}`).join("\n")}\n` : "";
      }
      case "bullet_list":
      case "ordered_list":
        return this.list(node);
      case "math_display":
        return this.display(node);
      case "page_command":
        return `${pageCommandLatex(node.attrs.command, node.attrs.value)}\n`;
      case "horizontal_rule":
        return "\\par\\noindent\\rule{\\linewidth}{0.4pt}\\par\n";
      case "code_block": {
        // listings: the code verbatim, in its language when listings knows it.
        this.packages.add("listings");
        const language = LISTINGS_LANGUAGES[node.attrs.language as string];
        const code = node.textContent.replace(/\\end\{lstlisting\}/g, "\\end {lstlisting}");
        return `\\begin{lstlisting}${language ? `[language=${language}]` : ""}\n${code}\n\\end{lstlisting}\n`;
      }
      case "table":
        return this.table(node);
      case "figure":
        return this.figure(node);
      case "table_figure":
        return this.tableFigure(node);
      case "bibliography": {
        // numbered as the editor numbers them: [1], [2], … in the list's order
        const items: string[] = [];
        node.forEach((entry) => {
          const key = entry.attrs.id ? this.citeKeys.get(entry.attrs.id) : undefined;
          if (key) items.push(`\\bibitem{${key}} ${this.inline(entry).trim()}`);
        });
        if (!items.length) return "";
        return `\\begin{thebibliography}{${String(items.length).replace(/\d/g, "9")}}\n${items.join("\n")}\n\\end{thebibliography}\n`;
      }
      default:
        return "";
    }
  }

  blocks(node: Node): string {
    const parts: string[] = [];
    const children: Node[] = [];
    node.forEach((child) => children.push(child));
    for (let i = 0; i < children.length; i++) {
      // the first title block: \title, \author, \date and \maketitle
      const run = this.titled ? 0 : titleRun(children, i);
      if (run) {
        this.titled = true;
        parts.push(this.titleBlock(children.slice(i, i + run)));
        i += run - 1;
        continue;
      }
      const text = this.block(children[i]);
      if (text) parts.push(text);
    }
    return parts.join("\n");
  }

  /**
   * The title block (a title, its authors and its date, in that order) as \maketitle prints it: authors joined by
   * \and (a line break in one is \\, the tabular's next line). No date: \date{} — LaTeX's default would be \today,
   * a date the editor does not show (and one that changes with every run).
   */
  private titleBlock(lines: Node[]): string {
    this.thanks = true;
    const of = (name: string) => lines.filter((line) => line.type.name === name).map((line) => this.inline(line).trim());
    const title = of("title")[0] ?? "";
    const authors = of("author").filter(Boolean).join(" \\and ");
    const date = of("date")[0] ?? "";
    this.thanks = false;
    return `\\title{${title}}\n\\author{${authors}}\n\\date{${date}}\n\\maketitle\n`;
  }

  private list(node: Node): string {
    const ordered = node.type.name === "ordered_list";
    const env = ordered ? "enumerate" : "itemize";
    this.listDepth++;
    let out = `\\begin{${env}}\n`;
    if (ordered && node.attrs.start !== 1 && this.listDepth <= 4) {
      out += `\\setcounter{enum${"i".repeat(this.listDepth)}}{${node.attrs.start - 1}}\n`;
    }
    node.forEach((item) => {
      const body = this.blocks(item).trim();
      // a task: an empty or ticked box as its label (amssymb)
      const box = item.attrs.checked === true ? "[$\\boxtimes$]" : item.attrs.checked === false ? "[$\\square$]" : "";
      out += `\\item${box} ${body}\n`;
    });
    out += `\\end{${env}}\n`;
    this.listDepth--;
    return out;
  }

  private display(node: Node): string {
    const latex = this.formula(node.attrs.latex, true);
    if (!latex) return "";
    // amsmath's own displays (align, gather, multline, …) are written as they are: they are displays themselves
    if (/^\\begin\{(align|flalign|alignat|gather|multline|eqnarray)(\*?)\}[\s\S]*\\end\{\1\2\}$/.test(latex)) return `${latex}\n`;
    // a formula's own \tag replaces the number (an equation, so that \eqref finds its label)
    if (node.attrs.numbered || ownEquationTag(latex)) {
      const label = node.attrs.id ? `\\label{${node.attrs.id.startsWith("eq:") ? node.attrs.id : `eq:${node.attrs.id}`}}` : "";
      return `\\begin{equation}${label}\n${latex}\n\\end{equation}\n`;
    }
    return `\\[\n${latex}\n\\]\n`;
  }

  /** A cell's content for a tabular column ("block": a p{} column that may hold display math). */
  private cell(cell: Node, block: boolean): string {
    const parts: string[] = [];
    cell.forEach((child) => {
      if (child.type.name === "math_display") {
        const latex = this.formula(child.attrs.latex, true);
        if (latex) parts.push(block ? `\\[${latex}\\]` : `\\(\\displaystyle ${latex}\\)`);
      } else {
        const text = this.inline(child).trim();
        if (text) parts.push(text);
      }
    });
    return parts.join(block ? "\\par " : " ");
  }

  /**
   * A table: a ruled grid in a center (a table on its own), or with its caption (figures.ts) in a [H] table float, or
   * a longtable when it cannot be on one page; booktabs' rules for a table that has them.
   */
  private table(table: Node, float: { caption: string; label: string; rules: string; long: boolean } | null = null): string {
    const map = TableMap.get(table);
    const booktabs = float?.rules === "booktabs";
    const bar = booktabs ? "" : "|";
    // The columns as the page view sets them (tableColumns.ts): l c r, or p{} — a column given a width, or one whose
    // cells hold display math or several paragraphs (its share of \linewidth) — centred or right-aligned as its cells
    const columns = tableColumns(table);
    const aligned = (align: string | null) => (align === "center" || align === "right" ? align : null);
    const letter = (align: string | null) => ({ center: "c", right: "r" } as Record<string, string>)[align ?? ""] ?? "l";
    const columnSpec = (col: number, align: string | null) => {
      const width = columnWidth(columns[col], map.width);
      if (!width) return letter(align);
      const p = "pt" in width ? `p{${width.pt}pt}` : `p{${width.share}\\linewidth}`;
      if (align !== "center" && align !== "right") return p;
      this.packages.add("array");
      return `>{${align === "center" ? "\\centering" : "\\raggedleft"}\\arraybackslash}${p}`;
    };
    const isP = columns.map((column) => !!columnWidth(column, map.width));
    // Mathcha's tables are ruled grids.
    const spec = bar + columns.map((column, col) => columnSpec(col, column.align)).join(bar) + bar;
    const lines: string[] = [];
    for (let row = 0; row < map.height; row++) {
      const cells: string[] = [];
      for (let col = 0; col < map.width; ) {
        const offset = map.map[row * map.width + col];
        const cell = table.nodeAt(offset)!;
        const rect = map.findCell(offset);
        const colspan = rect.right - rect.left;
        let content = rect.top === row ? this.cell(cell, isP[col]) : "";
        // a line break in a cell: the next line of the cell (in an l c r column a \makecell of its lines)
        if (content.includes("\\\\\n")) {
          if (isP[col]) content = content.replace(/\\\\\n/g, "\\newline ");
          else {
            this.packages.add("makecell");
            content = `\\makecell[${letter(cell.attrs.align)}]{${content.replace(/\\\\\n/g, "\\\\ ")}}`;
          }
        }
        const rowspan = rect.bottom - rect.top;
        if (rowspan > 1 && rect.top === row) {
          this.packages.add("multirow");
          content = `\\multirow{${rowspan}}{*}{${content}}`;
        }
        if (colspan > 1) {
          content = `\\multicolumn{${colspan}}{${col === 0 ? bar : ""}${letter(cell.attrs.align)}${bar}}{${content}}`;
        } else if (rect.top === row && content && aligned(cell.attrs.align) !== aligned(columns[col].align)) {
          // a cell aligned otherwise than its column: its own \multicolumn, as the page view shows it
          content = `\\multicolumn{1}{${col === 0 ? bar : ""}${columnSpec(col, cell.attrs.align)}${bar}}{${content}}`;
        }
        cells.push(content);
        col += colspan;
      }
      lines.push(cells.join(" & ") + " \\\\");
    }
    // A rule under every row, except across a cell that goes on into the next row (\cline around it).
    const ruleAfter = (row: number) => {
      if (row === map.height - 1) return "\\hline";
      const ranges: Array<[number, number]> = [];
      for (let col = 0; col < map.width; col++) {
        const continues = map.map[row * map.width + col] === map.map[(row + 1) * map.width + col];
        if (continues) continue;
        const last = ranges[ranges.length - 1];
        if (last && last[1] === col) last[1] = col + 1;
        else ranges.push([col, col + 1]);
      }
      if (ranges.length === 1 && ranges[0][0] === 0 && ranges[0][1] === map.width) return "\\hline";
      return ranges.map(([a, b]) => `\\cline{${a + 1}-${b}}`).join(" ");
    };
    let body = booktabs ? "\\toprule\n" : "\\hline\n";
    lines.forEach((line, row) => {
      body += booktabs ? `${line}${row === map.height - 1 ? "\n\\bottomrule" : row === 0 ? "\n\\midrule" : ""}\n` : `${line} ${ruleAfter(row)}\n`;
    });
    if (booktabs) this.packages.add("booktabs");
    if (!float) return `\\begin{center}\n\\begin{tabular}{${spec}}\n${body}\\end{tabular}\n\\end{center}\n`;
    // its caption above it, as LaTeX puts a table's (the caption package's skip under it)
    this.packages.add("caption");
    if (float.long) {
      this.packages.add("longtable");
      return `\\begin{longtable}{${spec}}\n\\caption{${float.caption}}${float.label}\\\\\n${body}\\end{longtable}\n`;
    }
    this.packages.add("float");
    return `\\begin{table}[H]\n\\centering\n\\caption{${float.caption}}${float.label}\n\\begin{tabular}{${spec}}\n${body}\\end{tabular}\n\\end{table}\n`;
  }

  /** The document being written (a long table's height is measured on its page). */
  document: Node | null = null;

  /**
   * A figure (figures.ts): its image line or board and its caption in a [H] figure (LaTeX keeps it where it is, as the
   * editor shows it), numbered and labelled; a wrapped one as its wrapfigure with the caption under the image. An
   * empty caption still numbers ("Figure 1": the caption package drops the colon).
   */
  private figure(node: Node): string {
    const { body, caption } = figureParts(node);
    if (!body) return "";
    const text = `\\caption{${this.captionText(caption)}}${node.attrs.id ? `\\label{fig:${node.attrs.id}}` : ""}\n`;
    this.packages.add("caption");
    if (figureWrap(node)) {
      let image: Node | null = null;
      body.forEach((child) => {
        if (isImage(child)) image = child;
      });
      return image ? this.wrapFigure(image, text) : "";
    }
    this.packages.add("float");
    let inner: string;
    const row = body.type.name === "image_board" ? canonicalRow(body) : null;
    if (body.type.name === "image_board" && !row) inner = this.imageBoard(body);
    else {
      const images: string[] = [];
      if (row) for (const item of row.items) images.push(this.image(item.node));
      else body.forEach((child) => {
        if (isImage(child)) images.push(this.image(child));
      });
      const align = row ? row.align : ((body.attrs.align as string | null) ?? "left");
      const indent = !row && body.attrs.indent ? `\\hspace*{${body.attrs.indent * 2}em}` : "";
      const switch_ = align === "center" ? "\\centering" : align === "right" ? "\\raggedleft" : "\\raggedright";
      inner = `${switch_}\n${indent}${images.join(`\\hspace{${widthOf(IMAGE_GAP)}\\linewidth}%\n`)}\n`;
    }
    return `\\begin{figure}[H]\n${inner}${text}\\end{figure}\n`;
  }

  /** A caption's text: its inline content, a footnote in it protected (\caption's argument moves). */
  private captionText(caption: Node): string {
    return this.inline(caption).trim().replace(/\\footnote\{/g, "\\protect\\footnote{");
  }

  /** A table with its caption (figures.ts): a table float, or a longtable when it is longer than a page. */
  private tableFigure(node: Node): string {
    const { caption, table } = tableFigureParts(node);
    if (!table) return "";
    const long = this.document ? isLongTable(this.document, table) : false;
    return this.table(table, { caption: this.captionText(caption), label: node.attrs.id ? `\\label{tab:${node.attrs.id}}` : "", rules: node.attrs.rules, long });
  }

  numberOf(id: string): number | undefined {
    return this.numbers.get(id);
  }
}

/** How many blocks from `i` make the title block: [title] author* [date], in that order, with a title or an author
 *  that has text (0: none starts here). */
function titleRun(children: Node[], i: number): number {
  let k = i;
  if (children[k]?.type.name === "title") k++;
  while (children[k]?.type.name === "author") k++;
  if (children[k]?.type.name === "date") k++;
  const lines = children.slice(i, k);
  const filled = (line: Node) => !!line.textContent.trim() || line.content.content.some((child) => !child.isText);
  return lines.some((line) => (line.type.name === "title" || line.type.name === "author") && filled(line)) ? k - i : 0;
}

/** The document as LaTeX. */
export function documentToLatex(doc: Node, options: LatexExportOptions = {}): string {
  doc = withoutMarkdownSyntax(doc);
  const numbers = new Map<string, number>();
  const citeKeys = new Map(bibliographyEntries(doc).map((entry) => [entry.id, entry.key]));
  // the images' files: figures/fig-1.png, … in the order they appear
  const paths = imageFiles(doc, "figures").paths;
  const writer = new LatexWriter({ imagePath: (key) => paths.get(key) ?? null, ...options }, numbers, citeKeys);
  writer.document = doc;
  const body = writer.withFormulas(writer.blocks(doc));
  if (options.standalone === false) return body;
  const packages = [...writer.packages].sort().map((name) =>
    name === "ulem" ? "\\usepackage[normalem]{ulem}" : `\\usepackage{${name}}`,
  );
  // the rest of a wrapped image's lines skipped before a block wrapfig cannot wrap (a heading, a list, an image line)
  if (writer.packages.has("wrapfig")) packages.push(WRAP_CLEAR);
  // the page, body size, spacing and fonts the editor shows (settings.ts)
  const page = settingsPreamble(documentSettings(doc), chosenSettings(doc), writer.hasCjk);
  const classOptions = page.classOptions.length ? `[${page.classOptions.join(",")}]` : "";
  const documentClass = writer.hasCjk ? `\\documentclass${classOptions}{ctexart}` : `\\documentclass${classOptions}{article}`;
  const program = page.xelatex ? "% !TEX program = xelatex\n" : "";
  const title = options.title ? `\\title{${escapeLatexText(options.title)}}\n\\date{}\n` : "";
  const maketitle = options.title ? "\\maketitle\n\n" : "";
  return `${program}${documentClass}\n${[...packages, ...page.lines].join("\n")}\n${title}\n\\begin{document}\n${maketitle}${body}\n\\end{document}\n`;
}

/** A copied selection as LaTeX (Copy Latex): its blocks, or the inline content as one paragraph's text. */
export function sliceToLatex(slice: Slice, schema: Schema, options: LatexExportOptions = {}): string {
  slice = sliceWithoutMarkdownSyntax(slice, schema);
  const blocks: Node[] = [];
  let inline: Node[] = [];
  const flush = () => {
    if (inline.length) blocks.push(schema.nodes.paragraph.create(null, inline));
    inline = [];
  };
  slice.content.forEach((node) => {
    if (node.isInline) inline.push(node);
    else {
      flush();
      blocks.push(node);
    }
  });
  flush();
  const doc = schema.nodes.doc.create(null, blocks);
  return documentToLatex(doc, { ...options, standalone: false }).trim();
}
