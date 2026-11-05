// Document -> LaTeX (DESIGN.md §5). Every node type has one explicit
// mapping; only the packages the document uses are loaded.
import type { Mark, Node, Schema, Slice } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";
import { splitInlineMath } from "./mathText.ts";
import { bibliographyEntries } from "./notes.ts";
import { chosenSettings, documentSettings, settingsPreamble } from "./settings.ts";
import { IMAGE_GAP, imageFiles, isImage, isImageLine } from "./images.ts";
import { pageCommandLatex } from "./pageCommands.ts";
import { boardHeight, boardItems, canonicalRow } from "./imageBoards.ts";

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

const CJK = /[⺀-鿿豈-﫿＀-￯　-〿]/;

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
  [/\\(?:si|SI|num|qty|unit)\b/, "siunitx"],
  [/\\(?:ce|pu)\b/, "mhchem"],
  [/\\(?:overset|underset)\b/, "amsmath"],
];

export function defaultFormulaPackages(latex: string): string[] {
  return PACKAGE_HINTS.filter(([pattern]) => pattern.test(latex)).map(([, name]) => name);
}

class LatexWriter {
  readonly packages = new Set<string>(["amsmath", "amssymb"]);
  hasCjk = false;
  private listDepth = 0;
  private titled = false;
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
    return portable;
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
      else if (child.type.name === "footnote") out += `\\footnote{${this.noteText(child.attrs.text).trim()}}`;
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
  private wrapFigure(image: Node): string {
    // a wrap still going on ends first (wrapfig would collide), as the editor starts this paragraph below it
    const clear = this.wrapping ? "\\vtxwrapclear\n" : "";
    this.packages.add("graphicx");
    this.packages.add("wrapfig");
    this.wrapping = true;
    return `${clear}\\begin{wrapfigure}{${image.attrs.wrap === "left" ? "l" : "r"}}{${widthOf(image.attrs.width)}\\linewidth}\n\\centering\n${this.image(image, "\\linewidth")}\n\\end{wrapfigure}\n`;
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
      case "title": {
        // the first is the document's title; any other a centred large line
        const text = this.inline(node).trim();
        if (!text) return "";
        if (this.titled) return `\\begin{center}\n{\\LARGE ${text}\\par}\n\\end{center}\n`;
        this.titled = true;
        return `\\title{${text}}\n\\author{}\n\\date{}\n\\maketitle\n`;
      }
      case "heading": {
        const command = ["section", "subsection", "subsubsection", "paragraph", "subparagraph", "subparagraph"][Math.min(6, Math.max(1, node.attrs.level)) - 1];
        return `\\${command}{${this.inline(node).trim()}}\n`;
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
    node.forEach((child) => {
      const text = this.block(child);
      if (text) parts.push(text);
    });
    return parts.join("\n");
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
    if (node.attrs.numbered) {
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

  private table(table: Node): string {
    const map = TableMap.get(table);
    // A column holds blocks when a cell starting in it has display math or several paragraphs.
    const blockColumn = new Array(map.width).fill(false);
    const align = new Array<string | null>(map.width).fill(null);
    for (let row = 0; row < map.height; row++) {
      for (let col = 0; col < map.width; col++) {
        const offset = map.map[row * map.width + col];
        const cell = table.nodeAt(offset)!;
        const rect = map.findCell(offset);
        if (rect.left !== col || rect.top !== row) continue;
        if (cell.childCount > 1 || cell.firstChild?.type.name === "math_display") blockColumn[col] = true;
        align[col] ??= cell.attrs.align;
      }
    }
    const share = (1 / map.width).toFixed(3);
    // Mathcha's tables are ruled grids.
    const spec =
      "|" +
      blockColumn
        .map((block, col) =>
          block ? `p{${share}\\linewidth}` : ({ center: "c", right: "r" } as Record<string, string>)[align[col] ?? ""] ?? "l",
        )
        .join("|") +
      "|";
    const lines: string[] = [];
    for (let row = 0; row < map.height; row++) {
      const cells: string[] = [];
      for (let col = 0; col < map.width; ) {
        const offset = map.map[row * map.width + col];
        const cell = table.nodeAt(offset)!;
        const rect = map.findCell(offset);
        const colspan = rect.right - rect.left;
        let content = rect.top === row ? this.cell(cell, blockColumn[col]) : "";
        const rowspan = rect.bottom - rect.top;
        if (rowspan > 1 && rect.top === row) {
          this.packages.add("multirow");
          content = `\\multirow{${rowspan}}{*}{${content}}`;
        }
        if (colspan > 1) {
          const a = ({ center: "c", right: "r" } as Record<string, string>)[cell.attrs.align ?? ""] ?? "l";
          content = `\\multicolumn{${colspan}}{${col === 0 ? "|" : ""}${a}|}{${content}}`;
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
    let body = "\\hline\n";
    lines.forEach((line, row) => {
      body += `${line} ${ruleAfter(row)}\n`;
    });
    return `\\begin{center}\n\\begin{tabular}{${spec}}\n${body}\\end{tabular}\n\\end{center}\n`;
  }

  numberOf(id: string): number | undefined {
    return this.numbers.get(id);
  }
}

/** The document as LaTeX. */
export function documentToLatex(doc: Node, options: LatexExportOptions = {}): string {
  const numbers = new Map<string, number>();
  const citeKeys = new Map(bibliographyEntries(doc).map((entry) => [entry.id, entry.key]));
  // the images' files: figures/fig-1.png, … in the order they appear
  const paths = imageFiles(doc, "figures").paths;
  const writer = new LatexWriter({ imagePath: (key) => paths.get(key) ?? null, ...options }, numbers, citeKeys);
  const body = writer.blocks(doc);
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
