// LaTeX source read into the document (a paste, an opened .tex file): it is turned into the Markdown the document
// already reads (markdown.ts) — sections into headings, \textbf / \emph / \underline into marks, itemize /
// enumerate into lists, quote into a quote, tabular into a table, lstlisting / verbatim into code blocks, math as
// it is ($…$, \(…\), \[…\], equation / align / gather as numbered display formulas, \eqref as references).
// \maketitle prints the title block (\title, \author, \date) where it stands; \begin{abstract} is the abstract; \section*
// an unnumbered section. Commands without a counterpart (\usepackage, \centering…) are dropped; their text stays.
import { resolveImageSource } from "./images.ts";
import { bibliographyMarker, citeMarker, footnoteMarker } from "./latexMarkers.ts";
import { PAGE_COMMAND_PATTERN, pageCommandLatex, parsePageCommand } from "./pageCommands.ts";
import { CJK_TEXT } from "./titleBlock.ts";
import { parseTabular, tableMarker } from "./latexTables.ts";

/** Whether text is LaTeX document source (not just a formula). */
export function looksLikeLatexSource(text: string): boolean {
  return (
    /\\documentclass|\\begin\{document\}/.test(text) ||
    /\\(?:chapter|section|subsection|subsubsection|paragraph)\*?\s*\{/.test(text) ||
    /\\begin\{(?:itemize|enumerate|description|quote|quotation|tabular|table|figure|lstlisting|verbatim|center|abstract)\}/.test(text) ||
    /\\(?:textbf|textit|emph|underline|texttt)\s*\{/.test(text) ||
    /\\maketitle\b/.test(text) ||
    PAGE_COMMAND_PATTERN.test(text)
  );
}

/** The {…} group starting at `open` (a "{"): its content and the index after its "}". */
function group(text: string, open: number): { content: string; end: number } | null {
  if (text[open] !== "{") return null;
  let depth = 0;
  for (let i = open; i < text.length; i++) {
    const c = text[i];
    if (c === "\\") {
      i++;
      continue;
    }
    if (c === "{") depth++;
    else if (c === "}" && --depth === 0) return { content: text.slice(open + 1, i), end: i + 1 };
  }
  return null;
}

/** Skip spaces (not blank lines) from i. */
function skipSpaces(text: string, i: number): number {
  while (i < text.length && (text[i] === " " || text[i] === "\t")) i++;
  return i;
}

/** An optional [...] argument at i: the index after it. */
function skipOptional(text: string, i: number): number {
  const j = skipSpaces(text, i);
  if (text[j] !== "[") return i;
  const close = text.indexOf("]", j);
  return close < 0 ? i : close + 1;
}

/** The text width of an A4 page with Word's margins (the document's default), in pt, for widths given in units. */
const DEFAULT_TEXT_WIDTH_PT = ((210 - 63.6) / 25.4) * 72.27;

const UNIT_PT: Record<string, number> = { pt: 1, bp: 72.27 / 72, mm: 72.27 / 25.4, cm: 72.27 / 2.54, in: 72.27, pc: 12, em: 10, ex: 4.3 };

/** \includegraphics' width as a fraction of the line: 0.5\linewidth, \textwidth, 5cm, 120pt (null: not given). */
function graphicsWidth(options: string): number | null {
  const width = /(?:^|,)\s*width\s*=\s*([^,]+)/.exec(options)?.[1]?.trim();
  if (!width) return null;
  const relative = /^([\d.]*)\s*\\(?:linewidth|textwidth|columnwidth|hsize)$/.exec(width);
  if (relative) return Math.min(1, Number(relative[1] || "1"));
  const absolute = /^([\d.]+)\s*(pt|bp|mm|cm|in|pc|em|ex)$/.exec(width);
  if (absolute) return Math.min(1, (Number(absolute[1]) * UNIT_PT[absolute[2]]) / DEFAULT_TEXT_WIDTH_PT);
  return null;
}

/** \includegraphics[…]{path} as the HTML image the Markdown reader takes (markdown.ts). */
function graphicsHtml(options: string, path: string, align: string | null = null): string {
  // an image whose file is not at hand (LaTeX pasted without its figures) is left out, as before
  if (!resolveImageSource(path)) return "";
  const width = graphicsWidth(options);
  const escaped = path.trim().replace(/"/g, "&quot;");
  return `<img src="${escaped}" alt=""${width ? ` width="${Math.round(width * 1000) / 10}%"` : ""}${align ? ` align="${align}"` : ""}>`;
}

const GRAPHICS = /\\includegraphics\*?\s*(?:\[([^\]]*)\])?\s*\{([^{}]*)\}/g;

const PAGE_COMMANDS = /\\(?:newpage|clearpage|cleardoublepage|pagebreak)\b(?:\[\d\])?|\\pagenumbering\s*\{[^{}]*\}|\\(?:this)?pagestyle\s*\{[^{}]*\}|\\setcounter\s*\{\s*page\s*\}\s*\{[^{}]*\}/g;

/** \newpage, \pagenumbering{…} and the other page commands: blocks of their own (pageCommands.ts). */
function pageCommands(text: string, kept: string[]): string {
  return text.replace(PAGE_COMMANDS, (whole) => {
    const parsed = parsePageCommand(whole);
    if (!parsed) return whole;
    kept.push(`\n\n<!-- ${pageCommandLatex(parsed.command, parsed.value)} -->\n\n`);
    return `\u0000${kept.length - 1}\u0000`;
  });
}

/** \graphicspath{{dir/}{dir2/}}: the folders \includegraphics looks in. */
export function graphicsPaths(source: string): string[] {
  const match = /\\graphicspath\s*\{((?:\s*\{[^{}]*\})+)\s*\}/.exec(source);
  return match ? [...match[1].matchAll(/\{([^{}]*)\}/g)].map((m) => m[1]) : [];
}

/** A caption's LaTeX as Markdown now (for a title, or a table's marker): its kept pieces back in, its \label out. */
function captionMarkdown(latex: string, kept: string[]): string {
  let text = finishText(inline(latex.replace(/\\label\s*\{[^{}]*\}/g, "").replace(/\s*\n\s*/g, " "), kept));
  for (let pass = 0; pass < 3 && text.includes("\u0000"); pass++) text = text.replace(/\u0000(\d+)\u0000/g, (_m, n: string) => kept[Number(n)]);
  return text.trim();
}

/** A figure's captions: its last \caption, with those before it (subfigures') after it as (a) …, (b) … */
function captionOf(body: string): string {
  const captions: string[] = [];
  for (const at of body.matchAll(/\\(?:sub)?caption\s*(?:\[[^\]]*\])?\s*(?=\{)/g)) {
    const arg = group(body, at.index! + at[0].length);
    if (arg) captions.push(arg.content.replace(/\\label\s*\{[^{}]*\}/g, "").replace(/\s*\n\s*/g, " ").trim());
  }
  if (!captions.length) return "";
  const main = captions.pop()!;
  return [main, ...captions.map((caption, i) => `(${String.fromCharCode(97 + i)}) ${caption}`)].filter(Boolean).join(" ");
}

/**
 * Images (DESIGN.md §11): a wrapfigure becomes an image at the start of the next paragraph (wrapped on its side);
 * figure / center / \centering groups and lines holding only \includegraphics become image lines; an
 * \includegraphics among text stays in the text. A figure's caption is the figure's (figures.ts): the centred line
 * "Figure 1: …" the Markdown reader takes as the caption of the image line before it; a wrapfigure's the image's
 * title.
 */
function figures(text: string, kept: string[]): string {
  const keep = (markdown: string) => {
    kept.push(markdown);
    return `\u0000${kept.length - 1}\u0000`;
  };
  // an image board as VisualTeX writes it: a picture of \linewidth units, one \put per image
  text = text.replace(
    /\{\\noindent\s*\\setlength\s*\{\\unitlength\}\s*\{\\linewidth\}%?\s*\\begin\{picture\}\s*\(\s*1\s*,\s*([\d.]+)\s*\)([\s\S]*?)\\end\{picture\}\s*\\par\s*\}/g,
    (whole, height: string, body: string) => {
      const images: string[] = [];
      for (const put of body.matchAll(/\\put\s*\(\s*([\d.]+)\s*,\s*([\d.]+)\s*\)\s*\{\s*\\includegraphics\*?\s*(?:\[([^\]]*)\])?\s*\{([^{}]*)\}\s*\}/g)) {
        const html = graphicsHtml(put[3] ?? "", put[4]);
        if (!html) return whole;
        images.push(html.replace(/^<img/, `<img data-x="${put[1]}" data-y="${put[2]}"`));
      }
      return images.length ? keep(`\n\n<div data-vtx-board="1" data-height="${height}">${images.join("")}</div>\n\n`) : whole;
    },
  );
  // wrapfigure: its image goes to the start of the paragraph after it
  text = text.replace(/\\begin\{wrapfigure\}\s*(?:\[[^\]]*\])?\s*\{([rlioRLIO])\}\s*(?:\[[^\]]*\])?\s*\{([^{}]*)\}([\s\S]*?)\\end\{wrapfigure\}\s*/g, (_m, side: string, width: string, body: string) => {
    const graphics = [...body.matchAll(GRAPHICS)][0];
    if (!graphics) return "";
    const fraction = graphicsWidth(`width=${width.trim()}`);
    const inner = graphicsWidth(graphics[1] ?? "");
    const options = fraction ? `width=${(fraction * (inner ?? 1)).toFixed(3)}\\linewidth` : graphics[1] ?? "";
    const html = graphicsHtml(options, graphics[2], /[lLiI]/.test(side) ? "left" : "right");
    // its caption: the image's title, "Figure 1: …" (the Markdown reader's caption of a wrapped image)
    const caption = captionOf(body);
    const title = caption ? ` title="${`Figure 1: ${captionMarkdown(caption, kept)}`.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}"` : "";
    return html ? keep(html.replace(/>$/, `${title}>`) + " ") : "";
  });
  // figure / center / a \centering group with images: an image line (and the caption under it)
  text = text.replace(/\\begin\{(figure|center|subfigure|minipage)\*?\}(?:\[[^\]]*\])?(?:\{[^{}]*\})?([\s\S]*?)\\end\{\1\*?\}/g, (whole, env: string, body: string) => {
    if (env === "subfigure" || env === "minipage") return whole;
    const images = [...body.matchAll(GRAPHICS)];
    // a board (taken aside above) in a figure
    const board = env === "figure" ? [...body.matchAll(/\u0000(\d+)\u0000/g)].find((m) => /data-vtx-board/.test(kept[Number(m[1])] ?? "")) : undefined;
    if (!images.length && !board) return whole;
    const align = /\\raggedleft/.test(body) ? "right" : /\\raggedright/.test(body) ? "left" : env === "center" || /\\centering/.test(body) || env === "figure" ? "center" : "left";
    const caption = captionOf(body);
    const html = images.map((g) => graphicsHtml(g[1] ?? "", g[2])).filter(Boolean);
    const line = board ? `\n\n${board[0]}\n\n` : html.length ? keep(`\n\n<p align="${align}">${html.join(" ")}</p>\n\n`) : "";
    // a table in the same float stays (read with the tables, below)
    const tabular = [...body.matchAll(/\\begin\{(tabular\*?|tabularx|tabulary)\}[\s\S]*?\\end\{\1\}/g)].map((m) => `\n\n${m[0]}\n\n`).join("");
    // the caption: the figure's (figures.ts), a centred line the Markdown reader takes as it; without images, a line
    if (!line) return (caption ? `\n\n*${caption}*\n\n` : "") + tabular;
    return line + (caption && env === "figure" ? `\n\n<p align="center">Figure 1: ${caption}</p>\n\n` : caption ? `\n\n*${caption}*\n\n` : "") + tabular;
  });
  // {\centering …\par} / \noindent … lines with only images
  text = text.replace(/\{\\(centering|raggedleft|raggedright)\b((?:[^{}]|\{[^{}]*\})*?)\\par\s*\}/g, (whole, switch_: string, body: string) => {
    const images = [...body.matchAll(GRAPHICS)];
    if (!images.length || body.replace(GRAPHICS, "").replace(/\\(hspace\*?\{[^{}]*\}|hfill|quad|qquad|noindent)|%|\s/g, "")) return whole;
    const align = switch_ === "centering" ? "center" : switch_ === "raggedleft" ? "right" : "left";
    const html = images.map((g) => graphicsHtml(g[1] ?? "", g[2])).filter(Boolean);
    return html.length ? keep(`\n\n<p align="${align}">${html.join(" ")}</p>\n\n`) : "";
  });
  // any other \includegraphics: in the text where it is
  return text.replace(GRAPHICS, (_m, options: string | undefined, path: string) => {
    const html = graphicsHtml(options ?? "", path);
    return html ? keep(html) : "";
  });
}

const DISPLAY_ENVIRONMENTS = "equation|align|gather|multline|eqnarray|flalign|alignat";

/** Math kept aside while the text around it is converted (placeholders \u0000n\u0000). */
function protectMath(text: string, kept: string[]): string {
  const keep = (markdown: string) => {
    kept.push(markdown);
    return `\u0000${kept.length - 1}\u0000`;
  };
  // Display environments: numbered unless starred; align and gather become aligned and gathered inside $$.
  text = text.replace(new RegExp(`\\\\begin\\{(${DISPLAY_ENVIRONMENTS})(\\*?)\\}([\\s\\S]*?)\\\\end\\{\\1\\2\\}`, "g"), (_m, env: string, star: string, body: string) => {
    const inner = body.trim();
    if (star) {
      const name = env === "align" || env === "flalign" || env === "alignat" || env === "eqnarray" ? "aligned" : env === "gather" ? "gathered" : null;
      return keep(`\n\n$$\n${name ? `\\begin{${name}}${inner}\\end{${name}}` : inner}\n$$\n\n`);
    }
    const kind = env === "equation" || env === "multline" ? "equation" : env === "gather" ? "gather" : "align";
    return keep(`\n\n$$\n\\begin{${kind}}${inner}\\end{${kind}}\n$$\n\n`);
  });
  text = text.replace(/\$\$([\s\S]+?)\$\$/g, (_m, body: string) => keep(`\n\n$$\n${body.trim()}\n$$\n\n`));
  text = text.replace(/\\\[([\s\S]+?)\\\]/g, (_m, body: string) => keep(`\n\n$$\n${body.trim()}\n$$\n\n`));
  text = text.replace(/\\\(([\s\S]+?)\\\)/g, (_m, body: string) => keep(`$${body.trim()}$`));
  text = text.replace(/(?<!\\)\$([^$]+?)(?<!\\)\$/g, (_m, body: string) => keep(`$${body.trim()}$`));
  return text;
}

/** Code environments, kept verbatim as fences. */
function protectCode(text: string, kept: string[]): string {
  return text.replace(/\\begin\{(lstlisting|verbatim|minted)\}(\[[^\]]*\]|\{[^}]*\})*\n?([\s\S]*?)\\end\{\1\}/g, (m, _env: string, _opts: string, body: string) => {
    const language = /language\s*=\s*([\w+#-]+)/i.exec(m)?.[1] ?? /\\begin\{minted\}(?:\[[^\]]*\])?\{([^}]*)\}/.exec(m)?.[1] ?? "";
    kept.push(`\n\n\`\`\`${language.toLowerCase()}\n${body.replace(/\n$/, "")}\n\`\`\`\n\n`);
    return `\u0000${kept.length - 1}\u0000`;
  });
}

/** Commands with one argument turned into Markdown, innermost first so nesting works. */
const WRAPPERS: Record<string, (arg: string) => string> = {
  // HTML tags rather than **…** / *…*: Markdown's emphasis does not open or close beside CJK punctuation (、**粗体**、)
  textbf: (a) => `<b>${a}</b>`,
  textit: (a) => `<i>${a}</i>`,
  emph: (a) => `<i>${a}</i>`,
  textsl: (a) => `<i>${a}</i>`,
  underline: (a) => `<u>${a}</u>`,
  uline: (a) => `<u>${a}</u>`,
  texttt: (a) => `\`${a}\``,
  sout: (a) => `<s>${a}</s>`,
  url: (a) => `<${a}>`,
  caption: (a) => `*${a}*`,
  mbox: (a) => a,
  textrm: (a) => a,
  textnormal: (a) => a,
  text: (a) => a,
};

/** Accents: \"o, \'{e}, \c{c}… as the accented letter. */
const ACCENTS: Record<string, string> = { '"': "\u0308", "'": "\u0301", "`": "\u0300", "^": "\u0302", "~": "\u0303", "=": "\u0304", ".": "\u0307", c: "\u0327", v: "\u030C", u: "\u0306", H: "\u030B", r: "\u030A", k: "\u0328" };
const LETTERS: Record<string, string> = { ss: "ß", ae: "æ", AE: "Æ", oe: "œ", OE: "Œ", o: "ø", O: "Ø", aa: "å", AA: "Å", l: "ł", L: "Ł", i: "ı", j: "ȷ" };

/** Commands without arguments that print a sign or a logo. */
const SIGNS: Record<string, string> = {
  TeX: "TeX", LaTeX: "LaTeX", LaTeXe: "LaTeX2ε", BibTeX: "BibTeX", XeTeX: "XeTeX", XeLaTeX: "XeLaTeX", LuaTeX: "LuaTeX", LuaLaTeX: "LuaLaTeX", pdfTeX: "pdfTeX", pdfLaTeX: "pdfLaTeX", AmS: "AMS", AmSTeX: "AMS-TeX",
  copyright: "©", textcopyright: "©", textregistered: "®", texttrademark: "™", dag: "†", ddag: "‡", textdagger: "†", textdaggerdbl: "‡", S: "§", P: "¶", pounds: "£", textsterling: "£", euro: "€", texteuro: "€", textdegree: "°", textbullet: "•", textperiodcentered: "·",
  textbackslash: "\\\\", textendash: "–", textemdash: "—", textellipsis: "…", textbar: "|", textless: "<", textgreater: ">", textquotedblleft: "“", textquotedblright: "”", textquoteleft: "‘", textquoteright: "’", guillemotleft: "«", guillemotright: "»",
  quad: "\u2003", qquad: "\u2003\u2003", enspace: "\u2002", thinspace: "\u2009", slash: "/", textasciicircum: "^",
};

/** An accent's letter (`at` is after the accent command): the accented letter and the index after it. */
function accented(text: string, at: number, mark: string): { text: string; end: number } | null {
  let k = skipSpaces(text, at);
  let letter = "";
  if (text[k] === "{") {
    const arg = group(text, k);
    if (!arg) return null;
    letter = arg.content.trim();
    k = arg.end;
  } else if (/[A-Za-z]/.test(text[k] ?? "")) {
    letter = text[k];
    k++;
  } else return null;
  const special = /^\\([a-zA-Z]+)$/.exec(letter);
  // the dotless i / j take the accent as i / j
  if (special) letter = special[1] === "i" || special[1] === "j" ? special[1] : (LETTERS[special[1]] ?? special[1]);
  if (!letter || letter.length > 2) return null;
  return { text: (letter + mark).normalize("NFC"), end: k };
}

/** A note's or an entry's text (Markdown with placeholders): plain text with $…$ formulas. */
function plainText(markdown: string, kept: string[]): string {
  // marks and grouping braces go before the formulas come back (they keep theirs)
  let text = finishText(
    markdown
      .replace(/\*\*|<\/?(u|sup|sub|b|i|s)>|<\/?span[^>]*>|[*`]/g, "")
      .replace(/<((?:https?|ftp|mailto):[^>\s]+)>/g, "$1")
      .replace(/\[([^\]]*)\]\(([^)\s]*)\)/g, "$1"),
  ).replace(/\\\n/g, " ");
  for (let pass = 0; pass < 3 && text.includes("\u0000"); pass++) text = text.replace(/\u0000(\d+)\u0000/g, (_m, n: string) => kept[Number(n)]);
  return text.replace(/\s+/g, " ").trim();
}

/** The text's last steps: grouping braces go, literal ones come back; ties, quotes and dashes. */
function finishText(text: string): string {
  return text
    .replace(/[{}]/g, "")
    .replace(/\u0004/g, "{")
    .replace(/\u0005/g, "}")
    .replace(/(?<!\\)~/g, " ")
    .replace(/``/g, "“")
    .replace(/''/g, "”")
    .replace(/---/g, "—")
    .replace(/--/g, "–");
}

/**
 * What \ref prints for labels of sections, figures and tables (equations are references to their formulas): the
 * numbers LaTeX gives them, counted through the source. Set while a source is read.
 */
let labelNumbers = new Map<string, { number: string; kind: "section" | "figure" | "table" }>();

function countLabels(text: string): Map<string, { number: string; kind: "section" | "figure" | "table" }> {
  const labels = new Map<string, { number: string; kind: "section" | "figure" | "table" }>();
  const hasChapters = /\\chapter\*?\s*[[{]/.test(text);
  const levels = hasChapters ? ["chapter", "section", "subsection", "subsubsection"] : ["section", "subsection", "subsubsection"];
  const counts = [0, 0, 0, 0];
  let appendix = false;
  let figures = 0;
  let tables = 0;
  let current: { number: string; kind: "section" | "figure" | "table" } | null = null;
  let float: "figure" | "table" | null = null;
  let math = false;
  const pattern = /\\(chapter|section|subsection|subsubsection)(\*?)\s*(?:\[[^\]]*\])?\s*\{|\\appendix\b|\\begin\{(figure|table|equation|align|gather|multline|eqnarray|flalign|alignat)\*?\}|\\end\{(figure|table|equation|align|gather|multline|eqnarray|flalign|alignat)\*?\}|\\caption\b|\\label\s*\{([^{}]*)\}/g;
  const prefix = (level: number) => {
    const top = appendix ? String.fromCharCode(64 + counts[0]) : String(counts[0]);
    return [top, ...counts.slice(1, level + 1).map(String)].join(".");
  };
  for (const match of text.matchAll(pattern)) {
    if (match[1]) {
      if (match[2]) continue;
      const level = levels.indexOf(match[1]);
      if (level < 0) continue;
      counts[level]++;
      for (let k = level + 1; k < counts.length; k++) counts[k] = 0;
      if (hasChapters && level === 0) figures = tables = 0;
      current = { number: prefix(level), kind: "section" };
    } else if (match[0] === "\\appendix") {
      appendix = true;
      counts.fill(0);
    } else if (match[3]) {
      if (match[3] === "figure" || match[3] === "table") float = match[3];
      else math = true;
    } else if (match[4]) {
      if (match[4] === "figure" || match[4] === "table") float = null;
      else math = false;
    } else if (match[0].startsWith("\\caption") && float) {
      const n = float === "figure" ? ++figures : ++tables;
      current = { number: hasChapters ? `${prefix(0)}.${n}` : String(n), kind: float };
    } else if (match[5] && !math && current) labels.set(match[5].trim(), current);
  }
  return labels;
}

const REFERENCE_NAMES = { section: ["Section", "section"], figure: ["Figure", "figure"], table: ["Table", "table"] };

/** Inline commands: wrappers, links, references, escapes. Generated math goes into `kept` (placeholders). */
function inline(text: string, kept: string[]): string {
  const keep = (markdown: string) => {
    kept.push(markdown);
    return `\u0000${kept.length - 1}\u0000`;
  };
  let out = "";
  for (let i = 0; i < text.length; ) {
    if (text[i] !== "\\") {
      out += text[i++];
      continue;
    }
    const name = /^\\([A-Za-z]+)\*?/.exec(text.slice(i));
    if (!name) {
      // \% \& \# \_ \$ \{ \} \\ and the like
      const c = text[i + 1] ?? "";
      if (c === "\\") {
        out += "\\\n";
        i += 2;
      } else if (c === "{" || c === "}") {
        // literal braces survive the removal of grouping braces at the end
        out += c === "{" ? "\u0004" : "\u0005";
        i += 2;
      } else if ("%&".includes(c)) {
        out += c;
        i += 2;
      } else if ("#_$".includes(c)) {
        out += `\\${c}`;
        i += 2;
      } else if (c === ",") {
        out += " ";
        i += 2;
      } else if (ACCENTS[c] && c !== "c" && accented(text, i + 2, ACCENTS[c])) {
        const accent = accented(text, i + 2, ACCENTS[c])!;
        out += accent.text;
        i = accent.end;
      } else {
        out += c;
        i += 2;
      }
      continue;
    }
    const command = name[1];
    let j = i + name[0].length;
    if (/^[cvuHrk]$/.test(command)) {
      const accent = accented(text, j, ACCENTS[command]);
      if (accent) {
        out += accent.text;
        i = accent.end;
        continue;
      }
    }
    const sign = LETTERS[command] ?? SIGNS[command];
    if (sign !== undefined) {
      if (text[j] === "{" && text[j + 1] === "}") j += 2;
      out += sign;
      i = text[j] === " " ? j + 1 : j;
      continue;
    }
    if (command === "textsuperscript" || command === "textsubscript") {
      // raised / lowered text (<sup> / <sub>, read as the document's marks)
      const arg = group(text, skipSpaces(text, j));
      if (arg) {
        const tag = command === "textsuperscript" ? "sup" : "sub";
        out += `<${tag}>${inline(arg.content, kept)}</${tag}>`;
        i = arg.end;
        continue;
      }
    }
    // \thanks (in \author or \title): LaTeX prints it as a footnote
    if (command === "footnote" || command === "thanks") {
      const arg = group(text, skipSpaces(text, skipOptional(text, j)));
      if (arg) {
        out += footnoteMarker(plainText(inline(arg.content, kept), kept));
        i = arg.end;
        continue;
      }
    }
    const wrapper = WRAPPERS[command];
    if (wrapper) {
      const arg = group(text, skipSpaces(text, j));
      if (arg) {
        out += wrapper(inline(arg.content, kept));
        i = arg.end;
        continue;
      }
    }
    if (command === "textcolor" || command === "colorbox") {
      let k = skipSpaces(text, j);
      const model = text[k] === "[" ? /^\[([^\]]*)\]/.exec(text.slice(k)) : null;
      if (model) k = skipSpaces(text, k + model[0].length);
      const color = group(text, k);
      const body = color && group(text, skipSpaces(text, color.end));
      if (color && body) {
        const value = model && /html/i.test(model[1]) ? `#${color.content.trim()}` : color.content.trim();
        out += `<span style="${command === "textcolor" ? "color" : "background-color"}:${value}">${inline(body.content, kept)}</span>`;
        i = body.end;
        continue;
      }
    }
    if (command === "href") {
      const url = group(text, skipSpaces(text, j));
      const label = url && group(text, skipSpaces(text, url.end));
      if (url && label) {
        out += `[${inline(label.content, kept)}](${url.content.replace(/\s/g, "%20")})`;
        i = label.end;
        continue;
      }
    }
    if (command === "ref" || command === "eqref" || command === "autoref" || command === "cref" || command === "Cref") {
      const arg = group(text, skipSpaces(text, j));
      const counted = arg && command !== "eqref" ? labelNumbers.get(arg.content.trim()) : undefined;
      if (arg && counted) {
        const name = command === "autoref" || command === "Cref" ? REFERENCE_NAMES[counted.kind][0] : command === "cref" ? REFERENCE_NAMES[counted.kind][1] : "";
        out += name ? `${name} ${counted.number}` : counted.number;
        i = arg.end;
        continue;
      }
      if (arg) {
        out += keep(`$\\eqref{${arg.content.trim()}}$`);
        i = arg.end;
        continue;
      }
    }
    if (/^(cite|citep|citet|citealp|citealt|parencite|textcite|autocite|footcite|supercite|upcite)$/.test(command)) {
      j = skipOptional(text, skipOptional(text, j));
      const arg = group(text, skipSpaces(text, j));
      if (arg) {
        const keys = arg.content.split(",").map((key) => key.trim()).filter(Boolean);
        out += citeMarker(keys);
        i = arg.end;
        continue;
      }
    }
    if (command === "newline" || command === "linebreak") {
      out += "\\\n";
      i = j;
      continue;
    }
    if (command === "par") {
      out += "\n\n";
      i = j;
      continue;
    }
    if (command === "ldots" || command === "dots") {
      out += "…";
      i = j;
      continue;
    }
    // Commands with arguments that leave nothing (\label, \includegraphics, \vspace, …), and plain switches
    // (\noindent, \centering, \maketitle, …): dropped.
    j = skipOptional(text, j);
    const arg = group(text, skipSpaces(text, j));
    // \begin{…} / \end{…} of an environment with no counterpart (a class's own, minipage contents…): its text stays
    const dropsArgument = /^(begin|end|label|includegraphics|vspace|hspace|setlength|addtolength|bibliography|bibliographystyle|input|include|usepackage|documentclass|newcommand|renewcommand|pagestyle|thispagestyle|setcounter|addbibresource|nocite|printbibliography|pageref)$/.test(command);
    i = dropsArgument && arg ? arg.end : j;
    if (!dropsArgument && arg) {
      out += inline(arg.content, kept);
      i = arg.end;
    }
  }
  return out;
}

/** \begin{thebibliography} as the document's reference list (its \bibitem keys kept for the citations). */
function bibliographies(text: string, kept: string[]): string {
  return text.replace(/\\begin\{thebibliography\}\s*(\{[^{}]*\})?([\s\S]*?)\\end\{thebibliography\}/g, (_m, _widest: string, body: string) => {
    const entries: Array<{ key: string; text: string }> = [];
    const items = body.split(/\\bibitem\b/).slice(1);
    for (const item of items) {
      let i = skipOptional(item, 0);
      const key = group(item, skipSpaces(item, i));
      if (!key) continue;
      i = key.end;
      entries.push({ key: key.content.trim(), text: plainText(inline(item.slice(i).replace(/\\newblock\b/g, " "), kept), kept) });
    }
    if (!entries.length) return "";
    kept.push(`\n\n${bibliographyMarker(entries)}\n\n`);
    return `\u0000${kept.length - 1}\u0000`;
  });
}

/** Where an environment ends, nested ones of the same name skipped: the index after its \end, or -1. */
function environmentEnd(text: string, name: string, from: number): number {
  const pattern = new RegExp(`\\\\(begin|end)\\{${name.replace(/\*/g, "\\*")}\\}`, "g");
  pattern.lastIndex = from;
  let depth = 1;
  for (let match = pattern.exec(text); match; match = pattern.exec(text)) {
    depth += match[1] === "begin" ? 1 : -1;
    if (depth === 0) return match.index + match[0].length;
  }
  return -1;
}

/** The tabular-like environments (latexTables.ts) and the arguments before their column specification. */
const TABULARS: Record<string, number> = { tabular: 0, "tabular*": 1, tabularx: 1, tabulary: 1, longtable: 0, xtabular: 0, supertabular: 0, tabu: 0 };
const TABULAR_BEGIN = /\\begin\{(tabular\*?|tabularx|tabulary|longtable|xtabular|supertabular|tabu)\}/g;

/**
 * Tables (latexTables.ts): every tabular, tabular*, tabularx, tabulary and longtable as the document's table with
 * its caption — a table (or figure) float's \caption, a longtable's own — its spans, column alignment and rules,
 * carried to the Markdown reader as a marker; threeparttable's notes and whatever else a float holds as text after it.
 */
function tables(text: string, kept: string[]): string {
  const keep = (markdown: string) => {
    kept.push(markdown);
    return `\u0000${kept.length - 1}\u0000`;
  };
  const convert = (latex: string) => captionMarkdown(latex, kept);
  /** The tabulars in a text as markers (the first one with the caption). */
  const tabulars = (body: string, floatCaption: string | null): { text: string; found: number } => {
    let out = "";
    let last = 0;
    let found = 0;
    TABULAR_BEGIN.lastIndex = 0;
    for (let match = TABULAR_BEGIN.exec(body); match; match = TABULAR_BEGIN.exec(body)) {
      const name = match[1];
      const end = environmentEnd(body, name, match.index + match[0].length);
      if (end < 0) break;
      let at = match.index + match[0].length;
      // the width of tabular* / tabularx / tabulary, a position [t]
      for (let k = 0; k < (TABULARS[name] ?? 0); k++) {
        const arg = group(body, skipSpaces(body, at));
        if (arg) at = arg.end;
      }
      at = skipOptional(body, at);
      // tabu's "to \linewidth" / "spread 0pt"
      const width = /^\s*(?:to|spread)\s*[-\d.]*\s*\\?[a-zA-Z]+/.exec(body.slice(at));
      if (name === "tabu" && width) at += width[0].length;
      const spec = group(body, skipSpaces(body, at));
      const inside = body.slice(spec ? spec.end : at, end - `\\end{${name}}`.length);
      const parsed = parseTabular(spec?.content ?? "", inside, convert);
      const caption = found === 0 ? (floatCaption ?? parsed.caption) : parsed.caption;
      parsed.table.caption = caption ? convert(caption) : null;
      out += body.slice(last, match.index) + (parsed.table.rows.length ? keep(`\n\n${tableMarker(parsed.table)}\n\n`) : "");
      found++;
      last = end;
      TABULAR_BEGIN.lastIndex = end;
    }
    return { text: out + body.slice(last), found };
  };
  /** threeparttable's notes: \item[a] … as lines after the table. */
  const notes = (body: string) =>
    body.replace(/\\begin\{tablenotes\}(?:\[[^\]]*\])?([\s\S]*?)\\end\{tablenotes\}/g, (_m, items: string) =>
      items
        .split(/\\item\b/)
        .slice(1)
        .map((item) => {
          const mark = /^\s*\[([^\]]*)\]/.exec(item);
          return `\n\n${mark ? `<sup>${mark[1]}</sup> ` : ""}${item.slice(mark ? mark[0].length : 0).trim()}\n\n`;
        })
        .join(""),
    );
  // a table (or a figure) float holding tables: its caption the first table's; the rest of what it holds after them
  text = text.replace(/\\begin\{(table|figure)(\*?)\}(?:\[[^\]]*\])?([\s\S]*?)\\end\{\1\2\}/g, (whole, _env: string, _star: string, body: string) => {
    if (!/\\begin\{(tabular\*?|tabularx|tabulary|longtable|xtabular|supertabular|tabu)\}/.test(body)) return whole;
    const caption = captionOf(body.replace(/\\begin\{(tabular\*?|tabularx|tabulary|longtable)\}[\s\S]*?\\end\{\1\}/g, ""));
    const rest = body
      .replace(/\\(?:sub)?caption\s*(?:\[[^\]]*\])?\s*\{(?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*\}/g, "")
      .replace(/\\label\s*\{[^{}]*\}/g, "")
      .replace(/\\(?:begin|end)\{(?:threeparttable|adjustbox|center|small|footnotesize)\}(?:\{[^{}]*\})?/g, "")
      // \resizebox{\linewidth}{!}{…} and \scalebox{.8}{…} around a table: the group's braces go later
      .replace(/\\(?:resizebox\*?|scalebox)\s*\{[^{}]*\}(?:\s*\{[^{}]*\})?/g, "")
      .replace(/\\(?:centering|small|footnotesize|scriptsize|normalsize)\b/g, "");
    const result = tabulars(rest, caption || null);
    return `\n\n${notes(result.text)}\n\n`;
  });
  // the other tables (a longtable with its own caption)
  return notes(tabulars(text, null).text);
}

/** \section{…} and its kin on lines of their own, as headings (\paragraph{…} as a bold run-in); \section*{…} as an
 *  unnumbered heading (Pandoc's "{-}", which the Markdown reader takes). */
function sections(text: string): string {
  const pattern = /\\(chapter|section|subsection|subsubsection|paragraph)(\*?)\s*(\[[^\]]*\])?\s*(?=\{)/g;
  let out = "";
  let last = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(text))) {
    const title = group(text, match.index + match[0].length);
    if (!title) continue;
    const kind = match[1];
    const marker = kind === "chapter" || kind === "section" ? "#" : kind === "subsection" ? "##" : kind === "subsubsection" ? "###" : null;
    // (literal braces: grouping braces go at the end)
    const unnumbered = match[2] ? " \u0004-\u0005" : "";
    out += text.slice(last, match.index) + (marker ? `\n\n${marker} ${title.content.trim()}${unnumbered}\n\n` : `\n\n**${title.content.trim()}** `);
    last = title.end;
    pattern.lastIndex = title.end;
  }
  return out + text.slice(last);
}

/** Lists, quotes and the environments that only arrange text, line by line. */
function blocks(text: string): string {
  const out: string[] = [];
  const lists: Array<"-" | "1."> = [];
  let quote = 0;
  const prefix = () => "> ".repeat(quote);
  // center / flushright: their paragraphs aligned (<p align>, what the Markdown reader takes as an aligned paragraph)
  const aligns: Array<"center" | "right" | null> = [];
  let aligned: string[] = [];
  const flushAligned = () => {
    if (!aligned.length) return;
    const align = aligns[aligns.length - 1];
    const body = aligned.join(" ").trim();
    aligned = [];
    if (!body) return;
    // a block kept aside (a figure, a display formula) is not a paragraph
    if (!align || /^\u0000\d+\u0000$/.test(body)) out.push(`${prefix()}${body}`);
    else out.push("", `${prefix()}<p align="${align}">${body.replace(/\\\\\s*/g, "<br>")}</p>`, "");
  };
  for (const raw of text.split("\n")) {
    let line = raw;
    // Several \begin / \end / \item on one line are taken apart first.
    const pieces = line.split(/(?=\\begin\{|\\end\{|\\item\b)/);
    for (let piece of pieces) {
      const begin = /^\\begin\{(itemize|enumerate|description|quote|quotation|center|flushleft|flushright|abstract|figure|table|minipage)\*?\}(\[[^\]]*\])?(\{[^}]*\})?/.exec(piece);
      if (begin) {
        const env = begin[1];
        flushAligned();
        if (env === "center" || env === "flushright" || env === "flushleft") aligns.push(env === "center" ? "center" : env === "flushright" ? "right" : null);
        // the abstract: Markdown's form of it (markdown.ts) around its paragraphs
        if (env === "abstract") out.push("", `${prefix()}${MARKDOWN_ABSTRACT_OPEN}`);
        if (env === "itemize" || env === "description") lists.push("-");
        else if (env === "enumerate") lists.push("1.");
        else if (env === "quote" || env === "quotation") quote++;
        out.push("");
        piece = piece.slice(begin[0].length);
      }
      const end = /^\\end\{(itemize|enumerate|description|quote|quotation|center|flushleft|flushright|abstract|figure|table|minipage)\*?\}/.exec(piece);
      if (end) {
        const env = end[1];
        if (env === "center" || env === "flushright" || env === "flushleft") {
          flushAligned();
          aligns.pop();
        }
        if (env === "itemize" || env === "description" || env === "enumerate") lists.pop();
        else if (env === "quote" || env === "quotation") quote = Math.max(0, quote - 1);
        if (env === "abstract") out.push("", `${prefix()}${MARKDOWN_ABSTRACT_CLOSE}`);
        out.push("");
        piece = piece.slice(end[0].length);
      }
      const item = /^\\item\b\s*(\[([^\]]*)\])?/.exec(piece);
      if (item && lists.length) {
        const depth = lists.length - 1;
        const indent = lists.slice(0, depth).map((m) => (m === "-" ? "  " : "   ")).join("");
        const label = item[2] ? `**${item[2]}** ` : "";
        out.push(`${prefix()}${indent}${lists[depth]} ${label}${piece.slice(item[0].length).trim()}`);
        continue;
      }
      if (!piece.trim()) continue;
      if (aligns.length && aligns[aligns.length - 1] && !lists.length) {
        aligned.push(piece.trim());
        continue;
      }
      const indent = lists.length ? lists.map((m) => (m === "-" ? "  " : "   ")).join("") : "";
      out.push(`${prefix()}${indent}${piece.trim()}`);
    }
    if (!raw.trim()) {
      flushAligned();
      out.push(prefix().trim());
    }
  }
  flushAligned();
  return out.join("\n");
}

/**
 * The pieces kept aside back in place. One inside a list item or a quote (a display formula after \item, a table)
 * takes the item's indentation and the quote's "> " on each of its lines, so that it stays in the item; at the
 * start of an item it starts on the item's line.
 */
function restoreKept(text: string, kept: string[]): string {
  return text
    .split("\n")
    .map((line) => {
      if (!line.includes("\u0000")) return line;
      const lead = /^((?:> ?)*)([ \t]*)((?:[-*+]|\d+[.)])[ \t]+)?/.exec(line)!;
      const continuation = lead[1] + " ".repeat(lead[2].length + (lead[3]?.length ?? 0));
      return line.replace(/\u0000(\d+)\u0000/g, (_m, n: string, offset: number) => {
        let piece = kept[Number(n)];
        if (!continuation) return piece;
        if (offset === lead[0].length && lead[3]) piece = piece.replace(/^\n+/, "");
        return piece.replace(/\n/g, `\n${continuation}`);
      });
    })
    .join("\n");
}

/** Markdown has no title: the document's title is written as a centred level-1 heading in HTML (Typora and GitHub
 *  show it so) and read back as the title (markdown.ts). */
export const MARKDOWN_TITLE_OPEN = '<h1 align="center">';
export const MARKDOWN_TITLE_CLOSE = "</h1>";
/** The abstract: a div around its Markdown (markdown.ts). */
export const MARKDOWN_ABSTRACT_OPEN = '<div class="abstract">';
export const MARKDOWN_ABSTRACT_CLOSE = "</div>";

/** `\name[…]{…}` taken out of the text (the first one's argument returned). */
function takeCommand(text: string, name: string): { text: string; argument?: string } {
  const pattern = new RegExp(`\\\\${name}\\s*(?:\\[[^\\]]*\\])?\\s*(?=\\{)`, "g");
  let argument: string | undefined;
  let out = "";
  let last = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(text))) {
    const arg = group(text, match.index + match[0].length);
    if (!arg) continue;
    argument ??= arg.content;
    out += text.slice(last, match.index);
    last = arg.end;
    pattern.lastIndex = arg.end;
  }
  return { text: out + text.slice(last), argument };
}

/** The date \today prints: the day the file is read, in the document's language (ctex: 2026年11月7日). */
function today(chinese: boolean): string {
  const now = new Date();
  if (chinese) return `${now.getFullYear()}年${now.getMonth() + 1}月${now.getDate()}日`;
  return `${now.toLocaleString("en-US", { month: "long" })} ${now.getDate()}, ${now.getFullYear()}`;
}

/**
 * The title block \maketitle prints: the title, each author (\and between them; \\ a line break in one) and the
 * date, as lines marked \u0006 / \u0001 / \u0002 until the end, where they become Markdown's forms. No \date: none
 * (the editor shows no date where none was written).
 */
function titleBlock(title: string | undefined, author: string | undefined, date: string | undefined, chinese: boolean): string {
  const line = (marker: string, text: string) => `\n\n${marker}${text}\n\n`;
  let out = title?.trim() ? line("\u0006", title.replace(/\\\\(\[[^\]]*\])?/g, " ").replace(/\s*\n\s*/g, " ").trim()) : "";
  for (const one of (author ?? "").split(/\\and(?![A-Za-z])/)) {
    const text = one.replace(/\\\\(\[[^\]]*\])?/g, "\u0003").replace(/\s*\n\s*/g, " ").replace(/\s*\u0003\s*/g, "\u0003").trim();
    if (text) out += line("\u0001", text);
  }
  const day = date?.replace(/\\today(?![A-Za-z])/g, today(chinese)).replace(/\\\\(\[[^\]]*\])?/g, " ").replace(/\s*\n\s*/g, " ").trim();
  if (day) out += line("\u0002", day);
  return out;
}

/** LaTeX source as Markdown the document reads. */
export function latexToMarkdown(source: string): string {
  let text = source.replace(/\r\n?/g, "\n");
  // Comments (an unescaped % to the end of the line).
  text = text.replace(/(^|[^\\])%.*$/gm, "$1");
  // The body of a whole document; \title, \author and \date (in the preamble or the body) are shown by \maketitle where it
  // stands (titleBlock), and print nothing themselves.
  const chinese = /\\documentclass\s*(?:\[[^\]]*\])?\s*\{ctex|\\usepackage\s*(?:\[[^\]]*\])?\s*\{ctex/.test(text) || CJK_TEXT.test(text);
  const titled = takeCommand(text, "title");
  const authored = takeCommand(titled.text, "author");
  const dated = takeCommand(authored.text, "date");
  text = dated.text;
  const body = /\\begin\{document\}([\s\S]*?)(\\end\{document\}|$)/.exec(text);
  // page commands of the preamble act from the first page: they open the body
  const preamble = body ? text.slice(0, body.index) : "";
  const early = [...preamble.matchAll(PAGE_COMMANDS)].map((m) => m[0]).join("\n");
  if (body && early) body[1] = `${early}\n\n${body[1]}`;
  if (body) text = body[1];
  text = text.replace(/\\maketitle(?![A-Za-z])/, () => titleBlock(titled.argument, authored.argument, dated.argument, chinese));
  labelNumbers = countLabels(text);
  const kept: string[] = [];
  text = protectCode(text, kept);
  text = protectMath(text, kept);
  text = figures(text, kept);
  text = pageCommands(text, kept);
  text = bibliographies(text, kept);
  text = tables(text, kept);
  text = sections(text);
  // the table of contents where \tableofcontents stands (on a line of its own)
  text = text.replace(/\\tableofcontents\b/g, "\n\n[toc]\n\n");
  // Old switches in a group: {\bf …}, {\it …}, {\tt …}.
  text = text
    .replace(/\{\\(?:bf|bfseries)\s+([^{}]*)\}/g, "**$1**")
    .replace(/\{\\(?:it|itshape|em|sl|slshape)\s+([^{}]*)\}/g, "*$1*")
    .replace(/\{\\(?:tt|ttfamily)\s+([^{}]*)\}/g, "`$1`");
  text = blocks(text);
  text = inline(text, kept);
  // Grouping braces go, literal ones (\{ \}) come back; ties, quotes and dashes.
  text = finishText(text);
  for (let pass = 0; pass < 3 && text.includes("\u0000"); pass++) text = restoreKept(text, kept);
  text = text.replace(/^\u0006[ \t]*(.*)$/m, (_m, line: string) => `${MARKDOWN_TITLE_OPEN}${line.trim()}${MARKDOWN_TITLE_CLOSE}`);
  text = text.replace(/^([\u0001\u0002])[ \t]*(.*)$/gm, (_m, marker: string, line: string) => `<p align="center" class="${marker === "\u0001" ? "author" : "date"}">${line.trim().replace(/\u0003/g, "<br>")}</p>`);
  return text.replace(/\n{3,}/g, "\n\n").trim() + "\n";
}
