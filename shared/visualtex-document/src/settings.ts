// The document's page and body text (DESIGN.md §8): paper, margins, body fonts (one for Chinese, one for Latin
// script), body size and line spacing. Kept as attributes of the doc node (saved with the document, undoable);
// an attribute left null takes the default below, so documents written before these settings keep working.
import type { Node } from "prosemirror-model";
import type { Command } from "prosemirror-state";
import { documentHasCjk } from "./titleBlock.ts";

export type PaperName = "a4" | "a5" | "b5" | "letter";
/** Equation numbers: through the document (1), (2); or within level-1 sections, (1.1) / (1-1). */
export type EquationNumbering = "1" | "1.1" | "1-1";
export type EquationBrackets = "()" | "[]";
/** Section numbers: "1" / "1.1" (default: LaTeX's), "1." / "1.1.", or none. */
export type HeadingNumbering = "1." | "1" | "none";

/** Paper sizes in millimetres (portrait). B5 is the JIS size Word and WPS call B5. */
export const PAPER_SIZES: Record<PaperName, { width: number; height: number; label: string }> = {
  a4: { width: 210, height: 297, label: "A4" },
  a5: { width: 148, height: 210, label: "A5" },
  b5: { width: 182, height: 257, label: "B5" },
  letter: { width: 215.9, height: 279.4, label: "Letter" },
};

export interface DocumentSettings {
  paper: PaperName;
  /** millimetres */
  marginTop: number;
  marginBottom: number;
  marginLeft: number;
  marginRight: number;
  /** body text size in points */
  fontSize: number;
  /** body font for Chinese (and other CJK) characters: a family name the system knows */
  cjkFont: string;
  /** body font for Latin script, digits and punctuation that font has */
  latinFont: string;
  /** line spacing as a multiple of single spacing (1.3 x the font size, as ctex sets it) */
  lineSpacing: number;
  equationNumbering: EquationNumbering;
  equationBrackets: EquationBrackets;
  headingNumbering: HeadingNumbering;
}

/** The defaults: A4 with Word's margins (2.54 cm, 3.18 cm), 五号 (10.5 pt), SimSun and Times New Roman. */
export const DEFAULT_DOCUMENT_SETTINGS: DocumentSettings = {
  paper: "a4",
  marginTop: 25.4,
  marginBottom: 25.4,
  marginLeft: 31.8,
  marginRight: 31.8,
  fontSize: 10.5,
  cjkFont: "SimSun",
  latinFont: "Times New Roman",
  lineSpacing: 1,
  equationNumbering: "1",
  equationBrackets: "()",
  headingNumbering: "1",
};

export type MarginName = "marginTop" | "marginBottom" | "marginLeft" | "marginRight";

/** The text area keeps at least this much (mm) across and down. */
export const MIN_TEXT_MM = 30;

/** A margin as it can be set: 0.1 mm steps, from 0 to what leaves the text area MIN_TEXT_MM (and at most 100). */
export function clampMargin(settings: DocumentSettings, name: MarginName, mm: number): number {
  const paper = PAPER_SIZES[settings.paper];
  const across = name === "marginLeft" || name === "marginRight";
  const other =
    name === "marginLeft" ? settings.marginRight
    : name === "marginRight" ? settings.marginLeft
    : name === "marginTop" ? settings.marginBottom
    : settings.marginTop;
  const max = Math.min(100, (across ? paper.width : paper.height) - other - MIN_TEXT_MM);
  return Math.round(Math.max(0, Math.min(max, mm)) * 10) / 10;
}

/** Word's margin presets (mm). */
export const MARGIN_PRESETS: Array<{ name: string; en: string; zh: string; top: number; bottom: number; left: number; right: number }> = [
  { name: "normal", en: "Normal", zh: "常规", top: 25.4, bottom: 25.4, left: 31.8, right: 31.8 },
  { name: "narrow", en: "Narrow", zh: "窄", top: 12.7, bottom: 12.7, left: 12.7, right: 12.7 },
  { name: "moderate", en: "Moderate", zh: "适中", top: 25.4, bottom: 25.4, left: 19.1, right: 19.1 },
  { name: "wide", en: "Wide", zh: "宽", top: 25.4, bottom: 25.4, left: 50.8, right: 50.8 },
];

/** ctex's \linespread for Chinese text (on the class's \baselineskip, 1.2 x the size). */
export const SINGLE_LINE_HEIGHT = 1.3;

export const LINE_SPACINGS = [1, 1.15, 1.25, 1.5, 2];

/** Chinese named sizes (字号) in points, the way Word and ctex define them. */
export const CHINESE_FONT_SIZES: Array<{ name: string; pt: number; zihao: string }> = [
  { name: "初号", pt: 42, zihao: "0" },
  { name: "小初", pt: 36, zihao: "-0" },
  { name: "一号", pt: 26, zihao: "1" },
  { name: "小一", pt: 24, zihao: "-1" },
  { name: "二号", pt: 22, zihao: "2" },
  { name: "小二", pt: 18, zihao: "-2" },
  { name: "三号", pt: 16, zihao: "3" },
  { name: "小三", pt: 15, zihao: "-3" },
  { name: "四号", pt: 14, zihao: "4" },
  { name: "小四", pt: 12, zihao: "-4" },
  { name: "五号", pt: 10.5, zihao: "5" },
  { name: "小五", pt: 9, zihao: "-5" },
  { name: "六号", pt: 7.5, zihao: "6" },
  { name: "小六", pt: 6.5, zihao: "-6" },
  { name: "七号", pt: 5.5, zihao: "7" },
  { name: "八号", pt: 5, zihao: "8" },
];

/** Body sizes offered in points (besides the named ones). */
export const POINT_FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 13, 14, 16, 18, 20, 24];

/** The doc node's attributes (null: the default). */
export const DOCUMENT_SETTING_ATTRS: Record<keyof DocumentSettings, { default: null }> = {
  paper: { default: null },
  marginTop: { default: null },
  marginBottom: { default: null },
  marginLeft: { default: null },
  marginRight: { default: null },
  fontSize: { default: null },
  cjkFont: { default: null },
  latinFont: { default: null },
  lineSpacing: { default: null },
  equationNumbering: { default: null },
  equationBrackets: { default: null },
  headingNumbering: { default: null },
};

const finite = (value: unknown, min: number, max: number): number | null =>
  typeof value === "number" && Number.isFinite(value) && value >= min && value <= max ? value : null;

const family = (value: unknown): string | null =>
  typeof value === "string" && value.trim() && value.length <= 128 && !/["<>;{}\\]/.test(value) ? value.trim() : null;

/** The document's settings, each one its stored value when that is valid, else the default. */
export function documentSettings(doc: Node): DocumentSettings {
  const a = doc.attrs as Partial<Record<keyof DocumentSettings, unknown>>;
  const d = DEFAULT_DOCUMENT_SETTINGS;
  return {
    paper: typeof a.paper === "string" && a.paper in PAPER_SIZES ? (a.paper as PaperName) : d.paper,
    marginTop: finite(a.marginTop, 0, 100) ?? d.marginTop,
    marginBottom: finite(a.marginBottom, 0, 100) ?? d.marginBottom,
    marginLeft: finite(a.marginLeft, 0, 100) ?? d.marginLeft,
    marginRight: finite(a.marginRight, 0, 100) ?? d.marginRight,
    fontSize: finite(a.fontSize, 4, 72) ?? d.fontSize,
    cjkFont: family(a.cjkFont) ?? d.cjkFont,
    latinFont: family(a.latinFont) ?? d.latinFont,
    lineSpacing: finite(a.lineSpacing, 0.8, 3) ?? d.lineSpacing,
    equationNumbering: a.equationNumbering === "1.1" || a.equationNumbering === "1-1" ? a.equationNumbering : d.equationNumbering,
    equationBrackets: a.equationBrackets === "[]" ? "[]" : d.equationBrackets,
    headingNumbering: a.headingNumbering === "1." || a.headingNumbering === "none" ? a.headingNumbering : d.headingNumbering,
  };
}

/**
 * A display formula's own \tag{…} / \tag*{…} at its top level (outside \begin…\end), as amsmath sets it: it
 * replaces the document's number for that formula (no number is taken; references show it). `label` is what
 * \eqref shows inside its brackets.
 */
export function ownEquationTag(latex: string): { label: string; star: boolean } | null {
  let depth = 0;
  for (let i = 0; i < latex.length; i++) {
    if (latex[i] !== "\\") continue;
    const name = /^\\([a-zA-Z]+\*?)/.exec(latex.slice(i))?.[1];
    if (!name) {
      i++;
      continue;
    }
    if (name === "begin") depth++;
    else if (name === "end") depth--;
    else if (depth === 0 && (name === "tag" || name === "tag*")) {
      let j = i + 1 + name.length;
      while (latex[j] === " ") j++;
      if (latex[j] !== "{") return null;
      let level = 0;
      let k = j;
      for (; k < latex.length; k++) {
        if (latex[k] === "\\") k++;
        else if (latex[k] === "{") level++;
        else if (latex[k] === "}" && --level === 0) break;
      }
      const label = latex.slice(j + 1, k).trim().replace(/^\$(.*)\$$/, "$1");
      return { label, star: name === "tag*" };
    }
    i += name.length;
  }
  return null;
}

/**
 * The sections' numbers in the document's format, by the heading's position: "2.1." or "2.1" (none: no numbers).
 * Levels 1-3 are numbered, a section resets the deeper levels; an unnumbered section (\section*) takes no number and
 * leaves the counters as they are, as in LaTeX. Counted here, not by CSS counters: Chromium no longer carries a
 * heading's counter-reset over to the headings after it.
 */
export function headingNumbers(doc: Node): Map<number, string> {
  const style = documentSettings(doc).headingNumbering;
  const numbers = new Map<number, string>();
  if (style === "none") return numbers;
  const counters = [0, 0, 0];
  doc.descendants((node, pos) => {
    if (node.type.name !== "heading") return !node.isTextblock;
    const level = node.attrs.level as number;
    if (level > 3 || node.attrs.unnumbered) return false;
    counters[level - 1]++;
    for (let k = level; k < 3; k++) counters[k] = 0;
    numbers.set(pos, counters.slice(0, level).join(".") + (style === "1." ? "." : ""));
    return false;
  });
  return numbers;
}

/**
 * The numbered display formulas' labels in document order, by formula id: "3", or "2.3" / "2-3" (the 3rd in the
 * 2nd level-1 section; before the first section "0.1", as the Word add-in numbers them). An unnumbered section
 * (\section*) does not step the section, as in LaTeX.
 */
export function equationLabels(doc: Node): Map<string, string> {
  const style = documentSettings(doc).equationNumbering;
  const labels = new Map<string, string>();
  let section = 0;
  let n = 0;
  doc.descendants((node) => {
    if (node.type.name === "heading" && node.attrs.level === 1 && !node.attrs.unnumbered && style !== "1") {
      section += 1;
      n = 0;
    } else if (node.type.name === "math_display" && ownEquationTag(node.attrs.latex)) {
      // a formula's own \tag: no number; references show the tag
      if (node.attrs.id) labels.set(node.attrs.id, ownEquationTag(node.attrs.latex)!.label);
    } else if (node.type.name === "math_display" && node.attrs.numbered) {
      n += 1;
      const label = style === "1" ? String(n) : `${section}${style === "1.1" ? "." : "-"}${n}`;
      if (node.attrs.id) labels.set(node.attrs.id, label);
    }
    return node.type.name !== "math_display";
  });
  return labels;
}

/** A label in the document's brackets: (2.3) or [2.3]. */
export function bracketLabel(doc: Node, label: string): string {
  return documentSettings(doc).equationBrackets === "[]" ? `[${label}]` : `(${label})`;
}

/** The settings the document chose itself (the others are the defaults). */
export function chosenSettings(doc: Node): Set<keyof DocumentSettings> {
  const chosen = new Set<keyof DocumentSettings>();
  for (const name of Object.keys(DOCUMENT_SETTING_ATTRS) as Array<keyof DocumentSettings>) {
    if (doc.attrs[name] !== null && doc.attrs[name] !== undefined) chosen.add(name);
  }
  return chosen;
}

/** Change some of the document's settings (one undo step). */
export function setDocumentSettings(patch: Partial<DocumentSettings>): Command {
  return (state, dispatch) => {
    const current = state.doc.attrs as Record<string, unknown>;
    const entries = Object.entries(patch).filter(([name, value]) => name in DOCUMENT_SETTING_ATTRS && current[name] !== value);
    if (!entries.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const [name, value] of entries) tr.setDocAttribute(name, value ?? null);
      dispatch(tr);
    }
    return true;
  };
}

/** The page's text area in millimetres. */
export function textArea(settings: DocumentSettings): { width: number; height: number } {
  const paper = PAPER_SIZES[settings.paper];
  return {
    width: Math.max(20, paper.width - settings.marginLeft - settings.marginRight),
    height: Math.max(20, paper.height - settings.marginTop - settings.marginBottom),
  };
}

/** The class the LaTeX export writes (latex.ts): ctexart for a document with Chinese text, article otherwise. */
export type LatexClass = "article" | "ctexart";

export interface BodyLayout {
  documentClass: LatexClass;
  /** the body size, pt */
  fontSize: number;
  /** \baselineskip as a multiple of the size */
  lineHeight: number;
  /** the line spacing's factor on the class's \baselineskip (\linespread) */
  spread: number;
  /** \parindent, em of the body size */
  parindent: number;
}

/** article's \baselineskip and \parindent at its class options' sizes (pt). */
const ARTICLE_SIZES: Record<number, { baseline: number; parindent: number }> = {
  10: { baseline: 12, parindent: 15 },
  11: { baseline: 13.6, parindent: 17 },
  12: { baseline: 14.5, parindent: 18 },
};

/**
 * The body as the exported class sets it, which the page view shows: article's 10 pt, \baselineskip 1.2 x the size
 * and \parindent 1.5em; ctexart's 五号 (10.5 pt), \linespread 1.3 on the 1.2 (1.56 x the size) and 2em (two
 * characters). A size or spacing the document chose is the preamble's (settingsPreamble).
 */
export function bodyLayout(doc: Node): BodyLayout {
  const settings = documentSettings(doc);
  const chosen = chosenSettings(doc);
  const cjk = documentHasCjk(doc);
  const fontSize = chosen.has("fontSize") ? settings.fontSize : cjk ? 10.5 : 10;
  const article = ARTICLE_SIZES[fontSize];
  const spread = (cjk ? SINGLE_LINE_HEIGHT : 1) * settings.lineSpacing;
  return {
    documentClass: cjk ? "ctexart" : "article",
    fontSize,
    lineHeight: (cjk || !article ? 1.2 : article.baseline / fontSize) * spread,
    spread,
    parindent: cjk ? 2 : article ? article.parindent / fontSize : 1.5,
  };
}

/** A body size's Chinese name (五号 for 10.5 pt), if it has one. */
export function chineseSizeName(pt: number): string | null {
  return CHINESE_FONT_SIZES.find((size) => Math.abs(size.pt - pt) < 0.01)?.name ?? null;
}

/** CSS font-family for the body: the Latin font first (Latin text), the CJK font next (Chinese text), fallbacks. */
export function bodyFontFamily(settings: DocumentSettings): string {
  const quote = (name: string) => `"${name.replace(/"/g, "")}"`;
  return [quote(settings.latinFont), quote(settings.cjkFont), '"Times New Roman"', '"SimSun"', '"Songti SC"', "serif"].join(", ");
}

const formatMm = (value: number) => `${Math.round(value * 10) / 10}mm`;

/**
 * The preamble for the settings (DESIGN.md §8.4): the paper and margins always (they make the page the editor
 * shows); the body size, spacing and fonts when the document chose them (a font needs XeLaTeX: fontspec / ctex read
 * the system's font names, and the file says so in its first line).
 */
export function settingsPreamble(
  settings: DocumentSettings,
  chosen: ReadonlySet<keyof DocumentSettings>,
  cjk: boolean,
): { classOptions: string[]; lines: string[]; xelatex: boolean } {
  const classOptions: string[] = [];
  const lines: string[] = [];
  if (chosen.has("fontSize")) {
    const named = CHINESE_FONT_SIZES.find((size) => Math.abs(size.pt - settings.fontSize) < 0.01);
    if (cjk && named) classOptions.push(`zihao=${named.zihao}`);
    else if ([10, 11, 12].includes(settings.fontSize)) classOptions.push(`${settings.fontSize}pt`);
    else lines.push(`\\usepackage[fontsize=${settings.fontSize}pt]{scrextend}`);
  }
  if (chosen.has("lineSpacing")) {
    if (cjk) classOptions.push(`linespread=${Math.round(SINGLE_LINE_HEIGHT * settings.lineSpacing * 100) / 100}`);
    // the spacing multiplies the class's single spacing (article's \baselineskip, ctex's \linespread 1.3 on it), as the
    // page view shows it (bodyLayout)
    else if (settings.lineSpacing !== 1) lines.push(`\\linespread{${Math.round(settings.lineSpacing * 100) / 100}}`);
  }
  const paper = PAPER_SIZES[settings.paper];
  lines.push(
    `\\usepackage[paperwidth=${formatMm(paper.width)},paperheight=${formatMm(paper.height)},top=${formatMm(settings.marginTop)},bottom=${formatMm(settings.marginBottom)},left=${formatMm(settings.marginLeft)},right=${formatMm(settings.marginRight)}]{geometry}`,
  );
  // equation numbers within sections, their separator and brackets (\eqref shows the same)
  if (settings.equationNumbering !== "1") lines.push("\\numberwithin{equation}{section}");
  if (settings.equationNumbering === "1-1") lines.push("\\renewcommand{\\theequation}{\\thesection-\\arabic{equation}}");
  if (settings.equationBrackets === "[]") lines.push("\\makeatletter\\renewcommand{\\tagform@}[1]{\\maketag@@@{[\\ignorespaces#1\\unskip\\@@italiccorr]}}\\makeatother");
  if (settings.headingNumbering === "none") lines.push("\\setcounter{secnumdepth}{0}");
  // "1." / "1.1.": a period after the section's number in its heading (references keep "1.1")
  if (settings.headingNumbering === "1.") {
    lines.push(
      cjk
        ? "\\ctexset{section/aftername=.\\quad,subsection/aftername=.\\quad,subsubsection/aftername=.\\quad}"
        : "\\makeatletter\\renewcommand\\@seccntformat[1]{\\csname the#1\\endcsname.\\quad}\\makeatother",
    );
  }
  const latinFont = chosen.has("latinFont");
  const cjkFont = cjk && chosen.has("cjkFont");
  if (latinFont && !cjk) lines.push("\\usepackage{fontspec}");
  if (latinFont) lines.push(`\\setmainfont{${settings.latinFont}}`);
  if (cjkFont) lines.push(`\\setCJKmainfont{${settings.cjkFont}}`);
  return { classOptions, lines, xelatex: cjk || latinFont };
}
