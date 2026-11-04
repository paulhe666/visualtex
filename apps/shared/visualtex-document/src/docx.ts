// Document -> Word (.docx) (DESIGN.md §12; research ~/claude/document-export/research/docx.md). Every node has one
// mapping, as in latex.ts. Cross-references stay live: equation numbers are Word's own numbered equations (SEQ in
// an eqArr "#" number) with a bookmark, references are REF \h fields to it, citations and links are hyperlinks to
// bookmarks, footnotes are Word footnotes. Formulas come from the app (OMML through the shared Word writer, or
// VisualTeX formula pictures, or other objects) by key (docxFormulaRequests).
import type { Mark, Node } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";
import { documentSettings, equationLabels, PAPER_SIZES } from "./settings.ts";
import { bibliographyEntries, citeLabel } from "./notes.ts";
import { splitInlineMath } from "./mathText.ts";
import { FONT_SIZES } from "./schema.ts";
import { IMAGE_GAP, imageAspect, imageEntry, isImage, isImageLine, sha256Hex } from "./images.ts";
import { isPageBreak } from "./pageCommands.ts";
import type { CodeToken } from "./codeBlocks.ts";

/** A formula the app has to provide, by key. */
export interface DocxFormulaRequest {
  key: string;
  latex: string;
  display: boolean;
  /** a numbered display formula: what its Word number needs */
  number?: DocxEquationNumber;
}

export interface DocxEquationNumber {
  /** the bookmark around the number (VTEqNum_ + 32 hex): what references point to */
  bookmark: string;
  /** the number in the sequence ("3" of "2.3") */
  sequence: string;
  /** 0: one count through the document; 1: restarted at each level-1 heading, with its number first */
  chapterLevel: 0 | 1;
  /** "." or "-" between the chapter and the count */
  separator: string;
  /** the chapter number as text, when Word's heading numbers cannot give it (no numbered headings, before the first) */
  literalPrefix: string | null;
  /** the whole label, as the editor shows it ("2.3") */
  label: string;
  brackets: "()" | "[]";
  /** the formula id a picture of it carries (equationFormulaId) */
  formulaId: string;
}

export type DocxFormula =
  /** one m:oMath (a numbered one already carries its "#" number) */
  | { kind: "omml"; xml: string }
  /** a VisualTeX formula: a picture with the formula's metadata in its alt text */
  | {
      kind: "picture";
      png: Uint8Array;
      svg?: Uint8Array;
      widthPt: number;
      heightPt: number;
      /** below the baseline */
      depthPt: number;
      descr: string;
      title?: string;
    }
  /** an embedded OLE object (MathType on Windows): the run's whole XML and its parts */
  | { kind: "object"; runXml: (rel: (target: string, type: string) => string) => string; parts: Array<{ path: string; bytes: Uint8Array; contentType: string; extension: string }> };

export interface DocxExportOptions {
  /** A formula's LaTeX as it leaves the app (plain amsmath). */
  formulaLatex?: (latex: string, display: boolean) => string;
  title?: string;
  /** "参考文献" / "References" above the reference list */
  bibliographyTitle?: string;
  /** the math font named in the document's settings (Word's default math font) */
  mathFont?: string;
  /** the app's code highlighter (the colours the code blocks show) */
  codeTokens?: (code: string, language: string) => CodeToken[][] | null;
}

/** A section's page numbers (\pagenumbering, \pagestyle, \setcounter{page}). */
interface PageSection {
  fmt: string;
  start: number | null;
  shown: boolean;
  /** \thispagestyle{empty} on the section's first page */
  firstEmpty: boolean;
  /** how the section starts: on a new page (after \newpage) or where the command is */
  type: "nextPage" | "continuous";
}

/** What the page view starts below a wrapped image (document.css: everything but a plain text paragraph). */
function clearsWrap(node: Node): boolean {
  if (node.type.name !== "paragraph") return node.type.name !== "page_command";
  return isImageLine(node) || (!!node.firstChild && isImage(node.firstChild) && !!node.firstChild.attrs.wrap);
}

/** A text-wrapping break that clears both sides at the end of the last paragraph written (else a paragraph of it). */
function clearAfterLastParagraph(xml: string): string {
  const BREAK = '<w:r><w:br w:type="textWrapping" w:clear="all"/></w:r>';
  const at = xml.lastIndexOf("</w:p>");
  if (at < 0 || at + 6 !== xml.length) return xml + `<w:p>${BREAK}</w:p>`;
  return xml.slice(0, at) + BREAK + xml.slice(at);
}

/** a code block's text from the margins: the box's border 10 pt outside it */
const CODE_INSET = 220;
const PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

const WORD_PAGE_FORMATS: Record<string, string> = { arabic: "decimal", roman: "lowerRoman", Roman: "upperRoman", alph: "lowerLetter", Alph: "upperLetter" };

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS =
  `xmlns:w="${W_NS}" xmlns:r="${R_NS}" ` +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const MM_TWIPS = 1440 / 25.4;
const EMU_PER_TWIP = 635;
const EMU_PER_PT = 12700;

export const esc = (text: string) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
/** XML 1.0 forbids most control characters */
const clean = (text: string) => text.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "");

/** A hidden bookmark name (≤ 40 characters, letters, digits and _). */
function bookmarkName(prefix: string, id: string): string {
  return (prefix + id.replace(/[^A-Za-z0-9_]/g, "_")).slice(0, 40);
}

/** The bookmark of a numbered display formula's number (the shared Word writer's form: VTEqNum_<32 hex>). */
export function equationBookmark(id: string): string {
  return `VTEqNum_${sha256Hex(new TextEncoder().encode(`visualtex-equation:${id}`)).slice(0, 32)}`;
}

/** The formula id (a UUID) a numbered display formula's picture carries in its VisualTeX data. */
export function equationFormulaId(id: string): string {
  const hex = equationBookmark(id).slice("VTEqNum_".length);
  // a version-4 UUID's form, which the add-ins check
  const variant = ((parseInt(hex[16], 16) & 3) | 8).toString(16);
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-${variant}${hex.slice(17, 20)}-${hex.slice(20, 32)}`;
}

/** The formulas a document needs (body and footnotes), with the numbers of the numbered ones. */
export function docxFormulaRequests(doc: Node): DocxFormulaRequest[] {
  const settings = documentSettings(doc);
  const labels = equationLabels(doc);
  const requests: DocxFormulaRequest[] = [];
  let headingSeen = false;
  doc.descendants((node) => {
    if (node.type.name === "heading" && node.attrs.level === 1) headingSeen = true;
    if (node.type.name === "math_inline" || node.type.name === "math_display") {
      const display = node.type.name === "math_display";
      const request: DocxFormulaRequest = { key: formulaKey(node), latex: node.attrs.latex, display };
      const label = display && node.attrs.numbered && node.attrs.id ? labels.get(node.attrs.id) : undefined;
      if (label) {
        const chapter = settings.equationNumbering !== "1";
        const separator = settings.equationNumbering === "1-1" ? "-" : ".";
        const sequence = chapter ? label.slice(label.lastIndexOf(separator) + 1) : label;
        const prefix = chapter ? label.slice(0, label.lastIndexOf(separator)) : "";
        request.number = {
          bookmark: equationBookmark(node.attrs.id),
          sequence,
          chapterLevel: chapter ? 1 : 0,
          separator,
          // Word's STYLEREF reads the heading's own number; without numbered headings (or before the first) the
          // chapter number is written as it is
          literalPrefix: chapter && (settings.headingNumbering === "none" || !headingSeen) ? prefix : null,
          label,
          brackets: settings.equationBrackets,
          formulaId: equationFormulaId(node.attrs.id),
        };
      }
      requests.push(request);
      return false;
    }
    if (node.type.name === "footnote") {
      splitInlineMath(node.attrs.text ?? "").forEach((part, index) => {
        if (part.kind === "math") requests.push({ key: `fn:${node.attrs.id ?? "?"}:${index}`, latex: part.latex, display: false });
      });
    }
    return true;
  });
  return requests;
}

function formulaKey(node: Node): string {
  return `f:${node.attrs.id ?? sha256Hex(new TextEncoder().encode(node.attrs.latex)).slice(0, 16)}`;
}

interface Media {
  path: string;
  bytes: Uint8Array;
}

class DocxWriter {
  private rels: string[] = [];
  private footnoteRels: string[] = [];
  private relIds = new Map<string, string>();
  readonly media: Media[] = [];
  private mediaByKey = new Map<string, string>();
  readonly footnotes: string[] = [];
  readonly extraParts: Array<{ path: string; bytes: Uint8Array; contentType: string; extension: string }> = [];
  private bookmarkId = 100;
  private docPrId = 0;
  private lists: Array<{ numId: number; ordered: boolean; start: number }> = [];
  private nextNumId = 10;
  private inFootnote = false;
  readonly settings;
  readonly bodySizePt: number;
  readonly textWidthTwips: number;
  private labels: Map<string, string>;
  private bibNumbers: Map<string, number>;
  /** paragraphs an internal link (#id) points to: they get a bookmark */
  private linkTargets = new Set<string>();
  /** the page numbering of the section being written (page commands start new sections) */
  private section: PageSection = { fmt: "arabic", start: null, shown: true, firstEmpty: false, type: "continuous" };
  private sectionHasContent = false;
  /** a wrapped image's text is still beside it (top level) */
  private wrapOpen = false;
  /** a page break was written in the section: its first page is behind */
  private sectionPaged = false;
  /** a \newpage / \clearpage not yet written (it becomes the next section's start or a page break) */
  private pageBreak = false;
  readonly footers: string[] = [];
  private footerRels = new Map<string, string>();
  /** numbered display formulas drawn as pictures (or objects) */
  private pictureNumbers = new Set<string>();


  constructor(
    readonly doc: Node,
    private formulas: Map<string, DocxFormula>,
    readonly options: DocxExportOptions,
  ) {
    this.settings = documentSettings(doc);
    this.bodySizePt = this.settings.fontSize;
    const paper = PAPER_SIZES[this.settings.paper];
    this.textWidthTwips = Math.round((paper.width - this.settings.marginLeft - this.settings.marginRight) * MM_TWIPS);
    this.labels = equationLabels(doc);
    doc.descendants((node) => {
      if (node.type.name === "math_display" && node.attrs.numbered && node.attrs.id && this.formulas.get(formulaKey(node))?.kind !== "omml")
        this.pictureNumbers.add(node.attrs.id);
      return node.type.name !== "math_display";
    });
    this.bibNumbers = new Map(bibliographyEntries(doc).map((entry) => [entry.id, entry.number]));
    doc.descendants((node) => {
      for (const mark of node.marks) {
        const href = mark.type.name === "link" ? String(mark.attrs.href ?? "") : "";
        if (href.startsWith("#")) this.linkTargets.add(href.slice(1));
      }
      return true;
    });
  }

  // ------------------------------------------------------------ relationships --

  rel(target: string, type: string, external = false): string {
    const key = `${type} ${target}`;
    const list = this.inFootnote ? this.footnoteRels : this.rels;
    const known = this.inFootnote ? undefined : this.relIds.get(key);
    if (known) return known;
    const id = `rIdV${list.length + 1}${this.inFootnote ? "f" : ""}`;
    list.push(`<Relationship Id="${id}" Type="${R_NS}/${type}" Target="${esc(target)}"${external ? ' TargetMode="External"' : ""}/>`);
    if (!this.inFootnote) this.relIds.set(key, id);
    return id;
  }

  relsXml(footnotes = false): string {
    const fixed = footnotes
      ? ""
      : `<Relationship Id="rIdStyles" Type="${R_NS}/styles" Target="styles.xml"/>` +
        `<Relationship Id="rIdNumbering" Type="${R_NS}/numbering" Target="numbering.xml"/>` +
        `<Relationship Id="rIdSettings" Type="${R_NS}/settings" Target="settings.xml"/>` +
        `<Relationship Id="rIdFootnotes" Type="${R_NS}/footnotes" Target="footnotes.xml"/>`;
    return `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${fixed}${(footnotes ? this.footnoteRels : this.rels).join("")}</Relationships>`;
  }

  private mediaRel(key: string, bytes: Uint8Array, extension: string): string {
    let path = this.mediaByKey.get(key);
    if (!path) {
      path = `media/image${this.media.length + 1}.${extension}`;
      this.media.push({ path, bytes });
      this.mediaByKey.set(key, path);
    }
    return this.rel(path, "image");
  }

  // ------------------------------------------------------------------- pieces --

  bookmark(name: string, content: string): string {
    const id = ++this.bookmarkId;
    return `<w:bookmarkStart w:id="${id}" w:name="${esc(name)}"/>${content}<w:bookmarkEnd w:id="${id}"/>`;
  }

  run(text: string, rPr = ""): string {
    // tabs and line breaks inside text become Word's own elements
    const parts = clean(text).split(/(\t|\n)/);
    let out = "";
    for (const part of parts) {
      if (!part) continue;
      if (part === "\t") out += `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:tab/></w:r>`;
      else if (part === "\n") out += `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:br/></w:r>`;
      else out += `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${esc(part)}</w:t></w:r>`;
    }
    return out;
  }

  field(instruction: string, result: string, rPr = ""): string {
    const p = rPr ? `<w:rPr>${rPr}</w:rPr>` : "";
    return (
      `<w:r>${p}<w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r>${p}<w:instrText xml:space="preserve"> ${esc(instruction)} </w:instrText></w:r>` +
      `<w:r>${p}<w:fldChar w:fldCharType="separate"/></w:r>` +
      this.run(result, rPr) +
      `<w:r>${p}<w:fldChar w:fldCharType="end"/></w:r>`
    );
  }

  /** Character formatting of marks, in Word's w:rPr child order. */
  rPr(marks: readonly Mark[]): string {
    const has = (name: string) => marks.find((m) => m.type.name === name);
    const parts: Array<[number, string]> = [];
    if (has("code")) parts.push([0, '<w:rStyle w:val="VerbatimChar"/>']);
    else if (has("link")) parts.push([0, '<w:rStyle w:val="Hyperlink"/>']);
    if (has("strong")) parts.push([2, "<w:b/><w:bCs/>"]);
    if (has("em")) parts.push([4, "<w:i/><w:iCs/>"]);
    if (has("strike")) parts.push([8, "<w:strike/>"]);
    const color = has("textColor");
    if (color) {
      const hex = cssHex(color.attrs.color);
      if (hex) parts.push([12, `<w:color w:val="${hex}"/>`]);
    }
    const size = has("fontSize");
    if (size) {
      const ratio = FONT_SIZES.find((s) => s.name === size.attrs.size)?.ratio ?? 1;
      const half = Math.round(this.bodySizePt * ratio * 2);
      parts.push([17, `<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>`]);
    }
    const highlight = has("highlight");
    if (highlight) {
      const named = HIGHLIGHTS[cssHex(highlight.attrs.color) ?? ""];
      parts.push(named ? [19, `<w:highlight w:val="${named}"/>`] : [22, `<w:shd w:val="clear" w:color="auto" w:fill="${cssHex(highlight.attrs.color) ?? "FFFF00"}"/>`]);
    }
    if (has("underline")) parts.push([20, '<w:u w:val="single"/>']);
    return parts.sort((a, b) => a[0] - b[0]).map((p) => p[1]).join("");
  }

  // ------------------------------------------------------------------ formulas --

  private formula(key: string): DocxFormula | undefined {
    return this.formulas.get(key);
  }

  /** An inline formula among runs: OMML between the runs, a picture or an object in a run. */
  inlineFormula(node: Node): string {
    const script = /^\s*([\^_])\s*\{\s*\\text(?:rm|normal)?\s*\{([^{}$\\]*)\}\s*\}\s*$/.exec(String(node.attrs.latex ?? ""));
    if (script) return this.run(script[2], `${this.rPr(node.marks)}<w:vertAlign w:val="${script[1] === "^" ? "superscript" : "subscript"}"/>`);
    const formula = this.formula(formulaKey(node));
    if (!formula) return this.run(`$${node.attrs.latex}$`);
    return this.formulaRun(formula, false);
  }

  formulaRun(formula: DocxFormula, display: boolean): string {
    if (formula.kind === "omml") {
      // the writer gives every number's bookmark the same id; each needs its own
      const ids = new Map<string, number>();
      const xml = canonicalOmml(formula.xml).replace(/(<w:bookmark(?:Start|End)\b[^>]*?\bw:id=")(\d+)"/g, (_all, head: string, id: string) => {
        if (!ids.has(id)) ids.set(id, ++this.bookmarkId);
        return `${head}${ids.get(id)}"`;
      });
      return display ? `<m:oMathPara>${xml}</m:oMathPara>` : xml;
    }
    if (formula.kind === "object") return formula.runXml((target, type) => this.rel(target, type));
    return this.picture(formula);
  }

  /** A formula picture: the PNG (with its SVG when there is one), its baseline on the text's. */
  picture(formula: Extract<DocxFormula, { kind: "picture" }>): string {
    const key = sha256Hex(formula.png).slice(0, 24);
    const png = this.mediaRel(`png:${key}`, formula.png, "png");
    const svg = formula.svg ? this.mediaRel(`svg:${key}`, formula.svg, "svg") : null;
    const cx = Math.round(formula.widthPt * EMU_PER_PT);
    const cy = Math.round(formula.heightPt * EMU_PER_PT);
    const position = Math.round(-formula.depthPt * 2);
    const drawing = this.inlineDrawing(png, svg, cx, cy, formula.descr, formula.title);
    return `<w:r>${position ? `<w:rPr><w:position w:val="${position}"/></w:rPr>` : ""}${drawing}</w:r>`;
  }

  private graphic(blip: string, svg: string | null, cx: number, cy: number, name: string, descr: string): string {
    const ext = svg
      ? `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${svg}"/></a:ext></a:extLst>`
      : "";
    return (
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>` +
      `<pic:nvPicPr><pic:cNvPr id="0" name="${esc(name)}" descr="${esc(descr)}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${blip}">${ext}</a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic>`
    );
  }

  private inlineDrawing(blip: string, svg: string | null, cx: number, cy: number, descr: string, title?: string): string {
    const id = ++this.docPrId;
    const name = `Picture ${id}`;
    return (
      `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${id}" name="${name}" descr="${esc(descr)}"${title ? ` title="${esc(title)}"` : ""}/>` +
      `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      this.graphic(blip, svg, cx, cy, name, descr) +
      `</wp:inline></w:drawing>`
    );
  }

  /** A display formula's paragraph; a numbered picture or object has its number at the right margin (tabs). */
  displayFormula(node: Node): string {
    const formula = this.formula(formulaKey(node));
    const label = node.attrs.numbered && node.attrs.id ? this.labels.get(node.attrs.id) : undefined;
    if (!formula) return this.paragraph(this.run(`$$${node.attrs.latex}$$`), '<w:jc w:val="center"/>');
    if (formula.kind === "omml") return `<w:p><w:pPr>${this.displaySpacing()}<w:jc w:val="center"/></w:pPr>${this.formulaRun(formula, true)}</w:p>`;
    if (!label) return this.paragraph(this.formulaRun(formula, true), this.displaySpacing() + '<w:jc w:val="center"/>');
    // a picture (or object) and its number: centred by a centre tab, the number at a right tab
    const tabs = `<w:tabs><w:tab w:val="center" w:pos="${Math.round(this.textWidthTwips / 2)}"/><w:tab w:val="right" w:pos="${this.textWidthTwips}"/></w:tabs>`;
    const number = this.bodyEquationNumber(node.attrs.id, label);
    return this.paragraph(`<w:r><w:tab/></w:r>${this.formulaRun(formula, true)}<w:r><w:tab/></w:r>${number}`, tabs + this.displaySpacing() + '<w:jc w:val="left"/>');
  }

  /** The bookmark references point to: Word's own number's (VTEqNum_) for a Word equation; for a picture the
   * VisualTeX Word add-ins' (VT_N_ + the picture's formula id, equationFormulaId): they take such a picture as their
   * own when the document opens and renumber it under that name, so references keep their target. */
  private numberBookmark(id: string): string {
    return this.pictureNumbers.has(id) ? `VT_N_${equationFormulaId(id).replace(/-/g, "")}` : equationBookmark(id);
  }

  /** The page view's room above and below a display formula: 0.3 em of the formula's size (1.04 × the text's). */
  private displaySpacing(): string {
    const twips = Math.round(0.3 * 1.04 * this.bodySizePt * 20);
    return `<w:spacing w:before="${twips}" w:after="${twips}"/>`;
  }

  /** "(2.3)" in body text: SEQ (and the chapter) fields with the bookmark references point to. */
  private bodyEquationNumber(id: string, label: string): string {
    const style = this.settings.equationNumbering;
    const [open, close] = this.settings.equationBrackets === "[]" ? ["[", "]"] : ["(", ")"];
    let inner = "";
    if (style === "1") inner = this.field("SEQ Equation \\* ARABIC", label);
    else {
      const separator = style === "1-1" ? "-" : ".";
      const prefix = label.slice(0, label.lastIndexOf(separator));
      const sequence = label.slice(label.lastIndexOf(separator) + 1);
      const chapter = this.settings.headingNumbering === "none" || prefix === "0" ? this.run(prefix) : this.field("STYLEREF 1 \\w", prefix);
      inner = chapter + this.run(separator) + this.field("SEQ Equation \\s 1 \\* ARABIC", sequence);
    }
    return this.run(open) + this.bookmark(this.numberBookmark(id), inner) + this.run(close);
  }

  // ------------------------------------------------------------------- inline --

  inline(node: Node): string {
    let out = "";
    // links: consecutive runs with the same link mark in one w:hyperlink
    let link: Mark | null = null;
    let linkContent = "";
    const flush = () => {
      if (!link) return;
      const href = String(link.attrs.href ?? "");
      if (href.startsWith("#")) out += `<w:hyperlink w:anchor="${esc(bookmarkName("_VtxPara_", href.slice(1)))}" w:history="1">${linkContent}</w:hyperlink>`;
      else out += `<w:hyperlink r:id="${this.rel(href, "hyperlink", true)}" w:history="1"${link.attrs.title ? ` w:tooltip="${esc(link.attrs.title)}"` : ""}>${linkContent}</w:hyperlink>`;
      link = null;
      linkContent = "";
    };
    node.forEach((child, _offset, index) => {
      const marks = child.marks ?? [];
      const childLink = marks.find((m) => m.type.name === "link") ?? null;
      if (link && (!childLink || !childLink.eq(link))) flush();
      let piece = "";
      if (child.isText) piece = this.run(child.text ?? "", this.rPr(marks));
      else if (child.type.name === "math_inline") piece = this.inlineFormula(child);
      else if (child.type.name === "hard_break") piece = "<w:r><w:br/></w:r>";
      else if (isImage(child)) piece = child.attrs.wrap && index === 0 ? this.wrappedImage(child) : this.inlineImage(child);
      else if (child.type.name === "eqref") piece = this.equationReference(child);
      else if (child.type.name === "footnote") piece = this.footnoteReference(child);
      else if (child.type.name === "cite") piece = this.citation(child);
      if (childLink) {
        link = childLink;
        linkContent += piece;
      } else out += piece;
    });
    flush();
    // Word (and WPS) take an equation alone in its paragraph as a display one, centred; a zero-width space keeps it in the line
    let alone = node.childCount > 0;
    node.forEach((child) => {
      if (!(child.type.name === "math_inline" || (child.isText && !child.text?.trim()))) alone = false;
    });
    if (alone && out.includes("<m:oMath")) out += this.run("\u200b");
    return out;
  }

  /** (n): a REF \h field to the number's bookmark (Word updates it and jumps to the equation). */
  private equationReference(node: Node): string {
    const target = node.attrs.target as string | null;
    const label = target ? this.labels.get(target) : undefined;
    if (!target || !label) return this.run("(??)");
    const [open, close] = this.settings.equationBrackets === "[]" ? ["[", "]"] : ["(", ")"];
    return this.run(open) + this.field(`REF ${this.numberBookmark(target)} \\h`, label) + this.run(close);
  }

  private citation(node: Node): string {
    const targets = node.attrs.targets as string[];
    const label = citeLabel(targets.map((id) => this.bibNumbers.get(id) ?? null));
    const first = targets.find((id) => this.bibNumbers.has(id));
    if (!first) return this.run(label);
    return `<w:hyperlink w:anchor="${bookmarkName("_VtxBib_", first)}" w:history="1">${this.run(label)}</w:hyperlink>`;
  }

  private footnoteReference(node: Node): string {
    const id = this.footnotes.length + 1;
    this.inFootnote = true;
    let body = "";
    splitInlineMath(node.attrs.text ?? "").forEach((part, index) => {
      if (part.kind === "text") body += this.run(part.text);
      else {
        const formula = this.formulas.get(`fn:${node.attrs.id ?? "?"}:${index}`);
        body += formula ? this.formulaRun(formula, false) : this.run(`$${part.latex}$`);
      }
    });
    this.inFootnote = false;
    this.footnotes.push(
      `<w:footnote w:id="${id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>${this.run(" ")}${body}</w:p></w:footnote>`,
    );
    return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
  }

  // -------------------------------------------------------------------- images --

  private imageBlip(node: Node): { rel: string; aspect: number } | null {
    const entry = imageEntry(node.attrs.src);
    if (!entry) return null;
    return { rel: this.mediaRel(entry.key, entry.bytes, entry.mime === "image/png" ? "png" : "jpeg"), aspect: imageAspect(entry.key) };
  }

  private imageSize(node: Node, aspect: number, lineTwips = this.textWidthTwips): { cx: number; cy: number } {
    const cx = Math.round(node.attrs.width * lineTwips * EMU_PER_TWIP);
    return { cx, cy: Math.round(cx / aspect) };
  }

  inlineImage(node: Node): string {
    const blip = this.imageBlip(node);
    if (!blip) return "";
    const { cx, cy } = this.imageSize(node, blip.aspect);
    return `<w:r>${this.inlineDrawing(blip.rel, null, cx, cy, node.attrs.alt ?? "")}</w:r>`;
  }

  /** wrapfig's figure: anchored at the paragraph's top, at the text area's left or right edge, text around it. */
  wrappedImage(node: Node): string {
    const blip = this.imageBlip(node);
    if (!blip) return "";
    const { cx, cy } = this.imageSize(node, blip.aspect);
    const id = ++this.docPrId;
    const gap = 10 * EMU_PER_PT;
    const below = 12 * EMU_PER_PT;
    const side = node.attrs.wrap === "left" ? "left" : "right";
    const [distL, distR] = side === "right" ? [gap, 0] : [0, gap];
    // the editor and wrapfig put the top about 0.55 em under the first line's baseline: near the line's bottom
    const offset = Math.round(this.bodySizePt * 0.55 * EMU_PER_PT);
    return (
      `<w:r><w:drawing><wp:anchor distT="0" distB="${below}" distL="${distL}" distR="${distR}" simplePos="0" relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0">` +
      `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>${side}</wp:align></wp:positionH>` +
      `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${offset}</wp:posOffset></wp:positionV>` +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>` +
      `<wp:docPr id="${id}" name="Picture ${id}" descr="${esc(node.attrs.alt ?? "")}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      this.graphic(blip.rel, null, cx, cy, `Picture ${id}`, node.attrs.alt ?? "") +
      `</wp:anchor></w:drawing></w:r>`
    );
  }

  /** An image line: its images on one aligned line, 0.02 \linewidth apart (character spacing on a space). */
  imageLine(node: Node): string {
    const gapTwips = Math.round(IMAGE_GAP * this.textWidthTwips);
    const runs: string[] = [];
    node.forEach((child) => {
      if (isImage(child)) runs.push(this.inlineImage(child));
    });
    const gap = `<w:r><w:rPr><w:spacing w:val="${gapTwips}"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>`;
    const align = node.attrs.align === "center" ? "center" : node.attrs.align === "right" ? "right" : "left";
    return this.paragraph(runs.join(gap), `<w:spacing w:line="240" w:lineRule="auto"/><w:jc w:val="${align}"/>`);
  }

  // -------------------------------------------------------------------- blocks --

  paragraph(content: string, pPr = ""): string {
    return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${content}</w:p>`;
  }

  private paragraphPr(node: Node, extra: { style?: string; indentTwips?: number; numPr?: string } = {}): string {
    const parts: string[] = [];
    if (extra.style) parts.push(`<w:pStyle w:val="${extra.style}"/>`);
    if (extra.numPr) parts.push(extra.numPr);
    const indent = (extra.indentTwips ?? 0) + (node.attrs.indent ? Math.round(node.attrs.indent * 2 * this.bodySizePt * 20) : 0);
    if (indent) parts.push(`<w:ind w:left="${indent}"/>`);
    const align = node.attrs.align as string | null;
    if (align) parts.push(`<w:jc w:val="${align === "justify" ? "both" : align}"/>`);
    return parts.join("");
  }

  blocks(node: Node, indent = 0, quote = false): string {
    let out = "";
    let previousTable = false;
    node.forEach((child) => {
      if (child.type.name === "page_command") {
        if (isPageBreak(child)) {
          if (this.pageBreak) out += this.writePageBreak();
          this.pageBreak = true;
        } else {
          // the numbering changes: a new section (starting on a new page when a \newpage came just before), unless
          // nothing is in the current one yet
          if (this.sectionHasContent) out += this.sectionBreak(this.pageBreak ? "nextPage" : "continuous");
          this.applyPageCommand(child.attrs.command, child.attrs.value);
        }
        return;
      }
      if (this.pageBreak) out += this.writePageBreak();
      this.pageBreak = false;
      this.sectionHasContent = true;
      if (previousTable && child.type.name === "table") out += "<w:p/>";
      // a block the page view puts below a wrapped image (not a plain paragraph): the text before it ends below the
      // image (wrapfig's \vtxwrapclear in the LaTeX)
      const top = node === this.doc;
      if (top && this.wrapOpen && clearsWrap(child)) {
        out = clearAfterLastParagraph(out);
        this.wrapOpen = false;
      }
      out += this.block(child, indent, quote);
      if (top && child.type.name === "paragraph" && child.firstChild && isImage(child.firstChild) && child.firstChild.attrs.wrap) this.wrapOpen = true;
      previousTable = child.type.name === "table";
    });
    return out;
  }

  /** A \newpage at the end of the document. */
  trailingPageBreak(): string {
    return this.pageBreak ? this.writePageBreak() : "";
  }

  private writePageBreak(): string {
    this.sectionPaged = true;
    return PAGE_BREAK;
  }

  /** The paragraph that ends the current section (its sectPr), and a new section with the same numbering. */
  private sectionBreak(nextType: PageSection["type"]): string {
    const xml = `<w:p><w:pPr>${this.sectPr()}</w:pPr></w:p>`;
    this.section = { ...this.section, start: null, firstEmpty: false, type: nextType };
    this.sectionHasContent = false;
    this.sectionPaged = false;
    this.pageBreak = false;
    return xml;
  }

  private applyPageCommand(command: string, value: string) {
    if (command === "pagenumbering") {
      this.section.shown = value !== "gobble" && this.section.shown;
      if (value !== "gobble") this.section.fmt = value;
      this.section.start = 1;
    } else if (command === "setcounter") this.section.start = Number(value) || 0;
    else if (command === "pagestyle") this.section.shown = value !== "empty";
    // \thispagestyle{empty} on a section's first page (a title page): Word's different first page, footer empty;
    // later in a section Word has no per-page footer and the page keeps its section's
    else if (command === "thispagestyle" && !this.sectionPaged && !(this.pageBreak && this.sectionHasContent)) this.section.firstEmpty = value === "empty";
  }

  /** The footer with the page number (or an empty one), as a part of the package. */
  private footer(shown: boolean): string {
    const key = shown ? "number" : "empty";
    let rel = this.footerRels.get(key);
    if (rel) return rel;
    const path = `footer${this.footers.length + 1}.xml`;
    const content = shown ? this.field("PAGE", "1") : "";
    this.footers.push(`${XML}<w:ftr ${NS}><w:p><w:pPr><w:jc w:val="center"/></w:pPr>${content}</w:p></w:ftr>`);
    rel = this.rel(path, "footer");
    this.footerRels.set(key, rel);
    return rel;
  }

  block(node: Node, indent = 0, quote = false): string {
    switch (node.type.name) {
      case "paragraph": {
        if (isImageLine(node)) return this.imageLine(node);
        const pPr = this.paragraphPr(node, { style: quote ? "Quote" : undefined, indentTwips: indent });
        const content = this.inline(node);
        const target = node.attrs.id && this.linkTargets.has(node.attrs.id);
        return this.paragraph(target ? this.bookmark(bookmarkName("_VtxPara_", node.attrs.id), content || this.run(" ")) : content, pPr);
      }
      case "title":
        return this.paragraph(this.inline(node), '<w:pStyle w:val="Title"/>');
      case "heading": {
        const level = Math.min(3, Math.max(1, node.attrs.level));
        const content = this.inline(node);
        const marked = node.attrs.id ? this.bookmark(bookmarkName("_VtxSec_", node.attrs.id), content) : content;
        return this.paragraph(marked, `<w:pStyle w:val="Heading${level}"/>`);
      }
      case "blockquote":
        return this.blocks(node, indent + 420, true);
      case "bullet_list":
      case "ordered_list":
        return this.list(node, indent, 0, quote);
      case "math_display":
        return this.displayFormula(node);
      case "horizontal_rule":
        return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="9E9E9E"/></w:pBdr></w:pPr></w:p>';
      case "code_block":
        return this.codeBlock(node, indent);
      case "table":
        return this.table(node);
      case "bibliography":
        return this.bibliography(node);
      default:
        return node.isTextblock ? this.paragraph(this.inline(node)) : node.isBlock ? this.blocks(node, indent, quote) : "";
    }
  }

  /** A code block as the page shows it: VS Code's Dark+ box with the highlighter's colours, a paragraph a line. */
  private codeBlock(node: Node, indent: number): string {
    const code = node.textContent;
    let lines: CodeToken[][] | null = null;
    try {
      lines = this.options.codeTokens?.(code, String(node.attrs.language ?? "")) ?? null;
    } catch {
      lines = null;
    }
    // the box's border (the code's colour) stands 10 pt beside the text
    const pPr = `<w:pStyle w:val="SourceCode"/>${indent ? `<w:ind w:left="${indent + CODE_INSET}" w:right="${CODE_INSET}"/>` : ""}`;
    return code
      .split("\n")
      .map((line, index) => {
        const tokens = lines?.[index];
        if (!tokens || tokens.map((t) => t.content).join("") !== line) return this.paragraph(this.run(line), pPr);
        const runs = tokens
          .map((token) => {
            const color = cssHex(token.color);
            const style = token.fontStyle ?? 0;
            const rPr = `${style & 2 ? "<w:b/><w:bCs/>" : ""}${style & 1 ? "<w:i/><w:iCs/>" : ""}${color ? `<w:color w:val="${color}"/>` : ""}${style & 4 ? '<w:u w:val="single"/>' : ""}`;
            return this.run(token.content, rPr);
          })
          .join("");
        return this.paragraph(runs, pPr);
      })
      .join("");
  }

  private list(node: Node, indent: number, depth: number, quote: boolean): string {
    const ordered = node.type.name === "ordered_list";
    let numId: number;
    if (ordered) {
      numId = this.nextNumId++;
      this.lists.push({ numId, ordered: true, start: node.attrs.start ?? 1 });
    } else numId = 2;
    const level = Math.min(depth, ordered ? 3 : 2);
    const itemIndent = indent + 420 * (level + 1);
    let out = "";
    node.forEach((item) => {
      item.forEach((child, _offset, index) => {
        if (child.type.name === "bullet_list" || child.type.name === "ordered_list") {
          out += this.list(child, indent, depth + 1, quote);
          return;
        }
        if (index === 0 && child.type.name === "paragraph") {
          const numPr = `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr>`;
          const pPr = this.paragraphPr(child, { style: "ListParagraph", numPr, indentTwips: indent ? itemIndent : 0 });
          out += this.paragraph(this.inline(child), pPr);
        } else if (child.type.name === "paragraph") {
          out += this.paragraph(this.inline(child), this.paragraphPr(child, { indentTwips: itemIndent }));
        } else out += this.block(child, itemIndent, quote);
      });
    });
    return out;
  }

  /** A table cell's padding: half an em of the text (the page view's 7 px at 14 px). */
  private cellPad(): number {
    return Math.round(0.5 * this.bodySizePt * 20);
  }

  private table(table: Node): string {
    const map = TableMap.get(table);
    const columnWidth = Math.floor(this.textWidthTwips / map.width);
    let out = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="${this.cellPad()}" w:type="dxa"/><w:left w:w="${this.cellPad()}" w:type="dxa"/><w:bottom w:w="${this.cellPad()}" w:type="dxa"/><w:right w:w="${this.cellPad()}" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>`;
    out += `<w:tblGrid>${Array.from({ length: map.width }, () => `<w:gridCol w:w="${columnWidth}"/>`).join("")}</w:tblGrid>`;
    for (let row = 0; row < map.height; row++) {
      const rowNode = table.child(row);
      const header = rowNode.firstChild?.type.name === "table_header";
      out += `<w:tr>${header ? "<w:trPr><w:tblHeader/></w:trPr>" : ""}`;
      for (let col = 0; col < map.width; ) {
        const offset = map.map[row * map.width + col];
        const cell = table.nodeAt(offset)!;
        const rect = map.findCell(offset);
        const span = rect.right - rect.left;
        const width = columnWidth * span;
        const props = `<w:tcW w:w="${width}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ""}`;
        if (rect.top !== row) {
          out += `<w:tc><w:tcPr>${props}<w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p/></w:tc>`;
        } else {
          const merge = rect.bottom - rect.top > 1 ? '<w:vMerge w:val="restart"/>' : "";
          const jc = cell.attrs.align ? `<w:jc w:val="${cell.attrs.align}"/>` : "";
          let content = "";
          cell.forEach((child) => {
            content += child.type.name === "paragraph" ? this.paragraph(this.inline(child), jc) : this.block(child);
          });
          out += `<w:tc><w:tcPr>${props}${merge}<w:vAlign w:val="center"/></w:tcPr>${content || "<w:p/>"}</w:tc>`;
        }
        col += span;
      }
      out += "</w:tr>";
    }
    return out + "</w:tbl>";
  }

  private bibliography(node: Node): string {
    let out = this.options.bibliographyTitle ? this.paragraph(this.run(this.options.bibliographyTitle), '<w:pStyle w:val="BibliographyHeading"/>') : "";
    node.forEach((entry) => {
      const content = this.inline(entry);
      const marked = entry.attrs.id ? this.bookmark(bookmarkName("_VtxBib_", entry.attrs.id), content || this.run(" ")) : content;
      out += this.paragraph(marked, '<w:pStyle w:val="Bibliography"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr>');
    });
    return out;
  }

  // --------------------------------------------------------------------- parts --

  sectPr(): string {
    const s = this.settings;
    const paper = PAPER_SIZES[s.paper];
    const t = (mm: number) => Math.round(mm * MM_TWIPS);
    const section = this.section;
    // in the schema's order: footer, footnotePr, type, pgSz, pgMar, pgNumType, cols, docGrid
    const footer =
      `<w:footerReference w:type="default" r:id="${this.footer(section.shown)}"/>` +
      (section.firstEmpty ? `<w:footerReference w:type="first" r:id="${this.footer(false)}"/>` : "");
    const numbering =
      section.fmt !== "arabic" || section.start !== null
        ? `<w:pgNumType w:fmt="${WORD_PAGE_FORMATS[section.fmt] ?? "decimal"}"${section.start !== null ? ` w:start="${section.start}"` : ""}/>`
        : "";
    return (
      `<w:sectPr>${footer}<w:footnotePr><w:numFmt w:val="decimal"/></w:footnotePr><w:type w:val="${section.type}"/><w:pgSz w:w="${t(paper.width)}" w:h="${t(paper.height)}"/>` +
      `<w:pgMar w:top="${t(s.marginTop)}" w:right="${t(s.marginRight)}" w:bottom="${t(s.marginBottom)}" w:left="${t(s.marginLeft)}" w:header="851" w:footer="992" w:gutter="0"/>` +
      numbering + '<w:cols w:space="425"/>' + (section.firstEmpty ? "<w:titlePg/>" : "") + '<w:docGrid w:linePitch="312"/></w:sectPr>'
    );
  }

  numberingXml(): string {
    const lvl = (i: number, fmt: string, text: string, left: number, hanging = 420, pstyle?: string, font?: string) =>
      `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>${pstyle ? `<w:pStyle w:val="${pstyle}"/>` : ""}<w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="${hanging}"/></w:pPr>${font ? `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr>` : ""}</w:lvl>`;
    const headings = this.settings.headingNumbering;
    const h = (n: number) => (headings === "1" ? Array.from({ length: n }, (_, i) => `%${i + 1}`).join(".") : Array.from({ length: n }, (_, i) => `%${i + 1}`).join(".") + ".");
    let xml = `${XML}<w:numbering xmlns:w="${W_NS}">`;
    xml += `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>${lvl(0, "decimal", h(1), 432, 432, "Heading1")}${lvl(1, "decimal", h(2), 576, 576, "Heading2")}${lvl(2, "decimal", h(3), 720, 720, "Heading3")}</w:abstractNum>`;
    xml += `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${lvl(0, "bullet", "•", 420, 420, undefined, "Arial")}${lvl(1, "bullet", "◦", 840, 420, undefined, "Arial")}${lvl(2, "bullet", "▪", 1260, 420, undefined, "Arial")}</w:abstractNum>`;
    xml += `<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="hybridMultilevel"/>${lvl(0, "decimal", "%1.", 420)}${lvl(1, "lowerLetter", "(%2)", 840)}${lvl(2, "lowerRoman", "%3.", 1260)}${lvl(3, "upperLetter", "%4.", 1680)}</w:abstractNum>`;
    xml += `<w:abstractNum w:abstractNumId="3"><w:multiLevelType w:val="singleLevel"/>${lvl(0, "decimal", "[%1]", 420)}</w:abstractNum>`;
    xml += '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="5"><w:abstractNumId w:val="3"/></w:num>';
    for (const list of this.lists) {
      xml += `<w:num w:numId="${list.numId}"><w:abstractNumId w:val="2"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="${list.start}"/></w:lvlOverride></w:num>`;
    }
    return xml + "</w:numbering>";
  }

  stylesXml(): string {
    const s = this.settings;
    const body = this.bodySizePt;
    const half = (ratio: number) => Math.round(body * ratio * 2);
    const style = (kind: string, id: string, name: string, pPr = "", rPr = "", based: string | null = "Normal", extra = "", next?: string) =>
      `<w:style w:type="${kind}" w:styleId="${id}"><w:name w:val="${name}"/>${based ? `<w:basedOn w:val="${based}"/>` : ""}${next ? `<w:next w:val="${next}"/>` : ""}${extra}${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}</w:style>`;
    // the editor's "single" is 1.3 × the size (ctex); atLeast lets a line with a tall formula grow (\lineskip)
    const line = Math.round(1.3 * body * s.lineSpacing * 20);
    /** twips of `em` ems of the text */
    const pt20 = (em: number) => Math.round(em * body * 20);
    const numbered = s.headingNumbering !== "none";
    const headingNum = (level: number) => (numbered ? `<w:numPr>${level ? `<w:ilvl w:val="${level}"/>` : ""}<w:numId w:val="1"/></w:numPr>` : "");
    return (
      `${XML}<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>` +
      `<w:rFonts w:ascii="${esc(s.latinFont)}" w:eastAsia="${esc(s.cjkFont)}" w:hAnsi="${esc(s.latinFont)}" w:cs="${esc(s.latinFont)}"/>` +
      `<w:kern w:val="2"/><w:sz w:val="${half(1)}"/><w:szCs w:val="${half(1)}"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="${line}" w:lineRule="atLeast"/></w:pPr></w:pPrDefault></w:docDefaults>` +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/><w:jc w:val="left"/></w:pPr></w:style>' +
      '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
      '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
      // as the page view sets them (document.css): the title \LARGE-like 1.728 em, not bold, 0.6 / 0.9 em around it;
      // headings 1.2 / 1.1 / 1 em, bold, half an em above and below; every line 1.3 × its own size at least
      style("paragraph", "Title", "Title", `<w:spacing w:before="${pt20(0.6 * 1.728)}" w:after="${pt20(0.9 * 1.728)}" w:line="${pt20(1.3 * 1.728 * s.lineSpacing)}" w:lineRule="atLeast"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/>`, `<w:sz w:val="${half(1.728)}"/><w:szCs w:val="${half(1.728)}"/>`, "Normal", "<w:qFormat/>", "Normal") +
      style("paragraph", "Heading1", "heading 1", `<w:keepNext/><w:keepLines/>${headingNum(0)}<w:spacing w:before="${pt20(0.6)}" w:after="${pt20(0.6)}" w:line="${pt20(1.3 * 1.2 * s.lineSpacing)}" w:lineRule="atLeast"/><w:jc w:val="left"/><w:outlineLvl w:val="0"/>`, `<w:b/><w:bCs/><w:sz w:val="${half(1.2)}"/><w:szCs w:val="${half(1.2)}"/>`, "Normal", "<w:qFormat/>", "Normal") +
      style("paragraph", "Heading2", "heading 2", `<w:keepNext/><w:keepLines/>${headingNum(1)}<w:spacing w:before="${pt20(0.55)}" w:after="${pt20(0.55)}" w:line="${pt20(1.3 * 1.1 * s.lineSpacing)}" w:lineRule="atLeast"/><w:jc w:val="left"/><w:outlineLvl w:val="1"/>`, `<w:b/><w:bCs/><w:sz w:val="${half(1.1)}"/><w:szCs w:val="${half(1.1)}"/>`, "Normal", "<w:qFormat/>", "Normal") +
      style("paragraph", "Heading3", "heading 3", `<w:keepNext/><w:keepLines/>${headingNum(2)}<w:spacing w:before="${pt20(0.5)}" w:after="${pt20(0.5)}"/><w:jc w:val="left"/><w:outlineLvl w:val="2"/>`, "<w:b/><w:bCs/>", "Normal", "<w:qFormat/>", "Normal") +
      style("paragraph", "ListParagraph", "List Paragraph", "<w:contextualSpacing/>", "", "Normal", "<w:qFormat/>") +
      style("paragraph", "Quote", "Quote", '<w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="CCCCCC"/></w:pBdr><w:spacing w:before="60" w:after="60"/><w:ind w:left="420" w:right="420"/>', "", "Normal", "<w:qFormat/>", "Normal") +
      style("paragraph", "SourceCode", "Source Code", `<w:pBdr>${["top", "left", "bottom", "right"].map((side) => `<w:${side} w:val="single" w:sz="4" w:space="${side === "top" || side === "bottom" ? 7 : 10}" w:color="1E1E1E"/>`).join("")}</w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="1E1E1E"/><w:wordWrap w:val="0"/><w:spacing w:before="${pt20(0.5)}" w:after="${pt20(0.5)}" w:line="${pt20(1.27)}" w:lineRule="exact"/><w:ind w:left="${CODE_INSET}" w:right="${CODE_INSET}"/><w:contextualSpacing/><w:jc w:val="left"/>`, `<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:noProof/><w:color w:val="D4D4D4"/><w:sz w:val="${half(14 / 15)}"/>`) +
      style("character", "VerbatimChar", "Verbatim Char", "", '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>', "DefaultParagraphFont") +
      style("character", "Hyperlink", "Hyperlink", "", '<w:color w:val="0563C1"/><w:u w:val="single"/>', "DefaultParagraphFont", '<w:uiPriority w:val="99"/><w:unhideWhenUsed/>') +
      style("paragraph", "FootnoteText", "footnote text", '<w:snapToGrid w:val="0"/><w:jc w:val="left"/>', `<w:sz w:val="${half(0.8)}"/><w:szCs w:val="${half(0.8)}"/>`, "Normal", '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>') +
      style("character", "FootnoteReference", "footnote reference", "", '<w:vertAlign w:val="superscript"/>', "DefaultParagraphFont", '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>') +
      style("paragraph", "BibliographyHeading", "Bibliography Heading", `<w:keepNext/><w:spacing w:before="${pt20(0.8)}" w:after="${pt20(0.3)}"/><w:jc w:val="left"/>`, "<w:b/><w:bCs/>", "Normal", "", "Bibliography") +
      style("paragraph", "Bibliography", "Bibliography", '<w:jc w:val="left"/>', "", "Normal", '<w:uiPriority w:val="37"/><w:unhideWhenUsed/>') +
      style(
        "table",
        "TableGrid",
        "Table Grid",
        "",
        "",
        "TableNormal",
        '<w:uiPriority w:val="39"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>',
      ) +
      "</w:styles>"
    );
  }

  settingsXml(): string {
    const font = esc(this.options.mathFont || "Cambria Math");
    return (
      `${XML}<w:settings xmlns:w="${W_NS}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">` +
      '<w:defaultTabStop w:val="420"/><w:characterSpacingControl w:val="compressPunctuation"/>' +
      '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>' +
      '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
      `<m:mathPr><m:mathFont m:val="${font}"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr>` +
      "</w:settings>"
    );
  }

  footnotesXml(): string {
    return (
      `${XML}<w:footnotes ${NS}>` +
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
      this.footnotes.join("") +
      "</w:footnotes>"
    );
  }

  hasFootnoteRels(): boolean {
    return this.footnoteRels.length > 0;
  }
}

const HIGHLIGHTS: Record<string, string> = {
  FFFF00: "yellow", "00FF00": "green", "00FFFF": "cyan", FF00FF: "magenta", "0000FF": "blue", FF0000: "red",
  "000080": "darkBlue", "008080": "darkCyan", "008000": "darkGreen", "800080": "darkMagenta", "800000": "darkRed",
  "808000": "darkYellow", "808080": "darkGray", C0C0C0: "lightGray", "000000": "black", FFFFFF: "white",
};

/** A CSS colour as Word's RRGGBB (null: not a colour Word can take). */
/** CSS colour names a document may carry (\textcolor{red}, a pasted style): their hex. */
const CSS_COLORS: Record<string, string> = {
  black: "000000", white: "FFFFFF", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00", cyan: "00FFFF",
  magenta: "FF00FF", gray: "808080", grey: "808080", darkgray: "A9A9A9", lightgray: "D3D3D3", orange: "FFA500",
  purple: "800080", brown: "A52A2A", pink: "FFC0CB", olive: "808000", teal: "008080", navy: "000080", maroon: "800000",
  lime: "00FF00", violet: "EE82EE", silver: "C0C0C0", gold: "FFD700", darkred: "8B0000", darkgreen: "006400",
  darkblue: "00008B", lightblue: "ADD8E6", lightgreen: "90EE90",
};

function cssHex(color: unknown): string | null {
  const value = String(color ?? "").trim();
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(value);
  if (hex) {
    const h = hex[1].length === 3 ? hex[1].split("").map((c) => c + c).join("") : hex[1];
    return h.toUpperCase();
  }
  const named = CSS_COLORS[value.toLowerCase()];
  if (named) return named;
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(value);
  if (rgb) return [rgb[1], rgb[2], rgb[3]].map((n) => Math.min(255, Number(n)).toString(16).padStart(2, "0")).join("").toUpperCase();
  return null;
}

const MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const PREFIXES: Record<string, string> = { [MATH_NS]: "m", [W_NS]: "w", [XML_NS]: "xml" };

/** OMML with the document's m:/w: prefixes. The writer's numbered equations use default namespaces and p6:-style
 * prefixes; Word reads them, WPS stops reading the document there. */
function canonicalOmml(xml: string): string {
  if (!/xmlns="|xmlns:p\d/.test(xml) || typeof DOMParser === "undefined") return xml;
  const parsed = new DOMParser().parseFromString(`<vtx-root xmlns:m="${MATH_NS}" xmlns:w="${W_NS}">${xml}</vtx-root>`, "application/xml");
  if (parsed.getElementsByTagName("parsererror").length) return xml;
  const name = (namespace: string | null, local: string, fallback: string) => {
    const prefix = namespace ? PREFIXES[namespace] : undefined;
    return prefix ? `${prefix}:${local}` : fallback;
  };
  const write = (element: Element): string => {
    let attributes = "";
    for (const attribute of Array.from(element.attributes)) {
      if (attribute.name === "xmlns" || attribute.name.startsWith("xmlns:")) continue;
      attributes += ` ${name(attribute.namespaceURI, attribute.localName, attribute.name)}="${esc(attribute.value)}"`;
    }
    let inner = "";
    element.childNodes.forEach((child) => {
      if (child.nodeType === 1) inner += write(child as Element);
      else if (child.nodeType === 3 || child.nodeType === 4) inner += esc(child.nodeValue ?? "");
    });
    const tag = name(element.namespaceURI, element.localName, element.tagName);
    return inner ? `<${tag}${attributes}>${inner}</${tag}>` : `<${tag}${attributes}/>`;
  };
  return Array.from(parsed.documentElement.children).map(write).join("");
}

const strToU8 = (text: string) => new TextEncoder().encode(text);

/**
 * The document as a .docx package's files (the app zips them, [Content_Types].xml first). `formulas` are the app's
 * renderings of docxFormulaRequests(doc), by key.
 */
export function documentToDocx(doc: Node, formulas: Map<string, DocxFormula>, options: DocxExportOptions = {}): Record<string, Uint8Array> {
  const writer = new DocxWriter(doc, formulas, options);
  const body = writer.blocks(doc) + writer.trailingPageBreak();
  const document = `${XML}<w:document ${NS}><w:body>${body}${writer.sectPr()}</w:body></w:document>`;
  const extensions = new Set(writer.media.map((m) => m.path.slice(m.path.lastIndexOf(".") + 1)));
  for (const part of formulas.values()) {
    if (part.kind === "object") for (const p of part.parts) writer.extraParts.push(p);
  }
  const contentTypes =
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    (extensions.has("png") ? '<Default Extension="png" ContentType="image/png"/>' : "") +
    (extensions.has("jpeg") ? '<Default Extension="jpeg" ContentType="image/jpeg"/>' : "") +
    (extensions.has("svg") ? '<Default Extension="svg" ContentType="image/svg+xml"/>' : "") +
    [...new Map(writer.extraParts.map((p) => [p.extension, p.contentType])).entries()].map(([ext, type]) => `<Default Extension="${ext}" ContentType="${type}"/>`).join("") +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
    writer.footers.map((_f, i) => `<Override PartName="/word/footer${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`).join("") +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
    "</Types>";
  const rootRels =
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>` +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    `<Relationship Id="rId3" Type="${R_NS}/extended-properties" Target="docProps/app.xml"/></Relationships>`;
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  const core =
    `${XML}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `<dc:title>${esc(options.title ?? "")}</dc:title><dc:creator>VisualTeX</dc:creator>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
  const app = `${XML}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>VisualTeX</Application></Properties>`;
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(rootRels),
    "docProps/core.xml": strToU8(core),
    "docProps/app.xml": strToU8(app),
    "word/document.xml": strToU8(document),
    "word/styles.xml": strToU8(writer.stylesXml()),
    "word/numbering.xml": strToU8(writer.numberingXml()),
    "word/settings.xml": strToU8(writer.settingsXml()),
    "word/footnotes.xml": strToU8(writer.footnotesXml()),
    "word/_rels/document.xml.rels": strToU8(writer.relsXml()),
  };
  if (writer.hasFootnoteRels()) files["word/_rels/footnotes.xml.rels"] = strToU8(writer.relsXml(true));
  writer.footers.forEach((footer, i) => (files[`word/footer${i + 1}.xml`] = strToU8(footer)));
  for (const media of writer.media) files[`word/${media.path}`] = media.bytes;
  for (const part of writer.extraParts) files[`word/${part.path}`] = part.bytes;
  return files;
}
