// A document as its file: a Markdown file that holds nothing only VisualTeX reads, and beside it the document's
// memory — what Markdown has no syntax for (the page and its fonts, numbering, colours and sizes, alignment, which
// line is the title or a caption, page commands…). The two together give the document back exactly; the file alone
// gives its text, formulas, tables and pictures.
//
// The file is the document's lines, one per line; an object (a display formula, a table, a picture) stands with a
// blank line on each side, which is not a line of the document; a page command made here has no line (the memory
// keeps it with the block after it). Read, a file made elsewhere keeps what it says: a block that was not changed is
// written back as the file had it, so a file opened and saved unchanged is the same file, and an edit changes its own
// line.
//
// The memory keeps, for each run of lines whose reading is not the document's nodes, those nodes, named by the
// lines' text (a hash and which of the places with that text). A line another program rewrote no longer has its
// text: its entry finds no place and is left out, the others stay.
import type { Node, Schema } from "prosemirror-model";
import { isGeneratedId } from "./ids.ts";
import { imageEntry, isImageLine, resolveImageSource, setImageFileResolver } from "./images.ts";
import { FORMULA_JOINER, fileBlockWriter, ownImageWidth, readMarkdown } from "./markdown.ts";
import { withoutMarkdownSyntax } from "./markdownLive.ts";
import { parsePageCommand } from "./pageCommands.ts";
import { DOCUMENT_SETTING_ATTRS, ownEquationTag, type DocumentSettings } from "./settings.ts";
import { captionAfter, tableAfter, tableAt, tableEnd, toSourceDocument, withHeadings } from "./sourceMarkdown.ts";
import { captionLabelLength } from "./figures.ts";
import { NO_CAPTION } from "./markdown.ts";

export const DOCUMENT_MEMORY_FORMAT = "visualtex-document-memory";
export const DOCUMENT_MEMORY_VERSION = 1;

/** What the file's lines do not say about a run of them: the document's nodes for those lines. */
export interface MemoryBlock {
  /** the lines' text (joined by "\n"), hashed */
  hash: string;
  /** which of the places in the file with this text, from 0 */
  occurrence: number;
  /** how many lines */
  lines: number;
  /** the nodes those lines are (ProseMirror JSON); none: the lines are not part of the document's text */
  nodes: unknown[];
}

/** The record VisualTeX keeps beside a Markdown file. */
export interface DocumentMemory {
  format: typeof DOCUMENT_MEMORY_FORMAT;
  version: number;
  /** the settings the document chose (the others are the defaults) */
  settings: Partial<DocumentSettings>;
  blocks: MemoryBlock[];
  /** the files of the pictures the blocks show: image key -> path as the file would write it */
  images: Record<string, string>;
}

/** What a file was when it was read: enough to write its unchanged parts back as they were. */
export interface DocumentFileSource {
  /** the file's text */
  text: string;
  /** the line end most of its lines have */
  eol: LineEnd;
  finalNewline: boolean;
  /** each row's line end, when they are not all `eol` (a file edited on different systems); null when they are */
  ends: string[] | null;
  /** the document's blocks as read */
  signatures: string[];
  /** blocks the file wrote otherwise than this writer would: signature -> their lines, in the file's order */
  originals: Map<string, string[]>;
  /** the blank lines the file had before ("b" + signature) and after ("a" + signature) its objects, where one of
   *  them had another number than the one the writer leaves: for each such object in the file's order */
  junctions: Map<string, number[]>;
  /** the footnotes' definitions as the file had them (with the blank line before them, if any) */
  trailer: { signature: string; rows: string[] } | null;
  /** the pictures' files: image key -> path as written in the file */
  imagePaths: Map<string, string>;
  /** the names the file gave its formulas (\\label) */
  labels: string[];
}

export interface DocumentFileOptions {
  /** A formula's LaTeX as it leaves the app. */
  formulaLatex?: (latex: string) => string;
  /** The file the document was read from. */
  source?: DocumentFileSource | null;
  /** The pictures' files (image key -> path beside the file); a picture without one is named assets/image-<hash>. */
  imagePaths?: ReadonlyMap<string, string>;
  /** Names a file gave the document's formulas, for a document written anew from one read from a file (Save as). */
  labels?: Iterable<string>;
}

export interface WrittenDocumentFile {
  text: string;
  memory: DocumentMemory;
  /** every picture the document shows: image key -> its path in the file */
  images: Map<string, string>;
}

export interface ReadDocumentFile {
  doc: Node;
  source: DocumentFileSource;
  /** memory entries whose lines are no longer in the file */
  dropped: number;
}

// ------------------------------------------------------------------ hashing --

/** A 64-bit hash of a text as 16 hex digits (cyrb's mixing; not cryptographic: it names lines within one file). */
export function textHash(text: string): string {
  let h1 = 0xdeadbeef;
  let h2 = 0x41c6ce57;
  for (let i = 0; i < text.length; i++) {
    const ch = text.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (h2 >>> 0).toString(16).padStart(8, "0") + (h1 >>> 0).toString(16).padStart(8, "0");
}

interface NodeJson {
  type: string;
  attrs?: Record<string, unknown>;
  content?: NodeJson[];
  [key: string]: unknown;
}

/** A display formula's id that is a name a file gave it (its \label): part of what the formula is. */
const isNamedFormula = (type: string, id: unknown) => type === "math_display" && typeof id === "string" && !isGeneratedId(id);

/**
 * A node's JSON without the ids that are given anew when a document is read: every id but a formula's name and
 * the ids in `keep` (the ones references name). `labels`: the footnotes' labels stay.
 */
function withoutIds(json: NodeJson, keep: ReadonlySet<string> | null, labels: boolean): NodeJson {
  let attrs = json.attrs;
  if (attrs) {
    const id = attrs.id;
    if (typeof id === "string" && !isNamedFormula(json.type, id) && !keep?.has(id)) attrs = { ...attrs, id: null };
    if (!labels && json.type === "footnote" && typeof attrs.label === "string") attrs = { ...attrs, label: null };
  }
  const content = json.content?.map((child) => withoutIds(child, keep, labels));
  return { ...json, ...(attrs ? { attrs } : {}), ...(content ? { content } : {}) };
}

/** A node as the memory keeps it. */
function nodeJson(node: Node, keep: ReadonlySet<string> | null = null): unknown {
  return JSON.parse(JSON.stringify(withoutIds(node.toJSON() as NodeJson, keep, true)));
}

/**
 * A node as text for comparing: ids as the memory keeps them, and without the footnotes' labels (a footnote without
 * one takes a number in the file).
 */
function comparableJson(node: Node, keep: ReadonlySet<string> | null = null): string {
  return JSON.stringify(withoutIds(node.toJSON() as NodeJson, keep, false));
}

const signatures = new WeakMap<Node, string>();

/** What a block is, ids (but a formula's name) and footnote labels aside. */
function signature(node: Node): string {
  let value = signatures.get(node);
  if (value === undefined) {
    value = textHash(comparableJson(node));
    signatures.set(node, value);
  }
  return value;
}

/**
 * Whether two documents are the same document: the same nodes, marks and settings; ids may differ unless a
 * reference names them or they are a formula's name, and so may footnote labels.
 */
export function sameDocument(a: Node, b: Node): boolean {
  return comparableJson(a, referencedIds(a)) === comparableJson(b, referencedIds(b));
}

/** The ids the document's references name (formulas, reference list entries). */
function referencedIds(doc: Node): Set<string> {
  const ids = new Set<string>();
  doc.descendants((node) => {
    if (node.type.name === "eqref" && node.attrs.target) ids.add(node.attrs.target);
    else if (node.type.name === "cite") for (const id of node.attrs.targets as string[]) ids.add(id);
    return true;
  });
  return ids;
}

/** Whether a block carries an id of `ids` that its text cannot say (a formula's name — `names` too — is in its text). */
function hasId(node: Node, ids: ReadonlySet<string>, names: ReadonlySet<string>): boolean {
  if (!ids.size) return false;
  const carries = (n: Node) =>
    typeof n.attrs.id === "string" && ids.has(n.attrs.id) && !isNamedFormula(n.type.name, n.attrs.id) && !(n.type.name === "math_display" && names.has(n.attrs.id));
  let found = carries(node);
  if (!found)
    node.descendants((child) => {
      if (carries(child)) found = true;
      return !found;
    });
  return found;
}

// ------------------------------------------------------------------- reading --

/** A run of the file's rows and the block it is (null: a blank line that belongs to no block, or the footnotes). */
interface Item {
  start: number;
  end: number;
  node: Node | null;
}

const LINE_TYPES = new Set(["paragraph", "heading", "title", "author", "date", "page_command", "toc"]);

/** A block that stands apart from the lines around it (a blank line on each side in the file). */
function isObject(node: Node): boolean {
  return !LINE_TYPES.has(node.type.name) || isImageLine(node);
}

const isBlankLine = (node: Node | null) => !!node && node.type.name === "paragraph" && node.childCount === 0;

/** A picture in Markdown: ![alt](path), ![alt](<path>) with an optional "title". */
const PICTURE = /!\[((?:[^\]\\]|\\.)*)\]\(\s*(<[^>\n]*>|[^()\s]+)(?:\s+"[^"]*")?\s*\)/y;
const PICTURE_LINE = /^\s*(?:!\[(?:[^\]\\]|\\.)*\]\(\s*(?:<[^>\n]*>|[^()\s]+)(?:\s+"[^"]*")?\s*\)\s*)+$/;

/** A path as the file writes it, the way pictures are looked up: without <…> and "./", %-escapes read. */
export function picturePath(written: string): string {
  let value = written.trim().replace(/^<|>$/g, "");
  try {
    value = decodeURI(value);
  } catch {
    // as written
  }
  return value.replace(/^\.\//, "");
}

/** The paths of the pictures a Markdown text shows with ![…](…) (data: URLs aside). */
export function picturePaths(text: string): string[] {
  const found = new Set<string>();
  for (const m of text.matchAll(/!\[(?:[^\]\\]|\\.)*\]\(\s*(<[^>\n]*>|[^()\s]+)(?:\s+"[^"]*")?\s*\)/g)) {
    if (!/^<?data:/i.test(m[1])) found.add(picturePath(m[1]));
  }
  return [...found];
}

interface ReadContext {
  schema: Schema;
  /** the footnotes' texts by label (null: footnotes are not read) */
  notes: Map<string, string> | null;
  /** how often each footnote's reference was read */
  noteUses: Map<string, number>;
  /** the pictures read: image key -> path as written */
  imagePaths: Map<string, string>;
}

const lineCache = new Map<string, Node>();
const tableCache = new Map<string, Node | null>();
const CACHE_LIMIT = 20000;

function cached<T>(cache: Map<string, T>, key: string, make: () => T): T {
  if (cache.has(key)) return cache.get(key)!;
  const value = make();
  if (cache.size >= CACHE_LIMIT) cache.clear();
  cache.set(key, value);
  return value;
}

/** A row of text with nothing read in it (code, front matter): one line. */
function plainLine(schema: Schema, row: string): Node {
  return cached(lineCache, "p" + row, () => schema.nodes.paragraph.create(null, row ? schema.text(row) : null));
}

/**
 * A row as a line: its text, with `$…$` and `\(…\)` formulas (Pandoc's rule for dollars: nothing blank just inside
 * them, no digit after the closing one; `\$` inside is a dollar of the formula), a footnote's reference, and a
 * picture whose file is known. Nothing is read in a code span.
 */
function lineNode(context: ReadContext, row: string): Node {
  if (!/[$\\[!]/.test(row)) return plainLine(context.schema, row);
  // (a row with a picture or a footnote reads by the pictures and definitions at hand: read each time)
  const cacheable = !/!\[|\[\^/.test(row);
  const kept = cacheable ? lineCache.get("l" + row) : undefined;
  if (kept) return kept;
  const { schema, notes } = context;
  const children: Node[] = [];
  let text = "";
  const flush = () => {
    if (text) children.push(schema.text(text));
    text = "";
  };
  let i = 0;
  while (i < row.length) {
    const ch = row[i];
    if (ch === "`") {
      // a code span: to the next run of as many backticks
      let n = 1;
      while (row[i + n] === "`") n++;
      const run = "`".repeat(n);
      let close = row.indexOf(run, i + n);
      while (close >= 0 && (row[close + n] === "`" || row[close - 1] === "`")) close = row.indexOf(run, close + n + 1);
      const end = close >= 0 ? close + n : i + n;
      text += row.slice(i, end);
      i = end;
      continue;
    }
    if (ch === "\\") {
      if (row[i + 1] === "(") {
        const close = row.indexOf("\\)", i + 2);
        const latex = close >= 0 ? row.slice(i + 2, close).trim() : "";
        if (latex) {
          flush();
          children.push(schema.nodes.math_inline.create({ latex }));
          i = close + 2;
          continue;
        }
      }
      text += row.slice(i, i + 2);
      i += 2;
      continue;
    }
    if (ch === "$") {
      // a joiner and "$$": an empty inline formula, as the writer keeps one (markdown.ts: inlineFormulaSource; "$$"
      // alone stays text)
      if (row[i - 1] === FORMULA_JOINER && row[i + 1] === "$" && row[i + 2] !== "$") {
        if (text.endsWith(FORMULA_JOINER)) text = text.slice(0, -1);
        flush();
        children.push(schema.nodes.math_inline.create({ latex: "" }));
        i += 2;
        if (row[i] === FORMULA_JOINER && /[\d$]/.test(row[i + 1] ?? "")) i++;
        continue;
      }
      const next = row[i + 1];
      const opens = next !== undefined && !/\s/.test(next) && next !== "$" && next !== FORMULA_JOINER && row[i - 1] !== "$";
      let close = -1;
      if (opens) {
        for (let j = i + 1; j < row.length; j++) {
          if (row[j] === "\\") j++;
          else if (row[j] === "$") {
            close = j;
            break;
          }
        }
      }
      if (close > i + 1 && !/[\s\\]/.test(row[close - 1]) && !/[\d$]/.test(row[close + 1] ?? "")) {
        // the joiner the writer put between a dollar of the text and the formula
        if (text.endsWith("$" + FORMULA_JOINER)) text = text.slice(0, -1);
        flush();
        children.push(schema.nodes.math_inline.create({ latex: row.slice(i + 1, close) }));
        i = close + 1;
        // and between the formula and a digit or a dollar after it
        if (row[i] === FORMULA_JOINER && /[\d$]/.test(row[i + 1] ?? "")) i++;
        continue;
      }
    }
    if (ch === "[" && row[i + 1] === "^" && notes) {
      const match = /^\[\^([^\]\s]+)\]/.exec(row.slice(i, i + 200));
      const note = match ? notes.get(match[1]) : undefined;
      if (match && note !== undefined) {
        flush();
        children.push(schema.nodes.footnote.create({ text: note, label: match[1] }));
        context.noteUses.set(match[1], (context.noteUses.get(match[1]) ?? 0) + 1);
        i += match[0].length;
        continue;
      }
    }
    if (ch === "!" && row[i + 1] === "[") {
      PICTURE.lastIndex = i;
      const match = PICTURE.exec(row);
      const key = match ? resolveImageSource(match[2]) : null;
      if (match && key) {
        flush();
        children.push(schema.nodes.image.create({ src: key, alt: match[1].replace(/\\(.)/g, "$1"), width: ownImageWidth(key) }));
        if (!context.imagePaths.has(key)) context.imagePaths.set(key, match[2].replace(/^<|>$/g, ""));
        i += match[0].length;
        continue;
      }
    }
    text += ch;
    i++;
  }
  flush();
  const node = schema.nodes.paragraph.create(null, children);
  return cacheable ? cached(lineCache, "l" + row, () => node) : node;
}

/**
 * A display formula's LaTeX as the file has it: a \label names the formula (its id, which references use) and
 * numbers it unless it has a \tag of its own; a \tag stays the formula's own.
 */
function displayFormula(schema: Schema, body: string): Node {
  let latex = body.trim();
  const label = /\\label\{([^}]*)\}/.exec(latex);
  const id = label?.[1].trim() || null;
  if (label && id) latex = (latex.slice(0, label.index).replace(/[ \t]+$/, "") + latex.slice(label.index + label[0].length)).trim();
  return schema.nodes.math_display.create({ latex, id, numbered: !!id && !ownEquationTag(latex) });
}

/** The rows of a display formula that opens at row `i` ("$$" … "$$" or "\[" … "\]"): its last row and LaTeX. */
function displayAt(rows: string[], i: number): { last: number; body: string } | null {
  const first = rows[i].trim();
  const open = first.startsWith("$$") ? "$$" : first.startsWith("\\[") ? "\\[" : null;
  if (!open) return null;
  const close = open === "$$" ? "$$" : "\\]";
  if (first.length > 4 && first.endsWith(close)) {
    const body = first.slice(2, -2);
    // ("\[1\] … \[2\]" is escaped brackets, not a formula)
    return open === "\\[" && /\\[[\]]/.test(body) ? null : { last: i, body };
  }
  // ("\[" opens a formula of several rows only on a row of its own: at a line's start it is mostly an escaped bracket)
  if (open === "\\[" && first !== "\\[") return null;
  let j = i + 1;
  while (j < rows.length && !rows[j].trim().endsWith(close)) j++;
  if (j >= rows.length) return null;
  // the rows between, without the indent of the opening row (a formula under a list item)
  const indent = /^\s*/.exec(rows[i])![0].length;
  const inner = rows.slice(i + 1, j).map((row) => (/^\s*$/.test(row.slice(0, indent)) ? row.slice(indent) : row.trimStart()));
  const last = rows[j].trim().slice(0, -2);
  return { last: j, body: [first.slice(2), ...inner, last].join("\n") };
}

/** A footnote's definition and the rows under it that go on with it (indented; a blank row only before such a row). */
const DEFINITION = /^\[\^([^\]\s]+)\]:[ \t]?(.*)$/;

/**
 * The footnotes' definitions when they are the file's last rows: where they start (with the blank row before them)
 * and their texts by label. Null when there are none there, or two have one label.
 */
function footnoteTrailer(rows: string[]): { start: number; notes: Map<string, string> } | null {
  for (let start = 0; start < rows.length; start++) {
    if (!DEFINITION.test(rows[start])) continue;
    const notes = new Map<string, string>();
    let valid = true;
    let i = start;
    while (i < rows.length && valid) {
      if (rows[i] === "") {
        i++;
        continue;
      }
      const match = DEFINITION.exec(rows[i]);
      if (!match || notes.has(match[1])) {
        valid = false;
        break;
      }
      const lines = [match[2]];
      i++;
      while (i < rows.length) {
        if (/^(?: {4}|\t)/.test(rows[i])) lines.push(rows[i].replace(/^(?: {4}|\t)/, ""));
        else if (rows[i] === "" && /^(?: {4}|\t)/.test(rows[i + 1] ?? "")) lines.push("");
        else break;
        i++;
      }
      notes.set(match[1], lines.join("\n"));
    }
    if (valid && notes.size) return { start: start > 0 && rows[start - 1] === "" ? start - 1 : start, notes };
    // (a definition that is not among the last rows: none of them is read)
    return null;
  }
  return null;
}

/** A table's rows read as the document's table (with its caption), or null when they are not one. */
function tableNode(schema: Schema, markdown: string): Node | null {
  const read = () => {
    try {
      const doc = toSourceDocument(readMarkdown(schema, markdown).doc);
      return doc.childCount === 1 && doc.firstChild!.type.name === "table_figure" ? doc.firstChild! : null;
    } catch {
      return null;
    }
  };
  // (a table with pictures reads by the pictures at hand: read each time)
  return /!\[|<img/i.test(markdown) ? read() : cached(tableCache, markdown, read);
}

/** A row of pictures read as the document's figure, or null when a picture's file is not known. */
function pictureLine(context: ReadContext, row: string): Node | null {
  const paths: string[] = [];
  for (const m of row.matchAll(/!\[(?:[^\]\\]|\\.)*\]\(\s*(<[^>\n]*>|[^()\s]+)(?:\s+"[^"]*")?\s*\)/g)) paths.push(m[1].replace(/^<|>$/g, ""));
  const keys = paths.map((path) => resolveImageSource(path));
  if (!keys.length || keys.some((key) => !key)) return null;
  let read: Node;
  try {
    read = toSourceDocument(readMarkdown(context.schema, row.trim()).doc);
  } catch {
    return null;
  }
  if (read.childCount !== 1 || read.firstChild!.type.name !== "figure") return null;
  keys.forEach((key, k) => {
    if (!context.imagePaths.has(key!)) context.imagePaths.set(key!, paths[k]);
  });
  return read.firstChild!;
}

/**
 * The rows of a file as the document's blocks, each with the rows it is: lines of text, and the objects the text
 * writes (display formulas, tables with their captions, rows of pictures, page commands, [toc]). A blank row beside
 * an object, between it and another block, is the file's spacing, not a line. `notes`: the footnotes' definitions
 * are the file's last rows.
 */
function readRows(context: ReadContext, all: string[], end: number): Item[] {
  const { schema } = context;
  const rows = end === all.length ? all : all.slice(0, end);
  const items: Item[] = [];
  let fence: string | null = null;
  // front matter: "---" on the first row to the next "---" (or "...")
  let front = -1;
  if (rows[0] === "---") front = rows.findIndex((row, i) => i > 0 && i < end && (row === "---" || row === "..."));
  for (let i = 0; i < end; i++) {
    const row = rows[i];
    if (i <= front) {
      items.push({ start: i, end: i + 1, node: plainLine(schema, row) });
      continue;
    }
    const fenceMark = /^\s*(`{3,}|~{3,})/.exec(row)?.[1];
    if (fence) {
      if (fenceMark && fenceMark[0] === fence[0] && fenceMark.length >= fence.length && !row.trim().slice(fenceMark.length).trim()) fence = null;
      items.push({ start: i, end: i + 1, node: plainLine(schema, row) });
      continue;
    }
    if (fenceMark) {
      fence = fenceMark;
      items.push({ start: i, end: i + 1, node: plainLine(schema, row) });
      continue;
    }
    if (row === "") {
      items.push({ start: i, end: i + 1, node: plainLine(schema, "") });
      continue;
    }
    const display = /^\s*(?:\$\$|\\\[)/.test(row) ? displayAt(rows, i) : null;
    if (display) {
      items.push({ start: i, end: display.last + 1, node: displayFormula(schema, display.body) });
      i = display.last;
      continue;
    }
    // a table, with its caption above it ("Table: …"), below it, or the note that it has none (sourceMarkdown.ts)
    const body = rows;
    const captionAbove = captionLabelLength(row, "table") > 0 ? tableAfter(body, i + 1) : -1;
    const tableStart = tableAt(body, i) ? i : captionAbove;
    if (tableStart >= 0) {
      let j = tableEnd(body, tableStart);
      if (captionAbove < 0) {
        const k = body[j] === "" ? j + 1 : j;
        const following = k < body.length ? tableAfter(body, k + 1) : -1;
        const below = k < body.length && captionLabelLength(body[k], "table") > 0 && (following < 0 || captionAfter(body, tableEnd(body, following)));
        if (below || (k < body.length && NO_CAPTION.test(body[k]))) j = k + 1;
      }
      const node = tableNode(schema, body.slice(i, j).join("\n"));
      if (node) {
        items.push({ start: i, end: j, node });
        i = j - 1;
        continue;
      }
    }
    if (PICTURE_LINE.test(row)) {
      const node = pictureLine(context, row);
      if (node) {
        items.push({ start: i, end: i + 1, node });
        continue;
      }
    }
    const command = row[0] === "\\" || /^\s+\\/.test(row) ? parsePageCommand(row) : null;
    if (command) {
      items.push({ start: i, end: i + 1, node: schema.nodes.page_command.create(command) });
      continue;
    }
    if (/^\s*\[toc\]\s*$/i.test(row)) {
      items.push({ start: i, end: i + 1, node: schema.nodes.toc.create() });
      continue;
    }
    items.push({ start: i, end: i + 1, node: lineNode(context, row) });
  }
  return items;
}

/** A blank row beside an object, between it and another block, is the file's spacing: no line of the document. */
function withoutSpacing(items: Item[]): Item[] {
  const spacing = new Set<number>();
  items.forEach((item, i) => {
    if (!item.node || !isObject(item.node)) return;
    if (i - 1 >= 1 && isBlankLine(items[i - 1].node) && !spacing.has(i - 1)) spacing.add(i - 1);
    if (i + 1 <= items.length - 2 && isBlankLine(items[i + 1].node)) spacing.add(i + 1);
  });
  return items.map((item, i) => (spacing.has(i) ? { ...item, node: null } : item));
}

const REFERENCE = /^\s*\\(?:eq)?ref\{([^}]+)\}\s*$/;

/** `$\eqref{name}$` where a display formula has that name: a reference to it. */
function withReferences(schema: Schema, items: Item[]): Item[] {
  const names = new Set<string>();
  for (const item of items) if (item.node?.type.name === "math_display" && item.node.attrs.id) names.add(item.node.attrs.id);
  if (!names.size) return items;
  return items.map((item) => {
    const node = item.node;
    if (!node || node.type.name !== "paragraph" || !node.childCount) return item;
    let changed = false;
    const children: Node[] = [];
    node.forEach((child) => {
      const name = child.type.name === "math_inline" ? REFERENCE.exec(child.attrs.latex)?.[1]?.trim() : undefined;
      if (name && names.has(name)) {
        children.push(schema.nodes.eqref.create({ target: name }));
        changed = true;
      } else children.push(child);
    });
    return changed ? { ...item, node: node.type.create(node.attrs, children) } : item;
  });
}

interface RowsRead {
  /** every row of the file, in order: blocks, spacing, and (one item, no node) the footnotes' definitions */
  items: Item[];
  /** the footnotes' definitions: their rows start here (rows.length: none) */
  trailer: number;
  imagePaths: Map<string, string>;
}

/** The file's rows as the document's blocks, as the text alone gives them. */
function readText(schema: Schema, rows: string[]): RowsRead {
  const trailer = footnoteTrailer(rows);
  const read = (notes: Map<string, string> | null, end: number) => {
    const context: ReadContext = { schema, notes, noteUses: new Map(), imagePaths: new Map() };
    return { context, items: readRows(context, rows, end) };
  };
  let result = trailer ? read(trailer.notes, trailer.start) : read(null, rows.length);
  let end = trailer ? trailer.start : rows.length;
  if (trailer) {
    // every definition has its one reference, read as one; otherwise the footnotes stay the text they are
    const body = rows.slice(0, trailer.start).join("\n");
    const sound = [...trailer.notes.keys()].every((label) => result.context.noteUses.get(label) === 1 && body.split(`[^${label}]`).length === 2);
    if (!sound) {
      result = read(null, rows.length);
      end = rows.length;
    }
  }
  // lines typed as their text says (a "#" line a heading)
  const typed = withHeadings(schema.topNodeType.create(null, result.items.map((item) => item.node!)));
  let items: Item[] = result.items.map((item, i) => ({ ...item, node: typed.child(i) }));
  items = withReferences(schema, withoutSpacing(items));
  if (end < rows.length) items.push({ start: end, end: rows.length, node: null });
  return { items, trailer: end, imagePaths: result.context.imagePaths };
}

/** Where each run of `count` rows with one text is: text hash -> the rows they start at. */
function windows(rows: string[], count: number): Map<string, number[]> {
  const found = new Map<string, number[]>();
  for (let i = 0; i + count <= rows.length; i++) {
    const hash = textHash(count === 1 ? rows[i] : rows.slice(i, i + count).join("\n"));
    const list = found.get(hash);
    if (list) list.push(i);
    else found.set(hash, [i]);
  }
  return found;
}

type LineEnd = "\n" | "\r\n" | "\r";

/** A file's text as its rows, each row's line end ("" for a last row without one), and the line end most rows have. */
function splitRows(text: string): { rows: string[]; ends: string[]; eol: LineEnd; finalNewline: boolean } {
  const rows: string[] = [];
  const ends: string[] = [];
  const counts = { "\n": 0, "\r\n": 0, "\r": 0 };
  const breaks = /\r\n|\n|\r/g;
  let start = 0;
  for (let found = breaks.exec(text); found; found = breaks.exec(text)) {
    rows.push(text.slice(start, found.index));
    ends.push(found[0]);
    counts[found[0] as LineEnd]++;
    start = found.index + found[0].length;
  }
  rows.push(text.slice(start));
  ends.push("");
  const ended = rows.length > 1 && rows[rows.length - 1] === "";
  if (ended) {
    rows.pop();
    ends.pop();
  }
  const eol: LineEnd = counts["\r"] > counts["\n"] && counts["\r"] > counts["\r\n"] ? "\r" : counts["\r\n"] > counts["\n"] ? "\r\n" : "\n";
  // (an empty file has no habit: what is written into it ends with a newline)
  return { rows, ends, eol, finalNewline: ended || text === "" };
}

/**
 * The file's text from its rows. Where the file's rows end in different ways, the rows before the first row that
 * changed and after the last one keep their own line ends, and the changed rows take the ends of the rows they
 * replace, in order; rows added beyond those take the line end most rows have.
 */
function joinRows(rows: string[], source: DocumentFileSource | null): string {
  const eol = source?.eol ?? "\n";
  const last = (source?.finalNewline ?? true) ? eol : "";
  if (!source?.ends) return rows.join(eol) + last;
  const ends = source.ends;
  const old = splitRows(source.text).rows;
  let head = 0;
  while (head < rows.length && head < old.length && rows[head] === old[head]) head++;
  let tail = 0;
  while (tail < rows.length - head && tail < old.length - head && rows[rows.length - 1 - tail] === old[old.length - 1 - tail]) tail++;
  let text = "";
  for (let i = 0; i < rows.length; i++) {
    let end = i >= rows.length - tail ? ends[old.length - rows.length + i] : i < old.length - tail ? ends[i] : eol;
    // (the last row ends as the file did; a row that was the last one gets a line end when rows follow it now)
    if (i === rows.length - 1) end = last && end ? end : last;
    else if (!end) end = eol;
    text += rows[i] + end;
  }
  return text;
}

function settingsAttrs(settings: unknown): Record<string, unknown> {
  const attrs: Record<string, unknown> = {};
  if (settings && typeof settings === "object") {
    for (const [name, value] of Object.entries(settings)) {
      if (name in DOCUMENT_SETTING_ATTRS && (typeof value === "string" || typeof value === "number")) attrs[name] = value;
    }
  }
  return attrs;
}

/** A memory record read from its JSON text, or null when it is not one this version understands. */
export function parseDocumentMemory(text: string | null | undefined): DocumentMemory | null {
  if (!text) return null;
  try {
    const value = JSON.parse(text) as Partial<DocumentMemory>;
    if (value.format !== DOCUMENT_MEMORY_FORMAT || typeof value.version !== "number" || value.version > DOCUMENT_MEMORY_VERSION) return null;
    return {
      format: DOCUMENT_MEMORY_FORMAT,
      version: value.version,
      settings: (value.settings && typeof value.settings === "object" ? value.settings : {}) as Partial<DocumentSettings>,
      blocks: Array.isArray(value.blocks) ? value.blocks : [],
      images: value.images && typeof value.images === "object" ? value.images : {},
    };
  } catch {
    return null;
  }
}

export function serializeDocumentMemory(memory: DocumentMemory): string {
  return JSON.stringify(memory, null, 1);
}

/** Whether a memory says nothing (default settings, every line as the file's text reads). */
export function isEmptyMemory(memory: DocumentMemory): boolean {
  return !memory.blocks.length && !Object.keys(memory.settings).length;
}

/** The image keys of a memory's nodes given the keys their files have now (a picture changed on the disk). */
function withImageKeys(json: unknown, keys: Map<string, string>): unknown {
  if (!keys.size) return json;
  return JSON.parse(JSON.stringify(json), (key, value) => (key === "src" && typeof value === "string" && keys.has(value) ? keys.get(value) : value));
}

/**
 * A Markdown file as the document: its text read, then the memory's nodes put where their lines still stand. The
 * pictures the text and the memory name must be at hand (images.ts: setImageFileResolver) for them to be read.
 */
export function readDocumentFile(schema: Schema, text: string, memory: DocumentMemory | null = null): ReadDocumentFile {
  const { rows, ends, eol, finalNewline } = splitRows(text);
  const sameEnds = ends.every((end, i) => end === eol || (end === "" && i === ends.length - 1));
  const read = readText(schema, rows);
  const imagePaths = new Map(read.imagePaths);
  // the memory's pictures: the key their file has now
  const imageKeys = new Map<string, string>();
  for (const [key, path] of Object.entries(memory?.images ?? {})) {
    if (typeof path !== "string") continue;
    const now = resolveImageSource(path) ?? (imageEntry(key) ? key : null);
    if (now && now !== key) imageKeys.set(key, now);
    if (!imagePaths.has(now ?? key)) imagePaths.set(now ?? key, path);
  }
  // each entry's place: the rows with its text, if they are still whole blocks
  const starts = new Map<number, number>();
  read.items.forEach((item, i) => starts.set(item.start, i));
  const byCount = new Map<number, Map<string, number[]>>();
  const replaced = new Map<number, { until: number; nodes: Node[] }>();
  const taken = new Uint8Array(read.items.length);
  let dropped = 0;
  for (const block of memory?.blocks ?? []) {
    let nodes: Node[] | null = null;
    let first = -1;
    let until = -1;
    if (block && typeof block.hash === "string" && Number.isInteger(block.lines) && block.lines > 0 && Array.isArray(block.nodes)) {
      let places = byCount.get(block.lines);
      if (!places) byCount.set(block.lines, (places = windows(rows, block.lines)));
      const row = places.get(block.hash)?.[block.occurrence];
      first = row === undefined ? -1 : (starts.get(row) ?? -1);
      until = row === undefined ? -1 : row + block.lines === rows.length ? read.items.length : (starts.get(row + block.lines) ?? -1);
      if (first >= 0 && until > first && !taken.subarray(first, until).some(Boolean)) {
        try {
          nodes = block.nodes.map((json) => {
            const node = schema.nodeFromJSON(withImageKeys(json, imageKeys));
            node.check();
            return node;
          });
        } catch {
          nodes = null;
        }
      }
    }
    if (!nodes) {
      dropped++;
      continue;
    }
    taken.fill(1, first, until);
    replaced.set(first, { until, nodes });
  }
  // the document's blocks, each with where it came from
  const blocks: Array<{ node: Node; item: Item | null }> = [];
  const gaps: number[] = [];
  let gap = 0;
  let known = true;
  for (let i = 0; i < read.items.length; i++) {
    const swap = replaced.get(i);
    if (swap) {
      // (the spacing around what the memory gave is the writer's)
      for (const node of swap.nodes) {
        gaps.push(-1);
        blocks.push({ node, item: null });
      }
      gap = 0;
      known = false;
      i = swap.until - 1;
      continue;
    }
    const item = read.items[i];
    if (!item.node) {
      gap += item.end - item.start;
      continue;
    }
    gaps.push(known ? gap : -1);
    blocks.push({ node: item.node, item });
    gap = 0;
    known = true;
  }
  if (!blocks.length) blocks.push({ node: schema.nodes.paragraph.create(), item: null });
  let doc: Node;
  try {
    doc = schema.topNodeType.create(settingsAttrs(memory?.settings), blocks.map((block) => block.node));
    doc.check();
  } catch {
    // nodes that do not fit together: the text alone
    return readDocumentFile(schema, text, null);
  }
  // what the writer would write otherwise: the file's own lines for those blocks, and its spacing
  const labels = footnoteLabels(doc);
  // the names the file gave formulas (a \\label), which may look like the editor's own ids
  const names: string[] = [];
  for (const item of read.items) if (item.node?.type.name === "math_display" && item.node.attrs.id) names.push(item.node.attrs.id);
  const write = fileBlockWriter(doc, { footnoteLabels: labels, imagePath: (key) => imagePaths.get(key) ?? defaultImagePath(key), formulaNames: new Set(names) });
  const originals = new Map<string, string[]>();
  const differing = new Set<string>();
  const junctions = new Map<string, number[]>();
  const odd = new Set<string>();
  const push = (lists: Map<string, number[] | string[]>, key: string, value: number | string) => {
    const list = lists.get(key) as Array<number | string> | undefined;
    if (list) list.push(value);
    else lists.set(key, [value] as number[] | string[]);
  };
  blocks.forEach((block, i) => {
    const node = doc.child(i);
    const sig = signature(node);
    if (block.item) {
      const lines = rows.slice(block.item.start, block.item.end).join("\n");
      push(originals, sig, lines);
      if (lines !== write(node)) differing.add(sig);
    }
    if (isObject(node)) {
      // (at the file's ends and beside what the memory gave there is nothing to keep: the writer's one line)
      const before = i > 0 && gaps[i] >= 0 ? gaps[i] : 1;
      const after = i + 1 < blocks.length && gaps[i + 1] >= 0 ? gaps[i + 1] : 1;
      push(junctions, "b" + sig, before);
      push(junctions, "a" + sig, after);
      if (before !== 1) odd.add("b" + sig);
      if (after !== 1) odd.add("a" + sig);
    }
  });
  for (const sig of [...originals.keys()]) if (!differing.has(sig)) originals.delete(sig);
  for (const key of [...junctions.keys()]) if (!odd.has(key)) junctions.delete(key);
  const trailer = read.trailer < rows.length ? { signature: footnotesSignature(doc, labels), rows: rows.slice(read.trailer) } : null;
  return {
    doc,
    dropped,
    source: { text, eol, finalNewline, ends: sameEnds ? null : ends, signatures: blocks.map((_block, i) => signature(doc.child(i))), originals, junctions, trailer, imagePaths, labels: names },
  };
}

// ------------------------------------------------------------------- writing --

/** The file a picture without one gets: assets/image-<the start of its hash>.png|jpg. */
function defaultImagePath(key: string): string {
  const match = /([0-9a-f]{12})[0-9a-f]*\.(png|jpg)$/.exec(key);
  return match ? `assets/image-${match[1]}.${match[2]}` : `assets/${key.replace(/[^\w.-]+/g, "-")}`;
}

/** Each footnote's label in the file: the one it came with, else the first number no footnote has. */
function footnoteLabels(doc: Node): Map<Node, string> {
  const notes: Node[] = [];
  doc.descendants((node) => {
    if (node.type.name === "footnote") notes.push(node);
    return true;
  });
  const labels = new Map<Node, string>();
  const used = new Set<string>();
  for (const node of notes) {
    const label = typeof node.attrs.label === "string" && /^[^\]\s]+$/.test(node.attrs.label) && !used.has(node.attrs.label) ? node.attrs.label : null;
    if (label) {
      labels.set(node, label);
      used.add(label);
    }
  }
  let n = 1;
  for (const node of notes) {
    if (labels.has(node)) continue;
    while (used.has(String(n))) n++;
    labels.set(node, String(n));
    used.add(String(n));
  }
  return labels;
}

/** The footnotes in order, as [label, text] pairs. */
function footnoteList(doc: Node, labels: Map<Node, string>): Array<[string, string]> {
  const list: Array<[string, string]> = [];
  doc.descendants((node) => {
    if (node.type.name === "footnote") list.push([labels.get(node) ?? "?", String(node.attrs.text ?? "")]);
    return true;
  });
  return list;
}

function footnotesSignature(doc: Node, labels: Map<Node, string>): string {
  return JSON.stringify(footnoteList(doc, labels));
}

/** The image keys a document's nodes show. */
function imageKeysOf(node: Node, keys: Set<string>): void {
  node.descendants((child) => {
    if ((child.type.name === "image" || child.type.name === "board_image") && child.attrs.src) keys.add(child.attrs.src);
    return true;
  });
}

/**
 * The document as its file's text and its memory. Blocks still as the file had them are written as the file wrote
 * them (`source`); the memory holds the nodes of every run of lines that, read back, would not be the document's.
 */
export function writeDocumentFile(input: Node, options: DocumentFileOptions = {}): WrittenDocumentFile {
  const doc = withoutMarkdownSyntax(input);
  const source = options.source ?? null;
  const images = new Map<string, string>();
  const paths = new Set<string>();
  const imagePath = (key: string): string => {
    let path = images.get(key);
    if (path === undefined) {
      path = options.imagePaths?.get(key) ?? source?.imagePaths.get(key) ?? defaultImagePath(key);
      images.set(key, path);
      paths.add(path);
    }
    return path;
  };
  const labels = footnoteLabels(doc);
  // the names a file gave formulas stay theirs, however they look
  const names = new Set([...(source?.labels ?? []), ...(options.labels ?? [])]);
  const write = fileBlockWriter(doc, { formulaLatex: options.formulaLatex, footnoteLabels: labels, imagePath, formulaNames: names });
  // the blocks' lines: the file's own for a block still as it was, and the file's spacing between two such blocks
  const rows: string[] = [];
  const written: Array<{ start: number; end: number; nodes: Node[] }> = [];
  const used = new Map<string, number>();
  const take = <T>(lists: Map<string, T[]> | undefined, key: string): T | undefined => {
    const list = lists?.get(key);
    if (!list) return undefined;
    const index = used.get(key) ?? 0;
    used.set(key, index + 1);
    return list[index];
  };
  const sigs: string[] = [];
  let previous: Node | null = null;
  // the blank lines the file had after the object before
  let after: number | undefined;
  doc.forEach((node) => {
    const sig = signature(node);
    sigs.push(sig);
    const kept = take(source?.originals, sig);
    // a page command made here has no line in the file (a file's own "\\newpage" line is kept as its text)
    if (kept === undefined && node.type.name === "page_command") {
      written.push({ start: rows.length, end: rows.length, nodes: [node] });
      return;
    }
    const object = isObject(node);
    const before = object ? take(source?.junctions, "b" + sig) : undefined;
    if (previous) {
      const blank = before ?? after ?? (object || isObject(previous) ? 1 : 0);
      for (let k = 0; k < blank; k++) {
        written.push({ start: rows.length, end: rows.length + 1, nodes: [] });
        rows.push("");
      }
    }
    after = object ? take(source?.junctions, "a" + sig) : undefined;
    const lines = (kept ?? write(node)).split("\n");
    written.push({ start: rows.length, end: rows.length + lines.length, nodes: [node] });
    rows.push(...lines);
    previous = node;
  });
  // the footnotes' definitions after the text
  const notes = footnoteList(doc, labels);
  const sameNotes = !!source?.trailer && source.trailer.signature === JSON.stringify(notes);
  if (sameNotes || notes.length) {
    const start = rows.length;
    if (sameNotes) rows.push(...source!.trailer!.rows);
    else {
      rows.push("");
      for (const [label, text] of notes) {
        const lines = text.split("\n");
        rows.push(`[^${label}]: ${lines[0]}`.trimEnd(), ...lines.slice(1).map((line) => (line ? "    " + line : "")));
      }
    }
    written.push({ start, end: rows.length, nodes: [] });
  }
  // a page command without a line goes with the block after it (before it at the end), and the memory keeps it there;
  // a document of nothing else gets an empty line to keep it by
  const entries: typeof written = [];
  let pending: Node[] = [];
  for (const entry of written) {
    if (entry.start === entry.end) pending.push(...entry.nodes);
    else if (pending.length && entry.nodes.length) {
      entries.push({ ...entry, nodes: [...pending, ...entry.nodes] });
      pending = [];
    } else entries.push(entry);
  }
  if (pending.length) {
    let last = entries.length - 1;
    while (last >= 0 && !entries[last].nodes.length) last--;
    if (last >= 0) entries[last] = { ...entries[last], nodes: [...entries[last].nodes, ...pending] };
    else {
      entries.push({ start: rows.length, end: rows.length + 1, nodes: pending });
      rows.push("");
    }
  }
  // what the text alone reads as, with the pictures' paths as the file names them
  const byPath = new Map<string, string>();
  doc.descendants((node) => {
    if ((node.type.name === "image" || node.type.name === "board_image") && node.attrs.src) byPath.set(picturePath(imagePath(node.attrs.src)), node.attrs.src);
    return true;
  });
  const read = withImageFiles((path) => byPath.get(path) ?? null, () => readText(doc.type.schema, rows));
  // (an id something refers to must come back; only a formula's name is in the text)
  const referenced = referencedIds(doc);
  const keptIds = new Set([...referenced, ...names]);
  // runs of rows where the two differ: the document's nodes for them go into the memory
  const runs: Array<{ start: number; end: number; nodes: Node[] }> = [];
  let w = 0;
  let r = 0;
  while (w < entries.length && r < read.items.length) {
    const start = entries[w].start;
    const mine: Node[] = [...entries[w].nodes];
    const theirs: Node[] = read.items[r].node ? [read.items[r].node!] : [];
    let endW = entries[w++].end;
    let endR = read.items[r++].end;
    while (endW !== endR) {
      if (endW < endR) {
        mine.push(...entries[w].nodes);
        endW = entries[w++].end;
      } else {
        if (read.items[r].node) theirs.push(read.items[r].node!);
        endR = read.items[r++].end;
      }
    }
    const same = mine.length === theirs.length && mine.every((node, k) => signature(node) === signature(theirs[k]) && !hasId(node, referenced, names));
    if (!same) runs.push({ start, end: endW, nodes: mine });
  }
  // the file's spacing that read as a line goes with the block beside it
  const merged: typeof runs = [];
  for (let k = 0; k < runs.length; k++) {
    const run = runs[k];
    const last = merged[merged.length - 1];
    if (!run.nodes.length && runs[k + 1]?.start === run.end) runs[k + 1] = { ...runs[k + 1], start: run.start };
    else if (!run.nodes.length && last && last.end === run.start) last.end = run.end;
    else merged.push({ ...run });
  }
  const blocks: MemoryBlock[] = [];
  const memoryImages: Record<string, string> = {};
  const byCount = new Map<number, Map<string, number[]>>();
  for (const run of merged) {
    const count = run.end - run.start;
    let places = byCount.get(count);
    if (!places) byCount.set(count, (places = windows(rows, count)));
    const hash = textHash(rows.slice(run.start, run.end).join("\n"));
    blocks.push({ hash, occurrence: places.get(hash)!.indexOf(run.start), lines: count, nodes: run.nodes.map((node) => nodeJson(node, keptIds)) });
    const keys = new Set<string>();
    for (const node of run.nodes) imageKeysOf(node, keys);
    for (const key of keys) memoryImages[key] = imagePath(key);
  }
  const settings: Record<string, unknown> = {};
  for (const name of Object.keys(DOCUMENT_SETTING_ATTRS)) {
    const value = doc.attrs[name];
    if (value !== null && value !== undefined) settings[name] = value;
  }
  const unchanged = !!source && sigs.length === source.signatures.length && sigs.every((sig, i) => sig === source.signatures[i]) && (source.trailer ? sameNotes : !notes.length);
  const text = unchanged ? source!.text : joinRows(rows, source);
  return {
    text,
    memory: { format: DOCUMENT_MEMORY_FORMAT, version: DOCUMENT_MEMORY_VERSION, settings: settings as Partial<DocumentSettings>, blocks, images: memoryImages },
    images,
  };
}

/**
 * Paths for the document's pictures that have no file yet: beside the document in assets/, under the name of the
 * file the picture came from, or image-<date>-<time> for a pasted one. A name another picture has (`taken`: the
 * paths in use, of the document's pictures and of files already there) gets "-2", "-3"…; nothing is renumbered.
 */
export function newPicturePaths(doc: Node, known: ReadonlyMap<string, string>, taken: Iterable<string> = []): Map<string, string> {
  const used = new Set<string>();
  for (const path of [...known.values(), ...taken]) used.add(picturePath(path).toLowerCase());
  const paths = new Map<string, string>();
  const keys = new Set<string>();
  imageKeysOf(doc, keys);
  const two = (value: number) => String(value).padStart(2, "0");
  for (const key of keys) {
    if (known.has(key)) continue;
    const entry = imageEntry(key);
    if (!entry) continue;
    const extension = entry.mime === "image/png" ? "png" : "jpg";
    const when = new Date(entry.added ?? Date.now());
    const stamp = `image-${when.getFullYear()}${two(when.getMonth() + 1)}${two(when.getDate())}-${two(when.getHours())}${two(when.getMinutes())}${two(when.getSeconds())}`;
    // the file's own name, without what a file name cannot hold; its extension the one of the bytes kept
    const own = (entry.name ?? "").replace(/\.[^.]*$/, "").replace(/[\u0000-\u001f\\/:*?"<>|]+/g, "-").replace(/^[\s.-]+|[\s.-]+$/g, "").slice(0, 80);
    const stem = own || stamp;
    let path = `assets/${stem}.${extension}`;
    for (let k = 2; used.has(path.toLowerCase()); k++) path = `assets/${stem}-${k}.${extension}`;
    used.add(path.toLowerCase());
    paths.set(key, path);
  }
  return paths;
}

/** Runs `read` with the pictures' files resolved by `resolve` (a path as picturePath gives it -> image key). */
export function withImageFiles<T>(resolve: (path: string) => string | null, read: () => T): T {
  const before = setImageFileResolver(resolve);
  try {
    return read();
  } finally {
    setImageFileResolver(before);
  }
}

// ------------------------------------------------------------ bytes and text --

export type FileEncoding = "utf-8" | "utf-8-bom" | "utf-16le" | "utf-16be" | "gb18030" | "windows-1252";

/**
 * A file's bytes as text, and how they were encoded: UTF-8 (with or without a mark), UTF-16 with its mark, GB 18030
 * when they are not valid UTF-8, else Windows-1252 — which reads any bytes, so that a file in another encoding is
 * written back byte for byte where it was not edited.
 */
export function decodeFileText(bytes: Uint8Array): { text: string; encoding: FileEncoding } {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return { text: new TextDecoder("utf-8").decode(bytes.subarray(3)), encoding: "utf-8-bom" };
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return { text: new TextDecoder("utf-16le").decode(bytes.subarray(2)), encoding: "utf-16le" };
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return { text: new TextDecoder("utf-16be").decode(bytes.subarray(2)), encoding: "utf-16be" };
  try {
    return { text: new TextDecoder("utf-8", { fatal: true }).decode(bytes), encoding: "utf-8" };
  } catch {
    try {
      return { text: new TextDecoder("gb18030", { fatal: true }).decode(bytes), encoding: "gb18030" };
    } catch {
      return { text: new TextDecoder("windows-1252").decode(bytes), encoding: "windows-1252" };
    }
  }
}

/** The pointers of GB 18030's four-byte sequences for the characters of the Basic Multilingual Plane (0 to 39419). */
const GB_BMP_POINTERS = 39420;
/** The pointer of the first four-byte sequence of the supplementary planes (U+10000). */
const GB_SUPPLEMENTARY = 189000;

let gb18030: { pairs: Map<number, number>; pointers: Map<number, number> } | null = null;

/**
 * GB 18030 as the decoder reads it: code point -> two bytes (lead << 8 | trail), the first pair that reads as it
 * (0xA1A1 and 0xA3A0 both read as U+3000; 0xA1A1 is its encoding); and the other characters of the Basic Multilingual
 * Plane -> the pointer of their four bytes.
 */
function gb18030Tables(): { pairs: Map<number, number>; pointers: Map<number, number> } {
  if (gb18030) return gb18030;
  /** The character each `size`-byte sequence of `bytes` reads as (-1: none): all in one run when each reads as one
   *  character, else one by one. */
  const characters = (bytes: Uint8Array, size: number): (k: number) => number => {
    const count = bytes.length / size;
    const text = new TextDecoder("gb18030").decode(bytes);
    if (text.length === count) return (k) => text.charCodeAt(k);
    const strict = new TextDecoder("gb18030", { fatal: true });
    return (k) => {
      try {
        const one = strict.decode(bytes.subarray(k * size, (k + 1) * size));
        return one.length === 1 ? one.charCodeAt(0) : -1;
      } catch {
        return -1;
      }
    };
  };
  const pairBytes: number[] = [];
  for (let lead = 0x81; lead <= 0xfe; lead++) for (let trail = 0x40; trail <= 0xfe; trail++) if (trail !== 0x7f) pairBytes.push(lead, trail);
  const pairOf = characters(new Uint8Array(pairBytes), 2);
  const pairs = new Map<number, number>();
  for (let k = 0; k < pairBytes.length / 2; k++) {
    const code = pairOf(k);
    if (code >= 0 && code !== 0xfffd && !pairs.has(code)) pairs.set(code, (pairBytes[k * 2] << 8) | pairBytes[k * 2 + 1]);
  }
  const fourBytes = new Uint8Array(GB_BMP_POINTERS * 4);
  for (let pointer = 0; pointer < GB_BMP_POINTERS; pointer++) fourBytes.set(gbFourBytes(pointer), pointer * 4);
  const fourOf = characters(fourBytes, 4);
  const pointers = new Map<number, number>();
  for (let pointer = 0; pointer < GB_BMP_POINTERS; pointer++) {
    const code = fourOf(pointer);
    if (code >= 0 && !pairs.has(code) && !pointers.has(code)) pointers.set(code, pointer);
  }
  gb18030 = { pairs, pointers };
  return gb18030;
}

/** A four-byte sequence of GB 18030 by its pointer. */
function gbFourBytes(pointer: number): number[] {
  return [0x81 + Math.floor(pointer / 12600), 0x30 + (Math.floor(pointer / 1260) % 10), 0x81 + (Math.floor(pointer / 10) % 126), 0x30 + (pointer % 10)];
}

let windows1252: Map<number, number> | null = null;

/** Windows-1252 as the decoder reads it: code point -> byte (each of the 256 bytes reads as a character of its own). */
function windows1252Table(): Map<number, number> {
  if (windows1252) return windows1252;
  const all = new Uint8Array(256).map((_, byte) => byte);
  const text = new TextDecoder("windows-1252").decode(all);
  windows1252 = new Map();
  for (let byte = 0; byte < text.length; byte++) windows1252.set(text.charCodeAt(byte), byte);
  return windows1252;
}

/**
 * Text as a file's bytes in the encoding it was read in. A GB 18030 or Windows-1252 file whose text now has a
 * character that encoding has not is written as UTF-8 (the encoding returned says which was used).
 */
export function encodeFileText(text: string, encoding: FileEncoding): { bytes: Uint8Array; encoding: FileEncoding } {
  const utf8 = () => new TextEncoder().encode(text);
  if (encoding === "utf-8-bom") {
    const body = utf8();
    const bytes = new Uint8Array(body.length + 3);
    bytes.set([0xef, 0xbb, 0xbf]);
    bytes.set(body, 3);
    return { bytes, encoding };
  }
  if (encoding === "utf-16le" || encoding === "utf-16be") {
    const bytes = new Uint8Array(2 + text.length * 2);
    const view = new DataView(bytes.buffer);
    const little = encoding === "utf-16le";
    view.setUint16(0, 0xfeff, little);
    for (let i = 0; i < text.length; i++) view.setUint16(2 + i * 2, text.charCodeAt(i), little);
    return { bytes, encoding };
  }
  if (encoding === "gb18030") {
    const { pairs, pointers } = gb18030Tables();
    const bytes = new Uint8Array(text.length * 4);
    let length = 0;
    for (let i = 0; i < text.length; i++) {
      const code = text.charCodeAt(i);
      if (code < 0x80) {
        bytes[length++] = code;
        continue;
      }
      const pair = pairs.get(code);
      if (pair !== undefined) {
        bytes[length++] = pair >> 8;
        bytes[length++] = pair & 0xff;
        continue;
      }
      let pointer = pointers.get(code);
      // a character of the supplementary planes (a surrogate pair): its pointer follows from its code point
      const low = text.charCodeAt(i + 1);
      if (pointer === undefined && code >= 0xd800 && code <= 0xdbff && low >= 0xdc00 && low <= 0xdfff) {
        pointer = GB_SUPPLEMENTARY + ((code - 0xd800) << 10) + (low - 0xdc00);
        i++;
      }
      if (pointer === undefined) return { bytes: utf8(), encoding: "utf-8" };
      bytes.set(gbFourBytes(pointer), length);
      length += 4;
    }
    return { bytes: bytes.subarray(0, length), encoding };
  }
  if (encoding === "windows-1252") {
    const table = windows1252Table();
    const bytes = new Uint8Array(text.length);
    for (let i = 0; i < text.length; i++) {
      const byte = table.get(text.charCodeAt(i));
      if (byte === undefined) return { bytes: utf8(), encoding: "utf-8" };
      bytes[i] = byte;
    }
    return { bytes, encoding };
  }
  return { bytes: utf8(), encoding: "utf-8" };
}
