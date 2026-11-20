// LaTeX's tables read into the document's (latexImport.ts; their captions: figures.ts): tabular, tabular*, tabularx, tabulary and
// longtable, in a table float or not, with their caption and label; \multicolumn / \multirow spans, the columns'
// alignment (l c r, p{} m{} b{} X, >{\centering}, *{n}{…}), \hline / booktabs rules (booktabs: the table's rules),
// \makecell line breaks and threeparttable's notes. Nothing a cell holds is dropped: a cell under a \multirow that has
// text gives it to the spanning cell; what a float holds besides its tables stays as text after them. Only longtable's
// repeated page heads and continuation feet (\endhead after \endfirsthead, \endfoot) are left out: LaTeX prints them
// again on every page, the document's table has its rows once.
// Markdown has no table with spans, so the table travels to the Markdown reader as a marker (a paragraph of its own,
// as latexMarkers.ts carries footnotes) and is made into the table once the Markdown is read.
import { Fragment, type Node, type Schema } from "prosemirror-model";

type Align = "left" | "center" | "right" | null;

export interface LatexTableCell {
  /** the cell's text as Markdown */
  markdown: string;
  colspan: number;
  rowspan: number;
  align: Align;
}

export interface LatexTable {
  rows: LatexTableCell[][];
  rules: "grid" | "booktabs";
  /** the columns' widths in pt (p{} columns), null for the others */
  widths?: Array<number | null>;
  /** the caption as Markdown, or null */
  caption: string | null;
}

/** The {…} group starting at `open`: its content and the index after it. */
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

const skipSpaces = (text: string, i: number) => {
  while (i < text.length && /\s/.test(text[i])) i++;
  return i;
};

/** An optional [...] at i (spaces before it allowed): its content and the index after it, or null. */
function optional(text: string, i: number): { content: string; end: number } | null {
  const j = skipSpaces(text, i);
  if (text[j] !== "[") return null;
  let depth = 0;
  for (let k = j; k < text.length; k++) {
    if (text[k] === "{") depth++;
    else if (text[k] === "}") depth--;
    else if (text[k] === "]" && depth === 0) return { content: text.slice(j + 1, k), end: k + 1 };
  }
  return null;
}

/** A length in a column's width (3cm, 40mm, 2in, 120pt, 8em, 0.3\linewidth) in pt; null when it is not one. */
export function lengthPt(text: string): number | null {
  const match = /^\s*([\d.]+)\s*(pt|bp|cm|mm|in|em|ex|\\(?:linewidth|textwidth|columnwidth|hsize))\s*$/.exec(text);
  if (!match || !Number.isFinite(Number(match[1]))) return null;
  // \linewidth: A4's text area with the default margins (146.4 mm) — the document's own is not known here
  const unit: Record<string, number> = { pt: 1, bp: 72.27 / 72, cm: 72.27 / 2.54, mm: 72.27 / 25.4, in: 72.27, em: 10, ex: 4.3 };
  const factor = unit[match[2]] ?? (146.4 * 72.27) / 25.4;
  const value = Number(match[1]) * factor;
  return value > 0 ? Math.round(value * 2) / 2 : null;
}

/**
 * A column specification's columns, as alignments: l c r, p{} m{} b{} X (left), >{\centering} before one, *{n}{…};
 * `widths` (when given) gets each column's width in pt (p{}, m{}, b{}, w{}{}) or null.
 */
export function columnAligns(spec: string, widths?: Array<number | null>): Align[] {
  const out: Align[] = [];
  let pending: Align | undefined;
  for (let i = 0; i < spec.length; i++) {
    const c = spec[i];
    if (/\s|\|/.test(c)) continue;
    if (c === "@" || c === "!" || c === "<") {
      const arg = group(spec, skipSpaces(spec, i + 1));
      if (arg) i = arg.end - 1;
      continue;
    }
    if (c === ">") {
      const arg = group(spec, skipSpaces(spec, i + 1));
      if (arg) {
        i = arg.end - 1;
        if (/\\centering|\\Centering/.test(arg.content)) pending = "center";
        else if (/\\raggedleft|\\RaggedLeft/.test(arg.content)) pending = "right";
        else if (/\\raggedright|\\RaggedRight/.test(arg.content)) pending = "left";
      }
      continue;
    }
    if (c === "*") {
      const count = group(spec, skipSpaces(spec, i + 1));
      const repeated = count ? group(spec, skipSpaces(spec, count.end)) : null;
      if (count && repeated) {
        const n = Math.max(0, Math.min(64, Number(count.content.trim()) || 0));
        for (let k = 0; k < n; k++) out.push(...columnAligns(repeated.content, widths));
        i = repeated.end - 1;
      }
      continue;
    }
    let align: Align = null;
    if (c === "c" || c === "C") align = "center";
    else if (c === "r" || c === "R") align = "right";
    else if (c === "S") align = "center";
    else if (c === "D") {
      // dcolumn's D{.}{.}{2}
      let at = i + 1;
      for (let k = 0; k < 3; k++) {
        const arg = group(spec, skipSpaces(spec, at));
        if (arg) at = arg.end;
      }
      i = at - 1;
      align = "center";
    } else if (!/[A-Za-z]/.test(c)) continue;
    // a width argument (p{3cm}, m{}, b{}, X has none; a newcolumntype may take one)
    let columnWidth: number | null = null;
    if (/[pmbw]/.test(c)) {
      const arg = group(spec, skipSpaces(spec, i + 1));
      if (arg) i = arg.end - 1;
      if (arg && c !== "w") columnWidth = lengthPt(arg.content);
      if (c === "w") {
        // array's w{align}{width}
        const width = group(spec, skipSpaces(spec, i + 1));
        if (width) i = width.end - 1;
        if (width) columnWidth = lengthPt(width.content);
      }
    }
    widths?.push(columnWidth);
    out.push(pending !== undefined ? pending : align);
    pending = undefined;
  }
  return out;
}

/** Splits at a separator found outside braces (and not escaped). */
function splitTop(text: string, at: (text: string, i: number) => number): string[] {
  const parts: string[] = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === "\\") {
      if (depth === 0) {
        const length = at(text, i);
        if (length > 0) {
          parts.push(text.slice(start, i));
          start = i + length;
          i = start - 1;
          continue;
        }
      }
      i++;
      continue;
    }
    if (c === "{") depth++;
    else if (c === "}") depth = Math.max(0, depth - 1);
    else if (depth === 0) {
      const length = at(text, i);
      if (length > 0) {
        parts.push(text.slice(start, i));
        start = i + length;
        i = start - 1;
      }
    }
  }
  parts.push(text.slice(start));
  return parts;
}

/** A row's end: \\ (with * and [space] after it) or \tabularnewline / \cr. */
function rowEnd(text: string, i: number): number {
  const match = /^(?:\\\\\*?(?:\s*\[[^\]]*\])?|\\tabularnewline\b|\\cr\b|\\crcr\b)/.exec(text.slice(i, i + 40));
  return match ? match[0].length : 0;
}

/** The rules between rows (booktabs' and the grid's) and row colours: out of the rows. */
const RULES = /\\(?:hline|toprule|midrule|bottomrule|hdashline|morecmidrules|firsthline|lasthline)\b(?:\s*\[[^\]]*\])?|\\(?:cline|cdashline|hhline)\s*\{[^{}]*\}|\\cmidrule\s*(?:\([^)]*\))?\s*(?:\[[^\]]*\])?\s*\{[^{}]*\}|\\addlinespace\b(?:\s*\[[^\]]*\])?|\\specialrule\s*\{[^{}]*\}\s*\{[^{}]*\}\s*\{[^{}]*\}|\\(?:rowcolor|rowcolors)\s*(?:\[[^\]]*\])?\s*(?:\{[^{}]*\}\s*)+|\\arrayrulecolor\s*(?:\[[^\]]*\])?\s*\{[^{}]*\}|\\noalign\s*\{(?:[^{}]|\{[^{}]*\})*\}|\\(?:centering|raggedright|raggedleft|small|footnotesize|scriptsize|normalsize|large|tiny)\b/g;

/** A cell's own wrappers: \multicolumn, \multirow (\multirowcell), with their spans and alignment. */
function cellOf(raw: string, columnAlign: Align, convert: (latex: string) => string): LatexTableCell & { up: boolean } {
  let latex = raw.trim();
  let colspan = 1;
  let rowspan = 1;
  let align = columnAlign;
  let up = false;
  for (let guard = 0; guard < 4; guard++) {
    const multicolumn = /^\\multicolumn\s*(?=\{)/.exec(latex);
    if (multicolumn) {
      const n = group(latex, multicolumn[0].length);
      const spec = n ? group(latex, skipSpaces(latex, n.end)) : null;
      const body = spec ? group(latex, skipSpaces(latex, spec.end)) : null;
      if (n && spec && body) {
        colspan = Math.max(1, Math.min(64, Number(n.content.trim()) || 1));
        align = columnAligns(spec.content)[0] ?? null;
        latex = (body.content + latex.slice(body.end)).trim();
        continue;
      }
    }
    const multirow = /^\\multirow(?:cell)?\s*/.exec(latex);
    if (multirow) {
      let at = multirow[0].length;
      at = optional(latex, at)?.end ?? at;
      const n = group(latex, skipSpaces(latex, at));
      if (n) {
        at = n.end;
        at = optional(latex, at)?.end ?? at;
        // \multirow's width ({*}, {=}, {3cm}); \multirowcell has none
        if (!/cell/.test(multirow[0])) {
          const width = /^\s*\*/.exec(latex.slice(at));
          if (width) at += width[0].length;
          else {
            const w = group(latex, skipSpaces(latex, at));
            if (w) at = w.end;
          }
        }
        at = optional(latex, at)?.end ?? at;
        const body = group(latex, skipSpaces(latex, at));
        if (body) {
          const count = Number(n.content.trim()) || 1;
          up = count < 0;
          rowspan = Math.max(1, Math.min(256, Math.abs(Math.round(count))));
          latex = (body.content + latex.slice(body.end)).trim();
          continue;
        }
      }
    }
    break;
  }
  // line breaks inside a cell (\makecell{a\\b}, \shortstack, \newline): breaks in the cell's text
  latex = latex
    .replace(/\\(?:makecell|thead|shortstack|Centerstack|stackanchor)\s*(?:\[[^\]]*\])?\s*(?=\{)/g, "")
    .replace(/\\\\(?:\s*\[[^\]]*\])?/g, "<br>")
    .replace(/\\cellcolor\s*(?:\[[^\]]*\])?\s*\{[^{}]*\}/g, "")
    .replace(/\\tnote\s*\{([^{}]*)\}/g, "<sup>$1</sup>");
  return { markdown: convert(latex.replace(/\\(?:newline|linebreak)\b/g, "<br>")).trim(), colspan, rowspan, align, up };
}

/**
 * A tabular-like environment's rows. `body` is its content after the column specification; for a longtable the
 * caption row is taken out (its \caption and \label).
 */
export function parseTabular(
  spec: string,
  body: string,
  convert: (latex: string) => string,
): { table: LatexTable; caption: string | null; label: string | null } {
  let caption: string | null = null;
  let label: string | null = null;
  const booktabs = /\\(?:toprule|midrule|bottomrule|cmidrule)\b/.test(body);
  // longtable: the caption row, the heads and feet
  const captionAt = /\\caption\s*(?:\[[^\]]*\])?\s*(?=\{)/.exec(body);
  if (captionAt) {
    const arg = group(body, captionAt.index + captionAt[0].length);
    if (arg) {
      caption = arg.content;
      let end = arg.end;
      const labelAfter = /^\s*\\label\s*\{([^{}]*)\}/.exec(body.slice(end));
      if (labelAfter) {
        label = labelAfter[1].trim();
        end += labelAfter[0].length;
      }
      const rowBreak = /^\s*(?:\\\\\*?(?:\s*\[[^\]]*\])?|\\tabularnewline\b)/.exec(body.slice(end));
      if (rowBreak) end += rowBreak[0].length;
      body = body.slice(0, captionAt.index) + body.slice(end);
    }
  }
  const labelIn = /\\label\s*\{([^{}]*)\}/.exec(body);
  if (labelIn) {
    label ??= labelIn[1].trim();
    body = body.replace(/\\label\s*\{[^{}]*\}/g, "");
  }
  if (/\\end(?:firsthead|head|foot|lastfoot)\b/.test(body)) {
    const parts: Record<string, string> = {};
    let last = 0;
    for (const marker of body.matchAll(/\\end(firsthead|head|foot|lastfoot)\b/g)) {
      parts[marker[1]] = body.slice(last, marker.index);
      last = marker.index! + marker[0].length;
    }
    const rest = body.slice(last);
    // the first page's head (the repeated head only when there is no first one), the body, the last foot
    const head = parts.firsthead ?? parts.head ?? "";
    body = [head, rest, parts.lastfoot ?? ""].map((part) => part.trim().replace(/(\\\\|\\tabularnewline)?\s*$/, "")).filter(Boolean).join(" \\\\ ");
  }
  const widths: Array<number | null> = [];
  const aligns = columnAligns(spec, widths);
  const rows: Array<Array<LatexTableCell & { up: boolean; column: number }>> = [];
  for (const rawRow of splitTop(body.replace(RULES, " "), rowEnd)) {
    if (!rawRow.trim()) continue;
    const cells = splitTop(rawRow, (text, i) => (text[i] === "&" ? 1 : 0));
    const row: Array<LatexTableCell & { up: boolean; column: number }> = [];
    let column = 0;
    for (const raw of cells) {
      const cell = cellOf(raw, aligns[column] ?? null, convert);
      row.push({ ...cell, column });
      column += cell.colspan;
    }
    rows.push(row);
  }
  // a row of nothing (a rule's own row, a trailing \\): out
  while (rows.length && rows[rows.length - 1].every((cell) => !cell.markdown && cell.colspan === 1)) rows.pop();
  // \multirow: the cells it covers below (above, for a negative count) are the source's empty places; their text,
  // if any, goes into the spanning cell
  const removed = new Set<object>();
  const cellAt = (r: number, column: number) => rows[r]?.find((cell) => cell.column === column && !removed.has(cell));
  rows.forEach((row, r) => {
    for (const cell of row) {
      if (cell.rowspan < 2 || removed.has(cell)) continue;
      if (cell.up) {
        // \multirow{-n}: drawn from the cell n-1 rows up
        const top = Math.max(0, r - cell.rowspan + 1);
        const target = cellAt(top, cell.column);
        if (target && target !== cell) {
          target.markdown = [target.markdown, cell.markdown].filter(Boolean).join(" ");
          target.rowspan = r - top + 1;
          removed.add(cell);
          for (let k = top + 1; k < r; k++) {
            const covered = cellAt(k, cell.column);
            if (covered) {
              if (covered.markdown) target.markdown += ` ${covered.markdown}`;
              removed.add(covered);
            }
          }
        } else cell.rowspan = 1;
        continue;
      }
      const span = Math.min(cell.rowspan, rows.length - r);
      cell.rowspan = span;
      for (let k = 1; k < span; k++) {
        const covered = cellAt(r + k, cell.column);
        if (!covered) continue;
        if (covered.markdown) cell.markdown += ` ${covered.markdown}`;
        removed.add(covered);
      }
    }
  });
  const table: LatexTable = {
    rows: rows.map((row) => row.filter((cell) => !removed.has(cell)).map(({ markdown, colspan, rowspan, align }) => ({ markdown, colspan, rowspan, align }))).filter((row) => row.length),
    rules: booktabs ? "booktabs" : "grid",
    ...(widths.some((w) => w) ? { widths } : {}),
    caption: null,
  };
  return { table, caption, label };
}

// ------------------------------------------------------------------ marker --

const OPEN = "\uE001";
const CLOSE = "\uE002";
const MARKER = /^\s*\uE001T([A-Za-z0-9+/]*)\uE002\s*$/;

function encode(value: unknown): string {
  const bytes = new TextEncoder().encode(JSON.stringify(value));
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/=+$/, "");
}

function decode<T>(data: string): T | null {
  try {
    const binary = atob(data + "=".repeat((4 - (data.length % 4)) % 4));
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return JSON.parse(new TextDecoder().decode(bytes)) as T;
  } catch {
    return null;
  }
}

/** A table for the Markdown reader: a paragraph of its own. */
export const tableMarker = (table: LatexTable) => `${OPEN}T${encode(table)}${CLOSE}`;

/**
 * The table markers in a document read from Markdown, as tables with their captions. `parse` reads a cell's (or a
 * caption's) Markdown into a document.
 */
export function resolveTableMarkers(doc: Node, parse: (markdown: string) => Node): Node {
  if (!doc.textContent.includes(`${OPEN}T`)) return doc;
  const schema: Schema = doc.type.schema;
  const { table: tableType, table_row: rowType, table_cell: cellType, table_figure: floatType, caption: captionType, paragraph } = schema.nodes;
  if (!tableType || !floatType || !captionType) return doc;
  const blocks = (markdown: string): Node[] => {
    const out: Node[] = [];
    if (markdown.trim()) {
      parse(markdown).forEach((child) => {
        if (child.type === paragraph || child.type.name === "math_display") out.push(child);
        else if (child.isTextblock) out.push(paragraph.create(null, child.content));
      });
    }
    return out.length ? out : [paragraph.create()];
  };
  const inlineOf = (markdown: string): Fragment => {
    const first = blocks(markdown).find((b) => b.type === paragraph);
    return first ? first.content : Fragment.empty;
  };
  const map = (content: Fragment): Fragment => {
    const out: Node[] = [];
    content.forEach((child) => {
      const match = child.isTextblock ? MARKER.exec(child.textContent) : null;
      const data = match ? decode<LatexTable>(match[1]) : null;
      if (data && data.rows?.length) {
        // the p{} columns' widths on their cells (tableColumns.ts: colwidth, pt), by the column each cell starts in
        const taken: number[] = [];
        const rows = data.rows.map((row, r) => {
          let column = 0;
          return rowType.create(
            null,
            row.map((cell) => {
              while ((taken[column] ?? -1) >= r) column++;
              const colspan = cell.colspan || 1;
              const rowspan = cell.rowspan || 1;
              for (let k = 0; k < colspan; k++) if (rowspan > 1) taken[column + k] = r + rowspan - 1;
              const widths = data.widths ? Array.from({ length: colspan }, (_, k) => data.widths![column + k] ?? 0) : null;
              column += colspan;
              const colwidth = widths && widths.some((w) => w > 0) ? widths : null;
              return cellType.create({ colspan, rowspan, align: cell.align === "left" ? null : cell.align, colwidth }, blocks(cell.markdown));
            }),
          );
        });
        out.push(floatType.create({ rules: data.rules === "booktabs" ? "booktabs" : "grid" }, [captionType.create(null, data.caption ? inlineOf(data.caption) : Fragment.empty), tableType.create(null, rows)]));
        return;
      }
      out.push(child.isLeaf || child.isTextblock ? child : child.copy(map(child.content)));
    });
    return Fragment.from(out);
  };
  return doc.copy(map(doc.content));
}
