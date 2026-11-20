// What LaTeX has and Markdown has not — footnotes, citations, the reference list — carried through the Markdown
// latexImport.ts writes as markers (private-use characters around base64 data), and made into the document's
// footnote, cite and bibliography nodes once the Markdown is read (markdown.ts).
import { Fragment, type Node, type Schema } from "prosemirror-model";
import { newNodeId } from "./ids.ts";
import { splitInlineMath } from "./mathText.ts";

const OPEN = "\uE001";
const CLOSE = "\uE002";
const MARKER = /\uE001([FCB])([A-Za-z0-9+/]*)\uE002/g;

function encode(value: unknown): string {
  const bytes = new TextEncoder().encode(JSON.stringify(value));
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  // no "=" padding: "==" would read as a highlight in Markdown
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

/** A footnote: its text with $…$ formulas. */
export const footnoteMarker = (text: string) => `${OPEN}F${encode(text)}${CLOSE}`;
/** A citation of \bibitem / BibTeX keys. */
export const citeMarker = (keys: string[]) => `${OPEN}C${encode(keys)}${CLOSE}`;
/** A reference list (a paragraph of its own): its entries' keys and texts with $…$ formulas. */
export const bibliographyMarker = (entries: Array<{ key: string; text: string }>) => `${OPEN}B${encode(entries)}${CLOSE}`;

/**
 * The markers in a document read from Markdown, as nodes. `known` are the keys of reference list entries the
 * citations may also name (a paste's citations name the document's entries).
 */
export function resolveLatexMarkers(doc: Node, known: Map<string, string> = new Map()): Node {
  if (!doc.textContent.includes(OPEN)) return doc;
  const schema: Schema = doc.type.schema;
  const { footnote, cite, bibliography, bib_entry: entry, math_inline: math } = schema.nodes;
  // the entries first: citations may come before their list
  const keys = new Map(known);
  const lists = new Map<string, Array<{ key: string; text: string; id: string }>>();
  doc.descendants((node) => {
    if (!node.isText) return true;
    for (const match of node.text!.matchAll(MARKER)) {
      if (match[1] !== "B" || lists.has(match[2])) continue;
      const entries = (decode<Array<{ key: string; text: string }>>(match[2]) ?? []).map((e) => ({ ...e, id: newNodeId() }));
      for (const e of entries) if (!keys.has(e.key)) keys.set(e.key, e.id);
      lists.set(match[2], entries);
    }
    return true;
  });
  const inlineContent = (text: string) =>
    splitInlineMath(text)
      .map((part) => (part.kind === "text" ? (part.text ? schema.text(part.text) : null) : math ? math.create({ latex: part.latex }) : schema.text(`$${part.latex}$`)))
      .filter((node): node is Node => !!node);
  const listOnly = /^\s*\uE001B([A-Za-z0-9+/]*)\uE002\s*$/;
  const map = (content: Fragment): Fragment => {
    const out: Node[] = [];
    content.forEach((child) => {
      const list = child.isTextblock ? listOnly.exec(child.textContent) : null;
      if (list && bibliography && entry) {
        const entries = lists.get(list[1]) ?? [];
        if (entries.length) out.push(bibliography.create(null, entries.map((e) => entry.create({ id: e.id, key: e.key }, inlineContent(e.text)))));
        return;
      }
      if (child.isText && child.text!.includes(OPEN)) {
        let last = 0;
        const text = child.text!;
        for (const match of text.matchAll(MARKER)) {
          if (match.index! > last) out.push(schema.text(text.slice(last, match.index), child.marks));
          last = match.index! + match[0].length;
          if (match[1] === "F") {
            const note = decode<string>(match[2]) ?? "";
            out.push(footnote ? footnote.create({ text: note }) : schema.text(` (${note})`, child.marks));
          } else if (match[1] === "C") {
            const names = decode<string[]>(match[2]) ?? [];
            const targets = names.map((key) => keys.get(key)).filter((id): id is string => !!id);
            // a citation of keys with no entry keeps them as text, as LaTeX shows them ([?])
            out.push(cite && targets.length ? cite.create({ targets }) : schema.text(`[${names.join(", ")}]`, child.marks));
          }
        }
        if (last < text.length) out.push(schema.text(text.slice(last), child.marks));
        return;
      }
      out.push(child.isLeaf ? child : child.copy(map(child.content)));
    });
    return Fragment.from(out);
  };
  return doc.copy(map(doc.content));
}
