// Images in the document (DESIGN.md §11). An image is a box in a line, as LaTeX's \includegraphics is: alone in a
// paragraph it is an image line (2-3 of them side by side are one line), among text it is a character, and with
// `wrap` it is the paragraph's wrapfigure. Its width is a fraction of the line width (\linewidth); the bytes live in
// the document's image registry (keyed by their SHA-256), not in the node.
import { Fragment, type Node, type NodeSpec, type Schema } from "prosemirror-model";
import { EditorState, NodeSelection, Plugin, PluginKey, TextSelection, type Transaction } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView, type NodeView } from "prosemirror-view";
import { closeHistory } from "prosemirror-history";
import { documentSettings, PAPER_SIZES } from "./settings.ts";

// ------------------------------------------------------------------ registry --

export interface ImageEntry {
  key: string;
  mime: "image/png" | "image/jpeg";
  bytes: Uint8Array;
  /** natural size in pixels (0 when the header could not be read) */
  width: number;
  height: number;
  url?: string;
}

export const IMAGE_KEY_PREFIX = "vtx-image:";
/** The space between two images on an image line, as a fraction of the line (\hspace{0.02\linewidth}). */
export const IMAGE_GAP = 0.02;
/** What an image line's images and gaps may take of the line. */
export const IMAGE_LINE_ROOM = 0.99;
/** A new image's width at most (Overleaf's ½). */
export const DEFAULT_IMAGE_WIDTH = 0.5;

const registry = new Map<string, ImageEntry>();

const K = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01,
  0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc,
  0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147,
  0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116, 0x1e376c08,
  0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208,
  0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);

/** SHA-256 of the bytes, hex (synchronous: keys are needed while HTML is parsed). */
export function sha256Hex(bytes: Uint8Array): string {
  const length = bytes.length;
  const padded = new Uint8Array(((length + 9 + 63) >> 6) << 6);
  padded.set(bytes);
  padded[length] = 0x80;
  const view = new DataView(padded.buffer);
  view.setUint32(padded.length - 8, Math.floor((length * 8) / 0x100000000));
  view.setUint32(padded.length - 4, (length * 8) >>> 0);
  const h = new Uint32Array([0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19]);
  const w = new Uint32Array(64);
  const rotr = (x: number, n: number) => (x >>> n) | (x << (32 - n));
  for (let offset = 0; offset < padded.length; offset += 64) {
    for (let i = 0; i < 16; i++) w[i] = view.getUint32(offset + i * 4);
    for (let i = 16; i < 64; i++) {
      const s0 = rotr(w[i - 15], 7) ^ rotr(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      const s1 = rotr(w[i - 2], 17) ^ rotr(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) >>> 0;
    }
    let [a, b, c, d, e, f, g, hh] = h;
    for (let i = 0; i < 64; i++) {
      const s1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25);
      const ch = (e & f) ^ (~e & g);
      const t1 = (hh + s1 + ch + K[i] + w[i]) >>> 0;
      const s0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22);
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const t2 = (s0 + maj) >>> 0;
      hh = g;
      g = f;
      f = e;
      e = (d + t1) >>> 0;
      d = c;
      c = b;
      b = a;
      a = (t1 + t2) >>> 0;
    }
    h[0] = (h[0] + a) >>> 0;
    h[1] = (h[1] + b) >>> 0;
    h[2] = (h[2] + c) >>> 0;
    h[3] = (h[3] + d) >>> 0;
    h[4] = (h[4] + e) >>> 0;
    h[5] = (h[5] + f) >>> 0;
    h[6] = (h[6] + g) >>> 0;
    h[7] = (h[7] + hh) >>> 0;
  }
  return [...h].map((x) => x.toString(16).padStart(8, "0")).join("");
}

/** The pixel size a PNG or JPEG declares (its header), or null. */
export function imageSize(bytes: Uint8Array, mime: string): { width: number; height: number } | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (mime === "image/png") {
    if (bytes.length < 24 || bytes[0] !== 0x89 || bytes[1] !== 0x50) return null;
    return { width: view.getUint32(16), height: view.getUint32(20) };
  }
  if (bytes[0] !== 0xff || bytes[1] !== 0xd8) return null;
  let i = 2;
  while (i + 9 < bytes.length) {
    if (bytes[i] !== 0xff) {
      i++;
      continue;
    }
    const marker = bytes[i + 1];
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      i += 2;
      continue;
    }
    const length = view.getUint16(i + 2);
    // SOF0-SOF15 except DHT (C4), JPG (C8), DAC (CC)
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return { width: view.getUint16(i + 7), height: view.getUint16(i + 5) };
    }
    i += 2 + length;
  }
  return null;
}

function mimeOf(bytes: Uint8Array): "image/png" | "image/jpeg" | null {
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return "image/png";
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return "image/jpeg";
  return null;
}

/** PNG or JPEG bytes into the registry; their key (the same image twice is one entry). Null for other formats. */
export function registerImageBytes(bytes: Uint8Array): string | null {
  const mime = mimeOf(bytes);
  if (!mime) return null;
  const key = `${IMAGE_KEY_PREFIX}${sha256Hex(bytes)}.${mime === "image/png" ? "png" : "jpg"}`;
  if (!registry.has(key)) {
    const size = imageSize(bytes, mime) ?? { width: 0, height: 0 };
    registry.set(key, { key, mime, bytes, width: size.width, height: size.height });
  }
  return key;
}

export function imageEntry(key: string): ImageEntry | undefined {
  return registry.get(key);
}

/** The URL the editor shows the image from. */
export function imageUrl(key: string): string {
  const entry = registry.get(key);
  if (!entry) return "";
  if (!entry.url) {
    entry.url =
      typeof URL !== "undefined" && typeof Blob !== "undefined" && "createObjectURL" in URL
        ? URL.createObjectURL(new Blob([entry.bytes as BlobPart], { type: entry.mime }))
        : `data:${entry.mime};base64,${bytesToBase64(entry.bytes)}`;
  }
  return entry.url;
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

export function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64.replace(/\s+/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/** The image as a data: URL (copies, HTML). */
export function imageDataUrl(key: string): string {
  const entry = registry.get(key);
  return entry ? `data:${entry.mime};base64,${bytesToBase64(entry.bytes)}` : "";
}

/** A data:image/png|jpeg URL into the registry. */
export function registerDataUrl(url: string): string | null {
  const match = /^data:image\/(png|jpe?g);base64,(.*)$/is.exec(url.trim());
  if (!match) return null;
  try {
    return registerImageBytes(base64ToBytes(match[2]));
  } catch {
    return null;
  }
}

/** Files that come with a text being read (a .zip's figures/…): path -> key. */
let fileResolver: ((path: string) => string | null) | null = null;

/** While a project is read, the paths its text gives images resolve through this (null: none). */
export function setImageFileResolver(resolver: ((path: string) => string | null) | null): void {
  fileResolver = resolver;
}

/** An image source in a text being read (Markdown, LaTeX, HTML): a key, a data: URL or a file of the project. */
export function resolveImageSource(src: string): string | null {
  const value = src.trim();
  if (value.startsWith(IMAGE_KEY_PREFIX)) return registry.has(value) ? value : null;
  if (value.startsWith("data:")) return registerDataUrl(value);
  let path = value;
  try {
    path = decodeURI(value);
  } catch {
    // as written
  }
  return fileResolver?.(path.replace(/^<|>$/g, "").replace(/^\.\//, "")) ?? null;
}

/** Any image file the webview can decode: PNG and JPEG as they are, anything else drawn into a PNG. */
export async function registerImageFile(file: Blob): Promise<string | null> {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const direct = registerImageBytes(bytes);
  if (direct) return direct;
  try {
    const url = URL.createObjectURL(file);
    try {
      const img = new Image();
      img.src = url;
      await img.decode();
      const scale = /svg/i.test(file.type) ? 2 : 1;
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.round((img.naturalWidth || 300) * scale));
      canvas.height = Math.max(1, Math.round((img.naturalHeight || 150) * scale));
      canvas.getContext("2d")!.drawImage(img, 0, 0, canvas.width, canvas.height);
      const png = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, "image/png"));
      if (!png) return null;
      const key = registerImageBytes(new Uint8Array(await png.arrayBuffer()));
      // an SVG drawn at twice its size is shown at its own size
      if (key && scale !== 1) {
        const entry = registry.get(key)!;
        entry.width = Math.round(entry.width / scale);
        entry.height = Math.round(entry.height / scale);
      }
      return key;
    } finally {
      URL.revokeObjectURL(url);
    }
  } catch {
    return null;
  }
}

/** The image keys a document uses, in reading order. */
export function documentImageKeys(doc: Node): string[] {
  const keys: string[] = [];
  doc.descendants((node) => {
    if ((node.type.name === "image" || node.type.name === "board_image") && node.attrs.src && !keys.includes(node.attrs.src)) keys.push(node.attrs.src);
    return true;
  });
  return keys;
}

/** The document's images for its file: key -> base64. */
export function documentImagesForFile(doc: Node): Record<string, string> {
  const out: Record<string, string> = {};
  for (const key of documentImageKeys(doc)) {
    const entry = registry.get(key);
    if (entry) out[key] = bytesToBase64(entry.bytes);
  }
  return out;
}

/** A file's images back into the registry. */
export function loadDocumentImages(images: Record<string, string> | undefined): void {
  for (const [key, base64] of Object.entries(images ?? {})) {
    if (registry.has(key)) continue;
    try {
      registerImageBytes(base64ToBytes(base64));
    } catch {
      // a damaged image: the node shows nothing
    }
  }
}

/** Export files for the document's images: key -> path (`figures/fig-1.png`) and the bytes. */
export function imageFiles(doc: Node, folder: string): { paths: Map<string, string>; files: Array<{ path: string; bytes: Uint8Array }> } {
  const paths = new Map<string, string>();
  const files: Array<{ path: string; bytes: Uint8Array }> = [];
  for (const key of documentImageKeys(doc)) {
    const entry = registry.get(key);
    if (!entry) continue;
    const path = `${folder}/fig-${files.length + 1}.${entry.mime === "image/png" ? "png" : "jpg"}`;
    paths.set(key, path);
    files.push({ path, bytes: entry.bytes });
  }
  return { paths, files };
}

// --------------------------------------------------------------------- model --

export const imageNodeSpec: NodeSpec = {
  group: "inline",
  inline: true,
  atom: true,
  selectable: true,
  draggable: false,
  attrs: {
    id: { default: null as string | null },
    src: { default: "" },
    alt: { default: "" },
    width: { default: DEFAULT_IMAGE_WIDTH },
    wrap: { default: null as string | null },
  },
  parseDOM: [
    {
      tag: "img",
      getAttrs: (dom: HTMLElement | string) => {
        if (typeof dom === "string") return false;
        const key = dom.getAttribute("data-key");
        const src = key && registry.has(key) ? key : registerDataUrl(dom.getAttribute("src") ?? "");
        if (!src) return false;
        const width = Number(dom.getAttribute("data-width"));
        const wrap = dom.getAttribute("data-wrap");
        return {
          src,
          alt: dom.getAttribute("alt") ?? "",
          width: width > 0 && width <= 1 ? width : DEFAULT_IMAGE_WIDTH,
          wrap: wrap === "left" || wrap === "right" ? wrap : null,
        };
      },
    },
  ],
  // what a copy carries: the bytes, so a paste into another document (or another app) has the image
  toDOM: (node) => [
    "img",
    {
      src: imageDataUrl(node.attrs.src),
      alt: node.attrs.alt,
      "data-key": node.attrs.src,
      "data-width": String(node.attrs.width),
      ...(node.attrs.wrap ? { "data-wrap": node.attrs.wrap } : {}),
      style: `width:${(node.attrs.width * 100).toFixed(1)}%`,
    },
  ],
};

export const isImage = (node: Node | null | undefined): boolean => node?.type.name === "image";

/** A paragraph that holds images and nothing else (spaces aside): an image line. */
export function isImageLine(node: Node): boolean {
  if (node.type.name !== "paragraph" || node.childCount === 0) return false;
  let images = 0;
  let other = false;
  node.forEach((child) => {
    if (isImage(child) && !child.attrs.wrap) images++;
    else if (!(child.isText && !child.text!.trim())) other = true;
  });
  return images > 0 && !other;
}

/** What an image line's images and gaps take of the line. */
export function imageLineUse(line: Node): number {
  let total = 0;
  let count = 0;
  line.forEach((child) => {
    if (isImage(child)) {
      total += child.attrs.width;
      count++;
    }
  });
  return total + Math.max(0, count - 1) * IMAGE_GAP;
}

/** The text width of the document's page at zoom 1 (96 px per inch), px. */
export function lineWidthPx(doc: Node): number {
  const settings = documentSettings(doc);
  const paper = PAPER_SIZES[settings.paper];
  return Math.max(100, ((paper.width - settings.marginLeft - settings.marginRight) * 96) / 25.4);
}

export const roundWidth = (width: number) => Math.round(Math.max(0.01, Math.min(1, width)) * 1000) / 1000;

/** A new image's width: its natural size (96 px per inch) up to half the line. */
export function defaultImageWidth(doc: Node, key: string): number {
  const entry = registry.get(key);
  if (!entry?.width) return DEFAULT_IMAGE_WIDTH;
  return roundWidth(Math.min(DEFAULT_IMAGE_WIDTH, entry.width / lineWidthPx(doc)));
}

/** width / height of the image (1 when unknown). */
export function imageAspect(key: string): number {
  const entry = registry.get(key);
  return entry && entry.width && entry.height ? entry.width / entry.height : 1;
}

/**
 * The images of the image line at `linePos` made to fit the line: when they and the gaps take more than the room,
 * they get one height and fill it (Craft's and Paper's galleries).
 */
export function fitImageLine(tr: Transaction, linePos: number): Transaction {
  const line = tr.doc.nodeAt(linePos);
  if (!line || !isImageLine(line) || imageLineUse(line) <= IMAGE_LINE_ROOM + 1e-6) return tr;
  const images: Array<{ pos: number; node: Node }> = [];
  line.forEach((child, offset) => {
    if (isImage(child)) images.push({ pos: linePos + 1 + offset, node: child });
  });
  const room = IMAGE_LINE_ROOM - (images.length - 1) * IMAGE_GAP;
  const aspects = images.map(({ node }) => imageAspect(node.attrs.src));
  const sum = aspects.reduce((a, b) => a + b, 0);
  images.forEach(({ pos, node }, i) => {
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, width: Math.floor((room * aspects[i]) / sum * 1000) / 1000 });
  });
  return tr;
}

/** Whether images may be put in this textblock (not in tables, headings, the title, footnotes or code). */
export function imagesAllowedIn(state: EditorState, pos: number): boolean {
  const $pos = state.doc.resolve(pos);
  if ($pos.parent.type.name !== "paragraph") return false;
  for (let d = $pos.depth; d > 0; d--) {
    const name = $pos.node(d).type.name;
    if (name === "table_cell" || name === "table_header" || name === "bibliography") return false;
  }
  return true;
}

/** Whether a paragraph at this position can be wrapped: a body paragraph (top level), with text. */
function wrappable(doc: Node, paragraphPos: number): boolean {
  const $pos = doc.resolve(paragraphPos);
  const node = doc.nodeAt(paragraphPos);
  return $pos.depth === 0 && !!node && node.type.name === "paragraph" && node.textContent.trim().length > 0;
}

/**
 * Images put at the selection (paste, a dropped file, the image button): in an empty paragraph they take it; on an
 * image line they join it; in text (or anywhere else) they get an image line after that block (Mathcha: the
 * sentence is not split).
 */
export function insertImagesAtSelection(state: EditorState, keys: string[], schema: Schema): Transaction | null {
  return insertImageNodes(state, keys.map((key) => schema.nodes.image.create({ src: key, width: defaultImageWidth(state.doc, key) })), schema);
}

/** Image nodes put at the selection by the same rules (a paste of copied images). */
export function insertImageNodes(state: EditorState, images: Node[], schema: Schema): Transaction | null {
  if (!images.length) return null;
  const tr = closeHistory(state.tr);
  const sel = state.selection;
  const $from = sel.$from;
  // a selected board image: the new ones on its board, right of it (then below the board's images)
  if (sel instanceof NodeSelection && sel.node.type.name === "board_image" && schema.nodes.board_image) {
    const boardPos = $from.before();
    let x = sel.node.attrs.x + sel.node.attrs.width + IMAGE_GAP;
    let top = sel.node.attrs.top as number;
    let at = boardPos + $from.parent.nodeSize - 1;
    let first = -1;
    for (const image of images) {
      const width = Math.min(1, image.attrs.width);
      if (x + width > 1) {
        const board = tr.doc.nodeAt(boardPos)!;
        let bottom = 0;
        board.forEach((child) => {
          bottom = Math.max(bottom, child.attrs.top + child.attrs.width / imageAspect(child.attrs.src));
        });
        x = 0;
        top = bottom + IMAGE_GAP;
      }
      tr.insert(at, schema.nodes.board_image.create({ src: image.attrs.src, alt: image.attrs.alt, width, x: Math.round(x * 1000) / 1000, top: Math.round(top * 1000) / 1000 }));
      if (first < 0) first = at;
      at += 1;
      x += width + IMAGE_GAP;
    }
    tr.setSelection(NodeSelection.create(tr.doc, first));
    return tr;
  }
  // a selected image: beside it
  if (sel instanceof NodeSelection && isImage(sel.node)) {
    const parent = $from.parent;
    if (isImageLine(parent) && !sel.node.attrs.wrap) {
      tr.insert(sel.to, images);
      fitImageLine(tr, $from.before());
      tr.setSelection(NodeSelection.create(tr.doc, sel.to + images.slice(0, -1).reduce((n, img) => n + img.nodeSize, 0)));
      return tr;
    }
  }
  if ($from.parent.type.name === "paragraph" && imagesAllowedIn(state, $from.pos)) {
    const paragraph = $from.parent;
    const paragraphPos = $from.before();
    if (paragraph.content.size === 0) {
      tr.replaceWith(paragraphPos, paragraphPos + paragraph.nodeSize, schema.nodes.paragraph.create({ ...paragraph.attrs, align: paragraph.attrs.align ?? "center" }, images));
      fitImageLine(tr, paragraphPos);
      tr.setSelection(NodeSelection.create(tr.doc, paragraphPos + 1 + images.slice(0, -1).reduce((n, img) => n + img.nodeSize, 0)));
      return tr;
    }
    if (isImageLine(paragraph)) {
      const at = sel.empty ? $from.pos : sel.to;
      tr.insert(at, images);
      fitImageLine(tr, paragraphPos);
      tr.setSelection(NodeSelection.create(tr.doc, at + images.slice(0, -1).reduce((n, img) => n + img.nodeSize, 0)));
      return tr;
    }
  }
  // after the block the selection is in (its outermost block that is not a list or quote: inside a list item the
  // line goes after the paragraph; after a table, heading or display formula at their level)
  let depth = $from.depth;
  while (depth > 1 && !imagesAllowedIn(state, $from.before(depth)) && $from.node(depth).isTextblock) depth--;
  for (let d = $from.depth; d > 0; d--) {
    const name = $from.node(d).type.name;
    if (name === "table" || name === "bibliography") depth = d;
  }
  // a figure's or table's caption: after the figure or table (figures.ts)
  if (depth > 1 && ["figure", "table_figure"].includes($from.node(depth - 1).type.name)) depth--;
  if (depth < 1) depth = 1;
  const after = sel instanceof NodeSelection && !sel.node.isInline ? sel.to : $from.after(depth);
  const line = schema.nodes.paragraph.create({ align: "center" }, images);
  tr.insert(after, line);
  fitImageLine(tr, after);
  tr.setSelection(NodeSelection.create(tr.doc, after + 1 + images.slice(0, -1).reduce((n, img) => n + img.nodeSize, 0)));
  return tr;
}

/** The app's image picker asked for (the \\includegraphics command, the toolbar): an event on the editor. */
export const IMAGE_REQUEST_EVENT = "vtx-image-request";

export function requestImage(view: EditorView): boolean {
  view.dom.dispatchEvent(new CustomEvent(IMAGE_REQUEST_EVENT, { bubbles: true }));
  return true;
}

// ------------------------------------------------------------------- drops --

export type ImageDropTarget =
  /** a place in a line: beside an image on an image line, or in text */
  | { kind: "inline"; pos: number }
  /** a line of its own, between two blocks (pos: the boundary) */
  | { kind: "line"; pos: number }
  /** the wrapped image moved to the start of the body paragraph at `pos`, on that side */
  | { kind: "wrap"; pos: number; side: "left" | "right" }
  /** a place on the image board at `pos` (its left and top edges, fractions of the line width; imageBoards.ts) */
  | { kind: "board"; pos: number; x: number; top: number };

/** The block (outermost below a list, quote… at the pointer's textblock) whose boundary a "line" drop uses. */
function lineTarget(view: EditorView, pos: number, depth: number, y: number): ImageDropTarget {
  const $pos = view.state.doc.resolve(pos);
  const start = $pos.before(depth);
  const node = view.state.doc.nodeAt(start)!;
  const dom = view.nodeDOM(start) as HTMLElement | null;
  const rect = dom?.getBoundingClientRect();
  const below = rect ? y > (rect.top + rect.bottom) / 2 : false;
  return { kind: "line", pos: below ? start + node.nodeSize : start };
}

/** The top-level block at a height (or the nearest one), with its position. */
function blockAtHeight(view: EditorView, y: number): { pos: number; node: Node; rect: DOMRect } | null {
  let best: { pos: number; node: Node; rect: DOMRect; distance: number } | null = null;
  view.state.doc.forEach((node, offset) => {
    const dom = view.nodeDOM(offset) as HTMLElement | null;
    if (!dom || !(dom instanceof HTMLElement)) return;
    const rect = dom.getBoundingClientRect();
    const distance = y < rect.top ? rect.top - y : y > rect.bottom ? y - rect.bottom : 0;
    if (!best || distance < best.distance) best = { pos: offset, node, rect, distance };
  });
  return best;
}

/**
 * Where the image dragged from [from, to] lands at a page point: a line of its own between blocks (the pointer in
 * the margins, between blocks, or in the outer quarter of a block's first/last line), a place on an image line
 * (beside an image) or in text; a wrapped image wraps the body paragraph under the pointer on the nearer side.
 */
export function imageDropTarget(view: EditorView, x: number, y: number, image: Node, from: number, to: number): ImageDropTarget | null {
  const state = view.state;
  const text = view.dom.getBoundingClientRect();
  const style = getComputedStyle(view.dom);
  const left = text.left + parseFloat(style.paddingLeft || "0");
  const right = text.right - parseFloat(style.paddingRight || "0");
  if (image.attrs.wrap) {
    const block = blockAtHeight(view, y);
    if (block && wrappable(state.doc, block.pos) && !isImageLine(block.node) && y >= block.rect.top && y <= block.rect.bottom) {
      return { kind: "wrap", pos: block.pos, side: x < (block.rect.left + block.rect.right) / 2 ? "left" : "right" };
    }
  }
  // beside the text column: a line between the blocks there
  if (x < left - 2 || x > right + 2) {
    const block = blockAtHeight(view, y);
    if (!block) return null;
    return { kind: "line", pos: y > (block.rect.top + block.rect.bottom) / 2 ? block.pos + block.node.nodeSize : block.pos };
  }
  const hit = view.posAtCoords({ left: Math.max(left + 1, Math.min(right - 1, x)), top: y });
  if (!hit) return null;
  const $hit = state.doc.resolve(hit.pos);
  // the textblock under the pointer (an atom block such as a display formula: its own depth); a figure's line or
  // caption: the figure's (a line goes before or after the figure, figures.ts)
  const depth = $hit.depth > 1 && $hit.node($hit.depth - 1).type.name === "figure" ? $hit.depth - 1 : $hit.depth;
  const atom = hit.inside >= 0 ? state.doc.nodeAt(hit.inside) : null;
  if (atom && atom.isBlock && atom.isAtom) {
    const dom = view.nodeDOM(hit.inside) as HTMLElement | null;
    const rect = dom?.getBoundingClientRect();
    const below = rect ? y > (rect.top + rect.bottom) / 2 : false;
    return { kind: "line", pos: below ? hit.inside + atom.nodeSize : hit.inside };
  }
  if (!$hit.parent.isTextblock) {
    const block = blockAtHeight(view, y);
    if (!block) return null;
    return { kind: "line", pos: y > (block.rect.top + block.rect.bottom) / 2 ? block.pos + block.node.nodeSize : block.pos };
  }
  // outside tables (a line before or after the table)
  for (let d = $hit.depth; d > 0; d--) {
    if ($hit.node(d).type.name === "table") return lineTarget(view, hit.pos, d > 1 && $hit.node(d - 1).type.name === "table_figure" ? d - 1 : d, y);
  }
  const parent = $hit.parent;
  if (parent.type.name !== "paragraph" || !imagesAllowedIn(state, hit.pos)) return lineTarget(view, hit.pos, depth, y);
  const dom = view.nodeDOM($hit.before()) as HTMLElement | null;
  const rect = dom?.getBoundingClientRect();
  if (rect && parent.content.size > 0) {
    if (isImageLine(parent)) {
      const quarter = Math.min(24, rect.height / 4);
      if (y < rect.top + quarter || y > rect.bottom - quarter) return lineTarget(view, hit.pos, depth, y);
    } else {
      // the outer quarter of the first and the last line
      const first = view.coordsAtPos($hit.start());
      const last = view.coordsAtPos($hit.end());
      if (y < first.top + (first.bottom - first.top) / 4 || y > last.bottom - (last.bottom - last.top) / 4) {
        return lineTarget(view, hit.pos, depth, y);
      }
    }
  }
  // a place in the line: beside an atom under the pointer, the nearer side
  let pos = hit.pos;
  const inside = hit.inside >= 0 ? state.doc.nodeAt(hit.inside) : null;
  if (inside && inside.isInline) {
    const r = (view.nodeDOM(hit.inside) as HTMLElement | null)?.getBoundingClientRect();
    pos = r && x > (r.left + r.right) / 2 ? hit.inside + inside.nodeSize : hit.inside;
  }
  if (pos >= from && pos <= to) return null;
  return { kind: "inline", pos };
}

/** Shows where an image drop lands: a caret in a line, or a line across between blocks. */
export function imageDropMarker(view: EditorView, target: ImageDropTarget | null, element: HTMLElement): boolean {
  if (!target) return false;
  let style: Partial<CSSStyleDeclaration>;
  if (target.kind === "inline") {
    const c = view.coordsAtPos(target.pos);
    style = { left: `${c.left - 1}px`, top: `${c.top}px`, width: "2px", height: `${Math.max(10, c.bottom - c.top)}px` };
  } else if (target.kind === "line") {
    const $pos = view.state.doc.resolve(target.pos);
    const after = $pos.nodeAfter ? (view.nodeDOM(target.pos) as HTMLElement | null) : null;
    const before = !after && $pos.nodeBefore ? (view.nodeDOM(target.pos - $pos.nodeBefore.nodeSize) as HTMLElement | null) : null;
    const rect = (after ?? before)?.getBoundingClientRect();
    if (!rect) return false;
    const y = after ? rect.top - 2 : rect.bottom + 1;
    style = { left: `${rect.left}px`, top: `${y}px`, width: `${rect.width}px`, height: "2px" };
  } else if (target.kind === "wrap") {
    const dom = view.nodeDOM(target.pos) as HTMLElement | null;
    const rect = dom?.getBoundingClientRect();
    if (!rect) return false;
    const first = view.coordsAtPos(target.pos + 1);
    // the paragraph's start on that side: a bar the height of its first line at that edge
    const x = target.side === "left" ? rect.left : rect.right - 3;
    style = { left: `${x}px`, top: `${first.top}px`, width: "3px", height: `${Math.max(10, first.bottom - first.top) * 2}px` };
  } else return false;
  Object.assign(element.style, style);
  return true;
}

/** Text size and line width (px) of the paragraph a position is in. */
export function lineMetrics(view: EditorView, pos: number): { font: number; width: number } | null {
  const $pos = view.state.doc.resolve(pos);
  if (!$pos.parent.isTextblock) return null;
  let dom = view.nodeDOM($pos.before()) as HTMLElement | null;
  if (!dom) return null;
  // a wrapped figure's line is as wide as its image: the line it stands in (figures.ts)
  const wrapped = dom.parentElement?.closest<HTMLElement>(".vtx-figure.is-wrap-left, .vtx-figure.is-wrap-right");
  if (wrapped?.parentElement) dom = wrapped.parentElement;
  const style = getComputedStyle(dom);
  const width = dom.clientWidth - parseFloat(style.paddingLeft || "0") - parseFloat(style.paddingRight || "0");
  return { font: parseFloat(style.fontSize) || 16, width: Math.max(1, width) };
}

/**
 * The dragged image (from [from, to]) moved (or, with `copy`, copied) to the target, as one transaction: an image
 * line it leaves empty goes; on an image line it is fitted; in text a block-sized image is made a line high; on a
 * line of its own a line-high image gets its inserted size back.
 */
export function moveImage(view: EditorView, target: ImageDropTarget, from: number, to: number, copy: boolean): Transaction | null {
  const state = view.state;
  const image = state.doc.nodeAt(from);
  if (!image || !isImage(image)) return null;
  const schema = state.schema;
  // one undo step per gesture
  const tr = closeHistory(state.tr);
  let attrs = { ...image.attrs };
  // measured before the change (the target's line)
  const metrics = target.kind === "inline" ? lineMetrics(view, target.pos) : null;
  const sourceLine = state.doc.resolve(from).parent;
  const sourceMetrics = lineMetrics(view, from);
  const aspect = imageAspect(image.attrs.src);
  if (!copy) {
    const $from = state.doc.resolve(from);
    if ($from.parent.childCount === 1 && $from.parent.type.name === "paragraph" && $from.depth > 0 && target.kind !== "wrap") {
      // the image was alone in its paragraph: the paragraph goes with it (unless the drop lands in that paragraph)
      const start = $from.before();
      if (!(target.kind === "inline" && target.pos > start && target.pos < start + $from.parent.nodeSize)) {
        tr.delete(start, start + $from.parent.nodeSize);
      } else tr.delete(from, to);
    } else tr.delete(from, to);
  }
  // a line-high image (from a sentence) on a line of its own: its inserted size
  const lineHigh = sourceMetrics && !isImageLine(sourceLine) ? (image.attrs.width * sourceMetrics.width) / aspect <= sourceMetrics.font * 2 : false;
  if (target.kind === "board") {
    const boardPos = tr.mapping.map(target.pos);
    const board = tr.doc.nodeAt(boardPos)!;
    const at = boardPos + board.nodeSize - 1;
    const width = lineHigh || image.attrs.wrap ? defaultImageWidth(state.doc, image.attrs.src) : image.attrs.width;
    tr.insert(at, schema.nodes.board_image.create({ src: image.attrs.src, alt: image.attrs.alt, width: Math.min(1, width), x: target.x, top: target.top }));
    tr.setSelection(NodeSelection.create(tr.doc, at));
    return tr;
  }
  if (target.kind === "line") {
    attrs = { ...attrs, wrap: null, width: lineHigh || image.attrs.wrap ? defaultImageWidth(state.doc, image.attrs.src) : attrs.width };
    const pos = tr.mapping.map(target.pos);
    tr.insert(pos, schema.nodes.paragraph.create({ align: "center" }, schema.nodes.image.create(attrs)));
    tr.setSelection(NodeSelection.create(tr.doc, pos + 1));
    return tr;
  }
  if (target.kind === "wrap") {
    const pos = tr.mapping.map(target.pos);
    attrs = { ...attrs, wrap: target.side };
    tr.insert(pos + 1, schema.nodes.image.create(attrs));
    tr.setSelection(NodeSelection.create(tr.doc, pos + 1));
    return tr;
  }
  const pos = tr.mapping.map(target.pos);
  const $pos = tr.doc.resolve(pos);
  const parent = $pos.parent;
  const empty = parent.content.size === 0;
  const imageOnly = empty || isImageLine(parent);
  attrs = { ...attrs, wrap: null };
  if (!imageOnly && metrics) {
    // into a sentence: as high as the text (a block-sized image would open the line wide)
    const height = (attrs.width * metrics.width) / aspect;
    if (height > metrics.font * 2) attrs.width = roundWidth((metrics.font * aspect) / metrics.width);
  } else if (imageOnly && lineHigh) attrs.width = defaultImageWidth(state.doc, image.attrs.src);
  tr.insert(pos, schema.nodes.image.create(attrs));
  const linePos = $pos.before();
  if (empty && !parent.attrs.align) tr.setNodeMarkup(linePos, undefined, { ...parent.attrs, align: "center" });
  if (imageOnly) fitImageLine(tr, linePos);
  tr.setSelection(NodeSelection.create(tr.doc, pos));
  return tr;
}

// ------------------------------------------------------------------ editing --

/** The image line or paragraph of the selected image as an image line aligned left/centre/right (unwrapping it). */
export function setImageAlign(view: EditorView, align: "left" | "center" | "right"): boolean {
  const sel = view.state.selection;
  if (!(sel instanceof NodeSelection) || !isImage(sel.node)) return false;
  const state = view.state;
  const $pos = state.doc.resolve(sel.from);
  const tr = closeHistory(state.tr);
  const paragraph = $pos.parent;
  const paragraphPos = $pos.before();
  if (isImageLine(paragraph) && !sel.node.attrs.wrap) {
    tr.setNodeMarkup(paragraphPos, undefined, { ...paragraph.attrs, align: align === "left" ? null : align });
    tr.setSelection(NodeSelection.create(tr.doc, sel.from));
  } else if ($pos.depth > 1 && $pos.node($pos.depth - 1).type.name === "figure") {
    // a wrapped figure (figures.ts): its image line again, in place, the caption with it
    tr.setNodeMarkup(sel.from, undefined, { ...sel.node.attrs, wrap: null });
    tr.setNodeMarkup(paragraphPos, undefined, { ...paragraph.attrs, align: align === "left" ? null : align });
    tr.setSelection(NodeSelection.create(tr.doc, sel.from));
  } else {
    // out of the sentence (or out of the wrap): an image line of its own before the paragraph
    const attrs: Record<string, any> = { ...(sel.node.attrs as Record<string, any>), wrap: null };
    const metrics = lineMetrics(view, sel.from);
    const aspect = imageAspect(attrs.src);
    if (metrics && (attrs.width * metrics.width) / aspect <= metrics.font * 2) attrs.width = defaultImageWidth(state.doc, attrs.src);
    tr.delete(sel.from, sel.to);
    tr.insert(paragraphPos, state.schema.nodes.paragraph.create({ align: align === "left" ? null : align }, state.schema.nodes.image.create(attrs)));
    tr.setSelection(NodeSelection.create(tr.doc, paragraphPos + 1));
  }
  view.dispatch(tr.scrollIntoView());
  return true;
}

/** The selected image wrapped by text on a side: the start of its body paragraph (an image line: the next one). */
export function setImageWrap(view: EditorView, side: "left" | "right"): boolean {
  const sel = view.state.selection;
  if (!(sel instanceof NodeSelection) || !isImage(sel.node)) return false;
  const state = view.state;
  const target = wrapTargetFor(state, sel.from);
  if (target === null) return false;
  const tr = closeHistory(state.tr);
  const attrs = { ...sel.node.attrs, wrap: side };
  const $pos = state.doc.resolve(sel.from);
  const alone = isImageLine($pos.parent) && $pos.parent.childCount === 1;
  // the image line it leaves goes
  if (alone) tr.delete($pos.before(), $pos.after());
  else tr.delete(sel.from, sel.to);
  const pos = tr.mapping.map(target);
  tr.insert(pos + 1, state.schema.nodes.image.create(attrs));
  tr.setSelection(NodeSelection.create(tr.doc, pos + 1));
  view.dispatch(tr.scrollIntoView());
  return true;
}

/** The body paragraph a wrap of the image at `pos` goes to: its own (with text), else the next body paragraph. */
export function wrapTargetFor(state: EditorState, pos: number): number | null {
  const $pos = state.doc.resolve(pos);
  if ($pos.depth === 1 && wrappable(state.doc, $pos.before())) return $pos.before();
  // an image in a figure (figures.ts): the figure's neighbours
  const inFigure = $pos.depth === 2 && $pos.node(1).type.name === "figure";
  if ($pos.depth !== 1 && !inFigure) return null;
  const index = $pos.index(0);
  const next = index + 1 < state.doc.childCount ? state.doc.child(index + 1) : null;
  if (next && next.type.name === "paragraph" && next.textContent.trim()) return $pos.after(1);
  const previous = index > 0 ? state.doc.child(index - 1) : null;
  if (previous && previous.type.name === "paragraph" && previous.textContent.trim()) return $pos.before(1) - previous.nodeSize;
  return null;
}

/** An image's width set (a resize), as one step. */
export function setImageWidth(view: EditorView, pos: number, width: number): void {
  const node = view.state.doc.nodeAt(pos);
  if (!node || !isImage(node)) return;
  const tr = closeHistory(view.state.tr).setNodeMarkup(pos, undefined, { ...node.attrs, width: roundWidth(width) });
  tr.setSelection(NodeSelection.create(tr.doc, pos));
  view.dispatch(tr);
}

/** The widest an image on an image line may be: the line's room minus the others and the gaps. */
export function maxImageWidth(state: EditorState, pos: number): number {
  const $pos = state.doc.resolve(pos);
  const parent = $pos.parent;
  const node = state.doc.nodeAt(pos);
  if (!node) return 1;
  if (node.attrs.wrap) return 0.7;
  if (!isImageLine(parent)) return 1;
  return Math.max(0.05, IMAGE_LINE_ROOM - (imageLineUse(parent) - node.attrs.width));
}

// --------------------------------------------------------------------- view --

export interface ImageLabels {
  left: string;
  center: string;
  right: string;
  wrapLeft: string;
  wrapRight: string;
}

const ICONS: Record<string, string> = {
  left: '<path d="M4 5h16M4 19h16"/><rect x="4" y="8.5" width="9" height="7" rx="1"/>',
  center: '<path d="M4 5h16M4 19h16"/><rect x="7.5" y="8.5" width="9" height="7" rx="1"/>',
  right: '<path d="M4 5h16M4 19h16"/><rect x="11" y="8.5" width="9" height="7" rx="1"/>',
  wrapLeft: '<rect x="3" y="5" width="8" height="7" rx="1"/><path d="M14 6h7M14 10h7M3 15h18M3 19h18"/>',
  wrapRight: '<rect x="13" y="5" width="8" height="7" rx="1"/><path d="M3 6h7M3 10h7M3 15h18M3 19h18"/>',
};

const SNAPS = [0.25, 1 / 3, 0.5, 2 / 3, 0.75, 1];

class ImageView implements NodeView {
  dom: HTMLElement;
  private img: HTMLImageElement;
  private node: Node;
  private label: HTMLElement | null = null;

  constructor(
    node: Node,
    private view: EditorView,
    private getPos: () => number | undefined,
  ) {
    this.node = node;
    this.dom = document.createElement("span");
    this.dom.className = "vtx-image";
    this.dom.contentEditable = "false";
    this.img = document.createElement("img");
    this.img.draggable = false;
    this.img.alt = node.attrs.alt;
    this.dom.append(this.img);
    for (const side of ["left", "right"] as const) {
      const handle = document.createElement("span");
      handle.className = `vtx-image-handle is-${side}`;
      handle.addEventListener("pointerdown", (event) => this.resize(event, side));
      handle.addEventListener("dblclick", (event) => {
        event.preventDefault();
        event.stopPropagation();
        const pos = this.getPos();
        if (pos !== undefined) setImageWidth(this.view, pos, Math.min(defaultImageWidth(this.view.state.doc, this.node.attrs.src), maxImageWidth(this.view.state, pos)));
      });
      this.dom.append(handle);
    }
    this.render();
  }

  private render() {
    const { src, width, wrap, alt } = this.node.attrs;
    const url = imageUrl(src);
    if (this.img.getAttribute("src") !== url) this.img.src = url;
    this.img.alt = alt;
    this.dom.style.width = `${(width * 100).toFixed(2)}%`;
    const aspect = imageAspect(src);
    this.dom.style.setProperty("--vtx-image-aspect", String(aspect));
    this.dom.classList.toggle("is-wrap-left", wrap === "left");
    this.dom.classList.toggle("is-wrap-right", wrap === "right");
    this.dom.classList.toggle("is-missing", !url);
  }

  update(node: Node): boolean {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render();
    return true;
  }

  selectNode() {
    this.dom.classList.add("is-selected");
  }

  deselectNode() {
    this.dom.classList.remove("is-selected");
  }

  stopEvent(event: Event): boolean {
    return (event.target as Element | null)?.classList?.contains("vtx-image-handle") ?? false;
  }

  ignoreMutation(): boolean {
    return true;
  }

  /** A side handle dragged: the width follows the pointer (both sides of a centred line move), snapping. */
  private resize(event: PointerEvent, side: "left" | "right") {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    const pos = this.getPos();
    if (pos === undefined) return;
    const view = this.view;
    const $pos = view.state.doc.resolve(pos);
    // a wrapped figure (figures.ts) is as wide as its image: the line it stands in, and the figure follows the width
    const figure = this.dom.closest<HTMLElement>(".vtx-figure.is-wrap-left, .vtx-figure.is-wrap-right");
    const line = (figure?.parentElement ?? view.nodeDOM($pos.before())) as HTMLElement | null;
    if (!line) return;
    const style = getComputedStyle(line);
    const lineWidth = Math.max(1, line.clientWidth - parseFloat(style.paddingLeft || "0") - parseFloat(style.paddingRight || "0"));
    const wrap = this.node.attrs.wrap as string | null;
    const align = isImageLine($pos.parent) && !wrap ? ($pos.parent.attrs.align as string | null) ?? "left" : "left";
    // how the width follows the pointer: a centred line grows on both sides; dragging a handle outwards widens it
    const factor = align === "center" ? 2 : 1;
    const sign = side === "right" ? 1 : -1;
    const start = this.node.attrs.width as number;
    const max = maxImageWidth(view.state, pos);
    const others = new Set<number>();
    view.state.doc.descendants((node, at) => {
      if (isImage(node) && at !== pos) others.add(Math.round(node.attrs.width * 1000) / 1000);
      return true;
    });
    let width = start;
    const label = document.createElement("div");
    label.className = "vtx-image-size";
    document.body.append(label);
    this.label = label;
    this.dom.classList.add("is-resizing");
    const move = (e: PointerEvent) => {
      const dx = (e.clientX - event.clientX) * sign;
      let next = Math.max(0.03, Math.min(max, start + (dx * factor) / lineWidth));
      if (!e.altKey) {
        for (const snap of [...SNAPS, ...others]) {
          if (Math.abs(next - snap) <= 0.01 && snap <= max) {
            next = snap;
            break;
          }
        }
      }
      width = next;
      this.dom.style.width = `${(width * 100).toFixed(2)}%`;
      if (figure) figure.style.width = `${(width * 100).toFixed(2)}%`;
      label.textContent = `${Math.round(width * 100)}%`;
      const r = this.dom.getBoundingClientRect();
      label.style.left = `${side === "left" ? r.left : r.right}px`;
      label.style.top = `${r.top - 26}px`;
    };
    const up = () => {
      window.removeEventListener("pointermove", move, true);
      window.removeEventListener("pointerup", up, true);
      label.remove();
      this.label = null;
      this.dom.classList.remove("is-resizing");
      const now = this.getPos();
      if (now !== undefined && Math.abs(width - start) > 0.0005) setImageWidth(view, now, width);
      else this.render();
    };
    window.addEventListener("pointermove", move, true);
    window.addEventListener("pointerup", up, true);
    move(event);
  }

  destroy() {
    this.label?.remove();
  }
}

/** The bar above a selected image: alignment and wrap. */
class ImageBar {
  private element: HTMLElement;
  private buttons = new Map<string, HTMLButtonElement>();

  constructor(
    private view: EditorView,
    labels: ImageLabels,
  ) {
    this.element = document.createElement("div");
    this.element.className = "vtx-image-bar";
    this.element.setAttribute("role", "toolbar");
    const add = (name: keyof ImageLabels, run: () => void) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "vtx-image-bar-button";
      button.title = labels[name];
      button.setAttribute("aria-label", labels[name]);
      button.innerHTML = `<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">${ICONS[name]}</svg>`;
      button.addEventListener("pointerdown", (event) => event.preventDefault());
      button.addEventListener("click", (event) => {
        event.preventDefault();
        run();
        this.view.focus();
      });
      this.element.append(button);
      this.buttons.set(name, button);
    };
    add("left", () => setImageAlign(this.view, "left"));
    add("center", () => setImageAlign(this.view, "center"));
    add("right", () => setImageAlign(this.view, "right"));
    const gap = document.createElement("span");
    gap.className = "vtx-image-bar-gap";
    this.element.append(gap);
    add("wrapLeft", () => setImageWrap(this.view, "left"));
    add("wrapRight", () => setImageWrap(this.view, "right"));
    this.element.hidden = true;
    document.body.append(this.element);
    window.addEventListener("scroll", this.place, true);
    window.addEventListener("resize", this.place);
  }

  place = () => {
    const view = this.view;
    const sel = view.state.selection;
    if (!(sel instanceof NodeSelection) || !isImage(sel.node) || !view.hasFocus() && !this.element.contains(document.activeElement)) {
      this.element.hidden = true;
      return;
    }
    const dom = view.nodeDOM(sel.from) as HTMLElement | null;
    if (!dom) {
      this.element.hidden = true;
      return;
    }
    const $pos = view.state.doc.resolve(sel.from);
    const line = isImageLine($pos.parent) && !sel.node.attrs.wrap;
    const align = line ? (($pos.parent.attrs.align as string | null) ?? "left") : null;
    for (const [name, button] of this.buttons) {
      const on = name === align || (name === "wrapLeft" && sel.node.attrs.wrap === "left") || (name === "wrapRight" && sel.node.attrs.wrap === "right");
      button.classList.toggle("is-on", on);
      if (name === "wrapLeft" || name === "wrapRight") button.disabled = wrapTargetFor(view.state, sel.from) === null;
    }
    this.element.hidden = false;
    const r = dom.getBoundingClientRect();
    const bar = this.element.getBoundingClientRect();
    const left = Math.max(4, Math.min(window.innerWidth - bar.width - 4, (r.left + r.right) / 2 - bar.width / 2));
    const top = r.top - bar.height - 8 < 4 ? r.bottom + 8 : r.top - bar.height - 8;
    this.element.style.left = `${left}px`;
    this.element.style.top = `${top}px`;
  };

  destroy() {
    window.removeEventListener("scroll", this.place, true);
    window.removeEventListener("resize", this.place);
    this.element.remove();
  }
}

export const imagesKey = new PluginKey<DecorationSet>("visualtex-document-images");

/** Image lines and the images on them (the gap between neighbours): node classes. */
function imageDecorations(doc: Node): DecorationSet {
  const decorations: Decoration[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== "paragraph") return node.isBlock;
    let hasImage = false;
    node.forEach((child) => {
      if (isImage(child)) hasImage = true;
    });
    if (!hasImage) return false;
    if (isImageLine(node)) {
      decorations.push(Decoration.node(pos, pos + node.nodeSize, { class: "vtx-image-line" }));
      let previous = false;
      node.forEach((child, offset) => {
        if (!isImage(child)) return;
        if (previous) decorations.push(Decoration.node(pos + 1 + offset, pos + 1 + offset + child.nodeSize, { class: "has-gap" }));
        previous = true;
      });
    }
    return false;
  });
  return DecorationSet.create(doc, decorations);
}

/**
 * Keeps what LaTeX can do: a wrapped image is the first thing of a body paragraph with text (wrapfig), one per
 * paragraph; images stay out of tables, headings and the like (they move after that block).
 */
function normalize(state: EditorState): Transaction | null {
  const fixes: Array<{ pos: number; node: Node; kind: "unwrap" | "front" }> = [];
  // images where LaTeX cannot have them (table cells, headings, the title, the reference list): moved out
  const outside: Array<{ pos: number; node: Node; after: number }> = [];
  state.doc.descendants((node, pos, parent, index) => {
    if (!isImage(node)) return true;
    const $pos = state.doc.resolve(pos);
    let blocked = parent?.type.name !== "paragraph";
    for (let d = $pos.depth; d > 0; d--) {
      const name = $pos.node(d).type.name;
      if (name === "table" || name === "bibliography") blocked = true;
    }
    if (blocked) {
      outside.push({ pos, node, after: $pos.after(1) });
      return false;
    }
    if (!node.attrs.wrap) return false;
    // a wrapped image in its figure (figures.ts): the figure is placed beside its paragraph
    if ($pos.depth === 2 && $pos.node(1).type.name === "figure") return false;
    if (!parent || $pos.depth !== 1 || !parent.textContent.trim()) fixes.push({ pos, node, kind: "unwrap" });
    else if (index !== 0) {
      // a second wrapped image in the paragraph does not wrap
      const first = parent.firstChild;
      fixes.push({ pos, node, kind: isImage(first) && first!.attrs.wrap ? "unwrap" : "front" });
    }
    return false;
  });
  if (!fixes.length && !outside.length) return null;
  const tr = state.tr;
  for (const fix of fixes.reverse()) {
    const pos = tr.mapping.map(fix.pos);
    if (fix.kind === "unwrap") tr.setNodeMarkup(pos, undefined, { ...fix.node.attrs, wrap: null });
    else {
      const start = tr.doc.resolve(pos).start();
      tr.delete(pos, pos + fix.node.nodeSize);
      tr.insert(start, fix.node);
    }
  }
  for (const out of outside.reverse()) {
    const pos = tr.mapping.map(out.pos);
    tr.delete(pos, pos + out.node.nodeSize);
    const after = tr.mapping.map(out.after);
    tr.insert(after, state.schema.nodes.paragraph.create({ align: "center" }, state.schema.nodes.image.create({ ...out.node.attrs, wrap: null })));
  }
  return tr;
}

export interface ImagesPluginOptions {
  labels: () => ImageLabels;
}

export function imagesPlugin(options: ImagesPluginOptions): Plugin<DecorationSet> {
  return new Plugin<DecorationSet>({
    key: imagesKey,
    state: {
      init: (_config, state) => imageDecorations(state.doc),
      apply: (tr, value, _old, state) => (tr.docChanged ? imageDecorations(state.doc) : value),
    },
    appendTransaction(trs, _old, state) {
      if (!trs.some((tr) => tr.docChanged)) return null;
      return normalize(state);
    },
    view(view) {
      const bar = new ImageBar(view, options.labels());
      const raf = () => requestAnimationFrame(bar.place);
      raf();
      return {
        update: raf,
        destroy: () => bar.destroy(),
      };
    },
    props: {
      decorations: (state) => imagesKey.getState(state),
      nodeViews: {
        image: (node, view, getPos) => new ImageView(node, view, getPos),
      },
      handleKeyDown(view, event) {
        const sel = view.state.selection;
        if (!(sel instanceof NodeSelection) || !isImage(sel.node)) return false;
        if (event.key === "Enter" && !event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey) {
          // a paragraph after the image's line
          const $pos = view.state.doc.resolve(sel.from);
          if (isImageLine($pos.parent) || sel.node.attrs.wrap) {
            // after its figure (figures.ts), not between the line and the caption
            const inFigure = $pos.depth > 1 && $pos.node($pos.depth - 1).type.name === "figure";
            const after = inFigure ? $pos.after($pos.depth - 1) : $pos.after();
            const tr = view.state.tr.insert(after, view.state.schema.nodes.paragraph.create());
            tr.setSelection(TextSelection.create(tr.doc, after + 1));
            view.dispatch(tr.scrollIntoView());
            return true;
          }
          view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, sel.to)));
          return false;
        }
        return false;
      },
      // copied images pasted: by the insertion rules (a line after a sentence, not a block-sized image in it)
      handlePaste(view, _event, slice) {
        const images: Node[] = [];
        let other = false;
        let board = false;
        slice.content.descendants((node) => {
          if (node.type.name === "image_board") board = true;
          if (node.type.name === "board_image") images.push(node.type.schema.nodes.image.create({ src: node.attrs.src, alt: node.attrs.alt, width: node.attrs.width }));
          if (isImage(node)) images.push(node.type.create({ ...node.attrs, wrap: null, id: null }));
          else if (node.isText ? !!node.text?.trim() : node.isLeaf || (node.isTextblock && false)) other = true;
          return !isImage(node);
        });
        // a whole board pastes as a board
        if (!images.length || other || board) return false;
        const tr = insertImageNodes(view.state, images, view.state.schema);
        if (!tr) return false;
        view.dispatch(tr.scrollIntoView());
        return true;
      },
      // typing does not replace a selected image (Mathcha)
      handleTextInput(view) {
        const sel = view.state.selection;
        return sel instanceof NodeSelection && isImage(sel.node);
      },
    },
  });
}

/** Image nodes in a slice that are not allowed where it is pasted (tables…) are dropped. */
export function sliceImages(content: Fragment): number {
  let count = 0;
  content.descendants((node) => {
    if (isImage(node) || node.type.name === "board_image") count++;
    return true;
  });
  return count;
}
