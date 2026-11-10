// The .vtxdoc file: JSON with the ProseMirror document (DESIGN.md §2).
import type { Node, Schema } from "prosemirror-model";
import { documentImagesForFile, loadDocumentImages } from "./images.ts";
import { withoutMarkdownSyntax } from "./markdownLive.ts";
import { separateParagraphs, toSourceDocument } from "./sourceMarkdown.ts";

export const DOCUMENT_FILE_FORMAT = "visualtex-document";
/** 2: the file carries images (`images`); 3: it has image boards; 4: figures and tables with captions (figures.ts);
 *  5: the text is its Markdown source (the source editor, sourceMarkdown.ts: "**" and a heading's "#" are text);
 *  6: the document is lines of Markdown (lists, quotes, code fences and rules are lines of text, as in Obsidian).
 *  Read, an older file becomes lines (toSourceDocument). */
export const DOCUMENT_FILE_VERSION = 6;

export interface DocumentFile {
  format: typeof DOCUMENT_FILE_FORMAT;
  version: number;
  doc: unknown;
  meta?: { title?: string; modified?: string };
  /** the images the document shows: key (vtx-image:<sha-256>.png|jpg) -> base64 (DESIGN.md §11.2) */
  images?: Record<string, string>;
}

export function serializeDocument(doc: Node, meta: DocumentFile["meta"] = {}): string {
  // the Markdown shown around the caret is not part of the document
  doc = withoutMarkdownSyntax(doc);
  const images = documentImagesForFile(doc);
  const hasImages = Object.keys(images).length > 0;
  const file: DocumentFile = {
    format: DOCUMENT_FILE_FORMAT,
    version: DOCUMENT_FILE_VERSION,
    doc: doc.toJSON(),
    meta: { ...meta, modified: new Date().toISOString() },
    ...(hasImages ? { images } : {}),
  };
  return JSON.stringify(file, null, 1);
}

/** Parse a .vtxdoc file; throws with a readable message when it is not one this version understands. */
export function parseDocument(schema: Schema, text: string): { doc: Node; meta: DocumentFile["meta"] } {
  const file = JSON.parse(text) as Partial<DocumentFile>;
  if (file.format !== DOCUMENT_FILE_FORMAT) throw new Error("Not a VisualTeX document");
  if (typeof file.version !== "number" || file.version > DOCUMENT_FILE_VERSION) {
    throw new Error(`Document version ${file.version} is newer than this VisualTeX supports`);
  }
  loadDocumentImages(file.images);
  let doc = schema.nodeFromJSON(file.doc);
  doc.check();
  // written before the source editor: its formatting into its text ("**", a heading's "#")
  // before 6 a document was paragraphs and blocks: as source lines (a blank line between paragraphs; lists, quotes,
  // code blocks, rules their lines of text)
  if (file.version < 6) doc = toSourceDocument(separateParagraphs(doc));
  return { doc, meta: file.meta ?? {} };
}
