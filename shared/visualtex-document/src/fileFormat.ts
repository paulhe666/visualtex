// The .vtxdoc file: JSON with the ProseMirror document (DESIGN.md §2).
import type { Node, Schema } from "prosemirror-model";
import { documentImagesForFile, loadDocumentImages } from "./images.ts";
import { withoutMarkdownSyntax } from "./markdownLive.ts";

export const DOCUMENT_FILE_FORMAT = "visualtex-document";
/** 2: the file carries images (`images`); 3: it has image boards; 4: figures and tables with captions (figures.ts).
 *  A document without them is still written as 1. */
export const DOCUMENT_FILE_VERSION = 4;

export interface DocumentFile {
  format: typeof DOCUMENT_FILE_FORMAT;
  version: number;
  doc: unknown;
  meta?: { title?: string; modified?: string };
  /** the images the document shows: key (vtx-image:<sha-256>.png|jpg) -> base64 (DESIGN.md §11.2) */
  images?: Record<string, string>;
}

/** Whether the document has a node of this type. */
function hasNode(doc: Node, name: string): boolean {
  let found = false;
  doc.descendants((node) => {
    if (node.type.name === name) found = true;
    return !found;
  });
  return found;
}

/** Figures or tables with captions: an older VisualTeX would not know them (it refuses the file rather than fail). */
const hasCaptions = (doc: Node) => hasNode(doc, "figure") || hasNode(doc, "table_figure");

export function serializeDocument(doc: Node, meta: DocumentFile["meta"] = {}): string {
  // the Markdown shown around the caret is not part of the document
  doc = withoutMarkdownSyntax(doc);
  const images = documentImagesForFile(doc);
  const hasImages = Object.keys(images).length > 0;
  const file: DocumentFile = {
    format: DOCUMENT_FILE_FORMAT,
    version: hasCaptions(doc) ? 4 : hasNode(doc, "image_board") ? 3 : hasImages ? 2 : 1,
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
  const doc = schema.nodeFromJSON(file.doc);
  doc.check();
  return { doc, meta: file.meta ?? {} };
}
