// The .vtxdoc file: JSON with the ProseMirror document (DESIGN.md §2).
import type { Node, Schema } from "prosemirror-model";
import { documentImagesForFile, loadDocumentImages } from "./images.ts";

export const DOCUMENT_FILE_FORMAT = "visualtex-document";
/** 2: the file carries images (`images`); 3: it has image boards. A document without them is still written as 1. */
export const DOCUMENT_FILE_VERSION = 3;

export interface DocumentFile {
  format: typeof DOCUMENT_FILE_FORMAT;
  version: number;
  doc: unknown;
  meta?: { title?: string; modified?: string };
  /** the images the document shows: key (vtx-image:<sha-256>.png|jpg) -> base64 (DESIGN.md §11.2) */
  images?: Record<string, string>;
}

function hasBoards(doc: Node): boolean {
  let found = false;
  doc.descendants((node) => {
    if (node.type.name === "image_board") found = true;
    return !found;
  });
  return found;
}

export function serializeDocument(doc: Node, meta: DocumentFile["meta"] = {}): string {
  const images = documentImagesForFile(doc);
  const hasImages = Object.keys(images).length > 0;
  const file: DocumentFile = {
    format: DOCUMENT_FILE_FORMAT,
    version: hasBoards(doc) ? 3 : hasImages ? 2 : 1,
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
