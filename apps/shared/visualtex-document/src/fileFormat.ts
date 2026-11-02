// The .vtxdoc file: JSON with the ProseMirror document (DESIGN.md §2).
import type { Node, Schema } from "prosemirror-model";

export const DOCUMENT_FILE_FORMAT = "visualtex-document";
export const DOCUMENT_FILE_VERSION = 1;

export interface DocumentFile {
  format: typeof DOCUMENT_FILE_FORMAT;
  version: number;
  doc: unknown;
  meta?: { title?: string; modified?: string };
}

export function serializeDocument(doc: Node, meta: DocumentFile["meta"] = {}): string {
  const file: DocumentFile = {
    format: DOCUMENT_FILE_FORMAT,
    version: DOCUMENT_FILE_VERSION,
    doc: doc.toJSON(),
    meta: { ...meta, modified: new Date().toISOString() },
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
  const doc = schema.nodeFromJSON(file.doc);
  doc.check();
  return { doc, meta: file.meta ?? {} };
}
