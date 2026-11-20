// A document between the editor and its file (document mode): a Markdown file read with its memory and the pictures
// it names, and the document written back as the file's bytes, its memory and the pictures that have no file yet. A
// .vtxdoc of an earlier version is read and written in its own format.
import type { Node } from "prosemirror-model";
import {
  decodeFileText,
  documentSchema,
  encodeFileText,
  imageEntry,
  isEmptyMemory,
  newPicturePaths,
  parseDocument,
  parseDocumentMemory,
  picturePath,
  readDocumentFile,
  sameDocument,
  serializeDocument,
  serializeDocumentMemory,
  withImageFiles,
  writeDocumentFile,
  type DocumentFileSource,
  type FileEncoding,
} from "../../../visualtex-document/src/index.ts";
import type { DocumentContent, DocumentKind, LibraryDocument, OpenedDocument } from "./documentLibraryApi";
import { loadImportImages } from "./importImages";

/** A document open from a file: what its next save needs. */
export interface DocumentFile {
  entry: LibraryDocument;
  kind: DocumentKind;
  /** The document as read. */
  doc: Node;
  /** The Markdown file as read (null: a .vtxdoc, or a document that has no file yet). */
  source: DocumentFileSource | null;
  encoding: FileEncoding;
  /** What the file was when read or last saved: the next save is checked against it. */
  fingerprint: string | null;
  /** The pictures' files: image key -> path as the file names it. */
  imagePaths: Map<string, string>;
  /** Picture paths in use beside the file (other documents' pictures): a new picture takes another name. */
  taken: string[];
  /** Pictures given a path whose file is not written yet. */
  unwritten: Set<string>;
  /** Memory entries that found no place in the file (lines another program rewrote). */
  dropped: number;
}

export interface FileWriteOptions {
  /** A formula's LaTeX as it leaves the app. */
  formulaLatex?: (latex: string) => string;
  /** The names the document's file gave its formulas (a document read from a file, written to another). */
  labels?: Iterable<string>;
}

const utf8 = (text: string) => new TextEncoder().encode(text);

/** A file's text when it is one of the LaTeX documents of earlier versions (a pointer to a TeX project). */
export function isOldLatexPointer(text: string): boolean {
  if (!text.includes('"visualtex-latex-project"')) return false;
  try {
    return (JSON.parse(text) as { format?: unknown }).format === "visualtex-latex-project";
  } catch {
    return false;
  }
}

/** The folder a file is in, ready to put a relative path after it. */
function folderPrefix(dir: string): string | null {
  if (!dir) return null;
  return /[\\/]$/.test(dir) ? dir : dir + (dir.includes("\\") && !dir.includes("/") ? "\\" : "/");
}

/** A document read from the library: its text, memory and pictures as the editor's document. */
export async function loadDocumentFile(opened: OpenedDocument): Promise<DocumentFile> {
  const base = { entry: opened.entry, kind: opened.entry.kind, fingerprint: opened.fingerprint, taken: opened.assets.map((name) => `assets/${name}`), unwritten: new Set<string>() };
  if (opened.entry.kind === "vtxdoc") {
    const text = new TextDecoder("utf-8").decode(opened.bytes);
    if (isOldLatexPointer(text)) throw new Error("LaTeX documents are no longer supported");
    return { ...base, doc: parseDocument(documentSchema, text).doc, source: null, encoding: "utf-8", imagePaths: new Map(), dropped: 0 };
  }
  const { text, encoding } = decodeFileText(opened.bytes);
  const memory = parseDocumentMemory(opened.memory);
  // the pictures the text and the memory name: from the file's folder (the app), or kept with it (a browser)
  const pictures = await loadImportImages(
    text,
    "markdown",
    { archive: opened.pictures, base: "", folder: opened.pictures ? null : folderPrefix(opened.dir), folders: [] },
    Object.values(memory?.images ?? {}),
  );
  const read = withImageFiles((path) => pictures.get(path) ?? null, () => readDocumentFile(documentSchema, text, memory));
  return { ...base, doc: read.doc, source: read.source, encoding, imagePaths: new Map(read.source.imagePaths), dropped: read.dropped };
}

/** A picture path that is a file beside the document (not an address, not an absolute path, not above its folder). */
function besideDocument(path: string): boolean {
  const normal = picturePath(path).replace(/\\/g, "/");
  return !!normal && !/^(?:[a-zA-Z][a-zA-Z0-9+.-]*:|\/)/.test(normal) && !normal.split("/").includes("..");
}

/**
 * The document as its file's content. Pictures without a file get a path in assets/ (kept in `file`, so they keep
 * it); their bytes go along until a save has written them.
 */
export function documentContent(file: DocumentFile, doc: Node, options: FileWriteOptions = {}): DocumentContent {
  if (file.kind === "vtxdoc") return { bytes: utf8(serializeDocument(doc)), memory: null, assets: [] };
  for (const [key, path] of newPicturePaths(doc, file.imagePaths, file.taken)) {
    file.imagePaths.set(key, path);
    file.unwritten.add(key);
  }
  const written = writeDocumentFile(doc, { source: file.source, formulaLatex: options.formulaLatex, imagePaths: file.imagePaths });
  const encoded = encodeFileText(written.text, file.encoding);
  file.encoding = encoded.encoding;
  const assets: DocumentContent["assets"] = [];
  for (const key of file.unwritten) {
    const path = written.images.get(key);
    const bytes = imageEntry(key)?.bytes;
    if (path && bytes) assets.push({ path: picturePath(path), bytes });
  }
  return { bytes: encoded.bytes, memory: isEmptyMemory(written.memory) ? null : serializeDocumentMemory(written.memory), assets };
}

/**
 * After a save refused because a file of a new picture's name (`taken`) is beside the document now — another program
 * put it there since the document was opened: the new pictures take other names at the next save. Whether there were
 * new pictures to name again.
 */
export function renameNewPictures(file: DocumentFile, taken: string): boolean {
  if (!file.unwritten.size) return false;
  file.taken.push(taken);
  for (const key of file.unwritten) file.imagePaths.delete(key);
  file.unwritten.clear();
  return true;
}

/**
 * A document as a new Markdown file's content (a new document, Save as, a conversion): written without a file to
 * keep to, every picture that is a file beside the document going along. `keep`: the paths its pictures have. A
 * picture without one is named after the file it came from or the time it was pasted — or, with `byContent`, after
 * its bytes (many documents converted into one folder: two pictures never share a name).
 */
export function newDocumentContent(
  doc: Node,
  options: FileWriteOptions & { byContent?: boolean } = {},
  keep: ReadonlyMap<string, string> = new Map(),
  taken: string[] = [],
): { content: DocumentContent; text: string; imagePaths: Map<string, string> } {
  const imagePaths = new Map([...keep].filter(([, path]) => besideDocument(path)));
  if (!options.byContent) for (const [key, path] of newPicturePaths(doc, imagePaths, taken)) imagePaths.set(key, path);
  const written = writeDocumentFile(doc, { formulaLatex: options.formulaLatex, imagePaths, labels: options.labels });
  const assets: DocumentContent["assets"] = [];
  for (const [key, path] of written.images) {
    const bytes = imageEntry(key)?.bytes;
    if (bytes && besideDocument(path)) assets.push({ path: picturePath(path), bytes });
  }
  return { content: { bytes: utf8(written.text), memory: isEmptyMemory(written.memory) ? null : serializeDocumentMemory(written.memory), assets }, text: written.text, imagePaths: written.images };
}

/**
 * Whether a document written as a Markdown file and its memory reads back as the same document (the memory through
 * its stored text, the pictures by their paths).
 */
export function readsBackTheSame(doc: Node, content: DocumentContent, imagePaths: ReadonlyMap<string, string>): boolean {
  const byPath = new Map([...imagePaths].map(([key, path]) => [picturePath(path), key]));
  const read = withImageFiles(
    (path) => byPath.get(path) ?? null,
    () => readDocumentFile(documentSchema, decodeFileText(content.bytes).text, parseDocumentMemory(content.memory)),
  );
  return read.dropped === 0 && sameDocument(read.doc, doc);
}
