// The documents of earlier versions — `.vtxdoc` files in the library the app kept in its data folder — become Markdown
// files in the documents folder, once, at the first start of a version whose documents are files. Each one is
// written as Markdown and its memory, read back, and compared with the document it came from: only when the two
// are the same document is it converted; otherwise it is put into the documents folder as the `.vtxdoc` it is and
// reported. The earlier library stays where it is.
import {
  documentSchema,
  parseDocument,
} from "../../../visualtex-document/src/index.ts";
import { isOldLatexPointer, newDocumentContent, readsBackTheSame, type FileWriteOptions } from "./documentFiles";
import { documentTitle, type DocumentLibraryApi } from "./documentLibraryApi";

export interface MigrationState {
  version: 1;
  /** When every document of the earlier library had been dealt with (null: not yet). */
  finished: number | null;
  /** The documents dealt with, by their id in the earlier library: the document made (its id), and why it was kept
   *  as a .vtxdoc when it was. */
  documents: Record<string, { to: string; title: string; kept?: string }>;
  /** The documents kept as .vtxdoc were shown to the user. */
  reported: boolean;
}

function stateOf(value: unknown): MigrationState | null {
  const state = value as Partial<MigrationState> | null;
  if (!state || state.version !== 1 || typeof state.documents !== "object" || !state.documents) return null;
  return { version: 1, finished: typeof state.finished === "number" ? state.finished : null, documents: state.documents, reported: state.reported === true };
}

/** The documents the conversion kept as .vtxdoc. */
export function keptDocuments(state: MigrationState): Array<{ title: string; reason: string }> {
  return Object.values(state.documents)
    .filter((doc) => doc.kept !== undefined)
    .map((doc) => ({ title: doc.title, reason: doc.kept! }));
}

/**
 * Converts the earlier library's documents that are not converted yet (a run that was cut short goes on where it
 * stopped) and returns what was done. `differs`: the reason given for a document that does not read back the same.
 */
export async function migrateLegacyLibrary(api: DocumentLibraryApi, options: FileWriteOptions & { untitled: string; differs: string }): Promise<MigrationState> {
  const state: MigrationState = stateOf(await api.legacy.state()) ?? { version: 1, finished: null, documents: {}, reported: false };
  if (state.finished !== null) return state;
  const legacy = await api.legacy.list();
  const folders = new Map<string, string | null>();
  const folderFor = async (name: string | null): Promise<string | null> => {
    if (!name) return null;
    if (!folders.has(name)) {
      const existing = (await api.list()).folders.find((folder) => folder.id.toLowerCase() === name.trim().toLowerCase());
      folders.set(name, existing ? existing.id : await api.createFolder(name).then((made) => made.id, () => null));
    }
    return folders.get(name) ?? null;
  };
  // a document that could not even be copied is tried again at the next start
  let complete = true;
  for (const old of legacy) {
    if (state.documents[old.id]) continue;
    let text: string;
    try {
      text = await api.legacy.read(old.id);
    } catch {
      complete = false;
      continue;
    }
    // (the LaTeX documents of earlier versions were pointers to TeX projects: not documents any more)
    if (isOldLatexPointer(text)) continue;
    const folder = await folderFor(old.folder);
    let name = old.title.trim() || options.untitled;
    let kept: string | undefined;
    let to = "";
    try {
      const doc = parseDocument(documentSchema, text).doc;
      name = old.title.trim() || documentTitle(doc) || options.untitled;
      const made = newDocumentContent(doc, { formulaLatex: options.formulaLatex, byContent: true });
      if (!readsBackTheSame(doc, made.content, made.imagePaths)) throw new Error(options.differs);
      to = (await api.create(folder, name, "markdown", { ...made.content, modified: old.modified })).id;
    } catch (error) {
      kept = error instanceof Error ? error.message : String(error);
      try {
        to = (await api.create(folder, name, "vtxdoc", { bytes: new TextEncoder().encode(text), memory: null, assets: [], modified: old.modified })).id;
      } catch {
        complete = false;
        continue;
      }
    }
    state.documents[old.id] = { to, title: name, ...(kept === undefined ? {} : { kept }) };
    // (one kept in a run after the report was shown is reported too)
    if (kept !== undefined) state.reported = false;
    await api.legacy.setState(state);
  }
  if (complete) {
    state.finished = Date.now();
    if (!keptDocuments(state).length) state.reported = true;
  }
  await api.legacy.setState(state);
  return state;
}
