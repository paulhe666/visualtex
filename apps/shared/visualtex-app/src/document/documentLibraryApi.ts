// The document library (document mode's sidebar): the documents are Markdown files in the documents folder, in its
// subfolders, with recently deleted ones; files elsewhere are edited where they are (src-tauri document_files.rs).
// What VisualTeX remembers about a document besides its text (its memory) is kept beside the files by the library.
// In a browser (the document lab) the same calls are kept in localStorage.
import { invoke, isTauri } from "@tauri-apps/api/core";
import type { Node } from "prosemirror-model";
import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import { safeStorage } from "../runtime/safeStorage";
import { writeExportFile } from "@platform/exportFile";
import { renderedText } from "../../../visualtex-document/src/index.ts";

export type DocumentKind = "markdown" | "vtxdoc";

export interface LibraryDocument {
  /** Its path: relative to the documents folder ("Notes/Proof.md"), or "file:" and the path of a file elsewhere. */
  id: string;
  /** The file's name without its extension. */
  title: string;
  /** The subfolder of the documents folder it is in. */
  folder: string | null;
  created: number;
  modified: number;
  trashed: number | null;
  /** The file (the app only). */
  path: string;
  /** "vtxdoc": a document of an earlier version, kept in its own format. */
  kind: DocumentKind;
  /** A file outside the documents folder, edited where it is. */
  outside: boolean;
}

export interface LibraryFolder {
  /** Its path relative to the documents folder. */
  id: string;
  name: string;
  created: number;
}

export interface Library {
  /** The documents folder (the app only). */
  root: string;
  folders: LibraryFolder[];
  /** Newest first. */
  documents: LibraryDocument[];
  /** Files opened from elsewhere, the last opened first. */
  recent: LibraryDocument[];
}

/** What a save writes: the file's bytes, the memory record (null: none needed) and the pictures that have no file yet. */
export interface DocumentContent {
  bytes: Uint8Array;
  memory: string | null;
  assets: Array<{ path: string; bytes: Uint8Array }>;
  /** the file's modification time to set (a converted document keeps the time of the one it came from) */
  modified?: number;
}

export interface OpenedDocument {
  entry: LibraryDocument;
  bytes: Uint8Array;
  memory: string | null;
  /** What the file was when read: a save is checked against it. */
  fingerprint: string;
  /** The folder the file is in (the app): its pictures' paths start there. */
  dir: string;
  /** The files in "assets" beside it. */
  assets: string[];
  /** The pictures kept with it (a browser, where there is no folder to read them from): path -> bytes. */
  pictures: Map<string, Uint8Array> | null;
}

export type SaveOutcome = { status: "saved"; entry: LibraryDocument; fingerprint: string } | { status: "changed"; fingerprint: string };

/** A document of the library earlier versions kept in the app's data folder. */
export interface LegacyDocument {
  id: string;
  title: string;
  folder: string | null;
  created: number;
  modified: number;
}

export interface DocumentLibraryApi {
  list(): Promise<Library>;
  read(id: string): Promise<OpenedDocument>;
  /** The file's state now. */
  check(id: string): Promise<{ exists: boolean; fingerprint: string }>;
  /** A new document in a folder of the documents folder, named `name` or the first free "name 2", "name 3"… */
  create(folder: string | null, name: string, kind: DocumentKind, content: DocumentContent): Promise<LibraryDocument>;
  /** Saves if the file is still what was read (`base`), or anyway with `force`. */
  save(id: string, content: DocumentContent, base: string | null, force?: boolean): Promise<SaveOutcome>;
  /** Writes the document to a file the user chose (the app only). */
  saveAs(path: string, content: DocumentContent): Promise<LibraryDocument>;
  /** A file anywhere on the disk as a document (the app only). */
  openFile(path: string): Promise<LibraryDocument>;
  /** A file elsewhere leaves the recent files. */
  forget(id: string): Promise<Library>;
  rename(id: string, title: string): Promise<{ library: Library; id: string }>;
  duplicate(id: string, title: string): Promise<LibraryDocument>;
  move(id: string, folder: string | null): Promise<{ library: Library; id: string }>;
  trash(id: string): Promise<Library>;
  restore(id: string): Promise<Library>;
  remove(id: string): Promise<Library>;
  emptyTrash(): Promise<Library>;
  createFolder(name: string): Promise<{ library: Library; id: string }>;
  renameFolder(id: string, name: string): Promise<Library>;
  deleteFolder(id: string): Promise<Library>;
  /** Ids of the documents whose name or text has every word of the query. */
  search(query: string): Promise<string[]>;
  /** Another documents folder (null: the default one). */
  setRoot(path: string | null): Promise<Library>;
  /** The library of earlier versions, converted once (documentMigration.ts). */
  legacy: {
    list(): Promise<LegacyDocument[]>;
    read(id: string): Promise<string>;
    state(): Promise<unknown>;
    setState(state: unknown): Promise<void>;
  };
}

export function toBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

export function fromBase64(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

const wire = (content: DocumentContent) => ({
  base64: toBase64(content.bytes),
  memory: content.memory,
  assets: content.assets.map((asset) => ({ path: asset.path, base64: toBase64(asset.bytes) })),
  modified: content.modified ?? null,
});

const tauriLibrary: DocumentLibraryApi = {
  list: () => invoke("document_library_list"),
  read: async (id) => {
    const opened = await invoke<Omit<OpenedDocument, "bytes" | "pictures"> & { base64: string }>("document_library_read", { id });
    return { entry: opened.entry, bytes: fromBase64(opened.base64), memory: opened.memory, fingerprint: opened.fingerprint, dir: opened.dir, assets: opened.assets, pictures: null };
  },
  check: (id) => invoke("document_library_check", { id }),
  create: (folder, name, kind, content) => invoke("document_library_create", { folder, name, extension: kind === "vtxdoc" ? "vtxdoc" : "md", content: wire(content) }),
  save: (id, content, base, force = false) => invoke("document_library_save", { id, content: wire(content), base, force }),
  saveAs: (path, content) => invoke("document_library_save_as", { path, content: wire(content) }),
  openFile: (path) => invoke("document_library_open_file", { path }),
  forget: (id) => invoke("document_library_forget", { id }),
  rename: (id, title) => invoke("document_library_rename", { id, title }),
  duplicate: (id, title) => invoke("document_library_duplicate", { id, title }),
  move: (id, folder) => invoke("document_library_move", { id, folder }),
  trash: (id) => invoke("document_library_trash", { id }),
  restore: (id) => invoke("document_library_restore", { id }),
  remove: (id) => invoke("document_library_delete", { id }),
  emptyTrash: () => invoke("document_library_empty_trash"),
  createFolder: (name) => invoke("document_library_create_folder", { name }),
  renameFolder: (id, name) => invoke("document_library_rename_folder", { id, name }),
  deleteFolder: (id) => invoke("document_library_delete_folder", { id }),
  search: (query) => invoke("document_library_search", { query }),
  setRoot: (path) => invoke("document_library_set_root", { path }),
  legacy: {
    list: () => invoke("document_library_legacy_list"),
    read: (id) => invoke("document_library_legacy_read", { id }),
    state: () => invoke("document_library_migration"),
    setState: (state) => invoke("document_library_set_migration", { state }),
  },
};

// ---- the browser's copy (the document lab): the same behaviour as document_files.rs, in a Storage ----

const LOCAL_FILES = "visualtex.document.files.v1";
const LOCAL_FILE = "visualtex.document.file.";
const LOCAL_MIGRATION = "visualtex.document.files.migration";
// the library before documents were files: its documents are converted once
const OLD_INDEX = "visualtex.document.library.v1";
const OLD_DOC = "visualtex.document.library.doc.";
const TRASH_KEEP_MS = 30 * 24 * 60 * 60 * 1000;

/** A name as a file name (document_files.rs file_name). */
function fileName(name: string): string {
  return name
    .replace(/[\u0000-\u001f\u007f\\/:*?"<>|]/g, "-")
    .trim()
    .replace(/[. ]+$/, "")
    .replace(/^\.+/, "")
    .slice(0, 120)
    .trim();
}

/** SHA-256 of bytes as hex (what the app's library gives as a file's fingerprint). */
async function fingerprintOf(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", bytes as BufferSource);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

export function createLocalLibrary(storage: Storage): DocumentLibraryApi {
  type Stored = { base64: string; memory: string | null; assets: Record<string, string> };
  type Index = { folders: LibraryFolder[]; documents: Array<LibraryDocument & { was?: string }> };
  const load = (): Index => {
    let index: Index = { folders: [], documents: [] };
    try {
      index = { folders: [], documents: [], ...(JSON.parse(storage.getItem(LOCAL_FILES) ?? "{}") as Partial<Index>) };
    } catch {
      // an unreadable index starts again
    }
    const cutoff = Date.now() - TRASH_KEEP_MS;
    index.documents = index.documents.filter((doc) => {
      const keep = storage.getItem(LOCAL_FILE + doc.id) !== null && !(doc.trashed !== null && doc.trashed < cutoff);
      if (!keep) storage.removeItem(LOCAL_FILE + doc.id);
      return keep;
    });
    return index;
  };
  const store = (index: Index) => storage.setItem(LOCAL_FILES, JSON.stringify(index));
  const view = (index: Index): Library => ({
    root: "",
    folders: [...index.folders],
    documents: [...index.documents].sort((a, b) => b.modified - a.modified),
    recent: [],
  });
  const find = (index: Index, id: string) => {
    const doc = index.documents.find((d) => d.id === id);
    if (!doc) throw new Error("The document is not in the library");
    return doc;
  };
  const stored = (id: string): Stored => {
    const text = storage.getItem(LOCAL_FILE + id);
    if (text === null) throw new Error("The document is not in the library");
    return { assets: {}, ...(JSON.parse(text) as Partial<Stored>) } as Stored;
  };
  /** The assets of the documents in one folder are one "assets" folder, as on a disk. */
  const put = (id: string, content: DocumentContent, before: Stored | null) => {
    const assets = { ...(before?.assets ?? {}) };
    for (const asset of content.assets) assets[asset.path] ??= toBase64(asset.bytes);
    storage.setItem(LOCAL_FILE + id, JSON.stringify({ base64: toBase64(content.bytes), memory: content.memory, assets } satisfies Stored));
  };
  const freeId = (index: Index, folder: string | null, stem: string, extension: string) => {
    const at = (name: string) => (folder ? `${folder}/` : "") + `${name}.${extension}`;
    let id = at(stem);
    for (let n = 2; index.documents.some((doc) => doc.id.toLowerCase() === id.toLowerCase()); n++) id = at(`${stem} ${n}`);
    return id;
  };
  const extensionOf = (id: string) => (id.toLowerCase().endsWith(".vtxdoc") ? "vtxdoc" : "md");
  const rekey = (index: Index, doc: Index["documents"][number], id: string) => {
    const text = storage.getItem(LOCAL_FILE + doc.id);
    storage.removeItem(LOCAL_FILE + doc.id);
    if (text !== null) storage.setItem(LOCAL_FILE + id, text);
    doc.id = id;
    doc.title = id.slice(id.lastIndexOf("/") + 1).replace(/\.[^.]+$/, "");
    doc.folder = id.includes("/") && !id.startsWith(".trash/") ? id.slice(0, id.lastIndexOf("/")) : null;
    void index;
  };
  // a modification time strictly after the last one, so the newest is first even within a millisecond
  let lastTime = 0;
  const now = () => (lastTime = Math.max(Date.now(), lastTime + 1));
  const change = async (fn: (index: Index) => void) => {
    const index = load();
    fn(index);
    store(index);
    return view(index);
  };
  const api: DocumentLibraryApi = {
    list: async () => {
      const index = load();
      store(index);
      return view(index);
    },
    read: async (id) => {
      const entry = find(load(), id);
      const content = stored(id);
      const bytes = fromBase64(content.base64);
      return {
        entry: { ...entry },
        bytes,
        memory: content.memory,
        fingerprint: await fingerprintOf(bytes),
        dir: "",
        assets: Object.keys(content.assets).filter((path) => path.startsWith("assets/")).map((path) => path.slice(7)),
        pictures: new Map(Object.entries(content.assets).map(([path, base64]) => [path, fromBase64(base64)])),
      };
    },
    check: async (id) => {
      const text = storage.getItem(LOCAL_FILE + id);
      return text === null ? { exists: false, fingerprint: "" } : { exists: true, fingerprint: await fingerprintOf(fromBase64(stored(id).base64)) };
    },
    create: async (folder, name, kind, content) => {
      const index = load();
      const stem = fileName(name);
      if (!stem) throw new Error("The document needs a name");
      const time = content.modified ?? now();
      const place = folder && index.folders.some((f) => f.id === folder) ? folder : null;
      const id = freeId(index, place, stem, kind === "vtxdoc" ? "vtxdoc" : "md");
      const doc: LibraryDocument = { id, title: id.slice(id.lastIndexOf("/") + 1).replace(/\.[^.]+$/, ""), folder: place, created: time, modified: time, trashed: null, path: id, kind, outside: false };
      put(id, content, null);
      index.documents.push(doc);
      store(index);
      return { ...doc };
    },
    save: async (id, content, base, force = false) => {
      const index = load();
      const doc = find(index, id);
      const before = stored(id);
      const current = await fingerprintOf(fromBase64(before.base64));
      if (!force && base !== null && current !== base) return { status: "changed", fingerprint: current };
      put(id, content, before);
      doc.modified = now();
      store(index);
      return { status: "saved", entry: { ...doc }, fingerprint: await fingerprintOf(content.bytes) };
    },
    saveAs: async () => {
      throw new Error("Files are saved by the app only");
    },
    openFile: async () => {
      throw new Error("Files are opened by the app only");
    },
    forget: async () => view(load()),
    rename: async (id, title) => {
      const index = load();
      const doc = find(index, id);
      const stem = fileName(title);
      if (!stem) throw new Error("The document needs a name");
      if (stem !== doc.title) rekey(index, doc, stem.toLowerCase() === doc.title.toLowerCase() ? (doc.folder ? `${doc.folder}/` : "") + `${stem}.${extensionOf(id)}` : freeId(index, doc.folder, stem, extensionOf(id)));
      store(index);
      return { library: view(index), id: doc.id };
    },
    duplicate: async (id, title) => {
      const index = load();
      const doc = find(index, id);
      const content = stored(id);
      const stem = fileName(title);
      if (!stem) throw new Error("The document needs a name");
      const copy: LibraryDocument = { ...doc, id: freeId(index, doc.folder, stem, extensionOf(id)), created: now(), modified: now() };
      copy.title = copy.id.slice(copy.id.lastIndexOf("/") + 1).replace(/\.[^.]+$/, "");
      storage.setItem(LOCAL_FILE + copy.id, JSON.stringify(content));
      index.documents.push(copy);
      store(index);
      return { ...copy };
    },
    move: async (id, folder) => {
      const index = load();
      const doc = find(index, id);
      const place = folder && index.folders.some((f) => f.id === folder) ? folder : null;
      if (place !== doc.folder) rekey(index, doc, freeId(index, place, doc.title, extensionOf(id)));
      store(index);
      return { library: view(index), id: doc.id };
    },
    trash: (id) =>
      change((index) => {
        const doc = find(index, id);
        doc.was = doc.id;
        doc.trashed = now();
        rekey(index, doc, `.trash/${doc.trashed.toString(36)}/${doc.title}.${extensionOf(id)}`);
      }),
    restore: (id) =>
      change((index) => {
        const doc = find(index, id);
        const was = doc.was ?? `${doc.title}.${extensionOf(id)}`;
        const folder = was.includes("/") ? was.slice(0, was.lastIndexOf("/")) : null;
        doc.trashed = null;
        rekey(index, doc, freeId(index, folder && index.folders.some((f) => f.id === folder) ? folder : null, doc.title, extensionOf(id)));
        delete doc.was;
      }),
    remove: (id) =>
      change((index) => {
        index.documents = index.documents.filter((d) => d.id !== id);
        storage.removeItem(LOCAL_FILE + id);
      }),
    emptyTrash: () =>
      change((index) => {
        for (const doc of index.documents) if (doc.trashed !== null) storage.removeItem(LOCAL_FILE + doc.id);
        index.documents = index.documents.filter((d) => d.trashed === null);
      }),
    createFolder: async (name) => {
      const stem = fileName(name);
      if (!stem) throw new Error("The folder needs a name");
      let id = stem;
      const library = await change((index) => {
        for (let n = 2; index.folders.some((f) => f.id.toLowerCase() === id.toLowerCase()); n++) id = `${stem} ${n}`;
        index.folders.push({ id, name: id, created: now() });
      });
      return { library, id };
    },
    renameFolder: (id, name) =>
      change((index) => {
        const folder = index.folders.find((f) => f.id === id);
        const stem = fileName(name);
        if (!folder || !stem) throw new Error("The folder needs a name");
        if (index.folders.some((f) => f !== folder && f.id.toLowerCase() === stem.toLowerCase())) throw new Error("A folder with this name exists");
        for (const doc of index.documents) if (doc.folder === id) rekey(index, doc, `${stem}/${doc.id.slice(id.length + 1)}`);
        folder.id = stem;
        folder.name = stem;
      }),
    deleteFolder: (id) =>
      change((index) => {
        index.folders = index.folders.filter((f) => f.id !== id);
        for (const doc of index.documents) if (doc.folder === id) rekey(index, doc, freeId(index, null, doc.title, extensionOf(doc.id)));
      }),
    search: async (query) => {
      const words = query.toLowerCase().split(/\s+/).filter(Boolean);
      if (!words.length) return [];
      return load()
        .documents.filter((doc) => doc.trashed === null)
        .filter((doc) => {
          const text = `${doc.title}\n${new TextDecoder().decode(fromBase64(stored(doc.id).base64))}`.toLowerCase();
          return words.every((word) => text.includes(word));
        })
        .map((doc) => doc.id);
    },
    setRoot: async () => view(load()),
    legacy: {
      list: async () => {
        type Old = { folders?: Array<{ id: string; name: string }>; documents?: Array<{ id: string; title: string; folder: string | null; created: number; modified: number; trashed: number | null }> };
        let old: Old = {};
        try {
          old = JSON.parse(storage.getItem(OLD_INDEX) ?? "{}") as Old;
        } catch {
          // an unreadable index: nothing to convert
        }
        return (old.documents ?? [])
          .filter((doc) => doc.trashed === null && storage.getItem(OLD_DOC + doc.id) !== null)
          .map((doc) => ({ id: doc.id, title: doc.title, folder: old.folders?.find((f) => f.id === doc.folder)?.name ?? null, created: doc.created, modified: doc.modified }))
          .sort((a, b) => a.modified - b.modified);
      },
      read: async (id) => {
        const text = storage.getItem(OLD_DOC + id);
        if (text === null) throw new Error("The document is not in the library");
        return text;
      },
      state: async () => {
        try {
          return JSON.parse(storage.getItem(LOCAL_MIGRATION) ?? "null") as unknown;
        } catch {
          return null;
        }
      },
      setState: async (state) => storage.setItem(LOCAL_MIGRATION, JSON.stringify(state)),
    },
  };
  return api;
}

let library: DocumentLibraryApi | null = null;

export function documentLibrary(): DocumentLibraryApi {
  library ??= isTauri() ? tauriLibrary : createLocalLibrary(window.localStorage);
  return library;
}

/** A name for a document from its text: its title, else its first heading, else its first line with text. */
export function documentTitle(doc: Node): string {
  let title = "";
  let heading = "";
  let line = "";
  doc.descendants((node) => {
    if (title) return false;
    if (node.isTextblock) {
      // (as it reads: a heading's "#" and the markers left out)
      const read = renderedText(node).replace(/\s+/g, " ").trim();
      const text = node.type.name === "heading" ? read.replace(/^#{1,6}\s+/, "").replace(/\s+(?:#+|\{[^{}]*\})$/, "").trim() : read;
      if (text && node.type.name === "title") title = text;
      else if (text && node.type.name === "heading" && !heading) heading = text;
      else if (text && !line) line = text;
      return false;
    }
    return true;
  });
  return (title || heading || line).slice(0, 80);
}

export function libraryDocumentName(doc: Pick<LibraryDocument, "title">, isEn: boolean): string {
  return doc.title || untitledName(isEn);
}

/** A new document's name. */
export function untitledName(isEn: boolean): string {
  return isEn ? "Untitled" : "未命名文档";
}

/** The sidebar's open state and the document last open (both kept between sessions). */
interface DocumentLibraryState {
  open: boolean;
  lastId: string | null;
  collapsed: string[];
  setOpen: (open: boolean) => void;
  setLastId: (id: string | null) => void;
  toggleFolder: (id: string) => void;
}

export const useDocumentLibraryStore = create<DocumentLibraryState>()(
  persist(
    (set) => ({
      // shown at first when the window has room for it beside the page
      open: typeof window === "undefined" || window.innerWidth >= 1200,
      lastId: null,
      collapsed: [],
      setOpen: (open) => set({ open }),
      setLastId: (lastId) => set({ lastId }),
      toggleFolder: (id) =>
        set((state) => ({ collapsed: state.collapsed.includes(id) ? state.collapsed.filter((c) => c !== id) : [...state.collapsed, id] })),
    }),
    {
      name: "visualtex.document-library.v1",
      storage: createJSONStorage(() => safeStorage),
      partialize: (state) => ({ open: state.open, lastId: state.lastId, collapsed: state.collapsed }),
    },
  ),
);

/**
 * A binary file (an export .zip) to the disk: in the app through the system's save dialog (false when cancelled);
 * in a browser as a download.
 */
export async function saveBinaryFile(name: string, bytes: Uint8Array, extension: string, type: string, isEn: boolean): Promise<boolean> {
  if (isTauri()) {
    const { save } = await import("@tauri-apps/plugin-dialog");
    const chosen = await save({ defaultPath: name, filters: [{ name: extension.toUpperCase(), extensions: [extension] }], title: isEn ? "Save" : "保存" });
    if (!chosen) return false;
    const path = chosen.toLowerCase().endsWith(`.${extension}`) ? chosen : `${chosen}.${extension}`;
    let binary = "";
    for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    await writeExportFile(path, { base64: btoa(binary) });
    return true;
  }
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  const link = window.document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
  return true;
}

/** A path to save a file of `extension` to, from the system's save dialog (null when cancelled; the app only). */
export async function chooseExportPath(name: string, extension: string, isEn: boolean): Promise<string | null> {
  if (!isTauri()) return null;
  const { save } = await import("@tauri-apps/plugin-dialog");
  const chosen = await save({ defaultPath: name, filters: [{ name: extension.toUpperCase(), extensions: [extension] }], title: isEn ? "Save" : "保存" });
  if (!chosen) return null;
  return chosen.toLowerCase().endsWith(`.${extension}`) ? chosen : `${chosen}.${extension}`;
}

/**
 * A text file to the disk: in the app through the system's save dialog (false when cancelled); in a browser as a
 * download.
 */
export async function saveTextFile(name: string, text: string, extension: string, type: string, isEn: boolean): Promise<boolean> {
  if (isTauri()) {
    const { save } = await import("@tauri-apps/plugin-dialog");
    const chosen = await save({
      defaultPath: name,
      filters: [{ name: extension === "vtxdoc" ? "VisualTeX" : extension.toUpperCase(), extensions: [extension] }],
      title: isEn ? "Save" : "保存",
    });
    if (!chosen) return false;
    const path = chosen.toLowerCase().endsWith(`.${extension}`) ? chosen : `${chosen}.${extension}`;
    await writeExportFile(path, { text });
    return true;
  }
  const url = URL.createObjectURL(new Blob([text], { type }));
  const link = window.document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
  return true;
}

/** Documents open in place; a .zip (a Markdown export with its pictures) or a .json is taken into the library. */
const OPEN_EXTENSIONS = ["md", "markdown", "vtxdoc", "zip", "json"];

/** A file to open, from the system's open dialog (the app only; null when cancelled): its path. */
export async function chooseDocumentFile(isEn: boolean): Promise<string | null> {
  if (!isTauri()) return null;
  const { open } = await import("@tauri-apps/plugin-dialog");
  const chosen = await open({
    multiple: false,
    directory: false,
    filters: [{ name: isEn ? "Documents" : "文档", extensions: OPEN_EXTENSIONS }],
    title: isEn ? "Open" : "打开",
  });
  return !chosen || Array.isArray(chosen) ? null : chosen;
}

/** A file's bytes by its path (the app only). */
export async function readFileBytes(path: string): Promise<Uint8Array> {
  return fromBase64(await invoke<string>("read_import_file", { path }));
}

/** A folder from the system's dialog (the app only; null when cancelled). */
export async function chooseFolder(isEn: boolean, start?: string): Promise<string | null> {
  if (!isTauri()) return null;
  const { open } = await import("@tauri-apps/plugin-dialog");
  const chosen = await open({ multiple: false, directory: true, defaultPath: start || undefined, title: isEn ? "Documents folder" : "文档文件夹" });
  return !chosen || Array.isArray(chosen) ? null : chosen;
}

/** A picture file on the disk (an absolute path; the app only). */
export async function readImageFile(path: string): Promise<Uint8Array | null> {
  if (!isTauri()) return null;
  try {
    return fromBase64(await invoke<string>("read_image_file", { path }));
  } catch {
    return null;
  }
}

/** A picture on the web (http/https): through the app (no browser restrictions), or the browser's fetch. */
export async function fetchImage(url: string): Promise<Uint8Array | null> {
  try {
    if (isTauri()) return fromBase64(await invoke<string>("fetch_image_url", { url }));
    const response = await fetch(url);
    return response.ok ? new Uint8Array(await response.arrayBuffer()) : null;
  } catch {
    return null;
  }
}
