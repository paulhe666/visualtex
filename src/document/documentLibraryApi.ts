// The document library (document mode's sidebar): documents kept by the app, in folders, with recently deleted ones.
// In the app the files live in the app's data folder (src-tauri/src/document_library.rs); in a browser (the document
// lab) the same calls are kept in localStorage.
import { invoke, isTauri } from "@tauri-apps/api/core";
import type { Node } from "prosemirror-model";
import { create } from "zustand";
import { createJSONStorage, persist } from "zustand/middleware";
import { safeStorage } from "../runtime/safeStorage";
import { renderedText } from "../../shared/visualtex-document/src/index.ts";

export interface LibraryDocument {
  id: string;
  /** The user's name for it, or (until renamed) its first heading or line; "" while it has no text. */
  title: string;
  titleSet: boolean;
  folder: string | null;
  created: number;
  modified: number;
  trashed: number | null;
}

export interface LibraryFolder {
  id: string;
  name: string;
  created: number;
}

export interface Library {
  /** The folder holding the files (the app only). */
  root: string;
  folders: LibraryFolder[];
  /** Newest first. */
  documents: LibraryDocument[];
}

export interface DocumentLibraryApi {
  list(): Promise<Library>;
  read(id: string): Promise<string>;
  create(folder: string | null, text: string, title?: string | null): Promise<LibraryDocument>;
  save(id: string, text: string, title: string): Promise<LibraryDocument>;
  rename(id: string, title: string): Promise<Library>;
  duplicate(id: string, title: string): Promise<LibraryDocument>;
  move(id: string, folder: string | null): Promise<Library>;
  trash(id: string): Promise<Library>;
  restore(id: string): Promise<Library>;
  remove(id: string): Promise<Library>;
  emptyTrash(): Promise<Library>;
  createFolder(name: string): Promise<{ library: Library; id: string }>;
  renameFolder(id: string, name: string): Promise<Library>;
  deleteFolder(id: string): Promise<Library>;
  /** Ids of the documents whose title or text has every word of the query. */
  search(query: string): Promise<string[]>;
}

const tauriLibrary: DocumentLibraryApi = {
  list: () => invoke("document_library_list"),
  read: (id) => invoke("document_library_read", { id }),
  create: (folder, text, title) => invoke("document_library_create", { folder, text, title: title ?? null }),
  save: (id, text, title) => invoke("document_library_save", { id, text, title }),
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
};

// ---- the browser's copy (the document lab): the same behaviour as document_library.rs ----

const LOCAL_INDEX = "visualtex.document.library.v1";
const LOCAL_DOC = "visualtex.document.library.doc.";
const TRASH_KEEP_MS = 30 * 24 * 60 * 60 * 1000;

function createLocalLibrary(storage: Storage): DocumentLibraryApi {
  type Index = { folders: LibraryFolder[]; documents: LibraryDocument[] };
  const load = (): Index => {
    let index: Index = { folders: [], documents: [] };
    try {
      index = { folders: [], documents: [], ...(JSON.parse(storage.getItem(LOCAL_INDEX) ?? "{}") as Partial<Index>) };
    } catch {
      // an unreadable index starts again
    }
    const cutoff = Date.now() - TRASH_KEEP_MS;
    index.documents = index.documents.filter((doc) => {
      const keep = storage.getItem(LOCAL_DOC + doc.id) !== null && !(doc.trashed !== null && doc.trashed < cutoff);
      if (!keep) storage.removeItem(LOCAL_DOC + doc.id);
      return keep;
    });
    for (const doc of index.documents) if (doc.folder && !index.folders.some((f) => f.id === doc.folder)) doc.folder = null;
    return index;
  };
  const store = (index: Index) => storage.setItem(LOCAL_INDEX, JSON.stringify(index));
  const view = (index: Index): Library => ({
    root: "",
    folders: index.folders,
    documents: [...index.documents].sort((a, b) => b.modified - a.modified),
  });
  const newId = () => `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 10)}`.padEnd(16, "0");
  const clean = (name: string) => name.replace(/[\u0000-\u001f\u007f]/g, "").trim().slice(0, 200);
  const find = (index: Index, id: string) => {
    const doc = index.documents.find((d) => d.id === id);
    if (!doc) throw new Error("The document is not in the library");
    return doc;
  };
  const change = async (fn: (index: Index) => void) => {
    const index = load();
    fn(index);
    store(index);
    return view(index);
  };
  // a modification time strictly after the last one, so the newest is first even within a millisecond
  let lastTime = 0;
  const now = () => (lastTime = Math.max(Date.now(), lastTime + 1));
  const api: DocumentLibraryApi = {
    list: async () => {
      const index = load();
      store(index);
      return view(index);
    },
    read: async (id) => {
      const text = storage.getItem(LOCAL_DOC + id);
      if (text === null) throw new Error("The document is not in the library");
      return text;
    },
    create: async (folder, text, title) => {
      const index = load();
      const time = now();
      const name = clean(title ?? "");
      const doc: LibraryDocument = {
        id: newId(),
        title: name,
        titleSet: name !== "",
        folder: folder && index.folders.some((f) => f.id === folder) ? folder : null,
        created: time,
        modified: time,
        trashed: null,
      };
      storage.setItem(LOCAL_DOC + doc.id, text);
      index.documents.push(doc);
      store(index);
      return { ...doc };
    },
    save: async (id, text, title) => {
      const index = load();
      const doc = find(index, id);
      storage.setItem(LOCAL_DOC + id, text);
      doc.modified = now();
      if (!doc.titleSet) doc.title = clean(title);
      store(index);
      return { ...doc };
    },
    rename: (id, title) =>
      change((index) => {
        const doc = find(index, id);
        doc.title = clean(title);
        doc.titleSet = doc.title !== "";
      }),
    duplicate: async (id, title) => {
      const text = await api.read(id);
      return api.create(find(load(), id).folder, text, title);
    },
    move: (id, folder) =>
      change((index) => {
        find(index, id).folder = folder && index.folders.some((f) => f.id === folder) ? folder : null;
      }),
    trash: (id) =>
      change((index) => {
        find(index, id).trashed = now();
      }),
    restore: (id) =>
      change((index) => {
        find(index, id).trashed = null;
      }),
    remove: (id) =>
      change((index) => {
        index.documents = index.documents.filter((d) => d.id !== id);
        storage.removeItem(LOCAL_DOC + id);
      }),
    emptyTrash: () =>
      change((index) => {
        for (const doc of index.documents) if (doc.trashed !== null) storage.removeItem(LOCAL_DOC + doc.id);
        index.documents = index.documents.filter((d) => d.trashed === null);
      }),
    createFolder: async (name) => {
      const clean_ = clean(name);
      if (!clean_) throw new Error("The folder needs a name");
      const id = newId();
      const library = await change((index) => {
        index.folders.push({ id, name: clean_, created: now() });
      });
      return { library, id };
    },
    renameFolder: (id, name) =>
      change((index) => {
        const folder = index.folders.find((f) => f.id === id);
        if (!folder || !clean(name)) throw new Error("The folder needs a name");
        folder.name = clean(name);
      }),
    deleteFolder: (id) =>
      change((index) => {
        index.folders = index.folders.filter((f) => f.id !== id);
        for (const doc of index.documents) if (doc.folder === id) doc.folder = null;
      }),
    search: async (query) => {
      const words = query.toLowerCase().split(/\s+/).filter(Boolean);
      if (!words.length) return [];
      return load()
        .documents.filter((doc) => doc.trashed === null)
        .filter((doc) => {
          const text = `${doc.title}\n${searchableText(storage.getItem(LOCAL_DOC + doc.id) ?? "")}`.toLowerCase();
          return words.every((word) => text.includes(word));
        })
        .map((doc) => doc.id);
    },
  };
  return api;
}

function searchableText(source: string): string {
  const out: string[] = [];
  const walk = (value: unknown) => {
    if (Array.isArray(value)) value.forEach(walk);
    else if (value && typeof value === "object")
      for (const [key, item] of Object.entries(value)) {
        if (typeof item === "string" && ["text", "latex", "alt", "title", "href"].includes(key)) out.push(item);
        else walk(item);
      }
  };
  try {
    walk(JSON.parse(source));
  } catch {
    return source;
  }
  return out.join("\n");
}

let library: DocumentLibraryApi | null = null;

export function documentLibrary(): DocumentLibraryApi {
  library ??= isTauri() ? tauriLibrary : createLocalLibrary(window.localStorage);
  return library;
}

/** The name a document takes from its text: its title, else its first heading, else its first line with text. */
export function documentTitle(doc: Node): string {
  let title = "";
  let heading = "";
  let line = "";
  doc.descendants((node) => {
    if (title) return false;
    if (node.isTextblock) {
      // (as it reads: a heading's "#" and the markers left out)
      const text = renderedText(node).replace(/\s+/g, " ").trim();
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
  return doc.title || (isEn ? "Untitled" : "未命名文档");
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

function toBase64(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

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
    const dataBase64 = btoa(binary);
    const windows = /Windows/i.test(navigator.userAgent);
    await invoke("write_export_file", windows ? { request: { path, text: "", dataBase64 } } : { path, dataBase64 });
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
    await invoke("write_export_file", { path, dataBase64: toBase64(text) });
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

const IMPORT_EXTENSIONS = ["vtxdoc", "json", "md", "markdown", "zip"];

/**
 * A file to open, from the system's open dialog (the app only; null when cancelled): its bytes and the folder it is
 * in (a Markdown file's pictures are found from there).
 */
export async function chooseImportFile(isEn: boolean): Promise<{ file: File; folder: string } | null> {
  if (!isTauri()) return null;
  const { open } = await import("@tauri-apps/plugin-dialog");
  const chosen = await open({
    multiple: false,
    directory: false,
    filters: [{ name: isEn ? "Documents" : "文档", extensions: IMPORT_EXTENSIONS }],
    title: isEn ? "Open" : "打开",
  });
  if (!chosen || Array.isArray(chosen)) return null;
  const bytes = fromBase64(await invoke<string>("read_import_file", { path: chosen }));
  const name = chosen.split(/[\\/]/).pop() ?? chosen;
  return { file: new File([bytes as BlobPart], name), folder: chosen.slice(0, chosen.length - name.length) };
}

function fromBase64(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
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
