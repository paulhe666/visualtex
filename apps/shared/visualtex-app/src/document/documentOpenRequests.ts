// Documents the system hands to the app — a double click or "Open with" on a .md file, a file dropped on the window
// — reach document mode from here: the main window switches to document mode and the workspace opens them in place.
import { invoke, isTauri } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { useDocumentModeStore } from "../stores/documentModeStore";

/** Dispatched on the window when paths are waiting (takeOpenRequests). */
export const DOCUMENT_OPEN_FILES_EVENT = "visualtex-document-open-files";

const DOCUMENT_PATH = /\.(md|markdown|vtxdoc)$/i;
const waiting: string[] = [];

/** Whether a path names a document the app opens in place. */
export function isDocumentPath(path: string): boolean {
  return DOCUMENT_PATH.test(path);
}

/** Paths to open as documents: the main window shows document mode, whose workspace takes them. */
export function requestOpenDocuments(paths: string[]): void {
  const documents = paths.filter(isDocumentPath);
  if (!documents.length) return;
  for (const path of documents) if (!waiting.includes(path)) waiting.push(path);
  useDocumentModeStore.getState().setContent("document");
  window.dispatchEvent(new Event(DOCUMENT_OPEN_FILES_EVENT));
}

/** The paths waiting to be opened (each is given once). */
export function takeOpenRequests(): string[] {
  return waiting.splice(0);
}

/**
 * The main window's link to the system: the files the app was started with or handed while running (the backend
 * keeps them until asked: document_pending_files) and document files dropped on the window. Returns the undoing.
 */
export function watchSystemOpenRequests(): () => void {
  if (!isTauri()) return () => undefined;
  let disposed = false;
  const disposers: Array<() => void> = [];
  const keep = (dispose: () => void) => (disposed ? dispose() : disposers.push(dispose));
  // (an answer that is not a list opens nothing: a stand-in bridge answers null, and a throw here is a page error)
  const take = () =>
    void invoke<string[] | null>("document_pending_files").then(
      (paths) => requestOpenDocuments(Array.isArray(paths) ? paths : []),
      () => undefined,
    );
  take();
  void listen("document-open-files", take).then(keep, () => undefined);
  void getCurrentWebview()
    .onDragDropEvent((event) => {
      if (event.payload.type === "drop") requestOpenDocuments(event.payload.paths);
    })
    .then(keep, () => undefined);
  return () => {
    disposed = true;
    for (const dispose of disposers) dispose();
  };
}
