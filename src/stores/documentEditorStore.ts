// The document editor's live state for the workspace around it (docs/document-editor/DESIGN.md §6.4): the bottom
// bar's controls, the LaTeX source tab, the status bar and the zoom read the editor through this store.
import { create } from "zustand";
import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { createSafeStorage } from "../runtime/safeStorage";

const storage = createSafeStorage();
const ZOOM_KEY = "visualtex.document.zoom.v1";
const FIT_KEY = "visualtex.document.fit-width.v1";

/** The document zoom steps: 1 is the paper's real size (96 px per inch). */
export const DOCUMENT_ZOOMS = [0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3];

function readZoom(): number {
  const value = Number(storage.getItem(ZOOM_KEY));
  return DOCUMENT_ZOOMS.includes(value) ? value : 1;
}

interface DocumentEditorStore {
  view: EditorView | null;
  state: EditorState | null;
  /** the zoom chosen with + / - (used while the page does not fit the width) */
  zoom: number;
  /** the page is as wide as the window (the default); + / - turn it off */
  fitWidth: boolean;
  /** the zoom the page is drawn at (fitting the width: the fitted one) */
  effectiveZoom: number;
  setEditor(view: EditorView | null, state: EditorState | null): void;
  setZoom(zoom: number): void;
  setFitWidth(on: boolean): void;
  setEffectiveZoom(zoom: number): void;
}

export const useDocumentEditorStore = create<DocumentEditorStore>((set) => ({
  view: null,
  state: null,
  zoom: readZoom(),
  fitWidth: storage.getItem(FIT_KEY) !== "0",
  effectiveZoom: readZoom(),
  setEditor: (view, state) => set({ view, state }),
  setZoom: (zoom) => {
    storage.setItem(ZOOM_KEY, String(zoom));
    storage.setItem(FIT_KEY, "0");
    set({ zoom, fitWidth: false, effectiveZoom: zoom });
  },
  setFitWidth: (on) => {
    storage.setItem(FIT_KEY, on ? "1" : "0");
    set((state) => ({ fitWidth: on, effectiveZoom: on ? state.effectiveZoom : state.zoom }));
  },
  setEffectiveZoom: (effectiveZoom) => set({ effectiveZoom }),
}));

/** The next zoom step above (+1) or below (-1) the zoom the page has now. */
export function stepDocumentZoom(zoom: number, direction: 1 | -1): number {
  if (direction > 0) return DOCUMENT_ZOOMS.find((step) => step > zoom + 0.005) ?? DOCUMENT_ZOOMS[DOCUMENT_ZOOMS.length - 1];
  return [...DOCUMENT_ZOOMS].reverse().find((step) => step < zoom - 0.005) ?? DOCUMENT_ZOOMS[0];
}

/** Words (a Chinese character counts as one) and formulas in the document. */
export function documentCounts(state: EditorState | null): { words: number; formulas: number } {
  let words = 0;
  let formulas = 0;
  state?.doc.descendants((node) => {
    if (node.isText) words += (node.text ?? "").match(/[\p{Script=Han}]|[^\s\p{Script=Han}]+/gu)?.length ?? 0;
    else if (node.type.name === "math_inline" || node.type.name === "math_display") formulas += 1;
    return true;
  });
  return { words, formulas };
}
