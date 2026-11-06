// The document editor (docs/document-editor/DESIGN.md, MATHCHA-SPEC.md): one ProseMirror editor for text,
// formulas and tables, with the context menu, on a writing surface that fills the
// window and goes on downwards, and the "\" command menu; its controls, LaTeX and status are the workspace's
// (DocumentControls, documentEditorStore). Closed formulas are the engine's static markup; the open one is a
// DocumentFormulaField.
import { forwardRef, useCallback, useEffect, useImperativeHandle, useMemo, useRef, useState, type CSSProperties } from "react";
import { createRoot, type Root } from "react-dom/client";
import { flushSync } from "react-dom";
import { NodeSelection, TextSelection, type Command, type EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import type { Node } from "prosemirror-model";
import {
  createDocumentState,
  documentSchema,
  DEFAULT_BOARD_LABELS,
  documentToLatex,
  documentToMarkdown,
  markdownSourceKey,
  markdownSourceOf,
  readMarkdown,
  latexToMarkdown,
  FORMULAS_REFRESH_EVENT,
  LINK_REQUEST_EVENT,
  openFormulaId,
  parseDocument,
  serializeDocument,
  sliceToLatex,
  slicePlainText,
  splitInlineMath,
  bodyFontFamily,
  documentSettings,
  FOOTNOTE_EDIT_EVENT,
  CITATION_REQUEST_EVENT,
  PAPER_SIZES,
  refreshPagination,
  setDocumentSettings,
  SINGLE_LINE_HEIGHT,
  type DocumentSettings,
  type FootnoteEditRequest,
  type FormulaHost,
  type FormulaSession,
  IMAGE_REQUEST_EVENT,
  imageFiles,
  insertImagesAtSelection,
  registerImageFile,
} from "../../shared/visualtex-document/src/index.ts";
import { invoke, isTauri } from "@tauri-apps/api/core";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { cachedPreviewMarkup } from "../components/MathPreview";
import { convertVisualTexLatexToFieldMarkup } from "../editor/mathLiveIntegralCompatibility";
import { markVisualTexFormulaFontGlyphs } from "../editor/formulaFontPreferences";
import { getFormulaFontRevision, subscribeFormulaFontRevision } from "../editor/visualTexLayoutRuntime";
import { getCustomSymbolRevision, subscribeCustomSymbols } from "../math/customSymbolRegistry";
import { DocumentFormulaField, type DocumentFormulaFieldApi } from "./DocumentFormulaField";
import { stripVisualTexAlignmentMarkers } from "../editor/alignmentMarkers";
import { toPortableBoldLatex } from "../clipboard/portableBoldLatex";
import { normalizeChineseLatexAsWritten } from "../editor/normalizeChineseLatex";
import { useEditorStore } from "../stores/editorStore";
import { ContextMenu, type ContextMenuState } from "./MathchaChrome";
import { useDocumentEditorStore } from "../stores/documentEditorStore";
import { useDocumentFormatStore } from "../stores/documentFormatStore";
import { CommandMenu } from "./CommandMenu";
import { LinkDialog } from "./LinkDialog";
import { FootnoteDialog } from "./FootnoteDialog";
import { CitationPicker } from "./CitationPicker";
import { MarginGuides } from "./MarginGuides";
import { createPortal } from "react-dom";
import { documentMathCommands } from "./documentMathCommands";
import { codeHighlighter } from "./codeHighlighter";
import { drawDiagram } from "./diagrams";
import "./document.css";

/** Mathcha's defaults (MATHCHA-SPEC.md): text 15px, formulas 17px, at zoom 1. */
export const DOCUMENT_TEXT_SIZE = 15;
export const DOCUMENT_MATH_SIZE = 17;

/** Screen pixels per millimetre at zoom 1 (96 px per inch: the paper's real size). */
export const PX_PER_MM = 96 / 25.4;
/** The gap between two pages, and the room around the sheet, in screen pixels. */
export const PAGE_GAP = 16;
const SHEET_SIDE_ROOM = 24;

/** Body fonts whose letters are large for their size (sans): their formulas are set a little larger. */
const LARGE_BODY_FONTS = /arial|helvetica|calibri|segoe|verdana|tahoma|yahei|dengxian|pingfang|hei|sans/i;

/** A document formula's LaTeX as it leaves the app (the same conversions as Copy LaTeX). */
export function portableFormulaLatex(latex: string): string {
  // An unfilled slot (\placeholder{}) leaves its group empty.
  const filled = latex.replace(/\\placeholder(?:\[[^\]]*\])?\{\}/g, "");
  // the letters as written: the page shows the stored formula, and exports must say the same (no upright judgment)
  return toPortableBoldLatex(stripVisualTexAlignmentMarkers(normalizeChineseLatexAsWritten(filled)));
}


export interface DocumentEditorHandle {
  view(): EditorView | null;
  run(command: Command): boolean;
  toLatex(): string;
  /** The document's images as the export writes them (figures/fig-1.png …, or assets/ for Markdown). */
  imageFiles(folder: string): Array<{ path: string; bytes: Uint8Array }>;
  /** The document as Markdown (Typora / Obsidian / GitHub flavour). */
  toMarkdown(): string;
  serialize(): string;
  /** Open Markdown text as the document (undoable). */
  loadMarkdown(text: string): void;
  /** A .tex file: its document structure read in (latexImport.ts). */
  loadLatex(text: string): void;
  load(text: string): void;
  setDoc(doc: Node): void;
  focus(): void;
}

interface Props {
  initialJson?: unknown;
  onChange?: (doc: Node) => void;
  /** An image pasted into an open formula (a screenshot of a formula): the app recognises it (OCR) and puts it in
   *  the formula. In the text an image is inserted as an image (DESIGN.md §11). */
  onImage?: (file: File) => void;
}

/** The images a paste or drop carries, unless it also carries text (then the text is pasted). */
function imagesOf(data: DataTransfer | null): File[] {
  if (!data) return [];
  const images = [...data.files].filter((file) => file.type.startsWith("image/"));
  if (!images.length) return [];
  return data.getData("text/plain").trim() ? [] : images;
}

const IMAGE_FILE = /\.(png|jpe?g|gif|webp|bmp|svg|tiff?|heic|avif)$/i;

/** When the page last took a dropped image itself (a webview that gives drops to the page as well as to Tauri). */
let lastPageDrop = 0;

/** Image files into the document at the selection (or a page point): an image line, or onto the line there. */
async function insertImageFiles(view: EditorView, files: Blob[], at?: { x: number; y: number }) {
  const keys = (await Promise.all(files.map((file) => registerImageFile(file)))).filter((key): key is string => !!key);
  if (!keys.length || view.isDestroyed) return;
  if (at) {
    const hit = view.posAtCoords({ left: at.x, top: at.y });
    if (hit) view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(hit.pos))));
  }
  const tr = insertImagesAtSelection(view.state, keys, view.state.schema);
  if (tr) view.dispatch(tr.scrollIntoView());
  view.focus();
}

const documentMarkupCache = new Map<string, string>();

/**
 * A formula of the document drawn as its resting editor draws it (the editor's own layout, inline formulas in text
 * style): the live editor takes the picture's place without a pixel's change.
 */
function documentFormulaMarkup(latex: string, display: boolean): string {
  const key = `${getCustomSymbolRevision()}\u0000${getFormulaFontRevision()}\u0000${display ? "d" : "i"}\u0000${latex}`;
  const cached = documentMarkupCache.get(key);
  if (cached !== undefined) return cached;
  let markup = "";
  try {
    markup = convertVisualTexLatexToFieldMarkup(latex, {
      defaultMode: display ? "math" : "inline-math",
      visualTexEditorLayout: true,
    } as Parameters<typeof convertVisualTexLatexToFieldMarkup>[1]);
  } catch (error) {
    console.warn("VisualTeX could not draw a document formula.", error);
  }
  if (documentMarkupCache.size > 2000) documentMarkupCache.delete(documentMarkupCache.keys().next().value!);
  documentMarkupCache.set(key, markup);
  return markup;
}

function createFormulaHost(mathSize: () => number, onImage: () => ((file: File) => void) | undefined): FormulaHost {
  return {
    renderStatic(latex, display) {
      const markup = documentFormulaMarkup(latex, display);
      if (!markup) return "";
      const holder = document.createElement("span");
      holder.innerHTML = markup;
      markVisualTexFormulaFontGlyphs(holder);
      return holder.innerHTML;
    },
    open(container, { latex, display, entry, callbacks, passive }): FormulaSession {
      const apiRef: { current: DocumentFormulaFieldApi | null } = { current: null };
      const root: Root = createRoot(container);
      flushSync(() =>
        root.render(
          <DocumentFormulaField
            id="doc-formula"
            latex={latex}
            display={display}
            entry={entry}
            passive={passive}
            onPasteImage={onImage()}
            fontSize={parseFloat(getComputedStyle(container).fontSize) || mathSize()}
            callbacks={callbacks}
            apiRef={(api) => {
              apiRef.current = api;
            }}
          />,
        ),
      );
      return {
        showPart: (range) => apiRef.current?.showPart(range),
        stopAt: (x, y) => apiRef.current?.stopAt(x, y) ?? null,
        topLevelStop: (offset, side) => apiRef.current?.topLevelStop(offset, side) ?? offset,
        lastOffset: () => apiRef.current?.lastOffset() ?? null,
        selectStops: (anchor, focus) => apiRef.current?.selectStops(anchor, focus),
        anchorStop: () => apiRef.current?.anchorStop() ?? null,
        latexOf: (range) => apiRef.current?.latexOf(range) ?? "",
        withoutRange: (range) => apiRef.current?.withoutRange(range) ?? null,
        styled: (range, _kind, enable) => apiRef.current?.styled(range, enable) ?? null,
        rangeBold: (range) => apiRef.current?.rangeBold(range) ?? null,
        toggleBold: () => apiRef.current?.toggleBold() ?? false,
        enter: (entry) => apiRef.current?.enter(entry),
        rest: () => apiRef.current?.rest(),
        whenDrawn(done) {
          // The editor's first frame is not styled yet (its shadow styles apply a frame later): drawn once its formula
          // is laid out with them. (An editor opened at once shows only once it has the focus: until then the
          // document's selection still spans it, and WebKit would paint that selection in it for a frame.)
          let frames = 0;
          const check = () => {
            const root = container.querySelector("math-field")?.shadowRoot;
            const ink = root?.querySelector(".ML__latex");
            const speech = root?.querySelector<HTMLElement>(".ML__sr-only");
            const styled = !speech || getComputedStyle(speech).position === "absolute";
            const focused = passive || !!root?.activeElement;
            if ((ink && styled && focused) || ++frames > 30 || !container.isConnected) done();
            else requestAnimationFrame(check);
          };
          requestAnimationFrame(check);
        },
        update(next, selection) {
          apiRef.current?.update(next, selection);
        },
        typeKey(key) {
          apiRef.current?.typeKey(key);
        },
        focus() {
          apiRef.current?.focus();
        },
        stopPointerTracking() {
          apiRef.current?.stopPointerTracking();
        },
        flush() {
          apiRef.current?.flush();
        },
        destroy() {
          // Outside ProseMirror's update cycle: an edit not yet reported is stored, then the field unmounts.
          queueMicrotask(() => {
            apiRef.current?.flush();
            root.unmount();
          });
        },
      };
    },
  };
}

/** A footnote's text as HTML: its $…$ formulas drawn by the engine, the rest as text. */
function noteHtml(text: string): string {
  const escape = (value: string) => value.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!);
  return splitInlineMath(text)
    .map((part) => {
      if (part.kind === "text") return escape(part.text);
      const markup = cachedPreviewMarkup(`{\\textstyle ${part.latex}}`, getCustomSymbolRevision(), getFormulaFontRevision());
      return markup ? `<span class="vtx-note-math">${markup}</span>` : escape(`$${part.latex}$`);
    })
    .join("");
}

function emptyDoc(): Node {
  return documentSchema.nodes.doc.create(null, documentSchema.nodes.paragraph.create());
}

export const DocumentEditor = forwardRef<DocumentEditorHandle, Props>(function DocumentEditor({ initialJson, onChange, onImage }, ref) {
  const hostRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorView | null>(null);
  const [frame, setFrame] = useState<HTMLDivElement | null>(null);
  const frameRef = useRef<HTMLDivElement | null>(null);
  frameRef.current = frame;
  const [pmState, setPmState] = useState<EditorState | null>(null);
  const zoom = useDocumentEditorStore((store) => store.effectiveZoom);
  const fitWidth = useDocumentEditorStore((store) => store.fitWidth);
  const manualZoom = useDocumentEditorStore((store) => store.zoom);
  const [menu, setMenu] = useState<ContextMenuState | null>(null);
  // Ctrl+K and the menus' link entries ask for the link dialog.
  const [linkDialog, setLinkDialog] = useState(0);
  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const open = () => setLinkDialog((n) => n + 1);
    host.addEventListener(LINK_REQUEST_EVENT, open);
    return () => host.removeEventListener(LINK_REQUEST_EVENT, open);
  }, []);
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;
  // a footnote's editor (a new footnote, a click on its number or on its note at the foot of the page)
  const [footnote, setFootnote] = useState<FootnoteEditRequest | null>(null);
  // the citation picker at the caret ("\cite", "/cite")
  const [citation, setCitation] = useState<{ left: number; top: number } | null>(null);
  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const open = () => {
      const view = viewRef.current;
      if (!view) return;
      const caret = view.coordsAtPos(view.state.selection.head);
      setCitation({ left: Math.max(4, Math.min(caret.left, window.innerWidth - 360)), top: Math.min(caret.bottom + 6, window.innerHeight - 340) });
    };
    host.addEventListener(CITATION_REQUEST_EVENT, open);
    return () => host.removeEventListener(CITATION_REQUEST_EVENT, open);
  }, []);
  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const open = (event: Event) => setFootnote((event as CustomEvent<FootnoteEditRequest>).detail);
    host.addEventListener(FOOTNOTE_EDIT_EVENT, open);
    // the page layer (footnotes at the pages' feet) is outside the editor: its requests come from the sheet
    const sheet = host.closest(".vtx-sheet");
    sheet?.addEventListener(FOOTNOTE_EDIT_EVENT, open);
    return () => {
      host.removeEventListener(FOOTNOTE_EDIT_EVENT, open);
      sheet?.removeEventListener(FOOTNOTE_EDIT_EVENT, open);
    };
  }, []);
  const onChangeRef = useRef(onChange);
  onChangeRef.current = onChange;
  const onImageRef = useRef(onImage);
  onImageRef.current = onImage;
  // a footnote's formulas: the document's formula fields at the footnote panel's size
  const [noteFormulaHost] = useState(() => createFormulaHost(() => 14, () => onImageRef.current));
  const language = useEditorStore((state) => state.language);
  const lang = language === "en" ? "en" : "zh";


  const makeState = (doc: Node) =>
    createDocumentState(documentSchema, doc, {
      host: createFormulaHost(() => DOCUMENT_MATH_SIZE * zoomRef.current, () => onImageRef.current),
      formulaText: portableFormulaLatex,
      mathCommands: documentMathCommands,
      codeHighlighter,
      codeDiagram: drawDiagram,
      inputMode: () => useDocumentFormatStore.getState().inputMode,
      markdownRules: () => useDocumentFormatStore.getState().markdownRules,
      // the page view: the paper's pages, the footnotes at their feet, the page numbers
      pagination: {
        geometry: () => ({ pxPerMm: PX_PER_MM * zoomRef.current, gap: PAGE_GAP }),
        renderNote: noteHtml,
        pageLabel: (page, pages) => (lang === "en" ? `Page ${page} of ${pages}` : `第 ${page} 页，共 ${pages} 页`),
      },
      codeLabels:
        lang === "en"
          ? { plainText: "Plain Text", copy: "Copy", copied: "Copied", language: "Language" }
          : { plainText: "纯文本", copy: "复制", copied: "已复制", language: "语言" },
      boardLabels: () =>
        lang === "en"
          ? DEFAULT_BOARD_LABELS
          : {
              alignLeft: "左对齐",
              alignCenter: "水平居中",
              alignRight: "右对齐",
              alignTop: "顶端对齐",
              alignMiddle: "垂直居中",
              alignBottom: "底端对齐",
              distributeH: "横向分布",
              distributeV: "纵向分布",
              sameWidth: "等宽",
              tidyRow: "排成一行",
              tidyGrid: "排成网格",
              front: "置于顶层",
              back: "置于底层",
            },
      imageLabels: () =>
        lang === "en"
          ? { left: "Align left", center: "Centre", right: "Align right", wrapLeft: "Wrap text (image left)", wrapRight: "Wrap text (image right)" }
          : { left: "左对齐", center: "居中", right: "右对齐", wrapLeft: "文字环绕（图片在左）", wrapRight: "文字环绕（图片在右）" },
    });

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    let doc: Node;
    try {
      doc = initialJson ? documentSchema.nodeFromJSON(initialJson) : emptyDoc();
      doc.check();
    } catch {
      doc = emptyDoc();
    }
    const view = new EditorView(host, {
      state: makeState(doc),
      attributes: { class: "vtx-document", spellcheck: "false" },
      handlePaste(pasteView, event) {
        // an image in the text is inserted (an image pasted in an open formula is the formula's: OCR)
        const images = imagesOf(event.clipboardData);
        if (!images.length) return false;
        event.preventDefault();
        void insertImageFiles(pasteView, images);
        return true;
      },
      handleDrop(dropView, event) {
        const images = imagesOf((event as DragEvent).dataTransfer);
        if (!images.length) return false;
        event.preventDefault();
        lastPageDrop = Date.now();
        void insertImageFiles(dropView, images, { x: (event as DragEvent).clientX, y: (event as DragEvent).clientY });
        return true;
      },
      dispatchTransaction(tr) {
        const next = view.state.apply(tr);
        // Selecting a display formula (to open it) puts the browser's selection around the whole line; in a
        // window narrower than the page the browser would scroll sideways to its end. The sideways scroll stays.
        const canvas = frameRef.current;
        const keepLeft =
          canvas && next.selection instanceof NodeSelection && next.selection.node.type.name === "math_display"
            ? canvas.scrollLeft
            : null;
        view.updateState(next);
        if (canvas && keepLeft !== null) {
          canvas.scrollLeft = keepLeft;
          requestAnimationFrame(() => {
            canvas.scrollLeft = keepLeft;
          });
        }
        setPmState(next);
        useDocumentEditorStore.getState().setEditor(view, next);
        if (tr.docChanged) onChangeRef.current?.(next.doc);
      },
    });
    viewRef.current = view;
    // \includegraphics, "/image" and the toolbar's image button: the system's file picker
    const pick = () => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = "image/*";
      input.multiple = true;
      input.addEventListener("change", () => {
        const files = [...(input.files ?? [])];
        if (files.length) void insertImageFiles(view, files);
      });
      input.click();
    };
    host.addEventListener(IMAGE_REQUEST_EVENT, pick);
    // Files dropped from the system: the Windows webview gives them to Tauri (its drag-drop handler), not to the page
    let unlistenDrop: (() => void) | null = null;
    let disposed = false;
    if (isTauri()) {
      void getCurrentWebview()
        .onDragDropEvent((event) => {
          const payload = event.payload;
          if (payload.type !== "drop" || view.isDestroyed || Date.now() - lastPageDrop < 1500) return;
          const paths = payload.paths.filter((path) => IMAGE_FILE.test(path));
          if (!paths.length) return;
          const scale = window.devicePixelRatio || 1;
          const x = payload.position.x / scale;
          const y = payload.position.y / scale;
          const frame = (view.dom.closest(".mc-canvas") ?? view.dom).getBoundingClientRect();
          if (x < frame.left || x > frame.right || y < frame.top || y > frame.bottom) return;
          // the page may take the same drop itself (its own drop event, which can come after this one)
          window.setTimeout(() => {
            if (Date.now() - lastPageDrop < 1500 || view.isDestroyed) return;
            void Promise.all(
              paths.map((path) =>
                invoke<string>("read_image_file", { path })
                  .then((base64) => new Blob([Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))], { type: path.toLowerCase().endsWith(".svg") ? "image/svg+xml" : "" }))
                  .catch(() => null),
              ),
            ).then((blobs) => insertImageFiles(view, blobs.filter((b): b is Blob => !!b), { x, y }));
          }, 250);
        })
        .then((dispose) => {
          if (disposed) dispose();
          else unlistenDrop = dispose;
        })
        .catch(() => undefined);
    }
    setPmState(view.state);
    useDocumentEditorStore.getState().setEditor(view, view.state);
    view.focus();
    // Formula fonts or custom symbols changed: the closed formulas are drawn again.
    const refresh = () => window.dispatchEvent(new Event(FORMULAS_REFRESH_EVENT));
    const unsubscribeFonts = subscribeFormulaFontRevision(refresh);
    const unsubscribeSymbols = subscribeCustomSymbols(refresh);
    return () => {
      disposed = true;
      unlistenDrop?.();
      host.removeEventListener(IMAGE_REQUEST_EVENT, pick);
      unsubscribeFonts();
      unsubscribeSymbols();
      view.destroy();
      viewRef.current = null;
      if (useDocumentEditorStore.getState().view === view) useDocumentEditorStore.getState().setEditor(null, null);
    };
    // The editor is created once; later documents come through load/setDoc.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const replaceState = (doc: Node) => {
    const state = makeState(doc);
    viewRef.current?.updateState(state);
    setPmState(state);
    if (viewRef.current) useDocumentEditorStore.getState().setEditor(viewRef.current, state);
    onChangeRef.current?.(state.doc);
  };

  useImperativeHandle(ref, () => ({
    view: () => viewRef.current,
    run(command) {
      const view = viewRef.current;
      if (!view) return false;
      const done = command(view.state, view.dispatch, view);
      view.focus();
      return done;
    },
    toLatex() {
      const view = viewRef.current;
      return view ? documentToLatex(view.state.doc, { formulaLatex: portableFormulaLatex }) : "";
    },
    imageFiles(folder) {
      const view = viewRef.current;
      return view ? imageFiles(view.state.doc, folder).files : [];
    },
    toMarkdown() {
      const view = viewRef.current;
      if (!view) return "";
      return documentToMarkdown(view.state.doc, {
        formulaLatex: portableFormulaLatex,
        references: useDocumentFormatStore.getState().markdownReferences,
        source: markdownSourceOf(view.state),
      });
    },
    serialize() {
      const view = viewRef.current;
      return view ? serializeDocument(view.state.doc) : "";
    },
    loadMarkdown(text) {
      const view = viewRef.current;
      if (!view) return;
      // The file is kept with the document: blocks left as they were are written back exactly as they were read.
      const { doc, source } = readMarkdown(documentSchema, text);
      // Replaced in the history, so Ctrl+Z gives the previous document back; ids come from the ids plugin.
      view.dispatch(view.state.tr.replaceWith(0, view.state.doc.content.size, doc.content).setMeta(markdownSourceKey, source).scrollIntoView());
      view.focus();
    },
    loadLatex(text) {
      const view = viewRef.current;
      if (!view) return;
      const { doc } = readMarkdown(documentSchema, latexToMarkdown(text));
      view.dispatch(view.state.tr.replaceWith(0, view.state.doc.content.size, doc.content).setMeta(markdownSourceKey, null).scrollIntoView());
      view.focus();
    },
    load(text) {
      replaceState(parseDocument(documentSchema, text).doc);
    },
    setDoc(doc) {
      replaceState(doc);
    },
    focus() {
      viewRef.current?.focus();
    },
  }));

  // The canvas around the text is part of the writing surface: a click in a margin puts the caret on the nearest
  // line; a click below the last line goes on writing there (on a new line after a table, a list or a formula).
  const onCanvasMouseDown = useCallback((event: React.MouseEvent) => {
    const view = viewRef.current;
    const target = event.target;
    if (!view || event.button !== 0 || !(target instanceof Element) || view.dom.contains(target)) return;
    if (target.closest(".vtx-cmd-menu")) return;
    event.preventDefault();
    const editor = view.dom.getBoundingClientRect();
    const last = view.dom.lastElementChild?.getBoundingClientRect();
    const { doc, schema } = view.state;
    if (!last || event.clientY > last.bottom) {
      const tr = view.state.tr;
      if (doc.lastChild?.type.name === "paragraph") tr.setSelection(TextSelection.create(doc, doc.content.size - 1));
      else {
        tr.insert(doc.content.size, schema.nodes.paragraph.create());
        tr.setSelection(TextSelection.create(tr.doc, tr.doc.content.size - 1));
      }
      view.dispatch(tr.scrollIntoView());
    } else {
      const left = Math.min(Math.max(event.clientX, editor.left + 1), editor.right - 1);
      const hit = view.posAtCoords({ left, top: event.clientY });
      if (hit) view.dispatch(view.state.tr.setSelection(TextSelection.near(doc.resolve(hit.pos))));
    }
    view.focus();
  }, []);

  // Right click: Mathcha's menu. A click on a formula selects it; a click outside the selection puts the caret
  // there (the table items act on that cell).
  const onContextMenu = useCallback((event: React.MouseEvent) => {
    const view = viewRef.current;
    if (!view || !(event.target instanceof Element) || !event.target.closest(".mc-page")) return;
    event.preventDefault();
    const formula = event.target.closest(".vtx-math-inline, .vtx-math-display");
    if (formula && !openFormulaId(view.state)) {
      const pos = view.posAtDOM(formula, 0);
      for (const at of [pos, pos - 1]) {
        if (at >= 0 && view.state.doc.nodeAt(at)?.type.name.startsWith("math_")) {
          view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, at)));
          break;
        }
      }
    } else if (!formula) {
      const hit = view.posAtCoords({ left: event.clientX, top: event.clientY });
      const { from, to } = view.state.selection;
      if (hit && (hit.pos < from || hit.pos > to)) {
        view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(hit.pos))));
      }
    }
    setMenu({ x: event.clientX, y: event.clientY });
  }, []);

  // the page (DESIGN.md §8.2): the paper's width and margins at the zoom; the body's fonts, size and spacing
  const settingsDoc = pmState?.doc ?? null;
  // a margin being dragged on the sheet (MarginGuides): shown before it is set
  const [marginPreview, setMarginPreview] = useState<Partial<DocumentSettings> | null>(null);
  const [sheet, setSheet] = useState<HTMLDivElement | null>(null);
  const savedSettings = useMemo(
    () => (settingsDoc ? documentSettings(settingsDoc) : null),
    // the settings are the doc node's attributes
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [settingsDoc ? JSON.stringify(settingsDoc.attrs) : ""],
  );
  const settings = useMemo(
    () => (savedSettings && marginPreview ? { ...savedSettings, ...marginPreview } : savedSettings),
    [savedSettings, marginPreview],
  );
  // fitting the width: the zoom at which the page fills the window's width
  useEffect(() => {
    const store = useDocumentEditorStore.getState();
    if (!fitWidth || !frame || !settings) {
      if (!fitWidth && store.effectiveZoom !== manualZoom) store.setEffectiveZoom(manualZoom);
      return;
    }
    const paperWidth = PAPER_SIZES[settings.paper].width * PX_PER_MM;
    const fit = () => {
      // the page touches both sides; rounded down so that it never overflows into a horizontal scrollbar
      const width = frame.clientWidth;
      if (width <= 0) return;
      const next = Math.floor(Math.min(4, Math.max(0.3, width / paperWidth)) * 1000) / 1000;
      if (Math.abs(useDocumentEditorStore.getState().effectiveZoom - next) > 0.0005) useDocumentEditorStore.getState().setEffectiveZoom(next);
    };
    fit();
    const observer = new ResizeObserver(fit);
    observer.observe(frame);
    return () => observer.disconnect();
  }, [fitWidth, frame, settings, manualZoom]);
  // the zoom or the page changed: the pages are measured again
  useEffect(() => {
    const view = viewRef.current;
    if (view && settings) refreshPagination(view);
  }, [zoom, settings]);
  const sheetStyle = useMemo<CSSProperties>(() => {
    if (!settings) return { fontSize: DOCUMENT_TEXT_SIZE * zoom };
    const mm = PX_PER_MM * zoom;
    const paper = PAPER_SIZES[settings.paper];
    const pageHeight = paper.height * mm;
    return {
      width: paper.width * mm,
      minHeight: pageHeight,
      padding: `${settings.marginTop * mm}px ${settings.marginRight * mm}px ${settings.marginBottom * mm}px ${settings.marginLeft * mm}px`,
      margin: fitWidth ? `${SHEET_SIDE_ROOM}px 0` : `${SHEET_SIDE_ROOM}px auto`,
      fontSize: settings.fontSize * (96 / 72) * zoom,
      // the pages: white, then the gap between them (the canvas shows through)
      background: `repeating-linear-gradient(to bottom, var(--vtx-paper, #fff) 0, var(--vtx-paper, #fff) ${pageHeight}px, transparent ${pageHeight}px, transparent ${pageHeight + PAGE_GAP}px)`,
      ["--vtx-body-font" as string]: bodyFontFamily(settings),
      // one TeX point on the screen (images: wrapfig's \columnsep and \intextsep)
      // the left margin (what a page command's label may take)
      ["--vtx-margin-left" as string]: `${settings.marginLeft * mm}px`,
      ["--vtx-pt" as string]: `${(96 / 72) * zoom}px`,
      ["--vtx-line-height" as string]: String(SINGLE_LINE_HEIGHT * settings.lineSpacing),
      ["--vtx-math-scale" as string]: LARGE_BODY_FONTS.test(settings.latinFont) ? String(17 / 15) : "1.04",
      ["--vtx-bibliography-title" as string]: lang === "en" ? '"References"' : '"参考文献"',
    } as CSSProperties;
  }, [settings, zoom, lang, fitWidth]);

  const view = viewRef.current;
  const writeText = (text: string) => void navigator.clipboard.writeText(text);
  const menuActions = {
    copy: () => document.execCommand("copy"),
    cut: () => document.execCommand("cut"),
    copyPlainText: () => {
      if (view) writeText(slicePlainText(view.state.selection.content(), portableFormulaLatex));
    },
    copyLatex: () => {
      if (view) writeText(sliceToLatex(view.state.selection.content(), documentSchema, { formulaLatex: portableFormulaLatex }));
    },
    remove: () => {
      if (view) view.dispatch(view.state.tr.deleteSelection().scrollIntoView());
    },
  };

  return (
    <div className="mc-shell">
      <div className="mc-body">
        <div className="mc-canvas is-paged" ref={setFrame} onContextMenu={onContextMenu} onMouseDown={onCanvasMouseDown}>
          <div
            className={"mc-page vtx-sheet" + (settings?.headingNumbering === "1" ? " is-headings-plain" : settings?.headingNumbering === "none" ? " is-headings-none" : "")}
            style={sheetStyle}
            ref={setSheet}
          >
            <div ref={hostRef} className="vtx-document-host" />
            {settings && (
              <MarginGuides
                sheet={sheet}
                settings={settings}
                mm={PX_PER_MM * zoom}
                gap={PAGE_GAP}
                lang={lang}
                onPreview={setMarginPreview}
                onCommit={(patch) => {
                  const view = viewRef.current;
                  if (view) setDocumentSettings(patch)(view.state, view.dispatch);
                }}
              />
            )}
          </div>
          <CommandMenu view={view} state={pmState} lang={lang} frame={frame} />
        </div>
      </div>
      <ContextMenu view={view} state={pmState} lang={lang} at={menu} actions={menuActions} onClose={() => setMenu(null)} />
      {linkDialog > 0 && view && <LinkDialog key={linkDialog} view={view} lang={lang} onClose={() => setLinkDialog(0)} />}
      {citation &&
        view &&
        pmState &&
        createPortal(
          <>
            <div className="document-control-dismiss" onMouseDown={() => setCitation(null)} />
            <div className="document-control-panel" role="dialog" style={{ left: citation.left, top: citation.top }} onMouseDown={(event) => event.preventDefault()}>
              <CitationPicker lang={lang} view={view} state={pmState} close={() => setCitation(null)} />
            </div>
          </>,
          document.body,
        )}
      {footnote && view && (
        <FootnoteDialog
          key={`${footnote.pos}`}
          view={view}
          pos={footnote.pos}
          rect={footnote.rect}
          lang={lang}
          host={noteFormulaHost}
          mathCommands={documentMathCommands}
          font={settings ? bodyFontFamily(settings) : ""}
          onClose={() => setFootnote(null)}
        />
      )}
    </div>
  );
});
