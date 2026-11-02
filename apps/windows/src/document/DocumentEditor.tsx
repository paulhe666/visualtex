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
} from "../../../shared/visualtex-document/src/index.ts";
import { cachedPreviewMarkup } from "../components/MathPreview";
import { markVisualTexFormulaFontGlyphs } from "../editor/formulaFontPreferences";
import { getFormulaFontRevision, subscribeFormulaFontRevision } from "../editor/visualTexLayoutRuntime";
import { getCustomSymbolRevision, subscribeCustomSymbols } from "../math/customSymbolRegistry";
import { DocumentFormulaField, type DocumentFormulaFieldApi } from "./DocumentFormulaField";
import { stripVisualTexAlignmentMarkers } from "../editor/alignmentMarkers";
import { toPortableBoldLatex } from "../clipboard/portableBoldLatex";
import { normalizeChineseLatex } from "../editor/normalizeChineseLatex";
import { useEditorStore } from "../stores/editorStore";
import { ContextMenu, type ContextMenuState } from "./MathchaChrome";
import { useDocumentEditorStore } from "../stores/documentEditorStore";
import { useDocumentFormatStore } from "../stores/documentFormatStore";
import { openDocumentFormula } from "./openDocumentFormula";
import { CommandMenu } from "./CommandMenu";
import { LinkDialog } from "./LinkDialog";
import { FootnoteDialog } from "./FootnoteDialog";
import { CitationPicker } from "./CitationPicker";
import { MarginGuides } from "./MarginGuides";
import { createPortal } from "react-dom";
import { documentMathCommands } from "./documentMathCommands";
import { codeHighlighter } from "./codeHighlighter";
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
  return toPortableBoldLatex(stripVisualTexAlignmentMarkers(normalizeChineseLatex(filled)));
}


export interface DocumentEditorHandle {
  view(): EditorView | null;
  run(command: Command): boolean;
  toLatex(): string;
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
}

function createFormulaHost(mathSize: () => number): FormulaHost {
  return {
    renderStatic(latex, display) {
      const source = display ? latex : `{\\textstyle ${latex}}`;
      const markup = cachedPreviewMarkup(source, getCustomSymbolRevision(), getFormulaFontRevision());
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

export const DocumentEditor = forwardRef<DocumentEditorHandle, Props>(function DocumentEditor({ initialJson, onChange }, ref) {
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
  const language = useEditorStore((state) => state.language);
  const lang = language === "en" ? "en" : "zh";


  const makeState = (doc: Node) =>
    createDocumentState(documentSchema, doc, {
      host: createFormulaHost(() => DOCUMENT_MATH_SIZE * zoomRef.current),
      formulaText: portableFormulaLatex,
      mathCommands: documentMathCommands,
      codeHighlighter,
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
    setPmState(view.state);
    useDocumentEditorStore.getState().setEditor(view, view.state);
    view.focus();
    // Formula fonts or custom symbols changed: the closed formulas are drawn again.
    const refresh = () => window.dispatchEvent(new Event(FORMULAS_REFRESH_EVENT));
    const unsubscribeFonts = subscribeFormulaFontRevision(refresh);
    const unsubscribeSymbols = subscribeCustomSymbols(refresh);
    return () => {
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
      const width = frame.clientWidth - 2 * SHEET_SIDE_ROOM;
      if (width <= 0) return;
      const next = Math.round(Math.min(4, Math.max(0.3, width / paperWidth)) * 100) / 100;
      if (Math.abs(useDocumentEditorStore.getState().effectiveZoom - next) > 0.004) useDocumentEditorStore.getState().setEffectiveZoom(next);
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
      margin: `${SHEET_SIDE_ROOM}px auto`,
      fontSize: settings.fontSize * (96 / 72) * zoom,
      // the pages: white, then the gap between them (the canvas shows through)
      background: `repeating-linear-gradient(to bottom, var(--vtx-paper, #fff) 0, var(--vtx-paper, #fff) ${pageHeight}px, transparent ${pageHeight}px, transparent ${pageHeight + PAGE_GAP}px)`,
      ["--vtx-body-font" as string]: bodyFontFamily(settings),
      ["--vtx-line-height" as string]: String(SINGLE_LINE_HEIGHT * settings.lineSpacing),
      ["--vtx-math-scale" as string]: LARGE_BODY_FONTS.test(settings.latinFont) ? String(17 / 15) : "1.04",
      ["--vtx-bibliography-title" as string]: lang === "en" ? '"References"' : '"参考文献"',
    } as CSSProperties;
  }, [settings, zoom, lang]);

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
        <FootnoteDialog key={`${footnote.pos}`} view={view} pos={footnote.pos} rect={footnote.rect} lang={lang} onClose={() => setFootnote(null)} />
      )}
    </div>
  );
});
