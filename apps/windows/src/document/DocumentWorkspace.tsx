// The main window's document mode (docs/document-editor/DESIGN.md §6.4): the document editor in the workspace's
// editor area, with autosave, the file actions of the main menu (new, open, save, export LaTeX, copy LaTeX) and an
// editor handle so that formula commands from the rest of the app (tools, tiles, OCR) go into the document.
import { useCallback, useEffect, useRef, useState, type RefObject } from "react";
import type { Node } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import {
  hasMathDelimiters,
  insertDisplayLatex,
  insertInlineFormula,
  insertInlineLatex,
  mathTextToSlice,
  parseMathText,
} from "../../../shared/visualtex-document/src/index.ts";
import type { MathEditorHandle, MathEditorInsertionTarget } from "../editor/MathEditor";
import { createSafeStorage } from "../runtime/safeStorage";
import { useEditorStore } from "../stores/editorStore";
import type { LatexCommand } from "../types/command";
import { DocumentEditor, type DocumentEditorHandle } from "./DocumentEditor";
import { insertCommandIntoFormula, insertLatexIntoFormula, waitForOpenFormula } from "./formulaInsert";
import { openDocumentFormula } from "./openDocumentFormula";
import {
  DOCUMENT_COPY_LATEX_EVENT,
  DOCUMENT_COPY_MARKDOWN_EVENT,
  DOCUMENT_EXPORT_LATEX_EVENT,
  DOCUMENT_EXPORT_MARKDOWN_EVENT,
  DOCUMENT_NEW_REQUEST_EVENT,
  DOCUMENT_OPEN_REQUEST_EVENT,
  DOCUMENT_SAVE_REQUEST_EVENT,
} from "./documentEvents";

const AUTOSAVE_KEY = "visualtex.document.autosave.v1";
const storage = createSafeStorage();

function readAutosave(): unknown {
  try {
    const text = storage.getItem(AUTOSAVE_KEY);
    if (!text) return undefined;
    return (JSON.parse(text) as { doc?: unknown }).doc;
  } catch {
    return undefined;
  }
}

function download(name: string, text: string, type: string) {
  const url = URL.createObjectURL(new Blob([text], { type }));
  const link = window.document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

/**
 * A recognised formula (OCR) at a place in the document: text with $…$ / \[…\] becomes text and formulas, as when
 * pasted; a formula of several lines, or one on an empty line, a display formula; otherwise an inline formula.
 */
function insertRecognized(view: EditorView, latex: string, place: { from: number; to: number } | null): boolean {
  const source = latex.trim();
  if (!source) return false;
  if (place) {
    const max = view.state.doc.content.size;
    try {
      view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, Math.min(place.from, max), Math.min(place.to, max))));
    } catch {
      // The place is gone: the caret's place now.
    }
  }
  // Text around formulas ($…$, \(…\), \[…\]) or a whole display environment: as pasted text. A bare formula
  // (\frac…, \begin{aligned}…) is one formula.
  const mixed = /\$|\\\(|\\\[/.test(source) || /^\\begin\{(?:equation|align|gather|multline|displaymath)\*?\}/.test(source);
  if (mixed && hasMathDelimiters(source)) {
    view.dispatch(view.state.tr.replaceSelection(mathTextToSlice(view.state.schema, parseMathText(source))).scrollIntoView());
    return true;
  }
  const display = /\\\\|\\begin\{/.test(source) || view.state.selection.$from.parent.content.size === 0;
  const command = display ? insertDisplayLatex(source) : insertInlineLatex(source);
  return command(view.state, view.dispatch, view);
}

/** The editor handle in document mode: formula commands go into the open formula (or a new inline one). */
function createDocumentEditorAdapter(document: () => DocumentEditorHandle | null): MathEditorHandle {
  const target = async () => {
    const open = openDocumentFormula();
    if (open) return open;
    const editor = document();
    if (!editor?.run(insertInlineFormula)) return null;
    return waitForOpenFormula();
  };
  const insertCommand = (command: LatexCommand) => {
    void target().then((open) => open && insertCommandIntoFormula(open, command));
  };
  const insertLatex = (latex: string) => {
    void target().then((open) => open && insertLatexIntoFormula(open, latex));
  };
  // Quick OCR keeps the caret's place when the screenshot starts; the recognised formula goes there (the place
  // holds while the document is unchanged; otherwise the caret's place then).
  const places = new Map<string, { doc: Node; from: number; to: number }>();
  let placeCount = 0;
  const captureInsertionTarget = (): MathEditorInsertionTarget | null => {
    const view = document()?.view();
    if (!view) return null;
    const { from, to } = view.state.selection;
    const lineId = `document-place-${++placeCount}`;
    places.set(lineId, { doc: view.state.doc, from, to });
    return { lineId, ranges: [[from, to]], direction: "forward" };
  };
  const insertLatexAt = (target: MathEditorInsertionTarget, latex: string) => {
    const view = document()?.view();
    if (!view) return false;
    const place = places.get(target.lineId);
    places.delete(target.lineId);
    const done = insertRecognized(view, latex, place && place.doc === view.state.doc ? place : null);
    view.focus();
    return done;
  };
  return {
    insertCommand,
    insertLatex,
    insertLatexAt,
    appendLatex: insertLatex,
    focus: () => document()?.focus(),
    addLine: () => undefined,
    commitPendingTransaction: () => undefined,
    getSelectionMap: () => ({}),
    restoreSelection: async () => false,
    captureSelectionTarget: () => null,
    captureInsertionTarget,
    applySelectionStyle: () => false,
    refreshLayout: () => undefined,
    setActiveLineMode: () => false,
  };
}

interface Props {
  /** The workspace's editor handle: formula commands from the app insert through it. */
  editorRef: RefObject<MathEditorHandle | null>;
}

export function DocumentWorkspace({ editorRef }: Props) {
  const documentRef = useRef<DocumentEditorHandle>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const [initial] = useState(readAutosave);
  const [notice, setNotice] = useState<string | null>(null);
  const noticeTimer = useRef(0);

  useEffect(() => {
    const adapter = createDocumentEditorAdapter(() => documentRef.current);
    editorRef.current = adapter;
    // The headless tests drive the handle the way the app's tools and OCR do.
    if (import.meta.env.DEV) (window as unknown as { __documentAdapter?: MathEditorHandle }).__documentAdapter = adapter;
    return () => {
      if (editorRef.current === adapter) editorRef.current = null;
    };
  }, [editorRef]);

  // Autosave: the document is kept between sessions (debounced).
  const saveTimer = useRef(0);
  const onChange = useCallback((doc: Node) => {
    window.clearTimeout(saveTimer.current);
    saveTimer.current = window.setTimeout(() => {
      storage.setItem(AUTOSAVE_KEY, JSON.stringify({ version: 1, doc: doc.toJSON() }));
    }, 400);
  }, []);

  const flash = (message: string) => {
    setNotice(message);
    window.clearTimeout(noticeTimer.current);
    noticeTimer.current = window.setTimeout(() => setNotice(null), 1800);
  };

  const fileName = (ext: string) => {
    const doc = documentRef.current?.view()?.state.doc;
    const heading = doc?.firstChild?.type.name === "heading" ? doc.firstChild.textContent : "";
    const base = heading.trim().replace(/[\\/:*?"<>|]/g, "-").slice(0, 60) || (isEn ? "Untitled document" : "未命名文档");
    return base + ext;
  };

  const actions = useRef({
    save: () => {},
    open: () => {},
    create: () => {},
    exportLatex: () => {},
    copyLatex: () => {},
    exportMarkdown: () => {},
    copyMarkdown: () => {},
  });
  actions.current = {
    save: () => {
      const text = documentRef.current?.serialize();
      if (!text) return;
      download(fileName(".vtxdoc"), text, "application/json");
      flash(isEn ? "Document saved" : "文档已保存");
    },
    open: () => fileRef.current?.click(),
    create: () => {
      // An undoable clear: the old document stays one Ctrl+Z away.
      const view = documentRef.current?.view();
      if (!view) return;
      const paragraph = view.state.schema.nodes.paragraph.create();
      view.dispatch(view.state.tr.replaceWith(0, view.state.doc.content.size, paragraph).scrollIntoView());
      view.focus();
    },
    exportLatex: () => {
      const tex = documentRef.current?.toLatex();
      if (!tex) return;
      download(fileName(".tex"), tex, "text/x-tex");
      flash(isEn ? "LaTeX exported" : "LaTeX 已导出");
    },
    exportMarkdown: () => {
      const markdown = documentRef.current?.toMarkdown();
      if (!markdown) return;
      download(fileName(".md"), markdown, "text/markdown");
      flash(isEn ? "Markdown exported" : "Markdown 已导出");
    },
    copyMarkdown: () => {
      const markdown = documentRef.current?.toMarkdown();
      if (!markdown) return;
      void navigator.clipboard.writeText(markdown).then(
        () => flash(isEn ? "Markdown copied" : "Markdown 已复制"),
        () => flash(isEn ? "Copy failed" : "复制失败"),
      );
    },
    copyLatex: () => {
      const tex = documentRef.current?.toLatex();
      if (!tex) return;
      void navigator.clipboard.writeText(tex).then(
        () => flash(isEn ? "LaTeX copied" : "LaTeX 已复制"),
        () => flash(isEn ? "Copy failed" : "复制失败"),
      );
    },
  };

  // The main menu and Ctrl+S / Ctrl+O / Ctrl+N reach the document through window events.
  useEffect(() => {
    const bind: Array<[string, () => void]> = [
      [DOCUMENT_SAVE_REQUEST_EVENT, () => actions.current.save()],
      [DOCUMENT_OPEN_REQUEST_EVENT, () => actions.current.open()],
      [DOCUMENT_NEW_REQUEST_EVENT, () => actions.current.create()],
      [DOCUMENT_EXPORT_LATEX_EVENT, () => actions.current.exportLatex()],
      [DOCUMENT_COPY_LATEX_EVENT, () => actions.current.copyLatex()],
      [DOCUMENT_EXPORT_MARKDOWN_EVENT, () => actions.current.exportMarkdown()],
      [DOCUMENT_COPY_MARKDOWN_EVENT, () => actions.current.copyMarkdown()],
    ];
    for (const [name, fn] of bind) window.addEventListener(name, fn);
    return () => {
      for (const [name, fn] of bind) window.removeEventListener(name, fn);
    };
  }, []);

  return (
    <section className="vtx-document-workspace" aria-label={isEn ? "Document" : "文档"}>
      <DocumentEditor ref={documentRef} initialJson={initial} onChange={onChange} />
      {notice && <div className="mc-notice" role="status">{notice}</div>}
      <input
        ref={fileRef}
        type="file"
        accept=".vtxdoc,.json,.md,.markdown,.tex"
        className="visually-hidden"
        onChange={async (event) => {
          const file = event.target.files?.[0];
          event.target.value = "";
          if (!file) return;
          try {
            // A Markdown file is read as Markdown (Typora, Obsidian, GitHub); otherwise the document format.
            if (/\.(md|markdown)$/i.test(file.name)) documentRef.current?.loadMarkdown(await file.text());
            else if (/\.tex$/i.test(file.name)) documentRef.current?.loadLatex(await file.text());
            else documentRef.current?.load(await file.text());
            flash(isEn ? "Document opened" : "文档已打开");
          } catch (error) {
            flash((isEn ? "Cannot open: " : "无法打开：") + (error instanceof Error ? error.message : String(error)));
          }
        }}
      />
    </section>
  );
}
