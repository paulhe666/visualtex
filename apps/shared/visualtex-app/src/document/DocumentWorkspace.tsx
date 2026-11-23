// The main window's document mode (docs/document-editor/DESIGN.md §6.4): the document editor in the workspace's
// editor area beside the document library (saved as you type), the file actions of the main menu (new, import, save,
// export) and an editor handle so that formula commands from the rest of the app (tools, tiles, OCR) go into the
// document.
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type RefObject } from "react";
import type { Node } from "prosemirror-model";
import { NodeSelection, Selection, TextSelection } from "prosemirror-state";
import { invoke, isTauri } from "@tauri-apps/api/core";
import type { EditorView } from "prosemirror-view";
import { strToU8, unzipSync, zipSync } from "fflate";
import { GraduationCap } from "lucide-react";
import {
  decodeText,
  documentSchema,
  findMarkdownFile,
  hasMathDelimiters,
  insertDisplayLatex,
  insertInlineFormula,
  insertInlineLatex,
  mathTextToSlice,
  parseDocument,
  parseMathText,
  picturePath,
  picturePaths,
} from "../../../visualtex-document/src/index.ts";
import type { MathEditorHandle, MathEditorInsertionTarget } from "../editor/MathEditor";
import type { FormulaEditSource } from "../history/historyTypes";
import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../editor/alignmentMarkers";
import { createSafeStorage } from "../runtime/safeStorage";
import { useEditorStore } from "../stores/editorStore";
import type { LatexCommand } from "../types/command";
import { DocumentTextToolbar } from "./DocumentControls";
import { DocumentEditor, portableFormulaLatex, type DocumentEditorHandle } from "./DocumentEditor";
import { DocumentLibrary, DocumentLibraryToggle, type DocumentExportFormat, type DocumentLibraryActions } from "./DocumentLibrary";
import {
  chooseDocumentFile,
  chooseExportPath,
  chooseFolder,
  documentLibrary,
  documentTitle,
  libraryDocumentName,
  readFileBytes,
  saveBinaryFile,
  saveTextFile,
  untitledName,
  useDocumentLibraryStore,
  type DocumentContent,
  type Library,
  type LibraryDocument,
} from "./documentLibraryApi";
import { documentContent, loadDocumentFile, newDocumentContent, renameNewPictures, type DocumentFile } from "./documentFiles";
import { keptDocuments, migrateLegacyLibrary } from "./documentMigration";
import { DOCUMENT_OPEN_FILES_EVENT, takeOpenRequests } from "./documentOpenRequests";
import { DocumentFileDialog } from "./DocumentFileDialog";
import { insertCommandIntoFormula, insertLatexIntoFormula, waitForOpenFormula } from "./formulaInsert";
import { openDocumentFormula } from "./openDocumentFormula";
import { documentToPdf, printDocumentInBrowser } from "./documentPdf";
import {
  DOCUMENT_COPY_MARKDOWN_EVENT,
  DOCUMENT_EXPORT_MARKDOWN_EVENT,
  DOCUMENT_EXPORT_PDF_EVENT,
  DOCUMENT_NEW_REQUEST_EVENT,
  DOCUMENT_OPEN_REQUEST_EVENT,
  DOCUMENT_SAVE_AS_EVENT,
  DOCUMENT_SAVE_REQUEST_EVENT,
  DOCUMENT_TUTORIAL_EVENT,
} from "./documentEvents";
import { TutorialDialog } from "./TutorialDialog";
import { getEditorDocumentSnapshot } from "../history/documentHistory";

// Before the library the one document was kept here; the first start with the library takes it in (once; the key
// stays for an older version of the app).
const AUTOSAVE_KEY = "visualtex.document.autosave.v1";
const AUTOSAVE_TAKEN_KEY = "visualtex.document.autosave.taken";
const storage = createSafeStorage();
// set the first time document mode opens; the tutorials open by themselves then, unless document mode was used before
const TUTORIAL_SHOWN_KEY = "visualtex.document.tutorial-shown.v1";

// Document mode was opened in a version before the tutorials (the library, or the single autosave before it).
function documentModeUsedBefore(): boolean {
  return storage.getItem(AUTOSAVE_TAKEN_KEY) !== null || storage.getItem(AUTOSAVE_KEY) !== null;
}

function readAutosave(): unknown {
  try {
    const text = storage.getItem(AUTOSAVE_KEY);
    if (!text) return undefined;
    return (JSON.parse(text) as { doc?: unknown }).doc;
  } catch {
    return undefined;
  }
}

const SAVE_DELAY_MS = 400;

/** The library's answer when a picture file of the name given is there with other bytes (document_files.rs). */
const PICTURE_NAME_TAKEN = /^Another picture is named (.+)$/;

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
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

/**
 * A recognised formula as the app hands it over (OCR): a row a line, the formula rows' alignment markers at the rows'
 * top-level &. In a document it is one formula: its rows in an aligned environment (gathered when no row has an
 * alignment point), so the rows and their alignment stay as recognised.
 */
function recognizedRowsLatex(latex: string): string {
  const rows = latex
    .split(/\n+/)
    .map((row) => row.trim())
    .filter(Boolean);
  if (rows.length <= 1) return (rows[0] ?? "").split(VISUALTEX_ALIGNMENT_MARKER_LATEX).join("");
  const aligned = rows.some((row) => row.includes(VISUALTEX_ALIGNMENT_MARKER_LATEX));
  const environment = aligned ? "aligned" : "gathered";
  const body = rows.map((row) => row.split(VISUALTEX_ALIGNMENT_MARKER_LATEX).join("&")).join(" \\\\ ");
  return `\\begin{${environment}}${body}\\end{${environment}}`;
}

/** A recognised formula's LaTeX without its delimiters ($…$, $$…$$, \(…\), \[…\]); lines joined as rows. */
function bareFormulaLatex(latex: string): string {
  let source = latex.trim();
  const wrapped = /^\$\$([\s\S]*)\$\$$/.exec(source) ?? /^\$([\s\S]*)\$$/.exec(source) ??
    /^\\\[([\s\S]*)\\\]$/.exec(source) ?? /^\\\(([\s\S]*)\\\)$/.exec(source);
  if (wrapped) source = wrapped[1].trim();
  return source.replace(/\s*\n\s*/g, " ");
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
  // In an open formula the place is the formula's caret (its selection then), as in a formula row.
  const places = new Map<string, { doc: Node; from: number; to: number; formula?: { field: HTMLElement; selection: MathEditorInsertionTarget } }>();
  let placeCount = 0;
  const captureInsertionTarget = (): MathEditorInsertionTarget | null => {
    const view = document()?.view();
    if (!view) return null;
    const { from, to } = view.state.selection;
    const lineId = `document-place-${++placeCount}`;
    const open = openDocumentFormula();
    const formula = open
      ? {
          field: open.field,
          selection: {
            lineId,
            ranges: open.field.selection.ranges.map(([start, end]) => [start, end] as [number, number]),
            direction: open.field.selection.direction ?? "none",
          },
        }
      : undefined;
    places.set(lineId, { doc: view.state.doc, from, to, formula });
    return formula?.selection ?? { lineId, ranges: [[from, to]], direction: "forward" };
  };
  const insertLatexAt = (target: MathEditorInsertionTarget, given: string, source?: FormulaEditSource) => {
    // rows recognised by OCR: one formula of those rows (text with $…$ around formulas is pasted as it is)
    const latex = source === "ocr" && !/\$|\\\(|\\\[/.test(given) ? recognizedRowsLatex(given) : given;
    const view = document()?.view();
    if (!view) return false;
    const place = places.get(target.lineId);
    places.delete(target.lineId);
    const open = openDocumentFormula();
    if (place?.formula && open && open.field === place.formula.field && open.field.isConnected) {
      // into the formula, at its caret then: the formula's own LaTeX, without the delimiters of a recognised formula
      const bare = bareFormulaLatex(latex);
      if (!bare) return false;
      const { ranges, direction } = place.formula.selection;
      open.field.selection = { ranges, direction };
      insertLatexIntoFormula(open, bare);
      return true;
    }
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
    // the formula rows' history snapshots: the document keeps its own history
    captureDocumentSnapshot: () => getEditorDocumentSnapshot(),
    restoreDocumentSelection: () => false,
  };
}

interface Props {
  /** The workspace's editor handle: formula commands from the app insert through it. */
  editorRef: RefObject<MathEditorHandle | null>;
  /** An image pasted or dropped into the document: the app's OCR, the formula going in at `target`. */
  onPasteImage?: (file: File, target: MathEditorInsertionTarget) => void;
}

/**
 * Where the caret and the page were in each document opened in this run of the app: leaving document mode (for
 * formula mode) takes the editor down, and coming back the document opened with the caret at its start and the page
 * at its top. Kept per document: the caret and how far the page was scrolled (the same document lays out the same).
 */
const documentPlaces = new Map<string, { selection: unknown; scrollTop: number }>();

function scrollerOf(view: EditorView): HTMLElement | null {
  for (let node = view.dom.parentElement; node; node = node.parentElement) {
    const overflow = getComputedStyle(node).overflowY;
    if (overflow === "auto" || overflow === "scroll") return node;
  }
  return null;
}

function rememberPlace(id: string | null, view: EditorView | null | undefined) {
  if (!id || !view || view.isDestroyed || !view.dom.isConnected) return;
  const scroller = scrollerOf(view);
  documentPlaces.set(id, { selection: view.state.selection.toJSON(), scrollTop: scroller?.scrollTop ?? 0 });
}

function restorePlace(id: string, view: EditorView | null | undefined) {
  const place = documentPlaces.get(id);
  if (!place || !view || view.isDestroyed) return;
  try {
    let selection = Selection.fromJSON(view.state.doc, place.selection as Parameters<typeof Selection.fromJSON>[1]);
    // a formula that was selected or open: the caret after it (it is not reopened)
    if (selection instanceof NodeSelection) selection = TextSelection.near(view.state.doc.resolve(selection.to));
    view.dispatch(view.state.tr.setSelection(selection).setMeta("addToHistory", false));
  } catch {
    return;
  }
  // the page where it was — again while the page settles (the pages measured, the
  // formulas near it given their editors), until the user scrolls, clicks or types
  let stopped = false;
  const stop = () => {
    stopped = true;
    for (const name of ["wheel", "pointerdown", "keydown", "touchstart"]) window.removeEventListener(name, stop, true);
  };
  for (const name of ["wheel", "pointerdown", "keydown", "touchstart"]) window.addEventListener(name, stop, true);
  const settle = () => {
    if (stopped || view.isDestroyed || !view.dom.isConnected) return;
    const scroller = scrollerOf(view);
    if (!scroller) return;
    if (Math.abs(scroller.scrollTop - place.scrollTop) > 0.5) scroller.scrollTop = place.scrollTop;
  };
  requestAnimationFrame(settle);
  for (const delay of [150, 400, 800, 1300, 2000]) window.setTimeout(settle, delay);
  window.setTimeout(stop, 2100);
}

export function DocumentWorkspace({ editorRef, onPasteImage }: Props) {
  const documentRef = useRef<DocumentEditorHandle>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const language = useEditorStore((state) => state.language);
  const isEn = language === "en";
  const t = (en: string, zh: string) => (isEn ? en : zh);
  const [notice, setNotice] = useState<string | null>(null);
  const [tutorialOpen, setTutorialOpen] = useState(
    () => storage.getItem(TUTORIAL_SHOWN_KEY) !== "1" && !documentModeUsedBefore(),
  );
  useEffect(() => {
    storage.setItem(TUTORIAL_SHOWN_KEY, "1");
  }, []);
  const noticeTimer = useRef(0);
  const api = useMemo(documentLibrary, []);
  const libraryOpen = useDocumentLibraryStore((state) => state.open);
  const [library, setLibrary] = useState<Library | null>(null);
  const [currentId, setCurrentId] = useState<string | null>(null);
  // The open document: its file, the document last saved (or read) and the latest one.
  const current = useRef<{ file: DocumentFile | null; saved: Node | null; latest: Node | null }>({ file: null, saved: null, latest: null });
  const saveTimer = useRef(0);
  // saves run one after another: each is checked against the file the one before it wrote
  const saving = useRef<Promise<boolean>>(Promise.resolve(true));
  // The file changed on the disk while the document was edited here: nothing is saved until the user decides.
  const [conflict, setConflict] = useState(false);
  const conflictRef = useRef(false);
  const setConflicted = (value: boolean) => {
    conflictRef.current = value;
    setConflict(value);
  };
  // documents of earlier versions that stay .vtxdoc, shown once
  const [kept, setKept] = useState<Array<{ title: string; reason: string }> | null>(null);

  const adapterRef = useRef<MathEditorHandle | null>(null);
  useEffect(() => {
    const adapter = createDocumentEditorAdapter(() => documentRef.current);
    adapterRef.current = adapter;
    editorRef.current = adapter;
    // The headless tests drive the handle the way the app's tools and OCR do.
    if (import.meta.env.DEV) (window as unknown as { __documentAdapter?: MathEditorHandle }).__documentAdapter = adapter;
    return () => {
      if (editorRef.current === adapter) editorRef.current = null;
    };
  }, [editorRef]);

  const flash = (message: string) => {
    setNotice(message);
    window.clearTimeout(noticeTimer.current);
    noticeTimer.current = window.setTimeout(() => setNotice(null), 1800);
  };

  const entryOf = (id: string | null) => (id ? (library?.documents.find((doc) => doc.id === id) ?? library?.recent.find((doc) => doc.id === id)) : undefined);
  const withEntry = (entry: LibraryDocument) =>
    setLibrary((old) => {
      if (!old) return old;
      const put = (list: LibraryDocument[]) => [entry, ...list.filter((doc) => doc.id !== entry.id)].sort((a, b) => b.modified - a.modified);
      return entry.outside ? { ...old, recent: [entry, ...old.recent.filter((doc) => doc.id !== entry.id)] } : { ...old, documents: put(old.documents) };
    });

  const fileName = (ext: string, id = current.current.file?.entry.id ?? null) => {
    const entry = id === current.current.file?.entry.id ? current.current.file?.entry : entryOf(id);
    const base = (entry ? libraryDocumentName(entry, isEn) : "").trim().replace(/[\\/:*?"<>|]/g, "-").slice(0, 60) || t("Untitled document", "未命名文档");
    return base + ext;
  };

  // The open document is saved a moment after each change, and at once before another one opens.
  const onChange = useCallback((doc: Node) => {
    current.current.latest = doc;
    if (!current.current.file) return;
    window.clearTimeout(saveTimer.current);
    saveTimer.current = window.setTimeout(() => void ops.current.saveNow(), SAVE_DELAY_MS);
  }, []);

  /** The document as a new Markdown file's content, its pictures going along. */
  const asNewFile = (doc: Node, file: DocumentFile | null) =>
    newDocumentContent(doc, { formulaLatex: portableFormulaLatex, labels: file?.source?.labels }, file?.imagePaths ?? new Map()).content;

  const ops = useRef({
    saveNow: async (_force?: boolean): Promise<boolean> => true,
    open: async (_id: string, _reload?: boolean): Promise<boolean> => false,
    openPath: async (_path: string) => {},
    create: async (_folder: string | null) => {},
    importFile: async (_file: File) => {},
    saveAs: async () => {},
    exportFile: async (_id: string, _format?: DocumentExportFormat) => {},
    afterLoss: async (_library: Library) => {},
    checkDisk: async () => {},
    renamed: (_from: string, _to: string, _library: Library) => {},
  });
  const write = async (force: boolean, renamed = false): Promise<boolean> => {
    window.clearTimeout(saveTimer.current);
    const { file, saved, latest } = current.current;
    if (!file || !latest || (saved && saved.eq(latest) && !force)) return true;
    if (conflictRef.current && !force) return false;
    current.current.saved = latest;
    try {
      const content = documentContent(file, latest, { formulaLatex: portableFormulaLatex });
      const outcome = await api.save(file.entry.id, content, file.fingerprint, force);
      if (outcome.status === "changed") {
        // changed on the disk too: the user chooses which stays
        if (current.current.saved === latest) current.current.saved = saved;
        if (current.current.file === file) setConflicted(true);
        return false;
      }
      file.fingerprint = outcome.fingerprint;
      file.unwritten.clear();
      file.entry = outcome.entry;
      withEntry(outcome.entry);
      return true;
    } catch (error) {
      if (current.current.saved === latest) current.current.saved = saved;
      // another program put a file of a new picture's name beside the document meanwhile: the new pictures take
      // other names (nothing was written)
      const taken = PICTURE_NAME_TAKEN.exec(errorText(error))?.[1];
      if (taken && !renamed && renameNewPictures(file, taken)) return write(force, true);
      flash(t("Not saved: ", "未保存：") + errorText(error));
      return false;
    }
  };
  ops.current = {
    saveNow: (force = false) => {
      const next = saving.current.then(
        () => write(force),
        () => write(force),
      );
      saving.current = next;
      return next;
    },
    // `reload`: the file as it is on the disk now, what is open of it dropped
    open: async (id, reload = false) => {
      if (!reload && id === current.current.file?.entry.id) {
        documentRef.current?.focus();
        return true;
      }
      rememberPlace(current.current.file?.entry.id ?? null, documentRef.current?.view());
      if (!reload) await ops.current.saveNow();
      let file: DocumentFile;
      try {
        file = await loadDocumentFile(await api.read(id));
        documentRef.current?.open(file.doc);
      } catch (error) {
        const reason = errorText(error);
        flash(/LaTeX documents/.test(reason) ? t("LaTeX documents are no longer supported", "已不再支持 LaTeX 文档") : t("Cannot open: ", "无法打开：") + reason);
        return false;
      }
      const doc = documentRef.current?.view()?.state.doc ?? null;
      current.current = { file, saved: doc, latest: doc };
      setConflicted(false);
      setCurrentId(id);
      useDocumentLibraryStore.getState().setLastId(id);
      restorePlace(id, documentRef.current?.view());
      documentRef.current?.focus();
      return true;
    },
    // A file anywhere on the disk, in place; an archive or a .json is taken into the documents folder.
    openPath: async (path) => {
      try {
        if (/\.(zip|json)$/i.test(path)) {
          await ops.current.importFile(new File([(await readFileBytes(path)) as BlobPart], path.split(/[\\/]/).pop() ?? path));
          return;
        }
        await ops.current.saveNow();
        const entry = await api.openFile(path);
        setLibrary(await api.list());
        await ops.current.open(entry.id);
      } catch (error) {
        flash(t("Cannot open: ", "无法打开：") + errorText(error));
      }
    },
    create: async (folder) => {
      await ops.current.saveNow();
      try {
        const entry = await api.create(folder, untitledName(isEn), "markdown", { bytes: new Uint8Array(), memory: null, assets: [] });
        withEntry(entry);
        const store = useDocumentLibraryStore.getState();
        if (folder && store.collapsed.includes(folder)) store.toggleFolder(folder);
        await ops.current.open(entry.id);
      } catch (error) {
        flash(errorText(error));
      }
    },
    // Into the documents folder as a new document: a Markdown export's .zip (the text and the pictures it names), a
    // .json or .vtxdoc of an earlier version as it is; in a browser also a Markdown file (it cannot be edited in place).
    importFile: async (file) => {
      const stem = file.name.replace(/\.[^.]+$/, "");
      try {
        const bytes = new Uint8Array(await file.arrayBuffer());
        let content: DocumentContent;
        let kind: "markdown" | "vtxdoc" = "markdown";
        if (/\.zip$/i.test(file.name)) {
          const picture = /\.(png|jpe?g|gif|webp|bmp|svg|avif)$/i;
          const wanted = /\.(md|markdown)$/i;
          const entries = unzipSync(bytes, { filter: (entry) => wanted.test(entry.name) || picture.test(entry.name) });
          const markdownFile = findMarkdownFile(Object.entries(entries).filter(([path]) => !picture.test(path)).map(([path, data]) => ({ path, bytes: data })));
          if (!markdownFile) throw new Error(t("no .md file in the archive", "压缩包中没有 .md 文件"));
          const base = markdownFile.path.includes("/") ? markdownFile.path.slice(0, markdownFile.path.lastIndexOf("/") + 1) : "";
          // the pictures the text names, at the paths it names them by
          const assets: DocumentContent["assets"] = [];
          for (const path of picturePaths(decodeText(markdownFile.bytes))) {
            const data = entries[(base + path).replace(/\/\.\//g, "/")] ?? entries[path];
            if (data && !/^(?:[a-zA-Z][a-zA-Z0-9+.-]*:|\/)/.test(path) && !path.split("/").includes("..")) assets.push({ path: picturePath(path), bytes: data });
          }
          content = { bytes: markdownFile.bytes, memory: null, assets };
        } else if (/\.(md|markdown)$/i.test(file.name)) {
          content = { bytes, memory: null, assets: [] };
        } else {
          if (/\.(tex|ltx)$/i.test(file.name)) throw new Error(t("not a VisualTeX or Markdown file", "不是 VisualTeX 或 Markdown 文件"));
          parseDocument(documentSchema, decodeText(bytes));
          content = { bytes, memory: null, assets: [] };
          kind = "vtxdoc";
        }
        await ops.current.saveNow();
        const name = stem || untitledName(isEn);
        let entry: LibraryDocument;
        try {
          entry = await api.create(null, name, kind, content);
        } catch (error) {
          // pictures of the same names are in the documents folder already (another export's assets/fig-1.png…):
          // the document goes into a folder of its own with its pictures
          if (!PICTURE_NAME_TAKEN.test(errorText(error))) throw error;
          entry = await api.create((await api.createFolder(name)).id, name, kind, content);
        }
        setLibrary(await api.list());
        if (await ops.current.open(entry.id)) flash(t("Document imported", "文档已导入"));
      } catch (error) {
        flash(t("Cannot open: ", "无法打开：") + errorText(error));
      }
    },
    // The document as a Markdown file somewhere else, which becomes the open one.
    saveAs: async () => {
      const { file, latest } = current.current;
      if (!file || !latest) return;
      try {
        if (!isTauri()) {
          await saveTextFile(fileName(".md"), newDocumentContent(latest, { formulaLatex: portableFormulaLatex }).text, "md", "text/markdown", isEn);
          return;
        }
        const path = await chooseExportPath(fileName(".md"), "md", isEn);
        if (!path) return;
        await ops.current.saveNow();
        const entry = await api.saveAs(path, asNewFile(latest, file));
        setLibrary(await api.list());
        if (await ops.current.open(entry.id)) flash(t("Saved", "已保存"));
      } catch (error) {
        flash(t("Not saved: ", "未保存：") + errorText(error));
      }
    },
    exportFile: async (id, format = "markdown") => {
      try {
        // the formats come from the editor: the document is opened (its pages laid out) first
        if (id !== current.current.file?.entry.id) {
          await ops.current.open(id);
          await new Promise((resolve) => window.setTimeout(resolve, 400));
        }
        const run = { markdown: actions.current.exportMarkdown, pdf: actions.current.exportPdf }[format];
        run();
      } catch (error) {
        flash(errorText(error));
      }
    },
    // The open document is gone (deleted, forgotten, another documents folder): the newest other one opens, or a new one.
    afterLoss: async (next) => {
      current.current = { file: null, saved: null, latest: null };
      setConflicted(false);
      const other = next.documents.find((doc) => doc.trashed === null);
      if (other) await ops.current.open(other.id);
      else await ops.current.create(null);
    },
    // The file as it is on the disk now: when another program changed it and nothing was edited here, the document
    // is read again; when both changed, the user chooses.
    checkDisk: async () => {
      const { file, saved, latest } = current.current;
      if (!file || conflictRef.current) return;
      const state = await api.check(file.entry.id).catch(() => null);
      if (!state || !state.exists || state.fingerprint === file.fingerprint || current.current.file !== file) return;
      // (a save of its own may have finished meanwhile)
      await saving.current.catch(() => undefined);
      if (state.fingerprint === file.fingerprint || current.current.file !== file) return;
      if (saved && current.current.latest && !saved.eq(current.current.latest)) setConflicted(true);
      else if (latest === current.current.latest) await ops.current.open(file.entry.id, true);
    },
    // A document's path changed (renamed, moved): the open one and its remembered place follow.
    renamed: (from, to, next) => {
      setLibrary(next);
      const place = documentPlaces.get(from);
      if (place) {
        documentPlaces.delete(from);
        documentPlaces.set(to, place);
      }
      const file = current.current.file;
      if (!file || file.entry.id !== from) return;
      const entry = next.documents.find((doc) => doc.id === to) ?? next.recent.find((doc) => doc.id === to);
      if (entry) file.entry = entry;
      setCurrentId(to);
      useDocumentLibraryStore.getState().setLastId(to);
    },
  };

  // The documents of earlier versions become files once; then the library is read and the document last open
  // opens — or the files the system handed over.
  useEffect(() => {
    let live = true;
    void (async () => {
      try {
        const migration = await migrateLegacyLibrary(api, {
          formulaLatex: portableFormulaLatex,
          untitled: untitledName(isEn),
          differs: t("does not convert exactly", "无法精确转换"),
        });
        let lib = await api.list();
        // the one document kept before there was a library becomes the first one
        if (!lib.documents.length && !Object.keys(migration.documents).length && !storage.getItem(AUTOSAVE_TAKEN_KEY)) {
          const before = readAutosave();
          if (before) {
            try {
              const doc = documentSchema.nodeFromJSON(before);
              doc.check();
              await api.create(null, documentTitle(doc) || untitledName(isEn), "markdown", asNewFile(doc, null));
              lib = await api.list();
            } catch {
              // an unreadable autosave: nothing to take
            }
          }
        }
        storage.setItem(AUTOSAVE_TAKEN_KEY, "1");
        if (!live) return;
        setLibrary(lib);
        if (!migration.reported) setKept(keptDocuments(migration));
        const requested = takeOpenRequests();
        if (requested.length) {
          for (const path of requested) await ops.current.openPath(path);
          if (current.current.file) return;
        }
        const documents = lib.documents.filter((doc) => doc.trashed === null);
        const last = useDocumentLibraryStore.getState().lastId;
        const target = [...documents, ...lib.recent].find((doc) => doc.id === last) ?? documents[0];
        if (!target || !(await ops.current.open(target.id))) await ops.current.afterLoss(lib);
      } catch (error) {
        if (live) flash(t("Documents unavailable: ", "文档库不可用：") + errorText(error));
      }
    })();
    return () => {
      live = false;
    };
    // (read once: the language of a new document's name is the one at the start)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api]);

  // Files the system hands over while document mode is open (while the question about the open file is asked, they
  // wait for its answer).
  useEffect(() => {
    const take = () => {
      if (conflictRef.current) return;
      const requested = takeOpenRequests();
      void (async () => {
        for (const path of requested) await ops.current.openPath(path);
      })();
    };
    window.addEventListener(DOCUMENT_OPEN_FILES_EVENT, take);
    return () => window.removeEventListener(DOCUMENT_OPEN_FILES_EVENT, take);
  }, []);

  // Leaving document mode: the caret's and the page's place are kept for coming back (before the editor is taken
  // down: a layout effect's cleanup runs while it is still on the page).
  useLayoutEffect(
    () => () => rememberPlace(current.current.file?.entry.id ?? null, documentRef.current?.view()),
    [],
  );

  // Leaving document mode, the window or the app: what is not saved yet is saved now. Coming back to the window:
  // the file may have been changed by another program meanwhile, and documents added to or taken from the documents
  // folder.
  useEffect(() => {
    const flush = () => void ops.current.saveNow();
    const check = () => {
      void ops.current.checkDisk();
      void api.list().then(setLibrary, () => undefined);
    };
    window.addEventListener("blur", flush);
    window.addEventListener("pagehide", flush);
    window.addEventListener("beforeunload", flush);
    window.addEventListener("focus", check);
    return () => {
      window.removeEventListener("blur", flush);
      window.removeEventListener("pagehide", flush);
      window.removeEventListener("beforeunload", flush);
      window.removeEventListener("focus", check);
      flush();
    };
  }, [api]);

  // Open: the system's dialog in the app (the file is edited where it is), the file input in a browser
  const openFile = () => {
    if (!isTauri()) {
      fileRef.current?.click();
      return;
    }
    void chooseDocumentFile(isEn).then(
      (path) => {
        if (path) void ops.current.openPath(path);
      },
      (error) => flash(errorText(error)),
    );
  };

  const isCurrent = (id: string) => id === current.current.file?.entry.id;
  const libraryActions: DocumentLibraryActions = {
    open: (id) => void ops.current.open(id),
    create: (folder) => void ops.current.create(folder),
    openFile: () => openFile(),
    rename: (id, title) =>
      void (async () => {
        try {
          if (isCurrent(id)) await ops.current.saveNow();
          const done = await api.rename(id, title);
          ops.current.renamed(id, done.id, done.library);
        } catch (error) {
          flash(errorText(error));
        }
      })(),
    duplicate: (id) =>
      void (async () => {
        try {
          if (isCurrent(id)) await ops.current.saveNow();
          const source = entryOf(id);
          const entry = await api.duplicate(id, `${source ? libraryDocumentName(source, isEn) : ""} ${t("copy", "副本")}`.trim());
          setLibrary(await api.list());
          await ops.current.open(entry.id);
        } catch (error) {
          flash(errorText(error));
        }
      })(),
    move: (id, folder) =>
      void (async () => {
        try {
          if (isCurrent(id)) await ops.current.saveNow();
          const done = await api.move(id, folder);
          ops.current.renamed(id, done.id, done.library);
        } catch (error) {
          flash(errorText(error));
        }
      })(),
    trash: (id) =>
      void (async () => {
        try {
          if (isCurrent(id)) await ops.current.saveNow();
          const next = await api.trash(id);
          setLibrary(next);
          if (isCurrent(id)) await ops.current.afterLoss(next);
          flash(t("Moved to Recently deleted", "已移到最近删除"));
        } catch (error) {
          flash(errorText(error));
        }
      })(),
    restore: (id) => void api.restore(id).then(setLibrary, (error) => flash(errorText(error))),
    remove: (id) => void api.remove(id).then(setLibrary, (error) => flash(errorText(error))),
    emptyTrash: () => void api.emptyTrash().then(setLibrary, (error) => flash(errorText(error))),
    forget: (id) =>
      void (async () => {
        try {
          if (isCurrent(id)) await ops.current.saveNow();
          const next = await api.forget(id);
          setLibrary(next);
          if (isCurrent(id)) await ops.current.afterLoss(next);
        } catch (error) {
          flash(errorText(error));
        }
      })(),
    createFolder: async () => {
      try {
        const { library: next, id } = await api.createFolder(t("New folder", "新建文件夹"));
        setLibrary(next);
        return id;
      } catch (error) {
        flash(errorText(error));
        return null;
      }
    },
    renameFolder: (id, name) =>
      void (async () => {
        try {
          const open = current.current.file?.entry;
          if (open && open.folder === id) await ops.current.saveNow();
          const next = await api.renameFolder(id, name);
          // the open document's path changes with its folder's
          const inside = open && !open.outside && open.id.startsWith(`${id}/`) ? open.id : null;
          const moved = inside ? next.documents.find((doc) => doc.title === open!.title && doc.id.endsWith(inside.slice(id.length)) && !next.documents.some((other) => other.id === inside)) : undefined;
          if (inside && moved) ops.current.renamed(inside, moved.id, next);
          else setLibrary(next);
        } catch (error) {
          flash(errorText(error));
        }
      })(),
    deleteFolder: (id) =>
      void (async () => {
        try {
          const open = current.current.file?.entry;
          const inside = open && !open.outside && open.id.startsWith(`${id}/`);
          if (inside) await ops.current.saveNow();
          const known = new Set(library?.documents.map((doc) => doc.id));
          const next = await api.deleteFolder(id);
          setLibrary(next);
          // (the open document moved to the documents folder itself, under its name or "name 2": opened again there)
          if (inside) {
            const moved = next.documents.find((doc) => doc.trashed === null && !doc.folder && !known.has(doc.id) && (doc.title === open!.title || doc.title.startsWith(`${open!.title} `)));
            current.current = { file: null, saved: null, latest: null };
            if (moved) await ops.current.open(moved.id);
            else await ops.current.afterLoss(next);
          }
        } catch (error) {
          flash(errorText(error));
        }
      })(),
    exportFile: (id, format) => void ops.current.exportFile(id, format),
    reveal: isTauri() ? (id) => void invoke("document_library_reveal", { id }).catch((error) => flash(errorText(error))) : null,
    // Another documents folder: its documents are listed as they are; the open one is saved first.
    chooseRoot: isTauri()
      ? () =>
          void (async () => {
            try {
              const path = await chooseFolder(isEn, library?.root);
              if (!path || path === library?.root) return;
              await ops.current.saveNow();
              const next = await api.setRoot(path);
              setLibrary(next);
              await ops.current.afterLoss(next);
            } catch (error) {
              flash(errorText(error));
            }
          })()
      : null,
    search: (query) => api.search(query),
  };

  // The file changed on the disk and here: read it again, write this one over it, or keep this one as a copy.
  const resolveConflict = (choice: "reload" | "overwrite" | "copy") =>
    void (async () => {
      const { file, latest } = current.current;
      if (!file || !latest) return;
      try {
        if (choice === "reload") await ops.current.open(file.entry.id, true);
        else if (choice === "overwrite") {
          setConflicted(false);
          if (await ops.current.saveNow(true)) flash(t("Saved", "已保存"));
        } else {
          const name = `${libraryDocumentName(file.entry, isEn)} ${t("copy", "副本")}`;
          const content = file.kind === "vtxdoc" ? documentContent(file, latest) : asNewFile(latest, file);
          let entry: LibraryDocument;
          if (file.entry.outside && isTauri()) {
            const path = await chooseExportPath(`${name}.${file.kind === "vtxdoc" ? "vtxdoc" : "md"}`, file.kind === "vtxdoc" ? "vtxdoc" : "md", isEn);
            if (!path) return;
            entry = await api.saveAs(path, content);
          } else entry = await api.create(file.entry.folder, name, file.kind, content);
          // the copy has what was edited here; the file itself stays as the other program left it
          current.current = { file: null, saved: null, latest: null };
          setConflicted(false);
          setLibrary(await api.list());
          await ops.current.open(entry.id);
        }
      } catch (error) {
        flash(errorText(error));
      }
      // files the system handed over meanwhile
      window.dispatchEvent(new Event(DOCUMENT_OPEN_FILES_EVENT));
    })();

  // PDF: one at a time
  const pdfBusy = useRef(false);

  const actions = useRef({
    save: () => {},
    open: () => {},
    create: () => {},
    saveAs: () => {},
    exportMarkdown: () => {},
    copyMarkdown: () => {},
    exportPdf: () => {},
  });
  // (New, Open and Save as wait while the question about the open file is asked: its edits would be dropped)
  actions.current = {
    save: () => void ops.current.saveNow().then((done) => done && flash(t("Saved", "已保存"))),
    open: () => {
      if (!conflictRef.current) openFile();
    },
    create: () => {
      if (!conflictRef.current) void ops.current.create(current.current.file?.entry.folder ?? null);
    },
    saveAs: () => {
      if (!conflictRef.current) void ops.current.saveAs();
    },
    exportMarkdown: () => {
      const markdown = documentRef.current?.toMarkdown();
      if (!markdown) return;
      const assets = documentRef.current?.imageFiles("assets") ?? [];
      if (assets.length) {
        const zip = zipSync({ [fileName(".md")]: strToU8(markdown), ...Object.fromEntries(assets.map((file) => [file.path, file.bytes])) });
        void saveBinaryFile(fileName(".zip"), zip, "zip", "application/zip", isEn).then(
          (done) => done && flash(t("Markdown exported", "Markdown 已导出")),
          (error) => flash(errorText(error)),
        );
        return;
      }
      void saveTextFile(fileName(".md"), markdown, "md", "text/markdown", isEn).then(
        (done) => done && flash(t("Markdown exported", "Markdown 已导出")),
        (error) => flash(errorText(error)),
      );
    },
    copyMarkdown: () => {
      const markdown = documentRef.current?.toMarkdown();
      if (!markdown) return;
      void navigator.clipboard.writeText(markdown).then(
        () => flash(t("Markdown copied", "Markdown 已复制")),
        () => flash(t("Copy failed", "复制失败")),
      );
    },
    exportPdf: () => {
      const view = documentRef.current?.view();
      if (!view || pdfBusy.current) return;
      pdfBusy.current = true;
      void (async () => {
        try {
          // in a browser (the web edition): its print dialog ("Save as PDF")
          if (!isTauri()) {
            await printDocumentInBrowser(view, fileName(""));
            return;
          }
          const path = await chooseExportPath(fileName(".pdf"), "pdf", isEn);
          if (!path) return;
          flash(t("Exporting PDF…", "正在导出 PDF…"));
          await documentToPdf(view, path, fileName(""));
          flash(t("PDF exported", "PDF 已导出"));
        } catch (error) {
          flash(errorText(error));
        } finally {
          pdfBusy.current = false;
        }
      })();
    },
  };

  // The main menu and Ctrl+S / Ctrl+O / Ctrl+N reach the document through window events.
  useEffect(() => {
    const bind: Array<[string, () => void]> = [
      [DOCUMENT_SAVE_REQUEST_EVENT, () => actions.current.save()],
      [DOCUMENT_OPEN_REQUEST_EVENT, () => actions.current.open()],
      [DOCUMENT_NEW_REQUEST_EVENT, () => actions.current.create()],
      [DOCUMENT_SAVE_AS_EVENT, () => actions.current.saveAs()],
      [DOCUMENT_EXPORT_MARKDOWN_EVENT, () => actions.current.exportMarkdown()],
      [DOCUMENT_COPY_MARKDOWN_EVENT, () => actions.current.copyMarkdown()],
      [DOCUMENT_EXPORT_PDF_EVENT, () => actions.current.exportPdf()],
      [DOCUMENT_TUTORIAL_EVENT, () => setTutorialOpen(true)],
    ];
    for (const [name, fn] of bind) window.addEventListener(name, fn);
    return () => {
      for (const [name, fn] of bind) window.removeEventListener(name, fn);
    };
  }, []);

  return (
    <section className={"vtx-document-workspace" + (libraryOpen ? " has-library" : "")} aria-label={t("Document", "文档")}>
      {libraryOpen && <DocumentLibrary isEn={isEn} library={library} currentId={currentId} actions={libraryActions} />}
      <div className="vtx-document-main">
        {/* the document's own buttons above it; the formula tools stay in the bottom dock */}
        <DocumentTextToolbar
          leading={libraryOpen ? null : <DocumentLibraryToggle isEn={isEn} />}
          trailing={
            <button
              type="button"
              className="icon-button compact formula-formatting-button vtx-tutorial-button"
              title={t("Tutorials", "教程")}
              aria-label={t("Tutorials", "教程")}
              onClick={() => setTutorialOpen(true)}
            >
              <GraduationCap size={16} />
            </button>
          }
        />
        <DocumentEditor
          ref={documentRef}
          onChange={onChange}
          onNotice={flash}
          onImage={
            onPasteImage
              ? (file) => {
                  const target = adapterRef.current?.captureInsertionTarget();
                  if (target) onPasteImage(file, target);
                }
              : undefined
          }
        />
      </div>
      {notice && <div className="mc-notice" role="status">{notice}</div>}
      {conflict && (
        <DocumentFileDialog
          title={t("The file changed on disk", "文件已在磁盘上更改")}
          lines={[current.current.file?.entry.path || libraryDocumentName(current.current.file?.entry ?? { title: "" }, isEn)]}
          actions={[
            { label: t("Reload", "重新载入"), run: () => resolveConflict("reload") },
            { label: t("Overwrite", "覆盖"), run: () => resolveConflict("overwrite") },
            { label: t("Save a copy", "另存副本"), run: () => resolveConflict("copy"), primary: true },
          ]}
        />
      )}
      {kept && kept.length > 0 && !conflict && (
        <DocumentFileDialog
          title={t("Kept as .vtxdoc", "仍为 .vtxdoc")}
          lines={kept.map((doc) => `${doc.title} — ${/newer than this VisualTeX/.test(doc.reason) ? t("made by a newer VisualTeX", "由更新版本的 VisualTeX 创建") : doc.reason}`)}
          actions={[
            {
              label: t("OK", "确定"),
              primary: true,
              run: () => {
                setKept(null);
                void api.legacy.state().then((state) => state && api.legacy.setState({ ...(state as object), reported: true }));
              },
            },
          ]}
        />
      )}
      <TutorialDialog open={tutorialOpen} isEn={isEn} onClose={() => setTutorialOpen(false)} />
      <input
        ref={fileRef}
        type="file"
        accept=".vtxdoc,.json,.md,.markdown,.zip"
        className="visually-hidden"
        onChange={(event) => {
          const file = event.target.files?.[0];
          event.target.value = "";
          if (file) void ops.current.importFile(file);
        }}
      />
    </section>
  );
}
