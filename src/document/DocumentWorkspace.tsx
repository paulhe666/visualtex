// The main window's document mode (docs/document-editor/DESIGN.md §6.4): the document editor in the workspace's
// editor area beside the document library (saved as you type), the file actions of the main menu (new, import, save,
// export) and an editor handle so that formula commands from the rest of the app (tools, tiles, OCR) go into the
// document.
import { useCallback, useEffect, useMemo, useRef, useState, type RefObject } from "react";
import type { Node } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import { invoke, isTauri } from "@tauri-apps/api/core";
import type { EditorView } from "prosemirror-view";
import { strToU8, unzipSync, zipSync } from "fflate";
import { GraduationCap } from "lucide-react";
import {
  decodeText,
  documentSchema,
  findMarkdownFile,
  graphicsPaths,
  setImageFileResolver,
  hasMathDelimiters,
  insertDisplayLatex,
  insertInlineFormula,
  insertInlineLatex,
  mathTextToSlice,
  parseDocument,
  parseMathText,
  readLatexProject,
  serializeDocument,
} from "../../shared/visualtex-document/src/index.ts";
import type { MathEditorHandle, MathEditorInsertionTarget } from "../editor/MathEditor";
import type { FormulaEditSource } from "../history/historyTypes";
import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../editor/alignmentMarkers";
import { createSafeStorage } from "../runtime/safeStorage";
import { useEditorStore } from "../stores/editorStore";
import type { LatexCommand } from "../types/command";
import { DocumentTextToolbar } from "./DocumentControls";
import { DocumentEditor, type DocumentEditorHandle } from "./DocumentEditor";
import { DocumentLibrary, DocumentLibraryToggle, type DocumentExportFormat, type DocumentLibraryActions } from "./DocumentLibrary";
import {
  documentLibrary,
  documentTitle,
  libraryDocumentName,
  saveTextFile,
  saveBinaryFile,
  chooseExportPath,
  chooseImportFile,
  useDocumentLibraryStore,
  type Library,
  type LibraryDocument,
} from "./documentLibraryApi";
import { insertCommandIntoFormula, insertLatexIntoFormula, waitForOpenFormula } from "./formulaInsert";
import { openDocumentFormula } from "./openDocumentFormula";
import { documentToPdf, printDocumentInBrowser } from "./documentPdf";
import { imagePathKey, loadImportImages } from "./importImages";
import {
  DOCUMENT_COPY_LATEX_EVENT,
  DOCUMENT_COPY_MARKDOWN_EVENT,
  DOCUMENT_EXPORT_FILE_EVENT,
  DOCUMENT_EXPORT_LATEX_EVENT,
  DOCUMENT_EXPORT_MARKDOWN_EVENT,
  DOCUMENT_EXPORT_PDF_EVENT,
  DOCUMENT_NEW_REQUEST_EVENT,
  DOCUMENT_OPEN_REQUEST_EVENT,
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

function emptyDocumentText(): string {
  return serializeDocument(documentSchema.topNodeType.createAndFill()!);
}

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
  // The open document: its id, the text last saved (or read) and the latest text.
  const current = useRef<{ id: string | null; saved: Node | null; latest: Node | null }>({ id: null, saved: null, latest: null });
  const saveTimer = useRef(0);

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

  const entryOf = (id: string | null) => (id ? library?.documents.find((doc) => doc.id === id) : undefined);
  const withEntry = (entry: LibraryDocument) =>
    setLibrary((old) => old && { ...old, documents: [entry, ...old.documents.filter((doc) => doc.id !== entry.id)].sort((a, b) => b.modified - a.modified) });

  const fileName = (ext: string, id = current.current.id) => {
    const entry = entryOf(id);
    const base = (entry ? libraryDocumentName(entry, isEn) : "").trim().replace(/[\\/:*?"<>|]/g, "-").slice(0, 60) || t("Untitled document", "未命名文档");
    return base + ext;
  };

  // The open document is saved a moment after each change, and at once before another one opens.
  const onChange = useCallback((doc: Node) => {
    current.current.latest = doc;
    if (!current.current.id) return;
    window.clearTimeout(saveTimer.current);
    saveTimer.current = window.setTimeout(() => void ops.current.saveNow(), SAVE_DELAY_MS);
  }, []);

  const ops = useRef({
    saveNow: async (): Promise<boolean> => true,
    open: async (_id: string): Promise<boolean> => false,
    create: async (_folder: string | null) => {},
    importFile: async (_file: File, _folder?: string | null) => {},
    exportFile: async (_id: string, _format?: DocumentExportFormat) => {},
    afterLoss: async (_library: Library) => {},
  });
  ops.current = {
    saveNow: async () => {
      window.clearTimeout(saveTimer.current);
      const { id, saved, latest } = current.current;
      if (!id || !latest || (saved && saved.eq(latest))) return true;
      current.current.saved = latest;
      try {
        withEntry(await api.save(id, serializeDocument(latest), documentTitle(latest)));
        return true;
      } catch (error) {
        if (current.current.saved === latest) current.current.saved = saved;
        flash(t("Not saved: ", "未保存：") + errorText(error));
        return false;
      }
    },
    open: async (id) => {
      if (id === current.current.id) {
        documentRef.current?.focus();
        return true;
      }
      await ops.current.saveNow();
      try {
        const text = await api.read(id);
        documentRef.current?.load(text);
      } catch (error) {
        flash(t("Cannot open: ", "无法打开：") + errorText(error));
        return false;
      }
      const doc = documentRef.current?.view()?.state.doc ?? null;
      current.current = { id, saved: doc, latest: doc };
      setCurrentId(id);
      useDocumentLibraryStore.getState().setLastId(id);
      documentRef.current?.focus();
      return true;
    },
    create: async (folder) => {
      await ops.current.saveNow();
      try {
        const entry = await api.create(folder, emptyDocumentText(), null);
        withEntry(entry);
        const store = useDocumentLibraryStore.getState();
        if (folder && store.collapsed.includes(folder)) store.toggleFolder(folder);
        await ops.current.open(entry.id);
      } catch (error) {
        flash(errorText(error));
      }
    },
    // A .vtxdoc (or .json) as it is; Markdown and LaTeX read into a new document (one Ctrl+Z from empty); a .zip of
    // a LaTeX project (its main file, the files it inputs, the macros of its class and packages) or of Markdown.
    importFile: async (file, folder = null) => {
      const stem = file.name.replace(/\.[^.]+$/, "");
      try {
        const bytes = new Uint8Array(await file.arrayBuffer());
        let kind: "document" | "markdown" | "latex";
        let text: string;
        // the archive's images, for the paths its text gives them (figures/x.png, assets/x.png)
        let images: Map<string, Uint8Array> | null = null;
        let base = "";
        if (/\.zip$/i.test(file.name)) {
          const picture = /\.(png|jpe?g|gif|webp|bmp|svg|avif)$/i;
          const entries = unzipSync(bytes, { filter: (entry) => /\.(tex|cls|sty|def|ltx|bib|md|markdown)$/i.test(entry.name) || picture.test(entry.name) });
          images = new Map(Object.entries(entries).filter(([path]) => picture.test(path)));
          const files = Object.entries(entries).filter(([path]) => !picture.test(path)).map(([path, data]) => ({ path, bytes: data }));
          const project = readLatexProject(files);
          const markdownFile = project ? null : findMarkdownFile(files);
          if (project) {
            kind = "latex";
            text = project.source;
            base = project.main.includes("/") ? project.main.slice(0, project.main.lastIndexOf("/") + 1) : "";
          } else if (markdownFile) {
            kind = "markdown";
            text = decodeText(markdownFile.bytes);
            base = markdownFile.path.includes("/") ? markdownFile.path.slice(0, markdownFile.path.lastIndexOf("/") + 1) : "";
          } else throw new Error(t("no .tex or .md file in the archive", "压缩包中没有 .tex 或 .md 文件"));
        } else {
          text = decodeText(bytes);
          kind = /\.(md|markdown)$/i.test(file.name) ? "markdown" : /\.(tex|ltx)$/i.test(file.name) ? "latex" : "document";
          if (kind === "document") parseDocument(documentSchema, text);
        }
        // the pictures it names: from the web, the file's folder, its archive (before it is read)
        const pictures =
          kind === "document"
            ? new Map<string, string>()
            : await loadImportImages(text, kind, { archive: images, base, folder, folders: kind === "latex" ? graphicsPaths(text) : [] });
        await ops.current.saveNow();
        const entry = await api.create(null, kind === "document" ? text : emptyDocumentText(), stem);
        withEntry(entry);
        if (!(await ops.current.open(entry.id))) return;
        if (pictures.size) setImageFileResolver((path) => pictures.get(imagePathKey(path)) ?? null);
        try {
          if (kind === "markdown") documentRef.current?.loadMarkdown(text);
          else if (kind === "latex") documentRef.current?.loadLatex(text);
        } finally {
          setImageFileResolver(null);
        }
        await ops.current.saveNow();
        flash(t("Document imported", "文档已导入"));
      } catch (error) {
        flash(t("Cannot open: ", "无法打开：") + errorText(error));
      }
    },
    exportFile: async (id, format = "vtxdoc") => {
      try {
        if (format !== "vtxdoc") {
          // the other formats come from the editor: the document is opened (its pages laid out) first
          if (id !== current.current.id) {
            await ops.current.open(id);
            await new Promise((resolve) => window.setTimeout(resolve, 400));
          }
          const run = { latex: actions.current.exportLatex, markdown: actions.current.exportMarkdown, pdf: actions.current.exportPdf }[format];
          run();
          return;
        }
        let text: string;
        if (id === current.current.id) {
          await ops.current.saveNow();
          text = documentRef.current?.serialize() ?? "";
        } else text = await api.read(id);
        if (await saveTextFile(fileName(".vtxdoc", id), text, "vtxdoc", "application/json", isEn)) flash(t("Document exported", "文档已导出"));
      } catch (error) {
        flash(errorText(error));
      }
    },
    // The open document was deleted: the newest other one opens, or a new one.
    afterLoss: async (next) => {
      current.current = { id: null, saved: null, latest: null };
      const other = next.documents.find((doc) => doc.trashed === null);
      if (other) await ops.current.open(other.id);
      else await ops.current.create(null);
    },
  };

  // The library is read once; the document last open opens. On the first start with the library the document kept
  // before it becomes the first one.
  useEffect(() => {
    let live = true;
    void (async () => {
      try {
        let lib = await api.list();
        if (!lib.documents.some((doc) => doc.trashed === null)) {
          let text = emptyDocumentText();
          let title = "";
          const kept = lib.documents.length || storage.getItem(AUTOSAVE_TAKEN_KEY) ? undefined : readAutosave();
          if (kept) {
            try {
              const doc = documentSchema.nodeFromJSON(kept);
              doc.check();
              text = serializeDocument(doc);
              title = documentTitle(doc);
            } catch {
              // an unreadable autosave: an empty document
            }
          }
          const entry = await api.create(null, text, null);
          if (title) await api.save(entry.id, text, title);
          storage.setItem(AUTOSAVE_TAKEN_KEY, "1");
          lib = await api.list();
        }
        if (!live) return;
        setLibrary(lib);
        const documents = lib.documents.filter((doc) => doc.trashed === null);
        const last = useDocumentLibraryStore.getState().lastId;
        await ops.current.open((documents.find((doc) => doc.id === last) ?? documents[0]).id);
      } catch (error) {
        if (live) flash(t("Documents unavailable: ", "文档库不可用：") + errorText(error));
      }
    })();
    return () => {
      live = false;
    };
  }, [api]);

  // Leaving document mode, the window or the app: what is not saved yet is saved now.
  useEffect(() => {
    const flush = () => void ops.current.saveNow();
    window.addEventListener("blur", flush);
    window.addEventListener("pagehide", flush);
    window.addEventListener("beforeunload", flush);
    return () => {
      window.removeEventListener("blur", flush);
      window.removeEventListener("pagehide", flush);
      window.removeEventListener("beforeunload", flush);
      flush();
    };
  }, []);

  // Open: the system's dialog in the app (the file's folder then finds its pictures), the file input in a browser
  const openFile = () => {
    if (!isTauri()) {
      fileRef.current?.click();
      return;
    }
    void chooseImportFile(isEn).then(
      (chosen) => chosen && ops.current.importFile(chosen.file, chosen.folder),
      (error) => flash(errorText(error)),
    );
  };

  const libraryActions: DocumentLibraryActions = {
    open: (id) => void ops.current.open(id),
    create: (folder) => void ops.current.create(folder),
    importFile: () => openFile(),
    rename: (id, title) => void api.rename(id, title).then(setLibrary, (error) => flash(errorText(error))),
    duplicate: (id) =>
      void (async () => {
        try {
          if (id === current.current.id) await ops.current.saveNow();
          const source = entryOf(id);
          const entry = await api.duplicate(id, `${source ? libraryDocumentName(source, isEn) : ""} ${t("copy", "副本")}`.trim());
          withEntry(entry);
          await ops.current.open(entry.id);
        } catch (error) {
          flash(errorText(error));
        }
      })(),
    move: (id, folder) => void api.move(id, folder).then(setLibrary, (error) => flash(errorText(error))),
    trash: (id) =>
      void (async () => {
        try {
          if (id === current.current.id) await ops.current.saveNow();
          const next = await api.trash(id);
          setLibrary(next);
          if (id === current.current.id) await ops.current.afterLoss(next);
          flash(t("Moved to Recently deleted", "已移到最近删除"));
        } catch (error) {
          flash(errorText(error));
        }
      })(),
    restore: (id) => void api.restore(id).then(setLibrary, (error) => flash(errorText(error))),
    remove: (id) => void api.remove(id).then(setLibrary, (error) => flash(errorText(error))),
    emptyTrash: () => void api.emptyTrash().then(setLibrary, (error) => flash(errorText(error))),
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
    renameFolder: (id, name) => void api.renameFolder(id, name).then(setLibrary, (error) => flash(errorText(error))),
    deleteFolder: (id) => void api.deleteFolder(id).then(setLibrary, (error) => flash(errorText(error))),
    exportFile: (id, format) => void ops.current.exportFile(id, format),
    reveal: isTauri() ? (id) => void invoke("document_library_reveal", { id }).catch((error) => flash(errorText(error))) : null,
    search: (query) => api.search(query),
  };

  // PDF: one at a time
  const pdfBusy = useRef(false);

  const actions = useRef({
    save: () => {},
    open: () => {},
    create: () => {},
    exportDocument: () => {},
    exportLatex: () => {},
    copyLatex: () => {},
    exportMarkdown: () => {},
    copyMarkdown: () => {},
    exportPdf: () => {},
  });
  actions.current = {
    save: () => void ops.current.saveNow().then((done) => done && flash(t("Saved", "已保存"))),
    open: () => openFile(),
    create: () => void ops.current.create(entryOf(current.current.id)?.folder ?? null),
    exportDocument: () => current.current.id && void ops.current.exportFile(current.current.id),
    exportLatex: () => {
      const tex = documentRef.current?.toLatex();
      if (!tex) return;
      // with images: the .tex and its figures/ in one .zip (Overleaf's upload)
      const figures = documentRef.current?.imageFiles("figures") ?? [];
      if (figures.length) {
        const zip = zipSync({ [fileName(".tex")]: strToU8(tex), ...Object.fromEntries(figures.map((file) => [file.path, file.bytes])) });
        void saveBinaryFile(fileName(".zip"), zip, "zip", "application/zip", isEn).then(
          (done) => done && flash(t("LaTeX exported", "LaTeX 已导出")),
          (error) => flash(errorText(error)),
        );
        return;
      }
      void saveTextFile(fileName(".tex"), tex, "tex", "text/x-tex", isEn).then(
        (done) => done && flash(t("LaTeX exported", "LaTeX 已导出")),
        (error) => flash(errorText(error)),
      );
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
          // in a browser: its print dialog ("Save as PDF")
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
    copyLatex: () => {
      const tex = documentRef.current?.toLatex();
      if (!tex) return;
      void navigator.clipboard.writeText(tex).then(
        () => flash(t("LaTeX copied", "LaTeX 已复制")),
        () => flash(t("Copy failed", "复制失败")),
      );
    },
  };

  // The main menu and Ctrl+S / Ctrl+O / Ctrl+N reach the document through window events.
  useEffect(() => {
    const bind: Array<[string, () => void]> = [
      [DOCUMENT_SAVE_REQUEST_EVENT, () => actions.current.save()],
      [DOCUMENT_OPEN_REQUEST_EVENT, () => actions.current.open()],
      [DOCUMENT_NEW_REQUEST_EVENT, () => actions.current.create()],
      [DOCUMENT_EXPORT_FILE_EVENT, () => actions.current.exportDocument()],
      [DOCUMENT_EXPORT_LATEX_EVENT, () => actions.current.exportLatex()],
      [DOCUMENT_COPY_LATEX_EVENT, () => actions.current.copyLatex()],
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
      <TutorialDialog open={tutorialOpen} isEn={isEn} onClose={() => setTutorialOpen(false)} />
      <input
        ref={fileRef}
        type="file"
        accept=".vtxdoc,.json,.md,.markdown,.tex,.ltx,.zip"
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
