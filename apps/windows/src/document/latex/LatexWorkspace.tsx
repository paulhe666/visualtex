// A LaTeX document (personal edition): the project's files are the document, and it is edited on its pages as
// TeX makes them — SVG from the same DVI/XDV as the PDF, so the screen is the export. A click puts the caret
// between the glyphs clicked between; typing goes into the source there and the paragraph is typeset at once by
// the live server (the document's own TeX, the preamble already read) and drawn over its old lines
// (visualEditor.ts, livePara.ts); "$" opens the formula editor in place, "\" the LaTeX commands. TeX makes the
// whole document again a moment after the typing stops; the pages change when the new ones are there. The source
// (a tab a file) is a view of its own (the toolbar's switch): there a click on a page goes to its source line.
// TeX's errors are listed and marked on their lines; the last good pages stay.
import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore, type MutableRefObject, type ReactNode } from "react";
import type { ChangeSet } from "@codemirror/state";
import { gunzipSync, strToU8, zipSync } from "fflate";
import { isTauri } from "@tauri-apps/api/core";
import {
  AlertTriangle,
  ChevronDown,
  CircleX,
  Code2,
  Download,
  FileText,
  Loader2,
  Minus,
  Play,
  Plus,
  Sigma,
  X,
} from "lucide-react";
import { decodeText } from "../../../../shared/visualtex-document/src/index.ts";
import type { FormulaSessionCallbacks } from "../../../../shared/visualtex-document/src/index.ts";
import { chooseExportPath, saveBinaryFile, saveTextFile } from "../documentLibraryApi";
import { DocumentFormulaField, type DocumentFormulaFieldApi } from "../DocumentFormulaField";
import { setPageFormulaOpener } from "../openDocumentFormula";
import { createSafeStorage } from "../../runtime/safeStorage";
import { LatexPages, type PageContent, type PageMarker } from "./LatexPages";
import { FormulaOnPage, KeyboardSink, LatexCommandMenu, PageMarks } from "./LatexVisual";
import { VisualEditor, type TypesetOutcome } from "./visualEditor";
import { documentCommands } from "./texCaret";
import { LatexSource, type LatexSourceHandle, type LineMark } from "./LatexSource";
import {
  beginsParagraphCommand,
  bodyPrefix,
  changedParagraph,
  docMacros,
  endsProof,
  enumerated,
  proofOpening,
  proofText,
  isListItem,
  listContext,
  listText,
  liveEligible,
  liveText,
  overlayOf,
  pageGlyphs,
  paragraphArea,
  placesOf,
  SP_PER_BP,
  verify,
  type DocMacros,
  type LiveOverlay,
  type PageGlyph,
  type ParagraphChange,
  lineNumberAt,
  type ParagraphArea,
  type ParagraphPlace,
} from "./livePara";
import { formulaForEditor, mathRegionAt, type MathRegion } from "./mathRegion";
import { projectPath, SyncTex } from "./synctex";
import { parseTexLog, type TexMessage } from "./texLog";
import { base64ToBytes, texService, type CompileResult, type LiveResult, type PageMeta, type ProjectFile, type TexService } from "./texService";
import { projectTitle, type LatexPointer } from "./latexDocs";
import "./latex.css";

export interface LatexActions {
  save(): Promise<boolean>;
  exportPdf(): void;
  exportProject(): void;
  exportMain(): void;
}

interface Props {
  docId: string;
  pointer: LatexPointer;
  isEn: boolean;
  leading?: ReactNode;
  trailing?: ReactNode;
  /** the pointer changed (open tabs, engine): the library keeps it */
  onPointer(pointer: LatexPointer): void;
  /** the main file's \title changed */
  onTitle(title: string): void;
  flash(message: string): void;
  actionsRef: MutableRefObject<LatexActions | null>;
}

interface Buffer {
  /** the text on disk (as last read or written) */
  disk: string;
  text: string;
}

const storage = createSafeStorage();
const ZOOM_KEY = "visualtex.latex.zoom.v1";
const VIEW_KEY = "visualtex.latex.view.v1";
const SAVE_DELAY_MS = 700;
// TeX starts at the first keystroke (a run in progress is not stopped, the next one follows it); while the
// paragraph being typed is shown live, once typing pauses (the live lines are what TeX will make; the run then
// settles the page breaks and references)
const COMPILE_DELAY_MS = 0;
const LIVE_COMPILE_DELAY_MS = 400;
// a formula or command half typed (TeX stops at it): TeX runs only after a longer pause
const UNFINISHED_COMPILE_DELAY_MS = 1500;
const ENGINES = ["xelatex", "pdflatex", "lualatex"] as const;
const ENGINE_LABEL: Record<string, string> = { xelatex: "XeLaTeX", pdflatex: "pdfLaTeX", lualatex: "LuaLaTeX" };

function isUtf8(bytes: Uint8Array): boolean {
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    return true;
  } catch {
    return false;
  }
}

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export function LatexWorkspace({ docId, pointer, isEn, leading, trailing, onPointer, onTitle, flash, actionsRef }: Props) {
  const t = (en: string, zh: string) => (isEn ? en : zh);
  const service = useMemo<TexService | null>(() => texService(), []);
  const sourceRef = useRef<LatexSourceHandle>(null);
  const pointerRef = useRef(pointer);
  pointerRef.current = pointer;
  const buffers = useRef(new Map<string, Buffer>());
  /** each file's text as first read from disk (what a build made before the file was open used) */
  const pristine = useRef(new Map<string, string>());
  const legacy = useRef(new Set<string>());
  const [files, setFiles] = useState<ProjectFile[]>([]);
  const [tabs, setTabs] = useState<string[]>(() => pointer.tabs?.length ? pointer.tabs : [pointer.main]);
  const [active, setActive] = useState<string>(pointer.main);
  const tabsRef = useRef(tabs);
  tabsRef.current = tabs;
  const [problem, setProblem] = useState<string | null>(null);
  const [fileMenu, setFileMenu] = useState(false);
  const [exportMenu, setExportMenu] = useState(false);
  const [engineMenu, setEngineMenu] = useState(false);

  // ---- pages ----
  const [pages, setPages] = useState<PageMeta[]>([]);
  const [contents, setContents] = useState<Array<PageContent | undefined>>([]);
  const [build, setBuild] = useState(0);
  const [result, setResult] = useState<CompileResult | null>(null);
  const [busy, setBusy] = useState(false);
  const [messages, setMessages] = useState<TexMessage[]>([]);
  const [showMessages, setShowMessages] = useState(false);
  const [, setMarker] = useState<PageMarker | null>(null);
  const [reveal, setReveal] = useState<{ page: number; y: number; nonce: number; ifHidden?: boolean } | null>(null);
  const [zoomSetting, setZoomSetting] = useState<number | "fit">(() => {
    const z = storage.getItem(ZOOM_KEY);
    return z && z !== "fit" ? Number(z) || "fit" : "fit";
  });
  const [paneWidth, setPaneWidth] = useState(800);
  const [caretInMath, setCaretInMath] = useState(false);
  /** the pages edited (the visual editor), or the source */
  const [view, setViewState] = useState<"visual" | "source">(() => (storage.getItem(VIEW_KEY) === "source" ? "source" : "visual"));
  const viewRef = useRef(view);
  viewRef.current = view;

  const cache = useRef(new Map<string, PageContent>());
  const buildRef = useRef(0);
  const pagesRef = useRef<PageMeta[]>([]);
  const contentsRef = useRef<Array<PageContent | undefined>>([]);
  contentsRef.current = contents;
  const visibleRef = useRef<number[]>([1, 2]);
  const loading = useRef(new Set<string>());
  const sync = useRef<{ build: number; tex: SyncTex; files: Array<string | null>; work: string } | null>(null);

  const maxWidth = pages.reduce((m, p) => Math.max(m, p.width), 0) || 595;
  const zoom = zoomSetting === "fit" ? Math.max(0.3, Math.min(3, (paneWidth - 36) / maxWidth)) : zoomSetting;
  const setZoom = (z: number | "fit") => {
    setZoomSetting(z);
    storage.setItem(ZOOM_KEY, String(z));
  };

  // ---- files ----
  const readFile = useCallback(
    async (path: string): Promise<string> => {
      const known = buffers.current.get(path);
      if (known) return known.text;
      const bytes = base64ToBytes(await service!.read(pointerRef.current.root, path));
      const text = decodeText(bytes);
      // a file that is not UTF-8 (GB18030 from older CTeX projects) is shown, not written back: saving would
      // change its encoding under \documentclass[GBK]
      if (!isUtf8(bytes)) legacy.current.add(path);
      buffers.current.set(path, { disk: text, text });
      pristine.current.set(path, text);
      return text;
    },
    [service],
  );

  const saveTimers = useRef(new Map<string, number>());
  const saveFile = useCallback(
    async (path: string) => {
      window.clearTimeout(saveTimers.current.get(path));
      saveTimers.current.delete(path);
      const b = buffers.current.get(path);
      if (!b || b.text === b.disk || !service) return true;
      const text = b.text;
      try {
        await service.write(pointerRef.current.root, path, { text });
        b.disk = text;
        return true;
      } catch (error) {
        flash(t("Not saved: ", "未保存：") + errorText(error));
        return false;
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [service],
  );
  const saveAll = useCallback(async () => {
    let ok = true;
    for (const path of [...buffers.current.keys()]) ok = (await saveFile(path)) && ok;
    return ok;
  }, [saveFile]);

  // ---- compiling ----
  const run = useRef({ inflight: false, pending: false, started: 0, expected: 1500, reruns: 0, timer: 0, generation: 0, rerun: false });
  /** the texts the shown build was made from (the live paragraph is found against them) */
  const built = useRef<{ build: number; texts: Map<string, string>; seq: number } | null>(null);
  /** each file's changes since the texts of a build (in order, numbered): an offset of the build's text is found
   *  in the text now through them, and back */
  const changeLog = useRef(new Map<string, Array<{ seq: number; changes: ChangeSet }>>());
  const seqRef = useRef(0);
  const toCurrent = useCallback((path: string, pos: number, assoc: -1 | 1) => {
    const since = built.current?.seq ?? 0;
    let p = pos;
    for (const e of changeLog.current.get(path) ?? []) if (e.seq > since) p = e.changes.mapPos(p, assoc);
    return p;
  }, []);
  const toBuilt = useCallback((path: string, pos: number, assoc: -1 | 1) => {
    const since = built.current?.seq ?? 0;
    const log = changeLog.current.get(path) ?? [];
    let p = pos;
    for (let i = log.length - 1; i >= 0; i--) if (log[i].seq > since) p = log[i].changes.invertedDesc.mapPos(p, assoc);
    return p;
  }, []);

  const toProject = useCallback((path: string) => {
    const work = sync.current?.work ?? "";
    return projectPath(path, work, pointerRef.current.main, ["_vtx_preamble.tex", "_vtx_body.tex"]);
  }, []);

  const loadSync = useCallback(
    async (buildId: number) => {
      if (!service || sync.current?.build === buildId) return;
      try {
        const out = await service.synctex(docId, buildId);
        const text = new TextDecoder().decode(gunzipSync(base64ToBytes(out.base64)));
        const tex = SyncTex.parse(text);
        const filesOf = tex.files.map((f) => projectPath(f, out.work, pointerRef.current.main, out.aliases));
        sync.current = { build: buildId, tex, files: filesOf, work: out.work };
      } catch {
        // no SyncTeX for this build
      }
    },
    [docId, service],
  );

  const contentOf = (p: { svg: string | null; png: string | null; hash: string }): PageContent | undefined =>
    p.svg ? { kind: "svg", hash: p.hash, svg: p.svg } : p.png ? { kind: "png", hash: p.hash, url: `data:image/png;base64,${p.png}` } : undefined;

  /** The pages near the view that do not show their current content yet. */
  const loadVisible = useCallback(async () => {
    if (!service) return;
    const buildId = buildRef.current;
    const metas = pagesRef.current;
    const wanted = visibleRef.current.filter((n) => {
      const meta = metas[n - 1];
      if (!meta) return false;
      const shown = contentsRef.current[n - 1];
      if (shown?.hash === meta.hash) return false;
      const cached = cache.current.get(meta.hash);
      if (cached) {
        setContents((old) => {
          const next = old.slice();
          next[n - 1] = cached;
          return next;
        });
        return false;
      }
      return !loading.current.has(`${buildId}:${n}`);
    });
    if (!wanted.length) return;
    for (const n of wanted) loading.current.add(`${buildId}:${n}`);
    try {
      const out = await service.pages(docId, buildId, wanted, Math.min(4, Math.max(1, zoomRef.current * (window.devicePixelRatio || 1))));
      if (buildRef.current !== buildId) return;
      setContents((old) => {
        const next = old.slice();
        for (const p of out) {
          const c = contentOf(p);
          if (!c) continue;
          cache.current.set(p.hash, c);
          next[p.page - 1] = c;
        }
        return next;
      });
    } catch {
      // the build is gone (a newer one): the next one loads
    } finally {
      for (const n of wanted) loading.current.delete(`${buildId}:${n}`);
    }
  }, [docId, service]);
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;

  const applyResult = useCallback(
    (r: CompileResult) => {
      setResult(r);
      // a tool that failed beside TeX (Biber, BibTeX, the PDF driver): said above the pages
      setProblem(r.error || null);
      const parsed = parseTexLog(r.log, toProject);
      setMessages(parsed);
      if (r.status !== "ok" || !r.pages.length) {
        if (parsed.some((m) => m.kind === "error")) setShowMessages(true);
        return;
      }
      for (const s of r.svgs) cache.current.set(s.hash, { kind: "svg", hash: s.hash, svg: s.svg });
      buildRef.current = r.build;
      pagesRef.current = r.pages;
      setBuild(r.build);
      setPages(r.pages);
      // a page shows its new content if it is here, else what it showed (no blank pages while loading)
      setContents((old) => r.pages.map((p, i) => cache.current.get(p.hash) ?? old[i]));
      void loadSync(r.build);
      // the main file's title names the document
      const main = buffers.current.get(pointerRef.current.main)?.text;
      if (main) onTitle(projectTitle(main, pointerRef.current.main));
    },
    [loadSync, onTitle, toProject],
  );

  const compile = useCallback(
    async (full = false): Promise<CompileResult | null> => {
      if (!service) return null;
      const state = run.current;
      window.clearTimeout(state.timer);
      if (state.inflight && !full) {
        // a run that has only begun on a long document gives way; else this edit waits for it
        const elapsed = performance.now() - state.started;
        if (!(state.expected > 2500 && elapsed < state.expected * 0.3)) {
          state.pending = true;
          return null;
        }
      }
      const overlays: Record<string, string> = {};
      for (const [path, b] of buffers.current) if (b.text !== b.disk) overlays[path] = b.text;
      const texts = new Map([...buffers.current].map(([path, b]) => [path, b.text] as const));
      const seqAtStart = seqRef.current;
      const generation = ++state.generation;
      state.inflight = true;
      state.started = performance.now();
      setBusy(true);
      try {
        const p = pointerRef.current;
        // the pages in view as soon as TeX has written them (a long document's end comes later)
        const ticket = `${docId}:${generation}:${Date.now()}`;
        if (!full)
          void service.preview(docId, ticket).then(
            (early) => {
              if (early.status !== "early" || generation !== state.generation || !state.inflight) return;
              for (const s of early.svgs) cache.current.set(s.hash, { kind: "svg", hash: s.hash, svg: s.svg });
              setContents((old) => {
                const next = old.slice();
                for (const s of early.svgs) if (s.page - 1 < next.length) next[s.page - 1] = cache.current.get(s.hash);
                return next;
              });
            },
            () => undefined,
          );
        const r = await service.compile({
          ticket,
          doc: docId,
          root: p.root,
          main: p.main,
          engine: p.engine,
          shellEscape: p.shellEscape,
          overlays,
          visible: visibleRef.current.slice(0, 6),
          full,
        });
        if (r.status === "superseded") return null;
        if (generation === state.generation || full) {
          state.expected = r.timings.total || state.expected;
          state.rerun = r.status === "ok" && r.needsRerun;
          if (r.status === "ok" && r.pages.length) {
            built.current = { build: r.build, texts, seq: seqAtStart };
            // (the changes the build has: no longer needed)
            for (const [path, log] of changeLog.current) changeLog.current.set(path, log.filter((e) => e.seq > seqAtStart));
          }
          applyResult(r);
          if (r.status === "notex") setProblem(t("TeX is not installed", "未安装 TeX"));
        }
        return r;
      } catch (error) {
        setProblem(errorText(error));
        return null;
      } finally {
        if (generation === state.generation) {
          state.inflight = false;
          setBusy(false);
          if (state.pending) {
            state.pending = false;
            state.reruns = 0;
            void compile();
          } else if (state.rerun && state.reruns < 3) {
            state.rerun = false;
            state.reruns++;
            void compile();
          } else state.reruns = 0;
        }
      }
    },
    [applyResult, docId, service],
  );

  const rootRef = useRef<HTMLDivElement>(null);
  const scheduleCompile = useCallback(
    (delay = COMPILE_DELAY_MS) => {
      window.clearTimeout(run.current.timer);
      // (a run is due: the workspace says so, for the tests)
      rootRef.current?.setAttribute("data-due", "true");
      run.current.timer = window.setTimeout(() => {
        rootRef.current?.removeAttribute("data-due");
        void compile();
      }, delay);
    },
    [compile],
  );
  const scheduleCompileRef = useRef(scheduleCompile);
  scheduleCompileRef.current = scheduleCompile;

  // ---- opening ----
  useEffect(() => {
    let live = true;
    if (!service) {
      setProblem(t("LaTeX documents need the VisualTeX app", "LaTeX 文档需要在 VisualTeX 应用中打开"));
      return;
    }
    void (async () => {
      try {
        const status = await service.status();
        if (!live) return;
        if (!status.found) {
          setProblem(t("TeX is not installed", "未安装 TeX"));
          return;
        }
        const info = await service.open(pointer.root, pointer.main);
        if (!live) return;
        setFiles(info.files);
        const text = await readFile(pointer.main);
        if (!live) return;
        sourceRef.current?.show(pointer.main, text, legacy.current.has(pointer.main));
        setActive(pointer.main);
        void compile();
      } catch (error) {
        if (live) setProblem(errorText(error));
      }
    })();
    return () => {
      live = false;
      void saveAll();
      void service?.close(docId);
    };
    // one document per mount (the workspace is keyed by the document)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const openTab = useCallback(
    async (path: string, line?: number) => {
      try {
        const text = await readFile(path);
        sourceRef.current?.show(path, text, legacy.current.has(path));
        setActive(path);
        if (!tabsRef.current.includes(path)) {
          const next = [...tabsRef.current, path];
          tabsRef.current = next;
          setTabs(next);
          onPointer({ ...pointerRef.current, tabs: next });
        }
        if (line) requestAnimationFrame(() => sourceRef.current?.goto(line));
      } catch (error) {
        flash(t("Cannot open: ", "无法打开：") + errorText(error));
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [readFile, onPointer],
  );

  const closeTab = (path: string) => {
    if (path === pointer.main) return;
    void saveFile(path);
    const next = tabs.filter((p) => p !== path);
    setTabs(next);
    onPointer({ ...pointerRef.current, tabs: next });
    if (active === path) void openTab(pointer.main);
  };

  // marks of the active file
  useEffect(() => {
    const marks: LineMark[] = messages.filter((m) => m.file === active && m.line > 0).map((m) => ({ line: m.line, kind: m.kind, message: m.message }));
    sourceRef.current?.mark(marks);
  }, [messages, active]);

  // ---- the live paragraph ----
  const [, setLive] = useState<LiveOverlay | null>(null);
  const liveState = useRef({
    busy: false,
    pending: null as { path: string; text: string } | null,
    /** a paragraph of a build checked against its page: how it is indented, its lines and where its box starts */
    verified: new Map<string, { indent: number; place: ParagraphPlace; x0: number; y0: number; tx: number; ty: number; result: LiveResult; hsize: number } | null>(),
    /** glyph outlines by "<font>:<code>", as the server sent them to this session */
    paths: new Map<string, string>(),
    session: `${docId}:${Math.random().toString(36).slice(2)}`,
    /** the paragraph shown live */
    shown: "",
    glyphs: new Map<string, PageGlyph[]>(),
    areas: new Map<string, ParagraphArea | null>(),
    /** the paragraph typed is not TeX yet (a formula or command half typed) */
    unfinished: false,
  });
  const macrosRef = useRef<{ preamble: string; macros: DocMacros } | null>(null);
  /** The document's own commands (its main file's preamble). */
  const macrosOf = useCallback((): DocMacros => {
    const main = buffers.current.get(pointerRef.current.main)?.text ?? "";
    const at = main.search(/\\begin\s*\{document\}/);
    const preamble = at >= 0 ? main.slice(0, at) : "";
    if (macrosRef.current?.preamble !== preamble) macrosRef.current = { preamble, macros: docMacros(preamble) };
    return macrosRef.current.macros;
  }, []);
  const clearLive = useCallback(() => {
    liveState.current.shown = "";
    setLive(null);
  }, []);

  /**
   * What the live server makes of a change of a paragraph of `path` (`text` its text now): checked first, on the
   * paragraph's text as the shown build had it, to come out where the page has it (else nothing is drawn); then
   * the new text typeset there. `transform`: what the paragraph is sent as (a formula being edited: its room).
   */
  const typesetChange = useCallback(
    async (path: string, text: string, change: ParagraphChange, options: { displays?: boolean; transform?: (paragraph: string) => string } = {}): Promise<TypesetOutcome> => {
      const state = liveState.current;
      const base = built.current;
      if (!service || !base || base.build !== buildRef.current) return { ok: false, why: "no build text", retry: !base };
      const macros = macrosOf();
      const main = path === pointerRef.current.main;
      if (!liveEligible(change, text, main, options.displays)) return { ok: false, why: "not a plain paragraph" };
      // the body's declarations before it (a size, a family), then the paragraph with its indent
      const prefix = bodyPrefix(text, change.start, main);
      if (sync.current?.build !== base.build) await loadSync(base.build);
      const s = sync.current;
      if (!s || s.build !== base.build || buildRef.current !== base.build) return { ok: false, why: "no SyncTeX", retry: true };
      const key = `${base.build}:${path}:${change.oldFirst}:${change.oldLast}`;
      let area = state.areas.get(key);
      if (area === undefined) {
        const files = s.files.map((f, i) => (f === path ? i : -1)).filter((i) => i >= 0);
        area = paragraphArea(s.tex, files, change.oldFirst, change.oldLast);
        if (state.areas.size > 200) state.areas.clear();
        state.areas.set(key, area);
      }
      if (!area) return { ok: false, why: "not found on the pages" };
      const meta = pagesRef.current[area.page - 1];
      const shown = contentsRef.current[area.page - 1];
      // the page as this build made it (its glyphs are what the paragraph is checked against)
      if (!meta || !shown || shown.hash !== meta.hash) return { ok: false, why: "page not shown", retry: true };
      let glyphs: PageGlyph[] | null = null;
      if (shown.kind === "svg") {
        glyphs = state.glyphs.get(shown.hash) ?? pageGlyphs(shown.svg);
        if (state.glyphs.size > 40) state.glyphs.clear();
        state.glyphs.set(shown.hash, glyphs);
      }
      // a list item: typeset in its lists (the text's width; the list makes its lines' width and indent)
      const item = isListItem(change.oldText) && isListItem(change.newText);
      const list = item ? listContext(text, change.start, main) : null;
      if (item && !list) return { ok: false, why: "not in a list" };
      const numbered = list ? enumerated(list.name, macrosRef.current?.preamble ?? "") : false;
      // the line width: the lines' own; a list item's lines are the list's (its margins inside the width given):
      // the width of the box the lines are in, made right by one answer when the box is not the list's (a list in
      // a frame, a minipage)
      let hsize = Math.round((list ? area.width : area.lines[area.core[0]].width) * SP_PER_BP);
      const lineWidth = area.lines[area.core[0]].width;
      if (import.meta.env.DEV) {
        const w = window as unknown as { __latexLiveDebug?: unknown[] };
        w.__latexLiveDebug = [...(w.__latexLiveDebug ?? []).slice(-5), { key, list: list ? { begins: list.begins, before: list.before } : null, hsize: hsize / SP_PER_BP, lineWidth, core: area.core, lines: area.lines.map((l) => [l.h, l.v, l.width]) }];
      }
      // a proof's first paragraph: in its proof (its head); a proof's last: with its QED symbol (\end{proof} puts it)
      const proof = list ? null : proofOpening(text, change);
      const last = endsProof(text, change);
      const qed = last && !proof ? " \\qed" : "";
      const ask = async (paragraph: string, indent: number, size?: number) => {
        const sent = list
          ? listText(prefix, list, numbered, paragraph + qed)
          : proof
            ? proofText(prefix, proof, paragraph, last)
            : liveText(prefix, indent, paragraph + qed);
        const r = await service.live({ doc: docId, text: sent, hsize: size ?? hsize, indent: -2, session: state.session });
        if (import.meta.env.DEV) (window as unknown as { __latexLiveSent?: unknown }).__latexLiveSent = { sent, hsize: (size ?? hsize) / SP_PER_BP, map: r.map, lines: r.lines };
        for (const [k, d] of Object.entries(r.paths)) state.paths.set(k, d);
        return r;
      };
      let checked = state.verified.get(key);
      if (checked === undefined) {
        checked = null;
        // indented as the document indents it, or not (after a heading, \noindent; one of the document's paragraph
        // commands sets its own)
        const indents = list || proof ? [0] : beginsParagraphCommand(change.oldText, macros) ? [0, -1] : [-1, 0];
        let unindented: LiveResult | null = null;
        let measuredTried = 0;
        for (let k = 0; k < indents.length; k++) {
          const indent = indents[k];
          let r0 = await ask(change.oldText, indent);
          if (indent === 0) unindented = r0;
          // no server yet: asked again at the next edit
          if (r0.status === "starting" || r0.status === "unavailable") return { ok: false, why: "server " + r0.status, retry: r0.status === "starting" };
          if (list && r0.status === "ok" && r0.lines.length && Math.abs(r0.lines[0][3] - lineWidth) > 0.01) {
            // the list's lines narrower or wider than the page's: the box it is in is not the one SyncTeX gave
            hsize = Math.round(hsize + (lineWidth - r0.lines[0][3]) * SP_PER_BP);
            r0 = await ask(change.oldText, indent);
          }
          // its lines among the area's (the one run where its glyphs, or on a picture of the page its glues, are)
          const places = placesOf(area, r0.lines.length);
          for (const place of places) {
            const report = import.meta.env.DEV ? { matched: 0, onPage: 0, misses: [] as [number, number][] } : undefined;
            const at = verify(r0, place, glyphs, report);
            if (import.meta.env.DEV) {
              const w = window as unknown as { __latexLiveDebug?: unknown[] };
              w.__latexLiveDebug = [...(w.__latexLiveDebug ?? []).slice(-5), { key, indent, place, status: r0.status, error: r0.error, lines: r0.lines, count: r0.glyphs.length, report, ok: !!at }];
            }
            if (at) {
              checked = { indent, place, ...at, result: r0, hsize };
              break;
            }
          }
          if (checked) break;
          // indented neither as the document's paragraphs nor not at all (a list's own paragraph indent, a proof's
          // later paragraphs): as much as the page shows — its first glyph's place against the unindented answer's
          const first = (map: LiveResult["map"], baseline: number) =>
            Math.min(...(map ?? []).filter((g) => Math.abs(g[1] - baseline) < 0.5).map((g) => g[0]));
          if (k === indents.length - 1 && measuredTried < 2 && glyphs && unindented?.status === "ok" && unindented.lines.length && !list && !proof) {
            for (const place of placesOf(area, unindented.lines.length)) {
              const b = place.boxes[0];
              const onPage = Math.min(...glyphs.filter((g) => Math.abs(g.y - b.v) < 0.3 && g.x >= b.h - 3 && g.x <= b.h + b.width).map((g) => g.x));
              // first a kern as wide as the page's first glyph is in; then less what the glyph itself stands off it
              // in that answer
              let measured = onPage - b.h;
              if (measuredTried === 1 && r0.status === "ok" && r0.lines.length && indent > 0) measured = indent / SP_PER_BP - (first(r0.map, r0.lines[0][0]) - r0.lines[0][4] - indent / SP_PER_BP);
              if (!isFinite(measured) || measured < 0.2 || measured > 80 || (indent > 0 && Math.abs(measured * SP_PER_BP - indent) < 300)) continue;
              measuredTried++;
              indents.push(Math.round(measured * SP_PER_BP));
              break;
            }
          }
        }
        if (state.verified.size > 200) state.verified.clear();
        state.verified.set(key, checked);
      }
      if (!checked) return { ok: false, why: "does not match the page" };
      const unchanged = change.newText === change.oldText && !options.transform;
      if (import.meta.env.DEV) (window as unknown as { __latexLiveWhy?: string }).__latexLiveWhy = unchanged ? "the page is current" : "";
      const r = unchanged ? checked.result : await ask(options.transform ? options.transform(change.newText) : change.newText, checked.indent, checked.hsize);
      // TeX for the whole document once what is typed is TeX again (or after a longer pause)
      const unfinished = r.status === "error";
      if (unfinished !== state.unfinished) {
        state.unfinished = unfinished;
        scheduleCompileRef.current(unfinished ? UNFINISHED_COMPILE_DELAY_MS : LIVE_COMPILE_DELAY_MS);
      }
      if (buildRef.current !== base.build) return { ok: false, why: "a new build", retry: true };
      if (r.status === "ok" && r.undrawable === 0 && r.lines.length) {
        if (import.meta.env.DEV) (window as unknown as { __latexLive?: unknown }).__latexLive = { key, texMs: r.texMs, ms: r.ms, lines: r.lines.length, at: performance.now() };
        // its first line on the old first line's baseline (TeX keeps the baselines' distance; a taller first line —
        // a formula's room — reaches up, it does not push the paragraph down)
        const y0 = unchanged ? checked.y0 : checked.place.boxes[0].v - r.lines[0][0];
        return {
          ok: true,
          result: r,
          page: checked.place.page,
          x0: checked.x0,
          y0,
          tx: checked.tx,
          ty: checked.ty,
          overlay: unchanged ? null : overlayOf(r, checked.place, checked.x0, y0, state.paths, checked.tx, checked.ty),
        };
      }
      if (import.meta.env.DEV) (window as unknown as { __latexLiveWhy?: string }).__latexLiveWhy = `live ${r.status} ${r.error}`;
      return { ok: false, why: `live ${r.status} ${r.error}`, unfinished };
    },
    [docId, loadSync, macrosOf, service],
  );

  /** The paragraph `path`'s edit changed (the source edited), typeset by the live server and drawn over its old lines. */
  const liveFor = useCallback(
    async (path: string, text: string): Promise<void> => {
      const state = liveState.current;
      // (why there is none: for the tests)
      const quit = (why: string) => {
        if (import.meta.env.DEV) (window as unknown as { __latexLiveWhy?: string }).__latexLiveWhy = why;
        clearLive();
      };
      const base = built.current;
      // the text the shown build was made from (a file opened since: as it was on disk)
      const old = base?.texts.get(path) ?? pristine.current.get(path);
      if (!service || !base || old === undefined || base.build !== buildRef.current) return quit("no build text");
      if (old === text) return quit("the page is current");
      const change = changedParagraph(old, text, macrosOf());
      if (!change) return quit("not a plain paragraph");
      const key = `${base.build}:${path}:${change.oldFirst}:${change.oldLast}`;
      const out = await typesetChange(path, text, change);
      if (buildRef.current !== base.build) return;
      if (out.ok) {
        state.shown = key;
        if (import.meta.env.DEV) (window as unknown as { __latexLiveWhy?: string }).__latexLiveWhy = "";
        setLive(out.overlay);
      } else if (!out.unfinished || state.shown !== key) {
        // a command being typed (TeX's error) keeps what was shown; anything else shows the page
        quit(out.why);
      } else if (import.meta.env.DEV) {
        (window as unknown as { __latexLiveWhy?: string }).__latexLiveWhy = out.why;
      }
    },
    [clearLive, macrosOf, service, typesetChange],
  );

  const requestLive = useRef<(path: string, text: string) => void>(() => undefined);
  requestLive.current = (path: string, text: string) => {
    const state = liveState.current;
    if (state.busy) {
      state.pending = { path, text };
      return;
    }
    state.busy = true;
    void liveFor(path, text)
      .catch(() => clearLive())
      .finally(() => {
        state.busy = false;
        const next = state.pending;
        state.pending = null;
        if (next) requestLive.current(next.path, next.text);
      });
  };

  // a new build shown: the paragraph again against it (or nothing, when it has the text)
  useEffect(() => {
    if (viewRef.current === "visual") return;
    const path = sourceRef.current?.active();
    const text = path ? buffers.current.get(path)?.text : undefined;
    if (!path || text === undefined || (built.current?.texts.get(path) ?? pristine.current.get(path)) === text) {
      clearLive();
      return;
    }
    requestLive.current(path, text);
  }, [build, contents, clearLive]);

  // ---- editing ----
  const caretTimer = useRef(0);
  const placeOfRef = useRef<(path: string, line: number) => { page: number; x: number; y: number; width: number; height: number } | null>(() => null);
  const placeOf = (path: string, line: number) => {
    const s = sync.current;
    if (!s) return null;
    for (let i = 0; i < s.files.length; i++) {
      if (s.files[i] !== path) continue;
      const found = s.tex.view(i, line);
      if (found) return found;
    }
    return null;
  };
  placeOfRef.current = placeOf;

  const onChange = useCallback(
    (path: string, text: string) => {
      const b = buffers.current.get(path);
      if (b) b.text = text;
      else buffers.current.set(path, { disk: "", text });
      window.clearTimeout(saveTimers.current.get(path));
      saveTimers.current.set(path, window.setTimeout(() => void saveFile(path), SAVE_DELAY_MS));
      // (on the pages the editor typesets its paragraph and says when TeX runs)
      if (viewRef.current === "visual") return;
      // a paragraph the live server can show: TeX once typing pauses
      const base = built.current;
      const old = base ? base.texts.get(path) ?? pristine.current.get(path) : undefined;
      const change = base && old !== undefined ? changedParagraph(old, text, macrosOf()) : null;
      const live =
        !!change &&
        liveEligible(change, text, path === pointerRef.current.main) &&
        liveState.current.verified.get(`${base!.build}:${path}:${change.oldFirst}:${change.oldLast}`) !== null;
      scheduleCompile(live ? (liveState.current.unfinished ? UNFINISHED_COMPILE_DELAY_MS : LIVE_COMPILE_DELAY_MS) : COMPILE_DELAY_MS);
      requestLive.current(path, text);
      const caret = sourceRef.current?.caret();
      if (caret) {
        const at = placeOf(path, caret.line);
        if (at) setMarker({ ...at, kind: "pending" });
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [saveFile, scheduleCompile],
  );

  const onCaret = useCallback((path: string, line: number, pos: number) => {
    window.clearTimeout(caretTimer.current);
    caretTimer.current = window.setTimeout(() => {
      const text = sourceRef.current?.text(path) ?? "";
      setCaretInMath(!!mathRegionAt(text, pos));
      if (run.current.inflight) return;
      const at = placeOf(path, line);
      if (!at) return;
      setMarker({ ...at, kind: "sync" });
      setReveal({ page: at.page, y: at.y, nonce: Date.now(), ifHidden: true });
    }, 160);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // a new build: the caret's place again
  useEffect(() => {
    const caret = sourceRef.current?.caret();
    if (!caret) {
      setMarker(null);
      return;
    }
    const at = placeOf(caret.path, caret.line);
    setMarker(at ? { ...at, kind: "sync" } : null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [build]);

  // ---- the pages edited ----
  const editorRef = useRef<VisualEditor | null>(null);
  const flashRef = useRef(flash);
  flashRef.current = flash;
  const isEnRef = useRef(isEn);
  isEnRef.current = isEn;
  const overlayTick = useRef(0);
  const [, setOverlayTick] = useState(0);
  if (!editorRef.current) {
    editorRef.current = new VisualEditor({
      source: () => sourceRef.current,
      textOf: (path) => buffers.current.get(path)?.text,
      openFile: async (path) => {
        try {
          const text = await readFile(path);
          if (sourceRef.current?.active() !== path) sourceRef.current?.show(path, text, legacy.current.has(path));
          setActive(path);
          return !legacy.current.has(path);
        } catch {
          return false;
        }
      },
      built: () => built.current,
      builtText: (path) => built.current?.texts.get(path) ?? pristine.current.get(path),
      toCurrent: (path, pos, assoc) => toCurrent(path, pos, assoc),
      toBuilt: (path, pos, assoc) => toBuilt(path, pos, assoc),
      sync: async () => {
        await loadSync(buildRef.current);
        const s = sync.current;
        return s && s.build === buildRef.current ? { tex: s.tex, files: s.files } : null;
      },
      pageSvg: (page) => {
        const meta = pagesRef.current[page - 1];
        const shown = contentsRef.current[page - 1];
        return meta && shown && shown.hash === meta.hash && shown.kind === "svg" ? shown.svg : null;
      },
      pageCount: () => pagesRef.current.length,
      pageHeight: (page) => pagesRef.current[page - 1]?.height ?? 842,
      macros: () => macrosOf(),
      commands: () => {
        const main = buffers.current.get(pointerRef.current.main)?.text ?? "";
        const at = main.search(/\\begin\s*\{document\}/);
        return documentCommands(at >= 0 ? main.slice(0, at) : "");
      },
      typeset: (path, text, change, options) => typesetChangeRef.current(path, text, change, { displays: true, ...options }),
      overlaysChanged: () => setOverlayTick(++overlayTick.current),
      cannotEdit: () => flashRef.current(isEnRef.current ? "Not editable here on the page · ⌘/ source" : "此处不能在页面上编辑 · ⌘/ 源码"),
      revealSource: (path, offset) => {
        const text = buffers.current.get(path)?.text;
        if (text === undefined) return;
        const at = placeOfRef.current(path, lineNumberAt(text, offset));
        if (at) setReveal({ page: at.page, y: at.y, nonce: Date.now(), ifHidden: true });
      },
      scheduleCompile: (live) => scheduleCompileRef.current(live ? (liveState.current.unfinished ? UNFINISHED_COMPILE_DELAY_MS : LIVE_COMPILE_DELAY_MS) : COMPILE_DELAY_MS),
      isMac: /Mac|iPhone|iPad/.test(navigator.platform),
    });
  }
  const editor = editorRef.current;
  if (import.meta.env.DEV) (window as unknown as { __latexEditor?: VisualEditor }).__latexEditor = editor;
  const typesetChangeRef = useRef(typesetChange);
  typesetChangeRef.current = typesetChange;
  useSyncExternalStore(editor.subscribe, editor.getVersion);
  useEffect(() => {
    editor.revive();
    // the formula tiles and toolbar: a new formula at the caret on the pages
    setPageFormulaOpener(() => {
      if (viewRef.current !== "visual" || !editor.caret || editor.formula) return false;
      editor.newFormula(false);
      return true;
    });
    return () => {
      setPageFormulaOpener(null);
      editor.dispose();
    };
  }, [editor]);
  const onChanges = useCallback(
    (path: string, changes: ChangeSet) => {
      const log = changeLog.current.get(path) ?? [];
      log.push({ seq: ++seqRef.current, changes });
      changeLog.current.set(path, log);
      editor.changed(path, changes);
    },
    [editor],
  );
  // a new build: the units the caret is in again from it
  useEffect(() => {
    if (view === "visual") editor.newBuild();
  }, [build, editor, view]);
  const sinkRef = useRef<HTMLTextAreaElement | null>(null);
  const clickRef = useRef({ t: 0, x: 0, y: 0, n: 0 });
  const [scroller, setScroller] = useState<HTMLElement | null>(null);
  const focusSink = () => requestAnimationFrame(() => sinkRef.current?.focus({ preventScroll: true }));
  const pageEdit = useMemo(
    () => ({
      down: (page: number, x: number, y: number, event: React.PointerEvent) => {
        event.preventDefault();
        setScroller((event.currentTarget as HTMLElement).closest(".vtx-tex-pages") as HTMLElement | null);
        sinkRef.current?.focus({ preventScroll: true });
        (event.currentTarget as HTMLElement).setPointerCapture?.(event.pointerId);
        // (a pointer's click count: pointer events do not carry it)
        const last = clickRef.current;
        const now = performance.now();
        const clicks = now - last.t < 450 && Math.hypot(event.clientX - last.x, event.clientY - last.y) < 5 ? last.n + 1 : 1;
        clickRef.current = { t: now, x: event.clientX, y: event.clientY, n: clicks };
        void editor.pointerDown(page, x, y, event.shiftKey, { x: event.clientX, y: event.clientY }).then(() => {
          if (clicks >= 2 && !editor.formula) editor.selectWord(clicks);
          if (!editor.formula) focusSink();
        });
      },
      move: (page: number, x: number, y: number, event: React.PointerEvent) => {
        if (event.buttons & 1) editor.pointerMove(page, x, y);
      },
      up: () => editor.pointerUp(),
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [editor],
  );
  const decorate = useCallback(
    (page: number) => (
      <>
        <PageMarks editor={editor} page={page} zoom={zoomRef.current} />
        <FormulaOnPage editor={editor} page={page} zoom={zoomRef.current} onClosed={focusSink} />
        <LatexCommandMenu editor={editor} page={page} zoom={zoomRef.current} isEn={isEn} />
      </>
    ),
    // (drawn again with every change of the editor: the version is read by the components)
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [editor, editor.version, isEn, zoom],
  );
  const setViewRef = useRef<(next: "visual" | "source") => void>(() => undefined);
  const setView = (next: "visual" | "source") => {
    setViewState(next);
    storage.setItem(VIEW_KEY, next);
    if (next === "source") {
      // the source at the caret's place
      const c = editor.caret;
      const text = c ? buffers.current.get(c.path)?.text : undefined;
      if (c && text !== undefined) {
        const line = lineNumberAt(text, c.offset);
        const column = c.offset - (text.lastIndexOf("\n", c.offset - 1) + 1);
        void openTab(c.path).then(() => requestAnimationFrame(() => sourceRef.current?.caretToLine(line, column)));
      }
    } else focusSink();
  };
  setViewRef.current = setView;

  // ---- pages ↔ source ----
  const onPageClick = useCallback(
    async (page: number, x: number, y: number) => {
      await loadSync(buildRef.current);
      const s = sync.current;
      if (!s) return;
      const hit = s.tex.edit(page, x, y, (f) => !!s.files[f]);
      if (!hit) return;
      const path = s.files[hit.file];
      if (!path) return;
      await openTab(path, hit.line);
      sourceRef.current?.focus();
    },
    [loadSync, openTab],
  );

  const onVisible = useCallback(
    (visible: number[]) => {
      visibleRef.current = visible;
      void loadVisible();
    },
    [loadVisible],
  );
  useEffect(() => {
    void loadVisible();
  }, [build, loadVisible]);

  // ---- formulas ----
  type OpenFormula = { region: MathRegion; latex: string; back: (s: string) => string; left: number; top: number; path: string };
  const [formula, setFormulaState] = useState<OpenFormula | null>(null);
  const formulaRef = useRef<OpenFormula | null>(null);
  const setFormula = (f: OpenFormula | null) => {
    formulaRef.current = f;
    setFormulaState(f);
  };
  const formulaLatex = useRef("");
  const formulaApi = useRef<DocumentFormulaFieldApi>(null);
  const openFormula = useCallback(() => {
    const caret = sourceRef.current?.caret();
    if (!caret) return;
    const text = sourceRef.current?.text(caret.path) ?? "";
    const region = mathRegionAt(text, caret.pos);
    if (!region) {
      flash(t("No formula at the caret", "光标处没有公式"));
      return;
    }
    const { latex, back } = formulaForEditor(text, region);
    const at = sourceRef.current?.coords(region.from);
    const host = document.querySelector(".vtx-tex-source")?.getBoundingClientRect();
    formulaLatex.current = latex;
    setFormula({ region, latex, back, path: caret.path, left: at && host ? Math.max(8, at.left - host.left) : 24, top: at && host ? at.bottom - host.top + 6 : 60 });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  const closeFormula = useCallback((apply: boolean) => {
    const f = formulaRef.current;
    if (!f) return;
    formulaApi.current?.flush();
    setFormula(null);
    if (apply && sourceRef.current?.active() === f.path) sourceRef.current.replace(f.region.innerFrom, f.region.innerTo, f.back(formulaLatex.current));
    requestAnimationFrame(() => sourceRef.current?.focus());
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  const formulaCallbacks = useMemo<FormulaSessionCallbacks>(
    () => ({
      edit: (latex) => {
        formulaLatex.current = latex;
      },
      exit: () => closeFormula(true),
      remove: () => {
        formulaLatex.current = "";
      },
      dollar: () => closeFormula(true),
      undo: () => false,
      redo: () => false,
      handBack: () => undefined,
    }),
    [closeFormula],
  );

  // ---- export ----
  const fileStem = () => (pointer.main.split("/").pop() ?? "document").replace(/\.tex$/i, "");
  const exportPdf = useCallback(async () => {
    if (!service) return;
    flash(t("Exporting PDF…", "正在导出 PDF…"));
    const r = await compile(true);
    if (!r || r.status !== "ok") {
      flash(t("TeX could not make the PDF", "TeX 未能生成 PDF"));
      return;
    }
    try {
      const name = fileStem() + ".pdf";
      if (isTauri()) {
        const dest = await chooseExportPath(name, "pdf", isEn);
        if (!dest) return;
        await service.pdf(docId, r.build, { dest });
      } else {
        const out = await service.pdf(docId, r.build, { bytes: true });
        await saveBinaryFile(name, base64ToBytes(out.base64 ?? ""), "pdf", "application/pdf", isEn);
      }
      flash(t("PDF exported", "PDF 已导出"));
    } catch (error) {
      flash(errorText(error));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [compile, docId, service, isEn]);

  const exportProject = useCallback(async () => {
    if (!service) return;
    try {
      const info = await service.files(pointer.root);
      const entries: Record<string, Uint8Array> = {};
      for (const f of info.files) {
        const b = buffers.current.get(f.path);
        entries[f.path] = b ? strToU8(b.text) : base64ToBytes(await service.read(pointer.root, f.path));
      }
      if (await saveBinaryFile(fileStem() + ".zip", zipSync(entries), "zip", "application/zip", isEn)) flash(t("Project exported", "项目已导出"));
    } catch (error) {
      flash(errorText(error));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [service, pointer.root, isEn]);

  const exportMain = useCallback(async () => {
    const text = await readFile(pointer.main);
    if (await saveTextFile(fileStem() + ".tex", text, "tex", "text/x-tex", isEn)) flash(t("LaTeX exported", "LaTeX 已导出"));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [readFile, pointer.main, isEn]);

  // the headless tests drive the source the way the keyboard does, by line
  if (import.meta.env.DEV) (window as unknown as { __latexSource?: unknown }).__latexSource = sourceRef;
  actionsRef.current = { save: saveAll, exportPdf: () => void exportPdf(), exportProject: () => void exportProject(), exportMain: () => void exportMain() };

  // Mod+Enter: typeset now
  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if ((event.metaKey || event.ctrlKey) && event.key === "Enter") {
        event.preventDefault();
        void compile();
      }
      if ((event.metaKey || event.ctrlKey) && event.key === "/") {
        event.preventDefault();
        setViewRef.current(viewRef.current === "visual" ? "source" : "visual");
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [compile]);

  const bodyRef = useRef<HTMLDivElement>(null);

  const errors = messages.filter((m) => m.kind === "error");
  const warnings = messages.filter((m) => m.kind !== "error");
  const sourceFiles = files.filter((f) => f.kind === "tex" || f.kind === "bib");
  const name = (path: string) => path.split("/").pop() ?? path;
  const engineName = result?.engine ? ENGINE_LABEL[result.engine] ?? result.engine : pointer.engine ? ENGINE_LABEL[pointer.engine] : "TeX";
  const timing = result ? `${(result.timings.total / 1000).toFixed(2)} s` : "";

  return (
    <div ref={rootRef} className="vtx-tex-workspace" data-doc={docId} data-build={build} data-busy={busy ? "true" : undefined}>
      <div className="document-controls document-top-toolbar vtx-tex-toolbar" role="toolbar" aria-label="LaTeX">
        {leading}
        <button
          type="button"
          className="icon-button compact formula-formatting-button vtx-tex-compile"
          title={t("Typeset (⌘↵)", "排版 (⌘↵)")}
          aria-label={t("Typeset", "排版")}
          onClick={() => void compile()}
        >
          {busy ? <Loader2 size={16} className="vtx-tex-spin" /> : <Play size={16} />}
        </button>
        <span className="vtx-tex-menu-anchor">
          <button type="button" className="vtx-tex-chip" title={t("Engine", "引擎")} onClick={() => setEngineMenu((v) => !v)}>
            {engineName}
            <ChevronDown size={13} />
          </button>
          {engineMenu && (
            <div className="vtx-tex-menu" role="menu" onMouseLeave={() => setEngineMenu(false)}>
              {[null, ...ENGINES].map((e) => (
                <button
                  key={e ?? "auto"}
                  type="button"
                  role="menuitemradio"
                  aria-checked={pointer.engine === e}
                  className={pointer.engine === e ? "is-checked" : ""}
                  onClick={() => {
                    setEngineMenu(false);
                    onPointer({ ...pointer, engine: e });
                    pointerRef.current = { ...pointer, engine: e };
                    void compile();
                  }}
                >
                  {e ? ENGINE_LABEL[e] : t("Automatic", "自动")}
                </button>
              ))}
              <div className="vtx-tex-menu-rule" />
              <button
                type="button"
                role="menuitemcheckbox"
                aria-checked={pointer.shellEscape}
                className={pointer.shellEscape ? "is-checked" : ""}
                onClick={() => {
                  setEngineMenu(false);
                  const next = { ...pointer, shellEscape: !pointer.shellEscape };
                  onPointer(next);
                  pointerRef.current = next;
                  void compile();
                }}
              >
                Shell escape
              </button>
            </div>
          )}
        </span>
        <button
          type="button"
          className={"vtx-tex-chip vtx-tex-messages-button" + (errors.length ? " has-errors" : "")}
          data-errors={errors.length}
          data-warnings={warnings.length}
          title={t("Messages", "消息")}
          aria-pressed={showMessages}
          onClick={() => setShowMessages((v) => !v)}
        >
          <CircleX size={14} />
          {errors.length}
          <AlertTriangle size={14} />
          {warnings.length}
        </button>
        <span className="vtx-tex-time" title={result ? `${t("Last typesetting", "上次排版")}: TeX ${result.timings.tex} ms${result.warm ? "" : t(", preamble read", "，含导言")}` : t("Last typesetting", "上次排版")}>
          {timing}
        </span>
        <span className="formula-formatting-divider" aria-hidden="true" />
        <button
          type="button"
          className="icon-button compact formula-formatting-button"
          title={t("Formula (⌘E)", "公式 (⌘E)")}
          aria-label={t("Formula", "公式")}
          disabled={view === "source" ? !caretInMath : !editor.caret}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => (view === "source" ? openFormula() : editor.newFormula(false))}
        >
          <Sigma size={16} />
        </button>
        <button
          type="button"
          className={"icon-button compact formula-formatting-button vtx-tex-view-toggle" + (view === "source" ? " is-active" : "")}
          title={t("Source (⌘/)", "源码 (⌘/)")}
          aria-label={t("Source", "源码")}
          aria-pressed={view === "source"}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => setView(view === "visual" ? "source" : "visual")}
        >
          <Code2 size={16} />
        </button>
        <span className="formula-formatting-divider" aria-hidden="true" />
        <button type="button" className="icon-button compact formula-formatting-button" title={t("Zoom out", "缩小")} aria-label={t("Zoom out", "缩小")} onClick={() => setZoom(Math.max(0.3, zoom / 1.15))}>
          <Minus size={15} />
        </button>
        <button type="button" className="vtx-tex-chip vtx-tex-zoom" title={t("Fit width", "适合宽度")} onClick={() => setZoom("fit")}>
          {Math.round(zoom * 100)}%
        </button>
        <button type="button" className="icon-button compact formula-formatting-button" title={t("Zoom in", "放大")} aria-label={t("Zoom in", "放大")} onClick={() => setZoom(Math.min(4, zoom * 1.15))}>
          <Plus size={15} />
        </button>
        <span className="formula-formatting-divider" aria-hidden="true" />
        <span className="vtx-tex-menu-anchor">
          <button type="button" className="icon-button compact formula-formatting-button" title={t("Export", "导出")} aria-label={t("Export", "导出")} onClick={() => setExportMenu((v) => !v)}>
            <Download size={16} />
          </button>
          {exportMenu && (
            <div className="vtx-tex-menu" role="menu" onMouseLeave={() => setExportMenu(false)}>
              <button type="button" onClick={() => (setExportMenu(false), void exportPdf())}>
                PDF (.pdf)
              </button>
              <button type="button" onClick={() => (setExportMenu(false), void exportProject())}>
                {t("Project", "项目")} (.zip)
              </button>
              <button type="button" onClick={() => (setExportMenu(false), void exportMain())}>
                LaTeX (.tex)
              </button>
            </div>
          )}
        </span>
        <span className="vtx-tex-toolbar-fill" />
        {trailing}
      </div>
      {problem && <div className="vtx-tex-problem">{problem}</div>}
      <div className={"vtx-tex-body is-" + view} ref={bodyRef}>
        <div className="vtx-tex-source latex-source-editor document-source-view" style={view === "visual" ? { display: "none" } : undefined}>
          <div className="vtx-tex-tabs" role="tablist">
            {tabs.map((path) => (
              <span
                key={path}
                className={"vtx-tex-tab" + (path === active ? " is-active" : "") + (legacy.current.has(path) ? " is-read-only" : "")}
                role="tab"
                aria-selected={path === active}
                title={legacy.current.has(path) ? `${path} (GB18030)` : path}
              >
                <button type="button" className="vtx-tex-tab-name" onClick={() => void openTab(path)}>
                  {name(path)}
                </button>
                {path !== pointer.main && (
                  <button type="button" className="vtx-tex-tab-close" aria-label={t("Close", "关闭")} title={t("Close", "关闭")} onClick={() => closeTab(path)}>
                    <X size={12} />
                  </button>
                )}
              </span>
            ))}
            <span className="vtx-tex-menu-anchor">
              <button type="button" className="vtx-tex-tab-add" title={t("Files", "文件")} aria-label={t("Files", "文件")} onClick={() => setFileMenu((v) => !v)}>
                <FileText size={14} />
                <ChevronDown size={12} />
              </button>
              {fileMenu && (
                <div className="vtx-tex-menu vtx-tex-file-menu" role="menu" onMouseLeave={() => setFileMenu(false)}>
                  {sourceFiles.map((f) => (
                    <button key={f.path} type="button" className={tabs.includes(f.path) ? "is-checked" : ""} onClick={() => (setFileMenu(false), void openTab(f.path))}>
                      {f.path}
                    </button>
                  ))}
                </div>
              )}
            </span>
          </div>
          <LatexSource
            ref={sourceRef}
            onChange={onChange}
            onChanges={onChanges}
            onCaret={onCaret}
            onFormula={openFormula}
            onSave={() => void saveAll().then((ok) => ok && flash(t("Saved", "已保存")))}
          />
          {formula && (
            <>
              <div className="vtx-tex-formula-dismiss" onMouseDown={() => closeFormula(true)} />
              <div
                className="vtx-tex-formula"
                style={{ left: formula.left, top: formula.top }}
                onKeyDown={(e) => {
                  if (e.key === "Escape") closeFormula(false);
                  if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) closeFormula(true);
                }}
              >
                <DocumentFormulaField
                  id={`latex-formula-${formula.region.from}`}
                  latex={formula.latex}
                  display={formula.region.display}
                  entry={{ kind: "end" }}
                  fontSize={22}
                  callbacks={formulaCallbacks}
                  apiRef={formulaApi}
                />
              </div>
            </>
          )}
        </div>
        {view === "visual" && (
          <LatexPages
            pages={pages}
            contents={contents}
            zoom={zoom}
            marker={null}
            onPageClick={(p, x, y) => void onPageClick(p, x, y)}
            onVisible={onVisible}
            onZoom={(z) => setZoom(z)}
            reveal={reveal}
            onWidth={setPaneWidth}
            busy={busy}
            overlays={editor.overlays()}
            edit={pageEdit}
            decorate={decorate}
          >
            <KeyboardSink editor={editor} zoom={zoom} scroller={scroller} sinkRef={sinkRef} onCopy={() => undefined} />
          </LatexPages>
        )}
        {showMessages && (
          <div className="vtx-tex-messages" role="list">
            {messages.length === 0 && <div className="vtx-tex-message is-none">{t("No messages", "没有消息")}</div>}
            {messages.map((m, i) => (
              <button
                key={i}
                type="button"
                role="listitem"
                className={"vtx-tex-message is-" + m.kind}
                title={m.context ?? m.message}
                onClick={() => {
                  if (!m.file) return;
                  setView("source");
                  void openTab(m.file, m.line || undefined);
                }}
              >
                {m.kind === "error" ? <CircleX size={13} /> : <AlertTriangle size={13} />}
                <span className="vtx-tex-message-place">{m.file ? `${name(m.file)}${m.line ? `:${m.line}` : ""}` : ""}</span>
                <span className="vtx-tex-message-text">{m.message}</span>
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
