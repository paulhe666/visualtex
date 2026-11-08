// A LaTeX document (personal edition): the project's files are the document. Its source on the left (a tab a
// file), the pages TeX made of it on the right — SVG from the same DVI/XDV as the PDF, so the screen is the
// export. Typing saves the files a moment later and has TeX typeset again (the engine is already waiting with
// the preamble read); the pages change when the new ones are there. A click on a page goes to its source line;
// the caret's line is marked on the pages. TeX's errors are listed and marked on their lines; the last good
// pages stay. A formula at the caret opens in the formula editor (Mod+E) and goes back as LaTeX.
import { useCallback, useEffect, useMemo, useRef, useState, type MutableRefObject, type ReactNode } from "react";
import { gunzipSync, strToU8, zipSync } from "fflate";
import { isTauri } from "@tauri-apps/api/core";
import {
  AlertTriangle,
  ChevronDown,
  CircleX,
  Download,
  FileText,
  Loader2,
  Minus,
  Play,
  Plus,
  Sigma,
  X,
} from "lucide-react";
import { decodeText } from "../../../shared/visualtex-document/src/index.ts";
import type { FormulaSessionCallbacks } from "../../../shared/visualtex-document/src/index.ts";
import { chooseExportPath, saveBinaryFile, saveTextFile } from "../documentLibraryApi";
import { DocumentFormulaField, type DocumentFormulaFieldApi } from "../DocumentFormulaField";
import { createSafeStorage } from "../../runtime/safeStorage";
import { LatexPages, type PageContent, type PageMarker } from "./LatexPages";
import { LatexSource, type LatexSourceHandle, type LineMark } from "./LatexSource";
import { formulaForEditor, mathRegionAt, type MathRegion } from "./mathRegion";
import { projectPath, SyncTex } from "./synctex";
import { parseTexLog, type TexMessage } from "./texLog";
import { base64ToBytes, texService, type CompileResult, type PageMeta, type ProjectFile, type TexService } from "./texService";
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
const SPLIT_KEY = "visualtex.latex.split.v1";
const SAVE_DELAY_MS = 700;
const COMPILE_DELAY_MS = 220;
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
  const [marker, setMarker] = useState<PageMarker | null>(null);
  const [reveal, setReveal] = useState<{ page: number; y: number; nonce: number; ifHidden?: boolean } | null>(null);
  const [zoomSetting, setZoomSetting] = useState<number | "fit">(() => {
    const z = storage.getItem(ZOOM_KEY);
    return z && z !== "fit" ? Number(z) || "fit" : "fit";
  });
  const [paneWidth, setPaneWidth] = useState(800);
  const [split, setSplit] = useState(() => Number(storage.getItem(SPLIT_KEY)) || 0.44);
  const [caretInMath, setCaretInMath] = useState(false);

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

  const scheduleCompile = useCallback(() => {
    window.clearTimeout(run.current.timer);
    run.current.timer = window.setTimeout(() => void compile(), COMPILE_DELAY_MS);
  }, [compile]);

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

  // ---- editing ----
  const caretTimer = useRef(0);
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

  const onChange = useCallback(
    (path: string, text: string) => {
      const b = buffers.current.get(path);
      if (b) b.text = text;
      else buffers.current.set(path, { disk: "", text });
      window.clearTimeout(saveTimers.current.get(path));
      saveTimers.current.set(path, window.setTimeout(() => void saveFile(path), SAVE_DELAY_MS));
      scheduleCompile();
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
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [compile]);

  // ---- layout: the divider ----
  const bodyRef = useRef<HTMLDivElement>(null);
  const startDrag = (event: React.PointerEvent) => {
    const body = bodyRef.current;
    if (!body) return;
    event.preventDefault();
    const rect = body.getBoundingClientRect();
    const move = (e: PointerEvent) => setSplit(Math.min(0.75, Math.max(0.2, (e.clientX - rect.left) / rect.width)));
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      setSplit((s) => {
        storage.setItem(SPLIT_KEY, String(s));
        return s;
      });
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };

  const errors = messages.filter((m) => m.kind === "error");
  const warnings = messages.filter((m) => m.kind !== "error");
  const sourceFiles = files.filter((f) => f.kind === "tex" || f.kind === "bib");
  const name = (path: string) => path.split("/").pop() ?? path;
  const engineName = result?.engine ? ENGINE_LABEL[result.engine] ?? result.engine : pointer.engine ? ENGINE_LABEL[pointer.engine] : "TeX";
  const timing = result ? `${(result.timings.total / 1000).toFixed(2)} s` : "";

  return (
    <div className="vtx-tex-workspace" data-doc={docId} data-build={build} data-busy={busy ? "true" : undefined}>
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
          disabled={!caretInMath}
          onMouseDown={(e) => e.preventDefault()}
          onClick={openFormula}
        >
          <Sigma size={16} />
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
      <div className="vtx-tex-body" ref={bodyRef} style={{ gridTemplateColumns: `${split * 100}% 6px minmax(0, 1fr)` }}>
        <div className="vtx-tex-source latex-source-editor document-source-view">
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
          <LatexSource ref={sourceRef} onChange={onChange} onCaret={onCaret} onFormula={openFormula} onSave={() => void saveAll().then((ok) => ok && flash(t("Saved", "已保存")))} />
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
                  onClick={() => m.file && void openTab(m.file, m.line || undefined)}
                >
                  {m.kind === "error" ? <CircleX size={13} /> : <AlertTriangle size={13} />}
                  <span className="vtx-tex-message-place">{m.file ? `${name(m.file)}${m.line ? `:${m.line}` : ""}` : ""}</span>
                  <span className="vtx-tex-message-text">{m.message}</span>
                </button>
              ))}
            </div>
          )}
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
        <div className="vtx-tex-divider" onPointerDown={startDrag} role="separator" aria-orientation="vertical" />
        <LatexPages
          pages={pages}
          contents={contents}
          zoom={zoom}
          marker={marker}
          onPageClick={(p, x, y) => void onPageClick(p, x, y)}
          onVisible={onVisible}
          onZoom={(z) => setZoom(z)}
          reveal={reveal}
          onWidth={setPaneWidth}
          busy={busy}
        />
      </div>
    </div>
  );
}
