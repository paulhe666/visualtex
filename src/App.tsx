// Web shell for the macOS editor. Mirrors apps/macos/src/App.tsx, minus the
// desktop-only parts (Office, updates, keypad window, local/quick/silent OCR),
// plus the web-only parts (API OCR, landing showcase, default zoom).
import { ChangeEvent, useCallback, useEffect, useLayoutEffect, useRef, useState } from "react";
import {
  AlertCircle,
  BookOpenText,
  Check,
  Code2,
  FileDown,
  FilePlus2,
  FolderOpen,
  GraduationCap,
  History,
  Languages,
  LoaderCircle,
  Menu,
  PanelLeftClose,
  PanelLeftOpen,
  Redo2,
  Save,
  ScanLine,
  Settings2,
  Undo2,
  X,
} from "lucide-react";
import {
  type MathEditorHandle,
  type MathEditorInsertionTarget,
} from "./editor/MathEditor";
import { SettingsDialog } from "./components/SettingsDialog";
import { FormulaHotkeyManagerDialog } from "./components/FormulaHotkeyManagerDialog";
import { HistoryPanel } from "./components/HistoryPanel";
import { HelpManualDialog } from "./components/HelpManualDialog";
import { ExportDialog } from "./components/ExportDialog";
import { WebOcrDialog } from "./components/WebOcrDialog";
import { EditorWorkspace } from "./workspace/EditorWorkspace";
import {
  EDITOR_ZOOM_STEP,
  joinFormulaLines,
  useEditorStore,
} from "./stores/editorStore";
import { useHistoryManager, useHistorySnapshot, useDocumentSession } from "./history/EditorSession";
import {
  createBlankDocumentSnapshot,
  reconcileFormulaLines,
} from "./history/documentHistory";
import type {
  DocumentSnapshot,
  ReplaceDocumentEntry,
} from "./history/historyTypes";
import { copyFormulaLinesUniversal } from "./clipboard/LatexCopyService";
import { normalizeChineseLatex } from "./editor/normalizeChineseLatex";
import type {
  FormulaDocument,
  LatexFormatProfile,
} from "./types/formula";
import { applyDocumentTheme, publishSynchronizedTheme } from "./themeSync";
import { copyFormulaDocumentPngToClipboard } from "./export/pngClipboard";
import { readLocalStorage, writeLocalStorage } from "./runtime/safeStorage";
import { isLandingPreview, LANDING_PREVIEW_ZOOM, tutorialLanguage, tutorialLessonId } from "./runtime/landingPreview";
import { findTutorialLesson, type TutorialSnapshot } from "./tutorial/lessons";
import { useFormulaHotkeyStore } from "./stores/formulaHotkeyStore";
import {
  loadWebOcrConfiguration,
  recognizeFormulaWithWebApi,
} from "./ocr/webOcrService";

type InlineOcrStatus = "running" | "success" | "error";

interface InlineOcrState {
  status: InlineOcrStatus;
  message: string;
  seconds: number;
}

const WEB_DEFAULT_ZOOM = 0.45;
const TUTORIAL_ZOOM = 0.6;
const WEB_DEFAULT_ZOOM_MIGRATION_KEY = "visualtex.web.default-zoom.45.v1";
const LANDING_PREVIEW_LINES = [
  String.raw`J_\nu(x)=\sum_{k=0}^{\infty}\frac{(-1)^k}{k!\Gamma(k+\nu+1)}\left(\frac{x}{2}\right)^{2k+\nu}`,
  String.raw`R_{\mu\nu}-\frac{1}{2}Rg_{\mu\nu}+\Lambda g_{\mu\nu}=\frac{8\pi G}{c^4}T_{\mu\nu}`,
  String.raw`\mathrm{d}s^2=-\left(1-\frac{2GMr}{c^2\Sigma}\right)c^2\mathrm{d}t^2-\frac{4GMar\sin^2\theta}{c^2\Sigma}c\mathrm{d}t\mathrm{d}\phi+\frac{\Sigma}{\Delta}\mathrm{d}r^2+\Sigma\mathrm{d}\theta^2+\left(r^2+a^2+\frac{2GMa^2r\sin^2\theta}{c^2\Sigma}\right)\sin^2\theta\mathrm{d}\phi^2`,
  String.raw`\mathcal{L}_{\mathrm{SM}}=-\frac{1}{4}F^a_{\mu\nu}F^{a\mu\nu}+i\bar{\psi}\gamma^\mu D_\mu\psi+(D_\mu\Phi)^\dagger(D^\mu\Phi)-V(\Phi)-\left(y_{ij}\bar{\psi}_{Li}\Phi\psi_{Rj}+\mathrm{h.c.}\right)`,
] as const;

// Custom tiles live in localStorage (in-memory inside the tutorial sandbox).
function readTutorialCustomTiles() {
  try {
    const library = JSON.parse(localStorage.getItem("visualtex-custom-formula-tiles") ?? "{}");
    return {
      customTiles: Array.isArray(library?.tiles) ? library.tiles : [],
      customSectionCount: Array.isArray(library?.sections) ? library.sections.length : 0,
    };
  } catch {
    return { customTiles: [], customSectionCount: 0 };
  }
}

function App() {
  const editorRef = useRef<MathEditorHandle>(null);
  const historyManager = useHistoryManager();
  const { openDocumentWithHistory,
    replaceDocumentWithHistory: replaceDocumentTransaction } = useDocumentSession(editorRef);
  const ocrInsertionTargetRef = useRef<MathEditorInsertionTarget | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const menuButtonRef = useRef<HTMLButtonElement>(null);
  const appMenuRef = useRef<HTMLDivElement>(null);
  const copyMenuButtonRef = useRef<HTMLButtonElement>(null);
  const copyMenuRef = useRef<HTMLDivElement>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [formulaHotkeyManagerOpen, setFormulaHotkeyManagerOpen] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [helpOpen, setHelpOpen] = useState(false);
  const [ocrOpen, setOcrOpen] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(() => window.innerWidth >= 1040);
  const [copyMenuOpen, setCopyMenuOpen] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [toast, setToast] = useState("");
  const [savedPulse, setSavedPulse] = useState(false);
  const [desktopTopToolsMount, setDesktopTopToolsMount] =
    useState<HTMLDivElement | null>(null);
  const [sourceDocumentRevision, setSourceDocumentRevision] = useState(0);
  const [inlineOcr, setInlineOcr] = useState<InlineOcrState | null>(null);
  const inlineOcrBusyRef = useRef(false);
  const inlineOcrClearTimerRef = useRef<number | null>(null);
  const initialEditorFocusDoneRef = useRef(false);
  const pngClipboardBusyRef = useRef(false);

  const title = useEditorStore((state) => state.title);
  const setTitle = useEditorStore((state) => state.setTitle);
  const lines = useEditorStore((state) => state.lines);
  const activeLineId = useEditorStore((state) => state.activeLineId);
  const formulaAlignment = useEditorStore((state) => state.formulaAlignment);
  const theme = useEditorStore((state) => state.theme);
  const synchronizedThemeRef = useRef(theme);
  const language = useEditorStore((state) => state.language);
  const setLanguage = useEditorStore((state) => state.setLanguage);
  const zoom = useEditorStore((state) => state.zoom);
  const setZoom = useEditorStore((state) => state.setZoom);
  const setSourceOpen = useEditorStore((state) => state.setSourceOpen);
  const replaceDocumentState = useEditorStore((state) => state.replaceDocumentState);
  const editorLayout = useEditorStore((state) => state.editorLayout);
  const pngExportBackground = useEditorStore(
    (state) => state.pngExportBackground,
  );
  const formulaLetterFont = useEditorStore((state) => state.formulaLetterFont);
  const formulaChineseFont = useEditorStore((state) => state.formulaChineseFont);
  const latexFormatProfile = useEditorStore((state) => state.latexFormatProfile);
  const setLatexFormatProfile = useEditorStore(
    (state) => state.setLatexFormatProfile,
  );
  const addHistory = useEditorStore((state) => state.addHistory);
  const toDocument = useEditorStore((state) => state.toDocument);
  const historyState = useHistorySnapshot();
  const isEn = language === "en";
  const latex = joinFormulaLines(lines);
  const inlineProfileLabel =
    latexFormatProfile.inlineWrapper === "paren" ? "\\( \\)" : "$ $";
  const displayProfileLabel =
    latexFormatProfile.displayWrapper === "bracket"
      ? "\\[ \\]"
      : latexFormatProfile.displayWrapper === "equation"
        ? latexFormatProfile.numbered
          ? "equation"
          : "equation*"
        : "$$ $$";
  const multilineProfileLabel =
    latexFormatProfile.multilineEnvironment +
    (latexFormatProfile.numbered ? "" : "*");
  const latexProfileSummary =
    `${inlineProfileLabel} · ${displayProfileLabel} · ${multilineProfileLabel}`;
  const inlineOcrIsBusy = inlineOcr?.status === "running";

  const replaceDocumentWithHistory = (after: DocumentSnapshot, source: ReplaceDocumentEntry["source"]) => {
    if (source !== "source-apply") setSourceDocumentRevision(revision => revision + 1);
    return replaceDocumentTransaction(after, source);
  };

  // The landing page embeds /editor?landing-preview as a read-only showcase.
  // Its storage is in-memory (see runtime/landingPreview), so this never
  // touches the visitor's own document.
  useLayoutEffect(() => {
    if (!isLandingPreview) return;
    replaceDocumentState({
      title: "示例公式",
      lines: LANDING_PREVIEW_LINES.map((latex, index) => ({
        id: `landing-preview-${index + 1}`,
        latex,
      })),
      activeLineId: "landing-preview-1",
      formulaAlignment,
      selectionByLineId: {},
    });
    setSourceOpen(false);
    setZoom(LANDING_PREVIEW_ZOOM);
    setSourceDocumentRevision((revision) => revision + 1);
    // The showcase is always drawn at the same scale.
    return useEditorStore.subscribe((state) => {
      if (state.zoom !== LANDING_PREVIEW_ZOOM) setZoom(LANDING_PREVIEW_ZOOM);
    });
  }, []);

  // The tutorial page embeds /editor?tutorial=<lesson> as a practice editor
  // (in-memory storage, see runtime/landingPreview) and polls its state.
  useLayoutEffect(() => {
    const lesson = findTutorialLesson(tutorialLessonId);
    if (!lesson) return;
    setLanguage(tutorialLanguage);
    replaceDocumentState({
      title: lesson.title[tutorialLanguage === "en" ? "en" : "zh"],
      lines: lesson.start.map((line, index) => ({
        id: `tutorial-${lesson.id}-${index + 1}`,
        latex: line.latex,
        mode: line.mode,
      })),
      activeLineId: `tutorial-${lesson.id}-${lesson.start.length}`,
      formulaAlignment,
      selectionByLineId: {},
    });
    setSourceDocumentRevision((revision) => revision + 1);
    setZoom(TUTORIAL_ZOOM);
    const host = window as Window & { visualtexTutorial?: { snapshot: () => TutorialSnapshot } };
    let copiedText = "";
    const clipboard = navigator.clipboard;
    if (clipboard) {
      const writeText = clipboard.writeText.bind(clipboard);
      clipboard.writeText = (text: string) => {
        copiedText = text;
        return writeText(text);
      };
    }
    host.visualtexTutorial = {
      snapshot: () => {
        const state = useEditorStore.getState();
        return {
          lines: state.lines.map((line) => ({ latex: line.latex, mode: line.mode === "inline" ? "inline" : "display" })),
          profile: state.latexFormatProfile,
          copiedText,
          hotkeyIds: useFormulaHotkeyStore.getState().bindings.map((binding) => binding.id),
          ...readTutorialCustomTiles(),
        };
      },
    };
  }, []);

  // The web editor defaults to a smaller zoom than the desktop window.
  useLayoutEffect(() => {
    if (isLandingPreview || tutorialLessonId) return;
    if (readLocalStorage(WEB_DEFAULT_ZOOM_MIGRATION_KEY) === "true") return;
    if (useEditorStore.getState().zoom === 0.6) setZoom(WEB_DEFAULT_ZOOM);
    writeLocalStorage(WEB_DEFAULT_ZOOM_MIGRATION_KEY, "true");
  }, [setZoom]);

  useEffect(() => {
    if (isLandingPreview || initialEditorFocusDoneRef.current) return;
    initialEditorFocusDoneRef.current = true;
    const frame = window.requestAnimationFrame(() => {
      const active = document.activeElement;
      const userOwnsFocus =
        active instanceof HTMLElement &&
        active !== document.body &&
        active !== document.documentElement;
      if (userOwnsFocus) return;
      editorRef.current?.focus({ target: "last", moveToEnd: true });
    });
    return () => window.cancelAnimationFrame(frame);
  }, []);

  useEffect(() => {
    const checkpointTimer = window.setInterval(() => {
      historyManager.commitPendingTransaction();
      void historyManager.createCheckpoint("autosave").catch(() => undefined);
    }, 30_000);
    const handleBeforeUnload = () => {
      historyManager.commitPendingTransaction();
      void historyManager.createCheckpoint("before-unload").catch(() => undefined);
    };
    window.addEventListener("beforeunload", handleBeforeUnload);
    return () => {
      window.clearInterval(checkpointTimer);
      window.removeEventListener("beforeunload", handleBeforeUnload);
    };
  }, []);

  useEffect(() => {
    if (synchronizedThemeRef.current === theme) {
      applyDocumentTheme(theme);
    } else {
      synchronizedThemeRef.current = theme;
      publishSynchronizedTheme(theme);
    }
  }, [theme]);

  useEffect(() => {
    document.documentElement.lang = isEn ? "en" : "zh-CN";
  }, [isEn]);

  useEffect(() => {
    const compactWindow = window.matchMedia("(max-width: 1040px)");
    const handleCompactWindow = (event: MediaQueryListEvent) => {
      if (event.matches) setSidebarOpen(false);
    };
    compactWindow.addEventListener("change", handleCompactWindow);
    return () => compactWindow.removeEventListener("change", handleCompactWindow);
  }, []);

  useEffect(() => {
    if (!latex.trim()) return;
    const timeout = window.setTimeout(() => addHistory(latex), 1800);
    return () => window.clearTimeout(timeout);
  }, [latex, addHistory]);

  useEffect(() => {
    if (!toast) return;
    const timeout = window.setTimeout(() => setToast(""), 1800);
    return () => window.clearTimeout(timeout);
  }, [toast]);

  useEffect(() => {
    const menu = menuOpen ? appMenuRef.current : copyMenuOpen ? copyMenuRef.current : null;
    const trigger = menuOpen ? menuButtonRef.current : copyMenuButtonRef.current;
    if (!menu || !trigger) return;

    const items = Array.from(
      menu.querySelectorAll<HTMLButtonElement>('button:not(:disabled)'),
    );
    const frame = window.requestAnimationFrame(() => items[0]?.focus());

    const handleMenuKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        setMenuOpen(false);
        setCopyMenuOpen(false);
        trigger.focus({ preventScroll: true });
        return;
      }

      if (event.key !== "ArrowDown" && event.key !== "ArrowUp") return;
      event.preventDefault();
      const currentIndex = items.indexOf(document.activeElement as HTMLButtonElement);
      const direction = event.key === "ArrowDown" ? 1 : -1;
      const nextIndex = currentIndex < 0
        ? 0
        : (currentIndex + direction + items.length) % items.length;
      items[nextIndex]?.focus();
    };

    menu.addEventListener("keydown", handleMenuKeyDown);
    return () => {
      window.cancelAnimationFrame(frame);
      menu.removeEventListener("keydown", handleMenuKeyDown);
    };
  }, [menuOpen, copyMenuOpen]);

  const inlineOcrStatus = inlineOcr?.status;
  useEffect(() => {
    if (inlineOcrStatus !== "running") return;
    const timer = window.setInterval(() => {
      setInlineOcr((current) =>
        current
          ? {
              ...current,
              seconds: current.seconds + 1,
            }
          : current,
      );
    }, 1000);
    return () => window.clearInterval(timer);
  }, [inlineOcrStatus]);

  useEffect(
    () => () => {
      if (inlineOcrClearTimerRef.current !== null) {
        window.clearTimeout(inlineOcrClearTimerRef.current);
      }
    },
    [],
  );

  const scheduleInlineOcrClear = (delay: number) => {
    if (inlineOcrClearTimerRef.current !== null) {
      window.clearTimeout(inlineOcrClearTimerRef.current);
    }
    inlineOcrClearTimerRef.current = window.setTimeout(() => {
      setInlineOcr(null);
      inlineOcrClearTimerRef.current = null;
    }, delay);
  };

  // Images pasted into a formula field are recognized through the configured
  // web OCR API and inserted at the caret captured when the paste happened.
  const handleEditorImagePaste = useCallback(async (
    file: File,
    target: MathEditorInsertionTarget,
  ) => {
    if (inlineOcrBusyRef.current) {
      setToast(isEn ? "Another pasted image is being recognized" : "已有一张粘贴图片正在识别");
      return;
    }

    if (inlineOcrClearTimerRef.current !== null) {
      window.clearTimeout(inlineOcrClearTimerRef.current);
      inlineOcrClearTimerRef.current = null;
    }

    inlineOcrBusyRef.current = true;
    setInlineOcr({
      status: "running",
      message: isEn ? "Recognizing the pasted image…" : "正在识别粘贴的图片…",
      seconds: 0,
    });
    try {
      const result = await recognizeFormulaWithWebApi(
        file,
        loadWebOcrConfiguration(),
        (progress) =>
          setInlineOcr((current) =>
            current
              ? { ...current, message: isEn ? progress.messageEn : progress.messageZh }
              : current,
          ),
      );
      const recognizedLatex = result.formulas
        .map((formula) => formula.trim())
        .filter(Boolean)
        .join("\n");
      if (!recognizedLatex) {
        throw new Error(isEn ? "OCR returned an empty formula" : "OCR 没有返回可用公式");
      }

      const inserted =
        editorRef.current?.insertLatexAt(target, recognizedLatex, "ocr") ?? false;
      if (!inserted) {
        throw new Error(
          isEn
            ? "The formula row was deleted, so the result was not inserted"
            : "原公式行已删除，识别结果未插入",
        );
      }

      setInlineOcr((current) => ({
        status: "success",
        message: isEn ? "Recognized and inserted" : "已识别并插入",
        seconds: current?.seconds ?? 0,
      }));
      setToast(isEn ? "Image recognized" : "图片已识别");
      scheduleInlineOcrClear(1800);
    } catch (error) {
      const message =
        (error instanceof Error ? error.message : typeof error === "string" ? error : "") ||
        (isEn ? "Image OCR failed" : "图片 OCR 失败");
      setInlineOcr((current) => ({
        status: "error",
        message,
        seconds: current?.seconds ?? 0,
      }));
      setToast(message);
      scheduleInlineOcrClear(4500);
    } finally {
      inlineOcrBusyRef.current = false;
    }
  }, [isEn]);

  // Browsers deliver clipboard images to the document rather than reliably to
  // the MathLive keyboard sink, so catch them here before MathLive reads text.
  useEffect(() => {
    const handleFormulaImagePaste = (event: ClipboardEvent) => {
      const formulaField = event.composedPath().find(
        (target): target is HTMLElement =>
          target instanceof HTMLElement && target.tagName === "MATH-FIELD",
      );
      if (!formulaField && document.activeElement?.tagName !== "MATH-FIELD") {
        return;
      }

      const clipboard = event.clipboardData;
      const item = Array.from(clipboard?.items ?? []).find(
        (candidate) =>
          candidate.kind === "file" && candidate.type.startsWith("image/"),
      );
      const image =
        item?.getAsFile() ??
        Array.from(clipboard?.files ?? []).find((file) =>
          file.type.startsWith("image/"),
        );
      if (!image) return;

      const target = editorRef.current?.captureInsertionTarget();
      if (!target) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      void handleEditorImagePaste(image, target);
    };

    document.addEventListener("paste", handleFormulaImagePaste, true);
    return () =>
      document.removeEventListener("paste", handleFormulaImagePaste, true);
  }, [handleEditorImagePaste]);

  const captureOcrInsertionTarget = () => {
    const target = editorRef.current?.captureInsertionTarget() ?? null;
    if (target) ocrInsertionTargetRef.current = target;
  };

  const openOcrDialog = () => {
    captureOcrInsertionTarget();
    setOcrOpen(true);
  };

  const updateLatexFormatProfile = (
    patch: Partial<LatexFormatProfile>,
  ) => {
    setLatexFormatProfile(patch);
  };

  const handleCopy = async () => {
    try {
      await copyFormulaLinesUniversal(lines, latexFormatProfile);
      addHistory(latex);
      setToast(isEn ? "LaTeX copied" : "已复制 LaTeX");
      return true;
    } catch {
      setToast(
        isEn
          ? "Copy failed: the browser blocked the clipboard."
          : "复制失败：浏览器不允许写入剪贴板。",
      );
      return false;
    }
  };

  const handleCopyPng = async () => {
    if (pngClipboardBusyRef.current) return;
    pngClipboardBusyRef.current = true;
    try {
      await copyFormulaDocumentPngToClipboard(
        lines.map((line) => line.latex),
        {
          background: pngExportBackground,
          formulaLetterFont,
          formulaChineseFont,
        },
      );
      setToast(isEn ? "PNG copied to Clipboard" : "PNG 已复制到剪贴板");
    } catch (cause) {
      const message = cause instanceof Error ? cause.message : String(cause);
      setToast(
        isEn
          ? `Unable to copy PNG: ${message}`
          : `复制 PNG 失败：${message}`,
      );
    } finally {
      pngClipboardBusyRef.current = false;
    }
  };

  const saveDocument = () => {
    historyManager.commitPendingTransaction();
    void historyManager.createCheckpoint("save-document").catch(() => undefined);
    const document = toDocument();
    const blob = new Blob([JSON.stringify(document, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const link = window.document.createElement("a");
    const safeTitle =
      title.trim().replace(/[\\/:*?"<>|]/g, "-") ||
      (isEn ? "Untitled Formula" : "未命名公式");
    link.href = url;
    link.download = safeTitle + ".visualtex.json";
    link.click();
    URL.revokeObjectURL(url);
    setSavedPulse(true);
    setToast(isEn ? "Formula document saved" : "公式文档已保存");
    window.setTimeout(() => setSavedPulse(false), 900);
  };

  const openDocument = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const parsed = JSON.parse(await file.text()) as FormulaDocument;
      if (!parsed.formulas || !Array.isArray(parsed.formulas)) {
        throw new Error("invalid");
      }
      openDocumentWithHistory(parsed);
      setSourceDocumentRevision((revision) => revision + 1);
      setToast(isEn ? "Formula document opened" : "公式文档已打开");
    } catch {
      setToast(
        isEn
          ? "Unable to open: invalid file format"
          : "无法打开：文件格式不正确",
      );
    } finally {
      event.target.value = "";
    }
  };

  const newFormula = () => {
    addHistory(latex);
    const after = createBlankDocumentSnapshot(
      isEn ? "Untitled Formula" : "未命名公式",
    );
    replaceDocumentWithHistory(after, "new-document");
    setToast(isEn ? "Created a blank formula" : "已新建空白公式");
  };

  const handleTitleChange = (nextTitle: string) => {
    const beforeTitle = useEditorStore.getState().title;
    setTitle(nextTitle);
    historyManager.recordTitleEdit({
      beforeTitle,
      afterTitle: nextTitle,
    });
  };

  const runMenuAction = (action: () => void) => {
    setMenuOpen(false);
    action();
  };

  useEffect(() => {
    const handleWindowKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setMenuOpen(false);
        setCopyMenuOpen(false);
        return;
      }

      if (
        settingsOpen ||
        formulaHotkeyManagerOpen ||
        ocrOpen ||
        historyOpen ||
        helpOpen ||
        exportOpen
      ) {
        return;
      }

      const target = event.target instanceof Element ? event.target : null;
      const focusedElement =
        document.activeElement instanceof Element ? document.activeElement : null;
      const inCodeMirror = Boolean(
        target?.closest(".cm-editor") || focusedElement?.closest(".cm-editor"),
      );
      const primaryModifier = (event.metaKey || event.ctrlKey) && !event.altKey;
      const key = event.key.toLowerCase();
      const requestsUndo = primaryModifier && key === "z" && !event.shiftKey;
      const requestsRedo =
        primaryModifier &&
        ((key === "z" && event.shiftKey) ||
          (key === "y" && !event.shiftKey));

      if (requestsUndo || requestsRedo) {
        if (inCodeMirror) return;
        event.preventDefault();
        if (requestsRedo) historyManager.requestRedo();
        else historyManager.requestUndo();
        return;
      }

      if (!primaryModifier) return;
      if (key === "n") {
        event.preventDefault();
        newFormula();
      } else if (key === "o") {
        event.preventDefault();
        fileInputRef.current?.click();
      } else if (key === "s") {
        event.preventDefault();
        saveDocument();
      } else if (key === ",") {
        event.preventDefault();
        setSettingsOpen(true);
      } else if (key === "0") {
        event.preventDefault();
        setZoom(1);
      } else if (key === "=" || key === "+") {
        event.preventDefault();
        setZoom(zoom + EDITOR_ZOOM_STEP);
      } else if (key === "-") {
        event.preventDefault();
        setZoom(zoom - EDITOR_ZOOM_STEP);
      }
    };

    window.addEventListener("keydown", handleWindowKeyDown);
    return () => window.removeEventListener("keydown", handleWindowKeyDown);
  }, [
    latex,
    title,
    isEn,
    zoom,
    latexFormatProfile,
    settingsOpen,
    formulaHotkeyManagerOpen,
    ocrOpen,
    historyOpen,
    helpOpen,
    exportOpen,
  ]);

  const renderLatexProfileMenu = () =>
    copyMenuOpen ? (
      <div
        ref={copyMenuRef}
        id="copy-format-menu"
        className="copy-menu code-format-menu latex-profile-menu"
        role="dialog"
        aria-label={isEn ? "LaTeX format profile" : "LaTeX 格式配置"}
      >
        <div className="code-format-menu-header">
          <span className="copy-menu-label">
            {isEn ? "LaTeX format" : "LaTeX 格式"}
          </span>
        </div>

        <div className="latex-profile-options">
          <div className="latex-profile-row">
            <span>{isEn ? "Text" : "文字"}</span>
            <div className="latex-profile-segments">
              <button
                type="button"
                className={
                  latexFormatProfile.inlineTextPolicy === "text-command"
                    ? "is-selected"
                    : ""
                }
                data-latex-text-policy="text-command"
                onClick={() =>
                  updateLatexFormatProfile({
                    inlineTextPolicy: "text-command",
                  })
                }
              >
                {"\\text{}"}
              </button>
              <button
                type="button"
                className={
                  latexFormatProfile.inlineTextPolicy === "outside-math"
                    ? "is-selected"
                    : ""
                }
                data-latex-text-policy="outside-math"
                onClick={() =>
                  updateLatexFormatProfile({
                    inlineTextPolicy: "outside-math",
                  })
                }
              >
                {isEn ? "outside" : "公式外"}
              </button>
            </div>
          </div>

          <div className="latex-profile-row">
            <span>{isEn ? "Inline" : "行内"}</span>
            <div className="latex-profile-segments">
              <button
                type="button"
                className={
                  latexFormatProfile.inlineWrapper === "dollar"
                    ? "is-selected"
                    : ""
                }
                data-latex-inline-wrapper="dollar"
                onClick={() =>
                  updateLatexFormatProfile({ inlineWrapper: "dollar" })
                }
              >
                $...$
              </button>
              <button
                type="button"
                className={
                  latexFormatProfile.inlineWrapper === "paren"
                    ? "is-selected"
                    : ""
                }
                data-latex-inline-wrapper="paren"
                onClick={() =>
                  updateLatexFormatProfile({ inlineWrapper: "paren" })
                }
              >
                {"\\(...\\)"}
              </button>
            </div>
          </div>

          <div className="latex-profile-row">
            <span>{isEn ? "Display" : "行间"}</span>
            <div className="latex-profile-segments">
              <button
                type="button"
                className={
                  latexFormatProfile.displayWrapper === "double-dollar"
                    ? "is-selected"
                    : ""
                }
                data-latex-display-wrapper="double-dollar"
                onClick={() =>
                  updateLatexFormatProfile({
                    displayWrapper: "double-dollar",
                  })
                }
              >
                $$...$$
              </button>
              <button
                type="button"
                className={
                  latexFormatProfile.displayWrapper === "bracket"
                    ? "is-selected"
                    : ""
                }
                data-latex-display-wrapper="bracket"
                onClick={() =>
                  updateLatexFormatProfile({ displayWrapper: "bracket" })
                }
              >
                {"\\[...\\]"}
              </button>
              <button
                type="button"
                className={
                  latexFormatProfile.displayWrapper === "equation"
                    ? "is-selected"
                    : ""
                }
                data-latex-display-wrapper="equation"
                onClick={() =>
                  updateLatexFormatProfile({ displayWrapper: "equation" })
                }
              >
                equation
              </button>
            </div>
          </div>

          <div className="latex-profile-row">
            <span>{isEn ? "Number" : "编号"}</span>
            <button
              type="button"
              className={
                "latex-profile-toggle" +
                (latexFormatProfile.numbered ? " is-selected" : "")
              }
              aria-pressed={latexFormatProfile.numbered}
              data-latex-numbered
              onClick={() =>
                updateLatexFormatProfile({
                  numbered: !latexFormatProfile.numbered,
                })
              }
            >
              <span aria-hidden="true">#</span>
              {latexFormatProfile.numbered
                ? isEn
                  ? "on"
                  : "开启"
                : isEn
                  ? "off"
                  : "关闭"}
            </button>
          </div>

          <div className="latex-profile-row">
            <span>{isEn ? "Multi-line" : "多行"}</span>
            <div className="latex-profile-segments">
              <button
                type="button"
                className={
                  latexFormatProfile.multilineEnvironment === "gather"
                    ? "is-selected"
                    : ""
                }
                data-latex-multiline="gather"
                onClick={() =>
                  updateLatexFormatProfile({
                    multilineEnvironment: "gather",
                  })
                }
              >
                gather
              </button>
              <button
                type="button"
                className={
                  latexFormatProfile.multilineEnvironment === "align"
                    ? "is-selected"
                    : ""
                }
                data-latex-multiline="align"
                onClick={() =>
                  updateLatexFormatProfile({
                    multilineEnvironment: "align",
                  })
                }
              >
                align
              </button>
            </div>
          </div>
        </div>
      </div>
    ) : null;

  const codeFormatControl = (
    <div className="copy-control code-format-control">
      <button
        ref={copyMenuButtonRef}
        type="button"
        className="copy-primary code-format-primary icon-only-toolbar-button"
        aria-expanded={copyMenuOpen}
        aria-haspopup="menu"
        aria-controls="copy-format-menu"
        title={isEn ? `LaTeX format: ${latexProfileSummary}` : `LaTeX 格式：${latexProfileSummary}`}
        onClick={() => {
          setMenuOpen(false);
          setCopyMenuOpen((open) => !open);
        }}
      >
        <Code2 size={17} />
      </button>
      {renderLatexProfileMenu()}
    </div>
  );

  return (
    <div
      className={
        `app-shell ${editorLayout === "classic" ? "is-classic-app-layout" : "is-standard-app-layout"}`
      }
    >
      <input
        ref={fileInputRef}
        type="file"
        accept=".json,.visualtex"
        className="visually-hidden"
        onChange={openDocument}
      />

      <header
        className={
          "app-header" + (menuOpen || copyMenuOpen ? " has-open-menu" : "")
        }
      >
        <div className="brand-area">
          <button
            ref={menuButtonRef}
            type="button"
            className={"menu-button " + (menuOpen ? "is-active" : "")}
            aria-label={isEn ? "Main menu" : "主菜单"}
            aria-expanded={menuOpen}
            aria-haspopup="menu"
            aria-controls="app-main-menu"
            onClick={() => {
              setCopyMenuOpen(false);
              setMenuOpen((open) => !open);
            }}
          >
            <Menu size={18} />
          </button>
          {editorLayout !== "classic" ? (
            <button
              type="button"
              className={"icon-button sidebar-toggle " + (sidebarOpen ? "is-active" : "")}
              aria-label={
                sidebarOpen
                  ? isEn
                    ? "Hide formula tools"
                    : "隐藏公式工具"
                  : isEn
                    ? "Show formula tools"
                    : "显示公式工具"
              }
              aria-pressed={sidebarOpen}
              onClick={() => setSidebarOpen((open) => !open)}
            >
              {sidebarOpen ? (
                <PanelLeftClose size={17} />
              ) : (
                <PanelLeftOpen size={17} />
              )}
            </button>
          ) : null}

          {menuOpen && (
            <div
              ref={appMenuRef}
              id="app-main-menu"
              className="app-menu-popover"
              role="menu"
              aria-label={isEn ? "VisualTeX menu" : "VisualTeX 菜单"}
            >
              <div className="app-menu-heading">
                <strong>VisualTeX</strong>
              </div>
              <button type="button" role="menuitem" onClick={() => runMenuAction(newFormula)}>
                <FilePlus2 size={16} />
                <span>{isEn ? "New formula" : "新建公式"}</span>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() =>
                  runMenuAction(() => fileInputRef.current?.click())
                }
              >
                <FolderOpen size={16} />
                <span>{isEn ? "Open document" : "打开文档"}</span>
                <kbd>Ctrl+O</kbd>
              </button>
              <button type="button" role="menuitem" onClick={() => runMenuAction(saveDocument)}>
                <Save size={16} />
                <span>{isEn ? "Save document" : "保存文档"}</span>
                <kbd>Ctrl+S</kbd>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setExportOpen(true))}
              >
                <FileDown size={16} />
                <span>{isEn ? "Export…" : "导出…"}</span>
              </button>
              <div className="app-menu-divider" />
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setHistoryOpen(true))}
              >
                <History size={16} />
                <span>{isEn ? "Formula history" : "公式历史"}</span>
              </button>
              <button
                type="button"
                role="menuitem"
                onPointerDown={captureOcrInsertionTarget}
                onClick={() => runMenuAction(openOcrDialog)}
              >
                <ScanLine size={16} />
                <span>{isEn ? "Formula image OCR" : "图片公式识别"}</span>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setSettingsOpen(true))}
              >
                <Settings2 size={16} />
                <span>{isEn ? "Settings" : "设置"}</span>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setHelpOpen(true))}
              >
                <BookOpenText size={16} />
                <span>{isEn ? "Help Manual" : "帮助手册"}</span>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => (window.top ?? window).location.assign("/tutorial"))}
              >
                <GraduationCap size={16} />
                <span>{isEn ? "Tutorial" : "新手教程"}</span>
              </button>
              <div className="app-menu-divider" />
              <div className="app-menu-language">
                <span>
                  <Languages size={15} />
                  {isEn ? "Language" : "语言"}
                </span>
                <div>
                  <button
                    type="button"
                    role="menuitemradio"
                    aria-checked={language === "cn"}
                    className={language === "cn" ? "is-active" : ""}
                    onClick={() => setLanguage("cn")}
                  >
                    CN
                  </button>
                  <button
                    type="button"
                    role="menuitemradio"
                    aria-checked={language === "en"}
                    className={language === "en" ? "is-active" : ""}
                    onClick={() => setLanguage("en")}
                  >
                    ENG
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>

        <div className="document-title-area">
          <input
            value={title}
            onChange={(event) => handleTitleChange(event.target.value)}
            onBlur={() => historyManager.commitPendingTransaction()}
            aria-label={isEn ? "Formula document title" : "公式文档标题"}
          />
          <span
            className={"save-state " + (savedPulse ? "is-saved" : "")}
            aria-label={isEn ? "Saved" : "已保存"}
            title={isEn ? "Saved" : "已保存"}
          >
            <Check size={13} />
          </span>
        </div>

        <div className="header-actions">
          <div className="action-group file-actions">
            <button type="button" className="icon-button" onClick={newFormula} aria-label={isEn ? "New" : "新建"} title={isEn ? "New" : "新建"}>
              <FilePlus2 size={17} />
            </button>
            <button type="button" className="icon-button" onClick={() => fileInputRef.current?.click()} aria-label={isEn ? "Open" : "打开"} title={isEn ? "Open · Ctrl+O" : "打开 · Ctrl+O"}>
              <FolderOpen size={17} />
            </button>
            <button type="button" className="icon-button" onClick={saveDocument} aria-label={isEn ? "Save" : "保存到本地"} title={isEn ? "Save · Ctrl+S" : "保存到本地 · Ctrl+S"}>
              <Save size={17} />
            </button>
          </div>
          <div className="action-group edit-actions">
            <button
              type="button"
              className="icon-button"
              onClick={() => historyManager.requestUndo()}
              disabled={
                !historyState.canUndo ||
                historyState.isReplaying
              }
              aria-label={isEn ? "Undo" : "撤销"}
              title={isEn ? "Undo · Ctrl+Z" : "撤销 · Ctrl+Z"}
            >
              <Undo2 size={17} />
            </button>
            <button
              type="button"
              className="icon-button"
              onClick={() => historyManager.requestRedo()}
              disabled={
                !historyState.canRedo ||
                historyState.isReplaying
              }
              aria-label={isEn ? "Redo" : "重做"}
              title={isEn ? "Redo · Ctrl+Y / Ctrl+Shift+Z" : "重做 · Ctrl+Y / Ctrl+Shift+Z"}
            >
              <Redo2 size={17} />
            </button>
          </div>
          <button type="button" className="icon-button workspace-action" onClick={() => setHistoryOpen(true)} aria-label={isEn ? "Formula history" : "公式历史"} title={isEn ? "Formula history" : "公式历史"}>
            <History size={17} />
          </button>
          <button type="button" className="icon-button workspace-action" onPointerDown={captureOcrInsertionTarget} onClick={openOcrDialog} aria-label={isEn ? "Recognize formula image" : "图片公式识别"} title={isEn ? "Recognize formula image" : "图片公式识别"}>
            <ScanLine size={17} />
          </button>
          <button type="button" className="icon-button settings-toggle" onClick={() => setSettingsOpen(true)} aria-label={isEn ? "Settings" : "设置"} title={isEn ? "Settings · Ctrl+," : "设置 · Ctrl+,"}>
            <Settings2 size={17} />
          </button>
          {codeFormatControl}
          <div
            ref={setDesktopTopToolsMount}
            className="desktop-top-tools-mount"
            data-desktop-top-tools-mount
          />
        </div>
      </header>

      {(menuOpen || copyMenuOpen) && (
        <button
          type="button"
          className="menu-dismiss-layer"
          aria-label={isEn ? "Close menu" : "关闭菜单"}
          onClick={() => {
            setMenuOpen(false);
            setCopyMenuOpen(false);
          }}
        />
      )}

      <EditorWorkspace
        mode="desktop"
        showFileActions
        desktopTopToolsMount={desktopTopToolsMount}
        showUpdateActions={false}
        showOfficeActions={false}
        showOcrActions
        onOpenExport={() => setExportOpen(true)}
        editorRef={editorRef}
        sidebarOpen={sidebarOpen}
        onSidebarOpenChange={setSidebarOpen}
        onPasteImage={handleEditorImagePaste}
        onCopyPng={handleCopyPng}
        onCopy={async () => {
          await handleCopy();
        }}
        onReplaceDocument={replaceDocumentWithHistory}
        sourceDocumentRevision={sourceDocumentRevision}
        ocrBusy={inlineOcrIsBusy}
        ocrOverlay={
          inlineOcr ? (
            <div
              className={`inline-ocr-progress is-${inlineOcr.status}`}
              role="status"
              aria-live="polite"
            >
              <span className="inline-ocr-progress-icon">
                {inlineOcr.status === "running" ? (
                  <LoaderCircle size={17} className="is-spinning" />
                ) : inlineOcr.status === "success" ? (
                  <Check size={17} />
                ) : (
                  <AlertCircle size={17} />
                )}
              </span>
              <div>
                <strong>{inlineOcr.message}</strong>
                <span>
                  {isEn ? "Web OCR" : "网页 OCR"}
                  {" · "}
                  {inlineOcr.seconds}
                  {isEn ? "s" : " 秒"}
                </span>
              </div>
              {inlineOcrIsBusy ? null : (
                <button
                  type="button"
                  className="inline-ocr-dismiss"
                  onClick={() => setInlineOcr(null)}
                  aria-label={isEn ? "Dismiss OCR status" : "关闭 OCR 状态"}
                >
                  <X size={13} />
                </button>
              )}
            </div>
          ) : null
        }
      />

      <ExportDialog
        open={exportOpen}
        title={title}
        formulas={lines.map((line) => line.latex)}
        language={language}
        onClose={() => setExportOpen(false)}
        onNotify={setToast}
      />
      <HelpManualDialog
        open={helpOpen}
        language={language}
        onClose={() => setHelpOpen(false)}
      />
      <SettingsDialog
        open={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        onOpenFormulaHotkeys={() => {
          setSettingsOpen(false);
          setFormulaHotkeyManagerOpen(true);
        }}
      />
      <FormulaHotkeyManagerDialog
        open={formulaHotkeyManagerOpen}
        onClose={() => setFormulaHotkeyManagerOpen(false)}
      />
      <HistoryPanel
        open={historyOpen}
        onClose={() => setHistoryOpen(false)}
        onRestore={(item) => {
          const restored = item.lines?.length
            ? item.lines.map((line) => ({
                latex: normalizeChineseLatex(line.latex),
                mode: line.mode,
              }))
            : item.latex
                .replace(/\r\n?/g, "\n")
                .split("\n")
                .map((latex) => ({ latex: normalizeChineseLatex(latex), mode: undefined }));
          const nextLines = reconcileFormulaLines(
            restored.map((line) => line.latex),
            lines,
            restored.map((line) => line.mode),
          );
          const nextActiveLineId = nextLines.some(
            (line) => line.id === activeLineId,
          )
            ? activeLineId
            : nextLines[0]?.id ?? null;
          replaceDocumentWithHistory(
            {
              title,
              lines: nextLines,
              activeLineId: nextActiveLineId,
              formulaAlignment,
              selectionByLineId:
                editorRef.current?.getSelectionMap() ?? {},
            },
            "history-restore",
          );
          setHistoryOpen(false);
          setToast(isEn ? "Formula restored" : "已恢复历史公式");
        }}
      />
      <WebOcrDialog
        open={ocrOpen}
        language={language}
        onClose={() => {
          ocrInsertionTargetRef.current = null;
          setOcrOpen(false);
        }}
        onInsert={(value) => {
          const target = ocrInsertionTargetRef.current;
          const inserted = target
            ? editorRef.current?.insertLatexAt(target, value, "ocr")
            : false;
          if (!inserted) editorRef.current?.insertLatex(value, "ocr");
        }}
        onAppend={(value) => editorRef.current?.appendLatex(value, "ocr")}
        onNotify={setToast}
      />

      {historyOpen && (
        <div className="panel-backdrop" onClick={() => setHistoryOpen(false)} />
      )}
      {toast && (
        <div className="toast">
          <Check size={15} />
          {toast}
        </div>
      )}
    </div>
  );
}

export default App;
