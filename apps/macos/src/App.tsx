import { ChangeEvent, useCallback, useEffect, useRef, useState } from "react";
import { invoke, isTauri } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import {
  AlertCircle,
  BookOpenText,
  Braces,
  Check,
  ChevronDown,
  CircleHelp,
  Code2,
  FileDown,
  FilePlus2,
  FolderOpen,
  History,
  Languages,
  Keyboard,
  LoaderCircle,
  Menu,
  Minus,
  PanelBottomClose,
  PanelBottomOpen,
  PanelLeftClose,
  PanelLeftOpen,
  PanelRightClose,
  PanelRightOpen,
  Plus,
  Redo2,
  RefreshCw,
  Save,
  ScanLine,
  Settings2,
  Undo2,
  X,
} from "lucide-react";
import {
  MathEditor,
  type MathEditorHandle,
  type MathEditorInsertionTarget,
} from "./editor/MathEditor";
import { FormulaToolbar } from "./toolbar/FormulaToolbar";
import { LatexSourceEditor } from "./source-editor/LatexSourceEditor";
import { SettingsDialog } from "./components/SettingsDialog";
import { FormulaHotkeyManagerDialog } from "./components/FormulaHotkeyManagerDialog";
import { HistoryPanel } from "./components/HistoryPanel";
import { HelpManualDialog } from "./components/HelpManualDialog";
import { OcrDialog } from "./components/OcrDialog";
import { ExportDialog } from "./components/ExportDialog";
import { OnboardingTour } from "./components/OnboardingTour";
import { MacOfficeFirstRunPrompt } from "./components/MacOfficeFirstRunPrompt";
import { decodeMacOfflineOfficeStatus } from "./components/macOfficeStatusValidation";
import { UpdateDialog } from "./components/UpdateDialog";
import { VisualTeXLogo } from "./components/VisualTeXLogo";
import { EditorWorkspace } from "./workspace/EditorWorkspace";
import {
  EDITOR_ZOOM_STEP,
  MAX_EDITOR_ZOOM,
  MIN_EDITOR_ZOOM,
  joinFormulaLines,
  useEditorStore,
} from "./stores/editorStore";
import {
  historyManager,
  useHistorySnapshot,
} from "./history/HistoryManager";
import {
  applyHistoryEntryToEditor,
  createBlankDocumentSnapshot,
  documentSnapshotsEquivalent,
  getEditorDocumentSnapshot,
  reconcileFormulaLines,
} from "./history/documentHistory";
import type {
  DocumentSnapshot,
  ReplaceDocumentEntry,
} from "./history/historyTypes";
import {
  copyFormulaLines,
  formatFormulaLines,
  getLatexCodeFormatDefinition,
  latexCodeFormats,
  parseLatexSource,
} from "./clipboard/LatexCopyService";
import { normalizeChineseLatex } from "./editor/normalizeChineseLatex";
import type { FormulaDocument, LatexCodeFormat } from "./types/formula";
import { applyDocumentTheme, publishSynchronizedTheme } from "./themeSync";
import { copyFormulaDocumentPngToClipboard } from "./export/pngClipboard";
import {
  OCR_MODELS,
  cancelOcrRecognition,
  fileToOcrRequest,
  getOcrProviderConfiguration,
  getOcrRuntimeStatus,
  isTauriEnvironment,
  listenOcrRecognitionProgress,
  normalizeOcrFormulaLines,
  recognizeFormulaImage,
  resolveAvailableOcrModel,
  prewarmOcrModel,
  type OcrModelName,
} from "./ocr/ocrService";
import {
  checkForUpdates,
  openReleasePage,
  type UpdateCheckResult,
} from "./update/updateService";
import {
  detectDesktopPlatform,
  onboardingStorageKey,
  shouldOpenOnboardingInitially,
} from "./platform";
import {
  readLocalStorage,
  writeLocalStorage,
} from "./runtime/safeStorage";
import {
  captureQuickOcrScreenshot,
  configureSilentOcr,
  isQuickOcrCaptureMode,
  quickOcrCaptureToFile,
  QUICK_OCR_CAPTURE_MODE_STORAGE_KEY,
  SILENT_OCR_SHORTCUT,
  SILENT_OCR_STORAGE_KEY,
  waitForQuickOcrSystemScreenshot,
  type QuickOcrCaptureMode,
} from "./ocr/quickOcr";
import { useOcrQuickSelection } from "./ocr/useOcrQuickSelection";

type InlineOcrStatus = "running" | "cancelling" | "success" | "error" | "cancelled";

interface InlineOcrState {
  status: InlineOcrStatus;
  message: string;
  seconds: number;
  model: OcrModelName;
}

const DEFAULT_OCR_MODEL: OcrModelName = "PP-FormulaNet_plus-M";
const OCR_MODEL_STORAGE_KEY = "visualtex.ocr.model";
const MAC_OFFICE_FIRST_RUN_STORAGE_KEY =
  "visualtex.office.macos.native-first-run.v1.2.0.completed";
const DESKTOP_PLATFORM = detectDesktopPlatform();
const ONBOARDING_STORAGE_KEY = onboardingStorageKey(
  DESKTOP_PLATFORM,
  isTauriEnvironment(),
);

function App() {
  const editorRef = useRef<MathEditorHandle>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const menuButtonRef = useRef<HTMLButtonElement>(null);
  const appMenuRef = useRef<HTMLDivElement>(null);
  const copyMenuButtonRef = useRef<HTMLButtonElement>(null);
  const copyMenuRef = useRef<HTMLDivElement>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [formulaHotkeyManagerOpen, setFormulaHotkeyManagerOpen] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [helpManualOpen, setHelpManualOpen] = useState(false);
  const [ocrOpen, setOcrOpen] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(() => window.innerWidth >= 1040);
  const isMacDesktop = DESKTOP_PLATFORM === "macos" && isTauriEnvironment();
  // macOS Office setup/update mode must be chosen from the real DOTM/PPAM
  // status first. Opening the legacy first-run UI before that check can turn a
  // repair/update into a false "register PowerPoint again" prompt.
  const [macOfficeFirstRunOpen, setMacOfficeFirstRunOpen] = useState(false);
  const [macOfficePromptMode, setMacOfficePromptMode] = useState<
    "setup" | "update" | "repair"
  >("setup");
  const [powerpointRegistrationRequired, setPowerpointRegistrationRequired] =
    useState(false);
  const [onboardingOpen, setOnboardingOpen] = useState(() =>
    isMacDesktop
      ? false
      : shouldOpenOnboardingInitially(
          readLocalStorage(ONBOARDING_STORAGE_KEY) === "true",
          false,
        ),
  );
  const [copyMenuOpen, setCopyMenuOpen] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [keypadMode, setKeypadMode] = useState(false);
  const keypadModeSwitchPendingRef = useRef(false);
  const [updateOpen, setUpdateOpen] = useState(false);
  const [updateChecking, setUpdateChecking] = useState(false);
  const [updateError, setUpdateError] = useState("");
  const [updateResult, setUpdateResult] = useState<UpdateCheckResult | null>(null);
  const [automaticUpdatePrompt, setAutomaticUpdatePrompt] = useState(false);
  const [toast, setToast] = useState("");
  const [savedPulse, setSavedPulse] = useState(false);
  const [editorHistoryBusy, setEditorHistoryBusy] = useState(false);
  const [ocrModel, setOcrModel] = useState<OcrModelName>(() => {
    const stored = readLocalStorage(OCR_MODEL_STORAGE_KEY);
    return OCR_MODELS.some((item) => item.id === stored)
      ? (stored as OcrModelName)
      : DEFAULT_OCR_MODEL;
  });
  const [silentOcrEnabled, setSilentOcrEnabled] = useState(
    () => readLocalStorage(SILENT_OCR_STORAGE_KEY) === "true",
  );
  const [quickOcrCaptureMode, setQuickOcrCaptureMode] = useState<QuickOcrCaptureMode>(() => {
    const stored = readLocalStorage(QUICK_OCR_CAPTURE_MODE_STORAGE_KEY);
    return isQuickOcrCaptureMode(stored) ? stored : "immediate";
  });
  const [quickOcrCaptureBusy, setQuickOcrCaptureBusy] = useState(false);
  const [inlineOcr, setInlineOcr] = useState<InlineOcrState | null>(null);
  const inlineOcrBusyRef = useRef(false);
  const inlineOcrCancelRequestedRef = useRef(false);
  const inlineOcrRunIdRef = useRef(0);
  const inlineOcrClearTimerRef = useRef<number | null>(null);
  const automaticUpdateCheckRef = useRef(false);
  const ocrPrewarmStartedRef = useRef(false);
  const macOfficeInstallStatusCheckedRef = useRef(false);
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
  const editorLayout = useEditorStore((state) => state.editorLayout);
  const pngExportBackground = useEditorStore(
    (state) => state.pngExportBackground,
  );
  const formulaLetterFont = useEditorStore((state) => state.formulaLetterFont);
  const formulaChineseFont = useEditorStore((state) => state.formulaChineseFont);
  const sourceOpen = useEditorStore((state) => state.sourceOpen);
  const setSourceOpen = useEditorStore((state) => state.setSourceOpen);
  const latexCodeFormat = useEditorStore((state) => state.latexCodeFormat);
  const setLatexCodeFormat = useEditorStore(
    (state) => state.setLatexCodeFormat,
  );
  const addHistory = useEditorStore((state) => state.addHistory);
  const loadDocument = useEditorStore((state) => state.loadDocument);
  const toDocument = useEditorStore((state) => state.toDocument);
  const checkUpdatesOnStartup = useEditorStore(
    (state) => state.checkUpdatesOnStartup,
  );
  const keypadMinimizeOnCopy = useEditorStore(
    (state) => state.keypadMinimizeOnCopy,
  );
  const setCheckUpdatesOnStartup = useEditorStore(
    (state) => state.setCheckUpdatesOnStartup,
  );
  const historyState = useHistorySnapshot();
  const isEn = language === "en";
  const latex = joinFormulaLines(lines);
  const sourceLatex = formatFormulaLines(lines, latexCodeFormat);
  const currentCodeFormat = getLatexCodeFormatDefinition(latexCodeFormat);
  const codeFormatGroups = [
    {
      id: "single" as const,
      title: isEn ? "Independent formula formats" : "单公式独立环境",
      description: isEn
        ? "Each non-empty formula field gets its own wrapper"
        : "每个非空公式框分别生成一个完整环境",
      formats: latexCodeFormats.filter((format) => format.group === "single"),
    },
    {
      id: "multi" as const,
      title: isEn ? "Combined multi-line environments" : "多公式合并环境",
      description: isEn
        ? "All non-empty formula fields become rows in one environment"
        : "所有非空公式框合并成一个多行公式环境",
      formats: latexCodeFormats.filter((format) => format.group === "multi"),
    },
  ];
  const selectedOcrModel =
    OCR_MODELS.find((item) => item.id === ocrModel) ?? OCR_MODELS[1];
  const inlineOcrModel =
    OCR_MODELS.find((item) => item.id === inlineOcr?.model) ?? selectedOcrModel;
  const inlineOcrIsBusy =
    inlineOcr?.status === "running" || inlineOcr?.status === "cancelling";

  const captureDocumentSnapshot = (): DocumentSnapshot =>
    getEditorDocumentSnapshot(editorRef.current?.getSelectionMap() ?? {});

  const restoreSnapshotFocus = (snapshot: DocumentSnapshot) => {
    const lineId = snapshot.activeLineId;
    if (!lineId) return;
    const line = snapshot.lines.find((item) => item.id === lineId);
    if (!line) return;
    void editorRef.current?.restoreSelection(
      lineId,
      line.latex,
      snapshot.selectionByLineId[lineId] ?? null,
    );
  };

  const replaceDocumentWithHistory = (
    after: DocumentSnapshot,
    source: ReplaceDocumentEntry["source"],
  ) => {
    if (source !== "source-apply") historyManager.commitPendingTransaction();
    const before = captureDocumentSnapshot();
    if (documentSnapshotsEquivalent(before, after)) return false;
    useEditorStore.getState().replaceDocumentState(after);
    const entry: ReplaceDocumentEntry = {
      type: "replace-document",
      before,
      after,
      source,
      timestamp: Date.now(),
    };
    if (source === "source-apply") {
      historyManager.recordSourceDocumentEdit(entry);
    } else {
      historyManager.push(entry);
      window.requestAnimationFrame(() => restoreSnapshotFocus(after));
    }
    return true;
  };

  const handleKeypadModeToggle = () => {
    if (keypadModeSwitchPendingRef.current) return;
    setMenuOpen(false);
    setCopyMenuOpen(false);

    const nextKeypadMode = !keypadMode;
    if (!isTauri()) {
      setKeypadMode(nextKeypadMode);
      window.requestAnimationFrame(() =>
        editorRef.current?.focus({ target: "last", moveToEnd: false }),
      );
      return;
    }

    keypadModeSwitchPendingRef.current = true;
    void invoke("switch_main_window_mode", { keypad: nextKeypadMode })
      .then(() => {
        setKeypadMode(nextKeypadMode);
        window.requestAnimationFrame(() =>
          editorRef.current?.focus({ target: "last", moveToEnd: false }),
        );
      })
      .catch((error) => {
        setToast(
          isEn
            ? `Unable to switch keypad window size: ${String(error)}`
            : `小键盘窗口尺寸切换失败：${String(error)}`,
        );
      })
      .finally(() => {
        keypadModeSwitchPendingRef.current = false;
      });
  };

  useEffect(() => {
    if (
      onboardingOpen ||
      macOfficeFirstRunOpen ||
      initialEditorFocusDoneRef.current
    ) {
      return;
    }
    initialEditorFocusDoneRef.current = true;
    let repairTimer = 0;
    const frame = window.requestAnimationFrame(() => {
      editorRef.current?.focus({ target: "last", moveToEnd: true });
      repairTimer = window.setTimeout(() => {
        editorRef.current?.focus({ target: "last", moveToEnd: true });
      }, 80);
    });
    return () => {
      window.cancelAnimationFrame(frame);
      window.clearTimeout(repairTimer);
    };
  }, [macOfficeFirstRunOpen, onboardingOpen]);

  useEffect(() => {
    historyManager.configure({
      getDocumentSnapshot: () =>
        getEditorDocumentSnapshot(editorRef.current?.getSelectionMap() ?? {}),
      applyEntry: async (entry, direction) => {
        const target = applyHistoryEntryToEditor(entry, direction);
        if (!target) return;
        // Yield once so React can mount any line restored by the history entry.
        // Do not wait on requestAnimationFrame here: background macOS windows
        // and headless release checks can throttle animation frames indefinitely,
        // leaving history replay active and the Redo action disabled.
        await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
        await editorRef.current?.restoreSelection(
          target.lineId,
          target.latex,
          target.selection,
        );
      },
    });
    return () => historyManager.configure(null);
  }, []);

  useEffect(() => {
    if (
      DESKTOP_PLATFORM !== "macos" ||
      !isTauriEnvironment() ||
      macOfficeInstallStatusCheckedRef.current
    ) {
      return;
    }
    macOfficeInstallStatusCheckedRef.current = true;
    let cancelled = false;

    void invoke<unknown>("get_macos_offline_office_install_status")
      .then(decodeMacOfflineOfficeStatus)
      .then((status) => {
        if (cancelled) return;
        if (!status.compiledArtifactsAvailable) {
          if (readLocalStorage(ONBOARDING_STORAGE_KEY) !== "true") {
            setOnboardingOpen(true);
          }
          return;
        }

        const wordNeedsCurrentAddin =
          status.word.applicationInstalled && !status.word.filesInstalled;
        const powerpointNeedsCurrentAddin =
          status.powerpoint.applicationInstalled &&
          !status.powerpoint.filesInstalled;
        const needsCurrentAddins =
          wordNeedsCurrentAddin || powerpointNeedsCurrentAddin;

        // Remember whether PowerPoint genuinely has never had a PPAM at the
        // start of this flow. A current PPAM that merely is not running must not
        // be turned into a false "register once" requirement because Word needs
        // repair or because the first-run storage flag is stale.
        setPowerpointRegistrationRequired(
          status.powerpoint.applicationInstalled &&
            !status.powerpoint.filesPresent,
        );

        if (!needsCurrentAddins) {
          setMacOfficeFirstRunOpen(false);
          if (readLocalStorage(ONBOARDING_STORAGE_KEY) !== "true") {
            setOnboardingOpen(true);
          }
          return;
        }

        const staleInstalledAddins =
          (wordNeedsCurrentAddin && status.word.filesPresent) ||
          (powerpointNeedsCurrentAddin && status.powerpoint.filesPresent);
        const previouslyConfigured =
          readLocalStorage(MAC_OFFICE_FIRST_RUN_STORAGE_KEY) === "true" ||
          status.word.filesPresent ||
          status.powerpoint.filesPresent;

        setOnboardingOpen(false);
        if (staleInstalledAddins) {
          setMacOfficePromptMode("update");
        } else if (previouslyConfigured) {
          setMacOfficePromptMode("repair");
        } else {
          setMacOfficePromptMode("setup");
        }
        // Never mutate DOTM/PPAM silently. The user sees whether this is first
        // setup, a version update, or a repair of missing files before install.
        setMacOfficeFirstRunOpen(true);
      })
      .catch(() => {
        if (cancelled) return;
        // A status failure is not evidence of a version update. Surface it as a
        // repair/inspection flow rather than misleading the user with setup or
        // update-specific PowerPoint registration guidance.
        setPowerpointRegistrationRequired(false);
        setMacOfficePromptMode("repair");
        setOnboardingOpen(false);
        setMacOfficeFirstRunOpen(true);
      });

    return () => {
      cancelled = true;
    };
  }, [isEn]);

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
    if (isTauriEnvironment()) {
      void invoke<string>("set_app_theme", { theme }).catch(() => undefined);
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
    if (inlineOcrStatus !== "running" && inlineOcrStatus !== "cancelling") {
      return;
    }
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

  useEffect(() => {
    if (!isTauriEnvironment()) return;

    let cancelled = false;
    const delay = ocrPrewarmStartedRef.current ? 250 : 1200;
    const timer = window.setTimeout(() => {
      ocrPrewarmStartedRef.current = true;
      void getOcrProviderConfiguration()
        .then((provider) => {
          if (cancelled || provider.activeProvider !== "local") return null;
          return getOcrRuntimeStatus();
        })
        .then((runtime) => {
          if (cancelled || !runtime?.installed) return;
          const availableModel = resolveAvailableOcrModel(runtime, ocrModel);
          return prewarmOcrModel(availableModel);
        })
        .catch(() => undefined);
    }, delay);

    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [ocrModel]);

  useEffect(() => {
    if (!isTauri()) return;
    let cancelled = false;
    void configureSilentOcr(silentOcrEnabled, ocrModel, latexCodeFormat).catch((error) => {
      if (cancelled || !silentOcrEnabled) return;
      const message =
        error instanceof Error
          ? error.message
          : typeof error === "string"
            ? error
            : isEn
              ? "Unable to enable silent OCR"
              : "无法启用静默 OCR";
      setSilentOcrEnabled(false);
      writeLocalStorage(SILENT_OCR_STORAGE_KEY, "false");
      setToast(message);
    });
    return () => {
      cancelled = true;
    };
  }, [silentOcrEnabled, ocrModel, latexCodeFormat, isEn]);

  const handleOcrModelChange = useCallback(
    (nextModel: OcrModelName) => {
      if (inlineOcrBusyRef.current || nextModel === ocrModel) return;
      setOcrModel(nextModel);
      writeLocalStorage(OCR_MODEL_STORAGE_KEY, nextModel);
    },
    [ocrModel],
  );

  const ocrQuickSelection = useOcrQuickSelection({
    model: ocrModel,
    busy: inlineOcrIsBusy || quickOcrCaptureBusy,
    isEn,
    onModelChange: handleOcrModelChange,
    onError: setToast,
  });

  const cancelInlineOcr = async () => {
    if (!inlineOcrBusyRef.current) return;
    inlineOcrCancelRequestedRef.current = true;
    setInlineOcr((current) =>
      current
        ? {
            ...current,
            status: "cancelling",
            message: isEn ? "Cancelling OCR…" : "正在取消 OCR…",
          }
        : current,
    );
    try {
      await cancelOcrRecognition();
    } catch {
      // The recognition promise will surface the final state. A worker that
      // already exited is equivalent to a successful cancellation.
    }
  };

  const handleEditorImagePaste = async (
    file: File,
    target: MathEditorInsertionTarget,
    source: "paste" | "quick" = "paste",
  ) => {
    if (inlineOcrBusyRef.current) {
      setToast(isEn ? "Another pasted image is being recognized" : "已有一张粘贴图片正在识别");
      return;
    }
    if (!isTauriEnvironment()) {
      setToast(isEn ? "Image OCR is available in the desktop app" : "图片 OCR 只能在桌面应用中使用");
      return;
    }

    if (inlineOcrClearTimerRef.current !== null) {
      window.clearTimeout(inlineOcrClearTimerRef.current);
      inlineOcrClearTimerRef.current = null;
    }

    const runId = ++inlineOcrRunIdRef.current;
    inlineOcrBusyRef.current = true;
    inlineOcrCancelRequestedRef.current = false;
    setInlineOcr({
      status: "running",
      message: isEn ? "Checking the OCR provider…" : "正在检查 OCR 提供器…",
      seconds: 0,
      model: ocrModel,
    });

    let unlisten: (() => void) | undefined;
    try {
      const providerConfiguration = await getOcrProviderConfiguration();
      const usingLocalProvider = providerConfiguration.activeProvider === "local";
      if (inlineOcrCancelRequestedRef.current) throw new Error("OCR_CANCELLED");
      if (usingLocalProvider) {
        const runtime = await getOcrRuntimeStatus();
        if (inlineOcrCancelRequestedRef.current) throw new Error("OCR_CANCELLED");
        if (!runtime.installed) {
          setOcrOpen(true);
          throw new Error(
            isEn
              ? "Install the OCR runtime before pasting an image"
              : "请先安装 OCR 运行环境，再在公式框中粘贴图片",
          );
        }

        if (!runtime.installedModels.includes(ocrModel)) {
          setOcrOpen(true);
          throw new Error(
            isEn
              ? `Install ${selectedOcrModel.labelEn} before using it for OCR`
              : `请先安装${selectedOcrModel.labelZh}模型，再使用该模型进行 OCR`,
          );
        }
      } else {
        setInlineOcr((current) =>
          current
            ? {
                ...current,
                message: isEn
                  ? "Sending the image to the configured OCR API…"
                  : "正在将图片发送到已配置的 OCR API…",
              }
            : current,
        );
      }
      const availableOcrModel = ocrModel;

      unlisten = await listenOcrRecognitionProgress((progress) => {
        if (
          inlineOcrRunIdRef.current !== runId ||
          progress.model !== ocrModel
        ) {
          return;
        }
        setInlineOcr((current) =>
          current
            ? {
                ...current,
                message: progress.message,
              }
            : current,
        );
      });

      const request = await fileToOcrRequest(file, availableOcrModel);
      if (inlineOcrCancelRequestedRef.current) throw new Error("OCR_CANCELLED");
      const result = await recognizeFormulaImage(request);
      if (
        inlineOcrCancelRequestedRef.current ||
        inlineOcrRunIdRef.current !== runId
      ) {
        throw new Error("OCR_CANCELLED");
      }

      const recognizedLatex = normalizeOcrFormulaLines(result.formulas).join("\n");
      if (!recognizedLatex) {
        throw new Error(isEn ? "OCR returned an empty formula" : "OCR 没有返回可用公式");
      }

      const inserted =
        editorRef.current?.insertLatexAt(target, recognizedLatex, "ocr") ?? false;
      if (!inserted) {
        throw new Error(
          isEn
            ? "The original formula line no longer exists; the OCR result was not inserted"
            : "原来的公式行已被删除，OCR 结果没有插入到其他位置",
        );
      }

      setInlineOcr((current) => ({
        status: "success",
        message: result.backgroundInverted
          ? isEn
            ? "Recognized and inserted · dark background inverted"
            : "识别完成并已插入 · 已自动反色"
          : isEn
            ? "Recognized and inserted at the saved cursor"
            : "识别完成，已插入原光标位置",
        seconds: current?.seconds ?? 0,
        model: ocrModel,
      }));
      setToast(
        source === "quick"
          ? isEn
            ? "Screenshot converted to LaTeX"
            : "截图已转换为 LaTeX"
          : isEn
            ? "Pasted image converted to LaTeX"
            : "粘贴图片已转换为 LaTeX",
      );
      scheduleInlineOcrClear(1800);
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : typeof error === "string" ? error : "";
      const cancelled =
        inlineOcrCancelRequestedRef.current || errorMessage.includes("OCR_CANCELLED");
      if (cancelled) {
        setInlineOcr((current) => ({
          status: "cancelled",
          message: isEn ? "OCR cancelled" : "OCR 已取消",
          seconds: current?.seconds ?? 0,
          model: ocrModel,
        }));
        scheduleInlineOcrClear(1200);
      } else {
        const message =
          errorMessage || (isEn ? "Image OCR failed" : "图片 OCR 失败");
        setInlineOcr((current) => ({
          status: "error",
          message,
          seconds: current?.seconds ?? 0,
          model: ocrModel,
        }));
        setToast(message);
        scheduleInlineOcrClear(4500);
      }
    } finally {
      unlisten?.();
      if (inlineOcrRunIdRef.current === runId) {
        inlineOcrBusyRef.current = false;
        inlineOcrCancelRequestedRef.current = false;
      }
    }
  };

  const handleQuickOcr = async () => {
    if (inlineOcrBusyRef.current || quickOcrCaptureBusy) {
      setToast(isEn ? "OCR is already running" : "已有 OCR 任务正在进行");
      return;
    }
    const target = editorRef.current?.captureInsertionTarget();
    if (!target) {
      setToast(isEn ? "Click a formula field before quick OCR" : "请先点击一个公式输入框再使用快捷 OCR");
      return;
    }
    setQuickOcrCaptureBusy(true);
    try {
      const waitingForSystemScreenshot = quickOcrCaptureMode === "system-screenshot";
      setToast(
        waitingForSystemScreenshot
          ? isEn
            ? "VisualTeX will minimize. Switch to the target page and take a macOS screenshot within 60 seconds."
            : "VisualTeX 将最小化；请切到目标页面，并在 60 秒内使用 macOS 系统截图键截图"
          : isEn
            ? "Select a formula area to capture"
            : "请框选要识别的公式区域",
      );
      const capture = waitingForSystemScreenshot
        ? await waitForQuickOcrSystemScreenshot()
        : await captureQuickOcrScreenshot();
      if (!capture) {
        setToast(
          waitingForSystemScreenshot
            ? isEn
              ? "No system screenshot was detected within 60 seconds"
              : "60 秒内未检测到新的系统截图"
            : isEn
              ? "Screenshot cancelled"
              : "已取消截图",
        );
        return;
      }
      await handleEditorImagePaste(quickOcrCaptureToFile(capture), target, "quick");
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : typeof error === "string"
            ? error
            : isEn
              ? "Quick OCR failed"
              : "快捷 OCR 失败";
      setToast(message);
    } finally {
      setQuickOcrCaptureBusy(false);
    }
  };

  const handleQuickOcrCaptureModeChange = (mode: QuickOcrCaptureMode) => {
    setQuickOcrCaptureMode(mode);
    writeLocalStorage(QUICK_OCR_CAPTURE_MODE_STORAGE_KEY, mode);
    setToast(
      mode === "system-screenshot"
        ? isEn
          ? "Quick OCR mode: wait for macOS screenshot"
          : "快捷 OCR 已切换为：等待系统截图"
        : isEn
          ? "Quick OCR mode: immediate selection"
          : "快捷 OCR 已切换为：立即框选",
    );
  };

  const handleSilentOcrEnabledChange = (enabled: boolean) => {
    setSilentOcrEnabled(enabled);
    writeLocalStorage(SILENT_OCR_STORAGE_KEY, String(enabled));
    setToast(
      enabled
        ? isEn
          ? `Silent OCR enabled · ${SILENT_OCR_SHORTCUT}`
          : `静默 OCR 已开启 · ${SILENT_OCR_SHORTCUT}`
        : isEn
          ? "Silent OCR disabled"
          : "静默 OCR 已关闭",
    );
  };

  const handleCodeFormatChange = (format: LatexCodeFormat) => {
    const definition = getLatexCodeFormatDefinition(format);
    setLatexCodeFormat(format);
    if (!keypadMode) setSourceOpen(true);
    setCopyMenuOpen(false);
    setToast(
      isEn
        ? `LaTeX code format: ${definition.titleEn}`
        : `LaTeX 代码格式已切换为：${definition.titleZh}`,
    );
  };

  const handleCopy = async () => {
    try {
      await copyFormulaLines(lines, latexCodeFormat);
      addHistory(latex);
      setToast(
        isEn
          ? `Copied ${currentCodeFormat.titleEn}`
          : `已复制：${currentCodeFormat.titleZh}`,
      );
      return true;
    } catch {
      setToast(
        isEn
          ? "Copy failed. Check clipboard permission."
          : "复制失败，请检查系统剪贴板权限",
      );
      return false;
    }
  };

  const handleKeypadCopy = async () => {
    const copied = await handleCopy();
    if (!copied || !keypadMinimizeOnCopy || !isTauri()) return;
    try {
      await getCurrentWindow().minimize();
    } catch {
      setToast(
        isEn
          ? "LaTeX copied, but VisualTeX could not be minimized."
          : "LaTeX 已复制，但 VisualTeX 最小化失败。",
      );
    }
  };

  const handleCopyPng = async () => {
    if (pngClipboardBusyRef.current) return;
    pngClipboardBusyRef.current = true;
    try {
      await copyFormulaDocumentPngToClipboard(lines.map((line) => line.latex), {
        background: pngExportBackground,
        formulaLetterFont,
        formulaChineseFont,
      });
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

  const getSafeDocumentTitle = () =>
    title.trim().replace(/[\\/:*?"<>|]/g, "-") ||
    (isEn ? "Untitled Formula" : "未命名公式");

  const downloadBlobFile = (blob: Blob, filename: string) => {
    const url = URL.createObjectURL(blob);
    const link = window.document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  };

  const downloadTextFile = (content: string, filename: string, type: string) => {
    downloadBlobFile(new Blob([content], { type }), filename);
  };

  const saveDocument = () => {
    historyManager.commitPendingTransaction();
    void historyManager.createCheckpoint("save-document").catch(() => undefined);
    const document = toDocument();
    downloadTextFile(
      JSON.stringify(document, null, 2),
      `${getSafeDocumentTitle()}.visualtex.json`,
      "application/json;charset=utf-8",
    );
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
      historyManager.commitPendingTransaction();
      const before = captureDocumentSnapshot();
      loadDocument(parsed);
      const after = getEditorDocumentSnapshot({});
      if (!documentSnapshotsEquivalent(before, after)) {
        historyManager.push({
          type: "replace-document",
          before,
          after,
          source: "open-document",
          timestamp: Date.now(),
        });
        window.requestAnimationFrame(() => restoreSnapshotFocus(after));
      }
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

  const finishMacOfficeFirstRun = useCallback((installed: boolean) => {
    writeLocalStorage(MAC_OFFICE_FIRST_RUN_STORAGE_KEY, "true");
    setMacOfficeFirstRunOpen(false);
    if (readLocalStorage(ONBOARDING_STORAGE_KEY) !== "true") {
      setOnboardingOpen(true);
    }
    setToast(
      installed
        ? macOfficePromptMode === "update"
          ? isEn
            ? "Office add-ins were updated to this VisualTeX version"
            : "Office 插件已更新到当前 VisualTeX 版本"
          : macOfficePromptMode === "repair"
            ? isEn
              ? "Office add-ins were repaired"
              : "Office 插件已修复"
            : isEn
              ? "Native Office add-ins are ready"
              : "原生 Office 加载项已就绪"
        : isEn
          ? "You can finish native Office setup later in Settings"
          : "之后可在设置中完成原生 Office 插件配置",
    );
  }, [isEn, macOfficePromptMode]);

  const finishOnboarding = useCallback(() => {
    writeLocalStorage(ONBOARDING_STORAGE_KEY, "true");
    setOnboardingOpen(false);
    window.requestAnimationFrame(() => editorRef.current?.focus());
  }, []);

  const runUpdateCheck = useCallback(async (manual = true) => {
    if (manual) {
      setAutomaticUpdatePrompt(false);
      setUpdateResult(null);
      setUpdateOpen(true);
    }
    setUpdateChecking(true);
    setUpdateError("");
    try {
      const result = await checkForUpdates();
      setUpdateResult(result);
      if (manual || result.updateAvailable) {
        setAutomaticUpdatePrompt(!manual && result.updateAvailable);
        setUpdateOpen(true);
      }
    } catch (error) {
      if (manual) {
        setUpdateError(
          error instanceof Error
            ? error.message
            : isEn
              ? "Unable to connect to the update server"
              : "无法连接更新服务器",
        );
        setUpdateOpen(true);
      } else {
        automaticUpdateCheckRef.current = false;
      }
    } finally {
      setUpdateChecking(false);
    }
  }, [isEn]);

  useEffect(() => {
    if (
      !checkUpdatesOnStartup ||
      macOfficeFirstRunOpen ||
      onboardingOpen ||
      automaticUpdateCheckRef.current
    ) {
      return;
    }

    let timer = 0;
    const runWhenOnline = () => {
      if (
        automaticUpdateCheckRef.current ||
        !useEditorStore.getState().checkUpdatesOnStartup
      ) {
        return;
      }
      automaticUpdateCheckRef.current = true;
      timer = window.setTimeout(() => void runUpdateCheck(false), 1200);
    };

    window.addEventListener("online", runWhenOnline);
    if (navigator.onLine) runWhenOnline();

    return () => {
      window.clearTimeout(timer);
      window.removeEventListener("online", runWhenOnline);
    };
  }, [checkUpdatesOnStartup, macOfficeFirstRunOpen, onboardingOpen, runUpdateCheck]);

  useEffect(() => {
    const handleWindowKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setMenuOpen(false);
        setCopyMenuOpen(false);
        return;
      }

      if (settingsOpen || formulaHotkeyManagerOpen || helpManualOpen || ocrOpen || historyOpen || exportOpen || macOfficeFirstRunOpen || onboardingOpen || updateOpen) {
        return;
      }

      const target = event.target instanceof Element ? event.target : null;
      const inCodeMirror = Boolean(target?.closest(".cm-editor"));
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
        if (keypadMode) void handleKeypadCopy();
        else saveDocument();
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
  }, [latex, title, isEn, zoom, keypadMode, keypadMinimizeOnCopy, latexCodeFormat, settingsOpen, formulaHotkeyManagerOpen, helpManualOpen, ocrOpen, historyOpen, exportOpen, macOfficeFirstRunOpen, onboardingOpen, updateOpen]);

  const codeFormatControl = (
      <div
        className={
          "copy-control code-format-control" +
          (keypadMode ? " editor-code-format-control" : "")
        }
      >
        <button
          type="button"
          className="copy-primary code-format-primary"
          aria-expanded={copyMenuOpen}
          aria-haspopup="menu"
          aria-controls="copy-format-menu"
          title={
            isEn
              ? `Current: ${currentCodeFormat.titleEn}`
              : `当前格式：${currentCodeFormat.titleZh}`
          }
          onClick={() => {
            setMenuOpen(false);
            setCopyMenuOpen((open) => !open);
          }}
        >
          <Code2 size={15} />
          <span>{isEn ? currentCodeFormat.titleEn : currentCodeFormat.titleZh}</span>
        </button>
        <button
          ref={copyMenuButtonRef}
          type="button"
          className="copy-chevron"
          aria-label={isEn ? "Choose LaTeX code format" : "选择 LaTeX 代码格式"}
          aria-expanded={copyMenuOpen}
          aria-haspopup="menu"
          aria-controls="copy-format-menu"
          onClick={() => {
            setMenuOpen(false);
            setCopyMenuOpen((open) => !open);
          }}
        >
          <ChevronDown size={14} />
        </button>
        {copyMenuOpen && (
          <div
            ref={copyMenuRef}
            id="copy-format-menu"
            className="copy-menu code-format-menu"
            role="menu"
            aria-label={isEn ? "LaTeX code format" : "LaTeX 代码格式"}
          >
            <div className="code-format-menu-header">
              <span className="copy-menu-label">
                {isEn ? "LaTeX code format" : "LaTeX 代码格式"}
              </span>
              <small>
                {isEn
                  ? "Controls source rendering and keypad copy output"
                  : "控制源码显示与小键盘模式复制格式"}
              </small>
            </div>
            {codeFormatGroups.map((group) => (
              <div
                className="code-format-group"
                role="group"
                aria-label={group.title}
                key={group.id}
              >
                <div className="code-format-group-heading">
                  <strong>{group.title}</strong>
                  <small>{group.description}</small>
                </div>
                {group.formats.map((format) => {
                  const selected = format.id === latexCodeFormat;
                  return (
                    <button
                      type="button"
                      role="menuitemradio"
                      aria-checked={selected}
                      aria-label={`${isEn ? format.titleEn : format.titleZh}: ${format.hint}`}
                      data-format={format.id}
                      className={selected ? "is-selected" : ""}
                      key={format.id}
                      onClick={() => handleCodeFormatChange(format.id)}
                    >
                      <span className="code-format-item-copy">
                        <small className="code-format-hint">{format.hint}</small>
                      </span>
                      <span className="code-format-check" aria-hidden="true">
                        {selected && <Check size={14} />}
                      </span>
                    </button>
                  );
                })}
              </div>
            ))}
          </div>
        )}
      </div>
  );

  const keypadToggleControl = (
    <button
      type="button"
      className={`keypad-mode-toggle${keypadMode ? " is-active" : ""}`}
      aria-pressed={keypadMode}
      data-keypad-mode-toggle
      title={
        keypadMode
          ? isEn
            ? "Exit keypad mode"
            : "退出小键盘模式"
          : isEn
            ? "Enter keypad mode"
            : "进入小键盘模式"
      }
      onClick={() => void handleKeypadModeToggle()}
    >
      <Keyboard size={15} />
      <span>{isEn ? "Keypad" : "小键盘"}</span>
    </button>
  );

  const desktopHeaderControls = (
    <>
      {keypadMode ? codeFormatControl : null}
      {keypadToggleControl}
    </>
  );

  return (
    <div className={`app-shell${keypadMode ? " is-keypad-mode" : ""}`}>
      <input
        ref={fileInputRef}
        type="file"
        accept=".json,.visualtex"
        className="visually-hidden"
        onChange={openDocument}
      />

      {!keypadMode && (
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
          <button
            type="button"
            className={"icon-button sidebar-toggle " + (sidebarOpen ? "is-active" : "")}
            aria-label={
              editorLayout === "classic"
                ? sidebarOpen
                  ? isEn
                    ? "Hide formula tiles"
                    : "隐藏公式磁贴"
                  : isEn
                    ? "Show formula tiles"
                    : "显示公式磁贴"
                : sidebarOpen
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
            {editorLayout === "classic" ? (
              sidebarOpen ? <PanelRightClose size={17} /> : <PanelRightOpen size={17} />
            ) : sidebarOpen ? (
              <PanelLeftClose size={17} />
            ) : (
              <PanelLeftOpen size={17} />
            )}
          </button>
          <div className="brand-mark" aria-hidden="true">
            <VisualTeXLogo className="visualtex-brand-logo" />
          </div>
          <div className="brand-copy">
            <strong>VisualTeX</strong>
          </div>

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
                <span>{isEn ? "Formula workspace" : "公式工作区"}</span>
              </div>
              <button type="button" role="menuitem" onClick={() => runMenuAction(newFormula)}>
                <FilePlus2 size={16} />
                <span>{isEn ? "New formula" : "新建公式"}</span>
                <kbd>⌘N</kbd>
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
                <kbd>⌘O</kbd>
              </button>
              <button type="button" role="menuitem" onClick={() => runMenuAction(saveDocument)}>
                <Save size={16} />
                <span>{isEn ? "Save document" : "保存文档"}</span>
                <kbd>⌘S</kbd>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setExportOpen(true))}
              >
                <FileDown size={16} />
                <span>{isEn ? "Export…" : "导出…"}</span>
                <kbd>MD/SVG/PNG</kbd>
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
                onClick={() => runMenuAction(() => setOcrOpen(true))}
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
                onClick={() => runMenuAction(() => setOnboardingOpen(true))}
              >
                <CircleHelp size={16} />
                <span>{isEn ? "Quick tour" : "新手教程"}</span>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => setHelpManualOpen(true))}
              >
                <BookOpenText size={16} />
                <span>{isEn ? "Help manual" : "帮助手册"}</span>
              </button>
              <button
                type="button"
                role="menuitem"
                onClick={() => runMenuAction(() => void runUpdateCheck(true))}
              >
                <RefreshCw size={16} />
                <span>{isEn ? "Check for updates" : "检查更新"}</span>
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
            <button type="button" className="icon-button" onClick={newFormula} aria-label={isEn ? "New" : "新建"} title={isEn ? "New · ⌘N" : "新建 · ⌘N"}>
              <FilePlus2 size={17} />
            </button>
            <button type="button" className="icon-button" onClick={() => fileInputRef.current?.click()} aria-label={isEn ? "Open" : "打开"} title={isEn ? "Open · ⌘O" : "打开 · ⌘O"}>
              <FolderOpen size={17} />
            </button>
            <button type="button" className="icon-button" onClick={saveDocument} aria-label={isEn ? "Save" : "保存到本地"} title={isEn ? "Save · ⌘S" : "保存到本地 · ⌘S"}>
              <Save size={17} />
            </button>
          </div>
          <div className="action-group edit-actions">
            <button
              type="button"
              className="icon-button"
              onClick={() => historyManager.requestUndo()}
              disabled={
                editorHistoryBusy ||
                !historyState.canUndo ||
                historyState.isReplaying
              }
              aria-label={isEn ? "Undo" : "撤销"}
              title={isEn ? "Undo · ⌘/Ctrl+Z" : "撤销 · ⌘/Ctrl+Z"}
            >
              <Undo2 size={17} />
            </button>
            <button
              type="button"
              className="icon-button"
              onClick={() => historyManager.requestRedo()}
              disabled={
                editorHistoryBusy ||
                !historyState.canRedo ||
                historyState.isReplaying
              }
              aria-label={isEn ? "Redo" : "重做"}
              title={isEn ? "Redo · ⇧⌘Z / Ctrl+Y" : "重做 · ⇧⌘Z / Ctrl+Y"}
            >
              <Redo2 size={17} />
            </button>
          </div>
          <button type="button" className="icon-button workspace-action" onClick={() => setHistoryOpen(true)} aria-label={isEn ? "Formula history" : "公式历史"} title={isEn ? "Formula history" : "公式历史"}>
            <History size={17} />
          </button>
          <button type="button" className="icon-button workspace-action" onClick={() => setOcrOpen(true)} aria-label={isEn ? "Recognize formula image" : "图片公式识别"} title={isEn ? "Recognize formula image" : "图片公式识别"}>
            <ScanLine size={17} />
          </button>
          <button type="button" className="icon-button settings-toggle" onClick={() => setSettingsOpen(true)} aria-label={isEn ? "Settings" : "设置"} title={isEn ? "Settings · ⌘," : "设置 · ⌘,"}>
            <Settings2 size={17} />
          </button>
          {codeFormatControl}
        </div>
      </header>
      )}

      {(menuOpen || (!keypadMode && copyMenuOpen)) && (
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
        desktopHeaderControls={desktopHeaderControls}
        keypadMode={keypadMode}
        showUpdateActions
        showOfficeActions={false}
        showOcrActions
        onOpenExport={() => setExportOpen(true)}
        editorRef={editorRef}
        sidebarOpen={sidebarOpen}
        onSidebarOpenChange={setSidebarOpen}
        onHistoryBusyChange={setEditorHistoryBusy}
        onPasteImage={handleEditorImagePaste}
        onCopyPng={handleCopyPng}
        onCopy={async () => {
          await handleCopy();
        }}
        onReplaceDocument={replaceDocumentWithHistory}
        ocrSelection={ocrQuickSelection.selection}
        ocrOptions={ocrQuickSelection.options}
        ocrBusy={ocrQuickSelection.busy}
        onOcrSelectionChange={(selection) =>
          void ocrQuickSelection.handleSelectionChange(selection)
        }
        onQuickOcr={() => void handleQuickOcr()}
        quickOcrCaptureMode={quickOcrCaptureMode}
        onQuickOcrCaptureModeChange={handleQuickOcrCaptureModeChange}
        silentOcrEnabled={silentOcrEnabled}
        onSilentOcrEnabledChange={handleSilentOcrEnabledChange}
        ocrOverlay={
          inlineOcr ? (
            <div
              className={`inline-ocr-progress is-${inlineOcr.status}`}
              role="status"
              aria-live="polite"
            >
              <span className="inline-ocr-progress-icon">
                {inlineOcr.status === "running" ||
                inlineOcr.status === "cancelling" ? (
                  <LoaderCircle size={17} className="is-spinning" />
                ) : inlineOcr.status === "success" ? (
                  <Check size={17} />
                ) : inlineOcr.status === "error" ? (
                  <AlertCircle size={17} />
                ) : (
                  <X size={17} />
                )}
              </span>
              <div>
                <strong>{inlineOcr.message}</strong>
                <span>
                  {ocrQuickSelection.activeOption
                    ? isEn
                      ? ocrQuickSelection.activeOption.labelEn
                      : ocrQuickSelection.activeOption.labelZh
                    : isEn
                      ? inlineOcrModel.labelEn
                      : inlineOcrModel.labelZh}
                  {" · "}
                  {inlineOcr.seconds}
                  {isEn ? "s" : " 秒"}
                </span>
              </div>
              {inlineOcrIsBusy ? (
                <button
                  type="button"
                  className="inline-ocr-cancel"
                  onClick={cancelInlineOcr}
                  disabled={inlineOcr.status === "cancelling"}
                >
                  <X size={13} />
                  {isEn ? "Cancel" : "取消"}
                </button>
              ) : (
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
      <SettingsDialog
        open={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        onCheckForUpdates={() => {
          setSettingsOpen(false);
          void runUpdateCheck(true);
        }}
        onOpenFormulaHotkeys={() => {
          setSettingsOpen(false);
          setFormulaHotkeyManagerOpen(true);
        }}
      />
      <FormulaHotkeyManagerDialog
        open={formulaHotkeyManagerOpen}
        onClose={() => setFormulaHotkeyManagerOpen(false)}
      />
      <HelpManualDialog
        open={helpManualOpen}
        language={language}
        onClose={() => setHelpManualOpen(false)}
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
      <OcrDialog
        open={ocrOpen}
        language={language}
        model={ocrModel}
        onModelChange={handleOcrModelChange}
        onClose={() => setOcrOpen(false)}
        onInsert={(value) => editorRef.current?.insertLatex(value, "ocr")}
        onAppend={(value) => editorRef.current?.appendLatex(value, "ocr")}
        onNotify={setToast}
      />
      <MacOfficeFirstRunPrompt
        open={macOfficeFirstRunOpen}
        language={language}
        mode={macOfficePromptMode}
        powerpointRegistrationRequired={powerpointRegistrationRequired}
        onComplete={finishMacOfficeFirstRun}
      />
      <OnboardingTour
        open={onboardingOpen}
        language={language}
        platform={DESKTOP_PLATFORM}
        onFinish={finishOnboarding}
      />
      <UpdateDialog
        open={updateOpen}
        language={language}
        checking={updateChecking}
        error={updateError}
        result={updateResult}
        checkOnStartup={checkUpdatesOnStartup}
        automaticPrompt={automaticUpdatePrompt}
        onCheckOnStartupChange={setCheckUpdatesOnStartup}
        onRetry={() => void runUpdateCheck(true)}
        onOpenRelease={() => {
          if (!updateResult) return;
          void openReleasePage(updateResult.releaseUrl).catch((error) => {
            setUpdateError(
              error instanceof Error
                ? error.message
                : isEn
                  ? "Unable to open the download page"
                  : "无法打开下载页面",
            );
          });
        }}
        onClose={() => setUpdateOpen(false)}
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
