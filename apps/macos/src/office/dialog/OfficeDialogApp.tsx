import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { listen } from "@tauri-apps/api/event";
import { MathfieldElement } from "mathlive";
import { AlertCircle, Check, LoaderCircle, Redo2, Undo2, X } from "lucide-react";
import { OcrDialog } from "../../components/OcrDialog";
import { EditorWorkspace } from "../../workspace/EditorWorkspace";
import {
  historyManager,
  useHistorySnapshot,
} from "../../history/HistoryManager";
import {
  applyHistoryEntryToEditor,
  documentSnapshotsEquivalent,
  getEditorDocumentSnapshot,
} from "../../history/documentHistory";
import type {
  DocumentSnapshot,
  ReplaceDocumentEntry,
} from "../../history/historyTypes";
import {
  joinFormulaLines,
  useEditorStore,
} from "../../stores/editorStore";
import {
  copyFormulaLines,
  isLatexCodeFormat,
} from "../../clipboard/LatexCopyService";
import type { LatexCodeFormat } from "../../types/formula";
import {
  applyDocumentTheme,
  normalizeSynchronizedTheme,
  readSynchronizedTheme,
  subscribeSynchronizedTheme,
} from "../../themeSync";
import type {
  MathEditorHandle,
  MathEditorInsertionTarget,
} from "../../editor/MathEditor";
import { createUuid } from "../../runtime/browserCompatibility";
import {
  normalizeFormulaChineseFont,
  normalizeFormulaLetterFont,
  readPersistedFormulaFontPreferences,
} from "../../editor/formulaFontPreferences";
import { errorMessage } from "../../runtime/errorMessage";
import {
  readLocalStorage,
  writeLocalStorage,
} from "../../runtime/safeStorage";
import {
  closeCurrentTauriWindow,
  invokeTauri,
  onCurrentTauriWindowCloseRequested,
} from "../shared/tauriTransport";
import { normalizeFormulaEditorDocument } from "../shared/formulaEditorDocument";
import {
  readWorkspacePanelOpen,
  writeWorkspacePanelOpen,
} from "../../workspace/workspacePanelPreferences";
import {
  renderOfficeFormulaArtifacts,
  tryRenderOfficeFormulaDraftArtifacts,
  type OfficeFormulaRenderArtifacts,
} from "../shared/formulaRenderArtifacts";
import {
  cancelMacosOfflineOfficeSession,
  commitMacosOfflineOfficeSession,
  getOfficeSession,
  isMacosOfflineTauriTransport,
  saveOfficeSessionKeepalive,
  type OfficeExportResult,
  type OfficeHost,
  type UpdateOfficeSessionInput,
} from "../api/sessionClient";
import { useOfficeSession } from "./useOfficeSession";
import { messageOfficeParent } from "./dialogMessages";
import { registerOfficeApplyShortcut } from "./officeApplyShortcut";
import {
  OCR_MODELS,
  cancelOcrRecognition,
  fileToOcrRequest,
  getOcrRuntimeStatus,
  listenOcrRecognitionProgress,
  recognizeFormulaImage,
  resolveAvailableOcrModel,
  prewarmOcrModel,
  type OcrModelName,
} from "../../ocr/ocrService";

type InlineOcrStatus =
  | "running"
  | "cancelling"
  | "success"
  | "error"
  | "cancelled";

interface InlineOcrState {
  status: InlineOcrStatus;
  message: string;
  seconds: number;
  model: OcrModelName;
}

const DEFAULT_OCR_MODEL: OcrModelName = "PP-FormulaNet_plus-M";
const EDITOR_PERSISTENCE_STORAGE_KEY = "visualtex-editor";
const OCR_MODEL_STORAGE_KEY = "visualtex.ocr.model";
const OFFICE_WORD_CREATE_NUMBERED_STORAGE_KEY =
  "visualtex.office.word.create.numbered";
const USE_NATIVE_POWERPOINT_COMMIT =
  document
    .querySelector<HTMLMetaElement>(
      'meta[name="visualtex-native-powerpoint-commit"]',
    )
    ?.content.toLowerCase() === "true";

const OFFICE_COMMIT_RESULT_TIMEOUT_MS = 45_000;

function readOfficeWordCreateNumberedPreference(fallback: boolean) {
  const stored = readLocalStorage(OFFICE_WORD_CREATE_NUMBERED_STORAGE_KEY);
  if (stored === "true") return true;
  if (stored === "false") return false;
  return fallback;
}

function writeOfficeWordCreateNumberedPreference(numbered: boolean) {
  writeLocalStorage(
    OFFICE_WORD_CREATE_NUMBERED_STORAGE_KEY,
    numbered ? "true" : "false",
  );
}

function syncOfficeEditorSystemSettings(raw?: string | null) {
  const source = raw === undefined
    ? readLocalStorage(EDITOR_PERSISTENCE_STORAGE_KEY)
    : raw;
  if (!source) return;
  let persisted: Record<string, unknown> | null = null;
  try {
    const envelope = JSON.parse(source) as { state?: unknown };
    if (envelope.state && typeof envelope.state === "object") {
      persisted = envelope.state as Record<string, unknown>;
    }
  } catch {
    return;
  }
  if (!persisted) return;

  const editor = useEditorStore.getState();
  const applyNumber = (
    key: string,
    current: number,
    setter: (value: number) => void,
  ) => {
    const value = persisted?.[key];
    if (typeof value === "number" && Number.isFinite(value) && value !== current) {
      setter(value);
    }
  };
  const applyBoolean = (
    key: string,
    current: boolean,
    setter: (value: boolean) => void,
  ) => {
    const value = persisted?.[key];
    if (typeof value === "boolean" && value !== current) setter(value);
  };

  if (
    (persisted.editorLayout === "standard" || persisted.editorLayout === "classic") &&
    persisted.editorLayout !== editor.editorLayout
  ) {
    editor.setEditorLayout(persisted.editorLayout);
  }
  if (
    (persisted.language === "cn" || persisted.language === "en") &&
    persisted.language !== editor.language
  ) {
    editor.setLanguage(persisted.language);
  }
  applyNumber("zoom", editor.zoom, editor.setZoom);
  applyBoolean("autoPairDelimiters", editor.autoPairDelimiters, editor.setAutoPairDelimiters);
  applyBoolean("showLineNumbers", editor.showLineNumbers, editor.setShowLineNumbers);
  applyBoolean("highlightActiveLine", editor.highlightActiveLine, editor.setHighlightActiveLine);
  applyNumber("formulaInsetLeft", editor.formulaInsetLeft, editor.setFormulaInsetLeft);
  applyNumber("formulaInsetRight", editor.formulaInsetRight, editor.setFormulaInsetRight);
  applyNumber("formulaToolButtonSize", editor.formulaToolButtonSize, editor.setFormulaToolButtonSize);
  applyNumber("formulaToolButtonPadding", editor.formulaToolButtonPadding, editor.setFormulaToolButtonPadding);
  applyNumber("formulaRowVerticalInset", editor.formulaRowVerticalInset, editor.setFormulaRowVerticalInset);
  applyBoolean("personalize", editor.personalize, editor.setPersonalize);
  applyNumber("suggestionCount", editor.suggestionCount, editor.setSuggestionCount);
  applyNumber("classicTileWidth", editor.classicTileWidth, editor.setClassicTileWidth);
  applyNumber("classicDockHeight", editor.classicDockHeight, editor.setClassicDockHeight);
  applyBoolean("keypadMinimizeOnCopy", editor.keypadMinimizeOnCopy, editor.setKeypadMinimizeOnCopy);

  const persistedInputBehavior = persisted.inputBehavior;
  if (persistedInputBehavior && typeof persistedInputBehavior === "object") {
    const inputBehavior = persistedInputBehavior as Record<string, unknown>;
    for (const key of Object.keys(editor.inputBehavior) as Array<keyof typeof editor.inputBehavior>) {
      const value = inputBehavior[key];
      if (typeof value === "boolean" && value !== editor.inputBehavior[key]) {
        editor.setInputBehavior(key, value);
      }
    }
  }
}

function officeExportResultFromArtifacts(
  artifacts: OfficeFormulaRenderArtifacts,
): OfficeExportResult {
  const { svg } = artifacts;
  const wordArtifacts = artifacts.omml;
  return {
    svg: svg.svg,
    svgBase64: svg.base64,
    ...(wordArtifacts
      ? {
          ommlBase64: wordArtifacts.ommlBase64,
          ommlDocxBase64: wordArtifacts.ommlDocxBase64,
        }
      : {}),
    width: svg.width,
    height: svg.height,
    baseline: svg.baseline,
  };
}

function normalizeOfficeFontSizePt(value: unknown, fallback: number) {
  const numeric = typeof value === "number" ? value : Number(value);
  const resolved = Number.isFinite(numeric) ? numeric : fallback;
  return Math.round(Math.min(200, Math.max(5, resolved)) * 2) / 2;
}

const OFFICE_CHINESE_FONT_SIZE_OPTIONS = [
  { name: "初号", fontSizePt: 42 },
  { name: "小初", fontSizePt: 36 },
  { name: "一号", fontSizePt: 26 },
  { name: "小一", fontSizePt: 24 },
  { name: "二号", fontSizePt: 22 },
  { name: "小二", fontSizePt: 18 },
  { name: "三号", fontSizePt: 16 },
  { name: "小三", fontSizePt: 15 },
  { name: "四号", fontSizePt: 14 },
  { name: "小四", fontSizePt: 12 },
  { name: "五号", fontSizePt: 10.5 },
  { name: "小五", fontSizePt: 9 },
  { name: "六号", fontSizePt: 7.5 },
  { name: "小六", fontSizePt: 6.5 },
  { name: "七号", fontSizePt: 5.5 },
  { name: "八号", fontSizePt: 5 },
] as const;

const OFFICE_CHINESE_FONT_SIZE_POINTS = new Set<number>(
  OFFICE_CHINESE_FONT_SIZE_OPTIONS.map((option) => option.fontSizePt),
);

const OFFICE_FONT_SIZE_OPTIONS = [
  ...Array.from({ length: 63 }, (_, index) => 5 + index * 0.5),
  ...Array.from({ length: 18 }, (_, index) => 38 + index * 2),
  80,
  90,
  96,
  100,
  120,
  144,
  160,
  180,
  200,
].filter((fontSizePt) => !OFFICE_CHINESE_FONT_SIZE_POINTS.has(fontSizePt));

function officePointFontSizeOptions(currentFontSizePt: number) {
  const current = normalizeOfficeFontSizePt(currentFontSizePt, currentFontSizePt);
  if (OFFICE_CHINESE_FONT_SIZE_POINTS.has(current)) {
    return OFFICE_FONT_SIZE_OPTIONS;
  }
  return OFFICE_FONT_SIZE_OPTIONS.includes(current)
    ? OFFICE_FONT_SIZE_OPTIONS
    : [...OFFICE_FONT_SIZE_OPTIONS, current].sort((left, right) => left - right);
}

function delay(milliseconds: number) {
  return new Promise<void>((resolve) => window.setTimeout(resolve, milliseconds));
}

async function waitForOfficeCommitResult(
  sessionId: string,
  host: OfficeHost,
) {
  const hostLabel = host === "word" ? "Word" : "PowerPoint";
  const deadline = Date.now() + OFFICE_COMMIT_RESULT_TIMEOUT_MS;
  while (Date.now() < deadline) {
    const current = await getOfficeSession(sessionId);
    if (current.status === "completed") return;
    if (current.status === "failed") {
      throw new Error(current.error || `${hostLabel} 公式写入失败。`);
    }
    if (current.status === "cancelled" || current.explicitCancel) {
      throw new Error(`${hostLabel} 公式写入已取消。`);
    }
    await delay(100);
  }
  throw new Error(`等待 ${hostLabel} 确认写入超时，请重试。`);
}

function documentFingerprint(
  title: string,
  lines: Array<{ id: string; latex: string }>,
  codeFormat: string,
  displayMode: "inline" | "block",
  numbered: boolean,
  fontSizePt: number,
  formulaLetterFont: string,
  formulaChineseFont: string,
) {
  return JSON.stringify({
    title,
    lines: lines.map((line) => line.latex),
    codeFormat,
    displayMode,
    numbered,
    fontSizePt: normalizeOfficeFontSizePt(fontSizePt, fontSizePt),
    formulaLetterFont,
    formulaChineseFont,
  });
}

export function OfficeDialogApp() {
  const tauriResidentEditor = isMacosOfflineTauriTransport();
  const editorRef = useRef<MathEditorHandle>(null);
  const loadedSessionIdRef = useRef("");
  const skipAutosaveForSessionRef = useRef("");
  const lastSavedFingerprintRef = useRef("");
  const readyMessageSentRef = useRef(false);
  const finalizingRef = useRef(false);
  const allowNativeCloseRef = useRef(false);
  const nativeCloseRequestInFlightRef = useRef(false);
  const exportRunIdRef = useRef(0);
  const activeSessionKeyRef = useRef("");
  const readyReportedSessionKeyRef = useRef("");
  const silentCommitSessionKeyRef = useRef("");
  const prewarmReportedRef = useRef(false);
  const latestCompleteExportRef = useRef<{
    fingerprint: string;
    exportResult: OfficeExportResult;
  } | null>(null);
  const completeExportInFlightRef = useRef<{
    fingerprint: string;
    promise: Promise<OfficeExportResult | null>;
  } | null>(null);
  const [sidebarOpen, setSidebarOpenState] = useState(() =>
    readWorkspacePanelOpen("office-edit", "tiles"),
  );
  const setSidebarOpen = useCallback((open: boolean) => {
    setSidebarOpenState(open);
    writeWorkspacePanelOpen("office-edit", "tiles", open);
  }, []);
  const [historyBusy, setHistoryBusy] = useState(false);
  const [autoCommitOnClose, setAutoCommitOnClose] = useState(true);
  const [displayMode, setDisplayMode] = useState<"inline" | "block">("inline");
  const [numbered, setNumbered] = useState(false);
  const [officeFontSizePt, setOfficeFontSizePt] = useState(14);
  const [toast, setToast] = useState("");
  const [ocrOpen, setOcrOpen] = useState(false);
  const [ocrModel, setOcrModel] = useState<OcrModelName>(() => {
    const stored = readLocalStorage(OCR_MODEL_STORAGE_KEY);
    return OCR_MODELS.some((item) => item.id === stored)
      ? (stored as OcrModelName)
      : DEFAULT_OCR_MODEL;
  });
  const [inlineOcr, setInlineOcr] = useState<InlineOcrState | null>(null);
  const [hydratedSessionKey, setHydratedSessionKey] = useState("");
  const [hydratedPerformanceMs, setHydratedPerformanceMs] = useState(0);
  const inlineOcrBusyRef = useRef(false);
  const inlineOcrCancelRequestedRef = useRef(false);
  const inlineOcrRunIdRef = useRef(0);
  const inlineOcrClearTimerRef = useRef<number | null>(null);
  const ocrPrewarmStartedRef = useRef(false);
  const {
    sessionId,
    generation,
    session,
    loading,
    error,
    save,
    activationPerformanceMs,
    sessionLoadedPerformanceMs,
  } = useOfficeSession();
  const sessionKey = sessionId ? `${generation}:${sessionId}` : "";
  const sessionHydrated = Boolean(
    session && sessionKey && hydratedSessionKey === sessionKey,
  );
  activeSessionKeyRef.current = sessionKey;

  useEffect(() => {
    if (!isMacosOfflineTauriTransport() || prewarmReportedRef.current) {
      return;
    }

    let disposed = false;
    let mountProbeTimer = 0;
    let stableMountedChecks = 0;
    let lastDiagnosticSignature = "";
    const startedAt =
      typeof performance === "undefined" ? Date.now() : performance.now();
    const deadline = startedAt + 30_000;
    const reportDiagnostic = (
      stage: "effect-start" | "waiting" | "ready" | "timeout",
      editorReady: boolean,
      mathfieldHosts: number,
      now: number,
    ) => {
      const signature = `${stage}:${editorReady ? 1 : 0}:${mathfieldHosts}`;
      if (signature === lastDiagnosticSignature) return;
      lastDiagnosticSignature = signature;
      void invokeTauri<void>(
        "report_macos_offline_office_editor_prewarm_diagnostic",
        {
          input: {
            stage,
            editorReady,
            mathfieldHosts,
            elapsedMs: Math.max(0, now - startedAt),
          },
        },
      ).catch((reason) => {
        console.error("Unable to report Office editor prewarm diagnostics", reason);
      });
    };

    reportDiagnostic("effect-start", false, 0, startedAt);
    const reportWhenEditorMounted = () => {
      if (disposed || prewarmReportedRef.current) return;
      const now =
        typeof performance === "undefined" ? Date.now() : performance.now();
      const editorReady = Boolean(editorRef.current);
      const mathfieldHosts = document.querySelectorAll(
        ".formula-line .mathfield-host",
      ).length;
      reportDiagnostic("waiting", editorReady, mathfieldHosts, now);
      if (editorReady && mathfieldHosts >= 1) {
        stableMountedChecks += 1;
      } else {
        stableMountedChecks = 0;
      }
      if (stableMountedChecks < 2 && now < deadline) {
        mountProbeTimer = window.setTimeout(reportWhenEditorMounted, 16);
        return;
      }
      if (stableMountedChecks < 2) {
        reportDiagnostic("timeout", editorReady, mathfieldHosts, now);
        console.error("Unable to prewarm the resident Office MathLive editor");
        return;
      }

      reportDiagnostic("ready", editorReady, mathfieldHosts, now);
      prewarmReportedRef.current = true;
      void invokeTauri<void>(
        "report_macos_offline_office_editor_prewarmed",
      ).catch((reason) => {
        prewarmReportedRef.current = false;
        console.error("Unable to report Office editor prewarming", reason);
      });
    };

    mountProbeTimer = window.setTimeout(reportWhenEditorMounted, 0);
    return () => {
      disposed = true;
      window.clearTimeout(mountProbeTimer);
    };
  }, []);

  useEffect(() => {
    const sync = (raw?: string | null) => syncOfficeEditorSystemSettings(raw);
    sync();
    const handleStorage = (event: StorageEvent) => {
      if (event.key === EDITOR_PERSISTENCE_STORAGE_KEY) sync(event.newValue);
    };
    const handleFocus = () => sync();
    window.addEventListener("storage", handleStorage);
    window.addEventListener("focus", handleFocus);
    return () => {
      window.removeEventListener("storage", handleStorage);
      window.removeEventListener("focus", handleFocus);
    };
  }, []);

  useEffect(() => {
    loadedSessionIdRef.current = "";
    skipAutosaveForSessionRef.current = "";
    lastSavedFingerprintRef.current = "";
    readyMessageSentRef.current = false;
    readyReportedSessionKeyRef.current = "";
    finalizingRef.current = false;
    allowNativeCloseRef.current = false;
    nativeCloseRequestInFlightRef.current = false;
    exportRunIdRef.current += 1;
    latestCompleteExportRef.current = null;
    completeExportInFlightRef.current = null;
    historyManager.clear();
    setHydratedSessionKey("");
    setHydratedPerformanceMs(0);
    setToast("");
    setOcrOpen(false);
    setInlineOcr(null);
    inlineOcrBusyRef.current = false;
    inlineOcrCancelRequestedRef.current = false;
    inlineOcrRunIdRef.current += 1;
    useEditorStore.getState().replaceDocumentState({
      title: "",
      lines: [{ id: createUuid(), latex: "" }],
      activeLineId: null,
      formulaAlignment: useEditorStore.getState().formulaAlignment,
      selectionByLineId: {},
    });
    useEditorStore.getState().setLatexCodeFormat("raw");
    setAutoCommitOnClose(true);
    setDisplayMode("inline");
    setNumbered(false);
    setOfficeFontSizePt(14);
  }, [sessionKey]);

  useEffect(() => {
    if (!sessionId && isMacosOfflineTauriTransport()) {
      document.title = "VisualTeX Office Formula — 待命";
      return;
    }
    if (loading || (session && !sessionHydrated)) {
      document.title = "VisualTeX Office Formula — 正在加载";
      return;
    }
    if (error || !session) {
      document.title = `VisualTeX Office Formula — 加载失败${error ? `：${error.slice(0, 80)}` : ""}`;
      return;
    }
    document.title = `VisualTeX Office Formula — ${session.host === "word" ? "Word" : "PowerPoint"} 已就绪`;
  }, [loading, error, session?.id, session?.host, sessionHydrated, sessionId]);

  const title = useEditorStore((state) => state.title);
  const lines = useEditorStore((state) => state.lines);
  const activeLineId = useEditorStore((state) => state.activeLineId);
  const language = useEditorStore((state) => state.language);
  const theme = useEditorStore((state) => state.theme);
  const setTheme = useEditorStore((state) => state.setTheme);
  const latexCodeFormat = useEditorStore((state) => state.latexCodeFormat);
  const formulaLetterFont = useEditorStore((state) => state.formulaLetterFont);
  const formulaChineseFont = useEditorStore((state) => state.formulaChineseFont);
  const powerPointDefaultFontSizePt = useEditorStore(
    (state) => state.powerPointDefaultFontSizePt,
  );
  const addHistory = useEditorStore((state) => state.addHistory);
  const historyState = useHistorySnapshot();
  const isEn = language === "en";
  const latex = joinFormulaLines(lines);

  useEffect(() => {
    const applyTheme = (nextThemeValue: unknown) => {
      const nextTheme = normalizeSynchronizedTheme(nextThemeValue);
      applyDocumentTheme(nextTheme);
      if (useEditorStore.getState().theme !== nextTheme) {
        setTheme(nextTheme);
      }
    };

    applyTheme(readSynchronizedTheme());
    const unsubscribeBrowser = subscribeSynchronizedTheme(applyTheme);
    let disposed = false;
    let unsubscribeTauri: (() => void) | undefined;
    if (isMacosOfflineTauriTransport()) {
      void listen<unknown>("visualtex-theme-changed", (event) => {
        applyTheme(event.payload);
      })
        .then((unsubscribe) => {
          if (disposed) unsubscribe();
          else unsubscribeTauri = unsubscribe;
        })
        .catch(() => undefined);
    }

    return () => {
      disposed = true;
      unsubscribeBrowser();
      unsubscribeTauri?.();
    };
  }, [setTheme]);

  useEffect(() => {
    applyDocumentTheme(theme);
  }, [theme]);
  const selectedOcrModel =
    OCR_MODELS.find((item) => item.id === ocrModel) ?? OCR_MODELS[1];
  const inlineOcrModel =
    OCR_MODELS.find((item) => item.id === inlineOcr?.model) ?? selectedOcrModel;
  const inlineOcrIsBusy =
    inlineOcr?.status === "running" || inlineOcr?.status === "cancelling";

  const resolvedSessionFormulaFonts = useMemo(() => {
    const current = useEditorStore.getState();
    const persistedGlobal = readPersistedFormulaFontPreferences();
    const globalLetterFont =
      persistedGlobal.formulaLetterFont ?? current.formulaLetterFont;
    const globalChineseFont =
      persistedGlobal.formulaChineseFont ?? current.formulaChineseFont;
    return {
      formulaLetterFont: normalizeFormulaLetterFont(
        session?.originalMetadata?.formulaLetterFont ??
          (tauriResidentEditor ? undefined : session?.formulaLetterFont) ??
          globalLetterFont,
      ),
      formulaChineseFont: normalizeFormulaChineseFont(
        session?.originalMetadata?.formulaChineseFont ??
          (tauriResidentEditor ? undefined : session?.formulaChineseFont) ??
          globalChineseFont,
      ),
    };
  }, [
    sessionKey,
    session?.formulaChineseFont,
    session?.formulaLetterFont,
    session?.originalMetadata?.formulaChineseFont,
    session?.originalMetadata?.formulaLetterFont,
    tauriResidentEditor,
  ]);

  const editableSessionDocument = useMemo(
    () =>
      session
        ? normalizeFormulaEditorDocument(session.lines, session.codeFormat)
        : null,
    [session],
  );
  const editableOriginalDocument = useMemo(() => {
    if (!session) return null;
    return normalizeFormulaEditorDocument(
      session.originalMetadata?.lines ?? session.lines,
      session.originalMetadata?.codeFormat ?? session.codeFormat,
    );
  }, [session]);

  const originalFingerprint = useMemo(() => {
    if (!session || !editableOriginalDocument) return "";
    const fallbackFontSizePt = session.host === "word" ? 11 : 18;
    const originalFontSizePt =
      session.host === "powerpoint" &&
      session.mode === "create" &&
      session.status === "created" &&
      !session.dirty
        ? powerPointDefaultFontSizePt
        : normalizeOfficeFontSizePt(
            session.originalMetadata?.fontSizePt ?? session.fontSizePt,
            fallbackFontSizePt,
          );
    return documentFingerprint(
      session.originalMetadata?.title ?? session.title,
      editableOriginalDocument.lines,
      editableOriginalDocument.codeFormat,
      session.originalMetadata?.displayMode ?? session.displayMode,
      session.originalMetadata?.numbered ?? session.numbered ?? false,
      originalFontSizePt,
      resolvedSessionFormulaFonts.formulaLetterFont,
      resolvedSessionFormulaFonts.formulaChineseFont,
    );
  }, [
    editableOriginalDocument,
    powerPointDefaultFontSizePt,
    resolvedSessionFormulaFonts.formulaChineseFont,
    resolvedSessionFormulaFonts.formulaLetterFont,
    session,
  ]);

  const currentFingerprint = useMemo(
    () =>
      documentFingerprint(
        title,
        lines,
        latexCodeFormat,
        displayMode,
        numbered,
        officeFontSizePt,
        formulaLetterFont,
        formulaChineseFont,
      ),
    [
      title,
      lines,
      latexCodeFormat,
      displayMode,
      numbered,
      officeFontSizePt,
      formulaLetterFont,
      formulaChineseFont,
    ],
  );
  const dirty = Boolean(session) && currentFingerprint !== originalFingerprint;

  useEffect(() => {
    if (
      !session ||
      !editableSessionDocument ||
      loadedSessionIdRef.current === sessionKey
    ) {
      return;
    }
    loadedSessionIdRef.current = sessionKey;
    skipAutosaveForSessionRef.current = session.id;
    const nextLines = editableSessionDocument.lines.length
      ? editableSessionDocument.lines
      : [{ id: createUuid(), latex: "" }];
    useEditorStore.setState({
      formulaLetterFont: resolvedSessionFormulaFonts.formulaLetterFont,
      formulaChineseFont: resolvedSessionFormulaFonts.formulaChineseFont,
    });
    useEditorStore.getState().replaceDocumentState({
      title: session.title || (isEn ? "Office Formula" : "Office 公式"),
      lines: nextLines,
      activeLineId:
        session.activeLineId &&
        nextLines.some((line) => line.id === session.activeLineId)
          ? session.activeLineId
          : nextLines[0]?.id ?? null,
      formulaAlignment: useEditorStore.getState().formulaAlignment,
      selectionByLineId: {},
    });
    if (isLatexCodeFormat(editableSessionDocument.codeFormat)) {
      useEditorStore
        .getState()
        .setLatexCodeFormat(editableSessionDocument.codeFormat as LatexCodeFormat);
    }
    setAutoCommitOnClose(session.autoCommitOnClose);
    setDisplayMode(session.displayMode);
    const sessionNumbered =
      session.displayMode === "block" && Boolean(session.numbered);
    const loadedNumbered =
      session.host === "word" &&
      session.mode === "create" &&
      session.displayMode === "block"
        ? readOfficeWordCreateNumberedPreference(sessionNumbered)
        : sessionNumbered;
    setNumbered(loadedNumbered);
    if (
      session.host === "word" &&
      session.mode === "create" &&
      session.displayMode === "block"
    ) {
      writeOfficeWordCreateNumberedPreference(loadedNumbered);
    }
    const loadedFontSizePt =
      session.host === "powerpoint" &&
      session.mode === "create" &&
      session.status === "created" &&
      !session.dirty
        ? powerPointDefaultFontSizePt
        : normalizeOfficeFontSizePt(
            session.fontSizePt ?? session.originalMetadata?.fontSizePt,
            session.host === "word" ? 11 : 18,
          );
    setOfficeFontSizePt(loadedFontSizePt);
    const loadedFingerprint = documentFingerprint(
      session.title,
      nextLines,
      editableSessionDocument.codeFormat,
      session.displayMode,
      session.displayMode === "block" && Boolean(session.numbered),
      loadedFontSizePt,
      resolvedSessionFormulaFonts.formulaLetterFont,
      resolvedSessionFormulaFonts.formulaChineseFont,
    );
    lastSavedFingerprintRef.current = loadedFingerprint;
    latestCompleteExportRef.current = session.exportResult?.pngBase64
      ? { fingerprint: loadedFingerprint, exportResult: session.exportResult }
      : null;
    const hydratedAt =
      typeof performance === "undefined" ? Date.now() : performance.now();
    setHydratedPerformanceMs(hydratedAt);
    setHydratedSessionKey(sessionKey);
  }, [
    editableSessionDocument,
    session?.id,
    sessionKey,
    isEn,
    powerPointDefaultFontSizePt,
    resolvedSessionFormulaFonts.formulaChineseFont,
    resolvedSessionFormulaFonts.formulaLetterFont,
  ]);

  useEffect(() => {
    if (
      !isMacosOfflineTauriTransport() ||
      !session ||
      !editableSessionDocument ||
      !sessionHydrated ||
      !sessionKey ||
      generation <= 0 ||
      readyReportedSessionKeyRef.current === sessionKey
    ) {
      return;
    }

    let disposed = false;
    let readinessCheck = 0;
    let readinessTimer = 0;
    let editorMountedMs = 0;
    const origin =
      activationPerformanceMs ||
      sessionLoadedPerformanceMs ||
      hydratedPerformanceMs;
    const contentReadyDeadlineMs = origin + 5_000;
    const hydrateMs = Math.max(0, hydratedPerformanceMs - origin);
    const expectedLineIds = editableSessionDocument.lines.map((line) => line.id);
    const expectedLineLatex = editableSessionDocument.lines.map((line) => line.latex);

    const inspectContent = () => {
      if (disposed || activeSessionKeyRef.current !== sessionKey) return;
      readinessCheck += 1;
      const now =
        typeof performance === "undefined" ? Date.now() : performance.now();
      if (readinessCheck === 1) {
        editorMountedMs = Math.max(hydrateMs, now - origin);
      }
      const mountedLineIds = new Set(
        Array.from(
          document.querySelectorAll<HTMLElement>(
            ".formula-line[data-line-id]",
          ),
        ).map((element) => element.dataset.lineId ?? ""),
      );
      const mathfields = Array.from(
        document.querySelectorAll<MathfieldElement>(".formula-line math-field"),
      );
      const lineIdsMounted = expectedLineIds.every((lineId) =>
        mountedLineIds.has(lineId),
      );
      const formulaContentMounted =
        mathfields.length >= expectedLineLatex.length &&
        expectedLineLatex.every(
          (latex, index) => mathfields[index]?.value === latex,
        );
      const contentMounted =
        Boolean(editorRef.current) &&
        (lineIdsMounted || formulaContentMounted) &&
        mathfields.length >= expectedLineLatex.length;
      // The resident WebView is intentionally behind Word/PowerPoint while it
      // hydrates. macOS can heavily throttle requestAnimationFrame for that
      // occluded window even when WebKit background throttling is disabled.
      // Use ordinary tasks for readiness polling so a fully mounted editor can
      // cross the application foreground boundary immediately.
      if (
        readinessCheck < 2 ||
        (!contentMounted && now < contentReadyDeadlineMs)
      ) {
        readinessTimer = window.setTimeout(inspectContent, 0);
        return;
      }
      if (!contentMounted) {
        readyReportedSessionKeyRef.current = "";
        setToast(
          isEn
            ? "The formula editor did not finish mounting in time."
            : "公式编辑器未能及时完成挂载。",
        );
        return;
      }

      // A resident MathLive field can keep geometry from the previous Office
      // Session because its React slot intentionally survives window parking.
      // Re-measure every live line before the native window presents it.
      editorRef.current?.refreshLayout();

      // refreshLayout() remounts resident MathLive fields synchronously. Give
      // the custom elements one normal task turn to reconnect, then ask AppKit
      // to present the dedicated editor. Once the app is foreground, its normal
      // animation-frame paint/focus path is no longer subject to occlusion.
      document.body.style.opacity = "1";
      readinessTimer = window.setTimeout(() => {
        if (disposed || activeSessionKeyRef.current !== sessionKey) return;
        const contentReadyAt =
          typeof performance === "undefined" ? Date.now() : performance.now();
        const contentReadyMs = Math.max(
          editorMountedMs,
          contentReadyAt - origin,
        );
        readyReportedSessionKeyRef.current = sessionKey;
        void invokeTauri<void>(
          "report_macos_offline_office_editor_ready",
          {
            input: {
              sessionId: session.id,
              generation,
              frontendEpochMs: Date.now(),
              hydrateMs,
              editorMountedMs,
              contentReadyMs,
            },
          },
        )
          .then(() => {
            if (activeSessionKeyRef.current !== sessionKey) return;
            // A formula opened from Office is an editing action. Focus only
            // after AppKit has made the resident window visible and key; an
            // earlier MathLive focus can race its first connected frame and
            // must never block the ready report.
            window.requestAnimationFrame(() => editorRef.current?.focus());
          })
          .catch((reason) => {
            if (activeSessionKeyRef.current !== sessionKey) return;
            document.body.style.opacity = "0";
            readyReportedSessionKeyRef.current = "";
            setToast(
              errorMessage(
                reason,
                isEn
                  ? "Unable to reveal the Office formula editor"
                  : "无法显示 Office 公式编辑器",
              ),
            );
          });
      }, 0);
    };

    readinessTimer = window.setTimeout(inspectContent, 0);
    return () => {
      disposed = true;
      window.clearTimeout(readinessTimer);
    };
  }, [
    activationPerformanceMs,
    editableSessionDocument,
    generation,
    hydratedPerformanceMs,
    isEn,
    session,
    sessionHydrated,
    sessionKey,
    sessionLoadedPerformanceMs,
  ]);

  const captureSnapshot = useCallback(
    (): DocumentSnapshot =>
      getEditorDocumentSnapshot(editorRef.current?.getSelectionMap() ?? {}),
    [],
  );

  const restoreSnapshotFocus = useCallback((snapshot: DocumentSnapshot) => {
    const lineId = snapshot.activeLineId;
    if (!lineId) return;
    const line = snapshot.lines.find((item) => item.id === lineId);
    if (!line) return;
    void editorRef.current?.restoreSelection(
      lineId,
      line.latex,
      snapshot.selectionByLineId[lineId] ?? null,
    );
  }, []);

  const replaceDocumentWithHistory = useCallback(
    (
      after: DocumentSnapshot,
      source: ReplaceDocumentEntry["source"],
    ) => {
      if (source !== "source-apply") historyManager.commitPendingTransaction();
      const before = captureSnapshot();
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
    },
    [captureSnapshot, restoreSnapshotFocus],
  );

  useEffect(() => {
    historyManager.configure({
      getDocumentSnapshot: captureSnapshot,
      applyEntry: async (entry, direction) => {
        const target = applyHistoryEntryToEditor(entry, direction);
        if (!target) return;
        await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
        await editorRef.current?.restoreSelection(
          target.lineId,
          target.latex,
          target.selection,
        );
      },
    });
    return () => historyManager.configure(null);
  }, [captureSnapshot]);

  const generateSvgExportResult = useCallback((): OfficeExportResult | null => {
    if (!latex.trim()) return null;
    return officeExportResultFromArtifacts(
      renderOfficeFormulaArtifacts({
        lines,
        codeFormat: latexCodeFormat,
        displayMode,
        host: session?.host,
        includeWordOmml: session?.host === "word",
        formulaLetterFont,
        formulaChineseFont,
      }),
    );
  }, [
    latex,
    displayMode,
    lines,
    latexCodeFormat,
    session?.host,
    formulaLetterFont,
    formulaChineseFont,
  ]);

  const generateDraftExportResult = useCallback((): OfficeExportResult | null => {
    if (!latex.trim()) return null;
    const artifacts = tryRenderOfficeFormulaDraftArtifacts({
      lines,
      codeFormat: latexCodeFormat,
      displayMode,
      host: session?.host,
      includeWordOmml: session?.host === "word",
      formulaLetterFont,
      formulaChineseFont,
    });
    return artifacts ? officeExportResultFromArtifacts(artifacts) : null;
  }, [
    latex,
    displayMode,
    lines,
    latexCodeFormat,
    session?.host,
    formulaLetterFont,
    formulaChineseFont,
  ]);

  const generateExportResult = useCallback(async (
    preparedBase?: OfficeExportResult | null,
  ): Promise<OfficeExportResult | null> => {
    const base =
      preparedBase === undefined ? generateSvgExportResult() : preparedBase;
    if (!base) return null;
    let pngBase64: string | undefined;
    let inkTopRatio: number | undefined;
    let inkBottomRatio: number | undefined;
    let inkCenterYRatio: number | undefined;
    try {
      const { svgToPng } = await import("../../export/svgToPng");
      const png = await svgToPng(
        {
          base64: base.svgBase64,
          width: base.width,
          height: base.height,
        },
        { scale: 2, background: "transparent" },
      );
      pngBase64 = png.base64;
      inkTopRatio = png.inkTopRatio;
      inkBottomRatio = png.inkBottomRatio;
      inkCenterYRatio = png.inkCenterYRatio;
    } catch {
      // SVG remains a complete Office fallback when PNG rasterization fails.
    }
    return {
      ...base,
      pngBase64,
      inkTopRatio,
      inkBottomRatio,
      inkCenterYRatio,
    };
  }, [generateSvgExportResult]);

  const getCompleteExportResult = useCallback(
    (
      fingerprint: string,
      preparedBase?: OfficeExportResult | null,
    ): Promise<OfficeExportResult | null> => {
      const cached = latestCompleteExportRef.current;
      if (
        cached?.fingerprint === fingerprint &&
        cached.exportResult.pngBase64
      ) {
        return Promise.resolve(cached.exportResult);
      }
      const inFlight = completeExportInFlightRef.current;
      if (inFlight?.fingerprint === fingerprint) return inFlight.promise;

      const promise = generateExportResult(preparedBase)
        .then((result) => {
          if (result?.pngBase64) {
            latestCompleteExportRef.current = {
              fingerprint,
              exportResult: result,
            };
          }
          return result;
        })
        .finally(() => {
          if (completeExportInFlightRef.current?.promise === promise) {
            completeExportInFlightRef.current = null;
          }
        });
      completeExportInFlightRef.current = { fingerprint, promise };
      return promise;
    },
    [generateExportResult],
  );

  useEffect(() => {
    if (!session || !sessionId || !sessionHydrated || finalizingRef.current) {
      return;
    }
    if (skipAutosaveForSessionRef.current === sessionId) {
      skipAutosaveForSessionRef.current = "";
      return;
    }
    if (
      lastSavedFingerprintRef.current === currentFingerprint &&
      session.autoCommitOnClose === autoCommitOnClose
    ) {
      return;
    }

    const runId = ++exportRunIdRef.current;
    // Autosave is a source-persistence path, not an explicit Office write.
    // MathLive templates temporarily contain placeholders, unclosed groups or
    // a partially typed command. Save those drafts without artifacts and wait
    // for a later valid edit instead of repeatedly surfacing strict MathJax/
    // OMML errors while the user is still typing.
    const exportResult = generateDraftExportResult();
    const draftUpdate = {
      title,
      lines,
      activeLineId,
      codeFormat: latexCodeFormat,
      displayMode,
      numbered: displayMode === "block" && numbered,
      fontSizePt: officeFontSizePt,
      formulaLetterFont,
      formulaChineseFont,
      dirty,
      status: "editing",
      autoCommitOnClose,
      exportResult,
      exportWidth: exportResult?.width ?? 0,
      exportHeight: exportResult?.height ?? 0,
      error: null,
    } as const;
    void save(draftUpdate)
      .then((saved) => {
        if (saved && runId === exportRunIdRef.current) {
          lastSavedFingerprintRef.current = currentFingerprint;
        }
      })
      .catch((reason) => {
        setToast(
          errorMessage(
            reason,
            isEn ? "Unable to save the Office formula" : "无法保存 Office 公式",
          ),
        );
      });
    // Windows OLE inserts a PNG file. Keep rasterization off the critical
    // keystroke-save path, but persist the full export as soon as it is
    // ready so the title-bar close button has a committable final draft.
    if (
      exportResult &&
      !(session.host === "powerpoint" && USE_NATIVE_POWERPOINT_COMMIT)
    ) {
      void getCompleteExportResult(currentFingerprint, exportResult)
        .then((completeExport) => {
          if (
            !completeExport?.pngBase64 ||
            runId !== exportRunIdRef.current ||
            finalizingRef.current
          ) {
            return;
          }
          return save({
            ...draftUpdate,
            exportResult: completeExport,
            exportWidth: completeExport.width,
            exportHeight: completeExport.height,
          }).then((saved) => {
            if (saved && runId === exportRunIdRef.current) {
              lastSavedFingerprintRef.current = currentFingerprint;
            }
          });
        })
        .catch(() => {
          // The immediate SVG save is still recoverable. The explicit
          // insert/update path reports rasterization errors to the user.
        });
    } else {
      latestCompleteExportRef.current = null;
    }
  }, [
    sessionId,
    sessionHydrated,
    session?.id,
    session?.autoCommitOnClose,
    currentFingerprint,
    title,
    lines,
    activeLineId,
    latexCodeFormat,
    displayMode,
    numbered,
    officeFontSizePt,
    formulaLetterFont,
    formulaChineseFont,
    dirty,
    autoCommitOnClose,
    save,
    isEn,
    generateDraftExportResult,
    getCompleteExportResult,
  ]);

  useEffect(() => {
    if (!sessionId || !sessionHydrated) return;
    const finalDraftUpdate = (status: "editing" | "committing") => {
      const cached = latestCompleteExportRef.current;
      const exportResult =
        cached?.fingerprint === currentFingerprint
          ? cached.exportResult
          : status === "editing"
            ? generateDraftExportResult()
            : generateSvgExportResult();
      return {
        title,
        lines,
        activeLineId,
        codeFormat: latexCodeFormat,
        displayMode,
        numbered: displayMode === "block" && numbered,
        fontSizePt: officeFontSizePt,
        formulaLetterFont,
        formulaChineseFont,
        dirty,
        status,
        autoCommitOnClose,
        exportResult,
        exportWidth: exportResult?.width ?? 0,
        exportHeight: exportResult?.height ?? 0,
        error: null,
      } as const;
    };
    const persistFinalDraft = () => {
      if (finalizingRef.current) return;
      try {
        void saveOfficeSessionKeepalive(
          sessionId,
          finalDraftUpdate("editing"),
        ).catch(() => undefined);
      } catch {
        // The regular save path reports export errors while the page is open.
      }
    };
    const commitFinalDraft = () => {
      const cached = latestCompleteExportRef.current;
      const nativePowerPoint =
        session?.host === "powerpoint" && USE_NATIVE_POWERPOINT_COMMIT;
      if (
        finalizingRef.current ||
        !autoCommitOnClose ||
        !latex.trim() ||
        (!nativePowerPoint &&
          (cached?.fingerprint !== currentFingerprint ||
            !cached.exportResult.pngBase64))
      ) {
        persistFinalDraft();
        return;
      }
      try {
        finalizingRef.current = true;
        const update = finalDraftUpdate("committing");
        // The hidden Office command page owns every host mutation. The dialog
        // only persists a complete committing Session, including for native
        // PowerPoint. Directly mutating PowerPoint from this child window used
        // to bypass the adapter's durable name/tag decoration and produced
        // uneditable generic `Graphic N` shapes.
        void saveOfficeSessionKeepalive(sessionId, update).catch(
          () => undefined,
        );
      } catch {
        // Closing a dialog is best-effort; the explicit insert button reports errors.
      }
    };
    const persistWhenHidden = () => {
      if (document.visibilityState === "hidden") persistFinalDraft();
    };
    window.addEventListener("pagehide", commitFinalDraft);
    window.addEventListener("beforeunload", commitFinalDraft);
    document.addEventListener("visibilitychange", persistWhenHidden);
    return () => {
      window.removeEventListener("pagehide", commitFinalDraft);
      window.removeEventListener("beforeunload", commitFinalDraft);
      document.removeEventListener("visibilitychange", persistWhenHidden);
    };
  }, [
    sessionId,
    sessionHydrated,
    session?.host,
    title,
    lines,
    activeLineId,
    latexCodeFormat,
    displayMode,
    numbered,
    officeFontSizePt,
    formulaLetterFont,
    formulaChineseFont,
    dirty,
    autoCommitOnClose,
    currentFingerprint,
    generateDraftExportResult,
    generateSvgExportResult,
    latex,
  ]);

  useEffect(() => {
    if (
      !session ||
      !sessionHydrated ||
      readyMessageSentRef.current ||
      isMacosOfflineTauriTransport()
    ) {
      return;
    }
    readyMessageSentRef.current = true;
    messageOfficeParent({ type: "visualtex-ready", sessionId: session.id });
  }, [session?.id, sessionHydrated]);

  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(""), 2200);
    return () => window.clearTimeout(timer);
  }, [toast]);

  useEffect(() => {
    if (!inlineOcrIsBusy) return;
    const startedAt = Date.now();
    const timer = window.setInterval(() => {
      setInlineOcr((current) =>
        current
          ? {
              ...current,
              seconds: Math.max(0, Math.floor((Date.now() - startedAt) / 1000)),
            }
          : current,
      );
    }, 250);
    return () => window.clearInterval(timer);
  }, [inlineOcrIsBusy]);

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
    if (!sessionHydrated) return;
    let cancelled = false;
    const delay = ocrPrewarmStartedRef.current ? 250 : 500;
    const timer = window.setTimeout(() => {
      ocrPrewarmStartedRef.current = true;
      void getOcrRuntimeStatus()
        .then((runtime) => {
          if (cancelled || !runtime.installed) return;
          const availableModel = resolveAvailableOcrModel(runtime, ocrModel);
          return prewarmOcrModel(availableModel);
        })
        .catch(() => undefined);
    }, delay);

    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [ocrModel, sessionHydrated]);

  const handleOcrModelChange = (nextModel: OcrModelName) => {
    if (inlineOcrBusyRef.current || nextModel === ocrModel) return;
    setOcrModel(nextModel);
    writeLocalStorage(OCR_MODEL_STORAGE_KEY, nextModel);
  };

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
      // A worker that already exited is equivalent to successful cancellation.
    }
  };

  const handleEditorImagePaste = async (
    file: File,
    target: MathEditorInsertionTarget,
  ) => {
    if (inlineOcrBusyRef.current) {
      setToast(
        isEn
          ? "Another pasted image is being recognized"
          : "已有一张粘贴图片正在识别",
      );
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
      message: isEn
        ? "Checking the local OCR runtime…"
        : "正在检查本地 OCR 环境…",
      seconds: 0,
      model: ocrModel,
    });

    let unlisten: (() => void) | undefined;
    try {
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
      const availableOcrModel = ocrModel;

      unlisten = await listenOcrRecognitionProgress((progress) => {
        if (
          inlineOcrRunIdRef.current !== runId ||
          progress.model !== ocrModel
        ) {
          return;
        }
        setInlineOcr((current) =>
          current ? { ...current, message: progress.message } : current,
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

      const recognizedLatex = result.formulas
        .map((formula) => formula.latex.trim())
        .filter(Boolean)
        .join("\n");
      if (!recognizedLatex) {
        throw new Error(
          isEn ? "OCR returned an empty formula" : "OCR 没有返回可用公式",
        );
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
        isEn ? "Pasted image converted to LaTeX" : "粘贴图片已转换为 LaTeX",
      );
      scheduleInlineOcrClear(1800);
    } catch (reason) {
      const message = errorMessage(reason, "");
      const cancelled =
        inlineOcrCancelRequestedRef.current || message.includes("OCR_CANCELLED");
      if (cancelled) {
        setInlineOcr((current) => ({
          status: "cancelled",
          message: isEn ? "OCR cancelled" : "OCR 已取消",
          seconds: current?.seconds ?? 0,
          model: ocrModel,
        }));
        scheduleInlineOcrClear(1200);
      } else {
        const visibleMessage =
          message || (isEn ? "Image OCR failed" : "图片 OCR 失败");
        setInlineOcr((current) => ({
          status: "error",
          message: visibleMessage,
          seconds: current?.seconds ?? 0,
          model: ocrModel,
        }));
        setToast(visibleMessage);
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

  const buildCurrentSessionUpdate = useCallback(
    async (
      status: "editing" | "committing" | "cancelled",
    ): Promise<UpdateOfficeSessionInput> => {
      if (!session) throw new Error("Office Session 尚未加载。");
      const exportResult =
        status === "cancelled"
          ? session.exportResult
          : session.host === "powerpoint" && USE_NATIVE_POWERPOINT_COMMIT
            ? generateSvgExportResult()
            : await getCompleteExportResult(currentFingerprint);
      if (status === "committing" && !exportResult) {
        throw new Error(isEn ? "Formula export is empty" : "公式导出结果为空");
      }
      return {
        title,
        lines,
        activeLineId,
        codeFormat: latexCodeFormat,
        displayMode,
        numbered: displayMode === "block" && numbered,
        fontSizePt: officeFontSizePt,
        formulaLetterFont,
        formulaChineseFont,
        dirty,
        status,
        autoCommitOnClose,
        explicitCancel: status === "cancelled",
        exportResult,
        exportWidth: exportResult?.width ?? 0,
        exportHeight: exportResult?.height ?? 0,
        error: null,
      };
    }, [
      session,
      title,
      lines,
      activeLineId,
      latexCodeFormat,
      displayMode,
      numbered,
      officeFontSizePt,
      formulaLetterFont,
      formulaChineseFont,
      dirty,
      autoCommitOnClose,
      currentFingerprint,
      generateSvgExportResult,
      getCompleteExportResult,
      isEn,
    ],
  );

  const saveCurrentSession = useCallback(
    async (status: "editing" | "committing" | "cancelled") => {
      const update = await buildCurrentSessionUpdate(status);
      const next = await save(update);
      lastSavedFingerprintRef.current = currentFingerprint;
      return next;
    },
    [buildCurrentSessionUpdate, currentFingerprint, save],
  );

  const closeOfficeEditorWindow = useCallback(async () => {
    if (!isMacosOfflineTauriTransport()) {
      window.close();
      return;
    }

    allowNativeCloseRef.current = true;
    try {
      await invokeTauri<void>("close_macos_offline_office_editor_window", {
        sessionId,
        generation,
      });
    } catch (error) {
      allowNativeCloseRef.current = false;
      throw error;
    }
  }, [generation, sessionId]);

  const handleCommit = useCallback(async (): Promise<boolean> => {
    // React state updates do not disable the button until the next render.
    // Keep a synchronous guard as well so a rapid double-click cannot enqueue
    // two commits for the same Office Session.
    if (finalizingRef.current) return false;
    const applyStartedEpochMs = Date.now();
    historyManager.commitPendingTransaction();
    if (!latex.trim()) {
      setToast(isEn ? "Enter a formula before inserting" : "请输入公式后再插入");
      return false;
    }
    const targetSessionKey = sessionKey;
    finalizingRef.current = true;
    try {
      if (isMacosOfflineTauriTransport()) {
        if (!session) throw new Error("Office Session 尚未加载。");
        const update = await buildCurrentSessionUpdate("committing");
        await commitMacosOfflineOfficeSession(
          session.id,
          update,
          applyStartedEpochMs,
        );
        lastSavedFingerprintRef.current = currentFingerprint;
        if (activeSessionKeyRef.current !== targetSessionKey) return false;
        try {
          await closeOfficeEditorWindow();
        } catch (closeError) {
          if (activeSessionKeyRef.current !== targetSessionKey) return false;
          finalizingRef.current = false;
          const detail = errorMessage(
            closeError,
            isEn ? "Unknown window error" : "未知窗口错误",
          );
          setToast(
            isEn
              ? `The formula was inserted, but the editor could not close: ${detail}`
              : `公式已经插入，但编辑窗口无法自动关闭：${detail}`,
          );
        }
        return true;
      }

      const next = await saveCurrentSession("committing");
      messageOfficeParent({ type: "visualtex-commit", sessionId: next.id });
      // The parent bridge owns both Word and PowerPoint mutations. Keep the
      // action busy until the host confirms the durable final state; a failed
      // PowerPoint decoration therefore leaves this editor open with a useful
      // error instead of closing after creating an anonymous Graphic shape.
      await waitForOfficeCommitResult(next.id, next.host);
      window.close();
      return true;
    } catch (error) {
      if (activeSessionKeyRef.current !== targetSessionKey) return false;
      finalizingRef.current = false;
      setToast(
        errorMessage(
          error,
          isEn ? "Unable to insert the Office formula" : "无法插入 Office 公式",
        ),
      );
      return false;
    }
  }, [
    buildCurrentSessionUpdate,
    closeOfficeEditorWindow,
    currentFingerprint,
    isEn,
    latex,
    saveCurrentSession,
    session,
    sessionKey,
  ]);

  useEffect(() => {
    if (
      !session ||
      !sessionHydrated ||
      (session.operation !== "nativeToImage" &&
        session.operation !== "imageToNative") ||
      silentCommitSessionKeyRef.current === sessionKey
    ) {
      return;
    }

    let cancelled = false;
    // Conversion editors stay hidden, so WebKit may never deliver an animation
    // frame. Use a task, as the hidden-editor readiness check does. Claim the
    // session only when that task starts: a hydration rerender can cancel an
    // earlier effect before it runs, and must still be able to schedule it again.
    const timer = window.setTimeout(() => {
      if (cancelled) return;
      silentCommitSessionKeyRef.current = sessionKey;
      void handleCommit().then(async (committed) => {
        if (committed || cancelled) return;
        // Direct conversion commands must never fall back to a visible editor.
        // Leave the Word object unchanged, cancel the failed hidden Session and
        // return the resident WebView to its parked state.
        try {
          await cancelMacosOfflineOfficeSession(session.id);
          await closeOfficeEditorWindow();
        } catch {
          // A superseding Office Session may already have cleared this one.
        }
      });
    }, 0);
    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [
    closeOfficeEditorWindow,
    handleCommit,
    session,
    sessionHydrated,
    sessionKey,
  ]);

  useEffect(() => {
    const handleApplyShortcut = (event: KeyboardEvent) => {
      if (
        event.isComposing ||
        event.key !== "Enter" ||
        (!event.metaKey && !event.ctrlKey) ||
        event.altKey ||
        event.shiftKey
      ) {
        return;
      }
      event.preventDefault();
      void handleCommit();
    };
    window.addEventListener("keydown", handleApplyShortcut, true);
    return () =>
      window.removeEventListener("keydown", handleApplyShortcut, true);
  }, [handleCommit]);

  useEffect(
    () =>
      registerOfficeApplyShortcut({
        onApply: async () => {
          await handleCommit();
        },
        isEnabled: () =>
          !ocrOpen &&
          !inlineOcrBusyRef.current &&
          !historyState.isReplaying &&
          !finalizingRef.current,
      }),
    [handleCommit, historyState.isReplaying, ocrOpen],
  );

  const handleCancel = useCallback(async () => {
    if (finalizingRef.current) return;
    const targetSessionKey = sessionKey;
    finalizingRef.current = true;
    try {
      const next = await saveCurrentSession("cancelled");
      if (isMacosOfflineTauriTransport()) {
        await cancelMacosOfflineOfficeSession(next.id);
        if (activeSessionKeyRef.current !== targetSessionKey) return;
        await closeOfficeEditorWindow();
        return;
      }
      if (next.host === "powerpoint") {
        window.close();
        return;
      }
      messageOfficeParent({ type: "visualtex-cancel", sessionId: next.id });
    } catch (error) {
      if (activeSessionKeyRef.current !== targetSessionKey) return;
      finalizingRef.current = false;
      setToast(
        errorMessage(
          error,
          isEn ? "Unable to cancel the Office formula" : "无法取消 Office 公式",
        ),
      );
    }
  }, [closeOfficeEditorWindow, isEn, saveCurrentSession, sessionKey]);

  useEffect(() => {
    if (!isMacosOfflineTauriTransport() || !sessionId) return;

    let disposed = false;
    let unlisten: (() => void) | undefined;
    void onCurrentTauriWindowCloseRequested((event) => {
      if (allowNativeCloseRef.current || disposed) return;
      event.preventDefault();
      if (nativeCloseRequestInFlightRef.current) return;
      nativeCloseRequestInFlightRef.current = true;

      const finalize = async () => {
        if (latex.trim() && autoCommitOnClose) {
          const committed = await handleCommit();
          if (committed || allowNativeCloseRef.current || disposed) return;
        }
        // Closing the window must never trap the user behind a malformed or
        // partially rendered formula. If close-to-commit cannot export the
        // current draft, cancel this edit Session and leave the Office object
        // unchanged instead of blocking the native close request.
        await handleCancel();
        if (allowNativeCloseRef.current || disposed) return;

        // A stale or already-cleaned Office Session can make both commit and
        // cancel fail. The native close request is still authoritative: allow
        // one final close event and destroy this WebView without mutating the
        // Office object, rather than trapping the user behind an error toast.
        allowNativeCloseRef.current = true;
        try {
          await closeCurrentTauriWindow();
        } catch {
          allowNativeCloseRef.current = false;
        }
      };
      void finalize().finally(() => {
        nativeCloseRequestInFlightRef.current = false;
      });
    })
      .then((dispose) => {
        if (disposed) {
          dispose();
        } else {
          unlisten = dispose;
        }
      })
      .catch((reason) => {
        const message = errorMessage(
          reason,
          isEn ? "Unknown window error" : "未知窗口错误",
        );
        setToast(
          isEn
            ? `Unable to register window close handling: ${message}`
            : `无法注册窗口关闭处理：${message}`,
        );
      });

    return () => {
      disposed = true;
      unlisten?.();
    };
  }, [autoCommitOnClose, handleCancel, handleCommit, isEn, latex, sessionId]);

  const handleCopy = async () => {
    await copyFormulaLines(lines, latexCodeFormat);
    addHistory(latex);
    setToast(isEn ? "LaTeX copied" : "LaTeX 已复制");
  };

  const editorAvailable = Boolean(session && sessionHydrated);
  const officeHeaderLeadingControls = editorAvailable && session ? (
    <>
      {session.host === "word" ? (
        <div
          className="office-display-mode-setting"
          role="group"
          aria-label={isEn ? "Word formula layout" : "Word 公式排版"}
        >
          <button
            type="button"
            className={displayMode === "inline" ? "is-active" : ""}
            onClick={() => {
              setDisplayMode("inline");
              setNumbered(false);
            }}
            disabled={session.mode === "edit"}
          >
            {isEn ? "Inline" : "行内"}
          </button>
          <button
            type="button"
            className={displayMode === "block" ? "is-active" : ""}
            onClick={() => {
              setDisplayMode("block");
              if (session.mode === "create") {
                setNumbered(
                  readOfficeWordCreateNumberedPreference(
                    Boolean(session.numbered),
                  ),
                );
              }
            }}
            disabled={session.mode === "edit"}
          >
            {isEn ? "Display" : "行间"}
          </button>
        </div>
      ) : null}
      <label
        className="office-font-size-setting"
        title={
          session.host === "word" && session.mode === "create"
            ? isEn
              ? "Starts from the current Word paragraph font size"
              : "默认读取当前 Word 段落正文的字号"
            : isEn
              ? "Formula font size"
              : "公式字号"
        }
      >
        <span>{isEn ? "Size" : "字号"}</span>
        <select
          value={officeFontSizePt}
          data-office-font-size
          aria-label={isEn ? "Formula font size" : "公式字号"}
          onChange={(event) =>
            setOfficeFontSizePt(
              normalizeOfficeFontSizePt(event.target.value, officeFontSizePt),
            )
          }
        >
          <optgroup label={isEn ? "Chinese sizes" : "中文字号"}>
            {OFFICE_CHINESE_FONT_SIZE_OPTIONS.map((option) => (
              <option key={option.name} value={option.fontSizePt}>
                {isEn
                  ? `${option.name} (${option.fontSizePt} pt)`
                  : `${option.name}（${option.fontSizePt} 磅）`}
              </option>
            ))}
          </optgroup>
          <optgroup label={isEn ? "Point sizes" : "磅值"}>
            {officePointFontSizeOptions(officeFontSizePt).map((fontSizePt) => (
              <option key={fontSizePt} value={fontSizePt}>
                {isEn ? `${fontSizePt} pt` : `${fontSizePt} 磅`}
              </option>
            ))}
          </optgroup>
        </select>
      </label>
      {session.host === "word" && displayMode === "block" ? (
        <label className="office-auto-commit-setting is-numbering-setting">
          <input
            type="checkbox"
            checked={numbered}
            onChange={(event) => {
              const nextNumbered = event.target.checked;
              setNumbered(nextNumbered);
              if (session.mode === "create") {
                writeOfficeWordCreateNumberedPreference(nextNumbered);
              }
            }}
            disabled={session.mode === "edit"}
          />
          <span>{isEn ? "Add equation number" : "添加公式编号"}</span>
        </label>
      ) : null}
    </>
  ) : null;
  const officeHeaderTrailingActions = editorAvailable && session ? (
    <>
      <button
        type="button"
        className="icon-button compact office-history-icon-button"
        data-office-undo-action
        aria-label={isEn ? "Undo" : "撤销"}
        title={isEn ? "Undo" : "撤销"}
        onClick={() => historyManager.requestUndo()}
        disabled={historyBusy || !historyState.canUndo || historyState.isReplaying}
      >
        <Undo2 size={16} strokeWidth={2} />
      </button>
      <button
        type="button"
        className="icon-button compact office-history-icon-button"
        data-office-redo-action
        aria-label={isEn ? "Redo" : "重做"}
        title={isEn ? "Redo" : "重做"}
        onClick={() => historyManager.requestRedo()}
        disabled={historyBusy || !historyState.canRedo || historyState.isReplaying}
      >
        <Redo2 size={16} strokeWidth={2} />
      </button>
      <span className="office-dialog-action-divider" aria-hidden="true" />
      <button
        type="button"
        className="secondary-button"
        data-office-cancel-action
        onClick={() => void handleCancel()}
        disabled={historyBusy || inlineOcrIsBusy || historyState.isReplaying}
      >
        {isEn ? "Cancel" : "取消"}
      </button>
      <button
        type="button"
        className="primary-button"
        data-office-primary-action
        onClick={() => void handleCommit()}
        disabled={
          historyBusy ||
          inlineOcrIsBusy ||
          historyState.isReplaying ||
          !latex.trim()
        }
        aria-keyshortcuts="Control+Enter Meta+Enter"
        title={
          session.mode === "edit"
            ? isEn
              ? "Update formula (Ctrl/Command+Enter)"
              : "更新公式（Ctrl/Command+Enter）"
            : isEn
              ? "Finish and insert (Ctrl/Command+Enter)"
              : "完成并插入（Ctrl/Command+Enter）"
        }
      >
        {session.mode === "edit"
          ? isEn
            ? "Update formula"
            : "更新公式"
          : isEn
            ? "Finish and insert"
            : "完成并插入"}
      </button>
    </>
  ) : null;
  const residentEditorWorkspace = (
    <div
      key="resident-office-editor-workspace"
      className="office-resident-editor-workspace"
      aria-hidden={!editorAvailable}
      style={
        editorAvailable
          ? { position: "relative", opacity: 1, pointerEvents: "auto" }
          : {
              position: "absolute",
              inset: 0,
              opacity: 0,
              pointerEvents: "none",
            }
      }
    >
      <EditorWorkspace
        mode={session?.mode === "edit" ? "office-edit" : "office-create"}
        showFileActions={false}
        showUpdateActions={false}
        showOfficeActions={false}
        showOcrActions={editorAvailable}
        officeHeaderLeadingControls={officeHeaderLeadingControls}
        officeHeaderTrailingActions={officeHeaderTrailingActions}
        editorRef={editorRef}
        editorInstanceKey="resident-office-editor"
        reuseEditorLineSlots
        sidebarOpen={sidebarOpen}
        onSidebarOpenChange={setSidebarOpen}
        onHistoryBusyChange={setHistoryBusy}
        onPasteImage={editorAvailable ? handleEditorImagePaste : undefined}
        onCopy={handleCopy}
        onReplaceDocument={replaceDocumentWithHistory}
        ocrModel={ocrModel}
        ocrModels={OCR_MODELS}
        ocrBusy={inlineOcrIsBusy}
        onOcrModelChange={(model) =>
          handleOcrModelChange(model as OcrModelName)
        }
        ocrOverlay={
          editorAvailable && inlineOcr ? (
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
                  {isEn ? inlineOcrModel.labelEn : inlineOcrModel.labelZh}
                  {" · "}
                  {inlineOcr.seconds}
                  {isEn ? "s" : " 秒"}
                </span>
              </div>
              {inlineOcrIsBusy ? (
                <button
                  type="button"
                  className="inline-ocr-cancel"
                  onClick={() => void cancelInlineOcr()}
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
    </div>
  );

  if (!tauriResidentEditor && (loading || (session && !sessionHydrated))) {
    return (
      <div className="office-dialog-state">
        <LoaderCircle className="is-spinning" size={28} />
        <strong>{isEn ? "Loading Office Session…" : "正在加载 Office Session…"}</strong>
      </div>
    );
  }

  if (!tauriResidentEditor && (error || !session)) {
    return (
      <div className="office-dialog-state is-error">
        <X size={28} />
        <strong>{isEn ? "Unable to open VisualTeX" : "无法打开 VisualTeX"}</strong>
        <p>{error || (isEn ? "Session not found" : "Session 不存在")}</p>
      </div>
    );
  }

  return (
    <div
      className="app-shell office-dialog-shell"
      style={{ position: "relative" }}
    >
      {tauriResidentEditor &&
      (loading || (session && !sessionHydrated) || error || !session) ? (
        <div
          className={`office-dialog-state ${error ? "is-error" : ""}`}
          style={{ position: "absolute", inset: 0, zIndex: 2 }}
        >
          {error ? <X size={28} /> : <LoaderCircle className="is-spinning" size={28} />}
          <strong>
            {error
              ? isEn
                ? "Unable to open VisualTeX"
                : "无法打开 VisualTeX"
              : isEn
                ? "Loading Office Session…"
                : "正在加载 Office Session…"}
          </strong>
          {error ? (
            <p>{error || (isEn ? "Session not found" : "Session 不存在")}</p>
          ) : null}
        </div>
      ) : null}

      {residentEditorWorkspace}

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

      {toast && (
        <div className="toast">
          <Check size={15} />
          {toast}
        </div>
      )}
    </div>
  );
}
