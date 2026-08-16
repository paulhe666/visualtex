import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { createUuid } from "../../runtime/browserCompatibility";
import {
  AlertCircle,
  Check,
  LoaderCircle,
  Redo2,
  ScanLine,
  Undo2,
  X,
} from "lucide-react";
import { OcrDialog } from "../../components/OcrDialog";
import { EditorWorkspace } from "../../workspace/EditorWorkspace";
import {
  readWorkspacePanelOpen,
  writeWorkspacePanelOpen,
} from "../../workspace/workspacePanelPreferences";
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
  normalizeEditorLayout,
  useEditorStore,
} from "../../stores/editorStore";
import {
  copyLatex,
  isLatexCodeFormat,
} from "../../clipboard/LatexCopyService";
import type {
  InputBehaviorSettingKey,
  LatexCodeFormat,
} from "../../types/formula";
import type {
  MathEditorHandle,
  MathEditorInsertionTarget,
} from "../../editor/MathEditor";
import { readErrorMessage } from "../../errors/readErrorMessage";
import { normalizeChineseLatex } from "../../editor/normalizeChineseLatex";
import {
  DEFAULT_FORMULA_CHINESE_FONT,
  DEFAULT_FORMULA_LETTER_FONT,
  type FormulaChineseFont,
  type FormulaLetterFont,
} from "../../editor/formulaFontPreferences";
import { readLocalStorage, writeLocalStorage } from "../../runtime/safeStorage";
import { latexToMathMl, latexToSvg } from "../../export/latexToSvg";
import { isIncompleteLatexDraft } from "../../math/latexCompatibility";
import {
  closeOfficeSessionWindow,
  getOfficePreferences,
  getOfficeSession,
  getOfficeTheme,
  saveOfficeSessionKeepalive,
  takeOfficeConverterBatch,
  updateOfficeSession,
  type OfficeExportResult,
  type OfficeFormulaSession,
  type OfficeHost,
  type OfficeObjectMode,
  type OfficePreferences,
} from "../api/sessionClient";
import { useOfficeSession } from "./useOfficeSession";
import { attachFormulaEquationTag } from "../shared/formulaEquationTag";
import {
  normalizeFormulaEditorDocument,
  serializeFormulaEditorRenderDocument,
  type FormulaEditorLine,
} from "../shared/formulaEditorDocument";
import { messageOfficeParent } from "./dialogMessages";
import { registerOfficeApplyShortcut } from "./officeApplyShortcut";
import {
  applyDocumentTheme,
  normalizeSynchronizedTheme,
  readPublishedSynchronizedTheme,
  readSynchronizedTheme,
  subscribeSynchronizedTheme,
} from "../../themeSync";
import { saveCustomTheme } from "../../themeCustomization";
import {
  DEFAULT_OCR_MODEL,
  OCR_MODELS,
  cancelOcrRecognition,
  fileToOcrRequest,
  getOcrRuntimeStatus,
  listenOcrRecognitionProgress,
  recognizeFormulaImage,
  warmupOcrModel,
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

const OFFICE_EDITOR_ZOOM_60_MIGRATION_KEY =
  "visualtex-office-editor-zoom-60-migration-v1";
const OCR_MODEL_STORAGE_KEY = "visualtex.ocr.model";
const USE_NATIVE_POWERPOINT_COMMIT =
  document
    .querySelector<HTMLMetaElement>(
      'meta[name="visualtex-native-powerpoint-commit"]',
    )
    ?.content.toLowerCase() === "true";

const OFFICE_COMMIT_RESULT_TIMEOUT_MS = 45_000;
const VSTO_RUNTIME = new URLSearchParams(window.location.search).get("runtime");
const IS_VSTO_DESKTOP_RUNTIME = VSTO_RUNTIME === "vsto-desktop";
const IS_VSTO_CONVERT_RUNTIME = VSTO_RUNTIME === "vsto-convert";

function delay(milliseconds: number) {
  return new Promise<void>((resolve) => window.setTimeout(resolve, milliseconds));
}

function officeExportCanCommit(
  session: Pick<OfficeFormulaSession, "host" | "objectMode">,
  exportResult: OfficeExportResult | null | undefined,
) {
  if (!exportResult) return false;
  const hasSvg = Boolean(exportResult.svg?.trim() || exportResult.svgBase64?.trim());
  if (
    session.objectMode === "wordOmml" ||
    session.objectMode === "mathTypeOle"
  ) {
    return exportResult.mathMl?.trimStart().startsWith("<math") === true;
  }
  if (session.host === "powerpoint" && USE_NATIVE_POWERPOINT_COMMIT) {
    return hasSvg;
  }
  if (session.objectMode === "nativeOle") {
    return hasSvg && Boolean(exportResult.pngBase64?.trim());
  }
  return Boolean(exportResult.pngBase64?.trim());
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

function normalizeOfficeCodeFormat(codeFormat: string): LatexCodeFormat {
  if (isLatexCodeFormat(codeFormat)) return codeFormat;
  // Older Office metadata used the generic value "latex" for raw formula
  // source. Treat it as the editor's equivalent raw format instead of
  // falling back to the persisted desktop preference and marking the
  // untouched formula as modified immediately after opening.
  return "raw";
}

function normalizeOfficeFormulaDocument(
  lines: FormulaEditorLine[],
  codeFormat: unknown,
) {
  return normalizeFormulaEditorDocument(
    lines,
    typeof codeFormat === "string" ? normalizeOfficeCodeFormat(codeFormat) : codeFormat,
  );
}

function serializeOfficeRenderLatex(
  lines: FormulaEditorLine[],
  codeFormat: unknown,
  displayMode: "inline" | "block",
  equationTag?: string | null,
) {
  const document = normalizeOfficeFormulaDocument(lines, codeFormat);
  const tag = displayMode === "block" ? equationTag?.trim() : "";
  if (!tag || document.lines.length === 0) {
    return serializeFormulaEditorRenderDocument(document);
  }
  const taggedLines = document.lines.map((line, index) =>
    index === document.lines.length - 1
      ? { ...line, latex: attachFormulaEquationTag(line.latex, tag) }
      : line,
  );
  return serializeFormulaEditorRenderDocument({ ...document, lines: taggedLines });
}

function normalizeOfficeFontSizePt(value: unknown, fallback: number) {
  const numeric = typeof value === "number" ? value : Number(value);
  const resolved = Number.isFinite(numeric) ? numeric : fallback;
  return Math.round(Math.min(200, Math.max(5, resolved)) * 2) / 2;
}

function applyOfficeEditorPreferences(
  preferences: OfficePreferences,
  applyTheme: (value: unknown) => void,
  applyEditorLayout: (value: unknown) => void,
) {
  const payload = preferences.editorPreferences;
  if (!payload) return;
  if (payload.customTheme) saveCustomTheme(payload.customTheme);
  const settings = payload.settings;
  if (!settings) return;

  const editor = useEditorStore.getState();
  if (settings.theme !== undefined) applyTheme(settings.theme);
  if (settings.editorLayout !== undefined) applyEditorLayout(settings.editorLayout);
  if (settings.language === "cn" || settings.language === "en") {
    editor.setLanguage(settings.language);
  }
  if (typeof settings.zoom === "number") editor.setZoom(settings.zoom);
  if (settings.formulaAlignment) {
    editor.setFormulaAlignment(settings.formulaAlignment);
  }
  if (typeof settings.autoPairDelimiters === "boolean") {
    editor.setAutoPairDelimiters(settings.autoPairDelimiters);
  }
  if (typeof settings.showLineNumbers === "boolean") {
    editor.setShowLineNumbers(settings.showLineNumbers);
  }
  if (typeof settings.highlightActiveLine === "boolean") {
    editor.setHighlightActiveLine(settings.highlightActiveLine);
  }
  if (typeof settings.formulaInsetLeft === "number") {
    editor.setFormulaInsetLeft(settings.formulaInsetLeft);
  }
  if (typeof settings.formulaInsetRight === "number") {
    editor.setFormulaInsetRight(settings.formulaInsetRight);
  }
  if (typeof settings.formulaToolButtonSize === "number") {
    editor.setFormulaToolButtonSize(settings.formulaToolButtonSize);
  }
  if (typeof settings.formulaToolButtonPadding === "number") {
    editor.setFormulaToolButtonPadding(settings.formulaToolButtonPadding);
  }
  if (typeof settings.formulaRowVerticalInset === "number") {
    editor.setFormulaRowVerticalInset(settings.formulaRowVerticalInset);
  }
  if (settings.pngExportBackground !== undefined) {
    editor.setPngExportBackground(settings.pngExportBackground);
  }
  if (settings.formulaLetterFont !== undefined) {
    editor.setFormulaLetterFont(settings.formulaLetterFont);
  }
  if (settings.formulaChineseFont !== undefined) {
    editor.setFormulaChineseFont(settings.formulaChineseFont);
  }
  if (settings.inputBehavior) {
    for (const [key, enabled] of Object.entries(settings.inputBehavior)) {
      if (typeof enabled === "boolean") {
        editor.setInputBehavior(key as InputBehaviorSettingKey, enabled);
      }
    }
  }
  if (typeof settings.keypadMinimizeOnCopy === "boolean") {
    editor.setKeypadMinimizeOnCopy(settings.keypadMinimizeOnCopy);
  }
}

function requireOfficeSessionFontSizePt(value: unknown, host: OfficeHost) {
  const numeric = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(numeric) || numeric < 5 || numeric > 200) {
    throw new Error(
      host === "word"
        ? "Word 公式 Session 缺少当前正文位置的有效字号。"
        : "PowerPoint 公式 Session 缺少有效字号。",
    );
  }
  return Math.round(numeric * 2) / 2;
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

function documentFingerprint(
  title: string,
  lines: Array<{ id: string; latex: string }>,
  codeFormat: string,
  displayMode: "inline" | "block",
  objectMode: OfficeObjectMode,
  numbered: boolean,
  fontSizePt: number,
  formulaLetterFont: FormulaLetterFont,
  formulaChineseFont: FormulaChineseFont,
) {
  return JSON.stringify({
    title,
    // Compare the same canonical LaTeX representation that MathEditor uses
    // after mount (upright e/i/d, Chinese text normalization, etc.). Raw
    // Session source can be semantically identical but serialize differently;
    // comparing raw text here can otherwise leave initial autosave suppressed
    // forever.
    lines: lines.map((line) => normalizeChineseLatex(line.latex)),
    codeFormat,
    displayMode,
    objectMode,
    numbered,
    fontSizePt: normalizeOfficeFontSizePt(fontSizePt, fontSizePt),
    formulaLetterFont,
    formulaChineseFont,
  });
}

export function OfficeDialogApp() {
  const editorRef = useRef<MathEditorHandle>(null);
  const loadedSessionIdRef = useRef("");
  const skipAutosaveForSessionRef = useRef("");
  const originalFingerprintRef = useRef("");
  const loadedUiFingerprintRef = useRef("");
  const lastSavedFingerprintRef = useRef("");
  const readyMessageSentRef = useRef(false);
  const finalizingRef = useRef(false);
  const commitFromShortcutRef = useRef<() => void>(() => undefined);
  const closeFromNativeWindowRef = useRef<() => void>(() => undefined);
  const exportRunIdRef = useRef(0);
  const conversionStartedRef = useRef(false);
  const batchConversionQueueRef = useRef<Promise<void>>(Promise.resolve());
  const initialEditorFocusSessionRef = useRef("");
  const latestCompleteExportRef = useRef<{
    fingerprint: string;
    exportResult: OfficeExportResult;
  } | null>(null);
  const [sidebarOpen, setSidebarOpen] = useState(() =>
    readWorkspacePanelOpen("office-edit", "tiles", window.innerWidth >= 1040),
  );
  const [historyBusy, setHistoryBusy] = useState(false);
  const [autoCommitOnClose, setAutoCommitOnClose] = useState(true);
  const [displayMode, setDisplayMode] = useState<"inline" | "block">("inline");
  const [objectMode, setObjectMode] = useState<OfficeObjectMode>("nativeOle");
  const [numbered, setNumbered] = useState(false);
  // This is only a pre-Session UI placeholder. Word never uses it as a
  // default: every Word Session must supply the font size read at the current
  // document insertion point. PowerPoint's configured default is stored in a
  // separate state below.
  const [officeFontSizePt, setOfficeFontSizePt] = useState(14);
  const [powerPointDefaultFontSizePt, setPowerPointDefaultFontSizePt] =
    useState(20);
  const [officePreferencesReady, setOfficePreferencesReady] = useState(false);
  const [toast, setToast] = useState("");
  const [ocrOpen, setOcrOpen] = useState(false);
  const [ocrModel, setOcrModel] = useState<OcrModelName>(() => {
    const stored = window.localStorage.getItem(OCR_MODEL_STORAGE_KEY);
    return OCR_MODELS.some((item) => item.id === stored)
      ? (stored as OcrModelName)
      : DEFAULT_OCR_MODEL;
  });
  const [inlineOcr, setInlineOcr] = useState<InlineOcrState | null>(null);
  const startupOcrModelRef = useRef(ocrModel);
  const inlineOcrBusyRef = useRef(false);
  const inlineOcrCancelRequestedRef = useRef(false);
  const inlineOcrRunIdRef = useRef(0);
  const inlineOcrClearTimerRef = useRef<number | null>(null);
  const { sessionId, session, loading, error, reload, save } = useOfficeSession();

  useEffect(() => {
    if (readLocalStorage(OFFICE_EDITOR_ZOOM_60_MIGRATION_KEY) !== "done") {
      useEditorStore.getState().setZoom(0.6);
      writeLocalStorage(OFFICE_EDITOR_ZOOM_60_MIGRATION_KEY, "done");
    }
  }, []);

  useEffect(() => {
    loadedSessionIdRef.current = "";
    skipAutosaveForSessionRef.current = "";
    originalFingerprintRef.current = "";
    loadedUiFingerprintRef.current = "";
    lastSavedFingerprintRef.current = "";
    readyMessageSentRef.current = false;
    finalizingRef.current = false;
    conversionStartedRef.current = false;
    initialEditorFocusSessionRef.current = "";
    latestCompleteExportRef.current = null;
    exportRunIdRef.current += 1;
  }, [sessionId]);

  const title = useEditorStore((state) => state.title);
  const lines = useEditorStore((state) => state.lines);
  const activeLineId = useEditorStore((state) => state.activeLineId);
  const language = useEditorStore((state) => state.language);
  const theme = useEditorStore((state) => state.theme);
  const setTheme = useEditorStore((state) => state.setTheme);
  const setEditorLayout = useEditorStore((state) => state.setEditorLayout);
  const latexCodeFormat = useEditorStore((state) => state.latexCodeFormat);
  const formulaLetterFont = useEditorStore((state) => state.formulaLetterFont);
  const formulaChineseFont = useEditorStore((state) => state.formulaChineseFont);
  const addHistory = useEditorStore((state) => state.addHistory);
  const historyState = useHistorySnapshot();
  const isEn = language === "en";
  const latex = joinFormulaLines(lines);

  useEffect(() => {
    let disposed = false;
    void getOfficePreferences()
      .then((preferences) => {
        if (!disposed) {
          setPowerPointDefaultFontSizePt(
            normalizeOfficeFontSizePt(
              preferences.powerpointDefaultFontSizePt,
              20,
            ),
          );
        }
      })
      .catch(() => {
        if (!disposed) setPowerPointDefaultFontSizePt(20);
      })
      .finally(() => {
        // Do not mark the full Office editor preferences as ready here. This
        // request only seeds the PowerPoint default font size; the synchronized
        // editor-preference request below is the one that actually applies
        // formula fonts/theme/layout and owns officePreferencesReady.
      });
    return () => {
      disposed = true;
    };
  }, []);

  useEffect(() => {
    let disposed = false;
    const applyTheme = (nextThemeValue: unknown) => {
      const nextTheme = normalizeSynchronizedTheme(nextThemeValue);
      applyDocumentTheme(nextTheme);
      if (useEditorStore.getState().theme !== nextTheme) {
        setTheme(nextTheme);
      }
    };
    const applyEditorLayout = (nextLayoutValue: unknown) => {
      const nextLayout = normalizeEditorLayout(nextLayoutValue);
      if (useEditorStore.getState().editorLayout !== nextLayout) {
        setEditorLayout(nextLayout);
      }
    };
    const syncFromCompanion = async () => {
      try {
        const [status, preferences] = await Promise.all([
          getOfficeTheme(),
          getOfficePreferences(),
        ]);
        if (!disposed) {
          // Shared visual/editor preferences stay live, but Office-window view
          // state (source tab and resizable panel dimensions) is deliberately
          // excluded inside applyOfficeEditorPreferences. Those values persist
          // in this Office WebView and must never be overwritten by the main
          // app's 500ms companion snapshot.
          applyOfficeEditorPreferences(
            preferences,
            applyTheme,
            applyEditorLayout,
          );
          applyTheme(status.theme);
          applyEditorLayout(status.editorLayout);
          setPowerPointDefaultFontSizePt(
            normalizeOfficeFontSizePt(
              preferences.powerpointDefaultFontSizePt,
              20,
            ),
          );
          setOfficePreferencesReady(true);
        }
      } catch {
        // Keep the last applied appearance while the companion is restarting.
      }
    };

    applyTheme(readSynchronizedTheme());
    const unsubscribeBrowser = subscribeSynchronizedTheme(applyTheme);
    void syncFromCompanion();
    const interval = window.setInterval(() => void syncFromCompanion(), 500);
    // Browser/Vite parity and independent Office WebViews can miss a storage or
    // BroadcastChannel notification during creation. Re-read the published
    // active theme without the one-shot ?theme= bootstrap value so later main
    // window theme changes always win.
    const publishedThemeInterval = window.setInterval(() => {
      applyTheme(readPublishedSynchronizedTheme());
    }, 80);
    const handleFocus = () => void syncFromCompanion();
    const handleVisibility = () => {
      if (document.visibilityState === "visible") void syncFromCompanion();
    };
    window.addEventListener("focus", handleFocus);
    document.addEventListener("visibilitychange", handleVisibility);

    return () => {
      disposed = true;
      unsubscribeBrowser();
      window.clearInterval(interval);
      window.clearInterval(publishedThemeInterval);
      window.removeEventListener("focus", handleFocus);
      document.removeEventListener("visibilitychange", handleVisibility);
    };
  }, [setEditorLayout, setTheme]);

  useEffect(() => {
    applyDocumentTheme(theme);
  }, [theme]);

  const handleSidebarOpenChange = useCallback((open: boolean) => {
    setSidebarOpen(open);
    writeWorkspacePanelOpen("office-edit", "tiles", open);
  }, []);

  const selectedOcrModel =
    OCR_MODELS.find((item) => item.id === ocrModel) ??
    OCR_MODELS.find((item) => item.id === DEFAULT_OCR_MODEL)!;
  const inlineOcrModel =
    OCR_MODELS.find((item) => item.id === inlineOcr?.model) ?? selectedOcrModel;
  const inlineOcrIsBusy =
    inlineOcr?.status === "running" || inlineOcr?.status === "cancelling";

  const currentFingerprint = useMemo(
    () =>
      documentFingerprint(
        title,
        lines,
        latexCodeFormat,
        displayMode,
        objectMode,
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
      objectMode,
      numbered,
      officeFontSizePt,
      formulaLetterFont,
      formulaChineseFont,
    ],
  );
  const dirty =
    Boolean(session) &&
    Boolean(originalFingerprintRef.current) &&
    currentFingerprint !== originalFingerprintRef.current;

  useEffect(() => {
    if (!session || loadedSessionIdRef.current === session.id) return;
    // Font/layout preferences now participate in the immutable Office formula
    // fingerprint and rendered OLE/OMML output. Do not establish the Session
    // baseline before companion preferences have been applied, otherwise a
    // later preference sync looks like a half-loaded render and autosave can
    // suppress the required font redraw indefinitely.
    if (!officePreferencesReady) return;
    loadedSessionIdRef.current = session.id;
    skipAutosaveForSessionRef.current = session.id;
    const sourceLines = session.lines.length
      ? session.lines
      : [{ id: createUuid(), latex: "" }];
    const normalizedDocument = normalizeOfficeFormulaDocument(
      sourceLines,
      session.codeFormat,
    );
    const nextLines = normalizedDocument.lines;
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
    const loadedCodeFormat = normalizedDocument.codeFormat;
    useEditorStore.getState().setLatexCodeFormat(loadedCodeFormat);
    setAutoCommitOnClose(session.autoCommitOnClose);
    setDisplayMode(session.displayMode);
    setObjectMode(session.objectMode);
    const loadedNumbered =
      session.objectMode !== "mathTypeOle" &&
      session.displayMode === "block" &&
      Boolean(session.numbered);
    setNumbered(loadedNumbered);
    const loadedFontSizePt =
      session.host === "powerpoint" &&
      session.mode === "create" &&
      session.status === "created" &&
      !session.dirty
        ? powerPointDefaultFontSizePt
        : requireOfficeSessionFontSizePt(
            session.fontSizePt ??
              session.originalMetadata?.fontSizePt ??
              session.originalMetadata?.renderFontSizePt,
            session.host,
          );
    setOfficeFontSizePt(loadedFontSizePt);
    const loadedFingerprint = documentFingerprint(
      session.title,
      nextLines,
      loadedCodeFormat,
      session.displayMode,
      session.objectMode,
      loadedNumbered,
      loadedFontSizePt,
      session.originalMetadata?.formulaLetterFont ?? DEFAULT_FORMULA_LETTER_FONT,
      session.originalMetadata?.formulaChineseFont ?? DEFAULT_FORMULA_CHINESE_FONT,
    );
    const loadedUiFingerprint = documentFingerprint(
      session.title,
      nextLines,
      loadedCodeFormat,
      session.displayMode,
      session.objectMode,
      loadedNumbered,
      loadedFontSizePt,
      formulaLetterFont,
      formulaChineseFont,
    );
    originalFingerprintRef.current = loadedFingerprint;
    loadedUiFingerprintRef.current = loadedUiFingerprint;
    lastSavedFingerprintRef.current = loadedFingerprint;
    latestCompleteExportRef.current = session.exportResult?.pngBase64
      ? { fingerprint: loadedFingerprint, exportResult: session.exportResult }
      : null;
  }, [
    session?.id,
    isEn,
    officePreferencesReady,
    powerPointDefaultFontSizePt,
    formulaLetterFont,
    formulaChineseFont,
  ]);

  useEffect(() => {
    if (
      IS_VSTO_CONVERT_RUNTIME ||
      !session ||
      loadedSessionIdRef.current !== session.id ||
      initialEditorFocusSessionRef.current === session.id
    ) {
      return;
    }
    initialEditorFocusSessionRef.current = session.id;
    const repairTimers: number[] = [];
    let repairInterval = 0;
    const focusFirstLine = () => {
      window.focus();
      editorRef.current?.focus({ target: "first", moveToEnd: true });
      // MathLive can finish mounting its shadow input one frame after the
      // editor imperative handle becomes available. Keep a direct custom-
      // element fallback so the Office dialog does not leave focus on BODY.
      const field = document.querySelector<HTMLElement>("math-field");
      if (field && document.activeElement !== field) {
        field.focus({ preventScroll: true });
      }
    };
    const formulaHasFocus = () =>
      document.activeElement?.tagName === "MATH-FIELD";
    const focusWhenVisible = () => {
      if (document.visibilityState === "visible") focusFirstLine();
    };
    const stopActivationRepair = () => {
      window.clearInterval(repairInterval);
      repairInterval = 0;
      window.removeEventListener("focus", focusFirstLine);
      document.removeEventListener("visibilitychange", focusWhenVisible);
    };
    const repairFocus = () => {
      focusFirstLine();
      if (formulaHasFocus()) stopActivationRepair();
    };
    const frame = window.requestAnimationFrame(() => {
      repairFocus();
      if (!formulaHasFocus()) {
        repairInterval = window.setInterval(repairFocus, 60);
        repairTimers.push(window.setTimeout(stopActivationRepair, 3000));
      }
    });
    window.addEventListener("focus", focusFirstLine, { once: true });
    document.addEventListener("visibilitychange", focusWhenVisible);
    return () => {
      window.cancelAnimationFrame(frame);
      repairTimers.forEach((timer) => window.clearTimeout(timer));
      stopActivationRepair();
    };
  }, [session?.id]);

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
      historyManager.commitPendingTransaction();
      const before = captureSnapshot();
      if (documentSnapshotsEquivalent(before, after)) return false;
      useEditorStore.getState().replaceDocumentState(after);
      historyManager.push({
        type: "replace-document",
        before,
        after,
        source,
        timestamp: Date.now(),
      });
      window.requestAnimationFrame(() => restoreSnapshotFocus(after));
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

  const currentRenderedLatex = useMemo(
    () =>
      serializeOfficeRenderLatex(
        lines,
        latexCodeFormat,
        displayMode,
        session?.originalMetadata?.equationTag,
      ),
    [
      lines,
      latexCodeFormat,
      displayMode,
      session?.originalMetadata?.equationTag,
    ],
  );

  const generateSvgExportResult = useCallback((
    sourceLatex: string = currentRenderedLatex,
    sourceDisplayMode: "inline" | "block" = displayMode,
    sourceFontSizePt: number = officeFontSizePt,
    sourceFormulaLetterFont: FormulaLetterFont = formulaLetterFont,
    sourceFormulaChineseFont: FormulaChineseFont = formulaChineseFont,
  ): OfficeExportResult | null => {
    if (!sourceLatex.trim()) return null;
    const svg = latexToSvg(sourceLatex, {
      displayMode: sourceDisplayMode === "block",
      fontSizePt: sourceFontSizePt,
      paddingPx: sourceDisplayMode === "inline" ? 1 : 10,
      background: "transparent",
      formulaLetterFont: sourceFormulaLetterFont,
      formulaChineseFont: sourceFormulaChineseFont,
    });
    return {
      svg: svg.svg,
      svgBase64: svg.base64,
      mathMl: latexToMathMl(sourceLatex, sourceDisplayMode === "block"),
      width: svg.width,
      height: svg.height,
      baseline: svg.baseline,
      formulaLetterFont: sourceFormulaLetterFont,
      formulaChineseFont: sourceFormulaChineseFont,
    };
  }, [
    currentRenderedLatex,
    displayMode,
    officeFontSizePt,
    formulaLetterFont,
    formulaChineseFont,
  ]);

  const rasterizeSvgExportResult = useCallback(async (
    base: OfficeExportResult,
  ): Promise<OfficeExportResult> => {
    let pngBase64: string | undefined;
    try {
      const { svgToPng } = await import("../../export/svgToPng");
      pngBase64 = (
        await svgToPng(
          {
            svg: base.svg,
            base64: base.svgBase64,
            width: base.width,
            height: base.height,
            baseline: base.baseline,
          },
          { scale: 2, background: "transparent" },
        )
      ).base64;
    } catch {
      // SVG remains a complete Office fallback when PNG rasterization fails.
    }
    return { ...base, pngBase64 };
  }, []);

  const generateExportResult = useCallback(async (
    sourceLatex: string = currentRenderedLatex,
    sourceDisplayMode: "inline" | "block" = displayMode,
    sourceFontSizePt: number = officeFontSizePt,
  ): Promise<OfficeExportResult | null> => {
    const base = generateSvgExportResult(
      sourceLatex,
      sourceDisplayMode,
      sourceFontSizePt,
    );
    return base ? rasterizeSvgExportResult(base) : null;
  }, [
    currentRenderedLatex,
    displayMode,
    officeFontSizePt,
    generateSvgExportResult,
    rasterizeSvgExportResult,
  ]);

  const generateSessionExportResult = useCallback(async (
    sourceSession: OfficeFormulaSession,
  ): Promise<OfficeExportResult> => {
    const sourceLines = sourceSession.lines.length
      ? sourceSession.lines
      : [{ id: createUuid(), latex: "" }];
    const conversionDocument = normalizeOfficeFormulaDocument(
      sourceLines,
      sourceSession.codeFormat,
    );
    const conversionLines = conversionDocument.lines;
    const conversionLatex = serializeOfficeRenderLatex(
      conversionLines,
      conversionDocument.codeFormat,
      sourceSession.displayMode,
      sourceSession.originalMetadata?.equationTag,
    );
    const conversionFontSizePt = requireOfficeSessionFontSizePt(
      sourceSession.fontSizePt ??
        sourceSession.originalMetadata?.fontSizePt ??
        sourceSession.originalMetadata?.renderFontSizePt,
      sourceSession.host,
    );
    const sourceFormulaLetterFont =
      sourceSession.originalMetadata?.formulaLetterFont ?? formulaLetterFont;
    const sourceFormulaChineseFont =
      sourceSession.originalMetadata?.formulaChineseFont ?? formulaChineseFont;
    if (sourceSession.objectMode === "wordOmml") {
      const mathMl = latexToMathMl(
        conversionLatex,
        sourceSession.displayMode === "block",
      );
      if (!mathMl?.trim()) {
        throw new Error("Unable to generate MathML for Office conversion.");
      }
      const width =
        sourceSession.originalMetadata?.renderWidthPx ??
        (sourceSession.exportWidth > 0 ? sourceSession.exportWidth : 240);
      const height =
        sourceSession.originalMetadata?.renderHeightPx ??
        (sourceSession.exportHeight > 0 ? sourceSession.exportHeight : 80);
      return {
        svg: "",
        svgBase64: "",
        mathMl,
        width,
        height,
        baseline: sourceSession.originalMetadata?.baseline ?? height * 0.75,
        formulaLetterFont: sourceFormulaLetterFont,
        formulaChineseFont: sourceFormulaChineseFont,
      };
    }
    const base = generateSvgExportResult(
      conversionLatex,
      sourceSession.displayMode,
      conversionFontSizePt,
      sourceFormulaLetterFont,
      sourceFormulaChineseFont,
    );
    if (!base?.mathMl) {
      throw new Error("Unable to generate MathML for Office conversion.");
    }
    const complete = await rasterizeSvgExportResult(base);
    if (!complete?.mathMl || !complete.pngBase64) {
      throw new Error("Unable to generate a complete Office formula preview.");
    }
    return complete;
  }, [
    generateSvgExportResult,
    rasterizeSvgExportResult,
    formulaLetterFont,
    formulaChineseFont,
  ]);

  useEffect(() => {
    if (
      !IS_VSTO_CONVERT_RUNTIME ||
      !session ||
      !sessionId ||
      conversionStartedRef.current
    ) {
      return;
    }
    conversionStartedRef.current = true;
    void (async () => {
      try {
        // Use the immutable Session snapshot. The editor store is populated in
        // another React effect and can still contain a previous formula during
        // the first hidden conversion render.
        const sourceLines = session.lines.length
          ? session.lines
          : [{ id: createUuid(), latex: "" }];
        const conversionDocument = normalizeOfficeFormulaDocument(
          sourceLines,
          session.codeFormat,
        );
        const conversionLines = conversionDocument.lines;
        const conversionDisplayMode = session.displayMode;
        const conversionNumbered =
          conversionDisplayMode === "block" && Boolean(session.numbered);
        const conversionActiveLineId =
          session.activeLineId &&
          conversionLines.some((line) => line.id === session.activeLineId)
            ? session.activeLineId
            : conversionLines[0]?.id ?? null;
        // The hidden converter can run in the same React commit that loads a
        // new Session. Rendering from component state here used the previous
        // 20 pt PowerPoint value before Word's paragraph size had reached the
        // state hook. Always render from the immutable Session value instead.
        const conversionFontSizePt = requireOfficeSessionFontSizePt(
          session.fontSizePt ??
            session.originalMetadata?.fontSizePt ??
            session.originalMetadata?.renderFontSizePt,
          session.host,
        );
        const exportResult = await generateSessionExportResult(session);
        await save({
          title: session.title,
          lines: conversionLines,
          activeLineId: conversionActiveLineId,
          codeFormat: conversionDocument.codeFormat,
          displayMode: conversionDisplayMode,
          numbered: conversionNumbered,
          fontSizePt: conversionFontSizePt,
          dirty: false,
          status: "committing",
          autoCommitOnClose: false,
          exportResult,
          exportWidth: exportResult.width,
          exportHeight: exportResult.height,
          error: null,
        });
      } catch (reason) {
        const detail = readErrorMessage(
          reason,
          isEn ? "Formula format conversion failed" : "公式格式转换失败",
        );
        const sourceFormula = joinFormulaLines(session.lines).trim();
        const formulaPreview = sourceFormula.length <= 500
          ? sourceFormula
          : `${sourceFormula.slice(0, 500)}…`;
        const message = isEn
          ? `Formula rendering failed: ${detail}\nFormula: ${formulaPreview}`
          : `公式渲染失败：${detail}\n公式：${formulaPreview}`;
        try {
          await save({ status: "failed", error: message });
        } catch {
          // VSTO reports a timeout if the failure cannot be persisted.
        }
      } finally {
        try {
          await closeOfficeSessionWindow(sessionId);
        } catch {
          // Completion also removes the hidden conversion window.
        }
      }
    })();
  }, [session?.id, sessionId, isEn, save, generateSessionExportResult]);

  useEffect(() => {
    if (!IS_VSTO_CONVERT_RUNTIME) return;

    let disposed = false;
    let requestInFlight = false;
    const enqueue = (sessionIds: string[]) => {
      const uniqueSessionIds = Array.from(
        new Set(sessionIds.map((value) => value.trim()).filter(Boolean)),
      );
      if (!uniqueSessionIds.length) return;
      batchConversionQueueRef.current = batchConversionQueueRef.current
        .catch(() => undefined)
        .then(async () => {
          for (const queuedSessionId of uniqueSessionIds) {
            let queuedSession: OfficeFormulaSession | null = null;
            try {
              queuedSession = await getOfficeSession(queuedSessionId);
              const sourceLines = queuedSession.lines.length
                ? queuedSession.lines
                : [{ id: createUuid(), latex: "" }];
              const conversionDocument = normalizeOfficeFormulaDocument(
                sourceLines,
                queuedSession.codeFormat,
              );
              const conversionLines = conversionDocument.lines;
              const conversionFontSizePt = requireOfficeSessionFontSizePt(
                queuedSession.fontSizePt ??
                  queuedSession.originalMetadata?.fontSizePt ??
                  queuedSession.originalMetadata?.renderFontSizePt,
                queuedSession.host,
              );
              const exportResult = await generateSessionExportResult(queuedSession);
              await updateOfficeSession(queuedSessionId, {
                title: queuedSession.title,
                lines: conversionLines,
                activeLineId:
                  queuedSession.activeLineId &&
                  conversionLines.some((line) => line.id === queuedSession?.activeLineId)
                    ? queuedSession.activeLineId
                    : conversionLines[0]?.id ?? null,
                codeFormat: conversionDocument.codeFormat,
                displayMode: queuedSession.displayMode,
                numbered:
                  queuedSession.displayMode === "block" && Boolean(queuedSession.numbered),
                fontSizePt: conversionFontSizePt,
                dirty: false,
                status: "committing",
                autoCommitOnClose: false,
                exportResult,
                exportWidth: exportResult.width,
                exportHeight: exportResult.height,
                error: null,
              });
            } catch (reason) {
              const detail = readErrorMessage(reason, "Formula format conversion failed");
              const sourceFormula = queuedSession
                ? joinFormulaLines(queuedSession.lines).trim()
                : "";
              const formulaPreview = sourceFormula.length <= 500
                ? sourceFormula
                : `${sourceFormula.slice(0, 500)}…`;
              const message = formulaPreview
                ? `Formula rendering failed: ${detail}\nFormula: ${formulaPreview}`
                : `Formula rendering failed: ${detail}`;
              try {
                await updateOfficeSession(queuedSessionId, {
                  status: "failed",
                  error: message,
                });
              } catch {
                // VSTO reports a timeout if the failure cannot be persisted.
              }
            }
          }
        });
    };

    const drainCompanionQueue = async () => {
      if (disposed || requestInFlight) return;
      requestInFlight = true;
      try {
        while (!disposed) {
          const batch = await takeOfficeConverterBatch();
          if (!batch.sessionIds.length) break;
          enqueue(batch.sessionIds);
        }
      } catch {
        // The next poll retries after transient companion or startup failures.
      } finally {
        requestInFlight = false;
      }
    };
    const handleBatchConversion = () => {
      void drainCompanionQueue();
    };

    window.addEventListener(
      "visualtex-office-batch-conversion",
      handleBatchConversion,
    );
    void drainCompanionQueue();
    const pollTimer = window.setInterval(() => {
      void drainCompanionQueue();
    }, 100);
    return () => {
      disposed = true;
      window.clearInterval(pollTimer);
      window.removeEventListener(
        "visualtex-office-batch-conversion",
        handleBatchConversion,
      );
    };
  }, [generateSessionExportResult]);

  useEffect(() => {
    if (IS_VSTO_CONVERT_RUNTIME) return;
    if (!session || !sessionId || finalizingRef.current) return;
    if (loadedSessionIdRef.current !== sessionId) return;
    if (skipAutosaveForSessionRef.current === sessionId) {
      // Loading a Session updates the external editor store and several React
      // states in the same effect. Intermediate renders can therefore contain
      // the new LaTeX with the previous display mode or font size. Keep initial
      // autosave suppressed until every field matches the immutable Session
      // fingerprint; otherwise the half-loaded render is briefly persisted as
      // a false dirty edit.
      if (currentFingerprint !== loadedUiFingerprintRef.current) return;
      skipAutosaveForSessionRef.current = "";
      if (currentFingerprint === originalFingerprintRef.current) {
        lastSavedFingerprintRef.current = currentFingerprint;
        return;
      }
      // The UI is fully loaded, but a persisted/global visual preference (for
      // example formula fonts) differs from the formula's original metadata.
      // Fall through so this stable difference is exported instead of being
      // mistaken for a half-loaded Session.
    }
    if (
      lastSavedFingerprintRef.current === currentFingerprint &&
      session.autoCommitOnClose === autoCommitOnClose
    ) {
      return;
    }

    const runId = ++exportRunIdRef.current;
    const saveIncompleteDraft = () => {
      latestCompleteExportRef.current = null;
      void save({
        title,
        lines,
        activeLineId,
        codeFormat: latexCodeFormat,
        displayMode,
        objectMode,
        numbered: displayMode === "block" && numbered,
        fontSizePt: officeFontSizePt,
        dirty,
        status: "editing",
        autoCommitOnClose,
        exportResult: null,
        exportWidth: 0,
        exportHeight: 0,
        error: null,
      })
        .then((saved) => {
          if (saved && runId === exportRunIdRef.current) {
            lastSavedFingerprintRef.current = currentFingerprint;
          }
        })
        .catch((reason) => {
          setToast(
            readErrorMessage(
              reason,
              isEn ? "Unable to save the Office formula" : "无法保存 Office 公式",
            ),
          );
        });
    };
    if (isIncompleteLatexDraft(latex)) {
      saveIncompleteDraft();
      return;
    }
    try {
      // MathJax SVG generation is synchronous. Persist it immediately instead
      // of waiting for PNG rasterization, so closing the Office dialog cannot
      // lose the final keystrokes.
      const exportResult = generateSvgExportResult();
      const draftUpdate = {
        title,
        lines,
        activeLineId,
        codeFormat: latexCodeFormat,
        displayMode,
        objectMode,
        numbered: displayMode === "block" && numbered,
        fontSizePt: officeFontSizePt,
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
          const message =
            reason instanceof Error
              ? reason.message
              : isEn
                ? "Unable to save the Office formula"
                : "无法保存 Office 公式";
          setToast(message);
        });
      // Windows OLE inserts a PNG file. Keep rasterization off the critical
      // keystroke-save path, but persist the full export as soon as it is
      // ready so the title-bar close button has a committable final draft.
      if (
        exportResult &&
        !(session.host === "powerpoint" && USE_NATIVE_POWERPOINT_COMMIT)
      ) {
        void generateExportResult()
          .then((completeExport) => {
            if (
              !completeExport?.pngBase64 ||
              runId !== exportRunIdRef.current ||
              finalizingRef.current
            ) {
              return;
            }
            latestCompleteExportRef.current = {
              fingerprint: currentFingerprint,
              exportResult: completeExport,
            };
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
    } catch (reason) {
      if (isIncompleteLatexDraft(latex, reason)) {
        saveIncompleteDraft();
        return;
      }
      const message =
        reason instanceof Error
          ? reason.message
          : isEn
            ? "Unable to export the Office formula"
            : "无法导出 Office 公式";
      setToast(message);
    }
  }, [
    sessionId,
    session?.id,
    session?.autoCommitOnClose,
    currentFingerprint,
    title,
    lines,
    activeLineId,
    latexCodeFormat,
    displayMode,
    objectMode,
    numbered,
    officeFontSizePt,
    dirty,
    autoCommitOnClose,
    save,
    isEn,
    generateSvgExportResult,
    generateExportResult,
  ]);

  useEffect(() => {
    if (IS_VSTO_CONVERT_RUNTIME || !sessionId) return;
    const finalDraftUpdate = (status: "editing" | "committing") => {
      const cached = latestCompleteExportRef.current;
      const unchangedEdit = session?.mode === "edit" && !dirty;
      const exportResult = unchangedEdit
        ? cached?.fingerprint === currentFingerprint
          ? cached.exportResult
          : session?.exportResult ?? null
        : cached?.fingerprint === currentFingerprint
          ? cached.exportResult
          : isIncompleteLatexDraft(latex)
            ? null
            : generateSvgExportResult();
      return {
        title,
        lines,
        activeLineId,
        codeFormat: latexCodeFormat,
        displayMode,
        objectMode,
        numbered: displayMode === "block" && numbered,
        fontSizePt: officeFontSizePt,
        dirty,
        status,
        autoCommitOnClose,
        explicitCancel: false,
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
    const cancelFinalDraft = () => {
      try {
        void saveOfficeSessionKeepalive(sessionId, {
          ...finalDraftUpdate("editing"),
          status: "cancelled",
          explicitCancel: true,
        }).catch(() => undefined);
      } catch {
        void saveOfficeSessionKeepalive(sessionId, {
          status: "cancelled",
          explicitCancel: true,
          error: null,
        }).catch(() => undefined);
      }
    };
    const commitFinalDraft = () => {
      if (finalizingRef.current) return;
      const unchangedEdit = session?.mode === "edit" && !dirty;
      if (!autoCommitOnClose || (!unchangedEdit && !latex.trim())) {
        cancelFinalDraft();
        return;
      }
      try {
        const update = finalDraftUpdate("committing");
        if (
          !session ||
          (!unchangedEdit &&
            !officeExportCanCommit(
              { host: session.host, objectMode },
              update.exportResult,
            ))
        ) {
          cancelFinalDraft();
          return;
        }
        finalizingRef.current = true;
        // This is only the browser-unload fallback. The Windows native title-bar
        // close is intercepted and uses the awaited normal commit path below.
        void saveOfficeSessionKeepalive(sessionId, update).catch(
          () => undefined,
        );
      } catch {
        // Never leave a disappearing editor in `created`/`editing`: cancelling
        // preserves the draft Session on disk and releases the Office host.
        cancelFinalDraft();
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
    session?.host,
    session?.mode,
    session?.objectMode,
    session?.exportResult,
    title,
    lines,
    activeLineId,
    latexCodeFormat,
    displayMode,
    objectMode,
    numbered,
    officeFontSizePt,
    dirty,
    autoCommitOnClose,
    currentFingerprint,
    generateSvgExportResult,
    latex,
  ]);

  useEffect(() => {
    if (!session || readyMessageSentRef.current) return;
    readyMessageSentRef.current = true;
    messageOfficeParent({ type: "visualtex-ready", sessionId: session.id });
  }, [session?.id]);

  useEffect(() => {
    if (
      !IS_VSTO_DESKTOP_RUNTIME ||
      !session ||
      !sessionId ||
      loadedSessionIdRef.current !== sessionId ||
      skipAutosaveForSessionRef.current === sessionId ||
      currentFingerprint !== originalFingerprintRef.current ||
      session.status !== "created"
    ) {
      return;
    }
    let cancelled = false;
    let attempt = 0;
    const markEditorReady = () => {
      if (cancelled) return;
      const field = document.querySelector<HTMLElement>("math-field");
      if (!field?.isConnected && attempt < 20) {
        attempt += 1;
        window.setTimeout(markEditorReady, 16);
        return;
      }
      void save({ status: "editing", dirty: false }).catch(() => undefined);
    };
    const frame = window.requestAnimationFrame(markEditorReady);
    return () => {
      cancelled = true;
      window.cancelAnimationFrame(frame);
    };
  }, [
    session?.id,
    session?.status,
    sessionId,
    currentFingerprint,
    save,
  ]);

  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(""), 2200);
    return () => window.clearTimeout(timer);
  }, [toast]);

  useEffect(() => {
    if (IS_VSTO_CONVERT_RUNTIME) return;
    const timer = window.setTimeout(() => {
      void warmupOcrModel(startupOcrModelRef.current).catch(() => undefined);
    }, 300);
    return () => window.clearTimeout(timer);
  }, []);

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

  const handleOcrModelChange = (nextModel: OcrModelName) => {
    if (inlineOcrBusyRef.current || nextModel === ocrModel) return;
    startupOcrModelRef.current = nextModel;
    setOcrModel(nextModel);
    window.localStorage.setItem(OCR_MODEL_STORAGE_KEY, nextModel);
    void warmupOcrModel(nextModel).catch(() => undefined);
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
      const message =
        reason instanceof Error
          ? reason.message
          : typeof reason === "string"
            ? reason
            : "";
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

  const saveCurrentSession = useCallback(
    async (status: "editing" | "committing" | "cancelled") => {
      if (!session) throw new Error("Office Session 尚未加载。");
      const exportResult =
        status === "cancelled"
          ? session.exportResult
          : objectMode === "wordOmml" ||
              objectMode === "mathTypeOle" ||
              (session.host === "powerpoint" && USE_NATIVE_POWERPOINT_COMMIT)
            ? generateSvgExportResult()
            : await generateExportResult();
      if (status === "committing" && !exportResult) {
        throw new Error(isEn ? "Formula export is empty" : "公式导出结果为空");
      }
      const next = await save({
        title,
        lines,
        activeLineId,
        codeFormat: latexCodeFormat,
        displayMode,
        objectMode,
        numbered: displayMode === "block" && numbered,
        fontSizePt: officeFontSizePt,
        dirty,
        status,
        autoCommitOnClose,
        explicitCancel: status === "cancelled",
        exportResult,
        exportWidth: exportResult?.width ?? 0,
        exportHeight: exportResult?.height ?? 0,
        error: null,
      });
      lastSavedFingerprintRef.current = currentFingerprint;
      return next;
    }, [
      session,
      save,
      title,
      lines,
      activeLineId,
      latexCodeFormat,
      displayMode,
      objectMode,
      numbered,
      officeFontSizePt,
      dirty,
      autoCommitOnClose,
      currentFingerprint,
      generateSvgExportResult,
      generateExportResult,
      isEn,
    ],
  );

  const handleCommit = async () => {
    // React state updates do not disable the button until the next render.
    // Keep a synchronous guard as well so a rapid double-click cannot enqueue
    // two commits for the same Office Session.
    if (finalizingRef.current) return;
    historyManager.commitPendingTransaction();
    if (!latex.trim()) {
      setToast(isEn ? "Enter a formula before inserting" : "请输入公式后再插入");
      return;
    }
    finalizingRef.current = true;
    try {
      const next = await saveCurrentSession("committing");

      messageOfficeParent({ type: "visualtex-commit", sessionId: next.id });
      // The parent bridge owns both Word and PowerPoint mutations. Keep the
      // action busy until the host confirms the durable final state; a failed
      // PowerPoint decoration therefore leaves this editor open with a useful
      // error instead of closing after creating an anonymous Graphic shape.
      await waitForOfficeCommitResult(next.id, next.host);
      if (IS_VSTO_DESKTOP_RUNTIME) {
        await closeOfficeSessionWindow(next.id);
        return;
      }
      window.close();
    } catch (error) {
      finalizingRef.current = false;
      setToast(readErrorMessage(
        error,
        isEn ? "Unable to insert the Office formula" : "无法插入 Office 公式",
      ));
    }
  };

  commitFromShortcutRef.current = () => {
    void handleCommit();
  };

  useEffect(() => {
    if (
      !session ||
      loading ||
      error ||
      IS_VSTO_CONVERT_RUNTIME
    ) {
      return;
    }
    return registerOfficeApplyShortcut({
      onApply: () => commitFromShortcutRef.current(),
      isEnabled: () =>
        !ocrOpen &&
        !inlineOcrBusyRef.current &&
        !historyState.isReplaying,
    });
  }, [error, historyState.isReplaying, loading, ocrOpen, session?.id]);

  const handleCancel = async () => {
    finalizingRef.current = true;
    try {
      const next = await saveCurrentSession("cancelled");
      if (IS_VSTO_DESKTOP_RUNTIME) {
        await closeOfficeSessionWindow(next.id);
        return;
      }
      if (next.host === "powerpoint") {
        window.close();
        return;
      }
      const delivered = messageOfficeParent({
        type: "visualtex-cancel",
        sessionId: next.id,
      });
      if (!delivered) window.close();
    } catch (error) {
      finalizingRef.current = false;
      throw error;
    }
  };

  closeFromNativeWindowRef.current = () => {
    if (finalizingRef.current || !session) return;
    if (
      !autoCommitOnClose ||
      (session.mode === "create" && !latex.trim())
    ) {
      void handleCancel().catch((reason) => {
        setToast(readErrorMessage(
          reason,
          isEn ? "Unable to close the Office editor" : "无法关闭 Office 编辑器",
        ));
      });
      return;
    }
    void handleCommit();
  };

  useEffect(() => {
    if (!IS_VSTO_DESKTOP_RUNTIME || !sessionId) return;
    const handleNativeCloseRequest = () => {
      closeFromNativeWindowRef.current();
    };
    window.addEventListener(
      "visualtex-office-close-requested",
      handleNativeCloseRequest,
    );
    return () => {
      window.removeEventListener(
        "visualtex-office-close-requested",
        handleNativeCloseRequest,
      );
    };
  }, [sessionId]);

  const handleCopy = async () => {
    await copyLatex(latex, latexCodeFormat);
    addHistory(latex);
    setToast(isEn ? "LaTeX copied" : "LaTeX 已复制");
  };

  if (
    loading ||
    (session?.host === "powerpoint" &&
      session.mode === "create" &&
      session.status === "created" &&
      !session.dirty &&
      !officePreferencesReady)
  ) {
    return (
      <div className="office-dialog-state">
        <LoaderCircle className="is-spinning" size={28} />
        <strong>{isEn ? "Loading Office Session…" : "正在加载 Office Session…"}</strong>
      </div>
    );
  }

  if (error || !session) {
    return (
      <div className="office-dialog-state is-error">
        <X size={28} />
        <strong>{isEn ? "Unable to open VisualTeX" : "无法打开 VisualTeX"}</strong>
        <p>{error || (isEn ? "Session not found" : "Session 不存在")}</p>
        <button type="button" onClick={() => void reload()}>
          {isEn ? "Retry" : "重新加载"}
        </button>
      </div>
    );
  }

  const officeHeaderLeadingControls = (
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
            onClick={() => setDisplayMode("block")}
            disabled={session.mode === "edit"}
          >
            {isEn ? "Display" : "行间"}
          </button>
        </div>
      ) : null}
      {session.host === "word" &&
      session.mode === "edit" &&
      (session.objectMode === "wordOmml" ||
        session.objectMode === "mathTypeOle" ||
        session.objectMode === "nativeOle") ? (
        <label
          className="office-font-size-setting office-object-mode-setting"
          title={
            session.objectMode === "mathTypeOle"
              ? isEn
                ? "Keep this equation as an editable MathType OLE object, or convert it to VisualTeX OLE"
                : "本次编辑后继续保持可由 MathType 编辑的 OLE，或转换为 VisualTeX OLE"
              : isEn
                ? "Choose whether this edit stays as native Word OMML or becomes a VisualTeX OLE object"
                : "选择本次编辑后继续保持 Word OMML，或转换为 VisualTeX OLE"
          }
        >
          <span>{isEn ? "Save as" : "保存为"}</span>
          <select
            value={objectMode}
            data-office-object-mode
            aria-label={isEn ? "Formula object format" : "公式对象格式"}
            onChange={(event) => {
              const nextMode = event.target.value as OfficeObjectMode;
              setObjectMode(nextMode);
              // Numbering surrounding a MathType source belongs to MathType/Word,
              // not to VisualTeX.  Do not create a second VisualTeX numbering
              // owner during the same edit or conversion.  After conversion, a
              // later VisualTeX edit can opt into VisualTeX numbering normally.
              if (session.objectMode === "mathTypeOle") setNumbered(false);
            }}
          >
            {session.objectMode === "mathTypeOle" ? (
              <option value="mathTypeOle">MathType OLE</option>
            ) : (
              <option value="wordOmml">Word OMML</option>
            )}
            <option value="nativeOle">VisualTeX OLE</option>
          </select>
        </label>
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
        <label
          className="office-auto-commit-setting"
          title={
            session.objectMode === "mathTypeOle"
              ? isEn
                ? "This MathType equation keeps its existing MathType/Word numbering. VisualTeX will not add a second number during this edit."
                : "此 MathType 公式保留现有的 MathType/Word 编号；本次编辑不会再叠加一套 VisualTeX 编号。"
              : undefined
          }
        >
          <input
            type="checkbox"
            checked={session.objectMode === "mathTypeOle" ? false : numbered}
            disabled={session.objectMode === "mathTypeOle"}
            onChange={(event) => setNumbered(event.target.checked)}
          />
          <span>{isEn ? "Number" : "编号"}</span>
        </label>
      ) : null}
    </>
  );

  const officeHeaderTrailingActions = (
    <div
      className="office-inline-history-actions"
      aria-label={isEn ? "History actions" : "历史操作"}
    >
      <button
        type="button"
        className="secondary-button"
        onClick={() => setOcrOpen(true)}
        disabled={inlineOcrIsBusy}
      >
        <ScanLine size={15} />
        <span>{isEn ? "OCR" : "OCR"}</span>
      </button>
      <button
        type="button"
        className="icon-button compact office-history-icon-button"
        data-office-undo-action
        aria-label={isEn ? "Undo" : "撤销"}
        title={isEn ? "Undo" : "撤销"}
        onClick={() => void historyManager.undo()}
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
        onClick={() => void historyManager.redo()}
        disabled={historyBusy || !historyState.canRedo || historyState.isReplaying}
      >
        <Redo2 size={16} strokeWidth={2} />
      </button>
      <span className="office-inline-action-divider" aria-hidden="true" />
      <button
        type="button"
        className="secondary-button office-inline-cancel"
        data-office-cancel-action
        onClick={() => void handleCancel()}
        aria-label={isEn ? "Cancel" : "取消"}
      >
        {isEn ? "Cancel" : "取消"}
      </button>
      <button
        type="button"
        className="primary-button office-inline-primary"
        data-office-primary-action
        onClick={() => void handleCommit()}
        aria-keyshortcuts="Control+S"
        title={isEn ? "Apply and close (Ctrl+S)" : "应用并关闭（Ctrl+S）"}
      >
        {session.mode === "edit"
          ? isEn
            ? "Update"
            : "更新公式"
          : isEn
            ? "Finish and insert"
            : "完成并插入"}
      </button>
    </div>
  );

  return (
    <div className="app-shell office-dialog-shell">
      <EditorWorkspace
        mode={session.mode === "edit" ? "office-edit" : "office-create"}
        showFileActions={false}
        showUpdateActions={false}
        showOfficeActions={false}
        showOcrActions={true}
        officeHeaderLeadingControls={officeHeaderLeadingControls}
        officeHeaderTrailingActions={officeHeaderTrailingActions}
        editorRef={editorRef}
        editorInstanceKey={session.id}
        sidebarOpen={sidebarOpen}
        onSidebarOpenChange={handleSidebarOpenChange}
        onHistoryBusyChange={setHistoryBusy}
        onPasteImage={handleEditorImagePaste}
        onCopy={handleCopy}
        onReplaceDocument={replaceDocumentWithHistory}
        ocrModel={ocrModel}
        ocrModels={OCR_MODELS}
        ocrBusy={inlineOcrIsBusy}
        onOcrModelChange={(model) =>
          handleOcrModelChange(model as OcrModelName)
        }
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
