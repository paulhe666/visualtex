import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { AlertCircle, Check, LoaderCircle, ScanLine, X } from "lucide-react";
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
  copyLatex,
  isLatexCodeFormat,
} from "../../clipboard/LatexCopyService";
import type { LatexCodeFormat } from "../../types/formula";
import type {
  MathEditorHandle,
  MathEditorInsertionTarget,
} from "../../editor/MathEditor";
import { latexToSvg } from "../../export/latexToSvg";
import { latexLinesToOmmlArtifacts } from "../omml/latexToOmml";
import {
  invokeTauri,
  onCurrentTauriWindowCloseRequested,
} from "../shared/tauriTransport";
import {
  cancelMacosOfflineOfficeSession,
  commitMacosOfflineOfficeSession,
  getOfficeSession,
  isMacosOfflineTauriTransport,
  saveOfficeSessionKeepalive,
  type OfficeExportResult,
  type OfficeHost,
} from "../api/sessionClient";
import { useOfficeSession } from "./useOfficeSession";
import { messageOfficeParent } from "./dialogMessages";
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
const OCR_MODEL_STORAGE_KEY = "visualtex.ocr.model";
const USE_NATIVE_POWERPOINT_COMMIT =
  document
    .querySelector<HTMLMetaElement>(
      'meta[name="visualtex-native-powerpoint-commit"]',
    )
    ?.content.toLowerCase() === "true";

const OFFICE_COMMIT_RESULT_TIMEOUT_MS = 45_000;

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
) {
  return JSON.stringify({
    title,
    lines: lines.map((line) => line.latex),
    codeFormat,
    displayMode,
    numbered,
  });
}

export function OfficeDialogApp() {
  const editorRef = useRef<MathEditorHandle>(null);
  const loadedSessionIdRef = useRef("");
  const skipAutosaveForSessionRef = useRef("");
  const lastSavedFingerprintRef = useRef("");
  const readyMessageSentRef = useRef(false);
  const finalizingRef = useRef(false);
  const allowNativeCloseRef = useRef(false);
  const nativeCloseRequestInFlightRef = useRef(false);
  const exportRunIdRef = useRef(0);
  const latestCompleteExportRef = useRef<{
    fingerprint: string;
    exportResult: OfficeExportResult;
  } | null>(null);
  const [sidebarOpen, setSidebarOpen] = useState(() => window.innerWidth >= 1040);
  const [historyBusy, setHistoryBusy] = useState(false);
  const [autoCommitOnClose, setAutoCommitOnClose] = useState(true);
  const [displayMode, setDisplayMode] = useState<"inline" | "block">("inline");
  const [numbered, setNumbered] = useState(false);
  const [toast, setToast] = useState("");
  const [ocrOpen, setOcrOpen] = useState(false);
  const [ocrModel, setOcrModel] = useState<OcrModelName>(() => {
    const stored = window.localStorage.getItem(OCR_MODEL_STORAGE_KEY);
    return OCR_MODELS.some((item) => item.id === stored)
      ? (stored as OcrModelName)
      : DEFAULT_OCR_MODEL;
  });
  const [inlineOcr, setInlineOcr] = useState<InlineOcrState | null>(null);
  const inlineOcrBusyRef = useRef(false);
  const inlineOcrCancelRequestedRef = useRef(false);
  const inlineOcrRunIdRef = useRef(0);
  const inlineOcrClearTimerRef = useRef<number | null>(null);
  const ocrPrewarmStartedRef = useRef(false);
  const { sessionId, session, loading, error, save } = useOfficeSession();

  useEffect(() => {
    if (loading) {
      document.title = "VisualTeX Office Formula — 正在加载";
      return;
    }
    if (error || !session) {
      document.title = `VisualTeX Office Formula — 加载失败${error ? `：${error.slice(0, 80)}` : ""}`;
      return;
    }
    document.title = `VisualTeX Office Formula — ${session.host === "word" ? "Word" : "PowerPoint"} 已就绪`;
  }, [loading, error, session?.id, session?.host]);

  const title = useEditorStore((state) => state.title);
  const lines = useEditorStore((state) => state.lines);
  const activeLineId = useEditorStore((state) => state.activeLineId);
  const language = useEditorStore((state) => state.language);
  const latexCodeFormat = useEditorStore((state) => state.latexCodeFormat);
  const addHistory = useEditorStore((state) => state.addHistory);
  const historyState = useHistorySnapshot();
  const isEn = language === "en";
  const latex = joinFormulaLines(lines);
  const selectedOcrModel =
    OCR_MODELS.find((item) => item.id === ocrModel) ?? OCR_MODELS[1];
  const inlineOcrModel =
    OCR_MODELS.find((item) => item.id === inlineOcr?.model) ?? selectedOcrModel;
  const inlineOcrIsBusy =
    inlineOcr?.status === "running" || inlineOcr?.status === "cancelling";

  const originalFingerprint = useMemo(() => {
    if (!session) return "";
    return documentFingerprint(
      session.originalMetadata?.title ?? session.title,
      session.originalMetadata?.lines ?? session.lines,
      session.originalMetadata?.codeFormat ?? session.codeFormat,
      session.originalMetadata?.displayMode ?? session.displayMode,
      session.originalMetadata?.numbered ?? session.numbered ?? false,
    );
  }, [session?.id]);

  const currentFingerprint = useMemo(
    () => documentFingerprint(title, lines, latexCodeFormat, displayMode, numbered),
    [title, lines, latexCodeFormat, displayMode, numbered],
  );
  const dirty = Boolean(session) && currentFingerprint !== originalFingerprint;

  useEffect(() => {
    if (!session || loadedSessionIdRef.current === session.id) return;
    loadedSessionIdRef.current = session.id;
    skipAutosaveForSessionRef.current = session.id;
    const nextLines = session.lines.length
      ? session.lines
      : [{ id: crypto.randomUUID(), latex: "" }];
    useEditorStore.getState().replaceDocumentState({
      title: session.title || (isEn ? "Office Formula" : "Office 公式"),
      lines: nextLines,
      activeLineId:
        session.activeLineId &&
        nextLines.some((line) => line.id === session.activeLineId)
          ? session.activeLineId
          : nextLines[0]?.id ?? null,
      selectionByLineId: {},
    });
    if (isLatexCodeFormat(session.codeFormat)) {
      useEditorStore
        .getState()
        .setLatexCodeFormat(session.codeFormat as LatexCodeFormat);
    }
    setAutoCommitOnClose(session.autoCommitOnClose);
    setDisplayMode(session.displayMode);
    setNumbered(session.displayMode === "block" && Boolean(session.numbered));
    const loadedFingerprint = documentFingerprint(
      session.title,
      nextLines,
      session.codeFormat,
      session.displayMode,
      session.displayMode === "block" && Boolean(session.numbered),
    );
    lastSavedFingerprintRef.current = loadedFingerprint;
    latestCompleteExportRef.current = session.exportResult?.pngBase64
      ? { fingerprint: loadedFingerprint, exportResult: session.exportResult }
      : null;
  }, [session?.id, isEn]);

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

  const generateSvgExportResult = useCallback((): OfficeExportResult | null => {
    if (!latex.trim()) return null;
    const svg = latexToSvg(latex, {
      displayMode: displayMode === "block",
      fontSizePt: 14,
      paddingPx: displayMode === "inline" ? 1 : 10,
      background: "transparent",
    });
    const wordArtifacts =
      session?.host === "word"
        ? latexLinesToOmmlArtifacts(
            lines.map((line) => line.latex),
            displayMode,
          )
        : null;
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
  }, [latex, displayMode, lines, session?.host]);

  const generateExportResult = useCallback(async (): Promise<OfficeExportResult | null> => {
    const base = generateSvgExportResult();
    if (!base) return null;
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
  }, [generateSvgExportResult]);

  useEffect(() => {
    if (!session || !sessionId || finalizingRef.current) return;
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
        numbered: displayMode === "block" && numbered,
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
    numbered,
    dirty,
    autoCommitOnClose,
    save,
    isEn,
    generateSvgExportResult,
    generateExportResult,
  ]);

  useEffect(() => {
    if (!sessionId) return;
    const finalDraftUpdate = (status: "editing" | "committing") => {
      const cached = latestCompleteExportRef.current;
      const exportResult =
        cached?.fingerprint === currentFingerprint
          ? cached.exportResult
          : generateSvgExportResult();
      return {
        title,
        lines,
        activeLineId,
        codeFormat: latexCodeFormat,
        displayMode,
        numbered: displayMode === "block" && numbered,
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
    session?.host,
    title,
    lines,
    activeLineId,
    latexCodeFormat,
    displayMode,
    numbered,
    dirty,
    autoCommitOnClose,
    currentFingerprint,
    generateSvgExportResult,
    latex,
  ]);

  useEffect(() => {
    if (
      !session ||
      readyMessageSentRef.current ||
      isMacosOfflineTauriTransport()
    ) {
      return;
    }
    readyMessageSentRef.current = true;
    messageOfficeParent({ type: "visualtex-ready", sessionId: session.id });
  }, [session?.id]);

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
    let cancelled = false;
    const delay = ocrPrewarmStartedRef.current ? 250 : 500;
    const timer = window.setTimeout(() => {
      ocrPrewarmStartedRef.current = true;
      void getOcrRuntimeStatus()
        .then((runtime) => {
          if (cancelled || !runtime.installed) return;
          const availableModel = resolveAvailableOcrModel(runtime, ocrModel);
          if (availableModel !== ocrModel) {
            setOcrModel(availableModel);
            window.localStorage.setItem(OCR_MODEL_STORAGE_KEY, availableModel);
          }
          return prewarmOcrModel(availableModel);
        })
        .catch(() => undefined);
    }, delay);

    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [ocrModel]);

  const handleOcrModelChange = (nextModel: OcrModelName) => {
    if (inlineOcrBusyRef.current || nextModel === ocrModel) return;
    setOcrModel(nextModel);
    window.localStorage.setItem(OCR_MODEL_STORAGE_KEY, nextModel);
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

      const availableOcrModel = resolveAvailableOcrModel(runtime, ocrModel);
      if (availableOcrModel !== ocrModel) {
        setOcrModel(availableOcrModel);
        window.localStorage.setItem(OCR_MODEL_STORAGE_KEY, availableOcrModel);
      }

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
          : session.host === "powerpoint" && USE_NATIVE_POWERPOINT_COMMIT
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
        numbered: displayMode === "block" && numbered,
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
      numbered,
      dirty,
      autoCommitOnClose,
      currentFingerprint,
      generateSvgExportResult,
      generateExportResult,
      isEn,
    ],
  );

  const closeOfficeEditorWindow = useCallback(async () => {
    if (!isMacosOfflineTauriTransport()) {
      window.close();
      return;
    }

    allowNativeCloseRef.current = true;
    try {
      await invokeTauri<void>("close_macos_offline_office_editor_window");
    } catch (error) {
      allowNativeCloseRef.current = false;
      throw error;
    }
  }, []);

  const handleCommit = useCallback(async () => {
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

      if (isMacosOfflineTauriTransport()) {
        await commitMacosOfflineOfficeSession(next.id);
        try {
          await closeOfficeEditorWindow();
        } catch (closeError) {
          finalizingRef.current = false;
          const detail =
            closeError instanceof Error ? closeError.message : String(closeError);
          setToast(
            isEn
              ? `The formula was inserted, but the editor could not close: ${detail}`
              : `公式已经插入，但编辑窗口无法自动关闭：${detail}`,
          );
        }
        return;
      }

      messageOfficeParent({ type: "visualtex-commit", sessionId: next.id });
      // The parent bridge owns both Word and PowerPoint mutations. Keep the
      // action busy until the host confirms the durable final state; a failed
      // PowerPoint decoration therefore leaves this editor open with a useful
      // error instead of closing after creating an anonymous Graphic shape.
      await waitForOfficeCommitResult(next.id, next.host);
      window.close();
    } catch (error) {
      finalizingRef.current = false;
      const message =
        error instanceof Error
          ? error.message
          : isEn
            ? "Unable to insert the Office formula"
            : "无法插入 Office 公式";
      setToast(message);
    }
  }, [closeOfficeEditorWindow, isEn, latex, saveCurrentSession]);

  const handleCancel = useCallback(async () => {
    if (finalizingRef.current) return;
    finalizingRef.current = true;
    try {
      const next = await saveCurrentSession("cancelled");
      if (isMacosOfflineTauriTransport()) {
        await cancelMacosOfflineOfficeSession(next.id);
        await closeOfficeEditorWindow();
        return;
      }
      if (next.host === "powerpoint") {
        window.close();
        return;
      }
      messageOfficeParent({ type: "visualtex-cancel", sessionId: next.id });
    } catch (error) {
      finalizingRef.current = false;
      const message =
        error instanceof Error
          ? error.message
          : isEn
            ? "Unable to cancel the Office formula"
            : "无法取消 Office 公式";
      setToast(message);
    }
  }, [closeOfficeEditorWindow, isEn, saveCurrentSession]);

  useEffect(() => {
    if (!isMacosOfflineTauriTransport() || !sessionId) return;

    let disposed = false;
    let unlisten: (() => void) | undefined;
    void onCurrentTauriWindowCloseRequested((event) => {
      if (allowNativeCloseRef.current || disposed) return;
      event.preventDefault();
      if (nativeCloseRequestInFlightRef.current) return;
      nativeCloseRequestInFlightRef.current = true;

      const finalize = latex.trim() && autoCommitOnClose
        ? handleCommit()
        : handleCancel();
      void finalize.finally(() => {
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
        const message =
          reason instanceof Error ? reason.message : String(reason);
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
    await copyLatex(latex, latexCodeFormat);
    addHistory(latex);
    setToast(isEn ? "LaTeX copied" : "LaTeX 已复制");
  };

  if (loading) {
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
      </div>
    );
  }

  return (
    <div className="app-shell office-dialog-shell">
      <header className="office-dialog-header">
        <div>
          <strong>VisualTeX</strong>
          <span>
            {session.host === "word" ? "Microsoft Word" : "Microsoft PowerPoint"}
          </span>
        </div>
        <div className="office-dialog-options">
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
          {session.host === "word" && displayMode === "block" ? (
            <label className="office-auto-commit-setting">
              <input
                type="checkbox"
                checked={numbered}
                onChange={(event) => setNumbered(event.target.checked)}
                disabled={session.mode === "edit"}
              />
              <span>{isEn ? "Add equation number" : "添加公式编号"}</span>
            </label>
          ) : null}
          <label className="office-auto-commit-setting">
            <input
              type="checkbox"
              checked={autoCommitOnClose}
              onChange={(event) => setAutoCommitOnClose(event.target.checked)}
            />
            <span>
              {isEn ? "Apply when the window closes" : "关闭窗口时自动应用"}
            </span>
          </label>
        </div>
        <div className="office-history-actions" aria-label={isEn ? "History actions" : "历史操作"}>
          <button
            type="button"
            className="secondary-button"
            onClick={() => setOcrOpen(true)}
            disabled={inlineOcrIsBusy}
          >
            <ScanLine size={15} />
            {isEn ? "Image OCR" : "图片 OCR"}
          </button>
          <button
            type="button"
            className="secondary-button"
            onClick={() => void historyManager.undo()}
            disabled={historyBusy || !historyState.canUndo || historyState.isReplaying}
          >
            {isEn ? "Undo" : "撤销"}
          </button>
          <button
            type="button"
            className="secondary-button"
            onClick={() => void historyManager.redo()}
            disabled={historyBusy || !historyState.canRedo || historyState.isReplaying}
          >
            {isEn ? "Redo" : "重做"}
          </button>
        </div>
      </header>

      <EditorWorkspace
        mode={session.mode === "edit" ? "office-edit" : "office-create"}
        showFileActions={false}
        showUpdateActions={false}
        showOfficeActions
        showOcrActions={true}
        primaryActionLabel={
          session.mode === "edit"
            ? isEn
              ? "Update formula"
              : "更新公式"
            : isEn
              ? "Finish and insert"
              : "完成并插入"
        }
        onPrimaryAction={handleCommit}
        onCancel={handleCancel}
        editorRef={editorRef}
        sidebarOpen={sidebarOpen}
        onSidebarOpenChange={setSidebarOpen}
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
