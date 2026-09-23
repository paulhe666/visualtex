import {
  forwardRef,
  useEffect,
  useImperativeHandle,
  useMemo,
  useRef,
  useState,
  type ReactNode,
  type CSSProperties,
} from "react";
import { EngineEditorHost } from "../../../../engine/web/host/editor-host.ts";
import { EngineEditorSessionOwner } from "../../../../engine/web/host/editor-session.ts";
import type { OcrInsertionTicket } from "../../../../engine/web/host/editor-session.ts";
import type { CommandInputSnapshot } from "../../../../engine/web/host/command-input.ts";
import type { EditorSnapshot } from "../../../../engine/web/host/editing.ts";
import type { LayoutOptions } from "../../../../engine/web/host/layout.ts";
import type { CoreEditorExportPort } from "../engine/coreEditorExportPort";
import {
  buildCoreProductCommandDescriptors,
  coreCommandUsage,
  coreEditCommandForProductCommand,
  coreInlineShortcuts,
} from "../engine/coreCommandAdapter";
import {
  getCoreEngineRuntime,
  isCoreRuntimeFault,
  resolveCoreLayoutForPreferences,
  resolveCoreFontPreferences,
  type CoreFontResolution,
} from "../engine/coreEngineRuntime";
import {
  createFormulaLine,
  useEditorStore,
} from "../stores/editorStore";
import { getRuntimeCommandRegistry } from "../autocomplete/runtimeCommandRegistry";
import { CORE_PREVIEW_RETRY_EVENT, retryCoreStaticPreviews } from "../engine/coreStaticPreview";
import { CommandSuggestionPopup } from "../autocomplete/CommandSuggestionPopup";
import { readCustomSymbolLibrary } from "../math/customSymbolRegistry";
import type { FormulaAlignment, FormulaLine, LatexCodeFormat } from "../types/formula";
import type {
  AddLineEntry,
  CoreMathSelectionSnapshot,
  MathSelectionSnapshot,
  RemoveLineEntry,
} from "../history/historyTypes";
import { historyManager } from "../history/HistoryManager";
import { getEditorDocumentSnapshot } from "../history/documentHistory";
import type {
  MathEditorFocusOptions,
  MathEditorHandle,
  MathEditorHistoryState,
  MathEditorInsertionTarget,
  MathEditorSelectionStyle,
  MathEditorSelectionTarget,
} from "./mathEditorTypes";
import "./coreMathEditor.css";

interface Props {
  lines: FormulaLine[];
  activeLineId: string | null;
  formulaAlignment: FormulaAlignment;
  latexCodeFormat: LatexCodeFormat;
  zoom: number;
  readOnly?: boolean;
  previewOnly?: boolean;
  onPreviewActivate?: () => void;
  draftError?: string;
  onPasteImage?: (file: File, target: MathEditorInsertionTarget) => void;
  onCopyPng?: () => Promise<void>;
  onHistoryBusyChange?: (busy: boolean) => void;
  overlay?: ReactNode;
  onInitialReject: (reason: string) => void;
  onCoreHistoryStateChange?: (state: MathEditorHistoryState | null) => void;
}

type HostRecord = {
  owner: EngineEditorSessionOwner;
  host: EngineEditorHost;
  mount: HTMLDivElement;
  layout: LayoutOptions;
  focusListener: () => void;
  keyListener: (event: KeyboardEvent) => void;
};

type CoreCommandPopupState = {
  lineId: string;
  snapshot: CommandInputSnapshot;
  position: { left: number; top: number };
};

export interface CoreEditorProbeResult {
  eligible: boolean;
  reason?: string;
}

export async function probeCoreEditorSession(
  lines: readonly FormulaLine[],
): Promise<CoreEditorProbeResult> {
  if (!lines.length) return { eligible: false, reason: "A formula document requires at least one line." };
  const settings = useEditorStore.getState();
  try {
    const context = await getCoreEngineRuntime();
    const styles = new Set(lines.map((line) => line.mode === "inline" ? "text" as const : "display" as const));
    for (const style of styles) {
      const layout = await resolveCoreLayoutForPreferences(
        context,
        settings.formulaLetterFont,
        settings.formulaChineseFont,
        style,
      );
      if (!layout) {
        return {
          eligible: false,
          reason: "The selected formula font preferences cannot be resolved to verified core font resources.",
        };
      }
    }
    return { eligible: true };
  } catch (error) {
    return {
      eligible: false,
      reason: error instanceof Error ? error.message : String(error),
    };
  }
}

function hexColor(value: string) {
  const match = /^#([0-9a-f]{6})$/iu.exec(value.trim());
  if (!match) return null;
  const number = Number.parseInt(match[1], 16);
  return {
    r: (number >>> 16) & 0xff,
    g: (number >>> 8) & 0xff,
    b: number & 0xff,
  };
}

export const CoreMathEditor = forwardRef<MathEditorHandle, Props>(
  function CoreMathEditor(
    {
      lines,
      activeLineId,
      formulaAlignment,
      latexCodeFormat,
      draftError,
      zoom,
      readOnly = false,
      previewOnly = false,
      onPreviewActivate,
      onPasteImage,
      overlay,
      onInitialReject,
      onCoreHistoryStateChange,
    },
    ref,
  ) {
    const showLineNumbers = useEditorStore((state) => state.showLineNumbers);
    const formulaInsetLeft = useEditorStore((state) => state.formulaInsetLeft);
    const formulaInsetRight = useEditorStore((state) => state.formulaInsetRight);
    const formulaRowVerticalInset = useEditorStore((state) => state.formulaRowVerticalInset);
    const inputBehavior = useEditorStore((state) => state.inputBehavior);
    const personalize = useEditorStore((state) => state.personalize);
    const suggestionCount = useEditorStore((state) => state.suggestionCount);
    const usage = useEditorStore((state) => state.usage);
    const language = useEditorStore((state) => state.language);
    const autoPairDelimiters = useEditorStore((state) => state.autoPairDelimiters);
    const formulaLetterFont = useEditorStore((state) => state.formulaLetterFont);
    const formulaChineseFont = useEditorStore((state) => state.formulaChineseFont);
    const recordCommand = useEditorStore((state) => state.recordCommand);
    const [frozenSymbols] = useState(() => structuredClone(readCustomSymbolLibrary().symbols));
    const [frozenRegistry] = useState(() => structuredClone(getRuntimeCommandRegistry()));
    const commandDescriptors = useMemo(
      () => buildCoreProductCommandDescriptors(inputBehavior, language, frozenRegistry),
      [inputBehavior, language, frozenRegistry],
    );
    const commandUsage = useMemo(
      () => coreCommandUsage(usage, personalize),
      [usage, personalize],
    );
    const inlineShortcuts = useMemo(
      () => coreInlineShortcuts(inputBehavior.autoEscapeShortcuts),
      [inputBehavior.autoEscapeShortcuts],
    );
    const commandById = useMemo(
      () => new Map(frozenRegistry.map((command) => [command.id, command] as const)),
      [frozenRegistry],
    );
    const [accepted, setAccepted] = useState(false);
    const [retryGeneration, setRetryGeneration] = useState(0);
    const fontRecoveryAttemptsRef = useRef(0);
    const [fontResolution, setFontResolution] = useState<CoreFontResolution | null>(null);
    const [rowErrors, setRowErrors] = useState<Record<string, string>>({});
    const editorLayout = (layout: LayoutOptions): LayoutOptions => ({ ...layout, fontSizePx: 54, paddingEm: 0 });
    const [commandPopup, setCommandPopup] = useState<CoreCommandPopupState | null>(null);
    const [probeError, setProbeError] = useState<string | null>(null);
    const [layoutNotice, setLayoutNotice] = useState<string | null>(null);
    const [runtimeRecoveryGeneration, setRuntimeRecoveryGeneration] = useState(0);
    const runtimeRecoveryScheduledRef = useRef(false);
    const hostRecords = useRef(new Map<string, HostRecord>());
    const mountRefs = useRef(new Map<string, HTMLDivElement>());
    const initialLinesRef = useRef(lines.map((line) => ({ ...line })));
    const initialFontPreferencesRef = useRef({
      formulaLetterFont,
      formulaChineseFont,
    });
    const initialLayoutRef = useRef<LayoutOptions | null>(null);
    const rejectedRef = useRef(false);
    const onInitialRejectRef = useRef(onInitialReject);
    onInitialRejectRef.current = onInitialReject;

    useEffect(() => {
      const retry = () => setRetryGeneration((value) => value + 1);
      window.addEventListener(CORE_PREVIEW_RETRY_EVENT, retry);
      return () => window.removeEventListener(CORE_PREVIEW_RETRY_EVENT, retry);
    }, []);

    const activeRecord = () => {
      const id = useEditorStore.getState().activeLineId ?? activeLineId;
      if (id) {
        const record = hostRecords.current.get(id);
        if (record) return record;
      }
      return hostRecords.current.values().next().value as HostRecord | undefined;
    };

    const focusHostSoon = (lineId: string) => {
      window.setTimeout(() => hostRecords.current.get(lineId)?.host.focus(), 0);
    };

    const setFormulaLineMode = (lineId: string, mode: "inline" | "display") => {
      if (readOnly || previewOnly || latexCodeFormat !== "mixed-inline-display") return;
      const state = useEditorStore.getState();
      const line = state.lines.find((item) => item.id === lineId);
      if (!line || (line.mode ?? "display") === mode) return;
      historyManager.commitPendingTransaction();
      const before = getEditorDocumentSnapshot({});
      const after = { ...before, lines: before.lines.map((item) => item.id === lineId ? { ...item, mode } : { ...item }), activeLineId: lineId };
      state.replaceDocumentState(after);
      historyManager.push({ type: "replace-document", before, after, source: "source-apply", timestamp: Date.now() });
      focusHostSoon(lineId);
    };

    const addOuterLineAfter = (lineId?: string) => {
      historyManager.commitPendingTransaction();
      const state = useEditorStore.getState();
      const anchorIndex = lineId
        ? state.lines.findIndex((line) => line.id === lineId)
        : state.lines.length - 1;
      const safeAnchorIndex = Math.max(0, anchorIndex);
      const nextIndex = Math.min(safeAnchorIndex + 1, state.lines.length);
      const inheritedMode = state.lines[safeAnchorIndex]?.mode ?? "display";
      const line = createFormulaLine("", undefined, inheritedMode);
      const beforeActiveLineId = state.activeLineId;
      state.insertFormulaLine(line, nextIndex);
      useEditorStore.getState().setActiveLineId(line.id);
      const entry: AddLineEntry = {
        type: "add-line",
        line,
        index: nextIndex,
        beforeActiveLineId,
        afterActiveLineId: line.id,
        beforeSelection: null,
        afterSelection: null,
        timestamp: Date.now(),
      };
      historyManager.push(entry);
      focusHostSoon(line.id);
    };

    const removeEmptyOuterLine = (lineId: string) => {
      historyManager.commitPendingTransaction();
      const state = useEditorStore.getState();
      if (state.lines.length <= 1) return false;
      const index = state.lines.findIndex((line) => line.id === lineId);
      const line = index >= 0 ? state.lines[index] : undefined;
      if (!line || line.latex.trim()) return false;
      const beforeActiveLineId = state.activeLineId;
      const nextActiveLineId = state.lines[index - 1]?.id ?? state.lines[index + 1]?.id ?? null;
      state.removeFormulaLine(lineId);
      useEditorStore.getState().setActiveLineId(nextActiveLineId);
      const entry: RemoveLineEntry = {
        type: "remove-line",
        line: { ...line },
        index,
        beforeActiveLineId,
        afterActiveLineId: nextActiveLineId,
        beforeSelection: null,
        afterSelection: null,
        timestamp: Date.now(),
      };
      historyManager.push(entry);
      if (nextActiveLineId) focusHostSoon(nextActiveLineId);
      return true;
    };

    const scheduleRuntimeRecovery = (error: unknown) => {
      if (!isCoreRuntimeFault(error) || runtimeRecoveryScheduledRef.current) return false;
      runtimeRecoveryScheduledRef.current = true;
      setLayoutNotice(
        "核心运行实例异常，正在从最后一次已提交源码恢复；恢复后本公式的内核撤销历史将从当前检查点重新开始。",
      );
      window.setTimeout(() => {
        for (const record of hostRecords.current.values()) {
          record.mount.removeEventListener("focusin", record.focusListener);
          record.mount.removeEventListener("keydown", record.keyListener, true);
          try {
            record.host.dispose();
          } catch {
            // A faulted runtime may reject pending OCR teardown.
          }
          try {
            if (!record.owner.isDestroyed) record.owner.destroy();
          } catch {
            // The old runtime is already invalid; local host retirement is enough.
          }
        }
        hostRecords.current.clear();
        initialLayoutRef.current = null;
        onCoreHistoryStateChange?.({ engine: "core", canUndo: false, canRedo: false });
        setRuntimeRecoveryGeneration((value) => value + 1);
      }, 0);
      return true;
    };

    useEffect(() => {
      if (accepted) return;
      let cancelled = false;
      setProbeError(null);
      const probe = async () => {
        const context = await getCoreEngineRuntime();
        const style = initialLinesRef.current[0]?.mode === "inline" ? "text" : "display";
        let resolution: CoreFontResolution | null = null;
        let layout: LayoutOptions | null = null;
        try {
          resolution = await resolveCoreFontPreferences(
            context,
            formulaLetterFont,
            formulaChineseFont,
            style,
          );
          layout = resolution.layout;
        } catch (error) {
          if (isCoreRuntimeFault(error)) throw error;
          console.warn("VisualTeX exact formula-font plan is unavailable; using the bundled safety layout:", error);
          layout = await resolveCoreLayoutForPreferences(
            context,
            formulaLetterFont,
            formulaChineseFont,
            style,
          );
        }
        if (!layout) throw new Error("VisualTeX core layout resources are unavailable.");
        if (cancelled) return;
        initialFontPreferencesRef.current = { formulaLetterFont, formulaChineseFont };
        initialLayoutRef.current = editorLayout(layout);
        rejectedRef.current = false;
        setFontResolution(resolution);
        setAccepted(true);
        if (!resolution && fontRecoveryAttemptsRef.current < 2) {
          fontRecoveryAttemptsRef.current += 1;
          window.setTimeout(() => retryCoreStaticPreviews(), 180 * fontRecoveryAttemptsRef.current);
        } else if (resolution) {
          fontRecoveryAttemptsRef.current = 0;
        }
      };
      void probe().catch((error) => {
        if (cancelled) return;
        const reason = error instanceof Error ? error.message : String(error);
        setProbeError(reason);
        onInitialRejectRef.current(reason);
      });
      return () => { cancelled = true; };
    }, [accepted, formulaLetterFont, formulaChineseFont, retryGeneration]);

    useEffect(() => {
      if (!accepted) return;
      let cancelled = false;
      const createHosts = async () => {
        const context = await getCoreEngineRuntime();
        if (cancelled) return;
        const checkpointLines = runtimeRecoveryGeneration > 0
          ? useEditorStore.getState().lines
          : lines;
        for (const renderedLine of lines) {
          const line = checkpointLines.find((item) => item.id === renderedLine.id) ?? renderedLine;
          if (hostRecords.current.has(line.id)) continue;
          const mount = mountRefs.current.get(line.id);
          if (!mount) continue;
          const layout = initialLayoutRef.current
            ? { ...structuredClone(initialLayoutRef.current), style: line.mode === "inline" ? ("text" as const) : ("display" as const) }
            : await resolveCoreLayoutForPreferences(
                context,
                initialFontPreferencesRef.current.formulaLetterFont,
                initialFontPreferencesRef.current.formulaChineseFont,
                line.mode === "inline" ? "text" : "display",
              );
          if (cancelled || hostRecords.current.has(line.id) || !mount.isConnected) return;
          if (!layout) throw new Error("The initial verified core layout is unavailable.");
          const owner = new EngineEditorSessionOwner(
            context.runtime,
            `desktop-core:${line.id}`,
            line.latex,
            frozenSymbols,
          );
          let host: EngineEditorHost;
          host = new EngineEditorHost(
            context.runtime,
            mount,
            {
              documentId: owner.documentId,
              layout: editorLayout(layout),
              commands: commandDescriptors,
              commandUsage,
              commandLimit: suggestionCount,
              externalCommandPopup: true,
              onCommandState: (snapshot) => {
                if (!snapshot.active) {
                  setCommandPopup((current) => current?.lineId === line.id ? null : current);
                  return;
                }
                const anchor = host.input.textarea.getBoundingClientRect();
                setCommandPopup({
                  lineId: line.id,
                  snapshot,
                  position: {
                    left: Math.max(8, anchor.left + window.scrollX),
                    top: Math.max(8, anchor.bottom + 4 + window.scrollY),
                  },
                });
              },
              inlineShortcuts,
              inputBehavior: {
                autoPairDelimiters,
                autoExitSuperscript: inputBehavior.autoExitSuperscript,
                autoExitSubscript: inputBehavior.autoExitSubscript,
                autoExitAccent: inputBehavior.autoExitAccent,
                autoExitWrapperCommand: inputBehavior.autoExitWrapperCommand,
              },
              onCommandAccepted: (id, query) => recordCommand(id, query, "candidate"),
              onEnter: () => addOuterLineAfter(line.id),
              onLayout: (snapshot) => {
                // Preserve the pre-core VisualTeX row geometry: the product
                // shell owns vertical insets and separators, while the core
                // reports only the actual formula scene height.
                const row = mount.closest<HTMLElement>(".formula-line");
                if (!row) return;
                const verticalInset = Math.max(0, Math.min(24, Math.round(formulaRowVerticalInset)));
                const legacySimpleFloor = 54 * zoom * 1.12 + verticalInset * 2;
                const sceneHeight = snapshot.height * zoom + verticalInset * 2;
                const rowHeight = Math.ceil(Math.max(30, legacySimpleFloor, sceneHeight));
                row.style.setProperty("--formula-row-height", `${rowHeight}px`);
                mount.style.setProperty("--core-row-min-height", `${rowHeight}px`);
              },
              onChange: (snapshot) => {
                setRowErrors((current) => {
                  if (!(line.id in current)) return current;
                  const next = { ...current }; delete next[line.id]; return next;
                });
                onCoreHistoryStateChange?.({
                  engine: "core",
                  canUndo: snapshot.undoDepth > 0,
                  canRedo: snapshot.redoDepth > 0,
                });
                const state = useEditorStore.getState();
                const current = state.lines.find((item) => item.id === line.id);
                if (current && current.latex !== snapshot.document.source) {
                  state.replaceFormulaLineExact(line.id, snapshot.document.source);
                }
              },
              onError: (error) => {
                console.error("VisualTeX core editor:", error);
                setRowErrors((current) => ({ ...current, [line.id]: String(error.message ?? error) }));
                scheduleRuntimeRecovery(error);
              },
            },
            owner,
          );
          host.setZoom(zoom);
          const focusListener = () => {
            useEditorStore.getState().setActiveLineId(line.id);
            const snapshot = owner.current();
            onCoreHistoryStateChange?.({
              engine: "core",
              canUndo: snapshot.undoDepth > 0,
              canRedo: snapshot.redoDepth > 0,
            });
          };
          const keyListener = (event: KeyboardEvent) => {
            if (
              event.key !== "Backspace" ||
              event.altKey || event.ctrlKey || event.metaKey || event.shiftKey ||
              owner.current().document.source.trim()
            ) return;
            if (!removeEmptyOuterLine(line.id)) return;
            event.preventDefault();
            event.stopImmediatePropagation();
          };
          mount.addEventListener("focusin", focusListener);
          mount.addEventListener("keydown", keyListener, true);
          const record: HostRecord = {
            owner,
            host,
            mount,
            layout: editorLayout(layout),
            focusListener,
            keyListener,
          };
          hostRecords.current.set(line.id, record);
        }
        if (runtimeRecoveryGeneration > 0 && !cancelled) {
          runtimeRecoveryScheduledRef.current = false;
          setLayoutNotice(null);
        }
      };
      createHosts().catch((error) => {
        if (cancelled || rejectedRef.current) return;
        if (scheduleRuntimeRecovery(error)) return;
        rejectedRef.current = true;
        onInitialRejectRef.current(error instanceof Error ? error.message : String(error));
      });
      return () => {
        cancelled = true;
      };
    }, [accepted, lines, runtimeRecoveryGeneration, zoom]);

    useEffect(() => {
      if (!accepted) return;
      let cancelled = false;
      const updateFonts = async () => {
        const context = await getCoreEngineRuntime();
        let unresolved = false;
        for (const [lineId, record] of hostRecords.current) {
          const line = lines.find((item) => item.id === lineId);
          if (!line) continue;
          const style = line.mode === "inline" ? "text" : "display";
          let resolution: CoreFontResolution | null = null;
          let nextLayout: LayoutOptions | null = null;
          try {
            resolution = await resolveCoreFontPreferences(
              context,
              formulaLetterFont,
              formulaChineseFont,
              style,
            );
            nextLayout = resolution.layout;
          } catch (error) {
            if (isCoreRuntimeFault(error)) throw error;
            console.warn("VisualTeX exact formula-font plan is unavailable during a font switch:", error);
            nextLayout = await resolveCoreLayoutForPreferences(
              context,
              formulaLetterFont,
              formulaChineseFont,
              style,
            );
          }
          if (cancelled) return;
          if (!nextLayout) {
            unresolved = true;
            continue;
          }
          const layout = editorLayout(nextLayout);
          try {
            record.host.setLayoutOptions(layout);
            record.layout = structuredClone(layout);
            initialLayoutRef.current = structuredClone(layout);
            initialFontPreferencesRef.current = { formulaLetterFont, formulaChineseFont };
            if (resolution) {
              setFontResolution(resolution);
              fontRecoveryAttemptsRef.current = 0;
            }
          } catch (error) {
            unresolved = true;
            console.warn("VisualTeX core kept the previous verified font layout:", error);
          }
        }
        if (!cancelled) {
          setLayoutNotice(
            unresolved
              ? "当前字体偏好无法解析为已验证字体资源；本公式会话继续使用上一个已验证布局。"
              : null,
          );
        }
      };
      void updateFonts().catch((error) => {
        if (!cancelled && !scheduleRuntimeRecovery(error)) setLayoutNotice(`字体资源暂时不可用，公式源码与当前会话保留。${String(error)}`);
      });
      return () => {
        cancelled = true;
      };
    }, [accepted, formulaChineseFont, formulaLetterFont, lines, retryGeneration]);

    useEffect(() => {
      if (!accepted) return;
      for (const record of hostRecords.current.values()) {
        record.host.setInlineShortcuts(inlineShortcuts);
        record.host.setInputBehavior({
          autoPairDelimiters,
          autoExitSuperscript: inputBehavior.autoExitSuperscript,
          autoExitSubscript: inputBehavior.autoExitSubscript,
          autoExitAccent: inputBehavior.autoExitAccent,
          autoExitWrapperCommand: inputBehavior.autoExitWrapperCommand,
        });
        record.host.setCommandConfiguration(commandDescriptors, commandUsage, suggestionCount);
      }
    }, [
      accepted,
      autoPairDelimiters,
      commandDescriptors,
      commandUsage,
      inputBehavior.autoExitAccent,
      inputBehavior.autoExitSubscript,
      inputBehavior.autoExitSuperscript,
      inputBehavior.autoExitWrapperCommand,
      inlineShortcuts,
      suggestionCount,
    ]);

    useEffect(() => {
      if (!accepted) return;
      for (const [lineId, record] of hostRecords.current) {
        const line = lines.find((item) => item.id === lineId);
        if (!line) {
          record.mount.removeEventListener("focusin", record.focusListener);
          record.mount.removeEventListener("keydown", record.keyListener, true);
          try { record.host.dispose(); } catch (error) { console.warn("VisualTeX core host cleanup:", error); }
          try { if (!record.owner.isDestroyed) record.owner.destroy(); } catch (error) { console.warn("VisualTeX core owner cleanup:", error); }
          hostRecords.current.delete(lineId);
          continue;
        }
        if (record.host.inspect().document.source !== line.latex) {
          record.host.replaceSource(line.latex);
        }
        record.host.setZoom(zoom);
      }
    }, [accepted, lines, zoom]);

    useEffect(
      () => () => {
        onCoreHistoryStateChange?.(null);
        for (const record of hostRecords.current.values()) {
          record.mount.removeEventListener("focusin", record.focusListener);
          record.mount.removeEventListener("keydown", record.keyListener, true);
          try { record.host.dispose(); } catch (error) { console.warn("VisualTeX core host cleanup:", error); }
          try { if (!record.owner.isDestroyed) record.owner.destroy(); } catch (error) { console.warn("VisualTeX core owner cleanup:", error); }
        }
        hostRecords.current.clear();
      },
      [],
    );

    useImperativeHandle(ref, () => ({
      insertCommand: (command) => {
        activeRecord()?.host.applyEditCommand(coreEditCommandForProductCommand(command));
      },
      insertLatex: (latex) => {
        activeRecord()?.host.applyEditCommand({ kind: "pasteLatex", latex });
      },
      insertLatexAt: (target, latex) => {
        if (target.engine !== "core" || !target.coreRequest) return false;
        const record = hostRecords.current.get(target.lineId);
        if (!record) return false;
        const result = record.host.applyOcrResult(
          target.coreRequest as OcrInsertionTicket,
          latex,
        );
        return result.status === "applied";
      },
      appendLatex: (latex) => {
        const record = activeRecord();
        if (!record) return;
        record.host.applyEditCommand({ kind: "pasteLatex", latex });
      },
      focus: (options?: MathEditorFocusOptions) => {
        const line =
          options?.target === "first"
            ? lines[0]
            : options?.target === "last"
              ? lines.at(-1)
              : lines.find((item) => item.id === activeLineId) ?? lines[0];
        if (line) hostRecords.current.get(line.id)?.host.focus();
      },
      addLine: () => {
        const state = useEditorStore.getState();
        addOuterLineAfter(state.activeLineId ?? state.lines.at(-1)?.id);
      },
      commitPendingTransaction: () => {
        // BrowserInputBridge owns composition; there is no second pending
        // MathLive/history transaction to flush in the core path.
      },
      getSelectionMap: () => ({}),
      getCoreSelectionMap: () =>
        Object.fromEntries(
          [...hostRecords.current.entries()].map(([lineId, record]) => [
            lineId,
            record.host.captureSelectionToken() as CoreMathSelectionSnapshot,
          ]),
        ),
      restoreSelection: async (lineId, latex, _selection: MathSelectionSnapshot | null) => {
        const record = hostRecords.current.get(lineId);
        if (!record) return false;
        if (record.host.inspect().document.source !== latex) record.host.replaceSource(latex);
        record.host.focus();
        return true;
      },
      restoreCoreSelection: async (
        lineId,
        latex,
        selection: CoreMathSelectionSnapshot | null,
      ) => {
        const record = hostRecords.current.get(lineId);
        if (!record) return false;
        if (record.host.inspect().document.source !== latex) return false;
        if (!selection) {
          record.host.focus();
          return true;
        }
        return record.host.restoreSelectionToken(selection);
      },
      captureSelectionTarget: (): MathEditorSelectionTarget | null => {
        const lineId = useEditorStore.getState().activeLineId ?? activeLineId;
        return lineId && hostRecords.current.has(lineId)
          ? { engine: "core", coreLineId: lineId, selections: [] }
          : null;
      },
      captureInsertionTarget: (): MathEditorInsertionTarget | null => {
        const lineId = useEditorStore.getState().activeLineId ?? activeLineId;
        if (!lineId) return null;
        const record = hostRecords.current.get(lineId);
        if (!record) return null;
        return {
          engine: "core",
          lineId,
          ranges: [],
          direction: "none",
          coreRequest: record.host.beginOcrInsertion() as object,
        };
      },
      applySelectionStyle: (
        style: MathEditorSelectionStyle,
        target?: MathEditorSelectionTarget | null,
      ) => {
        const lineId = target?.engine === "core" ? target.coreLineId : activeLineId;
        const host = lineId ? hostRecords.current.get(lineId)?.host : activeRecord()?.host;
        if (!host) return false;
        if (style.kind === "bold") host.applyStyle("\\mathbf");
        else if (style.kind === "italic") host.applyStyle("\\mathit");
        else if (style.kind === "color") {
          const color = hexColor(style.value);
          if (!color) return false;
          host.applyColor(color);
        } else {
          // Background-color wrappers are not yet a core P10 capability.
          return false;
        }
        return true;
      },
      refreshLayout: () => {
        for (const record of hostRecords.current.values()) record.host.setZoom(zoom);
      },
      replaceLineSource: (lineId, latex) => {
        const record = hostRecords.current.get(lineId);
        if (!record) return false;
        if (record.host.inspect().document.source !== latex) record.host.replaceSource(latex);
        return true;
      },
      getCoreExportPort: (): CoreEditorExportPort | null => {
        const preferredLineId = useEditorStore.getState().activeLineId ?? activeLineId;
        const entry = preferredLineId && hostRecords.current.has(preferredLineId)
          ? [preferredLineId, hostRecords.current.get(preferredLineId)!] as const
          : hostRecords.current.entries().next().value as [string, HostRecord] | undefined;
        if (!entry) return null;
        const [lineId, record] = entry;
        return {
          lineId,
          documentId: record.owner.documentId,
          revision: record.owner.revision,
          owner: record.owner,
          layout: structuredClone(record.layout),
          resolveLayout: async (letterFont, chineseFont, style) => {
            const context = await getCoreEngineRuntime();
            return resolveCoreLayoutForPreferences(
              context,
              letterFont,
              chineseFont,
              style,
            );
          },
        };
      },
      history: (redo) => {
        const record = activeRecord();
        if (!record) return false;
        record.host.history(redo);
        return true;
      },
    }));

    const interceptImagePaste = (event: React.ClipboardEvent<HTMLDivElement>) => {
      if (!onPasteImage) return;
      const file = Array.from(event.clipboardData.files).find((item) =>
        item.type.startsWith("image/"),
      );
      if (!file) return;
      const lineId = useEditorStore.getState().activeLineId ?? activeLineId;
      const record = lineId ? hostRecords.current.get(lineId) : undefined;
      if (!lineId || !record) return;
      const target: MathEditorInsertionTarget = {
        engine: "core",
        lineId,
        ranges: [],
        direction: "none",
        coreRequest: record.host.beginOcrInsertion() as object,
      };
      event.preventDefault();
      onPasteImage(file, target);
    };

    const popupSuggestions = commandPopup
      ? commandPopup.snapshot.candidates.flatMap((candidate) => {
          const command = commandById.get(candidate.id);
          return command ? [command] : [];
        })
      : [];

    return (
      <section
        className={"editor-surface multi-line-editor core-math-editor"
          + (showLineNumbers ? " has-line-numbers" : "")
          + (latexCodeFormat === "mixed-inline-display" && !readOnly && !previewOnly ? " has-mixed-line-modes" : "")
          + (previewOnly || readOnly ? " is-preview-only is-read-only-preview" : "")
          + (previewOnly ? " is-source-preview-only" : "")}
        style={{
          "--formula-area-inset-left": `${formulaInsetLeft}px`,
          "--formula-area-inset-right": `${formulaInsetRight}px`,
          "--formula-row-vertical-inset": `${formulaRowVerticalInset}px`,
          "--core-row-min-height": `${Math.ceil(Math.max(30, 54 * zoom * 1.2 + 2 * formulaRowVerticalInset))}px`,
        } as CSSProperties}
        data-engine="visualtex-core"
        data-core-ready={accepted ? "true" : "false"}
        data-requested-font={formulaLetterFont}
        data-resolved-font={fontResolution?.actualLetterFamily}
        data-font-source={fontResolution?.letterSource}
        data-command-query={commandPopup?.snapshot.query ?? ""}
        data-source-draft-error={draftError}
        data-active-line-id={activeLineId ?? ""}
        data-formula-alignment={formulaAlignment}
        onPasteCapture={interceptImagePaste}
        aria-busy={!accepted && !probeError}
      >
        {!accepted && (
          <div className="core-math-editor-notice" role="status">
            {probeError
              ? (language === "en" ? "Formula resources could not be loaded. Your source is preserved." : "公式资源加载失败，源码已保留。")
              : (language === "en" ? "Loading bundled formula fonts…" : "正在加载软件内置公式字体…")}
            {probeError && <button type="button" onClick={() => retryCoreStaticPreviews()}>
              {language === "en" ? "Retry" : "重试"}
            </button>}
            {probeError && <details><summary>{language === "en" ? "Details" : "错误详情"}</summary>{probeError}</details>}
          </div>
        )}
        {accepted && layoutNotice ? <span className="visually-hidden" role="status">{layoutNotice}</span> : null}
        {overlay}
        <div className="mathfield-stack">
          {lines.map((line, index) => (
            <div key={line.id} className={`formula-line${line.id === activeLineId ? " is-active" : ""}`} data-line-id={line.id}
              style={{ "--formula-row-height": "var(--core-row-min-height)" } as CSSProperties}>
              {showLineNumbers && <span className="formula-line-number">{String(index + 1).padStart(2, "0")}</span>}
              {latexCodeFormat === "mixed-inline-display" && !previewOnly && !readOnly && (
                <div className="formula-line-mode-toggle" role="group" aria-label={language === "en" ? "Formula row mode" : "公式行模式"}>
                  <button type="button" className={line.mode === "inline" ? "is-active" : ""}
                    onPointerDown={(event) => event.preventDefault()} onClick={() => setFormulaLineMode(line.id, "inline")}>{language === "en" ? "Inline" : "行内"}</button>
                  <button type="button" className={line.mode !== "inline" ? "is-active" : ""}
                    onPointerDown={(event) => event.preventDefault()} onClick={() => setFormulaLineMode(line.id, "display")}>{language === "en" ? "Display" : "行间"}</button>
                </div>
              )}
              <div className="core-formula-cell">
                <div className={`mathfield-host core-formula-line is-${line.mode === "inline" ? "inline" : "display"}`}
                  data-line-id={line.id}
                  ref={(node) => {
                    if (node) mountRefs.current.set(line.id, node);
                    else mountRefs.current.delete(line.id);
                  }}
                  onPointerDown={(event) => {
                    if (!previewOnly && !readOnly && (event.target === event.currentTarget || (event.target as HTMLElement).classList?.contains("vt-core-view"))) {
                      hostRecords.current.get(line.id)?.host.focus();
                    }
                  }}
                />
                {(!accepted || rowErrors[line.id]) && line.latex && (
                  <pre className="core-row-source-preserved" title={rowErrors[line.id]}>{line.latex}</pre>
                )}
              </div>
            </div>
          ))}
        </div>
        {(previewOnly || readOnly) && (
          <button
            type="button"
            className="core-math-editor-preview-shield"
            aria-label="Activate formula editor"
            onClick={() => onPreviewActivate?.()}
          />
        )}
        {commandPopup && popupSuggestions.length && !previewOnly && !readOnly ? (
          <CommandSuggestionPopup
            suggestions={popupSuggestions}
            selectedIndex={Math.min(
              commandPopup.snapshot.selectedIndex,
              popupSuggestions.length - 1,
            )}
            position={commandPopup.position}
            usage={usage}
            onHighlight={(index) => {
              hostRecords.current.get(commandPopup.lineId)?.host.selectCommandCandidate(index);
            }}
            onCommit={(command) => {
              const index = popupSuggestions.findIndex((item) => item.id === command.id);
              if (index >= 0) {
                hostRecords.current.get(commandPopup.lineId)?.host.confirmCommandCandidate(index);
              }
            }}
          />
        ) : null}
      </section>
    );
  },
);
