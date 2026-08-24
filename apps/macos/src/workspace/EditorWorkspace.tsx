import {
  useEffect,
  useRef,
  useState,
  type CSSProperties,
  type PointerEvent as ReactPointerEvent,
  type ReactNode,
} from "react";
import { createPortal } from "react-dom";
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Bold,
  Braces,
  Camera,
  ChevronDown,
  Code2,
  Copy,
  FileDown,
  Highlighter,
  Italic,
  Minus,
  Palette,
  PanelBottomClose,
  PanelBottomOpen,
  PanelRightOpen,
  Plus,
  ScanLine,
  X,
} from "lucide-react";
import {
  MathEditor,
  type MathEditorSelectionTarget,
} from "../editor/MathEditor";
import { InputBehaviorMenu } from "../components/InputBehaviorMenu";
import { FormulaToolbar } from "../toolbar/FormulaToolbar";
import { LatexSourceEditor } from "../source-editor/LatexSourceEditor";
import {
  DEFAULT_CLASSIC_DOCK_HEIGHT,
  DEFAULT_CLASSIC_TILE_WIDTH,
  EDITOR_ZOOM_STEP,
  MAX_CLASSIC_DOCK_HEIGHT,
  MAX_CLASSIC_TILE_WIDTH,
  MAX_EDITOR_ZOOM,
  MIN_CLASSIC_DOCK_HEIGHT,
  MIN_CLASSIC_TILE_WIDTH,
  MIN_EDITOR_ZOOM,
  joinFormulaLines,
  useEditorStore,
} from "../stores/editorStore";
import {
  formatFormulaLines,
  parseLatexSourceDraft,
} from "../clipboard/LatexCopyService";
import { normalizeChineseLatex } from "../editor/normalizeChineseLatex";
import { reconcileFormulaLines } from "../history/documentHistory";
import type { FormulaAlignment, FormulaLine } from "../types/formula";
import { normalizeCustomFormulaColor } from "./formulaColor";
import type { EditorWorkspaceProps } from "./workspaceTypes";
import {
  readWorkspacePanelOpen,
  writeWorkspacePanelOpen,
} from "./workspacePanelPreferences";

const formulaTextColorPresets = [
  "#111827",
  "#dc2626",
  "#ea580c",
  "#ca8a04",
  "#16a34a",
  "#0891b2",
  "#2563eb",
  "#7c3aed",
  "#db2777",
] as const;

const formulaBackgroundColorPresets = [
  "#fef3c7",
  "#fed7aa",
  "#fecaca",
  "#fbcfe8",
  "#ddd6fe",
  "#bfdbfe",
  "#a5f3fc",
  "#bbf7d0",
  "#e5e7eb",
] as const;

type FormulaColorMenu = "color" | "backgroundColor";
type ClassicResizeTarget = "tiles" | "dock";

const compactOfficeTileBreakpoint = 760;
const compactOfficeEditorReserve = 220;

const customFormulaTextColorsStorageKey = "visualtex-custom-formula-text-colors";
const customFormulaBackgroundColorsStorageKey =
  "visualtex-custom-formula-background-colors";
const maximumCustomFormulaColors = 8;

function PortalOrInline({
  target,
  children,
}: {
  target: HTMLElement | null;
  children: ReactNode;
}) {
  return target ? createPortal(children, target) : children;
}

function clampPanelSize(value: number, minimum: number, maximum: number) {
  return Math.min(maximum, Math.max(minimum, value));
}

function loadCustomFormulaColors(storageKey: string) {
  try {
    const parsed = JSON.parse(localStorage.getItem(storageKey) ?? "[]");
    if (!Array.isArray(parsed)) return [];
    return Array.from(
      new Set(parsed.map(normalizeCustomFormulaColor).filter(Boolean)),
    ).slice(0, maximumCustomFormulaColors) as string[];
  } catch {
    return [];
  }
}

function persistCustomFormulaColors(storageKey: string, colors: string[]) {
  try {
    localStorage.setItem(storageKey, JSON.stringify(colors));
  } catch {
    // Custom colors still work for the current session without storage.
  }
}

export function EditorWorkspace({
  mode,
  showFileActions,
  showOfficeActions,
  showOcrActions,
  primaryActionLabel,
  officeHeaderLeadingControls,
  officeHeaderTrailingActions,
  desktopHeaderControls,
  keypadMode = false,
  onPrimaryAction,
  onCancel,
  onOpenExport,
  editorRef,
  editorInstanceKey,
  reuseEditorLineSlots = false,
  sidebarOpen,
  onSidebarOpenChange,
  onHistoryBusyChange,
  onPasteImage,
  onCopyPng,
  onCopy,
  onReplaceDocument,
  ocrModel,
  ocrModels = [],
  ocrBusy = false,
  onOcrModelChange,
  onQuickOcr,
  quickOcrCaptureMode = "immediate",
  onQuickOcrCaptureModeChange,
  silentOcrEnabled = false,
  onSilentOcrEnabledChange,
  ocrOverlay,
}: EditorWorkspaceProps) {
  const [primaryBusy, setPrimaryBusy] = useState(false);
  const [quickOcrModeMenuOpen, setQuickOcrModeMenuOpen] = useState(false);
  const quickOcrModeMenuRef = useRef<HTMLDivElement>(null);
  const [classicDockOpen, setClassicDockOpenState] = useState(() =>
    readWorkspacePanelOpen(mode, "toolbar"),
  );
  const setClassicDockOpen = (
    next: boolean | ((current: boolean) => boolean),
  ) => {
    setClassicDockOpenState((current) => {
      const resolved = typeof next === "function" ? next(current) : next;
      writeWorkspacePanelOpen(mode, "toolbar", resolved);
      return resolved;
    });
  };
  const [officeFormattingMount, setOfficeFormattingMount] =
    useState<HTMLDivElement | null>(null);
  const [formulaColorMenu, setFormulaColorMenu] =
    useState<FormulaColorMenu | null>(null);
  const formulaColorMenuRef = useRef<HTMLDivElement>(null);
  const formulaSelectionTargetRef = useRef<MathEditorSelectionTarget | null>(null);
  const customFormulaColorDraftRef = useRef<
    Record<FormulaColorMenu, string | null>
  >({ color: null, backgroundColor: null });
  const workspaceRef = useRef<HTMLElement>(null);
  const classicEditorBodyRef = useRef<HTMLDivElement>(null);
  const activeResizeCleanupRef = useRef<(() => void) | null>(null);
  const resizeFrameRef = useRef<number | null>(null);
  const persistedClassicTileWidth = useEditorStore(
    (state) => state.classicTileWidth,
  );
  const persistClassicTileWidth = useEditorStore(
    (state) => state.setClassicTileWidth,
  );
  const persistedClassicDockHeight = useEditorStore(
    (state) => state.classicDockHeight,
  );
  const persistClassicDockHeight = useEditorStore(
    (state) => state.setClassicDockHeight,
  );
  const [classicTileWidth, setClassicTileWidth] = useState(
    persistedClassicTileWidth,
  );
  const [classicDockHeight, setClassicDockHeight] = useState(
    persistedClassicDockHeight,
  );
  const [formulaTextColor, setFormulaTextColor] = useState("#2563eb");
  const [formulaBackgroundColor, setFormulaBackgroundColor] = useState("#fef3c7");
  const [formulaTextColorPickerValue, setFormulaTextColorPickerValue] =
    useState("#2563eb");
  const [formulaBackgroundColorPickerValue, setFormulaBackgroundColorPickerValue] =
    useState("#fef3c7");
  const [customFormulaTextColors, setCustomFormulaTextColors] = useState(() =>
    loadCustomFormulaColors(customFormulaTextColorsStorageKey),
  );
  const [customFormulaBackgroundColors, setCustomFormulaBackgroundColors] =
    useState(() =>
      loadCustomFormulaColors(customFormulaBackgroundColorsStorageKey),
    );
  const [sourceDraftFallback, setSourceDraftFallback] = useState<{
    source: string;
    error: string;
    previewLines: FormulaLine[] | null;
  } | null>(null);
  const sourceDraftFallbackRef = useRef(sourceDraftFallback);
  const [sourceFocused, setSourceFocused] = useState(false);
  const sourceFocusedRef = useRef(false);
  const title = useEditorStore((state) => state.title);
  const lines = useEditorStore((state) => state.lines);
  const activeLineId = useEditorStore((state) => state.activeLineId);
  const language = useEditorStore((state) => state.language);
  const theme = useEditorStore((state) => state.theme);
  const zoom = useEditorStore((state) => state.zoom);
  const setZoom = useEditorStore((state) => state.setZoom);
  const formulaAlignment = useEditorStore((state) => state.formulaAlignment);
  const setFormulaAlignment = useEditorStore(
    (state) => state.setFormulaAlignment,
  );
  const editorLayout = useEditorStore((state) => state.editorLayout);
  const highlightActiveLine = useEditorStore(
    (state) => state.highlightActiveLine,
  );
  const sourceOpen = useEditorStore((state) => state.sourceOpen);
  const setSourceOpen = useEditorStore((state) => state.setSourceOpen);
  const latexCodeFormat = useEditorStore((state) => state.latexCodeFormat);
  const isEn = language === "en";
  const isOfficeWorkspace = mode !== "desktop";
  const latex = joinFormulaLines(lines);
  const sourceLatex = formatFormulaLines(lines, latexCodeFormat);

  const acceptSourcePreview = () => {
    const previewLines = sourceDraftFallbackRef.current?.previewLines;
    if (!previewLines?.length) return;

    sourceDraftFallbackRef.current = null;
    setSourceDraftFallback(null);
    const nextActiveLineId = previewLines.some(
      (line) => line.id === activeLineId,
    )
      ? activeLineId
      : previewLines[0]?.id ?? null;
    onReplaceDocument(
      {
        title,
        lines: previewLines,
        activeLineId: nextActiveLineId,
        formulaAlignment,
        selectionByLineId: {},
      },
      "source-apply",
    );
  };

  const handleSourceFocusChange = (focused: boolean) => {
    sourceFocusedRef.current = focused;
    document.documentElement.classList.toggle(
      "visualtex-source-editor-focused",
      focused,
    );
    setSourceFocused(focused);
    if (!focused) acceptSourcePreview();
  };

  useEffect(() => {
    document.documentElement.classList.toggle(
      "visualtex-source-editor-focused",
      sourceFocused,
    );
    return () => {
      document.documentElement.classList.remove(
        "visualtex-source-editor-focused",
      );
    };
  }, [sourceFocused]);

  useEffect(() => {
    setClassicTileWidth(persistedClassicTileWidth);
  }, [persistedClassicTileWidth]);

  useEffect(() => {
    setClassicDockHeight(persistedClassicDockHeight);
  }, [persistedClassicDockHeight]);

  useEffect(() => {
    if (!sourceOpen) {
      sourceDraftFallbackRef.current = null;
      setSourceDraftFallback(null);
    }
  }, [sourceOpen]);

  useEffect(() => {
    sourceDraftFallbackRef.current = null;
    setSourceDraftFallback(null);
  }, [latexCodeFormat]);

  useEffect(() => {
    if (!quickOcrModeMenuOpen) return;
    const close = (event: PointerEvent) => {
      if (
        event.target instanceof Node &&
        quickOcrModeMenuRef.current?.contains(event.target)
      ) {
        return;
      }
      setQuickOcrModeMenuOpen(false);
    };
    const closeFromKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") setQuickOcrModeMenuOpen(false);
    };
    document.addEventListener("pointerdown", close, true);
    document.addEventListener("keydown", closeFromKey, true);
    return () => {
      document.removeEventListener("pointerdown", close, true);
      document.removeEventListener("keydown", closeFromKey, true);
    };
  }, [quickOcrModeMenuOpen]);

  useEffect(() => {
    if (!formulaColorMenu) return;
    const close = (event: PointerEvent) => {
      if (
        event.target instanceof Node &&
        formulaColorMenuRef.current?.contains(event.target)
      ) {
        return;
      }
      setFormulaColorMenu(null);
      formulaSelectionTargetRef.current = null;
    };
    const closeFromKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setFormulaColorMenu(null);
        formulaSelectionTargetRef.current = null;
      }
    };
    document.addEventListener("pointerdown", close, true);
    document.addEventListener("keydown", closeFromKey, true);
    return () => {
      document.removeEventListener("pointerdown", close, true);
      document.removeEventListener("keydown", closeFromKey, true);
    };
  }, [formulaColorMenu]);

  useEffect(
    () => () => {
      activeResizeCleanupRef.current?.();
      if (resizeFrameRef.current !== null) {
        window.cancelAnimationFrame(resizeFrameRef.current);
        resizeFrameRef.current = null;
      }
    },
    [],
  );

  const classicTileWidthLimit = () => {
    const workspaceWidth = workspaceRef.current?.getBoundingClientRect().width;
    if (!workspaceWidth) return MAX_CLASSIC_TILE_WIDTH;
    const editorReserve =
      isOfficeWorkspace && window.innerWidth <= compactOfficeTileBreakpoint
        ? compactOfficeEditorReserve
        : 360;
    return Math.max(MIN_CLASSIC_TILE_WIDTH, workspaceWidth - editorReserve);
  };

  const classicDockHeightLimit = () => {
    const editorHeight =
      classicEditorBodyRef.current?.getBoundingClientRect().height;
    if (!editorHeight) return MAX_CLASSIC_DOCK_HEIGHT;
    return Math.max(MIN_CLASSIC_DOCK_HEIGHT, editorHeight - 120);
  };

  const commitClassicPanelSize = (
    target: ClassicResizeTarget,
    value: number,
    persist = false,
  ) => {
    if (target === "tiles") {
      const next = clampPanelSize(
        value,
        MIN_CLASSIC_TILE_WIDTH,
        classicTileWidthLimit(),
      );
      workspaceRef.current?.style.setProperty(
        "--classic-tile-width",
        `${next}px`,
      );
      setClassicTileWidth(next);
      if (persist) persistClassicTileWidth(next);
      return next;
    }

    const next = clampPanelSize(
      value,
      MIN_CLASSIC_DOCK_HEIGHT,
      classicDockHeightLimit(),
    );
    classicEditorBodyRef.current?.style.setProperty(
      "--classic-dock-height",
      `${next}px`,
    );
    setClassicDockHeight(next);
    if (persist) persistClassicDockHeight(next);
    return next;
  };

  const startClassicResize = (
    target: ClassicResizeTarget,
    event: ReactPointerEvent<HTMLDivElement>,
  ) => {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    activeResizeCleanupRef.current?.();

    const handle = event.currentTarget;
    const pointerId = event.pointerId;
    let latestValue =
      target === "tiles" ? classicTileWidth : classicDockHeight;
    let finished = false;

    const applyLatestValue = () => {
      resizeFrameRef.current = null;
      commitClassicPanelSize(target, latestValue);
    };
    const scheduleValue = (value: number) => {
      latestValue = value;
      if (resizeFrameRef.current !== null) return;
      resizeFrameRef.current = window.requestAnimationFrame(applyLatestValue);
    };
    const move = (pointerEvent: PointerEvent) => {
      if (pointerEvent.pointerId !== pointerId) return;
      if (target === "tiles") {
        const bounds = workspaceRef.current?.getBoundingClientRect();
        if (bounds) scheduleValue(bounds.right - pointerEvent.clientX);
      } else {
        const bounds = classicEditorBodyRef.current?.getBoundingClientRect();
        if (bounds) scheduleValue(bounds.bottom - pointerEvent.clientY);
      }
      pointerEvent.preventDefault();
    };
    const finish = () => {
      if (finished) return;
      finished = true;
      document.removeEventListener("pointermove", move);
      document.removeEventListener("pointerup", finish);
      document.removeEventListener("pointercancel", finish);
      window.removeEventListener("blur", finish);
      if (resizeFrameRef.current !== null) {
        window.cancelAnimationFrame(resizeFrameRef.current);
        resizeFrameRef.current = null;
      }
      latestValue = commitClassicPanelSize(target, latestValue, true);
      if (handle.hasPointerCapture(pointerId)) {
        handle.releasePointerCapture(pointerId);
      }
      delete document.body.dataset.workspaceResize;
      if (activeResizeCleanupRef.current === finish) {
        activeResizeCleanupRef.current = null;
      }
    };

    document.body.dataset.workspaceResize = target;
    handle.setPointerCapture(pointerId);
    document.addEventListener("pointermove", move, { passive: false });
    document.addEventListener("pointerup", finish);
    document.addEventListener("pointercancel", finish);
    window.addEventListener("blur", finish);
    activeResizeCleanupRef.current = finish;
  };

  const adjustClassicPanelFromKeyboard = (
    target: ClassicResizeTarget,
    delta: number,
  ) => {
    const current = target === "tiles" ? classicTileWidth : classicDockHeight;
    commitClassicPanelSize(target, current + delta, true);
  };

  const resetClassicPanelSize = (target: ClassicResizeTarget) => {
    commitClassicPanelSize(
      target,
      target === "tiles"
        ? DEFAULT_CLASSIC_TILE_WIDTH
        : DEFAULT_CLASSIC_DOCK_HEIGHT,
      true,
    );
  };

  useEffect(() => {
    if (editorLayout !== "classic") return;
    const workspace = workspaceRef.current;
    const editorBody = classicEditorBodyRef.current;
    if (!workspace || !editorBody) return;

    let frame = 0;
    const constrainPanelSizes = () => {
      window.cancelAnimationFrame(frame);
      frame = window.requestAnimationFrame(() => {
        const workspaceWidth = workspace.getBoundingClientRect().width;
        const editorReserve =
          isOfficeWorkspace && window.innerWidth <= compactOfficeTileBreakpoint
            ? compactOfficeEditorReserve
            : 360;
        const tileMaximum = Math.max(
          MIN_CLASSIC_TILE_WIDTH,
          workspaceWidth - editorReserve,
        );
        setClassicTileWidth((current) => {
          const next = clampPanelSize(
            current,
            MIN_CLASSIC_TILE_WIDTH,
            tileMaximum,
          );
          if (next !== current) persistClassicTileWidth(next);
          return next;
        });

        const editorHeight = editorBody.getBoundingClientRect().height;
        const dockMaximum = Math.max(
          MIN_CLASSIC_DOCK_HEIGHT,
          editorHeight - 120,
        );
        setClassicDockHeight((current) => {
          const next = clampPanelSize(
            current,
            MIN_CLASSIC_DOCK_HEIGHT,
            dockMaximum,
          );
          if (next !== current) persistClassicDockHeight(next);
          return next;
        });
      });
    };

    const observer = new ResizeObserver(constrainPanelSizes);
    observer.observe(workspace);
    observer.observe(editorBody);
    constrainPanelSizes();
    return () => {
      window.cancelAnimationFrame(frame);
      observer.disconnect();
    };
  }, [editorLayout, isOfficeWorkspace]);

  const preserveFormulaFocus = (event: ReactPointerEvent<HTMLButtonElement>) => {
    event.preventDefault();
    event.stopPropagation();
  };

  const rememberFormulaSelection = () => {
    formulaSelectionTargetRef.current =
      editorRef.current?.captureSelectionTarget() ?? null;
  };

  const preserveFormulaSelection = (
    event: ReactPointerEvent<HTMLButtonElement>,
  ) => {
    preserveFormulaFocus(event);
    const liveTarget = editorRef.current?.captureSelectionTarget() ?? null;
    if (liveTarget) formulaSelectionTargetRef.current = liveTarget;
  };

  const applySelectedFormulaStyle = (kind: "bold" | "italic") => {
    const target =
      formulaSelectionTargetRef.current ??
      editorRef.current?.captureSelectionTarget() ??
      null;
    if (target) editorRef.current?.applySelectionStyle({ kind }, target);
    formulaSelectionTargetRef.current = null;
  };

  const applySelectedFormulaStyleFromPointer = (
    event: ReactPointerEvent<HTMLButtonElement>,
    kind: "bold" | "italic",
  ) => {
    preserveFormulaSelection(event);
    applySelectedFormulaStyle(kind);
  };

  const toggleFormulaColorMenu = (kind: FormulaColorMenu) => {
    const selection =
      formulaSelectionTargetRef.current ??
      editorRef.current?.captureSelectionTarget() ??
      null;
    if (!selection) {
      setFormulaColorMenu(null);
      formulaSelectionTargetRef.current = null;
      return;
    }
    formulaSelectionTargetRef.current = selection;
    setFormulaColorMenu((current) => {
      if (current === kind) {
        formulaSelectionTargetRef.current = null;
        return null;
      }
      return kind;
    });
  };

  const applySelectedFormulaColor = (
    kind: FormulaColorMenu,
    value: string,
  ) => {
    const target = formulaSelectionTargetRef.current;
    if (!target) return;
    editorRef.current?.applySelectionStyle({ kind, value }, target);
    if (kind === "color") {
      setFormulaTextColor(value);
      setFormulaTextColorPickerValue(value);
    } else {
      setFormulaBackgroundColor(value);
      setFormulaBackgroundColorPickerValue(value);
    }
    setFormulaColorMenu(null);
    formulaSelectionTargetRef.current = null;
  };

  const updateCustomFormulaColors = (
    kind: FormulaColorMenu,
    updater: (colors: string[]) => string[],
  ) => {
    const storageKey =
      kind === "color"
        ? customFormulaTextColorsStorageKey
        : customFormulaBackgroundColorsStorageKey;
    const setter =
      kind === "color"
        ? setCustomFormulaTextColors
        : setCustomFormulaBackgroundColors;
    setter((current) => {
      const next = updater(current);
      persistCustomFormulaColors(storageKey, next);
      return next;
    });
  };

  const beginCustomFormulaColorSelection = (kind: FormulaColorMenu) => {
    customFormulaColorDraftRef.current[kind] = null;
  };

  const saveCustomFormulaColor = (
    kind: FormulaColorMenu,
    value: string,
  ) => {
    const normalized = normalizeCustomFormulaColor(value);
    if (!normalized) return;
    if (kind === "color") setFormulaTextColorPickerValue(normalized);
    else setFormulaBackgroundColorPickerValue(normalized);

    const presets =
      kind === "color"
        ? formulaTextColorPresets
        : formulaBackgroundColorPresets;
    if ((presets as readonly string[]).includes(normalized)) return;

    const previousDraft = customFormulaColorDraftRef.current[kind];
    updateCustomFormulaColors(kind, (current) => [
      normalized,
      ...current.filter(
        (color) => color !== normalized && color !== previousDraft,
      ),
    ].slice(0, maximumCustomFormulaColors));
    customFormulaColorDraftRef.current[kind] = normalized;
  };

  const removeCustomFormulaColor = (
    kind: FormulaColorMenu,
    color: string,
  ) => {
    updateCustomFormulaColors(kind, (current) =>
      current.filter((item) => item !== color),
    );
    if (customFormulaColorDraftRef.current[kind] === color) {
      customFormulaColorDraftRef.current[kind] = null;
    }
  };

  const applyFormulaAlignment = (alignment: FormulaAlignment) => {
    setFormulaAlignment(alignment);
    editorRef.current?.focus();
  };

  const applySource = (source: string, sourceFormat: typeof latexCodeFormat) => {
    const parsed = parseLatexSourceDraft(source, sourceFormat);
    if (!parsed.valid) {
      const previewValues = parsed.values.map(normalizeChineseLatex);
      const fallback = {
        source,
        error: parsed.error ?? "invalid-latex",
        previewLines: previewValues.length
          ? reconcileFormulaLines(previewValues, lines, parsed.modes)
          : null,
      };
      sourceDraftFallbackRef.current = fallback;
      setSourceDraftFallback(fallback);
      return parsed;
    }

    sourceDraftFallbackRef.current = null;
    setSourceDraftFallback(null);
    const values = parsed.values.map(normalizeChineseLatex);
    const nextLines = reconcileFormulaLines(values, lines, parsed.modes);
    const nextActiveLineId = nextLines.some(
      (line) => line.id === activeLineId,
    )
      ? activeLineId
      : nextLines[0]?.id ?? null;
    onReplaceDocument(
      {
        title,
        lines: nextLines,
        activeLineId: nextActiveLineId,
        formulaAlignment,
        selectionByLineId: sourceFocusedRef.current
          ? {}
          : editorRef.current?.getSelectionMap() ?? {},
      },
      "source-apply",
    );
    return parsed;
  };
  const renderSourceEditor = ({
    showCollapseAction = true,
    showCopyAction = true,
    compact = false,
  }: {
    showCollapseAction?: boolean;
    showCopyAction?: boolean;
    compact?: boolean;
  } = {}) => (
    <LatexSourceEditor
      latex={sourceLatex}
      theme={theme}
      format={latexCodeFormat}
      onCollapse={() => setSourceOpen(false)}
      showCollapseAction={showCollapseAction}
      showCopyAction={showCopyAction}
      compact={compact}
      onLiveChange={applySource}
      onFocusChange={handleSourceFocusChange}
      onCopy={() => void onCopy()}
    />
  );

  const renderVisualEditor = () => {
    const previewLines = sourceDraftFallback?.previewLines;
    if (sourceDraftFallback && !previewLines) {
      return (
        <section
          className="editor-surface source-draft-fallback"
          data-source-draft-error={sourceDraftFallback.error}
          aria-live="polite"
          aria-label={
            isEn
              ? "Incomplete LaTeX source preview"
              : "未完成 LaTeX 源码预览"
          }
        >
          <div className="source-draft-fallback-heading">
            <Code2 size={16} />
            <span>
              {isEn
                ? "Complete the current LaTeX fragment to resume formula rendering."
                : "当前 LaTeX 片段尚未完成，补全后恢复公式渲染。"}
            </span>
          </div>
          <pre className="source-draft-fallback-code">
            {sourceDraftFallback.source || " "}
          </pre>
        </section>
      );
    }

    const visualLines = previewLines ?? lines;
    const visualActiveLineId = visualLines.some(
      (line) => line.id === activeLineId,
    )
      ? activeLineId
      : visualLines[0]?.id ?? null;
    return (
      <MathEditor
        key={editorInstanceKey}
        ref={editorRef}
        lines={visualLines}
        activeLineId={visualActiveLineId}
        reuseLineSlots={reuseEditorLineSlots}
        formulaAlignment={formulaAlignment}
        latexCodeFormat={latexCodeFormat}
        zoom={zoom}
        readOnly={false}
        previewOnly={sourceFocused}
        onPreviewActivate={() => handleSourceFocusChange(false)}
        draftError={sourceDraftFallback?.error}
        onPasteImage={
          previewLines ? undefined : showOcrActions ? onPasteImage : undefined
        }
        onCopyPng={previewLines ? undefined : onCopyPng}
        onHistoryBusyChange={onHistoryBusyChange}
        overlay={previewLines ? undefined : ocrOverlay}
      />
    );
  };

  const runPrimaryAction = async () => {
    if (!onPrimaryAction || primaryBusy) return;
    setPrimaryBusy(true);
    try {
      await onPrimaryAction();
    } finally {
      setPrimaryBusy(false);
    }
  };

  return (
    <>
      {showOfficeActions && (
        <div className="office-workspace-actions" data-workspace-mode={mode}>
          <div>
            <strong>
              {mode === "office-edit"
                ? isEn
                  ? "Edit selected formula"
                  : "编辑所选公式"
                : isEn
                  ? "Create Office formula"
                  : "新建 Office 公式"}
            </strong>
            <span>
              {isEn
                ? "The document is updated only after you finish or close this window."
                : "点击完成或关闭本窗口后，公式才会写入 Office 文档。"}
            </span>
          </div>
          <div>
            {onCancel && (
              <button
                type="button"
                className="secondary-button"
                onClick={() => void onCancel()}
                disabled={primaryBusy}
              >
                {isEn ? "Cancel" : "取消"}
              </button>
            )}
            {onPrimaryAction && (
              <button
                type="button"
                className="primary-button"
                onClick={() => void runPrimaryAction()}
                disabled={primaryBusy}
              >
                {primaryBusy
                  ? isEn
                    ? "Applying…"
                    : "正在应用…"
                  : primaryActionLabel ??
                    (mode === "office-edit"
                      ? isEn
                        ? "Update formula"
                        : "更新公式"
                      : isEn
                        ? "Finish and insert"
                        : "完成并插入")}
              </button>
            )}
          </div>
        </div>
      )}

      <main
        ref={workspaceRef}
        className={
          `workspace ${editorLayout === "classic" ? "is-classic-layout" : "is-standard-layout"}` +
          (!keypadMode && sidebarOpen ? " has-sidebar" : "") +
          (isOfficeWorkspace ? " is-office-workspace" : "") +
          (keypadMode ? " is-keypad-mode" : "") +
          (highlightActiveLine ? " has-active-line-highlight" : "") +
          (sourceFocused ? " is-source-editor-focused" : "")
        }
        style={
          {
            "--classic-tile-width": `${classicTileWidth}px`,
          } as CSSProperties
        }
        data-editor-layout={editorLayout}
      >
        {!keypadMode && editorLayout === "standard" && sidebarOpen && (
          <FormulaToolbar
            stabilizeTileLayout
            onInsert={(command) => editorRef.current?.insertCommand(command)}
          />
        )}

        {!keypadMode && editorLayout === "classic" && !sidebarOpen && (
          <button
            type="button"
            className="classic-tile-expand-button"
            data-formula-tile-expand
            aria-label={isEn ? "Expand formula tiles" : "展开公式磁贴"}
            title={isEn ? "Expand formula tiles" : "展开公式磁贴"}
            onClick={() => onSidebarOpenChange(true)}
          >
            <PanelRightOpen size={16} />
          </button>
        )}

        <section className="formula-workspace editor-pane">
          <header
            className={
              "workspace-heading pane-header editor-pane-header" +
              (isOfficeWorkspace ? " is-office-editor-header" : "")
            }
          >
            <div className="pane-title-group">
              {isOfficeWorkspace ? (
                officeHeaderLeadingControls ? (
                  <div className="office-inline-options">
                    {officeHeaderLeadingControls}
                  </div>
                ) : null
              ) : null}
              {!isOfficeWorkspace || officeFormattingMount ? (
                <PortalOrInline
                  target={isOfficeWorkspace ? officeFormattingMount : null}
                >
                  <div
                    className="formula-alignment-controls"
                role="toolbar"
                aria-label={isEn ? "Formula alignment" : "公式对齐方式"}
              >
                {(
                  [
                    ["left", AlignLeft, isEn ? "Align left" : "左对齐"],
                    ["center", AlignCenter, isEn ? "Centre" : "居中"],
                    ["right", AlignRight, isEn ? "Align right" : "右对齐"],
                  ] as const
                ).map(([alignment, Icon, label]) => (
                  <button
                    key={alignment}
                    type="button"
                    className={
                      "icon-button compact formula-alignment-button" +
                      (formulaAlignment === alignment ? " is-active" : "")
                    }
                    aria-label={label}
                    title={label}
                    aria-pressed={formulaAlignment === alignment}
                    data-formula-alignment={alignment}
                    onMouseDown={(event) => event.preventDefault()}
                    onClick={() => applyFormulaAlignment(alignment)}
                  >
                    <Icon size={16} strokeWidth={2} />
                  </button>
                ))}
                <span className="formula-formatting-divider" aria-hidden="true" />
                <div
                  ref={formulaColorMenuRef}
                  className="formula-formatting-controls"
                  role="group"
                  aria-label={isEn ? "Formula formatting" : "公式格式"}
                >
                  <button
                    type="button"
                    className="icon-button compact formula-formatting-button is-selection-action"
                    aria-label={
                      isEn
                        ? "Toggle bold for selected content"
                        : "切换选中内容的粗体状态"
                    }
                    title={
                      isEn
                        ? "Toggle bold while preserving math italic/upright shape"
                        : "切换粗体 · 保留原有数学斜体/正体 · 仅作用于选中内容"
                    }
                    data-formula-selection-bold
                    onPointerEnter={rememberFormulaSelection}
                    onPointerDown={(event) =>
                      applySelectedFormulaStyleFromPointer(event, "bold")
                    }
                    onClick={(event) => {
                      if (event.detail === 0) applySelectedFormulaStyle("bold");
                    }}
                  >
                    <Bold size={15} strokeWidth={2.2} />
                  </button>
                  <button
                    type="button"
                    className="icon-button compact formula-formatting-button is-selection-action"
                    aria-label={
                      isEn
                        ? "Toggle italic or upright for selected content"
                        : "切换选中内容的斜体或正体状态"
                    }
                    title={
                      isEn
                        ? "Toggle default math italic and \\mathrm upright"
                        : "切换默认数学斜体与 \\mathrm 正体 · 仅作用于选中内容"
                    }
                    data-formula-selection-italic
                    onPointerEnter={rememberFormulaSelection}
                    onPointerDown={(event) =>
                      applySelectedFormulaStyleFromPointer(event, "italic")
                    }
                    onClick={(event) => {
                      if (event.detail === 0) applySelectedFormulaStyle("italic");
                    }}
                  >
                    <Italic size={15} strokeWidth={2.2} />
                  </button>
                  <button
                    type="button"
                    className={
                      "icon-button compact formula-formatting-button is-color-action" +
                      (formulaColorMenu === "color" ? " is-active" : "")
                    }
                    style={
                      {
                        "--formula-format-color": formulaTextColor,
                      } as CSSProperties
                    }
                    aria-label={isEn ? "Selected text color" : "选中内容字体颜色"}
                    title={
                      isEn
                        ? "Apply a font color to the selection only"
                        : "字体颜色 · 仅应用于选中内容"
                    }
                    aria-pressed={formulaColorMenu === "color"}
                    data-formula-selection-color
                    onPointerEnter={rememberFormulaSelection}
                    onPointerDown={preserveFormulaSelection}
                    onMouseDown={(event) => event.preventDefault()}
                    onClick={() => toggleFormulaColorMenu("color")}
                  >
                    <Palette size={15} strokeWidth={2} />
                  </button>
                  <button
                    type="button"
                    className={
                      "icon-button compact formula-formatting-button is-color-action" +
                      (formulaColorMenu === "backgroundColor"
                        ? " is-active"
                        : "")
                    }
                    style={
                      {
                        "--formula-format-color": formulaBackgroundColor,
                      } as CSSProperties
                    }
                    aria-label={
                      isEn
                        ? "Selected text background color"
                        : "选中内容字体背景颜色"
                    }
                    title={
                      isEn
                        ? "Apply a background color to the selection only"
                        : "字体背景颜色 · 仅应用于选中内容"
                    }
                    aria-pressed={formulaColorMenu === "backgroundColor"}
                    data-formula-selection-background
                    onPointerEnter={rememberFormulaSelection}
                    onPointerDown={preserveFormulaSelection}
                    onMouseDown={(event) => event.preventDefault()}
                    onClick={() => toggleFormulaColorMenu("backgroundColor")}
                  >
                    <Highlighter size={15} strokeWidth={2} />
                  </button>

                  {formulaColorMenu && (
                    <div
                      className="formula-color-popover"
                      data-formula-color-popover={formulaColorMenu}
                      data-visualtex-floating-layer
                      role="dialog"
                      aria-label={
                        formulaColorMenu === "color"
                          ? isEn
                            ? "Formula text color"
                            : "公式字体颜色"
                          : isEn
                            ? "Formula background color"
                            : "公式背景颜色"
                      }
                    >
                      <strong>
                        {formulaColorMenu === "color"
                          ? isEn
                            ? "Text color"
                            : "字体颜色"
                          : isEn
                            ? "Background color"
                            : "背景颜色"}
                      </strong>
                      <div className="formula-color-content">
                        <section className="formula-color-presets">
                          <span className="formula-color-section-label">
                            {isEn ? "Preset" : "固定颜色"}
                          </span>
                          <div className="formula-color-swatches" role="group">
                            {(formulaColorMenu === "color"
                              ? formulaTextColorPresets
                              : formulaBackgroundColorPresets
                            ).map((color) => (
                              <button
                                key={color}
                                type="button"
                                className="formula-color-swatch"
                                style={{ backgroundColor: color }}
                                aria-label={`${isEn ? "Use" : "使用"} ${color}`}
                                title={color}
                                data-formula-color={color}
                                onMouseDown={(event) => event.preventDefault()}
                                onClick={() =>
                                  applySelectedFormulaColor(formulaColorMenu, color)
                                }
                              />
                            ))}
                            <label
                              className="formula-custom-color"
                              title={isEn ? "Add custom color" : "添加自定义颜色"}
                            >
                              <input
                                type="color"
                                value={
                                  formulaColorMenu === "color"
                                    ? formulaTextColorPickerValue
                                    : formulaBackgroundColorPickerValue
                                }
                                aria-label={
                                  isEn ? "Add custom color" : "添加自定义颜色"
                                }
                                onPointerDown={(event) => {
                                  event.stopPropagation();
                                  beginCustomFormulaColorSelection(formulaColorMenu);
                                }}
                                onInput={(event) =>
                                  saveCustomFormulaColor(
                                    formulaColorMenu,
                                    event.currentTarget.value,
                                  )
                                }
                                onChange={(event) =>
                                  saveCustomFormulaColor(
                                    formulaColorMenu,
                                    event.currentTarget.value,
                                  )
                                }
                              />
                              <Plus size={13} />
                            </label>
                          </div>
                        </section>
                        <section className="formula-custom-colors-panel">
                          <div className="formula-custom-colors-heading">
                            <span>{isEn ? "Custom" : "自定义颜色"}</span>
                            <small>
                              {(formulaColorMenu === "color"
                                ? customFormulaTextColors
                                : customFormulaBackgroundColors
                              ).length}
                              /{maximumCustomFormulaColors}
                            </small>
                          </div>
                          {(formulaColorMenu === "color"
                            ? customFormulaTextColors
                            : customFormulaBackgroundColors
                          ).length > 0 ? (
                            <div className="formula-custom-colors-grid">
                              {(formulaColorMenu === "color"
                                ? customFormulaTextColors
                                : customFormulaBackgroundColors
                              ).map((color) => (
                                <div
                                  key={color}
                                  className="formula-custom-color-item"
                                  data-formula-custom-color={color}
                                >
                                  <button
                                    type="button"
                                    className="formula-color-swatch"
                                    style={{ backgroundColor: color }}
                                    aria-label={`${isEn ? "Use custom" : "使用自定义颜色"} ${color}`}
                                    title={color}
                                    onMouseDown={(event) => event.preventDefault()}
                                    onClick={() =>
                                      applySelectedFormulaColor(
                                        formulaColorMenu,
                                        color,
                                      )
                                    }
                                  />
                                  <button
                                    type="button"
                                    className="formula-custom-color-delete"
                                    aria-label={`${isEn ? "Delete custom" : "删除自定义颜色"} ${color}`}
                                    title={isEn ? "Delete" : "删除"}
                                    data-delete-formula-custom-color={color}
                                    onPointerDown={(event) => {
                                      event.preventDefault();
                                      event.stopPropagation();
                                    }}
                                    onClick={(event) => {
                                      event.stopPropagation();
                                      removeCustomFormulaColor(
                                        formulaColorMenu,
                                        color,
                                      );
                                    }}
                                  >
                                    <X size={9} strokeWidth={2.4} />
                                  </button>
                                </div>
                              ))}
                            </div>
                          ) : (
                            <span className="formula-custom-colors-empty">
                              {isEn
                                ? "Pick + to save a color, then click its swatch to apply."
                                : "点击 + 保存颜色，再点击色块应用。"}
                            </span>
                          )}
                        </section>
                      </div>
                    </div>
                  )}
                </div>
                  </div>
                </PortalOrInline>
              ) : null}
              {!isOfficeWorkspace && desktopHeaderControls ? (
                <div className="desktop-editor-header-controls">
                  {desktopHeaderControls}
                </div>
              ) : null}
            </div>
            <div className="canvas-tool-group">
              {showFileActions && onOpenExport && (
                <button
                  type="button"
                  className="export-trigger-button workspace-export-trigger"
                  onClick={onOpenExport}
                  aria-label={isEn ? "Export" : "导出"}
                  title={isEn ? "Export Markdown, SVG or PNG" : "导出 Markdown、SVG 或 PNG"}
                >
                  <FileDown size={16} />
                  <span>{isEn ? "Export" : "导出"}</span>
                </button>
              )}
              <InputBehaviorMenu />
              {showOcrActions && onQuickOcr && (
                <div className="quick-ocr-controls">
                  <div className="quick-ocr-split" ref={quickOcrModeMenuRef}>
                    <button
                      type="button"
                      className="quick-ocr-button quick-ocr-primary"
                      onClick={() => {
                        setQuickOcrModeMenuOpen(false);
                        onQuickOcr();
                      }}
                      disabled={ocrBusy}
                      data-quick-ocr-button
                      title={
                        quickOcrCaptureMode === "system-screenshot"
                          ? isEn
                            ? "Minimize VisualTeX and wait for your next macOS screenshot"
                            : "最小化 VisualTeX，等待你下一次使用 macOS 系统截图键截图"
                          : isEn
                            ? "Minimize VisualTeX and immediately select a formula region"
                            : "最小化 VisualTeX 后立即框选公式区域"
                      }
                    >
                      <Camera size={15} />
                      <span>{isEn ? "Quick OCR" : "快捷 OCR"}</span>
                    </button>
                    {onQuickOcrCaptureModeChange && (
                      <button
                        type="button"
                        className={`quick-ocr-mode-trigger${quickOcrModeMenuOpen ? " is-open" : ""}${quickOcrCaptureMode === "system-screenshot" ? " is-system-screenshot" : ""}`}
                        onClick={() => setQuickOcrModeMenuOpen((open) => !open)}
                        disabled={ocrBusy}
                        aria-label={isEn ? "Choose Quick OCR capture mode" : "选择快捷 OCR 截图模式"}
                        aria-expanded={quickOcrModeMenuOpen}
                        data-quick-ocr-mode-trigger
                      >
                        <ChevronDown size={12} />
                      </button>
                    )}
                    {quickOcrModeMenuOpen && onQuickOcrCaptureModeChange && (
                      <div className="quick-ocr-mode-menu" role="menu" data-quick-ocr-mode-menu>
                        <button
                          type="button"
                          className={quickOcrCaptureMode === "immediate" ? "is-active" : ""}
                          onClick={() => {
                            onQuickOcrCaptureModeChange("immediate");
                            setQuickOcrModeMenuOpen(false);
                          }}
                          role="menuitemradio"
                          aria-checked={quickOcrCaptureMode === "immediate"}
                          data-quick-ocr-mode-option="immediate"
                        >
                          <strong>{isEn ? "Immediate selection" : "立即框选"}</strong>
                          <span>{isEn ? "Start the macOS selection tool right away" : "点击后立即进入 macOS 框选截图"}</span>
                        </button>
                        <button
                          type="button"
                          className={quickOcrCaptureMode === "system-screenshot" ? "is-active" : ""}
                          onClick={() => {
                            onQuickOcrCaptureModeChange("system-screenshot");
                            setQuickOcrModeMenuOpen(false);
                          }}
                          role="menuitemradio"
                          aria-checked={quickOcrCaptureMode === "system-screenshot"}
                          data-quick-ocr-mode-option="system-screenshot"
                        >
                          <strong>{isEn ? "Wait for system screenshot" : "等待系统截图"}</strong>
                          <span>{isEn ? "Switch pages first, then use ⌘⇧3 / 4 / 5" : "先切到目标页面，再使用 ⌘⇧3 / 4 / 5"}</span>
                        </button>
                      </div>
                    )}
                  </div>
                  {onSilentOcrEnabledChange && (
                    <label
                      className={`silent-ocr-toggle${silentOcrEnabled ? " is-active" : ""}`}
                      title={
                        isEn
                          ? "When enabled, press ⌘⇧O anywhere to capture, recognize, and copy LaTeX without opening the main window"
                          : "开启后可在任意应用中按 ⌘⇧O 框选截图，后台识别并复制 LaTeX，无需打开主窗口"
                      }
                    >
                      <input
                        type="checkbox"
                        checked={silentOcrEnabled}
                        onChange={(event) =>
                          onSilentOcrEnabledChange(event.target.checked)
                        }
                        data-silent-ocr-toggle
                      />
                      <span className="silent-ocr-indicator" aria-hidden="true" />
                      <span>{isEn ? "Silent" : "静默"}</span>
                      <kbd>⌘⇧O</kbd>
                    </label>
                  )}
                </div>
              )}
              {!keypadMode && showOcrActions && ocrModels.length > 0 && ocrModel && (
                <label
                  className="canvas-ocr-model"
                  title={
                    isEn
                      ? "Model used when an image is pasted into a formula field"
                      : "在公式输入框中粘贴图片时使用的 OCR 模型"
                  }
                >
                  <ScanLine size={14} />
                  <select
                    value={ocrModel}
                    disabled={ocrBusy}
                    onChange={(event) =>
                      onOcrModelChange?.(event.target.value)
                    }
                    aria-label={isEn ? "OCR recognition model" : "OCR 识别模型"}
                  >
                    {ocrModels.map((item) => (
                      <option key={item.id} value={item.id}>
                        {isEn ? item.labelEn : item.labelZh}
                      </option>
                    ))}
                  </select>
                </label>
              )}
              {!keypadMode && <div className="canvas-controls">
                <button
                  type="button"
                  className="icon-button compact"
                  onClick={() => setZoom(zoom - EDITOR_ZOOM_STEP)}
                  disabled={zoom <= MIN_EDITOR_ZOOM + 0.0001}
                  aria-label={isEn ? "Zoom out" : "缩小公式"}
                  title={
                    zoom <= MIN_EDITOR_ZOOM + 0.0001
                      ? isEn
                        ? "Minimum zoom: 20%"
                        : "最小缩放：20%"
                      : undefined
                  }
                >
                  <Minus size={14} />
                </button>
                <span aria-live="polite" aria-atomic="true">
                  {Math.round(zoom * 100)}%
                </span>
                <button
                  type="button"
                  className="icon-button compact"
                  onClick={() => setZoom(zoom + EDITOR_ZOOM_STEP)}
                  disabled={zoom >= MAX_EDITOR_ZOOM - 0.0001}
                  aria-label={isEn ? "Zoom in" : "放大公式"}
                  title={
                    zoom >= MAX_EDITOR_ZOOM - 0.0001
                      ? isEn
                        ? "Maximum zoom: 160%"
                        : "最大缩放：160%"
                      : undefined
                  }
                >
                  <Plus size={14} />
                </button>
              </div>}
            </div>
            {isOfficeWorkspace && officeHeaderTrailingActions ? (
              <div className="office-inline-actions">
                {officeHeaderTrailingActions}
              </div>
            ) : null}
          </header>

          {keypadMode ? (
            <div className="keypad-editor-pane-body">
              <div className="editor-pane-scroll">
                {renderVisualEditor()}
              </div>
            </div>
          ) : editorLayout === "classic" ? (
            <div
              ref={classicEditorBodyRef}
              className={
                "classic-editor-pane-body" +
                (classicDockOpen ? "" : " is-dock-collapsed")
              }
              style={
                {
                  "--classic-dock-height": `${classicDockHeight}px`,
                } as CSSProperties
              }
            >
              <div className="editor-pane-scroll">
                {renderVisualEditor()}
              </div>

              {classicDockOpen && (
                <div
                  className="workspace-panel-resizer classic-dock-resizer"
                  role="separator"
                  tabIndex={0}
                  aria-orientation="horizontal"
                  aria-valuemin={MIN_CLASSIC_DOCK_HEIGHT}
                  aria-valuemax={Math.round(classicDockHeightLimit())}
                  aria-valuenow={Math.round(classicDockHeight)}
                  aria-label={
                    isEn
                      ? "Resize formula tools and source panel"
                      : "调整公式工具与源码区高度"
                  }
                  title={
                    isEn
                      ? "Drag to resize · Double-click to reset"
                      : "拖动调整高度 · 双击恢复默认"
                  }
                  onPointerDown={(event) => startClassicResize("dock", event)}
                  onDoubleClick={() => resetClassicPanelSize("dock")}
                  onKeyDown={(event) => {
                    if (event.key === "ArrowUp") {
                      event.preventDefault();
                      adjustClassicPanelFromKeyboard("dock", 16);
                    } else if (event.key === "ArrowDown") {
                      event.preventDefault();
                      adjustClassicPanelFromKeyboard("dock", -16);
                    } else if (event.key === "Home") {
                      event.preventDefault();
                      commitClassicPanelSize(
                        "dock",
                        MIN_CLASSIC_DOCK_HEIGHT,
                        true,
                      );
                    } else if (event.key === "End") {
                      event.preventDefault();
                      commitClassicPanelSize(
                        "dock",
                        classicDockHeightLimit(),
                        true,
                      );
                    }
                  }}
                />
              )}

              <section
                className={
                  "classic-bottom-dock" +
                  (classicDockOpen ? "" : " is-collapsed") +
                  (sourceOpen ? " is-source-panel" : " is-tools-panel")
                }
                aria-label={
                  isEn ? "Formula tools and LaTeX source" : "公式工具与 LaTeX 源码"
                }
              >
                <nav
                  className="classic-bottom-tabs"
                  aria-label={isEn ? "Bottom editor panel" : "底部编辑面板"}
                >
                  {isOfficeWorkspace ? (
                    <div
                      ref={setOfficeFormattingMount}
                      className="classic-bottom-formatting-slot"
                      aria-label={
                        isEn
                          ? "Formula alignment and formatting"
                          : "公式对齐与格式"
                      }
                    />
                  ) : (
                    <span
                      className="classic-bottom-tab-spacer"
                      aria-hidden="true"
                    />
                  )}
                  <div
                    className="classic-bottom-tab-group"
                    role="tablist"
                    aria-label={isEn ? "Bottom editor panel" : "底部编辑面板"}
                  >
                    <button
                      type="button"
                      role="tab"
                      className={!sourceOpen ? "is-active" : ""}
                      aria-selected={!sourceOpen}
                      data-classic-bottom-view="tools"
                      onClick={() => {
                        setSourceOpen(false);
                        setClassicDockOpen(true);
                      }}
                    >
                      <Braces size={16} />
                      <span className="classic-bottom-tab-label">
                        {isEn ? "Formula tools" : "公式工具"}
                      </span>
                    </button>
                    <button
                      type="button"
                      role="tab"
                      className={sourceOpen ? "is-active" : ""}
                      aria-selected={sourceOpen}
                      data-classic-bottom-view="source"
                      onClick={() => {
                        setSourceOpen(true);
                        setClassicDockOpen(true);
                      }}
                    >
                      <Code2 size={16} />
                      <span className="classic-bottom-tab-label">
                        {isEn ? "LaTeX source" : "LaTeX 源码"}
                      </span>
                    </button>
                  </div>
                  <div className="classic-bottom-actions">
                    {sourceOpen && classicDockOpen && (
                      <button
                        type="button"
                        className="icon-button compact classic-bottom-copy"
                        data-classic-bottom-copy
                        onClick={() => void onCopy()}
                        aria-label={isEn ? "Copy LaTeX source" : "复制 LaTeX 源码"}
                        title={isEn ? "Copy LaTeX source" : "复制 LaTeX 源码"}
                      >
                        <Copy size={14} />
                      </button>
                    )}
                    <button
                      type="button"
                      className="icon-button compact classic-bottom-collapse"
                      data-classic-bottom-collapse
                      aria-expanded={classicDockOpen}
                      aria-label={
                        classicDockOpen
                          ? isEn
                            ? "Collapse formula tools and source"
                            : "收起公式工具与源码"
                          : isEn
                            ? "Expand formula tools and source"
                            : "展开公式工具与源码"
                      }
                      title={
                        classicDockOpen
                          ? isEn
                            ? "Collapse bottom panel"
                            : "收起底部面板"
                          : isEn
                            ? "Expand bottom panel"
                            : "展开底部面板"
                      }
                      onClick={() => setClassicDockOpen((open) => !open)}
                    >
                      {classicDockOpen ? (
                        <PanelBottomClose size={14} />
                      ) : (
                        <PanelBottomOpen size={14} />
                      )}
                    </button>
                  </div>
                </nav>
                {classicDockOpen && (
                  <div className="classic-bottom-content">
                    {sourceOpen ? (
                      <div className="source-pane-slot classic-source-pane-slot">
                        {renderSourceEditor({
                          showCollapseAction: false,
                          showCopyAction: false,
                          compact: true,
                        })}
                      </div>
                    ) : (
                      <FormulaToolbar
                        view="tools"
                        layout="horizontal"
                        className="classic-bottom-toolbar"
                        onInsert={(command) =>
                          editorRef.current?.insertCommand(command)
                        }
                      />
                    )}
                  </div>
                )}
              </section>
            </div>
          ) : (
            <div className={`editor-pane-body${sourceOpen ? " has-source" : ""}`}>
              <div className="editor-pane-scroll">
                {renderVisualEditor()}
              </div>

              {sourceOpen ? (
                <div className="source-pane-slot">{renderSourceEditor()}</div>
              ) : (
                <div className="source-toggle-row">
                  <span className="source-toggle-label" aria-hidden="true">
                    <Code2 size={15} />
                  </span>
                  <button
                    type="button"
                    className="source-toggle"
                    onClick={() => setSourceOpen(true)}
                    aria-label={isEn ? "Show LaTeX source" : "展开 LaTeX 源码"}
                    title={isEn ? "Show LaTeX source" : "展开 LaTeX 源码"}
                  >
                    <PanelBottomOpen size={15} />
                  </button>
                </div>
              )}
            </div>
          )}
        </section>

        {!keypadMode && editorLayout === "classic" && sidebarOpen && (
          <>
            <div
              className="workspace-panel-resizer classic-tile-resizer"
              role="separator"
              tabIndex={0}
              aria-orientation="vertical"
              aria-valuemin={MIN_CLASSIC_TILE_WIDTH}
              aria-valuemax={Math.round(classicTileWidthLimit())}
              aria-valuenow={Math.round(classicTileWidth)}
              aria-label={isEn ? "Resize formula tiles" : "调整公式磁贴区宽度"}
              title={
                isEn
                  ? "Drag to resize · Double-click to reset"
                  : "拖动调整宽度 · 双击恢复默认"
              }
              onPointerDown={(event) => startClassicResize("tiles", event)}
              onDoubleClick={() => resetClassicPanelSize("tiles")}
              onKeyDown={(event) => {
                if (event.key === "ArrowLeft") {
                  event.preventDefault();
                  adjustClassicPanelFromKeyboard("tiles", 16);
                } else if (event.key === "ArrowRight") {
                  event.preventDefault();
                  adjustClassicPanelFromKeyboard("tiles", -16);
                } else if (event.key === "Home") {
                  event.preventDefault();
                  commitClassicPanelSize(
                    "tiles",
                    MIN_CLASSIC_TILE_WIDTH,
                    true,
                  );
                } else if (event.key === "End") {
                  event.preventDefault();
                  commitClassicPanelSize(
                    "tiles",
                    classicTileWidthLimit(),
                    true,
                  );
                }
              }}
            />
            <FormulaToolbar
              view="tiles"
              className="classic-tile-toolbar"
              stabilizeTileLayout
              onCollapseTiles={() => onSidebarOpenChange(false)}
              onInsert={(command) => editorRef.current?.insertCommand(command)}
            />
          </>
        )}
      </main>

      <footer className="status-bar">
        <div>
          <span className="status-live-dot" />
          {isEn ? "Ready" : "就绪"}
        </div>
        <div>
          <span>
            {lines.length} {isEn ? "lines" : "行"}
          </span>
          <span>
            · {latex.length} {isEn ? "characters" : "字符"}
          </span>
        </div>
      </footer>
    </>
  );
}
