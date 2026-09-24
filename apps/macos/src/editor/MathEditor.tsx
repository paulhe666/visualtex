import { formatLatexLines } from "../clipboard/LatexCopyService";
import {
  useEffect,
  useLayoutEffect,
  useImperativeHandle,
  useRef,
  useState,
  forwardRef,
  Component,
  type CSSProperties,
  type ErrorInfo,
  type ReactNode,
} from "react";
import {
  MathfieldElement,
  type Style,
} from "mathlive";
import { flushSync } from "react-dom";
import { ClipboardCopy } from "lucide-react";
import type {
  CommandSource,
  LatexCommand,
} from "../types/command";
import type {
  FormulaAlignment,
  FormulaLine,
  InputBehaviorSettingKey,
  LatexCodeFormat,
  InputBehaviorSettings,
} from "../types/formula";
import type {
  AddLineEntry,
  EditKind,
  FormulaEditSource,
  MathSelectionSnapshot,
  RemoveLineEntry,
  ReplaceDocumentEntry,
  ReplaceFormulaEntry,
} from "../history/historyTypes";
import {
  clampSelection,
  historyManager,
} from "../history/HistoryManager";
import { getEditorDocumentSnapshot } from "../history/documentHistory";
import {
  findRuntimeCommandByCommand,
} from "../autocomplete/runtimeCommandRegistry";
import {
  compatibilityRawPlaceholderTemplates,
  compatibilityWrapperPreviews,
} from "../autocomplete/compatibilityCommands";
import {
  createFormulaLine,
  useEditorStore,
} from "../stores/editorStore";
import { useFormulaHotkeyStore } from "../stores/formulaHotkeyStore";
import {
  matchFormulaHotkey,
  resolveFormulaHotkeyCommand,
} from "../shortcuts/formulaHotkeys";
import {
  greekLetterHotkeyCommandFromEvent,
  isGreekLetterHotkeyPrefix,
} from "../shortcuts/greekLetterHotkeys";
import {
  normalizeChineseLatex,
  normalizeContextualUprightSymbols,
  normalizeMathLiveCanonicalUprightCommands,
  resolveVisualTexInlineShortcuts,
  type VisualTexInlineShortcutDefinitions,
} from "./normalizeChineseLatex";
import {
  convertVisualTexLatexToMarkup,
  installMathLiveContourIntegralShadowStyle,
} from "./mathLiveIntegralCompatibility";
import {
  inspectMathLiveSourceSafety,
  mathLiveSourceSafetyMessage,
} from "./mathLiveSourceSafety";
import { composeCustomSymbolMacrosForMathfield } from "../math/customSymbolRegistry";
import { isSingleCompleteLatexEnvironment } from "../math/latexEnvironment";
import {
  installCustomSymbolGlobalStyle,
  installCustomSymbolShadowStyle,
  refreshCustomSymbolMathfield,
} from "../math/customSymbolRendering";
import { useCustomSymbolRevision } from "../math/customSymbolReact";
import { ImeCompositionGuard } from "./imeCompositionGuard";
import { installMathLiveOptionMutationGuard } from "./mathLiveOptionCompatibility";
import { VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS } from "../math/mathLiveCompatibilityMacros";
import { hasBoundedOperatorPlaceholderOrder } from "./boundedOperatorTemplate";
import {
  configureMathLiveCompletion,
  dismissMathLiveSuggestions,
  insertMathLiveAlignmentPoint,
  insertMathLiveRowBreak,
  setMathLivePersistentTypingStyle,
  toggleMathLiveSelectionStyle,
  type MathLivePersistentTypingStyle,
} from "./mathLiveKernelCommands";
import { isSafeFormulaStyleColor } from "../workspace/formulaColor";
import { copyFormulaDocumentPngToClipboard } from "../export/pngClipboard";
import {
  formulaChineseFontFamily,
  formulaLetterFontFamilies,
  markVisualTexFormulaFontGlyphs,
  VISUALTEX_FORMULA_CHINESE_GLYPH_CLASS,
  VISUALTEX_FORMULA_LETTER_GLYPH_CLASS,
  type FormulaChineseFont,
  type FormulaLetterFont,
} from "./formulaFontPreferences";
import {
  hasVisualTexAlignmentMarker,
  VISUALTEX_ALIGNMENT_MARKER_CLASS,
  VISUALTEX_ALIGNMENT_MARKER_LATEX,
} from "./alignmentMarkers";

export interface MathEditorInsertionTarget {
  lineId: string;
  ranges: Array<[number, number]>;
  direction: "forward" | "backward" | "none";
}

export interface MathEditorSelectionTarget {
  selections: MathEditorInsertionTarget[];
}

export type MathEditorSelectionStyle =
  | { kind: "bold" }
  | { kind: "italic" }
  | { kind: "color"; value: string }
  | { kind: "backgroundColor"; value: string };

export interface MathEditorFocusOptions {
  target?: "active" | "first" | "last";
  moveToEnd?: boolean;
}

interface FormulaFieldErrorBoundaryProps {
  recoveryKey: string;
  children: ReactNode;
}

interface FormulaFieldErrorBoundaryState {
  error: unknown | null;
  recoveryKey: string;
}

function FormulaFieldRenderFallback({ message }: { message: string }) {
  return (
    <div
      role="status"
      className="formula-field-render-fallback"
      style={{
        minHeight: 42,
        boxSizing: "border-box",
        display: "flex",
        alignItems: "center",
        padding: "8px 12px",
        border: "1px solid rgba(190, 18, 60, 0.28)",
        borderRadius: 6,
        color: "#9f1239",
        background: "rgba(255, 241, 242, 0.72)",
        fontSize: 12,
      }}
    >
      {message}
    </div>
  );
}

class FormulaFieldErrorBoundary extends Component<
  FormulaFieldErrorBoundaryProps,
  FormulaFieldErrorBoundaryState
> {
  state: FormulaFieldErrorBoundaryState = {
    error: null,
    recoveryKey: this.props.recoveryKey,
  };

  static getDerivedStateFromError(error: unknown) {
    return { error };
  }

  static getDerivedStateFromProps(
    props: FormulaFieldErrorBoundaryProps,
    state: FormulaFieldErrorBoundaryState,
  ) {
    if (props.recoveryKey === state.recoveryKey) return null;
    return { error: null, recoveryKey: props.recoveryKey };
  }

  componentDidCatch(error: unknown, info: ErrorInfo) {
    console.error("VisualTeX isolated a formula renderer failure", error, info);
  }

  render() {
    if (!this.state.error) return this.props.children;
    return (
      <FormulaFieldRenderFallback message="该行公式渲染失败，源码已保留。可在 LaTeX 源码区修改或删除这一行。" />
    );
  }
}

const EDITOR_LAYOUT_REFRESH_EVENT = "visualtex-editor-layout-refresh";
const VISUALTEX_MULTILINE_LATEX_CLIPBOARD_TYPE =
  "application/x-visualtex-multiline-latex";

function safeConvertVisualTexLatexToMarkup(
  ...args: Parameters<typeof convertVisualTexLatexToMarkup>
) {
  const source = args[0];
  if (typeof source === "string") {
    const issue = inspectMathLiveSourceSafety(source);
    if (issue) {
      console.warn(
        "VisualTeX skipped a structurally unsafe auxiliary MathLive render.",
        { sourceLength: source.length, issue },
      );
      return "";
    }
  }
  try {
    return convertVisualTexLatexToMarkup(...args);
  } catch (error) {
    console.warn(
      "VisualTeX skipped a non-essential MathLive markup render after an exception.",
      {
        sourceLength: typeof source === "string" ? source.length : 0,
        error,
      },
    );
    return "";
  }
}

export interface MathEditorHandle {
  insertCommand: (command: LatexCommand, source?: "toolbar" | "history" | "shortcut") => void;
  insertLatex: (latex: string, source?: FormulaEditSource) => void;
  insertLatexAt: (
    target: MathEditorInsertionTarget,
    latex: string,
    source?: FormulaEditSource,
  ) => boolean;
  appendLatex: (latex: string, source?: FormulaEditSource) => void;
  focus: (options?: MathEditorFocusOptions) => void;
  addLine: () => void;
  commitPendingTransaction: () => void;
  getSelectionMap: () => Record<string, MathSelectionSnapshot>;
  restoreSelection: (
    lineId: string,
    latex: string,
    selection: MathSelectionSnapshot | null,
  ) => Promise<boolean>;
  captureSelectionTarget: () => MathEditorSelectionTarget | null;
  captureInsertionTarget: () => MathEditorInsertionTarget | null;
  applySelectionStyle: (
    style: MathEditorSelectionStyle,
    target?: MathEditorSelectionTarget | null,
  ) => boolean;
  refreshLayout: () => void;
}

interface Props {
  lines: FormulaLine[];
  activeLineId: string | null;
  formulaAlignment: FormulaAlignment;
  latexCodeFormat: LatexCodeFormat;
  zoom: number;
  persistentTypingStyle?: MathLivePersistentTypingStyle;
  reuseLineSlots?: boolean;
  readOnly?: boolean;
  previewOnly?: boolean;
  showLineModeControls?: boolean;
  onPreviewActivate?: () => void;
  draftError?: string;
  onPasteImage?: (file: File, target: MathEditorInsertionTarget) => void;
  onCopyPng?: () => Promise<void>;
  onHistoryBusyChange?: (busy: boolean) => void;
  overlay?: ReactNode;
}

interface FormulaFieldEdit {
  lineId: string;
  beforeLatex: string;
  afterLatex: string;
  beforeSelection: MathSelectionSnapshot;
  afterSelection: MathSelectionSnapshot;
  editKind: EditKind;
  source: FormulaEditSource;
}

interface FormulaFieldProps {
  lineId: string;
  index: number;
  latex: string;
  zoom: number;
  formulaRowVerticalInset: number;
  language: "cn" | "en";
  formulaLetterFont: FormulaLetterFont;
  formulaChineseFont: FormulaChineseFont;
  autoPairDelimiters: boolean;
  inputBehavior: InputBehaviorSettings;
  persistentTypingStyle: MathLivePersistentTypingStyle;
  readOnly: boolean;
  freshExternalSync: boolean;
  register: (
    lineId: string,
    field: MathfieldElement | null,
    expectedField?: MathfieldElement,
  ) => void;
  onEdit: (edit: FormulaFieldEdit, field: MathfieldElement) => void;
  onInputActivity: (field: MathfieldElement) => void;
  onSelectionChange: (
    lineId: string,
    selection: MathSelectionSnapshot,
  ) => void;
  onFocus: (index: number, field: MathfieldElement) => void;
  onCommitPending: () => void;
  onKeyDown: (index: number, event: KeyboardEvent, field: MathfieldElement) => void;
  onPasteImage?: (file: File, target: MathEditorInsertionTarget) => void;
  onCopyPng?: () => Promise<void>;
  onContextMenu?: (clientX: number, clientY: number) => void;
  onPasteLatexLines?: (
    lineId: string,
    field: MathfieldElement,
    lines: string[],
  ) => void;
}

interface MultiLineSelectionPoint {
  lineId: string;
  lineIndex: number;
  offset: number;
}

interface MultiLineSelectionState {
  anchor: MultiLineSelectionPoint;
  focus: MultiLineSelectionPoint;
}

interface PointerSelectionSession {
  pointerId: number;
  startX: number;
  startY: number;
  anchor: MultiLineSelectionPoint;
  allowSameLine: boolean;
  active: boolean;
}

const trailingCommand = /\\([\p{L}]*)$/u;

function hasRawLatexInput(field: MathfieldElement) {
  return Boolean(field.shadowRoot?.querySelector(".ML__raw-latex"));
}

function rawLatexInput(field: MathfieldElement) {
  return Array.from(
    field.shadowRoot?.querySelectorAll<HTMLElement>(".ML__raw-latex") ?? [],
  )
    .filter((node) => !node.classList.contains("ML__suggestion"))
    .map((node) => node.textContent ?? "")
    .join("");
}

// Opt-in diagnostics for real WKWebView input transactions. Keep the trace in
// the accessibility tree so the macOS probe can read it from a packaged or
// development Tauri app without relying on Web Inspector/CDP.
const VISUALTEX_IME_DIAGNOSTIC_LABEL = "VisualTeX IME Diagnostic Trace";
const VISUALTEX_IME_DIAGNOSTICS_ENABLED =
  import.meta.env.VITE_VISUALTEX_IME_DIAGNOSTICS === "1";
const visualTexImeDiagnosticEntries: string[] = [];
const visualTexImeDiagnosticEventIds = new WeakMap<Event, number>();
let visualTexImeDiagnosticSequence = 0;
let visualTexImeDiagnosticEventSequence = 0;

function visualTexImeDiagnosticEventId(event: Event) {
  const existing = visualTexImeDiagnosticEventIds.get(event);
  if (existing) return existing;
  const next = ++visualTexImeDiagnosticEventSequence;
  visualTexImeDiagnosticEventIds.set(event, next);
  return next;
}

function ensureVisualTexImeDiagnosticElement() {
  let element = document.querySelector<HTMLTextAreaElement>(
    `textarea[aria-label="${VISUALTEX_IME_DIAGNOSTIC_LABEL}"]`,
  );
  if (element) return element;
  element = document.createElement("textarea");
  element.readOnly = true;
  element.tabIndex = -1;
  element.setAttribute("aria-label", VISUALTEX_IME_DIAGNOSTIC_LABEL);
  Object.assign(element.style, {
    position: "fixed",
    left: "2px",
    bottom: "2px",
    width: "1px",
    height: "1px",
    opacity: "0.01",
    pointerEvents: "none",
    zIndex: "-1",
  });
  document.body.append(element);
  return element;
}

function traceVisualTexImeEvent(
  stage: string,
  field: MathfieldElement,
  keyboardSink: HTMLElement | null,
  event: Event,
) {
  if (!VISUALTEX_IME_DIAGNOSTICS_ENABLED) return;
  const keyboard = event instanceof KeyboardEvent ? event : null;
  const input = event instanceof InputEvent ? event : null;
  const composition = event instanceof CompositionEvent ? event : null;
  let raw = "";
  try {
    raw = rawLatexInput(field);
  } catch {
    raw = "";
  }
  const activeElement = field.shadowRoot?.activeElement;
  const entry = {
    seq: ++visualTexImeDiagnosticSequence,
    stage,
    eventId: visualTexImeDiagnosticEventId(event),
    timeStamp: event.timeStamp,
    now: performance.now(),
    type: event.type,
    eventPhase: event.eventPhase,
    composed: event.composed,
    bubbles: event.bubbles,
    cancelable: event.cancelable,
    key: keyboard?.key ?? "",
    code: keyboard?.code ?? "",
    keyCode: keyboard?.keyCode ?? 0,
    repeat: keyboard?.repeat ?? false,
    isComposing: keyboard?.isComposing ?? input?.isComposing ?? false,
    inputType: input?.inputType ?? "",
    data: input?.data ?? composition?.data ?? null,
    defaultPrevented: event.defaultPrevented,
    ctrlKey: keyboard?.ctrlKey ?? false,
    metaKey: keyboard?.metaKey ?? false,
    altKey: keyboard?.altKey ?? false,
    shiftKey: keyboard?.shiftKey ?? false,
    lineId: field.dataset.visualtexLineId ?? "",
    value: field.value,
    raw,
    mode: field.mode,
    position: field.position,
    sinkText: keyboardSink?.textContent ?? "",
    activePart:
      activeElement instanceof HTMLElement
        ? activeElement.getAttribute("part") ?? activeElement.tagName
        : "",
  };
  visualTexImeDiagnosticEntries.push(JSON.stringify(entry));
  if (visualTexImeDiagnosticEntries.length > 800) {
    visualTexImeDiagnosticEntries.splice(
      0,
      visualTexImeDiagnosticEntries.length - 800,
    );
  }
  ensureVisualTexImeDiagnosticElement().value =
    visualTexImeDiagnosticEntries.join("\n");
}

type MathLiveInternalField = {
  _mathfield?: {
    model?: {
      root?: unknown;
      position?: number;
      parentEnvironment?: {
        environmentName?: string;
      } | null;
    };
  };
};

function resetMathLiveModelRootForExternalSync(field: MathfieldElement) {
  const targetModel = (field as unknown as MathLiveInternalField)._mathfield?.model;
  if (!targetModel || typeof document === "undefined" || !document.body) {
    return false;
  }

  const previouslyFocused = document.activeElement as HTMLElement | null;
  const stagingHost = document.createElement("div");
  stagingHost.setAttribute("aria-hidden", "true");
  stagingHost.style.cssText =
    "position:fixed;left:-100000px;top:0;width:1px;height:1px;overflow:hidden;visibility:hidden;pointer-events:none";
  const freshField = new MathfieldElement();
  freshField.readOnly = true;
  freshField.tabIndex = -1;

  try {
    document.body.append(stagingHost);
    stagingHost.append(freshField);
    freshField.setValue("", {
      mode: "math",
      format: "latex",
      insertionMode: "replaceAll",
      selectionMode: "after",
      silenceNotifications: true,
    });
    const freshModel = (freshField as unknown as MathLiveInternalField)._mathfield
      ?.model;
    if (!freshModel || freshModel.root === undefined) return false;
    targetModel.root = freshModel.root;
    targetModel.position = 0;
    return true;
  } finally {
    stagingHost.remove();
    if (
      previouslyFocused?.isConnected &&
      document.activeElement !== previouslyFocused
    ) {
      previouslyFocused.focus({ preventScroll: true });
    }
  }
}

const structuralExternalSyncPattern = /\\(?:begin\s*\{|left\b|right\b)/;

function needsCleanMathLiveModelForExternalSync(
  currentLatex: string,
  nextLatex: string,
) {
  return (
    structuralExternalSyncPattern.test(currentLatex) ||
    structuralExternalSyncPattern.test(nextLatex)
  );
}

function activeMathLiveEnvironmentName(field: MathfieldElement) {
  return (
    (field as unknown as MathLiveInternalField)._mathfield?.model?.parentEnvironment
      ?.environmentName ?? null
  );
}

function structuralBoundaryOffsetFromPoint(
  field: MathfieldElement,
  clientX: number,
  clientY: number,
) {
  const elementInfo = Array.from(
    { length: field.lastOffset + 1 },
    (_, offset) => ({ offset, info: field.getElementInfo(offset) }),
  );
  // This correction exists only for bounds-less top-level model offsets in the
  // horizontal gap between rendered structures. If the pointer is actually on
  // any MathLive atom, let MathLive's native two-dimensional hit testing own the
  // selection. In particular this preserves PlaceholderAtom's built-in
  // setSelection(anchor - 1, anchor) behavior.
  const hitsRenderedAtom = elementInfo.some(({ info }) => {
    const bounds = info?.bounds;
    return Boolean(
      bounds &&
        clientX >= bounds.left &&
        clientX <= bounds.right &&
        clientY >= bounds.top &&
        clientY <= bounds.bottom,
    );
  });
  if (hitsRenderedAtom) return null;
  const nextTopLevelBounds = new Array<DOMRect | undefined>(
    elementInfo.length,
  );
  let upcomingTopLevelBounds: DOMRect | undefined;
  for (let index = elementInfo.length - 1; index >= 0; index -= 1) {
    nextTopLevelBounds[index] = upcomingTopLevelBounds;
    const info = elementInfo[index]?.info;
    if (info?.depth === 0 && info.bounds) {
      upcomingTopLevelBounds = info.bounds;
    }
  }

  let previousBounds: DOMRect | undefined;
  for (const candidate of elementInfo) {
    // MathLive's point hit testing uses two-dimensional distance. In the
    // horizontal gap after a scripted expression this can select a superscript
    // or subscript instead of the real top-level boundary represented by this
    // bounds-less model offset. Restrict the correction to the actual gap so
    // clicks on either neighboring structure keep their native behavior.
    if (candidate.info?.depth === 0 && !candidate.info.bounds) {
      const next = nextTopLevelBounds[candidate.offset];
      if (
        previousBounds &&
        next &&
        previousBounds.right <= next.left &&
        clientX >= previousBounds.right &&
        clientX <= next.left
      ) {
        return candidate.offset;
      }
    }

    if (candidate.info?.bounds) previousBounds = candidate.info.bounds;
  }

  return null;
}

function rawCommandQuery(field: MathfieldElement) {
  if (!hasRawLatexInput(field)) return "";
  const match = rawLatexInput(field).match(trailingCommand);
  return match && match[1].length > 0 ? "\\" + match[1] : "";
}

function trailingCommandQuery(
  field: MathfieldElement,
  normalizedValue = field.value,
) {
  if (hasRawLatexInput(field)) return rawCommandQuery(field);
  for (const source of [normalizedValue, field.value]) {
    const match = source.match(trailingCommand);
    if (match) return "\\" + match[1];
  }
  return "";
}

interface WrapperCaretAnchor {
  left: number;
  centerY: number;
  height: number;
}

interface EditorScrollSnapshot {
  scroller: HTMLElement;
  anchorLineId: string | null;
  anchorOffset: number;
  scrollTop: number;
}

type VerticalPlaceholderKind =
  | "fraction"
  | "operator-limit"
  | "script"
  | "stack";

type VerticalPlaceholderAnchor = {
  kind: VerticalPlaceholderKind;
  region: "upper" | "lower";
  centerX: number;
  centerY: number;
  relativeX: number;
  relativeY: number;
};

interface RawCommandAnchor {
  latex: string;
  selection: MathSelectionSnapshot;
  position: number;
  visualCaret: WrapperCaretAnchor | null;
  selectedPlaceholder: boolean;
  selectedPlaceholderIndex: number | null;
  selectedPlaceholderTextIndex: number | null;
  verticalPlaceholder: VerticalPlaceholderAnchor | null;
  autoExitSetting: InputBehaviorSettingKey | null;
  autoExitScriptKey: string | null;
}

const rawCommandAnchors = new WeakMap<MathfieldElement, RawCommandAnchor>();

function latexPlaceholderRanges(field: MathfieldElement) {
  const ranges: Array<[number, number]> = [];
  for (let offset = 1; offset <= field.lastOffset; offset += 1) {
    if (
      field.getValue(offset - 1, offset, "latex").trim() ===
      "\\placeholder{}"
    ) {
      ranges.push([offset - 1, offset]);
    }
  }
  return ranges;
}

const placeholderMarkerCharacters =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

function markedPlaceholderLatex(latex: string) {
  let index = 0;
  const markers: string[] = [];
  const markedLatex = latex.replace(/\\placeholder\{\}/g, () => {
    const marker = placeholderMarkerCharacters[index] ?? "Q";
    markers.push(marker);
    index += 1;
    return marker;
  });
  return { markedLatex, markers };
}

function textualPlaceholderIndexForRange(
  latex: string,
  range: [number, number],
) {
  const { markedLatex, markers } = markedPlaceholderLatex(latex);
  if (!markers.length) return null;
  const verifier = new MathfieldElement();
  verifier.setValue(markedLatex, {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "after",
    silenceNotifications: true,
  });
  const marker = verifier
    .getValue(
      Math.min(range[0], range[1]),
      Math.max(range[0], range[1]),
      "latex",
    )
    .trim();
  const index = markers.indexOf(marker);
  return index >= 0 ? index : null;
}

function replaceTextualPlaceholder(
  latex: string,
  targetIndex: number,
  replacement: string,
) {
  let index = 0;
  let replaced = false;
  const value = latex.replace(/\\placeholder\{\}/g, (placeholder) => {
    if (index !== targetIndex) {
      index += 1;
      return placeholder;
    }
    index += 1;
    replaced = true;
    return replacement;
  });
  return replaced ? value : null;
}

function modelRangeForTextualPlaceholder(
  field: MathfieldElement,
  latex: string,
  targetIndex: number,
) {
  const { markedLatex, markers } = markedPlaceholderLatex(latex);
  const targetMarker = markers[targetIndex];
  if (!targetMarker) return null;
  const verifier = new MathfieldElement();
  verifier.setValue(markedLatex, {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "after",
    silenceNotifications: true,
  });
  return (
    latexPlaceholderRanges(field).find(([start, end]) =>
      verifier.getValue(start, end, "latex").trim() === targetMarker,
    ) ?? null
  );
}

function captureWrapperCaretAnchor(
  field: MathfieldElement,
): WrapperCaretAnchor | null {
  const host = field.closest<HTMLElement>(".mathfield-host");
  if (!host) return null;
  const hostBounds = host.getBoundingClientRect();
  const candidateOffsets = Array.from(
    new Set(
      [field.position, field.position - 1, field.position + 1].filter(
        (offset) => offset >= 0 && offset <= field.lastOffset,
      ),
    ),
  );
  const modelAnchors = candidateOffsets
    .flatMap((offset) => {
      const bounds = field.getElementInfo(offset)?.bounds;
      if (
        !bounds ||
        !Number.isFinite(bounds.right) ||
        !Number.isFinite(bounds.top) ||
        bounds.height <= 0
      ) {
        return [];
      }
      return [{
        left: bounds.right - hostBounds.left,
        centerY: bounds.top - hostBounds.top + bounds.height / 2,
        height: bounds.height,
      }];
    });
  const modelAnchor = modelAnchors[0] ?? null;
  const markerAnchors = Array.from(
    field.shadowRoot?.querySelectorAll<HTMLElement>(
      ".visualtex-structural-placeholder-caret, .ML__caret, .ML__text-caret, .ML__latex-caret",
    ) ?? [],
  )
    .flatMap((marker) => {
      const bounds = marker.getBoundingClientRect();
      const style = getComputedStyle(marker);
      const overlapsHost =
        bounds.right >= hostBounds.left &&
        bounds.left <= hostBounds.right &&
        bounds.bottom >= hostBounds.top &&
        bounds.top <= hostBounds.bottom;
      if (
        bounds.height <= 0 ||
        style.display === "none" ||
        style.visibility === "hidden" ||
        !overlapsHost
      ) {
        return [];
      }
      return [{
        left: bounds.left - hostBounds.left,
        centerY: bounds.top - hostBounds.top + bounds.height / 2,
        height: bounds.height,
        width: bounds.width,
        priority: marker.classList.contains(
          "visualtex-structural-placeholder-caret",
        )
          ? 0
          : marker.classList.contains("ML__caret")
            ? 1
            : 2,
      }];
    })
    .sort(
      (first, second) =>
        first.priority - second.priority ||
        first.width - second.width ||
        first.height - second.height,
    );
  const markerAnchor = markerAnchors[0] ?? null;
  if (markerAnchor) {
    return {
      left: markerAnchor.left,
      centerY: markerAnchor.centerY,
      height: markerAnchor.height,
    };
  }
  return modelAnchor;
}

function visiblePlaceholderNodes(
  field: MathfieldElement,
  scope: ParentNode = field.shadowRoot ?? field,
) {
  return Array.from(
    new Set(
      Array.from(
        scope.querySelectorAll<HTMLElement>(
          `.ML__placeholder, .${visualTexPlaceholderClass}`,
        ),
      ),
    ),
  ).filter((placeholder) => {
    const bounds = placeholder.getBoundingClientRect();
    const style = getComputedStyle(placeholder);
    return (
      bounds.width > 0 &&
      bounds.height > 0 &&
      style.display !== "none" &&
      style.visibility !== "hidden"
    );
  });
}

function closestPlaceholderNodeToMarker(
  field: MathfieldElement,
  marker: HTMLElement,
) {
  if (
    marker.classList.contains("ML__placeholder") ||
    marker.classList.contains(visualTexPlaceholderClass)
  ) {
    return marker;
  }
  const markerBounds = (marker.parentElement ?? marker).getBoundingClientRect();
  const markerX = markerBounds.left + markerBounds.width / 2;
  const markerY = markerBounds.top + markerBounds.height / 2;
  return visiblePlaceholderNodes(field)
    .map((placeholder) => {
      const bounds = placeholder.getBoundingClientRect();
      const centerX = bounds.left + bounds.width / 2;
      const centerY = bounds.top + bounds.height / 2;
      return {
        placeholder,
        distance: Math.hypot(centerX - markerX, centerY - markerY),
      };
    })
    .sort((first, second) => first.distance - second.distance)[0]
    ?.placeholder ?? null;
}

function describeVerticalPlaceholderNode(
  field: MathfieldElement,
  placeholder: HTMLElement,
): VerticalPlaceholderAnchor | null {
  if (placeholder.closest(".ML__sqrt, .ML__accent-body")) return null;

  const placeholderCount = (container: HTMLElement) =>
    visiblePlaceholderNodes(field, container).length;
  const fraction = placeholder.closest<HTMLElement>(".ML__mfrac");
  const operator = placeholder.closest<HTMLElement>(".ML__op-group");
  const script = placeholder.closest<HTMLElement>(".ML__msubsup");
  let stack: HTMLElement | null = null;
  for (
    let current = placeholder.parentElement;
    current;
    current = current.parentElement
  ) {
    if (
      (current.classList.contains("ML__vlist") ||
        current.classList.contains("ML__vlist-r") ||
        current.classList.contains("ML__vlist-t")) &&
      placeholderCount(current) >= 2
    ) {
      stack = current;
      break;
    }
  }

  const directContainer = fraction ?? operator ?? script;
  let container =
    directContainer && placeholderCount(directContainer) >= 2
      ? directContainer
      : stack;
  if (!container) {
    const base = placeholder.closest<HTMLElement>(".ML__base");
    const centers = base
      ? visiblePlaceholderNodes(field, base).map((node) => {
          const bounds = node.getBoundingClientRect();
          return bounds.top + bounds.height / 2;
        })
      : [];
    const verticalSpread = centers.length
      ? Math.max(...centers) - Math.min(...centers)
      : 0;
    if (base && centers.length >= 2 && verticalSpread >= 6) {
      container = base;
    }
  }
  if (!container) return null;

  const placeholderBounds = placeholder.getBoundingClientRect();
  const containerBounds = container.getBoundingClientRect();
  if (
    placeholderBounds.height <= 0 ||
    containerBounds.width <= 0 ||
    containerBounds.height <= 0
  ) {
    return null;
  }

  const centerX = placeholderBounds.left + placeholderBounds.width / 2;
  const centerY = placeholderBounds.top + placeholderBounds.height / 2;
  const relativeX = (centerX - containerBounds.left) / containerBounds.width;
  const relativeY = (centerY - containerBounds.top) / containerBounds.height;
  const kind: VerticalPlaceholderKind = fraction
    ? "fraction"
    : operator
      ? "operator-limit"
      : script
        ? "script"
        : "stack";
  return {
    kind,
    region: relativeY < 0.5 ? "upper" : "lower",
    centerX,
    centerY,
    relativeX,
    relativeY,
  };
}

function describeSelectedVerticalPlaceholder(
  field: MathfieldElement,
): VerticalPlaceholderAnchor | null {
  const marker = activeMathCaretMarker(field);
  if (!marker) return null;
  const placeholder = closestPlaceholderNodeToMarker(field, marker);
  return placeholder
    ? describeVerticalPlaceholderNode(field, placeholder)
    : null;
}

function placeholderRangeForVisualNode(
  field: MathfieldElement,
  placeholder: HTMLElement,
  ranges: Array<[number, number]>,
) {
  const bounds = placeholder.getBoundingClientRect();
  const pointOffset = field.getOffsetFromPoint(
    bounds.left + bounds.width / 2,
    bounds.top + bounds.height / 2,
    { bias: 0 },
  );
  return ranges
    .map((range) => ({
      range,
      distance: Math.min(
        Math.abs(range[0] - pointOffset),
        Math.abs(range[1] - pointOffset),
      ),
    }))
    .sort((first, second) => first.distance - second.distance)[0]
    ?.range ?? null;
}

function findMatchingVerticalPlaceholderRange(
  field: MathfieldElement,
  target: VerticalPlaceholderAnchor,
) {
  const ranges = latexPlaceholderRanges(field);
  const candidates = visiblePlaceholderNodes(field).flatMap((placeholder) => {
    const description = describeVerticalPlaceholderNode(field, placeholder);
    if (
      !description ||
      description.kind !== target.kind ||
      description.region !== target.region
    ) {
      return [];
    }
    const range = placeholderRangeForVisualNode(field, placeholder, ranges);
    if (!range) return [];
    return [{
      range,
      score:
        Math.abs(description.relativeY - target.relativeY) * 1200 +
        Math.abs(description.relativeX - target.relativeX) * 300 +
        Math.abs(description.centerY - target.centerY) * 4 +
        Math.abs(description.centerX - target.centerX),
    }];
  });
  return candidates.sort((first, second) => first.score - second.score)[0]
    ?.range ?? null;
}

function canonicalRawCommandAnchorSkeleton(latex: string) {
  const verifier = new MathfieldElement();
  verifier.setValue(latex.replace(/\\placeholder\{\}/g, ""), {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "after",
    silenceNotifications: true,
  });
  return normalizeChineseLatex(verifier.value)
    .replace(/\s+/g, "")
    .replace(/\{([A-Za-z0-9])\}/g, "$1");
}

function restoreCancelledRawCommandAnchor(
  field: MathfieldElement,
  anchor: RawCommandAnchor,
) {
  field.executeCommand(["complete", "reject"]);
  field.mode = "math";
  field.setValue(anchor.latex, {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "after",
    silenceNotifications: true,
  });
  const selection = clampSelection(anchor.selection, field.lastOffset);
  field.selection = selection;
  if (selection.ranges.every(([start, end]) => start === end)) {
    field.position = Math.max(0, Math.min(field.lastOffset, anchor.position));
  }
  return selection;
}

function restoreSelectedPlaceholderAnchor(
  field: MathfieldElement,
  anchor: RawCommandAnchor,
) {
  const placeholderIndex = anchor.selectedPlaceholderIndex;
  if (placeholderIndex === null) return null;

  field.executeCommand(["complete", "reject"]);
  field.mode = "math";
  // `replaceAll` and an ordinary full-range selection can leave the empty
  // container that held the rejected raw group in MathLive's atom tree. Exit
  // every nested parent first, then remove every atom from the root model.
  for (let depth = 0; depth < 12; depth += 1) {
    if (!field.executeCommand("moveAfterParent")) break;
  }
  field.executeCommand("deleteAll");
  field.setValue(anchor.latex, {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "after",
    silenceNotifications: true,
  });
  const targetRange = latexPlaceholderRanges(field)[placeholderIndex] ?? null;
  if (!targetRange) return null;

  const selection: MathSelectionSnapshot = {
    ranges: [targetRange],
    direction: "none",
  };
  field.selection = selection;
  return selection;
}

function restoreRawCommandInsertionAnchor(
  field: MathfieldElement,
  anchor: RawCommandAnchor,
) {
  if (anchor.selectedPlaceholder) {
    const restored = restoreSelectedPlaceholderAnchor(field, anchor);
    if (restored) return restored;
    const selection = clampSelection(anchor.selection, field.lastOffset);
    field.selection = selection;
    return selection;
  }
  return restoreRawCommandAnchor(field, anchor);
}

function restoreRawCommandAnchor(
  field: MathfieldElement,
  anchor: RawCommandAnchor,
) {
  const rejectedRawGroup = field.executeCommand(["complete", "reject"]);
  if (!rejectedRawGroup) field.mode = "math";

  const currentLatex = normalizeChineseLatex(field.value);
  const anchorContainsPlaceholder = anchor.latex.includes("\\placeholder{}");
  const retainedPlaceholderContainer =
    anchorContainsPlaceholder &&
    canonicalRawCommandAnchorSkeleton(currentLatex) ===
      canonicalRawCommandAnchorSkeleton(anchor.latex);

  // MathLive keeps the original empty container when it rejects nested raw
  // input. Keep that model, but explicitly restore the original caret: the
  // selection produced by reject can drift outside an accent/script/fraction
  // and makes the accepted command appear in a surprising location.
  if (retainedPlaceholderContainer) {
    field.mode = "math";
    const preferredRange = anchor.selection.ranges.at(-1);
    const preferredStart = preferredRange
      ? Math.min(preferredRange[0], preferredRange[1])
      : anchor.position;

    // Preserve the original, proven single-slot behaviour for accents, roots
    // and other ordinary placeholders. Multi-branch vertical structures are
    // intercepted separately before this function is called.
    if (anchor.selectedPlaceholder) {
      const closestPlaceholder = latexPlaceholderRanges(field).sort(
        (first, second) =>
          Math.abs(first[0] - preferredStart) -
          Math.abs(second[0] - preferredStart),
      )[0];
      if (closestPlaceholder) {
        const selection: MathSelectionSnapshot = {
          ranges: [closestPlaceholder],
          direction: "none",
        };
        field.selection = selection;
        return selection;
      }

      const position = Math.max(
        0,
        Math.min(field.lastOffset, preferredStart),
      );
      const selection: MathSelectionSnapshot = {
        ranges: [[position, position]],
        direction: "none",
      };
      field.selection = selection;
      field.position = position;
      return selection;
    }

    const selection = clampSelection(anchor.selection, field.lastOffset);
    field.selection = selection;
    if (selection.ranges.every(([start, end]) => start === end)) {
      field.position = Math.max(
        0,
        Math.min(field.lastOffset, anchor.position),
      );
    }
    return selection;
  }

  if (currentLatex !== anchor.latex) {
    field.setValue(anchor.latex, {
      mode: "math",
      format: "latex",
      insertionMode: "replaceAll",
      selectionMode: "after",
      silenceNotifications: true,
    });
  }
  const selection = clampSelection(anchor.selection, field.lastOffset);
  field.selection = selection;
  const selectionIsCollapsed = selection.ranges.every(
    ([start, end]) => start === end,
  );
  if (selectionIsCollapsed) {
    field.position = Math.max(0, Math.min(field.lastOffset, anchor.position));
  }
  return selection;
}

const structuredSuggestionCommands = new Set([
  ...compatibilityWrapperPreviews.keys(),
  ...compatibilityRawPlaceholderTemplates.keys(),
  "\\sum",
  "\\prod",
  "\\coprod",
  "\\int",
  "\\iint",
  "\\iiint",
  "\\oint",
  "\\oiint",
  "\\oiiint",
  "\\lim",
  "\\bigcup",
  "\\bigcap",
]);
const nativePlaceholderSelectionCommands = new Set([
  "\\frac",
  "\\dfrac",
  "\\tfrac",
  "\\sqrt",
]);
const accentCommandTemplates = new Map<string, string>([
  ["\\acute", "\\acute{\\placeholder{}}"],
  ["\\grave", "\\grave{\\placeholder{}}"],
  ["\\dot", "\\dot{\\placeholder{}}"],
  ["\\ddot", "\\ddot{\\placeholder{}}"],
  ["\\dddot", "\\dddot{\\placeholder{}}"],
  ["\\ddddot", "\\ddddot{\\placeholder{}}"],
  ["\\tilde", "\\tilde{\\placeholder{}}"],
  ["\\bar", "\\bar{\\placeholder{}}"],
  ["\\breve", "\\breve{\\placeholder{}}"],
  ["\\check", "\\check{\\placeholder{}}"],
  ["\\hat", "\\hat{\\placeholder{}}"],
  ["\\vec", "\\vec{\\placeholder{}}"],
  ["\\widehat", "\\widehat{\\placeholder{}}"],
  ["\\widetilde", "\\widetilde{\\placeholder{}}"],
  ["\\overline", "\\overline{\\placeholder{}}"],
  ["\\overrightarrow", "\\overrightarrow{\\placeholder{}}"],
  ["\\overleftarrow", "\\overleftarrow{\\placeholder{}}"],
  ["\\overleftrightarrow", "\\overleftrightarrow{\\placeholder{}}"],
  ["\\mathring", "\\mathring{\\placeholder{}}"],
]);
const CASES_ENVIRONMENT_COMMAND = "\\begin{cases}";
const CASES_ENVIRONMENT_TEMPLATE =
  "\\begin{cases}\\placeholder{} & \\placeholder{}\\end{cases}";

const rawPlaceholderCommandTemplates = new Map<string, string>([
  ...accentCommandTemplates,
  ...compatibilityRawPlaceholderTemplates,
  [CASES_ENVIRONMENT_COMMAND, CASES_ENVIRONMENT_TEMPLATE],
  ["\\sqrt", "\\sqrt{\\placeholder{}}"],
  ["\\frac", "\\frac{\\placeholder{}}{\\placeholder{}}"],
  ["\\dfrac", "\\dfrac{\\placeholder{}}{\\placeholder{}}"],
  ["\\tfrac", "\\tfrac{\\placeholder{}}{\\placeholder{}}"],
  ["\\binom", "\\binom{\\placeholder{}}{\\placeholder{}}"],
  ["\\overset", "\\overset{\\placeholder{}}{\\placeholder{}}"],
  ["\\underset", "\\underset{\\placeholder{}}{\\placeholder{}}"],
  [
    "\\overunderset",
    "\\overset{\\placeholder{}}{\\underset{\\placeholder{}}{\\placeholder{}}}",
  ],
  ["\\stackrel", "\\stackrel{\\placeholder{}}{\\placeholder{}}"],
  ["\\stackbin", "\\stackbin{\\placeholder{}}{\\placeholder{}}"],
  ["\\overarc", "\\overarc{\\placeholder{}}"],
  ["\\overbrace", "\\overbrace{\\placeholder{}}"],
  ["\\overgroup", "\\overgroup{\\placeholder{}}"],
  ["\\overparen", "\\overparen{\\placeholder{}}"],
  ["\\overleftharpoon", "\\overleftharpoon{\\placeholder{}}"],
  ["\\overlinesegment", "\\overlinesegment{\\placeholder{}}"],
  ["\\overrightharpoon", "\\overrightharpoon{\\placeholder{}}"],
  ["\\underarc", "\\underarc{\\placeholder{}}"],
  ["\\underline", "\\underline{\\placeholder{}}"],
  ["\\underbrace", "\\underbrace{\\placeholder{}}"],
  ["\\undergroup", "\\undergroup{\\placeholder{}}"],
  ["\\underparen", "\\underparen{\\placeholder{}}"],
  ["\\underleftarrow", "\\underleftarrow{\\placeholder{}}"],
  ["\\underrightarrow", "\\underrightarrow{\\placeholder{}}"],
  ["\\underlinesegment", "\\underlinesegment{\\placeholder{}}"],
  ["\\underleftrightarrow", "\\underleftrightarrow{\\placeholder{}}"],
]);
const reverseModelPlaceholderOrderCommands = new Set([
  "\\overset",
  "\\underset",
  "\\overunderset",
  "\\stackrel",
  "\\stackbin",
]);
function selectFirstLatexPlaceholder(
  field: MathfieldElement,
  command: string,
  insertionTemplate = command,
) {
  if (hasBoundedOperatorPlaceholderOrder(insertionTemplate)) {
    if (
      field.selectionIsCollapsed &&
      field.position > 0 &&
      field.getValue(field.position - 1, field.position, "latex").trim() ===
        "\\placeholder{}"
    ) {
      field.selection = {
        ranges: [[field.position - 1, field.position]],
        direction: "none",
      };
    }
    return;
  }
  if (!reverseModelPlaceholderOrderCommands.has(command)) return;
  for (let offset = field.lastOffset; offset > 0; offset -= 1) {
    if (field.getElementInfo(offset)?.latex?.trim() !== "\\placeholder{}") {
      continue;
    }
    field.selection = {
      ranges: [[offset - 1, offset]],
      direction: "none",
    };
    return;
  }
}

function isAccentContainerLatex(latex: string) {
  const normalized = latex.trim();
  return Array.from(accentCommandTemplates.keys()).some((command) =>
    normalized.startsWith(command + "{"),
  );
}

function selectAdjacentAccentPlaceholder(
  field: MathfieldElement,
  direction: "left" | "right",
) {
  if (!field.selectionIsCollapsed || field.mode === "latex") return false;

  const position = field.position;
  for (let end = 1; end <= field.lastOffset; end += 1) {
    if (
      field.getValue(end - 1, end, "latex").trim() !== "\\placeholder{}"
    ) {
      continue;
    }
    const containerLatex =
      field.getElementInfo(Math.min(field.lastOffset, end + 1))?.latex ?? "";
    if (!isAccentContainerLatex(containerLatex)) continue;

    const start = end - 1;
    const adjacentBoundary =
      direction === "left" ? end + 1 : start - 1;
    if (position !== adjacentBoundary) continue;

    field.selection = {
      ranges: [[start, end]],
      direction: "none",
    };
    return true;
  }
  return false;
}

const BASE_FORMULA_FONT_SIZE = 54;
const MIN_FORMULA_FONT_SIZE = BASE_FORMULA_FONT_SIZE * 0.2;

const formulaFontSize = (zoom: number) =>
  Math.max(MIN_FORMULA_FONT_SIZE, BASE_FORMULA_FONT_SIZE * zoom);

function formulaRowHeightMetrics(
  latex: string,
  zoom: number,
  formulaRowVerticalInset: number,
) {
  const fontSize = formulaFontSize(zoom);
  const hasTallStructure = tallFormulaPattern.test(latex);
  const safeVerticalInset = Math.max(
    0,
    Math.min(24, Math.round(formulaRowVerticalInset)),
  );
  const verticalPadding = safeVerticalInset * 2;
  const baseContentHeight = hasTallStructure
    ? Math.max(20, fontSize * 1.34)
    : Math.max(22, fontSize * 1.12);
  const baseHeight = Math.max(
    hasTallStructure ? 36 : 30,
    baseContentHeight + verticalPadding,
  );
  return {
    fontSize,
    hasTallStructure,
    baseHeight,
    verticalPadding,
  };
}

function predictedFormulaRowHeight(
  latex: string,
  zoom: number,
  formulaRowVerticalInset: number,
  measurementHost?: HTMLElement | null,
) {
  const metrics = formulaRowHeightMetrics(
    latex,
    zoom,
    formulaRowVerticalInset,
  );
  let renderedHeight = latex.trim() ? 0 : metrics.fontSize * 1.2;

  // Keep the accurate pre-mount measurement for ordinary formulas because
  // early pointer hit testing depends on the correct row geometry. The render
  // itself is now guarded: structurally dangerous LaTeX is rejected before it
  // reaches MathLive, and any remaining renderer exception falls back to the
  // heuristic height instead of taking down the React root.
  if (measurementHost && latex.trim()) {
    const markup = safeConvertVisualTexLatexToMarkup(latex, {
      defaultMode: "math",
    });
    if (markup) {
      const probe = document.createElement("span");
      probe.setAttribute("aria-hidden", "true");
      probe.style.position = "absolute";
      probe.style.left = "-100000px";
      probe.style.top = "0";
      probe.style.display = "inline-block";
      probe.style.width = "max-content";
      probe.style.maxWidth = "none";
      probe.style.whiteSpace = "nowrap";
      probe.style.visibility = "hidden";
      probe.style.pointerEvents = "none";
      probe.style.fontSize = metrics.fontSize + "px";
      probe.style.lineHeight = "normal";
      probe.innerHTML = markup;
      measurementHost.append(probe);
      renderedHeight = probe.getBoundingClientRect().height;
      probe.remove();
    }
  }

  return Math.ceil(
    Math.max(
      metrics.baseHeight,
      renderedHeight + metrics.verticalPadding,
    ),
  );
}

function captureSelection(field: MathfieldElement): MathSelectionSnapshot {
  return {
    ranges: field.selection.ranges.map(
      ([start, end]) => [start, end] as [number, number],
    ),
    direction: field.selection.direction ?? "none",
  };
}

function selectionHasContent(selection: MathSelectionSnapshot) {
  return selection.ranges.some(([start, end]) => start !== end);
}

function captureFieldSnapshot(field: MathfieldElement) {
  return {
    latex: normalizeChineseLatex(field.value),
    selection: captureSelection(field),
  };
}

const visualTexPlaceholderStyleId = "visualtex-structural-placeholder-style";
const visualTexFormulaFontStyleId = "visualtex-formula-font-style";
const visualTexAlignmentMarkerStyleId = "visualtex-alignment-marker-style";
const visualTexFormulaUprightFontProperty =
  "--visualtex-formula-upright-font-family";
const visualTexFormulaItalicFontProperty =
  "--visualtex-formula-italic-font-family";
const visualTexFormulaChineseFontProperty =
  "--visualtex-formula-chinese-font-family";
const visualTexPlaceholderClass = "visualtex-structural-placeholder";
const visualTexAccentPlaceholderClass =
  "visualtex-accent-structural-placeholder";
const visualTexPlaceholderCaretClass =
  "visualtex-structural-placeholder-caret";
const visualTexPlaceholderSelectionClass =
  "has-visualtex-structural-placeholder-selection";
const visualTexRawLatexClass = "has-visualtex-raw-latex-command";
const visualTexPointerSelectingClass = "visualtex-pointer-selecting";
const visualTexSourcePreviewClass = "visualtex-source-preview-only";
const visualTexCaretRepaintClass = "visualtex-caret-repaint";
const visualTexPostOperatorCaretClass = "visualtex-post-operator-caret";
const visualTexPostOperatorCaretShiftProperty =
  "--visualtex-post-operator-caret-shift";
const visualTexCaretRepaintFrames = new WeakMap<MathfieldElement, number>();

function isMathLiveInterAtomGlue(node: Element | null): node is HTMLElement {
  if (!(node instanceof HTMLElement)) return false;
  if (node.dataset.atomId || node.className || node.textContent) return false;
  return (
    node.style.display === "inline-block" &&
    /^\d*\.?\d+(?:em|ex|px)$/.test(node.style.width)
  );
}

function syncPostOperatorCaretSpacing(field: MathfieldElement) {
  const caret = field.shadowRoot?.querySelector<HTMLElement>(".ML__caret");
  if (!caret) return;

  const previousAtom = caret.previousElementSibling;
  const leadingGlue = previousAtom?.previousElementSibling ?? null;
  const trailingGlue = caret.nextElementSibling;
  const shouldShift =
    previousAtom instanceof HTMLElement &&
    Boolean(previousAtom.dataset.atomId) &&
    isMathLiveInterAtomGlue(leadingGlue) &&
    isMathLiveInterAtomGlue(trailingGlue);

  if (!shouldShift) {
    if (caret.classList.contains(visualTexPostOperatorCaretClass)) {
      caret.classList.remove(visualTexPostOperatorCaretClass);
    }
    caret.style.removeProperty(visualTexPostOperatorCaretShiftProperty);
    return;
  }

  const shift = trailingGlue.getBoundingClientRect().width;
  if (!(shift > 0)) return;
  const nextShift = `${shift}px`;
  if (
    caret.style.getPropertyValue(visualTexPostOperatorCaretShiftProperty) !==
    nextShift
  ) {
    caret.style.setProperty(
      visualTexPostOperatorCaretShiftProperty,
      nextShift,
    );
  }
  if (!caret.classList.contains(visualTexPostOperatorCaretClass)) {
    caret.classList.add(visualTexPostOperatorCaretClass);
  }
}

function repaintMathLiveCaret(field: MathfieldElement) {
  const previousFrame = visualTexCaretRepaintFrames.get(field);
  if (previousFrame) window.cancelAnimationFrame(previousFrame);

  field.classList.remove(visualTexCaretRepaintClass);
  // Separate the previous WebKit compositing state from the class mutation.
  // This matters when the focused keyboard sink that dispatched Backspace has
  // just been removed from the document.
  field.shadowRoot
    ?.querySelector<HTMLElement>(
      ".ML__caret, .ML__text-caret, .ML__latex-caret",
    )
    ?.getBoundingClientRect();
  field.classList.add(visualTexCaretRepaintClass);

  const frame = window.requestAnimationFrame(() => {
    visualTexCaretRepaintFrames.delete(field);
    if (!field.isConnected) return;
    field.classList.remove(visualTexCaretRepaintClass);
  });
  visualTexCaretRepaintFrames.set(field, frame);
}
const visualTexPointerSelectingFields = new WeakSet<MathfieldElement>();

function markVisualTexFormulaGlyphFonts(field: MathfieldElement) {
  if (!field.shadowRoot) return;
  markVisualTexFormulaFontGlyphs(field.shadowRoot);
}

function syncStructuralPlaceholderSelection(field: MathfieldElement) {
  const root = field.shadowRoot;
  const container = root?.querySelector<HTMLElement>(".ML__container");
  if (!container) return;
  const [start, end] = field.selection.ranges[0] ?? [-1, -1];
  const selected =
    field.selection.ranges.length === 1 &&
    Math.abs(end - start) === 1 &&
    field.getValue(Math.min(start, end), Math.max(start, end), "latex").trim() ===
      "\\placeholder{}" &&
    Boolean(
      root?.querySelector(
        `.${visualTexPlaceholderClass}.ML__selected, .ML__placeholder.ML__selected`,
      ),
    );
  container.classList.toggle(visualTexPlaceholderSelectionClass, selected);
}

function installVisualTexFormulaFontStyle(field: MathfieldElement) {
  const shadowRoot = field.shadowRoot;
  if (!shadowRoot) return;
  if (shadowRoot.getElementById(visualTexFormulaFontStyleId)) return;

  const style = document.createElement("style");
  style.id = visualTexFormulaFontStyleId;
  style.textContent = `
    .ML__cmr.${VISUALTEX_FORMULA_LETTER_GLYPH_CLASS}:not(.ML__it) {
      font-family: var(${visualTexFormulaUprightFontProperty}, KaTeX_Main, serif) !important;
      font-style: normal !important;
    }

    .ML__cmr.${VISUALTEX_FORMULA_LETTER_GLYPH_CLASS}:not(.ML__bold):not(.ML__it) {
      font-weight: 400 !important;
    }

    .ML__cmr.${VISUALTEX_FORMULA_LETTER_GLYPH_CLASS}.ML__bold:not(.ML__it),
    .ML__mathbf.${VISUALTEX_FORMULA_LETTER_GLYPH_CLASS}:not(.lcGreek) {
      font-family: var(${visualTexFormulaUprightFontProperty}, KaTeX_Main, serif) !important;
      font-style: normal !important;
      font-weight: 700 !important;
    }

    .ML__cmr.${VISUALTEX_FORMULA_LETTER_GLYPH_CLASS}.ML__it,
    .ML__mathit.${VISUALTEX_FORMULA_LETTER_GLYPH_CLASS} {
      font-family: var(${visualTexFormulaItalicFontProperty}, KaTeX_Math, KaTeX_Main, serif) !important;
      font-style: italic !important;
    }

    .ML__cmr.${VISUALTEX_FORMULA_LETTER_GLYPH_CLASS}.ML__it:not(.ML__bold),
    .ML__mathit.${VISUALTEX_FORMULA_LETTER_GLYPH_CLASS} {
      font-weight: 400 !important;
    }

    .ML__cmr.${VISUALTEX_FORMULA_LETTER_GLYPH_CLASS}.ML__bold.ML__it,
    .lcGreek.ML__mathbf.${VISUALTEX_FORMULA_LETTER_GLYPH_CLASS},
    .ML__mathbfit.${VISUALTEX_FORMULA_LETTER_GLYPH_CLASS} {
      font-family: var(${visualTexFormulaItalicFontProperty}, KaTeX_Math, KaTeX_Main, serif) !important;
      font-style: italic !important;
      font-weight: 700 !important;
    }

    .ML__text,
    .${VISUALTEX_FORMULA_CHINESE_GLYPH_CLASS} {
      font-family: var(${visualTexFormulaChineseFontProperty}, var(--_text-font-family)) !important;
    }
  `;
  shadowRoot.append(style);
}

function applyVisualTexFormulaFonts(
  field: MathfieldElement,
  letterFont: FormulaLetterFont,
  chineseFont: FormulaChineseFont,
) {
  installVisualTexFormulaFontStyle(field);
  markVisualTexFormulaGlyphFonts(field);
  const letterFamilies = formulaLetterFontFamilies(letterFont);
  field.style.setProperty(
    visualTexFormulaUprightFontProperty,
    letterFamilies.upright,
  );
  field.style.setProperty(
    visualTexFormulaItalicFontProperty,
    letterFamilies.italic,
  );
  field.style.setProperty(
    visualTexFormulaChineseFontProperty,
    formulaChineseFontFamily(chineseFont),
  );
}

function installVisualTexStructuralPlaceholderStyle(field: MathfieldElement) {
  const shadowRoot = field.shadowRoot;
  if (!shadowRoot) return;
  if (!shadowRoot.getElementById(visualTexPlaceholderStyleId)) {
    const style = document.createElement("style");
    style.id = visualTexPlaceholderStyleId;
    style.textContent = `
      /*
       * Structural placeholders must never participate in formula geometry.
       * Keep MathLive's native placeholder metrics/baseline and paint the blue
       * marker as an absolutely-positioned overlay. This makes the same marker
       * work in fractions, scripts, radicals, accents, matrices and nested
       * structures without per-structure vertical offsets.
       */
      .ML__placeholder,
      .${visualTexPlaceholderClass} {
        position: relative !important;
        overflow: visible !important;
        border: 0 !important;
        background: transparent !important;
        color: transparent !important;
        opacity: 1 !important;
        box-shadow: none !important;
      }

      .ML__placeholder::before,
      .${visualTexPlaceholderClass}::before {
        content: "" !important;
        position: absolute !important;
        z-index: 0 !important;
        left: 50% !important;
        top: 50% !important;
        display: block !important;
        width: 0.40em !important;
        height: 0.74em !important;
        border-radius: 0.075em !important;
        background: var(--formula-placeholder) !important;
        transform: translate(-50%, -50%) !important;
        pointer-events: none !important;
      }

      .ML__contains-highlight,
      .ML__highlight {
        border-color: transparent !important;
        outline: 0 !important;
        background: transparent !important;
        background-color: transparent !important;
        box-shadow: none !important;
      }

      .ML__container.${visualTexPlaceholderSelectionClass} .ML__selected {
        border-color: transparent !important;
        outline: 0 !important;
        background: transparent !important;
        background-color: transparent !important;
        box-shadow: none !important;
      }

      .ML__container.${visualTexPlaceholderSelectionClass} .ML__placeholder-selected,
      .ML__container.${visualTexPlaceholderSelectionClass} .ML__selected .ML__placeholder,
      .ML__container.${visualTexPlaceholderSelectionClass} .${visualTexPlaceholderClass}.ML__selected,
      .ML__container.${visualTexPlaceholderSelectionClass} .ML__selected .${visualTexPlaceholderClass} {
        overflow: visible !important;
        border: 0 !important;
        background: transparent !important;
        color: transparent !important;
        opacity: 1 !important;
        box-shadow: none !important;
      }

      .ML__container.${visualTexPlaceholderSelectionClass}
        .ML__placeholder-selected::before,
      .ML__container.${visualTexPlaceholderSelectionClass}
        .ML__selected .ML__placeholder::before,
      .ML__container.${visualTexPlaceholderSelectionClass}
        .${visualTexPlaceholderClass}.ML__selected::before,
      .ML__container.${visualTexPlaceholderSelectionClass}
        .ML__selected .${visualTexPlaceholderClass}::before {
        background: var(--formula-placeholder-selected) !important;
      }

      :host(.${visualTexCaretRepaintClass})
        .ML__focused .ML__caret::after,
      :host(.${visualTexCaretRepaintClass})
        .ML__focused .ML__text-caret::after,
      :host(.${visualTexCaretRepaintClass})
        .ML__focused .ML__latex-caret::after {
        visibility: visible !important;
        opacity: 1 !important;
        animation: none !important;
        transform: translateZ(0);
      }

      .ML__caret.${visualTexPostOperatorCaretClass} {
        transform: translateX(
          var(${visualTexPostOperatorCaretShiftProperty}, 0px)
        ) !important;
      }

      .${visualTexPlaceholderCaretClass} {
        pointer-events: none !important;
      }

      .ML__container.${visualTexPlaceholderSelectionClass} .ML__selection {
        display: none !important;
      }

      :host(.${visualTexPointerSelectingClass})
        .ML__container.${visualTexPlaceholderSelectionClass}
        .ML__selection {
        display: block !important;
        background: var(--_selection-background-color) !important;
      }

      :host(.${visualTexPointerSelectingClass})
        .${visualTexPlaceholderCaretClass} {
        opacity: 0 !important;
        animation: none !important;
      }

      .ML__container.${visualTexRawLatexClass} .ML__latex,
      .ML__container.${visualTexRawLatexClass} .ML__raw-latex,
      .ML__container.${visualTexRawLatexClass} .ML__selected,
      .ML__container.${visualTexRawLatexClass} .ML__contains-highlight,
      .ML__container.${visualTexRawLatexClass} .ML__selection {
        background: transparent !important;
        background-color: transparent !important;
        box-shadow: none !important;
        outline: 0 !important;
      }

      .ML__container.${visualTexRawLatexClass} .ML__selection {
        display: none !important;
      }

      :host(.has-visualtex-multi-line-selection)
        .ML__container .ML__selection {
        display: block !important;
        background: var(--_selection-background-color) !important;
      }

      :host(.${visualTexSourcePreviewClass}) .ML__selection {
        display: none !important;
      }

      :host(.${visualTexSourcePreviewClass}) .ML__selected,
      :host(.${visualTexSourcePreviewClass}) .ML__contains-highlight,
      :host(.${visualTexSourcePreviewClass}) .ML__highlight,
      :host(.${visualTexSourcePreviewClass}) .ML__placeholder-selected,
      :host(.${visualTexSourcePreviewClass}) .ML__selected .ML__placeholder,
      :host(.${visualTexSourcePreviewClass})
        .ML__selected .${visualTexPlaceholderClass} {
        border-color: transparent !important;
        outline: 0 !important;
        background: transparent !important;
        background-color: transparent !important;
        box-shadow: none !important;
      }

      :host(.${visualTexSourcePreviewClass}) .ML__caret,
      :host(.${visualTexSourcePreviewClass}) .ML__text-caret,
      :host(.${visualTexSourcePreviewClass}) .ML__latex-caret,
      :host(.${visualTexSourcePreviewClass})
        .${visualTexPlaceholderCaretClass} {
        display: none !important;
        opacity: 0 !important;
        animation: none !important;
      }
    `;
    shadowRoot.append(style);
  }
}

function normalizeCompletedDifferentialDisplay(field: MathfieldElement) {
  if (field.mode === "latex" || !field.selectionIsCollapsed) return false;

  const originalValue = field.value;
  const portableValue = normalizeMathLiveCanonicalUprightCommands(originalValue);
  // Contextual upright typography is a semantic normalization pass, not an
  // inline-shortcut preference. It must remain active even when the user turns
  // off "common math input auto-convert".
  const contextualValue = normalizeContextualUprightSymbols(portableValue);
  if (contextualValue === originalValue) return false;

  const distanceFromEnd = Math.max(0, field.lastOffset - field.position);
  field.setValue(contextualValue, {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "after",
    silenceNotifications: true,
  });
  const nextPosition = Math.max(0, field.lastOffset - distanceFromEnd);
  field.selection = {
    ranges: [[nextPosition, nextPosition]],
    direction: "none",
  };
  field.position = nextPosition;
  return true;
}

function selectionIsCollapsed(selection: MathSelectionSnapshot) {
  return selection.ranges.every(([start, end]) => start === end);
}

function inferEditKind(
  inputType: string,
  beforeSelection: MathSelectionSnapshot,
): EditKind {
  if (inputType.includes("deleteContentBackward")) return "delete-backward";
  if (inputType.includes("deleteContentForward")) return "delete-forward";
  if (inputType.includes("Composition")) return "composition";
  if (inputType.includes("insert") && selectionIsCollapsed(beforeSelection)) {
    return "insert";
  }
  return "replace";
}

function inferEditSource(inputType: string): FormulaEditSource {
  return inputType.toLocaleLowerCase().includes("paste") ? "paste" : "keyboard";
}

const tallFormulaPattern =
  /\\(?:d?frac|tfrac|sqrt|sum|prod|int|iint|iiint|oint|oiint|oiiint|lim|begin|overset|underset|overline|underline)\b|[_^]/;
const bareStructuredOperatorPattern =
  /^\\(?:int|iint|iiint|oint|oiint|oiiint|sum|prod|lim|bigcup|bigcap)\s*$/;
const scriptContainerPattern = /^[_^]\{[\s\S]*\}$/;

function findTrailingCommandRange(
  field: MathfieldElement,
  activeQuery: string,
): [number, number] | null {
  const normalizedQuery = activeQuery.trim();
  if (!normalizedQuery) return null;

  const candidateEnds = Array.from(
    new Set([field.position, field.lastOffset].filter((offset) => offset >= 0)),
  );
  for (const end of candidateEnds) {
    for (let start = end; start >= 0; start -= 1) {
      const rangeLatex = field.getValue(start, end, "latex").trim();
      if (rangeLatex === normalizedQuery) return [start, end];
    }
  }

  return null;
}

function dismissNativeSuggestionPopover(field: MathfieldElement) {
  dismissMathLiveSuggestions(field);
}

function keepCaretAfterBareStructuredOperator(
  field: MathfieldElement,
  previousPosition: number,
) {
  if (field.position >= previousPosition || field.position >= field.lastOffset) {
    return;
  }

  const operatorOffset = field.position + 1;
  const operatorLatex =
    field.getElementInfo(operatorOffset)?.latex?.trim() ||
    field.getValue(field.position, operatorOffset, "latex").trim();
  if (!bareStructuredOperatorPattern.test(operatorLatex)) return;

  field.selection = {
    ranges: [[operatorOffset, operatorOffset]],
    direction: "none",
  };
  field.position = operatorOffset;
}

type ScriptCaretRegion = "upper" | "lower";

type ScriptCaretContext = {
  region: ScriptCaretRegion;
  kind: "script" | "operator-limit";
  caret: HTMLElement | null;
  container: HTMLElement | null;
  autoExitKey: string | null;
};

const consumedScriptAutoExitKeys = new WeakMap<
  MathfieldElement,
  Set<string>
>();

function activeMathCaretMarker(field: MathfieldElement) {
  return Array.from(
    field.shadowRoot?.querySelectorAll<HTMLElement>(
      ".ML__placeholder-selected, .ML__selected, .ML__caret",
    ) ?? [],
  ).find((marker) => marker.getBoundingClientRect().height > 0) ?? null;
}

function scriptBranchAutoExitKey(
  caret: HTMLElement | null,
  container: HTMLElement | null,
  region: ScriptCaretRegion,
) {
  if (!caret || !container || !container.classList.contains("ML__msubsup")) {
    return null;
  }

  // MathLive keeps a stable branch sentinel atom inside each super/subscript.
  // Its id survives caret movement and content edits, unlike model offsets,
  // which shift whenever another character is inserted.
  const branch = caret.parentElement ?? caret;
  const branchAtomId = branch
    .querySelector<HTMLElement>("[data-atom-id]")
    ?.getAttribute("data-atom-id");
  if (branchAtomId) return `${region}:branch:${branchAtomId}`;

  const base = container.previousElementSibling as HTMLElement | null;
  const baseAtomId =
    base?.getAttribute("data-atom-id") ??
    base
      ?.querySelector<HTMLElement>("[data-atom-id]")
      ?.getAttribute("data-atom-id");
  return baseAtomId ? `${region}:base:${baseAtomId}` : null;
}

function getScriptCaretContext(
  field: MathfieldElement,
): ScriptCaretContext | null {
  const caret = activeMathCaretMarker(field);
  const container = caret?.closest<HTMLElement>(
    ".ML__msubsup, .ML__op-group",
  ) ?? null;

  const currentOffset = Math.max(
    field.position,
    ...field.selection.ranges.flatMap(([start, end]) => [start, end]),
  );
  const currentDepth = field.getElementInfo(currentOffset)?.depth;
  let modelRegion: ScriptCaretRegion | null = null;
  if (typeof currentDepth === "number" && currentDepth > 0) {
    for (
      let offset = currentOffset + 1;
      offset <= Math.min(field.lastOffset, currentOffset + 3);
      offset += 1
    ) {
      const info = field.getElementInfo(offset);
      if (typeof info?.depth !== "number" || info.depth >= currentDepth) continue;
      const containerLatex = (info.latex ?? "").trim();
      if (/^\^\s*\{/.test(containerLatex)) modelRegion = "upper";
      else if (/^_\s*\{/.test(containerLatex)) modelRegion = "lower";
      break;
    }
  }

  let region = modelRegion;
  if (!region && caret && container) {
    const caretBounds = (caret.parentElement ?? caret).getBoundingClientRect();
    const containerBounds = container.getBoundingClientRect();
    if (caretBounds.height && containerBounds.height) {
      region =
        caretBounds.top + caretBounds.height / 2 <
        containerBounds.top + containerBounds.height / 2
          ? "upper"
          : "lower";
    }
  }
  if (!region) return null;

  const operatorGroup = container?.closest<HTMLElement>(".ML__op-group") ?? null;
  const outerScript =
    container?.parentElement?.closest<HTMLElement>(".ML__msubsup") ?? null;
  const kind =
    container?.classList.contains("ML__op-group") ||
    (operatorGroup && !outerScript)
      ? "operator-limit"
      : "script";
  return {
    region,
    kind,
    caret,
    container,
    autoExitKey:
      kind === "script"
        ? scriptBranchAutoExitKey(caret, container, region)
        : null,
  };
}

function scriptAutoExitWasConsumed(
  field: MathfieldElement,
  autoExitKey: string | null,
) {
  return Boolean(
    autoExitKey && consumedScriptAutoExitKeys.get(field)?.has(autoExitKey),
  );
}

function markScriptAutoExitConsumed(
  field: MathfieldElement,
  autoExitKey: string | null,
) {
  if (!autoExitKey) return;
  let keys = consumedScriptAutoExitKeys.get(field);
  if (!keys) {
    keys = new Set<string>();
    consumedScriptAutoExitKeys.set(field, keys);
  }
  keys.add(autoExitKey);
}

const accentContainerPattern =
  /^\\(?:acute|grave|dot|ddot|dddot|ddddot|tilde|bar|breve|check|hat|vec|widehat|widetilde|overline|overrightarrow|overleftarrow|overleftrightarrow)\s*\{/;

function caretIsInsideLatexContainer(
  field: MathfieldElement,
  containerPattern: RegExp,
) {
  const currentOffset = Math.max(
    field.position,
    ...field.selection.ranges.flatMap(([start, end]) => [start, end]),
  );
  const currentDepth = field.getElementInfo(currentOffset)?.depth;
  if (typeof currentDepth !== "number" || currentDepth <= 0) return false;

  for (
    let offset = currentOffset + 1;
    offset <= Math.min(field.lastOffset, currentOffset + 2);
    offset += 1
  ) {
    const info = field.getElementInfo(offset);
    if (
      typeof info?.depth === "number" &&
      info.depth < currentDepth &&
      containerPattern.test(info.latex ?? "")
    ) {
      return true;
    }
  }
  return false;
}

function caretIsInsideAccent(field: MathfieldElement) {
  return caretIsInsideLatexContainer(field, accentContainerPattern);
}

function getCaretAutoExitSetting(
  field: MathfieldElement,
): InputBehaviorSettingKey | null {
  if (caretIsInsideAccent(field)) return "autoExitAccent";

  const scriptContext = getScriptCaretContext(field);
  // Only the directly edited ordinary script participates in one-character
  // auto-exit. An outer radical, fraction, integral or other structure must not
  // suppress a nested x^a/x_a script; large-operator limits themselves remain
  // excluded.
  if (!scriptContext || scriptContext.kind !== "script") return null;
  if (scriptContext.region === "upper") return "autoExitSuperscript";
  if (scriptContext.region === "lower") return "autoExitSubscript";
  return null;
}

function isSingleDirectInput(event: InputEvent, field: MathfieldElement) {
  if (event.isComposing || field.mode === "latex") return false;
  if (event.inputType !== "insertText") return false;
  const data = event.data ?? "";
  if (data === "\\" || Array.from(data).length !== 1) return false;
  return true;
}

function moveCaretThroughEnabledAutoExitContainers(
  field: MathfieldElement,
  settings: InputBehaviorSettings,
  capturedSetting?: InputBehaviorSettingKey | null,
  capturedScriptKey?: string | null,
) {
  const moveOne = (
    setting: InputBehaviorSettingKey,
    preferredScriptKey: string | null = null,
  ) => {
    if (!settings[setting]) return false;
    if (setting === "autoExitAccent" && !caretIsInsideAccent(field)) {
      return false;
    }

    let scriptKey: string | null = null;
    if (
      setting === "autoExitSuperscript" ||
      setting === "autoExitSubscript"
    ) {
      const context = getScriptCaretContext(field);
      if (context?.kind === "operator-limit") return false;
      scriptKey = preferredScriptKey ?? context?.autoExitKey ?? null;
      if (scriptAutoExitWasConsumed(field, scriptKey)) return false;
    }

    const previousPosition = field.position;
    const changed = field.executeCommand(
      setting === "autoExitAccent" ? "moveToNextChar" : "moveAfterParent",
    );
    const moved = Boolean(changed || field.position !== previousPosition);
    if (moved && scriptKey) markScriptAutoExitConsumed(field, scriptKey);
    return moved;
  };

  if (capturedSetting) {
    return moveOne(capturedSetting, capturedScriptKey ?? null);
  }

  let moved = false;
  for (let attempt = 0; attempt < 4; attempt += 1) {
    const setting = getCaretAutoExitSetting(field);
    if (!setting) break;
    if (!moveOne(setting)) break;
    moved = true;
  }
  return moved;
}

const previousTargetToolbarCommandIds = new Set([
  "scripts",
  "lower-script",
  "upper-script",
  "power",
  "subscript",
  "degree",
  "overline",
  "underline",
  "hat",
  "widehat",
  "tilde",
  "widetilde",
  "dotaccent",
  "ddotaccent",
  "checkaccent",
  "breveaccent",
  "acuteaccent",
  "graveaccent",
  "ringaccent",
  "overrightarrow",
  "overleftarrow",
  "overleftrightarrow",
  "overleftharpoon",
  "overrightharpoon",
  "overarc",
  "overgroup",
  "overparen",
  "overlinesegment",
  "underarc",
  "undergroup",
  "underparen",
  "underleftarrow",
  "underrightarrow",
  "underleftrightarrow",
  "underlinesegment",
  "overbrace",
  "underbrace",
  "overset",
  "underset",
  "overunderset",
  "stackrel",
  "stackbin",
  "vector",
  "unitvector",
  "timederivative",
  "timesecond",
  "boxed",
  "evalbar",
  "boldsymbol",
  "blackboard-bold",
  "math-italic",
  "math-roman",
  "math-sans",
  "math-typewriter",
  "math-calligraphic",
  "math-script",
  "math-fraktur",
  "bold-symbol",
  "math-normal",
]);

const previousTargetToolbarCommandPattern =
  /^\\(?:acute|grave|dot|ddot|dddot|ddddot|tilde|widetilde|bar|breve|check|hat|widehat|vec|overline|underline|overrightarrow|overleftarrow|overleftrightarrow|overleftharpoon|overrightharpoon|overarc|overgroup|overparen|overlinesegment|underarc|undergroup|underparen|underleftarrow|underrightarrow|underleftrightarrow|underlinesegment|mathring|boxed|mathbf|boldsymbol|mathbb|mathit|mathrm|mathsf|mathtt|mathcal|mathscr|mathfrak|mathnormal)$/;

const nonTargetablePreviousLatex =
  /^(?:[+\-*/=<>:,;.!?]+|\\(?:pm|mp|times|div|cdot|ast|star|circ|bullet|cap|cup|land|lor|leq?|geq?|neq?|approx|equiv|sim|simeq|cong|propto|in|notin|subseteq?|supseteq?|parallel|perp|mid|nmid|to|leftarrow|rightarrow|leftrightarrow|Rightarrow|Leftarrow|Leftrightarrow))$/;

function commandTargetsPreviousExpression(command: LatexCommand) {
  return (
    previousTargetToolbarCommandIds.has(command.id) ||
    previousTargetToolbarCommandPattern.test(command.command.trim())
  );
}

function previousExpressionSelection(field: MathfieldElement) {
  if (
    field.mode !== "math" ||
    !field.selectionIsCollapsed ||
    field.position <= 0 ||
    hasRawLatexInput(field)
  ) {
    return null;
  }

  const originalSelection = captureSelection(field);
  const extended = field.executeCommand("extendSelectionBackward");
  if (!extended || field.selectionIsCollapsed) {
    field.selection = originalSelection;
    return null;
  }

  const selectedLatex = field.getValue(field.selection).trim();
  if (
    !selectedLatex ||
    selectedLatex === "\\placeholder{}" ||
    nonTargetablePreviousLatex.test(selectedLatex)
  ) {
    field.selection = originalSelection;
    return null;
  }

  return {
    latex: selectedLatex,
    selection: captureSelection(field),
  };
}

function templateForSelection(
  command: LatexCommand,
  selectedLatex: string,
): string {
  if (!selectedLatex) return command.insertTemplate;

  switch (command.id) {
    case "scripts":
      return selectedLatex + "_{\\placeholder{}}^{\\placeholder{}}";
    case "lower-script":
    case "subscript":
      return selectedLatex + "_{\\placeholder{}}";
    case "upper-script":
    case "power":
      return selectedLatex + "^{\\placeholder{}}";
    case "degree":
      return selectedLatex + "^{\\circ}";
    case "overset":
    case "underset":
    case "stackrel":
    case "stackbin":
      return (
        replaceTextualPlaceholder(command.insertTemplate, 1, selectedLatex) ??
        command.insertTemplate
      );
    case "overunderset":
      return (
        replaceTextualPlaceholder(command.insertTemplate, 2, selectedLatex) ??
        command.insertTemplate
      );
    case "sum":
    case "series":
      return "\\sum_{\\placeholder{}}^{\\placeholder{}} " + selectedLatex;
    case "prod":
    case "productseries":
      return "\\prod_{\\placeholder{}}^{\\placeholder{}} " + selectedLatex;
    case "int":
      return "\\int_{\\placeholder{}}^{\\placeholder{}} " + selectedLatex + "\\,\\mathrm{d}\\placeholder{}";
    case "intplain":
      return "\\int " + selectedLatex + "\\,\\mathrm{d}\\placeholder{}";
    case "lineintegral":
      return "\\int_{C} " + selectedLatex + "\\,\\mathrm{d}s";
    case "surfaceintegral":
      return "\\iint_{S} " + selectedLatex + "\\,\\mathrm{d}S";
    case "volumeintegral":
      return "\\iiint_{V} " + selectedLatex + "\\,\\mathrm{d}V";
    case "closed-surface-integral":
      return "\\oiint_{S} " + selectedLatex + "\\,\\mathrm{d}S";
    case "closed-volume-integral":
      return "\\oiiint_{V} " + selectedLatex + "\\,\\mathrm{d}V";
    case "frac":
    case "smallfrac":
    case "displayfrac":
      return command.command + "{" + selectedLatex + "}{\\placeholder{}}";
    case "sqrt":
      return "\\sqrt{" + selectedLatex + "}";
    case "parentheses":
      return "\\left(" + selectedLatex + "\\right)";
    case "brackets":
      return "\\left[" + selectedLatex + "\\right]";
    case "braces":
      return "\\left\\{" + selectedLatex + "\\right\\}";
    case "absolute":
      return "\\left|" + selectedLatex + "\\right|";
    default:
      return command.insertTemplate.replace("\\placeholder{}", selectedLatex);
  }
}

function FormulaField(props: FormulaFieldProps) {
  const hostRef = useRef<HTMLDivElement>(null);
  const fieldRef = useRef<MathfieldElement | null>(null);
  const registeredLineIdRef = useRef(props.lineId);
  const syncFrameSizeRef = useRef<(() => void) | null>(null);
  const defaultInlineShortcutsRef =
    useRef<VisualTexInlineShortcutDefinitions | null>(null);
  const lastSnapshotRef = useRef<ReturnType<typeof captureFieldSnapshot> | null>(null);
  const compositionStartRef = useRef<ReturnType<typeof captureFieldSnapshot> | null>(null);
  const sizingZoomRef = useRef(props.zoom);
  const sizingRowVerticalInsetRef = useRef(props.formulaRowVerticalInset);
  const propsRef = useRef(props);
  propsRef.current = props;

  useLayoutEffect(() => {
    const host = hostRef.current;
    if (!host) return;

    const initialLineId = propsRef.current.lineId;
    registeredLineIdRef.current = initialLineId;
    const lineId = initialLineId;
    const field = new MathfieldElement();
    MathfieldElement.locale = propsRef.current.language === "en" ? "en" : "zh-cn";
    field.value = propsRef.current.latex;
    field.className = "visual-mathfield";
    field.smartMode = false;
    field.smartFence = propsRef.current.autoPairDelimiters;
    field.readOnly = propsRef.current.readOnly;
    // VisualTeX handles superscript and subscript auto-exit independently.
    // Keep MathLive's built-in superscript heuristic disabled so it cannot
    // move out of an empty superscript before the first character is inserted.
    field.smartSuperscript = false;
    field.dataset.visualtexAutoExitSuperscript = String(
      propsRef.current.inputBehavior.autoExitSuperscript,
    );
    field.dataset.visualtexAutoExitSubscript = String(
      propsRef.current.inputBehavior.autoExitSubscript,
    );
    field.dataset.visualtexAutoExitAccent = String(
      propsRef.current.inputBehavior.autoExitAccent,
    );
    field.dataset.visualtexAutoExitWrapperCommand = String(
      propsRef.current.inputBehavior.autoExitWrapperCommand,
    );
    field.popoverPolicy = "auto";
    field.maxMatrixCols = 10;
    field.setAttribute("math-virtual-keyboard-policy", "manual");
    const isEn = propsRef.current.language === "en";
    field.setAttribute(
      "aria-label",
      isEn
        ? "Formula line " + (propsRef.current.index + 1)
        : "第 " + (propsRef.current.index + 1) + " 行公式",
    );
    field.removeAttribute("placeholder");
    const initialMetrics = formulaRowHeightMetrics(
      propsRef.current.latex,
      propsRef.current.zoom,
      propsRef.current.formulaRowVerticalInset,
    );
    const initialRowHeight = predictedFormulaRowHeight(
      propsRef.current.latex,
      propsRef.current.zoom,
      propsRef.current.formulaRowVerticalInset,
      host,
    );
    let initialHeightFloor = initialRowHeight;
    field.style.fontSize = initialMetrics.fontSize + "px";
    field.classList.toggle(
      "is-simple-formula",
      !initialMetrics.hasTallStructure,
    );
    field.style.height = initialRowHeight + "px";
    field.style.minHeight = initialRowHeight + "px";
    host.closest<HTMLElement>(".formula-line")?.style.setProperty(
      "--formula-row-height",
      initialRowHeight + "px",
    );
    installMathLiveContourIntegralShadowStyle(field);
    installCustomSymbolShadowStyle(field);
    installVisualTexStructuralPlaceholderStyle(field);
    applyVisualTexFormulaFonts(
      field,
      propsRef.current.formulaLetterFont,
      propsRef.current.formulaChineseFont,
    );

    let resizeFrame = 0;
    const measureFrameSize = () => {
      const metrics = formulaRowHeightMetrics(
        field.value,
        propsRef.current.zoom,
        propsRef.current.formulaRowVerticalInset,
      );
      const { fontSize, hasTallStructure, baseHeight, verticalPadding } =
        metrics;
      const content = field.shadowRoot?.querySelector<HTMLElement>(
        '[part="content"]',
      );
      const formulaRects = content
        ? Array.from(
            content.querySelectorAll<HTMLElement>(
              "[data-atom-id], .ML__base, .ML__mfrac, .ML__sqrt, " +
                ".ML__op-group, .ML__vlist",
            ),
          )
            .map((node) => node.getBoundingClientRect())
            .filter((rect) => rect.height > 0 && rect.width >= 0)
        : [];
      const formulaHeight = formulaRects.length
        ? Math.max(...formulaRects.map((rect) => rect.bottom)) -
          Math.min(...formulaRects.map((rect) => rect.top))
        : fontSize;
      const measuredHeight = Math.ceil(
        Math.max(baseHeight, formulaHeight + verticalPadding),
      );
      if (measuredHeight >= initialHeightFloor - 1) {
        initialHeightFloor = 0;
      }
      const nextHeight = Math.max(measuredHeight, initialHeightFloor);

      const shouldBeSimple = !hasTallStructure;
      if (field.classList.contains("is-simple-formula") !== shouldBeSimple) {
        field.classList.toggle("is-simple-formula", shouldBeSimple);
      }
      const nextHeightPx = nextHeight + "px";
      if (field.style.height !== nextHeightPx) field.style.height = nextHeightPx;
      if (field.style.minHeight !== nextHeightPx) {
        field.style.minHeight = nextHeightPx;
      }
      const row = host.closest<HTMLElement>(".formula-line");
      if (row?.style.getPropertyValue("--formula-row-height") !== nextHeightPx) {
        row?.style.setProperty("--formula-row-height", nextHeightPx);
      }
    };
    const syncFrameSize = (measureImmediately = false) => {
      window.cancelAnimationFrame(resizeFrame);
      if (measureImmediately) {
        // Shadow-DOM mutations can expose a fully rendered tall structure
        // before ResizeObserver/rAF updates the host row height. Measure once
        // in the same mutation turn so unrelated UI state changes cannot be
        // the first event that expands the formula row.
        measureFrameSize();
      }
      resizeFrame = window.requestAnimationFrame(measureFrameSize);
    };
    syncFrameSizeRef.current = syncFrameSize;
    const handleLayoutRefresh = () => {
      syncFrameSize();
    };
    window.addEventListener(
      EDITOR_LAYOUT_REFRESH_EVENT,
      handleLayoutRefresh,
    );

    const imeGuard = new ImeCompositionGuard();
    let pendingAutoExitSetting: InputBehaviorSettingKey | null = null;
    let pendingAutoExitScriptKey: string | null = null;
    const capturePendingAutoExit = () => {
      pendingAutoExitSetting = getCaretAutoExitSetting(field);
      const context = getScriptCaretContext(field);
      pendingAutoExitScriptKey =
        pendingAutoExitSetting === "autoExitSuperscript" ||
        pendingAutoExitSetting === "autoExitSubscript"
          ? context?.autoExitKey ?? null
          : null;
    };
    const clearPendingAutoExit = () => {
      pendingAutoExitSetting = null;
      pendingAutoExitScriptKey = null;
    };
    let restoringRawCommandAnchor = false;
    let compositionDeleteObserved = false;
    let suppressPostCompositionDeleteUntil = 0;

    const emitEdit = (
      before: ReturnType<typeof captureFieldSnapshot>,
      after: ReturnType<typeof captureFieldSnapshot>,
      editKind: EditKind,
      source: FormulaEditSource,
    ) => {
      lastSnapshotRef.current = after;
      if (before.latex === after.latex) return;
      propsRef.current.onEdit(
        {
          lineId,
          beforeLatex: before.latex,
          afterLatex: after.latex,
          beforeSelection: before.selection,
          afterSelection: after.selection,
          editKind,
          source,
        },
        field,
      );
      field.resetUndo();
    };
    const handleCompositionStart = () => {
      compositionDeleteObserved = false;
      suppressPostCompositionDeleteUntil = 0;
      capturePendingAutoExit();
      propsRef.current.onCommitPending();
      imeGuard.compositionStart();
      const liveCompositionStart = captureFieldSnapshot(field);
      compositionStartRef.current = liveCompositionStart;

    };
    const handleCompositionEnd = (event: CompositionEvent) => {
      const cancelledByCompositionDelete =
        event.data === "" && compositionDeleteObserved;
      const before =
        compositionStartRef.current ??
        lastSnapshotRef.current ??
        captureFieldSnapshot(field);
      imeGuard.compositionEnd(event.timeStamp);
      suppressPostCompositionDeleteUntil = cancelledByCompositionDelete
        ? event.timeStamp + 160
        : 0;
      compositionDeleteObserved = false;

      // Cancelling an uncommitted macOS IME candidate with Backspace can make
      // WKWebView/MathLive apply the same physical key to the confirmed formula.
      // The composition contains no committed text in this case, so restore the
      // exact pre-composition formula and selection instead of recording a delete.
      if (event.data === "" && field.value !== before.latex) {
        field.setValue(before.latex, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        const restored = clampSelection(before.selection, field.lastOffset);
        field.selection = restored;
        const restoredRange = restored.ranges.at(-1);
        if (restoredRange) field.position = restoredRange[1];
        field.resetUndo();
        lastSnapshotRef.current = captureFieldSnapshot(field);
        clearPendingAutoExit();
        compositionStartRef.current = null;
        syncFrameSize();
        return;
      }

      let after = captureFieldSnapshot(field);
      clearPendingAutoExit();
      compositionStartRef.current = null;
      emitEdit(before, after, "composition", "keyboard");
      syncFrameSize();
    };
  const handleBeforeInput = (event: InputEvent) => {
    if (restoringRawCommandAnchor) return;
      if (
        event.inputType === "deleteContentBackward" &&
        Boolean(rawCommandAnchors.get(field))
      ) {
        const anchor = rawCommandAnchors.get(field);
        const rawInput = rawLatexInput(field);
        if (anchor && Array.from(rawInput).length <= 1) {
          event.preventDefault();
          event.stopImmediatePropagation();
          const before = captureFieldSnapshot(field);
          restoringRawCommandAnchor = true;
          try {
            restoreCancelledRawCommandAnchor(field, anchor);
          } finally {
            restoringRawCommandAnchor = false;
          }
          rawCommandAnchors.delete(field);
          const after = captureFieldSnapshot(field);
          emitEdit(before, after, "delete-backward", "keyboard");
          propsRef.current.onInputActivity(field);
          syncFrameSize();
          return;
        }
      }
      if (
        event.inputType === "deleteContentBackward" &&
        event.timeStamp <= suppressPostCompositionDeleteUntil
      ) {
        suppressPostCompositionDeleteUntil = 0;
        event.preventDefault();
        event.stopImmediatePropagation();
        return;
      }
      if (!event.isComposing && !imeGuard.isComposing()) {
        if (isSingleDirectInput(event, field)) {
          // Keep the script type captured during keydown when WebKit's
          // beforeinput geometry is temporarily incomplete. A non-null
          // beforeinput result may refine it, but null must not erase it.
          const nextSetting = getCaretAutoExitSetting(field);
          if (nextSetting) {
            pendingAutoExitSetting = nextSetting;
            if (
              nextSetting === "autoExitSuperscript" ||
              nextSetting === "autoExitSubscript"
            ) {
              pendingAutoExitScriptKey =
                getScriptCaretContext(field)?.autoExitKey ??
                pendingAutoExitScriptKey;
            }
          }
        } else {
          clearPendingAutoExit();
        }
      }
    };
  const handleInput = (event: Event) => {
    if (restoringRawCommandAnchor) return;
      if (imeGuard.isComposing()) {
        if (
          event instanceof InputEvent &&
          event.inputType === "deleteCompositionText"
        ) {
          compositionDeleteObserved = true;
        }
        return;
      }
      const before = lastSnapshotRef.current ?? captureFieldSnapshot(field);
      const isDirectSingleInput =
        event instanceof InputEvent && isSingleDirectInput(event, field);
      // Direct single-character script/accent semantics are owned by the
      // VisualTeX MathLive fork. Only non-direct paths may consume a pending
      // outer auto-exit captured before composition/programmatic insertion.
      const autoExitSetting = isDirectSingleInput
        ? null
        : pendingAutoExitSetting;
      const autoExitScriptKey = isDirectSingleInput
        ? null
        : pendingAutoExitScriptKey;
      let autoExitMoved = false;
      if (
        autoExitSetting &&
        propsRef.current.inputBehavior[autoExitSetting]
      ) {
        autoExitMoved = moveCaretThroughEnabledAutoExitContainers(
          field,
          propsRef.current.inputBehavior,
          autoExitSetting,
          autoExitScriptKey,
        );
      }
      clearPendingAutoExit();
    normalizeCompletedDifferentialDisplay(field);
    const inputType =
      event instanceof InputEvent ? event.inputType || "insertText" : "insertText";
    const after = captureFieldSnapshot(field);
      emitEdit(
        before,
        after,
        inferEditKind(inputType, before.selection),
        inferEditSource(inputType),
      );
      propsRef.current.onInputActivity(field);
      syncFrameSize();
      window.requestAnimationFrame(() => {
        if (!field.isConnected) return;
        const deferredBefore =
          lastSnapshotRef.current ?? captureFieldSnapshot(field);
        if (
          normalizeCompletedDifferentialDisplay(field) || autoExitMoved
        ) {
          emitEdit(
            deferredBefore,
            captureFieldSnapshot(field),
            "replace",
            "keyboard",
          );
        }
        propsRef.current.onInputActivity(field);
      });
    };
    const handleSelectionChange = () => {
      syncStructuralPlaceholderSelection(field);
      syncPostOperatorCaretSpacing(field);
      syncFrameSize();
      const selection = captureSelection(field);
      propsRef.current.onSelectionChange(lineId, selection);
      if (imeGuard.isComposing() || !lastSnapshotRef.current) return;
      // MathLive emits selection-change after mutating the model but before
      // input. Keep the pre-edit selection paired with the pre-edit LaTeX;
      // otherwise Undo restores a post-edit offset (outside an empty cell).
      if (normalizeChineseLatex(field.value) !== lastSnapshotRef.current.latex) {
        return;
      }
      lastSnapshotRef.current = {
        ...lastSnapshotRef.current,
        selection,
      };
    };
    const handleFocus = () => {
      if (field.classList.contains(visualTexSourcePreviewClass)) {
        field.blur();
        queueMicrotask(() => {
          document
            .querySelector<HTMLElement>(".source-panel .cm-content")
            ?.focus({ preventScroll: true });
        });
        return;
      }
      propsRef.current.onFocus(propsRef.current.index, field);
      lastSnapshotRef.current = captureFieldSnapshot(field);
    };
    const handleBlur = () => {
      rawCommandAnchors.delete(field);

      propsRef.current.onCommitPending();
    };
    const scheduleInputActivity = () => {
      window.requestAnimationFrame(() => {
        if (!field.isConnected) return;
        propsRef.current.onInputActivity(field);
      });
    };
  const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") rawCommandAnchors.delete(field);
      if (
        event.key === "Backspace" &&
        rawCommandAnchors.has(field) &&
        Array.from(rawLatexInput(field)).length <= 1
      ) {
        event.preventDefault();
        event.stopImmediatePropagation();
        const anchor = rawCommandAnchors.get(field);
        if (anchor) {
          const before = captureFieldSnapshot(field);
          restoringRawCommandAnchor = true;
          try {
            restoreCancelledRawCommandAnchor(field, anchor);
          } finally {
            restoringRawCommandAnchor = false;
          }
          rawCommandAnchors.delete(field);
          const after = captureFieldSnapshot(field);
          emitEdit(before, after, "delete-backward", "keyboard");
          propsRef.current.onInputActivity(field);
          syncFrameSize();
        }
        return;
      }

      const imeDecision = imeGuard.keyDown(event, event.timeStamp);
      if (imeDecision === "composition") return;
      if (imeDecision === "post-composition-enter") {
        // macOS WebKit can replay the Enter used to confirm an IME candidate
        // after compositionend. Consume only that one synthetic follow-up so
        // VisualTeX does not also create a new formula line.
        event.preventDefault();
        event.stopPropagation();
        return;
      }
      if (compositionStartRef.current) {
        compositionStartRef.current = null;
        lastSnapshotRef.current = captureFieldSnapshot(field);
      }

      const accentNavigationDirection =
        !event.metaKey &&
        !event.ctrlKey &&
        !event.altKey &&
        !event.shiftKey &&
        event.key === "ArrowLeft"
          ? "left"
          : !event.metaKey &&
              !event.ctrlKey &&
              !event.altKey &&
              !event.shiftKey &&
              event.key === "ArrowRight"
            ? "right"
            : null;
      if (
        accentNavigationDirection &&
        selectAdjacentAccentPlaceholder(field, accentNavigationDirection)
      ) {
        event.preventDefault();
        event.stopImmediatePropagation();
        propsRef.current.onInputActivity(field);
        return;
      }

      const capturesDirectMathInput =
        !event.isComposing &&
        !event.metaKey &&
        !event.ctrlKey &&
        !event.altKey &&
        field.mode !== "latex" &&
        event.key !== "\\" &&
        Array.from(event.key).length === 1;
      if (capturesDirectMathInput) {
        // Direct single-character script/accent auto-exit is owned entirely by
        // the VisualTeX MathLive fork. Wrapper/package paths intercept their
        // own input before MathLive emits a normal math input event.
        clearPendingAutoExit();
      } else if (
        event.key !== "Shift" &&
        event.key !== "Control" &&
        event.key !== "Alt" &&
        event.key !== "Meta"
      ) {
        clearPendingAutoExit();
      }

      propsRef.current.onKeyDown(propsRef.current.index, event, field);
    };
    const handleKeyUp = (event: KeyboardEvent) => {
    };
    const suppressMathLiveContextMenu = (event: Event) => {
      event.preventDefault();
      event.stopImmediatePropagation();
      const pointer = event as MouseEvent;
      propsRef.current.onContextMenu?.(pointer.clientX, pointer.clientY);
    };
    const handlePaste = (event: ClipboardEvent) => {
      const clipboard = event.clipboardData;
      if (!clipboard) return;

      const multiLinePayload = clipboard.getData(
        VISUALTEX_MULTILINE_LATEX_CLIPBOARD_TYPE,
      );
      if (multiLinePayload && propsRef.current.onPasteLatexLines) {
        try {
          const parsed = JSON.parse(multiLinePayload) as {
            lines?: unknown;
          };
          const lines = Array.isArray(parsed.lines)
            ? parsed.lines.filter(
                (line): line is string => typeof line === "string",
              )
            : [];
          if (lines.length > 1) {
            event.preventDefault();
            event.stopImmediatePropagation();
            propsRef.current.onCommitPending();
            propsRef.current.onFocus(propsRef.current.index, field);
            propsRef.current.onPasteLatexLines(
              propsRef.current.lineId,
              field,
              lines,
            );
            return;
          }
        } catch {
          // Ignore malformed custom clipboard data and fall back to MathLive.
        }
      }

      if (!propsRef.current.onPasteImage) return;
      const item = Array.from(clipboard.items).find(
        (candidate) =>
          candidate.kind === "file" && candidate.type.startsWith("image/"),
      );
      const image = item?.getAsFile() ??
        Array.from(clipboard.files).find((file) => file.type.startsWith("image/"));
      if (!image) return;

      event.preventDefault();
      event.stopPropagation();
      propsRef.current.onCommitPending();
      propsRef.current.onFocus(propsRef.current.index, field);

      const selection = field.selection;
      propsRef.current.onPasteImage(image, {
        lineId: propsRef.current.lineId,
        ranges: selection.ranges.map(
          ([start, end]) => [start, end] as [number, number],
        ),
        direction: selection.direction ?? "none",
      });
    };
    const handlePointerSelectionEnd = () => {
      if (!visualTexPointerSelectingFields.delete(field)) return;
      field.classList.remove(visualTexPointerSelectingClass);
      window.requestAnimationFrame(() => {
        if (!field.isConnected) return;
      });
    };
    const handlePointerDown = (event: PointerEvent) => {
      const opensContextMenu =
        event.button === 2 || (event.button === 0 && event.ctrlKey);
      if (opensContextMenu) return;
      if (event.button !== 0) return;
      const pointerModelOffset = field.getOffsetFromPoint(
        event.clientX,
        event.clientY,
        { bias: 0 },
      );
      const modelHitIsPlaceholder =
        pointerModelOffset > 0 &&
        field.getElementInfo(pointerModelOffset)?.latex?.trim() ===
          "\\placeholder{}";
      visualTexPointerSelectingFields.add(field);
      field.classList.add(visualTexPointerSelectingClass);
      rawCommandAnchors.delete(field);
      propsRef.current.onCommitPending();

      const content = field.shadowRoot?.querySelector<HTMLElement>(
        '[part="content"]',
      );
      const contentBounds = content?.getBoundingClientRect();
      const hostBounds = host.getBoundingClientRect();
      const clickedInsideHost =
        event.clientX >= hostBounds.left &&
        event.clientX <= hostBounds.right &&
        event.clientY >= hostBounds.top &&
        event.clientY <= hostBounds.bottom;
      if (!clickedInsideHost) return;

      const pointerPath = event.composedPath();
      const clickedStructuralPlaceholder =
        modelHitIsPlaceholder ||
        pointerPath.some(
          (target) =>
            target instanceof HTMLElement &&
            (target.classList.contains("visualtex-structural-placeholder") ||
              target.classList.contains("ML__placeholder")),
        );
      // MathLive already gives PlaceholderAtom its own model selection semantics.
      // Never let the top-level X-gap correction intercept a real placeholder
      // pointer hit: doing so collapses the selection and bypasses MathLive's
      // native setSelection(anchor - 1, anchor) path.
      const structuralBoundaryOffset = clickedStructuralPlaceholder
        ? null
        : structuralBoundaryOffsetFromPoint(
            field,
            event.clientX,
            event.clientY,
          );
      if (
        structuralBoundaryOffset !== null &&
        !event.altKey &&
        !event.ctrlKey &&
        !event.metaKey &&
        !event.shiftKey
      ) {
        event.preventDefault();
        event.stopPropagation();
        field.focus();
        field.selection = {
          ranges: [[structuralBoundaryOffset, structuralBoundaryOffset]],
          direction: "none",
        };
        field.position = structuralBoundaryOffset;
        field.shadowRoot
          ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
          ?.focus({ preventScroll: true });
        propsRef.current.onFocus(propsRef.current.index, field);
        return;
      }

      const hasVisibleFormula = Boolean(field.value.trim()) && contentBounds;
      const clickedInLeftBlankArea = hasVisibleFormula
        ? event.clientX < contentBounds.left - 6
        : false;
      const clickedInRightBlankArea = hasVisibleFormula
        ? event.clientX > contentBounds.right + 6
        : true;

      // Preserve MathLive's native hit testing on rendered formula atoms.
      // Blank row space maps to the nearest mathematical edge so centered and
      // right-aligned formulas remain easy to focus and edit.
      if (!clickedInLeftBlankArea && !clickedInRightBlankArea) return;

      event.preventDefault();
      event.stopPropagation();

      const target = clickedInLeftBlankArea ? 0 : field.lastOffset;
      field.focus();
      field.selection = {
        ranges: [[target, target]],
        direction: "none",
      };
      field.position = target;
      field.shadowRoot
        ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
        ?.focus({ preventScroll: true });
      propsRef.current.onFocus(propsRef.current.index, field);
    };
    host.replaceChildren(field);
    installMathLiveOptionMutationGuard(field);
    // WebView2 exposes a richer mounted macro dictionary than MathLive's
    // deferred pre-mount options (for example, it includes native \\bmod).
    // Merge compatibility aliases only after mount so Windows does not lose
    // those native commands, while the mutation guard still handles transient
    // model initialization safely.
    field.macros = composeCustomSymbolMacrosForMathfield(
      field,
      field.macros,
      VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS,
    );
    defaultInlineShortcutsRef.current = { ...field.inlineShortcuts };
    field.inlineShortcuts = resolveVisualTexInlineShortcuts(
      defaultInlineShortcutsRef.current,
      propsRef.current.inputBehavior.autoEscapeShortcuts,
    );
    field.menuItems = [];
    installMathLiveContourIntegralShadowStyle(field);
    installCustomSymbolShadowStyle(field);
    installVisualTexStructuralPlaceholderStyle(field);
    applyVisualTexFormulaFonts(
      field,
      propsRef.current.formulaLetterFont,
      propsRef.current.formulaChineseFont,
    );
    // MathLive mounts a pre-filled field with the whole formula selected.
    // Collapse that implicit selection so toolbar commands insert at the end
    // instead of unexpectedly replacing/wrapping the entire line.
    field.position = field.lastOffset;
    setMathLivePersistentTypingStyle(
      field,
      propsRef.current.persistentTypingStyle,
    );
    field.resetUndo();
    lastSnapshotRef.current = captureFieldSnapshot(field);
    fieldRef.current = field;
    propsRef.current.register(initialLineId, field);
    // Capture before MathLive's keyboard sink replaces a selected placeholder;
    // the bubble phase is already too late because the parent structure is gone.
    field.addEventListener("compositionstart", handleCompositionStart, true);
    field.addEventListener("compositionend", handleCompositionEnd);
    field.addEventListener("beforeinput", handleBeforeInput, true);
    field.addEventListener("input", handleInput);
    field.addEventListener("selection-change", handleSelectionChange);
    field.addEventListener("focus", handleFocus);
    field.addEventListener("blur", handleBlur);
    field.addEventListener("keydown", handleKeyDown, true);
    field.addEventListener("keyup", handleKeyUp, true);
    const keyboardSink =
      field.shadowRoot?.querySelector<HTMLElement>('[part="keyboard-sink"]') ??
      null;
    const imeDiagnosticEventTypes = [
      "keydown",
      "keypress",
      "compositionstart",
      "compositionupdate",
      "compositionend",
      "beforeinput",
      "input",
      "keyup",
    ] as const;
    const handleWindowImeDiagnosticEvent = (event: Event) => {
      if (!event.composedPath().includes(field)) return;
      traceVisualTexImeEvent("window.capture." + event.type, field, keyboardSink, event);
    };
    const handleFieldImeDiagnosticEvent = (event: Event) => {
      traceVisualTexImeEvent("field.capture." + event.type, field, keyboardSink, event);
    };
    const handleSinkImeDiagnosticEvent = (event: Event) => {
      traceVisualTexImeEvent("sink.capture." + event.type, field, keyboardSink, event);
    };
    if (VISUALTEX_IME_DIAGNOSTICS_ENABLED) {
      ensureVisualTexImeDiagnosticElement();
      for (const type of imeDiagnosticEventTypes) {
        window.addEventListener(type, handleWindowImeDiagnosticEvent, true);
        field.addEventListener(type, handleFieldImeDiagnosticEvent, true);
        keyboardSink?.addEventListener(type, handleSinkImeDiagnosticEvent, true);
      }
    }
    keyboardSink?.addEventListener("input", scheduleInputActivity, true);
    keyboardSink?.addEventListener("keyup", scheduleInputActivity, true);
    field.addEventListener("paste", handlePaste, true);
    field.shadowRoot?.addEventListener(
      "contextmenu",
      suppressMathLiveContextMenu,
      true,
    );
    host.addEventListener("contextmenu", suppressMathLiveContextMenu, true);
    host.addEventListener("pointerdown", handlePointerDown, true);
    window.addEventListener("pointerup", handlePointerSelectionEnd, true);
    window.addEventListener("pointercancel", handlePointerSelectionEnd, true);
    const content = field.shadowRoot?.querySelector<HTMLElement>('[part="content"]');
    const resizeObserver = content
      ? new ResizeObserver(() => {
          syncFrameSize();
        })
      : null;
    const inputMutationObserver = field.shadowRoot
      ? new MutationObserver(() => {
          markVisualTexFormulaGlyphFonts(field);
          syncStructuralPlaceholderSelection(field);
          syncPostOperatorCaretSpacing(field);
          syncFrameSize(true);
          scheduleInputActivity();
        })
      : null;
    if (content) resizeObserver?.observe(content);
    if (field.shadowRoot) {
      inputMutationObserver?.observe(field.shadowRoot, {
        childList: true,
        characterData: true,
        attributes: true,
        attributeFilter: ["class"],
        subtree: true,
      });
    }
    // The field is now connected, but React's layout effect still runs before
    // the browser's first paint. Resolve the actual MathLive geometry here so
    // a newly inserted line never paints with the CSS fallback or predicted
    // height and then visibly shrinks on the next animation frame.
    measureFrameSize();
    syncPostOperatorCaretSpacing(field);
    queueMicrotask(() => {
      if (!field.isConnected) return;
      measureFrameSize();
      syncPostOperatorCaretSpacing(field);
    });

    return () => {
      window.cancelAnimationFrame(resizeFrame);
      window.removeEventListener(
        EDITOR_LAYOUT_REFRESH_EVENT,
        handleLayoutRefresh,
      );
      resizeObserver?.disconnect();
      inputMutationObserver?.disconnect();
      syncFrameSizeRef.current = null;
      field.removeEventListener("compositionstart", handleCompositionStart, true);
      field.removeEventListener("compositionend", handleCompositionEnd);
      field.removeEventListener("beforeinput", handleBeforeInput, true);
      field.removeEventListener("input", handleInput);
      field.removeEventListener("selection-change", handleSelectionChange);
      field.removeEventListener("focus", handleFocus);
      field.removeEventListener("blur", handleBlur);
      field.removeEventListener("keydown", handleKeyDown, true);
      field.removeEventListener("keyup", handleKeyUp, true);
      if (VISUALTEX_IME_DIAGNOSTICS_ENABLED) {
        for (const type of imeDiagnosticEventTypes) {
          window.removeEventListener(type, handleWindowImeDiagnosticEvent, true);
          field.removeEventListener(type, handleFieldImeDiagnosticEvent, true);
          keyboardSink?.removeEventListener(type, handleSinkImeDiagnosticEvent, true);
        }
      }
      keyboardSink?.removeEventListener("input", scheduleInputActivity, true);
      keyboardSink?.removeEventListener("keyup", scheduleInputActivity, true);
      field.removeEventListener("paste", handlePaste, true);
      field.shadowRoot?.removeEventListener(
        "contextmenu",
        suppressMathLiveContextMenu,
        true,
      );
      host.removeEventListener("contextmenu", suppressMathLiveContextMenu, true);
      host.removeEventListener("pointerdown", handlePointerDown, true);
      window.removeEventListener("pointerup", handlePointerSelectionEnd, true);
      window.removeEventListener("pointercancel", handlePointerSelectionEnd, true);
      visualTexPointerSelectingFields.delete(field);
      field.classList.remove(visualTexPointerSelectingClass);
      host.closest<HTMLElement>(".formula-line")?.style.removeProperty(
        "--formula-row-height",
      );
      propsRef.current.register(registeredLineIdRef.current, null, field);
      fieldRef.current = null;
      defaultInlineShortcutsRef.current = null;
      lastSnapshotRef.current = null;
      compositionStartRef.current = null;
      clearPendingAutoExit();
      rawCommandAnchors.delete(field);
      const caretRepaintFrame = visualTexCaretRepaintFrames.get(field);
      if (caretRepaintFrame) {
        window.cancelAnimationFrame(caretRepaintFrame);
        visualTexCaretRepaintFrames.delete(field);
      }
      field.classList.remove(visualTexCaretRepaintClass);
      host.replaceChildren();
    };
  }, []);

  useLayoutEffect(() => {
    const field = fieldRef.current;
    if (!field) return;
    const previousLineId = registeredLineIdRef.current;
    if (previousLineId === props.lineId) return;
    propsRef.current.register(previousLineId, null, field);
    propsRef.current.register(props.lineId, field);
    registeredLineIdRef.current = props.lineId;
  }, [props.lineId]);

  useLayoutEffect(() => {
    const field = fieldRef.current;
    if (!field) return;
    setMathLivePersistentTypingStyle(field, props.persistentTypingStyle);
  }, [
    props.persistentTypingStyle.bold,
    props.persistentTypingStyle.italic,
    props.persistentTypingStyle.color,
    props.persistentTypingStyle.backgroundColor,
  ]);

  useLayoutEffect(() => {
    const field = fieldRef.current;
    const host = hostRef.current;
    if (!field || !host) return;
    const zoomChanged = sizingZoomRef.current !== props.zoom;
    const rowVerticalInsetChanged =
      sizingRowVerticalInsetRef.current !== props.formulaRowVerticalInset;
    sizingZoomRef.current = props.zoom;
    sizingRowVerticalInsetRef.current = props.formulaRowVerticalInset;

    normalizeCompletedDifferentialDisplay(field);
    installCustomSymbolShadowStyle(field);

    // 本地输入仅因中文规范化而与 store 等值时，不重建 MathLive 模型；
    // 只更新事务基准，保留当前光标、选区和删除键内部状态。
    if (normalizeChineseLatex(field.value) === props.latex) {
      lastSnapshotRef.current = {
        latex: props.latex,
        selection: captureSelection(field),
      };
      if (!zoomChanged && !rowVerticalInsetChanged) return;
    } else {
      const currentLatex = normalizeChineseLatex(field.value);
      const applyExternalValue = () =>
        field.setValue(props.latex, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
      if (
        props.freshExternalSync &&
        needsCleanMathLiveModelForExternalSync(currentLatex, props.latex)
      ) {
        resetMathLiveModelRootForExternalSync(field);
      }
      applyExternalValue();
      if (
        props.freshExternalSync &&
        normalizeChineseLatex(field.value) !== props.latex &&
        resetMathLiveModelRootForExternalSync(field)
      ) {
        applyExternalValue();
      }
      installCustomSymbolShadowStyle(field);
      field.resetUndo();
    }

    const metrics = formulaRowHeightMetrics(
      props.latex,
      props.zoom,
      props.formulaRowVerticalInset,
    );
    const predictedHeight = predictedFormulaRowHeight(
      props.latex,
      props.zoom,
      props.formulaRowVerticalInset,
      host,
    );
    field.classList.toggle("is-simple-formula", !metrics.hasTallStructure);
    field.style.fontSize = metrics.fontSize + "px";
    field.style.height = predictedHeight + "px";
    field.style.minHeight = predictedHeight + "px";
    host.closest<HTMLElement>(".formula-line")?.style.setProperty(
      "--formula-row-height",
      predictedHeight + "px",
    );
    lastSnapshotRef.current = captureFieldSnapshot(field);
    syncFrameSizeRef.current?.();
  }, [props.formulaRowVerticalInset, props.latex, props.zoom]);

  useEffect(() => {
    if (fieldRef.current) {
      fieldRef.current.smartFence = props.autoPairDelimiters;
    }
  }, [props.autoPairDelimiters]);

  useEffect(() => {
    if (fieldRef.current) fieldRef.current.readOnly = props.readOnly;
  }, [props.readOnly]);

  useEffect(() => {
    const field = fieldRef.current;
    const defaults = defaultInlineShortcutsRef.current;
    if (!field || !defaults) return;
    field.inlineShortcuts = resolveVisualTexInlineShortcuts(
      defaults,
      props.inputBehavior.autoEscapeShortcuts,
    );
  }, [props.inputBehavior.autoEscapeShortcuts]);

  useEffect(() => {
    const field = fieldRef.current;
    if (!field) return;
    field.dataset.visualtexAutoExitSuperscript = String(
      props.inputBehavior.autoExitSuperscript,
    );
    field.dataset.visualtexAutoExitSubscript = String(
      props.inputBehavior.autoExitSubscript,
    );
    field.dataset.visualtexAutoExitAccent = String(
      props.inputBehavior.autoExitAccent,
    );
    field.dataset.visualtexAutoExitWrapperCommand = String(
      props.inputBehavior.autoExitWrapperCommand,
    );
  }, [
    props.inputBehavior.autoExitAccent,
    props.inputBehavior.autoExitSubscript,
    props.inputBehavior.autoExitSuperscript,
    props.inputBehavior.autoExitWrapperCommand,
  ]);

  useEffect(() => {
    const field = fieldRef.current;
    if (!field) return;
    applyVisualTexFormulaFonts(
      field,
      props.formulaLetterFont,
      props.formulaChineseFont,
    );
    syncFrameSizeRef.current?.();
    const frame = window.requestAnimationFrame(() => {
      if (!field.isConnected) return;
      syncFrameSizeRef.current?.();
    });
    return () => window.cancelAnimationFrame(frame);
  }, [props.formulaChineseFont, props.formulaLetterFont]);

  useEffect(() => {
    const field = fieldRef.current;
    if (!field) return;
    MathfieldElement.locale = props.language === "en" ? "en" : "zh-cn";
    const isEn = props.language === "en";
    field.setAttribute(
      "aria-label",
      isEn
        ? "Formula line " + (props.index + 1)
        : "第 " + (props.index + 1) + " 行公式",
    );
    field.removeAttribute("placeholder");
  }, [props.index, props.language]);

  return <div ref={hostRef} className="mathfield-host" />;
}

export const MathEditor = forwardRef<MathEditorHandle, Props>(
  function MathEditor(
    {
      lines,
      activeLineId,
      formulaAlignment,
      latexCodeFormat,
      zoom,
      persistentTypingStyle = {
        bold: false,
        italic: null,
        color: null,
        backgroundColor: null,
      },
      reuseLineSlots = false,
      readOnly = false,
      previewOnly = false,
      showLineModeControls = true,
      onPreviewActivate,
      draftError,
      onPasteImage,
      onCopyPng,
      onHistoryBusyChange,
      overlay,
    },
    ref,
  ) {
    const surfaceRef = useRef<HTMLDivElement>(null);
    const fieldRefs = useRef(new Map<string, MathfieldElement>());
    const linesRef = useRef(lines);
    const activeIndexRef = useRef(0);
    const activeLineIdRef = useRef<string | null>(activeLineId);
    const focusRequestRef = useRef(0);
    const previewOnlyRef = useRef(previewOnly);
    const greekLetterHotkeyLineIdRef = useRef<string | null>(null);
    const suppressedHistoryLineIdRef = useRef<string | null>(null);
    const multiLineSelectionRef = useRef<MultiLineSelectionState | null>(null);
    const documentSelectionRef = useRef<Record<string, MathSelectionSnapshot> | null>(null);
    const documentSelectionNavigationRef = useRef<string | null>(null);
    const lastSelectionTargetRef = useRef<MathEditorSelectionTarget | null>(null);
    const pointerSelectionSessionRef = useRef<PointerSelectionSession | null>(
      null,
    );
    const multiLineSelectedIdsRef = useRef(new Set<string>());
    const [, setMultiLineSelectionRevision] = useState(0);
    const pendingFocusRef = useRef<{
      lineId: string;
      latex: string | null;
      selection: MathSelectionSnapshot | null;
      moveToEnd: boolean;
      deferredRepair: boolean;
    } | null>(null);
    const [activeIndex, setActiveIndex] = useState(() =>
      Math.max(0, lines.findIndex((line) => line.id === activeLineId)),
    );
    const [fieldRenderEpoch, setFieldRenderEpoch] = useState(0);
    const [query, setQuery] = useState("");
    const [contextMenu, setContextMenu] = useState<{
      left: number;
      top: number;
    } | null>(null);
    const [contextMenuBusy, setContextMenuBusy] = useState(false);
    const queryRef = useRef("");
    const suppressedSuggestionRef = useRef<{
      lineId: string;
      value: string;
    } | null>(null);
    const usage = useEditorStore((state) => state.usage);
    const personalize = useEditorStore((state) => state.personalize);
    const suggestionCount = useEditorStore((state) => state.suggestionCount);
    const recordCommand = useEditorStore((state) => state.recordCommand);
    const language = useEditorStore((state) => state.language);
    const pngExportBackground = useEditorStore(
      (state) => state.pngExportBackground,
    );
    const autoPairDelimiters = useEditorStore(
      (state) => state.autoPairDelimiters,
    );
    const showLineNumbers = useEditorStore((state) => state.showLineNumbers);
    const formulaInsetLeft = useEditorStore((state) => state.formulaInsetLeft);
    const formulaInsetRight = useEditorStore((state) => state.formulaInsetRight);
    const formulaRowVerticalInset = useEditorStore(
      (state) => state.formulaRowVerticalInset,
    );
    const formulaLetterFont = useEditorStore((state) => state.formulaLetterFont);
    const formulaChineseFont = useEditorStore((state) => state.formulaChineseFont);
    const inputBehavior = useEditorStore((state) => state.inputBehavior);
    const latexFormatProfile = useEditorStore((state) => state.latexFormatProfile);
    const formulaHotkeyBindings = useFormulaHotkeyStore(
      (state) => state.bindings,
    );
    const customSymbolRevision = useCustomSymbolRevision();
    const isEn = language === "en";
    const interactionReadOnly = readOnly || previewOnly;
    const showLineModeMarkers =
      showLineModeControls && (!readOnly || previewOnly);
    previewOnlyRef.current = previewOnly;

    useEffect(() => {
      // Do not eagerly materialize every persisted custom-symbol SVG mask on
      // initial startup. Runtime edits still refresh the complete global style
      // so already-mounted non-shadow previews update immediately.
      installCustomSymbolGlobalStyle(undefined, customSymbolRevision > 0);
      fieldRefs.current.forEach((field) => refreshCustomSymbolMathfield(field));
    }, [customSymbolRevision]);

    linesRef.current = lines;
    const resolvedActiveLineId =
      lines.find((line) => line.id === activeLineId)?.id ?? lines[0]?.id ?? null;
    activeLineIdRef.current = resolvedActiveLineId;
    // A plain row edit must not tear down every earlier explicit-alignment
    // field. Include the row position for Office's reusable line slots, where
    // inserting a row can move a marker to a different mounted mathfield.
    const alignmentRowsSignature = lines.flatMap((line, index) =>
      hasVisualTexAlignmentMarker(line.latex)
        ? [`${index}\u0000${line.id}\u0000${line.latex}`]
        : [],
    ).join("\u0001");

    const resolveExplicitAlignmentVisualAnchor = (marker: HTMLElement) => {
      const markerBounds = marker.getBoundingClientRect();
      const parent = marker.parentElement;
      if (!parent) return markerBounds.left;

      for (
        let sibling = marker.nextElementSibling;
        sibling;
        sibling = sibling.nextElementSibling
      ) {
        const candidates: Element[] = sibling.hasAttribute("data-atom-id")
          ? [sibling, ...sibling.querySelectorAll("[data-atom-id]")]
          : [...sibling.querySelectorAll("[data-atom-id]")];
        for (const candidate of candidates) {
          if (!(candidate instanceof HTMLElement)) continue;
          const bounds = candidate.getBoundingClientRect();
          if (bounds.width <= 0.01 || bounds.height <= 0.01) continue;
          return bounds.left;
        }
      }

      return markerBounds.left;
    };

    const refreshExplicitAlignmentLayout = () => {
      const registered = Array.from(fieldRefs.current.values());
      for (const field of registered) {
        const host = field.closest<HTMLElement>(".mathfield-host");
        host?.classList.remove("has-explicit-align-marker");
        field.style.removeProperty("margin-left");
        const previousMarkerCount = Number.parseInt(
          field.dataset.visualtexAlignmentMarkerCount ?? "0",
          10,
        );
        if (Number.isFinite(previousMarkerCount)) {
          for (let markerIndex = 0; markerIndex < previousMarkerCount; markerIndex += 1) {
            field.style.removeProperty(`--visualtex-align-marker-${markerIndex}-left`);
            field.style.removeProperty(`--visualtex-align-marker-${markerIndex}-right`);
          }
        }
        delete field.dataset.visualtexAlignmentMarkerCount;
      }
      const hasOcrAlignmentMarkers = linesRef.current.some((line) =>
        hasVisualTexAlignmentMarker(line.latex),
      );
      if (
        latexFormatProfile.multilineEnvironment !== "align" &&
        !hasOcrAlignmentMarkers
      ) {
        return;
      }

      const marked = linesRef.current.flatMap((line) => {
        const field = fieldRefs.current.get(line.id);
        if (!field?.isConnected) return [];
        const host = field.closest<HTMLElement>(".mathfield-host");
        const markers = Array.from(
          field.shadowRoot?.querySelectorAll<HTMLElement>(
            `.${VISUALTEX_ALIGNMENT_MARKER_CLASS}`,
          ) ?? [],
        );
        if (!host || !markers.length) return [];
        const fieldBounds = field.getBoundingClientRect();
        const markerOffsets = markers.map(
          (marker) => resolveExplicitAlignmentVisualAnchor(marker) - fieldBounds.left,
        );
        const segmentWidths = markerOffsets.map((offset, index) =>
          Math.max(0, offset - (index === 0 ? 0 : markerOffsets[index - 1])),
        );
        segmentWidths.push(
          Math.max(0, fieldBounds.right - fieldBounds.left - markerOffsets.at(-1)!),
        );
        return [{ field, host, markers, segmentWidths }];
      });
      if (!marked.length) return;

      const maxMarkerCount = Math.max(...marked.map((entry) => entry.markers.length));
      const columnWidths = Array.from({ length: maxMarkerCount + 1 }, (_, column) =>
        Math.max(0, ...marked.map((entry) => entry.segmentWidths[column] ?? 0)),
      );
      // TeX align/aligned groups columns in right/left pairs. There is no gap
      // inside a pair, while consecutive pairs are separated by 2em. A single
      // marker therefore keeps the exact historical VisualTeX layout; extra
      // markers add only the inter-pair spacing needed by `&&`, `&&&`, ... .
      const fontSizePx = Number.parseFloat(getComputedStyle(marked[0].field).fontSize);
      const pairGap = Number.isFinite(fontSizePx) ? fontSizePx * 2 : 0;
      const pairGapCount = Math.floor(maxMarkerCount / 2);
      const alignedWidth =
        columnWidths.reduce((total, width) => total + width, 0) +
        pairGap * pairGapCount;

      for (const entry of marked) {
        const shadowRoot = entry.field.shadowRoot;
        if (shadowRoot) {
          let alignmentStyle = shadowRoot.getElementById(
            visualTexAlignmentMarkerStyleId,
          ) as HTMLStyleElement | null;
          if (!alignmentStyle) {
            alignmentStyle = document.createElement("style");
            alignmentStyle.id = visualTexAlignmentMarkerStyleId;
            shadowRoot.append(alignmentStyle);
          }
          // MathLive may rebuild its rendered atom tree after the math-field is
          // laid out. Keep compensation in field-owned CSS variables and a
          // stable shadow-root stylesheet instead of mutating transient marker
          // nodes directly. Top-level alignment markers are siblings, so the
          // Selectors-4 `of <selector>` form addresses the Nth marker rather
          // than the Nth arbitrary MathLive atom.
          alignmentStyle.textContent = Array.from(
            { length: maxMarkerCount },
            (_, markerIndex) => `
              .${VISUALTEX_ALIGNMENT_MARKER_CLASS}:nth-child(${markerIndex + 1} of .${VISUALTEX_ALIGNMENT_MARKER_CLASS}) {
                margin-left: var(--visualtex-align-marker-${markerIndex}-left, 0px) !important;
                margin-right: var(--visualtex-align-marker-${markerIndex}-right, 0px) !important;
              }
            `,
          ).join("\n");
        }

        const hostWidth = entry.host.getBoundingClientRect().width;
        const groupStart =
          formulaAlignment === "left"
            ? 0
            : formulaAlignment === "right"
              ? Math.max(0, hostWidth - alignedWidth)
              : Math.max(0, (hostWidth - alignedWidth) / 2);
        const fieldOffset = Math.max(
          0,
          groupStart + columnWidths[0] - entry.segmentWidths[0],
        );
        entry.host.classList.add("has-explicit-align-marker");
        entry.field.style.marginLeft = `${fieldOffset}px`;

        entry.field.dataset.visualtexAlignmentMarkerCount = String(entry.markers.length);
        entry.markers.forEach((_marker, markerIndex) => {
          let left = 0;
          // Apply each column's compensation immediately before its own marker.
          // Putting the next-column fill on the preceding marker's right margin
          // leaves one MathLive marker-atom width behind for consecutive `&&`
          // markers in WebView2/Chromium, which visibly shifts later anchors.
          if (markerIndex > 0) {
            left = Math.max(
              0,
              columnWidths[markerIndex] -
                (entry.segmentWidths[markerIndex] ?? 0),
            );
            if (markerIndex % 2 !== 0) left += pairGap;
          }
          entry.field.style.setProperty(
            `--visualtex-align-marker-${markerIndex}-left`,
            `${left}px`,
          );
          entry.field.style.setProperty(
            `--visualtex-align-marker-${markerIndex}-right`,
            "0px",
          );
        });
      }

      // MathLive's zero-width class atom can report a different inline box
      // origin when two markers are consecutive (`&&`). Reconcile the rendered
      // marker positions from left to right after the analytical column pass;
      // each non-negative correction also advances every later column.
      for (let markerIndex = 1; markerIndex < maxMarkerCount; markerIndex += 1) {
        const positioned = marked.flatMap((entry) => {
          const marker = entry.markers[markerIndex];
          return marker
            ? [{ entry, marker, left: marker.getBoundingClientRect().left }]
            : [];
        });
        if (!positioned.length) continue;
        const targetLeft = Math.max(...positioned.map((item) => item.left));
        for (const item of positioned) {
          const correction = Math.max(0, targetLeft - item.left);
          if (correction <= 0.01) continue;
          const property = `--visualtex-align-marker-${markerIndex}-left`;
          const current = Number.parseFloat(
            item.entry.field.style.getPropertyValue(property) || "0",
          );
          item.entry.field.style.setProperty(
            property,
            `${(Number.isFinite(current) ? current : 0) + correction}px`,
          );
        }
      }
    };

    useLayoutEffect(() => {
      let firstFrame = 0;
      let secondFrame = 0;
      const schedule = () => {
        window.cancelAnimationFrame(firstFrame);
        window.cancelAnimationFrame(secondFrame);
        firstFrame = window.requestAnimationFrame(() => {
          refreshExplicitAlignmentLayout();
          secondFrame = window.requestAnimationFrame(
            refreshExplicitAlignmentLayout,
          );
        });
      };
      schedule();
      window.addEventListener("resize", schedule);
      window.addEventListener(EDITOR_LAYOUT_REFRESH_EVENT, schedule);
      return () => {
        window.cancelAnimationFrame(firstFrame);
        window.cancelAnimationFrame(secondFrame);
        window.removeEventListener("resize", schedule);
        window.removeEventListener(EDITOR_LAYOUT_REFRESH_EVENT, schedule);
      };
    }, [
      customSymbolRevision,
      fieldRenderEpoch,
      formulaAlignment,
      formulaInsetLeft,
      formulaInsetRight,
      formulaLetterFont,
      formulaChineseFont,
      formulaRowVerticalInset,
      latexFormatProfile.multilineEnvironment,
      alignmentRowsSignature,
      showLineNumbers,
      zoom,
    ]);

    const captureEditorScrollSnapshot = (
      preferredLineId?: string | null,
    ): EditorScrollSnapshot | null => {
      const scroller =
        surfaceRef.current?.closest<HTMLElement>(".editor-pane-scroll") ?? null;
      if (!scroller) return null;
      const scrollerRect = scroller.getBoundingClientRect();
      const rows = Array.from(
        surfaceRef.current?.querySelectorAll<HTMLElement>(".formula-line") ?? [],
      );
      const preferredRow = preferredLineId
        ? rows.find((row) => row.dataset.lineId === preferredLineId) ?? null
        : null;
      const visibleRow = rows.find((row) => {
        const rect = row.getBoundingClientRect();
        return rect.bottom > scrollerRect.top && rect.top < scrollerRect.bottom;
      });
      const anchorRow = preferredRow ?? visibleRow ?? null;
      return {
        scroller,
        anchorLineId: anchorRow?.dataset.lineId ?? null,
        anchorOffset: anchorRow
          ? anchorRow.getBoundingClientRect().top - scrollerRect.top
          : 0,
        scrollTop: scroller.scrollTop,
      };
    };

    const applyEditorScrollSnapshot = (
      snapshot: EditorScrollSnapshot | null,
      targetLineId?: string | null,
      ensureTargetVisible = false,
    ) => {
      if (!snapshot?.scroller.isConnected) return;
      const { scroller } = snapshot;
      const scrollerRect = scroller.getBoundingClientRect();
      const anchorRow = snapshot.anchorLineId
        ? surfaceRef.current?.querySelector<HTMLElement>(
            `[data-line-id="${snapshot.anchorLineId}"]`,
          ) ?? null
        : null;
      if (anchorRow?.isConnected) {
        const currentOffset =
          anchorRow.getBoundingClientRect().top - scrollerRect.top;
        const offsetDelta = currentOffset - snapshot.anchorOffset;
        if (Math.abs(offsetDelta) > 0.5) {
          scroller.scrollTop += offsetDelta;
        }
      } else if (Math.abs(scroller.scrollTop - snapshot.scrollTop) > 0.5) {
        scroller.scrollTop = snapshot.scrollTop;
      }

      if (!ensureTargetVisible || !targetLineId) return;
      const targetRow = surfaceRef.current?.querySelector<HTMLElement>(
        `[data-line-id="${targetLineId}"]`,
      );
      if (!targetRow?.isConnected) return;
      const viewport = scroller.getBoundingClientRect();
      const target = targetRow.getBoundingClientRect();
      const inset = 8;
      const top = viewport.top + inset;
      const bottom = viewport.bottom - inset;
      const availableHeight = Math.max(0, bottom - top);
      if (target.height > availableHeight) {
        scroller.scrollTop += target.top - top;
      } else if (target.top < top) {
        scroller.scrollTop -= top - target.top;
      } else if (target.bottom > bottom) {
        scroller.scrollTop += target.bottom - bottom;
      }
    };

    const stabilizeEditorScroll = (
      snapshot: EditorScrollSnapshot | null,
      targetLineId?: string | null,
      ensureTargetVisible = false,
    ) => {
      const apply = () =>
        applyEditorScrollSnapshot(
          snapshot,
          targetLineId,
          ensureTargetVisible,
        );
      apply();
      queueMicrotask(apply);
      window.requestAnimationFrame(apply);
    };

    useEffect(() => {
      for (const field of fieldRefs.current.values()) {
        configureMathLiveCompletion(field, { usage, personalize, count: suggestionCount });
      }
    }, [lines, usage, personalize, suggestionCount, fieldRenderEpoch]);

    useEffect(() => {
      const surface = surfaceRef.current;
      const accepted = (event: Event) => {
        const detail = (event as CustomEvent<{ id: string; query: string }>).detail;
        if (detail?.id) recordCommand(detail.id, detail.query, "candidate");
      };
      surface?.addEventListener("visualtex-command-accepted", accepted);
      return () => surface?.removeEventListener("visualtex-command-accepted", accepted);
    }, [recordCommand]);

    const applyFocusState = (
      lineId: string,
      expectedLatex: string | null,
      selection: MathSelectionSnapshot | null,
      moveToEnd: boolean,
      requestId: number,
    ) => {
      if (
        requestId !== focusRequestRef.current ||
        activeLineIdRef.current !== lineId
      ) {
        return false;
      }

      const field = fieldRefs.current.get(lineId);
      if (!field?.isConnected) return false;

      if (
        expectedLatex !== null &&
        normalizeChineseLatex(field.value) !== expectedLatex
      ) {
        field.setValue(expectedLatex, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
      }
      field.resetUndo();

      // Establish the final model position before focusing MathLive's hidden
      // keyboard sink. WebKit can fail to repaint the caret when a focused
      // sink receives a programmatic selection change, leaving the correct
      // selection in the model but no visible caret until the next input.
      if (selection) {
        const clamped = clampSelection(selection, field.lastOffset);
        field.selection = clamped;
        const range = clamped.ranges[0];
        if (range && range[0] === range[1]) field.position = range[1];
      } else if (moveToEnd) {
        const end = field.lastOffset;
        field.selection = {
          ranges: [[end, end]],
          direction: "none",
        };
        field.position = end;
      }

      (field as HTMLElement).focus({ preventScroll: true });
      field.shadowRoot
        ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
        ?.focus({ preventScroll: true });
      return true;
    };

    const focusLine = (
      lineId: string,
      options: {
        latex?: string | null;
        selection?: MathSelectionSnapshot | null;
        moveToEnd?: boolean;
        deferredRepair?: boolean;
      } = {},
      remainingAttempts = 10,
      requestId = ++focusRequestRef.current,
    ) => {
      const index = linesRef.current.findIndex((line) => line.id === lineId);
      if (index < 0) return false;

      const expectedLatex = options.latex ?? null;
      const selection = options.selection ?? null;
      const moveToEnd = options.moveToEnd ?? false;
      const deferredRepair = options.deferredRepair ?? true;
      activeIndexRef.current = index;
      activeLineIdRef.current = lineId;
      setActiveIndex((current) => (current === index ? current : index));
      const currentState = useEditorStore.getState();
      if (currentState.activeLineId !== lineId) {
        currentState.setActiveLineId(lineId);
      }
      pendingFocusRef.current = {
        lineId,
        latex: expectedLatex,
        selection,
        moveToEnd,
        deferredRepair,
      };

      const apply = () =>
        applyFocusState(
          lineId,
          expectedLatex,
          selection,
          moveToEnd,
          requestId,
        );
      const finish = () => {
        if (!apply()) return false;
        pendingFocusRef.current = null;
        if (!deferredRepair) return true;
        const appliedField = fieldRefs.current.get(lineId);
        const repairSelection =
          appliedField?.isConnected ? captureSelection(appliedField) : null;
        const selectionStillMatchesRepair = () => {
          if (!repairSelection) return true;
          const currentField = fieldRefs.current.get(lineId);
          if (!currentField?.isConnected) return false;
          const currentSelection = captureSelection(currentField);
          return (
            currentSelection.direction === repairSelection.direction &&
            JSON.stringify(currentSelection.ranges) ===
              JSON.stringify(repairSelection.ranges)
          );
        };
        window.requestAnimationFrame(() => {
          // A pointer click, arrow key or drag that moved the selection after
          // the initial focus is authoritative user input. Never replay an old
          // focus snapshot over it merely because the LaTeX source is unchanged.
          if (!selectionStillMatchesRepair() || !apply()) return;
          window.setTimeout(() => {
            const currentField = fieldRefs.current.get(lineId);
            // Do not let the delayed focus repair overwrite input or selection
            // changes that started immediately after a toolbar insertion.
            if (
              !currentField?.isConnected ||
              currentField.mode !== "math" ||
              rawLatexInput(currentField) ||
              !selectionStillMatchesRepair() ||
              (expectedLatex !== null &&
                normalizeChineseLatex(currentField.value) !== expectedLatex)
            ) {
              return;
            }
            apply();
          }, 80);
        });
        return true;
      };

      if (finish()) return true;
      window.requestAnimationFrame(() => {
        if (finish() || remainingAttempts <= 0) return;
        window.setTimeout(
          () => focusLine(lineId, options, remainingAttempts - 1, requestId),
          16,
        );
      });
      return true;
    };

    const registerField = (
      lineId: string,
      field: MathfieldElement | null,
      expectedField?: MathfieldElement,
    ) => {
      if (field) {
        fieldRefs.current.set(lineId, field);
      } else if (
        !expectedField ||
        fieldRefs.current.get(lineId) === expectedField
      ) {
        fieldRefs.current.delete(lineId);
      }

      const pending = pendingFocusRef.current;
      if (field && pending?.lineId === lineId) {
        focusLine(lineId, {
          latex: pending.latex,
          selection: pending.selection,
          moveToEnd: pending.moveToEnd,
          deferredRepair: pending.deferredRepair,
        });
      }
    };

    const prepareFocusBeforeStructuralRemoval = (
      lineId: string,
      selection: MathSelectionSnapshot,
    ) => {
      const field = fieldRefs.current.get(lineId);
      if (!field?.isConnected) return false;

      // Transfer the native focus while both keyboard sinks are still in the
      // document. WKWebView drops a post-removal focus request made after the
      // keydown target has already been detached.
      focusRequestRef.current += 1;
      const index = linesRef.current.findIndex((line) => line.id === lineId);
      if (index >= 0) {
        activeIndexRef.current = index;
        activeLineIdRef.current = lineId;
        setActiveIndex((current) => (current === index ? current : index));
      }
      const currentState = useEditorStore.getState();
      if (currentState.activeLineId !== lineId) {
        currentState.setActiveLineId(lineId);
      }
      const clamped = clampSelection(selection, field.lastOffset);
      field.selection = clamped;
      const range = clamped.ranges[0];
      if (range && range[0] === range[1]) field.position = range[1];
      (field as HTMLElement).focus({ preventScroll: true });
      field.shadowRoot
        ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
        ?.focus({ preventScroll: true });
      return true;
    };

    const finalizeFocusAfterStructuralRemoval = (
      lineId: string,
      selection: MathSelectionSnapshot,
    ) => {
      queueMicrotask(() => {
        if (
          activeLineIdRef.current !== lineId ||
          !fieldRefs.current.get(lineId)?.isConnected
        ) {
          return;
        }
        const field = fieldRefs.current.get(lineId);
        if (!field?.isConnected) return;
        const clamped = clampSelection(selection, field.lastOffset);
        field.selection = clamped;
        const range = clamped.ranges[0];
        if (range && range[0] === range[1]) field.position = range[1];
        repaintMathLiveCaret(field);
      });
    };

    const refreshSuggestionQuery = (
      lineId: string,
      field: MathfieldElement,
      normalized: string,
    ) => {
      if (lineId !== activeLineIdRef.current) return;
      const rawCommandActive = hasRawLatexInput(field);
      const activeCommandQuery = rawCommandActive
        ? ""
        : trailingCommandQuery(field, normalized);
      const suppressed = suppressedSuggestionRef.current;
      if (suppressed) {
        if (
          suppressed.lineId === lineId &&
          suppressed.value.trim() === normalized.trim()
        ) {
          queryRef.current = "";
          setQuery("");
          return;
        }
        suppressedSuggestionRef.current = null;
      }

      if (activeCommandQuery) {
        queryRef.current = activeCommandQuery;
        setQuery(activeCommandQuery);
      } else {
        queryRef.current = "";
        setQuery("");
      }
    };

    const commitPendingVisualTransaction = () => {
      if (historyManager.getState().pendingTransaction?.kind === "source-document") {
        return;
      }
      historyManager.commitPendingTransaction();
    };

    const setActiveLine = (lineId: string) => {
      const index = linesRef.current.findIndex((line) => line.id === lineId);
      if (index < 0) return;
      if (activeLineIdRef.current !== lineId) {
        historyManager.commitPendingTransaction();
        focusRequestRef.current += 1;
      }
      activeIndexRef.current = index;
      activeLineIdRef.current = lineId;
      setActiveIndex((current) => (current === index ? current : index));
      const state = useEditorStore.getState();
      if (state.activeLineId !== lineId) state.setActiveLineId(lineId);
    };

    const clearMultiLineSelection = () => {
      for (const lineId of multiLineSelectedIdsRef.current) {
        const field = fieldRefs.current.get(lineId);
        if (!field?.isConnected) continue;
        field.classList.remove("has-visualtex-multi-line-selection");
        field
          .closest<HTMLElement>(".formula-line")
          ?.classList.remove("is-multi-line-selected");
        if (!field.selectionIsCollapsed) {
          const position = Math.max(0, Math.min(field.position, field.lastOffset));
          field.selection = {
            ranges: [[position, position]],
            direction: "none",
          };
        }
      }
      const hadSelection = multiLineSelectedIdsRef.current.size > 0;
      multiLineSelectedIdsRef.current.clear();
      multiLineSelectionRef.current = null;
      documentSelectionRef.current = null;
      documentSelectionNavigationRef.current = null;
      if (hadSelection) {
        setMultiLineSelectionRevision((revision) => revision + 1);
      }
    };

    useLayoutEffect(() => {
      const fields = Array.from(fieldRefs.current.values()).filter(
        (field) => field.isConnected,
      );

      if (!previewOnly) {
        for (const field of fields) {
          field.classList.remove(visualTexSourcePreviewClass);
          field.readOnly = readOnly;
        }
        return;
      }

      commitPendingVisualTransaction();
      queryRef.current = "";
      setQuery("");
      suppressedSuggestionRef.current = null;
      lastSelectionTargetRef.current = null;
      pointerSelectionSessionRef.current = null;
      clearMultiLineSelection();

      for (const field of fields) {
        field.classList.add(visualTexSourcePreviewClass);
        field.readOnly = true;
        const position = Math.max(0, Math.min(field.position, field.lastOffset));
        if (!field.selectionIsCollapsed) {
          field.selection = {
            ranges: [[position, position]],
            direction: "none",
          };
        }
        field.blur();
        delete field.dataset.pendingNativeSuggestion;
        field.classList.remove("has-visualtex-multi-line-selection");
        field
          .closest<HTMLElement>(".formula-line")
          ?.classList.remove("is-multi-line-selected");
        dismissNativeSuggestionPopover(field);
      }

      return () => {
        if (previewOnlyRef.current) return;
        for (const field of fields) {
          if (!field.isConnected) continue;
          field.classList.remove(visualTexSourcePreviewClass);
          field.readOnly = readOnly;
        }
      };
    }, [lines, previewOnly, readOnly]);

    const resolveMultiLineSelectionPoint = (
      clientX: number,
      clientY: number,
      preferredLineId?: string,
    ): MultiLineSelectionPoint | null => {
      const mountedLines = linesRef.current.flatMap((line, lineIndex) => {
        const field = fieldRefs.current.get(line.id);
        const row = field?.closest<HTMLElement>(".formula-line");
        return field?.isConnected && row
          ? [{ lineId: line.id, lineIndex, field, rowRect: row.getBoundingClientRect() }]
          : [];
      });
      if (!mountedLines.length) return null;

      const preferred = preferredLineId
        ? mountedLines.find((line) => line.lineId === preferredLineId)
        : null;
      const target =
        preferred ??
        mountedLines.find(
          ({ rowRect }) => clientY >= rowRect.top && clientY <= rowRect.bottom,
        ) ??
        mountedLines.reduce((nearest, line) => {
          const distance =
            clientY < line.rowRect.top
              ? line.rowRect.top - clientY
              : clientY > line.rowRect.bottom
                ? clientY - line.rowRect.bottom
                : 0;
          const nearestDistance =
            clientY < nearest.rowRect.top
              ? nearest.rowRect.top - clientY
              : clientY > nearest.rowRect.bottom
                ? clientY - nearest.rowRect.bottom
                : 0;
          return distance < nearestDistance ? line : nearest;
        });
      const content = target.field.shadowRoot?.querySelector<HTMLElement>(
        '[part="content"]',
      );
      const contentRect = content?.getBoundingClientRect();
      const sampleY = contentRect
        ? Math.max(contentRect.top + 1, Math.min(clientY, contentRect.bottom - 1))
        : clientY;
      const structuralOffset = structuralBoundaryOffsetFromPoint(
        target.field,
        clientX,
        sampleY,
      );
      const offset = Math.max(
        0,
        Math.min(
          structuralOffset ??
            target.field.getOffsetFromPoint(clientX, sampleY, { bias: 0 }),
          target.field.lastOffset,
        ),
      );
      return {
        lineId: target.lineId,
        lineIndex: target.lineIndex,
        offset,
      };
    };

    const rememberSelectionTarget = (
      lineId: string,
      selection: MathSelectionSnapshot,
    ) => {
      if (documentSelectionNavigationRef.current === lineId &&
          documentSelectionRef.current && multiLineSelectionRef.current) {
        documentSelectionNavigationRef.current = null;
        const field = fieldRefs.current.get(lineId);
        if (field) {
          multiLineSelectionRef.current.focus = {
            lineId, lineIndex: linesRef.current.findIndex(line => line.id === lineId), offset: field.position,
          };
          documentSelectionRef.current = { ...documentSelectionRef.current, [lineId]: selection };
        }
      }
      if (!selectionHasContent(selection)) return;
      lastSelectionTargetRef.current = {
        selections: [
          {
            lineId,
            ranges: selection.ranges,
            direction: selection.direction,
          },
        ],
      };
    };

    const applyMultiLineSelection = (
      anchor: MultiLineSelectionPoint,
      focus: MultiLineSelectionPoint,
    ) => {
      const isForward =
        anchor.lineIndex < focus.lineIndex ||
        (anchor.lineIndex === focus.lineIndex && anchor.offset <= focus.offset);
      const start = isForward ? anchor : focus;
      const end = isForward ? focus : anchor;
      const nextSelectedIds = new Set<string>();

      for (const lineId of multiLineSelectedIdsRef.current) {
        if (
          linesRef.current.findIndex((line) => line.id === lineId) >=
            start.lineIndex &&
          linesRef.current.findIndex((line) => line.id === lineId) <=
            end.lineIndex
        ) {
          continue;
        }
        const field = fieldRefs.current.get(lineId);
        if (!field?.isConnected) continue;
        field.classList.remove("has-visualtex-multi-line-selection");
        field
          .closest<HTMLElement>(".formula-line")
          ?.classList.remove("is-multi-line-selected");
        const position = Math.max(0, Math.min(field.position, field.lastOffset));
        field.selection = {
          ranges: [[position, position]],
          direction: "none",
        };
      }

      for (let lineIndex = start.lineIndex; lineIndex <= end.lineIndex; lineIndex += 1) {
        const line = linesRef.current[lineIndex];
        const field = line ? fieldRefs.current.get(line.id) : null;
        if (!line || !field?.isConnected) continue;
        const rangeStart =
          lineIndex === start.lineIndex ? start.offset : 0;
        const rangeEnd =
          lineIndex === end.lineIndex ? end.offset : field.lastOffset;
        field.selection = {
          ranges: [[Math.min(rangeStart, rangeEnd), Math.max(rangeStart, rangeEnd)]],
          direction: isForward ? "forward" : "backward",
        };
        field.classList.add("has-visualtex-multi-line-selection");
        field
          .closest<HTMLElement>(".formula-line")
          ?.classList.add("is-multi-line-selected");
        nextSelectedIds.add(line.id);
      }

      multiLineSelectedIdsRef.current = nextSelectedIds;
      multiLineSelectionRef.current = { anchor, focus };
      const selections = linesRef.current.flatMap((line) => {
        if (!nextSelectedIds.has(line.id)) return [];
        const field = fieldRefs.current.get(line.id);
        if (!field?.isConnected) return [];
        const selection = captureSelection(field);
        return selectionHasContent(selection)
          ? [
              {
                lineId: line.id,
                ranges: selection.ranges,
                direction: selection.direction,
              } satisfies MathEditorInsertionTarget,
            ]
          : [];
      });
      if (selections.length) lastSelectionTargetRef.current = { selections };
      setMultiLineSelectionRevision((revision) => revision + 1);
    };

    // MathLive owns edits within a formula. The document owns replacing a
    // selection that spans independently mounted formula fields.
    const finishDocumentSelectionEdit = (
      before: ReplaceDocumentEntry["before"],
      lineId: string,
      field: MathfieldElement,
    ) => {
      const line = before.lines.find((item) => item.id === lineId);
      if (!line) return false;
      const selection = captureSelection(field);
      const after: ReplaceDocumentEntry["after"] = {
        ...before,
        lines: [{ ...line, latex: normalizeChineseLatex(field.value) }],
        activeLineId: lineId,
        selectionByLineId: { [lineId]: selection },
      };
      clearMultiLineSelection();
      flushSync(() => useEditorStore.getState().replaceDocumentState(after));
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(lineId);
      field.selection = selection;
      historyManager.push({ type: "replace-document", before, after,
        source: "replace-multi-line", timestamp: Date.now() });
      return true;
    };

    const handleFieldEdit = (
      edit: FormulaFieldEdit,
      field: MathfieldElement,
    ) => {
      if (suppressedHistoryLineIdRef.current === edit.lineId) return;
      if (documentSelectionRef.current && !historyManager.getState().isReplaying) {
        historyManager.commitPendingTransaction();
        const before = getEditorDocumentSnapshot(documentSelectionRef.current);
        if (finishDocumentSelectionEdit(before, edit.lineId, field)) return;
      }
      lastSelectionTargetRef.current = null;
      if (multiLineSelectionRef.current) clearMultiLineSelection();
      const state = useEditorStore.getState();
      const currentLine = state.lines.find((line) => line.id === edit.lineId);
      if (!currentLine) return;
      const beforeActiveLineId = state.activeLineId;
      const beforeLatex = currentLine.latex;

      state.replaceFormulaLine(edit.lineId, edit.afterLatex);
      state.setActiveLineId(edit.lineId);
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(edit.lineId);
      refreshSuggestionQuery(edit.lineId, field, edit.afterLatex);

      if (
        historyManager.getState().isReplaying ||
        suppressedHistoryLineIdRef.current === edit.lineId ||
        beforeLatex === edit.afterLatex
      ) {
        return;
      }

      historyManager.recordFormulaEdit({
        ...edit,
        beforeLatex,
        beforeActiveLineId,
        afterActiveLineId: edit.lineId,
      });
    };

    const applyDiscreteFormulaMutation = (
      lineId: string,
      field: MathfieldElement,
      source: FormulaEditSource,
      mutate: () => boolean,
      replaceDocumentSelection = true,
      recordHistory = true,
    ) => {
      historyManager.commitPendingTransaction();
      const state = useEditorStore.getState();
      const currentLine = state.lines.find((line) => line.id === lineId);
      if (!currentLine) return false;

      const documentBefore = replaceDocumentSelection && documentSelectionRef.current
        ? getEditorDocumentSnapshot(documentSelectionRef.current) : null;

      const before = captureFieldSnapshot(field);
      const beforeActiveLineId = state.activeLineId;
      suppressedHistoryLineIdRef.current = lineId;
      let changed = false;
      try {
        changed = mutate();
      } finally {
        suppressedHistoryLineIdRef.current = null;
      }
      if (!changed) return false;

      normalizeCompletedDifferentialDisplay(field);
      const after = captureFieldSnapshot(field);
      if (documentBefore) {
        field.resetUndo();
        return finishDocumentSelectionEdit(documentBefore, lineId, field);
      }
      if (before.latex === after.latex) {
        field.resetUndo();
        return true;
      }
      state.replaceFormulaLine(lineId, after.latex);
      state.setActiveLineId(lineId);
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(lineId);

      field.resetUndo();

      const entry: ReplaceFormulaEntry = {
        type: "replace-formula",
        lineId,
        beforeLatex: before.latex,
        afterLatex: after.latex,
        beforeSelection: before.selection,
        afterSelection: after.selection,
        beforeActiveLineId,
        afterActiveLineId: lineId,
        timestamp: Date.now(),
        source,
      };
      if (recordHistory) historyManager.push(entry);
      return true;
    };

    const applyInFormulaRowBreak = (
      lineId: string,
      field: MathfieldElement,
      environment: "aligned" | "gathered",
    ) => {
      historyManager.commitPendingTransaction();
      const state = useEditorStore.getState();
      const currentLine = state.lines.find((line) => line.id === lineId);
      if (!currentLine) return false;

      const selectionByLineId = Object.fromEntries(
        linesRef.current.flatMap((line) => {
          const currentField = fieldRefs.current.get(line.id);
          return currentField?.isConnected
            ? [[line.id, captureSelection(currentField)] as const]
            : [];
        }),
      );
      const replacesDocument = Boolean(documentSelectionRef.current);
      const before = getEditorDocumentSnapshot(
        documentSelectionRef.current ?? selectionByLineId,
      );

      suppressedHistoryLineIdRef.current = lineId;
      let changed = false;
      try {
        changed = insertMathLiveRowBreak(field, environment);
      } finally {
        suppressedHistoryLineIdRef.current = null;
      }
      if (!changed) return false;

      normalizeCompletedDifferentialDisplay(field);
      const afterLatex = normalizeChineseLatex(field.value);
      const afterSelection = captureSelection(field);
      field.resetUndo();

      const targetLine =
        before.lines.find((line) => line.id === lineId) ?? currentLine;
      const after: ReplaceDocumentEntry["after"] = {
        ...before,
        lines: replacesDocument
          ? [{ ...targetLine, latex: afterLatex, mode: "display" }]
          : before.lines.map((line) =>
              line.id === lineId
                ? { ...line, latex: afterLatex, mode: "display" }
                : { ...line },
            ),
        activeLineId: lineId,
        selectionByLineId: replacesDocument
          ? { [lineId]: afterSelection }
          : {
              ...before.selectionByLineId,
              [lineId]: afterSelection,
            },
      };

      if (replacesDocument) clearMultiLineSelection();
      flushSync(() => state.replaceDocumentState(after));
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(lineId);
      historyManager.push({
        type: "replace-document",
        before,
        after,
        source: "multiline-row-break",
        timestamp: Date.now(),
      });
      return true;
    };

    const resolveTargetField = () => {
      let targetLineId = activeLineIdRef.current;
      let field = targetLineId
        ? fieldRefs.current.get(targetLineId)
        : undefined;
      if (!field?.isConnected) {
        targetLineId =
          linesRef.current.find((line) =>
            Boolean(fieldRefs.current.get(line.id)?.isConnected),
          )?.id ?? null;
        field = targetLineId
          ? fieldRefs.current.get(targetLineId)
          : undefined;
      }
      return targetLineId && field?.isConnected
        ? { lineId: targetLineId, field }
        : null;
    };

    const insertCommand = (
      command: LatexCommand,
      source: CommandSource = "toolbar",
      activeQuery = "",
    ) => {
      historyManager.commitPendingTransaction();
      const target = resolveTargetField();
      if (!target) return;
      const { lineId: targetLineId, field } = target;
      setActiveLine(targetLineId);
      field.focus();

      const rawAnchor = activeQuery
        ? rawCommandAnchors.get(field) ?? null
        : null;
      const originalSelection = rawAnchor?.selection ?? captureSelection(field);
      let insertionSelection = originalSelection;
      let implicitPreviousLatex = "";
      if (
        !activeQuery &&
        field.selectionIsCollapsed &&
        commandTargetsPreviousExpression(command)
      ) {
        const previousTarget = previousExpressionSelection(field);
        if (previousTarget) {
          insertionSelection = previousTarget.selection;
          implicitPreviousLatex = previousTarget.latex;
        }
      }
      const queryRange = activeQuery
        ? findTrailingCommandRange(field, activeQuery)
        : null;
      const replacesRawCommand = Boolean(
        activeQuery && !queryRange && rawAnchor,
      );
      if (activeQuery && !queryRange && !replacesRawCommand) {
        setQuery("");
        return;
      }

      const selectedLatex = activeQuery
        ? ""
        : implicitPreviousLatex ||
          (insertionSelection.ranges[0]?.[0] ===
          insertionSelection.ranges[0]?.[1]
            ? ""
            : field.getValue(insertionSelection));
      const insertionTemplate = activeQuery
        ? command.insertTemplate
        : templateForSelection(command, selectedLatex);
      const autoExitSetting =
        rawAnchor?.autoExitSetting ?? getCaretAutoExitSetting(field);
      const autoExitScriptKey = rawAnchor?.autoExitScriptKey ?? null;
      const historySource: FormulaEditSource =
        source === "candidate"
          ? "candidate"
          : source === "shortcut"
            ? "shortcut"
            : "toolbar";

      if (
        queryRange &&
        field.getValue(queryRange[0], queryRange[1], "latex").trim() ===
          insertionTemplate.trim()
      ) {
        const normalizedValue = normalizeChineseLatex(field.value);
        suppressedSuggestionRef.current = {
          lineId: targetLineId,
          value: normalizedValue,
        };
        recordCommand(command.id, activeQuery, source);
        setQuery("");
        if (
          source === "candidate" &&
          !structuredSuggestionCommands.has(command.command)
        ) {
          dismissNativeSuggestionPopover(field);
        }
        field.focus();
        return;
      }

      const tryInsert = () => {
        if (rawAnchor) restoreRawCommandInsertionAnchor(field, rawAnchor);
        else field.selection = insertionSelection;
        const inserted = applyDiscreteFormulaMutation(
          targetLineId,
          field,
          historySource,
          () => {
            const isBareOperator = command.id.endsWith("-bare");
            const hasPlaceholder = insertionTemplate.includes("\\placeholder{}");
            if (queryRange) {
              field.selection = {
                ranges: [queryRange],
                direction: "forward",
              };
            }
            const insertCurrentTemplate = () =>
              field.insert(insertionTemplate, {
                mode: "math",
                format: "latex",
                insertionMode: "replaceSelection",
                selectionMode: hasPlaceholder ? "placeholder" : "after",
                style: {
                  variant: "normal",
                  variantStyle: undefined,
                },
                focus: true,
                scrollIntoView: false,
              });
            const valueBeforeInsertion = normalizeChineseLatex(field.value);
            const bareOperatorCountBefore = isBareOperator
              ? valueBeforeInsertion.split(insertionTemplate).length - 1
              : 0;
            const insertedBareOperator = () =>
              normalizeChineseLatex(field.value).split(insertionTemplate).length - 1 >
              bareOperatorCountBefore;
            let inserted = insertCurrentTemplate();
            if (isBareOperator && !insertedBareOperator()) {
              const retryValue = field.value;
              const retrySelection = captureSelection(field);
              field.executeCommand("deleteBackward");
              inserted = insertCurrentTemplate();
              if (!insertedBareOperator()) {
                field.position = field.lastOffset;
                field.selection = {
                  ranges: [[field.lastOffset, field.lastOffset]],
                  direction: "none",
                };
                inserted = insertCurrentTemplate();
              }
              if (!insertedBareOperator()) {
                field.setValue(retryValue, {
                  mode: "math",
                  format: "latex",
                  insertionMode: "replaceAll",
                  selectionMode: "after",
                  silenceNotifications: true,
                });
                field.selection = retrySelection;
              }
            }
            const insertionChangedValue =
              normalizeChineseLatex(field.value) !== valueBeforeInsertion &&
              (!isBareOperator || insertedBareOperator());
            const insertionAccepted =
              inserted && insertionChangedValue;
            if (
              inserted &&
              insertionChangedValue &&
              hasPlaceholder
            ) {
              // MathLive already selects the first editable argument for
              // \nicefrac. Reconstructing that selection from exported
              // offsets can point at its compatibility wrapper instead and
              // make the next toolbar insertion a no-op.
              if (command.id !== "skewed-fraction" && !isBareOperator) {
                selectFirstLatexPlaceholder(
                  field,
                  command.command,
                  insertionTemplate,
                );
              }
            }
            if (
              inserted &&
              insertionChangedValue &&
              !hasPlaceholder &&
              autoExitSetting &&
              inputBehavior[autoExitSetting]
            ) {
              moveCaretThroughEnabledAutoExitContainers(
                field,
                inputBehavior,
                autoExitSetting,
                autoExitScriptKey,
              );
            }
            return insertionAccepted;
          },
        );
        if (inserted) rawCommandAnchors.delete(field);
        else field.selection = originalSelection;
        return inserted;
      };

      const finishInsertion = () => {
        if (
          hasBoundedOperatorPlaceholderOrder(insertionTemplate) &&
          field.selectionIsCollapsed &&
          field.position > 0 &&
          field.getValue(field.position - 1, field.position, "latex").trim() ===
            "\\placeholder{}"
        ) {
          field.selection = {
            ranges: [[field.position - 1, field.position]],
            direction: "none",
          };
        }
        const normalizedValue = normalizeChineseLatex(field.value);
        suppressedSuggestionRef.current = {
          lineId: targetLineId,
          value: normalizedValue,
        };
        recordCommand(command.id, activeQuery, source);
        setQuery("");
        if (
          source === "candidate" &&
          !structuredSuggestionCommands.has(command.command)
        ) {
          dismissNativeSuggestionPopover(field);
        }
        focusLine(targetLineId, {
          latex: normalizedValue,
          selection: captureSelection(field),
          // Placeholder templates already synchronously focus MathLive's keyboard
          // sink after insertion. Replaying that focus/selection 80 ms later can
          // race the user's very first key (notably a physical backslash) before
          // raw-LaTeX mode has had a chance to appear.
          deferredRepair: !insertionTemplate.includes("\\placeholder{}"),
        });
      };

      if (tryInsert()) {
        finishInsertion();
        return;
      }

      window.requestAnimationFrame(() => {
        if (!field.isConnected) return;
        field.focus();
        if (tryInsert()) finishInsertion();
      });
    };

    const setFormulaLineMode = (
      lineId: string,
      mode: FormulaLine["mode"],
    ) => {
      if (
        interactionReadOnly ||
        (mode !== "inline" && mode !== "display")
      ) {
        return;
      }
      const state = useEditorStore.getState();
      const current = state.lines.find((line) => line.id === lineId);
      const currentMode = current?.mode === "inline" ? "inline" : "display";
      if (!current || currentMode === mode) return;

      historyManager.commitPendingTransaction();
      const selectionByLineId = Object.fromEntries(
        linesRef.current.flatMap((line) => {
          const currentField = fieldRefs.current.get(line.id);
          return currentField?.isConnected
            ? [[line.id, captureSelection(currentField)] as const]
            : [];
        }),
      );
      const before = getEditorDocumentSnapshot(selectionByLineId);
      const after: ReplaceDocumentEntry["after"] = {
        ...before,
        lines: before.lines.map((line) =>
          line.id === lineId ? { ...line, mode } : { ...line },
        ),
        activeLineId: lineId,
      };
      flushSync(() => state.replaceDocumentState(after));
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(lineId);
      historyManager.push({
        type: "replace-document",
        before,
        after,
        source: "source-apply",
        timestamp: Date.now(),
      });
      focusLine(lineId, {
        latex: current.latex,
        selection: before.selectionByLineId[lineId] ?? null,
      });
    };

    const addLineAfter = (
      index: number,
      requestedMode?: FormulaLine["mode"],
    ) => {
      historyManager.commitPendingTransaction();
      const state = useEditorStore.getState();
      const beforeActiveLineId = state.activeLineId;
      const beforeField = beforeActiveLineId
        ? fieldRefs.current.get(beforeActiveLineId)
        : undefined;
      const beforeSelection = beforeField
        ? captureSelection(beforeField)
        : null;
      const nextIndex = Math.max(0, Math.min(index + 1, state.lines.length));
      const anchorLineId =
        state.lines[Math.max(0, Math.min(index, state.lines.length - 1))]?.id ??
        beforeActiveLineId;
      const scrollSnapshot = captureEditorScrollSnapshot(anchorLineId);
      const anchorLine =
        state.lines[Math.max(0, Math.min(index, state.lines.length - 1))];
      const inheritedMode =
        requestedMode ??
        anchorLine?.mode ??
        "display";
      const line = createFormulaLine(
        "",
        undefined,
        inheritedMode,
        anchorLine?.displayStyle ?? "default",
      );
      const afterSelection: MathSelectionSnapshot = {
        ranges: [[0, 0]],
        direction: "none",
      };

      const nextLines = state.lines.map((currentLine) => ({ ...currentLine }));
      nextLines.splice(nextIndex, 0, line);
      const nextDocument: ReplaceDocumentEntry["after"] = {
        title: state.title,
        lines: nextLines,
        activeLineId: line.id,
        formulaAlignment: state.formulaAlignment,
        selectionByLineId: {},
      };

      flushSync(() => state.replaceDocumentState(nextDocument));
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(line.id);

      const entry: AddLineEntry = {
        type: "add-line",
        line,
        index: nextIndex,
        beforeActiveLineId,
        afterActiveLineId: line.id,
        beforeSelection,
        afterSelection,
        timestamp: Date.now(),
      };
      historyManager.push(entry);
      setQuery("");
      focusLine(line.id, {
        latex: line.latex,
        selection: afterSelection,
        deferredRepair: false,
      });
      stabilizeEditorScroll(scrollSnapshot, line.id, false);
    };

    const pasteLatexLines = (
      lineId: string,
      field: MathfieldElement,
      pastedLines: string[],
    ) => {
      if (interactionReadOnly || pastedLines.length <= 1) return;
      const replacesDocument = Boolean(documentSelectionRef.current);
      const state = useEditorStore.getState();
      const currentIndex = state.lines.findIndex((line) => line.id === lineId);
      const currentLine = state.lines[currentIndex];
      if (currentIndex < 0 || !currentLine) return;

      const activeRange = field.selection.ranges.at(-1) ?? [
        field.position,
        field.position,
      ];
      const selectionStart = Math.max(
        0,
        Math.min(activeRange[0], activeRange[1], field.lastOffset),
      );
      const selectionEnd = Math.max(
        selectionStart,
        Math.min(Math.max(activeRange[0], activeRange[1]), field.lastOffset),
      );
      const liveLatex = normalizeChineseLatex(field.value);
      const leftLatex = normalizeChineseLatex(
        field.getValue(0, selectionStart, "latex"),
      );
      const rightLatex = normalizeChineseLatex(
        field.getValue(selectionEnd, field.lastOffset, "latex"),
      );
      const canonicalize = (latex: string) => {
        const verifier = new MathfieldElement();
        verifier.setValue(latex, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        return normalizeChineseLatex(verifier.value);
      };
      const concatenate = (left: string, right: string) => {
        const separator =
          /\\[A-Za-z]+$/.test(left) && /^[A-Za-z]/.test(right) ? " " : "";
        return canonicalize(`${left}${separator}${right}`);
      };
      const normalizedLines = pastedLines.map(canonicalize);
      const firstLatex = concatenate(leftLatex, normalizedLines[0] ?? "");
      const lastPastedLatex = normalizedLines.at(-1) ?? "";
      const lastLatex = concatenate(lastPastedLatex, rightLatex);
      const insertedLines = normalizedLines.slice(1).map((latex, index, lines) =>
        createFormulaLine(
          index === lines.length - 1 ? lastLatex : latex,
          undefined,
          currentLine.mode,
          currentLine.displayStyle ?? "default",
        ),
      );
      const lastLine = insertedLines.at(-1);
      if (!lastLine) return;

      historyManager.commitPendingTransaction();
      const selectionByLineId = Object.fromEntries(
        linesRef.current.flatMap((line) => {
          const currentField = fieldRefs.current.get(line.id);
          return currentField?.isConnected
            ? [[line.id, captureSelection(currentField)] as const]
            : [];
        }),
      );
      const before = getEditorDocumentSnapshot(selectionByLineId);
      before.lines = before.lines.map((line) =>
        line.id === lineId ? { ...line, latex: liveLatex } : line,
      );
      const nextLines = (replacesDocument ? before.lines.filter(line => line.id === lineId) : before.lines).map((line) =>
        line.id === lineId ? { ...line, latex: firstLatex } : { ...line },
      );
      nextLines.splice(replacesDocument ? 1 : currentIndex + 1, 0, ...insertedLines);

      const afterSelection: MathSelectionSnapshot = {
        ranges: [[Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER]],
        direction: "none",
      };
      const after: ReplaceDocumentEntry["after"] = {
        title: before.title,
        lines: nextLines,
        activeLineId: lastLine.id,
        formulaAlignment: before.formulaAlignment,
        selectionByLineId: {
          ...before.selectionByLineId,
          [lineId]: {
            ranges: [[Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER]],
            direction: "none",
          },
          [lastLine.id]: afterSelection,
        },
      };

      const scrollSnapshot = captureEditorScrollSnapshot(lineId);
      clearMultiLineSelection();
      lastSelectionTargetRef.current = null;
      flushSync(() => useEditorStore.getState().replaceDocumentState(after));
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(lastLine.id);
      historyManager.push({
        type: "replace-document",
        before,
        after,
        source: "paste-multi-line",
        timestamp: Date.now(),
      });
      setQuery("");
      focusLine(lastLine.id, {
        latex: lastLatex,
        moveToEnd: true,
        deferredRepair: false,
      });
      stabilizeEditorScroll(scrollSnapshot, lastLine.id, false);
    };

    const splitLineAtCaret = (
      index: number,
      lineId: string,
      field: MathfieldElement,
      requestedMode?: FormulaLine["mode"],
    ) => {
      // Raw LaTeX input still belongs to MathLive's command transaction. In
      // ordinary math input, Enter follows Word: a selection is removed first,
      // then the remaining content is split at the selection boundary.
      if (hasRawLatexInput(field)) {
        addLineAfter(index, requestedMode);
        return;
      }

      const selectedRange = field.selection.ranges[0];
      if (field.selection.ranges.length !== 1 || !selectedRange) {
        addLineAfter(index, requestedMode);
        return;
      }
      const splitStart = Math.max(
        0,
        Math.min(
          field.selectionIsCollapsed
            ? field.position
            : Math.min(selectedRange[0], selectedRange[1]),
          field.lastOffset,
        ),
      );
      const splitEnd = Math.max(
        splitStart,
        Math.min(
          field.selectionIsCollapsed
            ? field.position
            : Math.max(selectedRange[0], selectedRange[1]),
          field.lastOffset,
        ),
      );
      if (field.selectionIsCollapsed && splitEnd >= field.lastOffset) {
        addLineAfter(index, requestedMode);
        return;
      }

      const originalLatex = normalizeChineseLatex(field.value);
      const leftLatex = normalizeChineseLatex(
        field.getValue(0, splitStart, "latex"),
      );
      const rightLatex = normalizeChineseLatex(
        field.getValue(splitEnd, field.lastOffset, "latex"),
      );
      const comparable = (value: string) =>
        value
          .replace(/\s+/g, "")
          .replace(/\{([A-Za-z0-9])\}/g, "$1");
      const verifier = new MathfieldElement();
      const canonicalize = (latex: string) => {
        verifier.setValue(latex, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        return comparable(verifier.value);
      };
      // MathLive offsets can point inside a structured atom such as a fraction
      // or matrix. Reparse the two serialized ranges before comparing them:
      // MathLive legitimately normalizes braces and command whitespace at a
      // safe boundary, which made direct string concatenation reject complex
      // trailing structures even though their model was lossless.
      if (
        field.selectionIsCollapsed &&
        (!rightLatex.trim() ||
          canonicalize(`${leftLatex} ${rightLatex}`) !==
            canonicalize(originalLatex))
      ) {
        addLineAfter(index, requestedMode);
        return;
      }

      historyManager.commitPendingTransaction();
      const state = useEditorStore.getState();
      const currentIndex = state.lines.findIndex((line) => line.id === lineId);
      if (currentIndex < 0) {
        addLineAfter(index, requestedMode);
        return;
      }

      const selectionByLineId = Object.fromEntries(
        linesRef.current.flatMap((line) => {
          const currentField = fieldRefs.current.get(line.id);
          return currentField?.isConnected
            ? [[line.id, captureSelection(currentField)] as const]
            : [];
        }),
      );
      const before = getEditorDocumentSnapshot(selectionByLineId);
      before.lines = before.lines.map((line) =>
        line.id === lineId ? { ...line, latex: originalLatex } : line,
      );

      const scrollSnapshot = captureEditorScrollSnapshot(lineId);
      const nextLine = createFormulaLine(
        rightLatex,
        undefined,
        requestedMode ?? before.lines[currentIndex]?.mode ?? "display",
        before.lines[currentIndex]?.displayStyle ?? "default",
      );
      const nextLines = before.lines.map((line) =>
        line.id === lineId ? { ...line, latex: leftLatex } : { ...line },
      );
      nextLines.splice(currentIndex + 1, 0, nextLine);
      const startSelection: MathSelectionSnapshot = {
        ranges: [[0, 0]],
        direction: "none",
      };
      const after: ReplaceDocumentEntry["after"] = {
        title: before.title,
        lines: nextLines,
        activeLineId: nextLine.id,
        formulaAlignment: before.formulaAlignment,
        selectionByLineId: {
          ...before.selectionByLineId,
          [lineId]: {
            ranges: [[Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER]],
            direction: "none",
          },
          [nextLine.id]: startSelection,
        },
      };

      flushSync(() => useEditorStore.getState().replaceDocumentState(after));
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(nextLine.id);

      const entry: ReplaceDocumentEntry = {
        type: "replace-document",
        before,
        after,
        source: "split-line",
        timestamp: Date.now(),
      };
      historyManager.push(entry);
      setQuery("");
      focusLine(nextLine.id, {
        latex: rightLatex,
        selection: startSelection,
        deferredRepair: false,
      });
      stabilizeEditorScroll(scrollSnapshot, nextLine.id, false);
    };

    const mergeLineWithPrevious = (
      index: number,
      lineId: string,
      field: MathfieldElement,
    ) => {
      if (index <= 0) return false;
      const state = useEditorStore.getState();
      const currentIndex = state.lines.findIndex((line) => line.id === lineId);
      if (currentIndex <= 0) return false;
      const currentLine = state.lines[currentIndex];
      const previousLine = state.lines[currentIndex - 1];
      if (!currentLine || !previousLine) return false;

      historyManager.commitPendingTransaction();
      const previousField = fieldRefs.current.get(previousLine.id);
      const previousLatex = normalizeChineseLatex(
        previousField?.value ?? previousLine.latex,
      );
      const currentLatex = normalizeChineseLatex(field.value);
      const joinOffset = previousField?.lastOffset ?? 0;
      const verifier = new MathfieldElement();
      const commandSeparator =
        /\\[A-Za-z]+$/.test(previousLatex) && /^[A-Za-z]/.test(currentLatex)
          ? " "
          : "";
      verifier.setValue(`${previousLatex}${commandSeparator}${currentLatex}`, {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      const mergedLatex = normalizeChineseLatex(verifier.value);

      const selectionByLineId = Object.fromEntries(
        linesRef.current.flatMap((line) => {
          const currentField = fieldRefs.current.get(line.id);
          return currentField?.isConnected
            ? [[line.id, captureSelection(currentField)] as const]
            : [];
        }),
      );
      const before = getEditorDocumentSnapshot(selectionByLineId);
      before.lines = before.lines.map((line) => {
        if (line.id === previousLine.id) {
          return { ...line, latex: previousLatex };
        }
        if (line.id === currentLine.id) {
          return { ...line, latex: currentLatex };
        }
        return line;
      });
      const joinSelection: MathSelectionSnapshot = {
        ranges: [[joinOffset, joinOffset]],
        direction: "none",
      };
      const nextSelectionByLineId = { ...before.selectionByLineId };
      delete nextSelectionByLineId[currentLine.id];
      nextSelectionByLineId[previousLine.id] = joinSelection;
      const after: ReplaceDocumentEntry["after"] = {
        title: before.title,
        lines: before.lines
          .filter((line) => line.id !== currentLine.id)
          .map((line) =>
            line.id === previousLine.id
              ? { ...line, latex: mergedLatex }
              : { ...line },
          ),
        activeLineId: previousLine.id,
        formulaAlignment: before.formulaAlignment,
        selectionByLineId: nextSelectionByLineId,
      };

      const scrollSnapshot = captureEditorScrollSnapshot(previousLine.id);
      prepareFocusBeforeStructuralRemoval(previousLine.id, joinSelection);
      flushSync(() => useEditorStore.getState().replaceDocumentState(after));
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(previousLine.id);
      historyManager.push({
        type: "replace-document",
        before,
        after,
        source: "merge-line",
        timestamp: Date.now(),
      });
      setQuery("");
      finalizeFocusAfterStructuralRemoval(previousLine.id, joinSelection);
      stabilizeEditorScroll(scrollSnapshot, previousLine.id, false);
      return true;
    };

    const deleteMultiLineSelection = () => {
      const selection = multiLineSelectionRef.current;
      if (!selection) return false;
      const anchorIndex = linesRef.current.findIndex(
        (line) => line.id === selection.anchor.lineId,
      );
      const focusIndex = linesRef.current.findIndex(
        (line) => line.id === selection.focus.lineId,
      );
      if (anchorIndex < 0 || focusIndex < 0 || anchorIndex === focusIndex) {
        return false;
      }

      const anchor = { ...selection.anchor, lineIndex: anchorIndex };
      const focus = { ...selection.focus, lineIndex: focusIndex };
      const isForward = anchor.lineIndex < focus.lineIndex;
      const start = isForward ? anchor : focus;
      const end = isForward ? focus : anchor;
      const startLine = linesRef.current[start.lineIndex];
      const endLine = linesRef.current[end.lineIndex];
      const startField = startLine ? fieldRefs.current.get(startLine.id) : null;
      const endField = endLine ? fieldRefs.current.get(endLine.id) : null;
      if (!startLine || !endLine || !startField || !endField) return false;

      historyManager.commitPendingTransaction();
      const startOffset = Math.max(0, Math.min(start.offset, startField.lastOffset));
      const endOffset = Math.max(0, Math.min(end.offset, endField.lastOffset));
      const leftLatex = normalizeChineseLatex(
        startField.getValue(0, startOffset, "latex"),
      );
      const rightLatex = normalizeChineseLatex(
        endField.getValue(endOffset, endField.lastOffset, "latex"),
      );
      const commandSeparator =
        /\\[A-Za-z]+$/.test(leftLatex) && /^[A-Za-z]/.test(rightLatex)
          ? " "
          : "";
      const verifier = new MathfieldElement();
      verifier.setValue(`${leftLatex}${commandSeparator}${rightLatex}`, {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      const mergedLatex = normalizeChineseLatex(verifier.value);
      const prefixVerifier = new MathfieldElement();
      prefixVerifier.setValue(leftLatex, {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      const joinOffset = prefixVerifier.lastOffset;
      const joinSelection: MathSelectionSnapshot = {
        ranges: [[joinOffset, joinOffset]],
        direction: "none",
      };

      const selectionByLineId = Object.fromEntries(
        linesRef.current.flatMap((line) => {
          const currentField = fieldRefs.current.get(line.id);
          return currentField?.isConnected
            ? [[line.id, captureSelection(currentField)] as const]
            : [];
        }),
      );
      const before = getEditorDocumentSnapshot(selectionByLineId);
      before.lines = before.lines.map((line) => {
        const currentField = fieldRefs.current.get(line.id);
        return currentField?.isConnected
          ? { ...line, latex: normalizeChineseLatex(currentField.value) }
          : line;
      });
      const removedIds = new Set(
        before.lines
          .slice(start.lineIndex + 1, end.lineIndex + 1)
          .map((line) => line.id),
      );
      const nextSelectionByLineId = { ...before.selectionByLineId };
      for (const lineId of removedIds) delete nextSelectionByLineId[lineId];
      nextSelectionByLineId[startLine.id] = joinSelection;
      const after: ReplaceDocumentEntry["after"] = {
        title: before.title,
        lines: before.lines
          .filter((line) => !removedIds.has(line.id))
          .map((line) =>
            line.id === startLine.id
              ? { ...line, latex: mergedLatex }
              : { ...line },
          ),
        activeLineId: startLine.id,
        formulaAlignment: before.formulaAlignment,
        selectionByLineId: nextSelectionByLineId,
      };

      const scrollSnapshot = captureEditorScrollSnapshot(startLine.id);
      prepareFocusBeforeStructuralRemoval(startLine.id, joinSelection);
      clearMultiLineSelection();
      flushSync(() => useEditorStore.getState().replaceDocumentState(after));
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(startLine.id);
      historyManager.push({
        type: "replace-document",
        before,
        after,
        source: "delete-multi-line",
        timestamp: Date.now(),
      });
      setQuery("");
      finalizeFocusAfterStructuralRemoval(startLine.id, joinSelection);
      stabilizeEditorScroll(scrollSnapshot, startLine.id, false);
      return true;
    };

    const removeEmptyLine = (index: number) => {
      const state = useEditorStore.getState();
      if (state.lines.length <= 1) return;
      const removedLine = state.lines[index];
      if (!removedLine) return;

      historyManager.commitPendingTransaction();
      const removedField = fieldRefs.current.get(removedLine.id);
      const beforeSelection = removedField
        ? captureSelection(removedField)
        : null;
      const remainingLines = state.lines.filter(
        (line) => line.id !== removedLine.id,
      );
      const targetIndex = Math.max(0, index - 1);
      const targetLine = remainingLines[targetIndex] ?? remainingLines[0];
      if (!targetLine) return;
      const targetField = fieldRefs.current.get(targetLine.id);
      const targetEnd = targetField?.lastOffset ?? targetLine.latex.length;
      const afterSelection: MathSelectionSnapshot = {
        ranges: [[targetEnd, targetEnd]],
        direction: "none",
      };

      const scrollSnapshot = captureEditorScrollSnapshot(targetLine.id);
      const nextDocument: ReplaceDocumentEntry["after"] = {
        title: state.title,
        lines: remainingLines.map((line) => ({ ...line })),
        activeLineId: targetLine.id,
        formulaAlignment: state.formulaAlignment,
        selectionByLineId: {
          [targetLine.id]: afterSelection,
        },
      };
      prepareFocusBeforeStructuralRemoval(targetLine.id, afterSelection);
      flushSync(() => state.replaceDocumentState(nextDocument));
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(targetLine.id);

      const entry: RemoveLineEntry = {
        type: "remove-line",
        line: { ...removedLine },
        index,
        beforeActiveLineId: removedLine.id,
        afterActiveLineId: targetLine.id,
        beforeSelection,
        afterSelection,
        timestamp: Date.now(),
      };
      historyManager.push(entry);
      setQuery("");
      finalizeFocusAfterStructuralRemoval(targetLine.id, afterSelection);
      stabilizeEditorScroll(scrollSnapshot, targetLine.id, false);
    };

    const scheduleUnconsumedMathSpace = (
      lineId: string,
      field: MathfieldElement,
    ) => {
      const beforeNativeSpace = captureFieldSnapshot(field);
      queueMicrotask(() => {
        if (!field.isConnected) return;
        const afterNativeSpace = captureFieldSnapshot(field);
        if (
          beforeNativeSpace.latex !== afterNativeSpace.latex ||
          JSON.stringify(beforeNativeSpace.selection) !==
            JSON.stringify(afterNativeSpace.selection)
        ) {
          return;
        }

        const inserted = applyDiscreteFormulaMutation(
          lineId,
          field,
          "keyboard",
          () =>
            field.insert("\\ ", {
              mode: "math",
              format: "latex",
              insertionMode: "replaceSelection",
              selectionMode: "after",
              focus: true,
              scrollIntoView: false,
            }),
        );
        if (!inserted) return;
        suppressedSuggestionRef.current = null;
        queryRef.current = "";
        setQuery("");
        focusLine(lineId, {
          latex: normalizeChineseLatex(field.value),
          selection: captureSelection(field),
        });
      });
    };

    const handleKeyDown = (
      index: number,
      lineId: string,
      event: KeyboardEvent,
      field: MathfieldElement,
    ) => {
      setActiveLine(lineId);
      documentSelectionNavigationRef.current = null;

      if ((event.ctrlKey || event.metaKey) && !event.altKey &&
          !event.shiftKey && event.key.toLowerCase() === "a") {
        const first = linesRef.current[0];
        const lastIndex = linesRef.current.length - 1;
        const last = linesRef.current[lastIndex];
        const lastField = last && fieldRefs.current.get(last.id);
        if (first && last && lastField) {
          event.preventDefault();
          event.stopImmediatePropagation();
          historyManager.commitPendingTransaction();
          lastField.focus();
          setActiveLine(last.id);
          applyMultiLineSelection(
            { lineId: first.id, lineIndex: 0, offset: 0 },
            { lineId: last.id, lineIndex: lastIndex, offset: lastField.lastOffset },
          );
          if (lastIndex > 0) documentSelectionRef.current = getSelectionMap();
        }
        return;
      }

      if (documentSelectionRef.current && event.shiftKey &&
          ["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown", "Home", "End"].includes(event.key)) {
        // Capture handlers run before MathLive navigation. Its selection-change
        // event is the authoritative endpoint, not a timed callback.
        documentSelectionNavigationRef.current = lineId;
      }

      if (multiLineSelectionRef.current && !event.shiftKey &&
          ["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown", "Home", "End", "Escape"].includes(event.key)) {
        const selectedIds = [...multiLineSelectedIdsRef.current];
        const backwards = ["ArrowLeft", "ArrowUp", "Home"].includes(event.key);
        const targetId = event.key === "Escape" ? lineId : backwards ? selectedIds[0] : selectedIds.at(-1);
        const targetField = targetId ? fieldRefs.current.get(targetId) : null;
        if (targetId && targetField) {
          const ranges = targetField.selection.ranges.flat();
          const offset = backwards ? Math.min(...ranges) : Math.max(...ranges);
          clearMultiLineSelection();
          event.preventDefault();
          event.stopImmediatePropagation();
          focusLine(targetId, { selection: { ranges: [[offset, offset]], direction: "none" } });
          return;
        }
      }

      if (
        (event.key === "Backspace" || event.key === "Delete") &&
        !event.altKey &&
        !event.ctrlKey &&
        !event.metaKey &&
        !event.shiftKey &&
        deleteMultiLineSelection()
      ) {
        event.preventDefault();
        event.stopImmediatePropagation();
        return;
      }

      if (
        greekLetterHotkeyLineIdRef.current &&
        greekLetterHotkeyLineIdRef.current !== lineId
      ) {
        greekLetterHotkeyLineIdRef.current = null;
      }

      const rawLatexInputActive = hasRawLatexInput(field);
      if (!rawLatexInputActive && isGreekLetterHotkeyPrefix(event)) {
        event.preventDefault();
        event.stopImmediatePropagation();
        greekLetterHotkeyLineIdRef.current = lineId;
        delete field.dataset.pendingNativeSuggestion;
        suppressedSuggestionRef.current = null;
        queryRef.current = "";
        setQuery("");
        dismissNativeSuggestionPopover(field);
        return;
      }

      if (greekLetterHotkeyLineIdRef.current === lineId) {
        if (event.code === "ShiftLeft" || event.code === "ShiftRight") {
          return;
        }
        if (event.key === "Escape") {
          event.preventDefault();
          event.stopImmediatePropagation();
          greekLetterHotkeyLineIdRef.current = null;
          return;
        }
        const greekCommand = greekLetterHotkeyCommandFromEvent(event);
        greekLetterHotkeyLineIdRef.current = null;
        if (greekCommand) {
          event.preventDefault();
          event.stopImmediatePropagation();
          delete field.dataset.pendingNativeSuggestion;
          suppressedSuggestionRef.current = null;
          queryRef.current = "";
          setQuery("");
          insertCommand(greekCommand, "shortcut");
          return;
        }
      }

      const forcedNewLineMode =
        showLineModeControls &&
        event.key === "Enter" &&
        !event.isComposing &&
        !event.metaKey &&
        !event.shiftKey &&
        !rawLatexInputActive
          ? event.ctrlKey && !event.altKey
            ? "inline"
            : event.altKey && !event.ctrlKey
              ? "display"
              : null
          : null;
      if (forcedNewLineMode) {
        event.preventDefault();
        event.stopImmediatePropagation();
        splitLineAtCaret(index, lineId, field, forcedNewLineMode);
        return;
      }

      const formulaHotkey = rawLatexInputActive
        ? null
        : matchFormulaHotkey(event, formulaHotkeyBindings);
      if (formulaHotkey) {
        event.preventDefault();
        event.stopImmediatePropagation();
        delete field.dataset.pendingNativeSuggestion;
        suppressedSuggestionRef.current = null;
        queryRef.current = "";
        setQuery("");
        insertCommand(
          resolveFormulaHotkeyCommand(formulaHotkey.target),
          "shortcut",
        );
        return;
      }

      const shortcutKey = event.key.toLocaleLowerCase();
      const primaryModifier = (event.metaKey || event.ctrlKey) && !event.altKey;
      const requestsUndo = primaryModifier && shortcutKey === "z" && !event.shiftKey;
      const requestsRedo =
        primaryModifier &&
        ((shortcutKey === "z" && event.shiftKey) ||
          (shortcutKey === "y" && !event.shiftKey));
      if (requestsUndo || requestsRedo) {
        clearMultiLineSelection();
        event.preventDefault();
        event.stopPropagation();
        if (requestsRedo) historyManager.requestRedo();
        else historyManager.requestUndo();
        return;
      }

      const requestsInFormulaRowBreak =
        event.key === "Enter" &&
        event.shiftKey &&
        !event.altKey &&
        !event.ctrlKey &&
        !event.metaKey &&
        !event.isComposing &&
        !rawLatexInputActive;
      if (requestsInFormulaRowBreak) {
        event.preventDefault();
        event.stopImmediatePropagation();

        const inserted = applyInFormulaRowBreak(
          lineId,
          field,
          latexFormatProfile.multilineEnvironment === "align"
            ? "aligned"
            : "gathered",
        );
        if (inserted) {
          suppressedSuggestionRef.current = null;
          queryRef.current = "";
          setQuery("");
          focusLine(lineId, {
            latex: normalizeChineseLatex(field.value),
            selection: captureSelection(field),
          });
        }
        return;
      }

      if (
        event.key !== "ArrowDown" &&
        event.key !== "ArrowUp" &&
        event.key !== "Enter" &&
        event.key !== "Tab" &&
        event.key !== " " &&
        event.code !== "Space"
      ) {
        delete field.dataset.pendingNativeSuggestion;
      }

      const rawCommandActive = hasRawLatexInput(field);
      // Let native completion own keys while a command is actually pending.
      // Escape can leave MathLive in latex mode with no raw input; that state
      // must still allow formula-row navigation.
      if (rawCommandActive || (field.mode === "latex" &&
          event.key !== "ArrowUp" && event.key !== "ArrowDown")) return;
      const currentState = useEditorStore.getState();

      const activeEnvironment = activeMathLiveEnvironmentName(field);
      const requestsAlignmentPoint =
        event.key === "&" &&
        !event.isComposing &&
        !event.altKey &&
        !event.ctrlKey &&
        !event.metaKey &&
        !rawCommandActive;
      const insertsInternalAlignmentPoint =
        requestsAlignmentPoint &&
        (activeEnvironment === "aligned" ||
          activeEnvironment === "align" ||
          activeEnvironment === "align*");
      if (insertsInternalAlignmentPoint) {
        const inserted = applyDiscreteFormulaMutation(
          lineId,
          field,
          "keyboard",
          () => insertMathLiveAlignmentPoint(field),
        );
        if (inserted) {
          event.preventDefault();
          event.stopImmediatePropagation();
          suppressedSuggestionRef.current = null;
          queryRef.current = "";
          setQuery("");
          focusLine(lineId, {
            latex: normalizeChineseLatex(field.value),
            selection: captureSelection(field),
          });
          return;
        }
      }

      const alignmentCaretDepth =
        field.getElementInfo(field.position)?.depth ??
        field.getElementInfo(Math.max(0, field.position - 1))?.depth ??
        0;
      const insertsExplicitAlignmentPoint =
        latexFormatProfile.multilineEnvironment === "align" &&
        alignmentCaretDepth === 0 &&
        requestsAlignmentPoint &&
        activeEnvironment === null;
      if (insertsExplicitAlignmentPoint) {
        event.preventDefault();
        event.stopImmediatePropagation();
        const inserted = applyDiscreteFormulaMutation(
          lineId,
          field,
          "keyboard",
          () =>
            field.insert(VISUALTEX_ALIGNMENT_MARKER_LATEX, {
              mode: "math",
              format: "latex",
              insertionMode: "replaceSelection",
              selectionMode: "after",
              focus: true,
              scrollIntoView: false,
            }),
        );
        if (inserted) {
          suppressedSuggestionRef.current = null;
          queryRef.current = "";
          setQuery("");
          focusLine(lineId, {
            latex: normalizeChineseLatex(field.value),
            selection: captureSelection(field),
          });
        }
        return;
      }

      const insertsVisibleMathSpace =
        !event.isComposing &&
        !event.altKey &&
        !event.ctrlKey &&
        !event.metaKey &&
        !event.shiftKey &&
        (event.key === " " || event.code === "Space") &&
        !rawCommandActive;
      if (insertsVisibleMathSpace) {
        scheduleUnconsumedMathSpace(lineId, field);
        return;
      }

      if (
        event.key === "Enter" &&
        activeMathLiveEnvironmentName(field) === "cases"
      ) {
        event.preventDefault();
        event.stopImmediatePropagation();
        const insertedRow = applyDiscreteFormulaMutation(
          lineId,
          field,
          "keyboard",
          () => field.executeCommand("addRowAfter"),
        );
        if (insertedRow) {
          suppressedSuggestionRef.current = null;
          queryRef.current = "";
          setQuery("");
          focusLine(lineId, {
            latex: normalizeChineseLatex(field.value),
            selection: captureSelection(field),
          });
        }
        return;
      }

      if (event.key === "Enter") {
        event.preventDefault();
        event.stopImmediatePropagation();
        const currentLineMode =
          currentState.lines.find((line) => line.id === lineId)?.mode ??
          "display";
        splitLineAtCaret(index, lineId, field, currentLineMode);
        return;
      }

      if (event.key === "Backspace" || event.key === "Delete") {
        const visibleLatex = field
          .getValue("latex-without-placeholders")
          .trim();
        const rawCommandInput = rawLatexInput(field);
        const state = useEditorStore.getState();
        const currentLine = state.lines.find((line) => line.id === lineId);
        const trulyEmpty = Boolean(
          currentLine &&
            visibleLatex.length === 0 &&
            rawCommandInput.length === 0 &&
            normalizeChineseLatex(field.value).trim().length === 0 &&
            currentLine.latex.trim().length === 0,
        );

        // Backspace on a truly empty row restores the multiline editor's
        // previous-row behavior. This check must happen before the raw-LaTeX
        // mode guard because MathLive can remain in `latex` mode after the
        // final raw command character has already been deleted.
        if (
          event.key === "Backspace" &&
          trulyEmpty &&
          field.selectionIsCollapsed &&
          !event.altKey &&
          !event.ctrlKey &&
          !event.metaKey &&
          !event.shiftKey
        ) {
          event.preventDefault();
          event.stopImmediatePropagation();
          removeEmptyLine(index);
          return;
        }

        // While MathLive is collecting a non-empty raw LaTeX command such as
        // `\\mat`, its public formula value is intentionally still empty. Let
        // MathLive consume Backspace/Delete natively so one physical key
        // removes one raw command character instead of deleting the row.
        if (rawCommandInput) return;

        // Preserve MathLive's native word/line deletion shortcuts such as
        // Option+Backspace and Command+Backspace on macOS.
        if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) {
          return;
        }

        if (
          event.key === "Backspace" &&
          field.selectionIsCollapsed &&
          field.position === 0
        ) {
          event.preventDefault();
          event.stopImmediatePropagation();
          mergeLineWithPrevious(index, lineId, field);
          return;
        }

        event.preventDefault();
        event.stopImmediatePropagation();

        if (!currentLine) return;
        const before = captureFieldSnapshot(field);
        const beforePosition = field.position;
        const command =
          event.key === "Backspace" ? "deleteBackward" : "deleteForward";

        const trailingStructuredLatex =
          event.key === "Backspace"
            ? field.getElementInfo(beforePosition)?.latex?.trim() ?? ""
            : "";
        suppressedHistoryLineIdRef.current = lineId;
        try {
          field.executeCommand(command);
          if (
            event.key === "Backspace" &&
            field.value === before.latex &&
            scriptContainerPattern.test(trailingStructuredLatex)
          ) {
            // MathLive first moves into a terminal super/subscript container
            // without deleting anything. Complete the same Backspace action by
            // deleting its final atom so the editor cannot become stuck there.
            field.executeCommand("deleteBackward");
          }
        } finally {
          suppressedHistoryLineIdRef.current = null;
        }

        if (event.key === "Backspace" && field.selectionIsCollapsed) {
          keepCaretAfterBareStructuredOperator(field, beforePosition);
        }

        const after = captureFieldSnapshot(field);
        if (before.latex === after.latex) {
          field.focus();
          return;
        }

        state.replaceFormulaLine(lineId, after.latex);
        state.setActiveLineId(lineId);
        linesRef.current = useEditorStore.getState().lines;
        field.resetUndo();
        historyManager.recordFormulaEdit({
          lineId,
          beforeLatex: currentLine.latex,
          afterLatex: after.latex,
          beforeSelection: before.selection,
          afterSelection: after.selection,
          beforeActiveLineId: state.activeLineId,
          afterActiveLineId: lineId,
          editKind:
            event.key === "Backspace"
              ? "delete-backward"
              : "delete-forward",
          source: "keyboard",
        });
        return;
      }

      if (
        (event.key === "ArrowUp" || event.key === "ArrowDown") &&
        !event.altKey &&
        !event.ctrlKey &&
        !event.metaKey &&
        !event.shiftKey
      ) {
        // Ask the kernel for every formula, regardless of its LaTeX command
        // spelling. A changed selection means an actual inner branch move;
        // otherwise the arrow may move to a different formula row below.
        const beforeSelection = captureSelection(field);
        field.executeCommand(event.key === "ArrowUp" ? "moveUp" : "moveDown");
        const afterSelection = captureSelection(field);
        if (JSON.stringify(beforeSelection) !== JSON.stringify(afterSelection)) {
          event.preventDefault();
          event.stopImmediatePropagation();
          return;
        }

        const direction = event.key === "ArrowUp" ? -1 : 1;
        const targetIndex = index + direction;
        const targetLine = linesRef.current[targetIndex];
        if (targetLine) {
          event.preventDefault();
          event.stopImmediatePropagation();
          historyManager.commitPendingTransaction();
          queryRef.current = "";
          setQuery("");

          const targetField = fieldRefs.current.get(targetLine.id);
          const targetPosition = Math.max(
            0,
            Math.min(
              field.position,
              targetField?.lastOffset ?? targetLine.latex.length,
            ),
          );
          const targetSelection: MathSelectionSnapshot = {
            ranges: [[targetPosition, targetPosition]],
            direction: "none",
          };
          setActiveLine(targetLine.id);
          if (targetField?.isConnected) {
            const applyTargetFocus = () => {
              setActiveLine(targetLine.id);
              targetField.focus();
              targetField.shadowRoot
                ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
                ?.focus({ preventScroll: true });
              targetField.selection = targetSelection;
              targetField.position = targetPosition;
            };
            applyTargetFocus();
            window.requestAnimationFrame(applyTargetFocus);
            window.setTimeout(applyTargetFocus, 0);
            window.setTimeout(applyTargetFocus, 80);
          } else {
            focusLine(targetLine.id, {
              latex: targetLine.latex,
              selection: targetSelection,
            });
          }
          return;
        }
      }

      if (
        event.key.startsWith("Arrow") ||
        event.key === "Home" ||
        event.key === "End"
      ) {
        historyManager.commitPendingTransaction();
      }
    };

    const normalizeInsertedFormulaLines = (latex: string) => {
      const normalized = latex.replace(/\r\n?/g, "\n");
      const trimmed = normalized.trim();
      if (isSingleCompleteLatexEnvironment(trimmed)) {
        return [normalizeChineseLatex(trimmed)];
      }
      return normalized
        .split("\n")
        .map((line) => normalizeChineseLatex(line.trim()))
        .filter(Boolean);
    };

    const normalizeInsertedLatex = (latex: string) =>
      normalizeInsertedFormulaLines(latex).join("\\quad ");

    const getSelectionMap = (): Record<string, MathSelectionSnapshot> =>
      Object.fromEntries(
        linesRef.current.flatMap((line) => {
          const field = fieldRefs.current.get(line.id);
          return field?.isConnected
            ? [[line.id, captureSelection(field)] as const]
            : [];
        }),
      );

    const captureInsertionTarget = (): MathEditorInsertionTarget | null => {
      const target = resolveTargetField();
      if (!target) return null;
      const selection = captureSelection(target.field);
      return {
        lineId: target.lineId,
        ranges: selection.ranges,
        direction: selection.direction,
      };
    };

    const captureSelectionTarget = (): MathEditorSelectionTarget | null => {
      // MathLive fields can retain a visual/model selection after focus moves to
      // another formula. A formatting button must never sweep those stale
      // selections from unrelated rows. Only an explicit VisualTeX multi-line
      // drag is allowed to target more than the active formula row.
      const explicitMultiLineSelection = Boolean(multiLineSelectionRef.current);
      const activeLineId =
        activeLineIdRef.current ?? useEditorStore.getState().activeLineId;
      const candidateLines = explicitMultiLineSelection
        ? linesRef.current.filter((line) =>
            multiLineSelectedIdsRef.current.has(line.id),
          )
        : linesRef.current.filter((line) => line.id === activeLineId);
      const selections = candidateLines.flatMap((line) => {
        const field = fieldRefs.current.get(line.id);
        if (!field?.isConnected) return [];
        const selection = captureSelection(field);
        if (!selectionHasContent(selection)) return [];
        return [
          {
            lineId: line.id,
            ranges: selection.ranges,
            direction: selection.direction,
          } satisfies MathEditorInsertionTarget,
        ];
      });
      if (selections.length) {
        const target = { selections };
        lastSelectionTargetRef.current = target;
        return target;
      }
      lastSelectionTargetRef.current = null;
      return null;
    };

    const applySelectionStyle = (
      style: MathEditorSelectionStyle,
      target: MathEditorSelectionTarget | null = captureSelectionTarget(),
    ) => {
      if (interactionReadOnly || !target?.selections.length) return false;
      const clearsBackground =
        style.kind === "backgroundColor" && style.value === "none";
      if (
        (style.kind === "color" || style.kind === "backgroundColor") &&
        !clearsBackground &&
        !isSafeFormulaStyleColor(style.value)
      ) {
        return false;
      }
      const resolved = target.selections.flatMap((selectionTarget) => {
        const field = fieldRefs.current.get(selectionTarget.lineId);
        if (!field?.isConnected) return [];
        const selection = clampSelection(
          {
            ranges: selectionTarget.ranges,
            direction: selectionTarget.direction,
          },
          field.lastOffset,
        );
        if (!selectionHasContent(selection)) return [];
        field.selection = selection;
        return [{ lineId: selectionTarget.lineId, field, selection }];
      });
      if (!resolved.length) return false;

      historyManager.commitPendingTransaction();
      const documentBefore = resolved.length > 1
        ? getEditorDocumentSnapshot(getSelectionMap()) : null;
      for (const entry of resolved) {
        const { field, lineId } = entry;
        field.selection = entry.selection;
        if (style.kind === "bold" || style.kind === "italic") {
          const toggled = applyDiscreteFormulaMutation(
            lineId,
            field,
            "toolbar",
            () => {
              return toggleMathLiveSelectionStyle(field, style.kind, entry.selection);
            },
            false,
            !documentBefore,
          );
          if (!toggled) field.selection = entry.selection;
          entry.selection = captureSelection(field);
          continue;
        }

        applyDiscreteFormulaMutation(lineId, field, "toolbar", () => {
          let mathLiveStyle: Pick<
            Style,
            "variant" | "variantStyle" | "color" | "backgroundColor"
          >;
          if (style.kind === "color") {
            mathLiveStyle = { color: style.value };
          } else {
            mathLiveStyle = { backgroundColor: style.value };
          }
          for (const [start, end] of entry.selection.ranges) {
            const range: [number, number] = [
              Math.min(start, end),
              Math.max(start, end),
            ];
            if (range[0] === range[1]) continue;
            field.applyStyle(mathLiveStyle, {
              range,
              operation: "set",
            });
          }
          // MathLive normally derives the implicit style for future input from
          // the atom immediately before the caret. If the formatted selection
          // includes the last atom, typing at the end would otherwise inherit
          // its color/background even though these controls are selection-only.
          const resetPosition = Math.max(
            ...entry.selection.ranges.flatMap(([start, end]) => [start, end]),
          );
          field.selection = {
            ranges: [[resetPosition, resetPosition]],
            direction: "none",
          };
          field.applyStyle(
            style.kind === "color"
              ? { color: "none" }
              : { backgroundColor: "none" },
            { operation: "set" },
          );
          field.selection = entry.selection;
          return true;
        }, false, !documentBefore);
      }
      historyManager.commitPendingTransaction();

      if (documentBefore) {
        historyManager.push({ type: "replace-document", before: documentBefore,
          after: getEditorDocumentSnapshot(getSelectionMap()), source: "format-multi-line", timestamp: Date.now() });
      }

      const activeSelection =
        resolved.find(({ lineId }) => lineId === activeLineIdRef.current) ??
        resolved.at(-1);
      if (activeSelection) {
        setActiveLine(activeSelection.lineId);
        activeSelection.field.focus();
        activeSelection.field.selection = activeSelection.selection;
        activeSelection.field.shadowRoot
          ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
          ?.focus({ preventScroll: true });
      }
      lastSelectionTargetRef.current = null;
      return true;
    };

    const restoreSelection = (
      lineId: string,
      latex: string,
      selection: MathSelectionSnapshot | null,
    ): Promise<boolean> =>
      new Promise((resolve) => {
        const index = linesRef.current.findIndex((line) => line.id === lineId);
        if (index < 0) {
          resolve(false);
          return;
        }
        historyManager.commitPendingTransaction();
        const requestId = ++focusRequestRef.current;
        activeIndexRef.current = index;
        activeLineIdRef.current = lineId;
        setActiveIndex(index);
        useEditorStore.getState().setActiveLineId(lineId);
        pendingFocusRef.current = {
          lineId,
          latex,
          selection,
          moveToEnd: !selection,
          deferredRepair: true,
        };

        let attempts = 12;
        const attempt = () => {
          const applied = applyFocusState(
            lineId,
            latex,
            selection,
            !selection,
            requestId,
          );
          if (applied) {
            pendingFocusRef.current = null;
            const restoredField = fieldRefs.current.get(lineId);
            const restoredSelection =
              restoredField?.isConnected ? captureSelection(restoredField) : null;
            const historySnapshotStillCurrent = () => {
              const currentField = fieldRefs.current.get(lineId);
              if (!currentField?.isConnected) return false;
              if (normalizeChineseLatex(currentField.value) !== latex) return false;
              if (!restoredSelection) return true;
              const currentSelection = captureSelection(currentField);
              return (
                currentSelection.direction === restoredSelection.direction &&
                JSON.stringify(currentSelection.ranges) ===
                  JSON.stringify(restoredSelection.ranges)
              );
            };
            window.requestAnimationFrame(() => {
              // History focus repair is only valid while the restored snapshot
              // is still untouched. A user can type immediately after Undo/Redo;
              // replaying the old LaTeX 1 frame or 80 ms later would otherwise
              // erase that first new character and make history feel broken.
              if (!historySnapshotStillCurrent()) return;
              if (!applyFocusState(lineId, latex, selection, !selection, requestId)) {
                return;
              }
              window.setTimeout(() => {
                if (!historySnapshotStillCurrent()) return;
                applyFocusState(
                  lineId,
                  latex,
                  selection,
                  !selection,
                  requestId,
                );
              }, 80);
            });
            resolve(true);
            return;
          }
          attempts -= 1;
          if (attempts <= 0) {
            resolve(false);
            return;
          }
          window.requestAnimationFrame(() => window.setTimeout(attempt, 16));
        };
        attempt();
      });

    const insertMultipleLatexRowsAt = (
      target: MathEditorInsertionTarget,
      values: readonly string[],
      source: FormulaEditSource,
    ): boolean => {
      if (values.length < 2) return false;
      const targetIndex = linesRef.current.findIndex(
        (line) => line.id === target.lineId,
      );
      if (targetIndex < 0) return false;
      const field = fieldRefs.current.get(target.lineId);
      if (!field?.isConnected) return false;

      const selection = clampSelection(
        {
          ranges: target.ranges.length
            ? target.ranges
            : [[field.lastOffset, field.lastOffset]],
          direction: target.direction,
        },
        field.lastOffset,
      );
      historyManager.commitPendingTransaction();
      const replacesDocument = Boolean(documentSelectionRef.current);
      const before = getEditorDocumentSnapshot(documentSelectionRef.current ?? getSelectionMap());
      if (replacesDocument) clearMultiLineSelection();
      setActiveLine(target.lineId);
      field.focus();

      suppressedHistoryLineIdRef.current = target.lineId;
      let inserted = false;
      try {
        field.selection = selection;
        inserted = Boolean(
          field.insert(values[0], {
            mode: "math",
            format: "latex",
            insertionMode: "replaceSelection",
            selectionMode: "after",
            focus: true,
            scrollIntoView: false,
          }),
        );
        if (!inserted) return false;
        const normalized = normalizeChineseLatex(field.value);
        if (normalized !== field.value) {
          field.setValue(normalized, { silenceNotifications: true });
        }
      } finally {
        suppressedHistoryLineIdRef.current = null;
      }
      if (!inserted) return false;

      const firstLatex = normalizeChineseLatex(field.value);
      const firstSelection = captureSelection(field);
      const currentLines = useEditorStore.getState().lines;
      const currentTargetIndex = currentLines.findIndex(
        (line) => line.id === target.lineId,
      );
      if (currentTargetIndex < 0) return false;
      const additionalLines = values.slice(1).map((value) => createFormulaLine(value));
      const nextLines = (replacesDocument ? currentLines.filter(line => line.id === target.lineId) : currentLines).map((line) =>
        line.id === target.lineId ? { ...line, latex: firstLatex } : { ...line },
      );
      nextLines.splice(replacesDocument ? 1 : currentTargetIndex + 1, 0, ...additionalLines);
      const lastLine = additionalLines[additionalLines.length - 1];
      const lastSelection: MathSelectionSnapshot = {
        ranges: [[Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER]],
        direction: "none",
      };
      const after: ReplaceDocumentEntry["after"] = {
        title: before.title,
        lines: nextLines,
        activeLineId: lastLine.id,
        formulaAlignment: before.formulaAlignment,
        selectionByLineId: {
          ...before.selectionByLineId,
          [target.lineId]: firstSelection,
          [lastLine.id]: lastSelection,
        },
      };

      flushSync(() => useEditorStore.getState().replaceDocumentState(after));
      linesRef.current = useEditorStore.getState().lines;
      field.resetUndo();
      historyManager.push({
        type: "replace-document",
        before,
        after,
        source: source === "paste" ? "paste-multi-line" : "ocr",
        timestamp: Date.now(),
      });
      setQuery("");
      focusLine(lastLine.id, {
        latex: lastLine.latex,
        moveToEnd: true,
      });
      return true;
    };

    const prepareOcrAlignmentFormat = (
      values: readonly string[],
      source: FormulaEditSource,
    ) => {
      if (
        source === "ocr" &&
        values.some((value) => hasVisualTexAlignmentMarker(value))
      ) {
        useEditorStore.getState().setLatexCodeFormat("aligned");
      }
    };

    const insertLatex = (
      latex: string,
      source: FormulaEditSource = "ocr",
    ) => {
      const values = normalizeInsertedFormulaLines(latex);
      if (!values.length) return;
      prepareOcrAlignmentFormat(values, source);
      const target = resolveTargetField();
      if (!target) return;
      if (values.length > 1) {
        const selection = captureSelection(target.field);
        insertMultipleLatexRowsAt(
          {
            lineId: target.lineId,
            ranges: selection.ranges,
            direction: selection.direction,
          },
          values,
          source,
        );
        return;
      }

      const value = values[0];
      const { lineId, field } = target;
      setActiveLine(lineId);
      field.focus();
      const inserted = applyDiscreteFormulaMutation(
        lineId,
        field,
        source,
        () =>
          field.insert(value, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceSelection",
            selectionMode: "after",
            focus: true,
            scrollIntoView: false,
          }),
      );
      if (!inserted) return;
      setQuery("");
      focusLine(lineId, {
        latex: normalizeChineseLatex(field.value),
        selection: captureSelection(field),
      });
    };

    const insertLatexAt = (
      target: MathEditorInsertionTarget,
      latex: string,
      source: FormulaEditSource = "ocr",
    ): boolean => {
      const values = normalizeInsertedFormulaLines(latex);
      if (!values.length) return false;
      prepareOcrAlignmentFormat(values, source);
      if (values.length > 1) {
        return insertMultipleLatexRowsAt(target, values, source);
      }
      const value = values[0];
      if (!linesRef.current.some((line) => line.id === target.lineId)) {
        return false;
      }
      const field = fieldRefs.current.get(target.lineId);
      if (!field?.isConnected) return false;

      const selection = clampSelection(
        {
          ranges: target.ranges.length
            ? target.ranges
            : [[field.lastOffset, field.lastOffset]],
          direction: target.direction,
        },
        field.lastOffset,
      );
      setActiveLine(target.lineId);
      field.focus();
      const inserted = applyDiscreteFormulaMutation(
        target.lineId,
        field,
        source,
        () => {
          field.selection = selection;
          return field.insert(value, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceSelection",
            selectionMode: "after",
            focus: true,
            scrollIntoView: false,
          });
        },
      );
      if (!inserted) return false;
      setQuery("");
      focusLine(target.lineId, {
        latex: normalizeChineseLatex(field.value),
        selection: captureSelection(field),
      });
      return true;
    };

    const appendLatex = (
      latex: string,
      source: FormulaEditSource = "ocr",
    ) => {
      const values = latex
        .replace(/\r\n?/g, "\n")
        .split("\n")
        .map((line) => normalizeChineseLatex(line.trim()))
        .filter(Boolean);
      if (!values.length) return;
      prepareOcrAlignmentFormat(values, source);

      historyManager.commitPendingTransaction();
      const before = getEditorDocumentSnapshot(getSelectionMap());
      const currentLines = useEditorStore.getState().lines;
      const replacesOnlyBlankLine =
        currentLines.length === 1 && !currentLines[0].latex.trim();
      const nextLines = replacesOnlyBlankLine
        ? values.map((value, index) =>
            index === 0
              ? { ...currentLines[0], latex: value }
              : createFormulaLine(value),
          )
        : [...currentLines, ...values.map((value) => createFormulaLine(value))];
      const lastLine = nextLines[nextLines.length - 1];
      const afterSelection: MathSelectionSnapshot = {
        ranges: [[Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER]],
        direction: "none",
      };
      const after: ReplaceDocumentEntry["after"] = {
        title: before.title,
        lines: nextLines.map((line) => ({ ...line })),
        activeLineId: lastLine.id,
        formulaAlignment: before.formulaAlignment,
        selectionByLineId: {
          ...before.selectionByLineId,
          [lastLine.id]: afterSelection,
        },
      };

      flushSync(() => useEditorStore.getState().replaceDocumentState(after));
      linesRef.current = useEditorStore.getState().lines;
      const entry: ReplaceDocumentEntry = {
        type: "replace-document",
        before,
        after,
        source: "ocr",
        timestamp: Date.now(),
      };
      historyManager.push(entry);
      setQuery("");
      focusLine(lastLine.id, {
        latex: lastLine.latex,
        moveToEnd: true,
      });
    };

    useImperativeHandle(ref, () => ({
      insertCommand,
      insertLatex,
      insertLatexAt,
      appendLatex,
      focus: (options = {}) => {
        const target = options.target ?? "active";
        const lineId =
          target === "first"
            ? linesRef.current[0]?.id
            : target === "last"
              ? linesRef.current.at(-1)?.id
              : activeLineIdRef.current ?? linesRef.current[0]?.id;
        if (lineId) {
          focusLine(lineId, {
            moveToEnd: options.moveToEnd ?? false,
          });
        }
      },
      addLine: () => addLineAfter(linesRef.current.length - 1),
      commitPendingTransaction: () => historyManager.commitPendingTransaction(),
      getSelectionMap,
      restoreSelection,
      captureSelectionTarget,
      captureInsertionTarget,
      applySelectionStyle,
      refreshLayout: () => {
        flushSync(() => {
          setFieldRenderEpoch((epoch) => epoch + 1);
        });
        window.dispatchEvent(new Event(EDITOR_LAYOUT_REFRESH_EVENT));
      },
    }));

    useEffect(() => {
      queryRef.current = query;
    }, [query]);

    useEffect(() => {
      const surface = surfaceRef.current;
      if (!surface) return;

      const handleSurfacePointerDown = (event: PointerEvent) => {
        if (event.button !== 0 || !event.isPrimary || event.shiftKey) return;
        lastSelectionTargetRef.current = null;
        const path = event.composedPath();
        const entry = Array.from(fieldRefs.current.entries()).find(
          ([, field]) =>
            path.includes(field) ||
            Boolean(field.parentElement && path.includes(field.parentElement)),
        );
        if (!entry) return;

        if (previewOnlyRef.current) {
          event.preventDefault();
          event.stopImmediatePropagation();
          if (!onPreviewActivate) {
            document
              .querySelector<HTMLElement>(".source-panel .cm-content")
              ?.focus({ preventScroll: true });
            return;
          }
          previewOnlyRef.current = false;
          for (const field of fieldRefs.current.values()) {
            if (!field.isConnected) continue;
            field.classList.remove(visualTexSourcePreviewClass);
            field.readOnly = readOnly;
          }
          onPreviewActivate?.();

          const [lineId, field] = entry;
          const offset = Math.max(
            0,
            Math.min(
              field.getOffsetFromPoint(event.clientX, event.clientY, { bias: 0 }),
              field.lastOffset,
            ),
          );
          field.focus();
          field.position = offset;
          field.shadowRoot
            ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
            ?.focus({ preventScroll: true });
          setActiveLine(lineId);
          return;
        }

        clearMultiLineSelection();
        const contentBounds = entry[1].shadowRoot
          ?.querySelector<HTMLElement>('[part="content"]')
          ?.getBoundingClientRect();
        const startedOutsideFormula = Boolean(
          contentBounds &&
            (event.clientX < contentBounds.left - 6 ||
              event.clientX > contentBounds.right + 6),
        );
        const anchor = resolveMultiLineSelectionPoint(
          event.clientX,
          event.clientY,
          entry[0],
        );
        if (!anchor) return;
        pointerSelectionSessionRef.current = {
          pointerId: event.pointerId,
          startX: event.clientX,
          startY: event.clientY,
          anchor,
          allowSameLine: startedOutsideFormula,
          active: false,
        };
      };

      const handleWindowPointerMove = (event: PointerEvent) => {
        const session = pointerSelectionSessionRef.current;
        if (!session) {
          if (multiLineSelectionRef.current && event.buttons === 0) {
            event.stopImmediatePropagation();
          }
          return;
        }
        if (event.pointerId !== session.pointerId) return;
        const distance = Math.hypot(
          event.clientX - session.startX,
          event.clientY - session.startY,
        );
        if (!session.active && distance < 5) return;

        const focus = resolveMultiLineSelectionPoint(
          event.clientX,
          event.clientY,
        );
        if (!focus) return;
        if (
          !session.active &&
          focus.lineId === session.anchor.lineId &&
          !session.allowSameLine
        ) {
          return;
        }
        session.active = true;
        event.preventDefault();
        event.stopImmediatePropagation();
        applyMultiLineSelection(session.anchor, focus);
      };

      const handleWindowPointerEnd = (event: PointerEvent) => {
        const session = pointerSelectionSessionRef.current;
        if (!session || event.pointerId !== session.pointerId) return;
        pointerSelectionSessionRef.current = null;
        // WebKit can deliver pointerup past the last pointermove. Use the
        // release coordinates as the final selection endpoint; otherwise a
        // quick reverse drag can leave only part of a ket/fraction highlighted.
        const releaseFocus = resolveMultiLineSelectionPoint(
          event.clientX,
          event.clientY,
        );
        if (!session.active) {
          const distance = Math.hypot(
            event.clientX - session.startX,
            event.clientY - session.startY,
          );
          if (
            distance < 5 ||
            !releaseFocus ||
            (releaseFocus.lineId === session.anchor.lineId &&
              !session.allowSameLine)
          ) return;
          session.active = true;
        }
        if (releaseFocus) applyMultiLineSelection(session.anchor, releaseFocus);

        // Do not prevent or stop this pointerup. MathLive owns pointer capture
        // on the active mathfield and must finish its own drag tracker first.
        // Reapply the cross-line ranges in the next task, after MathLive's
        // target-phase pointerup handlers have completed, so later free mouse
        // movement cannot collapse the final line again.
        for (const field of fieldRefs.current.values()) {
          visualTexPointerSelectingFields.delete(field);
          field.classList.remove(visualTexPointerSelectingClass);
        }
        const selection = multiLineSelectionRef.current;
        if (!selection) return;
        window.setTimeout(() => {
          const currentSelection = multiLineSelectionRef.current;
          if (!currentSelection) return;
          applyMultiLineSelection(
            currentSelection.anchor,
            currentSelection.focus,
          );
          setActiveLine(currentSelection.focus.lineId);
        }, 0);
      };

      const handleWindowMouseMove = (event: MouseEvent) => {
        if (
          !pointerSelectionSessionRef.current &&
          multiLineSelectionRef.current &&
          event.buttons === 0
        ) {
          event.stopImmediatePropagation();
        }
      };

      const handleWindowPointerCancel = (event: PointerEvent) => {
        const session = pointerSelectionSessionRef.current;
        if (!session || event.pointerId !== session.pointerId) return;
        pointerSelectionSessionRef.current = null;
        for (const field of fieldRefs.current.values()) {
          visualTexPointerSelectingFields.delete(field);
          field.classList.remove(visualTexPointerSelectingClass);
        }
        clearMultiLineSelection();
      };

      const handleMultiLineCopy = (event: ClipboardEvent) => {
        const selection = multiLineSelectionRef.current;
        if (!selection || !event.clipboardData) return;
        const selectedLines = linesRef.current.flatMap((line) => {
          if (!multiLineSelectedIdsRef.current.has(line.id)) return [];
          const field = fieldRefs.current.get(line.id);
          if (!field?.isConnected) return [];
          const fieldSelection = captureSelection(field);
          const latex = fieldSelection.ranges
            .map(([start, end]) =>
              field.getValue(
                Math.min(start, end),
                Math.max(start, end),
                "latex-expanded",
              ),
            )
            .join("");
          return [normalizeChineseLatex(latex)];
        });
        if (selectedLines.length <= 1) return;

        const latex = selectedLines.join("\n");
        event.clipboardData.setData(
          VISUALTEX_MULTILINE_LATEX_CLIPBOARD_TYPE,
          JSON.stringify({ version: 1, lines: selectedLines }),
        );
        event.clipboardData.setData("application/x-latex", latex);
        event.clipboardData.setData("text/plain", formatLatexLines(selectedLines, latexCodeFormat));
        event.preventDefault();
        event.stopImmediatePropagation();
        if (event.type === "cut") deleteMultiLineSelection();
      };

      surface.addEventListener("pointerdown", handleSurfacePointerDown, true);
      surface.addEventListener("copy", handleMultiLineCopy, true);
      surface.addEventListener("cut", handleMultiLineCopy, true);
      window.addEventListener("pointermove", handleWindowPointerMove, true);
      window.addEventListener("mousemove", handleWindowMouseMove, true);
      window.addEventListener("pointerup", handleWindowPointerEnd, true);
      window.addEventListener("pointercancel", handleWindowPointerCancel, true);
      return () => {
        surface.removeEventListener("pointerdown", handleSurfacePointerDown, true);
        surface.removeEventListener("copy", handleMultiLineCopy, true);
        surface.removeEventListener("cut", handleMultiLineCopy, true);
        window.removeEventListener("pointermove", handleWindowPointerMove, true);
        window.removeEventListener("mousemove", handleWindowMouseMove, true);
        window.removeEventListener("pointerup", handleWindowPointerEnd, true);
        window.removeEventListener("pointercancel", handleWindowPointerCancel, true);
      };
    });

    useEffect(() => {
      const lineId =
        lines.find((line) => line.id === activeLineId)?.id ??
        lines[0]?.id ??
        null;
      const index = lineId
        ? Math.max(0, lines.findIndex((line) => line.id === lineId))
        : 0;
      activeLineIdRef.current = lineId;
      activeIndexRef.current = index;
      setActiveIndex(index);
    }, [lines, activeLineId]);

    useEffect(() => {
      if (!contextMenu) return;
      const close = () => setContextMenu(null);
      const handleKeyDown = (event: KeyboardEvent) => {
        if (event.key === "Escape") close();
      };
      window.addEventListener("pointerdown", close);
      window.addEventListener("blur", close);
      window.addEventListener("keydown", handleKeyDown);
      return () => {
        window.removeEventListener("pointerdown", close);
        window.removeEventListener("blur", close);
        window.removeEventListener("keydown", handleKeyDown);
      };
    }, [contextMenu]);

    useEffect(() => {
      if (previewOnly) setContextMenu(null);
    }, [previewOnly]);

    const openContextMenu = (clientX: number, clientY: number) => {
      if (previewOnly) return;
      const menuWidth = 188;
      const menuHeight = 42;
      setContextMenu({
        left: Math.max(8, Math.min(clientX, window.innerWidth - menuWidth - 8)),
        top: Math.max(8, Math.min(clientY, window.innerHeight - menuHeight - 8)),
      });
    };

    const copyPngFromContextMenu = async () => {
      if (contextMenuBusy) return;
      setContextMenuBusy(true);
      setContextMenu(null);
      try {
        if (onCopyPng) {
          await onCopyPng();
        } else {
          await copyFormulaDocumentPngToClipboard(
            linesRef.current.map((line) => line.latex),
            {
              background: pngExportBackground,
              formulaLetterFont,
              formulaChineseFont,
            },
          );
        }
      } catch (error) {
        console.error("Unable to copy formula PNG from the context menu", error);
      } finally {
        setContextMenuBusy(false);
      }
    };

    return (
      <div
        ref={surfaceRef}
        className={
          "editor-surface multi-line-editor" +
          (showLineNumbers ? " has-line-numbers" : "") +
          (showLineModeMarkers ? " has-mixed-line-modes" : "") +
          (interactionReadOnly ? " is-read-only-preview" : "") +
          (previewOnly ? " is-source-preview-only" : "")
        }
        style={
          {
            "--formula-area-inset-left": `${formulaInsetLeft}px`,
            "--formula-area-inset-right": `${formulaInsetRight}px`,
            "--formula-row-vertical-inset": `${formulaRowVerticalInset}px`,
          } as CSSProperties
        }
        data-command-query={previewOnly ? "" : query}
        data-source-draft-error={draftError}
        data-active-line-id={activeLineIdRef.current ?? ""}
        data-formula-alignment={formulaAlignment}
      >
        {overlay}
        <div className="mathfield-stack">
          {lines.map((line, index) => {
            const lineId = line.id;
            const lineHasInternalMultiline =
              /^\\begin\{(?:aligned|gathered)\}/.test(line.latex.trim());
            const effectiveLineMode = lineHasInternalMultiline
              ? "display"
              : line.mode === "inline"
                ? "inline"
                : "display";
            const safetyIssue = inspectMathLiveSourceSafety(line.latex);
            return (
              <div
                className={
                  "formula-line " +
                  (lineId === activeLineIdRef.current ? "is-active " : "") +
                  (multiLineSelectedIdsRef.current.has(lineId)
                    ? "is-multi-line-selected"
                    : "")
                }
                data-line-id={lineId}
                key={reuseLineSlots ? `resident-line-slot-${index}` : lineId}
              >
                {showLineNumbers ? (
                  <span className="formula-line-number">
                    {String(index + 1).padStart(2, "0")}
                  </span>
                ) : null}
                {showLineModeMarkers ? (
                  <button
                    type="button"
                    className="formula-line-mode-toggle is-active"
                    disabled={lineHasInternalMultiline || interactionReadOnly}
                    data-formula-line-mode-toggle
                    data-formula-line-mode={effectiveLineMode}
                    aria-label={
                      language === "en"
                        ? lineHasInternalMultiline
                          ? "Multi-line formula: display mode"
                          : effectiveLineMode === "inline"
                            ? "Switch row to display formula"
                            : "Switch row to inline formula"
                        : lineHasInternalMultiline
                          ? "多行公式固定为行间公式"
                          : effectiveLineMode === "inline"
                            ? "切换为行间公式"
                            : "切换为行内公式"
                    }
                    title={
                      language === "en"
                        ? lineHasInternalMultiline
                          ? "Multi-line formulas are always display formulas"
                          : effectiveLineMode === "inline"
                            ? "Inline · click for display · Alt+Enter creates display"
                            : "Display · click for inline · Ctrl+Enter creates inline"
                        : lineHasInternalMultiline
                          ? "多行公式固定为行间公式"
                          : effectiveLineMode === "inline"
                            ? "行内 · 点击切换行间 · Alt+Enter 新建行间"
                            : "行间 · 点击切换行内 · Ctrl+Enter 新建行内"
                    }
                    onPointerDown={(event) => event.preventDefault()}
                    onClick={() =>
                      setFormulaLineMode(
                        lineId,
                        effectiveLineMode === "inline" ? "display" : "inline",
                      )
                    }
                  >
                    <span aria-hidden="true">
                      {effectiveLineMode === "inline" ? "$" : "$$"}
                    </span>
                  </button>
                ) : null}
                {safetyIssue ? (
                  <FormulaFieldRenderFallback
                    message={`${mathLiveSourceSafetyMessage(safetyIssue, language)} ${
                      language === "en"
                        ? "The LaTeX source is preserved; edit or remove this row in the source panel."
                        : "LaTeX 源码仍完整保留，可在源码区修改或删除这一行。"
                    }`}
                  />
                ) : (
                  <FormulaFieldErrorBoundary
                    recoveryKey={[
                      lineId,
                      line.latex,
                      zoom,
                      formulaRowVerticalInset,
                      formulaLetterFont,
                      formulaChineseFont,
                      autoPairDelimiters,
                      interactionReadOnly,
                      previewOnly,
                      fieldRenderEpoch,
                    ].join("\u0000")}
                  >
                <FormulaField
                  key={`formula-field-${reuseLineSlots ? index : lineId}-${fieldRenderEpoch}`}
                  lineId={lineId}
                  index={index}
                  latex={line.latex}
                  zoom={zoom}
                  formulaRowVerticalInset={formulaRowVerticalInset}
                  language={language}
                  formulaLetterFont={formulaLetterFont}
                  formulaChineseFont={formulaChineseFont}
                  autoPairDelimiters={autoPairDelimiters}
                  inputBehavior={inputBehavior}
                  persistentTypingStyle={persistentTypingStyle}
                  readOnly={interactionReadOnly}
                  freshExternalSync={previewOnly}
                  register={registerField}
                  onEdit={handleFieldEdit}
                  onInputActivity={(field) =>
                    refreshSuggestionQuery(
                      lineId,
                      field,
                      normalizeChineseLatex(field.value),
                    )
                  }
                  onSelectionChange={rememberSelectionTarget}
                  onCommitPending={commitPendingVisualTransaction}
                  onFocus={(_lineIndex, field) => {
                    setActiveLine(lineId);
                    const normalizedValue = normalizeChineseLatex(field.value);
                    suppressedSuggestionRef.current = {
                      lineId,
                      value: normalizedValue,
                    };
                    queryRef.current = "";
                    setQuery("");
                  }}
                  onKeyDown={(lineIndex, event, field) =>
                    handleKeyDown(lineIndex, lineId, event, field)
                  }
                  onPasteImage={onPasteImage}
                  onContextMenu={openContextMenu}
                  onPasteLatexLines={pasteLatexLines}
                />
                  </FormulaFieldErrorBoundary>
                )}
              </div>
            );
          })}
        </div>
        {contextMenu && (
          <div
            className="formula-editor-context-menu"
            role="menu"
            aria-label={language === "en" ? "Formula actions" : "公式操作"}
            style={{ left: contextMenu.left, top: contextMenu.top }}
            onPointerDown={(event) => event.stopPropagation()}
            onContextMenu={(event) => event.preventDefault()}
          >
            <button
              type="button"
              role="menuitem"
              disabled={contextMenuBusy}
              onClick={() => void copyPngFromContextMenu()}
            >
              <ClipboardCopy size={15} />
              <span>
                {language === "en"
                  ? "Copy PNG to Clipboard"
                  : "复制 PNG 到剪贴板"}
              </span>
            </button>
          </div>
        )}
      </div>
    );
  },
);
