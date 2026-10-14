import { useHistoryManager } from "../history/EditorSession";
import {
  formatFormulaLinesUniversal,
  formatFormulaSelectionUniversal,
} from "../clipboard/LatexCopyService";
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
  LatexCommand,
} from "../types/command";
import type {
  FormulaAlignment,
  FormulaDisplayStyle,
  FormulaLine,
  FormulaLineMode,
  LatexCodeFormat,
  LatexFormatProfile,
  InputBehaviorSettings,
} from "../types/formula";
import type {
  AddLineEntry,
  DocumentSelection,
  DocumentSelectionPoint,
  DocumentSnapshot,
  EditKind,
  FormulaEditSource,
  MathSelectionSnapshot,
  RemoveLineEntry,
  ReplaceDocumentEntry,
} from "../history/historyTypes";
import {
  clampSelection,
} from "../history/HistoryManager";
import { commitFormulaEdit, getEditorDocumentSnapshot } from "../history/documentHistory";
import {
  createFormulaLine,
  useEditorStore,
} from "../stores/editorStore";
import { formatLatexSourceForEditor } from "../source-editor/latexSourceEditorSupport";
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
  documentSelection?: DocumentSelection | null;
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
  captureDocumentSnapshot: () => DocumentSnapshot;
  restoreDocumentSelection: (snapshot: DocumentSnapshot) => boolean;
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
  mode?: FormulaLineMode;
  displayStyle?: FormulaDisplayStyle;
  latexFormatProfile: LatexFormatProfile;
  zoom: number;
  formulaRowVerticalInset: number;
  language: "cn" | "en";
  formulaLetterFont: FormulaLetterFont;
  formulaChineseFont: FormulaChineseFont;
  autoPairDelimiters: boolean;
  inputBehavior: InputBehaviorSettings;
  persistentTypingStyle: MathLivePersistentTypingStyle;
  readOnly: boolean;
  register: (
    lineId: string,
    field: MathfieldElement | null,
    expectedField?: MathfieldElement,
  ) => void;
  onEdit: (edit: FormulaFieldEdit, field: MathfieldElement) => void;
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

interface VisualTexClipboardSelectionRequest {
  latex: string;
  wholeLine: boolean;
}

type VisualTexClipboardMathfield = MathfieldElement & {
  visualTexClipboardFormatter?: (
    request: VisualTexClipboardSelectionRequest,
  ) => string;
};

interface MultiLineSelectionPoint extends DocumentSelectionPoint {
  lineIndex: number;
}

interface PointerSelectionSession {
  pointerId: number;
  startX: number;
  startY: number;
  anchor: MultiLineSelectionPoint;
  active: boolean;
}


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

interface EditorScrollSnapshot {
  scroller: HTMLElement;
  anchorLineId: string | null;
  anchorOffset: number;
  scrollTop: number;
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
    const visualTexClipboardField = field as VisualTexClipboardMathfield;
    visualTexClipboardField.visualTexClipboardFormatter = ({
      latex,
      wholeLine,
    }) => {
      const current = propsRef.current;
      const source = formatFormulaSelectionUniversal(
        {
          id: current.lineId,
          latex: current.latex,
          mode: current.mode,
          displayStyle: current.displayStyle,
        },
        latex,
        current.latexFormatProfile,
        wholeLine,
      );
      return formatLatexSourceForEditor(source);
    };
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
    let measuredLatex: string | null = null;
    const measureFrameSize = () => {
      measuredLatex = field.value;
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
    const syncFrameSize = () => {
      window.cancelAnimationFrame(resizeFrame);
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
      lastSnapshotRef.current = captureFieldSnapshot(field);
      field.resetUndo();
    };
    const handleCompositionStart = () => {
      compositionDeleteObserved = false;
      suppressPostCompositionDeleteUntil = 0;
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
        field.resetUndo();
        lastSnapshotRef.current = captureFieldSnapshot(field);
        compositionStartRef.current = null;
        syncFrameSize();
        return;
      }

      const after = captureFieldSnapshot(field);
      compositionStartRef.current = null;
      emitEdit(before, after, "composition", "keyboard");
      syncFrameSize();
    };
    const handleBeforeInput = (event: InputEvent) => {
      if (
        event.inputType === "deleteContentBackward" &&
        event.timeStamp <= suppressPostCompositionDeleteUntil
      ) {
        suppressPostCompositionDeleteUntil = 0;
        event.preventDefault();
        event.stopImmediatePropagation();
        return;
      }
    };
    const handleInput = (event: Event) => {
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
      syncFrameSize();
      window.requestAnimationFrame(() => {
        if (!field.isConnected || imeGuard.isComposing()) return;
        const deferredBefore =
          lastSnapshotRef.current ?? captureFieldSnapshot(field);
        if (
          normalizeCompletedDifferentialDisplay(field)
        ) {
          emitEdit(
            deferredBefore,
            captureFieldSnapshot(field),
            "replace",
            "keyboard",
          );
        }
      });
    };
    const handleSelectionChange = () => {
      syncStructuralPlaceholderSelection(field);
      syncPostOperatorCaretSpacing(field);
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
      propsRef.current.onCommitPending();
    };
    const handleKeyDown = (event: KeyboardEvent) => {
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
        return;
      }

      propsRef.current.onKeyDown(propsRef.current.index, event, field);
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
    };
    const handlePointerDown = (event: PointerEvent) => {
      const opensContextMenu =
        event.button === 2 || (event.button === 0 && event.ctrlKey);
      if (opensContextMenu) return;
      if (event.button !== 0) return;
      visualTexPointerSelectingFields.add(field);
      field.classList.add(visualTexPointerSelectingClass);
      propsRef.current.onCommitPending();
      if (!event.composedPath().includes(field)) field.visualTexStartPointerSelection(event);
    };
    host.replaceChildren(field);
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
          if (field.value !== measuredLatex) syncFrameSize();
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

      if (VISUALTEX_IME_DIAGNOSTICS_ENABLED) {
        for (const type of imeDiagnosticEventTypes) {
          window.removeEventListener(type, handleWindowImeDiagnosticEvent, true);
          field.removeEventListener(type, handleFieldImeDiagnosticEvent, true);
          keyboardSink?.removeEventListener(type, handleSinkImeDiagnosticEvent, true);
        }
      }
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
      delete (field as VisualTexClipboardMathfield).visualTexClipboardFormatter;
      propsRef.current.register(registeredLineIdRef.current, null, field);
      fieldRef.current = null;
      defaultInlineShortcutsRef.current = null;
      lastSnapshotRef.current = null;
      compositionStartRef.current = null;

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
      if (!zoomChanged && !rowVerticalInsetChanged) {
        propsRef.current.register(props.lineId, field);
        return;
      }
    } else {
      // replaceAll creates a fresh root in the MathLive adapter.
      field.setValue(props.latex, {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
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
    propsRef.current.register(props.lineId, field);
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
      overlay,
    },
    ref,
  ) {
    const surfaceRef = useRef<HTMLDivElement>(null);
    const historyManager = useHistoryManager();
    const fieldRefs = useRef(new Map<string, MathfieldElement>());
    const linesRef = useRef(lines);
    const activeIndexRef = useRef(0);
    const activeLineIdRef = useRef<string | null>(activeLineId);
    const previewOnlyRef = useRef(previewOnly);
    const greekLetterHotkeyLineIdRef = useRef<string | null>(null);
    const suppressedHistoryLineIdRef = useRef<string | null>(null);
    const documentSelectionRef = useRef<DocumentSelection | null>(null);
    const documentSelectionNavigationRef = useRef<string | null>(null);
    const pointerSelectionSessionRef = useRef<PointerSelectionSession | null>(
      null,
    );
    const [, setMultiLineSelectionRevision] = useState(0);
    const pendingFocusRef = useRef<{
      lineId: string;
      latex: string | null;
      selection: MathSelectionSnapshot | null;
      moveToEnd: boolean;
      document?: DocumentSnapshot;
    } | null>(null);
    const [fieldRenderEpoch, setFieldRenderEpoch] = useState(0);
    const [contextMenu, setContextMenu] = useState<{
      left: number;
      top: number;
    } | null>(null);
    const [contextMenuBusy, setContextMenuBusy] = useState(false);
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

    const applyPendingFocus = () => {
      const request = pendingFocusRef.current;
      if (!request) return false;
      const { lineId, latex, selection, moveToEnd } = request;
      const line = useEditorStore.getState().lines.find(item => item.id === lineId);
      if (!line || activeLineIdRef.current !== lineId || (latex !== null && line.latex !== latex)) {
        pendingFocusRef.current = null;
        return false;
      }
      const field = fieldRefs.current.get(lineId);
      if (!field?.isConnected || (latex !== null && normalizeChineseLatex(field.value) !== latex)) return false;
      if (request.document) {
        for (const item of request.document.lines) {
          const mounted = fieldRefs.current.get(item.id);
          if (!mounted?.isConnected || normalizeChineseLatex(mounted.value) !== item.latex) return false;
        }
        clearDocumentSelection();
        for (const [id, saved] of Object.entries(request.document.selectionByLineId)) {
          const mounted = fieldRefs.current.get(id);
          if (mounted) mounted.selection = clampSelection(saved, mounted.lastOffset);
        }
        if (request.document.documentSelection) {
          const { anchor, focus } = request.document.documentSelection;
          applyDocumentSelection(anchor, focus);
        }
      }
      // Content is applied by FormulaField's layout effect. Focusing never
      // rewrites LaTeX, and the mount/content notification consumes this once.
      pendingFocusRef.current = null;
      if (selection && !request.document?.documentSelection) {
        field.selection = clampSelection(selection, field.lastOffset);
      } else if (moveToEnd) {
        field.selection = { ranges: [[field.lastOffset, field.lastOffset]], direction: "none" };
      }
      (field as HTMLElement).focus({ preventScroll: true });
      field.shadowRoot?.querySelector<HTMLElement>('[part="keyboard-sink"]')?.focus({ preventScroll: true });
      return true;
    };

    const focusLine = (
      lineId: string,
      options: { latex?: string | null; selection?: MathSelectionSnapshot | null; moveToEnd?: boolean; document?: DocumentSnapshot } = {},
    ) => {
      const index = linesRef.current.findIndex(line => line.id === lineId);
      if (index < 0) return false;
      activeIndexRef.current = index;
      activeLineIdRef.current = lineId;
      const state = useEditorStore.getState();
      if (state.activeLineId !== lineId) state.setActiveLineId(lineId);
      pendingFocusRef.current = {
        lineId, latex: options.latex ?? null, selection: options.selection ?? null,
        moveToEnd: options.moveToEnd ?? false, document: options.document,
      };
      applyPendingFocus();
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

      if (field) applyPendingFocus();
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
      pendingFocusRef.current = null;
      const index = linesRef.current.findIndex((line) => line.id === lineId);
      if (index >= 0) {
        activeIndexRef.current = index;
        activeLineIdRef.current = lineId;
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
        pendingFocusRef.current = null;
      }
      activeIndexRef.current = index;
      activeLineIdRef.current = lineId;
      const state = useEditorStore.getState();
      if (state.activeLineId !== lineId) state.setActiveLineId(lineId);
    };

    const clearDocumentSelection = () => {
      for (const field of fieldRefs.current.values()) {
        if (!field.isConnected || !field.classList.contains("has-visualtex-multi-line-selection")) continue;
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
      const hadSelection = documentSelectionRef.current !== null;
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
      pointerSelectionSessionRef.current = null;
      clearDocumentSelection();

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
      const offset = Math.max(0, Math.min(
        target.field.getOffsetFromPoint(clientX, sampleY, { bias: 0 }),
        target.field.lastOffset,
      ));
      return {
        lineId: target.lineId,
        lineIndex: target.lineIndex,
        offset,
      };
    };

    const documentSelectionBounds = (selection = documentSelectionRef.current) => {
      if (!selection) return null;
      const anchorIndex = linesRef.current.findIndex(line => line.id === selection.anchor.lineId);
      const focusIndex = linesRef.current.findIndex(line => line.id === selection.focus.lineId);
      if (anchorIndex < 0 || focusIndex < 0) return null;
      const forward = anchorIndex < focusIndex ||
        (anchorIndex === focusIndex && selection.anchor.offset <= selection.focus.offset);
      return {
        start: { ...(forward ? selection.anchor : selection.focus), lineIndex: Math.min(anchorIndex, focusIndex) },
        end: { ...(forward ? selection.focus : selection.anchor), lineIndex: Math.max(anchorIndex, focusIndex) },
        forward,
      };
    };

    const selectedDocumentLineIds = () => {
      const bounds = documentSelectionBounds();
      return new Set(bounds ? linesRef.current.slice(bounds.start.lineIndex, bounds.end.lineIndex + 1).map(line => line.id) : []);
    };

    const applyDocumentSelection = (anchor: DocumentSelectionPoint, focus: DocumentSelectionPoint) => {
      const bounds = documentSelectionBounds({ anchor, focus });
      if (!bounds) return;
      clearDocumentSelection();
      const { start, end, forward } = bounds;
      for (let index = start.lineIndex; index <= end.lineIndex; index++) {
        const line = linesRef.current[index];
        const field = fieldRefs.current.get(line.id);
        if (!field?.isConnected) continue;
        const from = index === start.lineIndex ? start.offset : 0;
        const to = index === end.lineIndex ? end.offset : field.lastOffset;
        field.visualTexSelectRange(forward ? from : to, forward ? to : from);
        // The kernel expands a partial structural boundary. Store that same
        // boundary so replacement, copy and the painted selection agree.
        const range = field.selection.ranges[0];
        if (index === start.lineIndex) start.offset = range[0];
        if (index === end.lineIndex) end.offset = range[1];
        field.classList.add("has-visualtex-multi-line-selection");
        field.closest<HTMLElement>(".formula-line")?.classList.add("is-multi-line-selected");
      }
      const point = ({ lineId, offset }: DocumentSelectionPoint) => ({ lineId, offset });
      documentSelectionRef.current = {
        anchor: point(forward ? start : end), focus: point(forward ? end : start),
      };
      setMultiLineSelectionRevision(revision => revision + 1);
    };

    const rememberSelectionTarget = (lineId: string) => {
      if (documentSelectionNavigationRef.current !== lineId) return;
      documentSelectionNavigationRef.current = null;
      const selection = documentSelectionRef.current;
      const field = fieldRefs.current.get(lineId);
      if (selection && field) {
        applyDocumentSelection(selection.anchor, { lineId, offset: field.position });
      }
    };

    // MathLive owns edits within a formula. The document owns replacing a
    // selection that spans independently mounted formula fields.
    const finishDocumentSelectionEdit = (
      before: DocumentSnapshot,
      lineId: string,
      field: MathfieldElement,
      source: ReplaceDocumentEntry["source"] = "replace-multi-line",
    ) => {
      const bounds = documentSelectionBounds(before.documentSelection);
      if (!bounds || bounds.start.lineIndex === bounds.end.lineIndex) return false;
      const { start, end } = bounds;
      const first = fieldRefs.current.get(start.lineId);
      const last = fieldRefs.current.get(end.lineId);
      const line = before.lines.find(item => item.id === lineId);
      if (!first || !last || !line) return false;
      const prefix = lineId === start.lineId ? "" : first.getValue(0, start.offset, "latex");
      const suffix = lineId === end.lineId ? "" : last.getValue(end.offset, last.lastOffset, "latex");
      field.visualTexInsertDocumentEdges(prefix, suffix);
      const selection = captureSelection(field);
      const latex = normalizeChineseLatex(field.value);
      const lines = [...before.lines];
      lines.splice(start.lineIndex, end.lineIndex - start.lineIndex + 1, {
        ...line, latex, ...(source === "multiline-row-break" ? { mode: "display" as const } : {}),
      });
      const after: DocumentSnapshot = {
        ...before, lines, activeLineId: lineId, documentSelection: null,
        selectionByLineId: { [lineId]: selection },
      };
      clearDocumentSelection();
      flushSync(() => useEditorStore.getState().replaceDocumentState(after));
      linesRef.current = useEditorStore.getState().lines;
      focusLine(lineId, { latex, selection });
      historyManager.push({ type: "replace-document", before, after, source, timestamp: Date.now() });
      return true;
    };

    const handleFieldEdit = (
      edit: FormulaFieldEdit,
      field: MathfieldElement,
    ) => {
      if (suppressedHistoryLineIdRef.current === edit.lineId) return;
      if (documentSelectionRef.current && !historyManager.getState().isReplaying) {
        historyManager.commitPendingTransaction();
        const before = captureDocumentSnapshot();
        before.selectionByLineId[edit.lineId] = edit.beforeSelection;
        if (finishDocumentSelectionEdit(before, edit.lineId, field)) return;
      }
      if (documentSelectionRef.current) clearDocumentSelection();
      commitFormulaEdit(historyManager, {
        ...edit,
        beforeActiveLineId: activeLineIdRef.current,
        afterActiveLineId: edit.lineId,
      });
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(edit.lineId);
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
        ? captureDocumentSnapshot() : null;

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
        if (finishDocumentSelectionEdit(documentBefore, lineId, field)) return true;
        clearDocumentSelection();
      }
      if (before.latex === after.latex) {
        field.resetUndo();
        return true;
      }
      field.resetUndo();
      commitFormulaEdit(historyManager, {
        lineId,
        beforeLatex: before.latex,
        afterLatex: after.latex,
        beforeSelection: before.selection,
        afterSelection: after.selection,
        beforeActiveLineId,
        afterActiveLineId: lineId,
        timestamp: Date.now(),
        source,
        editKind: "replace",
      }, recordHistory ? "discrete" : "silent");
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(lineId);
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
      const before = captureDocumentSnapshot(selectionByLineId);

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

      if (finishDocumentSelectionEdit(before, lineId, field, "multiline-row-break")) return true;
      const after: ReplaceDocumentEntry["after"] = {
        ...before,
        lines: before.lines.map((line) =>
              line.id === lineId
                ? { ...line, latex: afterLatex, mode: "display" }
                : { ...line },
            ),
        activeLineId: lineId,
        documentSelection: null,
        selectionByLineId: { ...before.selectionByLineId, [lineId]: afterSelection },
      };

      clearDocumentSelection();
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
      source: "toolbar" | "history" | "shortcut" = "toolbar",
    ) => {
      historyManager.commitPendingTransaction();
      const target = resolveTargetField();
      if (!target) return;
      const { lineId: targetLineId, field } = target;
      setActiveLine(targetLineId);
      field.focus();

      const originalSelection = captureSelection(field);
      let insertionSelection = originalSelection;
      let implicitPreviousLatex = "";
      if (
        field.selectionIsCollapsed &&
        commandTargetsPreviousExpression(command)
      ) {
        const previousTarget = previousExpressionSelection(field);
        if (previousTarget) {
          insertionSelection = previousTarget.selection;
          implicitPreviousLatex = previousTarget.latex;
        }
      }
      const selectedLatex =
        implicitPreviousLatex ||
        (insertionSelection.ranges[0]?.[0] === insertionSelection.ranges[0]?.[1]
          ? ""
          : field.getValue(insertionSelection));
      const insertionTemplate = templateForSelection(command, selectedLatex);
      const historySource: FormulaEditSource =
        source === "shortcut" ? "shortcut" : "toolbar";

      const performInsertion = () => {
        field.selection = insertionSelection;
        const inserted = applyDiscreteFormulaMutation(
          targetLineId,
          field,
          historySource,
          () => {
            const hasPlaceholder = insertionTemplate.includes("\\placeholder{}");
            const inserted = field.visualTexInsertToolbarTemplate(insertionTemplate, {
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
            return inserted;
          },
        );
        if (!inserted) field.selection = originalSelection;
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
        recordCommand(command.id, "", source);
        focusLine(targetLineId, {
          latex: normalizedValue,
          selection: captureSelection(field),
        });
      };

      if (performInsertion()) finishInsertion();
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
      const before = captureDocumentSnapshot(selectionByLineId);
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
      focusLine(line.id, {
        latex: line.latex,
        selection: afterSelection,
      });
      stabilizeEditorScroll(scrollSnapshot, line.id, false);
    };

    const pasteLatexLines = (lineId: string, field: MathfieldElement, values: string[]) => {
      insertMultipleLatexRowsAt({ lineId, ...captureSelection(field) }, values, "paste");
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
      const before = captureDocumentSnapshot(selectionByLineId);
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
      focusLine(nextLine.id, {
        latex: rightLatex,
        selection: startSelection,
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
      const before = captureDocumentSnapshot(selectionByLineId);
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
      finalizeFocusAfterStructuralRemoval(previousLine.id, joinSelection);
      stabilizeEditorScroll(scrollSnapshot, previousLine.id, false);
      return true;
    };

    const deleteMultiLineSelection = () => {
      const bounds = documentSelectionBounds();
      if (!bounds || bounds.start.lineIndex === bounds.end.lineIndex || interactionReadOnly) return false;
      const field = fieldRefs.current.get(bounds.start.lineId);
      return Boolean(field && applyDiscreteFormulaMutation(
        bounds.start.lineId, field, "keyboard",
        () => field.selectionIsCollapsed || field.executeCommand("deleteBackward"),
      ));
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
          applyDocumentSelection(
            { lineId: first.id, offset: 0 },
            { lineId: last.id, offset: lastField.lastOffset },
          );
        }
        return;
      }

      if (documentSelectionRef.current && event.shiftKey &&
          ["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown", "Home", "End"].includes(event.key)) {
        // Capture handlers run before MathLive navigation. Its selection-change
        // event is the authoritative endpoint, not a timed callback.
        documentSelectionNavigationRef.current = lineId;
      }

      if (documentSelectionRef.current && !event.shiftKey &&
          ["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown", "Home", "End", "Escape"].includes(event.key)) {
        const selectedIds = [...selectedDocumentLineIds()];
        const backwards = ["ArrowLeft", "ArrowUp", "Home"].includes(event.key);
        const targetId = event.key === "Escape" ? lineId : backwards ? selectedIds[0] : selectedIds.at(-1);
        const targetField = targetId ? fieldRefs.current.get(targetId) : null;
        if (targetId && targetField) {
          const ranges = targetField.selection.ranges.flat();
          const offset = backwards ? Math.min(...ranges) : Math.max(...ranges);
          clearDocumentSelection();
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
        clearDocumentSelection();
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
          focusLine(lineId, {
            latex: normalizeChineseLatex(field.value),
            selection: captureSelection(field),
          });
        }
        return;
      }

      const rawCommandActive = hasRawLatexInput(field);
      // Let native completion own keys while a command is actually pending.
      // Escape can leave MathLive in latex mode with no raw input; that state
      // must still allow formula-row navigation.
      if (rawCommandActive || (field.mode === "latex" &&
          event.key !== "ArrowUp" && event.key !== "ArrowDown")) return;
      const currentState = useEditorStore.getState();

      const activeEnvironment = field.visualTexParentEnvironment;
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
        field.visualTexParentEnvironment === "cases"
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

        field.resetUndo();
        commitFormulaEdit(historyManager, {
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
        linesRef.current = useEditorStore.getState().lines;
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
          focusLine(targetLine.id, {
            selection: targetSelection,
          });
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

    const getSelectionMap = (): Record<string, MathSelectionSnapshot> =>
      Object.fromEntries(
        linesRef.current.flatMap((line) => {
          const field = fieldRefs.current.get(line.id);
          return field?.isConnected
            ? [[line.id, captureSelection(field)] as const]
            : [];
        }),
      );

    const captureDocumentSnapshot = (selections = getSelectionMap()): DocumentSnapshot => {
      const snapshot = getEditorDocumentSnapshot(selections);
      const selection = documentSelectionRef.current;
      snapshot.documentSelection = selection ? {
        anchor: { ...selection.anchor }, focus: { ...selection.focus },
      } : null;
      return snapshot;
    };

    const restoreDocumentSelection = (snapshot: DocumentSnapshot) => {
      const line = snapshot.lines.find(item => item.id === snapshot.activeLineId);
      return Boolean(line && focusLine(line.id, {
        latex: line.latex, selection: snapshot.selectionByLineId[line.id], document: snapshot,
      }));
    };

    const captureInsertionTarget = (): MathEditorInsertionTarget | null => {
      const target = resolveTargetField();
      if (!target) return null;
      const selection = captureSelection(target.field);
      return {
        lineId: target.lineId,
        ranges: selection.ranges,
        direction: selection.direction,
        documentSelection: captureDocumentSnapshot().documentSelection,
      };
    };

    const captureSelectionTarget = (): MathEditorSelectionTarget | null => {
      // MathLive fields can retain a visual/model selection after focus moves to
      // another formula. A formatting button must never sweep those stale
      // selections from unrelated rows. Only an explicit VisualTeX multi-line
      // drag is allowed to target more than the active formula row.
      const explicitMultiLineSelection = Boolean(documentSelectionRef.current);
      const activeLineId =
        activeLineIdRef.current ?? useEditorStore.getState().activeLineId;
      const selectedIds = selectedDocumentLineIds();
      const candidateLines = explicitMultiLineSelection
        ? linesRef.current.filter((line) =>
            selectedIds.has(line.id),
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
        return target;
      }
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
        ? captureDocumentSnapshot() : null;
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
          after: captureDocumentSnapshot(), source: "format-multi-line", timestamp: Date.now() });
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
      return true;
    };

    const restoreSelection = async (
      lineId: string,
      latex: string,
      selection: MathSelectionSnapshot | null,
    ): Promise<boolean> => {
      historyManager.commitPendingTransaction();
      return focusLine(lineId, { latex, selection, moveToEnd: !selection });
    };

    const insertMultipleLatexRowsAt = (
      target: MathEditorInsertionTarget,
      values: readonly string[],
      source: FormulaEditSource,
    ): boolean => {
      if (interactionReadOnly || values.length < 2) return false;
      const field = fieldRefs.current.get(target.lineId);
      if (!field?.isConnected) return false;
      historyManager.commitPendingTransaction();
      const before = captureDocumentSnapshot();
      const bounds = documentSelectionBounds();
      const targetIndex = before.lines.findIndex(line => line.id === target.lineId);
      if (targetIndex < 0) return false;
      const selected = clampSelection(target, field.lastOffset).ranges[0];
      const start = bounds?.start ?? { lineId: target.lineId, lineIndex: targetIndex, offset: selected[0] };
      const end = bounds?.end ?? { lineId: target.lineId, lineIndex: targetIndex, offset: selected[1] };
      const first = fieldRefs.current.get(start.lineId);
      const last = fieldRefs.current.get(end.lineId);
      if (!first || !last) return false;
      const prefix = first.getValue(0, start.offset, "latex");
      const suffix = last.getValue(end.offset, last.lastOffset, "latex");
      const selectionByLineId: Record<string, MathSelectionSnapshot> = {};
      const insertOptions = { mode: "math" as const, format: "latex" as const, selectionMode: "after" as const, silenceNotifications: true };
      // Use the mounted model to measure the cursor before the preserved tail.
      // The last pasted row must not place the caret after that tail.
      field.setValue(values[values.length - 1], { ...insertOptions });
      field.position = field.lastOffset;
      const lastSelection = captureSelection(field);
      field.visualTexInsertDocumentEdges("", suffix);
      const lastLatex = normalizeChineseLatex(field.value);
      field.setValue(values[0], { ...insertOptions });
      field.position = field.lastOffset;
      field.visualTexInsertDocumentEdges(prefix, "");
      const firstLatex = normalizeChineseLatex(field.value);
      const template = before.lines[start.lineIndex];
      const inserted = [
        { ...template, id: target.lineId, latex: firstLatex },
        ...values.slice(1).map((latex, index) => createFormulaLine(
          index === values.length - 2 ? lastLatex : normalizeChineseLatex(latex),
          undefined, template.mode, template.displayStyle,
        )),
      ];
      const lastLine = inserted[inserted.length - 1];
      selectionByLineId[lastLine.id] = lastSelection;
      const lines = [...before.lines];
      lines.splice(start.lineIndex, end.lineIndex - start.lineIndex + 1, ...inserted);
      const after: DocumentSnapshot = {
        ...before, lines, activeLineId: lastLine.id, selectionByLineId, documentSelection: null,
      };
      clearDocumentSelection();
      flushSync(() => useEditorStore.getState().replaceDocumentState(after));
      linesRef.current = useEditorStore.getState().lines;
      field.resetUndo();
      historyManager.push({ type: "replace-document", before, after,
        source: source === "paste" ? "paste-multi-line" : "ocr", timestamp: Date.now() });
      focusLine(lastLine.id, { latex: lastLine.latex, selection: lastSelection });
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
      const field = fieldRefs.current.get(target.lineId);
      if (!field?.isConnected) return false;
      // Async insertions belong to the captured target, even if the user has
      // selected different rows while OCR was running.
      if (target.documentSelection) {
        if (!documentSelectionBounds(target.documentSelection)) return false;
        applyDocumentSelection(target.documentSelection.anchor, target.documentSelection.focus);
      } else clearDocumentSelection();
      prepareOcrAlignmentFormat(values, source);
      if (values.length > 1) {
        return insertMultipleLatexRowsAt(target, values, source);
      }
      const value = values[0];
      if (!linesRef.current.some((line) => line.id === target.lineId)) {
        return false;
      }
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
      field.selection = selection;
      const inserted = applyDiscreteFormulaMutation(
        target.lineId,
        field,
        source,
        () => {
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
      const before = captureDocumentSnapshot();
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
      captureDocumentSnapshot,
      restoreDocumentSelection,
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
      const surface = surfaceRef.current;
      if (!surface) return;

      const handleSurfacePointerDown = (event: PointerEvent) => {
        if (event.button !== 0 || !event.isPrimary || event.shiftKey) return;
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

        clearDocumentSelection();
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
          active: false,
        };
      };

      const handleWindowPointerMove = (event: PointerEvent) => {
        const session = pointerSelectionSessionRef.current;
        if (!session) return;
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
          focus.lineId === session.anchor.lineId
        ) {
          return;
        }
        if (!session.active) fieldRefs.current.get(session.anchor.lineId)?.visualTexStopPointerTracking();
        session.active = true;
        event.preventDefault();
        event.stopImmediatePropagation();
        applyDocumentSelection(session.anchor, focus);
      };

      const handleWindowPointerEnd = (event: PointerEvent) => {
        const session = pointerSelectionSessionRef.current;
        if (!session || event.pointerId !== session.pointerId) return;
        pointerSelectionSessionRef.current = null;
        const focus = resolveMultiLineSelectionPoint(event.clientX, event.clientY);
        if (!focus || (!session.active && focus.lineId === session.anchor.lineId)) return;
        fieldRefs.current.get(session.anchor.lineId)?.visualTexStopPointerTracking();
        const field = fieldRefs.current.get(focus.lineId);
        field?.focus();
        setActiveLine(focus.lineId);
        applyDocumentSelection(session.anchor, focus);
        for (const mounted of fieldRefs.current.values()) {
          visualTexPointerSelectingFields.delete(mounted);
          mounted.classList.remove(visualTexPointerSelectingClass);
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
        clearDocumentSelection();
      };

      const handleMultiLineCopy = (event: ClipboardEvent) => {
        const selection = documentSelectionRef.current;
        if (!selection || !event.clipboardData) return;
        const selectedIds = selectedDocumentLineIds();
        const selectedFormulaLines = linesRef.current.flatMap((line) => {
          if (!selectedIds.has(line.id)) return [];
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
          return [
            {
              ...line,
              latex: normalizeChineseLatex(latex),
            },
          ];
        });
        if (selectedFormulaLines.length <= 1) return;

        const selectedLines = selectedFormulaLines.map((line) => line.latex);
        const latex = selectedLines.join("\n");
        event.clipboardData.setData(
          VISUALTEX_MULTILINE_LATEX_CLIPBOARD_TYPE,
          JSON.stringify({ version: 1, lines: selectedLines }),
        );
        event.clipboardData.setData("application/x-latex", latex);
        event.clipboardData.setData(
          "text/plain",
          formatLatexSourceForEditor(
            formatFormulaLinesUniversal(
              selectedFormulaLines,
              latexFormatProfile,
            ),
          ),
        );
        event.preventDefault();
        event.stopImmediatePropagation();
        if (event.type === "cut") deleteMultiLineSelection();
      };

      surface.addEventListener("pointerdown", handleSurfacePointerDown, true);
      surface.addEventListener("copy", handleMultiLineCopy, true);
      surface.addEventListener("cut", handleMultiLineCopy, true);
      window.addEventListener("pointermove", handleWindowPointerMove, true);
      window.addEventListener("pointerup", handleWindowPointerEnd, true);
      window.addEventListener("pointercancel", handleWindowPointerCancel, true);
      return () => {
        surface.removeEventListener("pointerdown", handleSurfacePointerDown, true);
        surface.removeEventListener("copy", handleMultiLineCopy, true);
        surface.removeEventListener("cut", handleMultiLineCopy, true);
        window.removeEventListener("pointermove", handleWindowPointerMove, true);
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

    const selectedLineIds = selectedDocumentLineIds();
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
                  (selectedLineIds.has(lineId)
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
                  mode={line.mode}
                  displayStyle={line.displayStyle}
                  latexFormatProfile={latexFormatProfile}
                  zoom={zoom}
                  formulaRowVerticalInset={formulaRowVerticalInset}
                  language={language}
                  formulaLetterFont={formulaLetterFont}
                  formulaChineseFont={formulaChineseFont}
                  autoPairDelimiters={autoPairDelimiters}
                  inputBehavior={inputBehavior}
                  persistentTypingStyle={persistentTypingStyle}
                  readOnly={interactionReadOnly}
                  register={registerField}
                  onEdit={handleFieldEdit}
                  onSelectionChange={rememberSelectionTarget}
                  onCommitPending={commitPendingVisualTransaction}
                  onFocus={() => setActiveLine(lineId)}
                  onKeyDown={(lineIndex, event, field) =>
                    handleKeyDown(lineIndex, lineId, event, field)
                  }
                  onPasteImage={onPasteImage ? (file, target) => onPasteImage(file, {
                    ...target, documentSelection: captureDocumentSnapshot().documentSelection,
                  }) : undefined}
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
