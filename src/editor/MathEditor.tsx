import {
  useEffect,
  useLayoutEffect,
  useImperativeHandle,
  useMemo,
  useRef,
  useState,
  forwardRef,
  type ReactNode,
} from "react";
import { MathfieldElement } from "mathlive";
import { flushSync } from "react-dom";
import { Plus } from "lucide-react";
import type { CommandSource, LatexCommand } from "../types/command";
import type { FormulaLine } from "../types/formula";
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
import { searchCommands } from "../autocomplete/CommandSearchEngine";
import { CommandSuggestionPopup } from "../autocomplete/CommandSuggestionPopup";
import {
  createFormulaLine,
  useEditorStore,
} from "../stores/editorStore";
import { normalizeChineseLatex } from "./normalizeChineseLatex";

export interface MathEditorInsertionTarget {
  lineId: string;
  ranges: Array<[number, number]>;
  direction: "forward" | "backward" | "none";
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
  focus: () => void;
  addLine: () => void;
  commitPendingTransaction: () => void;
  getSelectionMap: () => Record<string, MathSelectionSnapshot>;
  restoreSelection: (
    lineId: string,
    latex: string,
    selection: MathSelectionSnapshot | null,
  ) => Promise<boolean>;
}

interface Props {
  lines: FormulaLine[];
  activeLineId: string | null;
  zoom: number;
  onPasteImage?: (file: File, target: MathEditorInsertionTarget) => void;
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
  language: "cn" | "en";
  register: (lineId: string, field: MathfieldElement | null) => void;
  onEdit: (edit: FormulaFieldEdit, field: MathfieldElement) => void;
  onFocus: (index: number, field: MathfieldElement) => void;
  onCommitPending: () => void;
  onKeyDown: (index: number, event: KeyboardEvent, field: MathfieldElement) => void;
  onPasteImage?: (file: File, target: MathEditorInsertionTarget) => void;
}

const trailingCommand = /\\([\p{L}]*)$/u;
const BASE_FORMULA_FONT_SIZE = 54;
const MIN_FORMULA_FONT_SIZE = BASE_FORMULA_FONT_SIZE * 0.2;

const formulaFontSize = (zoom: number) =>
  Math.max(MIN_FORMULA_FONT_SIZE, BASE_FORMULA_FONT_SIZE * zoom);

function captureSelection(field: MathfieldElement): MathSelectionSnapshot {
  return {
    ranges: field.selection.ranges.map(
      ([start, end]) => [start, end] as [number, number],
    ),
    direction: field.selection.direction ?? "none",
  };
}

function captureFieldSnapshot(field: MathfieldElement) {
  return {
    latex: normalizeChineseLatex(field.value),
    selection: captureSelection(field),
  };
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

function getVisibleNativeSuggestionItems(): HTMLElement[] {
  const panel = document.getElementById("mathlive-suggestion-popover");
  if (!panel?.classList.contains("is-visible")) return [];
  return Array.from(
    panel.querySelectorAll<HTMLElement>("li[data-command]"),
  );
}

function moveNativeSuggestionSelection(direction: 1 | -1): boolean {
  const items = getVisibleNativeSuggestionItems();
  if (!items.length) return false;

  const currentIndex = items.findIndex((item) =>
    item.classList.contains("ML__popover__current"),
  );
  const nextIndex =
    currentIndex < 0
      ? direction > 0
        ? 0
        : items.length - 1
      : (currentIndex + direction + items.length) % items.length;

  items.forEach((item, index) => {
    const selected = index === nextIndex;
    item.classList.toggle("ML__popover__current", selected);
    item.setAttribute("aria-selected", String(selected));
  });
  items[nextIndex].scrollIntoView({ block: "nearest" });
  return true;
}

function commitNativeSuggestion(): boolean {
  const items = getVisibleNativeSuggestionItems();
  if (!items.length) return false;
  const selected =
    items.find((item) => item.classList.contains("ML__popover__current")) ??
    items[0];
  selected.click();
  return true;
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

function templateForSelection(
  command: LatexCommand,
  selectedLatex: string,
): string {
  if (!selectedLatex) return command.insertTemplate;

  switch (command.id) {
    case "scripts":
      return selectedLatex + "_{\\placeholder{}}^{\\placeholder{}}";
    case "lower-script":
      return selectedLatex + "_{\\placeholder{}}";
    case "upper-script":
      return selectedLatex + "^{\\placeholder{}}";
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
  const syncFrameSizeRef = useRef<(() => void) | null>(null);
  const lastSnapshotRef = useRef<ReturnType<typeof captureFieldSnapshot> | null>(null);
  const compositionStartRef = useRef<ReturnType<typeof captureFieldSnapshot> | null>(null);
  const propsRef = useRef(props);
  propsRef.current = props;

  useLayoutEffect(() => {
    const host = hostRef.current;
    if (!host) return;

    const lineId = propsRef.current.lineId;
    const field = new MathfieldElement();
    field.value = propsRef.current.latex;
    field.className = "visual-mathfield";
    field.smartMode = false;
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
    field.style.fontSize = formulaFontSize(propsRef.current.zoom) + "px";

    let resizeFrame = 0;
    const syncFrameSize = () => {
      window.cancelAnimationFrame(resizeFrame);
      resizeFrame = window.requestAnimationFrame(() => {
        const fontSize = formulaFontSize(propsRef.current.zoom);
        const content = field.shadowRoot?.querySelector<HTMLElement>(
          '[part="content"]',
        );
        const atomRects = content
          ? Array.from(
              content.querySelectorAll<HTMLElement>("[data-atom-id]"),
            )
              .map((atom) => atom.getBoundingClientRect())
              .filter((rect) => rect.height > 0 && rect.width >= 0)
          : [];
        const formulaHeight = atomRects.length
          ? Math.max(...atomRects.map((rect) => rect.bottom)) -
            Math.min(...atomRects.map((rect) => rect.top))
          : fontSize;
        const hasTallStructure = tallFormulaPattern.test(field.value);
        const baseHeight = hasTallStructure
          ? Math.max(36, fontSize * 1.34 + 16)
          : Math.max(30, fontSize * 1.12 + 8);
        const verticalPadding = hasTallStructure
          ? Math.max(10, fontSize * 0.44)
          : Math.max(8, fontSize * 0.26);
        const nextHeight = Math.ceil(
          Math.max(baseHeight, formulaHeight + verticalPadding),
        );

        field.classList.toggle("is-simple-formula", !hasTallStructure);
        field.style.height = nextHeight + "px";
        field.style.minHeight = nextHeight + "px";
        host.closest<HTMLElement>(".formula-line")?.style.setProperty(
          "--formula-row-height",
          nextHeight + "px",
        );
      });
    };
    syncFrameSizeRef.current = syncFrameSize;

    let composing = false;
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
      propsRef.current.onCommitPending();
      composing = true;
      compositionStartRef.current =
        lastSnapshotRef.current ?? captureFieldSnapshot(field);
    };
    const handleCompositionEnd = () => {
      composing = false;
      const before =
        compositionStartRef.current ??
        lastSnapshotRef.current ??
        captureFieldSnapshot(field);
      const after = captureFieldSnapshot(field);
      compositionStartRef.current = null;
      emitEdit(before, after, "composition", "keyboard");
      syncFrameSize();
    };
    const handleInput = (event: Event) => {
      if (composing) return;
      const before = lastSnapshotRef.current ?? captureFieldSnapshot(field);
      const after = captureFieldSnapshot(field);
      const inputType =
        event instanceof InputEvent ? event.inputType || "insertText" : "insertText";
      emitEdit(
        before,
        after,
        inferEditKind(inputType, before.selection),
        inferEditSource(inputType),
      );
      syncFrameSize();
    };
    const handleSelectionChange = () => {
      if (composing || !lastSnapshotRef.current) return;
      lastSnapshotRef.current = {
        ...lastSnapshotRef.current,
        selection: captureSelection(field),
      };
    };
    const handleFocus = () => {
      propsRef.current.onFocus(propsRef.current.index, field);
      lastSnapshotRef.current = captureFieldSnapshot(field);
    };
    const handleBlur = () => {
      propsRef.current.onCommitPending();
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.isComposing || composing) return;
      propsRef.current.onKeyDown(propsRef.current.index, event, field);
    };
    const handlePaste = (event: ClipboardEvent) => {
      const clipboard = event.clipboardData;
      if (!clipboard || !propsRef.current.onPasteImage) return;

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
        lineId,
        ranges: selection.ranges.map(
          ([start, end]) => [start, end] as [number, number],
        ),
        direction: selection.direction ?? "none",
      });
    };
    const handlePointerDown = (event: PointerEvent) => {
      if (event.button !== 0) return;
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

      const hasVisibleFormula = Boolean(field.value.trim()) && contentBounds;
      const clickedInRightBlankArea = hasVisibleFormula
        ? event.clientX > contentBounds.right + 6
        : true;

      // Preserve MathLive's native hit testing for every click on or between
      // rendered formula atoms. Only the unused row area to the right maps to
      // the mathematical end of the line.
      if (!clickedInRightBlankArea) return;

      event.preventDefault();
      event.stopPropagation();

      const end = field.lastOffset;
      field.focus();
      field.selection = {
        ranges: [[end, end]],
        direction: "none",
      };
      field.position = end;
      field.shadowRoot
        ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
        ?.focus({ preventScroll: true });
      propsRef.current.onFocus(propsRef.current.index, field);
    };
    host.replaceChildren(field);
    // MathLive mounts a pre-filled field with the whole formula selected.
    // Collapse that implicit selection so toolbar commands insert at the end
    // instead of unexpectedly replacing/wrapping the entire line.
    field.position = field.lastOffset;
    field.resetUndo();
    lastSnapshotRef.current = captureFieldSnapshot(field);
    fieldRef.current = field;
    propsRef.current.register(lineId, field);
    field.addEventListener("compositionstart", handleCompositionStart);
    field.addEventListener("compositionend", handleCompositionEnd);
    field.addEventListener("input", handleInput);
    field.addEventListener("selection-change", handleSelectionChange);
    field.addEventListener("focus", handleFocus);
    field.addEventListener("blur", handleBlur);
    field.addEventListener("keydown", handleKeyDown, true);
    field.addEventListener("paste", handlePaste, true);
    host.addEventListener("pointerdown", handlePointerDown, true);
    const content = field.shadowRoot?.querySelector<HTMLElement>('[part="content"]');
    const resizeObserver = content ? new ResizeObserver(syncFrameSize) : null;
    if (content) resizeObserver?.observe(content);
    syncFrameSize();

    return () => {
      window.cancelAnimationFrame(resizeFrame);
      resizeObserver?.disconnect();
      syncFrameSizeRef.current = null;
      field.removeEventListener("compositionstart", handleCompositionStart);
      field.removeEventListener("compositionend", handleCompositionEnd);
      field.removeEventListener("input", handleInput);
      field.removeEventListener("selection-change", handleSelectionChange);
      field.removeEventListener("focus", handleFocus);
      field.removeEventListener("blur", handleBlur);
      field.removeEventListener("keydown", handleKeyDown, true);
      field.removeEventListener("paste", handlePaste, true);
      host.removeEventListener("pointerdown", handlePointerDown, true);
      host.closest<HTMLElement>(".formula-line")?.style.removeProperty(
        "--formula-row-height",
      );
      propsRef.current.register(lineId, null);
      fieldRef.current = null;
      lastSnapshotRef.current = null;
      compositionStartRef.current = null;
      host.replaceChildren();
    };
  }, []);

  useEffect(() => {
    const field = fieldRef.current;
    if (!field) return;

    // 本地输入仅因中文规范化而与 store 等值时，不重建 MathLive 模型；
    // 只更新事务基准，保留当前光标、选区和删除键内部状态。
    if (normalizeChineseLatex(field.value) === props.latex) {
      lastSnapshotRef.current = {
        latex: props.latex,
        selection: captureSelection(field),
      };
      return;
    }

    field.setValue(props.latex, {
      mode: "math",
      format: "latex",
      insertionMode: "replaceAll",
      selectionMode: "after",
      silenceNotifications: true,
    });
    field.resetUndo();
    lastSnapshotRef.current = captureFieldSnapshot(field);
    syncFrameSizeRef.current?.();
  }, [props.latex]);

  useEffect(() => {
    if (fieldRef.current) {
      fieldRef.current.style.fontSize = formulaFontSize(props.zoom) + "px";
      syncFrameSizeRef.current?.();
    }
  }, [props.zoom]);

  useEffect(() => {
    const field = fieldRef.current;
    if (!field) return;
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
    { lines, activeLineId, zoom, onPasteImage, overlay },
    ref,
  ) {
    const surfaceRef = useRef<HTMLDivElement>(null);
    const fieldRefs = useRef(new Map<string, MathfieldElement>());
    const linesRef = useRef(lines);
    const activeIndexRef = useRef(0);
    const activeLineIdRef = useRef<string | null>(activeLineId);
    const focusRequestRef = useRef(0);
    const suppressedHistoryLineIdRef = useRef<string | null>(null);
    const pendingFocusRef = useRef<{
      lineId: string;
      latex: string | null;
      selection: MathSelectionSnapshot | null;
      moveToEnd: boolean;
    } | null>(null);
    const [activeIndex, setActiveIndex] = useState(() =>
      Math.max(0, lines.findIndex((line) => line.id === activeLineId)),
    );
    const [query, setQuery] = useState("");
    const [selectedIndex, setSelectedIndex] = useState(0);
    const [popupPosition, setPopupPosition] = useState({ left: 72, top: 132 });
    const selectedIndexRef = useRef(0);
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
    const isEn = language === "en";

    linesRef.current = lines;
    const resolvedActiveLineId =
      lines.find((line) => line.id === activeLineId)?.id ?? lines[0]?.id ?? null;
    activeLineIdRef.current = resolvedActiveLineId;

    const suggestions = useMemo(
      () =>
        query
          ? searchCommands(query, usage, personalize, suggestionCount)
          : [],
      [query, usage, personalize, suggestionCount],
    );

    const selectSuggestionIndex = (index: number) => {
      selectedIndexRef.current = index;
      setSelectedIndex(index);
    };

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
      field.focus();
      field.shadowRoot
        ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
        ?.focus({ preventScroll: true });

      if (selection) {
        const clamped = clampSelection(selection, field.lastOffset);
        field.selection = clamped;
        const [start, end] = clamped.ranges[0] ?? [field.lastOffset, field.lastOffset];
        field.position = clamped.direction === "backward" ? start : end;
      } else if (moveToEnd) {
        const end = field.lastOffset;
        field.selection = {
          ranges: [[end, end]],
          direction: "none",
        };
        field.position = end;
      }
      return true;
    };

    const focusLine = (
      lineId: string,
      options: {
        latex?: string | null;
        selection?: MathSelectionSnapshot | null;
        moveToEnd?: boolean;
      } = {},
      remainingAttempts = 10,
      requestId = ++focusRequestRef.current,
    ) => {
      const index = linesRef.current.findIndex((line) => line.id === lineId);
      if (index < 0) return false;

      const expectedLatex = options.latex ?? null;
      const selection = options.selection ?? null;
      const moveToEnd = options.moveToEnd ?? false;
      activeIndexRef.current = index;
      activeLineIdRef.current = lineId;
      setActiveIndex(index);
      useEditorStore.getState().setActiveLineId(lineId);
      pendingFocusRef.current = {
        lineId,
        latex: expectedLatex,
        selection,
        moveToEnd,
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
        window.requestAnimationFrame(() => {
          if (!apply()) return;
          window.setTimeout(apply, 80);
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
    ) => {
      if (field) fieldRefs.current.set(lineId, field);
      else fieldRefs.current.delete(lineId);

      const pending = pendingFocusRef.current;
      if (field && pending?.lineId === lineId) {
        focusLine(lineId, {
          latex: pending.latex,
          selection: pending.selection,
          moveToEnd: pending.moveToEnd,
        });
      }
    };

    const updatePopupPosition = (field: MathfieldElement) => {
      const fieldWithCaret = field as MathfieldElement & {
        getCaretPoint?: () => { x: number; y: number; height?: number };
      };
      const surface = surfaceRef.current;
      if (!surface) return;
      const surfaceRect = surface.getBoundingClientRect();
      const popupWidth = Math.min(440, Math.max(280, surfaceRect.width - 24));
      const clampLeft = (left: number) =>
        Math.max(12, Math.min(surfaceRect.width - popupWidth - 12, left));
      const caret = fieldWithCaret.getCaretPoint?.();

      if (caret) {
        setPopupPosition({
          left: clampLeft(caret.x - surfaceRect.left),
          top: Math.max(
            64,
            caret.y - surfaceRect.top + (caret.height ?? 28) + 8,
          ),
        });
      } else {
        const fieldRect = field.getBoundingClientRect();
        setPopupPosition({
          left: clampLeft(fieldRect.left - surfaceRect.left + 48),
          top: fieldRect.bottom - surfaceRect.top + 8,
        });
      }
    };

    const refreshSuggestionQuery = (
      lineId: string,
      field: MathfieldElement,
      normalized: string,
    ) => {
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

      const match = normalized.match(trailingCommand);
      if (match) {
        setQuery("\\" + match[1]);
        selectSuggestionIndex(0);
        requestAnimationFrame(() => updatePopupPosition(field));
      } else {
        setQuery("");
      }
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
      setActiveIndex(index);
      useEditorStore.getState().setActiveLineId(lineId);
    };

    const handleFieldEdit = (
      edit: FormulaFieldEdit,
      field: MathfieldElement,
    ) => {
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
    ) => {
      historyManager.commitPendingTransaction();
      const state = useEditorStore.getState();
      const currentLine = state.lines.find((line) => line.id === lineId);
      if (!currentLine) return false;

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

      const after = captureFieldSnapshot(field);
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
      historyManager.push(entry);
      return true;
    };

    const applyDeferredDiscreteFormulaMutation = async (
      lineId: string,
      field: MathfieldElement,
      source: FormulaEditSource,
      mutate: () => boolean,
    ) => {
      historyManager.commitPendingTransaction();
      const state = useEditorStore.getState();
      if (!state.lines.some((line) => line.id === lineId)) return false;

      const before = captureFieldSnapshot(field);
      const beforeActiveLineId = state.activeLineId;
      suppressedHistoryLineIdRef.current = lineId;
      let changed = false;
      let after = before;
      try {
        changed = mutate();
        if (!changed) return false;

        const startedAt = performance.now();
        while (performance.now() - startedAt < 1000) {
          await new Promise<void>((resolve) =>
            window.requestAnimationFrame(() => resolve()),
          );
          after = captureFieldSnapshot(field);
          if (after.latex !== before.latex) break;
        }
      } finally {
        suppressedHistoryLineIdRef.current = null;
      }

      if (before.latex === after.latex) {
        field.resetUndo();
        return false;
      }

      state.replaceFormulaLine(lineId, after.latex);
      state.setActiveLineId(lineId);
      linesRef.current = useEditorStore.getState().lines;
      setActiveLine(lineId);
      field.resetUndo();
      historyManager.push({
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

      const originalSelection = captureSelection(field);
      const queryRange = activeQuery
        ? findTrailingCommandRange(field, activeQuery)
        : null;
      if (activeQuery && !queryRange) {
        setQuery("");
        selectSuggestionIndex(0);
        return;
      }

      const selectedLatex = activeQuery || field.selectionIsCollapsed
        ? ""
        : field.getValue(field.selection);
      const insertionTemplate = activeQuery
        ? command.insertTemplate
        : templateForSelection(command, selectedLatex);
      const historySource: FormulaEditSource =
        source === "candidate" ? "candidate" : "toolbar";

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
        selectSuggestionIndex(0);
        field.focus();
        return;
      }

      const tryInsert = () => {
        field.selection = originalSelection;
        const inserted = applyDiscreteFormulaMutation(
          targetLineId,
          field,
          historySource,
          () => {
            if (queryRange) {
              field.selection = {
                ranges: [queryRange],
                direction: "forward",
              };
            }
            return field.insert(insertionTemplate, {
              mode: "math",
              format: "latex",
              insertionMode: "replaceSelection",
              selectionMode: "placeholder",
              focus: true,
              scrollIntoView: false,
            });
          },
        );
        if (!inserted) field.selection = originalSelection;
        return inserted;
      };

      const finishInsertion = () => {
        const normalizedValue = normalizeChineseLatex(field.value);
        suppressedSuggestionRef.current = {
          lineId: targetLineId,
          value: normalizedValue,
        };
        recordCommand(command.id, activeQuery, source);
        setQuery("");
        selectSuggestionIndex(0);
        focusLine(targetLineId, {
          latex: normalizedValue,
          selection: captureSelection(field),
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

    const addLineAfter = (index: number) => {
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
      const line = createFormulaLine("");
      const afterSelection: MathSelectionSnapshot = {
        ranges: [[0, 0]],
        direction: "none",
      };

      flushSync(() => {
        state.insertFormulaLine(line, nextIndex);
        useEditorStore.getState().setActiveLineId(line.id);
      });
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
      });
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

      removedField?.blur();
      flushSync(() => {
        state.removeFormulaLine(removedLine.id);
        useEditorStore.getState().setActiveLineId(targetLine.id);
      });
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
      focusLine(targetLine.id, {
        latex: targetLine.latex,
        selection: afterSelection,
      });
    };

    const handleKeyDown = (
      index: number,
      lineId: string,
      event: KeyboardEvent,
      field: MathfieldElement,
    ) => {
      setActiveLine(lineId);

      const shortcutKey = event.key.toLocaleLowerCase();
      const primaryModifier = (event.metaKey || event.ctrlKey) && !event.altKey;
      const requestsUndo = primaryModifier && shortcutKey === "z" && !event.shiftKey;
      const requestsRedo =
        primaryModifier &&
        ((shortcutKey === "z" && event.shiftKey) ||
          (shortcutKey === "y" && !event.shiftKey));
      if (requestsUndo || requestsRedo) {
        event.preventDefault();
        event.stopPropagation();
        if (requestsRedo) void historyManager.redo();
        else void historyManager.undo();
        return;
      }

      const liveMatch = field.value.match(trailingCommand);
      const liveQuery = liveMatch ? "\\" + liveMatch[1] : "";
      const candidateQuery = liveQuery || query;
      const liveSuggestions = candidateQuery
        ? searchCommands(
            candidateQuery,
            useEditorStore.getState().usage,
            useEditorStore.getState().personalize,
            useEditorStore.getState().suggestionCount,
          )
        : [];

      if (liveSuggestions.length) {
        if (event.key === "ArrowDown") {
          event.preventDefault();
          event.stopPropagation();
          selectSuggestionIndex(
            (selectedIndexRef.current + 1) % liveSuggestions.length,
          );
          return;
        }
        if (event.key === "ArrowUp") {
          event.preventDefault();
          event.stopPropagation();
          selectSuggestionIndex(
            (selectedIndexRef.current - 1 + liveSuggestions.length) % liveSuggestions.length,
          );
          return;
        }
        if (event.key === "Enter" || event.key === "Tab") {
          event.preventDefault();
          event.stopPropagation();
          const suggestionIndex = Math.min(
            selectedIndexRef.current,
            liveSuggestions.length - 1,
          );
          suppressedSuggestionRef.current = {
            lineId,
            value: normalizeChineseLatex(field.value),
          };
          queryRef.current = "";
          setQuery("");
          insertCommand(
            liveSuggestions[suggestionIndex],
            "candidate",
            candidateQuery,
          );
          return;
        }
        if (event.key === "Escape") {
          event.preventDefault();
          event.stopPropagation();
          suppressedSuggestionRef.current = {
            lineId,
            value: normalizeChineseLatex(field.value),
          };
          setQuery("");
          return;
        }
      }

      const nativeRecommendationVisible =
        document
          .getElementById("mathlive-suggestion-popover")
          ?.classList.contains("is-visible") ?? false;
      if (!liveSuggestions.length && nativeRecommendationVisible) {
        if (event.key === "ArrowDown" || event.key === "ArrowUp") {
          event.preventDefault();
          event.stopPropagation();
          moveNativeSuggestionSelection(event.key === "ArrowDown" ? 1 : -1);
          return;
        }
        if (event.key === "Enter" || event.key === "Tab") {
          event.preventDefault();
          event.stopPropagation();
          void applyDeferredDiscreteFormulaMutation(
            lineId,
            field,
            "candidate",
            commitNativeSuggestion,
          ).then((committed) => {
            if (!committed) return;
            setQuery("");
            selectSuggestionIndex(0);
            focusLine(lineId, {
              latex: normalizeChineseLatex(field.value),
              selection: captureSelection(field),
            });
          });
          return;
        }
        if (event.key === "Escape") {
          // Escape only dismisses the panel and does not rebuild its rows, so
          // MathLive can safely handle it itself.
          return;
        }
      }

      if (event.key === "Enter") {
        event.preventDefault();
        event.stopPropagation();
        addLineAfter(index);
        return;
      }

      if (event.key === "Backspace" || event.key === "Delete") {
        const visibleLatex = field
          .getValue("latex-without-placeholders")
          .trim();

        if (visibleLatex === "" && linesRef.current.length > 1) {
          event.preventDefault();
          event.stopPropagation();
          removeEmptyLine(index);
          return;
        }

        // Preserve MathLive's native word/line deletion shortcuts such as
        // Option+Backspace and Command+Backspace on macOS.
        if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) {
          return;
        }

        event.preventDefault();
        event.stopPropagation();

        const state = useEditorStore.getState();
        const currentLine = state.lines.find((line) => line.id === lineId);
        if (!currentLine) return;
        const before = captureFieldSnapshot(field);
        const beforePosition = field.position;
        const command =
          event.key === "Backspace" ? "deleteBackward" : "deleteForward";

        suppressedHistoryLineIdRef.current = lineId;
        try {
          field.executeCommand(command);
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
        event.key.startsWith("Arrow") ||
        event.key === "Home" ||
        event.key === "End"
      ) {
        historyManager.commitPendingTransaction();
      }
    };

    const normalizeInsertedLatex = (latex: string) =>
      latex
        .replace(/\r\n?/g, "\n")
        .split("\n")
        .map((line) => line.trim())
        .filter(Boolean)
        .join("\\quad ");

    const getSelectionMap = (): Record<string, MathSelectionSnapshot> =>
      Object.fromEntries(
        linesRef.current.flatMap((line) => {
          const field = fieldRefs.current.get(line.id);
          return field?.isConnected
            ? [[line.id, captureSelection(field)] as const]
            : [];
        }),
      );

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
            window.requestAnimationFrame(() => {
              applyFocusState(lineId, latex, selection, !selection, requestId);
              window.setTimeout(
                () =>
                  applyFocusState(
                    lineId,
                    latex,
                    selection,
                    !selection,
                    requestId,
                  ),
                80,
              );
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

    const insertLatex = (
      latex: string,
      source: FormulaEditSource = "ocr",
    ) => {
      const value = normalizeInsertedLatex(latex);
      if (!value) return;
      const target = resolveTargetField();
      if (!target) return;
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
      selectSuggestionIndex(0);
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
      const value = normalizeInsertedLatex(latex);
      if (!value) return false;
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
      selectSuggestionIndex(0);
      focusLine(target.lineId, {
        latex: normalizeChineseLatex(field.value),
        selection: captureSelection(field),
      });
      return true;
    };

    const appendLatex = (
      latex: string,
      _source: FormulaEditSource = "ocr",
    ) => {
      const values = latex
        .replace(/\r\n?/g, "\n")
        .split("\n")
        .map((line) => normalizeChineseLatex(line.trim()))
        .filter(Boolean);
      if (!values.length) return;

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
      const after = {
        title: before.title,
        lines: nextLines.map((line) => ({ ...line })),
        activeLineId: lastLine.id,
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
      focus: () => {
        const lineId = activeLineIdRef.current ?? linesRef.current[0]?.id;
        if (lineId) focusLine(lineId);
      },
      addLine: () => addLineAfter(linesRef.current.length - 1),
      commitPendingTransaction: () => historyManager.commitPendingTransaction(),
      getSelectionMap,
      restoreSelection,
    }));

    useEffect(() => {
      selectedIndexRef.current = selectedIndex;
      if (selectedIndex >= suggestions.length) selectSuggestionIndex(0);
    }, [suggestions.length, selectedIndex]);

    useEffect(() => {
      queryRef.current = query;
    }, [query]);

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

    return (
      <div
        ref={surfaceRef}
        className="editor-surface multi-line-editor"
        data-command-query={query}
        data-active-line-id={activeLineIdRef.current ?? ""}
      >
        {overlay}
        <div className="mathfield-stack">
          {lines.map((line, index) => {
            const lineId = line.id;
            return (
            <div
              className={"formula-line " + (lineId === activeLineIdRef.current ? "is-active" : "")}
              data-line-id={lineId}
              key={lineId}
            >
              <span className="formula-line-number">
                {String(index + 1).padStart(2, "0")}
              </span>
              <FormulaField
                lineId={lineId}
                index={index}
                latex={line.latex}
                zoom={zoom}
                language={language}
                register={registerField}
                onEdit={handleFieldEdit}
                onCommitPending={() => historyManager.commitPendingTransaction()}
                onFocus={(_lineIndex, field) => {
                  setActiveLine(lineId);
                  refreshSuggestionQuery(
                    lineId,
                    field,
                    normalizeChineseLatex(field.value),
                  );
                }}
                onKeyDown={(lineIndex, event, field) =>
                  handleKeyDown(lineIndex, lineId, event, field)
                }
                onPasteImage={onPasteImage}
              />
            </div>
            );
          })}

          <button
            type="button"
            className="add-formula-line"
            onClick={() => addLineAfter(linesRef.current.length - 1)}
            aria-label={isEn ? "Add formula line" : "添加公式行"}
            title={isEn ? "Add formula line · Enter" : "添加公式行 · Enter"}
          >
            <Plus size={15} />
          </button>
        </div>
        <CommandSuggestionPopup
          suggestions={suggestions}
          selectedIndex={selectedIndex}
          position={popupPosition}
          usage={usage}
          onHighlight={selectSuggestionIndex}
          onCommit={(command) =>
            insertCommand(command, "candidate", queryRef.current)
          }
        />
      </div>
    );
  },
);
