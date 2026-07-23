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
import { MathfieldElement, convertLatexToMarkup } from "mathlive";
import { flushSync } from "react-dom";
import { Plus } from "lucide-react";
import type {
  CommandSource,
  CommandUsage,
  LatexCommand,
} from "../types/command";
import type {
  FormulaLine,
  InputBehaviorSettingKey,
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
import { searchCommands } from "../autocomplete/CommandSearchEngine";
import { commandRegistry } from "../autocomplete/commandRegistry";
import { CommandSuggestionPopup } from "../autocomplete/CommandSuggestionPopup";
import {
  createFormulaLine,
  useEditorStore,
} from "../stores/editorStore";
import {
  normalizeChineseLatex,
  normalizeChineseTextLatex,
} from "./normalizeChineseLatex";

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

interface WrapperInputController {
  begin: (
    command: LatexCommand,
    activeQuery: string,
    source: FormulaEditSource,
  ) => boolean;
}

interface FormulaFieldProps {
  lineId: string;
  index: number;
  latex: string;
  zoom: number;
  language: "cn" | "en";
  autoPairDelimiters: boolean;
  inputBehavior: InputBehaviorSettings;
  register: (lineId: string, field: MathfieldElement | null) => void;
  registerWrapperInputController: (
    lineId: string,
    controller: WrapperInputController | null,
  ) => void;
  onEdit: (edit: FormulaFieldEdit, field: MathfieldElement) => void;
  onInputActivity: (field: MathfieldElement) => void;
  onFocus: (index: number, field: MathfieldElement) => void;
  onCommitPending: () => void;
  onKeyDown: (index: number, event: KeyboardEvent, field: MathfieldElement) => void;
  onPasteImage?: (file: File, target: MathEditorInsertionTarget) => void;
}

const trailingCommand = /\\([\p{L}]*)$/u;

function rawLatexInput(field: MathfieldElement) {
  return Array.from(
    field.shadowRoot?.querySelectorAll<HTMLElement>(".ML__raw-latex") ?? [],
  )
    .map((node) => node.textContent ?? "")
    .join("");
}

function trailingCommandQuery(
  field: MathfieldElement,
  normalizedValue = field.value,
) {
  for (const source of [normalizedValue, field.value, rawLatexInput(field)]) {
    const match = source.match(trailingCommand);
    if (match) return "\\" + match[1];
  }
  return "";
}

function fieldValueIncludingRawQuery(
  field: MathfieldElement,
  activeQuery: string,
) {
  const value = field.value;
  if (value.trimEnd().endsWith(activeQuery)) return value;
  const rawInput = rawLatexInput(field);
  if (rawInput.trimEnd().endsWith(activeQuery)) return value + rawInput;
  return value;
}

const structuredSuggestionCommands = new Set([
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
const wrapperCommandPreviews = new Map<string, string>([
  ["\\mathbb", "\\mathbb{ABC}"],
  ["\\mathbf", "\\mathbf{ABC}"],
  ["\\mathit", "\\mathit{ABC}"],
  ["\\mathop", "\\mathop{f(x)}"],
  ["\\mathrm", "\\mathrm{ABC}"],
  ["\\mathsf", "\\mathsf{ABC}"],
  ["\\mathtt", "\\mathtt{ABC}"],
  ["\\mathcal", "\\mathcal{ABC}"],
  ["\\mathscr", "\\mathscr{gG}"],
  ["\\mathfrak", "\\mathfrak{ABC}"],
  ["\\boldsymbol", "\\boldsymbol{\\alpha A}"],
  ["\\mathnormal", "\\mathnormal{ABC}"],
]);

function visibleCommandSuggestions(
  rawQuery: string,
  usage: Record<string, CommandUsage>,
  personalize: boolean,
  limit: number,
  settings: InputBehaviorSettings,
) {
  return searchCommands(
    rawQuery,
    usage,
    personalize,
    commandRegistry.length,
  )
    .filter((command) =>
      structuredSuggestionCommands.has(command.command)
        ? settings.showStructuredCommandSuggestions
        : settings.showOtherCommandSuggestions,
    )
    .slice(0, limit);
}

function exactWrapperCommand(rawQuery: string) {
  const normalizedQuery = rawQuery.trim();
  if (!wrapperCommandPreviews.has(normalizedQuery)) return null;
  return (
    commandRegistry.find((command) => command.command === normalizedQuery) ?? null
  );
}

function decorateNativeSuggestionPreviews() {
  const panel = document.getElementById("mathlive-suggestion-popover");
  if (!panel) return;

  panel.querySelectorAll<HTMLElement>("li[data-command]").forEach((item) => {
    const command = item.dataset.command ?? "";
    const previewLatex = wrapperCommandPreviews.get(command);
    const preview = item.querySelector<HTMLElement>(".ML__popover__command");
    if (!previewLatex || !preview) return;
    if (preview.dataset.visualtexPreview === previewLatex) return;

    preview.innerHTML = convertLatexToMarkup(previewLatex, {
      defaultMode: "math",
    });
    preview.dataset.visualtexPreview = previewLatex;
    preview.setAttribute("aria-label", command);
    item.classList.add("has-visualtex-command-preview");
  });
}

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

function getScriptCaretRegion(field: MathfieldElement): "upper" | "lower" | null {
  const markers = Array.from(
    field.shadowRoot?.querySelectorAll<HTMLElement>(
      ".ML__placeholder-selected, .ML__selected, .ML__caret",
    ) ?? [],
  );
  const caret = markers.find((marker) =>
    marker.closest(".ML__msubsup, .ML__op-group"),
  );
  // Side scripts use ML__msubsup; large operators such as sum/product render
  // their over/under limits directly inside ML__op-group.
  const script = caret?.closest<HTMLElement>(".ML__msubsup, .ML__op-group");
  if (!caret || !script) return null;

  // MathLive's caret itself can span most of a script or large-operator box.
  // Its immediate row wrapper has the actual upper/lower branch geometry.
  const caretBounds = (caret.parentElement ?? caret).getBoundingClientRect();
  if (!caretBounds.height) return null;

  if (script.classList.contains("ML__msubsup")) {
    const table = script.querySelector<HTMLElement>(":scope > .ML__vlist-t");
    const branchList = table?.querySelector<HTMLElement>(
      ":scope > .ML__vlist-r > .ML__vlist",
    );
    const branches = Array.from(branchList?.children ?? []).filter(
      (branch): branch is HTMLElement => branch instanceof HTMLElement,
    );
    const activeBranchIndex = branches.findIndex((branch) =>
      branch.contains(caret),
    );

    if (activeBranchIndex >= 0) {
      // MathLive emits subscript first and superscript second when both exist.
      if (branches.length > 1) {
        return activeBranchIndex === 0 ? "lower" : "upper";
      }

      // With one branch, ML__vlist-t2 is emitted only for a branch extending
      // below the baseline (a subscript). A lone superscript uses ML__vlist-t.
      // This structural distinction avoids the old midpoint bug where a lone
      // superscript was incorrectly governed by the subscript setting.
      return table?.classList.contains("ML__vlist-t2")
        ? "lower"
        : "upper";
    }
  }

  // Large operators keep the operator glyph inside ML__op-group, so their
  // over/under limits can be classified relative to that stable nucleus.
  const nucleus = script.querySelector<HTMLElement>(".ML__op-symbol");
  const nucleusBounds = nucleus?.getBoundingClientRect();
  const caretCenter = caretBounds.top + caretBounds.height / 2;
  if (nucleusBounds?.height) {
    return caretCenter < nucleusBounds.top + nucleusBounds.height / 2
      ? "upper"
      : "lower";
  }

  const scriptBounds = script.getBoundingClientRect();
  if (!scriptBounds.height) return null;
  return caretCenter < scriptBounds.top + scriptBounds.height / 2
    ? "upper"
    : "lower";
}

const accentContainerPattern =
  /\\(?:acute|grave|dot|ddot|dddot|ddddot|tilde|bar|breve|check|hat|vec|widehat|widetilde|overline|overrightarrow|overleftarrow|overleftrightarrow)\s*\{/;

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

  const scriptRegion = getScriptCaretRegion(field);
  if (scriptRegion === "upper") return "autoExitSuperscript";
  if (scriptRegion === "lower") return "autoExitSubscript";
  return null;
}

function isSingleDirectInput(event: InputEvent, field: MathfieldElement) {
  if (event.isComposing || field.mode === "latex") return false;
  if (event.inputType !== "insertText") return false;
  const data = event.data ?? "";
  if (data === "\\" || Array.from(data).length !== 1) return false;
  return true;
}

const placeholderToken = "\\placeholder{}";
const accentPlaceholderPrefixPattern =
  /\\(?:acute|grave|dot|ddot|dddot|ddddot|tilde|bar|breve|check|hat|vec|widehat|widetilde|overline|overrightarrow|overleftarrow|overleftrightarrow)\s*\{$/;

function autoExitSettingForPlaceholderAt(
  latex: string,
  placeholderIndex: number,
): InputBehaviorSettingKey | null {
  const prefix = latex.slice(0, placeholderIndex);
  if (/\^\s*\{$/.test(prefix)) return "autoExitSuperscript";
  if (/_\s*\{$/.test(prefix)) return "autoExitSubscript";
  if (accentPlaceholderPrefixPattern.test(prefix)) return "autoExitAccent";
  return null;
}

function inferPlaceholderAutoExitSetting(
  beforeLatex: string,
  afterLatex: string,
): InputBehaviorSettingKey | null {
  let searchStart = 0;
  while (searchStart < beforeLatex.length) {
    const placeholderIndex = beforeLatex.indexOf(placeholderToken, searchStart);
    if (placeholderIndex < 0) return null;
    const prefix = beforeLatex.slice(0, placeholderIndex);
    const suffix = beforeLatex.slice(placeholderIndex + placeholderToken.length);
    if (
      afterLatex.startsWith(prefix) &&
      afterLatex.endsWith(suffix) &&
      afterLatex.length > prefix.length + suffix.length
    ) {
      return autoExitSettingForPlaceholderAt(beforeLatex, placeholderIndex);
    }
    searchStart = placeholderIndex + placeholderToken.length;
  }
  return null;
}

function moveCaretThroughEnabledAutoExitContainers(
  field: MathfieldElement,
  settings: InputBehaviorSettings,
  initialSetting: InputBehaviorSettingKey | null = null,
) {
  let moved = false;
  let forcedSetting = initialSetting;
  for (let attempt = 0; attempt < 4; attempt += 1) {
    const setting = forcedSetting ?? getCaretAutoExitSetting(field);
    forcedSetting = null;
    if (!setting || !settings[setting]) break;

    const collapsePosition = Math.max(
      field.position,
      ...field.selection.ranges.flatMap(([start, end]) => [start, end]),
    );
    field.position = collapsePosition;
    const previousPosition = field.position;
    const changed = field.executeCommand(
      setting === "autoExitAccent" ? "moveToNextChar" : "moveAfterParent",
    );
    if (!changed && field.position === previousPosition) break;
    moved = true;
  }
  return moved;
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
  const nativeInputPendingRef = useRef(false);
  const propsRef = useRef(props);
  propsRef.current = props;

  useLayoutEffect(() => {
    const host = hostRef.current;
    if (!host) return;

    const lineId = propsRef.current.lineId;
    const field = new MathfieldElement();
    const wrapperMeasure = document.createElement("span");
    wrapperMeasure.className = "pending-wrapper-measure";
    MathfieldElement.locale = propsRef.current.language === "en" ? "en" : "zh-cn";
    field.value = propsRef.current.latex;
    field.className = "visual-mathfield";
    field.smartMode = false;
    field.smartFence = propsRef.current.autoPairDelimiters;
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
    let resizeTimer = 0;
    let resizePass = 0;
    const measureFrameSize = () => {
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

      resizePass += 1;
      if (resizePass < 4) {
        resizeTimer = window.setTimeout(() => {
          resizeFrame = window.requestAnimationFrame(measureFrameSize);
        }, resizePass * 50);
      }
    };
    const syncFrameSize = () => {
      window.cancelAnimationFrame(resizeFrame);
      window.clearTimeout(resizeTimer);
      resizePass = 0;
      resizeFrame = window.requestAnimationFrame(measureFrameSize);
    };
    syncFrameSizeRef.current = syncFrameSize;

    let composing = false;
    let pendingAutoExitSetting: InputBehaviorSettingKey | null = null;
    let pendingWrapperConfirmationKey = "";
    let renderingPendingWrapperInput = false;
    let pendingWrapperInput: {
      command: string;
      prefix: string;
      content: string;
      suffix: string;
    } | null = null;
    let wrapperPlaceholderFrame = 0;

    const clearPendingWrapperPlaceholderPosition = () => {
      host.style.removeProperty("--pending-wrapper-left");
      host.style.removeProperty("--pending-wrapper-top");
      host.style.removeProperty("--pending-wrapper-width");
      host.style.removeProperty("--pending-wrapper-height");
    };
    const schedulePendingWrapperPlaceholderPosition = () => {
      window.cancelAnimationFrame(wrapperPlaceholderFrame);
      if (!host.classList.contains("has-pending-wrapper-placeholder")) {
        clearPendingWrapperPlaceholderPosition();
        return;
      }
      wrapperPlaceholderFrame = window.requestAnimationFrame(() => {
        if (!field.isConnected) return;
        const hostBounds = host.getBoundingClientRect();
        let wrapperWidth = 18;
        let wrapperHeight = 29;
        if (pendingWrapperInput?.content) {
          wrapperMeasure.style.fontSize = field.style.fontSize;
          wrapperMeasure.innerHTML = convertLatexToMarkup(
            `${pendingWrapperInput.command}{${pendingWrapperInput.content}}`,
            { defaultMode: "math" },
          );
          const measureBounds = wrapperMeasure.getBoundingClientRect();
          wrapperWidth = Math.max(24, measureBounds.width + 12);
          wrapperHeight = Math.max(
            29,
            Math.min(72, measureBounds.height + 8),
          );
        } else {
          wrapperMeasure.replaceChildren();
        }

        const caret = field.shadowRoot?.querySelector<HTMLElement>(".ML__caret");
        const caretBounds = caret?.getBoundingClientRect();
        if (!caretBounds || caretBounds.height <= 0) {
          clearPendingWrapperPlaceholderPosition();
          return;
        }
        const caretCenterY =
          caretBounds.top - hostBounds.top + caretBounds.height / 2;
        const caretX = caretBounds.left - hostBounds.left;
        const wrapperCenterX = pendingWrapperInput?.content
          ? caretX - wrapperWidth / 2
          : caretX;
        host.style.setProperty(
          "--pending-wrapper-left",
          `${wrapperCenterX}px`,
        );
        host.style.setProperty("--pending-wrapper-top", `${caretCenterY}px`);
        host.style.setProperty(
          "--pending-wrapper-width",
          `${wrapperWidth}px`,
        );
        host.style.setProperty(
          "--pending-wrapper-height",
          `${Math.max(wrapperHeight, Math.min(72, caretBounds.height + 4))}px`,
        );
      });
    };
    const syncPendingWrapperPlaceholder = () => {
      const showPlaceholder = Boolean(pendingWrapperInput);
      host.classList.toggle("has-pending-wrapper-placeholder", showPlaceholder);
      if (showPlaceholder && pendingWrapperInput) {
        host.dataset.pendingWrapperCommand = pendingWrapperInput.command;
      } else {
        delete host.dataset.pendingWrapperCommand;
      }
      schedulePendingWrapperPlaceholderPosition();
    };
    const clearPendingWrapperInput = () => {
      pendingWrapperInput = null;
      delete field.dataset.pendingWrapperCommand;
      syncPendingWrapperPlaceholder();
    };

    const pendingWrapperLatex = () => {
      if (!pendingWrapperInput) return field.value;
      return (
        pendingWrapperInput.prefix +
        pendingWrapperInput.command +
        "{" +
        pendingWrapperInput.content +
        "}" +
        pendingWrapperInput.suffix
      );
    };

    const renderPendingWrapperInput = () => {
      if (!pendingWrapperInput) return false;
      const marker = "\\placeholder{}";
      const cleanValue = pendingWrapperLatex();
      const markerValue =
        pendingWrapperInput.prefix +
        pendingWrapperInput.command +
        "{" +
        pendingWrapperInput.content +
        "}" +
        marker +
        pendingWrapperInput.suffix;

      renderingPendingWrapperInput = true;
      let markerRemoved = false;
      try {
        field.mode = "math";
        field.setValue(markerValue, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "placeholder",
          silenceNotifications: true,
        });
        markerRemoved = field.insert("", {
          mode: "math",
          format: "latex",
          insertionMode: "replaceSelection",
          selectionMode: "after",
          focus: true,
          scrollIntoView: false,
        });
        if (!markerRemoved && field.value.includes(marker)) {
          markerRemoved = field.executeCommand("deleteBackward");
        }
        if (field.value.includes(marker)) {
          field.setValue(cleanValue, {
            mode: "math",
            format: "latex",
            insertionMode: "replaceAll",
            selectionMode: "after",
            silenceNotifications: true,
          });
          field.position = field.lastOffset;
        }
      } finally {
        renderingPendingWrapperInput = false;
      }
      field.focus();
      field.shadowRoot
        ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
        ?.focus({ preventScroll: true });
      syncPendingWrapperPlaceholder();
      syncFrameSize();
      return markerRemoved || field.value === cleanValue;
    };

    const beginPendingWrapperInput = (
      command: LatexCommand,
      activeQuery: string,
      source: FormulaEditSource,
    ) => {
      if (!wrapperCommandPreviews.has(command.command)) return false;
      if (!field.selectionIsCollapsed && !activeQuery) return false;

      propsRef.current.onCommitPending();
      clearPendingWrapperInput();
      const before = captureFieldSnapshot(field);
      let prefix = "";
      let suffix = "";

      if (activeQuery) {
        const rawFieldValue = fieldValueIncludingRawQuery(field, activeQuery);
        const rawFieldValueWithoutTrailingSpace = rawFieldValue.trimEnd();
        if (!rawFieldValueWithoutTrailingSpace.endsWith(activeQuery)) return false;
        const queryEnd = rawFieldValueWithoutTrailingSpace.length;
        const queryStart = queryEnd - activeQuery.length;
        prefix = rawFieldValue.slice(0, queryStart);
        suffix = rawFieldValue.slice(queryEnd);
      } else {
        const [selectionStart, selectionEnd] = field.selection.ranges[0] ?? [
          field.position,
          field.position,
        ];
        prefix = field.getValue(0, selectionStart, "latex");
        suffix = field.getValue(selectionEnd, field.lastOffset, "latex");
      }

      pendingWrapperInput = {
        command: command.command,
        prefix,
        content: "",
        suffix,
      };
      field.dataset.pendingWrapperCommand = command.command;
      if (!renderPendingWrapperInput()) {
        clearPendingWrapperInput();
        return false;
      }
      const after = captureFieldSnapshot(field);
      emitEdit(before, after, "replace", source);
      nativeInputPendingRef.current = false;
      propsRef.current.onInputActivity(field);
      return true;
    };

    const reconcilePendingWrapperConfirmation = () => {
      if (!pendingWrapperInput || !field.isConnected) return;
      const before = captureFieldSnapshot(field);
      renderPendingWrapperInput();
      const after = captureFieldSnapshot(field);
      emitEdit(before, after, "replace", "candidate");
      nativeInputPendingRef.current = false;
    };

    const emitEdit = (
      before: ReturnType<typeof captureFieldSnapshot>,
      after: ReturnType<typeof captureFieldSnapshot>,
      editKind: EditKind,
      source: FormulaEditSource,
    ) => {
      let committedAfter = after;
      if (
        field.mode !== "latex" &&
        !rawLatexInput(field) &&
        normalizeChineseTextLatex(field.value) !== after.latex
      ) {
        field.setValue(after.latex, {
          mode: "math",
          format: "latex",
          insertionMode: "replaceAll",
          selectionMode: "after",
          silenceNotifications: true,
        });
        try {
          const restored = clampSelection(after.selection, field.lastOffset);
          field.selection = restored;
          const [start, end] = restored.ranges[0] ?? [field.lastOffset, field.lastOffset];
          field.position = restored.direction === "backward" ? start : end;
        } catch {
          field.position = field.lastOffset;
        }
        committedAfter = captureFieldSnapshot(field);
      }

      lastSnapshotRef.current = committedAfter;
      if (before.latex === committedAfter.latex) return;
      propsRef.current.onEdit(
        {
          lineId,
          beforeLatex: before.latex,
          afterLatex: committedAfter.latex,
          beforeSelection: before.selection,
          afterSelection: committedAfter.selection,
          editKind,
          source,
        },
        field,
      );
      field.resetUndo();
    };
    const handleCompositionStart = () => {
      nativeInputPendingRef.current = true;
      const setting = getCaretAutoExitSetting(field);
      pendingAutoExitSetting =
        setting && propsRef.current.inputBehavior[setting] ? setting : null;
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
      let after = captureFieldSnapshot(field);
      const inferredAutoExitSetting =
        pendingAutoExitSetting ??
        inferPlaceholderAutoExitSetting(before.latex, after.latex);
      const liveAutoExitSetting = getCaretAutoExitSetting(field);
      const autoExitSetting =
        liveAutoExitSetting ??
        (field.position < field.lastOffset ? inferredAutoExitSetting : null);
      if (autoExitSetting && before.latex !== after.latex) {
        moveCaretThroughEnabledAutoExitContainers(
          field,
          propsRef.current.inputBehavior,
          autoExitSetting,
        );
        after = captureFieldSnapshot(field);
      }
      pendingAutoExitSetting = null;
      compositionStartRef.current = null;
      emitEdit(before, after, "composition", "keyboard");
      nativeInputPendingRef.current = false;
      propsRef.current.onInputActivity(field);
      syncFrameSize();
    };
    const handleBeforeInput = (event: InputEvent) => {
      if (renderingPendingWrapperInput) return;
      if (
        pendingWrapperConfirmationKey &&
        pendingWrapperInput &&
        !event.isComposing
      ) {
        event.preventDefault();
        event.stopImmediatePropagation();
        nativeInputPendingRef.current = false;
        return;
      }
      if (!event.isComposing && !pendingWrapperInput) {
        nativeInputPendingRef.current = true;
      }
      if (pendingWrapperInput && !event.isComposing) {
        if (event.inputType === "insertText" && event.data !== null) {
          event.preventDefault();
          event.stopImmediatePropagation();
          const before = captureFieldSnapshot(field);
          const inputCharacters = Array.from(event.data);
          const autoExit =
            propsRef.current.inputBehavior.autoExitWrapperCommand &&
            inputCharacters.length > 0;
          const wrappedInput = autoExit
            ? inputCharacters.slice(0, 1).join("")
            : event.data;
          const trailingInput = autoExit
            ? inputCharacters.slice(1).join("")
            : "";
          if (
            pendingWrapperInput.command === "\\mathcal" &&
            pendingWrapperInput.content.length === 0 &&
            /^[a-z]$/.test(wrappedInput)
          ) {
            pendingWrapperInput.command = "\\mathscr";
            field.dataset.pendingWrapperCommand = "\\mathscr";
          }
          pendingWrapperInput.content += wrappedInput;
          renderPendingWrapperInput();
          if (autoExit) {
            clearPendingWrapperInput();
            if (trailingInput) {
              field.insert(trailingInput, {
                mode: "math",
                format: "latex",
                insertionMode: "replaceSelection",
                selectionMode: "after",
                focus: true,
                scrollIntoView: false,
              });
            }
          }
          const after = captureFieldSnapshot(field);
          emitEdit(before, after, "insert", "keyboard");
          nativeInputPendingRef.current = false;
          propsRef.current.onInputActivity(field);
          syncFrameSize();
          return;
        }
        if (event.inputType === "deleteContentBackward") {
          if (pendingWrapperInput.content) {
            event.preventDefault();
            event.stopImmediatePropagation();
            const before = captureFieldSnapshot(field);
            pendingWrapperInput.content = Array.from(
              pendingWrapperInput.content,
            )
              .slice(0, -1)
              .join("");
            renderPendingWrapperInput();
            const after = captureFieldSnapshot(field);
            emitEdit(before, after, "delete-backward", "keyboard");
            nativeInputPendingRef.current = false;
            propsRef.current.onInputActivity(field);
            syncFrameSize();
            return;
          }
          clearPendingWrapperInput();
        }
      }

      if (!event.isComposing && !composing) {
        pendingAutoExitSetting = null;
        if (isSingleDirectInput(event, field)) {
          const setting = getCaretAutoExitSetting(field);
          if (setting && propsRef.current.inputBehavior[setting]) {
            const before = captureFieldSnapshot(field);
            let inserted = field.insert(event.data ?? "", {
              mode: "math",
              format: "latex",
              insertionMode: "replaceSelection",
              selectionMode: "after",
              focus: true,
              scrollIntoView: false,
            });
            if (!inserted && event.data) {
              inserted = field.executeCommand(["insert", event.data]);
            }
            if (inserted) {
              event.preventDefault();
              event.stopImmediatePropagation();
              moveCaretThroughEnabledAutoExitContainers(
                field,
                propsRef.current.inputBehavior,
                setting,
              );
              const after = captureFieldSnapshot(field);
              emitEdit(before, after, "insert", "keyboard");
              nativeInputPendingRef.current = false;
              propsRef.current.onInputActivity(field);
              syncFrameSize();
              return;
            }
            nativeInputPendingRef.current = false;
          }
        }
      }
    };
    const handleInput = (event: Event) => {
      if (composing) return;
      const before = lastSnapshotRef.current ?? captureFieldSnapshot(field);
      let after = captureFieldSnapshot(field);
      const directInputSetting =
        event instanceof InputEvent && isSingleDirectInput(event, field)
          ? getCaretAutoExitSetting(field)
          : null;
      const placeholderSetting = inferPlaceholderAutoExitSetting(
        before.latex,
        after.latex,
      );
      const inferredAutoExitSetting =
        pendingAutoExitSetting ?? directInputSetting ?? placeholderSetting;
      const liveAutoExitSetting = getCaretAutoExitSetting(field);
      const autoExitSetting =
        liveAutoExitSetting ??
        (field.position < field.lastOffset ? inferredAutoExitSetting : null);
      if (
        autoExitSetting &&
        propsRef.current.inputBehavior[autoExitSetting]
      ) {
        moveCaretThroughEnabledAutoExitContainers(
          field,
          propsRef.current.inputBehavior,
          autoExitSetting,
        );
        after = captureFieldSnapshot(field);
      }
      pendingAutoExitSetting = null;
      const inputType =
        event instanceof InputEvent ? event.inputType || "insertText" : "insertText";
      emitEdit(
        before,
        after,
        inferEditKind(inputType, before.selection),
        inferEditSource(inputType),
      );
      nativeInputPendingRef.current = false;
      propsRef.current.onInputActivity(field);
      syncFrameSize();
      window.requestAnimationFrame(() => {
        if (!field.isConnected) return;
        propsRef.current.onInputActivity(field);
        decorateNativeSuggestionPreviews();
      });
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
      nativeInputPendingRef.current = false;
      clearPendingWrapperInput();
      propsRef.current.onCommitPending();
    };
    const confirmRawWrapperCommand = (event: KeyboardEvent) => {
      if (
        pendingWrapperInput ||
        event.isComposing ||
        event.metaKey ||
        event.ctrlKey ||
        event.altKey ||
        !(
          event.key === "Enter" ||
          event.key === "Tab" ||
          event.key === " " ||
          event.code === "Space"
        )
      ) {
        return false;
      }

      const wrapperQuery = trailingCommandQuery(field);
      const wrapperCommand = wrapperQuery
        ? exactWrapperCommand(wrapperQuery)
        : null;
      if (!wrapperCommand) return false;

      pendingAutoExitSetting = null;
      if (
        !beginPendingWrapperInput(
          wrapperCommand,
          wrapperQuery,
          "candidate",
        )
      ) {
        return false;
      }
      pendingWrapperConfirmationKey =
        event.code === "Space" ? "Space" : event.key;
      window.setTimeout(() => {
        if (!pendingWrapperConfirmationKey) return;
        reconcilePendingWrapperConfirmation();
        pendingWrapperConfirmationKey = "";
      }, 120);
      event.preventDefault();
      event.stopImmediatePropagation();
      return true;
    };
    const handleWrapperConfirmationKeyUp = (event: KeyboardEvent) => {
      if (!pendingWrapperConfirmationKey) return;
      const releasedKey = event.code === "Space" ? "Space" : event.key;
      if (releasedKey !== pendingWrapperConfirmationKey) return;
      reconcilePendingWrapperConfirmation();
      pendingWrapperConfirmationKey = "";
      nativeInputPendingRef.current = false;
    };
    const commitUnreportedKeyboardMutation = () => {
      if (composing || !field.isConnected) return;
      const before = lastSnapshotRef.current ?? captureFieldSnapshot(field);
      let after = captureFieldSnapshot(field);
      if (before.latex === after.latex) return;

      const inferredAutoExitSetting =
        pendingAutoExitSetting ??
        inferPlaceholderAutoExitSetting(before.latex, after.latex);
      const liveAutoExitSetting = getCaretAutoExitSetting(field);
      const autoExitSetting =
        liveAutoExitSetting ??
        (field.position < field.lastOffset ? inferredAutoExitSetting : null);
      if (
        autoExitSetting &&
        propsRef.current.inputBehavior[autoExitSetting]
      ) {
        moveCaretThroughEnabledAutoExitContainers(
          field,
          propsRef.current.inputBehavior,
          autoExitSetting,
        );
        after = captureFieldSnapshot(field);
      }
      pendingAutoExitSetting = null;
      emitEdit(
        before,
        after,
        inferEditKind("insertText", before.selection),
        "keyboard",
      );
      syncFrameSize();
    };
    const scheduleInputActivity = () => {
      window.requestAnimationFrame(() => {
        if (!field.isConnected) return;
        // MathLive normally emits an `input` event on the custom element, but
        // some native keyboard paths update its internal model without
        // redispatching that event. Reconcile the field snapshot on keyup/input
        // so React never writes the previous placeholder value back over a
        // correctly typed character. The normal input path is a no-op here
        // because lastSnapshotRef has already advanced.
        commitUnreportedKeyboardMutation();
        propsRef.current.onInputActivity(field);
        decorateNativeSuggestionPreviews();
      });
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      // `composing` can remain true after MathLive commits a command without a
      // matching compositionend event. A new non-composing key proves that the
      // IME transaction has ended, so release the stale local guard as well.
      if (event.isComposing) return;
      if (composing) {
        composing = false;
        compositionStartRef.current = null;
        lastSnapshotRef.current = captureFieldSnapshot(field);
      }
      if (pendingWrapperInput && event.key === "Enter") {
        event.preventDefault();
        event.stopImmediatePropagation();
        const exitPosition = field.position;
        clearPendingWrapperInput();
        field.position = Math.min(exitPosition, field.lastOffset);
        field.focus();
        field.shadowRoot
          ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
          ?.focus({ preventScroll: true });
        propsRef.current.onCommitPending();
        syncFrameSize();
        return;
      }
      if (confirmRawWrapperCommand(event)) return;
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
      clearPendingWrapperInput();
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
    host.replaceChildren(field, wrapperMeasure);
    // MathLive mounts a pre-filled field with the whole formula selected.
    // Collapse that implicit selection so toolbar commands insert at the end
    // instead of unexpectedly replacing/wrapping the entire line.
    field.position = field.lastOffset;
    field.resetUndo();
    lastSnapshotRef.current = captureFieldSnapshot(field);
    fieldRef.current = field;
    propsRef.current.register(lineId, field);
    propsRef.current.registerWrapperInputController(lineId, {
      begin: beginPendingWrapperInput,
    });
    field.addEventListener("compositionstart", handleCompositionStart);
    field.addEventListener("compositionend", handleCompositionEnd);
    field.addEventListener("beforeinput", handleBeforeInput, true);
    field.addEventListener("input", handleInput);
    field.addEventListener("selection-change", handleSelectionChange);
    field.addEventListener("focus", handleFocus);
    field.addEventListener("blur", handleBlur);
    field.addEventListener("keydown", handleKeyDown, true);
    const keyboardSink = field.shadowRoot?.querySelector<HTMLElement>(
      '[part="keyboard-sink"]',
    );
    keyboardSink?.addEventListener("input", scheduleInputActivity, true);
    keyboardSink?.addEventListener("keyup", scheduleInputActivity, true);
    field.addEventListener("paste", handlePaste, true);
    window.addEventListener("keyup", handleWrapperConfirmationKeyUp, true);
    host.addEventListener("pointerdown", handlePointerDown, true);
    const content = field.shadowRoot?.querySelector<HTMLElement>('[part="content"]');
    const resizeObserver = content
      ? new ResizeObserver(() => {
          syncFrameSize();
          schedulePendingWrapperPlaceholderPosition();
        })
      : null;
    const inputMutationObserver = field.shadowRoot
      ? new MutationObserver(() => {
          scheduleInputActivity();
          schedulePendingWrapperPlaceholderPosition();
        })
      : null;
    if (content) resizeObserver?.observe(content);
    if (field.shadowRoot) {
      inputMutationObserver?.observe(field.shadowRoot, {
        childList: true,
        characterData: true,
        subtree: true,
      });
    }
    syncFrameSize();

    return () => {
      window.cancelAnimationFrame(resizeFrame);
      window.cancelAnimationFrame(wrapperPlaceholderFrame);
      window.clearTimeout(resizeTimer);
      resizeObserver?.disconnect();
      inputMutationObserver?.disconnect();
      syncFrameSizeRef.current = null;
      field.removeEventListener("compositionstart", handleCompositionStart);
      field.removeEventListener("compositionend", handleCompositionEnd);
      field.removeEventListener("beforeinput", handleBeforeInput, true);
      field.removeEventListener("input", handleInput);
      field.removeEventListener("selection-change", handleSelectionChange);
      field.removeEventListener("focus", handleFocus);
      field.removeEventListener("blur", handleBlur);
      field.removeEventListener("keydown", handleKeyDown, true);
      keyboardSink?.removeEventListener("input", scheduleInputActivity, true);
      keyboardSink?.removeEventListener("keyup", scheduleInputActivity, true);
      field.removeEventListener("paste", handlePaste, true);
      window.removeEventListener("keyup", handleWrapperConfirmationKeyUp, true);
      host.removeEventListener("pointerdown", handlePointerDown, true);
      host.closest<HTMLElement>(".formula-line")?.style.removeProperty(
        "--formula-row-height",
      );
      propsRef.current.registerWrapperInputController(lineId, null);
      propsRef.current.register(lineId, null);
      fieldRef.current = null;
      lastSnapshotRef.current = null;
      compositionStartRef.current = null;
      pendingAutoExitSetting = null;
      pendingWrapperConfirmationKey = "";
      nativeInputPendingRef.current = false;
      clearPendingWrapperInput();
      host.replaceChildren();
    };
  }, []);

  useEffect(() => {
    const field = fieldRef.current;
    if (!field || nativeInputPendingRef.current) return;

    const currentStoredLatex = useEditorStore
      .getState()
      .lines.find((line) => line.id === props.lineId)?.latex;
    if (currentStoredLatex !== undefined && currentStoredLatex !== props.latex) {
      return;
    }

    // 中文文本规范化不改变 MathLive 的数学原子结构，因此无需重建；
    // 正体微分、常用函数等数学排版规范化则必须写回 MathLive，才能
    // 在编辑器中立即显示正体。写回时恢复原选区，避免光标跳到末尾。
    if (normalizeChineseTextLatex(field.value) === props.latex) {
      lastSnapshotRef.current = {
        latex: props.latex,
        selection: captureSelection(field),
      };
      return;
    }

    const selection = captureSelection(field);
    field.setValue(props.latex, {
      mode: "math",
      format: "latex",
      insertionMode: "replaceAll",
      selectionMode: "after",
      silenceNotifications: true,
    });
    try {
      field.selection = clampSelection(selection, field.lastOffset);
    } catch {
      field.position = field.lastOffset;
    }
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
    if (fieldRef.current) {
      fieldRef.current.smartFence = props.autoPairDelimiters;
    }
  }, [props.autoPairDelimiters]);

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
      zoom,
      onPasteImage,
      onHistoryBusyChange,
      overlay,
    },
    ref,
  ) {
    const surfaceRef = useRef<HTMLDivElement>(null);
    const fieldRefs = useRef(new Map<string, MathfieldElement>());
    const wrapperInputControllersRef = useRef(
      new Map<string, WrapperInputController>(),
    );
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
    const autoPairDelimiters = useEditorStore(
      (state) => state.autoPairDelimiters,
    );
    const inputBehavior = useEditorStore((state) => state.inputBehavior);
    const isEn = language === "en";

    linesRef.current = lines;
    const resolvedActiveLineId =
      lines.find((line) => line.id === activeLineId)?.id ?? lines[0]?.id ?? null;
    activeLineIdRef.current = resolvedActiveLineId;

    const suggestions = useMemo(
      () =>
        query
          ? visibleCommandSuggestions(
              query,
              usage,
              personalize,
              suggestionCount,
              inputBehavior,
            )
          : [],
      [query, usage, personalize, suggestionCount, inputBehavior],
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

    const registerWrapperInputController = (
      lineId: string,
      controller: WrapperInputController | null,
    ) => {
      if (controller) wrapperInputControllersRef.current.set(lineId, controller);
      else wrapperInputControllersRef.current.delete(lineId);
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
      const activeCommandQuery = trailingCommandQuery(field, normalized);
      const hasRawCommandInput = field.mode === "latex" || Boolean(rawLatexInput(field));
      const suppressed = suppressedSuggestionRef.current;
      if (suppressed) {
        if (
          suppressed.lineId === lineId &&
          suppressed.value.trim() === normalized.trim() &&
          !hasRawCommandInput
        ) {
          queryRef.current = "";
          setQuery("");
          return;
        }
        suppressedSuggestionRef.current = null;
      }

      if (activeCommandQuery) {
        setQuery(activeCommandQuery);
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
      onHistoryBusyChange?.(true);
      suppressedHistoryLineIdRef.current = lineId;
      let changed = false;
      let after = before;
      try {
        changed = mutate();
        if (!changed) return false;

        const startedAt = performance.now();
        while (performance.now() - startedAt < 1000) {
          await new Promise<void>((resolve) =>
            window.setTimeout(resolve, 16),
          );
          after = captureFieldSnapshot(field);
          const nativeRecommendationVisible =
            document
              .getElementById("mathlive-suggestion-popover")
              ?.classList.contains("is-visible") ?? false;
          if (
            after.latex !== before.latex &&
            !nativeRecommendationVisible
          ) {
            break;
          }
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
      } finally {
        suppressedHistoryLineIdRef.current = null;
        onHistoryBusyChange?.(false);
      }
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
      const rawFieldValue = activeQuery
        ? fieldValueIncludingRawQuery(field, activeQuery)
        : field.value;
      const rawFieldValueWithoutTrailingSpace = rawFieldValue.trimEnd();
      const replacesRawCommand = Boolean(
        activeQuery &&
          rawFieldValueWithoutTrailingSpace.endsWith(activeQuery) &&
          (field.mode === "latex" || field.lastOffset === 0),
      );
      if (activeQuery && !queryRange && !replacesRawCommand) {
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
      const autoExitSetting = getCaretAutoExitSetting(field);
      const historySource: FormulaEditSource =
        source === "candidate" ? "candidate" : "toolbar";

      if (
        !selectedLatex &&
        wrapperCommandPreviews.has(command.command) &&
        wrapperInputControllersRef.current
          .get(targetLineId)
          ?.begin(command, activeQuery, historySource)
      ) {
        const normalizedValue = normalizeChineseLatex(field.value);
        suppressedSuggestionRef.current = {
          lineId: targetLineId,
          value: normalizedValue,
        };
        recordCommand(command.id, activeQuery, source);
        setQuery("");
        selectSuggestionIndex(0);
        return;
      }

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
            const hasPlaceholder = insertionTemplate.includes("\\placeholder{}");
            if (replacesRawCommand) {
              const queryEnd = rawFieldValueWithoutTrailingSpace.length;
              const queryStart = queryEnd - activeQuery.length;
              const trailingSpace = rawFieldValue.slice(queryEnd);
              const nextValue =
                rawFieldValue.slice(0, queryStart) +
                insertionTemplate +
                trailingSpace;
              field.setValue(nextValue, {
                mode: "math",
                format: "latex",
                insertionMode: "replaceAll",
                selectionMode: hasPlaceholder ? "placeholder" : "after",
                silenceNotifications: true,
              });
              field.focus();
              return true;
            }
            if (queryRange) {
              field.selection = {
                ranges: [queryRange],
                direction: "forward",
              };
            }
            const inserted = field.insert(insertionTemplate, {
              mode: "math",
              format: "latex",
              insertionMode: "replaceSelection",
              selectionMode: hasPlaceholder ? "placeholder" : "after",
              focus: true,
              scrollIntoView: false,
            });
            if (
              inserted &&
              !hasPlaceholder &&
              autoExitSetting &&
              inputBehavior[autoExitSetting]
            ) {
              moveCaretThroughEnabledAutoExitContainers(field, inputBehavior);
            }
            return inserted;
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

      const insertsChineseIdeographicComma =
        event.code === "Backslash" &&
        event.key === "、" &&
        !event.metaKey &&
        !event.ctrlKey &&
        !event.altKey;
      if (insertsChineseIdeographicComma) {
        // macOS 中文输入法会在这个物理键上先触发 Backslash keydown，
        // 随后再通过输入法事务提交“、”。这里只阻止 MathLive 根据
        // 物理键位额外插入反斜杠，顿号本身交给后续 composition/input。
        event.preventDefault();
        event.stopPropagation();
        return;
      }

      const normalizedFieldValue = normalizeChineseLatex(field.value);
      const suppressedSuggestion = suppressedSuggestionRef.current;
      const suppressesCurrentTrailingCommand =
        suppressedSuggestion?.lineId === lineId &&
        suppressedSuggestion.value.trim() === normalizedFieldValue.trim();
      const liveQuery = suppressesCurrentTrailingCommand
        ? ""
        : trailingCommandQuery(field);
      const candidateQuery = suppressesCurrentTrailingCommand
        ? ""
        : liveQuery || query;
      const currentState = useEditorStore.getState();
      const liveSuggestions = candidateQuery
        ? visibleCommandSuggestions(
            candidateQuery,
            currentState.usage,
            currentState.personalize,
            currentState.suggestionCount,
            currentState.inputBehavior,
          )
        : [];
      const wrapperQuery = trailingCommandQuery(field) || candidateQuery;
      const wrapperCommand = wrapperQuery
        ? exactWrapperCommand(wrapperQuery)
        : null;
      const confirmsWrapperCommand =
        wrapperCommand &&
        !event.metaKey &&
        !event.ctrlKey &&
        !event.altKey &&
        (event.key === "Enter" || event.key === "Tab");

      if (confirmsWrapperCommand) {
        event.preventDefault();
        event.stopPropagation();
        insertCommand(wrapperCommand, "candidate", wrapperQuery);
        return;
      }

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
      if (
        !suppressesCurrentTrailingCommand &&
        !liveSuggestions.length &&
        nativeRecommendationVisible
      ) {
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
        const rawCommandInput = rawLatexInput(field);

        // MathLive keeps an unfinished command such as `\\mat` outside the
        // public formula value. It is not an empty line: one physical delete
        // must remove one raw command character and must not remove the row.
        if (field.mode === "latex" || rawCommandInput) return;

        // Preserve the established VisualTeX line behavior: Backspace/Delete
        // on a genuinely empty line removes that line and focuses the previous
        // formula at its mathematical end.
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
        const scriptRegion = getScriptCaretRegion(field);
        const requestedRegion = event.key === "ArrowUp" ? "upper" : "lower";
        if (scriptRegion) {
          event.preventDefault();
          event.stopImmediatePropagation();
          if (scriptRegion !== requestedRegion) {
            field.executeCommand("moveToOpposite");
          }
          return;
        }

        // MathLive can report a successful vertical move even for a flat,
        // single-atom formula. That traps ArrowUp/ArrowDown inside the row.
        // Only ask MathLive to navigate vertically when the formula actually
        // contains a structure with meaningful upper/lower geometry.
        if (tallFormulaPattern.test(field.value)) {
          const beforeSelection = captureSelection(field);
          const movedInsideFormula = field.executeCommand(
            event.key === "ArrowUp" ? "moveUp" : "moveDown",
          );
          const afterSelection = captureSelection(field);
          if (
            movedInsideFormula &&
            JSON.stringify(beforeSelection) !== JSON.stringify(afterSelection)
          ) {
            event.preventDefault();
            event.stopImmediatePropagation();
            return;
          }
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
                autoPairDelimiters={autoPairDelimiters}
                inputBehavior={inputBehavior}
                register={registerField}
                registerWrapperInputController={registerWrapperInputController}
                onEdit={handleFieldEdit}
                onInputActivity={(field) =>
                  refreshSuggestionQuery(
                    lineId,
                    field,
                    normalizeChineseLatex(field.value),
                  )
                }
                onCommitPending={() => historyManager.commitPendingTransaction()}
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
