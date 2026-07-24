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
import {
  MathfieldElement,
  convertLatexToMarkup,
  type Style,
} from "mathlive";
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
  normalizeContextualUprightSymbols,
  normalizeMathLiveCanonicalUprightCommands,
  visualTexUprightInlineShortcuts,
} from "./normalizeChineseLatex";
import { ImeCompositionGuard } from "./imeCompositionGuard";

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

interface FormulaFieldProps {
  lineId: string;
  index: number;
  latex: string;
  zoom: number;
  language: "cn" | "en";
  autoPairDelimiters: boolean;
  inputBehavior: InputBehaviorSettings;
  register: (lineId: string, field: MathfieldElement | null) => void;
  onEdit: (edit: FormulaFieldEdit, field: MathfieldElement) => void;
  onInputActivity: (field: MathfieldElement) => void;
  onFocus: (index: number, field: MathfieldElement) => void;
  onCommitPending: () => void;
  onKeyDown: (index: number, event: KeyboardEvent, field: MathfieldElement) => void;
  onPasteImage?: (file: File, target: MathEditorInsertionTarget) => void;
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

interface RawCommandAnchor {
  latex: string;
  selection: MathSelectionSnapshot;
  position: number;
  visualCaret: WrapperCaretAnchor | null;
}

const rawCommandAnchors = new WeakMap<MathfieldElement, RawCommandAnchor>();

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

function rememberRawCommandAnchor(field: MathfieldElement) {
  rawCommandAnchors.set(field, {
    latex: normalizeChineseLatex(field.value),
    selection: captureSelection(field),
    position: field.position,
    visualCaret: captureWrapperCaretAnchor(field),
  });
}

function restoreRawCommandAnchor(
  field: MathfieldElement,
  anchor: RawCommandAnchor,
) {
  const rejectedRawGroup = field.executeCommand(["complete", "reject"]);
  if (!rejectedRawGroup) field.mode = "math";
  if (normalizeChineseLatex(field.value) !== anchor.latex) {
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
  if (!panel) {
    scheduleStableNativeInputPopoverSync();
    return;
  }

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
  scheduleStableNativeInputPopoverSync();
}

const STABLE_NATIVE_INPUT_POPOVER_ID =
  "visualtex-native-input-suggestion-popover";
let stableNativeInputPopoverFrame = 0;
let stableNativeInputPopoverHideTimer = 0;
let nativeInputPopoverBodyObserver: MutationObserver | null = null;
let nativeInputPopoverSourceObserver: MutationObserver | null = null;
let observedNativeInputPopoverSource: HTMLElement | null = null;

function getNativeInputPopoverSource() {
  return document.getElementById("mathlive-suggestion-popover");
}

function ensureStableNativeInputPopover() {
  let panel = document.getElementById(STABLE_NATIVE_INPUT_POPOVER_ID);
  if (panel) return panel;

  panel = document.createElement("div");
  panel.id = STABLE_NATIVE_INPUT_POPOVER_ID;
  panel.setAttribute("aria-hidden", "true");
  panel.addEventListener("pointerdown", (event) => event.preventDefault());
  panel.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof Element)) return;
    const item = target.closest<HTMLElement>("li[data-command]");
    const command = item?.dataset.command ?? "";
    if (!command) return;
    const sourceItem = Array.from(
      getNativeInputPopoverSource()?.querySelectorAll<HTMLElement>(
        "li[data-command]",
      ) ?? [],
    ).find((candidate) => candidate.dataset.command === command);
    sourceItem?.dispatchEvent(
      new MouseEvent("click", {
        bubbles: true,
        cancelable: true,
        view: window,
      }),
    );
  });
  document.body.append(panel);
  return panel;
}

function bindNativeInputPopoverSource() {
  const source = getNativeInputPopoverSource();
  if (source === observedNativeInputPopoverSource) return;
  nativeInputPopoverSourceObserver?.disconnect();
  observedNativeInputPopoverSource = source;
  if (!source) return;
  nativeInputPopoverSourceObserver = new MutationObserver(() =>
    scheduleStableNativeInputPopoverSync(),
  );
  nativeInputPopoverSourceObserver.observe(source, {
    attributes: true,
    attributeFilter: ["class", "style", "aria-hidden"],
    childList: true,
    characterData: true,
    subtree: true,
  });
}

function ensureNativeInputPopoverObservers() {
  if (nativeInputPopoverBodyObserver || !document.body) return;
  nativeInputPopoverBodyObserver = new MutationObserver(() => {
    bindNativeInputPopoverSource();
    scheduleStableNativeInputPopoverSync();
  });
  nativeInputPopoverBodyObserver.observe(document.body, { childList: true });
  bindNativeInputPopoverSource();
}

function syncStableNativeInputPopoverSelection(command: string) {
  for (const panel of [
    getNativeInputPopoverSource(),
    document.getElementById(STABLE_NATIVE_INPUT_POPOVER_ID),
  ]) {
    const items = Array.from(
      panel?.querySelectorAll<HTMLElement>("li[data-command]") ?? [],
    );
    for (const item of items) {
      item.classList.toggle(
        "ML__popover__current",
        item.dataset.command === command,
      );
    }
    items
      .find((item) => item.dataset.command === command)
      ?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }
}

function syncStableNativeInputPopover() {
  bindNativeInputPopoverSource();
  const source = getNativeInputPopoverSource();
  const sourceVisible = Boolean(
    source?.classList.contains("is-visible") &&
      source.querySelector("li[data-command]"),
  );
  const stable = ensureStableNativeInputPopover();

  if (!source || !sourceVisible) {
    window.clearTimeout(stableNativeInputPopoverHideTimer);
    stableNativeInputPopoverHideTimer = window.setTimeout(() => {
      const latest = getNativeInputPopoverSource();
      if (
        latest?.classList.contains("is-visible") &&
        latest.querySelector("li[data-command]")
      ) {
        scheduleStableNativeInputPopoverSync();
        return;
      }
      stable.classList.remove("is-visible");
      stable.setAttribute("aria-hidden", "true");
    }, 64);
    return;
  }

  window.clearTimeout(stableNativeInputPopoverHideTimer);
  source.dataset.visualtexInputPopoverSource = "true";
  const nextHtml = source.innerHTML;
  if (stable.innerHTML !== nextHtml) stable.innerHTML = nextHtml;
  stable.classList.toggle("top-tip", source.classList.contains("top-tip"));
  stable.classList.toggle(
    "bottom-tip",
    source.classList.contains("bottom-tip"),
  );
  stable.style.left = source.style.left;
  stable.style.top = source.style.top;
  stable.classList.add("is-visible");
  stable.setAttribute("aria-hidden", "false");
}

function scheduleStableNativeInputPopoverSync() {
  ensureNativeInputPopoverObservers();
  window.cancelAnimationFrame(stableNativeInputPopoverFrame);
  stableNativeInputPopoverFrame = window.requestAnimationFrame(
    syncStableNativeInputPopover,
  );
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

const visualTexPlaceholderStyleId = "visualtex-structural-placeholder-style";
const visualTexPlaceholderClass = "visualtex-structural-placeholder";
const visualTexPlaceholderCaretClass =
  "visualtex-structural-placeholder-caret";
const visualTexPlaceholderSelectionClass =
  "has-visualtex-structural-placeholder-selection";
const visualTexRawLatexClass = "has-visualtex-raw-latex-command";
const visualTexPointerSelectingClass = "visualtex-pointer-selecting";
const visualTexPointerSelectingFields = new WeakSet<MathfieldElement>();

function markVisualTexStructuralPlaceholders(field: MathfieldElement) {
  const shadowRoot = field.shadowRoot;
  if (!shadowRoot || visualTexPointerSelectingFields.has(field)) return;
  const placeholderSymbol = field.placeholderSymbol || "▢";
  const selectionRanges = field.selection.ranges;
  const [selectionStart, selectionEnd] = selectionRanges[0] ?? [-1, -1];
  const isFocusedPlaceholderSelection =
    !visualTexPointerSelectingFields.has(field) &&
    selectionRanges.length === 1 &&
    Math.abs(selectionEnd - selectionStart) <= 1 &&
    !shadowRoot.querySelector(".ML__raw-latex");
  let selectedPlaceholder = false;

  shadowRoot
    .querySelectorAll<HTMLElement>(".ML__cmr[data-atom-id], .ML__placeholder")
    .forEach((node) => {
      const visibleText = Array.from(node.childNodes)
        .filter(
          (child) =>
            !(child instanceof HTMLElement) ||
            !child.classList.contains(visualTexPlaceholderCaretClass),
        )
        .map((child) => child.textContent ?? "")
        .join("")
        .trim();
      const isPlaceholder =
        node.classList.contains(visualTexPlaceholderClass) ||
        node.classList.contains("ML__placeholder") ||
        visibleText === placeholderSymbol;
      node.classList.toggle(visualTexPlaceholderClass, isPlaceholder);

      const isSelected = Boolean(
        isFocusedPlaceholderSelection &&
          isPlaceholder &&
          (node.classList.contains("ML__placeholder-selected") ||
            node.classList.contains("ML__selected") ||
            node.closest(".ML__selected")),
      );
      let caret = node.querySelector<HTMLElement>(
        `:scope > .${visualTexPlaceholderCaretClass}`,
      );
      if (isSelected) {
        selectedPlaceholder = true;
        if (!caret) {
          caret = document.createElement("span");
          caret.className = visualTexPlaceholderCaretClass;
          caret.setAttribute("aria-hidden", "true");
          node.append(caret);
        }
      } else {
        caret?.remove();
      }
    });

  const container = shadowRoot.querySelector<HTMLElement>(".ML__container");
  container?.classList.toggle(
    visualTexPlaceholderSelectionClass,
    selectedPlaceholder,
  );
  container?.classList.toggle(
    visualTexRawLatexClass,
    Boolean(shadowRoot.querySelector(".ML__raw-latex")),
  );
}

function installVisualTexStructuralPlaceholderStyle(field: MathfieldElement) {
  const shadowRoot = field.shadowRoot;
  if (!shadowRoot) return;
  if (!shadowRoot.getElementById(visualTexPlaceholderStyleId)) {
    const style = document.createElement("style");
    style.id = visualTexPlaceholderStyleId;
    style.textContent = `
      .ML__placeholder,
      .${visualTexPlaceholderClass} {
        position: relative !important;
        display: inline-block !important;
        box-sizing: border-box !important;
        width: 0.40em !important;
        min-width: 0.32em !important;
        height: 0.74em !important;
        margin: 0 0.065em !important;
        padding: 0 !important;
        overflow: hidden !important;
        border: 0 !important;
        border-radius: 0.075em !important;
        background: #d9edf9 !important;
        color: transparent !important;
        opacity: 1 !important;
        text-indent: -999px !important;
        vertical-align: -0.075em !important;
        box-shadow: none !important;
      }

      .ML__mfrac .ML__placeholder,
      .ML__mfrac .${visualTexPlaceholderClass} {
        width: 0.38em !important;
        min-width: 0.31em !important;
        height: 0.70em !important;
        border-radius: 0.07em !important;
      }

      .ML__msubsup .ML__placeholder,
      .ML__msubsup .${visualTexPlaceholderClass},
      .ML__op-group .ML__placeholder,
      .ML__op-group .${visualTexPlaceholderClass} {
        width: 0.35em !important;
        min-width: 0.28em !important;
        height: 0.66em !important;
        border-radius: 0.065em !important;
      }

      .ML__sqrt .ML__placeholder,
      .ML__sqrt .${visualTexPlaceholderClass} {
        width: 0.40em !important;
        min-width: 0.32em !important;
        height: 0.71em !important;
      }

      .ML__mtable .ML__placeholder,
      .ML__mtable .${visualTexPlaceholderClass} {
        width: 0.42em !important;
        min-width: 0.34em !important;
        height: 0.73em !important;
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
        border: 0 !important;
        background: #cfe8f7 !important;
        color: transparent !important;
        opacity: 1 !important;
        box-shadow: none !important;
      }

      .${visualTexPlaceholderCaretClass} {
        position: absolute !important;
        z-index: 2 !important;
        left: 0 !important;
        top: 0.04em !important;
        display: block !important;
        width: 0 !important;
        height: 0.66em !important;
        margin: 0 !important;
        padding: 0 !important;
        border: 0 !important;
        border-left: max(1px, 0.045em) solid var(--_caret-color) !important;
        border-radius: 1px !important;
        background: transparent !important;
        opacity: 1;
        text-indent: 0 !important;
        animation: visualtex-placeholder-caret-blink 1.05s step-end infinite !important;
        pointer-events: none !important;
      }

      @keyframes visualtex-placeholder-caret-blink {
        0%, 48% { opacity: 1; }
        49%, 100% { opacity: 0; }
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
    `;
    shadowRoot.append(style);
  }
  markVisualTexStructuralPlaceholders(field);
}

const visualTexPointerPlaceholderStyleId =
  "visualtex-pointer-placeholder-style";
const visualTexPointerPlaceholderStyles = new WeakMap<
  MathfieldElement,
  string
>();

function installPointerPlaceholderSnapshotStyle(field: MathfieldElement) {
  const shadowRoot = field.shadowRoot;
  if (!shadowRoot) return;
  shadowRoot.getElementById(visualTexPointerPlaceholderStyleId)?.remove();

  const placeholderSymbol = field.placeholderSymbol || "▢";
  const rules = Array.from(
    shadowRoot.querySelectorAll<HTMLElement>(
      ".ML__cmr[data-atom-id], .ML__placeholder",
    ),
  ).flatMap((node) => {
    const visibleText = node.textContent?.trim() ?? "";
    if (
      !node.classList.contains(visualTexPlaceholderClass) &&
      !node.classList.contains("ML__placeholder") &&
      visibleText !== placeholderSymbol
    ) {
      return [];
    }
    const atomId = node.dataset.atomId;
    if (!atomId) return [];
    const style = getComputedStyle(node);
    const escapedAtomId = CSS.escape(atomId);
    return [
      `
      :host(.${visualTexPointerSelectingClass}) [data-atom-id="${escapedAtomId}"] {
        position: relative !important;
        display: inline-block !important;
        box-sizing: border-box !important;
        width: ${style.width} !important;
        min-width: ${style.minWidth} !important;
        height: ${style.height} !important;
        margin: ${style.margin} !important;
        padding: 0 !important;
        overflow: hidden !important;
        border: 0 !important;
        border-radius: ${style.borderRadius} !important;
        background: #d9edf9 !important;
        color: transparent !important;
        opacity: 1 !important;
        text-indent: -999px !important;
        vertical-align: ${style.verticalAlign} !important;
        box-shadow: none !important;
      }
      `,
    ];
  });
  if (!rules.length) return;

  const style = document.createElement("style");
  style.id = visualTexPointerPlaceholderStyleId;
  style.textContent = rules.join("\n");
  visualTexPointerPlaceholderStyles.set(field, style.textContent);
  shadowRoot.append(style);
}

function restorePointerPlaceholderSnapshotStyle(field: MathfieldElement) {
  const shadowRoot = field.shadowRoot;
  const cssText = visualTexPointerPlaceholderStyles.get(field);
  if (
    !shadowRoot ||
    !cssText ||
    shadowRoot.getElementById(visualTexPointerPlaceholderStyleId)
  ) {
    return;
  }
  const style = document.createElement("style");
  style.id = visualTexPointerPlaceholderStyleId;
  style.textContent = cssText;
  shadowRoot.append(style);
}

function removePointerPlaceholderSnapshotStyle(field: MathfieldElement) {
  visualTexPointerPlaceholderStyles.delete(field);
  field.shadowRoot
    ?.getElementById(visualTexPointerPlaceholderStyleId)
    ?.remove();
}

function normalizeCompletedDifferentialDisplay(field: MathfieldElement) {
  if (field.mode === "latex" || !field.selectionIsCollapsed) return false;

  const portableValue = normalizeMathLiveCanonicalUprightCommands(field.value);
  const contextualValue = normalizeContextualUprightSymbols(portableValue);
  if (contextualValue === portableValue) return false;

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

type PhysicalBackslashInputKind = "ideographic-comma" | "latin-backslash";

function normalizePhysicalBackslashInput(
  beforeValue: string,
  afterValue: string,
  kind: PhysicalBackslashInputKind,
) {
  let prefixLength = 0;
  while (
    prefixLength < beforeValue.length &&
    prefixLength < afterValue.length &&
    beforeValue[prefixLength] === afterValue[prefixLength]
  ) {
    prefixLength += 1;
  }

  let suffixLength = 0;
  while (
    suffixLength < beforeValue.length - prefixLength &&
    suffixLength < afterValue.length - prefixLength &&
    beforeValue[beforeValue.length - 1 - suffixLength] ===
      afterValue[afterValue.length - 1 - suffixLength]
  ) {
    suffixLength += 1;
  }

  const insertedEnd = afterValue.length - suffixLength;
  const inserted = afterValue.slice(prefixLength, insertedEnd);
  let normalizedInserted = inserted;
  if (
    kind === "ideographic-comma" &&
    inserted.includes("、") &&
    inserted.replaceAll("\\", "") === "、"
  ) {
    normalizedInserted = "、";
  } else if (
    kind === "latin-backslash" &&
    inserted.length > 1 &&
    /^\\+$/.test(inserted)
  ) {
    normalizedInserted = "\\";
  }

  if (normalizedInserted === inserted) return afterValue;
  return (
    afterValue.slice(0, prefixLength) +
    normalizedInserted +
    afterValue.slice(insertedEnd)
  );
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

function moveNativeSuggestionSelection(
  _field: MathfieldElement,
  direction: 1 | -1,
): string | null {
  const items = getVisibleNativeSuggestionItems();
  if (!items.length) return null;
  const currentIndex = items.findIndex((item) =>
    item.classList.contains("ML__popover__current"),
  );
  const nextIndex =
    currentIndex < 0
      ? direction > 0
        ? 0
        : items.length - 1
      : (currentIndex + direction + items.length) % items.length;
  const command = items[nextIndex]?.dataset.command ?? "";
  if (!command) return null;

  // MathLive rebuilds and remounts its entire suggestion popover when the
  // nextSuggestion/previousSuggestion commands run. VisualTeX commits the
  // remembered command itself, so moving only the current-row marker preserves
  // the same popover node and avoids a visible flash on every arrow key.
  syncStableNativeInputPopoverSelection(command);
  return command;
}

function commitNativeSuggestion(
  field: MathfieldElement,
  rememberedCommand = "",
): boolean {
  const selectedCommand =
    rememberedCommand ||
    getVisibleNativeSuggestionItems().find((item) =>
      item.classList.contains("ML__popover__current"),
    )?.dataset.command ||
    "";

  if (selectedCommand) {
    const rawInput = rawLatexInput(field).trim();
    const anchor = rawCommandAnchors.get(field);
    const queryRange = rawInput
      ? findTrailingCommandRange(field, rawInput)
      : null;
    if (anchor) {
      restoreRawCommandAnchor(field, anchor);
    } else {
      field.mode = "math";
      if (queryRange) {
        field.selection = {
          ranges: [queryRange],
          direction: "forward",
        };
      }
    }
    const inserted = field.insert(selectedCommand, {
      mode: "math",
      format: "latex",
      insertionMode: "replaceSelection",
      selectionMode: "after",
      focus: true,
      scrollIntoView: false,
    });
    if (inserted) {
      rawCommandAnchors.delete(field);
      dismissNativeSuggestionPopover(field);
    }
    return inserted;
  }

  if (!getVisibleNativeSuggestionItems().length && field.mode !== "latex") {
    return false;
  }
  return field.executeCommand(["complete", "accept-all"]);
}

function dismissNativeSuggestionPopover(field: MathfieldElement) {
  const panel = document.getElementById("mathlive-suggestion-popover");
  const stablePanel = document.getElementById(STABLE_NATIVE_INPUT_POPOVER_ID);
  window.clearTimeout(stableNativeInputPopoverHideTimer);
  field.popoverPolicy = "off";
  panel?.classList.remove("is-visible");
  panel?.setAttribute("aria-hidden", "true");
  stablePanel?.classList.remove("is-visible");
  stablePanel?.setAttribute("aria-hidden", "true");
  window.setTimeout(() => {
    if (!field.isConnected) return;
    field.popoverPolicy = "auto";
  }, 120);
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
  const currentOffset = Math.max(
    field.position,
    ...field.selection.ranges.flatMap(([start, end]) => [start, end]),
  );
  const currentDepth = field.getElementInfo(currentOffset)?.depth;
  if (typeof currentDepth === "number" && currentDepth > 0) {
    for (
      let offset = currentOffset + 1;
      offset <= Math.min(field.lastOffset, currentOffset + 3);
      offset += 1
    ) {
      const info = field.getElementInfo(offset);
      if (typeof info?.depth !== "number" || info.depth >= currentDepth) continue;
      const containerLatex = (info.latex ?? "").trim();
      if (/^\^\s*\{/.test(containerLatex)) return "upper";
      if (/^_\s*\{/.test(containerLatex)) return "lower";
      break;
    }
  }

  const markers = Array.from(
    field.shadowRoot?.querySelectorAll<HTMLElement>(
      ".ML__placeholder-selected, .ML__selected, .ML__caret",
    ) ?? [],
  );
  const caret = markers.find((marker) =>
    marker.closest(".ML__msubsup, .ML__op-group"),
  );
  // Side scripts use ML__msubsup; large operators such as sum/product render
  // their over/under limits directly inside ML__op-group. Model markers above
  // distinguish ordinary ^{...} and _{...} exactly; geometry is only the
  // fallback for combined upper/lower operator limits.
  const script = caret?.closest<HTMLElement>(".ML__msubsup, .ML__op-group");
  if (!caret || !script) return null;

  const caretBounds = (caret.parentElement ?? caret).getBoundingClientRect();
  const scriptBounds = script.getBoundingClientRect();
  if (!caretBounds.height || !scriptBounds.height) return null;
  return caretBounds.top + caretBounds.height / 2 <
    scriptBounds.top + scriptBounds.height / 2
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

function moveCaretThroughEnabledAutoExitContainers(
  field: MathfieldElement,
  settings: InputBehaviorSettings,
  capturedSetting?: InputBehaviorSettingKey | null,
) {
  if (capturedSetting) {
    if (!settings[capturedSetting]) return false;
    const previousPosition = field.position;
    const changed = field.executeCommand(
      capturedSetting === "autoExitAccent"
        ? "moveToNextChar"
        : "moveAfterParent",
    );
    return Boolean(changed || field.position !== previousPosition);
  }

  let moved = false;
  for (let attempt = 0; attempt < 4; attempt += 1) {
    const setting = getCaretAutoExitSetting(field);
    if (!setting || !settings[setting]) break;

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
  const propsRef = useRef(props);
  propsRef.current = props;

  useLayoutEffect(() => {
    const host = hostRef.current;
    if (!host) return;

    const lineId = propsRef.current.lineId;
    const field = new MathfieldElement();
    MathfieldElement.locale = propsRef.current.language === "en" ? "en" : "zh-cn";
    field.value = propsRef.current.latex;
    field.className = "visual-mathfield";
    field.smartMode = false;
    field.smartFence = propsRef.current.autoPairDelimiters;
    // VisualTeX handles superscript and subscript auto-exit independently.
    // Keep MathLive's built-in superscript heuristic disabled so it cannot
    // move out of an empty superscript before the first character is inserted.
    field.smartSuperscript = false;
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
    installVisualTexStructuralPlaceholderStyle(field);

    let pointerPlaceholderFrame = 0;
    const schedulePointerPlaceholderSnapshotStyle = () => {
      window.cancelAnimationFrame(pointerPlaceholderFrame);
      pointerPlaceholderFrame = window.requestAnimationFrame(() => {
        if (
          field.isConnected &&
          visualTexPointerSelectingFields.has(field)
        ) {
          if (!visualTexPointerPlaceholderStyles.has(field)) {
            installPointerPlaceholderSnapshotStyle(field);
          }
          restorePointerPlaceholderSnapshotStyle(field);
        }
      });
    };
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

    const imeGuard = new ImeCompositionGuard();
    let physicalBackslashGuard: {
      kind: PhysicalBackslashInputKind;
      beforeValue: string;
      expiresAt: number;
    } | null = null;
    let suppressBackslashReplayUntil = 0;
    let backslashGuardTimer = 0;
    let pendingAutoExitSetting: InputBehaviorSettingKey | null = null;
    let pendingWrapperInput: {
      command: string;
      content: string;
      range: [number, number];
      anchorStyle: Readonly<Style>;
      visualCaret: WrapperCaretAnchor | null;
    } | null = null;
    let replacingPendingWrapperInput = false;
    let restoringRawCommandAnchor = false;
    let wrapperPlaceholderFrame = 0;
    let wrapperPlaceholderPass = 0;
    let wrapperMeasure: HTMLSpanElement | null = null;
    const clearPendingWrapperPlaceholderPosition = () => {
      host.style.removeProperty("--pending-wrapper-left");
      host.style.removeProperty("--pending-wrapper-top");
    host.style.removeProperty("--pending-wrapper-width");
    host.style.removeProperty("--pending-wrapper-height");
    delete host.dataset.pendingWrapperAnchorY;
    };
    const measurePendingWrapperPlaceholderPosition = () => {
      if (!field.isConnected || !pendingWrapperInput) return;
      const hostBounds = host.getBoundingClientRect();
    const anchor =
      pendingWrapperInput.visualCaret ?? captureWrapperCaretAnchor(field);
    if (!pendingWrapperInput.visualCaret && anchor) {
      pendingWrapperInput.visualCaret = anchor;
    }
    const markers = Array.from(
      field.shadowRoot?.querySelectorAll<HTMLElement>(
        ".visualtex-structural-placeholder-caret, .ML__caret, .ML__text-caret, .ML__latex-caret",
      ) ?? [],
      )
        .map((marker) => {
          const bounds = marker.getBoundingClientRect();
          return {
            left: bounds.right - hostBounds.left,
            centerY: bounds.top - hostBounds.top + bounds.height / 2,
            height: bounds.height,
          };
        })
        .filter((marker) => marker.height > 0)
        .sort(
          (first, second) =>
            Math.abs(first.centerY - (anchor?.centerY ?? first.centerY)) -
            Math.abs(second.centerY - (anchor?.centerY ?? second.centerY)),
    );
    const closestMarker = markers[0] ?? null;
    const verticalCaret = anchor ?? closestMarker;
    if (!verticalCaret) {
      clearPendingWrapperPlaceholderPosition();
      return;
    }

      if (!wrapperMeasure) {
        wrapperMeasure = document.createElement("span");
        wrapperMeasure.className = "pending-wrapper-measure";
        wrapperMeasure.setAttribute("aria-hidden", "true");
        host.append(wrapperMeasure);
      }
      const content = pendingWrapperInput.content;
      let renderedWidth = 0;
      if (content) {
        wrapperMeasure.style.fontSize = field.style.fontSize;
        wrapperMeasure.innerHTML = convertLatexToMarkup(
          pendingWrapperInput.command + "{" + content + "}",
          { defaultMode: "math" },
        );
        renderedWidth = wrapperMeasure.getBoundingClientRect().width;
      } else {
        wrapperMeasure.replaceChildren();
    }
    const frameWidth = Math.max(18, Math.ceil(renderedWidth + 10));
    const fallbackFrameLeft = content
      ? (closestMarker?.left ?? verticalCaret.left) - renderedWidth
      : closestMarker?.left ?? verticalCaret.left;
    const frameLeft = anchor?.left ?? fallbackFrameLeft;
    const frameCenterX = frameLeft + frameWidth / 2;
    const formulaFontSize =
      Number.parseFloat(field.style.fontSize) || BASE_FORMULA_FONT_SIZE;
    const minimumFrameHeight = Math.max(12, formulaFontSize * 0.52);
    const maximumFrameHeight = Math.max(
      minimumFrameHeight,
      formulaFontSize * 1.08,
    );
    const frameHeight = Math.max(
      minimumFrameHeight,
      Math.min(maximumFrameHeight, verticalCaret.height + 4),
    );
      host.style.setProperty("--pending-wrapper-left", `${frameCenterX}px`);
      host.style.setProperty(
        "--pending-wrapper-top",
        `${verticalCaret.centerY}px`,
      );
    host.style.setProperty("--pending-wrapper-width", `${frameWidth}px`);
    host.style.setProperty(
      "--pending-wrapper-height",
      `${frameHeight}px`,
    );
    host.dataset.pendingWrapperAnchorY = String(verticalCaret.centerY);

      wrapperPlaceholderPass += 1;
      if (
        wrapperPlaceholderPass < 3 &&
        host.classList.contains("has-pending-wrapper-placeholder")
      ) {
        wrapperPlaceholderFrame = window.requestAnimationFrame(
          measurePendingWrapperPlaceholderPosition,
        );
      }
    };
    const schedulePendingWrapperPlaceholderPosition = () => {
      window.cancelAnimationFrame(wrapperPlaceholderFrame);
      if (!host.classList.contains("has-pending-wrapper-placeholder")) {
        clearPendingWrapperPlaceholderPosition();
        return;
      }
      wrapperPlaceholderPass = 0;
      wrapperPlaceholderFrame = window.requestAnimationFrame(
        measurePendingWrapperPlaceholderPosition,
      );
    };
    const syncPendingWrapperPlaceholder = () => {
      const showPlaceholder = Boolean(pendingWrapperInput);
      host.classList.toggle("has-pending-wrapper-placeholder", showPlaceholder);
      if (showPlaceholder && pendingWrapperInput) {
        host.dataset.pendingWrapperCommand = pendingWrapperInput.command;
        host.dataset.pendingWrapperLength = String(
          Array.from(pendingWrapperInput.content).length,
        );
      } else {
        delete host.dataset.pendingWrapperCommand;
        delete host.dataset.pendingWrapperLength;
      }
      schedulePendingWrapperPlaceholderPosition();
    };
    const clearPendingWrapperInput = () => {
      pendingWrapperInput = null;
      delete field.dataset.pendingWrapperCommand;
      syncPendingWrapperPlaceholder();
    };
    const replacePendingWrapperInput = (
      command: string,
      content: string,
      trailingInput = "",
    ) => {
      const pending = pendingWrapperInput;
      if (!pending) return false;

      const [rangeStart, rangeEnd] = pending.range;
      field.mode = "math";
      field.selection = {
        ranges: [[rangeStart, rangeEnd]],
        direction: "forward",
      };
      replacingPendingWrapperInput = true;
      let inserted = false;
      try {
        inserted = field.insert(
          command + "{" + content + "}" + trailingInput,
          {
            mode: "math",
            format: "latex",
            insertionMode: "replaceSelection",
            selectionMode: "after",
            focus: true,
            scrollIntoView: false,
          },
        );
      } finally {
        replacingPendingWrapperInput = false;
      }
      if (!inserted) return false;

      field.applyStyle(pending.anchorStyle);
      pending.command = command;
      pending.content = content;
      if (!trailingInput) pending.range = [rangeStart, field.position];
      return true;
    };
    let compositionDeleteObserved = false;
    let suppressPostCompositionDeleteUntil = 0;


    const armPhysicalBackslashGuard = (
      kind: PhysicalBackslashInputKind,
      timeStamp: number,
    ) => {
      physicalBackslashGuard = {
        kind,
        beforeValue: field.value,
        expiresAt: timeStamp + 240,
      };
      window.clearTimeout(backslashGuardTimer);
      const expectedGuard = physicalBackslashGuard;
      backslashGuardTimer = window.setTimeout(() => {
        if (physicalBackslashGuard === expectedGuard) {
          physicalBackslashGuard = null;
        }
      }, 280);
    };

    const normalizeGuardedBackslashInput = (timeStamp: number) => {
      const guard = physicalBackslashGuard;
      if (!guard || timeStamp > guard.expiresAt) {
        physicalBackslashGuard = null;
        return;
      }
      const rawValue = field.value;
      const normalizedValue = normalizePhysicalBackslashInput(
        guard.beforeValue,
        rawValue,
        guard.kind,
      );
      if (normalizedValue === rawValue) return;

      const previousPosition = field.position;
      const removedCharacters = rawValue.length - normalizedValue.length;
      field.setValue(normalizedValue, {
        mode: "math",
        format: "latex",
        insertionMode: "replaceAll",
        selectionMode: "after",
        silenceNotifications: true,
      });
      const correctedPosition = Math.max(
        0,
        Math.min(field.lastOffset, previousPosition - removedCharacters),
      );
      field.position = correctedPosition;
      field.selection = {
        ranges: [[correctedPosition, correctedPosition]],
        direction: "none",
      };
      physicalBackslashGuard = null;
    };

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
      pendingAutoExitSetting = getCaretAutoExitSetting(field);
      propsRef.current.onCommitPending();
      imeGuard.compositionStart();
      compositionStartRef.current =
        lastSnapshotRef.current ?? captureFieldSnapshot(field);
    };
    const handleCompositionEnd = (event: CompositionEvent) => {
      const cancelledByCompositionDelete =
        event.data === "" && compositionDeleteObserved;
      imeGuard.compositionEnd(event.timeStamp);
      suppressPostCompositionDeleteUntil = cancelledByCompositionDelete
        ? event.timeStamp + 160
        : 0;
      compositionDeleteObserved = false;
      normalizeGuardedBackslashInput(event.timeStamp);
      const before =
        compositionStartRef.current ??
        lastSnapshotRef.current ??
        captureFieldSnapshot(field);

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
        pendingAutoExitSetting = null;
        compositionStartRef.current = null;
        syncFrameSize();
        return;
      }

      let after = captureFieldSnapshot(field);
      if (pendingAutoExitSetting && before.latex !== after.latex) {
        moveCaretThroughEnabledAutoExitContainers(
          field,
          propsRef.current.inputBehavior,
          pendingAutoExitSetting,
        );
        after = captureFieldSnapshot(field);
      }
      pendingAutoExitSetting = null;
      compositionStartRef.current = null;
      emitEdit(before, after, "composition", "keyboard");
      syncFrameSize();
    };
    const handleBeforeInput = (event: InputEvent) => {
      if (restoringRawCommandAnchor) return;
      if (
        event.inputType === "deleteContentBackward" &&
        !pendingWrapperInput &&
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
            restoreRawCommandAnchor(field, anchor);
          } finally {
            restoringRawCommandAnchor = false;
          }
          rawCommandAnchors.delete(field);
          markVisualTexStructuralPlaceholders(field);
          const after = captureFieldSnapshot(field);
          emitEdit(before, after, "delete-backward", "keyboard");
          propsRef.current.onInputActivity(field);
          syncFrameSize();
          return;
        }
      }
      if (
        pendingWrapperInput &&
        !replacingPendingWrapperInput &&
        !event.isComposing
      ) {
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
          const nextCommand =
            pendingWrapperInput.command === "\\mathcal" &&
            pendingWrapperInput.content.length === 0 &&
            /^[a-z]$/.test(wrappedInput)
              ? "\\mathscr"
              : pendingWrapperInput.command;
          const nextContent = pendingWrapperInput.content + wrappedInput;
          if (
            !replacePendingWrapperInput(
              nextCommand,
              nextContent,
              trailingInput,
            )
          ) {
            return;
          }
          field.dataset.pendingWrapperCommand = nextCommand;
          syncPendingWrapperPlaceholder();
          field.focus();
          if (autoExit) clearPendingWrapperInput();
          field.shadowRoot
            ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
            ?.focus({ preventScroll: true });
          const after = captureFieldSnapshot(field);
          emitEdit(before, after, "insert", "keyboard");
          syncFrameSize();
          return;
        }
        if (event.inputType === "deleteContentBackward") {
          if (pendingWrapperInput.content) {
            event.preventDefault();
            event.stopImmediatePropagation();
            const before = captureFieldSnapshot(field);
            const nextContent = Array.from(pendingWrapperInput.content)
              .slice(0, -1)
              .join("");
            if (
              !replacePendingWrapperInput(
                pendingWrapperInput.command,
                nextContent,
              )
            ) {
              return;
            }
            syncPendingWrapperPlaceholder();
            field.focus();
            const after = captureFieldSnapshot(field);
            emitEdit(before, after, "delete-backward", "keyboard");
            syncFrameSize();
            return;
          }
          clearPendingWrapperInput();
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
          pendingAutoExitSetting =
            getCaretAutoExitSetting(field) ?? pendingAutoExitSetting;
        } else {
          pendingAutoExitSetting = null;
        }
      }
      if (
        event.data === "\\" &&
        event.timeStamp <= suppressBackslashReplayUntil
      ) {
        pendingAutoExitSetting = null;
        event.preventDefault();
        event.stopPropagation();
      }
    };
    const handleInput = (event: Event) => {
      if (replacingPendingWrapperInput || restoringRawCommandAnchor) return;
      if (imeGuard.isComposing()) {
        if (
          event instanceof InputEvent &&
          event.inputType === "deleteCompositionText"
        ) {
          compositionDeleteObserved = true;
        }
        return;
      }
      normalizeGuardedBackslashInput(event.timeStamp);
      const before = lastSnapshotRef.current ?? captureFieldSnapshot(field);
      const directInputSetting =
        event instanceof InputEvent && isSingleDirectInput(event, field)
          ? getCaretAutoExitSetting(field)
          : null;
      const autoExitSetting = pendingAutoExitSetting ?? directInputSetting;
      if (
        autoExitSetting &&
        propsRef.current.inputBehavior[autoExitSetting]
      ) {
        moveCaretThroughEnabledAutoExitContainers(
          field,
          propsRef.current.inputBehavior,
          autoExitSetting,
        );
      }
      pendingAutoExitSetting = null;
      normalizeCompletedDifferentialDisplay(field);
      const after = captureFieldSnapshot(field);
      const inputType =
        event instanceof InputEvent ? event.inputType || "insertText" : "insertText";
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
        if (normalizeCompletedDifferentialDisplay(field)) {
          emitEdit(
            deferredBefore,
            captureFieldSnapshot(field),
            "replace",
            "keyboard",
          );
        }
        propsRef.current.onInputActivity(field);
        decorateNativeSuggestionPreviews();
      });
    };
    const handleSelectionChange = () => {
      markVisualTexStructuralPlaceholders(field);
      schedulePointerPlaceholderSnapshotStyle();
      if (imeGuard.isComposing() || !lastSnapshotRef.current) return;
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
      clearPendingWrapperInput();
      rawCommandAnchors.delete(field);
      propsRef.current.onCommitPending();
    };
    const confirmPendingWrapperInput = (event: KeyboardEvent) => {
      if (
        !pendingWrapperInput ||
        event.isComposing ||
        event.metaKey ||
        event.ctrlKey ||
        event.altKey ||
        event.key !== "Enter"
      ) {
        return false;
      }

      event.preventDefault();
      event.stopImmediatePropagation();
      const confirmedPosition = field.position;
      clearPendingWrapperInput();
      field.mode = "math";
      field.selection = {
        ranges: [[confirmedPosition, confirmedPosition]],
        direction: "none",
      };
      field.position = confirmedPosition;
      field.focus();
      field.shadowRoot
        ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
        ?.focus({ preventScroll: true });
      propsRef.current.onCommitPending();
      syncFrameSize();
      return true;
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

      const wrapperQuery = rawCommandQuery(field) || trailingCommandQuery(field);
      const selectedNativeCommand =
        field.dataset.pendingNativeSuggestion ||
        getVisibleNativeSuggestionItems().find((item) =>
          item.classList.contains("ML__popover__current"),
        )?.dataset.command ||
        "";
      const wrapperCommand =
        (wrapperQuery ? exactWrapperCommand(wrapperQuery) : null) ??
        (selectedNativeCommand
          ? exactWrapperCommand(selectedNativeCommand)
          : null);
      if (
        !wrapperCommand ||
        (field.mode !== "latex" && field.lastOffset !== 0)
      ) {
        return false;
      }

      event.preventDefault();
      event.stopImmediatePropagation();
      pendingAutoExitSetting = null;
      const before = captureFieldSnapshot(field);
      const anchor = rawCommandAnchors.get(field) ?? {
        latex: before.latex,
        selection: before.selection,
        position: field.position,
        visualCaret: captureWrapperCaretAnchor(field),
      };
      const anchorSelection = restoreRawCommandAnchor(field, anchor);
      const anchorRange = anchorSelection.ranges.at(-1) ?? [anchor.position, anchor.position];
      const rangeStart = Math.min(anchorRange[0], anchorRange[1]);
      rawCommandAnchors.delete(field);
      delete field.dataset.pendingNativeSuggestion;
      const styleAtAnchor =
        field.getElementInfo(rangeStart)?.style ??
        field.getElementInfo(Math.max(0, rangeStart - 1))?.style ??
        {};
      const anchorStyle: Readonly<Style> = {
        variant: styleAtAnchor.variant ?? "normal",
        variantStyle: styleAtAnchor.variantStyle ?? "",
        fontFamily: styleAtAnchor.fontFamily ?? "none",
        fontShape: styleAtAnchor.fontShape ?? "",
        fontSeries: styleAtAnchor.fontSeries ?? "",
      };
      pendingWrapperInput = {
        command: wrapperCommand.command,
        content: "",
        range: [rangeStart, rangeStart],
        anchorStyle,
        visualCaret: anchor.visualCaret,
      };
      field.dataset.pendingWrapperCommand = wrapperCommand.command;
      field.focus();
      field.shadowRoot
        ?.querySelector<HTMLElement>('[part="keyboard-sink"]')
        ?.focus({ preventScroll: true });
      const pendingReference = pendingWrapperInput;
      window.requestAnimationFrame(() => {
        window.requestAnimationFrame(() => {
          if (
            !field.isConnected ||
            !pendingReference ||
            pendingWrapperInput !== pendingReference
          ) {
            return;
          }
          if (!pendingReference.visualCaret) {
            pendingReference.visualCaret = captureWrapperCaretAnchor(field);
          }
          syncPendingWrapperPlaceholder();
        });
      });
      const after = captureFieldSnapshot(field);
      emitEdit(before, after, "replace", "candidate");
      syncFrameSize();
      return true;
    };
    const handleRawWrapperKeyDown = (event: KeyboardEvent) => {
      if (confirmPendingWrapperInput(event)) return;
      confirmRawWrapperCommand(event);
    };
    const handleWindowRawWrapperKeyDown = (event: KeyboardEvent) => {
      if (!event.composedPath().includes(field)) return;
      const isUnmodifiedPhysicalBackslash =
        event.code === "Backslash" &&
        !event.metaKey &&
        !event.ctrlKey &&
        !event.altKey;
      if (
        isUnmodifiedPhysicalBackslash &&
        event.key === "\\" &&
        !pendingWrapperInput &&
        !rawLatexInput(field)
      ) {
        rememberRawCommandAnchor(field);
      }
      if (confirmPendingWrapperInput(event)) return;
      confirmRawWrapperCommand(event);
    };
    const scheduleInputActivity = () => {
      window.requestAnimationFrame(() => {
        if (!field.isConnected) return;
        propsRef.current.onInputActivity(field);
        decorateNativeSuggestionPreviews();
      });
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      const isUnmodifiedPhysicalBackslash =
        event.code === "Backslash" &&
        !event.metaKey &&
        !event.ctrlKey &&
        !event.altKey;
      if (event.key === "Escape") rawCommandAnchors.delete(field);
      if (isUnmodifiedPhysicalBackslash && event.key === "、") {
        armPhysicalBackslashGuard("ideographic-comma", event.timeStamp);
        suppressBackslashReplayUntil = event.timeStamp + 180;
        event.preventDefault();
        event.stopPropagation();
        return;
      }
      if (isUnmodifiedPhysicalBackslash && event.key === "\\") {
        if (event.timeStamp <= suppressBackslashReplayUntil) {
          event.preventDefault();
          event.stopPropagation();
          return;
        }
        armPhysicalBackslashGuard("latin-backslash", event.timeStamp);
      }

      if (
        event.key === "Backspace" &&
        !pendingWrapperInput &&
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
            restoreRawCommandAnchor(field, anchor);
          } finally {
            restoringRawCommandAnchor = false;
          }
          rawCommandAnchors.delete(field);
          markVisualTexStructuralPlaceholders(field);
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

      const capturesDirectMathInput =
        !event.isComposing &&
        !event.metaKey &&
        !event.ctrlKey &&
        !event.altKey &&
        field.mode !== "latex" &&
        event.key !== "\\" &&
        Array.from(event.key).length === 1;
      if (capturesDirectMathInput) {
        pendingAutoExitSetting = getCaretAutoExitSetting(field);
      } else if (
        event.key !== "Shift" &&
        event.key !== "Control" &&
        event.key !== "Alt" &&
        event.key !== "Meta"
      ) {
        pendingAutoExitSetting = null;
      }

      if (confirmPendingWrapperInput(event)) return;
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
    const handlePointerSelectionEnd = () => {
      if (!visualTexPointerSelectingFields.delete(field)) return;
      field.classList.remove(visualTexPointerSelectingClass);
      removePointerPlaceholderSnapshotStyle(field);
      window.requestAnimationFrame(() => {
        if (!field.isConnected) return;
        markVisualTexStructuralPlaceholders(field);
      });
    };
    const handlePointerDown = (event: PointerEvent) => {
      if (event.button !== 0) return;
      installPointerPlaceholderSnapshotStyle(field);
      visualTexPointerSelectingFields.add(field);
      field.classList.add(visualTexPointerSelectingClass);
      schedulePointerPlaceholderSnapshotStyle();
      clearPendingWrapperInput();
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
    installVisualTexStructuralPlaceholderStyle(field);
    field.inlineShortcuts = {
      ...field.inlineShortcuts,
      ...visualTexUprightInlineShortcuts,
    };
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
    field.addEventListener("beforeinput", handleBeforeInput, true);
    field.addEventListener("input", handleInput);
    field.addEventListener("selection-change", handleSelectionChange);
    field.addEventListener("focus", handleFocus);
    field.addEventListener("blur", handleBlur);
    window.addEventListener("keydown", handleWindowRawWrapperKeyDown, true);
    field.addEventListener("keydown", handleKeyDown, true);
    const keyboardSink = field.shadowRoot?.querySelector<HTMLElement>(
      '[part="keyboard-sink"]',
    );
    keyboardSink?.addEventListener("keydown", handleRawWrapperKeyDown, true);
    keyboardSink?.addEventListener("input", scheduleInputActivity, true);
    keyboardSink?.addEventListener("keyup", scheduleInputActivity, true);
    field.addEventListener("paste", handlePaste, true);
    host.addEventListener("pointerdown", handlePointerDown, true);
    window.addEventListener("pointerup", handlePointerSelectionEnd, true);
    window.addEventListener("pointercancel", handlePointerSelectionEnd, true);
    const content = field.shadowRoot?.querySelector<HTMLElement>('[part="content"]');
    const resizeObserver = content
      ? new ResizeObserver(() => {
          syncFrameSize();
          schedulePendingWrapperPlaceholderPosition();
        })
      : null;
    const inputMutationObserver = field.shadowRoot
      ? new MutationObserver(() => {
          markVisualTexStructuralPlaceholders(field);
          schedulePointerPlaceholderSnapshotStyle();
          scheduleInputActivity();
          schedulePendingWrapperPlaceholderPosition();
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
    syncFrameSize();

    return () => {
      window.cancelAnimationFrame(resizeFrame);
      window.cancelAnimationFrame(pointerPlaceholderFrame);
      window.cancelAnimationFrame(wrapperPlaceholderFrame);
      window.clearTimeout(resizeTimer);
      window.clearTimeout(backslashGuardTimer);
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
      window.removeEventListener("keydown", handleWindowRawWrapperKeyDown, true);
      field.removeEventListener("keydown", handleKeyDown, true);
      keyboardSink?.removeEventListener("keydown", handleRawWrapperKeyDown, true);
      keyboardSink?.removeEventListener("input", scheduleInputActivity, true);
      keyboardSink?.removeEventListener("keyup", scheduleInputActivity, true);
      field.removeEventListener("paste", handlePaste, true);
      host.removeEventListener("pointerdown", handlePointerDown, true);
      window.removeEventListener("pointerup", handlePointerSelectionEnd, true);
      window.removeEventListener("pointercancel", handlePointerSelectionEnd, true);
      visualTexPointerSelectingFields.delete(field);
      field.classList.remove(visualTexPointerSelectingClass);
      removePointerPlaceholderSnapshotStyle(field);
      host.closest<HTMLElement>(".formula-line")?.style.removeProperty(
        "--formula-row-height",
      );
      propsRef.current.register(lineId, null);
      fieldRef.current = null;
      lastSnapshotRef.current = null;
      compositionStartRef.current = null;
      pendingAutoExitSetting = null;
      clearPendingWrapperInput();
      rawCommandAnchors.delete(field);
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
        const queryChanged = queryRef.current !== activeCommandQuery;
        queryRef.current = activeCommandQuery;
        setQuery(activeCommandQuery);
        if (queryChanged) selectSuggestionIndex(0);
        requestAnimationFrame(() => updatePopupPosition(field));
      } else {
        queryRef.current = "";
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

      normalizeCompletedDifferentialDisplay(field);
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
      const queryRange = activeQuery
        ? findTrailingCommandRange(field, activeQuery)
        : null;
      const replacesRawCommand = Boolean(
        activeQuery && !queryRange && rawAnchor,
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
        if (rawAnchor) restoreRawCommandAnchor(field, rawAnchor);
        else field.selection = originalSelection;
        const inserted = applyDiscreteFormulaMutation(
          targetLineId,
          field,
          historySource,
          () => {
            const hasPlaceholder = insertionTemplate.includes("\\placeholder{}");
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
              moveCaretThroughEnabledAutoExitContainers(
                field,
                inputBehavior,
                autoExitSetting,
              );
            }
            return inserted;
          },
        );
        if (inserted) rawCommandAnchors.delete(field);
        else field.selection = originalSelection;
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
        if (
          source === "candidate" &&
          !structuredSuggestionCommands.has(command.command)
        ) {
          dismissNativeSuggestionPopover(field);
        }
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

      if (
        event.key !== "ArrowDown" &&
        event.key !== "ArrowUp" &&
        event.key !== "Enter" &&
        event.key !== "Tab"
      ) {
        delete field.dataset.pendingNativeSuggestion;
      }

      const normalizedFieldValue = normalizeChineseLatex(field.value);
      const suppressedSuggestion = suppressedSuggestionRef.current;
      const suppressesCurrentTrailingCommand =
        suppressedSuggestion?.lineId === lineId &&
        suppressedSuggestion.value.trim() === normalizedFieldValue.trim();
      const rawCommandActive = hasRawLatexInput(field);
      const liveQuery =
        suppressesCurrentTrailingCommand || rawCommandActive
          ? ""
          : trailingCommandQuery(field);
      const candidateQuery =
        suppressesCurrentTrailingCommand || rawCommandActive
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
      const wrapperQuery =
        rawCommandQuery(field) || trailingCommandQuery(field) || candidateQuery;
      const wrapperCommand = wrapperQuery
        ? exactWrapperCommand(wrapperQuery)
        : null;
      const confirmsWrapperCommand =
        wrapperCommand &&
        !event.metaKey &&
        !event.ctrlKey &&
        !event.altKey &&
        (event.key === "Enter" ||
          event.key === "Tab" ||
          event.key === " " ||
          event.code === "Space");

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
          queryRef.current = "";
          setQuery("");
          selectSuggestionIndex(0);
          return;
        }
      }

      const nativeRecommendationVisible =
        document
          .getElementById("mathlive-suggestion-popover")
          ?.classList.contains("is-visible") ?? false;
      const pendingNativeSuggestion =
        field.dataset.pendingNativeSuggestion ?? "";
      if (
        !suppressesCurrentTrailingCommand &&
        !liveSuggestions.length &&
        (nativeRecommendationVisible || Boolean(pendingNativeSuggestion))
      ) {
        if (
          nativeRecommendationVisible &&
          (event.key === "ArrowDown" || event.key === "ArrowUp")
        ) {
          event.preventDefault();
          event.stopPropagation();
          const selectedNativeCommand = moveNativeSuggestionSelection(
            field,
            event.key === "ArrowDown" ? 1 : -1,
          );
          if (selectedNativeCommand) {
            field.dataset.pendingNativeSuggestion = selectedNativeCommand;
          }
          return;
        }
        if (
          event.key === "Enter" ||
          event.key === "Tab" ||
          event.key === " " ||
          event.code === "Space"
        ) {
          event.preventDefault();
          event.stopPropagation();
          // The remembered native command is inserted synchronously. Waiting
          // for MathLive's popover mutation here used to add up to one second
          // before the editor state and history caught up with the field.
          const committed = applyDiscreteFormulaMutation(
            lineId,
            field,
            "candidate",
            () => commitNativeSuggestion(field, pendingNativeSuggestion),
          );
          if (committed) {
            delete field.dataset.pendingNativeSuggestion;
            setQuery("");
            selectSuggestionIndex(0);
            focusLine(lineId, {
              latex: normalizeChineseLatex(field.value),
              selection: captureSelection(field),
            });
          }
          return;
        }
        if (event.key === "Escape") {
          delete field.dataset.pendingNativeSuggestion;
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
          event.stopPropagation();
          removeEmptyLine(index);
          return;
        }

        // While MathLive is collecting a non-empty raw LaTeX command such as
        // `\\mat`, its public formula value is intentionally still empty. Let
        // MathLive consume Backspace/Delete natively so one physical key
        // removes one raw command character instead of deleting the row.
        if (rawCommandInput || (field.mode === "latex" && !trulyEmpty)) return;

        // Preserve MathLive's native word/line deletion shortcuts such as
        // Option+Backspace and Command+Backspace on macOS.
        if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) {
          return;
        }

        event.preventDefault();
        event.stopPropagation();

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
      // A caret/selection refresh can briefly hide the candidate list before
      // restoring the same query. Do not reset the highlighted row during that
      // transient empty state; a genuinely new query is reset explicitly in
      // refreshSuggestionQuery().
      if (
        suggestions.length > 0 &&
        selectedIndex >= suggestions.length
      ) {
        selectSuggestionIndex(suggestions.length - 1);
      }
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
