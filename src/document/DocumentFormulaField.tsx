// The live editor of an open document formula (docs/document-editor/DESIGN.md
// §3): the row editor's own FormulaField, so typing behaves exactly as in the
// formula rows. The document owns history and leaving; this component turns
// the field's edits and keys into the session callbacks.
import { useCallback, useEffect, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState, type Ref } from "react";
import type { MathfieldElement } from "mathlive";
import {
  REPLAYED_KEY,
  type FormulaEntry,
  type FormulaExit,
  type FormulaKey,
  type FormulaSelection,
  type FormulaSessionCallbacks,
} from "../../shared/visualtex-document/src/index.ts";
import { FormulaField, type FormulaFieldEdit, type FormulaFieldProps } from "../editor/MathEditor";
import { normalizeChineseLatex } from "../editor/normalizeChineseLatex";
import { configureMathLiveCompletion } from "../editor/mathLiveKernelCommands";
import { editingModelOf, placeCaretAtPoint, visualTexEditingCore } from "../editor/visualTexLayoutRuntime";
import { topLevelStop } from "../editor/rowPointer";
import { VISUALTEX_ALIGNMENT_MARKER_LATEX } from "../editor/alignmentMarkers";
import {
  collapseMathLiveSelectionToEnd,
  insertMathLiveAlignmentPoint,
  insertMathLiveRowBreak,
  moveMathLiveToNextColumn,
  toggleMathLiveSelectionStyle,
} from "../editor/mathLiveKernelCommands";
import { useEditorStore } from "../stores/editorStore";
import { useFormulaHotkeyStore } from "../stores/formulaHotkeyStore";
import { matchFormulaHotkey, resolveFormulaHotkeyCommand } from "../shortcuts/formulaHotkeys";
import { greekLetterHotkeyCommandFromEvent, isGreekLetterHotkeyPrefix } from "../shortcuts/greekLetterHotkeys";
import type { LatexCommand } from "../types/command";
import { insertCommandIntoFormula } from "./formulaInsert";
import { setOpenDocumentFormula } from "./openDocumentFormula";

/** The row editor's formula size is 54px x zoom. */
const ROW_FORMULA_FONT_SIZE = 54;

export interface DocumentFormulaFieldApi {
  update(latex: string, selection: FormulaSelection | null): void;
  /** Report the field's LaTeX if the document does not have it yet. */
  flush(): void;
  typeKey(key: FormulaKey): void;
  focus(): void;
  /** The formula opened: the field takes the focus, the caret where `entry` says. */
  enter(entry: FormulaEntry): void;
  /** The formula closed: the field rests (no focus, nothing selected); it stays on the page. */
  rest(): void;
  /** A drag that began in the field left the formula: the field stops following the pointer. */
  stopPointerTracking(): void;
  /** Show a part of the formula selected (a selection through the text takes it); null: none. */
  showPart(range: [number, number] | null): void;
  stopAt(x: number, y: number): number | null;
  topLevelStop(offset: number, side: "start" | "end"): number;
  lastOffset(): number | null;
  selectStops(anchor: number, focus: number): void;
  anchorStop(): number | null;
  latexOf(range: [number, number]): string;
  withoutRange(range: [number, number]): string | null;
  styled(range: [number, number], enable: boolean): string | null;
  rangeBold(range: [number, number]): boolean | null;
  /** Ctrl+B in the formula: its selection bold, or not bold when all of it is; false without a selection. */
  toggleBold(): boolean;
}

interface Props {
  id: string;
  latex: string;
  display: boolean;
  entry: FormulaEntry;
  fontSize: number;
  callbacks: FormulaSessionCallbacks;
  apiRef: Ref<DocumentFormulaFieldApi>;
  /** Only shown, with a part selected (a selection through the text): not focused, no caret placed. */
  passive?: boolean;
  /** An image pasted into the formula: the app's OCR, the result going in at the formula's caret. */
  onPasteImage?: (file: File) => void;
}

function captureSelection(field: MathfieldElement): FormulaSelection {
  return {
    ranges: field.selection.ranges.map(([start, end]) => [start, end] as [number, number]),
    direction: (field.selection.direction ?? "none") as FormulaSelection["direction"],
  };
}

function hasRawLatexInput(field: MathfieldElement) {
  return Boolean(field.shadowRoot?.querySelector(".ML__raw-latex"));
}

type KernelField = Parameters<typeof visualTexEditingCore.caretRect>[0];

function kernelOf(field: MathfieldElement): KernelField | null {
  return (field as unknown as { _mathfield?: KernelField })._mathfield ?? null;
}

const DOCUMENT_FIELD_STYLE_ID = "visualtex-document-field-style";

/**
 * In a document the field is the formula's own size: none of the row editor's padding or minimum height, so
 * opening a formula moves neither the formula nor the text around it.
 */
function installDocumentFieldStyle(field: MathfieldElement) {
  const root = field.shadowRoot;
  if (!root || root.getElementById(DOCUMENT_FIELD_STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = DOCUMENT_FIELD_STYLE_ID;
  style.textContent = `
    .ML__container { min-height: 0 !important; padding: 0 !important; width: auto !important; }
    .ML__content { padding: 0 !important; overflow: visible !important; }
  `;
  root.append(style);
}

/**
 * Focus the field without scrolling: the document already scrolled the formula into view, and the kernel's
 * focus() would scroll the whole formula element into view again (sideways, for a wide display formula).
 */
function focusWithoutScroll(field: MathfieldElement) {
  const kernel = (field as unknown as { _mathfield?: { focus(options?: FocusOptions): void } })._mathfield;
  if (kernel) kernel.focus({ preventScroll: true });
  else field.focus();
}

function keyboardSink(field: MathfieldElement): HTMLElement {
  return field.shadowRoot?.querySelector<HTMLElement>('[part="keyboard-sink"]') ?? field;
}

/**
 * A key typed into the field the way the keyboard types into it (keydown, then keypress, beforeinput and input
 * for a character, then keyup, at the field's keyboard sink), so the editor's own handling applies — as
 * VisualTeX Anywhere replays keys.
 */
function typeKeyInto(field: MathfieldElement, key: FormulaKey) {
  const sink = keyboardSink(field);
  const init: KeyboardEventInit = {
    key: key.key,
    code: key.code,
    shiftKey: key.shift,
    altKey: key.alt,
    ctrlKey: key.ctrl,
    metaKey: key.meta,
    bubbles: true,
    cancelable: true,
    composed: true,
  };
  const down = new KeyboardEvent("keydown", init);
  // (the document does not take it for a key of its own to hand over: it bubbles through the document's view)
  Object.defineProperty(down, REPLAYED_KEY, { value: true });
  sink.dispatchEvent(down);
  if (key.text && !down.defaultPrevented) {
    const press = new KeyboardEvent("keypress", { ...init, charCode: key.text.codePointAt(0) ?? 0 });
    sink.dispatchEvent(press);
    if (!press.defaultPrevented) {
      const before = new InputEvent("beforeinput", { data: key.text, inputType: "insertText", bubbles: true, cancelable: true, composed: true });
      if (sink.dispatchEvent(before)) {
        sink.dispatchEvent(new InputEvent("input", { data: key.text, inputType: "insertText", bubbles: true, composed: true }));
      }
    }
  }
  sink.dispatchEvent(new KeyboardEvent("keyup", init));
}

/** Put the caret where the formula was entered. */
function applyEntry(field: MathfieldElement, entry: FormulaEntry) {
  const model = editingModelOf(field);
  if (!model) return;
  if (entry.kind === "point") {
    if (placeCaretAtPoint(field, entry.x, entry.y)) return;
    entry = { kind: "end" };
  }
  if (entry.kind === "selection") {
    field.selection = { ranges: entry.selection.ranges, direction: entry.selection.direction };
    return;
  }
  const stops = visualTexEditingCore.allowedStops(model);
  const stop = entry.kind === "start" ? stops[0] : stops[stops.length - 1];
  if (stop !== undefined) visualTexEditingCore.goToStop(model, stop);
  // a template (\begin{align}, a fraction): its first slot
  // a template (\begin{align}, a fraction): its first slot selected, so that typing fills it
  if (entry.kind === "start" && /\\placeholder\b/.test(field.value)) {
    for (let offset = 0; offset < Math.min(field.lastOffset, 400); offset++) {
      if (/^\s*\\placeholder\b/.test(field.getValue([offset, offset + 1], "latex"))) {
        field.selection = { ranges: [[offset, offset + 1]], direction: "forward" };
        break;
      }
    }
  }
}

export function DocumentFormulaField({ id, latex, display, entry, fontSize, callbacks, apiRef, passive = false, onPasteImage }: Props) {
  const [currentLatex, setCurrentLatex] = useState(latex);
  // a value the document gave the field (undo, a split) that the field has not taken yet: nothing to report until then
  const incomingRef = useRef<string | null>(null);
  const fieldRef = useRef<MathfieldElement | null>(null);
  const pendingSelection = useRef<FormulaSelection | null>(null);
  const callbacksRef = useRef(callbacks);
  callbacksRef.current = callbacks;
  // The formula's selection dragged out of it: the document takes it (formulas.ts FormulaDragOut).
  const selectionDrag = useMemo<NonNullable<FormulaFieldProps["selectionDrag"]>>(
    () => ({
      outside: (x, y, _source, start, end) => callbacksRef.current.dragOutAt?.(x, y, start, end) ?? false,
      dropOutside: (_source, start, end, latex, copy) => callbacksRef.current.dragOutDrop?.(start, end, latex, copy),
      end: () => callbacksRef.current.dragOutEnd?.(),
    }),
    [],
  );
  const entryRef = useRef(entry);
  // Keys pressed while the formula was opening, replayed in order once the field holds the focus.
  const pendingKeys = useRef<FormulaKey[]>([]);
  const readyRef = useRef(false);
  // The formula was left (or removed): later keys belong to the document.
  const leftRef = useRef(false);
  // The LaTeX the document has (the field's edits are reported on the kernel's input events).
  const reportedRef = useRef(latex);
  /** Report the field's LaTeX if the document does not have it yet (an edit whose input event is still due). */
  const flush = () => {
    const field = fieldRef.current;
    if (!field || incomingRef.current !== null) return;
    const value = normalizeChineseLatex(field.value);
    if (value === reportedRef.current) return;
    reportedRef.current = value;
    callbacksRef.current.edit(value, null, captureSelection(field));
  };
  const leaveFormula = (how: FormulaExit, x?: number) => {
    flush();
    leftRef.current = true;
    callbacksRef.current.exit(how, x);
  };
  // a "$" typed in the empty inline formula: the third of "$$$$", shown as nothing until the fourth
  const heldDollarRef = useRef(false);
  /** "$" in the formula: the closing dollar, or (in an empty inline formula) the third and fourth "$" of "$$$$". */
  const typedDollar = (field: MathfieldElement) => {
    const empty = !field.getValue("latex-without-placeholders").trim();
    // An empty display formula waits for its content.
    if (empty && displayRef.current) return;
    if (empty && !heldDollarRef.current) {
      heldDollarRef.current = true;
      return;
    }
    heldDollarRef.current = false;
    flush();
    leftRef.current = true;
    callbacksRef.current.dollar(empty);
  };
  const displayRef = useRef(display);
  const passiveRef = useRef(passive);
  // the part shown selected (set again after the field takes a new value)
  const partRef = useRef<[number, number] | null>(null);
  const showPart = (field: MathfieldElement, range: [number, number] | null) => {
    partRef.current = range;
    field.classList.toggle("has-visualtex-multi-line-selection", !!range);
    if (range) field.selection = { ranges: [range], direction: "forward" };
    else if (!field.selectionIsCollapsed) field.selection = { ranges: [[0, 0]], direction: "none" };
  };

  const language = useEditorStore((state) => state.language);
  const autoPairDelimiters = useEditorStore((state) => state.autoPairDelimiters);
  const formulaLetterFont = useEditorStore((state) => state.formulaLetterFont);
  const formulaChineseFont = useEditorStore((state) => state.formulaChineseFont);
  const inputBehavior = useEditorStore((state) => state.inputBehavior);

  useImperativeHandle(apiRef, () => ({
    update(next, selection) {
      reportedRef.current = next;
      pendingSelection.current = selection;
      if (next !== currentLatex) incomingRef.current = next;
      setCurrentLatex(next);
    },
    focus() {
      passiveRef.current = false;
      const field = fieldRef.current;
      if (!field) return;
      setOpenDocumentFormula({ field, display: displayRef.current, mutate: (change) => mutateRef.current(field, change) });
      focusWithoutScroll(field);
    },
    enter(entry) {
      const field = fieldRef.current;
      entryRef.current = entry;
      passiveRef.current = false;
      if (field) activateRef.current(field, entry);
    },
    rest() {
      passiveRef.current = true;
      leftRef.current = false;
      readyRef.current = true;
      pendingKeys.current = [];
      greekPrefixRef.current = false;
      const field = fieldRef.current;
      if (!field) return;
      flush();
      setOpenDocumentFormula(null, field);
      if (field.shadowRoot?.activeElement) field.blur();
      showPart(field, null);
    },
    stopPointerTracking() {
      fieldRef.current?.visualTexStopPointerTracking?.();
    },
    showPart(range) {
      if (fieldRef.current) showPart(fieldRef.current, range);
      else partRef.current = range;
    },
    stopAt(x, y) {
      const field = fieldRef.current;
      const kernel = field ? kernelOf(field) : null;
      return kernel ? visualTexEditingCore.hitTest(kernel, x, y) : null;
    },
    topLevelStop(offset, side) {
      return fieldRef.current ? topLevelStop(fieldRef.current, offset, side) : offset;
    },
    lastOffset() {
      return fieldRef.current?.lastOffset ?? null;
    },
    selectStops(anchor, focus) {
      fieldRef.current?.visualTexSelectStops(anchor, focus);
    },
    anchorStop() {
      const field = fieldRef.current;
      return field ? editingModelOf(field)?.anchor ?? null : null;
    },
    latexOf(range) {
      const field = fieldRef.current;
      if (!field) return "";
      return normalizeChineseLatex(field.getValue(Math.min(...range), Math.max(...range), "latex"));
    },
    withoutRange(range) {
      const result = fieldRef.current?.visualTexDropResult({ remove: range });
      return result ? normalizeChineseLatex(result.latex) : null;
    },
    styled(range, enable) {
      const latex = fieldRef.current?.visualTexStyledResult(range, "bold", enable);
      return latex ? normalizeChineseLatex(latex) : latex ?? null;
    },
    rangeBold(range) {
      return fieldRef.current?.visualTexRangeFont(range)?.bold ?? null;
    },
    toggleBold() {
      const field = fieldRef.current;
      if (!field || field.selectionIsCollapsed) return false;
      return mutate(field, () => toggleMathLiveSelectionStyle(field, "bold", field.selection));
    },
    flush,
    typeKey(key) {
      const field = fieldRef.current;
      // a field at rest (a formula just made again, not yet entered) keeps the keys for when it is entered
      if (!field || !readyRef.current || passiveRef.current) {
        pendingKeys.current.push(key);
        return;
      }
      typeKeyInto(field, key);
    },
  }));

  // After the field took a document change (undo/redo), its caret goes where it was at that state.
  useLayoutEffect(() => {
    incomingRef.current = null;
    const field = fieldRef.current;
    const selection = pendingSelection.current;
    pendingSelection.current = null;
    if (field && selection) field.selection = { ranges: selection.ranges, direction: selection.direction };
    // a field showing a part shows it again in its new value (the offsets stay: bold changes no atom)
    else if (field && partRef.current) showPart(field, partRef.current);
  }, [currentLatex]);

  /** Run a kernel mutation as one formula edit. */
  const mutate = useCallback((field: MathfieldElement, change: () => boolean) => {
    const beforeLatex = normalizeChineseLatex(field.value);
    const before = captureSelection(field);
    if (!change()) return false;
    const afterLatex = normalizeChineseLatex(field.value);
    field.resetUndo();
    if (afterLatex !== beforeLatex) {
      reportedRef.current = afterLatex;
      setCurrentLatex(afterLatex);
      callbacksRef.current.edit(afterLatex, before, captureSelection(field));
    }
    return true;
  }, []);

  const mutateRef = useRef(mutate);
  mutateRef.current = mutate;
  // Ctrl+G (Greek letter shortcut) waits for the letter, as in the formula rows.
  const greekPrefixRef = useRef(false);
  /** A formula command (shortcut, Greek letter) at the caret, as the formula rows insert it. */
  const insertCommand = (field: MathfieldElement, command: LatexCommand) =>
    insertCommandIntoFormula({ field, display: displayRef.current, mutate: (change) => mutateRef.current(field, change) }, command);
  // only this formula's own field: by the time an earlier formula goes, the next one may already be open
  useEffect(
    () => () => {
      if (fieldRef.current) setOpenDocumentFormula(null, fieldRef.current);
    },
    [],
  );

  /** Enter in a display formula: two formulas at the caret (at the formula's top level), else a text line after it. */
  const splitAtCaret = (field: MathfieldElement) => {
    const [start, end] = field.selectionIsCollapsed
      ? [field.position, field.position]
      : [Math.min(...field.selection.ranges.flat()), Math.max(...field.selection.ranges.flat())];
    const last = field.lastOffset;
    const atTop = topLevelStop(field, start, "start") === start && topLevelStop(field, end, "end") === end;
    const before = normalizeChineseLatex(field.getValue(0, start, "latex")).trim();
    const after = normalizeChineseLatex(field.getValue(end, last, "latex")).trim();
    if (!atTop || start <= 0 || end >= last || !before || !after || !callbacksRef.current.split) {
      leaveFormula("new-paragraph");
      return;
    }
    flush();
    leftRef.current = true;
    callbacksRef.current.split(before, after);
  };

  const handleKeyDown = useCallback(
    (event: KeyboardEvent, field: MathfieldElement) => {
      if (event.isComposing || event.keyCode === 229) return;
      const stop = () => {
        event.preventDefault();
        event.stopImmediatePropagation();
      };
      const key = event.key.toLocaleLowerCase();
      const mod = (event.ctrlKey || event.metaKey) && !event.altKey;
      // The formula rows' shortcuts: Greek letters (Ctrl+G, then the letter) and the formula shortcuts of Settings.
      const rawInput = hasRawLatexInput(field);
      if (greekPrefixRef.current) {
        if (event.code === "ShiftLeft" || event.code === "ShiftRight") return;
        greekPrefixRef.current = false;
        const greek = event.key === "Escape" ? null : greekLetterHotkeyCommandFromEvent(event);
        if (event.key === "Escape" || greek) {
          stop();
          if (greek) insertCommand(field, greek);
          return;
        }
      }
      if (!rawInput && isGreekLetterHotkeyPrefix(event)) {
        stop();
        greekPrefixRef.current = true;
        return;
      }
      const hotkey = rawInput ? null : matchFormulaHotkey(event, useFormulaHotkeyStore.getState().bindings);
      if (hotkey) {
        const command = resolveFormulaHotkeyCommand(hotkey.target);
        if (command) {
          stop();
          insertCommand(field, command);
          return;
        }
      }
      if (mod && ((key === "z" && event.shiftKey) || (key === "y" && !event.shiftKey))) {
        stop();
        callbacksRef.current.redo();
        return;
      }
      if (mod && key === "z") {
        stop();
        callbacksRef.current.undo();
        return;
      }
      // Ctrl+B: the selection bold, every symbol (\bm), or not bold when all of it is
      if (mod && key === "b" && !event.shiftKey) {
        stop();
        if (!field.selectionIsCollapsed) mutate(field, () => toggleMathLiveSelectionStyle(field, "bold", field.selection));
        return;
      }
      const model = editingModelOf(field);
      const rawLatex = hasRawLatexInput(field);
      if (event.key === "Escape" && model && visualTexEditingCore.escapeAction(model, event, rawLatex) === "collapse") {
        stop();
        collapseMathLiveSelectionToEnd(field);
        return;
      }
      // Ctrl+Enter: out of the formula, to the text after it (Typora's way out of a math block).
      if (event.key === "Enter" && mod && !event.shiftKey) {
        stop();
        leaveFormula("after");
        return;
      }
      if (event.key === "Enter" && model) {
        const action = visualTexEditingCore.enterAction(model, event, rawLatex);
        if (action === "complete") return;
        stop();
        // As in the formula rows: Shift+Enter is a new line inside the formula (its lines aligned rows,
        // left-aligned until an & aligns them; a formula written as gathered stays gathered), Enter ends the line —
        // the display formula is split at the caret, or a new text line follows it; an inline formula is left.
        if (!display) leaveFormula("after");
        else if (action === "row-break") {
          // an empty display formula takes no empty line
          if (field.getValue("latex-without-placeholders").trim())
            mutate(field, () => insertMathLiveRowBreak(field, /^\s*\\begin\{gathered\}/.test(field.value) ? "gathered" : "aligned"));
        } else if (action === "case-row") mutate(field, () => field.executeCommand("addRowAfter"));
        else splitAtCaret(field);
        return;
      }
      if (rawLatex || field.mode === "latex") return;
      if (event.key === "$" && !mod && !event.altKey && field.mode === "math") {
        stop();
        typedDollar(field);
        return;
      }
      if (event.key === "&" && model) {
        // in a display formula, & at the top level is an alignment point (formula mode's rows): the next line
        // makes the formula aligned rows, split at it
        const action = visualTexEditingCore.ampersandAction(model, event, rawLatex, display);
        if (action === "alignment-marker") {
          if (mutate(field, () => field.insert(VISUALTEX_ALIGNMENT_MARKER_LATEX, { mode: "math", format: "latex", insertionMode: "replaceSelection", selectionMode: "after" }))) stop();
        } else if (action === "alignment-point") {
          if (mutate(field, () => insertMathLiveAlignmentPoint(field))) stop();
        } else if (action === "next-column") {
          let moved = false;
          mutate(field, () => {
            const before = JSON.stringify([field.value, captureSelection(field)]);
            moveMathLiveToNextColumn(field);
            moved = JSON.stringify([field.value, captureSelection(field)]) !== before;
            return moved;
          });
          if (moved) stop();
        }
        return;
      }
      if (event.key === "Backspace" && !event.shiftKey && !mod) {
        const visible = field.getValue("latex-without-placeholders").trim();
        if (!visible) {
          stop();
          leftRef.current = true;
          callbacksRef.current.remove();
        }
      }
    },
    [display, mutate],
  );

  const typedDollarRef = useRef(typedDollar);
  typedDollarRef.current = typedDollar;
  // "$" that an input method types ("￥" with Chinese punctuation) reaches the field as text, not as a key.
  const detachDollar = useRef<(() => void) | null>(null);
  /** The formula opened: the field takes the focus, the caret where `entry` says; keys typed meanwhile are replayed. */
  const activate = (field: MathfieldElement, entry: FormulaEntry) => {
    passiveRef.current = false;
    leftRef.current = false;
    readyRef.current = false;
    heldDollarRef.current = false;
    partRef.current = null;
    field.classList.remove("has-visualtex-multi-line-selection");
    setOpenDocumentFormula({ field, display: displayRef.current, mutate: (change) => mutateRef.current(field, change) });
    focusWithoutScroll(field);
    applyEntry(field, entry);
    requestAnimationFrame(() => {
      if (!field.isConnected || entry.kind !== "point") return;
      applyEntry(field, entry);
    });
    // The kernel takes the focus only after a timeout (and drops keys until then): keys pressed meanwhile reach
    // the document, which hands them over (typeKey); they are replayed in order right after the kernel's focus,
    // before any later key.
    const replay = () => {
      if (readyRef.current || !field.isConnected) return;
      readyRef.current = true;
      const keys = pendingKeys.current;
      pendingKeys.current = [];
      for (let i = 0; i < keys.length; i++) {
        // A key left the formula (Tab, an arrow at its edge, Enter, the closing "$"): the document takes the rest.
        // (Leaving puts the field to rest at once, which clears leftRef: passiveRef tells it then.)
        if (leftRef.current || passiveRef.current) {
          callbacksRef.current.handBack(keys.slice(i));
          return;
        }
        typeKeyInto(field, keys[i]);
      }
    };
    if (field.shadowRoot?.activeElement) queueMicrotask(replay);
    else field.addEventListener("focusin", () => queueMicrotask(replay), { once: true, capture: true });
  };
  const activateRef = useRef(activate);
  activateRef.current = activate;
  const register = useCallback((_lineId: string, field: MathfieldElement | null, expected?: MathfieldElement) => {
    fieldRef.current = field;
    detachDollar.current?.();
    detachDollar.current = null;
    if (!field) {
      setOpenDocumentFormula(null, expected);
      return;
    }
    const onBeforeInput = (event: Event) => {
      const input = event as InputEvent;
      if (input.inputType !== "insertText" || (input.data !== "$" && input.data !== "￥")) {
        // anything else typed: a held "$" is no longer the start of "$$$$"
        if (input.data) heldDollarRef.current = false;
        return;
      }
      if (field.mode !== "math" || hasRawLatexInput(field)) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      typedDollarRef.current(field);
    };
    field.addEventListener("beforeinput", onBeforeInput, true);
    // the candidates as the formula rows have them: ranked by what was chosen before, and every choice remembered
    const configure = () => {
      const { usage, personalize, suggestionCount } = useEditorStore.getState();
      configureMathLiveCompletion(field, { usage, personalize, count: suggestionCount });
    };
    configure();
    const onAccepted = (event: Event) => {
      const detail = (event as CustomEvent<{ id?: string; query?: string }>).detail;
      if (!detail?.id) return;
      useEditorStore.getState().recordCommand(detail.id, detail.query ?? "", "candidate");
      configure();
    };
    field.addEventListener("visualtex-command-accepted", onAccepted);
    detachDollar.current = () => {
      field.removeEventListener("beforeinput", onBeforeInput, true);
      field.removeEventListener("visualtex-command-accepted", onAccepted);
    };
    // exactly the formula's picture: no room kept around it for the caret (drawn over what is around it)
    const kernel = kernelOf(field) as { visualTexNoCaretRoom?: boolean } | null;
    if (kernel) kernel.visualTexNoCaretRoom = true;
    // An inline formula is set in text style, as in the text (and as its static markup).
    if (!displayRef.current) field.defaultMode = "inline-math";
    installDocumentFieldStyle(field);
    // the formula as the field reads it: only an edit is reported back (not the field's own form of the same LaTeX)
    if (normalizeChineseLatex(field.value) !== reportedRef.current && reportedRef.current === latex) reportedRef.current = normalizeChineseLatex(field.value);
    if (passiveRef.current) {
      // only shown: the part a selection through the text takes, once the engine has drawn it
      readyRef.current = true;
      if (partRef.current) showPart(field, partRef.current);
      requestAnimationFrame(() => {
        if (field.isConnected && partRef.current) showPart(field, partRef.current);
      });
      return;
    }
    // The engine draws the field on its next frame; the caret is placed in that geometry.
    activateRef.current(field, entryRef.current);
  }, []);

  const onEdit = useCallback((edit: FormulaFieldEdit) => {
    reportedRef.current = edit.afterLatex;
    setCurrentLatex(edit.afterLatex);
    callbacksRef.current.edit(edit.afterLatex, edit.beforeSelection, edit.afterSelection);
  }, []);

  const caretX = (field: MathfieldElement) => {
    const kernel = kernelOf(field);
    const rect = kernel ? visualTexEditingCore.caretRect(kernel) : null;
    return rect ? rect.left : undefined;
  };

  return (
    <FormulaField
      lineId={id}
      index={0}
      latex={currentLatex}
      zoom={fontSize / ROW_FORMULA_FONT_SIZE}
      formulaRowVerticalInset={0}
      language={language}
      formulaLetterFont={formulaLetterFont}
      formulaChineseFont={formulaChineseFont}
      autoPairDelimiters={autoPairDelimiters}
      inputBehavior={inputBehavior}
      persistentTypingStyle={{ bold: false, italic: null, color: null, backgroundColor: null }}
      readOnly={false}
      freshExternalSync={false}
      register={register}
      onEdit={onEdit}
      onInputActivity={() => undefined}
      onSelectionChange={() => undefined}
      onFocus={() => undefined}
      onCommitPending={() => undefined}
      onKeyDown={(_index, event, field) => handleKeyDown(event, field)}
      onVerticalExit={(_index, direction, field) => {
        // Up/Down with nowhere to go inside the formula: the text line above/below, at the caret's x.
        leaveFormula(direction < 0 ? "up" : "down", caretX(field));
        return true;
      }}
      onHorizontalExit={(_index, direction) => {
        leaveFormula(direction < 0 ? "before" : "after");
        return true;
      }}
      selectionDrag={selectionDrag}
      onPasteImage={onPasteImage ? (file) => onPasteImage(file) : undefined}
    />
  );
}
