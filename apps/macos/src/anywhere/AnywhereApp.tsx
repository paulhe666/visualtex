import { useCallback, useEffect, useLayoutEffect, useRef, useState, type CSSProperties } from "react";
import { formatFormulaLinesUniversal } from "../../../shared/visualtex-app/src/clipboard/LatexCopyService";
import { MathEditor, type MathEditorHandle, type MathEditorInsertionTarget } from "../../../shared/visualtex-app/src/editor/MathEditor";
import {
  editingModelOf,
  redrawOpenFormulas,
  setFormulaCaretBand,
  setFormulaLineLayout,
  visualTexEditingCore,
} from "../../../shared/visualtex-app/src/editor/visualTexLayoutRuntime";
import { useHistoryManager } from "../../../shared/visualtex-app/src/history/EditorSession";
import type { DocumentSnapshot } from "../../../shared/visualtex-app/src/history/historyTypes";
import { useEditorStore } from "../../../shared/visualtex-app/src/stores/editorStore";
import { applyDocumentTheme } from "../../../shared/visualtex-app/src/themeSync";
import { useAnywhereHotkeyStore } from "../../../shared/visualtex-app/src/stores/anywhereHotkeyStore";
import { useFormulaRowHotkeyStore } from "../../../shared/visualtex-app/src/stores/formulaRowHotkeyStore";
import {
  anywhereHotkeyLabel,
  matchAnywhereHotkey,
  type AnywhereHotkeyAction,
} from "../../../shared/visualtex-app/src/shortcuts/anywhereHotkeys";
import {
  formulaMarkdown,
  formulaPng,
  formulaUnicode,
  nextOutputMode,
  opaqueHex,
  OUTPUT_MODE_LABELS,
  rememberedOutputMode,
  rememberOutputMode,
  warmOutputMode,
  type AnywhereOutputMode,
} from "../../../shared/visualtex-app/src/anywhere/outputModes";
import {
  ANYWHERE_ACTIVATE_EVENT,
  ANYWHERE_COMPOSE_EVENT,
  ANYWHERE_KEY_EVENT,
  ANYWHERE_PIECES_EVENT,
  ANYWHERE_STYLE_EVENT,
  ANYWHERE_SUSPEND_EVENT,
  captureAnywhereScreenshot,
  readAnywhereClipboard,
  commitAnywhere,
  commitAnywhereImage,
  dismissAnywhere,
  isAnywhereActivation,
  isAnywhereCompose,
  isAnywhereKey,
  isAnywherePieces,
  isAnywhereStyle,
  isAnywhereSuspension,
  listenAnywhere,
  logAnywhere,
  reportAnywhereLayout,
  reportAnywhereMarked,
  reportAnywhereMode,
  reportAnywhereReady,
  type AnywhereActivation,
  type AnywhereKey,
  type AnywhereLayout,
} from "../../../shared/visualtex-app/src/anywhere/anywhereBridge";

/** MathEditor's formula font size at zoom 1 (BASE_FORMULA_FONT_SIZE). */
const EDITOR_BASE_FONT_PX = 54;
/** Room around the strip for its shadow (anywhere.css: .anywhere-strip). */
const PANEL_MARGIN = 8;
/** The candidate table: under the formula, its left edge where the command
 * being typed starts (anywhere.css lays it out, at the formula's size); the
 * panel grows to hold it and its shadow. */
const CANDIDATE_GAP = 5;
const CANDIDATE_INSET = 1;
const CANDIDATE_MARGIN = 14;
const COPIED_NOTICE_MS = 1800;
/** How long the output mode shows under the formula when it opens or changes. */
const MODE_TAG_MS = 1600;
/** The output mode's preview, cut to this many characters. */
const MODE_PREVIEW_CHARS = 48;

const NOTICE = {
  accessibility: "未授权辅助功能 · ↩ 复制",
  copied: "已复制 · ⌘V 粘贴",
  inputMethod: "输入源中没有 VisualTeX · 系统设置 › 键盘 › 输入法",
  reopen: (app: string) => `${app} 需重新打开 · 之后公式写在输入框里`,
} as const;

type CoreMathfield = Parameters<typeof visualTexEditingCore.caretRect>[0];

/** x-height over font size of the formula's fonts (TeX's Computer Modern
 * and Latin Modern; the Times profiles are a little larger). */
const FORMULA_X_HEIGHT = 0.431;
/** The field's font when it does not tell (system UI fonts). */
const FIELD_X_HEIGHT = 0.52;

/** The smallest formula, CSS pixels: the native side sizes it from believable
 * reports only (else from the system's text size), so this only keeps a very
 * small field's formula legible. */
const FORMULA_MIN_PX = 12;

/** The formula reads as part of the host's text: at the field's font size
 * when it tells it (the formula is drawn in its line), scaled so that its
 * letters are as tall as the field's (TeX's fonts have a smaller x-height
 * than UI fonts at the same size); else from its line height. */
function formulaZoom(caretHeight: number, fontSize?: number | null, xHeight?: number | null) {
  if (fontSize && Number.isFinite(fontSize) && fontSize > 0) {
    const fieldX = xHeight && Number.isFinite(xHeight) && xHeight > 0 ? xHeight : FIELD_X_HEIGHT;
    const scale = Math.min(1.35, Math.max(1, fieldX / FORMULA_X_HEIGHT));
    return Math.min(56, Math.max(FORMULA_MIN_PX, fontSize * scale)) / EDITOR_BASE_FONT_PX;
  }
  const height = Number.isFinite(caretHeight) && caretHeight > 0 ? caretHeight : 18;
  const fontPx = Math.min(40, Math.max(FORMULA_MIN_PX, height * 0.92));
  return fontPx / EDITOR_BASE_FONT_PX;
}

/** The host field's line (its caret) in CSS pixels, and how far it reaches
 * below the baseline. */
type FieldLine = { height: number; descent: number };

function fieldLineOf(height?: number | null, descent?: number | null): FieldLine | null {
  if (!height || !Number.isFinite(height) || height <= 0) return null;
  const below = descent && Number.isFinite(descent) && descent > 0 && descent < height ? descent : height * 0.22;
  return { height, descent: below };
}

/** The formula's caret as tall as the host field's line, standing on the
 * formula's baseline as the field's caret stands on its own: in em of the
 * formula's font (the engine scales it in scripts). Null: the editor's caret. */
function caretBandFor(line: FieldLine | null, formulaPx: number) {
  if (!line || !(formulaPx > 0)) return null;
  const up = (line.height - line.descent) / formulaPx;
  const down = line.descent / formulaPx;
  return up > 0.3 && up < 2.5 && down >= 0 && down < 1 ? { up, down } : null;
}

/** The text colour of the system's appearance, for a field that does not
 * tell its own (Chrome, Electron). */
function systemTextColor() {
  const dark = window.matchMedia?.("(prefers-color-scheme: dark)").matches ?? false;
  return dark ? "rgba(255, 255, 255, 0.92)" : "rgba(0, 0, 0, 0.9)";
}

/** #rrggbbaa as a CSS colour. */
function cssColor(color?: string | null) {
  if (!color || !/^#[0-9a-f]{8}$/i.test(color)) return null;
  const channel = (index: number) => parseInt(color.slice(index, index + 2), 16);
  return `rgba(${channel(1)}, ${channel(3)}, ${channel(5)}, ${(channel(7) / 255).toFixed(3)})`;
}

/** sp per em of the layout engine (10pt at 65536 sp/pt). */
const SP_PER_EM = 10 * 65536;

type FormulaLines = { originX: number; pieces: { width: number; x: number; y: number }[] };

/** The first row's formula set on several lines (the engine's geometry), if it is. */
function formulaLines(): FormulaLines | null {
  const mathfield = (firstField() as unknown as { _mathfield?: { visualTexGeometry?: { lines?: FormulaLines } } } | null)?._mathfield;
  return mathfield?.visualTexGeometry?.lines ?? null;
}

/**
 * The drawn formula in the strip's field: its origin (where its first
 * piece starts: on a formula set on several lines, lines that start further
 * left lie left of it), its baseline (the first line's: the bottom of the
 * layout's strut, whose top is the box's top), the width its ink takes and
 * its font size.
 */
function formulaBox() {
  const latex = firstField()?.shadowRoot?.querySelector<HTMLElement>(".ML__latex");
  if (!latex) return null;
  const rect = latex.getBoundingClientRect();
  const style = getComputedStyle(latex);
  const inkLeft = parseFloat(style.marginLeft) || 0;
  const inkRight = parseFloat(style.marginRight) || 0;
  const fontPx = parseFloat(style.fontSize) || 0;
  const strut = latex.querySelector<HTMLElement>(":scope > .ML__strut");
  const baseline = strut ? strut.getBoundingClientRect().bottom : rect.top + rect.height * 0.75;
  const origin = ((formulaLines()?.originX ?? 0) / SP_PER_EM) * fontPx;
  return { left: rect.left - inkLeft + origin, baseline, width: rect.width + inkLeft + inkRight - origin, fontPx };
}

/**
 * Covers over the host field's held space, in its background colour, under the formula: the field
 * draws its composition underline (Chrome in its accent colour, whatever the input method asks),
 * the selected anchor's highlight and its caret there; the held space is only spaces, so nothing of
 * the field's own text is covered. One per line of the formula, the field's line tall, from where
 * the field put the line's first piece to where the line's formula ends (the last line: to where
 * the held space ends, when the field tells). `line`: the field's line height and descent.
 */
/** How long the last line's cover keeps a width the formula lost (a deleted character): the field
 *  takes the narrower space a moment after the formula is drawn narrower (scheduleInline), and its
 *  underline and caret stood past the formula meanwhile. */
const SHRINK_COVER_MS = 200;
/** Spare room in the in-line panel, right of and below the formula and its covers: they grow into it
 *  in the frame the formula grows, without waiting for the window to be made larger. */
const PANEL_SLACK = 96;

function placeCovers(
  layer: HTMLElement | null,
  background: string | null,
  line: FieldLine | null,
  places: [number, number][],
  start: [number, number] | null,
  end: [number, number] | null,
  keep: { y: number; right: number } | null,
  caretLine: [number, number] | null,
): { right: number; bottom: number; last: { y: number; right: number } | null } | null {
  if (!layer) return null;
  const box = formulaBox();
  const lines = formulaLines();
  const covers: { left: number; right: number; top: number; bottom: number }[] = [];
  let lastLine: { y: number; right: number } | null = null;
  if (box && background && line && !places.length) {
    // the field holds a space before it has told where: the formula's line from the anchor, at
    // least a filler's width (its underline and caret are there)
    const baseline = box.baseline;
    const right = box.left + Math.max(box.width, line.height * 0.75) + 3;
    covers.push({ left: box.left - 3, right, top: baseline - (line.height - line.descent) - 1.5, bottom: baseline + line.descent + 1 });
  } else if (box && lines && background && line && places.length) {
    const px = box.fontPx / SP_PER_EM;
    const byLine = new Map<number, { left: number; right: number }>();
    // (a piece the field has no place for is drawn after the one before it: held with that one, its
    // last piece while narrower than a filler, or not held yet)
    lines.pieces.forEach((piece) => {
      const y = Math.round(piece.y * px * 4) / 4;
      const left = box.left + piece.x * px;
      const right = left + piece.width * px;
      const known = byLine.get(y);
      if (known) {
        known.left = Math.min(known.left, left);
        known.right = Math.max(known.right, right);
      } else byLine.set(y, { left, right });
    });
    const ys = [...byLine.keys()].sort((a, b) => a - b);
    ys.forEach((y, index) => {
      const span = byLine.get(y)!;
      // the first line from the anchor (the selected hair space before the formula)
      // (not past it: the letter before the formula ends there, and lost its last pixel under the cover)
      const left = index === 0 ? Math.min(span.left, box.left + (start?.[0] ?? -3)) - 0.25 : span.left - 1;
      // a line the field broke after holds its last piece to about a space past the formula; the
      // last line's held space ends where the field says (else a little past the formula)
      const last = index === ys.length - 1;
      // (the field's caret, a point or two wide, stands at the held space's end: never past the
      // formula's on its last line, FormulaSpace)
      let right = span.right + (last ? 3 : Math.max(4, line.height * 0.6));
      if (last && end && Math.abs(end[1] - y) < 2) right = Math.max(right, box.left + end[0] + 3);
      if (last) {
        lastLine = { y, right: right - box.left };
        // a width just lost stays covered until the field has taken the narrower space
        if (keep && Math.abs(keep.y - y) < 0.5) right = Math.max(right, box.left + keep.right);
      }
      const baseline = box.baseline + y;
      // (a field's caret is drawn a pixel taller than its line: Chrome)
      covers.push({ left, right, top: baseline - (line.height - line.descent) - 1.5, bottom: baseline + line.descent + 1 });
      // The field's own caret stands where the held space ends, under the formula, and may be a few
      // points taller than its line (WeChat's showed above the cover as a green dash, 2026-11-11):
      // a cover of its own, a caret wide, taller than the line's.
      if (last && end && Math.abs(end[1] - y) < 2) {
        const at = box.left + end[0];
        const over = Math.max(3, line.height * 0.2);
        let top = baseline - (line.height - line.descent) - over;
        let bottom = baseline + line.descent + over * 0.7;
        // (as tall as the field has the space's last character: its caret is; WebKit sets the blanks
        // in a taller font than the text, and its caret showed above the cover as a dot)
        if (caretLine) {
            const low = baseline + line.descent + caretLine[0];
            top = Math.min(top, low - caretLine[1] - 2);
            bottom = Math.max(bottom, low + 1.5);
        }
        covers.push({ left: at - 2, right: at + 3, top, bottom });
      }
    });
  }
  while (layer.children.length > covers.length) layer.lastElementChild?.remove();
  while (layer.children.length < covers.length) layer.appendChild(document.createElement("div"));
  // the panel holds the covers too (one may reach past the formula, to the end of the held space)
  const extent = covers.length
    ? {
        right: Math.max(...covers.map((cover) => cover.right)),
        bottom: Math.max(...covers.map((cover) => cover.bottom)),
        last: lastLine,
      }
    : null;
  covers.forEach((cover, index) => {
    const element = layer.children[index] as HTMLElement;
    element.style.cssText =
      `position:fixed;pointer-events:none;background:${background};left:${cover.left.toFixed(2)}px;` +
      `top:${cover.top.toFixed(2)}px;width:${(cover.right - cover.left).toFixed(2)}px;height:${(cover.bottom - cover.top).toFixed(2)}px`;
  });
  return extent;
}

/** The widths (points) of the pieces the host field holds for the formula. */
function pieceWidths(box: { fontPx: number } | null) {
  const lines = formulaLines();
  if (!box || !lines) return [];
  return lines.pieces.map((piece) => Math.round((piece.width / SP_PER_EM) * box.fontPx * 100) / 100);
}

/** The line of the formula the editor's caret is on, from the first line (points, y down): the
 *  line of the item of the formula's top level the caret is in (or after). */
function caretLineOffset(box: { fontPx: number } | null) {
  type Atom = { parent?: Atom };
  type Model = { root: Atom; position: number; at(offset: number): Atom | undefined; offsetOf(atom: Atom): number };
  const mathfield = (firstField() as unknown as {
    _mathfield?: { model?: Model; visualTexGeometry?: { lines?: FormulaLines; stops: ({ baseline: number } | null)[] } };
  } | null)?._mathfield;
  const model = mathfield?.model;
  const geometry = mathfield?.visualTexGeometry;
  if (!box || !model || !geometry?.lines) return 0;
  let top = model.at(model.position);
  while (top?.parent && top.parent !== model.root) top = top.parent;
  const stop = top && top.parent === model.root ? geometry.stops[model.offsetOf(top)] : null;
  return stop ? (stop.baseline / SP_PER_EM) * box.fontPx : 0;
}

/**
 * The strip's one formula row keeps this id across sessions: the editor's
 * resident field reports its edits under the id it was mounted with (a new
 * id would need a remount, see OfficeDialogApp's refreshLayout).
 */
const FORMULA_LINE_ID = "anywhere-formula";

function freshDocument(mode: "inline" | "display" = "inline") {
  useEditorStore.getState().replaceDocumentState({
    title: "",
    lines: [{ id: FORMULA_LINE_ID, latex: "", mode }],
    activeLineId: FORMULA_LINE_ID,
    formulaAlignment: "left",
    selectionByLineId: {},
  });
}

function currentDocument(): DocumentSnapshot {
  const { title, lines, activeLineId, formulaAlignment } = useEditorStore.getState();
  return { title, lines, activeLineId, formulaAlignment, selectionByLineId: {} };
}

function hasFormula() {
  return useEditorStore.getState().lines.some((line) => line.latex.trim());
}

/** The strip's input follows the system's light or dark appearance, as an
 * input method's composition window does, not the app's theme. */
function applySystemAppearance() {
  const dark = window.matchMedia?.("(prefers-color-scheme: dark)").matches ?? false;
  applyDocumentTheme(dark ? "dark" : "light");
}

function stripFields(): HTMLElement[] {
  return Array.from(document.querySelectorAll<HTMLElement>(".anywhere-strip math-field"));
}

/** The first row's field: it sits on the host's line (the panel's anchor). */
function firstField(): HTMLElement | null {
  return stripFields()[0] ?? null;
}

/** The field being edited (the formula may have several rows): the focused
 *  one, else the one the kernel draws as focused, else the first. */
function stripField(): HTMLElement | null {
  const fields = stripFields();
  if (fields.length <= 1) return fields[0] ?? null;
  const focused = fields.find((field) => field === document.activeElement || field.contains(document.activeElement));
  return focused ?? fields.find((field) => field.shadowRoot?.querySelector(".ML__focused")) ?? fields[0];
}

/** What the strip shows: the editor's row, or the field itself if the row
 * has not caught up (it never should). */
function composedFormulaLines() {
  const lines = useEditorStore.getState().lines.filter((line) => line.latex.trim());
  if (lines.length) return lines;
  const live = (stripField() as unknown as { value?: unknown } | null)?.value;
  if (typeof live !== "string" || !live.trim()) return [];
  void logAnywhere("commit read the field: the editor row was empty");
  return [{ id: FORMULA_LINE_ID, latex: live, mode: "inline" as const }];
}

function suggestionsVisible() {
  return Boolean(
    document.getElementById("mathlive-suggestion-popover")?.classList.contains("is-visible"),
  );
}

/** The strip's offset from the panel's top (anywhere.css: --anywhere-strip-top):
 * more than PANEL_MARGIN while the candidate bar sits above the formula. */
function setStripTop(strip: HTMLElement, top: number) {
  const value = top > PANEL_MARGIN ? `${top}px` : "";
  if (strip.style.getPropertyValue("--anywhere-strip-top") !== value) {
    if (value) strip.style.setProperty("--anywhere-strip-top", value);
    else strip.style.removeProperty("--anywhere-strip-top");
  }
}

/** Places the candidate bar under the formula (anywhere.css reads the
 * position; the kernel's own, centred on the caret, is for a list), or over
 * it when the screen has no room below the line (the strip then moves down
 * by the bar, and the panel up, so the formula stays on its line), and
 * returns where its far edges are, for the panel's size. */
function placeCandidateBar(stripElement: HTMLElement) {
  const popover = document.getElementById("mathlive-suggestion-popover");
  if (!popover) return null;
  const root = stripField()?.shadowRoot;
  const typed = root?.querySelector(".ML__raw-latex")?.getBoundingClientRect();
  // Right under the formula's ink (its depth included), not the strip's
  // line box, as an input method's bar sits under the line.
  const inkOf = () => {
    const drawn = root?.querySelector(".ML__latex")?.getBoundingClientRect();
    return drawn && drawn.height > 0 ? drawn : stripElement.getBoundingClientRect();
  };
  let ink = inkOf();
  let left = Math.max(0, Math.round((typed && typed.height > 0 ? typed.left : ink.left) - CANDIDATE_INSET));
  popover.style.setProperty("--anywhere-candidates-left", `${left}px`);
  const box = popover.getBoundingClientRect();
  // The bar's room under the line, on the screen: the same whichever side
  // the bar is on (the panel moves with the strip).
  const screenWithTop = window.screen as Screen & { availTop?: number };
  const screenBottom = (screenWithTop.availTop ?? 0) + window.screen.availHeight;
  const roomBelow = screenBottom - (window.screenY + ink.bottom + CANDIDATE_GAP);
  const above = box.height + CANDIDATE_MARGIN > roomBelow;
  const stripTop = parseFloat(stripElement.style.getPropertyValue("--anywhere-strip-top")) || PANEL_MARGIN;
  // Over the formula: the bar from the panel's margin, the formula's ink
  // (the strip's box, when it has one) under it.
  const boxed = !stripElement.classList.contains("is-input-method");
  const upperEdge = () => (boxed ? stripElement.getBoundingClientRect().top : ink.top);
  const edgeFromStrip = upperEdge() - stripElement.getBoundingClientRect().top;
  const wantedStripTop = above
    ? Math.max(PANEL_MARGIN, Math.ceil(PANEL_MARGIN + box.height + CANDIDATE_GAP - edgeFromStrip))
    : PANEL_MARGIN;
  if (wantedStripTop !== stripTop) {
    setStripTop(stripElement, wantedStripTop);
    ink = inkOf();
  }
  const top = above
    ? Math.round(upperEdge() - CANDIDATE_GAP - box.height)
    : Math.round(ink.bottom + CANDIDATE_GAP);
  popover.style.setProperty("--anywhere-candidates-top", `${top}px`);
  // Kept on the screen when the formula is near its right edge.
  const screenLeft = (window.screen as Screen & { availLeft?: number }).availLeft ?? 0;
  const roomRight = screenLeft + window.screen.availWidth - window.screenX - CANDIDATE_MARGIN;
  if (left + box.width > roomRight && roomRight > box.width) {
    left = Math.round(roomRight - box.width);
    popover.style.setProperty("--anywhere-candidates-left", `${left}px`);
  }
  return { right: left + box.width, bottom: top + box.height };
}

/** The row being edited, as the main window's $ / $$ marker says it. */
function activeRowLabel() {
  const { lines, activeLineId } = useEditorStore.getState();
  const line = lines.find((item) => item.id === activeLineId) ?? lines[0];
  const rows = lines.length > 1 ? ` ${lines.indexOf(line) + 1}/${lines.length}` : "";
  return `${line?.mode === "inline" ? "行内" : "独立"}${rows}`;
}

/** What Return would insert in `mode`, for the mode's tag. */
async function outputPreview(mode: AnywhereOutputMode, lines: ReturnType<typeof composedFormulaLines>) {
  let text: string;
  if (mode === "image") return "PNG";
  if (mode === "unicode") text = (await formulaUnicode(lines.map((line) => line.latex))).text;
  else if (mode === "markdown") text = formulaMarkdown(lines);
  else {
    text = formatFormulaLinesUniversal(lines, useEditorStore.getState().latexFormatProfile);
  }
  const chars = Array.from(text);
  return chars.length > MODE_PREVIEW_CHARS ? `${chars.slice(0, MODE_PREVIEW_CHARS - 1).join("")}…` : text;
}

/** The shortcut (Settings > Keyboard shortcuts) a key is while a formula
 *  is open, if any. */
function formulaShortcut(key: { code: string; ctrl?: boolean; alt?: boolean; shift?: boolean; meta?: boolean }) {
  return matchAnywhereHotkey(useAnywhereHotkeyStore.getState().bindings, key);
}

const OUTPUT_SHORTCUTS: Partial<Record<AnywhereHotkeyAction, AnywhereOutputMode>> = {
  "anywhere-output-source": "source",
  "anywhere-output-markdown": "markdown",
  "anywhere-output-unicode": "unicode",
  "anywhere-output-image": "image",
};

/** The output mode's tag: under the formula's start, while the candidates
 *  are not up; returns its far edges for the panel's size. */
function placeModeTag(strip: DOMRect) {
  const tag = document.querySelector<HTMLElement>(".anywhere-mode-tag");
  if (!tag) return null;
  // under the last row, at the formula's start
  const first = firstField()?.shadowRoot?.querySelector(".ML__latex")?.getBoundingClientRect();
  const drawn = stripFields().at(-1)?.shadowRoot?.querySelector(".ML__latex")?.getBoundingClientRect();
  const left = Math.max(0, Math.round((first && first.height > 0 ? first.left : strip.left) - 1));
  const top = Math.round((drawn && drawn.height > 0 ? drawn.bottom : strip.bottom) + CANDIDATE_GAP);
  tag.style.left = `${left}px`;
  tag.style.top = `${top}px`;
  const box = tag.getBoundingClientRect();
  return { right: left + box.width, bottom: top + box.height };
}

/** A dropped composition takes its candidate list with it (as the kernel's
 * hideSuggestionPopover does). */
function hideSuggestions() {
  const popover = document.getElementById("mathlive-suggestion-popover");
  popover?.classList.remove("is-visible");
  popover?.setAttribute("aria-hidden", "true");
}

/** A command is being typed in the LaTeX box: Enter completes it. */
function hasRawLatexInput(field: HTMLElement) {
  return Boolean(field.shadowRoot?.querySelector(".ML__raw-latex"));
}

/** Nothing typed yet, a half-typed command included (the editor's row only
 * hears of a command once it is complete). */
function fieldIsEmpty(field: HTMLElement) {
  const value = (field as unknown as { value?: unknown }).value;
  return (typeof value !== "string" || value.trim() === "") && !hasRawLatexInput(field);
}

/** The text the host field gets for the formula (the copy format's inline
 * delimiters), empty while there is no formula. */
function formulaText(latexFormatProfile: Parameters<typeof formatFormulaLinesUniversal>[1]) {
  const lines = useEditorStore.getState().lines.filter((line) => line.latex.trim());
  if (!lines.length) return "";
  return formatFormulaLinesUniversal(lines, latexFormatProfile);
}

/**
 * A key the input method took from the host field, typed into the formula
 * the way the keyboard types into it: keydown, then keypress, beforeinput
 * and input for a character, then keyup, all at the field's keyboard sink,
 * so the editor's own handling (and this page's Enter and Escape) apply.
 */
function typeKey(key: AnywhereKey) {
  const field = stripField();
  const sink = field?.shadowRoot?.querySelector<HTMLElement>('[part="keyboard-sink"]') ?? field;
  if (!sink) return;
  const init: KeyboardEventInit = {
    key: key.key,
    code: key.code,
    shiftKey: key.shift,
    altKey: key.alt,
    ctrlKey: key.ctrl,
    metaKey: key.meta,
    repeat: key.repeat,
    bubbles: true,
    cancelable: true,
    composed: true,
  };
  const down = new KeyboardEvent("keydown", init);
  sink.dispatchEvent(down);
  if (key.text && !down.defaultPrevented) {
    const press = new KeyboardEvent("keypress", { ...init, charCode: key.text.codePointAt(0) ?? 0 });
    sink.dispatchEvent(press);
    if (!press.defaultPrevented) {
      const before = new InputEvent("beforeinput", {
        data: key.text,
        inputType: "insertText",
        bubbles: true,
        cancelable: true,
        composed: true,
      });
      if (sink.dispatchEvent(before)) {
        sink.dispatchEvent(
          new InputEvent("input", { data: key.text, inputType: "insertText", bubbles: true, composed: true }),
        );
      }
    }
  }
  sink.dispatchEvent(new KeyboardEvent("keyup", init));
}

/**
 * The input method's pinyin in the formula's text, as the keyboard's input
 * method would give it to the field: compositionstart once, compositionupdate
 * with the spelling, compositionend with the chosen characters (or nothing).
 */
function compose(phase: "update" | "end", text: string, composing: { current: boolean }) {
  const field = stripField();
  const sink = field?.shadowRoot?.querySelector<HTMLElement>('[part="keyboard-sink"]') ?? field;
  if (!sink) return;
  const event = (type: string, data: string) =>
    sink.dispatchEvent(new CompositionEvent(type, { data, bubbles: true, cancelable: true, composed: true }));
  if (!composing.current && (phase === "update" || text)) {
    event("compositionstart", "");
    composing.current = true;
  }
  if (phase === "update") {
    event("compositionupdate", text);
  } else if (composing.current) {
    event("compositionend", text);
    composing.current = false;
  }
}

/** The editor's caret is in text (\text{}), not in math. */
function caretInText() {
  const field = stripField() as unknown as { mode?: unknown } | null;
  return field?.mode === "text";
}

/** The editor's caret from the drawn formula's start, in points. */
function caretOffset(box: { left: number } | null) {
  const field = stripField();
  const mathfield = (field as unknown as { _mathfield?: CoreMathfield } | null)?._mathfield;
  const caret = mathfield ? visualTexEditingCore.caretRect(mathfield) : null;
  if (!box || !caret || !Number.isFinite(caret.left)) return 0;
  // on a later line of a formula set on several lines the caret may stand left of the formula's start
  return formulaLines() ? caret.left - box.left : Math.max(0, caret.left - box.left);
}

function sameLayout(a: AnywhereLayout | null, b: AnywhereLayout) {
  return (
    a !== null &&
    Math.abs(a.width - b.width) < 0.5 &&
    Math.abs(a.height - b.height) < 0.5 &&
    Math.abs(a.anchorX - b.anchorX) < 0.5 &&
    Math.abs(a.anchorY - b.anchorY) < 0.5
  );
}

/**
 * VisualTeX Anywhere's composition strip: the formula being typed at the
 * host application's caret, with VisualTeX's candidates below it. Enter
 * types it into the host, Escape drops it.
 */
export function AnywhereApp() {
  const editorRef = useRef<MathEditorHandle>(null);
  const stripRef = useRef<HTMLDivElement>(null);
  const historyManager = useHistoryManager();
  const lines = useEditorStore((state) => state.lines);
  const activeLineId = useEditorStore((state) => state.activeLineId);
  const latexCodeFormat = useEditorStore((state) => state.latexCodeFormat);
  const [zoom, setZoom] = useState(() => formulaZoom(18));
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;
  /** The host field's line: the formula's caret is as tall. */
  const [fieldLine, setFieldLine] = useState<FieldLine | null>(null);
  const fieldLineRef = useRef<FieldLine | null>(null);
  fieldLineRef.current = fieldLine;
  const lastSizeLogRef = useRef("");
  const [notice, setNotice] = useState<string | null>(null);
  const [outputMode, setOutputMode] = useState<AnywhereOutputMode>("source");
  /** The output mode's tag under the formula (with a preview of the output),
   *  shown for a moment when the formula opens or the mode changes. */
  const [modeTag, setModeTag] = useState<
    { mode: AnywhereOutputMode; row: string; preview: string; status?: string } | null
  >(null);
  const outputModeRef = useRef<AnywhereOutputMode>("source");
  const appNameRef = useRef("");
  const modeTagTimerRef = useRef(0);
  const [session, setSession] = useState(0);
  /** VisualTeX's input method types this session's formula: keys from the
   * input method, the formula drawn in the field's line over the space the
   * field holds for it. */
  const [inputMethod, setInputMethod] = useState(false);
  const [textColor, setTextColor] = useState<string | null>(null);
  const inputMethodRef = useRef(false);
  const lastMarkedRef = useRef<string | null>(null);
  /** Where the host field put the formula's pieces (points from the first's start), and where its
   *  held space ends. */
  const piecePlacesRef = useRef<[number, number][]>([]);
  const pieceStartRef = useRef<[number, number] | null>(null);
  const pieceEndRef = useRef<[number, number] | null>(null);
  const pieceCaretRef = useRef<[number, number] | null>(null);
  /** The last line's cover widths of the last moments (placeCovers keeps the widest of them for
   *  SHRINK_COVER_MS), the timer that lets it shrink after, and the frame the field is told in
   *  (scheduleInline). */
  const coverWidthsRef = useRef<{ at: number; y: number; right: number }[]>([]);
  const shrinkTimerRef = useRef<number | null>(null);
  const inlineFrameRef = useRef<number | null>(null);
  /** The panel's last layout report (its window is set to the new size when it has been answered). */
  const layoutReportRef = useRef<Promise<unknown> | null>(null);
  const coverRef = useRef<HTMLDivElement>(null);
  const lastModeRef = useRef<boolean | null>(null);
  const composingRef = useRef(false);
  const generationRef = useRef(0);
  const committingRef = useRef(false);
  const lastLayoutRef = useRef<AnywhereLayout | null>(null);
  /** The field's text and background colours (#rrggbb) when known: the
   *  Image output is drawn in them. */
  const fieldInkRef = useRef<string | null>(null);
  const fieldBackgroundRef = useRef<string | null>(null);
  /** Where the formula starts inside the panel: the empty formula's caret. */
  const formulaStartRef = useRef<number | null>(null);

  /** Shows the output mode under the formula for a moment, with what Return
   *  would insert. */
  const showModeTag = useCallback((mode: AnywhereOutputMode) => {
    window.clearTimeout(modeTagTimerRef.current);
    const generation = generationRef.current;
    const lines = composedFormulaLines();
    const row = activeRowLabel();
    setModeTag({ mode, row, preview: "" });
    modeTagTimerRef.current = window.setTimeout(() => setModeTag(null), MODE_TAG_MS);
    if (!lines.length) return;
    void outputPreview(mode, lines)
      .then((preview) => {
        if (generationRef.current !== generation || outputModeRef.current !== mode) return;
        setModeTag((tag) => (tag && tag.mode === mode ? { ...tag, preview } : tag));
      })
      .catch(() => undefined);
  }, []);
  const showModeTagRef = useRef(showModeTag);
  showModeTagRef.current = showModeTag;

  /** An output mode (its shortcut, or the next one), remembered for this
   *  application. */
  const selectOutputMode = useCallback((mode: AnywhereOutputMode) => {
    outputModeRef.current = mode;
    setOutputMode(mode);
    rememberOutputMode(appNameRef.current, mode);
    warmOutputMode(mode);
    showModeTag(mode);
    void logAnywhere(`output mode ${mode} app=${appNameRef.current}`);
  }, [showModeTag]);

  /** Screenshot OCR: the area the user selects, recognized with the main
   *  window's OCR settings, goes in at the formula's caret. */
  const ocrBusyRef = useRef(false);
  const ocrStatus = useCallback((text: string | null) => {
    window.clearTimeout(modeTagTimerRef.current);
    setModeTag(text ? { mode: outputModeRef.current, row: "", preview: "", status: text } : null);
  }, []);

  /** Recognizes an image (a screenshot, a pasted image) with the main
   *  window's OCR settings and puts the formula at `target`. */
  const recognizeIntoFormula = useCallback(
    async (image: File, target: MathEditorInsertionTarget, generation: number) => {
      ocrStatus("正在识别公式…");
      const { fileToOcrRequest, recognizeFormulaImage, normalizeOcrFormulaLines, DEFAULT_OCR_MODEL, OCR_MODELS } =
        await import("../ocr/ocrService");
      let model = DEFAULT_OCR_MODEL;
      try {
        const stored = window.localStorage.getItem("visualtex.ocr.model");
        if (OCR_MODELS.some((item) => item.id === stored)) model = stored as typeof model;
      } catch {
        // the default model
      }
      const result = await recognizeFormulaImage(await fileToOcrRequest(image, model));
      if (generationRef.current !== generation) return;
      const latex = normalizeOcrFormulaLines(result.formulas).join("\n");
      if (!latex || !editorRef.current?.insertLatexAt(target, latex, "ocr")) {
        throw new Error("OCR 没有返回可用公式");
      }
      void logAnywhere(`ocr inserted chars=${latex.length}`);
      ocrStatus(null);
    },
    [ocrStatus],
  );

  /** Runs `work` as the formula's one OCR at a time, its failure shown in the tag. */
  const withOcr = useCallback(
    async (work: (target: MathEditorInsertionTarget, generation: number) => Promise<void>) => {
      const generation = generationRef.current;
      if (!generation || ocrBusyRef.current) return;
      const target = editorRef.current?.captureInsertionTarget();
      if (!target) {
        void logAnywhere("ocr: no insertion target");
        return;
      }
      ocrBusyRef.current = true;
      try {
        await work(target, generation);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        void logAnywhere(`ocr failed: ${message}`);
        ocrStatus(`OCR 失败：${message}`);
        modeTagTimerRef.current = window.setTimeout(() => setModeTag(null), 3000);
      } finally {
        ocrBusyRef.current = false;
      }
    },
    [ocrStatus],
  );

  /** Screenshot OCR: the area the user selects. */
  const runOcr = useCallback(
    () =>
      withOcr(async (target, generation) => {
        ocrStatus("框选公式区域（Esc 取消）");
        const data = await captureAnywhereScreenshot();
        if (!data || generationRef.current !== generation) {
          ocrStatus(null);
          return;
        }
        const { quickOcrCaptureToFile } = await import("../ocr/quickOcr");
        await recognizeIntoFormula(quickOcrCaptureToFile({ dataBase64: data, extension: "png" }), target, generation);
      }),
    [withOcr, ocrStatus, recognizeIntoFormula],
  );

  /** ⌘V with the input method: an image on the clipboard is recognized into
   *  the formula, text goes in as the main window pastes it. */
  const pasteIntoFormula = useCallback(async () => {
    const generation = generationRef.current;
    if (!generation) return;
    let clipboard;
    try {
      clipboard = await readAnywhereClipboard();
    } catch (error) {
      void logAnywhere(`paste failed: ${error instanceof Error ? error.message : String(error)}`);
      return;
    }
    if (generationRef.current !== generation) return;
    if (clipboard.kind === "text") {
      const target = editorRef.current?.captureInsertionTarget();
      if (target) editorRef.current?.insertLatexAt(target, clipboard.text, "paste");
      void logAnywhere(`paste text chars=${clipboard.text.length}`);
      return;
    }
    if (clipboard.kind !== "image") return;
    const { dataBase64, extension } = clipboard;
    void logAnywhere(`paste image ${extension}`);
    await withOcr(async (target, inGeneration) => {
      const { quickOcrCaptureToFile } = await import("../ocr/quickOcr");
      await recognizeIntoFormula(quickOcrCaptureToFile({ dataBase64, extension }), target, inGeneration);
    });
  }, [withOcr, recognizeIntoFormula]);

  /** An image pasted into the strip (no input method): recognized likewise. */
  const pastedImage = useCallback(
    (file: File) => {
      void withOcr((target, generation) => recognizeIntoFormula(file, target, generation));
    },
    [withOcr, recognizeIntoFormula],
  );

  /** A formula shortcut: an output mode, the next one, or OCR. */
  const runFormulaShortcut = useCallback(
    (action: AnywhereHotkeyAction) => {
      void logAnywhere(`shortcut ${action}`);
      if (action === "anywhere-next-output") selectOutputMode(nextOutputMode(outputModeRef.current));
      else if (action === "anywhere-row-display" || action === "anywhere-row-inline" || action === "anywhere-row-toggle") {
        const mode = action === "anywhere-row-display" ? "display" : action === "anywhere-row-inline" ? "inline" : "toggle";
        if (editorRef.current?.setActiveLineMode(mode)) {
          // the tag says what the row now is (the row-shape effect)
        } else {
          ocrStatus("多行公式保持独立公式");
          modeTagTimerRef.current = window.setTimeout(() => setModeTag(null), MODE_TAG_MS);
        }
      }
      else if (action === "anywhere-ocr") void runOcr();
      else if (OUTPUT_SHORTCUTS[action]) selectOutputMode(OUTPUT_SHORTCUTS[action]!);
    },
    [selectOutputMode, runOcr, ocrStatus],
  );

  const startNewFormula = useCallback((mode: "inline" | "display" = "inline") => {
    hideSuggestions();
    // A half-typed command lives in the field only (the row is still empty),
    // so emptying the row would leave it there: remount the field instead.
    const live = (firstField() as unknown as { value?: unknown } | null)?.value;
    const leftover = stripFields().some((field) => field.shadowRoot?.querySelector(".ML__raw-latex")) ||
      (typeof live === "string" && live.trim() !== "");
    freshDocument(mode);
    historyManager.clear();
    if (leftover) editorRef.current?.refreshLayout();
  }, [historyManager]);

  /** The input method keeps the formula's text (for a field that ends the
   *  composition itself) and its caret (pinyin candidates go under it). */
  const reportInline = useCallback(() => {
    const generation = generationRef.current;
    if (!generation || !inputMethodRef.current) return;
    const box = formulaBox();
    const width = box ? Math.ceil(box.width + 1) : 0;
    const caretX = Math.round(caretOffset(box));
    const caretY = Math.round(caretLineOffset(box));
    const pieces = pieceWidths(box);
    const text = formulaText(useEditorStore.getState().latexFormatProfile);
    const key = `${width}\u0000${caretX}\u0000${caretY}\u0000${pieces.join(",")}\u0000${text}`;
    if (key !== lastMarkedRef.current) {
      lastMarkedRef.current = key;
      void reportAnywhereMarked(generation, text, width, caretX, caretY, pieces).catch((error) =>
        console.error("Unable to hold the formula's space in the field", error),
      );
    }
    const inText = caretInText();
    if (inText !== lastModeRef.current) {
      lastModeRef.current = inText;
      void reportAnywhereMode(generation, inText).catch(() => undefined);
    }
  }, []);

  // The field is told the formula's new widths once the panel has drawn the formula and its cover as
  // they are now (two frames on): told at once, the field drew its wider space a frame before the
  // panel drew the formula, and its underline and caret showed past the formula at every key
  // (2026-11-10). Its space is never wider than the formula (FormulaSpace), so it grows under the
  // cover drawn already.
  // (A window made larger for them is set first: its layout report answered and a frame on.)
  const scheduleInline = useCallback(() => {
    if (inlineFrameRef.current !== null) return;
    const report = () => {
      inlineFrameRef.current = null;
      reportInline();
    };
    inlineFrameRef.current = window.requestAnimationFrame(() => {
      inlineFrameRef.current = window.requestAnimationFrame(() => {
        const layout = layoutReportRef.current;
        if (!layout) {
          report();
          return;
        }
        inlineFrameRef.current = -1;
        void layout.finally(() => {
          inlineFrameRef.current = window.requestAnimationFrame(report);
        });
      });
    });
  }, [reportInline]);

  const measure = useCallback(() => {
    const generation = generationRef.current;
    const strip = stripRef.current;
    if (!generation || !strip) return;
    // The candidate bar first: over the formula it moves the strip down.
    const candidates = suggestionsVisible() ? placeCandidateBar(strip) : null;
    if (!candidates) setStripTop(strip, PANEL_MARGIN);
    const rect = strip.getBoundingClientRect();
    // The layout is the page's answer to an activation: a strip not laid out
    // yet still answers (the native side drops a session that gets none).
    if (rect.width <= 0 && rect.height <= 0) void logAnywhere("measure: the strip has no size yet");
    const field = firstField();
    if (field && fieldIsEmpty(field) && stripFields().length === 1) {
      const mathfield = (field as unknown as { _mathfield?: CoreMathfield })._mathfield;
      const caret = mathfield ? visualTexEditingCore.caretRect(mathfield) : null;
      if (caret && Number.isFinite(caret.left)) formulaStartRef.current = caret.left;
    }
    const tag = candidates ? null : placeModeTag(rect);
    // In the field's line the formula's origin sits at the space the field
    // holds for it, on the line's baseline; on the caret line (pasting) the
    // strip's middle sits on the line.
    const box = inputMethodRef.current ? formulaBox() : null;
    const now = performance.now();
    const widths = coverWidthsRef.current.filter((width) => now - width.at < SHRINK_COVER_MS);
    const lastY = widths.length ? widths[widths.length - 1].y : null;
    const kept = widths.filter((width) => width.y === lastY);
    const keep = kept.length ? { y: lastY as number, right: Math.max(...kept.map((width) => width.right)) } : null;
    const covers = placeCovers(
      coverRef.current,
      inputMethodRef.current ? fieldBackgroundRef.current : null,
      fieldLineRef.current,
      piecePlacesRef.current,
      pieceStartRef.current,
      pieceEndRef.current,
      keep,
      pieceCaretRef.current,
    );
    if (covers?.last) {
      widths.push({ at: now, ...covers.last });
      // narrower than a width kept: drawn again once that is no longer kept
      if (keep && keep.y === covers.last.y && keep.right > covers.last.right + 0.5 && shrinkTimerRef.current === null) {
        shrinkTimerRef.current = window.setTimeout(() => {
          shrinkTimerRef.current = null;
          measureRef.current();
        }, SHRINK_COVER_MS + 20);
      }
    }
    coverWidthsRef.current = widths;
    const layout: AnywhereLayout = {
      width: Math.ceil(
        Math.max(
          rect.right + PANEL_MARGIN,
          candidates ? candidates.right + CANDIDATE_MARGIN : 0,
          tag ? tag.right + CANDIDATE_MARGIN : 0,
        ),
      ),
      height: Math.ceil(
        Math.max(
          rect.bottom + PANEL_MARGIN,
          candidates ? candidates.bottom + CANDIDATE_MARGIN : 0,
          tag ? tag.bottom + CANDIDATE_MARGIN : 0,
        ),
      ),
      anchorX: box ? box.left : formulaStartRef.current ?? rect.left + 1,
      anchorY: box ? box.baseline : rect.top + rect.height / 2,
    };
    if (inputMethodRef.current) scheduleInline();
    if (covers) {
      layout.width = Math.max(layout.width, Math.ceil(covers.right + 1));
      layout.height = Math.max(layout.height, Math.ceil(covers.bottom + 1));
    }
    if (inputMethodRef.current) {
      // the window a little larger than the formula, and made smaller only when much larger: the
      // formula and its cover grow into the spare room in the frame they grow
      const previous = lastLayoutRef.current;
      const slackY = Math.max(12, fieldLineRef.current?.height ?? 0);
      if (previous && layout.width <= previous.width && previous.width - layout.width < PANEL_SLACK * 2) layout.width = previous.width;
      else layout.width += PANEL_SLACK;
      if (previous && layout.height <= previous.height && previous.height - layout.height < slackY * 2) layout.height = previous.height;
      else layout.height += slackY;
    }
    if (sameLayout(lastLayoutRef.current, layout)) return;
    lastLayoutRef.current = layout;
    const report = reportAnywhereLayout(generation, layout).catch((error) =>
      console.error("Unable to place VisualTeX Anywhere", error),
    );
    layoutReportRef.current = report;
    void report.finally(() => {
      if (layoutReportRef.current === report) layoutReportRef.current = null;
    });
  }, [scheduleInline]);

  const measureRef = useRef(measure);
  measureRef.current = measure;

  // In the field's line (input method) the formula is set on lines as the
  // field wraps its text: each piece where the field put the space it holds
  // for it (a piece the field has not placed yet continues the line). A
  // formula of several rows stays as it is.
  useEffect(() => {
    setFormulaLineLayout(() => {
      if (!inputMethodRef.current || stripFields().length > 1) return null;
      return (widths) => {
        const box = formulaBox();
        const fontPx = box?.fontPx || EDITOR_BASE_FONT_PX * zoomRef.current;
        // the field decides each piece's line and where each line starts; on
        // a line, the pieces are spaced as in the formula (the field's space
        // for them is the same to half a filler)
        const out: ({ x: number; y: number } | null)[] = [];
        widths.forEach((_, index) => {
          const place = piecePlacesRef.current[index];
          const before = piecePlacesRef.current[index - 1];
          const drawn = out[index - 1];
          if (!place) out.push(null);
          else if (index > 0 && before && drawn && Math.abs(place[1] - before[1]) < 0.5) {
            out.push({ x: drawn.x + widths[index - 1], y: place[1] / fontPx });
          } else out.push({ x: place[0] / fontPx, y: place[1] / fontPx });
        });
        return out;
      };
    });
    return () => setFormulaLineLayout(null);
  }, []);

  const activate = useCallback(
    async (activation: AnywhereActivation) => {
      generationRef.current = activation.generation;
      committingRef.current = false;
      lastLayoutRef.current = null;
      lastMarkedRef.current = null;
      lastModeRef.current = null;
      composingRef.current = false;
      piecePlacesRef.current = [];
      pieceStartRef.current = null;
      pieceEndRef.current = null;
      pieceCaretRef.current = null;
      coverWidthsRef.current = [];
      inputMethodRef.current = Boolean(activation.inputMethod);
      setInputMethod(Boolean(activation.inputMethod));
      // set on several lines only in the field's line (input method)
      redrawOpenFormulas();
      applySystemAppearance();
      const kept = activation.resume && hasFormula() ? currentDocument() : null;
      // The size first: a size the field reports while the settings are read
      // (the anchor's) comes after this one and must stay.
      setZoom(formulaZoom(
        activation.caretHeight,
        activation.inputMethod ? activation.fontSize : null,
        activation.xHeight,
      ));
      setFieldLine(fieldLineOf(activation.lineHeight ?? activation.caretHeight, activation.lineDescent));
      // The main window's settings (fonts, symbols, copy format) as of now;
      // this window keeps its own formula (editorPersistenceStorage).
      const started = performance.now();
      void logAnywhere(`activate gen=${activation.generation}`);
      try {
        await Promise.all([
          useEditorStore.persist.rehydrate(),
          useAnywhereHotkeyStore.persist.rehydrate(),
          useFormulaRowHotkeyStore.persist.rehydrate(),
        ]);
        void logAnywhere(`activate gen=${activation.generation} settings read ms=${Math.round(performance.now() - started)}`);
      } catch (error) {
        console.error("Unable to read VisualTeX settings", error);
      }
      if (kept) {
        // The display shortcut makes the kept formula's row a display row.
        useEditorStore.getState().replaceDocumentState(
          activation.display
            ? {
                ...kept,
                lines: kept.lines.map((line) =>
                  line.id === kept.activeLineId || kept.lines.length === 1 ? { ...line, mode: "display" as const } : line,
                ),
              }
            : kept,
        );
      } else startNewFormula(activation.display ? "display" : "inline");
      fieldInkRef.current = activation.inputMethod ? opaqueHex(activation.textColor) : null;
      fieldBackgroundRef.current = activation.inputMethod ? opaqueHex(activation.backgroundColor) : null;
      setTextColor(activation.inputMethod ? cssColor(activation.textColor) ?? systemTextColor() : null);
      setNotice(
        !activation.accessibilityTrusted
          ? NOTICE.accessibility
          : activation.reopenApp
            ? NOTICE.reopen(activation.reopenApp)
            : activation.inputMethodMissing
              ? NOTICE.inputMethod
              : null,
      );
      appNameRef.current = activation.appName;
      const mode = rememberedOutputMode(activation.appName);
      outputModeRef.current = mode;
      setOutputMode(mode);
      warmOutputMode(mode);
      showModeTagRef.current(mode);
      setSession(activation.generation);
    },
    [startNewFormula],
  );

  const cancel = useCallback(() => {
    const generation = generationRef.current;
    if (!generation) return;
    startNewFormula();
    void dismissAnywhere(generation, false).catch((error) =>
      console.error("Unable to close VisualTeX Anywhere", error),
    );
  }, [startNewFormula]);

  const commit = useCallback(async () => {
    const generation = generationRef.current;
    if (!generation || committingRef.current) return;
    editorRef.current?.commitPendingTransaction();
    const { latexFormatProfile } = useEditorStore.getState();
    const formula = composedFormulaLines();
    if (!formula.length) {
      cancel();
      return;
    }
    const mode = outputModeRef.current;
    // each row as it is: inline ($…$) or display ($$…$$ or the user's form)
    const source = formatFormulaLinesUniversal(formula, latexFormatProfile);
    committingRef.current = true;
    try {
      let outcome;
      if (mode === "image") {
        const ink = fieldInkRef.current;
        const background = fieldBackgroundRef.current;
        const png = await formulaPng(
          formula.map((line) => line.latex),
          ink && background ? { ink, background } : null,
        );
        outcome = await commitAnywhereImage(generation, png);
      } else {
        let text = mode === "markdown" ? formulaMarkdown(formula) : source;
        if (mode === "unicode") {
          try {
            const converted = await formulaUnicode(formula.map((line) => line.latex));
            if (converted.text) text = converted.text;
            if (converted.notes.length) void logAnywhere(`unicode notes: ${converted.notes.join(" | ")}`);
          } catch (error) {
            // A formula the export path cannot read still goes in, as LaTeX.
            void logAnywhere(`unicode failed, LaTeX instead: ${error instanceof Error ? error.message : String(error)}`);
          }
        }
        outcome = await commitAnywhere(generation, text);
      }
      void logAnywhere(`commit gen=${generation} mode=${mode} outcome=${outcome}`);
      if (outcome === "inserted") {
        startNewFormula();
      } else if (outcome === "copied") {
        setNotice(NOTICE.copied);
        window.setTimeout(() => {
          if (generationRef.current === generation) cancel();
        }, COPIED_NOTICE_MS);
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      void logAnywhere(`commit gen=${generation} failed: ${message}`);
      setNotice(message);
    } finally {
      committingRef.current = false;
    }
  }, [cancel, startNewFormula]);

  // Activations and suspensions from the native panel.
  useEffect(() => {
    const disposers = [
      listenAnywhere(ANYWHERE_ACTIVATE_EVENT, isAnywhereActivation, (activation) => {
        void activate(activation);
      }),
      listenAnywhere(ANYWHERE_SUSPEND_EVENT, isAnywhereSuspension, (suspension) => {
        if (suspension.generation === generationRef.current && !suspension.keepDraft) {
          startNewFormula();
        }
      }),
      listenAnywhere(ANYWHERE_KEY_EVENT, isAnywhereKey, (key) => {
        if (key.generation !== generationRef.current || !inputMethodRef.current) return;
        if (key.meta && key.code === "KeyV" && !key.ctrl && !key.alt) {
          void pasteIntoFormula();
          return;
        }
        const shortcut = formulaShortcut(key);
        if (shortcut) runFormulaShortcut(shortcut);
        else typeKey(key);
      }),
      listenAnywhere(ANYWHERE_COMPOSE_EVENT, isAnywhereCompose, (message) => {
        if (message.generation === generationRef.current && inputMethodRef.current) {
          compose(message.phase, message.text, composingRef);
        }
      }),
      listenAnywhere(ANYWHERE_PIECES_EVENT, isAnywherePieces, (pieces) => {
        if (pieces.generation !== generationRef.current || !inputMethodRef.current) return;
        const same =
          pieces.places.length === piecePlacesRef.current.length &&
          pieces.places.every((place, index) => {
            const known = piecePlacesRef.current[index];
            return Math.abs(place[0] - known[0]) < 0.25 && Math.abs(place[1] - known[1]) < 0.25;
          });
        pieceStartRef.current = pieces.start ?? null;
        // an estimated end is no place to cover to (it may lie past the held space, over the field's text)
        pieceEndRef.current = pieces.endEstimated ? null : pieces.end ?? null;
        pieceCaretRef.current = pieces.caretLine ?? null;
        if (same) {
          measureRef.current();
          return;
        }
        piecePlacesRef.current = pieces.places;
        redrawOpenFormulas();
        measureRef.current();
      }),
      listenAnywhere(ANYWHERE_STYLE_EVENT, isAnywhereStyle, (style) => {
        if (style.generation !== generationRef.current || !inputMethodRef.current) return;
        if (style.fontSize) setZoom(formulaZoom(0, style.fontSize, style.xHeight));
        if (style.lineHeight) setFieldLine(fieldLineOf(style.lineHeight, style.lineDescent));
        if (style.textColor) {
          setTextColor(cssColor(style.textColor));
          fieldInkRef.current = opaqueHex(style.textColor) ?? fieldInkRef.current;
        }
        if (style.backgroundColor) {
          fieldBackgroundRef.current = opaqueHex(style.backgroundColor);
          measureRef.current();
        }
      }),
    ];
    return () => {
      for (const disposer of disposers) void disposer.then((unlisten) => unlisten());
    };
  }, [activate, startNewFormula, runFormulaShortcut, pasteIntoFormula]);

  // Ready once the editor has mounted its field: the native side then sends
  // activations here instead of holding them.
  useEffect(() => {
    let disposed = false;
    let timer = 0;
    const probe = () => {
      if (disposed) return;
      if (firstField()) {
        void reportAnywhereReady().catch((error) =>
          console.error("Unable to report VisualTeX Anywhere ready", error),
        );
        return;
      }
      timer = window.setTimeout(probe, 30);
    };
    startNewFormula();
    // The field mounted with the store's first row; remount it on the
    // strip's own row (FORMULA_LINE_ID) once React has rendered that row.
    timer = window.setTimeout(() => {
      if (disposed) return;
      editorRef.current?.refreshLayout();
      probe();
    }, 0);
    return () => {
      disposed = true;
      window.clearTimeout(timer);
    };
  }, [startNewFormula]);

  // The formula's caret as tall as the host field's line, set before the
  // panel measures the strip (a taller caret makes the strip taller).
  useLayoutEffect(() => {
    const formulaPx = EDITOR_BASE_FONT_PX * zoom;
    const band = caretBandFor(fieldLine, formulaPx);
    setFormulaCaretBand(band);
    if (!session) return;
    const line =
      `size gen=${generationRef.current} formula=${formulaPx.toFixed(1)}px line=${fieldLine ? fieldLine.height.toFixed(1) : "-"}px ` +
      `caret=${band ? `${((band.up + band.down) * formulaPx).toFixed(1)}px` : "editor"}`;
    if (line === lastSizeLogRef.current) return;
    lastSizeLogRef.current = line;
    void logAnywhere(line);
  }, [session, zoom, fieldLine]);

  // A new session: put the caret in the formula and place the panel. The
  // layout report is the page's answer to the activation (the native side
  // shows the panel on it), so it goes out with this commit: the parked,
  // hidden page gets no animation frames until it is shown.
  useLayoutEffect(() => {
    if (!session) return;
    // The candidate table is drawn at the formula's size.
    document.documentElement.style.setProperty(
      "--anywhere-candidate-size",
      `${(EDITOR_BASE_FONT_PX * zoom).toFixed(2)}px`,
    );
    editorRef.current?.focus({ target: "first", moveToEnd: true });
    measure();
    const frame = window.requestAnimationFrame(() => {
      editorRef.current?.focus({ target: "first", moveToEnd: true });
      measure();
    });
    return () => window.cancelAnimationFrame(frame);
  }, [session, zoom, measure]);

  // The strip follows the formula's size; the panel grows while candidates
  // are open so the bar has room below the formula, and follows the bar as
  // its candidates change.
  useEffect(() => {
    const strip = stripRef.current;
    if (!strip) return;
    const resize = new ResizeObserver(() => measure());
    resize.observe(strip);
    let candidates = suggestionsVisible();
    const popover = new MutationObserver(() => {
      const visible = suggestionsVisible();
      if (visible === candidates && !visible) return;
      candidates = visible;
      measure();
    });
    popover.observe(document.body, {
      subtree: true,
      childList: true,
      attributes: true,
      attributeFilter: ["class"],
    });
    return () => {
      resize.disconnect();
      popover.disconnect();
    };
  }, [measure]);

  // The field holds the formula's width as it changes (the formula is
  // drawn over that space); the input method follows the caret (its
  // candidates, pinyin in text).
  useEffect(() => {
    if (!session || !inputMethod) return;
    reportInline();
    const follow = () => scheduleInline();
    document.addEventListener("selection-change", follow, true);
    document.addEventListener("mode-change", follow, true);
    return () => {
      document.removeEventListener("selection-change", follow, true);
      document.removeEventListener("mode-change", follow, true);
    };
  }, [session, inputMethod, lines, zoom, reportInline, scheduleInline]);

  // The panel became key: the keyboard goes to the formula.
  useEffect(() => {
    const focusFormula = () => {
      if (generationRef.current) editorRef.current?.focus({ target: "active" });
    };
    window.addEventListener("focus", focusFormula);
    return () => window.removeEventListener("focus", focusFormula);
  }, []);

  // Input method composition first, then the candidate list, then the
  // editor's own Enter and Escape; what remains commits or cancels.
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (!generationRef.current || event.isComposing || event.keyCode === 229) return;
      const shortcut = inputMethodRef.current
        ? null
        : formulaShortcut({ code: event.code, ctrl: event.ctrlKey, alt: event.altKey, shift: event.shiftKey, meta: event.metaKey });
      if (shortcut) {
        event.preventDefault();
        event.stopImmediatePropagation();
        runFormulaShortcut(shortcut);
        return;
      }
      if (event.key !== "Enter" && event.key !== "Escape") return;
      if (event.shiftKey || event.altKey || event.metaKey || event.ctrlKey) return;
      const candidates = suggestionsVisible();
      const field = stripField();
      const model = field ? editingModelOf(field) : null;
      if (event.key === "Escape") {
        // An input method's Escape drops the composition, a half-typed
        // command and its candidates included; only a selection collapses
        // first (the editor's own Escape).
        if (!candidates && model &&
            visualTexEditingCore.escapeAction(model, event, false) === "collapse") return;
        event.preventDefault();
        event.stopImmediatePropagation();
        cancel();
        return;
      }
      // Enter takes the highlighted candidate.
      if (candidates) return;
      // Enter completes a command being typed, and adds a row in cases.
      if (field && hasRawLatexInput(field)) return;
      const action = model ? visualTexEditingCore.enterAction(model, event, false) : "split-row";
      if (action !== "split-row") {
        void logAnywhere(`enter left to the editor: ${action}`);
        return;
      }
      event.preventDefault();
      event.stopImmediatePropagation();
      void commit();
    };
    window.addEventListener("keydown", handleKeyDown, true);
    return () => window.removeEventListener("keydown", handleKeyDown, true);
  }, [cancel, commit, runFormulaShortcut]);

  // A new row (Ctrl/Alt+Enter) or a row switched between inline and display
  // (the row hotkeys): the tag says what the row now is.
  const activeRowMode = useEditorStore(
    (state) => (state.lines.find((line) => line.id === state.activeLineId) ?? state.lines[0])?.mode ?? "display",
  );
  const rowShapeRef = useRef<string | null>(null);
  useEffect(() => {
    const shape = `${session}:${lines.length}:${activeLineId}:${activeRowMode}`;
    const previous = rowShapeRef.current;
    rowShapeRef.current = shape;
    if (!session || !previous || !previous.startsWith(`${session}:`) || previous === shape) return;
    if (previous.split(":")[1] === String(lines.length) && previous.split(":")[3] === activeRowMode) return;
    showModeTag(outputModeRef.current);
  }, [session, lines.length, activeLineId, activeRowMode, showModeTag]);

  // The panel makes room for the mode's tag while it shows.
  useLayoutEffect(() => {
    if (session) measure();
  }, [modeTag, session, measure]);

  return (
    <>
    <div
      ref={stripRef}
      className={inputMethod ? "anywhere-strip is-input-method" : "anywhere-strip"}
      data-session={session || undefined}
      data-output-mode={outputMode}
      style={textColor ? ({ "--anywhere-text-color": textColor } as CSSProperties) : undefined}
    >
      <div ref={coverRef} className="anywhere-covers" aria-hidden="true" />
      <div className="anywhere-formula">
        <MathEditor
          ref={editorRef}
          lines={lines}
          activeLineId={activeLineId}
          formulaAlignment="left"
          latexCodeFormat={latexCodeFormat}
          zoom={zoom}
          reuseLineSlots
          showLineModeControls={false}
          rowModeHotkeys
          onPasteImage={pastedImage}
        />
      </div>
      {notice ? (
        <div className="anywhere-notice" role="status">
          {notice}
        </div>
      ) : null}
    </div>
    {modeTag && session ? (
      <div className="anywhere-mode-tag" role="status" data-mode={modeTag.mode}>
        {modeTag.status ? (
          <span className="anywhere-mode-tag__status">{modeTag.status}</span>
        ) : (
          <>
            <span className="anywhere-mode-tag__name">{OUTPUT_MODE_LABELS[modeTag.mode]}</span>
            <span className="anywhere-mode-tag__row">{modeTag.row}</span>
          </>
        )}
        {modeTag.preview && !modeTag.status ? <span className="anywhere-mode-tag__preview">{modeTag.preview}</span> : null}
        {modeTag.status ? null : (
          <span className="anywhere-mode-tag__key">{anywhereHotkeyLabel(useAnywhereHotkeyStore.getState().bindings["anywhere-next-output"])}</span>
        )}
      </div>
    ) : null}
    </>
  );
}
