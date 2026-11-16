import { invoke } from "@tauri-apps/api/core";
import { listen, type UnlistenFn } from "@tauri-apps/api/event";

/** src-tauri/src/anywhere/mod.rs */
export const ANYWHERE_ACTIVATE_EVENT = "visualtex-anywhere-activate";
export const ANYWHERE_SUSPEND_EVENT = "visualtex-anywhere-suspend";
export const ANYWHERE_KEY_EVENT = "visualtex-anywhere-key";
export const ANYWHERE_STYLE_EVENT = "visualtex-anywhere-style";
export const ANYWHERE_SHORTCUT_EVENT = "visualtex-anywhere-shortcut";
export const ANYWHERE_PIECES_EVENT = "visualtex-anywhere-pieces";
/** macOS: where the panel stands on the screen (the native side's report each time it sets the panel's frame). */
export const ANYWHERE_FRAME_EVENT = "visualtex-anywhere-frame";

export type AnywhereAnchorSource = "caret" | "field" | "pointer";

export interface AnywhereActivation {
  generation: number;
  /** The host's line height in CSS pixels (points on macOS); the formula is sized from it. */
  caretHeight: number;
  anchorSource: AnywhereAnchorSource;
  appName: string;
  accessibilityTrusted: boolean;
  /** macOS: the field's application has been running since before VisualTeX's input method was an
   *  input source (macOS gives an application its input methods when it starts): to be reopened */
  reopenApp?: string | null;
  /** The previous session hid without a commit or cancel: keep its formula. */
  resume: boolean;
  /**
   * VisualTeX's input method types the formula: the host field keeps the
   * keyboard and holds the formula's width in its line, the input method
   * sends its keys here (ANYWHERE_KEY_EVENT) and the formula is drawn over
   * that space, in the field's text size and colour.
   */
  inputMethod?: boolean;
  /** The display shortcut opened the formula: its row is a display row
   * (else a new formula starts inline). */
  display?: boolean;
  /** The field's font size in CSS pixels (points), when the field tells it. */
  fontSize?: number | null;
  /** The field's text colour (#rrggbbaa), when the field tells it (or as
   * read off the screen). */
  textColor?: string | null;
  /** The field font's x-height over its size. */
  xHeight?: number | null;
  /** The field's line (its caret) in CSS pixels (points), and how far it reaches below
   * the baseline. */
  lineHeight?: number | null;
  lineDescent?: number | null;
  /** The field's background (#rrggbbaa), as read off the screen. */
  backgroundColor?: string | null;
}

/** The field's text style, learnt after the formula opened. */
export interface AnywhereStyle {
  generation: number;
  fontSize?: number | null;
  textColor?: string | null;
  xHeight?: number | null;
  lineHeight?: number | null;
  lineDescent?: number | null;
  backgroundColor?: string | null;
}

/**
 * Where the host field put the formula's pieces (it holds each piece's width
 * in its text and wraps them as its own text): each piece's start from the
 * first piece's, in CSS pixels (points), y down (one place per piece the field has laid
 * out, in order).
 */
export interface AnywherePieces {
  generation: number;
  places: [number, number][];
  /** Where the held space starts (its anchor, before the first piece) and ends (after the last
   *  piece), when the field tells. */
  start?: [number, number] | null;
  end?: [number, number] | null;
  /** The end is the input method's estimate (the field did not tell it). */
  endEstimated?: boolean;
  /** macOS: the field's own caret after the held space: how far its line's bottom is below the last
   *  line's (points), and its height. */
  caretLine?: [number, number] | null;
}

export function isAnywherePieces(value: unknown): value is AnywherePieces {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<AnywherePieces>;
  return (
    typeof candidate.generation === "number" &&
    Array.isArray(candidate.places) &&
    candidate.places.every(
      (place) => Array.isArray(place) && place.length === 2 && place.every((n) => typeof n === "number" && Number.isFinite(n)),
    )
  );
}

/**
 * macOS: where the panel stands, and the visible area of its display (menu bar and Dock left out), in CSS pixels (points) from
 * the main display's top left, as a window's screenX/screenY would be. The page cannot read this from its window: WebKit gives
 * a panel's window.screenX/Y as 0 and the screen's height, whatever the panel's place.
 */
export interface AnywhereFrame {
  x: number;
  y: number;
  width: number;
  height: number;
  visibleLeft: number;
  visibleTop: number;
  visibleRight: number;
  visibleBottom: number;
}

export function isAnywhereFrame(value: unknown): value is AnywhereFrame {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Record<string, unknown>;
  return ["x", "y", "width", "height", "visibleLeft", "visibleTop", "visibleRight", "visibleBottom"].every(
    (key) => typeof candidate[key] === "number" && Number.isFinite(candidate[key]),
  );
}

export function isAnywhereStyle(value: unknown): value is AnywhereStyle {
  return Boolean(value && typeof value === "object" && typeof (value as AnywhereStyle).generation === "number");
}

/** A key the input method took from the host field: DOM KeyboardEvent fields
 * and, for a character key, the text it types. */
export interface AnywhereKey {
  generation: number;
  key: string;
  code: string;
  text?: string;
  shift: boolean;
  alt: boolean;
  ctrl: boolean;
  meta: boolean;
  repeat: boolean;
}

export interface AnywhereSuspension {
  generation: number;
  keepDraft: boolean;
}

/** Windows: a shortcut of the open formula pressed in the host application (registered
 * with Windows while a formula is open: some applications keep modifier
 * chords from the input method). */
export interface AnywhereShortcut {
  generation: number;
  action: string;
}

export function isAnywhereShortcut(value: unknown): value is AnywhereShortcut {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<AnywhereShortcut>;
  return typeof candidate.generation === "number" && typeof candidate.action === "string";
}

/**
 * The panel's size and the point of the strip that sits on the host's caret
 * line (where the formula starts, at its vertical centre), in CSS pixels
 * (points) from the window's top left.
 */
export interface AnywhereLayout {
  width: number;
  height: number;
  anchorX: number;
  anchorY: number;
}

export type AnywhereCommitOutcome = "inserted" | "copied" | "stale";

export function isAnywhereActivation(value: unknown): value is AnywhereActivation {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<AnywhereActivation>;
  return (
    typeof candidate.generation === "number" &&
    typeof candidate.caretHeight === "number" &&
    typeof candidate.appName === "string" &&
    typeof candidate.accessibilityTrusted === "boolean" &&
    typeof candidate.resume === "boolean"
  );
}

export function isAnywhereKey(value: unknown): value is AnywhereKey {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<AnywhereKey>;
  return (
    typeof candidate.generation === "number" &&
    typeof candidate.key === "string" &&
    typeof candidate.code === "string" &&
    (candidate.text === undefined || typeof candidate.text === "string")
  );
}

export function isAnywhereSuspension(value: unknown): value is AnywhereSuspension {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<AnywhereSuspension>;
  return typeof candidate.generation === "number" && typeof candidate.keepDraft === "boolean";
}

export function listenAnywhere<T>(
  event: string,
  accept: (value: unknown) => value is T,
  handler: (payload: T) => void,
): Promise<UnlistenFn> {
  return listen<unknown>(event, (message) => {
    if (accept(message.payload)) handler(message.payload);
  });
}

export const reportAnywhereReady = () => invoke<void>("anywhere_ready");

export const reportAnywhereLayout = (generation: number, layout: AnywhereLayout) =>
  invoke<void>("anywhere_layout", { generation, layout });

export const commitAnywhere = (generation: number, text: string) =>
  invoke<AnywhereCommitOutcome>("anywhere_commit", { generation, text });

/** Screenshot OCR in the formula: the area the user selects, as a PNG
 *  (base64); null when cancelled. */
export const captureAnywhereScreenshot = () => invoke<string | null>("anywhere_capture_screenshot");

/** Ctrl+V (⌘V) in the formula: the clipboard's image (for OCR) or text. */
export type AnywhereClipboard =
  | { kind: "image"; dataBase64: string; extension: string }
  | { kind: "text"; text: string }
  | { kind: "nothing" };

export const readAnywhereClipboard = () => invoke<AnywhereClipboard>("anywhere_read_clipboard");

/** The Image output mode: a PNG (base64) pasted into the field. */
export const commitAnywhereImage = (generation: number, dataBase64: string) =>
  invoke<AnywhereCommitOutcome>("anywhere_commit_image", { generation, dataBase64 });

export const dismissAnywhere = (generation: number, keepDraft: boolean) =>
  invoke<void>("anywhere_dismiss", { generation, keepDraft });

/** The formula's text (what the field gets if it ends the composition), its
 * caret from the formula's start (x, and y down to the caret's line), and
 * the widths of the pieces the field holds for it (TeX's break places). */
export const reportAnywhereMarked = (
  generation: number,
  text: string,
  width: number,
  caretX: number,
  caretY: number,
  pieces: number[],
) => invoke<void>("anywhere_marked", { generation, text, width, caretX, caretY, pieces });

/** Windows: the shortcuts that act while a formula is open, for Windows to give them
 * to this page whatever the host application does with modifier chords. */
export const reportAnywhereFormulaShortcuts = (
  shortcuts: { action: string; chord: { code: string; ctrlKey: boolean; altKey: boolean; shiftKey: boolean; metaKey: boolean } }[],
) => invoke<void>("anywhere_formula_shortcuts", { shortcuts });

/** A line in the native anywhere.log (no formula text). */
export const logAnywhere = (line: string) =>
  invoke<void>("anywhere_log", { line }).catch(() => undefined);
