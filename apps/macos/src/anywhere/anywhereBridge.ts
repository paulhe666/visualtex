import { invoke } from "@tauri-apps/api/core";
import { listen, type UnlistenFn } from "@tauri-apps/api/event";

/** src-tauri/src/anywhere/mod.rs */
export const ANYWHERE_ACTIVATE_EVENT = "visualtex-anywhere-activate";
export const ANYWHERE_SUSPEND_EVENT = "visualtex-anywhere-suspend";
export const ANYWHERE_KEY_EVENT = "visualtex-anywhere-key";
export const ANYWHERE_STYLE_EVENT = "visualtex-anywhere-style";
export const ANYWHERE_COMPOSE_EVENT = "visualtex-anywhere-compose";
export const ANYWHERE_PIECES_EVENT = "visualtex-anywhere-pieces";

export type AnywhereAnchorSource = "caret" | "field" | "pointer";

export interface AnywhereActivation {
  generation: number;
  /** The host's line height in points; the formula is sized from it. */
  caretHeight: number;
  anchorSource: AnywhereAnchorSource;
  appName: string;
  accessibilityTrusted: boolean;
  /** VisualTeX's input method is installed but not among the input sources */
  inputMethodMissing?: boolean;
  /** The field's application selected the input method without starting it: to be reopened */
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
  /** The field's font size in points, when the field tells it. */
  fontSize?: number | null;
  /** The field's text colour (#rrggbbaa), when the field tells it (or as
   * read off the screen). */
  textColor?: string | null;
  /** The field font's x-height over its size. */
  xHeight?: number | null;
  /** The field's line (its caret) in points, and how far it reaches below
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

/** Pinyin spelled in the formula's text by the input method: the editor's
 * composition ("update" with the spelling, "end" with the chosen text or
 * nothing). */
export interface AnywhereCompose {
  generation: number;
  phase: "update" | "end";
  text: string;
}

export function isAnywhereCompose(value: unknown): value is AnywhereCompose {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<AnywhereCompose>;
  return (
    typeof candidate.generation === "number" &&
    (candidate.phase === "update" || candidate.phase === "end") &&
    typeof candidate.text === "string"
  );
}

/**
 * Where the host field put the formula's pieces (it holds each piece's width
 * in its text and wraps them as its own text): each piece's start from the
 * first piece's, in points, y down (one place per piece the field has laid
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

/** ⌘V in the formula: the clipboard's image (for OCR) or text. */
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

/** The formula's caret is in text (\text{}): the input method spells pinyin there. */
export const reportAnywhereMode = (generation: number, text: boolean) =>
  invoke<void>("anywhere_mode", { generation, text });

/** A line in the native anywhere.log (no formula text). */
export const logAnywhere = (line: string) =>
  invoke<void>("anywhere_log", { line }).catch(() => undefined);
