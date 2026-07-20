export const POST_COMPOSITION_ENTER_WINDOW_MS = 80;

export interface ImeKeyboardEventLike {
  key: string;
  isComposing: boolean;
  keyCode: number;
  metaKey?: boolean;
  ctrlKey?: boolean;
  altKey?: boolean;
  shiftKey?: boolean;
}

export type ImeKeyDownDecision =
  | "composition"
  | "post-composition-enter"
  | null;

function normalizedTimestamp(value: number) {
  return Number.isFinite(value) && value >= 0 ? value : 0;
}

function isModifierOnlyKey(key: string) {
  return (
    key === "Shift" ||
    key === "Control" ||
    key === "Alt" ||
    key === "Meta" ||
    key === "CapsLock"
  );
}

/**
 * Tracks the browser IME transaction independently from MathLive.
 *
 * WebKit can dispatch compositionend and then replay the physical Enter key as
 * a normal keydown. That replay must be swallowed once, while a later Enter
 * remains available for VisualTeX's "new formula line" command.
 */
export class ImeCompositionGuard {
  private composing = false;
  private postCompositionEnterDeadline = -1;

  compositionStart() {
    this.composing = true;
    this.postCompositionEnterDeadline = -1;
  }

  compositionEnd(timestamp: number) {
    this.composing = false;
    this.postCompositionEnterDeadline =
      normalizedTimestamp(timestamp) + POST_COMPOSITION_ENTER_WINDOW_MS;
  }

  isComposing() {
    return this.composing;
  }

  keyDown(
    event: ImeKeyboardEventLike,
    timestamp: number,
  ): ImeKeyDownDecision {
    const browserStillComposing = event.isComposing || event.keyCode === 229;
    if (browserStillComposing) return "composition";

    // MathLive/WebKit can occasionally omit compositionend after committing a
    // command. A genuinely non-composing key releases that stale local state.
    if (this.composing) {
      this.composing = false;
      this.postCompositionEnterDeadline = -1;
      return null;
    }

    const now = normalizedTimestamp(timestamp);
    if (
      event.key === "Enter" &&
      now <= this.postCompositionEnterDeadline
    ) {
      this.postCompositionEnterDeadline = -1;
      return "post-composition-enter";
    }

    // Keep the one-shot guard across modifier-only events, but do not let an
    // unrelated real key leave it armed for a later Enter.
    if (!isModifierOnlyKey(event.key)) {
      this.postCompositionEnterDeadline = -1;
    }
    return null;
  }
}
