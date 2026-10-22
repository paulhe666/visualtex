/**
 * VisualTeX: macOS text-input guards of the keyboard sink.
 *
 * - A physical Backslash key is not a character until the text-input system
 *   commits it. In macOS Pinyin, WKWebView can report key="\\" on keydown for
 *   the Backslash-position key and then commit the ideographic comma "、"
 *   through beforeinput/input. Running the keybinding on keydown would enter
 *   command mode before the input method owns the character, so the key waits
 *   until the sink shows the committed text: a literal "\\" still starts a
 *   command, remapped punctuation stays normal input.
 * - macOS Pinyin can send the Space that confirmed a candidate twice: in the
 *   IME transaction and once more as plain input after compositionend.
 * - A finished or cancelled composition may be followed by a stray literal
 *   backslash input without a new physical keydown.
 *
 * Windows keeps binding the physical key code instead (keybindings-definitions).
 */

export interface MacTextInputDelegate {
  onKeystroke(event: KeyboardEvent): boolean;
}

type PendingKey = {
  key: string;
  code: string;
  ctrlKey: boolean;
  metaKey: boolean;
  altKey: boolean;
  shiftKey: boolean;
  repeat: boolean;
  location: number;
};

const CHINESE_COMMIT = /^[㐀-鿿豈-﫿，。；：！？、（）【】《》“”‘’\s]+$/;

function isUnmodifiedPhysicalBackslash(event: KeyboardEvent): boolean {
  return (
    (event.code === 'Backslash' || event.code === 'IntlBackslash') &&
    !event.ctrlKey &&
    !event.metaKey &&
    !event.altKey &&
    !event.shiftKey
  );
}

/** Unidentified/code="" keydowns inside one printable-key transaction: input-method plumbing */
function isInputMethodNoiseKeydown(event: KeyboardEvent): boolean {
  return (
    event.key === 'Unidentified' &&
    !event.code &&
    !event.ctrlKey &&
    !event.metaKey &&
    !event.altKey
  );
}

export class MacTextInputGuard {
  private deferredBackslash: PendingKey | null = null;
  private backslashMayReplayInput = false;
  private backslashKeyDown = false;
  private suppressRepeatedBackslash = false;
  private suppressOrphanBackslash = false;
  private candidateMayReplaySpace = false;
  private candidateReplayTimer: ReturnType<typeof setTimeout> | 0 = 0;

  constructor(private readonly delegate: MacTextInputDelegate) {}

  private clearCandidateReplay(): void {
    this.candidateMayReplaySpace = false;
    if (this.candidateReplayTimer) clearTimeout(this.candidateReplayTimer);
    this.candidateReplayTimer = 0;
  }

  private commitDeferredBackslash(): boolean {
    const pending = this.deferredBackslash;
    if (!pending) return false;
    this.deferredBackslash = null;
    const expectsTextInput = this.delegate.onKeystroke(
      new KeyboardEvent('keypress', pending)
    );
    this.backslashMayReplayInput = !expectsTextInput;
    return expectsTextInput;
  }

  /** The Space a Chinese candidate replays after compositionend */
  isReplayedCandidateSpace(event: KeyboardEvent, composing: boolean): boolean {
    return (
      !composing &&
      this.candidateMayReplaySpace &&
      (event.key === ' ' || event.code === 'Space')
    );
  }

  /**
   * Start of a sink keydown. Returns true when the event is used up; `reset`
   * clears the sink's pending keydown/keypress.
   */
  keydownStart(event: KeyboardEvent, composing: boolean, reset: () => void): boolean {
    if (this.isReplayedCandidateSpace(event, composing)) {
      // WebKit may replay the Space that confirmed the candidate: the marker
      // stays for its input event
      reset();
      event.preventDefault();
      event.stopPropagation();
      return true;
    }
    if (!composing && event.keyCode === 229 && isUnmodifiedPhysicalBackslash(event)) {
      // a delayed semantic keydown after the IME already committed its text:
      // not a new command-prefix transaction
      this.backslashKeyDown = false;
      this.suppressRepeatedBackslash = false;
      reset();
      return true;
    }
    const physicalBackslash = isUnmodifiedPhysicalBackslash(event);
    if (physicalBackslash && this.backslashKeyDown && !event.repeat) {
      this.suppressRepeatedBackslash = true;
      reset();
      event.preventDefault();
      event.stopPropagation();
      return true;
    }
    if (physicalBackslash) this.backslashKeyDown = true;
    else if (!isInputMethodNoiseKeydown(event)) {
      this.backslashKeyDown = false;
      this.suppressRepeatedBackslash = false;
    }
    if (!composing) {
      this.clearCandidateReplay();
      if (!isInputMethodNoiseKeydown(event)) {
        this.deferredBackslash = null;
        this.backslashMayReplayInput = false;
        this.suppressOrphanBackslash = false;
      }
    }
    return false;
  }

  /** After the sink's composition and modifier filters: defers a physical Backslash */
  keydownDefer(event: KeyboardEvent, composing: boolean, reset: () => void): boolean {
    if (isUnmodifiedPhysicalBackslash(event) && event.key === 'Unidentified') {
      this.suppressOrphanBackslash = true;
      reset();
      return true;
    }
    if (!composing && event.key === '\\' && isUnmodifiedPhysicalBackslash(event)) {
      // a fresh physical Backslash is a new transaction, not an orphan replay
      this.suppressOrphanBackslash = false;
      this.deferredBackslash = {
        key: event.key,
        code: event.code,
        ctrlKey: event.ctrlKey,
        metaKey: event.metaKey,
        altKey: event.altKey,
        shiftKey: event.shiftKey,
        repeat: event.repeat,
        location: event.location,
      };
      reset();
      return true;
    }
    return false;
  }

  keyup(event: KeyboardEvent): void {
    if (event.code !== 'Backslash' && event.code !== 'IntlBackslash') return;
    this.backslashKeyDown = false;
    this.suppressRepeatedBackslash = false;
  }

  /** A sink keypress (after the composition filter); true when used up */
  keypress(event: KeyboardEvent, clearKeypress: () => void): boolean {
    if (this.suppressRepeatedBackslash && event.key === '\\') {
      event.preventDefault();
      event.stopPropagation();
      return true;
    }
    if (this.deferredBackslash) {
      clearKeypress();
      if (event.key === '\\' && !this.commitDeferredBackslash()) {
        event.preventDefault();
        event.stopPropagation();
      }
      return true;
    }
    return false;
  }

  compositionStart(): void {
    this.deferredBackslash = null;
    this.backslashMayReplayInput = false;
    this.suppressOrphanBackslash = false;
    this.clearCandidateReplay();
  }

  compositionEnd(data: string): void {
    this.suppressOrphanBackslash = true;
    this.candidateMayReplaySpace = CHINESE_COMMIT.test(data) && /[^\s]/.test(data);
    if (this.candidateReplayTimer) clearTimeout(this.candidateReplayTimer);
    this.candidateReplayTimer = this.candidateMayReplaySpace
      ? setTimeout(() => {
          this.candidateMayReplaySpace = false;
          this.candidateReplayTimer = 0;
        }, 250)
      : 0;
  }

  beforeInput(ev: InputEvent): void {
    const data = ev.data ?? '';
    const backslash = ev.inputType === 'insertText' && data === '\\';
    if (this.suppressRepeatedBackslash && backslash) {
      this.suppressRepeatedBackslash = false;
      ev.preventDefault();
    } else if (
      !this.deferredBackslash &&
      !this.backslashMayReplayInput &&
      backslash &&
      this.suppressOrphanBackslash
    ) {
      this.suppressOrphanBackslash = false;
      ev.preventDefault();
    } else if (this.deferredBackslash && backslash) {
      if (!this.commitDeferredBackslash()) ev.preventDefault();
    } else if (this.backslashMayReplayInput && backslash) {
      // the same payload replayed after the deferred backslash was committed
      // (a genuine second key has its own keydown, which clears the marker)
      this.backslashMayReplayInput = false;
      ev.preventDefault();
    }
    if (data !== '\\') this.suppressOrphanBackslash = false;
  }

  /** A non-composition sink input; true when it is used up (not typed) */
  input(ev: InputEvent): boolean {
    const data = ev.data ?? '';
    const consume = (): boolean => {
      ev.preventDefault();
      ev.stopPropagation();
      return true;
    };
    if (this.suppressRepeatedBackslash && data === '\\') {
      this.suppressRepeatedBackslash = false;
      return consume();
    }
    if (
      !this.deferredBackslash &&
      !this.backslashMayReplayInput &&
      data === '\\' &&
      this.suppressOrphanBackslash
    ) {
      this.suppressOrphanBackslash = false;
      return consume();
    }
    if (this.deferredBackslash) {
      if (data === '\\' && !this.commitDeferredBackslash()) return consume();
      // the text-input system remapped the key (Pinyin Backslash -> "、"):
      // the committed text owns the transaction
      this.deferredBackslash = null;
    }
    if (this.backslashMayReplayInput && data === '\\') {
      this.backslashMayReplayInput = false;
      return consume();
    }
    this.backslashMayReplayInput = false;
    if (data !== '\\') this.suppressOrphanBackslash = false;
    if (this.candidateMayReplaySpace) {
      this.clearCandidateReplay();
      if (/^\s+$/.test(data)) return consume();
    }
    return false;
  }
}
