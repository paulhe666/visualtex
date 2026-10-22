/**
 * This modules handles low-level keyboard events and normalize them across
 * browsers.
 *
 * See https://dvcs.w3.org/hg/d4e/raw-file/tip/key-event-test.html
 * (also at https://w3c.github.io/uievents/tools/key-event-viewer.html)
 *
 *
 * - **KeyboardEvent.key** (the printable value associated with the key or a string
 * for special keys)
 *  https://developer.mozilla.org/en-US/docs/Web/API/KeyboardEvent/key/Key_Values
 *  https://www.w3.org/TR/uievents-key/
 *
 * - **KeyboardEvent.code** (the physical key being pressed. On an AZERTY keyboard
 * the key labelled "A" is KeyQ)
 *  https://developer.mozilla.org/en-US/docs/Web/API/KeyboardEvent/code/code_values
 *  https://www.w3.org/TR/uievents-code/
 *
 * Note:
 * - `charCode`, `keyCode` and `which` are deprecated
 *
 * For background, see this info regarding keybinding in VSCode:
 * - https://github.com/microsoft/vscode/tree/master/src/vs/workbench/services/keybinding
 * - https://github.com/microsoft/vscode/wiki/Keybinding-Issues
 */

import { normalizeKeyboardEvent } from './keyboard-layout';
import {
  deepActiveElement,
  mightProducePrintableCharacter,
} from '../ui/events/utils';
import { osPlatform } from '../ui/utils/capabilities';
import { MacTextInputGuard } from './macos-text-input';

export interface KeyboardDelegate {
  cancelComposition: () => void;
  blur: () => void;
  focus: () => void;
  hasFocus: () => boolean;
  setValue: (value: string) => void;
  setAriaLabel: (value: string) => void;
  /** VisualTeX: also the caret's height, so the IME window opens beside the caret */
  moveTo: (x: number, y: number, height?: number) => void;
  dispose: () => void;
  /**
   * VisualTeX: whether a keydown is the Enter WebKit replays just after an IME
   * candidate was confirmed (it is used up once). A keydown while a stale
   * composition is open, or any other real key, ends the wait.
   */
  takePostCompositionEnter: (event: KeyboardEvent) => boolean;
}

export interface KeyboardDelegateInterface {
  onKeystroke: (ev: KeyboardEvent) => boolean;
  onInput: (text: string) => void;
  onCut: (ev: ClipboardEvent) => void;
  onCopy: (ev: ClipboardEvent) => void;
  onPaste: (ev: ClipboardEvent) => boolean;
  onFocus: () => void;
  onBlur: () => void;
  onCompositionStart: (composition: string) => void;
  onCompositionUpdate: (composition: string) => void;
  onCompositionEnd: (composition: string) => void;
}

/** VisualTeX: how long after compositionend the Enter that WebKit replays is swallowed (ms) */
const POST_COMPOSITION_ENTER_MS = 80;

/**
 * Setup to capture the keyboard events from a `TextArea` and redispatch them to
 * handlers.
 *
 * In general, commands (arrows, delete, etc..) should be handled
 * in the `keystroke()` handler while text input should be handled in
 * `typedtext()`.
 *
 * @param {HTMLElement} keyboardSink The element that captures the keyboard
 * events.
 * @param {KeyboardDelegateInterface} delegate
 * @param {function} handlers.keystroke invoked on a key down event, including
 * for special keys such as ESC, arrow keys, tab, etc... and their variants
 * with modifiers.
 * @param {function} handlers.typedtext invoked on a keypress or other events
 * when a key corresponding to a character has been pressed. This include `a-z`,
 * `0-9`, `{}`, `^_()`, etc...
 * This does not include arrow keys, tab, etc... but does include 'space'
 * When a 'character' key is pressed, both `keystroke()` and `typedtext()` will
 * be invoked. When a control/function key is pressed, only `keystroke()` will
 * be invoked. In some cases, for example when using input methods or entering
 * emoji, only `onInput()` will be invoked.
 */
export function delegateKeyboardEvents(
  keyboardSink: HTMLElement,
  element: HTMLElement,
  delegate: KeyboardDelegateInterface
): KeyboardDelegate {
  let keydownEvent: KeyboardEvent | null = null;
  let keypressEvent: KeyboardEvent | null = null;
  let compositionInProgress = false;
  // VisualTeX: when the Enter WebKit replays after compositionend stops being
  // swallowed (POST_COMPOSITION_ENTER_MS after the compositionend)
  let postCompositionEnterDeadline = -1;
  let focusInProgress = false;
  let blurInProgress = false;

  const controller = new AbortController();
  const signal = controller.signal;

  // VisualTeX: macOS text-input guards (Pinyin Backslash, replayed candidate
  // Space, orphan backslash after a composition)
  const macTextInput = osPlatform() === 'macos' ? new MacTextInputGuard(delegate) : null;
  const resetPendingKeys = (): void => {
    keydownEvent = null;
    keypressEvent = null;
    keyboardSink.textContent = '';
  };

  // VisualTeX: WebView2 and some IMEs drop events of a composition (most
  // often compositionend while switching input methods). These keep one
  // composition transaction at a time and end a stale one, so the field is
  // never left locked in composition mode.
  const visualTexCompositionOwnsKeyboardEvent = (event?: KeyboardEvent | null): boolean => {
    if (!event) return false;
    if (event.isComposing || event.keyCode === 229) return true;
    if (event.key === 'Process' || event.key === 'Dead') return true;
    if (
      event.key === 'Shift' ||
      event.key === 'Control' ||
      event.key === 'Alt' ||
      event.key === 'Meta' ||
      event.key === 'CapsLock'
    )
      return true;
    return /(Control|Meta|Alt|Shift)(Left|Right)/.test(event.code || '');
  };
  const visualTexEndCompositionTransaction = (data = ''): boolean => {
    if (!compositionInProgress) return false;
    compositionInProgress = false;
    macTextInput?.compositionEnd(data);
    keydownEvent = null;
    keypressEvent = null;
    keyboardSink.textContent = '';
    delegate.onCompositionEnd(data);
    return true;
  };
  const visualTexNotifyRecoveredCompositionEnd = (data: string): void => {
    // When WebView2 replaces a missing compositionend with a final
    // non-composing input event, MathLive can repair its model internally, but
    // the host still needs the terminal composition signal to close its own
    // history/source transaction. Dispatch on the host, not the keyboard sink,
    // so this cannot re-enter MathLive's native composition listener.
    let event: Event;
    try {
      event = new CompositionEvent('compositionend', {
        bubbles: true,
        composed: true,
        data,
      });
    } catch (_error) {
      event = new Event('compositionend', {
        bubbles: true,
        composed: true,
      });
      Object.defineProperty(event, 'data', { value: data });
    }
    postCompositionEnterDeadline = event.timeStamp + POST_COMPOSITION_ENTER_MS;
    element.dispatchEvent(event);
  };

  if (macTextInput) {
    // the replayed Space is taken at the window and at the host (capture)
    // before application handlers on the math-field can queue it
    const consumeReplayedSpace = (event: KeyboardEvent): void => {
      if (!macTextInput.isReplayedCandidateSpace(event, compositionInProgress)) return;
      keydownEvent = null;
      event.preventDefault();
      event.stopImmediatePropagation();
    };
    window.addEventListener('keydown', consumeReplayedSpace, { capture: true, signal });
    element.addEventListener('keydown', consumeReplayedSpace, { capture: true, signal });
    keyboardSink.addEventListener('keyup', (event) => macTextInput.keyup(event), {
      capture: true,
      signal,
    });
  }

  keyboardSink.addEventListener(
    'keydown',
    (event: KeyboardEvent): void => {
      if (macTextInput?.keydownStart(event, compositionInProgress, resetPendingKeys)) return;
      if (compositionInProgress) {
        if (visualTexCompositionOwnsKeyboardEvent(event)) {
          keydownEvent = null;
          return;
        }
        // WebView2/IME can drop compositionend while switching input methods.
        // A definite non-composition key proves that transaction is stale.
        // Cancel the old transaction before handling this key so MathLive
        // cannot remain permanently locked in composition mode.
        visualTexEndCompositionTransaction('');
      }
      // "Process" key indicates commit of IME session (on Firefox)
      // It's handled with compositionEnd so it can be safely ignored
      if (
        event.key === 'Process' ||
        event.code === 'CapsLock' ||
        /(Control|Meta|Alt|Shift)(Left|Right)/.test(event.code)
      ) {
        keydownEvent = null;
        return;
      }
      if (macTextInput?.keydownDefer(event, compositionInProgress, resetPendingKeys)) return;

      keydownEvent = event;
      keypressEvent = null;
      if (!delegate.onKeystroke(event)) keydownEvent = null;
      else keyboardSink.textContent = '';
    },
    { capture: true, signal }
  );

  keyboardSink.addEventListener(
    'keypress',
    (event) => {
      if (compositionInProgress) {
        if (visualTexCompositionOwnsKeyboardEvent(event)) return;
        visualTexEndCompositionTransaction('');
      }
      if (macTextInput?.keypress(event, () => (keypressEvent = null))) return;
      // If this is not the first keypress after a keydown, that is,
      // if this is a repeated keystroke, call the keystroke handler.
      if (keydownEvent && keypressEvent) delegate.onKeystroke(keydownEvent);

      keypressEvent = event;
    },
    { capture: true, signal }
  );

  keyboardSink.addEventListener(
    'compositionstart',
    (event: CompositionEvent) => {
      keyboardSink.textContent = '';
      macTextInput?.compositionStart();
      // A second compositionstart without a matching compositionend means the
      // previous IME transaction was abandoned. Roll it back before starting
      // the new one instead of layering two transactions on one model.
      if (compositionInProgress)
        visualTexEndCompositionTransaction('');
      compositionInProgress = true;
      postCompositionEnterDeadline = -1;
      keydownEvent = null;
      keypressEvent = null;

      delegate.onCompositionStart(event.data);
    },
    { capture: true, signal }
  );

  keyboardSink.addEventListener(
    'compositionupdate',
    (ev: CompositionEvent) => {
      if (!compositionInProgress) return;

      delegate.onCompositionUpdate(ev.data);
    },
    { capture: true, signal }
  );

  keyboardSink.addEventListener(
    'compositionend',
    (ev: CompositionEvent) => {
      keyboardSink.textContent = '';
      postCompositionEnterDeadline = ev.timeStamp + POST_COMPOSITION_ENTER_MS;
      if (!compositionInProgress) return;

      visualTexEndCompositionTransaction(ev.data);
    },
    { capture: true, signal }
  );

  keyboardSink.addEventListener(
    'beforeinput',
    (ev) => {
      macTextInput?.beforeInput(ev as InputEvent);
      ev.stopImmediatePropagation();
    },
    { signal }
  );

  // The `input` events is dispatched when the field is changed,
  // but no other relevant events have been triggered
  // for example with emoji input...
  // (the DOM types `input` events as Event; they are InputEvents)
  (keyboardSink as any).addEventListener(
    'input',
    (ev: InputEvent) => {
      if (compositionInProgress) {
        if (ev.isComposing || ev.inputType === 'insertCompositionText') return;
        if (ev.inputType === 'insertFromPaste') {
          visualTexEndCompositionTransaction('');
        } else {
          // Some WebView2 IME paths deliver the committed text as a final
          // non-composing input event but omit compositionend. Treat that
          // event as the missing terminal composition event exactly once.
          const recoveredCompositionData = ev.data ?? '';
          visualTexEndCompositionTransaction(recoveredCompositionData);
          if (recoveredCompositionData)
            visualTexNotifyRecoveredCompositionEnd(recoveredCompositionData);
          ev.preventDefault();
          ev.stopPropagation();
          return;
        }
      }

      keyboardSink.textContent = '';

      // If this was an `input` event sent as a result of a commit of
      // IME, ignore it.
      // (This is what FireFox does, even though the spec says it
      // shouldn't happen). See https://github.com/w3c/uievents/issues/202
      if (ev.inputType === 'insertCompositionText') return;

      // Paste is handled in paste handler
      if (ev.inputType === 'insertFromPaste') {
        ev.preventDefault();
        ev.stopPropagation();
        return;
      }
      if (macTextInput?.input(ev)) return;

      delegate.onInput(ev.data ?? '');

      // Do not propagate the event (it crosses the shadow dom barrier)
      ev.preventDefault();
      ev.stopPropagation();
    },
    { signal }
  );

  keyboardSink.addEventListener(
    'paste',
    (event: ClipboardEvent) => {
      // In some cases (Linux browsers), the keyboard sink might not be focused
      // when doing a middle-click paste command.
      keyboardSink.focus({ preventScroll: true });
      keyboardSink.textContent = '';
      if (!delegate.onPaste(event)) event.preventDefault();
      event.stopImmediatePropagation();
    },
    { signal }
  );

  keyboardSink.addEventListener('cut', (ev) => delegate.onCut(ev), {
    capture: true,
    signal,
  });

  keyboardSink.addEventListener('copy', (ev) => delegate.onCopy(ev), {
    capture: true,
    signal,
  });

  keyboardSink.addEventListener(
    'blur',
    (event) => {
      // Composition is tied to the hidden contenteditable keyboard sink. Once
      // that sink blurs, keeping the transaction alive is invalid and can
      // permanently block all future key/input events if WebView2 omits
      // compositionend during an input-method switch.
      if (compositionInProgress)
        visualTexEndCompositionTransaction('');
      // If we're attempting to focus the mathfield (which can happen on iOS if
      // clicking right on the border of the mathfield) ignore it
      // (preventDefault on the event doesn't work)
      if (((event as FocusEvent)['relatedTarget'] as any)?.['_mathfield']?.['element'] === element) {
        keyboardSink.focus({ preventScroll: true });
        event.preventDefault();
        event.stopPropagation();
        return;
      }

      // If the relatedTarget (the element that is gaining the focus)
      // is contained in our shadow host, ignore the blur event
      if (
        event.relatedTarget ===
        ((event.target as HTMLElement).getRootNode() as any as ShadowRoot).host
      ) {
        event.preventDefault();
        event.stopPropagation();
        return;
      }

      if (blurInProgress || focusInProgress) return;

      blurInProgress = true;
      keydownEvent = null;
      keypressEvent = null;
      delegate.onBlur();
      blurInProgress = false;
    },
    { capture: true, signal }
  );

  keyboardSink.addEventListener(
    'focus',
    (_evt) => {
      if (blurInProgress || focusInProgress) return;

      focusInProgress = true;
      delegate.onFocus();

      focusInProgress = false;
    },
    { capture: true, signal }
  );

  return {
    dispose: (): void => controller.abort(),
    takePostCompositionEnter: (event: KeyboardEvent): boolean => {
      if (compositionInProgress) {
        postCompositionEnterDeadline = -1;
        return false;
      }
      if (event.key === 'Enter' && event.timeStamp <= postCompositionEnterDeadline) {
        postCompositionEnterDeadline = -1;
        return true;
      }
      // a modifier alone keeps the wait (Shift, ...), another key ends it
      if (!/^(?:Shift|Control|Alt|Meta|CapsLock)$/.test(event.key))
        postCompositionEnterDeadline = -1;
      return false;
    },
    cancelComposition: (): void => {
      // Due to the order of event, we may not have
      // received `compositionstart` by the time we decide
      // to cancel the composition.
      if (!compositionInProgress) return;
      // Do not rely on blur/focus to make the browser emit compositionend.
      // Clear MathLive's transaction first, then bounce focus only to dismiss
      // any native IME candidate UI that may still be open.
      visualTexEndCompositionTransaction('');
      if (typeof keyboardSink.blur === 'function') keyboardSink.blur();
      requestAnimationFrame(() => keyboardSink.focus({ preventScroll: true }));
    },

    blur: (): void => {
      if (typeof keyboardSink.blur === 'function') {
        blurInProgress = true;
        keyboardSink.blur();
        blurInProgress = false;
      }
    },

    focus: (): void => {
      if (!focusInProgress && typeof keyboardSink.focus === 'function') {
        focusInProgress = true;
        keyboardSink.focus({ preventScroll: true });
        focusInProgress = false;
      }
    },

    hasFocus: (): boolean => {
      return deepActiveElement() === keyboardSink;
    },

    setAriaLabel: (value: string): void =>
      keyboardSink.setAttribute('aria-label', value),

    setValue: (value: string): void => {
      if (keyboardSink.textContent === value) return;
      keyboardSink.textContent = value;
      // Select the elements in the sink (Safari will not enable copy/paste if there isn't a selection)
      window.getSelection()?.selectAllChildren(keyboardSink);
    },

    // VisualTeX: the sink stands on the caret (its height too), so the IME
    // candidate window opens beside it (MathLive parked the sink off screen)
    moveTo: (x: number, y: number, height?: number): void => {
      keyboardSink.style.top = `${y}px`;
      keyboardSink.style.left = `${x}px`;
      if (height! > 0) {
        keyboardSink.style.height = `${height}px`;
        keyboardSink.style.lineHeight = `${height}px`;
      }
    },
  };
}

export function keyboardEventToChar(evt?: KeyboardEvent): string {
  if (!evt || !mightProducePrintableCharacter(evt)) return '';
  let result: string | undefined;
  if (evt.key === 'Unidentified') {
    // On Android, the evt.key seems to always be 'Unidentified'.
    // Get the value entered in the event target
    if (evt.target) result = (evt.target as HTMLInputElement).value;
  }

  // Note that in some  rare cases, the evt.key can be a string of multiple
  // char. This happens on Windows Swedish keyboard with Firefox, where the
  // `¨` key is returned as `¨¨`.
  result = result ?? evt.key ?? evt.code;
  if (
    /^(Dead|Return|Enter|Tab|Escape|Delete|PageUp|PageDown|Home|End|Help|ArrowLeft|ArrowRight|ArrowUp|ArrowDown)$/.test(
      result
    )
  )
    result = '';

  return result;
}

/**
 * Create a normalized representation of a keyboard event,
 * i.e., key code and modifier keys. For example:
 * - `ctrl+Shift+alt+[KeyF]`
 *
 * Note: the key code corresponds to a physical key, e.g. 'KeyQ' is
 * the key labeled 'A' on a French keyboard
 *
 */
export function keyboardEventToString(evt: KeyboardEvent): string {
  evt = normalizeKeyboardEvent(evt);

  const modifiers: string[] = [];

  if (evt.ctrlKey) modifiers.push('ctrl');
  if (evt.metaKey) modifiers.push('meta');
  if (evt.altKey) modifiers.push('alt');
  if (evt.shiftKey) modifiers.push('shift');

  // If no modifiers, simply return the key name
  if (modifiers.length === 0) return `[${evt.code}]`;

  modifiers.push(`[${evt.code}]`);

  return modifiers.join('+');
}
