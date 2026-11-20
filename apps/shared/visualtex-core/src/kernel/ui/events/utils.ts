
const PRINTABLE_KEYCODE = new Set([
  'Backquote', // Japanese keyboard: hankaku/zenkaku/kanji key, which is non-printable
  'Digit0',
  'Digit1',
  'Digit2',
  'Digit3',
  'Digit4',
  'Digit5',
  'Digit6',
  'Digit7',
  'Digit8',
  'Digit9',
  'Minus',
  'Equal',
  'IntlYen', // Japanese Keyboard. Russian keyboard: \/

  'KeyQ', // AZERTY keyboard: labeled 'a'
  'KeyW', // AZERTY keyboard: labeled 'z'
  'KeyE',
  'KeyR',
  'KeyT',
  'KeyY', // QWERTZ keyboard: labeled 'z'
  'KeyU',
  'KeyI',
  'KeyO',
  'KeyP',
  'BracketLeft',
  'BracketRight',
  'Backslash', // May be labeled #~ on UK 102 keyboard
  'KeyA', // AZERTY keyboard: labeled 'q'
  'KeyS',
  'KeyD',
  'KeyF',
  'KeyG',
  'KeyH',
  'KeyJ',
  'KeyK',
  'KeyL',
  'Semicolon',
  'Quote',
  'IntlBackslash', // QWERTZ keyboard '><'
  'KeyZ', // AZERTY: 'w', QWERTZ: 'y'
  'KeyX',
  'KeyC',
  'KeyV',
  'KeyB',
  'KeyN',
  'KeyM',
  'Comma',
  'Period',
  'Slash',
  'IntlRo', // Japanse keyboard '\ろ'

  'Space',

  'Numpad0',
  'Numpad1',
  'Numpad2',
  'Numpad3',
  'Numpad4',
  'Numpad5',
  'Numpad6',
  'Numpad7',
  'Numpad8',
  'Numpad9',
  'NumpadAdd',
  'NumpadComma',
  'NumpadDecimal',
  'NumpadDivide',
  'NumpadEqual',
  'NumpadHash',
  'NumpadMultiply',
  'NumpadParenLeft',
  'NumpadParenRight',
  'NumpadStar',
  'NumpadSubstract',
]);

export function mightProducePrintableCharacter(evt: KeyboardEvent): boolean {
  // Ignore ctrl/cmd-combinations but not shift/alt-combinations
  if (evt.ctrlKey || evt.metaKey) return false;

  // https://developer.mozilla.org/en-US/docs/Web/API/KeyboardEvent/key/Key_Values
  if (['Dead', 'Process'].includes(evt.key)) return false;

  // When issued via a composition, the `code` field is empty
  if (evt.code === '') return true;

  return PRINTABLE_KEYCODE.has(evt.code);
}

export function deepActiveElement(): HTMLOrSVGElement | null {
  let a = document.activeElement;
  while (a?.shadowRoot?.activeElement) a = a.shadowRoot.activeElement;

  return a as unknown as HTMLOrSVGElement;
}
