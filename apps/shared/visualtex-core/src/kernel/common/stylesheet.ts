import MATHFIELD_STYLESHEET from '../css/mathfield.css?raw';

import CORE_STYLESHEET from '../css/core.css?raw';

import SUGGESTION_POPOVER_STYLESHEET from '../css/suggestion-popover.css?raw';

import UI_STYLESHEET from '../css/ui.css?raw';


type StylesheetId =
  | 'ui'
  | 'core'
  | 'mathfield-element'
  | 'mathfield'
  | 'suggestion-popover';

let gStylesheets: Partial<Record<StylesheetId, CSSStyleSheet>>;

export function getStylesheetContent(id: StylesheetId): string {
  let content = '';

  switch (id) {
    //
    // Note: the `position: relative` is required to fix https://github.com/arnog/mathlive/issues/971
    //
    case 'mathfield-element':
      content = `
    :host { display: inline-block; background-color: field; color: fieldtext; border-width: 1px; border-style: solid; border-color: #acacac; border-radius: 2px;}
    :host([hidden]) { display: none; }
    :host([disabled]), :host([disabled]:focus), :host([disabled]:focus-within) { outline: none; opacity:  .5; }
    :host(:focus), :host(:focus-within) {
      outline: Highlight auto 1px;    /* For Firefox */
      outline: -webkit-focus-ring-color auto 1px;
    }
    :host([readonly]:focus), :host([readonly]:focus-within),
    :host([read-only]:focus), :host([read-only]:focus-within) {
      outline: none;
    }`;
      break;
    case 'core':
      content = CORE_STYLESHEET;
      break;
    case 'mathfield':
      content = MATHFIELD_STYLESHEET;
      break;
    case 'suggestion-popover':
      content = SUGGESTION_POPOVER_STYLESHEET;
      break;
    case 'ui':
      content = UI_STYLESHEET;
      break;
    default:
      debugger;
  }
  return content;
}

export function getStylesheet(id: StylesheetId): CSSStyleSheet {
  if (!gStylesheets) gStylesheets = {};

  if (gStylesheets[id]) return gStylesheets[id]!;

  gStylesheets[id] = new CSSStyleSheet();

  gStylesheets[id]!.replaceSync(getStylesheetContent(id));

  return gStylesheets[id]!;
}

let gInjectedStylesheets: Partial<Record<StylesheetId, number>>;

export function injectStylesheet(id: StylesheetId): void {
  try {
    if (!('adoptedStyleSheets' in document)) {
      if (window.document.getElementById(`mathlive-style-${id}`)) return;
      const styleNode = window.document.createElement('style');
      styleNode.id = `mathlive-style-${id}`;
      styleNode.append(
        window.document.createTextNode(getStylesheetContent(id))
      );
      window.document.head.appendChild(styleNode);
      return;
    }

    if (!gInjectedStylesheets) gInjectedStylesheets = {};
    if ((gInjectedStylesheets[id] ?? 0) !== 0) gInjectedStylesheets[id]! += 1;
    else {
      const stylesheet = getStylesheet(id);
      document.adoptedStyleSheets = [
        ...document.adoptedStyleSheets,
        stylesheet,
      ];
      gInjectedStylesheets[id] = 1;
    }
  } catch (error) {
    console.error('Error injecting stylesheet', id, error);
  }
}

export function releaseStylesheet(id: StylesheetId): void {
  if (!('adoptedStyleSheets' in document)) return;

  if (!gInjectedStylesheets?.[id]) return;

  gInjectedStylesheets[id]! -= 1;
  if (gInjectedStylesheets[id]! <= 0) {
    const stylesheet = gStylesheets[id]!;
    document.adoptedStyleSheets = document.adoptedStyleSheets.filter(
      (x) => x !== stylesheet
    );
  }
}
