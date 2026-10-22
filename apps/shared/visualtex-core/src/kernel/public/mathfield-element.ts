import type { Selector } from './commands';
import type {
  LatexSyntaxError,
  LatexValue,
  MacroDictionary,
  Offset,
  ParseMode,
  Registers,
  Style,
  Selection,
  Range,
  OutputFormat,
  ElementInfo,
  InsertOptions,
} from './core-types';
import type { InsertStyleHook, Mathfield } from './mathfield';
import type {
  InlineShortcutDefinitions,
  Keybinding,
  MathfieldOptions,
} from './options';

import {
  get as getOptions,
  getDefault as getDefaultOptions,
  update as updateOptions,
} from '../editor-mathfield/options';
import { _Mathfield } from '../editor-mathfield/mathfield-private';
import { offsetFromPoint, PointerTracker } from '../editor-mathfield/pointer-input';
import { getElementInfo, getHref } from '../editor-mathfield/utils';
import {
  isBrowser,
  isInIframe,
  isTouchCapable,
} from '../ui/utils/capabilities';
import {
  reparseAllMathfields,
  requestUpdate,
} from '../editor-mathfield/render';
import { loadFonts } from '../core/fonts';

import { l10n } from '../core/l10n';
import { getStylesheet, getStylesheetContent } from '../common/stylesheet';
import {
  isOffset,
  isRange,
  isSelection,
} from '../editor-model/selection-utils';
import { defaultInsertStyleHook } from '../editor-mathfield/styling';
import { _MathEnvironment } from '../core/math-environment';
import { keyboardEventToString } from '../editor/keyboard';
import {
  visualTexAutoExitAccentAfterDirectInput,
  visualTexAutoExitScriptAfterDirectInput,
  visualTexJudgeKeystroke,
} from '../editor-mathfield/keyboard-input';

if (!isBrowser()) {
  console.error(
    `MathLive 0.109.2: this version of the MathLive library is for use in the browser. A subset of the API is available on the server side in the "mathlive-ssr" library. If using server side rendering (with React for example) you may want to do a dynamic import of the MathLive library inside a \`useEffect()\` call.`
  );
}

//
// Custom Events
//

/**
 *  **Event re-targeting**
 *
 *  Some events bubble up through the DOM tree, so that they are detectable by
 *   any element on the page.
 *
 * Bubbling events fired from within shadow DOM are re-targeted so that, to any
 *  listener external to your component, they appear to come from your
 *  component itself.
 *
 *  **Custom Event Bubbling**
 *
 *  By default, a bubbling custom event fired inside shadow DOM will stop
 *  bubbling when it reaches the shadow root.
 *
 *  To make a custom event pass through shadow DOM boundaries, you must set
 *  both the `composed` and `bubbles` flags to true.
 *
 * The `move-out` event signals that the user pressed an **arrow** key or
 * **tab** key but there was no navigation possible inside the mathfield.
 *
 * This event provides an opportunity to handle this situation, for example
 * by focusing an element adjacent to the mathfield.
 *
 * If the event is canceled (i.e. `evt.preventDefault()` is called inside your
 * event handler), the default behavior is to play a "plonk" sound.
 *
 * @category Mathfield
 */
export type MoveOutEvent = {
  direction: 'forward' | 'backward' | 'upward' | 'downward';
};

declare global {
  /**
   * Map the custom event names to types
   * @internal
   */
  interface HTMLElementEventMap {
    // Mathfield Element events
    'mode-change': CustomEvent;
    'mount': Event;
    'unmount': Event;
    'move-out': CustomEvent<MoveOutEvent>;
    'selection-change': Event;
    'undo-state-change': CustomEvent;
  }
}

//
// Deferred State
//
// Operations that modify the state of the mathfield before it has been
// connected to the DOM will be stashed in this object and they
// will be applied to the element when it gets connected to the DOM.
//
const gDeferredState = new WeakMap<
  MathfieldElement,
  {
    value: string | undefined;
    selection: Selection;
    options: Partial<MathfieldOptions>;
  }
>();

/**
 *
 * The `MathfieldElement` class is a DOM element that provides a math input
 * field.
 *
 * It is a subclass of the standard
 * [`HTMLElement`](https://developer.mozilla.org/en-US/docs/Web/API/HTMLElement)
 * class and as such inherits all of its properties and methods, such
 * as `style`, `tabIndex`, `addEventListener()`, `getAttribute()`, etc...
 *
 * The `MathfieldElement` class provides additional properties and methods to
 * control the display and behavior of `<math-field>` elements.
 *
 * **To instantiate a `MathfieldElement`** use the `<math-field>` tag in HTML.
 * You can also instantiate a `MathfieldElement` programmatically using
 * `new MathfieldElement()`.
 *
 * ```javascript
 * // 1. Create a new MathfieldElement
 * const mf = new MathfieldElement();
 *
 * // 2. Attach it to the DOM
 * document.body.appendChild(mf);
 *
 * // 3. Modifying options after the mathfield has been attached to the DOM
 * mf.addEventListener("mount"), () => {
 *  mf.smartFence = true;
 * });
 * ```
 *
 * Read more about customizing the appearance and behavior of the mathfield in
 * the [Customizing the Mathfield](/mathfield/guides/customizing/) guide.
 *
 *
 * #### MathfieldElement CSS Variables
 *
 * **To customize the appearance of the mathfield**, declare the following CSS
 * variables (custom properties) in a ruleset that applies to the mathfield.
 *
 * ```css
 * math-field {
 *  --hue: 10       // Set the highlight color and caret to a reddish hue
 * }
 * ```
 *
 * Alternatively you can set these CSS variables programatically:
 *
 * ```js
 * document.body.style.setProperty("--hue", "10");
 * ```
 *
 * Read more about the [CSS variables](/mathfield/guides/customizing/#css-variables) available for customization.
 *
 * You can customize the appearance and zindex of the virtual keyboard panel
 * with some CSS variables associated with a selector that applies to the
 * virtual keyboard panel container.
 *
 * Read more about [customizing the virtual keyboard appearance](/mathfield/guides/virtual-keyboards/#custom-appearance)
 *
 * #### MathfieldElement CSS Parts
 *
 * In addition to the CSS variables, the mathfield exposes [CSS
 * parts that can be used to style the mathfield](/mathfield/guides/customizing/#mathfield-parts).
 *
 * For example, to hide the menu button:
 *
 * ```css
 * math-field::part(menu-toggle) {
 *    display: none;
 * }
 * ```
 *
 *
 * #### MathfieldElement Attributes
 *
 * An attribute is a key-value pair set as part of the `<math-field>` tag:
 *
 * ```html
 * <math-field letter-shape-style="tex"></math-field>
 * ```
 *
 * The supported attributes are listed in the table below with their
 * corresponding property, which can be changed directly on the
 * `MathfieldElement` object:
 *
 * ```javascript
 *  mf.value = "\\sin x";
 *  mf.letterShapeStyle = "tex";
 * ```
 *
 * The values of attributes and properties are reflected, which means you can
 * change one or the other, for example:
 *
 * ```javascript
 * mf.setAttribute("letter-shape-style",  "french");
 * console.log(mf.letterShapeStyle);
 * // Result: "french"
 *
 * mf.letterShapeStyle ="tex";
 * console.log(mf.getAttribute("letter-shape-style");
 * // Result: "tex"
 * ```
 *
 * An exception is the `value` property, which is not reflected on the `value`
 * attribute. For consistency with other DOM elements, the `value` attribute
 * remains at its initial value.
 *
 *
 * <div className='symbols-table' style={{"--first-col-width":"32ex"}}>
 *
 * | Attribute | Property |
 * |:---|:---|
 * | `disabled` | `mf.disabled` |
 * | `default-mode` | `mf.defaultMode` |
 * | `letter-shape-style` | `mf.letterShapeStyle` |
 * | `min-font-scale` | `mf.minFontScale` |
 * | `max-matrix-cols` | `mf.maxMatrixCols` |
 * | `popover-policy` | `mf.popoverPolicy` |
 * | `math-mode-space` | `mf.mathModeSpace` |
 * | `read-only` | `mf.readOnly` |
 * | `remove-extraneous-parentheses` | `mf.removeExtraneousParentheses` |
 * | `smart-fence` | `mf.smartFence` |
 * | `smart-mode` | `mf.smartMode` |
 * | `smart-superscript` | `mf.smartSuperscript` |
 * | `inline-shortcut-timeout` | `mf.inlineShortcutTimeout` |
 * | `script-depth` | `mf.scriptDepth` |
 * | `value` | `value` |
 *
 * </div>
 *
 * See [more details about these attributes](#mathfieldelementattributes).
 *
 * In addition, the following DOM elements [global attributes](https://developer.mozilla.org/en-US/docs/Web/HTML/Global_attributes)
 * are supported:
 * - `class`
 * - `data-*`
 * - `hidden`
 * - `id`
 * - `item*`
 * - `style`
 * - `tabindex`
 *
 *
 * #### MathfieldElement Events
 *
 * **To listen to these events** use `mf.addEventListener()`. For events with
 * additional arguments, the arguments are available in `event.detail`.
 *
 * <div className='symbols-table' style={{"--first-col-width":"27ex"}}>
 *
 * | Event Name  | Description |
 * |:---|:---|
 * | `beforeinput` | The value of the mathfield is about to be modified.  |
 * | `input` | The value of the mathfield has been modified. This happens on almost every keystroke in the mathfield. The `evt.data` property includes a copy of `evt.inputType`. See `InputEvent` |
 * | `change` | The user has committed the value of the mathfield. This happens when the user presses **Return** or leaves the mathfield. |
 * | `selection-change` | The selection (or caret position) in the mathfield has changed |
 * | `mode-change` | The mode (`math`, `text`) of the mathfield has changed |
 * | `undo-state-change` |  The state of the undo stack has changed. The `evt.detail.type` indicate if a snapshot was taken or an undo performed. |
 * | `read-aloud-status-change` | The status of a read aloud operation has changed |
 * | `before-virtual-keyboard-toggle` | The visibility of the virtual keyboard panel is about to change. The `evt.detail.visible` property indicate if the keyboard will be visible or not. Listen for this event on `window.mathVirtualKeyboard` |
 * | `virtual-keyboard-toggle` | The visibility of the virtual keyboard panel has changed. Listen for this event on `window.mathVirtualKeyboard` |
 * | `geometrychange` | The geometry of the virtual keyboard has changed. The `evt.detail.boundingRect` property is the new bounding rectangle of the virtual keyboard. Listen for this event on `window.mathVirtualKeyboard` |
 * | `blur` | The mathfield is losing focus |
 * | `focus` | The mathfield is gaining focus |
 * | `move-out` | The user has pressed an **arrow** key or the **tab** key, but there is nowhere to go. This is an opportunity to change the focus to another element if desired. <br/> `detail: \{direction: 'forward' | 'backward' | 'upward' | 'downward'\}` **cancellable**|
 * | `keypress` | The user pressed a physical keyboard key |
 * | `mount` | The element has been attached to the DOM |
 * | `unmount` | The element is about to be removed from the DOM |
 *
 * </div>
 * @category Mathfield
 * @keywords zindex, events, attribute, attributes, property, properties, parts, variables, css, mathfield, mathfieldelement
 *
 */
export class MathfieldElement extends HTMLElement implements Mathfield {
  static version = '0.109.2';
  /** @internal */
  static get formAssociated(): boolean {
    return isElementInternalsSupported();
  }
  /**
   * Private lifecycle hooks.
   * If adding a 'boolean' attribute, add its default value to getOptionsFromAttributes
   * @internal
   */
  static get optionsAttributes(): Readonly<
    Record<string, 'number' | 'boolean' | 'string' | 'on/off'>
  > {
    return {
      'default-mode': 'string',
      'letter-shape-style': 'string',
      'min-font-scale': 'number',
      'max-matrix-cols': 'number',
      'popover-policy': 'string',

      'math-mode-space': 'string',
      'read-only': 'boolean',
      'remove-extraneous-parentheses': 'on/off',
      'smart-fence': 'on/off',
      'smart-mode': 'on/off',
      'smart-superscript': 'on/off',
      'inline-shortcut-timeout': 'string',
      'script-depth': 'string',
      'placeholder': 'string',
    };
  }

  /**
   * Custom elements lifecycle hooks
   * @internal
   */
  static get observedAttributes(): readonly string[] {
    return [
      ...Object.keys(this.optionsAttributes),
      'contenteditable', // Global attribute
      'disabled', // Global attribute
      'readonly', // A semi-global attribute (not all standard elements support it, but some do)
      'read-only', // Alternate spelling for `readonly`
    ];
  }

  static openUrl = (href: string): void => {
    if (!href) return;
    const url = new URL(href);
    if (!['http:', 'https:', 'file:'].includes(url.protocol.toLowerCase()))
      return;
    window.open(url, '_blank');
  };

  /**
   * Support for [Trusted Type](https://www.w3.org/TR/trusted-types/).
   *
   * This optional function will be called before a string of HTML is
   * injected in the DOM, allowing that string to be sanitized
   * according to a policy defined by the host.
   *
   * Consider using this option if you are displaying untrusted content. Read more about [Security Considerations](/mathfield/guides/security/)
   *
   */
  static createHTML: (html: string) => any = (x) => x;
  // @todo https://github.com/microsoft/TypeScript/issues/30024

  /**
   * Sets the speed of the selected voice.
   *
   * One of `x-slow`, `slow`, `medium`, `fast`, `x-fast` or a value as a
   * percentage.
   *
   * Range is `20%` to `200%` For example `200%` to indicate a speaking rate
   * twice the default rate.
   * @category Speech
   */
  static get speechEngineRate(): string {
    return this._speechEngineRate;
  }
  static set speechEngineRate(value: string) {
    this._speechEngineRate = value;
  }
  /** @internal */
  private static _speechEngineRate = '100%';

  /**
   * The markup syntax to use for the output of conversion to spoken text.
   *
   * Possible values are `ssml` for the SSML markup or `mac` for the macOS
   * markup, i.e. `&#91;&#91;ltr&#93;&#93;`.
   * @category Speech
   *
   */
  static get textToSpeechMarkup(): '' | 'ssml' | 'ssml_step' | 'mac' {
    return this._textToSpeechMarkup;
  }
  static set textToSpeechMarkup(value: '' | 'ssml' | 'ssml_step' | 'mac') {
    this._textToSpeechMarkup = value;
  }
  /** @internal */
  private static _textToSpeechMarkup: '' | 'ssml' | 'ssml_step' | 'mac' = '';

  /**
   * Specify which set of text to speech rules to use.
   *
   * A value of `mathlive` indicates that the simple rules built into MathLive
   * should be used.
   *
   * A value of `sre` indicates that the Speech Rule Engine from Volker Sorge
   * should be used.
   *
   * **(Caution)** SRE is not included or loaded by MathLive. For this option to
   * work SRE should be loaded separately.
   *
   * **See**
   * {@link mathfield/guides/speech/ | Guide: Speech}
   * @category Speech
   */
  static get textToSpeechRules(): 'mathlive' | 'sre' {
    return this._textToSpeechRules;
  }
  static set textToSpeechRules(value: 'mathlive' | 'sre') {
    this._textToSpeechRules = value;
  }
  /** @internal */
  private static _textToSpeechRules: 'mathlive' | 'sre' = 'mathlive';

  /**
   * A set of key/value pairs that can be used to configure the speech rule
   * engine.
   *
   * Which options are available depends on the speech rule engine in use.
   * There are no options available with MathLive's built-in engine. The
   * options for the SRE engine are documented
   * {@link https://github.com/zorkow/speech-rule-engine | here}
   * @category Speech
   */
  static get textToSpeechRulesOptions(): Readonly<Record<string, string>> {
    return this._textToSpeechRulesOptions;
  }
  static set textToSpeechRulesOptions(value: Record<string, string>) {
    this._textToSpeechRulesOptions = value;
  }
  /** @internal */
  private static _textToSpeechRulesOptions: Record<string, string> = {};

  /**
   * The locale (language + region) to use for string localization.
   *
   * If none is provided, the locale of the browser is used.
   * @category Localization
   *
   */
  static get locale(): string {
    return l10n.locale;
  }
  static set locale(value: string) {
    if (value === 'auto') value = navigator.language.slice(0, 5);
    l10n.locale = value;
  }

  /**
   * When switching from a tab to one that contains a mathfield that was
   * previously focused, restore the focus to the mathfield.
   *
   * This is behavior consistent with `<textarea>`, however it can be
   * disabled if it is not desired.
   *
   * **Default**: `true`
   * @category Customization
   */
  static restoreFocusWhenDocumentFocused = true;

  /**
   * The symbol used to separate the integer part from the fractional part of a
   * number.
   *
   * When `","` is used, the corresponding LaTeX string is `{,}`, in order
   * to ensure proper spacing (otherwise an extra gap is displayed after the
   * comma).
   *
   * This affects:
   * - what happens when the `,` key is pressed (if `decimalSeparator` is
   * `","`, the `{,}` LaTeX string is inserted when following some digits)
   * - the label and behavior of the "." key in the default virtual keyboard
   *
   * **Default**: `"."`
   * @category Localization
   */
  static get decimalSeparator(): ',' | '.' {
    return this._decimalSeparator;
  }
  static set decimalSeparator(value: ',' | '.') {
    this._decimalSeparator = value;
  }

  /** @internal */
  get decimalSeparator(): never {
    throw new Error('Use MathfieldElement.decimalSeparator instead');
  }
  /** @internal */
  set decimalSeparator(_val: unknown) {
    throw new Error('Use MathfieldElement.decimalSeparator instead');
  }

  /** @internal */
  private static _decimalSeparator: ',' | '.' = '.';

  /** The template used to format numbers in scientific notation.
   * The template should include the placeholders `#1` and `#2`, which will
   * be replaced by the significand and exponent, respectively.
   *
   * The template is used when typing a number in scientific notation, e.g.
   * `1.23e4`, which will be rendered as `1.23×10^4`.
   *
   * **Default**: `'#1\\times10^{#2}'`
   *
   * Other common formats include:
   * - `'#1\\,\\mathrm{E}\\mathop{#2}'` (e.g. `1.23\,\mathrm{E}\mathop{-4}`)
   *
   * @category Localization
   */
  static set scientificNotationTemplate(value: string | null) {
    this._scientificNotationTemplate = value;
  }

  static get scientificNotationTemplate(): string | null {
    return this._scientificNotationTemplate;
  }

  /** @internal */
  private static _scientificNotationTemplate: string | null =
    '#1\\times10^{#2}';

  /**
   * When using the keyboard to navigate a fraction, the order in which the
   * numerator and navigator are traversed:
   * - `"numerator-denominator"`: first the elements in the numerator, then
   *   the elements in the denominator.
   * - `"denominator-numerator"`: first the elements in the denominator, then
   *   the elements in the numerator. In some East-Asian cultures, fractions
   *   are read and written denominator first ("fēnzhī"). With this option
   *   the keyboard navigation follows this convention.
   *
   * **Default**: `"numerator-denominator"`
   * @category Localization
   */
  static get fractionNavigationOrder():
    | 'numerator-denominator'
    | 'denominator-numerator' {
    return _MathEnvironment.fractionNavigationOrder;
  }
  static set fractionNavigationOrder(
    s: 'numerator-denominator' | 'denominator-numerator'
  ) {
    if (s !== 'numerator-denominator' && s !== 'denominator-numerator')
      throw new Error('Invalid value');
    if (_MathEnvironment.fractionNavigationOrder === s) return;

    _MathEnvironment.fractionNavigationOrder = s;

    // Invalidate all mathfields on the page
    reparseAllMathfields();
  }

  /** @internal */
  private static _isFunction: (command: string) => boolean = () => false;

  static get isFunction(): (command: string) => boolean {
    if (typeof this._isFunction !== 'function') return () => false;
    return this._isFunction;
  }

  static set isFunction(value: (command: string) => boolean) {
    this._isFunction = value;

    reparseAllMathfields();
  }

  /** @internal */
  private _mathfield!: null | _Mathfield;

  /** @internal
   * Supported by some browser: allows some (static) attributes to be set
   * without being reflected on the element instance.
   */
  private _internals!: ElementInternals;

  /** @internal */
  private _observer: MutationObserver | null = null;

  // The content of <style> tags inside the element.
  /** @internal */
  private _style!: string;

  /**
     * To create programmatically a new mathfield use:
     *
     ```javascript
    let mfe = new MathfieldElement();

    // Set initial value and options
    mfe.value = "\\frac{\\sin(x)}{\\cos(x)}";

    // Options can be set either as an attribute (for simple options)...
    mfe.setAttribute("letter-shape-style", "french");

    // ... or using properties
    mfe.letterShapeStyle = "french";

    // Attach the element to the DOM
    document.body.appendChild(mfe);
    ```
    */
  constructor(options?: Partial<MathfieldOptions>) {
    super();

    if (isElementInternalsSupported()) {
      this._internals = this.attachInternals();
      this._internals['role'] = 'math';
      this._internals.ariaLabel = 'math input field';
      this._internals.ariaMultiLine = 'false';
    }

    this.attachShadow({ mode: 'open', delegatesFocus: true });

    if (this.shadowRoot && 'adoptedStyleSheets' in this.shadowRoot) {
      // @ts-ignore
      this.shadowRoot!.adoptedStyleSheets = [
        getStylesheet('core'),
        getStylesheet('mathfield'),
        getStylesheet('mathfield-element'),
        getStylesheet('ui'),
      ];

      // @ts-ignore
      this.shadowRoot!.appendChild(document.createElement('span'));

      const slot = document.createElement('slot');
      slot.style.display = 'none';
      // @ts-ignore
      this.shadowRoot!.appendChild(slot);
    } else {
      // @ts-ignore
      this.shadowRoot!.innerHTML =
        '<style>' +
        getStylesheetContent('core') +
        getStylesheetContent('mathfield') +
        getStylesheetContent('mathfield-element') +
        getStylesheetContent('ui') +
        '</style>' +
        '<span></span><slot style="display:none"></slot>';
    }

    // Record the (optional) configuration options, as a deferred state
    if (options) this._setOptions(options);
  }

  /**
   * VisualTeX kernel: runs the input judgment for a keystroke the host
   * handles before the kernel sees it (Space). Returns true when the
   * keystroke is used up (it finished a converted word).
   */
  visualTexJudgeKeystroke(evt: KeyboardEvent): boolean {
    const mf = this._mathfield;
    return mf ? visualTexJudgeKeystroke(mf, evt, keyboardEventToString(evt)) : false;
  }

  /** VisualTeX kernel: lays the formula out and draws it again (the formula fonts changed). */
  visualTexRender(): void {
    requestUpdate(this._mathfield);
  }

  /**
   * VisualTeX kernel: inserts a toolbar template. A template typed straight
   * into a script or accent (no placeholder left) exits it the way direct
   * input does.
   */
  visualTexInsertToolbarTemplate(latex: string, options?: InsertOptions): boolean {
    const mathfield = this._mathfield;
    if (!mathfield) return false;
    const model = mathfield.model;
    const before = model.at(Math.max(model.position, model.anchor));
    const parent = before?.parent;
    const branch = before?.parentBranch;
    const inserted = this.insert(latex, options);
    if (inserted && options?.selectionMode !== 'placeholder') {
      const atom = model.at(model.position);
      if (atom?.parent === parent && atom?.parentBranch === branch) {
        if (
          visualTexAutoExitAccentAfterDirectInput(mathfield, atom) ||
          visualTexAutoExitScriptAfterDirectInput(mathfield, atom)
        )
          requestUpdate(mathfield);
      }
    }
    return inserted;
  }

  /** VisualTeX kernel: selects from `anchor` to `focus` (the focus keeps the caret). */
  visualTexSelectRange(anchor: Offset, focus: Offset): void {
    const mathfield = this._mathfield;
    if (!mathfield) return;
    mathfield.model.extendSelectionTo(anchor, focus);
    requestUpdate(mathfield);
  }

  /**
   * VisualTeX kernel: puts document fragments before and after the formula
   * (they keep their own style) and keeps the selection on the same content.
   */
  visualTexInsertDocumentEdges(prefix: string, suffix: string): void {
    if (!this._mathfield) return;
    const selection: Selection = {
      ranges: this.selection.ranges.map((range) => [...range] as Range),
      direction: this.selection.direction,
    };
    const options: InsertOptions = {
      mode: 'math',
      format: 'latex',
      selectionMode: 'after',
      style: {},
      silenceNotifications: true,
      visualTexPreserveStyle: true,
    };
    if (prefix) {
      const length = this.lastOffset;
      this.position = 0;
      this.insert(prefix, { ...options });
      const shift = this.lastOffset - length;
      selection.ranges = selection.ranges.map(
        ([a, b]) => [a + shift, b + shift] as Range
      );
    }
    if (suffix) {
      this.position = this.lastOffset;
      this.insert(suffix, { ...options });
    }
    this.selection = selection;
  }

  /** VisualTeX kernel: ends the field's own drag tracking (the host took the drag over). */
  visualTexStopPointerTracking(): void {
    const field = this._mathfield?.field;
    if (field && PointerTracker.trackedElement === field) PointerTracker.cancel();
  }

  /** @internal */
  onPointerDown(): void {
    window.addEventListener(
      'pointerup',
      (evt) => {
        const mf = this._mathfield;
        if (!mf) return;
        // Disabled elements do not dispatch 'click' events
        if (evt.target === this && !mf.disabled) {
          this.dispatchEvent(
            new MouseEvent('click', {
              altKey: evt.altKey,
              button: evt.button,
              buttons: evt.buttons,
              clientX: evt.clientX,
              clientY: evt.clientY,
              ctrlKey: evt.ctrlKey,
              metaKey: evt.metaKey,
              movementX: evt.movementX,
              movementY: evt.movementY,
              relatedTarget: evt.relatedTarget,
              screenX: evt.screenX,
              screenY: evt.screenY,
              shiftKey: evt.shiftKey,
            })
          );

          // Check if a \href command was clicked on (or its children)
          const offset = this.getOffsetFromPoint(evt.clientX, evt.clientY);
          if (offset >= 0) MathfieldElement.openUrl(getHref(mf, offset));
          // set cursor position if selection is collapsed on touch events
          if (evt.pointerType === 'touch' && this.selectionIsCollapsed)
            this.position = offset;
        }
      },
      { once: true }
    );
  }

  /**
   * Return the content of the `\placeholder{}` command with the `placeholderId`
   * @category Prompts
   */
  getPromptValue(placeholderId: string, format?: OutputFormat): string {
    return this._mathfield?.getPromptValue(placeholderId, format) ?? '';
  }

  /** @category Prompts */
  setPromptValue(
    id: string,
    content: string,
    insertOptions: Omit<InsertOptions, 'insertionMode'>
  ): void {
    this._mathfield?.setPromptValue(id, content, insertOptions);
  }

  /**
   * Return the selection range for the specified prompt.
   *
   * This can be used for example to select the content of the prompt.
   *
   * ```js
   * mf.selection = mf.getPromptRange('my-prompt-id');
   * ```
   *
   * @category Prompts
   *
   */

  getPromptRange(id: string): Range | null {
    return this._mathfield?.getPromptRange(id) ?? null;
  }

  /** Return the id of the prompts matching the filter.
   * @category Prompts
   */
  getPrompts(filter?: {
    id?: string;
    locked?: boolean;
    correctness?: 'correct' | 'incorrect' | 'undefined';
  }): string[] {
    return this._mathfield?.getPrompts(filter) ?? [];
  }

  /** True if the mathfield has editable content, such as unlocked prompts */
  get hasEditableContent(): boolean {
    return this._mathfield?.hasEditableContent ?? false;
  }

  /** @internal */
  get form(): HTMLFormElement | null {
    return this._internals?.['form'];
  }

  /** @internal */
  get name(): string {
    return this.getAttribute('name') ?? '';
  }

  /** @internal */
  get type(): string {
    return this.localName;
  }

  get mode(): ParseMode {
    return (
      this._mathfield?.model.mode ??
      (this.defaultMode === 'text' ? 'text' : 'math')
    );
  }

  set mode(value: ParseMode) {
    this._mathfield?.switchMode(value);
  }

  /**
   * Return an array of LaTeX syntax errors, if any.
   * @category Accessing and changing the content
   */
  get errors(): readonly LatexSyntaxError[] {
    return this._mathfield?.errors ?? [];
  }

  /** @internal */
  private _getOptions<K extends keyof MathfieldOptions>(
    keys: K[]
  ): Pick<MathfieldOptions, K>;
  private _getOptions(): MathfieldOptions;
  private _getOptions(
    keys?: keyof MathfieldOptions | (keyof MathfieldOptions)[]
  ): null | Partial<MathfieldOptions> {
    if (this._mathfield) return getOptions(this._mathfield.options, keys);

    if (!gDeferredState.has(this)) return null;
    return {
      ...getOptions(
        {
          ...getDefaultOptions(),
          ...updateOptions(gDeferredState.get(this)!.options),
        },
        keys
      ),
    };
  }

  /** @internal */
  private reflectAttributes() {
    // options by name (attributes map to options by name)
    const defaultOptions: Record<string, any> = getDefaultOptions();
    const options: Record<string, any> = this._getOptions();
    Object.keys(MathfieldElement.optionsAttributes).forEach((x) => {
      // Handle special case: 'placeholder' attribute maps to 'contentPlaceholder' option
      const prop = x === 'placeholder' ? 'contentPlaceholder' : toCamelCase(x);
      if (MathfieldElement.optionsAttributes[x] === 'on/off') {
        if (defaultOptions[prop] !== options[prop])
          this.setAttribute(x, options[prop] ? 'on' : 'off');
        else this.removeAttribute(x);
      } else if (defaultOptions[prop] !== options[prop]) {
        if (MathfieldElement.optionsAttributes[x] === 'boolean') {
          if (options[prop]) {
            // Add attribute
            this.setAttribute(x, '');
          } else {
            // Remove attribute
            this.removeAttribute(x);
          }
        } else {
          // Set attribute (as string)
          if (
            typeof options[prop] === 'string' ||
            typeof options[prop] === 'number'
          )
            this.setAttribute(x, options[prop].toString());
        }
      }
    });
  }

  /** @internal */
  private _getOption<K extends keyof MathfieldOptions>(
    key: K
  ): MathfieldOptions[K] {
    return this._getOptions([key])[key];
  }

  /** @internal */
  private _setOptions(options: Partial<MathfieldOptions>): void {
    if (this._mathfield) this._mathfield.setOptions(options);
    else if (gDeferredState.has(this)) {
      const mergedOptions = {
        ...gDeferredState.get(this)!.options,
        ...options,
      };
      gDeferredState.set(this, {
        ...gDeferredState.get(this)!,
        selection: { ranges: mergedOptions.readOnly ? [[0, 0]] : [[0, -1]] },
        options: mergedOptions,
      });
    } else {
      gDeferredState.set(this, {
        value: undefined,
        selection: { ranges: [[0, 0]] },
        options,
      });
    }

    // Reflect options to attributes
    this.reflectAttributes();
  }

  /**
   * Execute a [`command`](#commands) defined by a selector.
   * ```javascript
   * mfe.executeCommand('add-column-after');
   * mfe.executeCommand(['switch-mode', 'math']);
   * ```
   *
   * @param selector - A selector, or an array whose first element
   * is a selector, and whose subsequent elements are arguments to the selector.
   *
   * Selectors can be passed either in camelCase or kebab-case.
   *
   * ```javascript
   * // Both calls do the same thing
   * mfe.executeCommand('selectAll');
   * mfe.executeCommand('select-all');
   * ```
   * @category Commands
   */
  executeCommand(selector: Selector): boolean;
  executeCommand(selector: Selector, ...args: unknown[]): boolean;
  executeCommand(selector: [Selector, ...unknown[]]): boolean;
  executeCommand(...args: unknown[]): boolean {
    let selector: Selector | [Selector, ...unknown[]];
    if (args.length === 1)
      selector = args[0] as Selector | [Selector, ...unknown[]];
    else selector = [args[0] as Selector, ...args.slice(1)];

    if (selector) return this._mathfield?.executeCommand(selector) ?? false;

    throw new Error('Invalid selector');
  }

  /**
   * Return a textual representation of the content of the mathfield.
   *
   * @param format - The format of the result.
   *
   * **Default:** `"latex"`
   *
   * @category Accessing and changing the content
   */
  getValue(format?: OutputFormat): string;
  getValue(start: Offset, end: Offset, format?: OutputFormat): string;
  getValue(range: Range, format?: OutputFormat): string;
  getValue(selection: Selection, format?: OutputFormat): string;
  getValue(
    arg1?: Offset | Range | Selection | OutputFormat,
    arg2?: Offset | OutputFormat,
    arg3?: OutputFormat
  ): string {
    if (this._mathfield)
      return this._mathfield.model.getValue(arg1 as any, arg2 as any, arg3);

    if (gDeferredState.has(this)) {
      let start: Offset;
      let end: Offset;
      let format: OutputFormat | undefined = undefined;
      if (isSelection(arg1)) {
        [start, end] = arg1.ranges[0];
        format = arg2 as OutputFormat;
      } else if (isRange(arg1)) {
        [start, end] = arg1;
        format = arg2 as OutputFormat;
      } else if (isOffset(arg1) && isOffset(arg2)) {
        start = arg1;
        end = arg2;
        format = arg3;
      } else {
        start = 0;
        end = -1;
        format = arg1 as OutputFormat;
      }

      if (
        (format === undefined || format === 'latex') &&
        start === 0 &&
        end === -1
      )
        return gDeferredState.get(this)!.value ?? this.textContent ?? '';
    }

    return '';
  }

  /**
   * Set the content of the mathfield to the text interpreted as a
   * LaTeX expression.
   *
   * @category Accessing and changing the content
   */
  setValue(value?: string, options?: InsertOptions): void {
    if (this._mathfield && value !== undefined) {
      const currentValue = this._mathfield.model.getValue();
      // VisualTeX: an explicit mode is part of the requested editor state,
      // not merely a parsing hint. A field can have identical LaTeX while its
      // live model is still stuck in text mode from a previous caret/IME
      // context. Do not let the same-value fast path preserve that stale mode.
      if (currentValue === value) {
        const requestedMode = options?.mode;
        if (
          requestedMode &&
          requestedMode !== 'auto' &&
          this._mathfield.model.mode !== requestedMode
        ) {
          this._mathfield.model.mode = requestedMode;
          requestUpdate(this._mathfield);
        }
        return;
      }

      options ??= { silenceNotifications: true, mode: 'math' };
      this._mathfield.setValue(value, options);

      return;
    }

    if (gDeferredState.has(this)) {
      const options = gDeferredState.get(this)!.options;
      gDeferredState.set(this, {
        value,
        selection: { ranges: [[-1, -1]], direction: 'forward' },
        options,
      });
      return;
    }

    const attrOptions = getOptionsFromAttributes(this);
    gDeferredState.set(this, {
      value,
      selection: { ranges: [[-1, -1]], direction: 'forward' },
      options: attrOptions,
    });
  }

  /**
   * Return true if the mathfield is currently focused (responds to keyboard
   * input).
   *
   * @category Focus
   *
   */
  hasFocus(): boolean {
    return this._mathfield?.hasFocus() ?? false;
  }

  /**
   * Sets the focus to the mathfield (will respond to keyboard input).
   *
   * @category Focus
   *
   */
  focus(): void {
    if (this.disabled) return;
    this._mathfield?.focus();
  }

  /**
   * Remove the focus from the mathfield (will no longer respond to keyboard
   * input).
   *
   * @category Focus
   *
   */
  blur(): void {
    this._mathfield?.blur();
  }

  /**
   * Select the content of the mathfield.
   * @category Selection
   */
  select(): void {
    this._mathfield?.select();
  }

  /**
   * Insert a block of text at the current insertion point.
   *
   * This method can be called explicitly or invoked as a selector with
   * `executeCommand("insert")`.
   *
   * After the insertion, the selection will be set according to the
   * `options.selectionMode`.

   *  @category Accessing and changing the content
   */
  insert(s: string, options?: InsertOptions): boolean {
    return this._mathfield?.insert(s, options) ?? false;
  }

  /**
   * Update the style (color, bold, italic, etc...) of the selection or sets
   * the style to be applied to future input.
   *
   * If there is no selection and no range is specified, the style will
   * apply to the next character typed.
   *
   * If a range is specified, the style is applied to the range, otherwise,
   * if there is a selection, the style is applied to the selection.
   *
   * If the operation is `"toggle"` and the range already has this style,
   * remove it. If the range
   * has the style partially applied (i.e. only some sections), remove it from
   * those sections, and apply it to the entire range.
   *
   * If the operation is `"set"`, the style is applied to the range,
   * whether it already has the style or not.
   *
   * The default operation is `"set"`.
   *
   *
   *
   * @category Styles
   */
  applyStyle(
    style: Readonly<Style>,
    options?: Range | { range?: Range; operation?: 'set' | 'toggle' }
  ): void {
    return this._mathfield?.applyStyle(style, options);
  }

  /**
   * If there is a selection, return if all the atoms in the selection,
   * some of them or none of them match the `style` argument.
   *
   * If there is no selection, return 'all' if the current implicit style
   * (determined by a combination of the style of the previous atom and
   * the current style) matches the `style` argument, 'none' if it does not.
   *
   * @category Styles
   */
  queryStyle(style: Readonly<Style>): 'some' | 'all' | 'none' {
    return this._mathfield?.queryStyle(style) ?? 'none';
  }

  /** The offset closest to the location `(x, y)` in viewport coordinate.
   *
   * **`bias`**:  if `0`, the vertical midline is considered to the left or
   * right sibling. If `-1`, the left sibling is favored, if `+1`, the right
   * sibling is favored.
   *
   * @category Selection
   */
  getOffsetFromPoint(
    x: number,
    y: number,
    options?: { bias?: -1 | 0 | 1 }
  ): Offset {
    if (!this._mathfield) return -1;
    return offsetFromPoint(this._mathfield, x, y, options);
  }

  getElementInfo(offset: Offset): ElementInfo | undefined {
    return getElementInfo(this._mathfield, offset);
  }

  /**
   * Reset the undo stack
   *
   * @category Undo
   */
  resetUndo(): void {
    this._mathfield?.resetUndo();
  }

  /**
   * Return whether there are undoable items
   * @category Undo
   */
  canUndo(): boolean {
    if (!this._mathfield) return false;
    return this._mathfield.canUndo();
  }

  /**
   * Return whether there are redoable items
   * @category Undo
   */
  canRedo(): boolean {
    if (!this._mathfield) return false;
    return this._mathfield.canRedo();
  }

  /** @internal */
  handleEvent(evt: Event): void {
    // only once the field is mounted
    if (!this._mathfield) return;

    if (evt.type === 'pointerdown') this.onPointerDown();

    // The private mathfield handles focus/blur events directly,
    // so we don't need to call focus() or blur() here
    if (evt.type === 'focus') return;

    // Ignore blur events if we're in an iFrame on a touch device (see #2350).
    if (evt.type !== 'blur') return;
    if (isTouchCapable() && isInIframe()) return;

    // Call onBlur directly to handle the blur, without dispatching events
    // (the DOM already dispatched them)
    this._mathfield?.onBlur({ dispatchEvents: false });
  }

  /**
   * Custom elements lifecycle hooks
   * @internal
   */
  connectedCallback(): void {
    const shadowRoot = this.shadowRoot!;
    const host = shadowRoot.host;
    const computedStyle = window.getComputedStyle(this);
    const userSelect = computedStyle.userSelect !== 'none';

    if (userSelect) host.addEventListener('pointerdown', this, true);
    else {
      const span = shadowRoot.querySelector('span');
      span!.style.pointerEvents = 'none';
    }
    // Listen for an element *inside* the mathfield to get focus, e.g. the virtual keyboard toggle
    host.addEventListener('focus', this, true);
    host.addEventListener('blur', this, true);

    // Create an observer instance to detect when the innerHTML or textContent
    // of the element is modified
    this._observer = new MutationObserver(() => {
      this.value = this.textContent ?? '';
    });
    this._observer.observe(this, {
      childList: true,
      characterData: true,
      subtree: true,
    });

    if (!isElementInternalsSupported()) {
      if (!this.hasAttribute('role')) this.setAttribute('role', 'math');
      if (!this.hasAttribute('aria-label'))
        this.setAttribute('aria-label', 'math input field');
      this.setAttribute('aria-multiline', 'false');
    }

    // NVDA on Firefox seems to require this attribute
    if (userSelect && !this.hasAttribute('contenteditable'))
      this.setAttribute('contenteditable', 'true');

    // When the elements get focused (through tabbing for example)
    // focus the mathfield
    if (!this.hasAttribute('tabindex')) this.setAttribute('tabindex', '0');

    const slot = shadowRoot.querySelector<HTMLSlotElement>('slot:not([name])');
    if (slot) {
      try {
        this._style = slot
          .assignedElements()
          .filter((x) => x.tagName.toLowerCase() === 'style')
          .map((x) => x.textContent)
          .join('');
      } catch (error: unknown) {
        console.error(error);
      }
    }
    // Add shadowed stylesheet if one was provided
    // (this is important to support the `\class{}{}` command)
    if (this._style) {
      const styleElement = document.createElement('style');
      styleElement.textContent = this._style;
      shadowRoot.appendChild(styleElement);
    }

    let value = '';
    // Check if there is a `value` attribute and set the initial value
    // of the mathfield from it
    if (this.hasAttribute('value')) value = this.getAttribute('value')!;
    else {
      value =
        slot
          ?.assignedNodes()
          .map((x) => (x.nodeType === 3 ? x.textContent : ''))
          .join('')
          .trim() ?? '';
    }

    this._mathfield = new _Mathfield(
      shadowRoot.querySelector(':host > span')!,
      {
        ...(gDeferredState.get(this)?.options ??
          getOptionsFromAttributes(this)),
        eventSink: this,
        value,
      }
    );

    if (!gDeferredState.has(this)) {
      this.upgradeProperty('disabled');
      this.upgradeProperty('readonly');
      for (const attr of Object.keys(MathfieldElement.optionsAttributes))
        this.upgradeProperty(toCamelCase(attr));
    }

    // The mathfield creation could have failed
    if (!this._mathfield?.model) {
      this._mathfield = null;
      return;
    }

    if (gDeferredState.has(this)) {
      const mf = this._mathfield!;
      const state = gDeferredState.get(this)!;
      mf.model.deferNotifications({ content: false, selection: false }, () => {
        const value = state.value;
        if (value !== undefined) mf.setValue(value);
        mf.model.selection = state.selection;

        gDeferredState.delete(this);
      });
    }

    // Notify listeners that we're mounted and ready
    window.queueMicrotask(() => {
      if (!this.isConnected) return;
      this.dispatchEvent(
        new Event('mount', {
          cancelable: false,
          bubbles: true,
          composed: true,
        })
      );
    });

    // Load the fonts
    void loadFonts();
  }

  /**
   * Custom elements lifecycle hooks
   * @internal
   */
  disconnectedCallback(): void {
    this.shadowRoot!.host.removeEventListener('pointerdown', this, true);

    if (!this._mathfield) return;

    this._observer?.disconnect();
    this._observer = null;

    window.queueMicrotask(() =>
      // Notify listeners that we have been unmounted
      this.dispatchEvent(
        new Event('unmount', {
          cancelable: false,
          bubbles: true,
          composed: true,
        })
      )
    );

    // Save the state (in case the element gets reconnected later)
    const options = getOptions(
      this._mathfield.options,
      Object.keys(MathfieldElement.optionsAttributes).map((x) => toCamelCase(x))
    );
    gDeferredState.set(this, {
      value: this._mathfield.getValue(),
      selection: this._mathfield.model.selection,
      options,
    });

    // Dispose of the mathfield
    this._mathfield.dispose();
    this._mathfield = null;
  }

  /**
   * Private lifecycle hooks
   * @internal
   */
  upgradeProperty(prop: string): void {
    if (this.hasOwnProperty(prop)) {
      const value: unknown = (this as Record<string, any>)[prop];
      // A property may have already been set on the object, before
      // the element was connected: delete the property (after saving its value)
      // and use the setter to (re-)set its value.
      delete (this as Record<string, any>)[prop];
      if (prop === 'readonly' || prop === 'read-only') prop = 'readOnly';
      (this as Record<string, any>)[prop] = value;
    }
  }

  /**
   * Custom elements lifecycle hooks
   * @internal
   */
  attributeChangedCallback(
    name: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    if (oldValue === newValue) return;
    const hasValue: boolean = newValue !== null;
    switch (name) {
      case 'contenteditable':
        requestUpdate(this._mathfield);
        break;
      case 'placeholder':
        if (newValue === false) newValue = '';
        this.placeholder = newValue as string;
        break;

      case 'disabled':
        this.disabled = hasValue;
        break;
      case 'read-only':
      case 'readonly':
        this.readOnly = hasValue;
        break;
      default:
    }
  }

  get readonly(): boolean {
    return this.hasAttribute('readonly') || this.hasAttribute('read-only');
  }

  set readonly(value: boolean) {
    const isReadonly = Boolean(value);

    // Note that `readonly` is a "boolean attribute" as
    // per the HTML5 spec. Its value must be the empty string to indicate
    // a value of true, or it must be absent to indicate a value of false.
    // https://html.spec.whatwg.org/#boolean-attribute
    if (isReadonly) {
      // The canonical spelling is "readonly" (no dash. It's a global attribute
      // name and follows HTML attribute conventions)
      this.setAttribute('readonly', '');
      if (isElementInternalsSupported()) this._internals.ariaReadOnly = 'true';
      else this.setAttribute('aria-readonly', 'true');

      this.setAttribute('aria-readonly', 'true');
    } else {
      if (isElementInternalsSupported()) this._internals.ariaReadOnly = 'false';
      else this.removeAttribute('aria-readonly');

      this.removeAttribute('readonly');
      this.removeAttribute('read-only');
    }

    this._setOptions({ readOnly: isReadonly });
  }

  get disabled(): boolean {
    return this.hasAttribute('disabled');
  }

  set disabled(value: boolean) {
    const isDisabled = Boolean(value);
    if (isDisabled) this.setAttribute('disabled', '');
    else this.removeAttribute('disabled');

    if (isElementInternalsSupported())
      this._internals.ariaDisabled = isDisabled ? 'true' : 'false';
    else this.setAttribute('aria-disabled', isDisabled ? 'true' : 'false');
  }

  /**
   * The content of the mathfield as a LaTeX expression.
   * ```js
   * document.querySelector('mf').value = '\\frac{1}{\\pi}'
   * ```
   *  @category Accessing and changing the content
   */
  get value(): string {
    return this.getValue();
  }
  set value(value: string) {
    this.setValue(value);
  }

  /**
   * The mode of the element when it is empty:
   * - `"math"`: equivalent to `\displaystyle` (display math mode)
   * - `"inline-math"`: equivalent to `\inlinestyle` (inline math mode)
   * - `"text"`: text mode
   * @category Customization
   */
  get defaultMode(): 'inline-math' | 'math' | 'text' {
    return this._getOption('defaultMode');
  }
  set defaultMode(value: 'inline-math' | 'math' | 'text') {
    this._setOptions({ defaultMode: value });
  }

  /**
   * A dictionary of LaTeX macros to be used to interpret and render the content.
   *
   * For example, to add a new macro to the default macro dictionary:
   *
```javascript
mf.macros = {
  ...mf.macros,
  smallfrac: '^{#1}\\!\\!/\\!_{#2}',
};
```
   *
   * Note that `...mf.macros` is used to keep the existing macros and add to
   * them.
   * Otherwise, all the macros are replaced with the new definition.
   *
   * The code above will support the following notation:
   *
    ```tex
    \smallfrac{5}{16}
    ```
    @category Macros
   */
  get macros(): Readonly<MacroDictionary> {
    if (!this._mathfield) throw new Error('Mathfield not mounted');
    return this._getOption('macros');
  }
  set macros(value: MacroDictionary) {
    this._setOptions({ macros: value });
  }

  /**
   * @inheritDoc Registers
   * @category Registers
   */
  get registers(): Registers {
    if (!this._mathfield) throw new Error('Mathfield not mounted');
    const that = this;
    return new Proxy(
      {},
      {
        get: (_, prop): number | string | LatexValue | undefined => {
          if (typeof prop !== 'string') return undefined;
          return that._getOption('registers')[prop];
        },
        set(_, prop, value): boolean {
          if (typeof prop !== 'string') return false;
          that._setOptions({
            registers: { ...that._getOption('registers'), [prop]: value },
          });
          return true;
        },
        ownKeys: (): (string | symbol)[] =>
          Reflect.ownKeys(that._getOption('registers')),
        getOwnPropertyDescriptor: (_, prop) => {
          const value = that._getOption('registers')[prop as string];
          if (value !== undefined) {
            return {
              configurable: true,
              enumerable: true,
              value,
              writable: true,
            };
          }
          return undefined;
        },
      }
    );
  }

  set registers(value: Registers) {
    this._setOptions({ registers: value });
  }

  /**
   * Map a color name as used in commands such as `\textcolor{}{}` or
   * `\colorbox{}{}` to a CSS color value.
   *
   * Use this option to override the standard mapping of colors such as "yellow"
   * or "red".
   *
   * If the name is not one you expected, return `undefined` and the default
   * color mapping will be applied.
   *
   * If a `backgroundColorMap()` function is not provided, the `colorMap()`
   * function will be used instead.
   *
   * If `colorMap()` is not provided, default color mappings are applied.
   *
   * The following color names have been optimized for a legible foreground
   * and background values, and are recommended:
   * - `red`, `orange`, `yellow`, `lime`, `green`, `teal`, `blue`, `indigo`,
   * `purple`, `magenta`, `black`, `dark-grey`, `grey`, `light-grey`, `white`
   *
   * @category Customization
   */
  get colorMap(): (name: string) => string | undefined {
    return this._getOption('colorMap');
  }
  set colorMap(value: (name: string) => string | undefined) {
    this._setOptions({ colorMap: value });
  }

  /** @category Customization */
  get backgroundColorMap(): (name: string) => string | undefined {
    return this._getOption('backgroundColorMap');
  }
  set backgroundColorMap(value: (name: string) => string | undefined) {
    this._setOptions({ backgroundColorMap: value });
  }

  /**
  * Control the letter shape style:

  | `letterShapeStyle` | xyz | ABC | αβɣ | ΓΔΘ |
  | ------------------ | --- | --- | --- | --- |
  | `iso`              | it  | it  | it  | it  |
  | `tex`              | it  | it  | it  | up  |
  | `french`           | it  | up  | up  | up  |
  | `upright`          | up  | up  | up  | up  |

  (it) = italic (up) = upright

  * The default letter shape style is `auto`, which indicates that `french`
  * should be used if the locale is "french", and `tex` otherwise.
  *
  * **Historical Note**
  *
  * Where do the "french" rules come from? The
  * TeX standard font, Computer Modern, is based on Monotype 155M, itself
  * based on the Porson greek font which was one of the most widely used
  * Greek fonts in english-speaking countries. This font had upright
  * capitals, but slanted lowercase. In France, the traditional font for
  * greek was Didot, which has both upright capitals and lowercase.
  *
  * As for roman uppercase, they are recommended by "Lexique des règles
  * typographiques en usage à l’Imprimerie Nationale". It should be noted
  * that this convention is not universally followed.
  *
  * @category Customization
*/
  get letterShapeStyle(): 'auto' | 'tex' | 'iso' | 'french' | 'upright' {
    return this._getOption('letterShapeStyle');
  }
  set letterShapeStyle(value: 'auto' | 'tex' | 'iso' | 'french' | 'upright') {
    this._setOptions({ letterShapeStyle: value });
  }

  /**
   * Set the minimum relative font size for nested superscripts and fractions. The value
   * should be a number between `0` and `1`. The size is in relative `em` units relative to the
   * font size of the `math-field` element. Specifying a value of `0` allows the `math-field`
   * to use its default sizing logic.
   *
   * **Default**: `0`
   *
   * @category Customization
   */
  get minFontScale(): number {
    return this._getOption('minFontScale');
  }
  set minFontScale(value: number) {
    this._setOptions({ minFontScale: value });
  }

  /**
   * Sets the maximum number of columns for the matrix environment. The default is
   * 10 columns to match the behavior of the amsmath matrix environment.
   * **Default**: `10`
   *
   * @category Customization
   */
  get maxMatrixCols(): number {
    return this._getOption('maxMatrixCols');
  }
  set maxMatrixCols(value: number) {
    this._setOptions({ maxMatrixCols: value });
  }

  /**
   * When `true`, during text input the field will switch automatically between
   * 'math' and 'text' mode depending on what is typed and the context of the
   * formula. If necessary, what was previously typed will be 'fixed' to
   * account for the new info.
   *
   * For example, when typing "if x >0":
   *
   * | Type  | Interpretation |
   * |---:|:---|
   * | `i` | math mode, imaginary unit |
   * | `if` | text mode, english word "if" |
   * | `if x` | all in text mode, maybe the next word is xylophone? |
   * | `if x >` | "if" stays in text mode, but now "x >" is in math mode |
   * | `if x > 0` | "if" in text mode, "x > 0" in math mode |
   *
   * **Default**: `false`
   *
   * Manually switching mode (by typing `alt/option+=`) will temporarily turn
   * off smart mode.
   *
   *
   * **Examples**
   *
   * - `slope = rise/run`
   * - `If x > 0, then f(x) = sin(x)`
   * - `x^2 + sin (x) when x > 0`
   * - `When x&lt;0, x^{2n+1}&lt;0`
   * - `Graph x^2 -x+3 =0 for 0&lt;=x&lt;=5`
   * - `Divide by x-3 and then add x^2-1 to both sides`
   * - `Given g(x) = 4x – 3, when does g(x)=0?`
   * - `Let D be the set {(x,y)|0&lt;=x&lt;=1 and 0&lt;=y&lt;=x}`
   * - `\int\_{the unit square} f(x,y) dx dy`
   * - `For all n in NN`
   *
   * @category Customization
   */
  get smartMode(): boolean {
    return this._getOption('smartMode');
  }
  set smartMode(value: boolean) {
    this._setOptions({ smartMode: value });
  }

  /**
   *
   * When `true` and an open fence is entered via `typedText()` it will
   * generate a contextually appropriate markup, for example using
   * `\left...\right` if applicable.
   *
   * When `false`, the literal value of the character will be inserted instead.
   * @category Customization
   */
  get smartFence(): boolean {
    return this._getOption('smartFence');
  }
  set smartFence(value: boolean) {
    this._setOptions({ smartFence: value });
  }

  /**
   * When `true` and a digit is entered in an empty superscript, the cursor
   * leaps automatically out of the superscript. This makes entry of common
   * polynomials easier and faster. If entering other characters (for example
   * "n+1") the navigation out of the superscript must be done manually (by
   * using the cursor keys or the spacebar to leap to the next insertion
   * point).
   *
   * When `false`, the navigation out of the superscript must always be done
   * manually.
   * @category Customization
   */
  get smartSuperscript(): boolean {
    return this._getOption('smartSuperscript');
  }
  set smartSuperscript(value: boolean) {
    this._setOptions({ smartSuperscript: value });
  }

  /**
   * This option controls how many levels of subscript/superscript can be entered. For
   * example, if `scriptDepth` is "1", there can be one level of superscript or
   * subscript. Attempting to enter a superscript while inside a superscript will
   * be rejected. Setting a value of 0 will prevent entry of any superscript or
   * subscript (but not limits for sum, integrals, etc...)
   *
   * This can make it easier to enter equations that fit what's expected for the
   * domain where the mathfield is used.
   *
   * To control the depth of superscript and subscript independently, provide an
   * array: the first element indicate the maximum depth for subscript and the
   * second element the depth of superscript. Thus, a value of `[0, 1]` would
   * suppress the entry of subscripts, and allow one level of superscripts.
   * @category Customization
   */
  get scriptDepth(): number | [number, number] {
    return this._getOption('scriptDepth');
  }
  set scriptDepth(value: number | [number, number]) {
    this._setOptions({ scriptDepth: value });
  }

  /**
   * If `true`, extra parentheses around a numerator or denominator are
   * removed automatically.
   *
   * **Default**: `true`
   * @category Customization
   */
  get removeExtraneousParentheses(): boolean {
    return this._getOption('removeExtraneousParentheses');
  }
  set removeExtraneousParentheses(value: boolean) {
    this._setOptions({ removeExtraneousParentheses: value });
  }

  /**
   * The LaTeX string to insert when the spacebar is pressed (on the physical or
   * virtual keyboard).
   *
   * Use `"\;"` for a thick space, `"\:"` for a medium space, `"\,"` for a
   * thin space.
   *
   * Do not use `" "` (a regular space), as whitespace is skipped by LaTeX
   * so this will do nothing.
   *
   * **Default**: `""` (empty string)
   * @category Customization
   */
  get mathModeSpace(): string {
    return this._getOption('mathModeSpace');
  }
  set mathModeSpace(value: string) {
    this._setOptions({ mathModeSpace: value });
  }

  /**
   * The symbol used to represent a placeholder in an expression.
   *
   * **Default**: `▢` `U+25A2 WHITE SQUARE WITH ROUNDED CORNERS`
   *
   * @category Customization
   */

  get placeholderSymbol(): string {
    return this._getOption('placeholderSymbol');
  }
  set placeholderSymbol(value: string) {
    this._setOptions({ placeholderSymbol: value });
  }

  /**
   * A LaTeX string displayed inside the mathfield when there is no content.
   * @category Customization
   */
  get placeholder(): string {
    return this._getOption('contentPlaceholder');
  }

  set placeholder(value: string) {
    this._setOptions({ contentPlaceholder: value });
  }

  /**
   * If `"auto"` a popover with suggestions may be displayed when a LaTeX
   * command is input.
   *
   * **Default**: `"auto"`
   *
   * @category Customization
   */
  get popoverPolicy(): 'auto' | 'off' {
    return this._getOption('popoverPolicy');
  }
  set popoverPolicy(value: 'auto' | 'off') {
    this._setOptions({ popoverPolicy: value });
  }

  /**
   * VisualTeX shows its own context menu; the field has none. Accepted (and
   * ignored) for compatibility with callers that clear MathLive's menu.
   */
  get menuItems(): readonly unknown[] {
    return [];
  }

  set menuItems(_menuItems: readonly unknown[]) {}

  /**
   * The keys of this object literal indicate the sequence of characters
   * that will trigger an inline shortcut.
   * @category Keyboard Shortcuts
   */
  get inlineShortcuts(): Readonly<InlineShortcutDefinitions> {
    if (!this._mathfield) throw new Error('Mathfield not mounted');
    return this._getOption('inlineShortcuts');
  }
  set inlineShortcuts(value: InlineShortcutDefinitions) {
    if (!this._mathfield) throw new Error('Mathfield not mounted');
    this._setOptions({ inlineShortcuts: value });
  }

  /**
   * Maximum time, in milliseconds, between consecutive characters for them to be
   * considered part of the same shortcut sequence.
   *
   * A value of 0 is the same as infinity: any consecutive character will be
   * candidate for an inline shortcut, regardless of the interval between this
   * character and the previous one.
   *
   * A value of 750 will indicate that the maximum interval between two
   * characters to be considered part of the same inline shortcut sequence is
   * 3/4 of a second.
   *
   * This is useful to enter "+-" as a sequence of two characters, while also
   * supporting the "±" shortcut with the same sequence.
   *
   * The first result can be entered by pausing slightly between the first and
   * second character if this option is set to a value of 250 or so.
   *
   * Note that some operations, such as clicking to change the selection, or
   * losing the focus on the mathfield, will automatically timeout the
   * shortcuts.
   * @category Keyboard Shortcuts
   */
  get inlineShortcutTimeout(): number {
    return this._getOption('inlineShortcutTimeout');
  }
  set inlineShortcutTimeout(value: number) {
    this._setOptions({ inlineShortcutTimeout: value });
  }

  /** @category Keyboard Shortcuts   */
  get keybindings(): readonly Keybinding[] {
    if (!this._mathfield) throw new Error('Mathfield not mounted');
    return this._getOption('keybindings');
  }
  set keybindings(value: readonly Keybinding[]) {
    if (!this._mathfield) throw new Error('Mathfield not mounted');
    this._setOptions({ keybindings: value });
  }

  /** @category Styles
   */
  get onInsertStyle(): InsertStyleHook | undefined | null {
    const hook = this._getOption('onInsertStyle');
    if (hook === undefined) return defaultInsertStyleHook as InsertStyleHook;
    return hook;
  }
  set onInsertStyle(value: InsertStyleHook | undefined | null) {
    this._setOptions({ onInsertStyle: value });
  }

  /**
   * A hook invoked when a string of characters that could be
   * interpreted as shortcut has been typed.
   *
   * If not a special shortcut, return the empty string `""`.
   *
   * Use this handler to detect multi character symbols, and return them wrapped appropriately,
   * for example `\mathrm{${symbol}}`.
   * @category Hooks
   */
  get onInlineShortcut(): (sender: Mathfield, symbol: string) => string {
    return this._getOption('onInlineShortcut');
  }
  set onInlineShortcut(value: (sender: Mathfield, symbol: string) => string) {
    this._setOptions({ onInlineShortcut: value });
  }

  /**
   * A hook invoked when scrolling the mathfield into view is necessary.
   *
   * Use when scrolling the page would not solve the problem, e.g.
   * when the mathfield is in another div that has scrollable content.
   * @category Hooks
   */
  get onScrollIntoView(): ((sender: Mathfield) => void) | null {
    return this._getOption('onScrollIntoView');
  }
  set onScrollIntoView(value: ((sender: Mathfield) => void) | null) {
    this._setOptions({ onScrollIntoView: value });
  }

  /**
   * This hook is invoked when the user has requested to export the content
   * of the mathfield, for example when pressing ctrl/command+C.
   *
   * This hook should return as a string what should be exported.
   *
   * The `range` argument indicates which portion of the mathfield should be
   * exported. It is not always equal to the current selection, but it can
   * be used to export a format other than LaTeX.
   *
   * By default this is:
   *
   * ```js
   *  return `\\begin{equation*}${latex}\\end{equation*}`;
   * ```
   * @category Hooks
   */
  get onExport(): (from: Mathfield, latex: string, range: Range) => string {
    return this._getOption('onExport');
  }
  set onExport(
    value: (from: Mathfield, latex: string, range: Range) => string
  ) {
    this._setOptions({ onExport: value });
  }

  get readOnly(): boolean {
    return this._getOption('readOnly');
  }
  set readOnly(value: boolean) {
    this._setOptions({ readOnly: value });
  }

  /**
   * This is a standard DOM property
   * @internal */
  get isSelectionEditable(): boolean {
    return this._mathfield?.isSelectionEditable ?? false;
  }

  /** @category Prompts */
  setPromptState(
    id: string,
    state: 'correct' | 'incorrect' | 'undefined' | undefined,
    locked?: boolean
  ): void {
    this._mathfield?.setPromptState(id, state, locked);
  }
  /** @category Prompts */
  getPromptState(id: string): ['correct' | 'incorrect' | undefined, boolean] {
    return this._mathfield?.getPromptState(id) ?? [undefined, true];
  }

  /**
   * An array of ranges representing the selection.
   *
   * It is guaranteed there will be at least one element. If a discontinuous
   * selection is present, the result will include more than one element.
   *
   * @category Selection
   *
   */
  get selection(): Readonly<Selection> {
    if (this._mathfield) return this._mathfield.model.selection;

    if (gDeferredState.has(this)) return gDeferredState.get(this)!.selection;

    return { ranges: [[0, 0]], direction: 'forward' };
  }

  /**
   *
   * @category Selection
   */
  set selection(sel: Selection | Offset) {
    if (typeof sel === 'number') sel = { ranges: [[sel, sel]] };

    if (this._mathfield) {
      this._mathfield.model.selection = sel;
      requestUpdate(this._mathfield);
      return;
    }

    if (gDeferredState.has(this)) {
      gDeferredState.set(this, {
        ...gDeferredState.get(this)!,
        selection: sel,
      });
      return;
    }

    gDeferredState.set(this, {
      value: undefined,
      selection: sel,
      options: getOptionsFromAttributes(this),
    });
  }

  /**
   * @category Selection
   */

  get selectionIsCollapsed(): boolean {
    const selection = this.selection;
    return (
      selection.ranges.length === 1 &&
      selection.ranges[0][0] === selection.ranges[0][1]
    );
  }

  /**
   * The position of the caret/insertion point, from 0 to `lastOffset`.
   *
   * @category Selection
   *
   */
  get position(): Offset {
    if (this._mathfield) return this._mathfield.model.position;

    if (gDeferredState.has(this))
      return gDeferredState.get(this)!.selection.ranges[0][0];

    return 0;
  }

  /**
   * @category Selection
   */
  set position(offset: Offset) {
    if (this._mathfield) {
      this._mathfield.model.position = offset;
      requestUpdate(this._mathfield);
    }

    if (gDeferredState.has(this)) {
      gDeferredState.set(this, {
        ...gDeferredState.get(this)!,
        selection: { ranges: [[offset, offset]] },
      });
      return;
    }

    gDeferredState.set(this, {
      value: undefined,
      selection: { ranges: [[offset, offset]] },
      options: getOptionsFromAttributes(this),
    });
  }

  /**
   * The last valid offset.
   * @category Selection
   */
  get lastOffset(): Offset {
    return this._mathfield?.model.lastOffset ?? -1;
  }
}

function toCamelCase(s: string): string {
  return s.replace(/[^a-zA-Z\d]+(.)/g, (_m, c) => c.toUpperCase());
}

// Function toKebabCase(s: string): string {
//     return s
//         .match(
//             /[A-Z]{2,}(?=[A-Z][a-z]+[0-9]*|\b)|[A-Z]?[a-z]+[0-9]*|[A-Z]|[0-9]+/g
//         )
//         .map((x: string) => x.toLowerCase())
//         .join('-');
// }

function getOptionsFromAttributes(
  mfe: MathfieldElement
): Partial<MathfieldOptions> {
  const result: Partial<MathfieldOptions> & Record<string, any> = { readOnly: false };
  const attribs = MathfieldElement.optionsAttributes;
  Object.keys(attribs).forEach((x) => {
    if (mfe.hasAttribute(x)) {
      let value = mfe.getAttribute(x);

      if (x === 'placeholder') result.contentPlaceholder = value ?? '';
      else if (attribs[x] === 'boolean') result[toCamelCase(x)] = true;
      else if (attribs[x] === 'on/off') {
        value = value?.toLowerCase() ?? '';
        if (value === 'on' || value === 'true') result[toCamelCase(x)] = true;
        else if (value === 'off' || value === 'false')
          result[toCamelCase(x)] = false;
        else result[toCamelCase(x)] = undefined;
      } else if (attribs[x] === 'number')
        result[toCamelCase(x)] = Number.parseFloat(value ?? '0');
      else result[toCamelCase(x)] = value;
    }
    // else if (attribs[x] === 'boolean') result[toCamelCase(x)] = false;
  });
  return result;
}

function isElementInternalsSupported(): boolean {
  if (!('ElementInternals' in window) || !HTMLElement.prototype.attachInternals)
    return false;
  if (!('role' in window.ElementInternals.prototype)) return false;
  return true;
}

export default MathfieldElement;

type MathfieldElementClass = typeof MathfieldElement;

declare global {
  interface Window {
    MathfieldElement: MathfieldElementClass;
  }
  // The element class, also published on the global object (the kernel's
  // modules read options such as the decimal separator from it)
  // eslint-disable-next-line no-var
  var MathfieldElement: MathfieldElementClass;
}

if (isBrowser() && !window.customElements?.get('math-field')) {
  // The `window[Symbol.for('io.cortexjs.mathlive')]` global is used  to coordinate between mathfield
  // instances that may have been instantiated by different versions of the
  // library
  (window as any)[Symbol.for('io.cortexjs.mathlive')] ??= {};
  const global = (window as any)[Symbol.for('io.cortexjs.mathlive')];
  global.version = '0.109.2';

  globalThis.MathfieldElement = MathfieldElement;
  window.customElements?.define('math-field', MathfieldElement);
}
