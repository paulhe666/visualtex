import type {
  ContentChangeOptions,
  Keybinding,
  KeyboardLayoutName,
} from '../public/options';
import type { Mathfield } from '../public/mathfield';
import type {
  InsertOptions,
  OutputFormat,
  Offset,
  Range,
  Selection,
  ApplyStyleOptions,
} from '../public/core-types';


import { Atom } from '../core/atom-class';
import { gFontsState } from '../core/fonts';
import { defaultBackgroundColorMap, defaultColorMap } from '../core/color';
import {
  getMacroDefinition,
  getMacros,
} from '../latex-commands/definitions-utils';
import { LatexGroupAtom } from '../atoms/latex';
import { parseLatex, validateLatex } from '../core/parser';
import { getDefaultRegisters } from '../core/registers';

import { applyStyle } from '../editor-model/styling';
import { getMode, isRange, range } from '../editor-model/selection-utils';
import {
  removeComposition,
  updateComposition,
} from '../editor-model/composition';
import {
  addRowAfter,
  addColumnAfter,
  type VisualTexPersistentTypingStyle,
} from '../editor-model/array';

import {
  delegateKeyboardEvents,
  KeyboardDelegate,
  KeyboardDelegateInterface,
} from '../editor/keyboard';
import { UndoManager } from '../editor/undo';
import {
  disposeSuggestionPopover,
  hideSuggestionPopover,
  updateSuggestionPopoverPosition,
} from '../editor/suggestion-popover';
import { l10n } from '../core/l10n';
import { SelectorPrivate, perform } from '../editor/commands';
import {
  _MathfieldOptions,
  update as updateOptions,
  getDefault as getDefaultOptions,
  get as getOptions,
  effectiveMode,
} from './options';
import { normalizeKeybindings } from '../editor/keybindings';
import {
  setKeyboardLayoutLocale,
  getActiveKeyboardLayout,
  gKeyboardLayout,
  getDefaultKeyboardLayout,
} from '../editor/keyboard-layout';

import { onInput, onKeystroke } from './keyboard-input';
import { complete } from './autocomplete';
import {
  requestUpdate,
  render,
  renderSelection,
  contentMarkup,
  reparse,
} from './render';

import './commands';
import './styling';
import { isValidMathfield, Rect } from './utils';

import { onPointerDown } from './pointer-input';

import { ModeEditor } from './mode-editor';
import './mode-editor-math';
import './mode-editor-text';

import { computeInsertStyle, validateStyle } from './styling';
import { PromptAtom } from '../atoms/prompt';
import '../public/mathfield-element';


import type {
  ParseMode,
  Style,
  NormalizedMacroDictionary,
  LatexSyntaxError,
} from '../public/core-types';
import type { ContextInterface, PrivateStyle } from '../core/types';
import type { ModelState } from '../editor-model/types';
import { _Model } from '../editor-model/model-private';
import { deleteRange } from '../editor-model/delete';

import '../editor-model/commands-delete';
import '../editor-model/commands-move';
import '../editor-model/commands-select';
import MathfieldElement from '../public/mathfield-element';
import { parseMathString } from '../formats/parse-math-string';
import { TextAtom } from '../atoms/text';
import { getLatexGroup } from './mode-editor-latex';
import {
  keepCaretInScrollingAncestor,
  scheduleCaretFollowUp,
  visualTexAnchorKeyboardSink,
} from './render';
import {
  visualTexFreshStructuralSlot,
  visualTexLiveStructuralSlot,
} from '../editor-model/delete';
import { visualTexEditingCore } from './visualtex-hooks';
import {
  visualTexAutoExitAccentAfterDirectInput,
  visualTexAutoExitAfterComposition,
  visualTexAutoExitScriptAfterDirectInput,
  visualTexCaretIsInsideTextRun,
  visualTexFinalizeSourceWrapper,
  visualTexFinalizeSourceWrapperOnCaretExit,
  visualTexJudgeInput,
  visualTexNormalizeTextBoundaryMode,
  visualTexSourceWrapperOwnsAtom,
  visualTexSyncSourceWrapper,
} from './keyboard-input';

/** A style read and written by property name */
type AnyStyle = Record<string, unknown>;

/** @internal */
export class _Mathfield implements Mathfield, KeyboardDelegateInterface {
  /**
   * Global tracker for the currently focused mathfield.
   * Used to handle cases where browsers don't fire blur events reliably
   * (e.g., when focus() is called rapidly on multiple mathfields).
   */
  private static _globallyFocusedMathfield: _Mathfield | undefined;

  readonly model: _Model;

  readonly undoManager: UndoManager;

  options: Readonly<Required<_MathfieldOptions>>;

  // When inserting new characters, if not `"none"`, adopt the style
  // (color, up variant, etc..) from the previous or following atom.
  styleBias: 'left' | 'right' | 'none';

  /** VisualTeX: the style new glyphs are typed in (set by the app) */
  visualTexPersistentTypingStyle?: VisualTexPersistentTypingStyle;

  /** VisualTeX: the last input judgment, undone by an immediate Backspace */
  visualTexLastJudgment?: { after: string; revert: string; position: number };
  /** VisualTeX: scripts that already auto-exited once (`branch:id`) */
  visualTexConsumedScriptAutoExitKeys?: Set<string>;

  // The style used when `styleBias` is set to 'none'
  defaultStyle: Readonly<Style>;

  dirty: boolean; // If true, need to be redrawn

  element: HTMLElement & { mathfield?: _Mathfield };

  /** The element from which events are emitted, usually a MathfieldElement */
  readonly host: HTMLElement | undefined;

  container!: HTMLElement;
  field!: HTMLElement;
  readonly ariaLiveText: HTMLElement | undefined;
  // readonly accessibleMathML: HTMLElement;

  atomBoundsCache?: Map<string, Rect>;

  // VisualTeX editing state kept on the field
  /** The command being typed in the LaTeX box as a source wrapper (keyboard-input.ts) */
  visualTexSourceWrapper?: any;
  /** The empty slot a command typed in the LaTeX box is to fill */
  visualTexPendingStructuralCommandSlot?: any;
  visualTexPendingPlaceholderSelection?: any;
  /** The candidates of the suggestion list and the typed command they answer */
  visualTexCompletion?: { command: string; suggestions: string[]; group: any } | null;
  visualTexCompletionRecords?: Map<string, any>;
  /** IME composition: the state to return to if it is cancelled, and where it started */
  visualTexCompositionState?: any;
  visualTexCompositionStructuralSlot?: any;
  visualTexCompositionStartedInsideText?: boolean;
  visualTexCompositionCommit?: boolean;
  /** A composition inside a source wrapper: the wrapper and the atom it types after */
  visualTexCompositionSourceWrapper?: any;
  visualTexCompositionSourceAfter?: any;
  /** A selection-change notification held back until the input event went out */
  visualTexDeferredSelectionAfterInput?: boolean;
  /** Pointer: the event already handled (field and host both receive it), drag auto-scroll */
  visualTexLastPointerEvent?: Event;
  visualTexDragScrolling?: boolean;
  /** Keeping the caret in view (render.ts) */
  caretScrollPending?: boolean;
  caretScrollObserver!: ResizeObserver;

  suggestionIndex: number;


  readonly keyboardDelegate!: Readonly<KeyboardDelegate>;

  _keybindings?: readonly Keybinding[]; // Normalized keybindings (raw ones in config)
  keyboardLayout!: KeyboardLayoutName;

  inlineShortcutBuffer: {
    state: ModelState;
    keystrokes: string[];
    leftSiblings: Atom[];
  }[];
  inlineShortcutBufferFlushTimer: ReturnType<typeof setTimeout> | 0;
  scientificNotationTimer: ReturnType<typeof setTimeout>;

  private blurred: boolean;



  // The value of the mathfield when it is focussed.
  // If this value is different when the field is blured
  // the `change` event is dispatched
  private valueOnFocus!: string;
  private focusBlurInProgress = false;
  private programmaticFocusInProgress = false;

  private geometryChangeTimer!: ReturnType<typeof requestAnimationFrame>;

  /** When true, the mathfield is listening to the virtual keyboard */

  private eventController: AbortController;
  resizeObserver!: ResizeObserver;
  // ResizeObserver, sadly, fire at least once, even if no size has changed.
  // Workaround this by tracking its state. See https://github.com/WICG/resize-observer/issues/38#issuecomment-532833484
  resizeObserverStarted!: boolean;

  /**
   * @param element - The DOM element that this mathfield is attached to.
   * Note that `element.mathfield` is this object.
   */
  constructor(
    element: HTMLElement & { mathfield?: _Mathfield },
    options: Partial<_MathfieldOptions> & {
      eventSink?: HTMLElement;
    }
  ) {
    // Setup default config options
    this.options = {
      ...getDefaultOptions(),
      macros: getMacros(),
      registers: getDefaultRegisters(),
      ...updateOptions(options),
    };

    this.eventController = new AbortController();
    const signal = this.eventController.signal;

    if (options.eventSink) this.host = options.eventSink;

    this.element = element;
    element.mathfield = this;

    // Focus/blur state
    this.blurred = true;

    // This index indicates which of the suggestions available to
    // display in the popover panel
    this.suggestionIndex = 0;

    this.inlineShortcutBuffer = [];
    this.inlineShortcutBufferFlushTimer = 0 as unknown as ReturnType<
      typeof setTimeout
    >;
    this.scientificNotationTimer = 0 as unknown as ReturnType<
      typeof setTimeout
    >;

    // Default style (color, weight, italic, etc...):
    // reflects the style to be applied on next insertion
    // if styleBias is "none".
    this.defaultStyle = {};
    // Adopt the style of the left sibling by default
    this.styleBias = 'left';

    if (this.options.defaultMode === 'inline-math')
      this.element.classList.add('ML__is-inline');
    else this.element.classList.remove('ML__is-inline');

    this.dirty = false;

    // Use the content of the element for the initial value of the mathfield
    let elementText = options.value ?? this.element.textContent ?? '';
    elementText = elementText.trim();

    // The initial input mode (text or math): the mode the next character
    // typed will be interpreted in, which may be different from the mode
    // of the current selection.
    const mode = effectiveMode(this.options);

    // Setup the model
    const body = parseLatex(elementText, { context: this.context });
    let root: Atom;
    if (body.length === 1 && body[0].isRoot) root = body[0];
    else root = new Atom({ type: 'root', mode, body });

    this.model = new _Model(this, mode, root);

    // Prepare to manage undo/redo
    this.undoManager = new UndoManager(this.model);

    // Additional elements used for UI.
    const markup: string[] = [];

    // const accessibleNodeID =
    //   Date.now().toString(36).slice(-2) +
    //   Math.floor(Math.random() * 0x186a0).toString(36);
    // Add "aria-labelledby="${accessibleNodeID}"" to the keyboard sink

    // 1/ The keyboard event capture element.
    markup.push(
      `<span contenteditable=true role=textbox aria-autocomplete=none aria-multiline=false part=keyboard-sink class=ML__keyboard-sink autocapitalize=off autocomplete=off autocorrect=off spellcheck=false inputmode=none tabindex=0></span>`
    );

    // 2/ The field, where the math equation will be displayed

    // Start with hidden content to minimize flashing during creation
    // The visibility will be reset during render
    markup.push(
      '<span part=container class=ML__container  style="visibility:hidden">'
    );
    markup.push('<span part=content class=ML__content>');
    markup.push(contentMarkup(this));
    markup.push('</span>');

    markup.push('</span>'); // end container

    // 3.1/ The aria-live region for announcements

    markup.push('<span class=ML__sr-only>');
    markup.push(
      '<span role=status aria-live=assertive aria-atomic=true></span>'
    );

    // markup.push(
    //   `<span class=accessibleMathML id="${accessibleNodeID}"></span>`
    // );
    markup.push('</span>');

    this.element.innerHTML = globalThis.MathfieldElement.createHTML(
      markup.join('')
    );
    if (!this.element.children) {
      console.error(
        `%cMathLive 0.109.2: Something went wrong and the mathfield could not be created.%c
If you are using Vue, this may be because you are using the runtime-only build of Vue. Make sure to include \`runtimeCompiler: true\` in your Vue configuration. There may a warning from Vue in the log above.`,
        'color:red;font-family:system-ui;font-size:1.2rem;font-weight:bold',
        'color:inherit;font-family:system-ui;font-size:inherit;font-weight:inherit'
      );
      return;
    }

    this.container = this.element.querySelector('[part=container]')!;
    this.field = this.container.querySelector('[part=content]')!;

    // Listen to 'click' events on the part of the field that doesn't have
    // content, so we avoid sending two 'click' events
    this.field.addEventListener(
      'click',
      (evt) => evt.stopImmediatePropagation(),
      { capture: false, signal }
    );

    // Listen to 'wheel' events to scroll (horizontally) the field when it overflows
    this.field.addEventListener('wheel', this, { passive: false, signal });

    // Delegate pointer events
    const pointerTarget = this.element;
    if ('PointerEvent' in window)
      pointerTarget.addEventListener('pointerdown', this, { signal });
    else pointerTarget.addEventListener('mousedown', this, { signal });

    // Capture pointer events that occur on the shadow host (i.e. padding around the container)
    if (this.host) {
      const hostPointerHandler = (evt: Event): void => {
        // a shadow-retargeted event that already reached the field
        const path = typeof evt.composedPath === 'function' ? evt.composedPath() : [];
        if (path.includes(this.element!)) return;
        if (evt.target !== this.host) return;
        this.handleEvent(evt);
      };
      if ('PointerEvent' in window) {
        this.host.addEventListener('pointerdown', hostPointerHandler, {
          signal,
        });
      } else {
        this.host.addEventListener('mousedown', hostPointerHandler, {
          signal,
        });
      }
    }

    this.ariaLiveText =
      this.element.querySelector<HTMLElement>('[role=status]') ?? undefined;
    // this.accessibleMathML = this.element.querySelector('.accessibleMathML')!;

    // Capture clipboard events
    // Delegate keyboard events
    this.keyboardDelegate = delegateKeyboardEvents(
      this.element.querySelector('.ML__keyboard-sink')!,
      this.element,
      this
    );

    // VisualTeX: the field's key entry. Every keydown on the field (aimed at
    // the keyboard sink, or dispatched on the field itself) is offered first
    // to the app's key commands, which the editing core keeps for the field;
    // the kernel's own key handling follows on the keyboard sink. The Enter
    // WebKit replays after a confirmed IME candidate is used up here.
    this.host?.addEventListener(
      'keydown',
      (ev) => {
        if (
          !ev.isComposing &&
          ev.keyCode !== 229 &&
          this.keyboardDelegate.takePostCompositionEnter(ev)
        ) {
          ev.preventDefault();
          ev.stopPropagation();
          return;
        }
        visualTexEditingCore?.offerKey(this.host, ev);
      },
      { capture: true, signal }
    );

    // Request notification for when the window is resized, the device
    // switched from portrait to landscape or the document is scrolled
    // to adjust the UI (popover, etc...)
    window.addEventListener('resize', this, { signal });
    document.addEventListener('scroll', this, { signal });
    this.resizeObserver = new ResizeObserver((_entries) => {
      if (this.resizeObserverStarted) {
        this.resizeObserverStarted = false;
        return;
      }
      requestUpdate(this);
    });
    this.resizeObserverStarted = true;
    this.resizeObserver.observe(this.field);
    this.resizeObserver.observe(this.container);
    // VisualTeX kernel: the editor restyles glyphs after a render (formula
    // letter font), which can widen the line after render() has already
    // scrolled the caret into view. Follow the settled width.
    this.caretScrollObserver = new ResizeObserver(() => {
      if (this.hasFocus()) keepCaretInScrollingAncestor(this);
    });
    this.caretScrollObserver.observe(this.field);

    if (gKeyboardLayout && !l10n.locale.startsWith(gKeyboardLayout.locale))
      setKeyboardLayoutLocale(l10n.locale);

    // When fonts are done loading, re-render
    // (the selection highlighting may be out of date due to the HTML layout
    // having been updated with the new font metrics)
    if (gFontsState !== 'ready')
      document.fonts.ready.then(() => renderSelection(this));

    // The mathfield container is initially set with a visibility of hidden
    // to minimize flashing during construction.
    element
      .querySelector<HTMLElement>('[part=container]')!
      .style.removeProperty('visibility');

    // Now start recording potentially undoable actions
    this.undoManager.startRecording();
    // Snapshot as 'set-value' operation, so that any other subsequent
    // `setValue()` gets coalesced
    this.undoManager.snapshot('set-value');

    // Request an update (this is necessary in the case where the mathfield
    // is empty but does have a contentPlaceholder)
    requestUpdate(this);
  }

  get colorMap(): (name: string) => string | undefined {
    return (name) => this.options.colorMap?.(name) ?? defaultColorMap(name);
  }

  get backgroundColorMap(): (name: string) => string | undefined {
    return (name) =>
      this.options.backgroundColorMap?.(name) ??
      this.options.colorMap?.(name) ??
      defaultBackgroundColorMap(name);
  }

  get smartFence(): boolean {
    return this.options.smartFence ?? false;
  }

  get readOnly(): boolean {
    return this.options.readOnly ?? false;
  }

  get disabled(): boolean {
    return (this.host as { disabled?: boolean } | undefined)?.['disabled'] ?? false;
  }

  // This reflects the contenteditable attribute.
  // Use hasEditableContent instead to take into account readonly and disabled
  // states.
  get contentEditable(): boolean {
    if (!this.host) return false;
    return this.host.getAttribute('contenteditable') !== 'false';
  }

  // This reflect the `user-select` CSS property
  get userSelect(): string {
    if (!this.host) return '';
    const style = getComputedStyle(this.host);
    // Safari uses '-webkit-user-select'. Other browsers use 'user-select'
    return (
      style.getPropertyValue('user-select') ||
      style.getPropertyValue('-webkit-user-select')
    );
  }

  get hasEditableContent(): boolean {
    if (this.disabled || !this.contentEditable) return false;
    return !this.readOnly || this.hasEditablePrompts;
  }

  get hasEditablePrompts(): boolean {
    return (
      this.readOnly &&
      !this.disabled &&
      this.contentEditable &&
      this.model.findAtom(
        (a) => a.type === 'prompt' && !(a as PromptAtom).locked
      ) !== undefined
    );
  }

  /** Returns true if the selection is editable:
   * - mathfield is not disabled, and has contentEditable
   * - if mathfield is readonly, the current selection is in a prompt which is editable (not locked)
   */
  get isSelectionEditable(): boolean {
    if (this.disabled || !this.contentEditable) return false;
    if (!this.readOnly) return true;

    const anchor = this.model.at(this.model.anchor);
    const cursor = this.model.at(this.model.position);

    const ancestor = Atom.commonAncestor(anchor, cursor);

    if (ancestor?.type === 'prompt' || ancestor?.parentPrompt) return true;

    return false;
  }

  get letterShapeStyle(): 'tex' | 'iso' | 'french' | 'upright' {
    return (
      (this.options.letterShapeStyle as 'tex' | 'iso' | 'french' | 'upright') ??
      'tex'
    );
  }

  get minFontScale(): number {
    return this.options.minFontScale;
  }

  get maxMatrixCols(): number {
    return this.options.maxMatrixCols;
  }

  /**
   *
   * If there is a selection, return if all the atoms in the selection,
   * some of them or none of them match the `style` argument.
   *
   * If there is no selection, return 'all' if the current implicit style
   * (determined by a combination of the style of the previous atom and
   * the current style) matches the `style` argument, 'none' if it does not.
   */
  queryStyle(inStyle: Readonly<Style>): 'some' | 'all' | 'none' {
    const style = validateStyle(this, inStyle);
    if ('verbatimColor' in style) delete style.verbatimColor;
    if ('verbatimBackgroundColor' in style)
      delete style.verbatimBackgroundColor;
    const keyCount = Object.keys(style).length;
    if (keyCount === 0) return 'all';
    if (keyCount > 1) {
      for (const prop of Object.keys(style)) {
        const result = this.queryStyle({ [prop]: (style as AnyStyle)[prop] });
        if (result === 'none') return 'none';
        if (result === 'some') return 'some';
      }
      return 'all';
    }

    const prop = Object.keys(style)[0] as keyof Style;
    const value = style[prop];

    if (this.model.selectionIsCollapsed) {
      const style = computeInsertStyle(this);
      return style[prop] === value ? 'all' : 'none';
    }

    const atoms = this.model.getAtoms(this.model.selection, {
      includeChildren: true,
    });
    let length = atoms.length;
    if (length === 0) return 'none';
    let count = 0;

    for (const atom of atoms) {
      if (atom.type === 'first') {
        length -= 1;
        continue;
      }
      if (atom.style[prop] === value) count += 1;
    }
    if (count === 0) return 'none';
    if (count === length) return 'all';
    return 'some';
  }

  get keybindings(): readonly Keybinding[] {
    if (this._keybindings) return this._keybindings;

    const [keybindings, errors] = normalizeKeybindings(
      this.options.keybindings,
      getActiveKeyboardLayout() ?? getDefaultKeyboardLayout()
    );

    if (getActiveKeyboardLayout()?.score > 0) {
      this._keybindings = keybindings;

      if (errors.length > 0) {
        console.error(
          `MathLive 0.109.2: Invalid keybindings for current keyboard layout`,
          errors
        );
      }
    }
    return keybindings;
  }

  setOptions(config: Partial<_MathfieldOptions>): void {
    this.options = { ...this.options, ...updateOptions(config) };

    this._keybindings = undefined;

    if (this.options.defaultMode === 'inline-math')
      this.element!.classList.add('ML__is-inline');
    else this.element!.classList.remove('ML__is-inline');

    // The mode of the 'first' atom is the mode of the  expression when empty
    let mode = this.options.defaultMode;
    if (mode === 'inline-math') mode = 'math';
    const visualTexRootFirstChild = this.model.root.firstChild;
    if (visualTexRootFirstChild && visualTexRootFirstChild.mode !== mode)
      visualTexRootFirstChild.mode = mode;

    // Changing some config options (i.e. `macros`) may
    // require the content to be reparsed and re-rendered
    const content = Atom.serialize([this.model.root], {
      expandMacro: false,
      defaultMode: this.options.defaultMode,
    });
    if ('macros' in config || this.model.getValue() !== content) reparse(this);

    if (
      'value' in config ||
      'registers' in config ||
      'colorMap' in config ||
      'backgroundColorMap' in config ||
      'letterShapeStyle' in config ||
      'minFontScale' in config ||
      'maxMatrixCols' in config ||
      'readOnly' in config ||
      'contentPlaceholder' in config ||
      'placeholderSymbol' in config
    )
      requestUpdate(this);
  }

  getOptions<K extends keyof _MathfieldOptions>(
    keys: K[]
  ): Pick<_MathfieldOptions, K>;
  getOptions(): _MathfieldOptions;
  getOptions(
    keys?: keyof _MathfieldOptions | (keyof _MathfieldOptions)[]
  ): any | Partial<_MathfieldOptions> {
    return getOptions(this.options, keys);
  }

  getOption<K extends keyof _MathfieldOptions>(key: K): _MathfieldOptions[K] {
    return getOptions(this.options, key) as _MathfieldOptions[K];
  }

  /*
   * handleEvent is a function invoked when an event is registered with an
   * object.
   * The name is defined by `addEventListener()` and cannot be changed.
   * This pattern is used to be able to release bound event handlers,
   * (event handlers that need access to `this`) as the `bind()` function
   * would create a new function that would have to be kept track of
   * to be able to properly remove the event handler later.
   */
  async handleEvent(evt: Event): Promise<void> {
    if (!isValidMathfield(this)) return;

    switch (evt.type) {
      case 'focus':
        // Skip handling DOM focus events that result from programmatic focus()
        // calls to prevent double onFocus() invocation (fixes #2685)
        if (!this.programmaticFocusInProgress)
          this.onFocus({ suppressEvents: true });
        break;

      case 'blur':
        this.onBlur({ dispatchEvents: false });
        break;

      // Safari on iOS <= 13 and Firefox on Android
      case 'mousedown':
        // pointerdown already handled this physical click
        if (typeof PointerEvent !== 'undefined') break;
        if (this.userSelect !== 'none') onPointerDown(this, evt as PointerEvent);

        break;

      case 'pointerdown':
        if (!evt.defaultPrevented && this.userSelect !== 'none')
          onPointerDown(this, evt as PointerEvent);
        break;

      case 'resize':
        if (this.geometryChangeTimer)
          cancelAnimationFrame(this.geometryChangeTimer);

        this.geometryChangeTimer = requestAnimationFrame(
          () => isValidMathfield(this) && this.onGeometryChange()
        );
        break;

      case 'scroll':
        if (this.geometryChangeTimer)
          cancelAnimationFrame(this.geometryChangeTimer);

        this.geometryChangeTimer = requestAnimationFrame(
          () => isValidMathfield(this) && this.onGeometryChange()
        );
        break;

      case 'wheel':
        this.onWheel(evt as WheelEvent);
        break;

      default:
        console.warn('Unexpected event type', evt.type);
    }
  }

  dispose(): void {
    if (!isValidMathfield(this)) return;

    this.keyboardDelegate.dispose();
    (this as any).keyboardDelegate = undefined;
    this.eventController.abort();
    (this as any).eventController = undefined;

    this.resizeObserver.disconnect();
    this.caretScrollObserver.disconnect();

    this.model.dispose();

    const element = this.element!;
    delete element.mathfield;
    (this.element as any) = undefined;

    (this as any).host = undefined;
    (this as any).field = undefined;
    (this as any).ariaLiveText = undefined;

    disposeSuggestionPopover(this);
  }

  flushInlineShortcutBuffer(options?: { defer: boolean }): void {
    options ??= { defer: false };
    if (!options.defer) {
      this.inlineShortcutBuffer.length = 0;
      clearTimeout(this.inlineShortcutBufferFlushTimer);
      this.inlineShortcutBufferFlushTimer = 0;
      return;
    }
    // If there is a timeout greater than 0, defer the reset
    // If the timeout is 0, never do the reset: regardless of the amount
    // of time between keystrokes, consider them as candidates for
    // a shortcut
    if (this.options.inlineShortcutTimeout > 0) {
      // Set a timer to reset the shortcut buffer
      clearTimeout(this.inlineShortcutBufferFlushTimer);
      this.inlineShortcutBufferFlushTimer = setTimeout(
        () => this.flushInlineShortcutBuffer(),
        this.options.inlineShortcutTimeout
      );
    }
  }

  executeCommand(
    command: SelectorPrivate | [SelectorPrivate, ...unknown[]]
  ): boolean {
    return perform(this, command);
  }

  get errors(): readonly LatexSyntaxError[] {
    return validateLatex(this.model.getValue(), { context: this.context });
  }

  getValue(): string;
  getValue(format: OutputFormat): string;
  getValue(start: Offset, end: Offset, format?: OutputFormat): string;
  getValue(range: Range, format?: OutputFormat): string;
  getValue(selection: Selection, format?: OutputFormat): string;
  getValue(
    arg1?: Offset | OutputFormat | Range | Selection,
    arg2?: Offset | OutputFormat,
    arg3?: OutputFormat
  ): string {
    return this.model.getValue(arg1, arg2, arg3);
  }

  setValue(value: string, options?: InsertOptions): void {
    // VisualTeX kernel: keystrokes typed before the formula was replaced must
    // not combine with later ones into a shortcut (hc, then hnu -> h ch nu).
    this.flushInlineShortcutBuffer();
    options = options ?? { mode: 'math' };
    if (options.insertionMode === undefined)
      options.insertionMode = 'replaceAll';

    if (options.format === undefined || options.format === 'auto')
      options.format = 'latex';

    if (options.mode === undefined || options.mode === 'auto')
      options.mode = getMode(this.model, this.model.position) ?? 'math';

    const couldUndo = this.undoManager.canUndo();
    if (ModeEditor.insert(this.model, value, options)) {
      // VisualTeX: replacing the whole formula with an explicitly requested
      // mode must also replace the live editing context. Upstream MathLive
      // otherwise parses as math but leaves model.mode untouched, so a field
      // that once entered text mode can stay text across reloads/syncs.
      if (
        options.insertionMode === 'replaceAll' &&
        options.mode &&
        (options.mode as string) !== 'auto'
      ) {
        this.model.mode = options.mode;
      }
      requestUpdate(this);
      // If this is the first insertion, drop the previous (empty) snapshot
      if (!couldUndo) this.undoManager.reset();
      this.undoManager.snapshot('set-value');
    }
  }

  /** Make sure the caret is visible within the matfield.
   * If using mathfield element, make sure the mathfield element is visible in
   * the page
   */
  scrollIntoView(): void {
    if (!this.element) return;
    //
    // 1/ If using a mathfield element, make sure that the element is visible.
    //

    if (this.host) {
      if (this.options.onScrollIntoView) this.options.onScrollIntoView(this);
      else {
        // 1.1/ Bring the mathfield into the viewport
        this.host.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
    }

    //
    // 2/ If a render is pending, do it now to make sure we have correct layout
    // and caret position
    //
    if (this.dirty) render(this, { interactive: true });

    // VisualTeX: the field and its host grow to the formula and never scroll;
    // the caret is kept in view in the formula's row (and the keyboard sink
    // follows it) by the render follow-up.
    scheduleCaretFollowUp(this);
  }

  insert(s: string, options?: InsertOptions): boolean {
    if (typeof s !== 'string') return false;

    if (
      s.length === 0 &&
      (options?.insertionMode === 'insertBefore' ||
        options?.insertionMode === 'insertAfter')
    )
      return false;

    if (s.length === 0 && this.model.selectionIsCollapsed) return false;

    // This code path is used when inserting content from the virtual keyboard
    // (i.e. inserting `\sin`). We need to ignore previous key combinations
    // in this case
    this.flushInlineShortcutBuffer();

    options = options ?? { mode: 'math' };
    if (options.focus) this.focus();

    if (s === '\\\\') {
      // This string is interpreted as an "insert row after" command
      addRowAfter(this.model);
    } else if (s === '&') addColumnAfter(this.model);
    else {
      if (this.model.selectionIsCollapsed) {
        const style = { ...computeInsertStyle(this) };
        // If we're inserting a non-alphanumeric character, reset the variant
        if (!/^[a-zA-Z0-9]$/.test(s) && this.styleBias !== 'none') {
          style.variant = 'normal';
          style.variantStyle = undefined;
        }

        ModeEditor.insert(this.model, s, { style, ...options });
      } else ModeEditor.insert(this.model, s, options);
    }

    this.snapshot(`insert-${this.model.at(this.model.position).type}`);

    requestUpdate(this);
    if (options.scrollIntoView) this.scrollIntoView();

    return true;
  }

  /**
   * Switch from the current mode to the new mode, if different.
   * Prefix and suffix are optional strings to be inserted before and after
   * the mode change, so prefix is interpreted with the current mode and
   * suffix with the new mode.
   */
  switchMode(mode: ParseMode, prefix = '', suffix = ''): void {
    if (this.model.mode === mode || !this.hasEditableContent || !this.contentEditable || this.disabled)
      return;
    const { model } = this;
    // VisualTeX: a physical command prefix must never reinterpret an
    // accidental residual selection left by IME composition. Preserve the
    // selected content, collapse to its forward edge, then start raw LaTeX.
    // Structural placeholders remain selected because typing a command there
    // intentionally replaces the placeholder.
    if (
      mode === "latex" &&
      suffix === "\\" &&
      !model.selectionIsCollapsed &&
      !model.selectionIsPlaceholder
    ) {
      model.collapseSelection("forward");
    }
    if (mode === "latex" && model.selectionIsPlaceholder) {
      const selectedPlaceholder = model.at(Math.max(model.position, model.anchor));
      if (selectedPlaceholder?.type === "placeholder") {
        this.visualTexPendingStructuralCommandSlot = visualTexFreshStructuralSlot(selectedPlaceholder);
      }
    }
    if (mode === "latex" && model.selectionIsCollapsed) {
      const at = model.at(model.position);
      const placeholder = at?.type === "placeholder" ? at : at?.rightSibling?.type === "placeholder" ? at.rightSibling : null;
      if (placeholder) {
        this.visualTexPendingStructuralCommandSlot = visualTexFreshStructuralSlot(placeholder);
        model.setSelection(model.offsetOf(placeholder.leftSibling), model.offsetOf(placeholder));
      } else if (visualTexLiveStructuralSlot(at)) {
        this.visualTexPendingStructuralCommandSlot = visualTexLiveStructuralSlot(at);
      }
    }
    const previousMode = model.mode;
    model.mode = mode;
    if (this.host && !this.host.dispatchEvent(
      new Event("mode-change", {
        bubbles: true,
        composed: true,
        cancelable: true
      })
    )) {
      model.mode = previousMode;
      return;
    }
    model.mode = previousMode;
    model.deferNotifications(
      {
        content: Boolean(suffix) || Boolean(prefix),
        selection: true,
        // Boolean(suffix) || Boolean(prefix),
        type: "insertText"
      },
      () => {
        let cursor = model.at(model.position);
        const insertString = (s: string, options: { select: boolean }) => {
          if (!s) return;
          const atoms = model.mode === "math" ? parseLatex(parseMathString(s, { format: "ascii-math" })[1], {
            context: this.context
          }) : [...s].map((c) => new TextAtom(c, c, {}));
          if (options.select) {
            const end = cursor.parent!.addChildrenAfter(atoms, cursor);
            model.setSelection(
              model.offsetOf(atoms[0].leftSibling),
              model.offsetOf(end)
            );
          } else {
            model.position = model.offsetOf(
              cursor.parent!.addChildrenAfter(atoms, cursor)
            );
          }
          contentChanged = true;
        };
        const insertLatexGroup = (latex: string, options: { select: boolean }) => {
          const atom = new LatexGroupAtom(latex);
          cursor.parent!.addChildAfter(atom, cursor);
          if (options.select) {
            model.setSelection(
              model.offsetOf(atom.firstChild),
              model.offsetOf(atom.lastChild)
            );
          } else model.position = model.offsetOf(atom.lastChild);
          contentChanged = true;
        };
        let visualTexReplacingPlaceholder = false;
        const getContent = () => {
          const format = mode === "latex" ? "latex" : mode === "math" ? "plain-text" : "ascii-math";
          const selRange = range(model.selection);
          let content = this.model.getValue(selRange, format);
          const atoms = this.model.extractAtoms(selRange);
          if (atoms.length === 1 && atoms[0].type === "placeholder") {
            visualTexReplacingPlaceholder = true;
            content = suffix;
          }
          cursor = model.at(selRange[0]);
          return content;
        };
        let contentChanged = false;
        this.flushInlineShortcutBuffer();
        this.stopCoalescingUndo();
        complete(this, "accept");
        if (model.selectionIsCollapsed) {
          insertString(prefix, { select: false });
          model.mode = mode;
          if (mode === "latex") insertLatexGroup(suffix, { select: false });
          else insertString(suffix, { select: false });
        } else {
          const content = getContent();
          model.mode = mode;
          if (mode === "latex") insertLatexGroup(content, { select: !visualTexReplacingPlaceholder });
          else insertString(content, { select: true });
        }
        requestUpdate(this);
        this.undoManager.snapshot(mode === "latex" ? "insert-latex" : "insert");
        return contentChanged;
      }
    );
    model.mode = mode;
  }

  hasFocus(): boolean {
    return !this.blurred;
  }

  focus(options?: FocusOptions): void {
    if (this.disabled || this.focusBlurInProgress) return;
    if (!this.hasFocus()) {
      this.programmaticFocusInProgress = true;
      this.onFocus();
      this.model.announce('line');
    }
    if (!(options?.preventScroll ?? false)) this.scrollIntoView();
  }

  blur(): void {
    if (!this.hasFocus()) return;
    this.keyboardDelegate.blur();
  }

  select(): void {
    this.model.selection = { ranges: [[0, this.model.lastOffset]] };
    // The behavior of select() is not clearly defined. Some implementations
    // focus, others don't. Selecting the sink may focus it in some cases
    // so, to be safe, we focus the field as well
    this.focus();
  }

  applyStyle(inStyle: Style, inOptions: Range | ApplyStyleOptions = {}): void {
    let range: Range | undefined;
    let operation: 'set' | 'toggle' = 'set';
    let silenceNotifications = false;

    if (isRange(inOptions)) range = inOptions;
    else {
      if (inOptions.operation === 'toggle') operation = 'toggle';
      range = inOptions.range;
      silenceNotifications = inOptions.silenceNotifications ?? false;
    }

    if (range) range = this.model.normalizeRange(range);
    if (range && range[0] === range[1]) range = undefined;

    const style = validateStyle(this, inStyle);

    if (range === undefined && this.model.selectionIsCollapsed) {
      // We don't have a selection. Set the global style instead.
      if (operation === 'set') {
        const newStyle: PrivateStyle = { ...this.defaultStyle };
        if ('color' in style) delete newStyle.verbatimColor;
        if ('backgroundColor' in style) delete newStyle.verbatimBackgroundColor;
        this.defaultStyle = { ...newStyle, ...style };
        this.styleBias = 'none';
        return;
      }

      // Toggle the properties
      const currentStyle = computeInsertStyle(this);
      const newStyle: PrivateStyle = { ...this.defaultStyle };
      for (const prop of Object.keys(style)) {
        if ((currentStyle as AnyStyle)[prop] === (style as AnyStyle)[prop]) {
          if (prop === 'color') delete newStyle.verbatimColor;
          if (prop === 'backgroundColor')
            delete newStyle.verbatimBackgroundColor;

          delete (newStyle as AnyStyle)[prop];
        } else (newStyle as AnyStyle)[prop] = (style as AnyStyle)[prop];
      }
      this.defaultStyle = newStyle;
      this.styleBias = 'none';
      return;
    }

    this.model.deferNotifications(
      { content: !silenceNotifications, type: 'insertText' },
      () => {
        if (range === undefined) {
          for (const range of this.model.selection.ranges)
            applyStyle(this.model, range, style, { operation });
        } else applyStyle(this.model, range, style, { operation });
      }
    );
    requestUpdate(this);
  }

  getPrompt(id: string): PromptAtom | undefined {
    const prompt = this.model.findAtom(
      (a) => a.type === 'prompt' && (a as PromptAtom).placeholderId === id
    ) as PromptAtom | undefined;
    console.assert(
      prompt !== undefined,
      `MathLive 0.109.2:  no prompts with matching ID found`
    );
    return prompt;
  }

  getPromptValue(id: string, format?: OutputFormat): string {
    const prompt = this.getPrompt(id);
    if (!prompt) return '';

    const first = this.model.offsetOf(prompt.firstChild);
    const last = this.model.offsetOf(prompt.lastChild);

    return this.model.getValue(first, last, format);
  }

  getPrompts(filter?: {
    id?: string;
    locked?: boolean;
    correctness?: 'correct' | 'incorrect' | 'undefined';
  }): string[] {
    return (this.model.atoms as PromptAtom[])
      .filter((a) => {
        if (a.type !== 'prompt') return false;
        if (!filter) return true;

        if (filter.id && a.placeholderId !== filter.id) return false;
        if (filter.locked && a.locked !== filter.locked) return false;
        if (filter.correctness === 'undefined' && a.correctness) return false;
        if (filter.correctness && a.correctness !== filter.correctness)
          return false;

        return true;
      })
      .map((a) => a.placeholderId!);
  }

  setPromptValue(
    id: string,
    value?: string,
    insertOptions?: Omit<InsertOptions, 'insertionMode'>
  ): void {
    if (value !== undefined) {
      const prompt = this.getPrompt(id);
      if (!prompt) {
        console.error(`MathLive 0.109.2: unknown prompt ${id}`);
        return;
      }

      const branchRange = this.model.getBranchRange(
        this.model.offsetOf(prompt),
        'body'
      );

      this.model.setSelection(branchRange);
      this.insert(value, {
        ...insertOptions,
        insertionMode: 'replaceSelection',
      });
    }
    if (insertOptions?.silenceNotifications)
      this.valueOnFocus = this.getValue();
    requestUpdate(this);
  }

  setPromptState(
    id: string,
    state: 'correct' | 'incorrect' | 'undefined' | undefined,
    locked?: boolean
  ): void {
    const prompt = this.getPrompt(id);
    if (!prompt) {
      console.error(`MathLive 0.109.2: unknown prompt ${id}`);
      return;
    }
    if (state === 'undefined') prompt.correctness = undefined;
    else if (typeof state === 'string') prompt.correctness = state;

    if (typeof locked === 'boolean') {
      prompt.locked = locked;
      prompt.captureSelection = locked;
    }

    requestUpdate(this);
  }

  getPromptState(id: string): ['correct' | 'incorrect' | undefined, boolean] {
    const prompt = this.getPrompt(id);
    if (!prompt) {
      console.error(`MathLive 0.109.2: unknown prompt ${id}`);
      return [undefined, true];
    }
    return [prompt.correctness, prompt.locked];
  }

  getPromptRange(id: string): Range {
    const prompt = this.getPrompt(id);
    if (!prompt) {
      console.error(`MathLive 0.109.2: unknown prompt ${id}`);
      return [0, 0];
    }
    return this.model.getBranchRange(this.model.offsetOf(prompt), 'body');
  }

  canUndo(): boolean {
    return this.undoManager.canUndo();
  }

  canRedo(): boolean {
    return this.undoManager.canRedo();
  }

  popUndoStack(): void {
    this.undoManager.pop();
  }

  snapshot(op?: string): void {
    if (this.undoManager.snapshot(op)) {
      this.host?.dispatchEvent(
        new CustomEvent('undo-state-change', {
          bubbles: true,
          composed: true,
          detail: { type: 'snapshot' },
        })
      );
    }
  }

  stopCoalescingUndo(): void {
    this.undoManager.stopCoalescing(this.model.selection);
  }

  stopRecording(): void {
    this.undoManager.stopRecording();
  }

  startRecording(): void {
    this.undoManager.startRecording();
  }

  undo(): void {
    if (!this.undoManager.undo()) return;
    this.host?.dispatchEvent(
      new CustomEvent('undo-state-change', {
        bubbles: true,
        composed: true,
        detail: { type: 'undo' },
      })
    );
  }

  redo(): void {
    if (!this.undoManager.redo()) return;
    this.host?.dispatchEvent(
      new CustomEvent('undo-state-change', {
        bubbles: true,
        composed: true,
        detail: { type: 'redo' },
      })
    );
  }

  resetUndo(): void {
    this.undoManager?.reset();
  }

  onSelectionDidChange(): void {
    const model = this.model;
    // VisualTeX: a source wrapper the caret left is finished where it is
    visualTexFinalizeSourceWrapperOnCaretExit(this);

    // Keep the content of the keyboard sink in sync with the selection.
    // Safari will not dispatch cut/copy/paste unless there is a DOM selection.
    // No need to sync if the mathfield is not focused: syncing has a side-effect
    // of focusing the mathfield if not focused, which is not desirable
    if (model.mathfield.hasFocus()) {
      this.keyboardDelegate.setValue(
        model.getValue(model.selection, 'latex-expanded')
      );
    }

    // If we move the selection outside of a LaTeX group, close the group
    if (model.selectionIsCollapsed) {
      const latexGroup = getLatexGroup(model);
      const pos = model.position;
      const cursor = model.at(pos);
      const sibling = model.at(pos + 1);
      // VisualTeX editing core (text.ts): text only inside a text run
      const mode = effectiveMode(this.options) === 'math'
        ? visualTexEditingCore.caretMode(model, pos)
        : cursor.mode ?? effectiveMode(this.options);
      if (
        latexGroup &&
        (pos < model.offsetOf(latexGroup.firstChild) ||
          pos > model.offsetOf(latexGroup.lastChild))
      ) {
        // We moved outside a LaTeX group
        complete(this, 'accept', { mode });
        const cursorOffset = model.offsetOf(cursor);
        if (cursorOffset >= 0) model.position = cursorOffset;
      } else {
        // If we're at the start or the end of a LaTeX group,
        // move inside the group and don't switch mode.
        if (sibling?.type === 'first' && sibling.mode === 'latex')
          model.position = pos + 1;
        else if (latexGroup && sibling?.mode !== 'latex')
          model.position = pos - 1;
        else {
          // We may have moved from math to text, or text to math.
          this.switchMode(mode);
        }
      }
    }

    // Dispatch `selection-change` event
    this.host?.dispatchEvent(
      new Event('selection-change', {
        bubbles: true,
        composed: true,
      })
    );
  }

  onContentWillChange(options: ContentChangeOptions): boolean {
    return (
      this.host?.dispatchEvent(
        new InputEvent('beforeinput', {
          ...options,
          // To work around a bug in WebKit/Safari (the inputType property gets stripped), include the inputType as the 'data' property. (see #1843)
          data: options.data ? options.data : (options.inputType ?? ''),
          cancelable: true,
          bubbles: true,
          composed: true,
        })
      ) ?? true
    );
  }

  onFocus(options?: { suppressEvents?: boolean }): void {
    if (this.disabled || this.focusBlurInProgress || !this.blurred) return;

    // If another mathfield is globally tracked as focused, blur it first.
    // This handles cases where browsers don't fire blur events reliably
    // (e.g., rapid focus() calls on multiple mathfields in Chromium).
    const previouslyFocusedMathfield = _Mathfield._globallyFocusedMathfield;
    if (
      previouslyFocusedMathfield &&
      previouslyFocusedMathfield !== this &&
      !previouslyFocusedMathfield.disabled &&
      previouslyFocusedMathfield.hasFocus()
    ) {
      // Call onBlur directly to avoid DOM event timing issues
      previouslyFocusedMathfield.onBlur({ dispatchEvents: true });
    }

    this.focusBlurInProgress = true;
    this.blurred = false;

    // Update the global tracker to point to this mathfield
    _Mathfield._globallyFocusedMathfield = this;

    this.stopCoalescingUndo();

    // Save the current value.
    // It will be compared in `onBlur()` to see if the
    // `change` event needs to be dispatched. This
    // mimic the `<input>` and `<textarea>` behavior
    this.valueOnFocus = this.model.getValue();

    // If we're in prompt mode, and the selection is
    // not in a prompt, move it to a prompt
    if (
      this.hasEditablePrompts &&
      !this.model.at(this.model.anchor).parentPrompt
    )
      this.executeCommand('moveToNextPlaceholder');

    render(this, { interactive: true });

    setTimeout(() => {
      if (!isValidMathfield(this)) return;

      // Only suppress events when responding to a DOM focus event to avoid
      // double-dispatching (fixes #2665). When focus() is called
      // programmatically, we want the events to fire normally (fixes #2816).
      const suppressEvents = options?.suppressEvents ?? false;
      if (suppressEvents) {
        //
        // Capture the focus/blur events to avoid double-dispatching
        // When responding to a DOM focus event, just focus the keyboard
        // delegate without the blur/focus cycle
        //
        const abortController = new AbortController();
        const signal = abortController.signal;
        const captureEvent = (evt: Event) => {
          evt.preventDefault();
          evt.stopPropagation();
        };

        // Add listeners that will capture focus/focusin events
        for (const event of ['focus', 'focusin']) {
          this.host?.addEventListener(event, captureEvent, {
            capture: true,
            signal,
          });
        }

        this.focusBlurInProgress = false;
        this.keyboardDelegate.focus();

        // Abort the event listeners after a short delay to ensure all
        // events from the focus have been captured
        setTimeout(() => abortController.abort(), 0);
      } else {
        // Programmatic focus() call - just focus the keyboard delegate
        // without blurring first (no need since we're not focused yet)
        this.focusBlurInProgress = false;
        this.keyboardDelegate.focus();

        // Clear the flag after the DOM focus event has been processed
        setTimeout(() => {
          this.programmaticFocusInProgress = false;
        }, 0);
      }
    }, 60);
  }

  onBlur(options?: { dispatchEvents?: boolean }): void {
    if (this.focusBlurInProgress || this.blurred) return;
    // VisualTeX kernel: leaving the field finishes the word at the caret.
    visualTexJudgeInput(this, { finishing: true });
    this.focusBlurInProgress = true;

    this.stopCoalescingUndo();
    // VisualTeX: leaving the field finishes a source wrapper (not during a
    // composition inside it: its end finishes the wrapper)
    if (!this.visualTexCompositionSourceWrapper)
      visualTexFinalizeSourceWrapper(this, { notify: false });

    this.blurred = true;

    // Clear the global tracker if it points to this mathfield
    if (_Mathfield._globallyFocusedMathfield === this)
      _Mathfield._globallyFocusedMathfield = undefined;

    if (this.ariaLiveText) this.ariaLiveText.textContent = '';

    hideSuggestionPopover(this);

    if (this.model.getValue() !== this.valueOnFocus) {
      this.host?.dispatchEvent(
        new Event('change', { bubbles: true, composed: true })
      );
    }

    // When called from keyboard delegate or programmatically, dispatch events.
    // When responding to a DOM blur event, don't dispatch again to avoid
    // double-firing (fixes #2816).
    const dispatchEvents = options?.dispatchEvents ?? true;
    if (dispatchEvents) {
      this.host?.dispatchEvent(
        new Event('blur', {
          bubbles: false, // DOM 'focus' and 'blur' don't bubble
          composed: true,
        })
      );

      this.host?.dispatchEvent(
        new UIEvent('focusout', {
          bubbles: true, // unlike 'blur', focusout does bubble
          composed: true,
        })
      );
    }

    requestUpdate(this);

    this.focusBlurInProgress = false;

    if (MathfieldElement.restoreFocusWhenDocumentFocused) {
      //
      // When the document/window loses focus, for example by switching
      // to another tab, the mathfield will be blured. When the window
      // regains focus, we'd like the focus to be restored on the mathfield,
      // like the browsers do for `<textarea>` elements. However, they
      // don't do that for custom elements, so we do it ourselves. @futureweb
      //

      // Wait for the window focus to change
      // (the mathfield gets blurred before the window)
      const controller = new AbortController();
      const signal = controller.signal;
      window.addEventListener(
        'blur',
        () => {
          window.addEventListener(
            'focus',
            () => {
              if (isValidMathfield(this)) this.focus({ preventScroll: true });
            },
            { once: true, signal }
          );
        },
        { once: true, signal }
      );

      // If something else gets the focus (could be the mathfield too),
      // cancel the above
      document.addEventListener('focusin', () => controller.abort(), {
        once: true,
      });

      // If user clicks anywhere, cancel the refocus (covers case where
      // click doesn't cause an element to come into focus)
      document.addEventListener('click', () => controller.abort(), {
        once: true,
      });
    }
  }

  onInput(text: string): void {
    onInput(this, text);
  }

  onKeystroke(evt: KeyboardEvent): boolean {
    visualTexAnchorKeyboardSink(this);
    return onKeystroke(this, evt);
  }

  onCompositionStart(_composition: string): void {
    if (!this.isSelectionEditable) return;
    visualTexAnchorKeyboardSink(this);
    // A caret at the visual boundary of \text{} is a math caret in VisualTeX.
    // Normalize before the composition atom is created so IME cannot inherit
    // "text" merely because the atom to the left is a TextAtom.
    visualTexNormalizeTextBoundaryMode(this);
    this.visualTexCompositionStartedInsideText =
      this.model.mode === "text" && visualTexCaretIsInsideTextRun(this.model);
    const visualTexSourceWrapper = this.visualTexSourceWrapper;
    this.visualTexCompositionSourceWrapper =
      visualTexSourceWrapper && visualTexSourceWrapperOwnsAtom(visualTexSourceWrapper, this.model.at(this.model.position))
        ? visualTexSourceWrapper
        : null;
    this.visualTexCompositionSourceAfter = this.visualTexCompositionSourceWrapper
      ? this.model.at(this.model.position)
      : null;
    this.visualTexCompositionState = this.model.getState();
    const visualTexCompositionPlaceholder = this.model.selectionIsPlaceholder
      ? this.model.at(Math.max(this.model.position, this.model.anchor))
      : null;
    this.visualTexCompositionStructuralSlot =
      visualTexCompositionPlaceholder?.type === "placeholder"
        ? visualTexFreshStructuralSlot(visualTexCompositionPlaceholder)
        : (() => {
            const current = this.model.at(this.model.position);
            if (!current?.parent || current.parentBranch == null) return null;
            return {
              parent: current.parent,
              branch: current.parentBranch,
              after: current.leftSibling
            };
          })();
    this.stopCoalescingUndo();
    // Composition replaces the selection without invoking structural
    // Backspace promotion rules.
    if (!this.model.selectionIsCollapsed)
      this.model.deleteAtoms(range(this.model.selection));
  }

  onCompositionUpdate(composition: string): void {
    updateComposition(this.model, composition);
    requestUpdate(this);
    visualTexAnchorKeyboardSink(this);
  }

  onCompositionEnd(composition: string): void {
    removeComposition(this.model);
    if (!composition && this.visualTexCompositionState) {
      this.model.setState(this.visualTexCompositionState);
      visualTexNormalizeTextBoundaryMode(this);
      requestUpdate(this);
    } else {
      this.visualTexCompositionCommit = true;
      try {
        const visualTexChineseComposition =
          /^[\u3400-\u9fff\uf900-\ufaff，。；：！？、（）【】《》“”‘’\s]+$/.test(composition);
        // the Space that picked a candidate is not part of the text
        const visualTexCommittedComposition =
          visualTexChineseComposition && /[^\s]/.test(composition)
            ? composition.replace(/\s+$/u, "")
            : composition;
        const visualTexSourceWrapper = this.visualTexCompositionSourceWrapper;
        const visualTexSourceAfter = this.visualTexCompositionSourceAfter;
        if (
          visualTexSourceWrapper &&
          this.visualTexSourceWrapper === visualTexSourceWrapper &&
          visualTexSourceWrapperOwnsAtom(visualTexSourceWrapper, visualTexSourceAfter)
        ) {
          // typed inside a source wrapper: the text joins its source
          const visualTexSourceComposition = visualTexChineseComposition
            ? "\\text{" + visualTexCommittedComposition + "}"
            : visualTexCommittedComposition;
          this.model.position = this.model.offsetOf(visualTexSourceAfter);
          ModeEditor.insert(this.model, visualTexSourceComposition, {
            mode: "latex",
            insertionMode: "replaceSelection",
            selectionMode: "after"
          });
          this.model.mode = "latex";
          visualTexSyncSourceWrapper(this);
          this.snapshot("insert-composition");
          requestUpdate(this);
        } else if (visualTexChineseComposition) {
          if (this.visualTexCompositionStructuralSlot)
            this.visualTexPendingStructuralCommandSlot =
              this.visualTexCompositionStructuralSlot;
          if (this.visualTexCompositionStartedInsideText) {
            ModeEditor.insert(this.model, visualTexCommittedComposition, {
              mode: "text",
              insertionMode: "replaceSelection",
              selectionMode: "after",
              style: computeInsertStyle(this)
            });
            this.model.mode = "text";
          } else {
            ModeEditor.insert(this.model, "\\text{" + visualTexCommittedComposition + "}", {
              mode: "math",
              format: "latex",
              insertionMode: "replaceSelection",
              selectionMode: "after",
              style: computeInsertStyle(this)
            });
            if (!this.model.selectionIsCollapsed)
              this.model.collapseSelection("forward");
            this.model.mode = "math";
          }
          this.snapshot("insert-composition");
          requestUpdate(this);
        } else {
          onInput(this, composition, { simulateKeystroke: true });
          if (!this.visualTexCompositionStartedInsideText)
            visualTexNormalizeTextBoundaryMode(this, true);
        }
      } finally {
        this.visualTexCompositionCommit = false;
      }
      const visualTexCompositionAtom = this.model.at(this.model.position);
      if (!visualTexAutoExitAccentAfterDirectInput(
            this,
            visualTexCompositionAtom
          ) &&
          !visualTexAutoExitScriptAfterDirectInput(
            this,
            visualTexCompositionAtom
          ))
        visualTexAutoExitAfterComposition(
          this,
          this.visualTexCompositionStructuralSlot
        );
    }
    this.visualTexCompositionState = null;
    this.visualTexCompositionStructuralSlot = null;
    this.visualTexCompositionSourceWrapper = null;
    this.visualTexCompositionSourceAfter = null;
    this.visualTexCompositionStartedInsideText = false;
    if (this.blurred)
      visualTexFinalizeSourceWrapper(this, { notify: false });
  }

  onCut(ev: ClipboardEvent): void {
    // Ignore if in read-only mode
    if (!this.isSelectionEditable) {
      this.model.announce('plonk');
      return;
    }

    if (this.model.contentWillChange({ inputType: 'deleteByCut' })) {
      // Snapshot the undo state
      this.stopCoalescingUndo();

      // Copy to the clipboard
      if (ev.clipboardData) ModeEditor.onCopy(this, ev);
      else ModeEditor.copyToClipboard(this, 'latex');

      // Delete the selection
      deleteRange(this.model, range(this.model.selection), 'deleteByCut');

      this.snapshot('cut');

      requestUpdate(this);
    }
  }

  onCopy(ev: ClipboardEvent): void {
    if (ev.clipboardData) ModeEditor.onCopy(this, ev);
    else ModeEditor.copyToClipboard(this, 'latex');
  }

  onPaste(ev: ClipboardEvent): boolean {
    // Ignore if in read-only mode
    let result = this.isSelectionEditable;

    if (result) {
      // VisualTeX editing core (text.ts): the caret's mode, not its left atom's
      result = ModeEditor.onPaste(
        visualTexEditingCore.caretMode(this.model, this.model.position),
        this,
        ev.clipboardData
      );
    }

    if (!result) this.model.announce('plonk');

    ev.preventDefault();
    ev.stopPropagation();
    return result;
  }

  private onGeometryChange(): void {
    updateSuggestionPopoverPosition(this);
  }

  private onWheel(ev: WheelEvent): void {
    const wheelDelta = 5 * ev.deltaX;
    if (!Number.isFinite(wheelDelta) || wheelDelta === 0) return;

    const field = this.field!;

    if (wheelDelta < 0 && field.scrollLeft === 0) return;

    if (
      wheelDelta > 0 &&
      field.offsetWidth + field.scrollLeft >= field.scrollWidth
    )
      return;

    field.scrollBy({ top: 0, left: wheelDelta });
    ev.preventDefault();
    ev.stopPropagation();
  }

  getHTMLElement(atom: Atom): HTMLSpanElement {
    // Find an atom id in this atom or its children
    let target = atom;
    while (!target.id && target.hasChildren) target = atom.children[0];

    return this.field.querySelector(
      `[data-atom-id="${target.id}"]`
    ) as HTMLSpanElement;
  }

  get context(): Readonly<ContextInterface> {
    return {
      registers: this.options.registers ?? {},
      smartFence: this.smartFence,
      letterShapeStyle: this.letterShapeStyle,
      minFontScale: this.minFontScale,
      maxMatrixCols: this.maxMatrixCols,
      placeholderSymbol: this.options.placeholderSymbol ?? '▢',
      colorMap: (name) => this.colorMap(name),
      backgroundColorMap: (name) => this.backgroundColorMap(name),
      getMacro: (token) =>
        getMacroDefinition(
          token,
          this.options.macros as NormalizedMacroDictionary
        ),
      atomIdsSettings: { seed: 'random', groupNumbers: false },
    };
  }
}
