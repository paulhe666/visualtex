import type { ClipboardFragment } from './protocol';
import type { StructuredClipboardFragment } from './editing';

export const FRAGMENT_MIME = 'application/vnd.visualtex.engine-fragment+json';
export const LATEX_MIME = 'application/x-latex';
const MAX_CLIPBOARD = 1_048_576;
export interface InputPort {
  key(event: KeyboardEvent): boolean;
  insert(text: string): boolean;
  remove(backward: boolean): boolean;
  history(redo: boolean): boolean;
  lineBreak?(): boolean;
  paragraph?(): boolean;
  beginComposition(): number | null;
  updateComposition(id: number, text: string): boolean;
  commitComposition(id: number, text: string): boolean;
  cancelComposition(id: number): boolean;
  copy(): ClipboardFragment | null;
  cut(fragment: ClipboardFragment): boolean;
  paste(text: string, latex: string | null, structure: StructuredClipboardFragment | null): boolean;
  focus(focused: boolean): void;
  error(error: unknown): void;
}
export type PendingInput =
  | { kind: 'composition'; id: number; text: string; cancel: boolean }
  | { kind: 'input'; inputType: string; text: string };
export interface InputAudit {
  type: string;
  key: string | null;
  inputType: string | null;
  data: string | null;
  trusted: boolean;
  composing: boolean | null;
  compositionId: number | null;
  value: string;
}

/** A real textarea owns browser/OS input and candidate positioning. It does not
 * own document text, selections, math parsing, undo, or formula geometry. */
export class BrowserInputBridge {
  readonly textarea: HTMLTextAreaElement;
  readonly preedit: HTMLSpanElement;
  private readonly abort = new AbortController();
  private compositionId: number | null = null;
  private disposed = false;
  private terminalText: string | null = null;
  private flushTimer: ReturnType<typeof setTimeout> | null = null;
  private readonly audit: InputAudit[] = [];
  private pending: PendingInput | null = null;
  private latestPreedit = '';
  private ignoredNativeComposition = false;

  constructor(parent: HTMLElement, private readonly port: InputPort) {
    const doc = parent.ownerDocument;
    this.textarea = doc.createElement('textarea');
    this.textarea.className = 'vt-core-input';
    this.textarea.setAttribute('aria-label', '公式输入');
    this.textarea.setAttribute('aria-multiline', 'false');
    this.textarea.setAttribute('autocapitalize', 'off');
    this.textarea.setAttribute('autocorrect', 'off');
    this.textarea.autocomplete = 'off';
    this.textarea.spellcheck = false;
    this.textarea.wrap = 'off';
    this.textarea.rows = 1;
    this.textarea.inputMode = 'text';
    // It remains focusable and in the accessibility tree. display:none and
    // detached inputs prevent native IME/candidate windows from working.
    Object.assign(this.textarea.style, {
      position: 'fixed', width: '2px', height: '24px', margin: '0', padding: '0',
      border: '0', outline: '0', resize: 'none', overflow: 'hidden',
      color: 'transparent', caretColor: 'transparent', background: 'transparent',
      opacity: '0.01', pointerEvents: 'none', zIndex: '4', fontSize: '20px', lineHeight: '1',
    });
    this.preedit = doc.createElement('span');
    this.preedit.className = 'vt-core-preedit';
    this.preedit.hidden = true;
    this.preedit.setAttribute('aria-hidden', 'true');
    Object.assign(this.preedit.style, { position: 'fixed', pointerEvents: 'none', zIndex: '5' });
    parent.append(this.textarea, this.preedit);
    const signal = this.abort.signal;
    this.textarea.addEventListener('keydown', event => {
      this.record(event);
      if (this.pending !== null) {
        if (event.isComposing || event.keyCode === 229 || this.ignoredNativeComposition) return;
        event.preventDefault();
        if (event.key === 'Enter') this.retryPendingInput();
        else if (event.key === 'Escape') this.discardPendingInput();
        return;
      }
      if (event.isComposing || event.keyCode === 229 || this.compositionId !== null) return;
      this.clearTerminal();
      if (this.attempt(() => port.key(event), false)) event.preventDefault();
    }, { signal });
    this.textarea.addEventListener('beforeinput', event => this.beforeInput(event as InputEvent), { signal });
    this.textarea.addEventListener('input', event => this.afterInput(event as InputEvent), { signal });
    this.textarea.addEventListener('compositionstart', event => {
      this.record(event);
      if (this.pending !== null) { this.ignoredNativeComposition = true; return; }
      this.clearTerminal();
      if (this.compositionId !== null && !this.attempt(() => port.cancelComposition(this.compositionId!), false)) {
        this.hold({kind:'composition',id:this.compositionId,text:this.latestPreedit,cancel:true});
        this.ignoredNativeComposition = true;
        return;
      }
      this.compositionId = this.attempt(() => port.beginComposition(), null);
      this.latestPreedit = '';
    }, { signal });
    this.textarea.addEventListener('compositionupdate', event => {
      this.record(event);
      if (this.pending !== null || this.ignoredNativeComposition) return;
      this.latestPreedit = event.data;
      if (this.compositionId !== null) this.attempt(() => port.updateComposition(this.compositionId!, event.data), false);
    }, { signal });
    this.textarea.addEventListener('compositionend', event => {
      this.record(event);
      if (this.pending !== null || this.ignoredNativeComposition) {
        this.ignoredNativeComposition = false;
        if (this.pending) this.textarea.value = this.pending.text;
        return;
      }
      const id = this.compositionId;
      if (id !== null) {
        const accepted = this.attempt(() => event.data === '' ? port.cancelComposition(id) : port.commitComposition(id, event.data), false);
        if (!accepted) {
          this.hold({kind:'composition',id,text:event.data || this.latestPreedit,cancel:event.data===''});
          return;
        }
      }
      this.compositionId = null;
      this.latestPreedit = '';
      // Engines differ in whether the terminal input event precedes or follows
      // compositionend. Keep one same-turn receipt; a subsequent real keydown
      // clears it, so typing the same character again is not discarded.
      this.terminalText = event.data;
      this.flushTimer = setTimeout(() => {
        this.flushTimer = null;
        if (!this.disposed && this.compositionId === null) this.clearTerminal();
      }, 0);
    }, { signal });
    this.textarea.addEventListener('copy', event => this.copy(event, false), { signal });
    this.textarea.addEventListener('cut', event => this.copy(event, true), { signal });
    this.textarea.addEventListener('paste', event => {
      this.record(event);
      if (this.compositionId !== null || this.pending !== null || !event.clipboardData) {
        if (this.pending !== null) event.preventDefault();
        return;
      }
      event.preventDefault();
      this.clearTerminal();
      this.attempt(() => {
        const text = event.clipboardData!.getData('text/plain');
        const latexFlavor = event.clipboardData!.getData(LATEX_MIME);
        const custom = event.clipboardData!.getData(FRAGMENT_MIME);
        if (text.length > MAX_CLIPBOARD || latexFlavor.length > MAX_CLIPBOARD || custom.length > MAX_CLIPBOARD) throw new Error('粘贴内容超过当前公式输入预算');
        let latex: string | null = latexFlavor || null;
        let structure: StructuredClipboardFragment | null = null;
        if (custom) {
          try {
            const fragment = JSON.parse(custom);
            if (fragment?.schemaVersion === 2 && typeof fragment.latex === 'string' && fragment.structure && typeof fragment.structure === 'object') {
              latex = fragment.latex;
              structure = fragment.structure as StructuredClipboardFragment;
            } else if (fragment?.schemaVersion === 1 && typeof fragment.latex === 'string') latex = fragment.latex;
          } catch { /* Invalid custom flavor cannot suppress the plain/LaTeX fallback. */ }
        }
        return port.paste(text, latex, structure);
      }, false);
    }, { signal });
    this.textarea.addEventListener('focus', () => port.focus(true), { signal });
    this.textarea.addEventListener('blur', () => {
      port.focus(false);
      if (this.pending !== null) return;
      if (this.compositionId !== null) {
        const id = this.compositionId;
        if (!this.attempt(() => port.cancelComposition(id), false)) {
          this.hold({kind:'composition',id,text:this.latestPreedit,cancel:true});
          return;
        }
        this.compositionId = null;
      }
      this.clearTerminal();
    }, { signal });
  }
  private attempt<T>(operation: () => T, fallback: T): T {
    if (this.disposed) return fallback;
    try { return operation(); } catch (error) { this.port.error(error); return fallback; }
  }
  private record(event: Event): void {
    const input = event as InputEvent;
    this.audit.push({ type: event.type, key: event instanceof KeyboardEvent ? event.key : null,
      inputType: input.inputType ?? null, data: input.data ?? null, trusted: event.isTrusted,
      composing: input.isComposing ?? null, compositionId: this.compositionId, value: this.textarea.value.slice(0, 128) });
    if (this.audit.length > 512) this.audit.shift();
  }
  private terminal(event: InputEvent): boolean {
    if (this.terminalText === null) return false;
    return event.inputType === 'insertFromComposition' || event.inputType === 'insertCompositionText'
      || (event.inputType === 'insertText' && event.data === this.terminalText && this.textarea.value === this.terminalText);
  }
  private beforeInput(event: InputEvent): void {
    this.record(event);
    if (this.pending !== null) { if(event.cancelable) event.preventDefault(); return; }
    if (event.isComposing || this.compositionId !== null) return;
    if (this.terminal(event)) { if (event.cancelable) event.preventDefault(); return; }
    if (!event.cancelable) return; // Native DOM mutation is handled once in input.
    event.preventDefault();
    const data = event.data ?? '';
    if (this.applyInput(event.inputType, data)) this.textarea.value = '';
    else this.hold({kind:'input',inputType:event.inputType,text:data});
  }
  private afterInput(event: InputEvent): void {
    this.record(event);
    // Native IMEs anchor their candidate UI to the composition range. A narrow
    // scrolling proxy shifts that range away from the mathematical caret.
    this.textarea.scrollLeft = 0;
    this.textarea.scrollTop = 0;
    if (this.pending !== null) { this.textarea.value=this.pending.text; return; }
    if (event.isComposing || this.compositionId !== null || this.terminal(event)) return;
    const data=event.data ?? this.textarea.value;
    if (this.applyInput(event.inputType, data)) this.textarea.value = '';
    else this.hold({kind:'input',inputType:event.inputType,text:data});
  }
  private applyInput(inputType: string, data: string): boolean {
    return this.attempt(() => {
      switch (inputType) {
        case 'insertText':
        case 'insertReplacementText':
        case 'insertFromComposition': return data ? this.port.insert(data) : true;
        case 'insertFromPaste': return this.port.paste(data, null, null);
        case 'deleteContentBackward': return this.port.remove(true);
        case 'deleteContentForward': return this.port.remove(false);
        case 'historyUndo': return this.port.history(false);
        case 'historyRedo': return this.port.history(true);
        case 'insertLineBreak': return this.port.lineBreak?.() ?? true;
        case 'insertParagraph': return this.port.paragraph?.() ?? true;
        default: throw new Error(`当前输入桥尚未支持浏览器操作：${inputType}`);
      }
    }, false);
  }
  private copy(event: ClipboardEvent, cut: boolean): void {
    this.record(event);
    if (this.compositionId !== null || !event.clipboardData) return;
    this.attempt(() => {
      const fragment = this.port.copy();
      if (!fragment || fragment.latex === null) return false;
      event.preventDefault();
      event.clipboardData!.setData('text/plain', fragment.latex);
      event.clipboardData!.setData(LATEX_MIME, fragment.latex);
      // Never persist live node IDs, formula IDs, handles or session objects.
      event.clipboardData!.setData(FRAGMENT_MIME, fragment.structure
        ? JSON.stringify({ schemaVersion: 2, latex: fragment.latex, structure: fragment.structure })
        : JSON.stringify({ schemaVersion: 1, latex: fragment.latex }));
      if (cut) return this.port.cut(fragment);
      return true;
    }, false);
  }
  private hold(pending: PendingInput): void {
    this.pending = pending;
    if(this.flushTimer!==null){clearTimeout(this.flushTimer);this.flushTimer=null;}
    this.terminalText=null;
    this.textarea.value=pending.text;
    this.setPreedit(pending.text);
    this.textarea.setAttribute('title','输入尚未提交：按 Enter 重试，按 Escape 取消');
    this.port.error(new Error('输入未提交，已保留原文；按 Enter 重试或 Escape 取消。'));
  }
  get pendingInput(): PendingInput | null { return this.pending ? {...this.pending} : null; }
  retryPendingInput(): boolean {
    const pending=this.pending;if(!pending||this.disposed)return false;
    const accepted=pending.kind==='composition'
      ?this.attempt(()=>pending.cancel?this.port.cancelComposition(pending.id):this.port.commitComposition(pending.id,pending.text),false)
      :this.applyInput(pending.inputType,pending.text);
    if(accepted){if(pending.kind==='composition')this.compositionId=null;this.pending=null;this.latestPreedit='';this.clearTerminal();this.setPreedit('');this.textarea.removeAttribute('title');}
    return accepted;
  }
  discardPendingInput(): boolean {
    const pending=this.pending;if(!pending||this.disposed)return false;
    if(pending.kind==='composition'&&!this.attempt(()=>this.port.cancelComposition(pending.id),false))return false;
    if(pending.kind==='composition')this.compositionId=null;
    this.pending=null;this.latestPreedit='';this.clearTerminal();this.setPreedit('');this.textarea.removeAttribute('title');return true;
  }
  private clearTerminal(): void {
    if(this.pending!==null)return;
    if (this.flushTimer !== null) { clearTimeout(this.flushTimer); this.flushTimer = null; }
    this.terminalText = null;
    if (this.compositionId === null) this.textarea.value = '';
  }
  get events(): InputAudit[] { return this.audit.map(event => ({ ...event })); }
  get composing(): boolean { return this.compositionId !== null; }
  get hasUncommittedInput(): boolean { return this.compositionId !== null || this.pending !== null; }
  focus(): void { if (!this.disposed && !this.textarea.disabled) this.textarea.focus({ preventScroll: true }); }
  setEnabled(enabled: boolean): void { if (!this.disposed) this.textarea.disabled = !enabled; }
  setDescription(text: string): void { this.textarea.setAttribute('aria-description', text); }
  setPreedit(text: string): void { this.preedit.textContent = text.slice(0, 256); this.preedit.hidden = text === ''; }
  setAnchor(x: number, y: number, height: number): void {
    if (![x,y,height].every(Number.isFinite) || height < 0) throw new Error('Invalid native input anchor');
    const win = this.textarea.ownerDocument.defaultView!;
    const left = Math.max(0, Math.min(x, win.innerWidth - 4));
    const top = Math.max(0, Math.min(y, win.innerHeight - 16));
    const size = Math.max(12, Math.min(height, 128));
    this.textarea.style.left = `${left}px`; this.textarea.style.top = `${top}px`;
    // This is an invisible native input surface, not a formula box. Give the
    // current composition room within the viewport so the browser does not
    // horizontally scroll its start off the core-provided caret anchor.
    this.textarea.style.width = `${Math.max(4, win.innerWidth-left-4)}px`;
    this.textarea.style.height = `${size}px`; this.textarea.style.fontSize = `${size}px`;
    this.textarea.scrollLeft = 0;
    this.textarea.scrollTop = 0;
    this.preedit.style.left = `${left}px`; this.preedit.style.top = `${Math.min(top+size+4,win.innerHeight-28)}px`;
  }
  dispose(): void {
    if (this.disposed) return;
    if (this.compositionId !== null) this.attempt(() => this.port.cancelComposition(this.compositionId!), false);
    this.compositionId = null; this.pending = null; this.latestPreedit=''; this.clearTerminal();
    this.disposed = true; this.abort.abort(); this.textarea.remove(); this.preedit.remove();
  }
}
