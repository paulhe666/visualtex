// Code blocks (DESIGN.md §6.5): plain text in a language, highlighted the way VS Code highlights it — the app
// provides the highlighter (VS Code's own grammars and its Dark+ colours, through Shiki); this module draws the
// block (a header with the language and Copy over the code), turns the tokens into decorations and gives the block
// a code editor's keys (Tab indents, Enter keeps the indentation, the arrows step out at its ends).
import type { Node } from "prosemirror-model";
import { Plugin, PluginKey, Selection, TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView, type NodeView, type ViewMutationRecord } from "prosemirror-view";
import { newNodeId } from "./ids.ts";

/** A piece of a line in one colour. */
export interface CodeToken {
  content: string;
  color?: string;
  /** 1 italic, 2 bold, 4 underline (VS Code's font style bits) */
  fontStyle?: number;
}

/** The app's highlighter. */
export interface CodeHighlighter {
  /** The tokens of each line; null while the language's grammar is still loading (then `onReady` fires). */
  tokens(code: string, language: string): CodeToken[][] | null;
  /** Called (with nothing) whenever a grammar finished loading; returns the unsubscribe. */
  onReady(listener: () => void): () => void;
}

/** The languages the block offers ("" = plain text), with the names VS Code shows. */
export const CODE_LANGUAGES: Array<{ id: string; label: string }> = [
  { id: "", label: "Plain Text" },
  { id: "python", label: "Python" },
  { id: "javascript", label: "JavaScript" },
  { id: "typescript", label: "TypeScript" },
  { id: "c", label: "C" },
  { id: "cpp", label: "C++" },
  { id: "csharp", label: "C#" },
  { id: "java", label: "Java" },
  { id: "kotlin", label: "Kotlin" },
  { id: "swift", label: "Swift" },
  { id: "go", label: "Go" },
  { id: "rust", label: "Rust" },
  { id: "matlab", label: "MATLAB" },
  { id: "julia", label: "Julia" },
  { id: "r", label: "R" },
  { id: "fortran", label: "Fortran" },
  { id: "haskell", label: "Haskell" },
  { id: "scala", label: "Scala" },
  { id: "ruby", label: "Ruby" },
  { id: "php", label: "PHP" },
  { id: "lua", label: "Lua" },
  { id: "bash", label: "Shell" },
  { id: "powershell", label: "PowerShell" },
  { id: "sql", label: "SQL" },
  { id: "json", label: "JSON" },
  { id: "yaml", label: "YAML" },
  { id: "toml", label: "TOML" },
  { id: "xml", label: "XML" },
  { id: "html", label: "HTML" },
  { id: "css", label: "CSS" },
  { id: "latex", label: "LaTeX" },
  { id: "markdown", label: "Markdown" },
  { id: "mermaid", label: "Mermaid" },
  { id: "dockerfile", label: "Dockerfile" },
  { id: "ini", label: "INI" },
];

const INDENT = "    ";

function codeBlockAt(state: EditorState): { node: Node; pos: number } | null {
  const { $from } = state.selection;
  return $from.parent.type.name === "code_block" ? { node: $from.parent, pos: $from.before() } : null;
}

/** A code block from the caret: an empty line becomes one, otherwise the text after the caret goes into a new one. */
export function insertCodeBlock(language = ""): Command {
  return (state, dispatch) => {
    const { $from, empty } = state.selection;
    const type = state.schema.nodes.code_block;
    if (!$from.parent.isTextblock || $from.parent.type.spec.code) return false;
    const parent = $from.node(-1);
    if (!parent.canReplaceWith($from.index(-1), $from.index(-1) + 1, type)) return false;
    if (!dispatch) return true;
    const tr = state.tr;
    if (!empty) tr.deleteSelection();
    let $at = tr.selection.$from;
    if ($at.parentOffset > 0) {
      tr.split($at.pos);
      $at = tr.selection.$from;
    }
    tr.setBlockType($at.pos, $at.pos, type, { id: newNodeId(), language });
    // The text goes on after the block.
    const after = tr.selection.$from.after();
    if (after >= tr.doc.content.size || !tr.doc.resolve(after).nodeAfter) tr.insert(after, state.schema.nodes.paragraph.create());
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Enter in code: a new line indented as this one (one step more after "{", "[", "(" or ":"). */
const codeEnter: Command = (state, dispatch) => {
  const block = codeBlockAt(state);
  if (!block) return false;
  if (dispatch) {
    const { $from } = state.selection;
    const before = block.node.textBetween(0, $from.parentOffset);
    const line = before.slice(before.lastIndexOf("\n") + 1);
    let indent = /^[ \t]*/.exec(line)?.[0] ?? "";
    if (/[{[(:]\s*$/.test(line)) indent += INDENT;
    dispatch(state.tr.insertText("\n" + indent).scrollIntoView());
  }
  return true;
};

/** The lines the selection touches: their start positions. */
function selectedLineStarts(state: EditorState, block: { node: Node; pos: number }): number[] {
  const start = block.pos + 1;
  const text = block.node.textContent;
  const from = state.selection.from - start;
  const to = state.selection.to - start;
  const starts: number[] = [];
  let lineStart = text.lastIndexOf("\n", Math.max(0, from - 1));
  lineStart = from === 0 ? 0 : lineStart + 1;
  for (let at = lineStart; at <= to; ) {
    starts.push(start + at);
    const next = text.indexOf("\n", at);
    if (next < 0 || next >= to) break;
    at = next + 1;
  }
  return starts;
}

/** Tab in code: spaces to the next stop; with several lines selected, each one in by a step. */
const codeTab: Command = (state, dispatch) => {
  const block = codeBlockAt(state);
  if (!block) return false;
  if (!dispatch) return true;
  const lines = selectedLineStarts(state, block);
  if (state.selection.empty || lines.length < 2) {
    const { $from } = state.selection;
    const before = block.node.textBetween(0, $from.parentOffset);
    const column = before.length - (before.lastIndexOf("\n") + 1);
    dispatch(state.tr.insertText(" ".repeat(INDENT.length - (column % INDENT.length))).scrollIntoView());
    return true;
  }
  const tr = state.tr;
  for (const at of lines.reverse()) tr.insertText(INDENT, at);
  dispatch(tr);
  return true;
};

/** Shift+Tab in code: the lines out by a step. */
const codeShiftTab: Command = (state, dispatch) => {
  const block = codeBlockAt(state);
  if (!block) return false;
  if (!dispatch) return true;
  const tr = state.tr;
  for (const at of selectedLineStarts(state, block).reverse()) {
    const lead = /^ {1,4}|^\t/.exec(state.doc.textBetween(at, Math.min(at + INDENT.length, block.pos + block.node.nodeSize - 1)))?.[0];
    if (lead) tr.delete(at, at + lead.length);
  }
  if (tr.docChanged) dispatch(tr);
  return true;
};

/** Shift+Enter in code is a new line too (a code block has no line breaks other than newlines). */
const codeShiftEnter: Command = (state, dispatch) => {
  if (!codeBlockAt(state)) return false;
  if (dispatch) dispatch(state.tr.insertText("\n").scrollIntoView());
  return true;
};

/** Backspace at the very start of a code block: an empty one becomes body text; a full one stays as it is. */
const codeBackspace: Command = (state, dispatch) => {
  const block = codeBlockAt(state);
  if (!block || !state.selection.empty || state.selection.$from.parentOffset > 0) return false;
  if (block.node.content.size > 0) return true;
  if (dispatch) dispatch(state.tr.setBlockType(block.pos + 1, block.pos + 1, state.schema.nodes.paragraph, { id: null }));
  return true;
};

/** ↓ on the last line / ↑ on the first: out of the block (to a new line when nothing follows / precedes it). */
function codeExit(direction: 1 | -1): Command {
  return (state, dispatch, view) => {
    const block = codeBlockAt(state);
    if (!block || !state.selection.empty) return false;
    if (view && !view.endOfTextblock(direction > 0 ? "down" : "up")) return false;
    const outside = direction > 0 ? block.pos + block.node.nodeSize : block.pos;
    const $outside = state.doc.resolve(outside);
    if ((direction > 0 ? $outside.nodeAfter : $outside.nodeBefore) && Selection.findFrom($outside, direction, true)) return false;
    if (dispatch) {
      const tr = state.tr.insert(outside, state.schema.nodes.paragraph.create());
      dispatch(tr.setSelection(TextSelection.create(tr.doc, outside + 1)).scrollIntoView());
    }
    return true;
  };
}

/** Ctrl+Enter: out of the block, to the line after it (a new one when nothing follows), as in Typora. */
const codeLeave: Command = (state, dispatch) => {
  const block = codeBlockAt(state);
  if (!block) return false;
  if (dispatch) {
    const after = block.pos + block.node.nodeSize;
    const tr = state.tr;
    if (!tr.doc.resolve(after).nodeAfter?.isTextblock) tr.insert(after, state.schema.nodes.paragraph.create());
    dispatch(tr.setSelection(TextSelection.create(tr.doc, after + 1)).scrollIntoView());
  }
  return true;
};

/** The keys a code block takes before the document's own. */
export const codeBlockKeys: Record<string, Command> = {
  Enter: codeEnter,
  "Mod-Enter": codeLeave,
  "Shift-Enter": codeShiftEnter,
  Tab: codeTab,
  "Shift-Tab": codeShiftTab,
  Backspace: codeBackspace,
  ArrowDown: codeExit(1),
  ArrowUp: codeExit(-1),
};

// ------------------------------------------------------------- the plugin --

const codeKey = new PluginKey<DecorationSet>("visualtex-document-code");

function styleOf(token: CodeToken): string | null {
  const parts: string[] = [];
  if (token.color) parts.push(`color:${token.color}`);
  if (token.fontStyle && token.fontStyle & 1) parts.push("font-style:italic");
  if (token.fontStyle && token.fontStyle & 2) parts.push("font-weight:bold");
  if (token.fontStyle && token.fontStyle & 4) parts.push("text-decoration:underline");
  return parts.length ? parts.join(";") : null;
}

function highlight(doc: Node, highlighter: CodeHighlighter): DecorationSet {
  const decorations: Decoration[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name !== "code_block") return true;
    const language = node.attrs.language as string;
    if (!language || !node.content.size) return false;
    const lines = highlighter.tokens(node.textContent, language);
    if (!lines) return false;
    let at = pos + 1;
    for (const line of lines) {
      for (const token of line) {
        const style = styleOf(token);
        if (style && token.content) decorations.push(Decoration.inline(at, at + token.content.length, { style }));
        at += token.content.length;
      }
      at += 1; // the newline
    }
    return false;
  });
  return DecorationSet.create(doc, decorations);
}

export interface CodeBlockOptions {
  highlighter: CodeHighlighter | null;
  /** names in the app's language */
  labels: { plainText: string; copy: string; copied: string; language: string };
  /** a diagram's code drawn as SVG (Mermaid, as Typora draws it); null: shown as code only */
  diagram?: ((language: string, code: string) => Promise<string> | null) | null;
}

class CodeBlockView implements NodeView {
  dom: HTMLElement;
  contentDOM: HTMLElement;
  private header: HTMLElement;
  private select: HTMLSelectElement;
  private copyButton: HTMLButtonElement;
  private drawing: HTMLElement | null = null;
  private drawn = "";
  private drawTimer = 0;

  constructor(
    private node: Node,
    private view: EditorView,
    private getPos: () => number | undefined,
    private options: CodeBlockOptions,
  ) {
    this.dom = document.createElement("div");
    this.dom.className = "vtx-code-block";
    this.header = document.createElement("div");
    this.header.className = "vtx-code-header";
    this.header.contentEditable = "false";
    this.select = document.createElement("select");
    this.select.className = "vtx-code-language";
    this.select.title = options.labels.language;
    this.select.setAttribute("aria-label", options.labels.language);
    for (const language of CODE_LANGUAGES) {
      const option = document.createElement("option");
      option.value = language.id;
      option.textContent = language.id ? language.label : options.labels.plainText;
      this.select.append(option);
    }
    this.select.addEventListener("change", this.onLanguage);
    this.copyButton = document.createElement("button");
    this.copyButton.type = "button";
    this.copyButton.className = "vtx-code-copy";
    this.copyButton.textContent = options.labels.copy;
    this.copyButton.addEventListener("mousedown", (event) => event.preventDefault());
    this.copyButton.addEventListener("click", this.onCopy);
    this.header.append(this.select, this.copyButton);
    const pre = document.createElement("pre");
    pre.className = "vtx-code-pre";
    this.contentDOM = document.createElement("code");
    this.contentDOM.className = "vtx-code-content";
    this.contentDOM.spellcheck = false;
    pre.append(this.contentDOM);
    this.dom.append(this.header, pre);
    this.sync();
  }

  private sync() {
    const language = this.node.attrs.language as string;
    if (this.select.value !== language) this.select.value = CODE_LANGUAGES.some((l) => l.id === language) ? language : "";
    this.dom.dataset.language = language;
    this.scheduleDrawing();
  }

  /** A diagram block: its picture under the code, drawn again a moment after the code changes. */
  private scheduleDrawing() {
    const language = this.node.attrs.language as string;
    const code = this.node.textContent;
    const key = `${language}\n${code}`;
    if (!this.options.diagram || language !== "mermaid" || !code.trim()) {
      this.drawing?.remove();
      this.drawing = null;
      this.drawn = "";
      return;
    }
    if (key === this.drawn) return;
    window.clearTimeout(this.drawTimer);
    this.drawTimer = window.setTimeout(() => {
      this.drawn = key;
      const pending = this.options.diagram?.(language, code);
      if (!pending) return;
      void pending.then(
        (svg) => {
          if (this.drawn !== key) return;
          if (!this.drawing) {
            this.drawing = document.createElement("div");
            this.drawing.className = "vtx-code-diagram";
            this.drawing.contentEditable = "false";
            this.dom.append(this.drawing);
          }
          this.drawing.classList.remove("is-error");
          this.drawing.innerHTML = svg;
        },
        (error: unknown) => {
          if (this.drawn !== key || !this.drawing) return;
          this.drawing.classList.add("is-error");
          this.drawing.textContent = String((error as Error)?.message ?? error).split("\n")[0];
        },
      );
    }, this.drawn ? 400 : 0);
  }

  private onLanguage = () => {
    const pos = this.getPos();
    if (pos === undefined) return;
    this.view.dispatch(this.view.state.tr.setNodeAttribute(pos, "language", this.select.value));
    this.view.focus();
  };

  private onCopy = () => {
    void navigator.clipboard.writeText(this.node.textContent);
    this.copyButton.textContent = this.options.labels.copied;
    window.setTimeout(() => {
      this.copyButton.textContent = this.options.labels.copy;
    }, 1200);
  };

  update(node: Node) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.sync();
    return true;
  }

  // The header and the diagram are the block's own; the code is ProseMirror's.
  stopEvent(event: Event) {
    return this.header.contains(event.target as globalThis.Node) || !!this.drawing?.contains(event.target as globalThis.Node);
  }

  ignoreMutation(mutation: ViewMutationRecord) {
    if (mutation.type === "selection") return false;
    return !this.contentDOM.contains(mutation.target);
  }

  destroy() {
    window.clearTimeout(this.drawTimer);
    this.select.removeEventListener("change", this.onLanguage);
    this.copyButton.removeEventListener("click", this.onCopy);
  }
}

/** Code blocks: their view, their highlighting (kept as it is while an input method composes in one). */
export function codeBlockPlugin(options: CodeBlockOptions): Plugin<DecorationSet> {
  const { highlighter } = options;
  return new Plugin<DecorationSet>({
    key: codeKey,
    state: {
      init: (_config, state) => (highlighter ? highlight(state.doc, highlighter) : DecorationSet.empty),
      apply(tr: Transaction, value, _old, state) {
        if (!highlighter) return value;
        if (tr.getMeta(codeKey)) return highlight(state.doc, highlighter);
        if (!tr.docChanged) return value;
        if (tr.getMeta("composition")) return value.map(tr.mapping, tr.doc);
        return highlight(state.doc, highlighter);
      },
    },
    view(view) {
      const unsubscribe = highlighter?.onReady(() => {
        if (!view.isDestroyed) view.dispatch(view.state.tr.setMeta(codeKey, true));
      });
      return { destroy: () => unsubscribe?.() };
    },
    props: {
      decorations: (state) => codeKey.getState(state),
      nodeViews: {
        code_block: (node, view, getPos) => new CodeBlockView(node, view, getPos, options),
      },
    },
  });
}
