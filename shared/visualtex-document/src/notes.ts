// Footnotes, citations and the reference list (DESIGN.md §8.3). A footnote is a mark in the text that keeps its
// note's text (inline formulas as $…$); its number is its place among the document's footnotes, and the page view
// prints the note at the foot of the page its mark is on. A citation names entries of the reference list (the
// document's bibliography block, \begin{thebibliography}); it shows their numbers, [1] or [1, 3–5].
import type { DOMOutputSpec, Node, NodeSpec, Schema } from "prosemirror-model";
import { NodeSelection, Plugin, PluginKey, TextSelection, type Command, type EditorState } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { newNodeId } from "./ids.ts";
import { splitInlineMath } from "./mathText.ts";

// ------------------------------------------------------------------ schema --

export const noteNodes: Record<string, NodeSpec> = {
  /** A footnote: its number in the text; its note's text (with $…$ formulas) in `text`. */
  footnote: {
    group: "inline",
    inline: true,
    atom: true,
    selectable: true,
    attrs: { id: { default: null as string | null }, text: { default: "" } },
    parseDOM: [
      {
        tag: "span.vtx-footnote",
        getAttrs: (dom: HTMLElement | string) =>
          typeof dom === "string" ? {} : { id: dom.getAttribute("data-id"), text: dom.getAttribute("data-text") ?? "" },
      },
    ],
    toDOM: (node): DOMOutputSpec => [
      "span",
      { class: "vtx-footnote", ...(node.attrs.id ? { "data-id": node.attrs.id } : {}), "data-text": node.attrs.text },
    ],
  },

  /** A citation of reference list entries (their ids). */
  cite: {
    group: "inline",
    inline: true,
    atom: true,
    selectable: true,
    attrs: { targets: { default: [] as string[] } },
    parseDOM: [
      {
        tag: "span.vtx-cite",
        getAttrs: (dom: HTMLElement | string) =>
          typeof dom === "string" ? {} : { targets: (dom.getAttribute("data-targets") ?? "").split(",").filter(Boolean) },
      },
    ],
    toDOM: (node): DOMOutputSpec => ["span", { class: "vtx-cite", "data-targets": (node.attrs.targets as string[]).join(",") }],
  },

  /** The reference list: numbered entries, written where it stands (usually the end). */
  bibliography: {
    group: "block",
    content: "bib_entry+",
    defining: true,
    isolating: true,
    attrs: { id: { default: null as string | null } },
    parseDOM: [{ tag: "div.vtx-bibliography", contentElement: "ol" }],
    toDOM: (node): DOMOutputSpec => [
      "div",
      { class: "vtx-bibliography", ...(node.attrs.id ? { "data-id": node.attrs.id } : {}) },
      ["div", { class: "vtx-bibliography-title", contenteditable: "false" }],
      ["ol", 0],
    ],
  },

  /** One reference: its text as written (with formulas); `key` is its \bibitem key when it came with one. */
  bib_entry: {
    content: "(text | math_inline | hard_break)*",
    defining: true,
    attrs: { id: { default: null as string | null }, key: { default: null as string | null } },
    parseDOM: [
      {
        tag: "li.vtx-bib-entry",
        getAttrs: (dom: HTMLElement | string) =>
          typeof dom === "string" ? {} : { id: dom.getAttribute("data-id"), key: dom.getAttribute("data-key") },
      },
    ],
    toDOM: (node): DOMOutputSpec => [
      "li",
      { class: "vtx-bib-entry", ...(node.attrs.id ? { "data-id": node.attrs.id } : {}), ...(node.attrs.key ? { "data-key": node.attrs.key } : {}) },
      0,
    ],
  },
};

// --------------------------------------------------------------- numbering --

export interface BibEntryInfo {
  id: string;
  number: number;
  key: string;
  text: string;
  pos: number;
}

/** The reference list's entries in order (the first bibliography block's). */
export function bibliographyEntries(doc: Node): BibEntryInfo[] {
  const entries: BibEntryInfo[] = [];
  let done = false;
  doc.descendants((node, pos) => {
    if (done) return false;
    if (node.type.name !== "bibliography") return true;
    node.forEach((entry, offset) => {
      if (!entry.attrs.id) return;
      entries.push({ id: entry.attrs.id, number: entries.length + 1, key: citationKey(entry), text: entry.textContent, pos: pos + 1 + offset });
    });
    done = true;
    return false;
  });
  return entries;
}

/** An entry's \bibitem key: the one it came with, else one made of its id. */
export function citationKey(entry: Node): string {
  const key = typeof entry.attrs.key === "string" ? entry.attrs.key.trim() : "";
  return key && /^[^\s{},%#\\]+$/.test(key) ? key : `ref-${entry.attrs.id}`;
}

/** The footnotes in document order (their numbers are 1, 2, …). */
export function footnotes(doc: Node): Array<{ id: string | null; text: string; pos: number; number: number }> {
  const list: Array<{ id: string | null; text: string; pos: number; number: number }> = [];
  doc.descendants((node, pos) => {
    if (node.type.name === "footnote") list.push({ id: node.attrs.id, text: node.attrs.text, pos, number: list.length + 1 });
    return true;
  });
  return list;
}

/** Citation numbers as a label: [1], [1, 2], [1–3, 5]; a missing entry shows "?". */
export function citeLabel(numbers: Array<number | null>): string {
  const known = [...new Set(numbers.filter((n): n is number => n !== null))].sort((a, b) => a - b);
  const parts: string[] = [];
  for (let i = 0; i < known.length; ) {
    let j = i;
    while (j + 1 < known.length && known[j + 1] === known[j] + 1) j++;
    parts.push(j - i >= 2 ? `${known[i]}–${known[j]}` : known.slice(i, j + 1).join(", "));
    i = j + 1;
  }
  if (numbers.some((n) => n === null) || !numbers.length) parts.push("?");
  return `[${parts.join(", ")}]`;
}

export const notesKey = new PluginKey<DecorationSet>("visualtex-document-notes");

/** The event a footnote mark sends when its note is to be edited (the app shows the editor). */
export const FOOTNOTE_EDIT_EVENT = "vtx-footnote-edit";

export interface FootnoteEditRequest {
  pos: number;
  /** the mark on the screen, for placing the editor */
  rect: DOMRect | null;
}

/** The event that asks the app for its citation picker (the "\\cite" and "/cite" menu entries). */
export const CITATION_REQUEST_EVENT = "vtx-citation-request";

/** Ask the app to show the citation picker at the caret. */
export const requestCitation: Command = (state, _dispatch, view) => {
  if (!inText(state) || state.selection.$from.parent.type.name === "bib_entry") return false;
  view?.dom.dispatchEvent(new CustomEvent(CITATION_REQUEST_EVENT, { bubbles: true }));
  return true;
};

/** Ask the app to edit the footnote at `pos` (its editor is placed at `at`: what was clicked; default its number). */
export function requestFootnoteEdit(view: EditorView, pos: number, at?: DOMRect) {
  const dom = view.nodeDOM(pos);
  const rect = at ?? (dom instanceof Element ? dom.getBoundingClientRect() : null);
  view.dom.dispatchEvent(new CustomEvent<FootnoteEditRequest>(FOOTNOTE_EDIT_EVENT, { bubbles: true, detail: { pos, rect } }));
}

class FootnoteView {
  dom: HTMLElement;
  constructor(
    private node: Node,
    private view: EditorView,
    private getPos: () => number | undefined,
    decorations: readonly Decoration[],
  ) {
    this.dom = document.createElement("sup");
    this.dom.className = "vtx-footnote";
    this.dom.contentEditable = "false";
    this.dom.addEventListener("mousedown", this.onMouseDown);
    this.render(decorations);
  }
  private onMouseDown = (event: MouseEvent) => {
    if (event.button !== 0) return;
    const pos = this.getPos();
    if (pos === undefined) return;
    event.preventDefault();
    this.view.dispatch(this.view.state.tr.setSelection(NodeSelection.create(this.view.state.doc, pos)));
    requestFootnoteEdit(this.view, pos);
  };
  private render(decorations: readonly Decoration[]) {
    let number: number | null = null;
    for (const d of decorations) {
      const spec = d.spec as { footnoteNumber?: number };
      if (spec.footnoteNumber !== undefined) number = spec.footnoteNumber;
    }
    this.dom.textContent = number === null ? "?" : String(number);
    this.dom.title = this.node.attrs.text || "";
    this.dom.dataset.id = this.node.attrs.id ?? "";
  }
  update(node: Node, decorations: readonly Decoration[]) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render(decorations);
    return true;
  }
  ignoreMutation() {
    return true;
  }
  stopEvent(event: Event) {
    return event.type === "mousedown";
  }
  destroy() {
    this.dom.removeEventListener("mousedown", this.onMouseDown);
  }
}

class CiteView {
  dom: HTMLElement;
  constructor(
    private node: Node,
    private view: EditorView,
    _getPos: () => number | undefined,
    decorations: readonly Decoration[],
  ) {
    this.dom = document.createElement("span");
    this.dom.className = "vtx-cite";
    this.dom.contentEditable = "false";
    this.dom.addEventListener("mousedown", this.onMouseDown);
    this.render(decorations);
  }
  /** Ctrl+click goes to the (first) entry. */
  private onMouseDown = (event: MouseEvent) => {
    if (!(event.ctrlKey || event.metaKey)) return;
    const first = (this.node.attrs.targets as string[])[0];
    const entry = bibliographyEntries(this.view.state.doc).find((e) => e.id === first);
    if (!entry) return;
    event.preventDefault();
    const tr = this.view.state.tr.setSelection(TextSelection.create(this.view.state.doc, entry.pos + 1));
    this.view.dispatch(tr.scrollIntoView());
    this.view.focus();
  };
  private render(decorations: readonly Decoration[]) {
    let label = "[?]";
    let tip = "";
    for (const d of decorations) {
      const spec = d.spec as { citeLabel?: string; citeTip?: string };
      if (spec.citeLabel !== undefined) {
        label = spec.citeLabel;
        tip = spec.citeTip ?? "";
      }
    }
    this.dom.textContent = label;
    this.dom.title = tip;
    this.dom.classList.toggle("is-broken", label.includes("?"));
  }
  update(node: Node, decorations: readonly Decoration[]) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render(decorations);
    return true;
  }
  ignoreMutation() {
    return true;
  }
  stopEvent(event: Event) {
    return event.type === "mousedown" && ((event as MouseEvent).ctrlKey || (event as MouseEvent).metaKey);
  }
  destroy() {
    this.dom.removeEventListener("mousedown", this.onMouseDown);
  }
}

/** Footnote and citation numbers (decorations the node views read). */
export function notesPlugin(): Plugin<DecorationSet> {
  const build = (doc: Node) => {
    const decorations: Decoration[] = [];
    const entries = new Map(bibliographyEntries(doc).map((entry) => [entry.id, entry]));
    let footnote = 0;
    doc.descendants((node, pos) => {
      if (node.type.name === "footnote") {
        decorations.push(Decoration.node(pos, pos + node.nodeSize, {}, { footnoteNumber: ++footnote }));
      } else if (node.type.name === "cite") {
        const targets = node.attrs.targets as string[];
        const found = targets.map((id) => entries.get(id) ?? null);
        decorations.push(
          Decoration.node(pos, pos + node.nodeSize, {}, {
            citeLabel: citeLabel(found.map((entry) => entry?.number ?? null)),
            citeTip: found.map((entry) => (entry ? `[${entry.number}] ${entry.text}` : "?")).join("\n"),
          }),
        );
      }
      return true;
    });
    return DecorationSet.create(doc, decorations);
  };
  return new Plugin<DecorationSet>({
    key: notesKey,
    state: {
      init: (_config, state) => build(state.doc),
      apply: (tr, value) => (tr.docChanged ? build(tr.doc) : value),
    },
    props: {
      decorations: (state) => notesKey.getState(state),
      nodeViews: {
        footnote: (node, view, getPos, decorations) => new FootnoteView(node, view, getPos, decorations),
        cite: (node, view, getPos, decorations) => new CiteView(node, view, getPos, decorations),
      },
    },
  });
}

// ---------------------------------------------------------------- commands --

const inText = (state: EditorState) => state.selection.$from.parent.inlineContent && !state.selection.$from.parent.type.spec.code;

/** A footnote at the caret (its text: `text`), then its editor opens. */
export function insertFootnote(text = ""): Command {
  return (state, dispatch, view) => {
    if (!inText(state) || state.selection.$from.parent.type.name === "bib_entry") return false;
    if (!dispatch) return true;
    const node = state.schema.nodes.footnote.create({ id: newNodeId(), text });
    const tr = state.tr.replaceSelectionWith(node, false);
    const pos = tr.mapping.map(state.selection.from, -1);
    tr.setSelection(NodeSelection.create(tr.doc, pos));
    dispatch(tr.scrollIntoView());
    if (view && !text) requestFootnoteEdit(view, pos);
    return true;
  };
}

/** Change the text of the footnote at `pos`. */
export function setFootnoteText(pos: number, text: string): Command {
  return (state, dispatch) => {
    const node = state.doc.nodeAt(pos);
    if (!node || node.type.name !== "footnote") return false;
    if (node.attrs.text === text) return true;
    if (dispatch) dispatch(state.tr.setNodeAttribute(pos, "text", text));
    return true;
  };
}

/** A citation of the entries `targets` at the caret. */
export function insertCitation(targets: string[]): Command {
  return (state, dispatch) => {
    if (!targets.length || !inText(state) || state.selection.$from.parent.type.name === "bib_entry") return false;
    if (dispatch) {
      const tr = state.tr.replaceSelectionWith(state.schema.nodes.cite.create({ targets }), false);
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/**
 * A new reference list entry with `text` (a reference list is made at the end of the document when there is none);
 * cited at the caret when `cite` (the caret stays in the text).
 */
export function addReference(text: string, cite: boolean): Command {
  return (state, dispatch) => {
    const { schema } = state;
    if (cite && (!inText(state) || state.selection.$from.parent.type.name === "bib_entry")) return false;
    if (!dispatch) return true;
    const id = newNodeId();
    const entry = schema.nodes.bib_entry.create({ id }, text ? schema.text(text) : null);
    let tr = state.tr;
    if (cite) tr = tr.replaceSelectionWith(schema.nodes.cite.create({ targets: [id] }), false);
    let list: { node: Node; pos: number } | null = null;
    tr.doc.descendants((node, pos) => {
      if (list) return false;
      if (node.type.name === "bibliography") {
        list = { node, pos };
        return false;
      }
      return true;
    });
    const found = list as { node: Node; pos: number } | null;
    if (found) tr.insert(found.pos + found.node.nodeSize - 1, entry);
    else tr.insert(tr.doc.content.size, schema.nodes.bibliography.create({ id: newNodeId() }, entry));
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** The reference list: a new empty entry at its end (the list is made at the document's end), the caret in it. */
export const addReferenceEntry: Command = (state, dispatch) => {
  const { schema } = state;
  if (!dispatch) return true;
  const entry = schema.nodes.bib_entry.create({ id: newNodeId() });
  let list: { node: Node; pos: number } | null = null;
  state.doc.descendants((node, pos) => {
    if (list) return false;
    if (node.type.name === "bibliography") {
      list = { node, pos };
      return false;
    }
    return true;
  });
  const found = list as { node: Node; pos: number } | null;
  const tr = state.tr;
  // an empty entry the caret is in (the list's last) is used as it is
  if (found && found.node.lastChild && !found.node.lastChild.content.size) {
    const at = found.pos + found.node.nodeSize - 1 - found.node.lastChild.nodeSize;
    dispatch(tr.setSelection(TextSelection.create(tr.doc, at + 1)).scrollIntoView());
    return true;
  }
  let at: number;
  if (found) {
    at = found.pos + found.node.nodeSize - 1;
    tr.insert(at, entry);
  } else {
    at = tr.doc.content.size + 1;
    tr.insert(tr.doc.content.size, schema.nodes.bibliography.create({ id: newNodeId() }, entry));
  }
  dispatch(tr.setSelection(TextSelection.create(tr.doc, at + 1)).scrollIntoView());
  return true;
};

/** Footnotes and citations in plain text (copy): [^n] and [n]. */
export function noteText(node: Node, doc: Node | null): string {
  if (node.type.name === "footnote") return `[^${node.attrs.text}]`;
  if (node.type.name === "cite" && doc) {
    const entries = new Map(bibliographyEntries(doc).map((entry) => [entry.id, entry.number]));
    return citeLabel((node.attrs.targets as string[]).map((id) => entries.get(id) ?? null));
  }
  return "";
}

// ---------------------------------------------------------- note editing --

/** A footnote's text as the content of its editor: one paragraph, its formulas as formulas, its lines as breaks. */
export function footnoteEditorDoc(schema: Schema, text: string): Node {
  const content: Node[] = [];
  text.split("\n").forEach((line, index) => {
    if (index) content.push(schema.nodes.hard_break.create());
    for (const part of splitInlineMath(line)) {
      if (part.kind === "text") {
        if (part.text) content.push(schema.text(part.text));
      } else content.push(schema.nodes.math_inline.create({ id: newNodeId(), latex: part.latex }));
    }
  });
  return schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, content));
}

/** The editor's content back as a footnote's text: formulas as $…$ (\(…\) where $…$ would not read back). */
export function footnoteEditorText(doc: Node): string {
  const parts: Array<{ text: string } | { latex: string }> = [];
  doc.forEach((block, _offset, index) => {
    if (index) parts.push({ text: "\n" });
    block.descendants((node) => {
      if (node.isText) parts.push({ text: node.text!.replace(/\$/g, "\\$") });
      else if (node.type.name === "hard_break") parts.push({ text: "\n" });
      else if (node.type.name === "math_inline" && String(node.attrs.latex).trim()) parts.push({ latex: String(node.attrs.latex).trim() });
      return true;
    });
  });
  return parts
    .map((part, i) => {
      if ("text" in part) return part.text;
      const next = parts[i + 1];
      // "$x$2" does not close at the "$": then \(x\)
      const dollar = !part.latex.includes("$") && !(next && "text" in next && /^\d/.test(next.text));
      return dollar ? `$${part.latex}$` : `\\(${part.latex}\\)`;
    })
    .join("")
    .trim();
}
