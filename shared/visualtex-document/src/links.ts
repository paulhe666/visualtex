// Links (Markdown's [text](url), LaTeX's \href): Ctrl+K and the menus ask the app for its link dialog
// (LINK_REQUEST_EVENT on the editor); these commands make, change and remove the link it shows.
import type { Mark } from "prosemirror-model";
import { TextSelection, type Command, type EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

/** Dispatched (bubbling) on the editor's DOM when Ctrl+K or a menu asks for the link dialog. */
export const LINK_REQUEST_EVENT = "vtx-link-request";

export const requestLink: Command = (_state, _dispatch, view) => {
  if (!view) return false;
  view.dom.dispatchEvent(new CustomEvent(LINK_REQUEST_EVENT, { bubbles: true }));
  return true;
};

export interface LinkRange {
  from: number;
  to: number;
  href: string;
  title: string | null;
}

/** The link at the caret or the selection's start: the whole run of text it covers. */
export function linkAt(state: EditorState): LinkRange | null {
  const type = state.schema.marks.link;
  const { $from } = state.selection;
  if (!type || !$from.parent.inlineContent) return null;
  const find = (marks: readonly Mark[] | undefined) => marks?.find((m) => m.type === type);
  const mark = find($from.nodeAfter?.marks) ?? find($from.nodeBefore?.marks);
  if (!mark) return null;
  const parent = $from.parent;
  const start = $from.start();
  // The children around the caret that carry this same link.
  const spans: Array<{ from: number; to: number; linked: boolean }> = [];
  parent.forEach((child, offset) => spans.push({ from: start + offset, to: start + offset + child.nodeSize, linked: mark.isInSet(child.marks) }));
  let i = spans.findIndex((s) => s.linked && s.from <= $from.pos && $from.pos <= s.to);
  if (i < 0) return null;
  let first = i;
  while (first > 0 && spans[first - 1].linked) first--;
  while (i < spans.length - 1 && spans[i + 1].linked) i++;
  return { from: spans[first].from, to: spans[i].to, href: mark.attrs.href, title: mark.attrs.title ?? null };
}

/** An address as typed: "example.com" gets https://; mail addresses mailto:. */
export function normalizeHref(href: string): string {
  const value = href.trim();
  if (!value || /^[a-z][a-z0-9+.-]*:/i.test(value) || value.startsWith("#") || value.startsWith("/")) return value;
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)) return `mailto:${value}`;
  return `https://${value}`;
}

/** The text from..to links to `href` (replacing any link there); the caret goes after it. */
export function setLink(view: EditorView, from: number, to: number, href: string, title: string | null = null) {
  const type = view.state.schema.marks.link;
  const tr = view.state.tr.removeMark(from, to, type).addMark(from, to, type.create({ href: normalizeHref(href), title }));
  tr.setSelection(TextSelection.create(tr.doc, to)).removeStoredMark(type);
  view.dispatch(tr.scrollIntoView());
}

export function removeLink(view: EditorView, from: number, to: number) {
  view.dispatch(view.state.tr.removeMark(from, to, view.state.schema.marks.link).scrollIntoView());
}

/** `text` at the caret, linked to `href`; typing after it is plain again. */
export function insertLink(view: EditorView, text: string, href: string) {
  const type = view.state.schema.marks.link;
  const { from, to } = view.state.selection;
  const content = text || href;
  if (!content) return;
  const tr = view.state.tr.insertText(content, from, to);
  tr.addMark(from, from + content.length, type.create({ href: normalizeHref(href), title: null }));
  tr.setSelection(TextSelection.create(tr.doc, from + content.length)).removeStoredMark(type);
  view.dispatch(tr.scrollIntoView());
}
