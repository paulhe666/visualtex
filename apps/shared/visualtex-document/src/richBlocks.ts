// What Markdown files bring besides text (markdown.ts): the table of contents ([toc]), HTML the document has no
// element for (shown as the browser draws it, its source edited on double-click) and task lists (the box toggles on
// a click). Node views and a plugin both apps register.
import type { Node } from "prosemirror-model";
import { Plugin, TextSelection } from "prosemirror-state";
import type { EditorView, NodeView } from "prosemirror-view";
import MarkdownIt from "markdown-it";
import { headingNumbers } from "./settings.ts";
import { renderedText } from "./sourceMarkdown.ts";

// ------------------------------------------------------------------- [toc] --

const tocViews = new Set<TocView>();

class TocView implements NodeView {
  dom: HTMLElement;
  private shown = "";

  constructor(
    private node: Node,
    private view: EditorView,
  ) {
    this.dom = document.createElement("div");
    this.dom.className = "vtx-toc";
    this.dom.contentEditable = "false";
    this.dom.addEventListener("mousedown", (event) => {
      const entry = (event.target as HTMLElement).closest<HTMLElement>("[data-pos]");
      if (!entry) return;
      event.preventDefault();
      const pos = Number(entry.dataset.pos);
      const state = this.view.state;
      if (pos >= state.doc.content.size) return;
      this.view.dispatch(state.tr.setSelection(TextSelection.near(state.doc.resolve(pos + 1))).scrollIntoView());
      this.view.focus();
    });
    tocViews.add(this);
    this.render();
  }

  render() {
    const doc = this.view.state.doc;
    const numbers = headingNumbers(doc);
    const entries: Array<{ level: number; text: string; number: string; pos: number }> = [];
    doc.forEach((child, offset) => {
      // an unnumbered section (\section*) is not in LaTeX's table of contents
      if (child.type.name !== "heading" || child.attrs.unnumbered) return;
      entries.push({ level: child.attrs.level as number, text: renderedText(child), number: numbers.get(offset) ?? "", pos: offset });
    });
    const key = JSON.stringify(entries);
    if (key === this.shown) return;
    this.shown = key;
    this.dom.replaceChildren(
      ...entries.map((entry) => {
        const line = document.createElement("div");
        line.className = "vtx-toc-entry";
        line.dataset.level = String(entry.level);
        line.dataset.pos = String(entry.pos);
        if (entry.number) {
          const number = document.createElement("span");
          number.className = "vtx-toc-number";
          number.textContent = entry.number;
          line.append(number);
        }
        line.append(entry.text || " ");
        return line;
      }),
    );
    if (!entries.length) this.dom.dataset.empty = "";
    else delete this.dom.dataset.empty;
  }

  update(node: Node) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.render();
    return true;
  }

  ignoreMutation() {
    return true;
  }

  stopEvent(event: Event) {
    return event.type === "mousedown" && !!(event.target as HTMLElement).closest?.("[data-pos]");
  }

  destroy() {
    tocViews.delete(this);
  }
}

// -------------------------------------------------------------------- HTML --

const ALLOWED = new Set(
  "a abbr b bdi bdo blockquote br caption center cite code col colgroup dd del details dfn div dl dt em figcaption figure font h1 h2 h3 h4 h5 h6 hr i img ins kbd li mark ol p pre q rp rt ruby s samp small span strike strong sub summary sup table tbody td tfoot th thead time tr tt u ul var".split(" "),
);
const ALLOWED_ATTRIBUTES = new Set("align alt border cellpadding cellspacing class color colspan dir face height href lang open rowspan size span src start style title type valign width".split(" "));

/** HTML from a file as safe DOM: known elements and attributes only, no scripts, no javascript: addresses. */
export function safeHtml(html: string): DocumentFragment {
  const parsed = new DOMParser().parseFromString(`<body>${html}</body>`, "text/html");
  const clean = (from: globalThis.Node, to: globalThis.Node) => {
    from.childNodes.forEach((child) => {
      if (child.nodeType === 3) to.appendChild(document.createTextNode(child.nodeValue ?? ""));
      if (child.nodeType !== 1) return;
      const element = child as Element;
      const name = element.localName;
      if (!ALLOWED.has(name)) {
        // an unknown wrapper's text stays; scripts, styles and frames go
        if (!/^(script|style|iframe|object|embed|frame|frameset|link|meta|base|form|input|button|textarea|select)$/.test(name)) clean(element, to);
        return;
      }
      const copy = document.createElement(name);
      for (const attribute of Array.from(element.attributes)) {
        const attr = attribute.name.toLowerCase();
        if (!ALLOWED_ATTRIBUTES.has(attr)) continue;
        if ((attr === "href" || attr === "src") && /^\s*(javascript|vbscript|data:text)/i.test(attribute.value)) continue;
        if (attr === "style" && /expression|url\s*\(|@import/i.test(attribute.value)) continue;
        copy.setAttribute(attr, attribute.value);
      }
      if (name === "a") {
        copy.setAttribute("target", "_blank");
        copy.setAttribute("rel", "noopener noreferrer");
      }
      clean(element, copy);
      to.appendChild(copy);
    });
  };
  const fragment = document.createDocumentFragment();
  clean(parsed.body, fragment);
  return fragment;
}

let renderer: MarkdownIt | null = null;
const markdownHtml = () => (renderer ??= new MarkdownIt("default", { html: true, linkify: true }));

class HtmlView implements NodeView {
  dom: HTMLElement;
  private editor: HTMLTextAreaElement | null = null;

  constructor(
    private node: Node,
    private view: EditorView,
    private getPos: () => number | undefined,
  ) {
    this.dom = document.createElement("div");
    this.dom.className = "vtx-html";
    this.dom.contentEditable = "false";
    this.dom.addEventListener("dblclick", (event) => {
      if (this.editor || !this.view.editable) return;
      event.preventDefault();
      this.edit();
    });
    this.render();
  }

  private render() {
    const html = String(this.node.attrs.html ?? "");
    const comment = /^\s*<!--([\s\S]*?)-->\s*$/.exec(html);
    this.dom.classList.toggle("is-comment", !!comment);
    if (comment) {
      // a comment is not shown in the page; its text, faint, where it is
      const label = document.createElement("span");
      label.className = "vtx-html-comment";
      label.textContent = `<!-- ${comment[1].trim()} -->`;
      this.dom.replaceChildren(label);
      return;
    }
    // HTML with Markdown inside (<details> … </details> around text): the Markdown drawn too
    this.dom.replaceChildren(safeHtml(/\n\s*\n/.test(html) ? markdownHtml().render(html) : html));
  }

  private edit() {
    const area = document.createElement("textarea");
    area.className = "vtx-html-source";
    area.value = String(this.node.attrs.html ?? "");
    area.rows = Math.min(16, Math.max(2, area.value.split("\n").length));
    area.spellcheck = false;
    const finish = (save: boolean) => {
      if (!this.editor) return;
      const value = area.value;
      this.editor = null;
      const pos = this.getPos();
      if (save && pos !== undefined && value !== this.node.attrs.html) {
        this.view.dispatch(value.trim() ? this.view.state.tr.setNodeAttribute(pos, "html", value) : this.view.state.tr.delete(pos, pos + this.node.nodeSize));
      } else this.render();
      this.view.focus();
    };
    area.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        finish(false);
      } else if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
        event.preventDefault();
        finish(true);
      }
    });
    area.addEventListener("blur", () => finish(true));
    this.editor = area;
    this.dom.replaceChildren(area);
    area.focus();
  }

  update(node: Node) {
    if (node.type !== this.node.type) return false;
    if (this.editor) return true;
    const changed = node.attrs.html !== this.node.attrs.html;
    this.node = node;
    if (changed) this.render();
    return true;
  }

  stopEvent(event: Event) {
    return !!this.editor && this.editor.contains(event.target as globalThis.Node);
  }

  ignoreMutation() {
    return true;
  }
}

/** Inline HTML kept as written (a picture that could not be fetched): drawn from its address. */
class HtmlInlineView implements NodeView {
  dom: HTMLElement;

  constructor(node: Node) {
    this.dom = document.createElement("span");
    this.dom.className = "vtx-html-inline";
    this.dom.contentEditable = "false";
    const html = String(node.attrs.html ?? "");
    const picture = /^!\[([^\]]*)\]\(<?([^)\s>]+)>?/.exec(html);
    if (picture) {
      const img = document.createElement("img");
      img.alt = picture[1];
      img.src = /^(https?:|data:image\/)/i.test(picture[2]) ? picture[2] : "";
      img.title = picture[2];
      this.dom.append(img);
    } else this.dom.append(safeHtml(html));
  }

  ignoreMutation() {
    return true;
  }
}

/** The node views of [toc] and HTML blocks. */
export const richBlockViews = {
  html_inline: (node: Node) => new HtmlInlineView(node),
  toc: (node: Node, view: EditorView) => new TocView(node, view),
  html_block: (node: Node, view: EditorView, getPos: () => number | undefined) => new HtmlView(node, view, getPos),
};

// ------------------------------------------------------------------- tasks --

/** A click on a task's box ticks it or clears it; every [toc] follows the headings. */
export function richBlocksPlugin(): Plugin {
  return new Plugin({
    view: () => ({
      update: (view, previous) => {
        if (view.state.doc.eq(previous.doc)) return;
        for (const toc of tocViews) toc.render();
      },
    }),
    props: {
      nodeViews: richBlockViews,
      handleDOMEvents: {
        mousedown(view, event) {
          const item = (event.target as HTMLElement).closest?.("li.vtx-task");
          if (!item || !view.editable) return false;
          // the box: drawn in the item's marker area, left of its text (its first paragraph)
          const box = (item.firstElementChild ?? item).getBoundingClientRect();
          const line = parseFloat(getComputedStyle(item).lineHeight) || 20;
          if (event.clientX >= box.left || event.clientX < box.left - line * 1.6 || event.clientY > box.top + line) return false;
          const pos = view.posAtDOM(item, 0) - 1;
          const node = view.state.doc.nodeAt(pos);
          if (!node || node.type.name !== "list_item" || node.attrs.checked === null) return false;
          event.preventDefault();
          view.dispatch(view.state.tr.setNodeAttribute(pos, "checked", !node.attrs.checked));
          return true;
        },
      },
    },
  });
}
