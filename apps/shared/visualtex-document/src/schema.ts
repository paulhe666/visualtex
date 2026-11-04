// The VisualTeX document model (docs/document-editor/DESIGN.md §2, MATHCHA-SPEC.md): one ProseMirror
// document. Prose is text; formulas are atoms that store their LaTeX and nothing else (the kernel and the
// engine are views of it). Sections are Mathcha's: three numbered levels.
import { Schema, type DOMOutputSpec, type MarkSpec, type NodeSpec } from "prosemirror-model";
import { tableNodes } from "prosemirror-tables";
import { DOCUMENT_SETTING_ATTRS } from "./settings.ts";
import { noteNodes } from "./notes.ts";
import { imageNodeSpec } from "./images.ts";

/** Every block carries a stable id (labels, export mapping, compile errors). */
const id = { id: { default: null as string | null } };

/** Paragraph-level style: alignment (left/center/right/justify) and indent steps (Mathcha's Indent button). */
const paragraphStyle = { align: { default: null as string | null }, indent: { default: 0 } };

const idFrom = (dom: HTMLElement | string) =>
  typeof dom === "string" ? {} : { id: dom.getAttribute("data-id") };

const withId = (attrs: { id?: string | null }) => (attrs.id ? { "data-id": attrs.id } : {});

function styleOf(attrs: { align?: string | null; indent?: number }) {
  const style: string[] = [];
  if (attrs.align) style.push(`text-align:${attrs.align}`);
  if (attrs.indent) style.push(`--vtx-indent:${attrs.indent}`);
  // data-align: for the style sheet (the browser rewrites the style attribute: "text-align: center;")
  return style.length
    ? { style: style.join(";"), "data-indent": attrs.indent ? String(attrs.indent) : undefined, "data-align": attrs.align ?? undefined }
    : {};
}

function paragraphStyleFrom(dom: HTMLElement | string) {
  if (typeof dom === "string") return {};
  const align = dom.style.textAlign;
  const indent = Number(dom.getAttribute("data-indent") ?? 0) || 0;
  return { align: ["left", "center", "right", "justify"].includes(align) ? align : null, indent };
}

const nodes: Record<string, NodeSpec> = {
  /** The page and body text settings are the document's attributes (settings.ts). */
  doc: { content: "block+", attrs: DOCUMENT_SETTING_ATTRS },

  paragraph: {
    group: "block",
    content: "inline*",
    attrs: { ...id, ...paragraphStyle },
    parseDOM: [{ tag: "p", getAttrs: (dom) => ({ ...idFrom(dom), ...paragraphStyleFrom(dom) }) }],
    toDOM: (node): DOMOutputSpec => ["p", { ...withId(node.attrs), ...styleOf(node.attrs) }, 0],
  },

  /** A Mathcha section: levels 1-3, numbered 1. / 1.1. / 1.1.1. by the document. */
  heading: {
    group: "block",
    content: "inline*",
    defining: true,
    attrs: { ...id, level: { default: 1 }, ...paragraphStyle },
    parseDOM: [1, 2, 3, 4, 5, 6].map((level) => ({
      tag: "h" + level,
      getAttrs: (dom: HTMLElement | string) => ({ ...idFrom(dom), ...paragraphStyleFrom(dom), level: Math.min(3, level) }),
    })),
    toDOM: (node): DOMOutputSpec => ["h" + node.attrs.level, { ...withId(node.attrs), ...styleOf(node.attrs) }, 0],
  },

  /** The document's title (LaTeX's \title + \maketitle): centred, unnumbered, at the top. */
  title: {
    group: "block",
    content: "inline*",
    defining: true,
    attrs: id,
    parseDOM: [{ tag: "p.vtx-title", priority: 60, getAttrs: idFrom }],
    toDOM: (node): DOMOutputSpec => ["p", { ...withId(node.attrs), class: "vtx-title" }, 0],
  },

  blockquote: {
    group: "block",
    content: "block+",
    defining: true,
    attrs: id,
    parseDOM: [{ tag: "blockquote", getAttrs: idFrom }],
    toDOM: (node): DOMOutputSpec => ["blockquote", withId(node.attrs), 0],
  },

  bullet_list: {
    group: "block",
    content: "list_item+",
    attrs: id,
    parseDOM: [{ tag: "ul", getAttrs: idFrom }],
    toDOM: (node): DOMOutputSpec => ["ul", withId(node.attrs), 0],
  },

  ordered_list: {
    group: "block",
    content: "list_item+",
    attrs: { ...id, start: { default: 1 } },
    parseDOM: [
      {
        tag: "ol",
        getAttrs: (dom: HTMLElement | string) =>
          typeof dom === "string" ? {} : { ...idFrom(dom), start: Number(dom.getAttribute("start") ?? 1) || 1 },
      },
    ],
    toDOM: (node): DOMOutputSpec => [
      "ol",
      { ...withId(node.attrs), ...(node.attrs.start !== 1 ? { start: node.attrs.start } : {}) },
      0,
    ],
  },

  list_item: {
    content: "paragraph block*",
    defining: true,
    attrs: id,
    parseDOM: [{ tag: "li", getAttrs: idFrom }],
    toDOM: (node): DOMOutputSpec => ["li", withId(node.attrs), 0],
  },

  /** A code block: plain text in a language (highlighted by the app, as VS Code highlights it); no marks. */
  code_block: {
    group: "block",
    content: "text*",
    marks: "",
    code: true,
    defining: true,
    attrs: { ...id, language: { default: "" } },
    parseDOM: [
      {
        tag: "pre",
        preserveWhitespace: "full",
        getAttrs: (dom: HTMLElement | string) =>
          typeof dom === "string"
            ? {}
            : {
                ...idFrom(dom),
                language:
                  dom.getAttribute("data-language") ??
                  /(?:^|\s)language-([\w+#-]+)/.exec(dom.querySelector("code")?.className ?? "")?.[1] ??
                  "",
              },
      },
    ],
    toDOM: (node): DOMOutputSpec => [
      "pre",
      { ...withId(node.attrs), ...(node.attrs.language ? { "data-language": node.attrs.language } : {}) },
      ["code", 0],
    ],
  },

  /** A display formula (Mathcha's \math-container): its own line, centred; numbered only on request. */
  math_display: {
    group: "block",
    atom: true,
    selectable: true,
    draggable: false,
    attrs: { ...id, latex: { default: "" }, numbered: { default: false } },
    parseDOM: [
      {
        tag: "div.vtx-math-display",
        getAttrs: (dom: HTMLElement | string) =>
          typeof dom === "string"
            ? {}
            : {
                ...idFrom(dom),
                latex: dom.getAttribute("data-latex") ?? "",
                numbered: dom.getAttribute("data-numbered") === "true",
              },
      },
    ],
    toDOM: (node): DOMOutputSpec => [
      "div",
      {
        class: "vtx-math-display",
        ...withId(node.attrs),
        "data-latex": node.attrs.latex,
        "data-numbered": String(node.attrs.numbered),
      },
    ],
  },

  horizontal_rule: {
    group: "block",
    attrs: id,
    parseDOM: [{ tag: "hr", getAttrs: idFrom }],
    toDOM: (node): DOMOutputSpec => ["hr", withId(node.attrs)],
  },

  text: { group: "inline" },

  /** An inline formula: one unit in the text (the caret steps over it or into it). */
  math_inline: {
    group: "inline",
    inline: true,
    atom: true,
    selectable: true,
    attrs: { id: { default: null as string | null }, latex: { default: "" } },
    parseDOM: [
      {
        tag: "span.vtx-math-inline",
        getAttrs: (dom: HTMLElement | string) =>
          typeof dom === "string" ? {} : { latex: dom.getAttribute("data-latex") ?? "" },
      },
    ],
    toDOM: (node): DOMOutputSpec => ["span", { class: "vtx-math-inline", "data-latex": node.attrs.latex }],
  },

  /** An image (images.ts): a box in the line, as \includegraphics is. */
  image: imageNodeSpec,

  /** A reference to a numbered display formula: shows its number, exported as \eqref. */
  eqref: {
    group: "inline",
    inline: true,
    atom: true,
    selectable: true,
    attrs: { target: { default: null as string | null } },
    parseDOM: [
      {
        tag: "span.vtx-eqref",
        getAttrs: (dom: HTMLElement | string) => (typeof dom === "string" ? {} : { target: dom.getAttribute("data-target") }),
      },
    ],
    toDOM: (node): DOMOutputSpec => ["span", { class: "vtx-eqref", "data-target": node.attrs.target ?? "" }, "(?)"],
  },

  hard_break: {
    group: "inline",
    inline: true,
    selectable: false,
    parseDOM: [{ tag: "br" }],
    toDOM: (): DOMOutputSpec => ["br"],
  },

  ...noteNodes,

  ...tableNodes({
    tableGroup: "block",
    cellContent: "(paragraph | math_display)+",
    cellAttributes: {
      align: {
        default: null,
        getFromDOM: (dom) => dom.style.textAlign || null,
        setDOMAttr: (value, attrs) => {
          if (value) attrs.style = (attrs.style ?? "") + `text-align:${value};`;
        },
      },
    },
  }),
};

// Tables are blocks with an id too (\label{tab:...}).
nodes.table = {
  ...nodes.table,
  attrs: { ...(nodes.table.attrs ?? {}), ...id },
  parseDOM: [{ tag: "table", getAttrs: idFrom }],
  toDOM: (node): DOMOutputSpec => ["table", withId(node.attrs), ["tbody", 0]],
};

/** Mathcha's font sizes (LaTeX's names) as ratios of the document's text size. */
export const FONT_SIZES: Array<{ name: string; ratio: number }> = [
  { name: "tiny", ratio: 0.5 },
  { name: "scriptsize", ratio: 0.7 },
  { name: "footnotesize", ratio: 0.8 },
  { name: "small", ratio: 0.9 },
  { name: "normalsize", ratio: 1 },
  { name: "large", ratio: 1.2 },
  { name: "Large", ratio: 1.44 },
  { name: "LARGE", ratio: 1.728 },
  { name: "huge", ratio: 2.074 },
  { name: "Huge", ratio: 2.488 },
];

const marks: Record<string, MarkSpec> = {
  /** A link (Markdown's [text](url)); first, so it wraps the other marks. */
  link: {
    attrs: { href: {}, title: { default: null as string | null } },
    inclusive: false,
    parseDOM: [
      {
        tag: "a[href]",
        getAttrs: (dom: HTMLElement | string) =>
          typeof dom === "string" ? false : { href: dom.getAttribute("href"), title: dom.getAttribute("title") },
      },
    ],
    // The address shows on hover (the link is not followed from the editor: it is text being edited).
    toDOM: (mark): DOMOutputSpec => [
      "a",
      { href: mark.attrs.href, title: mark.attrs.title ? `${mark.attrs.title}\n${mark.attrs.href}` : mark.attrs.href, class: "vtx-link", rel: "noopener noreferrer" },
      0,
    ],
  },
  strong: {
    parseDOM: [
      { tag: "strong" },
      { tag: "b", getAttrs: (node: HTMLElement | string) => typeof node !== "string" && node.style.fontWeight !== "normal" && null },
      { style: "font-weight=bold" },
    ],
    toDOM: (): DOMOutputSpec => ["strong", 0],
  },
  em: {
    parseDOM: [{ tag: "i" }, { tag: "em" }, { style: "font-style=italic" }],
    toDOM: (): DOMOutputSpec => ["em", 0],
  },
  underline: {
    parseDOM: [{ tag: "u" }, { style: "text-decoration=underline" }],
    toDOM: (): DOMOutputSpec => ["u", 0],
  },
  strike: {
    parseDOM: [{ tag: "s" }, { tag: "del" }, { style: "text-decoration=line-through" }],
    toDOM: (): DOMOutputSpec => ["s", 0],
  },
  code: {
    excludes: "_",
    parseDOM: [{ tag: "code" }],
    toDOM: (): DOMOutputSpec => ["code", 0],
  },
  textColor: {
    attrs: { color: {} },
    parseDOM: [
      {
        tag: "span[data-color]",
        getAttrs: (dom: HTMLElement | string) => (typeof dom === "string" ? false : { color: dom.getAttribute("data-color") }),
      },
    ],
    toDOM: (mark): DOMOutputSpec => ["span", { "data-color": mark.attrs.color, style: `color:${mark.attrs.color}` }, 0],
  },
  highlight: {
    attrs: { color: {} },
    parseDOM: [
      {
        tag: "span[data-highlight]",
        getAttrs: (dom: HTMLElement | string) => (typeof dom === "string" ? false : { color: dom.getAttribute("data-highlight") }),
      },
    ],
    toDOM: (mark): DOMOutputSpec => [
      "span",
      { "data-highlight": mark.attrs.color, style: `background-color:${mark.attrs.color}` },
      0,
    ],
  },
  fontSize: {
    attrs: { size: {} },
    parseDOM: [
      {
        tag: "span[data-size]",
        getAttrs: (dom: HTMLElement | string) => (typeof dom === "string" ? false : { size: dom.getAttribute("data-size") }),
      },
    ],
    toDOM: (mark): DOMOutputSpec => {
      const ratio = FONT_SIZES.find((s) => s.name === mark.attrs.size)?.ratio ?? 1;
      return ["span", { "data-size": mark.attrs.size, style: `font-size:${ratio}em` }, 0];
    },
  },
};

export const documentSchema = new Schema({ nodes, marks });

export type DocumentSchema = typeof documentSchema;

/** Node types that are formulas. */
export function isFormulaNodeName(name: string): name is "math_inline" | "math_display" {
  return name === "math_inline" || name === "math_display";
}
