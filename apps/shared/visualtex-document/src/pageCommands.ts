// LaTeX's page commands in the document (DESIGN.md §8.7): \newpage, \clearpage, \pagenumbering{…},
// \pagestyle{…}, \thispagestyle{…}, \setcounter{page}{…}. Each is a block where it was typed or read, as in the
// source; the page view breaks the page at \newpage / \clearpage and numbers the pages as LaTeX does (the style
// and count change on the page the command is on), and every export writes it (LaTeX as it is, Word as a page
// break or a section's page numbering, Markdown as a page break or a comment).
import type { DOMOutputSpec, Node, NodeSpec } from "prosemirror-model";

export type PageCommandName = "newpage" | "clearpage" | "pagenumbering" | "pagestyle" | "thispagestyle" | "setcounter";

export const PAGE_NUMBER_STYLES = ["arabic", "roman", "Roman", "alph", "Alph", "gobble"] as const;
export type PageNumberStyle = (typeof PAGE_NUMBER_STYLES)[number];
export const PAGE_STYLES = ["plain", "empty", "headings", "myheadings"] as const;

/** The command as LaTeX writes it. */
export function pageCommandLatex(command: string, value: string): string {
  switch (command) {
    case "newpage":
    case "clearpage":
      return `\\${command}`;
    case "setcounter":
      return `\\setcounter{page}{${value || "1"}}`;
    default:
      return `\\${command}{${value}}`;
  }
}

/** A line of LaTeX that is one page command (null: it is not). */
export function parsePageCommand(text: string): { command: PageCommandName; value: string } | null {
  const line = text.trim();
  if (/^\\(newpage|clearpage|cleardoublepage|pagebreak)(\[\d\])?$/.test(line)) {
    return { command: /^\\newpage|^\\pagebreak/.test(line) ? "newpage" : "clearpage", value: "" };
  }
  const numbering = /^\\pagenumbering\s*\{\s*(arabic|roman|Roman|alph|Alph|gobble)\s*\}$/.exec(line);
  if (numbering) return { command: "pagenumbering", value: numbering[1] };
  const style = /^\\(pagestyle|thispagestyle)\s*\{\s*([A-Za-z]+)\s*\}$/.exec(line);
  if (style) return { command: style[1] as PageCommandName, value: style[2] };
  const counter = /^\\setcounter\s*\{\s*page\s*\}\s*\{\s*(-?\d+)\s*\}$/.exec(line);
  if (counter) return { command: "setcounter", value: counter[1] };
  return null;
}

/** Any page command in a text (LaTeX source detection). */
export const PAGE_COMMAND_PATTERN = /\\(?:newpage|clearpage|cleardoublepage|pagenumbering|pagestyle|thispagestyle)\b|\\setcounter\s*\{\s*page\s*\}/;

export const pageCommandNodeSpec: NodeSpec = {
  group: "block",
  atom: true,
  selectable: true,
  draggable: false,
  attrs: { id: { default: null as string | null }, command: { default: "newpage" }, value: { default: "" } },
  parseDOM: [
    {
      tag: "div.vtx-page-command",
      getAttrs: (dom: HTMLElement | string) =>
        typeof dom === "string" ? {} : { command: dom.getAttribute("data-command") ?? "newpage", value: dom.getAttribute("data-value") ?? "" },
    },
  ],
  // no height of its own (LaTeX's commands take none): the command shows in the left margin, a page break as a
  // dashed rule across
  toDOM: (node): DOMOutputSpec => [
    "div",
    {
      class: "vtx-page-command",
      "data-command": node.attrs.command,
      "data-value": node.attrs.value,
      contenteditable: "false",
    },
    ["span", { class: "vtx-page-command-label", title: pageCommandLatex(node.attrs.command, node.attrs.value) }, pageCommandLatex(node.attrs.command, node.attrs.value)],
  ],
};

export const isPageCommand = (node: Node | null | undefined) => node?.type.name === "page_command";
export const isPageBreak = (node: Node | null | undefined) => isPageCommand(node) && (node!.attrs.command === "newpage" || node!.attrs.command === "clearpage");

function roman(n: number): string {
  if (n <= 0 || n >= 4000) return String(n);
  const table: Array<[number, string]> = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
  let out = "";
  for (const [value, letters] of table) {
    while (n >= value) {
      out += letters;
      n -= value;
    }
  }
  return out;
}

/** A page number in a LaTeX numbering style. */
export function formatPageNumber(n: number, style: string): string {
  switch (style) {
    case "roman":
      return roman(n);
    case "Roman":
      return roman(n).toUpperCase();
    case "alph":
      return n >= 1 && n <= 26 ? String.fromCharCode(96 + n) : String(n);
    case "Alph":
      return n >= 1 && n <= 26 ? String.fromCharCode(64 + n) : String(n);
    case "gobble":
      return "";
    default:
      return String(n);
  }
}

/** Whether the document has page-numbering commands (then its pages are numbered as LaTeX numbers them). */
export function hasPageNumbering(doc: Node): boolean {
  let found = false;
  doc.forEach((node) => {
    if (isPageCommand(node) && !isPageBreak(node)) found = true;
  });
  return found;
}

export interface PageLabel {
  /** the number as printed ("" on an empty page style or \pagenumbering{gobble}) */
  text: string;
  style: string;
  number: number;
  /** the page's style: plain (article's and ctexart's default), empty, headings, myheadings */
  pageStyle: string;
}

/**
 * The pages' numbers, LaTeX's way: `commands[p]` are the commands on page p, in order. \pagenumbering resets the
 * count to 1 and sets the style from its page on; \setcounter{page}{n} makes its page n; \pagestyle{…} sets the page
 * style from its page on, \thispagestyle{…} on its page only (empty: no number).
 */
export function pageLabels(pages: number, commands: Array<Array<{ command: string; value: string }>>): PageLabel[] {
  const labels: PageLabel[] = [];
  let style = "arabic";
  let count = 1;
  let pageStyle = "plain";
  const known = (value: string) => (PAGE_STYLES as readonly string[]).includes(value);
  for (let page = 0; page < pages; page++) {
    let thisStyle: string | null = null;
    for (const { command, value } of commands[page] ?? []) {
      if (command === "pagenumbering") {
        style = value || "arabic";
        count = 1;
      } else if (command === "setcounter") count = Number(value) || 0;
      else if (command === "pagestyle" && known(value)) pageStyle = value;
      else if (command === "thispagestyle" && known(value)) thisStyle = value;
    }
    const current = thisStyle ?? pageStyle;
    labels.push({ text: current !== "empty" ? formatPageNumber(count, style) : "", style, number: count, pageStyle: current });
    count++;
  }
  return labels;
}
