// The PDF's pages (DESIGN.md §12): a copy of the page view as it is on the screen — same breaks, footnotes, page
// numbers, formulas drawn by the same engine — cut into one paper-sized box per page (research: the engines' own
// page cutting drifts), scaled back to 100 %. Equation references, citations, footnote marks and internal links
// become <a href="#…"> to ids on their targets: the PDF's internal links.
import type { EditorView } from "prosemirror-view";
import { documentImageKeys, documentSettings, hasPageNumbering, imageDataUrl, imageUrl, PAPER_SIZES } from "../../shared/visualtex-document/src/index.ts";

export interface PrintCopy {
  /** the pages, for the print page's body */
  html: string;
  pages: number;
  widthMm: number;
  heightMm: number;
}

/** What an editor shows only while editing: out of the copy. */
const EDITOR_ONLY = [
  ".vtx-margin-guide",
  ".vtx-margin-line",
  ".vtx-image-handle",
  ".ProseMirror-gapcursor",
  ".vtx-image-bar",
  ".vtx-drop-marker",
  ".vtx-code-copy",
  ".vtx-page-command-label",
  // the code block's language menu and copy button
  ".vtx-code-header",
  // an HTML comment's label; an empty [toc]'s placeholder is its ::before (left as it is)
  ".vtx-html-comment",
  ".vtx-board-handle",
  ".vtx-board-guide",
].join(",");

/** The child indices from `root` down to `el`. */
function pathOf(el: Element, root: Element): number[] | null {
  const path: number[] = [];
  for (let at: Element | null = el; at && at !== root; at = at.parentElement) {
    const parent: Element | null = at.parentElement;
    if (!parent) return null;
    path.unshift(Array.prototype.indexOf.call(parent.children, at));
  }
  return path;
}

function atPath(root: Element, path: number[]): Element | null {
  let at: Element | null = root;
  for (const index of path) at = at?.children[index] ?? null;
  return at;
}

/**
 * The print copy of the editor's sheet. `gap` is the screen gap between pages (px). The editor must show every
 * formula closed (static markup): the caller closes an open one first.
 */
export function buildPrintCopy(view: EditorView, gap: number): PrintCopy | null {
  const sheet = view.dom.closest<HTMLElement>(".vtx-sheet");
  if (!sheet) return null;
  const doc = view.state.doc;
  const settings = documentSettings(doc);
  const latexNumbers = hasPageNumbering(doc);
  const paper = PAPER_SIZES[settings.paper];
  const sheetRect = sheet.getBoundingClientRect();
  const pxPerMm = sheetRect.width / paper.width;
  const zoom = pxPerMm / (96 / 25.4);
  const pageHeight = paper.height * pxPerMm;
  const stride = pageHeight + gap;
  const pages = Math.max(1, Math.round((sheet.offsetHeight + gap) / stride));

  // the blocks' places on the sheet, to leave out of a page what is not on it
  const top = (el: Element) => el.getBoundingClientRect().top - sheetRect.top;
  const editorPath = pathOf(view.dom, sheet);
  const blocks = Array.from(view.dom.children) as HTMLElement[];
  const places = blocks.map((el) => ({ top: top(el), bottom: el.getBoundingClientRect().bottom - sheetRect.top, height: el.offsetHeight, style: getComputedStyle(el) }));

  // image URLs of this window → data: URLs (the print page is another document)
  const urls = new Map<string, string>();
  for (const key of documentImageKeys(doc)) urls.set(imageUrl(key), imageDataUrl(key));

  // footnote numbers at the feet: their footnote's id (for the two-way links)
  const footnoteIds = new Map<string, string>();
  sheet.querySelectorAll("[data-footnote-pos]").forEach((el) => {
    const node = doc.nodeAt(Number(el.getAttribute("data-footnote-pos")));
    if (node?.attrs.id) footnoteIds.set(el.getAttribute("data-footnote-pos")!, node.attrs.id);
  });
  const targets = new Set<string>();
  view.dom.querySelectorAll("[data-id]").forEach((el) => targets.add(el.getAttribute("data-id")!));
  // what node views draw without their ids: references, citations, display formulas (found again in each copy by path)
  const tagged: Array<{ path: number[]; name: string; value: string }> = [];
  doc.descendants((node, pos) => {
    const name = node.type.name;
    const value =
      name === "eqref" ? node.attrs.target : name === "cite" ? (node.attrs.targets as string[]).join(",") : name === "math_display" ? node.attrs.id : null;
    if (!value) return true;
    const dom = view.nodeDOM(pos);
    const path = dom instanceof Element ? pathOf(dom, sheet) : null;
    if (path) tagged.push({ path, name: name === "eqref" ? "data-target" : name === "cite" ? "data-targets" : "data-id", value });
    if (name === "math_display") targets.add(value);
    return true;
  });

  const used = new Set<string>();
  let html = "";
  for (let page = 0; page < pages; page++) {
    const from = page * stride;
    const to = from + pageHeight;
    const copy = sheet.cloneNode(true) as HTMLElement;
    for (const tag of tagged) atPath(copy, tag.path)?.setAttribute(tag.name, tag.value);
    // the editor's own element (found by its place in the sheet: the sheet holds other elements before it)
    const editor = (editorPath ? atPath(copy, editorPath) : null) as HTMLElement | null;
    if (editor) {
      const children = Array.from(editor.children) as HTMLElement[];
      children.forEach((child, index) => {
        const place = places[index];
        if (!place || (place.bottom > from - 1 && place.top < to + 1)) return;
        // a paragraph with a wrapped image: its float reaches past the paragraph and pushes what clears it down, so
        // it stays (unseen) as it is
        if (child.querySelector(".is-wrap-left, .is-wrap-right")) {
          child.style.visibility = "hidden";
          child.setAttribute("data-print-hidden", "");
          child.removeAttribute("data-id");
          return;
        }
        // off this page: an empty, unseen box of its kind (headings keep counting) with its height and margins keeps
        // everything else where it is
        const stub = child.cloneNode(false) as HTMLElement;
        stub.removeAttribute("data-id");
        stub.removeAttribute("id");
        stub.setAttribute("data-print-hidden", "");
        stub.style.visibility = "hidden";
        stub.style.boxSizing = "border-box";
        stub.style.height = `${place.height}px`;
        stub.style.minHeight = "0";
        stub.style.overflow = "hidden";
        stub.style.marginTop = place.style.marginTop;
        stub.style.marginBottom = place.style.marginBottom;
        stub.style.clear = place.style.clear;
        child.replaceWith(stub);
      });
      editor.removeAttribute("contenteditable");
    }
    copy.querySelectorAll(EDITOR_ONLY).forEach((el) => el.remove());
    // the page numbers as LaTeX prints them (the editor's "page n of m" is the editor's), unless page commands set them
    if (!latexNumbers) copy.querySelectorAll(".vtx-page-number").forEach((el, index) => (el.textContent = String(index + 1)));
    copy.querySelectorAll(".is-selected, .ProseMirror-selectednode").forEach((el) => el.classList.remove("is-selected", "ProseMirror-selectednode"));
    copy.querySelectorAll("[contenteditable]").forEach((el) => el.removeAttribute("contenteditable"));
    copy.querySelectorAll("img").forEach((img) => {
      const data = urls.get(img.getAttribute("src") ?? "");
      if (data) img.setAttribute("src", data);
    });
    linkTargets(copy, targets, footnoteIds, used);
    copy.style.boxShadow = "none";
    copy.style.margin = "0";
    html +=
      // a hair under the paper: rounding must not spill a box onto a page of its own
      `<section class="vtx-print-page" style="width:${paper.width}mm;height:calc(${paper.height}mm - 1px)">` +
      `<div class="vtx-print-window" style="width:${sheetRect.width}px;transform:scale(${1 / zoom});top:${-from / zoom}px">` +
      copy.outerHTML +
      "</div></section>";
  }
  return { html, pages, widthMm: paper.width, heightMm: paper.height };
}

/** Ids on the targets and <a> on the references, each id once in the whole copy. */
function linkTargets(root: HTMLElement, targets: Set<string>, footnoteIds: Map<string, string>, used: Set<string>) {
  const setId = (el: Element, id: string) => {
    // an unseen block (kept for a wrapped image's float) is not a target: the page it is seen on has it
    if (used.has(id) || el.closest("[data-print-hidden]")) return;
    used.add(id);
    el.id = id;
  };
  const wrap = (el: Element, href: string) => {
    if (el.closest("a")) return;
    const a = document.createElement("a");
    a.href = href;
    a.className = "vtx-print-link";
    while (el.firstChild) a.append(el.firstChild);
    el.append(a);
  };
  root.querySelectorAll("[data-id]").forEach((el) => {
    if (el.classList.contains("vtx-footnote")) setId(el, `fnref-${el.getAttribute("data-id")}`);
    else setId(el, `vtx-${el.getAttribute("data-id")}`);
  });
  root.querySelectorAll(".vtx-eqref[data-target]").forEach((el) => {
    const target = el.getAttribute("data-target");
    if (target) wrap(el, `#vtx-${target}`);
  });
  root.querySelectorAll(".vtx-cite[data-targets]").forEach((el) => {
    const first = (el.getAttribute("data-targets") ?? "").split(",").find((id) => targets.has(id));
    if (first) wrap(el, `#vtx-${first}`);
  });
  root.querySelectorAll(".vtx-footnote[data-id]").forEach((el) => wrap(el, `#fn-${el.getAttribute("data-id")}`));
  root.querySelectorAll("[data-footnote-pos]").forEach((el) => {
    const id = footnoteIds.get(el.getAttribute("data-footnote-pos")!);
    if (!id) return;
    setId(el, `fn-${id}`);
    const number = el.querySelector("sup");
    if (number) wrap(number, `#fnref-${id}`);
  });
  root.querySelectorAll("a.vtx-link[href^='#']").forEach((a) => {
    const id = a.getAttribute("href")!.slice(1);
    if (targets.has(id)) a.setAttribute("href", `#vtx-${id}`);
  });
}
