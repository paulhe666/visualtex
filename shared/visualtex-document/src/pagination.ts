// The page view's pages (DESIGN.md §8.2): the document is laid out on one sheet as wide as the paper's text area
// (settings.ts); this plugin measures where the paper's pages end and pushes what follows a page's end down to the
// next page's text area with a spacer (a widget decoration: between blocks, or inside a paragraph at the start of
// the line that no longer fits). A page's footnotes take room at its foot; the page layer, an overlay on the sheet,
// prints them there with the page numbers. The sheet's background draws the pages and the gaps between them.
import type { Node } from "prosemirror-model";
import { Plugin, PluginKey, type EditorState } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { bodyLayout, documentSettings, headingNumbers, PAPER_SIZES, type DocumentSettings, type LatexClass } from "./settings.ts";
import { requestFootnoteEdit } from "./notes.ts";
import { isPageBreak, pageLabels, type PageLabel } from "./pageCommands.ts";
import { CJK_TEXT } from "./titleBlock.ts";
import { shownContent } from "./sourceMarkdown.ts";

export interface PageGeometry {
  /** screen pixels per millimetre (the zoom included) */
  pxPerMm: number;
  /** the gap between two pages on the screen, px */
  gap: number;
}

export interface PaginationOptions {
  geometry: () => PageGeometry;
  /** A footnote's text as HTML (its $…$ formulas drawn); default: the text escaped. */
  renderNote?: (text: string) => string;
}

interface PageBreak {
  pos: number;
  /** inside a paragraph (at the start of a line) or between blocks */
  inline: boolean;
  height: number;
  /** a break before a textblock's first line: a transparent top border on the block (its heading number and first-line indent
   *  moves with its text), at this block position */
  block?: number;
  /** a break before a row of a long table (figures.ts): a row of space, as wide as the table's columns */
  row?: number;
}

interface PaginationState {
  decorations: DecorationSet;
  breaks: PageBreak[];
}

export const paginationKey = new PluginKey<PaginationState>("visualtex-document-pagination");

/** A line of text, or a block that cannot be divided (a display formula, a table, a rule). */
interface Unit {
  top: number;
  bottom: number;
  /** where a page break before it goes */
  pos: number;
  inline: boolean;
  /** the last line of a heading: it goes to the next page with what follows it */
  keepWithNext: boolean;
  /** a textblock's first line: the block's position */
  block?: number;
  footnotes: number[];
  /** \newpage / \clearpage: what follows starts a new page */
  breakAfter?: boolean;
  /** a page command (\pagenumbering…): the page it lands on numbers by it */
  command?: { command: string; value: string };
  /** a long table's row (its table's columns) */
  row?: number;
  /** a section's first line: its mark for the running head (\sectionmark), as HTML */
  mark?: string;
}

const SEPARATOR_HEIGHT_EM = 0.9;

function escapeHtml(text: string): string {
  return text.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!);
}

/** Text for the running head, in \slshape: Chinese upright (its font has no slanted face), the rest slanted. */
function slanted(text: string): string {
  return text.replace(new RegExp(`(${CJK_TEXT.source})+`, "gu"), (run) => `<span class="is-upright">${run}</span>`);
}

/**
 * A section's mark as the class's \sectionmark makes it for \pagestyle{headings} (one-sided): its number, \quad, its
 * title in capitals (\MakeUppercase); its formulas as drawn. With no section numbers article still prints the
 * counter (0, as LaTeX does: secnumdepth is not tested there), ctexart none.
 */
function sectionMark(view: EditorView, node: Node, pos: number, number: string | undefined, documentClass: LatexClass): string {
  let title = "";
  // (the heading as it reads: its "#" and markers left out)
  for (const piece of shownContent(node)) {
    if ("text" in piece) title += slanted(escapeHtml(piece.text.toUpperCase()));
    else if (piece.node.isInline) {
      const dom = view.nodeDOM(pos + 1 + piece.offset);
      if (dom instanceof HTMLElement && dom.classList.contains("vtx-math-inline")) title += dom.outerHTML;
    }
  }
  const shown = number?.replace(/\.$/, "") ?? (documentClass === "article" ? "0" : "");
  return shown ? `${slanted(shown)}<span class="vtx-page-quad"></span>${title}` : title;
}

function spacer(height: number, inline: boolean, row?: number): () => HTMLElement {
  if (row) {
    // between a long table's rows: a row of space (document.css: no rules through it)
    return () => {
      const tr = document.createElement("tr");
      tr.className = "vtx-page-break";
      tr.contentEditable = "false";
      tr.setAttribute("aria-hidden", "true");
      const td = document.createElement("td");
      td.colSpan = row;
      td.style.height = `${height}px`;
      tr.append(td);
      return tr;
    };
  }
  return () => {
    const el = document.createElement(inline ? "span" : "div");
    el.className = "vtx-page-break";
    el.contentEditable = "false";
    el.setAttribute("aria-hidden", "true");
    el.style.display = "block";
    el.style.height = `${height}px`;
    return el;
  };
}

function buildDecorations(doc: Node, breaks: PageBreak[]): DecorationSet {
  return DecorationSet.create(
    doc,
    breaks.map((b) => {
      const block = b.block !== undefined ? doc.nodeAt(b.block) : null;
      if (block?.isTextblock || block?.type.name === "abstract") {
        return Decoration.node(b.block!, b.block! + block.nodeSize, {
          class: "vtx-page-break-before",
          style: `border-top:${b.height}px solid transparent`,
          "data-page-break": String(b.height),
        });
      }
      return Decoration.widget(b.pos, spacer(b.height, b.inline, b.row), {
        side: -1,
        ignoreSelection: true,
        key: `page-break-${b.pos}-${b.inline ? "i" : "b"}${b.row ? "r" : ""}-${Math.round(b.height * 10)}`,
      });
    }),
  );
}

const sameBreaks = (a: PageBreak[], b: PageBreak[]) =>
  a.length === b.length && a.every((x, i) => x.pos === b[i].pos && x.inline === b[i].inline && x.block === b[i].block && Math.abs(x.height - b[i].height) < 0.5);

/** The lines of a textblock's DOM, from its text and inline boxes (top/bottom in client coordinates). */
export function lineBoxes(dom: HTMLElement): Array<{ top: number; bottom: number; left: number }> {
  const rects: DOMRect[] = [];
  const range = document.createRange();
  for (const child of Array.from(dom.childNodes)) {
    if (child.nodeType === 3) {
      range.selectNodeContents(child);
      for (const r of Array.from(range.getClientRects())) if (r.height > 0) rects.push(r);
    } else if (child instanceof HTMLElement) {
      if (child.classList.contains("vtx-page-break")) continue;
      // a wrapped image is beside the lines, not one of them (its height goes with the first line: collectUnits)
      if (child.classList.contains("is-wrap-left") || child.classList.contains("is-wrap-right")) continue;
      // a mark's span holds text; an inline formula or a footnote is one box
      if (getComputedStyle(child).display === "inline") {
        range.selectNodeContents(child);
        for (const r of Array.from(range.getClientRects())) if (r.height > 0) rects.push(r);
      } else {
        const r = child.getBoundingClientRect();
        if (r.height > 0) rects.push(r);
      }
    }
  }
  rects.sort((a, b) => a.top - b.top || a.left - b.left);
  const lines: Array<{ top: number; bottom: number; left: number }> = [];
  for (const r of rects) {
    const line = lines.find((l) => Math.min(l.bottom, r.bottom) - Math.max(l.top, r.top) > 0.5 * Math.min(l.bottom - l.top, r.height));
    if (line) {
      line.top = Math.min(line.top, r.top);
      line.bottom = Math.max(line.bottom, r.bottom);
      line.left = Math.min(line.left, r.left);
    } else lines.push({ top: r.top, bottom: r.bottom, left: r.left });
  }
  return lines.sort((a, b) => a.top - b.top);
}

/** The position at the start of a line (its first character or formula). */
function lineStart(view: EditorView, start: number, end: number, line: { top: number; bottom: number; left: number }): number | null {
  const hit = view.posAtCoords({ left: line.left + 1, top: (line.top + line.bottom) / 2 });
  if (!hit) return null;
  let pos = hit.inside >= start && hit.inside < end && view.state.doc.nodeAt(hit.inside)?.isInline ? hit.inside : hit.pos;
  pos = Math.max(start, Math.min(end, pos));
  // the place must be on this line: not the end of the line before
  for (let i = 0; i < 3 && pos < end; i++) {
    const at = view.coordsAtPos(pos, 1);
    if (at.top >= line.top - 2) break;
    pos++;
  }
  return pos;
}

/** The block at `pos` comes first in a table float with a caption (figures.ts): a caption above its table, or a
 *  table with its caption below it. It goes to a page with what follows it. */
function firstOfCaptionedTable(doc: Node, pos: number): boolean {
  const $pos = doc.resolve(pos);
  return $pos.parent.type.name === "table_figure" && $pos.index() === 0 && $pos.parent.childCount > 1;
}

/** What the page view divides into pages, in document order, at their places on the screen. */
/** Lines that go to the next page with what follows them: a section's, and the title block's (\maketitle's box). */
const KEEP_WITH_NEXT = new Set(["heading", "title", "author", "date"]);

function collectUnits(view: EditorView, pageHeight = Infinity): Unit[] {
  const units: Unit[] = [];
  /** `container`: the abstract whose first line this block holds (its heading goes along: the break is the abstract's) */
  const visit = (node: Node, pos: number, container?: { pos: number; top: number }) => {
    const dom = view.nodeDOM(pos);
    if (!(dom instanceof HTMLElement)) return;
    if (node.isTextblock && node.type.name !== "code_block") {
      const keep = KEEP_WITH_NEXT.has(node.type.name);
      const lines = lineBoxes(dom);
      if (!lines.length) {
        const r = dom.getBoundingClientRect();
        units.push({ top: container?.top ?? r.top, bottom: r.bottom, pos: pos + 1, inline: true, keepWithNext: keep, footnotes: [], block: container?.pos ?? pos });
        return;
      }
      let previous = -1;
      // a wrapped image takes its paragraph's first line along: wrapfig puts them on the next page together
      const wrapped = dom.querySelector(":scope > .is-wrap-left, :scope > .is-wrap-right");
      const wrapBottom = wrapped ? wrapped.getBoundingClientRect().bottom : -Infinity;
      lines.forEach((line, index) => {
        const start = index === 0 ? pos + 1 : lineStart(view, pos + 1, pos + node.nodeSize - 1, line);
        if (start === null || start <= previous) return;
        previous = start;
        units.push({
          top: index === 0 ? Math.min(line.top, container?.top ?? dom.getBoundingClientRect().top) : line.top,
          bottom: index === 0 ? Math.max(line.bottom, wrapBottom) : line.bottom,
          pos: start,
          inline: true,
          // a table's caption above it (figures.ts) goes with its table, as a heading with its text
          keepWithNext: (keep || (node.type.name === "caption" && firstOfCaptionedTable(view.state.doc, pos))) && index === lines.length - 1,
          footnotes: [],
          ...(index === 0 ? { block: container?.pos ?? pos } : {}),
        });
      });
      return;
    }
    // a list, a quote, a reference list, a list item, a table with its caption: their blocks
    // (a table, an image board and a figure are one unit each)
    if (!node.isAtom && !node.isTextblock && node.type.name !== "table" && node.type.name !== "image_board" && node.type.name !== "figure" && node.childCount && node.firstChild!.isBlock) {
      const first = node.type.name === "abstract" ? { pos, top: dom.getBoundingClientRect().top } : container;
      node.forEach((child, offset) => visit(child, pos + 1 + offset, offset === 0 ? first : undefined));
      return;
    }
    const r = dom.getBoundingClientRect();
    // a table with its caption below it: the table's end goes with the caption
    const captionBelow = node.type.name === "table" && firstOfCaptionedTable(view.state.doc, pos);
    // a table longer than a page (without the rows of space its breaks put in): its rows, as longtable breaks it
    // (figures.ts)
    const breaksIn = node.type.name === "table" ? Array.from(dom.querySelectorAll("tr.vtx-page-break")).reduce((sum, tr) => sum + tr.getBoundingClientRect().height, 0) : 0;
    if (node.type.name === "table" && r.height - breaksIn > pageHeight) {
      let columns = 0;
      node.firstChild?.forEach((cell) => (columns += cell.attrs.colspan ?? 1));
      node.forEach((_row, offset, index) => {
        const rowDom = view.nodeDOM(pos + 1 + offset);
        if (!(rowDom instanceof HTMLElement)) return;
        const rr = rowDom.getBoundingClientRect();
        units.push({ top: rr.top, bottom: rr.bottom, pos: pos + 1 + offset, inline: false, keepWithNext: captionBelow && index === node.childCount - 1, footnotes: [], row: Math.max(1, columns) });
      });
      return;
    }
    // a wrapped figure goes to the next page with the first line of the paragraph beside it (wrapfig)
    const unit: Unit = { top: r.top, bottom: r.bottom, pos, inline: false, keepWithNext: captionBelow || (node.type.name === "figure" && (dom.classList.contains("is-wrap-left") || dom.classList.contains("is-wrap-right"))), footnotes: [] };
    if (node.type.name === "page_command") {
      unit.breakAfter = isPageBreak(node);
      unit.command = { command: node.attrs.command, value: node.attrs.value };
    }
    units.push(unit);
  };
  view.state.doc.forEach((node, offset) => visit(node, offset));
  return units;
}

class PaginationView {
  private frame = 0;
  private timer = 0;
  private readonly layer: HTMLElement;
  private readonly measurer: HTMLElement;
  private lastKey = "";
  /** The page breaks measured since the document or the page last changed: one of them again is not taken. */
  private seen = new Set<string>();

  constructor(
    private view: EditorView,
    private options: PaginationOptions,
  ) {
    this.layer = document.createElement("div");
    this.layer.className = "vtx-page-layer";
    this.measurer = document.createElement("div");
    this.measurer.className = "vtx-page-footnotes is-measuring";
    this.layer.append(this.measurer);
    this.layer.addEventListener("mousedown", this.onLayerMouseDown);
    this.sheet()?.append(this.layer);
    this.schedule();
    window.addEventListener("resize", this.changed);
    document.fonts?.addEventListener?.("loadingdone", this.changed);
  }

  /** The sheet the editor lies on (its pages are drawn by its background; the layer is placed on it). */
  private sheet(): HTMLElement | null {
    return this.view.dom.closest<HTMLElement>(".vtx-sheet");
  }

  private onLayerMouseDown = (event: MouseEvent) => {
    const note = (event.target as Element | null)?.closest?.("[data-footnote-pos]");
    if (!note || event.button !== 0) return;
    event.preventDefault();
    // the editor opens at the note that was clicked (its number may be far above, off the screen)
    requestFootnoteEdit(this.view, Number(note.getAttribute("data-footnote-pos")), note.getBoundingClientRect());
  };

  update(view: EditorView, previous: EditorState) {
    this.view = view;
    if (this.layer.parentElement !== this.sheet()) this.sheet()?.append(this.layer);
    if (view.state.doc !== previous.doc) this.seen.clear();
    if (view.state.doc !== previous.doc || paginationKey.getState(view.state) !== paginationKey.getState(previous)) this.schedule();
  }

  /** Measure after the browser has laid the change out (a burst of changes is measured once). */
  schedule = () => {
    window.clearTimeout(this.timer);
    cancelAnimationFrame(this.frame);
    this.timer = window.setTimeout(() => {
      this.frame = requestAnimationFrame(() => this.measure());
    }, 30);
  };

  /** Measure again now (the geometry or the fonts changed). */
  refresh() {
    this.seen.clear();
    this.schedule();
  }

  private changed = () => {
    this.seen.clear();
    this.schedule();
  };

  private measure() {
    const view = this.view;
    const sheet = this.sheet();
    if (!sheet || !view.dom.isConnected) return;
    // an input method is composing: a spacer changed next to it would end the composition (Chrome); later
    if (view.composing) {
      this.schedule();
      return;
    }
    const settings = documentSettings(view.state.doc);
    const { pxPerMm, gap } = this.options.geometry();
    if (!(pxPerMm > 0)) return;
    const paper = PAPER_SIZES[settings.paper];
    const stride = paper.height * pxPerMm + gap;
    const contentHeight = (paper.height - settings.marginTop - settings.marginBottom) * pxPerMm;
    const origin = view.dom.getBoundingClientRect().top;

    // the spacers now on the screen: what is below them is that much lower than it would be without them
    const spacers = Array.from(view.dom.querySelectorAll<HTMLElement>(".vtx-page-break, [data-page-break]")).map((el) => {
      const r = el.getBoundingClientRect();
      // a block's top border: from its top, as wide as it is drawn (the browser rounds a border to whole device pixels:
      // what it asked for would put everything below a fraction of a pixel off, and the next measurement would move
      // a line on a page's edge back and forth)
      if (el.dataset.pageBreak === undefined) return { top: r.top, height: r.height };
      const drawn = parseFloat(getComputedStyle(el).borderTopWidth);
      return { top: r.top, height: Number.isFinite(drawn) ? drawn : Number(el.dataset.pageBreak) || 0 };
    });
    const natural = (y: number) => {
      let shift = 0;
      for (const s of spacers) if (s.top + s.height / 2 <= y) shift += s.height;
      return y - origin - shift;
    };
    const units = collectUnits(view, contentHeight).map((u) => ({ ...u, top: natural(u.top), bottom: natural(u.bottom) }));
    // the sections' marks (the running head of \pagestyle{headings}) and \maketitle's \thispagestyle{plain}
    const layout = bodyLayout(view.state.doc);
    const numbers = headingNumbers(view.state.doc);
    let titled = false;
    for (const unit of units) {
      const node = unit.block !== undefined ? view.state.doc.nodeAt(unit.block) : null;
      if (node?.type.name === "heading" && node.attrs.level === 1 && !node.attrs.unnumbered) {
        unit.mark = sectionMark(view, node, unit.block!, numbers.get(unit.block!), layout.documentClass);
      } else if (node?.type.name === "title" && node.textContent.trim() && !titled && !unit.command) {
        titled = true;
        unit.command = { command: "thispagestyle", value: "plain" };
      }
    }

    // footnotes: the line each one's mark is on, and the room its note takes at a page's foot
    const notes: Array<{ pos: number; text: string; number: number }> = [];
    view.state.doc.descendants((node, pos) => {
      if (node.type.name === "footnote") notes.push({ pos, text: node.attrs.text, number: notes.length + 1 });
      return true;
    });
    const noteHeights = this.measureNotes(notes, settings, pxPerMm);
    const separator = SEPARATOR_HEIGHT_EM * settings.fontSize * (96 / 72) * (pxPerMm / (96 / 25.4));
    notes.forEach((note, index) => {
      const dom = view.nodeDOM(note.pos);
      if (!(dom instanceof Element)) return;
      const r = dom.getBoundingClientRect();
      const y = natural((r.top + r.bottom) / 2);
      const unit = units.find((u) => y >= u.top - 1 && y <= u.bottom + 1) ?? units.filter((u) => u.top <= y).at(-1);
      unit?.footnotes.push(index);
    });

    // place the units page by page
    const breaks: PageBreak[] = [];
    const pageNotes: number[][] = [[]];
    const pageCommands: Array<Array<{ command: string; value: string }>> = [[]];
    const pageMarks: string[][] = [[]];
    let page = 0;
    let breakPending = false;
    let offset = 0;
    let reserve = 0;
    const roomFor = (unit: Unit, reserved: number) =>
      unit.footnotes.length ? unit.footnotes.reduce((sum, i) => sum + noteHeights[i], 0) + (reserved ? 0 : separator) : 0;
    for (let i = 0; i < units.length; i++) {
      const unit = units[i];
      let top = unit.top + offset;
      let bottom = unit.bottom + offset;
      // after \newpage: the next page (from its top)
      if (breakPending) {
        breakPending = false;
        if (top < (page + 1) * stride) {
          const push = (page + 1) * stride - top;
          breaks.push({ pos: unit.pos, inline: unit.inline, height: push, block: unit.block, row: unit.row });
          offset += push;
          top += push;
          bottom += push;
        }
      }
      // past the foot of its page (an earlier block ran over): it starts on the page it is on — unless a heading on
      // this page is right before it: then both go to the next page (below)
      const previous = i > 0 ? units[i - 1] : null;
      const headingBefore =
        !!previous?.keepWithNext && previous.top + offset > page * stride + 0.5 && previous.bottom + offset <= page * stride + contentHeight;
      if (!headingBefore) {
        while (top >= page * stride + contentHeight - 0.5) {
          page++;
          reserve = 0;
          pageNotes[page] = [];
        }
        // (not on the first page: a line's glyphs may reach above its line box, a tight \baselineskip's, but nothing is
        // above the first page's top)
        if (page > 0 && top < page * stride) {
          const push = page * stride - top;
          breaks.push({ pos: unit.pos, inline: unit.inline, height: push, block: unit.block, row: unit.row });
          offset += push;
          top += push;
          bottom += push;
        }
      }
      const extra = roomFor(unit, reserve);
      const atPageTop = top <= page * stride + 0.5;
      if (bottom > page * stride + contentHeight - reserve - extra && !atPageTop) {
        // a heading's last line goes with what follows it
        let move = i;
        while (move > 0 && units[move - 1].keepWithNext && units[move - 1].top + offset > page * stride + 0.5) move--;
        const first = units[move];
        if (move !== i) {
          // the units moved back onto this page come off it again
          for (let k = move; k < i; k++) for (const n of units[k].footnotes) pageNotes[page] = pageNotes[page].filter((x) => x !== n);
        }
        const firstTop = first.top + offset;
        const push = (page + 1) * stride - firstTop;
        breaks.push({ pos: first.pos, inline: first.inline, height: push, block: first.block, row: first.row });
        offset += push;
        page++;
        reserve = 0;
        pageNotes[page] = [];
        for (let k = move; k < i; k++) {
          reserve += roomFor(units[k], reserve);
          pageNotes[page].push(...units[k].footnotes);
        }
        top = unit.top + offset;
        bottom = unit.bottom + offset;
      }
      reserve += roomFor(unit, reserve);
      pageNotes[page].push(...unit.footnotes);
      if (unit.command) (pageCommands[page] ??= []).push(unit.command);
      if (unit.mark !== undefined) (pageMarks[page] ??= []).push(unit.mark);
      if (unit.breakAfter) breakPending = true;
      // a block taller than the rest of the page runs over its foot
      if (bottom > page * stride + contentHeight) {
        while (bottom > page * stride + contentHeight) {
          page++;
          pageNotes[page] = [];
        }
        reserve = 0;
      }
    }
    // breaks are kept in document order, one per place
    breaks.sort((a, b) => a.pos - b.pos);
    const pages = page + 1;
    const state = paginationKey.getState(view.state);
    if (state && !sameBreaks(state.breaks, breaks)) {
      // breaks measured before with nothing changed since: the pages would only go back and forth between them
      const key = breaks.map((b) => `${b.pos}${b.inline ? "i" : "b"}${b.block ?? ""}:${Math.round(b.height)}`).join(" ");
      if (!this.seen.has(key)) {
        this.seen.add(state.breaks.map((b) => `${b.pos}${b.inline ? "i" : "b"}${b.block ?? ""}:${Math.round(b.height)}`).join(" "));
        this.seen.add(key);
        view.dispatch(view.state.tr.setMeta(paginationKey, breaks).setMeta("addToHistory", false));
      }
    }
    // the pages' numbers and styles, LaTeX's; the running head's mark: the first on the page (\rightmark is the
    // \firstmark), else the last one before it
    const labels = pageLabels(pages, pageCommands);
    const marks: string[] = [];
    let mark = "";
    for (let p = 0; p < pages; p++) {
      const own = pageMarks[p] ?? [];
      marks.push(own.length ? own[0] : mark);
      if (own.length) mark = own[own.length - 1];
    }
    this.renderLayer(sheet, settings, pxPerMm, stride, pages, pageNotes, notes, labels, marks);
  }

  /** Each footnote's height at the foot of a page (its text set at the text area's width). */
  private measureNotes(notes: Array<{ text: string; number: number }>, settings: DocumentSettings, pxPerMm: number): number[] {
    const width = (PAPER_SIZES[settings.paper].width - settings.marginLeft - settings.marginRight) * pxPerMm;
    this.measurer.style.width = `${width}px`;
    this.measurer.innerHTML = notes.map((note) => this.noteHtml(note.number, note.text)).join("");
    return Array.from(this.measurer.children).map((el) => (el as HTMLElement).getBoundingClientRect().height);
  }

  private noteHtml(number: number, text: string, pos?: number): string {
    const body = this.options.renderNote ? this.options.renderNote(text) : escapeHtml(text);
    return `<div class="vtx-page-footnote"${pos !== undefined ? ` data-footnote-pos="${pos}"` : ""}><sup>${number}</sup>${body || "&#8203;"}</div>`;
  }

  private renderLayer(
    sheet: HTMLElement,
    settings: DocumentSettings,
    pxPerMm: number,
    stride: number,
    pages: number,
    pageNotes: number[][],
    notes: Array<{ pos: number; text: string; number: number }>,
    labels: PageLabel[],
    marks: string[],
  ) {
    const paper = PAPER_SIZES[settings.paper];
    const sheetHeight = pages * stride - (stride - paper.height * pxPerMm);
    sheet.style.minHeight = `${sheetHeight}px`;
    const key = JSON.stringify([pages, stride, pxPerMm, settings, pageNotes, notes.map((n) => [n.pos, n.text]), labels, marks]);
    if (key === this.lastKey) return;
    this.lastKey = key;
    const parts: string[] = [];
    for (let page = 0; page < pages; page++) {
      const pageTop = page * stride;
      const footTop = pageTop + (paper.height - settings.marginBottom) * pxPerMm;
      const numbers = pageNotes[page] ?? [];
      if (numbers.length) {
        parts.push(
          `<div class="vtx-page-footnotes" style="left:${settings.marginLeft * pxPerMm}px;width:${(paper.width - settings.marginLeft - settings.marginRight) * pxPerMm}px;top:${footTop}px">` +
            numbers.map((i) => this.noteHtml(notes[i].number, notes[i].text, notes[i].pos)).join("") +
            "</div>",
        );
      }
      // the page style (article.cls, as ctexart keeps it): plain, the number centred at the foot (its baseline
      // \footskip, 30pt, under the text); headings, the section's mark slanted at the left and the number at the right
      // in the head (its baseline \headsep, 25pt, over the text); myheadings, the number there alone; empty, nothing
      const { pageStyle, text } = labels[page] ?? { pageStyle: "plain", text: String(page + 1) };
      const pt = pxPerMm * (25.4 / 72.27);
      const left = settings.marginLeft * pxPerMm;
      const width = (paper.width - settings.marginLeft - settings.marginRight) * pxPerMm;
      if (pageStyle === "plain" && text) {
        parts.push(
          `<div class="vtx-page-number" style="top:${pageTop + (paper.height - settings.marginBottom) * pxPerMm + 30 * pt}px;left:${left}px;width:${width}px"><span class="vtx-baseline-probe"></span>${escapeHtml(text)}</div>`,
        );
      } else if (pageStyle === "headings" || pageStyle === "myheadings") {
        parts.push(
          `<div class="vtx-page-head" style="top:${pageTop + settings.marginTop * pxPerMm - 25 * pt}px;left:${left}px;width:${width}px">` +
            `<span class="vtx-baseline-probe"></span><span class="vtx-page-mark">${pageStyle === "headings" ? marks[page] ?? "" : ""}</span><span class="vtx-page-number">${escapeHtml(text)}</span></div>`,
        );
      }
    }
    const html = parts.join("");
    // the measurer stays in the layer
    this.layer.innerHTML = html;
    this.layer.append(this.measurer);
    // the heads and feet are placed by their baseline: each one's top moved up by its baseline's depth in it (a
    // zero-height box on the baseline tells it, whatever the fonts)
    for (const el of Array.from(this.layer.querySelectorAll<HTMLElement>(".vtx-page-head, .vtx-page-layer > .vtx-page-number"))) {
      const probe = el.querySelector(".vtx-baseline-probe");
      if (!probe) continue;
      const depth = probe.getBoundingClientRect().top - el.getBoundingClientRect().top;
      el.style.top = `${parseFloat(el.style.top) - depth}px`;
    }
  }

  destroy() {
    window.clearTimeout(this.timer);
    cancelAnimationFrame(this.frame);
    window.removeEventListener("resize", this.changed);
    document.fonts?.removeEventListener?.("loadingdone", this.changed);
    this.layer.removeEventListener("mousedown", this.onLayerMouseDown);
    this.layer.remove();
  }
}

const pageViews = new WeakMap<EditorView, PaginationView>();

/** The page view's pages: spacers where pages end, footnotes at their feet, page numbers. */
export function paginationPlugin(options: PaginationOptions): Plugin<PaginationState> {
  let pageView: PaginationView | null = null;
  return new Plugin<PaginationState>({
    key: paginationKey,
    state: {
      init: () => ({ decorations: DecorationSet.empty, breaks: [] as PageBreak[] }),
      apply: (tr, value) => {
        const breaks = tr.getMeta(paginationKey) as PageBreak[] | undefined;
        if (breaks) return { decorations: buildDecorations(tr.doc, breaks), breaks };
        if (!tr.docChanged) return value;
        // until the next measurement the spacers keep their places in the changed text
        const decorations = value.decorations.map(tr.mapping, tr.doc);
        const mapped = value.breaks.map((b) => ({ ...b, pos: tr.mapping.map(b.pos, -1) }));
        return { decorations, breaks: mapped };
      },
    },
    view: (view) => {
      pageView = new PaginationView(view, options);
      pageViews.set(view, pageView);
      return pageView;
    },
    props: {
      decorations: (state) => paginationKey.getState(state)?.decorations ?? DecorationSet.empty,
    },
  });
}

/** Measure the pages again (the zoom, the fonts or the window changed). */
export function refreshPagination(view: EditorView) {
  pageViews.get(view)?.refresh();
  view.dispatch(view.state.tr.setMeta(paginationKey, paginationKey.getState(view.state)?.breaks ?? []).setMeta("addToHistory", false));
}
