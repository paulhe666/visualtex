// Dragging a selection in the document (AxMath's drag, DESIGN.md §6.6), with the pointer instead of the browser's
// drag and drop: a press on the selection (text, formulas in it, a selected formula or block) and a move of 3 px drag
// it; a drop caret follows the pointer and the document scrolls at its edges; the release moves it there (Ctrl held:
// a copy) as one undo step, the moved content selected. Esc or the right button cancel; a press without a move is a
// click (on a formula in the selection it opens the formula). The browser's drag and drop is not used at all: the
// Windows app's webview keeps it for files (Tauri's drag-drop handler), so a native drag never reaches the page there.
import type { Node, ResolvedPos, Slice } from "prosemirror-model";
import { NodeSelection, Plugin, PluginKey, Selection, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { isFormulaNodeName } from "./schema.ts";
import { formulaParts, formulaViewOf, openFormula } from "./formulas.ts";
import { deleteSelectionWithParts, selectionSliceWithParts, wholeFormulasRange } from "./mixedSelection.ts";
import { DropMarker, dropPlace } from "./dropPlace.ts";
import { imageAspect, imageDropMarker, imageDropTarget, isImage, moveImage, type ImageDropTarget } from "./images.ts";
import { BoardDrag, boardDropMarker, boardDropTarget, isBoardImage, selectBoardImage } from "./imageBoards.ts";

/** Where a dragged image lands: a place on an image board under the pointer, else the image rules (images.ts). */
function anyImageTarget(view: EditorView, x: number, y: number, image: Node, from: number, to: number, exceptBoard = -1): ImageDropTarget | null {
  const aspect = imageAspect(image.attrs.src);
  return boardDropTarget(view, x, y, image.attrs.width, aspect, exceptBoard) ?? imageDropTarget(view, x, y, image, from, to);
}

function anyImageMarker(view: EditorView, target: ImageDropTarget | null, image: Node, element: HTMLElement): boolean {
  if (target?.kind === "board") return boardDropMarker(view, target, image.attrs.width, imageAspect(image.attrs.src), element);
  return imageDropMarker(view, target, element);
}

export const documentDragKey = new PluginKey("visualtex-document-drag");

const THRESHOLD = 3;

/** The element that scrolls the document. */
function scroller(view: EditorView): HTMLElement | null {
  for (let node = view.dom.parentElement; node; node = node.parentElement) {
    const style = getComputedStyle(node);
    if (/(auto|scroll)/.test(style.overflowY) && node.scrollHeight > node.clientHeight) return node;
  }
  return null;
}

/** Whether a page point is on the drawn selection: its text, a formula in it, or the selected node. */
function onSelection(view: EditorView, x: number, y: number): boolean {
  const sel = view.state.selection;
  const inRect = (r: DOMRect) => r.width > 0 && x >= r.left - 1 && x <= r.right + 1 && y >= r.top - 1 && y <= r.bottom + 1;
  // the formulas a selection through the text takes a part of (a press anywhere on them: a click there opens them)
  for (const part of formulaParts(view.state)) {
    const dom = formulaViewOf(view, part.id)?.dom;
    if (dom && inRect(dom.getBoundingClientRect())) return true;
  }
  if (sel.empty) return false;
  if (sel instanceof NodeSelection) {
    const dom = view.nodeDOM(sel.from) as HTMLElement | null;
    return !!dom && inRect(dom.getBoundingClientRect());
  }
  // The text: the DOM ranges of the selection, line by line.
  try {
    const start = view.domAtPos(sel.from);
    const end = view.domAtPos(sel.to);
    const range = document.createRange();
    range.setStart(start.node, start.offset);
    range.setEnd(end.node, end.offset);
    for (const rect of range.getClientRects()) if (inRect(rect)) return true;
  } catch {
    // positions without a DOM point: fall through to the formulas
  }
  // Formulas and other atoms fully inside the selection.
  let hit = false;
  view.state.doc.nodesBetween(sel.from, sel.to, (node, pos) => {
    if (hit || !node.isAtom || pos < sel.from || pos + node.nodeSize > sel.to) return !hit;
    const dom = view.nodeDOM(pos);
    if (dom instanceof Element && inRect(dom.getBoundingClientRect())) hit = true;
    return false;
  });
  return hit;
}

/**
 * A range from the end of a line (or the start of one) to the start of a later line (or its end) holds the blocks
 * between those edges, not the line breaks: those blocks. Null when an end is inside a line, or both are in one.
 */
function wholeBlocks(doc: Node, from: number, to: number): { from: number; to: number } | null {
  const $from = doc.resolve(from);
  const $to = doc.resolve(to);
  if ($from.parent.inlineContent && $to.parent.inlineContent && $from.sameParent($to)) return null;
  const edge = ($pos: ResolvedPos, pos: number) =>
    !$pos.parent.inlineContent ? pos : $pos.parentOffset === 0 ? $pos.before() : $pos.parentOffset === $pos.parent.content.size ? $pos.after() : null;
  const start = edge($from, from);
  const end = edge($to, to);
  if (start === null || end === null || start >= end || !doc.resolve(start).sameParent(doc.resolve(end))) return null;
  return { from: start, to: end };
}

class SelectionDrag {
  private dragging = false;
  private cancelled = false;
  private readonly marker = new DropMarker();
  private target: number | null = null;
  private last: { x: number; y: number; copy: boolean } | null = null;
  private timer = 0;
  private readonly slice: Slice;
  private readonly from: number;
  private readonly to: number;
  private readonly mixed: boolean;
  /** a selected image: dropped by the image rules (images.ts), not as a slice */
  private readonly image: boolean;
  private imageTarget: ImageDropTarget | null = null;
  private imageMarker: HTMLElement | null = null;

  constructor(
    private readonly view: EditorView,
    private readonly start: PointerEvent,
  ) {
    const sel = view.state.selection;
    // with formulas partly selected: their parts go with the text (mixedSelection.ts); a formula taken whole goes
    // whole, and ends at the edges of lines take the blocks between them (a display formula with no line joined)
    const whole = wholeFormulasRange(view);
    const blocks = whole ? wholeBlocks(view.state.doc, whole.from, whole.to) : null;
    const mixed = whole ? null : selectionSliceWithParts(view);
    const from = blocks?.from ?? whole?.from;
    const to = blocks?.to ?? whole?.to;
    this.slice = from !== undefined && to !== undefined ? view.state.doc.slice(from, to) : (mixed?.slice ?? sel.content());
    this.from = from ?? mixed?.from ?? sel.from;
    this.to = to ?? mixed?.to ?? sel.to;
    this.mixed = !!mixed;
    this.image = sel instanceof NodeSelection && isImage(sel.node);
    window.addEventListener("pointermove", this.onMove, true);
    window.addEventListener("pointerup", this.onUp, true);
    window.addEventListener("pointercancel", this.onCancel, true);
    window.addEventListener("mousedown", this.onOtherButton, true);
    window.addEventListener("keydown", this.onKey, true);
    this.timer = window.setInterval(this.autoScroll, 32);
  }

  private finish() {
    window.removeEventListener("pointermove", this.onMove, true);
    window.removeEventListener("pointerup", this.onUp, true);
    window.removeEventListener("pointercancel", this.onCancel, true);
    window.removeEventListener("mousedown", this.onOtherButton, true);
    window.removeEventListener("keydown", this.onKey, true);
    window.clearInterval(this.timer);
    this.marker.remove();
    this.imageMarker?.remove();
    this.imageMarker = null;
    document.documentElement.style.removeProperty("cursor");
    this.view.dom.classList.remove("is-dragging-selection");
  }

  /** Where the selection would land: a position it can be inserted at, outside itself. */
  private place() {
    const last = this.last;
    if (!last) return;
    if (this.image) {
      const node = this.view.state.doc.nodeAt(this.from)!;
      this.imageTarget = anyImageTarget(this.view, last.x, last.y, node, this.from, this.to);
      if (!this.imageMarker) {
        this.imageMarker = document.createElement("div");
        this.imageMarker.className = "vtx-drop-marker is-caret";
        document.body.append(this.imageMarker);
      }
      this.imageMarker.classList.toggle("is-wrap", this.imageTarget?.kind === "wrap");
      this.imageMarker.classList.toggle("is-box", this.imageTarget?.kind === "board");
      this.imageMarker.hidden = !anyImageMarker(this.view, this.imageTarget, node, this.imageMarker);
      return;
    }
    this.target = dropPlace(this.view, last.x, last.y, this.slice, this.from, this.to);
    this.marker.show(this.view, this.target);
  }

  private autoScroll = () => {
    if (!this.dragging || !this.last) return;
    const frame = scroller(this.view);
    if (!frame) return;
    const r = frame.getBoundingClientRect();
    const step = this.last.y < r.top + 24 ? -16 : this.last.y > r.bottom - 24 ? 16 : 0;
    if (!step) return;
    const before = frame.scrollTop;
    frame.scrollTop += step;
    if (frame.scrollTop !== before) this.place();
  };

  private onMove = (event: PointerEvent) => {
    if (event.pointerId !== this.start.pointerId || this.cancelled) return;
    if (event.buttons & 2) {
      this.cancel();
      return;
    }
    if (!this.dragging) {
      if (Math.abs(event.clientX - this.start.clientX) < THRESHOLD && Math.abs(event.clientY - this.start.clientY) < THRESHOLD) return;
      this.dragging = true;
      this.view.dom.classList.add("is-dragging-selection");
    }
    event.preventDefault();
    const copy = event.ctrlKey || event.altKey;
    document.documentElement.style.setProperty("cursor", copy ? "copy" : "move", "important");
    this.last = { x: event.clientX, y: event.clientY, copy };
    this.place();
  };

  private onUp = (event: PointerEvent) => {
    if (event.pointerId !== this.start.pointerId) return;
    this.finish();
    if (this.cancelled) return;
    const view = this.view;
    if (!this.dragging && this.image) {
      // a click on an image keeps it selected (its bar and handles)
      view.focus();
      return;
    }
    if (!this.dragging) {
      // A click on the selection: on a formula in it, the formula opens; elsewhere the caret goes there.
      const hit = view.posAtCoords({ left: event.clientX, top: event.clientY });
      if (!hit) return;
      const inside = hit.inside >= 0 ? view.state.doc.nodeAt(hit.inside) : null;
      if (inside && isFormulaNodeName(inside.type.name)) {
        view.dispatch(openFormula(view.state.tr, hit.inside, { kind: "point", x: event.clientX, y: event.clientY }));
      } else {
        view.dispatch(view.state.tr.setSelection(Selection.near(view.state.doc.resolve(hit.pos))));
      }
      view.focus();
      return;
    }
    event.preventDefault();
    if (this.image) {
      const tr = this.imageTarget ? moveImage(view, this.imageTarget, this.from, this.to, event.ctrlKey || event.altKey) : null;
      if (tr) view.dispatch(tr.setMeta("uiEvent", "drop").scrollIntoView());
      view.focus();
      return;
    }
    const target = this.target;
    if (target === null) return;
    const copy = event.ctrlKey || event.altKey;
    let tr = view.state.tr;
    if (!copy) {
      if (this.mixed) tr = deleteSelectionWithParts(view) ?? tr;
      else tr.deleteRange(this.from, this.to);
    }
    const pos = tr.mapping.map(target);
    const slice = this.slice;
    const single = slice.openStart === 0 && slice.openEnd === 0 && slice.content.childCount === 1;
    const before = tr.doc;
    if (single) tr.replaceRangeWith(pos, pos, slice.content.firstChild!);
    else tr.replaceRange(pos, pos, slice);
    if (tr.doc.eq(before)) return;
    // The moved content is selected.
    const $pos = tr.doc.resolve(pos);
    if (single && NodeSelection.isSelectable(slice.content.firstChild!) && $pos.nodeAfter?.sameMarkup(slice.content.firstChild!)) {
      tr.setSelection(new NodeSelection($pos));
    } else {
      let end = tr.mapping.map(target);
      tr.mapping.maps[tr.mapping.maps.length - 1].forEach((_f, _t, _nf, newTo) => (end = newTo));
      tr.setSelection(TextSelection.between($pos, tr.doc.resolve(end)));
    }
    view.dispatch(tr.setMeta("uiEvent", "drop").scrollIntoView());
    view.focus();
  };

  private onCancel = (event: PointerEvent) => {
    if (event.pointerId === this.start.pointerId) this.cancel();
  };

  private onKey = (event: KeyboardEvent) => {
    if (event.key !== "Escape") return;
    event.preventDefault();
    event.stopImmediatePropagation();
    this.cancel();
  };

  /** The right button during the drag cancels it (as a Windows drag does); its context menu does not open. */
  private onOtherButton = (event: MouseEvent) => {
    if (event.button === 0) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    this.cancel();
    window.addEventListener("contextmenu", (e) => e.preventDefault(), { once: true, capture: true });
  };

  private cancel() {
    this.cancelled = true;
    this.finish();
  }
}

export function documentDragPlugin(): Plugin {
  return new Plugin({
    key: documentDragKey,
    props: {
      handleDOMEvents: {
        pointerdown(view, event) {
          if (event.button !== 0 || !event.isPrimary) return false;
          // an image on a board: selected (Shift/Ctrl: added to the selection) and moved as on a slide
          const boardImage = (event.target as Element | null)?.closest?.(".vtx-board-image");
          if (boardImage && !(event.target as Element).classList.contains("vtx-board-handle")) {
            const pos = view.posAtDOM(boardImage, 0);
            const node = pos >= 0 ? view.state.doc.nodeAt(pos) : null;
            if (node && isBoardImage(node)) {
              const add = event.shiftKey || event.ctrlKey || event.metaKey;
              selectBoardImage(view, pos, add);
              if (!add) {
                const boardPos = view.state.doc.resolve(pos).before();
                new BoardDrag(
                  view,
                  event,
                  boardPos,
                  pos,
                  (x, y, image, from) => anyImageTarget(view, x, y, image, from, from + image.nodeSize, boardPos),
                  (target, element) => anyImageMarker(view, target, view.state.doc.nodeAt(pos) ?? node, element),
                );
              }
              event.preventDefault();
              view.focus();
              return true;
            }
          }
          if (event.shiftKey) return false;
          // An open formula (its editor, its number) handles its own presses; one showing a part is the document's.
          if ((event.target as Element | null)?.closest?.(".is-open")) return false;
          // a press on an image selects it and may drag it at once (Notion, Word)
          const imageDom = (event.target as Element | null)?.closest?.(".vtx-image");
          if (imageDom && !(event.target as Element).classList.contains("vtx-image-handle")) {
            const pos = view.posAtDOM(imageDom, 0);
            const node = pos >= 0 ? view.state.doc.nodeAt(pos) : null;
            if (node && isImage(node)) {
              view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)));
              new SelectionDrag(view, event);
              event.preventDefault();
              view.focus();
              return true;
            }
          }
          if (!onSelection(view, event.clientX, event.clientY)) return false;
          new SelectionDrag(view, event);
          // Neither the editor's own press handling nor the browser's selection or drag starts.
          event.preventDefault();
          return true;
        },
        // The browser's drag and drop of the selection is replaced by the drag above.
        dragstart(_view, event) {
          event.preventDefault();
          return true;
        },
      },
    },
  });
}
