// Where content dragged in the document lands (documentDrag.ts, and a formula's selection dragged out of the
// formula, formulas.ts) and the caret that marks the place.
import type { Slice } from "prosemirror-model";
import { dropPoint } from "prosemirror-transform";
import type { EditorView } from "prosemirror-view";
import { isFormulaNodeName } from "./schema.ts";

/** Where content (`slice`, now at [from, to]) dropped at a page point would go: a position outside [from, to]. */
export function dropPlace(view: EditorView, x: number, y: number, slice: Slice, from: number, to: number): number | null {
  const hit = view.posAtCoords({ left: x, top: y });
  if (!hit) return null;
  let pos = hit.pos;
  // Onto a formula: before or after it, whichever side is nearer.
  const inside = hit.inside >= 0 ? view.state.doc.nodeAt(hit.inside) : null;
  if (inside && isFormulaNodeName(inside.type.name) && inside.isInline) {
    const r = (view.nodeDOM(hit.inside) as HTMLElement | null)?.getBoundingClientRect();
    pos = r && x > (r.left + r.right) / 2 ? hit.inside + inside.nodeSize : hit.inside;
  }
  pos = dropPoint(view.state.doc, pos, slice) ?? pos;
  return pos >= from && pos <= to ? null : pos;
}

/** The drop caret: a bar at a place in a line, or a line across the width between blocks. */
export class DropMarker {
  private element: HTMLElement | null = null;

  show(view: EditorView, pos: number | null) {
    if (pos === null) {
      this.remove();
      return;
    }
    const $pos = view.state.doc.resolve(pos);
    let style: Partial<CSSStyleDeclaration>;
    if ($pos.parent.inlineContent) {
      const c = view.coordsAtPos(pos);
      style = { left: `${c.left - 1}px`, top: `${c.top}px`, width: "2px", height: `${Math.max(10, c.bottom - c.top)}px` };
    } else {
      const after = $pos.nodeAfter ? (view.nodeDOM(pos) as HTMLElement | null) : null;
      const before = !after && $pos.nodeBefore ? (view.nodeDOM(pos - $pos.nodeBefore.nodeSize) as HTMLElement | null) : null;
      const rect = (after ?? before)?.getBoundingClientRect();
      if (!rect) {
        this.remove();
        return;
      }
      const y = after ? rect.top - 2 : rect.bottom + 1;
      style = { left: `${rect.left}px`, top: `${y}px`, width: `${rect.width}px`, height: "2px" };
    }
    if (!this.element) {
      this.element = document.createElement("div");
      this.element.className = "vtx-drop-marker is-caret";
      document.body.append(this.element);
    }
    Object.assign(this.element.style, style);
  }

  remove() {
    this.element?.remove();
    this.element = null;
  }
}
