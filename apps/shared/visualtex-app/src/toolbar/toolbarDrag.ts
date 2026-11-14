// Dragging in the toolbars' edit mode (FormulaToolbar.tsx): a card, a tile or a group's tab follows the pointer as a
// ghost; where it would land shows as a line between the cards; it lands on release. Pointer events, not the
// browser's drag and drop (the same in WebKit and WebView2, and clear of the formula editor's own drops). The places
// a thing can land are found in the page by their attributes, so a drag goes from one toolbar to another (the
// classic bottom toolbar and the tiles panel are separate components).
//
//   data-drop-kind="group" | "common-tiles" | "custom-section" | "group-tab" | "tabs"   a place
//   data-drop-flow="column"   its cards run down then across (the horizontal toolbar); else across then down
//   data-group-id / data-section-id   which group or section
//   data-drag-index="n"   a card's position in its place

export type DragSource =
  | { type: "toolbar"; groupId: string; index: number }
  | { type: "common-tile"; index: number; tileId: string }
  | { type: "custom-tile"; tileId: string }
  | { type: "group-tab"; groupId: string };

export type DropTarget =
  | { type: "group"; groupId: string; index: number }
  | { type: "common-tiles"; index: number }
  | { type: "custom-section"; sectionId: string; index: number }
  | { type: "tabs"; index: number };

export interface DragHandlers {
  accepts: (source: DragSource, target: DropTarget) => boolean;
  drop: (source: DragSource, target: DropTarget) => void;
}

const START_DISTANCE = 5;
const EDGE = 36;
const TAB_HOVER_MS = 450;
/** an index past every card: the end */
export const END_INDEX = Number.MAX_SAFE_INTEGER;

interface Placed {
  index: number;
  rect: DOMRect;
}

function placedItems(container: Element): Placed[] {
  const items: Placed[] = [];
  container.querySelectorAll<HTMLElement>("[data-drag-index]").forEach((element) => {
    // a card that is itself a place (a group's tab among the tabs) belongs to the place around it
    const owner = element.hasAttribute("data-drop-kind") ? element.parentElement?.closest("[data-drop-kind]") : element.closest("[data-drop-kind]");
    if (owner !== container) return;
    const rect = element.getBoundingClientRect();
    if (!rect.width && !rect.height) return;
    items.push({ index: Number(element.dataset.dragIndex), rect });
  });
  return items.sort((a, b) => a.index - b.index);
}

/** The index a drop at a point means among a place's cards, and the line that shows it. */
function insertionAt(container: HTMLElement, x: number, y: number) {
  const items = placedItems(container);
  const column = container.dataset.dropFlow === "column";
  for (const item of items) {
    const r = item.rect;
    const before = column ? x < r.left || (x <= r.right && y < r.top + r.height / 2) : y < r.top || (y <= r.bottom && x < r.left + r.width / 2);
    if (before) return { index: item.index, line: lineBefore(r, column) };
  }
  const last = items[items.length - 1];
  if (last) return { index: last.index + 1, line: lineAfter(last.rect, column) };
  const box = container.getBoundingClientRect();
  return { index: 0, line: column ? { left: box.left + 4, top: box.top + 4, width: Math.min(48, box.width - 8), height: 2 } : { left: box.left + 4, top: box.top + 4, width: 2, height: Math.min(40, box.height - 8) } };
}

function lineBefore(r: DOMRect, column: boolean) {
  return column ? { left: r.left, top: r.top - 2, width: r.width, height: 3 } : { left: r.left - 2, top: r.top, width: 3, height: r.height };
}

function lineAfter(r: DOMRect, column: boolean) {
  return column ? { left: r.left, top: r.bottom - 1, width: r.width, height: 3 } : { left: r.right - 1, top: r.top, width: 3, height: r.height };
}

function targetFor(source: DragSource, element: Element | null, x: number, y: number) {
  if (!element) return null;
  if (source.type === "group-tab") {
    const tabs = element.closest<HTMLElement>("[data-drop-kind='tabs']");
    if (!tabs) return null;
    const { index, line } = insertionAt(tabs, x, y);
    return { target: { type: "tabs", index } as DropTarget, line, container: tabs };
  }
  const container = element.closest<HTMLElement>("[data-drop-kind]");
  if (!container) return null;
  const kind = container.dataset.dropKind;
  if (kind === "group-tab") {
    const box = container.getBoundingClientRect();
    return {
      target: { type: "group", groupId: container.dataset.groupId ?? "", index: END_INDEX } as DropTarget,
      line: { left: box.left, top: box.bottom - 3, width: box.width, height: 3 },
      container,
      tab: true,
    };
  }
  if (kind !== "group" && kind !== "common-tiles" && kind !== "custom-section") return null;
  const { index, line } = insertionAt(container, x, y);
  const target: DropTarget =
    kind === "group"
      ? { type: "group", groupId: container.dataset.groupId ?? "", index }
      : kind === "common-tiles"
        ? { type: "common-tiles", index }
        : { type: "custom-section", sectionId: container.dataset.sectionId ?? "", index };
  return { target, line, container };
}

function scrollParent(element: HTMLElement | null): HTMLElement | null {
  for (let node = element; node; node = node.parentElement) {
    const style = getComputedStyle(node);
    if (/(auto|scroll)/.test(style.overflowX + style.overflowY) && (node.scrollWidth > node.clientWidth || node.scrollHeight > node.clientHeight)) return node;
  }
  return null;
}

let active = false;

/** A press on a card in edit mode: a drag once the pointer has moved a little (else it stays a click). */
export function armDrag(event: PointerEvent | { clientX: number; clientY: number; button: number; pointerId?: number }, source: DragSource, element: HTMLElement, handlers: DragHandlers) {
  if (active || event.button !== 0) return;
  const startX = event.clientX;
  const startY = event.clientY;
  let started = false;
  let ghost: HTMLElement | null = null;
  let line: HTMLElement | null = null;
  let offsetX = 0;
  let offsetY = 0;
  let current: ReturnType<typeof targetFor> = null;
  let pointer = { x: startX, y: startY };
  let frame = 0;
  let tabTimer = 0;
  let hoveredTab: HTMLElement | null = null;

  const begin = () => {
    started = true;
    active = true;
    const rect = element.getBoundingClientRect();
    offsetX = startX - rect.left;
    offsetY = startY - rect.top;
    ghost = element.cloneNode(true) as HTMLElement;
    ghost.classList.add("vtx-toolbar-drag-ghost");
    ghost.removeAttribute("data-drag-index");
    ghost.removeAttribute("data-drop-kind");
    Object.assign(ghost.style, { position: "fixed", left: `${rect.left}px`, top: `${rect.top}px`, width: `${rect.width}px`, height: `${rect.height}px`, margin: "0", pointerEvents: "none", zIndex: "2147483000" });
    document.body.append(ghost);
    line = document.createElement("div");
    line.className = "vtx-toolbar-drop-line";
    line.style.display = "none";
    document.body.append(line);
    element.classList.add("is-drag-source");
    document.body.classList.add("vtx-toolbar-dragging");
    frame = requestAnimationFrame(tick);
  };

  const update = () => {
    if (!ghost || !line) return;
    ghost.style.left = `${pointer.x - offsetX}px`;
    ghost.style.top = `${pointer.y - offsetY}px`;
    const hit = document.elementFromPoint(pointer.x, pointer.y);
    const found = targetFor(source, hit, pointer.x, pointer.y);
    current = found && handlers.accepts(source, found.target) ? found : null;
    document.querySelectorAll(".is-drop-target").forEach((node) => node.classList.remove("is-drop-target"));
    if (current) {
      current.container.classList.add("is-drop-target");
      Object.assign(line.style, { display: "block", left: `${current.line.left}px`, top: `${current.line.top}px`, width: `${current.line.width}px`, height: `${current.line.height}px` });
    } else line.style.display = "none";
    // resting on a group's tab shows that group (as a click on the tab does), so the card can go among its cards
    const tab = current && "tab" in current ? current.container : null;
    if (tab !== hoveredTab) {
      hoveredTab = tab;
      window.clearTimeout(tabTimer);
      if (tab) tabTimer = window.setTimeout(() => tab.click(), TAB_HOVER_MS);
    }
  };

  // past the edge of a scrolling place: it scrolls (the place under the pointer, or the one last under it)
  const tick = () => {
    frame = requestAnimationFrame(tick);
    const scroller = scrollParent(current?.container ?? (document.elementFromPoint(pointer.x, pointer.y) as HTMLElement | null));
    if (!scroller) return;
    const box = scroller.getBoundingClientRect();
    let dx = 0;
    let dy = 0;
    if (scroller.scrollWidth > scroller.clientWidth) dx = pointer.x < box.left + EDGE ? -12 : pointer.x > box.right - EDGE ? 12 : 0;
    if (scroller.scrollHeight > scroller.clientHeight) dy = pointer.y < box.top + EDGE ? -12 : pointer.y > box.bottom - EDGE ? 12 : 0;
    if (dx || dy) {
      scroller.scrollBy(dx, dy);
      update();
    }
  };

  const finish = (dropIt: boolean) => {
    window.removeEventListener("pointermove", onMove, true);
    window.removeEventListener("pointerup", onUp, true);
    window.removeEventListener("pointercancel", onCancel, true);
    window.removeEventListener("keydown", onKey, true);
    cancelAnimationFrame(frame);
    window.clearTimeout(tabTimer);
    if (!started) return;
    active = false;
    ghost?.remove();
    line?.remove();
    element.classList.remove("is-drag-source");
    document.body.classList.remove("vtx-toolbar-dragging");
    document.querySelectorAll(".is-drop-target").forEach((node) => node.classList.remove("is-drop-target"));
    // the click the release makes is not a click on the card
    const swallow = (click: Event) => {
      click.preventDefault();
      click.stopPropagation();
    };
    window.addEventListener("click", swallow, { capture: true, once: true });
    window.setTimeout(() => window.removeEventListener("click", swallow, true), 0);
    if (dropIt && current) handlers.drop(source, current.target);
  };

  const onMove = (move: PointerEvent) => {
    pointer = { x: move.clientX, y: move.clientY };
    if (!started) {
      if (Math.hypot(move.clientX - startX, move.clientY - startY) < START_DISTANCE) return;
      begin();
    }
    move.preventDefault();
    update();
  };
  const onUp = () => finish(true);
  const onCancel = () => finish(false);
  const onKey = (key: KeyboardEvent) => {
    if (key.key !== "Escape" || !started) return;
    key.preventDefault();
    key.stopPropagation();
    finish(false);
  };
  window.addEventListener("pointermove", onMove, true);
  window.addEventListener("pointerup", onUp, true);
  window.addEventListener("pointercancel", onCancel, true);
  window.addEventListener("keydown", onKey, true);
}
