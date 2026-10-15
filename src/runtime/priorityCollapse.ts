// Collapses a toolbar row by priority, based on the row's own free space.
//
// The row gets `data-vt-collapse="1 2 … n"`: the smallest n at which
// `fits(n)` is true. CSS decides what each level hides or compacts
// (`[data-vt-collapse~="2"] .file-actions { display: none }`), so nothing is
// hidden while the row still has room, and the least important items go first.
//
// The level is a pure function of the row width and the natural widths of its
// items: every level is tried in order, synchronously, before the next paint.
// The row's width must not depend on its content, otherwise the result could
// feed back into the measurement.

export interface PriorityCollapseOptions {
  /** Highest level; it is accepted even when `fits` is still false. */
  maxLevel: number;
  /** Whether the row's content fits with the current level applied. */
  fits: (row: HTMLElement, level: number) => boolean;
}

function levelTokens(level: number) {
  let tokens = "";
  for (let index = 1; index <= level; index += 1) {
    tokens += (index > 1 ? " " : "") + index;
  }
  return tokens;
}

export function trackPriorityCollapse(
  row: HTMLElement,
  { maxLevel, fits }: PriorityCollapseOptions,
) {
  let frame = 0;
  let disposed = false;

  const apply = (level: number) => {
    const next = levelTokens(level);
    // The attribute is present from level 0 on, so CSS can tell a tracked row.
    if (row.dataset.vtCollapse !== next) row.dataset.vtCollapse = next;
  };

  const settle = () => {
    frame = 0;
    if (disposed || !row.isConnected) return;
    for (let level = 0; level <= maxLevel; level += 1) {
      apply(level);
      if (level === maxLevel || fits(row, level)) break;
    }
  };

  const schedule = () => {
    if (frame || disposed) return;
    frame = window.requestAnimationFrame(settle);
  };

  // Size changes of the row; items added, removed or relabelled inside it.
  // The attribute written by `apply` is not observed.
  const resizeObserver = new ResizeObserver(schedule);
  resizeObserver.observe(row);
  const mutationObserver = new MutationObserver(schedule);
  mutationObserver.observe(row, {
    childList: true,
    subtree: true,
    characterData: true,
  });
  void document.fonts?.ready.then(schedule);

  settle();

  return () => {
    disposed = true;
    if (frame) window.cancelAnimationFrame(frame);
    resizeObserver.disconnect();
    mutationObserver.disconnect();
    delete row.dataset.vtCollapse;
  };
}

/** Sum of the in-flow children's widths plus gaps and padding fits the row.
    Children must not shrink or grow (`flex: none`); margins are ignored, so an
    `auto` margin can be used to push a group to the far side. */
export function childrenFitRow(row: HTMLElement) {
  const style = window.getComputedStyle(row);
  const gap = Number.parseFloat(style.columnGap) || 0;
  let required =
    (Number.parseFloat(style.paddingLeft) || 0) +
    (Number.parseFloat(style.paddingRight) || 0);
  let count = 0;
  for (const child of Array.from(row.children)) {
    if (!(child instanceof HTMLElement)) continue;
    const childStyle = window.getComputedStyle(child);
    if (
      childStyle.display === "none" ||
      childStyle.position === "absolute" ||
      childStyle.position === "fixed"
    ) {
      continue;
    }
    const width = child.getBoundingClientRect().width;
    if (width <= 0) continue;
    required += width;
    count += 1;
  }
  required += gap * Math.max(0, count - 1);
  return required <= row.clientWidth + 0.5;
}
