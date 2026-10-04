import type { FieldBlock } from "./art.generated";

const CANVAS_WIDTH = 1600; // design width of the formula field
const GAP = 26; // breathing room around every block and cleared element
const SIZES = [15, 16, 17, 18, 19, 20, 22, 24];
const SEED = 20261002;
// Cards waiting for their reveal animation sit this far below their resting place.
export const REVEAL_OFFSET = 28;

type Rect = { x: number; y: number; w: number; h: number };
/** A placed block, in field (canvas) units. */
export type PlacedBlock = Rect & { el: HTMLElement };

function random(seed: number) {
  return () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function makeBlock(block: FieldBlock, size: number) {
  const el = document.createElement("div");
  el.className = block.box ? "landing-blk landing-blk-box" : "landing-blk";
  el.style.fontSize = `${size}px`;
  const typeset = document.createElement("div");
  typeset.className = "landing-blk-tex";
  typeset.innerHTML = block.html;
  const source = document.createElement("code");
  source.className = "landing-blk-src";
  source.textContent = block.src;
  el.append(typeset, source);
  return el;
}

/** Size the source text so it wraps inside the block it stands for (monospace ≈ 0.6em/char). */
function fitSource(el: HTMLElement, w: number, h: number, size: number) {
  const source = el.lastElementChild as HTMLElement;
  const chars = source.textContent!.length;
  // allowed to spill past the block a little: it is only ever seen through the lens
  const fit = Math.sqrt((w * Math.max(h, size * 1.3) * 2) / (chars * 0.6 * 1.3));
  source.style.fontSize = `${Math.max(10, Math.min(size * 0.8, fit)).toFixed(1)}px`;
}

/**
 * Packs the notebook blocks into one continuous field as tall as `page`, leaving every
 * `[data-clear]` element uncovered. Deterministic for a given page width.
 */
export function layoutFormulaField(page: HTMLElement, field: HTMLElement, blocks: readonly FieldBlock[]): PlacedBlock[] {
  const viewportWidth = page.clientWidth;
  const scale = Math.max(viewportWidth / CANVAS_WIDTH, 0.5);
  const offsetX = (viewportWidth - CANVAS_WIDTH * scale) / 2;
  field.style.transform = `translate(${offsetX}px, 0) scale(${scale})`;
  field.replaceChildren();
  const height = Math.ceil(page.scrollHeight / scale);
  field.style.height = `${height}px`;

  const pageTop = page.getBoundingClientRect().top;
  const taken: Rect[] = [...page.querySelectorAll<HTMLElement>("[data-clear]")].map((el) => {
    const r = el.getBoundingClientRect();
    const shift = el.matches("[data-reveal]:not(.is-in)") ? REVEAL_OFFSET : 0;
    // a fixed element (the nav) only needs to be kept clear where it sits at the very top
    const top = el.dataset.clear === "fixed" ? r.top : r.top - pageTop - shift;
    return {
      x: (r.left - offsetX) / scale - GAP,
      y: top / scale - GAP,
      w: r.width / scale + 2 * GAP,
      h: r.height / scale + 2 * GAP,
    };
  });
  const collides = (x: number, y: number, w: number, h: number) =>
    taken.some((r) => x < r.x + r.w && x + w > r.x && y < r.y + r.h && y + h > r.y);
  const rand = random(SEED);
  const placed: PlacedBlock[] = [];
  // Notes are sentences: keep them wholly inside the visible part of the canvas.
  const visibleLeft = Math.max(0, -offsetX / scale) + 12;
  const visibleRight = Math.min(CANVAS_WIDTH, (viewportWidth - offsetX) / scale) - 12;
  const place = (el: HTMLElement, x: number, y: number, w: number, h: number, size: number) => {
    el.style.left = `${x}px`;
    el.style.top = `${y}px`;
    fitSource(el, w, h, size);
    taken.push({ x: x - GAP, y: y - GAP, w: w + 2 * GAP, h: h + 2 * GAP });
    placed.push({ el, x, y, w, h });
  };

  for (const block of blocks.filter((b) => b.pin)) {
    const el = makeBlock(block, block.size);
    field.appendChild(el);
    const w = el.offsetWidth, h = el.offsetHeight;
    const x = Math.max(visibleLeft, Math.min(block.pin![0], visibleRight - w));
    let y = block.pin![1];
    while (collides(x, y, w, h) && y < 1200) y += 14; // e.g. below the nav on a phone
    place(el, x, y, w, h, block.size);
  }

  // Notes appear once; formulas are dealt from a reshuffled deck so repeats stay far apart.
  const notes = blocks.filter((b) => b.once && !b.pin).sort(() => rand() - 0.5);
  const pool = blocks.filter((b) => !b.once);
  const noteEvery = Math.max(3, Math.floor((pool.length * 3) / (notes.length + 1)));
  let deck: FieldBlock[] = [];
  let lowest = 0, misses = 0;
  for (let k = 0; misses < 60 && k < 900; k++) {
    const fromNotes = notes.length > 0 && k % noteEvery === 0;
    if (!deck.length) deck = pool.slice().sort(() => rand() - 0.5);
    const block = fromNotes ? notes.shift()! : deck.pop()!;
    const size = block.size || SIZES[Math.floor(rand() * SIZES.length)];
    const el = makeBlock(block, size);
    el.style.left = "-9999px";
    field.appendChild(el);
    const w = el.offsetWidth, h = el.offsetHeight;
    // formulas may run off the edge like a page that continues; notes may not
    const minX = fromNotes ? visibleLeft : -0.18 * w;
    const span = fromNotes ? Math.max(0, visibleRight - visibleLeft - w) : CANVAS_WIDTH - 0.64 * w;
    let spot: { x: number; y: number } | null = null;
    for (let y = Math.max(0, lowest - 420); y < height - h && !spot; y += 14) {
      for (let tries = 0; tries < 18; tries++) {
        const x = minX + rand() * span;
        if (!collides(x, y, w, h)) { spot = { x, y }; break; }
      }
    }
    if (!spot) {
      el.remove();
      misses++;
      if (fromNotes) notes.push(block);
      continue;
    }
    misses = 0;
    place(el, spot.x, spot.y, w, h, size);
    lowest = Math.max(lowest, spot.y);
  }
  return placed;
}
