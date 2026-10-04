// Hand-drawn marks for the landing page: brush swashes behind buttons, sketchy outlines, margin
// arrows, circling loops and ink blots. Deterministic: every shape comes from a fixed seed.

function rng(seed) {
  return () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const f = (v) => +v.toFixed(1);
const poly = (pts) => `M${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join("L")}Z`;
// smooth 1-D noise from a few sines with random phases
function wobble(rand, amp, waves = 3) {
  const parts = Array.from({ length: waves }, (_, i) => [rand() * 6.28, 1 + i * 1.7 + rand(), amp / (i + 1)]);
  return (t) => parts.reduce((sum, [phase, freq, a]) => sum + a * Math.sin(phase + freq * t * 6.28), 0);
}

/**
 * A broad brush stroke for a button background, in a 300×60 box stretched to the button.
 * Returns the solid body and a dry-brush tail that the page filters separately.
 */
export function swash(seed) {
  const rand = rng(seed), n = 90, top = [], bot = [];
  const wt = wobble(rand, 2.2), wb = wobble(rand, 2.2), lift = wobble(rand, 1.6, 2);
  for (let i = 0; i <= n; i++) {
    const t = i / n, x = -4 + t * 310;
    // round, loaded start; full body; thinning dry end
    const load = Math.pow(Math.sin((Math.PI / 2) * Math.min(1, t * 4 + 0.12)), 0.4);
    const fade = 1 - 0.32 * Math.max(0, t - 0.72) / 0.28;
    const half = 27 * load * fade;
    const mid = 31 + lift(t) - 2 * t;
    top.push([x, mid - half + wt(t)]);
    bot.unshift([x, mid + half + wb(t)]);
  }
  const d = poly([...top, ...bot]);
  return `<svg viewBox="0 0 300 60" preserveAspectRatio="none" aria-hidden="true"><path class="sw-body" d="${d}" mask="url(#vt-btn-solid)" filter="url(#vt-ink-solid)"/><path class="sw-dry" d="${d}" mask="url(#vt-btn-dry)" filter="url(#vt-dry-brush)"/></svg>`;
}

/** A sketchy rectangle drawn twice, slightly off, in a 200×56 box (stroke does not scale). */
export function sketchBox(seed) {
  const rand = rng(seed);
  const loop = (o) => {
    const j = () => (rand() - 0.5) * 3.2;
    const x0 = 2 + j() + o, y0 = 3 + j(), x1 = 198 + j() - o, y1 = 53 + j();
    // start a little in, overshoot the end: the hand never closes a box exactly
    return `M${f(x0 + 10)} ${f(y0 + j() * 0.4)}L${f(x1 + j())} ${f(y0 + j() * 0.5)}L${f(x1 + j() * 0.6)} ${f(y1 + j())}L${f(x0 + j())} ${f(y1 + j() * 0.5)}L${f(x0 + j() * 0.6)} ${f(y0 - 1)}L${f(x0 + 24 + j() * 2)} ${f(y0 + j() * 0.4)}`;
  };
  return `<svg viewBox="0 0 200 56" preserveAspectRatio="none" aria-hidden="true"><path pathLength="1" d="${loop(0)}"/><path pathLength="1" d="${loop(1.5)}"/></svg>`;
}

/**
 * A margin arrow from (0,0) to `to`, bowed by `bend`, with a two-stroke head; drawn as strokes
 * in a padded box. Each path has pathLength=1 so the page can draw it in with dashoffset.
 */
export function arrow(seed, to, bend) {
  const rand = rng(seed);
  const [tx, ty] = to, len = Math.hypot(tx, ty), nx = -ty / len, ny = tx / len;
  const c1 = [tx * 0.3 + nx * bend, ty * 0.3 + ny * bend];
  const c2 = [tx * 0.75 + nx * bend * 0.8, ty * 0.75 + ny * bend * 0.8];
  const shaft = `M0 0C${f(c1[0])} ${f(c1[1])} ${f(c2[0])} ${f(c2[1])} ${f(tx)} ${f(ty)}`;
  // head follows the final tangent
  const dx = tx - c2[0], dy = ty - c2[1], dl = Math.hypot(dx, dy), ux = dx / dl, uy = dy / dl;
  const head = (side) => {
    const a = side * (0.5 + rand() * 0.12), s = 13 + rand() * 4;
    const hx = -(ux * Math.cos(a) - uy * Math.sin(a)) * s, hy = -(ux * Math.sin(a) + uy * Math.cos(a)) * s;
    return `M${f(tx + hx)} ${f(ty + hy)}L${f(tx)} ${f(ty)}`;
  };
  const xs = [0, tx, c1[0], c2[0]], ys = [0, ty, c1[1], c2[1]], P = 18;
  const x0 = Math.min(...xs) - P, y0 = Math.min(...ys) - P;
  const w = Math.max(...xs) - Math.min(...xs) + 2 * P, h = Math.max(...ys) - Math.min(...ys) + 2 * P;
  return {
    w: f(w), h: f(h),
    svg: `<svg viewBox="${f(x0)} ${f(y0)} ${f(w)} ${f(h)}" width="${f(w)}" height="${f(h)}" aria-hidden="true"><path class="ar-shaft" pathLength="1" d="${shaft}"/><path class="ar-head" pathLength="1" d="${head(1)}"/><path class="ar-head" pathLength="1" d="${head(-1)}"/></svg>`,
  };
}

/**
 * A quick loop drawn around something: a little more than one turn, in a 200×80 box
 * (or a 200×200 box when `round`, for the lens).
 */
export function circleLoop(seed, round = false) {
  const rand = rng(seed), pts = [], turns = 1.16, n = round ? 120 : 70, wr = wobble(rand, 0.05, 2);
  const ry = round ? 94 : 34, cy = round ? 100 : 40;
  const start = -2.3 + rand() * 0.3;
  for (let i = 0; i <= n; i++) {
    const t = i / n, a = start + t * turns * 6.283;
    const grow = 1 + 0.08 * t + wr(t);
    pts.push([100 + 94 * grow * Math.cos(a), cy + ry * grow * Math.sin(a) - 4 * t]);
  }
  return `<svg viewBox="${round ? "-10 -10 220 220" : "-8 -8 216 96"}" preserveAspectRatio="none" aria-hidden="true"><path pathLength="1" d="M${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join("L")}"/></svg>`;
}

/** An ink blot with a few satellite drops, for the press splash; 100×100 box. */
export function blot(seed) {
  const rand = rng(seed), n = 40, pts = [], w = wobble(rand, 0.18, 3);
  for (let i = 0; i < n; i++) {
    const a = (i / n) * 6.283, r = 22 * (1 + w(i / n));
    pts.push([50 + r * Math.cos(a), 50 + r * Math.sin(a)]);
  }
  const drops = Array.from({ length: 6 }, () => {
    const a = rand() * 6.283, d = 30 + rand() * 16, r = 1.5 + rand() * 3.5;
    return `<circle cx="${f(50 + d * Math.cos(a))}" cy="${f(50 + d * Math.sin(a))}" r="${f(r)}"/>`;
  }).join("");
  return `<svg viewBox="0 0 100 100" aria-hidden="true"><g filter="url(#vt-ink-solid)"><path d="${poly(pts)}"/>${drops}</g></svg>`;
}

/** Masks for the swash: solid body until ~70%, then the dry tail. */
export const swashDefs = `
<linearGradient id="vt-btn-g"><stop offset=".8" stop-color="#fff"/><stop offset=".97" stop-color="#000"/></linearGradient>
<linearGradient id="vt-btn-gi"><stop offset=".8" stop-color="#000"/><stop offset=".97" stop-color="#fff"/></linearGradient>
<mask id="vt-btn-solid" maskContentUnits="objectBoundingBox" x="-.1" y="-.3" width="1.2" height="1.6"><rect x="-.1" y="-.3" width="1.2" height="1.6" fill="url(#vt-btn-g)"/></mask>
<mask id="vt-btn-dry" maskContentUnits="objectBoundingBox" x="-.1" y="-.3" width="1.2" height="1.6"><rect x="-.1" y="-.3" width="1.2" height="1.6" fill="url(#vt-btn-gi)"/></mask>`;
