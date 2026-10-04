import type { PlacedBlock } from "./formulaField";

const LENS_RADIUS = 120; // screen px
const TOUCH_HOLD_MS = 1600;

/**
 * Behaviour of the formula field once it is laid out:
 * - "compile": every block first shows its LaTeX source and typesets itself when it comes into
 *   view (the hero ripples outward from the logo on first load);
 * - "lens": around the pointer the field turns back into source, so the page itself shows the
 *   bijection the editor is built on. On touch, pressing the background opens the lens briefly.
 * Returns a cleanup function.
 */
export function startFieldEffects(
  page: HTMLElement,
  field: HTMLElement,
  ring: HTMLElement,
  blocks: PlacedBlock[],
  { compile }: { compile: boolean },
) {
  const cleanups: (() => void)[] = [];
  const reduce = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  const scale = () => field.getBoundingClientRect().width / field.offsetWidth || 1;

  // ---- compile on first sight
  if (compile && !reduce) {
    const heroBottom = window.innerHeight;
    const s = scale(), fieldTop = field.getBoundingClientRect().top + window.scrollY;
    const origin = { x: field.offsetWidth / 2, y: (heroBottom * 0.47) / s };
    const observer = new IntersectionObserver((entries) => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        const el = entry.target as HTMLElement;
        observer.unobserve(el);
        const block = blocks.find((b) => b.el === el)!;
        const inHero = block.y * s + fieldTop < heroBottom && window.scrollY < 40;
        // hero: ripple out from the logo while it is being written; later: a short stagger
        const delay = inHero
          ? 500 + Math.hypot(block.x + block.w / 2 - origin.x, block.y + block.h / 2 - origin.y) * 1.1
          : 120 + Math.random() * 380;
        window.setTimeout(() => el.classList.remove("is-raw"), delay);
      }
    }, { rootMargin: "0px 0px -8% 0px" });
    for (const block of blocks) {
      block.el.classList.add("is-raw");
      observer.observe(block.el);
    }
    cleanups.push(() => observer.disconnect());
  }

  // ---- lens
  let lensed = new Set<HTMLElement>();
  let frame = 0, last: { x: number; y: number } | null = null, hideTimer = 0;

  const clear = () => {
    for (const el of lensed) el.classList.remove("is-lensed");
    lensed = new Set();
    ring.classList.remove("is-on");
  };
  const paint = () => {
    frame = 0;
    if (!last) return clear();
    const rect = field.getBoundingClientRect(), s = rect.width / field.offsetWidth;
    const fx = (last.x - rect.left) / s, fy = (last.y - rect.top) / s, r = LENS_RADIUS / s;
    const next = new Set<HTMLElement>();
    for (const b of blocks) {
      if (fx + r < b.x || fx - r > b.x + b.w || fy + r < b.y || fy - r > b.y + b.h) continue;
      next.add(b.el);
      b.el.style.setProperty("--lx", `${(fx - b.x).toFixed(1)}px`);
      b.el.style.setProperty("--ly", `${(fy - b.y).toFixed(1)}px`);
      b.el.style.setProperty("--lr", `${r.toFixed(1)}px`);
    }
    for (const el of lensed) if (!next.has(el)) el.classList.remove("is-lensed");
    for (const el of next) el.classList.add("is-lensed");
    lensed = next;
    ring.style.transform = `translate(${last.x}px, ${last.y}px)`;
    ring.classList.add("is-on");
  };
  const aim = (x: number, y: number, target: EventTarget | null) => {
    // The lens belongs to the paper behind the sheets, not to the sheets or the nav.
    const over = target instanceof Element && target.closest("[data-clear], .landing-nav, a, button");
    last = over ? null : { x, y };
    if (!frame) frame = requestAnimationFrame(paint);
  };

  const onMove = (event: PointerEvent) => {
    if (event.pointerType !== "mouse") return;
    aim(event.clientX, event.clientY, event.target);
  };
  const onLeave = () => { last = null; if (!frame) frame = requestAnimationFrame(paint); };
  const onDown = (event: PointerEvent) => {
    if (event.pointerType === "mouse") return;
    window.clearTimeout(hideTimer);
    aim(event.clientX, event.clientY, event.target);
    hideTimer = window.setTimeout(onLeave, TOUCH_HOLD_MS);
  };
  const onScroll = () => {
    // Content moves under a still mouse; re-aim at the element now under it.
    if (!last) return;
    aim(last.x, last.y, document.elementFromPoint(last.x, last.y));
  };

  if (!reduce) {
    page.addEventListener("pointermove", onMove);
    page.addEventListener("pointerleave", onLeave);
    page.addEventListener("pointerdown", onDown);
    window.addEventListener("scroll", onScroll, { passive: true });
    cleanups.push(() => {
      page.removeEventListener("pointermove", onMove);
      page.removeEventListener("pointerleave", onLeave);
      page.removeEventListener("pointerdown", onDown);
      window.removeEventListener("scroll", onScroll);
      cancelAnimationFrame(frame);
      window.clearTimeout(hideTimer);
      clear();
    });
  }
  return () => cleanups.forEach((fn) => fn());
}
