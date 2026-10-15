/**
 * Breaks two-value layout oscillations.
 *
 * A component that sizes its content from its own measured width can change
 * the height of a scroll container; when that toggles a non-overlay scroll bar
 * the width changes back, and the next measurement undoes the previous one.
 * `scrollbar-gutter: stable` prevents that, but older macOS WebKit does not
 * support it. Feed every measured value through the guard: when it sees
 * A -> B -> A within `windowMs`, with A and B no further apart than
 * `tolerance` (a scroll bar is about 15px), it settles on the smaller value
 * and keeps returning it while measurements stay on that pair. Any other value
 * means a real resize and clears the lock.
 */
export function createLayoutOscillationGuard(
  tolerance = 24,
  windowMs = 1000,
) {
  let history: Array<{ value: number; at: number }> = [];
  let lockedPair: [number, number] | null = null;

  return (value: number) => {
    const same = (a: number, b: number) => Math.abs(a - b) < 0.5;
    if (lockedPair) {
      if (same(value, lockedPair[0]) || same(value, lockedPair[1])) {
        return Math.min(lockedPair[0], lockedPair[1]);
      }
      lockedPair = null;
      history = [];
    }
    const last = history[history.length - 1];
    if (!last || !same(last.value, value)) {
      history.push({ value, at: performance.now() });
      if (history.length > 3) history.shift();
    }
    if (history.length === 3) {
      const [first, second, third] = history;
      const gap = Math.abs(first.value - second.value);
      if (
        same(first.value, third.value) &&
        gap <= tolerance &&
        third.at - first.at <= windowMs
      ) {
        lockedPair = [first.value, second.value];
        return Math.min(first.value, second.value);
      }
    }
    return value;
  };
}
