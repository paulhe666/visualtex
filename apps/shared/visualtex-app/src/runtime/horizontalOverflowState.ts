/**
 * Marks a horizontally scrollable strip with data-overflow-start / -end so CSS
 * can fade the clipped edge, and lets a vertical mouse wheel scroll it
 * sideways. Only data attributes change; they drive a mask, never layout, so
 * this observer cannot feed back into the size it measures.
 */
export function trackHorizontalOverflow(strip: HTMLElement) {
  let frame = 0;
  const setFlag = (name: "overflowStart" | "overflowEnd", value: boolean) => {
    const next = value ? "true" : undefined;
    if (strip.dataset[name] === next) return;
    if (next) strip.dataset[name] = next;
    else delete strip.dataset[name];
  };
  const update = () => {
    if (frame) return;
    frame = window.requestAnimationFrame(() => {
      frame = 0;
      const maximum = strip.scrollWidth - strip.clientWidth;
      setFlag("overflowStart", maximum > 1 && strip.scrollLeft > 1);
      setFlag("overflowEnd", maximum > 1 && strip.scrollLeft < maximum - 1);
    });
  };
  const resizeObserver = new ResizeObserver(update);
  const observeChildren = () => {
    resizeObserver.disconnect();
    resizeObserver.observe(strip);
    for (const child of Array.from(strip.children)) resizeObserver.observe(child);
  };
  const mutationObserver = new MutationObserver(() => {
    observeChildren();
    update();
  });
  const handleWheel = (event: WheelEvent) => {
    if (event.ctrlKey || event.metaKey) return;
    if (Math.abs(event.deltaX) >= Math.abs(event.deltaY)) return;
    const maximum = strip.scrollWidth - strip.clientWidth;
    if (maximum <= 1) return;
    event.preventDefault();
    strip.scrollLeft = Math.max(0, Math.min(maximum, strip.scrollLeft + event.deltaY));
  };

  observeChildren();
  mutationObserver.observe(strip, { childList: true });
  strip.addEventListener("scroll", update, { passive: true });
  strip.addEventListener("wheel", handleWheel, { passive: false });
  update();

  return () => {
    window.cancelAnimationFrame(frame);
    resizeObserver.disconnect();
    mutationObserver.disconnect();
    strip.removeEventListener("scroll", update);
    strip.removeEventListener("wheel", handleWheel);
    delete strip.dataset.overflowStart;
    delete strip.dataset.overflowEnd;
  };
}
