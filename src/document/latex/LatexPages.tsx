// The pages of a LaTeX document as TeX made them: each page's SVG (from the same DVI/XDV as the PDF) put in at
// its paper size × zoom, only near the visible part; a picture of the PDF page where the SVG cannot show it. A
// page keeps what it showed until its new content is there (no flashing while TeX runs). A click asks for the
// source line (SyncTeX); the marker shows where the caret's line is.
import { memo, useCallback, useEffect, useLayoutEffect, useRef, type CSSProperties } from "react";
import type { PageMeta } from "./texService";

export type PageContent = { kind: "svg"; hash: string; svg: string } | { kind: "png"; hash: string; url: string };

export interface PageMarker {
  page: number;
  /** bp from the page's top left */
  x: number;
  y: number;
  width: number;
  height: number;
  /** "sync": the caret's place; "pending": an edit TeX has not typeset yet */
  kind: "sync" | "pending";
}

interface Props {
  pages: PageMeta[];
  /** what each page shows (1-based index − 1); missing: blank until loaded */
  contents: Array<PageContent | undefined>;
  zoom: number;
  marker: PageMarker | null;
  /** a page point clicked (bp) */
  onPageClick(page: number, x: number, y: number): void;
  /** the pages near the visible part (1-based), when they change */
  onVisible(pages: number[]): void;
  onZoom(zoom: number): void;
  /** scroll so that this page point is in view (changes with `nonce`) */
  reveal: { page: number; y: number; nonce: number; ifHidden?: boolean } | null;
  /** fits the width (for "fit width") */
  onWidth?(width: number): void;
  busy: boolean;
}

const GAP = 14;

const PageBody = memo(function PageBody({ content }: { content: PageContent | undefined }) {
  if (!content) return null;
  if (content.kind === "png") return <img className="vtx-tex-page-picture" src={content.url} alt="" draggable={false} />;
  return <div className="vtx-tex-page-svg" dangerouslySetInnerHTML={{ __html: content.svg }} />;
}, (a, b) => a.content?.hash === b.content?.hash && a.content?.kind === b.content?.kind);

export function LatexPages({ pages, contents, zoom, marker, onPageClick, onVisible, onZoom, reveal, onWidth, busy }: Props) {
  const scrollRef = useRef<HTMLDivElement>(null);
  const visibleRef = useRef(new Set<number>());
  const onVisibleRef = useRef(onVisible);
  onVisibleRef.current = onVisible;
  const down = useRef<{ x: number; y: number } | null>(null);

  // pages near the view: IntersectionObserver on each page frame
  useEffect(() => {
    const root = scrollRef.current;
    if (!root) return;
    const visible = visibleRef.current;
    visible.clear();
    let timer = 0;
    const observer = new IntersectionObserver(
      (entries) => {
        for (const e of entries) {
          const n = Number((e.target as HTMLElement).dataset.page);
          if (e.isIntersecting) visible.add(n);
          else visible.delete(n);
        }
        window.clearTimeout(timer);
        timer = window.setTimeout(() => onVisibleRef.current([...visible].sort((a, b) => a - b)), 30);
      },
      { root, rootMargin: "800px 0px" },
    );
    root.querySelectorAll<HTMLElement>(".vtx-tex-page").forEach((el) => observer.observe(el));
    return () => {
      window.clearTimeout(timer);
      observer.disconnect();
    };
  }, [pages.length, zoom]);

  // the width for "fit width"
  useEffect(() => {
    const el = scrollRef.current;
    if (!el || !onWidth) return;
    const report = () => onWidth(el.clientWidth);
    report();
    const ro = new ResizeObserver(report);
    ro.observe(el);
    return () => ro.disconnect();
  }, [onWidth]);

  // zoom keeps the point at the middle of the view where it was
  const lastZoom = useRef(zoom);
  useLayoutEffect(() => {
    const el = scrollRef.current;
    if (!el || lastZoom.current === zoom) return;
    const ratio = zoom / lastZoom.current;
    const middle = el.scrollTop + el.clientHeight / 2;
    el.scrollTop = middle * ratio - el.clientHeight / 2;
    lastZoom.current = zoom;
  }, [zoom]);

  useEffect(() => {
    const el = scrollRef.current;
    if (!el || !reveal) return;
    const frame = el.querySelector<HTMLElement>(`.vtx-tex-page[data-page="${reveal.page}"]`);
    if (!frame) return;
    const target = frame.offsetTop + reveal.y * zoom;
    if (reveal.ifHidden && target > el.scrollTop + 40 && target < el.scrollTop + el.clientHeight - 40) return;
    el.scrollTo({ top: Math.max(0, target - el.clientHeight / 3), behavior: reveal.ifHidden ? "smooth" : "auto" });
    // the scroll position depends on the zoom only at the time of the request
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [reveal?.nonce]);

  const onWheel = useCallback(
    (event: React.WheelEvent) => {
      if (!(event.ctrlKey || event.metaKey)) return;
      event.preventDefault();
      onZoom(Math.min(4, Math.max(0.3, zoom * Math.exp(-event.deltaY / 400))));
    },
    [onZoom, zoom],
  );

  // ctrl+wheel must be able to prevent the browser's own zoom: a non-passive listener
  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const block = (e: WheelEvent) => {
      if (e.ctrlKey || e.metaKey) e.preventDefault();
    };
    el.addEventListener("wheel", block, { passive: false });
    return () => el.removeEventListener("wheel", block);
  }, []);

  return (
    <div
      ref={scrollRef}
      className={"vtx-tex-pages" + (busy ? " is-busy" : "")}
      onWheel={onWheel}
      aria-label="Pages"
    >
      {pages.map((page, i) => {
        const n = i + 1;
        const style: CSSProperties = { width: page.width * zoom, height: page.height * zoom, marginBottom: GAP };
        return (
          <div
            key={n}
            className={"vtx-tex-page" + (contents[i] ? "" : " is-empty")}
            data-page={n}
            data-hash={contents[i]?.hash}
            data-current={contents[i]?.hash === page.hash ? "true" : undefined}
            style={style}
            onMouseDown={(event) => {
              down.current = { x: event.clientX, y: event.clientY };
            }}
            onMouseUp={(event) => {
              const start = down.current;
              down.current = null;
              if (!start || Math.hypot(event.clientX - start.x, event.clientY - start.y) > 4) return;
              const selection = window.getSelection();
              if (selection && !selection.isCollapsed) return;
              const rect = event.currentTarget.getBoundingClientRect();
              onPageClick(n, ((event.clientX - rect.left) / rect.width) * page.width, ((event.clientY - rect.top) / rect.height) * page.height);
            }}
          >
            <PageBody content={contents[i]} />
            {marker && marker.page === n && (
              <div
                className={"vtx-tex-marker is-" + marker.kind}
                style={{
                  left: marker.x * zoom - 3,
                  top: marker.y * zoom - 2,
                  width: Math.max(marker.width * zoom, 6) + 6,
                  height: Math.max(marker.height * zoom, 6) + 4,
                }}
              />
            )}
          </div>
        );
      })}
    </div>
  );
}
