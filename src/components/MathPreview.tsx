import { memo, useLayoutEffect, useMemo, useRef } from "react";
import { convertVisualTexLatexToMarkup } from "../editor/mathLiveIntegralCompatibility";
import { useCustomSymbolRevision } from "../math/customSymbolReact";

interface MathPreviewProps {
  latex: string;
  className?: string;
  showPlaceholders?: boolean;
  fit?: boolean;
  fluidHeight?: boolean;
  intrinsicWidth?: boolean;
  intrinsicMaxWidth?: number;
  minimumFluidScale?: number;
  maximumFluidScale?: number;
  maximumFitScale?: number;
  fitInsetRatio?: number;
  minimumFluidHeight?: number;
  maximumFluidHeight?: number;
  fluidVerticalPadding?: number;
  staticLayout?: boolean;
  onMeasure?: (size: { width: number; height: number }) => void;
}

const defaultFitInsetRatio = 0.9;
const minimumFluidFitScale = 0.1;
const defaultMaximumFitScale = 8;
const visiblePlaceholderLatex =
  "\\htmlClass{visualtex-tile-placeholder}{\\phantom{\\rule{0.40em}{0.66em}}}";
const mathPreviewMarkupCache = new Map<string, string>();
const mathPreviewMarkupCacheLimit = 1024;

function cachedPreviewMarkup(latex: string, customSymbolRevision: number) {
  const cacheKey = `${customSymbolRevision}\u0000${latex}`;
  const cached = mathPreviewMarkupCache.get(cacheKey);
  if (cached !== undefined) return cached;
  const markup = convertVisualTexLatexToMarkup(latex, { defaultMode: "math" });
  if (mathPreviewMarkupCache.size >= mathPreviewMarkupCacheLimit) {
    const oldestKey = mathPreviewMarkupCache.keys().next().value;
    if (typeof oldestKey === "string") mathPreviewMarkupCache.delete(oldestKey);
  }
  mathPreviewMarkupCache.set(cacheKey, markup);
  return markup;
}

export function latexWithVisiblePlaceholders(latex: string) {
  if (!latex.includes("\\placeholder")) return latex;

  let cursor = 0;
  let rendered = "";
  while (cursor < latex.length) {
    const commandStart = latex.indexOf("\\placeholder", cursor);
    if (commandStart < 0) {
      rendered += latex.slice(cursor);
      break;
    }

    rendered += latex.slice(cursor, commandStart);
    let commandEnd = commandStart + "\\placeholder".length;
    while (/\s/.test(latex[commandEnd] ?? "")) commandEnd += 1;

    if (latex[commandEnd] === "[") {
      let bracketDepth = 1;
      commandEnd += 1;
      while (commandEnd < latex.length && bracketDepth > 0) {
        if (latex[commandEnd] === "[") bracketDepth += 1;
        if (latex[commandEnd] === "]") bracketDepth -= 1;
        commandEnd += 1;
      }
      while (/\s/.test(latex[commandEnd] ?? "")) commandEnd += 1;
    }

    if (latex[commandEnd] !== "{") {
      rendered += "\\placeholder";
      cursor = commandStart + "\\placeholder".length;
      continue;
    }

    let braceDepth = 1;
    commandEnd += 1;
    while (commandEnd < latex.length && braceDepth > 0) {
      if (latex[commandEnd] === "{") braceDepth += 1;
      if (latex[commandEnd] === "}") braceDepth -= 1;
      commandEnd += 1;
    }
    if (braceDepth > 0) {
      rendered += latex.slice(commandStart);
      break;
    }

    rendered += visiblePlaceholderLatex;
    cursor = commandEnd;
  }
  return rendered;
}

function MathPreviewComponent({
  latex,
  className = "",
  showPlaceholders = false,
  fit = false,
  fluidHeight = false,
  intrinsicWidth = false,
  intrinsicMaxWidth = 280,
  minimumFluidScale = minimumFluidFitScale,
  maximumFluidScale = 1.35,
  maximumFitScale = defaultMaximumFitScale,
  fitInsetRatio = defaultFitInsetRatio,
  minimumFluidHeight = 52,
  maximumFluidHeight = 168,
  fluidVerticalPadding = 20,
  staticLayout = false,
  onMeasure,
}: MathPreviewProps) {
  const hostRef = useRef<HTMLSpanElement>(null);
  const contentRef = useRef<HTMLSpanElement>(null);
  const onMeasureRef = useRef(onMeasure);
  onMeasureRef.current = onMeasure;
  const customSymbolRevision = useCustomSymbolRevision();
  const previewLatex = useMemo(
    () => (showPlaceholders ? latexWithVisiblePlaceholders(latex) : latex),
    [latex, showPlaceholders],
  );
  const markup = useMemo(
    () => cachedPreviewMarkup(previewLatex, customSymbolRevision),
    [customSymbolRevision, previewLatex],
  );

  useLayoutEffect(() => {
    const host = hostRef.current;
    const content = contentRef.current;
    if (!host || !content) return;

    if (staticLayout) {
      host.style.removeProperty("--math-preview-fluid-height");
      host.style.removeProperty("--math-preview-intrinsic-width");
      let staticAnimationFrame = 0;
      let disposed = false;
      const measureStatic = () => {
        staticAnimationFrame = 0;
        content.style.setProperty("--math-preview-fit-scale", "1");
        let scale = 1;
        if (fit) {
          const visualRoot =
            content.querySelector<HTMLElement>(".ML__latex") ?? content;
          const visualRect = visualRoot.getBoundingClientRect();
          const contentRect = content.getBoundingClientRect();
          const naturalWidth = Math.max(
            1,
            content.scrollWidth,
            contentRect.width,
            visualRect.width,
          );
          const naturalHeight = Math.max(
            1,
            content.offsetHeight,
            contentRect.height,
            visualRect.height,
          );
          const containedScale = Math.min(
            Math.max(1, host.clientWidth * fitInsetRatio) / naturalWidth,
            Math.max(1, host.clientHeight * fitInsetRatio) / naturalHeight,
          );
          scale = Math.max(
            Number.EPSILON,
            Math.min(0.92, maximumFitScale, containedScale),
          );
        }
        content.style.setProperty(
          "--math-preview-fit-scale",
          scale.toFixed(4),
        );
        host.dataset.fitReady = "static";
        host.dataset.fitScale = scale.toFixed(4);
      };
      const scheduleStaticMeasure = () => {
        if (disposed) return;
        if (staticAnimationFrame) cancelAnimationFrame(staticAnimationFrame);
        staticAnimationFrame = requestAnimationFrame(measureStatic);
      };
      measureStatic();
      scheduleStaticMeasure();
      void document.fonts?.ready.then(scheduleStaticMeasure);
      return () => {
        disposed = true;
        if (staticAnimationFrame) cancelAnimationFrame(staticAnimationFrame);
      };
    }

    let animationFrame = 0;
    const measure = () => {
      animationFrame = 0;
      content.style.setProperty("--math-preview-fit-scale", "1");
      const visualRoot =
        content.querySelector<HTMLElement>(".ML__latex") ?? content;
      const visualRect = visualRoot.getBoundingClientRect();
      const contentRect = content.getBoundingClientRect();
      const naturalWidth = Math.max(
        1,
        content.scrollWidth,
        contentRect.width,
        visualRect.width,
      );
      const naturalHeight = Math.max(
        1,
        content.offsetHeight,
        contentRect.height,
        visualRect.height,
      );
      onMeasureRef.current?.({
        width: Math.max(1, content.offsetWidth),
        height: Math.max(1, content.offsetHeight),
      });
      if (intrinsicWidth) {
        const desiredWidth = Math.min(
          intrinsicMaxWidth,
          Math.max(34, Math.ceil(naturalWidth)),
        );
        host.style.setProperty(
          "--math-preview-intrinsic-width",
          `${desiredWidth}px`,
        );
      } else {
        host.style.removeProperty("--math-preview-intrinsic-width");
      }

      if (!fit) {
        content.style.setProperty("--math-preview-fit-scale", "1");
        host.style.removeProperty("--math-preview-fluid-height");
        host.dataset.fitReady = "false";
        host.dataset.fitScale = "1";
        return;
      }

      const availableWidth = Math.max(1, host.clientWidth * fitInsetRatio);
      let scale = 1;

      if (fluidHeight) {
        const availableHeight = Math.max(
          1,
          maximumFluidHeight - fluidVerticalPadding,
        );
        const widthScale = availableWidth / naturalWidth;
        const heightScale = availableHeight / naturalHeight;
        scale = Math.max(
          minimumFluidScale,
          Math.min(maximumFluidScale, widthScale, heightScale),
        );
        const renderedHeight = naturalHeight * scale;
        const rowHeight = Math.min(
          maximumFluidHeight,
          Math.max(
            minimumFluidHeight,
            Math.ceil(renderedHeight + fluidVerticalPadding),
          ),
        );
        host.style.setProperty("--math-preview-fluid-height", `${rowHeight}px`);
      } else {
        host.style.removeProperty("--math-preview-fluid-height");
        const availableHeight = Math.max(1, host.clientHeight * fitInsetRatio);
        const containedScale = Math.min(
          availableWidth / naturalWidth,
          availableHeight / naturalHeight,
        );
        scale = Math.max(
          Number.EPSILON,
          Math.min(maximumFitScale, containedScale),
        );
      }

      content.style.setProperty(
        "--math-preview-fit-scale",
        scale.toFixed(4),
      );
      host.dataset.fitReady = "true";
      host.dataset.fitScale = scale.toFixed(4);
    };
    const scheduleMeasure = () => {
      if (animationFrame) cancelAnimationFrame(animationFrame);
      animationFrame = requestAnimationFrame(measure);
    };

    measure();
    scheduleMeasure();
    void document.fonts?.ready.then(scheduleMeasure);
    const resizeObserver = new ResizeObserver(scheduleMeasure);
    resizeObserver.observe(host);
    resizeObserver.observe(content);
    const observedVisualRoot = content.querySelector<HTMLElement>(".ML__latex");
    if (observedVisualRoot) resizeObserver.observe(observedVisualRoot);

    return () => {
      if (animationFrame) cancelAnimationFrame(animationFrame);
      resizeObserver.disconnect();
    };
  }, [
    fit,
    fluidHeight,
    intrinsicMaxWidth,
    intrinsicWidth,
    markup,
    fluidVerticalPadding,
    maximumFluidHeight,
    maximumFluidScale,
    maximumFitScale,
    fitInsetRatio,
    minimumFluidHeight,
    minimumFluidScale,
    staticLayout,
  ]);

  return (
    <span
      ref={hostRef}
      className={"math-preview " + className}
      aria-hidden="true"
      data-fit={fit ? "contain" : "none"}
      data-show-placeholders={showPlaceholders ? "true" : "false"}
      data-fluid-height={fluidHeight ? "true" : "false"}
      data-intrinsic-width={intrinsicWidth ? "true" : "false"}
      data-static-layout={staticLayout ? "true" : "false"}
    >
      <span
        ref={contentRef}
        className="math-preview-fit-content"
        dangerouslySetInnerHTML={{ __html: markup }}
      />
    </span>
  );
}

export const MathPreview = memo(MathPreviewComponent);
