import { useLayoutEffect, useMemo, useRef } from "react";
import { convertLatexToMarkup } from "mathlive";

interface MathPreviewProps {
  latex: string;
  className?: string;
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
  onMeasure?: (size: { width: number; height: number }) => void;
}

const defaultFitInsetRatio = 0.9;
const minimumFluidFitScale = 0.1;
const defaultMaximumFitScale = 8;

export function MathPreview({
  latex,
  className = "",
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
  onMeasure,
}: MathPreviewProps) {
  const hostRef = useRef<HTMLSpanElement>(null);
  const contentRef = useRef<HTMLSpanElement>(null);
  const onMeasureRef = useRef(onMeasure);
  onMeasureRef.current = onMeasure;
  const markup = useMemo(
    () => convertLatexToMarkup(latex, { defaultMode: "math" }),
    [latex],
  );

  useLayoutEffect(() => {
    const host = hostRef.current;
    const content = contentRef.current;
    if (!host || !content) return;

    let animationFrame = 0;
    const measure = () => {
      animationFrame = 0;
      const naturalWidth = Math.max(1, content.offsetWidth);
      const naturalHeight = Math.max(1, content.offsetHeight);
      onMeasureRef.current?.({ width: naturalWidth, height: naturalHeight });
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
        scale = Math.max(
          minimumFluidScale,
          Math.min(maximumFluidScale, availableWidth / naturalWidth),
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
        // Non-fluid previews are strict contain boxes: never impose a visual
        // minimum that could make a tall integral, sum, or matrix overflow.
        // The caller may cap upscaling (the formula toolbar uses 1) while
        // oversized content is always allowed to shrink as far as required.
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

    scheduleMeasure();
    void document.fonts?.ready.then(scheduleMeasure);
    const resizeObserver = new ResizeObserver(scheduleMeasure);
    resizeObserver.observe(host);

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
  ]);

  return (
    <span
      ref={hostRef}
      className={"math-preview " + className}
      aria-hidden="true"
      data-fit={fit ? "contain" : "none"}
      data-fluid-height={fluidHeight ? "true" : "false"}
      data-intrinsic-width={intrinsicWidth ? "true" : "false"}
    >
      <span
        ref={contentRef}
        className="math-preview-fit-content"
        dangerouslySetInnerHTML={{ __html: markup }}
      />
    </span>
  );
}
