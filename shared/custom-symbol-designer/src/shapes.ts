import { svgTransformAttribute } from "./matrix.ts";
import type { CustomSymbolVectorShape } from "./symbolTypes.ts";

// Designer shapes as SVG path data (for TeX) and as SVG markup (for the SVG
// code and the exported image).

const n = (value: number) => {
  const rounded = Math.round(value * 1000) / 1000;
  return (Object.is(rounded, -0) ? 0 : rounded).toString();
};

/** The shape's outline as path data in its own coordinates (null: text). */
export function shapePathData(shape: CustomSymbolVectorShape): string | null {
  switch (shape.kind) {
    case "path":
      return shape.d;
    case "line":
      return `M${n(shape.x1)} ${n(shape.y1)}L${n(shape.x2)} ${n(shape.y2)}`;
    case "polygon":
      return shape.points.length
        ? `M${shape.points.map(([x, y]) => `${n(x)} ${n(y)}`).join("L")}Z`
        : null;
    case "circle":
      return ellipsePath(shape.cx, shape.cy, shape.r, shape.r);
    case "ellipse":
      return ellipsePath(shape.cx, shape.cy, shape.rx, shape.ry);
    case "rect": {
      const rx = Math.min(Math.abs(shape.rx ?? shape.ry ?? 0), shape.width / 2);
      const ry = Math.min(Math.abs(shape.ry ?? shape.rx ?? 0), shape.height / 2);
      const { x, y, width, height } = shape;
      if (!rx || !ry) return `M${n(x)} ${n(y)}H${n(x + width)}V${n(y + height)}H${n(x)}Z`;
      return (
        `M${n(x + rx)} ${n(y)}H${n(x + width - rx)}A${n(rx)} ${n(ry)} 0 0 1 ${n(x + width)} ${n(y + ry)}` +
        `V${n(y + height - ry)}A${n(rx)} ${n(ry)} 0 0 1 ${n(x + width - rx)} ${n(y + height)}` +
        `H${n(x + rx)}A${n(rx)} ${n(ry)} 0 0 1 ${n(x)} ${n(y + height - ry)}` +
        `V${n(y + ry)}A${n(rx)} ${n(ry)} 0 0 1 ${n(x + rx)} ${n(y)}Z`
      );
    }
    case "text":
      return null;
  }
}

function ellipsePath(cx: number, cy: number, rx: number, ry: number) {
  return (
    `M${n(cx - rx)} ${n(cy)}A${n(rx)} ${n(ry)} 0 1 0 ${n(cx + rx)} ${n(cy)}` +
    `A${n(rx)} ${n(ry)} 0 1 0 ${n(cx - rx)} ${n(cy)}Z`
  );
}

/** Whether the shape is filled, and its stroke width (0: none), as the canvas paints it. */
export function shapePaint(shape: CustomSymbolVectorShape) {
  return {
    fill: shape.fill ?? shape.kind !== "line",
    strokeWidth: Math.max(0, shape.strokeWidth ?? (shape.kind === "line" ? 50 : 0)),
  };
}

const escapeText = (value: string) =>
  value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/**
 * One shape as SVG markup, drawn as the registered symbol is: the user
 * transform outside, then the crop (so it follows the layer), then the
 * compiler's matrix. A crop becomes a clip path added to `defs`.
 */
export function shapeSvgMarkup(
  shape: CustomSymbolVectorShape,
  paint: string,
  clipId: string,
  defs: string[],
): string {
  const { fill, strokeWidth } = shapePaint(shape);
  const attributes = [
    `fill="${fill ? paint : "none"}"`,
    strokeWidth > 0 ? `stroke="${paint}" stroke-width="${n(strokeWidth)}"` : "",
    shape.lineCap ? `stroke-linecap="${shape.lineCap}"` : "",
    shape.lineJoin ? `stroke-linejoin="${shape.lineJoin}"` : "",
  ].filter(Boolean).join(" ");
  const outer = svgTransformAttribute(shape.transform ? { ...shape.transform, matrix: undefined } : undefined);
  const matrix = shape.transform?.matrix;
  const transformAttribute = matrix ? ` transform="matrix(${matrix.map(n).join(" ")})"` : "";
  let element: string;
  if (shape.kind === "text") {
    element =
      `<text x="${n(shape.x)}" y="${n(shape.y)}" font-family="${escapeText(shape.fontFamily)}" ` +
      `font-size="${n(shape.fontSize)}"${shape.fontStyle ? ` font-style="${shape.fontStyle}"` : ""}` +
      `${shape.fontWeight !== undefined ? ` font-weight="${shape.fontWeight}"` : ""} ${attributes}${transformAttribute}>` +
      `${escapeText(shape.text)}</text>`;
  } else {
    element = `<path d="${shapePathData(shape) ?? ""}" ${attributes}${transformAttribute}/>`;
  }
  if (shape.clipRect) {
    const { x, y, width, height } = shape.clipRect;
    defs.push(
      `<clipPath id="${clipId}" clipPathUnits="userSpaceOnUse"><rect x="${n(x)}" y="${n(y)}" width="${n(width)}" height="${n(height)}"/></clipPath>`,
    );
    element = `<g clip-path="url(#${clipId})">${element}</g>`;
  }
  return outer ? `<g transform="${outer}">${element}</g>` : element;
}
