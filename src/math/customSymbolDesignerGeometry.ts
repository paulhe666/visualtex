import { createUuid } from "../runtime/browserCompatibility";
import type {
  CustomSymbolDesignerBounds,
  CustomSymbolDesignerGeometryLayer,
  CustomSymbolDesignerGlyphLayer,
  CustomSymbolGeometryPreset,
} from "./customSymbolDesignerTypes";
import type { CustomSymbolMetrics, CustomSymbolVectorShape } from "./customSymbolTypes";

export type { CustomSymbolGeometryPreset } from "./customSymbolDesignerTypes";

export const CUSTOM_SYMBOL_GEOMETRY_PRESETS: readonly Exclude<
  CustomSymbolGeometryPreset,
  "eraser"
>[] = [
  "line",
  "circle",
  "ellipse",
  "rect",
  "triangle",
  "arrow",
  "arc",
];

const DEFAULT_STROKE_WIDTH = 22;

export interface CustomSymbolEraserPoint {
  x: number;
  y: number;
}

function rounded(value: number) {
  return Number(value.toFixed(2));
}

/**
 * Build a smooth Catmull-Rom style cubic spline through pointer samples.
 * The same path is used for live preview and the persisted erase artwork so
 * what the user sees while dragging is exactly what is exported later.
 */
export function buildSmoothCustomSymbolEraserPath(
  points: readonly CustomSymbolEraserPoint[],
) {
  if (points.length === 0) return "";
  if (points.length === 1) {
    const point = points[0];
    return `M${rounded(point.x)} ${rounded(point.y)}L${rounded(point.x + 0.01)} ${rounded(point.y)}`;
  }
  if (points.length === 2) {
    return `M${rounded(points[0].x)} ${rounded(points[0].y)}L${rounded(points[1].x)} ${rounded(points[1].y)}`;
  }

  const parts = [`M${rounded(points[0].x)} ${rounded(points[0].y)}`];
  for (let index = 0; index < points.length - 1; index += 1) {
    const p0 = points[Math.max(0, index - 1)];
    const p1 = points[index];
    const p2 = points[index + 1];
    const p3 = points[Math.min(points.length - 1, index + 2)];
    const c1x = p1.x + (p2.x - p0.x) / 6;
    const c1y = p1.y + (p2.y - p0.y) / 6;
    const c2x = p2.x - (p3.x - p1.x) / 6;
    const c2y = p2.y - (p3.y - p1.y) / 6;
    parts.push(
      `C${rounded(c1x)} ${rounded(c1y)} ${rounded(c2x)} ${rounded(c2y)} ${rounded(p2.x)} ${rounded(p2.y)}`,
    );
  }
  return parts.join(" ");
}

function clampDimension(value: number) {
  return Math.max(10, Math.min(8_000, Number.isFinite(value) ? value : 10));
}

function clampStroke(value: number) {
  return Math.max(0, Math.min(600, Number.isFinite(value) ? value : DEFAULT_STROKE_WIDTH));
}

function presetShape(
  preset: Exclude<CustomSymbolGeometryPreset, "eraser">,
): {
  shape: CustomSymbolVectorShape;
  bounds: CustomSymbolDesignerBounds;
} {
  switch (preset) {
    case "line":
      return {
        shape: {
          kind: "line",
          x1: 0,
          y1: 0,
          x2: 320,
          y2: 0,
          fill: false,
          strokeWidth: DEFAULT_STROKE_WIDTH,
          lineCap: "round",
        },
        bounds: { x: 0, y: -40, width: 320, height: 80 },
      };
    case "circle":
      return {
        shape: {
          kind: "circle",
          cx: 120,
          cy: 120,
          r: 100,
          fill: false,
          strokeWidth: DEFAULT_STROKE_WIDTH,
        },
        bounds: { x: 0, y: 0, width: 240, height: 240 },
      };
    case "ellipse":
      return {
        shape: {
          kind: "ellipse",
          cx: 160,
          cy: 100,
          rx: 140,
          ry: 80,
          fill: false,
          strokeWidth: DEFAULT_STROKE_WIDTH,
        },
        bounds: { x: 0, y: 0, width: 320, height: 200 },
      };
    case "rect":
      return {
        shape: {
          kind: "rect",
          x: 0,
          y: 0,
          width: 280,
          height: 180,
          rx: 4,
          ry: 4,
          fill: false,
          strokeWidth: DEFAULT_STROKE_WIDTH,
        },
        bounds: { x: 0, y: 0, width: 280, height: 180 },
      };
    case "triangle":
      return {
        shape: {
          kind: "polygon",
          points: [
            [150, 0],
            [300, 220],
            [0, 220],
          ],
          fill: false,
          strokeWidth: DEFAULT_STROKE_WIDTH,
          lineJoin: "round",
        },
        bounds: { x: 0, y: 0, width: 300, height: 220 },
      };
    case "arrow":
      return {
        shape: {
          kind: "path",
          d: "M0 90H300M220 10L300 90L220 170",
          fill: false,
          strokeWidth: DEFAULT_STROKE_WIDTH,
          lineCap: "round",
          lineJoin: "round",
        },
        bounds: { x: 0, y: 0, width: 300, height: 180 },
      };
    case "arc":
      return {
        shape: {
          kind: "path",
          d: "M0 170Q160 0 320 170",
          fill: false,
          strokeWidth: DEFAULT_STROKE_WIDTH,
          lineCap: "round",
        },
        bounds: { x: 0, y: 0, width: 320, height: 170 },
      };
  }
}

function centeredTransform(
  bounds: CustomSymbolDesignerBounds,
  canvasMetrics: CustomSymbolMetrics,
) {
  const canvasWidth = canvasMetrics.widthEm * 1000;
  const canvasHeight = (canvasMetrics.ascentEm + canvasMetrics.descentEm) * 1000;
  return {
    translateX: (canvasWidth - bounds.width) / 2 - bounds.x,
    translateY: (canvasHeight - bounds.height) / 2 - bounds.y,
    scaleX: 1,
    scaleY: 1,
    rotateDeg: 0,
    originX: bounds.x + bounds.width / 2,
    originY: bounds.y + bounds.height / 2,
  };
}

export function createCustomSymbolGeometryLayer(
  preset: Exclude<CustomSymbolGeometryPreset, "eraser">,
  canvasMetrics: CustomSymbolMetrics,
  options: { id?: string; name?: string } = {},
): CustomSymbolDesignerGeometryLayer {
  const { shape, bounds } = presetShape(preset);
  return {
    id: options.id ?? createUuid(),
    name: options.name ?? preset,
    kind: "geometry",
    geometryPreset: preset,
    visible: true,
    locked: false,
    transform: centeredTransform(bounds, canvasMetrics),
    clipRect: null,
    shape,
    bounds,
  };
}

export interface CustomSymbolGeometryProperties {
  width: number;
  height: number;
  strokeWidth: number;
  cornerRadius: number;
  fill: boolean;
}

export function customSymbolGeometryProperties(
  layer: CustomSymbolDesignerGeometryLayer,
): CustomSymbolGeometryProperties {
  const shape = layer.shape;
  const strokeWidth = shape.strokeWidth ?? 0;
  switch (layer.geometryPreset) {
    case "line":
      if (shape.kind === "line") {
        return {
          width: Math.hypot(shape.x2 - shape.x1, shape.y2 - shape.y1),
          height: Math.max(10, strokeWidth),
          strokeWidth,
          cornerRadius: 0,
          fill: false,
        };
      }
      break;
    case "circle":
      if (shape.kind === "circle") {
        return {
          width: shape.r * 2,
          height: shape.r * 2,
          strokeWidth,
          cornerRadius: shape.r,
          fill: shape.fill === true,
        };
      }
      break;
    case "ellipse":
      if (shape.kind === "ellipse") {
        return {
          width: shape.rx * 2,
          height: shape.ry * 2,
          strokeWidth,
          cornerRadius: 0,
          fill: shape.fill === true,
        };
      }
      break;
    case "rect":
      if (shape.kind === "rect") {
        return {
          width: shape.width,
          height: shape.height,
          strokeWidth,
          cornerRadius: shape.rx ?? 0,
          fill: shape.fill === true,
        };
      }
      break;
    case "triangle":
    case "arrow":
    case "arc":
      return {
        width: layer.bounds.width,
        height: layer.bounds.height,
        strokeWidth,
        cornerRadius: 0,
        fill: shape.fill === true,
      };
    case "eraser":
      return {
        width: layer.bounds.width,
        height: layer.bounds.height,
        strokeWidth,
        cornerRadius: 0,
        fill: false,
      };
    default:
      return {
        width: layer.bounds.width,
        height: layer.bounds.height,
        strokeWidth,
        cornerRadius: shape.kind === "rect" ? shape.rx ?? 0 : 0,
        fill: shape.fill === true,
      };
  }
  return {
    width: layer.bounds.width,
    height: layer.bounds.height,
    strokeWidth,
    cornerRadius: 0,
    fill: shape.fill === true,
  };
}

export function updateCustomSymbolGeometryLayer(
  layer: CustomSymbolDesignerGeometryLayer,
  patch: Partial<CustomSymbolGeometryProperties>,
): CustomSymbolDesignerGeometryLayer {
  if (layer.geometryPreset === "eraser") {
    const strokeWidth = clampStroke(patch.strokeWidth ?? layer.shape.strokeWidth ?? 60);
    return {
      ...layer,
      shape: { ...layer.shape, strokeWidth } as CustomSymbolVectorShape,
    };
  }

  const current = customSymbolGeometryProperties(layer);
  const width = clampDimension(patch.width ?? current.width);
  const height = clampDimension(patch.height ?? current.height);
  const strokeWidth = clampStroke(patch.strokeWidth ?? current.strokeWidth);
  const cornerRadius = Math.max(0, Math.min(Math.min(width, height) / 2, patch.cornerRadius ?? current.cornerRadius));
  const fill = patch.fill ?? current.fill;
  const preset = layer.geometryPreset;
  let shape: CustomSymbolVectorShape = layer.shape;
  let bounds = { ...layer.bounds };

  switch (preset) {
    case "line":
      shape = {
        kind: "line",
        x1: 0,
        y1: 0,
        x2: width,
        y2: 0,
        fill: false,
        strokeWidth,
        lineCap: layer.shape.lineCap ?? "round",
      };
      bounds = { x: 0, y: -Math.max(20, strokeWidth), width, height: Math.max(40, strokeWidth * 2) };
      break;
    case "circle": {
      const diameter = width;
      const r = diameter / 2;
      shape = {
        kind: "circle",
        cx: r,
        cy: r,
        r,
        fill,
        strokeWidth,
      };
      bounds = { x: 0, y: 0, width: diameter, height: diameter };
      break;
    }
    case "ellipse":
      shape = {
        kind: "ellipse",
        cx: width / 2,
        cy: height / 2,
        rx: width / 2,
        ry: height / 2,
        fill,
        strokeWidth,
      };
      bounds = { x: 0, y: 0, width, height };
      break;
    case "rect":
      shape = {
        kind: "rect",
        x: 0,
        y: 0,
        width,
        height,
        rx: cornerRadius,
        ry: cornerRadius,
        fill,
        strokeWidth,
      };
      bounds = { x: 0, y: 0, width, height };
      break;
    case "triangle":
      shape = {
        kind: "polygon",
        points: [
          [width / 2, 0],
          [width, height],
          [0, height],
        ],
        fill,
        strokeWidth,
        lineJoin: layer.shape.lineJoin ?? "round",
      };
      bounds = { x: 0, y: 0, width, height };
      break;
    case "arrow": {
      const mid = height / 2;
      const head = Math.min(width * 0.34, height * 0.6);
      shape = {
        kind: "path",
        d: `M0 ${mid}H${width}M${width - head} 0L${width} ${mid}L${width - head} ${height}`,
        fill: false,
        strokeWidth,
        lineCap: layer.shape.lineCap ?? "round",
        lineJoin: layer.shape.lineJoin ?? "round",
      };
      bounds = { x: 0, y: 0, width, height };
      break;
    }
    case "arc":
      shape = {
        kind: "path",
        d: `M0 ${height}Q${width / 2} 0 ${width} ${height}`,
        fill: false,
        strokeWidth,
        lineCap: layer.shape.lineCap ?? "round",
      };
      bounds = { x: 0, y: 0, width, height };
      break;
    default: {
      const nextShape = { ...layer.shape, strokeWidth, fill } as CustomSymbolVectorShape;
      if (nextShape.kind === "rect") {
        nextShape.width = width;
        nextShape.height = height;
        nextShape.rx = cornerRadius;
        nextShape.ry = cornerRadius;
      } else if (nextShape.kind === "ellipse") {
        nextShape.rx = width / 2;
        nextShape.ry = height / 2;
        nextShape.cx = width / 2;
        nextShape.cy = height / 2;
      } else if (nextShape.kind === "circle") {
        nextShape.r = width / 2;
        nextShape.cx = width / 2;
        nextShape.cy = width / 2;
      }
      shape = nextShape;
      bounds = { ...bounds, width, height };
    }
  }

  return {
    ...layer,
    shape,
    bounds,
    transform: {
      ...layer.transform,
      originX: bounds.x + bounds.width / 2,
      originY: bounds.y + bounds.height / 2,
    },
  };
}

export function createCustomSymbolEraserLayer(
  points: Array<{ x: number; y: number }>,
  strokeWidth: number,
  options: { id?: string; name?: string } = {},
): CustomSymbolDesignerGeometryLayer | null {
  if (points.length < 1) return null;
  const width = Math.max(4, clampStroke(strokeWidth || 40));
  const d = buildSmoothCustomSymbolEraserPath(points);
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  const minX = Math.min(...xs) - width / 2;
  const maxX = Math.max(...xs) + width / 2;
  const minY = Math.min(...ys) - width / 2;
  const maxY = Math.max(...ys) + width / 2;
  return {
    id: options.id ?? createUuid(),
    name: options.name ?? "Eraser",
    kind: "geometry",
    geometryPreset: "eraser",
    visible: true,
    locked: false,
    transform: {
      translateX: 0,
      translateY: 0,
      scaleX: 1,
      scaleY: 1,
      rotateDeg: 0,
      originX: (minX + maxX) / 2,
      originY: (minY + maxY) / 2,
    },
    clipRect: null,
    shape: {
      kind: "path",
      operation: "erase",
      d,
      fill: false,
      strokeWidth: width,
      lineCap: "round",
      lineJoin: "round",
    },
    bounds: {
      x: minX,
      y: minY,
      width: Math.max(1, maxX - minX),
      height: Math.max(1, maxY - minY),
    },
  };
}

function cloneGlyphLayer(layer: CustomSymbolDesignerGlyphLayer) {
  if (typeof structuredClone === "function") {
    return structuredClone(layer) as CustomSymbolDesignerGlyphLayer;
  }
  return JSON.parse(JSON.stringify(layer)) as CustomSymbolDesignerGlyphLayer;
}

export function createGlyphLayerSlices(
  layer: CustomSymbolDesignerGlyphLayer,
  orientation: "horizontal" | "vertical",
  count = 3,
) {
  const parts = Math.max(2, Math.min(8, Math.round(count)));
  const width = layer.asset.metrics.widthEm * 1000;
  const height =
    (layer.asset.metrics.ascentEm + layer.asset.metrics.descentEm) * 1000;
  return Array.from({ length: parts }, (_, index) => {
    const slice = cloneGlyphLayer(layer);
    slice.id = createUuid();
    slice.name = `${layer.name} ${index + 1}/${parts}`;
    if (orientation === "horizontal") {
      const top = (height * index) / parts;
      const bottom = (height * (index + 1)) / parts;
      slice.clipRect = {
        x: 0,
        y: top,
        width,
        height: bottom - top,
      };
    } else {
      const left = (width * index) / parts;
      const right = (width * (index + 1)) / parts;
      slice.clipRect = {
        x: left,
        y: 0,
        width: right - left,
        height,
      };
    }
    return slice;
  });
}

export function glyphLayerFullClip(layer: CustomSymbolDesignerGlyphLayer) {
  return {
    x: 0,
    y: 0,
    width: layer.asset.metrics.widthEm * 1000,
    height:
      (layer.asset.metrics.ascentEm + layer.asset.metrics.descentEm) * 1000,
  };
}
