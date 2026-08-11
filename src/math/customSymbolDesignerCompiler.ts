import type {
  CustomSymbolDefinition,
  CustomSymbolMetrics,
  CustomSymbolVectorShape,
  CustomSymbolVectorTransform,
} from "./customSymbolTypes";
import { createCustomSymbolDesignerSourceArchive } from "./customSymbolDesignerArchive";
import type {
  CustomSymbolDesignerDocument,
  CustomSymbolDesignerGlyphLayer,
  CustomSymbolDesignerLayer,
} from "./customSymbolDesignerTypes";

function compactTransform(
  transform: CustomSymbolVectorTransform,
): CustomSymbolVectorTransform | undefined {
  const result: CustomSymbolVectorTransform = {};
  if (transform.translateX) result.translateX = transform.translateX;
  if (transform.translateY) result.translateY = transform.translateY;
  if (transform.scaleX !== undefined && transform.scaleX !== 1) {
    result.scaleX = transform.scaleX;
  }
  if (transform.scaleY !== undefined && transform.scaleY !== 1) {
    result.scaleY = transform.scaleY;
  }
  if (transform.rotateDeg) result.rotateDeg = transform.rotateDeg;
  if (transform.originX) result.originX = transform.originX;
  if (transform.originY) result.originY = transform.originY;
  if (transform.matrix) result.matrix = [...transform.matrix];
  return Object.keys(result).length ? result : undefined;
}

function mergeLayerTransform(
  base: CustomSymbolVectorTransform | undefined,
  layer: CustomSymbolDesignerLayer["transform"],
) {
  return compactTransform({
    ...(base?.matrix ? { matrix: [...base.matrix] as CustomSymbolVectorTransform["matrix"] } : {}),
    translateX: (base?.translateX ?? 0) + (layer.translateX ?? 0),
    translateY: (base?.translateY ?? 0) + (layer.translateY ?? 0),
    scaleX: (base?.scaleX ?? 1) * (layer.scaleX ?? 1),
    scaleY: (base?.scaleY ?? 1) * (layer.scaleY ?? 1),
    rotateDeg: (base?.rotateDeg ?? 0) + (layer.rotateDeg ?? 0),
    originX: layer.originX ?? base?.originX,
    originY: layer.originY ?? base?.originY,
  });
}

function compileLayerShape(
  shape: CustomSymbolVectorShape,
  layer: CustomSymbolDesignerLayer,
): CustomSymbolVectorShape {
  return {
    ...shape,
    ...(mergeLayerTransform(shape.transform, layer.transform)
      ? { transform: mergeLayerTransform(shape.transform, layer.transform) }
      : { transform: undefined }),
    ...(layer.clipRect ? { clipRect: { ...layer.clipRect } } : {}),
  } as CustomSymbolVectorShape;
}

export function compileCustomSymbolDesignerArtwork(
  document: CustomSymbolDesignerDocument,
) {
  return document.layers.flatMap((layer) => {
    if (!layer.visible) return [];
    if (layer.kind === "glyph") {
      return layer.asset.shapes.map((shape) => compileLayerShape(shape, layer));
    }
    return [compileLayerShape(layer.shape, layer)];
  });
}

const svgNamespace = "http://www.w3.org/2000/svg";
const registrationPaddingUnits = 50;

function svgTransformParts(
  transform: CustomSymbolVectorTransform | undefined,
  includeMatrix: boolean,
) {
  if (!transform) return "";
  const parts: string[] = [];
  const tx = transform.translateX ?? 0;
  const ty = transform.translateY ?? 0;
  const sx = transform.scaleX ?? 1;
  const sy = transform.scaleY ?? 1;
  const angle = transform.rotateDeg ?? 0;
  const ox = transform.originX ?? 0;
  const oy = transform.originY ?? 0;
  if (tx || ty) parts.push(`translate(${tx} ${ty})`);
  if (ox || oy) parts.push(`translate(${ox} ${oy})`);
  if (angle) parts.push(`rotate(${angle})`);
  if (sx !== 1 || sy !== 1) parts.push(`scale(${sx} ${sy})`);
  if (ox || oy) parts.push(`translate(${-ox} ${-oy})`);
  if (includeMatrix && transform.matrix) {
    parts.push(`matrix(${transform.matrix.join(" ")})`);
  }
  return parts.join(" ");
}

function geometryElement(shape: CustomSymbolVectorShape) {
  const element = document.createElementNS(svgNamespace, shape.kind) as SVGGraphicsElement;
  switch (shape.kind) {
    case "path":
      element.setAttribute("d", shape.d);
      break;
    case "circle":
      element.setAttribute("cx", String(shape.cx));
      element.setAttribute("cy", String(shape.cy));
      element.setAttribute("r", String(shape.r));
      break;
    case "ellipse":
      element.setAttribute("cx", String(shape.cx));
      element.setAttribute("cy", String(shape.cy));
      element.setAttribute("rx", String(shape.rx));
      element.setAttribute("ry", String(shape.ry));
      break;
    case "line":
      element.setAttribute("x1", String(shape.x1));
      element.setAttribute("y1", String(shape.y1));
      element.setAttribute("x2", String(shape.x2));
      element.setAttribute("y2", String(shape.y2));
      break;
    case "rect":
      element.setAttribute("x", String(shape.x));
      element.setAttribute("y", String(shape.y));
      element.setAttribute("width", String(shape.width));
      element.setAttribute("height", String(shape.height));
      if (shape.rx !== undefined) element.setAttribute("rx", String(shape.rx));
      if (shape.ry !== undefined) element.setAttribute("ry", String(shape.ry));
      break;
    case "polygon":
      element.setAttribute(
        "points",
        shape.points.map(([x, y]) => `${x},${y}`).join(" "),
      );
      break;
  }
  const defaultFill = shape.kind !== "line";
  const fill = shape.fill ?? defaultFill;
  const strokeWidth = Math.max(
    0,
    shape.strokeWidth ?? (shape.kind === "line" ? 50 : 0),
  );
  element.setAttribute("fill", fill ? "black" : "none");
  if (strokeWidth > 0) {
    element.setAttribute("stroke", "black");
    element.setAttribute("stroke-width", String(strokeWidth));
    if (shape.lineCap) element.setAttribute("stroke-linecap", shape.lineCap);
    if (shape.lineJoin) element.setAttribute("stroke-linejoin", shape.lineJoin);
  } else {
    element.setAttribute("stroke", "none");
  }
  return element;
}

function appendMeasuredShape(
  parent: SVGGElement,
  defs: SVGDefsElement,
  shape: CustomSymbolVectorShape,
  index: number,
) {
  const outer = document.createElementNS(svgNamespace, "g");
  const userTransform = svgTransformParts(shape.transform, false);
  if (userTransform) outer.setAttribute("transform", userTransform);

  let geometryParent: SVGElement = outer;
  if (shape.clipRect) {
    const clipId = `visualtex-register-clip-${index}`;
    const clipPath = document.createElementNS(svgNamespace, "clipPath");
    clipPath.id = clipId;
    clipPath.setAttribute("clipPathUnits", "userSpaceOnUse");
    const rect = document.createElementNS(svgNamespace, "rect");
    rect.setAttribute("x", String(shape.clipRect.x));
    rect.setAttribute("y", String(shape.clipRect.y));
    rect.setAttribute("width", String(shape.clipRect.width));
    rect.setAttribute("height", String(shape.clipRect.height));
    clipPath.append(rect);
    defs.append(clipPath);
    const clipped = document.createElementNS(svgNamespace, "g");
    clipped.setAttribute("clip-path", `url(#${clipId})`);
    outer.append(clipped);
    geometryParent = clipped;
  }

  const geometry = geometryElement(shape);
  if (shape.transform?.matrix) {
    geometry.setAttribute(
      "transform",
      `matrix(${shape.transform.matrix.join(" ")})`,
    );
  }
  geometryParent.append(geometry);
  parent.append(outer);
}

function measureArtworkBounds(
  artwork: CustomSymbolVectorShape[],
  metrics: CustomSymbolMetrics,
) {
  if (typeof document === "undefined" || !document.body) return null;
  const paintShapes = artwork.filter((shape) => shape.operation !== "erase");
  if (!paintShapes.length) return null;

  const svg = document.createElementNS(svgNamespace, "svg");
  svg.setAttribute(
    "viewBox",
    `0 0 ${Math.max(1, metrics.widthEm * 1000)} ${Math.max(
      1,
      (metrics.ascentEm + metrics.descentEm) * 1000,
    )}`,
  );
  svg.style.position = "fixed";
  svg.style.left = "-100000px";
  svg.style.top = "-100000px";
  svg.style.width = "1000px";
  svg.style.height = "1000px";
  svg.style.pointerEvents = "none";
  svg.style.visibility = "hidden";
  const defs = document.createElementNS(svgNamespace, "defs");
  const group = document.createElementNS(svgNamespace, "g");
  svg.append(defs, group);
  paintShapes.forEach((shape, index) => appendMeasuredShape(group, defs, shape, index));
  document.body.append(svg);
  try {
    const graphics = group as SVGGElement & {
      getBBox(options?: {
        fill?: boolean;
        stroke?: boolean;
        markers?: boolean;
        clipped?: boolean;
      }): DOMRect;
    };
    let bounds: DOMRect;
    try {
      bounds = graphics.getBBox({
        fill: true,
        stroke: true,
        markers: true,
        clipped: true,
      });
    } catch {
      bounds = group.getBBox();
    }
    if (
      !Number.isFinite(bounds.x) ||
      !Number.isFinite(bounds.y) ||
      !Number.isFinite(bounds.width) ||
      !Number.isFinite(bounds.height) ||
      bounds.width <= 0 ||
      bounds.height <= 0
    ) {
      return null;
    }
    return {
      left: bounds.x,
      top: bounds.y,
      right: bounds.x + bounds.width,
      bottom: bounds.y + bounds.height,
    };
  } finally {
    svg.remove();
  }
}

function shiftArtwork(
  artwork: CustomSymbolVectorShape[],
  shiftX: number,
  shiftY: number,
) {
  return artwork.map((shape) => ({
    ...shape,
    transform: compactTransform({
      ...shape.transform,
      translateX: (shape.transform?.translateX ?? 0) + shiftX,
      translateY: (shape.transform?.translateY ?? 0) + shiftY,
    }),
  })) as CustomSymbolVectorShape[];
}

function autoCropRegisteredArtwork(
  artwork: CustomSymbolVectorShape[],
  designerMetrics: CustomSymbolMetrics,
) {
  const measured = measureArtworkBounds(artwork, designerMetrics);
  if (!measured) {
    return { artwork, metrics: { ...designerMetrics } };
  }
  const baseline = designerMetrics.ascentEm * 1000;
  const left = measured.left - registrationPaddingUnits;
  const right = measured.right + registrationPaddingUnits;
  const top = Math.min(
    measured.top - registrationPaddingUnits,
    baseline - 20,
  );
  const bottom = Math.max(
    measured.bottom + registrationPaddingUnits,
    baseline,
  );
  const width = Math.max(20, right - left);
  const height = Math.max(20, bottom - top);
  const normalizedBaseline = baseline - top;
  return {
    artwork: shiftArtwork(artwork, -left, -top),
    metrics: {
      widthEm: Number((width / 1000).toFixed(6)),
      ascentEm: Number((normalizedBaseline / 1000).toFixed(6)),
      descentEm: Number(((height - normalizedBaseline) / 1000).toFixed(6)),
    },
  };
}

export function customSymbolDefinitionFromDesignerDocument(
  document: CustomSymbolDesignerDocument,
  options: {
    id: string;
    createdAt?: number;
    updatedAt?: number;
  },
): CustomSymbolDefinition {
  const now = Date.now();
  const compiledArtwork = compileCustomSymbolDesignerArtwork(document);
  const cropped = autoCropRegisteredArtwork(compiledArtwork, document.metrics);
  return {
    id: options.id,
    command: document.command.trim().replace(/^\\/, ""),
    name: document.name.trim(),
    role: document.role,
    limitsBehavior: document.limitsBehavior,
    metrics: cropped.metrics,
    artwork: {
      shapes: cropped.artwork,
    },
    ommlFallback: document.ommlFallback?.trim() || null,
    designerSource: createCustomSymbolDesignerSourceArchive(document),
    createdAt: options.createdAt ?? now,
    updatedAt: options.updatedAt ?? now,
  };
}

export function glyphLayerFromAsset(
  asset: CustomSymbolDesignerGlyphLayer["asset"],
  options: {
    id: string;
    name?: string;
  },
): CustomSymbolDesignerGlyphLayer {
  return {
    id: options.id,
    name: options.name?.trim() || asset.sourceLatex,
    kind: "glyph",
    visible: true,
    locked: false,
    transform: {
      translateX: 0,
      translateY: 0,
      scaleX: 1,
      scaleY: 1,
      rotateDeg: 0,
      originX: (asset.metrics.widthEm * 1000) / 2,
      originY:
        ((asset.metrics.ascentEm + asset.metrics.descentEm) * 1000) / 2,
    },
    clipRect: null,
    asset,
  };
}
