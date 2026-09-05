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
  if (transform.skewXDeg) result.skewXDeg = transform.skewXDeg;
  if (transform.skewYDeg) result.skewYDeg = transform.skewYDeg;
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
  // Glyph-compiler matrices remain the innermost exact transform. The designer
  // transform is deliberately kept as the outer translate/rotate/scale so the
  // source glyph can always be restored without decomposing its matrix.
  return compactTransform({
    ...(base?.matrix ? { matrix: [...base.matrix] as CustomSymbolVectorTransform["matrix"] } : {}),
    translateX: (base?.translateX ?? 0) + (layer.translateX ?? 0),
    translateY: (base?.translateY ?? 0) + (layer.translateY ?? 0),
    scaleX: (base?.scaleX ?? 1) * (layer.scaleX ?? 1),
    scaleY: (base?.scaleY ?? 1) * (layer.scaleY ?? 1),
    skewXDeg: (base?.skewXDeg ?? 0) + (layer.skewXDeg ?? 0),
    skewYDeg: (base?.skewYDeg ?? 0) + (layer.skewYDeg ?? 0),
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

function shapeWithLocalOffset(
  shape: CustomSymbolVectorShape,
  offsetX: number,
  offsetY: number,
) {
  return {
    ...shape,
    transform: compactTransform({
      ...shape.transform,
      translateX: (shape.transform?.translateX ?? 0) + offsetX,
      translateY: (shape.transform?.translateY ?? 0) + offsetY,
    }),
  } as CustomSymbolVectorShape;
}

function outlinedShape(
  shape: CustomSymbolVectorShape,
  layer: CustomSymbolDesignerLayer,
) {
  const outline = layer.effects?.outline;
  if (!outline?.enabled || shape.operation === "erase") return shape;
  const scaleX = Math.max(0.02, Math.abs(layer.transform.scaleX ?? 1));
  const scaleY = Math.max(0.02, Math.abs(layer.transform.scaleY ?? 1));
  const scaleCompensation = Math.sqrt(scaleX * scaleY);
  return {
    ...shape,
    fill: false,
    strokeWidth: Math.max(1, outline.width / scaleCompensation),
    lineCap: shape.lineCap ?? "round",
    lineJoin: shape.lineJoin ?? "round",
  } as CustomSymbolVectorShape;
}

/**
 * Return the editable layer artwork before the outer layer transform is
 * applied. The same function is used by the designer preview and registration
 * compiler so outline and perspective effects cannot diverge.
 */
export function customSymbolDesignerLayerLocalShapes(
  layer: CustomSymbolDesignerLayer,
): CustomSymbolVectorShape[] {
  const sourceShapes =
    layer.kind === "glyph" ? layer.asset.shapes : [layer.shape];
  if (sourceShapes.some((shape) => shape.operation === "erase")) {
    return sourceShapes;
  }

  const frontShapes = sourceShapes.map((shape) => outlinedShape(shape, layer));
  const perspective = layer.effects?.perspective;
  if (!perspective?.enabled || perspective.depth <= 0) return frontShapes;

  const steps = Math.max(1, Math.min(24, Math.round(perspective.steps)));
  const radians = (perspective.angleDeg * Math.PI) / 180;
  const depthX = Math.cos(radians) * perspective.depth;
  const depthY = Math.sin(radians) * perspective.depth;
  const extrusion: CustomSymbolVectorShape[] = [];
  for (let step = steps; step >= 1; step -= 1) {
    const ratio = step / steps;
    for (const shape of sourceShapes) {
      extrusion.push(shapeWithLocalOffset(shape, depthX * ratio, depthY * ratio));
    }
  }
  return [...extrusion, ...frontShapes];
}

export function compileCustomSymbolDesignerArtwork(
  document: CustomSymbolDesignerDocument,
) {
  return document.layers.flatMap((layer) => {
    if (!layer.visible) return [];
    return customSymbolDesignerLayerLocalShapes(layer).map((shape) =>
      compileLayerShape(shape, layer),
    );
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
  const skewX = transform.skewXDeg ?? 0;
  const skewY = transform.skewYDeg ?? 0;
  const angle = transform.rotateDeg ?? 0;
  const ox = transform.originX ?? 0;
  const oy = transform.originY ?? 0;
  if (tx || ty) parts.push(`translate(${tx} ${ty})`);
  if (ox || oy) parts.push(`translate(${ox} ${oy})`);
  if (angle) parts.push(`rotate(${angle})`);
  if (skewX) parts.push(`skewX(${skewX})`);
  if (skewY) parts.push(`skewY(${skewY})`);
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
    case "text":
      element.setAttribute("x", String(shape.x));
      element.setAttribute("y", String(shape.y));
      element.setAttribute("font-family", shape.fontFamily);
      element.setAttribute("font-size", String(shape.fontSize));
      if (shape.fontStyle) element.setAttribute("font-style", shape.fontStyle);
      if (shape.fontWeight !== undefined) {
        element.setAttribute("font-weight", String(shape.fontWeight));
      }
      element.textContent = shape.text;
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

function transformStrokeScale(
  transform: CustomSymbolVectorTransform | undefined,
) {
  if (!transform) return 1;
  const scaleX = Math.abs(transform.scaleX ?? 1);
  const scaleY = Math.abs(transform.scaleY ?? 1);
  const skewX = Math.abs(Math.tan(((transform.skewXDeg ?? 0) * Math.PI) / 180));
  const skewY = Math.abs(Math.tan(((transform.skewYDeg ?? 0) * Math.PI) / 180));
  const shearBound = 1 + skewX + skewY;
  const matrix = transform.matrix;
  const matrixBound = matrix
    ? Math.max(
        1,
        Math.hypot(matrix[0], matrix[1]) +
          Math.hypot(matrix[2], matrix[3]),
      )
    : 1;
  return Math.max(0.02, scaleX, scaleY) * shearBound * matrixBound;
}

function fallbackArtworkStrokePadding(artwork: CustomSymbolVectorShape[]) {
  return artwork.reduce((maximum, shape) => {
    if (shape.operation === "erase") return maximum;
    const defaultWidth = shape.kind === "line" ? 50 : 0;
    const width = Math.max(0, shape.strokeWidth ?? defaultWidth);
    if (!width) return maximum;
    const radius = (width * transformStrokeScale(shape.transform)) / 2;
    return Math.max(maximum, radius * 1.25);
  }, 0);
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
    let fallbackStrokePadding = 0;
    try {
      bounds = graphics.getBBox({
        fill: true,
        stroke: true,
        markers: true,
        clipped: true,
      });
    } catch {
      bounds = group.getBBox();
      fallbackStrokePadding = fallbackArtworkStrokePadding(paintShapes);
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
      left: bounds.x - fallbackStrokePadding,
      top: bounds.y - fallbackStrokePadding,
      right: bounds.x + bounds.width + fallbackStrokePadding,
      bottom: bounds.y + bounds.height + fallbackStrokePadding,
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

export function fitCustomSymbolDesignerDocumentToArtwork(
  source: CustomSymbolDesignerDocument,
  paddingUnits = registrationPaddingUnits,
): CustomSymbolDesignerDocument {
  const artwork = compileCustomSymbolDesignerArtwork(source);
  const measured = measureArtworkBounds(artwork, source.metrics);
  if (!measured) return source;
  const baseline = source.metrics.ascentEm * 1000;
  const left = measured.left - paddingUnits;
  const right = measured.right + paddingUnits;
  const top = Math.min(measured.top - paddingUnits, baseline - 20);
  const bottom = Math.max(measured.bottom + paddingUnits, baseline);
  const width = Math.max(20, right - left);
  const height = Math.max(20, bottom - top);
  const normalizedBaseline = baseline - top;
  return {
    ...source,
    metrics: {
      widthEm: Number((width / 1000).toFixed(6)),
      ascentEm: Number((normalizedBaseline / 1000).toFixed(6)),
      descentEm: Number(((height - normalizedBaseline) / 1000).toFixed(6)),
    },
    layers: source.layers.map((layer) => ({
      ...layer,
      transform: {
        ...layer.transform,
        translateX: (layer.transform.translateX ?? 0) - left,
        translateY: (layer.transform.translateY ?? 0) - top,
      },
    })),
  };
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
