import type { CustomSymbolDesignerBounds, CustomSymbolDesignerDocument, CustomSymbolDesignerLayer } from "../designerTypes.ts";
import { apply, transformMatrix, type Affine } from "../matrix.ts";

/** Where TeX's math axis is above the baseline (cmsy10's axis height), in designer units. */
export const MATH_AXIS = 250;
/** The x-height and cap height of Computer Modern, for the guides. */
export const X_HEIGHT = 431;
export const CAP_HEIGHT = 683;

/** The designer's baseline (y grows downward). */
export function baselineOf(document: CustomSymbolDesignerDocument) {
  return document.metrics.ascentEm * 1000;
}

/** The layer's box in its own coordinates (a glyph's TeX box or crop; a shape's bounds). */
export function layerBox(layer: CustomSymbolDesignerLayer): CustomSymbolDesignerBounds {
  if (layer.kind === "glyph") {
    return (
      layer.clipRect ?? {
        x: 0,
        y: 0,
        width: layer.asset.metrics.widthEm * 1000,
        height: (layer.asset.metrics.ascentEm + layer.asset.metrics.descentEm) * 1000,
      }
    );
  }
  return layer.bounds;
}

/**
 * The layer's reference point in its own coordinates: a glyph's TeX reference
 * point (left end of its baseline); a shape's bottom left corner.
 */
export function layerAnchor(layer: CustomSymbolDesignerLayer) {
  if (layer.kind === "glyph") return { x: 0, y: layer.asset.metrics.ascentEm * 1000 };
  return { x: layer.bounds.x, y: layer.bounds.y + layer.bounds.height };
}

export function layerMatrix(layer: CustomSymbolDesignerLayer): Affine {
  return transformMatrix(layer.transform, false);
}

/** The layer's box corners in designer coordinates, and their bounds. */
export function layerWorldBounds(layer: CustomSymbolDesignerLayer) {
  const box = layerBox(layer);
  const matrix = layerMatrix(layer);
  const corners = [
    apply(matrix, box.x, box.y),
    apply(matrix, box.x + box.width, box.y),
    apply(matrix, box.x + box.width, box.y + box.height),
    apply(matrix, box.x, box.y + box.height),
  ];
  const xs = corners.map((point) => point.x);
  const ys = corners.map((point) => point.y);
  return {
    x: Math.min(...xs),
    y: Math.min(...ys),
    width: Math.max(...xs) - Math.min(...xs),
    height: Math.max(...ys) - Math.min(...ys),
  };
}

/** The reference point's position: from the left edge (x = 0) and above the baseline, in em. */
export function anchorPosition(layer: CustomSymbolDesignerLayer, baseline: number) {
  const anchor = layerAnchor(layer);
  const point = apply(layerMatrix(layer), anchor.x, anchor.y);
  return { xEm: point.x / 1000, raiseEm: (baseline - point.y) / 1000 };
}

/** The layer moved so its reference point is at (xEm, raiseEm). */
export function withAnchorPosition(
  layer: CustomSymbolDesignerLayer,
  baseline: number,
  position: { xEm?: number; raiseEm?: number },
): CustomSymbolDesignerLayer {
  const current = anchorPosition(layer, baseline);
  const dx = ((position.xEm ?? current.xEm) - current.xEm) * 1000;
  const dy = -((position.raiseEm ?? current.raiseEm) - current.raiseEm) * 1000;
  return translated(layer, dx, dy);
}

export function translated(layer: CustomSymbolDesignerLayer, dx: number, dy: number): CustomSymbolDesignerLayer {
  if (!dx && !dy) return layer;
  return {
    ...layer,
    transform: {
      ...layer.transform,
      translateX: (layer.transform.translateX ?? 0) + dx,
      translateY: (layer.transform.translateY ?? 0) + dy,
    },
  };
}

/**
 * The same drawing with its rotation and scaling about the box centre, so
 * the next rotation or flip turns the layer in place.
 */
export function withCenteredOrigin(layer: CustomSymbolDesignerLayer): CustomSymbolDesignerLayer {
  const box = layerBox(layer);
  const ox = box.x + box.width / 2;
  const oy = box.y + box.height / 2;
  const transform = layer.transform;
  if ((transform.originX ?? 0) === ox && (transform.originY ?? 0) === oy) return layer;
  // M(p) = t + o + L(p - o): keep c = t + o - L o while o changes.
  const linear = transformMatrix({ ...transform, translateX: 0, translateY: 0, originX: 0, originY: 0 }, false);
  const oldOrigin = { x: transform.originX ?? 0, y: transform.originY ?? 0 };
  const oldLinear = apply(linear, oldOrigin.x, oldOrigin.y);
  const newLinear = apply(linear, ox, oy);
  return {
    ...layer,
    transform: {
      ...transform,
      originX: ox,
      originY: oy,
      translateX: (transform.translateX ?? 0) + oldOrigin.x - oldLinear.x - (ox - newLinear.x),
      translateY: (transform.translateY ?? 0) + oldOrigin.y - oldLinear.y - (oy - newLinear.y),
    },
  };
}

/** The union of the visible layers' boxes, without `exceptId`. */
export function contentBounds(document: CustomSymbolDesignerDocument, exceptId?: string) {
  let bounds: CustomSymbolDesignerBounds | null = null;
  for (const layer of document.layers) {
    if (!layer.visible || layer.id === exceptId) continue;
    if (layer.kind === "geometry" && layer.shape.operation === "erase") continue;
    const box = layerWorldBounds(layer);
    if (!bounds) {
      bounds = box;
      continue;
    }
    const x = Math.min(bounds.x, box.x);
    const y = Math.min(bounds.y, box.y);
    bounds = {
      x,
      y,
      width: Math.max(bounds.x + bounds.width, box.x + box.width) - x,
      height: Math.max(bounds.y + bounds.height, box.y + box.height) - y,
    };
  }
  return bounds;
}
