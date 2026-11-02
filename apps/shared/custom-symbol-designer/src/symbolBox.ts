import type { CustomSymbolDesignerDocument, CustomSymbolDesignerGlyphLayer, CustomSymbolDesignerLayer } from "./designerTypes.ts";
import { apply, transformMatrix } from "./matrix.ts";
import type { CustomSymbolMetrics } from "./symbolTypes.ts";

// The registered symbol's box. As TeX sets a character in the box of its
// font metrics, a designed symbol takes the box of its base character (the
// largest upright LaTeX material, or the one the user chose): the symbol then
// sits, spaces and takes scripts and limits as that character does, and
// overlays may reach outside the box as accents and slashes do. Without a
// base the box is the ink's, with a margin.

/** The margin around the ink when the box is the ink's (designer units). */
export const INK_BOX_PADDING = 50;

/** The italic correction of a material (em), as the engine measured it; undefined: not known yet. */
export type ItalicLookup = (sourceLatex: string) => number | undefined;

export interface InkBounds {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface CustomSymbolBox {
  /** The box's left edge and top in designer coordinates (y grows downward). */
  left: number;
  top: number;
  /** The baseline's y in designer coordinates. */
  baseline: number;
  metrics: CustomSymbolMetrics;
  /** The operator's italic correction (em); 0 unless the symbol is an operator on a base character. */
  italicEm: number;
  /** The layer the box comes from, or null (the ink's box). */
  base: CustomSymbolDesignerGlyphLayer | null;
  /** Where the ink reaches, from the box's top left, when outside the box. */
  artworkBounds?: { x: number; y: number; width: number; height: number };
}

/** Whether a glyph layer's material is LaTeX the engine typeset (else a system-font outline). */
export function isLatexMaterial(layer: CustomSymbolDesignerLayer): boolean {
  if (layer.kind !== "glyph") return false;
  const { sourceLatex, shapes } = layer.asset;
  if (!/^[\x20-\x7e]+$/.test(sourceLatex)) return false;
  // The engine draws glyphs as <use> of font outlines, each with its font matrix;
  // a system-font character is one outline (or text) without one.
  return shapes.every((shape) => shape.kind !== "text") && shapes.some((shape) => Boolean(shape.transform?.matrix));
}

/** Not rotated or slanted (flips and scaling keep a box a box). */
export function isUprightLayer(layer: CustomSymbolDesignerLayer) {
  const { rotateDeg = 0, skewXDeg = 0, skewYDeg = 0 } = layer.transform;
  const turns = (((rotateDeg % 360) + 360) % 360);
  return (turns < 1e-6 || turns > 360 - 1e-6) && Math.abs(skewXDeg) < 1e-6 && Math.abs(skewYDeg) < 1e-6;
}

/** The layers that can be the base: visible, upright LaTeX materials. */
export function baseCandidates(document: CustomSymbolDesignerDocument) {
  return document.layers.filter(
    (layer): layer is CustomSymbolDesignerGlyphLayer =>
      layer.kind === "glyph" && layer.visible && isLatexMaterial(layer) && isUprightLayer(layer),
  );
}

/** The layer whose TeX box the symbol takes (see CustomSymbolDesignerDocument.baseLayerId). */
export function baseLayerOf(document: CustomSymbolDesignerDocument): CustomSymbolDesignerGlyphLayer | null {
  if (document.baseLayerId === null) return null;
  const candidates = baseCandidates(document);
  const chosen = document.baseLayerId ? candidates.find((layer) => layer.id === document.baseLayerId) : undefined;
  if (chosen) return chosen;
  let best: CustomSymbolDesignerGlyphLayer | null = null;
  let bestArea = 0;
  for (const layer of candidates) {
    const { widthEm, ascentEm, descentEm } = layer.asset.metrics;
    const area = widthEm * (ascentEm + descentEm) * Math.abs((layer.transform.scaleX ?? 1) * (layer.transform.scaleY ?? 1));
    if (area > bestArea) {
      best = layer;
      bestArea = area;
    }
  }
  return best;
}

/** A layer's TeX box in designer coordinates (an upright layer: still a box). */
export function layerTexBox(layer: CustomSymbolDesignerGlyphLayer) {
  const matrix = transformMatrix(layer.transform, false);
  const width = layer.asset.metrics.widthEm * 1000;
  const height = (layer.asset.metrics.ascentEm + layer.asset.metrics.descentEm) * 1000;
  const corners = [apply(matrix, 0, 0), apply(matrix, width, 0), apply(matrix, width, height), apply(matrix, 0, height)];
  const xs = corners.map((point) => point.x);
  const ys = corners.map((point) => point.y);
  return { left: Math.min(...xs), top: Math.min(...ys), right: Math.max(...xs), bottom: Math.max(...ys) };
}

const round6 = (value: number) => Number(value.toFixed(6));

/** The symbol's box for `document` whose painted ink reaches `ink` (designer coordinates). */
export function computeSymbolBox(
  document: CustomSymbolDesignerDocument,
  ink: InkBounds,
  italicOf: ItalicLookup = () => undefined,
  /** the ink of everything but the base (null: nothing else is painted) */
  overlayInk: InkBounds | null = null,
): CustomSymbolBox {
  const baseline = document.metrics.ascentEm * 1000;
  const base = baseLayerOf(document);
  let left: number, top: number, right: number, bottom: number;
  let italicEm = 0;
  if (base) {
    ({ left, top, right, bottom } = layerTexBox(base));
    // overlays inside the base's box leave it as it is (the symbol sets as the
    // base does); overlays reaching out widen it, so what follows and the
    // limits make room for them
    // (the base's own outline may poke out of its box a little, as fonts' do)
    const baseRight = right;
    if (overlayInk) {
      left = Math.min(left, overlayInk.left);
      right = Math.max(right, overlayInk.right);
      top = Math.min(top, overlayInk.top);
      bottom = Math.max(bottom, overlayInk.bottom);
    }
    // a raised or lowered base keeps a box TeX accepts: some height, no negative depth
    top = Math.min(top, baseline - 20);
    bottom = Math.max(bottom, baseline);
    if (document.role === "operator") {
      // the italic correction tucks the subscript under the base's slant: an
      // overlay reaching out on the right takes that room, as much as it reaches
      const italic = (italicOf(base.asset.sourceLatex) ?? 0) * Math.abs(base.transform.scaleX ?? 1);
      italicEm = Math.max(0, italic - (right - baseRight) / 1000);
    }
  } else {
    left = ink.left - INK_BOX_PADDING;
    right = ink.right + INK_BOX_PADDING;
    top = Math.min(ink.top - INK_BOX_PADDING, baseline - 20);
    bottom = Math.max(ink.bottom + INK_BOX_PADDING, baseline);
  }
  const width = Math.max(20, right - left);
  const box: CustomSymbolBox = {
    left,
    top,
    baseline,
    metrics: {
      widthEm: round6(width / 1000),
      ascentEm: round6((baseline - top) / 1000),
      descentEm: round6((bottom - baseline) / 1000),
    },
    italicEm: round6(italicEm),
    base,
  };
  const margin = 4;
  if (ink.left < left - margin || ink.top < top - margin || ink.right > left + width + margin || ink.bottom > bottom + margin) {
    const x0 = Math.min(ink.left, left) - margin;
    const y0 = Math.min(ink.top, top) - margin;
    const x1 = Math.max(ink.right, left + width) + margin;
    const y1 = Math.max(ink.bottom, bottom) + margin;
    box.artworkBounds = { x: round6(x0 - left), y: round6(y0 - top), width: round6(x1 - x0), height: round6(y1 - y0) };
  }
  return box;
}
