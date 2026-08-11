import type {
  CustomSymbolClipRect,
  CustomSymbolLimitsBehavior,
  CustomSymbolMathRole,
  CustomSymbolMetrics,
  CustomSymbolVectorShape,
  CustomSymbolVectorTransform,
} from "./customSymbolTypes";

export const CUSTOM_SYMBOL_DESIGNER_DOCUMENT_VERSION = 1 as const;

export type CustomSymbolGeometryPreset =
  | "line"
  | "circle"
  | "ellipse"
  | "rect"
  | "triangle"
  | "arrow"
  | "arc"
  | "eraser";

export interface CustomSymbolGlyphAsset {
  sourceLatex: string;
  metrics: CustomSymbolMetrics;
  shapes: CustomSymbolVectorShape[];
}

export interface CustomSymbolDesignerLayerBase {
  id: string;
  name: string;
  visible: boolean;
  locked: boolean;
  transform: Omit<CustomSymbolVectorTransform, "matrix">;
  clipRect?: CustomSymbolClipRect | null;
}

export interface CustomSymbolDesignerGlyphLayer
  extends CustomSymbolDesignerLayerBase {
  kind: "glyph";
  asset: CustomSymbolGlyphAsset;
}

export interface CustomSymbolDesignerBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface CustomSymbolDesignerGeometryLayer
  extends CustomSymbolDesignerLayerBase {
  kind: "geometry";
  geometryPreset?: CustomSymbolGeometryPreset;
  shape: CustomSymbolVectorShape;
  bounds: CustomSymbolDesignerBounds;
}

export type CustomSymbolDesignerLayer =
  | CustomSymbolDesignerGlyphLayer
  | CustomSymbolDesignerGeometryLayer;

export interface CustomSymbolDesignerDocument {
  version: typeof CUSTOM_SYMBOL_DESIGNER_DOCUMENT_VERSION;
  symbolId: string | null;
  name: string;
  command: string;
  role: CustomSymbolMathRole;
  limitsBehavior: CustomSymbolLimitsBehavior;
  metrics: CustomSymbolMetrics;
  ommlFallback: string | null;
  layers: CustomSymbolDesignerLayer[];
}

export function createEmptyCustomSymbolDesignerDocument(): CustomSymbolDesignerDocument {
  return {
    version: CUSTOM_SYMBOL_DESIGNER_DOCUMENT_VERSION,
    symbolId: null,
    name: "",
    command: "",
    role: "ordinary",
    limitsBehavior: "auto",
    metrics: {
      widthEm: 3.2,
      ascentEm: 3,
      descentEm: 1.5,
    },
    ommlFallback: null,
    layers: [],
  };
}
