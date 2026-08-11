export type CustomSymbolMathRole =
  | "ordinary"
  | "binary"
  | "relation"
  | "operator"
  | "open"
  | "close"
  | "punctuation";

export type CustomSymbolLimitsBehavior = "auto" | "limits" | "nolimits";

export interface CustomSymbolMetrics {
  widthEm: number;
  ascentEm: number;
  descentEm: number;
}

export type CustomSymbolVectorMatrix = [
  number,
  number,
  number,
  number,
  number,
  number,
];

export interface CustomSymbolVectorTransform {
  translateX?: number;
  translateY?: number;
  scaleX?: number;
  scaleY?: number;
  rotateDeg?: number;
  originX?: number;
  originY?: number;
  matrix?: CustomSymbolVectorMatrix;
}

export interface CustomSymbolClipRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface CustomSymbolShapeBase {
  operation?: "erase";
  fill?: boolean;
  strokeWidth?: number;
  lineCap?: "butt" | "round" | "square";
  lineJoin?: "miter" | "round" | "bevel";
  transform?: CustomSymbolVectorTransform;
  clipRect?: CustomSymbolClipRect;
}

export interface CustomSymbolPathShape extends CustomSymbolShapeBase {
  kind: "path";
  d: string;
}

export interface CustomSymbolCircleShape extends CustomSymbolShapeBase {
  kind: "circle";
  cx: number;
  cy: number;
  r: number;
}

export interface CustomSymbolLineShape extends CustomSymbolShapeBase {
  kind: "line";
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

export interface CustomSymbolRectShape extends CustomSymbolShapeBase {
  kind: "rect";
  x: number;
  y: number;
  width: number;
  height: number;
  rx?: number;
  ry?: number;
}

export interface CustomSymbolEllipseShape extends CustomSymbolShapeBase {
  kind: "ellipse";
  cx: number;
  cy: number;
  rx: number;
  ry: number;
}

export interface CustomSymbolPolygonShape extends CustomSymbolShapeBase {
  kind: "polygon";
  points: Array<[number, number]>;
}

export type CustomSymbolVectorShape =
  | CustomSymbolPathShape
  | CustomSymbolCircleShape
  | CustomSymbolLineShape
  | CustomSymbolRectShape
  | CustomSymbolEllipseShape
  | CustomSymbolPolygonShape;

export interface CustomSymbolArtwork {
  shapes: CustomSymbolVectorShape[];
}

export interface CustomSymbolDesignerSourceAsset {
  id: string;
  sourceLatex: string;
  metrics: CustomSymbolMetrics;
  shapes: CustomSymbolVectorShape[];
}

export interface CustomSymbolDesignerSourceLayerBase {
  id: string;
  name: string;
  visible: boolean;
  locked: boolean;
  transform: CustomSymbolVectorTransform;
  clipRect?: CustomSymbolClipRect | null;
}

export interface CustomSymbolDesignerSourceGlyphLayer
  extends CustomSymbolDesignerSourceLayerBase {
  kind: "glyph";
  assetId: string;
}

export interface CustomSymbolDesignerSourceGeometryLayer
  extends CustomSymbolDesignerSourceLayerBase {
  kind: "geometry";
  geometryPreset?: "line" | "circle" | "ellipse" | "rect" | "triangle" | "arrow" | "arc" | "eraser";
  shape: CustomSymbolVectorShape;
  bounds: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
}

export type CustomSymbolDesignerSourceLayer =
  | CustomSymbolDesignerSourceGlyphLayer
  | CustomSymbolDesignerSourceGeometryLayer;

export interface CustomSymbolDesignerSourceArchive {
  version: 1;
  metrics?: CustomSymbolMetrics;
  assets: CustomSymbolDesignerSourceAsset[];
  layers: CustomSymbolDesignerSourceLayer[];
}

export interface CustomSymbolDefinition {
  id: string;
  command: string;
  name: string;
  role: CustomSymbolMathRole;
  limitsBehavior: CustomSymbolLimitsBehavior;
  metrics: CustomSymbolMetrics;
  artwork: CustomSymbolArtwork;
  ommlFallback?: string | null;
  designerSource?: CustomSymbolDesignerSourceArchive | null;
  createdAt: number;
  updatedAt: number;
}

export interface CustomSymbolLibrary {
  version: 1;
  symbols: CustomSymbolDefinition[];
}
