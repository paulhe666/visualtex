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
  /** TeX advance width, in em. */
  widthEm: number;
  /** Height above the mathematical baseline, in em. */
  ascentEm: number;
  /** Depth below the mathematical baseline, in em. */
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
  skewXDeg?: number;
  skewYDeg?: number;
  rotateDeg?: number;
  originX?: number;
  originY?: number;
  /**
   * Optional immutable base transform produced by the glyph compiler. User
   * translate/scale/skew/rotate values are applied outside this matrix so imported
   * MathJax glyph geometry stays exact while remaining editable as one layer.
   */
  matrix?: CustomSymbolVectorMatrix;
}

export interface CustomSymbolClipRect {
  /** Layer-local coordinates, 1000 units/em; the crop follows layer transforms. */
  x: number;
  y: number;
  width: number;
  height: number;
}

interface CustomSymbolShapeBase {
  /** Paint is the default. Erase shapes subtract from the composed monochrome artwork. */
  operation?: "erase";
  fill?: boolean;
  strokeWidth?: number;
  lineCap?: "butt" | "round" | "square";
  lineJoin?: "miter" | "round" | "bevel";
  transform?: CustomSymbolVectorTransform;
  /** Non-destructive crop in final symbol-canvas coordinates. */
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

export interface CustomSymbolTextShape extends CustomSymbolShapeBase {
  kind: "text";
  text: string;
  x: number;
  /** Baseline position in designer coordinates. */
  y: number;
  fontFamily: string;
  fontSize: number;
  fontStyle?: "normal" | "italic";
  fontWeight?: number;
}

export type CustomSymbolVectorShape =
  | CustomSymbolPathShape
  | CustomSymbolCircleShape
  | CustomSymbolLineShape
  | CustomSymbolRectShape
  | CustomSymbolEllipseShape
  | CustomSymbolPolygonShape
  | CustomSymbolTextShape;

/**
 * Compiled monochrome vector artwork. Coordinates use a fixed 1000 units/em,
 * x grows rightward and y grows downward from the top of the symbol box.
 * The baseline is therefore `metrics.ascentEm * 1000`.
 */
export interface CustomSymbolArtwork {
  shapes: CustomSymbolVectorShape[];
}

export interface CustomSymbolOutlineEffect {
  enabled: boolean;
  /** Target outline thickness in designer units (1000 units/em). */
  width: number;
}

export interface CustomSymbolPerspectiveEffect {
  enabled: boolean;
  /** Total extrusion depth in designer units. */
  depth: number;
  /** Extrusion direction in SVG canvas degrees. */
  angleDeg: number;
  /** Number of vector copies used to form the extrusion. */
  steps: number;
}

export interface CustomSymbolLayerEffects {
  outline?: CustomSymbolOutlineEffect;
  perspective?: CustomSymbolPerspectiveEffect;
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
  effects?: CustomSymbolLayerEffects;
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
  /** Original designer-canvas metrics. Runtime symbol metrics may be auto-cropped. */
  metrics?: CustomSymbolMetrics;
  assets: CustomSymbolDesignerSourceAsset[];
  layers: CustomSymbolDesignerSourceLayer[];
}

export interface CustomSymbolDefinition {
  id: string;
  /** TeX control word without the leading backslash. Letters only. */
  command: string;
  name: string;
  role: CustomSymbolMathRole;
  limitsBehavior: CustomSymbolLimitsBehavior;
  metrics: CustomSymbolMetrics;
  artwork: CustomSymbolArtwork;
  /** Optional semantic fallback used only by MathML/Word OMML export. */
  ommlFallback?: string | null;
  /** Optional editable designer source; ignored by runtime rendering. */
  designerSource?: CustomSymbolDesignerSourceArchive | null;
  createdAt: number;
  updatedAt: number;
}

export interface CustomSymbolLibrary {
  version: 1;
  symbols: CustomSymbolDefinition[];
}
