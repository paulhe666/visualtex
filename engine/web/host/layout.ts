// Internal geometry only. Office export units and DTOs remain unchanged.
import type { Position, Selection } from './editing';
export type MathStyle = 'display' | 'text' | 'script' | 'scriptScript';
export type SceneQuery =
  | { kind: 'hitTest'; xEm: number; yEm: number }
  | { kind: 'vertical'; from: Position; down: boolean; desiredX: number | null }
  | { kind: 'navigation'; selection: Selection; down: boolean; desiredX: number | null }
  | { kind: 'boundary'; selection: Selection; toEnd: boolean; document: boolean }
  | { kind: 'selection'; selection: Selection }
  | { kind: 'exportSvg' };
export interface PositionDescription {
  path: string[];
  offset: number;
  length: number;
  text: string | null;
}
export type SceneQueryResult =
  | { kind: 'hitTest'; position: Position }
  | { kind: 'vertical'; position: Position; origin: Position; desiredX: number }
  | { kind: 'boundary'; position: Position; origin: Position }
  | { kind: 'selection'; caret: Position | null; rectangles: Rect[]; description: PositionDescription }
  | { kind: 'svg'; svg: string; width: number; height: number; baseline: number };
export interface CustomSymbolTextFontBinding {
  fontFamily:string;
  fontStyle:'normal'|'italic';
  fontWeight:number;
  /** Registered core font identity: immutable bytes SHA-256 plus face index. */
  fontId:string;
}
export interface LayoutOptions {
  /** Structural mathematical face. It must expose an OpenType MATH table. */
  mathFont: string;
  textFont: string | null;
  /** Optional real product-facing alphabet faces. These may be ordinary text
   * fonts without a MATH table; the core uses them only for Latin/Greek/digit
   * glyphs while structural geometry continues to come from mathFont. */
  alphabetUprightFont?: string | null;
  alphabetItalicFont?: string | null;
  alphabetBoldFont?: string | null;
  alphabetBoldItalicFont?: string | null;
  /** Verified TeX face roles; these fonts use ordinary cmap positions. */
  texFontSet?: Partial<Record<
    'main' | 'mainItalic' | 'mainBold' | 'mainBoldItalic' | 'mathItalic' | 'mathBoldItalic' |
    'ams' | 'calligraphic' | 'calligraphicBold' | 'script' | 'fraktur' | 'frakturBold' |
    'sans' | 'sansBold' | 'sansItalic' | 'monospace' | 'size1' | 'size2' | 'size3' | 'size4', string>>;
  /** Explicit verified font IDs, not CSS family names or implicit OS fallback. */
  mathFallbacks?: string[];
  customSymbolTextFonts?:CustomSymbolTextFontBinding[];
  style: MathStyle;
  fontSizePx: number;
  paddingEm: number;
}
export interface FontIdentity {
  id: string;
  sha256: string;
  faceIndex: number;
  unitsPerEm: number;
  variationAxes: [string, number][];
}
export interface Rect { xMin: number; yMin: number; xMax: number; yMax: number }
export interface RgbColor { r: number; g: number; b: number }
export type PathOp =
  | { op: 'move' | 'line'; points: [number, number] }
  | { op: 'quad'; points: [number, number, number, number] }
  | { op: 'cubic'; points: [number, number, number, number, number, number] }
  | { op: 'close' };
export interface GlyphShape {
  glyphId: number;
  advance: number;
  ink: [number, number, number, number] | null;
  italicCorrection: number;
  path: PathOp[];
}
export interface CaretStop {
  position: Position;
  x: number;
  baseline: number;
  top: number;
  bottom: number;
  depth: number;
}
export interface LayoutSnapshot {
  artworks?:import('./custom-artwork').ArtworkInstance[];
  paintOrder?:import('./custom-artwork').PaintItem[];
  documentId: string;
  revision: number;
  sourceId: number;
  fontSizePx: number;
  width: number;
  height: number;
  baseline: number;
  widthEm: number;
  heightEm: number;
  baselineEm: number;
  advanceEm: number;
  ink: Rect | null;
  fonts: FontIdentity[];
  resources: { fontId: string; shape: GlyphShape }[];
  glyphs: { nodeId: number; resource: number; x: number; y: number; scale: number; color?: RgbColor }[];
  rules: { nodeId: number; rect: Rect; color?: RgbColor }[];
  carets: CaretStop[];
  regions: { nodeId: number; role: string; logical: Rect; ink: Rect | null; depth: number }[];
  gridColumns?: { gridId: number; columnId: number; index: number; x: number; width: number }[];
  exportReady: boolean;
  rebuiltFragments: number;
  reusedFragments: number;
}
