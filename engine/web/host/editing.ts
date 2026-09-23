import type { DocumentSnapshot } from './protocol';
export type Affinity = 'before' | 'after';
export type Position =
  | { kind: 'row'; nodeId: number; offset: number; affinity: Affinity }
  | { kind: 'text'; nodeId: number; grapheme: number; affinity: Affinity };
export type Selection =
  | { kind: 'range'; anchor: Position; focus: Position }
  | { kind: 'node'; nodeId: number }
  | { kind: 'gridCells'; gridId: number; anchorCell: number; focusCell: number };
export type NormalizedRange =
  | { kind: 'row'; nodeId: number; start: number; end: number; promoted: boolean }
  | { kind: 'text'; nodeId: number; start: number; end: number }
  | { kind: 'gridCells'; rectangle: { gridId: number; rowStart: number; rowEnd: number; columnStart: number; columnEnd: number } };
export type Horizontal = 'left' | 'right';
export type Vertical = 'up' | 'down';
export type InputMode = 'math' | 'text';
export interface ClipboardAssetRef { kind: string; assetId: string; sha256: string; definition?:Record<string,unknown> }
export interface StructuredClipboardFragment {
  schema: string;
  version: number;
  sourceSha256: string;
  shape: unknown;
  assets: ClipboardAssetRef[];
}
export type EditCommand =
  | { kind: 'insertSymbol'; command: string }
  | { kind: 'insertCustomSymbol'; command:string }
  | { kind: 'insertText' | 'replaceRange'; text: string; mode: InputMode }
  | { kind: 'replaceRecentMath'; expected:string; latex:string }
  | { kind: 'insertStructure'; structure: 'fraction' | 'binomial' | 'skewedFraction' | 'superscript' | 'subscript' | 'text' | 'squareRoot' | 'indexedRoot' }
  | { kind: 'insertFence'; left: string; right: string }
  | { kind: 'typeDelimiter'; text: string }
  | { kind: 'insertAccent' | 'insertAnnotation'; command: string }
  | { kind: 'insertOperator'; command: string; withLower: boolean; withUpper: boolean }
  | { kind: 'insertGeneralOperator'; withLower: boolean; withUpper: boolean }
  | { kind: 'setOperatorLimits'; mode: import('./syntax').LimitMode }
  | { kind: 'exitStructure' }
  | { kind: 'splitLine' }
  | { kind: 'insertMatrix'; rows: number; columns: number }
  | { kind: 'insertCases'; rows: number }
  | { kind: 'insertAligned'; rows: number; pairs: number }
  | { kind: 'insertGridRow'; after: boolean }
  | { kind: 'deleteGridRow' }
  | { kind: 'insertGridColumn'; after: boolean }
  | { kind: 'deleteGridColumn' }
  | { kind: 'moveGridCell'; direction: 'left' | 'right' }
  | { kind: 'mergeLine'; direction: Horizontal }
  | { kind: 'deleteBackward' | 'deleteForward' | 'duplicateSelection' | 'splitText' }
  | { kind: 'pasteLatex' | 'pasteGridCells'; latex: string }
  | { kind: 'pasteStructured'; latex: string; fragment: StructuredClipboardFragment }
  | { kind: 'clearGridCells' }
  | { kind: 'mergeText'; direction: Horizontal }
  | { kind: 'setStyle'; style: string }
  | { kind: 'insertSpace' | 'insertPhantom'; command: string }
  | { kind: 'insertBoxed' }
  | { kind: 'setColor'; color: import('./layout').RgbColor };
export interface OcrInsertionTarget {
  sessionHandle:number;requestId:number;documentId:string;revision:number;sourceId:number;
  selectionEpoch:number;selection:Selection;
}
export type EditorAction =
  | { kind: 'edit'; command: EditCommand; groupId?: string | null }
  | { kind: 'applyOcrLatex'; target:OcrInsertionTarget; latex:string }
  | { kind: 'setSelection'; selection: Selection }
  | { kind: 'pointerSelection'; anchor: Position; focus: Position }
  | { kind: 'extendGridSelection'; direction: Horizontal | Vertical }
  | { kind: 'moveHorizontal'; direction: Horizontal; extend: boolean }
  | { kind: 'moveVertical'; direction: Vertical; desiredX?: number | null }
  | { kind: 'undo' | 'redo' }
  | { kind: 'beginComposition'; mode: InputMode }
  | { kind: 'updateComposition'; id: number; preedit: string }
  | { kind: 'commitComposition'; id: number; text?: string | null }
  | { kind: 'cancelComposition'; id: number };
export interface Composition {
  id: number;
  revision: number;
  selectionEpoch: number;
  selection: Selection;
  mode: InputMode;
  preedit: string;
}
export type PositionStep =
  | { kind: 'rowSplice'; nodeId: number; start: number; removed: number; inserted: number }
  | { kind: 'rowSplit'; nodeId: number; leftId: number; rightId: number; at: number; rightOffset: number }
  | { kind: 'rowOffset'; nodeId: number; targetId: number; offset: number }
  | { kind: 'textRemap'; nodeId: number; targetId: number; before: number[]; after: number[] }
  | { kind: 'textSplit'; nodeId: number; rightId: number; at: number }
  | { kind: 'removed'; nodeId: number; fallback: Position };
export interface PositionMap {
  retainedIds: number[];
  removedIds: number[];
  insertedIds: number[];
  steps: PositionStep[];
}
export interface EditorSnapshot {
  document: DocumentSnapshot;
  selection: Selection;
  normalizedSelection: NormalizedRange;
  selectionEpoch: number;
  undoDepth: number;
  redoDepth: number;
  historyWeight: number;
  composition: Composition | null;
}
export interface EditorActionResult {
  editor: EditorSnapshot;
  contentChanged: boolean;
  replayed: boolean;
  positionMap: PositionMap | null;
  verticalMove: { selection: Selection; direction: Vertical; desiredX: number | null } | null;
}
