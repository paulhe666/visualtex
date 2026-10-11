// Internal engine v1 only. These types are NOT Office DTOs or persisted Office metadata.
import type { MathNode, SourceAnalysis } from './syntax';
import type { SourceEdit, SourceChangeMap } from './source-edit';
import type { EditorAction, EditorSnapshot, EditorActionWireResult, StructuredClipboardFragment, OcrInsertionTarget } from './editing';
import type { LayoutOptions, LayoutSnapshot, FontIdentity, SceneQuery, SceneQueryResult } from './layout';
export type NodeId = number;
export type Revision = number;
export type SessionHandle = number;
export interface SessionSnapshot {
  documentId: string;
  revision: Revision;
  rootId: NodeId;
  source: string;
}
export interface DocumentSnapshot {
  documentId: string;
  revision: Revision;
  sourceId: number;
  source: string;
  analysis: SourceAnalysis;
  retainedTree: MathNode | null;
  canonicalLatex: string | null;
}
export interface ClipboardFragment {
  schemaVersion: 1;
  latex: string | null;
  structure?: StructuredClipboardFragment;
  revision: number;
  selectionEpoch: number;
}
export type EngineCommand =
  | { op: 'create'; documentId: string; source?: string | null }
  | { op: 'createWithSymbols'; documentId:string;source:string;symbols:unknown[] }
  | { op: 'analyzeCustomSymbol'; definition:unknown }
  | { op: 'analyzeSourceWithSymbols'; source:string;symbols:unknown[] }
  | { op: 'inspect'; handle: SessionHandle; expectedRevision?: Revision }
  | { op: 'destroy'; handle: SessionHandle; expectedRevision: Revision }
  | { op: 'parkSession'; handle: SessionHandle; expectedRevision: Revision }
  | { op: 'resumeSession'; handle: SessionHandle; expectedRevision: Revision }
  | { op: 'analyzeSource'; source: string }
  | { op: 'replaceSource'; handle: SessionHandle; expectedRevision: Revision; source: string }
  | { op: 'patchSource'; handle: SessionHandle; expectedRevision: Revision; edit: SourceEdit }
  | { op: 'inspectDocument'; handle: SessionHandle; expectedRevision?: Revision }
  | { op: 'inspectEditor'; handle: SessionHandle; expectedRevision?: Revision }
  | { op: 'copySelection'; handle: SessionHandle; expectedRevision: Revision }
  | { op: 'captureOcrTarget'; handle:SessionHandle; expectedRevision:Revision }
  | { op: 'cancelOcrTarget'; handle:SessionHandle; requestId:number }
  | { op: 'editorAction'; handle: SessionHandle; expectedRevision: Revision; eventSequence: number; action: EditorAction; omitUnchangedDocument?: boolean }
  | { op: 'layout'; handle: SessionHandle; expectedRevision: Revision; options: LayoutOptions; includeSvg: boolean }
  | { op: 'exportPresentationMathMl'; handle: SessionHandle; expectedRevision: Revision; displayMode: boolean }
  | { op: 'queryScene'; handle: SessionHandle; expectedRevision: Revision; sceneToken: number; query: SceneQuery }
  | { op: 'stats' };
export interface EngineRequest { protocol: 1; command: EngineCommand }
export type EngineReply =
  | { kind: 'session'; handle: SessionHandle; snapshot: SessionSnapshot }
  | { kind: 'customSymbolAnalysis'; asset:import('./custom-artwork').CustomSymbolAsset }
  | { kind: 'document'; handle: SessionHandle; snapshot: DocumentSnapshot }
  | { kind: 'editor'; handle: SessionHandle; snapshot: EditorSnapshot }
  | { kind: 'editorAction'; handle: SessionHandle; result: EditorActionWireResult }
  | { kind: 'sourcePatched'; handle: SessionHandle; snapshot: DocumentSnapshot; mapping: SourceChangeMap }
  | { kind: 'layout'; handle: SessionHandle; sceneToken: number; snapshot: LayoutSnapshot; svg: string | null }
  | { kind: 'presentationMathMl'; handle: SessionHandle; result: { mathMl: string; alignedGridIds: number[] } }
  | { kind: 'sceneQuery'; handle: SessionHandle; sceneToken: number; result: SceneQueryResult }
  | { kind: 'fontRegistered'; font: FontIdentity }
  | { kind: 'clipboard'; handle: SessionHandle; fragment: ClipboardFragment }
  | { kind: 'ocrTarget'; handle:SessionHandle; target:OcrInsertionTarget }
  | { kind: 'ocrTargetCancelled'; handle:SessionHandle; cancelled:boolean }
  | { kind: 'destroyed'; handle: SessionHandle }
  | { kind: 'sessionParked'; handle: SessionHandle; snapshot: SessionSnapshot }
  | { kind: 'sessionResumed'; handle: SessionHandle; snapshot: SessionSnapshot }
  | { kind: 'sourceAnalysis'; analysis: SourceAnalysis; lossless: string; canonical: string | null }
  | { kind: 'stats'; protocol: 1; activeSessions: number };
export type EngineErrorCode =
  | 'invalidInput' | 'unsupportedProtocol' | 'invalidHandle' | 'conflict'
  | 'staleRevision' | 'resourceLimit' | 'invalidPosition' | 'unsupportedSyntax'
  | 'incompleteSource' | 'missingFont';
export type EngineResponse =
  | { status: 'ok'; result: EngineReply }
  | { status: 'error'; error: { code: EngineErrorCode; message: string } };
