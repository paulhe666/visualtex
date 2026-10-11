// Internal syntax-analysis schema, independent of Office metadata and Session DTOs.
// IDs returned by analyzeSource are local to that analysis, never edit handles.
import type { MathStyle } from './layout';
export interface SourceSpan {
  sourceId: number;
  byteStart: number;
  byteEnd: number;
  utf16Start: number;
  utf16End: number;
}
export type AtomClass = 'ordinary' | 'number' | 'binary' | 'relation' | 'open' | 'close' | 'punctuation' | 'operator' | 'inner';
export type ScriptKind = 'superscript' | 'subscript';
export interface ColorSpec {
  beforeModel: string; model: string | null; beforeValue: string; value: string;
  color: import('./layout').RgbColor;
}
export interface RowEntry { leading: string; node: MathNode }
export interface ScriptSlot { script: ScriptKind; beforeMarker: string; beforeArgument: string; argument: MathNode }
export type LimitMode = 'auto' | 'limits' | 'noLimits' | 'displayLimits';
export type MathVariant = 'default' | 'upright' | 'bold' | 'italic' | 'boldItalic' | 'sans' | 'sansBold'
  | 'sansItalic' | 'sansBoldItalic' | 'monospace' | 'doubleStruck' | 'script' | 'boldScript' | 'fraktur' | 'boldFraktur';
export type OperatorPart = { kind: 'script'; slot: ScriptSlot } | { kind: 'limits'; before: string; command: string };
export type GridColumnAlignment = 'left' | 'center' | 'right';
export interface ArrayColumnSpec { beforeSpec: string; source: string; alignments: GridColumnAlignment[]; verticalRules: number[] }
export type BuiltinMacroKind = 'absolute' | 'norm' | 'differential' | 'derivative' | 'partialDerivative'
  | 'commutator' | 'anticommutator' | 'poissonBracket' | 'bra' | 'ket' | 'braKet' | 'expectation'
  | 'matrixElement' | 'ketBra' | 'boldVector' | 'arrowVector' | 'unitVector'
  | 'extensibleRightArrow' | 'extensibleLeftArrow' | 'parenthesizedMod';
export type MathNode = { id: number; origin: SourceSpan | null } & (
  | { kind: 'row'; grouping: 'bare' | 'braced'; entries: RowEntry[]; trailing: string }
  | { kind: 'grid'; environment: 'gathered' | 'matrix' | 'pmatrix' | 'bmatrix' | 'Bmatrix' | 'vmatrix' | 'Vmatrix' | 'array' | 'cases' | 'aligned' | 'align' | 'alignStar' | 'substack'; columns: number[]; arraySpec: ArrayColumnSpec | null; beforeBeginName: string; beforeEndName: string; rows: MathNode[] }
  | { kind: 'gridRow'; cells: MathNode[]; separators: string[]; rowBreak: string }
  | { kind: 'symbol'; value: string; class: AtomClass; command: string | null }
  | { kind: 'text'; command: string; beforeBody: string; text: string }
  | { kind: 'fraction'; command: string; beforeNumerator: string; numerator: MathNode; beforeDenominator: string; denominator: MathNode }
  | { kind: 'radical'; command: string; beforeDegree: string; degree: MathNode | null; beforeRadicand: string; radicand: MathNode }
  | { kind: 'fence'; beforeLeft: string; left: string; body: MathNode; beforeRight: string; right: string }
  | { kind: 'middleDelimiter'; beforeDelimiter: string; delimiter: string }
  | { kind: 'accent'; command: string; beforeBody: string; body: MathNode }
  | { kind: 'annotation'; command: string; above: boolean; beforeAnnotation: string; annotation: MathNode; beforeBase: string; base: MathNode }
  | { kind: 'largeOperator'; command: string; parts: OperatorPart[] }
  | { kind: 'generalOperator'; command: string; beforeBody: string; body: MathNode; parts: OperatorPart[] }
  | { kind: 'customSymbol'; command: string; assetSha256:string; parts:OperatorPart[] }
  | { kind: 'macroCall'; command: string; macroKind: BuiltinMacroKind; beforeArguments: string[]; arguments: MathNode[] }
  | { kind: 'style'; command: string; beforeBody: string; variant: MathVariant; body: MathNode }
  | { kind: 'space'; command: string; widthMu: number }
  | { kind: 'phantom'; command: string; beforeBody: string; axis: 'full' | 'horizontal' | 'vertical'; body: MathNode }
  | { kind: 'boxed'; beforeBody: string; body: MathNode }
  | { kind: 'color'; spec: ColorSpec; beforeBody: string; body: MathNode }
  | { kind: 'colorChange'; spec: ColorSpec }
  | { kind: 'styleChange'; command: string; style: MathStyle }
  | { kind: 'scripts'; base: MathNode; slots: ScriptSlot[] }
  | { kind: 'emptySlot'; label: string }
  | { kind: 'unparsedSource'; raw: string }
);
export interface SyntaxDiagnostic {
  code: 'missingArgument' | 'unclosedGroup' | 'unexpectedClosingBrace' | 'missingBase' | 'duplicateScript'
    | 'unsupportedColor' | 'unclosedColorModel' | 'unknownCommand' | 'unsupportedTextCommand' | 'unfinishedCommand' | 'unsupportedToken' | 'unbracedStructure' | 'unclosedOptionalArgument' | 'missingDelimiter' | 'unsupportedDelimiter' | 'unclosedFence' | 'unexpectedRightDelimiter' | 'unexpectedMiddleDelimiter'
    | 'missingEnvironmentName' | 'mismatchedEnvironment' | 'unclosedEnvironment' | 'unsupportedEnvironment'
    | 'unsupportedEnvironmentOption' | 'unexpectedEnvironmentEnd' | 'unsupportedCustomArtwork';
  severity: 'error' | 'unsupported';
  span: SourceSpan;
  message: string;
}
export interface SyntaxToken {
  kind: 'controlWord' | 'controlSymbol' | 'character' | 'openBrace' | 'closeBrace'
    | 'superscript' | 'subscript' | 'space' | 'comment';
  span: SourceSpan;
}
export interface SyntaxFacts {
  nodes: number; rows: number; fractions: number; radicals: number; fences: number; accents: number; annotations: number; largeOperators: number; generalOperators: number; styles: number; spaces: number; phantoms: number; boxes: number; colors: number; colorChanges: number; scripts: number;
  textRuns: number; emptySlots: number; unparsedSources: number;
  grids: number; gridRows: number; gridColumns: number; customSymbols:number;
}
export interface SourceAnalysis {
  symbolAssets?:Record<string,import('./custom-artwork').CustomSymbolAsset>;
  tree: MathNode;
  tokens: SyntaxToken[];
  diagnostics: SyntaxDiagnostic[];
  facts: SyntaxFacts;
  complete: boolean;
  supported: boolean;
  exportable: boolean;
}
