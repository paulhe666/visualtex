import type { SourceSpan } from './syntax';
export interface SourceEdit { startUtf16: number; endUtf16: number; replacement: string }
export type ParseScope =
  | { kind: 'none' }
  | { kind: 'subtree'; nodeId: number; parsedNodes: number }
  | { kind: 'formula'; parsedNodes: number; reason: string };
export interface SourceChangeMap {
  oldRange: SourceSpan;
  newRange: SourceSpan;
  retainedIds: number[];
  removedIds: number[];
  insertedIds: number[];
  changedIds: number[];
  parseScope: ParseScope;
  lexedTokens: number;
}
