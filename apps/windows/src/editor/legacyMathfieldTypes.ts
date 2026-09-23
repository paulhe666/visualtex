export type LegacyMacroDictionary = Record<string, any>;

/**
 * Structural compatibility surface retained only for the isolated legacy
 * MathLive oracle. Production core code must not import MathLive types.
 */
export interface LegacyMathfieldLike extends HTMLElement {
  value: string;
  macros: LegacyMacroDictionary;
  readonly lastOffset: number;
  selection: {
    ranges: Array<[number, number]>;
    direction?: "forward" | "backward" | "none" | string;
  };
  getValue(...args: unknown[]): string;
  setValue(value: string, options?: Record<string, unknown>): void;
  resetUndo(): void;
}
