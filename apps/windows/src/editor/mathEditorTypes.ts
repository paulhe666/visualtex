import type { LatexCommand } from "../types/command";
import type { CoreEditorExportPort } from "../engine/coreEditorExportPort";
import type {
  CoreMathSelectionSnapshot,
  FormulaEditSource,
  MathSelectionSnapshot,
} from "../history/historyTypes";

export interface MathEditorInsertionTarget {
  lineId: string;
  ranges: Array<[number, number]>;
  direction: "forward" | "backward" | "none";
  /** Core targets carry an opaque one-shot request object. The optional legacy
   * discriminator remains only in the shared archive/read boundary during P15. */
  engine?: "legacy" | "core";
  coreRequest?: object;
}

export interface MathEditorSelectionTarget {
  selections: MathEditorInsertionTarget[];
  engine?: "legacy" | "core";
  coreLineId?: string;
}

export type MathEditorSelectionStyle =
  | { kind: "bold" }
  | { kind: "italic" }
  | { kind: "color"; value: string }
  | { kind: "backgroundColor"; value: string };

export interface MathEditorFocusOptions {
  target?: "active" | "first" | "last";
  moveToEnd?: boolean;
}

export interface MathEditorHistoryState {
  engine: "core";
  canUndo: boolean;
  canRedo: boolean;
}

export interface MathEditorHandle {
  insertCommand: (command: LatexCommand, source?: "toolbar" | "history" | "shortcut") => void;
  insertLatex: (latex: string, source?: FormulaEditSource) => void;
  insertLatexAt: (
    target: MathEditorInsertionTarget,
    latex: string,
    source?: FormulaEditSource,
  ) => boolean;
  appendLatex: (latex: string, source?: FormulaEditSource) => void;
  focus: (options?: MathEditorFocusOptions) => void;
  addLine: () => void;
  commitPendingTransaction: () => void;
  getSelectionMap: () => Record<string, MathSelectionSnapshot>;
  getCoreSelectionMap?: () => Record<string, CoreMathSelectionSnapshot>;
  restoreSelection: (
    lineId: string,
    latex: string,
    selection: MathSelectionSnapshot | null,
  ) => Promise<boolean>;
  restoreCoreSelection?: (
    lineId: string,
    latex: string,
    selection: CoreMathSelectionSnapshot | null,
  ) => Promise<boolean>;
  captureSelectionTarget: () => MathEditorSelectionTarget | null;
  captureInsertionTarget: () => MathEditorInsertionTarget | null;
  applySelectionStyle: (
    style: MathEditorSelectionStyle,
    target?: MathEditorSelectionTarget | null,
  ) => boolean;
  refreshLayout: () => void;
  replaceLineSource?: (lineId: string, latex: string) => boolean;
  getCoreExportPort?: () => CoreEditorExportPort | null;
  history?: (redo: boolean) => boolean;
}
