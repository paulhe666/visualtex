export type LatexCodeFormat =
  | "raw"
  | "inline-dollar"
  | "inline-paren"
  | "display-dollar"
  | "display-bracket"
  | "equation"
  | "equation-star"
  | "align"
  | "align-star"
  | "aligned"
  | "gather"
  | "gather-star"
  | "multline"
  | "multline-star"
  | "equation-split"
  | "equation-star-split";

export type FormulaAlignment = "left" | "center" | "right";
export type Theme = "light" | "beige" | "dark" | "purple" | "green";

export interface FormulaLine {
  id: string;
  latex: string;
}

export interface FormulaBlock {
  id: string;
  latex: string;
  displayMode: "inline" | "block";
  alignment: FormulaAlignment;
  fontSize: number;
  createdAt: number;
  updatedAt: number;
}

export interface FormulaDocument {
  version: number;
  title: string;
  formulas: FormulaBlock[];
  macros: Record<string, string>;
  settings: {
    theme: Theme;
    zoom: number;
    formulaAlignment?: FormulaAlignment;
    latexCodeFormat?: LatexCodeFormat;
  };
}

export interface FormulaHistoryItem {
  id: string;
  latex: string;
  createdAt: number;
}

export interface InputBehaviorSettings {
  autoEscapeShortcuts: boolean;
  autoExitSuperscript: boolean;
  autoExitSubscript: boolean;
  autoExitAccent: boolean;
  autoExitWrapperCommand: boolean;
  showStructuredCommandSuggestions: boolean;
  showOtherCommandSuggestions: boolean;
}

export type InputBehaviorSettingKey =
  | "autoEscapeShortcuts"
  | "autoExitSuperscript"
  | "autoExitSubscript"
  | "autoExitAccent"
  | "autoExitWrapperCommand"
  | "showStructuredCommandSuggestions"
  | "showOtherCommandSuggestions";
