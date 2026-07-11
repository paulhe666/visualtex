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

export interface FormulaLine {
  id: string;
  latex: string;
}

export interface FormulaBlock {
  id: string;
  latex: string;
  displayMode: "inline" | "block";
  alignment: "left" | "center" | "right";
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
    theme: "light" | "dark";
    zoom: number;
    latexCodeFormat?: LatexCodeFormat;
  };
}

export interface FormulaHistoryItem {
  id: string;
  latex: string;
  createdAt: number;
}
