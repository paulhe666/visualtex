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
  };
}

export interface FormulaHistoryItem {
  id: string;
  latex: string;
  createdAt: number;
}
