export type LatexCodeFormat =
  | "raw"
  | "mixed-inline-display"
  | "inline-dollar"
  | "inline-text-double-dollar"
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
export type Theme =
  | "light"
  | "beige"
  | "dark"
  | "purple"
  | "green"
  | "codex"
  | "notion"
  | "one"
  | "proof"
  | "raycast"
  | "rose-pine"
  | "solarized"
  | "vercel"
  | "vscode-plus"
  | "xcode"
  | "custom";

export type FormulaLineMode = "inline" | "display";

export interface FormulaLine {
  id: string;
  latex: string;
  mode?: FormulaLineMode;
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
    editorLayout?: "standard" | "classic";
    language?: "cn" | "en";
    sourceOpen?: boolean;
    autoPairDelimiters?: boolean;
    showLineNumbers?: boolean;
    highlightActiveLine?: boolean;
    formulaInsetLeft?: number;
    formulaInsetRight?: number;
    formulaToolButtonSize?: number;
    formulaToolButtonPadding?: number;
    formulaRowVerticalInset?: number;
    pngExportBackground?: "transparent" | `#${string}`;
    formulaLetterFont?:
      | "katex"
      | "times"
      | "cambria"
      | "stix"
      | "palatino"
      | "helvetica";
    formulaChineseFont?: "system" | "pingfang" | "songti" | "kaiti" | "heiti";
    inputBehavior?: InputBehaviorSettings;
    personalize?: boolean;
    suggestionCount?: number;
    checkUpdatesOnStartup?: boolean;
    powerPointDefaultFontSizePt?: number;
    classicTileWidth?: number;
    classicDockHeight?: number;
    keypadMinimizeOnCopy?: boolean;
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
