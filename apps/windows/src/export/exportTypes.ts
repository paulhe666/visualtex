import type {
  FormulaChineseFont,
  FormulaLetterFont,
} from "../editor/formulaFontPreferences";

export interface SvgExportOptions {
  displayMode: boolean;
  fontSizePt: number;
  paddingPx: number;
  /** Optional horizontal padding override. Defaults to paddingPx. */
  paddingXPx?: number;
  /** Optional vertical padding override. Defaults to paddingPx. */
  paddingYPx?: number;
  background: "transparent" | "white";
  forceExplicitBlack?: boolean;
  formulaLetterFont?: FormulaLetterFont;
  formulaChineseFont?: FormulaChineseFont;
}

export interface SvgExportResult {
  svg: string;
  base64: string;
  width: number;
  height: number;
  baseline?: number;
}

export interface PngExportOptions {
  scale?: number;
  background?: "transparent" | "white" | `#${string}`;
}

export interface PngExportResult {
  blob: Blob;
  base64: string;
  width: number;
  height: number;
}
