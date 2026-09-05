import { latexToSvg, svgToPng } from "./runtime";
import {
  normalizePngExportBackground,
  type PngExportBackground,
} from "./pngBackground";
import type { PngExportResult } from "./exportTypes";
import type {
  FormulaChineseFont,
  FormulaLetterFont,
} from "../editor/formulaFontPreferences";

export interface FormulaPngRenderPreferences {
  background: PngExportBackground;
  formulaLetterFont?: FormulaLetterFont;
  formulaChineseFont?: FormulaChineseFont;
}

export async function renderFormulaDocumentPng(
  formulas: readonly string[],
  preferences: FormulaPngRenderPreferences,
): Promise<PngExportResult> {
  const nonEmptyFormulas = formulas
    .map((formula) => formula.trim())
    .filter(Boolean);
  if (!nonEmptyFormulas.length) {
    throw new Error("There is no formula to export.");
  }

  const svg = latexToSvg(nonEmptyFormulas.join("\n"), {
    displayMode: true,
    fontSizePt: 18,
    paddingPx: 18,
    background: "transparent",
    formulaLetterFont: preferences.formulaLetterFont,
    formulaChineseFont: preferences.formulaChineseFont,
  });
  return svgToPng(svg, {
    scale: 3,
    background: normalizePngExportBackground(preferences.background),
  });
}

async function writePngClipboard(blob: Blob) {
  if (
    typeof navigator === "undefined" ||
    !navigator.clipboard?.write ||
    typeof ClipboardItem === "undefined"
  ) {
    throw new Error("PNG clipboard access is unavailable in this environment.");
  }
  await navigator.clipboard.write([
    new ClipboardItem({
      "image/png": blob,
    }),
  ]);
}

/**
 * WebView2 exposes the standard image clipboard API while VisualTeX is
 * foregrounded. Keep the Windows implementation on that native browser path
 * instead of importing the macOS AppKit clipboard command.
 */
export async function copyFormulaDocumentPngToClipboard(
  formulas: readonly string[],
  preferences: FormulaPngRenderPreferences,
): Promise<PngExportResult> {
  const png = await renderFormulaDocumentPng(formulas, preferences);
  await writePngClipboard(png.blob);
  return png;
}
