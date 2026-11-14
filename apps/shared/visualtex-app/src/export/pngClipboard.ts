import { invoke, isTauri } from "@tauri-apps/api/core";
import { APP_PLATFORM } from "@platform/appInfo";
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
  /** The formula's ink (#rrggbb); near-black when absent. */
  color?: string;
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

  const svg = await latexToSvg(nonEmptyFormulas.join("\n"), {
    displayMode: true,
    fontSizePt: 18,
    paddingPx: 18,
    background: "transparent",
    formulaLetterFont: preferences.formulaLetterFont,
    formulaChineseFont: preferences.formulaChineseFont,
    color: preferences.color,
  });
  return svgToPng(svg, {
    scale: 3,
    background: normalizePngExportBackground(preferences.background),
  });
}

async function writeBrowserPngClipboard(blob: Blob) {
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
 * foregrounded, so Windows uses that browser path; macOS writes the PNG through
 * its AppKit clipboard command (WKWebView's image clipboard is not reliable).
 */
export async function copyFormulaDocumentPngToClipboard(
  formulas: readonly string[],
  preferences: FormulaPngRenderPreferences,
): Promise<PngExportResult> {
  const png = await renderFormulaDocumentPng(formulas, preferences);
  if (APP_PLATFORM === "macos" && isTauri()) {
    await invoke("copy_png_to_clipboard", {
      dataBase64: png.base64,
    });
  } else {
    await writeBrowserPngClipboard(png.blob);
  }
  return png;
}
