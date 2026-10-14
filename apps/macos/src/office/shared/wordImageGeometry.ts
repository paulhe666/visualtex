import type { FormulaLetterFont } from "../../editor/formulaFontPreferences";
import { svgToBase64 } from "../../export/runtime";
import type { SvgExportResult } from "../../export/exportTypes";

const referenceFontPt = 14;
const maxWidthPt = 500;

export function wordImageScales(font?: FormulaLetterFont) {
  return font === "times" ? { width: 1.067, height: 1 } : { width: 1.1, height: 1.1 };
}

export function wordImageReferenceGeometry(width: number, height: number, baseline: number, font?: FormulaLetterFont) {
  const scales = wordImageScales(font);
  const scale = Math.min(1, maxWidthPt / (width * 0.75 * scales.width));
  return {
    referenceWidthPt: width * 0.75 * scales.width * scale,
    referenceHeightPt: height * 0.75 * scales.height * scale,
    referenceBaselinePt: -(height - baseline) * 0.75 * scales.height * scale,
  };
}

/**
 * Word's VBA Font.Position only accepts whole points. Quantizing the position
 * of an unchanged tight image shifts its mathematical baseline by up to half
 * a point in opposite directions for adjacent formulas. Add transparent space
 * below the SVG instead: its descent becomes an exact whole point at the
 * requested size, while its glyphs, baseline, width and aspect remain intact.
 */
export function alignWordInlineSvg(svg: SvgExportResult, fontSizePt: number, font?: FormulaLetterFont): SvgExportResult {
  if (!Number.isFinite(fontSizePt) || fontSizePt < 1 || fontSizePt > 512) return svg;
  const geometry = wordImageReferenceGeometry(svg.width, svg.height, svg.baseline, font);
  const pointsPerPx = geometry.referenceHeightPt / svg.height * fontSizePt / referenceFontPt;
  const descentPt = (svg.height - svg.baseline) * pointsPerPx;
  const alignedDescentPt = Math.ceil(descentPt - 1e-9);
  const paddingPx = Math.max(0, (alignedDescentPt - descentPt) / pointsPerPx);
  if (paddingPx < 1e-8) return svg;
  const viewBox = svg.svg.match(/\bviewBox="([^"]+)"/);
  if (!viewBox) throw new Error("Word formula SVG has no viewBox.");
  const bounds = viewBox[1].trim().split(/\s+/).map(Number);
  if (bounds.length !== 4 || !bounds.every(Number.isFinite)) throw new Error("Word formula SVG viewBox is invalid.");
  const height = svg.height + paddingPx;
  bounds[3] *= height / svg.height;
  const source = svg.svg.replace(viewBox[0], `viewBox="${bounds.join(" ")}"`)
    .replace(/^(<svg\b[^>]*\bheight=")[^"]+"/, `$1${height}"`);
  return { ...svg, svg: source, base64: svgToBase64(source), height };
}
