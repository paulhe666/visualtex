import type {
  FormulaChineseFont,
  FormulaLetterFont,
} from "../../editor/formulaFontPreferences";
import {
  createHeadlessCoreFormulaOwner,
  resolveCoreLayoutForPreferences,
} from "../../engine/coreEngineRuntime";
import { readCustomSymbolLibrary } from "../../math/customSymbolRegistry";
import type {
  OfficeExportResult,
  OfficeObjectMode,
} from "../shared/sessionClient";
import type { PreservedOfficeGeometry } from "./coreOfficeCompatibility";
import {
  generateCoreOfficeExport,
  rasterizeCoreOfficeExport,
} from "./coreOfficeCompatibility";

export interface HeadlessCoreOfficeConversionRequest {
  identity: string;
  latex: string;
  objectMode: OfficeObjectMode;
  displayMode: "inline" | "block";
  fontSizePt: number;
  formulaLetterFont: FormulaLetterFont;
  formulaChineseFont: FormulaChineseFont;
  preservedGeometry?: PreservedOfficeGeometry;
  equationTag?: string;
}

function coreDocumentId(identity: string) {
  const normalized = identity.trim();
  if (!normalized) throw new Error("Office conversion requires a stable request identity.");
  return `office-convert:${normalized}`;
}

const CSS_PX_PER_PT = 96 / 72;

function encodeUtf8Base64(value: string) {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(binary);
}

function escapeXmlText(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeXmlAttribute(value: string) {
  return escapeXmlText(value).replaceAll('"', "&quot;");
}

function escapeCoreText(value: string) {
  return value
    .replaceAll("\\", "\\textbackslash{}")
    .replaceAll("{", "\\{")
    .replaceAll("}", "\\}")
    .replaceAll("%", "\\%")
    .replaceAll("$", "\\$")
    .replaceAll("#", "\\#")
    .replaceAll("&", "\\&")
    .replaceAll("_", "\\_")
    .replaceAll("^", "\\^{}")
    .replaceAll("~", "\\~{}");
}

function wrapLabeledMathMl(mathMl: string, equationTag: string) {
  const trimmed = mathMl.trim();
  const match = /^<math\b[^>]*>([\s\S]*)<\/math>$/u.exec(trimmed);
  if (!match) throw new Error("Core Office equation-tag export requires a MathML root.");
  const text = escapeXmlText(equationTag);
  const id = escapeXmlAttribute(equationTag);
  return `<math xmlns="http://www.w3.org/1998/Math/MathML" display="block"><mtable displaystyle="true"><mlabeledtr><mtd id="mjx-eqn:${id}"><mtext>(${text})</mtext></mtd><mtd>${match[1]}</mtd></mlabeledtr></mtable></math>`;
}

interface ParsedSvgRoot {
  opening: string;
  inner: string;
  x: number;
  y: number;
  widthUnits: number;
  heightUnits: number;
}

function parseSvgRoot(svg: string): ParsedSvgRoot {
  const trimmed = svg.trim();
  const opening = /^<svg\b[^>]*>/u.exec(trimmed)?.[0];
  if (!opening || !trimmed.endsWith("</svg>")) {
    throw new Error("Core Office equation-tag export did not produce an SVG root.");
  }
  const viewBoxValue = /\bviewBox="([^"]+)"/u.exec(opening)?.[1] ?? "";
  const viewBox = viewBoxValue.trim().split(/[\s,]+/u).map(Number);
  if (viewBox.length !== 4 || !viewBox.every(Number.isFinite) || viewBox[2] <= 0 || viewBox[3] <= 0) {
    throw new Error("Core Office equation-tag SVG has invalid viewBox geometry.");
  }
  const inner = trimmed
    .slice(opening.length, -"</svg>".length)
    .replace(/<rect\b[^>]*\bfill-opacity="0\.001"[^>]*\/>/giu, "");
  return {
    opening,
    inner,
    x: viewBox[0],
    y: viewBox[1],
    widthUnits: viewBox[2],
    heightUnits: viewBox[3],
  };
}

function mergeEquationTagSvg(
  body: OfficeExportResult,
  label: OfficeExportResult,
  fontSizePt: number,
): OfficeExportResult {
  const bodySvg = parseSvgRoot(body.svg);
  const labelSvg = parseSvgRoot(label.svg);
  const fontSizePx = fontSizePt * CSS_PX_PER_PT;
  const gapUnits = 1;
  const bodyBaselineUnits = bodySvg.y + (body.baseline ?? body.height * 0.75) / fontSizePx;
  const labelBaselineUnits = labelSvg.y + (label.baseline ?? label.height * 0.75) / fontSizePx;
  const labelTranslateX = bodySvg.x + bodySvg.widthUnits + gapUnits - labelSvg.x;
  const labelTranslateY = bodyBaselineUnits - labelBaselineUnits;
  const labelTop = labelSvg.y + labelTranslateY;
  const labelBottom = labelTop + labelSvg.heightUnits;
  const finalY = Math.min(bodySvg.y, labelTop);
  const finalBottom = Math.max(bodySvg.y + bodySvg.heightUnits, labelBottom);
  const finalWidthUnits = bodySvg.widthUnits + gapUnits + labelSvg.widthUnits;
  const finalHeightUnits = finalBottom - finalY;
  const width = finalWidthUnits * fontSizePx;
  const height = finalHeightUnits * fontSizePx;
  const baseline = (bodyBaselineUnits - finalY) * fontSizePx;
  const hitTarget = `<rect x="${bodySvg.x}" y="${finalY}" width="${finalWidthUnits}" height="${finalHeightUnits}" fill="#000000" fill-opacity="0.001"/>`;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="${bodySvg.x} ${finalY} ${finalWidthUnits} ${finalHeightUnits}">${hitTarget}${bodySvg.inner}<g transform="translate(${labelTranslateX} ${labelTranslateY})">${labelSvg.inner}</g></svg>`;
  return {
    ...body,
    svg,
    svgBase64: encodeUtf8Base64(svg),
    width,
    height,
    baseline,
  };
}

function requiresPng(objectMode: OfficeObjectMode) {
  // MathType creation/rewrite consumes Presentation MathML plus the SVG-derived
  // vector preview. Word OMML consumes MathML only. Native VisualTeX OLE and the
  // cross-platform picture path retain the existing PNG contract.
  return objectMode === "nativeOle" || objectMode === "crossPlatformPicture";
}

export async function generateHeadlessCoreOfficeExport(
  request: HeadlessCoreOfficeConversionRequest,
): Promise<OfficeExportResult> {
  if (!request.latex.trim()) {
    throw new Error("Office conversion cannot export an empty formula.");
  }

  const symbols = structuredClone(readCustomSymbolLibrary().symbols);
  const prepared = await createHeadlessCoreFormulaOwner(
    coreDocumentId(request.identity),
    request.latex,
    symbols,
  );

  try {
    const outputLetterFont: FormulaLetterFont =
      request.objectMode === "mathTypeOle"
        ? "times"
        : request.formulaLetterFont;
    const style: "display" | "text" =
      request.objectMode === "mathTypeOle" || request.displayMode === "block"
        ? "display"
        : "text";

    // The Word OMML path is intentionally MathML-only. Presentation MathML is
    // structural and does not need a system-font lookup or layout snapshot.
    const resolvedLayout = request.objectMode === "wordOmml"
      ? { ...prepared.context.baseLayout, style }
      : await resolveCoreLayoutForPreferences(
          prepared.context,
          outputLetterFont,
          request.formulaChineseFont,
          style,
        );
    // The hidden converter has no interactive font surface. In a plain-browser
    // acceptance host there is deliberately no trusted system-font bridge, so a
    // MathType output preference such as `times` cannot resolve to installed
    // font bytes. Keep the Office output metadata (`formulaLetterFont=times`)
    // while using the bundled, SHA-verified core math font for conversion
    // geometry. Real Tauri/WebView2 hosts still resolve the requested system
    // math face first and use the same bundled font only as the P12 fallback.
    // Interactive CoreMathEditor sessions do not use this compatibility
    // fallback and continue to report non-Tauri system-font preferences as
    // unresolved.
    const layout = resolvedLayout ?? { ...prepared.context.baseLayout, style };

    let base = generateCoreOfficeExport({
      owner: prepared.owner,
      layout,
      objectMode: request.objectMode,
      displayMode: request.displayMode,
      fontSizePt: request.fontSizePt,
      formulaLetterFont: outputLetterFont,
      formulaChineseFont: request.formulaChineseFont,
      preservedGeometry: request.preservedGeometry,
    });

    const equationTag =
      request.displayMode === "block" ? request.equationTag?.trim() ?? "" : "";
    if (equationTag) {
      base = {
        ...base,
        mathMl: wrapLabeledMathMl(base.mathMl ?? "", equationTag),
      };
      if (base.svg) {
        const labelPrepared = await createHeadlessCoreFormulaOwner(
          `${coreDocumentId(request.identity)}:equation-tag`,
          `\\text{(${escapeCoreText(equationTag)})}`,
          symbols,
        );
        try {
          const labelBase = generateCoreOfficeExport({
            owner: labelPrepared.owner,
            layout,
            objectMode: request.objectMode,
            displayMode: "block",
            fontSizePt: request.fontSizePt,
            formulaLetterFont: outputLetterFont,
            formulaChineseFont: request.formulaChineseFont,
          });
          base = mergeEquationTagSvg(base, labelBase, request.fontSizePt);
        } finally {
          labelPrepared.owner.destroy();
        }
      }
    }

    if (!requiresPng(request.objectMode)) return base;

    return await rasterizeCoreOfficeExport(base, async (svg) => {
      const { svgToPng } = await import("../../export/svgToPng");
      return (
        await svgToPng(
          {
            base64: svg.svgBase64,
            width: svg.width,
            height: svg.height,
          },
          { scale: 2, background: "transparent" },
        )
      ).base64;
    });
  } finally {
    prepared.owner.destroy();
  }
}
