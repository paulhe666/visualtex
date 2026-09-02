import { mathjax } from "mathjax-full/js/mathjax.js";
import { TeX } from "mathjax-full/js/input/tex.js";
import { SVG } from "mathjax-full/js/output/svg.js";
import { liteAdaptor } from "mathjax-full/js/adaptors/liteAdaptor.js";
import { RegisterHTMLHandler } from "mathjax-full/js/handlers/html.js";
import { AllPackages } from "mathjax-full/js/input/tex/AllPackages.js";
import { STATE } from "mathjax-full/js/core/MathItem.js";
import { SerializedMmlVisitor } from "mathjax-full/js/core/MmlTree/SerializedMmlVisitor.js";
import type { MmlNode } from "mathjax-full/js/core/MmlTree/MmlNode.js";
import { normalizeMathLiveCanonicalUprightCommands } from "../editor/normalizeChineseLatex.ts";
import { normalizeExtendedIntegralLatexCommands } from "../math/extendedIntegralCompatibility.ts";
import {
  isSingleCompleteLatexEnvironment,
  unwrapSingleLatexDisplayMath,
} from "../math/latexEnvironment.ts";
import { applyVisualTexIntegralSvgGlyphs } from "../math/integralSvgExportCompatibility.ts";
import {
  applyCustomSymbolArtworkToSvg,
  expandCustomSymbolsForMathMl,
  expandCustomSymbolsForSvg,
} from "../math/customSymbolRendering.ts";
import {
  assertResolvedMathJaxSvg,
  assertResolvedPresentationMathMl,
  VISUALTEX_MATHML_MACROS,
  VISUALTEX_SVG_MACROS,
  type VisualTexMathJaxMacro,
} from "../math/latexCompatibility.ts";
import type {
  PngExportOptions,
  PngExportResult,
  SvgExportOptions,
  SvgExportResult,
} from "./exportTypes";
import { errorMessage } from "../runtime/errorMessage.ts";
import {
  DEFAULT_FORMULA_CHINESE_FONT,
  DEFAULT_FORMULA_LETTER_FONT,
  formulaChineseFontFamily,
  formulaLetterFontFamilies,
  normalizeFormulaChineseFont,
  normalizeFormulaLetterFont,
} from "../editor/formulaFontPreferences.ts";

const DEFAULT_OPTIONS: SvgExportOptions = {
  displayMode: true,
  fontSizePt: 12,
  paddingPx: 8,
  background: "transparent",
};

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);

function createTexInput(macros: Record<string, VisualTexMathJaxMacro>) {
  return new TeX({
    packages: AllPackages,
    macros,
    formatError: (_jax: unknown, error: unknown) => {
      throw new Error(
        errorMessage(error, "MathJax could not parse this formula."),
        { cause: error },
      );
    },
  });
}

const mathMlTexInput = createTexInput(VISUALTEX_MATHML_MACROS);
const svgTexInput = createTexInput(VISUALTEX_SVG_MACROS);
const mathMlOutput = new SVG({
  fontCache: "local",
  internalSpeechTitles: false,
});
const svgOutput = new SVG({
  fontCache: "local",
  internalSpeechTitles: false,
});
const mathMlDocument = mathjax.document("", {
  InputJax: mathMlTexInput,
  OutputJax: mathMlOutput,
});
const mathDocument = mathjax.document("", {
  InputJax: svgTexInput,
  OutputJax: svgOutput,
});
const serializedMmlVisitor = new SerializedMmlVisitor(
  mathMlDocument.mmlFactory,
);

function positiveFinite(value: number, fallback: number) {
  return Number.isFinite(value) && value > 0 ? value : fallback;
}

function nonNegativeFinite(value: number, fallback: number) {
  return Number.isFinite(value) && value >= 0 ? value : fallback;
}

function prepareLatex(latex: string) {
  let normalized = normalizeMathLiveCanonicalUprightCommands(
    normalizeExtendedIntegralLatexCommands(latex.replace(/\r\n?/g, "\n")),
  ).trim();
  if (!normalized) throw new Error("Cannot export an empty formula.");

  // VisualTeX source formats legitimately serialize display formulas as
  // `\\[ ... \\]`. MathJax's direct conversion API is already invoked in
  // display mode, so those source delimiters must not become literal glyphs or
  // be split into extra rows around an inner aligned/align environment.
  normalized = unwrapSingleLatexDisplayMath(normalized) ?? normalized;

  const lines = normalized
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  if (lines.length <= 1) return normalized;

  // Preserve a source string that is itself one complete TeX environment.
  // A document with multiple VisualTeX formula rows may still contain an
  // inner matrix/cases environment on one row; that must not make all rows
  // collapse into a single horizontal TeX expression.
  if (isSingleCompleteLatexEnvironment(normalized)) return normalized;

  // `aligned` uses a right/left pair around every alignment marker. Without
  // an explicit marker MathJax right-aligns rows of different widths. Keep
  // the whole document as one image, but anchor every formula row on its left
  // edge and preserve the editor's vertical ordering.
  return `\\begin{aligned}${lines.map((line) => `&${line}`).join("\\\\")}\\end{aligned}`;
}

function extractSvg(markup: string) {
  const start = markup.indexOf("<svg");
  const end = markup.lastIndexOf("</svg>");
  if (start < 0 || end < start) {
    throw new Error("MathJax did not produce an SVG element.");
  }
  return markup.slice(start, end + "</svg>".length);
}

function parseViewBox(svg: string) {
  const match = svg.match(
    /\bviewBox=["']\s*([-+\d.eE]+)\s+([-+\d.eE]+)\s+([-+\d.eE]+)\s+([-+\d.eE]+)\s*["']/,
  );
  if (!match) throw new Error("Exported SVG is missing a valid viewBox.");
  const values = match.slice(1).map(Number);
  if (values.some((value) => !Number.isFinite(value))) {
    throw new Error("Exported SVG has an invalid viewBox.");
  }
  const [x, y, width, height] = values;
  if (width <= 0 || height <= 0) {
    throw new Error("Exported SVG has non-positive dimensions.");
  }
  return { x, y, width, height };
}

function assertSelfContained(svg: string) {
  if (/<foreignObject\b/i.test(svg)) {
    throw new Error("SVG export must not contain foreignObject.");
  }
  if (/<link\b|@import\b/i.test(svg)) {
    throw new Error("SVG export must not depend on external CSS.");
  }
  if (/\b(?:href|xlink:href)=["'](?!#|data:)[^"']+/i.test(svg)) {
    throw new Error("SVG export contains an external resource reference.");
  }
  if (/url\(\s*["']?https?:/i.test(svg)) {
    throw new Error("SVG export contains a remote URL.");
  }
}

function escapeSvgAttribute(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeSvgText(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function mathAlphabetBaseCharacter(codePointHex: string) {
  const codePoint = Number.parseInt(codePointHex, 16);
  if (!Number.isFinite(codePoint)) return "";
  const source = String.fromCodePoint(codePoint);
  const normalized = source.normalize("NFKD");
  const characters = Array.from(normalized);
  if (characters.length !== 1 || !/\p{L}/u.test(characters[0])) return "";
  return characters[0];
}

function applyVisualTexSvgFontPreferences(
  svg: string,
  options: SvgExportOptions,
) {
  const letterFont = normalizeFormulaLetterFont(
    options.formulaLetterFont ?? DEFAULT_FORMULA_LETTER_FONT,
  );
  const chineseFont = normalizeFormulaChineseFont(
    options.formulaChineseFont ?? DEFAULT_FORMULA_CHINESE_FONT,
  );
  let output = svg;

  const chineseFamily = escapeSvgAttribute(formulaChineseFontFamily(chineseFont));
  output = output.replace(
    /(<g\b[^>]*data-mml-node=["']mtext["'][^>]*>)([\s\S]*?)(<\/g>)/gi,
    (_whole, opening: string, body: string, closing: string) =>
      `${opening}${body.replace(
        /font-family=["'][^"']*["']/gi,
        `font-family="${chineseFamily}" data-visualtex-output-text-font="${escapeSvgAttribute(chineseFont)}"`,
      )}${closing}`,
  );

  if (letterFont === DEFAULT_FORMULA_LETTER_FONT) return output;

  const families = formulaLetterFontFamilies(letterFont);
  output = output.replace(/<use\b([^>]*)><\/use>/gi, (whole, attributes: string) => {
    const codePoint = attributes.match(/\bdata-c=["']([0-9A-F]+)["']/i)?.[1];
    const href = attributes.match(/\bxlink:href=["']#([^"']+)["']/i)?.[1] ?? "";
    const variant = href.match(/-TEX-(BI|B|I|N)-[0-9A-F]+$/i)?.[1]?.toUpperCase();
    if (!codePoint || !variant) return whole;
    const character = mathAlphabetBaseCharacter(codePoint);
    if (!character) return whole;

    const italic = variant === "I" || variant === "BI";
    const bold = variant === "B" || variant === "BI";
    const family = escapeSvgAttribute(italic ? families.italic : families.upright);
    const originalTransform = attributes.match(/\btransform=["']([^"']+)["']/i)?.[1]?.trim();
    const originalX = attributes.match(/\bx=["']([^"']+)["']/i)?.[1]?.trim();
    const originalY = attributes.match(/\by=["']([^"']+)["']/i)?.[1]?.trim();
    // MathJax often lays out multi-letter operators such as \\arccos or
    // \\operatorname{rank} as several <use> glyphs inside one translated
    // parent. Every glyph after the first then owns an additional local
    // translate(...). Dropping that transform while replacing <use> with a
    // system-font <text> puts every character at the same origin and produces
    // the severe overlap/"garbled" appearance seen in Word. Preserve all
    // glyph-local positioning before applying the y-axis flip required for SVG
    // text inside MathJax's outer scale(1,-1) coordinate system.
    const transform = originalTransform
      ? `${escapeSvgAttribute(originalTransform)} scale(1,-1)`
      : "scale(1,-1)";
    const position = `${
      originalX ? ` x="${escapeSvgAttribute(originalX)}"` : ""
    }${originalY ? ` y="${escapeSvgAttribute(originalY)}"` : ""}`;
    return `<text data-c="${codePoint}" data-visualtex-output-letter-font="${escapeSvgAttribute(letterFont)}"${position} transform="${transform}" font-size="1000px" font-family="${family}"${italic ? ' font-style="italic"' : ""}${bold ? ' font-weight="700"' : ""}>${escapeSvgText(character)}</text>`;
  });
  return output;
}

const WORD_EXPLICIT_BLACK = "#000000";

function wordCompatiblePaintValue(value: string) {
  const trimmed = value.trim();
  if (/^(?:none|transparent)$/i.test(trimmed)) {
    return trimmed.toLowerCase();
  }
  return WORD_EXPLICIT_BLACK;
}

function removeCssCustomProperties(value: string) {
  return value.replace(
    /(^|[;{])\s*--[a-zA-Z0-9_-]+\s*:[^;}]*;?/g,
    "$1",
  );
}

function forceStylePaintBlack(value: string) {
  return removeCssCustomProperties(value).replace(
    /(^|[;{]\s*)(color|fill|stroke)\s*:\s*([^;}]+)/gi,
    (_match, prefix: string, property: string, paint: string) =>
      `${prefix}${property}:${wordCompatiblePaintValue(paint)}`,
  );
}

/**
 * Word 16.89 can initially paint an SVG formula as transparent when its first
 * resolved colour comes from currentColor, a CSS variable, a white inherited
 * paint, or another deferred style carrier. Normalize every SVG paint carrier
 * before either the SVG or PNG is emitted so both compatibility representations
 * are byte-for-byte derived from the same explicit-black artwork.
 */
function forceWordCompatibleBlack(svg: string) {
  let output = svg.replace(/currentColor/gi, WORD_EXPLICIT_BLACK);
  output = output.replace(
    /\b(color|fill|stroke)=(['"])(.*?)\2/gi,
    (_match, property: string, quote: string, paint: string) =>
      `${property}=${quote}${wordCompatiblePaintValue(paint)}${quote}`,
  );
  output = output.replace(
    /\bstyle=(['"])(.*?)\1/gi,
    (_match, quote: string, style: string) =>
      `style=${quote}${forceStylePaintBlack(style)}${quote}`,
  );
  output = output.replace(
    /<style\b([^>]*)>([\s\S]*?)<\/style>/gi,
    (_match, attributes: string, css: string) =>
      `<style${attributes}>${forceStylePaintBlack(css)}</style>`,
  );

  const lower = output.toLowerCase();
  if (
    lower.includes("currentcolor") ||
    lower.includes("var(") ||
    /\b(?:color|fill|stroke)\s*[:=]\s*['"]?(?:inherit|white|#fff(?:fff)?)(?:['";\s>]|$)/i.test(
      output,
    )
  ) {
    throw new Error(
      "Word SVG export still contains a deferred or white paint style.",
    );
  }
  if (!/\b(?:fill|stroke)=["']#000000["']/i.test(output)) {
    throw new Error("Word SVG export is missing explicit black formula paint.");
  }
  return output;
}

function encodeUtf8Base64(value: string) {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary);
}

export function svgToBase64(svg: string) {
  return encodeUtf8Base64(svg);
}

export function latexToMathMl(latex: string, displayMode = true) {
  const source = expandCustomSymbolsForMathMl(prepareLatex(latex));
  const root = mathMlDocument.convert(source, {
    display: displayMode,
    end: STATE.COMPILED,
  }) as unknown as MmlNode;
  const mathMl = serializedMmlVisitor.visitTree(root).trim();
  if (!mathMl.startsWith("<math") || !mathMl.includes("MathML")) {
    throw new Error("MathJax did not produce valid Presentation MathML.");
  }
  assertResolvedPresentationMathMl(mathMl);
  return mathMl;
}

export function latexToSvg(
  latex: string,
  options: SvgExportOptions = DEFAULT_OPTIONS,
): SvgExportResult {
  const source = expandCustomSymbolsForSvg(prepareLatex(latex));
  const fontSizePt = positiveFinite(options.fontSizePt, DEFAULT_OPTIONS.fontSizePt);
  const paddingPx = nonNegativeFinite(options.paddingPx, DEFAULT_OPTIONS.paddingPx);
  const fontSizePx = fontSizePt * (96 / 72);
  const exPx = fontSizePx * 0.442;

  const container = mathDocument.convert(source, {
    display: options.displayMode,
    em: fontSizePx,
    ex: exPx,
    containerWidth: 100_000,
  });
  let svg = extractSvg(adaptor.outerHTML(container));
  svg = applyVisualTexIntegralSvgGlyphs(svg, options.displayMode);
  svg = applyCustomSymbolArtworkToSvg(svg);
  svg = applyVisualTexSvgFontPreferences(svg, options);
  const viewBox = parseViewBox(svg);

  const unitsPerPx = 1000 / fontSizePx;
  const paddingUnits = paddingPx * unitsPerPx;
  const padded = {
    x: viewBox.x - paddingUnits,
    y: viewBox.y - paddingUnits,
    width: viewBox.width + 2 * paddingUnits,
    height: viewBox.height + 2 * paddingUnits,
  };
  const width = Math.max(1, padded.width / unitsPerPx);
  const height = Math.max(1, padded.height / unitsPerPx);
  const baseline = Math.max(0, Math.min(height, -padded.y / unitsPerPx));

  svg = svg
    .replace(
      /\bviewBox=["'][^"']+["']/,
      `viewBox="${padded.x} ${padded.y} ${padded.width} ${padded.height}"`,
    )
    .replace(/^<svg\b([^>]*)>/, (_opening, rawAttributes: string) => {
      const attributes = rawAttributes
        .replace(
          /\s(?:xmlns|width|height|role|focusable|style)=["'][^"']*["']/g,
          "",
        )
        .trim();
      return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" role="img" focusable="false"${
        attributes ? ` ${attributes}` : ""
      }>`;
    })
    .replace(
      /currentColor/gi,
      options.forceExplicitBlack ? WORD_EXPLICIT_BLACK : "#111111",
    );

  const openingEnd = svg.indexOf(">");
  if (options.background === "white") {
    const background = `<rect x="${padded.x}" y="${padded.y}" width="${padded.width}" height="${padded.height}" fill="#ffffff"/>`;
    svg = `${svg.slice(0, openingEnd + 1)}${background}${svg.slice(openingEnd + 1)}`;
  } else {
    // PowerPoint otherwise hit-tests only the painted glyph paths of a
    // transparent SVG. A practically invisible filled rectangle makes the
    // entire formula bounds selectable and double-clickable at normal zoom.
    const hitTarget = `<rect x="${padded.x}" y="${padded.y}" width="${padded.width}" height="${padded.height}" fill="#000000" fill-opacity="0.001"/>`;
    svg = `${svg.slice(0, openingEnd + 1)}${hitTarget}${svg.slice(openingEnd + 1)}`;
  }

  if (options.forceExplicitBlack) {
    svg = forceWordCompatibleBlack(svg);
  }

  assertResolvedMathJaxSvg(svg);
  assertSelfContained(svg);
  return {
    svg,
    base64: svgToBase64(svg),
    width,
    height,
    baseline,
  };
}

function blobToBase64(blob: Blob) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error ?? new Error("Unable to read PNG blob."));
    reader.onload = () => {
      const result = typeof reader.result === "string" ? reader.result : "";
      const comma = result.indexOf(",");
      resolve(comma >= 0 ? result.slice(comma + 1) : result);
    };
    reader.readAsDataURL(blob);
  });
}

function pngDataUrlToBlob(value: string) {
  const prefix = "data:image/png;base64,";
  if (!value.startsWith(prefix)) {
    throw new Error("Canvas did not produce a PNG data URL.");
  }
  const binary = atob(value.slice(prefix.length));
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new Blob([bytes], { type: "image/png" });
}

async function encodeCanvasPng(canvas: HTMLCanvasElement) {
  if (typeof canvas.toBlob === "function") {
    const blob = await new Promise<Blob | null>((resolve) => {
      let settled = false;
      const finish = (value: Blob | null) => {
        if (settled) return;
        settled = true;
        window.clearTimeout(timeout);
        resolve(value);
      };
      const timeout = window.setTimeout(() => finish(null), 750);
      try {
        canvas.toBlob(finish, "image/png");
      } catch {
        finish(null);
      }
    });
    if (blob) return blob;
  }

  // WKWebView can expose canvas.toBlob() but return null for SVG-backed
  // canvases. toDataURL() uses a different WebKit encoding path and is stable
  // on the same canvas, so use it as the required Word compatibility fallback.
  return pngDataUrlToBlob(canvas.toDataURL("image/png"));
}

export async function svgToPng(
  svgResult: Pick<SvgExportResult, "base64" | "width" | "height">,
  options: PngExportOptions = {},
): Promise<PngExportResult> {
  if (typeof document === "undefined" || typeof Image === "undefined") {
    throw new Error("PNG export requires a browser canvas environment.");
  }

  const scale = positiveFinite(options.scale ?? 2, 2);
  const width = Math.max(1, Math.ceil(svgResult.width * scale));
  const height = Math.max(1, Math.ceil(svgResult.height * scale));
  const image = new Image();
  image.decoding = "async";
  const source = `data:image/svg+xml;base64,${svgResult.base64}`;

  await new Promise<void>((resolve, reject) => {
    image.onload = () => resolve();
    image.onerror = () => reject(new Error("Unable to rasterize the generated SVG."));
    image.src = source;
  });

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Unable to create a PNG canvas context.");
  const requestedBackground = options.background ?? "transparent";
  const opaqueBackground =
    requestedBackground === "white" ? "#ffffff" : requestedBackground;
  if (opaqueBackground !== "transparent") {
    context.fillStyle = opaqueBackground;
    context.fillRect(0, 0, width, height);
  }
  context.drawImage(image, 0, 0, width, height);

  const backgroundRgb = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(
    opaqueBackground,
  );
  const backgroundChannels = backgroundRgb
    ? backgroundRgb.slice(1).map((channel) => Number.parseInt(channel, 16))
    : null;
  const pixels = context.getImageData(0, 0, width, height).data;
  let inkTop = height;
  let inkBottom = -1;
  for (let index = 0; index < pixels.length; index += 4) {
    const alpha = pixels[index + 3];
    if (alpha < 16) continue;
    const differsFromBackground = backgroundChannels
      ? Math.abs(pixels[index] - backgroundChannels[0]) > 10 ||
        Math.abs(pixels[index + 1] - backgroundChannels[1]) > 10 ||
        Math.abs(pixels[index + 2] - backgroundChannels[2]) > 10
      : pixels[index] < 245 ||
        pixels[index + 1] < 245 ||
        pixels[index + 2] < 245;
    if (!differsFromBackground) continue;
    const row = Math.floor(index / 4 / width);
    if (row < inkTop) inkTop = row;
    if (row > inkBottom) inkBottom = row;
  }
  if (inkBottom < inkTop) {
    throw new Error("PNG rasterization produced no visible formula ink.");
  }
  const inkTopRatio = inkTop / height;
  const inkBottomRatio = (inkBottom + 1) / height;
  const inkCenterYRatio = (inkTopRatio + inkBottomRatio) / 2;

  const blob = await encodeCanvasPng(canvas);
  return {
    blob,
    base64: await blobToBase64(blob),
    width,
    height,
    inkTopRatio,
    inkBottomRatio,
    inkCenterYRatio,
  };
}
