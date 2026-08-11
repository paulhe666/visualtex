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
import { applyVisualTexIntegralSvgGlyphs } from "../math/integralSvgExportCompatibility.ts";
import {
  applyCustomSymbolArtworkToSvg,
  expandCustomSymbolsForMathMl,
  expandCustomSymbolsForSvg,
} from "../math/customSymbolRendering.ts";
import { readErrorMessage } from "../errors/readErrorMessage.ts";
import {
  assertNoUnfilledStructuralPlaceholders,
  assertResolvedMathJaxSvg,
  assertResolvedPresentationMathMl,
  normalizePackageLatexCommands,
  VISUALTEX_MATHML_MACROS,
  VISUALTEX_SVG_MACROS,
  type VisualTexMathJaxMacro,
} from "../math/latexCompatibility.ts";
import type {
  PngExportOptions,
  PngExportResult,
  SvgExportOptions,
  SvgExportResult,
} from "./exportTypes.ts";
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
      const message = readErrorMessage(error, "MathJax 无法解析当前 LaTeX 公式。");
      if (error instanceof Error && error.message.trim() === message) throw error;
      throw new Error(message, { cause: error });
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
const svgDocument = mathjax.document("", {
  InputJax: svgTexInput,
  OutputJax: svgOutput,
});
const serializedMmlVisitor = new SerializedMmlVisitor(mathMlDocument.mmlFactory);

function positiveFinite(value: number, fallback: number) {
  return Number.isFinite(value) && value > 0 ? value : fallback;
}

function nonNegativeFinite(value: number, fallback: number) {
  return Number.isFinite(value) && value >= 0 ? value : fallback;
}

function isSingleCompleteEnvironment(source: string) {
  const first = source.match(/^\\begin\s*\{([^{}]+)\}/);
  if (!first) return false;

  const environmentToken = /\\(begin|end)\s*\{([^{}]+)\}/g;
  const stack: string[] = [];
  let match: RegExpExecArray | null;
  let outerEnd = -1;

  while ((match = environmentToken.exec(source))) {
    const [, kind, name] = match;
    if (kind === "begin") {
      stack.push(name);
      continue;
    }
    if (stack.at(-1) !== name) return false;
    stack.pop();
    if (stack.length === 0) {
      outerEnd = environmentToken.lastIndex;
      break;
    }
  }

  return outerEnd >= 0 && source.slice(outerEnd).trim().length === 0;
}

function prepareLatex(latex: string) {
  const normalized = normalizeMathLiveCanonicalUprightCommands(
    normalizeExtendedIntegralLatexCommands(
      normalizePackageLatexCommands(latex.replace(/\r\n?/g, "\n")),
    ),
  ).trim();
  if (!normalized) throw new Error("Cannot export an empty formula.");
  assertNoUnfilledStructuralPlaceholders(normalized);

  const lines = normalized
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  if (lines.length <= 1) return normalized;

  // Preserve a source string that is itself one complete TeX environment.
  // A document with multiple VisualTeX formula rows may still contain an
  // inner matrix/cases environment on one row; that must not make all rows
  // collapse into a single horizontal TeX expression.
  if (isSingleCompleteEnvironment(normalized)) return normalized;

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

type SvgViewBox = {
  x: number;
  y: number;
  width: number;
  height: number;
};

type SvgRootGeometry = {
  viewBox: SvgViewBox;
  unitsPerPx: number;
  baselinePx: number | null;
  fullViewportNestedSvg: boolean;
};

function readSvgAttribute(opening: string, name: string) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return opening.match(new RegExp(`\\s${escaped}=["']([^"']*)["']`, "i"))?.[1] ?? null;
}

function readStyleDeclaration(style: string | null, name: string) {
  if (!style) return null;
  const normalizedName = name.toLowerCase();
  for (const declaration of style.split(";")) {
    const separator = declaration.indexOf(":");
    if (separator <= 0) continue;
    if (declaration.slice(0, separator).trim().toLowerCase() === normalizedName) {
      return declaration.slice(separator + 1).trim();
    }
  }
  return null;
}

function parseSvgViewBox(value: string | null) {
  if (!value) return null;
  const match = value.match(
    /^\s*([-+\d.eE]+)[\s,]+([-+\d.eE]+)[\s,]+([-+\d.eE]+)[\s,]+([-+\d.eE]+)\s*$/,
  );
  if (!match) throw new Error("Exported SVG has an invalid viewBox.");
  const values = match.slice(1).map(Number);
  if (values.some((number) => !Number.isFinite(number))) {
    throw new Error("Exported SVG has an invalid viewBox.");
  }
  const [x, y, width, height] = values;
  if (!(width > 0) || !(height > 0)) {
    throw new Error("Exported SVG has a non-positive viewBox.");
  }
  return { x, y, width, height };
}

function parseCssSvgLength(
  value: string | null,
  fontSizePx: number,
  exPx: number,
) {
  if (!value) return null;
  const match = value.trim().match(/^([-+\d.eE]+)\s*(px|ex|em)?$/i);
  if (!match) return null;
  const number = Number(match[1]);
  if (!Number.isFinite(number)) return null;
  const unit = (match[2] ?? "px").toLowerCase();
  return unit === "ex" ? number * exPx : unit === "em" ? number * fontSizePx : number;
}

function resolveSvgRootGeometry(
  svg: string,
  fontSizePx: number,
  exPx: number,
): SvgRootGeometry {
  const rootOpening = svg.match(/^<svg\b[^>]*>/i)?.[0];
  if (!rootOpening) throw new Error("MathJax did not produce an SVG root element.");

  const rootViewBox = parseSvgViewBox(readSvgAttribute(rootOpening, "viewBox"));
  if (rootViewBox) {
    return {
      viewBox: rootViewBox,
      unitsPerPx: 1000 / fontSizePx,
      baselinePx: null,
      fullViewportNestedSvg: false,
    };
  }

  const style = readSvgAttribute(rootOpening, "style");
  const widthPx = parseCssSvgLength(
    readStyleDeclaration(style, "min-width") ?? readSvgAttribute(rootOpening, "width"),
    fontSizePx,
    exPx,
  );
  const heightPx = parseCssSvgLength(
    readSvgAttribute(rootOpening, "height"),
    fontSizePx,
    exPx,
  );
  const verticalAlignPx = parseCssSvgLength(
    readStyleDeclaration(style, "vertical-align"),
    fontSizePx,
    exPx,
  ) ?? 0;
  if (!widthPx || widthPx <= 0 || !heightPx || heightPx <= 0) {
    throw new Error("Exported SVG is missing a valid root viewBox and intrinsic size.");
  }
  return {
    viewBox: { x: 0, y: 0, width: widthPx, height: heightPx },
    unitsPerPx: 1,
    baselinePx: Math.max(0, Math.min(heightPx, heightPx + verticalAlignPx)),
    fullViewportNestedSvg: true,
  };
}

function normalizeFullViewportNestedSvg(
  svg: string,
  width: number,
  height: number,
) {
  return svg.replace(
    /<svg\b([^>]*\bdata-(?:table|labels)=["'][^"']+["'][^>]*)>/gi,
    (_opening, rawAttributes: string) => {
      let attributes = rawAttributes;
      const append: string[] = [];
      if (!/\swidth=["']/i.test(attributes)) append.push(`width="${width}"`);
      if (!/\sheight=["']/i.test(attributes)) append.push(`height="${height}"`);
      if (!/\sx=["']/i.test(attributes)) append.push('x="0"');
      if (!/\sy=["']/i.test(attributes)) append.push('y="0"');
      if (!/\soverflow=["']/i.test(attributes)) append.push('overflow="visible"');
      attributes = attributes.trim();
      return `<svg${attributes ? ` ${attributes}` : ""}${append.length ? ` ${append.join(" ")}` : ""}>`;
    },
  );
}

function rewriteRootSvgOpening(
  svg: string,
  width: number,
  height: number,
  viewBox: SvgViewBox,
) {
  return svg.replace(/^<svg\b([^>]*)>/i, (_opening, rawAttributes: string) => {
    const attributes = rawAttributes
      .replace(
        /\s(?:xmlns|width|height|role|focusable|style|viewBox)=["'][^"']*["']/gi,
        "",
      )
      .trim();
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" role="img" focusable="false" viewBox="${viewBox.x} ${viewBox.y} ${viewBox.width} ${viewBox.height}"${
      attributes ? ` ${attributes}` : ""
    }>`;
  });
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
    return `<text data-c="${codePoint}" data-visualtex-output-letter-font="${escapeSvgAttribute(letterFont)}" transform="scale(1,-1)" font-size="1000px" font-family="${family}"${italic ? ' font-style="italic"' : ""}${bold ? ' font-weight="700"' : ""}>${escapeSvgText(character)}</text>`;
  });
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

  const container = svgDocument.convert(source, {
    display: options.displayMode,
    em: fontSizePx,
    ex: exPx,
    containerWidth: 100_000,
  });
  let svg = extractSvg(adaptor.outerHTML(container));
  svg = applyVisualTexIntegralSvgGlyphs(svg, options.displayMode);
  svg = applyCustomSymbolArtworkToSvg(svg);
  svg = applyVisualTexSvgFontPreferences(svg, options);
  const rootGeometry = resolveSvgRootGeometry(svg, fontSizePx, exPx);
  const viewBox = rootGeometry.viewBox;
  if (rootGeometry.fullViewportNestedSvg) {
    svg = normalizeFullViewportNestedSvg(svg, viewBox.width, viewBox.height);
  }

  const paddingUnits = paddingPx * rootGeometry.unitsPerPx;
  const padded = {
    x: viewBox.x - paddingUnits,
    y: viewBox.y - paddingUnits,
    width: viewBox.width + 2 * paddingUnits,
    height: viewBox.height + 2 * paddingUnits,
  };
  const width = Math.max(1, padded.width / rootGeometry.unitsPerPx);
  const height = Math.max(1, padded.height / rootGeometry.unitsPerPx);
  const baseline =
    rootGeometry.baselinePx === null
      ? Math.max(0, Math.min(height, -padded.y / rootGeometry.unitsPerPx))
      : Math.max(0, Math.min(height, paddingPx + rootGeometry.baselinePx));

  svg = rewriteRootSvgOpening(svg, width, height, padded).replaceAll(
    "currentColor",
    "#111111",
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

export async function svgToPng(
  svgResult: SvgExportResult,
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

  const blob = await new Promise<Blob>((resolve, reject) => {
    canvas.toBlob(
      (value) =>
        value ? resolve(value) : reject(new Error("Unable to encode PNG output.")),
      "image/png",
    );
  });
  return {
    blob,
    base64: await blobToBase64(blob),
    width,
    height,
  };
}
