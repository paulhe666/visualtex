import { mathjax } from "mathjax-full/js/mathjax.js";
import { TeX } from "mathjax-full/js/input/tex.js";
import { SVG } from "mathjax-full/js/output/svg.js";
import { liteAdaptor } from "mathjax-full/js/adaptors/liteAdaptor.js";
import { RegisterHTMLHandler } from "mathjax-full/js/handlers/html.js";
import { AllPackages } from "mathjax-full/js/input/tex/AllPackages.js";
import { normalizeMathLiveCanonicalUprightCommands } from "../editor/normalizeChineseLatex.ts";
import type {
  PngExportOptions,
  PngExportResult,
  SvgExportOptions,
  SvgExportResult,
} from "./exportTypes";

const DEFAULT_OPTIONS: SvgExportOptions = {
  displayMode: true,
  fontSizePt: 12,
  paddingPx: 8,
  background: "transparent",
};

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);
const texInput = new TeX({
  packages: AllPackages,
  formatError: (_jax: unknown, error: unknown) => {
    throw error instanceof Error ? error : new Error(String(error));
  },
});
const svgOutput = new SVG({
  fontCache: "local",
  internalSpeechTitles: false,
});
const mathDocument = mathjax.document("", {
  InputJax: texInput,
  OutputJax: svgOutput,
});

function positiveFinite(value: number, fallback: number) {
  return Number.isFinite(value) && value > 0 ? value : fallback;
}

function nonNegativeFinite(value: number, fallback: number) {
  return Number.isFinite(value) && value >= 0 ? value : fallback;
}

function prepareLatex(latex: string) {
  const normalized = normalizeMathLiveCanonicalUprightCommands(
    latex.replace(/\r\n?/g, "\n"),
  ).trim();
  if (!normalized) throw new Error("Cannot export an empty formula.");
  if (/\\begin\s*\{/.test(normalized)) return normalized;

  const lines = normalized
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  if (lines.length <= 1) return normalized;
  // `aligned` uses a right/left pair around every alignment marker. Without
  // an explicit marker MathJax right-aligns rows of different widths. Keep
  // the whole formula as one image, but anchor every row on its left edge.
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

export function latexToSvg(
  latex: string,
  options: SvgExportOptions = DEFAULT_OPTIONS,
): SvgExportResult {
  const source = prepareLatex(latex);
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
    .replaceAll("currentColor", "#111111");

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
  if (options.background === "white") {
    context.fillStyle = "#ffffff";
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
