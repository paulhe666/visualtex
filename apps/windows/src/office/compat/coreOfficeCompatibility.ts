import type { EngineEditorSessionOwner } from "../../../../../engine/web/host/editor-session";
import type { LayoutOptions, LayoutSnapshot } from "../../../../../engine/web/host/layout";
import type {
  FormulaChineseFont,
  FormulaLetterFont,
} from "../../editor/formulaFontPreferences";
import type {
  OfficeExportResult,
  OfficeObjectMode,
} from "../shared/sessionClient";

const CSS_PX_PER_PT = 96 / 72;
const MTEF_UNITS_PER_CSS_PX = 24;
const MTEF_MAX_RULER_STOP = 0xffff;
const CORE_GRID_ATTRIBUTE = "data-visualtex-core-grid-id";
const MTEF_TABS_ATTRIBUTE = "data-mtef-tabs";
const MTEF_RULER_ATTRIBUTE = "data-visualtex-mtef-ruler-stops";

export interface CoreOfficeExportOwner {
  layout(
    options: LayoutOptions,
    includeSvg: boolean,
  ): { sceneToken: number; snapshot: LayoutSnapshot; svg: string | null };
  presentationMathMl(displayMode: boolean): {
    mathMl: string;
    alignedGridIds: number[];
  };
}

export interface PreservedOfficeGeometry {
  width: number;
  height: number;
  baseline?: number;
}

export interface CoreOfficeExportRequest {
  owner: CoreOfficeExportOwner;
  layout: LayoutOptions;
  objectMode: OfficeObjectMode;
  displayMode: "inline" | "block";
  fontSizePt: number;
  formulaLetterFont: FormulaLetterFont;
  formulaChineseFont: FormulaChineseFont;
  /** Required only by the MathML-only Word OMML path. */
  preservedGeometry?: PreservedOfficeGeometry;
}

export type CoreOfficeRasterizer = (
  svg: Pick<OfficeExportResult, "svgBase64" | "width" | "height">,
) => Promise<string>;

function positiveFinite(value: number, name: string) {
  if (!Number.isFinite(value) || value <= 0) {
    throw new Error(`Invalid ${name} for Office export.`);
  }
  return value;
}

function nonNegativeFinite(value: number, name: string) {
  if (!Number.isFinite(value) || value < 0) {
    throw new Error(`Invalid ${name} for Office export.`);
  }
  return value;
}

function encodeUtf8Base64(value: string) {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(binary);
}

function stripPrivateGridAttributes(mathMl: string, retainMtefTabs: boolean) {
  let output = mathMl.replace(
    new RegExp(`\\s${CORE_GRID_ATTRIBUTE}="\\d+"`, "gu"),
    "",
  );
  if (!retainMtefTabs) {
    output = output.replace(
      new RegExp(`\\s${MTEF_TABS_ATTRIBUTE}="true"`, "gu"),
      "",
    );
  }
  return output;
}

function rulerStopsForGrid(snapshot: LayoutSnapshot, gridId: number) {
  const columns = (snapshot.gridColumns ?? [])
    .filter((column) => column.gridId === gridId)
    .slice()
    .sort((left, right) => left.index - right.index);
  if (columns.length < 2 || columns.length % 2 !== 0) {
    throw new Error("Core aligned-grid geometry is incomplete for MathType export.");
  }
  if (columns.some((column, index) => column.index !== index)) {
    throw new Error("Core aligned-grid columns are not contiguous.");
  }
  const origin = columns[0].x;
  const stops: number[] = [];
  let previous = 0;
  for (let index = 1; index < columns.length; index += 2) {
    const xPx = (columns[index].x - origin) * snapshot.fontSizePx;
    const stop = Math.round(xPx * MTEF_UNITS_PER_CSS_PX);
    if (
      !Number.isSafeInteger(stop) ||
      stop <= previous ||
      stop <= 0 ||
      stop > MTEF_MAX_RULER_STOP
    ) {
      throw new Error("Core aligned-grid ruler geometry is outside MathType bounds.");
    }
    stops.push(stop);
    previous = stop;
  }
  if (stops.length !== columns.length / 2) {
    throw new Error("Core aligned-grid ruler count does not match its pair count.");
  }
  return stops;
}

export function annotateCoreMathTypeAlignmentGeometry(
  mathMl: string,
  alignedGridIds: readonly number[],
  snapshot: LayoutSnapshot,
) {
  let output = mathMl;
  for (const gridId of alignedGridIds) {
    const marker = `${CORE_GRID_ATTRIBUTE}="${gridId}"`;
    const tablePattern = new RegExp(`<mtable\\b([^>]*\\s${marker.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&")}[^>]*)>`, "u");
    const match = tablePattern.exec(output);
    if (!match) {
      throw new Error("Core MathML/grid geometry identities disagree.");
    }
    const stops = rulerStopsForGrid(snapshot, gridId);
    let attributes = match[1]
      .replace(new RegExp(`\\s${CORE_GRID_ATTRIBUTE}="${gridId}"`, "u"), "")
      .replace(new RegExp(`\\s${MTEF_RULER_ATTRIBUTE}="[^"]*"`, "u"), "");
    if (!new RegExp(`\\s${MTEF_TABS_ATTRIBUTE}="true"`, "u").test(attributes)) {
      attributes += ` ${MTEF_TABS_ATTRIBUTE}="true"`;
    }
    attributes += ` ${MTEF_RULER_ATTRIBUTE}="${stops.join(",")}"`;
    output = `${output.slice(0, match.index)}<mtable${attributes}>${output.slice(match.index + match[0].length)}`;
  }
  return stripPrivateGridAttributes(output, true);
}

function parseRootSvg(svg: string) {
  const match = /^<svg\b([^>]*)>/u.exec(svg.trim());
  if (!match) throw new Error("Core Office export did not produce an SVG root.");
  const attribute = (name: string) => {
    const found = new RegExp(`\\b${name}="([^"]+)"`, "u").exec(match[1]);
    return found?.[1] ?? null;
  };
  const width = Number(attribute("width"));
  const height = Number(attribute("height"));
  const viewBox = (attribute("viewBox") ?? "")
    .trim()
    .split(/[\s,]+/u)
    .map(Number);
  if (
    !Number.isFinite(width) || width <= 0 ||
    !Number.isFinite(height) || height <= 0 ||
    viewBox.length !== 4 || !viewBox.every(Number.isFinite) ||
    viewBox[2] <= 0 || viewBox[3] <= 0
  ) {
    throw new Error("Core Office SVG geometry is invalid.");
  }
  return { match, width, height, viewBox };
}

function addMathTypeVerticalPadding(
  svg: string,
  snapshot: LayoutSnapshot,
  paddingPx: number,
) {
  const { match, viewBox } = parseRootSvg(svg);
  const paddingEm = paddingPx / snapshot.fontSizePx;
  const nextViewBox = [viewBox[0], viewBox[1] - paddingEm, viewBox[2], viewBox[3] + 2 * paddingEm];
  const nextHeight = snapshot.height + 2 * paddingPx;
  const opening = match[0]
    .replace(/\bheight="[^"]+"/u, `height="${nextHeight}"`)
    .replace(/\bviewBox="[^"]+"/u, `viewBox="${nextViewBox.join(" ")}"`);
  const trimmed = svg.trim();
  return {
    svg: `${opening}${trimmed.slice(match[0].length)}`,
    width: snapshot.width,
    height: nextHeight,
    baseline: snapshot.baseline + paddingPx,
  };
}

function ensureOfficeVectorSubset(svg: string) {
  const forbidden = [
    /<script\b/iu,
    /<foreignObject\b/iu,
    /<image\b/iu,
    /<text\b/iu,
    /<mask\b/iu,
    /<clipPath\b/iu,
    /\bmask\s*=/iu,
    /\bclip-path\s*=/iu,
    /\b(?:xlink:)?href\s*=\s*["'](?!#)/iu,
    /\burl\s*\(/iu,
  ];
  if (forbidden.some((pattern) => pattern.test(svg))) {
    throw new Error(
      "Core SVG requires vector features that the existing Office EMF consumer does not support.",
    );
  }
}

function addTransparentHitTarget(svg: string) {
  const { match, viewBox } = parseRootSvg(svg);
  const [x, y, width, height] = viewBox;
  const hitTarget = `<rect x="${x}" y="${y}" width="${width}" height="${height}" fill="#000000" fill-opacity="0.001"/>`;
  const trimmed = svg.trim();
  return `${trimmed.slice(0, match[0].length)}${hitTarget}${trimmed.slice(match[0].length)}`;
}

function outputMathMl(
  request: CoreOfficeExportRequest,
  snapshot?: LayoutSnapshot,
) {
  const isMathType = request.objectMode === "mathTypeOle";
  const presentation = request.owner.presentationMathMl(
    isMathType || request.displayMode === "block",
  );
  if (isMathType) {
    if (!snapshot && presentation.alignedGridIds.length) {
      throw new Error("MathType aligned export requires the matching core layout snapshot.");
    }
    return snapshot
      ? annotateCoreMathTypeAlignmentGeometry(
          presentation.mathMl,
          presentation.alignedGridIds,
          snapshot,
        )
      : stripPrivateGridAttributes(presentation.mathMl, true);
  }
  return stripPrivateGridAttributes(presentation.mathMl, false);
}

export function generateCoreOfficeExport(
  request: CoreOfficeExportRequest,
): OfficeExportResult {
  const fontSizePt = positiveFinite(request.fontSizePt, "Office semantic font size");
  const fontSizePx = fontSizePt * CSS_PX_PER_PT;
  if (request.objectMode === "wordOmml") {
    const geometry = request.preservedGeometry;
    const width = positiveFinite(geometry?.width ?? 240, "preserved OMML width");
    const height = positiveFinite(geometry?.height ?? 80, "preserved OMML height");
    const baseline = nonNegativeFinite(geometry?.baseline ?? height * 0.75, "preserved OMML baseline");
    if (baseline > height) throw new Error("Preserved OMML baseline exceeds its natural height.");
    return {
      svg: "",
      svgBase64: "",
      mathMl: outputMathMl(request),
      width,
      height,
      baseline,
      formulaLetterFont: request.formulaLetterFont,
      formulaChineseFont: request.formulaChineseFont,
    };
  }

  const isMathType = request.objectMode === "mathTypeOle";
  const uniformPaddingPx = isMathType
    ? 0
    : request.displayMode === "inline" ? 1 : 10;
  const options: LayoutOptions = {
    ...request.layout,
    style: isMathType || request.displayMode === "block" ? "display" : "text",
    fontSizePx,
    paddingEm: uniformPaddingPx / fontSizePx,
  };
  const rendered = request.owner.layout(options, true);
  if (!rendered.svg) throw new Error("Core Office export did not return SVG geometry.");
  let svg = rendered.svg;
  let width = rendered.snapshot.width;
  let height = rendered.snapshot.height;
  let baseline = rendered.snapshot.baseline;
  if (isMathType) {
    const padded = addMathTypeVerticalPadding(
      svg,
      rendered.snapshot,
      fontSizePt * 0.5,
    );
    svg = padded.svg;
    width = padded.width;
    height = padded.height;
    baseline = padded.baseline;
  }
  ensureOfficeVectorSubset(svg);
  svg = addTransparentHitTarget(svg);
  ensureOfficeVectorSubset(svg);
  const mathMl = outputMathMl(request, rendered.snapshot);
  return {
    svg,
    svgBase64: encodeUtf8Base64(svg),
    mathMl,
    width,
    height,
    baseline,
    formulaLetterFont: request.formulaLetterFont,
    formulaChineseFont: request.formulaChineseFont,
  };
}

export async function rasterizeCoreOfficeExport(
  base: OfficeExportResult,
  rasterize: CoreOfficeRasterizer,
): Promise<OfficeExportResult> {
  if (!base.svg || !base.svgBase64) return base;
  const pngBase64 = await rasterize({
    svgBase64: base.svgBase64,
    width: base.width,
    height: base.height,
  });
  if (!pngBase64.trim()) {
    throw new Error("Core Office PNG rasterizer returned an empty image.");
  }
  return { ...base, pngBase64 };
}

// Compile-time proof that the real owner satisfies the deliberately narrow
// P11 adapter surface without importing Session state or Office I/O.
const _ownerCompatibility: CoreOfficeExportOwner | null =
  null as EngineEditorSessionOwner | null;
void _ownerCompatibility;
