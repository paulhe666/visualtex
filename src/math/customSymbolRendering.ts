import type { MathfieldElement } from "mathlive";
import type {
  CustomSymbolDefinition,
  CustomSymbolVectorShape,
  CustomSymbolVectorTransform,
} from "./customSymbolTypes.ts";
import {
  applyCustomSymbolMacrosToMathfield,
  customSymbolCssClass,
  customSymbolMathLiveMacroDefinition,
  getAppliedCustomSymbolCommandsForMathfield,
  customSymbolSvgMacro,
  customSymbolSvgMarkerClass,
  getActiveCustomSymbols,
  getCustomSymbolRevision,
} from "./customSymbolRegistry.ts";

const globalStyleId = "visualtex-custom-symbol-runtime-style";
const shadowStyleId = "visualtex-custom-symbol-runtime-shadow-style";
const shadowStyleSignatures = new WeakMap<MathfieldElement, string>();
let cachedGlobalStyleRevision = -1;
let cachedGlobalStyleCss = "";

function number(value: number) {
  const normalized = Math.abs(value) < 0.000001 ? 0 : value;
  return Number(normalized.toFixed(5)).toString();
}

function escapeAttribute(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function userTransformParts(transform?: CustomSymbolVectorTransform) {
  if (!transform) return [];
  const tx = transform.translateX ?? 0;
  const ty = transform.translateY ?? 0;
  const sx = transform.scaleX ?? 1;
  const sy = transform.scaleY ?? 1;
  const angle = transform.rotateDeg ?? 0;
  const ox = transform.originX ?? 0;
  const oy = transform.originY ?? 0;
  const parts: string[] = [];
  if (tx || ty) parts.push(`translate(${number(tx)} ${number(ty)})`);
  if (ox || oy) parts.push(`translate(${number(ox)} ${number(oy)})`);
  if (angle) parts.push(`rotate(${number(angle)})`);
  if (sx !== 1 || sy !== 1) parts.push(`scale(${number(sx)} ${number(sy)})`);
  if (ox || oy) parts.push(`translate(${number(-ox)} ${number(-oy)})`);
  return parts;
}

function userTransformMarkup(transform?: CustomSymbolVectorTransform) {
  const parts = userTransformParts(transform);
  return parts.length ? ` transform="${parts.join(" ")}"` : "";
}

function matrixTransformMarkup(transform?: CustomSymbolVectorTransform) {
  return transform?.matrix
    ? ` transform="matrix(${transform.matrix.map(number).join(" ")})"`
    : "";
}

function transformMarkup(transform?: CustomSymbolVectorTransform) {
  const parts = userTransformParts(transform);
  if (transform?.matrix) {
    parts.push(`matrix(${transform.matrix.map(number).join(" ")})`);
  }
  return parts.length ? ` transform="${parts.join(" ")}"` : "";
}

function paintAttributes(shape: CustomSymbolVectorShape, paint: string) {
  const defaultFill = shape.kind !== "line";
  const fill = shape.fill ?? defaultFill;
  const strokeWidth = Math.max(0, shape.strokeWidth ?? (shape.kind === "line" ? 50 : 0));
  const parts = [fill ? `fill="${paint}"` : 'fill="none"'];
  if (strokeWidth > 0) {
    parts.push(`stroke="${paint}"`);
    parts.push(`stroke-width="${number(strokeWidth)}"`);
    if (shape.lineCap) parts.push(`stroke-linecap="${shape.lineCap}"`);
    if (shape.lineJoin) parts.push(`stroke-linejoin="${shape.lineJoin}"`);
  } else {
    parts.push('stroke="none"');
  }
  return parts.join(" ");
}

function geometryMarkup(
  shape: CustomSymbolVectorShape,
  paintColor: string,
  transform: string,
) {
  const paint = paintAttributes(shape, paintColor);
  switch (shape.kind) {
    case "path":
      return `<path d="${escapeAttribute(shape.d)}" ${paint}${transform}></path>`;
    case "circle":
      return `<circle cx="${number(shape.cx)}" cy="${number(shape.cy)}" r="${number(shape.r)}" ${paint}${transform}></circle>`;
    case "line":
      return `<line x1="${number(shape.x1)}" y1="${number(shape.y1)}" x2="${number(shape.x2)}" y2="${number(shape.y2)}" ${paint}${transform}></line>`;
    case "rect":
      return `<rect x="${number(shape.x)}" y="${number(shape.y)}" width="${number(shape.width)}" height="${number(shape.height)}"${shape.rx === undefined ? "" : ` rx="${number(shape.rx)}"`}${shape.ry === undefined ? "" : ` ry="${number(shape.ry)}"`} ${paint}${transform}></rect>`;
    case "ellipse":
      return `<ellipse cx="${number(shape.cx)}" cy="${number(shape.cy)}" rx="${number(shape.rx)}" ry="${number(shape.ry)}" ${paint}${transform}></ellipse>`;
    case "polygon":
      return `<polygon points="${shape.points.map(([x, y]) => `${number(x)},${number(y)}`).join(" ")}" ${paint}${transform}></polygon>`;
  }
}

function shapeMarkup(shape: CustomSymbolVectorShape, paintColor: string) {
  if (!shape.clipRect) {
    return geometryMarkup(shape, paintColor, transformMarkup(shape.transform));
  }
  const geometry = geometryMarkup(
    shape,
    paintColor,
    matrixTransformMarkup(shape.transform),
  );
  const { x, y, width, height } = shape.clipRect;
  const clipped =
    `<svg x="${number(x)}" y="${number(y)}" width="${number(width)}" height="${number(height)}" ` +
    `viewBox="${number(x)} ${number(y)} ${number(width)} ${number(height)}" ` +
    `preserveAspectRatio="none" overflow="hidden">${geometry}</svg>`;
  const userTransform = userTransformMarkup(shape.transform);
  return userTransform ? `<g${userTransform}>${clipped}</g>` : clipped;
}

function artworkHasErase(symbol: CustomSymbolDefinition) {
  return symbol.artwork.shapes.some((shape) => shape.operation === "erase");
}

function artworkMaskMarkup(
  symbol: CustomSymbolDefinition,
  maskId: string,
  width: number,
  height: number,
) {
  const paintShapes = symbol.artwork.shapes.filter((shape) => shape.operation !== "erase");
  const eraseShapes = symbol.artwork.shapes.filter((shape) => shape.operation === "erase");
  return (
    `<defs><mask id="${maskId}" maskUnits="userSpaceOnUse" x="0" y="0" width="${number(width)}" height="${number(height)}" style="mask-type:luminance">` +
    `<rect x="0" y="0" width="${number(width)}" height="${number(height)}" fill="black"></rect>` +
    paintShapes.map((shape) => shapeMarkup(shape, "white")).join("") +
    eraseShapes.map((shape) => shapeMarkup(shape, "black")).join("") +
    `</mask></defs>`
  );
}

export function customSymbolArtworkSvg(
  symbol: CustomSymbolDefinition,
  monochromeMask = true,
) {
  const width = Math.max(1, symbol.metrics.widthEm * 1000);
  const height = Math.max(1, (symbol.metrics.ascentEm + symbol.metrics.descentEm) * 1000);
  const paint = monochromeMask ? "black" : "inherit";
  if (!artworkHasErase(symbol)) {
    return (
      `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${number(width)} ${number(height)}">` +
      symbol.artwork.shapes.map((shape) => shapeMarkup(shape, paint)).join("") +
      "</svg>"
    );
  }
  const maskId = "visualtex-custom-symbol-erase";
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${number(width)} ${number(height)}">` +
    artworkMaskMarkup(symbol, maskId, width, height) +
    `<rect x="0" y="0" width="${number(width)}" height="${number(height)}" fill="${paint}" mask="url(#${maskId})"></rect>` +
    "</svg>"
  );
}

function maskLayerSvg(
  symbol: CustomSymbolDefinition,
  operation: "paint" | "erase",
) {
  const width = Math.max(1, symbol.metrics.widthEm * 1000);
  const height = Math.max(1, (symbol.metrics.ascentEm + symbol.metrics.descentEm) * 1000);
  const shapes = symbol.artwork.shapes.filter((shape) =>
    operation === "erase" ? shape.operation === "erase" : shape.operation !== "erase",
  );
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${number(width)} ${number(height)}">` +
    shapes.map((shape) => shapeMarkup(shape, "black")).join("") +
    "</svg>"
  );
}

function maskLayerDataUrl(
  symbol: CustomSymbolDefinition,
  operation: "paint" | "erase",
) {
  return `url("data:image/svg+xml,${encodeURIComponent(maskLayerSvg(symbol, operation))}")`;
}

export function customSymbolRuntimeCss(symbols = getActiveCustomSymbols()) {
  return symbols
    .map((symbol) => {
      const className = customSymbolCssClass(symbol);
      const paintMask = maskLayerDataUrl(symbol, "paint");
      const hasErase = artworkHasErase(symbol);
      const eraseMask = hasErase ? maskLayerDataUrl(symbol, "erase") : "";
      const maskImages = hasErase ? `${paintMask}, ${eraseMask}` : paintMask;
      const maskPosition = hasErase ? "0 0, 0 0" : "0 0";
      const maskRepeat = hasErase ? "no-repeat, no-repeat" : "no-repeat";
      const maskSize = hasErase ? "100% 100%, 100% 100%" : "100% 100%";
      const composite = hasErase
        ? `
  -webkit-mask-composite: source-out;
  mask-composite: subtract;`
        : "";
      return `
.${className} .ML__inner,
.${className} .ML__inner * {
  opacity: 1 !important;
}
.${className} .ML__rule {
  opacity: 1 !important;
  border-top-color: transparent !important;
  border-right-color: transparent !important;
  background-color: currentColor;
  background-clip: border-box;
  -webkit-mask-image: ${maskImages};
  mask-image: ${maskImages};
  -webkit-mask-position: ${maskPosition};
  mask-position: ${maskPosition};
  -webkit-mask-repeat: ${maskRepeat};
  mask-repeat: ${maskRepeat};
  -webkit-mask-size: ${maskSize};
  mask-size: ${maskSize};${composite}
}`;
    })
    .join("\n");
}

function installStyle(root: Document | ShadowRoot, id: string, css: string) {
  let style = root.getElementById(id) as HTMLStyleElement | null;
  if (!style) {
    style = document.createElement("style");
    style.id = id;
    if (root instanceof Document) root.head.append(style);
    else root.append(style);
  }
  if (style.textContent !== css) style.textContent = css;
}

export function customSymbolsUsedInSource(source: string) {
  if (!source.includes("\\")) return [];
  const activeSymbols = getActiveCustomSymbols();
  const byCommand = new Map(
    activeSymbols.map((symbol) => [symbol.command, symbol]),
  );
  const usedCommands = new Set<string>();
  let index = 0;
  while (index < source.length) {
    const slash = source.indexOf("\\", index);
    if (slash < 0) break;
    if (source[slash + 1] === "\\") {
      index = slash + 2;
      continue;
    }
    let end = slash + 1;
    while (/[A-Za-z]/.test(source[end] ?? "")) end += 1;
    if (end > slash + 1) {
      const command = source.slice(slash + 1, end);
      if (byCommand.has(command)) usedCommands.add(command);
    }
    index = Math.max(slash + 1, end);
  }
  return activeSymbols.filter((symbol) => usedCommands.has(symbol.command));
}

export function installCustomSymbolGlobalStyle() {
  if (typeof document === "undefined") return;
  const revision = getCustomSymbolRevision();
  if (revision !== cachedGlobalStyleRevision) {
    cachedGlobalStyleRevision = revision;
    cachedGlobalStyleCss = customSymbolRuntimeCss();
  }
  installStyle(document, globalStyleId, cachedGlobalStyleCss);
}

export function installCustomSymbolShadowStyle(field: MathfieldElement) {
  const root = field.shadowRoot;
  if (!root) return;
  const revision = getCustomSymbolRevision();
  const usedSymbols = customSymbolsUsedInSource(field.value);
  const signature = `${revision}\u0000${usedSymbols.map((symbol) => symbol.id).join("\u0000")}`;
  if (shadowStyleSignatures.get(field) === signature) return;
  installStyle(root, shadowStyleId, customSymbolRuntimeCss(usedSymbols));
  shadowStyleSignatures.set(field, signature);
}

function sourceContainsControlWord(source: string, commands: ReadonlySet<string>) {
  if (!commands.size || !source.includes("\\")) return false;
  let index = 0;
  while (index < source.length) {
    const slash = source.indexOf("\\", index);
    if (slash < 0) return false;
    if (source[slash + 1] === "\\") {
      index = slash + 2;
      continue;
    }
    let end = slash + 1;
    while (/[A-Za-z]/.test(source[end] ?? "")) end += 1;
    if (end > slash + 1 && commands.has(source.slice(slash + 1, end))) {
      return true;
    }
    index = Math.max(slash + 1, end);
  }
  return false;
}

function selectionSourceAnchor(field: MathfieldElement, offset: number) {
  const safeOffset = Math.max(0, Math.min(field.lastOffset, offset));
  if (safeOffset === 0) return "";
  if (safeOffset === field.lastOffset) return field.value;
  return field.getValue(0, safeOffset, "latex");
}

function modelOffsetForSourceAnchor(
  field: MathfieldElement,
  source: string,
  anchor: string,
) {
  if (!anchor) return 0;
  if (anchor === source) return field.lastOffset;
  for (let offset = 1; offset < field.lastOffset; offset += 1) {
    if (field.getValue(0, offset, "latex") === anchor) return offset;
  }
  return anchor.length <= source.length / 2 ? 0 : field.lastOffset;
}

export function refreshCustomSymbolMathfield(field: MathfieldElement) {
  const latex = field.value;
  const previousCommands = getAppliedCustomSymbolCommandsForMathfield(field);
  const currentCommands = new Set(
    getActiveCustomSymbols().map((symbol) => symbol.command),
  );
  const needsReparse =
    sourceContainsControlWord(latex, previousCommands) ||
    sourceContainsControlWord(latex, currentCommands);
  const selectionAnchors = needsReparse
    ? {
        ranges: field.selection.ranges.map(([start, end]) => [
          selectionSourceAnchor(field, start),
          selectionSourceAnchor(field, end),
        ] as const),
        direction: field.selection.direction,
      }
    : null;

  applyCustomSymbolMacrosToMathfield(field);
  installCustomSymbolShadowStyle(field);
  if (!latex || !needsReparse || !selectionAnchors) return;

  field.setValue(`${latex} `, {
    mode: "math",
    format: "latex",
    insertionMode: "replaceAll",
    selectionMode: "after",
    silenceNotifications: true,
  });
  field.resetUndo();
  const reparsedSource = field.value;
  field.selection = {
    ranges: selectionAnchors.ranges.map(([startAnchor, endAnchor]) => [
      modelOffsetForSourceAnchor(field, reparsedSource, startAnchor),
      modelOffsetForSourceAnchor(field, reparsedSource, endAnchor),
    ]),
    direction: selectionAnchors.direction,
  };
}

function replaceRegisteredCommands(
  source: string,
  replacement: (symbol: CustomSymbolDefinition) => string | null,
) {
  const symbols = new Map(
    getActiveCustomSymbols().map((symbol) => [symbol.command, symbol]),
  );
  let output = "";
  let index = 0;
  while (index < source.length) {
    if (source[index] !== "\\") {
      output += source[index];
      index += 1;
      continue;
    }
    if (source[index + 1] === "\\") {
      output += "\\\\";
      index += 2;
      continue;
    }
    let end = index + 1;
    while (/[A-Za-z]/.test(source[end] ?? "")) end += 1;
    if (end === index + 1) {
      output += source[index];
      index += 1;
      continue;
    }
    const command = source.slice(index + 1, end);
    const symbol = symbols.get(command);
    if (!symbol) {
      output += source.slice(index, end);
      index = end;
      continue;
    }
    const next = replacement(symbol);
    output += next ?? source.slice(index, end);
    index = end;
  }
  return output;
}

export function containsCustomSymbolCommand(source: string) {
  let found = false;
  replaceRegisteredCommands(source, () => {
    found = true;
    return null;
  });
  return found;
}

export function expandCustomSymbolsForMathLiveMarkup(source: string) {
  return replaceRegisteredCommands(
    source,
    (symbol) => customSymbolMathLiveMacroDefinition(symbol).def,
  );
}

export function expandCustomSymbolsForSvg(source: string) {
  return replaceRegisteredCommands(source, customSymbolSvgMacro);
}

export function expandCustomSymbolsForMathMl(source: string) {
  return replaceRegisteredCommands(source, (symbol) => symbol.ommlFallback || null);
}

function markerPattern(className: string) {
  return new RegExp(
    `<g\\b([^>]*\\bclass=["'][^"']*\\b${className}\\b[^"']*["'][^>]*)>` +
      `<g\\b[^>]*data-mml-node=["']mphantom["'][^>]*><\\/g>` +
      `<\\/g>`,
    "g",
  );
}

function exportedArtworkGroup(symbol: CustomSymbolDefinition, occurrence = 0) {
  const baseline = number(symbol.metrics.ascentEm * 1000);
  const width = Math.max(1, symbol.metrics.widthEm * 1000);
  const height = Math.max(1, (symbol.metrics.ascentEm + symbol.metrics.descentEm) * 1000);
  let artwork: string;
  if (artworkHasErase(symbol)) {
    const safeId = symbol.id.replace(/[^A-Za-z0-9_-]/g, "-");
    const maskId = `visualtex-custom-symbol-erase-${safeId}-${occurrence}`;
    artwork =
      artworkMaskMarkup(symbol, maskId, width, height) +
      `<rect x="0" y="0" width="${number(width)}" height="${number(height)}" fill="inherit" mask="url(#${maskId})"></rect>`;
  } else {
    artwork = symbol.artwork.shapes
      .map((shape) => shapeMarkup(shape, "inherit"))
      .join("");
  }
  return (
    `<g data-visualtex-custom-symbol="${escapeAttribute(symbol.id)}" ` +
    `transform="translate(0 ${baseline}) scale(1 -1)">` +
    artwork +
    "</g>"
  );
}

export function applyCustomSymbolArtworkToSvg(svg: string) {
  let output = svg;
  for (const symbol of getActiveCustomSymbols()) {
    const marker = customSymbolSvgMarkerClass(symbol);
    let occurrence = 0;
    output = output.replace(
      markerPattern(marker),
      (_whole, attributes: string) =>
        `<g${attributes}>${exportedArtworkGroup(symbol, occurrence++)}</g>`,
    );
  }
  return output;
}

installCustomSymbolGlobalStyle();
