import { compileCustomSymbolDesignerLayer } from "./compiler.ts";
import type { CustomSymbolDesignerDocument, CustomSymbolDesignerGlyphLayer } from "./designerTypes.ts";
import { isLargeOperatorMaterial, textStyleVariant } from "./glyphAsset.ts";
import { apply, lengthScale, multiply, transformMatrix, type Affine } from "./matrix.ts";
import { normalizePathData, type PathSegment } from "./pathData.ts";
import { shapePaint, shapePathData } from "./shapes.ts";
import { isLatexMaterial, isUprightLayer, layerTexBox, type CustomSymbolBox } from "./symbolBox.ts";
import type { CustomSymbolMathRole, CustomSymbolMetrics, CustomSymbolVectorShape } from "./symbolTypes.ts";

// The designed symbol as a LaTeX command built from LaTeX's own picture
// environment: a box the size of the registered symbol (its base character's
// TeX box), each LaTeX material placed at its reference point with \put and
// turned with graphicx (any matrix is a rotation, a scaling and a rotation),
// and shapes and system-font outlines as pict2e paths. Lengths are in em, so
// the symbol follows the font size, and \text makes it follow script sizes.
//
// An operator takes its limits as TeX gives them to its base character: the
// box includes the base's italic correction, which a \nolimits subscript
// gives back and \limits splits between the limits (tex.web make_op); the
// command reads its own ^ and _ to place them so. In text and script styles
// it uses the text size of a display operator (\textstyle\int), as TeX's
// own large operators do.

export interface CustomSymbolTexExport {
  /** The \newcommand and the packages it needs, ready to paste into a preamble. */
  code: string;
  /** What the code cannot reproduce (eraser strokes, hollow glyphs, ...). */
  omissions: string[];
}

/** The text size of a display-operator material: its box and italic correction, as the engine measured them. */
export interface TextStyleMaterial {
  metrics: CustomSymbolMetrics;
  italicEm: number;
}

export interface CustomSymbolTexOptions {
  /** the text-size version of display operators (\textstyle\int for \displaystyle\int), by display source */
  textStyle?: (displaySource: string) => TextStyleMaterial | undefined;
  /** the math font the designer measured with: Computer Modern, or Times (newtxmath) */
  mathFont?: "cm" | "ntx";
}

const UNITS = 1000;

const num = (value: number, digits = 4) => {
  const scale = 10 ** digits;
  const rounded = Math.round(value * scale) / scale;
  return (Object.is(rounded, -0) ? 0 : rounded).toString();
};
const em = (units: number) => num(units / UNITS, 5);

const ROLE_WRAPPERS: Record<CustomSymbolMathRole, string> = {
  ordinary: "\\mathord",
  binary: "\\mathbin",
  relation: "\\mathrel",
  operator: "\\mathop",
  open: "\\mathopen",
  close: "\\mathclose",
  punctuation: "\\mathpunct",
};

/** A LaTeX material set at its reference point; `matrix` maps the material's box (top left at 0, y down) into the symbol. */
interface NodeItem {
  kind: "node";
  source: string;
  /** the material's height above its baseline (units) */
  ascent: number;
  matrix: Affine;
  /** a crop in the material's own coordinates */
  clip?: { x: number; y: number; width: number; height: number } | null;
}

/** An outline with every transform applied (symbol coordinates). */
interface PathItem {
  kind: "path";
  segments: PathSegment[];
  fill: boolean;
  strokeWidth: number;
  lineCap?: "butt" | "round" | "square";
  lineJoin?: "miter" | "round" | "bevel";
  /** an axis-aligned crop in symbol coordinates */
  clip?: { left: number; top: number; right: number; bottom: number };
}

type Item = NodeItem | PathItem;

const isIdentityLinear = (m: Affine) =>
  Math.abs(m[0] - 1) < 1e-6 && Math.abs(m[1]) < 1e-6 && Math.abs(m[2]) < 1e-6 && Math.abs(m[3] - 1) < 1e-6;

/**
 * A 2x2 matrix [[p, q], [r, s]] (y up) as rotate(phi) * scale(sx, sy) *
 * rotate(theta), angles in degrees counterclockwise (graphicx's \rotatebox).
 */
export function decomposeLinear(p: number, q: number, r: number, s: number) {
  const e = (p + s) / 2;
  const f = (p - s) / 2;
  const g = (r + q) / 2;
  const h = (r - q) / 2;
  const big = Math.hypot(e, h);
  const small = Math.hypot(f, g);
  const a1 = Math.atan2(g, f);
  const a2 = Math.atan2(h, e);
  return {
    phi: ((a2 + a1) / 2) * (180 / Math.PI),
    sx: big + small,
    sy: big - small,
    theta: ((a2 - a1) / 2) * (180 / Math.PI),
  };
}

function layerItems(document: CustomSymbolDesignerDocument, omissions: string[]): Item[] {
  const items: Item[] = [];
  for (const layer of document.layers) {
    if (!layer.visible) continue;
    if (layer.kind === "glyph" && isLatexMaterial(layer)) {
      const outer = transformMatrix(layer.transform, false);
      const node = (dx: number, dy: number): NodeItem => ({
        kind: "node",
        source: layer.asset.sourceLatex,
        ascent: layer.asset.metrics.ascentEm * UNITS,
        matrix: multiply([1, 0, 0, 1, dx, dy], outer),
        clip: layer.clipRect,
      });
      // the extrusion's copies are shifted in the symbol's coordinates, as when drawn
      const perspective = layer.effects?.perspective;
      if (perspective?.enabled && perspective.depth > 0) {
        const steps = Math.max(1, Math.min(24, Math.round(perspective.steps)));
        const radians = (perspective.angleDeg * Math.PI) / 180;
        for (let step = steps; step >= 1; step -= 1) {
          const ratio = step / steps;
          items.push(node(Math.cos(radians) * perspective.depth * ratio, Math.sin(radians) * perspective.depth * ratio));
        }
      }
      items.push(node(0, 0));
      if (layer.effects?.outline?.enabled) omissions.push(`${layer.name}: hollow outline (drawn solid)`);
      continue;
    }
    // shapes and system-font outlines: the outlines as the symbol stores them
    const label = layer.kind === "glyph" ? layer.asset.sourceLatex : layer.name;
    for (const shape of compileCustomSymbolDesignerLayer(layer)) {
      const item = pathItem(shape, label, omissions);
      if (item) items.push(item);
    }
  }
  return items;
}

function pathItem(shape: CustomSymbolVectorShape, label: string, omissions: string[]): PathItem | null {
  if (shape.operation === "erase") {
    omissions.push(`${label}: eraser`);
    return null;
  }
  const data = shapePathData(shape);
  if (data === null) {
    omissions.push(`${label}: text without an outline`);
    return null;
  }
  const { fill, strokeWidth } = shapePaint(shape);
  if (!fill && strokeWidth <= 0) return null;
  const outer = transformMatrix(shape.transform, false);
  const total = shape.transform?.matrix ? multiply(outer, [...shape.transform.matrix] as Affine) : outer;
  const segments = normalizePathData(data).map((segment): PathSegment => {
    if (segment.op === "Z") return segment;
    const end = apply(total, segment.x, segment.y);
    if (segment.op === "C") {
      const c1 = apply(total, segment.x1, segment.y1);
      const c2 = apply(total, segment.x2, segment.y2);
      return { op: "C", x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y, x: end.x, y: end.y };
    }
    return { op: segment.op, x: end.x, y: end.y };
  });
  let clip: PathItem["clip"];
  if (shape.clipRect) {
    // the crop is in the layer's coordinates: kept only where it stays a box
    if (Math.abs(outer[1]) > 1e-6 || Math.abs(outer[2]) > 1e-6) {
      omissions.push(`${label}: crop of a turned outline`);
    } else {
      const a = apply(outer, shape.clipRect.x, shape.clipRect.y);
      const b = apply(outer, shape.clipRect.x + shape.clipRect.width, shape.clipRect.y + shape.clipRect.height);
      clip = { left: Math.min(a.x, b.x), top: Math.min(a.y, b.y), right: Math.max(a.x, b.x), bottom: Math.max(a.y, b.y) };
    }
  }
  return {
    kind: "path",
    segments,
    fill,
    strokeWidth: strokeWidth * lengthScale(total),
    lineCap: shape.lineCap,
    lineJoin: shape.lineJoin,
    clip,
  };
}

interface Frame {
  /** the box's left edge and the baseline, in symbol coordinates */
  left: number;
  baseline: number;
  metrics: CustomSymbolMetrics;
}

interface Usage {
  graphicx: boolean;
  pict2e: boolean;
  trimclip: boolean;
  pin: boolean;
}

const tx = (frame: Frame, x: number) => em(x - frame.left);
const ty = (frame: Frame, y: number) => em(frame.baseline - y);

function nodeCommand(item: NodeItem, frame: Frame, usage: Usage) {
  let content = `$${item.source}$`;
  if (item.clip) {
    usage.trimclip = true;
    const { x, y, width, height } = item.clip;
    content =
      `\\clipbox*{${em(x)}em ${em(item.ascent - (y + height))}em ` +
      `${em(x + width)}em ${em(item.ascent - y)}em}{${content}}`;
  }
  const reference = apply(item.matrix, 0, item.ascent);
  const [a, b, c, d] = item.matrix;
  if (!isIdentityLinear(item.matrix)) {
    usage.graphicx = true;
    // the designer's y grows downward: the same map with y up
    let { phi, sx, sy, theta } = decomposeLinear(a, -c, -b, d);
    if (Math.abs(sx - sy) < 1e-6) {
      // a uniform scaling turns with the rotation: one \rotatebox
      phi += theta;
      theta = 0;
    }
    phi = ((((phi + 180) % 360) + 360) % 360) - 180;
    // \scalebox scales about the reference point; turning and mirroring move
    // the box, so the material is pinned to a box of no size first
    if (Math.abs(theta) > 1e-4 || Math.abs(phi) > 1e-4 || sx < 0 || sy < 0) {
      usage.pin = true;
      content = `\\vtxpin{${content}}`;
    }
    if (Math.abs(theta) > 1e-4) content = `\\rotatebox{${num(theta)}}{${content}}`;
    if (Math.abs(sx - 1) > 1e-5 || Math.abs(sy - 1) > 1e-5) {
      content = Math.abs(sx - sy) < 1e-5 ? `\\scalebox{${num(sx, 5)}}{${content}}` : `\\scalebox{${num(sx, 5)}}[${num(sy, 5)}]{${content}}`;
    }
    if (Math.abs(phi) > 1e-4) content = `\\rotatebox{${num(phi)}}{${content}}`;
  }
  return `\\put(${tx(frame, reference.x)},${ty(frame, reference.y)}){${content}}`;
}

function pathCommands(item: PathItem, frame: Frame, usage: Usage) {
  usage.pict2e = true;
  const path = item.segments
    .map((segment) => {
      switch (segment.op) {
        case "M":
          return `\\moveto(${tx(frame, segment.x)},${ty(frame, segment.y)})`;
        case "L":
          return `\\lineto(${tx(frame, segment.x)},${ty(frame, segment.y)})`;
        case "C":
          return (
            `\\curveto(${tx(frame, segment.x1)},${ty(frame, segment.y1)})` +
            `(${tx(frame, segment.x2)},${ty(frame, segment.y2)})(${tx(frame, segment.x)},${ty(frame, segment.y)})`
          );
        case "Z":
          return "\\closepath";
      }
    })
    .join("");
  const parts: string[] = [];
  if (item.fill) parts.push(`${path}\\fillpath`);
  if (item.strokeWidth > 0) {
    const cap = item.lineCap === "round" ? "\\roundcap" : item.lineCap === "square" ? "\\squarecap" : "\\buttcap";
    const join = item.lineJoin === "round" ? "\\roundjoin" : item.lineJoin === "bevel" ? "\\beveljoin" : "\\miterjoin";
    parts.push(`\\linethickness{${em(item.strokeWidth)}em}${cap}${join}${path}\\strokepath`);
  }
  let body = parts.join("");
  if (item.clip) {
    // a picture of no size at the origin, cropped to the box (in its coordinates)
    usage.trimclip = true;
    const { left, top, right, bottom } = item.clip;
    body =
      `\\put(0,0){\\clipbox*{${tx(frame, left)}em ${ty(frame, bottom)}em ${tx(frame, right)}em ${ty(frame, top)}em}` +
      `{\\begin{picture}(0,0)${body}\\end{picture}}}`;
  }
  return body;
}

/** The picture of `items` in the box `frame`, inside \text so it follows the math style's size. */
function pictureBox(items: Item[], frame: Frame, usage: Usage, indent: string) {
  const { widthEm, ascentEm, descentEm } = frame.metrics;
  const lines = items.map((item) => (item.kind === "node" ? nodeCommand(item, frame, usage) : pathCommands(item, frame, usage)));
  return [
    `\\text{\\setlength{\\unitlength}{1em}\\raisebox{-${num(descentEm, 5)}em}{%`,
    `${indent}  \\begin{picture}(${num(widthEm, 5)},${num(ascentEm + descentEm, 5)})(0,-${num(descentEm, 5)})%`,
    ...lines.map((line) => `${indent}    ${line}%`),
    `${indent}  \\end{picture}}}`,
  ].join("\n");
}

/**
 * The text-size version of the symbol: display operators replaced by their
 * text size, everything else scaled about the base's reference point by the
 * ratio of the two sizes of the base.
 */
function textStyleItems(
  items: Item[],
  base: CustomSymbolDesignerGlyphLayer,
  textOf: (source: string) => TextStyleMaterial | undefined,
): { items: Item[]; ratio: number; widthRatio: number } | null {
  const baseText = textOf(base.asset.sourceLatex);
  if (!baseText) return null;
  // a text-size operator is not the display one made smaller (\int: two thirds
  // as wide, half as tall): positions follow each ratio, and what sits over
  // the operator keeps its proportions at the width's ratio, so it still
  // spans the signs
  const displayHeight = base.asset.metrics.ascentEm + base.asset.metrics.descentEm;
  const ratio = (baseText.metrics.ascentEm + baseText.metrics.descentEm) / displayHeight;
  const widthRatio = baseText.metrics.widthEm / base.asset.metrics.widthEm;
  const outer = transformMatrix(base.transform, false);
  const anchor = apply(outer, 0, base.asset.metrics.ascentEm * UNITS);
  const about: Affine = [widthRatio, 0, 0, ratio, anchor.x * (1 - widthRatio), anchor.y * (1 - ratio)];
  const scaled = items.map((item): Item => {
    if (item.kind === "path") {
      return {
        ...item,
        strokeWidth: item.strokeWidth * Math.sqrt(ratio * widthRatio),
        segments: item.segments.map((segment): PathSegment => {
          if (segment.op === "Z") return segment;
          const end = apply(about, segment.x, segment.y);
          if (segment.op === "C") {
            const c1 = apply(about, segment.x1, segment.y1);
            const c2 = apply(about, segment.x2, segment.y2);
            return { op: "C", x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y, x: end.x, y: end.y };
          }
          return { op: segment.op, x: end.x, y: end.y };
        }),
        clip: item.clip && {
          left: apply(about, item.clip.left, 0).x,
          right: apply(about, item.clip.right, 0).x,
          top: apply(about, 0, item.clip.top).y,
          bottom: apply(about, 0, item.clip.bottom).y,
        },
      };
    }
    const textSource = textStyleVariant(item.source);
    const text = textSource ? textOf(item.source) : undefined;
    if (!textSource || !text) {
      // placed by both ratios, sized by the width's
      const designed = apply(item.matrix, 0, item.ascent);
      const placed = apply(about, designed.x, designed.y);
      const [a, b, c, d] = item.matrix.map((value) => value * widthRatio);
      return { ...item, matrix: [a, b, c, d, placed.x - c * item.ascent, placed.y - d * item.ascent] };
    }
    // the operator in its text size, its reference point where the scaled design puts it
    const designed = apply(item.matrix, 0, item.ascent);
    const reference = apply(about, designed.x, designed.y);
    const ascent = text.metrics.ascentEm * UNITS;
    const [a, b, c, d] = item.matrix;
    return {
      kind: "node",
      source: textSource,
      ascent,
      matrix: [a, b, c, d, reference.x - c * ascent, reference.y - d * ascent],
      clip: item.clip && {
        x: item.clip.x * ratio,
        y: ascent - (item.ascent - item.clip.y) * ratio,
        width: item.clip.width * ratio,
        height: item.clip.height * ratio,
      },
    };
  });
  return { items: scaled, ratio, widthRatio };
}

/**
 * The base a symbol can keep as the real operator: an operator symbol on a
 * large operator (\int, \iint, \sum, ...) that is not scaled, turned,
 * cropped or given effects. TeX then sets that operator itself (its box,
 * italic correction, limits and sizes, in the document's own math font)
 * and the command only adds the other layers over it.
 */
export function nativeOperatorBase(document: CustomSymbolDesignerDocument, box: CustomSymbolBox | null) {
  const base = box?.base;
  if (document.role !== "operator" || !base || !isLargeOperatorMaterial(base.asset.sourceLatex)) return null;
  const t = base.transform;
  const unit = (value: number | undefined) => Math.abs((value ?? 1) - 1) < 1e-6;
  if (!unit(t.scaleX) || !unit(t.scaleY) || !isUprightLayer(base) || base.clipRect) return null;
  if (base.effects?.outline?.enabled || base.effects?.perspective?.enabled) return null;
  return base;
}

/** The other layers in a picture of no size whose origin is the base's reference point. */
function overlayPicture(items: Item[], reference: { x: number; y: number }, usage: Usage) {
  const frame: Frame = { left: reference.x, baseline: reference.y, metrics: { widthEm: 0, ascentEm: 0, descentEm: 0 } };
  const lines = items.map((item) => (item.kind === "node" ? nodeCommand(item, frame, usage) : pathCommands(item, frame, usage)));
  return `\\text{\\setlength{\\unitlength}{1em}\\begin{picture}(0,0)${lines.join("")}\\end{picture}}`;
}

/** The LaTeX command for `document`, whose registered box is `box`. */
export function customSymbolTexExport(
  document: CustomSymbolDesignerDocument,
  box: CustomSymbolBox | null,
  options: CustomSymbolTexOptions = {},
): CustomSymbolTexExport {
  const command = document.command.trim().replace(/^\\/, "") || "mysymbol";
  if (!box) return { code: `% \\${command}: nothing to draw yet`, omissions: [] };
  const omissions: string[] = [];
  const usage: Usage = { graphicx: false, pict2e: false, trimclip: false, pin: false };
  const items = layerItems(document, omissions);
  const displayFrame: Frame = { left: box.left, baseline: box.baseline, metrics: box.metrics };
  const wrapper = ROLE_WRAPPERS[document.role] ?? "\\mathord";
  const operator = document.role === "operator";
  const limitsWord = (display: boolean) =>
    document.limitsBehavior === "limits" || (document.limitsBehavior === "auto" && display) ? "\\limits" : "\\nolimits";

  let definition: string;
  const textOf = options.textStyle;
  const native = nativeOperatorBase(document, box);
  if (native) {
    return nativeOperatorExport(document, native, command, options, omissions);
  }
  const text = operator && box.base && textOf ? textStyleItems(items, box.base, textOf) : null;
  if (!operator || (box.italicEm <= 0 && !text)) {
    const limits = operator && document.limitsBehavior !== "auto" ? limitsWord(false) : "";
    definition = `\\newcommand{\\${command}}{${wrapper}{${pictureBox(items, displayFrame, usage, "")}}${limits}}`;
  } else {
    // the box and italic correction of each size; the scripts are set in each
    // branch so the italic correction is that size's
    let textFrame = displayFrame;
    let textItalic = box.italicEm;
    if (text && box.base) {
      const baseText = textOf!(box.base.asset.sourceLatex)!;
      const outer = transformMatrix(box.base.transform, false);
      const anchor = apply(outer, 0, box.base.asset.metrics.ascentEm * UNITS);
      const scaleX = Math.abs(box.base.transform.scaleX ?? 1);
      const scaleY = Math.abs(box.base.transform.scaleY ?? 1);
      const width = baseText.metrics.widthEm * UNITS * scaleX;
      const ascent = baseText.metrics.ascentEm * scaleY;
      const descent = baseText.metrics.descentEm * scaleY;
      const left = (box.base.transform.scaleX ?? 1) < 0 ? anchor.x - width : anchor.x;
      // the room the overlays take beyond the base's box, scaled with the design
      const baseBox = layerTexBox(box.base);
      const extraLeft = (baseBox.left - box.left) * text.widthRatio;
      const extraRight = (box.left + box.metrics.widthEm * UNITS - baseBox.right) * text.widthRatio;
      const extraTop = (baseBox.top - (box.baseline - box.metrics.ascentEm * UNITS)) * text.ratio;
      const extraBottom = (box.baseline + box.metrics.descentEm * UNITS - baseBox.bottom) * text.ratio;
      textFrame = {
        left: left - extraLeft,
        baseline: box.baseline,
        metrics: {
          widthEm: (width + extraLeft + extraRight) / UNITS,
          ascentEm: Math.max(0.02, ascent + (box.baseline - anchor.y + extraTop) / UNITS),
          descentEm: Math.max(0, descent - (box.baseline - anchor.y - extraBottom) / UNITS),
        },
      };
      // overlays reaching out on the right take the italic correction's room (symbolBox.ts)
      textItalic = Math.max(0, baseText.italicEm * scaleX - extraRight / UNITS);
    }
    // #1 \limits, #2 \nolimits typed after the command; #3, #4 the limits
    const scripts = (limits: boolean, italic: number, quadFont: string) => {
      // \relax ends the font number before a script that starts with a digit
      const delta = (factor: number) => `${num(italic * factor, 5)}\\fontdimen6${quadFont}\\relax`;
      if (italic <= 0) return `\\IfValueT{#3}{^{#3}}\\IfValueT{#4}{_{#4}}`;
      return limits
        ? `\\IfValueT{#3}{^{\\kern${delta(0.5)}#3\\kern-${delta(0.5)}}}\\IfValueT{#4}{_{\\kern-${delta(0.5)}#4\\kern${delta(0.5)}}}`
        : `\\IfValueT{#3}{^{#3}}\\IfValueT{#4}{_{\\kern-${delta(1)}#4}}`;
    };
    const branch = (frame: Frame, branchItems: Item[], italic: number, display: boolean, quadFont: string) => {
      const byDefault = limitsWord(display) === "\\limits";
      const placed = (limits: boolean) => `${limits ? "\\limits" : "\\nolimits"}${scripts(limits, italic, quadFont)}`;
      return (
        `{\\mathop{${pictureBox(branchItems, frame, usage, "    ")}}` +
        `\\IfBooleanTF{#1}{${placed(true)}}{\\IfBooleanTF{#2}{${placed(false)}}{${placed(byDefault)}}}}`
      );
    };
    const textItems = text ? text.items : items;
    definition = [
      `\\NewDocumentCommand{\\${command}}{t\\limits t\\nolimits e{^_}}{\\mathchoice`,
      `  ${branch(displayFrame, items, box.italicEm, true, "\\textfont2")}`,
      `  ${branch(textFrame, textItems, textItalic, false, "\\textfont2")}`,
      `  ${branch(textFrame, textItems, textItalic, false, "\\scriptfont2")}`,
      `  ${branch(textFrame, textItems, textItalic, false, "\\scriptscriptfont2")}}`,
    ].join("\n");
  }

  return { code: [...header(document, command, usage, options.mathFont), definition].join("\n"), omissions: [...new Set(omissions)] };
}

/** The comment lines (packages, the font the positions were designed in) and helpers before the definition. */
function header(document: CustomSymbolDesignerDocument, command: string, usage: Usage, mathFont: "cm" | "ntx" | undefined) {
  const packages = ["amsmath", usage.graphicx && "graphicx", usage.pict2e && "pict2e", usage.trimclip && "trimclip"].filter(Boolean);
  const name = document.name.trim();
  const lines = [
    `% ${name && name !== command ? `${name} — ` : ""}VisualTeX custom symbol \\${command}`,
    `% Preamble: \\usepackage{${packages.join(",")}}`,
  ];
  if (mathFont === "ntx") {
    lines.push("% Designed in the Times math font: with \\usepackage{newtxmath} it looks exactly as designed");
  }
  if (usage.pin) {
    lines.push(
      "% \\vtxpin: a box of no size at the material's reference point, so turning it turns it about that point",
      "\\providecommand*\\vtxpin[1]{\\makebox[0pt][l]{\\raisebox{0pt}[0pt][0pt]{#1}}}",
    );
  }
  return lines;
}

/**
 * An operator kept as the real operator (nativeOperatorBase): the other
 * layers in a picture of no size at its reference point, one per math style
 * (text and script styles use the text-size design), as an empty \mathop
 * whose space before the operator \mskip-\thinmuskip takes back; then the
 * operator itself, which TeX sets as always.
 */
function nativeOperatorExport(
  document: CustomSymbolDesignerDocument,
  base: CustomSymbolDesignerGlyphLayer,
  command: string,
  options: CustomSymbolTexOptions,
  omissions: string[],
): CustomSymbolTexExport {
  const usage: Usage = { graphicx: false, pict2e: false, trimclip: false, pin: false };
  const others: CustomSymbolDesignerDocument = { ...document, layers: document.layers.filter((layer) => layer.id !== base.id) };
  const items = layerItems(others, omissions);
  const reference = apply(transformMatrix(base.transform, false), 0, base.asset.metrics.ascentEm * UNITS);
  // \displaystyle\iint -> \iint: TeX picks the size of the style it is in
  const operator = base.asset.sourceLatex.trim().replace(/^\\(?:display|text)style\s*/, "");
  const limits =
    document.limitsBehavior === "limits" ? "\\limits" : document.limitsBehavior === "nolimits" ? "\\nolimits" : "\\displaylimits";
  let definition: string;
  if (!items.length) {
    definition = `\\newcommand{\\${command}}{${operator}${limits}}`;
  } else {
    const display = overlayPicture(items, reference, usage);
    const text = options.textStyle ? textStyleItems(items, base, options.textStyle) : null;
    const small = text ? overlayPicture(text.items, reference, usage) : display;
    definition = [
      `\\newcommand{\\${command}}{\\mathop{\\mathchoice`,
      `    {${display}}`,
      `    {${small}}`,
      `    {${small}}`,
      `    {${small}}}%`,
      `  \\mskip-\\thinmuskip${operator}${limits}}`,
    ].join("\n");
  }
  return { code: [...header(document, command, usage, options.mathFont), definition].join("\n"), omissions: [...new Set(omissions)] };
}
