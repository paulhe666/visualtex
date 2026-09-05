import { latexToSvg } from "../export/runtime.ts";
import type {
  CustomSymbolMetrics,
  CustomSymbolVectorMatrix,
  CustomSymbolVectorShape,
} from "./customSymbolTypes.ts";
import type { CustomSymbolGlyphAsset } from "./customSymbolDesignerTypes.ts";

type Matrix = CustomSymbolVectorMatrix;

const identityMatrix: Matrix = [1, 0, 0, 1, 0, 0];
const maximumGlyphSourceLength = 2_000;
const maximumCompiledShapes = 160;

function finite(value: string | null, fallback = 0) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : fallback;
}

function multiply(left: Matrix, right: Matrix): Matrix {
  const [a1, b1, c1, d1, e1, f1] = left;
  const [a2, b2, c2, d2, e2, f2] = right;
  return [
    a1 * a2 + c1 * b2,
    b1 * a2 + d1 * b2,
    a1 * c2 + c1 * d2,
    b1 * c2 + d1 * d2,
    a1 * e2 + c1 * f2 + e1,
    b1 * e2 + d1 * f2 + f1,
  ];
}

function translation(x: number, y: number): Matrix {
  return [1, 0, 0, 1, x, y];
}

function scale(x: number, y: number): Matrix {
  return [x, 0, 0, y, 0, 0];
}

function rotation(degrees: number): Matrix {
  const radians = (degrees * Math.PI) / 180;
  const cosine = Math.cos(radians);
  const sine = Math.sin(radians);
  return [cosine, sine, -sine, cosine, 0, 0];
}

function skewX(degrees: number): Matrix {
  return [1, 0, Math.tan((degrees * Math.PI) / 180), 1, 0, 0];
}

function skewY(degrees: number): Matrix {
  return [1, Math.tan((degrees * Math.PI) / 180), 0, 1, 0, 0];
}

function parseNumberList(value: string) {
  return value
    .trim()
    .split(/[\s,]+/)
    .filter(Boolean)
    .map(Number)
    .filter(Number.isFinite);
}

function parseSvgTransform(value: string | null): Matrix {
  if (!value?.trim()) return identityMatrix;
  let matrix = identityMatrix;
  const pattern = /([A-Za-z]+)\s*\(([^)]*)\)/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(value))) {
    const name = match[1].toLowerCase();
    const args = parseNumberList(match[2]);
    let operation = identityMatrix;
    if (name === "matrix" && args.length === 6) {
      operation = args as Matrix;
    } else if (name === "translate" && args.length >= 1) {
      operation = translation(args[0], args[1] ?? 0);
    } else if (name === "scale" && args.length >= 1) {
      operation = scale(args[0], args[1] ?? args[0]);
    } else if (name === "rotate" && args.length >= 1) {
      operation = rotation(args[0]);
      if (args.length >= 3) {
        operation = multiply(
          translation(args[1], args[2]),
          multiply(operation, translation(-args[1], -args[2])),
        );
      }
    } else if (name === "skewx" && args.length >= 1) {
      operation = skewX(args[0]);
    } else if (name === "skewy" && args.length >= 1) {
      operation = skewY(args[0]);
    } else {
      throw new Error(`Unsupported SVG transform in LaTeX material: ${match[1]}.`);
    }
    matrix = multiply(matrix, operation);
  }
  return matrix;
}

function approximatelyIdentity(matrix: Matrix) {
  return matrix.every((value, index) =>
    Math.abs(value - identityMatrix[index]) < 0.000001,
  );
}

function normalizedMatrix(matrix: Matrix) {
  if (approximatelyIdentity(matrix)) return undefined;
  return matrix.map((value) => Number(value.toFixed(6))) as Matrix;
}

function paintShapeBase(element: Element, defaultFill: boolean) {
  const fillAttribute = element.getAttribute("fill");
  const strokeAttribute = element.getAttribute("stroke");
  const fill = fillAttribute === "none" ? false : defaultFill;
  const strokeWidth = finite(element.getAttribute("stroke-width"), 0);
  return {
    fill,
    ...(strokeWidth > 0 || (strokeAttribute && strokeAttribute !== "none")
      ? { strokeWidth: strokeWidth > 0 ? strokeWidth : 50 }
      : {}),
    ...(element.getAttribute("stroke-linecap") === "round" ||
    element.getAttribute("stroke-linecap") === "square" ||
    element.getAttribute("stroke-linecap") === "butt"
      ? {
          lineCap: element.getAttribute("stroke-linecap") as
            | "round"
            | "square"
            | "butt",
        }
      : {}),
    ...(element.getAttribute("stroke-linejoin") === "round" ||
    element.getAttribute("stroke-linejoin") === "bevel" ||
    element.getAttribute("stroke-linejoin") === "miter"
      ? {
          lineJoin: element.getAttribute("stroke-linejoin") as
            | "round"
            | "bevel"
            | "miter",
        }
      : {}),
  };
}

function parseViewBox(svg: SVGSVGElement) {
  const values = parseNumberList(svg.getAttribute("viewBox") ?? "");
  if (values.length !== 4 || values[2] <= 0 || values[3] <= 0) {
    throw new Error("MathJax did not provide a usable SVG viewBox for this material.");
  }
  return { x: values[0], y: values[1], width: values[2], height: values[3] };
}

function metricsFromViewBox(viewBox: ReturnType<typeof parseViewBox>): CustomSymbolMetrics {
  const widthEm = viewBox.width / 1000;
  const ascentEm = Math.max(0.02, -viewBox.y / 1000);
  const descentEm = Math.max(0, (viewBox.y + viewBox.height) / 1000);
  return {
    widthEm: Number(widthEm.toFixed(6)),
    ascentEm: Number(ascentEm.toFixed(6)),
    descentEm: Number(descentEm.toFixed(6)),
  };
}

function rootDefinitions(svg: SVGSVGElement) {
  const definitions = new Map<string, Element>();
  svg.querySelectorAll("defs [id]").forEach((element) => {
    const id = element.getAttribute("id");
    if (id) definitions.set(id, element);
  });
  return definitions;
}

function directShape(
  element: Element,
  matrix: Matrix,
): CustomSymbolVectorShape | null {
  const tag = element.tagName.toLowerCase();
  const transform = normalizedMatrix(matrix);
  const withTransform = transform ? { transform: { matrix: transform } } : {};
  switch (tag) {
    case "path": {
      const d = element.getAttribute("d")?.trim() ?? "";
      if (!d) return null;
      return {
        kind: "path",
        d,
        ...paintShapeBase(element, true),
        ...withTransform,
      };
    }
    case "circle":
      return {
        kind: "circle",
        cx: finite(element.getAttribute("cx")),
        cy: finite(element.getAttribute("cy")),
        r: finite(element.getAttribute("r")),
        ...paintShapeBase(element, true),
        ...withTransform,
      };
    case "ellipse":
      return {
        kind: "ellipse",
        cx: finite(element.getAttribute("cx")),
        cy: finite(element.getAttribute("cy")),
        rx: finite(element.getAttribute("rx")),
        ry: finite(element.getAttribute("ry")),
        ...paintShapeBase(element, true),
        ...withTransform,
      };
    case "line":
      return {
        kind: "line",
        x1: finite(element.getAttribute("x1")),
        y1: finite(element.getAttribute("y1")),
        x2: finite(element.getAttribute("x2")),
        y2: finite(element.getAttribute("y2")),
        ...paintShapeBase(element, false),
        ...withTransform,
      };
    case "rect":
      return {
        kind: "rect",
        x: finite(element.getAttribute("x")),
        y: finite(element.getAttribute("y")),
        width: finite(element.getAttribute("width")),
        height: finite(element.getAttribute("height")),
        ...(element.hasAttribute("rx") ? { rx: finite(element.getAttribute("rx")) } : {}),
        ...(element.hasAttribute("ry") ? { ry: finite(element.getAttribute("ry")) } : {}),
        ...paintShapeBase(element, true),
        ...withTransform,
      };
    case "polygon": {
      const points = (element.getAttribute("points") ?? "")
        .trim()
        .split(/\s+/)
        .map((pair) => pair.split(",").map(Number))
        .filter(
          (pair): pair is [number, number] =>
            pair.length === 2 && pair.every(Number.isFinite),
        );
      if (points.length < 2) return null;
      return {
        kind: "polygon",
        points,
        ...paintShapeBase(element, true),
        ...withTransform,
      };
    }
    default:
      return null;
  }
}

function hrefId(element: Element) {
  const value =
    element.getAttribute("href") ??
    element.getAttribute("xlink:href") ??
    element.getAttributeNS("http://www.w3.org/1999/xlink", "href") ??
    "";
  return value.startsWith("#") ? value.slice(1) : "";
}

function isTransparentHitTarget(element: Element) {
  return (
    element.tagName.toLowerCase() === "rect" &&
    finite(element.getAttribute("fill-opacity"), 1) <= 0.002
  );
}

function flattenSvg(
  svg: SVGSVGElement,
  viewBox: ReturnType<typeof parseViewBox>,
) {
  const definitions = rootDefinitions(svg);
  const result: CustomSymbolVectorShape[] = [];
  const normalization = translation(-viewBox.x, -viewBox.y);
  const visitingReferences = new Set<string>();

  const visit = (element: Element, parentMatrix: Matrix, fromDefinition = false) => {
    const tag = element.tagName.toLowerCase();
    if (["defs", "style", "title", "desc", "metadata"].includes(tag)) return;
    if (!fromDefinition && isTransparentHitTarget(element)) return;
    if (["text", "image", "foreignobject", "a", "script"].includes(tag)) {
      throw new Error(
        "This LaTeX material contains text or external content that cannot be converted into an editable vector glyph.",
      );
    }

    const ownTransform = parseSvgTransform(element.getAttribute("transform"));
    let matrix = multiply(parentMatrix, ownTransform);

    if (tag === "use") {
      const id = hrefId(element);
      const referenced = definitions.get(id);
      if (!id || !referenced) {
        throw new Error("MathJax SVG contains an unresolved glyph reference.");
      }
      if (visitingReferences.has(id)) {
        throw new Error("MathJax SVG contains a recursive glyph reference.");
      }
      const x = finite(element.getAttribute("x"));
      const y = finite(element.getAttribute("y"));
      matrix = multiply(matrix, translation(x, y));
      visitingReferences.add(id);
      visit(referenced, matrix, true);
      visitingReferences.delete(id);
      return;
    }

    const shape = directShape(element, matrix);
    if (shape) {
      result.push(shape);
      if (result.length > maximumCompiledShapes) {
        throw new Error("This LaTeX material is too complex for one custom symbol layer.");
      }
      return;
    }

    if (tag === "svg" && element !== svg) {
      throw new Error(
        "Nested SVG material is not yet supported as a source glyph. Remove clipping from the source symbol or use the original LaTeX glyph.",
      );
    }

    for (const child of Array.from(element.children)) {
      visit(child, matrix, fromDefinition);
    }
  };

  for (const child of Array.from(svg.children)) {
    visit(child, normalization);
  }
  if (!result.length) {
    throw new Error("The LaTeX material did not contain editable vector geometry.");
  }
  return result;
}

export function compileLatexGlyphAsset(sourceLatex: string): CustomSymbolGlyphAsset {
  const source = sourceLatex.trim();
  if (!source || source.length > maximumGlyphSourceLength) {
    throw new Error("LaTeX symbol material is empty or too large.");
  }
  if (typeof DOMParser === "undefined") {
    throw new Error("LaTeX glyph compilation requires the VisualTeX browser runtime.");
  }
  const rendered = latexToSvg(source, {
    displayMode: false,
    fontSizePt: 12,
    paddingPx: 0,
    background: "transparent",
  });
  const documentNode = new DOMParser().parseFromString(
    rendered.svg,
    "image/svg+xml",
  );
  if (documentNode.querySelector("parsererror")) {
    throw new Error("MathJax produced SVG that could not be parsed for editing.");
  }
  const svg = documentNode.documentElement;
  if (svg.tagName.toLowerCase() !== "svg") {
    throw new Error("MathJax did not produce an SVG glyph.");
  }
  const svgElement = svg as unknown as SVGSVGElement;
  const viewBox = parseViewBox(svgElement);
  return {
    sourceLatex: source,
    metrics: metricsFromViewBox(viewBox),
    shapes: flattenSvg(svgElement, viewBox),
  };
}
