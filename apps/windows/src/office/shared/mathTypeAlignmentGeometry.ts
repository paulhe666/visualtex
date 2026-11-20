const MTEF_RULER_STOPS_ATTRIBUTE = "data-visualtex-mtef-ruler-stops";

function directChildrenByMmlNode(element: Element, nodeName: string) {
  return Array.from(element.children).filter(
    (child) => child.getAttribute("data-mml-node") === nodeName,
  );
}

function directMathMlChildren(element: Element, localName: string) {
  return Array.from(element.children).filter(
    (child) => child.localName === localName,
  );
}

function alignedColumnCount(table: Element) {
  const rows = directMathMlChildren(table, "mtr").concat(
    directMathMlChildren(table, "mlabeledtr"),
  );
  if (!rows.length) return 0;
  const columnCount = Math.max(
    ...rows.map((row) => directMathMlChildren(row, "mtd").length),
  );
  if (columnCount < 2 || columnCount % 2 !== 0) return 0;
  const alignment = (table.getAttribute("columnalign") ?? "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  if (alignment.length < columnCount) return 0;
  for (let column = 0; column < columnCount; column += 1) {
    const expected = column % 2 === 0 ? "right" : "left";
    if (alignment[column]?.toLowerCase() !== expected) return 0;
  }
  return columnCount;
}

function readTranslateX(value: string | null) {
  if (!value) return 0;
  const match = value.match(
    /(?:^|\s)translate\(\s*([-+]?\d+(?:\.\d+)?(?:e[-+]?\d+)?)(?:[ ,][^)]*)?\)/i,
  );
  if (!match) return 0;
  const parsed = Number.parseFloat(match[1]);
  return Number.isFinite(parsed) ? parsed : 0;
}

function readViewBoxWidth(svgRoot: Element) {
  const values = (svgRoot.getAttribute("viewBox") ?? "")
    .trim()
    .split(/[\s,]+/)
    .map(Number);
  if (values.length !== 4 || !values.every(Number.isFinite) || values[2] <= 0) {
    return null;
  }
  return values[2];
}

/**
 * Annotate MathJax's Presentation MathML with MathType-native RULER offsets.
 *
 * MathJax has already solved the TeX aligned layout when it creates the SVG:
 * direct mtd transforms are the exact right/left pair anchors.  Reusing those
 * transforms avoids a second MathType/MathPage render just to measure columns.
 * MTEF dimensional values are 1/32 printer points; CSS pixels are 3/4 point.
 */
export function annotateMathTypeAlignmentGeometry(
  mathMl: string,
  svg: string,
  svgWidthPx: number,
) {
  if (
    typeof DOMParser === "undefined" ||
    typeof XMLSerializer === "undefined" ||
    !mathMl.trim() ||
    !svg.trim() ||
    !Number.isFinite(svgWidthPx) ||
    svgWidthPx <= 0
  ) {
    return mathMl;
  }

  try {
    const parser = new DOMParser();
    const mathDocument = parser.parseFromString(mathMl, "application/xml");
    const mathRoot = mathDocument.documentElement;
    if (!mathRoot || mathRoot.localName === "parsererror") return mathMl;

    const mathTables = Array.from(mathRoot.getElementsByTagName("*")).filter(
      (element) => element.localName === "mtable" && alignedColumnCount(element) > 0,
    );
    if (mathRoot.localName === "mtable" && alignedColumnCount(mathRoot) > 0) {
      mathTables.unshift(mathRoot);
    }
    if (!mathTables.length) return mathMl;

    const svgDocument = parser.parseFromString(svg, "image/svg+xml");
    const svgRoot = svgDocument.documentElement;
    if (!svgRoot || svgRoot.localName === "parsererror") return mathMl;
    const viewBoxWidth = readViewBoxWidth(svgRoot);
    if (!viewBoxWidth) return mathMl;
    const unitsPerPx = viewBoxWidth / svgWidthPx;
    if (!Number.isFinite(unitsPerPx) || unitsPerPx <= 0) return mathMl;

    const svgTables = Array.from(svgDocument.querySelectorAll('[data-mml-node="mtable"]'));
    const usedSvgTables = new Set<Element>();

    for (const mathTable of mathTables) {
      const columnCount = alignedColumnCount(mathTable);
      if (!columnCount) continue;
      const mathRows = directMathMlChildren(mathTable, "mtr").concat(
        directMathMlChildren(mathTable, "mlabeledtr"),
      );

      const svgTable = svgTables.find((candidate) => {
        if (usedSvgTables.has(candidate)) return false;
        const rows = directChildrenByMmlNode(candidate, "mtr");
        if (rows.length !== mathRows.length || !rows.length) return false;
        return rows.every(
          (row) => directChildrenByMmlNode(row, "mtd").length === columnCount,
        );
      });
      if (!svgTable) continue;
      usedSvgTables.add(svgTable);

      const firstRow = directChildrenByMmlNode(svgTable, "mtr")[0];
      const cells = directChildrenByMmlNode(firstRow, "mtd");
      if (cells.length !== columnCount) continue;

      const stops: number[] = [];
      let previous = -1;
      for (let column = 1; column < columnCount; column += 2) {
        const xUnits = readTranslateX(cells[column].getAttribute("transform"));
        // CSS px -> printer pt -> MathType internal 1/32pt units:
        // xUnits / unitsPerPx * (72/96) * 32 = xUnits / unitsPerPx * 24.
        const internalUnits = Math.round((xUnits / unitsPerPx) * 24);
        if (
          !Number.isFinite(internalUnits) ||
          internalUnits <= previous ||
          internalUnits <= 0 ||
          internalUnits > 0xffff
        ) {
          stops.length = 0;
          break;
        }
        stops.push(internalUnits);
        previous = internalUnits;
      }
      if (stops.length !== columnCount / 2) continue;
      mathTable.setAttribute(MTEF_RULER_STOPS_ATTRIBUTE, stops.join(","));
    }

    return new XMLSerializer().serializeToString(mathRoot);
  } catch {
    // Geometry metadata is an optimization/correctness hint for MathType only.
    // Never turn an otherwise valid Office export into a hard frontend failure.
    return mathMl;
  }
}

export { MTEF_RULER_STOPS_ATTRIBUTE };
