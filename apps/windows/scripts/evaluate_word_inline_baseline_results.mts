import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

interface Metric {
  Name: string;
  FontSizePt: number;
  FormulaLetterFont: string;
  Position: number;
  RenderedDescentPt: number;
  ExpectedNearestPosition: number;
  PreviewInkLeftMarginPx: number;
  PreviewInkTopMarginPx: number;
  PreviewInkRightMarginPx: number;
  PreviewInkBottomMarginPx: number;
  BodyDeltaTopPt: number;
  BodyDeltaBottomPt: number;
  BodyDeltaCentroidPt: number;
  HTopSpreadPt: number;
  HBottomSpreadPt: number;
}

interface ProbeResult {
  Source: string;
  SourceUnchanged: boolean;
  TemporaryDocumentSaved: boolean;
  Metrics: Metric[];
}

function argument(name: string) {
  const prefix = `--${name}=`;
  return process.argv.find((value) => value.startsWith(prefix))?.slice(prefix.length);
}

function finite(value: unknown, label: string) {
  assert.equal(typeof value, "number", `${label} is not numeric`);
  assert.ok(Number.isFinite(value as number), `${label} is not finite`);
  return value as number;
}

function spread(values: number[]) {
  return Math.max(...values) - Math.min(...values);
}

async function readResult(path: string): Promise<ProbeResult> {
  return JSON.parse(await readFile(path, "utf8")) as ProbeResult;
}

function validateCommon(label: string, result: ProbeResult, expectedCount: number) {
  assert.equal(result.SourceUnchanged, true, `${label}: source Word document changed`);
  assert.equal(
    result.TemporaryDocumentSaved,
    false,
    `${label}: temporary Word document was unexpectedly saved`,
  );
  assert.equal(
    result.Metrics.length,
    expectedCount,
    `${label}: expected ${expectedCount} metrics, received ${result.Metrics.length}`,
  );

  for (const metric of result.Metrics) {
    const context = `${label}/${metric.Name}`;
    assert.equal(
      metric.Position,
      metric.ExpectedNearestPosition,
      `${context}: Word Position is not the nearest-point rendered descent (${metric.RenderedDescentPt}pt)`,
    );
    for (const [name, value] of [
      ["PreviewInkLeftMarginPx", metric.PreviewInkLeftMarginPx],
      ["PreviewInkTopMarginPx", metric.PreviewInkTopMarginPx],
      ["PreviewInkRightMarginPx", metric.PreviewInkRightMarginPx],
      ["PreviewInkBottomMarginPx", metric.PreviewInkBottomMarginPx],
    ] as const) {
      assert.ok(
        finite(value, `${context}/${name}`) > 0,
        `${context}: final preview ink touches or crosses the ${name} edge`,
      );
    }
    assert.ok(
      Math.abs(finite(metric.HTopSpreadPt, `${context}/HTopSpreadPt`)) <= 0.35,
      `${context}: surrounding prose top baseline drifted`,
    );
    assert.ok(
      Math.abs(finite(metric.HBottomSpreadPt, `${context}/HBottomSpreadPt`)) <= 0.35,
      `${context}: surrounding prose bottom baseline drifted`,
    );
    finite(metric.BodyDeltaTopPt, `${context}/BodyDeltaTopPt`);
    finite(metric.BodyDeltaBottomPt, `${context}/BodyDeltaBottomPt`);
    finite(metric.BodyDeltaCentroidPt, `${context}/BodyDeltaCentroidPt`);
  }
}

function validateFormulaIndependence(label: string, metrics: Metric[]) {
  const groups = new Map<string, Metric[]>();
  for (const metric of metrics) {
    const key = `${metric.FormulaLetterFont}/${metric.FontSizePt}`;
    const group = groups.get(key) ?? [];
    group.push(metric);
    groups.set(key, group);
  }

  const summary = [];
  for (const [groupName, group] of groups) {
    const fontSizePt = finite(group[0]?.FontSizePt, `${label}/${groupName}/fontSizePt`);
    const centroidValues = group.map((metric) => metric.BodyDeltaCentroidPt);
    const bottomValues = group.map((metric) => metric.BodyDeltaBottomPt);
    const centroidSpread = spread(centroidValues);
    const bottomSpread = spread(bottomValues);
    const allowedSpread = Math.max(1.25, fontSizePt * 0.025);
    assert.ok(
      centroidSpread <= allowedSpread,
      `${label}/${groupName}: identical anchor glyph has formula-dependent centroid spread ${centroidSpread.toFixed(3)}pt`,
    );
    assert.ok(
      bottomSpread <= allowedSpread,
      `${label}/${groupName}: identical anchor glyph has formula-dependent bottom spread ${bottomSpread.toFixed(3)}pt`,
    );
    summary.push({
      group: groupName,
      count: group.length,
      centroidMeanPt: Number(
        (centroidValues.reduce((sum, value) => sum + value, 0) / group.length).toFixed(3),
      ),
      centroidSpreadPt: Number(centroidSpread.toFixed(3)),
      bottomSpreadPt: Number(bottomSpread.toFixed(3)),
      minimumInkMarginPx: Math.min(
        ...group.flatMap((metric) => [
          metric.PreviewInkLeftMarginPx,
          metric.PreviewInkTopMarginPx,
          metric.PreviewInkRightMarginPx,
          metric.PreviewInkBottomMarginPx,
        ]),
      ),
    });
  }
  return summary;
}

async function main() {
  const structuresPath = argument("structures");
  const sizesPath = argument("sizes");
  const fontsPath = argument("fonts");
  if (!structuresPath || !sizesPath || !fontsPath) {
    throw new Error(
      "Usage: tsx scripts/evaluate_word_inline_baseline_results.mts --structures=<json> --sizes=<json> --fonts=<json>",
    );
  }

  const structures = await readResult(structuresPath);
  const sizes = await readResult(sizesPath);
  const fonts = await readResult(fontsPath);
  validateCommon("structures", structures, 19);
  validateCommon("sizes", sizes, 48);
  validateCommon("fonts", fonts, 36);

  const byName = new Map(structures.Metrics.map((metric) => [metric.Name, metric]));
  for (const name of [
    "plain-x-10.5pt-katex",
    "plain-L-10.5pt-katex",
    "superscript-10.5pt-katex",
  ]) {
    const metric = byName.get(name);
    assert.ok(metric, `structures: missing ${name}`);
    assert.equal(
      metric.Position,
      -1,
      `${name}: a sub-one-point rendered descent was incorrectly collapsed to Position=0`,
    );
  }

  const summary = [
    ...validateFormulaIndependence("structures", structures.Metrics),
    ...validateFormulaIndependence("sizes", sizes.Metrics),
    ...validateFormulaIndependence("fonts", fonts.Metrics),
  ];
  console.table(summary);
  console.log(
    `Real Word inline baseline acceptance passed for ${
      structures.Metrics.length + sizes.Metrics.length + fonts.Metrics.length
    } OLE/OMML comparisons.`,
  );
}

await main();
