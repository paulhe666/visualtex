import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import {
  annotateCoreMathTypeAlignmentGeometry,
  generateCoreOfficeExport,
  rasterizeCoreOfficeExport,
} from "../src/office/compat/coreOfficeCompatibility.ts";

const require = createRequire(import.meta.url);
const { EngineRuntime } = require(resolve("../../engine/web/pkg-node/visualtex_engine_wasm.js"));
const fontPath = resolve("src-windows/artifacts/office-fonts/latinmodern-math.otf");
const fontBytes = readFileSync(fontPath);
const fontSha256 = createHash("sha256").update(fontBytes).digest("hex");
assert.equal(fontSha256, "6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f");

const runtime = new EngineRuntime();
const registered = JSON.parse(runtime.registerFont(fontBytes, 0));
assert.equal(registered.status, "ok");
const fontId = registered.result.font.id as string;
const baseLayout = {
  mathFont: fontId,
  textFont: null,
  mathFallbacks: [],
  customSymbolTextFonts: [],
  style: "display" as const,
  fontSizePx: 16,
  paddingEm: 0,
};
const cssPx = (pt: number) => pt * (96 / 72);
const close = (actual: number, expected: number, label: string) =>
  assert.ok(Math.abs(actual - expected) < 1e-7, `${label}: ${actual} != ${expected}`);
const decodeBase64 = (value: string) => Buffer.from(value, "base64").toString("utf8");
let owners = 0;
class ProtocolOwner {
  readonly handle: number;
  revision = 0;
  constructor(source: string, symbols: unknown[]) {
    const command = symbols.length
      ? { op: "createWithSymbols", documentId: `p11-office-${++owners}`, source, symbols }
      : { op: "create", documentId: `p11-office-${++owners}`, source };
    const created = this.send(command, "session");
    this.handle = created.handle;
    this.revision = created.snapshot.revision;
  }
  private send(command: unknown, kind: string) {
    const response = JSON.parse(runtime.dispatch(JSON.stringify({ protocol: 1, command })));
    assert.equal(response.status, "ok", JSON.stringify(response));
    assert.equal(response.result.kind, kind, JSON.stringify(response));
    return response.result;
  }
  layout(options: typeof baseLayout, includeSvg: boolean) {
    const result = this.send({ op: "layout", handle: this.handle, expectedRevision: this.revision, options, includeSvg }, "layout");
    return { sceneToken: result.sceneToken, snapshot: result.snapshot, svg: result.svg };
  }
  presentationMathMl(displayMode: boolean) {
    return this.send({ op: "exportPresentationMathMl", handle: this.handle, expectedRevision: this.revision, displayMode }, "presentationMathMl").result;
  }
  destroy() {
    this.send({ op: "destroy", handle: this.handle, expectedRevision: this.revision }, "destroyed");
  }
}
const owner = (source: string, symbols: unknown[] = []) => new ProtocolOwner(source, symbols);

try {
  {
    const formula = owner(String.raw`\frac{x_1+1}{y^2}`);
    const zero = formula.layout({ ...baseLayout, style: "text", fontSizePx: cssPx(12), paddingEm: 0 }, true).snapshot;
    const inline = generateCoreOfficeExport({
      owner: formula,
      layout: baseLayout,
      objectMode: "nativeOle",
      displayMode: "inline",
      fontSizePt: 12,
      formulaLetterFont: "katex",
      formulaChineseFont: "songti",
    });
    close(inline.width, zero.width + 2, "inline width padding");
    close(inline.height, zero.height + 2, "inline height padding");
    close(inline.baseline!, zero.baseline + 1, "inline baseline padding");
    assert.equal(decodeBase64(inline.svgBase64), inline.svg);
    assert.match(inline.svg, /fill-opacity="0\.001"/u);
    assert.doesNotMatch(inline.svg, /MathJax|MathLive|<text\b|<image\b|<script\b/iu);
    assert.doesNotMatch(inline.mathMl!, /data-visualtex-core-grid-id|data-mtef-tabs/u);

    const blockZero = formula.layout({ ...baseLayout, style: "display", fontSizePx: cssPx(12), paddingEm: 0 }, true).snapshot;
    const block = generateCoreOfficeExport({
      owner: formula,
      layout: baseLayout,
      objectMode: "crossPlatformPicture",
      displayMode: "block",
      fontSizePt: 12,
      formulaLetterFont: "katex",
      formulaChineseFont: "songti",
    });
    close(block.width, blockZero.width + 20, "block width padding");
    close(block.height, blockZero.height + 20, "block height padding");
    close(block.baseline!, blockZero.baseline + 10, "block baseline padding");
    formula.destroy();
  }

  {
    const formula = owner(String.raw`\begin{aligned}a&=b&c&=d\\long&=z&u&=v\end{aligned}`);
    const zero = formula.layout({ ...baseLayout, style: "display", fontSizePx: cssPx(10.5), paddingEm: 0 }, true).snapshot;
    const result = generateCoreOfficeExport({
      owner: formula,
      layout: baseLayout,
      objectMode: "mathTypeOle",
      displayMode: "inline",
      fontSizePt: 10.5,
      formulaLetterFont: "times",
      formulaChineseFont: "songti",
    });
    close(result.width, zero.width, "MathType horizontal padding");
    close(result.height, zero.height + 10.5, "MathType vertical padding");
    close(result.baseline!, zero.baseline + 5.25, "MathType baseline padding");
    assert.match(result.mathMl!, /data-mtef-tabs="true"/u);
    assert.doesNotMatch(result.mathMl!, /data-visualtex-core-grid-id/u);
    const ruler = /data-visualtex-mtef-ruler-stops="([0-9,]+)"/u.exec(result.mathMl!)?.[1]
      ?.split(",").map(Number);
    assert.ok(ruler && ruler.length === 2, result.mathMl);
    const columns = zero.gridColumns!.slice().sort((a,b) => a.index-b.index);
    const origin = columns[0].x;
    const expected = [columns[1], columns[3]].map(column =>
      Math.round((column.x-origin)*zero.fontSizePx*24));
    assert.deepEqual(ruler, expected);
    assert.ok(ruler[0] > 0 && ruler[1] > ruler[0] && ruler[1] <= 0xffff);

    const presentation = formula.presentationMathMl(true);
    assert.equal(
      annotateCoreMathTypeAlignmentGeometry(presentation.mathMl, presentation.alignedGridIds, zero),
      result.mathMl,
    );
    formula.destroy();
  }

  {
    const formula = owner(String.raw`\sqrt{x^2+y^2}`);
    let layoutCalls = 0;
    const mathMlOwner = {
      presentationMathMl: (displayMode: boolean) => formula.presentationMathMl(displayMode),
      layout: () => { layoutCalls += 1; throw new Error("OMML fast path must not layout"); },
    };
    const result = generateCoreOfficeExport({
      owner: mathMlOwner,
      layout: baseLayout,
      objectMode: "wordOmml",
      displayMode: "block",
      fontSizePt: 13.25,
      formulaLetterFont: "cambria",
      formulaChineseFont: "songti",
      preservedGeometry: { width: 321.5, height: 87.25, baseline: 61.75 },
    });
    assert.equal(layoutCalls, 0);
    assert.equal(result.svg, "");
    assert.equal(result.svgBase64, "");
    assert.equal(result.width, 321.5);
    assert.equal(result.height, 87.25);
    assert.equal(result.baseline, 61.75);
    assert.match(result.mathMl!, /^<math\b/u);
    formula.destroy();
  }

  const simpleSymbol = {
    id: "p11-simple", command: "pelevensimple", name: "P11 Simple", role: "ordinary",
    limitsBehavior: "auto", metrics: { widthEm: .8, ascentEm: .7, descentEm: .1 },
    artwork: { shapes: [{ kind: "rect", x: 100, y: 100, width: 600, height: 500, fill: true }] },
    ommlFallback: String.raw`\star`, createdAt: 1, updatedAt: 1,
  };
  {
    const formula = owner(String.raw`1+\pelevensimple`, [simpleSymbol]);
    const result = generateCoreOfficeExport({
      owner: formula,
      layout: baseLayout,
      objectMode: "nativeOle",
      displayMode: "inline",
      fontSizePt: 12,
      formulaLetterFont: "katex",
      formulaChineseFont: "songti",
    });
    assert.match(result.mathMl!, /[⋆★∗]/u);
    assert.doesNotMatch(result.mathMl!, /pelevensimple/u);
    assert.doesNotMatch(result.svg, /<mask\b|<clipPath\b|\bmask=|\bclip-path=/iu);
    const withPng = await rasterizeCoreOfficeExport(result, async input => {
      assert.equal(input.width, result.width);
      assert.equal(input.height, result.height);
      assert.equal(input.svgBase64, result.svgBase64);
      return "cG5n";
    });
    assert.equal(withPng.pngBase64, "cG5n");
    await assert.rejects(() => rasterizeCoreOfficeExport(result, async () => ""), /empty image/u);
    formula.destroy();
  }

  {
    const eraseSymbol = {
      ...simpleSymbol,
      id: "p11-erase",
      command: "pelevenerase",
      name: "P11 Erase",
      artwork: { shapes: [
        { kind: "rect", x: 100, y: 100, width: 600, height: 500, fill: true },
        { kind: "rect", x: 260, y: 220, width: 200, height: 140, fill: true, operation: "erase" },
      ] },
    };
    const formula = owner(String.raw`\pelevenerase`, [eraseSymbol]);
    assert.throws(() => generateCoreOfficeExport({
      owner: formula,
      layout: baseLayout,
      objectMode: "nativeOle",
      displayMode: "inline",
      fontSizePt: 12,
      formulaLetterFont: "katex",
      formulaChineseFont: "songti",
    }), /existing Office EMF consumer does not support/u);
    formula.destroy();
  }

  {
    const formula = owner("x");
    assert.throws(() => generateCoreOfficeExport({
      owner: formula,
      layout: baseLayout,
      objectMode: "nativeOle",
      displayMode: "inline",
      fontSizePt: Number.NaN,
      formulaLetterFont: "katex",
      formulaChineseFont: "songti",
    }), /semantic font size/u);
    formula.destroy();
  }

  assert.equal(runtime.activeSessions, 0);
  console.log(JSON.stringify({
    status: "PASS",
    actualWasmExecution: true,
    owners,
    fontSha256,
    contracts: [
      "OfficeExportResult field names unchanged",
      "inline 1px and block 10px natural padding",
      "MathType x=0/y=0.5pt vertical padding",
      "aligned ruler stops from core gridColumns at 24 MTEF units per CSS px",
      "wordOmml MathML-only fast path performs zero layout calls",
      "custom symbol Office semantic fallback",
      "unsupported Office vector mask/clip features fail closed",
      "PNG rasterization remains an injected post-process",
    ],
    officeSessionCalls: 0,
    wordCalls: 0,
  }, null, 2));
} finally {
  try { runtime.free(); } catch {}
}
