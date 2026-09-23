import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { generateCoreOfficeExport } from "../src/office/compat/coreOfficeCompatibility.ts";

const require = createRequire(import.meta.url);
const { EngineRuntime } = require(resolve("../../engine/web/pkg-node/visualtex_engine_wasm.js"));
const fontBytes = readFileSync(resolve("src-windows/artifacts/office-fonts/latinmodern-math.otf"));
const fontSha256 = createHash("sha256").update(fontBytes).digest("hex");
const runtime = new EngineRuntime();
const registered = JSON.parse(runtime.registerFont(fontBytes, 0));
assert.equal(registered.status, "ok");
const fontId = registered.result.font.id as string;
const layout = {
  mathFont: fontId,
  textFont: null,
  mathFallbacks: [],
  customSymbolTextFonts: [],
  style: "display" as const,
  fontSizePx: 16,
  paddingEm: 0,
};
let next = 1;
class Owner {
  readonly handle: number;
  revision = 0;
  constructor(source: string) {
    const response = this.call({ op: "create", documentId: `p11-consumer-${next++}`, source }, "session");
    this.handle = response.handle;
    this.revision = response.snapshot.revision;
  }
  private call(command: unknown, kind: string) {
    const response = JSON.parse(runtime.dispatch(JSON.stringify({ protocol: 1, command })));
    if (response.status !== "ok" || response.result.kind !== kind) throw new Error(JSON.stringify(response));
    return response.result;
  }
  layout(options: typeof layout, includeSvg: boolean) {
    const result = this.call({ op: "layout", handle: this.handle, expectedRevision: this.revision, options, includeSvg }, "layout");
    return { sceneToken: result.sceneToken, snapshot: result.snapshot, svg: result.svg };
  }
  presentationMathMl(displayMode: boolean) {
    return this.call({ op: "exportPresentationMathMl", handle: this.handle, expectedRevision: this.revision, displayMode }, "presentationMathMl").result;
  }
  destroy() { this.call({ op: "destroy", handle: this.handle, expectedRevision: this.revision }, "destroyed"); }
}

try {
  const vectorOwner = new Owner(String.raw`\frac{x_1+1}{y^2}+\sqrt{z}`);
  const vector = generateCoreOfficeExport({
    owner: vectorOwner,
    layout,
    objectMode: "nativeOle",
    displayMode: "inline",
    fontSizePt: 12,
    formulaLetterFont: "katex",
    formulaChineseFont: "songti",
  });
  vectorOwner.destroy();

  const mathTypeOwner = new Owner(String.raw`\begin{aligned}a&=b&c&=d\\long&=z&u&=v\end{aligned}`);
  const mathType = generateCoreOfficeExport({
    owner: mathTypeOwner,
    layout,
    objectMode: "mathTypeOle",
    displayMode: "block",
    fontSizePt: 10.5,
    formulaLetterFont: "times",
    formulaChineseFont: "songti",
  });
  mathTypeOwner.destroy();
  const ruler = /data-visualtex-mtef-ruler-stops="([0-9,]+)"/u.exec(mathType.mathMl ?? "")?.[1];
  if (!ruler) throw new Error("MathType fixture has no core ruler geometry");

  assert.equal(runtime.activeSessions, 0);
  process.stdout.write(JSON.stringify({
    schema: "visualtex-p11-office-consumer-fixture-v1",
    fontSha256,
    vector: {
      svg: vector.svg,
      width: vector.width,
      height: vector.height,
      baseline: vector.baseline,
      mathMl: vector.mathMl,
    },
    mathType: {
      mathMl: mathType.mathMl,
      width: mathType.width,
      height: mathType.height,
      baseline: mathType.baseline,
      fontSizePt: 10.5,
      rulerStops: ruler,
    },
  }));
} finally {
  try { runtime.free(); } catch {}
}
