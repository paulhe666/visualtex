import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const read = async (name) => (await readFile(
  new URL("../src-windows/VisualTeX.PowerPointVsto/" + name, import.meta.url), "utf8",
)).replace(/\r\n?/g, "\n");
const [service, addin, bridge, native, kernel] = await Promise.all([
  read("PowerPointFormulaService.cs"), read("ThisAddIn.cs"),
  read("PowerPointOmmlBridge.cs"), read("PowerPointFormulaService.NativeOmml.cs"),
  read("PowerPointNativeOmmlKernel.cs"),
]);
assert.ok(!service.includes("PowerPointOmmlBridge.IsNativeEquation("),
  "Passive selection must not use MathZones detection");
assert.ok(!service.includes(".MathZones("),
  "Service selection must not use MathZones directly");
assert.ok(service.includes("objectId = NativeOmmlReference(slide, shape);"),
  "Passive selection must capture the exact candidate identity");
assert.ok(native.includes("PowerPointNativeOmmlPackage.ReadStandalone(presentation, slideId, shapeId)"),
  "Explicit native source must come from the exact saved shape");
assert.ok(bridge.includes("MathZones(1, 1)") && bridge.includes(".Count) != 1"),
  "Explicit native font writes must use one indexed math zone, not its collection span");
const selectedEntry = addin.slice(addin.indexOf("private void BeginSelectedSession("), addin.indexOf("private void BeginSession("));
assert.ok(!selectedEntry.includes("LoadNativeOmmlSelection(") && !selectedEntry.includes("LoadMathTypeSelection("),
  "Repeated Ribbon clicks must not overwrite an active source before the operation gate");
const resolver = addin.slice(addin.indexOf("private OfficeSelection ResolveFormulaSelection("), addin.indexOf("private void ReportError("));
assert.ok(!resolver.includes("_lastFormulaSelection") && resolver.includes("return service.ReadSelection();"),
  "An empty/prose current selection must not edit a previously selected object");
assert.ok(kernel.includes("false, // Do not promote the private Word inline scratch range."),
  "PowerPoint must supply the sixth direct-OMML argument explicitly");
assert.ok(kernel.includes("PowerPointNativeOmmlFormatting.ForWordReverseTransform(omml)"),
  "Native DrawingML colors must be adapted at the PowerPoint reverse boundary");
assert.ok(!kernel.includes("return linearText.Normalize(") && !kernel.includes("Text = text.Normalize("),
  "UnicodeMath mathematical alphabet scalars must not be compatibility-normalized away");
assert.ok(kernel.indexOf("foreach (var color in linear.Colors)") < kernel.indexOf("SelectTextRangeForEquation("),
  "Retained colors must be applied only to the exact inserted linear range before professional conversion");
console.log("PowerPoint basic source contracts passed (passive detection, exact source, indexed font range, gate ownership, selection authority, six-argument call, native color boundary)");
