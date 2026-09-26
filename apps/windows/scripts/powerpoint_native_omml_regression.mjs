import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

async function source(path) {
  return (await readFile(path, "utf8")).replace(/\r\n?/g, "\n");
}

const addin = await source(
  "src-windows/VisualTeX.PowerPointVsto/ThisAddIn.cs",
);
const localization = await source(
  "src-windows/VisualTeX.PowerPointVsto/PowerPointRibbonLocalization.cs",
);
const service = await source(
  "src-windows/VisualTeX.PowerPointVsto/PowerPointFormulaService.cs",
);
const kernel = await source(
  "src-windows/VisualTeX.PowerPointVsto/PowerPointNativeOmmlKernel.cs",
);

for (const callback of [
  "OnNewSvgFormula",
  "OnNewOmmlFormula",
  "OnNewOleFormula",
  "OnInsertInlineOmml",
  "OnInsertDisplayOmml",
]) {
  assert.ok(addin.includes(callback), `Missing callback ${callback}`);
}

for (const ribbonId of [
  "VisualTeX.PowerPointVsto.New",
  "VisualTeX.PowerPointVsto.NewOle",
  "VisualTeX.PowerPointVsto.InsertInlineOmml",
  "VisualTeX.PowerPointVsto.InsertDisplayOmml",
]) {
  assert.ok(addin.includes(`id="${ribbonId}"`), `Chinese Ribbon missing ${ribbonId}`);
  assert.ok(localization.includes(`id="${ribbonId}"`), `English Ribbon missing ${ribbonId}`);
}

assert.ok(
  addin.includes('BeginSession("create", "crossPlatformPicture", null)'),
  "New SVG does not use picture mode",
);
assert.ok(
  addin.includes("public void OnNewOmmlFormula(object control) =>\n        BeginContextAwareOmmlSession();"),
  "Legacy New OMML callback must remain COM-compatible",
);
assert.ok(
  !addin.includes('id="VisualTeX.PowerPointVsto.NewOmml"')
    && !localization.includes('id="VisualTeX.PowerPointVsto.NewOmml"'),
  "PowerPoint Ribbon must expose exactly the two explicit inline/display OMML insert buttons",
);
assert.ok(
  addin.includes('label="插入行内 OMML 公式"')
    && addin.includes('label="插入行间 OMML 公式"')
    && localization.includes('label="Insert Inline OMML"')
    && localization.includes('label="Insert Display OMML"'),
  "Explicit inline/display OMML Ribbon labels are missing",
);
assert.ok(
  addin.includes("ResolveOmmlTextInsertionTarget(service)"),
  "Context-aware New OMML does not resolve the current text caret",
);
assert.ok(
  addin.includes("var liveTarget = service.TryCaptureTextInsertionTarget();")
    && addin.includes("return _lastTextInsertionTarget;"),
  "Ribbon ppSelectionNone fallback does not preserve the last text insertion target",
);
assert.ok(
  addin.includes("selection.Type == PpSelectionType.ppSelectionText")
    && addin.includes("selection.Type != PpSelectionType.ppSelectionNone")
    && addin.includes("_lastTextInsertionTarget = null;"),
  "Text insertion target cache is not updated/invalidated with PowerPoint selection state",
);
assert.ok(
  addin.includes('requestedDisplayMode: "inline"')
    && addin.includes('BeginSession(\n                "create",\n                "wordOmml",\n                target,'),
  "Context-aware New OMML does not insert inline into the captured text box",
);
assert.ok(
  addin.includes('BeginSession(\n                "create",\n                "wordOmml",\n                null,\n                requestedDisplayMode: "block")'),
  "Context-aware New OMML does not retain the standalone fallback",
);
assert.ok(
  addin.includes('BeginSession("create", "nativeOle", null)'),
  "New OLE does not use native OLE mode",
);
assert.ok(addin.includes('BeginOmmlTextInsertion("inline")'));
assert.ok(addin.includes('BeginOmmlTextInsertion("block")'));
assert.ok(
  addin.includes('target is null\n                && !string.Equals(\n                    displayMode,\n                    "block"'),
  "Display OMML button does not retain standalone-slide fallback while inline still requires a text caret",
);

assert.ok(
  service.includes('TextInsertionReferencePrefix = "visualtex-ppt-vsto-text:"'),
  "Text insertion target reference is missing",
);
assert.ok(service.includes("CaptureTextInsertionTarget()"));
assert.ok(
  service.includes("TryCaptureTextInsertionTarget()")
    && service.includes("CaptureTextInsertionTargetCore(requireTextSelection: false)"),
  "PowerPoint service has no non-throwing text-caret capture for context-aware New OMML",
);
assert.ok(
  service.includes("PowerPointNativeOmmlKernel.InsertIntoText("),
  "Text insertion does not use the native OMML kernel",
);
assert.ok(
  service.includes("mathMl!,\n                    session.DisplayMode);")
    && !service.includes("session.DisplayMode,\n                    FormulaFontSize.ResolveSemanticFontSize(metadata)"),
  "Text-box OMML must not receive VisualTeX persistent/session formula size",
);
assert.ok(
  kernel.includes("var contextFontSize = ResolveTextContextFontSize(")
    && kernel.includes("surrounding PowerPoint body text, not VisualTeX's persistent"),
  "Text-box OMML does not resolve font size from the PowerPoint text context",
);
assert.ok(
  service.includes("PowerPointNativeOmmlKernel.CreateStandalone("),
  "Standalone/converted OMML does not use the native OMML kernel",
);
assert.ok(
  addin.includes("PowerPointFormulaService.IsTextInsertionReference(selection.ObjectId)"),
  "Captured text target is not preserved through the editor session",
);

const insertStart = service.indexOf(
  "public OfficeObjectResult InsertOmml(OfficeSessionDocument session)",
);
const replaceStart = service.indexOf(
  "public OfficeObjectResult ReplaceOmml(OfficeSessionDocument session)",
);
const replaceOleStart = service.indexOf("public OfficeObjectResult ReplaceOle(");
assert.ok(insertStart >= 0 && replaceStart > insertStart && replaceOleStart > replaceStart);
const insertOmml = service.slice(insertStart, replaceStart);
const replaceOmml = service.slice(replaceStart, replaceOleStart);

assert.ok(
  !insertOmml.includes("Configure(shape, metadata)"),
  "New PowerPoint OMML still persists VisualTeX metadata",
);
assert.ok(
  !replaceOmml.includes("Configure(replacement, metadata)"),
  "Converted PowerPoint OMML still persists VisualTeX metadata",
);
assert.ok(
  !service.includes("PowerPointOmmlBridge.IsNativeEquation(")
    && !service.includes("PowerPointOmmlBridge.IsStandaloneNativeEquation("),
  "PowerPoint service still uses mutating MathZones-based native-equation detection",
);
assert.ok(
  !service.includes(".MathZones("),
  "PowerPoint service must not call MathZones as a detection API",
);
assert.ok(
  service.includes('GetEnabledMso(\n                       "EquationProfessional")')
    || service.includes('"EquationProfessional"'),
  "Native equation selection is not detected through PowerPoint command state",
);
assert.ok(service.includes('"EquationLinearFormat"'));

for (const marker of [
  "TransformMathMlToOmml",
  "ReplaceWithPreparedOmmlDirect",
  "math.Linearize()",
  '"EquationInsertNew"',
  '"EquationProfessional"',
  "ActivateShapeView(application, shape)",
  "shape.Select(MsoTriState.msoTrue)",
  "WordHostSync",
  "EnsureWordHost()",
  "ResetWordHost()",
  "GetTextSelectionWithRetry",
  "needsLeadingBreak",
  "needsTrailingBreak",
]) {
  assert.ok(kernel.includes(marker), `Native OMML kernel missing ${marker}`);
}

for (const forbidden of [
  "VisualTeXFormulaId",
  "VisualTeXMetadata",
  "VisualTeXIdentityOwner",
  "Tags.Add(",
  "AlternativeText =",
  "OleSetClipboard",
  "SetClipboardDataObject",
  '"MathML Presentation"',
]) {
  assert.ok(
    !kernel.includes(forbidden),
    `Native OMML kernel must not contain/persist ${forbidden}`,
  );
}

assert.ok(
  kernel.includes("presentation = slide.Parent as Presentation"),
  "Kernel does not recover the owning presentation after Word automation",
);
assert.ok(
  kernel.includes("sourceRange2 = GetCharacterRange(")
    && kernel.includes("((dynamic)sourceFont2).Size =")
    && !kernel.includes("activeRange =\n                    selection.TextRange2"),
  "Formula font size must be applied to the exact pre-equation source range, not PowerPoint's post-Professional selection",
);
assert.ok(
  !kernel.includes("Thread.Sleep(120)")
    && !kernel.includes("Thread.Sleep(80)")
    && !kernel.includes("Thread.Sleep(25)"),
  "Native OMML hot path still contains unconditional fixed PowerPoint waits",
);
assert.ok(
  service.includes("public sealed class PowerPointFormulaService : IDisposable")
    && service.includes("PowerPointNativeOmmlKernel.Shutdown()"),
  "Private Word host has no deterministic PowerPoint service shutdown path",
);
assert.ok(
  kernel.includes("window.Activate()") && kernel.includes("view.GotoSlide(slide.SlideIndex)"),
  "Kernel does not restore the target PowerPoint view before selecting text",
);

console.log("PowerPoint native OMML regression passed.");
