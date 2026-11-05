using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.PowerPointVsto;

// Agent writes (docs/office-mcp/DESIGN.md §5): each one places the target and then
// runs the same command as the Ribbon, without its dialogs.
public sealed partial class ThisAddIn
{
    internal sealed class AutomationSessionInput
    {
        internal IReadOnlyList<string>? Lines { get; set; }
        internal double? FontSizePt { get; set; }
    }

    internal static readonly string[] AutomationWriteOperations =
    {
        "insert_formula", "update_formula", "delete_formula", "select_formula", "set_formula_font_size",
        "convert_formulas", "latex_text_to_formulas", "formulas_to_latex_text",
    };

    internal Task<object?>? ExecuteAutomationWrite(string operation, JsonElement arguments) =>
        Array.IndexOf(AutomationWriteOperations, operation) < 0
            ? null
            : Task.Run(() => RunAutomationWriteAsync(operation, arguments));

    private async Task<object?> RunAutomationWriteAsync(string operation, JsonElement arguments)
    {
        var dispatcher = _dispatcher ?? throw new OfficeAutomationException("internal", "The PowerPoint add-in is not ready.");
        var service = _formulaService ?? throw new OfficeAutomationException("internal", "The PowerPoint add-in is not ready.");
        var lifetime = _lifetime ?? throw new OfficeAutomationException("internal", "The PowerPoint add-in is not ready.");
        if (_operationGate.CurrentCount == 0)
            throw new OfficeAutomationException(
                "host_busy",
                "VisualTeX is busy with another PowerPoint operation (the formula editor may be open). Try again when it is finished.");
        var key = OfficeAutomationText.String(arguments, "document");
        string Key() => key is { Length: > 0 }
            ? key!
            : throw new OfficeAutomationException("invalid_arguments", "document_id is required.");
        string FormulaId() => OfficeAutomationText.RequiredString(arguments, "formula_id");

        using var run = OfficeAutomationRun.Begin();
        switch (operation)
        {
            case "select_formula":
                return await dispatcher.InvokeAsync(() => service.AutomationSelectFormula(Key(), FormulaId()))
                    .ConfigureAwait(false);

            case "insert_formula":
            {
                var lines = OfficeAutomationText.Lines(arguments)
                    ?? throw new OfficeAutomationException("invalid_arguments", "latex is required.");
                var mode = OfficeAutomationText.ObjectMode(OfficeAutomationText.String(arguments, "format"))
                    ?? FormulaOleContract.CrossPlatformPictureMode;
                var display = OfficeAutomationText.String(arguments, "display");
                if (display is not (null or "" or "inline" or "block"))
                    throw new OfficeAutomationException("invalid_arguments", $"Unknown display {display}: use inline or block.");
                var location = OfficeAutomationText.Element(arguments, "location");
                var input = new AutomationSessionInput { Lines = lines, FontSizePt = OfficeAutomationText.Number(arguments, "font_size_pt") };
                var placed = await dispatcher.InvokeAsync(() => service.AutomationPlaceInsertion(key, location))
                    .ConfigureAwait(false);
                OfficeSelection? target = null;
                if (mode == FormulaOleContract.WordOmmlMode && placed.AtCursor)
                    target = await dispatcher.InvokeAsync(() => service.TryCaptureTextInsertionTarget()).ConfigureAwait(false);
                if (mode == FormulaOleContract.WordOmmlMode && target is null && display == "inline")
                    throw new OfficeAutomationException("invalid_arguments", "An inline native equation goes into text: put the text caret in a text box first, or use display block.");
                await RunSessionAsync("create", mode, target, conversionOnly: true,
                    mode == FormulaOleContract.WordOmmlMode ? (target is null ? "block" : display is null or "" ? "inline" : display) : display is null or "" ? null : display,
                    lifetime.Token, input).ConfigureAwait(false);
                run.ThrowIfFailed();
                if (target is not null)
                    return await FinishWithSelectionAsync(run, dispatcher, service).ConfigureAwait(false);
                var inserted = await dispatcher.InvokeAsync(() => service.AutomationDescribeInserted(key, placed)).ConfigureAwait(false);
                var result = run.Finish();
                if (inserted is not null) result["formula"] = inserted;
                return result;
            }

            case "update_formula":
            {
                var lines = OfficeAutomationText.Lines(arguments);
                var mode = OfficeAutomationText.ObjectMode(OfficeAutomationText.String(arguments, "format"));
                var position = PowerPointFormulaService.AutomationFrameChange.From(OfficeAutomationText.Element(arguments, "position"));
                if (lines is null && mode is null && position is null)
                    throw new OfficeAutomationException("invalid_arguments", "Give latex (new content), format (convert) and/or position (move or resize).");
                if (lines is null && mode is null)
                {
                    // only a new place / size: the formula stays as it is
                    var moved = await dispatcher.InvokeAsync(() => service.AutomationSetFrame(Key(), FormulaId(), position!))
                        .ConfigureAwait(false);
                    var result = run.Finish();
                    result["result"] = moved;
                    return result;
                }
                var selection = await dispatcher.InvokeAsync(() =>
                {
                    service.AutomationSelectFormula(Key(), FormulaId());
                    return ResolveFormulaSelection(service);
                }).ConfigureAwait(false);
                await RunSessionAsync("edit", mode, selection, conversionOnly: true, null, lifetime.Token,
                    new AutomationSessionInput { Lines = lines, FontSizePt = OfficeAutomationText.Number(arguments, "font_size_pt") })
                    .ConfigureAwait(false);
                if (position is not null)
                {
                    // PowerPoint leaves the edited formula selected: it takes the new place / size
                    run.ThrowIfFailed();
                    await dispatcher.InvokeAsync(() => service.AutomationSetFrame(Key(), null, position)).ConfigureAwait(false);
                }
                return await FinishWithSelectionAsync(run, dispatcher, service).ConfigureAwait(false);
            }

            case "delete_formula":
                await dispatcher.InvokeAsync(() => service.AutomationSelectFormula(Key(), FormulaId())).ConfigureAwait(false);
                await DeleteSelectedAsync().ConfigureAwait(false);
                return run.Finish();

            case "set_formula_font_size":
            {
                var size = OfficeAutomationText.Number(arguments, "font_size_pt")
                    ?? throw new OfficeAutomationException("invalid_arguments", "font_size_pt is required.");
                if (OfficeAutomationText.String(arguments, "kind") is { Length: > 0 } kind)
                {
                    // every inline or every display formula of the presentation, as the Ribbon's 行内 / 行间 do
                    var display = kind switch
                    {
                        "inline" => false,
                        "display" => true,
                        _ => throw new OfficeAutomationException("invalid_arguments", $"Unknown kind {kind}: use inline or display."),
                    };
                    var batch = await dispatcher.InvokeAsync(() =>
                    {
                        service.AutomationActivate(Key());
                        return service.ApplyFontSizeToPresentationFormulas(display, size);
                    }).ConfigureAwait(false);
                    run.Status(T(
                        $"已把 {batch.Resized} 个{(display ? "行间" : "行内")}公式改为 {FormulaFontSize.Describe(size)}。",
                        $"{batch.Resized} {(display ? "display" : "inline")} formula(s) set to {FormulaFontSize.Describe(size, english: true)}."));
                    if (batch.MathTypeLeft > 0)
                        run.Warn(T($"{batch.MathTypeLeft} 个 MathType 公式保持原字号。", $"{batch.MathTypeLeft} MathType equation(s) keep their size."));
                    foreach (var failure in batch.Failures) run.Warn(failure);
                    return run.Finish();
                }
                await dispatcher.InvokeAsync(() =>
                {
                    service.AutomationSelectFormula(Key(), FormulaId());
                    ApplyFormulaFontSize(size);
                    return true;
                }).ConfigureAwait(false);
                return run.Finish();
            }

            case "convert_formulas":
            case "latex_text_to_formulas":
            case "formulas_to_latex_text":
            {
                string Family(string name, bool allowAny)
                {
                    var value = OfficeAutomationText.RequiredString(arguments, name);
                    return value switch
                    {
                        "visualtex_ole" or "picture" => PowerPointConversionRequest.VisualTeXFamily,
                        "office_equation" => FormulaOleContract.WordOmmlMode,
                        "mathtype" => FormulaOleContract.MathTypeOleMode,
                        "any" when allowAny => PowerPointConversionRequest.AnySource,
                        _ => throw new OfficeAutomationException("invalid_arguments", $"Unknown {name} {value}."),
                    };
                }
                var source = operation switch
                {
                    "latex_text_to_formulas" => PowerPointConversionRequest.LatexFamily,
                    _ => Family("from", allowAny: true),
                };
                var target = operation switch
                {
                    "formulas_to_latex_text" => PowerPointConversionRequest.LatexFamily,
                    "latex_text_to_formulas" => Family("format", allowAny: false),
                    _ => Family("to", allowAny: false),
                };
                var scope = OfficeAutomationText.String(arguments, "scope") switch
                {
                    null or "" or "presentation" or "document" => "presentation",
                    "slide" => "slide",
                    "selection" => "selection",
                    var other => throw new OfficeAutomationException("invalid_arguments", $"Unknown scope {other}: use presentation, slide or selection."),
                };
                var visualTeXMode = (OfficeAutomationText.String(arguments, operation == "latex_text_to_formulas" ? "format" : "to")) == "picture"
                    ? FormulaOleContract.CrossPlatformPictureMode
                    : FormulaOleContract.NativeOleMode;
                if (!PowerPointConversionRequest.TryParse($"{source}|{target}|{scope}", visualTeXMode, out var request))
                    throw new OfficeAutomationException("invalid_arguments", "This conversion is not available (source and target must differ).");
                await dispatcher.InvokeAsync(() => { ReleaseComObject(service.AutomationActivate(key)); return true; }).ConfigureAwait(false);
                await ConvertFormulaFormatAsync(request, lifetime.Token).ConfigureAwait(false);
                var outcome = LastFormatConversionOutcome;
                return run.Finish(outcome is null ? null : new
                {
                    converted = outcome.Converted,
                    attempted = outcome.Attempted,
                    total = outcome.Total,
                    stopped = outcome.Stopped,
                });
            }
        }
        throw OfficeAutomationText.UnknownOperation(operation);
    }

    private static async Task<object?> FinishWithSelectionAsync(
        OfficeAutomationRun run,
        OfficeUiDispatcher dispatcher,
        PowerPointFormulaService service)
    {
        // PowerPoint leaves the formula it wrote selected
        object? selection = null;
        try { selection = await dispatcher.InvokeAsync(() => service.AutomationGetSelection(null)).ConfigureAwait(false); }
        catch { }
        var result = run.Finish();
        if (selection is not null) result["selection"] = selection;
        return result;
    }
}
