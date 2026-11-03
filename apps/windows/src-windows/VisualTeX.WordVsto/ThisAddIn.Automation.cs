using System.Text.Json;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Task = System.Threading.Tasks.Task;

namespace VisualTeX.WordVsto;

// Agent writes (docs/office-mcp/DESIGN.md §5): each one places the target in the
// document and then runs the same command as the Ribbon, without its dialogs.
public sealed partial class ThisAddIn
{
    internal sealed class AutomationSessionInput
    {
        internal IReadOnlyList<string>? Lines { get; set; }
        internal bool? Numbered { get; set; }
        internal double? FontSizePt { get; set; }
    }

    internal static readonly string[] AutomationWriteOperations =
    {
        "insert_formula", "update_formula", "delete_formula", "select_formula", "set_formula_font_size",
        "convert_formulas", "latex_text_to_formulas", "formulas_to_latex_text", "insert_rich_text",
        "insert_equation_reference", "number_display_formulas", "update_equation_numbers",
        "set_equation_number_format",
    };

    internal Task<object?>? ExecuteAutomationWrite(string operation, JsonElement arguments) =>
        Array.IndexOf(AutomationWriteOperations, operation) < 0
            ? null
            : Task.Run(() => RunAutomationWriteAsync(operation, arguments));

    private async Task<object?> RunAutomationWriteAsync(string operation, JsonElement arguments)
    {
        var dispatcher = _dispatcher ?? throw new OfficeAutomationException("internal", "The Word add-in is not ready.");
        var service = _formulaService ?? throw new OfficeAutomationException("internal", "The Word add-in is not ready.");
        var lifetime = _lifetime ?? throw new OfficeAutomationException("internal", "The Word add-in is not ready.");
        if (_operationGate.CurrentCount == 0)
            throw new OfficeAutomationException(
                "host_busy",
                "VisualTeX is busy with another Word operation (the formula editor may be open). Try again when it is finished.");
        var key = OfficeAutomationText.String(arguments, "document");
        string Key() => key is { Length: > 0 }
            ? key!
            : throw new OfficeAutomationException("invalid_arguments", "document_id is required.");
        string Scope() => OfficeAutomationText.String(arguments, "scope") switch
        {
            null or "" or "document" => "document",
            "selection" => "selection",
            var other => throw new OfficeAutomationException("invalid_arguments", $"Unknown scope {other}: use document or selection."),
        };
        string RequiredMode(string name) =>
            OfficeAutomationText.ObjectMode(OfficeAutomationText.RequiredString(arguments, name))
            ?? throw new OfficeAutomationException("invalid_arguments", $"{name} is required.");

        using var run = OfficeAutomationRun.Begin();
        switch (operation)
        {
            case "select_formula":
                return await dispatcher.InvokeAsync(() =>
                    service.AutomationSelectFormula(Key(), OfficeAutomationText.RequiredString(arguments, "formula_id")))
                    .ConfigureAwait(false);

            case "insert_formula":
            {
                var lines = OfficeAutomationText.Lines(arguments)
                    ?? throw new OfficeAutomationException("invalid_arguments", "latex is required.");
                var display = OfficeAutomationText.String(arguments, "display") switch
                {
                    null or "" or "inline" => "inline",
                    "block" => "block",
                    var other => throw new OfficeAutomationException("invalid_arguments", $"Unknown display {other}: use inline or block."),
                };
                var mode = OfficeAutomationText.ObjectMode(OfficeAutomationText.String(arguments, "format"));
                if (mode == FormulaOleContract.CrossPlatformPictureMode)
                    throw new OfficeAutomationException("unsupported_in_host", "Pictures are a PowerPoint format; use visualtex_ole, office_equation or mathtype in Word.");
                var location = OfficeAutomationText.Element(arguments, "location");
                var selection = await dispatcher.InvokeAsync(() => service.AutomationPlaceInsertion(key, location))
                    .ConfigureAwait(false);
                await RunSessionAsync("create", display, mode, selection, conversionOnly: true, lifetime.Token,
                    new AutomationSessionInput
                    {
                        Lines = lines,
                        Numbered = OfficeAutomationText.Bool(arguments, "numbered"),
                        FontSizePt = OfficeAutomationText.Number(arguments, "font_size_pt"),
                    }).ConfigureAwait(false);
                return await FinishWithWrittenAsync(run, dispatcher, service).ConfigureAwait(false);
            }

            case "update_formula":
            {
                var lines = OfficeAutomationText.Lines(arguments);
                var mode = OfficeAutomationText.ObjectMode(OfficeAutomationText.String(arguments, "format"));
                // the editor's number switch: a display formula numbered or not
                var numbered = OfficeAutomationText.Bool(arguments, "numbered");
                if (lines is null && mode is null && numbered is null)
                    throw new OfficeAutomationException("invalid_arguments", "Give latex (new content), format (convert) and/or numbered.");
                if (mode == FormulaOleContract.CrossPlatformPictureMode)
                    throw new OfficeAutomationException("unsupported_in_host", "Pictures are a PowerPoint format.");
                var selection = await dispatcher.InvokeAsync(() =>
                {
                    service.AutomationSelectFormula(Key(), OfficeAutomationText.RequiredString(arguments, "formula_id"));
                    return service.ReadSelection();
                }).ConfigureAwait(false);
                if (selection.Metadata is null)
                    throw new OfficeAutomationException("formula_not_found", "The formula could not be read for editing.");
                var sourceMode = selection.ObjectMode;
                var converts = mode is not null && sourceMode is not null && mode != sourceMode;
                if (converts)
                {
                    // the format first, through the Ribbon's conversion of the selected formula;
                    // the converted formula stays selected
                    await ConvertFormulaFormatAsync(false, sourceMode!, mode!).ConfigureAwait(false);
                    if (lines is not null || numbered is not null)
                    {
                        run.ThrowIfFailed();
                        selection = await dispatcher.InvokeAsync(() => service.ReadSelection()).ConfigureAwait(false);
                        if (selection.Metadata is null)
                            throw new OfficeAutomationException("operation_failed", "The formula was converted but could not be found again to change its LaTeX.");
                    }
                }
                if (lines is not null || numbered is not null)
                {
                    await RunSessionAsync("edit", null, null, selection, conversionOnly: true, lifetime.Token,
                        new AutomationSessionInput
                        {
                            Lines = lines,
                            Numbered = numbered,
                            FontSizePt = OfficeAutomationText.Number(arguments, "font_size_pt"),
                        }).ConfigureAwait(false);
                    return await FinishWithWrittenAsync(run, dispatcher, service).ConfigureAwait(false);
                }
                var converted = await dispatcher.InvokeAsync(() => service.AutomationDescribeSelected()).ConfigureAwait(false);
                var result = run.Finish();
                if (converted is not null) result["formula"] = converted;
                return result;
            }

            case "delete_formula":
                await dispatcher.InvokeAsync(() =>
                    service.AutomationSelectFormula(Key(), OfficeAutomationText.RequiredString(arguments, "formula_id")))
                    .ConfigureAwait(false);
                await DeleteSelectedAsync().ConfigureAwait(false);
                return run.Finish();

            case "set_formula_font_size":
            {
                var size = OfficeAutomationText.Number(arguments, "font_size_pt")
                    ?? throw new OfficeAutomationException("invalid_arguments", "font_size_pt is required.");
                if (OfficeAutomationText.String(arguments, "kind") is { Length: > 0 } kind)
                {
                    // every inline or every display formula of the document, as the Ribbon's 行内 / 行间 do
                    var display = kind switch
                    {
                        "inline" => false,
                        "display" => true,
                        _ => throw new OfficeAutomationException("invalid_arguments", $"Unknown kind {kind}: use inline or display."),
                    };
                    BeginFormulaFormatMutation();
                    try
                    {
                        var batch = await dispatcher.InvokeAsync(() =>
                        {
                            service.AutomationActivate(key);
                            return service.ApplyFontSizeToDocumentFormulas(display, size);
                        }).ConfigureAwait(false);
                        run.Status(T(
                            $"已把 {batch.Resized} 个{(display ? "行间" : "行内")}公式改为 {FormulaFontSize.Describe(size)}。",
                            $"{batch.Resized} {(display ? "display" : "inline")} formula(s) set to {FormulaFontSize.Describe(size, english: true)}."));
                        foreach (var failure in batch.Failures) run.Warn(failure);
                    }
                    finally { EndFormulaFormatMutation(); }
                    return run.Finish();
                }
                await dispatcher.InvokeAsync(() =>
                {
                    service.AutomationSelectFormula(Key(), OfficeAutomationText.RequiredString(arguments, "formula_id"));
                    ApplyFormulaFontSize(size);
                    return true;
                }).ConfigureAwait(false);
                return run.Finish();
            }

            case "convert_formulas":
            {
                var source = RequiredMode("from");
                var target = RequiredMode("to");
                await dispatcher.InvokeAsync(() => { service.AutomationActivate(key); return true; }).ConfigureAwait(false);
                await ConvertFormulaFormatAsync(Scope() == "document", source, target).ConfigureAwait(false);
                return run.Finish();
            }

            case "latex_text_to_formulas":
            {
                var target = RequiredMode("format");
                run.NumberDisplayFormulas = OfficeAutomationText.Bool(arguments, "number_display_formulas") ?? false;
                await dispatcher.InvokeAsync(() => { service.AutomationActivate(key); return true; }).ConfigureAwait(false);
                await RedrawLatexAsync(Scope() == "document", target).ConfigureAwait(false);
                return run.Finish();
            }

            case "formulas_to_latex_text":
            {
                var source = RequiredMode("from");
                await dispatcher.InvokeAsync(() => { service.AutomationActivate(key); return true; }).ConfigureAwait(false);
                await ConvertFormulaObjectsToLatexAsync(Scope() == "document", source).ConfigureAwait(false);
                return run.Finish();
            }

            case "insert_rich_text":
            {
                run.BulkSource = OfficeAutomationText.RequiredString(arguments, "text");
                run.BulkSourceFormat = OfficeAutomationText.String(arguments, "source_format") ?? "auto";
                run.BulkObjectMode = OfficeAutomationText.ObjectMode(OfficeAutomationText.String(arguments, "format"));
                run.NumberDisplayFormulas = OfficeAutomationText.Bool(arguments, "number_display_formulas") ?? false;
                var location = OfficeAutomationText.Element(arguments, "location");
                await dispatcher.InvokeAsync(() => service.AutomationPlaceInsertion(key, location)).ConfigureAwait(false);
                await BulkImportAsync().ConfigureAwait(false);
                return run.Finish();
            }

            case "insert_equation_reference":
            {
                var number = OfficeAutomationText.RequiredString(arguments, "number").Trim().Trim('(', ')', ' ');
                var source = OfficeAutomationText.String(arguments, "source");
                var location = OfficeAutomationText.Element(arguments, "location");
                await dispatcher.InvokeAsync(() => service.AutomationPlaceInsertion(key, location)).ConfigureAwait(false);
                var inserted = await dispatcher.InvokeAsync(() => InsertAutomationEquationReference(service, number, source))
                    .ConfigureAwait(false);
                run.Status(inserted);
                return run.Finish();
            }

            case "number_display_formulas":
                await dispatcher.InvokeAsync(() => { service.AutomationActivate(key); return true; }).ConfigureAwait(false);
                await NumberDisplayFormulasAsync(Scope() == "document").ConfigureAwait(false);
                return run.Finish();

            case "update_equation_numbers":
                await dispatcher.InvokeAsync(() => { service.AutomationActivate(key); return true; }).ConfigureAwait(false);
                await UpdateEquationNumbersAsync().ConfigureAwait(false);
                return run.Finish();

            case "set_equation_number_format":
            {
                var format = OfficeAutomationText.RequiredString(arguments, "format_id");
                await dispatcher.InvokeAsync(() => { service.AutomationActivate(key); return true; }).ConfigureAwait(false);
                await SetEquationNumberFormatAsync(format).ConfigureAwait(false);
                return run.Finish();
            }
        }
        throw OfficeAutomationText.UnknownOperation(operation);
    }

    private static async Task<object?> FinishWithWrittenAsync(
        OfficeAutomationRun run,
        OfficeUiDispatcher dispatcher,
        WordFormulaService service)
    {
        var written = run.Result as OfficeObjectResult;
        object? formula = null;
        if (written is not null)
            formula = await dispatcher.InvokeAsync(() => service.AutomationDescribeWritten(written)).ConfigureAwait(false);
        var result = run.Finish();
        if (formula is not null) result["formula"] = formula;
        return result;
    }

    /// The Ribbon's "Insert equation reference" with the target named by its number
    /// instead of picked in the dialog.
    private string InsertAutomationEquationReference(WordFormulaService service, string number, string? source)
    {
        var application = _application ?? throw new OfficeAutomationException("internal", "Word is not available.");
        Document? document = null;
        Selection? selection = null;
        try
        {
            document = application.ActiveDocument;
            selection = application.Selection;
            var start = selection.Start;
            var end = selection.End;
            var color = selection.Font.Color;
            var targets = service.GetCanonicalEquationReferenceTargets(document)
                .Concat(MathTypeEquationReferences.GetTargets(document))
                .Where(target => source switch
                {
                    null or "" => true,
                    "visualtex_ole" => target.Source == EquationReferenceSource.VisualTeX,
                    "office_equation" => target.Source == EquationReferenceSource.WordOmml,
                    "mathtype" => target.Source == EquationReferenceSource.MathType,
                    _ => throw new OfficeAutomationException("invalid_arguments", $"Unknown source {source}."),
                })
                .ToList();
            static string Bare(string? text) => (text ?? string.Empty).Trim().Trim('(', ')', ' ');
            var matches = targets.Where(target => Bare(target.NumberText) == number).ToList();
            if (matches.Count != 1)
            {
                var available = string.Join(", ", targets.Select(target => $"({Bare(target.NumberText)}) {target.Source}"));
                throw new OfficeAutomationException(
                    "formula_not_found",
                    matches.Count == 0
                        ? $"No numbered equation ({number}) in this document. Numbered equations: {(available.Length == 0 ? "none" : available)}."
                        : $"Several numbered equations are ({number}); give source to choose one. Numbered equations: {available}.");
            }
            var chosen = matches[0];
            selection.SetRange(start, end);
            if (chosen.Source == EquationReferenceSource.MathType)
                MathTypeEquationReferences.InsertReference(document, selection, chosen, color);
            else
                service.InsertEquationReference(document, selection, chosen, EquationReferenceStyle.Parenthesized, color);
            return $"Inserted a reference to equation ({number}).";
        }
        finally
        {
            ReleaseComObject(selection);
            ReleaseComObject(document);
        }
    }
}
