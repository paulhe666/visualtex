using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

// Update Equation Format's font step (WordFormulaService.FontRefresh.cs): the fonts the main app draws in now come
// from the companion, the formulas drawn in other fonts are drawn again from their LaTeX by the hidden converter (a
// render batch, as conversions use) and written back by the service.
public sealed partial class ThisAddIn
{
    /// <summary>
    /// The main app's current formula fonts. Fonts is null when they are not checked: VisualTeX cannot be reached
    /// (Reached false), or it runs without having opened its main window, which publishes them.
    /// </summary>
    private async Task<(FormulaFontSettings? Fonts, bool Reached)> TryGetFormulaFontSettingsAsync(
        VisualTeXSessionClient client,
        CancellationToken token)
    {
        try
        {
            await client.EnsureHealthyAsync(token).ConfigureAwait(false);
            var fonts = await client.GetFormulaFontSettingsAsync(token).ConfigureAwait(false);
            if (fonts is null) WordDoubleClickHook.TraceMessage("font-refresh-settings-unpublished");
            return (fonts, true);
        }
        catch (Exception error) when (error is not OperationCanceledException || !token.IsCancellationRequested)
        {
            WordDoubleClickHook.TraceMessage($"font-refresh-settings-unavailable error={error.Message}");
            return (null, false);
        }
    }

    /// <summary>Why the formula fonts were not checked, for Update Equation Format's messages.</summary>
    private static string FontsNotCheckedNote(bool reached) => reached
        ? T("（请先打开 VisualTeX 主窗口，本次未检查公式字体。）",
            "(Open the VisualTeX main window first; formula fonts were not checked this time.)")
        : T("（无法连接 VisualTeX，未检查公式字体。）",
            "(VisualTeX could not be reached, so formula fonts were not checked.)");

    /// <summary>
    /// Draws the formulas again in the current fonts, each at its own size and display mode, 20 to a batch, and
    /// prepares them for writing with their ids, numbers, lines and titles. A formula that cannot be drawn is
    /// reported and keeps its picture.
    /// </summary>
    private async Task<List<(WordFormulaService.FontRefreshTarget Target, PreparedWordBulkFormula Prepared)>> RenderFontRefreshAsync(
        VisualTeXSessionClient client,
        IReadOnlyList<WordFormulaService.FontRefreshTarget> targets,
        WordOperationProgress progress,
        List<string> failures,
        CancellationToken token)
    {
        var prepared = new List<(WordFormulaService.FontRefreshTarget, PreparedWordBulkFormula)>();
        await client.PrewarmConverterAsync(token).ConfigureAwait(false);
        for (var offset = 0; offset < targets.Count; offset += 20)
        {
            progress.Report(T("正在用当前字体重绘公式…", "Drawing equations in the current fonts…"), offset, targets.Count);
            var chunk = targets.Skip(offset).Take(20).ToList();
            var documentId = chunk[0].DocumentId;
            var runs = chunk.Select(target => new WordBulkRun
            {
                Id = target.FormulaId,
                IsFormula = true,
                Latex = target.Metadata.Latex,
                DisplayMode = target.DisplayMode,
            }).ToList();
            // no fonts named: the converter draws in the current ones
            var batch = await client.RenderBatchAsync(
                chunk.Select((target, index) => new RenderBatchItemRequest
                {
                    Id = runs[index].Id,
                    Request = CreateBulkFormulaConversionRequest(
                        runs[index], FormulaOleContract.NativeOleMode, documentId, target.FontSizePt,
                        rows: target.Metadata.Lines?.Select(line => line.Latex).ToList()),
                }).ToList(),
                TimeSpan.FromMinutes(3),
                token).ConfigureAwait(false);
            var materialized = MaterializeBulkFormulaTemplates(
                client,
                runs.Select((run, index) => (run, batch.Items[index].Session)).ToList(),
                FormulaOleContract.NativeOleMode,
                materializePreview: true);
            for (var index = 0; index < chunk.Count; index++)
            {
                var target = chunk[index];
                if (materialized[index].Error is { } error)
                {
                    failures.Add(T(
                        $"无法重绘（{WordFormulaRenderException.Excerpt(target.Metadata.Latex, 60)}）：{OfficeErrorText.Describe(error)}",
                        $"Could not redraw ({WordFormulaRenderException.Excerpt(target.Metadata.Latex, 60)}): {OfficeErrorText.Describe(error)}"));
                    continue;
                }
                var template = materialized[index].Template!;
                var session = CloneBulkFormulaSession(
                    template.Session, runs[index], documentId, target.FontSizePt, FormulaOleContract.NativeOleMode);
                // the formula as it was: its id, number, rows, title and creation; the new picture's fonts and extent
                // come with the render (ToMetadata takes the export result's)
                var original = target.Metadata;
                session.Mode = "edit";
                session.FormulaId = target.FormulaId;
                session.Numbered = target.Numbered;
                session.DisplayMode = target.DisplayMode;
                session.Title = string.IsNullOrWhiteSpace(original.Title) ? session.Title : original.Title;
                if (original.Lines is { Count: > 0 })
                {
                    session.Lines = original.Lines.Select(line => new FormulaLine { Id = line.Id, Latex = line.Latex }).ToList();
                    session.CodeFormat = string.IsNullOrWhiteSpace(original.CodeFormat) ? session.CodeFormat : original.CodeFormat;
                }
                session.OriginalMetadata = original;
                original.WordInlinePreviewShiftPx = null;
                original.WordInlineSourceBottomWhitespacePt = null;
                prepared.Add((target, new PreparedWordBulkFormula
                {
                    Run = runs[index],
                    Session = session,
                    MathMl = template.MathMl,
                    PngPath = template.PngPath,
                    EmfPath = template.EmfPath,
                }));
            }
        }
        progress.Report(T("正在用当前字体重绘公式…", "Drawing equations in the current fonts…"), targets.Count, targets.Count);
        return prepared;
    }
}
