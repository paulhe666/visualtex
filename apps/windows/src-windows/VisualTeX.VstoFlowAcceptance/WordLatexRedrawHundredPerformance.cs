using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;
using WinForms = System.Windows.Forms;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordLatexRedrawHundredPerformance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        const int formulaCount = 100;
        var target =
            Environment.GetEnvironmentVariable("VISUALTEX_REDRAW_100_TARGET")
            ?? "visualtex";
        var mathType = string.Equals(
            target,
            "mathtype",
            StringComparison.OrdinalIgnoreCase);
        if (!mathType
            && !string.Equals(
                target,
                "visualtex",
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                $"Unsupported 100-formula redraw target '{target}'.");

        var suspendScreenUpdating = string.Equals(
            Environment.GetEnvironmentVariable(
                "VISUALTEX_REDRAW_100_SUSPEND_SCREEN_UPDATING"),
            "1",
            StringComparison.Ordinal);

        var suffix =
            $"{(mathType ? "mathtype" : "visualtex")}-screen-{(suspendScreenUpdating ? "off" : "on")}";
        var redrawLog = Path.Combine(
            artifactRoot,
            $"word-latex-redraw-100-{suffix}.log");
        var documentPath = Path.Combine(
            artifactRoot,
            $"VisualTeX-Word-Latex-Redraw-100-{suffix}.docx");
        TryDeleteAcceptanceFile(redrawLog);
        TryDeleteAcceptanceFile(documentPath);

        var previousRedrawLog =
            Environment.GetEnvironmentVariable(
                "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG");
        Environment.SetEnvironmentVariable(
            "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG",
            redrawLog);
        try
        {
            client.PrewarmConverterAsync(
                    CancellationToken.None)
                .GetAwaiter()
                .GetResult();
            Thread.Sleep(700);
            WinForms.Application.DoEvents();

            using (var host =
                   new WordPerformanceHost(documentPath: null))
            {
                PopulateHundredLatexRedrawPerformanceDocument(
                    host,
                    formulaCount);
                host.Document.Activate();
                host.Application.Selection.HomeKey(
                    Word.WdUnits.wdStory);

                var previousScreenUpdating =
                    host.Application.ScreenUpdating;
                var wall = Stopwatch.StartNew();
                try
                {
                    if (suspendScreenUpdating)
                        host.Application.ScreenUpdating = false;

                    if (mathType)
                        host.AddIn.OnRedrawDocumentToMathType(
                            new object());
                    else
                        host.AddIn.OnRedrawDocumentToOle(
                            new object());

                    var log = WaitForLatexRedraw(
                        redrawLog,
                        TimeSpan.FromMinutes(15));
                    WaitForAddInIdle(
                        host.AddIn,
                        TimeSpan.FromSeconds(180));
                    wall.Stop();

                    AssertHundredLatexRedrawSummary(
                        log,
                        formulaCount);
                    var timings =
                        ParseHundredLatexRedrawInsertTimings(
                            log,
                            formulaCount);
                    var firstTen =
                        timings.Take(10).Average();
                    var middleTen =
                        timings.Skip(45).Take(10).Average();
                    var lastTen =
                        timings.Skip(90).Take(10).Average();

                    AssertHundredLatexRedrawDocument(
                        host.Document,
                        formulaCount,
                        mathType);
                    host.Save(documentPath);

                    Console.WriteLine(
                        $"[Word LaTeX redraw 100] target={target}; wall={wall.ElapsedMilliseconds}ms; "
                        + $"screenUpdatingSuspended={suspendScreenUpdating}; "
                        + $"first10={firstTen:F1}ms; middle10={middleTen:F1}ms; last10={lastTen:F1}ms; "
                        + $"max={timings.Max()}ms.");
                }
                finally
                {
                    if (suspendScreenUpdating)
                    {
                        try
                        {
                            host.Application.ScreenUpdating =
                                previousScreenUpdating;
                        }
                        catch { }
                    }
                }

                // Partial-failure redraw now commits each formula independently
                // so one bad source can be rolled back without discarding earlier
                // successes. Save the successful document first, then verify that
                // one native Undo restores exactly the last committed formula.
                object undoCount = 1;
                if (!host.Document.Undo(ref undoCount))
                    throw new InvalidDataException(
                        "100-formula LaTeX redraw did not produce an undoable final formula mutation.");
                Word.Range? undoneContent = null;
                try
                {
                    undoneContent =
                        host.Document.Content;
                    var spans =
                        WordBulkImportParser.FindFormulaSpans(
                            undoneContent.Text
                            ?? string.Empty);
                    if (spans.Count != 1)
                        throw new InvalidDataException(
                            $"One Undo restored {spans.Count} LaTeX source formulas; expected exactly the last committed formula.");
                }
                finally
                {
                    Release(undoneContent);
                }
            }

            using (var reopened =
                   new WordPerformanceHost(documentPath))
            {
                AssertHundredLatexRedrawDocument(
                    reopened.Document,
                    formulaCount,
                    mathType);
            }

            Console.WriteLine(
                $"[Word LaTeX redraw 100] save/reopen + single-target undo validation passed: {documentPath}");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG",
                previousRedrawLog);
        }
    }

    private static void RunWordLatexRedrawHundredReopenVerify(string artifactRoot)
    {
        var documentPath = Environment.GetEnvironmentVariable("VISUALTEX_REDRAW_100_REOPEN_PATH");
        if (string.IsNullOrWhiteSpace(documentPath))
            throw new InvalidDataException(
                "VISUALTEX_REDRAW_100_REOPEN_PATH is required for reopen verification.");
        documentPath = Path.GetFullPath(documentPath);
        if (!File.Exists(documentPath))
            throw new FileNotFoundException(
                "The saved 100-formula redraw document does not exist.",
                documentPath);

        using var reopened = new WordPerformanceHost(documentPath);
        AssertHundredLatexRedrawDocument(
            reopened.Document,
            expectedCount: 100,
            mathType: true);

        Word.Range? content = null;
        try
        {
            content = reopened.Document.Content;
            var text = content.Text ?? string.Empty;
            if (WordBulkImportParser.FindFormulaSpans(text).Count != 0)
                throw new InvalidDataException(
                    "Reopened MathType document unexpectedly contains redraw-source LaTeX spans.");
        }
        finally
        {
            Release(content);
        }

        var evidencePath = Path.Combine(
            artifactRoot,
            "mathtype-100-reopen-verified.txt");
        File.WriteAllText(
            evidencePath,
            "PASS\r\n" + documentPath + "\r\n",
            System.Text.Encoding.UTF8);
        Console.WriteLine(
            $"[Word LaTeX redraw 100 reopen] MathType save/reopen metadata + FormulaId validation passed: {documentPath}");
    }

    private static void PopulateHundredLatexRedrawPerformanceDocument(
        WordPerformanceHost host,
        int formulaCount)
    {
        var selection = host.Application.Selection;
        selection.HomeKey(Word.WdUnits.wdStory);
        selection.Font.Name = "宋体";
        selection.Font.Size = 10.5f;
        selection.Font.Position = 0;

        for (var index = 1;
             index <= formulaCount;
             index++)
        {
            if ((index & 1) == 1)
            {
                selection.TypeText(
                    $"Inline {index} before $x_{{{index}}}+y_{{{index}}}={index}$ after");
                selection.TypeParagraph();
            }
            else
            {
                selection.TypeText(
                    $"Display {index} before");
                selection.TypeParagraph();
                selection.TypeText(
                    $"$$E_{{{index}}}=mc^2+{index}$$");
                selection.TypeParagraph();
                selection.TypeText(
                    $"Display {index} after");
                selection.TypeParagraph();
            }
        }
    }

    private static void AssertHundredLatexRedrawSummary(
        string log,
        int expectedCount)
    {
        var summary = Regex.Match(
            log,
            @"\bredraw-complete formulas=(?<formulas>\d+) unique=(?<unique>\d+) skipped=(?<skipped>\d+) "
            + @"renderAverageMs=(?<renderAverage>\d+) renderMaxMs=(?<renderMax>\d+) "
            + @"insertTotalMs=(?<insertTotal>\d+) insertMaxMs=(?<insertMax>\d+)");
        if (!summary.Success)
            throw new InvalidDataException(
                "100-formula redraw did not emit a complete summary.\n"
                + log);

        var formulas = int.Parse(
            summary.Groups["formulas"].Value,
            CultureInfo.InvariantCulture);
        var skipped = int.Parse(
            summary.Groups["skipped"].Value,
            CultureInfo.InvariantCulture);
        if (formulas != expectedCount
            || skipped != 0)
            throw new InvalidDataException(
                $"100-formula redraw completed {formulas}/{expectedCount} formulas with {skipped} skipped.");
    }

    private static long[] ParseHundredLatexRedrawInsertTimings(
        string log,
        int expectedCount)
    {
        var matches = Regex.Matches(
            log,
            @"\binsert-timing index=(?<index>\d+)/(?<count>\d+) elapsedMs=(?<elapsed>\d+)");
        if (matches.Count != expectedCount)
            throw new InvalidDataException(
                $"100-formula redraw emitted {matches.Count}/{expectedCount} per-insert timing rows.");

        var result = new long[expectedCount];
        foreach (Match match in matches)
        {
            var index = int.Parse(
                match.Groups["index"].Value,
                CultureInfo.InvariantCulture);
            var count = int.Parse(
                match.Groups["count"].Value,
                CultureInfo.InvariantCulture);
            var elapsed = long.Parse(
                match.Groups["elapsed"].Value,
                CultureInfo.InvariantCulture);
            if (count != expectedCount
                || index < 1
                || index > expectedCount)
                throw new InvalidDataException(
                    "100-formula redraw emitted an invalid timing index.");
            result[index - 1] =
                elapsed;
        }
        return result;
    }

    private static void AssertHundredLatexRedrawDocument(
        Word.Document document,
        int expectedCount,
        bool mathType)
    {
        var actualCount =
            mathType
                ? CountMathTypeOleFormulas(document)
                : CountVisualTeXNativeOleShapes(document);
        if (actualCount != expectedCount)
            throw new InvalidDataException(
                $"Expected {expectedCount} {(mathType ? "MathType" : "VisualTeX")} hosts; actual {actualCount}.");

        Word.Range? content = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        try
        {
            content = document.Content;
            var finalText =
                content.Text
                ?? string.Empty;
            if (finalText.IndexOf(
                    "$",
                    StringComparison.Ordinal) >= 0)
                throw new InvalidDataException(
                    "100-formula redraw left dollar-delimited LaTeX source in the document.");
            for (var index = 1;
                 index <= expectedCount;
                 index++)
            {
                var before =
                    (index & 1) == 1
                        ? $"Inline {index} before "
                        : $"Display {index} before";
                var after =
                    (index & 1) == 1
                        ? " after"
                        : $"Display {index} after";
                if (finalText.IndexOf(
                        before,
                        StringComparison.Ordinal) < 0
                    || finalText.IndexOf(
                        after,
                        StringComparison.Ordinal) < 0)
                    throw new InvalidDataException(
                        $"100-formula redraw damaged surrounding body text near formula {index}.");
            }

            shapes = document.InlineShapes;
            var formulaIds =
                new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                Release(shape);
                shape = shapes[index];
                if (mathType)
                {
                    if (!MathTypeOleInterop.IsMathTypeOle(shape))
                        continue;
                    if (!WordFormulaIdentityStore.TryResolveMathType(
                            shape,
                            out var formulaId)
                        || !Guid.TryParse(
                            formulaId,
                            out _)
                        || !formulaIds.Add(
                            formulaId))
                        throw new InvalidDataException(
                            "MathType redraw produced a missing, invalid, or duplicate FormulaId.");
                }
                else
                {
                    if (!WordFormulaMetadataReader.IsNativeOle(
                            shape))
                        continue;
                    var metadata =
                        WordFormulaMetadataReader.TryRead(
                            shape)
                        ?? throw new InvalidDataException(
                            "VisualTeX redraw produced an OLE host without embedded metadata.");
                    if (!Guid.TryParse(
                            metadata.FormulaId,
                            out _)
                        || !formulaIds.Add(
                            metadata.FormulaId)
                        || string.IsNullOrWhiteSpace(
                            metadata.Latex))
                        throw new InvalidDataException(
                            "VisualTeX redraw produced invalid or duplicate embedded metadata.");
                }
            }

            if (formulaIds.Count != expectedCount)
                throw new InvalidDataException(
                    $"100-formula redraw resolved {formulaIds.Count}/{expectedCount} unique FormulaIds.");
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(content);
        }
    }
}
