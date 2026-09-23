using System.Diagnostics;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordLatexRedrawImmediateNumberingAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var pngPath = Path.Combine(
            artifactRoot,
            "redraw-immediate-numbering.png");
        WriteAcceptancePng(
            pngPath,
            "redraw-numbering",
            360,
            112);
        var svgPath = Path.Combine(
            artifactRoot,
            "redraw-immediate-numbering.svg");
        File.WriteAllText(
            svgPath,
            @"<svg xmlns=""http://www.w3.org/2000/svg"" width=""360"" height=""112"" viewBox=""0 0 360 112""><text x=""4"" y=""72"" font-family=""Times New Roman"" font-size=""48"">x+1</text></svg>");
        var emfPath =
            OfficeOlePreview.CreateVectorEmfFromSvg(
                svgPath,
                360,
                112);

        var modes =
            new[]
            {
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.NativeOleMode,
            };
        var actual =
            new Dictionary<string, string[]>(
                StringComparer.Ordinal);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        Process? testOleServerProcess = null;
        object? oleServerKeepAlive = null;
        try
        {
            foreach (var objectMode in modes)
            {
                // Office can keep native-OLE activation ACL state in a Word
                // process after another formula-host family has been exercised.
                // Use a fresh isolated Word process per redraw target so this
                // acceptance measures numbering behavior, not cross-test OLE ACL
                // state. Production Ribbon redraw still runs inside the user's
                // existing Word process.
                if (string.Equals(
                        objectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal))
                {
                    var testOleServerPath =
                        Environment.GetEnvironmentVariable(
                            "VISUALTEX_TEST_OLE_SERVER_PATH");
                    if (!string.IsNullOrWhiteSpace(
                            testOleServerPath))
                    {
                        testOleServerProcess =
                            Process.Start(
                                new ProcessStartInfo
                                {
                                    FileName =
                                        testOleServerPath,
                                    Arguments =
                                        "-Embedding",
                                    UseShellExecute =
                                        false,
                                    CreateNoWindow =
                                        true,
                                })
                            ?? throw new InvalidOperationException(
                                "Failed to start the test VisualTeX OLE server.");
                        Thread.Sleep(500);
                        if (testOleServerProcess.HasExited)
                            throw new InvalidOperationException(
                                "The test VisualTeX OLE server exited before Word COM activation.");
                        var oleServerType =
                            Type.GetTypeFromProgID(
                                FormulaOleContract.ProgId,
                                throwOnError: true)
                            ?? throw new InvalidOperationException(
                                "VisualTeX OLE server class is not registered.");
                        oleServerKeepAlive =
                            Activator.CreateInstance(
                                oleServerType)
                            ?? throw new InvalidOperationException(
                                "VisualTeX OLE server keep-alive could not be created.");
                        Console.WriteLine(
                            $"[REDRAW IMMEDIATE NUMBERING] Test VisualTeX OLE server pinned: pid={testOleServerProcess.Id}, path={testOleServerPath}");
                    }
                }

                application =
                    string.Equals(
                        objectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal)
                        ? CreateFreshAcceptanceAutomationWord(
                            artifactRoot)
                        : CreateWordApplication(
                            visible: false);
                var modeName =
                    string.Equals(
                        objectMode,
                        FormulaOleContract.WordOmmlMode,
                        StringComparison.Ordinal)
                        ? "omml"
                        : string.Equals(
                            objectMode,
                            FormulaOleContract.NativeOleMode,
                            StringComparison.Ordinal)
                            ? "visualtex"
                            : "mathtype";
                var path = Path.Combine(
                    artifactRoot,
                    $"redraw-immediate-numbering-{modeName}.docx");

                document =
                    application.Documents.Add();
                document.Content.Text =
                    "before\r$$a+b$$\r$$c+d$$\r$$x+y$$\rafter\r";
                var compatibilityMode =
                    string.Equals(
                        objectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal)
                        ? 15
                        : 14;
                document.SaveAs2(
                    FileName: path,
                    FileFormat:
                        Word.WdSaveFormat.wdFormatXMLDocument,
                    CompatibilityMode:
                        compatibilityMode);
                document.Close(
                    Word.WdSaveOptions.wdSaveChanges);
                Release(document);
                document =
                    application.Documents.Open(
                        path,
                        ReadOnly: false,
                        Visible: false);
                document.Activate();

                AssertEqual(
                    compatibilityMode,
                    document.CompatibilityMode,
                    $"{modeName} redraw fixture did not reopen in compatibility mode {compatibilityMode}.");
                WordEquationNumbering
                    .SetEquationNumberFormatPreference(
                        document,
                        EquationNumberFormat.ContinuousId);

                var service =
                    new WordFormulaService(
                        application);
                var plan =
                    service.CaptureLatexRedrawPlan(
                        wholeDocument: true);
                AssertEqual(
                    3,
                    plan.Targets.Count,
                    $"{modeName} redraw did not capture all three display LaTeX sources.");
                plan.NumberDisplayFormulas =
                    true;

                var prepared =
                    new Dictionary<string, PreparedWordBulkFormula>(
                        StringComparer.Ordinal);
                foreach (var target in plan.Targets)
                {
                    var latex =
                        target.Latex.Trim();
                    var mathMl =
                        latex switch
                        {
                            "a+b" =>
                                FirstNumberedMathMl,
                            "c+d" =>
                                SecondNumberedMathMl,
                            _ =>
                                SimpleMathMl,
                        };
                    prepared[target.Id] =
                        new PreparedWordBulkFormula
                        {
                            Run =
                                new WordBulkRun
                                {
                                    Id =
                                        target.Id,
                                    IsFormula =
                                        true,
                                    Latex =
                                        target.Latex,
                                    DisplayMode =
                                        target.DisplayMode,
                                },
                            Session =
                                CreateOmmlMathTypeAcceptanceSession(
                                    mathMl,
                                    target.DisplayMode,
                                    numbered: false,
                                    objectMode),
                            MathMl =
                                mathMl,
                            PngPath =
                                string.Equals(
                                    objectMode,
                                    FormulaOleContract.NativeOleMode,
                                    StringComparison.Ordinal)
                                    ? pngPath
                                    : null,
                            EmfPath =
                                string.Equals(
                                    objectMode,
                                    FormulaOleContract.WordOmmlMode,
                                    StringComparison.Ordinal)
                                    ? null
                                    : emfPath,
                        };
                }

                var redraw =
                    service.ApplyLatexRedrawPlan(
                        plan,
                        prepared);
                AssertEqual(
                    3,
                    redraw.FormulaCount,
                    $"{modeName} redraw did not replace all three sources.");

                string[] numbers;
                if (string.Equals(
                        objectMode,
                        FormulaOleContract.WordOmmlMode,
                        StringComparison.Ordinal))
                {
                    var index =
                        WordFormulaHostResolver
                            .CaptureDocumentIndex(
                                document);
                    var hosts =
                        index.Omml
                            .OrderBy(
                                host =>
                                    host.Range.Start)
                            .ToArray();
                    AssertEqual(
                        3,
                        hosts.Length,
                        "OMML redraw did not create exactly three native hosts.");
                    numbers =
                        hosts.Select(
                                host =>
                                {
                                    host.Numbering =
                                        WordFormulaNumberingResolver
                                            .ResolveLocal(
                                                document,
                                                host);
                                    return WordNativeOmmlNumbering
                                        .ReadVisibleNumber(
                                            document,
                                            host.Numbering);
                                })
                            .ToArray();
                }
                else if (string.Equals(
                             objectMode,
                             FormulaOleContract.NativeOleMode,
                             StringComparison.Ordinal))
                {
                    var index =
                        WordFormulaHostResolver
                            .CaptureDocumentIndex(
                                document);
                    var hosts =
                        index.VisualTeX
                            .OrderBy(
                                host =>
                                    host.Range.Start)
                            .ToArray();
                    AssertEqual(
                        3,
                        hosts.Length,
                        "VisualTeX redraw did not create exactly three OLE hosts.");
                    numbers =
                        hosts.Select(
                                host =>
                                    WordVisualTeXParagraphNumbering
                                        .ReadVisibleNumberText(
                                            document,
                                            host))
                            .ToArray();
                }
                else
                {
                    var values =
                        new List<(int Start, string Number)>();
                    for (var index = 1;
                         index <=
                         document.InlineShapes.Count;
                         index++)
                    {
                        Release(shape);
                        shape =
                            document.InlineShapes[index];
                        if (!string.Equals(
                                shape.OLEFormat.ProgID,
                                "Equation.DSMT4",
                                StringComparison.OrdinalIgnoreCase))
                            continue;
                        var shapeRange =
                            shape.Range;
                        try
                        {
                            values.Add(
                                (shapeRange.Start,
                                 ReadMathTypeVisibleNumberForShape(
                                         shape)
                                     .Trim()
                                     .Trim('(', ')')
                                     .Trim()));
                        }
                        finally
                        {
                            Release(shapeRange);
                        }
                    }
                    numbers =
                        values
                            .OrderBy(
                                item =>
                                    item.Start)
                            .Select(
                                item =>
                                    item.Number)
                            .ToArray();
                    AssertEqual(
                        3,
                        numbers.Length,
                        "MathType redraw did not create exactly three numbered Equation.DSMT4 objects.");
                }

                actual[modeName] =
                    numbers;
                Console.WriteLine(
                    $"[REDRAW IMMEDIATE NUMBERING] {modeName}={string.Join(",", numbers)}");

                document.Close(
                    Word.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                document = null;
                try
                {
                    QuitWordApplicationIfOwned(
                        application);
                }
                catch { }
                Release(application);
                application = null;
                Release(oleServerKeepAlive);
                oleServerKeepAlive = null;
                if (testOleServerProcess is not null)
                {
                    try
                    {
                        if (!testOleServerProcess.HasExited)
                            testOleServerProcess.Kill();
                    }
                    catch { }
                    testOleServerProcess.Dispose();
                    testOleServerProcess = null;
                }
                ForceComCleanup();
            }

            foreach (var pair in actual)
            {
                AssertEqual(
                    3,
                    pair.Value.Length,
                    $"{pair.Key} immediate-number result count is wrong.");
                AssertEqual(
                    "1",
                    pair.Value[0],
                    $"{pair.Key} first redraw number is wrong.");
                AssertEqual(
                    "2",
                    pair.Value[1],
                    $"{pair.Key} second redraw number is wrong.");
                AssertEqual(
                    "3",
                    pair.Value[2],
                    $"{pair.Key} third redraw number is wrong.");
            }

            Console.WriteLine(
                "[REDRAW IMMEDIATE NUMBERING PASS] OMML, VisualTeX and MathType redraw all materialized 1,2,3 immediately without a manual Update Equation Numbers command.");
        }
        finally
        {
            Release(shape);
            Release(range);
            Release(oleServerKeepAlive);
            if (testOleServerProcess is not null)
            {
                try
                {
                    if (!testOleServerProcess.HasExited)
                        testOleServerProcess.Kill();
                }
                catch { }
                testOleServerProcess.Dispose();
            }
            if (document is not null)
            {
                try
                {
                    document.Close(
                        Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(document);
            try
            {
                QuitWordApplicationIfOwned(
                    application);
            }
            catch { }
            Release(application);
            ForceComCleanup();
        }
    }
}
