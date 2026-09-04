using System.Runtime.InteropServices;
using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunActiveDoc17NumberedMathTypeBulkAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var requestedName = Environment.GetEnvironmentVariable("VISUALTEX_BULK_LIVE_SOURCE_NAME");
        if (string.IsNullOrWhiteSpace(requestedName)) requestedName = "文档17";
        var logPath = Path.Combine(artifactRoot, "active-doc17-numbered-mathtype-bulk.log");
        var outputPath = Path.Combine(artifactRoot, "Active-Doc17-Numbered-MathType-Bulk.docx");
        try { File.Delete(logPath); } catch { }
        try { File.Delete(outputPath); } catch { }

        var previousSource = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE");
        var previousSourcePath = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH");
        var previousFormat = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT");
        var previousMode = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE");
        var previousNumber = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_NUMBER_DISPLAY_FORMULAS");
        var previousLog = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG");

        Word.Application? application = null;
        Word.Documents? documents = null;
        Word.Document? sourceDocument = null;
        Word.Document? targetDocument = null;
        Word.Document? reopenedDocument = null;
        Word.Range? sourceRange = null;
        ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = (Word.Application)Marshal.GetActiveObject("Word.Application");
            documents = application.Documents;
            for (var index = 1; index <= documents.Count; index++)
            {
                Word.Document? candidate = null;
                try
                {
                    candidate = documents[index];
                    if (!string.Equals(candidate.Name, requestedName, StringComparison.Ordinal))
                        continue;
                    sourceDocument = candidate;
                    candidate = null;
                    break;
                }
                finally { Release(candidate); }
            }
            if (sourceDocument is null)
                throw new FileNotFoundException($"The current Word instance does not contain '{requestedName}'.");

            sourceRange = sourceDocument.Content.Duplicate;
            var source = (sourceRange.Text ?? string.Empty).TrimEnd('\r', '\a');
            var parsed = WordBulkImportParser.Parse(
                source,
                WordBulkSourceFormat.Latex,
                WordBulkFormulaObjectMode.MathType);
            parsed.NumberDisplayFormulas = true;
            Console.WriteLine(
                $"[DOC17 BULK PARSE] source={sourceDocument.Name}; chars={source.Length}; blocks={parsed.Blocks.Count}; formulas={parsed.FormulaCount}; display={parsed.DisplayFormulaCount}; inline={parsed.InlineFormulaCount}; warnings={parsed.Warnings.Count}");
            foreach (var warning in parsed.Warnings)
                Console.WriteLine("  WARNING: " + warning);
            var parsedFormulas = parsed.Blocks
                .SelectMany(block => block.Runs)
                .Where(run => run.IsFormula)
                .ToArray();
            for (var index = 0; index < parsedFormulas.Length; index++)
            {
                var run = parsedFormulas[index];
                Console.WriteLine(
                    $"  PARSED#{index + 1} display={run.DisplayMode} latex={run.Latex}");
            }
            AssertEqual(7, parsed.FormulaCount,
                "Document17 source should parse into exactly seven formulas.");
            AssertEqual(7, parsed.DisplayFormulaCount,
                "Document17 source should parse into exactly seven display formulas.");
            AssertEqual(0, parsed.InlineFormulaCount,
                "Document17 source unexpectedly contains an inline formula according to the bulk parser.");

            targetDocument = application.Documents.Add();
            targetDocument.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                targetDocument,
                EquationNumberFormat.Heading2DotId);

            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE", source);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", "latex");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", "mathtype");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_NUMBER_DISPLAY_FORMULAS", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", logPath);

            addIn = new ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);
            addIn.OnBulkImport(new object());
            WaitForBulkImportCompletion(logPath, TimeSpan.FromMinutes(3));
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));

            var mathTypeCount = CountMathTypeOleShapes(targetDocument);
            var placeRefCount = CountMathTypePlaceRefFields(targetDocument);
            Console.WriteLine(
                $"[DOC17 BULK RESULT] MathType={mathTypeCount}; MTPlaceRef={placeRefCount}; InlineShapes={targetDocument.InlineShapes.Count}; Fields={targetDocument.Fields.Count}; Paragraphs={targetDocument.Paragraphs.Count}");
            DumpNumberedMathTypeBulkRows(targetDocument);

            targetDocument.SaveAs2(
                outputPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);

            AssertEqual(7, mathTypeCount,
                "Document17 numbered MathType bulk import did not retain all seven Equation.DSMT4 objects.");
            AssertEqual(7, placeRefCount,
                "Document17 numbered MathType bulk import did not retain exactly seven MTPlaceRef fields.");
            AssertMathTypeNumberTexts(
                targetDocument,
                "(0.0.1)", "(0.0.2)", "(0.0.3)", "(0.0.4)",
                "(0.0.5)", "(0.0.6)", "(0.0.7)");

            var service = new WordFormulaService(application);
            var metadata = ReadBulkMathTypeMetadata(
                service,
                targetDocument,
                "Document17 numbered MathType bulk import");
            AssertEqual(7, metadata.Count,
                "Document17 numbered MathType bulk import did not leave seven editable MathType formulas.");
            AssertDoc17BulkFormulaSemantics(metadata);
            AssertEveryMathTypeOleHasReadableMathMl(
                targetDocument,
                expectedCount: 7,
                "fresh Document17 numbered MathType bulk import");

            targetDocument.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(targetDocument);
            targetDocument = null;
            reopenedDocument = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            AssertEqual(7, CountMathTypeOleShapes(reopenedDocument),
                "Save/reopen lost a Document17 bulk-imported MathType OLE.");
            AssertEqual(7, CountMathTypePlaceRefFields(reopenedDocument),
                "Save/reopen changed the Document17 MathType number-row count.");
            AssertMathTypeNumberTexts(
                reopenedDocument,
                "(0.0.1)", "(0.0.2)", "(0.0.3)", "(0.0.4)",
                "(0.0.5)", "(0.0.6)", "(0.0.7)");
            AssertEveryMathTypeOleHasReadableMathMl(
                reopenedDocument,
                expectedCount: 7,
                "save/reopened Document17 numbered MathType bulk import");

            Console.WriteLine(
                "[DOC17 BULK PASS] Exact live Document17 source imported as seven numbered MathType display formulas; all seven OLEs and all seven MTPlaceRef rows survived save/reopen.");
            Console.WriteLine("Artifact: " + outputPath);
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE", previousSource);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", previousSourcePath);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", previousFormat);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", previousMode);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_NUMBER_DISPLAY_FORMULAS", previousNumber);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", previousLog);
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); }
                catch { }
            }
            if (reopenedDocument is not null)
            {
                try { reopenedDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            if (targetDocument is not null)
            {
                try { targetDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            if (sourceDocument is not null)
            {
                try { sourceDocument.Activate(); } catch { }
            }
            Release(sourceRange);
            Release(reopenedDocument);
            Release(targetDocument);
            Release(sourceDocument);
            Release(documents);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void AssertDoc17BulkFormulaSemantics(IReadOnlyList<FormulaMetadata> metadata)
    {
        var expectedTokens = new[]
        {
            new[] { @"\partial^{2}u", @"\partial x^{2}", @"f(x,y)" },
            new[] { "x=0", "x=a", "y", "b" },
            new[] { "y=0", "y=b", "x", "a" },
            new[] { @"\sum_{n=1}^{+\infty}", "c_{nm}", "d_{nm}" },
            new[] { @"\partial^{2}w", @"\partial^{2}v", "t>0" },
            new[] { @"w\right|_{t=0}", @"\partial u", @"\partial v" },
            new[] { @"\langle f|L|g\rangle", @"L^{\dagger}", "Q[" },
        };
        AssertEqual(expectedTokens.Length, metadata.Count,
            "Document17 semantic token table no longer matches the imported formula count.");
        for (var index = 0; index < metadata.Count; index++)
        {
            var latex = metadata[index].Latex ?? string.Empty;
            foreach (var token in expectedTokens[index])
            {
                AssertTrue(latex.IndexOf(token, StringComparison.Ordinal) >= 0,
                    $"Document17 MathType formula #{index + 1} lost semantic token '{token}'. actual='{latex}'");
            }
        }
    }

    private static void AssertEveryMathTypeOleHasReadableMathMl(
        Word.Document document,
        int expectedCount,
        string context)
    {
        var found = 0;
        Word.InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                    var mathMl = MathTypeOleStorage.ReadMathMl(shape);
                    AssertTrue(!string.IsNullOrWhiteSpace(mathMl),
                        $"{context}: MathType OLE #{index} has no readable MathML.");
                    found++;
                }
                finally { Release(shape); }
            }
        }
        finally { Release(shapes); }
        AssertEqual(expectedCount, found,
            $"{context}: readable MathType OLE count changed.");
    }

    private static void DumpNumberedMathTypeBulkRows(Word.Document document)
    {
        for (var index = 1; index <= document.Paragraphs.Count; index++)
        {
            Word.Paragraph? paragraph = null;
            Word.Range? range = null;
            Word.Fields? fields = null;
            Word.InlineShapes? shapes = null;
            try
            {
                paragraph = document.Paragraphs[index];
                range = paragraph.Range;
                fields = range.Fields;
                shapes = range.InlineShapes;
                var hasPlaceRef = false;
                for (var fieldIndex = 1; fieldIndex <= fields.Count; fieldIndex++)
                {
                    Word.Field? field = null;
                    Word.Range? code = null;
                    try
                    {
                        field = fields[fieldIndex];
                        code = field.Code;
                        if ((code.Text ?? string.Empty).IndexOf(
                                "MACROBUTTON MTPlaceRef",
                                StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            hasPlaceRef = true;
                            break;
                        }
                    }
                    finally
                    {
                        Release(code);
                        Release(field);
                    }
                }
                if (!hasPlaceRef && shapes.Count == 0) continue;
                Console.WriteLine(
                    $"  ROW P{index} range={range.Start}:{range.End} shapes={shapes.Count} fields={fields.Count} placeRef={hasPlaceRef} text='{NormalizeBulkDiagnosticText(range.Text)}'");
            }
            finally
            {
                Release(shapes);
                Release(fields);
                Release(range);
                Release(paragraph);
            }
        }
    }

    private static string NormalizeBulkDiagnosticText(string? text)
    {
        var value = text ?? string.Empty;
        value = value.Replace("\r", "<CR>").Replace("\a", "<CELL>").Replace("\t", "<TAB>");
        return value.Length <= 220 ? value : value.Substring(0, 217) + "...";
    }
}
