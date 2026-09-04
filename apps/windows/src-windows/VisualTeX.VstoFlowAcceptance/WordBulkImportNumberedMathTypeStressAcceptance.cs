using System.Runtime.InteropServices;
using System.Text;
using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordBulkImportNumberedMathTypeStressAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var logPath = Path.Combine(artifactRoot, "numbered-mathtype-bulk-stress.log");
        var outputPath = Path.Combine(artifactRoot, "Numbered-MathType-Bulk-Stress.docx");
        try { File.Delete(logPath); } catch { }
        try { File.Delete(outputPath); } catch { }

        var source = CreateNumberedMathTypeBulkStressSource();
        var parsed = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.MathType);
        parsed.NumberDisplayFormulas = true;
        AssertEqual(20, parsed.FormulaCount,
            "Numbered MathType stress source must contain exactly twenty formulas.");
        AssertEqual(10, parsed.DisplayFormulaCount,
            "Numbered MathType stress source must contain exactly ten display formulas.");
        AssertEqual(10, parsed.InlineFormulaCount,
            "Numbered MathType stress source must contain exactly ten inline formulas.");
        AssertEqual(0, parsed.Warnings.Count,
            "Numbered MathType stress source unexpectedly produced parser warnings.");

        var previousSource = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE");
        var previousSourcePath = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH");
        var previousFormat = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT");
        var previousMode = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE");
        var previousNumber = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_NUMBER_DISPLAY_FORMULAS");
        var previousLog = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = (Word.Application)Marshal.GetActiveObject("Word.Application");
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
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
            var elapsedMs = WaitForBulkImportCompletion(logPath, TimeSpan.FromMinutes(4));
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(45));

            AssertNumberedMathTypeBulkStressState(
                document,
                expectedOleCount: 20,
                expectedNumberedCount: 10,
                "fresh stress import");
            AssertStressProseOrder(document);
            var service = new WordFormulaService(application);
            var metadata = ReadBulkMathTypeMetadata(
                service,
                document,
                "numbered MathType stress import");
            AssertEqual(20, metadata.Count,
                "Numbered MathType stress import did not leave twenty editable formulas.");
            AssertStressFormulaIdentity(metadata);

            document.SaveAs2(
                outputPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            reopened = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            AssertNumberedMathTypeBulkStressState(
                reopened,
                expectedOleCount: 20,
                expectedNumberedCount: 10,
                "save/reopened stress import");
            AssertStressProseOrder(reopened);
            var reopenedService = new WordFormulaService(application);
            var reopenedMetadata = ReadBulkMathTypeMetadata(
                reopenedService,
                reopened,
                "save/reopened numbered MathType stress import");
            AssertEqual(20, reopenedMetadata.Count,
                "Save/reopen lost an editable MathType formula from the stress import.");
            AssertStressFormulaIdentity(reopenedMetadata);

            Console.WriteLine(
                $"[NUMBERED MATHTYPE BULK STRESS PASS] 10 inline + 10 display formulas with interleaved prose survived import and save/reopen; OLE=20, MTPlaceRef=10, elapsedMs={elapsedMs}.");
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
            try { reopened?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(reopened);
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static string CreateNumberedMathTypeBulkStressSource()
    {
        var source = new StringBuilder();
        for (var index = 1; index <= 10; index++)
        {
            source.AppendLine(
                $"STRESS-PROSE-{index:00}-BEFORE 普通正文在行内公式前。"
                + $" 行内公式 $x_{{{index}}}^2+y_{{{index}}}^2=z_{{{index}}}^2$ 后继续正文，"
                + "用于确认 MathType 行内 OLE 不吞前后文字。");
            source.AppendLine();
            source.AppendLine("\\begin{equation*}");
            if (index % 3 == 1)
            {
                source.AppendLine(
                    $"\\frac{{\\partial^2 u_{{{index}}}}}{{\\partial x^2}}"
                    + $"+\\frac{{\\partial^2 u_{{{index}}}}}{{\\partial y^2}}"
                    + $"=f_{{{index}}}(x,y)");
            }
            else if (index % 3 == 2)
            {
                source.AppendLine(
                    $"\\sum_{{n=1}}^{{+\\infty}} c_{{{index},n}}"
                    + $"\\sin\\frac{{n\\pi x}}{{a_{{{index}}}}}"
                    + $"=g_{{{index}}}(x)");
            }
            else
            {
                source.AppendLine(
                    $"\\left.\\frac{{\\partial v_{{{index}}}}}{{\\partial t}}\\right|_{{t=0}}"
                    + $"=-\\lambda_{{{index}}} v_{{{index}}}(x,0)");
            }
            source.AppendLine("\\end{equation*}");
            source.AppendLine(
                $"STRESS-PROSE-{index:00}-AFTER 行间编号公式之后立即继续普通正文，"
                + "专门验证后续文字不会覆盖刚插入的 Equation.DSMT4。 ");
            source.AppendLine();
        }
        return source.ToString();
    }

    private static void AssertNumberedMathTypeBulkStressState(
        Word.Document document,
        int expectedOleCount,
        int expectedNumberedCount,
        string context)
    {
        AssertEqual(expectedOleCount, CountMathTypeOleShapes(document),
            $"{context}: MathType OLE count changed.");
        AssertEqual(expectedOleCount, document.InlineShapes.Count,
            $"{context}: an unexpected non-MathType InlineShape appeared or a MathType OLE disappeared.");
        AssertEqual(expectedNumberedCount, CountMathTypePlaceRefFields(document),
            $"{context}: numbered MathType MTPlaceRef count changed.");
        AssertEveryMathTypeOleHasReadableMathMl(document, expectedOleCount, context);
        AssertEqual(0, CountNumberOnlyMathTypeRows(document),
            $"{context}: one or more numbered MathType rows retained MTPlaceRef but lost Equation.DSMT4.");
    }

    private static int CountNumberOnlyMathTypeRows(Word.Document document)
    {
        var count = 0;
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
                if (!hasPlaceRef) continue;
                shapes = range.InlineShapes;
                var hasMathType = false;
                for (var shapeIndex = 1; shapeIndex <= shapes.Count; shapeIndex++)
                {
                    Word.InlineShape? shape = null;
                    try
                    {
                        shape = shapes[shapeIndex];
                        if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                        hasMathType = true;
                        break;
                    }
                    finally { Release(shape); }
                }
                if (!hasMathType) count++;
            }
            finally
            {
                Release(shapes);
                Release(fields);
                Release(range);
                Release(paragraph);
            }
        }
        return count;
    }

    private static void AssertStressProseOrder(Word.Document document)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            var text = content.Text ?? string.Empty;
            var cursor = -1;
            for (var index = 1; index <= 10; index++)
            {
                var before = $"STRESS-PROSE-{index:00}-BEFORE";
                var after = $"STRESS-PROSE-{index:00}-AFTER";
                var beforePosition = text.IndexOf(before, cursor + 1, StringComparison.Ordinal);
                var afterPosition = text.IndexOf(after, Math.Max(0, beforePosition + before.Length), StringComparison.Ordinal);
                AssertTrue(beforePosition > cursor && afterPosition > beforePosition,
                    $"Stress prose markers for section {index:00} were lost or reordered.");
                cursor = afterPosition;
            }
        }
        finally { Release(content); }
    }

    private static void AssertStressFormulaIdentity(IReadOnlyList<FormulaMetadata> metadata)
    {
        AssertEqual(20, metadata.Count,
            "Stress MathType metadata count changed.");
        var inlineCount = metadata.Count(item => string.Equals(
            item.DisplayMode,
            "inline",
            StringComparison.Ordinal));
        var displayCount = metadata.Count(item => string.Equals(
            item.DisplayMode,
            "block",
            StringComparison.Ordinal));
        AssertEqual(10, inlineCount,
            "Stress import changed the inline MathType count.");
        AssertEqual(10, displayCount,
            "Stress import changed the display MathType count.");

        for (var index = 1; index <= 10; index++)
        {
            AssertTrue(metadata.Any(item =>
                    string.Equals(item.DisplayMode, "inline", StringComparison.Ordinal)
                    && (item.Latex ?? string.Empty).IndexOf($"x_{{{index}}}", StringComparison.Ordinal) >= 0),
                $"Stress inline formula {index} is missing after MathType import.");
            AssertTrue(metadata.Any(item =>
                    string.Equals(item.DisplayMode, "block", StringComparison.Ordinal)
                    && (item.Latex ?? string.Empty).IndexOf($"_{{{index}}}", StringComparison.Ordinal) >= 0),
                $"Stress display formula {index} is missing after MathType import.");
        }
    }
}
