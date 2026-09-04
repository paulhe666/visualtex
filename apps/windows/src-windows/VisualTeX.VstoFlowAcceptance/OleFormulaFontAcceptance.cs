using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.PowerPointVsto;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunOleFormulaFontAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var tempRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(tempRoot);
        var timesSvg = Path.Combine(tempRoot, $"{Guid.NewGuid():N}.svg");
        var arialSvg = Path.Combine(tempRoot, $"{Guid.NewGuid():N}.svg");
        var timesPng = Path.Combine(tempRoot, $"{Guid.NewGuid():N}.png");
        var arialPng = Path.Combine(tempRoot, $"{Guid.NewGuid():N}.png");
        string? timesEmf = null;
        string? arialEmf = null;
        try
        {
            File.WriteAllText(timesSvg, CreateFontAcceptanceSvg("Times New Roman", "SimSun"), new UTF8Encoding(false));
            // Keep the western font identical so this regression proves that the
            // Chinese font alone survives SVG -> vector EMF conversion.
            File.WriteAllText(arialSvg, CreateFontAcceptanceSvg("Times New Roman", "KaiTi"), new UTF8Encoding(false));
            WriteAcceptancePng(timesPng, "Emc 中文", 320, 96);
            WriteAcceptancePng(arialPng, "Emc 中文", 320, 96);
            timesEmf = OfficeOlePreview.CreateVectorEmfFromSvg(timesSvg, 320, 96);
            arialEmf = OfficeOlePreview.CreateVectorEmfFromSvg(arialSvg, 320, 96);
            AssertChineseEmfFontRenderingDiffers(timesEmf, arialEmf);
            RunWordOleFontReplacementAcceptance(artifactRoot, timesPng, timesEmf, arialPng, arialEmf);
            RunWordInlineOleReeditReplacementAcceptance(artifactRoot, timesPng, timesEmf, arialPng, arialEmf);
            RunPowerPointOleFontReplacementAcceptance(artifactRoot, timesPng, timesEmf, arialPng, arialEmf);
            Console.WriteLine("OLE formula font acceptance passed: the native EMF glyph outlines changed and Word/PowerPoint rebuilt embedded OLE objects for font-only edits.");
        }
        finally
        {
            foreach (var path in new[] { timesSvg, arialSvg, timesPng, arialPng, timesEmf, arialEmf })
            {
                if (string.IsNullOrWhiteSpace(path)) continue;
                try { File.Delete(path!); } catch { }
            }
        }
    }

    private static void RunWordOleFontReplacementAcceptance(
        string artifactRoot,
        string initialPng,
        string initialEmf,
        string updatedPng,
        string updatedEmf)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? originalShape = null;
        Word.InlineShape? updatedShape = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Activate();
            application.Selection.SetRange(0, 0);
            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var initial = CreateOleFontSession(
                "word", "create", formulaId, document.FullName, WordRangeReference(0, 0), null,
                "times", "songti");
            service.InsertOle(initial, initialPng, initialEmf);
            AssertEqual(1, document.InlineShapes.Count, "Word did not insert the initial OLE font fixture.");
            originalShape = document.InlineShapes[1];
            var originalMetadata = WordFormulaMetadataReader.TryRead(originalShape)
                ?? throw new InvalidOperationException("Word OLE metadata could not be read before the font-only edit.");

            var update = CreateOleFontSession(
                "word", "edit", formulaId, document.FullName, formulaId, originalMetadata,
                "times", "kaiti");
            service.ReplaceOle(update, updatedPng, updatedEmf);
            AssertEqual(1, document.InlineShapes.Count, "Word font-only OLE replacement changed the formula count.");
            updatedShape = document.InlineShapes[1];
            var updatedMetadata = WordFormulaMetadataReader.TryRead(updatedShape)
                ?? throw new InvalidOperationException("Word OLE metadata could not be read after the font-only edit.");
            AssertEqual("times", updatedMetadata.FormulaLetterFont,
                "Word OLE Chinese-font-only replacement changed formulaLetterFont.");
            AssertEqual("kaiti", updatedMetadata.FormulaChineseFont,
                "Word OLE font-only replacement lost formulaChineseFont.");
            AssertTrue(IsDeletedWordInlineShape(originalShape),
                "Word reused the old display OLE object for a font-only edit, so its presentation cache can remain stale.");

            document.SaveAs2(
                Path.Combine(artifactRoot, "word-ole-font-replacement.docx"),
                Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine("  Word OLE Chinese-font-only edit rebuilt the embedded object and persisted times/kaiti metadata.");
        }
        finally
        {
            Release(updatedShape);
            Release(originalShape);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordInlineOleReeditReplacementAcceptance(
        string artifactRoot,
        string initialPng,
        string initialEmf,
        string updatedPng,
        string updatedEmf)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? originalShape = null;
        Word.InlineShape? updatedShape = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Activate();
            application.Selection.SetRange(0, 0);
            application.Selection.TypeText("before  after");
            application.Selection.SetRange(7, 7);

            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var initial = CreateOleFontSession(
                "word", "create", formulaId, document.FullName, WordRangeReference(7, 7), null,
                "times", "songti", displayMode: "inline", latex: @"a^2+b^2=c^2");
            initial.ExportResult!.Width = 75.5938f;
            initial.ExportResult.Height = 14.8226f;
            initial.ExportResult.Baseline = 11.5f;
            service.InsertOle(initial, initialPng, initialEmf);
            AssertEqual(1, document.InlineShapes.Count,
                "Word did not insert exactly one initial inline VisualTeX OLE formula.");
            originalShape = document.InlineShapes[1];
            var originalStart = originalShape.Range.Start;
            var originalEnd = originalShape.Range.End;
            var originalMetadata = WordFormulaMetadataReader.TryRead(originalShape)
                ?? throw new InvalidOperationException(
                    "Word inline OLE metadata could not be read before re-editing.");
            Word.Range? selectedSourceRange = null;
            try
            {
                selectedSourceRange = originalShape.Range.Duplicate;
                selectedSourceRange.Select();
            }
            finally { Release(selectedSourceRange); }

            var update = CreateOleFontSession(
                "word", "edit", formulaId, document.FullName,
                WordRangeReference(originalStart, originalEnd), originalMetadata,
                "times", "songti", displayMode: "inline", latex: @"a^2+b^2=c^2+1");
            update.ExportResult!.Width = 141.1502f;
            update.ExportResult.Height = 14.8534f;
            update.ExportResult.Baseline = 11.5f;
            service.ReplaceOle(update, updatedPng, updatedEmf);

            AssertEqual(1, document.InlineShapes.Count,
                "Re-editing one inline VisualTeX OLE formula appended a stale duplicate.");
            updatedShape = document.InlineShapes[1];
            var updatedMetadata = WordFormulaMetadataReader.TryRead(updatedShape)
                ?? throw new InvalidOperationException(
                    "Word inline OLE metadata could not be read after re-editing.");
            AssertEqual(formulaId, updatedMetadata.FormulaId,
                "Word inline OLE re-edit changed the formula identity.");
            AssertEqual(@"a^2+b^2=c^2+1", updatedMetadata.Latex,
                "Word inline OLE re-edit retained the old LaTeX payload.");
            AssertEqual(originalStart, updatedShape.Range.Start,
                "Word inline OLE re-edit moved the formula away from its original text position.");
            AssertTrue(IsDeletedWordInlineShape(originalShape),
                "Word inline OLE re-edit retained the old embedded object.");

            document.SaveAs2(
                Path.Combine(artifactRoot, "word-inline-ole-reedit-replacement.docx"),
                Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                "  Word inline VisualTeX OLE re-edit replaced one object in place without appending the stale source OLE.");
        }
        finally
        {
            Release(updatedShape);
            Release(originalShape);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunPowerPointOleFontReplacementAcceptance(
        string artifactRoot,
        string initialPng,
        string initialEmf,
        string updatedPng,
        string updatedEmf)
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? originalShape = null;
        PowerPoint.Shape? updatedShape = null;
        try
        {
            application = new PowerPoint.Application { Visible = Office.MsoTriState.msoTrue };
            presentation = application.Presentations.Add(Office.MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            var service = new PowerPointFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var initial = CreateOleFontSession(
                "powerpoint", "create", formulaId, null, null, null, "times", "songti");
            var inserted = service.InsertOle(initial, initialPng, initialEmf);
            AssertTrue(!string.IsNullOrWhiteSpace(inserted.ObjectId),
                "PowerPoint did not return the initial OLE shape name.");
            var originalName = inserted.ObjectId!;
            originalShape = slide.Shapes[originalName];
            var originalMetadata = ReadPowerPointOleMetadata(originalShape)
                ?? throw new InvalidOperationException("PowerPoint OLE metadata could not be read before the font-only edit.");

            var update = CreateOleFontSession(
                "powerpoint", "edit", formulaId, null, originalName, originalMetadata,
                "times", "kaiti");
            var replaced = service.ReplaceOle(update, updatedPng, updatedEmf);
            AssertTrue(!string.IsNullOrWhiteSpace(replaced.ObjectId),
                "PowerPoint did not return the font-updated OLE shape name.");
            AssertTrue(IsDeletedPowerPointShape(originalShape),
                "PowerPoint reused the old OLE shape for a font-only edit, so its presentation cache can remain stale.");
            updatedShape = slide.Shapes[replaced.ObjectId!];
            var updatedMetadata = ReadPowerPointOleMetadata(updatedShape)
                ?? throw new InvalidOperationException("PowerPoint OLE metadata could not be read after the font-only edit.");
            AssertEqual("times", updatedMetadata.FormulaLetterFont,
                "PowerPoint OLE Chinese-font-only replacement changed formulaLetterFont.");
            AssertEqual("kaiti", updatedMetadata.FormulaChineseFont,
                "PowerPoint OLE font-only replacement lost formulaChineseFont.");

            presentation.SaveAs(Path.Combine(artifactRoot, "powerpoint-ole-font-replacement.pptx"));
            Console.WriteLine("  PowerPoint OLE Chinese-font-only edit rebuilt the embedded object and persisted times/kaiti metadata.");
        }
        finally
        {
            Release(updatedShape);
            Release(originalShape);
            Release(slide);
            try { presentation?.Close(); } catch { }
            Release(presentation);
            try { application?.Quit(); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static OfficeSessionDocument CreateOleFontSession(
        string host,
        string mode,
        string formulaId,
        string? sourceDocumentId,
        string? sourceObjectId,
        FormulaMetadata? originalMetadata,
        string letterFont,
        string chineseFont,
        string displayMode = "block",
        string latex = @"E=mc^2+\text{中文}")
    {
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = mode,
            Host = host,
            FormulaId = formulaId,
            SourceDocumentId = sourceDocumentId,
            SourceObjectId = sourceObjectId,
            Title = "OLE formula font acceptance",
            CodeFormat = "latex",
            DisplayMode = displayMode,
            ObjectMode = FormulaOleContract.NativeOleMode,
            Numbered = false,
            FontSizePt = 20,
            OriginalMetadata = originalMetadata,
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
            ExportResult = new OfficeExportDocument
            {
                Width = 320,
                Height = 96,
                Baseline = 72,
                FormulaLetterFont = letterFont,
                FormulaChineseFont = chineseFont,
            },
        };
    }

    private static FormulaMetadata? ReadPowerPointOleMetadata(
        PowerPoint.Slide slide,
        string shapeName)
    {
        PowerPoint.Shape? shape = null;
        return ReadPowerPointOleMetadata(shape = slide.Shapes[shapeName], releaseShape: true);
    }

    private static FormulaMetadata? ReadPowerPointOleMetadata(
        PowerPoint.Shape shape,
        bool releaseShape = false)
    {
        PowerPoint.OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            format = shape.OLEFormat;
            try { oleObject = format.Object; } catch { }
            if (oleObject is null)
            {
                try { format.DoVerb(); } catch { }
                try { oleObject = format.Object; } catch { }
            }
            return oleObject is IVisualTeXFormulaObject formula
                ? FormulaOleInterop.ReadMetadata(formula)
                : null;
        }
        finally
        {
            Release(oleObject);
            Release(format);
            if (releaseShape) Release(shape);
        }
    }

    private static bool IsDeletedWordInlineShape(Word.InlineShape shape)
    {
        try
        {
            _ = shape.Width;
            _ = shape.Range.Start;
            return false;
        }
        catch (COMException)
        {
            return true;
        }
    }

    private static bool IsDeletedPowerPointShape(PowerPoint.Shape shape)
    {
        try
        {
            _ = shape.Width;
            _ = shape.Left;
            return false;
        }
        catch (COMException)
        {
            return true;
        }
    }

    private static string CreateFontAcceptanceSvg(string westernFont, string chineseFont) =>
        $"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 320 96\" width=\"320\" height=\"96\">"
        + $"<text x=\"8\" y=\"66\" font-size=\"58\" font-family=\"{westernFont}\" font-style=\"italic\">Emc</text>"
        + $"<text x=\"150\" y=\"66\" font-size=\"48\" font-family=\"{chineseFont}\">中文</text>"
        + "</svg>";

    private static void AssertChineseEmfFontRenderingDiffers(string firstPath, string secondPath)
    {
        using var first = RenderEmfForFontAcceptance(firstPath);
        using var second = RenderEmfForFontAcceptance(secondPath);
        var westernChanged = 0;
        var chineseChanged = 0;
        var chineseInk = 0;
        for (var y = 0; y < first.Height; y++)
        for (var x = 0; x < first.Width; x++)
        {
            var a = first.GetPixel(x, y);
            var b = second.GetPixel(x, y);
            var difference = Math.Abs(a.R - b.R) + Math.Abs(a.G - b.G) + Math.Abs(a.B - b.B);
            if (x < 140)
            {
                if (difference >= 90) westernChanged++;
                continue;
            }
            if (a.R + a.G + a.B < 700 || b.R + b.G + b.B < 700) chineseInk++;
            if (difference >= 90) chineseChanged++;
        }
        AssertTrue(chineseInk > 200, "OLE Chinese-font EMF fixture rendered no meaningful Chinese glyph ink.");
        AssertTrue(westernChanged < 20,
            $"The western glyphs changed while only the Chinese font was switched ({westernChanged} differing pixels).");
        AssertTrue(chineseChanged > 150,
            $"Changing only the Chinese formula font did not materially change the native EMF Chinese glyph outlines ({chineseChanged} differing pixels).");
        Console.WriteLine($"  Native OLE Chinese EMF rendering differs at {chineseChanged} Chinese-region pixels while western region changed at {westernChanged} pixels.");
    }

    private static System.Drawing.Bitmap RenderEmfForFontAcceptance(string emfPath)
    {
        var bitmap = new System.Drawing.Bitmap(
            320,
            96,
            System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var metafile = new System.Drawing.Imaging.Metafile(emfPath);
        using var graphics = System.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(System.Drawing.Color.White);
        graphics.DrawImage(metafile, new System.Drawing.Rectangle(0, 0, 320, 96));
        return bitmap;
    }
}
