using System.Drawing;
using System.Drawing.Imaging;
using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class MathTypeColorPreviewEvidence
    {
        public int RedPixels { get; set; }
        public int BluePixels { get; set; }
        public int DarkPixels { get; set; }
        public string PngPath { get; set; } = string.Empty;
        public string PreviewSource { get; set; } = string.Empty;
    }

    private static void RunWordOwnedMathTypeColorEditAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var placeholderSvg = Path.Combine(artifactRoot, "black-placeholder.svg");
        File.WriteAllText(
            placeholderSvg,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"280\" height=\"80\" viewBox=\"0 0 280 80\">"
            + "<rect width=\"280\" height=\"80\" fill=\"white\"/><text x=\"4\" y=\"55\" "
            + "font-family=\"Times New Roman\" font-size=\"42\" fill=\"black\">xy+zq</text></svg>");
        var placeholderEmf = OfficeOlePreview.CreateVectorEmfFromSvg(
            placeholderSvg,
            280,
            80);

        const string plainMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mi>x</mi><mi>y</mi><mo>+</mo><mi>z</mi><mi>q</mi></math>";
        const string coloredMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mstyle mathcolor=\"#DC2626\"><mi>x</mi>"
            + "<mstyle mathcolor=\"#2563EB\"><mi>y</mi></mstyle>"
            + "<mo>+</mo><mi>z</mi></mstyle><mi>q</mi></math>";
        const string coloredLatex =
            @"\textcolor[HTML]{DC2626}{x\textcolor[HTML]{2563EB}{y}+z}q";

        var documentPath = Path.Combine(
            artifactRoot,
            "VisualTeX-MathType-Color-Edit.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            range = document.Range(0, 0);
            range.Select();
            var service = new WordFormulaService(application);

            service.InsertMathTypeOle(
                CreateMathTypeCreateSession("inline", false, "xy+zq"),
                plainMathMl,
                placeholderEmf);
            AssertEqual(1, document.InlineShapes.Count,
                "The isolated create did not produce exactly one MathType OLE.");
            shape = document.InlineShapes[1];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                "The isolated create did not produce Equation.DSMT4.");
            Release(range);
            range = shape.Range.Duplicate;
            range.Select();
            var source = service.ReadSelection();
            AssertEqual(FormulaOleContract.MathTypeOleMode, source.ObjectMode,
                "The newly created equation was not read through the MathType edit path.");

            var editSession = CreateMathTypeProductSession(
                source,
                coloredLatex,
                exportWidth: 280,
                exportHeight: 80);
            Release(shape);
            shape = null;
            Release(range);
            range = null;
            service.ReplaceMathTypeOle(editSession, coloredMathMl, placeholderEmf);

            AssertEqual(1, document.InlineShapes.Count,
                "Applying the editor color duplicated or removed the MathType OLE.");
            shape = document.InlineShapes[1];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                "Applying the editor color changed the object away from Equation.DSMT4.");
            var appliedReadback = MathTypeOleStorage.ReadMathMl(shape);
            AssertMathTypeColorReadback(appliedReadback, "immediately after editor apply");
            File.WriteAllText(
                Path.Combine(artifactRoot, "applied-readback.mathml"),
                appliedReadback);
            var appliedPreview = CaptureMathTypeColorPreview(
                shape,
                artifactRoot,
                "applied-preview.png");
            AssertTrue(appliedPreview.RedPixels >= 8,
                $"The Word MathType preview has no measurable red glyph ink ({appliedPreview.RedPixels} pixels).");
            AssertTrue(appliedPreview.BluePixels >= 8,
                $"The Word MathType preview has no measurable blue glyph ink ({appliedPreview.BluePixels} pixels).");
            AssertTrue(appliedPreview.DarkPixels >= 8,
                $"The uncolored trailing q is no longer rendered as dark ink ({appliedPreview.DarkPixels} pixels).");

            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            Release(shape);
            shape = null;

            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                Visible: false);
            AssertEqual(1, document.InlineShapes.Count,
                "Save/reopen changed the colored MathType OLE count.");
            shape = document.InlineShapes[1];
            var reopenedReadback = MathTypeOleStorage.ReadMathMl(shape);
            AssertMathTypeColorReadback(reopenedReadback, "after save/reopen");
            File.WriteAllText(
                Path.Combine(artifactRoot, "reopened-readback.mathml"),
                reopenedReadback);
            var reopenedPreview = CaptureMathTypeColorPreview(
                shape,
                artifactRoot,
                "reopened-preview.png");
            AssertTrue(reopenedPreview.RedPixels >= 8 && reopenedPreview.BluePixels >= 8,
                $"Save/reopen lost colored MathType preview ink: red={reopenedPreview.RedPixels}, blue={reopenedPreview.BluePixels}.");

            File.WriteAllText(
                Path.Combine(artifactRoot, "result.json"),
                JsonSerializer.Serialize(
                    new
                    {
                        DocumentPath = documentPath,
                        ObjectCount = document.InlineShapes.Count,
                        ProgId = shape.OLEFormat.ProgID,
                        AppliedReadback = appliedReadback,
                        ReopenedReadback = reopenedReadback,
                        AppliedPreview = appliedPreview,
                        ReopenedPreview = reopenedPreview,
                    },
                    new JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine(
                $"[MathType color edit PASS] red/blue/dark pixels="
                + $"{appliedPreview.RedPixels}/{appliedPreview.BluePixels}/{appliedPreview.DarkPixels}; "
                + $"reopened={reopenedPreview.RedPixels}/{reopenedPreview.BluePixels}/{reopenedPreview.DarkPixels}.");
            Console.WriteLine($"Document: {documentPath}");
        }
        finally
        {
            Release(shape);
            Release(range);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void AssertMathTypeColorReadback(string mathMl, string stage)
    {
        var tokens = XDocument.Parse(mathMl).Root!
            .DescendantsAndSelf()
            .Where(element => element.Name.LocalName is "mi" or "mo")
            .Where(element => element.Value is "x" or "y" or "+" or "z" or "q")
            .ToDictionary(
                element => element.Value,
                MathMlColorContract.Effective);
        AssertEqual("#DC2626", tokens["x"], $"{stage}: x lost the editor's red color.");
        AssertEqual("#2563EB", tokens["y"], $"{stage}: y lost the editor's blue color.");
        AssertEqual("#DC2626", tokens["+"], $"{stage}: + lost the inherited red color.");
        AssertEqual("#DC2626", tokens["z"], $"{stage}: z lost the inherited red color.");
        AssertEqual<string?>(null, tokens["q"], $"{stage}: the uncolored q inherited a preceding color.");
    }

    private static MathTypeColorPreviewEvidence CaptureMathTypeColorPreview(
        Word.InlineShape shape,
        string artifactRoot,
        string fileName)
    {
        byte[] emf;
        var previewSource = "word-clipboard-enhanced-metafile";
        try
        {
            emf = ReadInlineShapeEnhancedMetafile(shape);
        }
        catch (Exception clipboardError)
        {
            // A reopened invisible Word document can keep the clipboard busy
            // even though the OLE presentation stored in the document is valid.
            // Read that exact persisted MathType WMF cache through Flat OPC so
            // the visual assertion still measures Word's object presentation,
            // never the black SVG supplied as the fallback input preview.
            var fragment = MathTypeWordOpenXml.Read(shape);
            var wmfPath = Path.Combine(
                artifactRoot,
                Path.GetFileNameWithoutExtension(fileName) + ".wmf");
            File.WriteAllBytes(wmfPath, fragment.PreviewWmf);
            var converted = MathTypeWordOpenXml.ConvertPlaceableWmfToEnhancedMetafile(
                wmfPath,
                shape.Width,
                shape.Height,
                artifactRoot);
            emf = File.ReadAllBytes(converted);
            previewSource = "word-flat-opc-ole-presentation-cache; clipboard="
                + clipboardError.GetType().Name;
        }
        var width = Math.Max(560, (int)Math.Ceiling(shape.Width * 10d));
        var height = Math.Max(180, (int)Math.Ceiling(shape.Height * 10d));
        using var bitmap = RenderEmf(emf, width, height);
        var path = Path.Combine(artifactRoot, fileName);
        bitmap.Save(path, ImageFormat.Png);
        var evidence = new MathTypeColorPreviewEvidence
        {
            PngPath = path,
            PreviewSource = previewSource,
        };
        for (var y = 0; y < bitmap.Height; y++)
        {
            for (var x = 0; x < bitmap.Width; x++)
            {
                var color = bitmap.GetPixel(x, y);
                if (color.R >= 100 && color.R >= color.G * 3 / 2 && color.R >= color.B * 3 / 2)
                    evidence.RedPixels++;
                if (color.B >= 85 && color.B >= color.R * 6 / 5 && color.B >= color.G * 6 / 5)
                    evidence.BluePixels++;
                if (color.R < 110 && color.G < 110 && color.B < 110)
                    evidence.DarkPixels++;
            }
        }
        return evidence;
    }
}
