using System.Runtime.InteropServices;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeOleCreateAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "mathtype-create-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"96\" viewBox=\"0 0 240 96\"><text x=\"4\" y=\"64\" font-family=\"Times New Roman\" font-size=\"48\">x+1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 96);

        var previousCreateObjectMode = WordEquationNumbering.GetDefaultCreateObjectMode();
        try
        {
            WordEquationNumbering.SetDefaultCreateObjectMode(FormulaOleContract.MathTypeOleMode);
            AssertEqual(
                FormulaOleContract.MathTypeOleMode,
                WordEquationNumbering.GetDefaultCreateObjectMode(),
                "Word did not remember MathType OLE as the create object format.");
            WordEquationNumbering.SetDefaultCreateObjectMode(FormulaOleContract.NativeOleMode);
            AssertEqual(
                FormulaOleContract.NativeOleMode,
                WordEquationNumbering.GetDefaultCreateObjectMode(),
                "Word did not remember VisualTeX OLE as the create object format.");
        }
        finally
        {
            WordEquationNumbering.SetDefaultCreateObjectMode(previousCreateObjectMode);
        }

        RunWordMathTypeInlineCreateAcceptance(artifactRoot, emfPath);
        RunWordMathTypeDisplayCreateAcceptance(artifactRoot, emfPath);
    }

    private static void RunWordMathTypeInlineCreateAcceptance(
        string artifactRoot,
        string emfPath)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "LEFT RIGHT";
            range = document.Range(5, 5);
            range.Select();

            var service = new WordFormulaService(application);
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "inline",
                    numbered: false,
                    latex: @"\frac{x+1}{y}"),
                FractionMathMl,
                emfPath);

            AssertEqual(1, document.InlineShapes.Count,
                "Standalone MathType inline create did not insert exactly one OLE object.");
            shape = document.InlineShapes[1];
            AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                "Standalone MathType inline create did not materialize Equation.DSMT4.");
            AssertWordMathTypePreviewVisible(
                shape,
                "Standalone MathType inline create",
                FractionMathMl,
                inline: true,
                artifactRoot);
            Console.WriteLine($"[MathType create inline probe] paragraphs={document.Paragraphs.Count}, shape={shape.Range.Start}-{shape.Range.End}");
            for (var paragraphIndex = 1; paragraphIndex <= document.Paragraphs.Count; paragraphIndex++)
            {
                var probeParagraph = document.Paragraphs[paragraphIndex];
                try
                {
                    var probeText = probeParagraph.Range.Text ?? string.Empty;
                    Console.WriteLine(
                        $"  P{paragraphIndex}={probeParagraph.Range.Start}-{probeParagraph.Range.End} cp="
                        + string.Join(",", probeText.Select(character => $"U+{(int)character:X4}")));
                }
                finally { Release(probeParagraph); }
            }
            AssertEqual(1, document.Paragraphs.Count,
                "Standalone MathType inline create unexpectedly split the text paragraph.");
            AssertTrue((document.Content.Text ?? string.Empty).Contains("LEFT")
                && (document.Content.Text ?? string.Empty).Contains("RIGHT"),
                "Standalone MathType inline create damaged surrounding prose.");
            var readback = MathTypeOleStorage.ReadMathMl(shape);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(FractionMathMl),
                MathTypeMtefCodec.SemanticSignature(readback),
                "Standalone MathType inline create changed formula semantics.");

            var path = Path.Combine(artifactRoot, "VisualTeX-MathType-Create-Inline.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            Release(shape);
            shape = document.InlineShapes[1];
            AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                "Standalone MathType inline create lost its ProgID after Word reopen.");
            readback = MathTypeOleStorage.ReadMathMl(shape);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(FractionMathMl),
                MathTypeMtefCodec.SemanticSignature(readback),
                "Standalone MathType inline create changed after Word reopen.");
            AssertWordMathTypePreviewVisible(
                shape,
                "Reopened standalone MathType inline create",
                FractionMathMl,
                inline: true,
                artifactRoot);
            Console.WriteLine("[MathType create] Inline Equation.DSMT4 insert + save/reopen passed.");
        }
        finally
        {
            Release(shape);
            Release(range);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordMathTypeDisplayCreateAcceptance(
        string artifactRoot,
        string emfPath)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabs = null;
        Word.TabStop? tab = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "display create acceptance";
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            var service = new WordFormulaService(application);

            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: false,
                    latex: @"x+y"),
                SimpleMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "Standalone unnumbered MathType display create did not insert one OLE.");
            shape = document.InlineShapes[1];
            AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                "Standalone unnumbered MathType display create did not materialize Equation.DSMT4.");
            AssertMathTypeDisplayRow(
                shape,
                expectNumber: false,
                "Standalone unnumbered MathType display create");
            AssertWordMathTypePreviewVisible(
                shape,
                "Standalone unnumbered MathType display create",
                SimpleMathMl,
                inline: false,
                artifactRoot);

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(2, document.InlineShapes.Count,
                "First numbered MathType display create did not add exactly one OLE.");
            Release(shape);
            shape = document.InlineShapes[2];
            AssertMathTypeDisplayRow(
                shape,
                expectNumber: true,
                "First numbered MathType display create");
            AssertWordMathTypePreviewVisible(
                shape,
                "First numbered MathType display create",
                FirstNumberedMathMl,
                inline: false,
                artifactRoot);
            AssertTrue(CountMathTypePlaceRefFields(document) == 1,
                "First numbered MathType display create did not create exactly one MTPlaceRef field.");
            AssertNativeMathTypeSectionBreak(document, 1);

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"c+d"),
                SecondNumberedMathMl,
                emfPath);
            AssertEqual(3, document.InlineShapes.Count,
                "Second numbered MathType display create did not add exactly one OLE.");
            Release(shape);
            shape = document.InlineShapes[3];
            AssertMathTypeDisplayRow(
                shape,
                expectNumber: true,
                "Second numbered MathType display create");
            AssertWordMathTypePreviewVisible(
                shape,
                "Second numbered MathType display create",
                SecondNumberedMathMl,
                inline: false,
                artifactRoot);
            AssertTrue(CountMathTypePlaceRefFields(document) == 2,
                "Second numbered MathType display create did not preserve/clone MTPlaceRef numbering.");
            AssertNativeMathTypeSectionBreak(document, 1);
            var codesBeforeSave = ReadMathTypePlaceRefCodes(document);
            AssertEqual(2, codesBeforeSave.Count,
                "MathType numbered create did not produce two durable MTPlaceRef codes.");
            AssertEqual(codesBeforeSave[0], codesBeforeSave[1],
                "Second numbered MathType display create did not inherit the existing MathType numbering template.");

            var path = Path.Combine(artifactRoot, "VisualTeX-MathType-Create-Display.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            AssertEqual(3, document.InlineShapes.Count,
                "MathType display creates changed object count after Word reopen.");
            AssertTrue(CountMathTypePlaceRefFields(document) == 2,
                "MathType MTPlaceRef fields did not survive Word reopen.");
            AssertNativeMathTypeSectionBreak(document, 1);
            var codesAfterReopen = ReadMathTypePlaceRefCodes(document);
            AssertEqual(codesBeforeSave[0], codesAfterReopen[0],
                "First MathType numbering template changed after Word reopen.");
            AssertEqual(codesBeforeSave[1], codesAfterReopen[1],
                "Second MathType numbering template changed after Word reopen.");
            for (var index = 1; index <= document.InlineShapes.Count; index++)
            {
                Release(shape);
                shape = document.InlineShapes[index];
                AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                    $"MathType display create #{index} lost Equation.DSMT4 after reopen.");
                AssertMathTypeDisplayRow(
                    shape,
                    expectNumber: index >= 2,
                    $"Reopened MathType display create #{index}");
                var expectedMathMl = index == 1
                    ? SimpleMathMl
                    : index == 2
                        ? FirstNumberedMathMl
                        : SecondNumberedMathMl;
                AssertWordMathTypePreviewVisible(
                    shape,
                    $"Reopened MathType display create #{index}",
                    expectedMathMl,
                    inline: false,
                    artifactRoot);
            }
            Console.WriteLine(
                "[MathType create] Unnumbered + two numbered display Equation.DSMT4 inserts, native MTPlaceRef numbering, template inheritance and save/reopen passed.");
        }
        finally
        {
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraph);
            Release(shape);
            Release(range);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void AssertWordMathTypePreviewVisible(
        Word.InlineShape shape,
        string context,
        string mathMl,
        bool inline,
        string artifactRoot)
    {
        var preview = ReadInlineShapeEnhancedMetafile(shape);
        var ink = DescribeEmfInkBounds(preview);
        Console.WriteLine($"[MathType create preview] {context}: {ink}");
        AssertTrue(!string.Equals(ink, "empty", StringComparison.Ordinal),
            context + " is a valid Equation.DSMT4 object but Word renders its OLE preview as blank.");

        var generated = MathTypeMtefCodec.CreateEquationNative(mathMl, inline);
        if (!MathTypeNativePreviewRenderer.TryRender(
                generated.Mtef,
                artifactRoot,
                out var nativePreview))
            return;
        using (nativePreview)
        {
            var expectedNativeWmf = File.ReadAllBytes(nativePreview.WmfPath);
            var difference = MeasureEmfPixelDifference(expectedNativeWmf, preview);
            Console.WriteLine(
                $"[MathType create preview] {context}: native diff={difference:0.0000}, "
                + $"size={shape.Width:0.0}x{shape.Height:0.0}pt, "
                + $"native={nativePreview.WidthPt:0.0}x{nativePreview.HeightPt:0.0}pt");
            AssertTrue(
                difference < 0.03,
                context + " is visible but does not visually match MathType's native renderer.");
            AssertNear(
                nativePreview.WidthPt,
                shape.Width,
                0.7f,
                context + " does not use MathType's native width.");
            AssertNear(
                nativePreview.HeightPt,
                shape.Height,
                0.7f,
                context + " does not use MathType's native height.");
            AssertNear(
                nativePreview.WordPosition,
                ReadInlineOlePositionForAcceptance(shape),
                1.0f,
                context + " does not use MathType's native baseline.");
        }
    }

    private static void AssertMathTypeDisplayRow(
        Word.InlineShape shape,
        bool expectNumber,
        string context)
    {
        Word.Range? shapeRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabs = null;
        Word.TabStop? tab = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        object? paragraphStyleObject = null;
        Word.Style? paragraphStyle = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            AssertEqual(1, paragraphs.Count, context + " spans more than one paragraph.");
            paragraph = paragraphs[1];
            paragraphStyleObject = paragraph.Range.get_Style();
            paragraphStyle = paragraphStyleObject as Word.Style;
            AssertTrue(paragraphStyle is not null,
                context + " does not expose a Word paragraph style.");
            AssertEqual("MTDisplayEquation", paragraphStyle!.NameLocal,
                context + " does not use MathType's MTDisplayEquation style.");
            format = paragraph.Format;
            tabs = format.TabStops;
            var sawCenter = false;
            var sawRight = false;
            for (var index = 1; index <= tabs.Count; index++)
            {
                Release(tab);
                tab = tabs[index];
                sawCenter |= tab.Alignment == Word.WdTabAlignment.wdAlignTabCenter;
                sawRight |= tab.Alignment == Word.WdTabAlignment.wdAlignTabRight;
            }
            AssertTrue(sawCenter && sawRight,
                context + " does not have MathType center/right tab stops.");

            var paragraphText = paragraph.Range.Text ?? string.Empty;
            AssertTrue(
                paragraphText.StartsWith("\t\u0001", StringComparison.Ordinal),
                context + " does not begin with Word's native tab + OLE sequence.");

            fields = paragraph.Range.Fields;
            var sawPlaceRef = false;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                sawPlaceRef |= (code.Text ?? string.Empty).IndexOf(
                    "MACROBUTTON MTPlaceRef",
                    StringComparison.OrdinalIgnoreCase) >= 0;
            }
            AssertEqual(expectNumber, sawPlaceRef,
                context + " has the wrong MathType MTPlaceRef numbering ownership.");
            if (expectNumber)
                AssertTrue(
                    paragraphText.IndexOf("\u0001\t", StringComparison.Ordinal) >= 0,
                    context + " does not position the MathType number with a trailing tab.");
        }
        finally
        {
            Release(paragraphStyle);
            paragraphStyle = null;
            paragraphStyleObject = null;
            Release(code);
            Release(field);
            Release(fields);
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static void AssertNativeMathTypeSectionBreak(
        Word.Document document,
        int expectedCount)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Fields? nestedFields = null;
        Word.Field? nestedField = null;
        Word.Range? nestedCode = null;
        object? styleObject = null;
        Word.Style? style = null;
        var breakCount = 0;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var codeText = code.Text ?? string.Empty;
                if (codeText.IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                breakCount++;
                AssertEqual(document.Content.Start, code.Start - 1,
                    "The default MathType chapter/section break is not at the start of the document.");
                var normalizedOuter = NormalizeFieldCodeForMathTypeAcceptance(codeText);
                AssertTrue(normalizedOuter.IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0,
                    "MathType section break lost its MTEditEquationSection2 MacroButton.");

                nestedFields = code.Fields;
                AssertEqual(3, nestedFields.Count,
                    "MathType default section break does not contain exactly three nested SEQ fields.");
                var nestedCodes = new List<(int Start, string Code)>();
                for (var nestedIndex = 1; nestedIndex <= nestedFields.Count; nestedIndex++)
                {
                    Release(nestedCode);
                    nestedCode = null;
                    Release(nestedField);
                    nestedField = nestedFields[nestedIndex];
                    nestedCode = nestedField.Code;
                    nestedCodes.Add((
                        nestedCode.Start,
                        NormalizeFieldCodeForMathTypeAcceptance(nestedCode.Text ?? string.Empty)));
                }
                var ordered = nestedCodes.OrderBy(item => item.Start).Select(item => item.Code).ToArray();
                AssertEqual("SEQ MTEqn \\r \\h \\* MERGEFORMAT", ordered[0],
                    "MathType section break does not reset MTEqn using MathType's native field code.");
                AssertEqual("SEQ MTSec \\r 1 \\h \\* MERGEFORMAT", ordered[1],
                    "MathType section break does not initialize MTSec to 1 using MathType's native field code.");
                AssertEqual("SEQ MTChap \\r 1 \\h \\* MERGEFORMAT", ordered[2],
                    "MathType section break does not initialize MTChap to 1 using MathType's native field code.");

                styleObject = code.get_Style();
                style = styleObject as Word.Style;
                AssertTrue(style is not null,
                    "MathType section break does not expose its character style.");
                AssertEqual("MTEquationSection", style!.NameLocal,
                    "MathType section break does not use MTEquationSection.");
                AssertEqual(-1, style.Font.Hidden,
                    "MTEquationSection is not hidden like native MathType.");
                AssertEqual((int)Word.WdColor.wdColorRed, (int)style.Font.Color,
                    "MTEquationSection is not red like native MathType.");

                Release(style);
                style = null;
                styleObject = null;
                Release(nestedCode);
                nestedCode = null;
                Release(nestedField);
                nestedField = null;
                Release(nestedFields);
                nestedFields = null;
            }
            AssertEqual(expectedCount, breakCount,
                "MathType create inserted the wrong number of chapter/section breaks.");
        }
        finally
        {
            Release(style);
            styleObject = null;
            Release(nestedCode);
            Release(nestedField);
            Release(nestedFields);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static int CountMathTypePlaceRefFields(Word.Document document) =>
        ReadMathTypePlaceRefCodes(document).Count;

    private static List<string> ReadMathTypePlaceRefCodes(Word.Document document)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        var result = new List<(int Start, string Code)>();
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var text = code.Text ?? string.Empty;
                if (text.IndexOf("MACROBUTTON MTPlaceRef", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                result.Add((code.Start, NormalizeFieldCodeForMathTypeAcceptance(text)));
            }
            return result.OrderBy(item => item.Start).Select(item => item.Code).ToList();
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static string NormalizeFieldCodeForMathTypeAcceptance(string value) =>
        string.Join(
            " ",
            (value ?? string.Empty)
                .Replace("\u0013", " ")
                .Replace("\u0014", " ")
                .Replace("\u0015", " ")
                .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private static OfficeSessionDocument CreateMathTypeCreateSession(
        string displayMode,
        bool numbered,
        string latex) =>
        new()
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = Guid.NewGuid().ToString("D"),
            SourceDocumentId = null,
            SourceObjectId = null,
            Title = "MathType standalone create acceptance",
            CodeFormat = "latex",
            DisplayMode = displayMode,
            ObjectMode = FormulaOleContract.MathTypeOleMode,
            Numbered = numbered,
            FontSizePt = 12,
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
            ExportResult = new OfficeExportDocument
            {
                Width = 240,
                Height = 96,
                Baseline = 72,
            },
        };

    private const string FractionMathMl =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow><mi>y</mi></mfrac></math>";
    private const string SimpleMathMl =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mi>y</mi></math>";
    private const string FirstNumberedMathMl =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>a</mi><mo>+</mo><mi>b</mi></math>";
    private const string SecondNumberedMathMl =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>c</mi><mo>+</mo><mi>d</mi></math>";
}
