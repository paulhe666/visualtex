using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeOleProductRoundTripAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var fixture = ResolveMathTypeNativeEditorFixture();
        if (!File.Exists(fixture))
            throw new FileNotFoundException(
                "A genuine MathType-generated Equation.DSMT4 fixture is required.", fixture);

        var compatibilityCase = Environment.GetEnvironmentVariable(
            "VISUALTEX_MATHTYPE_PRODUCT_CASE")?.Trim();
        string? compatibilityLatex = null;
        string? compatibilityMathMl = null;
        if (!string.IsNullOrWhiteSpace(compatibilityCase))
        {
            switch (compatibilityCase.ToLowerInvariant())
            {
                case "underbrace":
                    compatibilityLatex = @"\underbrace{a+b}_{n}";
                    compatibilityMathMl =
                        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><munder><mrow data-mjx-texclass=\"OP\"><munder><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mo>⏟</mo></munder></mrow><mrow><mi>n</mi></mrow></munder></math>";
                    break;
                case "overbrace":
                    compatibilityLatex = @"\overbrace{a+b}^{n}";
                    compatibilityMathMl =
                        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mover><mrow data-mjx-texclass=\"OP\"><mover><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mo>⏞</mo></mover></mrow><mrow><mi>n</mi></mrow></mover></math>";
                    break;
                case "mathbb":
                    compatibilityLatex = @"\mathbb{R}";
                    compatibilityMathMl =
                        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi mathvariant=\"double-struck\">R</mi></math>";
                    break;
                case "max":
                    compatibilityLatex = @"\max_{x\in A} f(x)";
                    compatibilityMathMl =
                        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><munder><mo data-mjx-texclass=\"OP\" movablelimits=\"true\">max</mo><mrow><mi>x</mi><mo>∈</mo><mi>A</mi></mrow></munder><mi>f</mi><mo>(</mo><mi>x</mi><mo>)</mo></math>";
                    break;
                case "iiint":
                    compatibilityLatex = @"\iiint_{V} f\,dV";
                    compatibilityMathMl =
                        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msub><mo data-mjx-texclass=\"OP\">∭</mo><mi>V</mi></msub><mi>f</mi><mspace width=\"0.167em\"></mspace><mi>d</mi><mi>V</mi></math>";
                    break;
                case "xcancel":
                    compatibilityLatex = @"\xcancel{x+1}";
                    compatibilityMathMl =
                        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><menclose notation=\"updiagonalstrike downdiagonalstrike\"><mi>x</mi><mo>+</mo><mn>1</mn></menclose></math>";
                    break;
                case "mixed-bigop-sizing":
                    compatibilityLatex = @"\frac{a}{b}+\int c.fga^2+b^2=c^2(a+b)^n=\sum_{k=0}^{n}\frac{n}{k}a^{n-k}b^k\sum_c^a b";
                    compatibilityMathMl =
                        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                        + "<mfrac><mi>a</mi><mi>b</mi></mfrac><mo>+</mo>"
                        + "<mo>∫</mo><mi>c</mi><mo>.</mo><mi>f</mi><mi>g</mi><msup><mi>a</mi><mn>2</mn></msup>"
                        + "<mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo>"
                        + "<msup><mi>c</mi><mn>2</mn></msup><msup><mfenced><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow></mfenced><mi>n</mi></msup><mo>=</mo>"
                        + "<msubsup><mo>∑</mo><mrow><mi>k</mi><mo>=</mo><mn>0</mn></mrow><mi>n</mi></msubsup>"
                        + "<mfrac><mi>n</mi><mi>k</mi></mfrac><msup><mi>a</mi><mrow><mi>n</mi><mo>−</mo><mi>k</mi></mrow></msup><msup><mi>b</mi><mi>k</mi></msup>"
                        + "<msubsup><mo>∑</mo><mi>c</mi><mi>a</mi></msubsup><mi>b</mi></math>";
                    break;
                default:
                    throw new InvalidDataException(
                        $"Unknown MathType product compatibility case '{compatibilityCase}'.");
            }
        }

        var path = Path.Combine(
            artifactRoot,
            $"VisualTeX-MathType7-Product-{Guid.NewGuid():N}.docx");
        File.Copy(fixture, path, overwrite: false);
        var previewSvg = Path.Combine(artifactRoot, "mathtype-offline-preview.svg");
        File.WriteAllText(
            previewSvg,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"360\" height=\"64\" viewBox=\"0 0 360 64\"><text x=\"4\" y=\"44\" font-size=\"36\">VisualTeX MathType long formula</text></svg>");
        var previewEmf = OfficeOlePreview.CreateVectorEmfFromSvg(previewSvg, 360, 64);

        var mathTypeProcessesBefore = SnapshotMathTypeProcessIds();

        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.Range? range = null;
        Word.OLEFormat? format = null;
        try
        {
            application = CreateWordApplication(visible: true);
            document = application.Documents.Open(path, ReadOnly: false, Visible: true);
            document.Activate();
            AssertEqual(1, document.InlineShapes.Count,
                "MathType product fixture must begin with one inline OLE equation.");
            Word.Range? leftProseRange = null;
            Word.Range? rightProseRange = null;
            Word.InlineShape? proseShape = null;
            try
            {
                proseShape = document.InlineShapes[1];
                leftProseRange = proseShape.Range.Duplicate;
                leftProseRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                leftProseRange.Text = "VTLEFT ";
                Release(proseShape);
                proseShape = document.InlineShapes[1];
                rightProseRange = proseShape.Range.Duplicate;
                rightProseRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                rightProseRange.Text = " VTRIGHT";
            }
            finally
            {
                Release(rightProseRange);
                Release(leftProseRange);
                Release(proseShape);
            }
            var service = new WordFormulaService(application);

            Console.WriteLine("[MathType product 1/6] Reading the genuine Word MathType OLE through WordFormulaService...");
            var firstSelection = ReadMathTypeProductSelection(service, document, ref shape, ref range);
            var sourceOleWidth = shape!.Width;
            var sourceOleHeight = shape.Height;
            var sourceOlePosition = ReadInlineOlePositionForAcceptance(shape);
            var sourceFragmentForPreview = MathTypeWordOpenXml.Read(shape!);
            var sourceWordPreview = sourceFragmentForPreview.PreviewWmf;
            var sourceNative = MathTypeOleStorage.ReadEquationNative(
                sourceFragmentForPreview.CompoundFile);
            var sourceHeaderLength = BitConverter.ToUInt16(sourceNative, 0);
            var sourceMtefLength = checked((int)BitConverter.ToUInt32(sourceNative, 8));
            var sourceMtef = new byte[sourceMtefLength];
            Buffer.BlockCopy(
                sourceNative,
                sourceHeaderLength,
                sourceMtef,
                0,
                sourceMtefLength);
            AssertTrue(
                MathTypeNativePreviewRenderer.TryRender(
                    sourceMtef,
                    artifactRoot,
                    out var sourceNativePreview),
                "MathType native renderer was unavailable while validating the source presentation scale.");
            float sourceNativeWidthPt;
            float sourceNativeHeightPt;
            int sourceNativeWordPosition;
            using (sourceNativePreview)
            {
                sourceNativeWidthPt = sourceNativePreview.WidthPt;
                sourceNativeHeightPt = sourceNativePreview.HeightPt;
                sourceNativeWordPosition = sourceNativePreview.WordPosition;
            }
            AssertEqual(
                FormulaOleContract.MathTypeOleMode,
                firstSelection.ObjectMode,
                "WordFormulaService did not expose the source as mathTypeOle.");
            var firstLatex = (firstSelection.Metadata?.Latex ?? string.Empty).Replace(" ", string.Empty);
            Console.WriteLine("  source LaTeX=" + firstLatex);
            if (string.IsNullOrWhiteSpace(compatibilityCase))
            {
                AssertTrue(
                    firstLatex.IndexOf("sqrt", StringComparison.OrdinalIgnoreCase) >= 0
                    && (firstLatex.IndexOf("p^2", StringComparison.Ordinal) >= 0
                        || firstLatex.IndexOf("p^{2}", StringComparison.Ordinal) >= 0)
                    && (firstLatex.IndexOf("q^2", StringComparison.Ordinal) >= 0
                        || firstLatex.IndexOf("q^{2}", StringComparison.Ordinal) >= 0),
                    $"VisualTeX read the wrong source from the genuine MathType OLE. LaTeX='{firstLatex}'.");
            }

            Console.WriteLine("[MathType product 2/6] Editing through ReplaceMathTypeOle while keeping MathType OLE...");
            const string firstEditLatex = @"\pi\theta\mathrm{e}^{\mathrm{i}\pi}+1=0";
            const string firstEditMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>π</mi><mi>θ</mi><msup><mi mathvariant=\"normal\">e</mi>"
                + "<mrow><mi mathvariant=\"normal\">i</mi><mi>π</mi></mrow></msup>"
                + "<mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>";
            var firstEdit = CreateMathTypeProductSession(
                firstSelection,
                firstEditLatex,
                exportWidth: 360,
                exportHeight: 64);
            service.ReplaceMathTypeOle(firstEdit, firstEditMathMl, previewEmf);
            Release(range);
            range = null;
            Release(shape);
            shape = null;
            AssertEqual(1, document.InlineShapes.Count,
                "Keeping MathType OLE duplicated or lost the Word equation.");
            shape = document.InlineShapes[1];
            format = shape.OLEFormat;
            AssertTrue(
                MathTypeOleInterop.IsMathTypeOle(shape),
                $"The replacement object is no longer MathType OLE. ProgID='{format.ProgID}'.");
            Console.WriteLine("  replacement ProgID=" + format.ProgID);
            // The presentation standard is the native MathType renderer for the
            // *new* MTEF, not the old equation's height and not VisualTeX's SVG.
            // Genuine MathType changes the outer box when the structure changes
            // (for example sqrt -> a simple linear expression) while preserving
            // its Full Size, fonts and script rules.
            var replacementFragmentForPreview = MathTypeWordOpenXml.Read(shape);
            var replacementNative = MathTypeOleStorage.ReadEquationNative(
                replacementFragmentForPreview.CompoundFile);
            var replacementHeaderLength = BitConverter.ToUInt16(replacementNative, 0);
            var replacementMtefLength = checked((int)BitConverter.ToUInt32(replacementNative, 8));
            var replacementMtef = new byte[replacementMtefLength];
            Buffer.BlockCopy(
                replacementNative,
                replacementHeaderLength,
                replacementMtef,
                0,
                replacementMtefLength);
            AssertTrue(
                MathTypeNativePreviewRenderer.TryRender(
                    replacementMtef,
                    artifactRoot,
                    out var expectedNativePreview),
                "MathType native renderer was unavailable while validating the product presentation.");
            using (expectedNativePreview)
            {
                var expectedNativeWmf = File.ReadAllBytes(expectedNativePreview.WmfPath);
                var actualInlinePosition = ReadInlineOlePositionForAcceptance(shape);
                var sourceWidthScale = sourceNativeWidthPt > 0
                    ? sourceOleWidth / sourceNativeWidthPt
                    : 1f;
                var sourceHeightScale = sourceNativeHeightPt > 0
                    ? sourceOleHeight / sourceNativeHeightPt
                    : 1f;
                var expectedWidthPt = expectedNativePreview.WidthPt * sourceWidthScale;
                var expectedHeightPt = expectedNativePreview.HeightPt * sourceHeightScale;
                var expectedWordPosition = (int)Math.Round(
                    expectedNativePreview.WordPosition * sourceHeightScale,
                    MidpointRounding.AwayFromZero);
                Console.WriteLine(
                    $"  MathType presentation: source={sourceOleWidth:0.0}x{sourceOleHeight:0.0}pt pos={sourceOlePosition}, "
                    + $"sourceNative={sourceNativeWidthPt:0.0}x{sourceNativeHeightPt:0.0}pt pos={sourceNativeWordPosition}, "
                    + $"scale={sourceWidthScale:0.###}x{sourceHeightScale:0.###}, "
                    + $"replacement={shape.Width:0.0}x{shape.Height:0.0}pt pos={actualInlinePosition}, "
                    + $"native target={expectedNativePreview.WidthPt:0.0}x{expectedNativePreview.HeightPt:0.0}pt pos={expectedNativePreview.WordPosition}, "
                    + $"scaled expected={expectedWidthPt:0.0}x{expectedHeightPt:0.0}pt pos={expectedWordPosition}.");
                AssertNear(
                    expectedWidthPt,
                    shape.Width,
                    0.6f,
                    "Word did not preserve the source MathType width scale for the edited equation.");
                AssertNear(
                    expectedHeightPt,
                    shape.Height,
                    0.6f,
                    "Word did not preserve the source MathType height scale for the edited equation.");
                AssertNear(
                    expectedWordPosition,
                    actualInlinePosition,
                    1.0f,
                    "Word did not preserve the source MathType inline baseline scale for the edited equation.");

                Range? suffixBaselineRange = null;
                try
                {
                    suffixBaselineRange = document.Range(
                        shape.Range.End,
                        Math.Min(document.Content.End, shape.Range.End + 8));
                    AssertNear(
                        0f,
                        suffixBaselineRange.Font.Position,
                        0.1f,
                        "MathType inline replacement leaked its baseline into the following prose.");
                }
                finally { Release(suffixBaselineRange); }

                var replacementPreviewEmf = ReadInlineShapeEnhancedMetafile(shape);
                var replacementWordPreview = replacementFragmentForPreview.PreviewWmf;
                AssertTrue(
                    !sourceWordPreview.SequenceEqual(replacementWordPreview),
                    "Keeping MathType OLE left Word's persisted media preview unchanged.");
                var replayDifference = MeasureEmfPixelDifference(
                    expectedNativeWmf,
                    replacementPreviewEmf);
                var persistedDifference = MeasureEmfPixelDifference(
                    expectedNativeWmf,
                    replacementWordPreview);
                Console.WriteLine(
                    $"  native presentation diff: Word replay={replayDifference:0.0000}, persisted WMF={persistedDifference:0.0000}.");
                // Word converts the persisted WMF to CF_ENHMETAFILE when an OLE
                // object is copied. That replay conversion can introduce a few
                // antialiasing pixels, especially when the original MathType OLE
                // carries a small user/document scale. The persisted WMF itself
                // must remain essentially identical to MathType's renderer; allow
                // only the extra conversion noise on Word's replay surface.
                AssertTrue(
                    replayDifference < 0.04 && persistedDifference < 0.01,
                    "The edited MathType OLE does not visually match MathType's native renderer.");
                File.WriteAllBytes(
                    Path.Combine(artifactRoot, "preview-native-expected.wmf"),
                    expectedNativeWmf);
                File.WriteAllBytes(
                    Path.Combine(artifactRoot, "preview-word-replay.emf"),
                    replacementPreviewEmf);
            }
            Release(format);
            format = null;

            Console.WriteLine("[MathType product 3/6] Saving, closing, reopening, and reading the edited MathType OLE again...");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            document.Activate();
            var reopenedService = new WordFormulaService(application);
            var reopenedSelection = ReadMathTypeProductSelection(
                reopenedService,
                document,
                ref shape,
                ref range);
            var reopenedLatex = (reopenedSelection.Metadata?.Latex ?? string.Empty).Replace(" ", string.Empty);
            Console.WriteLine("  reopened LaTeX=" + reopenedLatex);
            AssertTrue(
                reopenedLatex.IndexOf(@"\pi", StringComparison.Ordinal) >= 0
                && reopenedLatex.IndexOf(@"\theta", StringComparison.Ordinal) >= 0
                && reopenedLatex.IndexOf(@"\mathrm{e}", StringComparison.Ordinal) >= 0
                && reopenedLatex.IndexOf(@"\mathrm{i}", StringComparison.Ordinal) >= 0
                && reopenedLatex.IndexOf("+1=0", StringComparison.Ordinal) >= 0,
                $"Save/reopen lost the upright Euler MathType edit. LaTeX='{reopenedLatex}'.");

            Console.WriteLine("[MathType product 4/6] Performing a second VisualTeX edit after Word reopen...");
            const string secondEditLatex = @"\frac{m}{n}+z";
            const string secondEditMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mfrac><mi>m</mi><mi>n</mi></mfrac><mo>+</mo><mi>z</mi></mrow></math>";
            var secondEdit = CreateMathTypeProductSession(reopenedSelection, secondEditLatex);
            reopenedService.ReplaceMathTypeOle(secondEdit, secondEditMathMl, previewEmf);

            Console.WriteLine("[MathType product 5/6] Reopening a second time and proving the equation is still MathType-editable data...");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            document.Activate();
            var finalService = new WordFormulaService(application);
            var finalSelection = ReadMathTypeProductSelection(
                finalService,
                document,
                ref shape,
                ref range);
            var finalLatex = (finalSelection.Metadata?.Latex ?? string.Empty).Replace(" ", string.Empty);
            Console.WriteLine("  final LaTeX=" + finalLatex);
            AssertTrue(
                finalLatex.IndexOf(@"\frac{m}{n}", StringComparison.Ordinal) >= 0
                && finalLatex.IndexOf("+z", StringComparison.Ordinal) >= 0,
                $"The second MathType edit did not survive Word reopen. LaTeX='{finalLatex}'.");
            format = shape!.OLEFormat;
            AssertTrue(
                MathTypeOleInterop.IsMathTypeOle(shape),
                $"Second save/reopen changed the object away from MathType OLE. ProgID='{format.ProgID}'.");
            var finalProse = document.Content.Text ?? string.Empty;
            AssertEqual(
                1,
                finalProse.Split(new[] { "VTLEFT" }, StringSplitOptions.None).Length - 1,
                "Two MathType edits duplicated or removed the ordinary text before the inline equation.");
            AssertEqual(
                1,
                finalProse.Split(new[] { "VTRIGHT" }, StringSplitOptions.None).Length - 1,
                "Two MathType edits duplicated or removed the ordinary text after the inline equation.");

            if (!string.IsNullOrWhiteSpace(compatibilityCase)
                && compatibilityLatex is not null
                && compatibilityMathMl is not null)
            {
                Console.WriteLine(
                    $"[MathType product compatibility] Performing a third VisualTeX edit with {compatibilityCase}...");
                Release(format);
                format = null;
                var compatibilityEdit = CreateMathTypeProductSession(
                    finalSelection,
                    compatibilityLatex,
                    exportWidth: 360,
                    exportHeight: 64);
                finalService.ReplaceMathTypeOle(
                    compatibilityEdit,
                    compatibilityMathMl,
                    previewEmf);

                document.Save();
                document.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = application.Documents.Open(path, ReadOnly: false, Visible: false);
                document.Activate();
                finalService = new WordFormulaService(application);
                finalSelection = ReadMathTypeProductSelection(
                    finalService,
                    document,
                    ref shape,
                    ref range);
                var persistedMathMl = MathTypeOleStorage.ReadMathMl(shape!);
                AssertEqual(
                    MathTypeMtefCodec.SemanticSignature(compatibilityMathMl),
                    MathTypeMtefCodec.SemanticSignature(persistedMathMl),
                    $"Production MathType OLE save/reopen changed the {compatibilityCase} formula before any MathType server was activated.");
                Console.WriteLine(
                    $"  production reopen source={MathMlToLatexConverter.Convert(persistedMathMl).Trim()}.");
            }

            var mathTypeProcessesAfter = SnapshotMathTypeProcessIds();
            var startedMathTypeProcesses = mathTypeProcessesAfter.Except(mathTypeProcessesBefore).ToArray();
            AssertEqual(
                0,
                startedMathTypeProcesses.Length,
                "Production MathType read/edit/keep flow unexpectedly started MathType process(es): "
                + string.Join(", ", startedMathTypeProcesses));
            Console.WriteLine("  MathType processes started by production VisualTeX flow: none.");

            if (!string.IsNullOrWhiteSpace(compatibilityCase)
                && compatibilityMathMl is not null)
            {
                Release(format);
                format = shape!.OLEFormat;
                application.Visible = true;
                if (document.Windows.Count > 0)
                    document.Windows[1].Visible = true;
                document.Activate();
                var verifyNativeResave = string.Equals(
                    compatibilityCase,
                    "mixed-bigop-sizing",
                    StringComparison.OrdinalIgnoreCase);
                var beforeMathTypeWidth = shape.Width;
                var beforeMathTypeHeight = shape.Height;
                var mathTypeReadback = InvokeWordOwnedMathTypeEditor(
                    application,
                    format,
                    replacementLatex: null,
                    saveChanges: verifyNativeResave);
                File.WriteAllText(
                    Path.Combine(artifactRoot, $"mathtype-product-{compatibilityCase}-readback.xml"),
                    mathTypeReadback);
                Console.WriteLine(
                    "  MathType 7 read-back codepoints="
                    + string.Join(",", mathTypeReadback
                        .Where(character => !char.IsWhiteSpace(character))
                        .Select(character => $"U+{(int)character:X4}").Take(160)));
                AssertEqual(
                    MathTypeMtefCodec.SemanticSignature(compatibilityMathMl),
                    MathTypeMtefCodec.SemanticSignature(mathTypeReadback),
                    $"Installed MathType 7 could not read the production VisualTeX {compatibilityCase} OLE after save/reopen.");
                Console.WriteLine(
                    $"  MathType 7 read-back source={MathMlToLatexConverter.Convert(mathTypeReadback).Trim()}.");
                if (verifyNativeResave)
                {
                    Release(format);
                    format = null;
                    document.Save();
                    Release(shape);
                    shape = document.InlineShapes[1];
                    format = shape.OLEFormat;
                    Console.WriteLine(
                        $"  MathType native re-save geometry: before={beforeMathTypeWidth:0.0}x{beforeMathTypeHeight:0.0}pt, after={shape.Width:0.0}x{shape.Height:0.0}pt.");
                    AssertTrue(
                        shape.Height >= beforeMathTypeHeight * 0.75f
                        && shape.Height <= beforeMathTypeHeight * 1.35f,
                        "MathType native re-save changed the VisualTeX-edited OLE height dramatically, indicating incompatible MTEF size state.");
                    var nativeResavedMathMl = MathTypeOleStorage.ReadMathMl(shape);
                    AssertEqual(
                        MathTypeMtefCodec.SemanticSignature(compatibilityMathMl),
                        MathTypeMtefCodec.SemanticSignature(nativeResavedMathMl),
                        "MathType native re-save changed the mixed BigOp formula semantics.");
                }
            }

            Console.WriteLine("[MathType product 6/6] Real MathType OLE read/edit/keep/save/reopen/re-edit passed through production WordFormulaService without launching MathType.");
        }
        finally
        {
            Release(format);
            Release(range);
            Release(shape);
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

    private static void RunWordMathTypeDisplayLayoutAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var fixture = ResolveMathTypeNativeEditorFixture();
        if (!File.Exists(fixture))
            throw new FileNotFoundException("A genuine MathType display-layout fixture is required.", fixture);
        var path = Path.Combine(artifactRoot, $"VisualTeX-MathType-Display-{Guid.NewGuid():N}.docx");
        File.Copy(fixture, path, overwrite: false);
        var previewSvg = Path.Combine(artifactRoot, "mathtype-display-preview.svg");
        File.WriteAllText(
            previewSvg,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"96\" viewBox=\"0 0 240 96\"><text x=\"4\" y=\"64\" font-size=\"48\">display</text></svg>");
        var previewEmf = OfficeOlePreview.CreateVectorEmfFromSvg(previewSvg, 240, 96);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.Range? range = null;
        Word.Range? insertion = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabs = null;
        Word.TabStop? tab = null;
        Word.Field? mathTypeNumberField = null;
        MathTypeNativePreviewRenderer.Result? sourceNative = null;
        MathTypeNativePreviewRenderer.Result? targetNative = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            AssertEqual(1, document.InlineShapes.Count,
                "MathType display fixture must begin with one Equation.DSMT4 OLE.");
            shape = document.InlineShapes[1];

            // The genuine fixture is intentionally inline and contains ordinary
            // prose next to the OLE.  Remove only that surrounding prose in this
            // temporary copy before constructing a true MathType display row.
            paragraph = shape.Range.Paragraphs[1];
            insertion = document.Range(shape.Range.End, paragraph.Range.End - 1);
            if (insertion.End > insertion.Start) insertion.Delete();
            Release(insertion); insertion = null;
            Release(paragraph); paragraph = null;
            Release(shape); shape = document.InlineShapes[1];
            paragraph = shape.Range.Paragraphs[1];
            insertion = document.Range(paragraph.Range.Start, shape.Range.Start);
            if (insertion.End > insertion.Start) insertion.Delete();
            Release(insertion); insertion = null;
            Release(paragraph); paragraph = null;
            Release(shape); shape = document.InlineShapes[1];

            insertion = shape.Range.Duplicate;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            insertion.Text = "\t";
            Release(insertion); insertion = null;
            Release(shape); shape = document.InlineShapes[1];
            insertion = shape.Range.Duplicate;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertion.Text = "\t";
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            // Reproduce MathType's real right-numbered Word structure.  The
            // equation number is not ordinary "(2.4)" text: it is an empty-result
            // MACROBUTTON MTPlaceRef field whose field code owns the number and
            // nested MTEqn/MTSec fields.  Paragraph.Range.Text therefore exposes
            // U+0015, which is the exact control character that previously caused
            // VisualTeX to misclassify the row as inline.
            mathTypeNumberField = document.Fields.Add(
                insertion,
                Word.WdFieldType.wdFieldEmpty,
                " MACROBUTTON MTPlaceRef (2.4) \\* MERGEFORMAT ",
                PreserveFormatting: false);
            Release(insertion); insertion = null;
            Release(shape); shape = document.InlineShapes[1];

            paragraph = shape.Range.Paragraphs[1];
            format = paragraph.Format;
            // Genuine MathType numbered rows use center/right tab stops rather
            // than paragraph-center alignment.  Keep that real layout model in
            // the acceptance so a lost tab stop cannot be hidden by wdAlignCenter.
            format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            format.LeftIndent = 11f;
            format.RightIndent = 7f;
            format.FirstLineIndent = 0f;
            format.SpaceBefore = 5f;
            format.SpaceAfter = 8f;
            format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
            format.LineSpacing = 24f;
            format.KeepTogether = -1;
            tabs = format.TabStops;
            tabs.ClearAll();
            tab = tabs.Add(208f, Word.WdTabAlignment.wdAlignTabCenter, Word.WdTabLeader.wdTabLeaderSpaces);
            Release(tab); tab = null;
            tab = tabs.Add(415f, Word.WdTabAlignment.wdAlignTabRight, Word.WdTabLeader.wdTabLeaderSpaces);
            Release(tab); tab = null;

            var sourceFragment = MathTypeWordOpenXml.Read(shape);
            var sourceNativeBytes = MathTypeOleStorage.ReadEquationNative(sourceFragment.CompoundFile);
            var sourceHeader = BitConverter.ToUInt16(sourceNativeBytes, 0);
            var sourceLength = checked((int)BitConverter.ToUInt32(sourceNativeBytes, 8));
            var sourceMtef = new byte[sourceLength];
            Buffer.BlockCopy(sourceNativeBytes, sourceHeader, sourceMtef, 0, sourceLength);
            AssertTrue(
                MathTypeNativePreviewRenderer.TryRender(sourceMtef, artifactRoot, out sourceNative),
                "MathType native renderer was unavailable for display source geometry.");
            var sourceHeightScale = shape.Height / sourceNative!.HeightPt;

            var displayProbeText = paragraph.Range.Text ?? string.Empty;
            Console.WriteLine(
                "[MathType display probe] paragraph codepoints="
                + string.Join(",", displayProbeText.Select(character => $"U+{(int)character:X4}")));
            var service = new WordFormulaService(application);
            range = shape.Range.Duplicate;
            range.Select();
            var selection = service.ReadSelection();
            AssertEqual("block", selection.Metadata?.DisplayMode,
                "A MathType equation with only tab/number decoration was misclassified as inline.");
            const string editedLatex = @"\frac{a+b}{c+d}";
            const string editedMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mrow><mi>c</mi><mo>+</mo><mi>d</mi></mrow></mfrac></math>";
            var session = CreateMathTypeProductSession(selection, editedLatex, 240, 96);
            service.ReplaceMathTypeOle(session, editedMathMl, previewEmf);

            Release(range); range = null;
            Release(shape); shape = document.InlineShapes[1];
            Release(paragraph); paragraph = shape.Range.Paragraphs[1];
            Release(format); format = paragraph.Format;
            AssertEqual(Word.WdParagraphAlignment.wdAlignParagraphJustify, format.Alignment,
                "VisualTeX changed the original MathType numbered-row paragraph alignment.");
            AssertNear(11f, format.LeftIndent, 0.1f, "MathType display left indent changed.");
            AssertNear(7f, format.RightIndent, 0.1f, "MathType display right indent changed.");
            AssertNear(5f, format.SpaceBefore, 0.1f, "MathType display SpaceBefore changed.");
            AssertNear(8f, format.SpaceAfter, 0.1f, "MathType display SpaceAfter changed.");
            AssertEqual(Word.WdLineSpacing.wdLineSpaceExactly, format.LineSpacingRule,
                "MathType display line-spacing rule changed.");
            AssertNear(24f, format.LineSpacing, 0.1f, "MathType display exact line spacing changed.");
            var postEditParagraphText = paragraph.Range.Text ?? string.Empty;
            var postEditDocumentText = document.Content.Text ?? string.Empty;
            Console.WriteLine(
                "[MathType display post-edit] paragraph="
                + string.Join(",", postEditParagraphText.Select(character => $"U+{(int)character:X4}"))
                + " document="
                + string.Join(",", postEditDocumentText.Select(character => $"U+{(int)character:X4}")));
            var postEditHasMathTypeNumber = false;
            for (var fieldIndex = 1; fieldIndex <= paragraph.Range.Fields.Count; fieldIndex++)
            {
                Word.Field? field = null;
                Word.Range? fieldCode = null;
                try
                {
                    field = paragraph.Range.Fields[fieldIndex];
                    fieldCode = field.Code;
                    postEditHasMathTypeNumber |= (fieldCode.Text ?? string.Empty)
                        .IndexOf("MACROBUTTON MTPlaceRef", StringComparison.OrdinalIgnoreCase) >= 0;
                }
                finally
                {
                    Release(fieldCode);
                    Release(field);
                }
            }
            AssertTrue(postEditHasMathTypeNumber,
                "VisualTeX removed or detached MathType's MTPlaceRef equation-number field.");

            var sawCenterTab = false;
            var sawRightTab = false;
            Release(tabs); tabs = format.TabStops;
            for (var index = 1; index <= tabs.Count; index++)
            {
                Release(tab); tab = tabs[index];
                sawCenterTab |= Math.Abs(tab.Position - 208f) <= 0.5f
                    && tab.Alignment == Word.WdTabAlignment.wdAlignTabCenter;
                sawRightTab |= Math.Abs(tab.Position - 415f) <= 0.5f
                    && tab.Alignment == Word.WdTabAlignment.wdAlignTabRight;
            }
            AssertTrue(sawCenterTab && sawRightTab,
                "VisualTeX lost MathType display equation center/right tab stops.");

            var targetFragment = MathTypeWordOpenXml.Read(shape);
            var targetNativeBytes = MathTypeOleStorage.ReadEquationNative(targetFragment.CompoundFile);
            var targetHeader = BitConverter.ToUInt16(targetNativeBytes, 0);
            var targetLength = checked((int)BitConverter.ToUInt32(targetNativeBytes, 8));
            var targetMtef = new byte[targetLength];
            Buffer.BlockCopy(targetNativeBytes, targetHeader, targetMtef, 0, targetLength);
            AssertTrue(
                MathTypeNativePreviewRenderer.TryRender(targetMtef, artifactRoot, out targetNative),
                "MathType native renderer was unavailable for edited display geometry.");
            AssertNear(
                targetNative!.HeightPt * sourceHeightScale,
                shape.Height,
                0.75f,
                "Edited display OLE did not preserve the source MathType presentation scale.");
            var expectedPosition = (int)Math.Round(
                targetNative.WordPosition * sourceHeightScale,
                MidpointRounding.AwayFromZero);
            AssertNear(
                expectedPosition,
                ReadInlineOlePositionForAcceptance(shape),
                1f,
                "Edited display OLE baseline does not match the new native MathType preview.");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            Release(shape); shape = document.InlineShapes[1];
            Release(paragraph); paragraph = shape.Range.Paragraphs[1];
            Release(format); format = paragraph.Format;
            AssertEqual(Word.WdParagraphAlignment.wdAlignParagraphJustify, format.Alignment,
                "MathType numbered-row alignment did not survive Word save/reopen.");
            AssertNear(24f, format.LineSpacing, 0.1f,
                "MathType display line spacing did not survive Word save/reopen.");
            var reopenedHasMathTypeNumber = false;
            for (var fieldIndex = 1; fieldIndex <= paragraph.Range.Fields.Count; fieldIndex++)
            {
                Word.Field? field = null;
                Word.Range? fieldCode = null;
                try
                {
                    field = paragraph.Range.Fields[fieldIndex];
                    fieldCode = field.Code;
                    reopenedHasMathTypeNumber |= (fieldCode.Text ?? string.Empty)
                        .IndexOf("MACROBUTTON MTPlaceRef", StringComparison.OrdinalIgnoreCase) >= 0;
                }
                finally
                {
                    Release(fieldCode);
                    Release(field);
                }
            }
            AssertTrue(reopenedHasMathTypeNumber,
                "MathType MTPlaceRef equation number did not survive Word save/reopen in the formula paragraph.");

            // Reproduce the exact damaged state created by older VisualTeX builds:
            // the MathType OLE remains in one paragraph while the trailing tab +
            // MTPlaceRef field is pushed to the immediately following paragraph,
            // and the formula paragraph loses the center/right MathType tab stops.
            var healthyParagraphCount = document.Paragraphs.Count;
            insertion = document.Range(shape.Range.End, shape.Range.End);
            insertion.Text = "\r";
            Release(insertion); insertion = null;
            AssertEqual(healthyParagraphCount + 1, document.Paragraphs.Count,
                "Acceptance could not reproduce the legacy detached MTPlaceRef paragraph.");
            Release(shape); shape = document.InlineShapes[1];
            Release(paragraph); paragraph = shape.Range.Paragraphs[1];
            Release(format); format = paragraph.Format;
            Release(tabs); tabs = format.TabStops;
            tabs.ClearAll();

            var recoveryService = new WordFormulaService(application);
            Release(range); range = shape.Range.Duplicate;
            range.Select();
            var recoverySelection = recoveryService.ReadSelection();
            AssertEqual("block", recoverySelection.Metadata?.DisplayMode,
                "A legacy split MathType display equation was not recognized as block layout.");
            const string recoveryLatex = @"x^2+y^2";
            const string recoveryMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><msup><mi>y</mi><mn>2</mn></msup></math>";
            recoveryService.ReplaceMathTypeOle(
                CreateMathTypeProductSession(recoverySelection, recoveryLatex, 200, 72),
                recoveryMathMl,
                previewEmf);
            AssertEqual(healthyParagraphCount, document.Paragraphs.Count,
                "VisualTeX did not merge a legacy detached MathType MTPlaceRef number paragraph back into the formula row.");
            Release(range); range = null;
            Release(shape); shape = document.InlineShapes[1];
            Release(paragraph); paragraph = shape.Range.Paragraphs[1];
            Release(format); format = paragraph.Format;
            var recoveredPlaceRef = false;
            for (var fieldIndex = 1; fieldIndex <= paragraph.Range.Fields.Count; fieldIndex++)
            {
                Word.Field? field = null;
                Word.Range? fieldCode = null;
                try
                {
                    field = paragraph.Range.Fields[fieldIndex];
                    fieldCode = field.Code;
                    recoveredPlaceRef |= (fieldCode.Text ?? string.Empty)
                        .IndexOf("MACROBUTTON MTPlaceRef", StringComparison.OrdinalIgnoreCase) >= 0;
                }
                finally
                {
                    Release(fieldCode);
                    Release(field);
                }
            }
            AssertTrue(recoveredPlaceRef,
                "Legacy MathType MTPlaceRef field was not restored to the equation paragraph.");
            sawCenterTab = false;
            sawRightTab = false;
            Release(tabs); tabs = format.TabStops;
            for (var index = 1; index <= tabs.Count; index++)
            {
                Release(tab); tab = tabs[index];
                sawCenterTab |= Math.Abs(tab.Position - 208f) <= 0.5f
                    && tab.Alignment == Word.WdTabAlignment.wdAlignTabCenter;
                sawRightTab |= Math.Abs(tab.Position - 415f) <= 0.5f
                    && tab.Alignment == Word.WdTabAlignment.wdAlignTabRight;
            }
            AssertTrue(sawCenterTab && sawRightTab,
                "Legacy split-row recovery did not restore MathType's center/right tab stops.");

            // Repeat with a pure, unnumbered display row.  MathType commonly leaves
            // a leading centering tab even when no equation number is present; the
            // Flat OPC insertion must not leave an extra empty paragraph behind.
            insertion = document.Range(shape.Range.End, paragraph.Range.End - 1);
            if (insertion.End > insertion.Start) insertion.Delete();
            Release(insertion); insertion = null;
            var paragraphCountBeforeUnnumberedEdit = document.Paragraphs.Count;
            var unnumberedService = new WordFormulaService(application);
            Release(range); range = shape.Range.Duplicate;
            range.Select();
            var unnumberedSelection = unnumberedService.ReadSelection();
            AssertEqual("block", unnumberedSelection.Metadata?.DisplayMode,
                "A pure MathType display equation with only positioning tabs was misclassified as inline.");
            const string unnumberedLatex = @"x+y";
            const string unnumberedMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mi>y</mi></math>";
            unnumberedService.ReplaceMathTypeOle(
                CreateMathTypeProductSession(unnumberedSelection, unnumberedLatex, 160, 64),
                unnumberedMathMl,
                previewEmf);
            Release(range); range = null;
            Release(shape); shape = document.InlineShapes[1];
            Release(paragraph); paragraph = shape.Range.Paragraphs[1];
            Release(format); format = paragraph.Format;
            AssertEqual(paragraphCountBeforeUnnumberedEdit, document.Paragraphs.Count,
                "Editing an unnumbered MathType display equation left an extra paragraph behind.");
            AssertEqual(Word.WdParagraphAlignment.wdAlignParagraphJustify, format.Alignment,
                "Unnumbered MathType display alignment changed after VisualTeX edit.");
            AssertNear(24f, format.LineSpacing, 0.1f,
                "Unnumbered MathType display line spacing changed after VisualTeX edit.");
            AssertEqual(1, paragraph.Range.Fields.Count,
                "The MathType MTPlaceRef number field leaked into the unnumbered display row.");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            Release(shape); shape = document.InlineShapes[1];
            Release(paragraph); paragraph = shape.Range.Paragraphs[1];
            Release(format); format = paragraph.Format;
            AssertEqual(Word.WdParagraphAlignment.wdAlignParagraphJustify, format.Alignment,
                "Unnumbered MathType display alignment did not survive Word reopen.");
            AssertNear(24f, format.LineSpacing, 0.1f,
                "Unnumbered MathType display line spacing did not survive Word reopen.");
            Console.WriteLine(
                "[MathType display] numbered and unnumbered rows preserved centering, spacing, tabs, native geometry/baseline and numbering ownership through edit + reopen.");
        }
        finally
        {
            sourceNative?.Dispose();
            targetNative?.Dispose();
            Release(mathTypeNumberField);
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraph);
            Release(insertion);
            Release(range);
            Release(shape);
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

    private static string ResolveMathTypeNativeEditorFixture()
    {
        const string fixtureName =
            "VisualTeX-MathType7-NativeEditor-5f04f8b3545e444a824705446e314ba1.docx";
        foreach (var parentDepth in new[] { 5, 6, 7 })
        {
            var parts = new List<string> { AppContext.BaseDirectory };
            for (var index = 0; index < parentDepth; index++) parts.Add("..");
            parts.Add("artifacts");
            parts.Add("mathtype-native-editor");
            parts.Add(fixtureName);
            var candidate = Path.GetFullPath(Path.Combine(parts.ToArray()));
            if (File.Exists(candidate)) return candidate;
        }
        return Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", "..",
            "artifacts", "mathtype-native-editor", fixtureName));
    }

    private static OfficeSelection ReadMathTypeProductSelection(
        WordFormulaService service,
        Word.Document document,
        ref Word.InlineShape? shape,
        ref Word.Range? range)
    {
        Release(range);
        range = null;
        Release(shape);
        shape = null;
        AssertEqual(1, document.InlineShapes.Count,
            "The product MathType acceptance expected exactly one inline equation.");
        shape = document.InlineShapes[1];
        range = shape.Range.Duplicate;
        range.Select();
        var selection = service.ReadSelection();
        if (selection.Metadata is null)
            throw new InvalidDataException(
                "WordFormulaService recognized the MathType OLE but did not return editable metadata.");
        return selection;
    }

    private readonly struct EmfInkVerticalMetrics
    {
        internal EmfInkVerticalMetrics(double inkHeightRatio, double bottomWhitespaceRatio)
        {
            InkHeightRatio = inkHeightRatio;
            BottomWhitespaceRatio = bottomWhitespaceRatio;
        }

        internal double InkHeightRatio { get; }
        internal double BottomWhitespaceRatio { get; }
    }

    private static int ReadInlineOlePositionForAcceptance(Word.InlineShape shape)
    {
        Word.Range? shapeRange = null;
        Word.Range? objectCharacter = null;
        Word.Document? document = null;
        Word.Font? font = null;
        try
        {
            shapeRange = shape.Range;
            document = shapeRange.Document;
            for (var position = shapeRange.Start; position < shapeRange.End; position++)
            {
                Release(font);
                font = null;
                Release(objectCharacter);
                objectCharacter = document.Range(position, position + 1);
                if (!string.Equals(objectCharacter.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                font = objectCharacter.Font;
                var value = font.Position;
                return value == (int)Word.WdConstants.wdUndefined ? 0 : value;
            }
            throw new InvalidDataException(
                "MathType acceptance could not find the U+0001 OLE result character inside the EMBED field.");
        }
        finally
        {
            Release(font);
            Release(objectCharacter);
            Release(shapeRange);
            Release(document);
        }
    }

    private static EmfInkVerticalMetrics MeasureEmfInkVerticalMetrics(byte[] bytes)
    {
        using var bitmap = RenderEmf(bytes, 640, 240);
        var minY = bitmap.Height;
        var maxY = -1;
        for (var y = 0; y < bitmap.Height; y++)
        {
            for (var x = 0; x < bitmap.Width; x++)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R >= 245 && pixel.G >= 245 && pixel.B >= 245) continue;
                minY = Math.Min(minY, y);
                maxY = Math.Max(maxY, y);
            }
        }
        if (maxY < minY)
            throw new InvalidDataException("EMF preview contained no visible equation ink.");
        return new EmfInkVerticalMetrics(
            (maxY - minY + 1d) / bitmap.Height,
            (bitmap.Height - 1d - maxY) / bitmap.Height);
    }

    private static string DescribeEmfInkBounds(byte[] bytes)
    {
        using var bitmap = RenderEmf(bytes, 360, 128);
        var minX = bitmap.Width;
        var minY = bitmap.Height;
        var maxX = -1;
        var maxY = -1;
        var ink = 0;
        for (var y = 0; y < bitmap.Height; y++)
        {
            for (var x = 0; x < bitmap.Width; x++)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R > 245 && pixel.G > 245 && pixel.B > 245) continue;
                ink++;
                minX = Math.Min(minX, x);
                minY = Math.Min(minY, y);
                maxX = Math.Max(maxX, x);
                maxY = Math.Max(maxY, y);
            }
        }
        return ink == 0
            ? "empty"
            : $"{minX},{minY}-{maxX},{maxY}; ink={ink}";
    }

    private static double MeasureEmfPixelDifference(byte[] expected, byte[] actual)
    {
        using var expectedBitmap = RenderEmf(expected, 360, 128);
        using var actualBitmap = RenderEmf(actual, 360, 128);
        long difference = 0;
        var channelCount = (long)expectedBitmap.Width * expectedBitmap.Height * 3;
        for (var y = 0; y < expectedBitmap.Height; y++)
        {
            for (var x = 0; x < expectedBitmap.Width; x++)
            {
                var left = expectedBitmap.GetPixel(x, y);
                var right = actualBitmap.GetPixel(x, y);
                difference += Math.Abs(left.R - right.R);
                difference += Math.Abs(left.G - right.G);
                difference += Math.Abs(left.B - right.B);
            }
        }
        return difference / (channelCount * 255d);
    }

    private static Bitmap RenderEmf(byte[] bytes, int width, int height)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var metafile = new Metafile(stream);
        var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        graphics.DrawImage(metafile, new System.Drawing.Rectangle(0, 0, width, height));
        return bitmap;
    }

    private static byte[] ReadInlineShapeEnhancedMetafile(Word.InlineShape shape)
    {
        Exception? lastError = null;
        foreach (var delayMs in new[] { 0, 25, 50, 100, 150 })
        {
            if (delayMs > 0) Thread.Sleep(delayMs);
            try { return ReadInlineShapeEnhancedMetafileOnce(shape); }
            catch (COMException error)
            {
                lastError = error;
            }
        }
        throw new InvalidDataException(
            "Word clipboard stayed busy while reading the MathType preview.",
            lastError);
    }

    private static byte[] ReadInlineShapeEnhancedMetafileOnce(Word.InlineShape shape)
    {
        Word.Range? copiedRange = null;
        object? clipboardObject = null;
        try
        {
            copiedRange = shape.Range;
            copiedRange.Copy();
            var hr = OleGetClipboard(out var clipboardPointer);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);
            if (clipboardPointer == IntPtr.Zero)
                throw new InvalidDataException("Word copy returned no OLE IDataObject for the MathType preview.");
            try { clipboardObject = Marshal.GetObjectForIUnknown(clipboardPointer); }
            finally { Marshal.Release(clipboardPointer); }
            if (clipboardObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                throw new InvalidDataException("Word MathType copy does not expose IDataObject.");

            var request = new FORMATETC
            {
                cfFormat = 14,
                dwAspect = DVASPECT.DVASPECT_CONTENT,
                lindex = -1,
                ptd = IntPtr.Zero,
                tymed = TYMED.TYMED_ENHMF,
            };
            dataObject.GetData(ref request, out var medium);
            try
            {
                if (medium.tymed != TYMED.TYMED_ENHMF || medium.unionmember == IntPtr.Zero)
                    throw new InvalidDataException(
                        $"Word MathType preview did not expose CF_ENHMETAFILE; actual={medium.tymed}.");
                var byteCount = GetEnhMetaFileBits(medium.unionmember, 0, null);
                if (byteCount == 0 || byteCount > 32 * 1024 * 1024)
                    throw new InvalidDataException($"Unexpected MathType preview EMF size: {byteCount}.");
                var bytes = new byte[byteCount];
                if (GetEnhMetaFileBits(medium.unionmember, byteCount, bytes) != byteCount)
                    throw new InvalidDataException("GetEnhMetaFileBits returned an incomplete Word MathType preview.");
                return bytes;
            }
            finally { ReleaseStgMedium(ref medium); }
        }
        finally
        {
            Release(clipboardObject);
            Release(copiedRange);
        }
    }

    private static OfficeSessionDocument CreateMathTypeProductSession(
        OfficeSelection source,
        string latex,
        float exportWidth = 180,
        float exportHeight = 64)
    {
        var metadata = source.Metadata
            ?? throw new InvalidDataException("MathType source metadata is unavailable.");
        var lineId = metadata.Lines.FirstOrDefault()?.Id;
        if (string.IsNullOrWhiteSpace(lineId)) lineId = Guid.NewGuid().ToString("D");
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "edit",
            Host = "word",
            FormulaId = source.FormulaId ?? metadata.FormulaId,
            SourceDocumentId = source.DocumentId,
            SourceObjectId = source.ObjectId,
            Title = "MathType product round-trip acceptance",
            CodeFormat = "raw",
            DisplayMode = metadata.DisplayMode,
            ObjectMode = FormulaOleContract.MathTypeOleMode,
            Numbered = false,
            FontSizePt = metadata.FontSizePt ?? 12,
            OriginalMetadata = metadata,
            Lines = new List<FormulaLine>
            {
                new() { Id = lineId!, Latex = latex },
            },
            ExportResult = new OfficeExportDocument
            {
                Width = exportWidth,
                Height = exportHeight,
                Baseline = exportHeight * 0.75f,
            },
        };
    }
}
