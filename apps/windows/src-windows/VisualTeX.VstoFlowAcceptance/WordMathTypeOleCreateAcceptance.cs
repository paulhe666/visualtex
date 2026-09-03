using System.Diagnostics;
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
        RunWordMathTypeRightThenLeftCreateAcceptance(artifactRoot, emfPath);
        RunWordMathTypeDisplayCreateAcceptance(artifactRoot, emfPath);
        RunWordMathTypeInsertionRollbackAcceptance(artifactRoot, emfPath);
        RunWordMathTypeCreateEditCreateLifecycleAcceptance(artifactRoot, emfPath);
        RunWordMathTypeSequentialCreateStressAcceptance(artifactRoot, emfPath);
    }

    private static void RunWordMathTypeOleCreateStructureAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "mathtype-create-structure-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"96\" viewBox=\"0 0 240 96\"><text x=\"4\" y=\"64\" font-family=\"Times New Roman\" font-size=\"48\">x+1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 96);
        RunWordMathTypeInsertionRollbackAcceptance(artifactRoot, emfPath);
        RunWordMathTypeSequentialCreateStressAcceptance(artifactRoot, emfPath);
    }

    private static void RunWordMathTypeAlignedCreateAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "mathtype-aligned-create-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"760\" height=\"180\" viewBox=\"0 0 760 180\"><text x=\"4\" y=\"70\" font-family=\"Times New Roman\" font-size=\"34\">VisualTeX aligned MathType acceptance</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 760, 180);

        const string latex =
            @"\begin{aligned}"
            + @"\langle p_1,p_0\rangle &\leftarrow \operatorname{umul}(a,b)=ab &&\text{Double word product}\\"
            + @"p_0 &\leftarrow \operatorname{umullo}(a,b)=(ab)\bmod\beta &&\text{Low word}\\"
            + @"p_1 &\leftarrow \operatorname{umulhi}(a,b)=\left\lfloor\frac{ab}{\beta}\right\rfloor &&\text{High word.}"
            + @"\end{aligned}";
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mtable displaystyle=\"true\" columnalign=\"right left right left\" columnspacing=\"0em 2em 0em\" rowspacing=\"3pt\" data-visualtex-mtef-ruler-stops=\"1191,6881\">"
            + "<mtr>"
            + "<mtd><mo fence=\"false\" stretchy=\"false\">&#x27E8;</mo><msub><mi>p</mi><mn>1</mn></msub><mo>,</mo><msub><mi>p</mi><mn>0</mn></msub><mo fence=\"false\" stretchy=\"false\">&#x27E9;</mo></mtd>"
            + "<mtd><mi></mi><mo stretchy=\"false\">&#x2190;</mo><mi>umul</mi><mo stretchy=\"false\">(</mo><mi>a</mi><mo>,</mo><mi>b</mi><mo stretchy=\"false\">)</mo><mo>=</mo><mi>a</mi><mi>b</mi></mtd>"
            + "<mtd></mtd><mtd><mtext>Double word product</mtext></mtd>"
            + "</mtr>"
            + "<mtr>"
            + "<mtd><msub><mi>p</mi><mn>0</mn></msub></mtd>"
            + "<mtd><mi></mi><mo stretchy=\"false\">&#x2190;</mo><mi>umullo</mi><mo stretchy=\"false\">(</mo><mi>a</mi><mo>,</mo><mi>b</mi><mo stretchy=\"false\">)</mo><mo>=</mo><mo stretchy=\"false\">(</mo><mi>a</mi><mi>b</mi><mo stretchy=\"false\">)</mo><mo lspace=\"thickmathspace\" rspace=\"thickmathspace\">mod</mo><mi>&#x3B2;</mi></mtd>"
            + "<mtd></mtd><mtd><mtext>Low word</mtext></mtd>"
            + "</mtr>"
            + "<mtr>"
            + "<mtd><msub><mi>p</mi><mn>1</mn></msub></mtd>"
            + "<mtd><mi></mi><mo stretchy=\"false\">&#x2190;</mo><mi>umulhi</mi><mo stretchy=\"false\">(</mo><mi>a</mi><mo>,</mo><mi>b</mi><mo stretchy=\"false\">)</mo><mo>=</mo><mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">&#x230A;</mo><mfrac><mrow><mi>a</mi><mi>b</mi></mrow><mi>&#x3B2;</mi></mfrac><mo data-mjx-texclass=\"CLOSE\">&#x230B;</mo></mrow></mtd>"
            + "<mtd></mtd><mtd><mtext>High word.</mtext></mtd>"
            + "</mtr>"
            + "</mtable></math>";

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        Word.OLEFormat? format = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            range = document.Range(0, 0);
            range.Select();
            var service = new WordFormulaService(application);

            // This is the production path that previously threw
            // "VisualTeX generated invalid standalone MathType MTEF" before Word
            // received an OLE object. The exact three-row user fixture exercises
            // multiple alignment pairs/&&, prose cells, \bmod, function runs and
            // the floor/fraction structure in one insertion.
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: false,
                    latex: latex),
                mathMl,
                emfPath);

            AssertEqual(1, document.InlineShapes.Count,
                "Aligned MathType product create did not insert exactly one OLE object.");
            shape = document.InlineShapes[1];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                "Aligned MathType product create did not materialize Equation.DSMT4.");

            var immediateReadback = MathTypeOleStorage.ReadMathMl(shape);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(mathMl),
                MathTypeMtefCodec.SemanticSignature(immediateReadback),
                "Aligned MathType product create changed MTEF semantics before save.");
            var immediateLatex = MathMlToLatexConverter.Convert(immediateReadback);
            AssertTrue(
                immediateLatex.IndexOf(@"\begin{aligned}", StringComparison.Ordinal) >= 0
                && immediateLatex.Count(character => character == '&') == 9
                && immediateLatex.IndexOf("umullo", StringComparison.Ordinal) >= 0
                && immediateLatex.IndexOf("Double word product", StringComparison.Ordinal) >= 0
                && immediateLatex.IndexOf("Low word", StringComparison.Ordinal) >= 0
                && immediateLatex.IndexOf("High word.", StringComparison.Ordinal) >= 0,
                $"Aligned MathType direct MTEF readback lost alignment/text semantics: '{immediateLatex}'.");

            var alignedCompound = MathTypeOleStorage.CaptureCompoundFile(shape);
            var alignedEquationNative = MathTypeOleStorage.ReadEquationNative(alignedCompound);
            File.WriteAllBytes(
                Path.Combine(artifactRoot, "VisualTeX-MathType-Aligned-Multipoint-Equation-Native.bin"),
                alignedEquationNative);
            var editStableRewrite = MathTypeMtefCodec.RewriteEquationNative(
                alignedEquationNative,
                immediateReadback,
                inline: false);
            var editStableReadback = MathTypeMtefCodec.ReadEquationNativeMathMl(
                editStableRewrite.EquationNative);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(immediateReadback),
                MathTypeMtefCodec.SemanticSignature(editStableReadback),
                "Aligned MathType semantics changed after a VisualTeX read/rewrite cycle.");
            AssertTrue(
                editStableReadback.IndexOf(
                    "data-visualtex-mtef-ruler-stops=\"1191,6881\"",
                    StringComparison.Ordinal) >= 0,
                "Aligned MathType read/rewrite cycle lost its native RULER stops.");
            if (!MathTypeNativePreviewRenderer.TryRender(
                    editStableRewrite.Mtef,
                    artifactRoot,
                    out var editStablePreview))
                throw new InvalidDataException(
                    "MathType native renderer was unavailable for aligned edit-stability validation.");
            using (editStablePreview)
            {
                var editStableGeometry = MeasureMathTypeAlignedPreviewGeometry(
                    File.ReadAllBytes(editStablePreview.WmfPath),
                    artifactRoot,
                    "VisualTeX-MathType-Aligned-Multipoint-Rewrite-Preview.png");
                AssertEqual(3, editStableGeometry.RowCount,
                    "VisualTeX read/rewrite changed the aligned MathType native row count.");
                AssertTrue(
                    editStableGeometry.FirstAnchorSpread <= 2.0
                    && editStableGeometry.TextAnchorSpread <= 2.0,
                    $"VisualTeX read/rewrite lost native aligned geometry: first={editStableGeometry.FirstAnchorSpread:0.###}px, text={editStableGeometry.TextAnchorSpread:0.###}px.");
            }
            var alignedPreview = ReadInlineShapeEnhancedMetafile(shape);
            var visualAlignment = MeasureMathTypeAlignedPreviewGeometry(
                alignedPreview,
                artifactRoot,
                "VisualTeX-MathType-Aligned-Multipoint-Preview.png");
            Console.WriteLine(
                $"[MathType aligned visual geometry] rows={visualAlignment.RowCount}, "
                + $"firstAnchor={string.Join("/", visualAlignment.FirstAnchorX)}, "
                + $"textAnchor={string.Join("/", visualAlignment.TextAnchorX)}, "
                + $"spreads={visualAlignment.FirstAnchorSpread:0.###}/"
                + $"{visualAlignment.TextAnchorSpread:0.###}px, "
                + $"preview={visualAlignment.PngPath}.");
            AssertEqual(3, visualAlignment.RowCount,
                "Aligned MathType native preview did not contain exactly three visual equation rows.");
            AssertTrue(
                visualAlignment.FirstAnchorSpread <= 2.0,
                $"Aligned MathType first '&' right-side anchors are not visually aligned; spread={visualAlignment.FirstAnchorSpread:0.###}px.");
            AssertTrue(
                visualAlignment.TextAnchorSpread <= 2.0,
                $"Aligned MathType '&&' prose-column anchors are not visually aligned; spread={visualAlignment.TextAnchorSpread:0.###}px.");

            var path = Path.Combine(artifactRoot, "VisualTeX-MathType-Aligned-Multipoint-Create.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            var embeddings = ReadDocxOleEmbeddings(path);
            AssertEqual(1, embeddings.Count,
                "Aligned MathType product create did not persist exactly one OLE package part.");
            var persistedReadback = MathTypeOleStorage.ReadMathMl(embeddings[0]);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(mathMl),
                MathTypeMtefCodec.SemanticSignature(persistedReadback),
                "Aligned MathType product create changed MTEF semantics in the saved DOCX package.");

            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            AssertEqual(1, document.InlineShapes.Count,
                "Aligned MathType product create changed the OLE count after save/reopen.");
            Release(shape);
            shape = document.InlineShapes[1];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                "Reopened aligned MathType product create is no longer Equation.DSMT4.");
            var reopenedReadback = MathTypeOleStorage.ReadMathMl(shape);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(mathMl),
                MathTypeMtefCodec.SemanticSignature(reopenedReadback),
                "Reopened aligned MathType product create changed MTEF semantics.");
            var reopenedVisualAlignment = MeasureMathTypeAlignedPreviewGeometry(
                ReadInlineShapeEnhancedMetafile(shape),
                artifactRoot,
                "VisualTeX-MathType-Aligned-Multipoint-Reopened-Preview.png");
            AssertEqual(3, reopenedVisualAlignment.RowCount,
                "Reopened aligned MathType native preview did not contain exactly three visual rows.");
            AssertTrue(
                reopenedVisualAlignment.FirstAnchorSpread <= 2.0
                && reopenedVisualAlignment.TextAnchorSpread <= 2.0,
                $"Reopened aligned MathType visual anchors drifted: first={reopenedVisualAlignment.FirstAnchorSpread:0.###}px, text={reopenedVisualAlignment.TextAnchorSpread:0.###}px.");
            AssertTrue(
                Math.Abs(reopenedVisualAlignment.FirstAnchorX[0] - visualAlignment.FirstAnchorX[0]) <= 2
                && Math.Abs(reopenedVisualAlignment.TextAnchorX[0] - visualAlignment.TextAnchorX[0]) <= 2,
                "Word save/reopen moved the aligned MathType ruler anchors relative to the original native preview.");

            // Final product-level validation: let the installed MathType 7 OLE
            // editor itself open the exact object produced by VisualTeX. MathType's
            // MathML exporter flattens native fnMARKER tabs to U+0009 text nodes;
            // SemanticSignature reconstructs only that tab-pile form for comparison.
            application.Visible = true;
            document.Activate();
            format = shape.OLEFormat;
            var mathTypeReadback = InvokeWordOwnedMathTypeEditor(
                application,
                format,
                replacementLatex: null,
                saveChanges: false);
            File.WriteAllText(
                Path.Combine(artifactRoot, "VisualTeX-MathType-Aligned-NativeEditor-Readback.xml"),
                mathTypeReadback);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(mathMl),
                MathTypeMtefCodec.SemanticSignature(mathTypeReadback),
                "Installed MathType 7 changed the exact aligned product-create semantics.");
            Release(format);
            format = null;

            Console.WriteLine(
                "[MATHTYPE ALIGNED PRODUCT CREATE] Exact three-row multipoint aligned fixture passed WordFormulaService.InsertMathTypeOle, direct Equation.DSMT4 readback, save/package readback, Word reopen and installed MathType 7 native-editor readback with all nine '&' alignment boundaries preserved.");
        }
        finally
        {
            Release(format);
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

    private sealed class MathTypeAlignedPreviewGeometry
    {
        internal int RowCount { get; set; }
        internal int[] FirstAnchorX { get; set; } = Array.Empty<int>();
        internal int[] TextAnchorX { get; set; } = Array.Empty<int>();
        internal double FirstAnchorSpread { get; set; }
        internal double TextAnchorSpread { get; set; }
        internal string PngPath { get; set; } = string.Empty;
    }

    private static MathTypeAlignedPreviewGeometry MeasureMathTypeAlignedPreviewGeometry(
        byte[] preview,
        string artifactRoot,
        string pngFileName)
    {
        const int renderWidth = 1800;
        const int renderHeight = 600;
        using var bitmap = RenderEmf(preview, renderWidth, renderHeight);
        var pngPath = Path.Combine(artifactRoot, pngFileName);
        bitmap.Save(pngPath, System.Drawing.Imaging.ImageFormat.Png);

        static bool IsInk(System.Drawing.Color pixel) =>
            pixel.R < 242 || pixel.G < 242 || pixel.B < 242;

        var yInk = new int[bitmap.Height];
        for (var y = 0; y < bitmap.Height; y++)
        {
            for (var x = 0; x < bitmap.Width; x++)
            {
                if (IsInk(bitmap.GetPixel(x, y))) yInk[y]++;
            }
        }

        var rawBands = new List<(int Start, int End)>();
        var bandStart = -1;
        for (var y = 0; y < yInk.Length; y++)
        {
            if (yInk[y] > 0)
            {
                if (bandStart < 0) bandStart = y;
                continue;
            }
            if (bandStart < 0) continue;
            rawBands.Add((bandStart, y - 1));
            bandStart = -1;
        }
        if (bandStart >= 0) rawBands.Add((bandStart, yInk.Length - 1));

        var mergedBands = new List<(int Start, int End)>();
        foreach (var band in rawBands)
        {
            if (mergedBands.Count > 0
                && band.Start - mergedBands[mergedBands.Count - 1].End <= 10)
            {
                var previous = mergedBands[mergedBands.Count - 1];
                mergedBands[mergedBands.Count - 1] = (previous.Start, band.End);
            }
            else
            {
                mergedBands.Add(band);
            }
        }

        var rowBands = mergedBands
            .Where(band => band.End - band.Start + 1 >= 4)
            .Where(band => Enumerable.Range(band.Start, band.End - band.Start + 1)
                .Sum(y => yInk[y]) >= 40)
            .ToList();

        var rowSpans = new List<List<(int Start, int End)>>();
        foreach (var rowBand in rowBands)
        {
            var xInk = new bool[bitmap.Width];
            for (var x = 0; x < bitmap.Width; x++)
            {
                for (var y = rowBand.Start; y <= rowBand.End; y++)
                {
                    if (!IsInk(bitmap.GetPixel(x, y))) continue;
                    xInk[x] = true;
                    break;
                }
            }

            var rawSpans = new List<(int Start, int End)>();
            var start = -1;
            for (var x = 0; x < xInk.Length; x++)
            {
                if (xInk[x])
                {
                    if (start < 0) start = x;
                    continue;
                }
                if (start < 0) continue;
                rawSpans.Add((start, x - 1));
                start = -1;
            }
            if (start >= 0) rawSpans.Add((start, xInk.Length - 1));

            var mergedSpans = new List<(int Start, int End)>();
            foreach (var span in rawSpans)
            {
                if (mergedSpans.Count > 0
                    && span.Start - mergedSpans[mergedSpans.Count - 1].End <= 5)
                {
                    var previous = mergedSpans[mergedSpans.Count - 1];
                    mergedSpans[mergedSpans.Count - 1] = (previous.Start, span.End);
                }
                else if (span.End - span.Start + 1 >= 2)
                {
                    mergedSpans.Add(span);
                }
            }
            if (mergedSpans.Count == 0)
                throw new InvalidDataException("Aligned MathType preview row contained no horizontal ink spans.");
            rowSpans.Add(mergedSpans);
        }

        if (rowSpans.Count != 3)
        {
            return new MathTypeAlignedPreviewGeometry
            {
                RowCount = rowSpans.Count,
                PngPath = pngPath,
            };
        }

        var commonStarts = new List<int[]>();
        foreach (var candidate in rowSpans[0].Select(span => span.Start))
        {
            var matched = new[] { candidate, -1, -1 };
            var valid = true;
            for (var rowIndex = 1; rowIndex < rowSpans.Count; rowIndex++)
            {
                var nearest = rowSpans[rowIndex]
                    .Select(span => span.Start)
                    .OrderBy(value => Math.Abs(value - candidate))
                    .First();
                if (Math.Abs(nearest - candidate) > 2)
                {
                    valid = false;
                    break;
                }
                matched[rowIndex] = nearest;
            }
            if (valid) commonStarts.Add(matched);
        }

        // For this deliberately asymmetric fixture, the only visible component
        // starts shared by all three rows are the first '&' right-hand glyph
        // (the identical left arrow) and the '&&' prose-column start. This is a
        // stronger visual assertion than measuring whitespace: different left
        // expression widths cannot trick the detector into choosing another tab.
        var distinctCommon = commonStarts
            .OrderBy(values => values[0])
            .GroupBy(values => values[0])
            .Select(group => group.First())
            .ToArray();
        if (distinctCommon.Length != 2)
            throw new InvalidDataException(
                "Aligned MathType preview did not expose exactly two cross-row visual alignment anchors: "
                + string.Join(", ", distinctCommon.Select(values => string.Join("/", values))));

        var firstAnchor = distinctCommon[0];
        var textAnchor = distinctCommon[1];
        static double Spread(IReadOnlyList<int> values) => values.Count == 0
            ? double.PositiveInfinity
            : values.Max() - values.Min();

        return new MathTypeAlignedPreviewGeometry
        {
            RowCount = rowBands.Count,
            FirstAnchorX = firstAnchor,
            TextAnchorX = textAnchor,
            FirstAnchorSpread = Spread(firstAnchor),
            TextAnchorSpread = Spread(textAnchor),
            PngPath = pngPath,
        };
    }

    private static void RunWordMathTypeRightThenLeftLiveAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "mathtype-right-left-live-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"96\" viewBox=\"0 0 240 96\"><text x=\"4\" y=\"64\" font-family=\"Times New Roman\" font-size=\"48\">x+1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 96);

        Word.Application? application = null;
        Word.Document? originalDocument = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        try
        {
            application = CreateWordApplication(visible: false);
            try { originalDocument = application.ActiveDocument; } catch { }
            document = application.Documents.Add();
            document.Content.Text = "VisualTeX MathType live right then left acceptance";
            var service = new WordFormulaService(application);

            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "right"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "Live right-then-left acceptance did not create the first MathType OLE.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Live right-then-left acceptance did not create the first MTPlaceRef.");
            shape = document.InlineShapes[1];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "right",
                "Live right-then-left first MathType row");

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"c+d",
                    mathTypeNumberPosition: "left"),
                SecondNumberedMathMl,
                emfPath);
            AssertEqual(2, document.InlineShapes.Count,
                "Live right-then-left acceptance failed to create the second MathType OLE.");
            AssertEqual(2, CountMathTypePlaceRefFields(document),
                "Live right-then-left acceptance lost or duplicated MTPlaceRef fields.");
            Release(shape);
            shape = document.InlineShapes[2];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "Live right-then-left second MathType row");

            var path = Path.Combine(
                artifactRoot,
                "VisualTeX-MathType-Live-Right-Then-Left.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            var embeddings = ReadDocxOleEmbeddings(path);
            AssertEqual(2, embeddings.Count,
                "Live right-then-left acceptance did not persist exactly two OLE package parts.");
            var signatures = embeddings
                .Select(MathTypeOleStorage.ReadMathMl)
                .Select(MathTypeMtefCodec.SemanticSignature)
                .ToList();
            AssertTrue(
                signatures.Contains(MathTypeMtefCodec.SemanticSignature(FirstNumberedMathMl))
                && signatures.Contains(MathTypeMtefCodec.SemanticSignature(SecondNumberedMathMl)),
                "Live right-then-left acceptance persisted the wrong MathType MTEF data.");
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            AssertEqual(2, document.InlineShapes.Count,
                "Live right-then-left acceptance changed the MathType OLE count after reopen.");
            Release(shape);
            shape = document.InlineShapes[1];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "right",
                "Reopened loaded-addin right-then-left first MathType row");
            Release(shape);
            shape = document.InlineShapes[2];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "Reopened loaded-addin right-then-left second MathType row");
            Console.WriteLine(
                "[MathType LIVE] In the already-running user Word process: right-numbered Equation.DSMT4 -> left-numbered Equation.DSMT4 passed; both native previews, MTPlaceRef rows and saved CFB/MTEF package parts are valid without reading the Word clipboard.");
        }
        finally
        {
            Release(shape);
            Release(range);
            try { originalDocument?.Activate(); } catch { }
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            Release(originalDocument);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordMathTypeLeftThenRightStabilityAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "mathtype-left-right-stability-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"360\" height=\"120\" viewBox=\"0 0 360 120\"><text x=\"4\" y=\"76\" font-family=\"Times New Roman\" font-size=\"48\">MathType stability</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 360, 120);
        const string quadraticMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>";
        const string eulerMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msup><mi>e</mi><mrow><mi>i</mi><mi>π</mi></mrow></msup><mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>";

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        Word.Window? hostWindow = null;
        Process? wordProcess = null;
        try
        {
            application = CreateWordApplication(visible: true);
            document = application.Documents.Add();
            document.Content.Text = "VisualTeX MathType real-environment left then right stability acceptance";
            hostWindow = application.ActiveWindow;
            var hwnd = new IntPtr(hostWindow.Hwnd);
            _ = GetWindowThreadProcessId(hwnd, out var processId);
            if (processId == 0)
                throw new InvalidOperationException("Could not resolve the real Word PID for MathType stability acceptance.");
            wordProcess = Process.GetProcessById(unchecked((int)processId));
            Console.WriteLine(
                $"[MathType stability] Word pid={wordProcess.Id} started={wordProcess.StartTime:O} responding={wordProcess.Responding}");

            try
            {
                object addInKey = "VisualTeX.WordVsto";
                var addIn = application.COMAddIns.Item(ref addInKey);
                try
                {
                    Console.WriteLine(
                        $"[MathType stability] installed VisualTeX add-in Connect={addIn.Connect} Description={addIn.Description}");
                }
                finally { Release(addIn); }
            }
            catch (Exception error)
            {
                Console.WriteLine($"[MathType stability] installed VisualTeX add-in inventory unavailable: {error.Message}");
            }

            var service = new WordFormulaService(application);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                    mathTypeNumberPosition: "left"),
                quadraticMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "Real-environment left-right stability acceptance did not create the first MathType OLE.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Real-environment left-right stability acceptance did not create the first MTPlaceRef.");
            shape = document.InlineShapes[1];
            AssertMathTypeDisplayRow(shape, "left", "Real-environment first left-numbered MathType row");

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"e^{i\pi}+1=0",
                    mathTypeNumberPosition: "right"),
                eulerMathMl,
                emfPath);
            AssertEqual(2, document.InlineShapes.Count,
                "Real-environment left-right stability acceptance did not create the second MathType OLE.");
            AssertEqual(2, CountMathTypePlaceRefFields(document),
                "Real-environment left-right stability acceptance lost or duplicated MTPlaceRef fields.");
            Release(shape);
            shape = document.InlineShapes[2];
            AssertMathTypeDisplayRow(shape, "right", "Real-environment second right-numbered MathType row");

            Console.WriteLine(
                "[MathType stability] both inserts returned; waiting 10 seconds for delayed OLE/clipboard callbacks before declaring success...");
            Thread.Sleep(TimeSpan.FromSeconds(10));
            wordProcess.Refresh();
            AssertTrue(wordProcess.Responding,
                "Word became non-responsive after left-numbered -> right-numbered MathType insertion.");

            var cpuBefore = wordProcess.TotalProcessorTime;
            Thread.Sleep(TimeSpan.FromSeconds(5));
            wordProcess.Refresh();
            var cpuDelta = (wordProcess.TotalProcessorTime - cpuBefore).TotalMilliseconds;
            Console.WriteLine(
                $"[MathType stability] delayed window responding={wordProcess.Responding} cpuDelta5s={cpuDelta:0.0}ms");
            AssertTrue(wordProcess.Responding,
                "Word became non-responsive during the delayed MathType stability window.");
            AssertTrue(cpuDelta < 2500,
                $"Word entered a sustained CPU loop after MathType insertion (5s CPU delta={cpuDelta:0.0}ms).");

            AssertEqual(2, document.InlineShapes.Count,
                "Word changed the MathType OLE count during the delayed stability window.");
            AssertEqual(2, CountMathTypePlaceRefFields(document),
                "Word changed MTPlaceRef fields during the delayed stability window.");
            Console.WriteLine(
                "[MathType stability] REAL loaded-addin left-numbered quadratic -> right-numbered Euler insertion remained responsive and CPU-idle for 15 seconds after both inserts.");
        }
        finally
        {
            wordProcess?.Dispose();
            Release(hostWindow);
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
            if (MathTypeOleStorage.TryCaptureCompoundFileFromWordOpenXml(
                    shape,
                    out var immediateCompoundFile))
            {
                var immediateReadback = MathTypeOleStorage.ReadMathMl(immediateCompoundFile);
                AssertEqual(
                    MathTypeMtefCodec.SemanticSignature(FractionMathMl),
                    MathTypeMtefCodec.SemanticSignature(immediateReadback),
                    "Standalone MathType inline create changed formula semantics.");
            }
            else
            {
                Console.WriteLine(
                    "[MathType create] Live Word deferred the OLE package from Range.WordOpenXML; semantic validation is deferred to the saved DOCX package without using the clipboard.");
            }

            var path = Path.Combine(artifactRoot, "VisualTeX-MathType-Create-Inline.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            var savedEmbeddings = ReadDocxOleEmbeddings(path);
            AssertEqual(1, savedEmbeddings.Count,
                "Standalone MathType inline create did not persist exactly one OLE package part.");
            var savedReadback = MathTypeOleStorage.ReadMathMl(savedEmbeddings[0]);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(FractionMathMl),
                MathTypeMtefCodec.SemanticSignature(savedReadback),
                "Standalone MathType inline create changed in the saved DOCX package.");
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            Release(shape);
            shape = document.InlineShapes[1];
            AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                "Standalone MathType inline create lost its ProgID after Word reopen.");
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

    private static void RunWordMathTypeRightThenLeftCreateAcceptance(
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
            document.Content.Text = "MathType right then left acceptance";
            var service = new WordFormulaService(application);

            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "right"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "Right-then-left acceptance did not create the first MathType OLE.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Right-then-left acceptance did not create the first MTPlaceRef.");
            shape = document.InlineShapes[1];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "right",
                "Right-then-left first MathType row");
            AssertWordMathTypePreviewVisible(
                shape,
                "Right-then-left first MathType row",
                FirstNumberedMathMl,
                inline: false,
                artifactRoot);

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"c+d",
                    mathTypeNumberPosition: "left"),
                SecondNumberedMathMl,
                emfPath);
            AssertEqual(2, document.InlineShapes.Count,
                "Right-then-left acceptance failed to create the second MathType OLE.");
            AssertEqual(2, CountMathTypePlaceRefFields(document),
                "Right-then-left acceptance lost or duplicated MTPlaceRef fields.");
            Release(shape);
            shape = document.InlineShapes[2];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "Right-then-left second MathType row");
            AssertWordMathTypePreviewVisible(
                shape,
                "Right-then-left second MathType row",
                SecondNumberedMathMl,
                inline: false,
                artifactRoot);

            var path = Path.Combine(
                artifactRoot,
                "VisualTeX-MathType-Right-Then-Left.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            var embeddings = ReadDocxOleEmbeddings(path);
            AssertEqual(2, embeddings.Count,
                "Right-then-left acceptance did not persist exactly two OLE package parts.");
            var signatures = embeddings
                .Select(MathTypeOleStorage.ReadMathMl)
                .Select(MathTypeMtefCodec.SemanticSignature)
                .ToList();
            AssertTrue(
                signatures.Contains(MathTypeMtefCodec.SemanticSignature(FirstNumberedMathMl))
                && signatures.Contains(MathTypeMtefCodec.SemanticSignature(SecondNumberedMathMl)),
                "Right-then-left acceptance persisted the wrong MathType MTEF data.");

            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            AssertEqual(2, document.InlineShapes.Count,
                "Right-then-left MathType OLE count changed after reopen.");
            AssertEqual(2, CountMathTypePlaceRefFields(document),
                "Right-then-left MTPlaceRef count changed after reopen.");
            Release(shape);
            shape = document.InlineShapes[1];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "right",
                "Reopened right-then-left first row");
            Release(shape);
            shape = document.InlineShapes[2];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "Reopened right-then-left second row");
            Console.WriteLine(
                "[MathType real sequence] right-numbered create -> left-numbered create passed in the same Word process; both Equation.DSMT4 CFB/MTEF packages and MTPlaceRef rows survived save + reopen.");
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
        Word.Selection? selection = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabs = null;
        Word.TabStop? tab = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = string.Empty;
            InsertNumberingHeading(
                application,
                document,
                level: 1,
                text: "Display create heading");
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Text = "display create acceptance";
            Release(range);
            range = null;
            // This scenario explicitly validates MathType's native heading-scope
            // state. Pin the format instead of inheriting the machine's current
            // continuous/heading default from HKCU.
            WordEquationNumbering.SetEquationNumberFormat(
                document,
                EquationNumberFormat.Heading1DotId);
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
                expectedNumberPosition: null,
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
                    latex: @"a+b",
                    mathTypeNumberPosition: "left"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(2, document.InlineShapes.Count,
                "First numbered MathType display create did not add exactly one OLE.");
            Release(shape);
            shape = document.InlineShapes[2];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "First left-numbered MathType display create");
            AssertEqual(
                "(1.1)",
                ReadMathTypeVisibleNumberForShape(shape),
                "First left-numbered MathType display create rendered the wrong visible number.");
            AssertWordMathTypePreviewVisible(
                shape,
                "First numbered MathType display create",
                FirstNumberedMathMl,
                inline: false,
                artifactRoot);
            AssertTrue(CountMathTypePlaceRefFields(document) == 1,
                "First numbered MathType display create did not create exactly one MTPlaceRef field.");
            AssertNativeMathTypeSectionBreak(
                document,
                expectedCount: 1,
                expectedChapter: 1,
                expectedSection: 0);

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"c+d",
                    mathTypeNumberPosition: "right"),
                SecondNumberedMathMl,
                emfPath);
            AssertEqual(3, document.InlineShapes.Count,
                "Second numbered MathType display create did not add exactly one OLE.");
            Release(shape);
            shape = document.InlineShapes[3];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "right",
                "Second right-numbered MathType display create after a left-numbered row");
            AssertEqual(
                "(1.2)",
                ReadMathTypeVisibleNumberForShape(shape),
                "Second right-numbered MathType display create rendered the wrong visible number.");
            AssertWordMathTypePreviewVisible(
                shape,
                "Second numbered MathType display create",
                SecondNumberedMathMl,
                inline: false,
                artifactRoot);
            AssertTrue(CountMathTypePlaceRefFields(document) == 2,
                "Second numbered MathType display create did not preserve/clone MTPlaceRef numbering.");
            AssertNativeMathTypeSectionBreak(
                document,
                expectedCount: 1,
                expectedChapter: 1,
                expectedSection: 0);
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
            document.Activate();
            AssertEqual(3, document.InlineShapes.Count,
                "MathType display creates changed object count after Word reopen.");
            AssertTrue(CountMathTypePlaceRefFields(document) == 2,
                "MathType MTPlaceRef fields did not survive Word reopen.");
            AssertNativeMathTypeSectionBreak(
                document,
                expectedCount: 1,
                expectedChapter: 1,
                expectedSection: 0);
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
                    expectedNumberPosition: index == 1
                        ? null
                        : index == 2
                            ? "left"
                            : "right",
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
                "[MathType create] Unnumbered + left-then-right numbered Equation.DSMT4 inserts completed without mutating MathType's global number-side document state; native MTPlaceRef numbering, template inheritance and save/reopen passed.");
        }
        finally
        {
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraph);
            Release(selection);
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

    private static void RunWordMathTypeInsertionRollbackAcceptance(
        string artifactRoot,
        string emfPath)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        var previousFailureStage = Environment.GetEnvironmentVariable(
            "VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE");
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "MathType rollback acceptance";
            var service = new WordFormulaService(application);
            var paragraphCountBefore = document.Paragraphs.Count;

            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            Environment.SetEnvironmentVariable(
                "VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE",
                "after-flat-opc");
            var failed = false;
            try
            {
                service.InsertMathTypeOle(
                    CreateMathTypeCreateSession(
                        displayMode: "block",
                        numbered: true,
                        latex: @"a+b",
                        mathTypeNumberPosition: "right"),
                    FirstNumberedMathMl,
                    emfPath);
            }
            catch (InvalidOperationException error)
            {
                failed = true;
                AssertTrue(
                    error.Message.IndexOf("insert-flat-opc", StringComparison.Ordinal) >= 0
                    && error.Message.IndexOf("8007000E", StringComparison.OrdinalIgnoreCase) >= 0,
                    "Injected MathType insertion failure did not preserve stage/HRESULT diagnostics.");
            }
            finally
            {
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE",
                    previousFailureStage);
            }
            AssertTrue(failed,
                "MathType rollback acceptance did not inject the expected failure.");
            AssertEqual(0, document.InlineShapes.Count,
                "Failed MathType create left a partial Equation.DSMT4 object behind.");
            AssertEqual(0, CountMathTypePlaceRefFields(document),
                "Failed MathType create left an orphan MTPlaceRef field behind.");
            AssertEqual(paragraphCountBefore, document.Paragraphs.Count,
                "Failed MathType create left an extra display paragraph behind.");
            AssertNativeMathTypeSectionBreak(document, 0);

            // Inject a failure after the complete right-side MTPlaceRef has already
            // been copied atomically to the left, but before the temporary right
            // field is removed. Rollback must clear the whole half-relocated row so
            // the next insertion starts from the original document state.
            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            Environment.SetEnvironmentVariable(
                "VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE",
                "left-relocate-after-copy");
            failed = false;
            try
            {
                service.InsertMathTypeOle(
                    CreateMathTypeCreateSession(
                        displayMode: "block",
                        numbered: true,
                        latex: @"c+d",
                        mathTypeNumberPosition: "left"),
                    SecondNumberedMathMl,
                    emfPath);
            }
            catch (InvalidOperationException error)
            {
                failed = true;
                AssertTrue(
                    error.Message.IndexOf("relocate-left-mathtype-field-scaffold", StringComparison.Ordinal) >= 0
                    && error.Message.IndexOf("800A1710", StringComparison.OrdinalIgnoreCase) >= 0,
                    "Injected left MathType relocation failure did not preserve stage/HRESULT diagnostics.");
            }
            finally
            {
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE",
                    previousFailureStage);
            }
            AssertTrue(failed,
                "MathType left-relocation rollback acceptance did not inject the expected failure.");
            AssertEqual(0, document.InlineShapes.Count,
                "Left-relocation failure left a partial Equation.DSMT4 object behind.");
            AssertEqual(0, document.Fields.Count,
                "Left-relocation failure left duplicate/orphan MathType number fields behind.");
            AssertEqual(paragraphCountBefore, document.Paragraphs.Count,
                "Left-relocation failure left an extra display paragraph behind.");

            // A retry in the same Word document/process must succeed after the
            // rollback; the exception must not poison the left-number scaffold.
            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "left"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "MathType create could not recover after a rolled-back failure.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "MathType retry after rollback produced the wrong number count.");

            // Model a document saved by an older build after PasteSpecial failed:
            // remove only the OLE and leave its MTPlaceRef/tabs as an orphan row.
            shape = document.InlineShapes[1];
            var orphanStart = shape.Range.Paragraphs[1].Range.Start;
            shape.Delete();
            Release(shape);
            shape = null;
            AssertEqual(0, document.InlineShapes.Count,
                "Legacy-orphan setup still contains a MathType OLE.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Legacy-orphan setup did not leave exactly one MTPlaceRef.");
            Release(range);
            range = document.Range(orphanStart, orphanStart);
            range.Select();
            var captured = service.ReadSelection();
            var retrySession = CreateMathTypeCreateSession(
                displayMode: "block",
                numbered: true,
                latex: @"c+d",
                mathTypeNumberPosition: "left");
            retrySession.SourceDocumentId = captured.DocumentId;
            retrySession.SourceObjectId = captured.ObjectId;
            service.InsertMathTypeOle(retrySession, SecondNumberedMathMl, emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "VisualTeX did not recover an old number-only MathType failure row.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Recovering an old number-only MathType row duplicated MTPlaceRef.");
            shape = document.InlineShapes[1];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "Recovered legacy MathType orphan row");
            Console.WriteLine(
                "[MathType rollback] injected E_OUTOFMEMORY and half-relocated left-number failures rolled back the complete MathType row + paragraph/section state; same-process left retry and orphan-row recovery passed.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE",
                previousFailureStage);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
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

    private static void RunWordMathTypeCreateEditCreateLifecycleAcceptance(
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
            document.Content.Text = "MathType lifecycle acceptance";
            var service = new WordFormulaService(application);

            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "right"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "MathType lifecycle first create did not insert one OLE.");

            Release(shape);
            shape = document.InlineShapes[1];
            shape.Range.Select();
            var firstSelection = service.ReadSelection();
            AssertTrue(firstSelection.Metadata is not null,
                "MathType lifecycle edit did not read source metadata.");
            AssertTrue(firstSelection.Metadata!.Numbered,
                "A numbered MathType display equation was read back as unnumbered.");
            AssertEqual("block", firstSelection.Metadata.DisplayMode,
                "A numbered MathType display equation was not read back as block display.");
            AssertEqual("right", service.GetMathTypeNumberPositionForRange(firstSelection.ObjectId),
                "A right-numbered MathType display equation was not read back on the right.");

            service.ReplaceMathTypeOle(
                CreateMathTypeEditSession(
                    firstSelection,
                    @"c+d",
                    mathTypeNumberPosition: "right"),
                SecondNumberedMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "MathType lifecycle edit changed the OLE count.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "MathType lifecycle edit lost or duplicated its native number.");

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "left"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(2, document.InlineShapes.Count,
                "MathType lifecycle second create did not insert a second OLE.");

            Release(shape);
            shape = document.InlineShapes[2];
            shape.Range.Select();
            var secondSelection = service.ReadSelection();
            AssertTrue(secondSelection.Metadata is not null && secondSelection.Metadata.Numbered,
                "The second numbered MathType display equation was read back as unnumbered.");
            AssertEqual("left", service.GetMathTypeNumberPositionForRange(secondSelection.ObjectId),
                "A left-numbered MathType display equation was not read back on the left.");
            // Reading the second equation without replacing it models an editor
            // open/cancel cycle. No Word write is allowed on cancellation.

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            var capturedCreateSelection = service.ReadSelection();
            var thirdCreate = CreateMathTypeCreateSession(
                displayMode: "block",
                numbered: true,
                latex: @"a+b",
                mathTypeNumberPosition: "right");
            thirdCreate.SourceDocumentId = capturedCreateSelection.DocumentId;
            thirdCreate.SourceObjectId = capturedCreateSelection.ObjectId;

            // Simulate Word moving the live Selection while the external editor is
            // in front. The create must still use the captured empty paragraph and
            // the numbering template nearest to that captured position.
            Release(shape);
            shape = document.InlineShapes[1];
            shape.Range.Select();
            service.InsertMathTypeOle(thirdCreate, FirstNumberedMathMl, emfPath);
            AssertEqual(3, document.InlineShapes.Count,
                "MathType lifecycle third create failed after create/edit/cancel state transitions.");
            AssertEqual(3, CountMathTypePlaceRefFields(document),
                "MathType lifecycle third create lost or duplicated MTPlaceRef fields.");
            Release(shape);
            shape = document.InlineShapes[3];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "right",
                "MathType lifecycle third create");

            var path = Path.Combine(artifactRoot, "VisualTeX-MathType-Create-Edit-Create.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            AssertEqual(3, document.InlineShapes.Count,
                "MathType lifecycle OLE count changed after reopen.");
            AssertEqual(3, CountMathTypePlaceRefFields(document),
                "MathType lifecycle numbering count changed after reopen.");
            Console.WriteLine(
                "[MathType lifecycle] create -> numbered edit/readback -> create -> edit/cancel -> captured-position third create + save/reopen passed.");
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

    private static OfficeSessionDocument CreateMathTypeEditSession(
        OfficeSelection source,
        string latex,
        string mathTypeNumberPosition)
    {
        var metadata = source.Metadata
            ?? throw new InvalidDataException("MathType lifecycle source metadata is unavailable.");
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
            Title = "MathType lifecycle edit acceptance",
            CodeFormat = "latex",
            DisplayMode = metadata.DisplayMode,
            ObjectMode = FormulaOleContract.MathTypeOleMode,
            Numbered = metadata.Numbered,
            MathTypeNumberPosition = mathTypeNumberPosition,
            FontSizePt = metadata.FontSizePt ?? 12,
            OriginalMetadata = metadata,
            Lines = new List<FormulaLine>
            {
                new() { Id = lineId!, Latex = latex },
            },
            ExportResult = new OfficeExportDocument
            {
                Width = 240,
                Height = 96,
                Baseline = 72,
            },
        };
    }

    private static void RunWordMathTypeSequentialCreateStressAcceptance(
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
            document.Content.Text = "MathType sequential create stress";
            var service = new WordFormulaService(application);
            const int targetCount = 24;
            for (var index = 1; index <= targetCount; index++)
            {
                Release(range);
                range = document.Range(document.Content.End - 1, document.Content.End - 1);
                range.Select();
                var side = index % 2 == 0 ? "right" : "left";
                service.InsertMathTypeOle(
                    CreateMathTypeCreateSession(
                        displayMode: "block",
                        numbered: true,
                        latex: @"a+b",
                        mathTypeNumberPosition: side),
                    FirstNumberedMathMl,
                    emfPath);
                AssertEqual(index, document.InlineShapes.Count,
                    $"MathType sequential create #{index} changed the OLE count unexpectedly.");
                AssertEqual(index, CountMathTypePlaceRefFields(document),
                    $"MathType sequential create #{index} lost or duplicated MTPlaceRef fields.");
                Release(shape);
                shape = document.InlineShapes[index];
                AssertMathTypeDisplayRow(
                    shape,
                    expectedNumberPosition: side,
                    $"MathType sequential create #{index}");
            }

            var path = Path.Combine(artifactRoot, "VisualTeX-MathType-Sequential-Stress.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            AssertEqual(targetCount, document.InlineShapes.Count,
                "MathType sequential create stress changed OLE count after reopen.");
            AssertEqual(targetCount, CountMathTypePlaceRefFields(document),
                "MathType sequential create stress changed MTPlaceRef count after reopen.");
            for (var index = 1; index <= targetCount; index++)
            {
                Release(shape);
                shape = document.InlineShapes[index];
                var side = index % 2 == 0 ? "right" : "left";
                AssertMathTypeDisplayRow(
                    shape,
                    expectedNumberPosition: side,
                    $"Reopened MathType sequential create #{index}");
            }
            Console.WriteLine(
                $"[MathType create stress] {targetCount} consecutive numbered Equation.DSMT4 inserts with alternating left/right numbering passed in one Word process + save/reopen.");
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
        string? expectedNumberPosition,
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
        Word.Fields? nestedFields = null;
        Word.Field? nestedField = null;
        Word.Range? nestedCode = null;
        Word.Range? fieldResult = null;
        Word.Range? separator = null;
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

            var paragraphRange = paragraph.Range;
            var paragraphText = paragraphRange.Text ?? string.Empty;
            var expectNumber = expectedNumberPosition is not null;

            fields = paragraphRange.Fields;
            var sawPlaceRef = false;
            var placeRefStart = -1;
            var placeRefEnd = -1;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(fieldResult);
                fieldResult = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var placeRefCodeText = code.Text ?? string.Empty;
                if (placeRefCodeText.IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                AssertTrue(
                    placeRefCodeText.Length > 0
                    && !char.IsWhiteSpace(placeRefCodeText[placeRefCodeText.Length - 1]),
                    context + " has trailing whitespace after the visible MathType equation number.");
                sawPlaceRef = true;
                Release(nestedCode);
                nestedCode = null;
                Release(nestedField);
                nestedField = null;
                Release(nestedFields);
                nestedFields = code.Fields;
                AssertTrue(
                    nestedFields.Count >= 2,
                    context + " has an MTPlaceRef whose nested MathType SEQ fields escaped the outer field code range.");
                var hasHiddenMteqn = false;
                var hasVisibleMteqn = false;
                for (var nestedIndex = 1; nestedIndex <= nestedFields.Count; nestedIndex++)
                {
                    Release(nestedCode);
                    nestedCode = null;
                    Release(nestedField);
                    nestedField = nestedFields[nestedIndex];
                    nestedCode = nestedField.Code;
                    var nestedText = nestedCode.Text ?? string.Empty;
                    if (nestedText.IndexOf("SEQ MTEqn", StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    hasHiddenMteqn |= nestedText.IndexOf("\\h", StringComparison.OrdinalIgnoreCase) >= 0;
                    hasVisibleMteqn |= nestedText.IndexOf("\\c", StringComparison.OrdinalIgnoreCase) >= 0;
                }
                AssertTrue(
                    hasHiddenMteqn && hasVisibleMteqn,
                    context + " has an incomplete MTPlaceRef field tree (missing owned hidden/current MTEqn fields).");
                fieldResult = field.Result;
                placeRefStart = code.Start - 1;
                placeRefEnd = fieldResult.End + 1;
            }
            AssertEqual(expectNumber, sawPlaceRef,
                context + " has the wrong MathType MTPlaceRef numbering ownership.");

            if (!expectNumber)
            {
                AssertTrue(
                    paragraphText.StartsWith("\t\u0001", StringComparison.Ordinal),
                    context + " does not begin with Word's native tab + OLE sequence.");
            }
            else if (string.Equals(expectedNumberPosition, "left", StringComparison.Ordinal))
            {
                AssertTrue(placeRefEnd <= shapeRange.Start,
                    context + " does not place its MathType number before the equation.");
                // For an empty-result MTPlaceRef Word can report Result.End + 1 at
                // the following OLE boundary even though the preceding story
                // character is the real separator tab. Validate the physical
                // character immediately before the OLE, matching production.
                separator = shapeRange.Document.Range(
                    Math.Max(paragraphRange.Start, shapeRange.Start - 1),
                    shapeRange.Start);
                AssertEqual("\t", separator.Text,
                    context + " does not have MathType's number-to-equation center tab.");
            }
            else
            {
                AssertEqual("right", expectedNumberPosition,
                    context + " uses an unsupported MathType number position in the acceptance fixture.");
                AssertTrue(
                    paragraphText.StartsWith("\t\u0001", StringComparison.Ordinal),
                    context + " does not begin with Word's native center tab + OLE sequence.");
                AssertTrue(placeRefStart >= shapeRange.End,
                    context + " does not place its MathType number after the equation.");
                separator = shapeRange.Document.Range(shapeRange.End, placeRefStart);
                AssertTrue((separator.Text ?? string.Empty).IndexOf('\t') >= 0,
                    context + " does not have MathType's equation-to-number right tab.");
            }
            Release(paragraphRange);
        }
        finally
        {
            Release(paragraphStyle);
            paragraphStyle = null;
            paragraphStyleObject = null;
            Release(separator);
            Release(fieldResult);
            Release(nestedCode);
            Release(nestedField);
            Release(nestedFields);
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

    private static string ReadMathTypeVisibleNumberForShape(Word.InlineShape shape)
    {
        Word.Range? shapeRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return string.Empty;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            fields = paragraphRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (!MathTypeEquationReferences.IsMathTypePlaceRefCode(code.Text))
                    continue;
                return MathTypeEquationReferences.ReadVisibleNumberText(field);
            }
            return string.Empty;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static void AssertNativeMathTypeReference(
        Word.Document document,
        string expectedNumberText)
    {
        Word.Fields? fields = null;
        Word.Field? outer = null;
        Word.Range? outerCode = null;
        Word.Fields? nestedFields = null;
        Word.Field? nested = null;
        Word.Range? nestedCode = null;
        Word.Range? nestedResult = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? bookmarkRange = null;
        try
        {
            fields = document.Fields;
            string? bookmarkName = null;
            var sawReference = false;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(outerCode);
                outerCode = null;
                Release(outer);
                outer = fields[index];
                outerCode = outer.Code;
                var outerText = outerCode.Text ?? string.Empty;
                var match = System.Text.RegularExpressions.Regex.Match(
                    outerText,
                    @"\bGOTOBUTTON\s+(ZEqnNum\d{6})\b",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (!match.Success) continue;

                bookmarkName = match.Groups[1].Value;
                nestedFields = outerCode.Fields;
                AssertEqual(1, nestedFields.Count,
                    "Native MathType GOTOBUTTON reference must contain exactly one nested REF field.");
                nested = nestedFields[1];
                nestedCode = nested.Code;
                var normalizedNested = NormalizeFieldCodeForMathTypeAcceptance(
                    nestedCode.Text ?? string.Empty);
                AssertTrue(
                    normalizedNested.IndexOf(
                        "REF " + bookmarkName,
                        StringComparison.OrdinalIgnoreCase) >= 0,
                    "Native MathType reference does not target the same ZEqnNum bookmark as its GOTOBUTTON field.");
                AssertTrue(
                    normalizedNested.IndexOf("\\!", StringComparison.Ordinal) >= 0,
                    "Native MathType REF field is missing the \\! non-recursive-update switch.");
                AssertTrue(
                    normalizedNested.IndexOf("\\* Charformat", StringComparison.OrdinalIgnoreCase) >= 0,
                    "Native MathType REF field is missing MathType's Charformat switch.");
                nestedResult = nested.Result;
                var rawReferenceText = nestedResult.Text ?? string.Empty;
                AssertEqual(
                    expectedNumberText.Trim(),
                    rawReferenceText.Trim(),
                    "Native MathType REF field shows the wrong equation number.");
                AssertTrue(
                    rawReferenceText.Length == 0
                    || !char.IsWhiteSpace(rawReferenceText[rawReferenceText.Length - 1]),
                    "Native MathType REF field contains trailing whitespace after the equation number.");
                sawReference = true;
                break;
            }

            AssertTrue(sawReference && !string.IsNullOrWhiteSpace(bookmarkName),
                "Document does not contain a native MathType GOTOBUTTON + nested REF equation reference.");
            bookmarks = document.Bookmarks;
            AssertTrue(bookmarks.Exists(bookmarkName!),
                "Native MathType reference target bookmark does not exist.");
            bookmark = bookmarks[bookmarkName!];
            bookmarkRange = bookmark.Range;
            AssertEqual(
                expectedNumberText.Trim(),
                (bookmarkRange.Text ?? string.Empty).Trim(),
                "Native MathType ZEqnNum bookmark does not cover the visible equation number.");
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(nestedResult);
            Release(nestedCode);
            Release(nested);
            Release(nestedFields);
            Release(outerCode);
            Release(outer);
            Release(fields);
        }
    }

    private static void AssertNativeMathTypeSectionBreak(
        Word.Document document,
        int expectedCount,
        int expectedChapter = 1,
        int expectedSection = 1)
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
                AssertEqual(
                    $"SEQ MTSec \\r {expectedSection} \\h \\* MERGEFORMAT",
                    ordered[1],
                    $"MathType section break does not initialize MTSec to {expectedSection} using MathType's native field code.");
                AssertEqual(
                    $"SEQ MTChap \\r {expectedChapter} \\h \\* MERGEFORMAT",
                    ordered[2],
                    $"MathType section break does not initialize MTChap to {expectedChapter} using MathType's native field code.");

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
            if (expectedCount > 0)
            {
                var firstPlaceRefStart = FindFirstMathTypePlaceRefStartForAcceptance(document);
                var firstSectionBreakStart = FindFirstMathTypeSectionBreakStartForAcceptance(document);
                AssertTrue(firstPlaceRefStart >= 0 && firstSectionBreakStart >= 0,
                    "MathType create could not resolve the native section break / MTPlaceRef ordering.");
                AssertTrue(firstSectionBreakStart < firstPlaceRefStart,
                    "The default MathType chapter/section break must precede the first numbered equation.");
            }
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

    private static int FindFirstMathTypePlaceRefStartForAcceptance(Word.Document document)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        var best = int.MaxValue;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                best = Math.Min(best, Math.Max(document.Content.Start, code.Start - 1));
            }
            return best == int.MaxValue ? -1 : best;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static int FindFirstMathTypeSectionBreakStartForAcceptance(Word.Document document)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        var best = int.MaxValue;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                best = Math.Min(best, Math.Max(document.Content.Start, code.Start - 1));
            }
            return best == int.MaxValue ? -1 : best;
        }
        finally
        {
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
        string latex,
        string mathTypeNumberPosition = "right") =>
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
            MathTypeNumberPosition = mathTypeNumberPosition,
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
