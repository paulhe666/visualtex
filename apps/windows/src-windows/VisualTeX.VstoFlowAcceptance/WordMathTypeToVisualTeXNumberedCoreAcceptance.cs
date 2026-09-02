using System.Text;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeToVisualTeXNumberedCoreAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var tempRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(tempRoot);
        var svgPath = Path.Combine(tempRoot, $"{Guid.NewGuid():N}.svg");
        var pngPath = Path.Combine(tempRoot, $"{Guid.NewGuid():N}.png");
        string? emfPath = null;
        Word.Application? application = null;
        Word.Document? document = null;
        var previousPreviewDisable = Environment.GetEnvironmentVariable(
            "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW");
        var previousPerf = Environment.GetEnvironmentVariable(
            "VISUALTEX_NUMBERED_PERF_TRACE");
        try
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                "1");
            Environment.SetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE", "1");
            File.WriteAllText(
                svgPath,
                CreateFontAcceptanceSvg("Times New Roman", "SimSun"),
                new UTF8Encoding(false));
            WriteAcceptancePng(pngPath, "x=1", 260, 96);
            emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 260, 96);

            application = CreateWordApplication(visible: false);
            RunSelectedFirstMathTypeRangeCollisionAcceptance(
                application,
                pngPath,
                emfPath);
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);
            var service = new WordFormulaService(application);

            var sources = new[]
            {
                (Latex: @"\hbar\omega+1", MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mi>ℏ</mi><mi>ω</mi><mo>+</mo><mn>1</mn></math>"),
                (Latex: @"\int_0^1 x^2\,dx", MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup><msup><mi>x</mi><mn>2</mn></msup><mi>d</mi><mi>x</mi></math>"),
            };

            for (var index = 0; index < sources.Length; index++)
            {
                application.Selection.EndKey(Word.WdUnits.wdStory);
                var insertion = application.Selection.Range;
                var session = CreateSimpleVisualTeXSourceSession(
                    sources[index].Latex,
                    numbered: true);
                session.ObjectMode = FormulaOleContract.MathTypeOleMode;
                session.SourceDocumentId = document.FullName;
                session.SourceObjectId = WordRangeReference(insertion.Start, insertion.End);
                session.MathTypeNumberPosition = "right";
                Release(insertion);
                service.InsertMathTypeOle(
                    session,
                    sources[index].MathMl,
                    emfPath,
                    updateCreatedMathTypeNumberFields: true);
                application.Selection.EndKey(Word.WdUnits.wdStory);
                application.Selection.TypeParagraph();
            }

            AssertEqual(2, CountMathTypeOleShapes(document),
                "Numbered MathType→VisualTeX core fixture did not create two Equation.DSMT4 sources.");
            AssertEqual(2, CountMathTypePlaceRefFields(document),
                "Numbered MathType→VisualTeX core fixture did not create two MTPlaceRef fields.");

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.NativeOleMode);
            AssertEqual(2, plan.Targets.Count,
                "Numbered MathType→VisualTeX core capture did not find two sources.");
            AssertTrue(plan.Targets.All(target => target.Numbered),
                "Numbered MathType→VisualTeX core capture lost numbered state.");

            var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
            for (var index = 0; index < plan.Targets.Count; index++)
            {
                var target = plan.Targets[index];
                var source = sources[index];
                prepared[target.Id] = new PreparedWordBulkFormula
                {
                    Run = new WordBulkRun
                    {
                        Id = target.Id,
                        IsFormula = true,
                        Latex = target.Latex,
                        DisplayMode = target.DisplayMode,
                    },
                    Session = CreateSimpleFormatTargetSession(
                        target,
                        FormulaOleContract.NativeOleMode,
                        source.MathMl),
                    MathMl = source.MathMl,
                    PngPath = pngPath,
                    EmfPath = emfPath,
                };
            }

            var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);
            Console.WriteLine(
                $"[MT→VT NUMBERED CORE] converted={result.FormulaCount} failed={result.FailedFormulaCount} failures={string.Join(" | ", result.Failures)}");
            AssertEqual(2, result.FormulaCount,
                "Numbered MathType→VisualTeX core conversion did not convert both formulas.");
            AssertEqual(0, result.FailedFormulaCount,
                "Numbered MathType→VisualTeX core conversion failed: "
                + string.Join(" | ", result.Failures));
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Numbered MathType→VisualTeX core conversion left MathType sources behind.");
            AssertEqual(2, CountVisualTeXNativeOleShapes(document),
                "Numbered MathType→VisualTeX core conversion did not create two VisualTeX OLE targets.");
            AssertEqual(2, CountInstalledVisualTeXNumberedFormulaHosts(document),
                "Numbered MathType→VisualTeX core conversion lost numbered VisualTeX hosts.");
            AssertEqual(0, CountMathTypePlaceRefFields(document),
                "Numbered MathType→VisualTeX core conversion left MTPlaceRef fields behind.");
            AssertCompressedVisualTeXCaptionParagraphs(document, expectedCount: 2);

            var outputPath = Path.Combine(
                artifactRoot,
                "MathType-To-VisualTeX-Numbered-Core.docx");
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Reopened numbered MathType→VisualTeX core document restored MathType sources.");
            AssertEqual(2, CountVisualTeXNativeOleShapes(document),
                "Reopened numbered MathType→VisualTeX core document lost VisualTeX targets.");
            AssertEqual(2, CountInstalledVisualTeXNumberedFormulaHosts(document),
                "Reopened numbered MathType→VisualTeX core document lost numbered state.");
            AssertCompressedVisualTeXCaptionParagraphs(document, expectedCount: 2);

            var paragraphsWithVisualTeXCaptions = document.Paragraphs.Count;
            var reversePlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.NativeOleMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(2, reversePlan.Targets.Count,
                "VisualTeX→MathType reverse core capture did not find both converted formulas.");
            var reversePrepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
            for (var index = 0; index < reversePlan.Targets.Count; index++)
            {
                var target = reversePlan.Targets[index];
                var source = sources[index];
                reversePrepared[target.Id] = new PreparedWordBulkFormula
                {
                    Run = new WordBulkRun
                    {
                        Id = target.Id,
                        IsFormula = true,
                        Latex = target.Latex,
                        DisplayMode = target.DisplayMode,
                    },
                    Session = CreateSimpleFormatTargetSession(
                        target,
                        FormulaOleContract.MathTypeOleMode,
                        source.MathMl),
                    MathMl = source.MathMl,
                    PngPath = pngPath,
                    EmfPath = emfPath,
                };
            }
            var reverseResult = service.ApplyFormulaFormatConversionPlan(
                reversePlan,
                reversePrepared);
            AssertEqual(2, reverseResult.FormulaCount,
                "VisualTeX→MathType reverse core conversion did not convert both formulas.");
            AssertEqual(0, reverseResult.FailedFormulaCount,
                "VisualTeX→MathType reverse core conversion failed: "
                + string.Join(" | ", reverseResult.Failures));
            AssertEqual(2, CountMathTypeOleShapes(document),
                "VisualTeX→MathType reverse core conversion did not recreate both MathType OLEs.");
            AssertEqual(0, CountVisualTeXNativeOleShapes(document),
                "VisualTeX→MathType reverse core conversion left VisualTeX OLEs behind.");
            AssertEqual(0, CountVisualTeXCaptionBookmarks(document),
                "VisualTeX→MathType reverse core conversion left VTEqCap helper identities behind.");
            AssertEqual(
                paragraphsWithVisualTeXCaptions - 2,
                document.Paragraphs.Count,
                "VisualTeX→MathType removed caption fields but left their paragraph marks as deletable blank lines.");
            AssertEqual(0, document.Frames.Count,
                "VisualTeX→MathType reverse core conversion left a clipped caption Frame behind.");

            Console.WriteLine(
                "Numbered MathType↔VisualTeX core acceptance passed: genuine MTPlaceRef sources converted to numbered VisualTeX OLE hosts with 1pt non-layout caption paragraphs; reverse conversion removed both caption paragraphs completely with no blank-line or Frame residue.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                previousPreviewDisable);
            Environment.SetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE", previousPerf);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            try { if (!string.IsNullOrWhiteSpace(emfPath)) File.Delete(emfPath); } catch { }
            try { File.Delete(pngPath); } catch { }
            try { File.Delete(svgPath); } catch { }
        }
    }

    private static int CountVisualTeXCaptionBookmarks(Word.Document document)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            var count = 0;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmark);
                bookmark = bookmarks[index];
                if (bookmark.Name.StartsWith("VTEqCap_", StringComparison.OrdinalIgnoreCase))
                    count++;
            }
            return count;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void AssertCompressedVisualTeXCaptionParagraphs(
        Word.Document document,
        int expectedCount)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? captionRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Range? paragraphMark = null;
        Word.ParagraphFormat? format = null;
        Word.Font? markFont = null;
        Word.Frames? frames = null;
        Word.Frame? frame = null;
        try
        {
            bookmarks = document.Bookmarks;
            var found = 0;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmark);
                bookmark = bookmarks[index];
                if (!bookmark.Name.StartsWith("VTEqCap_", StringComparison.OrdinalIgnoreCase))
                    continue;
                found++;
                Release(captionRange); captionRange = bookmark.Range;
                Release(paragraphs); paragraphs = captionRange.Paragraphs;
                AssertEqual(1, paragraphs.Count,
                    "VisualTeX caption helper no longer owns exactly one paragraph.");
                Release(paragraph); paragraph = paragraphs[1];
                Release(paragraphRange); paragraphRange = paragraph.Range;
                Release(format); format = paragraph.Format;
                AssertEqual(Word.WdLineSpacing.wdLineSpaceExactly, format.LineSpacingRule,
                    "VisualTeX caption helper still participates in normal body line spacing.");
                AssertNear(1f, format.LineSpacing, 0.1f,
                    "VisualTeX caption helper did not collapse to a 1pt body-flow line.");
                Release(paragraphMark);
                paragraphMark = document.Range(paragraphRange.End - 1, paragraphRange.End);
                Release(markFont); markFont = paragraphMark.Font;
                AssertNear(1f, markFont.Size, 0.1f,
                    "VisualTeX caption paragraph mark still occupies normal text height.");
                Release(frames); frames = captionRange.Frames;
                AssertTrue(frames.Count > 0,
                    "VisualTeX caption helper lost its clipping Frame.");
                Release(frame); frame = frames[1];
                AssertNear(0.1f, frame.Width, 0.05f,
                    "VisualTeX caption helper Frame width changed.");
                AssertNear(0.1f, frame.Height, 0.05f,
                    "VisualTeX caption helper Frame height changed.");
            }
            AssertEqual(expectedCount, found,
                "VisualTeX caption helper count changed unexpectedly.");
        }
        finally
        {
            Release(frame);
            Release(frames);
            Release(markFont);
            Release(format);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(captionRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void RunSelectedFirstMathTypeRangeCollisionAcceptance(
        Word.Application application,
        string pngPath,
        string emfPath)
    {
        Word.Document? document = null;
        Word.InlineShape? firstShape = null;
        Word.InlineShape? remainingMathType = null;
        Word.Range? firstRange = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);
            var service = new WordFormulaService(application);
            var sources = new[]
            {
                (Latex: @"a+b", MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mi>a</mi><mo>+</mo><mi>b</mi></math>"),
                (Latex: @"c+d", MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mi>c</mi><mo>+</mo><mi>d</mi></math>"),
            };
            foreach (var source in sources)
            {
                application.Selection.EndKey(Word.WdUnits.wdStory);
                var insertion = application.Selection.Range;
                try
                {
                    var session = CreateSimpleVisualTeXSourceSession(
                        source.Latex,
                        numbered: true);
                    session.ObjectMode = FormulaOleContract.MathTypeOleMode;
                    session.SourceDocumentId = document.FullName;
                    session.SourceObjectId = WordRangeReference(
                        insertion.Start,
                        insertion.End);
                    session.MathTypeNumberPosition = "right";
                    service.InsertMathTypeOle(
                        session,
                        source.MathMl,
                        emfPath,
                        updateCreatedMathTypeNumberFields: true);
                }
                finally { Release(insertion); }
            }

            AssertEqual(2, CountMathTypeOleShapes(document),
                "Range-collision fixture did not create two MathType sources.");
            firstShape = document.InlineShapes[1];
            firstRange = firstShape.Range.Duplicate;
            var originalFirstStart = firstRange.Start;
            var originalFirstEnd = firstRange.End;
            firstRange.Select();

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: false,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.NativeOleMode);
            AssertEqual(1, plan.Targets.Count,
                "Selecting the first MathType equation did not capture exactly one conversion source.");
            var target = plan.Targets[0];
            var sourceMathMl = target.SourceMathMl
                ?? throw new InvalidDataException(
                    "Selected MathType range-collision source has no MathML.");
            var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal)
            {
                [target.Id] = new PreparedWordBulkFormula
                {
                    Run = new WordBulkRun
                    {
                        Id = target.Id,
                        IsFormula = true,
                        Latex = target.Latex,
                        DisplayMode = target.DisplayMode,
                    },
                    Session = CreateSimpleFormatTargetSession(
                        target,
                        FormulaOleContract.NativeOleMode,
                        sourceMathMl),
                    MathMl = sourceMathMl,
                    PngPath = pngPath,
                    EmfPath = emfPath,
                },
            };

            var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);
            AssertEqual(1, result.FormulaCount,
                "Selected first MathType→VisualTeX conversion was not committed.");
            AssertEqual(0, result.FailedFormulaCount,
                "Selected first MathType→VisualTeX conversion falsely treated the shifted second formula as the deleted source: "
                + string.Join(" | ", result.Failures));
            AssertEqual(1, CountMathTypeOleShapes(document),
                "Selected first MathType conversion removed the wrong number of MathType sources.");
            AssertEqual(1, CountVisualTeXNativeOleShapes(document),
                "Selected first MathType conversion did not create one VisualTeX target.");

            for (var index = 1; index <= document.InlineShapes.Count; index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate = document.InlineShapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    remainingMathType = candidate;
                    candidate = null;
                    break;
                }
                finally { Release(candidate); }
            }
            AssertTrue(remainingMathType is not null,
                "The second MathType formula did not survive the selected-first conversion.");
            var shiftedRange = remainingMathType!.Range;
            try
            {
                AssertTrue(shiftedRange.Start > originalFirstStart,
                    "The surviving second MathType formula did not remain after the converted VisualTeX target was inserted at the first source position.");
                AssertEqual(originalFirstEnd - originalFirstStart,
                    shiftedRange.End - shiftedRange.Start,
                    "The surviving second MathType OLE changed its own one-object Word range length during selected-first conversion.");
            }
            finally { Release(shiftedRange); }
            Console.WriteLine(
                "[MT→VT RANGE COLLISION] Selected first of two numbered MathType rows converted successfully; the second Equation.DSMT4 survived and the stale numeric source range no longer caused a delete-source false positive.");
        }
        finally
        {
            Release(firstRange);
            Release(remainingMathType);
            Release(firstShape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }
}
