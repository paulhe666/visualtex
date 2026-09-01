using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeAdjacentFormatConversionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        const string firstLatex = @"\frac{d}{dx}a=b";
        const string secondLatex = @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}";
        const string firstMathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mi>d</mi><mrow><mi>d</mi><mi>x</mi></mrow></mfrac><mi>a</mi><mo>=</mo><mi>b</mi></math>";
        const string secondMathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>";

        var svgPath = Path.Combine(artifactRoot, "adjacent-format-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"96\" viewBox=\"0 0 240 96\"><text x=\"4\" y=\"64\" font-family=\"Times New Roman\" font-size=\"48\">x+1</text></svg>");
        var emfPath = VisualTeX.WindowsOffice.VstoShared.OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 96);

        ProbeAdjacentOmmlGroupInsertion();
        RunDirection(targetVisualTeX: true);
        RunDirection(targetVisualTeX: false);
        RunNumberedDisplayGroupTrailingSourceRegression();
        RunDisplayToVisualTeXBoundary();
        Console.WriteLine("[ADJACENT FORMAT PASS] zero-gap MathType OLE pair survived MT→VisualTeX and MT→OMML without cross-formula contamination; every unnumbered display conversion direction among MathType, VisualTeX and OMML preserved its paragraph boundary without an extra blank line.");

        void ProbeAdjacentOmmlGroupInsertion()
        {
            Word.Application? application = null;
            Word.Document? document = null;
            Word.Range? target = null;
            WordOmmlConverter.BatchSource? source = null;
            IReadOnlyList<Word.Range>? inserted = null;
            try
            {
                application = CreateWordApplication(visible: false);
                document = application.Documents.Add();
                document.Content.Text = "AB";
                var firstId = Guid.NewGuid().ToString("D");
                var secondId = Guid.NewGuid().ToString("D");
                source = WordOmmlConverter.CreateBatchSource(
                    application,
                    new[]
                    {
                        (FormulaId: firstId, MathMl: firstMathMl),
                        (FormulaId: secondId, MathMl: secondMathMl),
                    });
                target = document.Range(0, 2);
                inserted = source.InsertAdjacentInlineGroup(
                    application,
                    document,
                    target,
                    new[] { firstId, secondId });
                AssertEqual(2, inserted.Count,
                    "Adjacent OMML group paste did not return two equation ranges.");
                AssertEqual(2, document.OMaths.Count,
                    "Adjacent OMML group paste was normalized into one OMath by Word.");
                Console.WriteLine("[ADJACENT GROUP PROBE] one-shot sibling OMath paste retained 2 independent equations.");
            }
            finally
            {
                if (inserted is not null)
                    foreach (var range in inserted) Release(range);
                source?.Dispose();
                Release(target);
                try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                try { QuitWordApplicationIfOwned(application); } catch { }
                Release(document);
                Release(application);
                ForceComCleanup();
            }
        }

        void RunNumberedDisplayGroupTrailingSourceRegression()
        {
            const string thirdLatex = @"x_{n+1}=x_n-\frac{f(x_n)}{f'(x_n)}";
            const string thirdMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msub><mi>x</mi><mrow><mi>n</mi><mo>+</mo><mn>1</mn></mrow></msub><mo>=</mo><msub><mi>x</mi><mi>n</mi></msub><mo>−</mo><mfrac><mrow><mi>f</mi><mfenced><msub><mi>x</mi><mi>n</mi></msub></mfenced></mrow><mrow><msup><mi>f</mi><mo>′</mo></msup><mfenced><msub><mi>x</mi><mi>n</mi></msub></mfenced></mrow></mfrac></math>";
            var tracePath = Path.Combine(
                artifactRoot,
                "numbered-display-group-trailing-source.trace.log");
            var previousTracePath = Environment.GetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH");
            Word.Application? application = null;
            Word.Document? document = null;
            try
            {
                try { File.Delete(tracePath); } catch { }
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_WORD_HOOK_TRACE_PATH",
                    tracePath);
                application = CreateWordApplication(visible: false);
                document = application.Documents.Add();
                // Paragraphs 2 and 3 are deliberately consecutive numbered
                // MTDisplayEquation hosts. Their MathType number fields make each
                // paragraph ~190 Word characters even though the OLE shape itself
                // occupies only ~26. Paragraph 5 is an unrelated trailing MathType
                // source separated by ordinary prose. Replacing paragraphs 2-3 by
                // native OMaths makes paragraph 5 drift far into the frozen ranges
                // of the grouped sources—the exact production failure geometry.
                document.Content.Text =
                    "group-before\r\r\r"
                    + new string('G', 147)
                    + "\r\rgroup-after\r";
                var service = new WordFormulaService(application);

                void InsertAtParagraph(
                    int paragraphIndex,
                    string latex,
                    string mathMl,
                    bool numbered)
                {
                    Word.Paragraph? paragraph = null;
                    Word.Range? insertion = null;
                    try
                    {
                        paragraph = document.Paragraphs[paragraphIndex];
                        insertion = paragraph.Range.Duplicate;
                        insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        insertion.Select();
                        service.InsertMathTypeOle(
                            CreateMathTypeCreateSession(
                                "block",
                                numbered,
                                latex),
                            mathMl,
                            emfPath,
                            preserveExistingDisplayParagraphBoundary: true);
                    }
                    finally
                    {
                        Release(insertion);
                        Release(paragraph);
                    }
                }

                // End-to-start insertion keeps the blank paragraph indexes stable.
                InsertAtParagraph(5, thirdLatex, thirdMathMl, numbered: false);
                InsertAtParagraph(3, secondLatex, secondMathMl, numbered: true);
                InsertAtParagraph(2, firstLatex, firstMathMl, numbered: true);
                AssertEqual(3, CountMathTypeOleShapes(document),
                    "Numbered display-group drift setup did not create three MathType sources.");

                var plan = service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.MathTypeOleMode,
                    FormulaOleContract.WordOmmlMode);
                var ordered = plan.Targets.OrderBy(target => target.SourceStart).ToArray();
                AssertEqual(3, ordered.Length,
                    "Numbered display-group drift setup did not capture three MathType sources.");
                AssertTrue(ordered[0].Numbered && ordered[1].Numbered && !ordered[2].Numbered,
                    "Numbered display-group drift setup did not preserve the 2-numbered + 1-unnumbered source topology.");

                static (int Start, int End) ParseRangeReference(string reference)
                {
                    var parts = (reference ?? string.Empty).Split(':');
                    if (parts.Length < 3
                        || !int.TryParse(parts[parts.Length - 2], out var start)
                        || !int.TryParse(parts[parts.Length - 1], out var end))
                        throw new InvalidDataException(
                            "Invalid frozen Word range reference: " + reference);
                    return (start, end);
                }

                var frozenFirst = ParseRangeReference(ordered[0].SourceObjectId);
                var frozenSecond = ParseRangeReference(ordered[1].SourceObjectId);
                var prepared = PrepareOmmlMathTypeTargets(plan, emfPath);
                var expectedSignatures = ordered.ToDictionary(
                    target => prepared[target.Id].Session.FormulaId,
                    target => MathTypeMtefCodec.SemanticSignature(
                        target.SourceMathMl
                        ?? throw new InvalidDataException(
                            "A numbered display-group source lost its MathML.")),
                    StringComparer.OrdinalIgnoreCase);

                var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);
                AssertEqual(3, result.FormulaCount,
                    "Numbered display-group + trailing source did not convert all three MathType formulas. Failures: "
                    + string.Join(" | ", result.Failures));
                AssertEqual(0, result.FailedFormulaCount,
                    "Numbered display-group + trailing source reported a conversion failure: "
                    + string.Join(" | ", result.Failures));
                AssertEqual(0, CountMathTypeOleShapes(document),
                    "Numbered display-group + trailing source left a MathType source behind.");
                AssertEqual(3, document.OMaths.Count,
                    "Numbered display-group + trailing source did not create three OMath targets.");

                var verified = 0;
                var numberedHosts = 0;
                foreach (var entry in expectedSignatures)
                {
                    Word.Bookmark? bookmark = null;
                    Word.Range? range = null;
                    try
                    {
                        bookmark = WordOmmlFormulaStore.FindByFormulaId(document, entry.Key)
                            ?? throw new InvalidDataException(
                                $"Numbered display-group target {entry.Key} lost its VTOMML bookmark.");
                        range = WordOmmlFormulaStore.GetEquationRange(bookmark);
                        var actualMathMl = WordOmmlConverter.TransformOmmlToMathMl(
                            range.WordOpenXML,
                            display: true);
                        AssertEqual(
                            entry.Value,
                            MathTypeMtefCodec.SemanticSignature(actualMathMl),
                            $"Numbered display-group target {entry.Key} changed semantics.");
                        var metadata = WordOmmlFormulaStore.TryRead(document, entry.Key);
                        if (metadata?.Numbered == true)
                        {
                            AssertTrue(
                                WordEquationNumbering.HasReusableNumberedNativeOmmlDirectTableHost(
                                    document,
                                    range,
                                    entry.Key),
                                $"Numbered display-group target {entry.Key} is not the required 1x3 direct-SEQ host.");
                            numberedHosts++;
                        }
                        verified++;
                    }
                    finally
                    {
                        Release(range);
                        Release(bookmark);
                    }
                }
                AssertEqual(3, verified,
                    "Numbered display-group regression did not verify all three OMML targets.");
                AssertEqual(2, numberedHosts,
                    "Numbered display-group regression did not retain exactly two 1x3 direct-SEQ hosts.");

                var trace = File.Exists(tracePath)
                    ? File.ReadAllLines(tracePath)
                    : Array.Empty<string>();
                AssertTrue(
                    trace.Any(line => line.IndexOf(
                        "format-conversion-block-omml-groups groups=1 formulas=2",
                        StringComparison.Ordinal) >= 0),
                    "Numbered display-group regression did not exercise the 2-formula atomic block path.");
                var trailingLiveLine = trace.FirstOrDefault(line =>
                    line.IndexOf(
                        "format-conversion-forward-source-live formulaId=" + ordered[2].SourceFormulaId,
                        StringComparison.Ordinal) >= 0);
                AssertTrue(!string.IsNullOrWhiteSpace(trailingLiveLine),
                    "Numbered display-group regression did not capture the trailing source's live range after group replacement.");
                var rangeMarker = trailingLiveLine!.IndexOf(" range=", StringComparison.Ordinal);
                AssertTrue(rangeMarker >= 0,
                    "Trailing-source trace has no live range: " + trailingLiveLine);
                var liveRangeText = trailingLiveLine.Substring(rangeMarker + 7).Split(' ')[0];
                var liveParts = liveRangeText.Split(':');
                var liveStart = 0;
                var liveEnd = 0;
                AssertTrue(liveParts.Length == 2
                    && int.TryParse(liveParts[0], out liveStart)
                    && int.TryParse(liveParts[1], out liveEnd),
                    "Trailing-source live range could not be parsed: " + trailingLiveLine);
                static bool Overlaps(
                    (int Start, int End) left,
                    (int Start, int End) right) =>
                    left.Start < right.End && left.End > right.Start;
                var liveTrailing = (Start: liveStart, End: liveEnd);
                AssertTrue(
                    Overlaps(liveTrailing, frozenFirst)
                    || Overlaps(liveTrailing, frozenSecond),
                    "The regression fixture did not reproduce the stale-range collision: "
                    + $"frozen1={frozenFirst.Start}:{frozenFirst.End}; "
                    + $"frozen2={frozenSecond.Start}:{frozenSecond.End}; "
                    + $"trailingLive={liveStart}:{liveEnd}.");

                Console.WriteLine(
                    "[NUMBERED DISPLAY GROUP RANGE-DRIFT PASS] Two consecutive numbered MathType paragraphs were replaced atomically; "
                    + $"the trailing MathType drifted to {liveStart}:{liveEnd} and overlapped a frozen group-member range, "
                    + "yet all 3 targets converted with exact semantics and both numbered targets remained 1x3 direct-SEQ.");
            }
            finally
            {
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_WORD_HOOK_TRACE_PATH",
                    previousTracePath);
                try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                try { QuitWordApplicationIfOwned(application); } catch { }
                Release(document);
                Release(application);
                ForceComCleanup();
            }
        }

        void RunDisplayToVisualTeXBoundary()
        {
            var previousFormatAcceptance = Environment.GetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
            var previousTracePath = Environment.GetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH");
            var tracePath = Path.Combine(
                artifactRoot,
                "display-mt-to-vt-boundary.trace.log");
            try { File.Delete(tracePath); } catch { }
            Word.Application? application = null;
            Word.Document? document = null;
            Word.Document? reopened = null;
            Word.Range? insertion = null;
            Word.Paragraphs? paragraphs = null;
            Word.Paragraph? formulaParagraph = null;
            Word.Range? formulaRange = null;
            Word.InlineShapes? formulaShapes = null;
            Word.InlineShape? shape = null;
            VisualTeX.WordVsto.ThisAddIn? addIn = null;
            Array custom = Array.Empty<object>();
            try
            {
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                    "1");
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_WORD_HOOK_TRACE_PATH",
                    tracePath);
                var mathTypeBaseline = SnapshotMathTypeProcessIds();
                application = CreateWordApplication(visible: false);
                document = application.Documents.Add();
                document.Content.Text = "display-before\r\rdisplay-after\r";
                paragraphs = document.Paragraphs;
                AssertTrue(paragraphs.Count >= 3,
                    "Display MT→VisualTeX boundary setup did not create the source/formula/following paragraph sequence.");
                var sourceParagraphCount = paragraphs.Count;
                formulaParagraph = paragraphs[2];
                formulaRange = formulaParagraph.Range.Duplicate;
                insertion = document.Range(formulaRange.Start, formulaRange.Start);
                insertion.Select();
                var service = new WordFormulaService(application);
                service.InsertMathTypeOle(
                    CreateMathTypeCreateSession("block", false, firstLatex),
                    firstMathMl,
                    emfPath,
                    preserveExistingDisplayParagraphBoundary: true);
                var paragraphCountBefore = document.Paragraphs.Count;
                AssertEqual(sourceParagraphCount, paragraphCountBefore,
                    "Display MathType source setup changed the baseline paragraph count.");

                addIn = new VisualTeX.WordVsto.ThisAddIn();
                addIn.OnConnection(
                    application,
                    Extensibility.ext_ConnectMode.ext_cm_AfterStartup,
                    addIn,
                    ref custom);
                addIn.OnConvertMathTypeToVisualTeXDocument(new object());
                WaitForInstalledOmmlMathTypeConversion(
                    tracePath,
                    "source=MathType target=VisualTeX",
                    mathTypeBaseline);
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));

                AssertEqual(1, CountInstalledVisualTeXOleShapes(document),
                    "Display MT→VisualTeX did not create exactly one VisualTeX formula.");
                AssertEqual(0, CountMathTypeOleShapes(document),
                    "Display MT→VisualTeX left its MathType source behind.");
                AssertEqual(paragraphCountBefore, document.Paragraphs.Count,
                    "Display MT→VisualTeX changed the paragraph count and introduced an extra blank line.");
                AssertVisualTeXDisplayFollowedImmediatelyByText(
                    document,
                    "display-after",
                    "Display MT→VisualTeX");

                addIn.OnConvertVisualTeXToOmmlDocument(new object());
                WaitForInstalledOmmlMathTypeConversion(
                    tracePath,
                    "source=VisualTeX target=OMML",
                    mathTypeBaseline);
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));
                AssertEqual(1, document.OMaths.Count,
                    "Display VisualTeX→OMML did not create exactly one Word equation.");
                AssertEqual(0, CountInstalledVisualTeXOleShapes(document),
                    "Display VisualTeX→OMML left its VisualTeX source behind.");
                AssertEqual(paragraphCountBefore, document.Paragraphs.Count,
                    "Display VisualTeX→OMML changed the paragraph count.");
                AssertOmmlDisplayFollowedImmediatelyByText(
                    document,
                    "display-after",
                    "Display VisualTeX→OMML");

                addIn.OnConvertOmmlToVisualTeXDocument(new object());
                WaitForInstalledOmmlMathTypeConversion(
                    tracePath,
                    "source=OMML target=VisualTeX",
                    mathTypeBaseline);
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));
                AssertEqual(1, CountInstalledVisualTeXOleShapes(document),
                    "Display OMML→VisualTeX did not create exactly one VisualTeX formula.");
                AssertEqual(0, document.OMaths.Count,
                    "Display OMML→VisualTeX left its OMML source behind.");
                AssertEqual(paragraphCountBefore, document.Paragraphs.Count,
                    "Display OMML→VisualTeX changed the paragraph count.");
                AssertVisualTeXDisplayFollowedImmediatelyByText(
                    document,
                    "display-after",
                    "Display OMML→VisualTeX");

                addIn.OnConvertVisualTeXToMathTypeDocument(new object());
                WaitForInstalledOmmlMathTypeConversion(
                    tracePath,
                    "source=VisualTeX target=MathType",
                    mathTypeBaseline);
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));
                AssertEqual(1, CountMathTypeOleShapes(document),
                    "Display VisualTeX→MathType did not create exactly one MathType formula.");
                AssertEqual(0, CountInstalledVisualTeXOleShapes(document),
                    "Display VisualTeX→MathType left its VisualTeX source behind.");
                AssertEqual(paragraphCountBefore, document.Paragraphs.Count,
                    "Display VisualTeX→MathType changed the paragraph count.");
                AssertMathTypeDisplayFollowedImmediatelyByText(
                    document,
                    "display-after",
                    "Display VisualTeX→MathType");

                addIn.OnConvertMathTypeToOmmlDocument(new object());
                WaitForInstalledOmmlMathTypeConversion(
                    tracePath,
                    "source=MathType target=OMML",
                    mathTypeBaseline);
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));
                AssertEqual(1, document.OMaths.Count,
                    "Display MathType→OMML did not create exactly one Word equation.");
                AssertEqual(0, CountMathTypeOleShapes(document),
                    "Display MathType→OMML left its MathType source behind.");
                AssertEqual(paragraphCountBefore, document.Paragraphs.Count,
                    "Display MathType→OMML changed the paragraph count.");
                AssertOmmlDisplayFollowedImmediatelyByText(
                    document,
                    "display-after",
                    "Display MathType→OMML");

                addIn.OnConvertOmmlToMathTypeDocument(new object());
                WaitForInstalledOmmlMathTypeConversion(
                    tracePath,
                    "source=OMML target=MathType",
                    mathTypeBaseline);
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));
                AssertEqual(1, CountMathTypeOleShapes(document),
                    "Display OMML→MathType did not create exactly one MathType formula.");
                AssertEqual(0, document.OMaths.Count,
                    "Display OMML→MathType left its OMML source behind.");
                AssertEqual(paragraphCountBefore, document.Paragraphs.Count,
                    "Display OMML→MathType changed the paragraph count.");
                AssertMathTypeDisplayFollowedImmediatelyByText(
                    document,
                    "display-after",
                    "Display OMML→MathType");

                var outputPath = Path.Combine(
                    artifactRoot,
                    "Display-Format-Conversion-Boundary-Matrix.docx");
                document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
                document.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = null;
                reopened = application.Documents.Open(
                    outputPath,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);
                AssertEqual(paragraphCountBefore, reopened.Paragraphs.Count,
                    "Save/reopen changed the display conversion paragraph count.");
                AssertMathTypeDisplayFollowedImmediatelyByText(
                    reopened,
                    "display-after",
                    "Reopened display conversion matrix");
                AssertNoNewMathTypeProcess(
                    mathTypeBaseline,
                    "display format conversion boundary matrix");
                Console.WriteLine(
                    "[DISPLAY FORMAT] MT→VisualTeX→OMML→VisualTeX→MathType→OMML→MathType preserved the same display paragraph and following user paragraph with no generated blank line.");
            }
            finally
            {
                Release(shape);
                Release(formulaShapes);
                Release(formulaRange);
                Release(formulaParagraph);
                Release(paragraphs);
                Release(insertion);
                if (addIn is not null)
                {
                    try
                    {
                        addIn.OnDisconnection(
                            Extensibility.ext_DisconnectMode.ext_dm_UserClosed,
                            ref custom);
                    }
                    catch { }
                }
                try { reopened?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                try { QuitWordApplicationIfOwned(application); } catch { }
                Release(reopened);
                Release(document);
                Release(application);
                ForceComCleanup();
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                    previousFormatAcceptance);
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_WORD_HOOK_TRACE_PATH",
                    previousTracePath);
            }
        }

        void RunDirection(bool targetVisualTeX)
        {
            var previousFormatAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
            var previousTracePath = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
            var tracePath = Path.Combine(
                artifactRoot,
                targetVisualTeX ? "adjacent-mt-to-vt.trace.log" : "adjacent-mt-to-omml.trace.log");
            try { File.Delete(tracePath); } catch { }
            Word.Application? application = null;
            Word.Document? document = null;
            Word.Range? insertion = null;
            Word.InlineShape? first = null;
            Word.InlineShape? second = null;
            VisualTeX.WordVsto.ThisAddIn? addIn = null;
            Array custom = Array.Empty<object>();
            try
            {
                Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
                Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
                var mathTypeBaseline = SnapshotMathTypeProcessIds();
                application = CreateWordApplication(visible: false);
                document = application.Documents.Add();
                document.Activate();
                var service = new WordFormulaService(application);

                insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
                insertion.Select();
                service.InsertMathTypeOle(
                    CreateMathTypeCreateSession("inline", false, firstLatex),
                    firstMathMl,
                    emfPath,
                    createdObjectBookmarkName: "VTMT_ADJ_FMT_FIRST");
                Release(insertion); insertion = null;

                first = document.InlineShapes[1];
                var boundary = first.Range.End;
                insertion = document.Range(boundary, boundary);
                insertion.Select();
                service.InsertMathTypeOle(
                    CreateMathTypeCreateSession("inline", false, secondLatex),
                    secondMathMl,
                    emfPath,
                    createdObjectBookmarkName: "VTMT_ADJ_FMT_SECOND");
                Release(insertion); insertion = null;
                Release(first); first = null;

                AssertEqual(2, document.InlineShapes.Count,
                    "Zero-gap MathType setup did not retain exactly two OLE equations.");
                first = document.InlineShapes[1];
                second = document.InlineShapes[2];
                AssertEqual(first.Range.End, second.Range.Start,
                    "Zero-gap MathType setup unexpectedly inserted a separator between the two OLE equations.");
                AssertEqual(
                    MathTypeMtefCodec.SemanticSignature(firstMathMl),
                    MathTypeMtefCodec.SemanticSignature(MathTypeOleStorage.ReadMathMl(first)),
                    "Zero-gap first MathType equation changed before batch conversion.");
                AssertEqual(
                    MathTypeMtefCodec.SemanticSignature(secondMathMl),
                    MathTypeMtefCodec.SemanticSignature(MathTypeOleStorage.ReadMathMl(second)),
                    "Zero-gap second MathType equation changed before batch conversion.");
                Release(first); first = null;
                Release(second); second = null;

                addIn = new VisualTeX.WordVsto.ThisAddIn();
                addIn.OnConnection(
                    application,
                    Extensibility.ext_ConnectMode.ext_cm_AfterStartup,
                    addIn,
                    ref custom);
                if (targetVisualTeX)
                    addIn.OnConvertMathTypeToVisualTeXDocument(new object());
                else
                    addIn.OnConvertMathTypeToOmmlDocument(new object());
                WaitForInstalledOmmlMathTypeConversion(
                    tracePath,
                    targetVisualTeX
                        ? "source=MathType target=VisualTeX"
                        : "source=MathType target=OMML",
                    mathTypeBaseline);
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));

                if (targetVisualTeX)
                {
                    AssertEqual(2, CountInstalledVisualTeXOleShapes(document),
                        "Zero-gap MT→VisualTeX did not create exactly two VisualTeX formulas.");
                    var expected = new[]
                    {
                        NormalizeStressLatex(MathMlToLatexConverter.Convert(firstMathMl)),
                        NormalizeStressLatex(MathMlToLatexConverter.Convert(secondMathMl)),
                    };
                    var seen = 0;
                    for (var index = 1; index <= document.InlineShapes.Count; index++)
                    {
                        Word.InlineShape? shape = null;
                        try
                        {
                            shape = document.InlineShapes[index];
                            if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                            var metadata = WordFormulaMetadataReader.TryRead(shape)
                                ?? throw new InvalidDataException($"Zero-gap VisualTeX target #{seen + 1} has no metadata.");
                            AssertEqual(expected[seen], NormalizeStressLatex(metadata.Latex ?? string.Empty),
                                $"Zero-gap MT→VisualTeX formula #{seen + 1} inherited content from its neighbor.");
                            seen++;
                        }
                        finally { Release(shape); }
                    }
                    AssertEqual(2, seen, "Zero-gap MT→VisualTeX did not inspect two targets.");
                }
                else
                {
                    AssertEqual(2, document.OMaths.Count,
                        "Zero-gap MT→OMML did not create exactly two OMath equations.");
                    var expectedSignatures = new[]
                    {
                        MathTypeMtefCodec.SemanticSignature(firstMathMl),
                        MathTypeMtefCodec.SemanticSignature(secondMathMl),
                    };
                    for (var index = 1; index <= 2; index++)
                    {
                        Word.OMath? math = null;
                        Word.Range? range = null;
                        try
                        {
                            math = document.OMaths[index];
                            range = math.Range;
                            var roundTrip = WordOmmlConverter.TransformOmmlToMathMl(range.WordOpenXML, display: false);
                            AssertEqual(expectedSignatures[index - 1], MathTypeMtefCodec.SemanticSignature(roundTrip),
                                $"Zero-gap MT→OMML formula #{index} inherited content from its neighbor.");
                        }
                        finally
                        {
                            Release(range);
                            Release(math);
                        }
                    }
                }
                AssertNoNewMathTypeProcess(
                    mathTypeBaseline,
                    targetVisualTeX ? "zero-gap MT→VisualTeX" : "zero-gap MT→OMML");
                Console.WriteLine($"[ADJACENT FORMAT] {(targetVisualTeX ? "MT→VisualTeX" : "MT→OMML")} passed.");
            }
            finally
            {
                Release(first);
                Release(second);
                Release(insertion);
                if (addIn is not null)
                {
                    try
                    {
                        addIn.OnDisconnection(
                            Extensibility.ext_DisconnectMode.ext_dm_UserClosed,
                            ref custom);
                    }
                    catch { }
                }
                try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                try { QuitWordApplicationIfOwned(application); } catch { }
                Release(document);
                Release(application);
                ForceComCleanup();
                Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", previousFormatAcceptance);
                Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTracePath);
            }
        }
    }

    private static void AssertOmmlDisplayFollowedImmediatelyByText(
        Word.Document document,
        string expectedFollowingText,
        string context)
    {
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Paragraph? formulaParagraph = null;
        Word.Range? formulaRange = null;
        Word.OMaths? maths = null;
        try
        {
            paragraphs = document.Paragraphs;
            var followingIndex = -1;
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Release(paragraphRange);
                paragraphRange = null;
                Release(paragraph);
                paragraph = paragraphs[index];
                paragraphRange = paragraph.Range;
                var text = (paragraphRange.Text ?? string.Empty)
                    .Trim('\r', '\a', '\v', '\f', '\t', ' ');
                if (!string.Equals(text, expectedFollowingText, StringComparison.Ordinal))
                    continue;
                followingIndex = index;
                break;
            }
            AssertTrue(followingIndex > 1,
                $"{context} could not resolve the user paragraph following the formula.");

            formulaParagraph = paragraphs[followingIndex - 1];
            formulaRange = formulaParagraph.Range;
            maths = formulaRange.OMaths;
            AssertEqual(1, maths.Count,
                $"{context} left a plain blank paragraph between the OMML display formula and '{expectedFollowingText}'.");
        }
        finally
        {
            Release(maths);
            Release(formulaRange);
            Release(formulaParagraph);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void AssertVisualTeXDisplayFollowedImmediatelyByText(
        Word.Document document,
        string expectedFollowingText,
        string context)
    {
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Paragraph? formulaParagraph = null;
        Word.Range? formulaRange = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        try
        {
            paragraphs = document.Paragraphs;
            var followingIndex = -1;
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Release(paragraphRange);
                paragraphRange = null;
                Release(paragraph);
                paragraph = paragraphs[index];
                paragraphRange = paragraph.Range;
                var text = (paragraphRange.Text ?? string.Empty)
                    .Trim('\r', '\a', '\v', '\f', '\t', ' ');
                if (!string.Equals(text, expectedFollowingText, StringComparison.Ordinal))
                    continue;
                followingIndex = index;
                break;
            }
            AssertTrue(followingIndex > 1,
                $"{context} could not resolve the user paragraph following the formula.");

            formulaParagraph = paragraphs[followingIndex - 1];
            formulaRange = formulaParagraph.Range;
            shapes = formulaRange.InlineShapes;
            var visualTeXCount = 0;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (WordFormulaMetadataReader.IsNativeOle(shape)) visualTeXCount++;
            }
            AssertEqual(1, visualTeXCount,
                $"{context} left a plain blank paragraph between the VisualTeX display formula and '{expectedFollowingText}'.");
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(formulaRange);
            Release(formulaParagraph);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }
}
