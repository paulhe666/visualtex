using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordHostCoreMinimalAcceptance(
        string artifactRoot)
    {
        AssertTrue(
            !AttachActiveWord,
            "Host-core minimal acceptance refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);

        RunWordHostCorePureOmmlAcceptance(
            artifactRoot);
        return;

        // Historical pre-pure-OMML acceptance remains below for source-level
        // comparison only. It is intentionally unreachable because its durable
        // VTOMML/VTEqNum ownership assumptions are retired.
        const string initialMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></math>";
        const string editedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>y</mi><mo>+</mo><mn>2</mn></mrow></math>";

        var assetRoot = Path.Combine(
            Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            "host-core-minimal-"
            + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(assetRoot);
        var pngPath =
            Path.Combine(
                assetRoot,
                "preview.png");
        var svgPath =
            Path.Combine(
                assetRoot,
                "preview.svg");
        WriteAcceptancePng(
            pngPath,
            "y+2",
            260,
            96);
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"96\" viewBox=\"0 0 260 96\"><rect width=\"260\" height=\"96\" fill=\"white\"/><text x=\"10\" y=\"66\" font-family=\"Cambria Math\" font-size=\"44\">y+2</text></svg>");
        var emfPath =
            OfficeOlePreview.CreateVectorEmfFromSvg(
                svgPath,
                260,
                96);

        using var host =
            new WordPerformanceHost(
                documentPath: null);
        var document = host.Document;
        var service =
            new WordFormulaService(
                host.Application);

        document.Content.Text = "left right";
        Word.Range? insertion = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Range? prefix = null;
        Word.Range? suffix = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? identity = null;
        Word.Range? identityRange = null;
        var identityName = string.Empty;

        try
        {
            insertion =
                document.Range(
                    4,
                    4);
            insertion.Select();

            var createSession =
                CreateOmmlMathTypeAcceptanceSession(
                    initialMathMl,
                    "inline",
                    numbered: false,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    createSession,
                    initialMathMl);

            maths = document.OMaths;
            AssertEqual(
                1,
                maths.Count,
                "Inline OMML insert did not create exactly one Word OMath.");
            math = maths[1];
            AssertEqual(
                Word.WdOMathType.wdOMathInline,
                math.Type,
                "Inserted OMML is not a native inline OMath.");

            mathRange =
                math.Range.Duplicate;
            AssertEqual(
                4,
                mathRange.Start,
                "Inline OMML insert moved away from the captured caret.");

            prefix =
                document.Range(
                    0,
                    mathRange.Start);
            AssertEqual(
                "left",
                prefix.Text ?? string.Empty,
                "Text before the inline OMath changed.");

            paragraphs =
                mathRange.Paragraphs;
            AssertEqual(
                1,
                paragraphs.Count,
                "Inline OMath no longer belongs to one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            var bodyEnd =
                Math.Max(
                    paragraphRange.Start,
                    paragraphRange.End - 1);
            suffix =
                document.Range(
                    mathRange.End,
                    bodyEnd);
            AssertEqual(
                " right",
                suffix.Text ?? string.Empty,
                "Text after the inline OMath changed.");

            AssertNoHostCoreSentinelArtifacts(
                document,
                "after inline OMML insert");

            AssertNoVisualTeXOmmlArtifacts(
                document,
                "after inline OMML insert");

            mathRange.Select();
            var selected =
                service.ReadSelection();
            AssertEqual(
                FormulaOleContract.WordOmmlMode,
                selected.ObjectMode ?? string.Empty,
                "ReadSelection did not resolve the native OMath as OMML.");
            AssertTrue(
                Guid.TryParse(
                    selected.FormulaId,
                    out _),
                "ReadSelection did not assign a session-local OMML FormulaId.");
            AssertTrue(
                !string.IsNullOrWhiteSpace(
                    selected.ObjectId),
                "ReadSelection did not capture an exact local source range.");

            var editSession =
                CreateOmmlMathTypeAcceptanceSession(
                    editedMathMl,
                    "inline",
                    numbered: false,
                    FormulaOleContract.WordOmmlMode);
            editSession.Mode = "edit";
            editSession.FormulaId =
                selected.FormulaId!;
            editSession.SourceDocumentId =
                selected.DocumentId;
            editSession.SourceObjectId =
                selected.ObjectId;
            editSession.OriginalMetadata =
                selected.Metadata;

            _ = service.ReplaceOmml(
                editSession,
                editedMathMl);

            Release(identityRange);
            identityRange = null;
            Release(identity);
            identity = null;
            Release(bookmarks);
            bookmarks = null;
            Release(suffix);
            suffix = null;
            Release(paragraphRange);
            paragraphRange = null;
            Release(paragraph);
            paragraph = null;
            Release(paragraphs);
            paragraphs = null;
            Release(mathRange);
            mathRange = null;
            Release(math);
            math = null;
            Release(maths);
            maths = null;

            maths = document.OMaths;
            AssertEqual(
                1,
                maths.Count,
                "Inline OMML edit created or deleted an unrelated OMath.");
            math = maths[1];
            AssertEqual(
                Word.WdOMathType.wdOMathInline,
                math.Type,
                "Edited OMML lost inline type.");
            mathRange =
                math.Range.Duplicate;

            var resolved =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    mathRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Edited OMML could not be resolved from its actual local OMath range.");
            resolved.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    resolved);
            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    resolved);
            AssertTrue(
                payload.Latex.IndexOf(
                    "y",
                    StringComparison.OrdinalIgnoreCase) >= 0
                && payload.Latex.IndexOf(
                    "2",
                    StringComparison.OrdinalIgnoreCase) >= 0,
                "Edited OMML semantic payload does not contain the new formula.");

            AssertTrue(
                string.IsNullOrWhiteSpace(
                    resolved.FormulaId),
                "Resolved native OMML unexpectedly exposes a durable VisualTeX FormulaId.");

            AssertNoHostCoreSentinelArtifacts(
                document,
                "after inline OMML edit");

            prefix =
                document.Range(
                    0,
                    mathRange.Start);
            AssertEqual(
                "left",
                prefix.Text ?? string.Empty,
                "Text before the edited inline OMath changed.");

            paragraphs =
                mathRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            bodyEnd =
                Math.Max(
                    paragraphRange.Start,
                    paragraphRange.End - 1);
            suffix =
                document.Range(
                    mathRange.End,
                    bodyEnd);
            AssertEqual(
                " right",
                suffix.Text ?? string.Empty,
                "Text after the edited inline OMath changed.");

            AssertNoVisualTeXOmmlArtifacts(
                document,
                "after inline OMML edit");

            Release(identityRange);
            identityRange = null;
            Release(identity);
            identity = null;
            Release(bookmarks);
            bookmarks = null;
            Release(suffix);
            suffix = null;
            Release(paragraphRange);
            paragraphRange = null;
            Release(paragraph);
            paragraph = null;
            Release(paragraphs);
            paragraphs = null;

            mathRange.Select();
            var toVisualTeXPlan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: false,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.NativeOleMode);
            AssertEqual(
                1,
                toVisualTeXPlan.Targets.Count,
                "OMML→VisualTeX did not capture exactly the selected local OMath.");

            var toVisualTeXTarget =
                toVisualTeXPlan.Targets.Single();
            var toVisualTeXMathMl =
                toVisualTeXTarget.SourceMathMl
                ?? editedMathMl;
            var toVisualTeXPrepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal)
                {
                    [toVisualTeXTarget.Id] =
                        new PreparedWordBulkFormula
                        {
                            Run = new WordBulkRun
                            {
                                Id = toVisualTeXTarget.Id,
                                IsFormula = true,
                                Latex = toVisualTeXTarget.Latex,
                                DisplayMode =
                                    toVisualTeXTarget.DisplayMode,
                            },
                            Session =
                                CreateSimpleFormatTargetSession(
                                    toVisualTeXTarget,
                                    FormulaOleContract.NativeOleMode,
                                    toVisualTeXMathMl),
                            MathMl =
                                toVisualTeXMathMl,
                            PngPath = pngPath,
                            EmfPath = emfPath,
                        },
                };

            Release(mathRange);
            mathRange = null;
            Release(math);
            math = null;
            Release(maths);
            maths = null;

            var toVisualTeXResult =
                service.ApplyFormulaFormatConversionPlan(
                    toVisualTeXPlan,
                    toVisualTeXPrepared);
            AssertEqual(
                1,
                toVisualTeXResult.FormulaCount,
                "OMML→VisualTeX did not convert exactly one formula.");
            AssertEqual(
                0,
                toVisualTeXResult.FailedFormulaCount,
                "OMML→VisualTeX reported a conversion failure.");
            AssertEqual(
                0,
                document.OMaths.Count,
                "OMML→VisualTeX left a native OMath behind.");
            AssertEqual(
                1,
                CountVisualTeXNativeOleShapes(document),
                "OMML→VisualTeX did not create exactly one VisualTeX OLE.");

            Word.InlineShapes? visualTeXShapes = null;
            Word.InlineShape? visualTeXShape = null;
            Word.Range? visualTeXRange = null;
            try
            {
                visualTeXShapes =
                    document.InlineShapes;
                for (var index = 1;
                     index <= visualTeXShapes.Count;
                     index++)
                {
                    Word.InlineShape? candidate = null;
                    try
                    {
                        candidate =
                            visualTeXShapes[index];
                        if (!WordFormulaMetadataReader
                                .IsNativeOle(candidate))
                            continue;
                        if (visualTeXShape is not null)
                            throw new InvalidDataException(
                                "OMML→VisualTeX created more than one VisualTeX host.");
                        visualTeXShape = candidate;
                        candidate = null;
                    }
                    finally
                    {
                        Release(candidate);
                    }
                }

                AssertTrue(
                    visualTeXShape is not null,
                    "OMML→VisualTeX produced no native VisualTeX OLE.");
                visualTeXRange =
                    visualTeXShape!.Range.Duplicate;
                var embedded =
                    WordFormulaMetadataReader
                        .TryReadEmbeddedNativeOle(
                            visualTeXShape)
                    ?? throw new InvalidDataException(
                        "Converted VisualTeX OLE has no embedded metadata.");
                AssertEqual(
                    created.FormulaId,
                    embedded.FormulaId,
                    "OMML→VisualTeX changed FormulaId.");
                AssertEqual(
                    "inline",
                    embedded.DisplayMode,
                    "OMML→VisualTeX changed display mode.");
                AssertTrue(
                    !embedded.Numbered,
                    "OMML→VisualTeX unexpectedly numbered the inline formula.");

                var visualHost =
                    WordFormulaHostResolver.ResolveLocal(
                        document,
                        visualTeXRange,
                        WordFormulaHostKind.VisualTeX)
                    ?? throw new InvalidDataException(
                        "Converted VisualTeX host cannot be resolved from its actual InlineShape range.");
                AssertEqual(
                    created.FormulaId,
                    visualHost.FormulaId ?? string.Empty,
                    "Resolved VisualTeX host lost FormulaId.");

                prefix =
                    document.Range(
                        0,
                        visualTeXRange.Start);
                AssertEqual(
                    "left",
                    prefix.Text ?? string.Empty,
                    "OMML→VisualTeX changed text before the formula.");
                Release(prefix);
                prefix = null;

                paragraphs =
                    visualTeXRange.Paragraphs;
                paragraph =
                    paragraphs[1];
                paragraphRange =
                    paragraph.Range.Duplicate;
                bodyEnd =
                    Math.Max(
                        paragraphRange.Start,
                        paragraphRange.End - 1);
                suffix =
                    document.Range(
                        visualTeXRange.End,
                        bodyEnd);
                AssertEqual(
                    " right",
                    suffix.Text ?? string.Empty,
                    "OMML→VisualTeX changed text after the formula.");

                AssertNoHostCoreSentinelArtifacts(
                    document,
                    "after OMML→VisualTeX");

                visualTeXRange.Select();
                var toOmmlPlan =
                    service.CaptureFormulaFormatConversionPlan(
                        wholeDocument: false,
                        FormulaOleContract.NativeOleMode,
                        FormulaOleContract.WordOmmlMode);
                AssertEqual(
                    1,
                    toOmmlPlan.Targets.Count,
                    "VisualTeX→OMML did not capture exactly the selected local OLE.");

                var toOmmlTarget =
                    toOmmlPlan.Targets.Single();
                var toOmmlPrepared =
                    new Dictionary<string, PreparedWordBulkFormula>(
                        StringComparer.Ordinal)
                    {
                        [toOmmlTarget.Id] =
                            new PreparedWordBulkFormula
                            {
                                Run = new WordBulkRun
                                {
                                    Id = toOmmlTarget.Id,
                                    IsFormula = true,
                                    Latex = toOmmlTarget.Latex,
                                    DisplayMode =
                                        toOmmlTarget.DisplayMode,
                                },
                                Session =
                                    CreateSimpleFormatTargetSession(
                                        toOmmlTarget,
                                        FormulaOleContract.WordOmmlMode,
                                        editedMathMl),
                                MathMl =
                                    editedMathMl,
                            },
                    };

                Release(suffix);
                suffix = null;
                Release(paragraphRange);
                paragraphRange = null;
                Release(paragraph);
                paragraph = null;
                Release(paragraphs);
                paragraphs = null;
                Release(visualTeXRange);
                visualTeXRange = null;
                Release(visualTeXShape);
                visualTeXShape = null;
                Release(visualTeXShapes);
                visualTeXShapes = null;

                var toOmmlResult =
                    service.ApplyFormulaFormatConversionPlan(
                        toOmmlPlan,
                        toOmmlPrepared);
                AssertEqual(
                    1,
                    toOmmlResult.FormulaCount,
                    "VisualTeX→OMML did not convert exactly one formula.");
                AssertEqual(
                    0,
                    toOmmlResult.FailedFormulaCount,
                    "VisualTeX→OMML reported a conversion failure.");
            }
            finally
            {
                Release(visualTeXRange);
                Release(visualTeXShape);
                Release(visualTeXShapes);
            }

            maths =
                document.OMaths;
            AssertEqual(
                1,
                maths.Count,
                "VisualTeX→OMML did not restore exactly one native OMath.");
            AssertEqual(
                0,
                CountVisualTeXNativeOleShapes(document),
                "VisualTeX→OMML left a VisualTeX OLE behind.");
            math =
                maths[1];
            AssertEqual(
                Word.WdOMathType.wdOMathInline,
                math.Type,
                "VisualTeX→OMML lost inline type.");
            mathRange =
                math.Range.Duplicate;

            var roundTripHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    mathRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Round-tripped OMML cannot be resolved from its actual OMath range.");
            roundTripHost.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    roundTripHost);
            var roundTripPayload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    roundTripHost);
            AssertTrue(
                string.IsNullOrWhiteSpace(
                    roundTripHost.FormulaId),
                "Pure OMML unexpectedly exposed a durable VisualTeX FormulaId.");
            AssertTrue(
                roundTripPayload.Latex.IndexOf(
                    "y",
                    StringComparison.OrdinalIgnoreCase) >= 0
                && roundTripPayload.Latex.IndexOf(
                    "2",
                    StringComparison.OrdinalIgnoreCase) >= 0,
                "OMML→VisualTeX→OMML changed formula semantics.");

            AssertNoHostCoreSentinelArtifacts(
                document,
                "after OMML→VisualTeX→OMML");

            prefix =
                document.Range(
                    0,
                    mathRange.Start);
            AssertEqual(
                "left",
                prefix.Text ?? string.Empty,
                "Round-trip changed text before the inline formula.");
            paragraphs =
                mathRange.Paragraphs;
            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            bodyEnd =
                Math.Max(
                    paragraphRange.Start,
                    paragraphRange.End - 1);
            suffix =
                document.Range(
                    mathRange.End,
                    bodyEnd);
            AssertEqual(
                " right",
                suffix.Text ?? string.Empty,
                "Round-trip changed text after the inline formula.");

            bookmarks =
                document.Bookmarks;
            AssertEqual(
                0,
                CountHostCoreBookmarks(
                    bookmarks,
                    WordFormulaIdentityStore.OmmlBookmarkPrefix),
                "Pure OMML round-trip created a VTOMML identity bookmark.");

            Console.WriteLine(
                "[host-core minimal] inline OMML insert/edit and OMML↔VisualTeX round-trip passed.");

            Release(identityRange);
            identityRange = null;
            Release(identity);
            identity = null;
            Release(bookmarks);
            bookmarks = null;
            Release(suffix);
            suffix = null;
            Release(paragraphRange);
            paragraphRange = null;
            Release(paragraph);
            paragraph = null;
            Release(paragraphs);
            paragraphs = null;
            Release(prefix);
            prefix = null;
            Release(mathRange);
            mathRange = null;
            Release(math);
            math = null;
            Release(maths);
            maths = null;

            RunHostCoreCanonicalNumberingSmoke(
                host.Application,
                document,
                service);

            RunHostCoreNativeFormulaToLatexSmoke(
                host.Application,
                document);

            RunHostCoreNumberedFormatRoundTripSmoke(
                host.Application,
                document,
                pngPath,
                emfPath);

            RunHostCoreNativeLatexRedrawSmoke(
                host.Application,
                document);

            RunHostCoreNativeHeadingNumberingSmoke(
                host.Application,
                document);

            RunHostCoreAdjacentNativeDeleteSmoke(
                host.Application,
                document);

            RunHostCoreUserTableNumberingSmoke(
                host.Application,
                document,
                service);

            RunHostCoreNativeRawPasteProbe(
                host.Application,
                document,
                service);

            RunHostCoreNativeNumberCopyPasteUndoSmoke(
                host.Application,
                document,
                service);

            RunHostCoreLazyNativePasteAdoptionSmoke(
                host.Application,
                document,
                service);

            RunHostCoreUnownedNativeReferenceAdoptionSmoke(
                host.Application,
                document,
                service);

            RunHostCoreUnownedNativeNumberRefreshSmoke(
                host.Application,
                document);

            RunHostCoreNativeNumberGroupPasteUndoSmoke(
                host.Application,
                document);

            RunHostCoreNativeFontSizeAdoptionSmoke(
                host.Application,
                document);

            RunHostCoreNativeNumberToggleSmoke(
                host.Application,
                document);

            RunHostCoreMathTypeBoundarySmoke(
                host.Application,
                document,
                service,
                pngPath,
                emfPath);

            RunHostCoreNativeSaveReopenSmoke(
                host.Application,
                document,
                artifactRoot);

            RunHostCoreUnownedNativeSaveReopenAdoptionSmoke(
                host.Application,
                document,
                artifactRoot);
        }
        finally
        {
            Release(identityRange);
            Release(identity);
            Release(bookmarks);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(suffix);
            Release(prefix);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(insertion);
        }
    }

    private static void RunHostCoreNativeNumberCopyPasteSmoke(
        Word.Application application,
        Word.Document document,
        WordFormulaService service)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>p</mi><mo>=</mo><msup><mi>q</mi><mn>2</mn></msup></mrow></math>";

        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? sourceIdentity = null;
        Word.Bookmark? pastedIdentity = null;
        Word.Range? sourceIdentityRange = null;
        Word.Range? pastedIdentityRange = null;
        Word.Range? sourceRange = null;
        Word.Range? pastedRegion = null;
        Word.Range? pastedRange = null;
        Word.UndoRecord? undoRecord = null;
        var undoStarted = false;
        try
        {
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();

            var session =
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    session,
                    mathMl);

            bookmarks = document.Bookmarks;
            var sourceIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    created.FormulaId);
            AssertTrue(
                bookmarks.Exists(sourceIdentityName),
                "Native copy/paste source lost its OMML identity.");
            sourceIdentity = bookmarks[sourceIdentityName];
            sourceIdentityRange =
                sourceIdentity.Range.Duplicate;

            var sourceHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    sourceIdentityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Native copy/paste source cannot be resolved.");
            sourceHost.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    sourceHost);
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                sourceHost.Numbering.ContainerKind,
                "Native copy/paste source is not one Word-native #(SEQ) host.");

            sourceRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    sourceHost.Range);
            sourceRange.Select();

            var snapshot =
                service.CaptureSelectedFormulaForCopy()
                ?? throw new InvalidDataException(
                    "Host-core copy snapshot was not created for native numbered OMML.");
            AssertTrue(
                snapshot.UsesHostCore
                && snapshot.Metadata.Numbered,
                "Native numbered OMML copy snapshot lost host-core/numbered state.");

            selection.Copy();
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            var pasteStart =
                selection.Start;

            undoRecord =
                application.UndoRecord;
            if (!undoRecord.IsRecordingCustomRecord
                && undoRecord.CustomRecordLevel == 0)
            {
                undoRecord.StartCustomRecord(
                    "VisualTeX Native OMML Paste Acceptance");
                undoStarted = true;
            }

            service.ArmHostCorePaste(
                snapshot);
            selection.Paste();

            var repair =
                service.RepairPastedFormula(
                    snapshot);
            for (var attempt = 0;
                 repair == WordFormulaService.PastedFormulaRepairResult.NotReady
                 && attempt < 4;
                 attempt++)
            {
                System.Threading.Thread.Sleep(20);
                repair =
                    service.RepairPastedFormula(
                        snapshot);
            }
            AssertEqual(
                WordFormulaService.PastedFormulaRepairResult.Repaired,
                repair,
                "Native numbered OMML paste was not repaired by the host core.");

            if (undoStarted)
            {
                undoRecord.EndCustomRecord();
                undoStarted = false;
            }

            var pasteEnd =
                Math.Max(
                    pasteStart + 1,
                    selection.End);
            pastedRegion =
                document.Range(
                    pasteStart,
                    Math.Min(
                        document.Content.End,
                        pasteEnd));
            // Re-keying rebuilds the native OMath and Word is free to
            // expand its Range beyond the pre-repair paste caret. The acceptance
            // therefore reindexes the owned document after repair instead of
            // treating the stale physical paste range as semantic truth.
            var pastedIndex =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            var pastedCandidates =
                pastedIndex.Omml
                    .Where(item =>
                        item.Range.Start >= pasteStart)
                    .OrderBy(item =>
                        item.Range.Start)
                    .ToArray();
            AssertEqual(
                1,
                pastedCandidates.Length,
                "Native paste region does not contain exactly one pasted OMML host.");
            var pastedHost =
                pastedCandidates[0];
            AssertTrue(
                !string.IsNullOrWhiteSpace(
                    pastedHost.FormulaId),
                "Pasted native OMML received no fresh FormulaId.");
            AssertTrue(
                !string.Equals(
                    created.FormulaId,
                    pastedHost.FormulaId,
                    StringComparison.OrdinalIgnoreCase),
                "Pasted native OMML retained the source FormulaId.");

            pastedHost.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    pastedHost);
            pastedRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    pastedHost.Range);
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                pastedHost.Numbering.ContainerKind,
                "Pasted native OMML did not rebuild one canonical Word #(SEQ) host. "
                + WordNativeOmmlNumbering.DescribeNativeState(
                    document,
                    pastedRange));

            var sourceNumber =
                WordNativeOmmlNumbering.ReadVisibleNumber(
                    document,
                    sourceHost.Numbering);
            var pastedNumber =
                WordNativeOmmlNumbering.ReadVisibleNumber(
                    document,
                    pastedHost.Numbering);
            AssertEqual(
                "1",
                sourceNumber,
                "Native copy/paste source SEQ result changed unexpectedly.");
            AssertEqual(
                "2",
                pastedNumber,
                "Pasted native OMML did not receive the next Word SEQ result.");

            var sourceTarget =
                WordFormulaNumberingKernel.ReferenceBookmarkName(
                    created.FormulaId);
            var pastedTarget =
                WordFormulaNumberingKernel.ReferenceBookmarkName(
                    pastedHost.FormulaId!);
            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(sourceTarget),
                "Native copy/paste source lost its VTEqNum target.");
            AssertTrue(
                bookmarks.Exists(pastedTarget),
                "Pasted native OMML has no fresh VTEqNum target.");
            AssertTrue(
                !string.Equals(
                    sourceTarget,
                    pastedTarget,
                    StringComparison.OrdinalIgnoreCase),
                "Pasted native OMML reused the source reference target.");

            var pastedIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    pastedHost.FormulaId!);
            AssertTrue(
                bookmarks.Exists(pastedIdentityName),
                "Pasted native OMML has no fresh OMML identity.");
            pastedIdentity =
                bookmarks[pastedIdentityName];
            pastedIdentityRange =
                pastedIdentity.Range.Duplicate;
            pastedRange.Select();

            var editSelection =
                service.ReadSelection();
            AssertEqual(
                pastedHost.FormulaId!,
                editSelection.FormulaId,
                "ReadSelection changed the pasted native OMML FormulaId.");
            AssertTrue(
                editSelection.Metadata is not null
                && editSelection.Metadata.Numbered,
                "ReadSelection lost numbered metadata after native paste repair.");
            AssertTrue(
                editSelection.Metadata!.Latex.IndexOf(
                    "#",
                    StringComparison.Ordinal) < 0
                && editSelection.Metadata.Latex.IndexOf(
                    "VisualTeXEquation",
                    StringComparison.OrdinalIgnoreCase) < 0,
                "Pasted native OMML exposed numbering syntax to the editor.");

            AssertNoHostCoreSentinelArtifacts(
                document,
                "after Word-native OMML copy/paste");

            Console.WriteLine(
                "[host-core minimal] Word-native numbered OMML copy/paste + editor isolation passed.");
        }
        finally
        {
            if (undoStarted)
            {
                try { undoRecord?.EndCustomRecord(); } catch { }
            }
            Release(undoRecord);
            Release(pastedRange);
            Release(pastedRegion);
            Release(pastedIdentityRange);
            Release(sourceIdentityRange);
            Release(pastedIdentity);
            Release(sourceIdentity);
            Release(bookmarks);
            Release(sourceRange);
            Release(selection);
        }
    }

    private static void RunHostCoreNativeRawPasteProbe(
        Word.Application application,
        Word.Document document,
        WordFormulaService service)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>u</mi><mo>=</mo><mi>v</mi></mrow></math>";

        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? identity = null;
        Word.Range? identityRange = null;
        Word.Range? sourceRange = null;
        Word.Range? tail = null;
        Word.OMaths? maths = null;
        Word.OMath? pastedMath = null;
        Word.Range? pastedRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();

            var session =
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    session,
                    mathMl);

            bookmarks = document.Bookmarks;
            var identityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    created.FormulaId);
            AssertTrue(
                bookmarks.Exists(identityName),
                "Raw-paste source lost its OMML identity.");
            identity = bookmarks[identityName];
            identityRange = identity.Range.Duplicate;
            var sourceHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    identityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Raw-paste source cannot be resolved.");
            sourceRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    sourceHost.Range);
            sourceRange.Select();
            selection.Copy();

            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            var pasteStart = selection.Start;
            selection.Paste();
            System.Threading.Thread.Sleep(180);

            tail = document.Range(
                pasteStart,
                document.Content.End);
            maths = tail.OMaths;
            AssertTrue(
                maths.Count >= 1,
                "Native Word paste produced no OMath.");
            pastedMath = maths[1];
            pastedRange = pastedMath.Range.Duplicate;
            AssertEqual(
                Word.WdOMathType.wdOMathDisplay,
                pastedMath.Type,
                "Native Word paste changed numbered OMath display mode.");
            AssertTrue(
                WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                    pastedRange.WordOpenXML,
                    formulaId: null),
                "Native Word paste lost the m:eqArr + #(SEQ) structure.");

            fields = pastedRange.Fields;
            var sequenceCount = 0;
            var rawPastedSequenceText = string.Empty;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code.Duplicate;
                if ((code.Text ?? string.Empty)
                    .TrimStart()
                    .StartsWith(
                        "SEQ VisualTeXEquation",
                        StringComparison.OrdinalIgnoreCase))
                {
                    sequenceCount++;
                    Word.Range? rawResult = null;
                    try
                    {
                        rawResult = field.Result.Duplicate;
                        rawPastedSequenceText =
                            (rawResult.Text ?? string.Empty).Trim();
                    }
                    finally { Release(rawResult); }
                }
            }
            AssertEqual(
                1,
                sequenceCount,
                "Native Word paste did not preserve exactly one live SEQ field.");
            Console.WriteLine(
                $"[host-core raw paste] sequenceResultBeforeUpdate={rawPastedSequenceText}");

            // Native Word paste deliberately keeps the field's cached
            // display result until the next ordinary Word field refresh (F9,
            // print/update, or an explicit refresh command). Do not auto-update
            // it here: Field.Update is a separate native Undo action and would
            // make Ctrl+Z undo the number refresh before undoing the paste.

            var copiedIdentityCount = 0;
            var copiedNumberTargetCount = 0;
            var localBookmarks = pastedRange.Bookmarks;
            try
            {
                for (var index = 1; index <= localBookmarks.Count; index++)
                {
                    Word.Bookmark? bookmark = null;
                    try
                    {
                        bookmark = localBookmarks[index];
                        var name = bookmark.Name ?? string.Empty;
                        if (name.StartsWith(
                                WordFormulaIdentityStore.OmmlBookmarkPrefix,
                                StringComparison.OrdinalIgnoreCase))
                            copiedIdentityCount++;
                        if (name.StartsWith(
                                "VTEqNum_",
                                StringComparison.OrdinalIgnoreCase))
                            copiedNumberTargetCount++;
                    }
                    finally { Release(bookmark); }
                }
            }
            finally { Release(localBookmarks); }

            Console.WriteLine(
                $"[host-core raw paste] identityBookmarks={copiedIdentityCount} "
                + $"numberTargets={copiedNumberTargetCount} "
                + $"range={pastedRange.Start}:{pastedRange.End}");

            object one = 1;
            AssertTrue(
                document.Undo(ref one),
                "Word could not undo one raw native OMML paste.");

            Release(tail);
            tail = document.Range(
                pasteStart,
                document.Content.End);
            maths = tail.OMaths;
            AssertEqual(
                0,
                maths.Count,
                "One native Word Undo did not remove the raw pasted OMath.");

            Console.WriteLine(
                "[host-core minimal] raw Word-native numbered OMML paste + one native Undo passed.");
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(pastedRange);
            Release(pastedMath);
            Release(maths);
            Release(tail);
            Release(sourceRange);
            Release(identityRange);
            Release(identity);
            Release(bookmarks);
            Release(selection);
        }
    }

    private static void RunHostCoreNativeNumberCopyPasteUndoSmoke(
        Word.Application application,
        Word.Document document,
        WordFormulaService service)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>q</mi><mo>=</mo><msup><mi>r</mi><mn>2</mn></msup></mrow></math>";

        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? sourceIdentity = null;
        Word.Range? sourceIdentityRange = null;
        Word.Range? sourceRange = null;
        Word.Range? pastedRange = null;
        Word.UndoRecord? undoRecord = null;
        try
        {
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();

            var session =
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    session,
                    mathMl);

            bookmarks = document.Bookmarks;
            var sourceIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    created.FormulaId);
            AssertTrue(
                bookmarks.Exists(sourceIdentityName),
                "Native numbered copy source lost its OMML identity.");
            sourceIdentity = bookmarks[sourceIdentityName];
            sourceIdentityRange =
                sourceIdentity.Range.Duplicate;

            var sourceHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    sourceIdentityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Native numbered copy source cannot be resolved.");
            sourceHost.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    sourceHost);
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                sourceHost.Numbering.ContainerKind,
                "Copy source is not Word-native #(SEQ).");

            sourceRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    sourceHost.Range);
            sourceRange.Select();

            var snapshot =
                service.CaptureSelectedFormulaForCopy()
                ?? throw new InvalidDataException(
                    "Host core did not capture the selected native numbered OMML.");
            AssertTrue(
                snapshot.UsesHostCore,
                "Native numbered OMML copy fell back to the legacy copy path.");
            AssertTrue(
                snapshot.Metadata.Numbered,
                "Native numbered OMML copy snapshot lost numbered state.");
            AssertTrue(
                (snapshot.CoreLatex ?? string.Empty)
                    .IndexOf("#", StringComparison.Ordinal) < 0
                && (snapshot.CoreLatex ?? string.Empty)
                    .IndexOf(
                        "VisualTeXEquation",
                        StringComparison.OrdinalIgnoreCase) < 0,
                "Copy snapshot exposed native numbering to formula semantics.");

            selection.Copy();
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            var pasteStart =
                selection.Range.Start;

            // Product behavior: native OMML paste is one ordinary Word
            // paste. VisualTeX arms a bounded local observation, but performs no
            // post-paste document mutation and does not open a Custom UndoRecord.
            service.ArmHostCorePaste(
                snapshot);
            selection.Paste();
            System.Threading.Thread.Sleep(120);

            var repair =
                service.RepairPastedFormula(
                    snapshot);
            AssertEqual(
                WordFormulaService.PastedFormulaRepairResult.Repaired,
                repair,
                "Native numbered OMML paste was not recognized by the host core.");
            System.Threading.Thread.Sleep(120);

            var index =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            var pastedCandidates =
                index.Omml
                    .Where(item =>
                        item.Range.Start >= pasteStart)
                    .OrderBy(item => item.Range.Start)
                    .ToArray();
            AssertEqual(
                1,
                pastedCandidates.Length,
                "Native numbered paste did not create exactly one OMML host in the paste region.");

            var pastedHost =
                pastedCandidates[0];
            pastedHost.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    pastedHost);
            AssertTrue(
                !string.Equals(
                    pastedHost.FormulaId,
                    created.FormulaId,
                    StringComparison.OrdinalIgnoreCase),
                "Pasted native numbered OMML incorrectly retained the source FormulaId.");
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                pastedHost.Numbering.ContainerKind,
                "Pasted numbered OMML is not one Word-native #(SEQ) host.");

            var pastedPayload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    pastedHost);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    pastedPayload.MathMl
                    ?? throw new InvalidDataException(
                        "Pasted native numbered OMML returned no semantic MathML.")),
                "Pasted native numbered OMML changed formula semantics.");
            AssertTrue(
                pastedPayload.Latex.IndexOf(
                    "#",
                    StringComparison.Ordinal) < 0
                && pastedPayload.Latex.IndexOf(
                    "VisualTeXEquation",
                    StringComparison.OrdinalIgnoreCase) < 0,
                "Pasted native numbered OMML exposes the number to editing.");

            pastedRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    pastedHost.Range);
            pastedRange.Select();
            var editSelection =
                service.ReadSelection();
            AssertTrue(
                editSelection.Metadata is not null
                && editSelection.Metadata.Numbered,
                "ReadSelection did not recognize the pasted native equation as numbered.");
            AssertTrue(
                !string.IsNullOrWhiteSpace(
                    editSelection.FormulaId)
                && !string.Equals(
                    editSelection.FormulaId,
                    created.FormulaId,
                    StringComparison.OrdinalIgnoreCase),
                "ReadSelection did not allocate an independent session FormulaId for the unowned paste.");
            AssertTrue(
                editSelection.Metadata!.Latex.IndexOf(
                    "#",
                    StringComparison.Ordinal) < 0
                && editSelection.Metadata.Latex.IndexOf(
                    "VisualTeXEquation",
                    StringComparison.OrdinalIgnoreCase) < 0,
                "Double-click/read semantics exposed the Word-native number.");

            var sourceRef =
                WordNativeOmmlNumbering.ReferenceBookmarkName(
                    created.FormulaId);
            var transientIdentity =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    editSelection.FormulaId!);
            var transientRef =
                WordNativeOmmlNumbering.ReferenceBookmarkName(
                    editSelection.FormulaId!);
            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(sourceRef),
                "Native copy/paste source lost its VTEqNum target.");
            AssertTrue(
                !bookmarks.Exists(transientIdentity)
                && !bookmarks.Exists(transientRef),
                "ReadSelection mutated the pasted Word equation while merely opening it.");

            object one = 1;
            var undoResult =
                document.Undo(ref one);
            Release(bookmarks);
            bookmarks = document.Bookmarks;
            var undoIndex =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            var undoPastedHosts =
                undoIndex.Omml
                    .Where(item =>
                        item.Range.Start >= pasteStart)
                    .ToArray();
            Console.WriteLine(
                $"[host-core native paste undo] result={undoResult} "
                + $"pastedHosts={undoPastedHosts.Length} "
                + $"sourceIdentity={bookmarks.Exists(sourceIdentityName)} "
                + $"sourceRef={bookmarks.Exists(sourceRef)}");

            AssertTrue(
                undoResult,
                "Word returned false while undoing the native paste transaction.");
            AssertEqual(
                0,
                undoPastedHosts.Length,
                "One Undo did not remove the pasted native OMML host.");
            AssertTrue(
                bookmarks.Exists(sourceIdentityName),
                "Undo removed the original numbered OMML instead of only the paste transaction.");
            AssertTrue(
                bookmarks.Exists(sourceRef),
                "Undo removed the original equation reference target.");

            AssertNoHostCoreSentinelArtifacts(
                document,
                "after native numbered OMML copy/paste + one Undo");

            Console.WriteLine(
                "[host-core minimal] Word-native numbered OMML copy/paste + read-only lazy identity + one-record Undo passed.");
        }
        finally
        {
            if (undoRecord is not null)
            {
                try
                {
                    if (undoRecord.IsRecordingCustomRecord
                        || undoRecord.CustomRecordLevel > 0)
                        undoRecord.EndCustomRecord();
                }
                catch { }
            }
            Release(undoRecord);
            Release(pastedRange);
            Release(sourceRange);
            Release(sourceIdentityRange);
            Release(sourceIdentity);
            Release(bookmarks);
            Release(selection);
        }
    }

    private static void RunHostCoreLazyNativePasteAdoptionSmoke(
        Word.Application application,
        Word.Document document,
        WordFormulaService service)
    {
        const string sourceMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>p</mi><mo>=</mo><mi>q</mi></mrow></math>";
        const string editedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>p</mi><mo>=</mo><mi>q</mi><mo>+</mo><mn>1</mn></mrow></math>";

        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? sourceIdentity = null;
        Word.Range? sourceIdentityRange = null;
        Word.Range? sourceRange = null;
        Word.Range? tail = null;
        Word.OMaths? maths = null;
        Word.OMath? pastedMath = null;
        Word.Range? pastedRange = null;
        Word.Bookmark? adoptedIdentity = null;
        Word.Range? adoptedIdentityRange = null;
        try
        {
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();

            var createSession =
                CreateOmmlMathTypeAcceptanceSession(
                    sourceMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    createSession,
                    sourceMathMl);

            bookmarks = document.Bookmarks;
            var sourceIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    created.FormulaId);
            sourceIdentity = bookmarks[sourceIdentityName];
            sourceIdentityRange =
                sourceIdentity.Range.Duplicate;
            var sourceHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    sourceIdentityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Lazy-adoption source cannot be resolved.");
            sourceRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    sourceHost.Range);
            sourceRange.Select();

            var snapshot =
                service.CaptureSelectedFormulaForCopy()
                ?? throw new InvalidDataException(
                    "Lazy-adoption source was not captured for copy.");
            selection.Copy();

            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            var pasteStart = selection.Start;
            service.ArmHostCorePaste(snapshot);
            selection.Paste();
            System.Threading.Thread.Sleep(150);
            AssertEqual(
                WordFormulaService.PastedFormulaRepairResult.Repaired,
                service.RepairPastedFormula(snapshot),
                "Lazy-adoption paste observation did not complete.");

            tail = document.Range(
                pasteStart,
                document.Content.End);
            maths = tail.OMaths;
            AssertTrue(
                maths.Count >= 1,
                "Lazy-adoption paste produced no OMath.");
            pastedMath = maths[1];
            pastedRange =
                pastedMath.Range.Duplicate;
            pastedRange.Select();

            var opened =
                service.ReadSelection();
            AssertTrue(
                opened.Metadata is not null
                && opened.Metadata.Numbered,
                "Lazy-adoption ReadSelection lost numbered state.");
            AssertTrue(
                !string.IsNullOrWhiteSpace(opened.FormulaId)
                && !string.Equals(
                    opened.FormulaId,
                    created.FormulaId,
                    StringComparison.OrdinalIgnoreCase),
                "Lazy-adoption ReadSelection did not allocate a fresh session FormulaId.");
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(sourceMathMl),
                MathTypeMtefCodec.SemanticSignature(
                    WordFormulaHostSemanticReader.Read(
                        document,
                        WordFormulaHostResolver.ResolveLocal(
                            document,
                            pastedRange,
                            WordFormulaHostKind.Omml)
                        ?? throw new InvalidDataException(
                            "Unowned pasted OMath cannot be resolved before edit."))
                    .MathMl
                    ?? throw new InvalidDataException(
                        "Unowned pasted OMath returned no semantic MathML.")),
                "Unowned pasted OMath semantics changed before adoption.");

            var adoptedIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    opened.FormulaId!);
            var adoptedRefName =
                WordNativeOmmlNumbering.ReferenceBookmarkName(
                    opened.FormulaId!);
            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertTrue(
                !bookmarks.Exists(adoptedIdentityName)
                && !bookmarks.Exists(adoptedRefName),
                "ReadSelection persisted lazy identity before an edit occurred.");

            var editSession =
                CreateOmmlMathTypeAcceptanceSession(
                    editedMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            editSession.Mode = "edit";
            editSession.FormulaId =
                opened.FormulaId!;
            editSession.SourceDocumentId =
                opened.DocumentId;
            editSession.SourceObjectId =
                opened.ObjectId;
            editSession.OriginalMetadata =
                opened.Metadata;

            _ = service.ReplaceOmml(
                editSession,
                editedMathMl);

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(adoptedIdentityName),
                "First edit did not persist the lazy OMML FormulaId.");
            AssertTrue(
                bookmarks.Exists(adoptedRefName),
                "First edit did not create the fresh native VTEqNum target.");

            adoptedIdentity =
                bookmarks[adoptedIdentityName];
            adoptedIdentityRange =
                adoptedIdentity.Range.Duplicate;
            var adopted =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    adoptedIdentityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Adopted pasted OMML cannot be resolved from its persisted identity.");
            adopted.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    adopted);
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                adopted.Numbering.ContainerKind,
                "First edit did not preserve Word-native #(SEQ) numbering.");

            var adoptedPayload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    adopted);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(editedMathMl),
                MathTypeMtefCodec.SemanticSignature(
                    adoptedPayload.MathMl
                    ?? throw new InvalidDataException(
                        "Adopted pasted OMML returned no MathML.")),
                "First edit changed the requested semantic formula.");
            AssertTrue(
                adoptedPayload.Latex.IndexOf(
                    "#",
                    StringComparison.Ordinal) < 0
                && adoptedPayload.Latex.IndexOf(
                    "VisualTeXEquation",
                    StringComparison.OrdinalIgnoreCase) < 0,
                "Adopted pasted formula exposes its native number to the editor.");

            AssertNoHostCoreSentinelArtifacts(
                document,
                "after lazy native paste adoption");

            Console.WriteLine(
                "[host-core minimal] lazy native OMML paste adoption on first edit passed.");
        }
        finally
        {
            Release(adoptedIdentityRange);
            Release(adoptedIdentity);
            Release(pastedRange);
            Release(pastedMath);
            Release(maths);
            Release(tail);
            Release(sourceRange);
            Release(sourceIdentityRange);
            Release(sourceIdentity);
            Release(bookmarks);
            Release(selection);
        }
    }

    private static void RunHostCoreUnownedNativeReferenceAdoptionSmoke(
        Word.Application application,
        Word.Document document,
        WordFormulaService service)
    {
        const string sourceMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>r</mi><mo>=</mo><mi>s</mi><mo>+</mo><mn>4</mn></mrow></math>";

        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? sourceIdentity = null;
        Word.Range? sourceIdentityRange = null;
        Word.Range? sourceRange = null;
        Word.Range? tail = null;
        Word.OMaths? maths = null;
        Word.OMath? pastedMath = null;
        Word.Range? pastedRange = null;
        try
        {
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            selection = application.Selection;
            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();

            var createSession =
                CreateOmmlMathTypeAcceptanceSession(
                    sourceMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    createSession,
                    sourceMathMl);

            bookmarks = document.Bookmarks;
            var sourceIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    created.FormulaId);
            sourceIdentity =
                bookmarks[sourceIdentityName];
            sourceIdentityRange =
                sourceIdentity.Range.Duplicate;
            var sourceHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    sourceIdentityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Reference-adoption source cannot be resolved.");
            sourceRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    sourceHost.Range);
            sourceRange.Select();

            var snapshot =
                service.CaptureSelectedFormulaForCopy()
                ?? throw new InvalidDataException(
                    "Reference-adoption source was not captured for copy.");
            selection.Copy();
            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            var pasteStart =
                selection.Range.Start;
            service.ArmHostCorePaste(
                snapshot);
            selection.Paste();
            System.Threading.Thread.Sleep(120);
            AssertEqual(
                WordFormulaService.PastedFormulaRepairResult.Repaired,
                service.RepairPastedFormula(
                    snapshot),
                "Reference-adoption paste validation did not complete.");

            tail =
                document.Range(
                    pasteStart,
                    document.Content.End);
            maths = tail.OMaths;
            AssertTrue(
                maths.Count >= 1,
                "Reference-adoption paste produced no OMath.");
            pastedMath = maths[1];
            pastedRange =
                pastedMath.Range.Duplicate;
            var pastedStart =
                pastedRange.Start;
            var pastedEnd =
                pastedRange.End;

            var unowned =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    pastedRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Reference-adoption pasted OMML cannot be resolved.");
            unowned.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    unowned);
            AssertTrue(
                string.IsNullOrWhiteSpace(
                    unowned.FormulaId),
                "Reference target was already owned before target discovery.");
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                unowned.Numbering.ContainerKind,
                "Reference target is not one unowned native #(SEQ) OMath.");

            var targets =
                service.GetCanonicalEquationReferenceTargets(
                    document);
            var target =
                targets.SingleOrDefault(item =>
                    item.Source ==
                        EquationReferenceSource.VisualTeX
                    && item.Position == pastedStart
                    && item.EndPosition == pastedEnd)
                ?? throw new InvalidDataException(
                    "Read-only reference target discovery omitted the unowned native OMath.");
            AssertTrue(
                !string.Equals(
                    target.FormulaId,
                    created.FormulaId,
                    StringComparison.OrdinalIgnoreCase),
                "Unowned reference target reused the source FormulaId.");

            var adoptedIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    target.FormulaId);
            var adoptedRefName =
                WordNativeOmmlNumbering.ReferenceBookmarkName(
                    target.FormulaId);
            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertTrue(
                !bookmarks.Exists(
                    adoptedIdentityName)
                && !bookmarks.Exists(
                    adoptedRefName),
                "Reference target discovery mutated the anonymous native OMath.");

            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            selection.TypeText(
                "native-ref ");
            var referenceInsertionStart =
                selection.Range.Start;

            service.InsertEquationReference(
                document,
                selection,
                target,
                EquationReferenceStyle.Parenthesized,
                Word.WdColor.wdColorAutomatic);

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(
                    adoptedIdentityName),
                "Reference insertion did not adopt VTOMML ownership.");
            AssertTrue(
                bookmarks.Exists(
                    adoptedRefName),
                "Reference insertion did not create VTEqNum target.");
            AssertTrue(
                ContainsHostCoreReferenceField(
                    document,
                    adoptedRefName),
                "Reference insertion did not create a live Word REF field.");

            var adoptedProbe =
                document.Range(
                    pastedStart,
                    pastedEnd);
            try
            {
                var adopted =
                    WordFormulaHostResolver.ResolveLocal(
                        document,
                        adoptedProbe,
                        WordFormulaHostKind.Omml)
                    ?? throw new InvalidDataException(
                        "Reference-adopted native OMath cannot be resolved.");
                adopted.Numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        adopted);
                AssertEqual(
                    target.FormulaId,
                    adopted.FormulaId
                    ?? string.Empty,
                    "Reference insertion adopted the wrong FormulaId.");
                AssertEqual(
                    WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                    adopted.Numbering.ContainerKind,
                    "Reference insertion changed native #(SEQ) topology.");
                AssertEqual(
                    MathTypeMtefCodec.SemanticSignature(
                        sourceMathMl),
                    MathTypeMtefCodec.SemanticSignature(
                        WordFormulaHostSemanticReader.Read(
                            document,
                            adopted).MathMl
                        ?? throw new InvalidDataException(
                            "Reference-adopted OMath returned no MathML.")),
                    "Reference insertion changed formula semantics.");
            }
            finally
            {
                Release(adoptedProbe);
            }

            AssertTrue(
                selection.Range.Start
                    >= referenceInsertionStart,
                "Reference insertion moved before the captured insertion point.");

            AssertTrue(
                bookmarks.Exists(
                    sourceIdentityName),
                "Reference insertion damaged the original source identity.");
            AssertNoHostCoreSentinelArtifacts(
                document,
                "after unowned native reference adoption");
            Console.WriteLine(
                "[host-core minimal] unowned native OMML reference adoption passed.");
        }
        finally
        {
            Release(pastedRange);
            Release(pastedMath);
            Release(maths);
            Release(tail);
            Release(sourceRange);
            Release(sourceIdentityRange);
            Release(sourceIdentity);
            Release(bookmarks);
            Release(selection);
        }
    }

    private static bool ContainsHostCoreReferenceField(
        Word.Document document,
        string bookmarkName)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            fields = document.Fields;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                field = fields[index];
                code = field.Code;
                var instruction =
                    (code.Text ?? string.Empty)
                    .Trim();
                if (instruction.StartsWith(
                        "REF ",
                        StringComparison.OrdinalIgnoreCase)
                    && instruction.IndexOf(
                        bookmarkName,
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;

                Release(code);
                code = null;
                Release(field);
                field = null;
            }
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static void RunHostCoreUnownedNativeNumberRefreshSmoke(
        Word.Application application,
        Word.Document mainDocument)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>n</mi><mo>=</mo><mi>m</mi><mo>+</mo><mn>1</mn></mrow></math>";

        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? sourceIdentity = null;
        Word.Range? sourceIdentityRange = null;
        Word.Range? sourceRange = null;
        Word.Range? tail = null;
        Word.OMaths? tailMaths = null;
        Word.OMath? pastedMath = null;
        Word.Range? pastedRange = null;
        try
        {
            document =
                application.Documents.Add(
                    Visible: false);
            document.Activate();
            var service =
                new WordFormulaService(
                    application);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            selection = application.Selection;
            selection.SetRange(0, 0);
            var createSession =
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    createSession,
                    mathMl);

            bookmarks = document.Bookmarks;
            var sourceIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    created.FormulaId);
            sourceIdentity =
                bookmarks[sourceIdentityName];
            sourceIdentityRange =
                sourceIdentity.Range.Duplicate;
            var sourceHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    sourceIdentityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Number-refresh source cannot be resolved.");
            sourceRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    sourceHost.Range);
            sourceRange.Select();

            var snapshot =
                service.CaptureSelectedFormulaForCopy()
                ?? throw new InvalidDataException(
                    "Number-refresh source was not captured for copy.");
            selection.Copy();
            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            var pasteStart =
                selection.Range.Start;
            service.ArmHostCorePaste(
                snapshot);
            selection.Paste();
            System.Threading.Thread.Sleep(100);
            AssertEqual(
                WordFormulaService.PastedFormulaRepairResult.Repaired,
                service.RepairPastedFormula(
                    snapshot),
                "Number-refresh native paste validation did not complete.");

            tail =
                document.Range(
                    pasteStart,
                    document.Content.End);
            tailMaths = tail.OMaths;
            AssertEqual(
                1,
                tailMaths.Count,
                "Number-refresh paste region did not contain exactly one OMath.");
            pastedMath = tailMaths[1];
            pastedRange =
                pastedMath.Range.Duplicate;
            var unownedBefore =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    pastedRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Number-refresh pasted OMath cannot be resolved.");
            unownedBefore.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    unownedBefore);
            AssertTrue(
                string.IsNullOrWhiteSpace(
                    unownedBefore.FormulaId),
                "Number-refresh pasted OMath was adopted before refresh.");
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                unownedBefore.Numbering.ContainerKind,
                "Number-refresh pasted OMath is not native #(SEQ).");

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertEqual(
                1,
                CountHostCoreBookmarks(
                    bookmarks,
                    WordFormulaIdentityStore.OmmlBookmarkPrefix),
                "Native paste unexpectedly created VTOMML before number refresh.");
            AssertEqual(
                1,
                CountHostCoreBookmarks(
                    bookmarks,
                    "VTEqNum_"),
                "Native paste unexpectedly created VTEqNum before number refresh.");

            var updated =
                service.UpdateEquationNumbers();
            AssertTrue(
                updated >= 2,
                "Updating equation numbers did not visit source + pasted native OMML.");

            var index =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            AssertEqual(
                2,
                index.Omml.Count,
                "Number refresh changed the source + pasted OMath count.");
            var ownedAfter =
                index.Omml.Single(item =>
                    string.Equals(
                        item.FormulaId,
                        created.FormulaId,
                        StringComparison.OrdinalIgnoreCase));
            var unownedAfter =
                index.Omml.Single(item =>
                    string.IsNullOrWhiteSpace(
                        item.FormulaId));
            ownedAfter.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    ownedAfter);
            unownedAfter.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    unownedAfter);
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                ownedAfter.Numbering.ContainerKind,
                "Owned source lost native numbering after refresh.");
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                unownedAfter.Numbering.ContainerKind,
                "Anonymous pasted OMath lost native numbering after refresh.");

            var ownedNumber =
                WordNativeOmmlNumbering.ReadVisibleNumber(
                    document,
                    ownedAfter.Numbering);
            var unownedNumber =
                WordNativeOmmlNumbering.ReadVisibleNumber(
                    document,
                    unownedAfter.Numbering);
            AssertTrue(
                int.TryParse(
                    ownedNumber,
                    out var ownedOrdinal)
                && int.TryParse(
                    unownedNumber,
                    out var unownedOrdinal)
                && unownedOrdinal == ownedOrdinal + 1,
                "Word SEQ did not naturally renumber the copied native equation. "
                + $"source=[{ownedNumber}] pasted=[{unownedNumber}]");

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertEqual(
                1,
                CountHostCoreBookmarks(
                    bookmarks,
                    WordFormulaIdentityStore.OmmlBookmarkPrefix),
                "Number refresh adopted anonymous native OMML by creating VTOMML.");
            AssertEqual(
                1,
                CountHostCoreBookmarks(
                    bookmarks,
                    "VTEqNum_"),
                "Number refresh adopted anonymous native OMML by creating VTEqNum.");

            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    unownedAfter);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(
                    mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    payload.MathMl
                    ?? throw new InvalidDataException(
                        "Refreshed anonymous native OMML returned no MathML.")),
                "Number refresh changed copied native formula semantics.");
            AssertTrue(
                payload.Latex.IndexOf(
                    "#",
                    StringComparison.Ordinal) < 0
                && payload.Latex.IndexOf(
                    "VisualTeXEquation",
                    StringComparison.OrdinalIgnoreCase) < 0,
                "Number refresh leaked native numbering into edit semantics.");

            Console.WriteLine(
                "[host-core minimal] anonymous native OMML number refresh stayed unowned and Word SEQ renumbered naturally.");
        }
        finally
        {
            Release(pastedRange);
            Release(pastedMath);
            Release(tailMaths);
            Release(tail);
            Release(sourceRange);
            Release(sourceIdentityRange);
            Release(sourceIdentity);
            Release(bookmarks);
            Release(selection);
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
            try { mainDocument.Activate(); } catch { }
        }
    }

    private static void RunHostCoreNativeNumberGroupPasteUndoSmoke(
        Word.Application application,
        Word.Document mainDocument)
    {
        const string firstMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>g</mi><mo>=</mo><mn>1</mn></mrow></math>";
        const string secondMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>h</mi><mo>=</mo><mn>2</mn></mrow></math>";

        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? firstIdentity = null;
        Word.Bookmark? secondIdentity = null;
        Word.Range? firstIdentityRange = null;
        Word.Range? secondIdentityRange = null;
        Word.Range? firstRange = null;
        Word.Range? secondRange = null;
        Word.Range? groupRange = null;
        try
        {
            document =
                application.Documents.Add(
                    Visible: false);
            document.Activate();
            var service =
                new WordFormulaService(
                    application);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            selection = application.Selection;
            selection.SetRange(0, 0);

            var firstSession =
                CreateOmmlMathTypeAcceptanceSession(
                    firstMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var first =
                service.InsertOmml(
                    firstSession,
                    firstMathMl);

            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            var secondSession =
                CreateOmmlMathTypeAcceptanceSession(
                    secondMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var second =
                service.InsertOmml(
                    secondSession,
                    secondMathMl);

            bookmarks = document.Bookmarks;
            var firstIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    first.FormulaId);
            var secondIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    second.FormulaId);
            firstIdentity =
                bookmarks[firstIdentityName];
            secondIdentity =
                bookmarks[secondIdentityName];
            firstIdentityRange =
                firstIdentity.Range.Duplicate;
            secondIdentityRange =
                secondIdentity.Range.Duplicate;
            var firstHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    firstIdentityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Group-copy first source cannot be resolved.");
            var secondHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    secondIdentityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Group-copy second source cannot be resolved.");
            firstRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    firstHost.Range);
            secondRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    secondHost.Range);
            groupRange =
                document.Range(
                    Math.Min(
                        firstRange.Start,
                        secondRange.Start),
                    Math.Max(
                        firstRange.End,
                        secondRange.End));
            groupRange.Select();

            var snapshot =
                service.CaptureSelectedFormulaForCopy()
                ?? throw new InvalidDataException(
                    "Native OMML group was not captured for copy.");
            AssertTrue(
                snapshot.UsesHostCore
                && snapshot.GroupItems.Count == 2,
                "Two native numbered OMML equations did not capture as one host-core group.");

            selection.Copy();
            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            var pasteStart =
                selection.Range.Start;
            service.ArmHostCorePaste(
                snapshot);
            selection.Paste();
            System.Threading.Thread.Sleep(120);
            AssertEqual(
                WordFormulaService.PastedFormulaRepairResult.Repaired,
                service.RepairPastedFormula(
                    snapshot),
                "Native numbered OMML group paste validation did not complete.");

            var afterPaste =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            var pasted =
                afterPaste.Omml
                    .Where(item =>
                        item.Range.Start >= pasteStart)
                    .OrderBy(item =>
                        item.Range.Start)
                    .ToArray();
            AssertEqual(
                2,
                pasted.Length,
                "Native group paste did not create exactly two OMath hosts.");
            foreach (var item in pasted)
            {
                item.Numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        item);
                AssertTrue(
                    string.IsNullOrWhiteSpace(
                        item.FormulaId),
                    "Group paste repair assigned durable identity during Paste.");
                AssertEqual(
                    WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                    item.Numbering.ContainerKind,
                    "A pasted group item is not native #(SEQ).");
            }

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertEqual(
                2,
                CountHostCoreBookmarks(
                    bookmarks,
                    WordFormulaIdentityStore.OmmlBookmarkPrefix),
                "Group paste created VTOMML identities before adoption.");
            AssertEqual(
                2,
                CountHostCoreBookmarks(
                    bookmarks,
                    "VTEqNum_"),
                "Group paste created VTEqNum targets before adoption.");

            object one = 1;
            var undone =
                document.Undo(
                    ref one);
            AssertTrue(
                undone,
                "Word returned false while undoing one native group paste.");
            var afterUndo =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            AssertEqual(
                0,
                afterUndo.Omml.Count(item =>
                    item.Range.Start >= pasteStart),
                "One Undo did not remove the complete pasted native OMML group.");
            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(
                    firstIdentityName)
                && bookmarks.Exists(
                    secondIdentityName),
                "Group-paste Undo damaged the two original OMML identities.");
            AssertEqual(
                2,
                afterUndo.Omml.Count,
                "Group-paste Undo changed the two original OMath hosts.");

            Console.WriteLine(
                "[host-core minimal] two-equation native OMML group paste stayed unowned and one Word Undo removed the whole paste.");
        }
        finally
        {
            Release(groupRange);
            Release(secondRange);
            Release(firstRange);
            Release(secondIdentityRange);
            Release(firstIdentityRange);
            Release(secondIdentity);
            Release(firstIdentity);
            Release(bookmarks);
            Release(selection);
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
            try { mainDocument.Activate(); } catch { }
        }
    }

    private static void RunHostCoreNativeFontSizeAdoptionSmoke(
        Word.Application application,
        Word.Document mainDocument)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>f</mi><mo>=</mo><mfrac><mn>1</mn><mn>2</mn></mfrac></mrow></math>";

        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? sourceIdentity = null;
        Word.Bookmark? adoptedIdentity = null;
        Word.Range? sourceIdentityRange = null;
        Word.Range? sourceRange = null;
        Word.Range? tail = null;
        Word.OMaths? maths = null;
        Word.OMath? pastedMath = null;
        Word.Range? pastedRange = null;
        Word.Range? adoptedIdentityRange = null;
        Word.Range? adoptedRange = null;
        Word.Font? adoptedFont = null;
        try
        {
            document =
                application.Documents.Add(
                    Visible: false);
            document.Activate();
            var service =
                new WordFormulaService(
                    application);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            selection = application.Selection;
            selection.SetRange(0, 0);

            var createSession =
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    createSession,
                    mathMl);

            bookmarks = document.Bookmarks;
            var sourceIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    created.FormulaId);
            sourceIdentity =
                bookmarks[sourceIdentityName];
            sourceIdentityRange =
                sourceIdentity.Range.Duplicate;
            var sourceHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    sourceIdentityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Font-size adoption source cannot be resolved.");
            sourceRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    sourceHost.Range);
            sourceRange.Select();

            var snapshot =
                service.CaptureSelectedFormulaForCopy()
                ?? throw new InvalidDataException(
                    "Font-size adoption source was not captured for copy.");
            selection.Copy();
            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            var pasteStart =
                selection.Range.Start;
            service.ArmHostCorePaste(
                snapshot);
            selection.Paste();
            System.Threading.Thread.Sleep(100);
            AssertEqual(
                WordFormulaService.PastedFormulaRepairResult.Repaired,
                service.RepairPastedFormula(
                    snapshot),
                "Font-size adoption paste validation did not complete.");

            tail =
                document.Range(
                    pasteStart,
                    document.Content.End);
            maths = tail.OMaths;
            AssertEqual(
                1,
                maths.Count,
                "Font-size adoption paste did not create exactly one OMath.");
            pastedMath = maths[1];
            pastedRange =
                pastedMath.Range.Duplicate;
            pastedRange.Select();

            var opened =
                service.ReadSelection();
            AssertTrue(
                opened.Metadata is not null
                && opened.Metadata.Numbered,
                "Font-size adoption ReadSelection lost numbered state.");
            AssertTrue(
                !string.IsNullOrWhiteSpace(
                    opened.FormulaId)
                && !string.Equals(
                    opened.FormulaId,
                    created.FormulaId,
                    StringComparison.OrdinalIgnoreCase),
                "Font-size adoption did not allocate an independent session FormulaId.");
            var adoptedIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    opened.FormulaId!);
            var adoptedRefName =
                WordNativeOmmlNumbering.ReferenceBookmarkName(
                    opened.FormulaId!);

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertTrue(
                !bookmarks.Exists(
                    adoptedIdentityName)
                && !bookmarks.Exists(
                    adoptedRefName),
                "Opening an anonymous OMath for font-size editing persisted identity too early.");

            var applied =
                service.SetSelectedFormulaFontSize(
                    16);
            AssertNear(
                16f,
                applied,
                0.01f,
                "Native OMML font-size edit did not report 16 pt.");

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(
                    adoptedIdentityName),
                "Font-size edit did not persist VTOMML ownership.");
            AssertTrue(
                bookmarks.Exists(
                    adoptedRefName),
                "Font-size edit did not persist the native VTEqNum target.");

            adoptedIdentity =
                bookmarks[adoptedIdentityName];
            adoptedIdentityRange =
                adoptedIdentity.Range.Duplicate;
            var adopted =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    adoptedIdentityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Font-size adopted OMath cannot be resolved from VTOMML.");
            adopted.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    adopted);
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                adopted.Numbering.ContainerKind,
                "Font-size edit changed native #(SEQ) numbering topology.");
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(
                    mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    WordFormulaHostSemanticReader.Read(
                        document,
                        adopted).MathMl
                    ?? throw new InvalidDataException(
                        "Font-size adopted OMath returned no MathML.")),
                "Font-size edit changed native OMML semantics.");

            adoptedRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    adopted.Range);
            adoptedFont =
                adoptedRange.Font;
            AssertTrue(
                adoptedFont.Size > 0
                && adoptedFont.Size < 256,
                "Font-size adopted OMath returned an undefined Word font size.");
            AssertNear(
                16f,
                adoptedFont.Size,
                0.35f,
                "Font-size adopted OMath did not retain 16 pt.");
            AssertEqual(
                0,
                document.Tables.Count,
                "Font-size edit introduced a layout table for native OMML.");
            AssertEqual(
                0,
                document.Shapes.Count,
                "Font-size edit introduced a Shape for native OMML.");

            Console.WriteLine(
                "[host-core minimal] anonymous native OMML font-size edit adopted ownership and preserved native numbering.");
        }
        finally
        {
            Release(adoptedFont);
            Release(adoptedRange);
            Release(adoptedIdentityRange);
            Release(adoptedIdentity);
            Release(pastedRange);
            Release(pastedMath);
            Release(maths);
            Release(tail);
            Release(sourceRange);
            Release(sourceIdentityRange);
            Release(sourceIdentity);
            Release(bookmarks);
            Release(selection);
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
            try { mainDocument.Activate(); } catch { }
        }
    }

    private static void RunHostCoreNativeNumberToggleSmoke(
        Word.Application application,
        Word.Document mainDocument)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>t</mi><mo>=</mo><msup><mi>x</mi><mn>2</mn></msup></mrow></math>";

        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? identity = null;
        Word.Range? identityRange = null;
        Word.Range? hostRange = null;
        try
        {
            document =
                application.Documents.Add(
                    Visible: false);
            document.Activate();
            var service =
                new WordFormulaService(
                    application);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            selection = application.Selection;
            selection.SetRange(0, 0);

            var createSession =
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    createSession,
                    mathMl);
            var identityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    created.FormulaId);
            var refName =
                WordNativeOmmlNumbering.ReferenceBookmarkName(
                    created.FormulaId);

            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(
                    identityName)
                && bookmarks.Exists(
                    refName),
                "Number-toggle fixture did not start as owned native numbered OMML.");
            identity =
                bookmarks[identityName];
            identityRange =
                identity.Range.Duplicate;
            var numbered =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    identityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Number-toggle source cannot be resolved.");
            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    numbered.Range);
            hostRange.Select();
            var selected =
                service.ReadSelection();

            var offSession =
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: false,
                    FormulaOleContract.WordOmmlMode);
            offSession.Mode = "edit";
            offSession.FormulaId =
                selected.FormulaId!;
            offSession.SourceDocumentId =
                selected.DocumentId;
            offSession.SourceObjectId =
                selected.ObjectId;
            offSession.OriginalMetadata =
                selected.Metadata;
            _ = service.ReplaceOmml(
                offSession,
                mathMl);

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(
                    identityName),
                "Turning numbering off removed durable VTOMML ownership.");
            AssertTrue(
                !bookmarks.Exists(
                    refName),
                "Turning numbering off left VTEqNum alive.");

            Release(identityRange);
            identityRange = null;
            Release(identity);
            identity =
                bookmarks[identityName];
            identityRange =
                identity.Range.Duplicate;
            var unnumbered =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    identityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Unnumbered OMML cannot be resolved after number toggle.");
            unnumbered.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    unnumbered);
            AssertTrue(
                !unnumbered.Numbering.Numbered
                && unnumbered.Numbering.ContainerKind ==
                    WordFormulaNumberingContainerKind.None,
                "Turning numbering off did not produce an ordinary unnumbered OMML host.");
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(
                    mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    WordFormulaHostSemanticReader.Read(
                        document,
                        unnumbered).MathMl
                    ?? throw new InvalidDataException(
                        "Unnumbered toggle result returned no MathML.")),
                "Turning numbering off changed formula semantics.");

            Release(hostRange);
            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    unnumbered.Range);
            hostRange.Select();
            selected =
                service.ReadSelection();

            var onSession =
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            onSession.Mode = "edit";
            onSession.FormulaId =
                selected.FormulaId!;
            onSession.SourceDocumentId =
                selected.DocumentId;
            onSession.SourceObjectId =
                selected.ObjectId;
            onSession.OriginalMetadata =
                selected.Metadata;
            _ = service.ReplaceOmml(
                onSession,
                mathMl);

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(
                    identityName)
                && bookmarks.Exists(
                    refName),
                "Turning numbering back on did not restore VTOMML + VTEqNum.");

            Release(identityRange);
            identityRange = null;
            Release(identity);
            identity =
                bookmarks[identityName];
            identityRange =
                identity.Range.Duplicate;
            var renumbered =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    identityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Renumbered OMML cannot be resolved.");
            renumbered.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    renumbered);
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                renumbered.Numbering.ContainerKind,
                "Turning numbering on did not restore native #(SEQ).");
            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    renumbered);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(
                    mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    payload.MathMl
                    ?? throw new InvalidDataException(
                        "Renumbered OMML returned no MathML.")),
                "Number toggle round-trip changed formula semantics.");
            AssertTrue(
                payload.Latex.IndexOf(
                    "#",
                    StringComparison.Ordinal) < 0
                && payload.Latex.IndexOf(
                    "VisualTeXEquation",
                    StringComparison.OrdinalIgnoreCase) < 0,
                "Number toggle round-trip exposed native numbering to the editor.");
            AssertEqual(
                0,
                document.Tables.Count,
                "Native OMML number toggle introduced a layout table.");
            AssertEqual(
                0,
                document.Shapes.Count,
                "Native OMML number toggle introduced a Shape.");

            Console.WriteLine(
                "[host-core minimal] native OMML number off/on round-trip preserved identity, semantics and native #(SEQ).");
        }
        finally
        {
            Release(hostRange);
            Release(identityRange);
            Release(identity);
            Release(bookmarks);
            Release(selection);
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
            try { mainDocument.Activate(); } catch { }
        }
    }

    private static void RunHostCoreNativeFormulaToLatexSmoke(
        Word.Application application,
        Word.Document mainDocument)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>z</mi><mo>=</mo><mn>5</mn></mrow></math>";

        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? identity = null;
        Word.Range? identityRange = null;
        Word.Range? formulaRange = null;
        Word.Range? content = null;
        try
        {
            document =
                application.Documents.Add(
                    Visible: false);
            document.Activate();
            var service =
                new WordFormulaService(
                    application);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            selection = application.Selection;
            selection.SetRange(0, 0);

            var createSession =
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    createSession,
                    mathMl);
            var identityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    created.FormulaId);
            var refName =
                WordNativeOmmlNumbering.ReferenceBookmarkName(
                    created.FormulaId);

            var targets =
                service.GetCanonicalEquationReferenceTargets(
                    document);
            var target =
                targets.Single(item =>
                    string.Equals(
                        item.FormulaId,
                        created.FormulaId,
                        StringComparison.OrdinalIgnoreCase));
            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            selection.TypeText(
                "formula-ref ");
            service.InsertEquationReference(
                document,
                selection,
                target,
                EquationReferenceStyle.Parenthesized,
                Word.WdColor.wdColorAutomatic);
            AssertTrue(
                ContainsHostCoreReferenceField(
                    document,
                    refName),
                "Formula-to-LaTeX fixture did not create a real REF field.");

            bookmarks = document.Bookmarks;
            identity =
                bookmarks[identityName];
            identityRange =
                identity.Range.Duplicate;
            var host =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    identityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Formula-to-LaTeX source cannot be resolved.");
            formulaRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            formulaRange.Select();

            var result =
                service.ConvertFormulaObjectsToLatex(
                    wholeDocument: false,
                    FormulaOleContract.WordOmmlMode);
            AssertEqual(
                1,
                result.FormulaCount,
                "Formula-to-LaTeX did not convert exactly one native OMML host.");
            AssertEqual(
                0,
                document.OMaths.Count,
                "Formula-to-LaTeX left the native OMath alive.");

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertTrue(
                !bookmarks.Exists(
                    identityName),
                "Formula-to-LaTeX left VTOMML ownership after deleting the formula.");
            AssertTrue(
                !bookmarks.Exists(
                    refName),
                "Formula-to-LaTeX left VTEqNum after deleting the numbered formula.");
            AssertTrue(
                ContainsHostCoreReferenceField(
                    document,
                    refName),
                "Formula-to-LaTeX froze or deleted the existing Word REF field.");

            var broken =
                ReadHostCoreReferenceResult(
                    document,
                    refName,
                    updateNestedRef: true);
            AssertTrue(
                IsMissingReferenceResult(
                    broken),
                "Formula-to-LaTeX did not let Word naturally report a missing REF target.");

            content =
                document.Content;
            var text =
                content.Text
                ?? string.Empty;
            AssertTrue(
                text.IndexOf(
                    "$$",
                    StringComparison.Ordinal) >= 0
                && text.IndexOf(
                    "z",
                    StringComparison.OrdinalIgnoreCase) >= 0,
                "Formula-to-LaTeX did not restore visible display LaTeX source.");
            AssertEqual(
                0,
                document.Tables.Count,
                "Formula-to-LaTeX left a layout table.");
            AssertEqual(
                0,
                document.Shapes.Count,
                "Formula-to-LaTeX left a Shape.");

            Console.WriteLine(
                "[host-core minimal] native numbered OMML -> LaTeX preserved live REF and Word naturally reported missing target.");
        }
        finally
        {
            Release(content);
            Release(formulaRange);
            Release(identityRange);
            Release(identity);
            Release(bookmarks);
            Release(selection);
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
            try { mainDocument.Activate(); } catch { }
        }
    }

    private static void RunHostCoreNativeLatexRedrawSmoke(
        Word.Application application,
        Word.Document mainDocument)
    {
        const string inlineMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></math>";
        const string displayMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>y</mi><mo>=</mo><mn>2</mn></mrow></math>";

        Word.Document? document = null;
        try
        {
            document =
                application.Documents.Add(
                    Visible: false);
            document.Activate();
            document.Content.Text =
                "before $x+1$ after\r$$y=2$$\rtail";

            var service =
                new WordFormulaService(
                    application);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            var plan =
                service.CaptureLatexRedrawPlan(
                    wholeDocument: true);
            AssertEqual(
                2,
                plan.Targets.Count,
                "Native LaTeX redraw did not capture inline + display sources.");
            plan.NumberDisplayFormulas = true;

            var prepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal);
            foreach (var target in plan.Targets)
            {
                var mathMl =
                    string.Equals(
                        target.DisplayMode,
                        "block",
                        StringComparison.OrdinalIgnoreCase)
                        ? displayMathMl
                        : inlineMathMl;
                prepared[target.Id] =
                    new PreparedWordBulkFormula
                    {
                        Run = new WordBulkRun
                        {
                            Id = target.Id,
                            IsFormula = true,
                            Latex = target.Latex,
                            DisplayMode = target.DisplayMode,
                        },
                        Session =
                            CreateOmmlMathTypeAcceptanceSession(
                                mathMl,
                                target.DisplayMode,
                                numbered: false,
                                FormulaOleContract.WordOmmlMode),
                        MathMl = mathMl,
                    };
            }

            var result =
                service.ApplyLatexRedrawPlan(
                    plan,
                    prepared);
            AssertEqual(
                2,
                result.FormulaCount,
                "Native LaTeX redraw did not materialize both formulas.");

            var index =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            AssertEqual(
                2,
                index.Omml.Count,
                "Native LaTeX redraw did not leave exactly two OMML hosts.");
            AssertEqual(
                0,
                index.VisualTeX.Count,
                "Native LaTeX redraw unexpectedly created a VisualTeX OLE.");

            var inlineHosts =
                index.Omml
                    .Where(host =>
                        string.Equals(
                            host.DisplayMode,
                            "inline",
                            StringComparison.OrdinalIgnoreCase))
                    .ToArray();
            var displayHosts =
                index.Omml
                    .Where(host =>
                        string.Equals(
                            host.DisplayMode,
                            "block",
                            StringComparison.OrdinalIgnoreCase))
                    .ToArray();
            AssertEqual(
                1,
                inlineHosts.Length,
                "Native LaTeX redraw changed the inline/display split.");
            AssertEqual(
                1,
                displayHosts.Length,
                "Native LaTeX redraw changed the inline/display split.");

            inlineHosts[0].Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    inlineHosts[0]);
            displayHosts[0].Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    displayHosts[0]);
            AssertTrue(
                !inlineHosts[0].Numbering.Numbered,
                "Inline LaTeX redraw was numbered.");
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                displayHosts[0].Numbering.ContainerKind,
                "Display LaTeX redraw did not use native #(SEQ) numbering.");

            var visibleText =
                document.Content.Text
                ?? string.Empty;
            AssertTrue(
                visibleText.IndexOf(
                    "before ",
                    StringComparison.Ordinal) >= 0
                && visibleText.IndexOf(
                    " after",
                    StringComparison.Ordinal) >= 0
                && visibleText.IndexOf(
                    "tail",
                    StringComparison.Ordinal) >= 0,
                "Native LaTeX redraw changed surrounding user text.");
            AssertTrue(
                visibleText.IndexOf(
                    '$') < 0,
                "Native LaTeX redraw left source delimiters in Word text.");
            AssertEqual(
                0,
                document.Tables.Count,
                "Native LaTeX redraw created a legacy layout table.");
            AssertEqual(
                0,
                document.Shapes.Count,
                "Native LaTeX redraw created a Shape.");
            AssertNoHostCoreSentinelArtifacts(
                document,
                "after native LaTeX redraw");

            Console.WriteLine(
                "[host-core minimal] native LaTeX redraw used OMML host core, preserved inline/display semantics and native display numbering.");
        }
        finally
        {
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
            try { mainDocument.Activate(); } catch { }
        }
    }

    private static void RunHostCoreNativeHeadingNumberingSmoke(
        Word.Application application,
        Word.Document mainDocument)
    {
        const string firstMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>a</mi><mo>=</mo><mn>1</mn></mrow></math>";
        const string secondMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>b</mi><mo>=</mo><mn>2</mn></mrow></math>";

        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.ListTemplate? listTemplate = null;
        Word.ListLevel? listLevel = null;
        Word.Range? headingRange = null;
        Word.ListFormat? listFormat = null;
        try
        {
            document =
                application.Documents.Add(
                    Visible: false);
            document.Activate();
            var service =
                new WordFormulaService(
                    application);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);

            listTemplate =
                document.ListTemplates.Add(
                    OutlineNumbered: false,
                    Name:
                        "VisualTeXHostCoreHeading"
                        + Guid.NewGuid().ToString("N"));
            listLevel =
                listTemplate.ListLevels[1];
            listLevel.NumberStyle =
                Word.WdListNumberStyle.wdListNumberStyleArabic;
            listLevel.NumberFormat = "%1";
            listLevel.StartAt = 1;

            selection = application.Selection;
            selection.SetRange(0, 0);

            void AppendHeading(
                string text,
                bool continuePreviousList)
            {
                Release(listFormat);
                listFormat = null;
                Release(headingRange);
                headingRange = null;

                var start =
                    selection.Start;
                selection.TypeText(text);
                selection.TypeParagraph();
                headingRange =
                    document.Range(
                        start,
                        selection.Start);
                object headingStyle =
                    Word.WdBuiltinStyle.wdStyleHeading1;
                headingRange.set_Style(
                    ref headingStyle);
                listFormat =
                    headingRange.ListFormat;
                listFormat.ApplyListTemplateWithLevel(
                    listTemplate,
                    ContinuePreviousList:
                        continuePreviousList,
                    ApplyTo:
                        Word.WdListApplyTo.wdListApplyToWholeList,
                    DefaultListBehavior:
                        Word.WdDefaultListBehavior.wdWord10ListBehavior,
                    ApplyLevel: 1);
                object normalStyle =
                    Word.WdBuiltinStyle.wdStyleNormal;
                selection.set_Style(
                    ref normalStyle);
            }

            AppendHeading(
                "Chapter One",
                continuePreviousList: false);
            var first =
                service.InsertOmml(
                    CreateOmmlMathTypeAcceptanceSession(
                        firstMathMl,
                        "block",
                        numbered: true,
                        FormulaOleContract.WordOmmlMode),
                    firstMathMl);

            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            AppendHeading(
                "Chapter Two",
                continuePreviousList: true);
            var second =
                service.InsertOmml(
                    CreateOmmlMathTypeAcceptanceSession(
                        secondMathMl,
                        "block",
                        numbered: true,
                        FormulaOleContract.WordOmmlMode),
                    secondMathMl);

            var targets =
                service.GetCanonicalEquationReferenceTargets(
                    document)
                    .OrderBy(target => target.Position)
                    .ToArray();
            AssertEqual(
                2,
                targets.Length,
                "Heading-numbering smoke did not expose two native reference targets.");
            AssertEqual(
                first.FormulaId,
                targets[0].FormulaId,
                "First heading-numbered OMML changed FormulaId.");
            AssertEqual(
                second.FormulaId,
                targets[1].FormulaId,
                "Second heading-numbered OMML changed FormulaId.");
            AssertEqual(
                "1.1",
                targets[0].NumberText,
                "First native heading equation did not resolve through STYLEREF + SEQ.");
            AssertEqual(
                "2.1",
                targets[1].NumberText,
                "Second native heading equation did not restart SEQ at Heading 1.");

            var index =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            AssertEqual(
                2,
                index.Omml.Count,
                "Heading numbering did not leave exactly two OMML hosts.");
            foreach (var host in index.Omml)
            {
                host.Numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        host);
                AssertEqual(
                    WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                    host.Numbering.ContainerKind,
                    "Heading-numbered OMML is not one native #(SEQ) host.");

                Word.Range? hostRange = null;
                Word.Fields? fields = null;
                Word.Field? field = null;
                Word.Range? code = null;
                try
                {
                    hostRange =
                        WordFormulaHostSemanticReader.CreateRange(
                            document,
                            host.Range);
                    fields = hostRange.Fields;
                    var styleRefCount = 0;
                    var sequenceCount = 0;
                    for (var fieldIndex = 1;
                         fieldIndex <= fields.Count;
                         fieldIndex++)
                    {
                        Release(code); code = null;
                        Release(field); field = fields[fieldIndex];
                        code = field.Code.Duplicate;
                        var normalized =
                            (code.Text ?? string.Empty)
                                .Trim();
                        if (normalized.StartsWith(
                                "STYLEREF 1",
                                StringComparison.OrdinalIgnoreCase))
                            styleRefCount++;
                        if (normalized.StartsWith(
                                "SEQ VisualTeXEquation",
                                StringComparison.OrdinalIgnoreCase)
                            && normalized.IndexOf(
                                "\\s 1",
                                StringComparison.OrdinalIgnoreCase) >= 0)
                            sequenceCount++;
                    }
                    AssertEqual(
                        1,
                        styleRefCount,
                        "Heading-numbered OMML does not contain exactly one native STYLEREF 1 field.");
                    AssertEqual(
                        1,
                        sequenceCount,
                        "Heading-numbered OMML does not contain exactly one SEQ VisualTeXEquation \\s 1 field.");
                }
                finally
                {
                    Release(code);
                    Release(field);
                    Release(fields);
                    Release(hostRange);
                }
            }

            AssertEqual(
                0,
                document.Tables.Count,
                "Heading-numbered native OMML created a legacy layout table.");
            AssertEqual(
                0,
                document.Shapes.Count,
                "Heading-numbered native OMML created a Shape.");

            Console.WriteLine(
                "[host-core minimal] real Heading 1 numbering drove native STYLEREF + SEQ restart semantics (1.1, 2.1).");
        }
        finally
        {
            Release(listFormat);
            Release(headingRange);
            Release(listLevel);
            Release(listTemplate);
            Release(selection);
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
            try { mainDocument.Activate(); } catch { }
        }
    }

    private static void RunHostCoreAdjacentNativeDeleteSmoke(
        Word.Application application,
        Word.Document mainDocument)
    {
        const string firstMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></math>";
        const string secondMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>y</mi><mo>+</mo><mn>2</mn></mrow></math>";

        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? firstIdentity = null;
        Word.Range? firstIdentityRange = null;
        try
        {
            document =
                application.Documents.Add(
                    Visible: false);
            document.Activate();
            document.Content.Text = "LR";
            var service =
                new WordFormulaService(
                    application);

            selection = application.Selection;
            selection.SetRange(1, 1);
            var first =
                service.InsertOmml(
                    CreateOmmlMathTypeAcceptanceSession(
                        firstMathMl,
                        "inline",
                        numbered: false,
                        FormulaOleContract.WordOmmlMode),
                    firstMathMl);

            bookmarks = document.Bookmarks;
            var firstIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    first.FormulaId);
            firstIdentity =
                bookmarks[firstIdentityName];
            firstIdentityRange =
                firstIdentity.Range.Duplicate;
            var firstHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    firstIdentityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "First adjacent OMML cannot be resolved.");

            selection.SetRange(
                firstHost.Range.End,
                firstHost.Range.End);
            var second =
                service.InsertOmml(
                    CreateOmmlMathTypeAcceptanceSession(
                        secondMathMl,
                        "inline",
                        numbered: false,
                        FormulaOleContract.WordOmmlMode),
                    secondMathMl);

            var beforeDelete =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            AssertEqual(
                2,
                beforeDelete.Omml.Count,
                "Two adjacent inline OMML inserts collapsed into one Word OMath.");
            var ordered =
                beforeDelete.Omml
                    .OrderBy(host => host.Range.Start)
                    .ToArray();
            AssertEqual(
                ordered[0].Range.End,
                ordered[1].Range.Start,
                "Adjacent inline OMML hosts are not truly adjacent.");
            AssertEqual(
                first.FormulaId,
                ordered[0].FormulaId
                ?? string.Empty,
                "First adjacent OMML identity changed.");
            AssertEqual(
                second.FormulaId,
                ordered[1].FormulaId
                ?? string.Empty,
                "Second adjacent OMML identity changed.");

            WordFormulaHostMutationKernel.Delete(
                application,
                document,
                ordered[0]);

            var afterDelete =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            AssertEqual(
                1,
                afterDelete.Omml.Count,
                "Deleting the first adjacent OMML damaged the surviving host count.");
            AssertEqual(
                second.FormulaId,
                afterDelete.Omml[0].FormulaId
                ?? string.Empty,
                "Deleting the first adjacent OMML changed the second FormulaId.");
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(
                    secondMathMl),
                MathTypeMtefCodec.SemanticSignature(
                    WordFormulaHostSemanticReader.Read(
                        document,
                        afterDelete.Omml[0]).MathMl
                    ?? throw new InvalidDataException(
                        "Surviving adjacent OMML returned no MathML.")),
                "Deleting one adjacent OMML changed the survivor semantics.");

            var contentText =
                document.Content.Text
                ?? string.Empty;
            AssertTrue(
                contentText.StartsWith(
                    "L",
                    StringComparison.Ordinal)
                && contentText.TrimEnd('\r')
                    .EndsWith(
                        "R",
                        StringComparison.Ordinal),
                "Adjacent OMML deletion changed surrounding user text.");
            AssertNoHostCoreSentinelArtifacts(
                document,
                "after adjacent native OMML deletion");

            Console.WriteLine(
                "[host-core minimal] adjacent native OMML hosts stayed distinct and exact deletion preserved the neighbor + user text.");
        }
        finally
        {
            Release(firstIdentityRange);
            Release(firstIdentity);
            Release(bookmarks);
            Release(selection);
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
            try { mainDocument.Activate(); } catch { }
        }
    }

    private static void RunHostCoreNumberedFormatRoundTripSmoke(
        Word.Application application,
        Word.Document mainDocument,
        string pngPath,
        string emfPath)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>c</mi><mo>=</mo><mi>a</mi><mo>+</mo><mi>b</mi></mrow></math>";

        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? identity = null;
        Word.Range? identityRange = null;
        Word.Range? sourceRange = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? visualShape = null;
        Word.Range? visualRange = null;
        try
        {
            document =
                application.Documents.Add(
                    Visible: false);
            document.Activate();
            var service =
                new WordFormulaService(
                    application);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            selection = application.Selection;
            selection.SetRange(0, 0);

            var createSession =
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    createSession,
                    mathMl);
            var identityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    created.FormulaId);
            var refName =
                WordNativeOmmlNumbering.ReferenceBookmarkName(
                    created.FormulaId);

            var targets =
                service.GetCanonicalEquationReferenceTargets(
                    document);
            var target =
                targets.Single(item =>
                    string.Equals(
                        item.FormulaId,
                        created.FormulaId,
                        StringComparison.OrdinalIgnoreCase));
            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            selection.TypeText(
                "roundtrip-ref ");
            service.InsertEquationReference(
                document,
                selection,
                target,
                EquationReferenceStyle.Parenthesized,
                Word.WdColor.wdColorAutomatic);
            AssertTrue(
                ContainsHostCoreReferenceField(
                    document,
                    refName),
                "Numbered format round-trip fixture did not create direct REF.");

            bookmarks = document.Bookmarks;
            identity =
                bookmarks[identityName];
            identityRange =
                identity.Range.Duplicate;
            var source =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    identityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Numbered format round-trip OMML source cannot be resolved.");
            sourceRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    source.Range);
            sourceRange.Select();

            var toVisualPlan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: false,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.NativeOleMode);
            AssertEqual(
                1,
                toVisualPlan.Targets.Count,
                "Numbered OMML→VisualTeX did not capture exactly one target.");
            var toVisualTarget =
                toVisualPlan.Targets.Single();
            AssertTrue(
                toVisualTarget.Numbered,
                "Numbered OMML→VisualTeX capture lost numbered state.");
            AssertEqual(
                created.FormulaId,
                toVisualTarget.SourceFormulaId,
                "Numbered OMML→VisualTeX capture changed FormulaId.");
            var toVisualMathMl =
                toVisualTarget.SourceMathMl
                ?? mathMl;
            var toVisualPrepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal)
                {
                    [toVisualTarget.Id] =
                        new PreparedWordBulkFormula
                        {
                            Run = new WordBulkRun
                            {
                                Id = toVisualTarget.Id,
                                IsFormula = true,
                                Latex = toVisualTarget.Latex,
                                DisplayMode =
                                    toVisualTarget.DisplayMode,
                            },
                            Session =
                                CreateSimpleFormatTargetSession(
                                    toVisualTarget,
                                    FormulaOleContract.NativeOleMode,
                                    toVisualMathMl),
                            MathMl =
                                toVisualMathMl,
                            PngPath = pngPath,
                            EmfPath = emfPath,
                        },
                };

            Release(sourceRange);
            sourceRange = null;
            var toVisualResult =
                service.ApplyFormulaFormatConversionPlan(
                    toVisualPlan,
                    toVisualPrepared);
            AssertEqual(
                1,
                toVisualResult.FormulaCount,
                "Numbered OMML→VisualTeX did not convert exactly one formula.");
            AssertEqual(
                0,
                toVisualResult.FailedFormulaCount,
                "Numbered OMML→VisualTeX reported a failure.");
            AssertEqual(
                0,
                document.OMaths.Count,
                "Numbered OMML→VisualTeX left the source OMath alive.");
            AssertEqual(
                1,
                CountVisualTeXNativeOleShapes(
                    document),
                "Numbered OMML→VisualTeX did not create one VisualTeX OLE.");

            shapes =
                document.InlineShapes;
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate =
                        shapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(
                            candidate))
                        continue;
                    var metadata =
                        WordFormulaMetadataReader.TryReadEmbeddedNativeOle(
                            candidate);
                    if (metadata is null
                        || !string.Equals(
                            metadata.FormulaId,
                            created.FormulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    visualShape = candidate;
                    candidate = null;
                    break;
                }
                finally
                {
                    Release(candidate);
                }
            }
            AssertTrue(
                visualShape is not null,
                "Numbered OMML→VisualTeX target with the original FormulaId is missing.");
            visualRange =
                visualShape!.Range.Duplicate;
            var visualHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    visualRange,
                    WordFormulaHostKind.VisualTeX)
                ?? throw new InvalidDataException(
                    "Converted numbered VisualTeX host cannot be resolved.");
            visualHost.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    visualHost);
            AssertTrue(
                visualHost.Numbering.Numbered,
                "OMML→VisualTeX lost numbered state.");
            AssertTrue(
                visualHost.Numbering.ContainerKind is
                    WordFormulaNumberingContainerKind.CanonicalBodyTable
                    or WordFormulaNumberingContainerKind.CanonicalUserTableCell,
                "Converted VisualTeX OLE is not in its canonical external numbering layout.");

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(
                    refName),
                "OMML→VisualTeX changed or removed VTEqNum target identity.");
            var visualRefText =
                ReadHostCoreReferenceResult(
                    document,
                    refName,
                    updateNestedRef: true);
            AssertTrue(
                visualRefText.IndexOf(
                    target.NumberText,
                    StringComparison.Ordinal) >= 0,
                "Direct REF stopped resolving after OMML→VisualTeX.");

            visualRange.Select();
            var toOmmlPlan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: false,
                    FormulaOleContract.NativeOleMode,
                    FormulaOleContract.WordOmmlMode);
            AssertEqual(
                1,
                toOmmlPlan.Targets.Count,
                "Numbered VisualTeX→OMML did not capture exactly one target.");
            var toOmmlTarget =
                toOmmlPlan.Targets.Single();
            AssertTrue(
                toOmmlTarget.Numbered,
                "Numbered VisualTeX→OMML capture lost numbered state.");
            AssertEqual(
                created.FormulaId,
                toOmmlTarget.SourceFormulaId,
                "Numbered VisualTeX→OMML capture changed FormulaId.");
            var toOmmlPrepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal)
                {
                    [toOmmlTarget.Id] =
                        new PreparedWordBulkFormula
                        {
                            Run = new WordBulkRun
                            {
                                Id = toOmmlTarget.Id,
                                IsFormula = true,
                                Latex = toOmmlTarget.Latex,
                                DisplayMode =
                                    toOmmlTarget.DisplayMode,
                            },
                            Session =
                                CreateSimpleFormatTargetSession(
                                    toOmmlTarget,
                                    FormulaOleContract.WordOmmlMode,
                                    mathMl),
                            MathMl = mathMl,
                        },
                };

            Release(visualRange);
            visualRange = null;
            Release(visualShape);
            visualShape = null;
            Release(shapes);
            shapes = null;

            var toOmmlResult =
                service.ApplyFormulaFormatConversionPlan(
                    toOmmlPlan,
                    toOmmlPrepared);
            AssertEqual(
                1,
                toOmmlResult.FormulaCount,
                "Numbered VisualTeX→OMML did not convert exactly one formula.");
            AssertEqual(
                0,
                toOmmlResult.FailedFormulaCount,
                "Numbered VisualTeX→OMML reported a failure.");
            AssertEqual(
                0,
                CountVisualTeXNativeOleShapes(
                    document),
                "VisualTeX→OMML left the source OLE alive.");
            AssertEqual(
                1,
                document.OMaths.Count,
                "VisualTeX→OMML did not create exactly one native OMath.");

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(
                    identityName)
                && bookmarks.Exists(
                    refName),
                "VisualTeX→OMML did not restore VTOMML + VTEqNum using the same FormulaId.");
            Release(identityRange);
            identityRange = null;
            Release(identity);
            identity =
                bookmarks[identityName];
            identityRange =
                identity.Range.Duplicate;
            var finalHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    identityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Round-tripped OMML cannot be resolved from VTOMML.");
            finalHost.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    finalHost);
            AssertEqual(
                created.FormulaId,
                finalHost.FormulaId
                ?? string.Empty,
                "Format round-trip changed the OMML FormulaId.");
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                finalHost.Numbering.ContainerKind,
                "Format round-trip did not return to Word-native #(SEQ).");
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(
                    mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    WordFormulaHostSemanticReader.Read(
                        document,
                        finalHost).MathMl
                    ?? throw new InvalidDataException(
                        "Round-tripped OMML returned no MathML.")),
                "Numbered OMML↔VisualTeX round-trip changed formula semantics.");

            var finalRefText =
                ReadHostCoreReferenceResult(
                    document,
                    refName,
                    updateNestedRef: true);
            AssertTrue(
                finalRefText.IndexOf(
                    target.NumberText,
                    StringComparison.Ordinal) >= 0,
                "Direct REF stopped resolving after VisualTeX→OMML.");
            AssertEqual(
                0,
                document.Tables.Count,
                "Round-tripped native OMML retained a VisualTeX layout table.");
            AssertEqual(
                0,
                document.Shapes.Count,
                "Round-tripped native OMML introduced a Shape.");

            Console.WriteLine(
                "[host-core minimal] numbered OMML↔VisualTeX format round-trip preserved FormulaId, live REF and native final #(SEQ).");
        }
        finally
        {
            Release(visualRange);
            Release(visualShape);
            Release(shapes);
            Release(sourceRange);
            Release(identityRange);
            Release(identity);
            Release(bookmarks);
            Release(selection);
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
            try { mainDocument.Activate(); } catch { }
        }
    }

    private static void RunHostCoreNativeSaveReopenSmoke(
        Word.Application application,
        Word.Document mainDocument,
        string artifactRoot)
    {
        const string firstMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>a</mi><mo>=</mo><mn>1</mn></mrow></math>";
        const string secondMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>b</mi><mo>=</mo><mn>2</mn></mrow></math>";

        var path =
            Path.Combine(
                artifactRoot,
                "host-core-native-save-reopen.docx");

        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? identity = null;
        Word.Range? identityRange = null;
        Word.Fields? fields = null;
        Word.Field? referenceField = null;
        try
        {
            document =
                application.Documents.Add(
                    Visible: false);
            document.Activate();
            document.SaveAs2(
                path,
                Word.WdSaveFormat.wdFormatXMLDocument);
            var service =
                new WordFormulaService(
                    application);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            selection = application.Selection;
            selection.SetRange(0, 0);

            var firstSession =
                CreateOmmlMathTypeAcceptanceSession(
                    firstMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var first =
                service.InsertOmml(
                    firstSession,
                    firstMathMl);

            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            var secondSession =
                CreateOmmlMathTypeAcceptanceSession(
                    secondMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var second =
                service.InsertOmml(
                    secondSession,
                    secondMathMl);

            var targets =
                service.GetCanonicalEquationReferenceTargets(
                    document);
            var firstTarget =
                targets.Single(item =>
                    string.Equals(
                        item.FormulaId,
                        first.FormulaId,
                        StringComparison.OrdinalIgnoreCase));

            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            selection.TypeText(
                "Native reference: ");
            service.InsertEquationReference(
                document,
                selection,
                firstTarget,
                EquationReferenceStyle.Parenthesized,
                Word.WdColor.wdColorAutomatic);

            var firstRefName =
                WordNativeOmmlNumbering.ReferenceBookmarkName(
                    first.FormulaId);
            var firstIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    first.FormulaId);
            var secondIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    second.FormulaId);

            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(firstRefName)
                && bookmarks.Exists(firstIdentityName)
                && bookmarks.Exists(secondIdentityName),
                "Native save/reopen fixture did not create durable identities/reference target.");

            document.Save();
            document.Close(
                Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            document =
                application.Documents.Open(
                    path,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);
            document.Activate();
            service =
                new WordFormulaService(
                    application);

            // No reconcile/on-open repair is called here. Persistence itself must
            // be sufficient to resolve the same native hosts.
            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(firstIdentityName)
                && bookmarks.Exists(secondIdentityName)
                && bookmarks.Exists(firstRefName),
                "Save/reopen lost native OMML identity or VTEqNum target.");

            foreach (var tuple in new[]
                     {
                         (first.FormulaId, firstIdentityName, firstMathMl),
                         (second.FormulaId, secondIdentityName, secondMathMl),
                     })
            {
                Release(identityRange);
                identityRange = null;
                Release(identity);
                identity =
                    bookmarks[tuple.Item2];
                identityRange =
                    identity.Range.Duplicate;

                var host =
                    WordFormulaHostResolver.ResolveLocal(
                        document,
                        identityRange,
                        WordFormulaHostKind.Omml)
                    ?? throw new InvalidDataException(
                        "Saved/reopened native OMML cannot be resolved from its exact identity.");
                AssertEqual(
                    tuple.FormulaId,
                    host.FormulaId ?? string.Empty,
                    "Save/reopen changed the native OMML FormulaId.");

                host.Numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        host);
                AssertEqual(
                    WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                    host.Numbering.ContainerKind,
                    "Saved/reopened OMML is no longer Word-native #(SEQ).");

                var payload =
                    WordFormulaHostSemanticReader.Read(
                        document,
                        host);
                AssertEqual(
                    MathTypeMtefCodec.SemanticSignature(tuple.Item3),
                    MathTypeMtefCodec.SemanticSignature(
                        payload.MathMl
                        ?? throw new InvalidDataException(
                            "Saved/reopened native OMML has no semantic MathML.")),
                    "Save/reopen changed native OMML semantic content.");
                AssertTrue(
                    payload.Latex.IndexOf(
                        "#",
                        StringComparison.Ordinal) < 0
                    && payload.Latex.IndexOf(
                        "VisualTeXEquation",
                        StringComparison.OrdinalIgnoreCase) < 0,
                    "Saved/reopened native number leaked into editor semantics.");
            }

            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? field = null;
                Word.Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if ((code.Text ?? string.Empty)
                        .TrimStart()
                        .StartsWith(
                            "REF ",
                            StringComparison.OrdinalIgnoreCase))
                        field.Update();
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }

            var reopenedRefText =
                ReadHostCoreReferenceResult(
                    document,
                    firstRefName,
                    updateNestedRef: true);
            AssertTrue(
                reopenedRefText.IndexOf(
                    firstTarget.NumberText,
                    StringComparison.Ordinal) >= 0,
                "Saved/reopened native REF no longer resolves its equation number.");

            AssertEqual(
                0,
                document.Tables.Count,
                "Native save/reopen introduced a layout table.");
            AssertEqual(
                0,
                document.Shapes.Count,
                "Native save/reopen introduced a Shape.");
            AssertNoHostCoreSentinelArtifacts(
                document,
                "after native numbered OMML save/reopen");

            Console.WriteLine(
                "[host-core minimal] Word-native numbering + REF save/reopen persistence passed.");
        }
        finally
        {
            Release(referenceField);
            Release(fields);
            Release(identityRange);
            Release(identity);
            Release(bookmarks);
            Release(selection);
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
            try { mainDocument.Activate(); } catch { }
        }
    }

    private static void RunHostCoreUnownedNativeSaveReopenAdoptionSmoke(
        Word.Application application,
        Word.Document mainDocument,
        string artifactRoot)
    {
        const string sourceMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>u</mi><mo>=</mo><mi>v</mi></mrow></math>";
        const string editedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>u</mi><mo>=</mo><mi>v</mi><mo>+</mo><mn>3</mn></mrow></math>";

        var path =
            Path.Combine(
                artifactRoot,
                "host-core-unowned-native-save-reopen.docx");

        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? sourceIdentity = null;
        Word.Range? sourceIdentityRange = null;
        Word.Range? sourceRange = null;
        Word.OMaths? maths = null;
        Word.OMath? pastedMath = null;
        Word.Range? pastedRange = null;
        try
        {
            document =
                application.Documents.Add(
                    Visible: false);
            document.Activate();
            document.SaveAs2(
                path,
                Word.WdSaveFormat.wdFormatXMLDocument);
            var service =
                new WordFormulaService(
                    application);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            selection = application.Selection;
            selection.SetRange(0, 0);
            var sourceSession =
                CreateOmmlMathTypeAcceptanceSession(
                    sourceMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    sourceSession,
                    sourceMathMl);

            bookmarks = document.Bookmarks;
            var sourceIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    created.FormulaId);
            var sourceRefName =
                WordNativeOmmlNumbering.ReferenceBookmarkName(
                    created.FormulaId);
            sourceIdentity =
                bookmarks[sourceIdentityName];
            sourceIdentityRange =
                sourceIdentity.Range.Duplicate;
            var sourceHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    sourceIdentityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Unowned save/reopen source cannot be resolved.");
            sourceRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    sourceHost.Range);
            sourceRange.Select();

            var snapshot =
                service.CaptureSelectedFormulaForCopy()
                ?? throw new InvalidDataException(
                    "Unowned save/reopen source was not captured for copy.");
            selection.Copy();
            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            var pasteStart =
                selection.Range.Start;
            service.ArmHostCorePaste(
                snapshot);
            selection.Paste();
            System.Threading.Thread.Sleep(120);
            AssertEqual(
                WordFormulaService.PastedFormulaRepairResult.Repaired,
                service.RepairPastedFormula(
                    snapshot),
                "Unowned save/reopen paste validation did not complete.");

            Release(maths);
            maths = document.OMaths;
            AssertEqual(
                2,
                maths.Count,
                "Unowned save/reopen fixture did not contain source + pasted OMath.");
            Release(pastedMath);
            pastedMath = maths[maths.Count];
            Release(pastedRange);
            pastedRange =
                pastedMath.Range.Duplicate;
            AssertTrue(
                pastedRange.Start >= pasteStart,
                "Unowned save/reopen pasted OMath is outside the native paste region.");

            var beforeSave =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    pastedRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Unowned pasted OMath cannot be resolved before save.");
            beforeSave.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    beforeSave);
            AssertTrue(
                string.IsNullOrWhiteSpace(
                    beforeSave.FormulaId),
                "Native paste acquired a durable FormulaId before save.");
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                beforeSave.Numbering.ContainerKind,
                "Unowned pasted OMath is not native #(SEQ) before save.");
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(
                    sourceMathMl),
                MathTypeMtefCodec.SemanticSignature(
                    WordFormulaHostSemanticReader.Read(
                        document,
                        beforeSave).MathMl
                    ?? throw new InvalidDataException(
                        "Unowned pasted OMath has no semantic MathML before save.")),
                "Unowned pasted OMath semantics changed before save.");

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertEqual(
                1,
                CountHostCoreBookmarks(
                    bookmarks,
                    WordFormulaIdentityStore.OmmlBookmarkPrefix),
                "Unowned paste unexpectedly created a VTOMML identity before save.");
            AssertEqual(
                1,
                CountHostCoreBookmarks(
                    bookmarks,
                    "VTEqNum_"),
                "Unowned paste unexpectedly created a VTEqNum target before save.");
            AssertTrue(
                bookmarks.Exists(sourceIdentityName)
                && bookmarks.Exists(sourceRefName),
                "Source identity/reference target changed before save.");

            document.Save();
            document.Close(
                Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            Release(pastedRange);
            pastedRange = null;
            Release(pastedMath);
            pastedMath = null;
            Release(maths);
            maths = null;
            Release(sourceRange);
            sourceRange = null;
            Release(sourceIdentityRange);
            sourceIdentityRange = null;
            Release(sourceIdentity);
            sourceIdentity = null;
            Release(bookmarks);
            bookmarks = null;
            Release(selection);
            selection = null;

            document =
                application.Documents.Open(
                    path,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);
            document.Activate();
            service =
                new WordFormulaService(
                    application);
            selection = application.Selection;

            maths = document.OMaths;
            AssertEqual(
                2,
                maths.Count,
                "Save/reopen changed the source + unowned pasted OMath count.");
            pastedMath = maths[maths.Count];
            pastedRange =
                pastedMath.Range.Duplicate;
            var reopened =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    pastedRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Saved/reopened unowned native OMath cannot be resolved locally.");
            reopened.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    reopened);
            AssertTrue(
                string.IsNullOrWhiteSpace(
                    reopened.FormulaId),
                "Save/reopen implicitly adopted an unowned native OMath.");
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                reopened.Numbering.ContainerKind,
                "Saved/reopened unowned OMath is no longer Word-native #(SEQ).");
            var reopenedPayload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    reopened);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(
                    sourceMathMl),
                MathTypeMtefCodec.SemanticSignature(
                    reopenedPayload.MathMl
                    ?? throw new InvalidDataException(
                        "Saved/reopened unowned OMath returned no MathML.")),
                "Save/reopen changed unowned native formula semantics.");
            AssertTrue(
                reopenedPayload.Latex.IndexOf(
                    "#",
                    StringComparison.Ordinal) < 0
                && reopenedPayload.Latex.IndexOf(
                    "VisualTeXEquation",
                    StringComparison.OrdinalIgnoreCase) < 0,
                "Saved/reopened unowned number leaked into editor semantics.");

            bookmarks = document.Bookmarks;
            AssertEqual(
                1,
                CountHostCoreBookmarks(
                    bookmarks,
                    WordFormulaIdentityStore.OmmlBookmarkPrefix),
                "Save/reopen invented a VTOMML identity for the unowned paste.");
            AssertEqual(
                1,
                CountHostCoreBookmarks(
                    bookmarks,
                    "VTEqNum_"),
                "Save/reopen invented a VTEqNum target for the unowned paste.");

            pastedRange.Select();
            var opened =
                service.ReadSelection();
            AssertTrue(
                opened.Metadata is not null
                && opened.Metadata.Numbered,
                "ReadSelection lost saved/reopened unowned numbered state.");
            AssertTrue(
                !string.IsNullOrWhiteSpace(
                    opened.FormulaId)
                && !string.Equals(
                    opened.FormulaId,
                    created.FormulaId,
                    StringComparison.OrdinalIgnoreCase),
                "ReadSelection did not allocate a fresh session ID after reopening unowned OMML.");
            var adoptedIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    opened.FormulaId!);
            var adoptedRefName =
                WordNativeOmmlNumbering.ReferenceBookmarkName(
                    opened.FormulaId!);
            AssertTrue(
                !bookmarks.Exists(adoptedIdentityName)
                && !bookmarks.Exists(adoptedRefName),
                "ReadSelection persisted identity while merely reading reopened unowned OMML.");

            var editSession =
                CreateOmmlMathTypeAcceptanceSession(
                    editedMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            editSession.Mode = "edit";
            editSession.FormulaId =
                opened.FormulaId!;
            editSession.SourceDocumentId =
                opened.DocumentId;
            editSession.SourceObjectId =
                opened.ObjectId;
            editSession.OriginalMetadata =
                opened.Metadata;

            _ = service.ReplaceOmml(
                editSession,
                editedMathMl);

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(adoptedIdentityName),
                "First edit after save/reopen did not adopt VTOMML identity.");
            AssertTrue(
                bookmarks.Exists(adoptedRefName),
                "First edit after save/reopen did not create VTEqNum target.");

            Word.Bookmark? adoptedIdentity = null;
            Word.Range? adoptedRange = null;
            try
            {
                adoptedIdentity =
                    bookmarks[adoptedIdentityName];
                adoptedRange =
                    adoptedIdentity.Range.Duplicate;
                var adopted =
                    WordFormulaHostResolver.ResolveLocal(
                        document,
                        adoptedRange,
                        WordFormulaHostKind.Omml)
                    ?? throw new InvalidDataException(
                        "Adopted reopened OMML cannot be resolved from its new identity.");
                adopted.Numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        adopted);
                AssertEqual(
                    WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                    adopted.Numbering.ContainerKind,
                    "Adopted reopened OMML lost native #(SEQ) numbering.");
                AssertEqual(
                    MathTypeMtefCodec.SemanticSignature(
                        editedMathMl),
                    MathTypeMtefCodec.SemanticSignature(
                        WordFormulaHostSemanticReader.Read(
                            document,
                            adopted).MathMl
                        ?? throw new InvalidDataException(
                            "Adopted reopened OMML returned no MathML.")),
                    "First edit after reopen changed requested formula semantics.");
            }
            finally
            {
                Release(adoptedRange);
                Release(adoptedIdentity);
            }

            AssertNoHostCoreSentinelArtifacts(
                document,
                "after saved/reopened unowned native OMML first-edit adoption");
            Console.WriteLine(
                "[host-core minimal] unowned native OMML save/reopen + lazy first-edit adoption passed.");
        }
        finally
        {
            Release(pastedRange);
            Release(pastedMath);
            Release(maths);
            Release(sourceRange);
            Release(sourceIdentityRange);
            Release(sourceIdentity);
            Release(bookmarks);
            Release(selection);
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
            try { mainDocument.Activate(); } catch { }
        }
    }

    private static int CountHostCoreBookmarks(
        Word.Bookmarks bookmarks,
        string prefix)
    {
        var count = 0;
        Word.Bookmark? bookmark = null;
        try
        {
            for (var index = 1;
                 index <= bookmarks.Count;
                 index++)
            {
                bookmark = bookmarks[index];
                if ((bookmark.Name ?? string.Empty)
                    .StartsWith(
                        prefix,
                        StringComparison.OrdinalIgnoreCase))
                    count++;
                Release(bookmark);
                bookmark = null;
            }
            return count;
        }
        finally
        {
            Release(bookmark);
        }
    }

    private static void RunHostCoreMathTypeBoundarySmoke(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        string pngPath,
        string emfPath)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>m</mi><mo>+</mo><mn>3</mn></mrow></math>";
        const string prefixText = "mt-left ";
        const string suffixText = " mt-right";

        var previousNativePreview =
            Environment.GetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW");

        Word.Selection? selection = null;
        Word.Range? insertion = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? identity = null;
        Word.Range? identityRange = null;
        Word.Range? ommlRange = null;
        try
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                "1");

            selection = application.Selection;
            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            selection.TypeText(prefixText);

            insertion =
                selection.Range.Duplicate;
            var sourceSession =
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "inline",
                    numbered: false,
                    FormulaOleContract.MathTypeOleMode);
            sourceSession.SourceDocumentId =
                document.FullName;
            sourceSession.SourceObjectId =
                WordRangeReference(
                    insertion.Start,
                    insertion.End);

            _ = service.InsertMathTypeOle(
                sourceSession,
                mathMl,
                emfPath,
                updateCreatedMathTypeNumberFields: true);

            Release(insertion);
            insertion = null;
            selection.TypeText(suffixText);

            AssertEqual(
                1,
                CountMathTypeOleShapes(document),
                "MathType boundary smoke did not create exactly one MathType OLE.");

            shapes = document.InlineShapes;
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(
                            candidate))
                        continue;
                    if (shape is not null)
                        throw new InvalidDataException(
                            "MathType boundary smoke found more than one MathType source.");
                    shape = candidate;
                    candidate = null;
                }
                finally
                {
                    Release(candidate);
                }
            }
            AssertTrue(
                shape is not null,
                "MathType boundary smoke could not resolve its source OLE.");
            shapeRange =
                shape!.Range.Duplicate;
            AssertHostCoreImmediateText(
                document,
                shapeRange,
                prefixText,
                suffixText,
                "before MathType→OMML");
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    MathTypeOleStorage.ReadMathMl(shape)),
                "MathType source semantics changed before conversion.");

            shapeRange.Select();
            var toOmmlPlan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: false,
                    FormulaOleContract.MathTypeOleMode,
                    FormulaOleContract.WordOmmlMode);
            AssertEqual(
                1,
                toOmmlPlan.Targets.Count,
                "MathType→OMML did not capture exactly the selected MathType OLE.");
            var mathTypeSourceId =
                toOmmlPlan.Targets[0].SourceFormulaId;
            var toOmmlResult =
                service.ApplyFormulaFormatConversionPlan(
                    toOmmlPlan,
                    PrepareOmmlMathTypeTargets(
                        toOmmlPlan,
                        emfPath));
            AssertEqual(
                1,
                toOmmlResult.FormulaCount,
                "MathType→OMML did not convert exactly one formula.");
            AssertEqual(
                0,
                toOmmlResult.FailedFormulaCount,
                "MathType→OMML reported a failure.");
            AssertEqual(
                0,
                CountMathTypeOleShapes(document),
                "MathType→OMML left the source MathType OLE alive.");

            Release(shapeRange);
            shapeRange = null;
            Release(shape);
            shape = null;
            Release(shapes);
            shapes = null;

            bookmarks = document.Bookmarks;
            var ommlIdentityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    mathTypeSourceId);
            AssertTrue(
                bookmarks.Exists(
                    ommlIdentityName),
                "MathType→OMML did not bind an exact OMML identity.");
            identity =
                bookmarks[ommlIdentityName];
            identityRange =
                identity.Range.Duplicate;
            var ommlHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    identityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "MathType→OMML target cannot be resolved from its exact identity.");
            AssertEqual(
                mathTypeSourceId,
                ommlHost.FormulaId ?? string.Empty,
                "MathType→OMML changed FormulaId.");
            var ommlPayload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    ommlHost);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    ommlPayload.MathMl
                    ?? throw new InvalidDataException(
                        "MathType→OMML target has no MathML.")),
                "MathType→OMML changed formula semantics.");

            ommlRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    ommlHost.Range);
            AssertHostCoreImmediateText(
                document,
                ommlRange,
                prefixText,
                suffixText,
                "after MathType→OMML");
            AssertNoHostCoreSentinelArtifacts(
                document,
                "after MathType→OMML");

            ommlRange.Select();
            var toMathTypePlan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: false,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.MathTypeOleMode);
            AssertEqual(
                1,
                toMathTypePlan.Targets.Count,
                "OMML→MathType did not capture exactly the converted OMML.");
            var toMathTypeResult =
                service.ApplyFormulaFormatConversionPlan(
                    toMathTypePlan,
                    PrepareOmmlMathTypeTargets(
                        toMathTypePlan,
                        emfPath));
            AssertEqual(
                1,
                toMathTypeResult.FormulaCount,
                "OMML→MathType did not convert exactly one formula.");
            AssertEqual(
                0,
                toMathTypeResult.FailedFormulaCount,
                "OMML→MathType reported a failure.");
            AssertEqual(
                1,
                CountMathTypeOleShapes(document),
                "OMML→MathType did not recreate one MathType OLE.");

            Release(ommlRange);
            ommlRange = null;
            Release(identityRange);
            identityRange = null;
            Release(identity);
            identity = null;
            Release(bookmarks);
            bookmarks = null;

            shapes = document.InlineShapes;
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(
                            candidate))
                        continue;
                    if (shape is not null)
                        throw new InvalidDataException(
                            "OMML→MathType created more than one MathType OLE.");
                    shape = candidate;
                    candidate = null;
                }
                finally
                {
                    Release(candidate);
                }
            }
            AssertTrue(
                shape is not null,
                "OMML→MathType target OLE is missing.");
            shapeRange =
                shape!.Range.Duplicate;
            AssertHostCoreImmediateText(
                document,
                shapeRange,
                prefixText,
                suffixText,
                "after OMML→MathType");
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    MathTypeOleStorage.ReadMathMl(shape)),
                "OMML→MathType changed formula semantics.");

            shapeRange.Select();
            var toVisualTeXPlan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: false,
                    FormulaOleContract.MathTypeOleMode,
                    FormulaOleContract.NativeOleMode);
            AssertEqual(
                1,
                toVisualTeXPlan.Targets.Count,
                "MathType→VisualTeX did not capture exactly the selected MathType OLE.");
            var toVisualTeXTarget =
                toVisualTeXPlan.Targets.Single();
            var toVisualTeXMathMl =
                toVisualTeXTarget.SourceMathMl
                ?? mathMl;
            var toVisualTeXPrepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal)
                {
                    [toVisualTeXTarget.Id] =
                        new PreparedWordBulkFormula
                        {
                            Run = new WordBulkRun
                            {
                                Id = toVisualTeXTarget.Id,
                                IsFormula = true,
                                Latex = toVisualTeXTarget.Latex,
                                DisplayMode =
                                    toVisualTeXTarget.DisplayMode,
                            },
                            Session =
                                CreateSimpleFormatTargetSession(
                                    toVisualTeXTarget,
                                    FormulaOleContract.NativeOleMode,
                                    toVisualTeXMathMl),
                            MathMl =
                                toVisualTeXMathMl,
                            PngPath = pngPath,
                            EmfPath = emfPath,
                        },
                };
            var toVisualTeXResult =
                service.ApplyFormulaFormatConversionPlan(
                    toVisualTeXPlan,
                    toVisualTeXPrepared);
            AssertEqual(
                1,
                toVisualTeXResult.FormulaCount,
                "MathType→VisualTeX did not convert exactly one formula.");
            AssertEqual(
                0,
                toVisualTeXResult.FailedFormulaCount,
                "MathType→VisualTeX reported a failure.");
            AssertEqual(
                0,
                CountMathTypeOleShapes(document),
                "MathType→VisualTeX left a MathType source behind.");

            Release(shapeRange);
            shapeRange = null;
            Release(shape);
            shape = null;
            Release(shapes);
            shapes = null;

            shapes = document.InlineShapes;
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(
                            candidate))
                        continue;
                    if (shape is not null)
                        throw new InvalidDataException(
                            "MathType→VisualTeX created more than one VisualTeX OLE.");
                    shape = candidate;
                    candidate = null;
                }
                finally
                {
                    Release(candidate);
                }
            }
            AssertTrue(
                shape is not null,
                "MathType→VisualTeX target is missing.");
            shapeRange =
                shape!.Range.Duplicate;
            var visualMetadata =
                WordFormulaMetadataReader
                    .TryReadEmbeddedNativeOle(
                        shape)
                ?? throw new InvalidDataException(
                    "MathType→VisualTeX target has no embedded metadata.");
            AssertTrue(
                Guid.TryParse(
                    toVisualTeXTarget.SourceFormulaId,
                    out _),
                "MathType→VisualTeX capture did not assign a valid conversion identity.");
            AssertEqual(
                toVisualTeXTarget.SourceFormulaId,
                visualMetadata.FormulaId,
                "MathType→VisualTeX did not bind the identity captured for this conversion.");
            AssertHostCoreImmediateText(
                document,
                shapeRange,
                prefixText,
                suffixText,
                "after MathType→VisualTeX");
            AssertNoHostCoreSentinelArtifacts(
                document,
                "after MathType→VisualTeX");

            shapeRange.Select();
            var visualToMathTypePlan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: false,
                    FormulaOleContract.NativeOleMode,
                    FormulaOleContract.MathTypeOleMode);
            AssertEqual(
                1,
                visualToMathTypePlan.Targets.Count,
                "VisualTeX→MathType did not capture exactly the selected VisualTeX OLE.");
            var visualToMathTypeTarget =
                visualToMathTypePlan.Targets.Single();
            var visualToMathTypePrepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal)
                {
                    [visualToMathTypeTarget.Id] =
                        new PreparedWordBulkFormula
                        {
                            Run = new WordBulkRun
                            {
                                Id = visualToMathTypeTarget.Id,
                                IsFormula = true,
                                Latex =
                                    visualToMathTypeTarget.Latex,
                                DisplayMode =
                                    visualToMathTypeTarget.DisplayMode,
                            },
                            Session =
                                CreateSimpleMathTypeTargetSession(
                                    visualToMathTypeTarget,
                                    mathMl),
                            MathMl = mathMl,
                            EmfPath = emfPath,
                        },
                };
            var visualToMathTypeResult =
                service.ApplyFormulaFormatConversionPlan(
                    visualToMathTypePlan,
                    visualToMathTypePrepared);
            AssertEqual(
                1,
                visualToMathTypeResult.FormulaCount,
                "VisualTeX→MathType did not convert exactly one formula.");
            AssertEqual(
                0,
                visualToMathTypeResult.FailedFormulaCount,
                "VisualTeX→MathType reported a failure.");
            AssertEqual(
                1,
                CountMathTypeOleShapes(document),
                "VisualTeX→MathType did not recreate one MathType OLE.");
            AssertEqual(
                0,
                CountVisualTeXNativeOleShapes(document),
                "VisualTeX→MathType left a VisualTeX OLE behind.");

            Release(shapeRange);
            shapeRange = null;
            Release(shape);
            shape = null;
            Release(shapes);
            shapes = document.InlineShapes;
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(
                            candidate))
                        continue;
                    shape = candidate;
                    candidate = null;
                    break;
                }
                finally
                {
                    Release(candidate);
                }
            }
            AssertTrue(
                shape is not null,
                "VisualTeX→MathType final OLE is missing.");
            shapeRange =
                shape!.Range.Duplicate;
            AssertHostCoreImmediateText(
                document,
                shapeRange,
                prefixText,
                suffixText,
                "after VisualTeX→MathType");
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    MathTypeOleStorage.ReadMathMl(shape)),
                "VisualTeX→MathType changed formula semantics.");

            Console.WriteLine(
                "[host-core minimal] MathType↔OMML and MathType↔VisualTeX boundary conversions passed.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                previousNativePreview);
            Release(ommlRange);
            Release(identityRange);
            Release(identity);
            Release(bookmarks);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(insertion);
            Release(selection);
        }
    }

    private static void AssertHostCoreImmediateText(
        Word.Document document,
        Word.Range hostRange,
        string expectedPrefix,
        string expectedSuffix,
        string stage)
    {
        Word.Range? before = null;
        Word.Range? after = null;
        try
        {
            var beforeStart =
                Math.Max(
                    document.Content.Start,
                    hostRange.Start
                    - expectedPrefix.Length);
            before =
                document.Range(
                    beforeStart,
                    hostRange.Start);
            AssertEqual(
                expectedPrefix,
                before.Text ?? string.Empty,
                stage + ": text immediately before the host changed.");

            var afterEnd =
                Math.Min(
                    document.Content.End,
                    hostRange.End
                    + expectedSuffix.Length);
            after =
                document.Range(
                    hostRange.End,
                    afterEnd);
            AssertEqual(
                expectedSuffix,
                after.Text ?? string.Empty,
                stage + ": text immediately after the host changed.");
        }
        finally
        {
            Release(after);
            Release(before);
        }
    }

    private static void RunHostCoreUserTableNumberingSmoke(
        Word.Application application,
        Word.Document document,
        WordFormulaService service)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>a</mi><mo>=</mo><mi>b</mi></mrow></math>";

        Word.Selection? selection = null;
        Word.Range? anchor = null;
        Word.Table? userTable = null;
        Word.Cell? cell = null;
        Word.Range? cellRange = null;
        Word.Range? insertion = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? identity = null;
        Word.Range? identityRange = null;
        Word.Range? hostRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.OMaths? maths = null;
        try
        {
            var tableCountBefore =
                document.Tables.Count;

            selection = application.Selection;
            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            anchor =
                selection.Range.Duplicate;
            userTable =
                document.Tables.Add(
                    anchor,
                    1,
                    1);
            var tableCountAfterCreate =
                document.Tables.Count;
            AssertEqual(
                tableCountBefore + 1,
                tableCountAfterCreate,
                "User-table smoke did not create exactly one user table.");

            cell = userTable.Cell(1, 1);
            cellRange =
                cell.Range.Duplicate;
            insertion =
                document.Range(
                    cellRange.Start,
                    cellRange.Start);
            insertion.Select();

            var session =
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    session,
                    mathMl);

            AssertEqual(
                tableCountAfterCreate,
                document.Tables.Count,
                "Numbering inside a user table created a nested/top-level VisualTeX table.");

            bookmarks = document.Bookmarks;
            var identityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    created.FormulaId);
            AssertTrue(
                bookmarks.Exists(identityName),
                "Numbered user-table OMML has no exact host identity.");
            identity =
                bookmarks[identityName];
            identityRange =
                identity.Range.Duplicate;

            var host =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    identityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Numbered user-table OMML cannot be resolved locally.");
            AssertTrue(
                host.WithinTable,
                "User-table OMML resolver lost its table ownership.");

            var numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            AssertTrue(
                numbering.Numbered,
                "User-table display OMML was requested numbered but resolver reports unnumbered.");
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                numbering.ContainerKind,
                "User-table display OMML did not remain one Word-native #(SEQ) OMath.");

            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            paragraphs =
                hostRange.Paragraphs;
            AssertEqual(
                1,
                paragraphs.Count,
                "Numbered user-table formula no longer occupies one paragraph.");
            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            fields =
                paragraphRange.Fields;

            var sequenceCount = 0;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code =
                    field.Code.Duplicate;
                if ((code.Text ?? string.Empty)
                    .IndexOf(
                        "SEQ VisualTeXEquation",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    sequenceCount++;
            }
            AssertEqual(
                1,
                sequenceCount,
                "Word-native numbered OMath in the user table does not contain exactly one VisualTeX SEQ field.");

            var semanticPayload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    host);
            AssertTrue(
                (semanticPayload.Latex ?? string.Empty)
                    .IndexOf("#", StringComparison.Ordinal) < 0
                && (semanticPayload.Latex ?? string.Empty)
                    .IndexOf(
                        "VisualTeXEquation",
                        StringComparison.OrdinalIgnoreCase) < 0,
                "User-table double-click semantics exposed native numbering content.");

            var referenceName =
                WordFormulaNumberingKernel.ReferenceBookmarkName(
                    created.FormulaId);
            AssertTrue(
                bookmarks.Exists(referenceName),
                "Canonical user-table number has no VTEqNum target.");

            cellRange =
                cell.Range.Duplicate;
            maths = cellRange.OMaths;
            AssertEqual(
                1,
                maths.Count,
                "Canonical user-table numbering moved or duplicated the OMath host.");

            AssertNoHostCoreSentinelArtifacts(
                document,
                "after canonical user-table numbering");

            Release(maths);
            maths = null;
            Release(fields);
            fields = null;
            Release(paragraphRange);
            paragraphRange = null;
            Release(paragraph);
            paragraph = null;
            Release(paragraphs);
            paragraphs = null;
            Release(hostRange);
            hostRange = null;

            identityRange.Select();
            var deleted =
                service.DeleteSelectedFormula();
            AssertEqual(
                created.FormulaId,
                deleted,
                "Deleting the user-table formula returned another FormulaId.");
            AssertEqual(
                tableCountAfterCreate,
                document.Tables.Count,
                "Deleting a numbered user-table formula deleted or added a user table.");
            AssertTrue(
                !bookmarks.Exists(referenceName),
                "Deleting the user-table formula left its VTEqNum target alive.");

            Release(cellRange);
            cellRange =
                cell.Range.Duplicate;
            maths = cellRange.OMaths;
            AssertEqual(
                0,
                maths.Count,
                "Deleting the user-table formula left an OMath in the cell.");
            Release(maths);
            maths = null;

            fields = cellRange.Fields;
            sequenceCount = 0;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code =
                    field.Code.Duplicate;
                if ((code.Text ?? string.Empty)
                    .IndexOf(
                        "SEQ VisualTeXEquation",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    sequenceCount++;
            }
            AssertEqual(
                0,
                sequenceCount,
                "Deleting the user-table formula left a VisualTeX SEQ field behind.");

            AssertNoHostCoreSentinelArtifacts(
                document,
                "after deleting canonical user-table formula");

            Console.WriteLine(
                "[host-core minimal] Word-native user-table numbering lifecycle passed.");
        }
        finally
        {
            Release(maths);
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(hostRange);
            Release(identityRange);
            Release(identity);
            Release(bookmarks);
            Release(insertion);
            Release(cellRange);
            Release(cell);
            Release(userTable);
            Release(anchor);
            Release(selection);
        }
    }

    private static void RunHostCoreCanonicalNumberingSmoke(
        Word.Application application,
        Word.Document document,
        WordFormulaService service)
    {
        const string displayMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><msup><mi>z</mi><mn>2</mn></msup><mo>+</mo><mn>1</mn></mrow></math>";

        Word.Selection? selection = null;
        Word.Range? content = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? identity = null;
        Word.Range? identityRange = null;
        Word.Range? containerRange = null;
        Word.Tables? tables = null;
        Word.Table? table = null;
        Word.Rows? rows = null;
        Word.Columns? columns = null;
        Word.Cell? rightCell = null;
        Word.Range? rightRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? fieldCode = null;
        Word.Bookmarks? rightBookmarks = null;
        Word.Bookmark? rightBookmark = null;
        try
        {
            selection = application.Selection;
            content = document.Content;
            var paragraphEnd =
                Math.Max(
                    content.Start,
                    content.End - 1);
            selection.SetRange(
                paragraphEnd,
                paragraphEnd);
            selection.TypeParagraph();

            var session =
                CreateOmmlMathTypeAcceptanceSession(
                    displayMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode);
            var created =
                service.InsertOmml(
                    session,
                    displayMathMl);

            bookmarks = document.Bookmarks;
            var identityName =
                WordFormulaIdentityStore.OmmlBookmarkName(
                    created.FormulaId);
            AssertTrue(
                bookmarks.Exists(identityName),
                "Numbered display OMML lost its exact host identity.");
            identity = bookmarks[identityName];
            identityRange =
                identity.Range.Duplicate;

            var host =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    identityRange,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Numbered display OMML cannot be resolved from its own identity position.");
            AssertEqual(
                created.FormulaId,
                host.FormulaId ?? string.Empty,
                "Numbered display OMML resolved to another FormulaId.");
            AssertEqual(
                "block",
                host.DisplayMode,
                "Numbered display OMML is not wdOMathDisplay.");

            var numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            AssertTrue(
                numbering.Numbered,
                "Display OMML was requested numbered but resolver reports unnumbered.");
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                numbering.ContainerKind,
                "Display OMML did not remain one Word-native #(SEQ) OMath host.");
            AssertTrue(
                numbering.ContainerRange is not null,
                "Canonical numbered OMML has no native OMath container range.");
            AssertTrue(
                numbering.NumberRange is not null,
                "Canonical numbered OMML has no visible number range.");

            containerRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    numbering.ContainerRange!);
            AssertEqual(
                0,
                containerRange.Tables.Count,
                "Word-native numbered OMML unexpectedly created a layout table.");

            fields = containerRange.Fields;
            var sequenceCount = 0;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(fieldCode);
                fieldCode = null;
                Release(field);
                field = fields[index];
                fieldCode = field.Code.Duplicate;
                if ((fieldCode.Text ?? string.Empty).IndexOf(
                        "SEQ VisualTeXEquation",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    sequenceCount++;
            }
            AssertEqual(
                1,
                sequenceCount,
                "Word-native numbered OMML does not contain exactly one SEQ VisualTeXEquation field.");

            var nativeXml =
                containerRange.WordOpenXML
                ?? string.Empty;
            AssertTrue(
                nativeXml.IndexOf(
                    "<m:eqArr",
                    StringComparison.OrdinalIgnoreCase) >= 0,
                "Word-native numbered OMML lost its m:eqArr host.");

            var semanticPayload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    host);
            AssertTrue(
                (semanticPayload.Latex ?? string.Empty)
                    .IndexOf(
                        "VisualTeXEquation",
                        StringComparison.OrdinalIgnoreCase) < 0
                && (semanticPayload.Latex ?? string.Empty)
                    .IndexOf("#", StringComparison.Ordinal) < 0,
                "Semantic reader exposed native numbering content to the editor.");
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(displayMathMl),
                MathTypeMtefCodec.SemanticSignature(
                    semanticPayload.MathMl
                    ?? throw new InvalidDataException(
                        "Native numbered OMML semantic reader returned no MathML.")),
                "Native #(SEQ) changed the formula semantic body.");

            containerRange.Select();
            var editSelection =
                service.ReadSelection();
            AssertTrue(
                editSelection.Metadata is not null,
                "ReadSelection did not return editable metadata for native numbered OMML.");
            AssertEqual(
                created.FormulaId,
                editSelection.FormulaId,
                "ReadSelection changed the numbered OMML FormulaId.");
            AssertTrue(
                editSelection.Metadata!.Numbered,
                "ReadSelection lost the numbered flag for native #(SEQ) OMML.");
            AssertTrue(
                editSelection.Metadata.Latex.IndexOf(
                    "#",
                    StringComparison.Ordinal) < 0
                && editSelection.Metadata.Latex.IndexOf(
                    "VisualTeXEquation",
                    StringComparison.OrdinalIgnoreCase) < 0,
                "ReadSelection exposed Word numbering syntax to the formula editor.");

            var referenceName =
                WordFormulaNumberingKernel.ReferenceBookmarkName(
                    created.FormulaId);
            rightBookmarks =
                containerRange.Bookmarks;
            var targetCount = 0;
            for (var index = 1;
                 index <= rightBookmarks.Count;
                 index++)
            {
                Release(rightBookmark);
                rightBookmark =
                    rightBookmarks[index];
                if (string.Equals(
                        rightBookmark.Name,
                        referenceName,
                        StringComparison.OrdinalIgnoreCase))
                    targetCount++;
            }
            AssertEqual(
                1,
                targetCount,
                "Word-native numbered OMML does not contain exactly one VTEqNum target.");

            for (var index = 1;
                 index <= bookmarks.Count;
                 index++)
            {
                Release(rightBookmark);
                rightBookmark =
                    bookmarks[index];
                var name =
                    rightBookmark.Name
                    ?? string.Empty;
                AssertTrue(
                    !name.StartsWith(
                        "VTEq_",
                        StringComparison.OrdinalIgnoreCase)
                    && !name.StartsWith(
                        "VTEqCap_",
                        StringComparison.OrdinalIgnoreCase)
                    && !name.StartsWith(
                        "VTEqAnc_",
                        StringComparison.OrdinalIgnoreCase)
                    && !name.StartsWith(
                        "VTEqShape_",
                        StringComparison.OrdinalIgnoreCase),
                    $"Canonical numbering created retired bookmark {name}.");
            }

            AssertNoHostCoreSentinelArtifacts(
                document,
                "after canonical numbered display OMML");

            var referenceTargets =
                service.GetCanonicalEquationReferenceTargets(
                    document);
            var referenceTarget =
                referenceTargets.Single(
                    item => string.Equals(
                        item.FormulaId,
                        created.FormulaId,
                        StringComparison.OrdinalIgnoreCase));

            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            selection.TypeText(
                "Reference: ");
            service.InsertEquationReference(
                document,
                selection,
                referenceTarget,
                EquationReferenceStyle.Parenthesized,
                Word.WdColor.wdColorAutomatic);

            var liveReferenceText =
                ReadHostCoreReferenceResult(
                    document,
                    referenceName,
                    updateNestedRef: true);
            AssertTrue(
                liveReferenceText.IndexOf(
                    referenceTarget.NumberText,
                    StringComparison.Ordinal) >= 0,
                "Inserted REF does not display the current canonical equation number.");

            Release(identityRange);
            identityRange = null;
            Release(identity);
            identity = null;
            Release(bookmarks);
            bookmarks = document.Bookmarks;
            identity = bookmarks[identityName];
            identityRange =
                identity.Range.Duplicate;
            identityRange.Select();

            var deletedFormulaId =
                service.DeleteSelectedFormula();
            AssertEqual(
                created.FormulaId,
                deletedFormulaId,
                "Deleting the numbered formula returned another FormulaId.");

            AssertTrue(
                !bookmarks.Exists(
                    referenceName),
                "Deleting a numbered formula left its VTEqNum target bookmark alive.");

            var remainingTargets =
                service.GetCanonicalEquationReferenceTargets(
                    document);
            AssertTrue(
                !remainingTargets.Any(
                    item => string.Equals(
                        item.FormulaId,
                        created.FormulaId,
                        StringComparison.OrdinalIgnoreCase)),
                "Deleted formula still appears as a canonical REF target.");

            var brokenReferenceText =
                ReadHostCoreReferenceResult(
                    document,
                    referenceName,
                    updateNestedRef: true);
            AssertTrue(
                IsMissingReferenceResult(
                    brokenReferenceText),
                "REF field survived target deletion but Word did not report a missing reference source. "
                + $"result=[{brokenReferenceText}]");

            AssertNoHostCoreSentinelArtifacts(
                document,
                "after numbered formula deletion with live REF");

            Console.WriteLine(
                "[host-core minimal] canonical numbering + live REF missing-target lifecycle passed.");
        }
        finally
        {
            Release(rightBookmark);
            Release(rightBookmarks);
            Release(fieldCode);
            Release(field);
            Release(fields);
            Release(rightRange);
            Release(rightCell);
            Release(columns);
            Release(rows);
            Release(table);
            Release(tables);
            Release(containerRange);
            Release(identityRange);
            Release(identity);
            Release(bookmarks);
            Release(content);
            Release(selection);
        }
    }

    private static string ReadHostCoreReferenceResult(
        Word.Document document,
        string referenceName,
        bool updateNestedRef)
    {
        Word.Fields? fields = null;
        Word.Field? reference = null;
        Word.Range? code = null;
        Word.Range? result = null;
        try
        {
            fields = document.Fields;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(code);
                code = null;
                Release(reference);
                reference = fields[index];
                if (reference.Type !=
                    Word.WdFieldType.wdFieldRef)
                    continue;

                code =
                    reference.Code.Duplicate;
                var instruction =
                    (code.Text ?? string.Empty)
                    .Trim();
                if (!instruction.StartsWith(
                        "REF ",
                        StringComparison.OrdinalIgnoreCase)
                    || instruction.IndexOf(
                        referenceName,
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                // The canonical Host Core reference is one native Word REF field.
                // \h supplies Word's own navigation behavior; there is no private
                // GOTOBUTTON wrapper or nested field tree.
                AssertTrue(
                    instruction.IndexOf(
                        "\\h",
                        StringComparison.OrdinalIgnoreCase) >= 0,
                    "Canonical equation REF lost Word's hyperlink/navigation switch.");

                if (updateNestedRef)
                    _ = reference.Update();

                Release(result);
                result = null;
                result =
                    reference.Result.Duplicate;
                return (
                        result.Text
                        ?? string.Empty)
                    .Trim();
            }

            throw new InvalidDataException(
                $"Native REF reference for {referenceName} is missing.");
        }
        finally
        {
            Release(result);
            Release(code);
            Release(reference);
            Release(fields);
        }
    }

    private static bool IsMissingReferenceResult(
        string value)
    {
        var normalized =
            (value ?? string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();
        if (normalized.Length == 0)
            return false;

        return normalized.IndexOf(
                   "error",
                   StringComparison.OrdinalIgnoreCase) >= 0
            || normalized.IndexOf(
                   "错误",
                   StringComparison.Ordinal) >= 0
            || normalized.IndexOf(
                   "未找到",
                   StringComparison.Ordinal) >= 0
            || normalized.IndexOf(
                   "エラー",
                   StringComparison.Ordinal) >= 0
            || normalized.IndexOf(
                   "見つかりません",
                   StringComparison.Ordinal) >= 0
            || normalized.IndexOf(
                   "오류",
                   StringComparison.Ordinal) >= 0
            || normalized.IndexOf(
                   "erreur",
                   StringComparison.OrdinalIgnoreCase) >= 0
            || normalized.IndexOf(
                   "fehler",
                   StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static void AssertNoVisualTeXOmmlArtifacts(
        Word.Document document,
        string stage)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1;
                 index <= bookmarks.Count;
                 index++)
            {
                Release(bookmark);
                bookmark = bookmarks[index];
                var name =
                    bookmark.Name
                    ?? string.Empty;
                AssertTrue(
                    !name.StartsWith(
                        "VTOMML_",
                        StringComparison.OrdinalIgnoreCase),
                    $"{stage}: native OMML contains forbidden VisualTeX identity bookmark {name}.");
                AssertTrue(
                    !name.StartsWith(
                        "VTEqNum_",
                        StringComparison.OrdinalIgnoreCase),
                    $"{stage}: native OMML contains forbidden VisualTeX number bookmark {name}.");
            }

            fields = document.Fields;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code.Duplicate;
                var instruction =
                    code.Text
                    ?? string.Empty;
                AssertTrue(
                    instruction.IndexOf(
                        "VisualTeXEquation",
                        StringComparison.OrdinalIgnoreCase) < 0,
                    $"{stage}: native OMML contains forbidden VisualTeX SEQ namespace.");
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void AssertNoHostCoreSentinelArtifacts(
        Word.Document document,
        string stage)
    {
        Word.Range? content = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        try
        {
            content =
                document.Content;
            var text =
                content.Text ?? string.Empty;
            AssertTrue(
                text.IndexOf(
                    '\u200C') < 0,
                $"{stage}: document contains forbidden U+200C.");
            AssertTrue(
                text.IndexOf(
                    '\u200B') < 0,
                $"{stage}: document contains forbidden U+200B.");
            AssertTrue(
                text.IndexOf(
                    '\v') < 0,
                $"{stage}: inline formula introduced a vertical-tab separator.");
            AssertTrue(
                text.IndexOf(
                    '\uE000') < 0,
                $"{stage}: temporary bulk/MathType placeholder U+E000 leaked into the document.");

            bookmarks =
                document.Bookmarks;
            for (var index = 1;
                 index <= bookmarks.Count;
                 index++)
            {
                Release(bookmark);
                bookmark = bookmarks[index];
                AssertTrue(
                    !(bookmark.Name ?? string.Empty)
                        .StartsWith(
                            "VTBL_",
                            StringComparison.OrdinalIgnoreCase),
                    $"{stage}: document contains retired VTBL bookmark {bookmark.Name}.");
            }
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(content);
        }
    }
}
