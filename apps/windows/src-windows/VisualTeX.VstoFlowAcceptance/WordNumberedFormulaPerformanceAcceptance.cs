using System.Diagnostics;
using System.Text;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const string NumberedMiddleInsertionAnchorText =
        "VisualTeX numbered performance middle anchor";

    private static void RunWordNumberedFormulaPerformanceAcceptance(string artifactRoot)
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
        Word.Document? copyDocument = null;
        var originalNumberFormat = WordEquationNumbering.GetDefaultEquationNumberFormatId();
        try
        {
            File.WriteAllText(
                svgPath,
                CreateFontAcceptanceSvg("Times New Roman", "SimSun"),
                new UTF8Encoding(false));
            WriteAcceptancePng(pngPath, "x=1", 240, 72);
            emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 72);

            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(document, "continuous");
            var service = new WordFormulaService(application);
            var formulaIds = new List<string>();
            var insertTimings = new List<long>();
            var targetFormulaCount = int.TryParse(
                Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_COUNT"),
                out var parsedTargetFormulaCount)
                ? Math.Max(6, Math.Min(200, parsedTargetFormulaCount))
                : 6;
            var timingCheckpoints = new HashSet<int>(new[] { 1, 6, 10, 20, 40, 80, 100, 200 });

            for (var index = 1; index <= targetFormulaCount; index++)
            {
                application.Selection.EndKey(Word.WdUnits.wdStory);
                if (targetFormulaCount >= 100 && index == 21)
                {
                    // Create the structural-test anchor from a known main-story
                    // end position before formula #21 captures its insertion
                    // range. Never trust the selection returned by formula #20.
                    Word.Range? anchorProbe = null;
                    Word.Frames? anchorProbeFrames = null;
                    try
                    {
                        anchorProbe = application.Selection.Range;
                        if ((bool)anchorProbe.get_Information(Word.WdInformation.wdWithInTable))
                            throw new InvalidOperationException(
                                "The #20/#21 fixture anchor is still inside a table after EndKey(wdStory).");
                        anchorProbeFrames = anchorProbe.Frames;
                        if (anchorProbeFrames.Count > 0)
                            throw new InvalidOperationException(
                                "The #20/#21 fixture anchor is still inside a caption frame after EndKey(wdStory).");
                    }
                    finally
                    {
                        Release(anchorProbeFrames);
                        Release(anchorProbe);
                    }
                    application.Selection.TypeText(NumberedMiddleInsertionAnchorText);
                    application.Selection.TypeParagraph();
                    application.Selection.TypeText(
                        "VisualTeX numbered performance middle spacer");
                    application.Selection.TypeParagraph();
                    application.Selection.EndKey(Word.WdUnits.wdStory);
                }
                var range = application.Selection.Range;
                var formulaId = Guid.NewGuid().ToString("D");
                var session = CreateNumberedPerformanceSession(
                    "create",
                    formulaId,
                    document.FullName,
                    WordRangeReference(range.Start, range.End),
                    originalMetadata: null,
                    latex: $"x_{{{index}}}={index}");
                Release(range);

                var watch = Stopwatch.StartNew();
                service.InsertOle(session, pngPath, emfPath);
                watch.Stop();
                insertTimings.Add(watch.ElapsedMilliseconds);
                formulaIds.Add(formulaId);

                // A successful numbered display insertion must leave the caret
                // in ordinary body text. If Word keeps it inside the three-cell
                // equation table or the clipped native-caption frame, the next
                // user keystroke becomes part of numbering infrastructure.
                Word.Range? typingProbe = null;
                Word.Frames? typingProbeFrames = null;
                try
                {
                    typingProbe = application.Selection.Range;
                    if ((bool)typingProbe.get_Information(Word.WdInformation.wdWithInTable))
                        throw new InvalidOperationException(
                            $"Numbered OLE append #{index} left the caret inside its equation table.");
                    typingProbeFrames = typingProbe.Frames;
                    if (typingProbeFrames.Count > 0)
                        throw new InvalidOperationException(
                            $"Numbered OLE append #{index} left the caret inside a caption frame.");
                }
                finally
                {
                    Release(typingProbeFrames);
                    Release(typingProbe);
                }

                if (timingCheckpoints.Contains(index) || index == targetFormulaCount)
                {
                    Console.WriteLine(
                        $"    [perf] numbered OLE append #{index}: {watch.ElapsedMilliseconds}ms");
                }
            }

            AssertEqual(targetFormulaCount, document.Tables.Count,
                $"Numbered OLE performance fixture did not create {targetFormulaCount} equation tables.");
            AssertNumberedFormulaArtifacts(document, formulaIds);
            AssertVisibleEquationNumbers(document, formulaIds, 1);

            Word.InlineShape? editShape = null;
            Word.Range? editRange = null;
            try
            {
                var editIndex = targetFormulaCount >= 50 ? 49 : 2;
                editShape = FindNumberedOleByFormulaId(document, formulaIds[editIndex]);
                var originalMetadata = WordFormulaMetadataReader.TryRead(editShape)
                    ?? throw new InvalidOperationException("Numbered OLE edit metadata is missing.");
                editRange = editShape.Range;
                var editSession = CreateNumberedPerformanceSession(
                    "edit",
                    formulaIds[editIndex],
                    document.FullName,
                    WordRangeReference(editRange.Start, editRange.End),
                    originalMetadata,
                    $"x_{{{editIndex + 1}}}=333");
                var watch = Stopwatch.StartNew();
                service.ReplaceOle(editSession, pngPath, emfPath);
                watch.Stop();
                Console.WriteLine(
                    $"    [perf] numbered OLE edit #{editIndex + 1} at {targetFormulaCount} formulas: "
                    + $"{watch.ElapsedMilliseconds}ms");
                if (watch.ElapsedMilliseconds > 1500)
                    throw new InvalidDataException(
                        $"Numbered OLE edit still took {watch.ElapsedMilliseconds}ms.");
            }
            finally
            {
                Release(editRange);
                Release(editShape);
            }

            Word.InlineShape? ommlSourceShape = null;
            Word.Range? ommlSourceRange = null;
            Word.Bookmark? ommlBookmark = null;
            Word.Range? ommlRange = null;
            try
            {
                var formulaId = formulaIds[3];
                ommlSourceShape = FindNumberedOleByFormulaId(document, formulaId);
                var sourceMetadata = WordFormulaMetadataReader.TryRead(ommlSourceShape)
                    ?? throw new InvalidOperationException("Numbered OLE->OMML source metadata is missing.");
                ommlSourceRange = ommlSourceShape.Range;
                var convertSession = CreateNumberedPerformanceSession(
                    "edit",
                    formulaId,
                    document.FullName,
                    WordRangeReference(ommlSourceRange.Start, ommlSourceRange.End),
                    sourceMetadata,
                    @"y_4=4");
                convertSession.ObjectMode = FormulaOleContract.WordOmmlMode;
                service.ReplaceOmml(convertSession, PerformanceMathMl(4, 1));

                ommlBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                    ?? throw new InvalidOperationException("Converted numbered OMML bookmark is missing.");
                var ommlMetadata = WordOmmlFormulaStore.TryRead(document, ommlBookmark)
                    ?? throw new InvalidOperationException("Converted numbered OMML metadata is missing.");
                ommlRange = WordOmmlFormulaStore.GetEquationRange(ommlBookmark);
                var editSession = CreateNumberedPerformanceSession(
                    "edit",
                    formulaId,
                    document.FullName,
                    WordRangeReference(ommlRange.Start, ommlRange.End),
                    ommlMetadata,
                    @"y_4=44");
                editSession.ObjectMode = FormulaOleContract.WordOmmlMode;

                var watch = Stopwatch.StartNew();
                service.ReplaceOmml(editSession, PerformanceMathMl(4, 2));
                watch.Stop();
                Console.WriteLine($"    [perf] numbered OMML edit: {watch.ElapsedMilliseconds}ms");
                if (watch.ElapsedMilliseconds > 1500)
                    throw new InvalidDataException(
                        $"Numbered OMML edit still took {watch.ElapsedMilliseconds}ms.");
            }
            finally
            {
                Release(ommlRange);
                Release(ommlBookmark);
                Release(ommlSourceRange);
                Release(ommlSourceShape);
            }

            Word.Table? sourceTable = null;
            Word.Range? sourceTableRange = null;
            Word.Table? copySourceTable = null;
            Word.Range? copySourceRange = null;
            Word.InlineShape? copiedShape = null;
            Word.Range? copiedShapeRange = null;
            var copiedFormulaIds = new List<string>();
            try
            {
                sourceTable = document.Tables[1];
                sourceTableRange = sourceTable.Range.Duplicate;
                sourceTableRange.Copy();

                copyDocument = application.Documents.Add();
                copyDocument.Activate();
                WordEquationNumbering.SetEquationNumberFormatPreference(copyDocument, "continuous");
                application.Selection.Paste();
                AssertEqual(1, copyDocument.Tables.Count,
                    "Numbered OLE copy performance fixture could not seed the copy document.");

                copiedShape = copyDocument.Tables[1].Cell(1, 2).Range.InlineShapes[1];
                copiedShapeRange = copiedShape.Range;
                copiedShapeRange.Select();
                var firstSelection = service.ReadSelection();
                if (string.IsNullOrWhiteSpace(firstSelection.FormulaId)
                    || firstSelection.Metadata is null)
                    throw new InvalidOperationException(
                        "Seeded numbered OLE formula was not recognized in the copy document.");
                // A table copied into a different document carries the visible
                // number cell but not VisualTeX's hidden caption paragraph. Seed
                // the destination document with one complete local numbering
                // scaffold before measuring same-document duplication.
                WordEquationNumbering.ReconcileFormula(
                    copyDocument,
                    copiedShapeRange,
                    copiedShape.Height,
                    firstSelection.Metadata,
                    numberingOrderMayHaveChanged: true);
                copiedFormulaIds.Add(firstSelection.FormulaId!);
                Release(copiedShapeRange);
                copiedShapeRange = null;
                Release(copiedShape);
                copiedShape = null;

                long finalPasteMs = 0;
                long finalRepairMs = 0;
                for (var copyIndex = 2; copyIndex <= 7; copyIndex++)
                {
                    Release(copySourceRange);
                    copySourceRange = null;
                    Release(copySourceTable);
                    copySourceTable = copyDocument.Tables[copyIndex - 1];
                    copySourceRange = copySourceTable.Range.Duplicate;
                    copySourceRange.Copy();
                    application.Selection.EndKey(Word.WdUnits.wdStory);
                    application.Selection.TypeParagraph();

                    var pasteWatch = Stopwatch.StartNew();
                    application.Selection.Paste();
                    pasteWatch.Stop();
                    AssertEqual(copyIndex, copyDocument.Tables.Count,
                        $"Numbered OLE copy performance fixture did not paste table {copyIndex}.");

                    Release(copiedShapeRange);
                    copiedShapeRange = null;
                    Release(copiedShape);
                    copiedShape = copyDocument.Tables[copyIndex].Cell(1, 2).Range.InlineShapes[1];
                    copiedShapeRange = copiedShape.Range;
                    copiedShapeRange.Select();
                    var repairWatch = Stopwatch.StartNew();
                    var copiedSelection = service.ReadSelection();
                    repairWatch.Stop();
                    if (string.IsNullOrWhiteSpace(copiedSelection.FormulaId)
                        || copiedFormulaIds.Any(id => string.Equals(
                            id,
                            copiedSelection.FormulaId,
                            StringComparison.OrdinalIgnoreCase)))
                        throw new InvalidOperationException(
                            $"Copied numbered OLE formula {copyIndex} did not receive an independent FormulaId.");
                    copiedFormulaIds.Add(copiedSelection.FormulaId!);
                    if (copyIndex == 7)
                    {
                        finalPasteMs = pasteWatch.ElapsedMilliseconds;
                        finalRepairMs = repairWatch.ElapsedMilliseconds;
                    }
                }

                Console.WriteLine(
                    $"    [perf] seventh numbered OLE table paste: {finalPasteMs}ms; "
                    + $"identity/number repair: {finalRepairMs}ms");
                if (finalRepairMs > 1500)
                    throw new InvalidDataException(
                        $"Numbered OLE copy identity/number repair still took {finalRepairMs}ms at seven formulas.");
                AssertNumberedFormulaArtifacts(copyDocument, copiedFormulaIds);
                AssertVisibleEquationNumbers(copyDocument, copiedFormulaIds, 1);
            }
            finally
            {
                Release(copiedShapeRange);
                Release(copiedShape);
                Release(copySourceRange);
                Release(copySourceTable);
                Release(sourceTableRange);
                Release(sourceTable);
                if (copyDocument is not null)
                {
                    try { copyDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                    Release(copyDocument);
                    copyDocument = null;
                    document.Activate();
                }
            }

            AssertNumberedFormulaArtifacts(document, formulaIds);
            AssertVisibleEquationNumbers(document, formulaIds, 1);
            Console.WriteLine(
                $"    [perf] final numbered OLE insert #{targetFormulaCount}: "
                + $"{insertTimings[targetFormulaCount - 1]}ms");
            if (insertTimings[targetFormulaCount - 1] > 1500)
                throw new InvalidDataException(
                    $"The final numbered OLE insertion #{targetFormulaCount} still took "
                    + $"{insertTimings[targetFormulaCount - 1]}ms.");

            var baseDocumentPath = Path.Combine(
                artifactRoot,
                "word-numbered-formula-performance.docx");
            document.SaveAs2(
                baseDocumentPath,
                Word.WdSaveFormat.wdFormatXMLDocument);
            if (targetFormulaCount == 100)
            {
                RunHundredFormulaStructuralPerformanceScenarios(
                    application,
                    baseDocumentPath,
                    formulaIds,
                    pngPath,
                    emfPath,
                    artifactRoot);
                document.Activate();
            }
            var skipOmmlScale = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_SKIP_OMML"),
                "1",
                StringComparison.Ordinal);
            if (targetFormulaCount >= 10 && !skipOmmlScale)
            {
                RunNumberedOmmlAppendScaleAcceptance(
                    application,
                    targetFormulaCount,
                    artifactRoot);
                document.Activate();
            }
            Console.WriteLine(
                $"Word numbered formula performance acceptance passed: {targetFormulaCount}-formula OLE/OMML append, "
                + "OLE edit, OMML edit and copied-table identity repair stayed on localized numbering paths.");
        }
        finally
        {
            if (copyDocument is not null)
            {
                try { copyDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                Release(copyDocument);
            }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            WordEquationNumbering.SetDefaultEquationNumberFormatPreference(originalNumberFormat);
            ForceComCleanup();
            foreach (var path in new[] { svgPath, pngPath, emfPath })
            {
                if (string.IsNullOrWhiteSpace(path)) continue;
                try { File.Delete(path!); } catch { }
            }
        }
    }

    private static void RunWordExistingNumberedPerformanceAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var baseDocumentPath = Environment.GetEnvironmentVariable(
            "VISUALTEX_NUMBERED_PERF_BASE_DOC");
        if (string.IsNullOrWhiteSpace(baseDocumentPath) || !File.Exists(baseDocumentPath))
            throw new FileNotFoundException(
                "VISUALTEX_NUMBERED_PERF_BASE_DOC must point to an existing numbered Word fixture.",
                baseDocumentPath);

        var scenarioPath = Path.Combine(
            artifactRoot,
            "word-numbered-existing-performance.docx");
        File.Copy(baseDocumentPath, scenarioPath, overwrite: true);
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
        Word.InlineShape? editShape = null;
        Word.Range? editRange = null;
        Word.Range? appendRange = null;
        try
        {
            File.WriteAllText(
                svgPath,
                CreateFontAcceptanceSvg("Times New Roman", "SimSun"),
                new UTF8Encoding(false));
            WriteAcceptancePng(pngPath, "x=1", 240, 72);
            emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 72);

            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(
                scenarioPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            var formulaIds = ReadNumberedFormulaIdsInDocumentOrder(document)
                .Select(id => Guid.TryParseExact(id, "N", out var parsedId)
                    ? parsedId.ToString("D")
                    : id)
                .ToArray();
            if (formulaIds.Length < 50)
                throw new InvalidDataException(
                    $"Existing-numbered performance fixture needs at least 50 formulas; found {formulaIds.Length}.");
            var service = new WordFormulaService(application);

            var editIndex = 49;
            editShape = FindNumberedOleByFormulaId(document, formulaIds[editIndex]);
            var originalMetadata = WordFormulaMetadataReader.TryRead(editShape)
                ?? throw new InvalidOperationException("Existing numbered OLE edit metadata is missing.");
            editRange = editShape.Range;
            var editSession = CreateNumberedPerformanceSession(
                "edit",
                formulaIds[editIndex],
                document.FullName,
                WordRangeReference(editRange.Start, editRange.End),
                originalMetadata,
                "x_{50}=5050");
            var editWatch = Stopwatch.StartNew();
            service.ReplaceOle(editSession, pngPath, emfPath);
            editWatch.Stop();
            Console.WriteLine(
                $"    [perf] existing numbered OLE edit #50 at {formulaIds.Length} formulas: {editWatch.ElapsedMilliseconds}ms");

            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            appendRange = application.Selection.Range;
            var appendedFormulaId = Guid.NewGuid().ToString("D");
            var appendSession = CreateNumberedPerformanceSession(
                "create",
                appendedFormulaId,
                document.FullName,
                WordRangeReference(appendRange.Start, appendRange.End),
                originalMetadata: null,
                latex: "x_{101}=101");
            var appendWatch = Stopwatch.StartNew();
            service.InsertOle(appendSession, pngPath, emfPath);
            appendWatch.Stop();
            Console.WriteLine(
                $"    [perf] existing numbered OLE append at {formulaIds.Length + 1} formulas: {appendWatch.ElapsedMilliseconds}ms");

            var selectedAfterAppend = service.ReadSelection();
            if (!string.Equals(
                    selectedAfterAppend.FormulaId,
                    appendedFormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "Existing numbered append did not keep the newly inserted formula selected.");

            var updateWatch = Stopwatch.StartNew();
            var updatedFormulaCount = service.UpdateEquationNumbers();
            updateWatch.Stop();
            Console.WriteLine(
                $"    [perf] explicit equation-number refresh at {formulaIds.Length + 1} formulas: {updateWatch.ElapsedMilliseconds}ms");
            if (updatedFormulaCount < formulaIds.Length)
                throw new InvalidDataException(
                    $"Explicit equation-number refresh lost formulas: expected at least {formulaIds.Length}, actual {updatedFormulaCount}.");

            var targetWatch = Stopwatch.StartNew();
            var targets = WordEquationNumbering.GetEquationReferenceTargets(document);
            targetWatch.Stop();
            Console.WriteLine(
                $"    [perf] equation reference target inventory at {targets.Count} formulas: {targetWatch.ElapsedMilliseconds}ms");
            if (targets.Count < formulaIds.Length)
                throw new InvalidDataException(
                    $"Equation reference target inventory lost formulas: expected at least {formulaIds.Length}, actual {targets.Count}.");

            var target = targets.FirstOrDefault(item => string.Equals(
                    item.FormulaId,
                    formulaIds[Math.Min(79, formulaIds.Length - 1)],
                    StringComparison.OrdinalIgnoreCase))
                ?? targets[Math.Min(79, targets.Count - 1)];

            // First force a real number change while the target has only its own
            // generated visible REF. The explicit refresh must stay fully local and
            // must not pay for a document-wide Fields.Update in this common case.
            CorruptNativeEquationNumberForAcceptance(document, target.FormulaId, 998);
            RefreshEquationReferencesForAcceptance(document, target.FormulaId);
            var localRepairWatch = Stopwatch.StartNew();
            service.UpdateEquationNumbers();
            localRepairWatch.Stop();
            var locallyRepairedTarget = WordEquationNumbering.GetEquationReferenceTargets(document)
                .First(item => string.Equals(
                    item.FormulaId,
                    target.FormulaId,
                    StringComparison.OrdinalIgnoreCase));
            AssertEqual(
                target.NumberText,
                locallyRepairedTarget.NumberText,
                "Local explicit equation-number refresh did not repair the native SEQ result.");
            AssertEquationReferencesForAcceptance(
                document,
                target.FormulaId,
                target.NumberText,
                minimumMatches: 1);
            Console.WriteLine(
                $"    [perf] changed equation-number repair without body REF at {targets.Count} formulas: {localRepairWatch.ElapsedMilliseconds}ms");

            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            application.Selection.TypeText("See ");
            var referenceWatch = Stopwatch.StartNew();
            WordEquationNumbering.InsertEquationReference(
                document,
                application.Selection,
                target,
                EquationReferenceStyle.Parenthesized);
            referenceWatch.Stop();
            Console.WriteLine(
                $"    [perf] equation reference insertion at {targets.Count} formulas: {referenceWatch.ElapsedMilliseconds}ms");

            // Prove that the fast explicit update is not merely a no-op shortcut.
            // Corrupt one healthy native SEQ result, refresh its visible/body REF
            // fields to the bad value, then require UpdateEquationNumbers to repair
            // the target and every generated reference without structural rebuild.
            CorruptNativeEquationNumberForAcceptance(document, target.FormulaId, 999);
            RefreshEquationReferencesForAcceptance(document, target.FormulaId);
            var corruptedTarget = WordEquationNumbering.GetEquationReferenceTargets(document)
                .First(item => string.Equals(
                    item.FormulaId,
                    target.FormulaId,
                    StringComparison.OrdinalIgnoreCase));
            if (string.Equals(
                    corruptedTarget.NumberText,
                    target.NumberText,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Equation-number corruption probe did not change the target number.");

            var repairWatch = Stopwatch.StartNew();
            service.UpdateEquationNumbers();
            repairWatch.Stop();
            var repairedTarget = WordEquationNumbering.GetEquationReferenceTargets(document)
                .First(item => string.Equals(
                    item.FormulaId,
                    target.FormulaId,
                    StringComparison.OrdinalIgnoreCase));
            AssertEqual(
                target.NumberText,
                repairedTarget.NumberText,
                "Explicit equation-number refresh did not repair a stale native SEQ result.");
            AssertEquationReferencesForAcceptance(
                document,
                target.FormulaId,
                target.NumberText);
            Console.WriteLine(
                $"    [perf] changed equation-number repair at {targets.Count} formulas: {repairWatch.ElapsedMilliseconds}ms");

            document.Save();
            Console.WriteLine("Word existing numbered daily-operation performance acceptance passed.");
        }
        finally
        {
            Release(appendRange);
            Release(editRange);
            Release(editShape);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            foreach (var path in new[] { svgPath, pngPath, emfPath })
            {
                if (string.IsNullOrWhiteSpace(path)) continue;
                try { File.Delete(path!); } catch { }
            }
        }
    }

    private static void RunWordNumberFormatPerformanceAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var baseDocumentPath = Environment.GetEnvironmentVariable(
            "VISUALTEX_NUMBERED_PERF_BASE_DOC");
        if (string.IsNullOrWhiteSpace(baseDocumentPath) || !File.Exists(baseDocumentPath))
            throw new FileNotFoundException(
                "VISUALTEX_NUMBERED_PERF_BASE_DOC must point to the saved 100-formula Word fixture.",
                baseDocumentPath);

        var scenarioPath = Path.Combine(
            artifactRoot,
            "word-number-format-performance.docx");
        File.Copy(baseDocumentPath, scenarioPath, overwrite: true);
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? markerRange = null;
        Word.Find? markerFind = null;
        Word.Paragraphs? markerParagraphs = null;
        Word.Paragraph? markerParagraph = null;
        Word.Range? headingRange = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(
                scenarioPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            var formulaIds = ReadNumberedFormulaIdsInDocumentOrder(document);
            AssertEqual(100, formulaIds.Count,
                "Number-format performance fixture must contain 100 numbered formulas.");

            markerRange = document.Content.Duplicate;
            markerFind = markerRange.Find;
            markerFind.ClearFormatting();
            markerFind.Text = NumberedMiddleInsertionAnchorText;
            markerFind.Forward = true;
            markerFind.Wrap = Word.WdFindWrap.wdFindStop;
            markerFind.Format = false;
            if (!markerFind.Execute())
                throw new InvalidOperationException(
                    "Number-format performance fixture is missing the preserved heading anchor.");
            if ((bool)markerRange.get_Information(Word.WdInformation.wdWithInTable)
                || markerRange.Frames.Count > 0)
                throw new InvalidOperationException(
                    "Number-format performance heading anchor is not ordinary body text.");

            markerParagraphs = markerRange.Paragraphs;
            markerParagraph = markerParagraphs[1];
            headingRange = markerParagraph.Range;
            headingRange.Text = "1. 性能验收章节\r";
            object headingStyle = Word.WdBuiltinStyle.wdStyleHeading1;
            headingRange.set_Style(ref headingStyle);

            void AssertNumbers(string separator)
            {
                for (var index = 0; index < formulaIds.Count; index++)
                {
                    var expected = index < 20
                        ? $"0{separator}{index + 1}"
                        : $"1{separator}{index - 19}";
                    var visible = ReadEquationNumberBookmarkText(
                        document,
                        WordEquationNumbering.EquationBookmarkName(formulaIds[index]));
                    var native = ReadEquationNumberBookmarkText(
                        document,
                        WordEquationNumbering.NativeNumberBookmarkName(formulaIds[index]));
                    AssertEqual(expected, visible,
                        $"Visible formula #{index + 1} is stale after heading-format switch.");
                    AssertEqual(expected, native,
                        $"Native formula #{index + 1} is stale after heading-format switch.");
                }
            }

            var dotWatch = Stopwatch.StartNew();
            var dotUpdated = WordEquationNumbering.SetEquationNumberFormat(
                document,
                EquationNumberFormat.Heading1DotId);
            dotWatch.Stop();
            AssertEqual(100, dotUpdated,
                "Heading-dot format switch did not update all 100 formulas.");
            AssertNumbers(".");
            Console.WriteLine(
                $"    [perf] 100-formula heading1-dot format switch: {dotWatch.ElapsedMilliseconds}ms");

            var dashWatch = Stopwatch.StartNew();
            var dashUpdated = WordEquationNumbering.SetEquationNumberFormat(
                document,
                EquationNumberFormat.Heading1DashId);
            dashWatch.Stop();
            AssertEqual(100, dashUpdated,
                "Heading-dash format switch did not update all 100 formulas.");
            AssertNumbers("-");
            Console.WriteLine(
                $"    [perf] 100-formula heading1-dash format switch: {dashWatch.ElapsedMilliseconds}ms");

            document.Save();
            Console.WriteLine("Word equation-number format performance acceptance passed.");
        }
        finally
        {
            Release(headingRange);
            Release(markerParagraph);
            Release(markerParagraphs);
            Release(markerFind);
            Release(markerRange);
            try { document?.Close(Word.WdSaveOptions.wdSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordNumberedStructuralPerformanceAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var baseDocumentPath = Environment.GetEnvironmentVariable(
            "VISUALTEX_NUMBERED_PERF_BASE_DOC");
        if (string.IsNullOrWhiteSpace(baseDocumentPath) || !File.Exists(baseDocumentPath))
            throw new FileNotFoundException(
                "VISUALTEX_NUMBERED_PERF_BASE_DOC must point to the saved 100-formula Word fixture.",
                baseDocumentPath);

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
        Word.Document? inventoryDocument = null;
        try
        {
            File.WriteAllText(
                svgPath,
                CreateFontAcceptanceSvg("Times New Roman", "SimSun"),
                new UTF8Encoding(false));
            WriteAcceptancePng(pngPath, "x=1", 240, 72);
            emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 72);
            application = CreateWordApplication(visible: false);
            inventoryDocument = application.Documents.Open(
                baseDocumentPath,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false);
            var formulaIds = ReadNumberedFormulaIdsInDocumentOrder(inventoryDocument);
            if (formulaIds.Count != 100)
                throw new InvalidDataException(
                    $"Structural performance base document must contain 100 numbered formulas; found {formulaIds.Count}.");
            inventoryDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(inventoryDocument);
            inventoryDocument = null;

            RunHundredFormulaStructuralPerformanceScenarios(
                application,
                baseDocumentPath,
                formulaIds,
                pngPath,
                emfPath,
                artifactRoot);
            Console.WriteLine("Word numbered structural performance acceptance passed.");
        }
        finally
        {
            try { inventoryDocument?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(inventoryDocument);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            foreach (var path in new[] { svgPath, pngPath, emfPath })
            {
                if (string.IsNullOrWhiteSpace(path)) continue;
                try { File.Delete(path!); } catch { }
            }
        }
    }

    private static IReadOnlyList<string> ReadNumberedFormulaIdsInDocumentOrder(
        Word.Document document)
    {
        var entries = new List<(string FormulaId, int Position)>();
        Word.Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Word.Bookmark? bookmark = null;
                Word.Range? range = null;
                try
                {
                    bookmark = bookmarks[index];
                    const string prefix = "VTEqNum_";
                    var name = bookmark.Name;
                    if (!name.StartsWith(prefix, StringComparison.Ordinal)) continue;
                    var formulaId = name.Substring(prefix.Length);
                    range = bookmark.Range;
                    entries.Add((formulaId, range.Start));
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }
        return entries
            .OrderBy(entry => entry.Position)
            .Select(entry => entry.FormulaId)
            .ToArray();
    }

    private static void RunHundredFormulaStructuralPerformanceScenarios(
        Word.Application application,
        string baseDocumentPath,
        IReadOnlyList<string> formulaIds,
        string pngPath,
        string emfPath,
        string artifactRoot)
    {
        RunHundredFormulaMiddleInsertPerformance(
            application,
            baseDocumentPath,
            formulaIds,
            pngPath,
            emfPath,
            artifactRoot);
        RunHundredFormulaCopyToEndPerformance(
            application,
            baseDocumentPath,
            formulaIds,
            artifactRoot);
    }

    private static void RunHundredFormulaMiddleInsertPerformance(
        Word.Application application,
        string baseDocumentPath,
        IReadOnlyList<string> formulaIds,
        string pngPath,
        string emfPath,
        string artifactRoot)
    {
        var scenarioPath = Path.Combine(
            artifactRoot,
            "word-numbered-middle-insert-performance.docx");
        File.Copy(baseDocumentPath, scenarioPath, overwrite: true);
        Word.Document? document = null;
        Word.Range? insertionRange = null;
        try
        {
            document = application.Documents.Open(
                scenarioPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            var service = new WordFormulaService(application);
            Word.Find? anchorFind = null;
            Word.Frames? anchorFrames = null;
            try
            {
                insertionRange = document.Content.Duplicate;
                anchorFind = insertionRange.Find;
                anchorFind.ClearFormatting();
                anchorFind.Text = NumberedMiddleInsertionAnchorText;
                anchorFind.Forward = true;
                anchorFind.Wrap = Word.WdFindWrap.wdFindStop;
                anchorFind.Format = false;
                if (!anchorFind.Execute())
                    throw new InvalidOperationException(
                        "Middle-insert acceptance could not find the preserved body paragraph.");
                var anchorStart = insertionRange.Start;
                var anchorEnd = insertionRange.End;
                var anchorInTable = (bool)insertionRange.get_Information(
                    Word.WdInformation.wdWithInTable);
                var anchorTableCount = insertionRange.Tables.Count;
                var anchorFrameCount = insertionRange.Frames.Count;
                Console.WriteLine(
                    $"    [diag] preserved middle anchor: {anchorStart}:{anchorEnd}; "
                    + $"inTable={anchorInTable}; tables={anchorTableCount}; frames={anchorFrameCount}");
                insertionRange.Text = string.Empty;
                insertionRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                if (anchorInTable)
                    throw new InvalidOperationException(
                        "Middle-insert acceptance anchor unexpectedly moved inside a table.");
                anchorFrames = insertionRange.Frames;
                if (anchorFrames.Count > 0)
                    throw new InvalidOperationException(
                        "Middle-insert acceptance anchor unexpectedly moved inside a caption frame.");
                application.Selection.SetRange(insertionRange.Start, insertionRange.End);
            }
            finally
            {
                Release(anchorFrames);
                Release(anchorFind);
            }
            var formulaId = Guid.NewGuid().ToString("D");
            var session = CreateNumberedPerformanceSession(
                "create",
                formulaId,
                document.FullName,
                WordRangeReference(insertionRange.Start, insertionRange.End),
                originalMetadata: null,
                latex: @"x_{20.5}=205");

            var watch = Stopwatch.StartNew();
            service.InsertOle(session, pngPath, emfPath);
            watch.Stop();
            Console.WriteLine(
                $"    [perf] numbered OLE insert between #20/#21 at 100 formulas: "
                + $"{watch.ElapsedMilliseconds}ms");

            var expectedFormulaIds = formulaIds.ToList();
            expectedFormulaIds.Insert(20, formulaId);
            AssertEqual(101, document.Tables.Count,
                "Middle insertion did not create the 101st numbered table.");
            AssertNumberedFormulaArtifacts(document, expectedFormulaIds);
            AssertVisibleEquationNumbers(document, expectedFormulaIds, 1);
            document.Save();
        }
        finally
        {
            Release(insertionRange);
            try { document?.Close(Word.WdSaveOptions.wdSaveChanges); } catch { }
            Release(document);
        }
    }

    private static void RunWordNumberedMiddleArtifactDump(string artifactRoot)
    {
        var path = Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_MIDDLE_ARTIFACT");
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            throw new FileNotFoundException(
                "Numbered middle-artifact dump requires VISUALTEX_NUMBERED_MIDDLE_ARTIFACT.",
                path);
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(
                Path.GetFullPath(path),
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false);
            Console.WriteLine(
                $"[MIDDLE DUMP] tables={document.Tables.Count} shapes={document.InlineShapes.Count} "
                + $"frames={document.Frames.Count} bookmarks={document.Bookmarks.Count}");
            for (var index = 1; index <= Math.Min(25, document.InlineShapes.Count); index++)
            {
                Word.InlineShape? shape = null;
                Word.Range? range = null;
                Word.Table? table = null;
                Word.Range? tableRange = null;
                try
                {
                    shape = document.InlineShapes[index];
                    range = shape.Range;
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    var tableStart = -1;
                    var tableEnd = -1;
                    if (range.Tables.Count > 0)
                    {
                        table = range.Tables[1];
                        tableRange = table.Range;
                        tableStart = tableRange.Start;
                        tableEnd = tableRange.End;
                    }
                    Console.WriteLine(
                        $"SHAPE|{index}|range={range.Start}:{range.End}|table={tableStart}:{tableEnd}|" +
                        $"formulaId={metadata?.FormulaId}|latex={metadata?.Latex}|numbered={metadata?.Numbered}");
                }
                finally
                {
                    Release(tableRange);
                    Release(table);
                    Release(range);
                    Release(shape);
                }
            }

            for (var index = 1; index <= document.Bookmarks.Count; index++)
            {
                Word.Bookmark? bookmark = null;
                Word.Range? range = null;
                try
                {
                    bookmark = document.Bookmarks[index];
                    var name = bookmark.Name ?? string.Empty;
                    if (!name.StartsWith("VTO_", StringComparison.Ordinal)
                        && !name.StartsWith("VTEq_", StringComparison.Ordinal)
                        && !name.StartsWith("VTEqCap_", StringComparison.Ordinal)
                        && !name.StartsWith("VTEqNum_", StringComparison.Ordinal))
                        continue;
                    range = bookmark.Range;
                    if (range.Start < 2500 || range.Start > 3400) continue;
                    Console.WriteLine(
                        $"BM|{name}|{range.Start}:{range.End}|text={(range.Text ?? string.Empty).Replace("\r", "<CR>").Replace("\a", "<CELL>")}");
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally
        {
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunHundredFormulaCopyToEndPerformance(
        Word.Application application,
        string baseDocumentPath,
        IReadOnlyList<string> formulaIds,
        string artifactRoot)
    {
        var scenarioPath = Path.Combine(
            artifactRoot,
            "word-numbered-copy-50-to-end-performance.docx");
        File.Copy(baseDocumentPath, scenarioPath, overwrite: true);
        Word.Document? document = null;
        Word.Table? sourceTable = null;
        Word.Range? sourceRange = null;
        Word.Table? copiedTable = null;
        Word.InlineShape? copiedShape = null;
        Word.Range? copiedShapeRange = null;
        try
        {
            document = application.Documents.Open(
                scenarioPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            var service = new WordFormulaService(application);
            sourceTable = document.Tables[50];
            sourceRange = sourceTable.Range.Duplicate;
            sourceRange.Copy();
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();

            var pasteWatch = Stopwatch.StartNew();
            application.Selection.Paste();
            pasteWatch.Stop();
            AssertEqual(101, document.Tables.Count,
                "Copy-to-end scenario did not paste the 101st numbered table.");

            copiedTable = document.Tables[101];
            copiedShape = copiedTable.Cell(1, 2).Range.InlineShapes[1];
            copiedShapeRange = copiedShape.Range;
            copiedShapeRange.Select();
            var repairWatch = Stopwatch.StartNew();
            var copiedSelection = service.ReadSelection();
            repairWatch.Stop();
            if (string.IsNullOrWhiteSpace(copiedSelection.FormulaId)
                || string.Equals(
                    copiedSelection.FormulaId,
                    formulaIds[49],
                    StringComparison.OrdinalIgnoreCase)
                || formulaIds.Any(id => string.Equals(
                    id,
                    copiedSelection.FormulaId,
                    StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException(
                    "Copied #50 formula did not receive an independent FormulaId.");

            Console.WriteLine(
                $"    [perf] copy #50 table to end at 100 formulas: paste={pasteWatch.ElapsedMilliseconds}ms; "
                + $"identity/number repair={repairWatch.ElapsedMilliseconds}ms");
            var expectedFormulaIds = formulaIds.ToList();
            expectedFormulaIds.Add(copiedSelection.FormulaId!);
            AssertNumberedFormulaArtifacts(document, expectedFormulaIds);
            AssertVisibleEquationNumbers(document, expectedFormulaIds, 1);
            document.Save();
        }
        finally
        {
            Release(copiedShapeRange);
            Release(copiedShape);
            Release(copiedTable);
            Release(sourceRange);
            Release(sourceTable);
            try { document?.Close(Word.WdSaveOptions.wdSaveChanges); } catch { }
            Release(document);
        }
    }

    private static void RunWordNumberedOmmlPerformanceAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        var originalNumberFormat = WordEquationNumbering.GetDefaultEquationNumberFormatId();
        try
        {
            var targetFormulaCount = int.TryParse(
                Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_COUNT"),
                out var parsedTargetFormulaCount)
                ? Math.Max(10, Math.Min(200, parsedTargetFormulaCount))
                : 100;
            application = CreateWordApplication(visible: false);
            RunNumberedOmmlAppendScaleAcceptance(
                application,
                targetFormulaCount,
                artifactRoot);
            Console.WriteLine(
                $"Word numbered OMML performance acceptance passed at {targetFormulaCount} formulas.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            WordEquationNumbering.SetDefaultEquationNumberFormatPreference(originalNumberFormat);
            ForceComCleanup();
        }
    }

    private static void RunNumberedOmmlAppendScaleAcceptance(
        Word.Application application,
        int targetFormulaCount,
        string artifactRoot)
    {
        Word.Document? document = null;
        var formulaIds = new List<string>();
        var timingCheckpoints = new HashSet<int>(new[] { 1, 10, 20, 40, 80, 100, 200 });
        try
        {
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(document, "continuous");
            var service = new WordFormulaService(application);
            var insertTimings = new List<long>(targetFormulaCount);
            for (var index = 1; index <= targetFormulaCount; index++)
            {
                application.Selection.EndKey(Word.WdUnits.wdStory);
                var range = application.Selection.Range;
                var formulaId = Guid.NewGuid().ToString("D");
                var session = CreateNumberedPerformanceSession(
                    "create",
                    formulaId,
                    document.FullName,
                    WordRangeReference(range.Start, range.End),
                    originalMetadata: null,
                    latex: $"y_{{{index}}}={index}");
                session.ObjectMode = FormulaOleContract.WordOmmlMode;
                Release(range);

                var watch = Stopwatch.StartNew();
                service.InsertOmml(session, PerformanceMathMl(index, 0));
                watch.Stop();
                insertTimings.Add(watch.ElapsedMilliseconds);
                formulaIds.Add(formulaId);
                if (timingCheckpoints.Contains(index) || index == targetFormulaCount)
                {
                    Console.WriteLine(
                        $"    [perf] numbered OMML append #{index}: {watch.ElapsedMilliseconds}ms");
                }
            }

            AssertEqual(targetFormulaCount, document.Tables.Count,
                $"Numbered OMML scale fixture did not create {targetFormulaCount} equation tables.");
            AssertNumberedFormulaArtifacts(document, formulaIds);
            AssertVisibleEquationNumbers(document, formulaIds, 1);

            var editIndex = targetFormulaCount >= 50 ? 49 : targetFormulaCount / 2;
            Word.Bookmark? bookmark = null;
            Word.Range? equationRange = null;
            try
            {
                var formulaId = formulaIds[editIndex];
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                    ?? throw new InvalidOperationException(
                        $"Numbered OMML scale edit bookmark {editIndex + 1} is missing.");
                var metadata = WordOmmlFormulaStore.TryRead(document, bookmark)
                    ?? throw new InvalidOperationException(
                        $"Numbered OMML scale edit metadata {editIndex + 1} is missing.");
                equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                var editSession = CreateNumberedPerformanceSession(
                    "edit",
                    formulaId,
                    document.FullName,
                    WordRangeReference(equationRange.Start, equationRange.End),
                    metadata,
                    $"y_{{{editIndex + 1}}}=999");
                editSession.ObjectMode = FormulaOleContract.WordOmmlMode;
                var watch = Stopwatch.StartNew();
                service.ReplaceOmml(editSession, PerformanceMathMl(editIndex + 1, 9));
                watch.Stop();
                Console.WriteLine(
                    $"    [perf] numbered OMML edit #{editIndex + 1} at {targetFormulaCount} formulas: "
                    + $"{watch.ElapsedMilliseconds}ms");
                if (watch.ElapsedMilliseconds > 1500)
                    throw new InvalidDataException(
                        $"Numbered OMML edit #{editIndex + 1} took {watch.ElapsedMilliseconds}ms "
                        + $"at {targetFormulaCount} formulas.");
            }
            finally
            {
                Release(equationRange);
                Release(bookmark);
            }

            if (insertTimings[targetFormulaCount - 1] > 1500)
                throw new InvalidDataException(
                    $"The final numbered OMML insertion #{targetFormulaCount} still took "
                    + $"{insertTimings[targetFormulaCount - 1]}ms.");
            document.SaveAs2(
                Path.Combine(artifactRoot, "word-numbered-omml-scale-performance.docx"),
                Word.WdSaveFormat.wdFormatXMLDocument);
        }
        finally
        {
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
        }
    }

    private static OfficeSessionDocument CreateNumberedPerformanceSession(
        string mode,
        string formulaId,
        string sourceDocumentId,
        string sourceObjectId,
        FormulaMetadata? originalMetadata,
        string latex)
    {
        var session = CreateOleFontSession(
            "word",
            mode,
            formulaId,
            sourceDocumentId,
            sourceObjectId,
            originalMetadata,
            "times",
            "songti");
        session.Numbered = true;
        session.Title = "Numbered formula performance acceptance";
        session.Lines = new List<FormulaLine>
        {
            new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
        };
        return session;
    }

    private static string PerformanceMathMl(int index, int round) =>
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
        + $"<msub><mi>y</mi><mn>{index}</mn></msub><mo>=</mo><mn>{index * 10 + round}</mn>"
        + "</math>";

    private static Word.InlineShape FindNumberedOleByFormulaId(
        Word.Document document,
        string formulaId)
    {
        var shapes = document.InlineShapes;
        try
        {
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata?.Numbered == true
                        && string.Equals(
                            metadata.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        var result = shape;
                        shape = null;
                        return result;
                    }
                }
                finally { Release(shape); }
            }
            throw new InvalidOperationException($"Numbered OLE formula {formulaId} was not found.");
        }
        finally { Release(shapes); }
    }

    private static void CorruptNativeEquationNumberForAcceptance(
        Word.Document document,
        string formulaId,
        int forcedOrdinal)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? captionBookmark = null;
        Word.Range? captionRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            bookmarks = document.Bookmarks;
            captionBookmark = bookmarks[WordEquationNumbering.NativeCaptionBookmarkName(formulaId)];
            captionRange = captionBookmark.Range;
            fields = captionRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(field);
                field = fields[index];
                Release(code);
                code = field.Code;
                var codeText = code.Text ?? string.Empty;
                if (codeText.IndexOf("SEQ ", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                var replaced = System.Text.RegularExpressions.Regex.Replace(
                    codeText,
                    @"\\r\s+\d+",
                    $@"\r {forcedOrdinal}",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                    | System.Text.RegularExpressions.RegexOptions.CultureInvariant);
                if (string.Equals(replaced, codeText, StringComparison.Ordinal))
                    replaced = codeText + $@" \r {forcedOrdinal} ";
                code.Text = replaced;
                field.Update();
                return;
            }
            throw new InvalidDataException(
                $"Native equation caption for {formulaId} has no SEQ field to corrupt.");
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    private static void RefreshEquationReferencesForAcceptance(
        Word.Document document,
        string formulaId)
    {
        var targetBookmark = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(field);
                field = fields[index];
                Release(code);
                code = field.Code;
                var codeText = code.Text ?? string.Empty;
                if (codeText.IndexOf("REF", StringComparison.OrdinalIgnoreCase) < 0
                    || codeText.IndexOf(targetBookmark, StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                field.Update();
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static void AssertEquationReferencesForAcceptance(
        Word.Document document,
        string formulaId,
        string expectedNumber,
        int minimumMatches = 2)
    {
        var targetBookmark = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? result = null;
        var matched = 0;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(result);
                result = null;
                Release(field);
                field = fields[index];
                Release(code);
                code = field.Code;
                var codeText = code.Text ?? string.Empty;
                if (codeText.IndexOf("REF", StringComparison.OrdinalIgnoreCase) < 0
                    || codeText.IndexOf(targetBookmark, StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                result = field.Result;
                AssertEqual(
                    expectedNumber,
                    (result.Text ?? string.Empty).Trim(),
                    "Equation REF result remained stale after explicit number refresh.");
                matched++;
            }
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
        }
        if (matched < minimumMatches)
            throw new InvalidDataException(
                $"Expected at least {minimumMatches} REF fields for {formulaId}; found {matched}.");
    }

    private static void AssertNumberedFormulaArtifacts(
        Word.Document document,
        IReadOnlyList<string> formulaIds)
    {
        foreach (var formulaId in formulaIds)
        {
            AssertTrue(
                WordEquationNumbering.HasCompleteFormulaNumberingArtifacts(document, formulaId),
                $"Numbered formula {formulaId} lost its numbering artifacts.");
        }
    }

    private static void AssertVisibleEquationNumbers(
        Word.Document document,
        IReadOnlyList<string> formulaIds,
        int firstNumber)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 0; index < formulaIds.Count; index++)
            {
                var bookmarkName = WordEquationNumbering.EquationBookmarkName(formulaIds[index]);
                if (!bookmarks.Exists(bookmarkName))
                    throw new InvalidOperationException($"Visible number bookmark {bookmarkName} is missing.");
                Release(bookmark);
                bookmark = bookmarks[bookmarkName];
                var text = (bookmark.Range.Text ?? string.Empty).Trim();
                AssertEqual(
                    $"({firstNumber + index})",
                    text,
                    $"Visible numbered formula {index + 1} is stale.");
            }
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }
}
