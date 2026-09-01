using System.Diagnostics;
using System.Text;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordSparseNumberedOmmlPerformanceAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var scenarioPath = Path.Combine(
            artifactRoot,
            "word-sparse-numbered-omml-performance.docx");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Bookmark? targetBookmark = null;
        Word.Range? targetRange = null;
        Word.Bookmark? numberedBookmark = null;
        Word.Range? numberedRange = null;
        Word.Tables? numberedTables = null;
        Word.Table? numberedTable = null;
        Word.Range? numberedTableRange = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Activate();
            var formulaIds = PopulateSparseNumberedOmmlFixture(
                application,
                document,
                formulaCount: 100);
            var service = new WordFormulaService(application);

            // Model the user's real document: one hundred VisualTeX OMML objects,
            // but only four direct-SEQ numbered hosts. Fixture creation is outside
            // every measured interval.
            foreach (var index in new[] { 9, 29, 69, 89 })
                AddNumberToSparseFixtureFormula(
                    service,
                    document,
                    formulaIds[index]);
            document.SaveAs2(
                scenarioPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);

            var initialMathCount = document.OMaths.Count;
            var initialTableCount = document.Tables.Count;
            var initialBlankBeforeTables =
                CountPureBlankParagraphsImmediatelyBeforeTables(document);
            AssertEqual(100, initialMathCount,
                "Sparse-numbering performance fixture did not retain 100 OMML formulas.");
            AssertEqual(4, initialTableCount,
                "Sparse-numbering performance fixture did not create four numbered hosts.");

            // The hundred-formula fixture intentionally keeps one real inline OMath
            // and ninety-nine displays so all three Apply hot paths are measured in
            // the same document scale the user reported. These are content edits,
            // not create-time measurements.
            var inlineEditMilliseconds = MeasureSparseOmmlContentEdit(
                service,
                document,
                formulaIds[99],
                displayMode: "inline",
                numbered: false,
                editedMathMl: QuadraticFormulaMathMl()
                    .Replace(" display=\"block\"", string.Empty)
                    .Replace("</mrow></math>", "<mo>+</mo><mn>7</mn></mrow></math>"));
            AssertTrue(inlineEditMilliseconds < 700,
                $"Inline OMML content edit exceeded the 700ms target: {inlineEditMilliseconds}ms.");

            var unnumberedDisplayEditMilliseconds = MeasureSparseOmmlContentEdit(
                service,
                document,
                formulaIds[49],
                displayMode: "block",
                numbered: false,
                editedMathMl: QuadraticFormulaMathMl()
                    .Replace("</mrow></math>", "<mo>+</mo><mn>8</mn></mrow></math>"));
            AssertTrue(unnumberedDisplayEditMilliseconds < 700,
                $"Unnumbered display OMML content edit exceeded the 700ms target: {unnumberedDisplayEditMilliseconds}ms.");

            var middle = document.Content.End / 2;
            var candidates = new List<(
                string FormulaId,
                FormulaMetadata Metadata,
                int Start,
                int End)>();
            foreach (var formulaId in WordOmmlFormulaStore.FormulaIds(document))
            {
                Word.Bookmark? bookmark = null;
                Word.Range? range = null;
                try
                {
                    var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                    if (metadata is null
                        || metadata.Numbered
                        || !string.Equals(
                            metadata.DisplayMode,
                            "block",
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                    if (bookmark is null) continue;
                    range = WordOmmlFormulaStore.GetEquationRange(bookmark);
                    if ((bool)range.get_Information(Word.WdInformation.wdWithInTable)
                        || range.OMaths.Count != 1)
                        continue;
                    candidates.Add((formulaId, metadata, range.Start, range.End));
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
            var target = candidates
                .OrderBy(candidate => Math.Abs(candidate.Start - middle))
                .FirstOrDefault();
            if (string.IsNullOrWhiteSpace(target.FormulaId))
                throw new InvalidDataException(
                    "Sparse-numbering performance fixture has no standalone unnumbered display OMML formula.");

            targetBookmark = WordOmmlFormulaStore.FindByFormulaId(
                document,
                target.FormulaId)
                ?? throw new InvalidDataException("Target OMML bookmark is missing.");
            targetRange = WordOmmlFormulaStore.GetEquationRange(targetBookmark);
            application.Selection.SetRange(targetRange.Start, targetRange.End);
            try
            {
                object start = true;
                application.ActiveWindow.ScrollIntoView(targetRange, ref start);
            }
            catch { }
            var verticalBefore = application.ActiveWindow.VerticalPercentScrolled;

            var session = CreateNumberedOmmlTabSession(
                target.FormulaId,
                document.FullName,
                targetRange.Start,
                targetRange.End,
                latex: target.Metadata.Latex,
                originalMetadata: target.Metadata);
            session.FontSizePt = target.Metadata.FontSizePt
                ?? FormulaFontSize.DefaultPt;
            if (session.ExportResult is not null)
            {
                session.ExportResult.FormulaLetterFont =
                    target.Metadata.FormulaLetterFont ?? "katex";
                session.ExportResult.FormulaChineseFont =
                    target.Metadata.FormulaChineseFont ?? "system";
            }

            Release(targetRange);
            targetRange = null;
            Release(targetBookmark);
            targetBookmark = null;

            var addNumberWatch = Stopwatch.StartNew();
            service.ReplaceOmml(session, QuadraticFormulaMathMl());
            addNumberWatch.Stop();

            numberedBookmark = WordOmmlFormulaStore.FindByFormulaId(
                document,
                target.FormulaId)
                ?? throw new InvalidDataException(
                    "Unnumbered→numbered edit lost the target OMML bookmark.");
            numberedRange = WordOmmlFormulaStore.GetEquationRange(numberedBookmark);
            numberedTables = numberedRange.Tables;
            if (numberedTables.Count != 1)
                throw new InvalidDataException(
                    "Unnumbered→numbered edit did not create exactly one direct 1x3 host.");
            numberedTable = numberedTables[1];
            numberedTableRange = numberedTable.Range.Duplicate;

            AssertEqual(initialMathCount, document.OMaths.Count,
                "Adding a number changed the total OMML formula count.");
            AssertEqual(initialTableCount + 1, document.Tables.Count,
                "Adding a number did not create exactly one 1x3 layout table.");
            var blanksAfterAdd = CountPureBlankParagraphsImmediatelyBeforeTables(document);
            if (blanksAfterAdd != initialBlankBeforeTables)
                TraceSparseBlankParagraphsBeforeTables(document);
            AssertEqual(
                initialBlankBeforeTables,
                blanksAfterAdd,
                "Unnumbered→numbered OMML introduced a plain blank paragraph immediately above its 1x3 host.");
            AssertOmmlTabNumberingHost(
                document,
                target.FormulaId,
                "sparse-numbering unnumbered→numbered host",
                updateReference: true);

            var verticalAfter = application.ActiveWindow.VerticalPercentScrolled;
            var selectionInsideTarget = application.Selection.Start >= numberedTableRange.Start
                && application.Selection.End <= numberedTableRange.End;
            AssertTrue(selectionInsideTarget,
                "Adding a number did not restore the user's selection to the edited formula.");
            if (verticalBefore >= 10)
                AssertTrue(verticalAfter >= 5,
                    $"Adding a number jumped the Word viewport to the document top: before={verticalBefore}, after={verticalAfter}.");
            AssertTrue(addNumberWatch.ElapsedMilliseconds < 4500,
                $"Sparse 100-formula unnumbered→numbered edit remained too slow: {addNumberWatch.ElapsedMilliseconds}ms.");

            // Measure the user's two hot paths independently of initial fixture
            // construction: editing an already-numbered 1x3 OMML formula must keep
            // the existing direct-SEQ host, while toggling numbering is allowed to
            // dismantle/recreate exactly this one table.
            var editedMathMl = QuadraticFormulaMathMl().Replace(
                "</mrow></math>",
                "<mo>+</mo><mn>1</mn></mrow></math>");
            var numberedMetadata = WordOmmlFormulaStore.TryRead(
                    document,
                    target.FormulaId)
                ?? throw new InvalidDataException(
                    "Numbered-edit performance target lost its metadata.");
            Release(numberedTableRange); numberedTableRange = null;
            Release(numberedTable); numberedTable = null;
            Release(numberedTables); numberedTables = null;
            Release(numberedRange); numberedRange = null;
            Release(numberedBookmark); numberedBookmark = null;
            numberedBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    target.FormulaId)
                ?? throw new InvalidDataException(
                    "Numbered-edit performance target lost its formula bookmark.");
            numberedRange = WordOmmlFormulaStore.GetEquationRange(numberedBookmark);
            var numberedEditSession = CreateNumberedOmmlTabSession(
                target.FormulaId,
                document.FullName,
                numberedRange.Start,
                numberedRange.End,
                latex: numberedMetadata.Latex + "+1",
                originalMetadata: numberedMetadata);
            numberedEditSession.FontSizePt = numberedMetadata.FontSizePt
                ?? FormulaFontSize.DefaultPt;
            if (numberedEditSession.ExportResult is not null)
            {
                numberedEditSession.ExportResult.FormulaLetterFont =
                    numberedMetadata.FormulaLetterFont ?? "katex";
                numberedEditSession.ExportResult.FormulaChineseFont =
                    numberedMetadata.FormulaChineseFont ?? "system";
            }
            var numberedEditWatch = Stopwatch.StartNew();
            service.ReplaceOmml(numberedEditSession, editedMathMl);
            numberedEditWatch.Stop();
            AssertEqual(initialTableCount + 1, document.Tables.Count,
                "Editing a numbered OMML formula rebuilt or removed its 1x3 host.");
            AssertOmmlTabNumberingHost(
                document,
                target.FormulaId,
                "sparse-numbering numbered OMML content edit",
                updateReference: false);
            AssertTrue(numberedEditWatch.ElapsedMilliseconds < 800,
                $"Numbered OMML content edit exceeded the 800ms target: {numberedEditWatch.ElapsedMilliseconds}ms.");

            var editedNumberedMetadata = WordOmmlFormulaStore.TryRead(
                    document,
                    target.FormulaId)
                ?? throw new InvalidDataException(
                    "Numbered→unnumbered performance target lost edited metadata.");
            Release(numberedRange); numberedRange = null;
            Release(numberedBookmark); numberedBookmark = null;
            numberedBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    target.FormulaId)
                ?? throw new InvalidDataException(
                    "Numbered→unnumbered performance target lost its bookmark.");
            numberedRange = WordOmmlFormulaStore.GetEquationRange(numberedBookmark);
            var unnumberSession = CreateNumberedOmmlTabSession(
                target.FormulaId,
                document.FullName,
                numberedRange.Start,
                numberedRange.End,
                latex: editedNumberedMetadata.Latex,
                originalMetadata: editedNumberedMetadata);
            unnumberSession.Numbered = false;
            unnumberSession.FontSizePt = editedNumberedMetadata.FontSizePt
                ?? FormulaFontSize.DefaultPt;
            if (unnumberSession.ExportResult is not null)
            {
                unnumberSession.ExportResult.FormulaLetterFont =
                    editedNumberedMetadata.FormulaLetterFont ?? "katex";
                unnumberSession.ExportResult.FormulaChineseFont =
                    editedNumberedMetadata.FormulaChineseFont ?? "system";
            }
            var removeNumberWatch = Stopwatch.StartNew();
            service.ReplaceOmml(unnumberSession, editedMathMl);
            removeNumberWatch.Stop();
            AssertEqual(initialTableCount, document.Tables.Count,
                "Numbered→unnumbered did not dismantle exactly one 1x3 host.");

            var unnumberedMetadata = WordOmmlFormulaStore.TryRead(
                    document,
                    target.FormulaId)
                ?? throw new InvalidDataException(
                    "Unnumbered→numbered performance target lost metadata.");
            AssertTrue(!unnumberedMetadata.Numbered,
                "Numbered→unnumbered performance target still reports Numbered=true.");
            Release(numberedRange); numberedRange = null;
            Release(numberedBookmark); numberedBookmark = null;
            numberedBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    target.FormulaId)
                ?? throw new InvalidDataException(
                    "Unnumbered→numbered performance target lost its bookmark.");
            numberedRange = WordOmmlFormulaStore.GetEquationRange(numberedBookmark);
            AssertTrue(!(bool)numberedRange.get_Information(Word.WdInformation.wdWithInTable),
                "Numbered→unnumbered performance target remained inside a table.");
            var renumberSession = CreateNumberedOmmlTabSession(
                target.FormulaId,
                document.FullName,
                numberedRange.Start,
                numberedRange.End,
                latex: unnumberedMetadata.Latex,
                originalMetadata: unnumberedMetadata);
            renumberSession.FontSizePt = unnumberedMetadata.FontSizePt
                ?? FormulaFontSize.DefaultPt;
            if (renumberSession.ExportResult is not null)
            {
                renumberSession.ExportResult.FormulaLetterFont =
                    unnumberedMetadata.FormulaLetterFont ?? "katex";
                renumberSession.ExportResult.FormulaChineseFont =
                    unnumberedMetadata.FormulaChineseFont ?? "system";
            }
            var restoreNumberWatch = Stopwatch.StartNew();
            service.ReplaceOmml(renumberSession, editedMathMl);
            restoreNumberWatch.Stop();
            AssertEqual(initialTableCount + 1, document.Tables.Count,
                "Unnumbered→numbered did not restore exactly one 1x3 host.");
            AssertOmmlTabNumberingHost(
                document,
                target.FormulaId,
                "sparse-numbering restored numbered host",
                updateReference: false);
            AssertTrue(removeNumberWatch.ElapsedMilliseconds < 5000,
                $"Numbered→unnumbered toggle remained too slow: {removeNumberWatch.ElapsedMilliseconds}ms.");
            AssertTrue(restoreNumberWatch.ElapsedMilliseconds < 5000,
                $"Unnumbered→numbered toggle remained too slow: {restoreNumberWatch.ElapsedMilliseconds}ms.");

            var updateWatch = Stopwatch.StartNew();
            var updated = service.UpdateEquationNumbers();
            updateWatch.Stop();
            AssertTrue(updated >= initialTableCount + 1,
                $"Update Numbers returned too few VisualTeX formulas: {updated}.");
            AssertTrue(updateWatch.ElapsedMilliseconds < 1000,
                $"Sparse-numbering Update Numbers remained too slow: {updateWatch.ElapsedMilliseconds}ms.");

            var currentFormat = service.GetEquationNumberFormatId();
            var nextFormat = string.Equals(
                    currentFormat,
                    EquationNumberFormat.Heading1DashId,
                    StringComparison.Ordinal)
                ? EquationNumberFormat.Heading1DotId
                : EquationNumberFormat.Heading1DashId;
            var formatWatch = Stopwatch.StartNew();
            var formatted = service.SetEquationNumberFormat(nextFormat);
            formatWatch.Stop();
            AssertTrue(formatted >= initialTableCount + 1,
                $"Set Equation Number Format returned too few formulas: {formatted}.");
            AssertTrue(formatWatch.ElapsedMilliseconds < 1500,
                $"Sparse-numbering format switch remained too slow: {formatWatch.ElapsedMilliseconds}ms.");
            AssertEqual(
                initialBlankBeforeTables,
                CountPureBlankParagraphsImmediatelyBeforeTables(document),
                "Number update/format switching introduced a blank paragraph before a numbered OMML table.");

            document.Save();
            Console.WriteLine(
                $"Sparse numbered OMML performance passed: totalOMML={initialMathCount}, numbered={initialTableCount + 1}, "
                + $"inlineEdit={inlineEditMilliseconds}ms, unnumberedDisplayEdit={unnumberedDisplayEditMilliseconds}ms, "
                + $"addNumber={addNumberWatch.ElapsedMilliseconds}ms, numberedEdit={numberedEditWatch.ElapsedMilliseconds}ms, "
                + $"removeNumber={removeNumberWatch.ElapsedMilliseconds}ms, restoreNumber={restoreNumberWatch.ElapsedMilliseconds}ms, "
                + $"update={updateWatch.ElapsedMilliseconds}ms, format={formatWatch.ElapsedMilliseconds}ms, "
                + $"scroll={verticalBefore}->{verticalAfter}.");
        }
        finally
        {
            Release(numberedTableRange);
            Release(numberedTable);
            Release(numberedTables);
            Release(numberedRange);
            Release(numberedBookmark);
            Release(targetRange);
            Release(targetBookmark);
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

    private static IReadOnlyList<string> PopulateSparseNumberedOmmlFixture(
        Word.Application application,
        Word.Document document,
        int formulaCount)
    {
        var semanticOmml = WordOmmlConverter.ExtractSingleOMath(
            WordOmmlConverter.TransformMathMlToOmml(QuadraticFormulaMathMl()));
        var body = new StringBuilder(formulaCount * (semanticOmml.Length + 240));
        for (var index = 0; index < formulaCount; index++)
        {
            if (index == formulaCount - 1)
            {
                body.Append("<w:p><w:r><w:t xml:space=\"preserve\">Inline context: </w:t></w:r>")
                    .Append(semanticOmml)
                    .Append("<w:r><w:t xml:space=\"preserve\"> end.</w:t></w:r></w:p>");
                continue;
            }
            body.Append("<w:p><w:r><w:t xml:space=\"preserve\">Context paragraph ")
                .Append(index + 1)
                .Append(" for the sparse numbering performance fixture.</w:t></w:r></w:p>");
            body.Append("<w:p><m:oMathPara><m:oMathParaPr><m:jc m:val=\"center\" />")
                .Append("</m:oMathParaPr>")
                .Append(semanticOmml)
                .Append("</m:oMathPara></w:p>");
        }
        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
            + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
            + "<w:body>" + body + "</w:body></w:document>";

        WordOmmlConverter.WholeDocumentSource? source = null;
        Word.Range? anchor = null;
        Word.Range? inserted = null;
        Word.OMaths? maths = null;
        try
        {
            source = WordOmmlConverter.CreateWholeDocumentSource(
                application,
                documentXml);
            anchor = document.Range(document.Content.Start, document.Content.Start);
            inserted = source.Insert(document, anchor);
            maths = inserted.OMaths;
            AssertEqual(formulaCount, maths.Count,
                "Synthetic sparse-numbering source did not import every OMML formula.");

            var ids = new List<string>(formulaCount);
            var metadataItems = new List<FormulaMetadata>(formulaCount);
            for (var index = 1; index <= maths.Count; index++)
            {
                Word.OMath? math = null;
                Word.Range? range = null;
                Word.Bookmark? bookmark = null;
                try
                {
                    math = maths[index];
                    var isInline = index == maths.Count;
                    math.Type = isInline
                        ? Word.WdOMathType.wdOMathInline
                        : Word.WdOMathType.wdOMathDisplay;
                    range = math.Range.Duplicate;
                    var formulaId = Guid.NewGuid().ToString("D");
                    var latex = $@"x_{{{index}}}=\frac{{-b\pm\sqrt{{b^2-4ac}}}}{{2a}}";
                    var metadata = new FormulaMetadata
                    {
                        FormulaId = formulaId,
                        Title = "Sparse numbered OMML performance fixture",
                        Latex = latex,
                        Lines = new List<FormulaLine>
                        {
                            new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
                        },
                        CodeFormat = "latex",
                        DisplayMode = isInline ? "inline" : "block",
                        Numbered = false,
                        FontSizePt = 10.5,
                        FormulaLetterFont = "katex",
                        FormulaChineseFont = "system",
                        CreatedWithVersion = "1.2.5",
                        UpdatedWithVersion = "1.2.5",
                        CreatedAt = DateTimeOffset.UtcNow.ToString("O"),
                        UpdatedAt = DateTimeOffset.UtcNow.ToString("O"),
                    };
                    WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                        metadata,
                        range);
                    metadata.Validate();
                    bookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        range,
                        metadata,
                        replaceExisting: false);
                    metadataItems.Add(metadata);
                    ids.Add(formulaId);
                }
                finally
                {
                    Release(bookmark);
                    Release(range);
                    Release(math);
                }
            }
            WordOmmlFormulaStore.SaveNewBatch(document, metadataItems);
            return ids;
        }
        finally
        {
            Release(maths);
            Release(inserted);
            Release(anchor);
            source?.Dispose();
        }
    }

    private static void TraceSparseBlankParagraphsBeforeTables(Word.Document document)
    {
        for (var index = 1; index <= document.Tables.Count; index++)
        {
            Word.Table? table = null;
            Word.Range? tableRange = null;
            Word.Range? probe = null;
            Word.Paragraphs? paragraphs = null;
            Word.Paragraph? paragraph = null;
            Word.Range? paragraphRange = null;
            Word.Bookmarks? bookmarks = null;
            try
            {
                table = document.Tables[index];
                tableRange = table.Range;
                if (tableRange.Start <= document.Content.Start) continue;
                probe = document.Range(tableRange.Start - 1, tableRange.Start);
                if ((bool)probe.get_Information(Word.WdInformation.wdWithInTable)) continue;
                paragraphs = probe.Paragraphs;
                if (paragraphs.Count != 1) continue;
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
                if (paragraphRange.End != tableRange.Start
                    || !string.Equals(paragraphRange.Text, "\r", StringComparison.Ordinal))
                    continue;
                if (paragraphRange.Tables.Count != 0
                    || paragraphRange.InlineShapes.Count != 0
                    || paragraphRange.OMaths.Count != 0
                    || paragraphRange.Fields.Count != 0
                    || paragraphRange.Frames.Count != 0)
                    continue;
                bookmarks = paragraphRange.Bookmarks;
                var names = new List<string>();
                for (var bookmarkIndex = 1; bookmarkIndex <= bookmarks.Count; bookmarkIndex++)
                {
                    Word.Bookmark? bookmark = null;
                    try
                    {
                        bookmark = bookmarks[bookmarkIndex];
                        names.Add(bookmark.Name ?? string.Empty);
                    }
                    finally { Release(bookmark); }
                }
                Console.WriteLine(
                    $"    [blank-before-table] table={index} range={paragraphRange.Start}:{paragraphRange.End} table={tableRange.Start}:{tableRange.End} bookmarks=[{string.Join(",", names)}]");
            }
            finally
            {
                Release(bookmarks);
                Release(paragraphRange);
                Release(paragraph);
                Release(paragraphs);
                Release(probe);
                Release(tableRange);
                Release(table);
            }
        }
    }

    private static long MeasureSparseOmmlContentEdit(
        WordFormulaService service,
        Word.Document document,
        string formulaId,
        string displayMode,
        bool numbered,
        string editedMathMl)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        Word.Bookmark? updatedBookmark = null;
        Word.Range? updatedRange = null;
        Word.OMaths? updatedMaths = null;
        Word.OMath? updatedMath = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "Synthetic OMML performance formula metadata is missing.");
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    "Synthetic OMML performance formula bookmark is missing.");
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                range.Start,
                range.End,
                latex: metadata.Latex + "+1",
                originalMetadata: metadata);
            session.DisplayMode = displayMode;
            session.Numbered = numbered;
            session.FontSizePt = metadata.FontSizePt ?? FormulaFontSize.DefaultPt;
            if (session.ExportResult is not null)
            {
                session.ExportResult.FormulaLetterFont =
                    metadata.FormulaLetterFont ?? "katex";
                session.ExportResult.FormulaChineseFont =
                    metadata.FormulaChineseFont ?? "system";
            }

            var watch = Stopwatch.StartNew();
            service.ReplaceOmml(session, editedMathMl);
            watch.Stop();

            updatedBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    "OMML content edit lost its formula bookmark.");
            updatedRange = WordOmmlFormulaStore.GetEquationRange(updatedBookmark);
            AssertTrue(!(bool)updatedRange.get_Information(Word.WdInformation.wdWithInTable),
                $"{displayMode} unnumbered OMML content edit unexpectedly entered a table.");
            updatedMaths = updatedRange.OMaths;
            AssertEqual(1, updatedMaths.Count,
                $"{displayMode} OMML content edit no longer contains exactly one OMath.");
            updatedMath = updatedMaths[1];
            AssertEqual(
                string.Equals(displayMode, "inline", StringComparison.Ordinal)
                    ? Word.WdOMathType.wdOMathInline
                    : Word.WdOMathType.wdOMathDisplay,
                updatedMath.Type,
                $"{displayMode} OMML content edit changed the native OMath type.");
            return watch.ElapsedMilliseconds;
        }
        finally
        {
            Release(updatedMath);
            Release(updatedMaths);
            Release(updatedRange);
            Release(updatedBookmark);
            Release(range);
            Release(bookmark);
        }
    }

    private static void AddNumberToSparseFixtureFormula(
        WordFormulaService service,
        Word.Document document,
        string formulaId)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "Synthetic sparse-numbering formula metadata is missing.");
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    "Synthetic sparse-numbering formula bookmark is missing.");
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                range.Start,
                range.End,
                latex: metadata.Latex,
                originalMetadata: metadata);
            session.FontSizePt = metadata.FontSizePt ?? FormulaFontSize.DefaultPt;
            service.ReplaceOmml(session, QuadraticFormulaMathMl());
        }
        finally
        {
            Release(range);
            Release(bookmark);
        }
    }
}
