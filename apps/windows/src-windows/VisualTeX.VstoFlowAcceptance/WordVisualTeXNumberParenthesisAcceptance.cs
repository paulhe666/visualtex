using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordVisualTeXNumberParenthesisAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(artifactRoot, "visualtex-number-parenthesis.docx");
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"number-parenthesis-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var svgPath = Path.Combine(assetRoot, "visualtex-number-parenthesis.svg");
        File.WriteAllText(svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"220\" height=\"70\" viewBox=\"0 0 220 70\"><text x=\"4\" y=\"48\" font-size=\"36\">x = 1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 220, 70);
        var pngDataUrl = CreatePngDataUrl("number-parenthesis", 220, 70);
        var pngPath = Path.Combine(assetRoot, "visualtex-number-parenthesis.png");
        File.WriteAllBytes(
            pngPath,
            Convert.FromBase64String(pngDataUrl.Substring(pngDataUrl.IndexOf(',') + 1)));

        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);

            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            try
            {
                application.Selection.SetRange(insertion.Start, insertion.End);
                var createSession = CreateNumberedPerformanceSession(
                    "create",
                    formulaId,
                    document.FullName,
                    WordRangeReference(insertion.Start, insertion.End),
                    originalMetadata: null,
                    latex: @"x=1");
                createSession.ExportResult = new OfficeExportDocument
                {
                    Width = 220,
                    Height = 70,
                    Baseline = 52.5f,
                };
                service.InsertOle(createSession, pngPath, emfPath);
                AssertAndExerciseNumberedVisualTeXTypingParagraph(
                    application,
                    document,
                    formulaId);
            }
            finally { Release(insertion); }

            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "fresh numbered VisualTeX insertion");

            shape = FindVisualTeXOleByFormulaIdForNumberToggle(document, formulaId);
            var originalMetadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException("Fresh numbered VisualTeX OLE lost metadata before edit.");
            shapeRange = shape.Range;
            var editSession = CreateNumberedPerformanceSession(
                "edit",
                formulaId,
                document.FullName,
                WordRangeReference(shapeRange.Start, shapeRange.End),
                originalMetadata,
                latex: @"x=2");
            editSession.ExportResult = new OfficeExportDocument
            {
                Width = 220,
                Height = 70,
                Baseline = 52.5f,
            };
            service.ReplaceOle(editSession, pngPath, emfPath);
            Release(shapeRange); shapeRange = null;
            Release(shape); shape = null;

            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "numbered VisualTeX edit/reconcile");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(documentPath, ReadOnly: false, Visible: false);
            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "numbered VisualTeX save/reopen");

            AssertLazySelfContainedVisualTeXReference(
                application,
                document,
                service,
                formulaId);

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                Visible: false);
            AssertPersistedLazySelfContainedVisualTeXReference(
                document,
                formulaId);

            Console.WriteLine(
                "VisualTeX self-contained numbering acceptance passed: fresh insert, safe post-formula typing, edit/reconcile, lazy reference creation and save/reopen kept one MathType-style TAB + OLE + TAB + VisualTeXPlaceRef paragraph with no VTEqCap.");
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            try { Directory.Delete(assetRoot, recursive: true); } catch { }
            ForceComCleanup();
        }
    }

    private static void AssertAndExerciseNumberedVisualTeXTypingParagraph(
        Word.Application application,
        Word.Document document,
        string formulaId)
    {
        Word.Selection? selection = null;
        Word.Bookmarks? bookmarks = null;
        Word.Range? caretRange = null;
        Word.Range? ownerRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Frames? frames = null;
        try
        {
            selection =
                application.Selection;
            caretRange =
                selection.Range.Duplicate;
            bookmarks =
                document.Bookmarks;

            var captionName =
                "VTEqCap_"
                + Guid.Parse(
                    formulaId)
                    .ToString("N");
            var numberName =
                "VTEqNum_"
                + Guid.Parse(
                    formulaId)
                    .ToString("N");

            AssertTrue(
                !bookmarks.Exists(
                    captionName),
                "Fresh self-contained VisualTeX numbering unexpectedly created a VTEqCap hidden-caption bookmark.");
            AssertTrue(
                !bookmarks.Exists(
                    numberName),
                "Fresh self-contained VisualTeX numbering eagerly created VTEqNum before any body reference exists.");

            ownerRange =
                WordVisualTeXParagraphNumbering
                    .FindOwnerParagraphRange(
                        document,
                        formulaId)
                ?? throw new InvalidDataException(
                    "Fresh self-contained VisualTeX numbering has no owner paragraph.");

            AssertTrue(
                caretRange.Start >=
                    ownerRange.End,
                "Fresh numbered VisualTeX insertion left the caret inside its formula/number paragraph.");

            paragraphs =
                caretRange.Paragraphs;
            AssertEqual(
                1,
                paragraphs.Count,
                "Fresh numbered VisualTeX insertion caret does not belong to one ordinary typing paragraph.");
            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            frames =
                paragraphRange.Frames;
            AssertEqual(
                0,
                frames.Count,
                "Fresh numbered VisualTeX insertion left the caret inside a hidden Frame.");
            AssertEqual(
                0,
                paragraphRange.InlineShapes.Count,
                "Fresh numbered VisualTeX insertion typing paragraph already owns an OLE.");
            AssertEqual(
                0,
                paragraphRange.OMaths.Count,
                "Fresh numbered VisualTeX insertion typing paragraph already owns OMML.");
            AssertEqual(
                paragraphRange.Start,
                caretRange.Start,
                "Fresh numbered VisualTeX insertion did not place the caret at the start of its ordinary typing paragraph.");

            var ownerStartBefore =
                ownerRange.Start;
            var ownerEndBefore =
                ownerRange.End;
            var ownerTextBefore =
                ownerRange.Text
                ?? string.Empty;

            selection.TypeParagraph();
            selection.TypeText(
                "after-numbered");

            Release(ownerRange);
            ownerRange =
                WordVisualTeXParagraphNumbering
                    .FindOwnerParagraphRange(
                        document,
                        formulaId)
                ?? throw new InvalidDataException(
                    "Typing after a numbered VisualTeX formula lost its self-contained owner paragraph.");
            AssertEqual(
                ownerStartBefore,
                ownerRange.Start,
                "Typing after a numbered VisualTeX formula moved the owner paragraph start.");
            AssertEqual(
                ownerEndBefore,
                ownerRange.End,
                "Pressing Enter/typing after a numbered VisualTeX formula expanded its self-contained owner paragraph.");
            AssertEqual(
                ownerTextBefore,
                ownerRange.Text
                    ?? string.Empty,
                "Typing after a numbered VisualTeX formula changed its self-contained numbering paragraph.");
            AssertTrue(
                !bookmarks.Exists(
                    captionName),
                "Typing after a new numbered VisualTeX formula recreated a hidden VTEqCap bookmark.");
            AssertTrue(
                !bookmarks.Exists(
                    numberName),
                "Typing after a new numbered VisualTeX formula created a reference bookmark without a reference.");
        }
        finally
        {
            Release(frames);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(ownerRange);
            Release(caretRange);
            Release(bookmarks);
            Release(selection);
        }
    }

    private static void AssertLazySelfContainedVisualTeXReference(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        string formulaId)
    {
        document.Activate();
        Word.Bookmarks? bookmarks = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Selection? selection = null;
        Word.Bookmark? numberBookmark = null;
        Word.Range? numberRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? result = null;
        try
        {
            var bookmarkName =
                "VTEqNum_"
                + Guid.Parse(
                    formulaId)
                    .ToString("N");
            bookmarks =
                document.Bookmarks;
            AssertTrue(
                !bookmarks.Exists(
                    bookmarkName),
                "A fresh self-contained VisualTeX number already had VTEqNum before the first reference.");

            var targets =
                service.GetCanonicalEquationReferenceTargets(
                    document);
            var target =
                targets.SingleOrDefault(candidate =>
                    candidate.Source ==
                        EquationReferenceSource.VisualTeX
                    && string.Equals(
                        candidate.FormulaId,
                        formulaId,
                        StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidDataException(
                    "The self-contained VisualTeX number was not discoverable by the product reference picker.");

            selection =
                application.Selection;
            var referenceInsertion =
                Math.Max(
                    document.Content.Start,
                    document.Content.End - 1);
            selection.SetRange(
                referenceInsertion,
                referenceInsertion);
            selection.TypeParagraph();
            service.InsertEquationReferenceCore(
                document,
                selection,
                target,
                EquationReferenceStyle.NumberOnly,
                Word.WdColor.wdColorAutomatic);

            AssertTrue(
                bookmarks.Exists(
                    bookmarkName),
                "The first VisualTeX body reference did not lazily create VTEqNum.");
            numberBookmark =
                bookmarks[bookmarkName];
            numberRange =
                numberBookmark.Range.Duplicate;
            AssertEqual(
                target.NumberText,
                (numberRange.Text
                    ?? string.Empty)
                    .Trim(),
                "The lazy VTEqNum bookmark does not cover only the visible equation number.");

            fields =
                document.Fields;
            var matchingRefs = 0;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(result);
                result = null;
                Release(code);
                code = null;
                Release(field);
                field =
                    fields[index];
                if (field.Type !=
                    Word.WdFieldType.wdFieldRef)
                    continue;
                code =
                    field.Code.Duplicate;
                if ((code.Text
                        ?? string.Empty)
                    .IndexOf(
                        "REF "
                        + bookmarkName,
                        StringComparison.OrdinalIgnoreCase)
                    < 0)
                    continue;
                matchingRefs++;
                result =
                    field.Result.Duplicate;
                AssertEqual(
                    target.NumberText,
                    (result.Text
                        ?? string.Empty)
                        .Trim(),
                    "The lazy VisualTeX body REF does not display the current equation number.");
            }
            AssertEqual(
                1,
                matchingRefs,
                "The first VisualTeX reference did not create exactly one body REF.");
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(numberRange);
            Release(numberBookmark);
            Release(selection);
            Release(paragraphRange);
            Release(paragraph);
            Release(bookmarks);
        }
    }

    private static void AssertPersistedLazySelfContainedVisualTeXReference(
        Word.Document document,
        string formulaId)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? bookmarkRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? result = null;
        try
        {
            var bookmarkName =
                "VTEqNum_"
                + Guid.Parse(
                    formulaId)
                    .ToString("N");
            var host =
                WordFormulaHostResolver
                    .CaptureDocumentIndex(
                        document)
                    .VisualTeX
                    .Single(candidate =>
                        string.Equals(
                            candidate.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase));
            var expectedNumber =
                WordVisualTeXParagraphNumbering
                    .ReadVisibleNumberText(
                        document,
                        host);
            bookmarks =
                document.Bookmarks;
            AssertTrue(
                bookmarks.Exists(
                    bookmarkName),
                "Saved/reopened self-contained VisualTeX reference lost its lazy VTEqNum target.");
            bookmark =
                bookmarks[bookmarkName];
            bookmarkRange =
                bookmark.Range.Duplicate;
            AssertEqual(
                expectedNumber,
                (bookmarkRange.Text
                    ?? string.Empty)
                    .Trim(),
                "Saved/reopened lazy VTEqNum moved off the visible number.");

            fields =
                document.Fields;
            var matchingRefs = 0;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(result);
                result = null;
                Release(code);
                code = null;
                Release(field);
                field =
                    fields[index];
                if (field.Type !=
                    Word.WdFieldType.wdFieldRef)
                    continue;
                code =
                    field.Code.Duplicate;
                if ((code.Text
                        ?? string.Empty)
                    .IndexOf(
                        "REF "
                        + bookmarkName,
                        StringComparison.OrdinalIgnoreCase)
                    < 0)
                    continue;
                field.Update();
                result =
                    field.Result.Duplicate;
                AssertEqual(
                    expectedNumber,
                    (result.Text
                        ?? string.Empty)
                        .Trim(),
                    "Saved/reopened lazy VisualTeX REF no longer resolves to the self-contained number.");
                matchingRefs++;
            }
            AssertEqual(
                1,
                matchingRefs,
                "Saved/reopened document changed the lazy VisualTeX reference count.");

            var captionName =
                "VTEqCap_"
                + Guid.Parse(
                    formulaId)
                    .ToString("N");
            AssertTrue(
                !bookmarks.Exists(
                    captionName),
                "Saved/reopened self-contained VisualTeX reference recreated VTEqCap.");
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void TrimVisualTeXVisibleNumberLayoutTab(
        Word.Range range)
    {
        var text = range.Text ?? string.Empty;
        var trim = 0;
        while (trim < text.Length
               && text[trim] == '\t')
            trim++;
        if (trim > 0)
            range.SetRange(
                Math.Min(range.End, range.Start + trim),
                range.End);
    }

    private static Word.Range? ResolveVisualTeXVisibleRefDisplayRange(
        Word.Document document,
        Word.Range shapeRange,
        string formulaId)
    {
        Word.Range? result = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? fieldResult = null;
        Word.Range? before = null;
        Word.Range? after = null;
        Word.Range? found = null;
        try
        {
            var selfContained =
                WordVisualTeXParagraphNumbering
                    .FindVisibleLabelRange(
                        document,
                        formulaId);
            if (selfContained is not null)
                return selfContained;

            var host =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    shapeRange,
                    WordFormulaHostKind.VisualTeX);
            if (host is null)
                return null;
            if (string.IsNullOrWhiteSpace(host.FormulaId))
                host.FormulaId = formulaId;
            if (!string.Equals(
                    host.FormulaId,
                    formulaId,
                    StringComparison.OrdinalIgnoreCase))
                return null;

            var numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            if (!numbering.Numbered
                || numbering.ContainerKind !=
                    WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
                || numbering.NumberRange is null)
                return null;

            result =
                WordEquationNumbering.FindVisibleEquationNumberRange(
                    document,
                    formulaId);
            if (result is null
                || result.StoryType != shapeRange.StoryType
                || result.Start < shapeRange.End)
                return null;

            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (result.Start < paragraphRange.Start
                || result.End > paragraphRange.End)
                return null;

            var resolved = result;
            result = null;
            return resolved;

        }
        finally
        {
            Release(found);
            Release(after);
            Release(before);
            Release(fieldResult);
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(result);
        }
    }

    private static void AssertVisualTeXNumberedTabHost(
        Word.Document document,
        string formulaId,
        bool updateReference,
        string context,
        bool requireNativeOle = true,
        bool requireFormulaMetadata = true)
    {
        Word.Range? ownerRange = null;
        Word.Range? visibleRange = null;
        Word.Range? visibleTextRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabStops = null;
        Word.TabStop? tabStop = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Word.Range? objectResultRange = null;
        Word.Font? objectResultFont = null;
        Word.Range? precedingShape = null;
        Word.Range? betweenFormulaAndNumber = null;
        Word.Range? numberEnd = null;
        Word.Font? visibleFont = null;
        Word.Range? paragraphMark = null;
        Word.Font? paragraphMarkFont = null;
        Word.Fields? fields = null;
        Word.Field? reference = null;
        Word.Range? candidateCode = null;
        Word.Range? result = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
        Word.View? view = null;
        var restoreFieldCodes = false;
        try
        {
            document.Repaginate();
            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(document, formulaId)
                ?? throw new InvalidDataException(context + ": VisualTeX numbering owner is missing.");
            AssertTrue(
                !(bool)ownerRange.get_Information(Word.WdInformation.wdWithInTable),
                context + ": numbered VisualTeX OLE is still hosted by a table instead of a MathType-style tab paragraph.");

            paragraphs = ownerRange.Paragraphs;
            AssertEqual(1, paragraphs.Count, context + ": numbered VisualTeX OLE spans multiple paragraphs.");
            paragraph = paragraphs[1];
            format = paragraph.Format;
            AssertEqual(
                Word.WdParagraphAlignment.wdAlignParagraphJustify,
                format.Alignment,
                context + ": numbered VisualTeX paragraph does not match MathType's justified display style.");
            AssertNear(0f, format.LeftIndent, 0.5f,
                context + ": numbered VisualTeX paragraph has an unexpected left indent.");
            AssertNear(0f, format.RightIndent, 0.5f,
                context + ": numbered VisualTeX paragraph has an unexpected right indent.");

            tabStops = format.TabStops;
            var sawCenter = false;
            var sawRight = false;
            var centerPosition = 0f;
            var rightPosition = 0f;
            for (var index = 1; index <= tabStops.Count; index++)
            {
                Release(tabStop); tabStop = tabStops[index];
                if (tabStop.Alignment == Word.WdTabAlignment.wdAlignTabCenter)
                {
                    sawCenter = true;
                    centerPosition = tabStop.Position;
                }
                if (tabStop.Alignment == Word.WdTabAlignment.wdAlignTabRight)
                {
                    sawRight = true;
                    rightPosition = tabStop.Position;
                }
            }
            AssertTrue(sawCenter && sawRight,
                context + ": center/right equation tab stops are missing.");

            sections = ownerRange.Sections;
            AssertTrue(sections.Count > 0, context + ": paragraph has no Word section.");
            section = sections[1];
            pageSetup = section.PageSetup;
            var expected = WordEquationNumbering.CalculateEquationTabStops(
                pageSetup.PageWidth,
                pageSetup.LeftMargin,
                pageSetup.RightMargin,
                0f,
                0f);
            AssertNear(expected.Center, centerPosition, 0.75f,
                context + ": equation center tab is not at the writable-page midpoint.");
            AssertNear(expected.Right, rightPosition, 0.75f,
                context + ": equation-number tab is not at the writable-page right edge.");

            shapes = ownerRange.InlineShapes;
            AssertEqual(1, shapes.Count,
                context + ": tab paragraph does not contain exactly one VisualTeX OLE object.");
            shape = shapes[1];
            if (requireNativeOle)
            {
                AssertTrue(WordFormulaMetadataReader.IsNativeOle(shape),
                    context + ": the tab paragraph object is not a VisualTeX native OLE formula.");
            }
            FormulaMetadata? metadata = null;
            if (requireFormulaMetadata)
            {
                metadata = WordFormulaMetadataReader.TryRead(shape)
                    ?? throw new InvalidDataException(context + ": VisualTeX OLE metadata is missing.");
                AssertEqual(formulaId, metadata.FormulaId,
                    context + ": the tab paragraph belongs to another VisualTeX formula.");
            }
            shapeRange = shape.Range;
            var resolvedHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    shapeRange,
                    WordFormulaHostKind.VisualTeX);
            var selfContainedNumbering =
                resolvedHost is not null
                && WordVisualTeXParagraphNumbering
                    .IsSelfContainedHost(
                        document,
                        resolvedHost);

            if (requireNativeOle && metadata is not null)
            {
                for (var position = shapeRange.Start; position < shapeRange.End; position++)
                {
                    Release(objectResultRange); objectResultRange = null;
                    objectResultRange = document.Range(position, position + 1);
                    if (!string.Equals(objectResultRange.Text, "\u0001", StringComparison.Ordinal))
                        continue;
                    objectResultFont = objectResultRange.Font;
                    break;
                }
                if (objectResultFont is null)
                    throw new InvalidDataException(context + ": VisualTeX OLE has no U+0001 object-result character.");
                var semanticFontSize = FormulaFontSize.ResolveSemanticFontSize(metadata);
                var expectedOlePosition = WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                    shape.Height,
                    (float)(metadata.RenderHeightPx ?? 0d),
                    metadata.Baseline.HasValue ? (float?)metadata.Baseline.Value : null,
                    existingFontPosition: null,
                    sourceSemanticFontSizePoints: semanticFontSize,
                    targetSemanticFontSizePoints: semanticFontSize);
                AssertNear(expectedOlePosition, objectResultFont.Position, 0.1f,
                    context + ": VisualTeX OLE is not using its own exported baseline for display placement.");
                AssertTrue(objectResultFont.Position <= 0,
                    context + ": numbered VisualTeX OLE was raised above the Word text baseline.");
            }

            view = document.ActiveWindow.View;
            restoreFieldCodes = view.ShowFieldCodes;
            if (restoreFieldCodes)
            {
                view.ShowFieldCodes = false;
                System.Windows.Forms.Application.DoEvents();
                document.Repaginate();
            }

            Word.Range? visibleObjectStart = null;
            try
            {
                visibleObjectStart = (objectResultRange ?? shapeRange).Duplicate;
                visibleObjectStart.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                var shapeStartX = Convert.ToSingle(
                    visibleObjectStart.Information[Word.WdInformation.wdHorizontalPositionRelativeToTextBoundary]);
                var shapeCenterX = shapeStartX + shape.Width / 2f;
                AssertNear(expected.Center, shapeCenterX, 0.75f,
                    context + ": VisualTeX OLE is not physically centered on the center tab stop.");
            }
            finally { Release(visibleObjectStart); }

            AssertTrue(shapeRange.Start > ownerRange.Start,
                context + ": formula has no leading center-tab character.");
            precedingShape = document.Range(shapeRange.Start - 1, shapeRange.Start);
            AssertEqual("\t", precedingShape.Text,
                context + ": formula is not positioned after the center tab.");

            visibleRange = ResolveVisualTeXVisibleRefDisplayRange(
                    document,
                    shapeRange,
                    formulaId)
                ?? throw new InvalidDataException(context + ": visible equation-number REF range is missing.");
            AssertTrue(visibleRange.Start >= shapeRange.End,
                context + ": visible equation number precedes or overlaps the formula.");
            betweenFormulaAndNumber = document.Range(shapeRange.End, visibleRange.Start);
            var ownerText = ownerRange.Text ?? string.Empty;
            var betweenText = betweenFormulaAndNumber.Text ?? string.Empty;
            var rawVisible = visibleRange.Text ?? string.Empty;
            Console.WriteLine(
                $"  {context}: raw owner codes={string.Join(",", ownerText.Select(character => $"U+{(int)character:X4}"))}, owner={ownerRange.Start}-{ownerRange.End}, shape={shapeRange.Start}-{shapeRange.End}, visible={visibleRange.Start}-{visibleRange.End} codes={string.Join(",", rawVisible.Select(character => $"U+{(int)character:X4}"))}, between={betweenFormulaAndNumber.Start}-{betweenFormulaAndNumber.End} codes={string.Join(",", betweenText.Select(character => $"U+{(int)character:X4}"))}.");
            AssertTrue(
                ownerText.Count(character => character == '\t') >= 2
                && (betweenText.IndexOf('\t') >= 0 || rawVisible.StartsWith("\t", StringComparison.Ordinal)),
                context + ": formula and equation number are not separated by the right tab.");

            var visible = rawVisible.TrimStart('\t').TrimEnd('\r', '\a');
            visibleTextRange = visibleRange.Duplicate;
            TrimVisualTeXVisibleNumberLayoutTab(visibleTextRange);
            AssertEqual(
                visible,
                (visibleTextRange.Text ?? string.Empty).TrimEnd('\r', '\a'),
                context + ": pure equation-number text range still contains a layout tab or lost visible text.");
            AssertTrue(
                visible.StartsWith("(", StringComparison.Ordinal)
                && visible.EndsWith(")", StringComparison.Ordinal),
                context + ": visible equation number is not enclosed by both parentheses: '" + visible + "'.");

            visibleFont = visibleTextRange.Font;
            AssertNear(0f, visibleFont.Position, 0.1f,
                context + ": tab-layout equation number has a manual vertical offset.");
            paragraphMark = document.Range(ownerRange.End - 1, ownerRange.End);
            paragraphMarkFont = paragraphMark.Font;
            AssertNear(0f, paragraphMarkFont.Position, 0.1f,
                context + ": display paragraph mark has a manual vertical offset.");
            AssertNear(paragraphMarkFont.Size, visibleFont.Size, 0.1f,
                context + ": VisualTeX number does not inherit the display paragraph point size.");
            var paragraphFontName = paragraphMarkFont.NameAscii ?? paragraphMarkFont.Name ?? string.Empty;
            var visibleFontName = visibleFont.NameAscii ?? visibleFont.Name ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(paragraphFontName)
                && !string.IsNullOrWhiteSpace(visibleFontName))
            {
                AssertEqual(paragraphFontName, visibleFontName,
                    context + ": VisualTeX number does not inherit the display paragraph typeface.");
            }
            numberEnd = visibleTextRange.Duplicate;
            numberEnd.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            var numberEndX = Convert.ToSingle(
                numberEnd.Information[Word.WdInformation.wdHorizontalPositionRelativeToTextBoundary]);
            if (!selfContainedNumbering)
            {
                AssertNear(expected.Right, numberEndX, 0.75f,
                    context + ": legacy visible equation number does not end on the right tab stop.");
            }
            var numberY = Convert.ToSingle(
                visibleTextRange.Information[Word.WdInformation.wdVerticalPositionRelativeToPage]);
            var paragraphMarkY = Convert.ToSingle(
                paragraphMark.Information[Word.WdInformation.wdVerticalPositionRelativeToPage]);
            AssertNear(paragraphMarkY, numberY, 0.75f,
                context + ": equation number and paragraph mark are not on the same Word baseline.");

            if (selfContainedNumbering)
            {
                AssertSelfContainedVisualTeXNumberFieldTree(
                    document,
                    ownerRange,
                    formulaId,
                    visible,
                    updateReference,
                    context);
                Console.WriteLine(
                    $"  {context}: self-contained VisualTeXPlaceRef owner={ownerRange.Start}-{ownerRange.End}, "
                    + $"formula={shapeRange.Start}-{shapeRange.End}, visible='{visible}', "
                    + $"tabs={centerPosition:0.##}/{rightPosition:0.##}, "
                    + $"olePosition={objectResultFont?.Position ?? 0}, numberPosition={visibleFont.Position}, "
                    + $"numberFont='{visibleFontName}', numberY={numberY:0.##}, markY={paragraphMarkY:0.##}.");
                return;
            }

            fields = visibleRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? candidate = null;
                try
                {
                    candidate = fields[index];
                    Release(candidateCode); candidateCode = candidate.Code;
                    if ((candidateCode.Text ?? string.Empty).IndexOf(
                            "REF " + WordEquationNumbering.NativeNumberBookmarkName(formulaId),
                            StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    reference = candidate;
                    candidate = null;
                    break;
                }
                finally { Release(candidate); }
            }
            if (reference is null)
                throw new InvalidDataException(context + ": visible number REF field is missing.");

            if (updateReference)
            {
                reference.Update();
                if (requireNativeOle)
                {
                    AssertTrue(
                        WordEquationNumbering.UpdateEquationNumbers(document) >= 1,
                        context + ": explicit equation-number refresh did not see the numbered VisualTeX formula.");
                    WordEquationNumbering.UpdateNativeCrossReferences(document);
                }
            }
            Release(visibleRange); visibleRange = null;
            visibleRange = ResolveVisualTeXVisibleRefDisplayRange(
                    document,
                    shapeRange,
                    formulaId)
                ?? throw new InvalidDataException(context + ": REF update removed the visible equation-number REF range.");
            visible = (visibleRange.Text ?? string.Empty)
                .TrimStart('\t')
                .TrimEnd('\r', '\a');
            AssertTrue(
                visible.StartsWith("(", StringComparison.Ordinal)
                && visible.EndsWith(")", StringComparison.Ordinal),
                context + ": REF update removed a parenthesis: '" + visible + "'.");

            Release(visibleFont); visibleFont = null;
            Release(visibleTextRange); visibleTextRange = null;
            visibleTextRange = visibleRange.Duplicate;
            TrimVisualTeXVisibleNumberLayoutTab(visibleTextRange);
            visibleFont = visibleTextRange.Font;
            AssertNear(0f, visibleFont.Position, 0.1f,
                context + ": F9/update-number moved the tab-layout number off the paragraph baseline.");
            AssertNear(paragraphMarkFont.Size, visibleFont.Size, 0.1f,
                context + ": F9/update-number replaced the inherited paragraph point size.");
            visibleFontName = visibleFont.NameAscii ?? visibleFont.Name ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(paragraphFontName)
                && !string.IsNullOrWhiteSpace(visibleFontName))
            {
                AssertEqual(paragraphFontName, visibleFontName,
                    context + ": F9/update-number replaced the inherited paragraph typeface.");
            }

            result = reference.Result;
            var resultText = result.Text ?? string.Empty;
            AssertTrue(
                resultText.IndexOf('(') < 0 && resultText.IndexOf(')') < 0,
                context + ": parenthesis leaked inside REF.Result: '" + resultText + "'.");
            AssertTrue(
                result.Start >= visibleRange.Start && result.End < visibleRange.End,
                context + ": REF.Result reaches outside the visible-number bookmark; ')' is not safely outside the field result.");
            Console.WriteLine(
                $"  {context}: owner={ownerRange.Start}-{ownerRange.End}, formula={shapeRange.Start}-{shapeRange.End}, visible='{visible}', tabs={centerPosition:0.##}/{rightPosition:0.##}, olePosition={objectResultFont?.Position ?? 0}, numberPosition={visibleFont.Position}, numberFont='{visibleFontName}', numberY={numberY:0.##}, markY={paragraphMarkY:0.##}, REF.Result='{resultText}'.");
        }
        finally
        {
            if (view is not null && restoreFieldCodes)
            {
                try { view.ShowFieldCodes = true; } catch { }
            }
            Release(view);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(result);
            Release(reference);
            Release(paragraphMarkFont);
            Release(paragraphMark);
            Release(visibleFont);
            Release(numberEnd);
            Release(candidateCode);
            Release(fields);
            Release(betweenFormulaAndNumber);
            Release(precedingShape);
            Release(objectResultFont);
            Release(objectResultRange);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(tabStop);
            Release(tabStops);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(visibleTextRange);
            Release(visibleRange);
            Release(ownerRange);
        }
    }

    private static void AssertSelfContainedVisualTeXNumberFieldTree(
        Word.Document document,
        Word.Range ownerRange,
        string formulaId,
        string visibleNumber,
        bool updateNumber,
        string context)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Field? placeRef = null;
        Word.Fields? nested = null;
        Word.Field? child = null;
        Word.Range? childCode = null;
        Word.Bookmarks? bookmarks = null;
        Word.Range? refreshedVisible = null;
        try
        {
            fields =
                ownerRange.Fields;
            var placeRefCount = 0;
            var directVisualTeXRefCount = 0;
            var externalChapterStateCount = 0;
            var externalSectionStateCount = 0;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(code);
                code = null;
                Release(field);
                field =
                    fields[index];
                code =
                    field.Code.Duplicate;
                var instruction =
                    code.Text
                    ?? string.Empty;
                var normalizedInstruction =
                    instruction.Trim();
                if (normalizedInstruction.StartsWith(
                        "SEQ VisualTeXChapter ",
                        StringComparison.OrdinalIgnoreCase)
                    && normalizedInstruction.IndexOf(
                        "\\r",
                        StringComparison.OrdinalIgnoreCase) >= 0
                    && normalizedInstruction.IndexOf(
                        "\\h",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    externalChapterStateCount++;
                }
                if (normalizedInstruction.StartsWith(
                        "SEQ VisualTeXSection ",
                        StringComparison.OrdinalIgnoreCase)
                    && normalizedInstruction.IndexOf(
                        "\\r",
                        StringComparison.OrdinalIgnoreCase) >= 0
                    && normalizedInstruction.IndexOf(
                        "\\h",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    externalSectionStateCount++;
                }

                if (field.Type ==
                        Word.WdFieldType.wdFieldMacroButton
                    && WordVisualTeXParagraphNumbering
                        .IsPlaceRefCode(
                            instruction))
                {
                    placeRefCount++;
                    Release(placeRef);
                    placeRef =
                        field;
                    field = null;
                    continue;
                }

                if (field.Type ==
                        Word.WdFieldType.wdFieldRef
                    && instruction.IndexOf(
                        "VTEqNum_",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    directVisualTeXRefCount++;
            }

            AssertEqual(
                1,
                placeRefCount,
                context
                + ": self-contained VisualTeX paragraph does not own exactly one VisualTeXPlaceRef field.");
            AssertEqual(
                0,
                directVisualTeXRefCount,
                context
                + ": self-contained VisualTeX paragraph still contains the legacy REF-to-VTEqNum visible-number field.");
            AssertTrue(
                placeRef is not null,
                context
                + ": self-contained VisualTeXPlaceRef field was not retained for nested-field validation.");

            var expectedFormat =
                EquationNumberFormat.Resolve(
                    WordEquationNumbering
                        .GetEquationNumberFormatId(
                            document));
            AssertEqual(
                expectedFormat.HeadingLevel >= 1 ? 1 : 0,
                externalChapterStateCount,
                context
                + ": owner paragraph has the wrong external VisualTeXChapter state-field count.");
            AssertEqual(
                expectedFormat.HeadingLevel >= 2 ? 1 : 0,
                externalSectionStateCount,
                context
                + ": owner paragraph has the wrong external VisualTeXSection state-field count.");

            Release(code);
            code = null;
            code =
                placeRef!.Code.Duplicate;
            nested =
                code.Fields;
            var hiddenIncrementCount = 0;
            var currentValueCount = 0;
            var chapterRestartCount = 0;
            var chapterCurrentCount = 0;
            var sectionRestartCount = 0;
            var sectionCurrentCount = 0;
            var styleRefCount = 0;
            for (var index = 1;
                 index <= nested.Count;
                 index++)
            {
                Release(childCode);
                childCode = null;
                Release(child);
                child =
                    nested[index];
                childCode =
                    child.Code.Duplicate;
                var instruction =
                    (childCode.Text
                        ?? string.Empty)
                    .Trim();

                if (instruction.StartsWith(
                        "SEQ VisualTeXEquation ",
                        StringComparison.OrdinalIgnoreCase)
                    && instruction.IndexOf(
                        "\\h",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    hiddenIncrementCount++;
                    continue;
                }

                if (instruction.StartsWith(
                        "SEQ VisualTeXEquation ",
                        StringComparison.OrdinalIgnoreCase)
                    && instruction.IndexOf(
                        "\\c",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    currentValueCount++;
                    continue;
                }

                if (instruction.StartsWith(
                        "SEQ VisualTeXChapter ",
                        StringComparison.OrdinalIgnoreCase))
                {
                    if (instruction.IndexOf(
                            "\\r",
                            StringComparison.OrdinalIgnoreCase) >= 0
                        && instruction.IndexOf(
                            "\\h",
                            StringComparison.OrdinalIgnoreCase) >= 0)
                        chapterRestartCount++;
                    else if (instruction.IndexOf(
                                 "\\c",
                                 StringComparison.OrdinalIgnoreCase) >= 0)
                        chapterCurrentCount++;
                    else
                        throw new InvalidDataException(
                            context
                            + ": VisualTeXChapter field has an unsupported shape: '"
                            + instruction
                            + "'.");
                    continue;
                }

                if (instruction.StartsWith(
                        "SEQ VisualTeXSection ",
                        StringComparison.OrdinalIgnoreCase))
                {
                    if (instruction.IndexOf(
                            "\\r",
                            StringComparison.OrdinalIgnoreCase) >= 0
                        && instruction.IndexOf(
                            "\\h",
                            StringComparison.OrdinalIgnoreCase) >= 0)
                        sectionRestartCount++;
                    else if (instruction.IndexOf(
                                 "\\c",
                                 StringComparison.OrdinalIgnoreCase) >= 0)
                        sectionCurrentCount++;
                    else
                        throw new InvalidDataException(
                            context
                            + ": VisualTeXSection field has an unsupported shape: '"
                            + instruction
                            + "'.");
                    continue;
                }

                if (instruction.StartsWith(
                        "STYLEREF ",
                        StringComparison.OrdinalIgnoreCase))
                {
                    styleRefCount++;
                    continue;
                }

                throw new InvalidDataException(
                    context
                    + ": VisualTeXPlaceRef contains a non-canonical nested field: '"
                    + instruction
                    + "'.");
            }

            AssertEqual(
                1,
                hiddenIncrementCount,
                context
                + ": VisualTeXPlaceRef lost its single hidden SEQ increment.");
            AssertEqual(
                1,
                currentValueCount,
                context
                + ": VisualTeXPlaceRef lost its single visible current-value SEQ.");
            var format =
                EquationNumberFormat.Resolve(
                    WordEquationNumbering
                        .GetEquationNumberFormatId(
                            document));
            AssertEqual(
                0,
                chapterRestartCount,
                context
                + ": VisualTeXPlaceRef must not contain the external VisualTeXChapter restart field.");
            AssertEqual(
                format.HeadingLevel >= 1 ? 1 : 0,
                chapterCurrentCount,
                context
                + ": VisualTeXPlaceRef has the wrong VisualTeXChapter current-value count.");
            AssertEqual(
                0,
                sectionRestartCount,
                context
                + ": VisualTeXPlaceRef must not contain the external VisualTeXSection restart field.");
            AssertEqual(
                format.HeadingLevel >= 2 ? 1 : 0,
                sectionCurrentCount,
                context
                + ": VisualTeXPlaceRef has the wrong VisualTeXSection current-value count.");
            AssertEqual(
                0,
                styleRefCount,
                context
                + ": VisualTeXPlaceRef must not depend on Word/OMML STYLEREF state.");
            AssertTrue(
                visibleNumber.StartsWith(
                    "(",
                    StringComparison.Ordinal)
                && visibleNumber.EndsWith(
                    ")",
                    StringComparison.Ordinal)
                && visibleNumber.IndexOf(
                    "Error",
                    StringComparison.OrdinalIgnoreCase) < 0
                && visibleNumber.IndexOf(
                    "错误",
                    StringComparison.OrdinalIgnoreCase) < 0,
                context
                + ": self-contained VisualTeXPlaceRef does not expose a healthy parenthesized number: '"
                + visibleNumber
                + "'.");

            bookmarks =
                document.Bookmarks;
            var captionName =
                "VTEqCap_"
                + Guid.Parse(
                    formulaId)
                    .ToString("N");
            var numberName =
                "VTEqNum_"
                + Guid.Parse(
                    formulaId)
                    .ToString("N");
            AssertTrue(
                !bookmarks.Exists(
                    captionName),
                context
                + ": new self-contained VisualTeX numbering created a VTEqCap hidden-caption bookmark.");
            AssertTrue(
                !bookmarks.Exists(
                    numberName),
                context
                + ": new self-contained VisualTeX numbering eagerly created VTEqNum before a body reference exists.");

            if (updateNumber)
            {
                AssertTrue(
                    WordFormulaNumberingKernel
                        .RefreshCanonicalNumbers(
                            document) >= 1,
                    context
                    + ": canonical number refresh did not see the self-contained VisualTeXPlaceRef host.");

                refreshedVisible =
                    WordVisualTeXParagraphNumbering
                        .FindVisibleLabelRange(
                            document,
                            formulaId)
                    ?? throw new InvalidDataException(
                        context
                        + ": canonical number refresh removed the VisualTeXPlaceRef visible number.");
                var refreshedText =
                    (refreshedVisible.Text
                        ?? string.Empty)
                    .Trim();
                AssertEqual(
                    visibleNumber,
                    refreshedText,
                    context
                    + ": canonical number refresh changed the self-contained visible number unexpectedly.");
                AssertTrue(
                    !bookmarks.Exists(
                        captionName),
                    context
                    + ": number refresh recreated the retired VTEqCap hidden-caption bookmark.");
            }
        }
        finally
        {
            Release(refreshedVisible);
            Release(bookmarks);
            Release(childCode);
            Release(child);
            Release(nested);
            Release(placeRef);
            Release(code);
            Release(field);
            Release(fields);
        }
    }
}
