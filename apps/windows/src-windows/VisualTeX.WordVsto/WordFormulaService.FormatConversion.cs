using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    public WordFormulaFormatConversionPlan CaptureFormulaFormatConversionPlan(
        bool wholeDocument,
        string sourceMode,
        string targetMode)
    {
        ValidateSimpleFormatConversionPair(sourceMode, targetMode);
        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        InlineShapes? shapes = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            selection = _application.Selection;
            scope = wholeDocument
                ? document.Content.Duplicate
                : selection.Range.Duplicate;

            var plan = new WordFormulaFormatConversionPlan
            {
                DocumentId = DocumentIdentity(document),
                SourceMode = sourceMode,
                TargetMode = targetMode,
                WholeDocument = wholeDocument,
            };

            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                try
                {
                    shape = shapes[index];
                    FormulaMetadata? metadata = null;
                    var mathTypeNumberPosition = "right";
                    string sourceFormulaId;

                    if (string.Equals(
                            sourceMode,
                            FormulaOleContract.NativeOleMode,
                            StringComparison.Ordinal))
                    {
                        if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                        metadata = WordFormulaMetadataReader.TryRead(shape);
                        if (metadata is null) continue;
                        sourceFormulaId = metadata.FormulaId;
                    }
                    else
                    {
                        if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                        var mathMl = MathTypeOleStorage.ReadMathMl(shape);
                        metadata = MathTypeOleInterop.ReadMetadata(
                            _application,
                            shape,
                            mathMl);
                        if (MathTypeOleInterop.TryReadDisplayNumberPosition(
                                shape,
                                out var detectedPosition))
                            mathTypeNumberPosition = detectedPosition;
                        sourceFormulaId = metadata.FormulaId;
                    }

                    range = shape.Range;
                    if (!FormulaRangeMatchesScope(range, scope, wholeDocument))
                        continue;

                    var latex = string.IsNullOrWhiteSpace(metadata.Latex)
                        ? string.Join("\n", metadata.Lines.Select(line => line.Latex))
                        : metadata.Latex;
                    latex = (latex ?? string.Empty).Trim();
                    if (latex.Length == 0)
                        throw new InvalidDataException(
                            "A source formula has no recoverable LaTeX and was not converted.");

                    plan.Targets.Add(new WordFormulaFormatConversionTarget
                    {
                        Id = Guid.NewGuid().ToString("D"),
                        SourceFormulaId = sourceFormulaId,
                        SourceObjectId = $"{RangeReferencePrefix}{range.Start}:{range.End}",
                        SourceStart = range.Start,
                        Latex = latex,
                        DisplayMode = metadata.DisplayMode,
                        Numbered = metadata.Numbered,
                        MathTypeNumberPosition = mathTypeNumberPosition,
                        FontSizePt = FormulaFontSize.Normalize(metadata.FontSizePt),
                        Metadata = metadata,
                    });
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }

            return plan;
        }
        finally
        {
            Release(shapes);
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    public WordFormulaFormatConversionResult ApplyFormulaFormatConversionPlan(
        WordFormulaFormatConversionPlan plan,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        if (prepared is null) throw new ArgumentNullException(nameof(prepared));
        ValidateSimpleFormatConversionPair(plan.SourceMode, plan.TargetMode);

        Document? document = null;
        Selection? selection = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, plan.DocumentId);
            selection = _application.Selection;

            foreach (var target in plan.Targets)
            {
                if (!prepared.TryGetValue(target.Id, out var formula))
                    throw new InvalidDataException(
                        $"Missing rendered payload for formula '{target.Latex}'.");
                ValidatePreparedFormatConversionTarget(plan.TargetMode, target, formula);
                ValidateSimpleSourceHost(document, plan.SourceMode, target);
            }

            var result = new WordFormulaFormatConversionResult();
            foreach (var target in plan.Targets.OrderByDescending(item => item.SourceStart))
            {
                try
                {
                    var formula = prepared[target.Id];
                    var insertionStart = DeleteSimpleSourceHost(
                        document,
                        plan.SourceMode,
                        target);
                    var content = document.Content;
                    try
                    {
                        insertionStart = Math.Max(
                            content.Start,
                            Math.Min(insertionStart, content.End));
                    }
                    finally { Release(content); }

                    selection.SetRange(insertionStart, insertionStart);
                    selection.Collapse(WdCollapseDirection.wdCollapseStart);

                    var session = formula.Session;
                    session.Mode = "create";
                    session.SourceDocumentId = plan.DocumentId;
                    session.SourceObjectId = null;
                    session.DisplayMode = target.DisplayMode;
                    session.ObjectMode = plan.TargetMode;
                    session.Numbered = target.Numbered;
                    session.MathTypeNumberPosition = target.MathTypeNumberPosition;
                    session.FontSizePt = target.FontSizePt;
                    session.OriginalMetadata = target.Metadata;

                    if (string.Equals(
                            plan.TargetMode,
                            FormulaOleContract.MathTypeOleMode,
                            StringComparison.Ordinal))
                    {
                        InsertMathTypeOle(
                            session,
                            formula.MathMl!,
                            formula.EmfPath!);
                    }
                    else
                    {
                        InsertOle(
                            session,
                            formula.PngPath!,
                            formula.EmfPath!);
                    }
                    result.FormulaCount++;
                }
                catch (Exception error)
                {
                    result.FailedFormulaCount++;
                    result.Failures.Add($"{target.Latex}: {error.Message}");
                    // Stop at the first real Word failure. Continuing after a host
                    // replacement fails makes the document harder to reason about
                    // and was a major source of the previous conversion corruption.
                    break;
                }
            }

            if (result.FormulaCount > 0)
            {
                if (string.Equals(
                        plan.TargetMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal))
                    MathTypeEquationNumbering.UpdateEquationNumbers(document);
                else
                    WordEquationNumbering.TryReconcile(document);
            }
            return result;
        }
        finally
        {
            Release(selection);
            Release(document);
        }
    }

    private static void ValidateSimpleFormatConversionPair(
        string sourceMode,
        string targetMode)
    {
        var visualTeXToMathType =
            string.Equals(sourceMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal)
            && string.Equals(targetMode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal);
        var mathTypeToVisualTeX =
            string.Equals(sourceMode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal)
            && string.Equals(targetMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal);
        if (!visualTeXToMathType && !mathTypeToVisualTeX)
            throw new ArgumentOutOfRangeException(
                nameof(targetMode),
                $"Unsupported simple format conversion: {sourceMode} -> {targetMode}.");
    }

    private static void ValidatePreparedFormatConversionTarget(
        string targetMode,
        WordFormulaFormatConversionTarget target,
        PreparedWordBulkFormula formula)
    {
        if (string.IsNullOrWhiteSpace(formula.MathMl))
            throw new InvalidDataException(
                $"Formula '{target.Latex}' did not produce MathML.");
        if (string.IsNullOrWhiteSpace(formula.EmfPath)
            || !File.Exists(formula.EmfPath))
            throw new FileNotFoundException(
                $"Formula '{target.Latex}' did not produce an EMF preview.",
                formula.EmfPath);
        if (string.Equals(
                targetMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal)
            && (string.IsNullOrWhiteSpace(formula.PngPath)
                || !File.Exists(formula.PngPath)))
            throw new FileNotFoundException(
                $"Formula '{target.Latex}' did not produce a PNG preview.",
                formula.PngPath);
    }

    private void ValidateSimpleSourceHost(
        Document document,
        string sourceMode,
        WordFormulaFormatConversionTarget target)
    {
        InlineShape? shape = null;
        Table? table = null;
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                shape = FindByFormulaId(
                        document,
                        target.SourceFormulaId,
                        target.SourceObjectId)
                    ?? throw new InvalidOperationException(
                        "The VisualTeX source formula moved before conversion started.");
                if (!WordFormulaMetadataReader.IsNativeOle(shape))
                    throw new InvalidOperationException(
                        "The source object is no longer a VisualTeX OLE formula.");
                if (target.Numbered
                    && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                {
                    table = WordEquationNumbering.FindNumberedEquationTable(
                            document,
                            target.SourceFormulaId)
                        ?? throw new InvalidOperationException(
                            "The numbered VisualTeX source no longer owns its numbering table.");
                    if (table.Rows.Count != 1 || table.Columns.Count != 3)
                        throw new InvalidOperationException(
                            "The VisualTeX numbering table is not a normal 1x3 formula host.");
                    if (table.Range.InlineShapes.Count != 1)
                        throw new InvalidOperationException(
                            "The VisualTeX numbering table contains more than one formula object; conversion was refused before modifying the document.");
                }
                return;
            }

            shape = FindMathTypeOleByRange(document, target.SourceObjectId)
                ?? throw new InvalidOperationException(
                    "The MathType source formula moved before conversion started.");
            if (!MathTypeOleInterop.IsMathTypeOle(shape))
                throw new InvalidOperationException(
                    "The source object is no longer Equation.DSMT4.");
            if (!string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                return;
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display formula no longer occupies one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.InlineShapes.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display paragraph contains another inline object; conversion was refused.");
            if (!IsSafeMathTypeDisplayParagraph(paragraphRange))
                throw new InvalidOperationException(
                    "The MathType display paragraph contains ordinary user text; conversion was refused to avoid deleting prose.");
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
            Release(table);
            Release(shape);
        }
    }

    private int DeleteSimpleSourceHost(
        Document document,
        string sourceMode,
        WordFormulaFormatConversionTarget target)
    {
        InlineShape? shape = null;
        Range? shapeRange = null;
        Table? table = null;
        Range? tableRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? contentRange = null;
        try
        {
            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                shape = FindByFormulaId(
                        document,
                        target.SourceFormulaId,
                        target.SourceObjectId)
                    ?? throw new InvalidOperationException(
                        "The VisualTeX source formula moved before replacement.");
                shapeRange = shape.Range;
                var start = shapeRange.Start;
                if (target.Numbered
                    && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                {
                    table = WordEquationNumbering.FindNumberedEquationTable(
                            document,
                            target.SourceFormulaId)
                        ?? throw new InvalidOperationException(
                            "The numbered VisualTeX source lost its table before replacement.");
                    tableRange = table.Range.Duplicate;
                    start = tableRange.Start;
                    table.Delete();
                    TryDeleteBookmark(
                        document,
                        WordEquationNumbering.EquationBookmarkName(target.SourceFormulaId));
                    RemoveDetachedVisualTeXNumberingArtifacts(
                        document,
                        target.SourceFormulaId);
                }
                else
                {
                    if (string.Equals(target.DisplayMode, "inline", StringComparison.Ordinal))
                    {
                        RemoveInlineBaselineSentinel(document, target.SourceFormulaId);
                        RemoveInlineOleTypingAnchorAfter(shape);
                    }
                    shape.Delete();
                }
                TryDeleteBookmark(
                    document,
                    WordFormulaMetadataReader.IdentityBookmarkName(target.SourceFormulaId));
                return start;
            }

            shape = FindMathTypeOleByRange(document, target.SourceObjectId)
                ?? throw new InvalidOperationException(
                    "The MathType source formula moved before replacement.");
            shapeRange = shape.Range;
            var mathTypeStart = shapeRange.Start;
            if (!string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
            {
                RemoveInlineOleTypingAnchorAfter(shape);
                shape.Delete();
                return mathTypeStart;
            }

            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display source no longer occupies one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (!IsSafeMathTypeDisplayParagraph(paragraphRange))
                throw new InvalidOperationException(
                    "The MathType display paragraph contains ordinary user text; conversion was stopped.");
            contentRange = paragraphRange.Duplicate;
            mathTypeStart = contentRange.Start;
            var text = contentRange.Text ?? string.Empty;
            if (contentRange.End > contentRange.Start
                && text.Length > 0
                && (text[text.Length - 1] == '\r' || text[text.Length - 1] == '\a'))
                contentRange.SetRange(contentRange.Start, contentRange.End - 1);
            contentRange.Delete();
            return mathTypeStart;
        }
        finally
        {
            Release(contentRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(tableRange);
            Release(table);
            Release(shapeRange);
            Release(shape);
        }
    }

    private static bool IsSafeMathTypeDisplayParagraph(Range paragraphRange)
    {
        var text = paragraphRange.Text ?? string.Empty;
        foreach (var character in text)
        {
            if (char.IsWhiteSpace(character) || char.IsDigit(character)) continue;
            if (character < ' ') continue;
            if (character is '\u0001' or '\u0013' or '\u0014' or '\u0015') continue;
            if ("()[]{}.-–—_:;,+/\\".IndexOf(character) >= 0) continue;
            return false;
        }
        return true;
    }

    private static void RemoveDetachedVisualTeXNumberingArtifacts(
        Document document,
        string formulaId)
    {
        // The visible 1x3 numbered host has already been deleted. Only the hidden
        // native SEQ caption/frame remains. Delete that detached structure without
        // touching any former table/cell Range.
        TryDeleteBookmark(document, WordEquationNumbering.NativeNumberBookmarkName(formulaId));

        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Range? captionRange = null;
        Frames? frames = null;
        Frame? frame = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = WordEquationNumbering.NativeCaptionBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName)) return;
            captionBookmark = bookmarks[captionName];
            captionRange = captionBookmark.Range;
            try
            {
                frames = captionRange.Frames;
                if (frames.Count > 0)
                {
                    frame = frames[1];
                    frame.Delete();
                    Release(frame);
                    frame = null;
                    Release(frames);
                    frames = null;
                    Release(captionRange);
                    captionRange = captionBookmark.Range;
                }
            }
            catch
            {
                // The caption may already have lost its clipping frame. Deleting
                // the bookmarked caption contents is sufficient in that case.
            }
            captionRange.Delete();
        }
        finally
        {
            Release(frame);
            Release(frames);
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    private static void TryDeleteBookmark(Document document, string name)
    {
        if (document is null || string.IsNullOrWhiteSpace(name)) return;
        try
        {
            if (!document.Bookmarks.Exists(name)) return;
            Bookmark? bookmark = null;
            try
            {
                bookmark = document.Bookmarks[name];
                bookmark.Delete();
            }
            finally { Release(bookmark); }
        }
        catch
        {
            // The source host itself has already been removed. A collapsed stale
            // bookmark is harmless and must never turn cleanup into a conversion failure.
        }
    }
}
