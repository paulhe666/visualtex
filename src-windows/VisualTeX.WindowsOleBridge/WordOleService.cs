using System.Drawing;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOleBridge;

internal sealed class WordOleService : IWordFormulaService
{
    private const int MsoTrue = -1;
    private const int WdCollapseEnd = 0;
    private const int WdCollapseStart = 1;
    private const int WdAlignParagraphLeft = 0;
    private const int WdAlignParagraphCenter = 1;
    private const int WdTabAlignmentCenter = 1;
    private const int WdTabAlignmentRight = 2;
    private const int WdTabLeaderSpaces = 0;
    private const int WdFieldEmpty = -1;
    private const string RangeReferencePrefix = "visualtex-word-ole-range:";
    private const string EquationSequenceName = "VisualTeXEquation";
    private const string EquationBookmarkPrefix = "VTEq_";

    public OfficeSelection GetSelection()
    {
        object? app = null;
        object? document = null;
        object? selection = null;
        object? range = null;
        object? inlineShapes = null;
        object? shape = null;
        try
        {
            app = RunningOfficeLocator.GetWordApplication();
            dynamic word = app;
            document = word.ActiveDocument;
            selection = word.Selection;
            range = ((dynamic)selection).Range;
            inlineShapes = ((dynamic)range).InlineShapes;
            FormulaMetadata? metadata = null;
            if (Convert.ToInt32(((dynamic)inlineShapes).Count) == 1)
            {
                shape = ((dynamic)inlineShapes).Item(1);
                metadata = ReadMetadata(shape);
            }
            return new OfficeSelection
            {
                Host = "word",
                DocumentId = DocumentIdentity((dynamic)document),
                ObjectId = metadata?.FormulaId ?? RangeReference(range),
                ReadOnly = IsReadOnly((dynamic)document),
                FormulaId = metadata?.FormulaId,
                Metadata = metadata,
            };
        }
        finally
        {
            ComRelease.Final(shape);
            ComRelease.Final(inlineShapes);
            ComRelease.Final(range);
            ComRelease.Final(selection);
            ComRelease.Final(document);
            ComRelease.Final(app);
        }
    }

    public OfficeObjectResult InsertInlineFormula(SessionInfo session) =>
        InsertFormula(session, display: false);

    public OfficeObjectResult InsertDisplayFormula(SessionInfo session) =>
        InsertFormula(session, display: true);

    public OfficeObjectResult ReplaceFormula(SessionInfo session)
    {
        session.Metadata.Validate();
        object? app = null;
        object? document = null;
        object? oldShape = null;
        object? oldRange = null;
        object? insertionRange = null;
        object? newShape = null;
        object? undoRecord = null;
        var replacementCommitted = false;
        try
        {
            app = RunningOfficeLocator.GetWordApplication();
            undoRecord = BeginUndoRecord(app, "VisualTeX Replace Formula");
            document = ((dynamic)app).ActiveDocument;
            dynamic doc = document;
            EnsureWritable(doc);
            EnsureSourceDocument(doc, session.SourceDocumentId);
            oldShape = FindFormula(doc, session.FormulaId);
            if (oldShape is null)
                throw new InvalidOperationException("The target Word formula no longer exists.");

            dynamic original = oldShape;
            oldRange = original.Range;
            insertionRange = ((dynamic)oldRange).Duplicate;
            ((dynamic)insertionRange).Collapse(WdCollapseStart);
            newShape = ReplacementTransaction.Execute<object>(
                () => ((dynamic)insertionRange).InlineShapes.AddPicture(
                    session.ImagePath,
                    false,
                    true,
                    insertionRange),
                candidate => ConfigureShape(
                    candidate,
                    session,
                    session.Metadata.DisplayMode == "inline"),
                () => original.Delete(),
                TryDelete);
            replacementCommitted = true;
            TryReconcileEquationNumbers(doc);
            return Result(session, doc);
        }
        catch
        {
            if (!replacementCommitted)
                TryDelete(newShape);
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            ComRelease.Final(undoRecord);
            ComRelease.Final(newShape);
            ComRelease.Final(insertionRange);
            ComRelease.Final(oldRange);
            ComRelease.Final(oldShape);
            ComRelease.Final(document);
            ComRelease.Final(app);
        }
    }

    private static OfficeObjectResult InsertFormula(SessionInfo session, bool display)
    {
        session.Metadata.Validate();
        object? app = null;
        object? document = null;
        object? selection = null;
        object? sourceRange = null;
        object? insertionRange = null;
        object? paragraph = null;
        object? paragraphRange = null;
        object? shape = null;
        object? undoRecord = null;
        try
        {
            app = RunningOfficeLocator.GetWordApplication();
            undoRecord = BeginUndoRecord(
                app,
                display ? "VisualTeX Insert Display Formula" : "VisualTeX Insert Inline Formula");
            dynamic word = app;
            document = word.ActiveDocument;
            dynamic doc = document;
            EnsureWritable(doc);
            EnsureSourceDocument(doc, session.SourceDocumentId);
            selection = word.Selection;
            sourceRange = ResolveSourceRange(doc, session.SourceObjectId, selection);
            insertionRange = ((dynamic)sourceRange).Duplicate;
            ((dynamic)insertionRange).Collapse(WdCollapseEnd);

            if (display)
            {
                paragraph = TryGetReusableDisplayParagraph(sourceRange);
                var reuseCurrentParagraph = paragraph is not null;
                if (!reuseCurrentParagraph)
                {
                    ((dynamic)insertionRange).InsertParagraphAfter();
                    ((dynamic)insertionRange).Collapse(WdCollapseEnd);
                    paragraph = doc.Paragraphs.Add(insertionRange);
                }
                ((dynamic)paragraph).Alignment = WdAlignParagraphCenter;
                paragraphRange = ((dynamic)paragraph).Range;
                object pictureRange = reuseCurrentParagraph ? insertionRange : paragraphRange;
                shape = ((dynamic)pictureRange).InlineShapes.AddPicture(
                    session.ImagePath,
                    false,
                    true,
                    pictureRange);
            }
            else
            {
                shape = ((dynamic)insertionRange).InlineShapes.AddPicture(
                    session.ImagePath,
                    false,
                    true,
                    insertionRange);
            }

            ConfigureShape(
                shape,
                session,
                !display);
            if (display)
                TryReconcileEquationNumbers(doc);
            return Result(session, doc);
        }
        catch
        {
            TryDelete(shape);
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            ComRelease.Final(undoRecord);
            ComRelease.Final(shape);
            ComRelease.Final(paragraphRange);
            ComRelease.Final(paragraph);
            ComRelease.Final(insertionRange);
            ComRelease.Final(sourceRange);
            ComRelease.Final(selection);
            ComRelease.Final(document);
            ComRelease.Final(app);
        }
    }

    public int UpdateEquationNumbers()
    {
        object? app = null;
        object? document = null;
        object? undoRecord = null;
        try
        {
            app = RunningOfficeLocator.GetWordApplication();
            undoRecord = BeginUndoRecord(app, "VisualTeX Update Equation Numbers");
            document = ((dynamic)app).ActiveDocument;
            EnsureWritable((dynamic)document);
            return ReconcileEquationNumbers((dynamic)document);
        }
        finally
        {
            EndUndoRecord(undoRecord);
            ComRelease.Final(undoRecord);
            ComRelease.Final(document);
            ComRelease.Final(app);
        }
    }

    private static object? TryGetReusableDisplayParagraph(object sourceRange)
    {
        object? paragraphs = null;
        object? paragraph = null;
        object? paragraphRange = null;
        try
        {
            dynamic range = sourceRange;
            paragraphs = range.Paragraphs;
            if (Convert.ToInt32(((dynamic)paragraphs).Count) < 1)
                return null;
            paragraph = ((dynamic)paragraphs).Item(1);
            paragraphRange = ((dynamic)paragraph).Range;
            if (!ShouldReuseCurrentParagraphForDisplayFormula(
                    Convert.ToInt32(range.Start),
                    Convert.ToInt32(range.End),
                    ((dynamic)paragraphRange).Text as string))
                return null;

            var result = paragraph;
            paragraph = null;
            return result;
        }
        finally
        {
            ComRelease.Final(paragraphRange);
            ComRelease.Final(paragraph);
            ComRelease.Final(paragraphs);
        }
    }

    internal static bool ShouldReuseCurrentParagraphForDisplayFormula(
        int rangeStart,
        int rangeEnd,
        string? paragraphText)
    {
        if (rangeStart != rangeEnd || string.IsNullOrEmpty(paragraphText))
            return false;

        // Word exposes an empty paragraph as only its paragraph mark (\r), and
        // adds a cell marker (\a) when it is inside a table. Reusing that
        // paragraph prevents the user's empty caret line from being left above
        // a newly-created display-formula paragraph.
        return paragraphText.All(character => character is '\r' or '\a');
    }

    private static void TryReconcileEquationNumbers(dynamic document)
    {
        try { ReconcileEquationNumbers(document); }
        catch
        {
            // The formula image and its metadata are already durable. A user can
            // retry the numbering-only operation from the Ribbon without
            // creating a duplicate formula or losing a successful replacement.
        }
    }

    private static int ReconcileEquationNumbers(dynamic document)
    {
        var numberedFormulaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        object? inlineShapes = null;
        try
        {
            inlineShapes = document.InlineShapes;
            var count = Convert.ToInt32(((dynamic)inlineShapes).Count);
            for (var index = 1; index <= count; index++)
            {
                object? shape = null;
                try
                {
                    shape = ((dynamic)inlineShapes).Item(index);
                    var metadata = ReadMetadata(shape);
                    if (metadata is null || metadata.DisplayMode != "block")
                        continue;

                    if (metadata.Numbered)
                    {
                        ConfigureNumberedDisplayFormula(document, shape, metadata.FormulaId);
                        numberedFormulaIds.Add(metadata.FormulaId);
                    }
                    else
                    {
                        ConfigureUnnumberedDisplayFormula(document, shape, metadata.FormulaId);
                    }
                }
                finally { ComRelease.Final(shape); }
            }

            RemoveOrphanEquationNumbers(document, numberedFormulaIds);

            // Updating in document order makes every VisualTeX SEQ field derive
            // its result from the fields that precede it, producing (1), (2), ...
            for (var index = 1; index <= count; index++)
            {
                object? shape = null;
                try
                {
                    shape = ((dynamic)inlineShapes).Item(index);
                    var metadata = ReadMetadata(shape);
                    if (metadata?.DisplayMode == "block" && metadata.Numbered)
                        UpdateEquationNumberField(document, shape, metadata.FormulaId);
                }
                finally { ComRelease.Final(shape); }
            }
        }
        finally { ComRelease.Final(inlineShapes); }

        return numberedFormulaIds.Count;
    }

    private static void ConfigureNumberedDisplayFormula(
        dynamic document,
        object shape,
        string formulaId)
    {
        RemoveEquationNumber(document, formulaId);
        RemoveLeadingEquationTab(document, shape);
        ConfigureEquationParagraph(shape, numbered: true);
        InsertLeadingEquationTab(document, shape);
        InsertEquationNumber(document, shape, formulaId);
    }

    private static void ConfigureUnnumberedDisplayFormula(
        dynamic document,
        object shape,
        string formulaId)
    {
        RemoveEquationNumber(document, formulaId);
        RemoveLeadingEquationTab(document, shape);
        ConfigureEquationParagraph(shape, numbered: false);
    }

    private static void ConfigureEquationParagraph(object shape, bool numbered)
    {
        object? shapeRange = null;
        object? paragraphs = null;
        object? paragraph = null;
        object? paragraphRange = null;
        object? sections = null;
        object? section = null;
        object? pageSetup = null;
        object? format = null;
        object? tabStops = null;
        try
        {
            shapeRange = ((dynamic)shape).Range;
            paragraphs = ((dynamic)shapeRange).Paragraphs;
            paragraph = ((dynamic)paragraphs).Item(1);
            format = ((dynamic)paragraph).Format;
            ((dynamic)format).LeftIndent = 0f;
            ((dynamic)format).RightIndent = 0f;
            ((dynamic)format).FirstLineIndent = 0f;
            tabStops = ((dynamic)format).TabStops;
            ((dynamic)tabStops).ClearAll();

            if (!numbered)
            {
                ((dynamic)format).Alignment = WdAlignParagraphCenter;
                return;
            }

            var pageWidth = 612f;
            var leftMargin = 72f;
            var rightMargin = 72f;
            try
            {
                paragraphRange = ((dynamic)paragraph).Range;
                sections = ((dynamic)paragraphRange).Sections;
                if (Convert.ToInt32(((dynamic)sections).Count) > 0)
                {
                    section = ((dynamic)sections).Item(1);
                    pageSetup = ((dynamic)section).PageSetup;
                    pageWidth = Convert.ToSingle(((dynamic)pageSetup).PageWidth);
                    leftMargin = Convert.ToSingle(((dynamic)pageSetup).LeftMargin);
                    rightMargin = Convert.ToSingle(((dynamic)pageSetup).RightMargin);
                }
            }
            catch
            {
                // A standard 8.5-inch page with one-inch margins is a safe
                // fallback when a protected/custom story does not expose setup.
            }

            var positions = CalculateEquationTabStops(
                pageWidth,
                leftMargin,
                rightMargin,
                0,
                0);
            ((dynamic)format).Alignment = WdAlignParagraphLeft;
            ((dynamic)tabStops).Add(
                positions.Center,
                WdTabAlignmentCenter,
                WdTabLeaderSpaces);
            ((dynamic)tabStops).Add(
                positions.Right,
                WdTabAlignmentRight,
                WdTabLeaderSpaces);
        }
        finally
        {
            ComRelease.Final(tabStops);
            ComRelease.Final(format);
            ComRelease.Final(pageSetup);
            ComRelease.Final(section);
            ComRelease.Final(sections);
            ComRelease.Final(paragraphRange);
            ComRelease.Final(paragraph);
            ComRelease.Final(paragraphs);
            ComRelease.Final(shapeRange);
        }
    }

    internal static (float Center, float Right) CalculateEquationTabStops(
        float pageWidth,
        float leftMargin,
        float rightMargin,
        float leftIndent,
        float rightIndent)
    {
        var availableWidth = Math.Max(
            72f,
            pageWidth
                - Math.Max(0f, leftMargin)
                - Math.Max(0f, rightMargin)
                - Math.Max(0f, leftIndent)
                - Math.Max(0f, rightIndent));
        return (availableWidth / 2f, availableWidth);
    }

    internal static int CalculateEquationNumberFontPosition(
        float formulaHeightPoints,
        float numberFontSizePoints)
    {
        if (!float.IsFinite(formulaHeightPoints) || formulaHeightPoints <= 0)
            return 0;
        if (!float.IsFinite(numberFontSizePoints)
            || numberFontSizePoints <= 0
            || numberFontSizePoints > 256)
            numberFontSizePoints = 11f;

        // Word places an InlineShape and ordinary text on the same baseline.
        // Raising the number by half the difference between their visual boxes
        // aligns their centers while leaving the formula image itself untouched.
        return Math.Max(
            0,
            (int)Math.Round(
                (formulaHeightPoints - numberFontSizePoints) / 2f,
                MidpointRounding.AwayFromZero));
    }

    internal static (string Text, int FieldOffset) EquationNumberScaffold() =>
        ("\t()", 2);

    private static void InsertLeadingEquationTab(dynamic document, object shape)
    {
        object? shapeRange = null;
        object? insertion = null;
        try
        {
            shapeRange = ((dynamic)shape).Range;
            var start = Convert.ToInt32(((dynamic)shapeRange).Start);
            insertion = document.Range(start, start);
            ((dynamic)insertion).Text = "\t";
        }
        finally
        {
            ComRelease.Final(insertion);
            ComRelease.Final(shapeRange);
        }
    }

    private static void RemoveLeadingEquationTab(dynamic document, object shape)
    {
        object? shapeRange = null;
        object? paragraphs = null;
        object? paragraph = null;
        object? paragraphRange = null;
        object? preceding = null;
        try
        {
            shapeRange = ((dynamic)shape).Range;
            paragraphs = ((dynamic)shapeRange).Paragraphs;
            paragraph = ((dynamic)paragraphs).Item(1);
            paragraphRange = ((dynamic)paragraph).Range;
            var start = Convert.ToInt32(((dynamic)shapeRange).Start);
            var paragraphStart = Convert.ToInt32(((dynamic)paragraphRange).Start);
            if (start <= paragraphStart) return;
            preceding = document.Range(start - 1, start);
            if (string.Equals(((dynamic)preceding).Text as string, "\t", StringComparison.Ordinal))
                ((dynamic)preceding).Delete();
        }
        finally
        {
            ComRelease.Final(preceding);
            ComRelease.Final(paragraphRange);
            ComRelease.Final(paragraph);
            ComRelease.Final(paragraphs);
            ComRelease.Final(shapeRange);
        }
    }

    private static void InsertEquationNumber(dynamic document, object shape, string formulaId)
    {
        object? shapeRange = null;
        object? scaffoldRange = null;
        object? fieldRange = null;
        object? fields = null;
        object? field = null;
        object? fieldResult = null;
        object? bookmarkRange = null;
        object? bookmarks = null;
        try
        {
            shapeRange = ((dynamic)shape).Range;
            var suffixStart = Convert.ToInt32(((dynamic)shapeRange).End);
            var scaffold = EquationNumberScaffold();
            scaffoldRange = document.Range(suffixStart, suffixStart);
            ((dynamic)scaffoldRange).Text = scaffold.Text;

            // Insert the field into an already-complete pair of parentheses.
            // Field.Result.End is still inside Word's hidden field-end marker;
            // inserting ')' there makes it part of the result and the next
            // Field.Update removes it. Keeping ')' in the document first makes
            // it unambiguously external to the field across every update.
            var fieldStart = suffixStart + scaffold.FieldOffset;
            fieldRange = document.Range(fieldStart, fieldStart);
            fields = document.Fields;
            field = ((dynamic)fields).Add(
                fieldRange,
                WdFieldEmpty,
                $"SEQ {EquationSequenceName} \\* ARABIC",
                true);
            ((dynamic)field).Update();
            fieldResult = ((dynamic)field).Result;

            // The hidden field-end marker occupies one Word range position.
            // The pre-inserted ')' follows it, so +2 includes the full label.
            var labelEnd = Convert.ToInt32(((dynamic)fieldResult).End) + 2;
            bookmarkRange = document.Range(suffixStart, labelEnd);
            bookmarks = document.Bookmarks;
            ((dynamic)bookmarks).Add(EquationBookmarkName(formulaId), bookmarkRange);
            AlignEquationNumberVertically(bookmarkRange, shape);
        }
        finally
        {
            ComRelease.Final(bookmarks);
            ComRelease.Final(bookmarkRange);
            ComRelease.Final(fieldResult);
            ComRelease.Final(field);
            ComRelease.Final(fields);
            ComRelease.Final(fieldRange);
            ComRelease.Final(scaffoldRange);
            ComRelease.Final(shapeRange);
        }
    }

    private static void RemoveEquationNumber(dynamic document, string formulaId)
    {
        object? bookmarks = null;
        object? bookmark = null;
        object? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (!Convert.ToBoolean(((dynamic)bookmarks).Exists(name))) return;
            bookmark = ((dynamic)bookmarks).Item(name);
            range = ((dynamic)bookmark).Range;
            ((dynamic)range).Delete();
        }
        finally
        {
            ComRelease.Final(range);
            ComRelease.Final(bookmark);
            ComRelease.Final(bookmarks);
        }
    }

    private static void RemoveOrphanEquationNumbers(
        dynamic document,
        ISet<string> numberedFormulaIds)
    {
        object? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            var count = Convert.ToInt32(((dynamic)bookmarks).Count);
            for (var index = count; index >= 1; index--)
            {
                object? bookmark = null;
                object? range = null;
                try
                {
                    bookmark = ((dynamic)bookmarks).Item(index);
                    var name = ((dynamic)bookmark).Name as string;
                    if (!TryFormulaIdFromEquationBookmark(name, out var formulaId)
                        || numberedFormulaIds.Contains(formulaId))
                        continue;
                    range = ((dynamic)bookmark).Range;
                    ((dynamic)range).Delete();
                }
                finally
                {
                    ComRelease.Final(range);
                    ComRelease.Final(bookmark);
                }
            }
        }
        finally { ComRelease.Final(bookmarks); }
    }

    private static void UpdateEquationNumberField(
        dynamic document,
        object shape,
        string formulaId)
    {
        object? bookmarks = null;
        object? bookmark = null;
        object? range = null;
        object? fields = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (!Convert.ToBoolean(((dynamic)bookmarks).Exists(name))) return;
            bookmark = ((dynamic)bookmarks).Item(name);
            range = ((dynamic)bookmark).Range;
            fields = ((dynamic)range).Fields;
            var count = Convert.ToInt32(((dynamic)fields).Count);
            for (var index = 1; index <= count; index++)
            {
                object? field = null;
                object? code = null;
                try
                {
                    field = ((dynamic)fields).Item(index);
                    code = ((dynamic)field).Code;
                    if (IsVisualTeXSequenceFieldCode(((dynamic)code).Text as string))
                        ((dynamic)field).Update();
                }
                finally
                {
                    ComRelease.Final(code);
                    ComRelease.Final(field);
                }
            }
            AlignEquationNumberVertically(range, shape);
        }
        finally
        {
            ComRelease.Final(fields);
            ComRelease.Final(range);
            ComRelease.Final(bookmark);
            ComRelease.Final(bookmarks);
        }
    }

    private static void AlignEquationNumberVertically(object numberRange, object shape)
    {
        object? font = null;
        try
        {
            font = ((dynamic)numberRange).Font;
            var formulaHeight = Convert.ToSingle(((dynamic)shape).Height);
            var numberFontSize = 11f;
            try { numberFontSize = Convert.ToSingle(((dynamic)font).Size); }
            catch { }
            ((dynamic)font).Position = CalculateEquationNumberFontPosition(
                formulaHeight,
                numberFontSize);
        }
        finally { ComRelease.Final(font); }
    }

    internal static string EquationBookmarkName(string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var value))
            throw new InvalidOperationException("VisualTeX formulaId must be a UUID.");
        return $"{EquationBookmarkPrefix}{value:N}";
    }

    internal static bool TryFormulaIdFromEquationBookmark(
        string? bookmarkName,
        out string formulaId)
    {
        formulaId = string.Empty;
        if (string.IsNullOrWhiteSpace(bookmarkName)
            || !bookmarkName.StartsWith(EquationBookmarkPrefix, StringComparison.Ordinal)
            || !Guid.TryParseExact(
                bookmarkName.Substring(EquationBookmarkPrefix.Length),
                "N",
                out var value))
            return false;
        formulaId = value.ToString();
        return true;
    }

    internal static bool IsVisualTeXSequenceFieldCode(string? code) =>
        !string.IsNullOrWhiteSpace(code)
        && code.IndexOf(
            $"SEQ {EquationSequenceName}",
            StringComparison.OrdinalIgnoreCase) >= 0;

    private static object? FindFormula(dynamic document, string formulaId)
    {
        object? inlineShapes = null;
        try
        {
            inlineShapes = document.InlineShapes;
            var count = Convert.ToInt32(((dynamic)inlineShapes).Count);
            for (var index = 1; index <= count; index++)
            {
                object? candidate = null;
                try
                {
                    candidate = ((dynamic)inlineShapes).Item(index);
                    var metadata = ReadMetadata(candidate);
                    if (metadata?.FormulaId == formulaId)
                    {
                        var result = candidate;
                        candidate = null;
                        return result;
                    }
                }
                finally { ComRelease.Final(candidate); }
            }
            return null;
        }
        finally { ComRelease.Final(inlineShapes); }
    }

    private static FormulaMetadata? ReadMetadata(object shape)
    {
        dynamic candidate = shape;
        string? encoded = null;
        try { encoded = candidate.AlternativeText as string; } catch { }
        if (string.IsNullOrWhiteSpace(encoded))
        {
            try { encoded = candidate.Title as string; } catch { }
        }
        return MetadataCodec.Decode(encoded);
    }

    private static void ConfigureShape(
        object shape,
        SessionInfo session,
        bool alignInline)
    {
        dynamic target = shape;
        var encoded = MetadataCodec.Encode(session.Metadata);
        // Always size an edited formula from the current export. Reusing the old
        // InlineShape bounds makes a wider edit shrink to fit the previous box,
        // then compounds that shrink on every subsequent edit.
        var size = FitImage(session.ImagePath, session.Width, session.Height);
        target.LockAspectRatio = MsoTrue;
        target.Width = size.Width;
        target.Title = encoded;
        target.AlternativeText = encoded;
        if (alignInline)
            ApplyInlineBaseline(shape, size.Height, session.Height, session.Baseline);
    }

    private static object? BeginUndoRecord(object application, string name)
    {
        object? undoRecord = null;
        try
        {
            undoRecord = ((dynamic)application).UndoRecord;
            ((dynamic)undoRecord).StartCustomRecord(name);
            return undoRecord;
        }
        catch
        {
            ComRelease.Final(undoRecord);
            return null;
        }
    }

    private static void EndUndoRecord(object? undoRecord)
    {
        if (undoRecord is null) return;
        try { ((dynamic)undoRecord).EndCustomRecord(); } catch { }
    }

    private static void ApplyInlineBaseline(
        object shape,
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline)
    {
        object? range = null;
        object? font = null;
        try
        {
            range = ((dynamic)shape).Range;
            font = ((dynamic)range).Font;
            ((dynamic)font).Position = WordInlineAlignment.CalculateFontPosition(
                actualHeightPoints,
                exportedHeight,
                exportedBaseline);
        }
        finally
        {
            ComRelease.Final(font);
            ComRelease.Final(range);
        }
    }

    internal static SizeF FitImage(string path, float maxWidth, float maxHeight)
    {
        using var image = Image.FromFile(path);
        return FitImage(image.Width, image.Height, maxWidth, maxHeight);
    }

    internal static SizeF FitImage(
        int pixelWidth,
        int pixelHeight,
        float maxWidth,
        float maxHeight)
    {
        var ratio = pixelWidth / (float)Math.Max(1, pixelHeight);
        var width = Math.Max(12f, maxWidth);
        var height = width / ratio;
        if (maxHeight > 0 && height > maxHeight)
        {
            height = maxHeight;
            width = height * ratio;
        }
        return new SizeF(width, height);
    }

    private static OfficeObjectResult Result(SessionInfo session, dynamic document) =>
        new()
        {
            FormulaId = session.FormulaId,
            DocumentId = DocumentIdentity(document),
            ObjectId = session.FormulaId,
        };

    private static string RangeReference(object range)
    {
        dynamic value = range;
        return $"{RangeReferencePrefix}{Convert.ToInt32(value.Start)}:{Convert.ToInt32(value.End)}";
    }

    private static object ResolveSourceRange(
        dynamic document,
        string? sourceObjectId,
        object selection)
    {
        if (!TryParseRangeReference(sourceObjectId, out var start, out var end))
            return ((dynamic)selection).Range;
        object? content = null;
        try
        {
            content = document.Content;
            var maximum = Convert.ToInt32(((dynamic)content).End);
            if (start < 0 || end < start || end > maximum)
                throw new InvalidOperationException(
                    "The Word insertion range selected when the formula editor opened is no longer valid.");
            return document.Range(start, end);
        }
        finally { ComRelease.Final(content); }
    }

    private static bool TryParseRangeReference(
        string? value,
        out int start,
        out int end)
    {
        start = 0;
        end = 0;
        if (string.IsNullOrWhiteSpace(value)
            || !value.StartsWith(RangeReferencePrefix, StringComparison.Ordinal))
            return false;
        var payload = value.Substring(RangeReferencePrefix.Length);
        var separator = payload.IndexOf(':');
        if (separator <= 0 || separator >= payload.Length - 1) return false;
        return int.TryParse(payload.Substring(0, separator), out start)
            && int.TryParse(payload.Substring(separator + 1), out end);
    }

    private static void EnsureSourceDocument(dynamic document, string? expectedIdentity)
    {
        if (string.IsNullOrWhiteSpace(expectedIdentity)) return;
        var actual = DocumentIdentity(document);
        if (!string.Equals(actual, expectedIdentity, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "The active Word document changed while the VisualTeX editor was open.");
    }

    private static string DocumentIdentity(dynamic document)
    {
        try
        {
            var fullName = document.FullName as string;
            if (!string.IsNullOrWhiteSpace(fullName)) return fullName;
        }
        catch { }
        return document.Name as string ?? "Word";
    }

    private static bool IsReadOnly(dynamic document)
    {
        try { return Convert.ToBoolean(document.ReadOnly); }
        catch { return false; }
    }

    private static void EnsureWritable(dynamic document)
    {
        if (IsReadOnly(document))
            throw new UnauthorizedAccessException("The active Word document is read-only.");
    }

    private static void TryDelete(object? shape)
    {
        if (shape is null) return;
        try { ((dynamic)shape).Delete(); } catch { }
    }
}
