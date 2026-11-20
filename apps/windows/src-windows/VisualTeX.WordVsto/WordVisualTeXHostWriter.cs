using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// The only low-level writer for VisualTeX native OLE hosts in the rebuilt core.
/// It does not create numbering, move paragraphs, scan the document, or insert
/// persistent typing sentinels.
/// </summary>
internal static class WordVisualTeXHostWriter
{
    internal static WordFormulaHostWriteResult Insert(
        Document document,
        Range insertion,
        WordFormulaHostWriteRequest request,
        bool removeOmmlBoundaryGuard = false)
    {
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.VisualTeXHostWriter.Insert");
        if (request.Kind != WordFormulaHostKind.VisualTeX)
            throw new ArgumentException("VisualTeX writer received a non-VisualTeX request.");
        if (request.Metadata is null)
            throw new InvalidDataException(
                "VisualTeX insertion requires authoritative formula metadata.");
        if (string.IsNullOrWhiteSpace(request.PngPath)
            || string.IsNullOrWhiteSpace(request.EmfPath))
            throw new InvalidDataException(
                "VisualTeX insertion requires PNG and EMF preview paths.");

        var metadata = request.Metadata;
        if (!Guid.TryParse(metadata.FormulaId, out _))
            metadata.FormulaId = Guid.NewGuid().ToString("D");
        request.FormulaId = metadata.FormulaId;
        metadata.DisplayMode = request.DisplayMode;
        metadata.Numbered = request.Numbered;
        metadata.Validate();

        InlineShape? shape = null;
        Range? shapeRange = null;
        IVisualTeXFormulaObject? initializedFormula = null;
        try
        {
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            using (WordOperationMetrics.Measure(
                       "FormatConversion.VisualTeXHostWriter.AddOleObject"))
            {
                shape = AddOleObjectLocal(document, insertion);
            }
            using (WordOperationMetrics.Measure(
                       "FormatConversion.VisualTeXHostWriter.InitializeOle"))
            {
                initializedFormula = InitializeOle(shape, metadata, request.EmfPath!, request.PngPath!);
            }

            using (WordOperationMetrics.Measure(
                       "FormatConversion.VisualTeXHostWriter.ApplyGeometry"))
            {
                WordFormulaHostLayout.ApplyVisualTeXGeometry(
                    shape,
                    metadata,
                    request.WidthPoints > 0
                        ? request.WidthPoints
                        : 1f,
                    request.HeightPoints > 0
                        ? request.HeightPoints
                        : 1f,
                    request.ExportedHeightPixels,
                    request.ExportedBaselinePixels,
                    deferDisplayPosition: true);
            }

            if (!string.Equals(
                    metadata.DisplayMode,
                    "inline",
                    StringComparison.OrdinalIgnoreCase))
            {
                Range? displayRange = null;
                try
                {
                    displayRange = shape.Range;
                    using (WordOperationMetrics.Measure(
                               "FormatConversion.VisualTeXHostWriter.ConfigureDisplay"))
                    {
                        // ApplyVisualTeXGeometry immediately above already wrote
                        // this exact final height. Reading InlineShape.Height back
                        // here forces Word to resolve live layout for every display
                        // host. Reuse the authoritative request geometry instead.
                        var displayHeight = request.HeightPoints > 0
                            ? request.HeightPoints
                            : 1f;
                        WordFormulaHostLayout.ConfigureVisualTeXDisplayParagraph(
                            document,
                            displayRange,
                            displayHeight,
                            metadata,
                            request.DisplayCenterTabPoints,
                            request.DisplayRightTabPoints);
                    }
                    // Display baseline measurement can enrich Word-side preview
                    // metrics. Re-cache those values after the canonical layout is
                    // complete without reopening/re-writing the embedded OLE.
                    using (WordOperationMetrics.Measure(
                               "FormatConversion.VisualTeXHostWriter.CacheDisplayMetadata"))
                    {
                        WordFormulaMetadataReader.CacheMetadata(
                            shape,
                            metadata);
                    }
                }
                finally { Release(displayRange); }
            }

            using (WordOperationMetrics.Measure(
                       "FormatConversion.VisualTeXHostWriter.BindIdentity"))
            {
                WordFormulaIdentityStore.BindVisualTeX(shape, metadata.FormulaId);
            }
            using (WordOperationMetrics.Measure(
                       "FormatConversion.VisualTeXHostWriter.ReadbackValidate"))
            {
                ValidateInsertedHost(initializedFormula, metadata);
            }

            using (WordOperationMetrics.Measure(
                       "FormatConversion.VisualTeXHostWriter.CaptureDescriptorRange"))
            {
                shapeRange = shape.Range.Duplicate;
            }
            bool withinTable;
            using (WordOperationMetrics.Measure(
                       "FormatConversion.VisualTeXHostWriter.WithinTable"))
            {
                withinTable = IsWithinTable(shapeRange);
            }
            if (request.BatchMutation && metadata.Numbered
                && string.Equals(metadata.DisplayMode, "block", StringComparison.OrdinalIgnoreCase)
                && !withinTable && request.PreparedVisualTeXNumbering is not null)
            {
                request.PreparedVisualTeXNumbering.WrittenBodyLayout =
                    new WordVisualTeXParagraphNumbering.WrittenBodyLayout(shapeRange);
            }
            var result = new WordFormulaHostWriteResult
            {
                Host = new WordFormulaHostDescriptor
                {
                    Kind = WordFormulaHostKind.VisualTeX,
                    Range = new WordFormulaRangeAddress
                    {
                        StoryType = shapeRange.StoryType,
                        Start = shapeRange.Start,
                        End = shapeRange.End,
                    },
                    DisplayMode = metadata.DisplayMode,
                    FormulaId = metadata.FormulaId,
                    Metadata = metadata,
                    MetadataAuthoritative = true,
                    Numbering = new WordFormulaNumberingDescriptor
                    {
                        Numbered = metadata.Numbered,
                        FormulaId = metadata.FormulaId,
                    },
                    WithinTable = withinTable,
                },
            };
            if (removeOmmlBoundaryGuard)
                result.Host = WordOmmlHostWriter.RemovePlainTextBoundaryGuard(document, result.Host, shape);
            using (WordOperationMetrics.Measure("FormatConversion.VisualTeXHostWriter.CloseOle"))
                FormulaOleInterop.CloseAfterSave(initializedFormula);
            return result;
        }
        catch
        {
            try { shape?.Delete(); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            if (shape is not null)
                WordFormulaIdentityStore.RemoveVisualTeX(document, metadata.FormulaId);
            throw;
        }
        finally
        {
            Release(initializedFormula);
            Release(shapeRange);
            Release(shape);
        }
    }

    internal static WordFormulaHostDescriptor UpdateFontSize(
        Document document,
        WordFormulaHostDescriptor host,
        double targetFontSizePoints)
    {
        if (host.Kind != WordFormulaHostKind.VisualTeX)
            throw new ArgumentException(
                "VisualTeX font-size update received another host kind.");

        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            (shape, shapeRange) = FindOwnShape(document, host);
            FormulaMetadata metadata;
            using (WordOperationMetrics.Measure("FontSize.Ole.ReadPayload"))
                metadata =
                    WordFormulaMetadataReader.TryReadEmbeddedNativeOle(shape)
                    ?? throw new InvalidDataException(
                        "The VisualTeX host has no authoritative embedded metadata.");
            if (!string.IsNullOrWhiteSpace(host.FormulaId)
                && !string.Equals(
                    host.FormulaId,
                    metadata.FormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "The VisualTeX host identity changed before its font-size update.");

            ApplyFontSizeGeometry(document, shape, shapeRange, metadata, targetFontSizePoints);
            using (WordOperationMetrics.Measure("FontSize.Ole.WritePayload"))
                WordFormulaMetadataReader.Write(
                    shape,
                    metadata);

            return FontSizeResult(
                shapeRange,
                metadata,
                host,
                WordFormulaNumberingResolver.ResolveLocal(document, host));
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
        }
    }

    // A VisualTeX object whose font size a batch changes without reading its
    // payload: the object, where it is, and Word's cached copy of its metadata.
    internal sealed class CachedFontSizeTarget : IDisposable
    {
        internal CachedFontSizeTarget(InlineShape shape, Range shapeRange, FormulaMetadata metadata)
        {
            Shape = shape;
            ShapeRange = shapeRange;
            Metadata = metadata;
        }

        internal InlineShape Shape { get; }
        internal Range ShapeRange { get; }
        internal FormulaMetadata Metadata { get; }
        // The size the batch gives the object, and whether it is a display formula.
        internal (float Width, float Height) Size { get; set; }
        internal bool Display { get; set; }

        public void Dispose()
        {
            Release(ShapeRange);
            Release(Shape);
        }
    }

    // The object by Word's cached metadata, when that cache is this formula's
    // and describes the object as it is (null otherwise: the caller then reads
    // the payload, UpdateFontSize). Nothing is written: a size write before the
    // batch's undo record would make every move in it slower (PositionCached).
    internal static CachedFontSizeTarget? PrepareCachedFontSize(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (host.Kind != WordFormulaHostKind.VisualTeX
            || !Guid.TryParse(host.FormulaId, out _))
            return null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            (shape, shapeRange) = FindOwnShape(document, host);
            var metadata = WordFormulaMetadataReader.TryReadCachedPreview(shape);
            var width = shape.Width;
            var height = shape.Height;
            if (metadata is null
                || !string.Equals(metadata.FormulaId, host.FormulaId, StringComparison.OrdinalIgnoreCase)
                || !FormulaFontSize.CacheMatchesOleSize(metadata, width, height))
                return null;
            var target = new CachedFontSizeTarget(shape, shapeRange, metadata);
            shape = null;
            shapeRange = null;
            return target;
        }
        catch (Exception error) when (error is COMException or InvalidDataException or InvalidOperationException)
        {
            return null;
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
        }
    }

    // Inside the batch's undo record the three steps run for all objects one
    // after the other, positions and numbers before sizes. Measured in Word
    // (100 display objects, one custom undo record): moving an object on its
    // line costs 8 ms while Word has not resized it in this session, and from
    // 20 ms up to 500 ms per object once it has (Word runs the object's server
    // for every size write; nothing the server does changes that). So every
    // object is moved before any is resized, an object whose offset stays the
    // same is not written (WordNumberedDisplayLayout.TryMoveEmbeddedOleObject),
    // and the sizes come last.

    // 1. The object's position on its line for the new size, calculated from
    // the size it is about to get.
    internal static WordFormulaHostDescriptor PositionCached(
        Document document,
        WordFormulaHostDescriptor host,
        CachedFontSizeTarget target,
        double targetFontSizePoints)
    {
        var metadata = target.Metadata;
        var normalized = FormulaFontSize.Normalize(targetFontSizePoints);
        metadata.FontSizePt = normalized;
        metadata.UpdatedWithVersion = "1.2.7";
        metadata.UpdatedAt = DateTimeOffset.UtcNow.ToString("O");
        target.Size = FormulaFontSize.OleSizeAt(metadata, normalized);
        target.Display = !string.Equals(metadata.DisplayMode, "inline", StringComparison.OrdinalIgnoreCase);
        if (target.Display)
        {
            if (!WordNumberedDisplayLayout.TryMoveEmbeddedOleObject(
                    target.ShapeRange,
                    WordNumberedDisplayLayout.DisplayOlePosition(target.Size.Height, metadata)))
                WordEquationParagraphLayout.AlignVisualTeXDisplayHostBaseline(
                    document, target.ShapeRange, target.Size.Height, metadata);
        }
        else
            WordFormulaHostLayout.PositionVisualTeXInline(
                target.Shape, target.ShapeRange, metadata, target.Size.Height);
        return FontSizeResult(target.ShapeRange, metadata, host, host.Numbering);
    }

    // 2. The object's size in Word.
    internal static void ResizeCached(CachedFontSizeTarget target) =>
        WordFormulaHostLayout.SetVisualTeXSize(target.Shape, target.Size.Width, target.Size.Height);

    // 3. The new size in Word's cached metadata, with the size Word gave an
    // inline object. The payload keeps the old size until the formula is next
    // written (FormulaFontSize.WithDeferredFontSize reads the cached one
    // meanwhile).
    internal static void CacheResized(CachedFontSizeTarget target)
    {
        var metadata = target.Metadata;
        if (target.Display)
        {
            metadata.WordInlineOleWidthPt = null;
            metadata.WordInlineOleHeightPt = null;
            metadata.WordInlineOlePositionPt = null;
        }
        else
        {
            metadata.WordInlineOleWidthPt = target.Shape.Width;
            metadata.WordInlineOleHeightPt = target.Shape.Height;
            metadata.WordInlineSourceBottomWhitespacePt = null;
        }
        metadata.Validate();
        WordFormulaMetadataReader.CacheMetadata(target.Shape, metadata);
    }

    private static (InlineShape Shape, Range ShapeRange) FindOwnShape(
        Document document,
        WordFormulaHostDescriptor host)
    {
        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            range = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            shapes = range.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                var candidate = shapes[index];
                Range? candidateRange = null;
                try
                {
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate))
                        continue;
                    candidateRange = candidate.Range.Duplicate;
                    if (!WordFormulaHostSemanticReader.SameAddress(
                            candidateRange,
                            host.Range))
                        continue;
                    if (shape is not null)
                        throw new InvalidDataException(
                            "The VisualTeX range contains multiple native OLE hosts.");
                    shape = candidate;
                    candidate = null;
                    shapeRange = candidateRange;
                    candidateRange = null;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }

            if (shape is null || shapeRange is null)
                throw new InvalidDataException(
                    "The VisualTeX host moved before its font-size update.");
            var found = (shape, shapeRange);
            shape = null;
            shapeRange = null;
            return found;
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(range);
        }
    }

    // The object's size and position in Word for a font size, and the size in
    // its metadata.
    private static void ApplyFontSizeGeometry(
        Document document,
        InlineShape shape,
        Range shapeRange,
        FormulaMetadata metadata,
        double targetFontSizePoints)
    {
        var normalized =
            FormulaFontSize.Normalize(
                targetFontSizePoints);
        metadata.FontSizePt = normalized;
        metadata.UpdatedWithVersion = "1.2.7";
        metadata.UpdatedAt =
            DateTimeOffset.UtcNow.ToString("O");

        var size =
            FormulaFontSize.OleSizeAt(
                metadata,
                normalized);
        var display = !string.Equals(
            metadata.DisplayMode,
            "inline",
            StringComparison.OrdinalIgnoreCase);
        WordFormulaHostLayout.ApplyVisualTeXGeometry(
            shape,
            metadata,
            size.Width,
            size.Height,
            (float)(metadata.RenderHeightPx ?? 0d),
            metadata.Baseline.HasValue
                ? (float?)metadata.Baseline.Value
                : null,
            deferDisplayPosition: display);
        // A display OLE is lowered so the formula's own baseline sits on
        // the line that carries its TABs and number. That offset depends on
        // the height, so it is recalculated here as insertion does; the
        // paragraph's tab stops are left as they are.
        if (display)
            WordEquationParagraphLayout.AlignVisualTeXDisplayHostBaseline(
                document,
                shapeRange,
                size.Height,
                metadata);
    }

    private static WordFormulaHostDescriptor FontSizeResult(
        Range shapeRange,
        FormulaMetadata metadata,
        WordFormulaHostDescriptor host,
        WordFormulaNumberingDescriptor numbering) =>
        new()
        {
            Kind = WordFormulaHostKind.VisualTeX,
            Range = new WordFormulaRangeAddress
            {
                StoryType = shapeRange.StoryType,
                Start = shapeRange.Start,
                End = shapeRange.End,
            },
            DisplayMode = metadata.DisplayMode,
            FormulaId = metadata.FormulaId,
            Metadata = metadata,
            MetadataAuthoritative = true,
            Numbering = numbering,
            WithinTable = host.WithinTable,
            Latex = metadata.Latex,
        };

    internal static Range DeleteExactToExternalDisplayBoundary(
        Document document,
        WordFormulaHostDescriptor host,
        WordNativeConversionSource? preparedSource = null)
    {
        if (host.Kind != WordFormulaHostKind.VisualTeX
            || !host.Display)
            throw new ArgumentException(
                "VisualTeX external display deletion requires one display VisualTeX host.");

        Range? hostRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? leadingTab = null;
        Range? suffix = null;
        Range? deletedAnchor = null;
        try
        {
            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                return DeleteExact(
                    document,
                    host,
                    preparedSource);
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;

            // A numbered VisualTeX host is detached before an external
            // format writer runs. Its canonical unnumbered intermediate retains
            // exactly one center-tab character immediately before the OLE. That
            // tab belongs to VisualTeX's paragraph scaffold; leaving it behind
            // makes MathType/OMML inherit a second layout token (and Word can
            // materialize it as TAB + manual line break). Remove it only when it
            // is the first character of the same formula paragraph and there is
            // no user content after the OLE. The same invariant is valid inside
            // an ordinary user-table cell; paragraph/cell markers are structural
            // and are explicitly accepted by ContainsOnlyStructuralText below.
            if (hostRange.Start != paragraphRange.Start + 1)
                return DeleteExact(
                    document,
                    host,
                    preparedSource);
            leadingTab = document.Range(
                paragraphRange.Start,
                hostRange.Start);
            if (!string.Equals(
                    leadingTab.Text,
                    "\t",
                    StringComparison.Ordinal))
                return DeleteExact(
                    document,
                    host,
                    preparedSource);

            var bodyEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            if (hostRange.End > bodyEnd)
                throw new InvalidDataException(
                    "The VisualTeX display host escaped its paragraph body before external conversion.");
            suffix = document.Range(
                hostRange.End,
                bodyEnd);
            // User text after the equation: the tab may be the user's too, so
            // remove only the object itself, as the branches above do.
            if (!ContainsOnlyStructuralText(
                    suffix.Text))
                return DeleteExact(
                    document,
                    host,
                    preparedSource);

            var insertionStart = paragraphRange.Start;
            var expectedSuffix = suffix.Text ?? string.Empty;
            deletedAnchor = DeleteExact(
                document,
                host,
                preparedSource);

            Release(leadingTab);
            leadingTab = document.Range(
                insertionStart,
                insertionStart + 1);
            if (string.Equals(
                    leadingTab.Text,
                    "\t",
                    StringComparison.Ordinal))
                leadingTab.Delete();
            else
            {
                // Word's spacing cleanup for InlineShape.Delete can take the tab
                // as well; then the paragraph body holds only the structural
                // suffix. Anything else means another object was removed.
                var body = ReadParagraphBodyFrom(
                    document,
                    insertionStart);
                if (body is null
                    || !ContainsOnlyStructuralText(
                        body))
                    throw new InvalidDataException(
                        "The VisualTeX center tab moved while deleting the external source host. "
                        + $"Expected {DescribeText("\t" + expectedSuffix)} at {insertionStart}, found {DescribeText(body)}.");
                WordDoubleClickHook.TraceMessage(
                    $"visualtex-display-delete-tab-removed-by-word at={insertionStart} suffix={DescribeText(expectedSuffix)}");
            }

            return document.Range(
                insertionStart,
                insertionStart);
        }
        finally
        {
            Release(deletedAnchor);
            Release(suffix);
            Release(leadingTab);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(hostRange);
        }
    }

    internal static Range DeleteExact(
        Document document,
        WordFormulaHostDescriptor host,
        WordNativeConversionSource? preparedSource = null)
    {
        if (host.Kind != WordFormulaHostKind.VisualTeX)
            throw new ArgumentException("VisualTeX delete received another host kind.");

        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        Range? anchor = null;
        var ownsShape = true;
        try
        {
            // The prepared owner was bound before the batch began. Delete it only
            // while it still is the OLE at the host's address; otherwise find the
            // OLE at that address as an unprepared delete does.
            if (preparedSource?.VisualTeXShape is { } prepared)
            {
                var preparedRange = prepared.Range;
                if (WordFormulaHostSemanticReader.SameAddress(
                        preparedRange,
                        host.Range))
                {
                    shape = prepared;
                    ownsShape = false;
                    shapeRange = preparedRange;
                }
                else
                {
                    WordDoubleClickHook.TraceMessage(
                        $"visualtex-delete-prepared-owner-moved owner={preparedRange.Start}-{preparedRange.End} host={host.Range.Start}-{host.Range.End}");
                    Release(preparedRange);
                }
            }
            if (shape is null)
            {
            range = WordFormulaHostSemanticReader.CreateRange(document, host.Range);
            shapes = range.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                var candidate = shapes[index];
                Range? candidateRange = null;
                try
                {
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate))
                        continue;
                    candidateRange = candidate.Range.Duplicate;
                    if (!WordFormulaHostSemanticReader.SameAddress(
                            candidateRange,
                            host.Range))
                        continue;

                    if (shape is not null)
                        throw new InvalidDataException(
                            "The VisualTeX delete range contains multiple matching hosts.");
                    shape = candidate;
                    candidate = null;
                    shapeRange = candidateRange;
                    candidateRange = null;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }

            }
            if (shape is null || shapeRange is null)
                throw new InvalidDataException(
                    "The VisualTeX source host moved before deletion.");

            var embedded = (ownsShape
                ? WordFormulaMetadataReader.TryReadEmbeddedNativeOle(shape)
                : preparedSource!.Host.Metadata)
                ?? throw new InvalidDataException(
                    "The VisualTeX source host has no authoritative embedded metadata.");
            if (!string.IsNullOrWhiteSpace(host.FormulaId)
                && !string.Equals(
                    embedded.FormulaId,
                    host.FormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "The VisualTeX source identity changed before deletion.");

            anchor = shapeRange.Duplicate;
            anchor.Collapse(WdCollapseDirection.wdCollapseStart);
            WordFormulaIdentityStore.RemoveVisualTeX(
                document,
                embedded.FormulaId);
            // InlineShape.Delete applies Word's word-boundary spacing cleanup.
            // Exact range replacement removes only the OLE field, preserving the
            // neighboring prose space as ordinary source content.
            if (host.Display) shape.Delete();
            else shapeRange.Text = string.Empty;

            var returned = anchor;
            anchor = null;
            return returned;
        }
        finally
        {
            Release(anchor);
            Release(shapeRange);
            if (ownsShape) Release(shape);
            Release(shapes);
            Release(range);
        }
    }

    // The text from position to the end of its paragraph (paragraph mark
    // excluded), or null when position no longer starts a paragraph.
    private static string? ReadParagraphBodyFrom(
        Document document,
        int position)
    {
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? body = null;
        try
        {
            probe = document.Range(position, position);
            paragraphs = probe.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.Start != position)
                return null;
            body = document.Range(
                position,
                Math.Max(position, paragraphRange.End - 1));
            return body.Text ?? string.Empty;
        }
        finally
        {
            Release(body);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    private static string DescribeText(string? text)
    {
        if (text is null)
            return "no paragraph start";
        if (text.Length == 0)
            return "an empty paragraph";
        var shown = text.Length > 12 ? text.Substring(0, 12) : text;
        return string.Join(" ", shown.Select(character =>
                $"U+{(int)character:X4}"))
            + (text.Length > 12 ? $" ... ({text.Length} characters)" : string.Empty);
    }

    private static bool ContainsOnlyStructuralText(
        string? text)
    {
        if (string.IsNullOrEmpty(text))
            return true;
        foreach (var character in text!)
        {
            if (character is '\r' or '\n' or '\t' or '\v' or '\a'
                || char.IsWhiteSpace(character))
                continue;
            return false;
        }
        return true;
    }

    private static InlineShape AddOleObjectLocal(
        Document document,
        Range insertion)
    {
        var insertionStart = insertion.Start;
        Range? rightWitness = null;
        try
        {
            if (insertion.Start == insertion.End)
            {
                Range? content = null;
                try
                {
                    content = document.Content;
                    if (insertionStart < content.End)
                    {
                        rightWitness = insertion.Duplicate;
                        rightWitness.SetRange(
                            insertionStart,
                            insertionStart + 1);
                    }
                }
                catch
                {
                    Release(rightWitness);
                    rightWitness = null;
                }
                finally { Release(content); }
            }

            return document.InlineShapes.AddOLEObject(
                ClassType: FormulaOleContract.ProgId,
                LinkToFile: false,
                DisplayAsIcon: false,
                Range: insertion);
        }
        catch (COMException error) when (
            error.HResult == unchecked((int)0x800A1066))
        {
            var recovered = TryRecoverLocalMaterializedOle(
                document,
                insertionStart,
                rightWitness);
            if (recovered is null) throw;
            return recovered;
        }
        finally { Release(rightWitness); }
    }

    private static InlineShape? TryRecoverLocalMaterializedOle(
        Document document,
        int insertionStart,
        Range? rightWitness)
    {
        if (rightWitness is null) return null;

        Range? probe = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            var insertedEnd = rightWitness.End - 1;
            if (insertedEnd <= insertionStart) return null;

            Range? content = null;
            try
            {
                content = document.Content;
                var end = Math.Min(
                    content.End,
                    Math.Min(insertedEnd, insertionStart + 128));
                if (end <= insertionStart) return null;
                probe = document.Range(insertionStart, end);
            }
            finally { Release(content); }

            shapes = probe.InlineShapes;
            InlineShape? match = null;
            for (var index = 1; index <= shapes.Count; index++)
            {
                shape = shapes[index];
                if (!WordFormulaMetadataReader.IsNativeOle(shape))
                {
                    Release(shape);
                    shape = null;
                    continue;
                }

                shapeRange = shape.Range.Duplicate;
                if (shapeRange.Start < insertionStart
                    || shapeRange.End > insertedEnd
                    || !IsLocalVisualTeXEmbedField(probe, shapeRange))
                {
                    Release(shapeRange);
                    shapeRange = null;
                    Release(shape);
                    shape = null;
                    continue;
                }

                if (match is not null)
                {
                    Release(match);
                    return null;
                }
                match = shape;
                shape = null;
                Release(shapeRange);
                shapeRange = null;
            }
            return match;
        }
        catch { return null; }
        finally
        {
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(probe);
        }
    }

    private static bool IsLocalVisualTeXEmbedField(
        Range probe,
        Range shapeRange)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        try
        {
            fields = probe.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                field = fields[index];
                if (field.Type != WdFieldType.wdFieldEmbed)
                {
                    Release(field);
                    field = null;
                    continue;
                }

                code = field.Code;
                result = field.Result;
                var ownsShape =
                    result.Start <= shapeRange.Start
                    && result.End >= shapeRange.End;
                var namesVisualTeX =
                    (code.Text ?? string.Empty).IndexOf(
                        FormulaOleContract.ProgId,
                        StringComparison.OrdinalIgnoreCase) >= 0;
                if (ownsShape && namesVisualTeX) return true;

                Release(result); result = null;
                Release(code); code = null;
                Release(field); field = null;
            }
            return false;
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static IVisualTeXFormulaObject InitializeOle(
        InlineShape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            format = shape.OLEFormat;
            oleObject = WordOleObjectAccessor.GetRunningObject(format);
            var formula = oleObject as IVisualTeXFormulaObject
                ?? throw new InvalidOperationException(
                    "The inserted object does not expose the VisualTeX OLE interface.");
            FormulaOleInterop.Initialize(formula, metadata, emfPath, pngPath);
            // The writer owns this same object through geometry, metadata
            // readback and final save. Closing here forced immediate reactivation.
            oleObject = null;
            return formula;
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private static void ValidateInsertedHost(
        IVisualTeXFormulaObject formula,
        FormulaMetadata expected)
    {
        var metadata = FormulaOleInterop.ReadMetadata(formula);
        if (!string.Equals(
                metadata.FormulaId,
                expected.FormulaId,
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                "The inserted VisualTeX OLE returned a different FormulaId.");
        if (!string.Equals(
                metadata.Latex,
                expected.Latex,
                StringComparison.Ordinal))
            throw new InvalidDataException(
                "The inserted VisualTeX OLE returned different LaTeX source.");
        if (!string.Equals(
                metadata.DisplayMode,
                expected.DisplayMode,
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                "The inserted VisualTeX OLE returned a different display mode.");
    }

    // Table ownership is structural. Word's Information getter is backed by
    // pagination and can lay out the whole rest of the document per host.
    private static bool IsWithinTable(Range range) =>
        WordFormulaHostLookup.IsWhollyInsideTable(range);

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
    }
}
