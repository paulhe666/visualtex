using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    internal float SetSelectedFormulaFontSizeCore(
        double requestedFontSizePt)
    {
        var selected = ReadSelection();
        if (selected.Metadata is null
            || string.IsNullOrWhiteSpace(selected.ObjectId)
            || string.IsNullOrWhiteSpace(selected.ObjectMode))
            throw new InvalidOperationException(
                "请先选择一个公式。");

        var target =
            FormulaFontSize.Normalize(
                requestedFontSizePt);

        if (string.Equals(
                selected.ObjectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal))
        {
            Document? mathTypeDocument = null;
            InlineShape? shape = null;
            try
            {
                mathTypeDocument = _application.ActiveDocument
                    ?? throw new InvalidOperationException(
                        "No active Word document.");
                shape = FindMathTypeOleByRange(
                        mathTypeDocument,
                        selected.ObjectId,
                        allowGlobalFallback: false)
                    ?? throw new InvalidOperationException(
                        "The selected MathType equation no longer exists at the captured Word location.");
                var sourceFragment =
                    MathTypeWordOpenXml.Read(shape);
                var sourceMathMl =
                    MathTypeOleStorage.ReadMathMl(
                        sourceFragment.CompoundFile);
                var metadata = selected.Metadata;
                var resizeSession =
                    new OfficeSessionDocument
                    {
                        Id = Guid.NewGuid().ToString("D"),
                        Mode = "edit",
                        Host = "word",
                        FormulaId =
                            selected.FormulaId
                            ?? Guid.NewGuid().ToString("D"),
                        SourceDocumentId =
                            selected.DocumentId,
                        SourceObjectId =
                            selected.ObjectId,
                        Title = metadata.Title,
                        Lines = metadata.Lines.ConvertAll(
                            line => new FormulaLine
                            {
                                Id = line.Id,
                                Latex = line.Latex,
                            }),
                        CodeFormat = metadata.CodeFormat,
                        DisplayMode = metadata.DisplayMode,
                        ObjectMode =
                            FormulaOleContract.MathTypeOleMode,
                        Numbered = metadata.Numbered,
                        MathTypeNumberPosition =
                            metadata.Numbered
                                ? GetMathTypeNumberPositionForRange(
                                    selected.ObjectId)
                                : "right",
                        FontSizePt = target,
                        OriginalMetadata = metadata,
                        Dirty = true,
                        Status = "committing",
                    };

                Release(shape);
                shape = null;
                Release(mathTypeDocument);
                mathTypeDocument = null;
                _ = ReplaceMathTypeOle(
                    resizeSession,
                    sourceMathMl,
                    emfPath: null);
                return target;
            }
            finally
            {
                Release(shape);
                Release(mathTypeDocument);
            }
        }

        var kind =
            ObjectModeToHostKind(
                selected.ObjectMode);
        if (kind is not
            (WordFormulaHostKind.Omml
             or WordFormulaHostKind.VisualTeX))
            throw new NotSupportedException(
                "The selected formula host does not support direct font-size editing.");

        Document? document = null;
        Range? hostRange = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);
            var host =
                WordFormulaOperationLocator.ResolveCapturedHost(
                    _application,
                    document,
                    selected.ObjectId,
                    kind,
                    selected.FormulaId);

            _ = WordFormulaMutationTransaction.Execute(
                _application,
                document,
                "VisualTeX Set Formula Font Size",
                () =>
                {
                    if (kind == WordFormulaHostKind.Omml)
                    {
                        host =
                            WordFormulaHostMutationKernel
                                .AdoptUnownedOmmlForMutation(
                                    document,
                                    host,
                                    selected.FormulaId
                                        ?? host.FormulaId);

                        hostRange =
                            WordFormulaHostSemanticReader.CreateRange(
                                document,
                                host.Range);
                        WordFormulaHostLayout.ApplyOmmlLocalTypography(
                            hostRange,
                            target);

                        var resolved =
                            WordFormulaHostResolver.ResolveLocal(
                                document,
                                hostRange,
                                WordFormulaHostKind.Omml)
                            ?? throw new InvalidDataException(
                                "The OMML host disappeared during its font-size edit.");
                        var numbering =
                            WordFormulaNumberingResolver.ResolveLocal(
                                document,
                                resolved);
                        if (host.Numbering.Numbered
                            && !numbering.Numbered)
                            throw new InvalidDataException(
                                "The OMML font-size edit changed its numbering container.");
                    }
                    else
                    {
                        var updated =
                            WordVisualTeXHostWriter.UpdateFontSize(
                                document,
                                host,
                                target);
                        if (host.Numbering.Numbered
                            && !updated.Numbering.Numbered)
                            throw new InvalidDataException(
                                "The VisualTeX font-size edit changed its numbering container.");
                    }
                    return true;
                });

            if (kind == WordFormulaHostKind.Omml)
            {
                Release(hostRange);
                hostRange =
                    WordFormulaHostSemanticReader.CreateRange(
                        document,
                        host.Range);
                if (string.Equals(
                        host.DisplayMode,
                        "inline",
                        StringComparison.OrdinalIgnoreCase))
                    WordFormulaHostLayout.RestoreInlineCaret(
                        _application,
                        hostRange);
            }

            return target;
        }
        finally
        {
            Release(hostRange);
            Release(document);
        }
    }

    internal string DeleteSelectedFormulaCore()
    {
        var selected = ReadSelection();
        if (string.IsNullOrWhiteSpace(selected.ObjectId)
            || string.IsNullOrWhiteSpace(selected.ObjectMode))
            throw new InvalidOperationException(
                "请先选择一个公式。");

        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);

            var kind =
                ObjectModeToHostKind(
                    selected.ObjectMode);
            var host =
                WordFormulaOperationLocator.ResolveCapturedHost(
                    _application,
                    document,
                    selected.ObjectId,
                    kind,
                    selected.FormulaId);
            WordFormulaHostMutationKernel.Delete(
                _application,
                document,
                host);

            return selected.FormulaId
                ?? host.FormulaId
                ?? string.Empty;
        }
        finally
        {
            Release(document);
        }
    }

    internal OfficeObjectResult ApplyMathTypeTargetHostSession(
        OfficeSessionDocument session,
        string? sourceObjectMode,
        string mathMl,
        string emfPath)
    {
        if (string.Equals(
                session.Mode,
                "create",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                sourceObjectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal))
        {
            return string.Equals(
                    session.Mode,
                    "edit",
                    StringComparison.OrdinalIgnoreCase)
                ? ReplaceMathTypeOle(
                    session,
                    mathMl,
                    emfPath)
                : InsertMathTypeOle(
                    session,
                    mathMl,
                    emfPath);
        }

        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(
                document,
                session.SourceDocumentId);

            var sourceKind =
                ObjectModeToHostKind(sourceObjectMode);
            if (sourceKind is not
                (WordFormulaHostKind.Omml
                 or WordFormulaHostKind.VisualTeX))
                throw new NotSupportedException(
                    "Only OMML/VisualTeX sources are routed through the rebuilt host core when targeting MathType.");

            var source =
                WordFormulaOperationLocator.ResolveCapturedHost(
                    _application,
                    document,
                    session.SourceObjectId,
                    sourceKind,
                    session.OriginalMetadata?.FormulaId
                        ?? session.FormulaId);

            return WordFormulaHostMutationKernel.ReplaceSourceWithExternalTarget(
                _application,
                document,
                source,
                "VisualTeX Replace Formula Host With MathType",
                insertion =>
                {
                    var originalSourceObjectId =
                        session.SourceObjectId;
                    try
                    {
                        session.SourceObjectId =
                            RangeReference(insertion);
                        return InsertMathTypeOle(
                            session,
                            mathMl,
                            emfPath,
                            preserveCapturedInsertion: true);
                    }
                    finally
                    {
                        session.SourceObjectId =
                            originalSourceObjectId;
                    }
                });
        }
        finally
        {
            Release(document);
        }
    }

    internal OfficeObjectResult ApplyOmmlVisualTeXHostSession(
        OfficeSessionDocument session,
        string? sourceObjectMode,
        string? mathMl,
        string? pngPath,
        string? emfPath)
    {
        Document? document = null;
        Range? insertion = null;
        Range? finalRange = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(
                document,
                session.SourceDocumentId);

            var request = BuildHostWriteRequest(
                session,
                mathMl,
                pngPath,
                emfPath);

            WordFormulaHostDescriptor resultHost;
            if (string.Equals(
                    session.Mode,
                    "create",
                    StringComparison.OrdinalIgnoreCase))
            {
                insertion =
                    WordFormulaOperationLocator.ResolveCreateInsertion(
                        document,
                        session.SourceObjectId,
                        session.DisplayMode);
                resultHost =
                    WordFormulaHostMutationKernel.Insert(
                        _application,
                        document,
                        insertion,
                        request);
            }
            else
            {
                var sourceKind =
                    ObjectModeToHostKind(sourceObjectMode);
                var source =
                    WordFormulaOperationLocator.ResolveCapturedHost(
                        _application,
                        document,
                        session.SourceObjectId,
                        sourceKind,
                        session.OriginalMetadata?.FormulaId
                            ?? session.FormulaId);

                resultHost =
                    WordFormulaHostMutationKernel.Replace(
                        _application,
                        document,
                        source,
                        request);
            }

            if (!string.IsNullOrWhiteSpace(resultHost.FormulaId))
                session.FormulaId = resultHost.FormulaId!;

            finalRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    resultHost.Range);
            RestoreCaretAfterCoreMutation(
                document,
                resultHost,
                finalRange);

            return Result(
                session,
                document,
                documentIdentityAlreadyValidated: true);
        }
        finally
        {
            Release(finalRange);
            Release(insertion);
            Release(document);
        }
    }

    private static WordFormulaHostWriteRequest BuildHostWriteRequest(
        OfficeSessionDocument session,
        string? mathMl,
        string? pngPath,
        string? emfPath)
    {
        var targetKind =
            ObjectModeToHostKind(session.ObjectMode);
        if (targetKind is not
            (WordFormulaHostKind.Omml
             or WordFormulaHostKind.VisualTeX))
            throw new NotSupportedException(
                "The rebuilt host request owns OMML and VisualTeX targets only.");

        var metadata = session.ToMetadata();
        metadata.FormulaId =
            Guid.TryParse(session.FormulaId, out var parsed)
                ? parsed.ToString("D")
                : Guid.NewGuid().ToString("D");
        metadata.DisplayMode = session.DisplayMode;
        metadata.Numbered = session.Numbered;
        metadata.Validate();

        var export = session.ExportResult;
        return new WordFormulaHostWriteRequest
        {
            Kind = targetKind,
            DisplayMode = session.DisplayMode,
            Numbered = session.Numbered,
            FormulaId = metadata.FormulaId,
            Metadata = targetKind ==
                WordFormulaHostKind.VisualTeX
                    ? metadata
                    : null,
            MathMl = targetKind ==
                WordFormulaHostKind.Omml
                    ? mathMl
                    : null,
            PngPath = targetKind ==
                WordFormulaHostKind.VisualTeX
                    ? pngPath
                    : null,
            EmfPath = targetKind ==
                WordFormulaHostKind.VisualTeX
                    ? emfPath
                    : null,
            WidthPoints = Math.Max(
                1f,
                (export?.Width ?? 1f) * 0.75f),
            HeightPoints = Math.Max(
                1f,
                (export?.Height ?? 1f) * 0.75f),
            ExportedHeightPixels =
                export?.Height ?? 0f,
            ExportedBaselinePixels =
                export?.Baseline,
            FontSizePoints =
                FormulaFontSize.Normalize(
                    session.FontSizePt),
        };
    }

    private static void TryDeleteBookmark(
        Document document,
        string name)
    {
        if (document is null
            || string.IsNullOrWhiteSpace(name))
            return;

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name))
                return;
            bookmark = bookmarks[name];
            bookmark.Delete();
        }
        catch
        {
            // Cleanup of an explicitly named local locator is best-effort. It is
            // never formula-existence or semantic-success evidence.
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private string ResolveLegacyWrapperSourceMode(
        OfficeSessionDocument session)
    {
        Document? document = null;
        Range? captured = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureSourceDocument(
                document,
                session.SourceDocumentId);
            captured =
                WordFormulaOperationLocator.ResolveCapturedRange(
                    document,
                    session.SourceObjectId);

            var matches =
                new List<string>(3);

            try
            {
                if (WordFormulaHostResolver.ResolveLocal(
                        document,
                        captured,
                        WordFormulaHostKind.VisualTeX)
                    is not null)
                    matches.Add(
                        FormulaOleContract.NativeOleMode);
            }
            catch (InvalidDataException) { }

            try
            {
                if (WordFormulaHostResolver.ResolveLocal(
                        document,
                        captured,
                        WordFormulaHostKind.Omml)
                    is not null)
                    matches.Add(
                        FormulaOleContract.WordOmmlMode);
            }
            catch (InvalidDataException) { }

            try
            {
                if (WordMathTypeHostAdapter.ResolveLocal(
                        _application,
                        document,
                        captured)
                    is not null)
                    matches.Add(
                        FormulaOleContract.MathTypeOleMode);
            }
            catch (InvalidDataException) { }

            if (matches.Count != 1)
                throw new InvalidDataException(
                    $"The captured Word range resolves to {matches.Count} formula host families; exactly one is required.");

            return matches[0];
        }
        finally
        {
            Release(captured);
            Release(document);
        }
    }

    private static WordFormulaHostKind ObjectModeToHostKind(
        string? objectMode)
    {
        if (string.Equals(
                objectMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal))
            return WordFormulaHostKind.Omml;
        if (string.Equals(
                objectMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal))
            return WordFormulaHostKind.VisualTeX;
        if (string.Equals(
                objectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal))
            return WordFormulaHostKind.MathType;

        throw new NotSupportedException(
            $"Unsupported Word formula object mode '{objectMode ?? "<null>"}'.");
    }

    private void RestoreCaretAfterCoreMutation(
        Document document,
        WordFormulaHostDescriptor host,
        Range hostRange)
    {
        if (string.Equals(
                host.DisplayMode,
                "inline",
                StringComparison.OrdinalIgnoreCase))
        {
            WordFormulaHostLayout.RestoreInlineCaret(
                _application,
                hostRange);
            return;
        }

        Selection? selection = null;
        Range? target = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            selection = _application.Selection;

            // Pure OMML never owns an external numbering table/cell container:
            // its #(SEQ/STYLEREF) lives inside the OMath itself. Only external
            // OLE hosts need container discovery before restoring the caret.
            if (host.Kind != WordFormulaHostKind.Omml)
            {
                var numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        host);

                if (host.Kind ==
                        WordFormulaHostKind.VisualTeX
                    && !host.WithinTable
                    && !string.IsNullOrWhiteSpace(
                        host.FormulaId)
                    && (numbering.Numbered
                        || host.Numbering.Numbered
                        || host.Metadata?.Numbered == true)
                    && !WordVisualTeXParagraphNumbering
                        .IsSelfContainedHost(
                            document,
                            host))
                {
                    Range? typingRange = null;
                    try
                    {
                        // Legacy numbered VisualTeX hosts can still have a hidden
                        // caption paragraph. Keep that migration path isolated here.
                        // New self-contained VisualTeXPlaceRef paragraphs fall
                        // through to the ordinary next-paragraph caret path below.
                        typingRange =
                            WordEquationNumbering
                                .EnsureNormalTypingParagraphAfterNumberedDisplay(
                                    document,
                                    host.FormulaId!);
                        if (typingRange is null)
                            throw new InvalidDataException(
                                "VisualTeX could not create a safe typing paragraph after the legacy numbered display formula.");

                        selection.SetRange(
                            typingRange.Start,
                            typingRange.Start);
                        selection.Collapse(
                            WdCollapseDirection.wdCollapseStart);
                        return;
                    }
                    finally
                    {
                        Release(typingRange);
                    }
                }

                if (numbering.ContainerKind ==
                        WordFormulaNumberingContainerKind.CanonicalBodyTable
                    && numbering.ContainerRange is not null)
                {
                    var position =
                        numbering.ContainerRange.End;
                    Range? content = null;
                    try
                    {
                        content = document.Content;
                        position = Math.Max(
                            content.Start,
                            Math.Min(
                                position,
                                Math.Max(
                                    content.Start,
                                    content.End - 1)));
                        target = document.Range(
                            position,
                            position);
                        if (Convert.ToBoolean(
                                target.get_Information(
                                    WdInformation.wdWithInTable)))
                        {
                            target.InsertAfter("\r");
                            target.SetRange(
                                position,
                                position);
                        }
                    }
                    finally { Release(content); }
                    selection.SetRange(
                        target.Start,
                        target.End);
                    return;
                }

                if (numbering.ContainerKind ==
                    WordFormulaNumberingContainerKind.CanonicalUserTableCell)
                {
                    paragraphs = hostRange.Paragraphs;
                    if (paragraphs.Count == 1)
                    {
                        paragraph = paragraphs[1];
                        paragraphRange =
                            paragraph.Range.Duplicate;
                        var position = Math.Max(
                            paragraphRange.Start,
                            paragraphRange.End - 1);
                        selection.SetRange(
                            position,
                            position);
                        return;
                    }
                }
            }

            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count != 1)
            {
                selection.SetRange(
                    hostRange.End,
                    hostRange.End);
                return;
            }

            paragraph = paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            var nextStart = paragraphRange.End;
            Range? contentRange = null;
            try
            {
                contentRange = document.Content;
                if (nextStart >= contentRange.End)
                {
                    paragraphRange.InsertParagraphAfter();
                }
            }
            finally { Release(contentRange); }

            selection.SetRange(
                nextStart,
                nextStart);
        }
        catch
        {
            try
            {
                selection?.SetRange(
                    hostRange.End,
                    hostRange.End);
            }
            catch { }
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(target);
            Release(selection);
        }
    }
}
