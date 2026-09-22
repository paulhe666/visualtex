using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using WordApplication = Microsoft.Office.Interop.Word.Application;

namespace VisualTeX.WordVsto;

/// <summary>
/// Resolves only coordinates captured when the editor/session opened.
/// No live-selection fallback and no document-wide identity lookup.
/// </summary>
internal static class WordFormulaOperationLocator
{
    private const string RangeReferencePrefix =
        "visualtex-word-vsto-range:";

    internal static Range ResolveCapturedRange(
        Document document,
        string? sourceObjectId)
    {
        if (!TryParseRangeReference(
                sourceObjectId,
                out var start,
                out var end))
            throw new InvalidDataException(
                "The formula session has no captured Word range.");

        Range? content = null;
        try
        {
            content = document.Content;
            if (start < content.Start
                || end < start
                || end > content.End)
                throw new InvalidDataException(
                    "The captured Word range is no longer valid.");
            return document.Range(start, end);
        }
        finally { Release(content); }
    }

    internal static WordFormulaHostDescriptor ResolveCapturedHost(
        WordApplication application,
        Document document,
        string? sourceObjectId,
        WordFormulaHostKind kind,
        string? expectedFormulaId)
    {
        Range? captured = null;
        try
        {
            captured = ResolveCapturedRange(
                document,
                sourceObjectId);
            WordFormulaHostDescriptor? host =
                kind == WordFormulaHostKind.MathType
                    ? WordMathTypeHostAdapter.ResolveLocal(
                        application,
                        document,
                        captured)
                    : WordFormulaHostResolver.ResolveLocal(
                        document,
                        captured,
                        kind);

            if (host is null)
                throw new InvalidDataException(
                    $"The captured {kind} formula no longer exists locally.");

            // The captured range is authoritative for location. A stable owned
            // FormulaId is an additional consistency check, never a fallback
            // locator.
            if (!string.IsNullOrWhiteSpace(expectedFormulaId)
                && !string.IsNullOrWhiteSpace(host.FormulaId)
                && kind != WordFormulaHostKind.MathType
                && !string.Equals(
                    expectedFormulaId,
                    host.FormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "The formula at the captured range now has a different identity.");

            if (host.FormulaId is null
                && Guid.TryParse(expectedFormulaId, out var parsed))
                host.FormulaId = parsed.ToString("D");

            if (kind == WordFormulaHostKind.VisualTeX
                || (kind == WordFormulaHostKind.Omml
                    && host.Display))
            {
                host.Numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        host);
            }
            return host;
        }
        finally { Release(captured); }
    }

    internal static Range ResolveCreateInsertion(
        Document document,
        string? sourceObjectId,
        string displayMode)
    {
        Range? captured = null;
        try
        {
            captured = ResolveCapturedRange(
                document,
                sourceObjectId);
            captured.Collapse(WdCollapseDirection.wdCollapseEnd);

            if (!string.Equals(
                    displayMode,
                    "block",
                    StringComparison.OrdinalIgnoreCase))
                return captured.Duplicate;

            return ResolveDisplayCreateInsertion(
                document,
                captured);
        }
        finally { Release(captured); }
    }

    private static Range ResolveDisplayCreateInsertion(
        Document document,
        Range captured)
    {
        // The overwhelmingly common create case is a caret in a genuinely empty
        // paragraph. No formula host can occupy a zero-length paragraph body, so
        // avoid two local host-resolution COM passes in that case.
        if (captured.Start == captured.End)
        {
            Paragraphs? emptyParagraphs = null;
            Paragraph? emptyParagraph = null;
            Range? emptyParagraphRange = null;
            try
            {
                emptyParagraphs = captured.Paragraphs;
                if (emptyParagraphs.Count == 1)
                {
                    emptyParagraph = emptyParagraphs[1];
                    emptyParagraphRange =
                        emptyParagraph.Range.Duplicate;
                    var bodyEnd = Math.Max(
                        emptyParagraphRange.Start,
                        emptyParagraphRange.End - 1);
                    if (bodyEnd == emptyParagraphRange.Start)
                    {
                        return document.Range(
                            emptyParagraphRange.Start,
                            emptyParagraphRange.Start);
                    }
                }
            }
            finally
            {
                Release(emptyParagraphRange);
                Release(emptyParagraph);
                Release(emptyParagraphs);
            }
        }

        // If the caret is currently on an existing VisualTeX/OMML formula,
        // create after its complete local numbering container rather than inside
        // the old host.
        var existing = TryResolveAnyOwnedHost(
            document,
            captured);
        if (existing is not null)
        {
            var numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    existing);
            if (numbering.ContainerRange is not null
                && numbering.ContainerKind is
                    (WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
                     or WordFormulaNumberingContainerKind.CanonicalBodyTable))
                return CreateBodyParagraphAfter(
                    document,
                    numbering.ContainerRange.End);

            Range? hostRange = null;
            try
            {
                hostRange =
                    WordFormulaHostSemanticReader.CreateRange(
                        document,
                        existing.Range);
                return CreateParagraphAfterHost(
                    document,
                    hostRange);
            }
            finally { Release(hostRange); }
        }

        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? body = null;
        try
        {
            paragraphs = captured.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "The captured insertion point does not belong to one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;

            var bodyEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            body = document.Range(
                paragraphRange.Start,
                bodyEnd);
            var hasVisibleBody =
                ContainsVisibleText(body.Text);

            if (!hasVisibleBody)
                return document.Range(
                    paragraphRange.Start,
                    paragraphRange.Start);

            // A display equation never splices into ordinary prose. Create one
            // new paragraph after the paragraph containing the captured caret.
            var nextStart = paragraphRange.End;
            paragraphRange.InsertParagraphAfter();
            return document.Range(
                nextStart,
                nextStart);
        }
        finally
        {
            Release(body);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static WordFormulaHostDescriptor? TryResolveAnyOwnedHost(
        Document document,
        Range captured)
    {
        try
        {
            var visual = WordFormulaHostResolver.ResolveLocal(
                document,
                captured,
                WordFormulaHostKind.VisualTeX);
            if (visual is not null) return visual;
        }
        catch { }

        try
        {
            return WordFormulaHostResolver.ResolveLocal(
                document,
                captured,
                WordFormulaHostKind.Omml);
        }
        catch
        {
            return null;
        }
    }

    private static Range CreateParagraphAfterHost(
        Document document,
        Range hostRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "A display formula host must belong to one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;

            if (IsWithinTable(paragraphRange))
            {
                var start = paragraphRange.End;
                paragraphRange.InsertParagraphAfter();
                return document.Range(start, start);
            }

            return CreateBodyParagraphAfter(
                document,
                paragraphRange.End);
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static Range CreateBodyParagraphAfter(
        Document document,
        int position)
    {
        Range? content = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            content = document.Content;
            var last = Math.Max(
                content.Start,
                content.End - 1);
            position = Math.Max(
                content.Start,
                Math.Min(position, last));

            probe = document.Range(
                position,
                Math.Min(content.End, position + 1));
            if (IsWithinTable(probe))
            {
                // At a body->table boundary Word gives the forward probe table
                // affinity. Insert an explicit body paragraph before that table.
                probe.InsertBefore("\r");
                return document.Range(
                    position,
                    position);
            }

            paragraphs = probe.Paragraphs;
            if (paragraphs.Count == 1)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
                if (!ContainsVisibleText(paragraphRange.Text))
                    return document.Range(
                        paragraphRange.Start,
                        paragraphRange.Start);
            }

            probe.SetRange(position, position);
            probe.InsertBefore("\r");
            return document.Range(
                position,
                position);
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
            Release(content);
        }
    }

    private static bool ContainsVisibleText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return false;
        foreach (var value in text!)
        {
            if (value is '\r' or '\n' or '\t'
                or '\v' or '\a' or '\u0001'
                or '\u200B' or '\u200C')
                continue;
            if (!char.IsWhiteSpace(value))
                return true;
        }
        return false;
    }

    private static bool IsWithinTable(Range range)
    {
        try
        {
            return Convert.ToBoolean(
                range.get_Information(
                    WdInformation.wdWithInTable));
        }
        catch { return false; }
    }

    internal static bool TryParseRangeReference(
        string? value,
        out int start,
        out int end)
    {
        start = 0;
        end = 0;
        if (string.IsNullOrWhiteSpace(value)
            || !value!.StartsWith(
                RangeReferencePrefix,
                StringComparison.Ordinal))
            return false;

        var payload = value.Substring(
            RangeReferencePrefix.Length);
        var separator = payload.IndexOf(':');
        if (separator <= 0
            || separator >= payload.Length - 1)
            return false;

        return int.TryParse(
                payload.Substring(0, separator),
                out start)
            && int.TryParse(
                payload.Substring(separator + 1),
                out end);
    }

    private static void Release(object? value)
    {
        if (value is null
            || !Marshal.IsComObject(value))
            return;
        try { Marshal.ReleaseComObject(value); }
        catch { }
    }
}
