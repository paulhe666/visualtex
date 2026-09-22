using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Read-only, local numbering-container discovery.
///
/// OMML canonical numbering is Word's own native m:eqArr + #(SEQ) host,
/// and Word's native Equation cross-reference mechanism owns its REF targets.
/// VisualTeX OLE keeps the external Word SEQ layouts and VTEqNum_<FormulaId>
/// targets required by an InlineShape.
///
/// No document-wide bookmark/field/table enumeration is allowed for local host
/// resolution.
/// </summary>
internal static class WordFormulaNumberingResolver
{
    private const string NumberBookmarkPrefix = "VTEqNum_";
    private const string SequenceName = "VisualTeXEquation";

    internal static WordFormulaNumberingDescriptor ResolveLocal(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (host is null) throw new ArgumentNullException(nameof(host));

        Range? hostRange = null;
        Tables? tables = null;
        Table? table = null;
        try
        {
            hostRange = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);

            if (host.Kind == WordFormulaHostKind.Omml)
            {
                // Native OMML has no VisualTeX ownership layer. Numbering truth
                // comes only from the exact OMath + local OOXML/fields.
                if (WordNativeOmmlNumbering.TryResolveUnownedNativeStructure(
                        document,
                        hostRange,
                        out var nativeOmml))
                {
                    nativeOmml.FormulaId = null;
                    return nativeOmml;
                }

                return HasLocalLegacyNumberingEvidence(
                        hostRange,
                        formulaId: null)
                    ? Legacy(host)
                    : Unnumbered(host);
            }

            tables = hostRange.Tables;
            if (tables.Count > 0)
            {
                table = tables[1];
                if (TryResolveCanonicalBodyTable(
                        document,
                        table,
                        hostRange,
                        host.Kind,
                        host.FormulaId,
                        out var canonical))
                    return canonical;

                // A formula inside an ordinary user table is not allowed to be
                // reinterpreted as a VisualTeX-owned 1x3 merely because that user
                // table happens to have three columns.
                if (IsInsideUserTableCell(hostRange, table))
                {
                    if (TryResolveCanonicalUserTableCell(
                            hostRange,
                            host.FormulaId,
                            out var cellCanonical))
                        return cellCanonical;

                    return HasLocalLegacyNumberingEvidence(
                            hostRange,
                            host.FormulaId)
                        ? Legacy(host)
                        : Unnumbered(host);
                }
            }

            if (host.Kind == WordFormulaHostKind.VisualTeX
                && TryResolveCanonicalBodyTabParagraph(
                    document,
                    hostRange,
                    host.FormulaId,
                    out var bodyTab))
                return bodyTab;

            return HasLocalLegacyNumberingEvidence(
                    hostRange,
                    host.FormulaId)
                ? Legacy(host)
                : Unnumbered(host);
        }
        finally
        {
            Release(table);
            Release(tables);
            Release(hostRange);
        }
    }

    private static bool TryResolveCanonicalBodyTable(
        Document document,
        Table table,
        Range hostRange,
        WordFormulaHostKind hostKind,
        string? expectedFormulaId,
        out WordFormulaNumberingDescriptor descriptor)
    {
        descriptor = null!;
        if (!Guid.TryParse(expectedFormulaId, out var parsed))
            return false;

        Rows? rows = null;
        Columns? columns = null;
        Cell? leftCell = null;
        Cell? centerCell = null;
        Cell? rightCell = null;
        Range? tableRange = null;
        Range? leftRange = null;
        Range? centerRange = null;
        Range? rightRange = null;
        Fields? rightFields = null;
        Field? numberField = null;
        Range? fieldCode = null;
        Range? fieldResult = null;
        Bookmarks? rightBookmarks = null;
        Bookmark? numberBookmark = null;
        Range? numberRange = null;
        try
        {
            rows = table.Rows;
            columns = table.Columns;
            if (rows.Count != 1 || columns.Count != 3) return false;

            tableRange = table.Range.Duplicate;
            // Canonical numbered body tables are top-level. A 1x3 nested inside a
            // user table is never a production host in the rebuilt core.
            if (IsRangeNestedInsideAnotherTable(tableRange)) return false;

            leftCell = table.Cell(1, 1);
            centerCell = table.Cell(1, 2);
            rightCell = table.Cell(1, 3);
            leftRange = leftCell.Range.Duplicate;
            centerRange = centerCell.Range.Duplicate;
            rightRange = rightCell.Range.Duplicate;

            if (!ContainsOnlyStructuralWordText(leftRange.Text))
                return false;
            if (hostRange.Start < centerRange.Start
                || hostRange.End > centerRange.End)
                return false;

            // The center cell may contain the formula host and ordinary table
            // structural markers only. It must not own a second OMML/OLE host.
            if (!CenterCellOwnsExactlyHost(centerRange, hostRange, hostKind))
                return false;

            rightFields = rightRange.Fields;
            Field? matchedField = null;
            for (var index = 1; index <= rightFields.Count; index++)
            {
                numberField = rightFields[index];
                fieldCode = numberField.Code;
                if (IsVisualTeXSequenceField(fieldCode.Text))
                {
                    if (matchedField is not null)
                    {
                        Release(matchedField);
                        return false;
                    }
                    matchedField = numberField;
                    numberField = null;
                }
                Release(fieldCode); fieldCode = null;
                Release(numberField); numberField = null;
            }
            if (matchedField is null) return false;
            numberField = matchedField;
            fieldCode = numberField.Code;
            fieldResult = numberField.Result;

            var expectedName =
                NumberBookmarkPrefix + parsed.ToString("N");
            rightBookmarks = rightRange.Bookmarks;
            for (var index = 1; index <= rightBookmarks.Count; index++)
            {
                numberBookmark = rightBookmarks[index];
                if (!string.Equals(
                        numberBookmark.Name,
                        expectedName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    Release(numberBookmark);
                    numberBookmark = null;
                    continue;
                }
                numberRange = numberBookmark.Range.Duplicate;
                break;
            }

            if (numberRange is null)
                return false;
            if (numberRange.Start < rightRange.Start
                || numberRange.End > rightRange.End)
                return false;
            if (fieldResult.Start < numberRange.Start
                || fieldResult.End > numberRange.End)
                return false;

            descriptor = new WordFormulaNumberingDescriptor
            {
                Numbered = true,
                FormulaId = expectedFormulaId,
                Position = "right",
                ContainerKind =
                    WordFormulaNumberingContainerKind.CanonicalBodyTable,
                ContainerRange = Address(tableRange),
                NumberRange = Address(numberRange),
            };
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(numberRange);
            Release(numberBookmark);
            Release(rightBookmarks);
            Release(fieldResult);
            Release(fieldCode);
            Release(numberField);
            Release(rightFields);
            Release(rightRange);
            Release(centerRange);
            Release(leftRange);
            Release(tableRange);
            Release(rightCell);
            Release(centerCell);
            Release(leftCell);
            Release(columns);
            Release(rows);
        }
    }

    private static bool TryResolveCanonicalUserTableCell(
        Range hostRange,
        string? expectedFormulaId,
        out WordFormulaNumberingDescriptor descriptor)
    {
        descriptor = null!;
        if (!Guid.TryParse(expectedFormulaId, out var parsed))
            return false;

        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        try
        {
            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (!IsWithinTable(paragraphRange)) return false;

            var sequenceCount = 0;
            fields = paragraphRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                field = fields[index];
                code = field.Code;
                if (IsVisualTeXSequenceField(code.Text))
                    sequenceCount++;
                Release(code); code = null;
                Release(field); field = null;
            }
            if (sequenceCount != 1) return false;

            var expectedName =
                NumberBookmarkPrefix + parsed.ToString("N");
            bookmarks = paragraphRange.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                bookmark = bookmarks[index];
                if (!string.Equals(
                        bookmark.Name,
                        expectedName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    Release(bookmark);
                    bookmark = null;
                    continue;
                }
                bookmarkRange = bookmark.Range.Duplicate;
                break;
            }
            if (bookmarkRange is null) return false;
            if (bookmarkRange.Start <= hostRange.End) return false;

            descriptor = new WordFormulaNumberingDescriptor
            {
                Numbered = true,
                FormulaId = expectedFormulaId,
                Position = "right",
                ContainerKind =
                    WordFormulaNumberingContainerKind.CanonicalUserTableCell,
                ContainerRange = Address(paragraphRange),
                NumberRange = Address(bookmarkRange),
            };
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool TryResolveCanonicalBodyTabParagraph(
        Document document,
        Range hostRange,
        string? expectedFormulaId,
        out WordFormulaNumberingDescriptor descriptor)
    {
        descriptor = null!;

        // New canonical VisualTeX numbering is a MathType-style self-contained
        // field tree in the formula paragraph. Keep the legacy REF-to-hidden-
        // caption recognizer below only as migration input for older documents.
        if (WordVisualTeXParagraphNumbering.TryResolve(
                document,
                hostRange,
                expectedFormulaId,
                out descriptor))
            return true;

        bool Reject(string reason)
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_TRACE_BODY_TAB_RESOLVER"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"[body-tab reject] reason={reason} formulaId={expectedFormulaId ?? "<null>"} host={hostRange.Start}:{hostRange.End}");
            return false;
        }
        if (!Guid.TryParse(expectedFormulaId, out var parsed))
            return Reject("invalid-formula-id");
        if (IsWithinTable(hostRange))
            return Reject("host-in-table");

        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? leading = null;
        ParagraphFormat? format = null;
        TabStops? tabs = null;
        TabStop? tab = null;
        Range? visibleNumber = null;
        try
        {
            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                return Reject("host-spans-paragraphs");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (IsWithinTable(paragraphRange))
                return Reject("paragraph-in-table");
            if (hostRange.Start <= paragraphRange.Start)
                return Reject($"no-leading-slot paragraph={paragraphRange.Start}:{paragraphRange.End}");

            leading = document.Range(
                hostRange.Start - 1,
                hostRange.Start);
            if (!string.Equals(
                    leading.Text,
                    "\t",
                    StringComparison.Ordinal))
                return Reject(
                    $"leading-not-tab code={string.Join(",", (leading.Text ?? string.Empty).Select(ch => $"U+{(int)ch:X4}"))}");

            format = paragraphRange.ParagraphFormat;
            if (format.Alignment !=
                WdParagraphAlignment.wdAlignParagraphJustify)
                return Reject($"alignment={format.Alignment}");
            tabs = format.TabStops;
            var hasCenter = false;
            var hasRight = false;
            for (var index = 1;
                 index <= tabs.Count;
                 index++)
            {
                Release(tab);
                tab = tabs[index];
                hasCenter |=
                    tab.Alignment ==
                    WdTabAlignment.wdAlignTabCenter;
                hasRight |=
                    tab.Alignment ==
                    WdTabAlignment.wdAlignTabRight;
            }
            if (!hasCenter || !hasRight)
                return Reject($"tabs center={hasCenter} right={hasRight}");

            visibleNumber =
                TryResolveBodyTabVisibleReferenceResult(
                    paragraphRange,
                    parsed.ToString("D"));
            if (visibleNumber is null)
                return Reject("visible-number-missing");
            if (visibleNumber.StoryType != hostRange.StoryType)
                return Reject("visible-number-story");
            if (visibleNumber.Start < hostRange.End)
                return Reject($"visible-number-before-host number={visibleNumber.Start}:{visibleNumber.End}");
            if (visibleNumber.Start < paragraphRange.Start
                || visibleNumber.End > paragraphRange.End)
                return Reject($"visible-number-outside-paragraph number={visibleNumber.Start}:{visibleNumber.End} paragraph={paragraphRange.Start}:{paragraphRange.End}");

            descriptor = new WordFormulaNumberingDescriptor
            {
                Numbered = true,
                FormulaId = parsed.ToString("D"),
                Position = "right",
                ContainerKind =
                    WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph,
                ContainerRange = Address(paragraphRange),
                NumberRange = Address(visibleNumber),
            };
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(visibleNumber);
            Release(tab);
            Release(tabs);
            Release(format);
            Release(leading);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static Range? TryResolveBodyTabVisibleReferenceResult(
        Range paragraphRange,
        string formulaId)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        Range? found = null;
        try
        {
            var targetName =
                WordEquationNumbering.NativeNumberBookmarkName(
                    formulaId);
            fields = paragraphRange.Fields;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(result);
                result = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                if (field.Type != WdFieldType.wdFieldRef)
                    continue;
                code = field.Code.Duplicate;
                var fieldCode = code.Text ?? string.Empty;
                if (fieldCode.IndexOf(
                        "REF " + targetName,
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                result = field.Result.Duplicate;
                if (found is not null)
                {
                    Release(found);
                    found = null;
                    return null;
                }
                found = result;
                result = null;
            }

            var resolved = found;
            found = null;
            return resolved;
        }
        finally
        {
            Release(found);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static bool IsWithinTable(Range range)
    {
        try
        {
            return Convert.ToBoolean(
                range.get_Information(WdInformation.wdWithInTable));
        }
        catch { return false; }
    }

    private static bool HasLocalLegacyNumberingEvidence(
        Range hostRange,
        string? formulaId)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count < 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;

            fields = paragraphRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                field = fields[index];
                code = field.Code;
                if (IsVisualTeXSequenceField(code.Text))
                    return true;
                Release(code); code = null;
                Release(field); field = null;
            }

            bookmarks = paragraphRange.Bookmarks;
            var expectedSuffix = Guid.TryParse(formulaId, out var parsed)
                ? parsed.ToString("N")
                : null;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                bookmark = bookmarks[index];
                var name = bookmark.Name ?? string.Empty;
                if (name.StartsWith(
                        "VTEq",
                        StringComparison.OrdinalIgnoreCase)
                    && (expectedSuffix is null
                        || name.EndsWith(
                            expectedSuffix,
                            StringComparison.OrdinalIgnoreCase)))
                    return true;
                Release(bookmark);
                bookmark = null;
            }
            return false;
        }
        catch (Exception error)
        {
            // Do not substitute requested/cached Numbered state for Word's
            // current structure. If local inspection itself fails, stop instead
            // of guessing either legacy or unnumbered topology.
            throw new InvalidDataException(
                "Word could not inspect the formula's local numbering structure.",
                error);
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool CenterCellOwnsExactlyHost(
        Range centerRange,
        Range hostRange,
        WordFormulaHostKind hostKind)
    {
        OMaths? maths = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            maths = centerRange.OMaths;
            shapes = centerRange.InlineShapes;

            if (hostKind == WordFormulaHostKind.Omml)
            {
                if (maths.Count != 1) return false;
                var matchingShapes = 0;
                for (var index = 1; index <= shapes.Count; index++)
                {
                    shape = shapes[index];
                    if (WordFormulaMetadataReader.IsNativeOle(shape))
                        matchingShapes++;
                    Release(shape); shape = null;
                }
                if (matchingShapes != 0) return false;

                OMath? math = null;
                Range? mathRange = null;
                try
                {
                    math = maths[1];
                    mathRange = math.Range.Duplicate;
                    return mathRange.Start == hostRange.Start
                        && mathRange.End == hostRange.End
                        && mathRange.StoryType == hostRange.StoryType;
                }
                finally
                {
                    Release(mathRange);
                    Release(math);
                }
            }

            if (hostKind == WordFormulaHostKind.VisualTeX)
            {
                if (maths.Count != 0) return false;
                var matches = 0;
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
                    if (shapeRange.Start == hostRange.Start
                        && shapeRange.End == hostRange.End
                        && shapeRange.StoryType == hostRange.StoryType)
                        matches++;
                    Release(shapeRange); shapeRange = null;
                    Release(shape); shape = null;
                }
                return matches == 1;
            }

            return false;
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(maths);
        }
    }

    private static bool IsInsideUserTableCell(
        Range hostRange,
        Table immediateTable)
    {
        Range? tableRange = null;
        try
        {
            tableRange = immediateTable.Range;
            // Any non-canonical table is foreign user content for the resolver.
            return hostRange.Start >= tableRange.Start
                && hostRange.End <= tableRange.End;
        }
        catch
        {
            return true;
        }
        finally { Release(tableRange); }
    }

    private static bool IsRangeNestedInsideAnotherTable(Range tableRange)
    {
        Range? probe = null;
        Tables? tables = null;
        try
        {
            if (tableRange.Start <= 0) return false;
            var document = tableRange.Document;
            var start = Math.Max(0, tableRange.Start - 1);
            probe = document.Range(start, tableRange.Start);
            tables = probe.Tables;
            return tables.Count > 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(tables);
            Release(probe);
        }
    }

    private static bool IsVisualTeXSequenceField(string? code) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf("SEQ", StringComparison.OrdinalIgnoreCase) >= 0
        && code.IndexOf(
            SequenceName,
            StringComparison.OrdinalIgnoreCase) >= 0;

    private static bool ContainsOnlyStructuralWordText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return true;
        foreach (var value in text!)
        {
            if (value is '\r' or '\n' or '\t' or '\v' or '\a')
                continue;
            if (!char.IsWhiteSpace(value)) return false;
        }
        return true;
    }

    private static WordFormulaNumberingDescriptor Legacy(
        WordFormulaHostDescriptor host) =>
        new()
        {
            Numbered = true,
            FormulaId = host.FormulaId,
            Position = host.Numbering.Position,
            ContainerKind = WordFormulaNumberingContainerKind.Legacy,
        };

    private static WordFormulaNumberingDescriptor Unnumbered(
        WordFormulaHostDescriptor host) =>
        new()
        {
            Numbered = false,
            FormulaId = host.FormulaId,
            Position = host.Numbering.Position,
            ContainerKind = WordFormulaNumberingContainerKind.None,
        };

    private static WordFormulaRangeAddress Address(Range range) => new()
    {
        StoryType = range.StoryType,
        Start = range.Start,
        End = range.End,
    };

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
