using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Canonical host discovery for OMML and VisualTeX.
///
/// Local discovery is strictly bounded to the supplied range. It never falls
/// back to Document.OMaths, Document.InlineShapes, Document.Bookmarks, or a
/// document WordOpenXML snapshot.
///
/// Whole-document discovery is explicitly separate and builds one immutable
/// scheduling index. That index is invalid after the first mutation.
/// </summary>
internal static class WordFormulaHostResolver
{
    private sealed class IdentityAnchor
    {
        internal string Prefix { get; set; } = string.Empty;
        internal string FormulaId { get; set; } = string.Empty;
        internal WdStoryType StoryType { get; set; }
        internal int Start { get; set; }
        internal int End { get; set; }
    }

    internal static WordFormulaHostDescriptor? ResolveLocal(
        Document document,
        Range selectionRange,
        WordFormulaHostKind kind)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (selectionRange is null) throw new ArgumentNullException(nameof(selectionRange));

        return kind switch
        {
            WordFormulaHostKind.Omml => ResolveLocalOmml(document, selectionRange),
            WordFormulaHostKind.VisualTeX => ResolveLocalVisualTeX(document, selectionRange),
            _ => throw new NotSupportedException(
                $"Local host discovery is not owned by the OMML/VisualTeX core for {kind}."),
        };
    }

    internal static WordFormulaDocumentIndex CaptureDocumentIndex(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        Range? content = null;
        try
        {
            content = document.Content;
            return CaptureScopeIndex(document, content);
        }
        finally { Release(content); }
    }

    internal static WordFormulaDocumentIndex CaptureScopeIndex(
        Document document,
        Range scope)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (scope is null) throw new ArgumentNullException(nameof(scope));

        // One bookmark pass and one REF-result pass for exactly this scope.
        // Per-formula COM scans are forbidden after these snapshots are built.
        var anchors = CaptureIdentityAnchors(scope);
        var referenceResults = CaptureReferenceResultRanges(scope);
        var tableRanges = CaptureTableRanges(scope);
        var omml = CaptureOmmlIndex(
            scope,
            referenceResults,
            tableRanges);
        var visualTeX = CaptureVisualTeXIndex(
            scope,
            anchors,
            tableRanges);
        return new WordFormulaDocumentIndex(omml, visualTeX);
    }

    private static WordFormulaHostDescriptor? ResolveLocalOmml(
        Document document,
        Range selectionRange)
    {
        var candidates = new List<Range>();
        OMaths? maths = null;
        OMath? math = null;
        Range? candidate = null;
        Range? probe = null;
        try
        {
            try { maths = selectionRange.OMaths; } catch { maths = null; }
            AddMatchingOmmlRanges(maths, selectionRange, candidates);

            if (candidates.Count == 0)
            {
                Release(maths);
                maths = null;
                probe = selectionRange.Duplicate;
                try { probe.MoveStart(WdUnits.wdCharacter, -2); } catch { }
                try { probe.MoveEnd(WdUnits.wdCharacter, 2); } catch { }
                try { maths = probe.OMaths; } catch { maths = null; }
                AddMatchingOmmlRanges(maths, selectionRange, candidates);
            }

            DeduplicateRanges(candidates);
            if (candidates.Count == 0) return null;
            if (candidates.Count != 1)
                throw new InvalidDataException(
                    "The selected range resolves to more than one OMML equation.");

            candidate = candidates[0];
            candidates.Clear(); // ownership transferred to candidate

            // Word's native InsertCrossReference can cache an entire numbered
            // OMath inside the Result of a REF field. That OMath is rendering
            // state for the reference, not an independently editable equation
            // host. Keep Word's native field/result structure intact, but do not
            // surface its cached result as a VisualTeX formula.
            if (IsNativeReferenceResultOmml(candidate))
                return null;

            WdOMathType type;
            Release(maths);
            maths = null;
            try
            {
                maths = candidate.OMaths;
                if (maths.Count != 1)
                    throw new InvalidDataException(
                        "The resolved OMML range no longer contains exactly one equation.");
                math = maths[1];
                var exact = math.Range;
                try
                {
                    if (exact.StoryType != candidate.StoryType
                        || exact.Start != candidate.Start
                        || exact.End != candidate.End)
                        throw new InvalidDataException(
                            "The resolved OMML range is not the exact OMath boundary.");
                    type = math.Type;
                }
                finally { Release(exact); }
            }
            catch (COMException error)
            {
                throw new InvalidDataException(
                    "Word could not resolve the selected OMML equation locally.",
                    error);
            }

            // Native OMML has no durable VisualTeX identity. Any FormulaId
            // exposed to the editor is session-local and is assigned by the
            // service after resolving this exact OMath.
            string? formulaId = null;
            var display = type == WdOMathType.wdOMathDisplay;
            return new WordFormulaHostDescriptor
            {
                Kind = WordFormulaHostKind.Omml,
                Range = Address(candidate),
                DisplayMode = display ? "block" : "inline",
                FormulaId = formulaId,
                Metadata = null,
                MetadataAuthoritative = false,
                Numbering = new WordFormulaNumberingDescriptor
                {
                    Numbered = display && HasLocalNativeSequenceNumber(candidate),
                    FormulaId = formulaId,
                },
                WithinTable = IsWithinTable(candidate),
            };
        }
        finally
        {
            foreach (var range in candidates) Release(range);
            Release(candidate);
            Release(math);
            Release(maths);
            Release(probe);
        }
    }

    private static WordFormulaHostDescriptor? ResolveLocalVisualTeX(
        Document document,
        Range selectionRange)
    {
        Range? probe = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        var matches = new List<(InlineShape Shape, Range Range)>();
        try
        {
            probe = selectionRange.Duplicate;
            try { probe.MoveStart(WdUnits.wdCharacter, -2); } catch { }
            try { probe.MoveEnd(WdUnits.wdCharacter, 2); } catch { }
            shapes = probe.InlineShapes;

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
                if (!TouchesSelection(selectionRange, shapeRange))
                {
                    Release(shapeRange);
                    shapeRange = null;
                    Release(shape);
                    shape = null;
                    continue;
                }

                matches.Add((shape, shapeRange));
                shape = null;
                shapeRange = null;
            }

            if (matches.Count == 0) return null;
            if (matches.Count != 1)
                throw new InvalidDataException(
                    "The selected range resolves to more than one VisualTeX formula.");

            var selected = matches[0];
            matches.Clear();
            FormulaMetadata? metadata = null;
            try
            {
                // The embedded OLE payload is authoritative. VTO_ bookmarks are
                // not consulted to decide whether this object is a VisualTeX host.
                metadata = WordFormulaMetadataReader.TryReadEmbeddedNativeOle(
                    selected.Shape);
            }
            catch { }

            var displayMode = metadata?.DisplayMode;
            if (!string.Equals(displayMode, "block", StringComparison.OrdinalIgnoreCase))
                displayMode = "inline";

            return new WordFormulaHostDescriptor
            {
                Kind = WordFormulaHostKind.VisualTeX,
                Range = Address(selected.Range),
                DisplayMode = displayMode!,
                FormulaId = metadata?.FormulaId,
                Metadata = metadata,
                MetadataAuthoritative = metadata is not null,
                Numbering = new WordFormulaNumberingDescriptor
                {
                    Numbered = metadata?.Numbered == true,
                    FormulaId = metadata?.FormulaId,
                },
                WithinTable = IsWithinTable(selected.Range),
            };
        }
        finally
        {
            foreach (var item in matches)
            {
                Release(item.Range);
                Release(item.Shape);
            }
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(probe);
        }
    }

    private static IReadOnlyList<WordFormulaHostDescriptor> CaptureOmmlIndex(
        Range scope,
        IReadOnlyList<WordFormulaRangeAddress> referenceResults,
        IReadOnlyList<WordFormulaRangeAddress> tableRanges)
    {
        var result = new List<WordFormulaHostDescriptor>();
        OMaths? maths = null;
        OMath? math = null;
        Range? range = null;
        try
        {
            maths = scope.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                math = maths[index];
                range = math.Range;
                if (IsInsideReferenceResult(
                        range,
                        referenceResults))
                {
                    Release(range); range = null;
                    Release(math); math = null;
                    continue;
                }
                var address = Address(range);
                string? formulaId = null;
                result.Add(new WordFormulaHostDescriptor
                {
                    Kind = WordFormulaHostKind.Omml,
                    Range = address,
                    DisplayMode = math.Type == WdOMathType.wdOMathDisplay
                        ? "block"
                        : "inline",
                    FormulaId = formulaId,
                    Metadata = null,
                    MetadataAuthoritative = false,
                    Numbering = new WordFormulaNumberingDescriptor
                    {
                        // Numbering is re-proved locally immediately before a
                        // destructive mutation. The batch index is scheduling data,
                        // not a success/ownership oracle.
                        Numbered = false,
                        FormulaId = formulaId,
                    },
                    WithinTable = IsInsideTableRange(
                        range,
                        tableRanges),
                });
                Release(range); range = null;
                Release(math); math = null;
            }
            return result;
        }
        finally
        {
            Release(range);
            Release(math);
            Release(maths);
        }
    }

    private static IReadOnlyList<WordFormulaRangeAddress>
        CaptureTableRanges(
            Range scope)
    {
        var result =
            new List<WordFormulaRangeAddress>();
        Tables? tables = null;
        Table? table = null;
        Range? tableRange = null;
        try
        {
            tables = scope.Tables;
            for (var index = 1;
                 index <= tables.Count;
                 index++)
            {
                Release(tableRange);
                tableRange = null;
                Release(table);
                table = tables[index];
                tableRange =
                    table.Range.Duplicate;
                result.Add(
                    Address(tableRange));
            }

            return result
                .OrderBy(item => item.StoryType)
                .ThenBy(item => item.Start)
                .ThenBy(item => item.End)
                .ToArray();
        }
        finally
        {
            Release(tableRange);
            Release(table);
            Release(tables);
        }
    }

    private static bool IsInsideTableRange(
        Range hostRange,
        IReadOnlyList<WordFormulaRangeAddress> tableRanges)
    {
        foreach (var table in tableRanges)
        {
            if (table.StoryType < hostRange.StoryType)
                continue;
            if (table.StoryType > hostRange.StoryType)
                break;
            if (table.Start > hostRange.Start)
                break;
            if (table.Start <= hostRange.Start
                && table.End >= hostRange.End)
                return true;
        }
        return false;
    }

    private static IReadOnlyList<WordFormulaRangeAddress>
        CaptureReferenceResultRanges(
            Range scope)
    {
        var result =
            new List<WordFormulaRangeAddress>();
        Fields? fields = null;
        Field? field = null;
        Range? fieldResult = null;
        try
        {
            fields = scope.Fields;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(fieldResult);
                fieldResult = null;
                Release(field);
                field = fields[index];
                if (field.Type !=
                    WdFieldType.wdFieldRef)
                    continue;

                try
                {
                    fieldResult =
                        field.Result.Duplicate;
                    result.Add(
                        Address(fieldResult));
                }
                catch
                {
                    // A broken REF can reject Result access. It cannot be used
                    // as positive evidence that an OMath is only cached rendering.
                }
            }

            return result
                .OrderBy(item => item.StoryType)
                .ThenBy(item => item.Start)
                .ThenBy(item => item.End)
                .ToArray();
        }
        finally
        {
            Release(fieldResult);
            Release(field);
            Release(fields);
        }
    }

    private static bool IsInsideReferenceResult(
        Range equationRange,
        IReadOnlyList<WordFormulaRangeAddress> referenceResults)
    {
        foreach (var result in referenceResults)
        {
            if (result.StoryType < equationRange.StoryType)
                continue;
            if (result.StoryType > equationRange.StoryType)
                break;
            if (result.Start > equationRange.Start)
                break;
            if (result.Start <= equationRange.Start
                && result.End >= equationRange.End)
                return true;
        }
        return false;
    }

    private static bool IsNativeReferenceResultOmml(
        Range equationRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? result = null;
        try
        {
            paragraphs = equationRange.Paragraphs;
            for (var paragraphIndex = 1;
                 paragraphIndex <= paragraphs.Count;
                 paragraphIndex++)
            {
                Release(paragraphRange);
                paragraphRange = null;
                Release(paragraph);
                paragraph = paragraphs[paragraphIndex];
                paragraphRange = paragraph.Range.Duplicate;
                Release(fields);
                fields = paragraphRange.Fields;

                for (var fieldIndex = 1;
                     fieldIndex <= fields.Count;
                     fieldIndex++)
                {
                    Release(result);
                    result = null;
                    Release(field);
                    field = fields[fieldIndex];
                    if (field.Type != WdFieldType.wdFieldRef)
                        continue;

                    result = field.Result.Duplicate;
                    if (result.StoryType == equationRange.StoryType
                        && result.Start <= equationRange.Start
                        && result.End >= equationRange.End)
                        return true;
                }
            }
            return false;
        }
        catch
        {
            // Failure to inspect the local paragraph must not turn a genuine
            // equation into a false negative. The ordinary exact-host checks
            // still run after this bounded classification probe.
            return false;
        }
        finally
        {
            Release(result);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static IReadOnlyList<WordFormulaHostDescriptor> CaptureVisualTeXIndex(
        Range scope,
        IReadOnlyList<IdentityAnchor> anchors,
        IReadOnlyList<WordFormulaRangeAddress> tableRanges)
    {
        var result = new List<WordFormulaHostDescriptor>();
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? range = null;
        try
        {
            shapes = scope.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                shape = shapes[index];
                if (!WordFormulaMetadataReader.IsNativeOle(shape))
                {
                    Release(shape); shape = null;
                    continue;
                }

                range = shape.Range.Duplicate;
                var cached = WordFormulaMetadataReader.TryReadCachedPreview(shape);
                var indexedId = ResolveIndexedVisualTeXIdentity(Address(range), anchors);
                var formulaId = indexedId ?? cached?.FormulaId;
                var displayMode = cached?.DisplayMode;
                if (!string.Equals(displayMode, "block", StringComparison.OrdinalIgnoreCase))
                    displayMode = "inline";

                result.Add(new WordFormulaHostDescriptor
                {
                    Kind = WordFormulaHostKind.VisualTeX,
                    Range = Address(range),
                    DisplayMode = displayMode!,
                    FormulaId = formulaId,
                    Metadata = cached,
                    MetadataAuthoritative = false,
                    Numbering = new WordFormulaNumberingDescriptor
                    {
                        Numbered = cached?.Numbered == true,
                        FormulaId = formulaId,
                    },
                    WithinTable = IsInsideTableRange(
                        range,
                        tableRanges),
                });

                Release(range); range = null;
                Release(shape); shape = null;
            }
            return result;
        }
        finally
        {
            Release(range);
            Release(shape);
            Release(shapes);
        }
    }

    private static IReadOnlyList<IdentityAnchor> CaptureIdentityAnchors(
        Range scope)
    {
        var result = new List<IdentityAnchor>();
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        try
        {
            bookmarks = scope.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                bookmark = bookmarks[index];
                var name = bookmark.Name ?? string.Empty;
                string? prefix = null;
                string? formulaId = null;
                if (WordFormulaMetadataReader.TryFormulaIdFromIdentityBookmark(
                        name,
                        out var visualId))
                {
                    prefix = "VTO_";
                    formulaId = visualId;
                }
                if (formulaId is not null)
                {
                    range = bookmark.Range;
                    result.Add(new IdentityAnchor
                    {
                        Prefix = prefix!,
                        FormulaId = formulaId,
                        StoryType = range.StoryType,
                        Start = range.Start,
                        End = range.End,
                    });
                    Release(range); range = null;
                }
                Release(bookmark); bookmark = null;
            }
            return result;
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static string? ResolveIndexedVisualTeXIdentity(
        WordFormulaRangeAddress shape,
        IReadOnlyList<IdentityAnchor> anchors)
    {
        var matches = anchors
            .Where(anchor =>
                string.Equals(anchor.Prefix, "VTO_", StringComparison.Ordinal)
                && anchor.StoryType == shape.StoryType
                && ((anchor.Start == shape.Start && anchor.End == shape.End)
                    || (anchor.Start == anchor.End && anchor.Start == shape.Start)
                    || (anchor.Start <= shape.Start && anchor.End >= shape.End)))
            .Select(anchor => anchor.FormulaId)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        return matches.Length == 1 ? matches[0] : null;
    }

    private static void AddMatchingOmmlRanges(
        OMaths? maths,
        Range selectionRange,
        ICollection<Range> result)
    {
        if (maths is null) return;
        OMath? math = null;
        Range? range = null;
        try
        {
            for (var index = 1; index <= maths.Count; index++)
            {
                math = maths[index];
                range = math.Range.Duplicate;
                if (TouchesSelection(selectionRange, range))
                {
                    result.Add(range);
                    range = null;
                }
                Release(range); range = null;
                Release(math); math = null;
            }
        }
        finally
        {
            Release(range);
            Release(math);
        }
    }

    private static void DeduplicateRanges(List<Range> ranges)
    {
        var unique = new Dictionary<(WdStoryType Story, int Start, int End), Range>();
        foreach (var range in ranges)
        {
            var key = (range.StoryType, range.Start, range.End);
            if (unique.ContainsKey(key))
            {
                Release(range);
                continue;
            }
            unique[key] = range;
        }
        ranges.Clear();
        ranges.AddRange(unique.Values);
    }

    private static bool TouchesSelection(Range selection, Range candidate)
    {
        if (selection.StoryType != candidate.StoryType) return false;
        if (selection.Start == selection.End)
            return candidate.Start <= selection.Start
                && selection.Start <= candidate.End;
        return selection.Start < candidate.End
            && selection.End > candidate.Start;
    }

    private static bool HasLocalNativeSequenceNumber(
        Range equationRange)
    {
        try
        {
            // Numbering truth for OMML comes from this exact OMath only. The
            // structural detector recognizes Word's native m:eqArr + # + SEQ
            // delimiter regardless of the localized Equation sequence name.
            return WordOmmlConverter
                .HasVisualTeXDirectSequenceEquationNumber(
                    equationRange.WordOpenXML
                    ?? string.Empty,
                    formulaId: null);
        }
        catch
        {
            return false;
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
