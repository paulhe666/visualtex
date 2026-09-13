using System.Diagnostics;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    private static void ValidateCompletedOmmlRedraw(
        Document document, IReadOnlyList<FormulaMetadata> metadata, int expectedCount)
    {
        if (metadata.Count != expectedCount || metadata.Select(item => item.FormulaId).Distinct().Count() != expectedCount)
            throw new InvalidDataException("The completed OMML metadata inventory is incomplete.");
        var watch = Stopwatch.StartNew();
        var package = XDocument.Parse(WordDocumentXml.Read(document));
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var equations = package.Descendants(w + "body").Single().Descendants(m + "oMath").ToArray();
        OMaths? maths = null;
        Bookmarks? bookmarks = null;
        try
        {
            maths = document.OMaths;
            bookmarks = document.Bookmarks;
            if (maths.Count != equations.Length)
                throw new InvalidDataException("COM and XML disagree on the completed OMath inventory.");
            var owners = new Dictionary<(WdStoryType Story, int Start), (int Index, bool Display)>();
            for (var i = 1; i <= maths.Count; i++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[i]; range = math.Range;
                    owners.Add((range.StoryType, range.Start), (i - 1, math.Type == WdOMathType.wdOMathDisplay));
                }
                finally { Release(range); Release(math); }
            }
            var claimed = new HashSet<int>();
            foreach (var item in metadata)
            {
                Bookmark? bookmark = null;
                Range? anchor = null;
                try
                {
                    var name = WordOmmlFormulaStore.BookmarkName(item.FormulaId);
                    if (!bookmarks.Exists(name)) throw new InvalidDataException("The completed OMath lost its identity.");
                    bookmark = bookmarks[name]; anchor = bookmark.Range;
                    if (anchor.Start != anchor.End)
                        throw new InvalidDataException("The completed OMath identity bookmark is not collapsed.");

                    var ownerResolved = owners.TryGetValue(
                        (anchor.StoryType, anchor.Start),
                        out var owner);
                    if (!ownerResolved)
                    {
                        // Word display math can legally keep the canonical collapsed
                        // bookmark one story character before OMath.Start (the native
                        // vertical-tab display separator). WordOmmlFormulaStore.Wrap
                        // and IsCanonicalAnchor already define that as a valid owner;
                        // final redraw validation must use the same contract instead
                        // of rejecting a healthy formula after the Undo record closes.
                        if (owners.TryGetValue(
                                (anchor.StoryType, anchor.Start + 1),
                                out var adjacentOwner))
                        {
                            OMath? adjacentMath = null;
                            Range? adjacentRange = null;
                            try
                            {
                                adjacentMath = maths[adjacentOwner.Index + 1];
                                adjacentRange = adjacentMath.Range;
                                if (WordOmmlFormulaStore.IsCanonicalAnchor(
                                        bookmark,
                                        adjacentRange))
                                {
                                    owner = adjacentOwner;
                                    ownerResolved = true;
                                }
                            }
                            finally
                            {
                                Release(adjacentRange);
                                Release(adjacentMath);
                            }
                        }
                    }
                    var expectedDisplay = string.Equals(
                        item.DisplayMode,
                        "block",
                        StringComparison.Ordinal);
                    var duplicateOwner = ownerResolved && claimed.Contains(owner.Index);
                    if (!ownerResolved
                        || duplicateOwner
                        || owner.Display != expectedDisplay)
                    {
                        var nearby = string.Join(
                            ",",
                            owners
                                .Where(pair => pair.Key.Story == anchor.StoryType)
                                .OrderBy(pair => Math.Abs(pair.Key.Start - anchor.Start))
                                .Take(4)
                                .Select(pair =>
                                    $"{pair.Key.Start}:idx={pair.Value.Index}:display={pair.Value.Display}"));
                        WordDoubleClickHook.TraceMessage(
                            $"redraw-omml-owner-validation-failed formulaId={item.FormulaId} "
                            + $"anchor={anchor.StoryType}/{anchor.Start}:{anchor.End} "
                            + $"ownerResolved={ownerResolved} ownerIndex={(ownerResolved ? owner.Index : -1)} "
                            + $"ownerDisplay={(ownerResolved ? owner.Display.ToString() : "<none>")} "
                            + $"expectedDisplay={expectedDisplay} duplicate={duplicateOwner} nearby=[{nearby}]");
                        throw new InvalidDataException("The completed OMath identity has an incorrect or shared physical owner.");
                    }
                    claimed.Add(owner.Index);
                    WordDoubleClickHook.TraceMessage(
                        $"redraw-omml-owner-verified formulaId={item.FormulaId} "
                        + $"anchor={anchor.StoryType}/{anchor.Start}:{anchor.End} "
                        + $"ownerIndex={owner.Index} display={owner.Display}");
                    var actual = equations[owner.Index].ToString(SaveOptions.DisableFormatting);
                    if (!string.Equals(WordOmmlConverter.ComputeOmmlFingerprint(actual), item.NativeOmmlFingerprint, StringComparison.Ordinal))
                        throw new InvalidDataException($"Word changed the materialized OMML for {item.FormulaId}; the source fingerprint cannot be persisted.");
                }
                finally { Release(anchor); Release(bookmark); }
            }
            WordDoubleClickHook.TraceMessage($"redraw-omml-materialized-verified formulas={claimed.Count} elapsedMs={watch.ElapsedMilliseconds}");
        }
        finally { Release(bookmarks); Release(maths); }
    }

    private static void RetainMathTypeRedrawValidationIdentity(
        ICollection<(string FormulaId, Range EndAnchor, int OwnerLength, string Signature)> retained,
        string formulaId,
        Range owner,
        string signature)
    {
        Range? endAnchor = null;
        try
        {
            var ownerLength = owner.End - owner.Start;
            if (ownerLength <= 0)
                throw new InvalidDataException(
                    $"MathType redraw identity {formulaId} has no physical owner length.");
            endAnchor = owner.Duplicate;
            endAnchor.SetRange(owner.End, owner.End);
            retained.Add((formulaId, endAnchor, ownerLength, signature));
            endAnchor = null;
        }
        finally { Release(endAnchor); }
    }

    // Called only after the entire custom Undo record is closed. A single final
    // export replaces N per-object exports, without skipping any native data check.
    // Redraw mutates strictly from right to left. Retaining a spanning Word Range
    // across those structural edits is unsafe: Word can pull its Start all the way
    // into earlier body paragraphs when the owner lives in a user table. A collapsed
    // right-edge anchor remains a point while edits occur only to its left; pair it
    // with the immutable owner length captured immediately after insertion.
    private static void ValidateCompletedMathTypeRedraw(
        Document document,
        IReadOnlyList<(string FormulaId, Range EndAnchor, int OwnerLength, string Signature)> expected,
        int expectedCount)
    {
        if (expected.Count != expectedCount || expectedCount == 0)
            throw new InvalidDataException("The MathType redraw did not retain every inserted equation for validation.");
        var watch = Stopwatch.StartNew();
        var snapshots = MathTypeWordOpenXml.ReadOleSnapshots(WordDocumentXml.Read(document));
        InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            var ordinals = new Dictionary<(WdStoryType Story, int Start, int End), int>();
            var ordinal = 0;
            for (var i = 1; i <= shapes.Count; i++)
            {
                InlineShape? shape = null;
                Range? range = null;
                try
                {
                    shape = shapes[i];
                    if (shape.Type is not WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                        and not WdInlineShapeType.wdInlineShapeLinkedOLEObject) continue;
                    range = shape.Range;
                    var key = (range.StoryType, range.Start, range.End);
                    if (ordinals.ContainsKey(key))
                        throw new InvalidDataException("Two physical OLE objects share one redraw range.");
                    ordinals.Add(key, ordinal++);
                }
                finally { Release(range); Release(shape); }
            }
            if (ordinal != snapshots.Count)
                throw new InvalidDataException("COM and Flat OPC disagree on the completed OLE inventory.");

            var claimed = new HashSet<int>();
            var formulaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            WordDoubleClickHook.TraceMessage(
                "redraw-mathtype-owner-inventory actual=["
                + string.Join(",", ordinals
                    .OrderBy(pair => pair.Value)
                    .Select(pair => $"{pair.Value}:{pair.Key.Story}/{pair.Key.Start}:{pair.Key.End}"))
                + "]");
            foreach (var item in expected)
            {
                if (item.EndAnchor.Start != item.EndAnchor.End || item.OwnerLength <= 0)
                    throw new InvalidDataException(
                        $"MathType redraw identity {item.FormulaId} lost its collapsed owner anchor.");
                var ownerEnd = item.EndAnchor.Start;
                var ownerStart = ownerEnd - item.OwnerLength;
                var expectedKey = (item.EndAnchor.StoryType, ownerStart, ownerEnd);
                var uniqueFormulaId = formulaIds.Add(item.FormulaId);
                var index = -1;
                var ownerResolved = ownerStart >= 0
                    && ordinals.TryGetValue(expectedKey, out index);
                var independentOwner = ownerResolved && claimed.Add(index);
                WordDoubleClickHook.TraceMessage(
                    $"redraw-mathtype-owner-probe formulaId={item.FormulaId} "
                    + $"anchor={item.EndAnchor.StoryType}/{item.EndAnchor.Start}:{item.EndAnchor.End} "
                    + $"length={item.OwnerLength} expected={item.EndAnchor.StoryType}/{ownerStart}:{ownerEnd} "
                    + $"uniqueId={uniqueFormulaId} ownerResolved={ownerResolved} "
                    + $"ownerIndex={(ownerResolved ? index : -1)} independent={independentOwner}");
                if (!uniqueFormulaId || !ownerResolved || !independentOwner)
                    throw new InvalidDataException($"MathType redraw identity {item.FormulaId} has no independent physical owner.");
                var snapshot = snapshots[index];
                if (!string.Equals(snapshot.ProgId, "Equation.DSMT4", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("Word did not retain the generated MathType OLE class.");
                var actual = MathTypeOleStorage.ReadMathMl(snapshot.CompoundFile);
                if (!MathTypeMathMlRoundTripMatches(item.Signature, actual))
                    throw new InvalidDataException($"Word materialized different formula content for MathType redraw {item.FormulaId}. "
                        + DescribeSemanticSignatureDifference(item.Signature, MathTypeMtefCodec.SemanticSignature(actual)));
            }
            WordDoubleClickHook.TraceMessage(
                $"redraw-mathtype-materialized-verified formulas={claimed.Count} totalOles={ordinal} elapsedMs={watch.ElapsedMilliseconds}");
        }
        finally { Release(shapes); }
    }
}
