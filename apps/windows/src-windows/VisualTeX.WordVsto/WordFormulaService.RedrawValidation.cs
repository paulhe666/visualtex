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
                    if (anchor.Start != anchor.End
                        || !owners.TryGetValue((anchor.StoryType, anchor.Start), out var owner)
                        || !claimed.Add(owner.Index)
                        || owner.Display != string.Equals(item.DisplayMode, "block", StringComparison.Ordinal))
                        throw new InvalidDataException("The completed OMath identity has an incorrect or shared physical owner.");
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

    // Called only after the entire custom Undo record is closed. A single final
    // export replaces N per-object exports, without skipping any native data check.
    private static void ValidateCompletedMathTypeRedraw(
        Document document,
        IReadOnlyList<(string FormulaId, Range Range, string Signature)> expected,
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
            foreach (var item in expected)
            {
                if (!formulaIds.Add(item.FormulaId)
                    || !ordinals.TryGetValue((item.Range.StoryType, item.Range.Start, item.Range.End), out var index)
                    || !claimed.Add(index))
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
