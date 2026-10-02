using System.Diagnostics;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    // A bulk import owns its new text, not the surrounding document. Build all
    // positions in managed memory; after native materialization, bookmark each
    // exact placeholder once. Number/heading fields may subsequently insert text
    // before a pending formula, so bare integer offsets are not durable locators.
    private sealed class MathTypeBulkSkeletonPlan
    {
        internal string Text { get; set; } = string.Empty;
        internal List<(int Start, int Length, WordBulkBlock Block)> Blocks { get; } = new();
        internal List<(int Start, int Length, WordBulkRun Run)> Runs { get; } = new();
        internal List<MathTypeBulkSkeletonSlot> Slots { get; } = new();
    }

    private sealed class MathTypeBulkSkeletonSlot
    {
        internal int Start { get; set; }
        internal bool Display { get; set; }
        internal PreparedWordBulkFormula Formula { get; set; } = null!;
        internal string BookmarkName { get; } = "VTBS" + Guid.NewGuid().ToString("N");
    }

    private static MathTypeBulkSkeletonPlan PrepareMathTypeBulkSkeleton(
        WordBulkImportDocument source,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (prepared is null) throw new ArgumentNullException(nameof(prepared));
        var plan = new MathTypeBulkSkeletonPlan();
        var text = new StringBuilder();
        var runIds = new HashSet<string>(StringComparer.Ordinal);
        var formulaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var block in source.Blocks)
        {
            var blockStart = text.Length;
            if (block.Kind == WordBulkBlockKind.DisplayFormula
                && (block.Runs.Count != 1 || !block.Runs[0].IsFormula))
                throw new InvalidDataException("A display block must contain exactly one formula.");
            foreach (var run in block.Runs)
            {
                if (!run.IsFormula)
                {
                    // Newlines inside a single source block are native line
                    // breaks. Only the block boundary owns a paragraph mark.
                    var nativeText = (run.Text ?? string.Empty)
                        .Replace("\r\n", "\n").Replace('\r', '\n').Replace('\n', '\v');
                    if (nativeText.IndexOf('\a') >= 0 || nativeText.IndexOf('\0') >= 0)
                        throw new InvalidDataException("Bulk prose contains a reserved Word structure character.");
                    plan.Runs.Add((text.Length, nativeText.Length, run));
                    text.Append(nativeText);
                    continue;
                }
                if (!runIds.Add(run.Id)
                    || !prepared.TryGetValue(run.Id, out var formula))
                    throw new InvalidDataException($"Missing or duplicated bulk formula {run.Id}.");
                var display = block.Kind == WordBulkBlockKind.DisplayFormula;
                if (!string.Equals(formula.Session.ObjectMode,
                        FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal))
                    throw new InvalidDataException("The MathType skeleton cannot contain another formula object mode.");
                formula.Session.DisplayMode = display ? "block" : "inline";
                formula.Session.Numbered = display && source.NumberDisplayFormulas;
                formula.Session.ToMetadata().Validate();
                if (!formulaIds.Add(formula.Session.FormulaId))
                    throw new InvalidDataException("Bulk formulas must have distinct FormulaIds.");
                plan.Slots.Add(new MathTypeBulkSkeletonSlot
                {
                    Start = text.Length,
                    Display = display,
                    Formula = formula,
                });
                text.Append(BulkInlineFormulaPlaceholder);
            }
            text.Append('\r');
            plan.Blocks.Add((blockStart, text.Length - blockStart, block));
        }
        plan.Text = text.ToString();
        return plan;
    }

    private WordBulkInsertResult InsertMathTypeBulkSkeleton(
        WordBulkImportDocument source,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        string? expectedDocumentId,
        string? sourceObjectId)
    {
        var plan = PrepareMathTypeBulkSkeleton(source, prepared);
        if (plan.Blocks.Count == 0) return new WordBulkInsertResult();
        Document? document = null;
        Selection? selection = null;
        Range? target = null;
        Range? owner = null;
        MathTypeSectionStaging? sectionStaging = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, expectedDocumentId);
            selection = _application.Selection;
            target = ResolveSourceRange(document, sourceObjectId, selection);
            if (target.StoryType != WdStoryType.wdMainTextStory)
                throw new InvalidOperationException("Bulk import requires a main-story insertion range.");
            // Word's mandatory final paragraph mark is structure, not replaceable
            // prose. Keep it as the following typing paragraph even for Select All.
            var storyEnd = ReadMainStoryEnd(document);
            if (target.End == storyEnd && target.End > target.Start)
                target.End -= 1;
            // Replacing table-cell terminators with prose destroys table topology.
            // A collapsed caret or a selection contained inside a cell is valid;
            // a cross-cell selection must be rejected before the transaction.
            if ((target.Text ?? string.Empty).IndexOf('\a') >= 0)
                throw new InvalidOperationException("Bulk import cannot replace table-cell boundaries. Select text inside one cell.");
            owner = WordCharacterFormatting.ResolveParagraphAtPosition(target, target.Start);
            var joinsExistingPrefix = target.Start > owner.Start;
            // A structural block owns a whole paragraph. Never turn pre-existing
            // prefix prose into a heading/list/quote/code/display paragraph.
            var leadingBreak = joinsExistingPrefix
                && plan.Blocks[0].Block.Kind != WordBulkBlockKind.Paragraph;
            // Opening or closing a staging document terminates Word's active
            // Custom UndoRecord. Prepare one immutable native section field tree
            // now; parameterize each copied target tree using native ranges only.
            if (source.NumberDisplayFormulas && plan.Slots.Any(slot => slot.Display)
                && EquationNumberFormat.Resolve(WordEquationNumberFormatAuthority.Resolve(document)).UsesHeading)
                sectionStaging = new MathTypeSectionStaging(document, CreateMathTypeSectionFlatOpc(0, 0));
            var validation = new List<(string FormulaId, string Signature)>();
            return ExecuteDocumentEdit(document, "VisualTeX 批量导入 MathType",
                () => MaterializeMathTypeBulkSkeleton(document, selection, target,
                    plan, source.NumberDisplayFormulas, leadingBreak,
                    joinsExistingPrefix && !leadingBreak, validation, sectionStaging?.Paragraph),
                _ =>
                {
                    if (validation.Count != plan.Slots.Count)
                        throw new InvalidDataException("Bulk insertion did not retain every formula for completed validation.");
                    ValidateCompletedMathTypeBulkSkeleton(document, validation);
                });
        }
        finally
        {
            try
            {
                if (sectionStaging is not null)
                {
                    sectionStaging.Dispose();
                    document?.Activate();
                }
            }
            finally
            {
                Release(owner);
                Release(target);
                Release(selection);
                Release(document);
            }
        }
    }

    private WordBulkInsertResult MaterializeMathTypeBulkSkeleton(
        Document document,
        Selection selection,
        Range target,
        MathTypeBulkSkeletonPlan plan,
        bool numbered,
        bool leadingBreak,
        bool joinsExistingPrefix,
        ICollection<(string FormulaId, string Signature)> validation,
        Range? preparedSectionParagraph)
    {
        var watch = Stopwatch.StartNew();
        long checkpoint = 0;
        var tailBookmarkName = "VTBS" + Guid.NewGuid().ToString("N");
        Bookmarks? bookmarks = null;
        Range? skeleton = null;
        try
        {
            // The tail is an OWNED placeholder in the following typing
            // paragraph, not Word's mandatory final paragraph mark or an original
            // user character. The same locator contract therefore holds for an
            // empty document, tracked replacement, a cell and a display ending.
            const string tailPlaceholder = "\uE001";
            bookmarks = document.Bookmarks;
            var nativeText = (leadingBreak ? "\r" + plan.Text : plan.Text) + tailPlaceholder;
            target.Text = nativeText;
            // Verify actual Word character coordinates before trusting the plan.
            // This reads native text, NOT WordOpenXML, inside the UndoRecord.
            if (!string.Equals(target.Text, nativeText, StringComparison.Ordinal))
                throw new InvalidDataException("Word did not materialize the exact bulk prose skeleton.");
            var basePosition = target.Start + (leadingBreak ? 1 : 0);
            skeleton = document.Range(basePosition, basePosition + plan.Text.Length);
            TraceAcceptancePerformance("MathType.BulkSkeleton", "materialize-text", watch, ref checkpoint);
            Range? tail = null;
            Bookmark? tailBookmark = null;
            try
            {
                tail = document.Range(skeleton.End, skeleton.End + tailPlaceholder.Length);
                if (!string.Equals(tail.Text, tailPlaceholder, StringComparison.Ordinal))
                    throw new InvalidDataException("Bulk import did not materialize its exact owned tail placeholder.");
                tailBookmark = bookmarks.Add(tailBookmarkName, tail);
            }
            finally { Release(tailBookmark); Release(tail); }

            foreach (var slot in plan.Slots)
            {
                Range? placeholder = null;
                Bookmark? bookmark = null;
                try
                {
                    placeholder = document.Range(basePosition + slot.Start,
                        basePosition + slot.Start + BulkInlineFormulaPlaceholder.Length);
                    if (!string.Equals(placeholder.Text, BulkInlineFormulaPlaceholder, StringComparison.Ordinal))
                        throw new InvalidDataException("A bulk placeholder does not match its prepared native span.");
                    bookmark = bookmarks.Add(slot.BookmarkName, placeholder);
                }
                finally { Release(bookmark); Release(placeholder); }
            }
            TraceAcceptancePerformance("MathType.BulkSkeleton", "bind-slots", watch, ref checkpoint);

            ApplyMathTypeBulkSkeletonFormatting(document, skeleton, basePosition,
                plan, joinsExistingPrefix);
            // Never keep a large set of live ranges across OLE paragraph merges.
            Release(skeleton);
            skeleton = null;
            TraceAcceptancePerformance("MathType.BulkSkeleton", "native-formatting", watch, ref checkpoint);

            var widths = new Dictionary<int, float>();
            var nativeStylesPrepared = false;
            var repairs = new List<(string FormulaId, string RightBoundaryWhitespace)>();
            for (var index = plan.Slots.Count - 1; index >= 0; index--)
            {
                var slot = plan.Slots[index];
                Bookmark? bookmark = null;
                Range? insertion = null;
                try
                {
                    bookmark = bookmarks[slot.BookmarkName];
                    insertion = bookmark.Range;
                    if (!string.Equals(insertion.Text, BulkInlineFormulaPlaceholder, StringComparison.Ordinal))
                        throw new InvalidDataException($"Bulk formula {slot.Formula.Session.FormulaId} lost its exact placeholder owner.");
                    // Detach the locator while its character still exists.
                    // Word deletes a nonempty bookmark with its final character;
                    // calling Delete on that stale Bookmark RCW afterward fails.
                    bookmark.Delete();
                    Release(bookmark);
                    bookmark = null;
                    insertion.Text = string.Empty;
                    insertion.Collapse(WdCollapseDirection.wdCollapseStart);
                    var width = slot.Display
                        ? ResolveBatchMathTypeDisplayColumnWidth(document, insertion, widths)
                        : null;
                    InsertPreparedFormula(document, selection, slot.Formula,
                        display: slot.Display,
                        preserveExistingDisplayParagraphBoundary: slot.Display,
                        bulkImport: true,
                        mathTypeDisplayColumnWidth: width,
                        mathTypeNativeStylesPrepared: nativeStylesPrepared,
                        deferMathTypeInlineParagraphRepair: (id, whitespace) => repairs.Add((id, whitespace)),
                        retainMathTypeForCompletedValidation: (_, signature) =>
                            validation.Add((slot.Formula.Session.FormulaId, signature)),
                        exactMathTypeBulkInsertion: insertion,
                        preparedMathTypeSection: preparedSectionParagraph);
                    if (slot.Display) nativeStylesPrepared = true;
                }
                finally { Release(insertion); Release(bookmark); }
            }
            TraceAcceptancePerformance("MathType.BulkSkeleton", "fill-oles", watch, ref checkpoint);
            if (repairs.Count > 0)
                FinalizeDeferredMathTypeInlineParagraphRepairs(document, repairs, out _, out _, out _);
            TraceAcceptancePerformance("MathType.BulkSkeleton", "repair-owned-boundaries", watch, ref checkpoint);
            if (numbered && plan.Slots.Any(slot => slot.Display))
            {
                MathTypeEquationNumbering.UpdateEquationNumbers(document);
                WordEquationReferenceFields.UpdateReferences(document);
            }
            Bookmark? completedTail = null;
            Range? completedTailRange = null;
            try
            {
                completedTail = bookmarks[tailBookmarkName];
                completedTailRange = completedTail.Range;
                // Adjacent paragraph replacement may expand the bookmark's left
                // edge; its right endpoint still follows this owned literal.
                // Never delete the expanded span (which may now contain an OLE).
                if (completedTailRange.End - completedTailRange.Start < tailPlaceholder.Length)
                    throw new InvalidDataException("Bulk import lost its owned tail endpoint.");
                completedTailRange.Start = completedTailRange.End - tailPlaceholder.Length;
                if (!string.Equals(completedTailRange.Text, tailPlaceholder, StringComparison.Ordinal))
                    throw new InvalidDataException("Bulk import lost its exact owned tail placeholder.");
                completedTail.Delete();
                Release(completedTail); completedTail = null;
                completedTailRange.Text = string.Empty;
                completedTailRange.Collapse(WdCollapseDirection.wdCollapseStart);
                selection.SetRange(completedTailRange.Start, completedTailRange.Start);
            }
            finally { Release(completedTailRange); Release(completedTail); }
            TraceAcceptancePerformance("MathType.BulkSkeleton", "complete", watch, ref checkpoint);
            return new WordBulkInsertResult
            {
                BlockCount = plan.Blocks.Count,
                FormulaCount = plan.Slots.Count,
                FormulaIds = plan.Slots.Select(slot => slot.Formula.Session.FormulaId).ToList(),
            };
        }
        finally { Release(skeleton); Release(bookmarks); }
    }

    private static void ApplyMathTypeBulkSkeletonFormatting(
        Document document,
        Range skeleton,
        int basePosition,
        MathTypeBulkSkeletonPlan plan,
        bool joinsExistingPrefix)
    {
        Range? freshParagraphs = null;
        ParagraphFormat? paragraphFormat = null;
        ListFormat? listFormat = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            // Initialize new paragraphs together, before any OLE causes expensive
            // repagination. Do not reset a pre-existing prefix's paragraph style.
            var freshStart = joinsExistingPrefix
                ? basePosition + plan.Blocks[0].Length : basePosition;
            if (freshStart < skeleton.End)
            {
                freshParagraphs = document.Range(freshStart, skeleton.End);
                object normal = WdBuiltinStyle.wdStyleNormal;
                freshParagraphs.set_Style(ref normal);
                listFormat = freshParagraphs.ListFormat;
                listFormat.RemoveNumbers();
                paragraphFormat = freshParagraphs.ParagraphFormat;
                paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraphFormat.LeftIndent = 0;
                paragraphFormat.RightIndent = 0;
                paragraphFormat.FirstLineIndent = 0;
            }
            font = skeleton.Font;
            ResetTransientFont(font);
        }
        finally
        {
            Release(font); Release(paragraphFormat); Release(listFormat); Release(freshParagraphs);
        }
        foreach (var span in plan.Runs)
        {
            var run = span.Run;
            if (span.Length == 0 || !(run.Bold || run.Italic || run.Strike || run.Underline || run.Code))
                continue;
            Range? range = null;
            Microsoft.Office.Interop.Word.Font? runFont = null;
            try
            {
                range = document.Range(basePosition + span.Start, basePosition + span.Start + span.Length);
                runFont = range.Font;
                if (run.Bold) runFont.Bold = 1;
                if (run.Italic) runFont.Italic = 1;
                if (run.Strike) runFont.StrikeThrough = 1;
                if (run.Underline) runFont.Underline = WdUnderline.wdUnderlineSingle;
                if (run.Code)
                {
                    runFont.Name = "Consolas";
                    try { runFont.NameAscii = "Consolas"; } catch { }
                    try { runFont.NameFarEast = "Microsoft YaHei UI"; } catch { }
                }
            }
            finally { Release(runFont); Release(range); }
        }
        foreach (var span in plan.Blocks)
        {
            if (span.Block.Kind is WordBulkBlockKind.Paragraph or WordBulkBlockKind.DisplayFormula)
                continue;
            ApplyBulkParagraphFormatting(document, basePosition + span.Start,
                basePosition + span.Start + span.Length, span.Block);
        }
    }

    private static void RetargetPreparedMathTypeSection(
        Document document, Field sectionField, int chapter, int section)
    {
        // A prepared field tree is immutable and reusable. Its copy now belongs
        // to the target document; change only the literal label and reset values,
        // retaining all native nested field delimiters and character formatting.
        Range? code = null;
        Fields? nested = null;
        var replacements = new List<(int Start, string Before, string After)>();
        try
        {
            code = sectionField.Code;
            nested = code.Fields;
            if (nested.Count != 3)
                throw new InvalidDataException("A prepared MathType section must own exactly three native counters.");
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var firstChildStart = int.MaxValue;
            for (var index = 1; index <= nested.Count; index++)
            {
                Field? child = null;
                Range? childCode = null;
                try
                {
                    child = nested[index]; childCode = child.Code;
                    var text = childCode.Text ?? string.Empty;
                    var name = System.Text.RegularExpressions.Regex.Match(text,
                        @"^\s*SEQ\s+(MTEqn|MTChap|MTSec)\s", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (!name.Success || !names.Add(name.Groups[1].Value)
                        || childCode.Start <= code.Start || childCode.End >= code.End)
                        throw new InvalidDataException("The prepared MathType counter has ambiguous field ownership.");
                    firstChildStart = Math.Min(firstChildStart, childCode.Start - 1);
                    var counter = name.Groups[1].Value;
                    if (counter.Equals("MTEqn", StringComparison.OrdinalIgnoreCase)) continue;
                    var reset = System.Text.RegularExpressions.Regex.Match(text, @"\\r\s+(-?\d+)(?=\s|$)",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (!reset.Success)
                        throw new InvalidDataException("The prepared MathType heading counter has no literal reset value.");
                    var value = counter.Equals("MTChap", StringComparison.OrdinalIgnoreCase) ? chapter : section;
                    replacements.Add((childCode.Start + reset.Groups[1].Index, reset.Groups[1].Value,
                        value.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                }
                finally { Release(childCode); Release(child); }
            }
            var prefix = document.Range(code.Start, firstChildStart);
            try
            {
                var original = prefix.Text ?? string.Empty;
                const string marker = " MACROBUTTON MTEditEquationSection2 ";
                if (!original.StartsWith(marker, StringComparison.Ordinal))
                    throw new InvalidDataException("The prepared MathType section has no exact native label prefix.");
                var label = string.Equals(System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName,
                    "zh", StringComparison.OrdinalIgnoreCase)
                    ? $"公式章 {chapter} 节 {section}" : $"Equation Chapter {chapter} Section {section}";
                replacements.Add((prefix.Start, original, marker + label));
            }
            finally { Release(prefix); }
        }
        finally { Release(nested); Release(code); }
        foreach (var replacement in replacements.OrderByDescending(item => item.Start))
        {
            if (replacement.Before == replacement.After) continue;
            var target = document.Range(replacement.Start, replacement.Start + replacement.Before.Length);
            try
            {
                if (target.Text != replacement.Before)
                    throw new InvalidDataException("A prepared MathType section literal moved before native parameterization.");
                target.Text = replacement.After;
            }
            finally { Release(target); }
        }
    }

    private static void ValidateCompletedMathTypeBulkSkeleton(
        Document document,
        IReadOnlyCollection<(string FormulaId, string Signature)> validation)
    {
        // ExecuteDocumentEdit invokes this only AFTER EndUndoRecord. Serialize
        // the bounded imported region ONCE, then prove every exact bookmark/OLE/
        // relationship/CFB association in that same immutable native snapshot.
        // Re-exporting every shape separately adds a Word COM/layout turn per OLE.
        var watch = Stopwatch.StartNew();
        long checkpoint = 0;
        var ordered = validation.Reverse().ToArray();
        if (ordered.Length == 0) return;
        var expected = ordered.ToDictionary(
            item => WordFormulaIdentityStore.MathTypeBookmarkName(item.FormulaId),
            item => item, StringComparer.OrdinalIgnoreCase);
        Bookmarks? bookmarks = null;
        Bookmark? first = null;
        Bookmark? last = null;
        Range? firstRange = null;
        Range? lastRange = null;
        Range? scope = null;
        string xml;
        try
        {
            bookmarks = document.Bookmarks;
            first = bookmarks[WordFormulaIdentityStore.MathTypeBookmarkName(ordered[0].FormulaId)];
            last = bookmarks[WordFormulaIdentityStore.MathTypeBookmarkName(ordered[ordered.Length - 1].FormulaId)];
            firstRange = first.Range;
            lastRange = last.Range;
            if (firstRange.Start > lastRange.Start)
                throw new InvalidDataException("The imported MathType identity boundaries are reversed.");
            // Include a one-character witness so boundary bookmarks themselves
            // cannot be omitted merely because they coincide with export edges.
            scope = document.Range(Math.Max(0, firstRange.Start - 1),
                Math.Min(ReadMainStoryEnd(document), lastRange.End + 1));
            xml = scope.WordOpenXML;
        }
        finally
        {
            Release(scope); Release(lastRange); Release(firstRange);
            Release(last); Release(first); Release(bookmarks);
        }
        XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace office = "urn:schemas-microsoft-com:office:office";
        var package = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var documentPart = package.Descendants(pkg + "part").Single(
            part => (string?)part.Attribute(pkg + "name") == "/word/document.xml");
        var snapshots = MathTypeWordOpenXml.ReadOleSnapshots(xml);
        var active = new Dictionary<string, string>(StringComparer.Ordinal);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var actualOrder = new List<string>();
        var oleIndex = 0;
        foreach (var node in documentPart.Descendants())
        {
            var deleted = node.Ancestors().Any(owner => owner.Name == w + "del" || owner.Name == w + "moveFrom");
            if (node.Name == w + "bookmarkStart" && !deleted)
            {
                var name = (string?)node.Attribute(w + "name") ?? string.Empty;
                if (expected.ContainsKey(name))
                {
                    var id = (string?)node.Attribute(w + "id")
                        ?? throw new InvalidDataException("A MathType locator has no native bookmark ID.");
                    active.Add(id, name);
                }
            }
            else if (node.Name == w + "bookmarkEnd" && !deleted)
            {
                active.Remove((string?)node.Attribute(w + "id") ?? string.Empty);
            }
            else if (node.Name == office + "OLEObject")
            {
                if (oleIndex >= snapshots.Count)
                    throw new InvalidDataException("Native bulk OLE snapshot topology is inconsistent.");
                var snapshot = snapshots[oleIndex++];
                if (deleted || active.Count == 0) continue;
                if (active.Count != 1)
                    throw new InvalidDataException("A bulk OLE has ambiguous formula identity ownership.");
                var name = active.Values.Single();
                var item = expected[name];
                if (!seen.Add(name)
                    || !string.Equals(snapshot.ProgId, "Equation.DSMT4", StringComparison.OrdinalIgnoreCase)
                    || !MathTypeOleStorage.LooksLikeMathTypeCompoundFile(snapshot.CompoundFile)
                    || !MathTypeMathMlRoundTripMatches(item.Signature, MathTypeOleStorage.ReadMathMl(snapshot.CompoundFile)))
                    throw new InvalidDataException($"Bulk MathType {item.FormulaId} did not retain one native OLE with its prepared Equation Native semantics.");
                actualOrder.Add(item.FormulaId);
            }
        }
        if (oleIndex != snapshots.Count || active.Count != 0
            || seen.Count != expected.Count
            || !actualOrder.SequenceEqual(ordered.Select(item => item.FormulaId), StringComparer.OrdinalIgnoreCase))
            throw new InvalidDataException("Bulk native storage validation lost an identity, bookmark boundary, or source ordering.");
        TraceAcceptancePerformance("MathType.BulkSkeleton", "validate-native-storage", watch, ref checkpoint);
    }
}
