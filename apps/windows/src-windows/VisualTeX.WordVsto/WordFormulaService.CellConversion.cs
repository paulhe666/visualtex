using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    // A user cell is structural identity, not a long-lived COM RCW. Word can
    // rebind Cell/Table/Range RCWs after neighbouring OLE/OMML mutations. Keep a
    // transaction-only bookmark spanning the complete original cell and resolve
    // the current physical cell from that bookmark whenever ownership is needed.
    private sealed class CellConversionOwner : IDisposable
    {
        internal string CellBookmarkName { get; }
        internal int RowIndex { get; }
        internal int ColumnIndex { get; }
        internal string FormulaId { get; }
        internal int FormulaOrdinal { get; }
        internal int FormulaCount { get; }
        internal string? MathTypeBookmark { get; set; }

        internal CellConversionOwner(
            string cellBookmarkName,
            int rowIndex,
            int columnIndex,
            string formulaId,
            int formulaOrdinal,
            int formulaCount)
        {
            CellBookmarkName = cellBookmarkName;
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
            FormulaId = formulaId;
            FormulaOrdinal = formulaOrdinal;
            FormulaCount = formulaCount;
        }

        internal Cell AcquireCell(Document document)
        {
            Bookmarks? bookmarks = null;
            Bookmark? bookmark = null;
            Range? bookmarkRange = null;
            Range? probe = null;
            Cell? cell = null;
            Range? cellRange = null;
            try
            {
                bookmarks = document.Bookmarks;
                if (!bookmarks.Exists(CellBookmarkName))
                {
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-user-cell-owner-resolve-failed formulaId={FormulaId} bookmark={CellBookmarkName} reason=missing");
                    throw new InvalidDataException(
                        $"The user-cell conversion identity {CellBookmarkName} disappeared.");
                }
                bookmark = bookmarks[CellBookmarkName];
                bookmarkRange = bookmark.Range;
                if (bookmarkRange.Start != bookmarkRange.End)
                {
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-user-cell-owner-resolve-failed formulaId={FormulaId} bookmark={CellBookmarkName} reason=expanded range={bookmarkRange.Start}:{bookmarkRange.End}");
                    throw new InvalidDataException(
                        $"The user-cell conversion identity {CellBookmarkName} is no longer a structural point.");
                }

                // The collapsed structural bookmark is allowed to move as content
                // before/inside the table changes. Resolve the physical cell from
                // the current coordinate each time. At a Word cell boundary the
                // point can have affinity to either side, so probe forward first
                // and then backward, accepting only the originally captured row /
                // column. No long-lived Cell/Table/Range RCW participates here.
                var position = bookmarkRange.Start;
                var content = document.Content;
                try
                {
                    if (position < content.End)
                    {
                        probe = document.Range(position, Math.Min(content.End, position + 1));
                        var candidate = WordFormulaHost.TryGetOwningCell(probe);
                        if (candidate is not null
                            && candidate.RowIndex == RowIndex
                            && candidate.ColumnIndex == ColumnIndex)
                        {
                            cell = candidate;
                            candidate = null;
                        }
                        Release(candidate);
                        Release(probe); probe = null;
                    }
                    if (cell is null && position > content.Start)
                    {
                        probe = document.Range(Math.Max(content.Start, position - 1), position);
                        var candidate = WordFormulaHost.TryGetOwningCell(probe);
                        if (candidate is not null
                            && candidate.RowIndex == RowIndex
                            && candidate.ColumnIndex == ColumnIndex)
                        {
                            cell = candidate;
                            candidate = null;
                        }
                        Release(candidate);
                    }
                }
                finally { Release(content); }

                if (cell is null)
                {
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-user-cell-owner-resolve-failed formulaId={FormulaId} bookmark={CellBookmarkName} reason=no-matching-cell anchor={position} expected={RowIndex},{ColumnIndex}");
                    throw new InvalidDataException(
                        $"The user-cell conversion identity {CellBookmarkName} no longer resolves to its original cell.");
                }
                cellRange = cell.Range;
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-user-cell-owner-resolved formulaId={FormulaId} bookmark={CellBookmarkName} anchor={position} cell={cellRange.Start}:{cellRange.End} row={cell.RowIndex} column={cell.ColumnIndex}");
                var result = cell;
                cell = null;
                return result;
            }
            finally
            {
                Release(cellRange);
                Release(cell);
                Release(probe);
                Release(bookmarkRange);
                Release(bookmark);
                Release(bookmarks);
            }
        }

        internal void ReanchorAtCurrentCellTerminator(Document document)
        {
            Cell? cell = null;
            Range? cellRange = null;
            Range? anchor = null;
            Bookmarks? bookmarks = null;
            Bookmark? existing = null;
            Bookmark? replacement = null;
            Range? replacementRange = null;
            try
            {
                cell = AcquireCell(document);
                cellRange = cell.Range;
                if (cellRange.End <= cellRange.Start)
                    throw new InvalidDataException(
                        $"The user-cell conversion identity {CellBookmarkName} has no terminator to reanchor.");
                bookmarks = document.Bookmarks;
                if (bookmarks.Exists(CellBookmarkName))
                {
                    existing = bookmarks[CellBookmarkName];
                    existing.Delete();
                    Release(existing);
                    existing = null;
                }
                anchor = document.Range(cellRange.End - 1, cellRange.End - 1);
                replacement = bookmarks.Add(CellBookmarkName, anchor);
                replacementRange = replacement.Range;
                if (replacementRange.Start != cellRange.End - 1
                    || replacementRange.End != replacementRange.Start)
                    throw new InvalidDataException(
                        $"Word did not retain the reanchored user-cell identity {CellBookmarkName}.");
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-user-cell-owner-reanchored formulaId={FormulaId} bookmark={CellBookmarkName} anchor={replacementRange.Start} cell={cellRange.Start}:{cellRange.End} row={RowIndex} column={ColumnIndex}");
            }
            finally
            {
                Release(replacementRange);
                Release(replacement);
                Release(existing);
                Release(bookmarks);
                Release(anchor);
                Release(cellRange);
                Release(cell);
            }
        }

        internal void ReanchorAtVerifiedTargetCellTerminator(Document document, Range verifiedTarget)
        {
            if (verifiedTarget is null) throw new ArgumentNullException(nameof(verifiedTarget));
            Cell? cell = null;
            Range? cellRange = null;
            Range? anchor = null;
            Bookmarks? bookmarks = null;
            Bookmark? existing = null;
            Bookmark? replacement = null;
            Range? replacementRange = null;
            try
            {
                // A collapsed cell-terminator bookmark can become ambiguous after
                // Word inserts a managed numbering table at the same boundary. In
                // particular, both sides can report row=1/column=1. At this stage
                // the fresh target OMath has already passed the conversion's strict
                // content verification, so its actual owning cell is stronger
                // structural evidence than probing either side of VTFCC.
                cell = WordFormulaHost.TryGetOwningCell(verifiedTarget)
                    ?? throw new InvalidDataException(
                        $"The verified converted formula {FormulaId} is no longer wholly inside a user table cell.");
                if (cell.RowIndex != RowIndex || cell.ColumnIndex != ColumnIndex)
                    throw new InvalidDataException(
                        $"The verified converted formula {FormulaId} moved from its original user-cell coordinates.");
                cellRange = cell.Range;
                if (!WordFormulaHost.ContainsPhysicalRange(
                        cellRange.Start,
                        cellRange.End - 1,
                        verifiedTarget.Start,
                        verifiedTarget.End))
                    throw new InvalidDataException(
                        $"The verified converted formula {FormulaId} escaped its original user-cell contents.");
                if (cellRange.End <= cellRange.Start)
                    throw new InvalidDataException(
                        $"The verified converted formula {FormulaId} has no cell terminator to reanchor.");
                var formulaPositions = ReadCellFormulaPositions(cell);
                if (FormulaOrdinal < 0
                    || FormulaOrdinal >= formulaPositions.Count
                    || formulaPositions[FormulaOrdinal].Start != verifiedTarget.Start
                    || formulaPositions[FormulaOrdinal].End != verifiedTarget.End)
                    throw new InvalidDataException(
                        $"The verified converted formula {FormulaId} no longer owns its original cell-local formula slot.");

                bookmarks = document.Bookmarks;
                if (bookmarks.Exists(CellBookmarkName))
                {
                    existing = bookmarks[CellBookmarkName];
                    existing.Delete();
                    Release(existing);
                    existing = null;
                }
                anchor = document.Range(cellRange.End - 1, cellRange.End - 1);
                replacement = bookmarks.Add(CellBookmarkName, anchor);
                replacementRange = replacement.Range;
                if (replacementRange.Start != cellRange.End - 1
                    || replacementRange.End != replacementRange.Start)
                    throw new InvalidDataException(
                        $"Word did not retain the verified-target user-cell identity {CellBookmarkName}.");
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-user-cell-owner-reanchored-from-verified-target formulaId={FormulaId} bookmark={CellBookmarkName} target={verifiedTarget.Start}:{verifiedTarget.End} anchor={replacementRange.Start} cell={cellRange.Start}:{cellRange.End} row={RowIndex} column={ColumnIndex}");
            }
            finally
            {
                Release(replacementRange);
                Release(replacement);
                Release(existing);
                Release(bookmarks);
                Release(anchor);
                Release(cellRange);
                Release(cell);
            }
        }

        internal void RebindMathTypeBookmarkToVerifiedTarget(Document document, Range verifiedTarget)
        {
            if (string.IsNullOrWhiteSpace(MathTypeBookmark))
                throw new InvalidDataException(
                    $"The converted MathType formula {FormulaId} has no retained target locator to repair.");
            Bookmarks? bookmarks = null;
            Bookmark? existing = null;
            Bookmark? replacement = null;
            Range? replacementRange = null;
            InlineShapes? shapes = null;
            InlineShape? shape = null;
            try
            {
                bookmarks = document.Bookmarks;
                if (bookmarks.Exists(MathTypeBookmark))
                {
                    existing = bookmarks[MathTypeBookmark];
                    existing.Delete();
                    Release(existing);
                    existing = null;
                }
                replacement = bookmarks.Add(MathTypeBookmark, verifiedTarget);
                replacementRange = replacement.Range;
                shapes = replacementRange.InlineShapes;
                if (replacementRange.StoryType != verifiedTarget.StoryType
                    || replacementRange.Start > verifiedTarget.Start
                    || replacementRange.End < verifiedTarget.End
                    || shapes.Count != 1)
                    throw new InvalidDataException(
                        $"Word did not retain the repaired MathType locator {MathTypeBookmark} on one target object.");
                shape = shapes[1];
                if (!MathTypeOleInterop.IsMathTypeOle(shape))
                    throw new InvalidDataException(
                        $"The repaired MathType locator {MathTypeBookmark} no longer owns Equation.DSMT4.");
                var shapeRange = shape.Range;
                try
                {
                    if (shapeRange.Start != verifiedTarget.Start || shapeRange.End != verifiedTarget.End)
                        throw new InvalidDataException(
                            $"The repaired MathType locator {MathTypeBookmark} owns a different physical object.");
                }
                finally { Release(shapeRange); }
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-user-cell-mathtype-locator-rebound formulaId={FormulaId} bookmark={MathTypeBookmark} range={verifiedTarget.Start}:{verifiedTarget.End}");
            }
            finally
            {
                Release(shape);
                Release(shapes);
                Release(replacementRange);
                Release(replacement);
                Release(existing);
                Release(bookmarks);
            }
        }

        internal void DeleteTemporaryBookmark(Document document)
        {
            Bookmarks? bookmarks = null;
            Bookmark? bookmark = null;
            try
            {
                bookmarks = document.Bookmarks;
                if (!bookmarks.Exists(CellBookmarkName)) return;
                bookmark = bookmarks[CellBookmarkName];
                bookmark.Delete();
                if (bookmarks.Exists(CellBookmarkName))
                    throw new InvalidDataException(
                        $"The temporary user-cell identity {CellBookmarkName} was not removed.");
            }
            finally
            {
                Release(bookmark);
                Release(bookmarks);
            }
        }

        public void Dispose()
        {
            // No Office RCW is retained by design. Bookmark cleanup is explicit in
            // the owning conversion transaction/finally block.
        }
    }

    private readonly struct CellFormulaPosition
    {
        internal int Start { get; }
        internal int End { get; }
        internal string ObjectMode { get; }

        internal CellFormulaPosition(int start, int end, string objectMode)
        {
            Start = start;
            End = end;
            ObjectMode = objectMode;
        }
    }

    private static IReadOnlyList<CellFormulaPosition> ReadCellFormulaPositions(Cell cell)
    {
        Range? cellRange = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        var positions = new List<CellFormulaPosition>();
        try
        {
            cellRange = cell.Range;
            maths = cellRange.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    positions.Add(new CellFormulaPosition(
                        range.Start,
                        range.End,
                        FormulaOleContract.WordOmmlMode));
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }

            shapes = cellRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                try
                {
                    shape = shapes[index];
                    string? mode = null;
                    if (MathTypeOleInterop.IsMathTypeOle(shape))
                        mode = FormulaOleContract.MathTypeOleMode;
                    else if (WordFormulaMetadataReader.TryRead(shape) is not null)
                        mode = FormulaOleContract.NativeOleMode;
                    if (mode is null) continue;
                    range = shape.Range;
                    positions.Add(new CellFormulaPosition(
                        range.Start,
                        range.End,
                        mode));
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }
        }
        finally
        {
            Release(shapes);
            Release(maths);
            Release(cellRange);
        }

        return positions
            .OrderBy(position => position.Start)
            .ThenBy(position => position.End)
            .ToArray();
    }

    private static int CaptureFormulaOrdinalInCell(
        Cell cell,
        Range source,
        string sourceMode,
        out int formulaCount)
    {
        var positions = ReadCellFormulaPositions(cell);
        formulaCount = positions.Count;
        for (var index = 0; index < positions.Count; index++)
        {
            var position = positions[index];
            if (!string.Equals(position.ObjectMode, sourceMode, StringComparison.Ordinal)
                || position.Start != source.Start
                || position.End != source.End)
                continue;
            return index;
        }
        throw new InvalidDataException(
            "The table source formula could not be assigned a stable cell-local ordinal.");
    }

    private static Range? TryResolveConvertedOmmlByCellOrdinal(
        Document document,
        CellConversionOwner owner)
    {
        Cell? cell = null;
        Range? cellRange = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        try
        {
            cell = owner.AcquireCell(document);
            var positions = ReadCellFormulaPositions(cell);
            if (owner.FormulaOrdinal < 0 || owner.FormulaOrdinal >= positions.Count)
                return null;
            var position = positions[owner.FormulaOrdinal];
            if (!string.Equals(
                    position.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
                return null;

            cellRange = cell.Range;
            maths = cellRange.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                Release(mathRange); mathRange = null;
                Release(math); math = maths[index];
                mathRange = math.Range;
                if (mathRange.Start != position.Start || mathRange.End != position.End)
                    continue;
                return mathRange.Duplicate;
            }
            return null;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(cellRange);
            Release(cell);
        }
    }

    private static Range? TryResolveManagedOmmlByDurableIdentity(
        Document document,
        string formulaId)
    {
        Bookmark? bookmark = null;
        Range? anchor = null;
        OMaths? localMaths = null;
        OMath? localMath = null;
        Range? localRange = null;
        Range? range = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
            if (metadata is null) return null;
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            if (bookmark is null) return null;

            // A freshly verified VTOMML bookmark is a collapsed canonical anchor.
            // Ask Word for the OMath owned by THAT exact collapsed range first.
            // Expanding the probe across a body -> table boundary can omit the
            // table equation and return the preceding body OMath instead.
            anchor = bookmark.Range;
            localMaths = anchor.OMaths;
            if (localMaths.Count == 1)
            {
                localMath = localMaths[1];
                localRange = localMath.Range.Duplicate;
                if (WordOmmlFormulaStore.IsCanonicalAnchor(bookmark, localRange))
                {
                    var exact = localRange;
                    localRange = null;
                    return exact;
                }
            }

            // Keep the mature drift-recovery resolver for older/external documents
            // whose durable anchor genuinely moved. Fresh conversion never reaches
            // this branch when its exact anchor is still intact.
            range = WordOmmlFormulaStore.GetEquationRangeForCurrentRead(
                document,
                bookmark,
                metadata);
            var result = range;
            range = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(range);
            Release(localRange);
            Release(localMath);
            Release(localMaths);
            Release(anchor);
            Release(bookmark);
        }
    }

    private static void VerifyConvertedBodyHosts(
        Document document,
        string targetMode,
        IEnumerable<WordFormulaFormatConversionTarget> targets,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        IReadOnlyDictionary<string, string> mathTypeLocators,
        IReadOnlyDictionary<string, CellConversionOwner> userCells)
    {
        foreach (var target in targets)
        {
            if (userCells.ContainsKey(target.Id)) continue;
            var formulaId = prepared[target.Id].Session.FormulaId;
            Bookmark? bookmark = null;
            Range? locator = null;
            Range? range = null;
            InlineShapes? shapes = null;
            InlineShape? shape = null;
            try
            {
                if (targetMode == FormulaOleContract.WordOmmlMode)
                {
                    range = TryResolveManagedOmmlByDurableIdentity(document, formulaId)
                        ?? throw new InvalidDataException($"Converted body OMML {formulaId} lost its identity.");
                }
                else if (targetMode == FormulaOleContract.NativeOleMode)
                {
                    shape = FindByFormulaId(document, formulaId, allowGlobalFallback: false)
                        ?? throw new InvalidDataException($"Converted body OLE {formulaId} lost its identity.");
                    range = shape.Range.Duplicate;
                }
                else
                {
                    if (!mathTypeLocators.TryGetValue(target.Id, out var name)
                        || !document.Bookmarks.Exists(name))
                        throw new InvalidDataException($"Converted body MathType {formulaId} lost its exact locator.");
                    bookmark = document.Bookmarks[name];
                    locator = bookmark.Range;
                    shapes = locator.InlineShapes;
                    if (shapes.Count != 1)
                        throw new InvalidDataException($"Converted body MathType {formulaId} no longer owns one object.");
                    shape = shapes[1];
                    if (!MathTypeOleInterop.IsMathTypeOle(shape))
                        throw new InvalidDataException($"Converted body MathType {formulaId} has the wrong object type.");
                    range = shape.Range.Duplicate;
                }
                if (WordEquationNumbering.RangeIsWhollyWithinTable(range)
                    && !(targetMode == FormulaOleContract.WordOmmlMode
                        && target.Numbered
                        && WordEquationNumbering.HasReusableNumberedNativeOmmlDirectTableHost(
                            document, range, formulaId)))
                    throw new InvalidDataException(
                        $"Converted body formula {formulaId} entered an unrelated user table.");
            }
            finally
            {
                Release(range); Release(shape); Release(shapes); Release(locator); Release(bookmark);
            }
        }
    }

    private static Range ResolveVerifiedConvertedOleUserCellTarget(
        Document document,
        string targetMode,
        CellConversionOwner owner)
    {
        if (string.Equals(targetMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal))
        {
            InlineShape? nativeShape = null;
            Range? range = null;
            try
            {
                // A converted VisualTeX target carries a newly generated FormulaId
                // inside its own compound storage. That embedded identity is stronger
                // than any Word bookmark after unrelated body-numbering edits.
                nativeShape = ResolveConvertedVisualTeXByEmbeddedIdentity(
                        document,
                        owner.FormulaId,
                        string.Empty)
                    ?? throw new InvalidDataException(
                        $"Converted VisualTeX target {owner.FormulaId} disappeared before user-cell reanchoring.");
                if (!WordFormulaMetadataReader.IsNativeOle(nativeShape))
                    throw new InvalidDataException(
                        $"Converted VisualTeX target {owner.FormulaId} is no longer a VisualTeX OLE object.");
                var metadata = WordFormulaMetadataReader.TryReadEmbeddedNativeOle(nativeShape);
                if (!string.Equals(metadata?.FormulaId, owner.FormulaId, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException(
                        $"Converted VisualTeX target {owner.FormulaId} lost its embedded identity.");
                range = nativeShape.Range.Duplicate;
                var result = range;
                range = null;
                return result;
            }
            finally
            {
                Release(range);
                Release(nativeShape);
            }
        }

        if (!string.Equals(targetMode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal))
            throw new InvalidOperationException(
                $"OLE user-cell target recovery is not defined for {targetMode}.");
        if (string.IsNullOrWhiteSpace(owner.MathTypeBookmark))
            throw new InvalidDataException(
                $"Converted MathType target {owner.FormulaId} has no retained locator.");

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? locator = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        Cell? cell = null;
        Range? resolved = null;
        var matches = 0;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(owner.MathTypeBookmark))
                throw new InvalidDataException(
                    $"Converted MathType locator {owner.MathTypeBookmark} disappeared before user-cell recovery.");
            bookmark = bookmarks[owner.MathTypeBookmark];
            locator = bookmark.Range;
            shapes = locator.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shapeRange); shapeRange = null;
                Release(cell); cell = null;
                Release(shape); shape = shapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                shapeRange = shape.Range.Duplicate;
                cell = WordFormulaHost.TryGetOwningCell(shapeRange);
                if (cell is null
                    || cell.RowIndex != owner.RowIndex
                    || cell.ColumnIndex != owner.ColumnIndex)
                    continue;
                var positions = ReadCellFormulaPositions(cell);
                if (positions.Count != owner.FormulaCount
                    || owner.FormulaOrdinal < 0
                    || owner.FormulaOrdinal >= positions.Count
                    || !string.Equals(
                        positions[owner.FormulaOrdinal].ObjectMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal)
                    || positions[owner.FormulaOrdinal].Start != shapeRange.Start
                    || positions[owner.FormulaOrdinal].End != shapeRange.End)
                    continue;
                matches++;
                Release(resolved);
                resolved = shapeRange.Duplicate;
            }
            if (matches != 1 || resolved is null)
                throw new InvalidDataException(
                    $"Converted MathType locator {owner.MathTypeBookmark} has {matches} unique candidates in original cell ({owner.RowIndex},{owner.ColumnIndex}) slot {owner.FormulaOrdinal}.");
            var result = resolved;
            resolved = null;
            return result;
        }
        finally
        {
            Release(resolved);
            Release(cell);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(locator);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static Range? TryResolveConvertedOmmlUserCellOwner(
        Document document,
        IReadOnlyList<CellConversionOwner> owners,
        string formulaId)
    {
        var owner = owners.FirstOrDefault(item =>
            string.Equals(item.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase));
        if (owner is null) return null;
        return TryResolveConvertedOmmlByCellOrdinal(document, owner);
    }

    private CellConversionOwner? CaptureUserCellForConversion(
        Document document,
        string sourceMode,
        WordFormulaFormatConversionTarget target,
        string targetFormulaId)
    {
        if (!target.SourceWithinTable) return null;
        InlineShape? shape = null;
        Range? source = null;
        Table? managed = null;
        Cell? cell = null;
        Range? cellRange = null;
        Range? cellTerminator = null;
        Bookmarks? bookmarks = null;
        Bookmark? cellBookmark = null;
        Range? verifiedBookmarkRange = null;
        try
        {
            if (sourceMode == FormulaOleContract.WordOmmlMode)
                source = ResolveSimpleOmmlSourceRange(document, target);
            else
            {
                shape = sourceMode == FormulaOleContract.MathTypeOleMode
                    ? FindMathTypeOleByRange(document, target.SourceObjectId, allowGlobalFallback: false)
                    : FindByFormulaId(document, target.SourceFormulaId, target.SourceObjectId, allowGlobalFallback: false);
                if (shape is not null) source = shape.Range;
            }
            if (source is null)
                throw new InvalidDataException("The table source lost its physical owner before conversion.");
            if (sourceMode != FormulaOleContract.MathTypeOleMode && target.Numbered)
            {
                managed = TryGetVisualTeXNumberedTable(source, target.Metadata);
                if (managed is not null) return null; // Generated body host may be replaced.
            }
            cell = WordFormulaHost.TryGetOwningCell(source)
                ?? throw new InvalidDataException("The table source is not wholly inside one user cell.");
            var rowIndex = cell.RowIndex;
            var columnIndex = cell.ColumnIndex;
            var ordinal = CaptureFormulaOrdinalInCell(cell, source, sourceMode, out var formulaCount);
            cellRange = cell.Range.Duplicate;
            if (cellRange.End <= cellRange.Start)
                throw new InvalidDataException("The user table cell has no structural terminator.");
            // Important Word identity rule: any non-collapsed bookmark covering
            // editable cell contents can be pulled onto replacement formula text.
            // A zero-length bookmark at End-1 (immediately before the structural
            // cell terminator) survives replacement of this cell, other cells and
            // later body formulas. Its absolute coordinate moves with the table,
            // which is expected; AcquireCell resolves the current physical cell
            // from that coordinate plus the captured row/column.
            cellTerminator = document.Range(cellRange.End - 1, cellRange.End - 1);
            var bookmarkName = "VTFCC_" + Guid.Parse(targetFormulaId).ToString("N");
            bookmarks = document.Bookmarks;
            if (bookmarks.Exists(bookmarkName))
                throw new InvalidDataException(
                    $"The user-cell conversion identity {bookmarkName} already exists.");
            cellBookmark = bookmarks.Add(bookmarkName, cellTerminator);
            verifiedBookmarkRange = cellBookmark.Range;
            if (verifiedBookmarkRange.StoryType != cellRange.StoryType
                || verifiedBookmarkRange.Start != cellRange.End - 1
                || verifiedBookmarkRange.End != verifiedBookmarkRange.Start)
                throw new InvalidDataException(
                    "Word did not retain the collapsed source-cell structural identity.");
            var result = new CellConversionOwner(
                bookmarkName,
                rowIndex,
                columnIndex,
                targetFormulaId,
                ordinal,
                formulaCount);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-user-cell-owner-captured formulaId={targetFormulaId} row={rowIndex} column={columnIndex} ordinal={ordinal} bookmark={bookmarkName} cell={cellRange.Start}:{cellRange.End} source={source.Start}:{source.End}");
            return result;
        }
        finally
        {
            Release(verifiedBookmarkRange);
            Release(cellBookmark);
            Release(bookmarks);
            Release(cellTerminator);
            Release(cellRange);
            Release(cell);
            Release(managed);
            Release(source);
            Release(shape);
        }
    }

    private static void VerifyConvertedUserCell(
        Document document,
        string targetMode,
        CellConversionOwner owner,
        WordOmmlNativeSource.LiveInsertionOwners liveOmml)
    {
        Cell? cell = null;
        Range? cellRange = null;
        Range? targetRange = null;
        InlineShape? shape = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        InlineShapes? shapes = null;
        try
        {
            cell = owner.AcquireCell(document);
            cellRange = cell.Range;
            if (targetMode == FormulaOleContract.WordOmmlMode)
                targetRange = TryResolveManagedOmmlByDurableIdentity(
                        document,
                        owner.FormulaId)
                    ?? liveOmml.TryReadCurrentOwner(owner.FormulaId)
                    ?? TryResolveConvertedOmmlByCellOrdinal(document, owner);
            else if (targetMode == FormulaOleContract.NativeOleMode)
            {
                shape = FindByFormulaId(document, owner.FormulaId, allowGlobalFallback: false);
                if (shape is not null) targetRange = shape.Range;
            }
            else
            {
                bookmarks = document.Bookmarks;
                if (owner.MathTypeBookmark is not null && bookmarks.Exists(owner.MathTypeBookmark))
                {
                    bookmark = bookmarks[owner.MathTypeBookmark];
                    targetRange = bookmark.Range;
                    shapes = targetRange.InlineShapes;
                    if (shapes.Count != 1)
                        throw new InvalidDataException("A converted user-cell MathType locator does not own one formula.");
                    shape = shapes[1];
                    Release(targetRange);
                    targetRange = shape.Range;
                }
            }
            if (targetRange is null || targetRange.StoryType != cellRange.StoryType
                || !WordFormulaHost.ContainsPhysicalRange(cellRange.Start, cellRange.End - 1,
                    targetRange.Start, targetRange.End))
                throw new InvalidDataException($"Converted formula {owner.FormulaId} escaped its original user table cell.");
        }
        finally
        {
            Release(shapes);
            Release(bookmark);
            Release(bookmarks);
            Release(shape);
            Release(targetRange);
            Release(cellRange);
            Release(cell);
        }
    }
}
