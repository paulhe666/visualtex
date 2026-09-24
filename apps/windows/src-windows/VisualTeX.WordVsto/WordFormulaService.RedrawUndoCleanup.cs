using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    private sealed class OmmlRedrawUndoCleanupTicket
    {
        internal string DocumentId { get; set; } = string.Empty;
        internal string[] ProbeFormulaIds { get; set; } = Array.Empty<string>();
    }

    private readonly object _ommlRedrawUndoCleanupGate = new();
    private readonly List<OmmlRedrawUndoCleanupTicket> _ommlRedrawUndoCleanupTickets = new();

    internal bool HasPendingOmmlRedrawUndoCleanup
    {
        get
        {
            lock (_ommlRedrawUndoCleanupGate)
                return _ommlRedrawUndoCleanupTickets.Count > 0;
        }
    }

    internal void TrackCompletedOmmlRedrawForUndo(
        string documentId,
        IReadOnlyList<string> formulaIds)
    {
        if (string.IsNullOrWhiteSpace(documentId)
            || formulaIds is null
            || formulaIds.Count == 0)
            return;

        var validIds = formulaIds
            .Where(id => Guid.TryParse(id, out _))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (validIds.Length == 0) return;

        var probeIndices = new HashSet<int>
        {
            0,
            validIds.Length / 2,
            validIds.Length - 1,
        };
        var ticket = new OmmlRedrawUndoCleanupTicket
        {
            DocumentId = documentId,
            ProbeFormulaIds = probeIndices
                .OrderBy(index => index)
                .Select(index => validIds[index])
                .ToArray(),
        };

        lock (_ommlRedrawUndoCleanupGate)
        {
            _ommlRedrawUndoCleanupTickets.Add(ticket);
            // Keep enough history for ordinary repeated Ctrl+Z while bounding
            // process-lifetime memory. Cleanup itself is document-derived and can
            // remove orphans from older batches once any tracked undo is observed.
            if (_ommlRedrawUndoCleanupTickets.Count > 64)
                _ommlRedrawUndoCleanupTickets.RemoveRange(
                    0,
                    _ommlRedrawUndoCleanupTickets.Count - 64);
        }
    }

    internal int TryCleanupPendingOmmlRedrawUndoMetadata(Document document)
    {
        if (document is null) return 0;
        string documentId;
        try { documentId = DocumentIdentity(document); }
        catch { return 0; }

        OmmlRedrawUndoCleanupTicket[] tickets;
        lock (_ommlRedrawUndoCleanupGate)
        {
            tickets = _ommlRedrawUndoCleanupTickets
                .Where(ticket => string.Equals(
                    ticket.DocumentId,
                    documentId,
                    StringComparison.OrdinalIgnoreCase))
                .ToArray();
        }
        if (tickets.Length == 0) return 0;

        Bookmarks? bookmarks = null;
        var undoneTickets = new List<OmmlRedrawUndoCleanupTicket>();
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var ticket in tickets)
            {
                var allMissing = true;
                foreach (var formulaId in ticket.ProbeFormulaIds)
                {
                    if (bookmarks.Exists(WordOmmlFormulaStore.BookmarkName(formulaId)))
                    {
                        allMissing = false;
                        break;
                    }
                }
                if (allMissing) undoneTickets.Add(ticket);
            }
        }
        finally { Release(bookmarks); }

        if (undoneTickets.Count == 0) return 0;

        var removed = WordOmmlFormulaStore.RemoveOrphanedMetadataParts(document);
        lock (_ommlRedrawUndoCleanupGate)
        {
            foreach (var ticket in undoneTickets)
                _ommlRedrawUndoCleanupTickets.Remove(ticket);
        }
        WordDoubleClickHook.TraceMessage(
            $"redraw-omml-undo-orphan-metadata-cleaned tickets={undoneTickets.Count} removed={removed}");
        return removed;
    }
}
