using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using WordApplication = Microsoft.Office.Interop.Word.Application;

namespace VisualTeX.WordVsto;

/// <summary>
/// One structural mutation = one Word Custom UndoRecord.
///
/// The rebuilt host core never borrows/nests an older transaction. A caller that
/// already owns an undo record has not been migrated yet and must not mix the two
/// mutation models.
/// </summary>
internal static class WordFormulaMutationTransaction
{
    [ThreadStatic]
    private static int _activeDepth;

    internal static bool IsActiveOnCurrentThread =>
        _activeDepth > 0;

    internal static T Execute<T>(
        WordApplication application,
        Document document,
        string name,
        Func<T> mutationAndValidation,
        WordDocumentEditSnapshot? rollbackSnapshot = null,
        Action? verifyRollback = null)
    {
        if (application is null) throw new ArgumentNullException(nameof(application));
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (mutationAndValidation is null)
            throw new ArgumentNullException(nameof(mutationAndValidation));

        UndoRecord? undo = null;
        var started = false;
        var ended = false;
        try
        {
            undo = application.UndoRecord;
            if (undo.IsRecordingCustomRecord || undo.CustomRecordLevel > 0)
                throw new InvalidOperationException(
                    "The rebuilt OMML/VisualTeX mutation core cannot run inside a legacy Word undo transaction.");

            undo.StartCustomRecord(name);
            started = true;
            _activeDepth++;
            TraceBoundary(application, "transaction-start");

            try
            {
                var result = mutationAndValidation();
                try { WordEquationNumbering.RetireLegacyEquationNumberFormat(document); }
                catch { /* Never fail a completed edit over an obsolete variable. */ }
                TraceBoundary(application, "transaction-before-end");
                undo.EndCustomRecord();
                ended = true;
                return result;
            }
            catch (Exception primary)
            {
                if (started && !ended)
                {
                    try
                    {
                        undo.EndCustomRecord();
                        ended = true;
                    }
                    catch (Exception endError)
                    {
                        throw new AggregateException(
                            OfficePluginLanguage.Text(
                                "公式修改失败，而且 Word 无法关闭它的撤销记录。",
                                "The formula change failed, and Word could not close its undo record."),
                            primary,
                            endError);
                    }
                }

                if (rollbackSnapshot is not null)
                {
                    try
                    {
                        rollbackSnapshot.Restore(document);
                    }
                    catch (Exception recoveryError)
                    {
                        throw new AggregateException(
                            OfficePluginLanguage.Text(
                                "公式修改失败，而且无法用回滚快照完整恢复本批修改前的文档。",
                                "The formula change failed, and the rollback snapshot could not restore the complete batch."),
                            primary,
                            recoveryError);
                    }
                }
                else if (!TryUndoExactlyOne(document, out var undoError))
                {
                    throw new AggregateException(
                        OfficePluginLanguage.Text(
                            "公式修改失败，而且无法撤销它的 Word 撤销记录（文档可能处于部分修改状态）。",
                            "The formula change failed, and its Word undo record could not be rolled back (the document may be partly changed)."),
                        primary,
                        undoError
                            ?? new InvalidOperationException(OfficePluginLanguage.Text(
                                "Word 撤销失败的公式修改时返回了失败。",
                                "Word reported failure while undoing the failed formula change.")));
                }

                if (verifyRollback is not null)
                {
                    try { verifyRollback(); }
                    catch (Exception verificationError)
                    {
                        throw new AggregateException(
                            OfficePluginLanguage.Text(
                                "撤销失败的公式修改后，原公式没有被完全恢复。",
                                "Undoing the failed formula change did not fully restore the original formula."),
                            primary, verificationError);
                    }
                }
                throw;
            }
        }
        finally
        {
            if (started && _activeDepth > 0)
                _activeDepth--;

            if (started && !ended)
            {
                try { undo?.EndCustomRecord(); } catch { }
            }
            Release(undo);
        }
    }

    // Opt-in diagnostics only. Prove whether a Word-native operation ended the
    // application-wide custom record; never restart or nest it to conceal a split.
    internal static void TraceBoundary(WordApplication application, string stage)
    {
        if (!string.Equals(Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF"),
                "1", StringComparison.Ordinal)) return;
        UndoRecord? record = null;
        try
        {
            record = application.UndoRecord;
            WordDoubleClickHook.TraceMessage($"host-undo-boundary stage={stage} recording={record.IsRecordingCustomRecord} level={record.CustomRecordLevel}");
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage($"host-undo-boundary-unavailable stage={stage} error={error.Message}");
        }
        finally { Release(record); }
    }

    private static bool TryUndoExactlyOne(
        Document document,
        out Exception? error)
    {
        error = null;
        object times = 1;
        try
        {
            return document.Undo(ref times);
        }
        catch (Exception undoError)
        {
            error = undoError;
            return false;
        }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
