using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

public sealed partial class ThisAddIn
{
    private bool _ommlDisplayMigrationAttached;
    private bool _ommlDisplayMigrationRunning;

    private void AttachOmmlDisplayMigration()
    {
        if (_ommlDisplayMigrationAttached || _application is null) return;
        _ommlDisplayMigrationAttached = true;
        _application.DocumentChange += Application_DocumentChangeForOmmlDisplayMigration;
        TryMigrateActiveDocumentOmmlDisplayTypes();
    }

    private void DetachOmmlDisplayMigration()
    {
        if (!_ommlDisplayMigrationAttached) return;
        _ommlDisplayMigrationAttached = false;
        try
        {
            if (_application is not null)
                _application.DocumentChange -= Application_DocumentChangeForOmmlDisplayMigration;
        }
        catch
        {
            // Word may already be shutting down its event connection point.
        }
    }

    private void Application_DocumentChangeForOmmlDisplayMigration()
    {
        TryMigrateActiveDocumentOmmlDisplayTypes();
    }

    private void TryMigrateActiveDocumentOmmlDisplayTypes()
    {
        if (_ommlDisplayMigrationRunning) return;
        Document? document = null;
        try
        {
            document = _application?.ActiveDocument;
            if (document is null || document.ReadOnly) return;
            _ommlDisplayMigrationRunning = true;
            _ = MigrateManagedOmmlDisplayTypes(document);
        }
        catch (COMException)
        {
            // A transient protected/read-only/document-switch state must not block
            // Word startup. The next DocumentChange or explicit edit retries.
        }
        finally
        {
            _ommlDisplayMigrationRunning = false;
            ReleaseMigrationComObject(document);
        }
    }

    internal static int MigrateManagedOmmlDisplayTypes(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        OMaths? maths = null;
        var migrated = 0;
        // Include legacy VisualTeX OLE numbering hosts in the same one-pass
        // document upgrade. Previously this flag was raised only while walking
        // managed OMaths, so a document containing only an old numbered OLE table
        // could open without ever entering structural reconciliation.
        var requiresNumberingReconcile =
            WordEquationNumbering.NeedsLegacyManagedNumberingMigration(document);
        try
        {
            maths = document.OMaths;
            // Work backwards because changing OMath.Type can rematerialize ranges.
            for (var index = maths.Count; index >= 1; index--)
            {
                OMath? math = null;
                Range? range = null;
                Bookmark? bookmark = null;
                try
                {
                    math = maths[index];
                    range = math.Range.Duplicate;
                    bookmark = WordOmmlFormulaStore.FindAtRange(document, range);
                    if (bookmark is null) continue;
                    var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                    if (metadata is null) continue;
                    var block = string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.OrdinalIgnoreCase);
                    if (block && metadata.Numbered)
                    {
                        // Numbered OMML is a document structure, not merely an
                        // OMath.Type flag. Old 1x3/2x3 tables, inline-tab hosts and
                        // m:eqArr/# wrappers must enter one document-wide numbering
                        // reconciliation so FormulaId, SEQ, REF and cross-references
                        // are migrated together. Never force these formulas inline.
                        requiresNumberingReconcile |=
                            math.Type != WdOMathType.wdOMathDisplay
                            || !WordEquationNumbering
                                .HasStructurallyReusableNumberedNativeOmmlDisplayHost(
                                    document,
                                    range,
                                    metadata.FormulaId);
                        continue;
                    }

                    var target = block
                        ? WdOMathType.wdOMathDisplay
                        : WdOMathType.wdOMathInline;
                    if (math.Type == target) continue;

                    math.Type = target;
                    ReleaseMigrationComObject(range);
                    range = math.Range.Duplicate;
                    WordOmmlNativeSource.StampFingerprint(metadata, range);
                    ReleaseMigrationComObject(bookmark);
                    bookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        range,
                        metadata,
                        replaceExisting: true);
                    WordOmmlFormulaStore.Save(document, metadata);
                    migrated += 1;
                }
                finally
                {
                    ReleaseMigrationComObject(bookmark);
                    ReleaseMigrationComObject(range);
                    ReleaseMigrationComObject(math);
                }
            }
            if (requiresNumberingReconcile)
            {
                // UpdateEquationNumbers now treats only the final table-free OLE
                // tab paragraph and the pure m:oMathPara + external REF Shape host
                // as healthy. Every legacy table/inline/eqArr structure therefore
                // falls through to structural migration in this single pass.
                migrated += WordEquationNumbering.UpdateEquationNumbers(document);
            }
            if (migrated > 0)
            {
                try { document.Repaginate(); } catch { }
            }
            return migrated;
        }
        finally
        {
            ReleaseMigrationComObject(maths);
        }
    }

    private static void ReleaseMigrationComObject(object? value)
    {
        if (value is null) return;
        try
        {
            if (Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
        catch
        {
            // Best-effort release during Word event handling.
        }
    }
}
