using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    /// <summary>
    /// Keeps the complete OMML replacement/numbering transaction off screen.
    /// Word otherwise repaints while its hidden conversion source and the target
    /// document exchange focus, producing a visible top-of-document flash and
    /// forcing repeated layout of every unrelated formula in large documents.
    /// </summary>
    private sealed class NativeOmmlScreenUpdatingScope : IDisposable
    {
        private Microsoft.Office.Interop.Word.Application? _application;
        private readonly bool _previous;
        private readonly bool _active;
        private bool _documentChangeAttached;

        private NativeOmmlScreenUpdatingScope(Microsoft.Office.Interop.Word.Application application)
        {
            _application = application;
            try
            {
                _previous = application.ScreenUpdating;
                application.ScreenUpdating = false;
                // Word 2019/2021 can force ScreenUpdating back to true when a
                // hidden Documents.Open becomes ActiveDocument. Keep the
                // invariant for the entire source lifetime instead of assuming
                // the property remains sticky after one assignment.
                application.DocumentChange += KeepScreenUpdatingSuspended;
                _documentChangeAttached = true;
                _active = true;
            }
            catch
            {
                if (_documentChangeAttached)
                {
                    try { application.DocumentChange -= KeepScreenUpdatingSuspended; }
                    catch { }
                    _documentChangeAttached = false;
                }
                _application = null;
            }
        }

        private void KeepScreenUpdatingSuspended()
        {
            var application = _application;
            if (application is null) return;
            try
            {
                if (application.ScreenUpdating)
                {
                    application.ScreenUpdating = false;
                    WordDoubleClickHook.TraceMessage(
                        $"omml-screen-update-guard-reasserted activeDocument={TryReadActiveDocumentName(application)} screenUpdating={application.ScreenUpdating}");
                }
            }
            catch
            {
                // Word may be between document/window states. The target rebind
                // performed by the owning transaction will retry the invariant
                // before this scope is released.
            }
        }

        private static string TryReadActiveDocumentName(
            Microsoft.Office.Interop.Word.Application application)
        {
            Document? document = null;
            try
            {
                document = application.ActiveDocument;
                return document?.Name ?? string.Empty;
            }
            catch { return "<unavailable>"; }
            finally
            {
                if (document is not null)
                {
                    try { Marshal.FinalReleaseComObject(document); }
                    catch { }
                }
            }
        }

        internal static NativeOmmlScreenUpdatingScope Suspend(Microsoft.Office.Interop.Word.Application application) =>
            new(application);

        public void Dispose()
        {
            var application = _application;
            _application = null;
            if (!_active || application is null) return;
            if (_documentChangeAttached)
            {
                try { application.DocumentChange -= KeepScreenUpdatingSuspended; }
                catch { }
                _documentChangeAttached = false;
            }
            try { application.ScreenUpdating = _previous; }
            catch { }
        }
    }
}
