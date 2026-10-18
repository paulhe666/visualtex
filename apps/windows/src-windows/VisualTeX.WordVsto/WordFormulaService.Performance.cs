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
                // ActiveDocument is usually the same RCW the running batch holds
                // for its target. FinalReleaseComObject would invalidate that
                // shared reference mid-batch, so release only this reference.
                if (document is not null)
                {
                    try { Marshal.ReleaseComObject(document); }
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

    private const int MinimizeWindowFormulaThreshold = 10;

    /// <summary>
    /// Raised with true when a batch minimizes the document window and false
    /// when it restores it, so the progress window can stay on screen.
    /// </summary>
    internal Action<bool>? OperationWindowMinimized { get; set; }

    /// <summary>
    /// Minimizes the document window for a batch of host writes and restores
    /// its previous state (normal or maximized) afterwards.
    /// </summary>
    private sealed class MinimizedWindowScope : IDisposable
    {
        private Window? _window;
        private readonly Document _document;
        private readonly WdWindowState _previous;
        private readonly Action<bool>? _notify;

        private MinimizedWindowScope(Document document, Action<bool>? notify)
        {
            _document = document;
            _notify = notify;
            try
            {
                _window = document.ActiveWindow;
                _previous = _window.WindowState;
                if (_previous == WdWindowState.wdWindowStateMinimize || !_window.Visible)
                {
                    Marshal.ReleaseComObject(_window);
                    _window = null;
                    return;
                }
                _notify?.Invoke(true);
                _window.WindowState = WdWindowState.wdWindowStateMinimize;
                // With other documents open, Word activates one of them when this
                // window minimizes and then keeps laying out this (background)
                // document after every write. Re-activating it keeps the window
                // minimized and the layout suspended.
                var reactivated = false;
                if (!IsActiveDocument(document))
                {
                    document.Activate();
                    reactivated = true;
                }
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-window-minimized previous={_previous} reactivated={reactivated} state={_window.WindowState}");
            }
            catch
            {
                if (_window is not null) Marshal.ReleaseComObject(_window);
                _window = null;
                _notify?.Invoke(false);
            }
        }

        internal static MinimizedWindowScope Enter(Document document, Action<bool>? notify) =>
            new(document, notify);

        // Not released: when the target is active, ActiveDocument is the same
        // runtime-callable wrapper the conversion is still using.
        private static bool IsActiveDocument(Document document)
        {
            try
            {
                return string.Equals(
                    document.Application.ActiveDocument.FullName,
                    document.FullName,
                    StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        public void Dispose()
        {
            var window = _window;
            _window = null;
            if (window is null) return;
            // Reverse of Enter: the progress window gives up topmost and returns
            // to its owner before Word restores and activates the document.
            _notify?.Invoke(false);
            try { window.WindowState = _previous; } catch { }
            try { _document.Activate(); } catch { }
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-window-restored state={_previous} active={IsActiveDocument(_document)}");
            Marshal.ReleaseComObject(window);
        }
    }
}
