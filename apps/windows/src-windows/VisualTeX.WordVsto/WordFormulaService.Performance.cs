using System;
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

        private NativeOmmlScreenUpdatingScope(Microsoft.Office.Interop.Word.Application application)
        {
            _application = application;
            try
            {
                _previous = application.ScreenUpdating;
                application.ScreenUpdating = false;
                _active = true;
            }
            catch
            {
                _application = null;
            }
        }

        internal static NativeOmmlScreenUpdatingScope Suspend(Microsoft.Office.Interop.Word.Application application) =>
            new(application);

        public void Dispose()
        {
            var application = _application;
            _application = null;
            if (!_active || application is null) return;
            try { application.ScreenUpdating = _previous; }
            catch { }
        }
    }
}
