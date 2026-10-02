using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Prepare an immutable native MathType chapter/section paragraph before the
/// target document's undo transaction starts. Hidden-document creation and
/// closure must not split that transaction into unrelated native undo entries.
/// </summary>
internal sealed class MathTypeSectionStaging : IDisposable
{
    private Document? _staging;
    internal Range Paragraph { get; private set; } = null!;

    internal MathTypeSectionStaging(Document target, string flatOpc)
    {
        Microsoft.Office.Interop.Word.Application? app = null;
        Documents? documents = null;
        Range? insertion = null;
        Range? content = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        Range? full = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        InlineShapes? shapes = null;
        try
        {
            app = target.Application;
            documents = app.Documents;
            _staging = documents.Add(Visible: false);
            content = _staging.Content;
            insertion = content.Duplicate;
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            insertion.InsertXML(flatOpc);
            Release(content); content = _staging.Content;
            field = MathTypeFieldLocator.AtCodeStart(content, MathTypeFieldLocator.Section, content.Start + 1)
                ?? throw new InvalidDataException("The prepared MathType section has no exact native macro field.");
            code = field.Code;
            result = field.Result;
            full = _staging.Range(code.Start - 1, result.End + 1);
            paragraphs = full.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException("The prepared MathType section crosses a paragraph boundary.");
            paragraph = paragraphs[1];
            Paragraph = paragraph.Range.Duplicate;
            shapes = Paragraph.InlineShapes;
            if (shapes.Count != 0)
                throw new InvalidDataException("The prepared MathType section unexpectedly contains an OLE.");
        }
        catch { Dispose(); throw; }
        finally
        {
            Release(shapes); Release(paragraph); Release(paragraphs); Release(full);
            Release(result); Release(code); Release(field); Release(content); Release(insertion);
            Release(documents); Release(app);
            target.Activate();
        }
    }

    public void Dispose()
    {
        Release(Paragraph); Paragraph = null!;
        var staging = _staging;
        _staging = null;
        try { staging?.Close(WdSaveOptions.wdDoNotSaveChanges); }
        finally { Release(staging); }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
