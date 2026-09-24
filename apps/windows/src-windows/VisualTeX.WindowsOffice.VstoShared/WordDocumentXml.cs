using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordDocumentXml
{
    internal static bool CanProveNoMathTypePlaceRefFields(string wordOpenXml)
    {
        if (string.IsNullOrWhiteSpace(wordOpenXml)) return false;
        try
        {
            var package = XDocument.Parse(wordOpenXml);
            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var body = package.Descendants(word + "body").SingleOrDefault();
            if (body is null) return false;
            // Word may split one field instruction across several instrText runs.
            // Concatenate without separators so MTPlace|Ref remains detectable.
            // A false positive merely triggers the exact COM fallback; only full
            // absence is accepted as proof that no MathType number field exists.
            var instructions = string.Concat(
                body.Descendants(word + "instrText").Select(node => node.Value));
            return instructions.IndexOf("MTPlaceRef", StringComparison.OrdinalIgnoreCase) < 0;
        }
        catch
        {
            return false;
        }
    }

    internal static bool CanProveNoVisualTeXEquationNumberFields(string wordOpenXml)
    {
        if (string.IsNullOrWhiteSpace(wordOpenXml)) return false;
        try
        {
            var package = XDocument.Parse(wordOpenXml);
            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var body = package.Descendants(word + "body").SingleOrDefault();
            if (body is null) return false;
            // VisualTeX-owned OMML/OLE numbering always carries the fixed,
            // locale-neutral SEQ VisualTeXEquation field. MathType's native
            // MTPlaceRef/MTEqn/MTChap/MTSec fields are deliberately excluded.
            // Concatenate instruction runs because Word may split a field code
            // across several w:instrText nodes. Any ambiguity falls back to the
            // conservative path that still runs the VisualTeX updater.
            var instructions = string.Concat(
                body.Descendants(word + "instrText").Select(node => node.Value));
            return instructions.IndexOf(
                "SEQ VisualTeXEquation",
                StringComparison.OrdinalIgnoreCase) < 0;
        }
        catch
        {
            return false;
        }
    }

    internal static bool IsUnexpectedlyEmptyBody(XElement body, string liveText)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        if (body.Descendants(w + "t").Any(text => !string.IsNullOrWhiteSpace(text.Value))) return false;
        return liveText.Any(c => !char.IsWhiteSpace(c) && !char.IsControl(c)
            && c != '\u200B' && c != '\u200C' && c != '\u2060');
    }

    internal static string Read(Document document, Range? ownedScope = null)
    {
        Range? content = null;
        Range? export = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        try
        {
            content = ownedScope?.Duplicate ?? document.Content;
            if (content.StoryType != WdStoryType.wdMainTextStory)
                throw new InvalidDataException("Recovery XML must belong to the main Word story.");
            var start = content.Start;
            var end = content.End;
            maths = content.OMaths;
            shapes = content.InlineShapes;
            var mathCount = maths.Count;
            var shapeCount = shapes.Count;
            var oleCount = 0;
            for (var index = 1; index <= shapeCount; index++)
            {
                InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    if (shape.Type is WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                        or WdInlineShapeType.wdInlineShapeLinkedOLEObject) oleCount++;
                }
                finally { if (shape is not null) Marshal.ReleaseComObject(shape); }
            }
            // The first export immediately after an OLE/OMML replacement can
            // return Word's empty scratch body, even from a newly bound Range.
            // Read again only if the actual COM inventory disproves that export;
            // no retries of document mutations or relaxed identity checks occur.
            for (var attempt = 0; attempt < 2; attempt++)
            {
                export = document.Range(start, end);
                var xml = export.WordOpenXML;
                if (export.Start != start || export.End != end || content.End != end
                    || maths.Count != mathCount || shapes.Count != shapeCount)
                    throw new InvalidDataException("The document changed while its Word XML was being captured.");
                var package = XDocument.Parse(xml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
                XNamespace o = "urn:schemas-microsoft-com:office:office";
                var body = package.Descendants(w + "body").Single();
                var exportedMaths = body.Descendants(m + "oMath").Count();
                var exportedOles = body.Descendants(o + "OLEObject").Count();
                // Before redraw, an all-LaTeX document contains zero OMath/OLE
                // objects. An empty scratch export also contains zero of both,
                // so inventory equality alone previously accepted a blank undo
                // checkpoint for a nonempty source document.
                var missingPlainText = mathCount == 0 && shapeCount == 0
                    && IsUnexpectedlyEmptyBody(body, content.Text ?? string.Empty);
                if (exportedMaths == mathCount && exportedOles == oleCount && !missingPlainText) return xml;
                WordDoubleClickHook.TraceMessage($"word-xml-incomplete-export document={document.FullName} attempt={attempt + 1} maths={exportedMaths}/{mathCount} oles={exportedOles}/{oleCount}");
                Marshal.ReleaseComObject(export); export = null;
            }
            throw new InvalidDataException("Word could not export XML matching the actual formula and OLE inventory.");
        }
        finally
        {
            if (export is not null) Marshal.ReleaseComObject(export);
            if (shapes is not null) Marshal.ReleaseComObject(shapes);
            if (maths is not null) Marshal.ReleaseComObject(maths);
            if (content is not null) Marshal.ReleaseComObject(content);
        }
    }
}
