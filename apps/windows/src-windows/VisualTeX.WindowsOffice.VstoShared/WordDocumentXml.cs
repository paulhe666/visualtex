using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordDocumentXml
{
    internal static string Read(Document document)
    {
        Range? content = null;
        Range? export = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        try
        {
            content = document.Content;
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
                if (exportedMaths == mathCount && exportedOles == oleCount) return xml;
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
