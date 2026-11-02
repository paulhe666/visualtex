using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

// WPS writes fields differently from Word. Fields.Add and Field.Update give every field
// a separator and a result, including MACROBUTTON, GOTOBUTTON and hidden (\h) SEQ fields
// that Word keeps without one; WPS then displays the empty result instead of the field's
// display text. In WPS such field trees are inserted as Word's exact field XML (the form
// native MathType numbers and references also use).
internal static class WpsFieldXml
{
    private static readonly XNamespace PackageNamespace =
        "http://schemas.microsoft.com/office/2006/xmlPackage";
    private static readonly XNamespace RelationshipNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static XElement FieldChar(string type) =>
        new(
            W + "r",
            new XElement(
                W + "fldChar",
                new XAttribute(W + "fldCharType", type)));

    internal static XElement Instruction(string text) =>
        new(
            W + "r",
            new XElement(
                W + "instrText",
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                text));

    internal static XElement Result(string text) =>
        new(
            W + "r",
            new XElement(W + "rPr", new XElement(W + "noProof")),
            new XElement(
                W + "t",
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                text));

    // begin, instruction, [separate, result,] end
    internal static IEnumerable<XElement> ComplexField(string instruction, string? result)
    {
        yield return FieldChar("begin");
        yield return Instruction(instruction.StartsWith(" ", StringComparison.Ordinal) ? instruction : " " + instruction + " ");
        if (result is not null)
        {
            yield return FieldChar("separate");
            yield return Result(result);
        }
        yield return FieldChar("end");
    }

    // Inserts runs at a collapsed position through a one-paragraph Flat OPC package and
    // removes the paragraph mark WPS (and sometimes Word) appends after the inserted runs.
    internal static void InsertRuns(
        Document document,
        int position,
        IReadOnlyList<XElement> runs)
    {
        Paragraphs? paragraphs = null;
        Range? insertion = null;
        Range? probe = null;
        Paragraphs? probeParagraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? mark = null;
        try
        {
            paragraphs = document.Paragraphs;
            var before = paragraphs.Count;
            insertion = document.Range(position, position);
            // The import blanks the OLE formulas already in the document until they are laid out again (WPS).
            WpsOleDisplayRefresh.Mark(document);
            insertion.InsertXML(BuildRunsFlatOpc(runs));
            var after = paragraphs.Count;
            if (after == before) return;
            if (after != before + 1)
                throw new InvalidOperationException(
                    $"WPS changed the paragraph count unexpectedly while inserting fields: before={before}, after={after}.");
            probe = document.Range(position, position);
            probeParagraphs = probe.Paragraphs;
            paragraph = probeParagraphs[1];
            paragraphRange = paragraph.Range;
            mark = document.Range(paragraphRange.End - 1, paragraphRange.End);
            if (!string.Equals(mark.Text, "\r", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "WPS split the paragraph at a boundary VisualTeX could not identify safely.");
            mark.Delete();
        }
        finally
        {
            Release(mark);
            Release(paragraphRange);
            Release(paragraph);
            Release(probeParagraphs);
            Release(probe);
            Release(insertion);
            Release(paragraphs);
        }
    }

    // The field whose begin mark sits at fieldStart (in that position's paragraph).
    internal static Field? FindFieldStartingAt(
        Document document,
        int fieldStart,
        Func<Field, bool> predicate)
    {
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        try
        {
            probe = document.Range(fieldStart, fieldStart);
            paragraphs = probe.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            fields = paragraphRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                var field = fields[index];
                Range? code = null;
                var keep = false;
                try
                {
                    code = field.Code;
                    keep = code.Start - 1 == fieldStart && predicate(field);
                }
                finally
                {
                    Release(code);
                    if (!keep) Release(field);
                }
                if (keep) return field;
            }
            return null;
        }
        finally
        {
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    internal static string BuildRunsFlatOpc(IEnumerable<XElement> runs)
    {
        var rootRelationships = new XElement(
            RelationshipNamespace + "Relationships",
            new XElement(
                RelationshipNamespace + "Relationship",
                new XAttribute("Id", "rId1"),
                new XAttribute(
                    "Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
                new XAttribute("Target", "word/document.xml")));
        // WPS's own export of a partial-paragraph range marks its paragraph with
        // doNotImportParagraphEnd: InsertXML then inserts only the runs, without a new
        // paragraph mark and without applying the package paragraph's properties to the
        // host paragraph (see WpsXmlInsertion).
        var paragraph = new XElement(
            W + "p",
            new XElement(W + "pPr", new XElement(W + "doNotImportParagraphEnd")),
            runs);
        var wordDocument = new XElement(
            W + "document",
            new XAttribute(XNamespace.Xmlns + "w", W),
            new XElement(W + "body", paragraph, new XElement(W + "sectPr")));
        XElement Part(string name, string contentType, XElement xml) =>
            new(
                PackageNamespace + "part",
                new XAttribute(PackageNamespace + "name", name),
                new XAttribute(PackageNamespace + "contentType", contentType),
                new XElement(PackageNamespace + "xmlData", xml));
        return new XDocument(
            new XElement(
                PackageNamespace + "package",
                new XAttribute(XNamespace.Xmlns + "pkg", PackageNamespace),
                Part("/_rels/.rels",
                    "application/vnd.openxmlformats-package.relationships+xml",
                    rootRelationships),
                Part("/word/document.xml",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                    wordDocument)))
            .ToString(SaveOptions.DisableFormatting);
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
