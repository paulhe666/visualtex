using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Range.InsertXML with Word's semantics in WPS (measured on WPS 12.1.0.28488). Word keeps
/// its native call. WPS differs in three ways, handled here:
/// <list type="bullet">
/// <item>Paragraph end. Word drops the package's last paragraph mark when the content lands
/// in a paragraph that has text, and keeps it (with its paragraph properties) when the
/// target paragraph is empty or the range ends with a paragraph mark. WPS always imports
/// that mark, splitting the paragraph, unless the package paragraph carries
/// w:doNotImportParagraphEnd (WPS writes it into its own partial-range WordOpenXML). WPS
/// also cannot replace the story's final mark: it imports into the previous paragraph
/// instead. There the package mark is imported before the final mark and then deleted;
/// merging into the empty final paragraph keeps the imported paragraph's properties.</item>
/// <item>EMBED wrapper. WPS exports an embedded OLE object as an explicit EMBED field
/// (fldChar begin, instrText, separate, &lt;w:object&gt;, end). On import WPS creates its own
/// implicit EMBED field for the &lt;w:object&gt; and keeps the explicit one, so a WPS export
/// re-inserted as is nests the object inside a second EMBED field. Word writes the
/// &lt;w:object&gt; run alone; the explicit wrapper is removed.</item>
/// <item>Field instructions in equations. Word stores the instruction of a field inside an
/// OMath (the STYLEREF/SEQ of a #(number)) as w:instrText in m:r; WPS reads it as math
/// text (SEQ becomes italic variables, "*" becomes U+2217) and the field breaks. Written as
/// normal-text math (m:t with m:nor), which WPS keeps, as the macOS WPS add-in writes it.
/// Equations are spelled with the m: and w: prefixes; WPS does not take a default-namespace
/// spelling for an equation.</item>
/// <item>Replace deletes a non-collapsed range inside a paragraph first and inserts at the
/// collapsed position, inside the caller's undo record.</item>
/// </list>
/// The range is left collapsed at the insertion start, as WPS leaves it.
/// </summary>
internal static class WpsXmlInsertion
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M =
        "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace Package =
        "http://schemas.microsoft.com/office/2006/xmlPackage";

    private enum EmptyParagraph { None, Middle, Final }

    internal static void Insert(Range range, string flatOpc)
    {
        if (!OfficeHostEnvironment.IsWps)
        {
            range.InsertXML(flatOpc);
            return;
        }
        InsertInWps(range, flatOpc, deleteFirst: false);
    }

    internal static void Replace(Range range, string flatOpc)
    {
        if (!OfficeHostEnvironment.IsWps)
        {
            range.InsertXML(flatOpc);
            return;
        }
        InsertInWps(range, flatOpc, deleteFirst: true);
    }

    private static void InsertInWps(Range range, string flatOpc, bool deleteFirst)
    {
        var start = range.Start;
        if (range.End > range.Start && EndsWithParagraphMark(range))
        {
            if (range.End < StoryEnd(range))
            {
                range.InsertXML(PrepareForWps(flatOpc, importParagraphEnd: true).Xml);
                range.SetRange(start, start);
                return;
            }
            // The story's final mark stays; clear the content before it.
            range.SetRange(start, range.End - 1);
            if (range.End > range.Start) range.Delete();
            range.SetRange(start, start);
        }
        if (range.Start == range.End)
        {
            var empty = EmptyParagraphAt(range);
            WordDoubleClickHook.TraceMessage($"wps-insert-xml collapsed at {start} empty-paragraph={empty}");
            if (empty == EmptyParagraph.Middle)
            {
                range.SetRange(start, start + 1);
                range.InsertXML(PrepareForWps(flatOpc, importParagraphEnd: true).Xml);
                range.SetRange(start, start);
                return;
            }
            if (empty == EmptyParagraph.Final)
            {
                InsertIntoFinalEmptyParagraph(range, flatOpc);
                range.SetRange(start, start);
                return;
            }
        }
        var prepared = PrepareForWps(flatOpc, importParagraphEnd: false).Xml;
        if (deleteFirst && range.End > range.Start)
        {
            range.Delete();
            range.SetRange(start, start);
        }
        range.InsertXML(prepared);
        range.SetRange(start, start);
    }

    private static void InsertIntoFinalEmptyParagraph(Range range, string flatOpc)
    {
        var prepared = PrepareForWps(flatOpc, importParagraphEnd: true);
        var before = StoryEnd(range);
        range.InsertXML(prepared.Xml);
        if (!prepared.EndsWithParagraph) return;
        var after = StoryEnd(range);
        WordDoubleClickHook.TraceMessage($"wps-insert-xml final-paragraph story {before}->{after}");
        if (after <= before) return;
        Range? importedMark = null;
        try
        {
            importedMark = range.Duplicate;
            importedMark.SetRange(after - 2, after - 1);
            var mark = IsParagraphMark(importedMark);
            WordDoubleClickHook.TraceMessage(
                $"wps-insert-xml final-paragraph imported mark [{after - 2},{after - 1}) paragraph-mark={mark}");
            if (mark)
                importedMark.Delete();
        }
        finally
        {
            if (importedMark is not null) Marshal.ReleaseComObject(importedMark);
        }
    }

    private static EmptyParagraph EmptyParagraphAt(Range caret)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            // Story ends and cell markers elsewhere are not modelled; those keep the merge.
            if (caret.StoryType != WdStoryType.wdMainTextStory
                || Convert.ToBoolean(caret.get_Information(WdInformation.wdWithInTable)))
                return EmptyParagraph.None;
            paragraphs = caret.Paragraphs;
            if (paragraphs.Count != 1) return EmptyParagraph.None;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.End - paragraphRange.Start != 1
                || paragraphRange.Start != caret.Start
                || !string.Equals(paragraphRange.Text, "\r", StringComparison.Ordinal))
                return EmptyParagraph.None;
            return paragraphRange.End >= StoryEnd(caret)
                ? EmptyParagraph.Final
                : EmptyParagraph.Middle;
        }
        finally
        {
            if (paragraphRange is not null) Marshal.ReleaseComObject(paragraphRange);
            if (paragraph is not null) Marshal.ReleaseComObject(paragraph);
            if (paragraphs is not null) Marshal.ReleaseComObject(paragraphs);
        }
    }

    private static int StoryEnd(Range range) => range.StoryLength;

    internal static bool EndsWithParagraphMark(Range range)
    {
        if (range.End <= range.Start) return false;
        Range? last = null;
        try
        {
            last = range.Duplicate;
            last.SetRange(range.End - 1, range.End);
            return IsParagraphMark(last);
        }
        finally
        {
            if (last is not null) Marshal.ReleaseComObject(last);
        }
    }

    // One character that is a paragraph mark. WPS reads the structure of an equation
    // (the start and end of an equation array or a delimiter) as "\r" as well; only a
    // paragraph mark ends its paragraph.
    private static bool IsParagraphMark(Range character)
    {
        if (!string.Equals(character.Text, "\r", StringComparison.Ordinal)) return false;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = character.Paragraphs;
            if (paragraphs.Count < 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return paragraphRange.End == character.End;
        }
        finally
        {
            if (paragraphRange is not null) Marshal.ReleaseComObject(paragraphRange);
            if (paragraph is not null) Marshal.ReleaseComObject(paragraph);
            if (paragraphs is not null) Marshal.ReleaseComObject(paragraphs);
        }
    }

    internal static (string Xml, bool EndsWithParagraph) PrepareForWps(string flatOpc, bool importParagraphEnd)
    {
        if (string.IsNullOrEmpty(flatOpc)) return (flatOpc, false);
        XDocument document;
        try { document = XDocument.Parse(flatOpc, LoadOptions.PreserveWhitespace); }
        catch { return (flatOpc, false); }
        var changed = RemoveExplicitEmbedFields(document);
        changed |= WriteEquationFieldInstructionsAsNormalText(document);
        var last = LastBodyParagraph(document);
        if (last is not null)
            changed |= SetParagraphEndImport(last, importParagraphEnd);
        if (!changed) return (flatOpc, last is not null);
        var xml = document.Declaration is null
            ? document.ToString(SaveOptions.DisableFormatting)
            : document.Declaration + document.ToString(SaveOptions.DisableFormatting);
        return (xml, last is not null);
    }

    // The main document body's last block, when it is a paragraph.
    private static XElement? LastBodyParagraph(XDocument document)
    {
        var mainPart = document.Root?.Name == Package + "package"
            ? document.Root.Elements(Package + "part").FirstOrDefault(part =>
                string.Equals((string?)part.Attribute(Package + "name"), "/word/document.xml", StringComparison.OrdinalIgnoreCase))
            : document.Root;
        var body = mainPart?.Descendants(W + "body").FirstOrDefault();
        var last = body?.Elements().LastOrDefault(element => element.Name != W + "sectPr");
        return last is not null && last.Name == W + "p" ? last : null;
    }

    private static bool SetParagraphEndImport(XElement paragraph, bool importParagraphEnd)
    {
        var properties = paragraph.Element(W + "pPr");
        var marker = properties?.Element(W + "doNotImportParagraphEnd");
        if (importParagraphEnd)
        {
            if (marker is null) return false;
            marker.Remove();
            return true;
        }
        if (marker is not null) return false;
        if (properties is null)
        {
            properties = new XElement(W + "pPr");
            paragraph.AddFirst(properties);
        }
        // WPS writes the marker as the last paragraph property.
        properties.Add(new XElement(W + "doNotImportParagraphEnd"));
        return true;
    }

    private static bool WriteEquationFieldInstructionsAsNormalText(XDocument document)
    {
        var changed = false;
        foreach (var equation in document.Descendants(M + "oMath")
                     .Where(math => !math.Ancestors(M + "oMath").Any())
                     .ToList())
        {
            changed |= WriteFieldInstructionsAsNormalText(equation);
            changed |= UseMathAndWordPrefixes(equation);
        }
        return changed;
    }

    // Every run holding instruction text (between a field's begin and separate characters)
    // becomes a normal-text run whose switches use "*".
    internal static bool WriteFieldInstructionsAsNormalText(XElement equation)
    {
        var changed = false;
        // One entry per open field: true while its instruction is being read.
        var fields = new Stack<bool>();
        foreach (var run in equation.Descendants(M + "r").ToList())
        {
            var instruction = false;
            foreach (var child in run.Elements().ToList())
            {
                if (child.Name == W + "fldChar")
                {
                    switch ((string?)child.Attribute(W + "fldCharType"))
                    {
                        case "begin":
                            fields.Push(true);
                            break;
                        case "separate":
                            if (fields.Count > 0) { fields.Pop(); fields.Push(false); }
                            break;
                        case "end":
                            if (fields.Count > 0) fields.Pop();
                            break;
                    }
                }
                else if (fields.Count > 0 && fields.Peek()
                         && (child.Name == W + "instrText" || child.Name == M + "t"))
                {
                    instruction = true;
                    var text = child.Value.Replace('\u2217', '*');
                    if (child.Name == M + "t" && text == child.Value) continue;
                    child.ReplaceWith(new XElement(
                        M + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        text));
                    changed = true;
                }
            }
            if (!instruction) continue;
            var properties = run.Element(M + "rPr");
            if (properties is null)
            {
                properties = new XElement(M + "rPr");
                run.AddFirst(properties);
            }
            // m:nor and m:scr/m:sty are alternatives in a math run's properties.
            var styles = properties.Elements()
                .Where(element => element.Name == M + "scr" || element.Name == M + "sty")
                .ToList();
            if (styles.Count > 0)
            {
                styles.ForEach(style => style.Remove());
                changed = true;
            }
            if (properties.Element(M + "nor") is null)
            {
                var literal = properties.Element(M + "lit");
                if (literal is not null) literal.AddAfterSelf(new XElement(M + "nor"));
                else properties.AddFirst(new XElement(M + "nor"));
                changed = true;
            }
        }
        return changed;
    }

    // The equation spelled with the m: and w: prefixes: namespace declarations inside it
    // (an equation built from new elements carries default-namespace ones) are dropped, and
    // m and w are declared on the equation unless an ancestor already does.
    private static bool UseMathAndWordPrefixes(XElement equation)
    {
        var declarations = equation.DescendantsAndSelf()
            .SelectMany(element => element.Attributes())
            .Where(attribute => attribute.IsNamespaceDeclaration)
            .ToList();
        var changed = false;
        foreach (var declaration in declarations)
        {
            var prefix = declaration.Name.Namespace == XNamespace.None ? string.Empty : declaration.Name.LocalName;
            var keep = declaration.Parent == equation
                && (prefix == "m" && declaration.Value == M.NamespaceName
                    || prefix == "w" && declaration.Value == W.NamespaceName);
            if (keep) continue;
            declaration.Remove();
            changed = true;
        }
        foreach (var (prefix, ns) in new[] { ("m", M), ("w", W) })
        {
            if (equation.GetPrefixOfNamespace(ns) == prefix) continue;
            equation.SetAttributeValue(XNamespace.Xmlns + prefix, ns.NamespaceName);
            changed = true;
        }
        return changed;
    }

    private static bool RemoveExplicitEmbedFields(XDocument document)
    {
        var changed = false;
        foreach (var objectRun in document.Descendants(W + "r")
                     .Where(run => run.Element(W + "object") is not null)
                     .ToList())
        {
            var separate = PreviousRun(objectRun);
            var instruction = separate is null ? null : PreviousRun(separate);
            var begin = instruction is null ? null : PreviousRun(instruction);
            var end = NextRun(objectRun);
            if (!IsFieldChar(separate, "separate")
                || !IsEmbedInstruction(instruction)
                || !IsFieldChar(begin, "begin")
                || !IsFieldChar(end, "end"))
                continue;
            begin!.Remove();
            instruction!.Remove();
            separate!.Remove();
            end!.Remove();
            changed = true;
        }
        return changed;
    }

    // Runs only; bookmarks and proofing marks between the field runs are skipped.
    private static XElement? PreviousRun(XElement run)
    {
        for (var node = run.ElementsBeforeSelf().LastOrDefault(); node is not null; node = node.ElementsBeforeSelf().LastOrDefault())
        {
            if (node.Name == W + "r") return node;
            if (node.Name != W + "bookmarkStart" && node.Name != W + "bookmarkEnd" && node.Name != W + "proofErr")
                return null;
        }
        return null;
    }

    private static XElement? NextRun(XElement run)
    {
        foreach (var node in run.ElementsAfterSelf())
        {
            if (node.Name == W + "r") return node;
            if (node.Name != W + "bookmarkStart" && node.Name != W + "bookmarkEnd" && node.Name != W + "proofErr")
                return null;
        }
        return null;
    }

    private static bool IsFieldChar(XElement? run, string type) =>
        run is not null
        && run.Elements().Count(element => element.Name != W + "rPr") == 1
        && string.Equals(
            (string?)run.Element(W + "fldChar")?.Attribute(W + "fldCharType"),
            type,
            StringComparison.Ordinal);

    private static bool IsEmbedInstruction(XElement? run)
    {
        if (run is null) return false;
        var instruction = run.Element(W + "instrText");
        return instruction is not null
            && run.Elements().Count(element => element.Name != W + "rPr") == 1
            && instruction.Value.TrimStart().StartsWith("EMBED ", StringComparison.OrdinalIgnoreCase);
    }
}
