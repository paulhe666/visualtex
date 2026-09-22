using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Reads semantic content from a host that was already resolved by
/// <see cref="WordFormulaHostResolver"/>. This class never repairs identities,
/// rewrites metadata, refreshes numbering, or performs a document-wide scan.
/// </summary>
internal static class WordFormulaHostSemanticReader
{
    private static readonly XNamespace MathNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/math";

    internal static WordFormulaSemanticPayload Read(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (host is null) throw new ArgumentNullException(nameof(host));

        return host.Kind switch
        {
            WordFormulaHostKind.Omml => ReadOmml(document, host),
            WordFormulaHostKind.VisualTeX => ReadVisualTeX(document, host),
            _ => throw new NotSupportedException(
                $"Semantic reading is not owned by the OMML/VisualTeX core for {host.Kind}."),
        };
    }

    private static WordFormulaSemanticPayload ReadOmml(
        Document document,
        WordFormulaHostDescriptor host)
    {
        Range? range = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? exact = null;
        try
        {
            range = CreateRange(document, host.Range);
            maths = range.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    "The resolved OMML host no longer contains exactly one equation.");

            math = maths[1];
            exact = math.Range.Duplicate;
            if (!SameAddress(exact, host.Range))
                throw new InvalidDataException(
                    "The resolved OMML host moved before semantic capture.");

            var display = math.Type == WdOMathType.wdOMathDisplay;
            var hostXml = ReadLocalEquationXml(exact);
            var semanticXml = hostXml;
            if (display
                && WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                    hostXml,
                    formulaId: null))
            {
                // Word's native #(SEQ) is host structure, not formula semantics.
                // Every editor/conversion path receives only the mathematical body.
                semanticXml = WordOmmlConverter
                    .StripManagedVisualTeXNativeEquationNumber(
                        hostXml);
            }

            var mathMl = WordOmmlConverter.TransformOmmlToMathMl(
                semanticXml,
                display);
            var latex = MathMlToLatexConverter.Convert(mathMl);
            return new WordFormulaSemanticPayload
            {
                Latex = latex,
                MathMl = mathMl,
                WordOpenXml = semanticXml,
                Metadata = null,
            };
        }
        finally
        {
            Release(exact);
            Release(math);
            Release(maths);
            Release(range);
        }
    }

    private static WordFormulaSemanticPayload ReadVisualTeX(
        Document document,
        WordFormulaHostDescriptor host)
    {
        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            range = CreateRange(document, host.Range);
            shapes = range.InlineShapes;
            InlineShape? resolved = null;
            for (var index = 1; index <= shapes.Count; index++)
            {
                shape = shapes[index];
                if (!WordFormulaMetadataReader.IsNativeOle(shape))
                {
                    Release(shape);
                    shape = null;
                    continue;
                }

                shapeRange = shape.Range.Duplicate;
                if (!SameAddress(shapeRange, host.Range))
                {
                    Release(shapeRange);
                    shapeRange = null;
                    Release(shape);
                    shape = null;
                    continue;
                }

                if (resolved is not null)
                {
                    Release(resolved);
                    throw new InvalidDataException(
                        "More than one VisualTeX host occupies the resolved range.");
                }
                resolved = shape;
                shape = null;
                Release(shapeRange);
                shapeRange = null;
            }

            if (resolved is null)
                throw new InvalidDataException(
                    "The resolved VisualTeX host no longer exists.");

            try
            {
                var metadata =
                    WordFormulaMetadataReader.TryReadEmbeddedNativeOle(resolved)
                    ?? throw new InvalidDataException(
                        "The VisualTeX OLE payload contains no authoritative formula metadata.");
                return new WordFormulaSemanticPayload
                {
                    Latex = metadata.Latex,
                    Metadata = metadata,
                };
            }
            finally { Release(resolved); }
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(range);
        }
    }

    private static string ReadLocalEquationXml(Range exactEquation)
    {
        var direct = TryReadExactlyOneEquation(exactEquation);
        if (direct is not null) return direct;

        // A few Word builds serialize an exact OMath range incompletely at a
        // field/run boundary. Retry only with a tiny live range around the same
        // equation. This is bounded local work, never a document fallback.
        Range? probe = null;
        try
        {
            probe = exactEquation.Duplicate;
            try { probe.MoveStart(WdUnits.wdCharacter, -2); } catch { }
            try { probe.MoveEnd(WdUnits.wdCharacter, 2); } catch { }
            var expanded = TryReadExactlyOneEquation(probe);
            if (expanded is not null) return expanded;
        }
        finally { Release(probe); }

        throw new InvalidDataException(
            "Word could not export one local OMML equation from the resolved host.");
    }

    private static string? TryReadExactlyOneEquation(Range range)
    {
        try
        {
            var xml = range.WordOpenXML;
            var package = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            var count = package.Descendants(MathNs + "oMath").Count();
            if (count != 1) return null;
            _ = WordOmmlConverter.ExtractSingleOMath(xml);
            return xml;
        }
        catch
        {
            return null;
        }
    }

    internal static Range CreateRange(
        Document document,
        WordFormulaRangeAddress address)
    {
        if (address.StoryType == WdStoryType.wdMainTextStory)
            return document.Range(address.Start, address.End);

        Range? story = null;
        Range? result = null;
        try
        {
            story = document.StoryRanges[address.StoryType];
            result = story.Duplicate;
            result.SetRange(address.Start, address.End);
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(result);
            Release(story);
        }
    }

    internal static bool SameAddress(
        Range range,
        WordFormulaRangeAddress address) =>
        range.StoryType == address.StoryType
        && range.Start == address.Start
        && range.End == address.End;

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
