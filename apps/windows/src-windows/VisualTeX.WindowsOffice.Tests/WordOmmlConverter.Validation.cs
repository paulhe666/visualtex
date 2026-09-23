using System.Xml.Linq;

namespace VisualTeX.WordVsto;

internal static partial class WordOmmlConverter
{
    internal static string ComputeVerifiedMaterializedOmmlFingerprint(
        string preparedWordOpenXml,
        string materializedWordOpenXml)
    {
        var comparison = WordNativeOmmlSemanticComparer.CompareOmml(
            preparedWordOpenXml,
            materializedWordOpenXml,
            display: false);
        if (!comparison.Equivalent)
            throw new InvalidDataException(
                "Test-only OMML materialization semantic mismatch.");
        return ComputeOmmlFingerprint(materializedWordOpenXml);
    }

    internal static string ComputeVerifiedMaterializedOmmlFingerprint(
        string preparedWordOpenXml,
        string materializedWordOpenXml)
    {
        var comparison =
            WordNativeOmmlSemanticComparer.CompareOmml(
                preparedWordOpenXml,
                materializedWordOpenXml,
                display: false);
        if (!comparison.Equivalent)
            throw new InvalidDataException(
                "Test-only OMML materialization semantic comparison failed.");
        return ComputeOmmlFingerprint(materializedWordOpenXml);
    }

    internal static void ValidateOmmlResult(string omml, string mathMl)
    {
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        XNamespace officeMath = MathNamespace;

        var target = XDocument.Parse(
            ExtractSingleOMath(omml),
            LoadOptions.PreserveWhitespace);
        ValidateNoVisibleEmptyOmmlSlots(target);
        var source = XDocument.Parse(
            mathMl,
            LoadOptions.PreserveWhitespace);
        var sourceTables = source
            .Descendants(presentationMath + "mtable")
            .Where(table => IsMatrixLikeMathMlTable(
                table,
                presentationMath))
            .ToArray();
        if (sourceTables.Length == 0) return;

        var targetMatrices =
            target.Descendants(officeMath + "m").ToArray();
        foreach (var matrix in targetMatrices)
        {
            if (!IsMathBooleanEnabled(
                    matrix.Element(officeMath + "mPr"),
                    officeMath + "plcHide"))
                throw new InvalidDataException(
                    "Office matrix OMML did not hide dotted placeholder slots.");
        }

        var matchedMatrices = new HashSet<XElement>();
        for (var tableIndex = 0;
             tableIndex < sourceTables.Length;
             tableIndex++)
        {
            var sourceRows = sourceTables[tableIndex]
                .Elements(presentationMath + "mtr")
                .ToArray();
            var sourceColumnCounts = sourceRows
                .Select(row => row
                    .Elements(presentationMath + "mtd")
                    .Count())
                .ToArray();
            var targetMatrix = targetMatrices
                .FirstOrDefault(candidate =>
                {
                    if (matchedMatrices.Contains(candidate))
                        return false;
                    var candidateRows =
                        candidate
                            .Elements(officeMath + "mr")
                            .ToArray();
                    return candidateRows.Length == sourceRows.Length
                        && candidateRows
                            .Select(row => row
                                .Elements(officeMath + "e")
                                .Count())
                            .SequenceEqual(sourceColumnCounts);
                });
            if (targetMatrix is null)
                throw new InvalidDataException(
                    $"Office did not preserve matrix {tableIndex + 1} dimensions "
                    + $"({sourceRows.Length}x{string.Join("/", sourceColumnCounts)}).");

            matchedMatrices.Add(targetMatrix);
            var targetRows = targetMatrix
                .Elements(officeMath + "mr")
                .ToArray();
            for (var rowIndex = 0;
                 rowIndex < sourceRows.Length;
                 rowIndex++)
            {
                var sourceCells = sourceRows[rowIndex]
                    .Elements(presentationMath + "mtd")
                    .ToArray();
                var targetCells = targetRows[rowIndex]
                    .Elements(officeMath + "e")
                    .ToArray();
                for (var cellIndex = 0;
                     cellIndex < sourceCells.Length;
                     cellIndex++)
                {
                    var sourceHasVisibleContent =
                        sourceCells[cellIndex]
                            .DescendantsAndSelf()
                            .Any(element =>
                                element.Name.Namespace == presentationMath
                                && element.Name.LocalName is
                                    "mi" or "mn" or "mo" or "mtext"
                                && !string.IsNullOrWhiteSpace(
                                    element.Value));
                    if (!sourceHasVisibleContent)
                        continue;

                    var targetHasVisibleContent =
                        targetCells[cellIndex]
                            .Descendants(officeMath + "t")
                            .Any(text =>
                                !string.IsNullOrWhiteSpace(
                                    text.Value));
                    if (!targetHasVisibleContent)
                        throw new InvalidDataException(
                            $"Office produced an empty matrix slot at matrix {tableIndex + 1}, "
                            + $"row {rowIndex + 1}, column {cellIndex + 1}.");
                }
            }
        }
    }

    internal static void ValidateMaterializedOmml(
        string wordOpenXml)
    {
        var equation = ExtractSingleOMath(wordOpenXml);
        var document = XDocument.Parse(
            equation,
            LoadOptions.PreserveWhitespace);
        ValidateNoVisibleEmptyOmmlSlots(document);
    }
}
