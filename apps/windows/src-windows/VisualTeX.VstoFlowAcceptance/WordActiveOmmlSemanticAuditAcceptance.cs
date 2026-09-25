using System.Text.Json;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunActiveOmmlSemanticAuditAcceptance()
    {
        AssertTrue(
            AttachActiveWord,
            "Active OMML semantic audit must attach to the user's current Word instance.");

        var sessionPath =
            Environment.GetEnvironmentVariable(
                "VISUALTEX_ACTIVE_OMML_SESSION_PATH");
        if (string.IsNullOrWhiteSpace(sessionPath))
            throw new InvalidOperationException(
                "VISUALTEX_ACTIVE_OMML_SESSION_PATH must point to the source editor session.");
        sessionPath = Path.GetFullPath(sessionPath);
        if (!File.Exists(sessionPath))
            throw new FileNotFoundException(
                "The source editor session is missing.",
                sessionPath);

        var expectedMathMl =
            ReadSessionExportMathMl(
                sessionPath);
        var expectedLatex =
            ReadSessionLatex(
                sessionPath);
        var requestedDocumentName =
            Environment.GetEnvironmentVariable(
                "VISUALTEX_ACTIVE_OMML_DOCUMENT_NAME")
            ?? "文档1";

        Word.Application? application = null;
        Word.Document? document = null;
        Word.OMath? math = null;
        Word.Range? range = null;
        try
        {
            application =
                CreateWordApplication(
                    visible: false);
            for (var index = 1;
                 index <= application.Documents.Count;
                 index++)
            {
                Word.Document? candidate = null;
                try
                {
                    candidate =
                        application.Documents[index];
                    if (!string.Equals(
                            candidate.Name,
                            requestedDocumentName,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    document = candidate;
                    candidate = null;
                    break;
                }
                finally
                {
                    Release(candidate);
                }
            }

            document ??=
                application.Documents
                    .Cast<Word.Document>()
                    .FirstOrDefault(item =>
                        item.OMaths.Count > 0
                        && !item.Name.EndsWith(
                            ".docx",
                            StringComparison.OrdinalIgnoreCase))
                ?? throw new InvalidOperationException(
                    $"Unable to locate active OMML document '{requestedDocumentName}'.");

            var requestedIndexText =
                Environment.GetEnvironmentVariable(
                    "VISUALTEX_ACTIVE_OMML_INDEX");
            var equationIndex =
                int.TryParse(
                    requestedIndexText,
                    out var parsedIndex)
                    ? parsedIndex
                    : 1;
            if (string.IsNullOrWhiteSpace(requestedIndexText))
            {
                AssertEqual(
                    1,
                    document.OMaths.Count,
                    "Active OMML audit expected exactly one Document1 equation when no explicit index is supplied.");
            }
            if (equationIndex < 1
                || equationIndex > document.OMaths.Count)
            {
                throw new InvalidOperationException(
                    $"Requested OMath index {equationIndex} is outside 1..{document.OMaths.Count}.");
            }

            math = document.OMaths[equationIndex];
            range = math.Range.Duplicate;
            var resolved =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    range,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Document1 OMath could not be resolved locally.");
            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    resolved);
            if (string.IsNullOrWhiteSpace(
                    payload.WordOpenXml))
                throw new InvalidDataException(
                    "Document1 OMath returned no WordOpenXML.");

            var comparison =
                WordNativeOmmlSemanticComparer
                    .CompareMathMlToWordOpenXml(
                        expectedMathMl,
                        payload.WordOpenXml!,
                        display: math.Type ==
                            Word.WdOMathType.wdOMathDisplay);
            if (!comparison.Equivalent)
            {
                throw new InvalidDataException(
                    "Document1 current OMath is not semantically equivalent to the source editor MathML. "
                    + $"expectedMathMlSignature=[{comparison.ExpectedMathMlSignature}]; "
                    + $"actualMathMlSignature=[{comparison.ActualMathMlSignature}]; "
                    + $"expectedOmmlSignature=[{comparison.ExpectedOmmlSignature}]; "
                    + $"actualOmmlSignature=[{comparison.ActualOmmlSignature}]; "
                    + $"error=[{comparison.SemanticExtractionError ?? string.Empty}].");
            }

            Console.WriteLine(
                "[ACTIVE OMML SEMANTIC AUDIT PASS] "
                + $"document='{document.Name}'; "
                + $"omathIndex={equationIndex}; "
                + $"latex={expectedLatex}; "
                + $"wordLatex={payload.Latex}; "
                + $"exactOmml={comparison.ExactOmmlMatch}; "
                + $"semanticSignature={comparison.ActualMathMlSignature}");
        }
        finally
        {
            Release(range);
            Release(math);
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunActiveNativeNumberSemanticAuditAcceptance()
    {
        AssertTrue(
            AttachActiveWord,
            "Active native-number audit must attach to the user's current Word instance.");

        var requestedDocumentName =
            Environment.GetEnvironmentVariable(
                "VISUALTEX_ACTIVE_OMML_DOCUMENT_NAME")
            ?? "文档1";
        var requestedIndexText =
            Environment.GetEnvironmentVariable(
                "VISUALTEX_ACTIVE_OMML_INDEX");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.OMath? math = null;
        Word.Range? range = null;
        try
        {
            application =
                System.Runtime.InteropServices.Marshal.GetActiveObject(
                    "Word.Application")
                as Word.Application
                ?? throw new InvalidOperationException(
                    "No active Word instance is available.");

            for (var index = 1;
                 index <= application.Documents.Count;
                 index++)
            {
                Word.Document? candidate = null;
                try
                {
                    candidate = application.Documents[index];
                    if (!string.Equals(
                            candidate.Name,
                            requestedDocumentName,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    document = candidate;
                    candidate = null;
                    break;
                }
                finally
                {
                    Release(candidate);
                }
            }

            document ??=
                application.Documents
                    .Cast<Word.Document>()
                    .FirstOrDefault(item =>
                        item.OMaths.Count > 0)
                ?? throw new InvalidOperationException(
                    $"Unable to locate active OMML document '{requestedDocumentName}'.");

            var equationIndex =
                int.TryParse(
                    requestedIndexText,
                    out var parsedIndex)
                    ? parsedIndex
                    : document.OMaths.Count;
            if (equationIndex < 1
                || equationIndex > document.OMaths.Count)
            {
                throw new InvalidOperationException(
                    $"Requested OMath index {equationIndex} is outside 1..{document.OMaths.Count}.");
            }

            math = document.OMaths[equationIndex];
            range = math.Range.Duplicate;
            var rawWordOpenXml =
                range.WordOpenXML;
            AssertTrue(
                WordOmmlConverter.HasWordNativeEquationNumberHost(
                    rawWordOpenXml),
                "The selected active Word OMath is not a native '#(...)' numbered host.");

            var expectedSemanticOmml =
                WordOmmlConverter.StripWordNativeEquationNumberHost(
                    rawWordOpenXml);
            var resolved =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    range,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "The active Word-native numbered OMath could not be resolved locally.");
            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    resolved);
            if (string.IsNullOrWhiteSpace(
                    payload.WordOpenXml))
            {
                throw new InvalidDataException(
                    "The active Word-native numbered OMath returned no semantic WordOpenXML.");
            }

            AssertTrue(
                !WordOmmlConverter.HasWordNativeEquationNumberHost(
                    payload.WordOpenXml!),
                "The semantic reader leaked the Word-native number host into editor content.");

            var expectedSignature =
                WordOmmlConverter.ComputeImportedOmmlContentSignature(
                    expectedSemanticOmml);
            var actualSignature =
                WordOmmlConverter.ComputeImportedOmmlContentSignature(
                    payload.WordOpenXml!);
            AssertEqual(
                expectedSignature,
                actualSignature,
                "The semantic reader did not return exactly the native numbered OMath body.");

            if (payload.Latex.IndexOf(
                    "#",
                    StringComparison.Ordinal) >= 0)
            {
                throw new InvalidDataException(
                    "The semantic reader leaked a '#' equation-number token into LaTeX.");
            }

            Console.WriteLine(
                "[ACTIVE WORD NATIVE NUMBER SEMANTIC AUDIT PASS] "
                + $"document='{document.Name}'; "
                + $"omathIndex={equationIndex}; "
                + $"rawText='{range.Text.Replace("\r", " ").Replace("\a", " ")}'; "
                + $"semanticLatex='{payload.Latex}'; "
                + $"semanticSignature={actualSignature}");
        }
        finally
        {
            Release(range);
            Release(math);
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static string ReadSessionExportMathMl(
        string sessionPath)
    {
        using var document =
            JsonDocument.Parse(
                File.ReadAllText(
                    sessionPath));
        if (!document.RootElement.TryGetProperty(
                "exportResult",
                out var export)
            || !export.TryGetProperty(
                "mathMl",
                out var mathMl)
            || mathMl.ValueKind !=
                JsonValueKind.String
            || string.IsNullOrWhiteSpace(
                mathMl.GetString()))
        {
            throw new InvalidDataException(
                "The editor session has no exportResult.mathMl.");
        }
        return mathMl.GetString()!;
    }

    private static string ReadSessionLatex(
        string sessionPath)
    {
        using var document =
            JsonDocument.Parse(
                File.ReadAllText(
                    sessionPath));
        if (!document.RootElement.TryGetProperty(
                "lines",
                out var lines)
            || lines.ValueKind !=
                JsonValueKind.Array
            || lines.GetArrayLength() == 0)
            return string.Empty;
        var first = lines[0];
        return first.TryGetProperty(
                "latex",
                out var latex)
            && latex.ValueKind ==
                JsonValueKind.String
                ? latex.GetString()
                    ?? string.Empty
                : string.Empty;
    }
}
