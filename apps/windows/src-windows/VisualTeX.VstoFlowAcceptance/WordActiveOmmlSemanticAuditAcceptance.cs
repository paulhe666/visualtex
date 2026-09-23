using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunActiveOmmlRoundTripAuditAcceptance()
    {
        AssertTrue(
            AttachActiveWord,
            "Active OMML round-trip audit must attach to the user's current Word instance.");

        // First guard the semantic-signature layer itself. Word commonly rewrites
        // a set-builder bar from an explicit token into m:d/m:sepChr=U+2223.
        // Those two MathML spellings must compare equal, while genuinely different
        // separators and fraction forms must remain distinguishable.
        const string explicitSeparator =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mfenced open='{' close='}'><mrow><mi>x</mi><mo>|</mo><mi>a</mi></mrow></mfenced></math>";
        const string fencedSeparator =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mfenced open='{' close='}' separators='∣'><mrow><mi>x</mi></mrow><mrow><mi>a</mi></mrow></mfenced></math>";
        const string differentSeparator =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mfenced open='{' close='}' separators=';'><mrow><mi>x</mi></mrow><mrow><mi>a</mi></mrow></mfenced></math>";
        const string explicitDividesInsideSetBuilder =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mfenced open='{' close='}'><mrow><mi>x</mi><mo>∣</mo><mi>a</mi></mrow></mfenced></math>";
        const string explicitAsciiBarInsideSetBuilder =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mfenced open='{' close='}'><mrow><mi>x</mi><mo>|</mo><mi>a</mi></mrow></mfenced></math>";
        const string dividesRelation =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mi>a</mi><mo>∣</mo><mi>b</mi></math>";
        const string asciiBarRelation =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mi>a</mi><mo>|</mo><mi>b</mi></math>";
        const string mathTypeIdentifierDivides =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mi>x</mi><mi>∣</mi><mi>a</mi></math>";
        const string operatorAsciiBar =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mi>x</mi><mo>|</mo><mi>a</mi></math>";
        const string absoluteValueFence =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mfenced open='|' close='|'><mi>x</mi></mfenced></math>";
        const string binaryVerticalRelation =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mi>x</mi><mo>|</mo><mi>y</mi></math>";
        const string barFraction =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mfrac><mi>a</mi><mi>b</mi></mfrac></math>";
        const string noBarFraction =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mfrac linethickness='0pt'><mi>a</mi><mi>b</mi></mfrac></math>";
        const string bevelledFraction =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mfrac bevelled='true'><mi>a</mi><mi>b</mi></mfrac></math>";

        var explicitSeparatorSignature =
            WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(explicitSeparator);
        var fencedSeparatorSignature =
            WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(fencedSeparator);
        var differentSeparatorSignature =
            WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(differentSeparator);
        AssertEqual(
            explicitSeparatorSignature,
            fencedSeparatorSignature,
            "Explicit set-builder bar and Word mfenced separator must have identical semantic signatures.");
        AssertTrue(
            !string.Equals(explicitSeparatorSignature, differentSeparatorSignature, StringComparison.Ordinal),
            "A genuinely different mfenced separator must remain semantically distinct.");
        AssertEqual(
            WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(explicitDividesInsideSetBuilder),
            WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(explicitAsciiBarInsideSetBuilder),
            "Word U+2223/ASCII vertical-line materialization inside a set-builder fence must compare equal.");
        AssertEqual(
            WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(dividesRelation),
            WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(asciiBarRelation),
            "U+2223 DIVIDES and ASCII vertical line must compare equal when both are ordinary relation/operator tokens.");
        AssertEqual(
            WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(mathTypeIdentifierDivides),
            WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(operatorAsciiBar),
            "MathType's <mi> U+2223 export and Word's <mo>|</mo> materialization must compare equal.");
        AssertTrue(
            !string.Equals(
                WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(absoluteValueFence),
                WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(binaryVerticalRelation),
                StringComparison.Ordinal),
            "A vertical-bar fence must remain structurally distinct from an ordinary vertical relation token.");

        var barFractionSignature =
            WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(barFraction);
        var noBarFractionSignature =
            WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(noBarFraction);
        var bevelledFractionSignature =
            WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(bevelledFraction);
        AssertTrue(
            !string.Equals(barFractionSignature, noBarFractionSignature, StringComparison.Ordinal),
            "A no-bar fraction must not compare equal to a normal fraction.");
        AssertTrue(
            !string.Equals(barFractionSignature, bevelledFractionSignature, StringComparison.Ordinal),
            "A bevelled fraction must not compare equal to a normal fraction.");

        const string accentHat =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mover accent='true'><mi>x</mi><mo>^</mo></mover></math>";
        const string upperLimitHat =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mover><mi>x</mi><mo>^</mo></mover></math>";
        AssertTrue(
            !string.Equals(
                WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(accentHat),
                WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(upperLimitHat),
                StringComparison.Ordinal),
            "A real accent must not compare equal to an upper-limit/overset structure that happens to use the same glyph.");

        const string plainBox =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><menclose notation='box'><mi>x</mi></menclose></math>";
        const string boxWithVerticalStrike =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><menclose notation='box verticalstrike'><mi>x</mi></menclose></math>";
        const string topBorderOnly =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><menclose notation='top'><mi>x</mi></menclose></math>";
        const string noBorder =
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mi>x</mi></math>";
        AssertTrue(
            !string.Equals(
                WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(plainBox),
                WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(boxWithVerticalStrike),
                StringComparison.Ordinal),
            "A vertical strike inside a border box must remain visible to semantic validation.");
        AssertTrue(
            !string.Equals(
                WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(topBorderOnly),
                WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(noBorder),
                StringComparison.Ordinal),
            "A partial border must not collapse to plain formula content.");

        var preparedSeparatorOmml =
            WordOmmlConverter.TransformMathMlToOmml(explicitSeparator);
        var wordMaterializedSeparatorOmml =
            WordOmmlConverter.TransformMathMlToOmml(fencedSeparator);
        var differentSeparatorOmml =
            WordOmmlConverter.TransformMathMlToOmml(differentSeparator);
        var preparedRawSignature =
            WordOmmlConverter.ComputeImportedOmmlContentSignature(preparedSeparatorOmml);
        var materializedRawSignature =
            WordOmmlConverter.ComputeImportedOmmlContentSignature(wordMaterializedSeparatorOmml);
        AssertTrue(
            !string.Equals(preparedRawSignature, materializedRawSignature, StringComparison.Ordinal),
            "The regression fixture must preserve the raw OMML shape difference that Word introduces for delimiter separators.");
        var equivalentSeparatorComparison =
            WordNativeOmmlSemanticComparer.CompareOmml(
                preparedSeparatorOmml,
                wordMaterializedSeparatorOmml,
                display: false);
        AssertTrue(
            equivalentSeparatorComparison.Equivalent,
            "Test-only semantic comparer must accept the equivalent delimiter-separator rewrite.");

        var differentSeparatorComparison =
            WordNativeOmmlSemanticComparer.CompareOmml(
                preparedSeparatorOmml,
                differentSeparatorOmml,
                display: false);
        AssertTrue(
            !differentSeparatorComparison.Equivalent,
            "Test-only semantic comparer must reject a genuinely different delimiter separator.");

        XNamespace wordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace mathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var materializedEquation =
            XDocument.Parse(wordMaterializedSeparatorOmml)
                .Descendants(mathNamespace + "oMath")
                .Single();
        var identityBody = new XElement(
            wordNamespace + "body",
            new XElement(wordNamespace + "p", new XElement(materializedEquation)));
        var formulaId = Guid.NewGuid().ToString("D");
        var expectedSignatures = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            [formulaId] = preparedRawSignature,
        };
        var liveOwners = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            [formulaId] = 0,
        };
        var freshSources = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            [formulaId] = preparedSeparatorOmml,
        };
        var resolvedIdentity = WordOmmlNativeSource.IndexConversionEquationIdentities(
            identityBody,
            expectedSignatures,
            WordOmmlConverter.ComputeImportedOmmlContentSignature,
            liveInsertedOwners: liveOwners,
            allowFreshOwnerFingerprintRecovery: true,
            freshSourceOmml: freshSources);
        AssertTrue(
            resolvedIdentity.TryGetValue(formulaId, out var owner)
            && owner.Index == 0,
            "Fresh OMML identity indexing must retain the live physical owner when Word only changed the native representation.");

        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application =
                System.Runtime.InteropServices.Marshal.GetActiveObject(
                    "Word.Application")
                as Word.Application
                ?? throw new InvalidOperationException(
                    "No active Word instance is available.");
            document = application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "Word has no active document.");

            var failures = new List<string>();
            var structureCounts = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var index = 1; index <= document.OMaths.Count; index++)
            {
                Word.OMath? math = null;
                Word.Range? range = null;
                try
                {
                    math = document.OMaths[index];
                    range = math.Range.Duplicate;
                    var wordOpenXml = range.WordOpenXML ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(wordOpenXml))
                    {
                        failures.Add($"#{index}: empty WordOpenXML");
                        continue;
                    }

                    var display = math.Type == Word.WdOMathType.wdOMathDisplay;
                    var mathMl = WordOmmlConverter.TransformOmmlToMathMl(wordOpenXml, display);
                    var roundTripOmml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
                    var comparison = WordNativeOmmlSemanticComparer.CompareOmml(
                        wordOpenXml,
                        roundTripOmml,
                        display);

                    var localOmml = WordOmmlConverter.ExtractSingleOMath(wordOpenXml);
                    var source = System.Xml.Linq.XDocument.Parse(
                        localOmml,
                        System.Xml.Linq.LoadOptions.PreserveWhitespace);
                    const string officeMathNamespace =
                        "http://schemas.openxmlformats.org/officeDocument/2006/math";
                    foreach (var element in source.Descendants()
                                 .Where(element => element.Name.NamespaceName == officeMathNamespace)
                                 .Where(element => element.Name.LocalName is
                                     "d" or "f" or "rad" or "acc" or "bar" or "nary"
                                     or "groupChr" or "m" or "eqArr" or "limLow" or "limUpp"
                                     or "sSub" or "sSup" or "sSubSup" or "sPre"
                                     or "func" or "box" or "borderBox" or "phant"))
                    {
                        structureCounts.TryGetValue(element.Name.LocalName, out var count);
                        structureCounts[element.Name.LocalName] = count + 1;
                    }

                    var signature =
                        WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(mathMl);
                    Console.WriteLine(
                        $"[ACTIVE OMML ROUNDTRIP] #{index} "
                        + $"range={range.Start}:{range.End} "
                        + $"display={(display ? "block" : "inline")} "
                        + $"equivalent={comparison.Equivalent} "
                        + $"exactOmml={comparison.ExactOmmlMatch} "
                        + $"signature={signature}");
                    if (!comparison.Equivalent)
                    {
                        failures.Add(
                            $"#{index}: expectedMathMl=[{comparison.ExpectedMathMlSignature}] "
                            + $"actualMathMl=[{comparison.ActualMathMlSignature}] "
                            + $"error=[{comparison.SemanticExtractionError ?? string.Empty}]");
                    }
                }
                catch (Exception error)
                {
                    failures.Add(
                        $"#{index}: {error.GetType().Name}: {error.Message}");
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }

            Console.WriteLine(
                "[ACTIVE OMML STRUCTURES] "
                + string.Join(
                    ", ",
                    structureCounts
                        .OrderBy(pair => pair.Key, StringComparer.Ordinal)
                        .Select(pair => pair.Key + "=" + pair.Value)));

            if (failures.Count != 0)
            {
                throw new InvalidDataException(
                    "Active-document OMML round-trip semantic audit found "
                    + failures.Count
                    + " failure(s):\n"
                    + string.Join("\n", failures));
            }

            Console.WriteLine(
                $"[ACTIVE OMML ROUNDTRIP PASS] document='{document.Name}'; equations={document.OMaths.Count}");
        }
        finally
        {
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

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
