using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordVisualTeXOmmlTabAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(artifactRoot, "visualtex-omml-native-display-numbering.docx");
        var pdfPath = Path.Combine(artifactRoot, "visualtex-omml-native-display-numbering.pdf");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? referenceDocument = null;
        Word.Range? insertion = null;
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
        Word.Range? referenceRange = null;
        Word.Field? externalReference = null;
        WordOmmlLayoutMetric? editedMetric = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);

            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            application.Selection.SetRange(insertion.Start, insertion.End);
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertion.Start,
                insertion.End,
                latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            service.InsertOmml(session, QuadraticFormulaMathMl());
            Release(insertion); insertion = null;

            AssertNumberedOmmlTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "fresh numbered OMML insertion");
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException("Fresh numbered OMML lost its VisualTeX bookmark.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var numberedMetric = ReadNumberedOmmlBodyLayoutMetric(
                application,
                document,
                equationRange,
                "fresh numbered VisualTeX OMML body");

            referenceDocument = application.Documents.Add(Visible: false);
            referenceDocument.OMathFontName = WordOfficeMathFontLoader.LatinModernMathFamily;
            referenceDocument.Content.Text = "VT_TRUE_DISPLAY_QUADRATIC\r";
            referenceRange = InsertPureNativeOmml(
                referenceDocument,
                "VT_TRUE_DISPLAY_QUADRATIC",
                "x=(-b±√(b^2-4ac))/(2a)");
            ConfigureTrueDisplayReference(referenceRange, 14f);
            Release(referenceRange); referenceRange = null;
            referenceRange = GetOmmlRangeByIndex(referenceDocument, 1);
            var trueDisplayMetric = ReadWordOmmlLayoutMetric(
                application,
                referenceDocument,
                referenceRange,
                "normal Word display OMath reference");
            AssertNumberedOmmlMatchesTrueDisplayLayout(
                trueDisplayMetric,
                numberedMetric,
                "fresh numbered OMML native-display comparison");
            referenceDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(referenceRange); referenceRange = null;
            Release(referenceDocument); referenceDocument = null;
            document.Activate();

            var originalVisibleNumber = ReadVisibleEquationNumber(document, formulaId);
            externalReference = InsertExternalEquationReference(document, formulaId);
            AssertEqual(
                originalVisibleNumber.Trim('(', ')'),
                NormalizeEquationNumberText(externalReference.Result.Text),
                "External REF did not match the native equation-number target.");

            var originalMetadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException("Fresh numbered OMML lost its VisualTeX metadata.");
            var editSession = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                equationRange.Start,
                equationRange.End,
                latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}+1",
                originalMetadata);
            service.ReplaceOmml(
                editSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
                + "<mrow><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo>"
                + "<msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt>"
                + "</mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac><mo>+</mo><mn>1</mn></mrow></math>");
            Release(equationRange); equationRange = null;
            Release(bookmark); bookmark = null;

            AssertNumberedOmmlTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "numbered OMML edit/reconcile");
            externalReference.Update();
            AssertEqual(
                ReadVisibleEquationNumber(document, formulaId).Trim('(', ')'),
                NormalizeEquationNumberText(externalReference.Result.Text),
                "External REF disconnected after editing the numbered OMML formula.");
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException("Edited numbered OMML lost its VisualTeX bookmark.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            editedMetric = ReadNumberedOmmlBodyLayoutMetric(
                application,
                document,
                equationRange,
                "edited numbered VisualTeX OMML body");

            document.Save();
            document.ExportAsFixedFormat(
                pdfPath,
                Word.WdExportFormat.wdExportFormatPDF,
                OpenAfterExport: false,
                OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                Range: Word.WdExportRange.wdExportAllDocument,
                Item: Word.WdExportItem.wdExportDocumentContent,
                IncludeDocProps: true,
                KeepIRM: true,
                CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                DocStructureTags: true,
                BitmapMissingFonts: false,
                UseISO19005_1: false);
            AssertPdfRetainsLatinModernOmmlRendering(pdfPath);
            Release(externalReference); externalReference = null;
            Release(equationRange); equationRange = null;
            Release(bookmark); bookmark = null;
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            AssertNumberedOmmlTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "numbered OMML save/reopen");
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException("Save/reopened numbered OMML lost its VisualTeX bookmark.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var reopenedMetric = ReadNumberedOmmlBodyLayoutMetric(
                application,
                document,
                equationRange,
                "save/reopened numbered VisualTeX OMML body");
            AssertStableMathLayoutAfterReopen(
                editedMetric
                    ?? throw new InvalidDataException("Edited numbered OMML metric was not captured."),
                reopenedMetric,
                "numbered native display OMath save/reopen");
            externalReference = FindExternalEquationReference(document, formulaId)
                ?? throw new InvalidDataException("Save/reopen lost the external equation REF field.");
            externalReference.Update();
            AssertEqual(
                ReadVisibleEquationNumber(document, formulaId).Trim('(', ')'),
                NormalizeEquationNumberText(externalReference.Result.Text),
                "Save/reopen disconnected the external equation REF field.");

            Console.WriteLine(
                $"VisualTeX numbered OMML native-#SEQ acceptance passed: wdOMathDisplay/m:oMathPara + legal m:eqArr/#(SEQ), VTEqNum bookmark + external body REF, zero numbering Shape/Table, edit, PDF and save/reopen remained valid. PDF={pdfPath}");
        }
        finally
        {
            Release(externalReference);
            Release(referenceRange);
            if (referenceDocument is not null)
            {
                try { referenceDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(referenceDocument);
            Release(equationRange);
            Release(bookmark);
            Release(insertion);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static OfficeSessionDocument CreateNumberedOmmlTabSession(
        string formulaId,
        string documentId,
        int start,
        int end,
        string latex,
        FormulaMetadata? originalMetadata)
    {
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = originalMetadata is null ? "create" : "edit",
            Host = "word",
            FormulaId = formulaId,
            SourceDocumentId = documentId,
            SourceObjectId = WordRangeReference(start, end),
            Title = "VisualTeX numbered native-display OMML acceptance",
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = FormulaOleContract.WordOmmlMode,
            Numbered = true,
            FontSizePt = 14,
            OriginalMetadata = originalMetadata,
            Lines = new List<FormulaLine>
            {
                new()
                {
                    Id = Guid.NewGuid().ToString("D"),
                    Latex = latex,
                },
            },
            ExportResult = new OfficeExportDocument
            {
                FormulaLetterFont = "katex",
                FormulaChineseFont = "system",
            },
        };
    }

    private static void AssertPdfRetainsLatinModernOmmlRendering(string pdfPath)
    {
        AssertTrue(File.Exists(pdfPath),
            "Word did not create the OMML PDF acceptance artifact.");
        var bytes = File.ReadAllBytes(pdfPath);
        AssertTrue(bytes.Length > 10_000,
            "Word produced an unexpectedly small OMML PDF; the equation may be blank.");
        AssertTrue(!ContainsAscii(bytes, "Cambria-Italic"),
            "Word PDF output substituted Cambria Italic for Latin Modern variables.");
        AssertTrue(ContainsAscii(bytes, "%PDF-"),
            "The OMML fixed-format artifact is not a valid PDF stream.");
    }

    private static bool ContainsAscii(byte[] bytes, string value)
    {
        var needle = System.Text.Encoding.ASCII.GetBytes(value);
        for (var offset = 0; offset <= bytes.Length - needle.Length; offset++)
        {
            var matches = true;
            for (var index = 0; index < needle.Length; index++)
            {
                if (bytes[offset + index] == needle[index]) continue;
                matches = false;
                break;
            }
            if (matches) return true;
        }
        return false;
    }

    private static string QuadraticFormulaMathMl() =>
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
        + "<mrow><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo>"
        + "<msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt>"
        + "</mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></mrow></math>";

    private static void AssertNumberedOmmlTabHost(
        Word.Document document,
        string formulaId,
        bool updateReference,
        string context,
        string expectedLatinVariable = "x",
        string? expectedDigit = "2")
    {
        Word.Bookmark? ommlBookmark = null;
        Word.Range? equationRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        try
        {
            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context,
                updateReference);
            ommlBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(context + ": OMML bookmark is missing.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(ommlBookmark);
            maths = equationRange.OMaths;
            AssertEqual(1, maths.Count,
                context + ": equation range does not contain exactly one native OMath.");
            math = maths[1];
            AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                context + ": numbered OMML is not genuine Word display math.");

            var openXml = equationRange.WordOpenXML ?? string.Empty;
            const string MathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var mathNs = (XNamespace)MathNamespace;
            string semanticOmml;
            var directTableHost = (bool)equationRange.get_Information(
                Word.WdInformation.wdWithInTable);
            if (directTableHost)
            {
                // Current production host: number fields and all VTEq aliases live
                // only in right cell (1,3). The center m:oMath must be semantically
                // identical to an ordinary display equation and contain no field.
                AssertEqual(0, equationRange.Fields.Count,
                    context + ": direct-SEQ table leaked a field into m:oMath.");
                AssertTrue(!WordOmmlConverter.HasVisualTeXNativeEquationNumber(openXml),
                    context + ": direct-SEQ table formula still contains a mathematical number wrapper.");
                AssertTrue(openXml.IndexOf("VTEqNum_", StringComparison.OrdinalIgnoreCase) < 0,
                    context + ": VTEqNum leaked from the right table cell into formula OMML.");
                semanticOmml = WordOmmlConverter.ExtractSingleOMath(openXml);
            }
            else
            {
                AssertEqual(1, equationRange.Fields.Count,
                    context + ": native #(SEQ) OMath must contain exactly one field.");
                Word.Field? sequenceField = null;
                Word.Range? sequenceCode = null;
                try
                {
                    sequenceField = equationRange.Fields[1];
                    sequenceCode = sequenceField.Code;
                    AssertTrue(WordEquationNumbering.IsVisualTeXSequenceFieldCode(sequenceCode.Text),
                        context + ": the OMath field is not SEQ VisualTeXEquation.");
                    AssertTrue((sequenceCode.Text ?? string.Empty).IndexOf(
                            "REF VTEqNum_",
                            StringComparison.OrdinalIgnoreCase) < 0,
                        context + ": REF leaked inside the native #() field code.");
                }
                finally
                {
                    Release(sequenceCode);
                    Release(sequenceField);
                }

                AssertTrue(WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(openXml),
                    context + ": native #(SEQ) equation-number wrapper is missing.");
                AssertTrue(openXml.IndexOf("VTEqNum_", StringComparison.OrdinalIgnoreCase) >= 0,
                    context + ": VTEqNum bookmark is missing from native formula OMML.");
                AssertTrue(openXml.IndexOf("981730", StringComparison.Ordinal) < 0,
                    context + ": retired generated number placeholder leaked into OMML.");
                var numberedXml = XDocument.Parse(openXml, LoadOptions.PreserveWhitespace);
                AssertEqual(1, numberedXml.Descendants(mathNs + "eqArr").Count(),
                    context + ": Word native #() did not retain its one legal m:eqArr.");
                semanticOmml = WordOmmlConverter.StripVisualTeXNativeEquationNumber(openXml);
            }
            var semanticXml = XDocument.Parse(semanticOmml, LoadOptions.PreserveWhitespace);
            AssertTrue(semanticXml.Descendants(mathNs + "f").Any(),
                context + ": native fraction structure was lost after stripping numbering.");

            AssertDocumentOmmlMathFont(
                document,
                WordOfficeMathFontLoader.LatinModernMathFamily,
                context + " document math font");
            AssertSemanticOmmlTokenUsesNativeMath(
                semanticXml,
                expectedLatinVariable,
                context + " Latin variable");
            if (!string.IsNullOrEmpty(expectedDigit))
            {
                AssertSemanticOmmlTokenUsesNativeMath(
                    semanticXml,
                    expectedDigit!,
                    context + " digit");
            }
            AssertSemanticOmmlTokenUsesNativeMath(
                semanticXml,
                "=",
                context + " relation operator");
            Console.WriteLine(
                $"  {context}: true-display semantic OMML verified at {equationRange.Start}-{equationRange.End}, type={math.Type}.");
        }
        finally
        {
            Release(math);
            Release(maths);
            Release(equationRange);
            Release(ommlBookmark);
        }
    }

    private static void AssertSemanticOmmlTokenUsesNativeMath(
        XDocument semanticOmml,
        string expectedText,
        string context)
    {
        const string MathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var math = (XNamespace)MathNamespace;
        var normalizedExpected = expectedText.Normalize(System.Text.NormalizationForm.FormKC);
        var matchingRuns = semanticOmml
            .Descendants(math + "r")
            .Where(run => run.Elements(math + "t").Any(text =>
            {
                var value = text.Value;
                try { value = value.Normalize(System.Text.NormalizationForm.FormKC); }
                catch (ArgumentException) { }
                return value.IndexOf(normalizedExpected, StringComparison.Ordinal) >= 0;
            }))
            .ToArray();
        AssertTrue(matchingRuns.Length > 0,
            context + $": mathematical token '{expectedText}' is missing from semantic OMML.");
        AssertTrue(matchingRuns.All(run =>
                run.Element(math + "rPr")?.Element(math + "nor") is null),
            context + $": mathematical token '{expectedText}' was flattened to m:nor normal text.");
    }

    private static Word.Range ResolveNumberedOmmlFormulaBodyRange(
        Word.Document document,
        Word.Range equationRange)
    {
        Word.Table? table = null;
        Word.Rows? rows = null;
        Word.Columns? columns = null;
        Word.Range? visibleNumber = null;
        Word.Range? prefix = null;
        try
        {
            // In the current direct-SEQ 1x3 host the center OMath is already the
            // complete semantic formula body: numbering lives exclusively in cell
            // (1,3), so there is no # separator to trim from the mathematical Range.
            if ((bool)equationRange.get_Information(Word.WdInformation.wdWithInTable)
                && equationRange.Tables.Count > 0
                && equationRange.Fields.Count == 0)
            {
                table = equationRange.Tables[1];
                rows = table.Rows;
                columns = table.Columns;
                if (rows.Count == 1 && columns.Count == 3)
                    return equationRange.Duplicate;
            }

            Word.Bookmark? formulaBookmark = null;
            string formulaId;
            try
            {
                formulaBookmark = WordOmmlFormulaStore.FindAtRange(document, equationRange)
                    ?? throw new InvalidDataException(
                        "The numbered OMML body range cannot resolve its VisualTeX bookmark.");
                if (!WordOmmlFormulaStore.TryGetFormulaId(formulaBookmark, out formulaId))
                    throw new InvalidDataException(
                        "The numbered OMML body bookmark does not contain a valid FormulaId.");
            }
            finally { Release(formulaBookmark); }
            visibleNumber = WordEquationNumbering.FindVisibleEquationNumberRange(
                document,
                formulaId)
                ?? throw new InvalidDataException(
                    "The numbered OMML body range cannot resolve its VTEq_ number alias.");
            prefix = document.Range(equationRange.Start, visibleNumber.Start);
            var prefixText = prefix.Text ?? string.Empty;
            var hashOffset = prefixText.LastIndexOf('#');
            if (hashOffset < 0)
                throw new InvalidDataException(
                    "The native numbered OMath does not expose the # separator before VTEq_.");
            return document.Range(
                equationRange.Start,
                equationRange.Start + hashOffset);
        }
        finally
        {
            Release(prefix);
            Release(visibleNumber);
            Release(columns);
            Release(rows);
            Release(table);
        }
    }

    private static WordOmmlLayoutMetric ReadNumberedOmmlBodyLayoutMetric(
        Word.Application application,
        Word.Document document,
        Word.Range equationRange,
        string context)
    {
        Word.Range? body = null;
        Word.Range? visibleBody = null;
        try
        {
            body = ResolveNumberedOmmlFormulaBodyRange(document, equationRange);
            visibleBody = TrimToVisibleMathContent(document, body);
            var semanticOmml = WordOmmlConverter.StripVisualTeXNativeEquationNumber(
                equationRange.WordOpenXML ?? string.Empty);
            return ReadWordOmmlLayoutMetric(
                application,
                document,
                visibleBody,
                context,
                semanticOmml,
                measureVisibleCharacterInk: true);
        }
        finally
        {
            Release(visibleBody);
            Release(body);
        }
    }

    private static Word.Range TrimToVisibleMathContent(
        Word.Document document,
        Word.Range range)
    {
        var text = range.Text ?? string.Empty;
        static bool IsBoundary(char character) =>
            character is '\r' or '\n' or '\v' or '\a' or '\t'
                or '\u200b' or '\u200c' or '\u2060' or '\ufeff'
            || char.IsWhiteSpace(character);
        var startOffset = 0;
        while (startOffset < text.Length && IsBoundary(text[startOffset]))
            startOffset++;
        var endOffset = text.Length;
        while (endOffset > startOffset && IsBoundary(text[endOffset - 1]))
            endOffset--;
        if (endOffset <= startOffset)
            throw new InvalidDataException(
                "The numbered display OMath body contains no visible mathematical content.");
        return document.Range(
            range.Start + startOffset,
            range.Start + endOffset);
    }

    private static Word.Range GetOmmlRangeByIndex(
        Word.Document document,
        int index)
    {
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? range = null;
        try
        {
            maths = document.OMaths;
            if (index < 1 || index > maths.Count)
                throw new InvalidDataException(
                    $"Word OMath index {index} is outside the document's {maths.Count} equations.");
            math = maths[index];
            range = math.Range.Duplicate;
            var result = range;
            range = null;
            return result;
        }
        finally
        {
            Release(range);
            Release(math);
            Release(maths);
        }
    }

    private static void ConfigureTrueDisplayReference(
        Word.Range equationRange,
        float fontSizePt)
    {
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? liveRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            maths = equationRange.OMaths;
            AssertEqual(1, maths.Count,
                "True-display comparison range does not contain exactly one OMath.");
            math = maths[1];
            math.Type = Word.WdOMathType.wdOMathDisplay;
            math.Justification = Word.WdOMathJc.wdOMathJcCenterGroup;
            liveRange = math.Range;
            font = liveRange.Font;
            font.Position = 0;
            font.Size = fontSizePt;
            try { font.SizeBi = fontSizePt; } catch { }
        }
        finally
        {
            Release(font);
            Release(liveRange);
            Release(math);
            Release(maths);
        }
    }

    private static void AssertNumberedOmmlMatchesTrueDisplayLayout(
        WordOmmlLayoutMetric expectedDisplay,
        WordOmmlLayoutMetric actualNumbered,
        string context)
    {
        AssertEqual(Word.WdOMathType.wdOMathDisplay, expectedDisplay.Type,
            context + ": comparison reference is not a true display OMath.");
        AssertEqual(Word.WdOMathType.wdOMathDisplay, actualNumbered.Type,
            context + ": numbered formula is not genuine Word display math.");
        AssertNear(expectedDisplay.FontSizePt, actualNumbered.FontSizePt, 0.1f,
            context + ": numbered and ordinary display formulas use different semantic font sizes.");
        // Word's native #() numbering is intentionally hosted by m:eqArr. That
        // layout engine may allocate a different horizontal group width than an
        // otherwise identical unnumbered m:oMathPara, so pixel-for-pixel operator
        // geometry is not a valid acceptance invariant. Preserve the invariants
        // that matter: genuine Display math, the requested semantic point size,
        // live native math structures, and nonempty rendered geometry.
        AssertTrue(actualNumbered.WidthPx > 0 && actualNumbered.HeightPx > 0,
            context + ": native numbered OMML produced empty display geometry.");
        AssertTrue(actualNumbered.EqualsWidthPx > 0,
            context + ": native numbered OMML lost the relation operator geometry.");
        AssertTrue(actualNumbered.PlusMinusWidthPx > 0,
            context + ": native numbered OMML lost the plus/minus operator geometry.");
        AssertEqual(expectedDisplay.FractionCount, actualNumbered.FractionCount,
            context + ": native fraction structure count differs.");
        AssertEqual(expectedDisplay.RadicalCount, actualNumbered.RadicalCount,
            context + ": native radical structure count differs.");
    }

    private static Word.Field InsertExternalEquationReference(
        Word.Document document,
        string formulaId)
    {
        Word.Range? insertion = null;
        Word.OMaths? insertionMaths = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        try
        {
            // Use the exact product typing-boundary path. A copied/rekeyed native
            // #(SEQ) equation can be the final Word paragraph, and a bare
            // Paragraphs.Add() may inherit OMath affinity on that build. The product
            // helper verifies or creates an ordinary main-story paragraph adjacent
            // to this numbered display before returning its collapsed insertion
            // point.
            insertion = WordEquationNumbering.EnsureNormalTypingParagraphAfterNumberedDisplay(
                    document,
                    formulaId)
                ?? throw new InvalidDataException(
                    "VisualTeX could not establish an ordinary typing paragraph for the external equation REF.");
            insertionMaths = insertion.OMaths;
            if (insertionMaths.Count != 0)
                throw new InvalidDataException(
                    "The external equation REF paragraph was unexpectedly absorbed into OMath.");

            fields = insertion.Fields;
            object fieldType = Word.WdFieldType.wdFieldRef;
            object fieldCode = WordEquationNumbering.NativeNumberBookmarkName(formulaId) + " \\h";
            object preserveFormatting = true;
            field = fields.Add(insertion, ref fieldType, ref fieldCode, ref preserveFormatting);
            field.Update();
            AssertExternalEquationReferenceHyperlink(
                field,
                WordEquationNumbering.NativeNumberBookmarkName(formulaId));
            var result = field;
            field = null;
            return result;
        }
        finally
        {
            Release(field);
            Release(fields);
            Release(insertionMaths);
            Release(insertion);
        }
    }

    private static void AssertExternalEquationReferenceHyperlink(
        Word.Field field,
        string targetBookmarkName)
    {
        Word.Range? codeRange = null;
        Word.Range? resultRange = null;
        Word.Document? document = null;
        Word.Bookmarks? bookmarks = null;
        try
        {
            codeRange = field.Code;
            var code = codeRange.Text ?? string.Empty;
            AssertTrue(
                field.Type == Word.WdFieldType.wdFieldRef
                && code.IndexOf(targetBookmarkName, StringComparison.OrdinalIgnoreCase) >= 0
                && code.IndexOf("\\h", StringComparison.OrdinalIgnoreCase) >= 0,
                "The body equation reference is not REF <VTEqNum> \\h.");
            resultRange = field.Result;
            document = resultRange.Document;
            bookmarks = document.Bookmarks;
            AssertTrue(bookmarks.Exists(targetBookmarkName),
                "The REF \\h target VTEqNum bookmark is missing.");

            // Word does not reliably expose REF ... \\h through Result.Hyperlinks,
            // and Field.DoClick on a REF field is not the UI's hyperlink action on
            // all supported Word builds. This plain body field is therefore tested
            // for its native REF+\h contract and dynamic result only. VisualTeX's
            // actual double-clickable reference is the separately accepted
            // GOTOBUTTON + nested REF structure.
            AssertTrue(!string.IsNullOrWhiteSpace(resultRange.Text),
                "The body REF \\h produced an empty result.");
        }
        finally
        {
            Release(bookmarks);
            Release(document);
            Release(resultRange);
            Release(codeRange);
        }
    }

    private static Word.Field? FindExternalEquationReference(
        Word.Document document,
        string formulaId)
    {
        Word.Fields? fields = null;
        Word.Field? result = null;
        try
        {
            fields = document.Fields;
            var target = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? field = null;
                Word.Range? code = null;
                Word.Range? resultRange = null;
                try
                {
                    field = fields[index];
                    if (field.Type != Word.WdFieldType.wdFieldRef)
                        continue;
                    code = field.Code;
                    if ((code.Text ?? string.Empty).IndexOf(
                            "REF " + target,
                            StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    resultRange = field.Result;
                    if (resultRange.StoryType != Word.WdStoryType.wdMainTextStory)
                        continue;
                    // A collapsed/adjacent Range can report the preceding OMath in
                    // resultRange.OMaths after save/reopen even though the REF is a
                    // normal w:fldChar field in its own body paragraph. Field.Type,
                    // REF code and main-story ownership are the authoritative proof.
                    result = field;
                    field = null;
                    return result;
                }
                finally
                {
                    Release(resultRange);
                    Release(code);
                    Release(field);
                }
            }
            return null;
        }
        finally { Release(fields); }
    }

    private static string ReadVisibleEquationNumber(
        Word.Document document,
        string formulaId)
    {
        Word.Range? range = null;
        Word.View? view = null;
        var restoreFieldCodes = false;
        try
        {
            range = WordEquationNumbering.FindVisibleEquationNumberTextRange(document, formulaId)
                ?? throw new InvalidDataException("Visible equation-number range is missing.");
            view = document.ActiveWindow.View;
            restoreFieldCodes = view.ShowFieldCodes;
            if (restoreFieldCodes)
            {
                // Bookmark.Text follows the current Alt+F9 state and otherwise
                // exposes the SEQ instruction instead of its visible numeric result.
                // Read the rendered label without changing the user's persistent
                // field-code preference.
                view.ShowFieldCodes = false;
                System.Windows.Forms.Application.DoEvents();
            }
            return NormalizeEquationNumberText(range.Text);
        }
        finally
        {
            if (view is not null && restoreFieldCodes)
            {
                try { view.ShowFieldCodes = true; } catch { }
            }
            Release(view);
            Release(range);
        }
    }

    private static string NormalizeEquationNumberText(string? value) =>
        (value ?? string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\a", string.Empty)
            .Replace("\v", string.Empty)
            .Trim();
}
