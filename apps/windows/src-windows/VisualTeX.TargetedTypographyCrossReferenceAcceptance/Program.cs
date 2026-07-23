using System.Collections;
using System.Drawing;
using System.Drawing.Imaging;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Word = Microsoft.Office.Interop.Word;
using WordRange = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.TargetedTypographyCrossReferenceAcceptance;

internal sealed class UprightFixtureRoot
{
    public List<string> ForbiddenInternalCommands { get; set; } = new();
    public List<UprightFixtureCase> Cases { get; set; } = new();
}

internal sealed class UprightFixtureCase
{
    public string Name { get; set; } = string.Empty;
    public string Input { get; set; } = string.Empty;
    public string Normalized { get; set; } = string.Empty;
    public string MathMl { get; set; } = string.Empty;
    public List<string> ExpectedFragments { get; set; } = new();
    public List<string> NormalMathMlTokens { get; set; } = new();
}

internal static class Program
{
    private const float BodyFontSize = 14f;
    private static readonly List<string> Report = new();
    private static string _reportPath = string.Empty;

    [STAThread]
    private static int Main(string[] args)
    {
        var fixturePath = args.Length > 0
            ? Path.GetFullPath(args[0])
            : throw new ArgumentException("The upright MathML fixture path is required.");
        var artifactRoot = args.Length > 1
            ? Path.GetFullPath(args[1])
            : Path.Combine(
                Path.GetDirectoryName(fixturePath) ?? Environment.CurrentDirectory,
                "word");
        var skipNumberingFont = args.Any(argument =>
            string.Equals(argument, "--skip-numbering-font", StringComparison.OrdinalIgnoreCase));
        Directory.CreateDirectory(artifactRoot);
        _reportPath = Path.Combine(artifactRoot, "targeted-upright-crossref-report.txt");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", "1");

        try
        {
            Log("VisualTeX targeted acceptance: upright typography + Word native cross-reference font");
            Log($"Fixture: {fixturePath}");
            Log($"Artifacts: {artifactRoot}");
            var fixture = JsonSerializer.Deserialize<UprightFixtureRoot>(
                File.ReadAllText(fixturePath, Encoding.UTF8),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                ?? throw new InvalidDataException("The upright fixture is empty.");
            Assert(fixture.Cases.Count >= 10, "The upright fixture does not contain detailed coverage.");

            var wordAssembly = Assembly.Load("VisualTeX.WordVsto");
            var converterType = wordAssembly.GetType(
                "VisualTeX.WordVsto.WordOmmlConverter",
                throwOnError: true)!;

            ValidateWordOmmlUprightConversion(fixture, converterType, artifactRoot);
            if (skipNumberingFont)
            {
                Log("[Issue 2] SKIP - Word display-number font styling is outside the current task scope.");
            }
            else
            {
                var numberingType = wordAssembly.GetType(
                    "VisualTeX.WordVsto.WordEquationNumbering",
                    throwOnError: true)!;
                ValidateNativeCrossReferenceFormatting(
                    fixture,
                    numberingType,
                    converterType,
                    artifactRoot);
            }

            Log(skipNumberingFont
                ? "RESULT: PASS - upright OMML acceptance passed; numbering-font checks were explicitly skipped."
                : "RESULT: PASS - both targeted issues passed all detailed acceptance checks.");
            File.WriteAllLines(_reportPath, Report, new UTF8Encoding(false));
            return 0;
        }
        catch (Exception error)
        {
            Log($"RESULT: FAIL - {error}");
            try { File.WriteAllLines(_reportPath, Report, new UTF8Encoding(false)); } catch { }
            return 1;
        }
    }

    private static void ValidateWordOmmlUprightConversion(
        UprightFixtureRoot fixture,
        Type converterType,
        string artifactRoot)
    {
        Log("[Issue 1 / Stage 1] Converting every normalized MathML case to Word OMML...");
        var ommlDirectory = Path.Combine(artifactRoot, "omml-cases");
        Directory.CreateDirectory(ommlDirectory);
        var transform = FindStaticMethod(converterType, "TransformMathMlToOmml", 1);
        var checkedTokens = 0;

        for (var index = 0; index < fixture.Cases.Count; index++)
        {
            var item = fixture.Cases[index];
            var omml = (string)(Invoke(transform, null, item.MathMl)
                ?? throw new InvalidDataException($"{item.Name}: OMML conversion returned null."));
            foreach (var forbidden in fixture.ForbiddenInternalCommands)
            {
                Assert(
                    omml.IndexOf(forbidden, StringComparison.OrdinalIgnoreCase) < 0,
                    $"{item.Name}: Word OMML contains the literal internal command {forbidden}.");
            }

            var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
            XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var visibleText = string.Concat(
                document.Descendants(math + "t").Select(element => element.Value));
            Assert(
                visibleText.IndexOf("differentialD", StringComparison.OrdinalIgnoreCase) < 0,
                $"{item.Name}: Word-visible text contains differentialD.");
            Assert(
                visibleText.IndexOf("exponentialE", StringComparison.OrdinalIgnoreCase) < 0,
                $"{item.Name}: Word-visible text contains exponentialE.");
            Assert(
                visibleText.IndexOf("imaginaryI", StringComparison.OrdinalIgnoreCase) < 0,
                $"{item.Name}: Word-visible text contains imaginaryI.");
            Assert(
                visibleText.IndexOf("imaginaryJ", StringComparison.OrdinalIgnoreCase) < 0,
                $"{item.Name}: Word-visible text contains imaginaryJ.");

            var compactVisibleText = RemoveWhitespace(visibleText);
            var explicitNormalText = string.Concat(
                document
                    .Descendants(math + "r")
                    .Where(run => run.Element(math + "rPr")?.Element(math + "nor") is not null)
                    .SelectMany(run => run.Elements(math + "t"))
                    .Select(element => element.Value));
            var compactExplicitNormalText = RemoveWhitespace(explicitNormalText);
            XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
            var explicitNormalMathMlTokens = XDocument.Parse(item.MathMl)
                .Descendants(presentationMath + "mi")
                .Where(element =>
                {
                    var variant = element.Attribute("mathvariant")?.Value ?? string.Empty;
                    return variant.IndexOf("normal", StringComparison.OrdinalIgnoreCase) >= 0
                        || variant.IndexOf("upright", StringComparison.OrdinalIgnoreCase) >= 0;
                })
                .Select(element => RemoveWhitespace(element.Value))
                .Where(value => value.Length > 0)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var token in item.NormalMathMlTokens)
            {
                var compactToken = RemoveWhitespace(token);
                Assert(
                    compactVisibleText.IndexOf(compactToken, StringComparison.OrdinalIgnoreCase) >= 0,
                    $"{item.Name}: Word OMML no longer contains expected upright token '{token}'.");
                if (explicitNormalMathMlTokens.Contains(compactToken))
                {
                    Assert(
                        compactExplicitNormalText.IndexOf(compactToken, StringComparison.OrdinalIgnoreCase) >= 0,
                        $"{item.Name}: Word OMML token '{token}' is not marked with explicit m:nor normal style.");
                }
                checkedTokens++;
            }

            var safeName = string.Concat(
                item.Name.Select(character => char.IsLetterOrDigit(character) ? character : '_'));
            File.WriteAllText(
                Path.Combine(ommlDirectory, $"{index + 1:D2}-{safeName}.xml"),
                omml,
                new UTF8Encoding(false));
            Log($"  PASS {index + 1:D2}. {item.Name}");
            Log($"       normalized LaTeX: {item.Normalized}");
            Log($"       Word visible text: {visibleText}");
        }

        Log(
            $"[Issue 1 / Stage 1] PASS - {fixture.Cases.Count} formula classes, "
            + $"{checkedTokens} upright-token checks, and no MathLive internal command leaked into OMML.");
    }

    private static void ValidateNativeCrossReferenceFormatting(
        UprightFixtureRoot fixture,
        Type numberingType,
        Type converterType,
        string artifactRoot)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Table? table = null;
        Word.InlineShape? shape = null;
        WordRange? ommlRange = null;
        Word.Field? bodyReference = null;
        Word.Selection? selection = null;
        var imagePath = Path.Combine(artifactRoot, "target-formula.png");
        CreateFormulaImage(imagePath);

        try
        {
            Log("[Issue 2 / Stage 1] Starting a clean Word document and creating a numbered formula target...");
            application = new Word.Application
            {
                Visible = false,
                DisplayAlerts = WdAlertLevel.wdAlertsNone,
            };
            document = application.Documents.Add();
            table = CreateNumberedFormulaTable(document);
            shape = InsertFormulaPicture(document, table, imagePath);

            var formulaId = Guid.NewGuid().ToString("D");
            var metadata = new FormulaMetadata
            {
                FormulaId = formulaId,
                Title = "Targeted cross-reference formula",
                Latex = "x=y",
                Lines = new List<FormulaLine>
                {
                    new() { Id = Guid.NewGuid().ToString("D"), Latex = "x=y" },
                },
                DisplayMode = "block",
                Numbered = true,
                RenderWidthPx = 160,
                RenderHeightPx = 48,
                Baseline = 34,
                CreatedWithVersion = "targeted-acceptance",
                UpdatedWithVersion = "targeted-acceptance",
                CreatedAt = DateTime.UtcNow.ToString("O"),
                UpdatedAt = DateTime.UtcNow.ToString("O"),
            };
            var encodedMetadata = FormulaMetadataCodec.Encode(metadata);
            shape.Title = encodedMetadata;
            shape.AlternativeText = encodedMetadata;

            var reconcileFormula = FindStaticMethod(numberingType, "TryReconcileFormula", 4);
            Invoke(
                reconcileFormula,
                null,
                document,
                shape.Range,
                shape.Height,
                metadata);
            AssertHiddenSequenceRemainsOnePoint(document, "initial numbered formula insertion");
            Log("  PASS hidden native SEQ target is still 1 pt and white.");

            var getTargets = FindStaticMethod(numberingType, "GetEquationReferenceTargets", 1);
            var targets = (IEnumerable)(Invoke(getTargets, null, document)
                ?? throw new InvalidDataException("No equation-reference target collection was returned."));
            var target = targets.Cast<object>().SingleOrDefault()
                ?? throw new InvalidDataException("The numbered formula did not create one native reference target.");

            Log("[Issue 2 / Stage 2] Inserting a native cross-reference into 14 pt body text...");
            selection = application.Selection;
            selection.SetRange(document.Content.End - 1, document.Content.End - 1);
            selection.TypeParagraph();
            ApplyNormalBodyAppearance(selection.Range, BodyFontSize);
            selection.Font.Size = BodyFontSize;
            selection.Font.Hidden = 0;
            selection.Font.Color = WdColor.wdColorAutomatic;
            selection.TypeText("正文中的公式引用：");
            var referenceStart = selection.Start;

            var styleType = numberingType.Assembly.GetType(
                "VisualTeX.WordVsto.EquationReferenceStyle",
                throwOnError: true)!;
            var parenthesizedStyle = Enum.Parse(styleType, "Parenthesized");
            var insertReference = FindStaticMethod(numberingType, "InsertEquationReference", 4);
            Invoke(
                insertReference,
                null,
                document,
                selection,
                target,
                parenthesizedStyle);
            selection.TypeText("，用于字号专项验收。");

            bodyReference = FindBodyReferenceField(document, referenceStart);
            AssertVisibleReference(
                bodyReference,
                BodyFontSize,
                "normal native cross-reference insertion");
            AssertHiddenSequenceRemainsOnePoint(document, "normal reference insertion");
            LogReferenceDetails(bodyReference, "normal insertion");
            Log("  PASS inserted REF uses surrounding 14 pt body size, automatic color, visible text and CHARFORMAT.");

            Log("[Issue 2 / Stage 3] Forcing the REF back to 1 pt, then running the product refresh command...");
            DegradeReferenceToHiddenSequenceAppearance(bodyReference);
            Release(bodyReference);
            bodyReference = null;
            var updateReferences = FindStaticMethod(numberingType, "UpdateNativeCrossReferences", 1);
            Invoke(updateReferences, null, document);
            bodyReference = FindBodyReferenceField(document, referenceStart);
            AssertVisibleReference(
                bodyReference,
                BodyFontSize,
                "native reference refresh");
            AssertHiddenSequenceRemainsOnePoint(document, "native reference refresh");
            LogReferenceDetails(bodyReference, "after refresh");
            Log("  PASS refresh restores the body REF to 14 pt without enlarging the hidden SEQ target.");

            Log("[Issue 2 / Stage 4] Simulating picture/OLE → OMML format conversion through the exact production reconciliation entry...");
            DegradeReferenceToHiddenSequenceAppearance(bodyReference);
            Release(bodyReference);
            bodyReference = null;
            shape.Delete();
            Release(shape);
            shape = null;
            var insertion = GetCenterCellInsertionRange(table);
            try
            {
                var insertOmml = FindStaticMethod(converterType, "Insert", 7);
                ommlRange = (WordRange)(Invoke(
                    insertOmml,
                    null,
                    application,
                    document,
                    insertion,
                    fixture.Cases[0].MathMl,
                    true,
                    false,
                    false)
                    ?? throw new InvalidDataException("OMML conversion insertion returned no range."));
            }
            finally { Release(insertion); }
            metadata.Latex = fixture.Cases[0].Normalized;
            Invoke(
                reconcileFormula,
                null,
                document,
                ommlRange,
                24f,
                metadata);
            bodyReference = FindBodyReferenceField(document, referenceStart);
            AssertVisibleReference(
                bodyReference,
                BodyFontSize,
                "picture/OLE to OMML formula-format conversion");
            AssertHiddenSequenceRemainsOnePoint(document, "picture/OLE to OMML conversion");
            LogReferenceDetails(bodyReference, "after conversion to OMML");
            Log("  PASS OMML conversion reconciliation restores the existing body REF to 14 pt.");

            Log("[Issue 2 / Stage 5] Simulating OMML → picture/OLE format conversion through the same production entry...");
            DegradeReferenceToHiddenSequenceAppearance(bodyReference);
            Release(bodyReference);
            bodyReference = null;
            ommlRange.Delete();
            Release(ommlRange);
            ommlRange = null;
            shape = InsertFormulaPicture(document, table, imagePath);
            shape.Title = encodedMetadata;
            shape.AlternativeText = encodedMetadata;
            Invoke(
                reconcileFormula,
                null,
                document,
                shape.Range,
                shape.Height,
                metadata);
            bodyReference = FindBodyReferenceField(document, referenceStart);
            AssertVisibleReference(
                bodyReference,
                BodyFontSize,
                "OMML to picture/OLE formula-format conversion");
            AssertHiddenSequenceRemainsOnePoint(document, "OMML to picture/OLE conversion");
            LogReferenceDetails(bodyReference, "after conversion back to picture/OLE");
            Log("  PASS reverse conversion reconciliation restores the existing body REF to 14 pt.");

            var documentPath = Path.Combine(
                artifactRoot,
                "VisualTeX-Targeted-Upright-CrossReference.docx");
            document.SaveAs2(documentPath, WdSaveFormat.wdFormatXMLDocument);
            Log($"[Issue 2] Saved inspectable Word artifact: {documentPath}");
        }
        finally
        {
            Release(bodyReference);
            Release(ommlRange);
            Release(shape);
            Release(table);
            Release(selection);
            if (document is not null)
            {
                try { document.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            if (application is not null)
            {
                try { application.Quit(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(application);
            try { File.Delete(imagePath); } catch { }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static Word.Table CreateNumberedFormulaTable(Word.Document document)
    {
        object start = 0;
        object end = 0;
        var anchor = document.Range(ref start, ref end);
        try
        {
            var table = document.Tables.Add(anchor, 1, 3);
            table.Borders.Enable = 0;
            table.AllowAutoFit = false;
            table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 100f;
            table.Columns[1].PreferredWidth = 20f;
            table.Columns[2].PreferredWidth = 60f;
            table.Columns[3].PreferredWidth = 20f;
            return table;
        }
        finally { Release(anchor); }
    }

    private static Word.InlineShape InsertFormulaPicture(
        Word.Document document,
        Word.Table table,
        string imagePath)
    {
        var insertion = GetCenterCellInsertionRange(table);
        try
        {
            object link = false;
            object save = true;
            object rangeObject = insertion;
            var shape = document.InlineShapes.AddPicture(
                imagePath,
                ref link,
                ref save,
                ref rangeObject);
            shape.Width = 120f;
            shape.Height = 24f;
            return shape;
        }
        finally { Release(insertion); }
    }

    private static WordRange GetCenterCellInsertionRange(Word.Table table)
    {
        Word.Cell? cell = null;
        WordRange? cellRange = null;
        WordRange? editable = null;
        try
        {
            cell = table.Cell(1, 2);
            cellRange = cell.Range;
            editable = cellRange.Duplicate;
            editable.End = Math.Max(editable.Start, editable.End - 1);
            editable.Text = string.Empty;
            editable.Collapse(WdCollapseDirection.wdCollapseStart);
            return editable;
        }
        finally
        {
            Release(cellRange);
            Release(cell);
        }
    }

    private static Word.Field FindBodyReferenceField(
        Word.Document document,
        int minimumStart)
    {
        Word.Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? field = null;
                WordRange? result = null;
                try
                {
                    field = fields[index];
                    WordRange? code = null;
                    try
                    {
                        code = field.Code;
                        var codeText = code.Text ?? string.Empty;
                        if (codeText.IndexOf("REF ", StringComparison.OrdinalIgnoreCase) < 0)
                            continue;
                    }
                    finally { Release(code); }
                    result = field.Result;
                    var inTable = (bool)result.get_Information(WdInformation.wdWithInTable);
                    // Updating a field code from MERGEFORMAT to CHARFORMAT changes
                    // Word's hidden field-code character count, so absolute Range
                    // positions can move even though the visible REF remains in the
                    // same body paragraph. The generated formula-number REF is in
                    // the numbering table; the native body cross-reference is the
                    // only REF outside that table.
                    if (inTable) continue;
                    var selected = field;
                    field = null;
                    return selected;
                }
                finally
                {
                    Release(result);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
        LogFieldInventory(document, "body REF lookup failure");
        throw new InvalidDataException("No body native REF field was found.");
    }

    private static void LogFieldInventory(Word.Document document, string stage)
    {
        Word.Fields? fields = null;
        try
        {
            fields = document.Fields;
            Log($"       field inventory ({stage}): count={fields.Count}");
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? field = null;
                WordRange? code = null;
                WordRange? result = null;
                Word.Font? font = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    result = field.Result;
                    font = result.Font;
                    var inTable = false;
                    try
                    {
                        inTable = (bool)result.get_Information(WdInformation.wdWithInTable);
                    }
                    catch { }
                    Log(
                        $"         [{index}] type={field.Type}, start={result.Start}, "
                        + $"inTable={inTable}, size={font.Size:F2}, hidden={font.Hidden}, "
                        + $"code='{(code.Text ?? string.Empty).Trim()}', result='{result.Text}'");
                }
                finally
                {
                    Release(font);
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
    }

    private static void AssertVisibleReference(
        Word.Field field,
        float expectedSize,
        string stage)
    {
        WordRange? result = null;
        WordRange? code = null;
        Word.Font? resultFont = null;
        Word.Font? codeFont = null;
        try
        {
            result = field.Result;
            code = field.Code;
            resultFont = result.Font;
            codeFont = code.Font;
            Assert(
                Math.Abs(resultFont.Size - expectedSize) <= 0.25f,
                $"{stage}: REF result is {resultFont.Size:F2} pt, expected {expectedSize:F2} pt.");
            Assert(resultFont.Hidden == 0, $"{stage}: REF result is hidden.");
            Assert(
                resultFont.Color == WdColor.wdColorAutomatic,
                $"{stage}: REF result color is {resultFont.Color}, expected automatic.");
            Assert(resultFont.Position == 0, $"{stage}: REF baseline position is not zero.");
            Assert(
                Math.Abs(codeFont.Size - expectedSize) <= 0.25f,
                $"{stage}: REF field code is {codeFont.Size:F2} pt, expected {expectedSize:F2} pt.");
            var codeText = code.Text ?? string.Empty;
            Assert(
                codeText.IndexOf("CHARFORMAT", StringComparison.OrdinalIgnoreCase) >= 0,
                $"{stage}: REF code does not use CHARFORMAT: {codeText}");
            Assert(
                codeText.IndexOf("MERGEFORMAT", StringComparison.OrdinalIgnoreCase) < 0,
                $"{stage}: REF code still uses MERGEFORMAT and can inherit the 1 pt SEQ style.");
            Assert(
                !string.IsNullOrWhiteSpace(result.Text),
                $"{stage}: REF result text is empty.");
        }
        finally
        {
            Release(codeFont);
            Release(resultFont);
            Release(code);
            Release(result);
        }
    }

    private static void AssertHiddenSequenceRemainsOnePoint(
        Word.Document document,
        string stage)
    {
        Word.Fields? fields = null;
        Word.Field? sequence = null;
        WordRange? code = null;
        WordRange? result = null;
        Word.Font? font = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? candidate = null;
                WordRange? candidateCode = null;
                try
                {
                    candidate = fields[index];
                    candidateCode = candidate.Code;
                    var text = candidateCode.Text ?? string.Empty;
                    if (text.IndexOf("SEQ ", StringComparison.OrdinalIgnoreCase) < 0) continue;
                    sequence = candidate;
                    candidate = null;
                    break;
                }
                finally
                {
                    Release(candidateCode);
                    Release(candidate);
                }
            }
            Assert(sequence is not null, $"{stage}: hidden native SEQ field is missing.");
            code = sequence!.Code;
            result = sequence.Result;
            font = result.Font;
            Assert(
                Math.Abs(font.Size - 1f) <= 0.1f,
                $"{stage}: hidden SEQ changed to {font.Size:F2} pt instead of remaining 1 pt.");
            Assert(
                font.Color == WdColor.wdColorWhite,
                $"{stage}: hidden SEQ color changed from white.");
        }
        finally
        {
            Release(font);
            Release(result);
            Release(code);
            Release(sequence);
            Release(fields);
        }
    }

    private static void DegradeReferenceToHiddenSequenceAppearance(Word.Field field)
    {
        WordRange? result = null;
        WordRange? code = null;
        Word.Font? resultFont = null;
        Word.Font? codeFont = null;
        try
        {
            result = field.Result;
            code = field.Code;
            resultFont = result.Font;
            codeFont = code.Font;
            resultFont.Size = 1f;
            resultFont.Hidden = 1;
            resultFont.Color = WdColor.wdColorWhite;
            codeFont.Size = 1f;
            codeFont.Hidden = 1;
            codeFont.Color = WdColor.wdColorWhite;
        }
        finally
        {
            Release(codeFont);
            Release(resultFont);
            Release(code);
            Release(result);
        }
    }

    private static void LogReferenceDetails(Word.Field field, string stage)
    {
        WordRange? result = null;
        WordRange? code = null;
        Word.Font? resultFont = null;
        Word.Font? codeFont = null;
        try
        {
            result = field.Result;
            code = field.Code;
            resultFont = result.Font;
            codeFont = code.Font;
            Log(
                $"       {stage}: result='{result.Text}', resultSize={resultFont.Size:F2} pt, "
                + $"hidden={resultFont.Hidden}, color={resultFont.Color}, position={resultFont.Position}");
            Log(
                $"       {stage}: codeSize={codeFont.Size:F2} pt, code='{(code.Text ?? string.Empty).Trim()}'");
        }
        finally
        {
            Release(codeFont);
            Release(resultFont);
            Release(code);
            Release(result);
        }
    }

    private static void ApplyNormalBodyAppearance(WordRange range, float size)
    {
        Word.Font? font = null;
        Word.ParagraphFormat? format = null;
        try
        {
            object normalStyle = WdBuiltinStyle.wdStyleNormal;
            try { range.set_Style(ref normalStyle); } catch { }
            font = range.Font;
            font.Size = size;
            font.Hidden = 0;
            font.Color = WdColor.wdColorAutomatic;
            font.Position = 0;
            format = range.ParagraphFormat;
            format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
        }
        finally
        {
            Release(format);
            Release(font);
        }
    }

    private static void CreateFormulaImage(string path)
    {
        using var bitmap = new Bitmap(160, 48, PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.Transparent);
        using var pen = new Pen(Color.Black, 3f);
        graphics.DrawLine(pen, 8, 24, 152, 24);
        graphics.DrawEllipse(pen, 62, 8, 36, 30);
        bitmap.Save(path, ImageFormat.Png);
    }

    private static MethodInfo FindStaticMethod(Type type, string name, int parameterCount)
    {
        var methods = type.GetMethods(
                BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)
            .Where(method => method.Name == name
                && method.GetParameters().Length == parameterCount)
            .ToArray();
        if (methods.Length != 1)
            throw new MissingMethodException(
                type.FullName,
                $"{name} with {parameterCount} parameters (found {methods.Length})");
        return methods[0];
    }

    private static object? Invoke(MethodInfo method, object? instance, params object?[] arguments)
    {
        try
        {
            return method.Invoke(instance, arguments);
        }
        catch (TargetInvocationException error) when (error.InnerException is not null)
        {
            throw error.InnerException;
        }
    }

    private static string RemoveWhitespace(string value) =>
        string.Concat(value.Where(character => !char.IsWhiteSpace(character)));

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidDataException(message);
    }

    private static void Log(string message)
    {
        Console.WriteLine(message);
        Report.Add(message);
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.FinalReleaseComObject(value); } catch { }
    }
}
