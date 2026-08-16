using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeNativeNumberReferenceAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourceDocx = ResolveMathTypeNativeEditorFixture();
        if (!File.Exists(sourceDocx))
            throw new FileNotFoundException(
                "Genuine MathType Equation.DSMT4 fixture is missing.",
                sourceDocx);

        var targetPath = Path.Combine(
            artifactRoot,
            "MathType7-Genuine-OLE-Native-Number-Reference.docx");
        File.Copy(sourceDocx, targetPath, overwrite: true);

        var svgPath = Path.Combine(artifactRoot, "mathtype-number-template-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"160\" height=\"56\" viewBox=\"0 0 160 56\"><text x=\"4\" y=\"40\" font-family=\"Times New Roman\" font-size=\"32\">a+b</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 160, 56);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? stagingDocument = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Range? tail = null;
        Word.Range? insertion = null;
        Word.Field? numberField = null;
        Word.Range? numberCode = null;
        Word.Range? numberResult = null;
        Word.Range? numberFullRange = null;
        Word.Selection? selection = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(targetPath, ReadOnly: false, Visible: false);
            document.Activate();
            AssertEqual(1, document.InlineShapes.Count,
                "Genuine MathType fixture must contain exactly one Equation.DSMT4 object.");
            shape = document.InlineShapes[1];
            AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                "Genuine MathType fixture is not Equation.DSMT4.");
            var genuineCompoundBefore = MathTypeOleStorage.CaptureCompoundFile(shape);
            var genuineLatex = MathMlToLatexConverter.Convert(
                    MathTypeOleStorage.ReadMathMl(genuineCompoundBefore))
                .Trim();
            AssertTrue(!string.IsNullOrWhiteSpace(genuineLatex),
                "Genuine MathType fixture has no readable MTEF source.");
            AssertEqual(0, WordEquationNumbering.GetEquationReferenceTargets(document).Count,
                "Genuine MathType fixture unexpectedly starts with VisualTeX equation-number metadata.");

            // Remove only the fixture's trailing explanatory prose.  The original
            // Equation.DSMT4 character and CFB are never replaced or rewritten.
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.End - 1 > shapeRange.End)
            {
                tail = document.Range(shapeRange.End, paragraphRange.End - 1);
                tail.Delete();
            }
            Release(tail); tail = null;
            Release(paragraphRange); paragraphRange = null;
            Release(paragraph); paragraph = null;
            Release(paragraphs); paragraphs = null;
            Release(shapeRange); shapeRange = null;

            // Obtain an authentic MTPlaceRef field tree from the same production
            // generator used by VisualTeX.  Only the numbering field is copied;
            // the target equation remains the genuine MathType-created OLE above.
            stagingDocument = application.Documents.Add(Visible: false);
            stagingDocument.Activate();
            insertion = stagingDocument.Range(0, 0);
            insertion.Select();
            var stagingService = new WordFormulaService(application);
            stagingService.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "right"),
                FirstNumberedMathMl,
                emfPath);
            Release(insertion); insertion = null;
            numberField = FindMathTypePlaceRefFieldForAcceptance(stagingDocument)
                ?? throw new InvalidDataException(
                    "Staging MathType numbered equation did not expose MTPlaceRef.");
            numberCode = numberField.Code;
            numberResult = numberField.Result;
            numberFullRange = stagingDocument.Range(
                Math.Max(stagingDocument.Content.Start, numberCode.Start - 1),
                Math.Min(stagingDocument.Content.End, numberResult.End + 1));
            numberFullRange.Copy();

            document.Activate();
            // Match MathType's own default numbered-document state so the copied
            // MTSec/MTEqn current-value fields evaluate exactly as they do natively.
            insertion = document.Range(0, 0);
            insertion.InsertXML(MathTypeWordOpenXml.CreateDefaultSectionBreakFlatOpc("1.1"));
            Release(insertion); insertion = null;
            Release(shape); shape = document.InlineShapes[1];
            shapeRange = shape.Range;
            insertion = document.Range(shapeRange.Start, shapeRange.Start);
            insertion.InsertBefore("\t");
            Release(insertion); insertion = null;
            Release(shapeRange); shapeRange = null;
            Release(shape); shape = document.InlineShapes[1];
            shapeRange = shape.Range;
            insertion = document.Range(shapeRange.End, shapeRange.End);
            insertion.InsertAfter("\t");
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertion.Paste();
            try { document.Fields.Update(); } catch { }

            Release(shapeRange); shapeRange = null;
            Release(shape); shape = document.InlineShapes[1];
            var genuineCompoundAfterNumbering = MathTypeOleStorage.CaptureCompoundFile(shape);
            AssertTrue(genuineCompoundBefore.SequenceEqual(genuineCompoundAfterNumbering),
                "Adding native MathType numbering rewrote the genuine MathType OLE payload.");
            AssertEqual(0, WordEquationNumbering.GetEquationReferenceTargets(document).Count,
                "A genuine MathType numbered equation was incorrectly classified as a VisualTeX reference target.");

            var mathTypeTargets = MathTypeEquationReferences.GetTargets(document);
            AssertEqual(1, mathTypeTargets.Count,
                "VisualTeX did not discover the genuine MathType OLE after native MTPlaceRef numbering was attached.");
            AssertEqual(EquationReferenceSource.MathType, mathTypeTargets[0].Source,
                "Genuine MathType equation was assigned the wrong reference source.");
            AssertEqual(
                "MathType 公式",
                mathTypeTargets[0].LatexPreview,
                "MathType reference discovery should not synchronously read the OLE/MTEF payload merely to populate the picker preview.");

            paragraph = document.Paragraphs.Add();
            paragraphRange = paragraph.Range;
            selection = application.Selection;
            selection.SetRange(paragraphRange.Start, paragraphRange.Start);
            selection.Font.Color = Word.WdColor.wdColorAutomatic;
            var preDialogInsertionColor = selection.Font.Color;
            // Reproduce the real Ribbon path: Word can expose GOTOBUTTON's
            // built-in red typing format while the modal reference picker owns
            // focus. Production must use the color captured before the dialog,
            // then normalize the final materialized field tree after nested REF.
            selection.Font.Color = Word.WdColor.wdColorRed;
            MathTypeEquationReferences.InsertReference(
                document,
                selection,
                mathTypeTargets[0],
                preDialogInsertionColor);
            AssertNativeMathTypeReference(document, mathTypeTargets[0].NumberText);
            AssertNativeMathTypeReferenceColor(
                document,
                Word.WdColor.wdColorAutomatic);
            AssertEqual(
                Word.WdColor.wdColorAutomatic,
                selection.Font.Color,
                "VisualTeX MathType reference insertion left Word's typing color changed.");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            Release(shape); shape = null;
            document = application.Documents.Open(targetPath, ReadOnly: false, Visible: false);
            document.Activate();
            AssertEqual(1, document.InlineShapes.Count,
                "Genuine MathType OLE count changed after numbering/reference save + reopen.");
            shape = document.InlineShapes[1];
            var reopenedCompound = MathTypeOleStorage.CaptureCompoundFile(shape);
            AssertTrue(genuineCompoundBefore.SequenceEqual(reopenedCompound),
                "Genuine MathType OLE payload changed after save + reopen.");
            AssertEqual(0, WordEquationNumbering.GetEquationReferenceTargets(document).Count,
                "Reopened genuine MathType equation became a VisualTeX reference target unexpectedly.");
            var reopenedTargets = MathTypeEquationReferences.GetTargets(document);
            AssertEqual(1, reopenedTargets.Count,
                "Genuine MathType reference target did not survive Word reopen.");
            AssertNativeMathTypeReference(document, mathTypeTargets[0].NumberText);
            AssertNativeMathTypeReferenceColor(
                document,
                Word.WdColor.wdColorAutomatic);
            Console.WriteLine(
                "[MathType native reference] Genuine MathType-created Equation.DSMT4 remained byte-identical, was discovered through MTPlaceRef without VisualTeX metadata, and accepted a native ZEqnNum/GOTOBUTTON/REF reference through save + reopen.");
        }
        finally
        {
            Release(selection);
            Release(numberFullRange);
            Release(numberResult);
            Release(numberCode);
            Release(numberField);
            Release(insertion);
            Release(tail);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
            Release(shape);
            if (stagingDocument is not null)
            {
                try { stagingDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(stagingDocument);
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

    private static void AssertNativeMathTypeReferenceColor(
        Word.Document document,
        Word.WdColor expectedColor)
    {
        Word.Fields? fields = null;
        Word.Field? outer = null;
        Word.Range? outerCode = null;
        Word.Fields? nestedFields = null;
        Word.Field? nested = null;
        Word.Range? nestedCode = null;
        Word.Range? nestedResult = null;
        Word.Font? outerFont = null;
        Word.Font? nestedCodeFont = null;
        Word.Font? nestedResultFont = null;
        Word.Style? outerStyle = null;
        Word.Style? nestedCodeStyle = null;
        Word.Style? nestedResultStyle = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(outerCode); outerCode = null;
                Release(outer); outer = fields[index];
                outerCode = outer.Code;
                if ((outerCode.Text ?? string.Empty).IndexOf(
                        "GOTOBUTTON ZEqnNum",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                outerFont = outerCode.Font;
                AssertEqual(expectedColor, outerFont.Color,
                    "VisualTeX MathType GOTOBUTTON field kept Word's default red field color.");
                try { outerStyle = outerCode.get_Style() as Word.Style; } catch { }
                AssertTrue(
                    !string.Equals(
                        outerStyle?.NameLocal,
                        "MTEquationSection",
                        StringComparison.OrdinalIgnoreCase),
                    "VisualTeX MathType GOTOBUTTON retained MathType's internal MTEquationSection character style.");
                nestedFields = outerCode.Fields;
                AssertEqual(1, nestedFields.Count,
                    "VisualTeX MathType reference must contain one nested REF field.");
                nested = nestedFields[1];
                nestedCode = nested.Code;
                nestedCodeFont = nestedCode.Font;
                AssertEqual(expectedColor, nestedCodeFont.Color,
                    "VisualTeX MathType REF field code has the wrong character color.");
                try { nestedCodeStyle = nestedCode.get_Style() as Word.Style; } catch { }
                AssertTrue(
                    !string.Equals(
                        nestedCodeStyle?.NameLocal,
                        "MTEquationSection",
                        StringComparison.OrdinalIgnoreCase),
                    "VisualTeX MathType REF code retained MathType's internal MTEquationSection character style.");
                nestedResult = nested.Result;
                nestedResultFont = nestedResult.Font;
                AssertEqual(expectedColor, nestedResultFont.Color,
                    "VisualTeX MathType REF result inherited Word's default red GOTOBUTTON color.");
                try { nestedResultStyle = nestedResult.get_Style() as Word.Style; } catch { }
                AssertTrue(
                    !string.Equals(
                        nestedResultStyle?.NameLocal,
                        "MTEquationSection",
                        StringComparison.OrdinalIgnoreCase),
                    "VisualTeX MathType REF result retained MathType's internal MTEquationSection character style.");
                return;
            }
            throw new InvalidDataException(
                "No MathType GOTOBUTTON reference was available for color validation.");
        }
        finally
        {
            Release(nestedResultStyle);
            Release(nestedCodeStyle);
            Release(outerStyle);
            Release(nestedResultFont);
            Release(nestedCodeFont);
            Release(outerFont);
            Release(nestedResult);
            Release(nestedCode);
            Release(nested);
            Release(nestedFields);
            Release(outerCode);
            Release(outer);
            Release(fields);
        }
    }

    private static Word.Field? FindMathTypePlaceRefFieldForAcceptance(Word.Document document)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                var result = field;
                field = null;
                return result;
            }
            return null;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }
}
