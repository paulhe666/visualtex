using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeNumberFormatAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "mathtype-number-format-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"180\" height=\"64\" viewBox=\"0 0 180 64\"><text x=\"4\" y=\"44\" font-family=\"Times New Roman\" font-size=\"36\">a+b</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 180, 64);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.Selection? selection = null;
        Word.Paragraph? referenceParagraph = null;
        Word.Range? referenceRange = null;
        try
        {
            application = CreateWordApplication(visible: false);
            AssertFreshHeadingMathTypeReferenceIsolation(application, emfPath);
            Console.WriteLine("[MathType number formats] fresh-heading isolation complete.");
            document = application.Documents.Add();
            document.Activate();
            document.Content.Text = "MathType numbering acceptance\r";
            var service = new WordFormulaService(application);

            // A format selected before the first MathType equation must govern the
            // first native MTPlaceRef rather than falling back to MathType's old
            // hard-coded section.equation default.
            AssertEqual(0, service.SetEquationNumberFormat(EquationNumberFormat.ContinuousId),
                "Setting a number format in a document with no numbered equations should update zero objects.");

            insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            insertion.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "right"),
                FirstNumberedMathMl,
                emfPath);
            Release(insertion); insertion = null;

            insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            insertion.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"c+d",
                    mathTypeNumberPosition: "right"),
                SecondNumberedMathMl,
                emfPath);
            Release(insertion); insertion = null;

            AssertMathTypeNumberTexts(document, "(1)", "(2)");

            var targets = MathTypeEquationReferences.GetTargets(document);
            AssertEqual(2, targets.Count,
                "MathType number-format acceptance did not discover both native targets.");
            referenceParagraph = document.Paragraphs.Add();
            referenceRange = referenceParagraph.Range;
            selection = application.Selection;
            selection.SetRange(referenceRange.Start, referenceRange.Start);
            selection.Font.Color = Word.WdColor.wdColorAutomatic;
            MathTypeEquationReferences.InsertReference(document, selection, targets[0]);
            AssertNativeMathTypeReference(document, "(1)");
            AssertNativeMathTypeReferenceColor(document, Word.WdColor.wdColorAutomatic);

            var formats = new[]
            {
                (EquationNumberFormat.ContinuousId, new[] { "(1)", "(2)" }),
                (EquationNumberFormat.Heading1DotId, new[] { "(1.1)", "(1.2)" }),
                (EquationNumberFormat.Heading1DashId, new[] { "(1-1)", "(1-2)" }),
                (EquationNumberFormat.Heading2DotId, new[] { "(1.1.1)", "(1.1.2)" }),
                (EquationNumberFormat.Heading2DashId, new[] { "(1.1-1)", "(1.1-2)" }),
            };

            foreach (var (formatId, expected) in formats)
            {
                Console.WriteLine($"[MathType number formats] applying {formatId}...");
                var changed = service.SetEquationNumberFormat(formatId);
                AssertEqual(2, changed,
                    $"MathType format '{formatId}' did not rewrite exactly two native MTPlaceRef fields.");
                AssertMathTypeNumberTexts(document, expected);
                AssertNativeMathTypeReference(document, expected[0]);
                AssertNativeMathTypeReferenceColor(document, Word.WdColor.wdColorAutomatic);

                var refreshed = service.UpdateEquationNumbers();
                AssertEqual(2, refreshed,
                    $"MathType update after format '{formatId}' did not report both numbered equations.");
                AssertMathTypeNumberTexts(document, expected);
                AssertNativeMathTypeReference(document, expected[0]);
                Console.WriteLine($"[MathType number formats] {formatId} complete.");
            }

            // New VisualTeX-created MathType numbers must inherit the nearest
            // rewritten native MTPlaceRef template.
            Console.WriteLine("[MathType number formats] inserting inherited third equation...");
            Release(referenceRange); referenceRange = null;
            Release(referenceParagraph); referenceParagraph = null;
            insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            insertion.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"e+f",
                    mathTypeNumberPosition: "right"),
                FirstNumberedMathMl,
                emfPath);
            Release(insertion); insertion = null;
            AssertMathTypeNumberTexts(document, "(1.1-1)", "(1.1-2)", "(1.1-3)");
            AssertNativeMathTypeReference(document, "(1.1-1)");
            Console.WriteLine("[MathType number formats] third equation complete; saving...");

            var path = Path.Combine(
                artifactRoot,
                "VisualTeX-MathType-Number-Formats.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine("[MathType number formats] SaveAs2 complete; closing...");
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Console.WriteLine("[MathType number formats] close complete; reopening...");
            Release(document); document = null;
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            Console.WriteLine("[MathType number formats] reopen complete; validating...");
            document.Activate();
            Console.WriteLine("[MathType number formats] validating reopened number texts...");
            AssertMathTypeNumberTexts(document, "(1.1-1)", "(1.1-2)", "(1.1-3)");
            Console.WriteLine("[MathType number formats] reopened number texts ok; validating reference text...");
            AssertNativeMathTypeReference(document, "(1.1-1)");
            Console.WriteLine("[MathType number formats] reopened reference text ok; validating reference style/color...");
            AssertNativeMathTypeReferenceColor(document, Word.WdColor.wdColorAutomatic);
            Console.WriteLine("[MathType number formats] reopened reference style/color ok.");

            Console.WriteLine(
                "[MathType number formats] Continuous, chapter dot/dash and chapter.section dot/dash presets rewrote native MTPlaceRef fields, preserved ZEqnNum references, survived Update, propagated to newly inserted MathType equations and survived save/reopen.");
        }
        finally
        {
            Release(referenceRange);
            Release(referenceParagraph);
            Release(selection);
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

    private static void AssertFreshHeadingMathTypeReferenceIsolation(
        Word.Application application,
        string emfPath)
    {
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.Selection? selection = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            var service = new WordFormulaService(application);
            AssertEqual(
                0,
                service.SetEquationNumberFormat(EquationNumberFormat.Heading2DashId),
                "Fresh heading-format document should not report existing numbered equations.");

            insertion = document.Range(document.Content.Start, document.Content.Start);
            insertion.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "right"),
                FirstNumberedMathMl,
                emfPath);
            Release(insertion); insertion = null;

            var sectionCodeStart = int.MaxValue;
            var placeRefCodeStart = int.MaxValue;
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                var text = code.Text ?? string.Empty;
                if (text.IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    sectionCodeStart = Math.Min(sectionCodeStart, code.Start);
                    paragraphs = code.Paragraphs;
                    AssertEqual(1, paragraphs.Count,
                        "MathType section state must stay in one isolated paragraph.");
                    paragraph = paragraphs[1];
                    paragraphRange = paragraph.Range;
                    AssertEqual(-1, paragraphRange.Font.Hidden,
                        "MathType section-state paragraph mark must remain hidden.");
                    Release(paragraphRange); paragraphRange = null;
                    Release(paragraph); paragraph = null;
                    Release(paragraphs); paragraphs = null;
                }
                else if (text.IndexOf(
                             "MACROBUTTON MTPlaceRef",
                             StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    placeRefCodeStart = Math.Min(placeRefCodeStart, code.Start);
                }
            }
            AssertTrue(sectionCodeStart != int.MaxValue,
                "Fresh heading-format MathType equation did not create native section state.");
            AssertTrue(placeRefCodeStart != int.MaxValue,
                "Fresh heading-format MathType equation did not create MTPlaceRef.");
            Console.WriteLine(
                $"[MathType fresh heading reference] sectionStart={sectionCodeStart}, placeRefStart={placeRefCodeStart}, paragraphs={document.Paragraphs.Count}");
            AssertTrue(sectionCodeStart < placeRefCodeStart,
                "MathType section reset was placed after the first numbered equation.");
            AssertMathTypeNumberTexts(document, "(1.1-1)");

            // Reproduce the ordinary user path: put the caret in a normal paragraph
            // after the formula and invoke Insert Reference.  Do not pre-create a
            // sanitized Paragraph object as the old acceptance did.
            selection = application.Selection;
            selection.SetRange(document.Content.End - 1, document.Content.End - 1);
            selection.TypeParagraph();
            selection.Font.Color = Word.WdColor.wdColorAutomatic;
            // Reproduce the legacy production bug exactly: the caret inherited
            // MathType's internal hidden/red character style, not merely a red
            // direct font color. New reference insertion must strip this style.
            object pollutedMathTypeStyle = "MTEquationSection";
            selection.Range.set_Style(ref pollutedMathTypeStyle);
            var targets = MathTypeEquationReferences.GetTargets(document);
            AssertEqual(1, targets.Count,
                "Fresh heading-format MathType equation was not referenceable.");
            MathTypeEquationReferences.InsertReference(
                document,
                selection,
                targets[0],
                Word.WdColor.wdColorAutomatic);
            AssertNativeMathTypeReference(document, "(1.1-1)");
            AssertNativeMathTypeReferenceColor(
                document,
                Word.WdColor.wdColorAutomatic);
            AssertEqual(
                Word.WdColor.wdColorAutomatic,
                selection.Font.Color,
                "Fresh MathType reference polluted the following typing color.");
            Console.WriteLine(
                "[MathType fresh heading reference] Native section state stayed in a hidden paragraph before the first equation; first number was (1.1-1) and the ordinary reference path remained black.");
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(code);
            Release(field);
            Release(fields);
            Release(selection);
            Release(insertion);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }

    private static void AssertMathTypeNumberTexts(
        Word.Document document,
        params string[] expected)
    {
        var targets = MathTypeEquationReferences.GetTargets(document);
        AssertEqual(expected.Length, targets.Count,
            "MathType numbered-equation target count changed unexpectedly.");
        for (var index = 0; index < expected.Length; index++)
        {
            AssertEqual(
                expected[index],
                targets[index].NumberText.Trim(),
                $"MathType equation {index + 1} has the wrong visible number text.");
        }
    }
}
