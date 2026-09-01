using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlNativeNumberToggleAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-native-number-toggle.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.Bookmark? formulaBookmark = null;
        Word.Range? formulaRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Paragraphs? unnumberedParagraphs = null;
        Word.Paragraph? unnumberedParagraph = null;
        Word.ParagraphFormat? unnumberedParagraphFormat = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            application.Selection.SetRange(insertion.Start, insertion.End);
            var createSession = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertion.Start,
                insertion.End,
                latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            service.InsertOmml(createSession, QuadraticFormulaMathMl());
            Release(insertion); insertion = null;

            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                "native number-toggle initial numbered host",
                updateReference: true);

            var numberedMetadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "Native number-toggle acceptance lost initial metadata.");
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    "Native number-toggle acceptance lost initial VTOMML bookmark.");
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            var numberedStart = formulaRange.Start;
            var numberedEnd = formulaRange.End;
            Release(formulaRange); formulaRange = null;
            Release(formulaBookmark); formulaBookmark = null;
            var unnumberedSession = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                numberedStart,
                numberedEnd,
                latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: numberedMetadata);
            unnumberedSession.Numbered = false;
            service.ReplaceOmml(unnumberedSession, QuadraticFormulaMathMl());

            var unnumberedMetadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "Native number-toggle acceptance lost unnumbered metadata.");
            AssertTrue(!unnumberedMetadata.Numbered,
                "Numbered→unnumbered OMML did not persist Numbered=false.");
            AssertTrue(
                !document.Bookmarks.Exists(
                    WordEquationNumbering.NativeNumberBookmarkName(formulaId)),
                "Numbered→unnumbered OMML retained VTEqNum inside the formula.");
            AssertTrue(
                !document.Bookmarks.Exists(
                    WordEquationNumbering.EquationBookmarkName(formulaId)),
                "Numbered→unnumbered OMML retained the visible-number alias.");
            AssertTrue(
                !document.Bookmarks.Exists(
                    WordEquationNumbering.NativeCaptionBookmarkName(formulaId)),
                "Numbered→unnumbered OMML retained the caption alias.");

            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    "Numbered→unnumbered OMML lost its VTOMML identity.");
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            maths = formulaRange.OMaths;
            AssertEqual(1, maths.Count,
                "Numbered→unnumbered OMML no longer contains one native equation.");
            math = maths[1];
            AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                "Numbered→unnumbered OMML degraded from wdOMathDisplay.");
            AssertEqual(0, formulaRange.Fields.Count,
                "Numbered→unnumbered OMML retained a mathematical SEQ field.");
            AssertTrue(
                !WordOmmlConverter.HasVisualTeXNativeEquationNumber(
                    formulaRange.WordOpenXML ?? string.Empty),
                "Numbered→unnumbered OMML retained the generated #(...) wrapper.");
            AssertEqual(0, document.Shapes.Count,
                "Numbered→unnumbered OMML created a floating Shape.");
            AssertEqual(0, document.Tables.Count,
                "Numbered→unnumbered OMML created a Word table.");
            unnumberedParagraphs = formulaRange.Paragraphs;
            AssertEqual(1, unnumberedParagraphs.Count,
                "Numbered→unnumbered OMML no longer occupies one standalone paragraph.");
            unnumberedParagraph = unnumberedParagraphs[1];
            unnumberedParagraphFormat = unnumberedParagraph.Format;
            AssertTrue(
                !(unnumberedParagraphFormat.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceExactly
                  && unnumberedParagraphFormat.LineSpacing <= 2.01f),
                $"Numbered→unnumbered OMML inherited VisualTeX's compact structural line box: rule={unnumberedParagraphFormat.LineSpacingRule}, line={unnumberedParagraphFormat.LineSpacing:0.##}pt.");
            AssertEqual(Word.WdLineSpacing.wdLineSpaceSingle,
                unnumberedParagraphFormat.LineSpacingRule,
                "Numbered→unnumbered OMML did not restore ordinary single line spacing after dismantling its 1x3 host.");
            Release(unnumberedParagraphFormat); unnumberedParagraphFormat = null;
            Release(unnumberedParagraph); unnumberedParagraph = null;
            Release(unnumberedParagraphs); unnumberedParagraphs = null;

            var unnumberedStart = formulaRange.Start;
            var unnumberedEnd = formulaRange.End;
            Release(math); math = null;
            Release(maths); maths = null;
            Release(formulaRange); formulaRange = null;
            Release(formulaBookmark); formulaBookmark = null;
            var renumberSession = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                unnumberedStart,
                unnumberedEnd,
                latex: @"x = \frac{-b \pm \sqrt{b^2-4ac}}{2a}",
                originalMetadata: unnumberedMetadata);
            // Reproduce the real Office editor transition where legacy metadata
            // says codeFormat="latex", while
            // the editor internally normalizes the equivalent source mode to
            // "raw" and may normalize harmless LaTeX spacing. The MathML is still
            // semantically identical, so renumbering must keep the existing live
            // OMath instead of rebuilding it.
            renumberSession.CodeFormat = "raw";
            service.ReplaceOmml(renumberSession, QuadraticFormulaMathMl());

            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                "native number-toggle renumbered host",
                updateReference: true);
            var renumberedMetadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "Native number-toggle acceptance lost renumbered metadata.");
            AssertTrue(renumberedMetadata.Numbered,
                "Unnumbered→numbered OMML did not persist Numbered=true.");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                "native number-toggle save/reopened host",
                updateReference: true);

            Console.WriteLine(
                "Native OMML number-toggle acceptance passed: numbered #(SEQ) was atomically stripped to pure wdOMathDisplay and rebuilt with the same FormulaId, without Shape/Table artifacts, and survived save/reopen.");
        }
        finally
        {
            Release(unnumberedParagraphFormat);
            Release(unnumberedParagraph);
            Release(unnumberedParagraphs);
            Release(math);
            Release(maths);
            Release(formulaRange);
            Release(formulaBookmark);
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
}
