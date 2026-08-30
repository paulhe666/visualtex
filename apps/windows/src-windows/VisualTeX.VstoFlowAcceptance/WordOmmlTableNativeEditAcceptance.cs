using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlTableNativeEditAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The native-edit 1x3 acceptance refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-1x3-native-edit.docx");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Range? editToken = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            ConfigureOmmlTableNumberPage(document);
            document.SaveAs2(
                documentPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var insertion = document.Content.End - 1;
            application.Selection.SetRange(insertion, insertion);
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertion,
                insertion,
                latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            service.InsertOmml(session, QuadraticFormulaMathMl());

            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "01-before-native-edit");

            ResolveNativeEditMath(
                document,
                formulaId,
                out bookmark,
                out equationRange,
                out maths,
                out math,
                out mathRange);
            math.Linearize();
            Release(mathRange); mathRange = null;
            mathRange = math.Range.Duplicate;
            AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                "Linearize changed the center formula from wdOMathDisplay.");
            AssertTrue((bool)mathRange.get_Information(Word.WdInformation.wdWithInTable),
                "Linearize moved the formula outside the managed 1x3 table.");
            AssertEqual(0, mathRange.Fields.Count,
                "Linearize leaked the number SEQ into the center OMath.");
            Console.WriteLine(
                $"  native Linearize text='{NormalizeNativeEditText(mathRange.Text)}', range={mathRange.Start}:{mathRange.End}.");

            math.BuildUp();
            Release(mathRange); mathRange = null;
            Release(math); math = null;
            Release(maths); maths = null;
            Release(equationRange); equationRange = null;
            Release(bookmark); bookmark = null;
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "02-after-linearize-buildup");

            // Perform a genuine Word-native edit in linear math: replace only the
            // first visible x token in the center OMath, then ask Word to BuildUp.
            // The direct SEQ/right-tab cell must remain entirely outside the math.
            ResolveNativeEditMath(
                document,
                formulaId,
                out bookmark,
                out equationRange,
                out maths,
                out math,
                out mathRange);
            math.Linearize();
            Release(mathRange); mathRange = null;
            mathRange = math.Range.Duplicate;
            var linearText = mathRange.Text ?? string.Empty;
            const string MathematicalItalicX = "\U0001D465";
            var xOffset = linearText.IndexOf(
                MathematicalItalicX,
                StringComparison.Ordinal);
            var xLength = MathematicalItalicX.Length;
            if (xOffset < 0)
            {
                xOffset = linearText.IndexOf('x');
                xLength = 1;
            }
            AssertTrue(xOffset >= 0,
                "The linearized center formula does not expose the expected x/𝑥 token.");
            editToken = document.Range(
                mathRange.Start + xOffset,
                mathRange.Start + xOffset + xLength);
            editToken.Text = "y";
            Release(editToken); editToken = null;
            try { math.BuildUp(); }
            catch
            {
                // Reacquire the live OMath after Word replaced the edited linear run.
                Release(mathRange); mathRange = null;
                Release(math); math = null;
                Release(maths); maths = null;
                maths = equationRange.OMaths;
                AssertEqual(1, maths.Count,
                    "Native text edit destroyed the center OMath before BuildUp.");
                math = maths[1];
                math.BuildUp();
            }

            Release(mathRange); mathRange = null;
            Release(math); math = null;
            Release(maths); maths = null;
            Release(equationRange); equationRange = null;
            Release(bookmark); bookmark = null;

            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "03-after-native-content-edit");
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "Native Word edit lost VisualTeX OMML metadata.");
            equationRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                formulaId,
                metadata);
            var editedXml = equationRange.WordOpenXML ?? string.Empty;
            AssertTrue(editedXml.IndexOf("y", StringComparison.OrdinalIgnoreCase) >= 0,
                "The Word-native x→y edit did not survive BuildUp.");
            AssertEqual(0, equationRange.Fields.Count,
                "Native edit pulled the right-cell SEQ into the OMath.");

            document.Fields.Update();
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "04-after-native-edit-f9");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            Release(equationRange); equationRange = null;

            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "05-after-native-edit-reopen");
            metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "Save/reopen lost metadata after Word-native editing.");
            equationRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                formulaId,
                metadata);
            AssertTrue((equationRange.WordOpenXML ?? string.Empty)
                    .IndexOf("y", StringComparison.OrdinalIgnoreCase) >= 0,
                "Save/reopen lost the Word-native x→y content edit.");

            Console.WriteLine(
                "Word OMML 1x3 native-edit acceptance passed: Linearize/BuildUp and a direct Word-native x→y edit stayed inside the center wdOMathDisplay, the right-cell direct SEQ/right Tab never entered math, FormulaId recovered through VTEqNum→1x3 identity, and F9/save/reopen remained stable.");
        }
        finally
        {
            Release(editToken);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(equationRange);
            Release(bookmark);
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

    private static void ResolveNativeEditMath(
        Word.Document document,
        string formulaId,
        out Word.Bookmark bookmark,
        out Word.Range equationRange,
        out Word.OMaths maths,
        out Word.OMath math,
        out Word.Range mathRange)
    {
        var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
            ?? throw new InvalidDataException(
                "Native-edit acceptance cannot resolve formula metadata.");
        bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
            ?? throw new InvalidDataException(
                "Native-edit acceptance cannot resolve VTOMML identity.");
        equationRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
            document,
            formulaId,
            metadata);
        maths = equationRange.OMaths;
        AssertEqual(1, maths.Count,
            "Native-edit center cell does not contain exactly one OMath.");
        math = maths[1];
        AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
            "Native-edit center OMath is not wdOMathDisplay.");
        mathRange = math.Range.Duplicate;
    }

    private static string NormalizeNativeEditText(string? text) =>
        (text ?? string.Empty)
            .Replace("\r", "<P>")
            .Replace("\a", "<CELL>")
            .Replace("\t", "<TAB>")
            .Replace("\v", "<BR>");
}
