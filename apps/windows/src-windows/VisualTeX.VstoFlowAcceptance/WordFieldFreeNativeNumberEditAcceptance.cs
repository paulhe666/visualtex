using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordFieldFreeNativeNumberEditAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-field-free-native-number-edit.docx");
        TryDeleteAcceptanceFile(documentPath);

        const string editedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>y</mi><mo>+</mo><mn>2</mn></mrow></math>";

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? source = null;
        Word.Range? added = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? exact = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Content.Text = "x+1#(2)";
            source = document.Range(0, 7);
            added = document.OMaths.Add(source);
            maths = added.OMaths;
            AssertEqual(
                1,
                maths.Count,
                "Word did not create exactly one native #(2) OMath.");
            math = maths[1];
            math.Type = Word.WdOMathType.wdOMathDisplay;
            math.BuildUp();
            exact = math.Range.Duplicate;

            AssertEqual(
                0,
                exact.Fields.Count,
                "The Word-native #(2) fixture unexpectedly contains fields.");
            AssertTrue(
                WordOmmlConverter.HasWordNativeEquationNumberHost(
                    exact.WordOpenXML ?? string.Empty),
                "Word did not normalize #(2) to its native equation-number host.");

            var originalNumberXml =
                ReadNativeNumberDelimiterXml(
                    exact.WordOpenXML ?? string.Empty);
            exact.Select();

            var service = new WordFormulaService(application);
            var selected = service.ReadSelection();
            AssertTrue(
                selected.Metadata is not null,
                "VisualTeX did not expose field-free native OMML to the editor.");
            AssertTrue(
                selected.Metadata!.Numbered == false,
                "Field-free Word numbering was incorrectly adopted into VisualTeX numbering.");
            AssertTrue(
                selected.Metadata.Latex.IndexOf(
                    "#",
                    StringComparison.Ordinal) < 0,
                "The editor source still contains the Word-native number separator.");
            AssertTrue(
                selected.Metadata.Latex.IndexOf(
                    "2)",
                    StringComparison.Ordinal) < 0,
                "The editor source still contains the Word-native number payload.");

            var editSession = CreateOmmlMathTypeAcceptanceSession(
                editedMathMl,
                "block",
                numbered: false,
                FormulaOleContract.WordOmmlMode);
            editSession.Mode = "edit";
            editSession.FormulaId = selected.FormulaId!;
            editSession.SourceDocumentId = selected.DocumentId;
            editSession.SourceObjectId = selected.ObjectId;
            editSession.OriginalMetadata = selected.Metadata;

            _ = service.ReplaceOmml(
                editSession,
                editedMathMl);

            Release(exact); exact = null;
            Release(math); math = null;
            Release(maths); maths = null;
            Release(added); added = null;
            Release(source); source = null;

            maths = document.OMaths;
            AssertEqual(
                1,
                maths.Count,
                "Editing field-free numbered OMML changed the OMath count.");
            math = maths[1];
            exact = math.Range.Duplicate;

            AssertEqual(
                Word.WdOMathType.wdOMathDisplay,
                math.Type,
                "Editing field-free numbered OMML changed display mode.");
            AssertEqual(
                0,
                exact.Fields.Count,
                "Editing field-free numbered OMML introduced a field.");
            AssertTrue(
                WordOmmlConverter.HasWordNativeEquationNumberHost(
                    exact.WordOpenXML ?? string.Empty),
                "Editing field-free numbered OMML removed the native #(2) host.");
            AssertEqual(
                originalNumberXml,
                ReadNativeNumberDelimiterXml(
                    exact.WordOpenXML ?? string.Empty),
                "Editing the formula changed Word's field-free native number payload.");

            AssertEditedSemanticBody(
                exact.WordOpenXML ?? string.Empty,
                editedMathMl,
                "after edit");

            document.SaveAs2(
                documentPath,
                Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(
                Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();

            Release(exact); exact = null;
            Release(math); math = null;
            Release(maths); maths = null;
            maths = document.OMaths;
            AssertEqual(
                1,
                maths.Count,
                "Save/reopen changed the field-free numbered OMath count.");
            math = maths[1];
            exact = math.Range.Duplicate;
            AssertEqual(
                0,
                exact.Fields.Count,
                "Save/reopen introduced a field into the native #(2) host.");
            AssertEqual(
                originalNumberXml,
                ReadNativeNumberDelimiterXml(
                    exact.WordOpenXML ?? string.Empty),
                "Save/reopen changed Word's field-free native number payload.");
            AssertEditedSemanticBody(
                exact.WordOpenXML ?? string.Empty,
                editedMathMl,
                "after save/reopen");

            exact.Select();
            var reread = service.ReadSelection();
            AssertTrue(
                reread.Metadata is not null
                && reread.Metadata.Latex.IndexOf(
                    "#",
                    StringComparison.Ordinal) < 0,
                "The editor exposed the preserved native number after save/reopen.");

            Console.WriteLine(
                "Field-free Word-native #(2) edit acceptance passed: number stayed outside VisualTeX numbering, editor hid it, formula body changed, and the original number payload survived save/reopen.");
        }
        finally
        {
            Release(exact);
            Release(math);
            Release(maths);
            Release(added);
            Release(source);
            if (document is not null)
            {
                try
                {
                    document.Close(
                        Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(document);
            try
            {
                QuitWordApplicationIfOwned(application);
            }
            catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static string ReadNativeNumberDelimiterXml(
        string wordOpenXml)
    {
        var equation = XElement.Parse(
            WordOmmlConverter.ExtractSingleOMath(wordOpenXml),
            LoadOptions.PreserveWhitespace);
        XNamespace math =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var delimiter = equation
            .Descendants(math + "d")
            .Single();
        return delimiter.ToString(
            SaveOptions.DisableFormatting);
    }

    private static void AssertEditedSemanticBody(
        string wordOpenXml,
        string expectedMathMl,
        string stage)
    {
        var semanticOmml =
            WordOmmlConverter.StripWordNativeEquationNumberHost(
                wordOpenXml);
        var actualMathMl =
            WordOmmlConverter.TransformOmmlToMathMl(
                semanticOmml,
                display: true);
        AssertEqual(
            MathTypeMtefCodec.SemanticSignature(expectedMathMl),
            MathTypeMtefCodec.SemanticSignature(actualMathMl),
            $"Field-free native OMML semantic body mismatch {stage}.");
    }
}
