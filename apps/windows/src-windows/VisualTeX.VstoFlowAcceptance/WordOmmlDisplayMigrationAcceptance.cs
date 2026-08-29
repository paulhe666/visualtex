using System.Xml.Linq;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlDisplayMigrationAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "visualtex-omml-bad-eqarr-migration.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
        Word.Range? badRange = null;
        Word.Shapes? shapeInventory = null;
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
            var session = CreateOmmlNumberingSession(
                document,
                formulaId,
                mode: "create",
                sourceObjectId: WordRangeReference(insertion.Start, insertion.End),
                latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            service.InsertOmml(session, QuadraticMathMl("x"));
            Release(insertion); insertion = null;

            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context: "healthy numbered OMML before migration fixture",
                updateReference: true);

            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "Migration fixture lost its VisualTeX OMML metadata.");
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    "Migration fixture lost its VisualTeX OMML bookmark.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var semanticOmml = WordOmmlConverter.ExtractSingleOMath(
                equationRange.WordOpenXML);

            // Reproduce the broken 22:09 build without touching a user document:
            // it embedded '#(number)' into m:eqArr and therefore turned numbering
            // into mathematical content/full-width Word equation-array UI.
            WordEquationNumbering.RemoveVisibleEquationNumberForFormula(
                document,
                formulaId);
            shapeInventory = document.Shapes;
            Console.WriteLine(
                $"  bad-eqArr fixture after removing healthy visible number: shapes={shapeInventory.Count}.");
            AssertEqual(0, shapeInventory.Count,
                "The healthy external number Shape was not removed before materializing the bad m:eqArr fixture.");
            Release(shapeInventory); shapeInventory = null;
            var badOmml = BuildLegacyBadEqArrOmml(
                semanticOmml,
                formulaId);
            string mathFontName;
            try { mathFontName = document.OMathFontName ?? string.Empty; }
            catch { mathFontName = string.Empty; }
            badRange = WordOmmlConverter.ReplaceWithPreparedOmml(
                application,
                document,
                equationRange,
                badOmml,
                display: true,
                mathFontName);
            Release(equationRange); equationRange = null;
            Release(bookmark); bookmark = null;
            bookmark = WordOmmlFormulaStore.Wrap(
                document,
                badRange,
                metadata,
                replaceExisting: true);
            WordOmmlNativeSource.StampFingerprintFromResolvedRange(metadata, badRange);
            WordOmmlFormulaStore.Save(document, metadata);
            AssertTrue(
                WordOmmlConverter.HasVisualTeXNativeEquationNumber(badRange.WordOpenXML),
                "The regression fixture did not materialize the old generated m:eqArr number wrapper.");

            Release(badRange); badRange = null;
            Release(bookmark); bookmark = null;
            var refreshed = WordEquationNumbering.RefreshNumberedOmmlTabLayouts(document);
            AssertEqual(1, refreshed,
                "Opening/reconciling the broken generated m:eqArr formula did not repair it.");
            shapeInventory = document.Shapes;
            Console.WriteLine(
                $"  bad-eqArr fixture after pure-display migration: shapes={shapeInventory.Count}.");
            AssertEqual(0, shapeInventory.Count,
                "The pure-display migration created its external Shape inside the OMath replacement stack instead of deferring it.");
            Release(shapeInventory); shapeInventory = null;
            // The native #(SEQ) producer is complete in the same structural pass.
            // The historical finalizer name remains as a compatibility entry point,
            // but for a healthy direct-SEQ host it performs only an in-place field
            // update and must never create a drawing or anchor paragraph.
            AssertEqual(
                1,
                WordEquationNumbering.FinalizeNumberedOmmlDisplayShapeLayouts(document),
                "The repaired m:eqArr formula was not recognized as one healthy native #(SEQ) host.");
            shapeInventory = document.Shapes;
            Console.WriteLine(
                $"  bad-eqArr fixture after native finalization: shapes={shapeInventory.Count}.");
            AssertEqual(0, shapeInventory.Count,
                "The native bad-eqArr migration recreated a legacy Shape.");
            Release(shapeInventory); shapeInventory = null;
            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context: "broken m:eqArr formula after automatic tab-layout repair",
                updateReference: true);

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context: "repaired bad-package OMML after save/reopen",
                updateReference: true);

            Console.WriteLine(
                "Word bad-package OMML migration acceptance passed: the obsolete m:eqArr/#(REF) wrapper was stripped and replaced once by Word-native #(SEQ), VTEqNum stayed inside the mathematical number slot, and F9/save/reopen remained stable with zero Shape/Table objects.");
        }
        finally
        {
            Release(shapeInventory);
            Release(badRange);
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

    private static string BuildLegacyBadEqArrOmml(
        string omml,
        string formulaId)
    {
        const string MathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        const string WordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var formulaGuid = Guid.Parse(formulaId);
        var targetBookmarkName = "VTEqNum_" + formulaGuid.ToString("N");
        var math = (XNamespace)MathNamespace;
        var word = (XNamespace)WordNamespace;
        var equation = XElement.Parse(
            WordOmmlConverter.ExtractSingleOMath(omml),
            LoadOptions.PreserveWhitespace);
        var formulaNodes = equation
            .Elements()
            .Select(element => new XElement(element))
            .Cast<object>()
            .ToList();
        if (formulaNodes.Count == 0)
            throw new InvalidDataException(
                "The legacy bad-eqArr acceptance fixture requires a nonempty OMath body.");

        XElement FieldRun(XElement content) =>
            new(
                math + "r",
                new XElement(math + "rPr", new XElement(math + "nor")),
                new XElement(word + "rPr", new XElement(word + "noProof")),
                content);
        var delimiter = new XElement(
            math + "d",
            new XElement(
                math + "e",
                FieldRun(new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "begin"),
                    new XAttribute(word + "dirty", "true"))),
                FieldRun(new XElement(
                    word + "instrText",
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    $" REF {targetBookmarkName} \\h \\* CHARFORMAT ")),
                FieldRun(new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "separate"))),
                FieldRun(new XElement(math + "t", "0")),
                FieldRun(new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "end")))));
        var body = new XElement(math + "e", formulaNodes);
        body.Add(
            new XElement(math + "r", new XElement(math + "t", "#")),
            delimiter);
        return new XElement(
                math + "oMath",
                new XElement(
                    math + "eqArr",
                    new XElement(
                        math + "eqArrPr",
                        new XElement(
                            math + "maxDist",
                            new XAttribute(math + "val", "1"))),
                    body))
            .ToString(SaveOptions.DisableFormatting);
    }
}
