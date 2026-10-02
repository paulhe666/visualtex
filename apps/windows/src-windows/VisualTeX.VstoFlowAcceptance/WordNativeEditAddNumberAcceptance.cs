using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordNativeEditAddNumberAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        Word.ListTemplate? headingTemplate = null;
        Word.Range? firstInsertion = null;
        Word.Range? firstInserted = null;
        Word.Range? secondInsertion = null;
        Word.Range? secondInserted = null;
        Word.Range? secondExact = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            headingTemplate = CreateEquationNumberFormatHeadingTemplate(document);

            InsertNumberedFormatHeading(
                application,
                document,
                headingTemplate,
                level: 1,
                "Chapter",
                continuePreviousList: false);
            InsertNumberedFormatHeading(
                application,
                document,
                headingTemplate,
                level: 2,
                "Section",
                continuePreviousList: true);
            document.Content.InsertAfter("\r\r");

            const string mn = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            XNamespace m = mn;

            // Seed the exact legacy/native condition from the user's document:
            // a real Word #(label) display host whose number is literal text,
            // not a VisualTeX field or identifier.
            firstInsertion = document.Paragraphs[3].Range.Duplicate;
            firstInsertion.End = firstInsertion.Start;
            var firstPrepared = XElement.Parse(WordOmmlConverter.BuildWordNativeNumberedOmml(
                "<m:oMath xmlns:m=\"" + mn + "\"><m:r><m:t>x+1</m:t></m:r></m:oMath>",
                "Equation",
                numberBookmarkName: null,
                restartHeadingLevel: 2,
                headingSeparator: "-",
                initialSequenceResult: "1",
                literalHeadingPrefix: null,
                restartSequenceValue: 1));
            var firstArgument = firstPrepared
                .Element(m + "eqArr")!
                .Element(m + "e")!
                .Element(m + "d")!
                .Element(m + "e")!;
            firstArgument.ReplaceNodes(
                new XElement(
                    m + "r",
                    new XElement(m + "rPr", new XElement(m + "nor")),
                    new XElement(m + "t", "1.1-1")));
            firstInserted = WordOmmlConverter.ReplaceWithPreparedOmmlDirect(
                document,
                firstInsertion,
                firstPrepared.ToString(),
                display: true,
                mathFontName: document.OMathFontName);

            secondInsertion = document.Paragraphs[4].Range.Duplicate;
            secondInsertion.End = secondInsertion.Start;
            secondInserted = WordOmmlConverter.ReplaceWithPreparedOmmlDirect(
                document,
                secondInsertion,
                "<m:oMath xmlns:m=\"" + mn + "\"><m:r><m:t>y+2</m:t></m:r></m:oMath>",
                display: true,
                mathFontName: document.OMathFontName);

            WordEquationNumbering.RestoreEquationNumberFormatForConversion(
                document,
                EquationNumberFormat.Heading2DashId);
            // The user's large source document has Track Changes enabled.
            // Exercise the topology-changing edit under that exact ownership
            // model; the plain variant already passed before this stronger run.
            document.TrackRevisions = true;

            var source = WordFormulaHostLookup.ResolveLocal(
                document,
                secondInserted,
                WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException("The unnumbered OMML edit source could not be resolved.");
            source.Numbering = WordFormulaNumberingResolver.ResolveLocal(document, source);
            AssertTrue(!source.Numbering.Numbered, "The edit source unexpectedly began numbered.");

            var request = new WordFormulaHostWriteRequest
            {
                Kind = WordFormulaHostKind.Omml,
                DisplayMode = "block",
                Numbered = true,
                MathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>y</mi><mo>+</mo><mn>2</mn></math>",
                FontSizePoints = 12,
                WidthPoints = 60,
                HeightPoints = 24,
            };

            var before = ReadNativeEditAddNumberState(document);
            WordFormulaHostDescriptor result;
            try
            {
                result = WordFormulaHostMutationKernel.Replace(
                    application,
                    document,
                    source,
                    request);
            }
            catch (AggregateException aggregate)
            {
                Console.WriteLine("edit-add-number aggregate:");
                foreach (var inner in aggregate.Flatten().InnerExceptions)
                    Console.WriteLine(inner.ToString());
                throw;
            }
            var after = ReadNativeEditAddNumberState(document);

            AssertEqual(2, after.Labels.Length, "The edit-add-number path did not retain two native numbered equations.");
            AssertEqual("1.1-1", after.Labels[0], "The native predecessor changed number.");
            AssertEqual("1.1-2", after.Labels[1], "The edited OMML did not join the existing section-number sequence.");
            AssertTrue(after.CurrentSequenceFields >= 2,
                "The field-free predecessor was not adopted into the native sequence before final refresh.");
            AssertEqual(
                EquationNumberFormat.Heading2DashId,
                WordEquationNumberFormatAuthority.Resolve(document),
                "The edit-add-number path did not retain the document-wide number-format authority.");

            document.Undo(1);
            var undone = ReadNativeEditAddNumberState(document);
            AssertEqual(1, undone.Labels.Length, "One-step Undo did not restore the original unnumbered second OMML.");
            AssertEqual("1.1-1", undone.Labels[0], "Undo changed the original native literal number.");
            document.Redo(1);
            var redone = ReadNativeEditAddNumberState(document);
            AssertEqual("1.1-2", redone.Labels[1], "Redo changed the edited native number.");

            secondExact = WordFormulaHostSemanticReader.CreateRange(document, result.Range);
            File.WriteAllText(
                Path.Combine(artifactRoot, "result-host.xml"),
                secondExact.WordOpenXML);
            File.WriteAllText(
                Path.Combine(artifactRoot, "summary.json"),
                JsonSerializer.Serialize(
                    new { Before = before, After = after, Undo = undone, Redo = redone },
                    new JsonSerializerOptions { WriteIndented = true }));

            Console.WriteLine("PASS native OMML edit add-number uses the shared document format and native adoption sequence.");
        }
        finally
        {
            Release(secondExact);
            Release(secondInserted);
            Release(secondInsertion);
            Release(firstInserted);
            Release(firstInsertion);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(headingTemplate);
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private sealed class NativeEditAddNumberState
    {
        public string[] Labels { get; set; } = Array.Empty<string>();
        public int CurrentSequenceFields { get; set; }
    }

    private static NativeEditAddNumberState ReadNativeEditAddNumberState(Word.Document document)
    {
        var labels = new List<string>();
        var currentSequenceFields = 0;
        Word.OMaths? maths = null;
        try
        {
            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                Word.OMath? math = null;
                Word.Range? exact = null;
                Word.Range? number = null;
                Word.Fields? fields = null;
                try
                {
                    math = maths[index];
                    exact = math.Range;
                    if (!IsCurrentNativeEditRange(exact))
                        continue;
                    if (!WordOmmlConverter.HasWordNativeEquationNumberHost(exact.WordOpenXML))
                        continue;
                    labels.Add(WordOmmlConverter.ReadWordNativeEquationNumberText(exact.WordOpenXML));
                    number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
                    fields = number.Fields;
                    for (var fieldIndex = 1; fieldIndex <= fields.Count; fieldIndex++)
                    {
                        Word.Field? field = null;
                        try
                        {
                            field = fields[fieldIndex];
                            if (WordNumberingFieldCollection.IsCurrentField(field)
                                && field.Type == Word.WdFieldType.wdFieldSequence)
                                currentSequenceFields++;
                        }
                        finally { Release(field); }
                    }
                }
                finally
                {
                    Release(fields);
                    Release(number);
                    Release(exact);
                    Release(math);
                }
            }
        }
        finally { Release(maths); }

        return new NativeEditAddNumberState
        {
            Labels = labels.ToArray(),
            CurrentSequenceFields = currentSequenceFields,
        };
    }

    private static bool IsCurrentNativeEditRange(Word.Range range)
    {
        Word.Revisions? revisions = null;
        try
        {
            revisions = range.Revisions;
            if (revisions.Count == 0)
                return true;
            var sawDeletedHistory = false;
            for (var index = 1; index <= revisions.Count; index++)
            {
                Word.Revision? revision = null;
                try
                {
                    revision = revisions[index];
                    switch (revision.Type)
                    {
                        case Word.WdRevisionType.wdRevisionDelete:
                        case Word.WdRevisionType.wdRevisionMovedFrom:
                        case Word.WdRevisionType.wdRevisionCellDeletion:
                        case Word.WdRevisionType.wdRevisionConflictDelete:
                            sawDeletedHistory = true;
                            break;
                        default:
                            return true;
                    }
                }
                finally { Release(revision); }
            }
            return !sawDeletedHistory;
        }
        finally { Release(revisions); }
    }
}
