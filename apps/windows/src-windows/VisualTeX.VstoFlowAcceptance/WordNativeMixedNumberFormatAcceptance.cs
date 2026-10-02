using System.Text.Json;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordNativeMixedNumberFormatAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        Word.ListTemplate? headingTemplate = null;
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

            InsertNativeNumberedFixture(
                document,
                document.Paragraphs[3].Range,
                "x+1",
                headingLevel: 1,
                separator: "-",
                initialResult: "1",
                restart: 1);
            InsertNativeNumberedFixture(
                document,
                document.Paragraphs[4].Range,
                "y+2",
                headingLevel: 0,
                separator: string.Empty,
                initialResult: "2",
                restart: null);

            var before = ReadAllNativeLabels(document);
            File.WriteAllText(
                Path.Combine(artifactRoot, "before.json"),
                JsonSerializer.Serialize(before, new JsonSerializerOptions { WriteIndented = true }));

            document.TrackRevisions = true;
            var service = new WordFormulaService(application);
            var expected = new[]
            {
                (EquationNumberFormat.Heading1DashId, new[] { "1-1", "1-2" }),
                (EquationNumberFormat.Heading2DotId, new[] { "1.1.1", "1.1.2" }),
                (EquationNumberFormat.ContinuousId, new[] { "1", "2" }),
                (EquationNumberFormat.Heading2DashId, new[] { "1.1-1", "1.1-2" }),
                (EquationNumberFormat.Heading1DotId, new[] { "1.1", "1.2" }),
                (EquationNumberFormat.Heading1DashId, new[] { "1-1", "1-2" }),
            };
            var observations = new List<object>();
            foreach (var item in expected)
            {
                var changed = service.SetEquationNumberFormat(item.Item1);
                var after = ReadAllNativeLabels(document);
                File.WriteAllText(
                    Path.Combine(artifactRoot, item.Item1 + ".xml"),
                    document.Content.WordOpenXML);
                observations.Add(new { Format = item.Item1, Changed = changed, Labels = after.ToArray() });
                AssertEqual(2, changed, "Mixed native format rewrite count differs.");
                AssertEqual(2, after.Count, "Mixed native format fixture lost an equation.");
                AssertEqual(item.Item2[0], after[0], "First native number changed incorrectly.");
                AssertEqual(item.Item2[1], after[1], "Second native number did not adopt document format.");
            }
            File.WriteAllText(
                Path.Combine(artifactRoot, "after.json"),
                JsonSerializer.Serialize(observations, new JsonSerializerOptions { WriteIndented = true }));

            Console.WriteLine("PASS mixed native OMML number-format rewrite across all five formats.");
        }
        finally
        {
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(headingTemplate);
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void InsertNativeNumberedFixture(
        Word.Document document,
        Word.Range paragraphRange,
        string body,
        int headingLevel,
        string separator,
        string initialResult,
        int? restart)
    {
        Word.Range? insertion = null;
        Word.Range? inserted = null;
        try
        {
            insertion = paragraphRange.Duplicate;
            insertion.End = insertion.Start;
            const string mn = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var semantic = "<m:oMath xmlns:m=\"" + mn + "\"><m:r><m:t>" + body + "</m:t></m:r></m:oMath>";
            var prepared = WordOmmlConverter.BuildWordNativeNumberedOmml(
                semantic,
                "Equation",
                numberBookmarkName: null,
                restartHeadingLevel: headingLevel,
                headingSeparator: separator,
                initialSequenceResult: initialResult,
                literalHeadingPrefix: null,
                restartSequenceValue: restart);
            inserted = WordOmmlConverter.ReplaceWithPreparedOmmlDirect(
                document,
                insertion,
                prepared,
                display: true,
                mathFontName: document.OMathFontName);
            var maths = inserted.OMaths;
            try
            {
                var math = maths[1];
                try
                {
                    math.Type = Word.WdOMathType.wdOMathDisplay;
                    WordNativeOmmlNumbering.UpdateNativeFields(
                        document,
                        new WordFormulaHostDescriptor
                        {
                            Kind = WordFormulaHostKind.Omml,
                            DisplayMode = "block",
                            Range = new WordFormulaRangeAddress
                            {
                                StoryType = math.Range.StoryType,
                                Start = math.Range.Start,
                                End = math.Range.End,
                            },
                        });
                }
                finally { Release(math); }
            }
            finally { Release(maths); }
        }
        finally
        {
            Release(inserted);
            Release(insertion);
            Release(paragraphRange);
        }
    }

    private static List<string> ReadAllNativeLabels(Word.Document document)
    {
        var labels = new List<string>();
        Word.OMaths? maths = null;
        try
        {
            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                Word.OMath? math = null;
                Word.Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    if (!WordOmmlConverter.HasWordNativeEquationNumberHost(range.WordOpenXML))
                        continue;
                    var xmlLabel = WordOmmlConverter.ReadWordNativeEquationNumberText(range.WordOpenXML);
                    Word.Range? number = null;
                    Word.Fields? fields = null;
                    try
                    {
                        number = WordNativeOmmlNumbering.ResolveNumberContentRange(range);
                        fields = number.Fields;
                        var current = new List<string>();
                        for (var fieldIndex = 1; fieldIndex <= fields.Count; fieldIndex++)
                        {
                            Word.Field? field = null;
                            Word.Range? code = null;
                            Word.Range? result = null;
                            try
                            {
                                field = fields[fieldIndex];
                                if (!WordNumberingFieldCollection.IsCurrentField(field))
                                    continue;
                                code = field.Code;
                                result = field.Result;
                                current.Add($"{field.Type}:{code.Text}=>{result.Text}");
                            }
                            finally { Release(result); Release(code); Release(field); }
                        }
                        Console.WriteLine(
                            $"native-label xml=[{xmlLabel}] range=[{number.Text}] currentFields=[{string.Join(" | ", current)}]");
                    }
                    finally { Release(fields); Release(number); }
                    labels.Add(xmlLabel);
                }
                finally { Release(range); Release(math); }
            }
        }
        finally { Release(maths); }
        return labels;
    }
}
