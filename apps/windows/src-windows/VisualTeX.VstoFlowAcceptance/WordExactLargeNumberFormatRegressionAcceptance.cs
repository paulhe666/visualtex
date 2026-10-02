using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class ExactLargeFormatObservation
    {
        public string Format { get; set; } = string.Empty;
        public int Changed { get; set; }
        public long Milliseconds { get; set; }
        public string[] Before { get; set; } = Array.Empty<string>();
        public string[] After { get; set; } = Array.Empty<string>();
        public string[] Undo { get; set; } = Array.Empty<string>();
        public string[] Redo { get; set; } = Array.Empty<string>();
        public string[] Reopen { get; set; } = Array.Empty<string>();
        public int MathTypeOleBefore { get; set; }
        public int MathTypeOleAfter { get; set; }
        public int TablesBefore { get; set; }
        public int TablesAfter { get; set; }
        public bool TrackRevisions { get; set; }
    }

    private static void RunWordExactLargeNumberFormatRegressionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var source = Environment.GetEnvironmentVariable("VISUALTEX_EXACT_LARGE_NUMBER_SOURCE");
        if (string.IsNullOrWhiteSpace(source))
        {
            source = Path.GetFullPath(Path.Combine(
                Environment.CurrentDirectory,
                "artifacts",
                "mathtype-history-20260930",
                "inputs-baseline-ay9",
                "InputSnapshot.docx"));
        }
        source = Path.GetFullPath(source);
        if (!File.Exists(source))
            throw new FileNotFoundException("Exact large numbering source is missing.", source);

        string sourceHash;
        using (var sha256 = SHA256.Create())
        using (var stream = File.OpenRead(source))
        {
            sourceHash = BitConverter
                .ToString(sha256.ComputeHash(stream))
                .Replace("-", string.Empty);
        }
        AssertEqual(
            "868CB1460226AD432B98EEEE91EF49351146F8348B95E02702AF68E3F9E1F1C6",
            sourceHash,
            "Exact large numbering source is not the archived 1 (1).docx baseline.");

        var formats = new[]
        {
            EquationNumberFormat.ContinuousId,
            EquationNumberFormat.Heading1DotId,
            EquationNumberFormat.Heading1DashId,
            EquationNumberFormat.Heading2DotId,
            EquationNumberFormat.Heading2DashId,
        };

        var observations = new List<ExactLargeFormatObservation>();
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);
            foreach (var format in formats)
            {
                var path = Path.Combine(artifactRoot, "exact-large-" + format + ".docx");
                File.Copy(source, path, overwrite: true);

                document = application.Documents.Open(
                    path,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);
                document.Activate();

                var before = ReadCurrentNativeNumberLabelsBounded(document);
                AssertSequenceEqual(
                    new[] { "1-1", "1-2", "2-1" },
                    before,
                    $"Exact large source native labels changed before {format}.");
                var oleBefore = CountMathTypeOleShapes(document);
                var tablesBefore = document.Tables.Count;
                var tracking = document.TrackRevisions;

                var service = new WordFormulaService(application);
                var watch = Stopwatch.StartNew();
                var changed = service.SetEquationNumberFormat(format);
                watch.Stop();

                var after = ReadCurrentNativeNumberLabelsBounded(document);
                AssertEqual(3, after.Length,
                    $"Exact large {format} lost a native numbered OMath.");
                AssertTrue(changed >= 3,
                    $"Exact large {format} did not rewrite all three native numbers.");
                foreach (var label in after)
                {
                    AssertTrue(
                        WordNativeNumberingAdoptionPlan.TryParseLabel(label, out var style),
                        $"Exact large {format} produced an unparseable native label '{label}'.");
                    AssertEqual(
                        format,
                        style.FormatId,
                        $"Exact large {format} produced a label from another visible number format.");
                }

                document.Undo(1);
                var undone = ReadCurrentNativeNumberLabelsBounded(document);
                AssertSequenceEqual(
                    before,
                    undone,
                    $"Exact large {format} one-step Undo did not restore all native labels.");
                AssertEqual(oleBefore, CountMathTypeOleShapes(document),
                    $"Exact large {format} Undo changed MathType OLE count.");
                AssertEqual(tablesBefore, document.Tables.Count,
                    $"Exact large {format} Undo changed table count.");

                document.Redo(1);
                var redone = ReadCurrentNativeNumberLabelsBounded(document);
                AssertSequenceEqual(
                    after,
                    redone,
                    $"Exact large {format} Redo changed native labels.");
                AssertEqual(tracking, document.TrackRevisions,
                    $"Exact large {format} changed TrackRevisions.");
                AssertEqual(oleBefore, CountMathTypeOleShapes(document),
                    $"Exact large {format} changed MathType OLE count.");
                AssertEqual(tablesBefore, document.Tables.Count,
                    $"Exact large {format} changed table count.");

                document.Save();
                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                document = null;

                document = application.Documents.Open(
                    path,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);
                var reopened = ReadCurrentNativeNumberLabelsBounded(document);
                AssertSequenceEqual(
                    after,
                    reopened,
                    $"Exact large {format} save/reopen changed native labels.");
                AssertEqual(oleBefore, CountMathTypeOleShapes(document),
                    $"Exact large {format} save/reopen changed MathType OLE count.");
                AssertEqual(tablesBefore, document.Tables.Count,
                    $"Exact large {format} save/reopen changed table count.");

                observations.Add(new ExactLargeFormatObservation
                {
                    Format = format,
                    Changed = changed,
                    Milliseconds = watch.ElapsedMilliseconds,
                    Before = before,
                    After = after,
                    Undo = undone,
                    Redo = redone,
                    Reopen = reopened,
                    MathTypeOleBefore = oleBefore,
                    MathTypeOleAfter = CountMathTypeOleShapes(document),
                    TablesBefore = tablesBefore,
                    TablesAfter = document.Tables.Count,
                    TrackRevisions = tracking,
                });

                Console.WriteLine(
                    $"PASS exact-large format={format} changed={changed} ms={watch.ElapsedMilliseconds} labels=[{string.Join(",", after)}] tracking={tracking}");

                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                document = null;
            }

            File.WriteAllText(
                Path.Combine(artifactRoot, "summary.json"),
                JsonSerializer.Serialize(
                    new
                    {
                        Time = DateTimeOffset.Now,
                        Source = source,
                        SourceSha256 = sourceHash,
                        Passed = observations.Count,
                        Rows = observations,
                    },
                    new JsonSerializerOptions { WriteIndented = true }));
        }
        finally
        {
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

    private static string[] ReadCurrentNativeNumberLabelsBounded(Word.Document document)
    {
        var labels = new List<(int Start, string Label)>();
        Word.Range? scope = null;
        Word.Range? search = null;
        Word.Range? unit = null;
        Word.TextRetrievalMode? retrieval = null;
        Word.Find? find = null;
        Word.OMaths? maths = null;
        try
        {
            scope = document.Content;
            var cursor = scope.Start;
            var end = scope.End;
            var seen = new HashSet<int>();
            while (cursor < end)
            {
                Release(maths); maths = null;
                Release(unit); unit = null;
                Release(find); find = null;
                Release(retrieval); retrieval = null;
                Release(search); search = scope.Duplicate;
                search.SetRange(cursor, end);
                retrieval = search.TextRetrievalMode;
                retrieval.IncludeFieldCodes = true;
                retrieval.IncludeHiddenText = true;
                find = search.Find;
                find.ClearFormatting();
                find.Text = "#";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.MatchWildcards = false;
                find.MatchWholeWord = false;
                find.MatchCase = false;
                find.MatchSoundsLike = false;
                find.MatchAllWordForms = false;
                find.Format = false;
                if (!find.Execute())
                    break;

                unit = WordNumberingFieldCollection.CaptureSearchUnit(search);
                if (unit.End <= cursor || unit.Start >= end || !unit.InStory(scope))
                    throw new InvalidDataException(
                        "Exact large native-number scan repeated or escaped its bounded search unit.");
                cursor = unit.End;
                maths = unit.OMaths;
                for (var index = 1; index <= maths.Count; index++)
                {
                    Word.OMath? math = null;
                    Word.Range? exact = null;
                    try
                    {
                        math = maths[index];
                        if (math.Type != Word.WdOMathType.wdOMathDisplay)
                            continue;
                        exact = math.Range;
                        if (!IsCurrentNativeEditRange(exact)
                            || !seen.Add(exact.Start))
                            continue;
                        var xml = exact.WordOpenXML;
                        if (!WordOmmlConverter.HasWordNativeEquationNumberHost(xml))
                            continue;
                        labels.Add((
                            exact.Start,
                            WordOmmlConverter.ReadWordNativeEquationNumberText(xml)));
                    }
                    finally
                    {
                        Release(exact);
                        Release(math);
                    }
                }
            }
        }
        finally
        {
            Release(maths);
            Release(unit);
            Release(find);
            Release(retrieval);
            Release(search);
            Release(scope);
        }

        return labels
            .OrderBy(item => item.Start)
            .Select(item => item.Label)
            .ToArray();
    }

    private static void AssertSequenceEqual(
        IReadOnlyList<string> expected,
        IReadOnlyList<string> actual,
        string message)
    {
        if (expected.Count != actual.Count
            || !expected.SequenceEqual(actual, StringComparer.Ordinal))
        {
            throw new InvalidOperationException(
                message
                + $" Expected=[{string.Join(",", expected)}]"
                + $" actual=[{string.Join(",", actual)}].");
        }
    }
}
