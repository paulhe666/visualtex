using System.Text.Json;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordNativeFieldRevisionProbeAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.Range? inserted = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? exact = null;
        Word.Range? number = null;
        Word.Fields? fields = null;
        Word.Field? sequence = null;
        Word.Range? code = null;
        Word.Range? result = null;
        Word.Range? span = null;
        Word.Fields? writer = null;
        Word.Field? replacement = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "Chapter\r\r";
            Word.Range? heading = null;
            try
            {
                heading = document.Paragraphs[1].Range;
                object style = Word.WdBuiltinStyle.wdStyleHeading1;
                heading.set_Style(ref style);
                heading.ListFormat.ApplyNumberDefault();
            }
            finally { Release(heading); }

            insertion = document.Paragraphs[2].Range.Duplicate;
            insertion.End = insertion.Start;
            const string mn = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var semantic = "<m:oMath xmlns:m=\"" + mn + "\"><m:r><m:t>x+1</m:t></m:r></m:oMath>";
            var prepared = WordOmmlConverter.BuildWordNativeNumberedOmml(
                semantic, "Equation", null, 1, "-", "1", null, 1);
            inserted = WordOmmlConverter.ReplaceWithPreparedOmmlDirect(
                document, insertion, prepared, display: true, mathFontName: document.OMathFontName);
            maths = inserted.OMaths;
            math = maths[1];
            math.Type = Word.WdOMathType.wdOMathDisplay;
            exact = math.Range;
            number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);

            File.WriteAllText(Path.Combine(artifactRoot, "before.json"),
                JsonSerializer.Serialize(SnapshotFields(number), new JsonSerializerOptions { WriteIndented = true }));

            document.TrackRevisions = true;
            fields = number.Fields;
            for (var i = 1; i <= fields.Count; i++)
            {
                Release(sequence); sequence = fields[i];
                if (sequence.Type != Word.WdFieldType.wdFieldSequence) continue;
                code = sequence.Code; result = sequence.Result;
                span = number.Duplicate;
                span.SetRange(code.Start - 1, result.End + 1);
                break;
            }
            if (span is null) throw new InvalidDataException("Probe did not find the initial SEQ field.");
            Word.Range? capturedCode = null;
            try
            {
                AssertTrue(WordNumberingFieldCollection.TryCaptureCurrentCode(sequence!, out capturedCode, out var instruction),
                    "The original sequence must be current before replacement.");
                AssertTrue(instruction.Contains("\\s 1"), "Initial native sequence scope");
            }
            finally { Release(capturedCode); }
            writer = span.Fields;
            replacement = writer.Add(
                span,
                Word.WdFieldType.wdFieldSequence,
                "Equation \\s 2 \\* ARABIC",
                true);

            Release(number); number = null;
            Release(exact); exact = null;
            exact = math.Range;
            number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
            File.WriteAllText(Path.Combine(artifactRoot, "after-replace.json"),
                JsonSerializer.Serialize(SnapshotFields(number), new JsonSerializerOptions { WriteIndented = true }));

            capturedCode = null;
            try
            {
                AssertTrue(!WordNumberingFieldCollection.TryCaptureCurrentCode(sequence!, out capturedCode, out _),
                    "Revision-deleted predecessor must not re-enter the current field projection.");
                AssertTrue(capturedCode is null, "Rejected field must not transfer a code owner.");
                AssertTrue(WordNumberingFieldCollection.TryCaptureCurrentCode(replacement, out capturedCode, out var instruction),
                    "Inserted replacement must remain current next to deleted fields.");
                AssertTrue(instruction.Contains("\\s 2"), "Replacement native sequence scope");
            }
            finally { Release(capturedCode); }

            var currentSequences = new List<string>();
            using (var snapshot = WordNumberingFieldCollection.Capture(document))
                for (var i = 1; i <= snapshot.Count; i++)
                {
                    using var entry = snapshot.TakeEntry(i);
                    if (entry.Type == Word.WdFieldType.wdFieldSequence)
                        currentSequences.Add(entry.Instruction);
                }
            AssertEqual(1, currentSequences.Count, "Only the replacement SEQ belongs to the current snapshot.");
            AssertTrue(currentSequences[0].Contains("\\s 2"), "Snapshot must keep the replacement code.");
            var host = new WordFormulaHostDescriptor
            {
                Kind = WordFormulaHostKind.Omml,
                DisplayMode = "block",
                Range = new WordFormulaRangeAddress { Start=exact.Start, End=exact.End, StoryType=exact.StoryType },
            };
            var refreshed = WordNativeOmmlNumbering.RefreshCanonicalNativeFields(math, host);
            AssertTrue(refreshed.Numbering.Numbered, "Retained native owner preserves numbered topology under revisions.");
            File.WriteAllText(Path.Combine(artifactRoot, "current-projection-check.json"), JsonSerializer.Serialize(new {
                passed=true, assertions=9, currentSequences,
                scope="Actual isolated Word: initial/current, revision-deleted predecessor, inserted replacement, snapshot and retained-OMath refresh. No user document."
            }, new JsonSerializerOptions { WriteIndented=true }));

            Console.WriteLine("Native field revision probe complete.");
        }
        finally
        {
            Release(replacement); Release(writer); Release(span); Release(result); Release(code);
            Release(sequence); Release(fields); Release(number); Release(exact); Release(math); Release(maths);
            Release(inserted); Release(insertion);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static object[] SnapshotFields(Word.Range number)
    {
        var rows = new List<object>();
        Word.Fields? fields = null;
        try
        {
            fields = number.Fields;
            for (var i = 1; i <= fields.Count; i++)
            {
                Word.Field? field = null;
                Word.Range? code = null;
                Word.Range? result = null;
                Word.Range? span = null;
                try
                {
                    field = fields[i];
                    code = field.Code;
                    result = field.Result;
                    span = number.Duplicate;
                    span.SetRange(Math.Max(number.Start, code.Start - 1), Math.Min(number.End, result.End + 1));
                    rows.Add(new
                    {
                        Index = i,
                        Type = field.Type.ToString(),
                        Code = code.Text,
                        CodeStart = code.Start,
                        CodeEnd = code.End,
                        Result = result.Text,
                        ResultStart = result.Start,
                        ResultEnd = result.End,
                        SpanStart = span.Start,
                        SpanEnd = span.End,
                        CodeRevisions = RevisionTypes(code),
                        ResultRevisions = RevisionTypes(result),
                        SpanRevisions = RevisionTypes(span),
                    });
                }
                finally { Release(span); Release(result); Release(code); Release(field); }
            }
        }
        finally { Release(fields); }
        return rows.ToArray();
    }

    private static string[] RevisionTypes(Word.Range range)
    {
        var values = new List<string>();
        Word.Revisions? revisions = null;
        try
        {
            revisions = range.Revisions;
            for (var i = 1; i <= revisions.Count; i++)
            {
                Word.Revision? revision = null;
                try
                {
                    revision = revisions[i];
                    values.Add(revision.Type.ToString());
                }
                finally { Release(revision); }
            }
        }
        finally { Release(revisions); }
        return values.ToArray();
    }
}
