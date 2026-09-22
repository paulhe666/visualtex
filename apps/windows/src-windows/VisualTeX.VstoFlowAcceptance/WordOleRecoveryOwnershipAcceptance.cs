using System.Reflection;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // A deterministic ownership test, not an injected native Word failure. The
    // recovery predicate must reject a pre-existing formula when no insertion
    // happened, even when its ProgID and numeric position match exactly.
    private static void RunWordOleRecoveryOwnershipAcceptance(string artifactRoot)
    {
        var previewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX", "office", "temp", "recovery-owner-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(previewRoot);
        var png = Path.Combine(previewRoot, "formula.png");
        var svg = Path.Combine(previewRoot, "formula.svg");
        WriteAcceptancePng(png, "x+1", 260, 96);
        File.WriteAllText(svg,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"96\" viewBox=\"0 0 260 96\"><text x=\"6\" y=\"66\" font-size=\"48\">x+1</text></svg>");
        var emf = OfficeOlePreview.CreateVectorEmfFromSvg(svg, 260, 96);
        var method = typeof(WordFormulaService).GetMethod(
            "TryRecoverMaterializedOleAfterCommandFailure", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new MissingMethodException("OLE recovery predicate not found.");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? recovered = null;
        Word.InlineShape? existing = null;
        Word.Range? range = null;
        Word.Range? rightWitness = null;
        try
        {
            application = CreateFreshRollbackAutomationWord(artifactRoot);
            document = application.ActiveDocument;
            document.Content.Text = "before after\r";
            var service = new WordFormulaService(application);
            document.Range(7, 7).Select();
            var oldSession = CreateSimpleVisualTeXSourceSession("x+1", numbered: false);
            oldSession.DisplayMode = "inline";
            service.InsertOle(oldSession, png, emf);
            existing = document.InlineShapes[1];
            range = existing.Range;
            var oldStart = range.Start;
            var oldText = document.Content.Text;
            rightWitness = document.Range(oldStart, oldStart + 1);
            object?[] arguments = method.GetParameters().Length == 3
                ? new object?[] { document, oldStart, null }
                : new object?[] { document, oldStart, rightWitness, null };
            recovered = method.Invoke(null, arguments) as Word.InlineShape;
            Console.WriteLine($"[OLE RECOVERY NO INSERT] acceptedExisting={recovered is not null}; evidence={arguments[arguments.Length - 1]}");
            AssertTrue(recovered is null,
                "Recovery adopted an existing VisualTeX OLE without any insertion. Position and ProgID alone do not prove ownership.");
            AssertEqual(oldText, document.Content.Text, "The negative recovery probe modified document content.");
            Release(range); range = null;
            Release(existing); existing = null;

            range = document.Range(0, 0);
            Release(rightWitness);
            rightWitness = document.Range(0, 1);
            range.Select();
            var newSession = CreateSimpleVisualTeXSourceSession("y+2", numbered: false);
            newSession.DisplayMode = "inline";
            service.InsertOle(newSession, png, emf);
            arguments = method.GetParameters().Length == 3
                ? new object?[] { document, 0, null }
                : new object?[] { document, 0, rightWitness, null };
            recovered = method.Invoke(null, arguments) as Word.InlineShape;
            AssertTrue(recovered is not null, "Recovery did not recognize an actually materialized OLE at the exact insertion.");
            var metadata = WordFormulaMetadataReader.TryReadEmbeddedNativeOle(recovered!);
            AssertEqual(newSession.FormulaId, metadata?.FormulaId,
                "Recovery returned the neighboring old object instead of the newly inserted OLE.");
            Console.WriteLine("[OLE RECOVERY OWNERSHIP PASS] rejected-existing=True; accepted-new=True; neighbor-preserved=True");
        }
        finally
        {
            Release(recovered);
            Release(rightWitness);
            Release(range);
            Release(existing);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            try { Directory.Delete(previewRoot, recursive: true); } catch { }
        }
    }
}
