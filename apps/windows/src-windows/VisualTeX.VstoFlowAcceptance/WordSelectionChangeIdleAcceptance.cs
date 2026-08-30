using System.Diagnostics;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;
using WinForms = System.Windows.Forms;

namespace VisualTeX.VstoFlowAcceptance;

public sealed class AcceptanceRibbonProbe
{
    public int FormulaInvalidations { get; private set; }
    public int NumberFormatInvalidations { get; private set; }

    public void InvalidateControl(string controlId)
    {
        if (controlId.StartsWith("VisualTeX.WordVsto.FontSize", StringComparison.Ordinal))
            FormulaInvalidations++;
        else if (controlId.StartsWith("VisualTeX.WordVsto.NumberFormat", StringComparison.Ordinal))
            NumberFormatInvalidations++;
    }

    public void Reset()
    {
        FormulaInvalidations = 0;
        NumberFormatInvalidations = 0;
    }
}

internal static partial class Program
{
    private static void RunWordSelectionChangeIdleAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The SelectionChange idle acceptance refuses to attach to the user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        using var host = new WordPerformanceHost(documentPath: null);
        var document = host.Document;
        var application = host.Application;
        document.Activate();
        document.Content.Text = "VT ordinary click probe alpha beta gamma delta epsilon.\r"
            + "VT second ordinary paragraph for click probe.\r";

        var service = new WordFormulaService(application);
        var insertion = Math.Max(document.Content.Start, document.Content.End - 1);
        application.Selection.SetRange(insertion, insertion);
        var formulaId = Guid.NewGuid().ToString("D");
        var session = CreateNumberedOmmlTabSession(
            formulaId,
            document.FullName,
            insertion,
            insertion,
            @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
            originalMetadata: null);
        service.InsertOmml(session, QuadraticFormulaMathMl());
        AssertOmmlTableNumberLifecyclePhase(
            application,
            document,
            formulaId,
            "selection-change-idle fixture");

        var ribbon = new AcceptanceRibbonProbe();
        host.AddIn.OnRibbonLoad(ribbon);
        PumpWordUi();
        ribbon.Reset();
        host.AddIn.ResetSelectionChangeDiagnosticsForAcceptance();

        var ordinaryWatch = Stopwatch.StartNew();
        var ordinarySelection = application.Selection;
        try
        {
            for (var index = 0; index < 8; index++)
            {
                var position = document.Content.Start + 3 + index * 3;
                ordinarySelection.SetRange(position, position);
                PumpWordUi();
            }
        }
        finally { Release(ordinarySelection); }
        ordinaryWatch.Stop();
        var ordinary = host.AddIn.ReadSelectionChangeDiagnosticsForAcceptance();
        AssertTrue(ordinary.SelectionChanges >= 8,
            $"Ordinary click probe raised only {ordinary.SelectionChanges} SelectionChange events.");
        AssertEqual(0, ordinary.FormulaStateReads,
            "Ordinary prose-to-prose clicks still performed a full formula-state COM read.");
        AssertEqual(0, ordinary.DeferredCaretPasses,
            "Ordinary prose clicks still scheduled the retired/deferred caret pass.");
        AssertEqual(0, ordinary.EquationFormatReads,
            "Ordinary prose clicks still reread the document equation-number format.");
        AssertEqual(0, ribbon.NumberFormatInvalidations,
            "Ordinary prose clicks still invalidated equation-number Ribbon controls.");
        AssertEqual(0, ribbon.FormulaInvalidations,
            "Ordinary prose-to-prose clicks still invalidated formula-font Ribbon controls.");
        Console.WriteLine(
            $"[SELECTION CHANGE ORDINARY] events={ordinary.SelectionChanges} formulaReads={ordinary.FormulaStateReads} deferredCaret={ordinary.DeferredCaretPasses} numberFormatReads={ordinary.EquationFormatReads} fontInvalidations={ribbon.FormulaInvalidations} numberInvalidations={ribbon.NumberFormatInvalidations} elapsedMs={ordinaryWatch.ElapsedMilliseconds}.");

        ribbon.Reset();
        host.AddIn.ResetSelectionChangeDiagnosticsForAcceptance();
        Word.Table? table = null;
        Word.Cell? centerCell = null;
        Word.Range? centerRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Selection? mathSelection = null;
        var mathWatch = Stopwatch.StartNew();
        try
        {
            table = WordEquationNumbering.FindNumberedEquationTable(document, formulaId)
                ?? throw new InvalidDataException("SelectionChange fixture lost its numbered OMML table.");
            centerCell = table.Cell(1, 2);
            centerRange = centerCell.Range;
            maths = centerRange.OMaths;
            AssertEqual(1, maths.Count,
                "SelectionChange fixture center cell lost its OMath.");
            math = maths[1];
            mathRange = math.Range.Duplicate;
            mathSelection = application.Selection;
            var span = Math.Max(1, mathRange.End - mathRange.Start);
            for (var index = 0; index < 6; index++)
            {
                var position = mathRange.Start + Math.Min(span - 1, index);
                mathSelection.SetRange(position, position);
                PumpWordUi();
            }
        }
        finally
        {
            Release(mathSelection);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(centerRange);
            Release(centerCell);
            Release(table);
        }
        mathWatch.Stop();
        var omml = host.AddIn.ReadSelectionChangeDiagnosticsForAcceptance();
        AssertTrue(omml.SelectionChanges >= 1,
            "OMML click probe did not raise SelectionChange.");
        AssertTrue(omml.FormulaStateReads <= 1,
            $"Moving within one numbered OMML formula triggered {omml.FormulaStateReads} full formula-state reads; only the initial entry may read once.");
        AssertEqual(0, omml.DeferredCaretPasses,
            "Native OMML clicks still scheduled an OLE-only deferred caret pass.");
        AssertEqual(0, omml.EquationFormatReads,
            "Native OMML clicks still reread the equation-number format.");
        AssertEqual(0, ribbon.NumberFormatInvalidations,
            "Native OMML clicks still invalidated equation-number Ribbon controls.");
        AssertTrue(ribbon.FormulaInvalidations <= 3,
            $"Moving within one numbered OMML formula invalidated formula-font controls {ribbon.FormulaInvalidations} times; only one three-control refresh is allowed.");
        Console.WriteLine(
            $"[SELECTION CHANGE OMML] events={omml.SelectionChanges} formulaReads={omml.FormulaStateReads} deferredCaret={omml.DeferredCaretPasses} numberFormatReads={omml.EquationFormatReads} fontInvalidations={ribbon.FormulaInvalidations} numberInvalidations={ribbon.NumberFormatInvalidations} elapsedMs={mathWatch.ElapsedMilliseconds}.");
        Console.WriteLine(
            "Word SelectionChange idle acceptance passed: prose-to-prose clicks perform zero formula reads/zero Ribbon refreshes, repeated movement inside one numbered OMML performs one initial formula read/one three-control refresh total, equation-number controls never refresh on SelectionChange, and the OLE-only deferred caret retry is never queued for prose or OMML.");
    }

    private static void PumpWordUi()
    {
        for (var index = 0; index < 3; index++)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(20);
        }
    }
}
