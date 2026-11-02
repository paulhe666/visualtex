using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Office.Core;
using Microsoft.Win32;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using VisualTeX.PowerPointVsto;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // Native equations in table cells: inline and display insertion at the caret
    // (plain and merged cells), source capture, edit, double-click hit and delete,
    // each leaving the cell's other text and the slide's shapes as they were.
    private static void RunPowerPointTableCellAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var application = StartOwnedPowerPoint();
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? tableShape = null;
        PowerPoint.Table? table = null;
        var rows = new List<object>();
        try
        {
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            tableShape = slide.Shapes.AddTable(3, 3, 60, 60, 600, 240);
            table = tableShape.Table;
            table.Cell(1, 1).Merge(table.Cell(1, 2));
            CellText(table, 1, 1).Text = "merged";
            CellText(table, 2, 2).Text = "before after";
            CellText(table, 2, 2).Characters(1, 6).Font.Name = "Arial";
            CellText(table, 2, 2).Characters(8, 5).Font.Name = "Calibri";
            using var service = new PowerPointFormulaService(application);
            var shapeCount = slide.Shapes.Count;

            OfficeSelection Capture(int row, int column, int start, int length)
            {
                CellText(table, row, column).Characters(start, length).Select();
                return service.CaptureTextInsertionTarget();
            }
            void Insert(OfficeSelection target, string displayMode, string mathMl)
            {
                var session = CreatePowerPointAcceptanceSession("create", "wordOmml", Guid.NewGuid().ToString(),
                    target.ObjectId, null, displayMode);
                session.SourceDocumentId = target.DocumentId;
                session.ExportResult!.MathMl = mathMl;
                var result = service.InsertOmml(session);
                AssertEqual(tableShape!.Name, result.ObjectId, "The insertion did not report the table.");
            }

            // Inline, between two words of a cell.
            var inline = Capture(2, 2, 8, 0);
            AssertContains(inline.ObjectId, ":" + tableShape.Id + ".2.2:", "The caret's cell is missing from the target.");
            Insert(inline, "inline", PowerPointAcceptanceMathMl);
            AssertEqual(1, ZoneCount(table, 2, 2), "No native equation in the cell.");
            var text = CellText(table, 2, 2).Text;
            AssertTrue(text.StartsWith("before ", StringComparison.Ordinal) && text.EndsWith("after", StringComparison.Ordinal),
                "The cell's words changed: " + text);
            AssertEqual("Arial", CellText(table, 2, 2).Characters(1, 6).Font.Name, "The font before the equation changed.");
            AssertEqual("Calibri", CellText(table, 2, 2).Characters(text.Length - 4, 5).Font.Name, "The font after the equation changed.");
            AssertEqual(shapeCount, slide.Shapes.Count, "Inserting into a cell added a shape.");
            rows.Add(new { stage = "inline", text });

            // Display, into an empty cell, and inline into a merged cell.
            Insert(Capture(3, 1, 1, 0), "block", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msup><mi>e</mi><mi>x</mi></msup></math>");
            AssertEqual(1, ZoneCount(table, 3, 1), "No display equation in the empty cell.");
            var merged = Capture(1, 1, 4, 0);
            AssertContains(merged.ObjectId, ":" + tableShape.Id + ".1.1:", "The merged cell's first position is not the target.");
            Insert(merged, "inline", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>y</mi></math>");
            AssertEqual(1, ZoneCount(table, 1, 1), "No equation in the merged cell.");
            Insert(Capture(3, 3, 1, 0), "inline", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msub><mi mathvariant=\"normal\">M</mi><mrow><mi mathvariant=\"normal\">N</mi><mi mathvariant=\"normal\">a</mi></mrow></msub><mo>=</mo><mi mathvariant=\"bold\">v</mi></math>");
            AssertEqual(1, ZoneCount(table, 3, 3), "No equation with styled letters.");
            AssertEqual(shapeCount, slide.Shapes.Count, "Inserting into cells added a shape.");
            rows.Add(new { stage = "display+merged", display = CellText(table, 3, 1).Text, merged = CellText(table, 1, 1).Text });

            // Edit the inline equation from the caret inside it.
            var zone = Zone(table, 2, 2);
            CellText(table, 2, 2).Characters(zone.Start + 1, 0).Select();
            var selected = service.ReadSelection();
            AssertTrue(PowerPointFormulaService.IsNativeOmmlReference(selected.ObjectId), "The caret in the cell's equation is no native candidate.");
            selected = service.LoadNativeOmmlSelection(selected);
            AssertContains(selected.Metadata?.Latex, @"\frac", "The cell equation's source was not read.");
            var edit = CreatePowerPointAcceptanceSession("edit", "wordOmml", selected.FormulaId!, selected.ObjectId, selected.Metadata, "inline");
            edit.SourceDocumentId = selected.DocumentId;
            edit.ExportResult!.MathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msup><mi>z</mi><mn>9</mn></msup></math>";
            service.ReplaceOmml(edit);
            AssertEqual(1, ZoneCount(table, 2, 2), "The edit did not leave one equation in the cell.");
            zone = Zone(table, 2, 2);
            CellText(table, 2, 2).Characters(zone.Start + 1, 0).Select();
            var reopened = service.LoadNativeOmmlSelection(service.ReadSelection());
            AssertContains(reopened.Metadata?.Latex, "9", "The cell equation was not replaced.");
            text = CellText(table, 2, 2).Text;
            AssertTrue(text.StartsWith("before ", StringComparison.Ordinal) && text.EndsWith("after", StringComparison.Ordinal),
                "The edit changed the cell's words: " + text);
            AssertEqual(shapeCount, slide.Shapes.Count, "Editing a cell equation added a shape.");
            rows.Add(new { stage = "edit", text, reopened.Metadata?.Latex });

            // Double-click hit test over the equation (a cell's text reports its bounds
            // from the table's corner).
            var window = application.ActiveWindow;
            var bounds = ZoneBounds(table, 2, 2);
            var hit = service.ReadFormulaAtScreenPoint(
                window.PointsToScreenPixelsX(tableShape.Left + bounds.Left + bounds.Width / 2),
                window.PointsToScreenPixelsY(tableShape.Top + bounds.Top + bounds.Height / 2));
            AssertTrue(hit is not null && PowerPointFormulaService.IsNativeOmmlReference(hit.ObjectId), "The double-click missed the cell equation.");
            AssertContains(hit!.ObjectId, ":" + tableShape.Id + ".2.2:", "The double-click found another cell.");
            var miss = service.ReadFormulaAtScreenPoint(
                window.PointsToScreenPixelsX(tableShape.Left + 5), window.PointsToScreenPixelsY(tableShape.Top + tableShape.Height - 5));
            AssertTrue(miss is null, "A click on a cell without an equation found one.");
            rows.Add(new { stage = "hit", hit.ObjectId });

            // Delete only the equation.
            zone = Zone(table, 2, 2);
            CellText(table, 2, 2).Characters(zone.Start + 1, 0).Select();
            service.DeleteSelectedFormula();
            AssertEqual(0, ZoneCount(table, 2, 2), "The cell equation was not deleted.");
            AssertEqual("before after", CellText(table, 2, 2).Text, "Deleting the equation changed the cell's words.");
            AssertEqual(shapeCount, slide.Shapes.Count, "Deleting a cell equation removed a shape.");
            AssertEqual(1, ZoneCount(table, 3, 1), "Deleting one cell's equation touched another cell.");
            rows.Add(new { stage = "delete", text = CellText(table, 2, 2).Text });

            var saved = Path.Combine(artifactRoot, "table-cells.pptx");
            presentation.SaveCopyAs(saved, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            using (var archive = System.IO.Compression.ZipFile.OpenRead(saved))
            using (var reader = new StreamReader(archive.GetEntry("ppt/slides/slide1.xml")!.Open()))
            {
                var xml = reader.ReadToEnd();
                AssertTrue(xml.Contains("<a:tc><a:txBody>") && xml.Contains("<a14:m>"), "The saved table has no native equation.");
            }
            Console.WriteLine("PASS PowerPoint native equations in table cells");
        }
        finally
        {
            File.WriteAllText(Path.Combine(artifactRoot, "table-cells.json"),
                JsonSerializer.Serialize(rows, new JsonSerializerOptions { WriteIndented = true }));
            Release(table);
            Release(tableShape);
            Release(slide);
            if (presentation is not null)
            {
                try { presentation.Saved = MsoTriState.msoTrue; presentation.Close(); } catch { }
                Release(presentation);
            }
            try { application.Quit(); } catch { }
            Release(application);
        }
    }

    // WPS Pro can register itself as PowerPoint.Application's server: PowerPoint is
    // started from its executable and taken from the running-object table.
    private static PowerPoint.Application StartOwnedPowerPoint()
    {
        if (Process.GetProcessesByName("POWERPNT").Length != 0)
            throw new InvalidOperationException("This acceptance needs no running PowerPoint; user instances are not touched.");
        if (Process.GetProcessesByName("wpp").Length != 0)
            throw new InvalidOperationException("WPS Presentation is running and can answer for PowerPoint's class.");
        using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE");
        var executable = key?.GetValue(null) as string
            ?? throw new InvalidOperationException("PowerPoint is not installed.");
        using var process = Process.Start(executable.Trim('"'));
        var deadline = Stopwatch.StartNew();
        while (deadline.Elapsed < TimeSpan.FromSeconds(30))
        {
            try { return (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application"); }
            catch (COMException) { Thread.Sleep(300); }
        }
        throw new TimeoutException("PowerPoint did not register in the running-object table.");
    }

    private static PowerPoint.TextRange CellText(PowerPoint.Table table, int row, int column) =>
        table.Cell(row, column).Shape.TextFrame.TextRange;

    private static int ZoneCount(PowerPoint.Table table, int row, int column)
    {
        dynamic range = table.Cell(row, column).Shape.TextFrame2.TextRange;
        return Convert.ToInt32(range.MathZones().Count);
    }

    private static (int Start, int Length) Zone(PowerPoint.Table table, int row, int column)
    {
        dynamic range = table.Cell(row, column).Shape.TextFrame2.TextRange;
        dynamic zone = range.MathZones(1, 1);
        return (Convert.ToInt32(zone.Start), Convert.ToInt32(zone.Length));
    }

    private static (float Left, float Top, float Width, float Height) ZoneBounds(PowerPoint.Table table, int row, int column)
    {
        dynamic range = table.Cell(row, column).Shape.TextFrame2.TextRange;
        dynamic zone = range.MathZones(1, 1);
        return (Convert.ToSingle(zone.BoundLeft), Convert.ToSingle(zone.BoundTop),
            Convert.ToSingle(zone.BoundWidth), Convert.ToSingle(zone.BoundHeight));
    }
}
