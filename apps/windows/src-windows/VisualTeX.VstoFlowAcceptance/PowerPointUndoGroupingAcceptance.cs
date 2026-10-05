using System.Diagnostics;
using System.Text.Json;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // Capability investigation only. No product hook, VBA project, clipboard,
    // registry setting or existing presentation is used.
    private static void RunPowerPointUndoGroupingAcceptance(string artifactRoot)
    {
        if (Process.GetProcessesByName("POWERPNT").Length != 0)
            throw new InvalidOperationException("Undo probe requires a new owned PowerPoint instance.");
        var started = DateTime.UtcNow;
        PowerPoint.Application? app = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        CommandBar? temporaryBar = null;
        CommandBarButton? button = null;
        int ownedPid = 0;
        DateTime ownedStart = default;
        var observations = new List<object>();
        string[] History()
        {
            CommandBars? bars = null; CommandBar? standard = null; CommandBarControl? control = null;
            try
            {
                bars = app!.CommandBars; standard = bars["Standard"]; control = standard.FindControl(Id: 128);
                if (!bars.GetEnabledMso("Undo")) return Array.Empty<string>();
                var list = control as CommandBarComboBox ?? throw new InvalidOperationException("Undo list unavailable.");
                return Enumerable.Range(1, list.ListCount).Select(i => list.get_List(i)).ToArray();
            }
            finally { Release(control); Release(standard); Release(bars); }
        }
        try
        {
            app = new PowerPoint.Application();
            GetWindowThreadProcessId(new IntPtr(app.HWND), out var pid);
            using var process = Process.GetProcessById((int)pid);
            if (process.StartTime.ToUniversalTime() < started.AddSeconds(-2) || app.Presentations.Count != 0)
                throw new InvalidOperationException("PowerPoint ownership not established.");
            ownedPid = process.Id; ownedStart = process.StartTime.ToUniversalTime();
            app.Visible = MsoTriState.msoTrue;
            presentation = app.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            app.ActiveWindow.View.GotoSlide(1);
            foreach (var command in new[] { "EquationInsertNew", "InsertBuildingBlocksEquationsGallery" })
            foreach (var scope in new[] { "direct", "temporary-command-callback" })
            {
                shape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 30, 30, 500, 100);
                shape.TextFrame.TextRange.Text = "before SELECT after";
                shape.TextFrame.TextRange.Characters(1, 6).Font.Bold = MsoTriState.msoTrue;
                shape.TextFrame.TextRange.Characters(15, 5).Font.Italic = MsoTriState.msoTrue;
                app.StartNewUndoEntry();
                var before = History();
                var stages = new List<object>();
                void Record(string stage) => stages.Add(new { stage, history = History(), text = shape.TextFrame.TextRange.Text });
                void Write()
                {
                    app.StartNewUndoEntry();
                    Record("begin");
                    shape.TextFrame.TextRange.Characters(8, 6).Text = "x/y";
                    Record("replace-text");
                    shape.TextFrame.TextRange.Characters(8, 3).Font.Size = 24;
                    Record("format-text");
                    shape.Select(MsoTriState.msoTrue);
                    shape.TextFrame.TextRange.Characters(8, 3).Select();
                    Record("select-text");
                    if (!app.CommandBars.GetEnabledMso(command)) throw new InvalidOperationException(command + " unavailable");
                    app.CommandBars.ExecuteMso(command);
                    Record("insert-native-zone");
                    if (app.CommandBars.GetEnabledMso("EquationProfessional")) app.CommandBars.ExecuteMso("EquationProfessional");
                    Record("professional");
                    app.StartNewUndoEntry();
                    Record("end");
                }
                Exception? failure = null;
                try
                {
                    if (scope == "direct") Write();
                    else
                    {
                        temporaryBar = app.CommandBars.Add("VisualTeX owned Undo probe " + Guid.NewGuid(), Temporary: true);
                        button = (CommandBarButton)temporaryBar.Controls.Add(MsoControlType.msoControlButton, Temporary: true);
                        var invoked = false;
                        _CommandBarButtonEvents_ClickEventHandler handler = (CommandBarButton control, ref bool cancel) =>
                        {
                            invoked = true; cancel = true;
                            try { Write(); } catch (Exception error) { failure = error; }
                        };
                        button.Click += handler;
                        try { button.Execute(); }
                        finally { button.Click -= handler; }
                        AssertTrue(invoked, "Temporary command callback was not invoked.");
                    }
                    if (failure is null)
                    {
                        presentation.SaveCopyAs(Path.Combine(artifactRoot, scope + "-" + command + ".pptx"), PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                        app.CommandBars.ExecuteMso("Undo");
                        Record("one-undo");
                    }
                }
                catch (Exception error) { failure = error; }
                finally
                {
                    observations.Add(new { command, scope, before, stages, error = failure?.ToString(),
                        oneUndoTextRestored = failure is null && shape.TextFrame.TextRange.Text == "before SELECT after" });
                    Console.WriteLine(command + "/" + scope + ": " + (failure?.Message ?? shape.TextFrame.TextRange.Text));
                    button?.Delete(); Release(button); button = null;
                    temporaryBar?.Delete(); Release(temporaryBar); temporaryBar = null;
                    shape.Delete(); Release(shape); shape = null;
                }
            }
        }
        finally
        {
            Release(shape); Release(slide);
            if (presentation is not null) { presentation.Saved = MsoTriState.msoTrue; presentation.Close(); Release(presentation); }
            if (app is not null && ownedPid != 0)
            {
                using var process = Process.GetProcessById(ownedPid);
                if (process.StartTime.ToUniversalTime() == ownedStart && app.Presentations.Count == 0) app.Quit();
            }
            Release(app);
            File.WriteAllText(Path.Combine(artifactRoot, "undo-grouping.json"), JsonSerializer.Serialize(
                new { started, ownedPid, ownedStart, scope = "owned native command capability probe; not product acceptance", observations },
                new JsonSerializerOptions { WriteIndented = true }));
        }
    }
}
