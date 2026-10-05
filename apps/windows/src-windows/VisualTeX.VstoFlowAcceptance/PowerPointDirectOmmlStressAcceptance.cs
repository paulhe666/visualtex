using System.Diagnostics;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using VisualTeX.PowerPointVsto;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunPowerPointDirectOmmlStressAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        try
        {
            var wordProcessesBefore = Process.GetProcessesByName("WINWORD")
                .Select(process => process.Id)
                .ToHashSet();
            application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            var service = new PowerPointFormulaService(application);
            var timings = new List<double>();

            for (var iteration = 1; iteration <= 100; iteration++)
            {
                var formulaId = Guid.NewGuid().ToString();
                var session = CreatePowerPointAcceptanceSession(
                    "create",
                    "wordOmml",
                    formulaId,
                    sourceObjectId: null,
                    originalMetadata: null);
                var stopwatch = Stopwatch.StartNew();
                var result = service.InsertOmml(session);
                stopwatch.Stop();
                timings.Add(stopwatch.Elapsed.TotalMilliseconds);

                shape = slide.Shapes[result.ObjectId];
                AssertTrue(IsPowerPointNativeEquation(shape),
                    $"Direct PowerPoint MathML import did not create native Office Math at iteration {iteration}.");
                AssertPowerPointNativeOmmlHasNoVisualTeXIdentity(shape, formulaId);
                var standaloneFontSize = service.GetSelectedFormulaFontSize();
                AssertTrue(
                    standaloneFontSize.HasValue
                    && Math.Abs(standaloneFontSize.Value - 20f) < 0.01f,
                    $"Standalone OMML did not honor the persistent/session formula size at iteration {iteration}. Actual={standaloneFontSize?.ToString() ?? "<null>"}");
                shape.Delete();
                Release(shape);
                shape = null;
            }

            // Existing-text-box insertion must modify only the captured text range.
            // It must not create a new shape, replace the text box, or attach any
            // VisualTeX formula identity to the host text box.
            shape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                80,
                80,
                620,
                160);
            var originalShapeCount = slide.Shapes.Count;
            shape.AlternativeText = "user-authored-alt-text";
            shape.Tags.Add("UserTag", "keep-me");
            object? textFrame = null;
            object? fullRange = null;
            object? insertionRange = null;
            object? beforeRange = null;
            object? afterRange = null;
            object? beforeFont = null;
            object? afterFont = null;
            try
            {
                textFrame = shape.TextFrame2;
                fullRange = ((dynamic)textFrame).TextRange;
                ((dynamic)fullRange).Text = "before after";
                beforeRange = ((dynamic)fullRange).Characters(1, 6);
                beforeFont = ((dynamic)beforeRange).Font;
                ((dynamic)beforeFont).Name = "Arial";
                ((dynamic)beforeFont).Size = 18f;
                afterRange = ((dynamic)fullRange).Characters(7, 6);
                afterFont = ((dynamic)afterRange).Font;
                ((dynamic)afterFont).Name = "Calibri";
                ((dynamic)afterFont).Size = 18f;
                Release(afterFont);
                afterFont = null;
                Release(afterRange);
                afterRange = null;
                Release(beforeFont);
                beforeFont = null;
                Release(beforeRange);
                beforeRange = null;
                insertionRange = ((dynamic)fullRange).Characters(7, 0);
                ((dynamic)insertionRange).Select();

                var inlineTarget = service.CaptureTextInsertionTarget();
                var inlineSession = CreatePowerPointAcceptanceSession(
                    "create",
                    "wordOmml",
                    Guid.NewGuid().ToString(),
                    inlineTarget.ObjectId,
                    originalMetadata: null,
                    displayMode: "inline");
                inlineSession.SourceDocumentId = inlineTarget.DocumentId;
                service.InsertOmml(inlineSession);

                AssertEqual(originalShapeCount, slide.Shapes.Count,
                    "Inline OMML insertion created a replacement/new shape instead of staying in the existing text box.");
                Release(fullRange);
                fullRange = ((dynamic)textFrame).TextRange;
                var inlineLength = Convert.ToInt32(((dynamic)fullRange).Length);
                beforeRange = ((dynamic)fullRange).Characters(1, 6);
                beforeFont = ((dynamic)beforeRange).Font;
                afterRange = ((dynamic)fullRange).Characters(inlineLength - 5, 6);
                afterFont = ((dynamic)afterRange).Font;
                AssertEqual("Arial", Convert.ToString(((dynamic)beforeFont).Name),
                    "Inline OMML insertion changed the font before the formula.");
                AssertTrue(Math.Abs(Convert.ToDouble(((dynamic)beforeFont).Size) - 18.0) < 0.01,
                    "Inline OMML insertion changed the font size before the formula.");
                AssertEqual("Calibri", Convert.ToString(((dynamic)afterFont).Name),
                    "Inline OMML insertion changed the font after the formula.");
                AssertTrue(Math.Abs(Convert.ToDouble(((dynamic)afterFont).Size) - 18.0) < 0.01,
                    "Inline OMML insertion changed the font size after the formula.");
                Release(afterFont);
                afterFont = null;
                Release(afterRange);
                afterRange = null;
                Release(beforeFont);
                beforeFont = null;
                Release(beforeRange);
                beforeRange = null;
                AssertTrue(IsPowerPointNativeEquation(shape),
                    "Inline OMML insertion did not create a native math zone in the existing text box.");
                var inlineFormulaFontSize = service.GetSelectedFormulaFontSize();
                AssertTrue(
                    inlineFormulaFontSize.HasValue
                    && Math.Abs(inlineFormulaFontSize.Value - 18f) < 0.01f,
                    $"Inline OMML did not inherit the 18 pt body-text size; persistent/session 20 pt leaked into mixed text. Actual={inlineFormulaFontSize?.ToString() ?? "<null>"}");
                var inlineText = Convert.ToString(((dynamic)fullRange).Text) ?? string.Empty;
                AssertContains(inlineText, "before",
                    "Inline OMML insertion removed text before the captured caret.");
                AssertContains(inlineText, "after",
                    "Inline OMML insertion removed text after the captured caret.");
                AssertEqual("user-authored-alt-text", shape.AlternativeText,
                    "Inline OMML insertion changed user-authored alternative text.");
                AssertEqual("keep-me", shape.Tags["UserTag"],
                    "Inline OMML insertion changed unrelated PowerPoint Shape Tags.");

                Release(insertionRange);
                insertionRange = null;
                Release(fullRange);
                fullRange = ((dynamic)textFrame).TextRange;
                ((dynamic)fullRange).Text = "left right";
                beforeFont = ((dynamic)fullRange).Font;
                ((dynamic)beforeFont).Size = 24f;
                Release(beforeFont);
                beforeFont = null;
                insertionRange = ((dynamic)fullRange).Characters(5, 0);
                ((dynamic)insertionRange).Select();

                var displayTarget = service.CaptureTextInsertionTarget();
                var displaySession = CreatePowerPointAcceptanceSession(
                    "create",
                    "wordOmml",
                    Guid.NewGuid().ToString(),
                    displayTarget.ObjectId,
                    originalMetadata: null,
                    displayMode: "block");
                displaySession.SourceDocumentId = displayTarget.DocumentId;
                service.InsertOmml(displaySession);

                AssertEqual(originalShapeCount, slide.Shapes.Count,
                    "Display OMML insertion created a replacement/new shape instead of staying in the existing text box.");
                AssertTrue(IsPowerPointNativeEquation(shape),
                    "Display OMML insertion did not create a native math zone in the existing text box.");
                var displayFormulaFontSize = service.GetSelectedFormulaFontSize();
                AssertTrue(
                    displayFormulaFontSize.HasValue
                    && Math.Abs(displayFormulaFontSize.Value - 24f) < 0.01f,
                    $"Display OMML did not inherit the 24 pt body-text size; persistent/session 20 pt leaked into mixed text. Actual={displayFormulaFontSize?.ToString() ?? "<null>"}");
                var displayText = Convert.ToString(((dynamic)fullRange).Text) ?? string.Empty;
                AssertContains(displayText, "left",
                    "Display OMML insertion removed text before the captured caret.");
                AssertContains(displayText, "right",
                    "Display OMML insertion removed text after the captured caret.");
                AssertTrue(displayText.IndexOf("\r", StringComparison.Ordinal) >= 0,
                    "Display OMML insertion did not create a separate equation paragraph.");
                AssertEqual("user-authored-alt-text", shape.AlternativeText,
                    "Display OMML insertion changed user-authored alternative text.");
                AssertEqual("keep-me", shape.Tags["UserTag"],
                    "Display OMML insertion changed unrelated PowerPoint Shape Tags.");
            }
            finally
            {
                Release(afterFont);
                Release(afterRange);
                Release(beforeFont);
                Release(beforeRange);
                Release(insertionRange);
                Release(fullRange);
                Release(textFrame);
            }
            shape.Delete();
            Release(shape);
            shape = null;

            // The optimized kernel keeps one private hidden Word host warm while
            // PowerPoint is alive, rather than starting and stopping WINWORD for
            // every formula. Disposing the service must release that host fully.
            service.Dispose();
            var shutdownWatch = Stopwatch.StartNew();
            int[] spawnedWordProcesses;
            do
            {
                spawnedWordProcesses = Process.GetProcessesByName("WINWORD")
                    .Select(process => process.Id)
                    .Where(processId => !wordProcessesBefore.Contains(processId))
                    .ToArray();
                if (spawnedWordProcesses.Length == 0)
                    break;
                Thread.Sleep(20);
            }
            while (shutdownWatch.Elapsed < TimeSpan.FromSeconds(2));
            AssertEqual(0, spawnedWordProcesses.Length,
                "PowerPoint direct OMML conversion left its private Word host running after disposal.");

            var coldStart = timings[0];
            var warmOrdered = timings.Skip(1).OrderBy(value => value).ToArray();
            var warmP50 = warmOrdered[warmOrdered.Length / 2];
            var warmP95 = warmOrdered[(int)Math.Floor((warmOrdered.Length - 1) * 0.95)];
            var warmMaximum = warmOrdered[warmOrdered.Length - 1];
            AssertTrue(warmP95 < 750.0,
                $"PowerPoint native OMML warm-path p95 regressed to {warmP95:F1} ms.");
            Console.WriteLine(
                $"PowerPoint direct OMML stress passed: 100/100 pure native MathML imports with no persisted VisualTeX identity, " +
                $"inline/display insertion stayed inside the existing text box, private Word host shut down cleanly, " +
                $"cold={coldStart:F1} ms, warm p50={warmP50:F1} ms, warm p95={warmP95:F1} ms, warm max={warmMaximum:F1} ms.");
        }
        finally
        {
            Release(shape);
            Release(slide);
            if (presentation is not null)
            {
                try { presentation.Close(); } catch { }
            }
            Release(presentation);
            if (application is not null)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void AssertPowerPointNativeOmmlHasNoVisualTeXIdentity(
        PowerPoint.Shape shape,
        string forbiddenFormulaId)
    {
        AssertTrue(
            !shape.Name.StartsWith("VisualTeX_", StringComparison.OrdinalIgnoreCase),
            "Native PowerPoint OMML retained a VisualTeX_<FormulaId> Shape name.");
        AssertTrue(
            string.IsNullOrWhiteSpace(shape.AlternativeText)
            || shape.AlternativeText.IndexOf(
                forbiddenFormulaId,
                StringComparison.OrdinalIgnoreCase) < 0,
            "Native PowerPoint OMML retained FormulaId metadata in AlternativeText.");
        AssertTrue(
            string.IsNullOrWhiteSpace(shape.Tags["VisualTeXFormulaId"]),
            "Native PowerPoint OMML retained VisualTeXFormulaId Shape Tag.");
        AssertTrue(
            string.IsNullOrWhiteSpace(shape.Tags["VisualTeXMetadata"]),
            "Native PowerPoint OMML retained VisualTeXMetadata Shape Tag.");
        AssertTrue(
            string.IsNullOrWhiteSpace(shape.Tags["VisualTeXIdentityOwner"]),
            "Native PowerPoint OMML retained VisualTeXIdentityOwner Shape Tag.");
    }
}
