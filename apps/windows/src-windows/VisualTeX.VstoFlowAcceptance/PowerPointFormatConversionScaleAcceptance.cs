using System.Diagnostics;
using Extensibility;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using VisualTeX.PowerPointVsto;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // A whole-deck batch: 6 slides x 10 native equations unified to MathType,
    // VisualTeX and back to native. Checks every equation converts, stays on its
    // slide and centre, and that the Ribbon callback returns at once (the batch
    // must not hold PowerPoint's UI thread from start to finish).
    private static void RunPowerPointFormatConversionScaleAcceptance(string artifactRoot)
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        const int slideCount = 6;
        const int perSlide = 10;
        var linear = new[]
        {
            "a^2+b^2=c^2", "e^(iπ)+1=0", "x_1+x_2", "(a+b)^2", "√(x+1)",
            "f(x)=sin x", "1/(1+x)", "∑_(i=1)^n i", "x^(n+1)", "y=kx+b",
        };
        try
        {
            application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            for (var s = 1; s <= slideCount; s++)
                Release(presentation.Slides.Add(s, PowerPoint.PpSlideLayout.ppLayoutBlank));
            presentation.Windows[1].Activate();
            var centers = new List<(int Slide, float X, float Y)>();
            for (var s = 1; s <= slideCount; s++)
            {
                var slide = presentation.Slides[s];
                for (var k = 0; k < perSlide; k++)
                {
                    var shape = AddAcceptanceNativeEquation(application, slide, new[] { linear[(k + s) % linear.Length] },
                        40 + (k % 2) * 440, 20 + (k / 2) * 100, 400, 80, 24);
                    dynamic zone = ((dynamic)shape.TextFrame2.TextRange).MathZones(1, 1);
                    centers.Add((s, (float)zone.BoundLeft + (float)zone.BoundWidth / 2f, (float)zone.BoundTop + (float)zone.BoundHeight / 2f));
                    Release(shape);
                }
                Release(slide);
            }
            presentation.SaveCopyAs(Path.Combine(artifactRoot, "0-fixture.pptx"));

            addIn = new ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);
            SetAcceptanceVisualTeXForm(addIn, "nativeOle");

            void Run(string tag, string label, Func<PowerPoint.Shape, bool> isTarget)
            {
                var watch = Stopwatch.StartNew();
                addIn!.OnConvertFormulaFormat(new AcceptanceRibbonControl(tag));
                var returnedMs = watch.ElapsedMilliseconds;
                var outcome = WaitForAcceptanceConversion(addIn, label);
                Console.WriteLine($"  {label}: callback returned after {returnedMs} ms, batch {watch.ElapsedMilliseconds} ms");
                AssertTrue(outcome.Failure is null, label + " failed: " + outcome.Failure);
                AssertEqual(slideCount * perSlide, outcome.Converted, label + " converted count");
                AssertTrue(returnedMs < 1500, label + ": the Ribbon callback held PowerPoint's UI thread for the whole batch.");
                for (var s = 1; s <= slideCount; s++)
                {
                    var slide = presentation!.Slides[s];
                    var count = 0;
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        try { if (isTarget(shape)) count++; }
                        finally { Release(shape); }
                    }
                    AssertEqual(perSlide, count, $"{label}: slide {s} target count");
                    AssertCentersMatch(slide, isTarget, centers.Where(c => c.Slide == s).Select(c => (c.X, c.Y)).ToList(), 3f,
                        $"{label}: equations moved on slide {s}");
                    Release(slide);
                }
                presentation!.SaveCopyAs(Path.Combine(artifactRoot, label.Split(' ')[0] + ".pptx"));
            }
            Run("any|mathTypeOle|presentation", "1 unify -> MathType", IsAcceptanceMathType);
            Run("any|visualtex|presentation", "2 unify -> VisualTeX OLE", IsAcceptanceVisualTeXOle);
            Run("any|wordOmml|presentation", "3 unify -> OMML", IsAcceptanceNativeMath);
            Console.WriteLine("PASS PowerPoint whole-deck conversion scale acceptance");
        }
        catch
        {
            try { presentation?.SaveCopyAs(Path.Combine(artifactRoot, "failure.pptx")); } catch { }
            throw;
        }
        finally
        {
            try { addIn?.OnDisconnection(ext_DisconnectMode.ext_dm_HostShutdown, ref custom); } catch { }
            if (presentation is not null)
            {
                try { presentation.Saved = MsoTriState.msoTrue; presentation.Close(); } catch { }
            }
            Release(presentation);
            Release(application);
        }
    }
}
