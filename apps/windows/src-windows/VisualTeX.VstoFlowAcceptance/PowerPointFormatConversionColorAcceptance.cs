using System.Drawing;
using Extensibility;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using VisualTeX.PowerPointVsto;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // Colors the audience sees must survive a batch conversion: on a dark theme
    // an equation that inherits white text stays white as MathType/VisualTeX,
    // explicit colors stay explicit, and black MathType/VisualTeX equations do
    // not turn into theme-colored (white) native equations.
    private static void RunPowerPointFormatConversionColorAcceptance(string artifactRoot)
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            var scheme = presentation.SlideMaster.Theme.ThemeColorScheme;
            scheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark1).RGB = AcceptanceRgb(255, 255, 255);
            scheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight1).RGB = AcceptanceRgb(0x1F, 0x2A, 0x44);
            var slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            presentation.Windows[1].Activate();

            var inherited = AddAcceptanceNativeEquation(application, slide, new[] { "a^2+b^2=c^2" }, 60, 50, 420, 80, 36);
            inherited.Name = "inherited";
            var red = AddAcceptanceNativeEquation(application, slide, new[] { "x+y" }, 60, 200, 420, 80, 36);
            red.Name = "red";
            red.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = AcceptanceRgb(255, 0, 0);
            var partial = AddAcceptanceNativeEquation(application, slide, new[] { "p+q+r" }, 60, 350, 420, 80, 36);
            partial.Name = "partial";
            var partialText = partial.TextFrame2.TextRange;
            var last = partialText.Characters[partialText.Length, 1];
            last.Font.Fill.ForeColor.RGB = AcceptanceRgb(255, 0, 0);
            Release(last); Release(partialText);
            Release(inherited); Release(red); Release(partial);

            var bands = new Dictionary<string, RectangleF>
            {
                ["inherited"] = new RectangleF(40, 40, 460, 100),
                ["red"] = new RectangleF(40, 190, 460, 100),
                ["partial"] = new RectangleF(40, 340, 460, 100),
            };
            presentation.SaveCopyAs(Path.Combine(artifactRoot, "0-fixture.pptx"));
            var fixtureInk = MeasureAcceptanceInk(presentation, slide, bands, Path.Combine(artifactRoot, "0-fixture.png"), "0 fixture");
            AssertColorBand(fixtureInk["inherited"], white: true, red: false, black: false, "fixture inherited");
            AssertColorBand(fixtureInk["red"], white: false, red: true, black: false, "fixture red");
            AssertColorBand(fixtureInk["partial"], white: true, red: true, black: false, "fixture partial");

            addIn = new ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);
            SetAcceptanceVisualTeXForm(addIn, "crossPlatformPicture");

            void Step(string tag, string label, string file)
            {
                var outcome = RunAcceptanceConversion(addIn!, tag, label);
                AssertEqual(3, outcome.Converted, label + " converted count");
                presentation!.SaveCopyAs(Path.Combine(artifactRoot, file + ".pptx"));
                var ink = MeasureAcceptanceInk(presentation, slide, bands, Path.Combine(artifactRoot, file + ".png"), label);
                AssertColorBand(ink["inherited"], white: true, red: false, black: false, label + " inherited");
                AssertColorBand(ink["red"], white: false, red: true, black: false, label + " red");
                AssertColorBand(ink["partial"], white: true, red: true, black: false, label + " partial");
            }
            Step("wordOmml|mathTypeOle|slide", "1 OMML -> MathType", "1-mathtype");
            Step("mathTypeOle|wordOmml|slide", "2 MathType -> OMML", "2-omml");
            Step("wordOmml|visualtex|slide", "3 OMML -> VisualTeX SVG", "3-visualtex");
            Step("visualtex|mathTypeOle|slide", "4 VisualTeX -> MathType", "4-mathtype");
            Step("mathTypeOle|visualtex|slide", "5 MathType -> VisualTeX SVG", "5-visualtex");
            Step("visualtex|wordOmml|slide", "6 VisualTeX -> OMML", "6-omml");

            // A black equation stays black when it becomes a native equation on
            // a dark slide (its new text box would otherwise inherit white).
            var black = AddAcceptanceNativeEquation(application, slide, new[] { "k^2" }, 520, 200, 300, 80, 36);
            black.Name = "black";
            black.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = AcceptanceRgb(0, 0, 0);
            Release(black);
            var blackBand = new Dictionary<string, RectangleF> { ["black"] = new RectangleF(500, 190, 340, 100) };
            var outcome2 = RunAcceptanceConversion(addIn, "wordOmml|mathTypeOle|slide", "7 black OMML -> MathType");
            AssertEqual(4, outcome2.Converted, "7 converted count");
            var ink7 = MeasureAcceptanceInk(presentation, slide, blackBand, Path.Combine(artifactRoot, "7-mathtype.png"), "7 black OMML -> MathType");
            AssertColorBand(ink7["black"], white: false, red: false, black: true, "7 black");
            outcome2 = RunAcceptanceConversion(addIn, "mathTypeOle|wordOmml|slide", "8 black MathType -> OMML");
            AssertEqual(4, outcome2.Converted, "8 converted count");
            var ink8 = MeasureAcceptanceInk(presentation, slide, blackBand, Path.Combine(artifactRoot, "8-omml.png"), "8 black MathType -> OMML");
            AssertColorBand(ink8["black"], white: false, red: false, black: true, "8 black");
            presentation.SaveCopyAs(Path.Combine(artifactRoot, "8-omml.pptx"));
            Console.WriteLine("PASS PowerPoint batch conversion color acceptance");
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

    private static int AcceptanceRgb(int r, int g, int b) => r | (g << 8) | (b << 16);

    private sealed class AcceptanceInk
    {
        internal int White, Red, Black, Other;
        public override string ToString() => $"white={White} red={Red} black={Black} other={Other}";
    }

    // Exports the slide and classifies every pixel that differs from the slide
    // background inside each band (slide points) as white, red, black or other.
    private static Dictionary<string, AcceptanceInk> MeasureAcceptanceInk(PowerPoint.Presentation presentation,
        PowerPoint.Slide slide, Dictionary<string, RectangleF> bands, string png, string label)
    {
        const int width = 1600;
        var scale = width / presentation.PageSetup.SlideWidth;
        var height = (int)Math.Round(presentation.PageSetup.SlideHeight * scale);
        slide.Export(png, "PNG", width, height);
        var result = new Dictionary<string, AcceptanceInk>();
        using var bitmap = new Bitmap(png);
        var background = bitmap.GetPixel(2, 2);
        foreach (var band in bands)
        {
            var ink = new AcceptanceInk();
            var x0 = Math.Max(0, (int)(band.Value.Left * scale));
            var y0 = Math.Max(0, (int)(band.Value.Top * scale));
            var x1 = Math.Min(bitmap.Width, (int)(band.Value.Right * scale));
            var y1 = Math.Min(bitmap.Height, (int)(band.Value.Bottom * scale));
            for (var y = y0; y < y1; y++)
            {
                for (var x = x0; x < x1; x++)
                {
                    var c = bitmap.GetPixel(x, y);
                    if (Math.Abs(c.R - background.R) + Math.Abs(c.G - background.G) + Math.Abs(c.B - background.B) < 90) continue;
                    if (c.R > 200 && c.G > 200 && c.B > 200) ink.White++;
                    else if (c.R > 180 && c.G < 90 && c.B < 90) ink.Red++;
                    else if (c.R < 50 && c.G < 50 && c.B < 50) ink.Black++;
                    else ink.Other++;
                }
            }
            result[band.Key] = ink;
            Console.WriteLine($"  {label} [{band.Key}] {ink}");
        }
        return result;
    }

    private static void AssertColorBand(AcceptanceInk ink, bool white, bool red, bool black, string label)
    {
        const int present = 40;
        AssertTrue((ink.White >= present) == white, $"{label}: white ink {(white ? "missing" : "unexpected")} ({ink}).");
        AssertTrue((ink.Red >= present) == red, $"{label}: red ink {(red ? "missing" : "unexpected")} ({ink}).");
        AssertTrue((ink.Black >= present) == black, $"{label}: black ink {(black ? "missing" : "unexpected")} ({ink}).");
    }
}
