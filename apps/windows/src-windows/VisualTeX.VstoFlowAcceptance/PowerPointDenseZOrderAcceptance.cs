using System.Reflection;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using VisualTeX.PowerPointVsto;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunPowerPointDenseZOrderAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? target = null;
        try
        {
            application = new PowerPoint.Application();
            application.Visible = Office.MsoTriState.msoTrue;
            presentation = application.Presentations.Add(Office.MsoTriState.msoFalse);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);

            for (var index = 0; index < 620; index++)
            {
                PowerPoint.Shape? filler = null;
                try
                {
                    filler = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        1 + index % 30,
                        1 + index % 20,
                        2,
                        2);
                }
                finally { Release(filler); }
            }

            target = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeOval,
                120,
                80,
                24,
                24);
            AssertTrue(target.ZOrderPosition > 512,
                $"Dense PowerPoint fixture did not place the target beyond the legacy 512-shape ceiling: {target.ZOrderPosition}.");

            var moveToZOrder = typeof(PowerPointFormulaService).GetMethod(
                "MoveToZOrder",
                BindingFlags.Static | BindingFlags.NonPublic)
                ?? throw new MissingMethodException(
                    typeof(PowerPointFormulaService).FullName,
                    "MoveToZOrder");
            moveToZOrder.Invoke(null, new object[] { target, 3 });
            AssertEqual(3, target.ZOrderPosition,
                "PowerPoint formula replacement could not restore a z-order position beyond 512 shapes.");

            var outputPath = Path.Combine(
                artifactRoot,
                "PowerPoint-Dense-ZOrder.pptx");
            presentation.SaveAs(
                outputPath,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                Office.MsoTriState.msoFalse);
            Console.WriteLine(
                $"[POWERPOINT DENSE ZORDER] Restored target from >512 to position 3 among {slide.Shapes.Count} shapes. path={outputPath}");
        }
        finally
        {
            Release(target);
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
}
