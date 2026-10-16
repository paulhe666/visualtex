using System.Diagnostics;
using System.Reflection;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Win32;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using VisualTeX.PowerPointVsto;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class AcceptanceRibbonControl : IRibbonControl
    {
        internal AcceptanceRibbonControl(string tag) => Tag = tag;
        public string Id => "VisualTeX.Acceptance";
        public object Context => null!;
        public string Tag { get; }
    }

    // The add-in's batch outcome, read through reflection (its types are internal).
    private sealed class AcceptanceConversionOutcome
    {
        internal int Converted, Total;
        internal bool Cancelled;
        internal string? Failure;
        internal List<(string Reason, int Count)> Skips = new();
    }

    private const BindingFlags AcceptanceMembers = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;

    private static Task? AcceptanceConversionTask(ThisAddIn addIn) =>
        typeof(ThisAddIn).GetProperty("FormatConversionTask", AcceptanceMembers)!.GetValue(addIn) as Task;

    private static void RequestAcceptanceStop(ThisAddIn addIn) =>
        typeof(ThisAddIn).GetMethod("RequestFormatConversionStop", AcceptanceMembers)!.Invoke(addIn, null);

    private static AcceptanceConversionOutcome? ReadAcceptanceOutcome(ThisAddIn addIn)
    {
        var outcome = typeof(ThisAddIn).GetProperty("LastFormatConversionOutcome", AcceptanceMembers)!.GetValue(addIn);
        if (outcome is null) return null;
        T Read<T>(object owner, string name) => (T)owner.GetType().GetProperty(name, AcceptanceMembers)!.GetValue(owner)!;
        var result = new AcceptanceConversionOutcome
        {
            Converted = Read<int>(outcome, "Converted"),
            Total = Read<int>(outcome, "Total"),
            Cancelled = Read<bool>(outcome, "Cancelled"),
            Failure = (string?)outcome.GetType().GetProperty("Failure", AcceptanceMembers)!.GetValue(outcome),
        };
        foreach (var skip in Read<System.Collections.IEnumerable>(outcome, "Skips"))
            result.Skips.Add((Read<object>(skip, "Reason").ToString()!, Read<int>(skip, "Count")));
        return result;
    }

    // Ribbon -> plan -> render/encode -> PowerPoint write for batch format
    // conversion, on a scratch presentation in the running PowerPoint, through
    // the real add-in object and the installed VisualTeX companion.
    private static void RunPowerPointFormatConversionAcceptance(string artifactRoot)
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            var slide1 = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            var slide2 = presentation.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutText);
            var slide3 = presentation.Slides.Add(3, PowerPoint.PpSlideLayout.ppLayoutBlank);
            presentation.Windows[1].Activate();

            // Slide 1: a single display equation (animated), an equation-only box
            // with two equations, an equation inside a line of text and a hidden one.
            var standalone = AddAcceptanceNativeEquation(application, slide1, new[] { "a/b+c^2" }, 120, 60, 320, 70, 28);
            standalone.Name = "standalone";
            var split = AddAcceptanceNativeEquation(application, slide1, new[] { "x=1/2", "y=2" }, 120, 200, 320, 130, 24);
            split.Name = "split";
            var mixed = slide1.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 500, 60, 360, 50);
            mixed.Name = "mixed";
            mixed.TextFrame.TextRange.Text = "Let x^2 be large";
            MakeAcceptanceEquation(application, mixed, 5, 3);
            var hidden = AddAcceptanceNativeEquation(application, slide1, new[] { "e^x" }, 520, 220, 200, 60, 24);
            hidden.Name = "hidden";
            hidden.Visible = MsoTriState.msoFalse;
            var effect = slide1.TimeLine.MainSequence.AddEffect(standalone, PowerPoint.MsoAnimEffect.msoAnimEffectFade);
            Release(effect);
            // Triggers: clicking "button" spins the standalone equation; clicking the
            // equation spins "flag". Both must follow the equation into every format.
            var button = slide1.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 20, 440, 80, 30);
            button.Name = "button";
            var flag = slide1.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 140, 440, 30, 30);
            flag.Name = "flag";
            var byButton = slide1.TimeLine.InteractiveSequences.Add();
            Release(byButton.AddTriggerEffect(standalone, PowerPoint.MsoAnimEffect.msoAnimEffectSpin,
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnShapeClick, button));
            var byEquation = slide1.TimeLine.InteractiveSequences.Add();
            Release(byEquation.AddTriggerEffect(flag, PowerPoint.MsoAnimEffect.msoAnimEffectSpin,
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnShapeClick, standalone));
            Release(byEquation); Release(byButton); Release(flag); Release(button);

            // Slide 2: an equation that is the whole body placeholder.
            application.ActiveWindow.View.GotoSlide(2);
            var title = slide2.Shapes.Placeholders[1];
            title.TextFrame.TextRange.Text = "Placeholder title";
            var body = slide2.Shapes.Placeholders[2];
            body.TextFrame.TextRange.Text = "x^3";
            MakeAcceptanceEquation(application, body, 1, 3);
            Release(body);
            Release(title);

            // Slide 3: an equation inside a group.
            var grouped = AddAcceptanceNativeEquation(application, slide3, new[] { "z^2" }, 100, 100, 200, 60, 24);
            grouped.Name = "groupedEquation";
            var frame = slide3.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 90, 90, 220, 80);
            frame.Name = "groupFrame";
            var group = slide3.Shapes.Range(new object[] { "groupedEquation", "groupFrame" }).Group();
            Release(group); Release(frame); Release(grouped);

            application.ActiveWindow.View.GotoSlide(1);
            var sourceCenters = MathZoneCenters(slide1, "standalone", "split", "hidden");
            presentation.SaveCopyAs(Path.Combine(artifactRoot, "0-fixture.pptx"));

            addIn = new ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);
            SetAcceptanceVisualTeXForm(addIn, "nativeOle");

            // 1. Slide: native equations -> MathType.
            var outcome = RunAcceptanceConversion(addIn, "wordOmml|mathTypeOle|slide", "1 slide OMML -> MathType");
            AssertEqual(4, outcome.Converted, "Slide OMML -> MathType did not convert the standalone, both split and the hidden equation.");
            AssertSkip(outcome, "OmmlMixedWithText", 1);
            var counts = CountAcceptanceKinds(slide1);
            AssertEqual(4, counts.MathType, "Slide 1 does not hold four MathType objects.");
            AssertEqual(1, counts.NativeMath, "The equation inside a line of text was not left native.");
            AssertEqual(1, counts.HiddenMathType, "The hidden equation did not stay hidden as MathType.");
            AssertAnimationTarget(slide1, IsAcceptanceMathType, "The fade effect did not follow the converted equation.");
            AssertAcceptanceTriggers(slide1, IsAcceptanceMathType, "1 OMML -> MathType");
            AssertCentersMatch(slide1, IsAcceptanceMathType, sourceCenters, 1.5f, "MathType objects moved away from their equations.");
            AssertEqual("standalone", AcceptanceAnimatedName(slide1), "The converted equation lost its name.");
            AssertTrue(AcceptanceShapeNames(slide1, _ => true).Contains("split (2)")
                    && !AcceptanceShapeNames(slide1, _ => true).Any(name => name?.Contains("BatchFrame") == true),
                "Split equations are not named after their text box.");
            AssertEqual(2, CountAcceptanceKinds(slide2).NativeMath + CountGroupedNativeMath(slide3),
                "A slide-scope conversion touched other slides.");
            presentation.SaveCopyAs(Path.Combine(artifactRoot, "1-omml-to-mathtype.pptx"));

            // 2. Selection: two MathType objects -> VisualTeX OLE.
            var selectedNames = AcceptanceShapeNames(slide1, IsAcceptanceMathType).Where(name => name != null).Take(2).ToArray();
            slide1.Shapes.Range(selectedNames).Select(MsoTriState.msoTrue);
            outcome = RunAcceptanceConversion(addIn, "mathTypeOle|visualtex|selection", "2 selection MathType -> VisualTeX OLE");
            AssertEqual(2, outcome.Converted, "Selection MathType -> VisualTeX converted the wrong number of equations.");
            counts = CountAcceptanceKinds(slide1);
            AssertEqual(2, counts.VisualTeXOle, "Selection conversion did not create two VisualTeX OLE objects.");
            AssertEqual(2, counts.MathType, "Selection conversion changed unselected MathType objects.");
            var selection = application.ActiveWindow.Selection;
            AssertEqual(2, selection.ShapeRange.Count, "The converted objects were not left selected.");
            Release(selection);

            // 3. VisualTeX as SVG: remaining MathType -> SVG pictures (one of them hidden).
            SetAcceptanceVisualTeXForm(addIn, "crossPlatformPicture");
            outcome = RunAcceptanceConversion(addIn, "mathTypeOle|visualtex|slide", "3 slide MathType -> VisualTeX SVG");
            AssertEqual(2, outcome.Converted, "MathType -> VisualTeX SVG converted the wrong number of equations.");
            counts = CountAcceptanceKinds(slide1);
            AssertEqual(2, counts.VisualTeXPicture, "MathType -> VisualTeX SVG did not create two pictures.");
            AssertEqual(1, counts.HiddenVisualTeXPicture, "The hidden equation did not stay hidden as an SVG picture.");
            AssertEqual(0, counts.MathType, "MathType objects remain after the slide conversion.");
            AssertCentersMatch(slide1, shape => IsAcceptanceVisualTeX(shape), sourceCenters, 1.5f, "VisualTeX objects moved away from their equations.");

            // A rotated picture converts to a native equation (which can rotate)
            // but must be reported, not silently straightened, for MathType.
            var rotatedName = AcceptanceShapeNames(slide1, IsAcceptancePicture).First(name => name != null);
            slide1.Shapes[rotatedName].Rotation = 15;

            // 4. Slide: VisualTeX (OLE + SVG) -> native equations.
            outcome = RunAcceptanceConversion(addIn, "visualtex|wordOmml|slide", "4 slide VisualTeX -> OMML");
            AssertEqual(4, outcome.Converted, "VisualTeX -> OMML converted the wrong number of equations.");
            counts = CountAcceptanceKinds(slide1);
            AssertEqual(5, counts.NativeMath, "Slide 1 does not hold four converted native equations plus the text-line equation.");
            AssertEqual(0, counts.VisualTeXOle + counts.VisualTeXPicture, "VisualTeX objects remain after the slide conversion.");
            AssertEqual(1, counts.RotatedNativeMath, "The rotated picture's rotation was not kept on its native equation.");
            AssertAnimationTarget(slide1, IsAcceptanceNativeMath, "The fade effect did not follow into the native equation.");
            AssertAcceptanceTriggers(slide1, IsAcceptanceNativeMath, "4 VisualTeX -> OMML");
            AssertCentersMatch(slide1, shape => IsAcceptanceNativeMath(shape) && shape.Name != "mixed" && Math.Abs(shape.Rotation) < 0.01f,
                sourceCenters, 3f, "Native equations moved away from their source positions.");
            presentation.SaveCopyAs(Path.Combine(artifactRoot, "4-visualtex-to-omml.pptx"));

            // 5. Unify the whole presentation to MathType.
            outcome = RunAcceptanceConversion(addIn, "any|mathTypeOle|presentation", "5 unify presentation -> MathType");
            AssertEqual(4, outcome.Converted, "Unify -> MathType converted the wrong number of equations.");
            AssertSkip(outcome, "OmmlMixedWithText", 1);
            AssertSkip(outcome, "RotatedOrFlipped", 1);
            AssertSkip(outcome, "InGroup", 1);
            counts = CountAcceptanceKinds(slide1);
            AssertEqual(3, counts.MathType, "Unify left convertible native equations on slide 1.");
            var slide2Counts = CountAcceptanceKinds(slide2);
            AssertEqual(1, slide2Counts.MathType, "The placeholder equation was not converted.");
            AssertEqual(1, slide2.Shapes.Placeholders.Count, "An empty content placeholder was left behind.");
            AssertEqual("Placeholder title", slide2.Shapes.Placeholders[1].TextFrame.TextRange.Text, "The title placeholder changed.");
            AssertEqual(1, CountAcceptanceKinds(slide3).NativeMath + CountGroupedNativeMath(slide3), "The grouped equation was changed.");
            AssertAnimationTarget(slide1, IsAcceptanceMathType, "The fade effect did not follow the unified equation.");
            AssertAcceptanceTriggers(slide1, IsAcceptanceMathType, "5 unify -> MathType");
            var unified = Path.Combine(artifactRoot, "5-unify-mathtype.pptx");
            presentation.SaveCopyAs(unified);
            // The animations and triggers are in the saved file, not only in memory.
            var reopened = application.Presentations.Open(unified, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
            try
            {
                var savedSlide = reopened.Slides[1];
                AssertAnimationTarget(savedSlide, IsAcceptanceMathType, "The saved file lost the equation's fade effect.");
                AssertAcceptanceTriggers(savedSlide, IsAcceptanceMathType, "5 saved file");
                Release(savedSlide);
            }
            finally { reopened.Close(); Release(reopened); }
            presentation.Windows[1].Activate();

            // 6. Stop before anything is written leaves the presentation unchanged.
            SetAcceptanceVisualTeXForm(addIn, "nativeOle");
            addIn.OnConvertFormulaFormat(new AcceptanceRibbonControl("any|visualtex|presentation"));
            RequestAcceptanceStop(addIn);
            outcome = WaitForAcceptanceConversion(addIn, "6 stop");
            AssertTrue(outcome.Cancelled, "Stop did not cancel the conversion.");
            AssertEqual(0, outcome.Converted, "Stopping before the write phase still converted equations.");
            AssertEqual(3, CountAcceptanceKinds(slide1).MathType, "A stopped conversion changed slide 1.");

            // 7. Nothing selected: a clear message and no change.
            application.ActiveWindow.Selection.Unselect();
            outcome = RunAcceptanceConversion(addIn, "mathTypeOle|wordOmml|selection", "7 empty selection", allowFailure: true);
            AssertTrue(outcome.Failure is not null && outcome.Converted == 0, "An empty selection did not report a clear error.");
            Console.WriteLine("PASS PowerPoint batch format conversion acceptance");
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

    // The VisualTeX form is a per-user preference; the acceptance sets only the
    // add-in's in-memory choice so the user's saved setting is never touched.
    private static void SetAcceptanceVisualTeXForm(ThisAddIn addIn, string mode) =>
        typeof(ThisAddIn).GetField("_visualTeXTargetMode", AcceptanceMembers)!.SetValue(addIn, mode);

    private static AcceptanceConversionOutcome RunAcceptanceConversion(ThisAddIn addIn, string tag, string label, bool allowFailure = false)
    {
        addIn.OnConvertFormulaFormat(new AcceptanceRibbonControl(tag));
        var outcome = WaitForAcceptanceConversion(addIn, label);
        if (outcome.Failure is not null && !allowFailure)
            throw new InvalidOperationException($"{label} failed: {outcome.Failure}");
        return outcome;
    }

    private static AcceptanceConversionOutcome WaitForAcceptanceConversion(ThisAddIn addIn, string label)
    {
        var task = AcceptanceConversionTask(addIn) ?? throw new InvalidOperationException($"{label} did not start.");
        var watch = Stopwatch.StartNew();
        while (!task.IsCompleted)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(5);
            if (watch.Elapsed > TimeSpan.FromMinutes(4))
                throw new TimeoutException($"{label} did not finish within four minutes.");
        }
        task.GetAwaiter().GetResult();
        var outcome = ReadAcceptanceOutcome(addIn) ?? throw new InvalidOperationException($"{label} reported no outcome.");
        Console.WriteLine($"{label}: converted={outcome.Converted}/{outcome.Total} cancelled={outcome.Cancelled} elapsedMs={watch.ElapsedMilliseconds} skips="
            + string.Join(",", outcome.Skips.Select(skip => $"{skip.Reason}x{skip.Count}"))
            + (outcome.Failure is null ? string.Empty : " failure=" + outcome.Failure.Split('\n')[0]));
        return outcome;
    }

    private static void AssertSkip(AcceptanceConversionOutcome outcome, string reason, int count)
    {
        var actual = outcome.Skips.Where(skip => skip.Reason == reason).Sum(skip => skip.Count);
        AssertEqual(count, actual, $"Expected {count} skip(s) for {reason}.");
    }

    private static PowerPoint.Shape AddAcceptanceNativeEquation(
        PowerPoint.Application application, PowerPoint.Slide slide, string[] lines,
        float left, float top, float width, float height, float fontSize)
    {
        application.ActiveWindow.View.GotoSlide(slide.SlideIndex);
        var shape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
        shape.TextFrame.TextRange.Text = string.Join("\r", lines);
        shape.TextFrame.TextRange.Font.Size = fontSize;
        for (var index = lines.Length; index >= 1; index--)
        {
            var paragraph = shape.TextFrame.TextRange.Paragraphs(index);
            var start = paragraph.Start;
            var length = paragraph.Text.TrimEnd('\r').Length;
            Release(paragraph);
            MakeAcceptanceEquation(application, shape, start, length);
        }
        return shape;
    }

    private static void MakeAcceptanceEquation(PowerPoint.Application application, PowerPoint.Shape shape, int start, int length)
    {
        var before = Convert.ToInt32(((dynamic)shape.TextFrame2.TextRange).MathZones().Count);
        shape.Select(MsoTriState.msoTrue);
        shape.TextFrame.TextRange.Characters(start, length).Select();
        application.CommandBars.ExecuteMso(application.CommandBars.GetEnabledMso("InsertBuildingBlocksEquationsGallery")
            ? "InsertBuildingBlocksEquationsGallery" : "EquationInsertNew");
        PumpAcceptance(300);
        if (application.CommandBars.GetEnabledMso("EquationProfessional")) application.CommandBars.ExecuteMso("EquationProfessional");
        PumpAcceptance(300);
        var after = Convert.ToInt32(((dynamic)shape.TextFrame2.TextRange).MathZones().Count);
        AssertTrue(after == before + 1, $"Fixture equation was not created in {shape.Name}.");
    }

    private static void PumpAcceptance(int milliseconds)
    {
        var watch = Stopwatch.StartNew();
        while (watch.ElapsedMilliseconds < milliseconds)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(10);
        }
    }

    private static List<(float X, float Y)> MathZoneCenters(PowerPoint.Slide slide, params string[] names)
    {
        var centers = new List<(float X, float Y)>();
        foreach (var name in names)
        {
            var shape = slide.Shapes[name];
            var range = shape.TextFrame2.TextRange;
            var count = Convert.ToInt32(((dynamic)range).MathZones().Count);
            for (var index = 1; index <= count; index++)
            {
                dynamic zone = ((dynamic)range).MathZones(index, 1);
                centers.Add(((float)zone.BoundLeft + (float)zone.BoundWidth / 2f, (float)zone.BoundTop + (float)zone.BoundHeight / 2f));
            }
            Release(range);
            Release(shape);
        }
        return centers;
    }

    private static void AssertCentersMatch(PowerPoint.Slide slide, Func<PowerPoint.Shape, bool> select,
        List<(float X, float Y)> expected, float tolerance, string message)
    {
        foreach (PowerPoint.Shape shape in slide.Shapes)
        {
            try
            {
                if (!select(shape)) continue;
                float centerX, centerY;
                if (IsAcceptanceNativeMath(shape))
                {
                    dynamic zone = ((dynamic)shape.TextFrame2.TextRange).MathZones(1, 1);
                    centerX = (float)zone.BoundLeft + (float)zone.BoundWidth / 2f;
                    centerY = (float)zone.BoundTop + (float)zone.BoundHeight / 2f;
                }
                else
                {
                    centerX = shape.Left + shape.Width / 2f;
                    centerY = shape.Top + shape.Height / 2f;
                }
                var nearest = expected.Min(center => Math.Max(Math.Abs(center.X - centerX), Math.Abs(center.Y - centerY)));
                AssertTrue(nearest <= tolerance, $"{message} ({shape.Name} is {nearest:0.00} pt from the nearest source center).");
            }
            finally { Release(shape); }
        }
    }

    private static void AssertAnimationTarget(PowerPoint.Slide slide, Func<PowerPoint.Shape, bool> expected, string message)
    {
        var sequence = slide.TimeLine.MainSequence;
        AssertEqual(1, sequence.Count, "The slide lost or duplicated its animation effect.");
        var effect = sequence[1];
        var target = effect.Shape;
        AssertTrue(expected(target), message);
        Release(target); Release(effect); Release(sequence);
    }

    private static string AcceptanceAnimatedName(PowerPoint.Slide slide)
    {
        var sequence = slide.TimeLine.MainSequence;
        var effect = sequence[1];
        var shape = effect.Shape;
        var name = shape.Name;
        Release(shape); Release(effect); Release(sequence);
        return name;
    }

    private static void AssertAcceptanceTriggers(PowerPoint.Slide slide, Func<PowerPoint.Shape, bool> isEquation, string label)
    {
        var sequences = slide.TimeLine.InteractiveSequences;
        AssertEqual(2, sequences.Count, label + ": trigger sequences");
        var byButton = false;
        var byEquation = false;
        for (var index = 1; index <= sequences.Count; index++)
        {
            var sequence = sequences[index];
            AssertEqual(1, sequence.Count, label + ": effects in a trigger sequence");
            var effect = sequence[1];
            var timing = effect.Timing;
            var trigger = timing.TriggerShape;
            var animated = effect.Shape;
            if (trigger.Name == "button" && isEquation(animated)) byButton = true;
            if (animated.Name == "flag" && isEquation(trigger)) byEquation = true;
            Release(animated); Release(trigger); Release(timing); Release(effect); Release(sequence);
        }
        Release(sequences);
        AssertTrue(byButton, label + ": clicking the button no longer animates the equation.");
        AssertTrue(byEquation, label + ": clicking the equation no longer runs its trigger.");
    }

    private sealed class AcceptanceKindCounts
    {
        internal int MathType, HiddenMathType, VisualTeXOle, VisualTeXPicture, HiddenVisualTeXPicture, NativeMath, RotatedNativeMath;
    }

    private static AcceptanceKindCounts CountAcceptanceKinds(PowerPoint.Slide slide)
    {
        var counts = new AcceptanceKindCounts();
        foreach (PowerPoint.Shape shape in slide.Shapes)
        {
            try
            {
                var hidden = shape.Visible == MsoTriState.msoFalse;
                if (IsAcceptanceMathType(shape)) { counts.MathType++; if (hidden) counts.HiddenMathType++; }
                else if (IsAcceptanceVisualTeXOle(shape)) counts.VisualTeXOle++;
                else if (IsAcceptancePicture(shape)) { counts.VisualTeXPicture++; if (hidden) counts.HiddenVisualTeXPicture++; }
                else if (IsAcceptanceNativeMath(shape))
                {
                    counts.NativeMath += Convert.ToInt32(((dynamic)shape.TextFrame2.TextRange).MathZones().Count);
                    if (Math.Abs(shape.Rotation) > 0.01f) counts.RotatedNativeMath++;
                }
            }
            finally { Release(shape); }
        }
        return counts;
    }

    private static int CountGroupedNativeMath(PowerPoint.Slide slide)
    {
        var count = 0;
        foreach (PowerPoint.Shape shape in slide.Shapes)
        {
            try
            {
                if (shape.Type != MsoShapeType.msoGroup) continue;
                foreach (PowerPoint.Shape item in shape.GroupItems)
                {
                    try { if (IsAcceptanceNativeMath(item)) count++; }
                    finally { Release(item); }
                }
            }
            finally { Release(shape); }
        }
        return count;
    }

    private static string?[] AcceptanceShapeNames(PowerPoint.Slide slide, Func<PowerPoint.Shape, bool> select)
    {
        var names = new List<string?>();
        foreach (PowerPoint.Shape shape in slide.Shapes)
        {
            try { if (select(shape) && shape.Visible == MsoTriState.msoTrue) names.Add(shape.Name); }
            finally { Release(shape); }
        }
        return names.ToArray();
    }

    private static bool IsAcceptanceMathType(PowerPoint.Shape shape) =>
        shape.Type == MsoShapeType.msoEmbeddedOLEObject
        && (shape.OLEFormat.ProgID ?? string.Empty).StartsWith("Equation.DSMT", StringComparison.OrdinalIgnoreCase);

    private static bool IsAcceptanceVisualTeXOle(PowerPoint.Shape shape) =>
        shape.Type == MsoShapeType.msoEmbeddedOLEObject
        && string.Equals(shape.OLEFormat.ProgID, VisualTeX.WindowsOffice.Contracts.FormulaOleContract.ProgId, StringComparison.OrdinalIgnoreCase);

    private static bool IsAcceptancePicture(PowerPoint.Shape shape) =>
        shape.Type != MsoShapeType.msoEmbeddedOLEObject && DecodePowerPointMetadata(shape) is not null;

    private static bool IsAcceptanceVisualTeX(PowerPoint.Shape shape) =>
        IsAcceptanceVisualTeXOle(shape) || IsAcceptancePicture(shape);

    private static bool IsAcceptanceNativeMath(PowerPoint.Shape shape)
    {
        if (shape.HasTextFrame != MsoTriState.msoTrue) return false;
        try { return Convert.ToInt32(((dynamic)shape.TextFrame2.TextRange).MathZones().Count) > 0; }
        catch { return false; }
    }
}
