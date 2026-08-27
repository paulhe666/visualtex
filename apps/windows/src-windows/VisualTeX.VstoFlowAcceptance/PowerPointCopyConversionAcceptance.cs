using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using WinForms = System.Windows.Forms;
using VisualTeX.PowerPointVsto;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const string PowerPointAcceptanceLatex = @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}";
    private const string PowerPointAcceptanceMathMl =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" +
        "<mfrac><mi>a</mi><mi>b</mi></mfrac><mo>+</mo><msqrt><mi>x</mi></msqrt></math>";

    private static void RunPowerPointCopyConversionAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide1 = null;
        PowerPoint.Slide? slide2 = null;
        PowerPoint.Shape? shape = null;
        PowerPoint.Shape? copiedShape = null;
        PowerPoint.ShapeRange? pastedRange = null;
        VisualTeX.PowerPointVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        var officeTempRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(officeTempRoot);
        var svgPath = Path.Combine(officeTempRoot, $"{Guid.NewGuid():N}.svg");
        var pngPath = Path.Combine(officeTempRoot, $"{Guid.NewGuid():N}.png");
        string? emfPath = null;
        var presentationPath = Path.Combine(artifactRoot, "powerpoint-copy-conversion.pptx");
        var ommlSnapshotPath = Path.Combine(artifactRoot, "powerpoint-native-omml-snapshot.pptx");

        File.WriteAllText(svgPath, CreateSvg(220, 72));
        WriteAcceptancePng(pngPath, PowerPointAcceptanceLatex, 440, 144);

        try
        {
            application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide1 = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            slide2 = presentation.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            var service = new PowerPointFormulaService(application);
            emfPath = CreatePowerPointAcceptanceEmf(svgPath, 220, 72);

            var formulaId = Guid.NewGuid().ToString();
            var createPicture = CreatePowerPointAcceptanceSession(
                mode: "create",
                objectMode: "crossPlatformPicture",
                formulaId: formulaId,
                sourceObjectId: null,
                originalMetadata: null);
            var pictureResult = service.Insert(createPicture, svgPath);
            shape = slide1.Shapes[pictureResult.ObjectId];
            AssertTrue(IsPowerPointEditablePictureShape(shape), "Initial PowerPoint VisualTeX formula was not a picture.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "Initial PowerPoint picture FormulaId mismatch.");

            // Cross-slide copy is the hard case because PowerPoint shape names are
            // only unique per slide. The persisted SlideID:Shape.Id owner token must
            // still distinguish the copied picture from its source.
            shape.Copy();
            application.ActiveWindow.View.GotoSlide(2);
            pastedRange = PastePowerPointShapesWithRetry(slide2);
            copiedShape = pastedRange[1];
            copiedShape.Select(MsoTriState.msoTrue);
            var copiedPictureSelection = service.ReadSelection();
            AssertTrue(!string.IsNullOrWhiteSpace(copiedPictureSelection.FormulaId), "Copied picture was not readable by VisualTeX.");
            AssertTrue(!string.Equals(formulaId, copiedPictureSelection.FormulaId, StringComparison.OrdinalIgnoreCase), "Copied picture reused the source FormulaId.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "Reading the copied picture changed the source FormulaId.");
            var copiedPictureId = copiedPictureSelection.FormulaId!;
            var sourcePictureWidth = shape.Width;
            var copiedPictureWidth = copiedShape.Width;
            service.SetSelectedFormulaFontSize(30);
            AssertNear(sourcePictureWidth, shape.Width, 0.5f, "Formatting the copied picture resized the source picture.");
            AssertTrue(Math.Abs(copiedShape.Width - copiedPictureWidth) > 0.5f, "Formatting the copied picture did not target the copy.");
            var copiedPictureMetadataAfterFormatting = DecodePowerPointMetadata(copiedShape);
            AssertEqual(copiedPictureId, copiedPictureMetadataAfterFormatting?.FormulaId, "Copied picture lost its independent FormulaId after formatting.");
            AssertNear(30f, (float)(copiedPictureMetadataAfterFormatting?.FontSizePt ?? 0), 0.1f, "Copied picture metadata did not persist the requested font size.");
            Console.WriteLine("PowerPoint picture copy identity passed: cross-slide copy received an independent FormulaId and formatting targeted only the copy.");

            copiedShape.Delete();
            Release(copiedShape);
            copiedShape = null;
            Release(pastedRange);
            pastedRange = null;
            application.ActiveWindow.View.GotoSlide(1);
            shape.Select(MsoTriState.msoTrue);

            // Picture -> OLE.
            var sourceMetadata = DecodePowerPointMetadata(shape)
                ?? throw new InvalidDataException("PowerPoint picture metadata disappeared before OLE conversion.");
            var pictureToOle = CreatePowerPointAcceptanceSession(
                "edit",
                "nativeOle",
                formulaId,
                shape.Name,
                sourceMetadata);
            var oleResult = service.ReplaceOle(pictureToOle, pngPath, emfPath);
            Release(shape);
            shape = slide1.Shapes[oleResult.ObjectId];
            AssertPowerPointOle(shape, "Picture -> OLE conversion did not create a VisualTeX OLE object.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "Picture -> OLE changed FormulaId.");

            // Copy the OLE on the same slide. The copy inherits both embedded
            // metadata and shape tags, but the owner token must force a new id.
            shape.Copy();
            pastedRange = PastePowerPointShapesWithRetry(slide1);
            copiedShape = pastedRange[1];
            copiedShape.Select(MsoTriState.msoTrue);
            var copiedOleSelection = service.ReadSelection();
            AssertTrue(!string.IsNullOrWhiteSpace(copiedOleSelection.FormulaId), "Copied OLE was not readable by VisualTeX.");
            AssertTrue(!string.Equals(formulaId, copiedOleSelection.FormulaId, StringComparison.OrdinalIgnoreCase), "Copied OLE reused the source FormulaId.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "Reading the copied OLE changed the source FormulaId.");
            var copiedOleId = copiedOleSelection.FormulaId!;
            var sourceOleWidth = shape.Width;
            var copiedOleWidth = copiedShape.Width;
            service.SetSelectedFormulaFontSize(32);
            AssertNear(sourceOleWidth, shape.Width, 0.5f, "Formatting the copied OLE resized the source OLE.");
            AssertTrue(Math.Abs(copiedShape.Width - copiedOleWidth) > 0.5f, "Formatting the copied OLE did not target the copy.");
            var copiedOleMetadataAfterFormatting = DecodePowerPointMetadata(copiedShape);
            AssertEqual(copiedOleId, copiedOleMetadataAfterFormatting?.FormulaId, "Copied OLE lost its independent FormulaId after formatting.");
            AssertNear(32f, (float)(copiedOleMetadataAfterFormatting?.FontSizePt ?? 0), 0.1f, "Copied OLE metadata did not persist the requested font size.");
            Console.WriteLine("PowerPoint OLE copy identity passed: copied OLE received an independent FormulaId and editing geometry stayed isolated.");

            copiedShape.Delete();
            Release(copiedShape);
            copiedShape = null;
            Release(pastedRange);
            pastedRange = null;
            shape.Select(MsoTriState.msoTrue);

            // OLE -> picture.
            sourceMetadata = DecodePowerPointMetadata(shape)
                ?? throw new InvalidDataException("PowerPoint OLE metadata disappeared before picture conversion.");
            var oleToPicture = CreatePowerPointAcceptanceSession(
                "edit",
                "crossPlatformPicture",
                formulaId,
                shape.Name,
                sourceMetadata);
            var backToPicture = service.Replace(oleToPicture, svgPath);
            Release(shape);
            shape = slide1.Shapes[backToPicture.ObjectId];
            AssertTrue(IsPowerPointEditablePictureShape(shape), "OLE -> picture conversion did not create an SVG picture.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "OLE -> picture changed FormulaId.");

            // Regression: a DPI-scaled EMF frame used to make PowerPoint and the
            // OLE server disagree about the formula's physical extent.  Each
            // picture <-> OLE round trip then inherited the already-expanded box,
            // causing severe vertical stretching and exponential growth.
            var stablePictureWidth = shape.Width;
            var stablePictureHeight = shape.Height;
            for (var iteration = 1; iteration <= 12; iteration++)
            {
                sourceMetadata = DecodePowerPointMetadata(shape)
                    ?? throw new InvalidDataException($"PowerPoint picture metadata disappeared before round trip {iteration}.");
                var roundTripToOle = CreatePowerPointAcceptanceSession(
                    "edit",
                    "nativeOle",
                    formulaId,
                    shape.Name,
                    sourceMetadata);
                var roundTripOleResult = service.ReplaceOle(roundTripToOle, pngPath, emfPath);
                Release(shape);
                shape = slide1.Shapes[roundTripOleResult.ObjectId];
                WinForms.Application.DoEvents();
                Thread.Sleep(180);
                Release(shape);
                shape = slide1.Shapes[roundTripOleResult.ObjectId];
                AssertPowerPointOle(shape, $"Picture -> OLE round trip {iteration} did not create native OLE.");
                AssertNear(stablePictureWidth, shape.Width, 1.5f, $"Picture -> OLE round trip {iteration} drifted in width.");
                AssertNear(stablePictureHeight, shape.Height, 1.5f, $"Picture -> OLE round trip {iteration} drifted in height.");

                sourceMetadata = DecodePowerPointMetadata(shape)
                    ?? throw new InvalidDataException($"PowerPoint OLE metadata disappeared before round trip {iteration}.");
                var roundTripToPicture = CreatePowerPointAcceptanceSession(
                    "edit",
                    "crossPlatformPicture",
                    formulaId,
                    shape.Name,
                    sourceMetadata);
                var roundTripPictureResult = service.Replace(roundTripToPicture, svgPath);
                Release(shape);
                shape = slide1.Shapes[roundTripPictureResult.ObjectId];
                WinForms.Application.DoEvents();
                Thread.Sleep(180);
                Release(shape);
                shape = slide1.Shapes[roundTripPictureResult.ObjectId];
                AssertTrue(IsPowerPointEditablePictureShape(shape), $"OLE -> picture round trip {iteration} did not create SVG.");
                AssertNear(stablePictureWidth, shape.Width, 1.5f, $"OLE -> picture round trip {iteration} drifted in width.");
                AssertNear(stablePictureHeight, shape.Height, 1.5f, $"OLE -> picture round trip {iteration} drifted in height.");
            }
            Console.WriteLine(
                $"PowerPoint OLE/SVG geometry stability passed: 12 round trips stayed at "
                + $"{stablePictureWidth:0.0}x{stablePictureHeight:0.0} pt without cumulative growth.");

            // Existing presentations can already contain a catastrophically
            // enlarged Shape produced by the old DPI feedback loop. One normal
            // conversion must repair that geometry instead of preserving it.
            shape.LockAspectRatio = MsoTriState.msoFalse;
            shape.Width = 9000f;
            shape.Height = 5000f;
            shape.LockAspectRatio = MsoTriState.msoTrue;
            sourceMetadata = DecodePowerPointMetadata(shape)
                ?? throw new InvalidDataException("Corrupted PowerPoint picture lost metadata before recovery.");
            var recoveryToOle = CreatePowerPointAcceptanceSession(
                "edit",
                "nativeOle",
                formulaId,
                shape.Name,
                sourceMetadata);
            var recoveredOle = service.ReplaceOle(recoveryToOle, pngPath, emfPath);
            Release(shape);
            shape = slide1.Shapes[recoveredOle.ObjectId];
            WinForms.Application.DoEvents();
            Thread.Sleep(220);
            Release(shape);
            shape = slide1.Shapes[recoveredOle.ObjectId];
            AssertPowerPointOle(shape, "Corrupted picture recovery did not create native OLE.");
            AssertNear(stablePictureWidth, shape.Width, 1.5f, "Corrupted picture recovery did not restore natural width.");
            AssertNear(stablePictureHeight, shape.Height, 1.5f, "Corrupted picture recovery did not restore natural height.");

            sourceMetadata = DecodePowerPointMetadata(shape)
                ?? throw new InvalidDataException("Recovered PowerPoint OLE lost metadata.");
            var recoveryToPicture = CreatePowerPointAcceptanceSession(
                "edit",
                "crossPlatformPicture",
                formulaId,
                shape.Name,
                sourceMetadata);
            var recoveredPicture = service.Replace(recoveryToPicture, svgPath);
            Release(shape);
            shape = slide1.Shapes[recoveredPicture.ObjectId];
            AssertTrue(IsPowerPointEditablePictureShape(shape), "Recovered OLE -> picture did not restore SVG.");
            AssertNear(stablePictureWidth, shape.Width, 1.5f, "Recovered OLE -> picture changed repaired width.");
            AssertNear(stablePictureHeight, shape.Height, 1.5f, "Recovered OLE -> picture changed repaired height.");
            Console.WriteLine("PowerPoint pathological geometry recovery passed: a 9000x5000 pt corrupted formula returned to its metadata-derived natural size in one conversion.");

            // Picture -> native Office Math / OMML.
            sourceMetadata = DecodePowerPointMetadata(shape)
                ?? throw new InvalidDataException("PowerPoint picture metadata disappeared before OMML conversion.");
            var pictureToOmml = CreatePowerPointAcceptanceSession(
                "edit",
                "wordOmml",
                formulaId,
                shape.Name,
                sourceMetadata);
            var ommlResult = service.ReplaceOmml(pictureToOmml);
            Release(shape);
            shape = slide1.Shapes[ommlResult.ObjectId];
            AssertTrue(IsPowerPointNativeEquation(shape), "Picture -> OMML did not create native PowerPoint Office Math.");
            shape.Select(MsoTriState.msoTrue);
            var nativeSelection = service.ReadSelection();
            AssertEqual("wordOmml", nativeSelection.ObjectMode, "Native PowerPoint equation was not recognized as OMML mode.");
            AssertEqual(formulaId, nativeSelection.FormulaId, "Picture -> OMML changed FormulaId.");
            AssertContains(nativeSelection.Metadata?.Latex, @"\frac", "Native OMML did not round-trip the fraction back to LaTeX.");
            AssertContains(nativeSelection.Metadata?.Latex, @"\sqrt", "Native OMML did not round-trip the radical back to LaTeX.");

            presentation.SaveCopyAs(
                ommlSnapshotPath,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                MsoTriState.msoTrue);
            AssertPowerPointPptxContainsNativeOmml(ommlSnapshotPath);

            // OMML copy gets an independent identity as well.
            shape.Copy();
            pastedRange = PastePowerPointShapesWithRetry(slide1);
            copiedShape = pastedRange[1];
            copiedShape.Select(MsoTriState.msoTrue);
            var copiedOmmlSelection = service.ReadSelection();
            AssertTrue(!string.IsNullOrWhiteSpace(copiedOmmlSelection.FormulaId), "Copied OMML was not readable by VisualTeX.");
            AssertTrue(!string.Equals(formulaId, copiedOmmlSelection.FormulaId, StringComparison.OrdinalIgnoreCase), "Copied OMML reused the source FormulaId.");
            AssertTrue(IsPowerPointNativeEquation(copiedShape), "Copied OMML ceased to be native Office Math.");
            copiedShape.Delete();
            Release(copiedShape);
            copiedShape = null;
            Release(pastedRange);
            pastedRange = null;
            shape.Select(MsoTriState.msoTrue);

            // OMML -> OLE.
            sourceMetadata = service.ReadSelection().Metadata
                ?? throw new InvalidDataException("PowerPoint OMML metadata disappeared before OLE conversion.");
            var ommlToOle = CreatePowerPointAcceptanceSession(
                "edit",
                "nativeOle",
                formulaId,
                shape.Name,
                sourceMetadata);
            var ommlOleResult = service.ReplaceOle(ommlToOle, pngPath, emfPath);
            Release(shape);
            shape = slide1.Shapes[ommlOleResult.ObjectId];
            AssertPowerPointOle(shape, "OMML -> OLE conversion did not create a VisualTeX OLE object.");

            // OLE -> OMML.
            sourceMetadata = DecodePowerPointMetadata(shape)
                ?? throw new InvalidDataException("PowerPoint OLE metadata disappeared before OMML conversion.");
            var oleToOmml = CreatePowerPointAcceptanceSession(
                "edit",
                "wordOmml",
                formulaId,
                shape.Name,
                sourceMetadata);
            var secondOmml = service.ReplaceOmml(oleToOmml);
            Release(shape);
            shape = slide1.Shapes[secondOmml.ObjectId];
            AssertTrue(IsPowerPointNativeEquation(shape), "OLE -> OMML conversion did not create native Office Math.");

            // OMML -> picture.
            shape.Select(MsoTriState.msoTrue);
            sourceMetadata = service.ReadSelection().Metadata
                ?? throw new InvalidDataException("PowerPoint OMML metadata disappeared before picture conversion.");
            var ommlToPicture = CreatePowerPointAcceptanceSession(
                "edit",
                "crossPlatformPicture",
                formulaId,
                shape.Name,
                sourceMetadata);
            var finalPicture = service.Replace(ommlToPicture, svgPath);
            Release(shape);
            shape = slide1.Shapes[finalPicture.ObjectId];
            AssertTrue(IsPowerPointEditablePictureShape(shape), "OMML -> picture conversion did not create an SVG picture.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "Three-way conversion changed the original FormulaId.");

            // Stress the direct PowerPoint MathML importer. This specifically
            // guards against the old transient Word automation path, which could
            // fail intermittently even when a single conversion succeeded.
            var ommlWriteMilliseconds = new List<double>();
            for (var iteration = 1; iteration <= 25; iteration++)
            {
                sourceMetadata = DecodePowerPointMetadata(shape)
                    ?? throw new InvalidDataException($"PowerPoint stress picture metadata disappeared at iteration {iteration}.");
                var stressToOmml = CreatePowerPointAcceptanceSession(
                    "edit",
                    "wordOmml",
                    formulaId,
                    shape.Name,
                    sourceMetadata);
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                var stressOmmlResult = service.ReplaceOmml(stressToOmml);
                stopwatch.Stop();
                ommlWriteMilliseconds.Add(stopwatch.Elapsed.TotalMilliseconds);
                Release(shape);
                shape = slide1.Shapes[stressOmmlResult.ObjectId];
                AssertTrue(IsPowerPointNativeEquation(shape), $"Stress picture -> OMML failed at iteration {iteration}.");
                shape.Select(MsoTriState.msoTrue);
                var stressRead = service.ReadSelection();
                AssertContains(stressRead.Metadata?.Latex, @"\frac", $"Stress OMML fraction readback failed at iteration {iteration}.");
                AssertContains(stressRead.Metadata?.Latex, @"\sqrt", $"Stress OMML radical readback failed at iteration {iteration}.");

                sourceMetadata = stressRead.Metadata
                    ?? throw new InvalidDataException($"PowerPoint stress OMML metadata disappeared at iteration {iteration}.");
                var stressToPicture = CreatePowerPointAcceptanceSession(
                    "edit",
                    "crossPlatformPicture",
                    formulaId,
                    shape.Name,
                    sourceMetadata);
                var stressPictureResult = service.Replace(stressToPicture, svgPath);
                Release(shape);
                shape = slide1.Shapes[stressPictureResult.ObjectId];
                AssertTrue(IsPowerPointEditablePictureShape(shape), $"Stress OMML -> picture failed at iteration {iteration}.");
            }
            var orderedOmmlWrites = ommlWriteMilliseconds.OrderBy(value => value).ToArray();
            var ommlP50 = orderedOmmlWrites[orderedOmmlWrites.Length / 2];
            var ommlMax = orderedOmmlWrites[orderedOmmlWrites.Length - 1];
            Console.WriteLine(
                $"PowerPoint direct MathML stress passed: 25/25 picture -> OMML -> readback -> picture cycles; " +
                $"OMML write p50={ommlP50:F1} ms, max={ommlMax:F1} ms.");

            // Recreate the exact metadata/geometry reported by the user's fresh
            // quadratic formula: raw code, 20 pt, 264.7467x82.36 CSS px. This
            // keeps the Ribbon growth probe on the same session/render path as
            // the real presentation instead of inheriting the synthetic fixture.
            var userLikeMetadata = new FormulaMetadata
            {
                FormulaId = formulaId,
                Title = "PowerPoint Formula",
                Latex = PowerPointAcceptanceLatex,
                Lines = new List<FormulaLine>
                {
                    new() { Id = Guid.NewGuid().ToString(), Latex = PowerPointAcceptanceLatex },
                },
                CodeFormat = "raw",
                DisplayMode = "block",
                Numbered = false,
                RenderWidthPx = 264.74667358398438,
                RenderHeightPx = 82.36000061035156,
                Baseline = 53.8,
                FontSizePt = 20,
                RenderFontSizePt = 20,
                CreatedWithVersion = "1.0.18",
                UpdatedWithVersion = "1.0.18",
                CreatedAt = DateTimeOffset.UtcNow.ToString("O"),
                UpdatedAt = DateTimeOffset.UtcNow.ToString("O"),
            };
            var userLikeNatural = OfficeFormulaSizing.NaturalSize(
                (float)userLikeMetadata.RenderWidthPx.Value,
                (float)userLikeMetadata.RenderHeightPx.Value);
            var userLikeLeft = shape.Left;
            var userLikeTop = shape.Top;
            var userLikeShape = slide1.Shapes.AddPicture(
                svgPath,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                userLikeLeft,
                userLikeTop,
                userLikeNatural.Width,
                userLikeNatural.Height);
            userLikeShape.LockAspectRatio = MsoTriState.msoFalse;
            userLikeShape.Width = userLikeNatural.Width;
            userLikeShape.Height = userLikeNatural.Height;
            userLikeShape.LockAspectRatio = MsoTriState.msoTrue;
            userLikeShape.Name = $"VisualTeX_{formulaId}";
            var userLikeEncoded = FormulaMetadataCodec.Encode(userLikeMetadata);
            userLikeShape.AlternativeText = userLikeEncoded;
            var userLikeTags = userLikeShape.Tags;
            userLikeTags.Add("VisualTeXFormulaId", formulaId);
            userLikeTags.Add("VisualTeXMetadata", userLikeEncoded);
            Release(userLikeTags);
            shape.Delete();
            Release(shape);
            shape = userLikeShape;

            // Exercise the actual Ribbon -> Session -> converter -> PowerPoint path
            // for the new OMML mode and both existing conversion callbacks.
            addIn = new VisualTeX.PowerPointVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);
            application.ActiveWindow.Activate();
            application.ActiveWindow.View.GotoSlide(1);
            shape.Select(MsoTriState.msoTrue);
            var existing = SnapshotSessionIds();
            var converted = WaitForDirectConversion(
                client,
                existing,
                "powerpoint",
                "wordOmml",
                () => addIn.OnConvertSelectedOmml(new object()),
                TimeSpan.FromSeconds(45),
                out var ribbonPictureToOmmlElapsed,
                () => addIn.DiagnosticLastError);
            AssertEqual("completed", converted.Status, converted.Error ?? "PowerPoint Ribbon picture-to-OMML conversion failed.");
            Release(shape);
            shape = slide1.Shapes[1];
            AssertTrue(IsPowerPointNativeEquation(shape), "Ribbon picture -> OMML did not create native Office Math.");

            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            converted = WaitForDirectConversion(
                client,
                existing,
                "powerpoint",
                "nativeOle",
                () => addIn.OnConvertSelected(new object()),
                TimeSpan.FromSeconds(45),
                out _);
            AssertEqual("completed", converted.Status, converted.Error ?? "PowerPoint Ribbon OMML-to-OLE conversion failed.");
            Release(shape);
            shape = slide1.Shapes[1];
            AssertPowerPointOle(shape, "Ribbon OMML -> OLE did not create a VisualTeX OLE object.");

            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            converted = WaitForDirectConversion(
                client,
                existing,
                "powerpoint",
                "wordOmml",
                () => addIn.OnConvertSelectedOmml(new object()),
                TimeSpan.FromSeconds(45),
                out var ribbonOleToOmmlElapsed);
            AssertEqual("completed", converted.Status, converted.Error ?? "PowerPoint Ribbon OLE-to-OMML conversion failed.");
            Release(shape);
            shape = slide1.Shapes[1];
            AssertTrue(IsPowerPointNativeEquation(shape), "Ribbon OLE -> OMML did not create native Office Math.");

            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            converted = WaitForDirectConversion(
                client,
                existing,
                "powerpoint",
                "crossPlatformPicture",
                () => addIn.OnExportSelectedAsPicture(new object()),
                TimeSpan.FromSeconds(45),
                out _);
            AssertEqual("completed", converted.Status, converted.Error ?? "PowerPoint Ribbon OMML-to-picture conversion failed.");
            Release(shape);
            shape = slide1.Shapes[1];
            AssertTrue(IsPowerPointEditablePictureShape(shape), "Ribbon OMML -> picture did not create an SVG picture.");
            AssertEqual(formulaId, DecodePowerPointMetadata(shape)?.FormulaId, "Ribbon conversion chain changed the original FormulaId.");

            var ribbonRoundTripBaselineWidth = shape.Width;
            var ribbonRoundTripBaselineHeight = shape.Height;
            Console.WriteLine(
                $"PowerPoint Ribbon SVG/OLE growth probe baseline: {ribbonRoundTripBaselineWidth:F2}x{ribbonRoundTripBaselineHeight:F2} pt.");
            for (var roundTrip = 1; roundTrip <= 8; roundTrip++)
            {
                var beforeMetadata = DecodePowerPointMetadata(shape)
                    ?? throw new InvalidDataException($"Ribbon SVG/OLE growth probe lost picture metadata before round trip {roundTrip}.");
                Console.WriteLine(
                    $"  [Ribbon round {roundTrip}] SVG before={shape.Width:F2}x{shape.Height:F2} pt; " +
                    $"meta={beforeMetadata.RenderWidthPx:F2}x{beforeMetadata.RenderHeightPx:F2} px; " +
                    $"font={beforeMetadata.FontSizePt:F1}/{beforeMetadata.RenderFontSizePt:F1} pt.");

                shape.Select(MsoTriState.msoTrue);
                existing = SnapshotSessionIds();
                converted = WaitForDirectConversion(
                    client,
                    existing,
                    "powerpoint",
                    "nativeOle",
                    () => addIn.OnConvertSelected(new object()),
                    TimeSpan.FromSeconds(45),
                    out _);
                AssertEqual("completed", converted.Status,
                    converted.Error ?? $"Ribbon SVG-to-OLE growth probe failed at round trip {roundTrip}.");
                Release(shape);
                shape = slide1.Shapes[1];
                AssertPowerPointOle(shape, $"Ribbon SVG-to-OLE growth probe did not create OLE at round trip {roundTrip}.");
                var oleMetadata = DecodePowerPointMetadata(shape)
                    ?? throw new InvalidDataException($"Ribbon SVG/OLE growth probe lost OLE metadata at round trip {roundTrip}.");
                Console.WriteLine(
                    $"  [Ribbon round {roundTrip}] OLE after={shape.Width:F2}x{shape.Height:F2} pt; " +
                    $"export={converted.ExportResult?.Width:F2}x{converted.ExportResult?.Height:F2} px; " +
                    $"meta={oleMetadata.RenderWidthPx:F2}x{oleMetadata.RenderHeightPx:F2} px; " +
                    $"font={oleMetadata.FontSizePt:F1}/{oleMetadata.RenderFontSizePt:F1} pt.");

                if (roundTrip == 1)
                {
                    for (var sample = 1; sample <= 20; sample++)
                    {
                        WinForms.Application.DoEvents();
                        Thread.Sleep(250);
                        PowerPoint.OLEFormat? delayedFormat = null;
                        object? delayedObject = null;
                        try
                        {
                            delayedFormat = shape.OLEFormat;
                            delayedObject = delayedFormat.Object;
                            var serverWidth = double.NaN;
                            var serverHeight = double.NaN;
                            if (delayedObject is GeometryOleObjectNative nativeOle
                                && nativeOle.GetExtent(1, out var extent) >= 0)
                            {
                                serverWidth = extent.Cx * 72.0 / 2540.0;
                                serverHeight = extent.Cy * 72.0 / 2540.0;
                            }
                            Console.WriteLine(
                                $"    settle {sample * 250,4} ms: shape={shape.Width:F2}x{shape.Height:F2} pt; " +
                                $"server={serverWidth:F2}x{serverHeight:F2} pt.");
                        }
                        finally
                        {
                            Release(delayedObject);
                            Release(delayedFormat);
                        }
                    }
                }

                shape.Select(MsoTriState.msoTrue);
                existing = SnapshotSessionIds();
                converted = WaitForDirectConversion(
                    client,
                    existing,
                    "powerpoint",
                    "crossPlatformPicture",
                    () => addIn.OnExportSelectedAsPicture(new object()),
                    TimeSpan.FromSeconds(45),
                    out _);
                AssertEqual("completed", converted.Status,
                    converted.Error ?? $"Ribbon OLE-to-SVG growth probe failed at round trip {roundTrip}.");
                Release(shape);
                shape = slide1.Shapes[1];
                AssertTrue(IsPowerPointEditablePictureShape(shape),
                    $"Ribbon OLE-to-SVG growth probe did not create SVG at round trip {roundTrip}.");
                var pictureMetadata = DecodePowerPointMetadata(shape)
                    ?? throw new InvalidDataException($"Ribbon SVG/OLE growth probe lost picture metadata after round trip {roundTrip}.");
                Console.WriteLine(
                    $"  [Ribbon round {roundTrip}] SVG after={shape.Width:F2}x{shape.Height:F2} pt; " +
                    $"export={converted.ExportResult?.Width:F2}x{converted.ExportResult?.Height:F2} px; " +
                    $"meta={pictureMetadata.RenderWidthPx:F2}x{pictureMetadata.RenderHeightPx:F2} px; " +
                    $"font={pictureMetadata.FontSizePt:F1}/{pictureMetadata.RenderFontSizePt:F1} pt.");
            }
            AssertNear(ribbonRoundTripBaselineWidth, shape.Width, 1.0f,
                "Ribbon SVG/OLE round trips accumulated PowerPoint width growth.");
            AssertNear(ribbonRoundTripBaselineHeight, shape.Height, 1.0f,
                "Ribbon SVG/OLE round trips accumulated PowerPoint height growth.");

            Console.WriteLine(
                "PowerPoint Ribbon conversion path passed: SVG -> OMML -> OLE -> OMML -> SVG completed through VisualTeX Sessions. " +
                $"SVG->OMML={ribbonPictureToOmmlElapsed.TotalMilliseconds:F0} ms, " +
                $"OLE->OMML={ribbonOleToOmmlElapsed.TotalMilliseconds:F0} ms.");

            presentation.SaveAs(
                presentationPath,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                MsoTriState.msoTrue);
            Console.WriteLine("PowerPoint three-way conversion passed: picture <-> OLE, picture <-> OMML, and OLE <-> OMML all completed with native Office Math round-trip verification.");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            Release(copiedShape);
            Release(pastedRange);
            Release(shape);
            Release(slide2);
            Release(slide1);
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
            if (!string.IsNullOrWhiteSpace(emfPath))
            {
                try { File.Delete(emfPath); } catch { }
            }
            try { File.Delete(svgPath); } catch { }
            try { File.Delete(pngPath); } catch { }
            ForceComCleanup();
        }
    }

    private static OfficeSessionDocument CreatePowerPointAcceptanceSession(
        string mode,
        string objectMode,
        string formulaId,
        string? sourceObjectId,
        FormulaMetadata? originalMetadata)
    {
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString(),
            Mode = mode,
            Host = "powerpoint",
            FormulaId = formulaId,
            SourceDocumentId = null,
            SourceObjectId = sourceObjectId,
            Title = "PowerPoint copy/conversion acceptance",
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString(), Latex = PowerPointAcceptanceLatex },
            },
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = objectMode,
            Numbered = false,
            FontSizePt = 20,
            OriginalMetadata = originalMetadata,
            ExportResult = new OfficeExportDocument
            {
                Svg = CreateSvg(220, 72),
                SvgBase64 = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(CreateSvg(220, 72))),
                MathMl = PowerPointAcceptanceMathMl,
                PngBase64 = null,
                Width = 220,
                Height = 72,
                Baseline = 54,
            },
        };
    }

    private static void WriteAcceptancePng(
        string path,
        string text,
        int width,
        int height)
    {
        using var bitmap = new System.Drawing.Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using var graphics = System.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(System.Drawing.Color.Transparent);
        using var font = new System.Drawing.Font("Cambria Math", 30f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
        using var brush = new System.Drawing.SolidBrush(System.Drawing.Color.Black);
        graphics.DrawString(text, font, brush, 6, 18);
        bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
    }

    private static string CreatePowerPointAcceptanceEmf(
        string svgPath,
        float width,
        float height)
    {
        var type = typeof(PowerPointFormulaService).Assembly.GetType(
            "VisualTeX.WindowsOffice.VstoShared.OfficeOlePreview",
            throwOnError: true)
            ?? throw new InvalidOperationException("PowerPoint OfficeOlePreview type is unavailable.");
        var method = type.GetMethod(
            "CreateVectorEmfFromSvg",
            BindingFlags.Public | BindingFlags.Static)
            ?? throw new MissingMethodException(type.FullName, "CreateVectorEmfFromSvg");
        var result = (string)(method.Invoke(
            null,
            new object[] { svgPath, width, height, true, 0f })
            ?? throw new InvalidOperationException("PowerPoint EMF preview generation returned null."));
        var diagnostics = type.GetProperty(
            "LastRecordingDiagnostics",
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)
            ?.GetValue(null) as string;
        if (!string.IsNullOrWhiteSpace(diagnostics))
            Console.WriteLine($"  PowerPoint EMF recording: {diagnostics}");
        return result;
    }

    private static PowerPoint.ShapeRange PastePowerPointShapesWithRetry(
        PowerPoint.Slide slide)
    {
        PowerPoint.Shapes? shapes = null;
        Exception? lastError = null;
        try
        {
            shapes = slide.Shapes;
            var watch = System.Diagnostics.Stopwatch.StartNew();
            do
            {
                PowerPoint.ShapeRange? pasted = null;
                try
                {
                    pasted = shapes.Paste();
                    if (pasted is not null && pasted.Count > 0)
                    {
                        var result = pasted;
                        pasted = null;
                        return result;
                    }
                }
                catch (System.Runtime.InteropServices.COMException error)
                {
                    lastError = error;
                }
                finally { Release(pasted); }
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(50);
            }
            while (watch.Elapsed < TimeSpan.FromSeconds(3));
            throw new InvalidOperationException(
                "PowerPoint clipboard did not materialize the copied formula within 3 seconds.",
                lastError);
        }
        finally { Release(shapes); }
    }

    private static void AssertPowerPointOle(PowerPoint.Shape shape, string message)
    {
        PowerPoint.OLEFormat? format = null;
        try
        {
            if (shape.Type is not MsoShapeType.msoEmbeddedOLEObject
                and not MsoShapeType.msoLinkedOLEObject)
                throw new InvalidDataException(message);
            format = shape.OLEFormat;
            AssertEqual(
                FormulaOleContract.ProgId,
                format.ProgID,
                message);
        }
        finally { Release(format); }
    }

    private static void AssertPowerPointPptxContainsNativeOmml(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var xml = string.Join(
            "\n",
            archive.Entries
                .Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase)
                    && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .Select(entry =>
                {
                    using var stream = entry.Open();
                    using var reader = new StreamReader(stream);
                    return reader.ReadToEnd();
                }));
        AssertTrue(xml.IndexOf("<a14:m", StringComparison.Ordinal) >= 0, "Saved PowerPoint did not contain an a14:m native math wrapper.");
        AssertTrue(xml.IndexOf("<m:oMath", StringComparison.Ordinal) >= 0, "Saved PowerPoint did not contain OMML oMath markup.");
        AssertTrue(xml.IndexOf("<m:f>", StringComparison.Ordinal) >= 0, "Saved PowerPoint OMML did not contain the expected fraction structure.");
        AssertTrue(xml.IndexOf("<m:rad>", StringComparison.Ordinal) >= 0, "Saved PowerPoint OMML did not contain the expected radical structure.");
    }

    private static bool IsPowerPointNativeEquation(PowerPoint.Shape shape)
    {
        object? textFrame = null;
        object? range = null;
        object? mathZones = null;
        try
        {
            if (shape.HasTextFrame != MsoTriState.msoTrue) return false;
            textFrame = shape.TextFrame2;
            range = ((dynamic)textFrame).TextRange;
            try { mathZones = ((dynamic)range).MathZones(); }
            catch { mathZones = ((dynamic)range).MathZones(Type.Missing, Type.Missing); }
            return Convert.ToInt32(((dynamic)mathZones).Length) > 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(mathZones);
            Release(range);
            Release(textFrame);
        }
    }

    private static void AssertContains(string? value, string expected, string message)
    {
        if (value is null || value.IndexOf(expected, StringComparison.Ordinal) < 0)
            throw new InvalidDataException($"{message} Actual: {value ?? "<null>"}");
    }
}
