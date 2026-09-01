using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunActiveUserHundredFailureInspection()
    {
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = (Word.Application)Marshal.GetActiveObject("Word.Application");
            document = application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document to inspect.");
            Console.WriteLine(
                $"[ACTIVE USER DOC] name={document.Name}; saved={document.Saved}; "
                + $"inlineShapes={document.InlineShapes.Count}; omaths={document.OMaths.Count}; "
                + $"tables={document.Tables.Count}; frames={document.Frames.Count}; bookmarks={document.Bookmarks.Count}");

            for (var index = 1; index <= document.InlineShapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                Word.Range? range = null;
                Word.OLEFormat? format = null;
                try
                {
                    shape = document.InlineShapes[index];
                    range = shape.Range;
                    string progId;
                    try
                    {
                        format = shape.OLEFormat;
                        progId = format.ProgID ?? string.Empty;
                    }
                    catch { progId = string.Empty; }

                    if (string.Equals(progId, "Equation.DSMT4", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            var compoundFile = MathTypeOleStorage.CaptureCompoundFile(shape);
                            var equationNative = MathTypeOleStorage.ReadEquationNative(compoundFile);
                            var mathMl = MathTypeMtefCodec.ReadEquationNativeMathMl(equationNative);
                            var latex = MathMlToLatexConverter.Convert(mathMl).Trim();
                            var signature = MathTypeMtefCodec.SemanticSignature(mathMl);
                            var mtefLength = checked((int)BitConverter.ToUInt32(equationNative, 8));
                            var mtef = new byte[mtefLength];
                            Buffer.BlockCopy(equationNative, 28, mtef, 0, mtefLength);
                            var geometry = string.Empty;
                            if (index <= 20
                                && MathTypeNativePreviewRenderer.TryRender(
                                    mtef,
                                    Path.GetTempPath(),
                                    out var nativePreview))
                            {
                                using (nativePreview)
                                {
                                    var livePreview = ReadInlineShapeEnhancedMetafile(shape);
                                    using var liveBitmap = RenderEmf(
                                        livePreview,
                                        Math.Max(120, (int)Math.Ceiling(shape.Width * 6d)),
                                        Math.Max(72, (int)Math.Ceiling(shape.Height * 6d)));
                                    var nativeWmf = File.ReadAllBytes(nativePreview.WmfPath);
                                    using var nativeBitmap = RenderEmf(
                                        nativeWmf,
                                        Math.Max(120, (int)Math.Ceiling(nativePreview.WidthPt * 6d)),
                                        Math.Max(72, (int)Math.Ceiling(nativePreview.HeightPt * 6d)));
                                    geometry =
                                        $"|native={nativePreview.WidthPt:0.##}x{nativePreview.HeightPt:0.##},pos={nativePreview.WordPosition}"
                                        + $"|liveRight={MeasureRightWhiteMargin(liveBitmap)}|nativeRight={MeasureRightWhiteMargin(nativeBitmap)}";
                                }
                            }
                            Console.WriteLine(
                                $"MT|{index}|range={range.Start}:{range.End}|size={shape.Width:0.##}x{shape.Height:0.##}{geometry}|latex={latex}|signature={signature}");
                        }
                        catch (Exception error)
                        {
                            Console.WriteLine(
                                $"MT-READ-ERROR|{index}|range={range.Start}:{range.End}|{error.GetType().Name}|{error.Message}");
                        }
                    }
                    else if (string.Equals(progId, "VisualTeX.Formula.1", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            var metadata = WordFormulaMetadataReader.TryRead(shape);
                            Console.WriteLine(
                                $"VT|{index}|range={range.Start}:{range.End}|size={shape.Width:0.##}x{shape.Height:0.##}|formulaId={metadata?.FormulaId}|latex={metadata?.Latex}");
                        }
                        catch (Exception error)
                        {
                            Console.WriteLine(
                                $"VT-READ-ERROR|{index}|range={range.Start}:{range.End}|{error.GetType().Name}|{error.Message}");
                        }
                    }
                    else
                    {
                        Console.WriteLine(
                            $"OTHER|{index}|range={range.Start}:{range.End}|progId={progId}|type={shape.Type}");
                    }
                }
                finally
                {
                    Release(format);
                    Release(range);
                    Release(shape);
                }
            }

            for (var index = 1; index <= document.OMaths.Count; index++)
            {
                Word.OMath? math = null;
                Word.Range? range = null;
                try
                {
                    math = document.OMaths[index];
                    range = math.Range;
                    Console.WriteLine(
                        $"OMML|{index}|range={range.Start}:{range.End}|text={(range.Text ?? string.Empty).Replace("\r", "<CR>")}");
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }

            // Diagnose the exact live formula under the user's current Selection.
            // This is deliberately read-only: capture the production conversion
            // plan, then run only the standalone MathML -> Equation Native ->
            // MathML round-trip that InsertMathTypeOle validates before mutating
            // Word. It exposes the real mismatch from the user's unsaved document
            // instead of testing a reconstructed/equivalent formula elsewhere.
            var service = new WordFormulaService(application);
            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: false,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode);
            Console.WriteLine(
                $"LIVE-PLAN|selection={application.Selection.Start}:{application.Selection.End}|targets={plan.Targets.Count}");
            foreach (var target in plan.Targets.OrderBy(item => item.SourceStart))
            {
                var sourceMathMl = target.SourceMathMl
                    ?? throw new InvalidDataException(
                        $"Live OMML target {target.SourceFormulaId} has no SourceMathML.");
                var inline = string.Equals(
                    target.DisplayMode,
                    "inline",
                    StringComparison.OrdinalIgnoreCase);
                var generated = MathTypeMtefCodec.CreateEquationNative(sourceMathMl, inline);
                var compound = MathTypeOleStorage.CreateStandaloneCompoundFile(generated);
                var generatedMathMl = MathTypeOleStorage.ReadMathMl(compound);
                var expectedSignature = MathTypeMtefCodec.SemanticSignature(sourceMathMl);
                var actualSignature = MathTypeMtefCodec.SemanticSignature(generatedMathMl);
                Console.WriteLine(
                    $"LIVE-TARGET|start={target.SourceStart}|managed={target.SourceIsManagedOmml}|display={target.DisplayMode}|numbered={target.Numbered}|formulaId={target.SourceFormulaId}|objectId={target.SourceObjectId}|latex={target.Latex}");
                Console.WriteLine($"LIVE-SOURCE-MATHML|{sourceMathMl}");
                Console.WriteLine($"LIVE-GENERATED-MATHML|{generatedMathMl}");
                Console.WriteLine(
                    $"LIVE-SIGNATURE|match={string.Equals(expectedSignature, actualSignature, StringComparison.Ordinal)}|expected={expectedSignature}|actual={actualSignature}");
            }

            var wholePlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode);
            var liveMismatchCount = 0;
            var liveCodecErrorCount = 0;
            foreach (var target in wholePlan.Targets.OrderBy(item => item.SourceStart))
            {
                try
                {
                    var sourceMathMl = target.SourceMathMl
                        ?? throw new InvalidDataException("No SourceMathML.");
                    var generated = MathTypeMtefCodec.CreateEquationNative(
                        sourceMathMl,
                        inline: string.Equals(target.DisplayMode, "inline", StringComparison.OrdinalIgnoreCase));
                    var generatedMathMl = MathTypeOleStorage.ReadMathMl(
                        MathTypeOleStorage.CreateStandaloneCompoundFile(generated));
                    var expectedSignature = MathTypeMtefCodec.SemanticSignature(sourceMathMl);
                    var actualSignature = MathTypeMtefCodec.SemanticSignature(generatedMathMl);
                    if (string.Equals(expectedSignature, actualSignature, StringComparison.Ordinal))
                        continue;
                    liveMismatchCount++;
                    Console.WriteLine(
                        $"LIVE-AUDIT-MISMATCH|start={target.SourceStart}|display={target.DisplayMode}|managed={target.SourceIsManagedOmml}|latex={target.Latex}|expected={expectedSignature}|actual={actualSignature}|sourceMathMl={sourceMathMl}|generatedMathMl={generatedMathMl}");
                }
                catch (Exception error)
                {
                    liveCodecErrorCount++;
                    Console.WriteLine(
                        $"LIVE-AUDIT-ERROR|start={target.SourceStart}|display={target.DisplayMode}|managed={target.SourceIsManagedOmml}|latex={target.Latex}|{error.GetType().Name}|{error.Message}");
                }
            }
            Console.WriteLine(
                $"LIVE-AUDIT-SUMMARY|targets={wholePlan.Targets.Count}|mismatches={liveMismatchCount}|errors={liveCodecErrorCount}");

            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_ACTIVE_LIVE_CONVERT"),
                    "1",
                    StringComparison.Ordinal))
            {
                if (plan.Targets.Count != 1)
                    throw new InvalidDataException(
                        $"Live conversion requires exactly one selected OMML target, actual={plan.Targets.Count}.");
                var target = plan.Targets[0];
                var sourceMathMl = target.SourceMathMl
                    ?? throw new InvalidDataException("Live conversion target has no SourceMathML.");
                var generated = MathTypeMtefCodec.CreateEquationNative(
                    sourceMathMl,
                    inline: string.Equals(target.DisplayMode, "inline", StringComparison.OrdinalIgnoreCase));
                MathTypeNativePreviewRenderer.Result? nativePreview = null;
                try
                {
                    if (!MathTypeNativePreviewRenderer.TryRender(
                            generated.Mtef,
                            Path.GetTempPath(),
                            out nativePreview))
                        throw new InvalidDataException(
                            "Live conversion could not render the exact MathType native preview.");
                    var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal)
                    {
                        [target.Id] = new PreparedWordBulkFormula
                        {
                            Run = new WordBulkRun
                            {
                                Id = target.Id,
                                IsFormula = true,
                                Latex = target.Latex,
                                DisplayMode = target.DisplayMode,
                            },
                            Session = CreateOmmlMathTypeAcceptanceSession(
                                sourceMathMl,
                                target.DisplayMode,
                                target.Numbered,
                                FormulaOleContract.MathTypeOleMode),
                            MathMl = sourceMathMl,
                            MathTypeNativePreviewAttempted = true,
                            MathTypeNativePreview = nativePreview,
                        },
                    };
                    var beforeOmml = document.OMaths.Count;
                    var beforeMathType = CountMathTypeOleShapes(document);
                    var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);
                    Console.WriteLine(
                        $"LIVE-CONVERT-RESULT|converted={result.FormulaCount}|failed={result.FailedFormulaCount}|failures={string.Join(" || ", result.Failures)}|omml={beforeOmml}->{document.OMaths.Count}|mathType={beforeMathType}->{CountMathTypeOleShapes(document)}");
                    Word.InlineShape? convertedShape = null;
                    try
                    {
                        for (var index = 1; index <= document.InlineShapes.Count; index++)
                        {
                            Word.InlineShape? candidate = null;
                            Word.Range? candidateRange = null;
                            try
                            {
                                candidate = document.InlineShapes[index];
                                if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                                candidateRange = candidate.Range;
                                if (Math.Abs(candidateRange.Start - target.SourceStart) > 8) continue;
                                convertedShape = candidate;
                                candidate = null;
                                break;
                            }
                            finally
                            {
                                Release(candidateRange);
                                Release(candidate);
                            }
                        }
                        if (convertedShape is null)
                            throw new InvalidDataException(
                                "Live conversion reported success but no MathType OLE appeared at the selected source range.");
                        var liveMathMl = MathTypeOleStorage.ReadMathMl(convertedShape);
                        var expected = MathTypeMtefCodec.SemanticSignature(sourceMathMl);
                        var actual = MathTypeMtefCodec.SemanticSignature(liveMathMl);
                        Console.WriteLine(
                            $"LIVE-CONVERT-VERIFY|match={string.Equals(expected, actual, StringComparison.Ordinal)}|expected={expected}|actual={actual}|mathMl={liveMathMl}");
                    }
                    finally { Release(convertedShape); }
                }
                finally { nativePreview?.Dispose(); }
            }

            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_ACTIVE_SOURCE_ADDIN_DOUBLECLICK"),
                    "1",
                    StringComparison.Ordinal))
            {
                Microsoft.Office.Core.COMAddIns? comAddIns = null;
                Microsoft.Office.Core.COMAddIn? installedAddIn = null;
                ThisAddIn? sourceAddIn = null;
                Word.InlineShape? targetShape = null;
                Word.Range? targetRange = null;
                Word.Window? targetWindow = null;
                Array custom = Array.Empty<object>();
                var installedWasConnected = false;
                var sessionClient = new VisualTeXSessionClient();
                try
                {
                    comAddIns = application.COMAddIns;
                    installedAddIn = comAddIns.Item("VisualTeX.WordVsto");
                    installedWasConnected = installedAddIn.Connect;
                    Console.WriteLine(
                        $"LIVE-SOURCE-ADDIN|installedConnectedBefore={installedWasConnected}");

                    sourceAddIn = new ThisAddIn();
                    sourceAddIn.OnConnection(
                        application,
                        Extensibility.ext_ConnectMode.ext_cm_AfterStartup,
                        sourceAddIn,
                        ref custom);
                    // Keep the installed production add-in connected, but retain
                    // this temporary source instance's *new* low-level hook. Hooks
                    // are invoked newest-first, so the source hook can observe the
                    // live MathType OLE and dispatch VisualTeX without suppressing
                    // Word's second button-down. The older installed hook may still
                    // suppress downstream, which is exactly the user's current
                    // mixed-version situation we need to prove the new observer
                    // callback survives.
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(650);

                    var nearestDistance = int.MaxValue;
                    for (var index = 1; index <= document.InlineShapes.Count; index++)
                    {
                        Word.InlineShape? candidate = null;
                        Word.Range? candidateRange = null;
                        try
                        {
                            candidate = document.InlineShapes[index];
                            if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                            candidateRange = candidate.Range;
                            var distance = Math.Abs(candidateRange.Start - application.Selection.Start);
                            if (distance >= nearestDistance) continue;
                            Release(targetRange); targetRange = null;
                            Release(targetShape); targetShape = null;
                            targetShape = candidate;
                            candidate = null;
                            targetRange = candidateRange.Duplicate;
                            nearestDistance = distance;
                        }
                        finally
                        {
                            Release(candidateRange);
                            Release(candidate);
                        }
                    }
                    if (targetShape is null || targetRange is null)
                        throw new InvalidDataException(
                            "No live MathType OLE is available for current-source add-in double-click testing.");

                    targetRange.Select();
                    targetWindow = application.ActiveWindow;
                    targetWindow.Activate();
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(450);
                    targetWindow.GetPoint(out var left, out var top, out var width, out var height, targetRange);
                    if (width <= 0 || height <= 0)
                        throw new InvalidDataException(
                            "Word returned no visible rectangle for the live MathType OLE under the current-source add-in.");
                    var hwnd = new IntPtr(targetWindow.Hwnd);
                    const uint noMoveNoSizeShow = 0x0001 | 0x0002 | 0x0040;
                    SetWindowPos(hwnd, new IntPtr(-1), 0, 0, 0, 0, noMoveNoSizeShow);
                    SetForegroundWindow(hwnd);
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(300);

                    SetCursorPos(left + width / 2, top + height / 2);
                    mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                    mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                    Thread.Sleep(900);
                    targetRange.Select();
                    targetWindow.Activate();
                    SetForegroundWindow(hwnd);
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(250);
                    targetWindow.GetPoint(out left, out top, out width, out height, targetRange);

                    var sessionsBefore = SnapshotSessionIds();
                    SetCursorPos(left + width / 2, top + height / 2);
                    Thread.Sleep(120);
                    for (var click = 0; click < 2; click++)
                    {
                        mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                        mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                        Thread.Sleep(90);
                    }

                    var sessionId = WaitForNewSession(
                        sessionsBefore,
                        "word",
                        TimeSpan.FromSeconds(10));
                    var editSession = WaitForUnchangedEditorReady(
                        sessionClient,
                        sessionId,
                        TimeSpan.FromSeconds(10));
                    Console.WriteLine(
                        $"LIVE-SOURCE-ADDIN-DOUBLECLICK|opened=true|session={sessionId}|mode={editSession.Mode}|objectMode={editSession.ObjectMode}|latex={string.Join("\\n", editSession.Lines.Select(line => line.Latex))}|shapeRange={targetRange.Start}:{targetRange.End}|rect={left},{top},{width},{height}");
                    if (!string.Equals(editSession.Mode, "edit", StringComparison.Ordinal)
                        || !string.Equals(
                            editSession.ObjectMode,
                            FormulaOleContract.MathTypeOleMode,
                            StringComparison.Ordinal))
                        throw new InvalidDataException(
                            "Current-source MathType double-click opened the wrong VisualTeX Session mode.");
                    sessionClient.PatchAsync(
                            sessionId,
                            new Dictionary<string, object>
                            {
                                ["status"] = "cancelled",
                                ["explicitCancel"] = true,
                            },
                            CancellationToken.None)
                        .GetAwaiter().GetResult();
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(250);
                }
                finally
                {
                    if (sourceAddIn is not null)
                    {
                        try
                        {
                            sourceAddIn.OnDisconnection(
                                Extensibility.ext_DisconnectMode.ext_dm_UserClosed,
                                ref custom);
                        }
                        catch { }
                    }
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(250);
                    Console.WriteLine(
                        $"LIVE-SOURCE-ADDIN|installedConnectedAfter={(installedAddIn?.Connect == true)}");
                    Release(targetWindow);
                    Release(targetRange);
                    Release(targetShape);
                    Release(installedAddIn);
                    Release(comAddIns);
                }
            }

            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_ACTIVE_LIVE_DOUBLECLICK"),
                    "1",
                    StringComparison.Ordinal))
            {
                Word.InlineShape? targetShape = null;
                Word.Range? targetRange = null;
                Word.Window? targetWindow = null;
                var sessionClient = new VisualTeXSessionClient();
                try
                {
                    var nearestDistance = int.MaxValue;
                    for (var index = 1; index <= document.InlineShapes.Count; index++)
                    {
                        Word.InlineShape? candidate = null;
                        Word.Range? candidateRange = null;
                        try
                        {
                            candidate = document.InlineShapes[index];
                            if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                            candidateRange = candidate.Range;
                            var distance = Math.Abs(candidateRange.Start - application.Selection.Start);
                            if (distance >= nearestDistance) continue;
                            Release(targetRange); targetRange = null;
                            Release(targetShape); targetShape = null;
                            targetShape = candidate;
                            candidate = null;
                            targetRange = candidateRange.Duplicate;
                            nearestDistance = distance;
                        }
                        finally
                        {
                            Release(candidateRange);
                            Release(candidate);
                        }
                    }
                    if (targetShape is null || targetRange is null)
                        throw new InvalidDataException("No live MathType OLE is available for direct double-click testing.");

                    targetRange.Select();
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(250);
                    var selectedAsMathType = service.IsSelectedMathTypeOle();
                    var selectedFormula = service.ReadSelection();
                    targetWindow = application.ActiveWindow;
                    targetWindow.Activate();
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(400);
                    targetWindow.GetPoint(out var left, out var top, out var width, out var height, targetRange);
                    if (width <= 0 || height <= 0)
                        throw new InvalidDataException("Word returned no visible rectangle for the live MathType OLE.");
                    var centerX = left + width / 2;
                    var centerY = top + height / 2;
                    var hit = service.IsFormulaAtScreenPoint(selectedFormula, centerX, centerY);
                    var route = WordDoubleClickRouting.ShouldOpenVisualTeX(selectedFormula);
                    Console.WriteLine(
                        $"LIVE-DOUBLECLICK-PROBE|isSelectedMathType={selectedAsMathType}|formulaId={selectedFormula?.FormulaId}|objectMode={selectedFormula?.ObjectMode}|route={route}|hit={hit}|range={targetRange.Start}:{targetRange.End}|rect={left},{top},{width},{height}|latex={selectedFormula?.Metadata?.Latex}");
                    var hwnd = new IntPtr(targetWindow.Hwnd);
                    const uint noMoveNoSizeShow = 0x0001 | 0x0002 | 0x0040;
                    SetWindowPos(hwnd, new IntPtr(-1), 0, 0, 0, 0, noMoveNoSizeShow);
                    SetForegroundWindow(hwnd);
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(250);

                    // Prime Word with one ordinary click, wait beyond the system
                    // double-click interval, then send the real two-click gesture.
                    SetCursorPos(left + width / 2, top + height / 2);
                    mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                    mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                    Thread.Sleep(900);
                    targetRange.Select();
                    targetWindow.Activate();
                    SetForegroundWindow(hwnd);
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(250);
                    targetWindow.GetPoint(out left, out top, out width, out height, targetRange);

                    var sessionsBefore = SnapshotSessionIds();
                    SetCursorPos(left + width / 2, top + height / 2);
                    Thread.Sleep(120);
                    for (var click = 0; click < 2; click++)
                    {
                        mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                        mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                        Thread.Sleep(90);
                    }

                    try
                    {
                        var sessionId = WaitForNewSession(
                            sessionsBefore,
                            "word",
                            TimeSpan.FromSeconds(8));
                        var editSession = WaitForUnchangedEditorReady(
                            sessionClient,
                            sessionId,
                            TimeSpan.FromSeconds(8));
                        Console.WriteLine(
                            $"LIVE-DOUBLECLICK-RESULT|opened=true|session={sessionId}|mode={editSession.Mode}|objectMode={editSession.ObjectMode}|latex={string.Join("\\n", editSession.Lines.Select(line => line.Latex))}");
                        sessionClient.PatchAsync(
                                sessionId,
                                new Dictionary<string, object>
                                {
                                    ["status"] = "cancelled",
                                    ["explicitCancel"] = true,
                                },
                                CancellationToken.None)
                            .GetAwaiter().GetResult();
                    }
                    catch (TimeoutException error)
                    {
                        Console.WriteLine(
                            $"LIVE-DOUBLECLICK-RESULT|opened=false|error={error.Message}|shapeRange={targetRange.Start}:{targetRange.End}|rect={left},{top},{width},{height}");
                    }
                }
                finally
                {
                    Release(targetWindow);
                    Release(targetRange);
                    Release(targetShape);
                }
            }
        }
        finally
        {
            Release(document);
            Release(application);
        }
    }

    private static void RunActiveMathTypeSourceDoubleClickProbe()
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? targetShape = null;
        Word.Range? targetRange = null;
        Word.Window? window = null;
        ThisAddIn? sourceAddIn = null;
        Array custom = Array.Empty<object>();
        var client = new VisualTeXSessionClient();
        var liveWordHwnd = IntPtr.Zero;
        var restoreWordMinimized = false;
        try
        {
            application = (Word.Application)Marshal.GetActiveObject("Word.Application");
            document = application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document is available for live MathType double-click testing.");

            var selectionStart = application.Selection.Start;
            var nearestDistance = int.MaxValue;
            for (var index = 1; index <= document.InlineShapes.Count; index++)
            {
                Word.InlineShape? candidate = null;
                Word.Range? candidateRange = null;
                try
                {
                    candidate = document.InlineShapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    candidateRange = candidate.Range;
                    var distance = Math.Abs(candidateRange.Start - selectionStart);
                    if (distance >= nearestDistance) continue;
                    Release(targetRange); targetRange = null;
                    Release(targetShape); targetShape = null;
                    targetShape = candidate;
                    candidate = null;
                    targetRange = candidateRange.Duplicate;
                    nearestDistance = distance;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }
            if (targetShape is null || targetRange is null)
                throw new InvalidDataException("The active document has no MathType OLE available for live double-click testing.");

            sourceAddIn = new ThisAddIn();
            sourceAddIn.OnConnection(
                application,
                Extensibility.ext_ConnectMode.ext_cm_AfterStartup,
                sourceAddIn,
                ref custom);
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(650);

            targetRange.Select();
            window = application.ActiveWindow;
            window.Activate();
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(400);
            window.GetPoint(out var left, out var top, out var width, out var height, targetRange);
            if (width <= 0 || height <= 0)
                throw new InvalidDataException("Word returned no visible rectangle for the active MathType OLE.");

            var hwnd = new IntPtr(window.Hwnd);
            liveWordHwnd = hwnd;
            const uint noMoveNoSizeShow = 0x0001 | 0x0002 | 0x0040;
            if (left < -10000 || top < -10000)
            {
                // A minimized Word window reports formula coordinates around
                // (-32000,-32000); SetCursorPos then clamps to (0,0), which never
                // exercises the real equation. Restore it only for this live mouse
                // probe and put it back to minimized state in finally.
                restoreWordMinimized = true;
                ShowWindow(hwnd, 9); // SW_RESTORE
                document.Activate();
                window.Activate();
                SetForegroundWindow(hwnd);
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(650);
                SelectWordRangeWithRpcRetry(targetRange);
                window.GetPoint(out left, out top, out width, out height, targetRange);
                if (left < -10000 || top < -10000 || width <= 0 || height <= 0)
                    throw new InvalidDataException(
                        $"Word remained off-screen after restoring the minimized window: {left},{top},{width},{height}.");
            }
            SetWindowPos(hwnd, new IntPtr(-1), 0, 0, 0, 0, noMoveNoSizeShow);
            SetForegroundWindow(hwnd);
            if (GetWindowRect(hwnd, out var wordRect))
            {
                SetCursorPos(
                    wordRect.Left + Math.Max(40, (wordRect.Right - wordRect.Left) / 2),
                    wordRect.Top + 18);
                mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
            }
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(700);

            // Prime the actual OLE selection only after Word owns the foreground,
            // then wait beyond the system double-click interval before the pair.
            SelectWordRangeWithRpcRetry(targetRange);
            window.Activate();
            SetForegroundWindow(hwnd);
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(250);
            window.GetPoint(out left, out top, out width, out height, targetRange);
            if (width <= 0 || height <= 0)
                throw new InvalidDataException("Word lost the active MathType rectangle after foreground priming.");
            SetCursorPos(left + width / 2, top + height / 2);
            mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
            Thread.Sleep(900);
            SelectWordRangeWithRpcRetry(targetRange);
            window.Activate();
            SetForegroundWindow(hwnd);
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(250);
            window.GetPoint(out left, out top, out width, out height, targetRange);

            Console.WriteLine(
                $"[ACTIVE MATHTYPE SOURCE DOUBLECLICK PROBE] range={targetRange.Start}:{targetRange.End}; rect={left},{top},{width},{height}; cursor={System.Windows.Forms.Cursor.Position.X},{System.Windows.Forms.Cursor.Position.Y}");
            var sessionsBefore = SnapshotSessionIds();
            SetCursorPos(left + width / 2, top + height / 2);
            Thread.Sleep(120);
            for (var click = 0; click < 2; click++)
            {
                mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(90);
            }

            var sessionId = WaitForNewSession(
                sessionsBefore,
                "word",
                TimeSpan.FromSeconds(12));
            var editSession = WaitForUnchangedEditorReady(
                client,
                sessionId,
                TimeSpan.FromSeconds(12));
            if (!string.Equals(editSession.Mode, "edit", StringComparison.Ordinal)
                || !string.Equals(
                    editSession.ObjectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    $"Live MathType double-click opened the wrong Session: mode={editSession.Mode}, objectMode={editSession.ObjectMode}.");

            Console.WriteLine(
                $"[ACTIVE MATHTYPE SOURCE DOUBLECLICK PASS] document={document.Name}; range={targetRange.Start}:{targetRange.End}; "
                + $"rect={left},{top},{width},{height}; session={sessionId}; latex={string.Join("\\n", editSession.Lines.Select(line => line.Latex))}");

            client.PatchAsync(
                    sessionId,
                    new Dictionary<string, object>
                    {
                        ["status"] = "cancelled",
                        ["explicitCancel"] = true,
                    },
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            try
            {
                client.CloseEditorAsync(sessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
            }
            catch { }
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(250);
        }
        finally
        {
            if (sourceAddIn is not null)
            {
                try
                {
                    sourceAddIn.OnDisconnection(
                        Extensibility.ext_DisconnectMode.ext_dm_UserClosed,
                        ref custom);
                }
                catch { }
            }
            if (restoreWordMinimized && liveWordHwnd != IntPtr.Zero)
            {
                try { ShowWindow(liveWordHwnd, 6); } catch { } // SW_MINIMIZE
            }
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(200);
            Release(window);
            Release(targetRange);
            Release(targetShape);
            Release(document);
            Release(application);
        }
    }

    private static void SelectWordRangeWithRpcRetry(Word.Range range)
    {
        for (var attempt = 0; attempt < 20; attempt++)
        {
            try
            {
                range.Select();
                return;
            }
            catch (COMException error)
                when (error.HResult == unchecked((int)0x80010001) && attempt < 19)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
            }
        }
    }

    private static void RunUserHundredMathTypeSourceAudit()
    {
        var sourcePath = Environment.GetEnvironmentVariable("VISUALTEX_USER100_MATHTYPE_SOURCE");
        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            sourcePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Temp",
                "VisualTeX-VSTO-Flow-20260819-235014",
                "user-100-after-mathtype.docx");
        }
        sourcePath = Path.GetFullPath(sourcePath);
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException("The clean user-100 MathType source fixture is missing.", sourcePath);

        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(
                sourcePath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            var service = new WordFormulaService(application);
            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.NativeOleMode);
            Console.WriteLine(
                $"[USER100 MT SOURCE AUDIT] targets={plan.Targets.Count}; inlineShapes={document.InlineShapes.Count}; omaths={document.OMaths.Count}");
            var suspicious = 0;
            var ommlFailures = 0;
            var ommlRoundTripMismatches = 0;
            foreach (var pair in plan.Targets
                         .OrderBy(target => target.SourceStart)
                         .Select((target, index) => (Target: target, Index: index + 1)))
            {
                var latex = pair.Target.Latex ?? string.Empty;
                if (latex.IndexOf('?') >= 0
                    || latex.IndexOf('\uFFFD') >= 0)
                {
                    suspicious++;
                    Console.WriteLine(
                        $"SUSPICIOUS|{pair.Index}|start={pair.Target.SourceStart}|display={pair.Target.DisplayMode}|numbered={pair.Target.Numbered}|latex={latex}|mathml={pair.Target.SourceMathMl}");
                }
                try
                {
                    var sourceMathMl = pair.Target.SourceMathMl
                        ?? throw new InvalidDataException("MathType source target has no SourceMathMl.");
                    var omml = WordOmmlConverter.TransformMathMlToOmml(sourceMathMl);
                    var roundTripMathMl = WordOmmlConverter.TransformOmmlToMathMl(
                        omml,
                        display: string.Equals(pair.Target.DisplayMode, "block", StringComparison.Ordinal));
                    var expectedSignature = MathTypeMtefCodec.SemanticSignature(sourceMathMl);
                    var actualSignature = MathTypeMtefCodec.SemanticSignature(roundTripMathMl);
                    if (!string.Equals(expectedSignature, actualSignature, StringComparison.Ordinal))
                    {
                        ommlRoundTripMismatches++;
                        Console.WriteLine(
                            $"OMML-ROUNDTRIP-MISMATCH|{pair.Index}|latex={latex}|expected={expectedSignature}|actual={actualSignature}|mathml={roundTripMathMl}");
                    }
                }
                catch (Exception error)
                {
                    ommlFailures++;
                    Console.WriteLine(
                        $"OMML-PREFLIGHT-ERROR|{pair.Index}|start={pair.Target.SourceStart}|display={pair.Target.DisplayMode}|latex={latex}|error={error.GetType().Name}:{error.Message}|mathml={pair.Target.SourceMathMl}");
                }
            }
            Console.WriteLine(
                $"[USER100 MT SOURCE AUDIT DONE] suspicious={suspicious}; ommlFailures={ommlFailures}; ommlRoundTripMismatches={ommlRoundTripMismatches}");
        }
        finally
        {
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunUserHundredMathTypeReverseAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Environment.GetEnvironmentVariable("VISUALTEX_USER100_MATHTYPE_SOURCE");
        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            sourcePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Temp",
                "VisualTeX-VSTO-Flow-20260819-235014",
                "user-100-after-mathtype.docx");
        }
        sourcePath = Path.GetFullPath(sourcePath);
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException("The clean user-100 MathType source fixture is missing.", sourcePath);

        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_USER100_REVERSE_DIRECTION"),
                "omml",
                StringComparison.OrdinalIgnoreCase))
        {
            RunUserHundredMathTypeReverseDirection(
                sourcePath,
                Path.Combine(artifactRoot, "user-100-mathtype-to-visualtex.docx"),
                Path.Combine(artifactRoot, "user-100-mathtype-to-visualtex.trace.log"),
                targetVisualTeX: true);
        }
        RunUserHundredMathTypeReverseDirection(
            sourcePath,
            Path.Combine(artifactRoot, "user-100-mathtype-to-omml.docx"),
            Path.Combine(artifactRoot, "user-100-mathtype-to-omml.trace.log"),
            targetVisualTeX: false);
        Console.WriteLine("[USER100 REVERSE PASS] clean 100-MathType source survived MT→VisualTeX and MT→OMML with exact per-formula semantics.");
    }

    private static void RunUserHundredMathTypeReverseDirection(
        string sourcePath,
        string outputPath,
        string tracePath,
        bool targetVisualTeX)
    {
        File.Copy(sourcePath, outputPath, overwrite: true);
        try { File.Delete(tracePath); } catch { }
        var previousFormatAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousTracePath = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        var previousAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        Word.Application? application = null;
        Word.Document? document = null;
        Microsoft.Office.Core.COMAddIns? addIns = null;
        Microsoft.Office.Core.COMAddIn? installedAddIn = null;
        object? callbacksObject = null;
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            var mathTypeBaseline = SnapshotMathTypeProcessIds();
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            callbacksObject = GetInstalledStressCallbacks(
                application,
                out addIns,
                out installedAddIn);
            dynamic callbacks = callbacksObject;
            var sourceParagraphCount = document.Paragraphs.Count;
            var sourceBlankParagraphCount = CountStructurallyBlankParagraphs(document);

            var sourceService = new WordFormulaService(application);
            var targetMode = targetVisualTeX
                ? FormulaOleContract.NativeOleMode
                : FormulaOleContract.WordOmmlMode;
            var sourcePlan = sourceService.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                targetMode);
            var expected = sourcePlan.Targets
                .OrderBy(target => target.SourceStart)
                .Select(target => new
                {
                    Latex = NormalizeStressLatex(target.Latex ?? string.Empty),
                    Signature = MathTypeMtefCodec.SemanticSignature(
                        target.SourceMathMl
                        ?? throw new InvalidDataException("MathType source target has no SourceMathMl.")),
                    target.Numbered,
                    target.DisplayMode,
                    target.Metadata.FormulaId,
                })
                .ToArray();
            AssertEqual(100, expected.Length,
                "User-100 reverse acceptance did not capture exactly 100 MathType source formulas.");

            var watch = Stopwatch.StartNew();
            if (targetVisualTeX)
                callbacks.OnConvertMathTypeToVisualTeXDocument(new object());
            else
                callbacks.OnConvertMathTypeToOmmlDocument(new object());
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                targetVisualTeX
                    ? "source=MathType target=VisualTeX"
                    : "source=MathType target=OMML",
                mathTypeBaseline);
            watch.Stop();

            if (targetVisualTeX)
            {
                AssertEqual(100, CountInstalledVisualTeXOleShapes(document),
                    "MT→VisualTeX did not create exactly 100 VisualTeX OLE formulas.");
                var seen = 0;
                for (var index = 1; index <= document.InlineShapes.Count; index++)
                {
                    Word.InlineShape? shape = null;
                    try
                    {
                        shape = document.InlineShapes[index];
                        if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                        if (seen >= expected.Length)
                            throw new InvalidDataException("MT→VisualTeX created too many VisualTeX targets.");
                        var metadata = WordFormulaMetadataReader.TryRead(shape)
                            ?? throw new InvalidDataException($"MT→VisualTeX target #{seen + 1} has no metadata.");
                        var actualLatex = NormalizeStressLatex(metadata.Latex ?? string.Empty);
                        AssertEqual(
                            expected[seen].Latex,
                            actualLatex,
                            $"MT→VisualTeX formula #{seen + 1} changed LaTeX. expected='{expected[seen].Latex}', actual='{actualLatex}'.");
                        AssertEqual(
                            expected[seen].Numbered,
                            metadata.Numbered,
                            $"MT→VisualTeX formula #{seen + 1} changed numbered state.");
                        seen++;
                    }
                    finally { Release(shape); }
                }
                AssertEqual(100, seen, "MT→VisualTeX did not inspect exactly 100 VisualTeX formulas.");
            }
            else
            {
                AssertEqual(100, document.OMaths.Count,
                    "MT→OMML did not create exactly 100 OMML formulas.");
                var targetParagraphCount = document.Paragraphs.Count;
                var targetBlankParagraphCount = CountStructurallyBlankParagraphs(document);
                var numberedTargetCount = expected.Count(item => item.Numbered);
                Console.WriteLine(
                    $"[USER100 MT→OMML PARAGRAPHS] total={sourceParagraphCount}->{targetParagraphCount}; "
                    + $"blank={sourceBlankParagraphCount}->{targetBlankParagraphCount}; "
                    + $"numberedTargets={numberedTargetCount}.");
                AssertTrue(targetBlankParagraphCount <= sourceBlankParagraphCount,
                    $"MT→OMML introduced structurally blank paragraphs ({sourceBlankParagraphCount} -> {targetBlankParagraphCount}).");
                AssertTrue(sourceBlankParagraphCount - targetBlankParagraphCount <= numberedTargetCount,
                    $"MT→OMML removed non-numbering blank paragraphs ({sourceBlankParagraphCount} -> {targetBlankParagraphCount}; numbered={numberedTargetCount}).");
                AssertTrue(targetParagraphCount <= sourceParagraphCount + numberedTargetCount,
                    $"MT→OMML added non-numbering paragraphs ({sourceParagraphCount} -> {targetParagraphCount}; numbered={numberedTargetCount}).");
                for (var index = 0; index < expected.Length; index++)
                {
                    Word.OMath? math = null;
                    Word.Range? range = null;
                    try
                    {
                        math = document.OMaths[index + 1];
                        range = math.Range;
                        var mathMl = WordOmmlConverter.TransformOmmlToMathMl(
                            range.WordOpenXML,
                            display: string.Equals(
                                expected[index].DisplayMode,
                                "block",
                                StringComparison.Ordinal));
                        var actualSignature = MathTypeMtefCodec.SemanticSignature(mathMl);
                        if (!string.Equals(expected[index].Signature, actualSignature, StringComparison.Ordinal))
                            throw new InvalidDataException(
                                $"MT→OMML formula #{index + 1} changed semantics. expected={expected[index].Signature}; actual={actualSignature}; mathml={mathMl}");
                    }
                    finally
                    {
                        Release(range);
                        Release(math);
                    }
                }
            }

            AssertNoNewMathTypeProcess(
                mathTypeBaseline,
                targetVisualTeX ? "user-100 MT→VisualTeX" : "user-100 MT→OMML");
            document.Save();
            Console.WriteLine(
                $"[USER100 REVERSE] {(targetVisualTeX ? "MT→VisualTeX" : "MT→OMML")} passed in {watch.Elapsed.TotalSeconds:0.00}s.");
        }
        finally
        {
            Release(callbacksObject);
            Release(installedAddIn);
            Release(addIns);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", previousFormatAcceptance);
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTracePath);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", previousAcceptance);
        }
    }

    private static void RunWordUserHundredMathTypeConversionAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var formulas = new (string Latex, bool Display)[]
        {
            (@"a^2+b^2=c^2", false), // 1
            (@"\frac{a+b}{c+d}", true), // 2
            (@"\sqrt{x^2+y^2}", false), // 3
            (@"e^{i\pi}+1=0", true), // 4
            (@"x_{n+1}=x_n-\frac{f(x_n)}{f'(x_n)}", false), // 5
            (@"\alpha+\beta=\gamma", true), // 6
            (@"\sin^2 x+\cos^2 x=1", false), // 7
            (@"\log_a(xy)=\log_a x+\log_a y", true), // 8
            (@"\binom{n}{k}=\frac{n!}{k!(n-k)!}", false), // 9
            (@"|z|=\sqrt{x^2+y^2}", true), // 10
            (@"\lim_{x\to0}\frac{\sin x}{x}=1", false), // 11
            (@"\lim_{n\to\infty}\left(1+\frac1n\right)^n=e", true), // 12
            (@"\frac{d}{dx}x^n=nx^{n-1}", false), // 13
            (@"\frac{d}{dx}\sin x=\cos x", true), // 14
            (@"\frac{\partial f}{\partial x}+\frac{\partial f}{\partial y}=0", false), // 15
            (@"\int x^2\,dx=\frac{x^3}{3}+C", true), // 16
            (@"\int_0^1 x^n\,dx=\frac1{n+1}", false), // 17
            (@"\int_{-\infty}^{\infty}e^{-x^2}\,dx=\sqrt{\pi}", true), // 18
            (@"\oint_C\mathbf{F}\cdot d\mathbf{r}", false), // 19
            (@"\iint_D (x+y)\,dA", true), // 20
            (@"\iiint_V \rho\,dV", false), // 21
            (@"\sum_{k=1}^{n}k=\frac{n(n+1)}2", true), // 22
            (@"\sum_{n=0}^{\infty}x^n=\frac1{1-x}", false), // 23
            (@"\prod_{k=1}^{n}k=n!", true), // 24
            (@"\bigcup_{i=1}^{n}A_i", false), // 25
            (@"A\cap B=\varnothing", true), // 26
            (@"A\subseteq B", false), // 27
            (@"x\in\mathbb{R}", true), // 28
            (@"\forall x\in A,\ \exists y\in B", false), // 29
            (@"P(A\mid B)=\frac{P(A\cap B)}{P(B)}", true), // 30
            (@"E[X]=\sum_x xP(X=x)", false), // 31
            (@"\operatorname{Var}(X)=E[X^2]-E[X]^2", true), // 32
            (@"X\sim\mathcal{N}(\mu,\sigma^2)", false), // 33
            (@"f(x)=\frac1{\sqrt{2\pi}\sigma}e^{-\frac{(x-\mu)^2}{2\sigma^2}}", true), // 34
            (@"\rho_{XY}=\frac{\operatorname{Cov}(X,Y)}{\sigma_X\sigma_Y}", false), // 35
            (@"\vec{a}\cdot\vec{b}=|\vec a||\vec b|\cos\theta", true), // 36
            (@"\vec{a}\times\vec{b}", false), // 37
            (@"\nabla f=\left(\frac{\partial f}{\partial x},\frac{\partial f}{\partial y}\right)", true), // 38
            (@"\nabla\cdot\mathbf{F}=0", false), // 39
            (@"\nabla\times\mathbf{F}=\mathbf{0}", true), // 40
            (@"\mathbf{A}=\begin{pmatrix}a&b\\c&d\end{pmatrix}", false), // 41
            (@"\det\begin{pmatrix}a&b\\c&d\end{pmatrix}=ad-bc", true), // 42
            (@"\begin{bmatrix}1&0&0\\0&1&0\\0&0&1\end{bmatrix}", false), // 43
            (@"A^{-1}=\frac1{\det A}\operatorname{adj}(A)", true), // 44
            (@"A\mathbf{x}=\mathbf{b}", false), // 45
            (@"\lambda_{1,2}=\frac{\operatorname{tr}A\pm\sqrt{(\operatorname{tr}A)^2-4\det A}}2", true), // 46
            (@"\langle u,v\rangle=\sum_i u_i\overline{v_i}", false), // 47
            (@"\|x\|_2=\sqrt{\sum_i x_i^2}", true), // 48
            (@"\operatorname{rank}(A)\leq\min(m,n)", false), // 49
            (@"A^T A=I", true), // 50
            (@"f(x)=\begin{cases}x^2,&x\ge0\\-x,&x<0\end{cases}", false), // 51
            (@"|x|=\begin{cases}x,&x\ge0\\-x,&x<0\end{cases}", true), // 52
            (@"\max(a,b)=\begin{cases}a,&a\ge b\\b,&a<b\end{cases}", false), // 53
            (@"\delta_{ij}=\begin{cases}1,&i=j\\0,&i\ne j\end{cases}", true), // 54
            (@"\operatorname{sgn}(x)=\begin{cases}-1,&x<0\\0,&x=0\\1,&x>0\end{cases}", false), // 55
            (@"\begin{aligned}x+y&=3\\2x-y&=0\end{aligned}", true), // 56
            (@"\begin{aligned}a&=b+c\\&=d+e\end{aligned}", false), // 57
            (@"\begin{gathered}x_1+x_2=1\\x_1-x_2=0\end{gathered}", true), // 58
            (@"\left\{\begin{aligned}x+y&=1\\x-y&=2\end{aligned}\right.", false), // 59
            (@"\begin{array}{c|cc}&A&B\\\hline X&1&2\\Y&3&4\end{array}", true), // 60
            (@"\frac{1}{1+\frac{1}{1+x}}", false), // 61
            (@"\sqrt[3]{\frac{a^2+b^2}{c}}", true), // 62
            (@"\left(\frac{x+1}{x-1}\right)^2", false), // 63
            (@"\frac{\partial^2 u}{\partial x^2}+\frac{\partial^2 u}{\partial y^2}=0", true), // 64
            (@"\frac{d^2y}{dx^2}+\omega^2y=0", false), // 65
            (@"y''+py'+qy=0", true), // 66
            (@"\mathcal{L}\{f(t)\}=\int_0^\infty e^{-st}f(t)\,dt", false), // 67
            (@"\mathcal{F}\{f\}(\omega)=\int_{-\infty}^{\infty}f(t)e^{-i\omega t}\,dt", true), // 68
            (@"f(x)=\sum_{n=-\infty}^{\infty}c_ne^{inx}", false), // 69
            (@"c_n=\frac1{2\pi}\int_{-\pi}^{\pi}f(x)e^{-inx}\,dx", true), // 70
            (@"E=mc^2", false), // 71
            (@"F=ma", true), // 72
            (@"p=\hbar k", false), // 73
            (@"E=\hbar\omega", true), // 74
            (@"\Delta x\,\Delta p\ge\frac{\hbar}{2}", false), // 75
            (@"i\hbar\frac{\partial}{\partial t}\Psi=\hat H\Psi", true), // 76
            (@"\hat H=-\frac{\hbar^2}{2m}\nabla^2+V", false), // 77
            (@"\langle\psi|\phi\rangle", true), // 78
            (@"[\hat x,\hat p]=i\hbar", false), // 79
            (@"\psi(x)=Ae^{ikx}+Be^{-ikx}", true), // 80
            (@"R=\rho\frac{L}{A}", false), // 81
            (@"V=IR", true), // 82
            (@"P=VI=I^2R", false), // 83
            (@"C=\frac{Q}{V}", true), // 84
            (@"U=\frac12CV^2", false), // 85
            (@"\mathbf{E}=-\nabla V", true), // 86
            (@"\nabla\cdot\mathbf{E}=\frac{\rho}{\varepsilon_0}", false), // 87
            (@"\nabla\cdot\mathbf{B}=0", true), // 88
            (@"\nabla\times\mathbf{E}=-\frac{\partial\mathbf{B}}{\partial t}", false), // 89
            (@"\nabla\times\mathbf{B}=\mu_0\mathbf{J}+\mu_0\varepsilon_0\frac{\partial\mathbf{E}}{\partial t}", true), // 90
            (@"z=re^{i\theta}=r(\cos\theta+i\sin\theta)", false), // 91
            (@"\overline{z}=x-iy", true), // 92
            (@"\operatorname{Re}(z)=\frac{z+\overline z}{2}", false), // 93
            (@"\Gamma(n)=(n-1)!", true), // 94
            (@"B(x,y)=\frac{\Gamma(x)\Gamma(y)}{\Gamma(x+y)}", false), // 95
            (@"\zeta(s)=\sum_{n=1}^{\infty}\frac1{n^s}", true), // 96
            (@"\int_0^\infty x^{s-1}e^{-x}\,dx=\Gamma(s)", false), // 97
            (@"\frac{1}{2\pi i}\oint_C\frac{f(z)}{z-z_0}\,dz=f(z_0)", true), // 98
            (@"e^x=\sum_{n=0}^{\infty}\frac{x^n}{n!}", false), // 99
            (@"\cos x=\sum_{n=0}^{\infty}(-1)^n\frac{x^{2n}}{(2n)!}", true), // 100
        };
        var numberedFormulaIndices = new HashSet<int>
        {
            2, 4, 6, 8, 10, 14, 16, 18, 20, 22,
        };

        var sourceBuilder = new StringBuilder();
        for (var index = 0; index < formulas.Length; index++)
        {
            var formula = formulas[index];
            if (formula.Display)
                sourceBuilder.Append("$$").Append(formula.Latex).Append("$$\r\n");
            else
                sourceBuilder.Append("正文").Append(index + 1).Append("，$")
                    .Append(formula.Latex).Append("$。\r\n");
        }
        var sourcePath = Path.Combine(artifactRoot, "user-100-exact.tex");
        var bulkLogPath = Path.Combine(artifactRoot, "user-100.bulk.log");
        var tracePath = Path.Combine(artifactRoot, "user-100.conversion.trace.log");
        var beforePath = Path.Combine(artifactRoot, "user-100-before-conversion.docx");
        var outputPath = Path.Combine(artifactRoot, "user-100-after-mathtype.docx");
        File.WriteAllText(sourcePath, sourceBuilder.ToString(), new UTF8Encoding(false));
        DeleteBulkPerformanceArtifact(bulkLogPath);
        DeleteBulkPerformanceArtifact(tracePath);
        DeleteBulkPerformanceArtifact(beforePath);
        DeleteBulkPerformanceArtifact(outputPath);

        var previousBulkSource = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH");
        var previousBulkFormat = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT");
        var previousBulkMode = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE");
        var previousBulkLog = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG");
        var previousAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        var previousFormatAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousTracePath = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        Word.Application? application = null;
        Word.Document? document = null;
        Microsoft.Office.Core.COMAddIns? addIns = null;
        Microsoft.Office.Core.COMAddIn? installedAddIn = null;
        object? callbacksObject = null;
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", sourcePath);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", "latex");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", "ole");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", bulkLogPath);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);

            var mathTypeBaseline = SnapshotMathTypeProcessIds();
            if (mathTypeBaseline.Count != 0)
                throw new InvalidOperationException(
                    "User-100 acceptance requires MathType.exe to be closed before Word starts.");

            application = CreateWordApplication(visible: true);
            callbacksObject = GetInstalledStressCallbacks(
                application,
                out addIns,
                out installedAddIn);
            dynamic callbacks = callbacksObject;
            var reusableVisualTeXSource = Environment.GetEnvironmentVariable(
                "VISUALTEX_USER100_VT_SOURCE");
            var reusedVisualTeXSource =
                !string.IsNullOrWhiteSpace(reusableVisualTeXSource)
                && File.Exists(reusableVisualTeXSource);
            var importWatch = Stopwatch.StartNew();
            if (reusedVisualTeXSource)
            {
                File.Copy(reusableVisualTeXSource!, beforePath, true);
                document = application.Documents.Open(
                    beforePath,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: true);
                document.Activate();
                importWatch.Stop();
                Console.WriteLine(
                    $"[USER100 SOURCE] Reused and will revalidate VisualTeX-100 fixture: {reusableVisualTeXSource}");
            }
            else
            {
                document = application.Documents.Add();
                document.Activate();
                callbacks.OnBulkImport(new object());
                WaitForBulkImportCompletion(bulkLogPath, TimeSpan.FromMinutes(4));
                // The completion log is written at the end of the import
                // transaction; allow its async finally block to release the
                // operation gate before invoking the conversion callback.
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(500);
                importWatch.Stop();
            }
            AssertEqual(100, document.InlineShapes.Count,
                "Exact user-100 bulk import did not create 100 VisualTeX OLE formulas.");
            AssertEqual(100, CountInstalledVisualTeXOleShapes(document),
                "Exact user-100 bulk import created a non-VisualTeX formula object.");

            for (var index = 1; index <= formulas.Length; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = document.InlineShapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape)
                        ?? throw new InvalidDataException(
                            $"Exact user-100 VisualTeX formula #{index} has no readable metadata.");
                    AssertEqual(formulas[index - 1].Latex, metadata.Latex,
                        $"Exact user-100 VisualTeX formula #{index} changed before numbering/conversion.");
                }
                finally { Release(shape); }
            }

            if (!reusedVisualTeXSource)
            foreach (var formulaIndex in numberedFormulaIndices)
            {
                Word.InlineShape? shape = null;
                Word.Range? formulaRange = null;
                try
                {
                    shape = document.InlineShapes[formulaIndex];
                    var metadata = WordFormulaMetadataReader.TryRead(shape)
                        ?? throw new InvalidDataException(
                            $"Numbered source formula #{formulaIndex} has no metadata.");
                    metadata.Numbered = true;
                    metadata.DisplayMode = "block";
                    WordFormulaMetadataReader.CacheMetadata(shape, metadata);
                    formulaRange = shape.Range.Duplicate;
                    WordEquationNumbering.ReconcileFormula(
                        document,
                        formulaRange,
                        shape.Height,
                        metadata);
                }
                finally
                {
                    Release(formulaRange);
                    Release(shape);
                }
            }

            AssertEqual(100, document.InlineShapes.Count,
                "Numbering the exact user-100 source changed the formula count.");
            foreach (var formulaIndex in numberedFormulaIndices)
            {
                if (formulaIndex >= formulas.Length) continue;
                Word.InlineShape? following = null;
                Word.Range? followingRange = null;
                Word.Frames? frames = null;
                try
                {
                    following = document.InlineShapes[formulaIndex + 1];
                    followingRange = following.Range.Duplicate;
                    frames = followingRange.Frames;
                    AssertEqual(0, frames.Count,
                        $"Numbered formula #{formulaIndex} leaked its caption Frame into formula #{formulaIndex + 1}.");
                }
                finally
                {
                    Release(frames);
                    Release(followingRange);
                    Release(following);
                }
            }

            if (reusedVisualTeXSource)
                document.Save();
            else
                document.SaveAs2(beforePath, Word.WdSaveFormat.wdFormatXMLDocument);
            ResetInstalledFormatConversionTrace(tracePath);
            var conversionWatch = Stopwatch.StartNew();
            callbacks.OnConvertVisualTeXToMathTypeDocument(new object());
            var peakMathTypeProcessCount =
                WaitForInstalledOmmlMathTypeConversion(
                    tracePath,
                    "source=VisualTeX target=MathType",
                    mathTypeBaseline,
                    allowTransientMathTypeProcess: true);
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(500);
            conversionWatch.Stop();

            AssertEqual(0, CountInstalledVisualTeXOleShapes(document),
                "Exact user-100 conversion left VisualTeX OLE sources behind.");
            AssertEqual(100, CountMathTypeOleShapes(document),
                "Exact user-100 conversion did not create 100 MathType OLE formulas.");
            AssertEqual(100, document.InlineShapes.Count,
                "Exact user-100 conversion changed the total formula-object count.");
            // Persist the fully converted fixture before the expensive preview scan.
            // If a preview parser/regression assertion fails, the exact 100-formula
            // conversion result remains available for focused diagnostics instead of
            // forcing another five-minute conversion just to inspect the document.
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);

            var minRightMargin = int.MaxValue;
            var minRightMarginFormula = -1;
            for (var index = 1; index <= formulas.Length; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = document.InlineShapes[index];
                    var mathMl = MathTypeOleStorage.ReadMathMl(shape);
                    AssertTrue(!string.IsNullOrWhiteSpace(mathMl),
                        $"Converted MathType formula #{index} has no readable Equation Native MathML.");
                    var recoveredLatex = MathMlToLatexConverter.Convert(mathMl).Trim();
                    if (index == 36)
                        AssertTrue(recoveredLatex.IndexOf("\\theta", StringComparison.Ordinal) >= 0,
                            $"Converted user formula #36 lost its theta tail: '{recoveredLatex}'.");
                    if (index == 37)
                    {
                        AssertTrue(recoveredLatex.IndexOf("\\theta", StringComparison.Ordinal) < 0,
                            $"Converted user formula #37 inherited theta from formula #36: '{recoveredLatex}'.");
                        AssertTrue(recoveredLatex.IndexOf("\\times", StringComparison.Ordinal) >= 0,
                            $"Converted user formula #37 lost its multiplication operator: '{recoveredLatex}'.");
                    }

                    // Validate the presentation Word actually replays for the OLE.
                    // Re-decoding the package WMF through SetWinMetaFileBits can
                    // produce a blank/tight diagnostic bitmap for some native
                    // MathType records even though Word's live presentation is
                    // correct. The user-visible regression is therefore guarded by
                    // Word's own enhanced-metafile copy path, which must also stay
                    // completely offline (no MathType.exe activation).
                    var mathTypeProcessesBeforePreview = SnapshotMathTypeProcessIds();
                    var livePreview = ReadInlineShapeEnhancedMetafile(shape);
                    var startedDuringPreview = SnapshotMathTypeProcessIds()
                        .Except(mathTypeProcessesBeforePreview)
                        .ToArray();
                    AssertEqual(0, startedDuringPreview.Length,
                        $"Reading converted live preview #{index} started MathType.exe.");
                    AssertTrue(!string.Equals(
                            DescribeEmfInkBounds(livePreview),
                            "empty",
                            StringComparison.Ordinal),
                        $"Converted user formula #{index} has an empty live Word preview.");
                    using var bitmap = RenderEmf(
                        livePreview,
                        Math.Max(120, (int)Math.Ceiling(shape.Width * 6d)),
                        Math.Max(72, (int)Math.Ceiling(shape.Height * 6d)));
                    var rightMargin = MeasureRightWhiteMargin(bitmap);
                    if (rightMargin < minRightMargin)
                    {
                        minRightMargin = rightMargin;
                        minRightMarginFormula = index;
                    }
                    AssertTrue(rightMargin >= 2,
                        $"Converted user formula #{index} touches/clips the right preview edge: rightMargin={rightMargin}px; latex='{formulas[index - 1].Latex}'.");
                }
                finally { Release(shape); }
            }

            AssertNoNewMathTypeProcess(mathTypeBaseline, "exact user-100 VisualTeX→MathType conversion");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(100, CountMathTypeOleShapes(document),
                "Exact user-100 save/reopen lost MathType formulas.");
            for (var index = 1; index <= formulas.Length; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = document.InlineShapes[index];
                    AssertTrue(!string.IsNullOrWhiteSpace(MathTypeOleStorage.ReadMathMl(shape)),
                        $"Exact user-100 save/reopen made MathType formula #{index} unreadable.");
                    var mathTypeProcessesBeforePreview = SnapshotMathTypeProcessIds();
                    var reopenedPreview = ReadInlineShapeEnhancedMetafile(shape);
                    var startedDuringPreview = SnapshotMathTypeProcessIds()
                        .Except(mathTypeProcessesBeforePreview)
                        .ToArray();
                    AssertEqual(0, startedDuringPreview.Length,
                        $"Reading reopened live preview #{index} started MathType.exe.");
                    AssertTrue(!string.Equals(
                            DescribeEmfInkBounds(reopenedPreview),
                            "empty",
                            StringComparison.Ordinal),
                        $"Exact user-100 save/reopen made preview #{index} empty.");
                    using var reopenedBitmap = RenderEmf(
                        reopenedPreview,
                        Math.Max(120, (int)Math.Ceiling(shape.Width * 6d)),
                        Math.Max(72, (int)Math.Ceiling(shape.Height * 6d)));
                    var reopenedRightMargin = MeasureRightWhiteMargin(reopenedBitmap);
                    if (reopenedRightMargin < minRightMargin)
                    {
                        minRightMargin = reopenedRightMargin;
                        minRightMarginFormula = index;
                    }
                    AssertTrue(reopenedRightMargin >= 2,
                        $"Reopened user formula #{index} touches/clips the live right preview edge: rightMargin={reopenedRightMargin}px; latex='{formulas[index - 1].Latex}'.");
                }
                finally { Release(shape); }
            }

            Console.WriteLine(
                $"[USER100 PASS] import={importWatch.Elapsed.TotalSeconds:0.00}s conversion={conversionWatch.Elapsed.TotalSeconds:0.00}s; "
                + $"100/100 MathType OLE readable; numbered source Frames stayed local; #23 converted; #37 did not inherit theta; "
                + $"minimum preview right margin={minRightMargin}px at formula #{minRightMarginFormula}; save/reopen passed; "
                + $"peakTransientMathTypeProcessCount={peakMathTypeProcessCount}; finalMathTypeProcessCount=0.");
            Console.WriteLine($"Artifacts: before={beforePath}; after={outputPath}");
        }
        finally
        {
            Release(callbacksObject);
            Release(installedAddIn);
            Release(addIns);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", previousBulkSource);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", previousBulkFormat);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", previousBulkMode);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", previousBulkLog);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", previousAcceptance);
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", previousFormatAcceptance);
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTracePath);
        }
    }

    private static void RunUserHundredMathTypePreviewScan(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Environment.GetEnvironmentVariable("VISUALTEX_USER100_MATHTYPE_SOURCE");
        if (string.IsNullOrWhiteSpace(sourcePath))
            sourcePath = Path.Combine(artifactRoot, "user-100-after-mathtype.docx");
        sourcePath = Path.GetFullPath(sourcePath);
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException("User-100 MathType preview scan source is missing.", sourcePath);

        Word.Application? application = null;
        Word.Document? document = null;
        var minRightMargin = int.MaxValue;
        var minRightMarginFormula = -1;
        var edgeTouching = new List<int>();
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(
                sourcePath,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false);
            AssertEqual(100, CountMathTypeOleShapes(document),
                "User-100 preview scan source does not contain exactly 100 MathType OLE formulas.");

            for (var index = 1; index <= document.InlineShapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = document.InlineShapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                    var mathMl = MathTypeOleStorage.ReadMathMl(shape);
                    var latex = MathMlToLatexConverter.Convert(mathMl).Trim();
                    var mathTypeProcessesBeforePreview = SnapshotMathTypeProcessIds();
                    var livePreview = ReadInlineShapeEnhancedMetafile(shape);
                    var startedDuringPreview = SnapshotMathTypeProcessIds()
                        .Except(mathTypeProcessesBeforePreview)
                        .ToArray();
                    AssertEqual(0, startedDuringPreview.Length,
                        $"Reading live preview #{index} started MathType.exe.");
                    AssertTrue(!string.Equals(
                            DescribeEmfInkBounds(livePreview),
                            "empty",
                            StringComparison.Ordinal),
                        $"Live Word preview #{index} is empty.");
                    using var bitmap = RenderEmf(
                        livePreview,
                        Math.Max(120, (int)Math.Ceiling(shape.Width * 6d)),
                        Math.Max(72, (int)Math.Ceiling(shape.Height * 6d)));
                    var leftMargin = MeasureLeftWhiteMargin(bitmap);
                    var rightMargin = MeasureRightWhiteMargin(bitmap);
                    if (rightMargin < minRightMargin)
                    {
                        minRightMargin = rightMargin;
                        minRightMarginFormula = index;
                    }
                    if (rightMargin < 2) edgeTouching.Add(index);
                    Console.WriteLine(
                        $"PREVIEW|{index}|size={shape.Width:0.###}x{shape.Height:0.###}|left={leftMargin}|right={rightMargin}|latex={latex}");
                }
                finally { Release(shape); }
            }

            Console.WriteLine(
                $"[USER100 PREVIEW SCAN] count={document.InlineShapes.Count}; minRight={minRightMargin}px at #{minRightMarginFormula}; "
                + $"rightMarginBelow2={edgeTouching.Count}; indices={string.Join(",", edgeTouching)}");
        }
        finally
        {
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }
}
