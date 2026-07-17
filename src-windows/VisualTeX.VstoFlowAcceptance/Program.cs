using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Extensibility;
using Microsoft.Office.Core;
using VisualTeX.WindowsOffice.Contracts;
using WinForms = System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static class Program
{
    private const float ExportWidth = 160f;
    private const float ExportHeight = 32f;
    private const float ExportBaseline = 24f;
    private const uint MouseLeftDown = 0x0002;
    private const uint MouseLeftUp = 0x0004;
    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr windowHandle);

    [DllImport("user32.dll")]
    private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiContext);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GetConsoleWindow();

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr windowHandle, int command);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr windowHandle, out NativeRect rectangle);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(
        IntPtr windowHandle,
        IntPtr insertAfter,
        int x,
        int y,
        int width,
        int height,
        uint flags);

    [DllImport("user32.dll")]
    private static extern void mouse_event(
        uint flags,
        uint dx,
        uint dy,
        uint data,
        UIntPtr extraInfo);

    private static readonly string SessionRoot = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "com.visualtex.studio",
        "office",
        "sessions");

    [STAThread]
    private static int Main(string[] args)
    {
        try { SetProcessDpiAwarenessContext(new IntPtr(-4)); } catch { }
        using var instanceMutex = new Mutex(
            initiallyOwned: true,
            name: @"Local\VisualTeX.VstoFlowAcceptance",
            createdNew: out var createdNew);
        if (!createdNew)
        {
            Console.Error.WriteLine("Another VisualTeX VSTO acceptance instance is already running.");
            return 4;
        }

        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", "1");
        var mode = args
            .FirstOrDefault(argument => argument.StartsWith("--mode=", StringComparison.OrdinalIgnoreCase))
            ?.Substring("--mode=".Length)
            ?? "all";
        var artifactArgument = args.FirstOrDefault(argument => !argument.StartsWith("--", StringComparison.Ordinal));
        var artifactRoot = artifactArgument is not null
            ? Path.GetFullPath(artifactArgument)
            : Path.Combine(Path.GetTempPath(), $"VisualTeX-VSTO-Flow-{DateTime.Now:yyyyMMdd-HHmmss}");
        Directory.CreateDirectory(artifactRoot);

        using var log = new StreamWriter(
            Path.Combine(artifactRoot, "acceptance.log"),
            append: false,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false))
        {
            AutoFlush = true,
        };
        var originalOut = Console.Out;
        var originalError = Console.Error;
        Console.SetOut(new TeeTextWriter(originalOut, log));
        Console.SetError(new TeeTextWriter(originalError, log));
        Console.WriteLine($"Acceptance mode: {mode}");

        try
        {
            using var client = new VisualTeXSessionClient();
            client.EnsureHealthyAsync(CancellationToken.None).GetAwaiter().GetResult();
            if (string.Equals(mode, "word-native-crossref-probe", StringComparison.OrdinalIgnoreCase))
            {
                ProbeNativeEquationCrossReference();
            }
            else if (string.Equals(mode, "word-crossref", StringComparison.OrdinalIgnoreCase))
            {
                RunWordNativeCrossReference(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-create", StringComparison.OrdinalIgnoreCase))
            {
                RunWord(client, artifactRoot, initialOnly: true);
            }
            else if (string.Equals(mode, "word-omml-double-click-fixtures", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOmmlDoubleClickFixtures(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-ole-real-double-click", StringComparison.OrdinalIgnoreCase))
            {
                RunWordOleRealDoubleClick(client, artifactRoot);
            }
            else if (string.Equals(mode, "word-unchanged", StringComparison.OrdinalIgnoreCase))
            {
                RunWord(client, artifactRoot, stopAfterUnchanged: true);
            }
            else
            {
                RunWord(client, artifactRoot);
                RunPowerPoint(client, artifactRoot);
            }
            Console.WriteLine("VisualTeX real VSTO formula flow acceptance passed.");
            Console.WriteLine($"Artifacts: {artifactRoot}");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            Console.Error.WriteLine($"Acceptance artifacts retained: {artifactRoot}");
            return 1;
        }
    }

    private static void ProbeNativeEquationCrossReference()
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.Fields? fields = null;
        try
        {
            application = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
            document = application.Documents.Add();
            range = document.Range(0, 0);
            range.InsertCaption(
                Label: Word.WdCaptionLabelID.wdCaptionEquation,
                Title: " included",
                Position: Word.WdCaptionPosition.wdCaptionPositionBelow,
                ExcludeLabel: false);
            Release(range);
            range = null;

            var documentEnd = document.Content.End - 1;
            object secondStart = documentEnd;
            object secondEnd = documentEnd;
            range = document.Range(ref secondStart, ref secondEnd);
            range.InsertParagraphAfter();
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.InsertCaption(
                Label: Word.WdCaptionLabelID.wdCaptionEquation,
                Title: " excluded",
                Position: Word.WdCaptionPosition.wdCaptionPositionBelow,
                ExcludeLabel: true);
            Release(range);
            range = null;

            var label = application.CaptionLabels[Word.WdCaptionLabelID.wdCaptionEquation];
            var labelName = label.Name;
            Release(label);
            documentEnd = document.Content.End - 1;
            object thirdStart = documentEnd;
            object thirdEnd = documentEnd;
            range = document.Range(ref thirdStart, ref thirdEnd);
            range.InsertParagraphAfter();
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.Text = "(";
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            var manualField = document.Fields.Add(
                range,
                Word.WdFieldType.wdFieldEmpty,
                $"SEQ {labelName} \\* ARABIC",
                true);
            Release(manualField);
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.InsertAfter(") manual");
            Release(range);
            range = null;

            documentEnd = document.Content.End - 1;
            object fourthStart = documentEnd;
            object fourthEnd = documentEnd;
            range = document.Range(ref fourthStart, ref fourthEnd);
            range.InsertParagraphAfter();
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            range.InsertCaption(
                Label: Word.WdCaptionLabelID.wdCaptionEquation,
                Title: string.Empty,
                Position: Word.WdCaptionPosition.wdCaptionPositionBelow,
                ExcludeLabel: true);
            Word.Font? helperFont = null;
            Word.ParagraphFormat? helperParagraph = null;
            try
            {
                helperFont = range.Font;
                helperFont.Hidden = 0;
                helperFont.Size = 1f;
                helperFont.Color = Word.WdColor.wdColorWhite;
                helperParagraph = range.ParagraphFormat;
                helperParagraph.SpaceBefore = 0f;
                helperParagraph.SpaceAfter = 0f;
                helperParagraph.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                helperParagraph.LineSpacing = 1f;
            }
            finally
            {
                Release(helperParagraph);
                Release(helperFont);
            }
            Release(range);
            range = null;

            fields = document.Fields;
            Console.WriteLine($"Native equation caption field count: {fields.Count}");
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? field = null;
                Word.Range? code = null;
                Word.Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    result = field.Result;
                    Console.WriteLine(
                        $"Caption field {index}: code=[{code.Text?.Trim()}] result=[{result.Text?.Trim()}]");
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }

            var items = document.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionEquation);
            if (items is not Array array)
                throw new InvalidDataException("Word did not return an equation cross-reference array.");
            Console.WriteLine($"Native equation cross-reference item count: {array.Length}");
            for (var index = array.GetLowerBound(0); index <= array.GetUpperBound(0); index++)
                Console.WriteLine($"Native equation item {index}: [{array.GetValue(index)}]");
            if (array.Length < 4)
                throw new InvalidDataException("Word hidden native equation caption was not listed for cross-reference.");

            var referenceStart = document.Content.End - 1;
            object referenceStartObject = referenceStart;
            object referenceEndObject = referenceStart;
            range = document.Range(ref referenceStartObject, ref referenceEndObject);
            range.InsertParagraphAfter();
            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            Word.Font? visibleReferenceFont = null;
            try
            {
                visibleReferenceFont = range.Font;
                visibleReferenceFont.Hidden = 0;
            }
            finally { Release(visibleReferenceFont); }
            range.InsertCrossReference(
                ReferenceType: Word.WdCaptionLabelID.wdCaptionEquation,
                ReferenceKind: Word.WdReferenceKind.wdEntireCaption,
                ReferenceItem: 4,
                InsertAsHyperlink: true,
                IncludePosition: false);
            fields = document.Fields;
            Console.WriteLine($"Field count after native reference insertion: {fields.Count}");
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? nativeField = null;
                Word.Range? nativeCode = null;
                Word.Range? nativeResult = null;
                Word.Font? nativeResultFont = null;
                try
                {
                    nativeField = fields[index];
                    nativeCode = nativeField.Code;
                    nativeResult = nativeField.Result;
                    nativeResultFont = nativeResult.Font;
                    Console.WriteLine(
                        $"Post-reference field {index}: type={nativeField.Type} " +
                        $"code=[{nativeCode.Text?.Trim()}] result=[{nativeResult.Text?.Trim()}] " +
                        $"hidden={nativeResultFont.Hidden}");
                }
                finally
                {
                    Release(nativeResultFont);
                    Release(nativeResult);
                    Release(nativeCode);
                    Release(nativeField);
                }
            }
        }
        finally
        {
            Release(fields);
            Release(range);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            if (application is not null)
            {
                try { application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordNativeCrossReference(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? firstShape = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            Console.WriteLine("[Native cross-reference 1/7] Starting Word...");
            application = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
            document = application.Documents.Add();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            Console.WriteLine("[Native cross-reference 2/7] Creating equation (1)...");
            var existing = SnapshotSessionIds();
            addIn.OnInsertDisplay(new object());
            var firstSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var firstSession = client.GetSessionAsync(firstSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(client, firstSession, "block", "nativeOle", "a=b", numbered: true);
            var final = WaitForTerminal(client, firstSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "First numbered formula did not complete.");
            client.CloseEditorAsync(firstSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            var firstFormulaId = final.FormulaId
                ?? throw new InvalidDataException("First numbered formula has no formulaId.");
            WaitForWordInlineShapeCount(document, 1, TimeSpan.FromSeconds(15));

            Console.WriteLine("[Native cross-reference 3/7] Creating equation (2)...");
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            existing = SnapshotSessionIds();
            addIn.OnInsertDisplay(new object());
            var secondSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var secondSession = client.GetSessionAsync(secondSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(client, secondSession, "block", "nativeOle", "E=mc^2", numbered: true);
            final = WaitForTerminal(client, secondSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "Second numbered formula did not complete.");
            client.CloseEditorAsync(secondSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            var secondFormulaId = final.FormulaId
                ?? throw new InvalidDataException("Second numbered formula has no formulaId.");
            WaitForWordInlineShapeCount(document, 2, TimeSpan.FromSeconds(15));

            Console.WriteLine("[Native cross-reference 4/7] Checking Word's built-in Equation list...");
            var nativeItems = document.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionEquation) as Array;
            if (nativeItems is null || nativeItems.Length != 2)
                throw new InvalidDataException(
                    $"Word native Equation list should contain two VisualTeX formulas, actual count: {nativeItems?.Length ?? 0}.");
            var firstItem = Convert.ToString(nativeItems.GetValue(nativeItems.GetLowerBound(0)))?.Trim();
            var secondItem = Convert.ToString(nativeItems.GetValue(nativeItems.GetLowerBound(0) + 1))?.Trim();
            AssertEqual("1", firstItem, "First native Equation item is not the pure number 1.");
            AssertEqual("2", secondItem, "Second native Equation item is not the pure number 2.");

            Console.WriteLine("[Native cross-reference 5/7] Inserting a native REF to equation (2)...");
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            application.Selection.TypeText("See ");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REFERENCE_TARGET_INDEX", "1");
            addIn.OnInsertEquationReference(new object());
            var nativeReferenceCode = WaitForWordNativeReferenceResult(
                document,
                expectedResult: "2",
                expectedCode: null,
                TimeSpan.FromSeconds(15));
            if (!DocumentTextContains(document, "(2)"))
                throw new InvalidDataException("Native Word reference did not render as (2).");

            Console.WriteLine("[Native cross-reference 6/7] Deleting equation (1) and updating fields...");
            firstShape = document.InlineShapes[1];
            firstShape.Range.Select();
            addIn.OnDeleteSelected(new object());
            Release(firstShape);
            firstShape = null;
            WaitForWordInlineShapeCount(document, 1, TimeSpan.FromSeconds(15));
            addIn.OnUpdateEquationNumbers(new object());
            WaitForWordNativeReferenceResult(
                document,
                expectedResult: "1",
                expectedCode: nativeReferenceCode,
                TimeSpan.FromSeconds(15));
            if (!DocumentTextContains(document, "(1)"))
                throw new InvalidDataException("Native Word reference did not update to (1).");
            if (WordBookmarkExists(document, $"VTEq_{Guid.Parse(firstFormulaId):N}"))
                throw new InvalidDataException("Deleted formula retained its visible number bookmark.");
            if (!WordBookmarkExists(document, $"VTEq_{Guid.Parse(secondFormulaId):N}"))
                throw new InvalidDataException("Remaining formula lost its visible number bookmark.");
            nativeItems = document.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionEquation) as Array;
            if (nativeItems is null || nativeItems.Length != 1)
                throw new InvalidDataException(
                    $"Word native Equation list should contain one item after deletion, actual count: {nativeItems?.Length ?? 0}.");

            var path = Path.Combine(artifactRoot, "VisualTeX-Word-Native-CrossReference.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine($"[Native cross-reference 7/7] Saved {path}; native Equation list and REF update passed.");
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REFERENCE_TARGET_INDEX", null);
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            Release(firstShape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            if (application is not null)
            {
                try { application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordOleRealDoubleClick(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.Range? range = null;
        Word.Window? window = null;
        COMAddIns? installedAddIns = null;
        COMAddIn? installedAddIn = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        var consoleWindow = GetConsoleWindow();
        var hookTracePath = Path.Combine(artifactRoot, "word-ole-hook-trace.log");
        Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", hookTracePath);
        try
        {
            Console.WriteLine("[Word real OLE 1/6] Starting visible Word...");
            application = new Word.Application
            {
                Visible = true,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };

            // This acceptance must exercise the current source assembly, not the
            // formally installed previous build. Disconnect the registered add-in
            // only for this Word process and reconnect it before shutdown.
            installedAddIns = application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.WordVsto";
                installedAddIn = installedAddIns.Item(ref addInIndex);
                if (installedAddIn.Connect)
                    installedAddIn.Connect = false;
            }
            catch
            {
                Release(installedAddIn);
                installedAddIn = null;
            }

            document = application.Documents.Add();
            application.Selection.TypeText("Real OLE double-click: ");
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            Console.WriteLine("[Word real OLE 2/6] Creating a native OLE formula...");
            var existing = SnapshotSessionIds();
            addIn.OnInsertInline(new object());
            var createSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var createSession = client.GetSessionAsync(createSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(
                client,
                createSession,
                "inline",
                FormulaOleContract.NativeOleMode,
                "\\int_0^1 x^2\\,dx");
            var created = WaitForTerminal(client, createSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", created.Status,
                created.Error ?? "Real OLE fixture creation did not complete.");
            client.CloseEditorAsync(createSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            Console.WriteLine("[Word real OLE 3/6] Resolving the formula screen rectangle...");
            AssertEqual(1, document.InlineShapes.Count,
                "Real OLE fixture should contain exactly one inline shape.");
            shape = document.InlineShapes[1];
            AssertEqual(FormulaOleContract.ProgId, shape.OLEFormat.ProgID,
                "Real OLE fixture has the wrong ProgID.");
            range = shape.Range;
            range.Select();
            application.ActiveWindow.Activate();
            WinForms.Application.DoEvents();
            Thread.Sleep(500);
            var addInType = typeof(VisualTeX.WordVsto.ThisAddIn);
            var activeField = addInType.GetField(
                "_nativeOleTargetActive",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            var leftField = addInType.GetField(
                "_nativeOleTargetLeft",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            var topField = addInType.GetField(
                "_nativeOleTargetTop",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            var rightField = addInType.GetField(
                "_nativeOleTargetRight",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            var bottomField = addInType.GetField(
                "_nativeOleTargetBottom",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            Console.WriteLine(
                $"  OLE hook cache active={activeField?.GetValue(addIn)}; "
                + $"rect={leftField?.GetValue(addIn)},{topField?.GetValue(addIn)},"
                + $"{rightField?.GetValue(addIn)},{bottomField?.GetValue(addIn)}.");
            window = application.ActiveWindow;
            window.GetPoint(
                out var left,
                out var top,
                out var width,
                out var height,
                range);
            if (width <= 0 || height <= 0)
                throw new InvalidDataException("Word did not return a visible OLE formula rectangle.");

            Console.WriteLine("[Word real OLE 4/6] Sending a real mouse double-click...");
            existing = SnapshotSessionIds();
            if (consoleWindow != IntPtr.Zero) ShowWindow(consoleWindow, 0);
            var wordWindowHandle = new IntPtr(window.Hwnd);
            const uint noMoveNoSizeShow = 0x0001 | 0x0002 | 0x0040;
            SetWindowPos(wordWindowHandle, new IntPtr(-1), 0, 0, 0, 0, noMoveNoSizeShow);
            var foregroundSet = SetForegroundWindow(wordWindowHandle);
            if (GetWindowRect(wordWindowHandle, out var wordWindowRectangle))
            {
                var titleX = wordWindowRectangle.Left
                    + Math.Max(40, (wordWindowRectangle.Right - wordWindowRectangle.Left) / 2);
                var titleY = wordWindowRectangle.Top + 18;
                SetCursorPos(titleX, titleY);
                mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
            }
            WinForms.Application.DoEvents();
            Thread.Sleep(600);
            Console.WriteLine($"  Word foreground request accepted={foregroundSet}.");
            var x = left + width / 2;
            var y = top + height / 2;
            SetCursorPos(x, y);
            Thread.Sleep(150);
            for (var click = 0; click < 2; click++)
            {
                mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(90);
            }
            var editSessionId = WaitForNewSession(
                existing,
                "word",
                TimeSpan.FromSeconds(30));
            var editSession = client.GetSessionAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("edit", editSession.Mode,
                "Real OLE double-click did not create an edit Session.");
            AssertEqual(FormulaOleContract.NativeOleMode, editSession.ObjectMode,
                "Real OLE double-click created the wrong object mode.");
            Console.WriteLine(
                $"  Real mouse OLE Session={editSessionId}; rectangle={left},{top},{width},{height}.");

            Console.WriteLine("[Word real OLE 5/6] Closing the unchanged editor...");
            var ready = WaitForUnchangedEditorReady(
                client,
                editSessionId,
                TimeSpan.FromSeconds(10));
            AssertEqual(false, ready.Dirty,
                "Real OLE double-click editor became dirty before input.");
            client.CloseEditorAsync(editSessionId, CancellationToken.None).GetAwaiter().GetResult();
            var closed = WaitForTerminal(client, editSessionId, TimeSpan.FromSeconds(30));
            AssertEqual("completed", closed.Status,
                closed.Error ?? "Real OLE double-click editor did not close cleanly.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            var path = Path.Combine(artifactRoot, "VisualTeX-Word-Real-OLE-DoubleClick.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine($"[Word real OLE 6/6] Saved {path}; real mouse interception passed.");
        }
        finally
        {
            if (consoleWindow != IntPtr.Zero) ShowWindow(consoleWindow, 5);
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", null);
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            if (installedAddIn is not null)
            {
                try { installedAddIn.Connect = true; } catch { }
            }
            Release(installedAddIn);
            Release(installedAddIns);
            Release(window);
            Release(range);
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            if (application is not null)
            {
                try { application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordOmmlDoubleClickFixtures(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? oleShape = null;
        COMAddIns? installedAddIns = null;
        COMAddIn? installedAddIn = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        var consoleWindow = GetConsoleWindow();
        try
        {
            Console.WriteLine("[OMML fixtures 1/8] Starting Word...");
            application = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
            installedAddIns = application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.WordVsto";
                installedAddIn = installedAddIns.Item(ref addInIndex);
                if (installedAddIn.Connect)
                    installedAddIn.Connect = false;
            }
            catch
            {
                Release(installedAddIn);
                installedAddIn = null;
            }
            document = application.Documents.Add();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            void MoveToNewLabeledParagraph(string label, bool first = false)
            {
                application.Selection.EndKey(Word.WdUnits.wdStory);
                if (!first) application.Selection.TypeParagraph();
                application.Selection.TypeText(label);
            }

            void CreateOmml(
                Action<object> command,
                string displayMode,
                string latex,
                string mathMl,
                bool numbered)
            {
                var existing = SnapshotSessionIds();
                command(new object());
                var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
                var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
                AssertEqual(FormulaOleContract.WordOmmlMode, session.ObjectMode,
                    "OMML fixture command did not request Word OMML.");
                Commit(
                    client,
                    session,
                    displayMode,
                    FormulaOleContract.WordOmmlMode,
                    latex,
                    numbered: numbered,
                    mathMl: mathMl);
                var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
                AssertEqual("completed", final.Status,
                    final.Error ?? $"OMML fixture '{latex}' did not complete.");
                client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
                WaitForAddInIdle(addIn!, TimeSpan.FromSeconds(10));
                Console.WriteLine($"  OMML inventory after '{latex}': {document.OMaths.Count}.");
            }

            void ConvertSelected(
                Action<object> command,
                string displayMode,
                string objectMode,
                string latex,
                bool numbered,
                string? mathMl = null)
            {
                var existing = SnapshotSessionIds();
                command(new object());
                var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
                var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
                AssertEqual(objectMode, session.ObjectMode,
                    $"Conversion to {objectMode} requested the wrong object mode.");
                Commit(
                    client,
                    session,
                    displayMode,
                    objectMode,
                    latex,
                    numbered: numbered,
                    dirty: false,
                    mathMl: mathMl);
                var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
                AssertEqual("completed", final.Status,
                    final.Error ?? $"Conversion to {objectMode} did not complete.");
                client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
                WaitForAddInIdle(addIn!, TimeSpan.FromSeconds(10));
            }

            void AssertNumberedTableCellNormalized(int tableIndex, string stage)
            {
                Word.Table? table = null;
                Word.Cell? centerCell = null;
                Word.Cell? numberCell = null;
                Word.Range? centerRange = null;
                Word.Range? numberRange = null;
                Word.Paragraphs? centerParagraphs = null;
                Word.Paragraphs? numberParagraphs = null;
                try
                {
                    table = document.Tables[tableIndex];
                    centerCell = table.Cell(1, 2);
                    numberCell = table.Cell(1, 3);
                    centerRange = centerCell.Range;
                    numberRange = numberCell.Range;
                    centerParagraphs = centerRange.Paragraphs;
                    numberParagraphs = numberRange.Paragraphs;
                    AssertEqual(1, centerParagraphs.Count,
                        $"{stage}: formula cell contains extra paragraphs.");
                    AssertEqual(1, numberParagraphs.Count,
                        $"{stage}: number cell contains extra paragraphs.");
                    AssertEqual(
                        Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter,
                        centerCell.VerticalAlignment,
                        $"{stage}: formula cell is not vertically centered.");
                    AssertEqual(
                        Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter,
                        numberCell.VerticalAlignment,
                        $"{stage}: number cell is not vertically centered.");
                    AssertEqual(
                        Word.WdParagraphAlignment.wdAlignParagraphCenter,
                        centerRange.ParagraphFormat.Alignment,
                        $"{stage}: formula paragraph is not horizontally centered.");
                    AssertEqual(
                        Word.WdParagraphAlignment.wdAlignParagraphRight,
                        numberRange.ParagraphFormat.Alignment,
                        $"{stage}: number paragraph is not right aligned.");
                    AssertNear(0f, numberRange.Font.Position, 0.1f,
                        $"{stage}: number has a manual baseline offset.");

                    var cellXml = XDocument.Parse(centerRange.WordOpenXML);
                    XNamespace word =
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    AssertEqual(0, cellXml.Descendants(word + "br").Count(),
                        $"{stage}: formula cell retained hidden manual line breaks.");
                }
                finally
                {
                    Release(numberParagraphs);
                    Release(centerParagraphs);
                    Release(numberRange);
                    Release(centerRange);
                    Release(numberCell);
                    Release(centerCell);
                    Release(table);
                }
            }

            Console.WriteLine("[OMML fixtures 2/7] Creating inline OMML...");
            MoveToNewLabeledParagraph("1. Inline OMML: ", first: true);
            CreateOmml(
                addIn.OnInsertInlineOmml,
                "inline",
                "x+y",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mi>y</mi></math>",
                numbered: false);

            // Regression: OMML -> OLE used to leave the caret carrying the
            // formula's raised baseline. Exercise the actual ribbon conversion,
            // type after it, then return the fixture to OMML for later checks.
            document.OMaths[1].Range.Select();
            ConvertSelected(
                addIn.OnConvertSelected,
                "inline",
                FormulaOleContract.NativeOleMode,
                "x+y",
                numbered: false);
            AssertNear(0f, application.Selection.Font.Position, 0.1f,
                "Caret after inline OMML-to-OLE conversion inherited a baseline offset.");
            var inlineSuffixStart = application.Selection.Start;
            application.Selection.TypeText(" baseline-ok");
            object inlineSuffixRangeStart = inlineSuffixStart;
            object inlineSuffixRangeEnd = application.Selection.Start;
            Word.Range? inlineSuffixRange = document.Range(
                ref inlineSuffixRangeStart,
                ref inlineSuffixRangeEnd);
            try
            {
                AssertNear(0f, inlineSuffixRange.Font.Position, 0.1f,
                    "Text after inline OMML-to-OLE conversion inherited a baseline offset.");
            }
            finally { Release(inlineSuffixRange); }
            document.InlineShapes[1].Range.Select();
            ConvertSelected(
                addIn.OnConvertSelectedToOmml,
                "inline",
                FormulaOleContract.WordOmmlMode,
                "x+y",
                numbered: false,
                mathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mi>y</mi></math>");
            AssertEqual(1, document.OMaths.Count,
                "Inline OMML-to-OLE-to-OMML round-trip lost or duplicated the equation.");

            Console.WriteLine("[OMML fixtures 3/7] Creating unnumbered display OMML...");
            MoveToNewLabeledParagraph("2. Display OMML (unnumbered):");
            CreateOmml(
                addIn.OnInsertDisplayOmml,
                "block",
                "\\frac{a+b}{c+d}",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mfrac><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mrow><mi>c</mi><mo>+</mo><mi>d</mi></mrow></mfrac></math>",
                numbered: false);

            Console.WriteLine("[OMML fixtures 4/7] Creating numbered display OMML...");
            MoveToNewLabeledParagraph("3. Display OMML (numbered):");
            CreateOmml(
                addIn.OnInsertDisplayOmml,
                "block",
                "\\sum_{n=1}^{\\infty}\\frac{1}{n^2}",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><munderover><mo>∑</mo><mrow><mi>n</mi><mo>=</mo><mn>1</mn></mrow><mi>∞</mi></munderover><mfrac><mn>1</mn><msup><mi>n</mi><mn>2</mn></msup></mfrac></mrow></math>",
                numbered: true);

            var naryCases = new (string Latex, string MathMl)[]
            {
                ("\\sum_b^z c", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><munderover><mo>∑</mo><mi>b</mi><mi>z</mi></munderover><mi>c</mi></math>"),
                ("\\sum_{b}^{z} c", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><munderover><mo>∑</mo><mrow><mi>b</mi></mrow><mrow><mi>z</mi></mrow></munderover><mi>c</mi></math>"),
                ("\\oint_l^u x\\,dy", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msubsup><mo>∮</mo><mi>l</mi><mi>u</mi></msubsup><mi>x</mi><mstyle><mspace width=\"0.167em\"/></mstyle><mi>d</mi><mi>y</mi></math>"),
                ("\\oint_l x\\,dy", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msub><mo>∮</mo><mi>l</mi></msub><mi>x</mi><mstyle><mspace width=\"0.167em\"/></mstyle><mi>d</mi><mi>y</mi></math>"),
                ("\\int_0^1 x^2\\,dx", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup><msup><mi>x</mi><mn>2</mn></msup><mstyle><mspace width=\"0.167em\"/></mstyle><mi>d</mi><mi>x</mi></math>"),
                ("\\prod_{i=1}^{n} a_i", "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><munderover><mo>∏</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mrow><mi>n</mi></mrow></munderover><msub><mi>a</mi><mi>i</mi></msub></math>"),
            };
            foreach (var (latex, mathMl) in naryCases)
            {
                MoveToNewLabeledParagraph($"N-ary regression: {latex}");
                CreateOmml(
                    addIn.OnInsertDisplayOmml,
                    "block",
                    latex,
                    mathMl,
                    numbered: false);
            }

            Console.WriteLine("[OMML fixtures 5/7] Creating OLE then converting it to OMML...");
            MoveToNewLabeledParagraph("4. Numbered display OLE converted to OMML: ");
            var existing = SnapshotSessionIds();
            addIn.OnInsertDisplay(new object());
            var oleSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var oleSession = client.GetSessionAsync(oleSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(
                client,
                oleSession,
                "block",
                FormulaOleContract.NativeOleMode,
                "\\int_0^1 x^2\\,dx",
                numbered: true);
            var finalOle = WaitForTerminal(client, oleSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", finalOle.Status,
                finalOle.Error ?? "OLE fixture creation did not complete.");
            client.CloseEditorAsync(oleSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            oleShape = document.InlineShapes[document.InlineShapes.Count];
            AssertNumberedTableCellNormalized(2, "new numbered OLE");
            oleShape.Range.Select();
            existing = SnapshotSessionIds();
            addIn.OnConvertSelectedToOmml(new object());
            var conversionSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var conversionSession = client.GetSessionAsync(conversionSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual(FormulaOleContract.WordOmmlMode, conversionSession.ObjectMode,
                "OLE to OMML fixture conversion requested the wrong object mode.");
            Commit(
                client,
                conversionSession,
                "block",
                FormulaOleContract.WordOmmlMode,
                "\\int_0^1 x^2\\,dx",
                numbered: true,
                dirty: false,
                mathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup><msup><mi>x</mi><mn>2</mn></msup><mi>d</mi><mi>x</mi></math>");
            var converted = WaitForTerminal(client, conversionSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", converted.Status,
                converted.Error ?? "OLE to OMML fixture conversion did not complete.");
            client.CloseEditorAsync(conversionSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            Console.WriteLine($"  OMML inventory after numbered OLE conversion: {document.OMaths.Count}.");
            AssertNumberedTableCellNormalized(2, "initial OLE-to-OMML conversion");
            Release(oleShape);
            oleShape = null;

            // Regression: a numbered OMML -> OLE -> OMML round-trip used to
            // create a nested table and several empty paragraphs. Inserting the
            // next display formula could then remove the previous equation.
            for (var round = 1; round <= 3; round++)
            {
                document.OMaths[document.OMaths.Count].Range.Select();
                ConvertSelected(
                    addIn.OnConvertSelected,
                    "block",
                    FormulaOleContract.NativeOleMode,
                    "\\int_0^1 x^2\\,dx",
                    numbered: true);
                AssertEqual(2, document.Tables.Count,
                    $"Round {round} OMML-to-OLE created an extra equation table.");
                AssertEqual(9, document.OMaths.Count,
                    $"Round {round} OMML-to-OLE did not replace exactly one OMML equation.");
                AssertNumberedTableCellNormalized(
                    2,
                    $"round {round} OMML-to-OLE");

                document.InlineShapes[document.InlineShapes.Count].Range.Select();
                ConvertSelected(
                    addIn.OnConvertSelectedToOmml,
                    "block",
                    FormulaOleContract.WordOmmlMode,
                    "\\int_0^1 x^2\\,dx",
                    numbered: true,
                    mathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msubsup><mo>\u222b</mo><mn>0</mn><mn>1</mn></msubsup><msup><mi>x</mi><mn>2</mn></msup><mi>d</mi><mi>x</mi></math>");
                AssertEqual(2, document.Tables.Count,
                    $"Round {round} OLE-to-OMML created a nested/extra equation table.");
                AssertEqual(10, document.OMaths.Count,
                    $"Round {round} lost or duplicated the original display equation.");
                AssertNumberedTableCellNormalized(
                    2,
                    $"round {round} OLE-to-OMML");
            }

            MoveToNewLabeledParagraph("5. Display OMML inserted after round-trip:");
            CreateOmml(
                addIn.OnInsertDisplayOmml,
                "block",
                "a+b=c",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>a</mi><mo>+</mo><mi>b</mi><mo>=</mo><mi>c</mi></mrow></math>",
                numbered: true);
            AssertEqual(11, document.OMaths.Count,
                "Inserting the next display equation removed the round-tripped equation.");

            Console.WriteLine("[OMML fixtures 6/8] Validating formula inventory...");
            const int expectedEquationCount = 11;
            if (document.OMaths.Count != expectedEquationCount)
                throw new InvalidDataException(
                    $"OMML double-click fixture document should contain {expectedEquationCount} equations, actual: {document.OMaths.Count}.");
            XNamespace mathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            XNamespace wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            for (var equationIndex = 1; equationIndex <= document.OMaths.Count; equationIndex++)
            {
                Word.OMath? equation = null;
                Word.Range? equationRange = null;
                try
                {
                    equation = document.OMaths[equationIndex];
                    equationRange = equation.Range;
                    var xml = XDocument.Parse(equationRange.WordOpenXML);
                    if (xml.Descendants(wordNamespace + "color")
                            .Any(element => string.Equals(
                                (string?)element.Attribute(wordNamespace + "val"),
                                "FFFFFF",
                                StringComparison.OrdinalIgnoreCase))
                        || xml.Descendants(wordNamespace + "sz")
                            .Any(element => (string?)element.Attribute(wordNamespace + "val") == "2"))
                        throw new InvalidDataException(
                            $"Equation {equationIndex} inherited the hidden-caption white 1pt style.");
                    foreach (var nary in xml.Descendants(mathNamespace + "nary"))
                    {
                        var operand = nary.Element(mathNamespace + "e");
                        if (operand is null || !operand.Elements().Any())
                            throw new InvalidDataException(
                                $"Equation {equationIndex} contains an empty n-ary operand.");
                    }
                    if (equationIndex > 1 && equation.Type != Word.WdOMathType.wdOMathDisplay)
                        throw new InvalidDataException(
                            $"Display equation {equationIndex} degraded to inline OMath.");
                }
                finally
                {
                    Release(equationRange);
                    Release(equation);
                }
            }

            if (document.Tables.Count != 3 || document.Fields.Count != 6)
                throw new InvalidDataException(
                    $"Expected three numbered equation tables with SEQ+REF fields; "
                    + $"tables={document.Tables.Count}, fields={document.Fields.Count}.");
            for (var tableIndex = 1; tableIndex <= 3; tableIndex++)
            {
                Word.Table? table = null;
                Word.Cell? numberCell = null;
                Word.Range? numberCellRange = null;
                Word.Paragraphs? numberCellParagraphs = null;
                try
                {
                    table = document.Tables[tableIndex];
                    numberCell = table.Cell(1, 3);
                    numberCellRange = numberCell.Range;
                    numberCellParagraphs = numberCellRange.Paragraphs;
                    AssertEqual(
                        1,
                        numberCellParagraphs.Count,
                        $"Numbered equation table {tableIndex} contains an extra empty paragraph that shifts the number down.");
                    AssertNear(0f, numberCellRange.Font.Position, 0.1f,
                        $"Numbered equation table {tableIndex} applies a manual baseline shift instead of cell centering.");
                    var visibleNumber = (numberCellRange.Text ?? string.Empty)
                        .Trim('\r', '\a', ' ');
                    AssertEqual(
                        $"({tableIndex})",
                        visibleNumber,
                        $"Numbered equation table {tableIndex} has no visible right-aligned number.");
                }
                finally
                {
                    Release(numberCellParagraphs);
                    Release(numberCellRange);
                    Release(numberCell);
                    Release(table);
                }
            }

            var path = Path.Combine(artifactRoot, "VisualTeX-Word-OMML-DoubleClick-Fixtures.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            Console.WriteLine("[OMML fixtures 7/8] Reopening and exercising real native double-click editing...");
            document = application.Documents.Open(path, ReadOnly: false, Visible: true);
            application.Visible = true;
            Word.Window? focusWindow = null;
            try
            {
                focusWindow = application.ActiveWindow;
                focusWindow.Activate();
                if (consoleWindow != IntPtr.Zero) ShowWindow(consoleWindow, 0);
                var wordWindowHandle = new IntPtr(focusWindow.Hwnd);
                const uint noMoveNoSizeShow = 0x0001 | 0x0002 | 0x0040;
                SetWindowPos(wordWindowHandle, new IntPtr(-1), 0, 0, 0, 0, noMoveNoSizeShow);
                SetForegroundWindow(wordWindowHandle);
                if (GetWindowRect(wordWindowHandle, out var wordWindowRectangle))
                {
                    SetCursorPos(
                        wordWindowRectangle.Left
                            + Math.Max(40, (wordWindowRectangle.Right - wordWindowRectangle.Left) / 2),
                        wordWindowRectangle.Top + 18);
                    mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                    mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                }
                WinForms.Application.DoEvents();
                Thread.Sleep(600);
            }
            finally
            {
                Release(focusWindow);
            }

            for (var equationIndex = 1; equationIndex <= expectedEquationCount; equationIndex++)
            {
                Word.OMaths? maths = null;
                Word.OMath? math = null;
                Word.Range? equationRange = null;
                Word.Window? window = null;
                Word.OMaths? selectedMaths = null;
                try
                {
                    maths = document.OMaths;
                    math = maths[equationIndex];
                    equationRange = math.Range;
                    var beforeXml = equationRange.WordOpenXML;
                    equationRange.Select();
                    WinForms.Application.DoEvents();
                    Thread.Sleep(180);
                    window = application.ActiveWindow;
                    window.GetPoint(
                        out var left,
                        out var top,
                        out var width,
                        out var height,
                        equationRange);
                    if (width <= 0 || height <= 0)
                        throw new InvalidDataException(
                            $"Word did not return a visible rectangle for OMML equation {equationIndex}.");

                    var sessionsBefore = SnapshotSessionIds();
                    SetCursorPos(left + width / 2, top + height / 2);
                    Thread.Sleep(120);
                    for (var click = 0; click < 2; click++)
                    {
                        mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                        mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                        Thread.Sleep(90);
                    }

                    var deadline = DateTime.UtcNow.AddSeconds(5);
                    var nativeCaretEntered = false;
                    while (DateTime.UtcNow < deadline)
                    {
                        Release(selectedMaths);
                        selectedMaths = application.Selection.OMaths;
                        if (selectedMaths.Count == 1)
                        {
                            nativeCaretEntered = true;
                            break;
                        }
                        WinForms.Application.DoEvents();
                        Thread.Sleep(60);
                    }
                    if (!nativeCaretEntered)
                        throw new InvalidDataException(
                            $"Real double-click did not enter Word's native editor for OMML equation {equationIndex}.");

                    application.Selection.TypeText("q");
                    WinForms.Application.DoEvents();
                    Thread.Sleep(150);
                    Release(equationRange);
                    equationRange = null;
                    Release(math);
                    math = null;
                    Release(maths);
                    maths = document.OMaths;
                    math = maths[equationIndex];
                    equationRange = math.Range;
                    if (string.Equals(beforeXml, equationRange.WordOpenXML, StringComparison.Ordinal))
                        throw new InvalidDataException(
                            $"Typing after real double-click did not change OMML equation {equationIndex}.");
                    document.Undo(1);
                    WinForms.Application.DoEvents();
                    Thread.Sleep(120);

                    var sessionsAfter = SnapshotSessionIds();
                    if (sessionsAfter.Except(sessionsBefore, StringComparer.OrdinalIgnoreCase).Any())
                        throw new InvalidDataException(
                            $"OMML equation {equationIndex} incorrectly opened a VisualTeX Session.");
                    Console.WriteLine(
                        $"  OMML equation {equationIndex}: native caret entered, input changed OMML, undo passed.");
                }
                finally
                {
                    Release(selectedMaths);
                    Release(window);
                    Release(equationRange);
                    Release(math);
                    Release(maths);
                }
            }

            document.Save();
            Console.WriteLine($"[OMML fixtures 8/8] Saved {path}; current-source real mouse checks passed.");
        }
        finally
        {
            if (consoleWindow != IntPtr.Zero) ShowWindow(consoleWindow, 5);
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            if (installedAddIn is not null)
            {
                try { installedAddIn.Connect = true; } catch { }
            }
            Release(installedAddIn);
            Release(installedAddIns);
            Release(oleShape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            if (application is not null)
            {
                try { application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWord(
        VisualTeXSessionClient client,
        string artifactRoot,
        bool initialOnly = false,
        bool stopAfterUnchanged = false)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.OLEFormat? wordOleFormat = null;
        Word.Range? typedRange = null;
        Word.Selection? eventSelection = null;
        Word.InlineShape? numberedShape = null;
        COMAddIns? installedAddIns = null;
        COMAddIn? installedAddIn = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            Console.WriteLine("[Word 1/12] Starting Word and creating an inline Session...");
            application = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
            installedAddIns = application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.WordVsto";
                installedAddIn = installedAddIns.Item(ref addInIndex);
                if (installedAddIn.Connect)
                    installedAddIn.Connect = false;
            }
            catch
            {
                Release(installedAddIn);
                installedAddIn = null;
            }
            document = application.Documents.Add();
            application.Selection.TypeText("Before ");
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);
            var existing = SnapshotSessionIds();
            addIn.OnInsertInline(new object());
            var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("nativeOle", session.ObjectMode, "Word must request native OLE.");
            AssertEqual("inline", session.DisplayMode, "Word inline Session mode is wrong.");

            Console.WriteLine("[Word 2/12] Committing the initial vector formula...");
            Commit(client, session, "inline", "nativeOle", "x+y=z");
            var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status, final.Error ?? "Word Session did not complete.");
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();

            Console.WriteLine("[Word 3/12] Checking OLE size, aspect ratio and formula baseline...");
            AssertEqual(1, document.InlineShapes.Count, "Word should contain one inline formula.");
            shape = document.InlineShapes[1];
            AssertNear(120f, shape.Width, 0.5f, "Word formula width is incorrect.");
            AssertNear(24f, shape.Height, 0.5f, "Word formula height is incorrect.");
            AssertNear(5f, shape.Width / shape.Height, 0.05f,
                "Word formula aspect ratio is distorted.");
            var expectedPosition = WordInlineAlignment.CalculateFontPosition(
                shape.Height,
                ExportHeight,
                ExportBaseline);
            AssertNear(expectedPosition, shape.Range.Font.Position, 0.1f,
                "Word inline formula baseline is incorrect.");
            AssertEqual(FormulaOleContract.ProgId, shape.OLEFormat.ProgID,
                "Word inserted the wrong OLE class.");
            if (initialOnly)
            {
                var probePath = Path.Combine(artifactRoot, "VisualTeX-Word-Create-Probe.docx");
                document.SaveAs2(probePath, Word.WdSaveFormat.wdFormatXMLDocument);
                Console.WriteLine($"[Word probe] Saved {probePath}; initial OLE creation passed.");
                return;
            }

            Console.WriteLine("[Word 4/12] Verifying that typing after the formula returns to the text baseline...");
            AssertNear(0f, application.Selection.Font.Position, 0.1f,
                "Word caret inherited the formula baseline offset.");
            var textStart = application.Selection.Start;
            application.Selection.TypeText(" after");
            object rangeStart = textStart;
            object rangeEnd = application.Selection.Start;
            typedRange = document.Range(ref rangeStart, ref rangeEnd);
            AssertNear(0f, typedRange.Font.Position, 0.1f,
                "Text typed after the formula inherited the formula baseline offset.");
            Release(typedRange);
            typedRange = null;

            Console.WriteLine("[Word 5/12] Closing an unchanged edit and waiting for the add-in to unlock...");
            shape.Range.Select();
            existing = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            var unchangedSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var unchangedSession = WaitForUnchangedEditorReady(
                client,
                unchangedSessionId,
                TimeSpan.FromSeconds(10));
            AssertEqual(false, unchangedSession.Dirty,
                "Word unchanged edit Session became dirty before closing.");
            client.CloseEditorAsync(unchangedSessionId, CancellationToken.None).GetAwaiter().GetResult();
            final = WaitForTerminal(client, unchangedSessionId, TimeSpan.FromSeconds(30));
            AssertEqual("completed", final.Status,
                final.Error ?? "Word unchanged edit did not complete after closing the window.");
            AssertEqual(false, final.Dirty,
                "Word unchanged edit was incorrectly marked dirty.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            Console.WriteLine("[Word 6/12] Reopening immediately through the double-click interception...");
            shape.Range.Select();
            existing = SnapshotSessionIds();
            eventSelection = application.Selection;
            var handler = typeof(VisualTeX.WordVsto.ThisAddIn).GetMethod(
                "OnWindowBeforeDoubleClick",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                ?? throw new MissingMethodException("Word double-click handler is missing.");
            var handlerArguments = new object[] { eventSelection, false };
            handler.Invoke(addIn, handlerArguments);
            AssertEqual(true, (bool)handlerArguments[1],
                "Word did not suppress built-in OLE activation on double-click.");
            Release(eventSelection);
            eventSelection = null;
            var editSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var editSession = client.GetSessionAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("edit", editSession.Mode, "Word double-click did not create an edit Session.");
            AssertEqual("nativeOle", editSession.ObjectMode,
                "Word edit Session changed the object mode.");
            if (stopAfterUnchanged)
            {
                editSession = WaitForUnchangedEditorReady(
                    client,
                    editSessionId,
                    TimeSpan.FromSeconds(10));
                AssertEqual(false, editSession.Dirty,
                    "Word reopened edit Session was already dirty.");
                client.CloseEditorAsync(editSessionId, CancellationToken.None).GetAwaiter().GetResult();
                final = WaitForTerminal(client, editSessionId, TimeSpan.FromSeconds(30));
                AssertEqual("completed", final.Status,
                    final.Error ?? "Word reopened unchanged edit did not complete.");
                AssertEqual(false, final.Dirty,
                    "Word reopened unchanged edit was incorrectly marked dirty.");
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
                var probePath = Path.Combine(artifactRoot, "VisualTeX-Word-Unchanged-Probe.docx");
                document.SaveAs2(probePath, Word.WdSaveFormat.wdFormatXMLDocument);
                Console.WriteLine($"[Word unchanged probe] Saved {probePath}; close and immediate reopen passed.");
                return;
            }

            Console.WriteLine("[Word 7/12] Editing to a wider formula and checking natural resize...");
            Commit(
                client,
                editSession,
                "inline",
                "nativeOle",
                "x+y+z+a+b+c=d",
                renderWidth: 320f,
                renderHeight: 32f,
                baseline: 24f);
            final = WaitForTerminal(client, editSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status, final.Error ?? "Word edit Session did not complete.");
            client.CloseEditorAsync(editSessionId, CancellationToken.None).GetAwaiter().GetResult();
            Release(shape);
            shape = document.InlineShapes[1];
            Console.WriteLine(
                $"  Word edited OLE geometry: {shape.Width:F3} x {shape.Height:F3} pt, "
                + $"ratio={shape.Width / shape.Height:F4}");
            AssertNear(240f, shape.Width, 1.5f,
                "Word edited formula retained the old width.");
            AssertNear(24f, shape.Height, 0.5f,
                "Word edited formula height is incorrect.");
            AssertNear(10f, shape.Width / shape.Height, 0.08f,
                "Word edited formula is compressed into the old aspect ratio.");
            AssertNear(0f, application.Selection.Font.Position, 0.1f,
                "Word caret baseline was not restored after editing.");

            Console.WriteLine("[Word 8/12] Exporting the selected OLE formula as an editable picture...");
            shape.Range.Select();
            Release(shape);
            shape = null;
            addIn.OnExportSelectedAsPicture(new object());
            shape = WaitForWordShapeMode(document, nativeOle: false, TimeSpan.FromSeconds(15));
            AssertEqual(Word.WdInlineShapeType.wdInlineShapePicture, shape.Type,
                "Word OLE export did not create a normal picture.");

            Console.WriteLine("[Word 9/12] Converting the unchanged picture back to native OLE...");
            shape.Range.Select();
            existing = SnapshotSessionIds();
            addIn.OnConvertSelected(new object());
            var convertSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var convertSession = client.GetSessionAsync(convertSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("edit", convertSession.Mode,
                "Word convert command did not create an edit Session.");
            AssertEqual("nativeOle", convertSession.ObjectMode,
                "Word convert command did not request native OLE.");
            Commit(
                client,
                convertSession,
                "inline",
                "nativeOle",
                "x+y+z+a+b+c=d",
                renderWidth: 320f,
                renderHeight: 32f,
                baseline: 24f,
                dirty: false);
            final = WaitForTerminal(client, convertSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "Word OLE conversion did not complete.");
            client.CloseEditorAsync(convertSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            Release(shape);
            shape = WaitForWordShapeMode(document, nativeOle: true, TimeSpan.FromSeconds(15));
            wordOleFormat = shape.OLEFormat;
            AssertEqual(FormulaOleContract.ProgId, wordOleFormat.ProgID,
                "Word converted to the wrong OLE class.");

            Console.WriteLine("[Word 10/12] Exercising the converted OLE show verb...");
            wordOleFormat.DoVerb(-1);
            Release(wordOleFormat);
            wordOleFormat = null;

            Console.WriteLine("[Word 11/17] Rechecking size after picture-to-OLE conversion...");
            AssertNear(240f, shape.Width, 1.5f,
                "Word picture-to-OLE conversion changed the formula width.");
            AssertNear(24f, shape.Height, 0.5f,
                "Word picture-to-OLE conversion changed the formula height.");
            Release(shape);
            shape = null;

            Console.WriteLine("[Word 12/17] Creating the first numbered display formula...");
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            existing = SnapshotSessionIds();
            addIn.OnInsertDisplay(new object());
            var firstNumberSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var firstNumberSession = client.GetSessionAsync(firstNumberSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(
                client,
                firstNumberSession,
                "block",
                "nativeOle",
                "a=b",
                numbered: true);
            final = WaitForTerminal(client, firstNumberSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "First numbered Word formula did not complete.");
            client.CloseEditorAsync(firstNumberSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            var firstNumberFormulaId = final.FormulaId
                ?? throw new InvalidDataException("First numbered formula has no formulaId.");
            WaitForWordInlineShapeCount(document, 2, TimeSpan.FromSeconds(15));

            Console.WriteLine("[Word 13/17] Creating the second numbered display formula...");
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            existing = SnapshotSessionIds();
            addIn.OnInsertDisplay(new object());
            var secondNumberSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var secondNumberSession = client.GetSessionAsync(secondNumberSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(
                client,
                secondNumberSession,
                "block",
                "nativeOle",
                "E=mc^2",
                numbered: true);
            final = WaitForTerminal(client, secondNumberSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "Second numbered Word formula did not complete.");
            client.CloseEditorAsync(secondNumberSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            var secondNumberFormulaId = final.FormulaId
                ?? throw new InvalidDataException("Second numbered formula has no formulaId.");
            WaitForWordInlineShapeCount(document, 3, TimeSpan.FromSeconds(15));

            Console.WriteLine("[Word 14/17] Inserting a live reference to equation (2)...");
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            application.Selection.TypeText("See ");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REFERENCE_TARGET_INDEX", "1");
            var nativeItems = document.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionEquation) as Array;
            if (nativeItems is null || nativeItems.Length != 2)
                throw new InvalidDataException(
                    $"Word native Equation list should contain two VisualTeX formulas, actual count: {nativeItems?.Length ?? 0}.");
            addIn.OnInsertEquationReference(new object());
            var nativeReferenceCode = WaitForWordNativeReferenceResult(
                document,
                expectedResult: "2",
                expectedCode: null,
                TimeSpan.FromSeconds(15));
            if (!DocumentTextContains(document, "(2)"))
                throw new InvalidDataException("Word native reference did not include parenthesized equation number (2).");

            Console.WriteLine("[Word 15/17] Deleting equation (1) through the VisualTeX command...");
            numberedShape = document.InlineShapes[2];
            numberedShape.Range.Select();
            addIn.OnDeleteSelected(new object());
            Release(numberedShape);
            numberedShape = null;
            WaitForWordInlineShapeCount(document, 2, TimeSpan.FromSeconds(15));

            Console.WriteLine("[Word 16/17] Updating equation numbers and the REF field...");
            addIn.OnUpdateEquationNumbers(new object());
            WaitForWordNativeReferenceResult(
                document,
                expectedResult: "1",
                expectedCode: nativeReferenceCode,
                TimeSpan.FromSeconds(15));
            if (!DocumentTextContains(document, "(1)"))
                throw new InvalidDataException("Word native reference did not update to equation number (1).");
            if (WordBookmarkExists(document, $"VTEq_{Guid.Parse(firstNumberFormulaId):N}"))
                throw new InvalidDataException("Deleted equation retained its VisualTeX number bookmark.");
            if (!WordBookmarkExists(document, $"VTEq_{Guid.Parse(secondNumberFormulaId):N}"))
                throw new InvalidDataException("Referenced equation lost its persistent VisualTeX number bookmark.");
            nativeItems = document.GetCrossReferenceItems(Word.WdCaptionLabelID.wdCaptionEquation) as Array;
            if (nativeItems is null || nativeItems.Length != 1)
                throw new InvalidDataException(
                    $"Word native Equation list should contain one formula after deletion, actual count: {nativeItems?.Length ?? 0}.");

            Console.WriteLine("[Word 17/21] Creating a real Word OMML formula through the VisualTeX editor...");
            const string ommlInitialMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
                + "<msup><mi>x</mi><mn>2</mn></msup><mo>+</mo>"
                + "<msup><mi>y</mi><mn>2</mn></msup></math>";
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            existing = SnapshotSessionIds();
            addIn.OnInsertDisplayOmml(new object());
            var ommlCreateSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var ommlCreateSession = client.GetSessionAsync(ommlCreateSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("wordOmml", ommlCreateSession.ObjectMode,
                "Word OMML insert command did not create a wordOmml Session.");
            Commit(
                client,
                ommlCreateSession,
                "block",
                "wordOmml",
                "x^2+y^2",
                numbered: false,
                mathMl: ommlInitialMathMl);
            final = WaitForTerminal(client, ommlCreateSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "Word OMML creation did not complete.");
            client.CloseEditorAsync(ommlCreateSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            Console.WriteLine("[Word 18/21] Adding +z^3 through Word's native equation object...");
            AppendToLastWordOmmlAndSelect(document, "+z^3");

            Console.WriteLine("[Word 19/21] Opening the selected native-edited OMML through VisualTeX...");
            existing = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            var ommlEditSessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var ommlEditSession = WaitForUnchangedEditorReady(
                client,
                ommlEditSessionId,
                TimeSpan.FromSeconds(10));
            AssertEqual("edit", ommlEditSession.Mode,
                "Native-edited OMML did not open an edit Session.");
            AssertEqual("wordOmml", ommlEditSession.ObjectMode,
                "Native-edited OMML opened with the wrong object mode.");
            var importedNativeLatex = string.Join("\n", ommlEditSession.Lines.Select(line => line.Latex));
            if (importedNativeLatex.IndexOf("z", StringComparison.Ordinal) < 0
                || importedNativeLatex.IndexOf("^{3}", StringComparison.Ordinal) < 0)
                throw new InvalidDataException(
                    "VisualTeX editor Session did not include the +z^3 inserted by Word's native equation editor. "
                    + $"Imported source: {importedNativeLatex}");

            Console.WriteLine("[Word 20/21] Closing the unchanged OMML editor after verifying imported source...");
            client.CloseEditorAsync(ommlEditSessionId, CancellationToken.None).GetAwaiter().GetResult();
            final = WaitForTerminal(client, ommlEditSessionId, TimeSpan.FromSeconds(30));
            AssertEqual("completed", final.Status,
                final.Error ?? "Native-edited OMML unchanged edit did not complete.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            var path = Path.Combine(artifactRoot, "VisualTeX-Word-Flow.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                $"[Word 21/21] Saved {path}; conversion, foreground editor, live cross-reference, and native OMML source-import checks passed.");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            if (installedAddIn is not null)
            {
                try { installedAddIn.Connect = true; } catch { }
            }
            Release(installedAddIn);
            Release(installedAddIns);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REFERENCE_TARGET_INDEX", null);
            Release(numberedShape);
            Release(eventSelection);
            Release(typedRange);
            Release(wordOleFormat);
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            if (application is not null)
            {
                try { application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunPowerPoint(VisualTeXSessionClient client, string artifactRoot)
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        PowerPoint.OLEFormat? oleFormat = null;
        VisualTeX.PowerPointVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            Console.WriteLine("[PowerPoint 1/10] Starting PowerPoint and creating a formula Session...");
            application = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            addIn = new VisualTeX.PowerPointVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);
            var existing = SnapshotSessionIds();
            addIn.OnNewFormula(new object());
            var sessionId = WaitForNewSession(existing, "powerpoint", TimeSpan.FromSeconds(30));
            var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("crossPlatformPicture", session.ObjectMode,
                "PowerPoint must use the stable editable picture path by default.");

            Console.WriteLine("[PowerPoint 2/10] Committing the initial picture formula...");
            Commit(client, session, "block", "crossPlatformPicture", "x+y=z");
            var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status, final.Error ?? "PowerPoint Session did not complete.");
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();

            Console.WriteLine("[PowerPoint 3/10] Checking the inserted editable picture...");
            AssertEqual(1, slide.Shapes.Count, "PowerPoint should contain one formula shape.");
            shape = slide.Shapes[1];
            AssertEqual(MsoShapeType.msoPicture, shape.Type,
                "PowerPoint formula must be a picture, not an OLE placeholder.");
            AssertNear(120f, shape.Width, 0.5f, "PowerPoint formula width is incorrect.");
            AssertNear(24f, shape.Height, 0.5f, "PowerPoint formula height is incorrect.");
            AssertNear(5f, shape.Width / shape.Height, 0.05f,
                "PowerPoint formula aspect ratio is distorted.");
            ReportPowerPointMetadata("initial", shape);

            Console.WriteLine("[PowerPoint create reset] Creating again while the previous formula remains selected...");
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            addIn.OnNewFormula(new object());
            var blankCreateSessionId = WaitForNewSession(
                existing,
                "powerpoint",
                TimeSpan.FromSeconds(30));
            var blankCreateSession = client.GetSessionAsync(
                    blankCreateSessionId,
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("create", blankCreateSession.Mode,
                "PowerPoint New Formula reused the previous formula as an edit Session.");
            AssertEqual(1, blankCreateSession.Lines.Count,
                "PowerPoint New Formula should start with exactly one empty line.");
            AssertEqual(string.Empty, blankCreateSession.Lines[0].Latex,
                "PowerPoint New Formula inherited LaTeX from the selected previous formula.");
            if (string.Equals(blankCreateSession.FormulaId, final.FormulaId, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "PowerPoint New Formula reused the selected formulaId.");
            Commit(client, blankCreateSession, "block", "crossPlatformPicture", "u=v");
            var blankCreateFinal = WaitForTerminal(
                client,
                blankCreateSessionId,
                TimeSpan.FromSeconds(45));
            AssertEqual("completed", blankCreateFinal.Status,
                blankCreateFinal.Error ?? "Second blank PowerPoint create Session did not complete.");
            client.CloseEditorAsync(blankCreateSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            AssertEqual(2, slide.Shapes.Count,
                "PowerPoint second create did not insert an independent formula.");
            slide.Shapes[2].Delete();
            Release(shape);
            shape = slide.Shapes[1];

            Console.WriteLine("[PowerPoint 4/10] Closing an unchanged edit and waiting for the add-in to unlock...");
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            var unchangedSessionId = WaitForNewSession(existing, "powerpoint", TimeSpan.FromSeconds(30));
            var unchangedSession = WaitForUnchangedEditorReady(
                client,
                unchangedSessionId,
                TimeSpan.FromSeconds(10));
            AssertEqual(false, unchangedSession.Dirty,
                "PowerPoint unchanged edit Session became dirty before closing.");
            client.CloseEditorAsync(unchangedSessionId, CancellationToken.None).GetAwaiter().GetResult();
            final = WaitForTerminal(client, unchangedSessionId, TimeSpan.FromSeconds(30));
            AssertEqual("completed", final.Status,
                final.Error ?? "PowerPoint unchanged edit did not complete after closing the window.");
            AssertEqual(false, final.Dirty,
                "PowerPoint unchanged edit was incorrectly marked dirty.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            Console.WriteLine("[PowerPoint 5/10] Reopening immediately and editing through the Ribbon callback...");
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            var editSessionId = WaitForNewSession(existing, "powerpoint", TimeSpan.FromSeconds(30));
            var editSession = client.GetSessionAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("edit", editSession.Mode,
                "PowerPoint edit button did not create an edit Session.");
            AssertEqual("crossPlatformPicture", editSession.ObjectMode,
                "PowerPoint picture edit changed the object mode.");
            Commit(
                client,
                editSession,
                "block",
                "crossPlatformPicture",
                "x+y+z+a+b+c=d",
                renderWidth: 320f,
                renderHeight: 32f,
                baseline: 24f);
            final = WaitForTerminal(client, editSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "PowerPoint edit Session did not complete.");
            client.CloseEditorAsync(editSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            Release(shape);
            shape = slide.Shapes[1];
            AssertEqual(MsoShapeType.msoPicture, shape.Type,
                "PowerPoint edit unexpectedly changed the picture object type.");
            AssertNear(240f, shape.Width, 0.8f,
                "PowerPoint edited formula retained the old width.");
            AssertNear(24f, shape.Height, 0.5f,
                "PowerPoint edited formula height is incorrect.");
            AssertNear(10f, shape.Width / shape.Height, 0.08f,
                "PowerPoint edited formula is compressed into the old aspect ratio.");
            ReportPowerPointMetadata("edited", shape);

            Console.WriteLine("[PowerPoint 6/10] Exercising the double-click edit callback...");
            shape.Select(MsoTriState.msoTrue);
            existing = SnapshotSessionIds();
            var doubleClickHandler = typeof(VisualTeX.PowerPointVsto.ThisAddIn).GetMethod(
                "OnNativeDoubleClick",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                ?? throw new MissingMethodException("PowerPoint double-click callback is missing.");
            doubleClickHandler.Invoke(addIn, null);
            var doubleClickSessionId = WaitForNewSession(
                existing,
                "powerpoint",
                TimeSpan.FromSeconds(30));
            var doubleClickSession = client.GetSessionAsync(
                    doubleClickSessionId,
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("edit", doubleClickSession.Mode,
                "PowerPoint double-click did not create an edit Session.");
            Commit(
                client,
                doubleClickSession,
                "block",
                "crossPlatformPicture",
                "x+y+z+a+b+c=d",
                renderWidth: 320f,
                renderHeight: 32f,
                baseline: 24f,
                dirty: false);
            final = WaitForTerminal(client, doubleClickSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "PowerPoint double-click Session did not complete.");
            client.CloseEditorAsync(doubleClickSessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            Console.WriteLine("[PowerPoint 7/10] Converting the selected formula to native OLE...");
            application.ActiveWindow.Activate();
            application.ActiveWindow.View.GotoSlide(1);
            shape.Select(MsoTriState.msoTrue);
            WinForms.Application.DoEvents();
            Thread.Sleep(150);
            existing = SnapshotSessionIds();
            addIn.OnConvertSelected(new object());
            var convertSessionId = WaitForNewSession(
                existing,
                "powerpoint",
                TimeSpan.FromSeconds(30),
                () => string.IsNullOrWhiteSpace(addIn.DiagnosticLastError)
                    ? null
                    : Convert.ToBase64String(
                        Encoding.UTF8.GetBytes(addIn.DiagnosticLastError)));
            var convertSession = client.GetSessionAsync(convertSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("nativeOle", convertSession.ObjectMode,
                "PowerPoint convert command did not request native OLE.");
            Commit(
                client,
                convertSession,
                "block",
                "nativeOle",
                "x+y+z+a+b+c=d",
                renderWidth: 320f,
                renderHeight: 32f,
                baseline: 24f,
                dirty: false);
            final = WaitForTerminal(client, convertSessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "PowerPoint OLE conversion did not complete.");
            client.CloseEditorAsync(convertSessionId, CancellationToken.None).GetAwaiter().GetResult();
            Release(shape);
            shape = slide.Shapes[1];
            AssertEqual(MsoShapeType.msoEmbeddedOLEObject, shape.Type,
                "PowerPoint convert command did not create an embedded OLE object.");
            oleFormat = shape.OLEFormat;
            AssertEqual(FormulaOleContract.ProgId, oleFormat.ProgID,
                "PowerPoint converted to the wrong OLE class.");
            AssertNear(240f, shape.Width, 0.8f,
                "PowerPoint OLE conversion changed the formula width.");
            AssertNear(24f, shape.Height, 0.5f,
                "PowerPoint OLE conversion changed the formula height.");
            Console.WriteLine("[PowerPoint 8/10] Exercising the converted OLE show verb...");
            oleFormat.DoVerb(0);
            Release(oleFormat);
            oleFormat = null;

            Console.WriteLine("[PowerPoint 9/10] Exporting the final slide and checking visible content...");
            var presentationPath = Path.Combine(artifactRoot, "VisualTeX-PowerPoint-Flow.pptx");
            presentation.SaveAs(
                presentationPath,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                MsoTriState.msoFalse);
            var slidePng = Path.Combine(artifactRoot, "VisualTeX-PowerPoint-Flow.png");
            slide.Export(slidePng, "PNG", 960, 540);
            var preview = AnalyzeDarkPixels(slidePng);
            if (preview.Count < 40)
                throw new InvalidDataException(
                    $"PowerPoint export is blank or nearly blank ({preview.Count} dark pixels).");
            if (preview.Width > 140 || preview.Height > 45)
                throw new InvalidDataException(
                    $"PowerPoint OLE export still resembles the placeholder cache " +
                    $"({preview.Width}x{preview.Height} dark-pixel bounds).");
            Console.WriteLine(
                $"[PowerPoint 10/10] Saved {presentationPath}; preview has {preview.Count} dark pixels " +
                $"inside {preview.Width}x{preview.Height} bounds.");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            Release(oleFormat);
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

    private static void AppendToLastWordOmmlAndSelect(
        Word.Document document,
        string suffix)
    {
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? insertion = null;
        Word.Range? selectedRange = null;
        try
        {
            maths = document.OMaths;
            if (maths.Count == 0)
                throw new InvalidDataException("Word document contains no OMML equation to edit.");
            math = maths[maths.Count];
            math.Linearize();
            insertion = math.Range.Duplicate;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertion.InsertBefore(suffix);
            math.BuildUp();
            selectedRange = math.Range;
            selectedRange.Select();
        }
        finally
        {
            Release(selectedRange);
            Release(insertion);
            Release(math);
            Release(maths);
        }
    }

    private static void Commit(
        VisualTeXSessionClient client,
        OfficeSessionDocument session,
        string displayMode,
        string objectMode,
        string latex,
        float renderWidth = ExportWidth,
        float renderHeight = ExportHeight,
        float baseline = ExportBaseline,
        bool dirty = true,
        bool numbered = false,
        string? mathMl = null)
    {
        var lineId = session.Lines.First().Id;
        var svg = CreateSvg(renderWidth, renderHeight);
        var exportResult = new Dictionary<string, object?>
        {
            ["svg"] = svg,
            ["svgBase64"] = "data:image/svg+xml;base64," +
                Convert.ToBase64String(Encoding.UTF8.GetBytes(svg)),
            ["pngBase64"] = CreatePngDataUrl(latex, renderWidth, renderHeight),
            ["width"] = renderWidth,
            ["height"] = renderHeight,
            ["baseline"] = baseline,
        };
        if (!string.IsNullOrWhiteSpace(mathMl)) exportResult["mathMl"] = mathMl;
        var patch = new Dictionary<string, object>
        {
            ["lines"] = new[]
            {
                new Dictionary<string, object>
                {
                    ["id"] = lineId,
                    ["latex"] = latex,
                },
            },
            ["activeLineId"] = lineId,
            ["codeFormat"] = "latex",
            ["displayMode"] = displayMode,
            ["objectMode"] = objectMode,
            ["numbered"] = numbered,
            ["exportWidth"] = renderWidth,
            ["exportHeight"] = renderHeight,
            ["exportResult"] = exportResult,
            ["dirty"] = dirty,
            ["status"] = "committing",
        };
        client.PatchAsync(session.Id, patch, CancellationToken.None).GetAwaiter().GetResult();
    }

    private static void WaitForWordInlineShapeCount(
        Word.Document document,
        int expectedCount,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        var lastCount = -1;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            Word.InlineShapes? shapes = null;
            try
            {
                shapes = document.InlineShapes;
                lastCount = shapes.Count;
                if (lastCount == expectedCount) return;
            }
            finally { Release(shapes); }
        }
        throw new TimeoutException(
            $"Expected {expectedCount} Word inline shapes, last count was {lastCount}.");
    }

    private static string WaitForWordNativeReferenceResult(
        Word.Document document,
        string expectedResult,
        string? expectedCode,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        var lastResult = string.Empty;
        var lastCode = string.Empty;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            Word.Fields? fields = null;
            try
            {
                fields = document.Fields;
                for (var index = 1; index <= fields.Count; index++)
                {
                    Word.Field? field = null;
                    Word.Range? code = null;
                    Word.Range? result = null;
                    try
                    {
                        field = fields[index];
                        if (field.Type != Word.WdFieldType.wdFieldRef) continue;
                        code = field.Code;
                        var codeText = (code.Text ?? string.Empty).Trim();
                        if (!string.IsNullOrWhiteSpace(expectedCode)
                            && !string.Equals(
                                NormalizeFieldCode(codeText),
                                NormalizeFieldCode(expectedCode),
                                StringComparison.OrdinalIgnoreCase))
                            continue;
                        field.Update();
                        result = field.Result;
                        lastCode = codeText;
                        lastResult = (result.Text ?? string.Empty).Trim();
                        if (string.Equals(lastResult, expectedResult, StringComparison.Ordinal))
                            return codeText;
                    }
                    finally
                    {
                        Release(result);
                        Release(code);
                        Release(field);
                    }
                }
            }
            finally { Release(fields); }
        }
        throw new TimeoutException(
            $"Word native REF field did not become {expectedResult}; " +
            $"last result was [{lastResult}], last code was [{lastCode}].");
    }

    private static string NormalizeFieldCode(string value) =>
        string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private static bool DocumentTextContains(Word.Document document, string value)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            return (content.Text ?? string.Empty).IndexOf(value, StringComparison.Ordinal) >= 0;
        }
        finally { Release(content); }
    }

    private static bool WordBookmarkExists(Word.Document document, string name)
    {
        Word.Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            return bookmarks.Exists(name);
        }
        finally { Release(bookmarks); }
    }

    private static Word.InlineShape WaitForWordShapeMode(
        Word.Document document,
        bool nativeOle,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            Word.InlineShapes? shapes = null;
            Word.InlineShape? candidate = null;
            Word.OLEFormat? format = null;
            try
            {
                shapes = document.InlineShapes;
                if (shapes.Count != 1) continue;
                candidate = shapes[1];
                var isNativeOle = false;
                if (candidate.Type is Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                    or Word.WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                {
                    try
                    {
                        format = candidate.OLEFormat;
                        isNativeOle = string.Equals(
                            format.ProgID,
                            FormulaOleContract.ProgId,
                            StringComparison.OrdinalIgnoreCase);
                    }
                    catch
                    {
                        isNativeOle = false;
                    }
                }
                if (isNativeOle == nativeOle)
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }
            }
            finally
            {
                Release(format);
                Release(candidate);
                Release(shapes);
            }
        }
        throw new TimeoutException(nativeOle
            ? "Word formula did not become a VisualTeX native OLE object."
            : "Word formula did not become a cross-platform picture.");
    }

    private static OfficeSessionDocument WaitForUnchangedEditorReady(
        VisualTeXSessionClient client,
        string sessionId,
        TimeSpan timeout)
    {
        var startedAt = DateTime.UtcNow;
        var deadline = startedAt + timeout;
        OfficeSessionDocument? session = null;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            session = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            if (session.Dirty)
                throw new InvalidOperationException(
                    $"Session {sessionId} became dirty before the user changed the formula.");
            if (session.Status is "completed" or "failed" or "cancelled")
                throw new InvalidOperationException(
                    $"Session {sessionId} reached {session.Status} before the unchanged editor was closed: {session.Error}");
            // A pristine editor is allowed to remain in `created`; no autosave is
            // necessary until the user changes something. Wait long enough for
            // the WebView and its close handlers to mount before closing it.
            if (DateTime.UtcNow - startedAt >= TimeSpan.FromSeconds(3)
                && session.Status is "created" or "editing")
                return session;
        }
        throw new TimeoutException(
            $"Session {sessionId} was not ready for unchanged close; " +
            $"last status was {session?.Status ?? "unknown"}.");
    }

    private static OfficeSessionDocument WaitForTerminal(
        VisualTeXSessionClient client,
        string sessionId,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        OfficeSessionDocument? session = null;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            session = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            if (session.Status is "completed" or "failed" or "cancelled") return session;
        }
        throw new TimeoutException(
            $"Session {sessionId} did not finish; last status was {session?.Status ?? "unknown"}.");
    }

    private static void WaitForAddInIdle(object addIn, TimeSpan timeout)
    {
        var field = addIn.GetType().GetField(
            "_operationGate",
            System.Reflection.BindingFlags.Instance |
            System.Reflection.BindingFlags.NonPublic)
            ?? throw new MissingFieldException("Office add-in operation gate is missing.");
        var gate = field.GetValue(addIn) as SemaphoreSlim
            ?? throw new InvalidOperationException("Office add-in operation gate is unavailable.");
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            if (gate.CurrentCount == 1) return;
            Thread.Sleep(25);
        }
        throw new TimeoutException("Office add-in did not return to the idle state.");
    }

    private static HashSet<string> SnapshotSessionIds()
    {
        Directory.CreateDirectory(SessionRoot);
        return Directory.EnumerateDirectories(SessionRoot)
            .Select(Path.GetFileName)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Cast<string>()
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    private static string WaitForNewSession(
        HashSet<string> existing,
        string expectedHost,
        TimeSpan timeout,
        Func<string?>? errorProvider = null)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            var diagnosticError = errorProvider?.Invoke();
            if (!string.IsNullOrWhiteSpace(diagnosticError))
                throw new InvalidOperationException(diagnosticError);
            Thread.Sleep(100);
            foreach (var directory in Directory.EnumerateDirectories(SessionRoot)
                         .OrderByDescending(Directory.GetLastWriteTimeUtc))
            {
                var id = Path.GetFileName(directory);
                if (string.IsNullOrWhiteSpace(id) || existing.Contains(id)) continue;
                var sessionPath = Path.Combine(directory, "session.json");
                if (!File.Exists(sessionPath)) continue;
                var json = File.ReadAllText(sessionPath);
                if (json.IndexOf($"\"host\": \"{expectedHost}\"", StringComparison.OrdinalIgnoreCase) >= 0
                    || json.IndexOf($"\"host\":\"{expectedHost}\"", StringComparison.OrdinalIgnoreCase) >= 0)
                    return id;
            }
        }
        throw new TimeoutException($"No new {expectedHost} Office Session appeared.");
    }

    private static string CreateSvg(float width, float height) =>
        $"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 {width:F0} {height:F0}\">" +
        "<path fill=\"#111111\" d=\"" +
        "M4 5 L10 5 L18 14 L26 5 L32 5 L21 17 L33 29 L27 29 L18 20 L9 29 L3 29 L15 17 Z " +
        "M48 14 L60 14 L60 2 L66 2 L66 14 L78 14 L78 20 L66 20 L66 32 L60 32 L60 20 L48 20 Z " +
        "M94 5 L100 5 L108 14 L116 5 L122 5 L111 17 L123 29 L117 29 L108 20 L99 29 L93 29 L105 17 Z" +
        "\"/></svg>";

    private static string CreatePngDataUrl(string latex, float width, float height)
    {
        var pixelWidth = Math.Max(32, (int)Math.Ceiling(width * 2));
        var pixelHeight = Math.Max(24, (int)Math.Ceiling(height * 2));
        using var bitmap = new Bitmap(pixelWidth, pixelHeight, PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.Transparent);
        using var font = new Font("Cambria Math", 28f, FontStyle.Regular, GraphicsUnit.Pixel);
        using var brush = new SolidBrush(Color.Black);
        graphics.DrawString(latex, font, brush, new PointF(2, 4));
        using var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        return "data:image/png;base64," + Convert.ToBase64String(stream.ToArray());
    }

    private static void ReportPowerPointMetadata(string stage, PowerPoint.Shape shape)
    {
        var alternativeText = shape.AlternativeText ?? string.Empty;
        var decodedAlternative = FormulaMetadataCodec.Decode(alternativeText);
        string tagValue = string.Empty;
        PowerPoint.Tags? tags = null;
        try
        {
            tags = shape.Tags;
            try { tagValue = tags["VisualTeXMetadata"] ?? string.Empty; } catch { }
        }
        finally { Release(tags); }
        Console.WriteLine(
            $"  {stage} metadata: alt={alternativeText.Length} " +
            $"altDecoded={decodedAlternative is not null} tag={tagValue.Length} " +
            $"tagDecoded={FormulaMetadataCodec.Decode(tagValue) is not null}");
    }

    private static (int Count, int Width, int Height) AnalyzeDarkPixels(string path)
    {
        using var bitmap = new Bitmap(path);
        var count = 0;
        var minimumX = bitmap.Width;
        var minimumY = bitmap.Height;
        var maximumX = -1;
        var maximumY = -1;
        for (var y = 0; y < bitmap.Height; y++)
        for (var x = 0; x < bitmap.Width; x++)
        {
            var pixel = bitmap.GetPixel(x, y);
            if (pixel.R >= 120 && pixel.G >= 120 && pixel.B >= 120) continue;
            count++;
            minimumX = Math.Min(minimumX, x);
            minimumY = Math.Min(minimumY, y);
            maximumX = Math.Max(maximumX, x);
            maximumY = Math.Max(maximumY, y);
        }
        return maximumX < minimumX || maximumY < minimumY
            ? (0, 0, 0)
            : (count, maximumX - minimumX + 1, maximumY - minimumY + 1);
    }

    private static void AssertNear(float expected, float actual, float tolerance, string message)
    {
        if (Math.Abs(expected - actual) > tolerance)
            throw new InvalidOperationException(
                $"{message} Expected {expected:F3}, actual {actual:F3}.");
    }

    private static void AssertEqual<T>(T expected, T actual, string message)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new InvalidOperationException(
                $"{message} Expected {expected}, actual {actual}.");
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            try { Marshal.ReleaseComObject(value); } catch { }
        }
    }

    private static void ForceComCleanup()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
        Thread.Sleep(500);
    }

    private sealed class TeeTextWriter : TextWriter
    {
        private readonly TextWriter _primary;
        private readonly TextWriter _secondary;

        public TeeTextWriter(TextWriter primary, TextWriter secondary)
        {
            _primary = primary;
            _secondary = secondary;
        }

        public override Encoding Encoding => _primary.Encoding;

        public override void Write(char value)
        {
            _primary.Write(value);
            _secondary.Write(value);
        }

        public override void Write(string? value)
        {
            _primary.Write(value);
            _secondary.Write(value);
        }

        public override void WriteLine(string? value)
        {
            _primary.WriteLine(value);
            _secondary.WriteLine(value);
        }

        public override void Flush()
        {
            _primary.Flush();
            _secondary.Flush();
        }
    }
}
