using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using WinForms = System.Windows.Forms;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeOleRealDoubleClickAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var fixture = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", "..", "..",
            "artifacts", "mathtype-native-editor",
            "VisualTeX-MathType7-NativeEditor-5f04f8b3545e444a824705446e314ba1.docx"));
        if (!File.Exists(fixture))
            throw new FileNotFoundException(
                "A genuine MathType-generated Equation.DSMT4 fixture is required.", fixture);
        var path = Path.Combine(
            artifactRoot,
            $"VisualTeX-MathType7-RealDoubleClick-{Guid.NewGuid():N}.docx");
        File.Copy(fixture, path, overwrite: false);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Window? window = null;
        Word.InlineShape? shape = null;
        Word.Range? range = null;
        Office.COMAddIns? installedAddIns = null;
        Office.COMAddIn? installedAddIn = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        var consoleWindow = GetConsoleWindow();
        var hookTracePath = Path.Combine(artifactRoot, "mathtype-ole-hook-trace.log");
        Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", hookTracePath);
        var expectNativeMathType = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_MATHTYPE_DOUBLE_CLICK_EXPECT_NATIVE"),
            "1",
            StringComparison.Ordinal);
        var isolatedPreferencesPath = Path.Combine(
            artifactRoot,
            expectNativeMathType
                ? "office-preferences-mathtype-native.json"
                : "office-preferences-mathtype-visualtex.json");
        File.WriteAllText(
            isolatedPreferencesPath,
            expectNativeMathType
                ? "{\"powerpointDefaultFontSizePt\":20.0,\"mathtypeDoubleClickEditEnabled\":false}"
                : "{\"powerpointDefaultFontSizePt\":20.0,\"mathtypeDoubleClickEditEnabled\":true}");
        Environment.SetEnvironmentVariable(
            "VISUALTEX_OFFICE_PREFERENCES_PATH",
            isolatedPreferencesPath);

        try
        {
            Console.WriteLine("[MathType real double-click 1/7] Starting visible isolated Word with the current source add-in...");
            application = CreateWordApplication(visible: true);
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

            document = application.Documents.Open(path, ReadOnly: false, Visible: true);
            document.Activate();
            AssertEqual(1, document.InlineShapes.Count,
                "MathType double-click fixture must contain exactly one inline OLE equation.");
            shape = document.InlineShapes[1];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                "MathType double-click fixture is not a registered MathType OLE equation.");
            AssertEqual(0, GetMathTypeTopLevelWindows().Count,
                "A visible MathType equation window was already open before the VisualTeX double-click acceptance.");

            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            Console.WriteLine("[MathType real double-click 2/7] Selecting the equation once to cache only its ProgID/rectangle, without activating MathType...");
            range = shape.Range;
            range.Select();
            window = application.ActiveWindow;
            window.Activate();
            WinForms.Application.DoEvents();
            Thread.Sleep(500);
            AssertEqual(0, GetMathTypeTopLevelWindows().Count,
                "A single click on MathType OLE incorrectly opened the MathType editor.");
            window.GetPoint(out var left, out var top, out var width, out var height, range);
            if (width <= 0 || height <= 0)
                throw new InvalidDataException("Word did not return a visible MathType OLE rectangle.");

            var wordWindowHandle = new IntPtr(window.Hwnd);
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
            Thread.Sleep(650);
            range.Select();
            window.GetPoint(out left, out top, out width, out height, range);
            if (width <= 0 || height <= 0)
                throw new InvalidDataException("Word did not return a current MathType OLE rectangle before double-click.");
            if (consoleWindow != IntPtr.Zero) ShowWindow(consoleWindow, 0);

            Console.WriteLine("[MathType real double-click 3/7] Sending a real mouse double-click on Equation.DSMT4...");
            var sessionsBefore = SnapshotSessionIds();
            var centerX = left + width / 2;
            var centerY = top + height / 2;
            SetCursorPos(centerX, centerY);
            Thread.Sleep(120);
            for (var click = 0; click < 2; click++)
            {
                mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                if (click == 0 || !expectNativeMathType)
                    Thread.Sleep(90);
            }

            if (expectNativeMathType)
            {
                // The low-level hook has synchronously frozen the OLE range before
                // mouse_event returns from the second button-down. Deliberately move
                // Word Selection away before the 40ms callback / 300ms native-open
                // turn. The old implementation rediscovered MathType from Selection
                // and therefore silently skipped the open in this state; the current
                // implementation must resolve the original frozen OLE identity.
                Word.Range? driftRange = null;
                try
                {
                    var driftPosition = Math.Max(
                        document.Content.Start,
                        document.Content.End - 1);
                    if (driftPosition >= range.Start && driftPosition <= range.End)
                        driftPosition = document.Content.Start;
                    driftRange = document.Range(driftPosition, driftPosition);
                    driftRange.Select();
                }
                finally { Release(driftRange); }
                WinForms.Application.DoEvents();

                var deadline = DateTime.UtcNow.AddSeconds(12);
                while (DateTime.UtcNow < deadline
                    && GetMathTypeTopLevelWindows().Count == 0)
                {
                    WinForms.Application.DoEvents();
                    Thread.Sleep(100);
                }
                AssertTrue(
                    GetMathTypeTopLevelWindows().Count > 0,
                    "Disabling MathType double-click editing did not release the Word OLE double-click to native MathType.");
                Thread.Sleep(500);
                var sessionsAfterNativeDoubleClick = SnapshotSessionIds();
                AssertEqual(
                    0,
                    sessionsAfterNativeDoubleClick.Except(sessionsBefore).Count(),
                    "Disabling MathType double-click editing still created a VisualTeX edit Session.");
                var nativeHookTrace = File.ReadAllText(hookTracePath);
                AssertTrue(
                    nativeHookTrace.IndexOf(
                        "nativeOleTarget=True suppressSecondDown=True dispatchCallback=True",
                        StringComparison.Ordinal) >= 0,
                    "Native MathType mode did not take exclusive ownership of the second click.");
                AssertTrue(
                    nativeHookTrace.IndexOf(
                        "second-button-down-suppressed",
                        StringComparison.Ordinal) >= 0,
                    "Native MathType mode allowed Word to begin a competing OLE activation.");
                AssertTrue(
                    nativeHookTrace.IndexOf(
                        "callback-begin interceptedNativeOle=True",
                        StringComparison.Ordinal) >= 0,
                    "Native MathType mode never dispatched its one native-open callback.");
                AssertTrue(
                    nativeHookTrace.IndexOf(
                        "addin-native-mathtype-open-scheduled delayMs=300",
                        StringComparison.Ordinal) >= 0,
                    "Native MathType mode did not defer OLE activation until Word's double-click message had unwound.");
                AssertTrue(
                    nativeHookTrace.IndexOf(
                        "addin-native-mathtype-open started=True",
                        StringComparison.Ordinal) >= 0,
                    "The delayed native MathType Open verb was not invoked exactly through the intended path.");

                // Close the untouched OLE editor and prove Word accepts COM/input again.
                // This specifically guards the user-reported state where the MathType
                // window never appeared and Word remained stuck alternating the busy
                // cursor indefinitely.
                foreach (var nativeWindow in GetMathTypeTopLevelWindows())
                {
                    SetForegroundWindow(nativeWindow);
                    Thread.Sleep(120);
                    WinForms.SendKeys.SendWait("^{F4}");
                }
                var nativeCloseDeadline = DateTime.UtcNow.AddSeconds(8);
                while (DateTime.UtcNow < nativeCloseDeadline
                    && GetMathTypeTopLevelWindows().Count > 0)
                {
                    WinForms.Application.DoEvents();
                    Thread.Sleep(100);
                }
                AssertEqual(0, GetMathTypeTopLevelWindows().Count,
                    "Native MathType editor did not close after the untouched double-click test.");
                SetForegroundWindow(wordWindowHandle);
                Thread.Sleep(300);
                window.Activate();
                range.Select();
                application.ScreenRefresh();
                WinForms.Application.DoEvents();
                AssertEqual(1, document.InlineShapes.Count,
                    "Word was not responsive after closing the native MathType editor.");

                // Guard the user-reported stale-cache state after creating another
                // formula. The low-level hook object can exist while a particular
                // double-click was not claimed (for example because its cached
                // rectangle still belongs to the newly inserted equation). In that
                // case WordBeforeDoubleClick must be released to native MathType.
                // The previous implementation cancelled solely because the hook
                // object existed, creating a black hole with no editor.
                Thread.Sleep(1100);
                range.Select();
                WinForms.Application.DoEvents();
                var beforeDoubleClickMethod = typeof(VisualTeX.WordVsto.ThisAddIn)
                    .GetMethod(
                        "OnWindowBeforeDoubleClick",
                        System.Reflection.BindingFlags.Instance
                        | System.Reflection.BindingFlags.NonPublic)
                    ?? throw new MissingMethodException(
                        "Could not resolve WordBeforeDoubleClick routing for stale-cache acceptance.");
                var beforeDoubleClickArgs = new object[] { application.Selection, false };
                beforeDoubleClickMethod.Invoke(addIn, beforeDoubleClickArgs);
                AssertEqual(false, (bool)beforeDoubleClickArgs[1],
                    "An unclaimed MathType double-click was cancelled merely because the low-level hook object existed.");

                Console.WriteLine(
                    "MathType native-double-click acceptance passed: VisualTeX suppressed only a hook-claimed competing second click, deferred one native Open verb to a later Office UI turn, released unclaimed Word double-clicks, MathType opened, and Word was responsive after the editor closed.");
                return;
            }

            var editSessionId = WaitForNewSession(
                sessionsBefore,
                "word",
                TimeSpan.FromSeconds(30));
            var editSession = WaitForUnchangedEditorReady(
                client,
                editSessionId,
                TimeSpan.FromSeconds(15));
            Thread.Sleep(750);
            var sessionsAfterVisualTeXDoubleClick = SnapshotSessionIds();
            AssertEqual(
                1,
                sessionsAfterVisualTeXDoubleClick.Except(sessionsBefore).Count(),
                "One real MathType double-click created duplicate VisualTeX edit Sessions.");
            AssertEqual("edit", editSession.Mode,
                "Real MathType OLE double-click did not create an edit Session.");
            AssertEqual(FormulaOleContract.MathTypeOleMode, editSession.ObjectMode,
                "Real MathType OLE double-click opened the wrong object mode.");
            var sourceLatex = (editSession.Lines.FirstOrDefault()?.Latex ?? string.Empty)
                .Replace(" ", string.Empty);
            AssertTrue(
                sourceLatex.IndexOf("sqrt", StringComparison.OrdinalIgnoreCase) >= 0
                && (sourceLatex.IndexOf("p^2", StringComparison.Ordinal) >= 0
                    || sourceLatex.IndexOf("p^{2}", StringComparison.Ordinal) >= 0)
                && (sourceLatex.IndexOf("q^2", StringComparison.Ordinal) >= 0
                    || sourceLatex.IndexOf("q^{2}", StringComparison.Ordinal) >= 0),
                $"Real MathType OLE double-click opened the wrong source: '{sourceLatex}'.");
            AssertEqual(0, GetMathTypeTopLevelWindows().Count,
                "The intercepted MathType double-click also opened the native MathType editor.");
            Console.WriteLine(
                $"  VisualTeX Session={editSessionId}; source={sourceLatex}; native MathType window count=0.");

            Console.WriteLine("[MathType real double-click 4/7] Committing through VisualTeX while preserving the same MathType OLE object...");
            const string editedLatex = @"\sqrt{r^2+s^2}";
            const string editedMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<msqrt><mrow><msup><mi>r</mi><mn>2</mn></msup><mo>+</mo>"
                + "<msup><mi>s</mi><mn>2</mn></msup></mrow></msqrt></math>";
            Commit(
                client,
                editSession,
                editSession.DisplayMode ?? "inline",
                FormulaOleContract.MathTypeOleMode,
                editedLatex,
                numbered: false,
                mathMl: editedMathMl);
            var terminal = WaitForTerminal(
                client,
                editSessionId,
                TimeSpan.FromSeconds(60));
            AssertEqual("completed", terminal.Status,
                terminal.Error ?? "MathType OLE VisualTeX edit did not complete.");
            client.CloseEditorAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            AssertEqual(1, document.InlineShapes.Count,
                "Preserving MathType OLE duplicated or removed the source equation.");
            Release(shape);
            shape = document.InlineShapes[1];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                "VisualTeX preserve-mode commit changed the equation away from MathType OLE.");
            AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                "VisualTeX preserve-mode commit changed the MathType ProgID.");
            AssertEqual(0, GetMathTypeTopLevelWindows().Count,
                "MathType editor remained visibly open after VisualTeX finished the OLE commit.");

            Console.WriteLine("[MathType real double-click 5/7] Saving/reopening Word and resolving the same MathType OLE rectangle...");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: true);
            document.Activate();
            AssertEqual(1, document.InlineShapes.Count,
                "Save/reopen changed the MathType OLE inventory.");
            Release(shape);
            shape = document.InlineShapes[1];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                "Save/reopen changed the equation away from MathType OLE.");
            Release(range);
            range = shape.Range;
            range.Select();
            Release(window);
            window = application.ActiveWindow;
            window.Activate();
            WinForms.Application.DoEvents();
            Thread.Sleep(500);
            window.GetPoint(out left, out top, out width, out height, range);
            if (width <= 0 || height <= 0)
                throw new InvalidDataException("Word did not return the reopened MathType OLE rectangle.");
            wordWindowHandle = new IntPtr(window.Hwnd);
            SetWindowPos(wordWindowHandle, new IntPtr(-1), 0, 0, 0, 0, noMoveNoSizeShow);
            SetForegroundWindow(wordWindowHandle);
            if (GetWindowRect(wordWindowHandle, out var reopenedWordWindowRectangle))
            {
                SetCursorPos(
                    reopenedWordWindowRectangle.Left
                        + Math.Max(40, (reopenedWordWindowRectangle.Right - reopenedWordWindowRectangle.Left) / 2),
                    reopenedWordWindowRectangle.Top + 18);
                mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
            }
            WinForms.Application.DoEvents();
            Thread.Sleep(650);
            range.Select();
            window.Activate();
            SetForegroundWindow(wordWindowHandle);
            WinForms.Application.DoEvents();
            Thread.Sleep(250);
            window.GetPoint(out left, out top, out width, out height, range);
            if (width <= 0 || height <= 0)
                throw new InvalidDataException("Word did not return a current reopened MathType OLE rectangle before double-click.");

            Console.WriteLine("[MathType real double-click 6/7] Real-double-clicking the edited MathType OLE a second time...");
            // Saving/reopening can temporarily move foreground ownership away from
            // Word even after SetForegroundWindow succeeds. Prime the exact OLE
            // center with one real click, wait past the system double-click window,
            // then re-resolve both Selection and screen rectangle before sending
            // the actual double-click. This mirrors the proven installed re-edit
            // acceptance and prevents a false negative where mouse_event clicks the
            // stale cursor position instead of the reopened equation.
            centerX = left + width / 2;
            centerY = top + height / 2;
            if (!SetCursorPos(centerX, centerY))
                throw new InvalidOperationException(
                    $"Windows refused to move the cursor to the reopened MathType OLE center {centerX},{centerY}.");
            mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
            Thread.Sleep(900);
            range.Select();
            window.Activate();
            SetForegroundWindow(wordWindowHandle);
            WinForms.Application.DoEvents();
            Thread.Sleep(250);
            window.GetPoint(out left, out top, out width, out height, range);
            if (width <= 0 || height <= 0)
                throw new InvalidDataException(
                    "Word did not return a current reopened MathType OLE rectangle after foreground priming.");

            var reopenSessionsBefore = SnapshotSessionIds();
            centerX = left + width / 2;
            centerY = top + height / 2;
            if (!SetCursorPos(centerX, centerY))
                throw new InvalidOperationException(
                    $"Windows refused the final cursor move to the reopened MathType OLE center {centerX},{centerY}.");
            Thread.Sleep(120);
            for (var click = 0; click < 2; click++)
            {
                mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(90);
            }
            var reopenSessionId = WaitForNewSession(
                reopenSessionsBefore,
                "word",
                TimeSpan.FromSeconds(30));
            var reopenSession = WaitForUnchangedEditorReady(
                client,
                reopenSessionId,
                TimeSpan.FromSeconds(15));
            AssertEqual(FormulaOleContract.MathTypeOleMode, reopenSession.ObjectMode,
                "Second real MathType double-click changed the object mode.");
            var reopenedLatex = (reopenSession.Lines.FirstOrDefault()?.Latex ?? string.Empty)
                .Replace(" ", string.Empty);
            AssertTrue(
                reopenedLatex.IndexOf("sqrt", StringComparison.OrdinalIgnoreCase) >= 0
                && (reopenedLatex.IndexOf("r^2", StringComparison.Ordinal) >= 0
                    || reopenedLatex.IndexOf("r^{2}", StringComparison.Ordinal) >= 0)
                && (reopenedLatex.IndexOf("s^2", StringComparison.Ordinal) >= 0
                    || reopenedLatex.IndexOf("s^{2}", StringComparison.Ordinal) >= 0),
                $"Second real MathType double-click recovered the wrong source: '{reopenedLatex}'.");
            AssertEqual(0, GetMathTypeTopLevelWindows().Count,
                "Second intercepted MathType double-click also opened native MathType.");
            Console.WriteLine("[MathType real double-click 7/9] Switching this same double-click Session to VisualTeX OLE...");
            const string convertedLatex = @"\frac{r}{s}+1";
            Commit(
                client,
                reopenSession,
                reopenSession.DisplayMode ?? "inline",
                FormulaOleContract.NativeOleMode,
                convertedLatex,
                numbered: false);
            var convertedTerminal = WaitForTerminal(
                client,
                reopenSessionId,
                TimeSpan.FromSeconds(60));
            AssertEqual("completed", convertedTerminal.Status,
                convertedTerminal.Error ?? "MathType OLE to VisualTeX OLE Session conversion did not complete.");
            client.CloseEditorAsync(reopenSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            AssertEqual(1, document.InlineShapes.Count,
                "MathType-to-VisualTeX target-format edit duplicated or removed the equation.");
            Release(shape);
            shape = document.InlineShapes[1];
            AssertTrue(
                VisualTeX.WordVsto.WordFormulaMetadataReader.IsNativeOle(shape),
                "The target-format edit did not create VisualTeX native OLE.");
            AssertEqual(FormulaOleContract.ProgId, shape.OLEFormat.ProgID,
                "The target-format edit created the wrong VisualTeX OLE ProgID.");
            var convertedMetadata =
                VisualTeX.WordVsto.WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException("Converted VisualTeX OLE has no readable metadata.");
            AssertEqual(convertedLatex, convertedMetadata.Latex,
                "Converted VisualTeX OLE persisted the wrong source.");
            AssertEqual(0, GetMathTypeTopLevelWindows().Count,
                "Converting MathType OLE to VisualTeX OLE unexpectedly left MathType visible.");

            Console.WriteLine("[MathType real double-click 8/9] Saving/reopening the converted VisualTeX OLE...");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: true);
            document.Activate();
            AssertEqual(1, document.InlineShapes.Count,
                "Save/reopen changed the converted VisualTeX OLE inventory.");
            Release(shape);
            shape = document.InlineShapes[1];
            AssertTrue(
                VisualTeX.WordVsto.WordFormulaMetadataReader.IsNativeOle(shape),
                "Save/reopen changed the converted equation away from VisualTeX OLE.");
            Release(range);
            range = shape.Range;
            range.Select();
            Release(window);
            window = application.ActiveWindow;
            window.Activate();
            WinForms.Application.DoEvents();
            Thread.Sleep(500);
            window.GetPoint(out left, out top, out width, out height, range);
            wordWindowHandle = new IntPtr(window.Hwnd);
            SetWindowPos(wordWindowHandle, new IntPtr(-1), 0, 0, 0, 0, noMoveNoSizeShow);
            SetForegroundWindow(wordWindowHandle);
            WinForms.Application.DoEvents();
            Thread.Sleep(500);

            Console.WriteLine("[MathType real double-click 9/9] Real-double-clicking the converted VisualTeX OLE...");
            var convertedSessionsBefore = SnapshotSessionIds();
            SetCursorPos(left + width / 2, top + height / 2);
            Thread.Sleep(120);
            for (var click = 0; click < 2; click++)
            {
                mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(90);
            }
            var convertedSessionId = WaitForNewSession(
                convertedSessionsBefore,
                "word",
                TimeSpan.FromSeconds(30));
            var convertedSession = WaitForUnchangedEditorReady(
                client,
                convertedSessionId,
                TimeSpan.FromSeconds(15));
            AssertEqual(FormulaOleContract.NativeOleMode, convertedSession.ObjectMode,
                "Converted VisualTeX OLE reopened with the wrong object mode.");
            AssertEqual(convertedLatex,
                convertedSession.Lines.FirstOrDefault()?.Latex ?? string.Empty,
                "Converted VisualTeX OLE real double-click reopened the wrong source.");
            client.CloseEditorAsync(convertedSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            var convertedClosed = WaitForTerminal(
                client,
                convertedSessionId,
                TimeSpan.FromSeconds(30));
            AssertEqual("completed", convertedClosed.Status,
                convertedClosed.Error ?? "Converted VisualTeX OLE Session did not close cleanly.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            Console.WriteLine(
                "MathType real-double-click acceptance passed: MathType OLE opened only VisualTeX, preserved Equation.DSMT4 through edit/save/reopen, then the same double-click workflow converted it to VisualTeX.Formula.1 and reopened it in VisualTeX.");
        }
        finally
        {
            if (consoleWindow != IntPtr.Zero) ShowWindow(consoleWindow, 5);
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", null);
            if (expectNativeMathType)
                Environment.SetEnvironmentVariable("VISUALTEX_OFFICE_PREFERENCES_PATH", null);
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
            Release(range);
            Release(shape);
            Release(window);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }
}
