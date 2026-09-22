using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;
using WinForms = System.Windows.Forms;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const uint WmClose = 0x0010;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool PostMessage(
        IntPtr windowHandle,
        uint message,
        UIntPtr wParam,
        IntPtr lParam);

    private static void RunWordEditorNativeClose(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        AssertDirtyOmmlClosePolicy(client);

        Word.Application? application = null;
        Word.Document? document = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
        Array custom = Array.Empty<object>();
        string? firstSessionId = null;
        string? secondSessionId = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var lineId = Guid.NewGuid().ToString("D");
            var insertSession = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Host = "word",
                Mode = "create",
                FormulaId = formulaId,
                Title = "Native close acceptance",
                DisplayMode = "inline",
                ObjectMode = FormulaOleContract.WordOmmlMode,
                CodeFormat = "latex",
                FontSizePt = 11,
                Lines = new List<FormulaLine>
                {
                    new() { Id = lineId, Latex = "x+y" },
                },
            };
            const string mathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>x</mi><mo>+</mo><mi>y</mi></math>";
            service.InsertOmml(insertSession, mathMl);

            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);

            SelectOmmlFormula(document, formulaId, ref bookmark, ref equationRange);
            var sessionsBeforeFirst = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            firstSessionId = WaitForNewSession(
                sessionsBeforeFirst,
                "word",
                TimeSpan.FromSeconds(30));
            var editorWindow = WaitForVisibleOfficeEditorWindow(TimeSpan.FromSeconds(20));
            if (!PostMessage(editorWindow, WmClose, UIntPtr.Zero, IntPtr.Zero))
                throw new InvalidOperationException(
                    $"Unable to post WM_CLOSE to the VisualTeX editor (Win32 {Marshal.GetLastWin32Error()}).");

            var firstTerminal = WaitForTerminal(
                client,
                firstSessionId,
                TimeSpan.FromSeconds(45));
            AssertEqual(
                "completed",
                firstTerminal.Status,
                firstTerminal.Error
                ?? "Closing an unchanged OMML editor did not complete the Session.");
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(20));
            WaitForOfficeEditorHidden(TimeSpan.FromSeconds(15));

            Release(equationRange);
            equationRange = null;
            Release(bookmark);
            bookmark = null;
            SelectOmmlFormula(document, formulaId, ref bookmark, ref equationRange);
            var sessionsBeforeSecond = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            secondSessionId = WaitForNewSession(
                sessionsBeforeSecond,
                "word",
                TimeSpan.FromSeconds(30));
            AssertTrue(
                !string.Equals(firstSessionId, secondSessionId, StringComparison.Ordinal),
                "The second edit reused the closed Session instead of opening a new one.");
            _ = WaitForVisibleOfficeEditorWindow(TimeSpan.FromSeconds(20));

            // The native-close watcher for the first Session wakes after five
            // seconds. Keep the second editor alive past that boundary so a
            // stale cleanup can never hide it or clear its Session id.
            Thread.Sleep(TimeSpan.FromSeconds(6));
            AssertTrue(
                FindVisibleOfficeEditorWindow() != IntPtr.Zero,
                "A stale native-close watcher hid the newer Office editor Session.");
            var secondStillReadable = client.GetSessionAsync(
                    secondSessionId,
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual(
                secondSessionId,
                secondStillReadable.Id,
                "The newer Office editor Session disappeared after the previous close watcher fired.");

            client.PatchAsync(
                    secondSessionId,
                    new
                    {
                        status = "cancelled",
                        explicitCancel = true,
                        error = (string?)null,
                    },
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            client.CloseEditorAsync(secondSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(20));
            secondSessionId = null;

            AssertStaleRibbonSessionRecovery(
                client,
                addIn,
                document,
                formulaId);

            Console.WriteLine(
                "Word native editor close acceptance passed: WM_CLOSE finalized the first Session, "
                + "released the Word operation gate, the reused editor kept the second Session alive, "
                + "and a Ribbon click recovered a stale/missing active Session without restarting Word.");
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(secondSessionId))
            {
                try
                {
                    client.PatchAsync(
                            secondSessionId!,
                            new
                            {
                                status = "cancelled",
                                explicitCancel = true,
                                error = (string?)null,
                            },
                            CancellationToken.None)
                        .GetAwaiter().GetResult();
                    client.CloseEditorAsync(secondSessionId!, CancellationToken.None)
                        .GetAwaiter().GetResult();
                }
                catch { }
            }
            if (addIn is not null)
            {
                try
                {
                    addIn.OnDisconnection(
                        ext_DisconnectMode.ext_dm_UserClosed,
                        ref custom);
                }
                catch { }
            }
            Release(equationRange);
            Release(bookmark);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordComplexOmmlCommitClose(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        AssertTrue(
            !AttachActiveWord,
            "The complex OMML editor Apply acceptance must own its Word process.");
        Directory.CreateDirectory(artifactRoot);
        AssertDirectOfficeEditorCloseLatency(client);
        var sessionPath =
            Environment.GetEnvironmentVariable(
                "VISUALTEX_COMPLEX_OMML_SESSION_PATH");
        if (string.IsNullOrWhiteSpace(sessionPath))
            throw new InvalidOperationException(
                "VISUALTEX_COMPLEX_OMML_SESSION_PATH must point to the exact user-reported editor session.");
        sessionPath = Path.GetFullPath(sessionPath);
        if (!File.Exists(sessionPath))
            throw new FileNotFoundException(
                "The complex OMML source session is missing.",
                sessionPath);

        var expectedMathMl =
            ReadSessionExportMathMl(sessionPath);
        var expectedLatex =
            ReadSessionLatex(sessionPath);
        if (string.IsNullOrWhiteSpace(expectedLatex))
            throw new InvalidDataException(
                "The complex OMML source session has no LaTeX line.");

        Word.Application? application = null;
        Word.Document? document = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Word.OMath? math = null;
        Word.Range? range = null;
        Array custom = Array.Empty<object>();
        string? sessionId = null;
        string? reloadSessionId = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Range(0, 0).Select();

            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);

            var existing = SnapshotSessionIds();
            addIn.OnInsertDisplayOmml(new object());
            sessionId = WaitForNewSession(
                existing,
                "word",
                TimeSpan.FromSeconds(30));
            _ = WaitForVisibleOfficeEditorWindow(
                TimeSpan.FromSeconds(20));

            var session = client.GetSessionAsync(
                    sessionId,
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual(
                "wordOmml",
                session.ObjectMode,
                "Complex OMML editor acceptance did not create a wordOmml Session.");
            AssertEqual(
                "create",
                session.Mode,
                "Complex OMML editor acceptance did not create a create Session.");

            // Load the exact user-reported source into the real Office editor.
            // Reopen the same Session after patching so React/MathLive initializes
            // from this source instead of racing the original blank create payload.
            var lineId =
                session.Lines.FirstOrDefault()?.Id
                ?? throw new InvalidDataException(
                    "Complex OMML create Session has no editable line.");
            // Move the reusable WebView to a temporary Session first.
            // That stops the blank create page from autosaving over the source
            // while we seed the real Session from the user-reported formula.
            var reloadLineId = Guid.NewGuid().ToString("D");
            var reloadSession = client.CreateSessionAsync(
                    new CreateVstoSessionRequest
                    {
                        Mode = "create",
                        Host = "word",
                        FormulaId = Guid.NewGuid().ToString("D"),
                        SourceDocumentId = Guid.NewGuid().ToString("D"),
                        SourceObjectId = "acceptance-reload",
                        Title = "Complex OMML reload bridge",
                        Lines = new List<FormulaLine>
                        {
                            new()
                            {
                                Id = reloadLineId,
                                Latex = "x",
                            },
                        },
                        ActiveLineId = reloadLineId,
                        CodeFormat = "latex",
                        DisplayMode = "inline",
                        ObjectMode = FormulaOleContract.WordOmmlMode,
                        FontSizePt = 11,
                    },
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            reloadSessionId = reloadSession.Id;
            client.OpenEditorAsync(
                    reloadSessionId,
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            Thread.Sleep(500);

            session = client.PatchAsync(
                    sessionId,
                    new
                    {
                        lines = new[]
                        {
                            new
                            {
                                id = lineId,
                                latex = expectedLatex,
                            },
                        },
                        activeLineId = lineId,
                        codeFormat = "latex",
                        displayMode = "block",
                        objectMode = "wordOmml",
                        numbered = false,
                        dirty = true,
                        status = "editing",
                        error = (string?)null,
                    },
                    CancellationToken.None)
                .GetAwaiter().GetResult();

            // Switching back to a different Session id uses the production
            // visualtex-office-session event and forces useOfficeSession to GET
            // the freshly patched payload without closing/cancelling the window.
            client.OpenEditorAsync(
                    sessionId,
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            var editorWindow =
                WaitForVisibleOfficeEditorWindow(
                    TimeSpan.FromSeconds(20));
            Thread.Sleep(900);
            var reloaded = client.GetSessionAsync(
                    sessionId,
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual(
                expectedLatex,
                reloaded.Lines.FirstOrDefault()?.Latex ?? string.Empty,
                "OfficeDialog reloaded a source different from the exact user-reported formula.");

            client.PatchAsync(
                    reloadSessionId,
                    new
                    {
                        status = "cancelled",
                        explicitCancel = true,
                        error = (string?)null,
                    },
                    CancellationToken.None)
                .GetAwaiter().GetResult();

            if (!SetForegroundWindow(editorWindow))
                throw new InvalidOperationException(
                    "Unable to foreground the VisualTeX Office editor for Ctrl+S Apply.");
            Thread.Sleep(300);

            // Ctrl+S is OfficeDialogApp's production Apply-and-close shortcut.
            // From here onward the acceptance does not patch committing and
            // does not call CloseEditorAsync: the real editor owns export,
            // commit, host acknowledgement, and window closure.
            var apply = Stopwatch.StartNew();
            WinForms.SendKeys.SendWait("^s");

            OfficeSessionDocument? terminal = null;
            TimeSpan? completedElapsed = null;
            TimeSpan? hiddenElapsed = null;
            var observationDeadline =
                DateTime.UtcNow + TimeSpan.FromSeconds(30);
            while (DateTime.UtcNow < observationDeadline
                   && (terminal is null || hiddenElapsed is null))
            {
                WinForms.Application.DoEvents();
                if (hiddenElapsed is null
                    && FindVisibleOfficeEditorWindow() == IntPtr.Zero)
                {
                    hiddenElapsed = apply.Elapsed;
                }

                var current = client.GetSessionAsync(
                        sessionId,
                        CancellationToken.None)
                    .GetAwaiter().GetResult();
                if (current.Status is
                    "completed" or "failed" or "cancelled")
                {
                    terminal = current;
                    completedElapsed ??= apply.Elapsed;
                }
                if (terminal is not null && hiddenElapsed is not null)
                    break;
                Thread.Sleep(40);
            }

            if (terminal is null)
                throw new TimeoutException(
                    "Complex OMML editor Ctrl+S Apply did not reach a terminal Session state.");
            AssertEqual(
                "completed",
                terminal.Status,
                terminal.Error
                ?? "Complex OMML editor Ctrl+S Apply did not complete.");
            if (hiddenElapsed is null)
                throw new TimeoutException(
                    "Complex OMML editor did not hide after the completed Session.");

            // Gate cleanup is host bookkeeping and must not be counted as the
            // user-visible editor close latency.
            WaitForAddInIdle(
                addIn,
                TimeSpan.FromSeconds(10));
            if (hiddenElapsed.Value >= TimeSpan.FromSeconds(3))
            {
                throw new InvalidDataException(
                    "Complex OMML Apply completed but the Office editor did not close through the normal Session flow quickly enough. "
                    + $"completedMs={completedElapsed.Value.TotalMilliseconds:0}; "
                    + $"hiddenMs={hiddenElapsed.Value.TotalMilliseconds:0}.");
            }

            AssertEqual(
                1,
                document.OMaths.Count,
                "Complex OMML editor commit did not leave exactly one Word equation.");
            math = document.OMaths[1];
            range = math.Range.Duplicate;
            var resolved =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    range,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Complex OMML editor commit could not resolve the inserted Word equation.");
            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    resolved);
            if (string.IsNullOrWhiteSpace(payload.WordOpenXml))
                throw new InvalidDataException(
                    "Complex OMML editor commit returned no semantic WordOpenXML.");
            var comparison =
                WordNativeOmmlSemanticComparer
                    .CompareMathMlToWordOpenXml(
                        expectedMathMl,
                        payload.WordOpenXml!,
                        display: true);
            if (!comparison.Equivalent)
            {
                throw new InvalidDataException(
                    "Complex OMML editor commit changed mathematical semantics. "
                    + $"expectedMathMlSignature=[{comparison.ExpectedMathMlSignature}]; "
                    + $"actualMathMlSignature=[{comparison.ActualMathMlSignature}]; "
                    + $"error=[{comparison.SemanticExtractionError ?? string.Empty}].");
            }

            Console.WriteLine(
                "[COMPLEX OMML EDITOR APPLY PASS] "
                + $"sessionId={sessionId}; "
                + $"completedMs={completedElapsed.Value.TotalMilliseconds:0}; "
                + $"hiddenMs={hiddenElapsed.Value.TotalMilliseconds:0}; "
                + $"semanticSignature={comparison.ActualMathMlSignature}");
            sessionId = null;
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(reloadSessionId))
            {
                try
                {
                    client.PatchAsync(
                            reloadSessionId!,
                            new
                            {
                                status = "cancelled",
                                explicitCancel = true,
                                error = (string?)null,
                            },
                            CancellationToken.None)
                        .GetAwaiter().GetResult();
                }
                catch { }
            }
            if (!string.IsNullOrWhiteSpace(sessionId))
            {
                try
                {
                    var current = client.GetSessionAsync(
                            sessionId!,
                            CancellationToken.None)
                        .GetAwaiter().GetResult();
                    if (current.Status is not
                        ("completed" or "cancelled" or "failed"))
                    {
                        client.PatchAsync(
                                sessionId!,
                                new
                                {
                                    status = "cancelled",
                                    explicitCancel = true,
                                    error = (string?)null,
                                },
                                CancellationToken.None)
                            .GetAwaiter().GetResult();
                    }
                    client.CloseEditorAsync(
                            sessionId!,
                            CancellationToken.None)
                        .GetAwaiter().GetResult();
                }
                catch { }
            }
            if (addIn is not null)
            {
                try
                {
                    addIn.OnDisconnection(
                        ext_DisconnectMode.ext_dm_UserClosed,
                        ref custom);
                }
                catch { }
            }
            Release(range);
            Release(math);
            if (document is not null)
            {
                try
                {
                    document.Close(
                        Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            try
            {
                QuitWordApplicationIfOwned(application);
            }
            catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void AssertStaleRibbonSessionRecovery(
        VisualTeXSessionClient client,
        VisualTeX.WordVsto.ThisAddIn addIn,
        Word.Document document,
        string formulaId)
    {
        var flags = System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.NonPublic;
        var gateField = addIn.GetType().GetField("_operationGate", flags)
            ?? throw new MissingFieldException("Word add-in operation gate is missing.");
        var idField = addIn.GetType().GetField("_activeSessionId", flags)
            ?? throw new MissingFieldException("Word add-in active Session id field is missing.");
        var cancellationField = addIn.GetType().GetField("_activeSessionCancellation", flags)
            ?? throw new MissingFieldException("Word add-in active Session cancellation field is missing.");
        var gate = gateField.GetValue(addIn) as SemaphoreSlim
            ?? throw new InvalidOperationException("Word add-in operation gate is unavailable.");
        using var staleCancellation = new CancellationTokenSource();
        if (!gate.Wait(TimeSpan.FromSeconds(2)))
            throw new InvalidOperationException("Could not reserve the Word operation gate for stale-session acceptance.");

        var staleSessionId = Guid.NewGuid().ToString("D");
        idField.SetValue(addIn, staleSessionId);
        cancellationField.SetValue(addIn, staleCancellation);
        var released = 0;
        var staleUnwind = Task.Run(() =>
        {
            staleCancellation.Token.WaitHandle.WaitOne();
            if (Interlocked.Exchange(ref released, 1) == 0)
                gate.Release();
        });

        string? recoveredSessionId = null;
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
        try
        {
            SelectOmmlFormula(document, formulaId, ref bookmark, ref equationRange);
            var before = SnapshotSessionIds();
            addIn.OnEditSelected(new object());
            recoveredSessionId = WaitForNewSession(before, "word", TimeSpan.FromSeconds(30));
            AssertTrue(staleCancellation.IsCancellationRequested,
                "Ribbon stale-session recovery did not cancel the missing local Session waiter.");
            AssertTrue(!string.Equals(staleSessionId, recoveredSessionId, StringComparison.OrdinalIgnoreCase),
                "Ribbon stale-session recovery reused the missing Session id.");
            _ = WaitForVisibleOfficeEditorWindow(TimeSpan.FromSeconds(20));

            client.PatchAsync(
                    recoveredSessionId,
                    new
                    {
                        status = "cancelled",
                        explicitCancel = true,
                        error = (string?)null,
                    },
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            client.CloseEditorAsync(recoveredSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(20));
            recoveredSessionId = null;
            AssertTrue(gate.CurrentCount == 1,
                "Ribbon stale-session recovery did not restore the operation gate to idle.");
            Console.WriteLine(
                "    stale Ribbon Session recovery passed: missing Session was cancelled locally and the same click opened a fresh editor.");
        }
        finally
        {
            Release(equationRange);
            Release(bookmark);
            if (!string.IsNullOrWhiteSpace(recoveredSessionId))
            {
                try
                {
                    client.PatchAsync(
                            recoveredSessionId!,
                            new
                            {
                                status = "cancelled",
                                explicitCancel = true,
                                error = (string?)null,
                            },
                            CancellationToken.None)
                        .GetAwaiter().GetResult();
                    client.CloseEditorAsync(recoveredSessionId!, CancellationToken.None)
                        .GetAwaiter().GetResult();
                }
                catch { }
            }
            if (!staleCancellation.IsCancellationRequested)
            {
                try { staleCancellation.Cancel(); } catch { }
            }
            try { staleUnwind.Wait(TimeSpan.FromSeconds(3)); } catch { }
            // If the recovery path failed before consuming the synthetic gate,
            // restore the fixture to idle without touching a gate owned by a real
            // RunSessionAsync task.
            if (Interlocked.CompareExchange(ref released, 1, 1) == 0
                && gate.CurrentCount == 0)
            {
                gate.Release();
                Interlocked.Exchange(ref released, 1);
            }
        }
    }

    private static void AssertDirectOfficeEditorCloseLatency(
        VisualTeXSessionClient client)
    {
        var lineId = Guid.NewGuid().ToString("D");
        var session = client.CreateSessionAsync(
                new CreateVstoSessionRequest
                {
                    Mode = "create",
                    Host = "word",
                    FormulaId = Guid.NewGuid().ToString("D"),
                    SourceDocumentId = Guid.NewGuid().ToString("D"),
                    SourceObjectId = "direct-close-latency",
                    Title = "Direct close latency acceptance",
                    Lines = new List<FormulaLine>
                    {
                        new() { Id = lineId, Latex = "x" },
                    },
                    ActiveLineId = lineId,
                    CodeFormat = "latex",
                    DisplayMode = "inline",
                    ObjectMode = FormulaOleContract.WordOmmlMode,
                    FontSizePt = 11,
                    AutoCommitOnClose = false,
                },
                CancellationToken.None)
            .GetAwaiter().GetResult();

        try
        {
            client.OpenEditorAsync(
                    session.Id,
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            _ = WaitForVisibleOfficeEditorWindow(
                TimeSpan.FromSeconds(20));

            var watch = Stopwatch.StartNew();
            client.CloseEditorAsync(
                    session.Id,
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            var responseElapsed = watch.Elapsed;
            WaitForOfficeEditorHidden(
                TimeSpan.FromSeconds(10));
            var hiddenElapsed = watch.Elapsed;

            Console.WriteLine(
                "[DIRECT OFFICE EDITOR CLOSE] "
                + $"responseMs={responseElapsed.TotalMilliseconds:0}; "
                + $"hiddenMs={hiddenElapsed.TotalMilliseconds:0}");
        }
        finally
        {
            try
            {
                var current = client.GetSessionAsync(
                        session.Id,
                        CancellationToken.None)
                    .GetAwaiter().GetResult();
                if (current.Status is not ("completed" or "cancelled" or "failed"))
                {
                    client.PatchAsync(
                            session.Id,
                            new
                            {
                                status = "cancelled",
                                explicitCancel = true,
                                error = (string?)null,
                            },
                            CancellationToken.None)
                        .GetAwaiter().GetResult();
                }
            }
            catch { }
        }
    }

    private static void AssertDirtyOmmlClosePolicy(VisualTeXSessionClient client)
    {
        var formulaId = Guid.NewGuid().ToString("D");
        var lineId = Guid.NewGuid().ToString("D");
        var metadata = new FormulaMetadata
        {
            FormulaId = formulaId,
            Title = "Dirty OMML close policy",
            Latex = "x+y",
            Lines = new List<FormulaLine>
            {
                new() { Id = lineId, Latex = "x+y" },
            },
            CodeFormat = "latex",
            DisplayMode = "inline",
            FontSizePt = 11,
            CreatedWithVersion = "1.2.5",
            UpdatedWithVersion = "1.2.5",
            CreatedAt = DateTimeOffset.UtcNow.ToString("O"),
            UpdatedAt = DateTimeOffset.UtcNow.ToString("O"),
        };
        var session = client.CreateSessionAsync(
                new CreateVstoSessionRequest
                {
                    Mode = "edit",
                    Host = "word",
                    FormulaId = formulaId,
                    SourceDocumentId = Guid.NewGuid().ToString("D"),
                    SourceObjectId = "0:1",
                    Title = metadata.Title,
                    Lines = metadata.Lines,
                    ActiveLineId = lineId,
                    CodeFormat = "latex",
                    DisplayMode = "inline",
                    ObjectMode = FormulaOleContract.WordOmmlMode,
                    FontSizePt = 11,
                    OriginalMetadata = metadata,
                    AutoCommitOnClose = true,
                },
                CancellationToken.None)
            .GetAwaiter().GetResult();
        const string svg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 120 30\"></svg>";
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mi>x</mi><mo>+</mo><mi>z</mi></math>";
        try
        {
            client.PatchAsync(
                    session.Id,
                    new
                    {
                        lines = new[]
                        {
                            new { id = lineId, latex = "x+z" },
                        },
                        dirty = true,
                        status = "editing",
                        exportResult = new
                        {
                            svg,
                            svgBase64 = "data:image/svg+xml;base64,"
                                + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg)),
                            mathMl,
                            pngBase64 = (string?)null,
                            width = 120d,
                            height = 30d,
                            baseline = 22d,
                        },
                        exportWidth = 120d,
                        exportHeight = 30d,
                        error = (string?)null,
                    },
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            client.CloseEditorAsync(session.Id, CancellationToken.None)
                .GetAwaiter().GetResult();
            var closing = client.GetSessionAsync(session.Id, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual(
                "committing",
                closing.Status,
                "A dirty Word OMML Session with valid MathML but no PNG was not committed on close.");
            client.CompleteAsync(session.Id, CancellationToken.None)
                .GetAwaiter().GetResult();
            Console.WriteLine(
                "    dirty OMML close policy passed with valid MathML and no PNG payload");
        }
        finally
        {
            try
            {
                var current = client.GetSessionAsync(session.Id, CancellationToken.None)
                    .GetAwaiter().GetResult();
                if (current.Status is not ("completed" or "cancelled" or "failed"))
                {
                    client.PatchAsync(
                            session.Id,
                            new
                            {
                                status = "cancelled",
                                explicitCancel = true,
                                error = (string?)null,
                            },
                            CancellationToken.None)
                        .GetAwaiter().GetResult();
                }
            }
            catch { }
        }
    }

    private static void SelectOmmlFormula(
        Word.Document document,
        string formulaId,
        ref Word.Bookmark? bookmark,
        ref Word.Range? equationRange)
    {
        bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
        if (bookmark is not null)
        {
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            equationRange.Select();
            return;
        }

        // Current native OMML deliberately carries no VisualTeX durable
        // bookmark/identity. This acceptance fixture owns a document with one
        // OMath, so select that exact Word-native host instead of reintroducing
        // retired VisualTeX metadata just to satisfy the test harness.
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        try
        {
            maths = document.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    $"The native close fixture expected one bookmark-free OMath; found {maths.Count}.");
            math = maths[1];
            equationRange = math.Range.Duplicate;
            equationRange.Select();
        }
        finally
        {
            Release(math);
            Release(maths);
        }
    }

    private static IntPtr WaitForVisibleOfficeEditorWindow(TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            var handle = FindVisibleOfficeEditorWindow();
            if (handle != IntPtr.Zero) return handle;
            Thread.Sleep(100);
        }
        throw new TimeoutException("The VisualTeX Office editor window did not become visible.");
    }

    private static void WaitForOfficeEditorHidden(TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            if (FindVisibleOfficeEditorWindow() == IntPtr.Zero) return;
            Thread.Sleep(100);
        }
        throw new TimeoutException("The reusable VisualTeX Office editor remained visible after close.");
    }

    private static IntPtr FindVisibleOfficeEditorWindow()
    {
        var result = IntPtr.Zero;
        EnumWindows((windowHandle, _) =>
        {
            if (!IsWindowVisible(windowHandle)) return true;
            GetWindowThreadProcessId(windowHandle, out var processId);
            if (processId == 0) return true;
            try
            {
                using var process = Process.GetProcessById((int)processId);
                if (!string.Equals(
                        process.ProcessName,
                        "visualtex",
                        StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch
            {
                return true;
            }
            var length = GetWindowTextLength(windowHandle);
            if (length <= 0) return true;
            var text = new StringBuilder(length + 1);
            GetWindowText(windowHandle, text, text.Capacity);
            var title = text.ToString();
            if (title.IndexOf("Office", StringComparison.OrdinalIgnoreCase) < 0)
                return true;
            result = windowHandle;
            return false;
        }, IntPtr.Zero);
        return result;
    }
}
