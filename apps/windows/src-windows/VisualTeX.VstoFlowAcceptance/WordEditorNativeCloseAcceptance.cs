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
        bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
            ?? throw new InvalidDataException("The native close fixture lost its OMML bookmark.");
        equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
        equationRange.Select();
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
