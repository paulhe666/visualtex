using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunMathTypeAcceptanceWindowCleanup()
    {
        var processes = System.Diagnostics.Process.GetProcessesByName("MathType");
        try
        {
            foreach (var process in processes)
            {
                DismissMathTypeOleCloseDialogByCancel();
                process.Refresh();
                var originalHandle = process.MainWindowHandle;
                var originalTitle = process.MainWindowTitle ?? string.Empty;
                var temporaryWindows = new List<(IntPtr Handle, string Title)>();
                EnumWindowsForMathType((window, _) =>
                {
                    GetWindowThreadProcessId(window, out var processId);
                    if (processId != (uint)process.Id) return true;
                    var className = new StringBuilder(128);
                    GetClassNameForMathType(window, className, className.Capacity);
                    if (!string.Equals(className.ToString(), "EQNWINCLASS", StringComparison.Ordinal))
                        return true;
                    var title = GetWindowTitleForMathType(window);
                    if (IsAcceptanceUntitledMathTypeWindow(title))
                        temporaryWindows.Add((window, title));
                    return true;
                }, IntPtr.Zero);

                DismissMathTypeTooManyWindowsWarning(process.Id);
                Console.WriteLine(
                    $"MathType PID={process.Id}: cleaning {temporaryWindows.Count} acceptance-created Untitled windows; preserving 0x{originalHandle.ToInt64():X} '{originalTitle}'.");
                foreach (var temporary in temporaryWindows
                             .OrderByDescending(item => ExtractUntitledMathTypeNumber(item.Title)))
                {
                    Console.WriteLine(
                        $"  closing acceptance window 0x{temporary.Handle.ToInt64():X} '{temporary.Title}'");
                    PostMessageForMathType(
                        temporary.Handle,
                        0x0010,
                        IntPtr.Zero,
                        IntPtr.Zero); // WM_CLOSE
                    var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(3);
                    while (DateTime.UtcNow < deadline && IsWindowForMathType(temporary.Handle))
                    {
                        ClickMathTypeDialogButtonOwnedBy(
                            process.Id,
                            temporary.Handle,
                            7); // IDNO: discard only this exact acceptance temp doc
                        Thread.Sleep(100);
                    }
                    if (IsWindowForMathType(temporary.Handle))
                        throw new TimeoutException(
                            $"Acceptance-created MathType window '{temporary.Title}' did not close.");
                }

                process.Refresh();
                if (originalHandle != IntPtr.Zero && !IsWindowForMathType(originalHandle))
                    throw new InvalidOperationException(
                        "Acceptance cleanup unexpectedly closed the user's original MathType window.");
                Console.WriteLine(
                    $"  preserved original MathType window: '{GetWindowTitleForMathType(originalHandle)}'.");
            }
        }
        finally
        {
            foreach (var process in processes) process.Dispose();
        }
    }

    private static bool IsAcceptanceUntitledMathTypeWindow(string title)
    {
        if (string.IsNullOrWhiteSpace(title)) return false;
        return title.StartsWith("MathType - 无标题 ", StringComparison.Ordinal)
            || title.StartsWith("MathType - Untitled ", StringComparison.OrdinalIgnoreCase);
    }

    private static int ExtractUntitledMathTypeNumber(string title)
    {
        var lastSpace = title.LastIndexOf(' ');
        return lastSpace >= 0 && int.TryParse(title.Substring(lastSpace + 1), out var number)
            ? number
            : 0;
    }

    private static void DismissMathTypeTooManyWindowsWarning(int processId)
    {
        EnumWindowsForMathType((window, _) =>
        {
            GetWindowThreadProcessId(window, out var ownerPid);
            if (ownerPid != (uint)processId) return true;
            var className = new StringBuilder(64);
            GetClassNameForMathType(window, className, className.Capacity);
            if (!string.Equals(className.ToString(), "#32770", StringComparison.Ordinal))
                return true;
            var text = GetDialogStaticTextForMathType(window);
            if (text.IndexOf("打开的窗口太多", StringComparison.Ordinal) < 0)
                return true;
            var okButton = GetDlgItemForMathType(window, 2);
            if (okButton != IntPtr.Zero)
            {
                SendMessageForMathType(okButton, 0x00F5, IntPtr.Zero, IntPtr.Zero);
                Console.WriteLine("  dismissed acceptance-created 'too many windows' MathType warning.");
            }
            return false;
        }, IntPtr.Zero);
        Thread.Sleep(200);
    }

    private static void RunMathTypeWindowInventory()
    {
        var processes = System.Diagnostics.Process.GetProcessesByName("MathType");
        try
        {
            foreach (var process in processes)
            {
                process.Refresh();
                Console.WriteLine(
                    $"MathType PID={process.Id} main=0x{process.MainWindowHandle.ToInt64():X} title='{process.MainWindowTitle}'");
                EnumWindowsForMathType((window, _) =>
                {
                    GetWindowThreadProcessId(window, out var processId);
                    if (processId != (uint)process.Id) return true;
                    var className = new StringBuilder(128);
                    GetClassNameForMathType(window, className, className.Capacity);
                    var title = GetWindowTitleForMathType(window);
                    Console.WriteLine(
                        $"  HWND=0x{window.ToInt64():X} visible={IsWindowVisibleForMathType(window)} class='{className}' title='{title}'");
                    if (string.Equals(className.ToString(), "#32770", StringComparison.Ordinal))
                    {
                        var text = GetDialogStaticTextForMathType(window);
                        Console.WriteLine($"    dialogText='{text}'");
                        if (string.Equals(text, "打开的窗口太多。", StringComparison.Ordinal))
                        {
                            var button = GetDlgItemForMathType(window, 2);
                            if (button != IntPtr.Zero)
                            {
                                SendMessageForMathType(button, 0x00F5, IntPtr.Zero, IntPtr.Zero);
                                Console.WriteLine("    dismissed acceptance-created 'too many windows' warning.");
                            }
                        }
                    }
                    return true;
                }, IntPtr.Zero);
            }
        }
        finally
        {
            foreach (var process in processes) process.Dispose();
        }
    }

    private static string GetWindowTitleForMathType(IntPtr window)
    {
        var builder = new StringBuilder(512);
        GetWindowTextForMathType(window, builder, builder.Capacity);
        return builder.ToString();
    }

    private static string GetDialogStaticTextForMathType(IntPtr dialog)
    {
        var parts = new List<string>();
        EnumChildWindowsForMathType(dialog, (child, _) =>
        {
            var className = new StringBuilder(64);
            GetClassNameForMathType(child, className, className.Capacity);
            if (string.Equals(className.ToString(), "Static", StringComparison.Ordinal))
            {
                var text = GetWindowTitleForMathType(child);
                if (!string.IsNullOrWhiteSpace(text)) parts.Add(text);
            }
            return true;
        }, IntPtr.Zero);
        return string.Join(" ", parts);
    }

    private static void RunMathTypeSetFormatProbe()
    {
        Console.WriteLine("MathType standalone SetData format probe:");
        var result = MathTypeOleInterop.ProbeStandaloneSetFormats();
        Console.WriteLine("  " + result);
    }

    private static void RunMathTypeUiMathMlClipboardAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow><mi>y</mi></mfrac></math>";
        const string inputLatex = @"\frac{x+1}{y}";
        var path = Path.Combine(artifactRoot, "VisualTeX-MathType7-UI-MathML-Clipboard.docx");
        if (File.Exists(path)) File.Delete(path);

        System.Diagnostics.Process? mathType = null;
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.Shape? floatingShape = null;
        Word.OLEFormat? format = null;
        var clipboardBackup = TryCaptureClipboard();
        IntPtr originalWindow = IntPtr.Zero;
        IntPtr temporaryWindow = IntPtr.Zero;
        string originalTitle = string.Empty;
        string temporaryTitle = string.Empty;
        var temporaryDocumentCreated = false;
        try
        {
            var candidates = System.Diagnostics.Process.GetProcessesByName("MathType");
            try
            {
                mathType = candidates
                    .OrderBy(process => process.StartTime)
                    .FirstOrDefault(process => process.MainWindowHandle != IntPtr.Zero)
                    ?? throw new InvalidOperationException(
                        "No interactive MathType 7 window is available for the UI clipboard acceptance.");
            }
            finally
            {
                foreach (var candidate in candidates)
                {
                    if (!ReferenceEquals(candidate, mathType)) candidate.Dispose();
                }
            }
            mathType.Refresh();
            originalWindow = mathType.MainWindowHandle;
            originalTitle = mathType.MainWindowTitle ?? string.Empty;
            if (originalWindow == IntPtr.Zero || string.IsNullOrWhiteSpace(originalTitle))
                throw new InvalidOperationException("The interactive MathType 7 window has no stable HWND/title.");
            Console.WriteLine(
                $"[MathType UI clipboard 1/5] Preserving interactive MathType PID={mathType.Id}, title='{originalTitle}'.");

            var existingEquationWindows = GetMathTypeEquationWindows(mathType.Id)
                .Select(item => item.Handle)
                .ToHashSet();
            using (var launched = System.Diagnostics.Process.Start(
                       new System.Diagnostics.ProcessStartInfo
                       {
                           FileName = @"C:\Program Files (x86)\MathType\MathType.exe",
                           Arguments = "-new",
                           UseShellExecute = true,
                       }))
            {
            }
            var temporary = WaitForTemporaryMathTypeEquationWindow(
                mathType.Id,
                existingEquationWindows,
                originalWindow,
                originalTitle,
                TimeSpan.FromSeconds(30));
            temporaryWindow = temporary.Handle;
            temporaryTitle = temporary.Title;
            temporaryDocumentCreated = true;
            Console.WriteLine(
                $"  temporary MathType document HWND=0x{temporaryWindow.ToInt64():X}, title='{temporaryTitle}'.");

            Console.WriteLine("[MathType UI clipboard 2/5] Inspecting and clearing only the temporary MathType document before feeding MathML...");
            SetForegroundWindow(temporaryWindow);
            Thread.Sleep(350);
            System.Windows.Forms.SendKeys.SendWait("^a");
            Thread.Sleep(100);
            System.Windows.Forms.SendKeys.SendWait("^c");
            Thread.Sleep(500);
            var temporaryInitialClipboard = System.Windows.Forms.Clipboard.GetDataObject();
            if (temporaryInitialClipboard is not null)
            {
                try
                {
                    var initialMathMl = ReadMathMlFromClipboardDataObject(temporaryInitialClipboard);
                    var initialLatex = MathMlToLatexConverter.Convert(initialMathMl)
                        .Replace(" ", string.Empty);
                    Console.WriteLine("  temporary MathType initial LaTeX=" + initialLatex);
                }
                catch (InvalidDataException)
                {
                    Console.WriteLine("  temporary MathType document has no readable initial equation.");
                }
            }
            System.Windows.Forms.SendKeys.SendWait("^a");
            Thread.Sleep(100);
            System.Windows.Forms.SendKeys.SendWait("{DELETE}");
            Thread.Sleep(250);

            Console.WriteLine("  feeding requested TeX after clearing the temporary document...");
            System.Windows.Forms.Clipboard.SetText(
                inputLatex,
                System.Windows.Forms.TextDataFormat.UnicodeText);
            SetForegroundWindow(temporaryWindow);
            Thread.Sleep(250);
            System.Windows.Forms.SendKeys.SendWait("^v");
            // Pasted TeX is converted by MathType immediately even when keyboard
            // TeX input is disabled. Do not send Enter here: after paste it means
            // a real equation line break and would create an empty matrix row.
            Thread.Sleep(700);
            System.Windows.Forms.SendKeys.SendWait("^a");
            Thread.Sleep(120);
            System.Windows.Forms.SendKeys.SendWait("^c");
            Thread.Sleep(700);
            var copiedDataObject = System.Windows.Forms.Clipboard.GetDataObject()
                ?? throw new InvalidOperationException(
                    "MathType 7 did not populate the clipboard after importing MathML.");
            var copiedFormats = copiedDataObject.GetFormats();
            Console.WriteLine("  MathType copy formats: " + string.Join(", ", copiedFormats));
            AssertTrue(
                copiedFormats.Any(value =>
                    value.IndexOf("MathType", StringComparison.OrdinalIgnoreCase) >= 0
                    || value.IndexOf("Embed", StringComparison.OrdinalIgnoreCase) >= 0
                    || value.IndexOf("Object", StringComparison.OrdinalIgnoreCase) >= 0),
                "MathType 7 copied the MathML equation but did not expose an OLE/native clipboard format.");
            var copiedMathMl = ReadMathMlFromClipboardDataObject(copiedDataObject);
            var copiedLatex = MathMlToLatexConverter.Convert(copiedMathMl)
                .Replace(" ", string.Empty);
            Console.WriteLine("  MathType imported LaTeX=" + copiedLatex);
            AssertTrue(
                copiedLatex.IndexOf("x+1", StringComparison.Ordinal) >= 0
                && copiedLatex.IndexOf("y", StringComparison.Ordinal) >= 0
                && copiedLatex.IndexOf("begin{matrix}", StringComparison.OrdinalIgnoreCase) < 0,
                $"MathType 7 did not import the requested TeX cleanly before creating the OLE. LaTeX='{copiedLatex}'.");

            // MathType 7 exposes Embedded Object through delayed rendering. Its
            // OleFlushClipboard implementation returns E_FAIL for this clipboard,
            // but Word can consume the live OLE source correctly while the known
            // temporary equation is still open. Paste first, then close only that
            // temporary document and restore the user's original MathType session.
            Console.WriteLine("[MathType UI clipboard 3/5] Pasting MathType's live OLE clipboard into an isolated modern Word document...");
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Activate();
            application.Selection.TypeText("MathType UI clipboard acceptance: ");
            application.Selection.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteOLEObject);
            Thread.Sleep(400);
            Console.WriteLine(
                $"  Word paste inventory: inline={document.InlineShapes.Count}, floating={document.Shapes.Count}");
            if (document.InlineShapes.Count == 1)
            {
                shape = document.InlineShapes[1];
            }
            else if (document.InlineShapes.Count == 0 && document.Shapes.Count == 1)
            {
                floatingShape = document.Shapes[1];
                format = floatingShape.OLEFormat;
                AssertEqual("Equation.DSMT4", format.ProgID,
                    "MathType 7 clipboard paste created the wrong floating OLE class.");
                Release(format);
                format = null;
                shape = floatingShape.ConvertToInlineShape();
                Release(floatingShape);
                floatingShape = null;
            }
            else
            {
                throw new InvalidDataException(
                    "MathType 7 clipboard paste did not create exactly one Word OLE equation.");
            }
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                "MathType 7 clipboard-generated equation is not recognized as MathType OLE.");
            format = shape.OLEFormat;
            AssertEqual("Equation.DSMT4", format.ProgID,
                "MathType 7 clipboard-generated equation lost the Equation.DSMT4 ProgID.");
            Release(format);
            format = null;
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);

            Console.WriteLine("  Word consumed the live MathType OLE; restoring the original MathType session now.");
            CloseTemporaryMathTypeDocumentAndRestore(
                mathType,
                originalWindow,
                originalTitle,
                TimeSpan.FromSeconds(8));
            temporaryDocumentCreated = false;

            Console.WriteLine("[MathType UI clipboard 4/5] Verifying the generated OLE contains the requested equation without activating MathType...");
            var directFormats = MathTypeOleInterop.DescribeDataFormats(
                shape,
                activateForConversion: false);
            Console.WriteLine("  direct Word OLE formats: " + directFormats);
            if (!string.Equals(directFormats, "no-idataobject", StringComparison.Ordinal))
            {
                var directMathMl = MathTypeOleInterop.ReadMathMl(
                    shape,
                    activateForConversion: false);
                var directLatex = MathMlToLatexConverter.Convert(directMathMl)
                    .Replace(" ", string.Empty);
                Console.WriteLine("  direct LaTeX=" + directLatex);
                AssertTrue(
                    directLatex.IndexOf("x+1", StringComparison.Ordinal) >= 0
                    && directLatex.IndexOf("y", StringComparison.Ordinal) >= 0,
                    $"MathType clipboard OLE contains the wrong equation. LaTeX='{directLatex}'.");
            }

            Console.WriteLine("[MathType UI clipboard 5/5] Confirming the user's original MathType session remained restored after Word accepted the generated OLE...");
            mathType.Refresh();
            AssertEqual(originalTitle, mathType.MainWindowTitle,
                "Word OLE paste disturbed the user's restored MathType session.");
            Console.WriteLine(
                "MathType UI TeX-to-OLE acceptance passed: MathType itself generated the Equation.DSMT4 OLE, Word accepted the live OLE clipboard, and the original interactive MathType session stayed restored.");
        }
        finally
        {
            if (temporaryDocumentCreated && mathType is not null)
            {
                try
                {
                    CloseTemporaryMathTypeDocumentAndRestore(
                        mathType,
                        originalWindow,
                        originalTitle,
                        TimeSpan.FromSeconds(4));
                }
                catch
                {
                    // Never escalate to killing MathType. A failure to restore the
                    // temporary document is safer than discarding user state.
                }
            }
            TryRestoreClipboard(clipboardBackup);
            Release(format);
            Release(floatingShape);
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            mathType?.Dispose();
            ForceComCleanup();
        }
    }

    private static void RunWordMathTypeOleCopyFormatsAcceptance(string artifactRoot)
    {
        var path = Path.Combine(
            artifactRoot,
            "VisualTeX-MathType7-UI-MathML-Clipboard.docx");
        if (!File.Exists(path))
            throw new FileNotFoundException(
                "The MathType UI clipboard acceptance document is missing. Run mathtype-ui-mathml-clipboard first.",
                path);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        var clipboardBackup = TryCaptureClipboard();
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(path, ReadOnly: true, Visible: false);
            AssertEqual(1, document.InlineShapes.Count,
                "The MathType UI clipboard acceptance document no longer contains one inline OLE equation.");
            shape = document.InlineShapes[1];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                "The saved acceptance equation is no longer recognized as MathType OLE.");

            Console.WriteLine("[MathType Word-copy 1/3] Copying an existing Equation.DSMT4 from Word without activating MathType...");
            shape.Range.Copy();
            Thread.Sleep(500);
            var dataObject = System.Windows.Forms.Clipboard.GetDataObject()
                ?? throw new InvalidOperationException("Word copied the MathType OLE but the Windows clipboard is empty.");
            var formats = dataObject.GetFormats();
            Console.WriteLine("  Word copy formats: " + string.Join(", ", formats));

            Console.WriteLine("[MathType Word-copy 2/3] Converting Word's Embedded Object clipboard payload through a temporary MathType document...");
            string mathMl;
            try
            {
                mathMl = ReadMathMlFromClipboardDataObject(dataObject);
                Console.WriteLine("  Word clipboard already exposed MathML directly.");
            }
            catch (InvalidDataException)
            {
                mathMl = ConvertCurrentClipboardOleToMathMlThroughTemporaryMathType();
                Console.WriteLine("  MathType temporary document exported MathML from the Word Embedded Object.");
            }
            Console.WriteLine("  MathML=" + mathMl);
            var latex = MathMlToLatexConverter.Convert(mathMl).Replace(" ", string.Empty);
            Console.WriteLine("  LaTeX=" + latex);
            AssertTrue(
                latex.IndexOf("x+1", StringComparison.Ordinal) >= 0
                && latex.IndexOf("y", StringComparison.Ordinal) >= 0,
                $"Temporary MathType conversion recovered the wrong equation. LaTeX='{latex}'.");

            Console.WriteLine("[MathType Word-copy 3/3] Existing MathType OLE source can be imported through Word copy + an isolated temporary MathType document while restoring the user's original MathType session.");
        }
        finally
        {
            TryRestoreClipboard(clipboardBackup);
            Release(shape);
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

    private static string ConvertCurrentClipboardOleToMathMlThroughTemporaryMathType()
    {
        System.Diagnostics.Process? mathType = null;
        var temporaryCreated = false;
        string originalTitle = string.Empty;
        IntPtr mainWindow = IntPtr.Zero;
        try
        {
            var candidates = System.Diagnostics.Process.GetProcessesByName("MathType");
            try
            {
                mathType = candidates
                    .OrderBy(process => process.StartTime)
                    .FirstOrDefault(process => process.MainWindowHandle != IntPtr.Zero)
                    ?? throw new InvalidOperationException(
                        "No interactive MathType 7 window is available to decode the Word OLE clipboard payload.");
            }
            finally
            {
                foreach (var candidate in candidates)
                {
                    if (!ReferenceEquals(candidate, mathType)) candidate.Dispose();
                }
            }
            mathType.Refresh();
            originalTitle = mathType.MainWindowTitle ?? string.Empty;
            mainWindow = mathType.MainWindowHandle;
            if (mainWindow == IntPtr.Zero || string.IsNullOrWhiteSpace(originalTitle))
                throw new InvalidOperationException(
                    "MathType 7 has no stable interactive window for temporary OLE decoding.");

            using (var launched = System.Diagnostics.Process.Start(
                       new System.Diagnostics.ProcessStartInfo
                       {
                           FileName = @"C:\Program Files (x86)\MathType\MathType.exe",
                           Arguments = "-new",
                           UseShellExecute = true,
                       }))
            {
            }
            var temporaryTitle = WaitForMathTypeTitleChange(
                mathType,
                originalTitle,
                TimeSpan.FromSeconds(10));
            temporaryCreated = true;
            Console.WriteLine($"  decoder temporary MathType title='{temporaryTitle}'.");

            SetForegroundWindow(mainWindow);
            Thread.Sleep(350);
            System.Windows.Forms.SendKeys.SendWait("^v");
            Thread.Sleep(700);
            System.Windows.Forms.SendKeys.SendWait("^a");
            Thread.Sleep(100);
            System.Windows.Forms.SendKeys.SendWait("^c");
            Thread.Sleep(700);
            var convertedClipboard = System.Windows.Forms.Clipboard.GetDataObject()
                ?? throw new InvalidOperationException(
                    "MathType decoded the Word OLE but did not populate the clipboard.");
            var formats = convertedClipboard.GetFormats();
            Console.WriteLine("  decoder MathType copy formats: " + string.Join(", ", formats));
            var mathMl = ReadMathMlFromClipboardDataObject(convertedClipboard);

            CloseTemporaryMathTypeDocumentAndRestore(
                mathType,
                mainWindow,
                originalTitle,
                TimeSpan.FromSeconds(8));
            temporaryCreated = false;
            return mathMl;
        }
        finally
        {
            if (temporaryCreated && mathType is not null)
            {
                try
                {
                    CloseTemporaryMathTypeDocumentAndRestore(
                        mathType,
                        mainWindow,
                        originalTitle,
                        TimeSpan.FromSeconds(4));
                }
                catch { }
            }
            mathType?.Dispose();
        }
    }

    private static string ReadMathMlFromClipboardDataObject(System.Windows.Forms.IDataObject dataObject)
    {
        foreach (var formatName in new[] { "MathML", "MathML Presentation", "application/mathml+xml" })
        {
            if (!dataObject.GetDataPresent(formatName, false)) continue;
            var payload = dataObject.GetData(formatName, false);
            switch (payload)
            {
                case string text:
                {
                    var extracted = ExtractMathMlRoot(text);
                    if (extracted is not null) return extracted;
                    break;
                }
                case MemoryStream stream:
                {
                    var bytes = stream.ToArray();
                    var decoded = DecodeClipboardText(bytes);
                    var extracted = ExtractMathMlRoot(decoded);
                    if (extracted is not null) return extracted;
                    Console.WriteLine(
                        $"  {formatName} stream length={bytes.Length}, head={BitConverter.ToString(bytes.Take(Math.Min(bytes.Length, 96)).ToArray())}");
                    Console.WriteLine(
                        $"  {formatName} decoded head='{decoded.Substring(0, Math.Min(decoded.Length, 160)).Replace("\0", "<NUL>")}'");
                    break;
                }
                case byte[] bytes:
                {
                    var decoded = DecodeClipboardText(bytes);
                    var extracted = ExtractMathMlRoot(decoded);
                    if (extracted is not null) return extracted;
                    break;
                }
            }
            Console.WriteLine(
                $"  format {formatName} payload type={payload?.GetType().FullName ?? "<null>"}");
        }
        throw new InvalidDataException(
            "Word's clipboard representation of the MathType OLE did not expose readable MathML.");
    }

    private static string? ExtractMathMlRoot(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var text = value.Trim().TrimEnd('\0');
        var start = text.IndexOf("<math", StringComparison.OrdinalIgnoreCase);
        return start < 0 ? null : text.Substring(start).Trim().TrimEnd('\0');
    }

    private static string DecodeClipboardText(byte[] bytes)
    {
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2).TrimEnd('\0');
        if (bytes.Length >= 2 && bytes[0] == (byte)'<' && bytes[1] == 0)
            return Encoding.Unicode.GetString(bytes).TrimEnd('\0');
        try { return new UTF8Encoding(false, true).GetString(bytes).TrimEnd('\0'); }
        catch (DecoderFallbackException)
        {
            return Encoding.Default.GetString(bytes).TrimEnd('\0');
        }
    }

    private static void RunMathTypeStandaloneComAcceptance()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow><mi>y</mi></mfrac></math>";
        Console.WriteLine("[MathType standalone COM 1/3] Enumerating a background Equation.DSMT4 object while the interactive MathType window remains open...");
        var formats = MathTypeOleInterop.DescribeStandaloneServerFormats();
        Console.WriteLine("  " + formats);
        AssertTrue(formats.IndexOf("MathML", StringComparison.OrdinalIgnoreCase) >= 0,
            "Standalone MathType COM object did not advertise MathML formats.");

        Console.WriteLine("[MathType standalone COM 2/3] Initializing an in-memory OLE storage and inspecting SET formats...");
        var initializedFormats = MathTypeOleInterop.DescribeInitializedStandaloneServerFormats();
        Console.WriteLine("  initialized: " + initializedFormats);
        Console.WriteLine("  Creating the same MathType class through OleCreate + IOleClientSite...");
        var clientFormats = MathTypeOleInterop.DescribeOleCreateClientFormats();
        Console.WriteLine("  OleCreate client: " + clientFormats);
        Console.WriteLine("  Probing IOleObject.GetClipboardData as a possible AttachDataObject equivalent...");
        try
        {
            var attachedFormats = MathTypeOleInterop.DescribeOleCreateAttachedDataFormats();
            Console.WriteLine("  attached data object: " + attachedFormats);
        }
        catch (NotImplementedException)
        {
            Console.WriteLine("  GetClipboardData is E_NOTIMPL on MathType 7; MFC AttachDataObject uses another route.");
        }
        Console.WriteLine("  Round-tripping MathML through the OleCreate IDataObject with CLR-marshalled HGLOBAL...");
        var roundTrip = MathTypeOleInterop.RoundTripOleCreateMathMl(mathMl);
        var latex = MathMlToLatexConverter.Convert(roundTrip).Replace(" ", string.Empty);
        AssertTrue(latex.IndexOf("x+1", StringComparison.Ordinal) >= 0
                   && latex.IndexOf("y", StringComparison.Ordinal) >= 0,
            $"Standalone MathType COM object did not round-trip the requested equation. LaTeX='{latex}'.");

        Console.WriteLine("[MathType standalone COM 3/3] Background MathType COM SetData/GetData passed without closing the user's interactive MathType window.");
    }

    private static void RunWordMathTypeDirectSetRealOleAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var candidatePaths = new[]
        {
            Path.Combine(
                Directory.GetParent(artifactRoot)?.FullName ?? artifactRoot,
                "mathtype-product-roundtrip",
                "VisualTeX-MathType7-UI-MathML-Clipboard.docx"),
            Path.Combine(
                Directory.GetParent(artifactRoot)?.FullName ?? artifactRoot,
                "mathtype-native-editor",
                "VisualTeX-MathType7-UI-MathML-Clipboard.docx"),
        };
        var sourcePath = candidatePaths.FirstOrDefault(File.Exists)
            ?? throw new FileNotFoundException(
                "No genuine MathType-generated non-empty OLE fixture was found.");
        var testPath = Path.Combine(artifactRoot, "VisualTeX-MathType7-Direct-Set-Real-OLE.docx");
        File.Copy(sourcePath, testPath, overwrite: true);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.OLEFormat? format = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(testPath, ReadOnly: false, Visible: false);
            document.Activate();
            AssertEqual(1, document.InlineShapes.Count,
                "Direct MathType SetData acceptance expected one real inline OLE equation.");
            shape = document.InlineShapes[1];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                "The real OLE fixture is not recognized as MathType.");
            format = shape.OLEFormat;
            Console.WriteLine("Real MathType OLE ProgID=" + format.ProgID);
            Release(format);
            format = null;

            var before = MathTypeOleInterop.ReadMathMl(shape);
            var beforeLatex = MathMlToLatexConverter.Convert(before).Replace(" ", string.Empty);
            Console.WriteLine("  before=" + beforeLatex);

            const string editedMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>u</mi><mo>+</mo><mn>7</mn></mrow></math>";
            Console.WriteLine("  probing MathML SetData FORMATETC/HGLOBAL/ownership variants on the initialized real OLE...");
            var setProbe = MathTypeOleInterop.ProbeExistingOleMathMlSetVariants(shape, editedMathMl);
            Console.WriteLine("  set-probe=" + setProbe);
            if (setProbe.IndexOf("S_OK", StringComparison.Ordinal) < 0)
                throw new InvalidDataException(
                    "No MathType MathML SetData variant succeeded on the initialized real OLE. " + setProbe);
            var after = MathTypeOleInterop.ReadMathMl(shape);
            var afterLatex = MathMlToLatexConverter.Convert(after).Replace(" ", string.Empty);
            Console.WriteLine("  after=" + afterLatex);
            AssertTrue(afterLatex.IndexOf("u+7", StringComparison.Ordinal) >= 0,
                $"MathType real OLE did not retain direct SetData. LaTeX='{afterLatex}'.");
            format = shape.OLEFormat;
            AssertEqual("Equation.DSMT4", format.ProgID,
                "Direct SetData changed the real MathType OLE class.");
            Release(format);
            format = null;

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(testPath, ReadOnly: false, Visible: false);
            shape = document.InlineShapes[1];
            var reopened = MathTypeOleInterop.ReadMathMl(shape);
            var reopenedLatex = MathMlToLatexConverter.Convert(reopened).Replace(" ", string.Empty);
            Console.WriteLine("  reopened=" + reopenedLatex);
            AssertTrue(reopenedLatex.IndexOf("u+7", StringComparison.Ordinal) >= 0,
                $"Word save/reopen lost MathType direct SetData. LaTeX='{reopenedLatex}'.");
            Console.WriteLine("Real initialized MathType OLE direct SetData/save/reopen passed.");
        }
        finally
        {
            Release(format);
            Release(shape);
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

    private static void RunWordMathTypeOleInteropAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.InlineShape? shape = null;
        Word.OLEFormat? format = null;
        var path = Path.Combine(artifactRoot, "VisualTeX-MathType7-OLE-Interop.doc");
        if (File.Exists(path)) File.Delete(path);
        const string initialMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mi>c</mi></mfrac></math>";
        const string editedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow><mi>y</mi></mfrac></math>";
        try
        {
            DismissMathTypeOleCloseDialogByCancel();
            CloseAcceptanceCreatedBlankMathTypeIfSafe();
            if (!MathTypeOleInterop.TryResolveCapabilities("Equation.DSMT4", out var capabilities))
                throw new InvalidOperationException("Installed MathType 7 Equation.DSMT4 OLE server was not detected.");
            AssertTrue(capabilities.RegisteredMathMlGetSet,
                "Installed MathType 7 OLE server does not register MathML Get/Set.");
            Console.WriteLine(
                $"MathType OLE: ProgID={capabilities.ProgId}, CLSID={capabilities.ResolvedClsid:B}, "
                + $"server={capabilities.ServerPath}, conversionVerb={capabilities.RunForConversionVerb}.");

            application = CreateWordApplication(visible: false);
            var mathTypeSeed = @"C:\Program Files (x86)\MathType\Office Support\BlankEqn.doc";
            if (!File.Exists(mathTypeSeed))
                throw new FileNotFoundException("The installed MathType 7 BlankEqn.doc seed is missing.", mathTypeSeed);
            Console.WriteLine("[MathType OLE 1/5] Loading MathType 7's installed OLE seed after clearing the stale close prompt...");
            File.Copy(mathTypeSeed, path, overwrite: true);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            Console.WriteLine("  word-open complete");
            var seedShapeCount = document.InlineShapes.Count;
            Console.WriteLine($"  inline-shapes-count={seedShapeCount}");
            AssertEqual(1, seedShapeCount,
                "MathType 7 BlankEqn.doc did not contain exactly one OLE equation.");
            shape = document.InlineShapes[1];
            Console.WriteLine("  get-inline-shape complete");
            Console.WriteLine("  detect-mathtype begin");
            var isMathTypeSeed = MathTypeOleInterop.IsMathTypeOle(shape);
            Console.WriteLine($"  detect-mathtype={isMathTypeSeed}");
            AssertTrue(isMathTypeSeed,
                "Word inserted Equation.DSMT4 but VisualTeX did not recognize it as MathType OLE.");

            Console.WriteLine("  read-progid begin");
            format = shape.OLEFormat;
            AssertEqual("Equation.DSMT4", format.ProgID,
                "Inserted MathType OLE has the wrong ProgID.");
            Console.WriteLine($"  read-progid={format.ProgID}");
            Release(format);
            format = null;

            Console.WriteLine("  describe-data-formats begin");
            var seedFormats = MathTypeOleInterop.DescribeDataFormats(shape);
            Console.WriteLine("  describe-data-formats complete: " + seedFormats);
            Console.WriteLine("  initial-read-mathml begin");
            var nativeInitialRead = MathTypeOleInterop.ReadMathMl(shape);
            Console.WriteLine("  initial-read-mathml complete");
            var nativeInitialLatex = MathMlToLatexConverter.Convert(nativeInitialRead).Replace(" ", string.Empty);
            Console.WriteLine("  installed BlankEqn.doc seed LaTeX='" + nativeInitialLatex + "'.");

            Console.WriteLine("  Replacing the native equation through CF_TEXT/TeX SetData...");
            MathTypeOleInterop.WriteTextForExistingOle(shape, @"\frac{a+b}{c}");
            var firstRead = MathTypeOleInterop.ReadMathMl(shape);
            var firstLatex = MathMlToLatexConverter.Convert(firstRead).Replace(" ", string.Empty);
            AssertTrue(firstLatex.IndexOf("a+b", StringComparison.Ordinal) >= 0
                       && firstLatex.IndexOf("c", StringComparison.Ordinal) >= 0,
                $"MathType 7 did not round-trip the initial MathML. LaTeX='{firstLatex}'.");

            Console.WriteLine("[MathType OLE 2/5] Replacing MathML in-place without changing the OLE class...");
            var widthBefore = shape.Width;
            var heightBefore = shape.Height;
            MathTypeOleInterop.WriteTextForExistingOle(shape, @"\frac{x+1}{y}");
            var editedRead = MathTypeOleInterop.ReadMathMl(shape);
            var editedLatex = MathMlToLatexConverter.Convert(editedRead).Replace(" ", string.Empty);
            AssertTrue(editedLatex.IndexOf("x+1", StringComparison.Ordinal) >= 0
                       && editedLatex.IndexOf("y", StringComparison.Ordinal) >= 0,
                $"MathType 7 did not expose the edited MathML. LaTeX='{editedLatex}'.");
            format = shape.OLEFormat;
            AssertEqual("Equation.DSMT4", format.ProgID,
                "In-place MathML update changed MathType OLE into another object class.");
            Release(format);
            format = null;
            Console.WriteLine(
                $"  MathType redraw extent: {widthBefore:0.##}x{heightBefore:0.##} -> {shape.Width:0.##}x{shape.Height:0.##} pt.");

            Console.WriteLine("[MathType OLE 3/5] Saving and reopening the Word document...");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            AssertEqual(1, document.InlineShapes.Count,
                "Save/reopen changed the MathType OLE object count.");
            Release(shape);
            shape = document.InlineShapes[1];
            format = shape.OLEFormat;
            AssertEqual("Equation.DSMT4", format.ProgID,
                "Save/reopen lost the MathType 7 OLE ProgID.");
            Release(format);
            format = null;
            var reopenedRead = MathTypeOleInterop.ReadMathMl(shape);
            var reopenedLatex = MathMlToLatexConverter.Convert(reopenedRead).Replace(" ", string.Empty);
            AssertTrue(reopenedLatex.IndexOf("x+1", StringComparison.Ordinal) >= 0
                       && reopenedLatex.IndexOf("y", StringComparison.Ordinal) >= 0,
                $"Save/reopen lost the MathType-edited formula. LaTeX='{reopenedLatex}'.");

            Console.WriteLine("[MathType OLE 4/5] Editing the VisualTeX-written object through the installed MathType 7 UI...");
            format = shape.OLEFormat;
            InvokeWhileDrivingMathTypeEditor(
                application,
                "{END}{+}z",
                () =>
                {
                    object editVerb = (int)Word.WdOLEVerb.wdOLEVerbOpen;
                    format.DoVerb(ref editVerb);
                    return true;
                });
            Release(format);
            format = null;
            var afterNativeOpen = MathTypeOleInterop.ReadMathMl(shape);
            var afterNativeLatex = MathMlToLatexConverter.Convert(afterNativeOpen).Replace(" ", string.Empty);
            AssertTrue(afterNativeLatex.IndexOf("x+1", StringComparison.Ordinal) >= 0
                       && afterNativeLatex.IndexOf("y", StringComparison.Ordinal) >= 0
                       && afterNativeLatex.IndexOf("z", StringComparison.Ordinal) >= 0,
                $"MathType 7 native editor did not continue editing the VisualTeX-written equation. LaTeX='{afterNativeLatex}'.");

            Console.WriteLine("[MathType OLE 5/5] Real MathType 7 OLE GetData/SetData/save/reopen/native-open passed.");
        }
        finally
        {
            Release(format);
            Release(shape);
            Release(insertion);
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

    private static void RunWordMathTypeOleToVisualTeXAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var seedPath = @"C:\Program Files (x86)\MathType\Office Support\BlankEqn.doc";
        if (!File.Exists(seedPath))
            throw new FileNotFoundException("The installed MathType 7 BlankEqn.doc seed is missing.", seedPath);
        var documentPath = Path.Combine(artifactRoot, "VisualTeX-MathType7-To-VisualTeX-OLE.docx");
        var sourceSnapshotPath = Path.Combine(artifactRoot, "VisualTeX-MathType7-Source-OLE.docx");
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"mathtype-to-visualtex-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var svgPath = Path.Combine(assetRoot, "mathtype-to-visualtex.svg");
        var pngPath = Path.Combine(assetRoot, "mathtype-to-visualtex.png");
        string? emfPath = null;
        if (File.Exists(documentPath)) File.Delete(documentPath);
        if (File.Exists(sourceSnapshotPath)) File.Delete(sourceSnapshotPath);

        Word.Application? application = null;
        Word.Document? seedDocument = null;
        Word.Document? document = null;
        Word.InlineShape? seedShape = null;
        Word.InlineShape? source = null;
        Word.InlineShape? converted = null;
        Word.Range? sourceRange = null;
        Word.OLEFormat? format = null;
        try
        {
            application = CreateWordApplication(visible: false);
            seedDocument = application.Documents.Open(seedPath, ReadOnly: true, Visible: false);
            AssertEqual(1, seedDocument.InlineShapes.Count,
                "MathType 7 BlankEqn.doc did not contain one OLE equation.");
            seedShape = seedDocument.InlineShapes[1];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(seedShape),
                "MathType 7 BlankEqn.doc source was not recognized as Equation.DSMT4.");

            var clipboardBackup = TryCaptureClipboard();
            try
            {
                seedShape.Range.Copy();
                document = application.Documents.Add();
                document.Activate();
                application.Selection.TypeText("MathType OLE acceptance: ");
                application.Selection.PasteSpecial(
                    DataType: Word.WdPasteDataType.wdPasteOLEObject);
            }
            finally
            {
                TryRestoreClipboard(clipboardBackup);
            }
            Console.WriteLine(
                $"  modern paste inventory: inline={document.InlineShapes.Count}, floating={document.Shapes.Count}");
            if (document.InlineShapes.Count == 0 && document.Shapes.Count == 1)
            {
                Word.Shape? floating = null;
                try
                {
                    floating = document.Shapes[1];
                    var floatingProgId = floating.OLEFormat.ProgID;
                    AssertEqual("Equation.DSMT4", floatingProgId,
                        "The floating object pasted from MathType 7 is not Equation.DSMT4.");
                    source = floating.ConvertToInlineShape();
                }
                finally { Release(floating); }
            }
            else
            {
                AssertEqual(1, document.InlineShapes.Count,
                    "Copying the real MathType 7 OLE into a modern Word document did not create exactly one OLE equation.");
                source = document.InlineShapes[1];
            }
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(source),
                "MathType OLE lost its Equation.DSMT4 identity when copied into the modern Word document.");
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            File.Copy(documentPath, sourceSnapshotPath, overwrite: true);
            seedDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(seedShape);
            seedShape = null;
            Release(seedDocument);
            seedDocument = null;
            sourceRange = source.Range;
            var sourceObjectId = $"visualtex-word-vsto-range:{sourceRange.Start}:{sourceRange.End}";
            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var latex = @"\frac{x+1}{y}";
            const float width = 180;
            const float height = 72;
            File.WriteAllText(svgPath, CreateSvg(width, height));
            var pngDataUrl = CreatePngDataUrl(latex, width, height);
            var comma = pngDataUrl.IndexOf(',');
            if (comma < 0)
                throw new InvalidDataException("Acceptance PNG data URL is invalid.");
            File.WriteAllBytes(
                pngPath,
                Convert.FromBase64String(pngDataUrl.Substring(comma + 1)));
            emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, width, height);

            var session = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Mode = "edit",
                Host = "word",
                FormulaId = formulaId,
                SourceDocumentId = service.ReadActiveDocumentId(),
                SourceObjectId = sourceObjectId,
                Title = "MathType to VisualTeX OLE acceptance",
                CodeFormat = "latex",
                DisplayMode = "inline",
                ObjectMode = FormulaOleContract.NativeOleMode,
                Numbered = false,
                FontSizePt = 12,
                Lines = new List<FormulaLine>
                {
                    new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
                },
                ExportResult = new OfficeExportDocument
                {
                    Svg = CreateSvg(width, height),
                    SvgBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(CreateSvg(width, height))),
                    PngBase64 = pngDataUrl,
                    Width = width,
                    Height = height,
                    Baseline = 54,
                },
            };

            Console.WriteLine("[MathType->VisualTeX 1/3] Replacing the real Equation.DSMT4 source only after VisualTeX OLE initialization...");
            service.ReplaceOle(session, pngPath, emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "MathType-to-VisualTeX conversion lost or duplicated the OLE object.");
            Release(sourceRange);
            sourceRange = null;
            Release(source);
            source = null;
            converted = document.InlineShapes[1];
            format = converted.OLEFormat;
            AssertEqual(FormulaOleContract.ProgId, format.ProgID,
                "MathType-to-VisualTeX conversion left the old MathType OLE class behind.");
            Release(format);
            format = null;
            var metadata = WordFormulaMetadataReader.TryRead(converted)
                ?? throw new InvalidDataException(
                    "Converted VisualTeX OLE does not contain editable VisualTeX metadata.");
            AssertEqual(formulaId, metadata.FormulaId,
                "Converted VisualTeX OLE did not promote the session FormulaId.");
            AssertEqual(latex, metadata.Latex,
                "Converted VisualTeX OLE contains the wrong LaTeX source.");

            Console.WriteLine("[MathType->VisualTeX 2/3] Saving and reopening the modern Word document...");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(documentPath, ReadOnly: false, Visible: false);
            AssertEqual(1, document.InlineShapes.Count,
                "Save/reopen changed the converted OLE count.");
            Release(converted);
            converted = document.InlineShapes[1];
            format = converted.OLEFormat;
            AssertEqual(FormulaOleContract.ProgId, format.ProgID,
                "Save/reopen reverted the converted object to MathType OLE.");
            Release(format);
            format = null;
            var reopenedMetadata = WordFormulaMetadataReader.TryRead(converted)
                ?? throw new InvalidDataException(
                    "Save/reopen lost VisualTeX metadata after MathType conversion.");
            AssertEqual(formulaId, reopenedMetadata.FormulaId,
                "Save/reopen changed the converted formula identity.");

            Console.WriteLine("[MathType->VisualTeX 3/3] Real MathType 7 OLE -> VisualTeX OLE conversion passed without activating or modifying the MathType server.");
        }
        finally
        {
            Release(format);
            Release(sourceRange);
            Release(converted);
            Release(source);
            Release(seedShape);
            if (seedDocument is not null)
            {
                try { seedDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(seedDocument);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            if (emfPath is not null)
            {
                try { File.Delete(emfPath); } catch { }
            }
            try { Directory.Delete(assetRoot, recursive: true); } catch { }
            ForceComCleanup();
        }
    }

    private static void CloseAcceptanceCreatedBlankMathTypeIfSafe()
    {
        foreach (var process in System.Diagnostics.Process.GetProcessesByName("MathType"))
        {
            try
            {
                process.Refresh();
                if (!string.Equals(
                        process.MainWindowTitle,
                        "MathType - 无标题 1",
                        StringComparison.Ordinal)
                    || process.MainWindowHandle == IntPtr.Zero)
                    continue;
                var processId = process.Id;
                Console.WriteLine(
                    $"Closing acceptance-created blank MathType window PID={processId} before OLE server validation.");
                PostMessageForMathType(
                    process.MainWindowHandle,
                    0x0010,
                    IntPtr.Zero,
                    IntPtr.Zero); // WM_CLOSE
                var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(3);
                while (DateTime.UtcNow < deadline)
                {
                    Thread.Sleep(100);
                    try
                    {
                        using var current = System.Diagnostics.Process.GetProcessById(processId);
                        if (current.HasExited) return;
                    }
                    catch (ArgumentException)
                    {
                        return;
                    }
                }

                // This exact Untitled 1 window was created by the acceptance's
                // MathType.exe -new probe, and that probe timed out before it ever
                // sent input keys. Discard only this known-empty test document.
                if (DiscardAcceptanceCreatedBlankMathTypeClosePrompt())
                {
                    var discardDeadline = DateTime.UtcNow + TimeSpan.FromSeconds(3);
                    while (DateTime.UtcNow < discardDeadline)
                    {
                        Thread.Sleep(100);
                        try
                        {
                            using var current = System.Diagnostics.Process.GetProcessById(processId);
                            if (current.HasExited) return;
                        }
                        catch (ArgumentException)
                        {
                            return;
                        }
                    }
                }
                DismissMathTypeOleCloseDialogByCancel();
                throw new InvalidOperationException(
                    "The acceptance-created blank MathType window could not be closed safely; its close was cancelled.");
            }
            finally { process.Dispose(); }
        }
    }

    private static bool DiscardAcceptanceCreatedBlankMathTypeClosePrompt()
    {
        var discarded = false;
        EnumWindowsForMathType((window, _) =>
        {
            GetWindowThreadProcessId(window, out var processId);
            if (processId == 0) return true;
            try
            {
                using var process = System.Diagnostics.Process.GetProcessById((int)processId);
                if (!string.Equals(process.ProcessName, "MathType", StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch { return true; }
            var className = new System.Text.StringBuilder(64);
            if (GetClassNameForMathType(window, className, className.Capacity) <= 0
                || !string.Equals(className.ToString(), "#32770", StringComparison.Ordinal))
                return true;
            var noButton = GetDlgItemForMathType(window, 7); // IDNO
            if (noButton == IntPtr.Zero) return true;
            SendMessageForMathType(noButton, 0x00F5, IntPtr.Zero, IntPtr.Zero); // BM_CLICK
            discarded = true;
            return false;
        }, IntPtr.Zero);
        if (discarded)
        {
            Console.WriteLine("Discarded only the acceptance-created empty MathType Untitled 1 document.");
            Thread.Sleep(300);
        }
        return discarded;
    }

    private static void DismissMathTypeOleCloseDialogByCancel()
    {
        var cancelled = false;
        EnumWindowsForMathType((window, _) =>
        {
            GetWindowThreadProcessId(window, out var processId);
            if (processId == 0) return true;
            try
            {
                using var process = System.Diagnostics.Process.GetProcessById((int)processId);
                if (!string.Equals(process.ProcessName, "MathType", StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch { return true; }

            var className = new System.Text.StringBuilder(64);
            if (GetClassNameForMathType(window, className, className.Capacity) <= 0
                || !string.Equals(className.ToString(), "#32770", StringComparison.Ordinal))
                return true;
            var dialogTitle = GetWindowTitleForMathType(window);
            if (dialogTitle.IndexOf("MathType OLE", StringComparison.OrdinalIgnoreCase) < 0)
                return true;
            var cancelButton = GetDlgItemForMathType(window, 2); // IDCANCEL
            if (cancelButton == IntPtr.Zero) return true;
            SendMessageForMathType(cancelButton, 0x00F5, IntPtr.Zero, IntPtr.Zero); // BM_CLICK
            cancelled = true;
            return false;
        }, IntPtr.Zero);
        if (cancelled)
        {
            Console.WriteLine("Cancelled the stale MathType OLE close prompt without saving or discarding its equation.");
            Thread.Sleep(400);
        }
    }

    private static string WaitForMathTypeTitleChange(
        System.Diagnostics.Process process,
        string originalTitle,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            process.Refresh();
            if (process.HasExited)
                throw new InvalidOperationException(
                    "MathType exited while VisualTeX was waiting for the temporary equation document.");
            var title = process.MainWindowTitle ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(title)
                && !string.Equals(title, originalTitle, StringComparison.Ordinal))
                return title;
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(100);
        }
        throw new TimeoutException(
            $"MathType did not switch away from its original document '{originalTitle}' after -new.");
    }

    private static List<(IntPtr Handle, string Title)> GetMathTypeEquationWindows(int processId)
    {
        var result = new List<(IntPtr Handle, string Title)>();
        EnumWindowsForMathType((window, _) =>
        {
            GetWindowThreadProcessId(window, out var ownerPid);
            if (ownerPid != (uint)processId) return true;
            var className = new StringBuilder(64);
            if (GetClassNameForMathType(window, className, className.Capacity) <= 0
                || !string.Equals(className.ToString(), "EQNWINCLASS", StringComparison.Ordinal))
                return true;
            var title = GetWindowTitleForMathType(window);
            if (!string.IsNullOrWhiteSpace(title))
                result.Add((window, title));
            return true;
        }, IntPtr.Zero);
        return result;
    }

    private static (IntPtr Handle, string Title) WaitForTemporaryMathTypeEquationWindow(
        int processId,
        HashSet<IntPtr> existingHandles,
        IntPtr originalWindow,
        string originalTitle,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            DismissMathTypeTooManyWindowsWarning(processId);
            foreach (var candidate in GetMathTypeEquationWindows(processId))
            {
                // MathType 7 is inconsistent here: some -new calls create a new
                // EQNWINCLASS HWND, while others reuse the existing top-level
                // HWND and only switch its active document/title. Accept either
                // behavior, but only when the resulting document is an Untitled
                // acceptance page rather than the user's original equation.
                if (!existingHandles.Contains(candidate.Handle)
                    && IsAcceptanceUntitledMathTypeWindow(candidate.Title))
                    return candidate;
                if (candidate.Handle == originalWindow
                    && !string.Equals(candidate.Title, originalTitle, StringComparison.Ordinal)
                    && IsAcceptanceUntitledMathTypeWindow(candidate.Title))
                    return candidate;
            }
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(80);
        }
        throw new TimeoutException(
            "MathType did not expose a temporary Untitled equation document after -new.");
    }

    private static void CloseTemporaryMathTypeWindow(
        int processId,
        IntPtr temporaryWindow,
        string temporaryTitle,
        IntPtr originalWindow,
        string originalTitle,
        TimeSpan timeout)
    {
        if (temporaryWindow == IntPtr.Zero)
            throw new InvalidOperationException(
                "Refusing to close MathType because the temporary window is missing.");
        if (!IsWindowForMathType(temporaryWindow))
        {
            AssertOriginalMathTypeWindowPreserved(originalWindow, originalTitle);
            return;
        }
        var currentTitle = GetWindowTitleForMathType(temporaryWindow);
        if (!string.Equals(currentTitle, temporaryTitle, StringComparison.Ordinal)
            || !IsAcceptanceUntitledMathTypeWindow(currentTitle))
        {
            throw new InvalidOperationException(
                $"Refusing to close MathType window '{currentTitle}' because it is not the exact recorded temporary Untitled window '{temporaryTitle}'.");
        }

        PostMessageForMathType(
            temporaryWindow,
            0x0010,
            IntPtr.Zero,
            IntPtr.Zero); // WM_CLOSE
        var deadline = DateTime.UtcNow + timeout;
        if (temporaryWindow == originalWindow)
        {
            // When -new reused the user's EQNWINCLASS HWND, closing the temporary
            // document must not destroy that HWND. Wait for its title to switch
            // back to the recorded original document after answering IDNO only
            // for the temporary page's close prompt.
            while (DateTime.UtcNow < deadline)
            {
                if (!IsWindowForMathType(originalWindow))
                    throw new InvalidOperationException(
                        "The user's original MathType window was closed while restoring a reused temporary document.");
                currentTitle = GetWindowTitleForMathType(originalWindow);
                if (string.Equals(currentTitle, originalTitle, StringComparison.Ordinal))
                    return;
                ClickMathTypeDialogButtonOwnedBy(
                    processId,
                    originalWindow,
                    7); // IDNO: discard only the recorded Untitled page
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
            }
            throw new TimeoutException(
                $"MathType reused the original HWND but did not restore '{originalTitle}' after closing '{temporaryTitle}'.");
        }

        while (DateTime.UtcNow < deadline && IsWindowForMathType(temporaryWindow))
        {
            ClickMathTypeDialogButtonOwnedBy(
                processId,
                temporaryWindow,
                7); // IDNO: discard only the temporary conversion document
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(100);
        }
        if (IsWindowForMathType(temporaryWindow))
            throw new TimeoutException(
                $"Temporary MathType conversion window '{temporaryTitle}' did not close.");
        AssertOriginalMathTypeWindowPreserved(originalWindow, originalTitle);
    }

    private static bool ClickMathTypeDialogButtonOwnedBy(
        int processId,
        IntPtr ownerWindow,
        int buttonId)
    {
        var clicked = false;
        EnumWindowsForMathType((window, _) =>
        {
            GetWindowThreadProcessId(window, out var ownerPid);
            if (ownerPid != (uint)processId) return true;
            var className = new StringBuilder(64);
            GetClassNameForMathType(window, className, className.Capacity);
            if (!string.Equals(className.ToString(), "#32770", StringComparison.Ordinal))
                return true;
            if (GetWindowOwnerForMathType(window, 4) != ownerWindow) // GW_OWNER
                return true;
            var button = GetDlgItemForMathType(window, buttonId);
            if (button == IntPtr.Zero) return true;
            SendMessageForMathType(button, 0x00F5, IntPtr.Zero, IntPtr.Zero); // BM_CLICK
            clicked = true;
            return false;
        }, IntPtr.Zero);
        return clicked;
    }

    private static void AssertOriginalMathTypeWindowPreserved(
        IntPtr originalWindow,
        string originalTitle)
    {
        if (originalWindow == IntPtr.Zero || !IsWindowForMathType(originalWindow))
            throw new InvalidOperationException(
                "The user's original MathType equation window was closed during temporary conversion.");
        var currentTitle = GetWindowTitleForMathType(originalWindow);
        if (!string.Equals(currentTitle, originalTitle, StringComparison.Ordinal))
            throw new InvalidOperationException(
                $"The user's original MathType window changed from '{originalTitle}' to '{currentTitle}' during temporary conversion.");
    }

    private static void PutMathMlOnNativeClipboard(string mathMl)
    {
        // Mirror MathType 7's own copy payloads observed on this machine:
        //   MathML / application/mathml+xml => UTF-16LE HGLOBAL
        //   MathML Presentation             => narrow/ANSI HGLOBAL
        // Supplying all three lets MathType choose the same representation it
        // advertises through its OLE/clipboard subsystem.
        var payloads = new[]
        {
            (Name: "MathML", Bytes: Encoding.Unicode.GetBytes(mathMl + "\0")),
            (Name: "MathML Presentation", Bytes: Encoding.ASCII.GetBytes(mathMl + "\0")),
            (Name: "application/mathml+xml", Bytes: Encoding.Unicode.GetBytes(mathMl + "\0")),
        };
        var allocated = new List<(uint Format, IntPtr Memory, bool Transferred)>();
        try
        {
            foreach (var payload in payloads)
            {
                var format = RegisterClipboardFormatForMathType(payload.Name);
                if (format == 0)
                    throw new InvalidOperationException(
                        $"Windows could not register the {payload.Name} clipboard format.");
                var memory = AllocateMathTypeClipboardPayload(payload.Bytes);
                allocated.Add((format, memory, false));
            }

            var opened = false;
            for (var attempt = 0; attempt < 20 && !opened; attempt++)
            {
                opened = OpenClipboardForMathType(IntPtr.Zero);
                if (!opened) Thread.Sleep(25);
            }
            if (!opened)
                throw new InvalidOperationException("Unable to open the Windows clipboard for MathType MathML.");
            try
            {
                if (!EmptyClipboardForMathType())
                    throw new InvalidOperationException("Unable to clear the Windows clipboard for MathType MathML.");
                for (var index = 0; index < allocated.Count; index++)
                {
                    var item = allocated[index];
                    if (SetClipboardDataForMathType(item.Format, item.Memory) == IntPtr.Zero)
                        throw new InvalidOperationException(
                            $"Unable to place MathML clipboard format {item.Format} for MathType.");
                    allocated[index] = (item.Format, item.Memory, true);
                }
            }
            finally { CloseClipboardForMathType(); }
        }
        finally
        {
            foreach (var item in allocated)
            {
                if (!item.Transferred && item.Memory != IntPtr.Zero)
                    GlobalFreeForMathType(item.Memory);
            }
        }
    }

    private static IntPtr AllocateMathTypeClipboardPayload(byte[] payload)
    {
        const uint gmemMoveable = 0x0002;
        const uint gmemZeroInit = 0x0040;
        var global = GlobalAllocForMathType(
            gmemMoveable | gmemZeroInit,
            new UIntPtr((uint)payload.Length));
        if (global == IntPtr.Zero)
            throw new OutOfMemoryException("Unable to allocate the MathType clipboard payload.");
        var locked = GlobalLockForMathType(global);
        if (locked == IntPtr.Zero)
        {
            GlobalFreeForMathType(global);
            throw new InvalidOperationException("Unable to lock the MathType clipboard payload.");
        }
        try { Marshal.Copy(payload, 0, locked, payload.Length); }
        finally { GlobalUnlockForMathType(global); }
        return global;
    }

    private static void CloseTemporaryMathTypeDocumentAndRestore(
        System.Diagnostics.Process process,
        IntPtr mainWindow,
        string originalTitle,
        TimeSpan timeout)
    {
        process.Refresh();
        if (process.HasExited)
            throw new InvalidOperationException("MathType exited before its temporary document could be restored.");
        var currentTitle = process.MainWindowTitle ?? string.Empty;
        if (string.Equals(currentTitle, originalTitle, StringComparison.Ordinal))
            return;
        if (currentTitle.IndexOf("无标题", StringComparison.OrdinalIgnoreCase) < 0
            && currentTitle.IndexOf("Untitled", StringComparison.OrdinalIgnoreCase) < 0)
        {
            throw new InvalidOperationException(
                $"Refusing to close MathType document '{currentTitle}' because it is not the known temporary Untitled document.");
        }

        PostMessageForMathType(mainWindow, 0x0010, IntPtr.Zero, IntPtr.Zero); // WM_CLOSE
        var deadline = DateTime.UtcNow + timeout;
        var noClicked = false;
        while (DateTime.UtcNow < deadline)
        {
            Thread.Sleep(100);
            process.Refresh();
            if (process.HasExited)
                throw new InvalidOperationException(
                    "MathType exited while closing the temporary document instead of restoring the user's original session.");
            var title = process.MainWindowTitle ?? string.Empty;
            if (string.Equals(title, originalTitle, StringComparison.Ordinal))
                return;
            if (!noClicked && ClickMathTypeDialogButton(process.Id, 7)) // IDNO
            {
                noClicked = true;
                Console.WriteLine("  discarded only the temporary MathType Untitled document.");
            }
            System.Windows.Forms.Application.DoEvents();
        }
        throw new TimeoutException(
            $"MathType did not restore its original document '{originalTitle}' after closing the temporary document.");
    }

    private static bool ClickMathTypeDialogButton(int processId, int buttonId)
    {
        var clicked = false;
        EnumWindowsForMathType((window, _) =>
        {
            GetWindowThreadProcessId(window, out var ownerPid);
            if (ownerPid != (uint)processId) return true;
            var className = new StringBuilder(64);
            if (GetClassNameForMathType(window, className, className.Capacity) <= 0
                || !string.Equals(className.ToString(), "#32770", StringComparison.Ordinal))
                return true;
            var button = GetDlgItemForMathType(window, buttonId);
            if (button == IntPtr.Zero) return true;
            SendMessageForMathType(button, 0x00F5, IntPtr.Zero, IntPtr.Zero); // BM_CLICK
            clicked = true;
            return false;
        }, IntPtr.Zero);
        return clicked;
    }

    private static System.Windows.Forms.IDataObject? TryCaptureClipboard()
    {
        try { return System.Windows.Forms.Clipboard.GetDataObject(); }
        catch { return null; }
    }

    private static void TryRestoreClipboard(System.Windows.Forms.IDataObject? backup)
    {
        if (backup is null) return;
        try { System.Windows.Forms.Clipboard.SetDataObject(backup, true); }
        catch { }
    }

    private static IntPtr CopyEquationFromNewMathTypeWindow(string inputKeys)
    {
        var existing = GetMathTypeTopLevelWindows();
        var startInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = @"C:\Program Files (x86)\MathType\MathType.exe",
            Arguments = "-new",
            UseShellExecute = true,
        };
        using var launched = System.Diagnostics.Process.Start(startInfo);
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(12);
        IntPtr target = IntPtr.Zero;
        while (DateTime.UtcNow < deadline)
        {
            foreach (var window in GetMathTypeTopLevelWindows())
            {
                if (!existing.Contains(window))
                {
                    target = window;
                    break;
                }
            }
            if (target != IntPtr.Zero) break;
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(100);
        }
        if (target == IntPtr.Zero)
            throw new TimeoutException(
                "MathType 7 -new did not create a separate test equation window; the existing MathType window was left untouched.");

        SetForegroundWindow(target);
        Thread.Sleep(350);
        System.Windows.Forms.SendKeys.SendWait(inputKeys);
        Thread.Sleep(250);
        System.Windows.Forms.SendKeys.SendWait("^a");
        Thread.Sleep(100);
        System.Windows.Forms.SendKeys.SendWait("^c");
        Thread.Sleep(500);
        return target;
    }

    private static HashSet<IntPtr> GetMathTypeTopLevelWindows()
    {
        var result = new HashSet<IntPtr>();
        EnumWindowsForMathType((window, _) =>
        {
            if (window == IntPtr.Zero || !IsWindowVisibleForMathType(window)) return true;
            GetWindowThreadProcessId(window, out var processId);
            if (processId == 0) return true;
            try
            {
                using var process = System.Diagnostics.Process.GetProcessById((int)processId);
                if (string.Equals(process.ProcessName, "MathType", StringComparison.OrdinalIgnoreCase))
                    result.Add(window);
            }
            catch { }
            return true;
        }, IntPtr.Zero);
        return result;
    }

    private static void CloseOnlyMathTypeTestWindow(IntPtr testWindow)
    {
        if (testWindow == IntPtr.Zero) return;
        try
        {
            PostMessageForMathType(testWindow, 0x0010, IntPtr.Zero, IntPtr.Zero); // WM_CLOSE
            var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(2);
            while (DateTime.UtcNow < deadline)
            {
                if (!IsWindowForMathType(testWindow)) return;
                var dialog = IntPtr.Zero;
                EnumWindowsForMathType((window, _) =>
                {
                    if (GetWindowOwnerForMathType(window, 4) == testWindow) // GW_OWNER
                    {
                        dialog = window;
                        return false;
                    }
                    return true;
                }, IntPtr.Zero);
                if (dialog != IntPtr.Zero)
                {
                    var noButton = GetDlgItemForMathType(dialog, 7); // IDNO
                    if (noButton != IntPtr.Zero)
                        SendMessageForMathType(noButton, 0x00F5, IntPtr.Zero, IntPtr.Zero); // BM_CLICK
                }
                Thread.Sleep(100);
            }
        }
        catch { }
    }

    private static T InvokeWhileDrivingMathTypeEditor<T>(
        Word.Application application,
        string inputKeys,
        Func<T> invokeEditor)
    {
        Word.Window? wordWindow = null;
        Exception? inputError = null;
        using var inputCompleted = new ManualResetEventSlim(false);
        try
        {
            wordWindow = application.ActiveWindow;
            var wordWindowHandle = new IntPtr(wordWindow.Hwnd);
            SetForegroundWindow(wordWindowHandle);
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(250);

            // MathType's OLE create/open calls are synchronous from Word. Drive
            // only the external editor window from a helper STA thread while the
            // Word COM call is blocked, then let the OLE call return naturally.
            var inputThread = new Thread(() =>
            {
                try
                {
                    var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(15);
                    IntPtr mathTypeWindow = IntPtr.Zero;
                    while (DateTime.UtcNow < deadline)
                    {
                        var candidate = GetForegroundWindowForAcceptance();
                        if (candidate != IntPtr.Zero && candidate != wordWindowHandle)
                        {
                            GetWindowThreadProcessId(candidate, out var processId);
                            if (processId != 0)
                            {
                                try
                                {
                                    using var process = System.Diagnostics.Process.GetProcessById((int)processId);
                                    if (string.Equals(
                                            process.ProcessName,
                                            "MathType",
                                            StringComparison.OrdinalIgnoreCase))
                                    {
                                        mathTypeWindow = candidate;
                                        break;
                                    }
                                }
                                catch { }
                            }
                        }
                        Thread.Sleep(80);
                    }
                    if (mathTypeWindow == IntPtr.Zero)
                        throw new TimeoutException("The installed MathType 7 OLE editor did not become the foreground window.");

                    SetForegroundWindow(mathTypeWindow);
                    Thread.Sleep(350);
                    if (!string.IsNullOrEmpty(inputKeys))
                    {
                        System.Windows.Forms.SendKeys.SendWait(inputKeys);
                        Thread.Sleep(300);
                    }
                    System.Windows.Forms.SendKeys.SendWait("%{F4}");
                }
                catch (Exception error)
                {
                    inputError = error;
                }
                finally
                {
                    inputCompleted.Set();
                }
            })
            {
                IsBackground = true,
                Name = "VisualTeX MathType OLE acceptance input",
            };
            inputThread.SetApartmentState(ApartmentState.STA);
            inputThread.Start();

            var result = invokeEditor();
            if (!inputCompleted.Wait(TimeSpan.FromSeconds(4)))
                throw new TimeoutException("MathType 7 editor returned but the acceptance input thread did not complete.");
            if (inputError is not null)
                throw new InvalidOperationException("MathType 7 native editor automation failed.", inputError);

            SetForegroundWindow(wordWindowHandle);
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(700);
            return result;
        }
        finally
        {
            Release(wordWindow);
        }
    }

    private delegate bool EnumMathTypeWindowsProc(IntPtr window, IntPtr parameter);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "EnumWindows")]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    private static extern bool EnumWindowsForMathType(
        EnumMathTypeWindowsProc callback,
        IntPtr parameter);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "EnumChildWindows")]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    private static extern bool EnumChildWindowsForMathType(
        IntPtr parent,
        EnumMathTypeWindowsProc callback,
        IntPtr parameter);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetWindowTextW", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    private static extern int GetWindowTextForMathType(
        IntPtr window,
        StringBuilder text,
        int maxCount);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "IsWindowVisible")]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    private static extern bool IsWindowVisibleForMathType(IntPtr window);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "IsWindow")]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    private static extern bool IsWindowForMathType(IntPtr window);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "PostMessageW")]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    private static extern bool PostMessageForMathType(
        IntPtr window,
        uint message,
        IntPtr wParam,
        IntPtr lParam);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetWindow")]
    private static extern IntPtr GetWindowOwnerForMathType(IntPtr window, uint command);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetDlgItem")]
    private static extern IntPtr GetDlgItemForMathType(IntPtr dialog, int controlId);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    private static extern int GetClassNameForMathType(
        IntPtr window,
        System.Text.StringBuilder className,
        int maxCount);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SendMessageW")]
    private static extern IntPtr SendMessageForMathType(
        IntPtr window,
        uint message,
        IntPtr wParam,
        IntPtr lParam);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "RegisterClipboardFormatW", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    private static extern uint RegisterClipboardFormatForMathType(string format);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "OpenClipboard")]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    private static extern bool OpenClipboardForMathType(IntPtr ownerWindow);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "EmptyClipboard")]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    private static extern bool EmptyClipboardForMathType();

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetClipboardData")]
    private static extern IntPtr SetClipboardDataForMathType(uint format, IntPtr memory);

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "CloseClipboard")]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    private static extern bool CloseClipboardForMathType();

    [System.Runtime.InteropServices.DllImport("kernel32.dll", EntryPoint = "GlobalAlloc")]
    private static extern IntPtr GlobalAllocForMathType(uint flags, UIntPtr bytes);

    [System.Runtime.InteropServices.DllImport("kernel32.dll", EntryPoint = "GlobalLock")]
    private static extern IntPtr GlobalLockForMathType(IntPtr memory);

    [System.Runtime.InteropServices.DllImport("kernel32.dll", EntryPoint = "GlobalUnlock")]
    [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
    private static extern bool GlobalUnlockForMathType(IntPtr memory);

    [System.Runtime.InteropServices.DllImport("kernel32.dll", EntryPoint = "GlobalFree")]
    private static extern IntPtr GlobalFreeForMathType(IntPtr memory);

    [System.Runtime.InteropServices.DllImport("ole32.dll", EntryPoint = "OleFlushClipboard")]
    private static extern int OleFlushClipboardForMathType();

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetForegroundWindow")]
    private static extern IntPtr GetForegroundWindowForAcceptance();
}
