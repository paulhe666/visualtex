using System.Text;
using System.Windows.Automation;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;
using WinForms = System.Windows.Forms;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordInstalledMathTypeLeftUiE2eAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var previousAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Office.COMAddIns? addIns = null;
        Office.COMAddIn? installedAddIn = null;
        IntPtr editorWindow = IntPtr.Zero;
        var editorCommitted = false;
        var ownsDocument = true;
        var useActiveDocument = AttachActiveWord && string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_UI_E2E_USE_ACTIVE_DOCUMENT"),
            "1",
            StringComparison.OrdinalIgnoreCase);
        using var codeBaseOverride = CreateUiE2eWordVstoCodeBaseOverride();
        try
        {
            // This acceptance is intentionally UI-only for the operation under test.
            // It must exercise the exact installed Word Ribbon -> real editor UI ->
            // Finish and insert path. Never call ThisAddIn.OnInsertDisplay,
            // VisualTeXSessionClient.PatchAsync(committing), or WordFormulaService
            // insertion helpers from this acceptance.
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            var preExistingEditor = FindVisibleOfficeEditorWindow();
            if (preExistingEditor != IntPtr.Zero)
            {
                var existingRoot = AutomationElement.FromHandle(preExistingEditor);
                WriteAutomationTree(
                    existingRoot,
                    Path.Combine(artifactRoot, "pre-existing-editor.uia.txt"));
                var cancel = WaitForUiElement(
                    existingRoot,
                    "取消",
                    ControlType.Button,
                    TimeSpan.FromSeconds(5),
                    requireEnabled: true);
                InvokeUiElement(cancel, "pre-existing failed Office editor Cancel");
                WaitForOfficeEditorHidden(TimeSpan.FromSeconds(10));
                Console.WriteLine(
                    "[FULL UI MATHTYPE LEFT E2E] Closed the pre-existing failed Office editor through its real Cancel button before starting a clean UI run.");
            }

            application = CreateWordApplication(visible: true);
            application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            if (useActiveDocument)
            {
                document = application.ActiveDocument
                    ?? throw new InvalidOperationException("Active Word has no active document for UI E2E.");
                ownsDocument = false;
                WriteWordMathTypeUiE2eState(
                    document,
                    Path.Combine(artifactRoot, "active-document-before-ui.txt"));
                Console.WriteLine(
                    $"[FULL UI MATHTYPE LEFT E2E] Attaching to active document '{document.Name}' at Selection {application.Selection.Start}:{application.Selection.End}.");
            }
            else
            {
                document = application.Documents.Add();
                document.Activate();
            }
            try { application.ActiveWindow.WindowState = Word.WdWindowState.wdWindowStateMaximize; }
            catch { }

            addIns = application.COMAddIns;
            object addInKey = "VisualTeX.WordVsto";
            installedAddIn = addIns.Item(ref addInKey);
            if (!installedAddIn.Connect)
            {
                installedAddIn.Connect = true;
                PumpUi(TimeSpan.FromMilliseconds(800));
            }
            if (!installedAddIn.Connect)
                throw new InvalidOperationException("Installed VisualTeX.WordVsto add-in is not connected.");

            // This is a real physical Ribbon action through Windows UI Automation.
            var wordWindow = new IntPtr(application.ActiveWindow.Hwnd);
            SetForegroundWindow(wordWindow);
            PumpUi(TimeSpan.FromMilliseconds(400));
            var wordRoot = AutomationElement.FromHandle(wordWindow);
            var visualTeXTab = WaitForUiElement(
                wordRoot,
                "VisualTeX",
                ControlType.TabItem,
                TimeSpan.FromSeconds(15));
            ClickUiElementByMouse(visualTeXTab, "Word VisualTeX Ribbon tab");
            PumpUi(TimeSpan.FromMilliseconds(700));
            var displayOleButton = WaitForUiElement(
                wordRoot,
                "OLE 行间公式",
                ControlType.Button,
                TimeSpan.FromSeconds(15));
            ClickUiElementByMouse(displayOleButton, "Word Ribbon OLE 行间公式");

            editorWindow = WaitForVisibleOfficeEditorWindow(TimeSpan.FromSeconds(25));
            SetForegroundWindow(editorWindow);
            PumpUi(TimeSpan.FromMilliseconds(900));
            var editorRoot = AutomationElement.FromHandle(editorWindow);
            WriteAutomationTree(editorRoot, Path.Combine(artifactRoot, "editor-before-interaction.uia.txt"));

            // 1. Save as -> MathType OLE (real HTML select exposed through UIA).
            var objectMode = WaitForUiElement(
                editorRoot,
                "公式对象格式",
                ControlType.ComboBox,
                TimeSpan.FromSeconds(15));
            SelectComboBoxOption(
                editorRoot,
                objectMode,
                "MathType OLE",
                "保存为 MathType OLE");
            PumpUi(TimeSpan.FromMilliseconds(500));

            // 2. Number checkbox -> checked.
            var numberCheckbox = WaitForUiElement(
                editorRoot,
                "编号",
                ControlType.CheckBox,
                TimeSpan.FromSeconds(15));
            EnsureCheckboxCheckedByMouse(numberCheckbox, "编号");
            PumpUi(TimeSpan.FromMilliseconds(400));

            // 3. Number side -> Left.
            var numberSide = WaitForUiElement(
                editorRoot,
                "MathType 公式编号位置",
                ControlType.ComboBox,
                TimeSpan.FromSeconds(15));
            SelectComboBoxOption(
                editorRoot,
                numberSide,
                "左侧",
                "MathType 编号位置左侧");
            PumpUi(TimeSpan.FromMilliseconds(400));

            // 4. Enter an actual formula by clicking the real Common formula tile.
            // This avoids any hidden Session mutation or test-only editor API.
            var pythagorean = WaitForUiElement(
                editorRoot,
                "勾股定理",
                ControlType.Button,
                TimeSpan.FromSeconds(15));
            ClickUiElementByMouse(pythagorean, "常用公式 勾股定理");
            PumpUi(TimeSpan.FromMilliseconds(900));

            // 5. Click the real Finish and insert button.
            var finish = WaitForUiElement(
                editorRoot,
                "完成并插入",
                ControlType.Button,
                TimeSpan.FromSeconds(15),
                requireEnabled: true);
            WriteAutomationTree(editorRoot, Path.Combine(artifactRoot, "editor-before-finish.uia.txt"));
            ClickUiElementByMouse(finish, "完成并插入");

            // A successful production commit closes/hides the editor. A failure keeps
            // it visible and shows the same toast the user sees. Do not treat a mere
            // Session state change as success.
            var deadline = DateTime.UtcNow.AddSeconds(60);
            while (DateTime.UtcNow < deadline)
            {
                PumpUi(TimeSpan.FromMilliseconds(100));
                var visibleEditor = FindVisibleOfficeEditorWindow();
                if (visibleEditor == IntPtr.Zero)
                {
                    editorCommitted = true;
                    break;
                }
                Thread.Sleep(100);
            }

            if (!editorCommitted)
            {
                try
                {
                    var failedRoot = AutomationElement.FromHandle(editorWindow);
                    WriteAutomationTree(
                        failedRoot,
                        Path.Combine(artifactRoot, "editor-after-failed-finish.uia.txt"));
                }
                catch { }
                WriteWordMathTypeUiE2eState(
                    document,
                    Path.Combine(artifactRoot, "word-after-failed-finish.txt"));
                throw new InvalidOperationException(
                    "Real Office editor remained visible after clicking 完成并插入; the production UI commit failed. "
                    + DescribeVisibleEditorText(editorWindow));
            }

            WriteWordMathTypeUiE2eState(
                document,
                Path.Combine(artifactRoot, "word-after-success.txt"));
            AssertEqual(1, document.InlineShapes.Count,
                "UI E2E did not insert exactly one Word inline object.");
            shape = document.InlineShapes[1];
            AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                "UI E2E did not insert a real Equation.DSMT4 MathType OLE.");
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "full UI E2E first left-numbered MathType insertion");
            AssertMathTypeNumberTexts(document, "(0.1)");

            var mathMl = MathTypeOleStorage.ReadMathMl(shape);
            if (string.IsNullOrWhiteSpace(mathMl)
                || mathMl.IndexOf("a", StringComparison.OrdinalIgnoreCase) < 0
                || mathMl.IndexOf("b", StringComparison.OrdinalIgnoreCase) < 0
                || mathMl.IndexOf("c", StringComparison.OrdinalIgnoreCase) < 0
                || mathMl.IndexOf("msup", StringComparison.OrdinalIgnoreCase) < 0)
                throw new InvalidDataException(
                    "UI E2E inserted Equation.DSMT4, but its MathML does not match the clicked Pythagorean formula tile.");

            if (ownsDocument)
            {
                var targetPath = Path.Combine(artifactRoot, "Full-UI-MathType-Left.docx");
                document.SaveAs2(targetPath, Word.WdSaveFormat.wdFormatXMLDocument);
                document.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = application.Documents.Open(targetPath, ReadOnly: false, Visible: false);
                Release(shape);
                shape = document.InlineShapes[1];
                AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                    "Saved/reopened UI E2E MathType object lost Equation.DSMT4.");
                AssertMathTypeDisplayRow(
                    shape,
                    expectedNumberPosition: "left",
                    "saved/reopened full UI E2E first left-numbered MathType insertion");
                AssertMathTypeNumberTexts(document, "(0.1)");
                Console.WriteLine(
                    "[FULL UI MATHTYPE LEFT E2E] Word Ribbon OLE 行间公式 -> editor 保存为 MathType OLE -> 编号 -> 左侧 -> 勾股定理 -> 完成并插入 produced a real left-numbered Equation.DSMT4 and survived save/reopen.");
            }
            else
            {
                WriteWordMathTypeUiE2eState(
                    document,
                    Path.Combine(artifactRoot, "active-document-after-ui.txt"));
                Console.WriteLine(
                    "[FULL UI MATHTYPE LEFT E2E] The exact UI path inserted a real left-numbered Equation.DSMT4 directly into the user's active Word document; the document was left open and unsaved for visual inspection.");
            }
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", previousAcceptance);
            if (!editorCommitted && editorWindow != IntPtr.Zero)
            {
                try { PostMessage(editorWindow, WmClose, UIntPtr.Zero, IntPtr.Zero); }
                catch { }
                PumpUi(TimeSpan.FromMilliseconds(500));
            }
            Release(shape);
            if (ownsDocument && document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            Release(installedAddIn);
            Release(addIns);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void PumpUi(TimeSpan duration)
    {
        var deadline = DateTime.UtcNow + duration;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(20);
        }
    }

    private static AutomationElement WaitForUiElement(
        AutomationElement root,
        string name,
        ControlType type,
        TimeSpan timeout,
        bool requireEnabled = false)
    {
        var condition = new AndCondition(
            new PropertyCondition(AutomationElement.NameProperty, name),
            new PropertyCondition(AutomationElement.ControlTypeProperty, type));
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            try
            {
                var element = root.FindFirst(TreeScope.Descendants, condition);
                if (element is not null && (!requireEnabled || element.Current.IsEnabled))
                    return element;
            }
            catch { }
            PumpUi(TimeSpan.FromMilliseconds(80));
        }
        throw new TimeoutException(
            $"UI Automation could not find enabled={requireEnabled} {type.ProgrammaticName} '{name}'.");
    }

    private static void InvokeUiElement(AutomationElement element, string description)
    {
        if (!element.TryGetCurrentPattern(InvokePattern.Pattern, out var pattern)
            || pattern is not InvokePattern invoke)
            throw new InvalidOperationException(
                $"UI element '{description}' does not expose InvokePattern.");
        invoke.Invoke();
    }

    private static void ClickUiElementByMouse(AutomationElement element, string description)
    {
        var bounds = element.Current.BoundingRectangle;
        if (bounds.Width <= 1 || bounds.Height <= 1)
            throw new InvalidOperationException(
                $"UI element '{description}' has invalid screen bounds {bounds}.");
        try { element.SetFocus(); } catch { }
        var x = (int)Math.Round(bounds.Left + bounds.Width / 2d);
        var y = (int)Math.Round(bounds.Top + bounds.Height / 2d);
        if (!SetCursorPos(x, y))
            throw new InvalidOperationException(
                $"Unable to move the physical mouse cursor to UI element '{description}'.");
        Thread.Sleep(100);
        mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
        mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
        PumpUi(TimeSpan.FromMilliseconds(250));
        Console.WriteLine(
            $"    [FULL UI mouse] clicked '{description}' at screen ({x},{y}).");
    }

    private static void SelectUiItem(AutomationElement element, string description)
    {
        if (!element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var pattern)
            || pattern is not SelectionItemPattern selection)
            throw new InvalidOperationException(
                $"UI item '{description}' does not expose SelectionItemPattern.");
        selection.Select();
    }

    private static void EnsureCheckboxChecked(AutomationElement checkbox, string description)
    {
        if (!checkbox.TryGetCurrentPattern(TogglePattern.Pattern, out var pattern)
            || pattern is not TogglePattern toggle)
            throw new InvalidOperationException(
                $"UI checkbox '{description}' does not expose TogglePattern.");
        if (toggle.Current.ToggleState == ToggleState.Off)
            toggle.Toggle();
        if (toggle.Current.ToggleState != ToggleState.On)
            throw new InvalidOperationException(
                $"UI checkbox '{description}' did not become checked.");
    }

    private static void EnsureCheckboxCheckedByMouse(AutomationElement checkbox, string description)
    {
        if (!checkbox.TryGetCurrentPattern(TogglePattern.Pattern, out var pattern)
            || pattern is not TogglePattern toggle)
            throw new InvalidOperationException(
                $"UI checkbox '{description}' does not expose TogglePattern.");
        if (toggle.Current.ToggleState == ToggleState.Off)
            ClickUiElementByMouse(checkbox, description);
        PumpUi(TimeSpan.FromMilliseconds(150));
        if (!checkbox.TryGetCurrentPattern(TogglePattern.Pattern, out pattern)
            || pattern is not TogglePattern refreshed
            || refreshed.Current.ToggleState != ToggleState.On)
            throw new InvalidOperationException(
                $"UI checkbox '{description}' did not become checked after a physical mouse click.");
    }

    private static void SelectComboBoxOption(
        AutomationElement root,
        AutomationElement combo,
        string optionName,
        string description)
    {
        try
        {
            if (combo.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var expandObject)
                && expandObject is ExpandCollapsePattern expand
                && expand.Current.ExpandCollapseState != ExpandCollapseState.Expanded)
            {
                expand.Expand();
                PumpUi(TimeSpan.FromMilliseconds(250));
            }
        }
        catch { }

        var optionCondition = new AndCondition(
            new PropertyCondition(AutomationElement.NameProperty, optionName),
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
        var deadline = DateTime.UtcNow.AddSeconds(8);
        while (DateTime.UtcNow < deadline)
        {
            AutomationElement? option = null;
            try
            {
                option = combo.FindFirst(TreeScope.Descendants, optionCondition)
                    ?? root.FindFirst(TreeScope.Descendants, optionCondition);
            }
            catch { }
            if (option is not null)
            {
                if (option.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var selectionObject)
                    && selectionObject is SelectionItemPattern selection)
                {
                    selection.Select();
                    PumpUi(TimeSpan.FromMilliseconds(250));
                    return;
                }
                if (option.TryGetCurrentPattern(InvokePattern.Pattern, out var invokeObject)
                    && invokeObject is InvokePattern invoke)
                {
                    invoke.Invoke();
                    PumpUi(TimeSpan.FromMilliseconds(250));
                    return;
                }
            }
            PumpUi(TimeSpan.FromMilliseconds(100));
        }

        // WebView2's UIA provider can expose a native HTML <select> as a ComboBox
        // without exposing its option children. Keyboard navigation is still a real
        // UI action, so use deterministic option order as a final fallback.
        SetForegroundWindow((IntPtr)FindVisibleOfficeEditorWindow());
        combo.SetFocus();
        PumpUi(TimeSpan.FromMilliseconds(100));
        WinForms.SendKeys.SendWait("{HOME}");
        if (string.Equals(optionName, "MathType OLE", StringComparison.Ordinal))
            WinForms.SendKeys.SendWait("{DOWN}");
        WinForms.SendKeys.SendWait("{ENTER}");
        PumpUi(TimeSpan.FromMilliseconds(350));
    }

    private static void WriteAutomationTree(AutomationElement root, string path)
    {
        var lines = new List<string>
        {
            $"ROOT|{root.Current.ControlType.ProgrammaticName}|NAME={root.Current.Name}|ID={root.Current.AutomationId}",
        };
        try
        {
            var elements = root.FindAll(TreeScope.Descendants, Condition.TrueCondition);
            foreach (AutomationElement element in elements)
            {
                var name = element.Current.Name ?? string.Empty;
                var id = element.Current.AutomationId ?? string.Empty;
                if (name.Length == 0 && id.Length == 0) continue;
                lines.Add(
                    $"EL|{element.Current.ControlType.ProgrammaticName}|NAME={name}|ID={id}|ENABLED={element.Current.IsEnabled}");
            }
        }
        catch (Exception error)
        {
            lines.Add("TREE_ERROR|" + error);
        }
        File.WriteAllLines(path, lines, new UTF8Encoding(false));
    }

    private static string DescribeVisibleEditorText(IntPtr editorWindow)
    {
        if (editorWindow == IntPtr.Zero) return "No editor handle.";
        try
        {
            var root = AutomationElement.FromHandle(editorWindow);
            var elements = root.FindAll(TreeScope.Descendants, Condition.TrueCondition);
            var visible = new List<string>();
            foreach (AutomationElement element in elements)
            {
                var name = (element.Current.Name ?? string.Empty).Trim();
                if (name.Length == 0) continue;
                if (name.IndexOf("fail", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("失败", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("HRESULT", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("MathType", StringComparison.OrdinalIgnoreCase) >= 0)
                    visible.Add(name);
            }
            return visible.Count == 0
                ? "No failure text exposed through UI Automation."
                : string.Join(" | ", visible.Distinct());
        }
        catch (Exception error)
        {
            return "Unable to inspect editor UI text: " + error.Message;
        }
    }

    private static void WriteWordMathTypeUiE2eState(Word.Document document, string path)
    {
        var lines = new List<string>
        {
            $"DOC|paragraphs={document.Paragraphs.Count}|fields={document.Fields.Count}|inlineShapes={document.InlineShapes.Count}",
        };
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? result = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        try
        {
            for (var index = 1; index <= document.Fields.Count; index++)
            {
                Release(result); result = null;
                Release(code); code = null;
                Release(field); field = document.Fields[index];
                code = field.Code;
                result = field.Result;
                var nested = 0;
                try { nested = code.Fields.Count; } catch { }
                lines.Add(
                    $"FIELD|{index}|type={(int)field.Type}|code={code.Start}:{code.End}|result={result.Start}:{result.End}|nested={nested}|CODE=[{EscapeUiE2eText(code.Text)}]|RESULT=[{EscapeUiE2eText(result.Text)}]");
            }
            for (var index = 1; index <= document.InlineShapes.Count; index++)
            {
                Release(shapeRange); shapeRange = null;
                Release(shape); shape = document.InlineShapes[index];
                shapeRange = shape.Range;
                var progId = string.Empty;
                try { progId = shape.OLEFormat.ProgID ?? string.Empty; } catch { }
                lines.Add(
                    $"OLE|{index}|range={shapeRange.Start}:{shapeRange.End}|progId={progId}|width={shape.Width:0.###}|height={shape.Height:0.###}");
            }
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            Release(result);
            Release(code);
            Release(field);
        }
        File.WriteAllLines(path, lines, new UTF8Encoding(false));
    }

    private sealed class UiE2eWordVstoCodeBaseOverride : IDisposable
    {
        private readonly string _keyPath;
        private bool _disposed;

        internal UiE2eWordVstoCodeBaseOverride(string keyPath) => _keyPath = keyPath;

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(_keyPath, throwOnMissingSubKey: false);
                Console.WriteLine(
                    "[FULL UI MATHTYPE LEFT E2E] Removed temporary HKCU Word VSTO CodeBase override.");
            }
            catch (Exception error)
            {
                Console.WriteLine(
                    "[FULL UI MATHTYPE LEFT E2E] WARNING: failed to remove temporary HKCU Word VSTO override: "
                    + error.Message);
            }
        }
    }

    private static IDisposable? CreateUiE2eWordVstoCodeBaseOverride()
    {
        var assemblyPath = Environment.GetEnvironmentVariable(
            "VISUALTEX_UI_E2E_WORD_VSTO_CODEBASE");
        if (string.IsNullOrWhiteSpace(assemblyPath)) return null;
        assemblyPath = Path.GetFullPath(assemblyPath!);
        if (!File.Exists(assemblyPath))
            throw new FileNotFoundException(
                "UI E2E Word VSTO CodeBase override assembly does not exist.",
                assemblyPath);

        const string clsid = "{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}";
        var machinePath = $@"Software\Classes\CLSID\{clsid}";
        var userPath = $@"Software\Classes\CLSID\{clsid}";
        using (var existing = Registry.CurrentUser.OpenSubKey(userPath))
        {
            if (existing is not null)
                throw new InvalidOperationException(
                    "UI E2E refuses to overwrite an existing HKCU Word VSTO CLSID registration.");
        }

        try
        {
            using var source = Registry.LocalMachine.OpenSubKey(machinePath)
                ?? throw new InvalidOperationException(
                    "Installed machine-wide VisualTeX.WordVsto CLSID registration is missing.");
            using var destination = Registry.CurrentUser.CreateSubKey(userPath, writable: true)
                ?? throw new InvalidOperationException(
                    "Unable to create temporary HKCU Word VSTO CLSID registration.");
            CopyUiE2eRegistryTree(source, destination);
            var codeBase = new Uri(assemblyPath).AbsoluteUri;
            using var inproc = destination.OpenSubKey("InprocServer32", writable: true)
                ?? throw new InvalidOperationException(
                    "Temporary Word VSTO CLSID clone has no InprocServer32 key.");
            inproc.SetValue("CodeBase", codeBase, RegistryValueKind.String);
            foreach (var versionName in inproc.GetSubKeyNames())
            {
                using var version = inproc.OpenSubKey(versionName, writable: true);
                if (version is null) continue;
                if (version.GetValue("Class") is null) continue;
                version.SetValue("CodeBase", codeBase, RegistryValueKind.String);
            }
            Console.WriteLine(
                $"[FULL UI MATHTYPE LEFT E2E] Temporary HKCU Word VSTO CodeBase -> {codeBase}");
            return new UiE2eWordVstoCodeBaseOverride(userPath);
        }
        catch
        {
            try { Registry.CurrentUser.DeleteSubKeyTree(userPath, throwOnMissingSubKey: false); }
            catch { }
            throw;
        }
    }

    private static void CopyUiE2eRegistryTree(RegistryKey source, RegistryKey destination)
    {
        foreach (var valueName in source.GetValueNames())
        {
            var value = source.GetValue(
                valueName,
                null,
                RegistryValueOptions.DoNotExpandEnvironmentNames);
            if (value is null) continue;
            destination.SetValue(valueName, value, source.GetValueKind(valueName));
        }
        foreach (var subKeyName in source.GetSubKeyNames())
        {
            using var sourceChild = source.OpenSubKey(subKeyName);
            if (sourceChild is null) continue;
            using var destinationChild = destination.CreateSubKey(subKeyName, writable: true);
            if (destinationChild is null) continue;
            CopyUiE2eRegistryTree(sourceChild, destinationChild);
        }
    }

    private static string EscapeUiE2eText(string? text) =>
        (text ?? string.Empty)
            .Replace("\r", "<CR>")
            .Replace("\n", "<LF>")
            .Replace("\t", "<TAB>")
            .Replace("\u0013", "<FIELD_BEGIN>")
            .Replace("\u0014", "<FIELD_SEP>")
            .Replace("\u0015", "<FIELD_END>")
            .Replace("\u0001", "<OLE>");
}
