using System.Text;
using System.Windows.Automation;
using Office = Microsoft.Office.Core;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordInstalledInlineOleRibbonReeditAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var previousAcceptance =
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        Word.Application? application = null;
        Word.Document? sourceDocument = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Office.COMAddIns? addIns = null;
        Office.COMAddIn? installedAddIn = null;
        IntPtr editorWindow = IntPtr.Zero;
        using var codeBaseOverride = AttachActiveWord
            ? null
            : CreateUiE2eWordVstoCodeBaseOverride();
        try
        {
            // This regression deliberately follows the production UI path:
            // physical Ribbon clicks -> real Office editor -> physical commit button.
            // It does not invoke Ribbon callbacks, Session endpoints or formula
            // services directly.
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            application = CreateWordApplication(visible: true);
            if (AttachActiveWord)
                sourceDocument = application.ActiveDocument
                    ?? throw new InvalidOperationException(
                        "The user's Word instance has no active document.");

            document = application.Documents.Add();
            document.Activate();
            document.Content.Text = "VisualTeX inline OLE:  trailing text\r";
            application.Selection.SetRange(
                "VisualTeX inline OLE: ".Length,
                "VisualTeX inline OLE: ".Length);
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

            var wordWindow = new IntPtr(application.ActiveWindow.Hwnd);
            var wordRoot = AutomationElement.FromHandle(wordWindow);

            ClickWordRibbonControl(
                wordWindow,
                wordRoot,
                "OLE 行内公式",
                "Word Ribbon OLE 行内公式");

            editorWindow = WaitForVisibleOfficeEditorWindow(TimeSpan.FromSeconds(25));
            var editorRoot = AutomationElement.FromHandle(editorWindow);
            WriteAutomationTree(
                editorRoot,
                Path.Combine(artifactRoot, "create-editor.uia.txt"));
            ClickUiElementByMouse(
                WaitForUiElement(
                    editorRoot,
                    "勾股定理",
                    ControlType.Button,
                    TimeSpan.FromSeconds(15)),
                "常用公式 勾股定理");
            PumpUi(TimeSpan.FromMilliseconds(700));
            ClickUiElementByMouse(
                WaitForUiElement(
                    editorRoot,
                    "完成并插入",
                    ControlType.Button,
                    TimeSpan.FromSeconds(15),
                    requireEnabled: true),
                "完成并插入");
            WaitForOfficeEditorHidden(TimeSpan.FromSeconds(60));
            editorWindow = IntPtr.Zero;

            WriteInstalledInlineOleState(
                document,
                Path.Combine(artifactRoot, "after-create.txt"));
            AssertEqual(
                1,
                document.InlineShapes.Count,
                "The real Ribbon create path did not insert exactly one inline shape.");

            shape = document.InlineShapes[1];
            if (!WordFormulaMetadataReader.IsNativeOle(shape))
                throw new InvalidOperationException(
                    "The real Ribbon create path did not insert a VisualTeX native OLE formula.");
            var originalMetadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidOperationException(
                    "The created VisualTeX inline OLE metadata is unreadable.");
            var originalLatex = originalMetadata.Latex;

            shapeRange = shape.Range.Duplicate;
            shapeRange.Select();
            PumpUi(TimeSpan.FromMilliseconds(500));
            ClickWordRibbonControl(
                wordWindow,
                wordRoot,
                "编辑所选公式",
                "Word Ribbon 编辑所选公式");

            editorWindow = WaitForVisibleOfficeEditorWindow(TimeSpan.FromSeconds(25));
            editorRoot = AutomationElement.FromHandle(editorWindow);
            WriteAutomationTree(
                editorRoot,
                Path.Combine(artifactRoot, "edit-editor.uia.txt"));
            ClickUiElementByMouse(
                WaitForUiElement(
                    editorRoot,
                    "欧拉恒等式",
                    ControlType.Button,
                    TimeSpan.FromSeconds(15)),
                "常用公式 欧拉恒等式");
            PumpUi(TimeSpan.FromMilliseconds(700));
            ClickUiElementByMouse(
                WaitForUiElement(
                    editorRoot,
                    "更新公式",
                    ControlType.Button,
                    TimeSpan.FromSeconds(15),
                    requireEnabled: true),
                "更新公式");
            WaitForOfficeEditorHidden(TimeSpan.FromSeconds(60));
            editorWindow = IntPtr.Zero;

            WriteInstalledInlineOleState(
                document,
                Path.Combine(artifactRoot, "after-edit.txt"));

            AssertEqual(
                1,
                document.InlineShapes.Count,
                "Re-editing one VisualTeX inline OLE through the real Ribbon appended a stale duplicate.");
            Release(shapeRange);
            shapeRange = null;
            Release(shape);
            shape = document.InlineShapes[1];
            var updatedMetadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidOperationException(
                    "The edited VisualTeX inline OLE metadata is unreadable.");
            AssertTrue(
                !string.Equals(originalLatex, updatedMetadata.Latex, StringComparison.Ordinal),
                "The real editor interaction did not change the formula source.");

            Console.WriteLine(
                "[FULL UI INLINE OLE RE-EDIT] Real Word Ribbon create and edit kept one native VisualTeX OLE object.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_ACCEPTANCE",
                previousAcceptance);
            if (editorWindow != IntPtr.Zero)
            {
                try { PostMessage(editorWindow, WmClose, UIntPtr.Zero, IntPtr.Zero); }
                catch { }
                PumpUi(TimeSpan.FromMilliseconds(500));
            }
            Release(shapeRange);
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
                catch { }
            }
            Release(document);
            if (sourceDocument is not null)
            {
                try { sourceDocument.Activate(); }
                catch { }
            }
            Release(sourceDocument);
            Release(installedAddIn);
            Release(addIns);
            try { QuitWordApplicationIfOwned(application); }
            catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void ClickWordRibbonControl(
        IntPtr wordWindow,
        AutomationElement wordRoot,
        string controlName,
        string description)
    {
        SetForegroundWindow(wordWindow);
        PumpUi(TimeSpan.FromMilliseconds(350));
        var visualTeXTab = WaitForUiElement(
            wordRoot,
            "VisualTeX",
            ControlType.TabItem,
            TimeSpan.FromSeconds(15));
        ClickUiElementByMouse(visualTeXTab, "Word VisualTeX Ribbon tab");
        PumpUi(TimeSpan.FromMilliseconds(650));
        var control = WaitForUiElement(
            wordRoot,
            controlName,
            ControlType.Button,
            TimeSpan.FromSeconds(15),
            requireEnabled: true);
        ClickUiElementByMouse(control, description);
    }

    private static void WriteInstalledInlineOleState(
        Word.Document document,
        string path)
    {
        var output = new StringBuilder();
        output.AppendLine($"document={document.Name}");
        output.AppendLine($"inlineShapes={document.InlineShapes.Count}");
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? candidate = null;
            try
            {
                candidate = document.InlineShapes[index];
                var metadata = WordFormulaMetadataReader.TryRead(candidate);
                output.AppendLine(
                    $"#{index} range={candidate.Range.Start}:{candidate.Range.End} "
                    + $"progId={candidate.OLEFormat.ProgID} "
                    + $"formulaId={metadata?.FormulaId ?? "<none>"} "
                    + $"latex={metadata?.Latex ?? "<none>"}");
            }
            catch (Exception error)
            {
                output.AppendLine($"#{index} unreadable={error}");
            }
            finally { Release(candidate); }
        }
        File.WriteAllText(path, output.ToString(), new UTF8Encoding(false));
        Console.Write(output.ToString());
    }
}
