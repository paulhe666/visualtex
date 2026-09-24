using System.Windows.Automation;
using Office = Microsoft.Office.Core;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordInstalledVisualTeXNumberedSequenceUiAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var previousAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        Word.Application? application = null;
        Word.Document? document = null;
        Office.COMAddIns? addIns = null;
        Office.COMAddIn? installedAddIn = null;
        var ownsDocument = true;
        var useActiveDocument = AttachActiveWord && string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_UI_E2E_USE_ACTIVE_DOCUMENT"),
            "1",
            StringComparison.OrdinalIgnoreCase);
        using var codeBaseOverride = AttachActiveWord
            ? null
            : CreateUiE2eWordVstoCodeBaseOverride();
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            application = CreateWordApplication(visible: true);
            application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            if (useActiveDocument)
            {
                document = application.ActiveDocument
                    ?? throw new InvalidOperationException("Active Word has no document for VisualTeX numbered UI acceptance.");
                ownsDocument = false;
                Console.WriteLine($"[FULL UI VISUALTEX NUMBERING] Attaching to '{document.Name}'.");
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

            var wordWindow = new IntPtr(application.ActiveWindow.Hwnd);
            var wordRoot = AutomationElement.FromHandle(wordWindow);
            var initialShapes = document.InlineShapes.Count;
            var suffixes = new List<int>();
            for (var insertionIndex = 1; insertionIndex <= 3; insertionIndex++)
            {
                var insertionPosition = Math.Max(document.Content.Start, document.Content.End - 1);
                application.Selection.SetRange(insertionPosition, insertionPosition);
                PumpUi(TimeSpan.FromMilliseconds(350));
                ClickWordRibbonControl(
                    wordWindow,
                    wordRoot,
                    "OLE 行间公式",
                    $"Word Ribbon VisualTeX numbered display #{insertionIndex}");

                var editorWindow = WaitForVisibleOfficeEditorWindow(TimeSpan.FromSeconds(25));
                SetForegroundWindow(editorWindow);
                PumpUi(TimeSpan.FromMilliseconds(700));
                var editorRoot = AutomationElement.FromHandle(editorWindow);
                var objectMode = WaitForUiElement(
                    editorRoot,
                    "公式对象格式",
                    ControlType.ComboBox,
                    TimeSpan.FromSeconds(15));
                SelectComboBoxOption(
                    editorRoot,
                    objectMode,
                    "VisualTeX OLE",
                    "保存为 VisualTeX OLE");
                var numberCheckbox = WaitForUiElement(
                    editorRoot,
                    "编号",
                    ControlType.CheckBox,
                    TimeSpan.FromSeconds(15));
                EnsureCheckboxCheckedByMouse(numberCheckbox, "编号");
                var pythagorean = WaitForUiElement(
                    editorRoot,
                    "勾股定理",
                    ControlType.Button,
                    TimeSpan.FromSeconds(15));
                ClickUiElementByMouse(pythagorean, "常用公式 勾股定理");
                PumpUi(TimeSpan.FromMilliseconds(500));
                var finish = WaitForUiElement(
                    editorRoot,
                    "完成并插入",
                    ControlType.Button,
                    TimeSpan.FromSeconds(15),
                    requireEnabled: true);
                ClickUiElementByMouse(finish, "完成并插入");
                WaitForOfficeEditorHidden(TimeSpan.FromSeconds(60));
                PumpUi(TimeSpan.FromMilliseconds(500));

                AssertEqual(initialShapes + insertionIndex, document.InlineShapes.Count,
                    $"Real Ribbon insertion #{insertionIndex} did not create exactly one new VisualTeX OLE.");
                Word.InlineShape? shape = null;
                Word.Range? visible = null;
                try
                {
                    shape = document.InlineShapes[document.InlineShapes.Count];
                    AssertTrue(WordFormulaMetadataReader.IsNativeOle(shape),
                        $"Real Ribbon insertion #{insertionIndex} is not VisualTeX OLE.");
                    var metadata = WordFormulaMetadataReader.TryRead(shape)
                        ?? throw new InvalidDataException($"Insertion #{insertionIndex} lost VisualTeX metadata.");
                    AssertTrue(metadata.Numbered && string.Equals(metadata.DisplayMode, "block", StringComparison.OrdinalIgnoreCase),
                        $"Insertion #{insertionIndex} is not a numbered display formula.");
                    visible = WordEquationNumbering.FindVisibleEquationNumberRange(document, metadata.FormulaId)
                        ?? throw new InvalidDataException($"Insertion #{insertionIndex} has no visible equation number.");
                    var visibleText = (visible.Text ?? string.Empty).Trim('\t', '\r', '\a', ' ', '(', ')');
                    var lastPart = visibleText.Split(new[] { '.', '-' }, StringSplitOptions.RemoveEmptyEntries).LastOrDefault();
                    if (!int.TryParse(lastPart, out var suffix))
                        throw new InvalidDataException($"Insertion #{insertionIndex} has an unreadable number '{visibleText}'.");
                    suffixes.Add(suffix);
                    Console.WriteLine($"[FULL UI VISUALTEX NUMBERING] #{insertionIndex} formulaId={metadata.FormulaId} visible='{visibleText}' fields={document.Fields.Count} bookmarks={document.Bookmarks.Count}.");
                }
                finally
                {
                    Release(visible);
                    Release(shape);
                }
            }

            AssertEqual("1,2,3", string.Join(",", suffixes),
                "Three consecutive real Ribbon VisualTeX numbered insertions did not remain consecutive.");
            Console.WriteLine("[FULL UI VISUALTEX NUMBERING] Three real Ribbon insertions remained consecutive: " + string.Join(",", suffixes));
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", previousAcceptance);
            Release(installedAddIn);
            Release(addIns);
            if (ownsDocument && document is not null)
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
