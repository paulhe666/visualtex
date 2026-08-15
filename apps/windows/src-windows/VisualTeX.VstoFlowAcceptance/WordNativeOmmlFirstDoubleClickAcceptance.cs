using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using WinForms = System.Windows.Forms;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordNativeOmmlFirstDoubleClickAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var path = Path.Combine(artifactRoot, "word-native-omml-first-double-click.docx");
        var consoleWindow = GetConsoleWindow();

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Window? window = null;
        Word.Range? firstRange = null;
        Word.Range? secondRange = null;
        Word.InlineShape? convertedShape = null;
        Word.Bookmark? bookmark = null;
        Office.COMAddIns? installedAddIns = null;
        Office.COMAddIn? installedAddIn = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = CreateWordApplication(visible: true);
            installedAddIns = application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.WordVsto";
                installedAddIn = installedAddIns.Item(ref addInIndex);
                if (installedAddIn.Connect) installedAddIn.Connect = false;
            }
            catch
            {
                Release(installedAddIn);
                installedAddIn = null;
            }

            document = application.Documents.Add();
            document.Activate();
            application.Selection.TypeText("Coordinate fallback native OMML: ");
            firstRange = InsertUnmanagedNativeOmml(
                application,
                document,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>a</mi><mo>+</mo><mi>b</mi></math>",
                display: false);
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            application.Selection.TypeText("Real first double-click native OMML:");
            application.Selection.TypeParagraph();
            secondRange = InsertUnmanagedNativeOmml(
                application,
                document,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mi>m</mi><mo>+</mo><mi>n</mi></math>",
                display: true);

            bookmark = WordOmmlFormulaStore.FindAtRange(document, firstRange);
            AssertTrue(bookmark is null,
                "Coordinate-fallback native OMML unexpectedly had VisualTeX metadata before adoption.");
            Release(bookmark);
            bookmark = null;
            bookmark = WordOmmlFormulaStore.FindAtRange(document, secondRange);
            AssertTrue(bookmark is null,
                "Real-double-click native OMML unexpectedly had VisualTeX metadata before adoption.");
            Release(bookmark);
            bookmark = null;

            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: true);
            document.Activate();
            application.Visible = true;

            Release(firstRange);
            Release(secondRange);
            firstRange = document.OMaths[1].Range;
            secondRange = document.OMaths[2].Range;
            window = application.ActiveWindow;
            window.Activate();

            // First prove the Office-2021 coordinate fallback itself can adopt an
            // unmanaged native OMath. This is the path used when Word does not
            // raise WindowBeforeDoubleClick reliably for native equations.
            window.GetPoint(
                out var firstLeft,
                out var firstTop,
                out var firstWidth,
                out var firstHeight,
                firstRange);
            AssertTrue(firstWidth > 0 && firstHeight > 0,
                "Word did not return a visible rectangle for the coordinate-fallback OMML fixture.");
            var service = new WordFormulaService(application);
            var coordinateSelection = service.ReadVisualTeXOmmlAtScreenPoint(
                firstLeft + firstWidth / 2,
                firstTop + firstHeight / 2)
                ?? throw new InvalidDataException(
                    "Unmanaged native OMML was not adopted by the coordinate double-click fallback.");
            AssertEqual(FormulaOleContract.WordOmmlMode, coordinateSelection.ObjectMode,
                "Coordinate fallback changed unmanaged native OMML to the wrong object mode.");
            AssertEqual("a+b",
                (coordinateSelection.Metadata?.Latex ?? string.Empty).Replace(" ", string.Empty),
                "Coordinate fallback recovered the wrong native OMML source.");
            bookmark = WordOmmlFormulaStore.FindAtRange(document, firstRange);
            AssertTrue(bookmark is not null,
                "Coordinate fallback did not persist VTOMML identity after adopting native OMML.");
            Release(bookmark);
            bookmark = null;

            // The second equation is still completely unmanaged here. Exercise
            // the actual low-level mouse double-click path so the user-visible
            // behavior is covered end to end.
            bookmark = WordOmmlFormulaStore.FindAtRange(document, secondRange);
            AssertTrue(bookmark is null,
                "The real-double-click OMML fixture was adopted before its first double-click.");
            Release(bookmark);
            bookmark = null;

            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);

            secondRange.Select();
            WinForms.Application.DoEvents();
            Thread.Sleep(200);
            window.GetPoint(
                out var left,
                out var top,
                out var width,
                out var height,
                secondRange);
            AssertTrue(width > 0 && height > 0,
                "Word did not return a visible rectangle for unmanaged native OMML.");

            var wordWindowHandle = new IntPtr(window.Hwnd);
            if (consoleWindow != IntPtr.Zero) ShowWindow(consoleWindow, 0);
            const uint noMoveNoSizeShow = 0x0001 | 0x0002 | 0x0040;
            SetWindowPos(
                wordWindowHandle,
                new IntPtr(-1),
                0,
                0,
                0,
                0,
                noMoveNoSizeShow);
            SetForegroundWindow(wordWindowHandle);
            WinForms.Application.DoEvents();
            Thread.Sleep(250);
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

            var sessionsBefore = SnapshotSessionIds();
            SetCursorPos(left + width / 2, top + height / 2);
            Thread.Sleep(120);
            for (var click = 0; click < 2; click++)
            {
                mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(90);
            }

            var editSessionId = WaitForNewSession(
                sessionsBefore,
                "word",
                TimeSpan.FromSeconds(30));
            var editSession = WaitForUnchangedEditorReady(
                client,
                editSessionId,
                TimeSpan.FromSeconds(15));
            AssertEqual("edit", editSession.Mode,
                "First native OMML double-click did not create an edit Session.");
            AssertEqual(FormulaOleContract.WordOmmlMode, editSession.ObjectMode,
                "First native OMML double-click did not preserve OMML object mode.");
            AssertEqual("m+n",
                (editSession.Lines.FirstOrDefault()?.Latex ?? string.Empty).Replace(" ", string.Empty),
                "First native OMML double-click opened the wrong formula source.");
            var adoptedFormulaId = editSession.FormulaId;
            AssertTrue(!string.IsNullOrWhiteSpace(adoptedFormulaId),
                "First native OMML double-click did not assign a persistent FormulaId.");

            Commit(
                client,
                editSession,
                "block",
                FormulaOleContract.WordOmmlMode,
                "m+n+1",
                numbered: false,
                mathMl:
                    "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
                    + "<mi>m</mi><mo>+</mo><mi>n</mi><mo>+</mo><mn>1</mn></math>");
            var terminal = WaitForTerminal(
                client,
                editSessionId,
                TimeSpan.FromSeconds(45));
            AssertEqual("completed", terminal.Status,
                terminal.Error ?? "First native OMML VisualTeX edit did not complete.");
            client.CloseEditorAsync(editSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            AssertEqual(2, document.OMaths.Count,
                "Editing unmanaged native OMML lost or duplicated an equation.");
            AssertEqual(0, document.InlineShapes.Count,
                "Editing unmanaged native OMML unexpectedly converted it to OLE.");
            Release(secondRange);
            secondRange = document.OMaths[2].Range;
            var editedXml = System.Xml.Linq.XDocument.Parse(secondRange.WordOpenXML);
            System.Xml.Linq.XNamespace mathNs =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var editedText = string.Concat(
                editedXml.Descendants(mathNs + "t").Select(element => element.Value));
            AssertEqual("m+n+1", editedText.Replace(" ", string.Empty),
                "VisualTeX did not write the edited formula back as native OMML.");
            bookmark = WordOmmlFormulaStore.FindAtRange(document, secondRange);
            AssertTrue(bookmark is not null,
                "First native OMML edit did not persist VTOMML identity.");
            AssertTrue(
                WordOmmlFormulaStore.TryGetFormulaId(bookmark, out var persistedFormulaId)
                && string.Equals(
                    persistedFormulaId,
                    adoptedFormulaId,
                    StringComparison.OrdinalIgnoreCase),
                "First native OMML edit persisted the wrong FormulaId.");
            Release(bookmark);
            bookmark = null;

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: true);
            document.Activate();
            AssertEqual(2, document.OMaths.Count,
                "Save/reopen changed the native OMML equation inventory.");
            AssertEqual(0, document.InlineShapes.Count,
                "Save/reopen converted native OMML to OLE.");

            Release(secondRange);
            secondRange = document.OMaths[2].Range;
            Release(window);
            window = application.ActiveWindow;
            secondRange.Select();
            WinForms.Application.DoEvents();
            Thread.Sleep(180);
            window.GetPoint(out left, out top, out width, out height, secondRange);
            wordWindowHandle = new IntPtr(window.Hwnd);
            SetWindowPos(
                wordWindowHandle,
                new IntPtr(-1),
                0,
                0,
                0,
                0,
                noMoveNoSizeShow);
            SetForegroundWindow(wordWindowHandle);
            WinForms.Application.DoEvents();
            Thread.Sleep(250);
            if (GetWindowRect(wordWindowHandle, out wordWindowRectangle))
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
            var reopenSessionsBefore = SnapshotSessionIds();
            SetCursorPos(left + width / 2, top + height / 2);
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
            AssertEqual(adoptedFormulaId, reopenSession.FormulaId,
                "Save/reopen changed the FormulaId adopted from native OMML.");
            AssertEqual("m+n+1",
                (reopenSession.Lines.FirstOrDefault()?.Latex ?? string.Empty).Replace(" ", string.Empty),
                "Second native OMML double-click recovered the wrong edited source.");
            AssertEqual(FormulaOleContract.WordOmmlMode, reopenSession.ObjectMode,
                "Second native OMML double-click changed the formula out of OMML mode.");

            // Exercise the editor's second target-format choice: the same
            // double-click Session may be committed as VisualTeX native OLE
            // instead of remaining Word OMML.
            Commit(
                client,
                reopenSession,
                "block",
                FormulaOleContract.NativeOleMode,
                "m+n+2",
                numbered: false);
            var reopenTerminal = WaitForTerminal(
                client,
                reopenSessionId,
                TimeSpan.FromSeconds(45));
            AssertEqual("completed", reopenTerminal.Status,
                reopenTerminal.Error ?? "Native OMML to VisualTeX OLE edit did not complete.");
            client.CloseEditorAsync(reopenSessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            AssertEqual(1, document.OMaths.Count,
                "Converting the edited native OMML to VisualTeX OLE left the source OMath behind.");
            AssertEqual(1, document.InlineShapes.Count,
                "Converting the edited native OMML did not create exactly one VisualTeX OLE object.");
            convertedShape = document.InlineShapes[1];
            AssertTrue(WordFormulaMetadataReader.IsNativeOle(convertedShape),
                "The target-format edit created a non-VisualTeX OLE object.");
            var convertedMetadata = WordFormulaMetadataReader.TryRead(convertedShape)
                ?? throw new InvalidDataException(
                    "Converted VisualTeX OLE object has no readable formula metadata.");
            AssertEqual(adoptedFormulaId, convertedMetadata.FormulaId,
                "OMML to VisualTeX OLE conversion changed the adopted FormulaId.");
            AssertEqual("m+n+2", convertedMetadata.Latex.Replace(" ", string.Empty),
                "OMML to VisualTeX OLE conversion persisted the wrong edited source.");

            Console.WriteLine(
                "Native OMML first-double-click acceptance passed: unmanaged OMath adopted on first double-click, edited in place as OMML, reopened with the same FormulaId, and the same edit workflow can convert it to VisualTeX OLE.");
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
            Release(bookmark);
            Release(convertedShape);
            Release(firstRange);
            Release(secondRange);
            Release(window);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            Release(installedAddIn);
            Release(installedAddIns);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static Word.Range InsertUnmanagedNativeOmml(
        Word.Application application,
        Word.Document document,
        string mathMl,
        bool display)
    {
        Word.Range? insertion = null;
        Word.Range? equationRange = null;
        try
        {
            insertion = application.Selection.Range.Duplicate;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            equationRange = WordOmmlConverter.Insert(
                application,
                document,
                insertion,
                mathMl,
                display,
                sourceFingerprint: out _,
                includeLeadingTab: false);
            var result = equationRange.Duplicate;
            if (display)
                application.Selection.SetRange(equationRange.End, equationRange.End);
            else
                application.Selection.SetRange(equationRange.End, equationRange.End);
            return result;
        }
        finally
        {
            Release(equationRange);
            Release(insertion);
        }
    }
}
