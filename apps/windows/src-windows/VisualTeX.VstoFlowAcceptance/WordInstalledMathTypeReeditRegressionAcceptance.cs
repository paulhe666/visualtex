using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordInstalledMathTypeReeditRegressionAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var previousAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        var previousTracePath = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        var previousPreferencesPath = Environment.GetEnvironmentVariable("VISUALTEX_OFFICE_PREFERENCES_PATH");
        var previousFormatAcceptance = Environment.GetEnvironmentVariable(
            "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousCreateMode = WordEquationNumbering.GetDefaultCreateObjectMode();
        var previousNumbered = WordEquationNumbering.GetDefaultDisplayEquationNumbered();
        var tracePath = Path.Combine(artifactRoot, "installed-mathtype-reedit.trace.log");
        var preferencesPath = Path.Combine(artifactRoot, "office-preferences-mathtype-reedit.json");
        File.WriteAllText(
            preferencesPath,
            "{\"powerpointDefaultFontSizePt\":20.0,\"mathtypeDoubleClickEditEnabled\":true}");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Window? window = null;
        Word.InlineShape? shape = null;
        Word.Range? range = null;
        Microsoft.Office.Core.COMAddIns? addIns = null;
        Microsoft.Office.Core.COMAddIn? installedAddIn = null;
        object? callbacksObject = null;
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
            Environment.SetEnvironmentVariable("VISUALTEX_OFFICE_PREFERENCES_PATH", preferencesPath);

            var mathTypeBaseline = SnapshotMathTypeProcessIds();
            if (mathTypeBaseline.Count != 0)
                throw new InvalidOperationException(
                    "Installed MathType re-edit acceptance requires MathType.exe process count to be zero before Word starts.");

            application = CreateWordApplication(visible: true);
            addIns = application.COMAddIns;
            object addInKey = "VisualTeX.WordVsto";
            installedAddIn = addIns.Item(ref addInKey);
            if (!installedAddIn.Connect)
                installedAddIn.Connect = true;
            for (var index = 0; index < 80 && installedAddIn.Object is null; index++)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
            }
            callbacksObject = installedAddIn.Object
                ?? throw new InvalidOperationException(
                    "Installed VisualTeX.WordVsto automation object was unavailable for MathType re-edit acceptance.");
            dynamic callbacks = callbacksObject;

            WordEquationNumbering.SetDefaultCreateObjectMode(FormulaOleContract.MathTypeOleMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(false);

            const string latex = @"\lim_{x\to 0}\frac{\sin x}{x}";
            var export = CreateInstalledMathTypeProductExport(
                client,
                latex,
                FormulaOleContract.MathTypeOleMode,
                displayMode: "inline",
                numbered: false);

            document = application.Documents.Add();
            document.Activate();
            CommitInstalledMathTypeInlineFromRibbon(
                client,
                callbacks,
                document,
                latex,
                export,
                1,
                mathTypeBaseline,
                allowTransientMathType: true);

            shape = document.InlineShapes[1];
            AssertInstalledMathTypeFunctionRunExpanded(shape, "before VisualTeX re-edit lim-with-limits");
            var beforeWidth = shape.Width;
            var beforeHeight = shape.Height;
            AssertTrue(beforeWidth > 35f,
                $"The pre-edit inline lim fixture is unexpectedly narrow ({beforeWidth:0.###} pt), so it cannot detect the overlap regression.");
            AssertTrue(beforeHeight > 18f,
                $"The pre-edit inline lim fixture is unexpectedly short ({beforeHeight:0.###} pt), so it cannot detect the overlap regression.");
            WaitForInstalledMathTypeSidecarExit(
                mathTypeBaseline,
                "pre-edit inline MathType fixture");

            var path = Path.Combine(artifactRoot, "Installed-MathType-Reedit-Lim.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false, Visible: true);
            document.Activate();
            Release(shape); shape = document.InlineShapes[1];
            AssertInstalledMathTypeFunctionRunExpanded(shape, "reopened before VisualTeX re-edit inline lim-with-limits");
            beforeWidth = shape.Width;
            beforeHeight = shape.Height;
            var beforeNativeGeometry = SnapshotInstalledMathTypeNativeGeometry(
                shape,
                artifactRoot,
                "before VisualTeX inline MathType re-edit");
            var beforeObjectFormat = SnapshotInstalledMathTypeObjectFormat(
                document,
                shape,
                "before VisualTeX inline MathType re-edit");
            AssertNear(beforeNativeGeometry.WidthPt, beforeWidth, 0.6f,
                "Pre-edit inline OLE width does not match its MathType native WMF.");
            AssertNear(beforeNativeGeometry.HeightPt, beforeHeight, 0.6f,
                "Pre-edit inline OLE height does not match its MathType native WMF.");
            AssertEqual(beforeNativeGeometry.WordPosition, beforeObjectFormat.Position,
                "Pre-edit inline OLE baseline does not match its MathType native WMF.");
            WaitForInstalledMathTypeSidecarExit(
                mathTypeBaseline,
                "pre-edit native WMF verification");

            Release(range); range = shape.Range;
            range.Select();
            Release(window); window = application.ActiveWindow;
            window.Activate();
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(500);
            window.GetPoint(out var left, out var top, out var width, out var height, range);
            if (width <= 0 || height <= 0)
                throw new InvalidDataException(
                    "Word did not expose a visible rectangle for the MathType lim re-edit fixture.");

            var wordWindowHandle = new IntPtr(window.Hwnd);
            const uint noMoveNoSizeShow = 0x0001 | 0x0002 | 0x0040;
            SetWindowPos(wordWindowHandle, new IntPtr(-1), 0, 0, 0, 0, noMoveNoSizeShow);
            SetForegroundWindow(wordWindowHandle);
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(350);
            range.Select();
            window.GetPoint(out left, out top, out width, out height, range);

            // Windows can refuse SetForegroundWindow when the acceptance process
            // has not received recent foreground input. Prime Word with one real
            // click, then wait past the system double-click interval before the
            // actual double-click. This keeps the product hook test real while
            // avoiding a false negative where click #1 only grants foreground.
            SetCursorPos(left + width / 2, top + height / 2);
            mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
            Thread.Sleep(900);
            range.Select();
            window.Activate();
            SetForegroundWindow(wordWindowHandle);
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(250);
            window.GetPoint(out left, out top, out width, out height, range);

            var existingSessions = SnapshotSessionIds();
            SetCursorPos(left + width / 2, top + height / 2);
            Thread.Sleep(120);
            for (var click = 0; click < 2; click++)
            {
                mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
                Thread.Sleep(90);
            }

            var editSessionId = WaitForNewSession(
                existingSessions,
                "word",
                TimeSpan.FromSeconds(30));
            var editSession = WaitForUnchangedEditorReady(
                client,
                editSessionId,
                TimeSpan.FromSeconds(15));
            AssertEqual("edit", editSession.Mode,
                "Real installed MathType double-click did not open an edit Session.");
            AssertEqual(FormulaOleContract.MathTypeOleMode, editSession.ObjectMode,
                "Real installed MathType double-click changed the object mode.");
            var importedLatex = string.Join("\n", editSession.Lines.Select(line => line.Latex))
                .Replace(" ", string.Empty);
            AssertTrue(
                importedLatex.IndexOf("\\lim", StringComparison.Ordinal) >= 0
                && importedLatex.IndexOf("frac", StringComparison.OrdinalIgnoreCase) >= 0
                && importedLatex.IndexOf("sin", StringComparison.OrdinalIgnoreCase) >= 0,
                $"MathType re-edit recovered the wrong formula source: '{importedLatex}'.");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "opening MathType lim in VisualTeX");

            var reeditExport = CreateInstalledMathTypeProductExport(
                client,
                latex,
                FormulaOleContract.MathTypeOleMode,
                displayMode: "inline",
                numbered: false);
            var lineId = editSession.Lines.FirstOrDefault()?.Id
                ?? throw new InvalidDataException("MathType re-edit Session has no formula line.");
            client.PatchAsync(
                    editSessionId,
                    new Dictionary<string, object>
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
                        ["displayMode"] = "inline",
                        ["objectMode"] = FormulaOleContract.MathTypeOleMode,
                        ["numbered"] = false,
                        ["fontSizePt"] = 12d,
                        ["exportWidth"] = reeditExport.Width,
                        ["exportHeight"] = reeditExport.Height,
                        ["exportResult"] = reeditExport,
                        ["dirty"] = true,
                        ["status"] = "committing",
                        ["explicitCancel"] = false,
                    },
                    CancellationToken.None)
                .GetAwaiter().GetResult();

            var deadline = DateTime.UtcNow.AddSeconds(90);
            OfficeSessionDocument? terminal = null;
            while (DateTime.UtcNow < deadline)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
                terminal = client.GetSessionAsync(editSessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
                if (string.Equals(terminal.Status, "failed", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        terminal.Error ?? "Installed MathType re-edit failed while committing.");
                if (string.Equals(terminal.Status, "completed", StringComparison.OrdinalIgnoreCase)
                    && CountMathTypeOleShapes(document) == 1)
                    break;
            }
            if (terminal is null
                || !string.Equals(terminal.Status, "completed", StringComparison.OrdinalIgnoreCase))
                throw new TimeoutException("Installed MathType re-edit did not reach completed state.");
            WaitForInstalledRibbonSessionRelease(editSessionId);
            WaitForInstalledMathTypeSidecarExit(
                mathTypeBaseline,
                "committing VisualTeX inline MathType re-edit");

            Release(shape); shape = document.InlineShapes[1];
            AssertInstalledMathTypeFunctionRunExpanded(shape, "after VisualTeX re-edit inline lim-with-limits");
            var afterWidth = shape.Width;
            var afterHeight = shape.Height;
            var afterNativeGeometry = SnapshotInstalledMathTypeNativeGeometry(
                shape,
                artifactRoot,
                "after VisualTeX inline MathType re-edit");
            var afterObjectFormat = SnapshotInstalledMathTypeObjectFormat(
                document,
                shape,
                "after VisualTeX inline MathType re-edit");
            AssertNear(beforeWidth, afterWidth, 0.05f,
                "VisualTeX changed the inline MathType OLE width while recommitting unchanged MTEF.");
            AssertNear(beforeHeight, afterHeight, 0.05f,
                "VisualTeX changed the inline MathType OLE height while recommitting unchanged MTEF.");
            AssertNear(beforeNativeGeometry.WidthPt, afterNativeGeometry.WidthPt, 0.01f,
                "VisualTeX changed MathType native WMF width while recommitting unchanged MTEF.");
            AssertNear(beforeNativeGeometry.HeightPt, afterNativeGeometry.HeightPt, 0.01f,
                "VisualTeX changed MathType native WMF height while recommitting unchanged MTEF.");
            AssertEqual(beforeNativeGeometry.WordPosition, afterNativeGeometry.WordPosition,
                "VisualTeX changed MathType native WMF WordPosition while recommitting unchanged MTEF.");
            AssertEqual(beforeNativeGeometry.Signature, afterNativeGeometry.Signature,
                "VisualTeX changed Equation Native semantics while recommitting unchanged MTEF.");
            AssertNear(afterNativeGeometry.WidthPt, afterWidth, 0.6f,
                "Re-edited inline OLE width does not match its MathType native WMF.");
            AssertNear(afterNativeGeometry.HeightPt, afterHeight, 0.6f,
                "Re-edited inline OLE height does not match its MathType native WMF.");
            AssertEqual(afterNativeGeometry.WordPosition, afterObjectFormat.Position,
                "Re-edited inline OLE baseline does not match its MathType native WMF.");
            AssertEqual(beforeObjectFormat.Position, afterObjectFormat.Position,
                "VisualTeX changed the inline MathType Word baseline.");
            AssertEqual(beforeObjectFormat.FontName, afterObjectFormat.FontName,
                "VisualTeX changed the inline MathType object-character font.");
            AssertNear(beforeObjectFormat.FontSizePt, afterObjectFormat.FontSizePt, 0.01f,
                "VisualTeX changed the inline MathType object-character font size.");
            WaitForInstalledMathTypeSidecarExit(
                mathTypeBaseline,
                "post-edit native WMF verification");
            var readBackMathMl = MathTypeOleStorage.ReadMathMl(shape);
            var readBackLatex = MathMlToLatexConverter.Convert(readBackMathMl)
                .Replace(" ", string.Empty);
            AssertTrue(
                readBackLatex.IndexOf("\\lim", StringComparison.Ordinal) >= 0
                && readBackLatex.IndexOf("frac", StringComparison.OrdinalIgnoreCase) >= 0
                && readBackLatex.IndexOf("sin", StringComparison.OrdinalIgnoreCase) >= 0,
                $"MathType Equation Native changed semantically after VisualTeX re-edit: '{readBackLatex}'.");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false, Visible: true);
            Release(shape); shape = document.InlineShapes[1];
            AssertInstalledMathTypeFunctionRunExpanded(shape, "reopened after VisualTeX re-edit inline lim-with-limits");
            var reopenedNativeGeometry = SnapshotInstalledMathTypeNativeGeometry(
                shape,
                artifactRoot,
                "reopened after VisualTeX inline MathType re-edit");
            var reopenedObjectFormat = SnapshotInstalledMathTypeObjectFormat(
                document,
                shape,
                "reopened after VisualTeX inline MathType re-edit");
            AssertNear(beforeWidth, shape.Width, 0.05f,
                "Save/reopen changed the re-edited inline MathType OLE width.");
            AssertNear(beforeHeight, shape.Height, 0.05f,
                "Save/reopen changed the re-edited inline MathType OLE height.");
            AssertNear(beforeNativeGeometry.WidthPt, reopenedNativeGeometry.WidthPt, 0.01f,
                "Save/reopen changed the MathType native WMF width.");
            AssertNear(beforeNativeGeometry.HeightPt, reopenedNativeGeometry.HeightPt, 0.01f,
                "Save/reopen changed the MathType native WMF height.");
            AssertEqual(beforeNativeGeometry.WordPosition, reopenedNativeGeometry.WordPosition,
                "Save/reopen changed the MathType native WMF WordPosition.");
            AssertEqual(beforeObjectFormat.Position, reopenedObjectFormat.Position,
                "Save/reopen changed the inline MathType Word baseline.");
            AssertEqual(beforeObjectFormat.FontName, reopenedObjectFormat.FontName,
                "Save/reopen changed the inline MathType object-character font.");
            AssertNear(beforeObjectFormat.FontSizePt, reopenedObjectFormat.FontSizePt, 0.01f,
                "Save/reopen changed the inline MathType object-character font size.");
            WaitForInstalledMathTypeSidecarExit(
                mathTypeBaseline,
                "save/reopen VisualTeX inline MathType re-edit");

            Console.WriteLine(
                $"[MATHTYPE INLINE REEDIT REGRESSION] Real installed add-in double-clicked and recommitted {latex}; "
                + $"OLE before={beforeWidth:0.###}x{beforeHeight:0.###} pt, after={afterWidth:0.###}x{afterHeight:0.###} pt, reopened={shape.Width:0.###}x{shape.Height:0.###} pt; "
                + $"native={beforeNativeGeometry.WidthPt:0.###}x{beforeNativeGeometry.HeightPt:0.###} pt, WordPosition={beforeNativeGeometry.WordPosition}; "
                + $"font={beforeObjectFormat.FontName} {beforeObjectFormat.FontSizePt:0.###} pt; MathTypeProcessCount=0. Artifact={path}");

            // Finish the first-document lifecycle before exercising a second
            // installed Ribbon document. VisualTeX intentionally refuses to
            // commit a Session after Word's active-document identity changes.
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            Release(shape); shape = null;
            Release(range); range = null;
            Release(window); window = null;

            RunInstalledDocument1MathTypeRoundTrip(
                client,
                application,
                callbacks,
                artifactRoot,
                tracePath,
                mathTypeBaseline);
        }
        finally
        {
            WordEquationNumbering.SetDefaultCreateObjectMode(previousCreateMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(previousNumbered);
            Release(range);
            Release(shape);
            Release(window);
            Release(callbacksObject);
            Release(installedAddIn);
            Release(addIns);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            Environment.SetEnvironmentVariable("VISUALTEX_OFFICE_PREFERENCES_PATH", previousPreferencesPath);
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTracePath);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                previousFormatAcceptance);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", previousAcceptance);
        }
    }

    private static void RunInstalledDocument1MathTypeRoundTrip(
        VisualTeXSessionClient client,
        Word.Application application,
        dynamic callbacks,
        string artifactRoot,
        string tracePath,
        IReadOnlyCollection<int> mathTypeBaseline)
    {
        const string latex = @"(a+b)^{n}=\sum_{k=0}^{n}\left( \begin{matrix}n \\ k\end{matrix}\right) a^{n-k}b^{k}";
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.Range? content = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(true);
            WordEquationNumbering.SetDefaultCreateObjectMode(FormulaOleContract.NativeOleMode);

            var sourceExport = CreateInstalledMathTypeProductExport(
                client,
                latex,
                FormulaOleContract.NativeOleMode,
                displayMode: "block",
                numbered: true);
            if (string.IsNullOrWhiteSpace(sourceExport.MathMl))
                throw new InvalidDataException(
                    "The production VisualTeX renderer returned no MathML for the document1 matrix round-trip fixture.");
            var expectedSignature = MathTypeMtefCodec.SemanticSignature(sourceExport.MathMl!);

            CommitInstalledVisualTeXDisplayFromRibbon(
                client,
                callbacks,
                document,
                latex,
                sourceExport,
                expectedVisualTeXCount: 1,
                numbered: true);
            AssertEqual(1, CountInstalledVisualTeXOleShapes(document),
                "Document1 round-trip setup did not create exactly one VisualTeX source OLE.");
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Document1 round-trip setup unexpectedly created MathType before conversion.");

            WordEquationNumbering.SetDefaultCreateObjectMode(FormulaOleContract.MathTypeOleMode);
            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertVisualTeXToMathTypeDocument(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=VisualTeX target=MathType",
                mathTypeBaseline,
                allowTransientMathTypeProcess: true);
            AssertEqual(0, CountInstalledVisualTeXOleShapes(document),
                "Document1 VisualTeX→MathType left the source VisualTeX object behind.");
            AssertEqual(1, CountMathTypeOleShapes(document),
                "Document1 VisualTeX→MathType did not create exactly one MathType object.");
            shape = document.InlineShapes[1];
            var firstMathMl = MathTypeOleStorage.ReadMathMl(shape);
            AssertEqual(
                expectedSignature,
                MathTypeMtefCodec.SemanticSignature(firstMathMl),
                $"Document1 VisualTeX→MathType changed the matrix formula semantics. actual='{firstMathMl}'");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "document1 first VisualTeX→MathType conversion");
            Release(shape); shape = null;
            AssertDocument1RoundTripHasNoLatexBridge(document, "after first VisualTeX→MathType");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "document1 first VisualTeX→MathType");

            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertMathTypeToVisualTeXDocument(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=MathType target=VisualTeX",
                mathTypeBaseline);
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Document1 MathType→VisualTeX left the MathType source object behind.");
            AssertEqual(1, CountInstalledVisualTeXOleShapes(document),
                "Document1 MathType→VisualTeX did not restore exactly one VisualTeX OLE.");
            shape = document.InlineShapes[1];
            var metadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException(
                    "Document1 MathType→VisualTeX target has no readable VisualTeX metadata.");
            var recoveredLatex = (metadata.Latex ?? string.Empty).Trim();
            AssertTrue(
                recoveredLatex.IndexOf(@"\begin{matrix}", StringComparison.Ordinal) >= 0,
                $"Document1 MathType→VisualTeX lost the explicit matrix structure: '{recoveredLatex}'.");
            AssertTrue(
                recoveredLatex.IndexOf(@"\binom", StringComparison.Ordinal) < 0,
                $"Document1 MathType→VisualTeX incorrectly changed the explicit 2x1 matrix into binomial syntax: '{recoveredLatex}'.");
            Release(shape); shape = null;
            AssertDocument1RoundTripHasNoLatexBridge(document, "after MathType→VisualTeX");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "document1 MathType→VisualTeX");

            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertVisualTeXToMathTypeDocument(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=VisualTeX target=MathType",
                mathTypeBaseline,
                allowTransientMathTypeProcess: true);
            AssertEqual(0, CountInstalledVisualTeXOleShapes(document),
                "Document1 second VisualTeX→MathType left the VisualTeX source object behind.");
            AssertEqual(1, CountMathTypeOleShapes(document),
                "Document1 second VisualTeX→MathType did not create exactly one MathType object.");
            shape = document.InlineShapes[1];
            var secondMathMl = MathTypeOleStorage.ReadMathMl(shape);
            AssertEqual(
                expectedSignature,
                MathTypeMtefCodec.SemanticSignature(secondMathMl),
                $"Document1 VT→MT→VT→MT changed formula semantics. actual='{secondMathMl}'");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "document1 second VisualTeX→MathType conversion");
            AssertDocument1RoundTripHasNoLatexBridge(document, "after second VisualTeX→MathType");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "document1 second VisualTeX→MathType");

            var path = Path.Combine(
                artifactRoot,
                "Installed-Document1-VT-MT-VT-MT-RoundTrip.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            Release(shape); shape = null;

            document = application.Documents.Open(
                path,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: true);
            AssertEqual(1, CountMathTypeOleShapes(document),
                "Reopened document1 round-trip artifact lost the final MathType object.");
            AssertEqual(0, CountInstalledVisualTeXOleShapes(document),
                "Reopened document1 round-trip artifact restored a stale VisualTeX source object.");
            shape = document.InlineShapes[1];
            var reopenedMathMl = MathTypeOleStorage.ReadMathMl(shape);
            AssertEqual(
                expectedSignature,
                MathTypeMtefCodec.SemanticSignature(reopenedMathMl),
                $"Save/reopen changed document1 VT→MT→VT→MT semantics. actual='{reopenedMathMl}'");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "reopened document1 VT→MT→VT→MT conversion");
            AssertDocument1RoundTripHasNoLatexBridge(document, "after save/reopen");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "document1 VT→MT→VT→MT save/reopen");

            Console.WriteLine(
                $"[DOCUMENT1 INSTALLED ROUNDTRIP] Real installed Ribbon path passed VT→MT→VT→MT for the explicit parenthesized 2x1 matrix formula; matrix remained matrix, no temporary LaTeX bridge survived, save/reopen preserved one Equation.DSMT4, MathTypeProcessCount=0. Artifact={path}");
        }
        finally
        {
            Release(content);
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }

    private static void WaitForInstalledMathTypeSidecarExit(
        IReadOnlyCollection<int> baseline,
        string context)
    {
        var deadline = DateTime.UtcNow.AddSeconds(20);
        while (DateTime.UtcNow < deadline)
        {
            var started = SnapshotMathTypeProcessIds().Except(baseline).ToArray();
            if (started.Length == 0) return;
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(100);
        }
        var remaining = SnapshotMathTypeProcessIds().Except(baseline).ToArray();
        throw new InvalidOperationException(
            context + ": MathType native-preview sidecar process did not exit: "
            + string.Join(",", remaining));
    }

    private static (float WidthPt, float HeightPt, int WordPosition, string Signature)
        SnapshotInstalledMathTypeNativeGeometry(
            Word.InlineShape shape,
            string artifactRoot,
            string context)
    {
        var compoundFile = MathTypeOleStorage.CaptureCompoundFile(shape);
        var equationNative = MathTypeOleStorage.ReadEquationNative(compoundFile);
        var headerLength = BitConverter.ToUInt16(equationNative, 0);
        var mtefLength = checked((int)BitConverter.ToUInt32(equationNative, 8));
        var mtef = new byte[mtefLength];
        Buffer.BlockCopy(equationNative, headerLength, mtef, 0, mtefLength);
        AssertTrue(
            MathTypeNativePreviewRenderer.TryRender(mtef, artifactRoot, out var preview),
            context + ": MathType native preview renderer was unavailable.");
        using (preview)
        {
            var mathMl = MathTypeMtefCodec.ReadEquationNativeMathMl(equationNative);
            return (
                preview.WidthPt,
                preview.HeightPt,
                preview.WordPosition,
                MathTypeMtefCodec.SemanticSignature(mathMl));
        }
    }

    private static (string FontName, float FontSizePt, int Position)
        SnapshotInstalledMathTypeObjectFormat(
            Word.Document document,
            Word.InlineShape shape,
            string context)
    {
        Word.Range? shapeRange = null;
        Word.Font? font = null;
        try
        {
            shapeRange = shape.Range;
            font = shapeRange.Font;
            var fontSize = font.Size;
            if (fontSize == (float)Word.WdConstants.wdUndefined)
                throw new InvalidOperationException(
                    context + ": MathType object-character font size is undefined.");
            return (
                font.Name ?? string.Empty,
                fontSize,
                ReadInstalledMathTypeObjectCharacterPosition(document, shapeRange, context));
        }
        finally
        {
            Release(font);
            Release(shapeRange);
        }
    }

    private static void AssertDocument1RoundTripHasNoLatexBridge(
        Word.Document document,
        string context)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            var text = content.Text ?? string.Empty;
            AssertTrue(
                text.IndexOf(@"\begin{matrix}", StringComparison.Ordinal) < 0,
                context + ": temporary LaTeX bridge containing \\begin{matrix} remained in the Word body.");
            AssertTrue(
                text.IndexOf("$(a+b)", StringComparison.Ordinal) < 0
                && text.IndexOf("$$(a+b)", StringComparison.Ordinal) < 0,
                context + ": temporary dollar-delimited LaTeX bridge remained in the Word body.");
        }
        finally { Release(content); }
    }
}
