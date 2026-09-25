using System.Text;
using System.Windows.Automation;
using Office = Microsoft.Office.Core;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;
using WinForms = System.Windows.Forms;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordVisualTeXNumberToggleAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(artifactRoot, "visualtex-unnumbered-to-numbered-edit.docx");
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"number-toggle-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var svgPath = Path.Combine(assetRoot, "formula.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"180\" height=\"60\" viewBox=\"0 0 180 60\"><text x=\"4\" y=\"42\" font-size=\"32\">x = 2</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 180, 60);
        var pngDataUrl = CreatePngDataUrl("number-toggle", 180, 60);
        var pngPath = Path.Combine(assetRoot, "formula.png");
        File.WriteAllBytes(
            pngPath,
            Convert.FromBase64String(pngDataUrl.Substring(pngDataUrl.IndexOf(',') + 1)));

        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();

            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            try
            {
                application.Selection.SetRange(insertion.Start, insertion.End);
                var createSession = CreateNumberedPerformanceSession(
                    "create",
                    formulaId,
                    document.FullName,
                    WordRangeReference(insertion.Start, insertion.End),
                    originalMetadata: null,
                    latex: @"x=1");
                createSession.Numbered = false;
                createSession.ExportResult = new OfficeExportDocument
                {
                    Width = 180,
                    Height = 60,
                    Baseline = 45,
                };
                service.InsertOle(createSession, pngPath, emfPath);
            }
            finally { Release(insertion); }

            shape = FindVisualTeXOleByFormulaIdForNumberToggle(document, formulaId);
            var originalMetadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException("The unnumbered VisualTeX source lost metadata.");
            AssertTrue(!originalMetadata.Numbered,
                "The number-toggle fixture was unexpectedly numbered before editing.");
            AssertTrue(
                !document.Bookmarks.Exists(WordEquationNumbering.EquationBookmarkName(formulaId)),
                "The unnumbered VisualTeX source unexpectedly owns a visible equation number.");
            AssertVisualTeXUnnumberedDisplayBaseline(
                document,
                shape,
                originalMetadata,
                "fresh unnumbered VisualTeX display");
            document.Save();
            File.Copy(
                documentPath,
                Path.Combine(
                    artifactRoot,
                    "visualtex-unnumbered-display-before-toggle.docx"),
                overwrite: true);

            shapeRange = shape.Range;
            var editSession = CreateNumberedPerformanceSession(
                "edit",
                formulaId,
                document.FullName,
                WordRangeReference(shapeRange.Start, shapeRange.End),
                originalMetadata,
                latex: @"x=2");
            editSession.Numbered = true;
            editSession.ExportResult = new OfficeExportDocument
            {
                Width = 180,
                Height = 60,
                Baseline = 45,
            };

            service.ReplaceOle(editSession, pngPath, emfPath);
            Release(shapeRange); shapeRange = null;
            Release(shape); shape = null;

            shape = FindVisualTeXOleByFormulaIdForNumberToggle(document, formulaId);
            var updatedMetadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException("The numbered VisualTeX result lost metadata.");
            AssertTrue(updatedMetadata.Numbered,
                "Editing the VisualTeX formula did not persist Numbered=true.");
            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "unnumbered-to-numbered edit");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(documentPath, ReadOnly: false, Visible: false);
            Release(shape); shape = null;
            shape = FindVisualTeXOleByFormulaIdForNumberToggle(document, formulaId);
            updatedMetadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException("Saved/reopened numbered VisualTeX result lost metadata.");
            AssertTrue(updatedMetadata.Numbered,
                "Saved/reopened VisualTeX result lost Numbered=true.");
            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "saved/reopened unnumbered-to-numbered edit");
            document.Activate();

            Release(shapeRange);
            shapeRange = null;
            Release(shape);
            shape = FindVisualTeXOleByFormulaIdForNumberToggle(
                document,
                formulaId);
            updatedMetadata =
                WordFormulaMetadataReader.TryRead(
                    shape)
                ?? throw new InvalidDataException(
                    "Numbered-to-unnumbered source lost metadata.");
            shapeRange =
                shape.Range.Duplicate;
            var disableSession =
                CreateNumberedPerformanceSession(
                    "edit",
                    formulaId,
                    document.FullName,
                    WordRangeReference(
                        shapeRange.Start,
                        shapeRange.End),
                    updatedMetadata,
                    latex: @"x=3");
            disableSession.Numbered = false;
            disableSession.ExportResult =
                new OfficeExportDocument
                {
                    Width = 180,
                    Height = 60,
                    Baseline = 45,
                };

            service.ReplaceOle(
                disableSession,
                pngPath,
                emfPath);
            Release(shapeRange);
            shapeRange = null;
            Release(shape);
            shape = null;

            shape =
                FindVisualTeXOleByFormulaIdForNumberToggle(
                    document,
                    formulaId);
            var disabledMetadata =
                WordFormulaMetadataReader.TryRead(
                    shape)
                ?? throw new InvalidDataException(
                    "Numbered-to-unnumbered result lost metadata.");
            AssertTrue(
                !disabledMetadata.Numbered,
                "Numbered-to-unnumbered edit did not persist Numbered=false.");
            AssertVisualTeXUnnumberedDisplayBaseline(
                document,
                shape,
                disabledMetadata,
                "numbered-to-unnumbered self-contained detach");
            AssertNoSelfContainedVisualTeXNumberingArtifacts(
                document,
                formulaId,
                "numbered-to-unnumbered self-contained detach");

            document.Save();
            document.Close(
                Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document =
                application.Documents.Open(
                    documentPath,
                    ReadOnly: false,
                    Visible: false);
            Release(shape);
            shape = null;
            shape =
                FindVisualTeXOleByFormulaIdForNumberToggle(
                    document,
                    formulaId);
            disabledMetadata =
                WordFormulaMetadataReader.TryRead(
                    shape)
                ?? throw new InvalidDataException(
                    "Saved/reopened unnumbered result lost metadata.");
            AssertTrue(
                !disabledMetadata.Numbered,
                "Saved/reopened result restored Numbered=true after detach.");
            AssertVisualTeXUnnumberedDisplayBaseline(
                document,
                shape,
                disabledMetadata,
                "saved/reopened numbered-to-unnumbered self-contained detach");
            AssertNoSelfContainedVisualTeXNumberingArtifacts(
                document,
                formulaId,
                "saved/reopened numbered-to-unnumbered self-contained detach");

            Console.WriteLine(
                "VisualTeX number-toggle acceptance passed: unnumbered→self-contained VisualTeXPlaceRef→unnumbered completed through ReplaceOle, removed the number field/TAB cleanly, and both states survived save/reopen.");
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            try { Directory.Delete(assetRoot, recursive: true); } catch { }
            ForceComCleanup();
        }
    }

    private static void RunWordInstalledVisualTeXNumberToggleCloseAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "installed-visualtex-unnumbered-to-numbered-close.docx");
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"installed-number-toggle-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var svg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"180\" height=\"60\" viewBox=\"0 0 180 60\"><text x=\"4\" y=\"42\" font-size=\"32\">x = 2</text></svg>";
        var svgPath = Path.Combine(assetRoot, "formula.svg");
        File.WriteAllText(svgPath, svg);
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 180, 60);
        var pngDataUrl = CreatePngDataUrl("installed-number-toggle", 180, 60);
        var pngPath = Path.Combine(assetRoot, "formula.png");
        File.WriteAllBytes(
            pngPath,
            Convert.FromBase64String(pngDataUrl.Substring(pngDataUrl.IndexOf(',') + 1)));

        var previousAcceptanceMode = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_ACCEPTANCE");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Office.COMAddIns? addIns = null;
        Office.COMAddIn? installedAddIn = null;
        object? callbacksObject = null;
        string? sessionId = null;
        try
        {
            // Installed-addin validation must not inherit the manual-host
            // acceptance guard: that guard intentionally makes the installed
            // COMAddIn inert when VISUALTEX_VSTO_ACCEPTANCE=1.
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            application = CreateWordApplication(visible: true);
            document = application.Documents.Add();
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();

            // Fixture preparation is intentionally separate from the operation
            // under test.  The edit itself below must flow through Word's actual
            // installed VisualTeX COM add-in automation object.
            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            try
            {
                application.Selection.SetRange(insertion.Start, insertion.End);
                var createSession = CreateNumberedPerformanceSession(
                    "create",
                    formulaId,
                    document.FullName,
                    WordRangeReference(insertion.Start, insertion.End),
                    originalMetadata: null,
                    latex: @"x=1");
                createSession.Numbered = false;
                createSession.ExportResult = new OfficeExportDocument
                {
                    Width = 180,
                    Height = 60,
                    Baseline = 45,
                };
                service.InsertOle(createSession, pngPath, emfPath);
            }
            finally { Release(insertion); }

            shape = FindVisualTeXOleByFormulaIdForNumberToggle(document, formulaId);
            var sourceMetadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException("Installed number-toggle source lost metadata.");
            AssertTrue(!sourceMetadata.Numbered,
                "Installed number-toggle source was unexpectedly numbered before editing.");
            shapeRange = shape.Range;
            shapeRange.Select();

            addIns = application.COMAddIns;
            object addInKey = "VisualTeX.WordVsto";
            installedAddIn = addIns.Item(ref addInKey);
            if (!installedAddIn.Connect) installedAddIn.Connect = true;
            for (var index = 0; index < 50 && installedAddIn.Object is null; index++)
            {
                WinForms.Application.DoEvents();
                Thread.Sleep(100);
            }
            callbacksObject = installedAddIn.Object
                ?? throw new InvalidOperationException(
                    "Installed VisualTeX.WordVsto automation object is unavailable.");

            var sessionsBefore = SnapshotSessionIds();
            dynamic callbacks = callbacksObject;
            callbacks.OnEditSelected(null);
            sessionId = WaitForNewSession(
                sessionsBefore,
                "word",
                TimeSpan.FromSeconds(30));
            var editorWindow = WaitForVisibleOfficeEditorWindow(TimeSpan.FromSeconds(20));
            ToggleOfficeEditorNumberCheckbox(editorWindow, TimeSpan.FromSeconds(15));
            var numberedDraft = WaitForNumberedEditorDraft(
                client,
                sessionId,
                TimeSpan.FromSeconds(30));
            AssertTrue(numberedDraft.Dirty,
                "Toggling the real editor Number checkbox did not mark the Session dirty.");
            AssertTrue(numberedDraft.Numbered,
                "Toggling the real editor Number checkbox did not persist Numbered=true before close.");
            AssertTrue(numberedDraft.ExportResult?.PngBase64?.Length > 0,
                "The real editor did not finish the PNG export required for native OLE commit before close.");

            if (!PostMessage(editorWindow, WmClose, UIntPtr.Zero, IntPtr.Zero))
                throw new InvalidOperationException(
                    $"Unable to post WM_CLOSE to the installed VisualTeX editor (Win32 {System.Runtime.InteropServices.Marshal.GetLastWin32Error()}).");

            var terminal = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            AssertEqual(
                "completed",
                terminal.Status,
                terminal.Error
                ?? "Installed unnumbered->numbered edit did not complete after closing the editor.");
            WaitForOfficeEditorHidden(TimeSpan.FromSeconds(15));
            sessionId = null;

            Release(shapeRange); shapeRange = null;
            Release(shape); shape = null;
            shape = FindVisualTeXOleByFormulaIdForNumberToggle(document, formulaId);
            var updatedMetadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException("Installed numbered result lost metadata.");
            AssertTrue(updatedMetadata.Numbered,
                "Installed editor close did not persist Numbered=true.");
            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "installed editor unnumbered-to-numbered close");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            Release(shape); shape = null;
            shape = FindVisualTeXOleByFormulaIdForNumberToggle(document, formulaId);
            updatedMetadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException(
                    "Saved/reopened installed number-toggle result lost metadata.");
            AssertTrue(updatedMetadata.Numbered,
                "Saved/reopened installed number-toggle result lost Numbered=true.");
            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "saved/reopened installed number-toggle result");

            Console.WriteLine(
                "[INSTALLED ADD-IN NUMBER TOGGLE] Real VisualTeX.WordVsto editor Session changed Numbered=false to true, WM_CLOSE completed the Session, the editor hid, the host uses MathType-style center/right tabs, and save/reopen retained the numbered VisualTeX OLE.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_ACCEPTANCE",
                previousAcceptanceMode);
            if (!string.IsNullOrWhiteSpace(sessionId))
            {
                try
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
                    client.CloseEditorAsync(sessionId!, CancellationToken.None)
                        .GetAwaiter().GetResult();
                }
                catch { }
            }
            Release(shapeRange);
            Release(shape);
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
            try { Directory.Delete(assetRoot, recursive: true); } catch { }
            ForceComCleanup();
        }
    }

    private static void AssertNoSelfContainedVisualTeXNumberingArtifacts(
        Word.Document document,
        string formulaId,
        string context)
    {
        Word.Bookmarks? bookmarks = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            bookmarks =
                document.Bookmarks;
            var guid =
                Guid.Parse(
                    formulaId)
                    .ToString("N");
            AssertTrue(
                !bookmarks.Exists(
                    "VTEqCap_" + guid),
                context
                + ": detached self-contained VisualTeX host has a VTEqCap bookmark.");
            AssertTrue(
                !bookmarks.Exists(
                    "VTEqNum_" + guid),
                context
                + ": detached self-contained VisualTeX host has a VTEqNum bookmark without a live number.");

            shape =
                FindVisualTeXOleByFormulaIdForNumberToggle(
                    document,
                    formulaId);
            shapeRange =
                shape.Range.Duplicate;
            paragraphs =
                shapeRange.Paragraphs;
            AssertEqual(
                1,
                paragraphs.Count,
                context
                + ": detached VisualTeX OLE spans more than one paragraph.");
            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            fields =
                paragraphRange.Fields;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(code);
                code = null;
                Release(field);
                field =
                    fields[index];
                if (field.Type !=
                    Word.WdFieldType.wdFieldMacroButton)
                    continue;
                code =
                    field.Code.Duplicate;
                AssertTrue(
                    !WordVisualTeXParagraphNumbering
                        .IsPlaceRefCode(
                            code.Text),
                    context
                    + ": VisualTeXPlaceRef survived numbered detach.");
            }

            var owner =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    shapeRange,
                    WordFormulaHostKind.VisualTeX)
                ?? throw new InvalidDataException(
                    context
                    + ": detached VisualTeX OLE is no longer resolvable.");
            var numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    owner);
            AssertTrue(
                !numbering.Numbered
                && numbering.ContainerKind ==
                    WordFormulaNumberingContainerKind.None,
                context
                + ": numbering resolver still reports a container after detach.");
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
            Release(shape);
            Release(bookmarks);
        }
    }

    private static void ToggleOfficeEditorNumberCheckbox(
        IntPtr editorWindow,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        Exception? lastError = null;
        while (DateTime.UtcNow < deadline)
        {
            try
            {
                var root = AutomationElement.FromHandle(editorWindow);
                var checkboxes = root.FindAll(
                    TreeScope.Descendants,
                    new PropertyCondition(
                        AutomationElement.ControlTypeProperty,
                        ControlType.CheckBox));
                foreach (AutomationElement checkbox in checkboxes)
                {
                    var name = checkbox.Current.Name ?? string.Empty;
                    if (name.IndexOf("编号", StringComparison.OrdinalIgnoreCase) < 0
                        && name.IndexOf("Number", StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    if (!checkbox.TryGetCurrentPattern(
                            TogglePattern.Pattern,
                            out var patternObject)
                        || patternObject is not TogglePattern togglePattern)
                        continue;
                    if (togglePattern.Current.ToggleState == ToggleState.Off)
                        togglePattern.Toggle();
                    if (togglePattern.Current.ToggleState != ToggleState.On)
                        throw new InvalidOperationException(
                            "The VisualTeX Number checkbox did not enter the checked state.");
                    Console.WriteLine(
                        $"    [installed editor UI] toggled checkbox '{name}' to checked via Windows UI Automation.");
                    return;
                }
            }
            catch (Exception error)
            {
                lastError = error;
            }
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
        }
        throw new TimeoutException(
            "Could not locate and toggle the real VisualTeX editor Number checkbox through Windows UI Automation.",
            lastError);
    }

    private static OfficeSessionDocument WaitForNumberedEditorDraft(
        VisualTeXSessionClient client,
        string sessionId,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        OfficeSessionDocument? last = null;
        while (DateTime.UtcNow < deadline)
        {
            last = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            if (last.Numbered
                && last.Dirty
                && last.Status == "editing"
                && last.ExportResult?.PngBase64?.Length > 0)
                return last;
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
        }
        throw new TimeoutException(
            $"The real VisualTeX editor did not autosave a complete numbered draft before close. "
            + $"Last state: numbered={last?.Numbered}, dirty={last?.Dirty}, status={last?.Status}, png={last?.ExportResult?.PngBase64?.Length ?? 0}.");
    }

    private static Word.InlineShape FindVisualTeXOleByFormulaIdForNumberToggle(
        Word.Document document,
        string formulaId)
    {
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                var metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata is null
                    || !string.Equals(
                        metadata.FormulaId,
                        formulaId,
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                var result = shape;
                shape = null;
                return result;
            }
            throw new InvalidDataException($"VisualTeX OLE formula '{formulaId}' was not found.");
        }
        finally
        {
            Release(shape);
            Release(shapes);
        }
    }

    private static void AssertVisualTeXUnnumberedDisplayBaseline(
        Word.Document document,
        Word.InlineShape shape,
        FormulaMetadata metadata,
        string context)
    {
        Word.Range? shapeRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? ownerRange = null;
        Word.Range? objectResultRange = null;
        Word.Font? objectResultFont = null;
        Word.Range? paragraphMark = null;
        Word.Font? paragraphMarkFont = null;
        Word.ParagraphFormat? paragraphFormat = null;
        Word.TabStops? tabStops = null;
        Word.TabStop? tabStop = null;
        Word.Range? prefix = null;
        try
        {
            AssertTrue(!metadata.Numbered, context + ": fixture is unexpectedly numbered.");
            AssertEqual("block", metadata.DisplayMode, context + ": fixture is not display math.");

            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            AssertEqual(1, paragraphs.Count, context + ": formula does not own one paragraph.");
            paragraph = paragraphs[1];
            ownerRange = paragraph.Range;
            paragraphFormat = paragraph.Format;
            AssertEqual(
                Word.WdParagraphAlignment.wdAlignParagraphJustify,
                paragraphFormat.Alignment,
                context + ": unnumbered display paragraph is not using the shared tab-driven display layout.");
            tabStops = paragraphFormat.TabStops;
            var hasCenterTab = false;
            var hasRightTab = false;
            for (var index = 1; index <= tabStops.Count; index++)
            {
                Release(tabStop); tabStop = null;
                tabStop = tabStops[index];
                hasCenterTab |=
                    tabStop.Alignment == Word.WdTabAlignment.wdAlignTabCenter;
                hasRightTab |=
                    tabStop.Alignment == Word.WdTabAlignment.wdAlignTabRight;
            }
            AssertTrue(
                hasCenterTab && hasRightTab,
                context + ": unnumbered display paragraph lost the shared center/right tab stops.");
            prefix = document.Range(ownerRange.Start, shapeRange.Start);
            AssertEqual(
                "\t",
                prefix.Text,
                context + ": unnumbered display OLE is not anchored by one leading center tab.");

            for (var position = shapeRange.Start; position < shapeRange.End; position++)
            {
                Release(objectResultFont); objectResultFont = null;
                Release(objectResultRange); objectResultRange = null;
                objectResultRange = document.Range(position, position + 1);
                if (!string.Equals(objectResultRange.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                objectResultFont = objectResultRange.Font;
                break;
            }
            if (objectResultFont is null)
                throw new InvalidDataException(context + ": VisualTeX OLE has no U+0001 object-result character.");

            AssertTrue(
                metadata.WordDisplayPreviewInkHeightRatio.HasValue
                && metadata.WordDisplayPreviewBottomWhitespaceRatio.HasValue,
                context + ": display preview ink metrics were not persisted.");
            paragraphMark = document.Range(ownerRange.End - 1, ownerRange.End);
            AssertEqual("\r", paragraphMark.Text, context + ": formula is not followed by one ordinary paragraph mark.");
            paragraphMarkFont = paragraphMark.Font;
            var expectedPosition = WordInlineAlignment.CalculateDisplayInkCenterPosition(
                shape.Height,
                (float)metadata.WordDisplayPreviewInkHeightRatio!.Value,
                (float)metadata.WordDisplayPreviewBottomWhitespaceRatio!.Value,
                paragraphMarkFont.Size);
            AssertNear(
                expectedPosition,
                objectResultFont.Position,
                0.1f,
                context + ": OLE bottom remained on the paragraph baseline, leaving the paragraph mark at the lower-right corner.");
            AssertTrue(
                expectedPosition < 0,
                context + ": fixture did not exercise a visible display-baseline displacement.");
            AssertNear(
                0f,
                paragraphMarkFont.Position,
                0.1f,
                context + ": paragraph mark itself has a manual vertical offset.");
        }
        finally
        {
            Release(prefix);
            Release(tabStop);
            Release(tabStops);
            Release(paragraphFormat);
            Release(paragraphMarkFont);
            Release(paragraphMark);
            Release(objectResultFont);
            Release(objectResultRange);
            Release(ownerRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }
}
