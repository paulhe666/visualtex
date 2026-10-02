using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunNativeCreateEditAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        ThisAddIn? addIn = null;
        VisualTeXSessionClient? client = null;
        dynamic? callbacks = null;
        Array custom = Array.Empty<object>();
        using var codeBaseOverride = AttachActiveWord ? null : CreateUiE2eWordVstoCodeBaseOverride();
        try
        {
            var inProcess = Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_EDIT_IN_PROCESS") == "1";
            if (inProcess) Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            application = AttachActiveWord
                ? (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application")
                : CreateWordApplication(visible: inProcess);
            var clone = Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_EDIT_CLONE");
            document = string.IsNullOrEmpty(clone) ? application.Documents.Add() : application.Documents.Open(clone);
            if (Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_EDIT_TABLE") == "1")
            {
                var slot = document.Range(0, 0);
                var table = document.Tables.Add(slot, 1, 2);
                table.Cell(1, 2).Range.Text = "UNTOUCHED";
                var cell = table.Cell(1, 1).Range;
                cell.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                cell.Select();
                Release(cell); Release(table); Release(slot);
            }
            document.TrackRevisions = Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_EDIT_TRACKING") == "1";
            var format = Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_EDIT_FORMAT");
            if (!string.IsNullOrEmpty(format))
                WordEquationNumbering.RestoreEquationNumberFormatForConversion(document, format);
            var service = new WordFormulaService(application);
            var selection = service.ReadCreateSelection();
            var session = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Host = "word", Mode = "create",
                FormulaId = Guid.NewGuid().ToString("D"),
                SourceDocumentId = selection.DocumentId,
                SourceObjectId = selection.ObjectId,
                ObjectMode = FormulaOleContract.WordOmmlMode,
                DisplayMode = "block", Numbered = true,
                CodeFormat = "latex", FontSizePt = 12,
                Lines = new List<FormulaLine> { new() { Id = Guid.NewGuid().ToString("D"), Latex = "x+1" } },
            };
            var original = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mn>1</mn></math>";
            var edited = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mn>2</mn></math>";
            var replay = Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_EDIT_REPLAY");
            if (!string.IsNullOrEmpty(replay))
            {
                using var fixture = JsonDocument.Parse(File.ReadAllText(replay));
                original = fixture.RootElement.GetProperty("original").GetString()!;
                edited = fixture.RootElement.GetProperty("edited").GetString()!;
                session.FontSizePt = 10.5;
            }
            Console.WriteLine("Creating numbered OMML through the editor's service entry point.");
            if (inProcess || Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_EDIT_RIBBON") == "1")
            {
                client = new VisualTeXSessionClient();
                client.EnsureHealthyAsync(CancellationToken.None).GetAwaiter().GetResult();
                if (inProcess)
                {
                    object key = "VisualTeX.WordVsto";
                    var installed = application.COMAddIns.Item(ref key);
                    callbacks = installed.Object;
                    Release(installed);
                }
                else
                {
                    addIn = new ThisAddIn();
                    addIn.OnConnection(application, Extensibility.ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);
                    callbacks = addIn;
                }
                if (string.IsNullOrEmpty(clone))
                    NativeEditRibbonCommit(client, addIn, () => callbacks.OnInsertDisplayOmml(new object()), original, artifactRoot, "create");
            }
            else if (string.IsNullOrEmpty(clone)) service.ApplyOmmlVisualTeXHostSession(session, selection.ObjectMode, original, null, null);
            AssertEqual(1, document.OMaths.Count, "Creation must produce one formula.");
            if (Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_EDIT_OBSERVATIONAL") == "1")
            {
                var exact = document.OMaths[1].Range;
                var undo = application.UndoRecord;
                try
                {
                    undo.StartCustomRecord("Native number read remains observational");
                    AssertTrue(WordNativeOmmlNumbering.TryResolveUnownedNativeStructure(document, exact, out _),
                        "The existing numbered equation was not recognized.");
                    AssertTrue(undo.IsRecordingCustomRecord && undo.CustomRecordLevel == 1,
                        "Reading native numbering ended Word's active undo record.");
                }
                finally { undo.EndCustomRecord(); Release(undo); Release(exact); }
                Console.WriteLine("PASS native numbering lookup leaves the active Word undo record intact.");
                return;
            }
            if (Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_EDIT_SECOND") == "1")
            {
                application.Selection.EndKey(Word.WdUnits.wdStory);
                application.Selection.TypeParagraph();
                var secondSelection = service.ReadCreateSelection();
                session.SourceObjectId = secondSelection.ObjectId;
                service.ApplyOmmlVisualTeXHostSession(session, FormulaOleContract.WordOmmlMode, original, null, null);
                AssertEqual(2, document.OMaths.Count, "Second insertion must produce a second formula.");
                var labels = ReadNativeEditAddNumberState(document).Labels;
                var selectedFormat = EquationNumberFormat.Resolve(format);
                var prefix = selectedFormat.UsesHeading
                    ? string.Join(".", Enumerable.Repeat("0", selectedFormat.HeadingLevel)) + selectedFormat.Separator
                    : string.Empty;
                AssertEqual(prefix + "1", labels[0], "First number must use the document's pre-heading scope.");
                AssertEqual(prefix + "2", labels[1], "The next number must continue the same sequence.");
                Console.WriteLine("PASS second numbered insertion without a heading.");
                return;
            }
            document.OMaths[1].Range.Select();
            if (Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_EDIT_CARET") == "1")
            {
                var point = document.OMaths[1].Range.Start + 2;
                document.Range(point, point).Select();
            }
            var originalNumber = NativeEditNumberState(document);
            var history = Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_EDIT_HISTORY") == "1";
            using var before = history ? new RootDocumentEvidence(document, Path.Combine(artifactRoot, "before-edit")) : null;
            if (callbacks is null)
            {
                var editSelection = service.ReadSelection();
                File.WriteAllText(Path.Combine(artifactRoot, "selection.json"), JsonSerializer.Serialize(editSelection));
                session.Mode = "edit";
                session.SourceDocumentId = editSelection.DocumentId;
                session.SourceObjectId = editSelection.ObjectId;
                session.FormulaId = editSelection.FormulaId!;
                session.OriginalMetadata = editSelection.Metadata;
                session.DisplayMode = editSelection.Metadata!.DisplayMode;
                session.Numbered = editSelection.Metadata.Numbered;
                session.Lines[0].Latex = "x+2";
            }
            Console.WriteLine("Editing the native host without an extra XML read before the Ribbon callback.");
            try
            {
                if (callbacks is not null)
                    NativeEditRibbonCommit(client!, addIn, () => callbacks.OnEditSelected(new object()), edited, artifactRoot, "edit");
                else service.ApplyOmmlVisualTeXHostSession(session, FormulaOleContract.WordOmmlMode, edited, null, null);
            }
            catch (Exception error)
            {
                File.WriteAllText(Path.Combine(artifactRoot, "failure.txt"), error.ToString());
                File.WriteAllText(Path.Combine(artifactRoot, "failed.xml"), document.Content.WordOpenXML);
                throw;
            }
            File.WriteAllText(Path.Combine(artifactRoot, "after.xml"), document.Content.WordOpenXML);
            AssertEqual(1, document.OMaths.Count, "Editing must retain one formula.");
            document.OMaths[1].Range.Select();
            var actual = service.ReadSelection().Metadata!.Latex;
            var expected = MathMlToLatexConverter.Convert(edited);
            AssertEqual(WordFormulaMutationValidator.NormalizeWordNativeLatexSemantics(expected),
                WordFormulaMutationValidator.NormalizeWordNativeLatexSemantics(actual), "The edited body was not applied.");
            AssertEqual(originalNumber, NativeEditNumberState(document), "Editing changed the existing number or field code.");
            if (history)
            {
                using var after = new RootDocumentEvidence(document, Path.Combine(artifactRoot, "after-edit"));
                document.Undo(1); before!.VerifyRollback();
                document.Redo(1); after.VerifyRollback();
                var path = Path.Combine(artifactRoot, "edited.docx");
                document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
                document.Close(Word.WdSaveOptions.wdSaveChanges); Release(document);
                document = application.Documents.Open(path);
                AssertEqual(originalNumber, NativeEditNumberState(document), "Save/reopen changed native numbering.");
                document.OMaths[1].Range.Select();
                AssertEqual(WordFormulaMutationValidator.NormalizeWordNativeLatexSemantics(expected),
                    WordFormulaMutationValidator.NormalizeWordNativeLatexSemantics(service.ReadSelection().Metadata!.Latex),
                    "The reopened equation lost its edited mathematical content.");
                Console.WriteLine("PASS strict one-step Undo/Redo, number preservation, save/reopen and semantic readback.");
            }
            Console.WriteLine("PASS blank document numbered OMML create/read/edit/apply.");
        }
        finally
        {
            if (addIn is not null) addIn.OnDisconnection(Extensibility.ext_DisconnectMode.ext_dm_UserClosed, ref custom);
            if (addIn is null) Release(callbacks);
            client?.Dispose();
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
            finally { Release(document); QuitWordApplicationIfOwned(application); Release(application); }
        }
    }

    private static string NativeEditNumberState(Word.Document document)
    {
        var math = document.OMaths[1]; var exact = math.Range;
        try
        {
            AssertTrue(WordNativeOmmlNumbering.TryResolveUnownedNativeStructure(document, exact, out var number),
                "The exact native SEQ host must remain numbered.");
            var codes = new List<string>(); var fields = exact.Fields;
            try
            {
                for (var i = 1; i <= fields.Count; i++)
                {
                    var field = fields[i]; var code = field.Code;
                    try { if (WordNumberingFieldCollection.IsCurrentField(field)) codes.Add(code.Text); }
                    finally { Release(code); Release(field); }
                }
            }
            finally { Release(fields); }
            return WordNativeOmmlNumbering.ReadVisibleNumber(document, number) + "|" + string.Join("|", codes);
        }
        finally { Release(exact); Release(math); }
    }

    private static void NativeEditRibbonCommit(VisualTeXSessionClient client, ThisAddIn? addIn,
        Action command, string mathMl, string artifacts, string stage)
    {
        var existing = SnapshotSessionIds();
        command();
        var id = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
        var ready = DateTime.UtcNow.AddSeconds(3);
        while (DateTime.UtcNow < ready)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(50);
        }
        var session = client.GetSessionAsync(id, CancellationToken.None).GetAwaiter().GetResult();
        Commit(client, session, "block", FormulaOleContract.WordOmmlMode,
            MathMlToLatexConverter.Convert(mathMl), numbered: true, mathMl: mathMl);
        var terminal = WaitForTerminal(client, id, TimeSpan.FromSeconds(30));
        File.WriteAllText(Path.Combine(artifacts, stage + "-session.json"), JsonSerializer.Serialize(terminal));
        client.CloseEditorAsync(id, CancellationToken.None).GetAwaiter().GetResult();
        if (addIn is not null) WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
        AssertEqual("completed", terminal.Status, terminal.Error ?? stage + " failed");
    }
}
