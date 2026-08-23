using Extensibility;
using Microsoft.Office.Core;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordFontSizePersistenceAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var path = Path.Combine(
            artifactRoot,
            "VisualTeX-Word-Font-Size-Persistence.docx");
        if (File.Exists(path)) File.Delete(path);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        Word.InlineShape? shape = null;
        COMAddIns? installedAddIns = null;
        COMAddIn? installedAddIn = null;
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
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);

            Console.WriteLine("[FONT PERSIST 1/5] Creating a real inline VisualTeX native OLE formula...");
            var existing = SnapshotSessionIds();
            addIn.OnInsertInline(new object());
            var sessionId = WaitForNewSession(
                existing,
                "word",
                TimeSpan.FromSeconds(30));
            var session = client.GetSessionAsync(
                    sessionId,
                    CancellationToken.None)
                .GetAwaiter()
                .GetResult();
            Commit(
                client,
                session,
                "inline",
                FormulaOleContract.NativeOleMode,
                "x+y");
            var final = WaitForTerminal(
                client,
                sessionId,
                TimeSpan.FromSeconds(45));
            AssertEqual(
                "completed",
                final.Status,
                final.Error ?? "Font-size persistence fixture failed.");
            client.CloseEditorAsync(sessionId, CancellationToken.None)
                .GetAwaiter()
                .GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            shape = document.InlineShapes[1];
            shape.Range.Select();
            AssertTrue(
                addIn.GetFormulaFontSizeEnabled(null!),
                "Font-size control was disabled for the native OLE fixture.");

            Console.WriteLine("[FONT PERSIST 2/5] Saving 四号 / 14 pt and reopening Word...");
            addIn.OnFormulaFontSizeChanged(null!, "四号");
            Release(shape);
            shape = document.InlineShapes[1];
            shape.Range.Select();
            var authoritative14 = WordFormulaMetadataReader.TryReadAuthoritative(shape)
                ?? throw new InvalidDataException(
                    "14 pt authoritative native OLE metadata could not be read.");
            Console.WriteLine(
                $"  authoritative 14pt: font={authoritative14.FontSizePt}; display={authoritative14.DisplayMode}; "
                + $"stored={authoritative14.WordInlineOleWidthPt}x{authoritative14.WordInlineOleHeightPt}; "
                + $"live={shape.Width}x{shape.Height}");
            var metadata14 = new WordFormulaService(application)
                    .ReadSelection()
                    .Metadata
                ?? throw new InvalidDataException(
                    "14 pt native OLE metadata could not be read.");
            AssertTrue(
                metadata14.FontSizePt.HasValue,
                "四号 metadata lost its semantic font size before save.");
            AssertNear(
                14f,
                (float)metadata14.FontSizePt.GetValueOrDefault(),
                0.001f,
                "四号 was not persisted as exactly 14 pt before save.");
            AssertEqual(
                "四号",
                addIn.GetFormulaFontSizeText(null!),
                "Ribbon did not report 四号 before save.");
            var width14 = shape.Width;
            var height14 = shape.Height;

            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            reopened = application.Documents.Open(
                path,
                ReadOnly: false,
                Visible: true);
            reopened.Activate();
            AssertEqual(
                1,
                reopened.InlineShapes.Count,
                "Reopened 14 pt document lost the native OLE formula.");
            Release(shape);
            shape = reopened.InlineShapes[1];
            shape.Range.Select();
            var reopened14 = new WordFormulaService(application)
                    .ReadSelection()
                    .Metadata
                ?? throw new InvalidDataException(
                    "Reopened 14 pt native OLE metadata could not be read.");
            AssertTrue(
                reopened14.FontSizePt.HasValue,
                "Reopened 四号 metadata lost its semantic font size.");
            AssertNear(
                14f,
                (float)reopened14.FontSizePt.GetValueOrDefault(),
                0.001f,
                "四号 changed after Word save/reopen.");
            AssertEqual(
                "四号",
                addIn.GetFormulaFontSizeText(null!),
                "Ribbon did not restore 四号 after Word reopen.");
            AssertNear(
                width14,
                shape.Width,
                0.75f,
                "14 pt OLE width changed after save/reopen.");
            AssertNear(
                height14,
                shape.Height,
                0.75f,
                "14 pt OLE height changed after save/reopen.");

            Console.WriteLine("[FONT PERSIST 3/5] Saving a typed 13.25 pt value and reopening Word...");
            addIn.OnFormulaFontSizeChanged(null!, "13.25");
            Release(shape);
            shape = reopened.InlineShapes[1];
            shape.Range.Select();
            var metadata1325 = new WordFormulaService(application)
                    .ReadSelection()
                    .Metadata
                ?? throw new InvalidDataException(
                    "13.25 pt native OLE metadata could not be read.");
            AssertTrue(
                metadata1325.FontSizePt.HasValue,
                "Typed 13.25 pt metadata lost its semantic font size before save.");
            AssertNear(
                13.25f,
                (float)metadata1325.FontSizePt.GetValueOrDefault(),
                0.001f,
                "Typed 13.25 pt was quantized before save.");
            AssertEqual(
                "13.25",
                addIn.GetFormulaFontSizeText(null!),
                "Ribbon quantized 13.25 pt before save.");
            var width1325 = shape.Width;
            var height1325 = shape.Height;

            reopened.Save();
            reopened.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(reopened);
            reopened = application.Documents.Open(
                path,
                ReadOnly: false,
                Visible: true);
            reopened.Activate();
            Release(shape);
            shape = reopened.InlineShapes[1];
            shape.Range.Select();

            Console.WriteLine("[FONT PERSIST 4/5] Verifying authoritative metadata, Ribbon text, and geometry after second reopen...");
            var finalMetadata = new WordFormulaService(application)
                    .ReadSelection()
                    .Metadata
                ?? throw new InvalidDataException(
                    "Reopened 13.25 pt native OLE metadata could not be read.");
            AssertTrue(
                finalMetadata.FontSizePt.HasValue,
                "Reopened 13.25 pt metadata lost its semantic font size.");
            AssertNear(
                13.25f,
                (float)finalMetadata.FontSizePt.GetValueOrDefault(),
                0.001f,
                "Typed 13.25 pt changed after Word save/reopen.");
            AssertEqual(
                "13.25",
                addIn.GetFormulaFontSizeText(null!),
                "Ribbon did not restore the typed 13.25 pt value after reopen.");
            AssertNear(
                width1325,
                shape.Width,
                0.75f,
                "13.25 pt OLE width changed after save/reopen.");
            AssertNear(
                height1325,
                shape.Height,
                0.75f,
                "13.25 pt OLE height changed after save/reopen.");

            Console.WriteLine(
                $"[FONT PERSIST 5/5] Passed: 四号 stayed 14 pt; typed 13.25 stayed 13.25 pt; "
                + $"geometry remained stable across both reopen cycles. Artifact: {path}");
        }
        finally
        {
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
            if (installedAddIn is not null)
            {
                try { installedAddIn.Connect = true; } catch { }
            }
            Release(shape);
            if (reopened is not null)
            {
                try
                {
                    reopened.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(reopened);
            if (document is not null)
            {
                try
                {
                    document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(document);
            Release(installedAddIn);
            Release(installedAddIns);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }
}
