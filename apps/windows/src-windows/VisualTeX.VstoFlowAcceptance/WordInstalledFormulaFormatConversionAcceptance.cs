using Extensibility;
using Office = Microsoft.Office.Core;
using WinForms = System.Windows.Forms;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordInstalledFormatConversionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var fixturePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_FORMAT_CONVERSION_FIXTURE");
        if (string.IsNullOrWhiteSpace(fixturePath)
            || !File.Exists(fixturePath))
            throw new InvalidOperationException(
                "Installed format-conversion acceptance requires VISUALTEX_FORMAT_CONVERSION_FIXTURE pointing to a real Word document created by the installed VisualTeX add-in. Artificial service-created OLE fixtures are intentionally forbidden.");

        var previousAcceptanceMode = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_ACCEPTANCE");
        var previousFormatConversionAcceptanceMode = Environment.GetEnvironmentVariable(
            "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var useLocalAddIn = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_USE_LOCAL_ADDIN"),
            "1",
            StringComparison.Ordinal);
        var expectRollback = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_EXPECT_ROLLBACK"),
            "1",
            StringComparison.Ordinal);
        Word.Application? application = null;
        Word.Document? fixtureDocument = null;
        Word.Document? document = null;
        Office.COMAddIns? addIns = null;
        Office.COMAddIn? installedAddIn = null;
        ThisAddIn? localAddIn = null;
        Array? localCustom = null;
        object? callbacksObject = null;
        try
        {
            // Installed acceptance must use Word's normal product lifecycle: the
            // installed add-in is active from Word startup, exactly as it is for a
            // user opening a document and clicking the Ribbon command. Only local
            // source diagnostics suppress the installed instance to avoid two add-ins
            // handling the same Word application.
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_ACCEPTANCE",
                useLocalAddIn ? "1" : null);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                "1");
            application = CreateWordApplication(visible: true);
            int fixtureFormulaCount;
            int fixtureNumberedCount;
            if (string.Equals(
                    Path.GetExtension(fixturePath),
                    ".docx",
                    StringComparison.OrdinalIgnoreCase))
            {
                // A saved DOCX is the closest possible match to the user's real
                // workflow. Copy it at the filesystem level before Word opens it so
                // no FormattedText/paste event can race the installed add-in.
                var writableFixturePath = Path.Combine(
                    artifactRoot,
                    "installed-format-conversion-input.docx");
                File.Copy(
                    Path.GetFullPath(fixturePath),
                    writableFixturePath,
                    overwrite: true);
                document = application.Documents.Open(
                    writableFixturePath,
                    ReadOnly: false,
                    AddToRecentFiles: false);
                fixtureFormulaCount = CountInstalledVisualTeXOleShapes(document);
                fixtureNumberedCount = CountInstalledVisualTeXNumberedFormulaHosts(document);
                Console.WriteLine(
                    $"[INSTALLED ADD-IN FIXTURE] direct saved DOCX; formulas={fixtureFormulaCount}, numbered={fixtureNumberedCount}");
            }
            else
            {
                // AutoRecover input is retained only for source diagnostics. It
                // cannot be treated as a final installed acceptance because Word's
                // FormattedText clone can trigger add-in copy/reconciliation work.
                fixtureDocument = application.Documents.Open(
                    Path.GetFullPath(fixturePath),
                    ReadOnly: true,
                    AddToRecentFiles: false);
                fixtureFormulaCount = CountInstalledVisualTeXOleShapes(fixtureDocument);
                fixtureNumberedCount = CountInstalledVisualTeXNumberedFormulaHosts(fixtureDocument);
                document = application.Documents.Add();
                Word.Range? fixtureContent = null;
                Word.Range? writableContent = null;
                Word.Range? formattedFixture = null;
                try
                {
                    fixtureContent = fixtureDocument.Content;
                    writableContent = document.Content;
                    formattedFixture = fixtureContent.FormattedText;
                    writableContent.FormattedText = formattedFixture;
                }
                finally
                {
                    Release(formattedFixture);
                    Release(writableContent);
                    Release(fixtureContent);
                }
                WinForms.Application.DoEvents();
                Thread.Sleep(300);
            }
            var sourceFormulaCount = CountInstalledVisualTeXOleShapes(document);
            var sourceNumberedCount = CountInstalledVisualTeXNumberedFormulaHosts(document);
            var initialMathTypeCount = CountInstalledMathTypeOleShapes(document);
            var initialMathTypePlaceRefCount = CountMathTypePlaceRefFields(document);
            var sourceNumberingBookmarkCount = CountInstalledVisualTeXNumberingBookmarks(document);
            var sourceRawBridgeCounts = ReadInstalledVisualTeXRawBridgeCounts(document);
            AssertEqual(fixtureFormulaCount, sourceFormulaCount,
                "Preparing the real VisualTeX fixture changed the VisualTeX OLE count.");
            AssertEqual(fixtureNumberedCount, sourceNumberedCount,
                "Preparing the real VisualTeX fixture changed the numbered-state count.");
            if (sourceFormulaCount <= 0)
                throw new InvalidDataException(
                    "The real fixture does not contain any VisualTeX OLE formulas.");
            if (sourceNumberedCount < 0 || sourceNumberedCount > sourceFormulaCount)
                throw new InvalidDataException(
                    $"The real fixture reported an impossible numbered-state count; total={sourceFormulaCount}, numbered={sourceNumberedCount}.");
            Console.WriteLine(
                $"[INSTALLED ADD-IN SOURCE COUNTS] VT={sourceFormulaCount} VTNumbered={sourceNumberedCount} existingMT={initialMathTypeCount} existingMTPlaceRef={initialMathTypePlaceRefCount}");
            Console.WriteLine("[INSTALLED ADD-IN SOURCE SHAPES]");
            DumpInstalledFormulaShapes(document);

            if (useLocalAddIn)
            {
                Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
                localAddIn = new ThisAddIn();
                localCustom = Array.Empty<object>();
                localAddIn.OnConnection(
                    application,
                    ext_ConnectMode.ext_cm_AfterStartup,
                    localAddIn,
                    ref localCustom);
                callbacksObject = localAddIn;
                Console.WriteLine("[FORMAT CONVERSION DIAGNOSTIC] current source ThisAddIn manually hosted after the real fixture was prepared.");
            }
            else
            {
                addIns = application.COMAddIns;
                object addInKey = "VisualTeX.WordVsto";
                installedAddIn = addIns.Item(ref addInKey);
                if (!installedAddIn.Connect)
                    installedAddIn.Connect = true;
                for (var index = 0; index < 50 && installedAddIn.Object is null; index++)
                {
                    WinForms.Application.DoEvents();
                    Thread.Sleep(100);
                }
                callbacksObject = installedAddIn.Object
                    ?? throw new InvalidOperationException(
                        "Installed VisualTeX Word add-in automation object is unavailable after reconnect.");
            }
            var wholeDocumentCallback = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_WHOLE_DOCUMENT"),
                "1",
                StringComparison.Ordinal);
            dynamic callbacks = callbacksObject;
            if (wholeDocumentCallback)
            {
                callbacks.OnConvertVisualTeXToMathTypeDocument(null);
            }
            else
            {
                document.Content.Select();
                callbacks.OnConvertVisualTeXToMathTypeSelection(null);
            }

            var deadline = DateTime.UtcNow.AddSeconds(120);
            var tracePath = Environment.GetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH");
            var completionMarker = expectRollback
                ? "format-conversion-stopped"
                : "format-conversion-complete";
            var completionObserved = false;
            while (DateTime.UtcNow < deadline)
            {
                WinForms.Application.DoEvents();
                Thread.Sleep(200);
                if (string.IsNullOrWhiteSpace(tracePath)
                    || !File.Exists(tracePath))
                    continue;
                string trace;
                try { trace = File.ReadAllText(tracePath); }
                catch { continue; }
                if (trace.IndexOf(completionMarker, StringComparison.Ordinal) < 0)
                    continue;
                completionObserved = true;
                break;
            }
            AssertTrue(completionObserved,
                $"Installed add-in did not report '{completionMarker}' before the acceptance timeout.");
            // Do not poll Word's InlineShapes while the async Ribbon callback is
            // mutating OLE objects. Wait until the add-in has reported completion,
            // then give Word one UI cycle to settle before inspecting the document.
            for (var settle = 0; settle < 10; settle++)
            {
                WinForms.Application.DoEvents();
                Thread.Sleep(100);
            }

            var finalVisualTeXCount = CountInstalledVisualTeXOleShapes(document);
            var finalMathTypeCount = CountInstalledMathTypeOleShapes(document);
            var finalPlaceRefCount = CountMathTypePlaceRefFields(document);
            Console.WriteLine(
                $"[INSTALLED ADD-IN DIAGNOSTIC] VT={finalVisualTeXCount} MT={finalMathTypeCount} MTPlaceRef={finalPlaceRefCount}");
            DumpInstalledFormulaShapes(document);

            if (expectRollback)
            {
                AssertEqual(sourceFormulaCount, finalVisualTeXCount + finalMathTypeCount,
                    "A failed transactional format conversion changed the total formula-object count.");
                AssertEqual(sourceNumberedCount, CountInstalledVisualTeXNumberedFormulaHosts(document),
                    "The failed numbered source formula was not restored by Word Undo.");
                AssertEqual(sourceNumberingBookmarkCount, CountInstalledVisualTeXNumberingBookmarks(document),
                    "The failed numbered source formula did not recover its VisualTeX numbering bookmarks.");
                AssertEqual(0, finalPlaceRefCount,
                    "The failed numbered source unexpectedly left a MathType MTPlaceRef field behind.");
                AssertTrue(finalVisualTeXCount > 0,
                    "Rollback acceptance unexpectedly converted every VisualTeX source formula.");
                AssertInstalledVisualTeXRawBridgeCountsUnchanged(document, sourceRawBridgeCounts);
                Console.WriteLine(
                    $"[FORMAT CONVERSION ROLLBACK] Injected failure after deleting the numbered VisualTeX host; Word Undo restored the source and all numbering artifacts. total={sourceFormulaCount}, finalVT={finalVisualTeXCount}, finalMT={finalMathTypeCount}.");
                return;
            }

            AssertEqual(0, finalVisualTeXCount,
                "Installed add-in left VisualTeX source OLE objects after conversion.");
            AssertEqual(initialMathTypeCount + sourceFormulaCount, finalMathTypeCount,
                "Installed add-in did not retain existing MathType objects while creating one Equation.DSMT4 object per VisualTeX source formula.");
            AssertEqual(initialMathTypePlaceRefCount + sourceNumberedCount, finalPlaceRefCount,
                "Installed add-in did not retain existing MathType numbering while recreating the VisualTeX numbered states as fresh MTPlaceRef rows.");
            AssertEqual(0, CountInstalledVisualTeXNumberingBookmarks(document),
                "Installed add-in left old VisualTeX VTEq/VTEqCap/VTEqNum bookmarks behind.");
            AssertEqual(0, CountInstalledTemporaryMathTypeBookmarks(document),
                "Installed add-in left temporary VTMT MathType identity bookmarks behind.");
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_EXPECT_ZERO_HEADING_PREFIX"),
                    "1",
                    StringComparison.Ordinal))
            {
                var expectedNumbers = Enumerable.Range(1, finalMathTypeCount)
                    .Select(index => $"(0.{index})")
                    .ToArray();
                AssertMathTypeNumberTexts(document, expectedNumbers);
                AssertNativeMathTypeSectionBreak(document, 0);
                Console.WriteLine(
                    $"[INSTALLED ADD-IN NUMBERING] headingless heading1 format preserved zero prefix: {string.Join(", ", expectedNumbers)}; MTEquationSection=0.");
            }

            Console.WriteLine(
                $"[INSTALLED ADD-IN] Real VisualTeX.WordVsto COM callback converted a real VisualTeX document fixture: formulas={sourceFormulaCount}, numbered={sourceNumberedCount}; all sources became MathType, old VisualTeX numbering artifacts were removed, and numbered state was recreated as fresh MTPlaceRef fields.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_ACCEPTANCE",
                previousAcceptanceMode);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                previousFormatConversionAcceptanceMode);
            if (localAddIn is not null && localCustom is not null)
            {
                try
                {
                    localAddIn.OnDisconnection(
                        ext_DisconnectMode.ext_dm_UserClosed,
                        ref localCustom);
                }
                catch { }
            }
            Release(callbacksObject);
            Release(installedAddIn);
            Release(addIns);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            if (fixtureDocument is not null)
            {
                try { fixtureDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            Release(fixtureDocument);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static OfficeSessionDocument CreateInstalledFormatSourceSession(
        string latex,
        bool numbered)
    {
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = Guid.NewGuid().ToString("D"),
            Title = "Installed format conversion source",
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = FormulaOleContract.NativeOleMode,
            Numbered = numbered,
            FontSizePt = 12,
            ExportResult = new OfficeExportDocument
            {
                Width = 260,
                Height = 96,
                Baseline = 72,
            },
        };
    }

    private static void DumpInstalledFormulaShapes(Word.Document document)
    {
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                var progId = "<none>";
                try { progId = shape.OLEFormat.ProgID ?? "<none>"; } catch { }
                var formulaId = "";
                var displayMode = "";
                var numbered = false;
                var latex = "";
                if (WordFormulaMetadataReader.IsNativeOle(shape))
                {
                    try
                    {
                        var metadata = WordFormulaMetadataReader.TryRead(shape);
                        formulaId = metadata?.FormulaId ?? "";
                        displayMode = metadata?.DisplayMode ?? "";
                        numbered = metadata?.Numbered == true;
                        latex = (metadata?.Latex ?? "").Replace("\r", " ").Replace("\n", " ");
                    }
                    catch { }
                }
                Console.WriteLine(
                    $"[INSTALLED ADD-IN SHAPE] index={index} range={shape.Range.Start}:{shape.Range.End} progId={progId} visualTeXId={formulaId} display={displayMode} numbered={numbered} latex={latex}");
            }
            finally { Release(shape); }
        }
    }

    private static Dictionary<string, int> ReadInstalledVisualTeXRawBridgeCounts(
        Word.Document document)
    {
        var markers = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                var metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata is null || string.IsNullOrWhiteSpace(metadata.Latex)) continue;
                var latex = metadata.Latex.Replace("\r\n", "\n").Replace('\r', '\n').Trim();
                markers.Add(string.Equals(metadata.DisplayMode, "block", StringComparison.OrdinalIgnoreCase)
                    ? "$$" + latex + "$$"
                    : "$" + latex.Replace('\n', ' ') + "$");
            }
            finally { Release(shape); }
        }

        var text = document.Content.Text ?? string.Empty;
        return markers.ToDictionary(
            marker => marker,
            marker => CountOrdinalOccurrences(text, marker),
            StringComparer.Ordinal);
    }

    private static void AssertInstalledVisualTeXRawBridgeCountsUnchanged(
        Word.Document document,
        IReadOnlyDictionary<string, int> baseline)
    {
        var text = document.Content.Text ?? string.Empty;
        foreach (var pair in baseline)
        {
            AssertEqual(
                pair.Value,
                CountOrdinalOccurrences(text, pair.Key),
                $"Rollback left a temporary VisualTeX LaTeX bridge in the document: {pair.Key}");
        }
    }

    private static int CountOrdinalOccurrences(string text, string value)
    {
        if (string.IsNullOrEmpty(value)) return 0;
        var count = 0;
        var start = 0;
        while (start <= text.Length - value.Length)
        {
            var index = text.IndexOf(value, start, StringComparison.Ordinal);
            if (index < 0) break;
            count++;
            start = index + value.Length;
        }
        return count;
    }

    private static int CountInstalledVisualTeXOleShapes(Word.Document document)
    {
        var count = 0;
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                if (WordFormulaMetadataReader.IsNativeOle(shape)) count++;
            }
            finally { Release(shape); }
        }
        return count;
    }

    private static int CountInstalledMathTypeOleShapes(Word.Document document)
    {
        var count = 0;
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                if (MathTypeOleInterop.IsMathTypeOle(shape)) count++;
            }
            finally { Release(shape); }
        }
        return count;
    }

    private static int CountInstalledTemporaryMathTypeBookmarks(Word.Document document)
    {
        var count = 0;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmark);
                bookmark = bookmarks[index];
                if (bookmark.Name.StartsWith("VTMT_", StringComparison.Ordinal))
                    count++;
            }
            return count;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static int CountInstalledVisualTeXNumberingBookmarks(Word.Document document)
    {
        var count = 0;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmark);
                bookmark = bookmarks[index];
                var name = bookmark.Name;
                if (name.StartsWith("VTEq_", StringComparison.Ordinal)
                    || name.StartsWith("VTEqCap_", StringComparison.Ordinal)
                    || name.StartsWith("VTEqNum_", StringComparison.Ordinal))
                    count++;
            }
            return count;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static int CountInstalledVisualTeXNumberedFormulaHosts(Word.Document document)
    {
        var count = 0;
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                var metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata?.Numbered == true) count++;
            }
            finally { Release(shape); }
        }
        return count;
    }
}
