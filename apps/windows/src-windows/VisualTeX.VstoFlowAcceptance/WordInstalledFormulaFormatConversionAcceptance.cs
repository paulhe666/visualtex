using System.IO.Compression;
using System.Xml.Linq;
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
    private static void RunWordDoc4VisualTeXSourceFixture(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Environment.GetEnvironmentVariable("VISUALTEX_DOC4_SOURCE_PATH");
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
            throw new FileNotFoundException(
                "Doc4 source-fixture acceptance requires VISUALTEX_DOC4_SOURCE_PATH.",
                sourcePath);

        var bulkLogPath = Path.Combine(artifactRoot, "doc4-bulk-import.log");
        var outputPath = Path.Combine(artifactRoot, "doc4-visualtex-source.docx");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", sourcePath);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", "latex");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", "ole");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", bulkLogPath);

        Word.Application? application = null;
        Word.Document? document = null;
        ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            addIn = new ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);
            addIn.OnBulkImport(new object());
            WaitForBulkImportCompletion(bulkLogPath, TimeSpan.FromMinutes(3));
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));

            var visualTeXCount = CountInstalledVisualTeXOleShapes(document);
            AssertEqual(16, visualTeXCount,
                "The doc4 source fixture did not reproduce the 16 VisualTeX formulas in the exact LaTeX block supplied by the user.");
            AssertEqual(0, CountInstalledMathTypeOleShapes(document),
                "The doc4 source fixture unexpectedly contains MathType OLE objects before conversion.");
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                $"[DOC4 SOURCE FIXTURE] VisualTeX bulk-import path created VT={visualTeXCount}, MT=0. path={outputPath}");
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", null);
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); }
                catch { }
            }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordLiveFormatConversionFixtureCapture(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var requestedName = Environment.GetEnvironmentVariable(
            "VISUALTEX_FORMAT_CONVERSION_LIVE_SOURCE_NAME");
        if (string.IsNullOrWhiteSpace(requestedName)) requestedName = "文档5";

        Word.Application? application = null;
        Word.Documents? documents = null;
        Word.Document? document = null;
        Word.Range? content = null;
        try
        {
            application = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject(
                "Word.Application");
            documents = application.Documents;
            for (var index = 1; index <= documents.Count; index++)
            {
                Word.Document? candidate = null;
                try
                {
                    candidate = documents[index];
                    if (!string.Equals(candidate.Name, requestedName, StringComparison.Ordinal))
                        continue;
                    document = candidate;
                    candidate = null;
                    break;
                }
                finally { Release(candidate); }
            }
            if (document is null)
                throw new FileNotFoundException(
                    $"The currently running Word instance does not contain '{requestedName}'.");

            var sourceFormulaCount = CountInstalledVisualTeXOleShapes(document);
            var sourceNumberedCount = CountInstalledVisualTeXNumberedFormulaHosts(document);
            var sourceMathTypeCount = CountInstalledMathTypeOleShapes(document);
            content = document.Content;
            var flatOpc = content.WordOpenXML;
            var fixturePath = Path.Combine(
                artifactRoot,
                "VisualTeX-Live-Format-Conversion-Source.docx");
            WriteFlatOpcPackage(flatOpc, fixturePath);

            var package = XDocument.Parse(flatOpc, LoadOptions.PreserveWhitespace);
            XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
            var embeddingCount = package.Descendants(pkg + "part")
                .Count(part => ((string?)part.Attribute(pkg + "name") ?? string.Empty)
                    .StartsWith("/word/embeddings/", StringComparison.OrdinalIgnoreCase));
            Console.WriteLine(
                $"[LIVE FORMAT FIXTURE] source={document.Name} saved={document.Saved} "
                + $"VT={sourceFormulaCount} VTNumbered={sourceNumberedCount} MT={sourceMathTypeCount} "
                + $"embeddings={embeddingCount} content={content.Start}:{content.End} path={fixturePath}");
        }
        finally
        {
            Release(content);
            Release(document);
            Release(documents);
            Release(application);
        }
    }

    private static void RunWordLiveUnsavedBackup(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Documents? documents = null;
        try
        {
            application = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject(
                "Word.Application");
            documents = application.Documents;
            var backupIndex = 0;
            for (var index = 1; index <= documents.Count; index++)
            {
                Word.Document? document = null;
                Word.Range? content = null;
                try
                {
                    document = documents[index];
                    if (document.Saved) continue;
                    backupIndex++;
                    content = document.Content;
                    var flatOpc = content.WordOpenXML;
                    var path = Path.Combine(
                        artifactRoot,
                        $"unsaved-word-{backupIndex:D2}.docx");
                    WriteFlatOpcPackage(flatOpc, path);
                    if (!File.Exists(path) || new FileInfo(path).Length <= 0)
                        throw new InvalidDataException(
                            $"Unsaved Word backup #{backupIndex} was not written.");
                    Console.WriteLine(
                        $"[UNSAVED WORD BACKUP] index={backupIndex} source={document.Name} inline={document.InlineShapes.Count} tables={document.Tables.Count} path={path}");
                }
                finally
                {
                    Release(content);
                    Release(document);
                }
            }
            if (backupIndex == 0)
                Console.WriteLine("[UNSAVED WORD BACKUP] No unsaved documents were open.");
            else
                Console.WriteLine($"[UNSAVED WORD BACKUP] completed={backupIndex}");
        }
        finally
        {
            Release(documents);
            Release(application);
        }
    }

    private static void RunWordLiveMathTypeDump(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var requestedName = Environment.GetEnvironmentVariable(
            "VISUALTEX_FORMAT_CONVERSION_LIVE_SOURCE_NAME");
        if (string.IsNullOrWhiteSpace(requestedName)) requestedName = "文档4";

        Word.Application? application = null;
        Word.Documents? documents = null;
        Word.Document? document = null;
        try
        {
            application = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject(
                "Word.Application");
            documents = application.Documents;
            for (var index = 1; index <= documents.Count; index++)
            {
                Word.Document? candidate = null;
                try
                {
                    candidate = documents[index];
                    if (!string.Equals(candidate.Name, requestedName, StringComparison.Ordinal))
                        continue;
                    document = candidate;
                    candidate = null;
                    break;
                }
                finally { Release(candidate); }
            }
            if (document is null)
                throw new FileNotFoundException(
                    $"The currently running Word instance does not contain '{requestedName}'.");

            Console.WriteLine(
                $"[LIVE MATHTYPE DUMP] source={document.Name} inline={document.InlineShapes.Count}");
            var mathTypeIndex = 0;
            for (var index = 1; index <= document.InlineShapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                Word.Range? context = null;
                try
                {
                    shape = document.InlineShapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                    mathTypeIndex++;
                    var mathMl = MathTypeOleStorage.ReadMathMl(shape);
                    var latex = MathMlToLatexConverter.Convert(mathMl)
                        .Replace("\r", " ")
                        .Replace("\n", " ")
                        .Trim();
                    var contextStart = Math.Max(0, shape.Range.Start - 36);
                    var contextEnd = Math.Min(document.Content.End, shape.Range.End + 36);
                    context = document.Range(contextStart, contextEnd);
                    var contextText = (context.Text ?? string.Empty)
                        .Replace('\r', ' ')
                        .Replace('\a', ' ')
                        .Replace('\n', ' ')
                        .Trim();
                    Console.WriteLine(
                        $"[LIVE MATHTYPE #{mathTypeIndex}] inlineIndex={index} range={shape.Range.Start}:{shape.Range.End} latex={latex}");
                    Console.WriteLine($"[LIVE MATHTYPE #{mathTypeIndex} CONTEXT] {contextText}");
                    Console.WriteLine($"[LIVE MATHTYPE #{mathTypeIndex} MATHML] {mathMl}");
                }
                catch (Exception error)
                {
                    Console.WriteLine(
                        $"[LIVE MATHTYPE #{mathTypeIndex}] inlineIndex={index} READ_ERROR={error.GetType().Name}: {error.Message}");
                }
                finally
                {
                    Release(context);
                    Release(shape);
                }
            }
        }
        finally
        {
            Release(document);
            Release(documents);
            Release(application);
        }
    }

    private static void WriteFlatOpcPackage(string flatOpc, string targetPath)
    {
        if (string.IsNullOrWhiteSpace(flatOpc))
            throw new InvalidDataException("Word returned empty Flat OPC for the live fixture.");
        XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        var package = XDocument.Parse(flatOpc, LoadOptions.PreserveWhitespace);
        var parts = package.Descendants(pkg + "part").ToArray();
        if (parts.Length == 0)
            throw new InvalidDataException("Word Flat OPC contains no package parts.");

        if (File.Exists(targetPath)) File.Delete(targetPath);
        using var file = new FileStream(targetPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: false);

        XNamespace contentTypesNamespace =
            "http://schemas.openxmlformats.org/package/2006/content-types";
        var contentTypes = new XElement(contentTypesNamespace + "Types");
        foreach (var part in parts)
        {
            var packageName = ((string?)part.Attribute(pkg + "name") ?? string.Empty).Trim();
            var contentType = ((string?)part.Attribute(pkg + "contentType") ?? string.Empty).Trim();
            if (packageName.Length == 0 || packageName == "/" || contentType.Length == 0)
                continue;
            contentTypes.Add(new XElement(
                contentTypesNamespace + "Override",
                new XAttribute("PartName", packageName),
                new XAttribute("ContentType", contentType)));
        }
        var contentTypesEntry = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
        using (var contentTypesOutput = contentTypesEntry.Open())
        using (var contentTypesWriter = new StreamWriter(
                   contentTypesOutput,
                   new System.Text.UTF8Encoding(false),
                   4096,
                   leaveOpen: true))
        {
            contentTypesWriter.Write(contentTypes.ToString(SaveOptions.DisableFormatting));
        }

        foreach (var part in parts)
        {
            var packageName = ((string?)part.Attribute(pkg + "name") ?? string.Empty).Trim();
            if (packageName.Length == 0 || packageName == "/") continue;
            var entryName = packageName.TrimStart('/');
            var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using var output = entry.Open();
            var binary = part.Element(pkg + "binaryData");
            if (binary is not null)
            {
                var bytes = Convert.FromBase64String(binary.Value);
                output.Write(bytes, 0, bytes.Length);
                continue;
            }

            var xmlData = part.Element(pkg + "xmlData")
                ?? throw new InvalidDataException(
                    $"Flat OPC part '{packageName}' contains neither xmlData nor binaryData.");
            var root = xmlData.Elements().FirstOrDefault()
                ?? throw new InvalidDataException(
                    $"Flat OPC XML part '{packageName}' has no root element.");
            using var writer = new StreamWriter(output, new System.Text.UTF8Encoding(false), 4096, leaveOpen: true);
            writer.Write(root.ToString(SaveOptions.DisableFormatting));
            writer.Flush();
        }
    }

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
        var requireNoMathTypeRuntime = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_EXPECT_NO_MATHTYPE_RUNTIME"),
            "1",
            StringComparison.Ordinal);
        var requireNoUnknownGlyphs = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_EXPECT_NO_MATHTYPE_UNKNOWN_GLYPHS"),
            "1",
            StringComparison.Ordinal);
        var requireVisibleMathTypePreviews = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_EXPECT_MATHTYPE_VISIBLE_PREVIEWS"),
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
            var sourceHbarCount = CountInstalledVisualTeXHbarOccurrences(document);
            var initialMathTypeCount = CountInstalledMathTypeOleShapes(document);
            var initialMathTypePlaceRefCount = CountMathTypePlaceRefFields(document);
            var initialMathTypeNativeHbarCount = CountInstalledMathTypeNativeHbarCharacters(document);
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
                $"[INSTALLED ADD-IN SOURCE COUNTS] VT={sourceFormulaCount} VTNumbered={sourceNumberedCount} VTHbar={sourceHbarCount} existingMT={initialMathTypeCount} existingMTPlaceRef={initialMathTypePlaceRefCount} existingMTHbar={initialMathTypeNativeHbarCount}");
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
            var conversionMathTypeProcessesBefore = requireNoMathTypeRuntime
                ? SnapshotMathTypeProcessIds()
                : new HashSet<int>();
            if (requireNoMathTypeRuntime && conversionMathTypeProcessesBefore.Count != 0)
                throw new InvalidOperationException(
                    "No-MathType-runtime acceptance requires MathType.exe to be absent before conversion.");

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
                if (requireNoMathTypeRuntime)
                {
                    var startedMathType = SnapshotMathTypeProcessIds()
                        .Except(conversionMathTypeProcessesBefore)
                        .ToArray();
                    if (startedMathType.Length > 0)
                        throw new InvalidOperationException(
                            "VisualTeX format conversion started MathType.exe: "
                            + string.Join(", ", startedMathType));
                }
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
            if (requireNoUnknownGlyphs)
                AssertNoUnknownMathTypeGlyphTokens(document, "after conversion");
            if (requireVisibleMathTypePreviews)
                AssertVisibleMathTypeEmfPreviews(document, "after conversion");
            AssertEqual(
                initialMathTypeNativeHbarCount + sourceHbarCount,
                CountInstalledMathTypeNativeHbarCharacters(document),
                "Installed add-in did not persist every VisualTeX \\hbar using MathType 7's native MT Extra + encoded8 MTEF character record.");

            var persistedPath = document.FullName;
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                persistedPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            for (var settle = 0; settle < 10; settle++)
            {
                WinForms.Application.DoEvents();
                Thread.Sleep(100);
            }
            AssertEqual(0, CountInstalledVisualTeXOleShapes(document),
                "Saved/reopened installed-addin fixture restored a VisualTeX source OLE.");
            AssertEqual(initialMathTypeCount + sourceFormulaCount, CountInstalledMathTypeOleShapes(document),
                "Saved/reopened installed-addin fixture lost a converted MathType OLE.");
            AssertEqual(initialMathTypePlaceRefCount + sourceNumberedCount, CountMathTypePlaceRefFields(document),
                "Saved/reopened installed-addin fixture changed MathType numbering fields.");
            AssertEqual(0, CountInstalledVisualTeXNumberingBookmarks(document),
                "Saved/reopened installed-addin fixture restored old VisualTeX numbering bookmarks.");
            AssertEqual(0, CountInstalledTemporaryMathTypeBookmarks(document),
                "Saved/reopened installed-addin fixture retained a temporary VTMT bookmark.");
            if (requireNoUnknownGlyphs)
                AssertNoUnknownMathTypeGlyphTokens(document, "after save/reopen");
            if (requireVisibleMathTypePreviews)
                AssertVisibleMathTypeEmfPreviews(document, "after save/reopen");
            if (requireNoMathTypeRuntime)
            {
                var finalMathTypeProcesses = SnapshotMathTypeProcessIds()
                    .Except(conversionMathTypeProcessesBefore)
                    .ToArray();
                AssertEqual(0, finalMathTypeProcesses.Length,
                    "VisualTeX conversion left a MathType.exe process running.");
                Console.WriteLine("[NO MATHTYPE RUNTIME] Conversion and save/reopen completed with MathTypeProcessCount=0.");
            }
            AssertEqual(
                initialMathTypeNativeHbarCount + sourceHbarCount,
                CountInstalledMathTypeNativeHbarCharacters(document),
                "Saved/reopened installed-addin fixture lost or changed MathType 7 native \\hbar MTEF records.");
            for (var index = 1; index <= document.InlineShapes.Count; index++)
            {
                Word.InlineShape? persistedShape = null;
                try
                {
                    persistedShape = document.InlineShapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(persistedShape)) continue;
                    var persistedMathMl = MathTypeOleStorage.ReadMathMl(persistedShape);
                    AssertTrue(!string.IsNullOrWhiteSpace(persistedMathMl),
                        $"Saved/reopened MathType OLE #{index} has no readable MTEF/MathML payload.");
                }
                finally { Release(persistedShape); }
            }
            Console.WriteLine(
                $"[INSTALLED ADD-IN PERSISTENCE] Saved/reopened DOCX retained MT={CountInstalledMathTypeOleShapes(document)}, MTPlaceRef={CountMathTypePlaceRefFields(document)}, VT=0 and every MathType OLE remained readable.");

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

    private static void AssertNoUnknownMathTypeGlyphTokens(
        Word.Document document,
        string stage)
    {
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                var mathMl = MathTypeOleStorage.ReadMathMl(shape);
                var parsed = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
                var badTokens = parsed.Descendants()
                    .Where(element => element.Name.LocalName is "mi" or "mo" or "mn" or "mtext")
                    .Select(element => element.Value)
                    .Where(value => value.IndexOf('\uFFFD') >= 0 || value.Trim() == "?")
                    .ToArray();
                if (badTokens.Length > 0)
                    throw new InvalidDataException(
                        $"MathType OLE #{index} contains unknown/replacement glyph tokens {stage}: "
                        + string.Join(", ", badTokens.Select(value =>
                            value.IndexOf('\uFFFD') >= 0 ? "U+FFFD" : "?"))
                        + $". MathML='{mathMl}'");
            }
            finally { Release(shape); }
        }
        Console.WriteLine($"[NO UNKNOWN GLYPHS] Every MathType Equation Native payload is clean {stage}.");
    }

    private static void AssertVisibleMathTypeEmfPreviews(
        Word.Document document,
        string stage)
    {
        var mathTypeCount = 0;
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            Word.Range? range = null;
            try
            {
                shape = document.InlineShapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                mathTypeCount++;
                AssertTrue(shape.Width > 1f && shape.Height > 1f,
                    $"MathType OLE #{index} has no visible geometry {stage}: {shape.Width:0.###}x{shape.Height:0.###} pt.");

                range = shape.Range;
                var flatOpc = range.WordOpenXML;
                var hasExternalMetafile =
                    (flatOpc.IndexOf("image/x-wmf", StringComparison.OrdinalIgnoreCase) >= 0
                        && flatOpc.IndexOf(".wmf", StringComparison.OrdinalIgnoreCase) >= 0)
                    || (flatOpc.IndexOf("image/x-emf", StringComparison.OrdinalIgnoreCase) >= 0
                        && flatOpc.IndexOf(".emf", StringComparison.OrdinalIgnoreCase) >= 0);
                AssertTrue(hasExternalMetafile,
                    $"MathType OLE #{index} has no Word metafile presentation {stage}.");

                var fragment = MathTypeWordOpenXml.Read(flatOpc);
                var storageEntries = MathTypeOleStorage.ListCompoundFileEntries(fragment.CompoundFile);
                var olePresentationEntries = storageEntries
                    .Where(name => name.IndexOf("OlePres", StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToArray();
                AssertEqual(0, olePresentationEntries.Length,
                    $"MathType OLE #{index} contains an internal OlePres stream {stage}; genuine MathType Equation.DSMT4 storage leaves presentation ownership to Word.");

                var mathTypeProcessesBeforePreview = SnapshotMathTypeProcessIds();
                var preview = ReadInlineShapeEnhancedMetafile(shape);
                var startedDuringPreview = SnapshotMathTypeProcessIds()
                    .Except(mathTypeProcessesBeforePreview)
                    .ToArray();
                AssertEqual(0, startedDuringPreview.Length,
                    $"Reading the live Word preview for MathType OLE #{index} started MathType.exe {stage}.");
                var ink = DescribeEmfInkBounds(preview);
                AssertTrue(!string.Equals(ink, "empty", StringComparison.Ordinal),
                    $"MathType OLE #{index} copied from Word as an empty live preview {stage}.");
                Console.WriteLine(
                    $"[VISIBLE MATHTYPE PREVIEW] #{index} {stage}: {shape.Width:0.###}x{shape.Height:0.###} pt; {ink}");
            }
            finally
            {
                Release(range);
                Release(shape);
            }
        }
        AssertTrue(mathTypeCount > 0,
            $"No MathType OLE objects were available for live-preview validation {stage}.");
        Console.WriteLine(
            $"[VISIBLE MATHTYPE PREVIEW] Every MathType OLE used a non-empty Word metafile presentation with no internal OlePres stream {stage}; count={mathTypeCount}.");
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

    private static int CountInstalledVisualTeXHbarOccurrences(Word.Document document)
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
                count += CountOrdinalOccurrences(metadata?.Latex ?? string.Empty, @"\hbar");
            }
            finally { Release(shape); }
        }
        return count;
    }

    private static int CountInstalledMathTypeNativeHbarCharacters(Word.Document document)
    {
        // Genuine MathType 7 persists \hbar as:
        // CHAR, CharEncoded8, fnMTEXTRA(11), MTCode U+210F, encoded8 'h'.
        var pattern = new byte[] { 0x02, 0x04, 0x8B, 0x0F, 0x21, 0x68 };
        var count = 0;
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                var compoundFile = MathTypeOleStorage.CaptureCompoundFile(shape);
                var equationNative = MathTypeOleStorage.ReadEquationNative(compoundFile);
                for (var offset = 0; offset <= equationNative.Length - pattern.Length; offset++)
                {
                    var matches = true;
                    for (var probe = 0; probe < pattern.Length; probe++)
                    {
                        if (equationNative[offset + probe] == pattern[probe]) continue;
                        matches = false;
                        break;
                    }
                    if (matches) count++;
                }
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
