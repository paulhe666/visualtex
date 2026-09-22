using System.Text;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class DocumentChangeObservation
    {
        public string ActiveDocumentName { get; set; } = string.Empty;
        public bool ScreenUpdating { get; set; }
        public int SelectionStart { get; set; }
        public int SelectionEnd { get; set; }
    }

    private static void RunWordOmmlNativeEditToVisualTeXAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The native-edit OMML→VisualTeX acceptance refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"word-omml-native-edit-vt-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var pngPath = Path.Combine(assetRoot, "preview.png");
        var svgPath = Path.Combine(assetRoot, "preview.svg");
        WriteAcceptancePng(pngPath, "native-edit", 360, 112);
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"360\" height=\"112\" viewBox=\"0 0 360 112\"><rect width=\"360\" height=\"112\" fill=\"white\"/><text x=\"12\" y=\"74\" font-family=\"Cambria Math\" font-size=\"42\">x+1</text></svg>",
            new UTF8Encoding(false));
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 360, 112);
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></math>";
        var previousNativePreview = Environment.GetEnvironmentVariable(
            "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW");
        try
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                "1");
            using var host = new WordPerformanceHost(documentPath: null);
            var service = new WordFormulaService(host.Application);

            for (var index = 0; index < 2; index++)
            {
                SelectDocumentEnd(host.Document);
                var sourceSession = CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: false,
                    FormulaOleContract.MathTypeOleMode);
                service.InsertMathTypeOle(
                    sourceSession,
                    mathMl,
                    emfPath,
                    updateCreatedMathTypeNumberFields: true);
                if (index == 0)
                    AppendAcceptanceText(host.Document, "\r");
            }
            AssertEqual(2, CountMathTypeOleShapes(host.Document),
                "Native-edit regression setup did not create two identical MathType sources.");

            var toOmmlPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(2, toOmmlPlan.Targets.Count,
                "MathType→OMML setup did not capture both identical MathType sources.");
            var toOmmlPrepared = PrepareOmmlMathTypeTargets(toOmmlPlan, emfPath);
            var toOmmlResult = service.ApplyFormulaFormatConversionPlan(
                toOmmlPlan,
                toOmmlPrepared);
            AssertEqual(2, toOmmlResult.FormulaCount,
                "MathType→OMML setup did not convert both formulas.");
            AssertEqual(0, toOmmlResult.FailedFormulaCount,
                "MathType→OMML setup reported a failure: "
                + string.Join(" | ", toOmmlResult.Failures));
            AssertEqual(2, host.Document.OMaths.Count,
                "MathType→OMML setup did not leave two Word equations.");

            var ordered = WordOmmlFormulaStore.FormulaIds(host.Document)
                .Select(formulaId =>
                {
                    Word.Bookmark? bookmark = null;
                    Word.Range? range = null;
                    try
                    {
                        bookmark = WordOmmlFormulaStore.FindByFormulaId(
                            host.Document,
                            formulaId)
                            ?? throw new InvalidDataException(
                                $"Managed OMML bookmark {formulaId} is missing after MathType→OMML.");
                        range = WordOmmlFormulaStore.GetEquationRange(bookmark);
                        return (FormulaId: formulaId, Start: range.Start);
                    }
                    finally
                    {
                        Release(range);
                        Release(bookmark);
                    }
                })
                .OrderBy(item => item.Start)
                .ToArray();
            AssertEqual(2, ordered.Length,
                "MathType→OMML setup did not create exactly two managed identities.");
            var editedFormulaId = ordered[0].FormulaId;
            var siblingFormulaId = ordered[1].FormulaId;
            var stored = WordOmmlFormulaStore.TryRead(host.Document, editedFormulaId)
                ?? throw new InvalidDataException("Edited OMML source metadata is missing.");

            Word.Bookmark? editedBookmark = null;
            Word.Range? editedRange = null;
            Word.OMaths? maths = null;
            Word.OMath? math = null;
            Word.Range? linearRange = null;
            Word.Range? tokenRange = null;
            Word.Range? liveEditedRange = null;
            Word.Bookmarks? bookmarks = null;
            Word.Bookmark? driftedBookmark = null;
            try
            {
                editedBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    host.Document,
                    editedFormulaId)
                    ?? throw new InvalidDataException("Edited OMML source bookmark is missing.");
                editedRange = WordOmmlFormulaStore.GetEquationRange(editedBookmark);
                maths = editedRange.OMaths;
                AssertEqual(1, maths.Count,
                    "Edited OMML source does not contain exactly one OMath.");
                math = maths[1];
                math.Linearize();
                Release(linearRange); linearRange = math.Range.Duplicate;
                var linearText = linearRange.Text ?? string.Empty;
                const string MathematicalItalicX = "\U0001D465";
                var tokenOffset = linearText.IndexOf(
                    MathematicalItalicX,
                    StringComparison.Ordinal);
                var tokenLength = MathematicalItalicX.Length;
                if (tokenOffset < 0)
                {
                    tokenOffset = linearText.IndexOf('x');
                    tokenLength = 1;
                }
                AssertTrue(tokenOffset >= 0,
                    "Linearized Word equation does not expose the expected x token.");
                tokenRange = host.Document.Range(
                    linearRange.Start + tokenOffset,
                    linearRange.Start + tokenOffset + tokenLength);
                tokenRange.Text = "y";
                Release(tokenRange); tokenRange = null;
                try { math.BuildUp(); }
                catch
                {
                    Release(math); math = null;
                    Release(maths); maths = null;
                    maths = host.Document.OMaths;
                    AssertEqual(2, maths.Count,
                        "Native token edit changed the total OMath inventory.");
                    math = maths[1];
                    math.BuildUp();
                }

                Release(liveEditedRange); liveEditedRange = null;
                Release(math); math = null;
                Release(maths); maths = null;
                maths = host.Document.OMaths;
                AssertEqual(2, maths.Count,
                    "Native BuildUp changed the total OMath inventory.");
                math = maths[1];
                liveEditedRange = math.Range.Duplicate;
                var liveXml = WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                    host.Document,
                    liveEditedRange,
                    editedFormulaId);
                AssertTrue(
                    liveXml.IndexOf("y", StringComparison.OrdinalIgnoreCase) >= 0,
                    "The Word-native x→y edit did not survive BuildUp.");
                AssertTrue(
                    !WordOmmlConverter.MatchesStoredOmmlFingerprint(
                        liveXml,
                        stored.NativeOmmlFingerprint),
                    "Native edit did not change the stored OMML fingerprint.");

                // Word 2021 can leave a collapsed VTOMML bookmark at the OMath end
                // after Linearize/token editing/BuildUp. Normalize the fixture to
                // that exact observed post-edit state when this particular build
                // happens to retain the old canonical start anchor, so the regression
                // remains deterministic across Office patch levels.
                Release(editedBookmark); editedBookmark = null;
                editedBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    host.Document,
                    editedFormulaId);
                var recreateDriftedAnchor = editedBookmark is null
                    || WordOmmlFormulaStore.IsCanonicalAnchor(
                        editedBookmark,
                        liveEditedRange);
                if (recreateDriftedAnchor)
                {
                    var bookmarkName = WordOmmlFormulaStore.BookmarkName(
                        editedFormulaId);
                    if (editedBookmark is not null)
                    {
                        editedBookmark.Delete();
                        Release(editedBookmark); editedBookmark = null;
                    }
                    var end = liveEditedRange.End;
                    var driftRange = host.Document.Range(end, end);
                    try
                    {
                        bookmarks = host.Document.Bookmarks;
                        driftedBookmark = bookmarks.Add(bookmarkName, driftRange);
                    }
                    finally { Release(driftRange); }
                }

                // The second untouched formula intentionally still matches the old
                // fingerprint. A correct repair must use the explicitly selected
                // local OMath and must never jump to this global fingerprint decoy.
                var siblingMetadata = WordOmmlFormulaStore.TryRead(
                    host.Document,
                    siblingFormulaId)
                    ?? throw new InvalidDataException("Sibling OMML metadata is missing.");
                Word.Bookmark? siblingBookmark = null;
                Word.Range? siblingRange = null;
                try
                {
                    siblingBookmark = WordOmmlFormulaStore.FindByFormulaId(
                        host.Document,
                        siblingFormulaId)
                        ?? throw new InvalidDataException("Sibling OMML bookmark is missing.");
                    siblingRange = WordOmmlFormulaStore.GetEquationRange(siblingBookmark);
                    var siblingXml = WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                        host.Document,
                        siblingRange,
                        siblingFormulaId);
                    AssertTrue(
                        WordOmmlConverter.MatchesStoredOmmlFingerprint(
                            siblingXml,
                            stored.NativeOmmlFingerprint),
                        "The identical sibling no longer provides the required stale-fingerprint decoy.");
                }
                finally
                {
                    Release(siblingRange);
                    Release(siblingBookmark);
                }

                liveEditedRange.Select();
                var toVisualTeXPlan = service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: false,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.NativeOleMode);
                AssertEqual(1, toVisualTeXPlan.Targets.Count,
                    "Selected edited OMML was not captured as one conversion target.");
                var target = toVisualTeXPlan.Targets[0];
                AssertTrue(target.SourceIsManagedOmml,
                    "Selected edited OMML was demoted to unmanaged native OMath after bookmark/fingerprint drift.");
                AssertEqual(editedFormulaId, target.SourceFormulaId,
                    "Selected edited OMML rebound to the wrong managed FormulaId.");
                AssertTrue(
                    target.Latex.IndexOf("y", StringComparison.OrdinalIgnoreCase) >= 0,
                    "OMML→VisualTeX capture lost the Word-native y edit.");

                var prepared = PrepareOmmlVisualTeXStressTargets(
                    toVisualTeXPlan,
                    pngPath,
                    emfPath);
                var conversion = service.ApplyFormulaFormatConversionPlan(
                    toVisualTeXPlan,
                    prepared);
                AssertEqual(1, conversion.FormulaCount,
                    "Edited OMML→VisualTeX did not convert exactly one selected formula.");
                AssertEqual(0, conversion.FailedFormulaCount,
                    "Edited OMML→VisualTeX reported a failure: "
                    + string.Join(" | ", conversion.Failures));
                AssertEqual(1, host.Document.OMaths.Count,
                    "Edited OMML→VisualTeX deleted or retained the wrong native OMath count.");
                AssertEqual(1, CountVisualTeXNativeOleShapes(host.Document),
                    "Edited OMML→VisualTeX did not create exactly one VisualTeX OLE.");

                Release(driftedBookmark); driftedBookmark = null;
                Release(bookmarks); bookmarks = null;
                var stale = WordOmmlFormulaStore.FindByFormulaId(
                    host.Document,
                    editedFormulaId);
                try
                {
                    AssertTrue(stale is null,
                        "Edited OMML→VisualTeX left the old VTOMML bookmark behind.");
                }
                finally { Release(stale); }
                AssertTrue(
                    WordOmmlFormulaStore.TryRead(
                        host.Document,
                        editedFormulaId) is null,
                    "Edited OMML→VisualTeX left stale OMML metadata behind.");
            }
            finally
            {
                Release(driftedBookmark);
                Release(bookmarks);
                Release(liveEditedRange);
                Release(tokenRange);
                Release(linearRange);
                Release(math);
                Release(maths);
                Release(editedRange);
                Release(editedBookmark);
            }

            var savedPath = Path.Combine(
                artifactRoot,
                "Word-OMML-Native-Edit-To-VisualTeX.docx");
            host.Save(savedPath);
            host.Document.Close(Word.WdSaveOptions.wdSaveChanges);
            Word.Document? reopened = null;
            try
            {
                reopened = host.Application.Documents.Open(
                    savedPath,
                    ConfirmConversions: false,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false,
                    OpenAndRepair: false);
                reopened.Activate();
                AssertEqual(1, reopened.OMaths.Count,
                    "Save/reopen changed the untouched sibling OMML count.");
                AssertEqual(1, CountVisualTeXNativeOleShapes(reopened),
                    "Save/reopen lost the converted VisualTeX OLE.");
            }
            finally
            {
                if (reopened is not null)
                    try { reopened.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                Release(reopened);
            }
            Console.WriteLine(
                "[OMML NATIVE EDIT→VISUALTEX] Selected local OMath rebound its managed identity after Word-native edit/bookmark drift, ignored the stale-fingerprint sibling, converted once, cleaned the old OMML identity, and survived save/reopen.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                previousNativePreview);
            try { Directory.Delete(assetRoot, recursive: true); } catch { }
        }
    }

    private static System.Drawing.Bitmap CaptureAcceptanceWordViewport(
        IntPtr wordWindowHandle,
        string path)
    {
        if (wordWindowHandle == IntPtr.Zero
            || !GetWindowRect(wordWindowHandle, out var rectangle))
            throw new InvalidOperationException(
                "The isolated Word window rectangle is unavailable for viewport capture.");

        var windowWidth = Math.Max(1, rectangle.Right - rectangle.Left);
        var windowHeight = Math.Max(1, rectangle.Bottom - rectangle.Top);
        // Exclude Word's title/ribbon/status UI. The remaining rectangle is the
        // document viewport where a top-of-document flash would actually be
        // visible to the user.
        var topInset = Math.Min(190, Math.Max(110, windowHeight / 5));
        var leftInset = Math.Min(36, Math.Max(12, windowWidth / 40));
        var bottomInset = Math.Min(56, Math.Max(24, windowHeight / 18));
        var width = Math.Max(64, windowWidth - (leftInset * 2));
        var height = Math.Max(64, windowHeight - topInset - bottomInset);
        var bitmap = new System.Drawing.Bitmap(
            width,
            height,
            System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.CopyFromScreen(
                rectangle.Left + leftInset,
                rectangle.Top + topInset,
                0,
                0,
                new System.Drawing.Size(width, height),
                System.Drawing.CopyPixelOperation.SourceCopy);
        }
        bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
        return bitmap;
    }

    private static double CalculateAcceptanceFrameDifference(
        System.Drawing.Bitmap left,
        System.Drawing.Bitmap right)
    {
        if (left.Width != right.Width || left.Height != right.Height)
            return 1.0;

        var rectangle = new System.Drawing.Rectangle(
            0,
            0,
            left.Width,
            left.Height);
        var leftData = left.LockBits(
            rectangle,
            System.Drawing.Imaging.ImageLockMode.ReadOnly,
            System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        var rightData = right.LockBits(
            rectangle,
            System.Drawing.Imaging.ImageLockMode.ReadOnly,
            System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        try
        {
            var byteCount = Math.Abs(leftData.Stride) * left.Height;
            var leftBytes = new byte[byteCount];
            var rightBytes = new byte[byteCount];
            System.Runtime.InteropServices.Marshal.Copy(
                leftData.Scan0,
                leftBytes,
                0,
                byteCount);
            System.Runtime.InteropServices.Marshal.Copy(
                rightData.Scan0,
                rightBytes,
                0,
                byteCount);

            long changedPixels = 0;
            var pixelCount = (long)left.Width * left.Height;
            for (var y = 0; y < left.Height; y++)
            {
                var leftRow = y * Math.Abs(leftData.Stride);
                var rightRow = y * Math.Abs(rightData.Stride);
                for (var x = 0; x < left.Width; x++)
                {
                    var leftOffset = leftRow + (x * 4);
                    var rightOffset = rightRow + (x * 4);
                    // Ignore tiny ClearType/antialiasing fluctuations. A document
                    // jump replaces thousands of text pixels and is orders of
                    // magnitude larger than this threshold.
                    if (Math.Abs(leftBytes[leftOffset] - rightBytes[rightOffset]) > 24
                        || Math.Abs(leftBytes[leftOffset + 1] - rightBytes[rightOffset + 1]) > 24
                        || Math.Abs(leftBytes[leftOffset + 2] - rightBytes[rightOffset + 2]) > 24)
                        changedPixels++;
                }
            }
            return pixelCount == 0
                ? 0
                : (double)changedPixels / pixelCount;
        }
        finally
        {
            left.UnlockBits(leftData);
            right.UnlockBits(rightData);
        }
    }

    private static void RunWordOmmlInsertNoTopFlashAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The OMML no-top-flash acceptance refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        using var host = new WordPerformanceHost(documentPath: null);
        var service = new WordFormulaService(host.Application);
        host.Application.Visible = true;
        host.Document.Activate();
        var body = new StringBuilder();
        for (var index = 1; index <= 260; index++)
            body.Append($"VIEW-LINE-{index:D3} VisualTeX OMML insertion viewport regression.\r");
        host.Document.Content.Text = body.ToString();

        var markerText = "VIEW-LINE-220";
        var documentText = host.Document.Content.Text ?? string.Empty;
        var markerOffset = documentText.IndexOf(markerText, StringComparison.Ordinal);
        AssertTrue(markerOffset >= 0,
            "No-top-flash fixture could not locate the deep-document marker.");
        Word.Range? insertion = null;
        Word.Window? window = null;
        System.Drawing.Bitmap? beforeFrame = null;
        System.Drawing.Bitmap? sourceActivationFrame = null;
        System.Drawing.Bitmap? afterFrame = null;
        var targetWindowHandle = IntPtr.Zero;
        var targetName = host.Document.Name;
        var beforeFramePath = Path.Combine(
            artifactRoot,
            "omml-insert-before.png");
        var sourceFramePath = Path.Combine(
            artifactRoot,
            "omml-insert-hidden-source-activation.png");
        var afterFramePath = Path.Combine(
            artifactRoot,
            "omml-insert-after.png");
        var guardTracePath = Path.Combine(
            artifactRoot,
            "omml-insert-screen-updating-guard.log");
        var previousTracePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_WORD_HOOK_TRACE_PATH");
        var observations = new List<DocumentChangeObservation>();
        void ObserveDocumentChange()
        {
            Word.Document? active = null;
            Word.Selection? selection = null;
            try
            {
                active = host.Application.ActiveDocument;
                selection = host.Application.Selection;
                var activeName = active?.Name ?? string.Empty;
                observations.Add(new DocumentChangeObservation
                {
                    ActiveDocumentName = activeName,
                    ScreenUpdating = host.Application.ScreenUpdating,
                    SelectionStart = selection?.Start ?? -1,
                    SelectionEnd = selection?.End ?? -1,
                });
                if (sourceActivationFrame is null
                    && targetWindowHandle != IntPtr.Zero
                    && !string.Equals(
                        activeName,
                        targetName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    // This callback is delivered synchronously as Word changes
                    // ActiveDocument, before the next paint. Capture the actual
                    // visible target viewport at the dangerous transition.
                    sourceActivationFrame = CaptureAcceptanceWordViewport(
                        targetWindowHandle,
                        sourceFramePath);
                }
            }
            catch
            {
                observations.Add(new DocumentChangeObservation
                {
                    ActiveDocumentName = "<unavailable>",
                    ScreenUpdating = host.Application.ScreenUpdating,
                    SelectionStart = -1,
                    SelectionEnd = -1,
                });
            }
            finally
            {
                Release(selection);
                Release(active);
            }
        }

        try
        {
            insertion = host.Document.Range(markerOffset, markerOffset);
            insertion.Select();
            window = host.Application.ActiveWindow;
            targetWindowHandle = new IntPtr(window.Hwnd);
            _ = SetForegroundWindow(targetWindowHandle);
            object start = true;
            window.ScrollIntoView(insertion, ref start);
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(250);
            var verticalBefore = window.VerticalPercentScrolled;
            var horizontalBefore = window.HorizontalPercentScrolled;
            var selectionBefore = host.Application.Selection.Start;
            AssertTrue(verticalBefore >= 25,
                $"No-top-flash fixture did not reach the middle/lower document; vertical={verticalBefore}%.");
            beforeFrame = CaptureAcceptanceWordViewport(
                targetWindowHandle,
                beforeFramePath);
            try { File.Delete(guardTracePath); } catch { }
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                guardTracePath);

            host.Application.DocumentChange += ObserveDocumentChange;
            try
            {
                var session = CreateNumberedOmmlTabSession(
                    Guid.NewGuid().ToString("D"),
                    host.Document.FullName,
                    insertion.Start,
                    insertion.End,
                    @"x^2+y^2=r^2",
                    originalMetadata: null);
                session.Numbered = false;
                service.InsertOmml(
                    session,
                    "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><msup><mi>y</mi><mn>2</mn></msup><mo>=</mo><msup><mi>r</mi><mn>2</mn></msup></mrow></math>");
            }
            finally
            {
                host.Application.DocumentChange -= ObserveDocumentChange;
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_WORD_HOOK_TRACE_PATH",
                    previousTracePath);
            }

            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(250);
            afterFrame = CaptureAcceptanceWordViewport(
                targetWindowHandle,
                afterFramePath);
            var verticalAfter = window.VerticalPercentScrolled;
            var horizontalAfter = window.HorizontalPercentScrolled;
            var selectionAfter = host.Application.Selection.Start;
            var sourceActivations = observations
                .Where(item =>
                    !string.Equals(
                        item.ActiveDocumentName,
                        targetName,
                        StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(
                        item.ActiveDocumentName,
                        "<unavailable>",
                        StringComparison.Ordinal))
                .ToArray();
            var beforeViewport = beforeFrame
                ?? throw new InvalidDataException(
                    "The pre-insertion viewport frame is missing.");
            double? sourceActivationFrameDifference = sourceActivationFrame is null
                ? null
                : CalculateAcceptanceFrameDifference(
                    beforeViewport,
                    sourceActivationFrame);
            var afterFrameDifference =
                CalculateAcceptanceFrameDifference(
                    beforeViewport,
                    afterFrame
                        ?? throw new InvalidDataException(
                            "The post-insertion viewport frame is missing."));
            var guardTrace = File.Exists(guardTracePath)
                ? File.ReadAllText(guardTracePath)
                : string.Empty;
            var guardReassertionCount = guardTrace
                .Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries)
                .Count(line => line.IndexOf(
                    "omml-screen-update-guard-reasserted",
                    StringComparison.Ordinal) >= 0);

            File.WriteAllText(
                Path.Combine(artifactRoot, "omml-insert-document-change-observations.json"),
                JsonSerializer.Serialize(
                    new
                    {
                        targetDocument = targetName,
                        verticalBefore,
                        verticalAfter,
                        horizontalBefore,
                        horizontalAfter,
                        selectionBefore,
                        selectionAfter,
                        sourceActivationFrameDifference,
                        afterFrameDifference,
                        guardReassertionCount,
                        frames = new
                        {
                            before = beforeFramePath,
                            hiddenSourceActivation = sourceFramePath,
                            after = afterFramePath,
                        },
                        observations,
                    },
                    new JsonSerializerOptions { WriteIndented = true }),
                new UTF8Encoding(false));

            AssertEqual(0, sourceActivations.Length,
                "The target Word Application still activated a temporary OMML source document.");
            AssertTrue(sourceActivationFrame is null,
                "A hidden-source activation frame was captured even though the target Word Application must never activate the temporary source.");
            AssertEqual(0, guardReassertionCount,
                "The screen-updating guard had to recover from a target-Application document switch; the direct normalized-OMML path should require no such recovery.");
            AssertTrue(verticalAfter > 5,
                $"OMML insertion ended at the document top: before={verticalBefore}%, after={verticalAfter}%.");
            AssertTrue(Math.Abs(verticalAfter - verticalBefore) <= 3,
                $"OMML insertion changed the vertical viewport too much: before={verticalBefore}%, after={verticalAfter}%.");
            AssertTrue(Math.Abs(horizontalAfter - horizontalBefore) <= 3,
                $"OMML insertion changed the horizontal viewport: before={horizontalBefore}%, after={horizontalAfter}%.");
            AssertTrue(
                selectionAfter >= selectionBefore - 4
                && selectionAfter <= selectionBefore + 256,
                $"OMML insertion moved Selection outside the local insertion neighborhood: before={selectionBefore}, after={selectionAfter}.");
            AssertEqual(1, host.Document.OMaths.Count,
                "No-top-flash acceptance did not create exactly one OMML formula.");

            host.Save(Path.Combine(
                artifactRoot,
                "Word-OMML-Insert-No-Top-Flash.docx"));
            Console.WriteLine(
                $"[OMML INSERT VIEW] Hidden-source activations={sourceActivations.Length}; target-Application reassertions={guardReassertionCount}; before/after viewport changed pixels={afterFrameDifference:P3}; vertical {verticalBefore}%→{verticalAfter}%.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                previousTracePath);
            afterFrame?.Dispose();
            sourceActivationFrame?.Dispose();
            beforeFrame?.Dispose();
            Release(window);
            Release(insertion);
        }
    }
}
