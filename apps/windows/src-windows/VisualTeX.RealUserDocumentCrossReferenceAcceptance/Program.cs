using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using WordRange = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.RealUserDocumentCrossReferenceAcceptance;

internal static class Program
{
    private const string VisualTeXXmlNamespace = "urn:visualtex:word-omml:1";
    private const string NativeCaptionPrefix = "VTEqCap_";
    private const string OmmlBookmarkPrefix = "VTOMML_";

    [STAThread]
    private static int Main()
    {
        Console.OutputEncoding = Encoding.UTF8;
        var root = FindWindowsRoot();
        var artifactRoot = Path.Combine(
            root,
            "src-windows",
            "artifacts",
            $"real-user-document-crossref-{DateTime.Now:yyyyMMdd-HHmmss}");
        Directory.CreateDirectory(artifactRoot);
        var reportPath = Path.Combine(artifactRoot, "real-user-document-crossref-report.txt");
        var documentPath = Path.Combine(
            artifactRoot,
            "VisualTeX-Real-User-Document-CrossReference.docx");
        var report = new List<string>();

        Microsoft.Office.Interop.Word.Application? sourceWord = null;
        Document? sourceDocument = null;
        Microsoft.Office.Interop.Word.Application? testWord = null;
        Document? testDocument = null;
        try
        {
            Log(report, "VisualTeX REAL WORD acceptance: the currently open user document");
            Log(report, "This test reads the source document only and performs every mutation in a full WordOpenXML clone.");

            Log(report, "[1/8] Attaching to the user's currently active Word document...");
            sourceWord = (Microsoft.Office.Interop.Word.Application)
                Marshal.GetActiveObject("Word.Application");
            sourceDocument = sourceWord.ActiveDocument;
            var sourceXml = sourceDocument.Content.WordOpenXML;
            var visualTeXXmlParts = ReadVisualTeXXmlParts(sourceDocument);
            var sourceReferences = ReadBodyReferences(sourceDocument);
            Log(report,
                $"  Source='{sourceDocument.Name}', fields={sourceDocument.Fields.Count}, "
                + $"bookmarks={sourceDocument.Bookmarks.Count}, OMaths={sourceDocument.OMaths.Count}, "
                + $"VisualTeX XML parts={visualTeXXmlParts.Count}");
            LogReferences(report, sourceReferences, "source document");
            if (sourceReferences.Any(reference => IsTinyWhite(reference)))
                Log(report, "  Source document reproduces plain native REF fields at white 1 pt.");
            else
                Log(report, "  Source document currently has normal-size REF fields; continuing with the Frame-visibility fixture.");

            Log(report, "[2/8] Opening a separate visible real Word instance and cloning all content + VisualTeX metadata...");
            testWord = new Microsoft.Office.Interop.Word.Application
            {
                Visible = true,
                DisplayAlerts = WdAlertLevel.wdAlertsNone,
            };
            testDocument = testWord.Documents.Add();
            testDocument.Content.InsertXML(sourceXml);
            foreach (var xml in visualTeXXmlParts)
            {
                object? added = null;
                try { added = testDocument.CustomXMLParts.Add(xml); }
                finally { Release(added); }
            }
            var clonePageCount = testDocument.ComputeStatistics(WdStatistic.wdStatisticPages);
            Log(report,
                $"  Clone fields={testDocument.Fields.Count}, bookmarks={testDocument.Bookmarks.Count}, "
                + $"OMaths={testDocument.OMaths.Count}, copied VisualTeX XML parts={visualTeXXmlParts.Count}, "
                + $"pages={clonePageCount}");

            var numberingType = typeof(VisualTeX.WordVsto.ThisAddIn).Assembly.GetType(
                "VisualTeX.WordVsto.WordEquationNumbering",
                throwOnError: true,
                ignoreCase: false)
                ?? throw new TypeLoadException("WordEquationNumbering was not found.");
            const BindingFlags flags = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
            var reconcile = numberingType.GetMethod("Reconcile", flags)
                ?? throw new MissingMethodException("WordEquationNumbering.Reconcile was not found.");
            var reconcileFormula = numberingType.GetMethod("ReconcileFormula", flags)
                ?? throw new MissingMethodException("WordEquationNumbering.ReconcileFormula was not found.");

            var caption = FindCaption(testDocument);
            var beforeCaption = SnapshotCaption(caption.Range);
            LogCaption(report, beforeCaption, "clone before migration");
            var beforeReferences = ReadBodyReferences(testDocument);
            LogReferences(report, beforeReferences, "clone before migration");
            if (beforeCaption.FrameCount < 1)
                throw new InvalidDataException(
                    "The clone did not preserve the source document's visible legacy SEQ Frame.");

            Log(report, "[3/8] Running the compiled production 'Update Equation Numbers' reconciliation...");
            var reconciled = (int)(reconcile.Invoke(null, new object[] { testDocument }) ?? 0);
            Log(report, $"  Production reconciliation found {reconciled} numbered formula(s).");
            if (reconciled < 1)
                throw new InvalidDataException("Production reconciliation did not find the numbered formula.");
            Release(caption.Range);
            caption = FindCaption(testDocument);
            var afterUpdateCaption = SnapshotCaption(caption.Range);
            LogCaption(report, afterUpdateCaption, "after VisualTeX update");
            AssertClippedEdgeCaption(afterUpdateCaption, "VisualTeX update");
            if (testDocument.Frames.Count != 1)
                throw new InvalidDataException(
                    $"VisualTeX update left {testDocument.Frames.Count} Frames; legacy empty Frames were not removed.");
            var afterUpdatePageCount = testDocument.ComputeStatistics(WdStatistic.wdStatisticPages);
            if (afterUpdatePageCount != clonePageCount)
                throw new InvalidDataException(
                    $"VisualTeX update changed the document from {clonePageCount} to {afterUpdatePageCount} pages.");
            var afterUpdateReferences = ReadBodyReferences(testDocument);
            LogReferences(report, afterUpdateReferences, "after VisualTeX update");
            AssertNormalReferences(afterUpdateReferences, "VisualTeX update");
            var nativeItems = testDocument.GetCrossReferenceItems(
                WdCaptionLabelID.wdCaptionEquation) as Array;
            if (nativeItems is null || nativeItems.Length < 1)
                throw new InvalidDataException(
                    "Moving the SEQ source into an off-canvas Frame removed it from Word's native cross-reference list.");
            Log(report,
                $"  PASS Word's native Cross-reference dialog inventory still contains {nativeItems.Length} Equation item(s).");

            Log(report, "[4/8] Recreating the exact old plain REF state, then running Word's native F9 update...");
            ForceOldNativeReferenceAppearance(testDocument);
            var forcedReferences = ReadBodyReferences(testDocument);
            LogReferences(report, forcedReferences, "forced old state before F9");
            if (!forcedReferences.All(IsTinyWhite))
                throw new InvalidDataException("The F9 fixture was not fully restored to white 1 pt.");
            testDocument.Fields.Update();
            var afterF9References = ReadBodyReferences(testDocument);
            LogReferences(report, afterF9References, "after native Word F9");
            AssertNormalReferences(afterF9References, "native Word F9");
            Log(report,
                "  PASS plain Word-generated 'REF ... \\h' fields recover to normal black formatting without CHARFORMAT.");

            Log(report, "[5/8] Inserting a new reference through Word's native InsertCrossReference API...");
            InsertNativeCrossReference(testDocument);
            var afterNativeInsert = ReadBodyReferences(testDocument);
            LogReferences(report, afterNativeInsert, "after native InsertCrossReference");
            AssertNormalReferences(afterNativeInsert, "native InsertCrossReference");
            var newest = afterNativeInsert.OrderBy(reference => reference.Start).Last();
            if (newest.Code.IndexOf("CHARFORMAT", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new InvalidDataException(
                    "The new fixture unexpectedly used VisualTeX's CHARFORMAT field instead of a plain Word native REF.");
            Log(report,
                $"  PASS newest plain native REF is {newest.Size:F2} pt, color={newest.Color}, hidden={newest.Hidden}.");

            Log(report, "[6/8] Running the exact production formula-format-conversion reconciliation entry...");
            InvokeFormatConversionReconciliation(
                testDocument,
                caption.FormulaId,
                reconcileFormula);
            Release(caption.Range);
            caption = FindCaption(testDocument);
            var afterConversionCaption = SnapshotCaption(caption.Range);
            LogCaption(report, afterConversionCaption, "after format conversion reconciliation");
            AssertClippedEdgeCaption(afterConversionCaption, "format conversion");
            var afterConversionReferences = ReadBodyReferences(testDocument);
            LogReferences(report, afterConversionReferences, "after format conversion reconciliation");
            AssertNormalReferences(afterConversionReferences, "format conversion");
            Log(report,
                "  PASS formula conversion keeps the SEQ source off-canvas and every native body REF readable.");

            Log(report, "[7/8] Running a second native F9 after conversion...");
            testDocument.Fields.Update();
            var afterConversionF9 = ReadBodyReferences(testDocument);
            LogReferences(report, afterConversionF9, "after conversion + native F9");
            AssertNormalReferences(afterConversionF9, "conversion + native F9");
            Log(report, "  PASS native refresh remains stable after formula-format conversion.");

            Log(report, "[8/8] Saving the inspectable real Word acceptance document...");
            testDocument.SaveAs2(documentPath, WdSaveFormat.wdFormatXMLDocument);
            Log(report, "RESULT: PASS");
            Log(report, $"  Acceptance DOCX: {documentPath}");
            Log(report, $"  Detailed report: {reportPath}");
            File.WriteAllLines(reportPath, report, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            return 0;
        }
        catch (Exception error)
        {
            var actual = error is TargetInvocationException invocation && invocation.InnerException is not null
                ? invocation.InnerException
                : error;
            Log(report, $"RESULT: FAIL - {actual}");
            File.WriteAllLines(reportPath, report, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            return 1;
        }
        finally
        {
            if (testDocument is not null)
            {
                try { testDocument.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(testDocument);
            if (testWord is not null)
            {
                try { testWord.Quit(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(testWord);
            Release(sourceDocument);
            Release(sourceWord);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static string FindWindowsRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (Directory.Exists(Path.Combine(directory.FullName, "src-windows"))
                && File.Exists(Path.Combine(directory.FullName, "package.json")))
                return directory.FullName;
            directory = directory.Parent;
        }
        var fallback = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        if (Directory.Exists(Path.Combine(fallback, "src-windows"))) return fallback;
        throw new DirectoryNotFoundException("The apps/windows project root was not found.");
    }

    private static List<string> ReadVisualTeXXmlParts(Document document)
    {
        var result = new List<string>();
        var parts = document.CustomXMLParts;
        try
        {
            for (var index = 1; index <= parts.Count; index++)
            {
                object? part = null;
                try
                {
                    part = parts[index];
                    dynamic dynamicPart = part;
                    if (string.Equals(
                            (string?)dynamicPart.NamespaceURI,
                            VisualTeXXmlNamespace,
                            StringComparison.Ordinal))
                        result.Add((string)dynamicPart.XML);
                }
                finally { Release(part); }
            }
        }
        finally { Release(parts); }
        return result;
    }

    private static CaptionInfo FindCaption(Document document)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                try
                {
                    bookmark = bookmarks[index];
                    var name = bookmark.Name ?? string.Empty;
                    if (!name.StartsWith(NativeCaptionPrefix, StringComparison.Ordinal)) continue;
                    var range = bookmark.Range;
                    return new CaptionInfo(
                        name.Substring(NativeCaptionPrefix.Length),
                        range);
                }
                finally { Release(bookmark); }
            }
        }
        finally { Release(bookmarks); }
        throw new InvalidDataException("No VisualTeX native caption bookmark was found.");
    }

    private static List<ReferenceSnapshot> ReadBodyReferences(Document document)
    {
        var result = new List<ReferenceSnapshot>();
        Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                WordRange? code = null;
                WordRange? fieldResult = null;
                Microsoft.Office.Interop.Word.Font? font = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    var codeText = NormalizeCode(code.Text);
                    if (!Regex.IsMatch(
                            codeText,
                            @"^REF\s+VTEqNum_[0-9a-f]{32}\b",
                            RegexOptions.IgnoreCase))
                        continue;
                    fieldResult = field.Result;
                    if ((bool)fieldResult.get_Information(WdInformation.wdWithInTable))
                        continue;
                    font = fieldResult.Font;
                    result.Add(new ReferenceSnapshot(
                        index,
                        codeText,
                        (fieldResult.Text ?? string.Empty).Replace("\r", "¶").Replace("\n", "¶"),
                        font.Size,
                        font.Hidden,
                        (int)font.Color,
                        font.Position,
                        fieldResult.Start));
                }
                finally
                {
                    Release(font);
                    Release(fieldResult);
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
        return result;
    }

    private static CaptionSnapshot SnapshotCaption(WordRange range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        Frames? frames = null;
        Frame? frame = null;
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        try
        {
            font = range.Font;
            frames = range.Frames;
            var frameCount = frames.Count;
            var horizontal = double.NaN;
            var vertical = double.NaN;
            var width = double.NaN;
            var height = double.NaN;
            var horizontalPage = double.NaN;
            var verticalPage = double.NaN;
            var pageWidth = double.NaN;
            var pageHeight = double.NaN;
            if (frameCount > 0)
            {
                frame = frames[1];
                horizontal = frame.HorizontalPosition;
                vertical = frame.VerticalPosition;
                width = frame.Width;
                height = frame.Height;
                horizontalPage = Convert.ToDouble(
                    range.get_Information(WdInformation.wdHorizontalPositionRelativeToPage));
                verticalPage = Convert.ToDouble(
                    range.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
                sections = range.Sections;
                section = sections[1];
                pageSetup = section.PageSetup;
                pageWidth = pageSetup.PageWidth;
                pageHeight = pageSetup.PageHeight;
            }
            return new CaptionSnapshot(
                font.Size,
                font.Hidden,
                (int)font.Color,
                font.Position,
                frameCount,
                horizontal,
                vertical,
                width,
                height,
                horizontalPage,
                verticalPage,
                pageWidth,
                pageHeight);
        }
        finally
        {
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(frame);
            Release(frames);
            Release(font);
        }
    }

    private static void ForceOldNativeReferenceAppearance(Document document)
    {
        Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                WordRange? code = null;
                WordRange? result = null;
                Microsoft.Office.Interop.Word.Font? font = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    var normalized = NormalizeCode(code.Text);
                    if (!Regex.IsMatch(
                            normalized,
                            @"^REF\s+VTEqNum_[0-9a-f]{32}\b",
                            RegexOptions.IgnoreCase))
                        continue;
                    result = field.Result;
                    if ((bool)result.get_Information(WdInformation.wdWithInTable))
                        continue;
                    var plainCode = Regex.Replace(
                        code.Text ?? string.Empty,
                        @"\\\*\s+(?:CHARFORMAT|MERGEFORMAT)\b",
                        string.Empty,
                        RegexOptions.IgnoreCase);
                    code.Text = plainCode.TrimEnd() + " ";
                    font = result.Font;
                    font.Size = 1f;
                    font.Hidden = 0;
                    font.Color = WdColor.wdColorWhite;
                    font.Position = 0;
                }
                finally
                {
                    Release(font);
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
    }

    private static void InsertNativeCrossReference(Document document)
    {
        WordRange? insertion = null;
        try
        {
            var position = Math.Max(document.Content.Start, document.Content.End - 1);
            insertion = document.Range(position, position);
            insertion.InsertAfter("\rNative Word cross-reference acceptance: ");
            insertion.SetRange(document.Content.End - 1, document.Content.End - 1);
            insertion.Font.Size = 14f;
            insertion.Font.Color = WdColor.wdColorAutomatic;
            insertion.InsertCrossReference(
                ReferenceType: WdCaptionLabelID.wdCaptionEquation,
                ReferenceKind: WdReferenceKind.wdEntireCaption,
                ReferenceItem: 1,
                InsertAsHyperlink: true,
                IncludePosition: false);
        }
        finally { Release(insertion); }
    }

    private static void InvokeFormatConversionReconciliation(
        Document document,
        string formulaId,
        MethodInfo reconcileFormula)
    {
        Bookmarks? bookmarks = null;
        Bookmark? anchorBookmark = null;
        WordRange? anchorRange = null;
        OMaths? maths = null;
        WordRange? formulaRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = OmmlBookmarkPrefix + Guid.Parse(formulaId).ToString("N");
            if (!bookmarks.Exists(name))
                throw new InvalidDataException($"OMML formula bookmark {name} was not found.");
            anchorBookmark = bookmarks[name];
            anchorRange = anchorBookmark.Range;
            var anchor = anchorRange.Start;
            maths = document.OMaths;
            var bestDistance = int.MaxValue;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? candidate = null;
                WordRange? candidateRange = null;
                try
                {
                    candidate = maths[index];
                    candidateRange = candidate.Range;
                    var containsAnchor = anchor >= candidateRange.Start
                        && anchor <= candidateRange.End;
                    var distance = containsAnchor ? 0 : candidateRange.Start - anchor;
                    if (distance < 0 || distance > 8 || distance >= bestDistance) continue;
                    Release(formulaRange);
                    formulaRange = candidateRange.Duplicate;
                    bestDistance = distance;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }
            if (formulaRange is null)
                throw new InvalidDataException(
                    "The VTOMML formula bookmark is no longer adjacent to a native OMath.");
            var metadata = new FormulaMetadata
            {
                FormulaId = formulaId,
                DisplayMode = "block",
                Numbered = true,
            };
            reconcileFormula.Invoke(
                null,
                new object[] { document, formulaRange, 24f, metadata });
        }
        finally
        {
            Release(formulaRange);
            Release(maths);
            Release(anchorRange);
            Release(anchorBookmark);
            Release(bookmarks);
        }
    }

    private static void AssertNormalReferences(
        IReadOnlyCollection<ReferenceSnapshot> references,
        string stage)
    {
        if (references.Count == 0)
            throw new InvalidDataException($"{stage}: no native body REF fields were found.");
        foreach (var reference in references)
        {
            if (reference.Size < 8f || reference.Size > 72f)
                throw new InvalidDataException(
                    $"{stage}: REF #{reference.Index} is {reference.Size:F2} pt.");
            if (reference.Hidden != 0)
                throw new InvalidDataException($"{stage}: REF #{reference.Index} is hidden.");
            if (reference.Color == (int)WdColor.wdColorWhite)
                throw new InvalidDataException($"{stage}: REF #{reference.Index} is white.");
            if (reference.Position != 0)
                throw new InvalidDataException(
                    $"{stage}: REF #{reference.Index} has position {reference.Position}.");
        }
    }

    private static void AssertClippedEdgeCaption(CaptionSnapshot caption, string stage)
    {
        if (caption.Size < 8f || caption.Size > 72f)
            throw new InvalidDataException($"{stage}: SEQ source is {caption.Size:F2} pt.");
        if (caption.Hidden != 0)
            throw new InvalidDataException($"{stage}: SEQ source is hidden.");
        if (caption.Color == (int)WdColor.wdColorWhite)
            throw new InvalidDataException($"{stage}: SEQ source is white.");
        if (caption.Position != 0)
            throw new InvalidDataException($"{stage}: SEQ source is vertically shifted.");
        if (caption.FrameCount != 1)
            throw new InvalidDataException(
                $"{stage}: expected exactly one clipping Frame, found {caption.FrameCount}.");
        if (caption.Width > 0.5 || caption.Height > 0.5)
            throw new InvalidDataException(
                $"{stage}: clipping Frame is still {caption.Width:F2} × {caption.Height:F2} pt.");
        if (caption.HorizontalPage < caption.PageWidth - 1.0
            || caption.VerticalPage < caption.PageHeight - 1.0)
            throw new InvalidDataException(
                $"{stage}: clipping Frame is not at the bottom-right page edge; "
                + $"actual=({caption.HorizontalPage:F2},{caption.VerticalPage:F2}), "
                + $"page=({caption.PageWidth:F2},{caption.PageHeight:F2}).");
    }

    private static bool IsTinyWhite(ReferenceSnapshot reference) =>
        reference.Size <= 2f && reference.Color == (int)WdColor.wdColorWhite;

    private static bool IsTinyWhite(CaptionSnapshot caption) =>
        caption.Size <= 2f && caption.Color == (int)WdColor.wdColorWhite;

    private static string NormalizeCode(string? value) =>
        Regex.Replace(value ?? string.Empty, @"\s+", " ").Trim();

    private static void LogReferences(
        ICollection<string> report,
        IReadOnlyCollection<ReferenceSnapshot> references,
        string stage)
    {
        Log(report, $"  {stage}: {references.Count} native body REF field(s)");
        foreach (var reference in references)
        {
            Log(report,
                $"    #{reference.Index} text='{reference.Text}' size={reference.Size:F2} pt "
                + $"hidden={reference.Hidden} color={reference.Color} position={reference.Position} "
                + $"code='{reference.Code}'");
        }
    }

    private static void LogCaption(
        ICollection<string> report,
        CaptionSnapshot caption,
        string stage) =>
        Log(report,
            $"  {stage}: SEQ size={caption.Size:F2} pt hidden={caption.Hidden} "
            + $"color={caption.Color} position={caption.Position} frameCount={caption.FrameCount} "
            + $"frame=({caption.Horizontal:F2},{caption.Vertical:F2}) "
            + $"size=({caption.Width:F2}×{caption.Height:F2}) "
            + $"pagePos=({caption.HorizontalPage:F2},{caption.VerticalPage:F2}) "
            + $"pageSize=({caption.PageWidth:F2}×{caption.PageHeight:F2})");

    private static void Log(ICollection<string> report, string message)
    {
        report.Add(message);
        Console.WriteLine(message);
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.FinalReleaseComObject(value); } catch { }
    }

    private sealed record CaptionInfo(string FormulaId, WordRange Range);

    private sealed record CaptionSnapshot(
        float Size,
        int Hidden,
        int Color,
        int Position,
        int FrameCount,
        double Horizontal,
        double Vertical,
        double Width,
        double Height,
        double HorizontalPage,
        double VerticalPage,
        double PageWidth,
        double PageHeight);

    private sealed record ReferenceSnapshot(
        int Index,
        string Code,
        string Text,
        float Size,
        int Hidden,
        int Color,
        int Position,
        int Start);
}
