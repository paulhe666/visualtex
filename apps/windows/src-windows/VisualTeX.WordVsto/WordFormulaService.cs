using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    private const string RangeReferencePrefix = "visualtex-word-vsto-range:";
    private const string InlineBaselineBookmarkPrefix = "VTBL_";
    private const string InlineLatexBaselineBookmarkPrefix = "VTLB_";
    // Ordinary spaces are used only while Word materializes an inline OMath and
    // are deleted immediately afterwards. Inline OLE formulas keep a real zero-
    // width non-joiner after the object so keyboard input inherits an ordinary
    // text run instead of the OLE object's negative baseline offset.
    private const string InlineMathGuard = " ";
    private const string InlineBaselineSentinel = " ";
    private const string InlineOleTypingAnchor = "\u200C";
    private const string LegacyInlineMathGuard = "\u200B";
    private const string LegacyInlineBaselineSentinel = "\u2060";
    private const string LegacyInlineNonbreakingBaselineSentinel = "\u00A0";
    private const string BulkInlineFormulaPlaceholder = "\uE000";
    private const float ParagraphBeforeOleDisplaySpaceAfterPoints = 0f;
    private static int inlineTypingCaretNormalizationActive;
    private readonly Application _application;
    private readonly Dictionary<string, string> _unownedOmmlSessionFormulaIds =
        new(StringComparer.OrdinalIgnoreCase);

    internal sealed class WordViewState
    {
        internal int SelectionStart { get; set; }
        internal int SelectionEnd { get; set; }
        internal int? VerticalPercentScrolled { get; set; }
        internal int? HorizontalPercentScrolled { get; set; }
    }

    private sealed class ResolvedLatexRedrawTarget
    {
        internal WordLatexRedrawTarget Target { get; set; } = new();
        internal PreparedWordBulkFormula Formula { get; set; } = new();
        internal Range SourceRange { get; set; } = null!;
        internal int SourceStart { get; set; }
        internal int SourceEnd { get; set; }
        internal string ExpectedSource { get; set; } = string.Empty;
        internal float? MathTypeDisplayColumnWidth { get; set; }
    }

    private sealed class LatexRedrawSourceContext
    {
        internal bool HasVisibleSurroundingText { get; set; }
        internal int FontContextRelativePosition { get; set; } = -1;
    }

    private sealed class FormulaToLatexTarget
    {
        internal FormulaMetadata Metadata { get; set; } = new();
        internal string ObjectMode { get; set; } = string.Empty;
        internal string LatexSource { get; set; } = string.Empty;
        internal int Start { get; set; }
        internal int End { get; set; }
        internal Range FormulaRange { get; set; } = null!;
        internal InlineShape? OleShape { get; set; }
        internal Bookmark? OmmlBookmark { get; set; }
        internal int? SourceInlineWordPosition { get; set; }
        internal float? SourceInlineBottomWhitespacePoints { get; set; }
    }

    private readonly struct InlineLatexBaselineProvenance
    {
        internal InlineLatexBaselineProvenance(
            int start,
            int end,
            int wordPosition,
            float? bottomWhitespacePoints)
        {
            Start = start;
            End = end;
            WordPosition = wordPosition;
            BottomWhitespacePoints = bottomWhitespacePoints;
        }

        internal int Start { get; }
        internal int End { get; }
        internal int WordPosition { get; }
        internal float? BottomWhitespacePoints { get; }
    }

    private sealed class InlineFollowingTextVisibility
    {
        internal int CharacterCount { get; set; }
        internal int Hidden { get; set; }
    }

    private readonly struct NonProseHostRange
    {
        internal NonProseHostRange(int start, int end)
        {
            Start = start;
            End = end;
        }

        internal int Start { get; }
        internal int End { get; }

        internal bool Contains(int position) => position >= Start && position < End;
    }

    internal sealed class MathTypeDisplayParagraphLayout
    {
        internal WdParagraphAlignment Alignment { get; set; }
        internal float LeftIndent { get; set; }
        internal float RightIndent { get; set; }
        internal float FirstLineIndent { get; set; }
        internal float SpaceBefore { get; set; }
        internal float SpaceAfter { get; set; }
        internal WdLineSpacing LineSpacingRule { get; set; }
        internal float LineSpacing { get; set; }
        internal WdBaselineAlignment BaseLineAlignment { get; set; }
        internal int KeepTogether { get; set; }
        internal int KeepWithNext { get; set; }
        internal int WidowControl { get; set; }
        internal int PageBreakBefore { get; set; }
        internal List<(float Position, WdTabAlignment Alignment, WdTabLeader Leader)> SpecialTabStops { get; } = new();
    }

    private sealed class MathTypeSectionColumnLayout
    {
        internal int Start { get; set; }
        internal int End { get; set; }
        internal float LeftMargin { get; set; }
        internal float PageTextWidth { get; set; }
        internal float? UniformColumnWidth { get; set; }
        internal List<float> Widths { get; } = new();
        internal List<float> SpacesAfter { get; } = new();
    }

    public WordFormulaService(Application application)
    {
        _application = application;
    }

    private static void TraceAcceptancePerformance(
        string operation,
        string stage,
        Stopwatch stopwatch,
        ref long checkpoint)
    {
        var acceptanceTrace = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
            "1",
            StringComparison.Ordinal);
        var formatTrace = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF"),
            "1",
            StringComparison.Ordinal);
        var suppressAcceptanceConsoleTrace = string.Equals(
            Environment.GetEnvironmentVariable(
                "VISUALTEX_VSTO_SUPPRESS_ACCEPTANCE_PERF"),
            "1",
            StringComparison.Ordinal);
        if (!acceptanceTrace && !formatTrace)
            return;
        var elapsed = stopwatch.ElapsedMilliseconds;
        var message =
            $"[perf] {operation}.{stage}: +{elapsed - checkpoint}ms ({elapsed}ms total)";
        if (acceptanceTrace && !suppressAcceptanceConsoleTrace)
            Console.WriteLine("    " + message);
        if (formatTrace)
            WordDoubleClickHook.TraceMessage("format-conversion-subperf " + message);
        checkpoint = elapsed;
    }

    public OfficeSelection ReadSelection() => ReadSelection(null);

    // A create command needs only the exact insertion range that exists when the
    // Ribbon callback fires. Do not run the edit-oriented formula discovery path:
    // it can enumerate every VTOMML bookmark after a local miss, adopt native
    // equations, and replace a collapsed caret with an existing equation range.
    // Besides being needlessly expensive in formula-heavy documents, doing that
    // after an asynchronous health check lets Word move the live Selection before
    // we capture it. The result is both visible latency and insertion at the wrong
    // coordinate.
    public OfficeSelection ReadCreateSelection()
    {
        Document? document = null;
        Selection? selection = null;
        Range? range = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            selection = _application.Selection;
            range = selection.Range;
            return new OfficeSelection
            {
                Host = "word",
                DocumentId = DocumentIdentity(document),
                ObjectId = RangeReference(range),
                ReadOnly = document.ReadOnly,
            };
        }
        finally
        {
            Release(range);
            Release(selection);
            Release(document);
        }
    }

    public OfficeSelection ReadSelection(Selection? providedSelection)
    {
        Document? document = null;
        Selection? selection = null;
        Range? range = null;
        var ownsSelection = providedSelection is null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            selection = providedSelection ?? _application.Selection;
            range = selection.Range;

            var visualTeX = WordFormulaHostResolver.ResolveLocal(
                document,
                range,
                WordFormulaHostKind.VisualTeX);
            if (visualTeX is not null)
            {
                var payload = WordFormulaHostSemanticReader.Read(
                    document,
                    visualTeX);
                var metadata = payload.Metadata
                    ?? throw new InvalidDataException(
                        "The selected VisualTeX OLE has no embedded formula metadata.");
                var numbering = WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    visualTeX);
                metadata.Numbered = numbering.Numbered;
                metadata.DisplayMode = visualTeX.DisplayMode;
                return new OfficeSelection
                {
                    Host = "word",
                    DocumentId = DocumentIdentity(document),
                    ObjectId = RangeReferenceFromAddress(visualTeX.Range),
                    ReadOnly = document.ReadOnly,
                    FormulaId = metadata.FormulaId,
                    Metadata = metadata,
                    ObjectMode = FormulaOleContract.NativeOleMode,
                };
            }

            var mathType = TryReadLocalMathTypeSelection(document, range);
            if (mathType is not null)
                return mathType;

            var omml = WordFormulaHostResolver.ResolveLocal(
                document,
                range,
                WordFormulaHostKind.Omml);
            if (omml is not null)
            {
                // Inline OMML cannot own an equation number. Keep the local
                // descriptor returned by ResolveLocalOmml and avoid a second
                // numbering-container COM pass in large documents.
                if (omml.Display)
                {
                    omml.Numbering =
                        WordFormulaNumberingResolver.ResolveLocal(
                            document,
                            omml);
                }
                var payload = WordFormulaHostSemanticReader.Read(
                    document,
                    omml);
                var metadata = BuildReadOnlyOmmlSessionMetadata(
                    document,
                    omml,
                    payload);
                return new OfficeSelection
                {
                    Host = "word",
                    DocumentId = DocumentIdentity(document),
                    ObjectId = RangeReferenceFromAddress(omml.Range),
                    ReadOnly = document.ReadOnly,
                    FormulaId = metadata.FormulaId,
                    Metadata = metadata,
                    ObjectMode = FormulaOleContract.WordOmmlMode,
                };
            }

            return new OfficeSelection
            {
                Host = "word",
                DocumentId = DocumentIdentity(document),
                ObjectId = RangeReference(range),
                ReadOnly = document.ReadOnly,
            };
        }
        finally
        {
            Release(range);
            if (ownsSelection) Release(selection);
            Release(document);
        }
    }

    private OfficeSelection? TryReadLocalMathTypeSelection(
        Document document,
        Range selectionRange)
    {
        Range? probe = null;
        InlineShapes? shapes = null;
        InlineShape? selected = null;
        Range? selectedRange = null;
        try
        {
            probe = selectionRange.Duplicate;
            try { probe.MoveStart(WdUnits.wdCharacter, -2); } catch { }
            try { probe.MoveEnd(WdUnits.wdCharacter, 2); } catch { }
            shapes = probe.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate))
                        continue;
                    candidateRange = candidate.Range.Duplicate;
                    var touches = selectionRange.Start == selectionRange.End
                        ? candidateRange.Start <= selectionRange.Start
                          && selectionRange.Start <= candidateRange.End
                        : selectionRange.Start < candidateRange.End
                          && selectionRange.End > candidateRange.Start;
                    if (!touches) continue;
                    if (selected is not null)
                        throw new InvalidDataException(
                            "The selected range contains more than one MathType formula.");
                    selected = candidate;
                    candidate = null;
                    selectedRange = candidateRange;
                    candidateRange = null;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }

            if (selected is null || selectedRange is null)
                return null;

            var metadata = MathTypeOleInterop.ReadMetadata(
                _application,
                selected);
            return new OfficeSelection
            {
                Host = "word",
                DocumentId = DocumentIdentity(document),
                ObjectId = RangeReference(selectedRange),
                ReadOnly = document.ReadOnly,
                FormulaId = metadata.FormulaId,
                Metadata = metadata,
                ObjectMode = FormulaOleContract.MathTypeOleMode,
            };
        }
        finally
        {
            Release(selectedRange);
            Release(selected);
            Release(shapes);
            Release(probe);
        }
    }

    private FormulaMetadata BuildReadOnlyOmmlSessionMetadata(
        Document document,
        WordFormulaHostDescriptor host,
        WordFormulaSemanticPayload payload)
    {
        Range? range = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            var fontSize = FormulaFontSize.DefaultPt;
            try
            {
                font = range.Font;
                var candidate = font.Size;
                if (candidate > 0 && candidate < 256)
                    fontSize = FormulaFontSize.Normalize(candidate);
            }
            catch { }

            var formulaId = ResolveReadOnlyOmmlSessionFormulaId(
                document,
                host);
            var now = DateTimeOffset.UtcNow.ToString("O");
            return new FormulaMetadata
            {
                FormulaId = formulaId,
                Latex = payload.Latex,
                Lines = new List<FormulaLine>
                {
                    new()
                    {
                        Id = Guid.NewGuid().ToString("D"),
                        Latex = payload.Latex,
                    },
                },
                DisplayMode = host.DisplayMode,
                Numbered = host.Numbering.Numbered,
                FontSizePt = fontSize,
                RenderFontSizePt = fontSize,
                CreatedWithVersion = "1.0.18",
                UpdatedWithVersion = "1.0.18",
                CreatedAt = now,
                UpdatedAt = now,
            };
        }
        finally
        {
            Release(font);
            Release(range);
        }
    }

    private string ResolveReadOnlyOmmlSessionFormulaId(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (Guid.TryParse(
                host.FormulaId,
                out var durable))
            return durable.ToString("D");

        var address =
            host.Range;
        var key =
            DocumentIdentity(document)
            + "|"
            + ((int)address.StoryType).ToString(
                System.Globalization.CultureInfo.InvariantCulture)
            + ":"
            + address.Start.ToString(
                System.Globalization.CultureInfo.InvariantCulture)
            + ":"
            + address.End.ToString(
                System.Globalization.CultureInfo.InvariantCulture);

        if (_unownedOmmlSessionFormulaIds.TryGetValue(
                key,
                out var existing)
            && Guid.TryParse(
                existing,
                out var parsed))
            return parsed.ToString("D");

        var created =
            Guid.NewGuid().ToString("D");
        _unownedOmmlSessionFormulaIds[key] =
            created;
        return created;
    }

    private static string RangeReferenceFromAddress(
        WordFormulaRangeAddress address) =>
        $"{RangeReferencePrefix}{address.Start}:{address.End}";


    public OfficeSelection? ReadVisualTeXOmmlAtScreenPoint(
        int screenX,
        int screenY)
    {
        Document? document = null;
        Window? window = null;
        object? pointObject = null;
        Range? pointRange = null;
        Range? resolvedRange = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            window = _application.ActiveWindow;
            pointObject = window.RangeFromPoint(screenX, screenY);
            pointRange = pointObject as Range;
            if (pointRange is null) return null;
            pointObject = null;

            var host = WordFormulaHostResolver.ResolveLocal(
                document,
                pointRange,
                WordFormulaHostKind.Omml);
            if (host is null) return null;

            resolvedRange = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            if (!ScreenPointHitsRange(
                    window,
                    resolvedRange,
                    screenX,
                    screenY))
                return null;

            host.Numbering = WordFormulaNumberingResolver.ResolveLocal(
                document,
                host);
            var payload = WordFormulaHostSemanticReader.Read(
                document,
                host);
            var metadata = BuildReadOnlyOmmlSessionMetadata(
                document,
                host,
                payload);

            return new OfficeSelection
            {
                Host = "word",
                DocumentId = DocumentIdentity(document),
                ObjectId = RangeReferenceFromAddress(host.Range),
                ReadOnly = document.ReadOnly,
                FormulaId = metadata.FormulaId,
                Metadata = metadata,
                ObjectMode = FormulaOleContract.WordOmmlMode,
            };
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(resolvedRange);
            Release(pointRange);
            Release(pointObject);
            Release(window);
            Release(document);
        }
    }

    public bool IsFormulaAtScreenPoint(
        OfficeSelection? selected,
        int screenX,
        int screenY)
    {
        if (selected?.Metadata is null
            || string.IsNullOrWhiteSpace(selected.ObjectId))
            return false;

        Document? document = null;
        Window? window = null;
        Range? captured = null;
        Range? formulaRange = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null
                || !string.Equals(
                    DocumentIdentity(document),
                    selected.DocumentId,
                    StringComparison.OrdinalIgnoreCase))
                return false;

            captured = WordFormulaOperationLocator.ResolveCapturedRange(
                document,
                selected.ObjectId);
            window = _application.ActiveWindow;

            if (string.Equals(
                    selected.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                var host = WordFormulaHostResolver.ResolveLocal(
                    document,
                    captured,
                    WordFormulaHostKind.Omml);
                if (host is null) return false;
                formulaRange = WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            }
            else if (string.Equals(
                         selected.ObjectMode,
                         FormulaOleContract.NativeOleMode,
                         StringComparison.Ordinal))
            {
                var host = WordFormulaHostResolver.ResolveLocal(
                    document,
                    captured,
                    WordFormulaHostKind.VisualTeX);
                if (host is null) return false;
                formulaRange = WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            }
            else if (string.Equals(
                         selected.ObjectMode,
                         FormulaOleContract.MathTypeOleMode,
                         StringComparison.Ordinal))
            {
                var host = WordMathTypeHostAdapter.ResolveLocal(
                    _application,
                    document,
                    captured);
                if (host is null) return false;
                formulaRange = WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            }
            else
            {
                return false;
            }

            return ScreenPointHitsRange(
                window,
                formulaRange,
                screenX,
                screenY);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(formulaRange);
            Release(captured);
            Release(window);
            Release(document);
        }
    }

    private static bool ScreenPointHitsRange(
        Window window,
        Range range,
        int screenX,
        int screenY)
    {
        try
        {
            window.GetPoint(
                out var left,
                out var top,
                out var width,
                out var height,
                range);
            return WordDoubleClickRouting.ScreenPointHitsFormulaRectangle(
                screenX,
                screenY,
                left,
                top,
                width,
                height);
        }
        catch
        {
            return false;
        }
    }

    private static Range? TryResolveNativeOmmlAtRange(
        Document document,
        Range selectionRange)
    {
        Range? probe = null;
        OMaths? maths = null;
        Range? best = null;
        var candidates = new List<(int Start, int End)>();
        try
        {
            probe = selectionRange.Duplicate;
            maths = probe.OMaths;
            if (maths.Count == 0)
            {
                Release(maths);
                maths = null;
                Release(probe);
                probe = null;
                Range? content = null;
                try
                {
                    content = document.Content;
                    var start = Math.Max(content.Start, selectionRange.Start - 1);
                    var end = Math.Min(content.End, selectionRange.End + 1);
                    probe = document.Range(start, end);
                }
                finally { Release(content); }
                maths = probe.OMaths;
            }

            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    var containsCaret = selectionRange.Start == selectionRange.End
                        && selectionRange.Start >= range.Start
                        && selectionRange.Start <= range.End;
                    var overlapsSelection = selectionRange.Start < selectionRange.End
                        && range.Start < selectionRange.End
                        && range.End > selectionRange.Start;
                    if (!containsCaret && !overlapsSelection) continue;
                    candidates.Add((range.Start, range.End));
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }

            if (candidates.Count == 0) return null;
            var roots = candidates
                .Where(candidate => !candidates.Any(other =>
                    other != candidate
                    && other.Start <= candidate.Start
                    && other.End >= candidate.End))
                .Distinct()
                .ToArray();
            if (roots.Length != 1) return null;
            best = document.Range(roots[0].Start, roots[0].End);
            var result = best;
            best = null;
            return result;
        }
        catch { return null; }
        finally
        {
            Release(best);
            Release(maths);
            Release(probe);
        }
    }

    public string ReadActiveDocumentId()
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            return DocumentIdentity(document);
        }
        finally { Release(document); }
    }

    private enum SelectedOleKind
    {
        None,
        VisualTeX,
        MathType,
    }

    public bool IsSelectedNativeOle() =>
        ReadSelectedOleKind() == SelectedOleKind.VisualTeX;

    public bool IsSelectedMathTypeOle() =>
        ReadSelectedOleKind() == SelectedOleKind.MathType;

    public bool OpenMathTypeNativeEditorAtRange(
        int targetStart,
        int targetEnd,
        int screenX,
        int screenY)
    {
        Document? document = null;
        Window? window = null;
        Range? content = null;
        Range? targetRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        OLEFormat? format = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null) return false;
            content = document.Content;
            var start = Math.Max(content.Start, Math.Min(targetStart, content.End));
            var end = Math.Max(start, Math.Min(targetEnd, content.End));
            if (end <= start) return false;

            targetRange = document.Range(start, end);
            shapes = targetRange.InlineShapes;
            if (shapes.Count != 1) return false;
            shape = shapes[1];
            if (!MathTypeOleInterop.IsMathTypeOle(shape)) return false;
            shapeRange = shape.Range;
            if (shapeRange.End <= start || shapeRange.Start >= end) return false;

            window = _application.ActiveWindow;
            if (!ScreenPointHitsRange(window, shapeRange, screenX, screenY))
                return false;

            format = shape.OLEFormat;
            object openVerb = (int)WdOLEVerb.wdOLEVerbOpen;
            format.DoVerb(ref openVerb);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(format);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(targetRange);
            Release(content);
            Release(window);
            Release(document);
        }
    }

    public bool IsSelectedInterceptableOle() =>
        ReadSelectedOleKind() is SelectedOleKind.VisualTeX or SelectedOleKind.MathType;

    private SelectedOleKind ReadSelectedOleKind()
    {
        Selection? selection = null;
        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        OLEFormat? format = null;
        try
        {
            selection = _application.Selection;
            range = selection.Range;
            shapes = range.InlineShapes;
            if (shapes.Count != 1) return SelectedOleKind.None;
            shape = shapes[1];
            if (shape.Type is not WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                and not WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                return SelectedOleKind.None;
            format = shape.OLEFormat;
            if (string.Equals(
                    format.ProgID,
                    FormulaOleContract.ProgId,
                    StringComparison.OrdinalIgnoreCase))
                return SelectedOleKind.VisualTeX;
            return MathTypeOleInterop.TryResolveCapabilities(format.ProgID, out _)
                ? SelectedOleKind.MathType
                : SelectedOleKind.None;
        }
        catch
        {
            return SelectedOleKind.None;
        }
        finally
        {
            Release(format);
            Release(shape);
            Release(shapes);
            Release(range);
            Release(selection);
        }
    }

    public void NormalizeTypingCaretAfterInlineFormula(Selection selection)
    {
        if (selection is null) return;
        // Selection.SetRange can synchronously re-enter the VSTO
        // WindowSelectionChange handler. Tests also exercise the public service
        // from a second instance while the add-in's deferred pass is pending.
        // Serialize the tiny boundary repair across service instances so two
        // passes cannot both observe a missing VTBL owner and insert duplicate
        // U+200C typing anchors.
        if (Interlocked.CompareExchange(
                ref inlineTypingCaretNormalizationActive,
                1,
                0) != 0)
            return;
        try
        {
            NormalizeTypingCaretAfterInlineFormulaCore(selection);
        }
        finally
        {
            Interlocked.Exchange(ref inlineTypingCaretNormalizationActive, 0);
        }
    }

    private void NormalizeTypingCaretAfterInlineFormulaCore(
        Selection selection)
    {
        Range? caret = null;
        Document? document = null;
        Range? hostRange = null;
        try
        {
            caret = selection.Range;
            if (caret.Start != caret.End)
                return;
            document = caret.Document;

            WordFormulaHostDescriptor? host = null;
            try
            {
                host =
                    WordFormulaHostResolver.ResolveLocal(
                        document,
                        caret,
                        WordFormulaHostKind.VisualTeX);
            }
            catch (InvalidDataException) { }

            if (host is null)
            {
                try
                {
                    host =
                        WordFormulaHostResolver.ResolveLocal(
                            document,
                            caret,
                            WordFormulaHostKind.Omml);
                }
                catch (InvalidDataException) { }
            }

            if (host is null)
            {
                try
                {
                    host =
                        WordMathTypeHostAdapter.ResolveLocal(
                            _application,
                            document,
                            caret);
                }
                catch (InvalidDataException) { }
            }

            if (host is null
                || !string.Equals(
                    host.DisplayMode,
                    "inline",
                    StringComparison.OrdinalIgnoreCase)
                || caret.Start != host.Range.End)
                return;

            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            WordFormulaHostLayout.RestoreInlineCaret(
                _application,
                hostRange);
        }
        catch
        {
            // Selection-change normalization is best-effort and must never
            // interrupt ordinary Word input.
        }
        finally
        {
            Release(hostRange);
            Release(document);
            Release(caret);
        }
    }

    private static void ApplyInlineTypingFormattingToSelection(
        Selection selection,
        Range formulaRange)
    {
        Range? source = null;
        Range? caret = null;
        Microsoft.Office.Interop.Word.Font? sourceFont = null;
        Microsoft.Office.Interop.Word.Font? selectionFont = null;
        try
        {
            caret = selection.Range;
            source = FindInlineTypingFormatSource(formulaRange, caret);
            if (source is not null)
            {
                sourceFont = source.Font;
                selectionFont = selection.Font;
                selectionFont.Name = sourceFont.Name;
                try { selectionFont.NameAscii = sourceFont.NameAscii; } catch { }
                try { selectionFont.NameFarEast = sourceFont.NameFarEast; } catch { }
                try { selectionFont.NameOther = sourceFont.NameOther; } catch { }
                selectionFont.Size = sourceFont.Size;
                selectionFont.Bold = sourceFont.Bold;
                selectionFont.Italic = sourceFont.Italic;
                try { selectionFont.Underline = sourceFont.Underline; } catch { }
                try { selectionFont.Color = sourceFont.Color; } catch { }
                var sourcePosition = sourceFont.Position;
                selectionFont.Position = sourcePosition == (int)WdConstants.wdUndefined
                    ? 0
                    : sourcePosition;
            }
            else
            {
                selectionFont = selection.Font;
                selectionFont.Position = 0;
            }
            selectionFont.Hidden = 0;
            selectionFont.Subscript = 0;
            selectionFont.Superscript = 0;
            try { selectionFont.Spacing = 0; } catch { }
            try { selectionFont.Scaling = 100; } catch { }
        }
        finally
        {
            Release(selectionFont);
            Release(sourceFont);
            Release(source);
            Release(caret);
        }
    }

    private static bool TryResolveWordFontSize(float value, out float fontSizePt)
    {
        fontSizePt = FormulaFontSize.DefaultPt;
        if (float.IsNaN(value)
            || float.IsInfinity(value)
            || value < FormulaFontSize.MinimumPt
            || value > FormulaFontSize.MaximumPt)
            return false;
        fontSizePt = FormulaFontSize.Normalize(value);
        return true;
    }

    public float ReadCurrentTypingFontSize()
    {
        Selection? selection = null;
        Range? selectionRange = null;
        Range? probeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            selection = _application.Selection;
            selectionRange = selection.Range;
            font = selection.Font;
            if (TryResolveWordFontSize(font.Size, out var selectedSize))
                return selectedSize;
            Release(font);
            font = null;

            paragraphs = selectionRange.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
            }

            // A collapsed insertion point can report Word's mixed-size sentinel.
            // Prefer the character immediately before the caret in the current
            // paragraph, then the next character, before falling back to the
            // paragraph run as a whole. This makes a new formula inherit the
            // actual surrounding body text instead of an arbitrary global size.
            if (selectionRange.Start > (paragraphRange?.Start ?? 0))
            {
                probeRange = selectionRange.Duplicate;
                probeRange.SetRange(selectionRange.Start - 1, selectionRange.Start);
                font = probeRange.Font;
                if (TryResolveWordFontSize(font.Size, out var previousSize))
                    return previousSize;
                Release(font);
                font = null;
                Release(probeRange);
                probeRange = null;
            }

            var paragraphEnd = Math.Max(
                paragraphRange?.Start ?? selectionRange.Start,
                (paragraphRange?.End ?? selectionRange.End) - 1);
            if (selectionRange.Start < paragraphEnd)
            {
                probeRange = selectionRange.Duplicate;
                probeRange.SetRange(selectionRange.Start, selectionRange.Start + 1);
                font = probeRange.Font;
                if (TryResolveWordFontSize(font.Size, out var nextSize))
                    return nextSize;
                Release(font);
                font = null;
                Release(probeRange);
                probeRange = null;
            }

            if (paragraphRange is not null)
            {
                if (paragraphRange.End > paragraphRange.Start)
                    paragraphRange.End -= 1;
                font = paragraphRange.Font;
                if (TryResolveWordFontSize(font.Size, out var paragraphSize))
                    return paragraphSize;
            }
            return FormulaFontSize.DefaultPt;
        }
        catch
        {
            return FormulaFontSize.DefaultPt;
        }
        finally
        {
            Release(font);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probeRange);
            Release(selectionRange);
            Release(selection);
        }
    }

    public float? GetSelectedFormulaFontSize()
    {
        // Ribbon getText/getEnabled callbacks run immediately after a Word
        // SelectionChange. They must be strictly read-only: calling ReadSelection()
        // here can adopt an unowned native OMath by adding a bookmark/CustomXML
        // part, which mutates the document while Word is entering its native
        // equation editor and can suppress the normal equation editing frame.
        Document? document = null;
        Selection? selection = null;
        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? equationRange = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null) return null;
            selection = _application.Selection;
            range = selection.Range;

            shapes = range.InlineShapes;
            if (shapes.Count == 1)
            {
                shape = shapes[1];
                var metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata is not null)
                {
                    var stableSize = WordFormulaMetadataReader.IsNativeOle(shape)
                        ? WordInlineObjectGeometry.ReadStableVisualTeXSize(shape, metadata)
                        : (Width: shape.Width, Height: shape.Height);
                    return FormulaFontSize.InferOleFontSize(stableSize.Width, stableSize.Height, metadata);
                }
                if (MathTypeOleInterop.IsMathTypeOle(shape)
                    && MathTypeOleStorage.TryCaptureCompoundFileFromWordOpenXml(
                        shape,
                        out var mathTypeCompound))
                {
                    var equationNative = MathTypeOleStorage.ReadEquationNative(mathTypeCompound);
                    return FormulaFontSize.Normalize(
                        MathTypeMtefCodec.ReadEquationNativeFullFontSize(equationNative));
                }
            }

            var omml =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    range,
                    WordFormulaHostKind.Omml);
            if (omml is null)
                return null;
            equationRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    omml.Range);

            // Ribbon reads the live Word formatting of the physical OMath. Durable
            // identity/metadata is not semantic font-size evidence.
            return ReadFormulaFontSizeWithoutMutation(
                equationRange,
                metadata: null);
        }
        catch { return null; }
        finally
        {
            Release(equationRange);
            Release(shape);
            Release(shapes);
            Release(range);
            Release(selection);
            Release(document);
        }
    }

    private static float ReadFormulaFontSizeWithoutMutation(
        Range equationRange,
        FormulaMetadata? metadata)
    {
        // Numbered display OMML intentionally contains mixed native run sizes to
        // emulate Word display-style fraction arguments inside the OLE-compatible
        // tab host. The persisted VisualTeX metadata is therefore the semantic
        // size shown in the ribbon; Word may report its mixed-size sentinel here.
        if (metadata is not null && IsNumberedBlockOmml(metadata))
            return FormulaFontSize.ResolveSemanticFontSize(metadata);

        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = equationRange.Font;
            if (TryNormalizeDefinedWordFontSize(font.Size, out var size))
                return size;
        }
        catch { }
        finally { Release(font); }

        return metadata is null
            ? FormulaFontSize.DefaultPt
            : FormulaFontSize.ResolveSemanticFontSize(metadata);
    }

    public float SetSelectedFormulaFontSize(double requestedFontSizePt)
    {
        return SetSelectedFormulaFontSizeCore(
            requestedFontSizePt);
    }

    public string DeleteSelectedFormula()
    {
        return DeleteSelectedFormulaCore();
    }

    public int UpdateEquationNumbers()
    {
        return UpdateEquationNumbersCore();
    }

    public string GetEquationNumberFormatId()
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            return WordEquationNumbering.GetEquationNumberFormatId(document);
        }
        finally { Release(document); }
    }

    public string GetEquationNumberFormatDisplayName()
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            return WordEquationNumbering.GetEquationNumberFormatDisplayName(document);
        }
        finally { Release(document); }
    }

    internal void InsertEquationReference(Document document, Selection selection,
        EquationReferenceTarget target, EquationReferenceStyle style, WdColor color)
    {
        InsertEquationReferenceCore(
            document,
            selection,
            target,
            style,
            color);
    }

    private T ExecuteDocumentEdit<T>(Document document, string name, Func<T> edit,
        Action<T>? validateCompletedEdit = null)
    {
        var snapshot = new WordDocumentEditSnapshot(document);
        var undo = BeginUndoRecord(name)
            ?? throw new InvalidOperationException($"Word could not establish an independent transaction for '{name}'.");
        var ended = false;
        try
        {
            var result = edit();
            EndUndoRecord(undo); ended = true;
            // WordOpenXML can close Word's custom undo record as a side effect.
            // Perform read-only materialization verification after closing the
            // complete edit, but inside this try so failed validation still rolls
            // back the same whole transaction through the original snapshot.
            validateCompletedEdit?.Invoke(result);
            return result;
        }
        catch (Exception insertionError)
        {
            try
            {
                if (!ended) { EndUndoRecord(undo); ended = true; }
                snapshot.Restore(document);
            }
            catch (Exception recoveryError)
            {
                throw new AggregateException($"'{name}' failed and recovery could not be verified.", insertionError, recoveryError);
            }
            throw;
        }
        finally { if (!ended) EndUndoRecord(undo); Release(undo); }
    }

    // Canonical equation-field operations use the same transaction rule as all
    // rebuilt OMML/VisualTeX mutations: exactly one Word Custom UndoRecord, and
    // exactly one Undo of that record on failure. Non-undo state (for example the
    // persisted number-format preference) is restored explicitly after rollback.
    private T ExecuteFieldRefreshEdit<T>(
        Document document,
        string name,
        Func<T> edit,
        Action? restoreNonUndoState = null)
    {
        try
        {
            return WordFormulaMutationTransaction.Execute(
                _application,
                document,
                name,
                edit);
        }
        catch
        {
            restoreNonUndoState?.Invoke();
            throw;
        }
    }

    public int SetEquationNumberFormat(string formatId)
    {
        return SetEquationNumberFormatCore(
            formatId);
    }
    public string ExportSelectedOleAsPicture()
    {
        var selected = ReadSelection();
        if (!string.Equals(
                selected.ObjectMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal)
            || string.IsNullOrWhiteSpace(
                selected.FormulaId)
            || string.IsNullOrWhiteSpace(
                selected.ObjectId))
            throw new InvalidOperationException(
                "Please select one VisualTeX formula first.");

        var requiredFormulaId =
            selected.FormulaId!;
        Document? document = null;
        InlineShape? sourceShape = null;
        OLEFormat? format = null;
        object? oleObject = null;
        string? pngPath = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);

            var source =
                WordFormulaOperationLocator.ResolveCapturedHost(
                    _application,
                    document,
                    selected.ObjectId,
                    WordFormulaHostKind.VisualTeX,
                    requiredFormulaId);
            source.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    source);

            sourceShape =
                GetExactVisualTeXShape(
                    document,
                    source);
            var metadata =
                WordFormulaMetadataReader
                    .TryReadEmbeddedNativeOle(
                        sourceShape)
                ?? throw new InvalidDataException(
                    "The selected VisualTeX OLE has no authoritative embedded metadata.");
            metadata.DisplayMode =
                source.DisplayMode;
            metadata.Numbered =
                source.Numbering.Numbered;
            metadata.Validate();

            format = sourceShape.OLEFormat;
            if (!string.Equals(
                    format.ProgID,
                    FormulaOleContract.ProgId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "The captured VisualTeX host is no longer the expected OLE object.");
            oleObject =
                WordOleObjectAccessor.GetRunningObject(
                    format);
            pngPath =
                OlePngPreviewExtractor.MaterializePng(
                    oleObject,
                    requiredFormulaId);

            var sourceWidth =
                sourceShape.Width;
            var sourceHeight =
                sourceShape.Height;

            // Release the preflight shape before the transaction and re-resolve
            // the exact captured host once mutation begins.
            Release(sourceShape);
            sourceShape = null;

            return WordFormulaMutationTransaction.Execute(
                _application,
                document,
                "VisualTeX Export Formula As Picture",
                () =>
                {
                    var current =
                        WordFormulaOperationLocator.ResolveCapturedHost(
                            _application,
                            document,
                            selected.ObjectId,
                            WordFormulaHostKind.VisualTeX,
                            requiredFormulaId);
                    current.Numbering =
                        WordFormulaNumberingResolver.ResolveLocal(
                            document,
                            current);

                    InlineShape? liveSource = null;
                    Range? liveRange = null;
                    Range? insertion = null;
                    InlineShape? picture = null;
                    try
                    {
                        liveSource =
                            GetExactVisualTeXShape(
                                document,
                                current);
                        liveRange =
                            liveSource.Range.Duplicate;
                        insertion =
                            liveRange.Duplicate;
                        insertion.Collapse(
                            WdCollapseDirection.wdCollapseStart);

                        object link = false;
                        object save = true;
                        object rangeObject = insertion;
                        picture =
                            document.InlineShapes.AddPicture(
                                pngPath!,
                                ref link,
                                ref save,
                                ref rangeObject);
                        Configure(
                            picture,
                            metadata,
                            sourceWidth,
                            sourceHeight,
                            pngPath!,
                            (float)(
                                metadata.RenderHeightPx
                                ?? 0),
                            metadata.Baseline.HasValue
                                ? (float?)metadata.Baseline.Value
                                : null,
                            string.Equals(
                                current.DisplayMode,
                                "inline",
                                StringComparison.OrdinalIgnoreCase));

                        // Clean only historical anchors left by pre-core builds.
                        // No new VTBL/U+200C marker is created.
                        RemoveInlineBaselineSentinel(
                            document,
                            requiredFormulaId);
                        RemoveInlineOleTypingAnchorAfter(
                            liveSource);
                        liveSource.Delete();

                        if (string.Equals(
                                current.DisplayMode,
                                "inline",
                                StringComparison.OrdinalIgnoreCase))
                        {
                            Range? pictureRange = null;
                            try
                            {
                                pictureRange =
                                    picture.Range.Duplicate;
                                WordFormulaHostLayout.RestoreInlineCaret(
                                    _application,
                                    pictureRange);
                            }
                            finally
                            {
                                Release(pictureRange);
                            }
                        }

                        // Numbering is outside the host. Replacing the center
                        // InlineShape in-place leaves a canonical SEQ/REF container
                        // untouched; no legacy numbering rebuild is permitted.
                        return requiredFormulaId;
                    }
                    finally
                    {
                        Release(picture);
                        Release(insertion);
                        Release(liveRange);
                        Release(liveSource);
                    }
                });
        }
        finally
        {
            if (pngPath is not null)
            {
                try { File.Delete(pngPath); }
                catch { }
            }
            Release(oleObject);
            Release(format);
            Release(sourceShape);
            Release(document);
        }
    }

    public OfficeObjectResult Insert(OfficeSessionDocument session, string imagePath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShape? shape = null;
        UndoRecord? undoRecord = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            selection = _application.Selection;
            insertion = ResolveSessionInsertionRange(document, session, selection);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            undoRecord = BeginUndoRecord(
                session.DisplayMode == "inline"
                    ? "VisualTeX Insert Inline Formula"
                    : "VisualTeX Insert Display Formula");
            object link = false;
            object save = true;
            object rangeObject;
            if (session.DisplayMode == "inline")
            {
                rangeObject = insertion;
                shape = document.InlineShapes.AddPicture(
                    imagePath,
                    ref link,
                    ref save,
                    ref rangeObject);
            }
            else
            {
                CompactParagraphBeforeOleDisplayFormula(document, insertion);
                var displayInsertion = ResolveDisplayInsertionRange(document, insertion);
                Release(insertion);
                insertion = displayInsertion;
                rangeObject = insertion;
                shape = document.InlineShapes.AddPicture(
                    imagePath,
                    ref link,
                    ref save,
                    ref rangeObject);
            }
            Configure(
                shape,
                metadata,
                (session.ExportResult?.Width ?? 200) * 0.75f,
                (session.ExportResult?.Height ?? 60) * 0.75f,
                imagePath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            if (session.DisplayMode == "inline")
            {
                RestoreTypingBaselineAfter(shape);
            }
            else
            {
                TryReconcileShape(document, shape, metadata);
                Range? shapeRange = null;
                try
                {
                    shapeRange = shape.Range;
                    if (session.Numbered)
                    {
                        WordEquationNumbering.CleanupNumberedDisplayInsertionSpacing(
                            document,
                            metadata.FormulaId);
                        MoveSelectionAfterNumberedDisplayFormula(
                            document,
                            selection,
                            shapeRange,
                            metadata.FormulaId);
                    }
                    else
                    {
                        MoveSelectionAfterDisplayFormula(selection, shapeRange);
                    }
                }
                finally { Release(shapeRange); }
            }
            return Result(session, document);
        }
        catch
        {
            TryDelete(shape);
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(shape);
            Release(paragraphRange);
            Release(paragraph);
            Release(insertion);
            Release(selection);
            Release(document);
        }
    }

    private void CaptureCurrentSelectionForCreateSession(
        OfficeSessionDocument session)
    {
        if (!string.IsNullOrWhiteSpace(session.SourceObjectId)
            && !string.IsNullOrWhiteSpace(session.SourceDocumentId))
            return;

        Document? document = null;
        Selection? selection = null;
        Range? range = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            selection = _application.Selection;
            range = selection.Range.Duplicate;
            if (string.IsNullOrWhiteSpace(session.SourceDocumentId))
                session.SourceDocumentId = DocumentIdentity(document);
            if (string.IsNullOrWhiteSpace(session.SourceObjectId))
                session.SourceObjectId = RangeReference(range);
        }
        finally
        {
            Release(range);
            Release(selection);
            Release(document);
        }
    }

    public OfficeObjectResult InsertOle(
        OfficeSessionDocument session,
        string pngPath,
        string emfPath,
        bool deferNumberingLayout = false,
        bool numberingScaffoldOnly = false,
        bool preserveExistingDisplayParagraphBoundary = false,
        bool preserveCapturedInsertion = false,
        float? presentationScaleX = null,
        float? presentationScaleY = null)
    {
        if (deferNumberingLayout
            || numberingScaffoldOnly
            || preserveExistingDisplayParagraphBoundary
            || preserveCapturedInsertion
            || presentationScaleX.HasValue
            || presentationScaleY.HasValue)
            throw new NotSupportedException(
                "Retired VisualTeX OLE insertion compatibility flags are not supported by the rebuilt Word host core.");

        CaptureCurrentSelectionForCreateSession(session);
        session.Mode = "create";
        session.ObjectMode =
            FormulaOleContract.NativeOleMode;
        return ApplyOmmlVisualTeXHostSession(
            session,
            sourceObjectMode: null,
            mathMl: null,
            pngPath,
            emfPath);
    }

    internal string GetMathTypeNumberPositionPreference()
    {
        var saved = WordEquationNumbering.TryGetDefaultMathTypeNumberPosition();
        if (saved is not null) return saved;

        Document? document = null;
        try
        {
            // MathType stores its own document-wide default in
            // MTEqnNumsOnRight. VisualTeX may read that value once when no
            // VisualTeX preference exists, but must never write it while a Word
            // formula transaction is active: changing MathType's global document
            // state can make the MathType add-in re-enter Word layout/field code
            // while VisualTeX is still materializing an OLE object.
            document = _application.ActiveDocument;
            return document is null
                ? "right"
                : ReadMathTypeNumberPositionPreference(document);
        }
        catch { return "right"; }
        finally { Release(document); }
    }

    internal string GetMathTypeNumberPositionForRange(string? sourceObjectId)
    {
        Document? document = null;
        InlineShape? shape = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null) return "right";
            shape = FindMathTypeOleByRange(document, sourceObjectId);
            if (shape is not null
                && MathTypeOleInterop.TryReadDisplayNumberPosition(shape, out var position))
                return position;
            return GetMathTypeNumberPositionPreference();
        }
        catch { return "right"; }
        finally
        {
            Release(shape);
            Release(document);
        }
    }

    public OfficeObjectResult InsertMathTypeOle(
        OfficeSessionDocument session,
        string mathMl,
        string? emfPath,
        string? createdObjectBookmarkName = null,
        ResolvedEquationHeadingScope? preResolvedHeadingScope = null,
        ISet<int>? preparedHeadingScopeStarts = null,
        string? isolatedNativePreviewWmfPath = null,
        float isolatedNativePreviewWidthPt = 0,
        float isolatedNativePreviewHeightPt = 0,
        int isolatedNativePreviewWordPosition = 0,
        bool isolatedNativePreviewAttempted = false,
        bool reuseExistingInlineTypingBoundary = false,
        bool updateCreatedMathTypeNumberFields = false,
        bool preserveExistingDisplayParagraphBoundary = false,
        MathTypeDisplayParagraphLayout? preservedDisplayParagraphLayout = null,
        float? knownDisplayColumnWidth = null,
        Action<Range, string>? retainMathTypeForCompletedValidation = null,
        bool preserveCapturedInsertion = false)
    {
        if (string.IsNullOrWhiteSpace(mathMl)
            || !mathMl.TrimStart().StartsWith("<math", StringComparison.Ordinal))
            throw new InvalidDataException(
                "VisualTeX did not provide valid MathML for MathType OLE insertion.");
        var hasIsolatedNativePreview =
            !string.IsNullOrWhiteSpace(isolatedNativePreviewWmfPath)
            && File.Exists(isolatedNativePreviewWmfPath)
            && isolatedNativePreviewWidthPt > 0
            && isolatedNativePreviewHeightPt > 0;
        if (!hasIsolatedNativePreview
            && (string.IsNullOrWhiteSpace(emfPath) || !File.Exists(emfPath)))
            throw new FileNotFoundException(
                "VisualTeX did not provide a valid MathType native or vector preview for OLE insertion.",
                emfPath);

        var metadata = session.ToMetadata();
        metadata.Validate();
        if (string.IsNullOrWhiteSpace(metadata.Latex))
            throw new InvalidDataException(
                "VisualTeX did not provide LaTeX source for MathType OLE insertion.");

        var inline = string.Equals(
            session.DisplayMode,
            "inline",
            StringComparison.OrdinalIgnoreCase);
        var standalone = PrepareStandaloneMathTypeOleData(
            mathMl, inline, session.FontSizePt, metadata.Latex);
        var generated = standalone.Generated;
        var compoundFile = standalone.CompoundFile;
        var expectedSignature = standalone.SemanticSignature;

        // The visible Word presentation must use MathType's own MTEF geometry.
        // Using the frontend/MathJax EMF here made otherwise-valid Equation Native
        // objects look subtly different from equations inserted by MathType itself
        // (notably relation/operator spacing) and could size the OLE host a few
        // pixels too narrowly, clipping the right-most italic glyph. Prefer the
        // MathPage native renderer whenever it is installed; keep the frontend EMF
        // only as a compatibility fallback for machines without MathPage.
        MathTypeNativePreviewRenderer.Result? nativePreview = null;
        byte[] previewWmf;
        float widthPt;
        float heightPt;
        int wordPosition;
        var renderRoot = !string.IsNullOrWhiteSpace(emfPath)
            ? Path.GetDirectoryName(emfPath) ?? Path.GetTempPath()
            : Path.GetTempPath();
        if (hasIsolatedNativePreview)
        {
            previewWmf = File.ReadAllBytes(isolatedNativePreviewWmfPath);
            widthPt = isolatedNativePreviewWidthPt;
            heightPt = isolatedNativePreviewHeightPt;
            wordPosition = isolatedNativePreviewWordPosition;
        }
        else if (!isolatedNativePreviewAttempted
            && MathTypeNativePreviewRenderer.TryRender(
                generated.Mtef,
                renderRoot,
                out var renderedNativePreview))
        {
            nativePreview = renderedNativePreview;
            previewWmf = File.ReadAllBytes(nativePreview.WmfPath);
            widthPt = nativePreview.WidthPt;
            heightPt = nativePreview.HeightPt;
            wordPosition = nativePreview.WordPosition;
        }
        else
        {
            widthPt = (float)Math.Max(1d, (session.ExportResult?.Width ?? 200d) * 0.75d);
            heightPt = (float)Math.Max(1d, (session.ExportResult?.Height ?? 60d) * 0.75d);
            var alignToWordTextBaseline = inline || session.Numbered;
            wordPosition = alignToWordTextBaseline
                ? CalculateMathTypeOleWordPosition(
                    heightPt,
                    session.ExportResult?.Height ?? 0f,
                    session.ExportResult?.Baseline)
                : 0;
            if (string.IsNullOrWhiteSpace(emfPath) || !File.Exists(emfPath))
                throw new FileNotFoundException(
                    "MathType native preview was unavailable and no fallback EMF exists.",
                    emfPath);
            previewWmf = MathTypeWordOpenXml.ConvertEnhancedMetafileToPlaceableWmf(
                emfPath!,
                widthPt,
                heightPt);
        }

        // Keep genuine MathType storage free of OlePres. Word owns the external
        // WMF presentation in the DOCX package, avoiding the blank-object
        // regression that prompted the visibility change in 6a43aec.

        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Range? displaySpacingAnchor = null;
        InlineShape? shape = null;
        Field? sourceNumberTemplateField = null;
        Bookmark? createdObjectBookmark = null;
        Range? createdObjectBookmarkRange = null;
        UndoRecord? undoRecord = null;
        var sourceParagraphCount = -1;
        var paragraphCountBeforeDisplayPreparation = -1;
        var insertionStart = -1;
        var createdSectionBreakCodeStart = -1;
        // Batch format conversion supplies a temporary object-identity bookmark.
        // Its caller performs one document-wide numbering reconciliation after all
        // replacements, so the per-item path can avoid repeated global scans.
        var useLocalConversionLookup =
            !string.IsNullOrWhiteSpace(createdObjectBookmarkName);
        var traceInsertPerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF"),
            "1",
            StringComparison.Ordinal);
        var insertPerfWatch = traceInsertPerformance ? Stopwatch.StartNew() : null;
        long insertPerfLastMs = 0;
        void TraceInsertPerf(string perfStage)
        {
            if (insertPerfWatch is null) return;
            var totalMs = insertPerfWatch.ElapsedMilliseconds;
            WordDoubleClickHook.TraceMessage(
                $"mathtype-insert-perf stage={perfStage} numbered={session.Numbered} inline={inline} deltaMs={totalMs - insertPerfLastMs} totalMs={totalMs}");
            insertPerfLastMs = totalMs;
        }
        var stage = "initialize";
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            selection = _application.Selection;

            stage = "resolve-captured-insertion";
            insertion = ResolveSessionInsertionRange(document, session, selection, preserveCapturedInsertion);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            var insertionPosition = insertion.Start;
            // Capture before Flat OPC/style materialization: neither an OLE run
            // nor MTDisplayEquation's defaults are the surrounding body font.
            var bodyFormatting = inline ? null : WordCharacterFormatting.CaptureParagraphMarkAtPosition(
                document, insertionPosition);
            TraceInsertPerf("resolve-insertion");

            undoRecord = BeginUndoRecord(
                inline
                    ? "VisualTeX Insert MathType OLE Inline Formula"
                    : session.Numbered
                        ? "VisualTeX Insert MathType OLE Numbered Display Formula"
                        : "VisualTeX Insert MathType OLE Display Formula");
            if (!inline && session.Numbered)
                EnsureEquationFieldResultsVisible(document);

            if (!inline && session.Numbered)
            {
                stage = "repair-incomplete-number-row";
                ClearIncompleteMathTypeNumberRowAtInsertion(document, insertion);
                TraceInsertPerf("clear-incomplete-number-row");
            }

            MathTypeWordOpenXml.NumberTemplate? numberTemplate = null;
            if (!inline && session.Numbered)
            {
                stage = "resolve-number-template";
                var documentNumberFormat = EquationNumberFormat.Resolve(
                    WordEquationNumbering.GetEquationNumberFormatId(document));
                if (useLocalConversionLookup)
                {
                    // For VisualTeX/OMML -> MathType batch conversion, every
                    // numbered target follows the document-selected number format.
                    // Avoid scanning document.Fields for the nearest MTPlaceRef on
                    // every formula; the batch performs one final numbering update.
                    numberTemplate = MathTypeWordOpenXml.CreateVisualTeXNumberTemplate(
                        documentNumberFormat.Id);
                }
                else
                {
                    sourceNumberTemplateField = FindNearestMathTypePlaceRefField(
                        document,
                        insertion.Start,
                        excludeStart: -1,
                        excludeEnd: -1);
                    if (sourceNumberTemplateField is not null
                        && TryReadReusableMathTypePlaceRefTemplate(
                            document,
                            sourceNumberTemplateField,
                            out var reusableNumberTemplate))
                    {
                        numberTemplate = reusableNumberTemplate;
                    }
                    else
                    {
                        // A malformed legacy/direct-insert MTPlaceRef can retain only
                        // punctuation such as "(.)" while its nested MTEqn/MTChap
                        // fields have escaped outside the MACROBUTTON tree. Never
                        // clone that corruption into every later equation. The
                        // document-selected VisualTeX format is authoritative when
                        // no structurally complete native MathType template exists.
                        numberTemplate = MathTypeWordOpenXml.CreateVisualTeXNumberTemplate(
                            documentNumberFormat.Id);
                    }
                }

                // Number formatting and chapter/section state are two separate
                // pieces of MathType's native model.  The old code stopped here
                // whenever *any* nearby MTPlaceRef existed, so subsequent direct
                // insertions/conversions cloned the field template but never
                // established MTChap/MTSec state for their own Word heading scope.
                // Always reconcile the native heading state when this MTPlaceRef
                // actually uses chapter/section sequences.
                if (documentNumberFormat.UsesHeading
                    && numberTemplate is not null
                    && MathTypeNumberTemplateUsesHeading(numberTemplate))
                {
                    var logicalInsertionStart = insertion.Start;
                    var scopeAlreadyPrepared = preResolvedHeadingScope is not null
                        && preResolvedHeadingScope.ScopeStart != int.MinValue
                        && preparedHeadingScopeStarts?.Contains(
                            preResolvedHeadingScope.ScopeStart) == true;
                    if (!scopeAlreadyPrepared
                        && preResolvedHeadingScope?.ScopeStart != int.MinValue)
                    {
                        var insertedSectionLength = EnsureMathTypeHeadingScopeState(
                            document,
                            logicalInsertionStart,
                            documentNumberFormat,
                            out createdSectionBreakCodeStart,
                            preResolvedHeadingScope);
                        if (preResolvedHeadingScope is not null)
                            preparedHeadingScopeStarts?.Add(
                                preResolvedHeadingScope.ScopeStart);
                        if (insertedSectionLength > 0)
                        {
                            var shiftedStart = Math.Min(
                                document.Content.End,
                                logicalInsertionStart + insertedSectionLength);
                            insertion.SetRange(shiftedStart, shiftedStart);
                        }
                    }
                }
                TraceInsertPerf("number-template-heading-state");
            }

            if (!inline)
            {
                stage = "prepare-display-row";
                paragraphCountBeforeDisplayPreparation = ReadDocumentParagraphCount(document);
                displaySpacingAnchor = insertion.Duplicate;
                var displayInsertion = ResolveStandaloneMathTypeDisplayInsertionRange(
                    document,
                    insertion,
                    replaceAtExactInsertion: preserveExistingDisplayParagraphBoundary);
                Release(insertion);
                insertion = displayInsertion;
                TraceInsertPerf("prepare-display-row");
            }

            insertionStart = insertion.Start;
            var preservedRightBoundaryWhitespace =
                inline && preserveCapturedInsertion
                    ? CaptureLeadingHorizontalWhitespace(
                        document,
                        insertionStart)
                    : null;
            var relocateRightNumberToLeftAfterInsert =
                !inline
                && numberTemplate is not null
                && string.Equals(
                    session.MathTypeNumberPosition,
                    "left",
                    StringComparison.OrdinalIgnoreCase);
            stage = "build-flat-opc";
            // Word reliably preserves MathType's nested MTPlaceRef tree when the
            // number is imported on the right of the OLE. Importing the exact same
            // field tree on the left is not reliable in a real interactive Word
            // process: Word can promote the nested MTEqn/MTChap fields into sibling
            // document fields. Building those children later with Fields.Add has
            // the same nondeterministic failure mode. Therefore every left-numbered
            // direct insert is first materialized as the known-good right-numbered
            // Flat OPC structure. After Word has committed the complete outer field
            // atomically, relocate that already-healthy formatted field to the left.
            var wordOpenXml = MathTypeWordOpenXml.CreateWithPlaceableWmf(
                compoundFile,
                previewWmf,
                widthPt,
                heightPt,
                display: !inline,
                numberTemplate,
                mathTypeNumberPosition: relocateRightNumberToLeftAfterInsert
                    ? "right"
                    : session.MathTypeNumberPosition);
            TraceInsertPerf("build-flat-opc");

            // Resolve the inserted OLE from the exact mutation site for every path,
            // including ordinary interactive insertion. The previous interactive
            // safety check enumerated document.InlineShapes twice and read every
            // preceding shape.Range merely to predict the new ordinal. In a document
            // with N formulas that made each new insertion O(N), and a sequence of N
            // insertions O(N²). Word inserts one OLE character at the target range;
            // a bounded local probe plus the semantic CFB validation below is both
            // stricter and independent of total document size. Keep the global
            // nearest-position scan only as an exceptional compatibility fallback.
            sourceParagraphCount = ReadDocumentParagraphCount(document);
            TraceInsertPerf("pre-insert-bookkeeping");
            stage = "insert-flat-opc";
            var useIsolatedTableBoundaryTransfer =
                !inline
                && preserveExistingDisplayParagraphBoundary
                && string.Equals(
                    insertion.Text,
                    BulkInlineFormulaPlaceholder,
                    StringComparison.Ordinal);
            if (useIsolatedTableBoundaryTransfer)
            {
                InsertIsolatedMathTypeParagraphBody(
                    document,
                    insertion,
                    wordOpenXml);
                document.Activate();
                Release(selection);
                selection = _application.Selection;
            }
            else
            {
                insertion.InsertXML(wordOpenXml);
            }
            TraceInsertPerf(useIsolatedTableBoundaryTransfer
                ? "insert-formatted-body"
                : "insert-flat-opc");
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                && string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE"),
                    "after-flat-opc",
                    StringComparison.Ordinal))
                throw new COMException(
                    "Injected MathType Flat OPC failure for rollback acceptance.",
                    unchecked((int)0x8007000E));
            var insertedRangeEnd = Math.Min(
                document.Content.End,
                insertionStart + 8);
            var insertedRangeReference =
                $"{RangeReferencePrefix}{insertionStart}:{insertedRangeEnd}";
            shape = FindMathTypeOleByRange(
                    document,
                    insertedRangeReference,
                    allowGlobalFallback: false)
                ?? FindMathTypeOleInParagraphAtPosition(document, insertionStart)
                ?? FindMathTypeOleInLocalWindow(document, insertionStart)
                ?? FindMathTypeOleNearPosition(document, insertionStart);
            if (shape is null)
                throw new InvalidOperationException(
                    "Word inserted the MathType OLE data but VisualTeX could not resolve the new equation.");
            if (!MathTypeOleInterop.IsMathTypeOle(shape))
                throw new InvalidOperationException(
                    "Word did not materialize the standalone equation as Equation.DSMT4.");
            TraceInsertPerf("resolve-new-shape");
            stage = "validate-flat-opc-storage";
            // Interactive creation keeps the expensive post-insert Flat OPC
            // round-trip validation. Batch format conversion already validated
            // the exact standalone CFB before InsertXML and verifies the resulting
            // object class immediately above; serializing shape.Range.WordOpenXML
            // again for every converted equation costs several seconds per OLE in
            // large documents and turns a 50-formula conversion into minutes.
            if (!useLocalConversionLookup
                && retainMathTypeForCompletedValidation is null
                && MathTypeOleStorage.TryCaptureCompoundFileFromWordOpenXml(
                    shape,
                    out var materializedCompoundFile))
            {
                var materializedMathMl = MathTypeOleStorage.ReadMathMl(
                    materializedCompoundFile);
                if (!MathTypeMathMlRoundTripMatches(expectedSignature, materializedMathMl))
                    throw new InvalidDataException(
                        "Word materialized a different MathType equation than VisualTeX generated.");
            }

            // The Flat OPC already contains both VisualTeX's standalone MathType
            // CFB and its WMF/EMF presentation cache.  Do not round-trip the object
            // through Word PasteSpecial here: wdPasteOLEObject asks Windows to
            // instantiate the Equation.DSMT4 CLSID and therefore launches an
            // installed MathType OLE server.  Keeping this original InlineShape is
            // both sufficient for Word display and makes conversion fully offline.
            stage = "use-flat-opc-presentation";
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            shape.Width = widthPt;
            shape.Height = heightPt;
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            TraceInsertPerf("set-shape-geometry");

            if (sourceParagraphCount >= 0)
            {
                var repairParagraphSplit = inline;
                if (!repairParagraphSplit
                    && preserveExistingDisplayParagraphBoundary)
                {
                    Range? insertedShapeRange = null;
                    try
                    {
                        insertedShapeRange = shape.Range;
                        repairParagraphSplit =
                            WordEquationNumbering.RangeIsWhollyWithinTable(
                                insertedShapeRange);
                    }
                    finally { Release(insertedShapeRange); }
                }
                if (repairParagraphSplit)
                    RepairMathTypeInsertXmlParagraphSplit(
                        document,
                        shape,
                        sourceParagraphCount,
                        preservedRightBoundaryWhitespace);
            }

            TraceInsertPerf("repair-insertxml-paragraph");

            // Word structural edits can move the InlineShape range. Re-resolve it
            // before the final baseline/numbering work without activating its OLE
            // server.
            stage = "refresh-native-shape";
            var refreshedRangeEnd = Math.Min(
                document.Content.End,
                insertionStart + 8);
            var refreshedShape = FindMathTypeOleByRange(
                    document,
                    $"{RangeReferencePrefix}{insertionStart}:{refreshedRangeEnd}",
                    allowGlobalFallback: false)
                ?? FindMathTypeOleInParagraphAtPosition(document, insertionStart)
                ?? FindMathTypeOleInLocalWindow(document, insertionStart)
                ?? FindMathTypeOleNearPosition(document, insertionStart);
            if (refreshedShape is null)
                throw new InvalidOperationException(
                    "Word retained the MathType OLE but VisualTeX could not refresh its live InlineShape handle.");
            Release(shape);
            shape = refreshedShape;
            TraceInsertPerf("refresh-native-shape");

            if (inline)
            {
                stage = "apply-native-baseline";
                SetInlineOleWordPosition(shape, wordPosition);
                TraceInsertPerf("set-native-baseline");
                if (!reuseExistingInlineTypingBoundary)
                {
                    if (useLocalConversionLookup)
                        RestoreTypingBaselineAfterMathTypeConversion(shape);
                    else
                        RestoreTypingBaselineAfter(shape);
                }
                else if (!useLocalConversionLookup)
                {
                    // The typing-baseline restorers above already place the caret at
                    // shape.Range.End and copy the surrounding prose character format
                    // onto Word's insertion point. Calling Selection.SetRange again
                    // after that successful restoration makes Word re-inherit the OLE
                    // field run (typically the MathType formula size, e.g. 10.5 pt),
                    // so prose typed after a 12 pt paragraph silently drops back to
                    // the formula/default size. Only fall back to a raw caret move
                    // when the caller explicitly asked to reuse an existing boundary
                    // and therefore skipped the formatting restoration entirely.
                    var shapeRange = shape.Range;
                    try { selection.SetRange(shapeRange.End, shapeRange.End); }
                    finally { Release(shapeRange); }
                }
                TraceInsertPerf("restore-typing-baseline");
            }
            else
            {
                if (relocateRightNumberToLeftAfterInsert)
                {
                    stage = "relocate-left-mathtype-field-scaffold";
                    RelocateCompleteMathTypePlaceRefFromRightToLeft(
                        document,
                        shape);
                    TraceInsertPerf("relocate-left-mathtype-field-scaffold");
                }
                stage = "configure-display-numbering";
                ConfigureNewMathTypeDisplayEquation(
                    document,
                    shape,
                    session.Numbered,
                    session.MathTypeNumberPosition,
                    updateNestedNumberFields:
                        !useLocalConversionLookup || updateCreatedMathTypeNumberFields,
                    bodyFormatting: bodyFormatting,
                    preservedDisplayParagraphLayout: preservedDisplayParagraphLayout,
                    knownDisplayColumnWidth: knownDisplayColumnWidth);
                TraceInsertPerf("configure-display-numbering");
                // MTDisplayEquation setup resets direct character formatting on
                // the OLE run. Apply the exported math baseline after numbering
                // and paragraph style are final so the adjacent number stays
                // vertically aligned with tall display formulas.
                stage = "apply-native-display-baseline";
                SetInlineOleWordPosition(shape, wordPosition);
                TraceInsertPerf("display-baseline");
                var shapeRange = shape.Range;
                try
                {
                    if (useLocalConversionLookup && preserveExistingDisplayParagraphBoundary)
                    {
                        // The conversion caller owns a durable VTMT locator and the
                        // existing display paragraph boundary is already complete.
                        // Selecting this OLE here costs a full Word layout/UI turn
                        // and is repeated again by subsequent user interaction.
                    }
                    else if (session.Numbered && preserveExistingDisplayParagraphBoundary)
                        selection.SetRange(shapeRange.Start, shapeRange.End);
                    else if (preserveExistingDisplayParagraphBoundary)
                        selection.SetRange(shapeRange.End, shapeRange.End);
                    else
                    {
                        var typing = WordEquationNumbering.EnsureBodyTypingParagraphAfterDisplay(
                            document, shapeRange, bodyFormatting!);
                        try { selection.SetRange(typing.Start, typing.End); }
                        finally { Release(typing); }
                    }
                }
                finally { Release(shapeRange); }
            }

            TraceInsertPerf("typing-paragraph-selection");
            if (!inline && displaySpacingAnchor is not null)
            {
                stage = "finalize-display-spacing";
                CompactParagraphBeforeOleDisplayFormula(document, displaySpacingAnchor);
            }
            TraceInsertPerf("display-spacing");
            if (!string.IsNullOrWhiteSpace(createdObjectBookmarkName))
            {
                stage = "bind-created-object-identity";
                TryDeleteBookmark(document, createdObjectBookmarkName);
                createdObjectBookmarkRange = shape.Range;
                createdObjectBookmark = document.Bookmarks.Add(
                    createdObjectBookmarkName,
                    createdObjectBookmarkRange);
                TraceInsertPerf("bind-created-object-identity");
            }
            if (retainMathTypeForCompletedValidation is not null)
            {
                var verificationRange = shape.Range;
                try { retainMathTypeForCompletedValidation(verificationRange, expectedSignature); }
                finally { Release(verificationRange); }
            }

            // Keep VisualTeX ownership outside Equation.DSMT4/MTEF. The bookmark
            // is a locator only; MathType storage remains completely native.
            stage = "bind-mathtype-host-identity";
            WordFormulaIdentityStore.BindMathType(
                shape,
                metadata.FormulaId);

            stage = "complete";
            TraceInsertPerf("complete");
            return Result(
                session,
                document,
                documentIdentityAlreadyValidated: true);
        }
        catch (Exception error)
        {
            if (document is not null)
                WordFormulaIdentityStore.RemoveMathType(
                    document,
                    metadata.FormulaId);
            if (document is not null && !string.IsNullOrWhiteSpace(createdObjectBookmarkName))
                TryDeleteBookmark(document, createdObjectBookmarkName);
            if (!inline && document is not null && insertionStart >= 0)
            {
                RollbackStandaloneMathTypeDisplayInsertion(
                    document,
                    insertionStart,
                    paragraphCountBeforeDisplayPreparation);
            }
            else
            {
                TryDelete(shape);
            }
            if (createdSectionBreakCodeStart >= 0 && document is not null)
                RemoveMathTypeSectionBreakFieldAtCodeStart(
                    document,
                    createdSectionBreakCodeStart);
            var hresult = error is COMException
                ? $" HRESULT=0x{error.HResult:X8}."
                : string.Empty;
            throw new InvalidOperationException(
                $"MathType OLE insertion failed at stage '{stage}'.{hresult} {error.Message}",
                error);
        }
        finally
        {
            nativePreview?.Dispose();
            TraceInsertPerf("finally-preview");
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            TraceInsertPerf("finally-undo");
            Release(createdObjectBookmarkRange);
            Release(createdObjectBookmark);
            Release(sourceNumberTemplateField);
            TraceInsertPerf("finally-bookkeeping");
            Release(shape);
            TraceInsertPerf("finally-shape");
            Release(displaySpacingAnchor);
            Release(insertion);
            TraceInsertPerf("finally-ranges");
            Release(selection);
            TraceInsertPerf("finally-selection");
            Release(document);
            TraceInsertPerf("finally-document");
        }
    }

    // Legacy server-backed writer kept temporarily for focused diagnostics only.
    // Production insertion/conversion must never activate the MathType UI.
    public OfficeObjectResult InsertOmml(
        OfficeSessionDocument session,
        string mathMl,
        bool deferNumberingLayout = false,
        bool numberingScaffoldOnly = false,
        bool deferFinalFingerprint = false,
        WordOmmlConverter.BatchSource? ommlBatchSource = null,
        bool preserveExistingDisplayParagraphBoundary = false,
        bool normalizeMathTypeDisplayParagraph = false,
        Action<Range>? retainInsertedMath = null,
        bool preserveCapturedInsertion = false)
    {
        if (deferNumberingLayout
            || numberingScaffoldOnly
            || deferFinalFingerprint
            || ommlBatchSource is not null
            || preserveExistingDisplayParagraphBoundary
            || normalizeMathTypeDisplayParagraph
            || retainInsertedMath is not null
            || preserveCapturedInsertion)
            throw new NotSupportedException(
                "Retired OMML insertion compatibility flags are not supported by the rebuilt Word host core.");

        CaptureCurrentSelectionForCreateSession(session);
        session.Mode = "create";
        session.ObjectMode =
            FormulaOleContract.WordOmmlMode;
        return ApplyOmmlVisualTeXHostSession(
            session,
            sourceObjectMode: null,
            mathMl,
            pngPath: null,
            emfPath: null);
    }

    private static IReadOnlyList<InlineLatexBaselineProvenance>
        CaptureInlineLatexBaselineProvenance(Document document, Range scope)
    {
        var result = new List<InlineLatexBaselineProvenance>();
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmark);
                bookmark = null;
                Release(range);
                range = null;
                bookmark = bookmarks[index];
                if (!TryParseInlineLatexBaselineBookmarkName(
                        bookmark.Name,
                        out var wordPosition,
                        out var bottomWhitespacePoints))
                    continue;
                range = bookmark.Range;
                if (range.StoryType != scope.StoryType
                    || range.End <= scope.Start
                    || range.Start >= scope.End)
                    continue;
                result.Add(new InlineLatexBaselineProvenance(
                    range.Start,
                    range.End,
                    wordPosition,
                    bottomWhitespacePoints));
            }
            return result;
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static InlineLatexBaselineProvenance? ResolveInlineLatexBaselineProvenance(
        IReadOnlyList<InlineLatexBaselineProvenance> provenance,
        int sourceStart,
        int sourceEnd)
    {
        var matches = provenance
            .Where(item => item.Start == sourceStart && item.End == sourceEnd)
            .Take(2)
            .ToArray();
        return matches.Length == 1 ? matches[0] : null;
    }

    public WordLatexRedrawPlan CaptureLatexRedrawPlan(bool wholeDocument)
    {
        return CaptureLatexRedrawPlanCore(
            wholeDocument);
    }

    public WordLatexRedrawResult ApplyLatexRedrawPlan(
        WordLatexRedrawPlan plan,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared)
    {
        return ApplyLatexRedrawPlanCore(
            plan,
            prepared);
    }

    public int CountFormulaObjectsForLatex(bool wholeDocument, string objectMode)
    {
        return string.Equals(
                objectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal)
            ? CountFormulaObjectsForLatexLegacy(
                wholeDocument,
                objectMode)
            : CountFormulaObjectsForLatexCore(
                wholeDocument,
                objectMode);
    }

    private int CountFormulaObjectsForLatexLegacy(bool wholeDocument, string objectMode)
    {
        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        List<FormulaToLatexTarget>? targets = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            selection = _application.Selection;
            scope = wholeDocument
                ? document.Content.Duplicate
                : selection.Range.Duplicate;
            targets = CaptureFormulaToLatexTargets(
                document,
                scope,
                wholeDocument,
                objectMode,
                refreshOmmlMetadata: false);
            return targets.Count;
        }
        finally
        {
            ReleaseFormulaToLatexTargets(targets);
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    public WordFormulaToLatexResult ConvertFormulaObjectsToLatex(
        bool wholeDocument,
        string objectMode)
    {
        return string.Equals(
                objectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal)
            ? ConvertFormulaObjectsToLatexLegacy(
                wholeDocument,
                objectMode)
            : ConvertFormulaObjectsToLatexCore(
                wholeDocument,
                objectMode);
    }

    private WordFormulaToLatexResult ConvertFormulaObjectsToLatexLegacy(
        bool wholeDocument,
        string objectMode)
    {
        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        UndoRecord? undoRecord = null;
        WordViewState? viewState = null;
        List<FormulaToLatexTarget>? targets = null;
        WordDocumentEditSnapshot? rollbackSnapshot = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            selection = _application.Selection;
            scope = wholeDocument
                ? document.Content.Duplicate
                : selection.Range.Duplicate;
            targets = CaptureFormulaToLatexTargets(
                document,
                scope,
                wholeDocument,
                objectMode,
                refreshOmmlMetadata: true);
            if (targets.Count == 0)
            {
                var modeLabel = string.Equals(
                    objectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal)
                    ? "VisualTeX OLE"
                    : string.Equals(
                        objectMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal)
                        ? "MathType"
                        : "Word OMML";
                throw new InvalidDataException(
                    wholeDocument
                        ? $"当前 Word 文档中没有找到可转换的 {modeLabel} 公式。"
                        : $"所选内容中没有找到可转换的 {modeLabel} 公式。");
            }

            // Preflight every source before opening a destructive Word undo record.
            // Corrupt/empty metadata must never turn a visible formula into an empty
            // replacement merely because the object itself can still be located.
            foreach (var target in targets)
            {
                ApplyFormulaToLatexEmptySourceInjection(target);
                target.LatexSource = BuildFormulaLatexSource(target.Metadata);
            }

            // A single native Undo is not a sufficient recovery contract here.
            // Word can split the custom record while dismantling an OMath/table
            // and leave a harmless-looking body character or stale ownership
            // structure behind. Capture the same verified document evidence used
            // by format conversion so a failed formula-to-LaTeX transaction must
            // restore text, object geometry, bookmarks, metadata and variables.
            rollbackSnapshot = new WordDocumentEditSnapshot(
                document,
                targets.SelectMany(target =>
                    WordBookmarkRecoverySnapshot.NamesForFormula(
                        target.Metadata.FormulaId)));

            viewState = CaptureViewState();
            undoRecord = BeginUndoRecord("VisualTeX 公式转为 LaTeX 代码");
            if (undoRecord is null)
                throw new InvalidOperationException(
                    "Word 无法建立公式转 LaTeX 的撤销事务。为避免公式丢失，本次转换未开始。");

            var undoRecordEnded = false;
            var documentMutationStarted = false;
            try
            {
                var convertedIds = new List<string>(targets.Count);
                foreach (var target in targets.OrderByDescending(item => item.Start))
                {
                    ConvertFormulaTargetToLatex(
                        document,
                        target,
                        ref documentMutationStarted);
                    convertedIds.Add(target.Metadata.FormulaId);
                }
                WordEquationNumbering.TryReconcile(document);
                return new WordFormulaToLatexResult
                {
                    FormulaCount = convertedIds.Count,
                    FormulaIds = convertedIds,
                };
            }
            catch (Exception conversionError)
            {
                EndUndoRecord(undoRecord);
                undoRecordEnded = true;
                if (documentMutationStarted)
                {
                    try { rollbackSnapshot!.Restore(document); }
                    catch (Exception recoveryError)
                    {
                        throw new AggregateException(
                            "公式转 LaTeX 失败，而且 Word 无法完整恢复转换前的文档结构。请保留当前文档以便排查。",
                            conversionError,
                            recoveryError);
                    }
                }
                throw;
            }
            finally
            {
                if (!undoRecordEnded)
                    EndUndoRecord(undoRecord);
            }
        }
        finally
        {
            RestoreViewState(document, viewState, preferredSelection: null);
            Release(undoRecord);
            ReleaseFormulaToLatexTargets(targets);
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    private static List<FormulaToLatexTarget> CaptureFormulaToLatexTargets(
        Document document,
        Range scope,
        bool wholeDocument,
        string objectMode,
        bool refreshOmmlMetadata,
        ISet<string>? ignoredDeletedNumberedOmmlIds = null)
    {
        _ = refreshOmmlMetadata;
        _ = ignoredDeletedNumberedOmmlIds;

        if (!string.Equals(
                objectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal))
            throw new InvalidOperationException(
                "The legacy formula-to-LaTeX capture path is MathType-only.");

        using var operationMetric =
            VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure(
                "WordFormulaService.CaptureMathTypeFormulaToLatexTargets");

        var targets =
            new List<FormulaToLatexTarget>();
        InlineShapes? shapes = null;
        try
        {
            shapes = wholeDocument
                ? document.InlineShapes
                : scope.InlineShapes;

            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                InlineShape? shape = null;
                Range? formulaRange = null;
                try
                {
                    shape = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(
                            shape))
                        continue;

                    formulaRange =
                        shape.Range.Duplicate;
                    if (!FormulaRangeMatchesScope(
                            formulaRange,
                            scope,
                            wholeDocument))
                        continue;

                    var sourceMathMl =
                        MathTypeOleStorage.ReadMathMl(
                            shape);
                    var metadata =
                        MathTypeOleInterop.ReadMetadata(
                            document.Application,
                            shape,
                            sourceMathMl);

                    int? sourceInlineWordPosition =
                        null;
                    float? sourceInlineBottomWhitespacePoints =
                        null;
                    if (string.Equals(
                            metadata.DisplayMode,
                            "inline",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        sourceInlineWordPosition =
                            ReadInlineOleWordPosition(
                                shape);
                        var sourcePreviewMetrics =
                            TryMeasureInlineOlePreview(
                                shape);
                        if (sourcePreviewMetrics.HasValue)
                        {
                            sourceInlineBottomWhitespacePoints =
                                shape.Height
                                * sourcePreviewMetrics.Value
                                    .BottomWhitespaceRatio;
                        }
                    }

                    targets.Add(
                        new FormulaToLatexTarget
                        {
                            Metadata = metadata,
                            ObjectMode =
                                FormulaOleContract.MathTypeOleMode,
                            Start = formulaRange.Start,
                            End = formulaRange.End,
                            FormulaRange = formulaRange,
                            OleShape = shape,
                            SourceInlineWordPosition =
                                sourceInlineWordPosition,
                            SourceInlineBottomWhitespacePoints =
                                sourceInlineBottomWhitespacePoints,
                        });
                    formulaRange = null;
                    shape = null;
                }
                finally
                {
                    Release(formulaRange);
                    Release(shape);
                }
            }

            return targets;
        }
        catch
        {
            ReleaseFormulaToLatexTargets(
                targets);
            throw;
        }
        finally
        {
            Release(shapes);
        }
    }

    private static bool FormulaRangesMatch(Range left, Range right) =>
        left.Start == right.Start && left.End == right.End
        || left.Start <= right.Start && left.End >= right.End
        || right.Start <= left.Start && right.End >= left.End;

    private static bool FormulaRangeMatchesScope(
        Range formulaRange,
        Range scope,
        bool wholeDocument)
    {
        if (wholeDocument) return true;
        var scopeStart = scope.Start;
        var scopeEnd = scope.End;
        var formulaStart = formulaRange.Start;
        var formulaEnd = formulaRange.End;
        if (scopeStart == scopeEnd)
            return scopeStart >= formulaStart && scopeStart <= formulaEnd;
        return formulaStart >= scopeStart && formulaEnd <= scopeEnd;
    }

    private void ConvertFormulaTargetToLatex(
        Document document,
        FormulaToLatexTarget target,
        ref bool documentMutationStarted,
        IReadOnlyDictionary<string, int>? knownReferenceCounts = null,
        bool preserveCrossReferences = false,
        bool deferCaptionFlowTailCleanup = false)
    {
        // The legacy formula-to-LaTeX pipeline is now MathType-only. OMML and
        // VisualTeX public routes are handled by FormulaToLatexCore.
        if (!string.Equals(
                target.ObjectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal))
            throw new InvalidOperationException(
                "OMML/VisualTeX formula-to-LaTeX must use the rebuilt host core.");

        var metadata = target.Metadata;
        var latexSource = target.LatexSource;
        if (string.IsNullOrWhiteSpace(latexSource))
            throw new InvalidDataException(
                $"公式 {metadata.FormulaId} 没有可安全恢复的 LaTeX 源码。");
        if (target.OleShape is null)
            throw new InvalidDataException(
                "MathType 公式转 LaTeX 时无法定位源 Equation.DSMT4 对象。");

        Range? liveSourceRange = null;
        Range? insertion = null;
        Range? inserted = null;
        Range? deleteAnchor = null;
        try
        {
            documentMutationStarted = true;
            var referenceAliases =
                metadata.Numbered
                    ? MathTypeEquationReferences
                        .CaptureReferenceBookmarkAliases(
                            document,
                            target.OleShape)
                    : Array.Empty<string>();
            if (!preserveCrossReferences
                && referenceAliases.Count > 0)
            {
                MathTypeEquationReferences
                    .FreezeReferencesToPlainText(
                        document,
                        referenceAliases);
            }

            liveSourceRange =
                target.OleShape.Range.Duplicate;
            var source =
                WordMathTypeHostAdapter.ResolveLocal(
                    _application,
                    document,
                    liveSourceRange)
                ?? throw new InvalidDataException(
                    "The MathType source moved before conversion to LaTeX.");

            deleteAnchor =
                WordMathTypeHostAdapter.DeleteExact(
                    document,
                    source);
            var insertionStart =
                deleteAnchor.Start;

            ThrowIfFormulaToLatexFailureInjected(
                target);
            insertion =
                document.Range(
                    insertionStart,
                    insertionStart);
            insertion.Text = latexSource;
            inserted =
                document.Range(
                    insertionStart,
                    insertionStart
                    + latexSource.Length);
            VerifyLatexSourceRange(
                inserted,
                latexSource,
                metadata.FormulaId);
            NormalizeLatexSourceRange(
                inserted,
                metadata);
            BindInlineLatexBaselineProvenance(
                document,
                inserted,
                target.SourceInlineWordPosition,
                target.SourceInlineBottomWhitespacePoints);
        }
        finally
        {
            Release(deleteAnchor);
            Release(inserted);
            Release(insertion);
            Release(liveSourceRange);
        }
    }

    private static void DetachLatexSourceFromVisualTeXNumberingFrame(
        Range inserted,
        FormulaMetadata metadata)
    {
        if (!metadata.Numbered) return;
        Frames? frames = null;
        try
        {
            frames = inserted.Frames;
            for (var index = frames.Count; index >= 1; index--)
            {
                Frame? frame = null;
                try
                {
                    frame = frames[index];
                    var clippedVisualTeXCaptionFrame =
                        frame.Width <= 1f
                        && frame.Height <= 1f
                        && frame.LockAnchor
                        && !frame.TextWrap
                        && frame.RelativeHorizontalPosition
                            == WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                        && frame.RelativeVerticalPosition
                            == WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    if (clippedVisualTeXCaptionFrame)
                        frame.Delete();
                }
                finally { Release(frame); }
            }
        }
        finally { Release(frames); }

        Frames? remaining = null;
        try
        {
            remaining = inserted.Frames;
            for (var index = 1; index <= remaining.Count; index++)
            {
                Frame? frame = null;
                try
                {
                    frame = remaining[index];
                    if (frame.Width <= 1f && frame.Height <= 1f)
                        throw new InvalidDataException(
                            $"公式 {metadata.FormulaId} 的 LaTeX 源码仍位于隐藏编号框架中。为避免源码不可见，转换已回滚。");
                }
                finally { Release(frame); }
            }
        }
        finally { Release(remaining); }
    }

    private static string BuildFormulaLatexSource(FormulaMetadata metadata)
    {
        var latex = string.IsNullOrWhiteSpace(metadata.Latex)
            ? string.Join("\n", metadata.Lines.Select(line => line.Latex))
            : metadata.Latex;
        latex = (latex ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Trim();
        if (string.IsNullOrWhiteSpace(latex))
            throw new InvalidDataException(
                $"公式 {metadata.FormulaId} 的 LaTeX 元数据为空。为避免删除原公式，转换已中止。");
        if (string.Equals(
                metadata.DisplayMode,
                "block",
                StringComparison.Ordinal))
        {
            latex = FormulaEquationTag.Attach(latex, metadata.EquationTag);
            return "$$" + latex + "$$";
        }
        latex = FormulaEquationTag.Extract(latex).Latex
            .Replace('\n', ' ');
        return "$" + latex + "$";
    }

    private static void BindInlineLatexBaselineProvenance(
        Document document,
        Range latexRange,
        int? wordPosition,
        float? bottomWhitespacePoints)
    {
        if (!wordPosition.HasValue) return;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            var clamped = Math.Max(-256, Math.Min(256, wordPosition.Value));
            var sign = clamped < 0 ? "N" : "P";
            var bottomToken = string.Empty;
            if (bottomWhitespacePoints.HasValue
                && !float.IsNaN(bottomWhitespacePoints.Value)
                && !float.IsInfinity(bottomWhitespacePoints.Value)
                && bottomWhitespacePoints.Value >= 0)
            {
                var tenths = Math.Max(
                    0,
                    Math.Min(
                        999,
                        (int)Math.Round(
                            bottomWhitespacePoints.Value * 10f,
                            MidpointRounding.AwayFromZero)));
                bottomToken = $"_B{tenths:D3}";
            }
            var unique = Guid.NewGuid().ToString("N").Substring(0, 20);
            var name =
                $"{InlineLatexBaselineBookmarkPrefix}{sign}{Math.Abs(clamped):D3}{bottomToken}_{unique}";
            bookmarks = document.Bookmarks;
            bookmark = bookmarks.Add(name, latexRange);
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static bool TryParseInlineLatexBaselineBookmarkName(
        string? name,
        out int wordPosition,
        out float? bottomWhitespacePoints)
    {
        wordPosition = 0;
        bottomWhitespacePoints = null;
        if (string.IsNullOrWhiteSpace(name)
            || !name.StartsWith(InlineLatexBaselineBookmarkPrefix, StringComparison.Ordinal))
            return false;
        var payload = name.Substring(InlineLatexBaselineBookmarkPrefix.Length);
        if (payload.Length < 4
            || (payload[0] != 'N' && payload[0] != 'P')
            || !int.TryParse(payload.Substring(1, 3), out var magnitude))
            return false;
        wordPosition = payload[0] == 'N' ? -magnitude : magnitude;
        if (payload.Length >= 9
            && payload[4] == '_'
            && payload[5] == 'B'
            && int.TryParse(payload.Substring(6, 3), out var bottomWhitespaceTenths))
        {
            bottomWhitespacePoints = bottomWhitespaceTenths / 10f;
        }
        return wordPosition >= -256 && wordPosition <= 256;
    }

    private static string NormalizeFormulaToLatexVerificationText(string value) =>
        (value ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Replace('\v', '\n');

    private static void TraceFormulaToLatexBridgeState(
        Document document,
        Range inserted,
        string expected,
        string formulaId,
        string stage)
    {
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_LATEX_BRIDGE"),
                "1",
                StringComparison.Ordinal))
            return;

        Frames? frames = null;
        Range? content = null;
        try
        {
            var actual = inserted.Text ?? string.Empty;
            frames = inserted.Frames;
            content = document.Content;
            var documentText = content.Text ?? string.Empty;
            var expectedOffset = documentText.IndexOf(expected, StringComparison.Ordinal);
            var expectedStart = expectedOffset >= 0
                ? content.Start + expectedOffset
                : -1;
            WordDoubleClickHook.TraceMessage(
                $"formula-to-latex-bridge-state stage={stage} formulaId={formulaId} range={inserted.Start}:{inserted.End} actualLength={actual.Length} frames={frames.Count} expectedStart={expectedStart}");
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage(
                $"formula-to-latex-bridge-state-failed stage={stage} formulaId={formulaId} error={error.Message}");
        }
        finally
        {
            Release(content);
            Release(frames);
        }
    }

    private static void VerifyLatexSourceRange(
        Range inserted,
        string expected,
        string formulaId)
    {
        var actual = NormalizeFormulaToLatexVerificationText(
            inserted.Text ?? string.Empty);
        var normalizedExpected = NormalizeFormulaToLatexVerificationText(expected);
        if (!string.Equals(actual, normalizedExpected, StringComparison.Ordinal))
        {
            var actualCodes = string.Join(
                ",",
                actual.Take(128).Select(ch => $"U+{(int)ch:X4}"));
            var expectedCodes = string.Join(
                ",",
                normalizedExpected.Take(128).Select(ch => $"U+{(int)ch:X4}"));
            WordDoubleClickHook.TraceMessage(
                $"formula-to-latex-verify-mismatch formulaId={formulaId} range={inserted.Start}:{inserted.End} actualLength={actual.Length} expectedLength={normalizedExpected.Length} actualCodes={actualCodes} expectedCodes={expectedCodes}");
            throw new InvalidDataException(
                $"公式 {formulaId} 的 LaTeX 写回校验失败。Word 实际写入内容与预期源码不一致。");
        }

        Frames? frames = null;
        try
        {
            frames = inserted.Frames;
            if (frames.Count > 0)
                throw new InvalidDataException(
                    $"公式 {formulaId} 的 LaTeX 源码被 Word Frame 包围。为避免源码不可见，本次转换已撤销。");
        }
        finally { Release(frames); }
    }

    private static void ApplyFormulaToLatexEmptySourceInjection(
        FormulaToLatexTarget target)
    {
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            return;
        var requested = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_FORMULA_TO_LATEX_EMPTY_SOURCE");
        if (string.IsNullOrWhiteSpace(requested)) return;
        if (!string.Equals(requested, "1", StringComparison.Ordinal)
            && !string.Equals(
                requested,
                target.Metadata.FormulaId,
                StringComparison.OrdinalIgnoreCase))
            return;
        target.Metadata.Latex = string.Empty;
        foreach (var line in target.Metadata.Lines)
            line.Latex = string.Empty;
    }

    private static void ThrowIfFormulaToLatexFailureInjected(
        FormulaToLatexTarget target)
    {
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            return;
        var requested = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_FORMULA_TO_LATEX_FAIL_AFTER_DELETE");
        if (string.IsNullOrWhiteSpace(requested)) return;
        if (!string.Equals(requested, "1", StringComparison.Ordinal)
            && !string.Equals(
                requested,
                target.Metadata.FormulaId,
                StringComparison.OrdinalIgnoreCase))
            return;
        throw new InvalidOperationException(
            "Injected formula-to-LaTeX failure after deleting the source object.");
    }

    private static Table? TryGetVisualTeXNumberedTable(
        Document document,
        Range formulaRange,
        FormulaMetadata metadata)
    {
        if (!metadata.Numbered || string.IsNullOrWhiteSpace(metadata.FormulaId))
            return null;
        Table? table = null;
        Range? owner = null;
        Columns? columns = null;
        try
        {
            // A Word formula being inside a three-column table is not evidence that
            // the table belongs to VisualTeX. Real user tables are valid formula
            // containers. The shared resolver normally proves ownership through the
            // visible-number bookmark, and conservatively recovers only when the
            // formula's own identity plus the generated VTEqn field prove the exact
            // same row after Word has drifted that bookmark into column 2.
            table = WordEquationNumbering.FindNumberedEquationTableByFormulaOwner(
                document,
                formulaRange,
                metadata.FormulaId);
            if (table is null) return null;
            owner = table.Range;
            columns = table.Columns;
            if (columns.Count != 3
                || owner.StoryType != formulaRange.StoryType
                || formulaRange.Start < owner.Start
                || formulaRange.End > owner.End
                || !WordEquationNumbering.TryGetManagedNumberTableRowIndex(
                    table,
                    formulaRange,
                    expectedColumnIndex: 2,
                    out _))
                return null;

            var result = table;
            table = null;
            return result;
        }
        catch { return null; }
        finally
        {
            Release(columns);
            Release(owner);
            Release(table);
        }
    }

    private static Table? TryGetVisualTeXNumberedTable(
        Range formulaRange,
        FormulaMetadata metadata)
    {
        Document? document = null;
        try
        {
            document = formulaRange.Document;
            return TryGetVisualTeXNumberedTable(document, formulaRange, metadata);
        }
        finally { Release(document); }
    }

    private static void NormalizeFormerNumberedFormulaParagraph(
        Document document,
        Range latexRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefix = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        try
        {
            paragraphs = latexRange.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (latexRange.Start > paragraphRange.Start)
            {
                prefix = document.Range(paragraphRange.Start, latexRange.Start);
                var prefixText = prefix.Text ?? string.Empty;
                if (prefixText.All(character => character is '\t' or '\v' or ' '))
                    prefix.Delete();
            }
            format = paragraphRange.ParagraphFormat;
            format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            // A numbered 1x3 formula can be adjacent to VisualTeX's mandatory
            // compact table separator (1pt exact line spacing). When the table is
            // removed during OMML→LaTeX conversion Word may reuse that paragraph's
            // pPr for the newly inserted visible source text. Font normalization
            // alone is not enough: 10.5pt LaTeX rendered inside a 1pt exact line
            // box is visibly crushed. Repair only this unmistakable compact-tail
            // signature; preserve ordinary/custom user line spacing otherwise.
            if (format.LineSpacingRule == WdLineSpacing.wdLineSpaceExactly
                && format.LineSpacing <= 2.01f)
            {
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                format.SpaceBefore = 0f;
                format.SpaceAfter = 0f;
                try { format.DisableLineHeightGrid = -1; } catch { }
            }
            tabStops = format.TabStops;
            tabStops.ClearAll();
        }
        finally
        {
            Release(tabStops);
            Release(format);
            Release(prefix);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void NormalizeLatexSourceRange(
        Range range,
        FormulaMetadata metadata)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? paragraphFormat = null;
        try
        {
            font = range.Font;
            font.Hidden = 0;
            font.Position = 0;
            font.Superscript = 0;
            font.Subscript = 0;
            var size = FormulaFontSize.ResolveSemanticFontSize(metadata);
            if (size > 0) font.Size = (float)size;

            // Also protect unnumbered/cross-format LaTeX restores that happen to
            // land in a compact VisualTeX structural paragraph. Only the 1-2pt
            // exact-spacing sentinel is normalized, so intentional user paragraph
            // formatting remains untouched.
            paragraphs = range.Paragraphs;
            if (paragraphs.Count == 1)
            {
                paragraph = paragraphs[1];
                paragraphFormat = paragraph.Range.ParagraphFormat;
                if (paragraphFormat.LineSpacingRule == WdLineSpacing.wdLineSpaceExactly
                    && paragraphFormat.LineSpacing <= 2.01f)
                {
                    paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                    paragraphFormat.SpaceBefore = 0f;
                    paragraphFormat.SpaceAfter = 0f;
                    try { paragraphFormat.DisableLineHeightGrid = -1; } catch { }
                }
            }
        }
        finally
        {
            Release(paragraphFormat);
            Release(paragraph);
            Release(paragraphs);
            Release(font);
        }
    }

    private static void ReleaseFormulaToLatexTargets(
        IEnumerable<FormulaToLatexTarget>? targets)
    {
        if (targets is null) return;
        foreach (var target in targets)
        {
            Release(target.OmmlBookmark);
            Release(target.OleShape);
            Release(target.FormulaRange);
        }
    }

    private static List<ResolvedLatexRedrawTarget> ResolveLatexRedrawTargets(
        Document document,
        WordLatexRedrawPlan plan,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared)
    {
        var resolved = new List<ResolvedLatexRedrawTarget>(plan.Targets.Count);
        try
        {
            foreach (var target in plan.Targets.OrderBy(item => item.RelativeStart))
            {
                if (!prepared.TryGetValue(target.Id, out var formula))
                    throw new InvalidDataException($"缺少公式 {target.Id} 的渲染结果。");
                if (target.RelativeStart < 0
                    || target.SourceLength <= 0
                    || target.RelativeStart + target.SourceLength > plan.SourceText.Length)
                    throw new InvalidDataException($"公式 {target.Id} 的源文本范围无效。");

                var expectedSource = plan.SourceText.Substring(
                    target.RelativeStart,
                    target.SourceLength);
                var sourceRange = ResolveExactLatexSourceRange(
                    document,
                    plan,
                    target,
                    expectedSource,
                    resolved);
                resolved.Add(new ResolvedLatexRedrawTarget
                {
                    Target = target,
                    Formula = formula,
                    SourceRange = sourceRange,
                    SourceStart = sourceRange.Start,
                    SourceEnd = sourceRange.End,
                    ExpectedSource = expectedSource,
                });
            }
            return resolved;
        }
        catch
        {
            foreach (var item in resolved)
                Release(item.SourceRange);
            throw;
        }
    }

    private static void PopulateMathTypeRedrawDisplayColumnWidths(
        Document document,
        IReadOnlyList<ResolvedLatexRedrawTarget> targets)
    {
        if (!targets.Any(item => string.Equals(
                item.Target.DisplayMode,
                "block",
                StringComparison.Ordinal)))
            return;

        var sectionLayouts = CaptureMathTypeSectionColumnLayouts(document);
        foreach (var target in targets)
        {
            if (!string.Equals(
                    target.Target.DisplayMode,
                    "block",
                    StringComparison.Ordinal))
                continue;
            target.MathTypeDisplayColumnWidth = ResolveMathTypeDisplayColumnWidthFromLayouts(
                document,
                target.SourceRange,
                sectionLayouts);
        }
    }

    private static List<MathTypeSectionColumnLayout> CaptureMathTypeSectionColumnLayouts(
        Document document)
    {
        Sections? sections = null;
        Section? section = null;
        Range? sectionRange = null;
        PageSetup? pageSetup = null;
        TextColumns? textColumns = null;
        TextColumn? textColumn = null;
        var layouts = new List<MathTypeSectionColumnLayout>();
        try
        {
            sections = document.Sections;
            for (var sectionIndex = 1; sectionIndex <= sections.Count; sectionIndex++)
            {
                Release(textColumn);
                textColumn = null;
                Release(textColumns);
                textColumns = null;
                Release(pageSetup);
                pageSetup = null;
                Release(sectionRange);
                sectionRange = null;
                Release(section);
                section = sections[sectionIndex];
                sectionRange = section.Range;
                pageSetup = section.PageSetup;

                var layout = new MathTypeSectionColumnLayout
                {
                    Start = sectionRange.Start,
                    End = sectionRange.End,
                    LeftMargin = pageSetup.LeftMargin,
                    PageTextWidth = Math.Max(
                        72f,
                        pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin),
                };

                textColumns = pageSetup.TextColumns;
                var columnCount = textColumns?.Count ?? 0;
                if (columnCount <= 1)
                {
                    layout.UniformColumnWidth = layout.PageTextWidth;
                    layout.Widths.Add(layout.PageTextWidth);
                    layout.SpacesAfter.Add(0f);
                    layouts.Add(layout);
                    continue;
                }

                for (var columnIndex = 1; columnIndex <= columnCount; columnIndex++)
                {
                    Release(textColumn);
                    textColumn = null;
                    try { textColumn = textColumns![columnIndex]; }
                    catch { continue; }

                    float width;
                    try { width = textColumn.Width; }
                    catch { continue; }
                    if (!(width > 1f) || float.IsNaN(width) || float.IsInfinity(width))
                        continue;

                    var spaceAfter = 0f;
                    if (columnIndex < columnCount)
                    {
                        try
                        {
                            var candidate = textColumn.SpaceAfter;
                            if (!float.IsNaN(candidate) && !float.IsInfinity(candidate))
                                spaceAfter = Math.Max(0f, candidate);
                        }
                        catch { }
                    }
                    layout.Widths.Add(width);
                    layout.SpacesAfter.Add(spaceAfter);
                }

                if (layout.Widths.Count == 0)
                {
                    layout.UniformColumnWidth = layout.PageTextWidth;
                }
                else if (layout.Widths.Count == 1
                    || layout.Widths.Max() - layout.Widths.Min() <= 0.5f)
                {
                    layout.UniformColumnWidth = layout.Widths.Min();
                }
                layouts.Add(layout);
            }
            return layouts;
        }
        finally
        {
            Release(textColumn);
            Release(textColumns);
            Release(pageSetup);
            Release(sectionRange);
            Release(section);
            Release(sections);
        }
    }

    private static float ResolveMathTypeDisplayColumnWidthFromLayouts(
        Document document,
        Range sourceRange,
        IReadOnlyList<MathTypeSectionColumnLayout> layouts)
    {
        if (TryResolveOwningTableCellContentWidth(sourceRange, out var cellWidth))
            return cellWidth;

        var position = sourceRange.Start;
        MathTypeSectionColumnLayout? layout = null;
        for (var index = 0; index < layouts.Count; index++)
        {
            var candidate = layouts[index];
            if (position < candidate.Start || position >= candidate.End)
                continue;
            layout = candidate;
            break;
        }
        layout ??= layouts.LastOrDefault(candidate => position >= candidate.Start);
        if (layout is null)
            return ResolveMathTypeDisplayColumnWidth(document, sourceRange);
        if (layout.UniformColumnWidth.HasValue)
            return layout.UniformColumnWidth.Value;
        if (layout.Widths.Count == 0)
            return layout.PageTextWidth;

        try
        {
            var pageX = Convert.ToSingle(sourceRange.get_Information(
                WdInformation.wdHorizontalPositionRelativeToPage));
            if (!float.IsNaN(pageX) && !float.IsInfinity(pageX) && pageX >= 0f)
            {
                var relativeX = pageX - layout.LeftMargin;
                var cursor = 0f;
                for (var index = 0; index < layout.Widths.Count; index++)
                {
                    var gap = index < layout.SpacesAfter.Count
                        ? layout.SpacesAfter[index]
                        : 0f;
                    var next = cursor + layout.Widths[index] + gap;
                    if (relativeX < next || index == layout.Widths.Count - 1)
                        return layout.Widths[index];
                    cursor = next;
                }
            }
        }
        catch
        {
            // Unequal-width columns need a live horizontal coordinate. If Word is
            // not paginated yet, the narrowest valid column is the safe fallback.
        }
        return layout.Widths.Min();
    }

    private static bool TryResolveOwningTableCellContentWidth(Range range, out float width)
        => WordFormulaHost.TryGetCellContentWidth(range, out width);

    private static Range ResolveExactLatexSourceRange(
        Document document,
        WordLatexRedrawPlan plan,
        WordLatexRedrawTarget target,
        string expectedSource,
        IReadOnlyList<ResolvedLatexRedrawTarget> alreadyResolved)
    {
        var hasResolvedCoordinates =
            target.AbsoluteStart >= plan.ScopeStart
            && target.AbsoluteEnd > target.AbsoluteStart;
        var approximateStart = hasResolvedCoordinates
            ? target.AbsoluteStart
            : plan.ScopeStart + target.RelativeStart;
        var approximateEnd = hasResolvedCoordinates
            ? target.AbsoluteEnd
            : approximateStart + target.SourceLength;
        Range? direct = null;
        try
        {
            if (approximateStart >= plan.ScopeStart
                && approximateEnd >= approximateStart
                && approximateEnd <= plan.ScopeEnd)
            {
                direct = document.Range(approximateStart, approximateEnd);
                if (string.Equals(
                        direct.Text ?? string.Empty,
                        expectedSource,
                        StringComparison.Ordinal)
                    && !OverlapsResolvedLatexRange(direct, alreadyResolved))
                {
                    target.AbsoluteStart = approximateStart;
                    target.AbsoluteEnd = approximateEnd;
                    var result = direct;
                    direct = null;
                    return result;
                }
            }
        }
        finally { Release(direct); }

        // A table-owned formula must never fall through to a document-wide Word.Find.
        // Word exposes cell/row markers in Range.Text with different UTF-16/story
        // cardinality, and Find over a range spanning table boundaries can force an
        // expensive layout walk or hang Word 2021. Resolve duplicates from the one
        // owning cell's text instead and convert that local UTF-16 offset with the
        // same story-coordinate rules used during capture.
        var tableLocated = TryResolveExactLatexSourceRangeWithinTable(
            document,
            approximateStart,
            expectedSource,
            alreadyResolved);
        if (tableLocated is not null)
        {
            target.AbsoluteStart = tableLocated.Start;
            target.AbsoluteEnd = tableLocated.End;
            return tableLocated;
        }
        if (IsWordPositionWithinTable(document, approximateStart))
            throw new InvalidOperationException(
                $"无法在原表格单元格内重新定位公式：{target.Latex}。为避免跨单元格替换，本次重绘已停止。");

        const int localSearchRadius = 1024;
        var localStart = Math.Max(plan.ScopeStart, approximateStart - localSearchRadius);
        var localEnd = Math.Min(plan.ScopeEnd, approximateEnd + localSearchRadius);
        var located = FindExactLatexSourceRange(
            document,
            localStart,
            localEnd,
            approximateStart,
            expectedSource,
            alreadyResolved);
        if (located is not null)
        {
            target.AbsoluteStart = located.Start;
            target.AbsoluteEnd = located.End;
            return located;
        }

        located = FindExactLatexSourceRange(
            document,
            plan.ScopeStart,
            plan.ScopeEnd,
            approximateStart,
            expectedSource,
            alreadyResolved);
        if (located is not null)
        {
            target.AbsoluteStart = located.Start;
            target.AbsoluteEnd = located.End;
            return located;
        }

        throw new InvalidOperationException(
            $"无法在原位置附近重新定位公式：{target.Latex}。为避免替换错误内容，本次重绘已停止。");
    }

    private static Range? TryResolveExactLatexSourceRangeWithinTable(
        Document document,
        int approximateStart,
        string expectedSource,
        IReadOnlyList<ResolvedLatexRedrawTarget> alreadyResolved)
    {
        if (string.IsNullOrEmpty(expectedSource)) return null;
        Range? content = null;
        Range? probe = null;
        Cells? cells = null;
        Cell? cell = null;
        Range? cellRange = null;
        Range? candidate = null;
        Range? best = null;
        try
        {
            content = document.Content;
            if (approximateStart < content.Start || approximateStart >= content.End)
                return null;
            probe = document.Range(
                approximateStart,
                Math.Min(content.End, approximateStart + 1));
            cells = probe.Cells;
            if (cells.Count != 1) return null;
            cell = cells[1];
            cellRange = cell.Range.Duplicate;
            var cellText = cellRange.Text ?? string.Empty;
            var localStoryOffsets = BuildWordStoryOffsetIndex(cellText);
            var searchIndex = 0;
            var bestDistance = int.MaxValue;
            while (searchIndex <= cellText.Length - expectedSource.Length)
            {
                var match = cellText.IndexOf(
                    expectedSource,
                    searchIndex,
                    StringComparison.Ordinal);
                if (match < 0) break;
                var matchEnd = match + expectedSource.Length;
                var candidateStart = cellRange.Start + localStoryOffsets[match];
                var candidateEnd = cellRange.Start + localStoryOffsets[matchEnd];
                if (candidateEnd > candidateStart && candidateEnd <= cellRange.End)
                {
                    Release(candidate);
                    candidate = document.Range(candidateStart, candidateEnd);
                    if (string.Equals(
                            candidate.Text ?? string.Empty,
                            expectedSource,
                            StringComparison.Ordinal)
                        && !OverlapsResolvedLatexRange(candidate, alreadyResolved))
                    {
                        var distance = Math.Abs(candidateStart - approximateStart);
                        if (distance < bestDistance)
                        {
                            Release(best);
                            best = candidate.Duplicate;
                            bestDistance = distance;
                        }
                    }
                }
                searchIndex = match + 1;
            }
            var result = best;
            best = null;
            return result;
        }
        catch (COMException)
        {
            return null;
        }
        finally
        {
            Release(best);
            Release(candidate);
            Release(cellRange);
            Release(cell);
            Release(cells);
            Release(probe);
            Release(content);
        }
    }

    private static bool IsWordPositionWithinTable(Document document, int position)
    {
        Range? content = null;
        Range? probe = null;
        Tables? tables = null;
        try
        {
            content = document.Content;
            if (position < content.Start || position >= content.End) return false;
            probe = document.Range(position, Math.Min(content.End, position + 1));
            tables = probe.Tables;
            return tables.Count > 0;
        }
        catch (COMException)
        {
            return false;
        }
        finally
        {
            Release(tables);
            Release(probe);
            Release(content);
        }
    }

    private static Range? FindExactLatexSourceRange(
        Document document,
        int searchStart,
        int searchEnd,
        int approximateStart,
        string expectedSource,
        IReadOnlyList<ResolvedLatexRedrawTarget> alreadyResolved)
    {
        if (searchEnd <= searchStart || expectedSource.Length == 0) return null;
        var findText = BuildWordFindAnchor(expectedSource);
        if (findText.Length == 0) return null;

        Range? search = null;
        Range? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            search = document.Range(searchStart, searchEnd);
            while (search.Start < searchEnd)
            {
                Find? find = null;
                var matched = false;
                try
                {
                    find = search.Find;
                    find.ClearFormatting();
                    find.Text = findText;
                    find.Forward = true;
                    find.Wrap = WdFindWrap.wdFindStop;
                    find.Format = false;
                    find.MatchCase = true;
                    find.MatchWholeWord = false;
                    find.MatchWildcards = false;
                    find.MatchSoundsLike = false;
                    find.MatchAllWordForms = false;
                    matched = find.Execute();
                }
                finally { Release(find); }
                if (!matched) break;

                var matchStart = search.Start;
                var nextSearchStart = Math.Min(
                    searchEnd,
                    Math.Max(matchStart + 1, search.End));
                var candidate = TryCreateExactLatexRangeAt(
                    document,
                    matchStart,
                    searchEnd,
                    expectedSource);
                if (candidate is not null)
                {
                    if (!OverlapsResolvedLatexRange(candidate, alreadyResolved))
                    {
                        var distance = Math.Abs(candidate.Start - approximateStart);
                        if (distance < bestDistance)
                        {
                            Release(best);
                            best = candidate;
                            candidate = null;
                            bestDistance = distance;
                        }
                    }
                    Release(candidate);
                }
                search.SetRange(nextSearchStart, searchEnd);
            }

            var result = best;
            best = null;
            return result;
        }
        finally
        {
            Release(best);
            Release(search);
        }
    }

    private static Range? TryCreateExactLatexRangeAt(
        Document document,
        int start,
        int maximumEnd,
        string expectedSource)
    {
        const int maximumCoordinateAdjustment = 256;
        for (var adjustment = 0;
             adjustment <= maximumCoordinateAdjustment;
             adjustment++)
        {
            var deltas = adjustment == 0
                ? new[] { 0 }
                : new[] { -adjustment, adjustment };
            foreach (var delta in deltas)
            {
                var end = start + expectedSource.Length + delta;
                if (end <= start || end > maximumEnd) continue;
                Range? candidate = null;
                try
                {
                    candidate = document.Range(start, end);
                    if (string.Equals(
                            candidate.Text ?? string.Empty,
                            expectedSource,
                            StringComparison.Ordinal))
                    {
                        var result = candidate;
                        candidate = null;
                        return result;
                    }
                }
                catch (COMException)
                {
                    // Keep trying nearby Word story coordinates.
                }
                finally { Release(candidate); }
            }
        }
        return null;
    }

    private static bool OverlapsResolvedLatexRange(
        Range candidate,
        IReadOnlyList<ResolvedLatexRedrawTarget> alreadyResolved)
    {
        // Read the COM coordinates once. The previous implementation read
        // Range.Start/End again for every earlier formula, which turns a
        // 1000-formula document into hundreds of thousands of cross-process COM
        // calls even though the overlap comparison itself is trivial.
        var candidateStart = candidate.Start;
        var candidateEnd = candidate.End;
        foreach (var resolved in alreadyResolved)
        {
            if (candidateStart < resolved.SourceEnd
                && resolved.SourceStart < candidateEnd)
                return true;
        }
        return false;
    }

    private static string BuildWordFindAnchor(string source)
    {
        const int maximumFindTextLength = 180;
        var builder = new StringBuilder(Math.Min(source.Length, maximumFindTextLength));
        foreach (var character in source)
        {
            var token = character switch
            {
                '^' => "^^",
                '\r' => "^p",
                '\v' => "^l",
                '\t' => "^t",
                _ => character.ToString(),
            };
            if (builder.Length > 0
                && builder.Length + token.Length > maximumFindTextLength)
                break;
            builder.Append(token);
        }
        return builder.ToString();
    }

    private static double?[]? TryBuildWordOpenXmlFontSizeIndex(
        Range scope,
        string sourceText)
    {
        const string packageNamespace =
            "http://schemas.microsoft.com/office/2006/xmlPackage";
        const string wordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        try
        {
            var package = XDocument.Parse(
                scope.WordOpenXML,
                LoadOptions.PreserveWhitespace);
            XNamespace packageNs = packageNamespace;
            XNamespace word = wordNamespace;
            var documentPart = package
                .Descendants(packageNs + "part")
                .FirstOrDefault(part => string.Equals(
                    (string?)part.Attribute(packageNs + "name"),
                    "/word/document.xml",
                    StringComparison.OrdinalIgnoreCase));
            var body = documentPart?
                .Descendants(word + "body")
                .FirstOrDefault();
            if (body is null) return null;

            var reconstructed = new StringBuilder(sourceText.Length + 16);
            var fontSizes = new List<double?>(sourceText.Length + 16);
            foreach (var paragraph in body.Descendants(word + "p"))
            {
                var paragraphFontSize = ReadWordOpenXmlFontSize(
                    paragraph.Element(word + "pPr")?.Element(word + "rPr"),
                    word);
                foreach (var token in paragraph.Descendants())
                {
                    string? tokenText = null;
                    if (token.Name == word + "t"
                        || token.Name == word + "delText")
                    {
                        tokenText = token.Value;
                    }
                    else if (token.Name == word + "tab")
                    {
                        tokenText = "\t";
                    }
                    else if (token.Name == word + "br"
                        || token.Name == word + "cr")
                    {
                        tokenText = "\v";
                    }
                    else if (token.Name == word + "noBreakHyphen")
                    {
                        tokenText = "\u2011";
                    }
                    else if (token.Name == word + "softHyphen")
                    {
                        tokenText = "\u00AD";
                    }
                    if (tokenText is null) continue;

                    var run = token.Ancestors(word + "r").FirstOrDefault();
                    var runFontSize = ReadWordOpenXmlFontSize(
                        run?.Element(word + "rPr"),
                        word)
                        ?? paragraphFontSize;
                    reconstructed.Append(tokenText);
                    for (var index = 0; index < tokenText.Length; index++)
                        fontSizes.Add(runFontSize);
                }
                reconstructed.Append('\r');
                fontSizes.Add(paragraphFontSize);
            }

            var reconstructedText = reconstructed.ToString();
            if (string.Equals(
                    reconstructedText,
                    sourceText,
                    StringComparison.Ordinal))
                return fontSizes.ToArray();

            // A Word Open XML fragment can include one mandatory trailing empty
            // paragraph that is outside the requested story range. Accept only that
            // exact, unambiguous difference; every other mismatch falls back to COM.
            if (reconstructedText.Length == sourceText.Length + 1
                && reconstructedText[reconstructedText.Length - 1] == '\r'
                && string.Equals(
                    reconstructedText.Substring(0, sourceText.Length),
                    sourceText,
                    StringComparison.Ordinal))
            {
                fontSizes.RemoveAt(fontSizes.Count - 1);
                return fontSizes.ToArray();
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    private static double? ReadWordOpenXmlFontSize(
        XElement? runProperties,
        XNamespace word)
    {
        if (runProperties is null) return null;
        var value = (string?)runProperties
            .Element(word + "sz")?
            .Attribute(word + "val")
            ?? (string?)runProperties
                .Element(word + "szCs")?
                .Attribute(word + "val");
        if (!double.TryParse(
                value,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out var halfPoints)
            || halfPoints <= 0
            || double.IsNaN(halfPoints)
            || double.IsInfinity(halfPoints))
            return null;
        return halfPoints / 2d;
    }

    private static IReadOnlyDictionary<string, LatexRedrawSourceContext>
        BuildLatexRedrawSourceContexts(
            string sourceText,
            IReadOnlyList<WordLatexRedrawTarget> targets)
    {
        var contexts = new Dictionary<string, LatexRedrawSourceContext>(
            targets.Count,
            StringComparer.Ordinal);
        var formulaMask = new bool[sourceText.Length];
        foreach (var target in targets)
        {
            var start = Math.Max(0, Math.Min(sourceText.Length, target.RelativeStart));
            var end = Math.Max(start, Math.Min(
                sourceText.Length,
                target.RelativeStart + target.SourceLength));
            for (var index = start; index < end; index++)
                formulaMask[index] = true;
        }

        foreach (var target in targets)
        {
            var start = Math.Max(0, Math.Min(sourceText.Length, target.RelativeStart));
            var end = Math.Max(start, Math.Min(
                sourceText.Length,
                target.RelativeStart + target.SourceLength));
            var cellStart = FindSourceCellStart(sourceText, start);
            var cellEnd = FindSourceCellEnd(sourceText, end);
            var paragraphStart = Math.Max(
                cellStart,
                FindSourceParagraphStart(sourceText, start));
            var paragraphEnd = Math.Min(
                cellEnd,
                FindSourceParagraphEnd(sourceText, end));
            var previousInParagraph = FindPreviousVisibleSourcePosition(
                sourceText,
                start - 1,
                paragraphStart,
                formulaMask);
            var nextInParagraph = FindNextVisibleSourcePosition(
                sourceText,
                end,
                paragraphEnd,
                formulaMask);
            var hasVisibleSurroundingText =
                previousInParagraph >= 0 || nextInParagraph >= 0;
            var display = string.Equals(
                target.DisplayMode,
                "block",
                StringComparison.Ordinal);

            int fontContextPosition;
            if (!display || hasVisibleSurroundingText)
            {
                fontContextPosition = previousInParagraph >= 0
                    ? previousInParagraph
                    : nextInParagraph;
            }
            else
            {
                // A display formula that is the only visible content in a Word
                // table cell must inherit from that cell, not from text in a
                // neighboring cell. Word's story text concatenates cells with \a,
                // so an unrestricted previous/next search silently crosses table
                // geometry and can pick an unrelated font/size.
                var previousParagraph = FindPreviousVisibleSourcePosition(
                    sourceText,
                    paragraphStart - 1,
                    cellStart,
                    formulaMask);
                var nextParagraph = FindNextVisibleSourcePosition(
                    sourceText,
                    paragraphEnd,
                    cellEnd,
                    formulaMask);
                fontContextPosition = previousParagraph >= 0
                    ? previousParagraph
                    : nextParagraph;
            }

            contexts[target.Id] = new LatexRedrawSourceContext
            {
                HasVisibleSurroundingText = hasVisibleSurroundingText,
                FontContextRelativePosition = fontContextPosition,
            };
        }
        return contexts;
    }

    private static int FindSourceCellStart(string sourceText, int position)
    {
        for (var index = Math.Min(position, sourceText.Length) - 1;
             index >= 0;
             index--)
        {
            if (sourceText[index] == '\a') return index + 1;
        }
        return 0;
    }

    private static int FindSourceCellEnd(string sourceText, int position)
    {
        for (var index = Math.Max(0, position);
             index < sourceText.Length;
             index++)
        {
            if (sourceText[index] == '\a') return index;
        }
        return sourceText.Length;
    }

    private static int FindSourceParagraphStart(string sourceText, int position)
    {
        for (var index = Math.Min(position, sourceText.Length) - 1;
             index >= 0;
             index--)
        {
            if (IsSourceParagraphBoundary(sourceText[index]))
                return index + 1;
        }
        return 0;
    }

    private static int FindSourceParagraphEnd(string sourceText, int position)
    {
        for (var index = Math.Max(0, position);
             index < sourceText.Length;
             index++)
        {
            if (IsSourceParagraphBoundary(sourceText[index]))
                return index;
        }
        return sourceText.Length;
    }

    private static int FindPreviousVisibleSourcePosition(
        string sourceText,
        int startPosition,
        int minimumPosition,
        IReadOnlyList<bool> formulaMask)
    {
        for (var index = Math.Min(startPosition, sourceText.Length - 1);
             index >= Math.Max(0, minimumPosition);
             index--)
        {
            if (char.IsLowSurrogate(sourceText[index])
                && index > minimumPosition
                && char.IsHighSurrogate(sourceText[index - 1]))
                index--;
            if (formulaMask[index]) continue;
            if (char.IsHighSurrogate(sourceText[index])
                && index + 1 < sourceText.Length
                && char.IsLowSurrogate(sourceText[index + 1])
                && formulaMask[index + 1])
                continue;
            if (IsVisibleSourceCharacter(sourceText[index])) return index;
        }
        return -1;
    }

    private static int FindNextVisibleSourcePosition(
        string sourceText,
        int startPosition,
        int maximumPosition,
        IReadOnlyList<bool> formulaMask)
    {
        var maximum = Math.Min(sourceText.Length, maximumPosition);
        for (var index = Math.Max(0, startPosition);
             index < maximum;
             index++)
        {
            if (char.IsLowSurrogate(sourceText[index])
                && index > 0
                && char.IsHighSurrogate(sourceText[index - 1]))
                continue;
            if (formulaMask[index]) continue;
            if (char.IsHighSurrogate(sourceText[index])
                && index + 1 < sourceText.Length
                && char.IsLowSurrogate(sourceText[index + 1])
                && formulaMask[index + 1])
                continue;
            if (IsVisibleSourceCharacter(sourceText[index])) return index;
            if (char.IsHighSurrogate(sourceText[index])
                && index + 1 < maximum
                && char.IsLowSurrogate(sourceText[index + 1]))
                index++;
        }
        return -1;
    }

    private static bool IsSourceParagraphBoundary(char character) =>
        character is '\r' or '\n' or '\v' or '\a';

    private static bool IsVisibleSourceCharacter(char character)
    {
        if (character is '\r' or '\n' or '\t' or '\v' or '\a'
            or '\u0001' or '\u200B' or '\u200C')
            return false;
        return !char.IsWhiteSpace(character);
    }

    private static bool TryResolveSourceFormulaFontSizeFromContext(
        Document document,
        WordLatexRedrawPlan plan,
        WordLatexRedrawTarget target,
        int sourceStart,
        int sourceEnd,
        int contextRelativePosition,
        out float fontSizePt)
    {
        fontSizePt = FormulaFontSize.DefaultPt;
        if (contextRelativePosition < 0
            || contextRelativePosition >= plan.SourceText.Length)
            return false;

        var targetEnd = target.RelativeStart + target.SourceLength;
        int wordPosition;
        if (contextRelativePosition < target.RelativeStart)
        {
            wordPosition = sourceStart - CountWordStoryCharacters(
                plan.SourceText,
                contextRelativePosition,
                target.RelativeStart);
        }
        else if (contextRelativePosition >= targetEnd)
        {
            wordPosition = sourceEnd + CountWordStoryCharacters(
                plan.SourceText,
                targetEnd,
                contextRelativePosition);
        }
        else
        {
            return false;
        }

        if (wordPosition < plan.ScopeStart || wordPosition >= plan.ScopeEnd)
            return false;

        Range? probe = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            probe = document.Range(wordPosition, wordPosition + 1);
            var expectedText = ReadSourceCharacterAt(
                plan.SourceText,
                contextRelativePosition);
            var probeText = probe.Text ?? string.Empty;
            if (!string.Equals(
                    probeText,
                    expectedText,
                    StringComparison.Ordinal)
                || !ContainsVisibleBodyText(probeText))
                return false;
            font = probe.Font;
            return TryResolveWordFontSize(font.Size, out fontSizePt);
        }
        catch (COMException)
        {
            return false;
        }
        finally
        {
            Release(font);
            Release(probe);
        }
    }

    private static int[] BuildWordStoryOffsetIndex(string sourceText)
    {
        var offsets = new int[sourceText.Length + 1];
        var sourceIndex = 0;
        var wordOffset = 0;
        while (sourceIndex < sourceText.Length)
        {
            offsets[sourceIndex] = wordOffset;
            if (char.IsHighSurrogate(sourceText[sourceIndex])
                && sourceIndex + 1 < sourceText.Length
                && char.IsLowSurrogate(sourceText[sourceIndex + 1]))
            {
                // Word counts a valid UTF-16 surrogate pair as one story
                // character. The intermediate boundary is not a valid formula
                // boundary, but mapping it to the pair start keeps the table total.
                offsets[sourceIndex + 1] = wordOffset;
                sourceIndex += 2;
                wordOffset++;
                offsets[sourceIndex] = wordOffset;
                continue;
            }
            if (sourceText[sourceIndex] == '\r'
                && sourceIndex + 1 < sourceText.Length
                && sourceText[sourceIndex + 1] == '\a')
            {
                // Range.Text serializes every Word table cell/row terminator as
                // CR + BEL (\r\a), but Word's story coordinates count that pair as
                // one structural character. Treating both UTF-16 characters as two
                // story positions shifts every formula after the first table cell
                // and eventually sends relocation into a cross-table Word.Find.
                offsets[sourceIndex + 1] = wordOffset;
                sourceIndex += 2;
                wordOffset++;
                offsets[sourceIndex] = wordOffset;
                continue;
            }
            sourceIndex++;
            wordOffset++;
            offsets[sourceIndex] = wordOffset;
        }
        return offsets;
    }

    private static int CountWordStoryCharacters(
        string sourceText,
        int startPosition,
        int endPosition)
    {
        var start = Math.Max(0, Math.Min(sourceText.Length, startPosition));
        var end = Math.Max(start, Math.Min(sourceText.Length, endPosition));
        var count = 0;
        for (var index = start; index < end; index++, count++)
        {
            if (char.IsHighSurrogate(sourceText[index])
                && index + 1 < end
                && char.IsLowSurrogate(sourceText[index + 1]))
            {
                index++;
                continue;
            }
            if (sourceText[index] == '\r'
                && index + 1 < end
                && sourceText[index + 1] == '\a')
                index++;
        }
        return count;
    }

    private static string ReadSourceCharacterAt(string sourceText, int position)
    {
        if (position < 0 || position >= sourceText.Length) return string.Empty;
        if (char.IsHighSurrogate(sourceText[position])
            && position + 1 < sourceText.Length
            && char.IsLowSurrogate(sourceText[position + 1]))
            return sourceText.Substring(position, 2);
        return sourceText[position].ToString();
    }

    private static float ResolveSourceFormulaFontSize(
        Document document,
        Range source,
        bool display)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            paragraphs = source.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
            }

            var paragraphStart = paragraphRange?.Start ?? source.Start;
            var paragraphBodyEnd = Math.Max(
                paragraphStart,
                (paragraphRange?.End ?? source.End) - 1);

            // Inline formulas should inherit the prose beside them. Generated
            // documents often deliberately make the raw $...$ source smaller than
            // the surrounding text, so the source run itself is only a fallback.
            if (!display || HasVisibleSurroundingText(source))
            {
                if (TryResolveNearbyVisibleFontSize(
                        document,
                        source.Start - 1,
                        out var previousInline,
                        minimumPosition: paragraphStart,
                        step: -1))
                    return previousInline;
                if (TryResolveNearbyVisibleFontSize(
                        document,
                        source.End,
                        out var nextInline,
                        maximumPosition: paragraphBodyEnd,
                        step: 1))
                    return nextInline;
            }

            // A display formula normally occupies its own paragraph. In that case
            // inherit from the nearest visible prose outside the source paragraph,
            // preferring the preceding paragraph as Word users generally expect.
            if (TryResolveNearbyVisibleFontSize(
                    document,
                    paragraphStart - 1,
                    out var previousParagraph,
                    minimumPosition: 0,
                    step: -1))
                return previousParagraph;
            if (TryResolveNearbyVisibleFontSize(
                    document,
                    paragraphRange?.End ?? source.End,
                    out var nextParagraph,
                    maximumPosition: int.MaxValue,
                    step: 1))
                return nextParagraph;

            font = source.Font;
            if (TryResolveWordFontSize(font.Size, out var direct)) return direct;
            Release(font);
            font = null;

            if (paragraphRange is not null)
            {
                if (paragraphRange.End > paragraphRange.Start)
                    paragraphRange.End -= 1;
                font = paragraphRange.Font;
                if (TryResolveWordFontSize(font.Size, out var paragraphSize))
                    return paragraphSize;
            }
            return FormulaFontSize.DefaultPt;
        }
        finally
        {
            Release(font);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool TryResolveNearbyVisibleFontSize(
        Document document,
        int startPosition,
        out float fontSizePt,
        int minimumPosition = 0,
        int maximumPosition = int.MaxValue,
        int step = 1)
    {
        fontSizePt = FormulaFontSize.DefaultPt;
        if (step is not (-1 or 1)) return false;

        Range? content = null;
        try
        {
            content = document.Content;
            var contentStart = content.Start;
            var contentEnd = Math.Max(contentStart, content.End - 1);
            var lowerBound = Math.Max(contentStart, minimumPosition);
            var upperBound = Math.Min(contentEnd, maximumPosition);
            if (upperBound < lowerBound
                || (step < 0 && startPosition < lowerBound)
                || (step > 0 && startPosition > upperBound))
                return false;
            var position = startPosition;
            const int maximumProbeCharacters = 256;
            for (var probeIndex = 0;
                 probeIndex < maximumProbeCharacters
                 && position >= lowerBound
                 && position <= upperBound;
                 probeIndex++, position += step)
            {
                Range? probe = null;
                Microsoft.Office.Interop.Word.Font? font = null;
                try
                {
                    probe = document.Range(position, Math.Min(position + 1, content.End));
                    if (!ContainsVisibleBodyText(probe.Text)) continue;
                    font = probe.Font;
                    if (TryResolveWordFontSize(font.Size, out fontSizePt))
                        return true;
                }
                catch (COMException)
                {
                    // Keep probing neighboring Word story coordinates.
                }
                finally
                {
                    Release(font);
                    Release(probe);
                }
            }
            return false;
        }
        finally { Release(content); }
    }

    public WordBulkInsertResult InsertBulkDocument(
        WordBulkImportDocument source,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        string? expectedDocumentId,
        string? sourceObjectId)
    {
        var mathTypeOnly = prepared.Count > 0
            && prepared.Values.All(item => string.Equals(
                item.Session.ObjectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal));

        if (!mathTypeOnly)
            return InsertBulkDocumentHostCore(
                source,
                prepared,
                expectedDocumentId,
                sourceObjectId);

        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, expectedDocumentId);
            return ExecuteDocumentEdit(document, "VisualTeX 批量导入 MathType", () =>
                InsertBulkDocumentCore(source, prepared, expectedDocumentId, sourceObjectId));
        }
        finally { Release(document); }
    }

    private WordBulkInsertResult InsertBulkDocumentCore(
        WordBulkImportDocument source,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        string? expectedDocumentId,
        string? sourceObjectId)
    {
        if (source is null)
            throw new ArgumentNullException(nameof(source));
        if (prepared is null)
            throw new ArgumentNullException(nameof(prepared));
        if (prepared.Values.Any(item =>
                !string.Equals(
                    item.Session.ObjectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)))
            throw new InvalidOperationException(
                "The legacy bulk insertion core is MathType-only.");

        Document? document = null;
        Selection? selection = null;
        var insertedFormulaIds =
            new List<string>();
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(
                document,
                expectedDocumentId);

            selection = _application.Selection;
            Range? sourceRange = null;
            try
            {
                sourceRange =
                    ResolveSourceRange(
                        document,
                        sourceObjectId,
                        selection);
                selection.SetRange(
                    sourceRange.Start,
                    sourceRange.End);
            }
            finally
            {
                Release(sourceRange);
            }

            if (selection.Range.Start
                != selection.Range.End)
                selection.Text =
                    string.Empty;
            selection.Collapse(
                WdCollapseDirection.wdCollapseEnd);

            foreach (var preparedFormula
                     in prepared.Values)
            {
                preparedFormula.Session.Numbered =
                    source.NumberDisplayFormulas
                    && string.Equals(
                        preparedFormula.Run.DisplayMode,
                        "block",
                        StringComparison.Ordinal);
            }

            for (var blockIndex = 0;
                 blockIndex < source.Blocks.Count;
                 blockIndex++)
            {
                var block =
                    source.Blocks[blockIndex];
                var nextKind =
                    blockIndex + 1
                        < source.Blocks.Count
                        ? source.Blocks[
                            blockIndex + 1].Kind
                        : (WordBulkBlockKind?)null;

                if (block.Kind ==
                    WordBulkBlockKind.DisplayFormula)
                {
                    var formulaRun =
                        block.Runs.Single(
                            run => run.IsFormula);
                    if (!prepared.TryGetValue(
                            formulaRun.Id,
                            out var formula))
                        throw new InvalidDataException(
                            $"缺少行间公式 {formulaRun.Id} 的渲染结果。");

                    InsertPreparedFormula(
                        document,
                        selection,
                        formula,
                        display: true,
                        ommlBatchSource: null,
                        deferredOmmlMetadata: null,
                        bulkImport: true);
                    insertedFormulaIds.Add(
                        formula.Session.FormulaId);
                    continue;
                }

                EnsureWritableParagraph(
                    selection);
                var paragraphStart =
                    selection.Start;
                var pendingInlineFormulas =
                    new List<(
                        int Start,
                        PreparedWordBulkFormula Formula)>();

                foreach (var run in block.Runs)
                {
                    if (!run.IsFormula)
                    {
                        InsertNativeTextRun(
                            document,
                            selection,
                            run);
                        continue;
                    }

                    if (!prepared.TryGetValue(
                            run.Id,
                            out var formula))
                        throw new InvalidDataException(
                            $"缺少行内公式 {run.Id} 的渲染结果。");

                    var placeholderStart =
                        selection.Start;
                    selection.TypeText(
                        BulkInlineFormulaPlaceholder);
                    pendingInlineFormulas.Add(
                        (placeholderStart, formula));
                }

                selection.TypeParagraph();
                var paragraphEnd =
                    selection.Start;
                ApplyBulkParagraphFormatting(
                    document,
                    paragraphStart,
                    paragraphEnd,
                    block);

                for (var formulaIndex =
                         pendingInlineFormulas.Count - 1;
                     formulaIndex >= 0;
                     formulaIndex--)
                {
                    var pending =
                        pendingInlineFormulas[
                            formulaIndex];
                    selection.SetRange(
                        pending.Start,
                        pending.Start
                        + BulkInlineFormulaPlaceholder.Length);
                    selection.Text =
                        string.Empty;
                    selection.Collapse(
                        WdCollapseDirection.wdCollapseStart);
                    InsertPreparedFormula(
                        document,
                        selection,
                        pending.Formula,
                        display: false,
                        ommlBatchSource: null,
                        deferredOmmlMetadata: null,
                        bulkImport: true);
                    insertedFormulaIds.Add(
                        pending.Formula.Session.FormulaId);
                }

                MoveSelectionAfterBulkParagraph(
                    document,
                    selection,
                    paragraphStart);
                ResetNextParagraphFormatting(
                    selection,
                    block.Kind,
                    nextKind);
            }

            if (source.NumberDisplayFormulas
                && source.DisplayFormulaCount > 0)
            {
                MathTypeEquationNumbering
                    .UpdateEquationNumbers(
                        document);
                WordEquationReferenceFields
                    .UpdateReferences(
                        document);
            }

            return new WordBulkInsertResult
            {
                BlockCount = source.Blocks.Count,
                FormulaCount =
                    insertedFormulaIds.Count,
                FormulaIds =
                    insertedFormulaIds,
            };
        }
        finally
        {
            Release(selection);
            Release(document);
        }
    }

    private static void ApplyBulkOmmlTypographyXml(
        XElement equation,
        double fontSizePt,
        FormulaMetadata metadata,
        XNamespace word,
        XNamespace math)
    {
        var halfPoints = ((int)Math.Round(
            FormulaFontSize.NormalizeWordOmmlSize(fontSizePt) * 2.0,
            MidpointRounding.AwayFromZero))
            .ToString(System.Globalization.CultureInfo.InvariantCulture);

        // Keep ordinary variables, digits, Greek letters, operators and all native
        // structures as genuine Office Math. Their glyph selection and spacing must
        // come from the document-level m:mathFont and its OpenType MATH table—not
        // from per-run w:rFonts, run splitting, or m:nor normal-text conversion.
        foreach (var mathRun in equation.DescendantsAndSelf(math + "r"))
        {
            var mathProperties = mathRun.Element(math + "rPr");
            var normalText = mathProperties?.Element(math + "nor") is not null;
            var properties = mathRun.Element(word + "rPr");
            if (properties is null)
            {
                properties = new XElement(word + "rPr");
                if (mathProperties is not null) mathProperties.AddAfterSelf(properties);
                else mathRun.AddFirst(properties);
            }

            if (!normalText)
            {
                // A native math run must inherit the document's Office Math font.
                // Explicit run fonts bypass that mechanism and can flatten MATH-table
                // italic correction, operator spacing and extensible constructions.
                properties.Element(word + "rFonts")?.Remove();
            }
            else if (ContainsChineseOmmlText(mathRun, math))
            {
                // m:nor is retained only when it already carries real text semantics
                // (for example MathML mtext / LaTeX \text{中文}). Select only the
                // East-Asian text face; do not override the mathematical ASCII runs.
                var fonts = properties.Element(word + "rFonts");
                if (fonts is null)
                {
                    fonts = new XElement(word + "rFonts");
                    properties.AddFirst(fonts);
                }
                fonts.SetAttributeValue(word + "eastAsiaTheme", null);
                fonts.SetAttributeValue(
                    word + "eastAsia",
                    ResolveOmmlChineseFont(metadata.FormulaChineseFont));
            }

            ApplyNativeOmmlSizeAndPosition(properties, halfPoints, word);
        }

        // Fraction bars, radicals, delimiters and other structures may carry
        // m:ctrlPr/w:rPr independently of visible runs. Size them consistently,
        // while removing any control-level font override so the same MATH font
        // drives glyph assembly and spacing throughout the complete OMath tree.
        foreach (var controlProperties in equation.DescendantsAndSelf(math + "ctrlPr"))
        {
            var properties = controlProperties.Element(word + "rPr");
            if (properties is null)
            {
                properties = new XElement(word + "rPr");
                controlProperties.Add(properties);
            }
            properties.Element(word + "rFonts")?.Remove();
            ApplyNativeOmmlSizeAndPosition(properties, halfPoints, word);
        }

        // Genuine Word display OMML obtains fraction, radical, matrix and large-
        // operator sizing from the OpenType MATH table. Never apply the retired
        // numbered-inline 1.5x numerator/denominator compensation here.

        // Numbering is paragraph structure, never mathematical content. Keep the
        // OMath tree free of generated '#(...)' wrappers so Word does not turn a
        // normal formula into a full-width equation-array control. The visible REF
        // number is created outside OMath by WordEquationNumbering, exactly like the
        // accepted VisualTeX/MathType OLE tab layout.
    }

    private static bool ContainsChineseOmmlText(
        XElement mathRun,
        XNamespace math)
    {
        var text = string.Concat(mathRun.Elements(math + "t").Select(item => item.Value));
        for (var index = 0; index < text.Length;)
        {
            var width = char.IsHighSurrogate(text[index])
                && index + 1 < text.Length
                && char.IsLowSurrogate(text[index + 1])
                    ? 2
                    : 1;
            var codePoint = width == 2
                ? char.ConvertToUtf32(text[index], text[index + 1])
                : text[index];
            if (IsChineseOmmlCodePoint(codePoint)) return true;
            index += width;
        }
        return false;
    }

    private static void ApplyNativeOmmlSizeAndPosition(
        XElement properties,
        string halfPoints,
        XNamespace word)
    {
        var size = properties.Element(word + "sz");
        if (size is null)
        {
            size = new XElement(word + "sz");
            properties.Add(size);
        }
        size.SetAttributeValue(word + "val", halfPoints);

        var complexSize = properties.Element(word + "szCs");
        if (complexSize is null)
        {
            complexSize = new XElement(word + "szCs");
            properties.Add(complexSize);
        }
        complexSize.SetAttributeValue(word + "val", halfPoints);

        var position = properties.Element(word + "position");
        if (position is null)
        {
            position = new XElement(word + "position");
            properties.Add(position);
        }
        position.SetAttributeValue(word + "val", "0");
    }

    private void InsertPreparedFormula(
        Document document,
        Selection selection,
        PreparedWordBulkFormula prepared,
        bool display,
        bool preserveExistingDisplayParagraphBoundary = false,
        Range? preservedDisplayParagraphRange = null,
        string? preservedFollowingParagraphText = null,
        WordOmmlConverter.BatchSource? ommlBatchSource = null,
        ICollection<FormulaMetadata>? deferredOmmlMetadata = null,
        Action<Range>? retainInsertedOmml = null,
        bool deferOmmlBatchFinalization = false,
        Action<Range>? retainInsertedNativeOleIdentity = null,
        bool deferNativeOleBatchFinalization = false,
        bool bulkImport = false,
        float? mathTypeDisplayColumnWidth = null,
        Action<Range, string>? retainMathTypeForCompletedValidation = null)
    {
        _ = selection;
        _ = ommlBatchSource;
        _ = deferredOmmlMetadata;
        _ = retainInsertedOmml;
        _ = deferOmmlBatchFinalization;
        _ = retainInsertedNativeOleIdentity;
        _ = deferNativeOleBatchFinalization;
        _ = bulkImport;

        var session = prepared.Session;
        if (!string.Equals(
                session.ObjectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal))
            throw new InvalidOperationException(
                "The legacy prepared-formula writer is MathType-only.");

        session.DisplayMode =
            display ? "block" : "inline";
        session.Numbered =
            display && session.Numbered;

        var metadata =
            session.ToMetadata();
        metadata.Validate();

        if (string.IsNullOrWhiteSpace(
                prepared.MathMl))
            throw new InvalidDataException(
                $"公式 {metadata.FormulaId} 没有可用于 MathType 的 MathML。");
        if (string.IsNullOrWhiteSpace(
                prepared.EmfPath))
            throw new InvalidDataException(
                $"公式 {metadata.FormulaId} 没有可用于 MathType 的矢量预览。");

        var nativePreview =
            prepared.MathTypeNativePreview;
        var preservedLayout =
            display
            && preserveExistingDisplayParagraphBoundary
            && preservedDisplayParagraphRange is not null
                ? CaptureMathTypeDisplayParagraphLayout(
                    preservedDisplayParagraphRange)
                : null;

        InsertMathTypeOle(
            session,
            prepared.MathMl!,
            prepared.EmfPath,
            isolatedNativePreviewWmfPath:
                nativePreview?.WmfPath,
            isolatedNativePreviewWidthPt:
                nativePreview?.WidthPt ?? 0,
            isolatedNativePreviewHeightPt:
                nativePreview?.HeightPt ?? 0,
            isolatedNativePreviewWordPosition:
                nativePreview?.WordPosition ?? 0,
            isolatedNativePreviewAttempted:
                prepared.MathTypeNativePreviewAttempted,
            preserveExistingDisplayParagraphBoundary:
                preserveExistingDisplayParagraphBoundary,
            preservedDisplayParagraphLayout:
                preservedLayout,
            knownDisplayColumnWidth:
                mathTypeDisplayColumnWidth,
            retainMathTypeForCompletedValidation:
                retainMathTypeForCompletedValidation,
            preserveCapturedInsertion:
                true);

        if (display
            && preserveExistingDisplayParagraphBoundary
            && preservedDisplayParagraphRange is not null)
        {
            RemoveGeneratedParagraphAfterMathTypeRedraw(
                document,
                preservedDisplayParagraphRange,
                preservedFollowingParagraphText);
        }
    }

    private static void MoveSelectionAfterBulkParagraph(
        Document document,
        Selection selection,
        int paragraphStart)
    {
        var paragraph = WordCharacterFormatting.ResolveMainStoryParagraphAtPosition(document, paragraphStart);
        try { selection.SetRange(paragraph.End, paragraph.End); }
        finally { Release(paragraph); }
    }

    private static void EnsureWritableParagraph(Selection selection)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        try
        {
            paragraphs = selection.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            range = paragraph.Range;
            if (!ContainsVisibleBodyText(range.Text))
                return;
            if (selection.Start >= range.End - 1)
                selection.TypeParagraph();
        }
        finally
        {
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void CompactParagraphBeforeOleDisplayFormula(
        Document document,
        Range insertion)
    {
        Range? anchor = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = Math.Max(contentStart, document.Content.End - 1);
            var position = Math.Min(Math.Max(insertion.Start, contentStart), contentEnd);
            anchor = document.Range(position, position);
            paragraphs = anchor.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            // A caret at the start of an empty paragraph belongs to that empty
            // paragraph, while the display formula will still be visually tied
            // to the preceding prose. Resolve the preceding paragraph mark in
            // that case so the local spacing adjustment is applied to the text
            // the reader actually sees above the equation.
            if (!ContainsVisibleBodyText(paragraphRange.Text)
                && position > contentStart)
            {
                Release(paragraphRange);
                paragraphRange = null;
                Release(paragraph);
                paragraph = null;
                Release(paragraphs);
                paragraphs = null;
                Release(anchor);
                anchor = null;

                anchor = document.Range(position - 1, position - 1);
                paragraphs = anchor.Paragraphs;
                if (paragraphs.Count == 0) return;
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
            }

            if (!ContainsVisibleBodyText(paragraphRange.Text)) return;
            format = paragraph.Format;
            if (Math.Abs(format.SpaceAfter - ParagraphBeforeOleDisplaySpaceAfterPoints) > 0.01f)
                format.SpaceAfter = ParagraphBeforeOleDisplaySpaceAfterPoints;
        }
        finally
        {
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(anchor);
        }
    }

    private static Range DuplicateContainingParagraphRange(Range sourceRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = sourceRange.Paragraphs;
            if (paragraphs.Count == 0)
                throw new InvalidDataException("Word 未能定位行间公式所在段落。");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            var result = paragraphRange;
            paragraphRange = null;
            return result;
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static Range? DuplicateFollowingParagraphRange(
        Document document,
        Range paragraphRange)
    {
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? result = null;
        try
        {
            if (paragraphRange.End >= document.Content.End) return null;
            var start = paragraphRange.End;
            var end = Math.Min(document.Content.End, start + 1);
            probe = document.Range(start, end);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            result = paragraph.Range.Duplicate;
            var value = result;
            result = null;
            return value;
        }
        finally
        {
            Release(result);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    private static void RemoveGeneratedParagraphAfterMathTypeRedraw(
        Document document,
        Range preservedFormulaParagraphRange,
        string? preservedFollowingParagraphText)
    {
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaRange = null;
        Range? nextProbe = null;
        Paragraphs? nextParagraphs = null;
        Paragraph? nextParagraph = null;
        Range? nextRange = null;
        InlineShapes? nextShapes = null;
        Fields? nextFields = null;
        Range? afterNextProbe = null;
        Paragraphs? afterNextParagraphs = null;
        Paragraph? afterNextParagraph = null;
        Range? afterNextRange = null;
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF"),
            "1",
            StringComparison.Ordinal);
        var cleanupWatch = tracePerformance ? Stopwatch.StartNew() : null;
        long cleanupCheckpoint = 0;
        void TraceCleanup(string stage)
        {
            if (cleanupWatch is null) return;
            var elapsed = cleanupWatch.ElapsedMilliseconds;
            WordDoubleClickHook.TraceMessage(
                $"mathtype-redraw-cleanup-perf stage={stage} deltaMs={elapsed - cleanupCheckpoint} totalMs={elapsed}");
            cleanupCheckpoint = elapsed;
        }
        try
        {
            formulaParagraphs = preservedFormulaParagraphRange.Paragraphs;
            if (formulaParagraphs.Count != 1) return;
            formulaParagraph = formulaParagraphs[1];
            formulaRange = formulaParagraph.Range.Duplicate;
            TraceCleanup("formula-range");
            // This cleanup is reached only immediately after InsertMathTypeOle
            // successfully materialized and validated Equation.DSMT4 in this exact
            // preserved redraw paragraph. Re-enumerating InlineShapes and probing
            // OLEFormat.ProgID here repeats an expensive OLE identity/layout query
            // for every formula in a large redraw batch. The structural checks below
            // still require the generated paragraph to be empty and the following
            // paragraph to match the text captured before insertion exactly.
            if (formulaRange.End >= document.Content.End) return;

            nextProbe = document.Range(
                formulaRange.End,
                Math.Min(document.Content.End, formulaRange.End + 1));
            nextParagraphs = nextProbe.Paragraphs;
            if (nextParagraphs.Count != 1) return;
            nextParagraph = nextParagraphs[1];
            nextRange = nextParagraph.Range.Duplicate;
            TraceCleanup("next-range");
            if (ContainsVisibleBodyText(nextRange.Text)) return;
            TraceCleanup("next-text");
            nextShapes = nextRange.InlineShapes;
            nextFields = WordFormulaHost.GetLocalFields(nextRange);
            TraceCleanup("next-collections");
            if (nextShapes.Count > 0 || nextFields.Count > 0) return;
            if (preservedFollowingParagraphText is null) return;

            // If the source already had a blank paragraph immediately after the
            // display formula, this empty paragraph is user-authored and must stay.
            if (string.Equals(
                    nextRange.Text ?? string.Empty,
                    preservedFollowingParagraphText,
                    StringComparison.Ordinal))
                return;
            if (nextRange.End >= document.Content.End) return;

            afterNextProbe = document.Range(
                nextRange.End,
                Math.Min(document.Content.End, nextRange.End + 1));
            afterNextParagraphs = afterNextProbe.Paragraphs;
            if (afterNextParagraphs.Count != 1) return;
            afterNextParagraph = afterNextParagraphs[1];
            afterNextRange = afterNextParagraph.Range.Duplicate;
            TraceCleanup("after-next-range");
            if (!string.Equals(
                    afterNextRange.Text ?? string.Empty,
                    preservedFollowingParagraphText,
                    StringComparison.Ordinal))
                return;
            TraceCleanup("after-next-text");

            // Word's MathType Flat OPC generated exactly one extra empty paragraph
            // between the formula and the paragraph that originally followed the
            // LaTeX source. Delete only that proven generated paragraph.
            nextRange.Delete();
            TraceCleanup("delete-generated-paragraph");
        }
        finally
        {
            Release(afterNextRange);
            Release(afterNextParagraph);
            Release(afterNextParagraphs);
            Release(afterNextProbe);
            Release(nextFields);
            Release(nextShapes);
            Release(nextRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(nextProbe);
            Release(formulaRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
        }
    }

    private static bool IsIncompleteMathTypeNumberRow(Range paragraphRange)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        var placeRefCount = 0;
        var onlyMathTypeNumberFields = true;
        try
        {
            shapes = paragraphRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (MathTypeOleInterop.IsMathTypeOle(shape)) return false;
            }

            fields = WordFormulaHost.GetLocalFields(paragraphRange);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var fieldCode = (code.Text ?? string.Empty).Trim();
                if (fieldCode.IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    placeRefCount++;
                    continue;
                }
                if (fieldCode.StartsWith("SEQ MTEqn ", StringComparison.OrdinalIgnoreCase)
                    || fieldCode.StartsWith("SEQ MTSec ", StringComparison.OrdinalIgnoreCase)
                    || fieldCode.StartsWith("SEQ MTChap ", StringComparison.OrdinalIgnoreCase))
                    continue;
                onlyMathTypeNumberFields = false;
            }
            if (placeRefCount != 1) return false;
            // A failed constructor can leave the nested MathType SEQ fields as
            // document-level siblings. Their numeric results make Paragraph.Text
            // look non-empty even though the row owns no user content at all. If
            // every field in an OLE-free row belongs to this native numbering
            // scaffold, the row is incomplete by construction and can be removed
            // atomically on rollback/retry.
            if (onlyMathTypeNumberFields && fields.Count > 0) return true;

            var text = paragraphRange.Text ?? string.Empty;
            foreach (var character in text)
            {
                if (character == '\r' || character == '\t'
                    || character == '\u0013' || character == '\u0014' || character == '\u0015'
                    || character == '\u0001' || character == '\uFFFC'
                    || char.IsWhiteSpace(character))
                    continue;
                return false;
            }
            return true;
        }
        catch { return false; }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(shape);
            Release(shapes);
        }
    }

    private static bool ClearIncompleteMathTypeNumberRowAtInsertion(
        Document document,
        Range insertion)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? body = null;
        try
        {
            paragraphs = insertion.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (insertion.Start < paragraphRange.Start || insertion.Start > paragraphRange.End)
                return false;
            if (!IsIncompleteMathTypeNumberRow(paragraphRange)) return false;

            var start = paragraphRange.Start;
            var bodyEnd = Math.Max(start, paragraphRange.End - 1);
            body = document.Range(start, bodyEnd);
            body.Delete();
            insertion.SetRange(start, start);
            return true;
        }
        finally
        {
            Release(body);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool IsRollbackOwnedMathTypeDisplayRow(Range paragraphRange)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            shapes = paragraphRange.InlineShapes;
            if (shapes.Count == 0) return false;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape)) return false;
            }

            fields = WordFormulaHost.GetLocalFields(paragraphRange);
            if (fields.Count == 0) return false;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var fieldCode = (code.Text ?? string.Empty).Trim();
                if (fieldCode.IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) >= 0
                    || fieldCode.StartsWith("SEQ MTEqn ", StringComparison.OrdinalIgnoreCase)
                    || fieldCode.StartsWith("SEQ MTSec ", StringComparison.OrdinalIgnoreCase)
                    || fieldCode.StartsWith("SEQ MTChap ", StringComparison.OrdinalIgnoreCase)
                    || fieldCode.StartsWith("EMBED Equation.DSMT4", StringComparison.OrdinalIgnoreCase))
                    continue;
                return false;
            }

            // Detached native SEQ results can expose only digits and the supported
            // equation-number punctuation outside field control characters. The OLE
            // itself is U+0001/U+FFFC. Reject any ordinary prose before treating the
            // row as transaction-owned rollback state.
            foreach (var character in paragraphRange.Text ?? string.Empty)
            {
                if (character == '\r' || character == '\t'
                    || character == '\u0013' || character == '\u0014' || character == '\u0015'
                    || character == '\u0001' || character == '\uFFFC'
                    || char.IsWhiteSpace(character)
                    || char.IsDigit(character)
                    || character is '(' or ')' or '.' or '-')
                    continue;
                return false;
            }
            return true;
        }
        catch { return false; }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(shape);
            Release(shapes);
        }
    }

    private static void RollbackStandaloneMathTypeDisplayInsertion(
        Document document,
        int insertionStart,
        int paragraphCountBeforePreparation)
    {
        Range? anchor = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? body = null;
        Range? paragraphMark = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = Math.Max(contentStart, document.Content.End - 1);
            var position = Math.Max(contentStart, Math.Min(insertionStart, contentEnd));
            anchor = document.Range(position, position);
            paragraphs = anchor.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            // A half-relocated left number can temporarily contain two complete
            // MTPlaceRef trees plus the OLE. A detached-field failure can contain
            // one outer field plus sibling SEQ fields. If the entire row is proven
            // to consist only of MathType transaction artifacts, delete its body in
            // one Word range operation. This is substantially more reliable than
            // deleting OLE/field COM objects one by one after Word has invalidated
            // some of their ranges.
            if (IsRollbackOwnedMathTypeDisplayRow(paragraphRange))
            {
                var start = paragraphRange.Start;
                var end = Math.Max(start, paragraphRange.End - 1);
                body = document.Range(start, end);
                body.Delete();
            }
            else
            {
                // Conservative compatibility fallback for failure states that do
                // not prove full transaction ownership of the row.
                shapes = paragraphRange.InlineShapes;
                for (var index = shapes.Count; index >= 1; index--)
                {
                    Release(shape);
                    shape = shapes[index];
                    if (MathTypeOleInterop.IsMathTypeOle(shape))
                        shape.Delete();
                }
                Release(shape);
                shape = null;
                Release(shapes);
                shapes = null;

                if (IsIncompleteMathTypeNumberRow(paragraphRange)
                    || !ContainsVisibleBodyText(paragraphRange.Text))
                {
                    var start = paragraphRange.Start;
                    var end = Math.Max(start, paragraphRange.End - 1);
                    body = document.Range(start, end);
                    body.Delete();
                }
            }

            // ResolveStandaloneMathTypeDisplayInsertionRange may have created one
            // dedicated blank paragraph before InsertXML. Restore the original
            // paragraph count when that paragraph was created solely for the failed
            // transaction; preserve a blank paragraph that already existed.
            if (paragraphCountBeforePreparation >= 0
                && ReadDocumentParagraphCount(document) > paragraphCountBeforePreparation)
            {
                Release(paragraphRange);
                paragraphRange = paragraph.Range;
                if (!ContainsVisibleBodyText(paragraphRange.Text))
                {
                    paragraphMark = document.Range(
                        Math.Max(paragraphRange.Start, paragraphRange.End - 1),
                        paragraphRange.End);
                    if (string.Equals(paragraphMark.Text, "\r", StringComparison.Ordinal))
                        paragraphMark.Delete();
                }
            }
        }
        catch
        {
            // Rollback is best-effort and must never mask the original Word/OLE
            // failure. The caller rethrows the original exception with stage data.
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(paragraphMark);
            Release(body);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(anchor);
        }
    }

    private static Range ResolveStandaloneMathTypeDisplayInsertionRange(
        Document document,
        Range anchor,
        bool replaceAtExactInsertion = false)
    {
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? paragraphShapes = null;
        Fields? paragraphFields = null;
        try
        {
            // A collapsed Range at a paragraph/table boundary is frequently
            // reported by Word as belonging to the PREVIOUS paragraph. That is
            // catastrophic during right-to-left format conversion: a following
            // block equation can then be appended to the previous MathType row.
            // Resolve ownership with a one-character forward probe so the anchor
            // is interpreted in the paragraph that actually owns the insertion
            // position in the remaining document.
            var contentStart = document.Content.Start;
            var probeContentEnd = document.Content.End;
            var position = Math.Max(
                contentStart,
                Math.Min(anchor.Start, Math.Max(contentStart, probeContentEnd - 1)));
            var probeEnd = Math.Min(probeContentEnd, position + 1);
            if (replaceAtExactInsertion && position < probeContentEnd)
            {
                Range? exactCharacter = null;
                Cell? exactCell = null;
                try
                {
                    // Immediately before a user table, Word 2021 can give a
                    // forward probe at the BODY paragraph mark affinity to the
                    // first table-cell paragraph. Calling InsertParagraphBefore on
                    // that cell then extends the user table backwards and absorbs
                    // the redrawn display equations. Prove that the exact character
                    // is a real body CR and reuse that physical paragraph mark.
                    exactCharacter = document.Range(position, probeEnd);
                    if (string.Equals(exactCharacter.Text, "\r", StringComparison.Ordinal))
                    {
                        exactCell = WordFormulaHost.TryGetOwningCell(exactCharacter);
                        WordDoubleClickHook.TraceMessage(
                            $"mathtype-display-exact-owner position={position} cell={(exactCell is null ? "body" : exactCell.RowIndex + "," + exactCell.ColumnIndex)}");
                        if (exactCell is null)
                        {
                            Range? followingProbe = null;
                            Tables? followingTables = null;
                            Table? followingTable = null;
                            Range? followingTableRange = null;
                            Range? placeholderInsertion = null;
                            Range? placeholderRange = null;
                            try
                            {
                                // Word 2021 can adopt a Flat OPC <w:p> into the
                                // following table whenever InsertXML targets the last
                                // collapsed BODY position before that table. Make the
                                // replacement target unambiguously body-owned: insert
                                // one private-use character into the existing empty
                                // paragraph and return that exact non-collapsed range.
                                // InsertXML replaces the character atomically, so no
                                // user text or table structure participates.
                                var followingStart = position + 1;
                                if (followingStart < probeContentEnd)
                                {
                                    followingProbe = document.Range(
                                        followingStart,
                                        Math.Min(probeContentEnd, followingStart + 1));
                                    followingTables = followingProbe.Tables;
                                    if (followingTables.Count > 0)
                                    {
                                        followingTable = followingTables[1];
                                        followingTableRange = followingTable.Range;
                                        if (followingTableRange.Start == followingStart)
                                        {
                                            placeholderInsertion = document.Range(position, position);
                                            placeholderInsertion.Text = BulkInlineFormulaPlaceholder;
                                            placeholderRange = document.Range(
                                                position,
                                                position + BulkInlineFormulaPlaceholder.Length);
                                            WordDoubleClickHook.TraceMessage(
                                                $"mathtype-display-table-boundary-placeholder position={position} tableStart={followingStart}");
                                            var result = placeholderRange;
                                            placeholderRange = null;
                                            return result;
                                        }
                                    }
                                }
                            }
                            finally
                            {
                                Release(placeholderRange);
                                Release(placeholderInsertion);
                                Release(followingTableRange);
                                Release(followingTable);
                                Release(followingTables);
                                Release(followingProbe);
                            }
                            return document.Range(position, position);
                        }
                    }
                }
                finally
                {
                    Release(exactCell);
                    Release(exactCharacter);
                }
            }
            probe = document.Range(position, probeEnd);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count == 0) return anchor.Duplicate;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            paragraphShapes = paragraphRange.InlineShapes;
            paragraphFields = WordFormulaHost.GetLocalFields(paragraphRange);
            var hasInlineShapes = paragraphShapes.Count > 0;
            var hasFields = paragraphFields.Count > 0;
            if (IsMathTypeSectionStateParagraph(paragraphRange)
                && !hasInlineShapes
                && !ContainsVisibleBodyText(paragraphRange.Text))
            {
                // A collapsed Range exactly on a paragraph boundary is often
                // reported by Word as belonging to the previous paragraph.  When
                // that previous paragraph is MathType's hidden section state, the
                // user caret is logically in the following paragraph, so never
                // snap back to the hidden paragraph start.
                var nextParagraphStart = Math.Max(
                    document.Content.Start,
                    Math.Min(
                        paragraphRange.End,
                        Math.Max(document.Content.Start, document.Content.End - 1)));
                return document.Range(nextParagraphStart, nextParagraphStart);
            }
            if (!ContainsVisibleBodyText(paragraphRange.Text)
                && !hasInlineShapes
                && !hasFields)
                return document.Range(paragraphRange.Start, paragraphRange.Start);

            if (replaceAtExactInsertion)
            {
                // Format conversion has already deleted the original display host;
                // anchor.Start is the exact replacement position. At a paragraph
                // boundary Word reports the following paragraph as the owner of the
                // collapsed range. Inserting after that paragraph reverses adjacent
                // equations during right-to-left conversion. Reserve a fresh row
                // immediately BEFORE the following content instead.
                var exactStart = Math.Max(
                    document.Content.Start,
                    Math.Min(anchor.Start, Math.Max(document.Content.Start, document.Content.End - 1)));
                paragraphRange.InsertParagraphBefore();
                return document.Range(exactStart, exactStart);
            }

            // Flat OPC always carries its own <w:p>. Inserting it immediately before
            // the final paragraph mark does not automatically create a clean display
            // paragraph; Word can put the OLE beside existing prose. Create the one
            // required blank paragraph first, then materialize the MathType object
            // into that paragraph.
            var newParagraphStart = paragraphRange.End;
            paragraphRange.InsertParagraphAfter();
            var contentEnd = document.Content.End;
            newParagraphStart = Math.Max(
                document.Content.Start,
                Math.Min(newParagraphStart, Math.Max(document.Content.Start, contentEnd - 1)));
            return document.Range(newParagraphStart, newParagraphStart);
        }
        finally
        {
            Release(paragraphFields);
            Release(paragraphShapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    private static bool IsMathTypeSectionStateParagraph(Range range)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            fields = WordFormulaHost.GetLocalFields(range);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static Range PrepareNumberedOmmlReplacementTabPlaceholderPreservingOle(
        Document document,
        InlineShape sourceShape)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? shapeRange = null;
        Range? before = null;
        Range? after = null;
        Range? scaffold = null;
        Range? placeholder = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            shapeRange = sourceShape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The numbered OLE replacement no longer occupies one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (shapeRange.Start < paragraphRange.Start
                || shapeRange.End > paragraphRange.End)
                throw new InvalidOperationException(
                    "The numbered OLE replacement range escaped its paragraph.");

            maths = paragraphRange.OMaths;
            shapes = paragraphRange.InlineShapes;
            fields = WordFormulaHost.GetLocalFields(paragraphRange);
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                Console.WriteLine(
                    $"    [diagnostic] numbered OLE→OMML host: paragraph={paragraphRange.Start}:{paragraphRange.End}, shape={shapeRange.Start}:{shapeRange.End}, maths={maths.Count}, shapes={shapes.Count}, fields={fields.Count}.");
                for (var diagnosticIndex = 1; diagnosticIndex <= fields.Count; diagnosticIndex++)
                {
                    Field? diagnosticField = null;
                    Range? diagnosticCode = null;
                    Range? diagnosticResult = null;
                    try
                    {
                        diagnosticField = fields[diagnosticIndex];
                        diagnosticCode = diagnosticField.Code;
                        diagnosticResult = diagnosticField.Result;
                        Console.WriteLine(
                            $"    [diagnostic] numbered OLE→OMML field#{diagnosticIndex}: type={diagnosticField.Type}, code={diagnosticCode.Start}:{diagnosticCode.End} '{diagnosticCode.Text}', result={diagnosticResult.Start}:{diagnosticResult.End} '{diagnosticResult.Text}'.");
                    }
                    finally
                    {
                        Release(diagnosticResult);
                        Release(diagnosticCode);
                        Release(diagnosticField);
                    }
                }
            }
            if (maths.Count != 0
                || shapes.Count != 1
                || !ContainsOnlySourceOleFields(fields, sourceShape))
                throw new InvalidOperationException(
                    "VisualTeX refused to rebuild a numbered OLE paragraph containing another formula, object, or unrelated field.");

            var editableEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (shapeRange.End < editableEnd)
            {
                after = document.Range(shapeRange.End, editableEnd);
                if (ContainsVisibleBodyText(after.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text after the numbered OLE formula.");
                after.Text = string.Empty;
            }

            // Delete structural TAB/zero-width runs before the OLE only after the
            // trailing side has been removed, because deleting the prefix shifts the
            // InlineShape's live Word coordinates.
            Release(shapeRange);
            shapeRange = sourceShape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            if (paragraphRange.Start < shapeRange.Start)
            {
                before = document.Range(paragraphRange.Start, shapeRange.Start);
                if (ContainsVisibleBodyText(before.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text before the numbered OLE formula.");
                before.Text = string.Empty;
            }

            Release(shapeRange);
            shapeRange = sourceShape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            var start = paragraphRange.Start;
            scaffold = document.Range(start, start);
            scaffold.Text = "\t" + BulkInlineFormulaPlaceholder + "\t";
            placeholder = document.Range(
                start + 1,
                start + 1 + BulkInlineFormulaPlaceholder.Length);
            var result = placeholder;
            placeholder = null;
            return result;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(placeholder);
            Release(scaffold);
            Release(after);
            Release(before);
            Release(shapeRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool ContainsOnlySourceOleFields(
        Fields fields,
        InlineShape sourceShape)
    {
        if (fields.Count == 0) return true;
        Range? shapeRange = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        try
        {
            shapeRange = sourceShape.Range;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(result); result = null;
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                result = field.Result;
                var codeText = code.Text ?? string.Empty;
                var isEmbed = field.Type == WdFieldType.wdFieldEmbed
                    || codeText.IndexOf(
                        "EMBED ",
                        StringComparison.OrdinalIgnoreCase) >= 0;
                // For embedded OLE Word reports InlineShape.Range as the complete
                // field span including the begin/separator/end characters, while
                // Field.Code/Result exclude those boundary characters. Reconstruct
                // that full span instead of expecting Result alone to contain the
                // InlineShape range.
                var fullFieldStart = Math.Max(0, code.Start - 1);
                var fullFieldEnd = result.End + 1;
                var ownsSourceShape = fullFieldStart <= shapeRange.Start
                    && fullFieldEnd >= shapeRange.End;
                if (!isEmbed || !ownsSourceShape)
                    return false;
            }
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(shapeRange);
        }
    }

    private static Range PrepareNumberedOmmlTrueDisplayReplacementPlaceholder(
        Document document,
        Range equationRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? before = null;
        Range? after = null;
        Range? editableRange = null;
        Range? placeholder = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            paragraphs = equationRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The true-display numbered OMML replacement no longer occupies one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable)
                || equationRange.Start < paragraphRange.Start
                || equationRange.End > paragraphRange.End)
                throw new InvalidOperationException(
                    "The true-display numbered OMML replacement escaped its table-free formula paragraph.");

            maths = paragraphRange.OMaths;
            shapes = paragraphRange.InlineShapes;
            fields = paragraphRange.Fields;
            if (maths.Count != 1 || shapes.Count != 0 || fields.Count != 0)
                throw new InvalidOperationException(
                    "VisualTeX refused to replace a true-display OMML paragraph containing another formula, object, or field.");
            var paragraphXml = paragraphRange.WordOpenXML ?? string.Empty;
            if (paragraphXml.IndexOf("<m:oMathPara", StringComparison.OrdinalIgnoreCase) < 0
                || paragraphXml.IndexOf("<m:eqArr", StringComparison.OrdinalIgnoreCase) >= 0
                || paragraphXml.IndexOf("<w:fldChar", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new InvalidOperationException(
                    "The numbered OMML source is not a pure genuine-display paragraph.");

            if (paragraphRange.Start < equationRange.Start)
            {
                before = document.Range(paragraphRange.Start, equationRange.Start);
                if (ContainsVisibleBodyText(before.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text before the true-display OMML formula.");
            }
            var editableEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (equationRange.End < editableEnd)
            {
                after = document.Range(equationRange.End, editableEnd);
                if (ContainsVisibleBodyText(after.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text after the true-display OMML formula.");
            }

            var start = paragraphRange.Start;
            editableRange = document.Range(start, editableEnd);
            editableRange.Text = BulkInlineFormulaPlaceholder;
            placeholder = document.Range(
                start,
                start + BulkInlineFormulaPlaceholder.Length);
            var result = placeholder;
            placeholder = null;
            return result;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(placeholder);
            Release(editableRange);
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static Range PrepareNumberedOmmlReplacementTabPlaceholder(
        Document document,
        Range equationRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? before = null;
        Range? after = null;
        Range? editableRange = null;
        Range? placeholder = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            paragraphs = equationRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The numbered OMML replacement no longer occupies one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (equationRange.Start < paragraphRange.Start
                || equationRange.End > paragraphRange.End)
                throw new InvalidOperationException(
                    "The numbered OMML replacement range escaped its paragraph.");

            maths = paragraphRange.OMaths;
            shapes = paragraphRange.InlineShapes;
            fields = paragraphRange.Fields;
            if (maths.Count != 1 || shapes.Count != 0 || fields.Count != 0)
                throw new InvalidOperationException(
                    "VisualTeX refused to rebuild a numbered OMML paragraph containing another formula, object, or field.");

            if (paragraphRange.Start < equationRange.Start)
            {
                before = document.Range(paragraphRange.Start, equationRange.Start);
                if (ContainsVisibleBodyText(before.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text before the numbered OMML formula.");
            }
            var editableEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (equationRange.End < editableEnd)
            {
                after = document.Range(equationRange.End, editableEnd);
                if (ContainsVisibleBodyText(after.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text after the numbered OMML formula.");
            }

            editableRange = document.Range(paragraphRange.Start, editableEnd);
            editableRange.Text = "\t" + BulkInlineFormulaPlaceholder + "\t";
            placeholder = document.Range(
                paragraphRange.Start + 1,
                paragraphRange.Start + 1 + BulkInlineFormulaPlaceholder.Length);
            var result = placeholder;
            placeholder = null;
            return result;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(placeholder);
            Release(editableRange);
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static Range ResolveDisplayInsertionRange(
        Document document,
        Range anchor,
        bool replaceAtExactInsertion = false)
    {
        Range? probe = null;
        Range? exactCharacter = null;
        Cell? exactCell = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? content = null;
        try
        {
            content = document.Content;
            var lastInsertPosition = Math.Max(content.Start, content.End - 1);
            var exactPosition = Math.Max(
                content.Start,
                Math.Min(anchor.Start, lastInsertPosition));
            if (replaceAtExactInsertion)
            {
                // An atomic source replacement can leave one real empty BODY
                // paragraph immediately before a user table. Word 2021 gives a
                // one-character forward probe at that boundary table affinity, so
                // probe.Paragraphs can become the first cell paragraph and move the
                // converted display formula into the table. Prove the exact CR is
                // not owned by a cell and reuse that body paragraph directly.
                if (exactPosition < content.End)
                {
                    exactCharacter = document.Range(
                        exactPosition,
                        Math.Min(content.End, exactPosition + 1));
                    var exactText = exactCharacter.Text ?? string.Empty;
                    WordDoubleClickHook.TraceMessage(
                        $"display-insertion-exact-probe position={exactPosition} chars=[{string.Join(",", exactText.Select(character => ((int)character).ToString("X4")))}] range={exactCharacter.Start}:{exactCharacter.End}");
                    if (string.Equals(
                            exactText,
                            "\r",
                            StringComparison.Ordinal))
                    {
                        exactCell = WordFormulaHost.TryGetOwningCell(exactCharacter);
                        WordDoubleClickHook.TraceMessage(
                            $"display-insertion-exact-owner position={exactPosition} cell={(exactCell is null ? "body" : exactCell.RowIndex + "," + exactCell.ColumnIndex)}");
                        if (exactCell is null)
                            return document.Range(exactPosition, exactPosition);
                    }
                    Release(exactCell);
                    exactCell = null;
                    Release(exactCharacter);
                    exactCharacter = null;
                }

                // At other paragraph boundaries a collapsed Word range can report
                // the paragraph on either side. Probe one character forward so an
                // in-place format conversion never appends the target formula to
                // the paragraph that originally followed the source equation.
                probe = document.Range(
                    exactPosition,
                    Math.Min(content.End, exactPosition + 1));
                paragraphs = probe.Paragraphs;
            }
            else
            {
                paragraphs = anchor.Paragraphs;
            }
            if (paragraphs.Count == 0)
                return anchor.Duplicate;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            var hasVisibleText = ContainsVisibleBodyText(paragraphRange.Text);
            if (replaceAtExactInsertion)
            {
                if (!hasVisibleText)
                    return document.Range(paragraphRange.Start, paragraphRange.Start);

                // The source display host has been removed completely and the
                // exact replacement point now coincides with the start of the
                // following user paragraph. Reserve one paragraph before that
                // content; this is the formula host, not an additional blank line.
                if (exactPosition <= paragraphRange.Start)
                {
                    paragraphRange.InsertParagraphBefore();
                    return document.Range(exactPosition, exactPosition);
                }
                throw new InvalidOperationException(
                    "VisualTeX refused to insert a converted display formula inside user paragraph text.");
            }

            // Reuse an existing empty paragraph instead of creating another one.
            // At the physical document end, however, paragraphRange.End lies one
            // position beyond the last legal insertion point. Clamping it back to
            // Content.End - 1 places the caret at the preceding OMath boundary and
            // Word absorbs the next display formula into that same equation. Create
            // the following paragraph explicitly and insert at its start instead.
            if (hasVisibleText && paragraphRange.End > lastInsertPosition)
            {
                var nextParagraphStart = paragraphRange.End;
                paragraphRange.InsertParagraphAfter();
                return document.Range(nextParagraphStart, nextParagraphStart);
            }

            // For ordinary body text with a following paragraph, its existing
            // boundary remains a stable display insertion point.
            var position = hasVisibleText
                ? paragraphRange.End
                : paragraphRange.Start;
            position = Math.Max(
                content.Start,
                Math.Min(position, lastInsertPosition));
            return document.Range(position, position);
        }
        finally
        {
            Release(exactCell);
            Release(exactCharacter);
            Release(content);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    private static void FormatExistingDisplayParagraph(
        Range paragraphRange,
        bool preserveNativeOmmlSpacing)
    {
        ParagraphFormat? format = null;
        try
        {
            format = paragraphRange.ParagraphFormat;
            format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            if (!preserveNativeOmmlSpacing)
            {
                format.SpaceBefore = 0;
                format.SpaceAfter = 0;
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            }
            try { paragraphRange.ListFormat.RemoveNumbers(); } catch { }
        }
        finally { Release(format); }
    }

    private static bool CanUseDedicatedDisplayParagraphFastPath(
        Document document,
        Range source)
    {
        Range? previous = null;
        Range? next = null;
        Cells? cells = null;
        Cell? cell = null;
        Range? cellRange = null;
        try
        {
            if (source.StoryType != WdStoryType.wdMainTextStory)
                return false;

            // A Word table cell serializes its final structural position as CR+BEL
            // in Range.Text even though that pair occupies one story coordinate.
            // Therefore probing source.End with a one-character Range does not
            // reliably return a lone '\r' or '\a'. If the LaTeX source owns the
            // complete editable cell body, it already has the exact display host we
            // need: deleting only source.Start..source.End leaves the cell's single
            // paragraph/cell terminator in place. Do not insert synthetic CRs around
            // it, otherwise one source cell becomes three paragraphs after redraw.
            cell = WordFormulaHost.TryGetOwningCell(source);
            if (cell is not null)
            {
                cellRange = cell.Range;
                var editableCellEnd = Math.Max(cellRange.Start, cellRange.End - 1);
                if (source.Start == cellRange.Start
                    && source.End == editableCellEnd)
                    return true;
            }

            var contentStart = document.Content.Start;
            var contentEnd = document.Content.End;
            if (source.End >= contentEnd)
                return false;
            next = document.Range(source.End, Math.Min(contentEnd, source.End + 1));
            if (WordFormulaHost.ParagraphTerminatorStoryLength(next.Text) == 0)
                return false;
            if (source.Start <= contentStart)
                return true;
            previous = document.Range(source.Start - 1, source.Start);
            return WordFormulaHost.ParagraphTerminatorStoryLength(previous.Text) == 1;
        }
        finally
        {
            Release(cellRange);
            Release(cell);
            Release(cells);
            Release(next);
            Release(previous);
        }
    }

    // A display source owns its formula text and the immediately adjacent line
    // separators. Split both sides before deleting the source so every format
    // writes into the same exact empty paragraph, and existing prose keeps its
    // character/paragraph formatting. CR, manual BR and cell ends are distinct.
    private static Range IsolateDisplaySourceParagraph(Document document, Range source)
    {
        var start = source.Start;
        var end = source.End;
        var expected = source.Text;
        Range? boundary = null;
        Range? isolated = null;
        try
        {
            if (source.StoryType != WdStoryType.wdMainTextStory)
                throw new InvalidDataException("Display redraw requires a main-story source range.");
            boundary = document.Range(end, Math.Min(document.Content.End, end + 1));
            if (boundary.Text == "\v") PromoteDisplayLineBoundary(boundary);
            else if (WordFormulaHost.ParagraphTerminatorStoryLength(boundary.Text) == 0)
            {
                boundary.SetRange(end, end);
                boundary.Text = "\r";
            }
            Release(boundary); boundary = null;
            if (start > document.Content.Start)
            {
                boundary = document.Range(start - 1, start);
                if (boundary.Text == "\v") PromoteDisplayLineBoundary(boundary);
                else if (WordFormulaHost.ParagraphTerminatorStoryLength(boundary.Text) == 0)
                {
                    boundary.SetRange(start, start);
                    boundary.Text = "\r";
                    start++; end++;
                }
            }
            isolated = document.Range(start, end);
            if (!string.Equals(isolated.Text, expected, StringComparison.Ordinal))
                throw new InvalidDataException("The display source changed while isolating its paragraph.");

            // The explicit CR/cell-end checks above already prove that this
            // source starts and ends on paragraph boundaries. Re-enumerating
            // Range.Paragraphs here is redundant, and Word compatibility mode
            // makes that COM query progressively slower as more OLE objects are
            // inserted, turning a large redraw into an O(N^2) paragraph walk.
            var result = isolated; isolated = null;
            return result;
        }
        finally
        {
            Release(isolated); Release(boundary);
        }
    }

    private static void PromoteDisplayLineBoundary(Range boundary)
    {
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        Frames? frames = null;
        try
        {
            maths = boundary.OMaths; shapes = boundary.InlineShapes;
            fields = boundary.Fields; bookmarks = boundary.Bookmarks; frames = boundary.Frames;
            if (boundary.Text != "\v" || maths.Count != 0 || shapes.Count != 0
                || fields.Count != 0 || bookmarks.Count != 0 || frames.Count != 0)
                throw new InvalidDataException("A display line boundary contains another object or identity.");
            boundary.Text = "\r";
        }
        finally
        {
            Release(frames); Release(bookmarks); Release(fields); Release(shapes); Release(maths);
        }
    }

    private static void EnsureBlankDisplayParagraph(
        Selection selection,
        bool preserveNativeOmmlSpacing)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        try
        {
            paragraphs = selection.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                range = paragraph.Range;
                if (ContainsVisibleBodyText(range.Text))
                    selection.TypeParagraph();
            }
            selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            if (!preserveNativeOmmlSpacing)
            {
                selection.ParagraphFormat.SpaceBefore = 0;
                selection.ParagraphFormat.SpaceAfter = 0;
                selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            }
        }
        finally
        {
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void MoveSelectionAfterNumberedDisplayFormula(
        Document document,
        Selection selection,
        Range equationRange,
        string formulaId,
        Table? knownDirectTable = null)
    {
        Range? typingRange = null;
        Range? ownerRange = null;
        Range? captionRange = null;
        Range? content = null;
        try
        {
            typingRange =
                WordEquationNumbering.EnsureNormalTypingParagraphAfterNumberedDisplay(
                    document,
                    formulaId,
                    knownDirectTable);
            if (typingRange is not null)
            {
                selection.SetRange(typingRange.Start, typingRange.Start);
                selection.Collapse(WdCollapseDirection.wdCollapseStart);
                return;
            }

            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(
                document,
                formulaId);
            captionRange = WordEquationNumbering.FindNativeEquationCaptionRange(
                document,
                formulaId);
            content = document.Content;
            var target = equationRange.End;
            if (ownerRange is not null) target = Math.Max(target, ownerRange.End);
            if (captionRange is not null) target = Math.Max(target, captionRange.End);
            target = Math.Max(content.Start, Math.Min(target, content.End));
            selection.SetRange(target, target);
            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        }
        finally
        {
            Release(content);
            Release(captionRange);
            Release(ownerRange);
            Release(typingRange);
        }
    }

    private static int ReadParagraphStart(Range anchor)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = anchor.Paragraphs;
            if (paragraphs.Count == 0) return anchor.Start;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return paragraphRange.Start;
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool IsStructurallyEmptyParagraph(Range range)
    {
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            maths = range.OMaths;
            shapes = range.InlineShapes;
            fields = WordFormulaHost.GetLocalFields(range);
            var text = (range.Text ?? string.Empty)
                .Trim('\r', '\a', '\v', '\f', '\t', ' ');
            return text.Length == 0
                && maths.Count == 0
                && shapes.Count == 0
                && fields.Count == 0;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(maths);
        }
    }

    private static void RepairPreservedDisplayParagraphBoundary(
        Document document,
        Range formulaRange,
        int sourceParagraphStart,
        int sourceParagraphCount)
    {
        if (sourceParagraphStart < 0 || sourceParagraphCount < 0) return;
        var currentParagraphCount = ReadDocumentParagraphCount(document);
        if (currentParagraphCount == sourceParagraphCount) return;
        if (currentParagraphCount != sourceParagraphCount + 1)
            throw new InvalidOperationException(
                $"Word changed the paragraph count unexpectedly while converting a display formula to OMML: before={sourceParagraphCount}, after={currentParagraphCount}.");

        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Range? content = null;
        Range? probe = null;
        Paragraphs? candidateParagraphs = null;
        Paragraph? candidateParagraph = null;
        Range? candidateRange = null;
        try
        {
            formulaParagraphs = formulaRange.Paragraphs;
            if (formulaParagraphs.Count != 1)
                throw new InvalidOperationException(
                    "The converted OMML display formula spans multiple paragraphs.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range;
            content = document.Content;

            if (formulaParagraphRange.Start > sourceParagraphStart)
            {
                var previousStart = Math.Max(content.Start, formulaParagraphRange.Start - 1);
                probe = document.Range(previousStart, formulaParagraphRange.Start);
                candidateParagraphs = probe.Paragraphs;
                if (candidateParagraphs.Count == 0)
                    throw new InvalidOperationException(
                        "Word inserted an OMML paragraph before the source boundary, but the residual source paragraph could not be resolved.");
                candidateParagraph = candidateParagraphs[1];
                candidateRange = candidateParagraph.Range;
                if (candidateRange.Start != sourceParagraphStart
                    || !IsStructurallyEmptyParagraph(candidateRange))
                    throw new InvalidOperationException(
                        "The paragraph before the converted OMML formula is not the empty source paragraph VisualTeX expected to repair.");
            }
            else if (formulaParagraphRange.Start == sourceParagraphStart)
            {
                if (formulaParagraphRange.End >= content.End)
                    throw new InvalidOperationException(
                        "Word added a display-formula paragraph split at the document end, but no residual paragraph is available to repair.");
                probe = document.Range(
                    formulaParagraphRange.End,
                    Math.Min(content.End, formulaParagraphRange.End + 1));
                candidateParagraphs = probe.Paragraphs;
                if (candidateParagraphs.Count == 0)
                    throw new InvalidOperationException(
                        "Word inserted an OMML paragraph after the source boundary, but the residual paragraph could not be resolved.");
                candidateParagraph = candidateParagraphs[1];
                candidateRange = candidateParagraph.Range;
                if (!IsStructurallyEmptyParagraph(candidateRange))
                    throw new InvalidOperationException(
                        "The paragraph after the converted OMML formula contains user content and was not removed.");
            }
            else
            {
                throw new InvalidOperationException(
                    "The converted OMML display formula moved before its captured source paragraph boundary.");
            }

            candidateRange.Delete();
        }
        finally
        {
            Release(candidateRange);
            Release(candidateParagraph);
            Release(candidateParagraphs);
            Release(probe);
            Release(content);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
        }

        var repairedParagraphCount = ReadDocumentParagraphCount(document);
        if (repairedParagraphCount != sourceParagraphCount)
            throw new InvalidOperationException(
                $"VisualTeX could not restore the display formula's original paragraph structure: expected={sourceParagraphCount}, actual={repairedParagraphCount}.");
    }

    private static void MoveSelectionAfterDisplayFormula(
        Selection selection,
        Range formulaRange)
    {
        selection.SetRange(formulaRange.End, formulaRange.End);
        selection.TypeParagraph();
        selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        selection.ParagraphFormat.LeftIndent = 0;
        selection.ParagraphFormat.FirstLineIndent = 0;
        object normal = WdBuiltinStyle.wdStyleNormal;
        try { selection.Range.set_Style(ref normal); } catch { }
        try { selection.Range.ListFormat.RemoveNumbers(); } catch { }
        ResetSelectionTransientFormatting(selection);
    }

    private static void ResetSelectionTransientFormatting(Selection selection)
    {
        Range? caret = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Microsoft.Office.Interop.Word.Font? selectionFont = null;
        Microsoft.Office.Interop.Word.Font? paragraphMarkFont = null;
        try
        {
            selectionFont = selection.Font;
            ResetTransientFont(selectionFont);

            // A collapsed Word Selection can report neutral formatting while the
            // paragraph mark still stores the old direct italic/bold state. The
            // next typed character inherits the paragraph mark, not the temporary
            // Selection.Font value. Clear that mark only when the caret is at the
            // paragraph boundary used for subsequent typing.
            if (selection.Start != selection.End) return;
            caret = selection.Range;
            paragraphRange = WordCharacterFormatting.ResolveParagraphAtPosition(caret, selection.Start);
            if (selection.Start < paragraphRange.End - 1) return;
            paragraphMark = paragraphRange.Duplicate;
            paragraphMark.SetRange(
                Math.Max(paragraphRange.Start, paragraphRange.End - 1),
                paragraphRange.End);
            paragraphMarkFont = paragraphMark.Font;
            ResetTransientFont(paragraphMarkFont);
        }
        finally
        {
            Release(paragraphMarkFont);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(caret);
            Release(selectionFont);
        }
    }

    private static void ResetTransientFont(Microsoft.Office.Interop.Word.Font font)
    {
        font.Bold = 0;
        font.Italic = 0;
        font.StrikeThrough = 0;
        try { font.DoubleStrikeThrough = 0; } catch { }
        font.Underline = WdUnderline.wdUnderlineNone;
        font.Hidden = 0;
        font.Subscript = 0;
        font.Superscript = 0;
        font.Position = 0;
        try { font.AllCaps = 0; } catch { }
        try { font.SmallCaps = 0; } catch { }
    }

    private static void InsertNativeTextRun(
        Document document,
        Selection selection,
        WordBulkRun run)
    {
        if (string.IsNullOrEmpty(run.Text)) return;
        ResetSelectionTransientFormatting(selection);
        var start = selection.Start;
        selection.TypeText(run.Text);
        Range? inserted = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            inserted = document.Range(start, selection.Start);
            font = inserted.Font;
            font.Bold = run.Bold ? 1 : 0;
            font.Italic = run.Italic ? 1 : 0;
            font.StrikeThrough = run.Strike ? 1 : 0;
            try { font.DoubleStrikeThrough = 0; } catch { }
            font.Underline = run.Underline
                ? WdUnderline.wdUnderlineSingle
                : WdUnderline.wdUnderlineNone;
            font.Hidden = 0;
            font.Subscript = 0;
            font.Superscript = 0;
            font.Position = 0;
            try { font.AllCaps = 0; } catch { }
            try { font.SmallCaps = 0; } catch { }
            if (run.Code)
            {
                font.Name = "Consolas";
                try { font.NameAscii = "Consolas"; } catch { }
                try { font.NameFarEast = "Microsoft YaHei UI"; } catch { }
            }
        }
        finally
        {
            Release(font);
            Release(inserted);
            // Range formatting can update Word's collapsed typing state. Keep
            // placeholders, following runs and the next paragraph neutral; the
            // next source run reapplies its own explicit semantics.
            ResetSelectionTransientFormatting(selection);
        }
    }

    private static void ApplyBulkParagraphFormatting(
        Document document,
        int start,
        int end,
        WordBulkBlock block)
    {
        if (end < start) return;
        Range? range = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ListFormat? listFormat = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = document.Range(start, end);
            paragraphs = range.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            switch (block.Kind)
            {
                case WordBulkBlockKind.Heading:
                    object heading = block.Level switch
                    {
                        <= 1 => WdBuiltinStyle.wdStyleHeading1,
                        2 => WdBuiltinStyle.wdStyleHeading2,
                        3 => WdBuiltinStyle.wdStyleHeading3,
                        4 => WdBuiltinStyle.wdStyleHeading4,
                        5 => WdBuiltinStyle.wdStyleHeading5,
                        6 => WdBuiltinStyle.wdStyleHeading6,
                        7 => WdBuiltinStyle.wdStyleHeading7,
                        8 => WdBuiltinStyle.wdStyleHeading8,
                        _ => WdBuiltinStyle.wdStyleHeading9,
                    };
                    range.set_Style(ref heading);
                    break;
                case WordBulkBlockKind.Bullet:
                    listFormat = range.ListFormat;
                    listFormat.ApplyBulletDefault();
                    for (var level = 0; level < Math.Min(block.Level, 8); level++)
                        listFormat.ListIndent();
                    break;
                case WordBulkBlockKind.Numbered:
                    listFormat = range.ListFormat;
                    listFormat.ApplyNumberDefault();
                    for (var level = 0; level < Math.Min(block.Level, 8); level++)
                        listFormat.ListIndent();
                    break;
                case WordBulkBlockKind.Quote:
                    paragraph.LeftIndent = 18f;
                    paragraph.RightIndent = 9f;
                    font = range.Font;
                    font.Italic = 1;
                    break;
                case WordBulkBlockKind.Code:
                    font = range.Font;
                    font.Name = "Consolas";
                    try { font.NameAscii = "Consolas"; } catch { }
                    paragraph.LeftIndent = 18f;
                    paragraph.SpaceBefore = 3f;
                    paragraph.SpaceAfter = 3f;
                    break;
            }

        }
        finally
        {
            Release(font);
            Release(listFormat);
            Release(paragraph);
            Release(paragraphs);
            Release(range);
        }
    }

    private static void ResetNextParagraphFormatting(
        Selection selection,
        WordBulkBlockKind current,
        WordBulkBlockKind? next)
    {
        var continuingList =
            current == WordBulkBlockKind.Bullet && next == WordBulkBlockKind.Bullet
            || current == WordBulkBlockKind.Numbered && next == WordBulkBlockKind.Numbered;
        if (continuingList)
        {
            // Preserve list numbering/indentation, but never carry the previous
            // item's final bold/italic run into the next list item.
            ResetSelectionTransientFormatting(selection);
            return;
        }
        Range? caret = null;
        Range? nextParagraph = null;
        ListFormat? list = null;
        ParagraphFormat? format = null;
        try
        {
            caret = selection.Range;
            nextParagraph = WordCharacterFormatting.ResolveParagraphAtPosition(caret, selection.Start);
            // An import in the middle of existing text can leave user content in
            // the following paragraph. Only initialize an empty typing paragraph.
            if ((nextParagraph.Text ?? string.Empty).Trim('\r', '\a').Length != 0) return;
            list = nextParagraph.ListFormat;
            list.RemoveNumbers();
            format = nextParagraph.ParagraphFormat;
            format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            format.LeftIndent = 0;
            format.RightIndent = 0;
            format.FirstLineIndent = 0;
            object normal = WdBuiltinStyle.wdStyleNormal;
            nextParagraph.set_Style(ref normal);
            ResetSelectionTransientFormatting(selection);
        }
        finally
        {
            Release(format);
            Release(list);
            Release(nextParagraph);
            Release(caret);
        }
    }

    public OfficeObjectResult ReplaceMathTypeOle(
        OfficeSessionDocument session,
        string mathMl,
        string? emfPath)
    {
        if (string.IsNullOrWhiteSpace(mathMl)
            || !mathMl.TrimStart().StartsWith("<math", StringComparison.Ordinal))
            throw new InvalidDataException("VisualTeX did not provide valid MathML for MathType OLE.");

        var metadata = session.ToMetadata();
        metadata.Validate();
        if (string.IsNullOrWhiteSpace(metadata.Latex))
            throw new InvalidDataException(
                "VisualTeX did not provide LaTeX source for the MathType OLE update.");

        Document? document = null;
        InlineShape? oldShape = null;
        InlineShape? replacement = null;
        Range? oldRange = null;
        Range? insertion = null;
        Range? finalSelection = null;
        UndoRecord? undoRecord = null;
        WordViewState? viewState = null;
        string? rollbackWordOpenXml = null;
        var rollbackStart = -1;
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        var oldDeleted = false;
        MathTypeDisplayParagraphLayout? displayParagraphLayout = null;
        MathTypeDisplayParagraphLayout? inlineParagraphLayout = null;
        MathTypeNativePreviewRenderer.Result? sourceNativePreview = null;
        MathTypeNativePreviewRenderer.Result? nativePreview = null;
        var sourceParagraphCount = -1;
        var sourceWasNumbered = false;
        var sourceNumberPosition = "right";
        var numberingLayoutChanged = false;
        MathTypeWordOpenXml.NumberTemplate? sourceNumberTemplate = null;
        var createdEditSectionBreakCodeStart = -1;
        Options? editOptions = null;
        var previousSmartCutPaste = false;
        var smartCutPasteSuspended = false;
        var editUndoEnded = false;
        var editViewRestored = false;
        var alignInline = string.Equals(
            session.DisplayMode,
            "inline",
            StringComparison.OrdinalIgnoreCase);
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            viewState = CaptureViewState();
            try
            {
                previousScreenUpdating = _application.ScreenUpdating;
                _application.ScreenUpdating = false;
                screenUpdatingSuspended = true;
            }
            catch { }

            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=resolve-source formulaId={session.FormulaId}");
            oldShape = FindMathTypeOleByRange(document, session.SourceObjectId)
                ?? throw new InvalidOperationException(
                    "The MathType OLE equation no longer exists at the captured Word location.");
            if (!MathTypeOleInterop.IsMathTypeOle(oldShape))
                throw new InvalidOperationException(
                    "The selected OLE object is no longer recognized as MathType Equation.DSMT4.");

            oldRange = oldShape.Range.Duplicate;
            var oldStart = oldRange.Start;
            if (alignInline)
            {
                // InsertXML can silently import paragraph formatting from the
                // serialized Equation.DSMT4 fragment even when it does not split
                // the host paragraph. Preserve only paragraph-layout properties;
                // reapplying the paragraph Style to an OLE-containing range can
                // make Word/MathType regenerate the OLE presentation geometry.
                inlineParagraphLayout = CaptureMathTypeDisplayParagraphLayout(oldShape);
            }
            var sourceCount = document.InlineShapes.Count;
            // InsertXML can split the host paragraph for both display and inline
            // Equation.DSMT4 objects. Capture this before deleting the source so an
            // inline re-edit can restore surrounding prose to the same paragraph.
            sourceParagraphCount = ReadDocumentParagraphCount(document);
            if (!alignInline)
            {
                displayParagraphLayout = CaptureMathTypeDisplayParagraphLayout(oldShape);
                sourceWasNumbered = MathTypeOleInterop.TryReadDisplayNumberPosition(
                    oldShape,
                    out sourceNumberPosition);
                numberingLayoutChanged = sourceWasNumbered != session.Numbered
                    || sourceWasNumbered
                        && session.Numbered
                        && !string.Equals(
                            sourceNumberPosition,
                            session.MathTypeNumberPosition,
                            StringComparison.OrdinalIgnoreCase);
                if (sourceWasNumbered)
                    sourceNumberTemplate = ReadMathTypePlaceRefTemplateForShape(
                        document,
                        oldShape,
                        sourceNumberPosition);
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-replace-numbering source={sourceWasNumbered}:{sourceNumberPosition} target={session.Numbered}:{session.MathTypeNumberPosition} changed={numberingLayoutChanged}");
            }

            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=read-source-flat-opc formulaId={session.FormulaId}");
            var sourceFragment = MathTypeWordOpenXml.Read(oldShape);
            var originalProgId = sourceFragment.ProgId;
            var originalWordPosition = ReadInlineOleWordPosition(oldShape);
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=read-source-flat-opc-complete formulaId={session.FormulaId}");

            // Preserve the original Equation.DSMT4 CFB and replace only its MTEF.
            // Geometry is resolved from the rewritten MTEF below so a VisualTeX
            // re-edit uses the exact same MathType-native presentation model as
            // VisualTeX→MathType conversion and direct MathType insertion.
            var targetWidthPt = (float)Math.Max(
                1d,
                (session.ExportResult?.Width ?? 200d) * 0.75d);
            var targetHeightPt = (float)Math.Max(
                1d,
                (session.ExportResult?.Height ?? 60d) * 0.75d);
            var targetOriginalWidthPt = targetWidthPt;
            var targetOriginalHeightPt = targetHeightPt;
            var alignToWordTextBaseline = alignInline || session.Numbered;
            var targetWordPosition = alignToWordTextBaseline
                ? CalculateMathTypeOleWordPosition(
                    targetHeightPt,
                    session.ExportResult?.Height ?? 0f,
                    session.ExportResult?.Baseline)
                : 0;
            byte[] previewWmf;

            // Preserve the original Equation.DSMT4 CFB, replace only its MTEF
            // structure, and seed a fresh OLE presentation cache from VisualTeX's
            // current EMF. The source and result are serialized data; no MathType
            // COM server is needed for the semantic or visual update.
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=rewrite-cfb formulaId={session.FormulaId}");
            var rewritten = MathTypeOleStorage.RewriteMathTypeCompoundFile(
                sourceFragment.CompoundFile,
                mathMl,
                string.Equals(session.DisplayMode, "inline", StringComparison.Ordinal),
                session.FontSizePt);
            var expectedMathTypeSignature = MathTypeMtefCodec.SemanticSignature(mathMl);
            var generatedMathMl = MathTypeOleStorage.ReadMathMl(rewritten.CompoundFile);
            var generatedLatex = MathMlToLatexConverter.Convert(generatedMathMl);
            if (!MathTypeMathMlRoundTripMatches(expectedMathTypeSignature, generatedMathMl))
                throw new InvalidDataException(
                    $"VisualTeX generated invalid MathType MTEF. Expected '{metadata.Latex}', actual '{generatedLatex}'.");

            // MathType edit must not switch back to VisualTeX/MathJax geometry.
            // Render both the source and rewritten MTEF through the isolated
            // MathPage sidecar.  The source native extent tells us whether Word
            // was displaying the original equation at a user/document-specific
            // scale; apply that same scale to the rewritten native presentation.
            // This keeps MathType's own glyph/spacing model and prevents an inline
            // edit from changing the object baseline merely because VisualTeX's
            // frontend export has different pixel geometry.
            var renderRoot = !string.IsNullOrWhiteSpace(emfPath)
                ? Path.GetDirectoryName(emfPath!) ?? Path.GetTempPath()
                : Path.GetTempPath();
            var nativePreviewInputs = new Dictionary<string, byte[]>(StringComparer.Ordinal)
            {
                ["source"] = ReadMathTypeMtefFromCompoundFile(sourceFragment.CompoundFile),
                ["rewritten"] = rewritten.Mtef,
            };
            var renderedAllNativePreviews =
                MathTypeNativePreviewRenderer.TryRenderBatch(
                    nativePreviewInputs,
                    renderRoot,
                    out var nativePreviews);
            nativePreviews.TryGetValue("source", out sourceNativePreview);
            nativePreviews.TryGetValue("rewritten", out nativePreview);
            if (renderedAllNativePreviews
                && sourceNativePreview is not null
                && nativePreview is not null)
            {
                previewWmf = File.ReadAllBytes(nativePreview.WmfPath);
                float widthScale;
                float heightScale;
                if (alignInline)
                {
                    // Existing MathType objects in legacy/compatibility-mode Word
                    // documents are often stored at a document-specific presentation
                    // scale rather than at MathPage's raw native width/height.  An
                    // edit must preserve that horizontal/vertical scale; otherwise a
                    // previously compact inline object is silently reset to 1:1 and
                    // its Word object box grows relative to the surrounding prose.
                    widthScale = CalculateMathTypeNativePresentationScale(
                        sourceFragment.WidthPt,
                        sourceNativePreview.WidthPt);
                    heightScale = CalculateMathTypeNativePresentationScale(
                        sourceFragment.HeightPt,
                        sourceNativePreview.HeightPt);
                    targetWidthPt = Math.Max(1f, nativePreview.WidthPt * widthScale);
                    targetHeightPt = Math.Max(1f, nativePreview.HeightPt * heightScale);
                    // Genuine inline MathType objects keep w:dxaOrig/w:dyaOrig
                    // aligned with Word's displayed VML extent, including fractional
                    // point sizes.  Feeding MathPage's unscaled native extent here
                    // makes Word re-materialize a taller inline OLE and raises every
                    // OLE on the host line even when U+0001 Position stays unchanged.
                    targetOriginalWidthPt = targetWidthPt;
                    targetOriginalHeightPt = targetHeightPt;
                    targetWordPosition = originalWordPosition;
                }
                else
                {
                    widthScale = CalculateMathTypeNativePresentationScale(
                        sourceFragment.WidthPt,
                        sourceNativePreview.WidthPt);
                    heightScale = CalculateMathTypeNativePresentationScale(
                        sourceFragment.HeightPt,
                        sourceNativePreview.HeightPt);
                    targetWidthPt = Math.Max(1f, nativePreview.WidthPt * widthScale);
                    targetHeightPt = Math.Max(1f, nativePreview.HeightPt * heightScale);
                    targetWordPosition = alignToWordTextBaseline
                        ? (int)Math.Round(
                            nativePreview.WordPosition * heightScale,
                            MidpointRounding.AwayFromZero)
                        : 0;
                }
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-replace-native-preview formulaId={session.FormulaId} "
                    + $"sourceWord={sourceFragment.WidthPt:0.###}x{sourceFragment.HeightPt:0.###}@{originalWordPosition} "
                    + $"sourceNative={sourceNativePreview.WidthPt:0.###}x{sourceNativePreview.HeightPt:0.###}@{sourceNativePreview.WordPosition} "
                    + $"rewrittenNative={nativePreview.WidthPt:0.###}x{nativePreview.HeightPt:0.###}@{nativePreview.WordPosition} "
                    + $"scale={widthScale:0.###}x{heightScale:0.###} "
                    + $"target={targetWidthPt:0.###}x{targetHeightPt:0.###}@{targetWordPosition} "
                    + $"original={targetOriginalWidthPt:0.###}x{targetOriginalHeightPt:0.###}");
            }
            else
            {
                // Older MathType installations may provide only a 32-bit MathPage
                // library, and machines without MathPage have no native renderer at
                // all. Editing must still remain functional: use the already-rendered
                // VisualTeX EMF as a Word-owned presentation while preserving the
                // genuine MathType CFB/MTEF semantics. This is the same safe fallback
                // used by direct insertion and never activates the MathType server.
                if (string.IsNullOrWhiteSpace(emfPath) || !File.Exists(emfPath))
                    throw new InvalidOperationException(
                        "MathType native preview is unavailable, so VisualTeX cannot resize this MathType equation without a fallback vector preview.");
                previewWmf = MathTypeWordOpenXml.ConvertEnhancedMetafileToPlaceableWmf(
                    emfPath!,
                    targetWidthPt,
                    targetHeightPt);
                targetWordPosition = alignInline
                    // Without a native source/target preview pair there is no
                    // evidence for changing an already-correct Word baseline.
                    // Preserve the native source position rather than introducing
                    // a new VisualTeX/MathJax-derived offset on edit.
                    ? originalWordPosition
                    : alignToWordTextBaseline
                        ? CalculateMathTypeOleWordPosition(
                            targetHeightPt,
                            session.ExportResult?.Height ?? 0f,
                            session.ExportResult?.Baseline)
                        : 0;
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-replace-preview-fallback formulaId={session.FormulaId} "
                    + $"sourceNative={sourceNativePreview is not null} rewrittenNative={nativePreview is not null} "
                    + $"target={targetWidthPt:0.###}x{targetHeightPt:0.###}@{targetWordPosition}");
            }


            // Rewrite the existing Equation.DSMT4 CFB and its external Word WMF
            // presentation in one offline Flat OPC transaction using the native
            // MathType geometry whenever MathPage is available.
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=rewrite-flat-opc formulaId={session.FormulaId}");
            var replacementWordOpenXml = MathTypeWordOpenXml.RewriteWithPlaceableWmf(
                sourceFragment.WordOpenXml,
                rewritten.CompoundFile,
                previewWmf,
                targetWidthPt,
                targetHeightPt,
                targetOriginalWidthPt,
                targetOriginalHeightPt);
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=rewrite-flat-opc-complete formulaId={session.FormulaId}");
            var rewrittenFragment = MathTypeWordOpenXml.Read(replacementWordOpenXml);
            var rewrittenMathMl = MathTypeOleStorage.ReadMathMl(rewrittenFragment.CompoundFile);
            var rewrittenLatex = MathMlToLatexConverter.Convert(rewrittenMathMl);
            if (!MathTypeMathMlRoundTripMatches(expectedMathTypeSignature, rewrittenMathMl))
                throw new InvalidDataException(
                    $"VisualTeX generated invalid MathType Flat OPC. Expected '{metadata.Latex}', actual '{rewrittenLatex}'.");

            rollbackWordOpenXml = sourceFragment.WordOpenXml;
            rollbackStart = oldStart;

            // WordOpenXML exports above may terminate Word's custom Undo record.
            // Open it only after every source/template export and offline render,
            // immediately before the first document mutation.
            undoRecord = BeginUndoRecord("VisualTeX Update MathType OLE Formula")
                ?? throw new InvalidOperationException("Word could not establish an independent MathType edit transaction.");

            // An inline OLE object occupies exactly one Word character. Remove only
            // that character, then materialize the rewritten Flat OPC at the same
            // insertion point. Surrounding prose is untouched. If InsertXML fails,
            // the catch block restores the original serialized object.
            // Word's smart cut/paste deletes adjacent ordinary spaces together
            // with an inline OLE. Replacement is an exact object transaction, not
            // a user word deletion. Suspend whitespace adjustment through commit
            // or rollback and always restore the user's original option below.
            editOptions = _application.Options;
            previousSmartCutPaste = editOptions.SmartCutPaste;
            if (previousSmartCutPaste)
            {
                editOptions.SmartCutPaste = false;
                smartCutPasteSuspended = true;
            }
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=delete-source formulaId={session.FormulaId}");
            oldShape.Delete();
            oldDeleted = true;
            insertion = document.Range(oldStart, oldStart);
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=insert-rewritten-flat-opc formulaId={session.FormulaId}");
            insertion.InsertXML(replacementWordOpenXml);
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=insert-rewritten-flat-opc-complete formulaId={session.FormulaId}");

            if (document.InlineShapes.Count != sourceCount)
                throw new InvalidOperationException(
                    "Word changed the inline OLE object count while replacing the MathType equation.");

            // The replacement was just materialized at oldStart, so resolving it
            // must stay local to that mutation site. A one-character range is not
            // sufficient for Equation.DSMT4 in every Word field/view state; when
            // that narrow hint missed, FindMathTypeOleByRange previously fell back
            // to enumerating every InlineShape in the document and probing each
            // ProgID. In a 1000-MathType document that single lookup cost ~22 s.
            // Use the same bounded strategy as the mature direct-insert path and
            // fail/rollback rather than turning one-formula edit into O(N).
            var replacementRangeEnd = Math.Min(document.Content.End, oldStart + 8);
            replacement = FindMathTypeOleByRange(
                    document,
                    $"{RangeReferencePrefix}{oldStart}:{replacementRangeEnd}",
                    allowGlobalFallback: false)
                ?? FindMathTypeOleInParagraphAtPosition(document, oldStart)
                ?? FindMathTypeOleInLocalWindow(document, oldStart)
                ?? throw new InvalidOperationException(
                    "Word materialized the rewritten Flat OPC, but VisualTeX could not resolve the replacement MathType equation near the insertion point.");
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=resolve-replacement-complete formulaId={session.FormulaId}");

            if (sourceParagraphCount >= 0)
                RepairMathTypeInsertXmlParagraphSplit(
                    document,
                    replacement,
                    sourceParagraphCount);
            if (alignInline && inlineParagraphLayout is not null)
                RestoreMathTypeDisplayParagraphLayout(
                    replacement,
                    inlineParagraphLayout);

            MathTypeDisplayParagraphLayout? detachedNumberLayout = null;
            if (!alignInline)
            {
                // Older VisualTeX builds could misclassify a MathType MTPlaceRef
                // numbered row as inline and leave the equation and its number in
                // two adjacent paragraphs.  Repair that already-damaged shape on
                // the next edit as well: merge only a clean, immediately-following
                // MTPlaceRef-only paragraph and recover MathType's center/right tab
                // stops from that numbering paragraph.
                detachedNumberLayout = RepairDetachedMathTypeNumberParagraph(
                    document,
                    replacement);
            }
            var displayLayoutToRestore = detachedNumberLayout ?? displayParagraphLayout;
            if (!alignInline && numberingLayoutChanged)
            {
                var targetNumberTemplate = session.Numbered
                    ? ResolveMathTypeEditNumberTemplate(
                        document,
                        replacement,
                        sourceWasNumbered ? sourceNumberTemplate : null,
                        out createdEditSectionBreakCodeStart)
                    : null;
                RebuildMathTypeDisplayScaffold(
                    document,
                    replacement,
                    session.Numbered,
                    session.MathTypeNumberPosition,
                    targetNumberTemplate);
                if (displayLayoutToRestore is not null)
                    RestoreMathTypeDisplayParagraphLayout(
                        replacement,
                        displayLayoutToRestore);
                MathTypeEquationNumbering.UpdateEquationNumbers(document);
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-replace-numbering-rebuilt numbered={session.Numbered} position={session.MathTypeNumberPosition}");
            }
            else if (!alignInline && displayLayoutToRestore is not null)
            {
                RestoreMathTypeDisplayParagraphLayout(replacement, displayLayoutToRestore);
            }

            // Keep the U+0001 object-result baseline synchronized with the exact
            // same exported MathType geometry used to build the WMF. Reusing the
            // old object baseline after replacing the preview is what previously
            // let re-edited formulas acquire a second, incompatible layout model.
            SetInlineOleWordPosition(replacement, targetWordPosition);
            if (alignInline)
                RestoreTypingBaselineAfter(replacement);

            // Complete view/baseline changes inside the same Undo item. The final
            // materialized XML read remains mandatory but happens after that item
            // closes, so it cannot split content replacement from font writes.
            finalSelection = replacement.Range.Duplicate;
            RestoreViewState(document, viewState, finalSelection);
            editViewRestored = true;
            if (screenUpdatingSuspended)
            {
                _application.ScreenUpdating = previousScreenUpdating;
                screenUpdatingSuspended = false;
            }
            EndUndoRecord(undoRecord);
            editUndoEnded = true;
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=validate-replacement formulaId={session.FormulaId}");
            var replacementFragment = MathTypeWordOpenXml.Read(replacement);
            if (!string.IsNullOrWhiteSpace(originalProgId)
                && !string.Equals(
                    replacementFragment.ProgId,
                    originalProgId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    $"Word changed the MathType OLE ProgID from '{originalProgId}' to '{replacementFragment.ProgId}'.");
            var replacementMathMl = MathTypeOleStorage.ReadMathMl(replacementFragment.CompoundFile);
            var replacementLatex = MathMlToLatexConverter.Convert(replacementMathMl);
            if (!MathTypeMathMlRoundTripMatches(expectedMathTypeSignature, replacementMathMl))
                throw new InvalidDataException(
                    $"Word materialized the wrong MathType formula. Expected '{metadata.Latex}', actual '{replacementLatex}'.");

            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=complete formulaId={session.FormulaId}");
            return Result(session, document);
        }
        catch
        {
            editViewRestored = false;
            Release(finalSelection); finalSelection = null;
            if (oldDeleted && document is not null && rollbackStart >= 0
                && !string.IsNullOrWhiteSpace(rollbackWordOpenXml))
            {
                try
                {
                    TryDelete(replacement);
                    Range? rollbackRange = null;
                    InlineShape? rollbackShape = null;
                    try
                    {
                        rollbackRange = document.Range(rollbackStart, rollbackStart);
                        rollbackRange.InsertXML(rollbackWordOpenXml);
                        if (sourceParagraphCount >= 0)
                        {
                            var rollbackRangeEnd = Math.Min(
                                document.Content.End,
                                rollbackStart + 8);
                            rollbackShape = FindMathTypeOleByRange(
                                    document,
                                    $"{RangeReferencePrefix}{rollbackStart}:{rollbackRangeEnd}",
                                    allowGlobalFallback: false)
                                ?? FindMathTypeOleInParagraphAtPosition(document, rollbackStart)
                                ?? FindMathTypeOleInLocalWindow(document, rollbackStart);
                            if (rollbackShape is not null)
                            {
                                RepairMathTypeInsertXmlParagraphSplit(
                                    document,
                                    rollbackShape,
                                    sourceParagraphCount);
                                if (!alignInline && numberingLayoutChanged)
                                {
                                    RebuildMathTypeDisplayScaffold(
                                        document,
                                        rollbackShape,
                                        sourceWasNumbered,
                                        sourceNumberPosition,
                                        sourceWasNumbered ? sourceNumberTemplate : null);
                                    MathTypeEquationNumbering.UpdateEquationNumbers(document);
                                }
                                if (alignInline && inlineParagraphLayout is not null)
                                    RestoreMathTypeDisplayParagraphLayout(
                                        rollbackShape,
                                        inlineParagraphLayout);
                                if (displayParagraphLayout is not null)
                                    RestoreMathTypeDisplayParagraphLayout(
                                        rollbackShape,
                                        displayParagraphLayout);
                            }
                        }
                    }
                    finally
                    {
                        Release(rollbackShape);
                        Release(rollbackRange);
                    }
                }
                catch { }
                if (createdEditSectionBreakCodeStart >= 0)
                {
                    try
                    {
                        RemoveMathTypeSectionBreakFieldAtCodeStart(
                            document,
                            createdEditSectionBreakCodeStart);
                    }
                    catch { }
                }
            }
            else
            {
                TryDelete(replacement);
            }
            throw;
        }
        finally
        {
            nativePreview?.Dispose();
            sourceNativePreview?.Dispose();
            if (smartCutPasteSuspended && editOptions is not null)
            {
                try { editOptions.SmartCutPaste = previousSmartCutPaste; }
                catch (Exception optionError)
                {
                    WordDoubleClickHook.TraceMessage(
                        $"mathtype-edit-option-restore-failed error={optionError.Message}");
                }
            }
            Release(editOptions);
            if (screenUpdatingSuspended)
            {
                try { _application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
            if (!editViewRestored) RestoreViewState(document, viewState, finalSelection);
            if (!editUndoEnded) EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(finalSelection);
            Release(insertion);
            Release(oldRange);
            Release(replacement);
            Release(oldShape);
            Release(document);
        }
    }

    private static byte[] WaitForMathTypeOleMaterialization(
        InlineShape shape,
        string expectedLatex)
    {
        var delaysMs = new[] { 0, 15, 35, 70, 120, 200 };
        Exception? lastError = null;
        string? lastLatex = null;
        foreach (var delayMs in delaysMs)
        {
            if (delayMs > 0) Thread.Sleep(delayMs);
            try
            {
                var compound = MathTypeOleStorage.CaptureCompoundFile(shape);
                if (!MathTypeOleStorage.LooksLikeMathTypeCompoundFile(compound))
                {
                    lastError = new InvalidDataException(
                        "Word's replacement object is not yet a valid MathType compound storage.");
                    continue;
                }
                var mathMl = MathTypeOleStorage.ReadMathMl(compound);
                lastLatex = MathMlToLatexConverter.Convert(mathMl);
                if (MathTypeOleRoundTripMatches(expectedLatex, lastLatex))
                    return compound;
                lastError = new InvalidDataException(
                    $"MathType OLE materialization mismatch. Expected '{expectedLatex}', actual '{lastLatex}'.");
            }
            catch (Exception error)
            {
                lastError = error;
            }
        }

        throw new InvalidDataException(
            $"Word did not finish materializing the rewritten MathType OLE. Expected '{expectedLatex}', last='{lastLatex ?? "<unreadable>"}'.",
            lastError);
    }

    // Shared by redraw preflight and the actual insertion. Validate generated
    // bytes before the first source deletion; preview success alone does not
    // establish a lossless MathML -> Equation Native -> MathML round trip.
    internal static (MathTypeMtefCodec.RewriteResult Generated, byte[] CompoundFile,
        string SemanticSignature) PrepareStandaloneMathTypeOleData(
        string mathMl, bool inline, double fontSizePt, string latex)
    {
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(mathMl, inline, fontSizePt);
        var compoundFile = MathTypeOleStorage.CreateStandaloneCompoundFile(generated);
        var actualMathMl = MathTypeOleStorage.ReadMathMl(compoundFile);
        var expected = MathTypeMtefCodec.SemanticSignature(mathMl);
        if (!MathTypeMathMlRoundTripMatches(expected, actualMathMl))
            throw new InvalidDataException(
                $"VisualTeX generated invalid standalone MathType MTEF for '{latex}'. "
                + DescribeSemanticSignatureDifference(expected, MathTypeMtefCodec.SemanticSignature(actualMathMl)));
        return (generated, compoundFile, expected);
    }

    private static bool MathTypeMathMlRoundTripMatches(
        string expectedSignature,
        string actualMathMl) =>
        string.Equals(
            expectedSignature,
            MathTypeMtefCodec.SemanticSignature(actualMathMl),
            StringComparison.Ordinal);

    private static string DescribeSemanticSignatureDifference(
        string expected,
        string actual)
    {
        expected ??= string.Empty;
        actual ??= string.Empty;
        var common = Math.Min(expected.Length, actual.Length);
        var index = 0;
        while (index < common && expected[index] == actual[index]) index++;
        if (index == common && expected.Length == actual.Length)
            return $"identical(length={expected.Length})";

        static string DescribeAt(string value, int position)
        {
            if (position < 0 || position >= value.Length) return "<end>";
            var ch = value[position];
            return $"U+{(int)ch:X4}('{(char.IsControl(ch) ? '?' : ch)}')";
        }

        var start = Math.Max(0, index - 6);
        var expectedEnd = Math.Min(expected.Length, index + 7);
        var actualEnd = Math.Min(actual.Length, index + 7);
        static string Escape(string value) => string.Concat(value.Select(ch =>
            ch >= ' ' && ch <= '~' ? ch.ToString() : $"\\u{(int)ch:X4}"));
        return $"index={index}; expectedLen={expected.Length}; actualLen={actual.Length}; "
            + $"expectedChar={DescribeAt(expected, index)}; actualChar={DescribeAt(actual, index)}; "
            + $"expectedContext='{Escape(expected.Substring(start, expectedEnd - start))}'; "
            + $"actualContext='{Escape(actual.Substring(start, actualEnd - start))}'";
    }

    private static bool MathTypeOleRoundTripMatches(string expectedLatex, string actualLatex)
    {
        static string Normalize(string value) =>
            (value ?? string.Empty)
                .Replace(" ", string.Empty)
                .Replace("{", string.Empty)
                .Replace("}", string.Empty)
                .Trim();
        return string.Equals(
            Normalize(expectedLatex),
            Normalize(actualLatex),
            StringComparison.Ordinal);
    }

    private static InlineShape? FindNewMathTypeOleAtStart(
        Document document,
        int start,
        InlineShape sourceShape)
    {
        InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? range = null;
                var keepCandidate = false;
                try
                {
                    candidate = shapes[index];
                    range = candidate.Range;
                    if (range.Start != start
                        || IsSameComObject(candidate, sourceShape)
                        || !MathTypeOleInterop.IsMathTypeOle(candidate))
                        continue;
                    keepCandidate = true;
                    return candidate;
                }
                catch { }
                finally
                {
                    Release(range);
                    if (!keepCandidate) Release(candidate);
                }
            }
            return null;
        }
        finally { Release(shapes); }
    }

    private static bool IsSameComObject(object left, object right)
    {
        if (ReferenceEquals(left, right)) return true;
        if (!Marshal.IsComObject(left) || !Marshal.IsComObject(right)) return false;
        IntPtr leftIdentity = IntPtr.Zero;
        IntPtr rightIdentity = IntPtr.Zero;
        try
        {
            leftIdentity = Marshal.GetIUnknownForObject(left);
            rightIdentity = Marshal.GetIUnknownForObject(right);
            return leftIdentity == rightIdentity;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (leftIdentity != IntPtr.Zero) Marshal.Release(leftIdentity);
            if (rightIdentity != IntPtr.Zero) Marshal.Release(rightIdentity);
        }
    }

    public OfficeObjectResult ReplaceOle(
        OfficeSessionDocument session,
        string pngPath,
        string emfPath)
    {
        var sourceObjectMode =
            ResolveLegacyWrapperSourceMode(
                session);
        session.Mode = "edit";
        session.ObjectMode =
            FormulaOleContract.NativeOleMode;
        return ApplyOmmlVisualTeXHostSession(
            session,
            sourceObjectMode,
            mathMl: null,
            pngPath,
            emfPath);
    }

    public OfficeObjectResult ReplaceOmml(
        OfficeSessionDocument session,
        string mathMl)
    {
        var sourceObjectMode =
            ResolveLegacyWrapperSourceMode(
                session);
        session.Mode = "edit";
        session.ObjectMode =
            FormulaOleContract.WordOmmlMode;
        return ApplyOmmlVisualTeXHostSession(
            session,
            sourceObjectMode,
            mathMl,
            pngPath: null,
            emfPath: null);
    }

    public OfficeObjectResult Replace(OfficeSessionDocument session, string imagePath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        InlineShape? oldShape = null;
        Range? oldRange = null;
        Range? insertion = null;
        InlineShape? replacement = null;
        UndoRecord? undoRecord = null;
        WordViewState? viewState = null;
        Range? finalSelection = null;
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        try
        {
            undoRecord = BeginUndoRecord("VisualTeX Replace Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            viewState = CaptureViewState();
            try
            {
                previousScreenUpdating = _application.ScreenUpdating;
                _application.ScreenUpdating = false;
                screenUpdatingSuspended = true;
            }
            catch { }
            oldShape = ResolveCapturedFormulaPicture(
                    document,
                    session.SourceObjectId,
                    session.FormulaId)
                ?? throw new InvalidOperationException(
                    "The captured Word formula picture no longer exists.");
            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            var originalMetadata = WordFormulaMetadataReader.TryRead(oldShape)
                ?? session.OriginalMetadata;
            var editedSize = OfficeFormulaSizing.EditedSize(
                oldWidth,
                oldHeight,
                originalMetadata?.RenderWidthPx,
                originalMetadata?.RenderHeightPx,
                session.ExportResult?.Width ?? oldWidth / 0.75f,
                session.ExportResult?.Height ?? oldHeight / 0.75f,
                originalFontSizePt: originalMetadata?.FontSizePt,
                originalRenderFontSizePt: originalMetadata?.RenderFontSizePt);
            oldRange = oldShape.Range;
            insertion = oldRange.Duplicate;
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            object link = false;
            object save = true;
            object rangeObject = insertion;
            replacement = document.InlineShapes.AddPicture(
                imagePath,
                ref link,
                ref save,
                ref rangeObject);
            Configure(
                replacement,
                metadata,
                editedSize.Width,
                editedSize.Height,
                imagePath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            RemoveInlineBaselineSentinel(
                document,
                originalMetadata?.FormulaId ?? metadata.FormulaId);
            RemoveInlineOleTypingAnchorAfter(oldShape);
            oldShape.Delete();
            if (session.DisplayMode == "inline")
                RestoreTypingBaselineAfter(replacement);
            else
                TryReconcileShape(document, replacement, metadata);
            finalSelection = replacement.Range.Duplicate;
            return Result(session, document);
        }
        catch
        {
            TryDelete(replacement);
            throw;
        }
        finally
        {
            RestoreViewState(document, viewState, finalSelection);
            if (screenUpdatingSuspended)
            {
                try { _application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(finalSelection);
            Release(replacement);
            Release(insertion);
            Release(oldRange);
            Release(oldShape);
            Release(document);
        }
    }

    private static InlineShape? ResolveCapturedFormulaPicture(
        Document document,
        string? sourceObjectId,
        string formulaId)
    {
        Range? captured = null;
        InlineShapes? shapes = null;
        InlineShape? match = null;
        try
        {
            captured =
                WordFormulaOperationLocator.ResolveCapturedRange(
                    document,
                    sourceObjectId);
            shapes = captured.InlineShapes;
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (WordFormulaMetadataReader.IsNativeOle(
                            candidate)
                        || MathTypeOleInterop.IsMathTypeOle(
                            candidate))
                        continue;

                    var metadata =
                        WordFormulaMetadataReader
                            .TryReadCachedPreview(
                                candidate);
                    if (!string.Equals(
                            metadata?.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (match is not null)
                        throw new InvalidDataException(
                            "The captured Word range contains more than one formula picture with the same identity.");

                    match = candidate;
                    candidate = null;
                }
                finally
                {
                    Release(candidate);
                }
            }

            var result = match;
            match = null;
            return result;
        }
        finally
        {
            Release(match);
            Release(shapes);
            Release(captured);
        }
    }

    private static InlineShape AddOleObject(Document document, Range range)
    {
        var insertionStart = range.Start;
        Range? originalRightCharacter = null;
        try
        {
            // Keep one live boundary, not an inventory of document.InlineShapes.
            // On a collapsed insertion the original right-hand character must
            // remain AFTER anything newly inserted. Its live End distinguishes a
            // newly materialized OLE from an old object at the same numeric offset.
            // Failure to capture the witness disables recovery, not insertion.
            if (range.End == insertionStart)
            {
                try
                {
                    originalRightCharacter = range.Duplicate;
                    originalRightCharacter.SetRange(insertionStart, checked(insertionStart + 1));
                }
                catch (COMException)
                {
                    Release(originalRightCharacter);
                    originalRightCharacter = null;
                }
            }
            // Successful insertion is the overwhelmingly common path. Do not
            // enumerate document.InlineShapes merely to prepare evidence for the
            // rare wdErrorCommandFailed recovery case; that turns a 100-formula
            // batch into repeated whole-document COM work. The recovery path below
            // proves ownership from the exact EMBED field boundary instead.
            return document.InlineShapes.AddOLEObject(
                ClassType: FormulaOleContract.ProgId,
                LinkToFile: false,
                DisplayAsIcon: false,
                Range: range);
        }
        catch (COMException error) when (
            error.HResult == unchecked((int)0x800A1066))
        {
            // Word can finish embedding the OLE storage and then report
            // wdErrorCommandFailed after the server's creation verb. Recover only
            // a VisualTeX EMBED field that starts at this exact insertion point;
            // otherwise preserve the original failure.
            var recovered = TryRecoverMaterializedOleAfterCommandFailure(
                document,
                insertionStart,
                originalRightCharacter,
                out var recoveryEvidence);
            if (recovered is null)
                throw new COMException(
                    $"{error.Message} VisualTeX OLE recovery evidence: {recoveryEvidence}",
                    error.HResult);
            WordDoubleClickHook.TraceMessage(
                $"ole-add-command-failed-but-materialized start={insertionStart}");
            return recovered;
        }
        finally { Release(originalRightCharacter); }
    }

    private static InlineShape? TryRecoverMaterializedOleAfterCommandFailure(
        Document document,
        int insertionStart,
        Range? originalRightCharacter,
        out string evidence)
    {
        Range? probe = null;
        InlineShapes? shapes = null;
        InlineShape? candidate = null;
        Range? candidateRange = null;
        OLEFormat? format = null;
        try
        {
            if (originalRightCharacter is null
                || originalRightCharacter.StoryType != WdStoryType.wdMainTextStory)
            {
                evidence = $"start={insertionStart}; new-insertion-boundary=unproven";
                return null;
            }
            var newInsertionEnd = originalRightCharacter.End - 1;
            if (newInsertionEnd <= insertionStart)
            {
                evidence = $"start={insertionStart}; original-right-character={newInsertionEnd}; no-new-insertion";
                return null;
            }
            var content = document.Content;
            try
            {
                var probeStart = Math.Max(content.Start, insertionStart);
                var probeEnd = Math.Min(Math.Min(content.End, newInsertionEnd), insertionStart + 128);
                if (probeEnd <= probeStart)
                    probeEnd = Math.Min(content.End, probeStart + 1);
                probe = document.Range(probeStart, probeEnd);
            }
            finally { Release(content); }
            shapes = probe.InlineShapes;
            evidence = $"start={insertionStart}; localShapes={shapes.Count}";
            for (var index = shapes.Count; index >= 1; index--)
            {
                candidate = shapes[index];
                candidateRange = candidate.Range;
                // A successful AddOLEObject normally returns the result Range at
                // the insertion coordinate (or after the display TAB). When the
                // OLE server finishes storage creation and Word subsequently
                // raises wdErrorCommandFailed, however, Word can expose only the
                // one-character result of its EMBED field. The result starts after
                // the hidden field code (about 30 characters), not at the field's
                // opening coordinate. Accept that shape only when the surrounding
                // wdFieldEmbed is independently proven to start at this exact
                // insertion point and names the VisualTeX ProgID.
                if (candidateRange.Start < insertionStart
                    || candidateRange.End > newInsertionEnd
                    || (candidateRange.Start != insertionStart
                        && candidateRange.Start != insertionStart + 1
                        && !IsVisualTeXOleEmbedFieldAtInsertion(
                            document,
                            insertionStart,
                            candidateRange)))
                {
                    Release(candidateRange); candidateRange = null;
                    Release(candidate); candidate = null;
                    continue;
                }
                format = candidate.OLEFormat;
                if (!string.Equals(
                        format.ProgID,
                        FormulaOleContract.ProgId,
                        StringComparison.OrdinalIgnoreCase))
                {
                    Release(format); format = null;
                    Release(candidateRange); candidateRange = null;
                    Release(candidate); candidate = null;
                    continue;
                }

                var recovered = candidate;
                candidate = null;
                return recovered;
            }
            var coordinates = new List<string>(shapes.Count);
            for (var index = 1; index <= shapes.Count; index++)
            {
                candidate = shapes[index];
                candidateRange = candidate.Range;
                coordinates.Add($"{index}:{candidateRange.Start}-{candidateRange.End}");
                Release(candidateRange); candidateRange = null;
                Release(candidate); candidate = null;
            }
            evidence += "; ranges=" + string.Join(",", coordinates);
            return null;
        }
        catch (COMException)
        {
            evidence = $"start={insertionStart}; localInventory=COM-failed";
            return null;
        }
        finally
        {
            Release(format);
            Release(candidateRange);
            Release(candidate);
            Release(shapes);
            Release(probe);
        }
    }

    private static bool IsVisualTeXOleEmbedFieldAtInsertion(
        Document document,
        int insertionStart,
        Range candidateRange)
    {
        // Bound the recovery to Word's short EMBED field-code displacement. A
        // remotely placed object must never be adopted merely because its ProgID
        // matches.
        if (candidateRange.Start < insertionStart
            || candidateRange.Start - insertionStart > 128)
            return false;

        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        try
        {
            fields = document.Fields;
            for (var index = fields.Count; index >= 1; index--)
            {
                field = fields[index];
                if (field.Type != WdFieldType.wdFieldEmbed)
                {
                    Release(field); field = null;
                    continue;
                }
                code = field.Code;
                result = field.Result;
                var codeStartsAtInsertion = code.Start == insertionStart
                    || code.Start == insertionStart + 1;
                var ownsCandidate = result.Start <= candidateRange.Start
                    && result.End >= candidateRange.End;
                var namesVisualTeX = (code.Text ?? string.Empty).IndexOf(
                        FormulaOleContract.ProgId,
                        StringComparison.OrdinalIgnoreCase) >= 0;
                if (codeStartsAtInsertion && ownsCandidate && namesVisualTeX)
                    return true;

                Release(result); result = null;
                Release(code); code = null;
                Release(field); field = null;
            }
            return false;
        }
        catch (COMException)
        {
            return false;
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static void InitializeOle(
        InlineShape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        IVisualTeXFormulaObject? formula = null;
        var initialized = false;
        try
        {
            StoreWordDisplayPreviewMetrics(metadata, emfPath);
            format = shape.OLEFormat;
            oleObject = WordOleObjectAccessor.GetRunningObject(format);
            formula = oleObject as IVisualTeXFormulaObject;
            if (formula is null)
                throw new InvalidOperationException(
                    "The inserted Word object does not expose the VisualTeX native OLE interface.");
            FormulaOleInterop.Initialize(formula, metadata, emfPath, pngPath);
            WordFormulaMetadataReader.CacheMetadata(shape, metadata);
            initialized = true;
        }
        finally
        {
            if (formula is not null)
            {
                try { FormulaOleInterop.CloseAfterSave(formula); }
                catch when (!initialized) { }
            }
            Release(oleObject);
            Release(format);
        }
    }

    private static bool TryUpdateOle(
        InlineShape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        IVisualTeXFormulaObject? formula = null;
        var updated = false;
        try
        {
            StoreWordDisplayPreviewMetrics(metadata, emfPath);
            try { format = shape.OLEFormat; }
            catch { return false; }
            try { oleObject = WordOleObjectAccessor.GetRunningObject(format); }
            catch { return false; }
            formula = oleObject as IVisualTeXFormulaObject;
            if (formula is null) return false;
            FormulaOleInterop.Update(formula, metadata, emfPath, pngPath);
            WordFormulaMetadataReader.CacheMetadata(shape, metadata);
            updated = true;
            return true;
        }
        finally
        {
            if (formula is not null)
            {
                try { FormulaOleInterop.CloseAfterSave(formula); }
                catch when (!updated) { }
            }
            Release(oleObject);
            Release(format);
        }
    }

    private static void StoreWordDisplayPreviewMetrics(
        FormulaMetadata metadata,
        string emfPath)
    {
        if (!string.Equals(
                metadata.DisplayMode,
                "block",
                StringComparison.OrdinalIgnoreCase))
        {
            metadata.WordDisplayPreviewInkHeightRatio = null;
            metadata.WordDisplayPreviewBottomWhitespaceRatio = null;
            return;
        }

        var metrics = OfficeOlePreview.TryMeasureVerticalInkMetrics(emfPath);
        if (!metrics.HasValue)
        {
            metadata.WordDisplayPreviewInkHeightRatio = null;
            metadata.WordDisplayPreviewBottomWhitespaceRatio = null;
            return;
        }
        metadata.WordDisplayPreviewInkHeightRatio =
            metrics.Value.InkHeightRatio;
        metadata.WordDisplayPreviewBottomWhitespaceRatio =
            metrics.Value.BottomWhitespaceRatio;
        metadata.Validate();
    }

    private Range DuplicateCurrentSelectionRange()
    {
        Selection? selection = null;
        Range? range = null;
        try
        {
            selection = _application.Selection;
            range = selection.Range;
            return range.Duplicate;
        }
        finally
        {
            Release(range);
            Release(selection);
        }
    }

    internal WordViewState CaptureFormulaFormatConversionViewState() =>
        CaptureViewState();

    internal void RestoreFormulaFormatConversionViewState(WordViewState? state)
    {
        if (state is null) return;
        Document? document = null;
        try
        {
            document = _application.ActiveDocument;
            RestoreViewState(document, state, preferredSelection: null);
        }
        catch
        {
            // View restoration must never turn an otherwise successful formula
            // conversion into an error. The original document is reactivated by
            // the conversion write path before this method normally runs.
        }
        finally { Release(document); }
    }

    private WordViewState CaptureViewState()
    {
        Selection? selection = null;
        Range? range = null;
        Window? window = null;
        try
        {
            selection = _application.Selection;
            range = selection.Range;
            try { window = _application.ActiveWindow; } catch { }
            int? vertical = null;
            int? horizontal = null;
            try { vertical = window?.VerticalPercentScrolled; } catch { }
            try { horizontal = window?.HorizontalPercentScrolled; } catch { }
            return new WordViewState
            {
                SelectionStart = range.Start,
                SelectionEnd = range.End,
                VerticalPercentScrolled = vertical,
                HorizontalPercentScrolled = horizontal,
            };
        }
        finally
        {
            Release(window);
            Release(range);
            Release(selection);
        }
    }

    private void RestoreViewState(
        Document? document,
        WordViewState? state,
        Range? preferredSelection)
    {
        if (document is null || state is null) return;
        Selection? selection = null;
        Range? fallback = null;
        Range? content = null;
        Window? window = null;
        try
        {
            selection = _application.Selection;
            if (preferredSelection is not null)
            {
                if (selection.Start != preferredSelection.Start
                    || selection.End != preferredSelection.End)
                    selection.SetRange(preferredSelection.Start, preferredSelection.End);
            }
            else
            {
                content = document.Content;
                var start = Math.Max(content.Start, Math.Min(state.SelectionStart, content.End));
                var end = Math.Max(start, Math.Min(state.SelectionEnd, content.End));
                object startValue = start;
                object endValue = end;
                fallback = document.Range(ref startValue, ref endValue);
                selection.SetRange(fallback.Start, fallback.End);
            }
            try { window = _application.ActiveWindow; } catch { }
            if (window is not null)
            {
                try
                {
                    if (state.HorizontalPercentScrolled.HasValue
                        && window.HorizontalPercentScrolled != state.HorizontalPercentScrolled.Value)
                        window.HorizontalPercentScrolled = state.HorizontalPercentScrolled.Value;
                }
                catch { }
                try
                {
                    if (state.VerticalPercentScrolled.HasValue
                        && window.VerticalPercentScrolled != state.VerticalPercentScrolled.Value)
                        window.VerticalPercentScrolled = state.VerticalPercentScrolled.Value;
                }
                catch { }
            }
        }
        catch { }
        finally
        {
            Release(window);
            Release(content);
            Release(fallback);
            Release(selection);
        }
    }

    private static InlineShape? FindMathTypeOleByRange(
        Document document,
        string? sourceObjectId,
        bool allowGlobalFallback = true)
    {
        if (!TryParseRangeReference(sourceObjectId, out var start, out var end))
            return null;

        Range? hintedRange = null;
        InlineShapes? hintedShapes = null;
        Range? content = null;
        Range? localRange = null;
        InlineShapes? localShapes = null;
        InlineShapes? shapes = null;
        try
        {
            try
            {
                hintedRange = document.Range(start, end);
                hintedShapes = hintedRange.InlineShapes;
                for (var index = 1; index <= hintedShapes.Count; index++)
                {
                    InlineShape? candidate = null;
                    try
                    {
                        candidate = hintedShapes[index];
                        if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                        var result = candidate;
                        candidate = null;
                        return result;
                    }
                    finally { Release(candidate); }
                }
            }
            catch
            {
                // Fall back to a document scan below. The source range can move
                // if surrounding prose changes while the VisualTeX editor is open.
            }

            // During whole-document format conversion targets are processed from
            // the end toward the start, so the captured MathType range remains a
            // reliable local locator. After the source has been deleted the exact
            // range is expected to contain no MathType object. Probe only a tiny
            // neighborhood for Word's occasional one/two-character range drift;
            // never enumerate the whole document merely to prove absence.
            try
            {
                content = document.Content;
                var localStart = Math.Max(content.Start, start - 2);
                var localEnd = Math.Min(content.End, Math.Max(end, start) + 2);
                if (localEnd > localStart)
                {
                    localRange = document.Range(localStart, localEnd);
                    localShapes = localRange.InlineShapes;
                    InlineShape? nearestLocal = null;
                    var nearestLocalDistance = int.MaxValue;
                    try
                    {
                        for (var index = 1; index <= localShapes.Count; index++)
                        {
                            InlineShape? candidate = null;
                            Range? candidateRange = null;
                            try
                            {
                                candidate = localShapes[index];
                                if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                                candidateRange = candidate.Range;
                                var distance = Math.Abs(candidateRange.Start - start);
                                if (distance >= nearestLocalDistance) continue;
                                Release(nearestLocal);
                                nearestLocal = candidate;
                                candidate = null;
                                nearestLocalDistance = distance;
                            }
                            finally
                            {
                                Release(candidateRange);
                                Release(candidate);
                            }
                        }
                        if (nearestLocal is not null && nearestLocalDistance <= 2)
                        {
                            var result = nearestLocal;
                            nearestLocal = null;
                            return result;
                        }
                    }
                    finally { Release(nearestLocal); }
                }
            }
            catch { }
            finally
            {
                Release(localShapes);
                localShapes = null;
                Release(localRange);
                localRange = null;
                Release(content);
                content = null;
            }

            if (!allowGlobalFallback) return null;

            shapes = document.InlineShapes;
            InlineShape? nearest = null;
            var nearestDistance = int.MaxValue;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    candidateRange = candidate.Range;
                    if (candidateRange.Start == start && candidateRange.End == end)
                    {
                        Release(nearest);
                        nearest = null;
                        var exact = candidate;
                        candidate = null;
                        return exact;
                    }
                    var distance = Math.Abs(candidateRange.Start - start);
                    if (distance >= nearestDistance) continue;
                    Release(nearest);
                    nearest = candidate;
                    candidate = null;
                    nearestDistance = distance;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }

            // A range is only a transient locator for third-party OLE. Avoid
            // accidentally editing a different MathType equation after large
            // document changes: permit only a very small local shift.
            if (nearest is not null && nearestDistance <= 2)
            {
                var result = nearest;
                nearest = null;
                return result;
            }
            Release(nearest);
            return null;
        }
        finally
        {
            Release(shapes);
            Release(localShapes);
            Release(localRange);
            Release(content);
            Release(hintedShapes);
            Release(hintedRange);
        }
    }

    private UndoRecord? BeginUndoRecord(string name)
    {
        // Mature MathType insertion helpers can be invoked as the external
        // target writer of one rebuilt OMML/VisualTeX transaction. In that
        // case they must borrow the exact outer record and must never start or
        // end a second Word Custom UndoRecord.
        if (WordFormulaMutationTransaction.IsActiveOnCurrentThread)
            return null;

        UndoRecord? undoRecord = null;
        try
        {
            undoRecord = _application.UndoRecord;
            WordDoubleClickHook.TraceMessage(
                $"word-undo-begin name={name} recording={undoRecord.IsRecordingCustomRecord} level={undoRecord.CustomRecordLevel}");
            // Word exposes one process-wide Custom Undo Record stack. Starting a
            // second record from a helper that is already running inside a parent
            // transaction lets the inner EndCustomRecord close/reshape the outer
            // transaction. Format conversion deliberately wraps each destructive
            // source replacement in an outer record, while mature insertion paths
            // such as InsertMathTypeOle also call this helper. Treat an active
            // record as borrowed: the inner operation participates in it but must
            // neither start nor end another record.
            if (undoRecord.IsRecordingCustomRecord || undoRecord.CustomRecordLevel > 0)
            {
                Release(undoRecord);
                return null;
            }
            undoRecord.StartCustomRecord(name);
            WordDoubleClickHook.TraceMessage(
                $"word-undo-started name={name} recording={undoRecord.IsRecordingCustomRecord} level={undoRecord.CustomRecordLevel}");
            return undoRecord;
        }
        catch (Exception error)
        {
            Release(undoRecord);
            throw new InvalidOperationException($"Word could not start the undo record '{name}'.", error);
        }
    }

    private static void EndUndoRecord(UndoRecord? undoRecord)
    {
        if (undoRecord is null) return;
        WordDoubleClickHook.TraceMessage(
            $"word-undo-end recording={undoRecord.IsRecordingCustomRecord} level={undoRecord.CustomRecordLevel}");
        undoRecord.EndCustomRecord();
    }

    private void TraceConversionUndoState(string stage)
    {
        UndoRecord? record = null;
        Document? active = null;
        try
        {
            record = _application.UndoRecord;
            active = _application.ActiveDocument;
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-undo-state stage={stage} document={active.Name} recording={record.IsRecordingCustomRecord} level={record.CustomRecordLevel}");
        }
        finally { Release(active); Release(record); }
    }

    private static void TryReconcileShape(
        Document document,
        InlineShape shape,
        FormulaMetadata metadata,
        bool numberingOrderMayHaveChanged = true,
        bool reuseExistingNumberedTableFormatting = false,
        Table? knownNumberedTable = null,
        bool numberingScaffoldOnly = false)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            if (string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
            {
                RemoveInlineBaselineSentinel(document, metadata.FormulaId);
                ResetDisplayFormulaPosition(range);
            }
            if (numberingScaffoldOnly && metadata.Numbered)
            {
                WordEquationNumbering.BuildFormulaNumberingScaffoldForConversion(
                    document,
                    range,
                    shape.Height,
                    metadata,
                    knownNumberedTable);
            }
            else
            {
                WordEquationNumbering.ReconcileFormula(
                    document,
                    range,
                    shape.Height,
                    metadata,
                    numberingOrderMayHaveChanged,
                    reuseExistingNumberedTableFormatting,
                    knownNumberedTable);
            }
        }
        finally { Release(range); }
    }

    private static bool NumberingOrderMayHaveChanged(
        FormulaMetadata? originalMetadata,
        FormulaMetadata metadata)
    {
        if (originalMetadata is null) return true;
        return originalMetadata.Numbered != metadata.Numbered
            || !string.Equals(
                originalMetadata.DisplayMode,
                metadata.DisplayMode,
                StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryReadStoredWordInlineOleSize(
        FormulaMetadata? metadata,
        out float width,
        out float height)
    {
        width = 0;
        height = 0;
        if (metadata?.WordInlineOleWidthPt is not > 0
            || metadata.WordInlineOleHeightPt is not > 0)
            return false;
        width = (float)metadata.WordInlineOleWidthPt.Value;
        height = (float)metadata.WordInlineOleHeightPt.Value;
        return width > 0
            && height > 0
            && !float.IsNaN(width)
            && !float.IsInfinity(width)
            && !float.IsNaN(height)
            && !float.IsInfinity(height);
    }

    private static void StoreWordInlineOleSize(
        FormulaMetadata metadata,
        float width,
        float height,
        bool inline)
    {
        if (!inline
            || width <= 0
            || height <= 0
            || float.IsNaN(width)
            || float.IsInfinity(width)
            || float.IsNaN(height)
            || float.IsInfinity(height))
        {
            metadata.WordInlineOleWidthPt = null;
            metadata.WordInlineOleHeightPt = null;
            metadata.WordInlineOlePositionPt = null;
            return;
        }
        metadata.WordInlineOleWidthPt = width;
        metadata.WordInlineOleHeightPt = height;
    }

    private static void CacheFinalWordInlineOleGeometry(
        InlineShape shape,
        FormulaMetadata metadata,
        bool inline)
    {
        if (!inline) return;
        StoreWordInlineOleSize(metadata, shape.Width, shape.Height, inline: true);
        metadata.WordInlineOlePositionPt = ReadInlineOleWordPosition(shape);
        metadata.Validate();
        // Host geometry is Word-specific and can differ slightly from the OLE
        // server's requested extent after Font.Position materializes the EMBED
        // result. Keep that final Word state in the Word-side metadata cache so a
        // later baseline repair can restore the exact last-known-good extent.
        WordFormulaMetadataReader.CacheMetadata(shape, metadata);
    }

    private static bool IsNumberedBlockOmml(FormulaMetadata metadata) =>
        metadata.Numbered
        && string.Equals(
            metadata.DisplayMode,
            "block",
            StringComparison.OrdinalIgnoreCase);

    private static bool TryNormalizeDefinedWordFontSize(
        float value,
        out float fontSizePt)
    {
        fontSizePt = FormulaFontSize.DefaultPt;
        // Word uses 9,999,999 for mixed/undefined Font.Size. Fall back to the
        // persisted semantic size instead of treating that sentinel as a real value.
        if (value <= 0f
            || value >= 1000f
            || float.IsNaN(value)
            || float.IsInfinity(value))
            return false;
        fontSizePt = FormulaFontSize.Normalize(value);
        return true;
    }

    // Word exposes one native OpenType MATH face per document. This is intentionally
    // document-scoped: changing it reflows every existing OMath in the document; it is
    // not a per-formula font override and must never be emulated with m:nor/run splitting.
    private static void ApplyDocumentOmmlMathFont(
        Document document,
        FormulaMetadata metadata)
    {
        var requested = ResolveDocumentOmmlMathFont(metadata.FormulaLetterFont);
        if (string.Equals(
                requested,
                WordOfficeMathFontLoader.LatinModernMathFamily,
                StringComparison.OrdinalIgnoreCase))
            WordOfficeMathFontLoader.EnsureLoaded();

        string current;
        try { current = document.OMathFontName ?? string.Empty; }
        catch (COMException error)
        {
            throw new InvalidOperationException(
                "Word could not read this document's native Office Math font setting.",
                error);
        }
        if (string.Equals(current, requested, StringComparison.OrdinalIgnoreCase))
            return;

        try { document.OMathFontName = requested; }
        catch (COMException error)
        {
            throw new InvalidOperationException(
                $"Word could not set this document's native Office Math font to '{requested}'.",
                error);
        }

        string applied;
        try { applied = document.OMathFontName ?? string.Empty; }
        catch (COMException error)
        {
            throw new InvalidOperationException(
                "Word could not verify the document-level Office Math font after applying it.",
                error);
        }
        if (!string.Equals(applied, requested, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Word rejected '{requested}' as the document's native Office Math font. "
                + "Install a valid OpenType math font or repair the VisualTeX Office integration.");
        }
    }

    private static string ResolveDocumentOmmlMathFont(string? preference)
    {
        // Word native OMath has one document-wide math font and requires an
        // OpenType MATH table. VisualTeX's text-face choices (Times, Palatino,
        // Helvetica/Arial) cannot safely replace that math font, so they fall back
        // to the bundled Latin Modern Math instead of flattening individual runs.
        return (preference ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "cambria" => "Cambria Math",
            "stix" => "STIX Two Math",
            _ => WordOfficeMathFontLoader.LatinModernMathFamily,
        };
    }

    private static string ApplyOmmlTypographyXml(
        string omml,
        double fontSizePt,
        FormulaMetadata metadata)
    {
        if (string.IsNullOrWhiteSpace(omml))
            throw new InvalidDataException("VisualTeX produced an empty OMML payload.");
        var equation = XElement.Parse(
            omml,
            LoadOptions.PreserveWhitespace);
        XNamespace word =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace math =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        ApplyBulkOmmlTypographyXml(
            equation,
            fontSizePt,
            metadata,
            word,
            math);
        return equation.ToString(SaveOptions.DisableFormatting);
    }

    private static string ResolveOmmlChineseFont(string? preference) =>
        (preference ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "songti" => "SimSun",
            "kaiti" => "KaiTi",
            "heiti" => "SimHei",
            _ => "Microsoft YaHei",
        };

    private static bool IsChineseOmmlCodePoint(int codePoint) =>
        codePoint is >= 0x2E80 and <= 0x2FFF
        || codePoint is >= 0x3000 and <= 0x303F
        || codePoint is >= 0x31C0 and <= 0x31EF
        || codePoint is >= 0x3400 and <= 0x4DBF
        || codePoint is >= 0x4E00 and <= 0x9FFF
        || codePoint is >= 0xF900 and <= 0xFAFF
        || codePoint is >= 0xFF00 and <= 0xFFEF
        || codePoint is >= 0x20000 and <= 0x2FA1F;

    private static bool HasReusableInlineOmmlEditHost(
        Range formulaRange)
    {
        Tables? tables = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        Fields? fields = null;
        InlineShapes? shapes = null;
        try
        {
            tables = formulaRange.Tables;
            if (tables.Count > 0) return false;
            maths = formulaRange.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathInline) return false;
            mathRange = math.Range.Duplicate;
            if (mathRange.Start != formulaRange.Start
                || mathRange.End != formulaRange.End)
                return false;
            fields = formulaRange.Fields;
            shapes = formulaRange.InlineShapes;
            return fields.Count == 0 && shapes.Count == 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(shapes);
            Release(fields);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(tables);
        }
    }

    private static bool HasReusableContentOnlyDisplayOmmlEditHost(
        Range formulaRange)
    {
        Tables? tables = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        Fields? fields = null;
        InlineShapes? shapes = null;
        try
        {
            // Same-state unnumbered Display edits replace only this OMath. A
            // user's cell is also a valid unchanged host: routing it through
            // numbering reconciliation needlessly replaces paragraph runs and
            // then compares the new content against a provisional fingerprint.
            // Keep the content-only placeholder transaction and final full-tree
            // validation shared with body Display edits. Only the table case
            // needs a local physical-cell check; the large body fast path keeps
            // avoiding the layout-triggering Information(wdWithInTable) query.
            tables = formulaRange.Tables;
            if (tables.Count > 0 && !WordFormulaHost.IsWithinSingleCell(formulaRange))
                return false;
            maths = formulaRange.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay) return false;
            mathRange = math.Range.Duplicate;
            if (mathRange.Start != formulaRange.Start
                || mathRange.End != formulaRange.End)
                return false;
            fields = formulaRange.Fields;
            shapes = formulaRange.InlineShapes;
            return fields.Count == 0 && shapes.Count == 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(shapes);
            Release(fields);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(tables);
        }
    }

    private static bool HasReusableStandaloneDisplayOmmlHost(
        Document document,
        Range formulaRange)
    {
        Tables? tables = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefix = null;
        Range? suffix = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            tables = formulaRange.Tables;
            if (tables.Count > 0) return false;
            maths = formulaRange.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay) return false;
            mathRange = math.Range.Duplicate;
            if (mathRange.Start != formulaRange.Start
                || mathRange.End != formulaRange.End)
                return false;
            fields = formulaRange.Fields;
            shapes = formulaRange.InlineShapes;
            if (fields.Count != 0 || shapes.Count != 0) return false;

            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (formulaRange.Start < paragraphRange.Start
                || formulaRange.End > paragraphRange.End)
                return false;
            prefix = document.Range(paragraphRange.Start, formulaRange.Start);
            suffix = document.Range(formulaRange.End, paragraphRange.End);
            return !ContainsVisibleBodyText(prefix.Text)
                && !ContainsVisibleBodyText(suffix.Text);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(suffix);
            Release(prefix);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(tables);
        }
    }

    private static void Configure(
        InlineShape shape,
        FormulaMetadata metadata,
        float maxWidth,
        float maxHeight,
        string imagePath,
        float exportedHeight,
        float? exportedBaseline,
        bool alignInline,
        bool nativeOleKnown = false,
        bool trustExportDimensions = false)
    {
        // maxWidth/maxHeight are already the SVG's physical size after the
        // 96 dpi CSS-pixel to 72 dpi Word-point conversion. A 12 pt minimum
        // width scales narrow inline formulas (notably x) far above their
        // semantic font size. Keep only a one-point safety floor.
        var width = Math.Max(1f, maxWidth);
        var height = Math.Max(1f, maxHeight);
        if (!trustExportDimensions)
        {
            using var image = Image.FromFile(imagePath);
            var ratio = image.Width / (float)Math.Max(1, image.Height);
            height = width / ratio;
            if (maxHeight > 0 && height > maxHeight)
            {
                height = maxHeight;
                width = height * ratio;
            }
        }
        // An OLE object is initially created with the placeholder preview's 4:1
        // aspect ratio. Setting only Width while aspect-ratio locking is enabled
        // therefore distorts the real formula. Apply both natural dimensions
        // explicitly, then lock the resolved ratio for later user resizing.
        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
        shape.Width = width;
        shape.Height = height;
        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
        if (!nativeOleKnown && !WordFormulaMetadataReader.IsNativeOle(shape))
        {
            var encoded = FormulaMetadataCodec.Encode(metadata);
            shape.Title = encoded;
            shape.AlternativeText = encoded;
        }
        if (alignInline)
            ApplyInlineBaseline(
                shape,
                shape.Height,
                exportedHeight,
                exportedBaseline,
                FormulaFontSize.ResolveSemanticFontSize(metadata));
        else
            ResetDisplayFormulaPosition(shape);
    }

    private static bool ShouldAlignInline(InlineShape shape, FormulaMetadata metadata)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            return ShouldAlignInline(range, metadata);
        }
        finally { Release(range); }
    }

    private static bool ShouldAlignInline(Range formulaRange, FormulaMetadata metadata)
    {
        if (string.Equals(metadata.DisplayMode, "inline", StringComparison.Ordinal))
            return true;
        return HasVisibleSurroundingText(formulaRange);
    }

    private static bool HasVisibleFollowingInlineText(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? after = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var paragraphBodyEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (formulaRange.End >= paragraphBodyEnd) return false;
            after = paragraphRange.Duplicate;
            after.SetRange(formulaRange.End, paragraphBodyEnd);
            return ContainsVisibleBodyText(after.Text);
        }
        catch { return false; }
        finally
        {
            Release(after);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool HasVisibleSurroundingText(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? before = null;
        Range? after = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (formulaRange.Start > paragraphRange.Start)
            {
                before = paragraphRange.Duplicate;
                before.SetRange(paragraphRange.Start, formulaRange.Start);
                if (ContainsVisibleBodyText(before.Text)) return true;
            }
            if (formulaRange.End < paragraphRange.End)
            {
                after = paragraphRange.Duplicate;
                after.SetRange(formulaRange.End, paragraphRange.End);
                if (ContainsVisibleBodyText(after.Text)) return true;
            }
            return false;
        }
        catch { return false; }
        finally
        {
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool ContainsVisibleBodyText(string? value)
    {
        if (string.IsNullOrEmpty(value)) return false;
        foreach (var character in value!)
        {
            if (character is '\r' or '\n' or '\t' or '\v' or '\a' or '\u0001' or '\u200B' or '\u200C')
                continue;
            if (!char.IsWhiteSpace(character)) return true;
        }
        return false;
    }

    private readonly struct InlineOlePreviewMetrics
    {
        internal InlineOlePreviewMetrics(float inkHeightRatio, float bottomWhitespaceRatio)
        {
            InkHeightRatio = inkHeightRatio;
            BottomWhitespaceRatio = bottomWhitespaceRatio;
        }

        internal float InkHeightRatio { get; }
        internal float BottomWhitespaceRatio { get; }
    }

    private static float? ReadDefinedShapeFontPosition(InlineShape shape)
    {
        Range? range = null;
        Range? probe = null;
        Document? document = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = shape.Range;
            document = range.Document;

            // An embedded OLE is an EMBED field, not a single character range.
            // Current VisualTeX and genuine MathType objects keep the field
            // instruction on the prose baseline and position only the U+0001
            // object/result character. Read that character first so resizing a
            // legacy formula scales its real baseline instead of treating the
            // mixed field range as wdUndefined.
            for (var position = range.Start; position < range.End; position++)
            {
                Release(font);
                font = null;
                Release(probe);
                probe = document.Range(position, position + 1);
                if (!string.Equals(probe.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                font = probe.Font;
                var objectPosition = font.Position;
                if (objectPosition != (int)WdConstants.wdUndefined
                    && objectPosition >= -256
                    && objectPosition <= 256)
                    return objectPosition;
            }

            Release(font);
            font = null;
            font = range.Font;
            var fallbackPosition = font.Position;
            return fallbackPosition == (int)WdConstants.wdUndefined
                || fallbackPosition < -256
                || fallbackPosition > 256
                    ? null
                    : fallbackPosition;
        }
        catch { return null; }
        finally
        {
            Release(font);
            Release(probe);
            Release(range);
            Release(document);
        }
    }

    private static int ResolveInlineShapeInsertionIndex(Document document, int position)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? range = null;
        try
        {
            shapes = document.InlineShapes;
            var preceding = 0;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(range);
                range = null;
                Release(shape);
                shape = shapes[index];
                range = shape.Range;
                if (range.Start >= position) break;
                preceding++;
            }
            return preceding + 1;
        }
        finally
        {
            Release(range);
            Release(shape);
            Release(shapes);
        }
    }

    private static InlineShape? FindMathTypeOleAtIndex(Document document, int index)
    {
        InlineShapes? shapes = null;
        InlineShape? candidate = null;
        try
        {
            shapes = document.InlineShapes;
            if (index < 1 || index > shapes.Count) return null;
            candidate = shapes[index];
            if (!MathTypeOleInterop.IsMathTypeOle(candidate)) return null;
            var result = candidate;
            candidate = null;
            return result;
        }
        catch { return null; }
        finally
        {
            Release(candidate);
            Release(shapes);
        }
    }

    private static InlineShape? FindMathTypeOleInParagraphAtPosition(
        Document document,
        int position)
    {
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? shapes = null;
        try
        {
            var safePosition = Math.Max(
                document.Content.Start,
                Math.Min(document.Content.End - 1, position));
            probe = document.Range(safePosition, safePosition);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            shapes = paragraphRange.InlineShapes;
            InlineShape? match = null;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    if (match is not null)
                    {
                        Release(match);
                        return null;
                    }
                    match = candidate;
                    candidate = null;
                }
                finally { Release(candidate); }
            }
            return match;
        }
        catch { return null; }
        finally
        {
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    private static InlineShape? FindMathTypeOleInLocalWindow(
        Document document,
        int position)
    {
        Range? window = null;
        InlineShapes? shapes = null;
        InlineShape? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = document.Content.End;
            var start = Math.Max(contentStart, position - 8);
            var end = Math.Min(contentEnd, position + 256);
            if (end <= start) end = Math.Min(contentEnd, start + 1);
            window = document.Range(start, end);
            shapes = window.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    candidateRange = candidate.Range;
                    var distance = Math.Abs(candidateRange.Start - position);
                    if (distance >= bestDistance) continue;
                    Release(best);
                    best = candidate;
                    candidate = null;
                    bestDistance = distance;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }
            return best;
        }
        catch
        {
            Release(best);
            return null;
        }
        finally
        {
            Release(shapes);
            Release(window);
        }
    }

    private static InlineShape? FindMathTypeOleNearPosition(
        Document document,
        int position)
    {
        InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            InlineShape? best = null;
            var bestDistance = int.MaxValue;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? range = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    range = candidate.Range;
                    var distance = Math.Abs(range.Start - position);
                    if (distance > 3 || distance >= bestDistance) continue;
                    Release(best);
                    best = candidate;
                    candidate = null;
                    bestDistance = distance;
                }
                catch { }
                finally
                {
                    Release(range);
                    Release(candidate);
                }
            }
            return best;
        }
        finally { Release(shapes); }
    }

    private static void RelocateCompleteMathTypePlaceRefFromRightToLeft(
        Document document,
        InlineShape shape)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Field? sourcePlaceRef = null;
        Range? sourceCode = null;
        Range? sourceResult = null;
        Range? sourceSpan = null;
        Range? formattedField = null;
        Range? destination = null;
        Field? leftPlaceRef = null;
        Field? rightPlaceRef = null;
        Range? rightCode = null;
        Range? rightResult = null;
        Range? removal = null;
        Range? separator = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display equation does not occupy one paragraph while relocating its number.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            sourcePlaceRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: false)
                ?? throw new InvalidOperationException(
                    "The temporary right-numbered MathType row has no MTPlaceRef field to relocate.");
            if (!TryReadCompleteMathTypePlaceRefTemplate(
                    document,
                    sourcePlaceRef,
                    out _))
                throw new InvalidOperationException(
                    "The temporary right-numbered MathType row has an incomplete MTPlaceRef field tree.");

            sourceCode = sourcePlaceRef.Code;
            sourceResult = sourcePlaceRef.Result;
            var bodyEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            var sourceStart = Math.Max(paragraphRange.Start, sourceCode.Start - 1);
            var sourceEnd = Math.Min(bodyEnd, sourceResult.End + 1);
            if (sourceEnd <= sourceStart)
                throw new InvalidOperationException(
                    "Word did not expose a complete MTPlaceRef field span for relocation.");

            // Copy the already-complete outer field as one formatted Word range.
            // This preserves the nested field delimiters atomically; no SEQ field is
            // ever created while MTPlaceRef is half-built in the target position.
            sourceSpan = document.Range(sourceStart, sourceEnd);
            formattedField = sourceSpan.FormattedText;
            destination = document.Range(paragraphRange.Start, paragraphRange.Start);
            destination.FormattedText = formattedField;
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                && string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE"),
                    "left-relocate-after-copy",
                    StringComparison.Ordinal))
                throw new COMException(
                    "Injected MathType left-number relocation failure for rollback acceptance.",
                    unchecked((int)0x800A1710));

            Release(sourceResult);
            sourceResult = null;
            Release(sourceCode);
            sourceCode = null;
            Release(sourcePlaceRef);
            sourcePlaceRef = null;
            Release(shapeRange);
            shapeRange = shape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;

            leftPlaceRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: true)
                ?? throw new InvalidOperationException(
                    "Word did not preserve the relocated MathType number on the left of the OLE.");
            if (!TryReadCompleteMathTypePlaceRefTemplate(
                    document,
                    leftPlaceRef,
                    out _))
                throw new InvalidOperationException(
                    "Word damaged the MTPlaceRef field tree while copying it to the left of the OLE.");

            rightPlaceRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: false)
                ?? throw new InvalidOperationException(
                    "Word lost the original temporary right-side MTPlaceRef before relocation completed.");
            rightCode = rightPlaceRef.Code;
            rightResult = rightPlaceRef.Result;
            bodyEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            var rightFieldStart = Math.Max(paragraphRange.Start, rightCode.Start - 1);
            var removalStart = Math.Max(shapeRange.End, rightFieldStart - 1);
            var removalEnd = Math.Min(bodyEnd, rightResult.End + 1);
            if (removalEnd <= removalStart)
                throw new InvalidOperationException(
                    "Word did not expose the temporary right-side MathType number for removal.");
            removal = document.Range(removalStart, removalEnd);
            if ((removal.Text ?? string.Empty).IndexOf('\t') < 0)
                throw new InvalidOperationException(
                    "The temporary right-side MathType number is missing its separator tab.");
            removal.Delete();

            Release(rightResult);
            rightResult = null;
            Release(rightCode);
            rightCode = null;
            Release(rightPlaceRef);
            rightPlaceRef = null;
            Release(leftPlaceRef);
            leftPlaceRef = null;
            Release(shapeRange);
            shapeRange = shape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;

            leftPlaceRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: true)
                ?? throw new InvalidOperationException(
                    "The final left-numbered MathType row has no MTPlaceRef field.");
            if (!TryReadCompleteMathTypePlaceRefTemplate(
                    document,
                    leftPlaceRef,
                    out _))
                throw new InvalidOperationException(
                    "The final left-numbered MathType row has an incomplete MTPlaceRef field tree.");
            rightPlaceRef = FindMathTypePlaceRefFieldForShape(
                paragraphRange,
                shapeRange,
                numberOnLeft: false);
            if (rightPlaceRef is not null)
                throw new InvalidOperationException(
                    "The temporary right-side MTPlaceRef survived left-number relocation.");

            Range? finalCode = null;
            try
            {
                finalCode = leftPlaceRef.Code;
                var finalCodeText = finalCode.Text ?? string.Empty;
                if (finalCodeText.Length == 0
                    || char.IsWhiteSpace(finalCodeText[finalCodeText.Length - 1]))
                    throw new InvalidOperationException(
                        "The relocated MathType equation number contains trailing whitespace.");
            }
            finally { Release(finalCode); }

            var separatorStart = Math.Max(paragraphRange.Start, shapeRange.Start - 1);
            separator = document.Range(separatorStart, shapeRange.Start);
            if (!string.Equals(separator.Text, "\t", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "The relocated left MathType number is not separated from its OLE by one tab.");
        }
        finally
        {
            Release(separator);
            Release(removal);
            Release(rightResult);
            Release(rightCode);
            Release(rightPlaceRef);
            Release(leftPlaceRef);
            Release(destination);
            Release(formattedField);
            Release(sourceSpan);
            Release(sourceResult);
            Release(sourceCode);
            Release(sourcePlaceRef);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static void BuildIndependentMathTypeDisplayScaffold(
        Document document,
        InlineShape shape,
        bool numbered,
        string mathTypeNumberPosition,
        MathTypeWordOpenXml.NumberTemplate? numberTemplate)
    {
        Range? shapeRange = null;
        Range? separator = null;
        Field? placeRef = null;
        try
        {
            var numberOnLeft = numbered && string.Equals(
                mathTypeNumberPosition,
                "left",
                StringComparison.OrdinalIgnoreCase);
            var numberOnRight = numbered && string.Equals(
                mathTypeNumberPosition,
                "right",
                StringComparison.OrdinalIgnoreCase);
            if (numbered && !numberOnLeft && !numberOnRight)
                throw new InvalidDataException(
                    "MathType equation number position must be left or right.");
            if (numbered && numberTemplate is null)
                throw new InvalidDataException(
                    "A numbered MathType display equation requires an MTPlaceRef template.");

            shapeRange = shape.Range;
            if (!numberOnLeft)
            {
                // MathType's display style centers the equation on its center tab.
                separator = document.Range(shapeRange.Start, shapeRange.Start);
                separator.Text = "\t";
                Release(separator);
                separator = null;
                Release(shapeRange);
                shapeRange = shape.Range;
            }

            if (!numbered) return;

            if (numberOnLeft)
            {
                // Do not create MTPlaceRef first and then insert a tab at its end.
                // Word can expand the field's boundary to absorb that character,
                // leaving no ordinary separator between the number and the OLE.
                // Materialize the tab first, then create MTPlaceRef before it so
                // the final structure is unambiguously FIELD + TAB + OLE.
                var numberInsertionPosition = shapeRange.Start;
                separator = document.Range(
                    numberInsertionPosition,
                    numberInsertionPosition);
                separator.Text = "\t";
                Release(separator);
                separator = null;
                placeRef = CreateIndependentMathTypePlaceRef(
                    document,
                    numberInsertionPosition,
                    numberTemplate!);
            }
            else
            {
                separator = document.Range(shapeRange.End, shapeRange.End);
                separator.Text = "\t";
                var placeRefPosition = separator.End;
                placeRef = CreateIndependentMathTypePlaceRef(
                    document,
                    placeRefPosition,
                    numberTemplate!);
            }
            try { placeRef.ShowCodes = false; } catch { }
        }
        finally
        {
            Release(placeRef);
            Release(separator);
            Release(shapeRange);
        }
    }

    private static MathTypeWordOpenXml.NumberTemplate ReadMathTypePlaceRefTemplateForShape(
        Document document,
        InlineShape shape,
        string numberPosition)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Field? placeRef = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display equation does not occupy one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var numberOnLeft = string.Equals(
                numberPosition,
                "left",
                StringComparison.OrdinalIgnoreCase);
            placeRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft)
                ?? throw new InvalidOperationException(
                    "The numbered MathType equation has no readable MTPlaceRef field.");
            return ReadMathTypePlaceRefTemplate(document, placeRef);
        }
        finally
        {
            Release(placeRef);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static MathTypeWordOpenXml.NumberTemplate ResolveMathTypeEditNumberTemplate(
        Document document,
        InlineShape shape,
        MathTypeWordOpenXml.NumberTemplate? sourceTemplate,
        out int createdSectionBreakCodeStart)
    {
        createdSectionBreakCodeStart = -1;
        if (sourceTemplate is not null)
            return sourceTemplate;

        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Field? nearest = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display equation does not occupy one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            nearest = FindNearestMathTypePlaceRefField(
                document,
                shapeRange.Start,
                paragraphRange.Start,
                paragraphRange.End);
            var documentNumberFormat = EquationNumberFormat.Resolve(
                WordEquationNumbering.GetEquationNumberFormatId(document));
            var template = nearest is not null
                ? ReadMathTypePlaceRefTemplate(document, nearest)
                : MathTypeWordOpenXml.CreateVisualTeXNumberTemplate(
                    documentNumberFormat.Id);

            if (documentNumberFormat.UsesHeading
                && MathTypeNumberTemplateUsesHeading(template))
            {
                EnsureMathTypeHeadingScopeState(
                    document,
                    shapeRange.Start,
                    documentNumberFormat,
                    out createdSectionBreakCodeStart);
            }
            return template;
        }
        finally
        {
            Release(nearest);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static void RebuildMathTypeDisplayScaffold(
        Document document,
        InlineShape shape,
        bool numbered,
        string numberPosition,
        MathTypeWordOpenXml.NumberTemplate? numberTemplate,
        bool updateNumberFields = true)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? paragraphShapes = null;
        Field? placeRef = null;
        Range? placeRefCode = null;
        Range? placeRefResult = null;
        Range? fieldSpan = null;
        Range? prefix = null;
        Range? suffix = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "MathType numbering can only be changed for a standalone display equation.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            paragraphShapes = paragraphRange.InlineShapes;
            if (paragraphShapes.Count != 1)
                throw new InvalidOperationException(
                    "MathType numbering cannot be changed in a paragraph containing other inline objects.");

            // Remove only the native MathType number field. Nested SEQ fields live
            // inside its code range and disappear with the outer MACROBUTTON.
            placeRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: true)
                ?? FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: false);
            if (placeRef is not null)
            {
                placeRefCode = placeRef.Code;
                placeRefResult = placeRef.Result;
                var fieldStart = Math.Max(
                    paragraphRange.Start,
                    placeRefCode.Start - 1);
                var fieldEnd = Math.Min(
                    Math.Max(paragraphRange.Start, paragraphRange.End - 1),
                    placeRefResult.End + 1);
                fieldSpan = document.Range(fieldStart, Math.Max(fieldStart, fieldEnd));
                fieldSpan.Delete();
            }

            Release(shapeRange);
            shapeRange = shape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            var bodyEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            prefix = document.Range(paragraphRange.Start, shapeRange.Start);
            suffix = document.Range(shapeRange.End, bodyEnd);
            if (!IsMathTypeDisplayScaffoldWhitespace(prefix.Text)
                || !IsMathTypeDisplayScaffoldWhitespace(suffix.Text))
                throw new InvalidOperationException(
                    "MathType numbering was not changed because the display paragraph contains user text outside the equation.");

            // Delete from the end first so the OLE range remains stable while its
            // surrounding tabs are normalized back to a bare equation object.
            if (suffix.End > suffix.Start) suffix.Delete();
            Release(shapeRange);
            shapeRange = shape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            Release(prefix);
            prefix = document.Range(paragraphRange.Start, shapeRange.Start);
            if (prefix.End > prefix.Start) prefix.Delete();

            BuildIndependentMathTypeDisplayScaffold(
                document,
                shape,
                numbered,
                numberPosition,
                numberTemplate);
            ConfigureNewMathTypeDisplayEquation(
                document,
                shape,
                numbered,
                numberPosition,
                updateNestedNumberFields: updateNumberFields);
        }
        finally
        {
            Release(suffix);
            Release(prefix);
            Release(fieldSpan);
            Release(placeRefResult);
            Release(placeRefCode);
            Release(placeRef);
            Release(paragraphShapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static bool IsMathTypeDisplayScaffoldWhitespace(string? text)
    {
        if (string.IsNullOrEmpty(text)) return true;
        foreach (var character in text)
        {
            if (char.IsWhiteSpace(character)
                || character is '\t' or '\r' or '\n' or '\v'
                    or '\u0013' or '\u0014' or '\u0015')
                continue;
            return false;
        }
        return true;
    }

    internal static Field CreateIndependentMathTypePlaceRef(
        Document document,
        int position,
        MathTypeWordOpenXml.NumberTemplate template)
    {
        if (template.Segments.Count == 0)
            throw new InvalidDataException("MathType MTPlaceRef numbering template is empty.");

        Range? insertion = null;
        Range? outerCode = null;
        Fields? outerNestedFields = null;
        Field? outer = null;
        Field? nested = null;
        try
        {
            insertion = document.Range(position, position);
            outer = document.Fields.Add(
                insertion,
                WdFieldType.wdFieldMacroButton,
                "MTPlaceRef",
                false);
            // Word only treats Fields.Add inside another field's Code range as a
            // genuinely nested field while the outer field is exposing its code.
            // If MTPlaceRef remains in result view, the exact same numeric Range is
            // silently promoted to a sibling document field. That is the direct
            // cause of the left-number '(.)' + escaped SEQ corruption.
            outer.ShowCodes = true;

            // Word itself creates MACROBUTTON MTPlaceRef with one trailing ASCII
            // space in its native field instruction. Keep that native instruction
            // untouched. The previous implementation appended a temporary marker
            // and deleted it after building the nested fields; in the real in-process
            // VSTO path Word can refuse or silently ignore that final Range.Delete,
            // leaking the synthetic tail token or throwing 0x800A1710.
            //
            // Every supported MathType number template ends with literal closing
            // punctuation (currently ')'). Materialize that real final literal first
            // at Code.End. Its original insertion coordinate immediately becomes a
            // strictly interior point of MTPlaceRef. Build all preceding template
            // segments in reverse at that fixed point. No synthetic character and no
            // post-construction deletion are needed, so the visible number naturally
            // ends at its real closing punctuation with zero trailing whitespace.
            outerCode = outer.Code;
            var nativeCodeText = outerCode.Text ?? string.Empty;
            if (nativeCodeText.Length == 0
                || !char.IsWhiteSpace(nativeCodeText[nativeCodeText.Length - 1]))
                throw new InvalidOperationException(
                    "Word did not create MTPlaceRef with its native instruction separator.");

            var terminalSegmentIndex = template.Segments.Count - 1;
            var terminalSegment = template.Segments[terminalSegmentIndex];
            if (terminalSegment.IsField)
                throw new InvalidDataException(
                    "MathType MTPlaceRef numbering template must end with literal closing punctuation.");
            var terminalText = terminalSegment.Value.TrimEnd();
            if (terminalText.Length == 0)
                throw new InvalidDataException(
                    "MathType MTPlaceRef numbering template has no terminal closing punctuation.");

            var insertionPosition = outerCode.End;
            Release(insertion);
            insertion = document.Range(insertionPosition, insertionPosition);
            insertion.InsertAfter(terminalText);
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                && string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE"),
                    "left-scaffold-after-marker",
                    StringComparison.Ordinal))
                throw new COMException(
                    "Injected MathType left scaffold Range failure for rollback acceptance.",
                    unchecked((int)0x800A1710));

            Release(outerCode);
            outerCode = outer.Code;
            if (insertionPosition <= outerCode.Start || insertionPosition >= outerCode.End)
                throw new InvalidOperationException(
                    "Word did not retain the MathType terminal punctuation inside MTPlaceRef.");

            for (var segmentIndex = terminalSegmentIndex - 1;
                 segmentIndex >= 0;
                 segmentIndex--)
            {
                var segment = template.Segments[segmentIndex];
                Release(insertion);
                insertion = document.Range(insertionPosition, insertionPosition);
                if (!segment.IsField)
                {
                    var text = segment.Value;
                    // Word's native MTPlaceRef already supplies the instruction
                    // separator before the template. Avoid duplicating the leading
                    // whitespace stored in VisualTeX/MathType templates; trailing
                    // whitespace remains meaningful only between template tokens.
                    if (segmentIndex == 0)
                        text = text.TrimStart();
                    if (!string.IsNullOrEmpty(text))
                        insertion.InsertAfter(text);
                    continue;
                }

                Release(nested);
                Release(outerNestedFields);
                Release(outerCode);
                outerCode = outer.Code;
                outerNestedFields = outerCode.Fields;
                nested = outerNestedFields.Add(
                    insertion,
                    WdFieldType.wdFieldEmpty,
                    segment.Value.Trim(),
                    false);
                // Word automatically leaves a hidden SEQ ... \\h field with its
                // code displayed. In that state the parent Code.Text substitutes
                // the (empty) child result and the hidden increment disappears from
                // the outer instruction stream. MathType's native MTPlaceRef keeps
                // every child in result view, which also makes the complete outer
                // code serializable/reusable as one field tree.
                try { nested.ShowCodes = false; } catch { }
            }

            Release(outerCode);
            outerCode = outer.Code;
            var completedCodeText = outerCode.Text ?? string.Empty;
            if (completedCodeText.Length == 0
                || char.IsWhiteSpace(completedCodeText[completedCodeText.Length - 1]))
                throw new InvalidOperationException(
                    "Word left trailing whitespace after the completed MathType equation number.");

            if (!TryReadCompleteMathTypePlaceRefTemplate(document, outer, out _))
                throw new InvalidOperationException(
                    "Word detached one or more nested MathType number fields while constructing MTPlaceRef.");
            try { outer.ShowCodes = false; } catch { }
            var result = outer;
            outer = null;
            return result;
        }
        finally
        {
            Release(nested);
            Release(outerNestedFields);
            Release(outerCode);
            Release(insertion);
            Release(outer);
        }
    }

    private static void ConfigureNewMathTypeDisplayEquation(
        Document document,
        InlineShape shape,
        bool numbered,
        string mathTypeNumberPosition,
        bool updateNestedNumberFields = true,
        WordCharacterFormatting? bodyFormatting = null,
        MathTypeDisplayParagraphLayout? preservedDisplayParagraphLayout = null,
        float? knownDisplayColumnWidth = null)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        TabStops? tabs = null;
        TabStop? tab = null;
        Field? placeRef = null;
        Range? shapePrefix = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "A new MathType display equation must occupy one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            bodyFormatting ??= WordCharacterFormatting.CaptureParagraphMark(paragraphRange);
            EnsureMathTypeNativeStyles(document, bodyFormatting);
            object displayStyle = "MTDisplayEquation";
            // Clear direct formatting inherited from the source/blank paragraph
            // first, then apply MathType's native paragraph style. Paragraph.Reset
            // can itself restore Normal, so applying MTDisplayEquation before Reset
            // silently lost the native style in real Word.
            paragraph.Reset();
            paragraphRange.set_Style(ref displayStyle);
            NormalizeMathTypeDisplayParagraphFormat(
                document,
                paragraph,
                paragraphRange,
                preservedDisplayParagraphLayout,
                knownDisplayColumnWidth);
            format = paragraph.Format;
            tabs = format.TabStops;
            var hasCenter = false;
            var hasRight = false;
            for (var index = 1; index <= tabs.Count; index++)
            {
                Release(tab);
                tab = tabs[index];
                hasCenter |= tab.Alignment == WdTabAlignment.wdAlignTabCenter;
                hasRight |= tab.Alignment == WdTabAlignment.wdAlignTabRight;
            }
            if (!hasCenter || !hasRight)
                throw new InvalidOperationException(
                    "MTDisplayEquation does not contain MathType's center/right tab stops.");
            try { paragraphRange.ListFormat.RemoveNumbers(); } catch { }

            // Do not infer the OLE boundary from Paragraph.Range.Text. When Word's
            // field-code view is enabled, an embedded equation is exposed as
            // FIELD-BEGIN + " EMBED Equation..." rather than U+0001 even though
            // InlineShape.Range still identifies the exact OLE story interval.
            // Validate the real characters before that interval instead; this is
            // invariant under Alt+F9 / Field.ShowCodes.
            shapePrefix = document.Range(paragraphRange.Start, shapeRange.Start);
            if (!numbered)
            {
                if (!string.Equals(shapePrefix.Text, "\t", StringComparison.Ordinal))
                {
                    // Word can discard the Flat OPC leading tab when a VisualTeX
                    // display paragraph is replaced in-place. Rebuild the native
                    // MathType center-tab scaffold from the live OLE boundary
                    // instead of accepting an uncentered paragraph or failing the
                    // whole format conversion.
                    if (!IsMathTypeDisplayScaffoldWhitespace(shapePrefix.Text))
                        throw new InvalidOperationException(
                            "The unnumbered MathType display equation contains user text before its OLE object.");
                    shapePrefix.SetRange(shapeRange.Start, shapeRange.Start);
                    shapePrefix.Text = "\t";
                    Release(shapeRange);
                    shapeRange = shape.Range;
                    Release(paragraphRange);
                    paragraphRange = paragraph.Range;
                    Release(shapePrefix);
                    shapePrefix = document.Range(paragraphRange.Start, shapeRange.Start);
                }
                if (!string.Equals(shapePrefix.Text, "\t", StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "The unnumbered MathType display equation does not begin with Word's native center tab before its OLE object.");
                return;
            }

            var numberOnLeft = string.Equals(
                mathTypeNumberPosition,
                "left",
                StringComparison.OrdinalIgnoreCase);
            var numberOnRight = string.Equals(
                mathTypeNumberPosition,
                "right",
                StringComparison.OrdinalIgnoreCase);
            if (!numberOnLeft && !numberOnRight)
                throw new InvalidDataException(
                    "MathType equation number position must be left or right.");

            placeRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft)
                ?? throw new InvalidOperationException(
                    "The numbered MathType display equation has no MTPlaceRef field on the requested side of its OLE object.");
            if (!TryReadCompleteMathTypePlaceRefTemplate(
                    document,
                    placeRef,
                    out _))
                throw new InvalidOperationException(
                    "The numbered MathType display equation has a detached or incomplete MTPlaceRef field tree.");
            Range? placeRefCode = null;
            Range? placeRefResult = null;
            Range? separator = null;
            try
            {
                placeRefCode = placeRef.Code;
                placeRefResult = placeRef.Result;
                var fieldStart = Math.Max(paragraphRange.Start, placeRefCode.Start - 1);
                var fieldEnd = Math.Min(paragraphRange.End, placeRefResult.End + 1);
                if (numberOnLeft)
                {
                    if (fieldEnd > shapeRange.Start)
                        throw new InvalidOperationException(
                            "The MathType left equation number is not positioned before the equation object.");
                    // An empty-result MTPlaceRef can report a field-end boundary
                    // that coincides with the following OLE field. The one story
                    // character immediately before InlineShape.Range is still the
                    // actual separator TAB in both result and field-code views.
                    Release(shapePrefix);
                    shapePrefix = document.Range(
                        Math.Max(paragraphRange.Start, shapeRange.Start - 1),
                        shapeRange.Start);
                    if (!string.Equals(shapePrefix.Text, "\t", StringComparison.Ordinal))
                        throw new InvalidOperationException(
                            "The MathType left-numbered display equation has no tab between its number and equation.");
                }
                else
                {
                    if (fieldStart < shapeRange.End)
                        throw new InvalidOperationException(
                            "The MathType right equation number is not positioned after the equation object.");
                    separator = document.Range(shapeRange.End, fieldStart);
                    if ((separator.Text ?? string.Empty).IndexOf('\t') < 0)
                        throw new InvalidOperationException(
                            "The MathType right-numbered display equation has no tab between its equation and number.");
                    if (!string.Equals(shapePrefix.Text, "\t", StringComparison.Ordinal))
                    {
                        var prefix = shapePrefix.Text ?? string.Empty;
                        var codes = string.Join(
                            ",",
                            prefix.Select(ch => $"U+{(int)ch:X4}"));
                        throw new InvalidOperationException(
                            $"The MathType right-numbered display equation does not begin with Word's native center tab before its OLE object. paragraph={paragraphRange.Start}-{paragraphRange.End}; shape={shapeRange.Start}-{shapeRange.End}; prefix={codes}.");
                    }
                }
                if (updateNestedNumberFields)
                    UpdateNestedMathTypeNumberFields(placeRef);
            }
            finally
            {
                Release(separator);
                Release(placeRefResult);
                Release(placeRefCode);
            }
        }
        finally
        {
            Release(shapePrefix);
            Release(placeRef);
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static string StripWordFieldControlCharacters(string text) =>
        (text ?? string.Empty)
            .Replace("\u0013", string.Empty)
            .Replace("\u0014", string.Empty)
            .Replace("\u0015", string.Empty)
            .Replace("\u0001", string.Empty)
            .Replace("\uFFFC", string.Empty)
            .Replace("\r", string.Empty);

    private static MathTypeWordOpenXml.NumberTemplate ReadMathTypePlaceRefTemplate(
        Document document,
        Field source)
    {
        _ = document;
        Range? sourceCode = null;
        try
        {
            sourceCode = source.Code;
            var sourceText = sourceCode.Text ?? string.Empty;
            var macroIndex = sourceText.IndexOf(
                "MTPlaceRef",
                StringComparison.OrdinalIgnoreCase);
            if (macroIndex < 0)
                throw new InvalidDataException(
                    "The source MathType numbering field is not MTPlaceRef.");

            // Word exposes nested fields inside Field.Code using the same field
            // control characters stored in the document: U+0013 begin, optional
            // U+0014 separate, and U+0015 end. Parse that stream directly so
            // literal MathType punctuation such as '(', '.', '-' and ')' is kept
            // byte-for-byte instead of inferred from COM Range coordinates.
            var pattern = sourceText.Substring(
                macroIndex + "MTPlaceRef".Length);
            var template = new MathTypeWordOpenXml.NumberTemplate();
            var literal = new StringBuilder();
            for (var index = 0; index < pattern.Length;)
            {
                if (pattern[index] != '\u0013')
                {
                    literal.Append(pattern[index]);
                    index++;
                    continue;
                }

                if (literal.Length > 0)
                {
                    template.Segments.Add(
                        MathTypeWordOpenXml.NumberSegment.Text(literal.ToString()));
                    literal.Clear();
                }
                var end = pattern.IndexOf('\u0015', index + 1);
                if (end < 0)
                    throw new InvalidDataException(
                        "MathType MTPlaceRef contains an unterminated nested Word field.");
                var fieldBody = pattern.Substring(index + 1, end - index - 1);
                var separate = fieldBody.IndexOf('\u0014');
                if (separate >= 0)
                    fieldBody = fieldBody.Substring(0, separate);
                if (string.IsNullOrWhiteSpace(fieldBody))
                    throw new InvalidDataException(
                        "MathType MTPlaceRef contains an empty nested Word field.");
                template.Segments.Add(
                    MathTypeWordOpenXml.NumberSegment.Field(fieldBody));
                index = end + 1;
            }
            if (literal.Length > 0)
            {
                // Trailing whitespace is not part of a MathType equation-number
                // format. Never let it enter a reusable template, otherwise a later
                // equation/reference can inherit visible blanks after the number.
                var finalLiteral = literal.ToString().TrimEnd();
                if (finalLiteral.Length > 0)
                    template.Segments.Add(
                        MathTypeWordOpenXml.NumberSegment.Text(finalLiteral));
            }
            if (template.Segments.Count == 0)
                throw new InvalidDataException(
                    "MathType MTPlaceRef numbering template has no usable segments.");
            return template;
        }
        finally { Release(sourceCode); }
    }

    private static bool TryReadCompleteMathTypePlaceRefTemplate(
        Document document,
        Field source,
        out MathTypeWordOpenXml.NumberTemplate? template)
    {
        template = null;
        Range? sourceCode = null;
        Fields? nestedFields = null;
        try
        {
            sourceCode = source.Code;
            nestedFields = sourceCode.Fields;
            // A native MTPlaceRef always owns at least the hidden MTEqn increment
            // plus the visible MTEqn current-value field. Heading-aware templates
            // additionally own MTChap/MTSec fields. A literal-only outer field such
            // as "(.)" is a detached/corrupt scaffold and must never be accepted as
            // a valid numbered MathType row, even if it happens to sit on the
            // requested side of the OLE object.
            if (nestedFields.Count < 2) return false;

            var candidate = ReadMathTypePlaceRefTemplate(document, source);
            var hasHiddenEquationIncrement = candidate.Segments.Any(segment =>
                segment.IsField
                && segment.Value.IndexOf("SEQ MTEqn", StringComparison.OrdinalIgnoreCase) >= 0
                && segment.Value.IndexOf("\\h", StringComparison.OrdinalIgnoreCase) >= 0);
            var hasVisibleEquationValue = candidate.Segments.Any(segment =>
                segment.IsField
                && segment.Value.IndexOf("SEQ MTEqn", StringComparison.OrdinalIgnoreCase) >= 0
                && segment.Value.IndexOf("\\c", StringComparison.OrdinalIgnoreCase) >= 0);
            if (!hasHiddenEquationIncrement || !hasVisibleEquationValue)
                return false;

            template = candidate;
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(nestedFields);
            Release(sourceCode);
        }
    }

    private static bool TryReadReusableMathTypePlaceRefTemplate(
        Document document,
        Field source,
        out MathTypeWordOpenXml.NumberTemplate? template)
    {
        template = null;
        Range? visibleNumberRange = null;
        try
        {
            if (!MathTypeEquationReferences.TryGetVisibleNumberRange(
                    document,
                    source,
                    out visibleNumberRange)
                || visibleNumberRange is null)
                return false;
            if (string.IsNullOrWhiteSpace(
                    MathTypeEquationReferences.ReadVisibleNumberText(source)))
                return false;
            return TryReadCompleteMathTypePlaceRefTemplate(
                document,
                source,
                out template);
        }
        finally
        {
            Release(visibleNumberRange);
        }
    }

    private static bool MathTypeNumberTemplateUsesHeading(
        MathTypeWordOpenXml.NumberTemplate template) =>
        template.Segments.Any(segment =>
            segment.IsField
            && (segment.Value.IndexOf("SEQ MTChap", StringComparison.OrdinalIgnoreCase) >= 0
                || segment.Value.IndexOf("SEQ MTSec", StringComparison.OrdinalIgnoreCase) >= 0));

    private static Field? FindMathTypePlaceRefFieldForShape(
        Range range,
        Range shapeRange,
        bool numberOnLeft)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? resultRange = null;
        Field? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            fields = WordFormulaHost.GetLocalFields(range);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(resultRange);
                resultRange = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                resultRange = field.Result;
                var fieldStart = Math.Max(range.Start, code.Start - 1);
                var fieldEnd = Math.Min(range.End, resultRange.End + 1);
                int distance;
                if (numberOnLeft)
                {
                    if (fieldEnd > shapeRange.Start) continue;
                    distance = shapeRange.Start - fieldEnd;
                }
                else
                {
                    if (fieldStart < shapeRange.End) continue;
                    distance = fieldStart - shapeRange.End;
                }
                if (distance >= bestDistance) continue;
                Release(best);
                best = field;
                field = null;
                bestDistance = distance;
            }
            return best;
        }
        finally
        {
            Release(resultRange);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static Field? FindNearestMathTypePlaceRefField(
        Document document,
        int position,
        int excludeStart,
        int excludeEnd)
    {
        Range? content = null;
        Range? localBeforeRange = null;
        Range? localAfterRange = null;
        Range? beforeRange = null;
        Range? afterRange = null;
        Fields? localBeforeFields = null;
        Fields? localAfterFields = null;
        Fields? beforeFields = null;
        Fields? afterFields = null;
        Field? localBefore = null;
        Field? localAfter = null;
        Field? before = null;
        Field? after = null;
        try
        {
            content = document.Content;
            var safePosition = Math.Max(
                content.Start,
                Math.Min(position, content.End));

            // Ordinary insertion is usually adjacent to the previous MathType
            // equation. Search only a bounded character window first so adding the
            // Nth formula does not enumerate every field created by formulas 1..N-1.
            // A document-wide directional fallback remains for a genuinely distant
            // custom MathType numbering template.
            var localStart = Math.Max(content.Start, safePosition - 4096);
            var localEnd = Math.Min(content.End, safePosition + 4096);
            var localBeforeStart = -1;
            var localAfterStart = -1;
            if (safePosition > localStart)
            {
                localBeforeRange = document.Range(localStart, safePosition);
                localBeforeFields = localBeforeRange.Fields;
                localBefore = FindNearestMathTypePlaceRefFieldInCollection(
                    localBeforeFields,
                    safePosition,
                    excludeStart,
                    excludeEnd,
                    reverse: true,
                    stopAfterFirstMatch: true,
                    out localBeforeStart);
            }
            if (safePosition < localEnd)
            {
                localAfterRange = document.Range(safePosition, localEnd);
                localAfterFields = localAfterRange.Fields;
                localAfter = FindNearestMathTypePlaceRefFieldInCollection(
                    localAfterFields,
                    safePosition,
                    excludeStart,
                    excludeEnd,
                    reverse: false,
                    stopAfterFirstMatch: true,
                    out localAfterStart);
            }
            if (localBefore is not null || localAfter is not null)
            {
                if (localAfter is null
                    || localBefore is not null
                    && Math.Abs(localBeforeStart - safePosition)
                        <= Math.Abs(localAfterStart - safePosition))
                {
                    var localResult = localBefore;
                    localBefore = null;
                    return localResult;
                }
                var localAfterResult = localAfter;
                localAfter = null;
                return localAfterResult;
            }

            if (safePosition > content.Start)
            {
                beforeRange = document.Range(content.Start, safePosition);
                beforeFields = beforeRange.Fields;
                before = FindNearestMathTypePlaceRefFieldInCollection(
                    beforeFields,
                    safePosition,
                    excludeStart,
                    excludeEnd,
                    reverse: true,
                    stopAfterFirstMatch: true,
                    out var beforeStart);
                if (before is not null && safePosition - beforeStart == 0)
                {
                    var exact = before;
                    before = null;
                    return exact;
                }
            }

            if (safePosition < content.End)
            {
                afterRange = document.Range(safePosition, content.End);
                afterFields = afterRange.Fields;
                after = FindNearestMathTypePlaceRefFieldInCollection(
                    afterFields,
                    safePosition,
                    excludeStart,
                    excludeEnd,
                    reverse: false,
                    stopAfterFirstMatch: true,
                    out var afterStart);

                if (before is null)
                {
                    var result = after;
                    after = null;
                    return result;
                }
                if (after is null)
                {
                    var result = before;
                    before = null;
                    return result;
                }

                Range? beforeCode = null;
                try
                {
                    beforeCode = before.Code;
                    var beforeStart = Math.Max(content.Start, beforeCode.Start - 1);
                    if (Math.Abs(beforeStart - safePosition)
                        <= Math.Abs(afterStart - safePosition))
                    {
                        var result = before;
                        before = null;
                        return result;
                    }
                    var afterResult = after;
                    after = null;
                    return afterResult;
                }
                finally { Release(beforeCode); }
            }

            var onlyBefore = before;
            before = null;
            return onlyBefore;
        }
        finally
        {
            Release(after);
            Release(before);
            Release(localAfter);
            Release(localBefore);
            Release(afterFields);
            Release(beforeFields);
            Release(localAfterFields);
            Release(localBeforeFields);
            Release(afterRange);
            Release(beforeRange);
            Release(localAfterRange);
            Release(localBeforeRange);
            Release(content);
        }
    }

    private static Field? FindNearestMathTypePlaceRefFieldInCollection(
        Fields fields,
        int position,
        int excludeStart,
        int excludeEnd,
        bool reverse,
        bool stopAfterFirstMatch,
        out int matchedStart)
    {
        Field? field = null;
        Range? code = null;
        Field? best = null;
        var bestDistance = int.MaxValue;
        matchedStart = -1;
        try
        {
            var index = reverse ? fields.Count : 1;
            var end = reverse ? 1 : fields.Count;
            var step = reverse ? -1 : 1;
            for (; reverse ? index >= end : index <= end; index += step)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                var fieldStart = Math.Max(0, code.Start - 1);
                if (fieldStart >= excludeStart && fieldStart < excludeEnd) continue;
                var distance = Math.Abs(fieldStart - position);
                if (distance >= bestDistance) continue;
                Release(best);
                best = field;
                field = null;
                bestDistance = distance;
                matchedStart = fieldStart;
                if (stopAfterFirstMatch) break;
            }
            return best;
        }
        finally
        {
            Release(code);
            Release(field);
        }
    }

    internal static bool HasMathTypeSectionBreak(
        Document document,
        int beforePosition = int.MaxValue)
    {
        Range? content = null;
        Range? searchRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            content = document.Content;
            var searchEnd = beforePosition == int.MaxValue
                ? content.End
                : Math.Max(
                    content.Start,
                    Math.Min(content.End, beforePosition + 1));
            if (searchEnd <= content.Start) return false;
            searchRange = document.Range(content.Start, searchEnd);
            fields = searchRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(searchRange);
            Release(content);
        }
    }

    private static bool HasMathTypeSectionBreakBetween(
        Document document,
        int afterPosition,
        int beforePosition)
    {
        Range? content = null;
        Range? searchRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            content = document.Content;
            var start = Math.Max(
                content.Start,
                Math.Min(Math.Min(afterPosition, beforePosition), content.End));
            var end = Math.Max(
                start,
                Math.Min(Math.Max(afterPosition, beforePosition) + 1, content.End));
            if (end <= start) return false;
            searchRange = document.Range(start, end);
            fields = searchRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(searchRange);
            Release(content);
        }
    }

    internal static int EnsureMathTypeHeadingScopeState(
        Document document,
        int formulaPosition,
        EquationNumberFormat format) =>
        EnsureMathTypeHeadingScopeState(
            document,
            formulaPosition,
            format,
            out _);

    private static void EnsureExistingMathTypeHeadingScopes(
        Document document, string formatId, IReadOnlyList<int> numberedParagraphStarts)
    {
        var format = EquationNumberFormat.Resolve(formatId);
        if (!format.UsesHeading || numberedParagraphStarts.Count == 0) return;
        var scopes = WordEquationNumbering.CaptureHeadingScopesAtPositions(
            document, formatId, numberedParagraphStarts);
        // Insert once per actual heading, from the end of the document. Lower
        // captured positions then stay valid, including when several equations
        // share a scope. The shared inserter preserves existing native breaks.
        foreach (var group in scopes.Where(pair => pair.Value.ScopeStart != int.MinValue)
                     .GroupBy(pair => pair.Value.ScopeStart)
                     .OrderByDescending(group => group.Key))
        {
            var first = group.OrderBy(pair => pair.Key).First();
            var paragraphsBefore = ReadDocumentParagraphCount(document);
            var inserted = EnsureMathTypeHeadingScopeState(
                document, first.Key, format, out _, first.Value);
            if (ReadDocumentParagraphCount(document) != paragraphsBefore + (inserted > 0 ? 1 : 0))
                throw new InvalidDataException(
                    "MathType heading initialization changed paragraphs beyond its owned section state.");
        }
    }

    private static int EnsureMathTypeHeadingScopeState(
        Document document,
        int formulaPosition,
        EquationNumberFormat format,
        out int createdSectionBreakCodeStart,
        ResolvedEquationHeadingScope? preResolvedHeadingScope = null)
    {
        createdSectionBreakCodeStart = -1;
        if (!format.UsesHeading) return 0;

        var resolvedScope = preResolvedHeadingScope;
        var scope = resolvedScope is null
            ? WordEquationNumbering.ResolveHeadingScopeAtPosition(
                document,
                formulaPosition,
                format.Id)
            : (
                resolvedScope.ScopeStart,
                resolvedScope.ScopeEnd,
                resolvedScope.NumberText);
        // No real Heading paragraph exists before this equation. VisualTeX's
        // heading-aware numbering intentionally uses a zero prefix in that scope
        // (0.1 / 0.0-1). Do not manufacture MathType chapter 1 state.
        if (scope.ScopeStart == int.MinValue || scope.ScopeEnd == int.MinValue)
            return 0;
        if (scope.ScopeEnd > formulaPosition)
            return 0;

        var parts = (scope.NumberText ?? string.Empty)
            .Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0
            || !int.TryParse(
                parts[0],
                System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture,
                out var chapter))
            return 0;
        var section = 0;
        if (format.HeadingLevel >= 2
            && parts.Length >= 2
            && !int.TryParse(
                parts[1],
                System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture,
                out section))
            return 0;
        if (chapter == 0 && section == 0)
            return 0;

        // Scope state belongs immediately after the real Word heading, not before
        // whichever equation happens to be converted first. This is crucial for
        // descending whole-document conversion: every equation in the same scope
        // must remain after the same native MathType section state.
        if (HasMathTypeSectionBreakBetween(document, scope.ScopeEnd, formulaPosition))
            return 0;

        WordDoubleClickHook.TraceMessage(
            $"mathtype-heading-scope-state chapter={chapter} section={section} heading={scope.ScopeStart}:{scope.ScopeEnd} formula={formulaPosition}");
        var insertedLength = InsertMathTypeSectionBreakState(
            document,
            scope.ScopeEnd,
            chapter,
            section);
        if (insertedLength <= 0) return insertedLength;

        createdSectionBreakCodeStart = FindMathTypeSectionBreakCodeStartBetween(
            document,
            scope.ScopeEnd,
            Math.Min(document.Content.End, formulaPosition + insertedLength));
        if (createdSectionBreakCodeStart < 0)
            throw new InvalidOperationException(
                "VisualTeX inserted MathType heading state but could not identify the new section field for transactional rollback.");
        return insertedLength;
    }

    private static int FindMathTypeSectionBreakCodeStartBetween(
        Document document,
        int afterPosition,
        int beforePosition)
    {
        Range? content = null;
        Range? searchRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        var best = int.MaxValue;
        try
        {
            content = document.Content;
            var start = Math.Max(
                content.Start,
                Math.Min(Math.Min(afterPosition, beforePosition), content.End));
            var end = Math.Max(
                start,
                Math.Min(Math.Max(afterPosition, beforePosition) + 1, content.End));
            if (end <= start) return -1;
            searchRange = document.Range(start, end);
            fields = searchRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                best = Math.Min(best, code.Start);
            }
            return best == int.MaxValue ? -1 : best;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(searchRange);
            Release(content);
        }
    }

    private static void InsertIsolatedMathTypeParagraphBody(
        Document targetDocument,
        Range targetRange,
        string flatOpc)
    {
        Microsoft.Office.Interop.Word.Application? application = null;
        Documents? documents = null;
        Document? stagingDocument = null;
        Range? stagingInsertion = null;
        InlineShapes? stagingShapes = null;
        InlineShape? stagingShape = null;
        Range? stagingShapeRange = null;
        Paragraphs? stagingParagraphs = null;
        Paragraph? stagingParagraph = null;
        Range? stagingParagraphRange = null;
        Range? stagingBody = null;
        Range? formattedBody = null;
        try
        {
            if (!string.Equals(
                    targetRange.Text,
                    BulkInlineFormulaPlaceholder,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The isolated MathType paragraph transfer lost its body placeholder.");

            application = targetDocument.Application;
            documents = application.Documents;
            stagingDocument = documents.Add(Visible: false);
            stagingInsertion = stagingDocument.Range(
                stagingDocument.Content.Start,
                stagingDocument.Content.Start);
            stagingInsertion.InsertXML(flatOpc);

            stagingShapes = stagingDocument.InlineShapes;
            if (stagingShapes.Count != 1)
                throw new InvalidDataException(
                    $"The isolated MathType staging document materialized {stagingShapes.Count} OLE objects instead of one.");
            stagingShape = stagingShapes[1];
            if (!MathTypeOleInterop.IsMathTypeOle(stagingShape))
                throw new InvalidDataException(
                    "The isolated MathType staging document did not materialize Equation.DSMT4.");
            stagingShapeRange = stagingShape.Range;
            stagingParagraphs = stagingShapeRange.Paragraphs;
            if (stagingParagraphs.Count != 1)
                throw new InvalidDataException(
                    "The isolated MathType staging equation spans multiple paragraphs.");
            stagingParagraph = stagingParagraphs[1];
            stagingParagraphRange = stagingParagraph.Range;
            var bodyEnd = Math.Max(
                stagingParagraphRange.Start,
                stagingParagraphRange.End - 1);
            if (stagingShapeRange.Start < stagingParagraphRange.Start
                || stagingShapeRange.End > bodyEnd)
                throw new InvalidDataException(
                    "The isolated MathType staging equation escaped its paragraph body.");

            // Transfer only the completed paragraph BODY. The embedded OLE, WMF
            // presentation, native MathType field, tabs and optional MTPlaceRef
            // field tree are preserved by FormattedText, while the staging <w:p>
            // itself is deliberately excluded. This prevents Word from adopting a
            // new paragraph into a user table immediately following the target.
            stagingBody = stagingDocument.Range(
                stagingParagraphRange.Start,
                bodyEnd);
            formattedBody = stagingBody.FormattedText;
            targetRange.FormattedText = formattedBody;
        }
        finally
        {
            Release(formattedBody);
            Release(stagingBody);
            Release(stagingParagraphRange);
            Release(stagingParagraph);
            Release(stagingParagraphs);
            Release(stagingShapeRange);
            Release(stagingShape);
            Release(stagingShapes);
            Release(stagingInsertion);
            if (stagingDocument is not null)
            {
                try { stagingDocument.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(stagingDocument);
            Release(documents);
            Release(application);
            try { targetDocument.Activate(); } catch { }
        }
    }

    private static void InsertIsolatedMathTypeSectionParagraph(
        Document targetDocument,
        int insertionPosition,
        string sectionFlatOpc)
    {
        Microsoft.Office.Interop.Word.Application? application = null;
        Documents? documents = null;
        Document? stagingDocument = null;
        Range? stagingInsertion = null;
        Field? stagingSection = null;
        Range? stagingCode = null;
        Range? stagingResult = null;
        Range? stagingFull = null;
        Paragraphs? stagingParagraphs = null;
        Paragraph? stagingParagraph = null;
        Range? stagingParagraphRange = null;
        Range? targetBreak = null;
        Range? targetProbe = null;
        Paragraphs? targetParagraphs = null;
        Paragraph? targetParagraph = null;
        Range? targetParagraphRange = null;
        try
        {
            application = targetDocument.Application;
            documents = application.Documents;
            stagingDocument = documents.Add(Visible: false);
            stagingInsertion = stagingDocument.Range(
                stagingDocument.Content.Start,
                stagingDocument.Content.Start);
            stagingInsertion.InsertXML(sectionFlatOpc);
            stagingSection = FindFirstMathTypeSectionBreakField(stagingDocument)
                ?? throw new InvalidOperationException(
                    "The isolated MathType section-state document did not materialize MTEditEquationSection2.");
            stagingCode = stagingSection.Code;
            stagingResult = stagingSection.Result;
            stagingFull = stagingDocument.Range(
                Math.Max(stagingDocument.Content.Start, stagingCode.Start - 1),
                Math.Min(stagingDocument.Content.End, stagingResult.End + 1));
            stagingParagraphs = stagingFull.Paragraphs;
            if (stagingParagraphs.Count != 1)
                throw new InvalidOperationException(
                    "The isolated MathType section-state field spans multiple staging paragraphs.");
            stagingParagraph = stagingParagraphs[1];
            stagingParagraphRange = stagingParagraph.Range;
            if (stagingParagraphRange.InlineShapes.Count != 0)
                throw new InvalidOperationException(
                    "The isolated MathType section-state staging paragraph unexpectedly contains an OLE object.");

            targetBreak = targetDocument.Range(insertionPosition, insertionPosition);
            targetBreak.Text = "\r";
            targetProbe = targetDocument.Range(
                insertionPosition,
                Math.Min(targetDocument.Content.End, insertionPosition + 1));
            targetParagraphs = targetProbe.Paragraphs;
            if (targetParagraphs.Count != 1)
                throw new InvalidOperationException(
                    "Word did not create a dedicated target paragraph for MathType section state.");
            targetParagraph = targetParagraphs[1];
            targetParagraphRange = targetParagraph.Range;
            targetParagraphRange.FormattedText = stagingParagraphRange.FormattedText;
        }
        finally
        {
            Release(targetParagraphRange);
            Release(targetParagraph);
            Release(targetParagraphs);
            Release(targetProbe);
            Release(targetBreak);
            Release(stagingParagraphRange);
            Release(stagingParagraph);
            Release(stagingParagraphs);
            Release(stagingFull);
            Release(stagingResult);
            Release(stagingCode);
            Release(stagingSection);
            Release(stagingInsertion);
            if (stagingDocument is not null)
            {
                try { stagingDocument.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(stagingDocument);
            Release(documents);
            Release(application);
        }
    }

    internal static int EnsureDefaultMathTypeSectionBreak(
        Document document,
        int beforePosition = int.MaxValue)
    {
        if (HasMathTypeSectionBreak(document, beforePosition)) return 0;
        return InsertMathTypeSectionBreakState(
            document,
            beforePosition,
            chapter: 1,
            section: 1);
    }

    private static int InsertMathTypeSectionBreakState(
        Document document,
        int beforePosition,
        int chapter,
        int section)
    {
        EnsureMathTypeNativeStyles(document);

        var contentEndBefore = document.Content.End;
        var paragraphCountBefore = ReadDocumentParagraphCount(document);
        var label = string.Equals(
            System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName,
            "zh",
            StringComparison.OrdinalIgnoreCase)
            ? $"公式章 {chapter} 节 {section}"
            : $"Equation Chapter {chapter} Section {section}";
        var breakXml = MathTypeWordOpenXml.CreateSectionBreakFlatOpc(
            label,
            chapter,
            section);
        Range? insertion = null;
        Field? sectionBreak = null;
        Range? sectionCode = null;
        Range? sectionResult = null;
        Range? sectionFull = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Range? sectionParagraphSplit = null;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = document.Content.End;
            var insertionPosition = beforePosition == int.MaxValue
                ? contentStart
                : Math.Max(
                    contentStart,
                    Math.Min(beforePosition, Math.Max(contentStart, contentEnd - 1)));
            insertion = document.Range(insertionPosition, insertionPosition);
            // Always anchor the hidden section state at the containing paragraph
            // boundary. A shape/bookmark Start inside an MTDisplayEquation row can
            // otherwise place MTEditEquationSection2 after the row's leading tab.
            paragraphs = insertion.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
                insertionPosition = paragraphRange.Start;
                insertion.SetRange(insertionPosition, insertionPosition);
            }
            Release(paragraphRange);
            paragraphRange = null;
            Release(paragraph);
            paragraph = null;
            Release(paragraphs);
            paragraphs = null;

            // Target-document InsertXML can merge this complete <w:p> field tree
            // back into the following MTDisplayEquation paragraph. Materialize it
            // in an isolated hidden Word document first, then transfer the single
            // finished paragraph with FormattedText into a reserved target row.
            InsertIsolatedMathTypeSectionParagraph(
                document,
                insertionPosition,
                breakXml);
            sectionBreak = FindFirstMathTypeSectionBreakField(document, insertionPosition)
                ?? throw new InvalidOperationException(
                    "Word did not materialize MathType's MTEditEquationSection2 field.");
            sectionCode = sectionBreak.Code;
            sectionResult = sectionBreak.Result;
            sectionFull = document.Range(
                Math.Max(document.Content.Start, sectionCode.Start - 1),
                Math.Min(document.Content.End, sectionResult.End + 1));

            // Do not infer isolation from the total paragraph count. Word can add
            // a paragraph while still leaving MTEditEquationSection2 and the first
            // Equation.DSMT4 in the same row. Inspect the field's actual paragraph
            // and split immediately after the outer field whenever anything else
            // remains before that paragraph mark.
            paragraphs = sectionFull.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "MathType chapter/section break materialized across multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var paragraphContentEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            if (sectionFull.End < paragraphContentEnd)
            {
                sectionParagraphSplit = document.Range(sectionFull.End, sectionFull.End);
                sectionParagraphSplit.Text = "\r";
            }

            Release(paragraphRange);
            paragraphRange = null;
            Release(paragraph);
            paragraph = null;
            Release(paragraphs);
            paragraphs = null;

            Release(sectionFull);
            sectionFull = null;
            Release(sectionResult);
            sectionResult = null;
            Release(sectionCode);
            sectionCode = null;
            sectionCode = sectionBreak.Code;
            sectionResult = sectionBreak.Result;
            sectionFull = document.Range(
                Math.Max(document.Content.Start, sectionCode.Start - 1),
                Math.Min(document.Content.End, sectionResult.End + 1));
            paragraphs = sectionFull.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "MathType chapter/section break materialized across multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.InlineShapes.Count != 0)
                throw new InvalidOperationException(
                    "MathType chapter/section state still shares a paragraph with an OLE equation.");
            for (var fieldIndex = 1; fieldIndex <= paragraphRange.Fields.Count; fieldIndex++)
            {
                Field? paragraphField = null;
                Range? paragraphFieldCode = null;
                try
                {
                    paragraphField = paragraphRange.Fields[fieldIndex];
                    paragraphFieldCode = paragraphField.Code;
                    var fieldCodeText = paragraphFieldCode.Text ?? string.Empty;
                    if (fieldCodeText.IndexOf("MTPlaceRef", StringComparison.OrdinalIgnoreCase) >= 0
                        || fieldCodeText.IndexOf("EMBED ", StringComparison.OrdinalIgnoreCase) >= 0)
                        throw new InvalidOperationException(
                            "MathType chapter/section state still shares a paragraph with equation-owned fields.");
                }
                finally
                {
                    Release(paragraphFieldCode);
                    Release(paragraphField);
                }
            }
            paragraphMark = document.Range(
                Math.Max(paragraphRange.Start, paragraphRange.End - 1),
                paragraphRange.End);
            if (!string.Equals(paragraphMark.Text, "\r", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "MathType chapter/section break has no isolated paragraph mark.");

            // MTEquationSection is a hidden red *character* style.  The paragraph
            // mark must stay with the hidden field.  Deleting that mark merges the
            // hidden character formatting into the following user paragraph, so
            // the next field/text insertion inherits MathType's red state.
            object sectionStyle = "MTEquationSection";
            paragraphRange.set_Style(ref sectionStyle);
            UpdateNestedMathTypeNumberFields(sectionBreak);
            return document.Content.End - contentEndBefore;
        }
        finally
        {
            Release(sectionParagraphSplit);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(sectionFull);
            Release(sectionResult);
            Release(sectionCode);
            Release(sectionBreak);
            Release(insertion);
        }
    }

    private static Field? FindFirstMathTypeSectionBreakField(Document document, int? insertionPosition = null)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Field? result = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                if (insertionPosition.HasValue && code.Start != insertionPosition.Value + 1)
                    continue;
                result = field;
                field = null;
                break;
            }
            return result;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    internal static void RemoveAllMathTypeSectionBreakFields(Document document)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        var starts = new List<int>();
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    starts.Add(code.Start);
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }

        foreach (var start in starts.OrderByDescending(value => value))
            RemoveMathTypeSectionBreakFieldAtCodeStart(document, start);
    }

    private static void RemoveMathTypeSectionBreakFieldAtCodeStart(
        Document document,
        int codeStart)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        Range? full = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (code.Start != codeStart
                    || (code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                paragraphs = code.Paragraphs;
                if (paragraphs.Count == 1)
                {
                    paragraph = paragraphs[1];
                    paragraphRange = paragraph.Range.Duplicate;
                }

                // Delete the exact section field created by this operation. Never
                // delete "the first" MathType section field: later-chapter failures
                // must not destroy an earlier chapter's valid native state.
                try { field.Delete(); }
                catch
                {
                    result = field.Result;
                    var start = Math.Max(document.Content.Start, code.Start - 1);
                    var end = Math.Min(
                        document.Content.End,
                        Math.Max(code.End, result.End) + 1);
                    full = document.Range(start, end);
                    full.Delete();
                }

                if (paragraphRange is not null
                    && paragraphRange.InlineShapes.Count == 0)
                {
                    try
                    {
                        object normalStyle = WdBuiltinStyle.wdStyleNormal;
                        paragraphRange.set_Style(ref normalStyle);
                        paragraphRange.Font.Hidden = 0;
                        paragraphRange.Delete();
                    }
                    catch { }
                }
                return;
            }
        }
        catch { }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(full);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static void EnsureMathTypeNativeStyles(
        Document document,
        WordCharacterFormatting? displayBodyFormatting = null)
    {
        Styles? styles = null;
        Style? displayStyle = null;
        Style? sectionStyle = null;
        ParagraphFormat? displayFormat = null;
        TabStops? displayTabs = null;
        TabStop? tab = null;
        PageSetup? pageSetup = null;
        Microsoft.Office.Interop.Word.Font? sectionFont = null;
        Microsoft.Office.Interop.Word.Font? displayFont = null;
        try
        {
            styles = document.Styles;
            // Section-state initialization must not pre-create an incomplete
            // display style. Native MathType seeds that style from the first
            // insertion's body formatting and reuses it for subsequent equations.
            if (displayBodyFormatting is not null)
            {
                object displayName = "MTDisplayEquation";
                try { displayStyle = styles.get_Item(ref displayName); }
                catch
                {
                    object paragraphType = WdStyleType.wdStyleTypeParagraph;
                    displayStyle = styles.Add("MTDisplayEquation", ref paragraphType);
                    object normalStyle = WdBuiltinStyle.wdStyleNormal;
                    displayStyle.set_BaseStyle(ref normalStyle);
                    displayStyle.set_NextParagraphStyle(ref normalStyle);
                    displayFont = displayStyle.Font;
                    displayBodyFormatting.ApplyToParagraphStyle(displayFont);
                    // Match MathType's own behavior: create the style only when the
                    // document does not already have it. Never overwrite an existing
                    // native MathType style, because its tab geometry may have been
                    // adapted by MathType for the current document/template.
                    displayFormat = displayStyle.ParagraphFormat;
                    displayFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    displayFormat.SpaceBefore = 0;
                    displayFormat.SpaceAfter = 0;
                    displayFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                    displayTabs = displayFormat.TabStops;
                    displayTabs.ClearAll();
                    pageSetup = document.PageSetup;
                    var usableWidth = Math.Max(
                        72f,
                        pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin);
                    tab = displayTabs.Add(
                        usableWidth / 2f,
                        WdTabAlignment.wdAlignTabCenter,
                        WdTabLeader.wdTabLeaderSpaces);
                    Release(tab);
                    tab = displayTabs.Add(
                        usableWidth,
                        WdTabAlignment.wdAlignTabRight,
                        WdTabLeader.wdTabLeaderSpaces);
                    Release(tab);
                    tab = null;
                }
            }

            object sectionName = "MTEquationSection";
            try { sectionStyle = styles.get_Item(ref sectionName); }
            catch
            {
                object characterType = WdStyleType.wdStyleTypeCharacter;
                sectionStyle = styles.Add("MTEquationSection", ref characterType);
            }
            sectionFont = sectionStyle.Font;
            sectionFont.Hidden = -1;
            sectionFont.Color = WdColor.wdColorRed;
        }
        finally
        {
            Release(displayFont);
            Release(sectionFont);
            Release(pageSetup);
            Release(tab);
            Release(displayTabs);
            Release(displayFormat);
            Release(sectionStyle);
            Release(displayStyle);
            Release(styles);
        }
    }

    private static int FinalizeMathTypeRedrawDisplayLayouts(
        Document document,
        IReadOnlyList<(string FormulaId, Range Range, string Signature)> identities,
        ISet<string> displayFormulaIds)
    {
        var normalized = 0;
        foreach (var identity in identities)
        {
            if (!displayFormulaIds.Contains(identity.FormulaId)) continue;
            InlineShapes? shapes = null;
            InlineShape? shape = null;
            try
            {
                shapes = identity.Range.InlineShapes;
                if (shapes.Count != 1)
                    throw new InvalidDataException(
                        $"MathType redraw display {identity.FormulaId} no longer owns exactly one OLE object.");
                shape = shapes[1];
                if (!MathTypeOleInterop.IsMathTypeOle(shape))
                    throw new InvalidDataException(
                        $"MathType redraw display {identity.FormulaId} no longer owns Equation.DSMT4.");
                NormalizeMathTypeDisplayTabGeometry(document, shape);
                normalized++;
            }
            finally
            {
                Release(shape);
                Release(shapes);
            }
        }
        return normalized;
    }

    private static void NormalizeMathTypeDisplayTabGeometry(
        Document document,
        InlineShape shape)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        TabStops? tabs = null;
        TabStop? tab = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "A completed MathType display equation no longer occupies one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraph.Format;

            var columnWidth = ResolveMathTypeDisplayColumnWidth(document, paragraphRange);
            var firstLineLeft = Math.Max(
                0f,
                format.LeftIndent + format.FirstLineIndent);
            var rightEdge = Math.Max(
                firstLineLeft + 1f,
                columnWidth - Math.Max(0f, format.RightIndent));
            rightEdge = Math.Min(columnWidth, rightEdge);
            var center = firstLineLeft + ((rightEdge - firstLineLeft) / 2f);

            tabs = format.TabStops;
            tabs.ClearAll();
            Release(tabs);
            tabs = format.TabStops;
            tab = tabs.Add(
                center,
                WdTabAlignment.wdAlignTabCenter,
                WdTabLeader.wdTabLeaderSpaces);
            Release(tab);
            tab = null;
            Release(tabs);
            tabs = format.TabStops;
            tab = tabs.Add(
                rightEdge,
                WdTabAlignment.wdAlignTabRight,
                WdTabLeader.wdTabLeaderSpaces);
            Release(tab);
            tab = null;

            // Re-read the effective collection after both mutations. This catches
            // the Word case that originally caused the bug: Add() can return
            // successfully while a later style import makes the inherited full-page
            // tab pair effective again.
            Release(tabs);
            tabs = format.TabStops;
            var hasCenter = false;
            var hasRight = false;
            for (var index = 1; index <= tabs.Count; index++)
            {
                Release(tab);
                tab = tabs[index];
                if (tab.Alignment == WdTabAlignment.wdAlignTabCenter
                    && Math.Abs(tab.Position - center) <= 0.5f)
                    hasCenter = true;
                if (tab.Alignment == WdTabAlignment.wdAlignTabRight
                    && Math.Abs(tab.Position - rightEdge) <= 0.5f)
                    hasRight = true;
            }
            if (!hasCenter || !hasRight)
                throw new InvalidDataException(
                    "Word did not retain the completed MathType display equation's column-local tab geometry.");
        }
        finally
        {
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static void NormalizeMathTypeDisplayParagraphFormat(
        Document document,
        Paragraph paragraph,
        Range paragraphRange,
        MathTypeDisplayParagraphLayout? preservedLayout = null,
        float? knownColumnWidth = null)
    {
        ParagraphFormat? format = null;
        TabStops? tabs = null;
        TabStop? tab = null;
        try
        {
            format = paragraph.Format;
            if (preservedLayout is null)
            {
                format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                format.LeftIndent = 0;
                format.RightIndent = 0;
                format.FirstLineIndent = 0;
                format.SpaceBefore = 0;
                format.SpaceAfter = 0;
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            }
            else
            {
                // Redraw replaces source LaTeX in an already-owned paragraph. Keep
                // that paragraph's page-flow/spacing contract instead of silently
                // replacing it with MTDisplayEquation's document-wide defaults.
                // The native MathType center/right tab scaffold is rebuilt below,
                // so source custom tabs are intentionally not restored here.
                // MathType display equations are positioned by the native center/right
                // tab scaffold below. Preserving a centered LaTeX source paragraph here
                // applies paragraph centering on top of that tab geometry and visibly
                // shifts the OLE in both single- and multi-column layouts. Keep the
                // source paragraph's flow/spacing metrics, but restore MathType's
                // tab-driven paragraph alignment contract.
                format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                format.LeftIndent = preservedLayout.LeftIndent;
                format.RightIndent = preservedLayout.RightIndent;
                format.FirstLineIndent = preservedLayout.FirstLineIndent;
                format.SpaceBefore = preservedLayout.SpaceBefore;
                format.SpaceAfter = preservedLayout.SpaceAfter;
                format.LineSpacingRule = preservedLayout.LineSpacingRule;
                try { format.LineSpacing = preservedLayout.LineSpacing; } catch { }
                format.KeepTogether = preservedLayout.KeepTogether;
                format.KeepWithNext = preservedLayout.KeepWithNext;
                format.WidowControl = preservedLayout.WidowControl;
                format.PageBreakBefore = preservedLayout.PageBreakBefore;
            }

            tabs = format.TabStops;
            tabs.ClearAll();
            // Word invalidates the COM TabStops collection after ClearAll when the
            // cleared stops were inherited from a paragraph style. Reusing that
            // stale collection makes Add() return normally but silently leaves the
            // inherited MTDisplayEquation tabs in force. Reacquire after every
            // structural mutation so the direct paragraph overrides are real.
            Release(tabs);
            tabs = null;

            // MTDisplayEquation is a document-global style. MathType normally seeds
            // its tabs from the full page text width, which is wrong as soon as a
            // section uses newspaper columns: a two-column page can inherit a center
            // tab at half of the *whole page* width, already beyond the first column.
            // Direct-format every live display row from the width of the column that
            // actually owns this paragraph. This also overrides an old/full-page
            // MTDisplayEquation style without mutating the global style itself.
            // Large redraws resolve the section/column geometry once while the
            // source document is still intact and pass it here. Avoiding repeated
            // Sections/PageSetup/TextColumns COM calls is essential: on a 1000-formula
            // corpus those queries alone can add minutes of Word repagination work.
            var columnWidth = knownColumnWidth
                ?? ResolveMathTypeDisplayColumnWidth(document, paragraphRange);
            var firstLineLeft = Math.Min(
                Math.Max(0f, columnWidth - 1f),
                Math.Max(
                    0f,
                    format.LeftIndent + format.FirstLineIndent));
            var rightEdge = Math.Max(
                firstLineLeft + 1f,
                columnWidth - Math.Max(0f, format.RightIndent));
            rightEdge = Math.Min(columnWidth, rightEdge);
            var center = firstLineLeft + ((rightEdge - firstLineLeft) / 2f);

            tabs = format.TabStops;
            tab = tabs.Add(
                center,
                WdTabAlignment.wdAlignTabCenter,
                WdTabLeader.wdTabLeaderSpaces);
            Release(tab);
            tab = null;
            Release(tabs);
            tabs = format.TabStops;
            tab = tabs.Add(
                rightEdge,
                WdTabAlignment.wdAlignTabRight,
                WdTabLeader.wdTabLeaderSpaces);
        }
        finally
        {
            Release(tab);
            Release(tabs);
            Release(format);
        }
    }

    private static float ResolveMathTypeDisplayColumnWidth(
        Document document,
        Range paragraphRange)
    {
        if (TryResolveOwningTableCellContentWidth(paragraphRange, out var cellWidth))
            return cellWidth;

        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        TextColumns? textColumns = null;
        TextColumn? textColumn = null;
        try
        {
            try
            {
                sections = paragraphRange.Sections;
                if (sections.Count > 0)
                {
                    section = sections[1];
                    pageSetup = section.PageSetup;
                }
            }
            catch
            {
                Release(pageSetup);
                pageSetup = null;
                Release(section);
                section = null;
                Release(sections);
                sections = null;
            }

            pageSetup ??= document.PageSetup;
            var pageTextWidth = Math.Max(
                72f,
                pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin);
            textColumns = pageSetup.TextColumns;
            if (textColumns is null || textColumns.Count <= 1)
                return pageTextWidth;

            var widths = new List<float>(textColumns.Count);
            var spaces = new List<float>(textColumns.Count);
            for (var index = 1; index <= textColumns.Count; index++)
            {
                Release(textColumn);
                textColumn = textColumns[index];
                var width = textColumn.Width;
                if (!(width > 1f) || float.IsNaN(width) || float.IsInfinity(width))
                    continue;

                // Word exposes SpaceAfter only for columns that are followed by
                // another column. On the final TextColumn some Office builds return
                // a missing COM value (and others throw). Letting that escape made
                // the whole resolver fall back to full-page width, exactly the
                // geometry that breaks MathType display equations in newspaper
                // columns. A final column has no following gap by definition.
                var spaceAfter = 0f;
                if (index < textColumns.Count)
                {
                    try
                    {
                        var candidateSpace = textColumn.SpaceAfter;
                        if (candidateSpace > 0f
                            && !float.IsNaN(candidateSpace)
                            && !float.IsInfinity(candidateSpace))
                            spaceAfter = candidateSpace;
                    }
                    catch
                    {
                        // The width is still authoritative. Missing spacing must
                        // never downgrade a valid multi-column section to page width.
                    }
                }
                widths.Add(width);
                spaces.Add(spaceAfter);
            }
            if (widths.Count == 0) return pageTextWidth;
            if (widths.Count == 1) return widths[0];

            // Unequal-width newspaper columns are uncommon but legal. Word reports
            // a live paragraph's horizontal page coordinate, so use it to select
            // the owning column when pagination is available. If background layout
            // cannot provide a coordinate, the narrowest column is the conservative
            // fallback: it cannot place the equation outside any valid column.
            try
            {
                var pageX = Convert.ToSingle(paragraphRange.get_Information(
                    WdInformation.wdHorizontalPositionRelativeToPage));
                if (!float.IsNaN(pageX)
                    && !float.IsInfinity(pageX)
                    && pageX >= 0f)
                {
                    var relativeX = pageX - pageSetup.LeftMargin;
                    var cursor = 0f;
                    for (var index = 0; index < widths.Count; index++)
                    {
                        var next = cursor + widths[index] + spaces[index];
                        if (relativeX < next || index == widths.Count - 1)
                            return widths[index];
                        cursor = next;
                    }
                }
            }
            catch
            {
                // Background/protected stories can reject page-position queries.
            }

            return widths.Min();
        }
        catch
        {
            // Preserve the mature single-column behavior when section/column COM
            // metadata is unavailable in a custom story.
            try
            {
                pageSetup ??= document.PageSetup;
                return Math.Max(
                    72f,
                    pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin);
            }
            catch
            {
                return 468f;
            }
        }
        finally
        {
            Release(textColumn);
            Release(textColumns);
            Release(pageSetup);
            Release(section);
            Release(sections);
        }
    }

    private static string ReadMathTypeNumberPositionPreference(Document document)
    {
        object? propertiesObject = null;
        object? propertyObject = null;
        try
        {
            propertiesObject = document.CustomDocumentProperties;
            if (propertiesObject is null) return "right";
            dynamic properties = propertiesObject;
            try
            {
                propertyObject = properties["MTEqnNumsOnRight"];
                dynamic property = propertyObject;
                var value = property.Value;
                if (value is bool right) return right ? "right" : "left";
                if (value is int integer) return integer != 0 ? "right" : "left";
                var text = Convert.ToString((object)value);
                if (bool.TryParse(text, out bool parsed))
                    return parsed ? "right" : "left";
            }
            catch { }
            return "right";
        }
        finally
        {
            Release(propertyObject);
            Release(propertiesObject);
        }
    }

    private static void UpdateNestedMathTypeNumberFields(Field outer)
    {
        Range? code = null;
        Fields? fields = null;
        Field? field = null;
        try
        {
            code = outer.Code;
            fields = code.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(field);
                field = fields[index];
                try { field.Update(); } catch { }
            }
            try { outer.Update(); } catch { }
        }
        finally
        {
            Release(field);
            Release(fields);
            Release(code);
        }
    }

    private static MathTypeDisplayParagraphLayout? CaptureMathTypeDisplayParagraphLayout(
        Range range)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        try
        {
            paragraphs = range.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            return CaptureMathTypeDisplayParagraphLayout(paragraph);
        }
        catch { return null; }
        finally
        {
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static MathTypeDisplayParagraphLayout? CaptureMathTypeDisplayParagraphLayout(
        InlineShape shape)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            return CaptureMathTypeDisplayParagraphLayout(paragraph);
        }
        catch { return null; }
        finally
        {
            Release(tabStop);
            Release(tabStops);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static MathTypeDisplayParagraphLayout? CaptureMathTypeDisplayParagraphLayout(
        Paragraph paragraph)
    {
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            format = paragraph.Format;
            var layout = new MathTypeDisplayParagraphLayout
            {
                Alignment = format.Alignment,
                LeftIndent = format.LeftIndent,
                RightIndent = format.RightIndent,
                FirstLineIndent = format.FirstLineIndent,
                SpaceBefore = format.SpaceBefore,
                SpaceAfter = format.SpaceAfter,
                LineSpacingRule = format.LineSpacingRule,
                LineSpacing = format.LineSpacing,
                BaseLineAlignment = format.BaseLineAlignment,
                KeepTogether = format.KeepTogether,
                KeepWithNext = format.KeepWithNext,
                WidowControl = format.WidowControl,
                PageBreakBefore = format.PageBreakBefore,
            };
            tabStops = format.TabStops;
            for (var index = 1; index <= tabStops.Count; index++)
            {
                Release(tabStop);
                tabStop = tabStops[index];
                if (tabStop.Alignment == WdTabAlignment.wdAlignTabLeft) continue;
                layout.SpecialTabStops.Add((tabStop.Position, tabStop.Alignment, tabStop.Leader));
            }
            return layout;
        }
        catch { return null; }
        finally
        {
            Release(tabStop);
            Release(tabStops);
            Release(format);
        }
    }

    private static int ReadDocumentParagraphCount(Document document)
    {
        Paragraphs? paragraphs = null;
        try
        {
            paragraphs = document.Paragraphs;
            return paragraphs.Count;
        }
        finally { Release(paragraphs); }
    }

    private static string CaptureLeadingHorizontalWhitespace(
        Document document,
        int position)
    {
        Range? probe = null;
        Range? content = null;
        try
        {
            content = document.Content;
            var start = Math.Max(
                content.Start,
                Math.Min(position, content.End));
            var end = Math.Min(
                content.End,
                start + 64);
            if (end <= start)
                return string.Empty;

            probe = document.Range(
                start,
                end);
            var text = probe.Text
                ?? string.Empty;
            var length = 0;
            while (length < text.Length)
            {
                var character = text[length];
                if (character is not (' ' or '\t' or '\u00A0'))
                    break;
                length++;
            }
            return length == 0
                ? string.Empty
                : text.Substring(0, length);
        }
        finally
        {
            Release(content);
            Release(probe);
        }
    }

    private static void RepairMathTypeInsertXmlParagraphSplit(
        Document document,
        InlineShape shape,
        int sourceParagraphCount,
        string? preservedRightBoundaryWhitespace = null)
    {
        var currentParagraphCount = ReadDocumentParagraphCount(document);
        if (currentParagraphCount == sourceParagraphCount) return;
        if (currentParagraphCount != sourceParagraphCount + 1)
            throw new InvalidOperationException(
                $"Word changed the paragraph count unexpectedly while materializing a MathType OLE: before={sourceParagraphCount}, after={currentParagraphCount}.");

        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Range? followingProbe = null;
        Range? mergeRange = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The materialized MathType OLE spans multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.End <= paragraphRange.Start)
                throw new InvalidOperationException(
                    "The materialized MathType OLE paragraph has no paragraph mark to repair.");
            paragraphMark = document.Range(paragraphRange.End - 1, paragraphRange.End);
            if (!string.Equals(paragraphMark.Text, "\r", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "Word inserted a MathType OLE paragraph boundary that VisualTeX could not identify safely.");

            // InsertXML materializes an inline MathType OLE as a complete Word
            // paragraph and can split the user's original paragraph at the
            // insertion point. Deleting only that paragraph mark lets Word
            // normalize away leading spaces from the second paragraph. Merge the
            // exact structural boundary and its existing leading whitespace in
            // one replacement instead: the paragraph mark disappears while the
            // user's whitespace is recreated byte-for-byte as ordinary text.
            var probeStart = paragraphRange.End;
            var probeEnd = Math.Min(
                document.Content.End,
                probeStart + 64);
            if (probeEnd > probeStart)
            {
                followingProbe = document.Range(
                    probeStart,
                    probeEnd);
            }

            var followingText =
                followingProbe?.Text
                ?? string.Empty;
            var actualLeadingLength = 0;
            while (actualLeadingLength < followingText.Length)
            {
                var character =
                    followingText[actualLeadingLength];
                if (character is not (' ' or '\t' or '\u00A0'))
                    break;
                actualLeadingLength++;
            }

            var desiredLeadingWhitespace =
                preservedRightBoundaryWhitespace
                ?? (actualLeadingLength == 0
                    ? string.Empty
                    : followingText.Substring(
                        0,
                        actualLeadingLength));

            if (actualLeadingLength == 0
                && desiredLeadingWhitespace.Length == 0)
            {
                paragraphMark.Delete();
            }
            else
            {
                mergeRange = document.Range(
                    paragraphMark.Start,
                    paragraphRange.End
                    + actualLeadingLength);
                mergeRange.Text =
                    desiredLeadingWhitespace;
            }
        }
        finally
        {
            Release(mergeRange);
            Release(followingProbe);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }

        var repairedParagraphCount = ReadDocumentParagraphCount(document);
        if (repairedParagraphCount != sourceParagraphCount)
            throw new InvalidOperationException(
                $"VisualTeX could not restore the original MathType OLE paragraph structure: expected={sourceParagraphCount}, actual={repairedParagraphCount}.");
    }

    private static MathTypeDisplayParagraphLayout? RepairDetachedMathTypeNumberParagraph(
        Document document,
        InlineShape shape)
    {
        Range? shapeRange = null;
        Paragraphs? shapeParagraphs = null;
        Paragraph? shapeParagraph = null;
        Range? shapeParagraphRange = null;
        Range? nextProbe = null;
        Paragraphs? nextParagraphs = null;
        Paragraph? nextParagraph = null;
        Range? nextRange = null;
        Range? paragraphMark = null;
        try
        {
            shapeRange = shape.Range;
            shapeParagraphs = shapeRange.Paragraphs;
            if (shapeParagraphs.Count != 1) return null;
            shapeParagraph = shapeParagraphs[1];
            shapeParagraphRange = shapeParagraph.Range;
            if (shapeParagraphRange.End >= document.Content.End) return null;

            // Probe the first character after the formula paragraph mark.  A
            // non-empty one-character range reliably belongs to the following
            // paragraph, unlike a collapsed boundary range which Word can resolve
            // to either side depending on field state.
            nextProbe = document.Range(
                shapeParagraphRange.End,
                Math.Min(document.Content.End, shapeParagraphRange.End + 1));
            nextParagraphs = nextProbe.Paragraphs;
            if (nextParagraphs.Count < 1) return null;
            nextParagraph = nextParagraphs[1];
            nextRange = nextParagraph.Range;
            if (!IsDetachedMathTypeNumberParagraph(nextRange)) return null;

            var numberingLayout = CaptureMathTypeDisplayParagraphLayout(nextParagraph);
            paragraphMark = document.Range(
                shapeParagraphRange.End - 1,
                shapeParagraphRange.End);
            if (!string.Equals(paragraphMark.Text, "\r", StringComparison.Ordinal))
                return null;
            paragraphMark.Delete();
            return numberingLayout;
        }
        catch
        {
            // This is compatibility recovery for documents already damaged by an
            // older VisualTeX build.  Never make an otherwise valid MathType edit
            // fail merely because an adjacent paragraph only resembles MTPlaceRef.
            return null;
        }
        finally
        {
            Release(paragraphMark);
            Release(nextRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(nextProbe);
            Release(shapeParagraphRange);
            Release(shapeParagraph);
            Release(shapeParagraphs);
            Release(shapeRange);
        }
    }

    private static bool IsDetachedMathTypeNumberParagraph(Range range)
    {
        InlineShapes? inlineShapes = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            inlineShapes = range.InlineShapes;
            if (inlineShapes.Count != 0) return false;

            fields = WordFormulaHost.GetLocalFields(range);
            var hasPlaceRef = false;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    hasPlaceRef = true;
            }
            if (!hasPlaceRef) return false;

            var text = (range.Text ?? string.Empty)
                .Replace("\r", string.Empty)
                .Replace("\a", string.Empty)
                .Replace("\t", string.Empty)
                .Replace("\u0013", string.Empty)
                .Replace("\u0014", string.Empty)
                .Replace("\u0015", string.Empty)
                .Trim();
            if (text.Length == 0) return true;

            var sawNumber = false;
            foreach (var character in text)
            {
                if (char.IsWhiteSpace(character)) continue;
                if (char.IsDigit(character))
                {
                    sawNumber = true;
                    continue;
                }
                if ("()[]{}.,:;-–—/\\".IndexOf(character) >= 0) continue;
                return false;
            }
            return sawNumber;
        }
        catch { return false; }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(inlineShapes);
        }
    }

    private static void RestoreMathTypeDisplayParagraphLayout(
        InlineShape shape,
        MathTypeDisplayParagraphLayout layout)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            format = paragraph.Format;
            format.Alignment = layout.Alignment;
            format.LeftIndent = layout.LeftIndent;
            format.RightIndent = layout.RightIndent;
            format.FirstLineIndent = layout.FirstLineIndent;
            format.SpaceBefore = layout.SpaceBefore;
            format.SpaceAfter = layout.SpaceAfter;
            format.LineSpacingRule = layout.LineSpacingRule;
            try { format.LineSpacing = layout.LineSpacing; } catch { }
            format.BaseLineAlignment = layout.BaseLineAlignment;
            format.KeepTogether = layout.KeepTogether;
            format.KeepWithNext = layout.KeepWithNext;
            format.WidowControl = layout.WidowControl;
            format.PageBreakBefore = layout.PageBreakBefore;

            if (layout.SpecialTabStops.Count == 0) return;
            tabStops = format.TabStops;
            foreach (var special in layout.SpecialTabStops)
            {
                var exists = false;
                for (var index = 1; index <= tabStops.Count; index++)
                {
                    Release(tabStop);
                    tabStop = tabStops[index];
                    if (Math.Abs(tabStop.Position - special.Position) <= 0.5f
                        && tabStop.Alignment == special.Alignment
                        && tabStop.Leader == special.Leader)
                    {
                        exists = true;
                        break;
                    }
                }
                if (exists) continue;
                Release(tabStop);
                tabStop = tabStops.Add(special.Position, special.Alignment, special.Leader);
            }
        }
        finally
        {
            Release(tabStop);
            Release(tabStops);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static byte[] ReadMathTypeMtefFromCompoundFile(byte[] compoundFile)
    {
        var equationNative = MathTypeOleStorage.ReadEquationNative(compoundFile);
        if (equationNative.Length < 12)
            throw new InvalidDataException("MathType Equation Native is too short for MTEF extraction.");
        var headerLength = BitConverter.ToUInt16(equationNative, 0);
        var mtefLength = checked((int)BitConverter.ToUInt32(equationNative, 8));
        if (headerLength < 12
            || mtefLength <= 0
            || headerLength + mtefLength > equationNative.Length)
            throw new InvalidDataException("MathType Equation Native contains an invalid MTEF extent.");
        var mtef = new byte[mtefLength];
        Buffer.BlockCopy(equationNative, headerLength, mtef, 0, mtefLength);
        return mtef;
    }

    private static bool TryRenderMathTypeNativePreviewFromCompoundFile(
        byte[] compoundFile,
        string outputDirectory,
        out MathTypeNativePreviewRenderer.Result? result)
    {
        result = null;
        try
        {
            var mtef = ReadMathTypeMtefFromCompoundFile(compoundFile);
            if (!MathTypeNativePreviewRenderer.TryRender(
                    mtef,
                    outputDirectory,
                    out var rendered))
                return false;
            result = rendered;
            return true;
        }
        catch { return false; }
    }

    private static int ResolveVisualTeXInlineEditWordPosition(InlineShape shape)
    {
        var current = ReadInlineOleWordPosition(shape);
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? peers = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return current;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            peers = paragraphRange.InlineShapes;
            var positions = new List<int>();
            for (var index = 1; index <= peers.Count; index++)
            {
                InlineShape? peer = null;
                Range? peerRange = null;
                try
                {
                    peer = peers[index];
                    peerRange = peer.Range;
                    if (peerRange.Start == shapeRange.Start
                        && peerRange.End == shapeRange.End)
                        continue;
                    if (!MathTypeOleInterop.IsMathTypeOle(peer)) continue;
                    positions.Add(ReadInlineOleWordPosition(peer));
                }
                catch (COMException) { }
                finally
                {
                    Release(peerRange);
                    Release(peer);
                }
            }
            if (positions.Count < 2) return current;
            var profile = positions
                .GroupBy(position => position)
                .Select(group => new { Position = group.Key, Count = group.Count() })
                .OrderByDescending(item => item.Count)
                .ThenBy(item => Math.Abs(item.Position))
                .First();
            if (profile.Count < 2 || profile.Count * 2 <= positions.Count)
                return current;
            return profile.Position;
        }
        catch { return current; }
        finally
        {
            Release(peers);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static int ReadInlineOleWordPosition(InlineShape shape)
    {
        Range? shapeRange = null;
        Range? probe = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        Document? document = null;
        try
        {
            shapeRange = shape.Range;
            document = shapeRange.Document;
            // Word exposes an embedded MathType object as an EMBED field.  The
            // InlineShape.Range therefore spans the field instruction plus the
            // single U+0001 object/result character.  Genuine MathType keeps the
            // field instruction at Position=0 and applies the vertical offset only
            // to U+0001.  Reading range.Start therefore returns the wrong baseline.
            for (var position = shapeRange.Start; position < shapeRange.End; position++)
            {
                Release(font);
                font = null;
                Release(probe);
                probe = document.Range(position, position + 1);
                if (!string.Equals(probe.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                font = probe.Font;
                var wordPosition = font.Position;
                if (wordPosition != (int)WdConstants.wdUndefined
                    && wordPosition >= -256
                    && wordPosition <= 256)
                    return wordPosition;
            }
        }
        catch { }
        finally
        {
            Release(font);
            Release(probe);
            Release(shapeRange);
            Release(document);
        }
        return (int)Math.Round(ReadDefinedShapeFontPosition(shape) ?? 0f);
    }

    private static int CalculateMathTypeOleWordPosition(
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline)
    {
        if (!(actualHeightPoints > 0)
            || !(exportedHeight > 0)
            || !exportedBaseline.HasValue
            || float.IsNaN(exportedBaseline.Value)
            || float.IsInfinity(exportedBaseline.Value)
            || exportedBaseline.Value < 0
            || exportedBaseline.Value >= exportedHeight)
            return 0;

        // MathType's Word integration stores the OLE object's character position
        // from the equation baseline toward the bottom of the picture. Keep this
        // calculation local to MathType instead of sharing VisualTeX's ordinary
        // inline-OLE optical alignment code; changes to VisualTeX inline layout must
        // never alter Equation.DSMT4 placement again.
        var baselineFromBottomPoints =
            actualHeightPoints * (exportedHeight - exportedBaseline.Value) / exportedHeight;
        var rounded = Math.Max(
            0,
            (int)Math.Round(baselineFromBottomPoints, MidpointRounding.AwayFromZero));

        // The offline WMF presentation includes Word's one-point character-box
        // allowance. Native MathType samples (ordinary fractions, radicals and
        // mixed inline equations) place the U+0001 object one point above the raw
        // picture descent, while the WMF/dxaOrig dimensions retain the full bound.
        return -Math.Max(0, rounded - 1);
    }

    private static void SetInlineOleWordPosition(InlineShape shape, int position)
    {
        Range? shapeRange = null;
        Range? probe = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        Document? document = null;
        try
        {
            shapeRange = shape.Range;
            document = shapeRange.Document;
            // Match genuine MathType field formatting exactly: EMBED instruction
            // characters remain on the paragraph baseline; only the U+0001 object
            // result receives the equation's vertical offset.  Applying Position to
            // shape.Range shifts the whole field and can leak into following prose.
            font = shapeRange.Font;
            font.Position = 0;
            Release(font);
            font = null;

            var clamped = Math.Max(-256, Math.Min(256, position));
            for (var index = shapeRange.Start; index < shapeRange.End; index++)
            {
                Release(probe);
                probe = document.Range(index, index + 1);
                if (!string.Equals(probe.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                font = probe.Font;
                font.Position = clamped;
                return;
            }
        }
        finally
        {
            Release(font);
            Release(probe);
            Release(shapeRange);
            Release(document);
        }
    }

    private static InlineOlePreviewMetrics? TryMeasureInlineOlePreview(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            var bytes = range.EnhMetaFileBits as byte[];
            if (bytes is null || bytes.Length == 0) return null;
            using var stream = new System.IO.MemoryStream(bytes, writable: false);
            using var metafile = new System.Drawing.Imaging.Metafile(stream);
            return MeasureMetafilePreview(metafile);
        }
        catch { return null; }
        finally { Release(range); }
    }

    private static InlineOlePreviewMetrics? TryMeasureMetafilePreview(string emfPath)
    {
        try
        {
            using var metafile = new System.Drawing.Imaging.Metafile(emfPath);
            return MeasureMetafilePreview(metafile);
        }
        catch { return null; }
    }

    private static InlineOlePreviewMetrics? MeasureMetafilePreview(
        System.Drawing.Imaging.Metafile metafile)
    {
        const int width = 640;
        const int height = 240;
        using var bitmap = new Bitmap(width, height);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            graphics.DrawImage(metafile, 0, 0, width, height);
        }

        var minY = height;
        var maxY = -1;
        for (var y = 0; y < height; y++)
        {
            for (var x = 0; x < width; x++)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R >= 245 && pixel.G >= 245 && pixel.B >= 245) continue;
                minY = Math.Min(minY, y);
                maxY = Math.Max(maxY, y);
            }
        }
        if (maxY < minY) return null;

        var inkHeightRatio = (maxY - minY + 1f) / height;
        var bottomWhitespaceRatio = (height - 1f - maxY) / height;
        if (!(inkHeightRatio > 0.01f)
            || float.IsNaN(inkHeightRatio)
            || float.IsInfinity(inkHeightRatio))
            return null;
        return new InlineOlePreviewMetrics(inkHeightRatio, bottomWhitespaceRatio);
    }

    private static (float Width, float Height) CalculateMathTypeEditedPresentationSize(
        float oldWidth,
        float oldHeight,
        InlineOlePreviewMetrics? sourcePreview,
        InlineOlePreviewMetrics? editedPreview,
        float? newRenderWidth,
        float? newRenderHeight,
        double? originalRenderWidth,
        double? originalRenderHeight,
        double? originalFontSizePt,
        double? originalRenderFontSizePt)
    {
        var fallback = OfficeFormulaSizing.EditedSize(
            oldWidth,
            oldHeight,
            originalRenderWidth,
            originalRenderHeight,
            newRenderWidth ?? oldWidth / 0.75f,
            newRenderHeight ?? oldHeight / 0.75f,
            originalFontSizePt: originalFontSizePt,
            originalRenderFontSizePt: originalRenderFontSizePt);
        if (!sourcePreview.HasValue || !editedPreview.HasValue
            || !(oldHeight > 0)
            || !(sourcePreview.Value.InkHeightRatio > 0.01f)
            || !(editedPreview.Value.InkHeightRatio > 0.01f))
            return fallback;

        var height = oldHeight
            * sourcePreview.Value.InkHeightRatio
            / editedPreview.Value.InkHeightRatio;
        if (!(height > 0) || float.IsNaN(height) || float.IsInfinity(height))
            return fallback;

        var aspect = newRenderWidth is > 0 && newRenderHeight is > 0
            ? newRenderWidth.Value / newRenderHeight.Value
            : fallback.Width / Math.Max(0.01f, fallback.Height);
        if (!(aspect > 0) || float.IsNaN(aspect) || float.IsInfinity(aspect))
            return fallback;

        // Preserve the native MathType glyph scale, not the outer OLE box. A
        // MathType preview typically contains appreciable ascent/descent padding,
        // whereas VisualTeX's EMF is tightly cropped. Keeping the same outer
        // height would therefore enlarge the visible glyphs by 20–40 percent.
        height = Math.Max(1f, Math.Min(oldHeight * 4f, Math.Max(oldHeight * 0.25f, height)));
        return (Math.Max(1f, height * aspect), height);
    }

    private static void ResetShapeFontPosition(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            ResetRangeFontPosition(range);
        }
        finally { Release(range); }
    }

    private static void ResetDisplayFormulaPosition(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            ResetDisplayFormulaPosition(range);
        }
        finally { Release(range); }
    }

    private static void ResetDisplayFormulaPosition(Range formulaRange)
    {
        ResetRangeFontPosition(formulaRange);
        ResetParagraphTypingPosition(formulaRange);
    }

    private static void NormalizeFollowingInlineProseBaseline(Range formulaRange)
    {
        Document? document = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? precedingWindow = null;
        Range? preceding = null;
        Range? trailingHostWindow = null;
        Range? trailing = null;
        InlineShapes? shapes = null;
        InlineShape? nextShape = null;
        Range? nextShapeRange = null;
        OMaths? maths = null;
        OMath? nextMath = null;
        Range? nextMathRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            document = formulaRange.Document;
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var paragraphBodyEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);

            // Only the nearby prose can determine the insertion baseline. Capturing
            // every InlineShape/OMath in a paragraph made appending the Nth inline
            // formula rescan formulas 1..N-1 twice. Keep the format probe bounded;
            // hidden OLE field-code ranges inside this small window are still
            // excluded by the same host-aware logic used previously.
            const int maximumProbeCharacters = 256;
            var precedingStart = Math.Max(
                paragraphRange.Start,
                formulaRange.Start - maximumProbeCharacters);
            if (formulaRange.Start > precedingStart)
            {
                precedingWindow = document.Range(precedingStart, formulaRange.Start);
                var excludedHostRanges = CaptureNonProseHostRanges(precedingWindow);
                preceding = FindOrdinaryVisibleCharacterRange(
                    document,
                    formulaRange.Start - 1,
                    precedingStart,
                    step: -1,
                    excludedHostRanges);
            }

            var targetPosition = 0;
            if (preceding is not null)
            {
                font = preceding.Font;
                var precedingPosition = font.Position;
                if (precedingPosition != (int)WdConstants.wdUndefined
                    && precedingPosition >= -256
                    && precedingPosition <= 256)
                    targetPosition = precedingPosition;
                Release(font);
                font = null;
            }

            var trailingEnd = Math.Max(formulaRange.End, paragraphBodyEnd);
            if (trailingEnd <= formulaRange.End) return;
            trailingHostWindow = document.Range(formulaRange.End, trailingEnd);

            // Range collections are in document order. The first following host is
            // sufficient; do not enumerate every object in the paragraph.
            shapes = trailingHostWindow.InlineShapes;
            if (shapes.Count > 0)
            {
                nextShape = shapes[1];
                nextShapeRange = nextShape.Range;
                trailingEnd = Math.Min(trailingEnd, nextShapeRange.Start);
            }
            maths = trailingHostWindow.OMaths;
            if (maths.Count > 0)
            {
                nextMath = maths[1];
                nextMathRange = nextMath.Range;
                trailingEnd = Math.Min(trailingEnd, nextMathRange.Start);
            }

            if (trailingEnd <= formulaRange.End) return;
            trailing = document.Range(formulaRange.End, trailingEnd);
            if (!ContainsVisibleBodyText(trailing.Text)) return;
            font = trailing.Font;
            var currentPosition = font.Position;
            if (currentPosition == (int)WdConstants.wdUndefined
                || currentPosition != targetPosition)
                font.Position = targetPosition;
        }
        catch
        {
            // Baseline repair is best-effort and must not interrupt Word input.
        }
        finally
        {
            Release(font);
            Release(nextMathRange);
            Release(nextMath);
            Release(maths);
            Release(nextShapeRange);
            Release(nextShape);
            Release(shapes);
            Release(trailing);
            Release(trailingHostWindow);
            Release(preceding);
            Release(precedingWindow);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(document);
        }
    }

    private static void ResetParagraphTypingPosition(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Range? nextCharacter = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.End > paragraphRange.Start)
            {
                paragraphMark = paragraphRange.Duplicate;
                paragraphMark.SetRange(paragraphRange.End - 1, paragraphRange.End);
                ResetRangeFontPosition(paragraphMark);
            }

            if (formulaRange.End >= paragraphRange.End) return;
            nextCharacter = paragraphRange.Duplicate;
            nextCharacter.SetRange(
                formulaRange.End,
                Math.Min(formulaRange.End + 1, paragraphRange.End));
            if (nextCharacter.Text is "\v" or "\r" or "\n")
                ResetRangeFontPosition(nextCharacter);
        }
        catch
        {
            // Baseline restoration is best-effort and must not invalidate the
            // formula that has already been inserted or resized.
        }
        finally
        {
            Release(nextCharacter);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static string InlineBaselineBookmarkName(string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var parsed))
            throw new InvalidDataException("VisualTeX formulaId must be a UUID.");
        return InlineBaselineBookmarkPrefix + parsed.ToString("N");
    }


    private static bool IsKnownInlineBaselineSentinel(string? text) =>
        string.IsNullOrEmpty(text)
        || string.Equals(text, InlineOleTypingAnchor, StringComparison.Ordinal)
        || string.Equals(text, InlineBaselineSentinel, StringComparison.Ordinal)
        || string.Equals(text, LegacyInlineBaselineSentinel, StringComparison.Ordinal)
        || string.Equals(
            text,
            LegacyInlineNonbreakingBaselineSentinel,
            StringComparison.Ordinal);

    private static bool IsHiddenTextRange(Range range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            return font.Hidden != 0;
        }
        catch { return false; }
        finally { Release(font); }
    }

    private static bool RangeContainsMath(Range range)
    {
        OMaths? maths = null;
        try
        {
            maths = range.OMaths;
            return maths.Count > 0;
        }
        catch { return false; }
        finally { Release(maths); }
    }

    private static void RemoveInlineOleTypingAnchorAfter(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            RemoveInlineOleTypingAnchorAfter(range);
        }
        finally { Release(range); }
    }

    private static void RemoveInlineOleTypingAnchorAfter(Range formulaRange)
    {
        Document? document = null;
        Range? content = null;
        Range? anchor = null;
        try
        {
            document = formulaRange.Document;
            content = document.Content;
            var position = formulaRange.End;
            var contentStart = content.Start;
            var contentEnd = content.End;
            if (position < contentStart || position >= contentEnd) return;
            anchor = document.Range(position, Math.Min(position + 1, contentEnd));
            if (!string.Equals(
                    anchor.Text,
                    InlineOleTypingAnchor,
                    StringComparison.Ordinal))
                return;
            // Word can report an ordinary character immediately after an OLE as
            // math-affiliated because the adjacent object participates in the
            // same layout run. This helper is called only for a confirmed
            // VisualTeX inline OLE, whose first U+200C at Range.End is owned by us.
            anchor.Delete();
        }
        catch
        {
            // The formula may have been deleted by Word between range capture and
            // cleanup. Orphan-anchor removal is best-effort in rollback paths.
        }
        finally
        {
            Release(anchor);
            Release(content);
            Release(document);
        }
    }

    private static void RemoveInlineBaselineSentinel(
        Document document, string formulaId, Range? protectedFollowingText = null)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? sentinel = null;
        try
        {
            var name = InlineBaselineBookmarkName(formulaId);
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            sentinel = bookmark.Range;
            var sentinelStart = sentinel.Start;
            var ownedBoundaryHasWidth = sentinel.Start < sentinel.End
                && IsKnownInlineBaselineSentinel(sentinel.Text)
                && (protectedFollowingText is null || sentinel.End <= protectedFollowingText.Start);
            bookmark.Delete();
            if (ownedBoundaryHasWidth)
            {
                // Range.Delete at an OMath boundary can apply Word's smart
                // whitespace deletion and consume the next user-authored space.
                // Empty-text replacement addresses only the captured guard span.
                if (protectedFollowingText is not null) sentinel.Text = string.Empty;
                else sentinel.Delete();
            }

            // Remove a temporary guard from interrupted insertions. Probe both
            // sides of the former bookmark because deleting the bookmarked marker
            // shifts the following text one position to the left.
            TryDeleteTemporaryInlineBoundaryAt(document, sentinelStart - 1, protectedFollowingText);
            TryDeleteTemporaryInlineBoundaryAt(document, sentinelStart, protectedFollowingText);
        }
        catch
        {
            // A stale or externally edited sentinel must never block formula work.
        }
        finally
        {
            Release(sentinel);
            Release(bookmark);
            Release(bookmarks);
        }
    }


    private static bool TryDeleteTemporaryInlineBoundaryAt(
        Document document,
        int position,
        Range? protectedFollowingText = null)
    {
        Range? candidate = null;
        try
        {
            if (protectedFollowingText is not null && position >= protectedFollowingText.Start)
                return false;
            var contentStart = document.Content.Start;
            var contentEnd = document.Content.End;
            if (position < contentStart || position >= contentEnd) return false;
            object candidateStart = position;
            object candidateEnd = Math.Min(contentEnd, position + 1);
            candidate = document.Range(ref candidateStart, ref candidateEnd);
            var text = candidate.Text;
            var removable =
                string.Equals(text, BulkInlineFormulaPlaceholder, StringComparison.Ordinal)
                || string.Equals(text, LegacyInlineMathGuard, StringComparison.Ordinal)
                || string.Equals(text, LegacyInlineBaselineSentinel, StringComparison.Ordinal)
                || string.Equals(text, InlineMathGuard, StringComparison.Ordinal)
                    && IsHiddenTextRange(candidate);
            // U+00A0 is also the correct Word representation of explicit LaTeX
            // spacing (for example `~` and `\ `). Delete it only when the VTBL
            // bookmark itself owns that legacy marker, never by proximity.
            if (!removable || RangeContainsMath(candidate)) return false;
            if (protectedFollowingText is not null) candidate.Text = string.Empty;
            else candidate.Delete();
            return true;
        }
        catch { return false; }
        finally { Release(candidate); }
    }

    private static Range? FindInlineTypingFormatSource(
        Range formulaRange,
        Range typingAnchor)
    {
        Document? document = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? precedingWindow = null;
        Range? followingWindow = null;
        try
        {
            document = formulaRange.Document;
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var paragraphBodyEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            const int maximumProbeCharacters = 256;

            // The prose before an inline formula is the authoritative typing
            // style. Restrict host discovery to the same bounded region that can
            // actually be probed; scanning every OLE in a long formula paragraph
            // made caret restoration grow linearly with formula count.
            var precedingStart = Math.Max(
                paragraphRange.Start,
                formulaRange.Start - maximumProbeCharacters);
            if (formulaRange.Start > precedingStart)
            {
                precedingWindow = document.Range(precedingStart, formulaRange.Start);
                var precedingHosts = CaptureNonProseHostRanges(precedingWindow);
                var preceding = FindOrdinaryVisibleCharacterRange(
                    document,
                    formulaRange.Start - 1,
                    precedingStart,
                    step: -1,
                    precedingHosts);
                if (preceding is not null) return preceding;
            }

            // Paragraph-leading formulas have no preceding prose. In that case,
            // inherit from the first nearby ordinary character after the formula.
            var followingStart = Math.Max(typingAnchor.End, formulaRange.End);
            var followingEnd = Math.Min(
                paragraphBodyEnd,
                followingStart + maximumProbeCharacters);
            if (followingEnd <= followingStart) return null;
            followingWindow = document.Range(followingStart, followingEnd);
            var followingHosts = CaptureNonProseHostRanges(followingWindow);
            return FindOrdinaryVisibleCharacterRange(
                document,
                followingStart,
                followingEnd,
                step: 1,
                followingHosts);
        }
        finally
        {
            Release(followingWindow);
            Release(precedingWindow);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(document);
        }
    }

    private static Range? FindOrdinaryVisibleCharacterRange(
        Document document,
        int startPosition,
        int boundaryPosition,
        int step,
        IReadOnlyList<NonProseHostRange> excludedHostRanges)
    {
        if (step is not (-1 or 1)) return null;
        const int maximumProbeCharacters = 256;
        var position = startPosition;
        for (var probeIndex = 0;
             probeIndex < maximumProbeCharacters
             && (step > 0
                 ? position < boundaryPosition
                 : position >= boundaryPosition);
             probeIndex++, position += step)
        {
            Range? probe = null;
            try
            {
                if (position < 0) break;
                if (excludedHostRanges.Any(host => host.Contains(position)))
                    continue;
                probe = document.Range(position, position + 1);
                if (!ContainsVisibleBodyText(probe.Text)) continue;
                var result = probe.Duplicate;
                return result;
            }
            catch (COMException)
            {
                // Keep probing nearby ordinary prose.
            }
            finally
            {
                Release(probe);
            }
        }
        return null;
    }

    private static List<NonProseHostRange> CaptureNonProseHostRanges(
        Range paragraphRange)
    {
        var result = new List<NonProseHostRange>();
        InlineShapes? shapes = null;
        OMaths? maths = null;
        try
        {
            shapes = paragraphRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                try
                {
                    shape = shapes[index];
                    range = shape.Range;
                    result.Add(new NonProseHostRange(range.Start, range.End));
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }

            maths = paragraphRange.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    result.Add(new NonProseHostRange(range.Start, range.End));
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
            return result;
        }
        finally
        {
            Release(maths);
            Release(shapes);
        }
    }

    private static void ResetRangeFontPosition(Range range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            // A redundant format assignment still dirties the Word run and can
            // repaginate the paragraph. Healthy OMML boundaries are normally
            // already at the neutral baseline.
            if (font.Position != 0)
                font.Position = 0;
        }
        finally { Release(font); }
    }

    private static void ApplyInlineBaseline(
        InlineShape shape,
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline,
        double semanticFontSizePoints) =>
        ApplyInlineBaseline(
            shape,
            actualHeightPoints,
            exportedHeight,
            exportedBaseline,
            existingFontPosition: null,
            sourceSemanticFontSizePoints: semanticFontSizePoints,
            targetSemanticFontSizePoints: semanticFontSizePoints);

    private static void ApplyInlineBaseline(
        InlineShape shape,
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline,
        float? existingFontPosition,
        double sourceSemanticFontSizePoints,
        double targetSemanticFontSizePoints)
    {
        var position = CalculateVisualTeXInlinePosition(
            shape,
            actualHeightPoints,
            exportedHeight,
            exportedBaseline,
            existingFontPosition,
            sourceSemanticFontSizePoints,
            targetSemanticFontSizePoints);

        // Word stores an InlineShape OLE as an EMBED field whose instruction and
        // U+0001 object/result character share shape.Range. Applying Position to
        // the complete field lowers hidden instruction characters as well, which
        // inflates the line box and can leak the negative position into following
        // prose. Match Word/MathType behavior: keep the field instruction at the
        // paragraph baseline and move only the painted object character.
        SetInlineOleWordPosition(shape, position);
    }

    private void RestoreTypingBaselineAfterMathTypeConversion(InlineShape shape)
    {
        Range? range = null;
        Selection? selection = null;
        try
        {
            range = shape.Range;
            // Format conversion replaces an existing inline equation at the same
            // text boundary. Its surrounding prose already owns the correct
            // character baseline, so do not rescan/rewrite the entire following
            // prose segment here. Re-establish only the paragraph end and the
            // collapsed insertion format immediately after the new MathType OLE.
            ResetParagraphTypingPosition(range);
            selection = _application.Selection;
            selection.SetRange(range.End, range.End);
            ApplyInlineTypingFormattingToSelection(selection, range);
        }
        catch
        {
            // The MathType OLE is already structurally valid. A transient Word
            // insertion-format refusal must not invalidate the conversion.
        }
        finally
        {
            Release(selection);
            Release(range);
        }
    }

    private void RestoreTypingBaselineAfter(
        InlineShape shape,
        bool ensureTypingAnchor = false)
    {
        _ = ensureTypingAnchor;
        Range? range = null;
        try
        {
            range = shape.Range;
            NormalizeFollowingInlineProseBaseline(range);
            ResetParagraphTypingPosition(range);
            WordFormulaHostLayout.RestoreInlineCaret(
                _application,
                range);
        }
        finally
        {
            Release(range);
        }
    }

    private void RestoreTypingCaretAt(
        Document document,
        int caretPosition,
        Range formulaRange)
    {
        Range? content = null;
        Selection? selection = null;
        try
        {
            content = document.Content;
            var safePosition = Math.Max(
                content.Start,
                Math.Min(caretPosition, content.End));
            selection = _application.Selection;
            selection.SetRange(safePosition, safePosition);
            ApplyInlineTypingFormattingToSelection(selection, formulaRange);
        }
        catch
        {
            // Structural caret placement remains useful even if Word rejects a
            // transient insertion-format mutation at an unusual protected range.
        }
        finally
        {
            Release(selection);
            Release(content);
        }
    }

    private void RestoreTypingBaselineAfter(Range formulaRange) =>
        RestoreTypingBaselineAfter(formulaRange, null);

    private void RestoreTypingBaselineAfter(Range formulaRange, int? caretPosition)
    {
        Range? caret = null;
        Selection? selection = null;
        try
        {
            ResetParagraphTypingPosition(formulaRange);
            caret = formulaRange.Duplicate;
            if (caretPosition.HasValue)
                caret.SetRange(caretPosition.Value, caretPosition.Value);
            else
                caret.Collapse(WdCollapseDirection.wdCollapseEnd);

            selection = _application.Selection;
            selection.SetRange(caret.Start, caret.End);
            ApplyInlineTypingFormattingToSelection(selection, formulaRange);
        }
        catch
        {
            // Keep the caret outside the formula even if Word refuses to mutate
            // insertion formatting at an unusual protected boundary.
        }
        finally
        {
            Release(selection);
            Release(caret);
        }
    }

    private static OfficeObjectResult Result(
        OfficeSessionDocument session,
        Document document,
        bool documentIdentityAlreadyValidated = false) =>
        new()
        {
            FormulaId = session.FormulaId,
            DocumentId = documentIdentityAlreadyValidated
                && !string.IsNullOrWhiteSpace(session.SourceDocumentId)
                    ? session.SourceDocumentId!
                    : DocumentIdentity(document),
            ObjectId = session.FormulaId,
        };

    private static string RangeReference(Range range) =>
        $"{RangeReferencePrefix}{range.Start}:{range.End}";

    private static Range ResolveSessionInsertionRange(
        Document document,
        OfficeSessionDocument session,
        Selection selection,
        bool preserveCapturedInsertion = false)
    {
        var sourceRange = ResolveSourceRange(
            document,
            session.SourceObjectId,
            selection);
        // A new insertion may deliberately leave a generated numbering table.
        // In-place conversion/redraw owns the just-cleared source coordinate;
        // translating it as a new user insertion can move it out of its cell.
        if (preserveCapturedInsertion
            || !string.Equals(session.Mode, "create", StringComparison.OrdinalIgnoreCase))
            return sourceRange;
        try
        {
            return ResolveCreateInsertionRange(document, sourceRange);
        }
        finally { Release(sourceRange); }
    }

    private static Range ResolveCreateInsertionRange(
        Document document,
        Range sourceRange)
    {
        Tables? tables = null;
        Table? table = null;
        Range? safeTypingRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Range? captionRange = null;
        Range? content = null;
        try
        {
            if (!WordEquationNumbering.RangeIsWhollyWithinTable(sourceRange))
                return sourceRange.Duplicate;
            tables = sourceRange.Tables;
            if (tables.Count == 0) return sourceRange.Duplicate;
            table = tables[1];
            var formulaId = TryGetNumberedFormulaId(document, table, sourceRange);
            if (string.IsNullOrWhiteSpace(formulaId))
                return sourceRange.Duplicate;

            // When the captured caret belongs to a legacy VisualTeX numbered table,
            // place a new formula after that formula's dedicated native SEQ paragraph.
            // In Office 2019, using the live Selection at commit time can otherwise
            // insert new content between the old table and caption, reordering the
            // old number and corrupting adjacent cells.
            safeTypingRange =
                WordEquationNumbering.EnsureNormalTypingParagraphAfterNumberedDisplay(
                    document,
                    formulaId!);
            if (safeTypingRange is not null)
            {
                var result = safeTypingRange;
                safeTypingRange = null;
                return result;
            }

            bookmarks = document.Bookmarks;
            var captionName = WordEquationNumbering.NativeCaptionBookmarkName(
                formulaId!);
            if (!bookmarks.Exists(captionName))
                return sourceRange.Duplicate;
            captionBookmark = bookmarks[captionName];
            captionRange = captionBookmark.Range;
            content = document.Content;
            var lastInsertPosition = Math.Max(content.Start, content.End - 1);
            var position = Math.Max(
                content.Start,
                Math.Min(captionRange.End, lastInsertPosition));
            return document.Range(position, position);
        }
        catch
        {
            // Falling back to the captured range is safer than consulting the
            // mutable live Selection again. Normal document tables must retain
            // their existing in-cell insertion behavior.
            return sourceRange.Duplicate;
        }
        finally
        {
            Release(content);
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
            Release(safeTypingRange);
            Release(table);
            Release(tables);
        }
    }

    private static string? TryGetNumberedFormulaId(
        Document document,
        Table table,
        Range insertion)
    {
        Columns? columns = null;
        Cell? insertionCell = null;
        Cell? centerCell = null;
        Range? center = null;
        Range? tableRange = null;
        try
        {
            columns = table.Columns;
            if (columns.Count != 3)
                return null;

            insertionCell =
                WordFormulaHost.TryGetOwningCell(
                    insertion);
            if (insertionCell is null)
                return null;

            centerCell =
                table.Cell(
                    insertionCell.RowIndex,
                    2);
            center = centerCell.Range.Duplicate;

            WordFormulaHostDescriptor? host = null;
            try
            {
                host =
                    WordFormulaHostResolver.ResolveLocal(
                        document,
                        center,
                        WordFormulaHostKind.VisualTeX);
            }
            catch (InvalidDataException) { }

            if (host is null)
            {
                try
                {
                    host =
                        WordFormulaHostResolver.ResolveLocal(
                            document,
                            center,
                            WordFormulaHostKind.Omml);
                }
                catch (InvalidDataException) { }
            }
            if (host is null)
                return null;

            var numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            if (!numbering.Numbered
                || numbering.ContainerKind !=
                    WordFormulaNumberingContainerKind.CanonicalBodyTable
                || numbering.ContainerRange is null)
                return null;

            tableRange = table.Range.Duplicate;
            var owner = numbering.ContainerRange;
            if (owner.StoryType !=
                    tableRange.StoryType
                || owner.Start !=
                    tableRange.Start
                || owner.End !=
                    tableRange.End)
                return null;

            return numbering.FormulaId
                ?? host.FormulaId;
        }
        finally
        {
            Release(tableRange);
            Release(center);
            Release(centerCell);
            Release(insertionCell);
            Release(columns);
        }
    }

    private static Range ResolveSourceRange(
        Document document,
        string? sourceObjectId,
        Selection selection)
    {
        if (!TryParseRangeReference(sourceObjectId, out var start, out var end))
            return selection.Range.Duplicate;
        Range? content = null;
        try
        {
            content = document.Content;
            if (start < 0 || end < start || end > content.End)
                throw new InvalidOperationException(
                    "The Word insertion range selected when the formula editor opened is no longer valid.");
            object startValue = start;
            object endValue = end;
            return document.Range(ref startValue, ref endValue);
        }
        finally { Release(content); }
    }

    private static bool TryParseRangeReference(
        string? value,
        out int start,
        out int end)
    {
        start = 0;
        end = 0;
        if (string.IsNullOrWhiteSpace(value)) return false;
        var reference = value!;
        if (!reference.StartsWith(RangeReferencePrefix, StringComparison.Ordinal))
            return false;
        var payload = reference.Substring(RangeReferencePrefix.Length);
        var separator = payload.IndexOf(':');
        if (separator <= 0 || separator >= payload.Length - 1) return false;
        return int.TryParse(payload.Substring(0, separator), out start)
            && int.TryParse(payload.Substring(separator + 1), out end);
    }

    private static void EnsureSourceDocument(
        Document document,
        string? expectedIdentity)
    {
        if (string.IsNullOrWhiteSpace(expectedIdentity)) return;
        var actual = DocumentIdentity(document);
        if (!string.Equals(actual, expectedIdentity, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "The active Word document changed while the VisualTeX editor was open.");
    }

    private static string DocumentIdentity(Document document)
    {
        try
        {
            var fullName = ReadWordStateWithRetry(() => document.FullName);
            if (!string.IsNullOrWhiteSpace(fullName)) return fullName;
        }
        catch { }
        return ReadWordStateWithRetry(() => document.Name);
    }

    private static void EnsureWritable(Document document)
    {
        if (ReadWordStateWithRetry(() => document.ReadOnly))
            throw new UnauthorizedAccessException("The active Word document is read-only.");
    }

    private static void EnsureEquationFieldResultsVisible(Document document)
    {
        Window? window = null;
        Microsoft.Office.Interop.Word.View? view = null;
        try
        {
            window = document.ActiveWindow;
            if (window is null) return;
            view = window.View;
            if (!view.ShowFieldCodes) return;

            // VisualTeX equation numbers and references are genuine Word fields.
            // When a Word window is left in Alt+F9/ShowFieldCodes mode, Word expands
            // the complete SEQ/MACROBUTTON instruction inside the numbered row,
            // making a healthy formula appear corrupted and stretching the layout.
            // Per-field ShowCodes=false cannot override the window-level setting,
            // so user-initiated VisualTeX numbering operations normalize the active
            // view back to rendered field results. The user can still press Alt+F9
            // afterwards if they intentionally want to inspect field instructions.
            view.ShowFieldCodes = false;
        }
        catch
        {
            // Hidden/protected automation windows can reject View mutations. The
            // formula operation itself remains valid; only interactive presentation
            // normalization is best-effort in that environment.
        }
        finally
        {
            Release(view);
            Release(window);
        }
    }

    private static T ReadWordStateWithRetry<T>(Func<T> read)
    {
        const int rpcCallRejected = unchecked((int)0x80010001);
        const int rpcServerCallRetryLater = unchecked((int)0x8001010A);
        const int officeBusy = unchecked((int)0x800AC472);
        const int maximumAttempts = 40;
        for (var attempt = 0; ; attempt++)
        {
            try { return read(); }
            catch (COMException error)
                when ((error.HResult == rpcCallRejected
                        || error.HResult == rpcServerCallRetryLater
                        || error.HResult == officeBusy)
                    && attempt < maximumAttempts - 1)
            {
                // Word can reject harmless state reads for a few UI turns while a
                // just-created document, OLE server or imported OMath is settling.
                // Retry only the idempotent property read; callers still execute
                // every document mutation exactly once.
                System.Threading.Thread.Sleep(50);
            }
        }
    }

    private static bool HasLeadingTab(Document document, Range formulaRange)
    {
        if (formulaRange.Start <= 0) return false;
        Range? preceding = null;
        try
        {
            object start = formulaRange.Start - 1;
            object end = formulaRange.Start;
            preceding = document.Range(ref start, ref end);
            if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal)) return true;
            if (!string.Equals(preceding.Text, "\v", StringComparison.Ordinal)
                || formulaRange.Start <= 1)
                return false;
            preceding.SetRange(formulaRange.Start - 2, formulaRange.Start - 1);
            return string.Equals(preceding.Text, "\t", StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
        finally { Release(preceding); }
    }

    private static void NormalizeNumberedDisplayCell(Range formulaRange)
    {
        Document? document = null;
        Table? table = null;
        Columns? columns = null;
        Cell? centerCell = null;
        Range? cellRange = null;
        Range? character = null;
        try
        {
            if (!(bool)formulaRange.get_Information(WdInformation.wdWithInTable)
                || formulaRange.Tables.Count == 0)
                return;
            document = formulaRange.Document;
            table = formulaRange.Tables[1];
            columns = table.Columns;
            if (columns.Count != 3) return;
            centerCell = table.Cell(
                WordEquationNumbering.GetManagedNumberTableRowIndex(table, formulaRange, 2), 2);
            cellRange = centerCell.Range;

            // A display OMath inserted next to the source OLE can leave one
            // manual line break on each side. Delete only those exact control
            // characters, scanning backwards so Word's shifting ranges cannot
            // expand across and remove the replacement formula object.
            for (var position = cellRange.End - 2;
                 position >= cellRange.Start;
                 position--)
            {
                if (position >= formulaRange.Start
                    && position < formulaRange.End)
                    continue;
                object characterStart = position;
                object characterEnd = position + 1;
                character = document.Range(
                    ref characterStart,
                    ref characterEnd);
                if (string.Equals(character.Text, "\v", StringComparison.Ordinal))
                    character.Delete();
                Release(character);
                character = null;
            }
        }
        finally
        {
            Release(character);
            Release(cellRange);
            Release(centerCell);
            Release(columns);
            Release(table);
            Release(document);
        }
    }

    private static void NormalizeNumberedDisplayCell(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            NormalizeNumberedDisplayCell(range);
        }
        finally { Release(range); }
    }

    private static Range RestoreOmmlReplacementRollback(
        Document document,
        int position,
        string wordOpenXml)
    {
        Range? content = null;
        Range? insertion = null;
        Range? probe = null;
        OMaths? maths = null;
        Range? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            content = document.Content;
            var safePosition = Math.Max(content.Start, Math.Min(position, content.End));
            object insertionStart = safePosition;
            object insertionEnd = Math.Min(
                content.End,
                safePosition + BulkInlineFormulaPlaceholder.Length);
            insertion = document.Range(ref insertionStart, ref insertionEnd);
            if (string.Equals(
                    insertion.Text,
                    BulkInlineFormulaPlaceholder,
                    StringComparison.Ordinal))
                insertion.Text = string.Empty;
            insertion.SetRange(safePosition, safePosition);
            insertion.InsertXML(wordOpenXml);

            object probeStart = Math.Max(content.Start, safePosition - 1);
            object probeEnd = Math.Min(document.Content.End, safePosition + 8);
            probe = document.Range(ref probeStart, ref probeEnd);
            maths = probe.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    var distance = Math.Abs(range.Start - safePosition);
                    if (distance >= bestDistance) continue;
                    Release(best);
                    best = range.Duplicate;
                    bestDistance = distance;
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
            if (best is null)
                throw new InvalidOperationException(
                    "Word could not restore the original OMML equation after a failed replacement.");
            var result = best;
            best = null;
            return result;
        }
        finally
        {
            Release(best);
            Release(maths);
            Release(probe);
            Release(insertion);
            Release(content);
        }
    }

    private static void ValidateInsertedOmmlLiveRange(Range equationRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? range = null;
        try
        {
            maths = equationRange.OMaths;
            if (maths.Count != 1)
                throw new InvalidOperationException("Word did not create exactly one native OMML equation.");
            math = maths[1];
            range = math.Range;
            if (range.End <= range.Start || range.Start != equationRange.Start || range.End != equationRange.End)
                throw new InvalidOperationException("Word returned an invalid native OMML range.");
            // Structural/semantic XML validation is mandatory after the complete
            // local content transaction, before its metadata is accepted.
        }
        finally { Release(range); Release(math); Release(maths); }
    }

    private static void ValidateInsertedOmml(Range equationRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        try
        {
            maths = equationRange.OMaths;
            if (maths.Count != 1)
                throw new InvalidOperationException(
                    "Word did not create exactly one native OMML equation.");
            math = maths[1];
            mathRange = math.Range;
            var wordOpenXml = mathRange.WordOpenXML;
            if (mathRange.End <= mathRange.Start
                || string.IsNullOrWhiteSpace(wordOpenXml)
                || wordOpenXml.IndexOf("oMath", StringComparison.Ordinal) < 0)
                throw new InvalidOperationException(
                    "Word returned an empty native OMML equation.");
            // Validate the structure after Word has imported and normalized the
            // OMML. Pre-insertion XML checks cannot catch empty slots that Word
            // introduces while materializing the native equation tree.
            WordOmmlConverter.ValidateMaterializedOmml(wordOpenXml);
        }
        finally
        {
            Release(mathRange);
            Release(math);
            Release(maths);
        }
    }

    private static void TryDelete(InlineShape? shape)
    {
        if (shape is null) return;
        try { shape.Delete(); } catch { }
    }

    private static void TryDelete(Table? table)
    {
        if (table is null) return;
        try { table.Delete(); } catch { }
    }

    private static void TryDelete(Range? range)
    {
        if (range is null) return;
        try { range.Delete(); } catch { }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        // Office may return the same RCW to the host and to this service.
        // FinalReleaseComObject would invalidate every shared reference in the
        // add-in AppDomain, so release only the reference acquired here.
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
