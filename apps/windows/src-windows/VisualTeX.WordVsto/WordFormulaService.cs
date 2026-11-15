using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using static VisualTeX.WordVsto.WordBulkInsertion;
using static VisualTeX.WordVsto.WordDisplayInsertion;
using static VisualTeX.WordVsto.WordFormulaHostSetup;
using static VisualTeX.WordVsto.WordFormulaLatexSource;
using static VisualTeX.WordVsto.WordFormulaOle;
using static VisualTeX.WordVsto.WordInlineBaseline;
using static VisualTeX.WordVsto.WordInlineOlePosition;
using static VisualTeX.WordVsto.WordMathTypeDisplayLayout;
using static VisualTeX.WordVsto.WordMathTypeNumbering;
using static VisualTeX.WordVsto.WordMathTypeOleLookup;
using static VisualTeX.WordVsto.WordMathTypePresentation;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    private const string RangeReferencePrefix = "visualtex-word-vsto-range:";
    internal const string InlineBaselineBookmarkPrefix = "VTBL_";
    internal const string InlineLatexBaselineBookmarkPrefix = "VTLB_";
    // Ordinary spaces are used only while Word materializes an inline OMath and
    // are deleted immediately afterwards. Inline OLE formulas keep a real zero-
    // width non-joiner after the object so keyboard input inherits an ordinary
    // text run instead of the OLE object's negative baseline offset.
    internal const string InlineMathGuard = " ";
    internal const string InlineBaselineSentinel = " ";
    internal const string InlineOleTypingAnchor = "\u200C";
    internal const string LegacyInlineMathGuard = "\u200B";
    internal const string LegacyInlineBaselineSentinel = "\u2060";
    internal const string LegacyInlineNonbreakingBaselineSentinel = "\u00A0";
    internal const string BulkInlineFormulaPlaceholder = "\uE000";
    internal const float ParagraphBeforeOleDisplaySpaceAfterPoints = 0f;
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
        // Length changes of a completed batch since the capture. The restored
        // selection follows them, so it covers the same text afterwards.
        internal WordCoordinateShifts? Shifts { get; set; }
    }

    internal sealed class LatexRedrawSourceContext
    {
        internal bool HasVisibleSurroundingText { get; set; }
        internal int FontContextRelativePosition { get; set; } = -1;
    }

    internal sealed class FormulaToLatexTarget
    {
        internal int Ordinal { get; set; }
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

    internal readonly struct NonProseHostRange
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

            // An exact native OMath selection is common when the Ribbon edit
            // command is invoked. Detect it before probing both OLE families:
            // those negative InlineShapes/MathType probes are pure overhead for
            // OMML and are especially visible on the first edit in a document.
            if (SelectionContainsLocalOmml(range))
            {
                var directOmml = TryReadLocalOmmlSelection(document, range);
                if (directOmml is not null)
                    return directOmml;
            }

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

            var ommlSelection =
                TryReadLocalOmmlSelection(document, range);
            if (ommlSelection is not null)
                return ommlSelection;

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

    private static bool SelectionContainsLocalOmml(
        Range selectionRange)
    {
        OMaths? maths = null;
        try
        {
            maths = selectionRange.OMaths;
            return maths.Count > 0;
        }
        catch (COMException)
        {
            return false;
        }
        finally
        {
            Release(maths);
        }
    }

    private OfficeSelection? TryReadLocalOmmlSelection(
        Document document,
        Range selectionRange)
    {
        var omml = WordFormulaHostLookup.ResolveLocal(
            document,
            selectionRange,
            WordFormulaHostKind.Omml);
        if (omml is null)
            return null;

        // ResolveLocalOmml already proves the exact native numbering topology.
        // Re-running WordFormulaNumberingResolver here used to repeat the same
        // local COM walk for every display OMML edit.
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
            try { probe.MoveStart(WdUnits.wdCharacter, -2); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            try { probe.MoveEnd(WdUnits.wdCharacter, 2); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
                    {
                        // A caret between two adjacent equations touches both:
                        // it belongs to the one before it (shapes come in
                        // document order). A real selection of two stays ambiguous.
                        if (selectionRange.Start == selectionRange.End) continue;
                        throw new InvalidDataException(
                            "The selected range contains more than one MathType formula.");
                    }
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
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }

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


    // reportUnreadableEquation receives the reason when the double-clicked
    // equation was found but could not be read; every miss is traced.
    public OfficeSelection? ReadVisualTeXOmmlAtScreenPoint(
        int screenX,
        int screenY,
        Action<string>? reportUnreadableEquation = null)
    {
        Document? document = null;
        Window? window = null;
        object? pointObject = null;
        Range? pointRange = null;
        Range? resolvedRange = null;
        WordFormulaHostDescriptor? host = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            window = _application.ActiveWindow;
            pointObject = window.RangeFromPoint(screenX, screenY);
            pointRange = pointObject as Range;
            if (pointRange is null)
            {
                WordDoubleClickHook.TraceMessage($"omml-hit-miss no-text-range x={screenX} y={screenY}");
                return null;
            }
            pointObject = null;

            host = WordFormulaHostLookup.ResolveLocal(
                document,
                pointRange,
                WordFormulaHostKind.Omml);
            if (host is null)
            {
                WordDoubleClickHook.TraceMessage(
                    $"omml-hit-miss no-equation point={pointRange.Start}:{pointRange.End}");
                return null;
            }

            resolvedRange = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            if (!ScreenPointHitsRange(
                    window,
                    resolvedRange,
                    screenX,
                    screenY))
            {
                WordDoubleClickHook.TraceMessage(
                    $"omml-hit-miss outside-equation x={screenX} y={screenY} equation={host.Range.Start}:{host.Range.End}");
                return null;
            }

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
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage(
                $"omml-hit-read-failed equationFound={host is not null} {error.GetType().Name}: {error.Message}");
            if (host is not null) reportUnreadableEquation?.Invoke(error.Message);
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
                var host = WordFormulaHostLookup.ResolveLocal(
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

    public bool IsActiveSourceDocument(string? expectedIdentity)
    {
        if (string.IsNullOrWhiteSpace(expectedIdentity)) return true;
        Document? document = null;
        try
        {
            document = _application.ActiveDocument;
            return document is not null && IsSameSourceDocument(document, expectedIdentity!);
        }
        catch (COMException) { return false; }
        finally { Release(document); }
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
        Unresolved,
    }

    public bool IsSelectedNativeOle() =>
        ReadSelectedOleKind() == SelectedOleKind.VisualTeX;

    public bool IsSelectedMathTypeOle() =>
        ReadSelectedOleKind() == SelectedOleKind.MathType;

    // Enabling an explicit command grants permission to inspect a selected OLE,
    // not proof that it is a formula and not ownership of its native gestures.
    // The command validates CFB/MTEF before any mutation. Missing registration
    // must not make that data-based inspection unreachable from the Ribbon.
    internal bool IsSelectedOleFormulaCommandCandidate() =>
        ReadSelectedOleKind() != SelectedOleKind.None;

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
        try
        {
            selection = _application.Selection;
            range = selection.Range;
            shapes = range.InlineShapes;
            if (shapes.Count != 1) return SelectedOleKind.None;
            shape = shapes[1];
            return ReadOleHostKindWithoutContent(shape);
        }
        catch
        {
            return SelectedOleKind.None;
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(range);
            Release(selection);
        }
    }

    private static SelectedOleKind ReadOleHostKindWithoutContent(InlineShape shape)
    {
        // Observation may read a stored host identity, but must not fall back to
        // WordOpenXML/CFB or activate an OLE object when registration is missing.
        // First prove that this is an OLE host. An unavailable ProgID thereafter
        // means unresolved identity, not the absence of an inspectable object.
        if (shape.Type is not WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
            and not WdInlineShapeType.wdInlineShapeLinkedOLEObject)
            return SelectedOleKind.None;
        OLEFormat? format = null;
        try
        {
            format = shape.OLEFormat;
            var progId = WordOleObjectAccessor.ReadProgId(format);
            if (string.Equals(progId, FormulaOleContract.ProgId, StringComparison.OrdinalIgnoreCase))
                return SelectedOleKind.VisualTeX;
            return string.Equals(progId, MathTypeOleInterop.CanonicalProgId, StringComparison.OrdinalIgnoreCase)
                || MathTypeOleInterop.TryResolveCapabilities(progId, out _)
                ? SelectedOleKind.MathType
                : SelectedOleKind.Unresolved;
        }
        catch { return SelectedOleKind.Unresolved; }
        finally { Release(format); }
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
        Range? probe = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        InlineShape? host = null;
        Range? hostRange = null;
        try
        {
            caret = selection.Range;
            var position = caret.Start;
            if (position != caret.End) return;

            // A passive caret repair needs only the exact trailing boundary.
            // Generic host discovery also reads the OLE payload at the leading
            // boundary and asks layout-backed wdWithInTable for a descriptor
            // field unused here. Keep the same tiny structural/ownership probe,
            // but establish the trailing boundary before any payload access.
            probe = caret.Duplicate;
            try { probe.MoveStart(WdUnits.wdCharacter, -2); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            try { probe.MoveEnd(WdUnits.wdCharacter, 2); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            shapes = probe.InlineShapes;
            var touchingHosts = 0;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shapeRange); shapeRange = null;
                Release(shape); shape = shapes[index];
                shapeRange = shape.Range.Duplicate;
                if (shapeRange.StoryType != caret.StoryType
                    || position < shapeRange.Start || position > shapeRange.End
                    || !WordFormulaMetadataReader.IsNativeOle(shape))
                    continue;
                touchingHosts++;
                if (touchingHosts > 1) return;
                if (position != shapeRange.End) continue;
                host = shape; shape = null;
                hostRange = shapeRange; shapeRange = null;
            }
            if (touchingHosts != 1 || host is null || hostRange is null) return;

            // Preserve authoritative block/inline semantics at the one boundary
            // that can need normalization. Native OMML and third-party OLE never
            // enter this payload path; their insertion/edit lifecycle is owned
            // by Word/their own server, not passive selection notifications.
            var metadata = WordFormulaMetadataReader.TryReadEmbeddedNativeOle(host);
            if (string.Equals(metadata?.DisplayMode, "block", StringComparison.OrdinalIgnoreCase)
                || caret.Start != hostRange.End)
                return;
            WordFormulaHostLayout.RestoreInlineCaret(_application, hostRange);
        }
        catch (Exception ignored)
        {
            // Selection-change normalization is best-effort and must never
            // interrupt ordinary Word input.
            IgnoredErrors.Ignore(ignored);
        }
        finally
        {
            Release(hostRange);
            Release(host);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(probe);
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
                try { selectionFont.NameAscii = sourceFont.NameAscii; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                try { selectionFont.NameFarEast = sourceFont.NameFarEast; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                try { selectionFont.NameOther = sourceFont.NameOther; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                selectionFont.Size = sourceFont.Size;
                selectionFont.Bold = sourceFont.Bold;
                selectionFont.Italic = sourceFont.Italic;
                try { selectionFont.Underline = sourceFont.Underline; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                try { selectionFont.Color = sourceFont.Color; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
            try { selectionFont.Spacing = 0; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            try { selectionFont.Scaling = 100; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
        }
        finally
        {
            Release(selectionFont);
            Release(sourceFont);
            Release(source);
            Release(caret);
        }
    }

    internal static bool TryResolveWordFontSize(float value, out float fontSizePt)
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

    // Passive Ribbon/selection queries must not serialize or activate third-party OLE.
    // A third-party MathType semantic size is unknown until an explicit command
    // reads its MTEF; Word's surrounding text font is not a substitute for it.
    public float? GetSelectedFormulaFontSize() =>
        ReadSelectedFormulaFontSize(allowMathTypeContent: false);

    internal float? ReadSelectedFormulaFontSizeForCommand() =>
        ReadSelectedFormulaFontSize(allowMathTypeContent: true);

    private float? ReadSelectedFormulaFontSize(bool allowMathTypeContent)
    {
        // The command and observation entry points deliberately have different
        // authority. A nominally read-only WordOpenXML property saves OLE data
        // and is unsafe when native MathType is still constructing the object.
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
                if ((shape.Type is WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                        or WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                    && ReadOleHostKindWithoutContent(shape) != SelectedOleKind.VisualTeX)
                {
                    // Even IsMathTypeOle can export WordOpenXML when ProgID is
                    // unavailable. That semantic fallback belongs to commands,
                    // never to passive observation of another owner's OLE.
                    if (!allowMathTypeContent || !MathTypeOleInterop.IsMathTypeOle(shape))
                        return null;
                    if (!MathTypeOleStorage.TryCaptureCompoundFileFromWordOpenXml(
                            shape,
                            out var mathTypeCompound))
                        return null;
                    var equationNative = MathTypeOleStorage.ReadEquationNative(mathTypeCompound);
                    return FormulaFontSize.Normalize(
                        MathTypeMtefFontSize.ReadEquationNativeFullFontSizeOrDefault(equationNative));
                }
                var metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata is not null)
                {
                    var stableSize = WordFormulaMetadataReader.IsNativeOle(shape)
                        ? WordInlineObjectGeometry.ReadStableVisualTeXSize(shape, metadata)
                        : (Width: shape.Width, Height: shape.Height);
                    return FormulaFontSize.InferOleFontSize(stableSize.Width, stableSize.Height, metadata);
                }
            }

            var omml =
                WordFormulaHostLookup.ResolveLocal(
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
        catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
            return WordEquationNumberFormatAuthority.Resolve(document);
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
            return WordEquationNumberFormat.GetEquationNumberFormatDisplayName(document);
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
        Action? restoreNonUndoState = null,
        Action? verifyRollback = null)
    {
        try
        {
            return WordFormulaMutationTransaction.Execute(
                _application,
                document,
                name,
                edit,
                verifyRollback: verifyRollback);
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
                    WordOleObjectAccessor.ReadProgId(format),
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
                catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
                        WordNumberedTypingParagraphs.CleanupNumberedDisplayInsertionSpacing(
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
        var saved = WordEquationNumberFormat.TryGetDefaultMathTypeNumberPosition();
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
        bool preserveCapturedInsertion = false,
        bool boundedBatchMutation = false,
        Range? exactInsertionRange = null,
        bool mathTypeNativeStylesPrepared = false,
        Action<string, string>? deferInlineParagraphRepair = null,
        MathTypeMtefCodec.RewriteResult? preparedStandaloneGenerated = null,
        byte[]? preparedStandaloneCompoundFile = null,
        string? preparedStandaloneSemanticSignature = null,
        Range? preparedBatchMathTypeSection = null,
        WordParagraphSnapshot? sourceParagraph = null)
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
        var metadata = session.ToMetadata();
        metadata.Validate();
        if (string.IsNullOrWhiteSpace(metadata.Latex))
            throw new InvalidDataException(
                "VisualTeX did not provide LaTeX source for MathType OLE insertion.");

        var inline = string.Equals(
            session.DisplayMode,
            "inline",
            StringComparison.OrdinalIgnoreCase);
        var hasPreparedStandalone =
            preparedStandaloneGenerated is not null
            && preparedStandaloneCompoundFile is { Length: > 0 }
            && !string.IsNullOrWhiteSpace(preparedStandaloneSemanticSignature);
        var standalone = hasPreparedStandalone
            ? (
                Generated: preparedStandaloneGenerated!,
                CompoundFile: preparedStandaloneCompoundFile!,
                SemanticSignature: preparedStandaloneSemanticSignature!)
            : PrepareStandaloneMathTypeOleData(
                mathMl,
                inline,
                session.FontSizePt,
                metadata.Latex);
        var generated = standalone.Generated;
        var compoundFile = standalone.CompoundFile;
        var expectedSignature = standalone.SemanticSignature;

        // The visible Word presentation must use MathType's own MTEF geometry.
        // Using the frontend/MathJax EMF here made otherwise-valid Equation Native
        // objects look subtly different from equations inserted by MathType itself
        // (notably relation/operator spacing) and could size the OLE host a few
        // pixels too narrowly, clipping the right-most italic glyph. The private
        // renderer is required; installed MathType is not a rendering dependency.
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
            throw new InvalidOperationException(MathTypeNativePreviewRenderer.FailureMessage);
        }

        // Keep genuine MathType storage free of OlePres. Word owns the external
        // WMF presentation in the DOCX package, avoiding the blank-object
        // regression that prompted the visibility change in 6a43aec.

        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Range? displaySpacingAnchor = null;
        // Reserved for the Range-only compatibility path. The current batch writer
        // resolves one concrete InlineShape first; keeping this null makes all
        // downstream structural work use that same proven local host.
        Range? batchHostRange = null;
        InlineShape? shape = null;
        Bookmark? createdObjectBookmark = null;
        Range? createdObjectBookmarkRange = null;
        UndoRecord? undoRecord = null;
        WordParagraphFormatting? inlineParagraphFormatting = null;
        MathTypeLocalEditScope? insertionLocalScope = null;
        WordUndoHistorySnapshot? insertionUndoHistory = null;
        MathTypeSectionStaging? preparedHeadingParagraph = null;
        var sourceParagraphCount = -1;
        var paragraphCountBeforeDisplayPreparation = -1;
        var insertionStart = -1;
        var createdSectionBreakCodeStart = -1;
        // A bounded document/selection batch owns the captured insertion point and
        // one final numbering reconciliation after all mutations. Its per-item
        // writer therefore uses only local host probes and skips global document
        // scans that would otherwise grow with the number of formulas. A temporary
        // object-identity bookmark provides the same guarantee for legacy callers.
        var useLocalConversionLookup =
            boundedBatchMutation
            || !string.IsNullOrWhiteSpace(createdObjectBookmarkName);
        var traceInsertPerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF"),
            "1",
            StringComparison.Ordinal);
        var insertPerfWatch = traceInsertPerformance ? Stopwatch.StartNew() : null;
        var insertMetrics = traceInsertPerformance && !useLocalConversionLookup
            ? new WordOperationMetrics() : null;
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
        var insertionFailed = false;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            selection = _application.Selection;

            stage = "resolve-captured-insertion";
            insertion = exactInsertionRange is not null
                ? exactInsertionRange.Duplicate
                : ResolveSessionInsertionRange(
                    document,
                    session,
                    selection,
                    preserveCapturedInsertion);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            var insertionPosition = insertion.Start;
            if (inline)
            {
                // The imported Flat OPC paragraph is only a transport envelope;
                // the existing Word paragraph owns its style, spacing and vertical
                // alignment. InsertXML and paragraph merging can otherwise promote
                // the empty envelope's defaults onto the surrounding prose.
                inlineParagraphFormatting = WordParagraphFormatting.Capture(insertion);
                ParagraphFormat? sourceParagraphFormat = null;
                try
                {
                    sourceParagraphFormat = insertion.ParagraphFormat;
                    // MathPage's negative descent is a baseline-coordinate offset.
                    // Explicit top/center/Far-East alignment already places the
                    // picture box; applying descent again moves it off that axis.
                    if (sourceParagraphFormat.BaseLineAlignment is
                        WdBaselineAlignment.wdBaselineAlignCenter
                        or WdBaselineAlignment.wdBaselineAlignTop
                        or WdBaselineAlignment.wdBaselineAlignFarEast50)
                        wordPosition = 0;
                }
                finally { Release(sourceParagraphFormat); }
            }
            // Capture before Flat OPC/style materialization: neither an OLE run
            // nor MTDisplayEquation's defaults are the surrounding body font.
            var bodyFormatting = inline || mathTypeNativeStylesPrepared
                ? null
                : WordCharacterFormatting.CaptureParagraphMarkAtPosition(
                    document,
                    insertionPosition);
            TraceInsertPerf("resolve-insertion");

            void BeginInsertionMutation()
            {
                // Template WordOpenXML and hidden section staging belong to
                // preflight. Starting before them splits native Undo in Word.
                if (!useLocalConversionLookup)
                {
                    insertionLocalScope = MathTypeLocalEditScope.CaptureInsertion(document, insertion);
                    TraceInsertPerf("capture-local-insertion-scope");
                    insertionUndoHistory = new WordUndoHistorySnapshot(document);
                    TraceInsertPerf("capture-undo-history");
                }
                undoRecord = BeginUndoRecord(
                    inline ? "VisualTeX Insert MathType OLE Inline Formula"
                        : session.Numbered ? "VisualTeX Insert MathType OLE Numbered Display Formula"
                        : "VisualTeX Insert MathType OLE Display Formula");
                TraceInsertPerf("begin-undo-record");
                if (insertionLocalScope is not null && undoRecord is null)
                    throw new InvalidOperationException("Word could not establish an independent MathType insertion transaction.");
                if (!inline && session.Numbered)
                {
                    stage = "repair-incomplete-number-row";
                    ClearIncompleteMathTypeNumberRowAtInsertion(document, insertion);
                    TraceInsertPerf("clear-incomplete-number-row");
                }
            }
            if (!inline && session.Numbered) EnsureEquationFieldResultsVisible(document);

            MathTypeWordOpenXml.NumberTemplate? numberTemplate = null;
            if (!inline && session.Numbered)
            {
                stage = "resolve-number-template";
                var documentNumberFormat = EquationNumberFormat.Resolve(
                    WordEquationNumberFormatAuthority.Resolve(document));
                // A native donor supplies neither policy nor ownership for this
                // new number. Pasted/legacy donors can have a different format
                // from the explicit document policy even when their field trees
                // are perfectly valid. All entry paths materialize the SAME
                // policy-selected native MTPlaceRef template; sequence state,
                // left/right placement and reference identity remain separate.
                numberTemplate = MathTypeWordOpenXml.CreateVisualTeXNumberTemplate(
                    documentNumberFormat.Id);

                MathTypeHeadingPlan? preparedPlan = null;
                var scopeAlreadyPrepared = preResolvedHeadingScope is not null
                    && preResolvedHeadingScope.ScopeStart != int.MinValue
                    && preparedHeadingScopeStarts?.Contains(preResolvedHeadingScope.ScopeStart) == true;
                var needsHeadingState = documentNumberFormat.UsesHeading && numberTemplate is not null
                    && MathTypeNumberTemplateUsesHeading(numberTemplate) && !scopeAlreadyPrepared
                    && preResolvedHeadingScope?.ScopeStart != int.MinValue;
                if (!useLocalConversionLookup && needsHeadingState)
                {
                    preparedPlan = PrepareMathTypeHeadingScope(document, insertion.Start,
                        documentNumberFormat, preResolvedHeadingScope);
                    if (preparedPlan is not null)
                    {
                        using (WordOperationMetrics.Measure("MathType.Numbering.PrepareSection"))
                            preparedHeadingParagraph = new MathTypeSectionStaging(document,
                                CreateMathTypeSectionFlatOpc(preparedPlan.Chapter, preparedPlan.Section));
                    }
                }
                TraceInsertPerf("number-template-preflight");
                BeginInsertionMutation();

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
                    if (needsHeadingState)
                    {
                        var insertedSectionLength = useLocalConversionLookup
                            ? EnsureMathTypeHeadingScopeState(document, logicalInsertionStart,
                                documentNumberFormat, out createdSectionBreakCodeStart, preResolvedHeadingScope,
                                preparedBatchMathTypeSection)
                            : ApplyMathTypeHeadingScope(document, logicalInsertionStart,
                                preparedPlan, out createdSectionBreakCodeStart, preparedHeadingParagraph?.Paragraph);
                        if (preResolvedHeadingScope is not null)
                            preparedHeadingScopeStarts?.Add(
                                preResolvedHeadingScope.ScopeStart);
                        if (insertedSectionLength > 0)
                        {
                            if (preparedPlan is not null)
                                insertionLocalScope?.RebaseAfterPrecedingInsertion(
                                    preparedPlan.ScopeEnd, insertedSectionLength);
                            var shiftedStart = Math.Min(
                                document.Content.End,
                                logicalInsertionStart + insertedSectionLength);
                            insertion.SetRange(shiftedStart, shiftedStart);
                        }
                    }
                }
                TraceInsertPerf("number-template-heading-state");
            }
            else BeginInsertionMutation();

            if (!inline)
            {
                stage = "prepare-display-row";
                using (WordOperationMetrics.Measure(
                           "MathType.Insert.ParagraphCountBeforeDisplay"))
                    paragraphCountBeforeDisplayPreparation = insertionLocalScope?.ReadParagraphCount()
                        ?? ReadDocumentParagraphCount(document);
                if (!preserveExistingDisplayParagraphBoundary)
                    displaySpacingAnchor = insertion.Duplicate;
                Range displayInsertion;
                using (WordOperationMetrics.Measure(
                           "MathType.Insert.ResolveDisplayInsertion"))
                {
                    displayInsertion = boundedBatchMutation
                        && preserveExistingDisplayParagraphBoundary
                        && exactInsertionRange is not null
                        && CanUseExactBatchMathTypeDisplayInsertion(
                            document,
                            insertion)
                        ? insertion.Duplicate
                        : ResolveStandaloneMathTypeDisplayInsertionRange(
                            document,
                            insertion,
                            replaceAtExactInsertion: preserveExistingDisplayParagraphBoundary);
                }
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
            var batchIdentityBookmarkName = boundedBatchMutation
                ? WordFormulaIdentityStore.MathTypeBookmarkName(metadata.FormulaId)
                : null;
            string wordOpenXml;
            using (WordOperationMetrics.Measure(
                       "MathType.Insert.BuildFlatOpc"))
            {
                wordOpenXml = MathTypeWordOpenXml.CreateWithPlaceableWmf(
                    compoundFile,
                    previewWmf,
                    widthPt,
                    heightPt,
                    display: !inline,
                    numberTemplate,
                    mathTypeNumberPosition: relocateRightNumberToLeftAfterInsert
                        ? "right"
                        : session.MathTypeNumberPosition,
                    locatorBookmarkName: batchIdentityBookmarkName);
            }
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
            using (WordOperationMetrics.Measure(
                       "MathType.Insert.ParagraphCountBeforeInsert"))
                sourceParagraphCount = insertionLocalScope?.ReadParagraphCount()
                    ?? ReadDocumentParagraphCount(document);
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
                using (WordOperationMetrics.Measure(
                           "MathType.Insert.InsertXml"))
                    WpsXmlInsertion.Insert(insertion, wordOpenXml);
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
            var resolvedByBatchIdentityBookmark = false;
            using (WordOperationMetrics.Measure(
                       "MathType.Insert.ResolveNewShape"))
            {
                if (!string.IsNullOrWhiteSpace(batchIdentityBookmarkName))
                {
                    using (WordOperationMetrics.Measure(
                               "MathType.Resolve.IdentityBookmark"))
                    {
                        if (boundedBatchMutation)
                        {
                            batchHostRange = FindMathTypeHostRangeByIdentityBookmark(
                                document,
                                batchIdentityBookmarkName!);
                            resolvedByBatchIdentityBookmark = batchHostRange is not null;
                        }
                        else
                        {
                            shape = FindMathTypeOleByIdentityBookmark(
                                document,
                                batchIdentityBookmarkName!);
                            resolvedByBatchIdentityBookmark = shape is not null;
                        }
                    }
                }
                if (shape is null && batchHostRange is null)
                {
                    using (WordOperationMetrics.Measure(
                               "MathType.Resolve.Position"))
                    {
                        shape = FindSingleInlineShapeAtPosition(
                            document,
                            insertionStart);
                        if (shape is not null
                            && !MathTypeOleInterop.IsMathTypeOle(shape))
                        {
                            Release(shape);
                            shape = null;
                        }
                    }
                }
                if (shape is null && batchHostRange is null)
                {
                    using (WordOperationMetrics.Measure(
                               "MathType.Resolve.Window"))
                        shape = FindSingleInlineShapeInInsertionWindow(
                            document,
                            insertionStart);
                }
                if (shape is null && batchHostRange is null)
                {
                    using (WordOperationMetrics.Measure(
                               "MathType.Resolve.RangeReference"))
                        shape = FindMathTypeOleByRange(
                            document,
                            insertedRangeReference,
                            allowGlobalFallback: false);
                }
                if (shape is null && batchHostRange is null)
                {
                    using (WordOperationMetrics.Measure(
                               "MathType.Resolve.Paragraph"))
                        shape = FindMathTypeOleInParagraphAtPosition(
                            document,
                            insertionStart);
                }
                if (shape is null && batchHostRange is null)
                {
                    using (WordOperationMetrics.Measure(
                               "MathType.Resolve.LocalWindow"))
                        shape = FindMathTypeOleInLocalWindow(
                            document,
                            insertionStart);
                }
                if (shape is null && batchHostRange is null && !boundedBatchMutation)
                {
                    using (WordOperationMetrics.Measure(
                               "MathType.Resolve.NearPosition"))
                        shape = FindMathTypeOleNearPosition(
                            document,
                            insertionStart);
                }
            }
            if (shape is null && batchHostRange is null)
                throw new InvalidOperationException(
                    boundedBatchMutation
                        ? $"Word inserted the MathType OLE data but the bounded local resolver could not resolve formula {metadata.FormulaId} at story position {insertionStart}."
                        : "Word inserted the MathType OLE data but VisualTeX could not resolve the new equation.");
            if (shape is not null && !MathTypeOleInterop.IsMathTypeOle(shape))
                throw new InvalidOperationException(
                    "Word did not materialize the standalone equation as Equation.DSMT4.");
            if (batchHostRange is null
                && boundedBatchMutation
                && !string.IsNullOrWhiteSpace(batchIdentityBookmarkName))
            {
                batchHostRange = shape!.Range.Duplicate;
            }
            TraceInsertPerf("resolve-new-shape");
            // Materialized XML validation is deliberately deferred until all
            // style/paragraph/bookmark mutations are complete. WordOpenXML can
            // terminate a custom UndoRecord; reading it here split one insertion
            // into multiple native undo operations.

            // CreateWithPlaceableWmf already writes the final VML width/height
            // and w:dxaOrig/w:dyaOrig into the same Flat OPC transaction as the
            // Equation.DSMT4 storage and preview.  Do not immediately rewrite
            // InlineShape geometry through COM after InsertXML.  In Word
            // compatibility mode 14, a numbered Flat OPC paragraph that also
            // materializes nested MTPlaceRef/SEQ fields can keep the just-inserted
            // OLE result range temporarily non-editable until the field transaction
            // settles; assigning LockAspectRatio/Width/Height in that window raises
            // 0x800A1710 ("Range cannot be edited").  The geometry write was also
            // redundant because it repeated the dimensions already committed by
            // Flat OPC.  Keep Word's materialized presentation unchanged.
            stage = "use-flat-opc-presentation";
            TraceInsertPerf("use-flat-opc-geometry");

            var repairedInsertXmlParagraph = false;
            if (sourceParagraphCount >= 0)
            {
                var repairParagraphSplit = inline;
                if (!repairParagraphSplit
                    && preserveExistingDisplayParagraphBoundary)
                {
                    Range? insertedShapeRange = null;
                    try
                    {
                        insertedShapeRange = batchHostRange is not null
                            ? batchHostRange.Duplicate
                            : shape!.Range;
                        repairParagraphSplit =
                            WordNumberedDisplayLayout.RangeIsWhollyWithinTable(
                                insertedShapeRange);
                    }
                    finally { Release(insertedShapeRange); }
                }
                if (repairParagraphSplit)
                {
                    if (inline && deferInlineParagraphRepair is not null)
                    {
                        using (WordOperationMetrics.Measure(
                                   "MathType.Insert.DeferParagraphSplit"))
                        {
                            var currentParagraphCount =
                                insertionLocalScope?.ReadParagraphCount()
                                ?? ReadDocumentParagraphCount(document);
                            if (currentParagraphCount == sourceParagraphCount + 1)
                            {
                                deferInlineParagraphRepair(
                                    metadata.FormulaId,
                                    preservedRightBoundaryWhitespace
                                        ?? string.Empty);
                            }
                            else if (currentParagraphCount != sourceParagraphCount)
                            {
                                throw new InvalidOperationException(
                                    $"Word changed the paragraph count unexpectedly while materializing a deferred MathType OLE: before={sourceParagraphCount}, after={currentParagraphCount}.");
                            }
                        }
                    }
                    else
                    {
                        using (WordOperationMetrics.Measure(
                                   "MathType.Insert.RepairParagraphSplit"))
                        {
                            repairedInsertXmlParagraph = batchHostRange is not null
                                ? RepairMathTypeInsertXmlParagraphSplit(
                                    document,
                                    batchHostRange,
                                    sourceParagraphCount,
                                    preservedRightBoundaryWhitespace,
                                    insertionLocalScope)
                                : RepairMathTypeInsertXmlParagraphSplit(
                                    document,
                                    shape!,
                                    sourceParagraphCount,
                                    preservedRightBoundaryWhitespace,
                                    insertionLocalScope);
                        }
                    }
                }
            }

            TraceInsertPerf("repair-insertxml-paragraph");

            // Only an actual paragraph-boundary repair can invalidate the live
            // InlineShape COM handle. The old path re-resolved every equation even
            // when no structural mutation occurred, duplicating the same bounded
            // OLE lookup immediately after the successful insertion lookup above.
            if (repairedInsertXmlParagraph)
            {
                stage = "refresh-native-shape";
                if (batchHostRange is not null
                    && !string.IsNullOrWhiteSpace(batchIdentityBookmarkName))
                {
                    var refreshedRange = FindMathTypeHostRangeByIdentityBookmark(
                        document,
                        batchIdentityBookmarkName!,
                        out var bookmarkFailure);
                    if (refreshedRange is null)
                    {
                        // Merging the split paragraph can cost the locator
                        // bookmark in some documents (issue #35) although the
                        // OLE stays where it was inserted: the repair only
                        // removes the paragraph mark after it. Take it from the
                        // bounded insertion window, which must hold exactly
                        // that one MathType OLE, and restore the bookmark.
                        refreshedRange = RecoverMathTypeHostRangeAtInsertion(
                            document,
                            batchIdentityBookmarkName!,
                            insertionStart);
                        WordDoubleClickHook.TraceMessage(
                            $"mathtype-insert-bookmark-lost-after-repair formula={metadata.FormulaId} reason={bookmarkFailure} recovered={refreshedRange is not null}");
                    }
                    if (refreshedRange is null)
                        throw new InvalidOperationException(
                            $"Word retained the MathType OLE but lost its exact batch identity range after paragraph repair ({bookmarkFailure}).");
                    Release(batchHostRange);
                    batchHostRange = refreshedRange;
                }
                else
                {
                    var liveHandleValid = false;
                    Range? liveShapeRange = null;
                    try
                    {
                        liveShapeRange = shape!.Range;
                        liveHandleValid =
                            MathTypeOleInterop.IsMathTypeOle(shape)
                            && liveShapeRange.Start >= insertionStart - 1
                            && liveShapeRange.Start <= insertionStart + 8
                            && liveShapeRange.End > liveShapeRange.Start;
                    }
                    catch
                    {
                        liveHandleValid = false;
                    }
                    finally { Release(liveShapeRange); }

                    if (!liveHandleValid)
                    {
                        var refreshedRangeEnd = Math.Min(
                            document.Content.End,
                            insertionStart + 8);
                        var refreshedShape = FindSingleInlineShapeInInsertionWindow(
                                document,
                                insertionStart)
                            ?? FindMathTypeOleByRange(
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
                    }
                }
            }
            TraceInsertPerf("refresh-native-shape");

            if (inline)
            {
                Range? paragraphOwner = null;
                try
                {
                    paragraphOwner = batchHostRange is not null
                        ? batchHostRange.Duplicate
                        : shape!.Range;
                    inlineParagraphFormatting!.Apply(paragraphOwner);
                }
                finally { Release(paragraphOwner); }
                // This also restores the generated left fragment before a deferred
                // merge. Both fragments then carry the same source paragraph format.
                stage = "apply-native-baseline";
                using (WordOperationMetrics.Measure(
                           "MathType.Insert.SetInlineBaseline"))
                {
                    if (batchHostRange is not null)
                        SetInlineOleWordPosition(
                            batchHostRange,
                            wordPosition,
                            hostInstructionAlreadyBaseline: true);
                    else
                        SetInlineOleWordPosition(
                            shape!,
                            wordPosition,
                            hostInstructionAlreadyBaseline: boundedBatchMutation);
                }
                TraceInsertPerf("set-native-baseline");
                if (!reuseExistingInlineTypingBoundary)
                {
                    if (useLocalConversionLookup)
                    {
                        if (boundedBatchMutation)
                        {
                            if (deferInlineParagraphRepair is null)
                            {
                                Range? batchShapeRange = null;
                                try
                                {
                                    batchShapeRange = batchHostRange is not null
                                        ? batchHostRange.Duplicate
                                        : shape!.Range;
                                    ResetParagraphTypingPosition(batchShapeRange);
                                }
                                finally { Release(batchShapeRange); }
                            }
                            // Deferred redraw batches repair the InsertXML paragraph
                            // boundary and typing position once in finalization.
                        }
                        else
                        {
                            RestoreTypingBaselineAfterMathTypeConversion(shape);
                        }
                    }
                    else
                    {
                        RestoreTypingBaselineAfter(shape);
                    }
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
                using (WordOperationMetrics.Measure(
                           "MathType.Insert.ConfigureDisplay"))
                {
                    if (batchHostRange is not null)
                    {
                        ConfigureNewMathTypeDisplayEquation(
                            document,
                            batchHostRange,
                            session.Numbered,
                            session.MathTypeNumberPosition,
                            updateNestedNumberFields:
                                !useLocalConversionLookup || updateCreatedMathTypeNumberFields,
                            bodyFormatting: bodyFormatting,
                            nativeStylesAlreadyPrepared: mathTypeNativeStylesPrepared,
                            preservedDisplayParagraphLayout: preservedDisplayParagraphLayout,
                            knownDisplayColumnWidth: knownDisplayColumnWidth,
                            sourceParagraph: sourceParagraph);
                    }
                    else
                    {
                        ConfigureNewMathTypeDisplayEquation(
                            document,
                            shape!,
                            session.Numbered,
                            session.MathTypeNumberPosition,
                            updateNestedNumberFields:
                                !useLocalConversionLookup || updateCreatedMathTypeNumberFields,
                            bodyFormatting: bodyFormatting,
                            nativeStylesAlreadyPrepared: mathTypeNativeStylesPrepared,
                            preservedDisplayParagraphLayout: preservedDisplayParagraphLayout,
                            knownDisplayColumnWidth: knownDisplayColumnWidth,
                            sourceParagraph: sourceParagraph);
                    }
                }
                TraceInsertPerf("configure-display-numbering");
                // MTDisplayEquation setup resets direct character formatting on
                // the OLE run. Apply the exported math baseline after numbering
                // and paragraph style are final so the adjacent number stays
                // vertically aligned with tall display formulas.
                stage = "apply-native-display-baseline";
                using (WordOperationMetrics.Measure(
                           "MathType.Insert.SetDisplayBaseline"))
                {
                    if (batchHostRange is not null)
                        SetInlineOleWordPosition(
                            batchHostRange,
                            wordPosition,
                            hostInstructionAlreadyBaseline: true);
                    else
                        SetInlineOleWordPosition(
                            shape!,
                            wordPosition,
                            hostInstructionAlreadyBaseline: boundedBatchMutation);
                }
                TraceInsertPerf("display-baseline");
                var shapeRange = batchHostRange is not null
                    ? batchHostRange.Duplicate
                    : shape!.Range;
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
                        var typing = WordNumberedTableDisplay.EnsureBodyTypingParagraphAfterDisplay(
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
                var verificationRange = batchHostRange is not null
                    ? batchHostRange.Duplicate
                    : shape!.Range;
                try { retainMathTypeForCompletedValidation(verificationRange, expectedSignature); }
                finally { Release(verificationRange); }
            }

            // Keep VisualTeX ownership outside Equation.DSMT4/MTEF. The bookmark
            // is a locator only; MathType storage remains completely native.
            stage = "bind-mathtype-host-identity";
            using (WordOperationMetrics.Measure(
                       "MathType.Insert.BindIdentity"))
            {
                if (!resolvedByBatchIdentityBookmark)
                {
                    WordFormulaIdentityStore.BindMathType(
                        shape,
                        metadata.FormulaId);
                }
            }

            if (insertionUndoHistory is not null)
            {
                EndUndoRecord(undoRecord);
                Release(undoRecord);
                undoRecord = null;
            }
            stage = "validate-flat-opc-storage";
            if (!useLocalConversionLookup
                && retainMathTypeForCompletedValidation is null
                && MathTypeOleStorage.TryCaptureCompoundFileFromWordOpenXml(
                    shape!, out var materializedCompoundFile))
            {
                var materializedMathMl = MathTypeOleStorage.ReadMathMl(materializedCompoundFile);
                if (!MathTypeMathMlRoundTripMatches(expectedSignature, materializedMathMl))
                    throw new InvalidDataException(
                        "Word materialized a different MathType equation than VisualTeX generated.");
            }
            TraceInsertPerf("validate-final-storage");
            stage = "complete";
            TraceInsertPerf("complete");
            return Result(
                session,
                document,
                documentIdentityAlreadyValidated: true);
        }
        catch (Exception error)
        {
            insertionFailed = true;
            if (!useLocalConversionLookup && insertionUndoHistory is null)
                throw new InvalidOperationException(
                    $"MathType OLE preflight failed at stage '{stage}' before the target document was changed. {error.Message}", error);
            if (insertionUndoHistory is not null && insertionLocalScope is not null && document is not null)
            {
                try
                {
                    EndUndoRecord(undoRecord);
                    Release(undoRecord);
                    undoRecord = null;
                    insertionUndoHistory.UndoChanges(document);
                    insertionLocalScope.VerifyRollback();
                }
                catch (Exception recovery)
                {
                    throw new AggregateException(
                        $"MathType insertion failed at '{stage}' and its local Word undo recovery could not be verified.",
                        error, recovery);
                }
                throw new InvalidOperationException($"MathType OLE insertion failed at stage '{stage}'. {error.Message}", error);
            }
            // Preserve the primary insertion failure. Once Word has rejected an
            // operation (or its RPC server has terminated), cleanup COM calls may
            // fail as well; those are best-effort and must not replace the stage
            // and HRESULT that actually caused the batch to stop.
            try
            {
                if (document is not null)
                    WordFormulaIdentityStore.RemoveMathType(
                        document,
                        metadata.FormulaId);
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            try
            {
                if (document is not null && !string.IsNullOrWhiteSpace(createdObjectBookmarkName))
                    TryDeleteBookmark(document, createdObjectBookmarkName);
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            try
            {
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
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            try
            {
                if (createdSectionBreakCodeStart >= 0 && document is not null)
                    RemoveMathTypeSectionBreakFieldAtCodeStart(
                        document,
                        createdSectionBreakCodeStart);
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            var hresult = error is COMException
                ? $" HRESULT=0x{error.HResult:X8}."
                : string.Empty;
            throw new InvalidOperationException(
                $"MathType OLE insertion failed at stage '{stage}'.{hresult} {error.Message}",
                error);
        }
        finally
        {
            insertionLocalScope?.Dispose();
            inlineParagraphFormatting?.Dispose();
            nativePreview?.Dispose();
            TraceInsertPerf("finally-preview");
            try
            {
                EndUndoRecord(undoRecord);
            }
            catch (Exception ignored) when (insertionFailed)
            {
                // The primary insertion exception already carries the useful
                // stage/HRESULT. A dead Word RPC server cannot close the borrowed
                // undo record and must not mask that original failure.
                IgnoredErrors.Ignore(ignored);
            }
            Release(undoRecord);
            preparedHeadingParagraph?.Dispose();
            TraceInsertPerf("finally-undo");
            Release(createdObjectBookmarkRange);
            Release(createdObjectBookmark);
            TraceInsertPerf("finally-bookkeeping");
            Release(batchHostRange);
            batchHostRange = null;
            Release(shape);
            TraceInsertPerf("finally-shape");
            Release(displaySpacingAnchor);
            Release(insertion);
            TraceInsertPerf("finally-ranges");
            Release(selection);
            TraceInsertPerf("finally-selection");
            Release(document);
            TraceInsertPerf("finally-document");
            if (insertMetrics is not null)
            {
                foreach (var metric in insertMetrics.Entries)
                    WordDoubleClickHook.TraceMessage(
                        $"mathtype-insert-metric stage={metric.Key} calls={metric.Value.Calls} ms={metric.Value.Milliseconds:0.###}");
                insertMetrics.Dispose();
            }
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
        WordOmmlBatchSources.BatchSource? ommlBatchSource = null,
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

    internal WordLatexRedrawResult ApplyLatexRedrawPlan(
        WordLatexRedrawPlan plan,
        StreamingLatexRedrawFormulas streamed)
    {
        return ApplyLatexRedrawPlanCore(
            plan,
            new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal),
            streamed ?? throw new ArgumentNullException(nameof(streamed)));
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
        var coordinateShifts = new WordCoordinateShifts();
        var converted = false;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            selection = _application.Selection;
            scope = wholeDocument
                ? document.Content.Duplicate
                : selection.Range.Duplicate;
            var skipped = new List<(int Ordinal, string Reason)>();
            targets = CaptureFormulaToLatexTargets(
                document,
                scope,
                wholeDocument,
                objectMode,
                refreshOmmlMetadata: true,
                failures: skipped);
            // Preflight every source before opening a destructive Word undo record.
            // Corrupt/empty metadata must never turn a visible formula into an empty
            // replacement merely because the object itself can still be located.
            // Such a formula is left as it is; the others are still converted.
            foreach (var target in targets.ToArray())
            {
                try
                {
                    ApplyFormulaToLatexEmptySourceInjection(target);
                    target.LatexSource = BuildFormulaLatexSource(target.Metadata);
                }
                catch (Exception error) when (error is InvalidDataException or InvalidOperationException
                    or ArgumentException or NotSupportedException or COMException)
                {
                    WordDoubleClickHook.TraceMessage($"formula-to-latex-source-skipped start={target.Start} error={error}");
                    skipped.Add((target.Ordinal, OfficeErrorText.Describe(error)));
                    targets.Remove(target);
                    ReleaseFormulaToLatexTargets(new List<FormulaToLatexTarget> { target });
                }
            }
            if (targets.Count == 0 && skipped.Count > 0)
                throw new InvalidDataException(OfficePluginLanguage.Text(
                    "没有可以转为 LaTeX 的公式：",
                    "No equation could be converted to LaTeX: ")
                    + string.Join("; ", skipped.Select(item => OfficePluginLanguage.Text(
                        $"第 {item.Ordinal} 个公式：{item.Reason}", $"equation {item.Ordinal}: {item.Reason}"))));
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
                        ? OfficePluginLanguage.Text(
                            $"当前 Word 文档中没有找到可转换的 {modeLabel} 公式。",
                            $"No convertible {modeLabel} equations were found in the current Word document.")
                        : OfficePluginLanguage.Text(
                            $"所选内容中没有找到可转换的 {modeLabel} 公式。",
                            $"No convertible {modeLabel} equations were found in the current selection."));
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
            undoRecord = BeginUndoRecord(OfficePluginLanguage.Text(
                "VisualTeX 公式转为 LaTeX 代码",
                "VisualTeX Equation to LaTeX"));
            if (undoRecord is null)
                throw new InvalidOperationException(OfficePluginLanguage.Text(
                    "Word 无法建立公式转 LaTeX 的撤销事务。为避免公式丢失，本次转换未开始。",
                    "Word could not create the undo transaction for equation-to-LaTeX conversion. The conversion was not started to avoid losing the equation."));

            var undoRecordEnded = false;
            var documentMutationStarted = false;
            try
            {
                var convertedIds = new List<string>(targets.Count);
                foreach (var target in targets.OrderByDescending(item => item.Start))
                {
                    var storyEndBefore = ReadMainStoryEnd(document);
                    ConvertFormulaTargetToLatex(
                        document,
                        target,
                        ref documentMutationStarted);
                    // Back to front: earlier captured positions are still valid.
                    coordinateShifts.Add(target.End, ReadMainStoryEnd(document) - storyEndBefore);
                    convertedIds.Add(target.Metadata.FormulaId);
                }
                WordNumberingReconcile.TryReconcile(document);
                converted = true;
                return new WordFormulaToLatexResult
                {
                    FormulaCount = convertedIds.Count,
                    FormulaIds = convertedIds,
                    Skipped = skipped,
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
                            OfficePluginLanguage.Text(
                                "公式转 LaTeX 失败，而且 Word 无法完整恢复转换前的文档结构。请保留当前文档以便排查。",
                                "Equation-to-LaTeX conversion failed, and Word could not fully restore the pre-conversion document structure. Keep the current document for diagnosis."),
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
            // A failed conversion is rolled back to the original text, so only a
            // completed one moves the captured selection.
            if (converted && viewState is not null)
            {
                viewState.SelectionStart = coordinateShifts.Map(viewState.SelectionStart);
                viewState.SelectionEnd = coordinateShifts.Map(viewState.SelectionEnd);
            }
            RestoreViewState(document, viewState, preferredSelection: null);
            Release(undoRecord);
            ReleaseFormulaToLatexTargets(targets);
            Release(scope);
            Release(selection);
            Release(document);
        }
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
            throw new InvalidDataException(OfficePluginLanguage.Text(
                $"公式 {metadata.FormulaId} 没有可安全恢复的 LaTeX 源码。",
                $"Equation {metadata.FormulaId} has no LaTeX source that can be restored safely."));
        if (target.OleShape is null)
            throw new InvalidDataException(OfficePluginLanguage.Text(
                "MathType 公式转 LaTeX 时无法定位源 Equation.DSMT4 对象。",
                "The source Equation.DSMT4 object could not be located while converting the MathType equation to LaTeX."));

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

    public WordBulkInsertResult InsertBulkDocument(
        WordBulkImportDocument source,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        string? expectedDocumentId,
        string? sourceObjectId)
    {
        using var numberFormatScope =
            WordEquationNumberFormatAuthority.BeginOperationScope();
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

        return InsertMathTypeBulkSkeleton(
            source, prepared, expectedDocumentId, sourceObjectId);
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
        var mathTypeDisplayWidthsBySection =
            new Dictionary<int, float>();
        var mathTypeNativeStylesPrepared = false;
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

                    var displayParagraphStart =
                        selection.Start;
                    Range? displayInsertion = null;
                    try
                    {
                        displayInsertion =
                            selection.Range.Duplicate;
                        displayInsertion.Collapse(
                            WdCollapseDirection.wdCollapseStart);
                        var displayColumnWidth =
                            ResolveBatchMathTypeDisplayColumnWidth(
                                document,
                                displayInsertion,
                                mathTypeDisplayWidthsBySection);
                        InsertPreparedFormula(
                            document,
                            selection,
                            formula,
                            display: true,
                            preserveExistingDisplayParagraphBoundary: true,
                            ommlBatchSource: null,
                            deferredOmmlMetadata: null,
                            bulkImport: true,
                            mathTypeDisplayColumnWidth:
                                displayColumnWidth,
                            mathTypeNativeStylesPrepared:
                                mathTypeNativeStylesPrepared);
                        mathTypeNativeStylesPrepared = true;
                    }
                    finally
                    {
                        Release(displayInsertion);
                    }
                    insertedFormulaIds.Add(
                        formula.Session.FormulaId);
                    MoveSelectionAfterBulkParagraph(
                        document,
                        selection,
                        displayParagraphStart);
                    ResetNextParagraphFormatting(
                        selection,
                        block.Kind,
                        nextKind);
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

                var deferredInlineRepairs =
                    new List<(string FormulaId, string RightBoundaryWhitespace)>();
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
                        bulkImport: true,
                        deferMathTypeInlineParagraphRepair:
                            (id, rightWhitespace) =>
                                deferredInlineRepairs.Add(
                                    (id, rightWhitespace)));
                    insertedFormulaIds.Add(
                        pending.Formula.Session.FormulaId);
                }

                if (deferredInlineRepairs.Count > 0)
                {
                    FinalizeDeferredMathTypeInlineParagraphRepairs(
                        document,
                        deferredInlineRepairs,
                        out _,
                        out _,
                        out _);
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

    private void InsertPreparedFormula(
        Document document,
        Selection selection,
        PreparedWordBulkFormula prepared,
        bool display,
        bool preserveExistingDisplayParagraphBoundary = false,
        Range? preservedDisplayParagraphRange = null,
        string? preservedFollowingParagraphText = null,
        WordOmmlBatchSources.BatchSource? ommlBatchSource = null,
        ICollection<FormulaMetadata>? deferredOmmlMetadata = null,
        Action<Range>? retainInsertedOmml = null,
        bool deferOmmlBatchFinalization = false,
        Action<Range>? retainInsertedNativeOleIdentity = null,
        bool deferNativeOleBatchFinalization = false,
        bool bulkImport = false,
        float? mathTypeDisplayColumnWidth = null,
        bool mathTypeNativeStylesPrepared = false,
        Action<string, string>? deferMathTypeInlineParagraphRepair = null,
        Action<Range, string>? retainMathTypeForCompletedValidation = null,
        Range? exactMathTypeBulkInsertion = null,
        Range? preparedMathTypeSection = null)
    {
        _ = ommlBatchSource;
        _ = deferredOmmlMetadata;
        _ = retainInsertedOmml;
        _ = deferOmmlBatchFinalization;
        _ = retainInsertedNativeOleIdentity;
        _ = deferNativeOleBatchFinalization;

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

        Range? exactBulkInsertion = null;
        try
        {
            if (bulkImport)
            {
                if (exactMathTypeBulkInsertion is not null)
                    exactBulkInsertion = exactMathTypeBulkInsertion.Duplicate;
                else
                {
                    var selectedRange = selection.Range;
                    try { exactBulkInsertion = selectedRange.Duplicate; }
                    finally { Release(selectedRange); }
                }
                exactBulkInsertion.Collapse(
                    WdCollapseDirection.wdCollapseStart);
            }

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
                    true,
                boundedBatchMutation:
                    bulkImport,
                exactInsertionRange:
                    exactBulkInsertion,
                mathTypeNativeStylesPrepared:
                    mathTypeNativeStylesPrepared,
                deferInlineParagraphRepair:
                    deferMathTypeInlineParagraphRepair,
                preparedStandaloneGenerated:
                    prepared.MathTypeStandaloneGenerated,
                preparedStandaloneCompoundFile:
                    prepared.MathTypeStandaloneCompoundFile,
                preparedStandaloneSemanticSignature:
                    prepared.MathTypeStandaloneSemanticSignature,
                preparedBatchMathTypeSection:
                    preparedMathTypeSection);
        }
        finally
        {
            Release(exactBulkInsertion);
        }

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
        MathTypeLocalEditScope? localEdit = null;
        WordUndoHistorySnapshot? undoHistory = null;
        var mutationStarted = false;
        var previousTrackRevisions = false;
        var trackingSuspended = false;
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        MathTypeDisplayParagraphLayout? displayParagraphLayout = null;
        MathTypeDisplayParagraphLayout? inlineParagraphLayout = null;
        MathTypeNativePreviewRenderer.Result? sourceNativePreview = null;
        MathTypeNativePreviewRenderer.Result? nativePreview = null;
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
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }

            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=resolve-source formulaId={session.FormulaId}");
            oldShape = FindMathTypeOleByRange(document, session.SourceObjectId, allowGlobalFallback: false)
                ?? throw new InvalidOperationException(
                    "The MathType OLE equation no longer exists at the captured Word location.");
            if (!MathTypeOleInterop.IsMathTypeOle(oldShape))
                throw new InvalidOperationException(
                    "The selected OLE object is no longer recognized as MathType Equation.DSMT4.");

            oldRange = oldShape.Range.Duplicate;
            if (alignInline)
            {
                // InsertXML can silently import paragraph formatting from the
                // serialized Equation.DSMT4 fragment even when it does not split
                // the host paragraph. Preserve only paragraph-layout properties;
                // reapplying the paragraph Style to an OLE-containing range can
                // make Word/MathType regenerate the OLE presentation geometry.
                inlineParagraphLayout = CaptureMathTypeDisplayParagraphLayout(oldShape);
            }
            // The edit owns one local paragraph, not a document-wide count.
            // Keep an outside boundary and the original paragraph serialization
            // before the first mutation, including existing revision/bookmark data.
            localEdit = MathTypeLocalEditScope.Capture(document, oldRange);
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
            // The native renderer is required below. Do not seed any Word
            // presentation geometry from the frontend: its exported picture
            // uses a different layout engine and coordinate system.
            float targetWidthPt;
            float targetHeightPt;
            float targetOriginalWidthPt;
            float targetOriginalHeightPt;
            int targetWordPosition;
            byte[] previewWmf;

            // Preserve the original Equation.DSMT4 CFB, replace only its MTEF
            // structure, and obtain its native presentation through the private
            // data-only service. Word never activates a MathType COM server here.
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
                MathTypeNativePreviewRenderer.TryRenderBatchWithRetry(
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
                // Existing Word documents may scale MathType's native drawing
                // and may also store a different intrinsic OLE extent in
                // dxaOrig/dyaOrig. Both relationships belong to the document's
                // presentation. Preserve them across the edit; only the rewritten
                // mathematical bounds come from MathPage. The existing Word run
                // position is likewise presentation state and must not be replaced
                // by MathPage's intrinsic baseline for an already placed object.
                var widthScale = CalculateMathTypeNativePresentationScale(
                    sourceFragment.WidthPt,
                    sourceNativePreview.WidthPt);
                var heightScale = CalculateMathTypeNativePresentationScale(
                    sourceFragment.HeightPt,
                    sourceNativePreview.HeightPt);
                targetWidthPt = Math.Max(1f, nativePreview.WidthPt * widthScale);
                targetHeightPt = Math.Max(1f, nativePreview.HeightPt * heightScale);
                targetOriginalWidthPt = CalculateMathTypePreservedOriginalExtent(
                    sourceFragment.OriginalWidthPt,
                    sourceFragment.WidthPt,
                    targetWidthPt);
                targetOriginalHeightPt = CalculateMathTypePreservedOriginalExtent(
                    sourceFragment.OriginalHeightPt,
                    sourceFragment.HeightPt,
                    targetHeightPt);
                targetWordPosition = originalWordPosition;
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-replace-native-preview formulaId={session.FormulaId} "
                    + $"sourceWord={sourceFragment.WidthPt:0.###}x{sourceFragment.HeightPt:0.###}@{originalWordPosition} "
                    + $"sourceOriginal={sourceFragment.OriginalWidthPt:0.###}x{sourceFragment.OriginalHeightPt:0.###} "
                    + $"sourceNative={sourceNativePreview.WidthPt:0.###}x{sourceNativePreview.HeightPt:0.###}@{sourceNativePreview.WordPosition} "
                    + $"rewrittenNative={nativePreview.WidthPt:0.###}x{nativePreview.HeightPt:0.###}@{nativePreview.WordPosition} "
                    + $"scale={widthScale:0.###}x{heightScale:0.###} "
                    + $"target={targetWidthPt:0.###}x{targetHeightPt:0.###}@{targetWordPosition} "
                    + $"original={targetOriginalWidthPt:0.###}x{targetOriginalHeightPt:0.###}");
            }
            else
            {
                throw new InvalidOperationException(MathTypeNativePreviewRenderer.FailureMessage);
            }


            // Rewrite the existing Equation.DSMT4 CFB and its external Word WMF
            // presentation in one offline Flat OPC transaction using the native
            // MathType geometry returned by the private renderer.
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

            replacementWordOpenXml = MathTypeLocalFragment.Prepare(
                replacementWordOpenXml,
                localEdit.ParagraphXml,
                localEdit.LocatorName,
                localEdit.SourceOleOrdinal,
                localEdit.SourceOleCount);
            undoHistory = new WordUndoHistorySnapshot(document);

            // WordOpenXML exports above may terminate Word's custom Undo record.
            // Open it only after every source/template export and offline render,
            // immediately before the first document mutation.
            undoRecord = BeginUndoRecord("VisualTeX Update MathType OLE Formula")
                ?? throw new InvalidOperationException("Word could not establish an independent MathType edit transaction.");

            // Update the exact existing OLE range in one native replacement.
            // Its Word span can include an automatic EMBED field (26 characters
            // in real legacy documents), not just the visible object character.
            // Replacing an OLE's internals must not create a tracked deletion plus
            // a second visible object. Preserve existing revision wrappers from
            // the source XML and suspend only new tracking during this transaction.
            previousTrackRevisions = document.TrackRevisions;
            if (previousTrackRevisions)
            {
                document.TrackRevisions = false;
                trackingSuspended = true;
            }
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
                $"mathtype-replace-stage stage=insert-rewritten-flat-opc formulaId={session.FormulaId}");
            mutationStarted = true;
            WpsXmlInsertion.Replace(oldRange, replacementWordOpenXml);
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=insert-rewritten-flat-opc-complete formulaId={session.FormulaId}");

            insertion = localEdit.ResolveInsertedRange();
            localEdit.RepairAndVerifyParagraph(insertion);
            // Removing Word's synthetic paragraph mark can invalidate an OLE RCW;
            // acquire the object only after the local boundary is stable.
            Release(insertion);
            insertion = localEdit.ResolveInsertedRange();
            InlineShapes? localShapes = null;
            try
            {
                localShapes = insertion.InlineShapes;
                replacement = localShapes[1];
            }
            finally { Release(localShapes); }
            if (!MathTypeOleInterop.IsMathTypeOle(replacement))
                throw new InvalidOperationException("Word did not materialize the owned native MathType OLE.");
            WordDoubleClickHook.TraceMessage(
                $"mathtype-replace-stage stage=resolve-replacement-complete formulaId={session.FormulaId}");
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

            // Preserve the existing Word run position independently of the
            // native mathematical bounds, as required by the CP04 geometry model.
            SetInlineOleWordPosition(replacement, targetWordPosition);
            if (alignInline)
                RestoreTypingBaselineAfter(replacement);

            // Complete view/baseline changes inside the same Undo item. The final
            // materialized XML read remains mandatory but happens after that item
            // closes, so it cannot split content replacement from font writes.
            finalSelection = replacement.Range.Duplicate;
            localEdit.RestoreContainedBookmarks(finalSelection);
            localEdit.RemoveLocator();
            if (trackingSuspended)
            {
                document.TrackRevisions = previousTrackRevisions;
                trackingSuspended = false;
            }
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
        catch (Exception primary)
        {
            editViewRestored = false;
            Release(finalSelection); finalSelection = null;
            try
            {
                if (!editUndoEnded && undoRecord is not null)
                {
                    EndUndoRecord(undoRecord);
                    editUndoEnded = true;
                }
                if (mutationStarted && document is not null && undoHistory is not null)
                {
                    // Never insert another serialized source object as rollback.
                    // Undo only this operation's proven history prefix, then verify
                    // the original local content, geometry, revisions and bookmarks.
                    undoHistory.UndoChanges(document);
                    localEdit!.VerifyRollback();
                }
            }
            catch (Exception recovery)
            {
                throw new AggregateException(
                    "The MathType edit failed and its local Word undo recovery could not be verified.",
                    primary, recovery);
            }
            throw;
        }
        finally
        {
            nativePreview?.Dispose();
            sourceNativePreview?.Dispose();
            if (trackingSuspended && document is not null)
                document.TrackRevisions = previousTrackRevisions;
            localEdit?.Dispose();
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
                try { _application.ScreenUpdating = previousScreenUpdating; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
                try { _application.ScreenUpdating = previousScreenUpdating; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
        catch (Exception ignored)
        {
            // View restoration must never turn an otherwise successful formula
            // conversion into an error. The original document is reactivated by
            // the conversion write path before this method normally runs.
            IgnoredErrors.Ignore(ignored);
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
            try { window = _application.ActiveWindow; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            int? vertical = null;
            int? horizontal = null;
            try { vertical = window?.VerticalPercentScrolled; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            try { horizontal = window?.HorizontalPercentScrolled; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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
                var capturedStart = state.Shifts?.Map(state.SelectionStart) ?? state.SelectionStart;
                var capturedEnd = state.Shifts?.Map(state.SelectionEnd) ?? state.SelectionEnd;
                var start = Math.Max(content.Start, Math.Min(capturedStart, content.End));
                var end = Math.Max(start, Math.Min(capturedEnd, content.End));
                object startValue = start;
                object endValue = end;
                fallback = document.Range(ref startValue, ref endValue);
                selection.SetRange(fallback.Start, fallback.End);
            }
            try { window = _application.ActiveWindow; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            if (window is not null)
            {
                try
                {
                    if (state.HorizontalPercentScrolled.HasValue
                        && window.HorizontalPercentScrolled != state.HorizontalPercentScrolled.Value)
                        window.HorizontalPercentScrolled = state.HorizontalPercentScrolled.Value;
                }
                catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                try
                {
                    if (state.VerticalPercentScrolled.HasValue
                        && window.VerticalPercentScrolled != state.VerticalPercentScrolled.Value)
                        window.VerticalPercentScrolled = state.VerticalPercentScrolled.Value;
                }
                catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            }
        }
        catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
        finally
        {
            Release(window);
            Release(content);
            Release(fallback);
            Release(selection);
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
        catch (Exception ignored)
        {
            // The MathType OLE is already structurally valid. A transient Word
            // insertion-format refusal must not invalidate the conversion.
            IgnoredErrors.Ignore(ignored);
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
        catch (Exception ignored)
        {
            // Structural caret placement remains useful even if Word rejects a
            // transient insertion-format mutation at an unusual protected range.
            IgnoredErrors.Ignore(ignored);
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
        catch (Exception ignored)
        {
            // Keep the caret outside the formula even if Word refuses to mutate
            // insertion formatting at an unusual protected boundary.
            IgnoredErrors.Ignore(ignored);
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
            if (!WordNumberedDisplayLayout.RangeIsWhollyWithinTable(sourceRange))
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
                WordNumberedTypingParagraphs.EnsureNormalTypingParagraphAfterNumberedDisplay(
                    document,
                    formulaId!);
            if (safeTypingRange is not null)
            {
                var result = safeTypingRange;
                safeTypingRange = null;
                return result;
            }

            bookmarks = document.Bookmarks;
            var captionName = WordNumberingFields.NativeCaptionBookmarkName(
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
            catch (InvalidDataException ignored) { IgnoredErrors.Ignore(ignored); }

            if (host is null)
            {
                try
                {
                    host =
                        WordFormulaHostLookup.ResolveLocal(
                            document,
                            center,
                            WordFormulaHostKind.Omml);
                }
                catch (InvalidDataException ignored) { IgnoredErrors.Ignore(ignored); }
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

    internal static bool TryParseRangeReference(
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
        if (IsSameSourceDocument(document, expectedIdentity!)) return;
        throw new InvalidOperationException(
            "The active Word document changed while the VisualTeX editor was open.");
    }

    // Saving, Save As or a OneDrive move changes the path while the editor is
    // open; the same live document object is still the source.
    internal static bool IsSameSourceDocument(Document document, string expectedIdentity)
    {
        if (string.Equals(DocumentIdentity(document), expectedIdentity, StringComparison.OrdinalIgnoreCase))
            return true;
        IntPtr expectedObject;
        lock (IdentityObjects)
            if (!IdentityObjects.TryGetValue(expectedIdentity, out expectedObject)) expectedObject = IntPtr.Zero;
        return expectedObject != IntPtr.Zero && expectedObject == ObjectIdentity(document);
    }

    // Identity string -> the COM identity (IUnknown) of the document that had it.
    private static readonly Dictionary<string, IntPtr> IdentityObjects = new(StringComparer.OrdinalIgnoreCase);

    private static IntPtr ObjectIdentity(Document document)
    {
        try
        {
            var unknown = Marshal.GetIUnknownForObject(document);
            Marshal.Release(unknown);
            return unknown;
        }
        catch { return IntPtr.Zero; }
    }

    private static string DocumentIdentity(Document document)
    {
        string identity;
        try
        {
            var fullName = ReadWordStateWithRetry(() => document.FullName);
            identity = !string.IsNullOrWhiteSpace(fullName) ? fullName : ReadWordStateWithRetry(() => document.Name);
        }
        catch { identity = ReadWordStateWithRetry(() => document.Name); }
        var objectIdentity = ObjectIdentity(document);
        if (objectIdentity != IntPtr.Zero)
            lock (IdentityObjects) IdentityObjects[identity] = objectIdentity;
        return identity;
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
        catch (Exception ignored)
        {
            // Hidden/protected automation windows can reject View mutations. The
            // formula operation itself remains valid; only interactive presentation
            // normalization is best-effort in that environment.
            IgnoredErrors.Ignore(ignored);
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
                WordNumberedDisplayLayout.GetManagedNumberTableRowIndex(table, formulaRange, 2), 2);
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
            WpsXmlInsertion.Insert(insertion, wordOpenXml);

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
        try { shape.Delete(); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
    }

    private static void TryDelete(Table? table)
    {
        if (table is null) return;
        try { table.Delete(); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
    }

    private static void TryDelete(Range? range)
    {
        if (range is null) return;
        try { range.Delete(); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
    }

    internal static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        // Office may return the same RCW to the host and to this service.
        // FinalReleaseComObject would invalidate every shared reference in the
        // add-in AppDomain, so release only the reference acquired here.
        try { Marshal.ReleaseComObject(value); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
    }
}
