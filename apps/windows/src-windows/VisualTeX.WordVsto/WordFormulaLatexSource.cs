using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using static VisualTeX.WordVsto.WordFormulaService;
using static VisualTeX.WordVsto.WordInlineOlePosition;

namespace VisualTeX.WordVsto;

/// <summary>
/// Converting formulas to LaTeX text in the document: the formulas a conversion targets, the LaTeX source it
/// writes, checks and binds to an inline baseline, and the numbering frame it leaves behind.
/// </summary>
internal static class WordFormulaLatexSource
{
    internal static List<FormulaToLatexTarget> CaptureFormulaToLatexTargets(
        Document document,
        Range scope,
        bool wholeDocument,
        string objectMode,
        bool refreshOmmlMetadata,
        ISet<string>? ignoredDeletedNumberedOmmlIds = null,
        List<(int Ordinal, string Reason)>? failures = null)
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
        var candidates = 0;
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

                    var ordinal = ++candidates;
                    string? sourceMathMl;
                    FormulaMetadata metadata;
                    try
                    {
                        sourceMathMl =
                            MathTypeOleStorage.ReadMathMl(
                                shape);
                        metadata =
                            MathTypeOleInterop.ReadMetadata(
                                document.Application,
                                shape,
                                sourceMathMl);
                    }
                    catch (Exception error) when (failures is not null
                        && error is not OperationCanceledException and not OutOfMemoryException)
                    {
                        // One unreadable object stays as it is; the others convert.
                        WordDoubleClickHook.TraceMessage($"formula-to-latex-capture-skipped start={formulaRange.Start} error={error}");
                        failures.Add((ordinal, OfficeErrorText.Describe(error)));
                        continue;
                    }

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
                            Ordinal = ordinal,
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
                        throw new InvalidDataException(OfficePluginLanguage.Text(
                            $"公式 {metadata.FormulaId} 的 LaTeX 源码仍位于隐藏编号框架中。为避免源码不可见，转换已回滚。",
                            $"Equation {metadata.FormulaId} still has LaTeX source inside a hidden numbering frame. The conversion was rolled back to avoid invisible source."));
                }
                finally { Release(frame); }
            }
        }
        finally { Release(remaining); }
    }

    internal static string BuildFormulaLatexSource(FormulaMetadata metadata)
    {
        var latex = string.IsNullOrWhiteSpace(metadata.Latex)
            ? string.Join("\n", metadata.Lines.Select(line => line.Latex))
            : metadata.Latex;
        latex = (latex ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Trim();
        if (string.IsNullOrWhiteSpace(latex))
            throw new InvalidDataException(OfficePluginLanguage.Text(
                $"公式 {metadata.FormulaId} 的 LaTeX 元数据为空。为避免删除原公式，转换已中止。",
                $"Equation {metadata.FormulaId} has empty LaTeX metadata. Conversion was stopped to avoid deleting the original equation."));
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

    internal static void BindInlineLatexBaselineProvenance(
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

    internal static bool TryParseInlineLatexBaselineBookmarkName(
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

    internal static void VerifyLatexSourceRange(
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
            throw new InvalidDataException(OfficePluginLanguage.Text(
                $"公式 {formulaId} 的 LaTeX 写回校验失败。Word 实际写入内容与预期源码不一致。",
                $"Verifying the LaTeX written back for equation {formulaId} failed: the text Word wrote differs from the expected source."));
        }

        Frames? frames = null;
        try
        {
            frames = inserted.Frames;
            // Only a hidden (clipped) frame makes the source invisible; a formula
            // the user placed in a visible text frame keeps its source there.
            var hiddenFrame = false;
            for (var index = 1; index <= frames.Count && !hiddenFrame; index++)
            {
                Frame? frame = null;
                try
                {
                    frame = frames[index];
                    hiddenFrame = frame.Width <= 1f && frame.Height <= 1f;
                }
                finally { Release(frame); }
            }
            if (hiddenFrame)
                throw new InvalidDataException(OfficePluginLanguage.Text(
                    $"公式 {formulaId} 的 LaTeX 源码被 Word Frame 包围。为避免源码不可见，本次转换已撤销。",
                    $"The LaTeX source of equation {formulaId} ended up inside a Word frame. The conversion was undone so the source does not become invisible."));
        }
        finally { Release(frames); }
    }

    internal static void ApplyFormulaToLatexEmptySourceInjection(
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

    internal static void ThrowIfFormulaToLatexFailureInjected(
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
            table = WordNumberingReconcile.FindNumberedEquationTableByFormulaOwner(
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
                || !WordNumberedDisplayLayout.TryGetManagedNumberTableRowIndex(
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
                try { format.DisableLineHeightGrid = -1; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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

    internal static void NormalizeLatexSourceRange(
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
                    try { paragraphFormat.DisableLineHeightGrid = -1; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
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

    internal static void ReleaseFormulaToLatexTargets(
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

    internal static bool TryResolveOwningTableCellContentWidth(Range range, out float width)
        => WordFormulaHost.TryGetCellContentWidth(range, out width);
}
