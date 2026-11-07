using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Application = Microsoft.Office.Interop.Word.Application;
using Task = System.Threading.Tasks.Task;

namespace VisualTeX.WpsWriterVsto;

/// <summary>
/// Minimal WPS Writer integration for MathType OLE equations: create, edit and
/// renumber. All COM work is marshalled back to the host STA thread; the editor
/// Session itself runs on the background thread.
/// </summary>
internal sealed class WpsWriterFormulaService : IDisposable
{
    private const string NumberFormatVariable = "VT_EquationNumberFormat";
    private const string BookmarkPrefix = "VTMATH_";
    private static readonly TimeSpan SessionTimeout = TimeSpan.FromMinutes(30);

    private readonly Application _application;
    private readonly VisualTeXSessionClient _client;
    private readonly OfficeUiDispatcher _dispatcher;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private bool _disposed;

    internal WpsWriterFormulaService(
        Application application,
        VisualTeXSessionClient client,
        OfficeUiDispatcher dispatcher)
    {
        _application = application;
        _client = client;
        _dispatcher = dispatcher;
    }

    internal async Task CreateAsync(string displayMode, CancellationToken cancellationToken)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var capture = await _dispatcher
                .InvokeAsync(() => CaptureCreate(displayMode))
                .ConfigureAwait(false);
            if (capture is null) return;

            var committed = await RunEditorSessionAsync(capture.Request, cancellationToken)
                .ConfigureAwait(false);
            if (committed is null) return;

            var sessionId = committed.Id;
            var session = committed;
            var mathMl = session.ExportResult?.MathMl;
            if (string.IsNullOrWhiteSpace(mathMl))
                throw new InvalidDataException(
                    "VisualTeX 没有返回 MathML，无法创建 MathType 公式。");

            var prepared = PrepareMathType(session, mathMl!, inline: capture.Inline);
            try
            {
                await _dispatcher
                    .InvokeAsync(() => ApplyCreate(capture, session, prepared))
                    .ConfigureAwait(false);
                await _client.CompleteAsync(sessionId, cancellationToken).ConfigureAwait(false);
                WpsLog.Write($"Created MathType formula {session.FormulaId}.");
            }
            catch (Exception error)
            {
                await TryFailAsync(sessionId, error).ConfigureAwait(false);
                throw;
            }
            finally
            {
                await TryCloseEditorAsync(sessionId).ConfigureAwait(false);
            }
        }
        finally
        {
            _gate.Release();
        }
    }

    internal async Task EditSelectedAsync(CancellationToken cancellationToken)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var capture = await _dispatcher
                .InvokeAsync(CaptureEdit)
                .ConfigureAwait(false);
            if (capture is null) return;

            var request = new CreateVstoSessionRequest
            {
                Mode = "edit",
                Host = "word",
                FormulaId = capture.Metadata.FormulaId,
                SourceDocumentId = capture.DocumentKey,
                SourceObjectId = capture.ObjectKey,
                Title = capture.Metadata.Title,
                Lines = capture.Metadata.Lines,
                CodeFormat = capture.Metadata.CodeFormat,
                DisplayMode = capture.Metadata.DisplayMode,
                ObjectMode = FormulaOleContract.MathTypeOleMode,
                Numbered = capture.Metadata.Numbered,
                MathTypeNumberPosition = "right",
                FontSizePt = capture.Metadata.FontSizePt ?? FormulaFontSize.DefaultPt,
                OriginalMetadata = capture.Metadata,
                AutoCommitOnClose = true,
            };

            var committed = await RunEditorSessionAsync(request, cancellationToken)
                .ConfigureAwait(false);
            if (committed is null) return;

            var sessionId = committed.Id;
            var session = committed;
            var mathMl = session.ExportResult?.MathMl;
            if (string.IsNullOrWhiteSpace(mathMl))
                throw new InvalidDataException(
                    "VisualTeX 没有返回 MathML，无法更新 MathType 公式。");

            var prepared = PrepareMathType(
                session,
                mathMl!,
                inline: string.Equals(session.DisplayMode, "inline", StringComparison.OrdinalIgnoreCase));
            try
            {
                await _dispatcher
                    .InvokeAsync(() => ApplyEdit(capture, session, prepared))
                    .ConfigureAwait(false);
                await _client.CompleteAsync(sessionId, cancellationToken).ConfigureAwait(false);
                WpsLog.Write($"Updated MathType formula {session.FormulaId}.");
            }
            catch (Exception error)
            {
                await TryFailAsync(sessionId, error).ConfigureAwait(false);
                throw;
            }
            finally
            {
                await TryCloseEditorAsync(sessionId).ConfigureAwait(false);
            }
        }
        finally
        {
            _gate.Release();
        }
    }

    private async Task<OfficeSessionDocument?> RunEditorSessionAsync(
        CreateVstoSessionRequest request,
        CancellationToken cancellationToken)
    {
        await _client.EnsureHealthyAsync(cancellationToken).ConfigureAwait(false);
        var session = await _client.CreateSessionAsync(request, cancellationToken)
            .ConfigureAwait(false);
        await _client.OpenEditorAsync(session.Id, cancellationToken).ConfigureAwait(false);
        var committed = await _client
            .WaitForCommitAsync(session.Id, SessionTimeout, cancellationToken)
            .ConfigureAwait(false);

        switch (committed.Status)
        {
            case "cancelled":
                WpsLog.Write($"Session {session.Id} cancelled by the user.");
                await TryCloseEditorAsync(session.Id).ConfigureAwait(false);
                return null;
            case "failed":
                await TryCloseEditorAsync(session.Id).ConfigureAwait(false);
                throw new InvalidOperationException(
                    committed.Error ?? "VisualTeX 公式编辑会话失败。");
            case "committing":
            case "completed":
                return committed;
            default:
                throw new InvalidOperationException(
                    $"VisualTeX 会话返回了未知状态：{committed.Status}");
        }
    }

    private static PreparedMathType PrepareMathType(
        OfficeSessionDocument session,
        string mathMl,
        bool inline)
    {
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(
            mathMl,
            inline,
            session.FontSizePt);
        var compoundFile = MathTypeOleStorage.CreateStandaloneCompoundFile(generated);
        var root = Path.Combine(Path.GetTempPath(), "visualtex-wps-mathtype");
        Directory.CreateDirectory(root);
        if (!MathTypeNativePreviewRenderer.TryRender(
                generated.Mtef,
                root,
                out var preview))
        {
            throw new InvalidOperationException(
                MathTypeNativePreviewRenderer.FailureMessage);
        }

        try
        {
            return new PreparedMathType
            {
                CompoundFile = compoundFile,
                PreviewWmf = File.ReadAllBytes(preview.WmfPath),
                WidthPt = preview.WidthPt,
                HeightPt = preview.HeightPt,
            };
        }
        finally
        {
            preview.Dispose();
        }
    }

    private sealed class PreparedMathType
    {
        internal byte[] CompoundFile = Array.Empty<byte>();
        internal byte[] PreviewWmf = Array.Empty<byte>();
        internal float WidthPt;
        internal float HeightPt;
    }

    private sealed class CreateCapture
    {
        internal string DocumentKey = string.Empty;
        internal int InsertionStart;
        internal bool Inline;
        internal string NumberFormatId = "off";
        internal CreateVstoSessionRequest Request = new();
    }

    private sealed class EditCapture
    {
        internal string DocumentKey = string.Empty;
        internal string ObjectKey = string.Empty;
        internal int ParagraphStart;
        internal bool Display;
        internal FormulaMetadata Metadata = new();
    }

    private CreateCapture? CaptureCreate(string displayMode)
    {
        Document? document = null;
        Selection? selection = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("当前没有打开的文档。");
            selection = _application.Selection;
            var inline = string.Equals(displayMode, "inline", StringComparison.OrdinalIgnoreCase);
            var numberFormatId = inline ? "off" : ResolveNumberFormatIdCore(document);
            int insertionStart;
            Range? insertion = null;
            try
            {
                insertion = selection.Range.Duplicate;
                insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
                insertionStart = insertion.Start;
            }
            finally
            {
                ReleaseComObject(insertion);
            }
            var fontSizePt = ReadCaretFontSize(selection);
            var formulaId = Guid.NewGuid().ToString("D");
            var documentKey = ResolveDocumentKey(document);

            var line = new FormulaLine
            {
                Id = Guid.NewGuid().ToString("D"),
                Latex = string.Empty,
            };
            var request = new CreateVstoSessionRequest
            {
                Mode = "create",
                Host = "word",
                FormulaId = formulaId,
                SourceDocumentId = documentKey,
                Title = "MathType Formula",
                Lines = new List<FormulaLine> { line },
                ActiveLineId = line.Id,
                CodeFormat = "raw",
                DisplayMode = inline ? "inline" : "block",
                ObjectMode = FormulaOleContract.MathTypeOleMode,
                Numbered = !inline && numberFormatId != "off",
                MathTypeNumberPosition = "right",
                FontSizePt = fontSizePt,
                AutoCommitOnClose = true,
            };

            return new CreateCapture
            {
                DocumentKey = documentKey,
                InsertionStart = insertionStart,
                Inline = inline,
                NumberFormatId = numberFormatId,
                Request = request,
            };
        }
        finally
        {
            ReleaseComObject(selection);
            ReleaseComObject(document);
        }
    }

    private EditCapture? CaptureEdit()
    {
        Document? document = null;
        Selection? selection = null;
        InlineShape? shape = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("当前没有打开的文档。");
            selection = _application.Selection;
            if (selection.InlineShapes is null || selection.InlineShapes.Count != 1)
                throw new InvalidOperationException("请先选择一个 MathType 公式。");
            shape = selection.InlineShapes[1];
            if (!IsMathTypeShape(shape))
                throw new InvalidOperationException("所选对象不是 MathType 公式。");

            byte[] compoundFile;
            using (var transaction = MathTypeOleStorage.BeginClipboardTransaction(() =>
            {
                Range? copyRange = null;
                try
                {
                    copyRange = shape.Range;
                    copyRange.Copy();
                }
                finally
                {
                    ReleaseComObject(copyRange);
                }
            }))
            {
                compoundFile = transaction.CompoundFile;
            }

            var native = MathTypeOleStorage.ReadEquationNative(compoundFile);
            var mathMl = MathTypeMtefCodec.ReadEquationNativeMathMl(native);
            var latex = MathMlToLatexConverter.Convert(mathMl).Trim();
            if (string.IsNullOrWhiteSpace(latex))
                throw new InvalidDataException("所选 MathType 公式的内容为空或不受支持。");

            var fontSizePt = MathTypeMtefCodec.ReadEquationNativeFullFontSize(native);
            var display = InferDisplayMode(shape, out var numbered);
            var now = DateTimeOffset.UtcNow.ToString("O");
            var documentKey = ResolveDocumentKey(document);
            var objectKey = $"{documentKey}:{shape.Range.Start}:{shape.Range.End}";
            var formulaId = FindFormulaBookmark(document, shape.Range)
                ?? Guid.NewGuid().ToString("D");

            var widthPt = Math.Max(1d, shape.Width);
            var heightPt = Math.Max(1d, shape.Height);
            var metadata = new FormulaMetadata
            {
                FormulaId = formulaId,
                Title = "MathType Formula",
                Latex = latex,
                Lines = new List<FormulaLine>
                {
                    new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
                },
                CodeFormat = "raw",
                DisplayMode = display ? "block" : "inline",
                Numbered = display && numbered,
                FontSizePt = fontSizePt,
                RenderFontSizePt = fontSizePt,
                RenderWidthPx = widthPt / 0.75d,
                RenderHeightPx = heightPt / 0.75d,
                CreatedWithVersion = "1.2.9",
                UpdatedWithVersion = "1.2.9",
                CreatedAt = now,
                UpdatedAt = now,
            };
            metadata.Validate();

            return new EditCapture
            {
                DocumentKey = documentKey,
                ObjectKey = objectKey,
                ParagraphStart = shape.Range.Start,
                Display = display,
                Metadata = metadata,
            };
        }
        finally
        {
            ReleaseComObject(shape);
            ReleaseComObject(selection);
            ReleaseComObject(document);
        }
    }

    private void ApplyCreate(
        CreateCapture capture,
        OfficeSessionDocument session,
        PreparedMathType prepared)
    {
        Document? document = null;
        Range? insertion = null;
        try
        {
            document = FindDocument(capture.DocumentKey)
                ?? throw new InvalidOperationException("原始文档已关闭。");
            var position = Math.Min(capture.InsertionStart, document.Content.End);
            insertion = document.Range(position, position);
            var bookmark = BookmarkPrefix + session.FormulaId.Replace("-", string.Empty);
            InsertFlatOpcOrClipboard(
                document,
                insertion,
                prepared,
                display: !capture.Inline,
                numberFormatId: capture.NumberFormatId,
                bookmarkName: bookmark);

            var shape = FindMathTypeShapeNear(document, position)
                ?? throw new InvalidOperationException(
                    "WPS 未能物化 MathType OLE 对象。若此错误持续出现，请运行 wps_compat_probe.ps1 检查 InsertXML/OLE 支持。");
            try
            {
                try { document.Bookmarks.Add(bookmark, shape.Range); }
                catch (Exception error) { WpsLog.WriteException("Bind formula bookmark failed", error); }
                TryCacheMetadata(shape, session.ToMetadata());
                if (!capture.Inline) FormatDisplayParagraph(document, shape);
            }
            finally
            {
                ReleaseComObject(shape);
            }
        }
        finally
        {
            ReleaseComObject(insertion);
            ReleaseComObject(document);
        }
    }

    private void ApplyEdit(
        EditCapture capture,
        OfficeSessionDocument session,
        PreparedMathType prepared)
    {
        Document? document = null;
        InlineShape? existing = null;
        Range? insertion = null;
        try
        {
            document = FindDocument(capture.DocumentKey)
                ?? throw new InvalidOperationException("原始文档已关闭。");
            existing = FindMathTypeShapeAt(document, capture.ParagraphStart);
            if (existing is null)
                throw new InvalidOperationException("所选 MathType 公式已不存在。");

            int insertPosition;
            if (capture.Display)
            {
                Paragraphs? paragraphs = null;
                Paragraph? paragraph = null;
                Range? paragraphRange = null;
                try
                {
                    paragraphs = existing.Range.Paragraphs;
                    paragraph = paragraphs[1];
                    paragraphRange = paragraph.Range;
                    insertPosition = paragraphRange.Start;
                    // Clear content but keep the paragraph mark; this also removes
                    // the previous tab stops and number fields.
                    paragraphRange.Text = string.Empty;
                }
                finally
                {
                    ReleaseComObject(paragraphRange);
                    ReleaseComObject(paragraph);
                    ReleaseComObject(paragraphs);
                }
            }
            else
            {
                insertPosition = existing.Range.Start;
                existing.Delete();
            }

            insertPosition = Math.Min(insertPosition, document.Content.End);
            insertion = document.Range(insertPosition, insertPosition);
            var bookmark = BookmarkPrefix + session.FormulaId.Replace("-", string.Empty);
            InsertFlatOpcOrClipboard(
                document,
                insertion,
                prepared,
                display: capture.Display,
                numberFormatId: capture.Display ? ResolveNumberFormatIdCore(document) : "off",
                bookmarkName: bookmark);

            var shape = FindMathTypeShapeNear(document, insertPosition)
                ?? throw new InvalidOperationException("WPS 未能写入更新后的 MathType 公式。");
            try
            {
                try { document.Bookmarks.Add(bookmark, shape.Range); }
                catch (Exception error) { WpsLog.WriteException("Rebind formula bookmark failed", error); }
                TryCacheMetadata(shape, session.ToMetadata());
                if (capture.Display) FormatDisplayParagraph(document, shape);
            }
            finally
            {
                ReleaseComObject(shape);
            }
        }
        finally
        {
            ReleaseComObject(insertion);
            ReleaseComObject(existing);
            ReleaseComObject(document);
        }
    }

    private static void FormatDisplayParagraph(Document document, InlineShape shape)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        PageSetup? pageSetup = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            paragraphs = shape.Range.Paragraphs;
            paragraph = paragraphs[1];
            paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            pageSetup = document.PageSetup;
            var rightEdge = pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin;
            if (rightEdge <= 0) return;
            tabStops = paragraph.TabStops;
            tabStop = tabStops.Add(
                rightEdge,
                WdTabAlignment.wdAlignTabRight,
                WdTabLeader.wdTabLeaderSpaces);
        }
        catch (Exception error)
        {
            WpsLog.WriteException("Formatting display paragraph failed", error);
        }
        finally
        {
            ReleaseComObject(tabStop);
            ReleaseComObject(tabStops);
            ReleaseComObject(pageSetup);
            ReleaseComObject(paragraph);
            ReleaseComObject(paragraphs);
        }
    }

    private void InsertFlatOpcOrClipboard(
        Document document,
        Range insertion,
        PreparedMathType prepared,
        bool display,
        string numberFormatId,
        string bookmarkName)
    {
        var flatOpc = WpsMathTypeFlatOpc.Build(
            prepared.CompoundFile,
            prepared.PreviewWmf,
            prepared.WidthPt,
            prepared.HeightPt,
            display,
            numberFormatId,
            bookmarkName);

        try
        {
            insertion.InsertXML(flatOpc);
            return;
        }
        catch (Exception error)
        {
            WpsLog.WriteException("Range.InsertXML failed; trying clipboard paste", error);
        }

        selectionPaste(document, insertion);
        void selectionPaste(Document target, Range targetRange)
        {
            Range? caret = null;
            try
            {
                caret = targetRange.Duplicate;
                caret.Select();
                using var transaction = MathTypeOleStorage.BeginStandaloneClipboardTransaction(
                    prepared.CompoundFile,
                    null,
                    prepared.WidthPt,
                    prepared.HeightPt);
                _application.Selection.Paste();
            }
            finally
            {
                ReleaseComObject(caret);
            }
        }
    }

    internal void UpdateEquationNumbers()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(WpsWriterFormulaService));
        Document? document = null;
        Fields? fields = null;
        try
        {
            document = _application.ActiveDocument;
            if (document is null) return;
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    var instruction = code?.Text ?? string.Empty;
                    if (instruction.IndexOf("VisualTeXEquation", StringComparison.OrdinalIgnoreCase) >= 0
                        || instruction.IndexOf("STYLEREF", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        field.Update();
                    }
                }
                catch (Exception error)
                {
                    WpsLog.WriteException($"Updating field {index} failed", error);
                }
                finally
                {
                    ReleaseComObject(code);
                    ReleaseComObject(field);
                }
            }
            document.Fields.Update();
        }
        finally
        {
            ReleaseComObject(fields);
            ReleaseComObject(document);
        }
    }

    internal bool IsMathTypeSelection(Selection selection)
    {
        InlineShape? shape = null;
        try
        {
            if (selection?.InlineShapes is null || selection.InlineShapes.Count != 1)
                return false;
            shape = selection.InlineShapes[1];
            return IsMathTypeShape(shape);
        }
        catch
        {
            return false;
        }
        finally
        {
            ReleaseComObject(shape);
        }
    }

    internal string DescribeSelection(Selection selection)
    {
        InlineShape? shape = null;
        try
        {
            if (selection?.InlineShapes is null || selection.InlineShapes.Count != 1)
                return string.Empty;
            shape = selection.InlineShapes[1];
            return $"{selection.Document?.Name}:{shape.Range.Start}:{shape.Range.End}";
        }
        catch
        {
            return string.Empty;
        }
        finally
        {
            ReleaseComObject(shape);
        }
    }

    internal string ResolveNumberFormatId()
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument;
            return document is null ? "off" : ResolveNumberFormatIdCore(document);
        }
        finally
        {
            ReleaseComObject(document);
        }
    }

    internal void SetNumberFormat(string formatId)
    {
        var normalized = WpsMathTypeFlatOpc.IsKnownNumberFormat(formatId) ? formatId : "off";
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("当前没有打开的文档。");
            Variable? variable = null;
            try
            {
                variable = document.Variables[NumberFormatVariable];
                variable.Value = normalized;
            }
            catch
            {
                try { document.Variables.Add(NumberFormatVariable, normalized); }
                catch (Exception error) { WpsLog.WriteException("Persist number format failed", error); }
            }
            finally
            {
                ReleaseComObject(variable);
            }
        }
        finally
        {
            ReleaseComObject(document);
        }
    }

    private static string ResolveNumberFormatIdCore(Document document)
    {
        Variable? variable = null;
        try
        {
            variable = document.Variables[NumberFormatVariable];
            var value = variable?.Value as string;
            return WpsMathTypeFlatOpc.IsKnownNumberFormat(value) ? value! : "off";
        }
        catch
        {
            return "off";
        }
        finally
        {
            ReleaseComObject(variable);
        }
    }

    private static bool IsMathTypeShape(InlineShape shape)
    {
        OLEFormat? format = null;
        try
        {
            if (shape.Type is not WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                and not WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                return false;
            try
            {
                format = shape.OLEFormat;
                var progId = format?.ProgID;
                if (!string.IsNullOrWhiteSpace(progId)
                    && (string.Equals(
                            progId,
                            MathTypeOleInterop.CanonicalProgId,
                            StringComparison.OrdinalIgnoreCase)
                        || MathTypeOleInterop.TryResolveCapabilities(progId, out _)))
                {
                    return true;
                }
            }
            finally
            {
                ReleaseComObject(format);
                format = null;
            }

            using var transaction = MathTypeOleStorage.BeginClipboardTransaction(() =>
            {
                Range? range = null;
                try
                {
                    range = shape.Range;
                    range.Copy();
                }
                finally
                {
                    ReleaseComObject(range);
                }
            });
            return MathTypeOleStorage.LooksLikeMathTypeCompoundFile(transaction.CompoundFile);
        }
        catch
        {
            return false;
        }
        finally
        {
            ReleaseComObject(format);
        }
    }

    private static bool InferDisplayMode(InlineShape shape, out bool numbered)
    {
        numbered = false;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        try
        {
            paragraphs = shape.Range.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var text = (paragraphRange.Text ?? string.Empty)
                .Replace("\r", string.Empty)
                .Replace("\a", string.Empty)
                .Replace("\u0001", string.Empty)
                .Replace("\uFFFC", string.Empty)
                .Replace("\u0013", string.Empty)
                .Replace("\u0014", string.Empty)
                .Replace("\u0015", string.Empty)
                .Trim();
            var display = text.Length == 0;
            if (display)
            {
                fields = paragraphRange.Fields;
                for (var index = 1; index <= fields.Count; index++)
                {
                    Field? field = null;
                    Range? code = null;
                    try
                    {
                        field = fields[index];
                        code = field.Code;
                        if ((code?.Text ?? string.Empty).IndexOf(
                                "VisualTeXEquation",
                                StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            numbered = true;
                            break;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(code);
                        ReleaseComObject(field);
                    }
                }
            }
            return display;
        }
        catch
        {
            return false;
        }
        finally
        {
            ReleaseComObject(fields);
            ReleaseComObject(paragraphRange);
            ReleaseComObject(paragraph);
            ReleaseComObject(paragraphs);
        }
    }

    private static string? FindFormulaBookmark(Document document, Range shapeRange)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    bookmark = bookmarks[index];
                    var name = bookmark.Name ?? string.Empty;
                    if (!name.StartsWith(BookmarkPrefix, StringComparison.Ordinal)) continue;
                    range = bookmark.Range;
                    if (range.Start <= shapeRange.Start && range.End >= shapeRange.End)
                        return name.Substring(BookmarkPrefix.Length);
                }
                finally
                {
                    ReleaseComObject(range);
                    ReleaseComObject(bookmark);
                }
            }
            return null;
        }
        finally
        {
            ReleaseComObject(bookmarks);
        }
    }

    private static InlineShape? FindMathTypeShapeNear(Document document, int position)
    {
        InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    var start = shape.Range.Start;
                    if (start < position - 1 || start > position + 8) continue;
                    if (IsMathTypeShape(shape)) return shape;
                }
                catch
                {
                    // Ignore shapes whose Range is no longer readable.
                }
                ReleaseComObject(shape);
                shape = null;
            }
            return null;
        }
        finally
        {
            ReleaseComObject(shapes);
        }
    }

    private static InlineShape? FindMathTypeShapeAt(Document document, int position)
    {
        InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    var start = shape.Range.Start;
                    if (start < position - 8 || start > position + 8) continue;
                    if (IsMathTypeShape(shape)) return shape;
                }
                catch
                {
                    // Ignore shapes whose Range is no longer readable.
                }
                ReleaseComObject(shape);
                shape = null;
            }
            return null;
        }
        finally
        {
            ReleaseComObject(shapes);
        }
    }

    private static void TryCacheMetadata(InlineShape shape, FormulaMetadata metadata)
    {
        try
        {
            metadata.Validate();
            var encoded = FormulaMetadataCodec.Encode(metadata);
            shape.Title = "VisualTeX MathType Formula";
            shape.AlternativeText = encoded;
        }
        catch (Exception error)
        {
            WpsLog.WriteException("Cache formula metadata failed", error);
        }
    }

    private Document? FindDocument(string documentKey)
    {
        Documents? documents = null;
        try
        {
            documents = _application.Documents;
            for (var index = 1; index <= documents.Count; index++)
            {
                Document? document = null;
                try
                {
                    document = documents[index];
                    if (string.Equals(ResolveDocumentKey(document), documentKey, StringComparison.Ordinal))
                        return document;
                }
                catch
                {
                    // Ignore documents that closed while enumerating.
                }
                ReleaseComObject(document);
                document = null;
            }
            return null;
        }
        finally
        {
            ReleaseComObject(documents);
        }
    }

    private static string ResolveDocumentKey(Document document)
    {
        var fullName = document.FullName;
        return string.IsNullOrWhiteSpace(fullName)
            ? "unsaved:" + (document.Name ?? string.Empty)
            : fullName!;
    }

    private static double ReadCaretFontSize(Selection selection)
    {
        try
        {
            var size = selection.Font.Size;
            if (size >= FormulaFontSize.MinimumPt && size <= FormulaFontSize.MaximumPt)
                return FormulaFontSize.Normalize(size);
        }
        catch
        {
            // Mixed or unknown font sizes fall back to the default.
        }
        return FormulaFontSize.DefaultPt;
    }

    private async Task TryFailAsync(string sessionId, Exception error)
    {
        try
        {
            await _client.FailAsync(
                sessionId,
                OfficePluginLanguage.SafeErrorMessage(error.Message),
                CancellationToken.None).ConfigureAwait(false);
        }
        catch (Exception failError)
        {
            WpsLog.WriteException("Marking session failed", failError);
        }
    }

    private async Task TryCloseEditorAsync(string sessionId)
    {
        try
        {
            await _client.CloseEditorAsync(sessionId, CancellationToken.None)
                .ConfigureAwait(false);
        }
        catch (Exception error)
        {
            WpsLog.WriteException("Closing editor window", error);
        }
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            try { Marshal.ReleaseComObject(value); }
            catch { }
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _gate.Dispose();
    }
}
