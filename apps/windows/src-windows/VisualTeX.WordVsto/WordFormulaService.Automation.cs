using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Read operations for agents (docs/office-mcp/DESIGN.md §5, §6). They take an
/// explicit document (any open one, not only the active one), use the observational
/// resolvers and readers only, and never change the document: no bookmarks, no
/// identity repair, no selection change. UI thread only, like every Word call.
/// </summary>
internal sealed partial class WordFormulaService
{
    internal const string AutomationVisualTeXKind = "visualtex_ole";
    internal const string AutomationOmmlKind = "office_equation";
    internal const string AutomationMathTypeKind = "mathtype";

    private sealed class AutomationFormula
    {
        internal WordFormulaHostDescriptor Host { get; set; } = null!;
        internal string Kind { get; set; } = string.Empty;
        internal string? Id { get; set; }
    }

    private sealed class AutomationSemantic
    {
        internal string Latex { get; set; } = string.Empty;
        internal IReadOnlyList<string> Lines { get; set; } = Array.Empty<string>();
        internal string Display { get; set; } = "inline";
        internal bool Numbered { get; set; }
        internal string? Number { get; set; }
        internal double? FontSizePt { get; set; }
    }

    internal object AutomationListDocuments()
    {
        var foreground = OfficeAutomationText.ProcessIsForeground();
        var activeIdentity = TryActiveDocumentIdentity();
        var result = new List<object>();
        Documents? documents = null;
        try
        {
            documents = _application.Documents;
            foreach (var document in WordComEnumeration.Each<Document>(documents))
            {
                try
                {
                    if (!HasVisibleWindow(document)) continue;
                    var identity = DocumentIdentity(document);
                    var path = TryRead(() => document.Path) ?? string.Empty;
                    var active = string.Equals(identity, activeIdentity, StringComparison.OrdinalIgnoreCase);
                    result.Add(new
                    {
                        key = OfficeAutomationText.DocumentKey(identity),
                        name = TryRead(() => document.Name) ?? identity,
                        path = path.Length == 0 ? null : identity,
                        active,
                        foreground = foreground && active,
                        saved = TryRead(() => (bool?)document.Saved),
                        read_only = TryRead(() => (bool?)document.ReadOnly),
                    });
                }
                finally { Release(document); }
            }
        }
        finally { Release(documents); }
        return new { documents = result };
    }

    internal object AutomationGetSelection(string? key)
    {
        Document? document = null;
        Window? window = null;
        Selection? selection = null;
        Range? range = null;
        try
        {
            document = FindAutomationDocument(key);
            var identity = DocumentIdentity(document);
            window = document.ActiveWindow;
            selection = window.Selection;
            range = selection.Range;
            var active = string.Equals(identity, TryActiveDocumentIdentity(), StringComparison.OrdinalIgnoreCase);
            var location = AutomationLocation(document, range.StoryType, range.Start);
            var formula = TryResolveAutomationFormula(document, range);
            object? formulaData = null;
            if (formula is not null)
                formulaData = DescribeAutomationFormula(document, formula, 60, full: true);
            var kind = formula is not null ? "formula" : range.Start == range.End ? "none" : "text";
            return new
            {
                document = OfficeAutomationText.DocumentKey(identity),
                foreground = active && OfficeAutomationText.ProcessIsForeground(),
                kind,
                formula = formulaData,
                text_preview = kind == "text" ? Clip(OfficeAutomationText.Clean(range.Text), 200) : null,
                location,
                read_only = document.ReadOnly,
            };
        }
        finally
        {
            Release(range);
            Release(selection);
            Release(window);
            Release(document);
        }
    }

    internal object AutomationListFormulas(
        string key,
        int offset,
        int limit,
        int contextChars,
        HashSet<string> kinds)
    {
        Document? document = null;
        try
        {
            document = FindAutomationDocument(key);
            var formulas = CaptureAutomationFormulas(document, kinds);
            AssignAutomationIds(formulas);
            var numbers = AutomationNumberTargets(document);
            var page = formulas.Skip(offset).Take(limit)
                .Select(formula => DescribeAutomationFormula(document, formula, contextChars, full: false, numbers))
                .ToList();
            return new
            {
                revision = AutomationRevision(document, formulas),
                total = formulas.Count,
                offset,
                formulas = page,
            };
        }
        finally { Release(document); }
    }

    internal object AutomationReadFormula(string key, string formulaId)
    {
        Document? document = null;
        try
        {
            document = FindAutomationDocument(key);
            var formula = ResolveAutomationFormula(document, formulaId);
            return DescribeAutomationFormula(document, formula, 120, full: true);
        }
        finally { Release(document); }
    }

    // Writes: placing the target, then the add-in's own command does the work.

    /// Brings the document's window to the front: the add-in's commands work on
    /// the active document and its selection.
    internal void AutomationActivate(string? key)
    {
        Document? document = null;
        try
        {
            document = FindAutomationDocument(key);
            if (document.ReadOnly)
                throw new OfficeAutomationException("document_read_only", "The document is read-only.");
            document.Activate();
        }
        finally { Release(document); }
    }

    /// Selects a formula (as a click would) and returns it.
    internal object AutomationSelectFormula(string key, string formulaId)
    {
        Document? document = null;
        Range? range = null;
        try
        {
            document = FindAutomationDocument(key);
            document.Activate();
            var formula = ResolveAutomationFormula(document, formulaId);
            range = WordFormulaHostSemanticReader.CreateRange(document, formula.Host.Range);
            range.Select();
            return DescribeAutomationFormula(document, formula, 40, full: false);
        }
        finally
        {
            Release(range);
            Release(document);
        }
    }

    /// Puts the caret where a new formula goes and captures that insertion point.
    /// location: "cursor" (default), "document_end", or an object with
    /// after_formula / before_formula (a formula id) or paragraph_end (a paragraph number).
    internal OfficeSelection AutomationPlaceInsertion(string? key, JsonElement location)
    {
        Document? document = null;
        Range? target = null;
        try
        {
            document = FindAutomationDocument(key);
            if (document.ReadOnly)
                throw new OfficeAutomationException("document_read_only", "The document is read-only.");
            document.Activate();
            var where = location.ValueKind == JsonValueKind.String ? location.GetString() : null;
            if (location.ValueKind == JsonValueKind.Object)
            {
                if (OfficeAutomationText.String(location, "after_formula") is { Length: > 0 } after)
                {
                    var formula = ResolveAutomationFormula(document, after);
                    target = WordFormulaHostSemanticReader.CreateRange(document, formula.Host.Range);
                    target.Collapse(WdCollapseDirection.wdCollapseEnd);
                    LeaveMathZone(target, forward: true);
                }
                else if (OfficeAutomationText.String(location, "before_formula") is { Length: > 0 } before)
                {
                    var formula = ResolveAutomationFormula(document, before);
                    target = WordFormulaHostSemanticReader.CreateRange(document, formula.Host.Range);
                    target.Collapse(WdCollapseDirection.wdCollapseStart);
                    LeaveMathZone(target, forward: false);
                }
                else if (OfficeAutomationText.Number(location, "paragraph_end") is { } number)
                {
                    Paragraphs? paragraphs = null;
                    Paragraph? paragraph = null;
                    try
                    {
                        paragraphs = document.Paragraphs;
                        var index = (int)number;
                        if (index < 1 || index > paragraphs.Count)
                            throw new OfficeAutomationException("invalid_arguments", $"The document has {paragraphs.Count} paragraphs.");
                        paragraph = paragraphs[index];
                        target = paragraph.Range.Duplicate;
                        target.MoveEnd(WdUnits.wdCharacter, -1);
                        target.Collapse(WdCollapseDirection.wdCollapseEnd);
                    }
                    finally { Release(paragraph); Release(paragraphs); }
                }
                else where = OfficeAutomationText.String(location, "kind");
            }
            if (target is null && where == "document_end")
            {
                target = document.Content;
                target.MoveEnd(WdUnits.wdCharacter, -1);
                target.Collapse(WdCollapseDirection.wdCollapseEnd);
            }
            else if (target is null && where is not (null or "" or "cursor"))
            {
                throw new OfficeAutomationException("invalid_arguments", $"Unknown location {where}.");
            }
            target?.Select();
            return ReadCreateSelection();
        }
        finally
        {
            Release(target);
            Release(document);
        }
    }

    /// A caret at a native equation's edge is still inside its math zone, where a new
    /// object would become part of that equation: step out of the zone.
    private static void LeaveMathZone(Range caret, bool forward)
    {
        for (var step = 0; step < 2; step++)
        {
            OMaths? maths = null;
            try
            {
                maths = caret.OMaths;
                if (maths.Count == 0) return;
            }
            finally { Release(maths); }
            caret.Move(WdUnits.wdCharacter, forward ? 1 : -1);
        }
    }

    /// The formula at the user's selection (after a conversion of the selection).
    internal object? AutomationDescribeSelected()
    {
        Document? document = null;
        Selection? selection = null;
        Range? range = null;
        try
        {
            document = _application.ActiveDocument;
            selection = _application.Selection;
            range = selection.Range;
            var found = TryResolveAutomationFormula(document, range);
            if (found is null && range.Start == range.End)
            {
                Release(range);
                range = document.Range(Math.Max(0, selection.Start - 1), selection.Start);
                found = TryResolveAutomationFormula(document, range);
            }
            return found is null ? null : DescribeAutomationFormula(document, found, 40, full: true);
        }
        catch { return null; }
        finally
        {
            Release(range);
            Release(selection);
            Release(document);
        }
    }

    /// The formula a write produced: by its VisualTeX id, else the formula just
    /// before the caret (a write leaves the caret after it).
    private AutomationFormula? FindWritten(Document document, OfficeObjectResult? written)
    {
        if (written?.FormulaId is { Length: > 0 } id)
        {
            try { return ResolveAutomationFormula(document, "vtx:" + id); }
            catch (OfficeAutomationException) { }
        }
        Selection? selection = null;
        Range? probe = null;
        try
        {
            selection = _application.Selection;
            var caret = selection.Start;
            probe = document.Range(Math.Max(0, caret - 1), caret);
            return TryResolveAutomationFormula(document, probe);
        }
        finally
        {
            Release(probe);
            Release(selection);
        }
    }

    internal object? AutomationDescribeWritten(OfficeObjectResult? written)
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument;
            var formula = FindWritten(document, written);
            return formula is null ? null : DescribeAutomationFormula(document, formula, 40, full: true);
        }
        catch { return null; }
        finally { Release(document); }
    }

    // Documents

    private string? TryActiveDocumentIdentity()
    {
        Document? active = null;
        try
        {
            if (TryRead(() => (int?)_application.Documents.Count) is not > 0) return null;
            active = _application.ActiveDocument;
            return DocumentIdentity(active);
        }
        catch { return null; }
        finally { Release(active); }
    }

    private static bool HasVisibleWindow(Document document)
    {
        Windows? windows = null;
        Window? window = null;
        try
        {
            windows = document.Windows;
            if (windows.Count == 0) return false;
            window = windows[1];
            return window.Visible;
        }
        catch { return false; }
        finally { Release(window); Release(windows); }
    }

    /// The open document with this key; without a key the active document.
    private Document FindAutomationDocument(string? key)
    {
        if (string.IsNullOrEmpty(key))
        {
            try
            {
                if (_application.Documents.Count > 0) return _application.ActiveDocument;
            }
            catch { }
            throw new OfficeAutomationException("document_not_found", "No document is open.");
        }
        Documents? documents = null;
        try
        {
            documents = _application.Documents;
            foreach (var document in WordComEnumeration.Each<Document>(documents))
            {
                bool match;
                try { match = OfficeAutomationText.DocumentKey(DocumentIdentity(document)) == key; }
                catch { match = false; }
                if (match) return document;
                Release(document);
            }
        }
        finally { Release(documents); }
        throw OfficeAutomationText.DocumentNotFound();
    }

    // Formulas

    private List<AutomationFormula> CaptureAutomationFormulas(Document document, HashSet<string> kinds)
    {
        bool Wanted(string kind) => kinds.Count == 0 || kinds.Contains(kind);
        var formulas = new List<AutomationFormula>();
        if (Wanted(AutomationVisualTeXKind) || Wanted(AutomationOmmlKind))
        {
            var index = WordFormulaHostResolver.CaptureDocumentIndex(document);
            if (Wanted(AutomationVisualTeXKind))
                formulas.AddRange(index.VisualTeX.Select(host => new AutomationFormula { Host = host, Kind = AutomationVisualTeXKind }));
            if (Wanted(AutomationOmmlKind))
                formulas.AddRange(index.Omml.Select(host => new AutomationFormula { Host = host, Kind = AutomationOmmlKind }));
        }
        if (Wanted(AutomationMathTypeKind))
            formulas.AddRange(CaptureAutomationMathType(document)
                .Select(host => new AutomationFormula { Host = host, Kind = AutomationMathTypeKind }));
        return formulas
            .OrderBy(formula => formula.Host.Range.StoryType == WdStoryType.wdMainTextStory ? 0 : 1)
            .ThenBy(formula => (int)formula.Host.Range.StoryType)
            .ThenBy(formula => formula.Host.Range.Start)
            .ToList();
    }

    /// MathType objects by position only; their MTEF is decoded for the page that is returned.
    private static List<WordFormulaHostDescriptor> CaptureAutomationMathType(Document document)
    {
        var result = new List<WordFormulaHostDescriptor>();
        Range? content = null;
        InlineShapes? shapes = null;
        try
        {
            content = document.Content;
            shapes = content.InlineShapes;
            foreach (var shape in WordComEnumeration.Each<InlineShape>(shapes))
            {
                Range? range = null;
                try
                {
                    if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                    range = shape.Range;
                    result.Add(new WordFormulaHostDescriptor
                    {
                        Kind = WordFormulaHostKind.MathType,
                        Range = new WordFormulaRangeAddress
                        {
                            StoryType = range.StoryType,
                            Start = range.Start,
                            End = range.End,
                        },
                    });
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }
        }
        finally
        {
            Release(shapes);
            Release(content);
        }
        return result;
    }

    /// VisualTeX objects keep their own formula id (`vtx:`); the others are named by
    /// position and content, valid until the document changes.
    private static void AssignAutomationIds(List<AutomationFormula> formulas)
    {
        var durable = formulas
            .Where(formula => formula.Kind == AutomationVisualTeXKind && !string.IsNullOrEmpty(formula.Host.FormulaId))
            .GroupBy(formula => formula.Host.FormulaId!, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() == 1)
            .Select(group => group.Key)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var formula in formulas)
        {
            if (formula.Kind == AutomationVisualTeXKind
                && formula.Host.FormulaId is { } id
                && durable.Contains(id))
                formula.Id = "vtx:" + id;
        }
    }

    private static string KindLetter(string kind) => kind switch
    {
        AutomationVisualTeXKind => "v",
        AutomationOmmlKind => "o",
        _ => "m",
    };

    private static string PositionalId(AutomationFormula formula, string latex) =>
        $"w:{KindLetter(formula.Kind)}:{(int)formula.Host.Range.StoryType}:{formula.Host.Range.Start}:{formula.Host.Range.End}:{OfficeAutomationText.Hash8(latex)}";

    private static string AutomationRevision(Document document, List<AutomationFormula> formulas)
    {
        var end = TryRead(() => (int?)document.Content.End) ?? 0;
        var addresses = string.Join(
            ";",
            formulas.Select(formula => $"{formula.Kind}:{(int)formula.Host.Range.StoryType}:{formula.Host.Range.Start}:{formula.Host.Range.End}"));
        return OfficeAutomationText.Hash8($"{end}|{addresses}");
    }

    private AutomationSemantic ReadAutomationSemantic(Document document, AutomationFormula formula, bool authoritative)
    {
        var host = formula.Host;
        switch (formula.Kind)
        {
            case AutomationVisualTeXKind:
            {
                // listing uses the copy Word keeps beside the object; reading one formula
                // takes the object's own embedded metadata
                var metadata = !authoritative && host.Metadata is { } cached
                    ? cached
                    : WordFormulaHostSemanticReader.Read(document, host).Metadata
                      ?? throw new InvalidDataException("The VisualTeX object has no formula metadata.");
                var lines = metadata.Lines?.Select(line => line.Latex).Where(line => line is not null).ToList()
                    ?? new List<string>();
                return new AutomationSemantic
                {
                    Latex = metadata.Latex,
                    Lines = lines.Count > 0 ? lines : new List<string> { metadata.Latex },
                    Display = host.DisplayMode,
                    Numbered = host.Display && AutomationVisualTeXNumbered(document, host, metadata),
                    FontSizePt = metadata.FontSizePt > 0 ? metadata.FontSizePt : null,
                };
            }
            case AutomationOmmlKind:
                return ReadAutomationOmml(document, host);
            default:
            {
                var semantic = WordMathTypeHostAdapter.ReadSemantic(_application, document, host);
                var metadata = semantic.Metadata;
                return new AutomationSemantic
                {
                    Latex = metadata.Latex,
                    Lines = new List<string> { metadata.Latex },
                    Display = string.Equals(metadata.DisplayMode, "block", StringComparison.OrdinalIgnoreCase) ? "block" : "inline",
                    Numbered = metadata.Numbered,
                    FontSizePt = metadata.FontSizePt > 0 ? metadata.FontSizePt : null,
                };
            }
        }
    }

    /// The number scaffold in the document decides (the copy of the metadata Word keeps can be older).
    private static bool AutomationVisualTeXNumbered(Document document, WordFormulaHostDescriptor host, FormulaMetadata metadata)
    {
        try { return WordFormulaNumberingResolver.ResolveLocal(document, host).Numbered; }
        catch { return host.Numbering.Numbered || metadata.Numbered; }
    }

    private static AutomationSemantic ReadAutomationOmml(Document document, WordFormulaHostDescriptor host)
    {
        Range? range = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? exact = null;
        try
        {
            range = WordFormulaHostSemanticReader.CreateRange(document, host.Range);
            maths = range.OMaths;
            if (maths.Count != 1)
                throw new OfficeAutomationException("stale_formula_id", "The equation moved; call list_formulas again.");
            math = maths[1];
            exact = math.Range.Duplicate;
            var display = math.Type == WdOMathType.wdOMathDisplay;
            var hostXml = WordOmmlConverter.ReadLocalEquationXml(exact);
            var numbered = display && WordOmmlEquationNumbers.HasWordNativeEquationNumberHost(hostXml);
            string? number = null;
            if (numbered)
            {
                try { number = WordOmmlEquationNumbers.ReadWordNativeEquationNumberText(hostXml); }
                catch { }
            }
            var semanticXml = numbered ? WordOmmlEquationNumbers.StripWordNativeEquationNumberHost(hostXml) : hostXml;
            var mathMl = WordOmmlConverter.TransformOmmlToMathMl(semanticXml, display);
            var latex = MathMlToLatexConverter.ConvertWordEquation(mathMl);
            double? size = null;
            try
            {
                var font = exact.Font.Size;
                if (font > 0 && font < 1000) size = font;
            }
            catch { }
            return new AutomationSemantic
            {
                Latex = latex,
                Lines = new List<string> { latex },
                Display = display ? "block" : "inline",
                Numbered = numbered,
                Number = string.IsNullOrWhiteSpace(number) ? null : number,
                FontSizePt = size,
            };
        }
        finally
        {
            Release(exact);
            Release(math);
            Release(maths);
            Release(range);
        }
    }

    private object DescribeAutomationFormula(Document document, AutomationFormula formula, int contextChars, bool full,
        Lazy<List<EquationReferenceTarget>>? numbers = null)
    {
        AutomationSemantic? semantic = null;
        string? error = null;
        try { semantic = ReadAutomationSemantic(document, formula, authoritative: full); }
        catch (OfficeAutomationException) when (full) { throw; }
        catch (Exception failure) { error = failure is OfficeAutomationException ? failure.Message : OfficeErrorText.Describe(failure); }
        var id = formula.Id ?? PositionalId(formula, semantic?.Latex ?? string.Empty);
        var (before, after) = AutomationContext(document, formula.Host.Range, contextChars);
        return new
        {
            formula_id = id,
            kind = formula.Kind,
            latex = semantic?.Latex,
            lines = full && semantic is not null ? semantic.Lines : null,
            display = semantic?.Display ?? formula.Host.DisplayMode,
            numbered = semantic?.Numbered,
            number = semantic?.Number ?? (semantic?.Numbered == true ? AutomationNumber(formula, numbers ?? AutomationNumberTargets(document)) : null),
            font_size_pt = semantic?.FontSizePt,
            location = AutomationLocation(document, formula.Host.Range.StoryType, formula.Host.Range.Start),
            context_before = before,
            context_after = after,
            error,
        };
    }

    /// The numbers the reference picker offers, read once per answer.
    private Lazy<List<EquationReferenceTarget>> AutomationNumberTargets(Document document) =>
        new(() =>
        {
            try { return GetCanonicalEquationReferenceTargets(document).Concat(MathTypeEquationReferences.GetTargets(document)).ToList(); }
            catch { return new List<EquationReferenceTarget>(); }
        });

    private static string? AutomationNumber(AutomationFormula formula, Lazy<List<EquationReferenceTarget>> numbers)
    {
        var host = formula.Host;
        var target = numbers.Value.FirstOrDefault(item => !string.IsNullOrEmpty(host.FormulaId)
                && string.Equals(item.FormulaId, host.FormulaId, StringComparison.OrdinalIgnoreCase))
            ?? numbers.Value.FirstOrDefault(item => item.Position >= host.Range.Start && item.Position <= host.Range.End);
        var text = target?.NumberText?.Trim().Trim('(', ')', ' ');
        return string.IsNullOrEmpty(text) ? null : text;
    }

    private static object AutomationLocation(Document document, WdStoryType story, int position)
    {
        int? paragraph = null;
        if (story == WdStoryType.wdMainTextStory)
        {
            Range? head = null;
            Paragraphs? paragraphs = null;
            try
            {
                // through the character at the position: a range that ends where a
                // paragraph starts does not count that paragraph
                head = document.Range(0, Math.Min(position + 1, Math.Max(position, document.Content.End)));
                paragraphs = head.Paragraphs;
                paragraph = Math.Max(1, paragraphs.Count);
            }
            catch { }
            finally { Release(paragraphs); Release(head); }
        }
        return new
        {
            story = story == WdStoryType.wdMainTextStory ? "main" : story.ToString(),
            position,
            paragraph,
        };
    }

    private static (string Before, string After) AutomationContext(
        Document document,
        WordFormulaRangeAddress address,
        int characters)
    {
        if (characters <= 0) return (string.Empty, string.Empty);
        var storyEnd = StoryEnd(document, address.StoryType);
        // read a little more: object and field characters are dropped
        var before = AutomationText(document, address.StoryType, address.Start - characters * 2, address.Start);
        var after = AutomationText(document, address.StoryType, address.End, Math.Min(storyEnd, address.End + characters * 2));
        if (before.Length > characters) before = before.Substring(before.Length - characters);
        if (after.Length > characters) after = after.Substring(0, characters);
        return (before, after);
    }

    private static int StoryEnd(Document document, WdStoryType story)
    {
        Range? range = null;
        try
        {
            range = story == WdStoryType.wdMainTextStory ? document.Content : document.StoryRanges[story];
            return range.End;
        }
        catch { return int.MaxValue; }
        finally { Release(range); }
    }

    /// The text of a stretch of a story, other equations shown as "[formula]"
    /// (Word gives an equation's text as linear math, not readable context).
    private static string AutomationText(Document document, WdStoryType story, int start, int end)
    {
        start = Math.Max(0, start);
        if (end <= start) return string.Empty;
        string Piece(int from, int to)
        {
            if (to <= from) return string.Empty;
            Range? piece = null;
            try
            {
                piece = WordFormulaHostSemanticReader.CreateRange(document, new WordFormulaRangeAddress { StoryType = story, Start = from, End = to });
                return OfficeAutomationText.Clean(piece.Text);
            }
            catch { return string.Empty; }
            finally { Release(piece); }
        }
        Range? range = null;
        OMaths? maths = null;
        try
        {
            range = WordFormulaHostSemanticReader.CreateRange(document, new WordFormulaRangeAddress { StoryType = story, Start = start, End = end });
            maths = range.OMaths;
            var text = new System.Text.StringBuilder();
            var cursor = start;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? mathRange = null;
                try
                {
                    math = maths[index];
                    mathRange = math.Range;
                    var from = Math.Max(start, mathRange.Start);
                    var to = Math.Min(end, mathRange.End);
                    if (from > cursor) text.Append(Piece(cursor, from));
                    text.Append("[formula]");
                    cursor = Math.Max(cursor, to);
                }
                finally { Release(mathRange); Release(math); }
            }
            text.Append(Piece(cursor, end));
            return text.ToString();
        }
        catch { return Piece(start, end); }
        finally { Release(maths); Release(range); }
    }

    private AutomationFormula? TryResolveAutomationFormula(Document document, Range selection)
    {
        WordFormulaHostDescriptor? host = null;
        string kind = AutomationVisualTeXKind;
        try { host = WordFormulaHostResolver.ResolveLocal(document, selection, WordFormulaHostKind.VisualTeX); }
        catch { }
        if (host is null)
        {
            kind = AutomationMathTypeKind;
            try { host = WordMathTypeHostAdapter.ResolveLocal(_application, document, selection); }
            catch { }
        }
        if (host is null)
        {
            kind = AutomationOmmlKind;
            try { host = WordFormulaHostLookup.ResolveLocal(document, selection, WordFormulaHostKind.Omml); }
            catch { }
        }
        if (host is null) return null;
        var formula = new AutomationFormula { Host = host, Kind = kind };
        if (kind == AutomationVisualTeXKind && !string.IsNullOrEmpty(host.FormulaId))
            formula.Id = "vtx:" + host.FormulaId;
        return formula;
    }

    private AutomationFormula ResolveAutomationFormula(Document document, string formulaId)
    {
        if (formulaId.StartsWith("vtx:", StringComparison.Ordinal))
        {
            var id = formulaId.Substring(4);
            var matches = WordFormulaHostResolver.CaptureDocumentIndex(document).VisualTeX
                .Where(host => string.Equals(host.FormulaId, id, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (matches.Count == 0)
                throw new OfficeAutomationException("formula_not_found", "No VisualTeX formula in this document has that id.");
            return new AutomationFormula { Host = matches[0], Kind = AutomationVisualTeXKind, Id = formulaId };
        }

        var parts = formulaId.Split(':');
        if (parts.Length != 6 || parts[0] != "w"
            || !int.TryParse(parts[2], out var story)
            || !int.TryParse(parts[3], out var start)
            || !int.TryParse(parts[4], out var end))
            throw new OfficeAutomationException("formula_not_found", "This is not a Word formula id from list_formulas.");
        var kind = parts[1] switch
        {
            "v" => AutomationVisualTeXKind,
            "o" => AutomationOmmlKind,
            "m" => AutomationMathTypeKind,
            _ => throw new OfficeAutomationException("formula_not_found", "This is not a Word formula id from list_formulas."),
        };
        var hash = parts[5];
        var candidates = CaptureAutomationFormulas(document, new HashSet<string> { kind })
            .Where(formula => (int)formula.Host.Range.StoryType == story)
            .ToList();
        string LatexOf(AutomationFormula formula)
        {
            try { return ReadAutomationSemantic(document, formula, authoritative: false).Latex; }
            catch { return string.Empty; }
        }
        var exact = candidates.FirstOrDefault(formula => formula.Host.Range.Start == start && formula.Host.Range.End == end);
        if (exact is not null && OfficeAutomationText.Hash8(LatexOf(exact)) == hash)
            return exact;
        // the document changed: the same formula nearby
        const int window = 5000;
        var moved = candidates
            .Where(formula => Math.Abs(formula.Host.Range.Start - start) <= window)
            .OrderBy(formula => Math.Abs(formula.Host.Range.Start - start))
            .FirstOrDefault(formula => OfficeAutomationText.Hash8(LatexOf(formula)) == hash);
        return moved ?? throw new OfficeAutomationException(
            "stale_formula_id",
            "The document changed and this formula could not be found again; call list_formulas again.");
    }

    private static string Clip(string text, int length) =>
        text.Length <= length ? text : text.Substring(0, length) + "…";

    private static T? TryRead<T>(Func<T> read)
    {
        try { return ReadWordStateWithRetry(read); }
        catch { return default; }
    }
}
