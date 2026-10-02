using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

// Retain Word's actual native COM owner across a forward batch. A numeric
// address becomes stale and a generic boundary Range can absorb an adjacent
// insertion. OMath.Range/InlineShape.Range follow the specific source object.
internal sealed class WordNativeConversionSource : IDisposable
{
    private readonly WordFormulaHostDescriptor _host;
    private OMath? _math;
    private InlineShape? _shape;
    private Range? _number;
    private Range? _container;
    private readonly bool _containerIsHost;
    private WordVisualTeXParagraphNumbering.BoundBodyNumber? _visualTeXBody;

    // Exact captured location, one authoritative payload read, and one local
    // numbering proof. The same physical owner is consumed by the mutation.
    internal static WordNativeConversionSource BindVisualTeX(Document document, Range capturedExact)
    {
        InlineShapes? shapes = null;
        InlineShape? selected = null;
        Range? selectedRange = null;
        WordVisualTeXParagraphNumbering.BoundBodyNumber? body = null;
        try
        {
            shapes = capturedExact.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? range = null;
                try
                {
                    candidate = shapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate)) continue;
                    range = candidate.Range;
                    if (range.StoryType != capturedExact.StoryType
                        || range.Start != capturedExact.Start || range.End != capturedExact.End)
                        continue;
                    if (selected is not null)
                        throw new InvalidDataException("The captured VisualTeX range contains multiple exact owners.");
                    selected = candidate; candidate = null;
                    selectedRange = range; range = null;
                }
                finally { Release(range); Release(candidate); }
            }
            if (selected is null || selectedRange is null)
                throw new InvalidDataException("The captured VisualTeX source no longer has its exact OLE owner.");
            var metadata = WordFormulaMetadataReader.TryReadEmbeddedNativeOle(selected)
                ?? throw new InvalidDataException("The VisualTeX source has no authoritative embedded metadata.");
            var displayMode = string.Equals(metadata.DisplayMode, "block", StringComparison.OrdinalIgnoreCase)
                ? "block" : "inline";
            var host = new WordFormulaHostDescriptor
            {
                Kind = WordFormulaHostKind.VisualTeX,
                Range = new WordFormulaRangeAddress
                {
                    StoryType = selectedRange.StoryType,
                    Start = selectedRange.Start,
                    End = selectedRange.End,
                },
                DisplayMode = displayMode,
                FormulaId = metadata.FormulaId,
                Metadata = metadata,
                MetadataAuthoritative = true,
                WithinTable = Convert.ToBoolean(selectedRange.get_Information(WdInformation.wdWithInTable)),
            };
            host.Numbering = WordFormulaNumberingResolver.ResolveLocal(document, host, out body);
            if (host.Numbering.SelfContainedVisualTeXParagraph && body is null)
                throw new InvalidDataException("The canonical VisualTeX body number has no prepared owner.");

            // Constructor takes ownership even if acquiring its live address ranges fails.
            var ownedShape = selected; selected = null;
            var ownedBody = body; body = null;
            return new WordNativeConversionSource(document, host, ownedShape, null, ownedBody);
        }
        finally
        {
            body?.Dispose();
            Release(selectedRange); Release(selected); Release(shapes);
        }
    }

    internal WordNativeConversionSource(Document document, WordFormulaHostDescriptor host)
        : this(document, host, null, null)
    {
    }

    // Ownership of shape transfers to this binding, including constructor failure.
    internal static WordNativeConversionSource AdoptMathType(Document document,
        WordFormulaHostDescriptor host, InlineShape shape, string mathMl) =>
        new(document, host, shape, mathMl);

    private WordNativeConversionSource(Document document, WordFormulaHostDescriptor host,
        InlineShape? ownedShape, string? mathMl,
        WordVisualTeXParagraphNumbering.BoundBodyNumber? visualTeXBody = null)
    {
        _host = host;
        _shape = ownedShape;
        _visualTeXBody = visualTeXBody;
        MathMl = mathMl;
        Range? exact = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        try
        {
            if (_shape is null)
                exact = WordFormulaHostSemanticReader.CreateRange(document, host.Range);
            if (host.Kind == WordFormulaHostKind.Omml)
            {
                maths = exact!.OMaths;
                _math = maths[1];
            }
            else if (_shape is null)
            {
                shapes = exact!.InlineShapes;
                _shape = shapes[1];
            }
            if (host.Numbering.NumberRange is not null)
                _number = WordFormulaHostSemanticReader.CreateRange(document, host.Numbering.NumberRange);
            if (host.Numbering.ContainerRange is not null)
            {
                var address = host.Numbering.ContainerRange;
                _containerIsHost = address.Start == host.Range.Start && address.End == host.Range.End && address.StoryType == host.Range.StoryType;
                if (!_containerIsHost)
                    _container = WordFormulaHostSemanticReader.CreateRange(document, address);
            }
        }
        catch { Dispose(); throw; }
        finally { Release(shapes); Release(maths); Release(exact); }
    }

    internal WordFormulaHostDescriptor Host => _host;
    internal string? MathMl { get; }
    internal InlineShape? MathTypeShape => _host.Kind == WordFormulaHostKind.MathType ? _shape : null;
    internal InlineShape? VisualTeXShape => _host.Kind == WordFormulaHostKind.VisualTeX ? _shape : null;
    internal WordVisualTeXParagraphNumbering.BoundBodyNumber? VisualTeXBodyNumbering => _visualTeXBody;

    internal void RefreshAddresses()
    {
        Range? exact = null;
        try
        {
            exact = _math is not null ? _math.Range : _shape!.Range;
            Set(_host.Range, exact);
            if (_host.Numbering.ContainerRange is not null)
                Set(_host.Numbering.ContainerRange, _containerIsHost ? exact : _container!);
            if (_host.Numbering.NumberRange is not null)
                Set(_host.Numbering.NumberRange, _number!);
        }
        finally { Release(exact); }
    }

    private static void Set(WordFormulaRangeAddress address, Range range)
    {
        address.StoryType = range.StoryType;
        address.Start = range.Start;
        address.End = range.End;
    }

    public void Dispose()
    {
        var body = _visualTeXBody; _visualTeXBody = null;
        var container = _container; _container = null;
        var number = _number; _number = null;
        var math = _math; _math = null;
        var shape = _shape; _shape = null;
        try { body?.Dispose(); }
        finally
        {
            Release(container); Release(number); Release(math); Release(shape);
        }
    }

    private static void Release(object? value)
    {
        if (value is not null && System.Runtime.InteropServices.Marshal.IsComObject(value))
            System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
    }
}
