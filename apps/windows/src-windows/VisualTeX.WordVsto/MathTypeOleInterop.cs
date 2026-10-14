using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WordVsto;

internal static partial class MathTypeOleInterop
{
    internal static bool IsMathTypeOle(InlineShape shape)
    {
        OLEFormat? format = null;
        try
        {
            if (shape.Type is not WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                and not WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                return false;

            // Fast path when Word can expose the stored ProgID directly.
            try
            {
                format = shape.OLEFormat;
                var progId = format.ProgID;
                if (string.Equals(progId, CanonicalProgId, StringComparison.OrdinalIgnoreCase))
                    return true;
                if (TryResolveCapabilities(progId, out _))
                    return true;
            }
            catch
            {
                // A missing MathType registration can make OLEFormat access fail.
                // Fall through to the serialized WordOpenXML/CFB identity below.
            }
            finally
            {
                Release(format);
                format = null;
            }

            try
            {
                var fragment = MathTypeWordOpenXml.Read(shape);
                // Never identify an object merely because its ProgID starts with
                // "Equation.". Microsoft Equation Editor and unrelated equation
                // servers use that prefix too. The Flat OPC reader validates the
                // embedded CFB as a MathType MTEF-v5 object, including the stored
                // user type/class identity, which also works when MathType is not
                // installed on the current computer.
                return MathTypeOleStorage.LooksLikeMathTypeCompoundFile(
                    fragment.CompoundFile);
            }
            catch
            {
                return false;
            }
        }
        catch
        {
            return false;
        }
        finally { Release(format); }
    }

    internal static bool TryResolveCapabilities(
        InlineShape? shape,
        out Capabilities capabilities)
    {
        capabilities = new Capabilities();
        if (shape is null) return false;
        OLEFormat? format = null;
        try
        {
            if (shape.Type is not WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                and not WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                return false;

            try
            {
                format = shape.OLEFormat;
                if (TryResolveCapabilities(format.ProgID, out capabilities))
                    return true;
            }
            catch
            {
                // A damaged/missing ProgID registration can make OLEFormat fail.
                // Fall through to the serialized storage identity below.
            }
            finally
            {
                Release(format);
                format = null;
            }

            try
            {
                var fragment = MathTypeWordOpenXml.Read(shape);
                var identity = MathTypeOleStorage.ReadCompoundFileIdentity(
                    fragment.CompoundFile);
                return identity.Clsid != Guid.Empty
                    && TryResolveCapabilities(
                        CanonicalProgId,
                        identity.Clsid,
                        out capabilities);
            }
            catch
            {
                return false;
            }
        }
        catch
        {
            capabilities = new Capabilities();
            return false;
        }
        finally { Release(format); }
    }

    internal static bool IsInlineEquation(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            return string.Equals(InferDisplayMode(range, out _), "inline", StringComparison.Ordinal);
        }
        finally { Release(range); }
    }

    internal static FormulaMetadata ReadMetadata(
        Microsoft.Office.Interop.Word.Application application,
        InlineShape shape,
        string? knownMathMl = null,
        byte[]? knownCompoundFile = null)
    {
        var native = MathTypeOleStorage.ReadEquationNative(
            knownCompoundFile ?? MathTypeOleStorage.CaptureCompoundFile(shape));
        var fontSizePt = MathTypeMtefCodec.ReadEquationNativeFullFontSize(native);
        string mathMl;
        try
        {
            // Batch format conversion captures the CFB once and passes the parsed
            // MathML here. Avoid reading Equation Native a second time for every
            // MathType object in a large document.
            mathMl = !string.IsNullOrWhiteSpace(knownMathMl)
                ? knownMathMl!
                : MathTypeMtefCodec.ReadEquationNativeMathMl(native);
        }
        catch (Exception directError)
        {
            // Do not fall back to MathType's IDataObject MathML here. Genuine
            // MathType equations can carry explicit font/character styles in
            // Equation Native (Style > Other... Bold/Italic/etc.) that MathType's
            // exported MathML omits entirely. Treating that lossy export as an
            // authoritative fallback made VisualTeX occasionally reopen a styled
            // native formula as upright/unstyled. The direct MTEF reader is the
            // only style-authoritative read path; unsupported native structures
            // must be implemented there rather than silently normalized away.
            throw new InvalidDataException(
                "VisualTeX could not parse this MathType Equation Native structure without losing native formatting.",
                directError);
        }
        var latex = MathMlToLatexConverter.Convert(mathMl).Trim();
        if (string.IsNullOrWhiteSpace(latex))
            throw new InvalidDataException("MathType OLE returned MathML that VisualTeX could not convert to LaTeX.");

        Range? range = null;
        var displayMode = "inline";
        var numbered = false;
        try
        {
            range = shape.Range;
            using var layoutMetric = WordOperationMetrics.Measure("FormatConversion.MathTypeMetadata.CaptureLayout");
            displayMode = InferDisplayMode(range, out numbered);
        }
        finally
        {
            Release(range);
        }

        var now = DateTimeOffset.UtcNow.ToString("O");
        var formulaId =
            WordFormulaIdentityStore.TryResolveMathType(
                shape,
                out var durableFormulaId)
                ? durableFormulaId
                : Guid.NewGuid().ToString("D");
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
            DisplayMode = displayMode,
            Numbered = numbered,
            FontSizePt = fontSizePt,
            RenderFontSizePt = fontSizePt,
            CreatedWithVersion = "1.2.7",
            UpdatedWithVersion = "1.2.7",
            CreatedAt = now,
            UpdatedAt = now,
        };
        metadata.Validate();
        return metadata;
    }

    /// <summary>
    /// Reads MathType's server-exported MathML for diagnostics and interop
    /// acceptance only. This representation is not presentation-authoritative:
    /// MathType omits some explicit Style &gt; Other character formatting from it.
    /// Production MathType-to-VisualTeX reads must use Equation Native through
    /// <see cref="MathTypeOleStorage.ReadMathMl(InlineShape)"/> instead.
    /// </summary>
    internal static string ReadServerExportedMathMlForDiagnostics(
        InlineShape shape,
        bool activateForConversion = true)
    {
        OLEFormat? format = null;
        object? runningObject = null;
        try
        {
            format = shape.OLEFormat;
            if (!TryResolveCapabilities(format.ProgID, out var capabilities))
                throw new InvalidOperationException("The selected OLE object is not backed by an installed MathType server.");
            if (!capabilities.RegisteredMathMlGetSet)
            {
                // Some MathType 7.x builds expose MathML through IDataObject but
                // omit or vary the optional DataFormats\\GetSet registry entries.
                // Treat the registry as a discovery hint only; QueryGetData below
                // is the authoritative runtime capability probe.
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-mathml-registry-hint-missing progId={format.ProgID} clsid={capabilities.ResolvedClsid}");
            }

            runningObject = activateForConversion
                ? GetRunningMathTypeObject(format, capabilities)
                : TryGetObject(format);
            if (runningObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                throw new InvalidOperationException("MathType OLE does not expose IDataObject.");
            return ReadMathMlFromDataObject(dataObject);
        }
        finally
        {
            CloseOleObject(runningObject);
            Release(runningObject);
            Release(format);
        }
    }

    internal static string DescribeDataFormats(InlineShape shape, bool activateForConversion = true)
    {
        OLEFormat? format = null;
        object? runningObject = null;
        try
        {
            format = shape.OLEFormat;
            if (!TryResolveCapabilities(format.ProgID, out var capabilities))
                return "not-mathtype";
            runningObject = activateForConversion
                ? GetRunningMathTypeObject(format, capabilities)
                : TryGetObject(format);
            if (runningObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                return "no-idataobject";
            var parts = new List<string>();
            AppendFormats(parts, dataObject, DATADIR.DATADIR_GET, "GET");
            AppendFormats(parts, dataObject, DATADIR.DATADIR_SET, "SET");
            foreach (var name in MathMlFormats)
            {
                var id = RegisterClipboardFormat(name);
                if (id == 0 || id > ushort.MaxValue) continue;
                var request = CreateFormatEtc(unchecked((short)id));
                parts.Add($"QUERY {name}({id})=0x{dataObject.QueryGetData(ref request):X8}");
            }
            return string.Join(" | ", parts);
        }
        finally
        {
            CloseOleObject(runningObject);
            Release(runningObject);
            Release(format);
        }
    }

    internal static string ProbeExistingOleMathMlSetVariants(
        InlineShape shape,
        string mathMl)
    {
        OLEFormat? format = null;
        object? runningObject = null;
        try
        {
            format = shape.OLEFormat;
            if (!TryResolveCapabilities(format.ProgID, out var capabilities))
                throw new InvalidOperationException(
                    "The selected OLE object is not backed by an installed MathType server.");
            runningObject = GetRunningMathTypeObject(format, capabilities);
            if (runningObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                throw new InvalidOperationException("MathType OLE does not expose IDataObject.");

            var results = new List<string>();
            foreach (var formatName in MathMlFormats)
            {
                var clipboardId = RegisterClipboardFormat(formatName);
                if (clipboardId == 0 || clipboardId > ushort.MaxValue) continue;
                var request = CreateFormatEtc(unchecked((short)clipboardId));
                results.Add($"{formatName}:Query=0x{dataObject.QueryGetData(ref request):X8}");
                results.Add(formatName + ":Marshal/keep="
                    + ProbeSetDataWithMarshalHGlobal(dataObject, request, mathMl, false));
                results.Add(formatName + ":GlobalMoveable/keep="
                    + ProbeSetDataWithMoveableHGlobal(dataObject, request, mathMl, false));
                results.Add(formatName + ":GlobalMoveable/release="
                    + ProbeSetDataWithMoveableHGlobal(dataObject, request, mathMl, true));
            }
            return string.Join(" | ", results);
        }
        finally
        {
            CloseOleObject(runningObject);
            Release(runningObject);
            Release(format);
        }
    }

    internal static void WriteTextForExistingOle(
        InlineShape shape,
        string text,
        bool activateForConversion = true)
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new InvalidDataException("VisualTeX did not provide formula text for MathType OLE.");

        OLEFormat? format = null;
        object? runningObject = null;
        IntPtr global = IntPtr.Zero;
        try
        {
            format = shape.OLEFormat;
            if (!TryResolveCapabilities(format.ProgID, out var capabilities))
                throw new InvalidOperationException(
                    "The selected OLE object is not backed by an installed MathType server.");
            runningObject = activateForConversion
                ? GetRunningMathTypeObject(format, capabilities)
                : TryGetObject(format);
            if (runningObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                throw new InvalidOperationException("MathType OLE does not expose IDataObject.");

            // MathType 7.x does not expose one stable DATADIR_SET contract across
            // point releases. Some builds advertise CF_TEXT but reject it with
            // DV_E_FORMATETC, while accepting the registered TeX input format that
            // is absent from the SET enumeration. Treat enumeration as a hint and
            // perform the actual SetData calls in semantic-preference order.
            var failures = new List<string>();
            var texFormat = RegisterClipboardFormat("TeX Input Language");
            COMException? texError = null;
            if (texFormat != 0 && texFormat <= ushort.MaxValue)
            {
                if (TrySetAnsiData(
                        dataObject,
                        CreateFormatEtc(unchecked((short)texFormat)),
                        text,
                        out texError))
                    return;
                if (texError is not null)
                    failures.Add($"TeX Input Language=0x{texError.HResult:X8}");
            }

            if (TrySetAnsiData(
                    dataObject,
                    CreateFormatEtc(1),
                    text,
                    out var textError))
                return;
            if (textError is not null)
                failures.Add($"CF_TEXT=0x{textError.HResult:X8}");

            throw new InvalidOperationException(
                "The installed MathType OLE server rejected all supported text SetData formats"
                + (failures.Count == 0 ? "." : ": " + string.Join(", ", failures)));
        }
        finally
        {
            if (global != IntPtr.Zero) Marshal.FreeHGlobal(global);
            CloseOleObject(runningObject);
            Release(runningObject);
            Release(format);
        }
    }

    internal static void WriteMathMl(InlineShape shape, string mathMl, bool activateForConversion = true)
    {
        if (string.IsNullOrWhiteSpace(mathMl)
            || !mathMl.TrimStart().StartsWith("<math", StringComparison.Ordinal))
            throw new InvalidDataException("VisualTeX did not provide valid MathML for MathType OLE.");

        OLEFormat? format = null;
        object? runningObject = null;
        IntPtr global = IntPtr.Zero;
        var ownershipTransferred = false;
        try
        {
            format = shape.OLEFormat;
            if (!TryResolveCapabilities(format.ProgID, out var capabilities))
                throw new InvalidOperationException("The selected OLE object is not backed by an installed MathType server.");
            runningObject = activateForConversion
                ? GetRunningMathTypeObject(format, capabilities)
                : TryGetObject(format);
            if (runningObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                throw new InvalidOperationException("MathType OLE does not expose IDataObject.");
            WriteMathMlToDataObject(dataObject, mathMl);
            ownershipTransferred = true;
        }
        finally
        {
            if (!ownershipTransferred && global != IntPtr.Zero)
                GlobalFree(global);
            CloseOleObject(runningObject);
            Release(runningObject);
            Release(format);
        }
    }

    private static object GetRunningMathTypeObject(OLEFormat format, Capabilities capabilities)
    {
        // MathType's documented IDataObject contract requires the OLE server to
        // be started with its custom RunForConversion verb before Object is
        // queried. Word can expose a dormant OLE proxy through format.Object
        // before that transition; it may answer QueryGetData but reject SetData
        // with DV_E_FORMATETC. Always enter MathType's conversion mode first.
        object verb = capabilities.RunForConversionVerb;
        Exception? activationError = null;
        try { format.DoVerb(ref verb); }
        catch (Exception error) when (error is COMException or InvalidCastException)
        {
            activationError = error;
        }

        var watch = Stopwatch.StartNew();
        var delayMilliseconds = 20;
        while (watch.Elapsed < TimeSpan.FromSeconds(5))
        {
            var running = TryGetObject(format);
            if (running is not null) return running;
            Thread.Sleep(delayMilliseconds);
            delayMilliseconds = Math.Min(120, delayMilliseconds + 15);
        }

        throw new COMException(
            "MathType OLE server started but Word did not expose IDataObject within 5 seconds."
            + (activationError is null ? string.Empty : $" {activationError.Message}"));
    }

    private static object? TryGetObject(OLEFormat format)
    {
        try { return format.Object; }
        catch (Exception error) when (error is COMException or InvalidCastException)
        {
            return null;
        }
    }

    private static string InferDisplayMode(Range range, out bool numbered)
    {
        numbered = false;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? paragraphShapes = null;
        Document? document = null;
        Range? shapePrefix = null;
        ParagraphFormat? paragraphFormat = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            paragraphs = range.Paragraphs;
            if (paragraphs.Count != 1) return "inline";
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            // MathType's right-numbered display equation does not expose the
            // equation number as ordinary paragraph text.  The number is owned by
            // a MACROBUTTON MTPlaceRef field (with nested SEQ MTEqn/MTSec fields),
            // and Word commonly exposes only the field-end control character
            // U+0015 in Paragraph.Range.Text.  If that field shares the paragraph
            // with the MathType OLE, this is unambiguously a display equation.
            numbered = ContainsMathTypeDisplayNumberField(paragraphRange);
            if (numbered) return "block";

            // An unnumbered MathType display equation is a single equation object
            // occupying its paragraph. A paragraph containing multiple inline
            // objects is an inline formula run even when there is no prose between
            // them. Otherwise zero-gap OLE+OLE is misclassified as display math
            // and the conversion safety check refuses the neighboring object.
            paragraphShapes = paragraphRange.InlineShapes;
            if (paragraphShapes.Count != 1) return "inline";

            if (WordEquationNumbering.RangeIsWhollyWithinTable(range))
            {
                // A table cell is itself a paragraph-sized host, so "one OLE and no
                // prose" does not prove display math. VisualTeX/MathType display OLE
                // creation has a stricter structural contract: an exact center TAB
                // immediately before the OLE plus both center/right tab stops. An
                // isolated inline OLE in a cell has no such leading TAB. Preserve
                // that distinction during MathType -> OMML/VisualTeX round trips.
                document = range.Document;
                shapePrefix = document.Range(paragraphRange.Start, range.Start);
                paragraphFormat = paragraph.Format;
                tabStops = paragraphFormat.TabStops;
                var hasCenterTab = false;
                var hasRightTab = false;
                for (var index = 1; index <= tabStops.Count; index++)
                {
                    Release(tabStop);
                    tabStop = tabStops[index];
                    hasCenterTab |= tabStop.Alignment == WdTabAlignment.wdAlignTabCenter;
                    hasRightTab |= tabStop.Alignment == WdTabAlignment.wdAlignTabRight;
                }
                return WordEquationNumbering.IsTableMathTypeDisplayLayout(
                        shapePrefix.Text,
                        hasCenterTab,
                        hasRightTab)
                    ? "block"
                    : "inline";
            }

            var text = (paragraphRange.Text ?? string.Empty)
                .Replace("\r", string.Empty)
                .Replace("\a", string.Empty)
                .Replace("\u0001", string.Empty)
                .Replace("\uFFFC", string.Empty)
                // Word field begin/separator/end controls are structural, not
                // visible prose.  In particular U+0015 is what made genuine
                // MathType MTPlaceRef rows look non-empty and therefore inline.
                .Replace("\u0013", string.Empty)
                .Replace("\u0014", string.Empty)
                .Replace("\u0015", string.Empty)
                .Trim();
            if (text.Length == 0) return "block";

            // Some MathType/Word combinations materialize the display number as
            // visible numeric decoration instead of an empty MTPlaceRef result.
            // Keep supporting those forms without treating ordinary prose as a
            // display equation.
            return LooksLikeDisplayEquationDecoration(text) ? "block" : "inline";
        }
        catch { return "inline"; }
        finally
        {
            Release(tabStop);
            Release(tabStops);
            Release(paragraphFormat);
            Release(shapePrefix);
            Release(document);
            Release(paragraphShapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static bool TryReadDisplayNumberPosition(
        InlineShape shape,
        out string numberPosition)
    {
        numberPosition = "right";
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            fields = WordFormulaHost.GetLocalFields(paragraphRange);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(result);
                result = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                result = field.Result;
                var fieldStart = Math.Max(paragraphRange.Start, code.Start - 1);
                var fieldEnd = Math.Min(paragraphRange.End, result.End + 1);
                if (fieldEnd <= shapeRange.Start)
                    numberPosition = "left";
                else if (fieldStart >= shapeRange.End)
                    numberPosition = "right";
                else
                    numberPosition = code.Start < shapeRange.Start ? "left" : "right";
                return true;
            }
            return false;
        }
        catch
        {
            numberPosition = "right";
            return false;
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static bool ContainsMathTypeDisplayNumberFieldAtRange(Range range)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = range.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return ContainsMathTypeDisplayNumberField(paragraphRange);
        }
        catch { return false; }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool ContainsMathTypeDisplayNumberField(Range paragraphRange)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            fields = WordFormulaHost.GetLocalFields(paragraphRange);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var instruction = code.Text ?? string.Empty;
                if (instruction.IndexOf("MACROBUTTON MTPlaceRef", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static bool LooksLikeDisplayEquationDecoration(string text)
    {
        var sawNumber = false;
        foreach (var character in text)
        {
            if (char.IsWhiteSpace(character)) continue;
            if (char.IsDigit(character))
            {
                sawNumber = true;
                continue;
            }

            switch (character)
            {
                case '(':
                case ')':
                case '[':
                case ']':
                case '{':
                case '}':
                case '.':
                case ',':
                case ':':
                case ';':
                case '-':
                case '–':
                case '—':
                case '/':
                case '\\':
                    continue;
                default:
                    return false;
            }
        }
        return sawNumber;
    }
}
