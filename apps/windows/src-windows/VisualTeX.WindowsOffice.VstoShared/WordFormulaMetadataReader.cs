using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WordVsto;

internal static class WordFormulaMetadataReader
{
    public static FormulaMetadata? TryRead(InlineShape shape)
    {
        if (shape is null) return null;
        if (IsNativeOle(shape))
            return TryReadNativeOle(shape);

        string? encoded = null;
        try { encoded = shape.AlternativeText; } catch { }
        var metadata = FormulaMetadataCodec.Decode(encoded);
        if (metadata is not null) return metadata;
        try { encoded = shape.Title; } catch { encoded = null; }
        return FormulaMetadataCodec.Decode(encoded);
    }

    public static bool IsNativeOle(InlineShape shape)
    {
        OLEFormat? format = null;
        try
        {
            if (shape.Type is not WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                and not WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                return false;
            format = shape.OLEFormat;
            return string.Equals(
                format.ProgID,
                FormulaOleContract.ProgId,
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(format);
        }
    }

    private static FormulaMetadata? TryReadNativeOle(InlineShape shape)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            format = shape.OLEFormat;
            oleObject = WordOleObjectAccessor.GetRunningObject(format);
            return oleObject is IVisualTeXFormulaObject formula
                ? FormulaOleInterop.ReadMetadata(formula)
                : null;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
