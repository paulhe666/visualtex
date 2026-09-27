using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

internal static partial class MathTypeOleStorage
{
    internal static ClipboardTransaction BeginClipboardTransaction(InlineShape shape) =>
        new(() =>
        {
            var range = shape.Range;
            try { range.Copy(); }
            finally { Release(range); }
        });

    internal static bool TryCaptureCompoundFileFromWordOpenXml(
        InlineShape shape,
        out byte[] compoundFile)
    {
        if (shape is null) throw new ArgumentNullException(nameof(shape));
        try
        {
            compoundFile = MathTypeWordOpenXml.Read(shape).CompoundFile;
            ValidateCompoundFile(compoundFile);
            return true;
        }
        catch
        {
            compoundFile = Array.Empty<byte>();
            return false;
        }
    }

    internal static byte[] CaptureCompoundFile(InlineShape shape)
    {
        if (shape is null) throw new ArgumentNullException(nameof(shape));
        if (TryCaptureCompoundFileFromWordOpenXml(shape, out var compoundFile))
            return compoundFile;

        // Compatibility fallback for existing/legacy Word containers whose
        // Range.WordOpenXML does not expose the OLE package parts. New VisualTeX
        // MathType creation deliberately does not use this clipboard fallback.
        using var transaction = BeginClipboardTransaction(shape);
        return transaction.CompoundFile;
    }

    internal static string ReadMathMl(InlineShape shape) =>
        ReadMathMl(CaptureCompoundFile(shape));
}
