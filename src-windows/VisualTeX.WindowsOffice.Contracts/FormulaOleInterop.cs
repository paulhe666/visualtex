using System;
using System.IO;
using System.Runtime.InteropServices;

namespace VisualTeX.WindowsOffice.Contracts;

public static class FormulaOleInterop
{
    public static void Initialize(
        IVisualTeXFormulaObject formula,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        if (formula is null) throw new ArgumentNullException(nameof(formula));
        ThrowIfFailed(
            formula.InitializeFromFiles(
                FormulaMetadataCodec.SerializeJson(metadata),
                emfPath,
                pngPath),
            "Unable to initialize the VisualTeX native OLE object.");
    }

    public static void Update(
        IVisualTeXFormulaObject formula,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        if (formula is null) throw new ArgumentNullException(nameof(formula));
        ThrowIfFailed(
            formula.UpdateFromFiles(
                FormulaMetadataCodec.SerializeJson(metadata),
                emfPath,
                pngPath),
            "Unable to update the VisualTeX native OLE object.");
    }

    public static FormulaMetadata ReadMetadata(IVisualTeXFormulaObject formula)
    {
        if (formula is null) throw new ArgumentNullException(nameof(formula));
        ThrowIfFailed(
            formula.GetFormulaJson(out var metadataJson),
            "Unable to read metadata from the VisualTeX native OLE object.");
        return FormulaMetadataCodec.DeserializeJson(metadataJson)
            ?? throw new InvalidDataException(
                "The VisualTeX native OLE object contains invalid formula metadata.");
    }

    private static void ThrowIfFailed(int hresult, string message)
    {
        if (hresult >= 0) return;
        var error = Marshal.GetExceptionForHR(hresult);
        throw new COMException(
            error is null ? message : $"{message} {error.Message}",
            hresult);
    }
}
