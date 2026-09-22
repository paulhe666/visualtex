using System;
using System.IO;
using System.Runtime.InteropServices;

namespace VisualTeX.WindowsOffice.Contracts;

public static class FormulaOleInterop
{
    private const uint OleCloseSaveIfDirty = 0;

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

    public static void UpdateMetadata(
        IVisualTeXFormulaObject formula,
        FormulaMetadata metadata)
    {
        if (formula is null) throw new ArgumentNullException(nameof(formula));
        if (metadata is null) throw new ArgumentNullException(nameof(metadata));
        metadata.Validate();
        if (formula is not IVisualTeXFormulaMetadata metadataWriter)
            throw new NotSupportedException(
                "The installed VisualTeX OLE server does not support metadata-only updates.");
        ThrowIfFailed(
            metadataWriter.SetFormulaJson(FormulaMetadataCodec.SerializeJson(metadata)),
            "Unable to update metadata in the VisualTeX native OLE object.");
    }

    public static void CloseAfterSave(IVisualTeXFormulaObject formula)
    {
        if (formula is null) throw new ArgumentNullException(nameof(formula));
        if (formula is not IOleObjectClose oleObject)
            throw new NotSupportedException(
                "The VisualTeX native OLE object does not expose IOleObject.Close.");
        ThrowIfFailed(
            oleObject.Close(OleCloseSaveIfDirty),
            "Unable to close the persisted VisualTeX native OLE object.");
    }

    private static void ThrowIfFailed(int hresult, string message)
    {
        if (hresult >= 0) return;
        var error = Marshal.GetExceptionForHR(hresult);
        throw new COMException(
            error is null ? message : $"{message} {error.Message}",
            hresult);
    }

    [ComImport]
    [Guid("00000112-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleObjectClose
    {
        [PreserveSig] int SetClientSite(IntPtr clientSite);
        [PreserveSig] int GetClientSite(out IntPtr clientSite);
        [PreserveSig] int SetHostNames(
            [MarshalAs(UnmanagedType.LPWStr)] string containerApp,
            [MarshalAs(UnmanagedType.LPWStr)] string containerObject);
        [PreserveSig] int Close(uint saveOption);
    }
}
