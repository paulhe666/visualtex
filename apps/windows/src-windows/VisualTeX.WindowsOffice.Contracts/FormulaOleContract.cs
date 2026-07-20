using System.Runtime.InteropServices;

namespace VisualTeX.WindowsOffice.Contracts;

public static class FormulaOleContract
{
    public const int ProtocolVersion = 2;
    public const int StorageSchemaVersion = 1;

    public const string ProgId = "VisualTeX.Formula.1";
    public const string VersionIndependentProgId = "VisualTeX.Formula";
    public const string ClassId = "8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B";
    public const string InterfaceId = "6C672AF0-7321-4D21-B325-868CB34592C2";
    public const string AppId = "3C72FF7F-B04A-4FD0-AA7D-61D110D8B3C1";
    public const string TypeLibraryId = "DF66EC66-3B3A-4675-A7BE-30456A04EB96";

    public const string MetadataStream = "VisualTeX.Formula.json";
    public const string EmfPreviewStream = "VisualTeX.Preview.emf";
    public const string PngPreviewStream = "VisualTeX.Preview.png";

    public const string NativeOleMode = "nativeOle";
    public const string WordOmmlMode = "wordOmml";
    public const string CrossPlatformPictureMode = "crossPlatformPicture";
}

[ComImport]
[Guid(FormulaOleContract.InterfaceId)]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IVisualTeXFormulaObject
{
    [DispId(1)]
    [PreserveSig]
    int InitializeFromFiles(
        [MarshalAs(UnmanagedType.BStr)] string metadataJson,
        [MarshalAs(UnmanagedType.BStr)] string emfPath,
        [MarshalAs(UnmanagedType.BStr)] string pngPath);

    [DispId(2)]
    [PreserveSig]
    int UpdateFromFiles(
        [MarshalAs(UnmanagedType.BStr)] string metadataJson,
        [MarshalAs(UnmanagedType.BStr)] string emfPath,
        [MarshalAs(UnmanagedType.BStr)] string pngPath);

    [DispId(3)]
    [PreserveSig]
    int GetFormulaJson([MarshalAs(UnmanagedType.BStr)] out string metadataJson);
}
