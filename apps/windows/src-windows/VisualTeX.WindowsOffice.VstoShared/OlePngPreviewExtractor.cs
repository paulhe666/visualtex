using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace VisualTeX.WindowsOffice.VstoShared;

internal static class OlePngPreviewExtractor
{
    private const int DvAspectContent = 1;
    private const short CfEnhMetafile = 14;
    private const int MaxPngBytes = 64 * 1024 * 1024;
    private const int MaxEmfBytes = 64 * 1024 * 1024;

    internal sealed class PreviewFiles : IDisposable
    {
        internal string EmfPath { get; set; } = string.Empty;
        internal string PngPath { get; set; } = string.Empty;

        public void Dispose()
        {
            DeleteIfPresent(EmfPath);
            DeleteIfPresent(PngPath);
        }

        private static void DeleteIfPresent(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return;
            try { File.Delete(path); } catch { }
        }
    }

    public static PreviewFiles MaterializeForUpdate(
        object oleObject,
        string formulaId)
    {
        var previews = new PreviewFiles();
        try
        {
            previews.PngPath = MaterializePng(oleObject, formulaId);
            previews.EmfPath = MaterializeEmf(oleObject, formulaId);
            return previews;
        }
        catch
        {
            previews.Dispose();
            throw;
        }
    }

    public static string MaterializePng(object oleObject, string formulaId)
    {
        if (oleObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
            throw new InvalidOperationException(
                "The selected OLE object does not expose an IDataObject preview.");

        var clipboardFormat = RegisterClipboardFormat("PNG");
        if (clipboardFormat == 0 || clipboardFormat > ushort.MaxValue)
            throw new InvalidOperationException("The Windows PNG clipboard format is unavailable.");

        var format = new FORMATETC
        {
            cfFormat = unchecked((short)clipboardFormat),
            dwAspect = (DVASPECT)DvAspectContent,
            lindex = -1,
            ptd = IntPtr.Zero,
            tymed = TYMED.TYMED_HGLOBAL,
        };

        dataObject.GetData(ref format, out var medium);
        try
        {
            if (medium.tymed != TYMED.TYMED_HGLOBAL || medium.unionmember == IntPtr.Zero)
                throw new InvalidDataException("The OLE PNG preview is not an HGLOBAL payload.");

            var sizeValue = GlobalSize(medium.unionmember).ToUInt64();
            if (sizeValue < 8 || sizeValue > MaxPngBytes)
                throw new InvalidDataException("The OLE PNG preview size is invalid.");
            var size = checked((int)sizeValue);
            var source = GlobalLock(medium.unionmember);
            if (source == IntPtr.Zero)
                throw new InvalidOperationException("Unable to lock the OLE PNG preview.");
            byte[] bytes;
            try
            {
                bytes = new byte[size];
                Marshal.Copy(source, bytes, 0, size);
            }
            finally
            {
                GlobalUnlock(medium.unionmember);
            }

            ValidatePng(bytes);
            var root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX",
                "office",
                "temp");
            Directory.CreateDirectory(root);
            var safeFormulaId = Guid.TryParse(formulaId, out var id)
                ? id.ToString("N")
                : Guid.NewGuid().ToString("N");
            var path = Path.Combine(root, $"{safeFormulaId}-{Guid.NewGuid():N}-ole-export.png");
            File.WriteAllBytes(path, bytes);
            return path;
        }
        finally
        {
            ReleaseStgMedium(ref medium);
        }
    }

    private static string MaterializeEmf(
        object oleObject,
        string formulaId)
    {
        if (oleObject is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
            throw new InvalidOperationException(
                "The selected OLE object does not expose an IDataObject preview.");

        var format = new FORMATETC
        {
            cfFormat = CfEnhMetafile,
            dwAspect = (DVASPECT)DvAspectContent,
            lindex = -1,
            ptd = IntPtr.Zero,
            tymed = TYMED.TYMED_ENHMF,
        };

        dataObject.GetData(ref format, out var medium);
        try
        {
            if (medium.tymed != TYMED.TYMED_ENHMF
                || medium.unionmember == IntPtr.Zero)
                throw new InvalidDataException(
                    "The OLE enhanced-metafile preview is unavailable.");

            var byteCount =
                GetEnhMetaFileBits(
                    medium.unionmember,
                    0,
                    null);
            if (byteCount == 0
                || byteCount > MaxEmfBytes)
                throw new InvalidDataException(
                    "The OLE enhanced-metafile preview size is invalid.");

            var bytes = new byte[checked((int)byteCount)];
            if (GetEnhMetaFileBits(
                    medium.unionmember,
                    byteCount,
                    bytes) != byteCount)
                throw new InvalidDataException(
                    "The OLE enhanced-metafile preview could not be read completely.");

            var root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX",
                "office",
                "temp");
            Directory.CreateDirectory(root);
            var safeFormulaId = Guid.TryParse(formulaId, out var id)
                ? id.ToString("N")
                : Guid.NewGuid().ToString("N");
            var path = Path.Combine(
                root,
                $"{safeFormulaId}-{Guid.NewGuid():N}-ole-export.emf");
            File.WriteAllBytes(path, bytes);
            return path;
        }
        finally
        {
            ReleaseStgMedium(ref medium);
        }
    }

    private static void ValidatePng(byte[] bytes)
    {
        if (bytes.Length < 8
            || bytes[0] != 137
            || bytes[1] != 80
            || bytes[2] != 78
            || bytes[3] != 71
            || bytes[4] != 13
            || bytes[5] != 10
            || bytes[6] != 26
            || bytes[7] != 10)
            throw new InvalidDataException("The OLE PNG preview signature is invalid.");
    }

    [DllImport("gdi32.dll", SetLastError = true)]
    private static extern uint GetEnhMetaFileBits(
        IntPtr enhancedMetafile,
        uint bufferSize,
        [Out] byte[]? buffer);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint RegisterClipboardFormat(string format);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalLock(IntPtr memory);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GlobalUnlock(IntPtr memory);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern UIntPtr GlobalSize(IntPtr memory);

    [DllImport("ole32.dll")]
    private static extern void ReleaseStgMedium(ref STGMEDIUM medium);
}
