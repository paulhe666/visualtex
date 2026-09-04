using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WordVsto;

/// <summary>
/// Uses VisualTeX's own DPI-aware native OLE LocalServer purely as a presentation
/// converter. The returned CF_METAFILEPICT is then attached to a MathType CFB;
/// no VisualTeX OLE object is inserted into the Word document.
/// </summary>
internal static class VisualTeXOlePresentationProvider
{
    private const short CfMetafilePict = 3;
    private const string TinyPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII=";

    internal static STGMEDIUM CreateMetafilePicture(string emfPath)
    {
        if (string.IsNullOrWhiteSpace(emfPath) || !File.Exists(emfPath))
            throw new FileNotFoundException(
                "VisualTeX vector preview is unavailable for MathType presentation conversion.",
                emfPath);

        var officeTemp = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(officeTemp);
        var stagedEmf = Path.Combine(
            officeTemp,
            $"mathtype-presentation-{Guid.NewGuid():N}.emf");
        var stagedPng = Path.Combine(
            officeTemp,
            $"mathtype-presentation-{Guid.NewGuid():N}.png");

        object? instance = null;
        try
        {
            File.Copy(emfPath, stagedEmf, overwrite: true);
            File.WriteAllBytes(stagedPng, Convert.FromBase64String(TinyPngBase64));

            var type = Type.GetTypeFromProgID(
                FormulaOleContract.ProgId,
                throwOnError: true)
                ?? throw new InvalidOperationException(
                    "VisualTeX native Formula OLE server is not registered.");
            instance = Activator.CreateInstance(type)
                ?? throw new InvalidOperationException(
                    "VisualTeX native Formula OLE server could not be started.");
            if (instance is not IVisualTeXFormulaObject formula)
                throw new InvalidOperationException(
                    "VisualTeX native Formula OLE server does not expose IVisualTeXFormulaObject.");
            if (instance is not System.Runtime.InteropServices.ComTypes.IDataObject dataObject)
                throw new InvalidOperationException(
                    "VisualTeX native Formula OLE server does not expose IDataObject.");

            var now = DateTimeOffset.UtcNow.ToString("O");
            var metadata = new FormulaMetadata
            {
                FormulaId = Guid.NewGuid().ToString("D"),
                Title = "MathType presentation bridge",
                Latex = "x",
                Lines = new List<FormulaLine>
                {
                    new() { Id = Guid.NewGuid().ToString("D"), Latex = "x" },
                },
                CodeFormat = "raw",
                DisplayMode = "inline",
                Numbered = false,
                FontSizePt = FormulaFontSize.DefaultPt,
                RenderFontSizePt = FormulaFontSize.DefaultPt,
                CreatedWithVersion = "1.2.6",
                UpdatedWithVersion = "1.2.6",
                CreatedAt = now,
                UpdatedAt = now,
            };
            metadata.Validate();
            FormulaOleInterop.Initialize(formula, metadata, stagedEmf, stagedPng);

            var request = new FORMATETC
            {
                cfFormat = CfMetafilePict,
                dwAspect = DVASPECT.DVASPECT_CONTENT,
                lindex = -1,
                ptd = IntPtr.Zero,
                tymed = TYMED.TYMED_MFPICT,
            };
            dataObject.GetData(ref request, out var medium);
            if (medium.tymed != TYMED.TYMED_MFPICT || medium.unionmember == IntPtr.Zero)
            {
                if (medium.unionmember != IntPtr.Zero)
                    ReleaseStgMedium(ref medium);
                throw new InvalidDataException(
                    $"VisualTeX OLE presentation converter returned {medium.tymed} instead of TYMED_MFPICT.");
            }
            return medium;
        }
        finally
        {
            if (instance is not null && Marshal.IsComObject(instance))
            {
                try { Marshal.FinalReleaseComObject(instance); }
                catch { }
            }
            try { File.Delete(stagedEmf); } catch { }
            try { File.Delete(stagedPng); } catch { }
        }
    }

    [DllImport("ole32.dll")]
    private static extern void ReleaseStgMedium(ref STGMEDIUM medium);
}
