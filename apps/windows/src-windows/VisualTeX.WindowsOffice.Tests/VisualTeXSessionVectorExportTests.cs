using System.Text;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class VisualTeXSessionVectorExportTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MaterializeSvgWritesUtf8InsideTheControlledOfficeTempRoot(bool useBase64)
    {
        const string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\"><path d=\"M0 0 L10 10\" /></svg>";
        var sessionId = Guid.NewGuid();
        var export = new OfficeExportDocument();
        if (useBase64)
            export.SvgBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));
        else
            export.Svg = svg;
        var session = new OfficeSessionDocument
        {
            Id = sessionId.ToString("D"),
            ExportResult = export,
        };

        using var client = new VisualTeXSessionClient();
        var path = client.MaterializeSvg(session);
        try
        {
            var expectedRoot = Path.GetFullPath(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX",
                "office",
                "temp")) + Path.DirectorySeparatorChar;
            var fullPath = Path.GetFullPath(path);
            Assert.StartsWith(expectedRoot, fullPath, StringComparison.OrdinalIgnoreCase);
            Assert.Equal($"{sessionId:D}.svg", Path.GetFileName(path));
            var bytes = File.ReadAllBytes(path);
            Assert.NotEmpty(bytes);
            Assert.False(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF);
            Assert.Equal(svg, Encoding.UTF8.GetString(bytes));
        }
        finally
        {
            try { File.Delete(path); } catch { }
        }
    }

    [Fact]
    public void MaterializeSvgRejectsPathLikeSessionIdentifiers()
    {
        using var client = new VisualTeXSessionClient();
        var session = new OfficeSessionDocument
        {
            Id = "..\\..\\escape",
            ExportResult = new OfficeExportDocument
            {
                Svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 1 1\"></svg>",
            },
        };

        Assert.Throws<InvalidOperationException>(() => client.MaterializeSvg(session));
    }

    [Fact]
    public void MaterializeSvgRejectsEmbeddedRasterContent()
    {
        using var client = new VisualTeXSessionClient();
        var session = new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            ExportResult = new OfficeExportDocument
            {
                Svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 1 1\"><image href=\"data:image/png;base64,AA==\" /></svg>",
            },
        };

        Assert.Throws<InvalidDataException>(() => client.MaterializeSvg(session));
    }
}
