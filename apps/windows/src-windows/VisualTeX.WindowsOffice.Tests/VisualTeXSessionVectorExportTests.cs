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
    public void ToMetadataRemovesTypingAnchorWithoutDroppingFollowingDigit()
    {
        var line = new FormulaLine
        {
            Id = Guid.NewGuid().ToString("D"),
            Latex = "a^2+b^2=c^2\u200C1",
        };
        var session = new OfficeSessionDocument
        {
            FormulaId = Guid.NewGuid().ToString("D"),
            Title = "Word Formula",
            DisplayMode = "inline",
            Lines = new List<FormulaLine> { line },
        };

        var metadata = session.ToMetadata();

        Assert.Equal("a^2+b^2=c^21", metadata.Latex);
        Assert.Equal("a^2+b^2=c^21", metadata.Lines.Single().Latex);
        Assert.Equal("a^2+b^2=c^2\u200C1", line.Latex);
    }

    [Fact]
    public void ToMetadataCollapsesFormulaDuplicatedAcrossTypingAnchors()
    {
        const string latex = @"\mathrm{e}^{\mathrm{i}\pi}+1=0";
        var session = new OfficeSessionDocument
        {
            FormulaId = Guid.NewGuid().ToString("D"),
            Title = "Word Formula",
            DisplayMode = "inline",
            Lines = new List<FormulaLine>
            {
                new()
                {
                    Id = Guid.NewGuid().ToString("D"),
                    Latex = latex + "\u200C\u200C" + latex,
                },
            },
        };

        var metadata = session.ToMetadata();

        Assert.Equal(latex, metadata.Latex);
        Assert.Equal(latex, metadata.Lines.Single().Latex);
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
