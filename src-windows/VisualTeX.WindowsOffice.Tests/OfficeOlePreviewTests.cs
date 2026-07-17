using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class OfficeOlePreviewTests
{
    [Fact]
    public void MathJaxStyleSvgProducesAValidatedVectorEmf()
    {
        using var temp = new TemporaryDirectory();
        var svgPath = Path.Combine(temp.Path, "formula.svg");
        File.WriteAllText(svgPath, MathJaxStyleSvg, new UTF8Encoding(false));

        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 420, 130);

        Assert.True(File.Exists(emfPath));
        Assert.True(new FileInfo(emfPath).Length > 88);
        OfficeOlePreview.ValidateVectorEmf(emfPath);

        using var metafile = new Metafile(emfPath);
        using var bitmap = new Bitmap(420, 130, PixelFormat.Format32bppArgb);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            graphics.DrawImage(metafile, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
        }
        var visiblePixels = 0;
        for (var y = 0; y < bitmap.Height; y += 2)
        {
            for (var x = 0; x < bitmap.Width; x += 2)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R < 245 || pixel.G < 245 || pixel.B < 245) visiblePixels++;
            }
        }
        Assert.True(
            visiblePixels > 100,
            $"Vector EMF appears blank: {visiblePixels} visible samples. {OfficeOlePreview.LastRecordingDiagnostics}");
    }

    [Fact]
    public void FullViewBoxRectangleFillsTheEmfPixelFrame()
    {
        using var temp = new TemporaryDirectory();
        var svgPath = Path.Combine(temp.Path, "full-frame.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 100 50\"><rect x=\"0\" y=\"0\" width=\"100\" height=\"50\" fill=\"#111111\"/></svg>",
            new UTF8Encoding(false));

        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 100, 50);
        using var metafile = new Metafile(emfPath);
        using var bitmap = new Bitmap(100, 50, PixelFormat.Format32bppArgb);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            graphics.DrawImage(metafile, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
        }

        var bounds = FindDarkBounds(bitmap);
        Assert.True(
            bounds.Width >= 98 && bounds.Height >= 48,
            $"Full SVG frame recorded as only {bounds.Width}x{bounds.Height} pixels at {bounds.X},{bounds.Y}. {OfficeOlePreview.LastRecordingDiagnostics}");
    }

    [Theory]
    [InlineData("<image href=\"data:image/png;base64,AA==\" width=\"1\" height=\"1\" />")]
    [InlineData("<foreignObject width=\"10\" height=\"10\"></foreignObject>")]
    [InlineData("<script>throw new Error()</script>")]
    [InlineData("<text x=\"0\" y=\"10\">not converted to paths</text>")]
    public void UnsupportedOrRasterSvgFailsClosed(string forbiddenContent)
    {
        using var temp = new TemporaryDirectory();
        var svgPath = Path.Combine(temp.Path, "forbidden.svg");
        File.WriteAllText(
            svgPath,
            $"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\">{forbiddenContent}</svg>",
            new UTF8Encoding(false));

        Assert.Throws<InvalidDataException>(() =>
            OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 20, 20));
        Assert.Empty(Directory.GetFiles(temp.Path, "*.emf"));
    }

    [Fact]
    public void ExternalDefinitionReferenceFailsClosed()
    {
        using var temp = new TemporaryDirectory();
        var svgPath = Path.Combine(temp.Path, "external.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 20 20\"><use href=\"https://example.invalid/glyph.svg#x\" /></svg>",
            new UTF8Encoding(false));

        Assert.Throws<InvalidDataException>(() =>
            OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 20, 20));
    }

    private static Rectangle FindDarkBounds(Bitmap bitmap)
    {
        var left = bitmap.Width;
        var top = bitmap.Height;
        var right = -1;
        var bottom = -1;
        for (var y = 0; y < bitmap.Height; y++)
        for (var x = 0; x < bitmap.Width; x++)
        {
            var pixel = bitmap.GetPixel(x, y);
            if (pixel.R + pixel.G + pixel.B >= 660) continue;
            left = Math.Min(left, x);
            top = Math.Min(top, y);
            right = Math.Max(right, x);
            bottom = Math.Max(bottom, y);
        }
        return right < left || bottom < top
            ? Rectangle.Empty
            : Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
    }

    private const string MathJaxStyleSvg = """
        <svg xmlns="http://www.w3.org/2000/svg"
             xmlns:xlink="http://www.w3.org/1999/xlink"
             viewBox="0 0 420 130"
             color="#151515">
          <defs>
            <path id="glyph-x" d="M10 100 Q25 30 45 60 T80 35 L100 100 H80 V90 C70 110 55 110 45 90 S25 70 10 100 Z" />
            <path id="glyph-y" d="M0 0 A24 18 20 0 1 48 10 A24 18 20 1 0 0 0 Z" />
          </defs>
          <rect x="0" y="0" width="420" height="130" fill="transparent" opacity="0.001" stroke="none" />
          <g fill="currentColor" stroke="none" transform="translate(25 5)">
            <use xlink:href="#glyph-x" transform="scale(1.15 1.05)" />
            <use href="#glyph-y" transform="translate(135 60) rotate(-8) scale(1.2)" />
            <path d="M210 40 H370 V48 H210 Z M285 15 V100 H277 V15 Z" />
            <path d="M220 95 C245 65 265 65 285 95 S325 125 355 85" fill="none" stroke="currentColor" stroke-width="5" />
          </g>
        </svg>
        """;

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"visualtex-vector-emf-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, true); } catch { }
        }
    }
}
