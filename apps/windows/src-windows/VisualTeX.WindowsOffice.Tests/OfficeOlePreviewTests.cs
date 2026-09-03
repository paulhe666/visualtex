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

    [Fact]
    public void MathJaxChineseTextIsConvertedToVectorGlyphOutlines()
    {
        using var temp = new TemporaryDirectory();
        var svgPath = Path.Combine(temp.Path, "mixed-chinese.svg");
        File.WriteAllText(svgPath, MathJaxChineseTextSvg, new UTF8Encoding(false));

        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 80);

        Assert.True(File.Exists(emfPath));
        OfficeOlePreview.ValidateVectorEmf(emfPath);
        using var metafile = new Metafile(emfPath);
        using var bitmap = new Bitmap(240, 80, PixelFormat.Format32bppArgb);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            graphics.DrawImage(metafile, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
        }

        var bounds = FindDarkBounds(bitmap);
        Assert.True(
            bounds.Width >= 120 && bounds.Height >= 35,
            $"Chinese text glyph outlines were not recorded correctly: {bounds}. {OfficeOlePreview.LastRecordingDiagnostics}");
    }

    [Fact]
    public void MathJaxTallMatrixNestedSvgViewportKeepsBracketsContinuous()
    {
        using var temp = new TemporaryDirectory();
        var svgPath = Path.Combine(temp.Path, "tall-matrix.svg");
        File.WriteAllText(svgPath, MathJaxTallMatrixSvg, new UTF8Encoding(false));

        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 308, 270);
        OfficeOlePreview.ValidateVectorEmf(emfPath);
        using var metafile = new Metafile(emfPath);
        using var bitmap = new Bitmap(308, 270, PixelFormat.Format32bppArgb);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            graphics.DrawImage(metafile, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
        }

        var leftCoveredRows = 0;
        var rightCoveredRows = 0;
        for (var y = 42; y <= 228; y++)
        {
            if (RowContainsDarkPixel(bitmap, y, 8, 28)) leftCoveredRows++;
            if (RowContainsDarkPixel(bitmap, y, 280, 302)) rightCoveredRows++;
        }
        Assert.True(
            leftCoveredRows >= 175,
            $"The left tall-matrix bracket is structurally broken: {leftCoveredRows}/187 covered rows. {OfficeOlePreview.LastRecordingDiagnostics}");
        Assert.True(
            rightCoveredRows >= 175,
            $"The right tall-matrix bracket is structurally broken: {rightCoveredRows}/187 covered rows. {OfficeOlePreview.LastRecordingDiagnostics}");
    }

    [Fact]
    public void InkSafePreviewExpandsFinalSystemFontOutlinesWithoutClipping()
    {
        using var temp = new TemporaryDirectory();
        var svgPath = Path.Combine(temp.Path, "tight-system-font.svg");
        var browserPngPath = Path.Combine(temp.Path, "browser-preview.png");
        File.WriteAllBytes(browserPngPath, new byte[] { 1, 2, 3 });
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 -200 1000 300\" color=\"#111111\">"
            + "<text x=\"0\" y=\"0\" font-size=\"1000px\" font-family=\"Times New Roman\" fill=\"currentColor\">Lz</text>"
            + "</svg>",
            new UTF8Encoding(false));

        var preview = OfficeOlePreview.CreateInkSafePreviewFromSvg(
            svgPath,
            widthPixels: 100,
            heightPixels: 30,
            baselinePixels: 20,
            fallbackPngPath: browserPngPath,
            safetyPaddingPixels: 1);

        Assert.NotEqual(browserPngPath, preview.PngPath);
        Assert.True(
            preview.HeightPixels > 30,
            $"Final Times New Roman outlines did not expand the tight SVG frame: {preview.HeightPixels:0.###}px.");
        Assert.InRange(preview.BaselinePixels, 0.001f, preview.HeightPixels - 0.001f);
        OfficeOlePreview.ValidateVectorEmf(preview.EmfPath);

        using (var png = new Bitmap(preview.PngPath))
        {
            var ink = FindVisibleBounds(png);
            Assert.False(ink.IsEmpty);
            Assert.True(
                ink.Left > 0 && ink.Top > 0 && ink.Right < png.Width && ink.Bottom < png.Height,
                $"Ink-safe PNG still touches its clip edge: ink={ink}, bitmap={png.Width}x{png.Height}.");
        }

        using var metafile = new Metafile(preview.EmfPath);
        var emfWidth = Math.Max(1, (int)Math.Ceiling(preview.WidthPixels));
        var emfHeight = Math.Max(1, (int)Math.Ceiling(preview.HeightPixels));
        using var bitmap = new Bitmap(emfWidth, emfHeight, PixelFormat.Format32bppArgb);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            graphics.DrawImage(metafile, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
        }
        var emfInk = FindDarkBounds(bitmap);
        Assert.False(emfInk.IsEmpty);
        Assert.True(
            emfInk.Left > 0
            && emfInk.Top > 0
            && emfInk.Right < bitmap.Width
            && emfInk.Bottom < bitmap.Height,
            $"Ink-safe vector EMF still touches its clip edge: ink={emfInk}, bitmap={bitmap.Width}x{bitmap.Height}. "
            + OfficeOlePreview.LastRecordingDiagnostics);
    }

    [Fact]
    public void PathOnlyPreviewReusesTheAlreadyGeneratedPng()
    {
        using var temp = new TemporaryDirectory();
        var svgPath = Path.Combine(temp.Path, "path-only.svg");
        var pngPath = Path.Combine(temp.Path, "existing.png");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 100 50\"><path d=\"M10 10H90V40H10Z\"/></svg>",
            new UTF8Encoding(false));
        File.WriteAllBytes(pngPath, new byte[] { 1, 2, 3 });

        var preview = OfficeOlePreview.CreateInkSafePreviewFromSvg(
            svgPath,
            widthPixels: 100,
            heightPixels: 50,
            baselinePixels: 35,
            fallbackPngPath: pngPath);

        Assert.Equal(pngPath, preview.PngPath);
        Assert.Equal(100, preview.WidthPixels);
        Assert.Equal(50, preview.HeightPixels);
        Assert.Equal(35, preview.BaselinePixels);
        Assert.Single(Directory.GetFiles(temp.Path, "*.png"));
        OfficeOlePreview.ValidateVectorEmf(preview.EmfPath);
    }

    [Theory]
    [InlineData("<image href=\"data:image/png;base64,AA==\" width=\"1\" height=\"1\" />")]
    [InlineData("<foreignObject width=\"10\" height=\"10\"></foreignObject>")]
    [InlineData("<script>throw new Error()</script>")]
    [InlineData("<text x=\"0\" y=\"10\"><tspan>nested text</tspan></text>")]
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

    private static bool RowContainsDarkPixel(Bitmap bitmap, int y, int left, int right)
    {
        for (var x = left; x <= right; x++)
        {
            var pixel = bitmap.GetPixel(x, y);
            if (pixel.R + pixel.G + pixel.B < 660) return true;
        }
        return false;
    }

    private static Rectangle FindVisibleBounds(Bitmap bitmap)
    {
        var left = bitmap.Width;
        var top = bitmap.Height;
        var right = -1;
        var bottom = -1;
        for (var y = 0; y < bitmap.Height; y++)
        for (var x = 0; x < bitmap.Width; x++)
        {
            if (bitmap.GetPixel(x, y).A <= 8) continue;
            left = Math.Min(left, x);
            top = Math.Min(top, y);
            right = Math.Max(right, x);
            bottom = Math.Max(bottom, y);
        }
        return right < left || bottom < top
            ? Rectangle.Empty
            : Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
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

    private const string MathJaxTallMatrixSvg = """
        <svg xmlns="http://www.w3.org/2000/svg"
             xmlns:xlink="http://www.w3.org/1999/xlink"
             viewBox="-428.5714285714286 -7478.571428571428 16481.14285714286 14457.142857142857"
             color="#111111">
          <defs>
            <path id="left-top" d="M319 -645V1154H666V1070H403V-645H319Z" />
            <path id="left-middle" d="M319 0V602H403V0H319Z" />
            <path id="left-bottom" d="M319 -644V1155H403V-560H666V-644H319Z" />
            <path id="right-top" d="M0 1070V1154H347V-645H263V1070H0Z" />
            <path id="right-middle" d="M263 0V602H347V0H263Z" />
            <path id="right-bottom" d="M263 -560V1155H347V-644H0V-560H263Z" />
          </defs>
          <g stroke="currentColor" fill="currentColor" stroke-width="0" transform="scale(1,-1)">
            <g>
              <use href="#left-top" transform="translate(0,5896)" />
              <use href="#left-bottom" transform="translate(0,-5906)" />
              <svg width="667" height="10202" y="-4851" x="0" viewBox="0 2550.5 667 10202">
                <use href="#left-middle" transform="scale(1,25.42)" />
              </svg>
            </g>
            <g transform="translate(15000,0)">
              <use href="#right-top" transform="translate(0,5896)" />
              <use href="#right-bottom" transform="translate(0,-5906)" />
              <svg width="667" height="10202" y="-4851" x="0" viewBox="0 2550.5 667 10202">
                <use href="#right-middle" transform="scale(1,25.42)" />
              </svg>
            </g>
          </g>
        </svg>
        """;

    private const string MathJaxChineseTextSvg = """
        <svg xmlns="http://www.w3.org/2000/svg"
             viewBox="0 -1000 3000 1200"
             color="#111111">
          <g stroke="currentColor" fill="currentColor" stroke-width="0" transform="scale(1,-1)">
            <path d="M100 0 L500 0 L500 700 L100 700 Z" />
            <g transform="translate(700,0)">
              <text data-variant="normal" transform="scale(1,-1)" font-size="1000px" font-family="serif">的</text>
            </g>
            <g transform="translate(1700,0)">
              <text data-variant="normal" transform="scale(1,-1)" font-size="1000px" font-family="serif">地方</text>
            </g>
          </g>
        </svg>
        """;

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
