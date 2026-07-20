using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOleBridge;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MetadataAndTempPathTests
{
    [Fact]
    public void FormulaMetadataRoundTripsWithPersistentUuid()
    {
        var formulaId = Guid.NewGuid().ToString();
        var metadata = new FormulaMetadata
        {
            FormulaId = formulaId,
            Title = "Matrix",
            Latex = "a+b",
            CodeFormat = "latex",
            DisplayMode = "block",
            Numbered = true,
            CreatedWithVersion = "1.0.18",
            UpdatedWithVersion = "1.0.18",
            CreatedAt = DateTimeOffset.UtcNow.ToString("O"),
            UpdatedAt = DateTimeOffset.UtcNow.ToString("O"),
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString(), Latex = "a+b" },
            },
        };
        var encoded = MetadataCodec.Encode(metadata);
        var decoded = MetadataCodec.Decode(encoded);
        Assert.NotNull(decoded);
        Assert.Equal(formulaId, decoded!.FormulaId);
        Assert.Equal(metadata.Latex, decoded.Latex);
        Assert.True(decoded.Numbered);
    }

    [Fact]
    public void TempGuardAcceptsPngAndSvgOnlyInsideDedicatedDirectory()
    {
        var root = Path.Combine(Path.GetTempPath(), "VisualTeX.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var png = Path.Combine(root, $"{Guid.NewGuid()}.png");
        File.WriteAllBytes(png, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII="));
        var svg = Path.Combine(root, $"{Guid.NewGuid()}.svg");
        File.WriteAllText(svg, "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\"><path d=\"M0 0L10 10\"/></svg>");
        var guard = new TempPathGuard(root);
        Assert.Equal(Path.GetFullPath(png), guard.ValidatePng(png));
        Assert.Equal(Path.GetFullPath(svg), guard.ValidateSvg(svg));

        var outside = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.svg");
        File.WriteAllText(outside, File.ReadAllText(svg));
        Assert.Throws<UnauthorizedAccessException>(() => guard.ValidateSvg(outside));
        Assert.Throws<InvalidOperationException>(() => guard.ValidateSvg(png));
    }
}
