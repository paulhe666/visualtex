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
    public void MacPictureMetadataKeepsItsSizeWhenEditedHere()
    {
        // E=mc^2 inserted by VisualTeX for Mac in PowerPoint at 18 pt (80.3 x 35.45 pt): its render pixels are at the
        // Mac's 14 pt reference size (referenceWidthPt), not at fontSizePt
        const string encoded = "visualtex:v1:deflate:fZJRa9wwEIT_StHzyVgr2ZIMfSihoQcNHA00lNKCztqLxdlysOUmach_76p1L4GQgJ72G-2MBj2wue1wcKxhv8K8uD7hHT-M07D0jm1W-BWnOYyRNWLDVrb1dENIo5WSwF1tJFdSG269AV5aJf0eUEhoaUkKqUeS78ZbnHZjiOnd-cmhd-RI8OP7of0JeRAizqz5_sBC9qixRTh4y1ULiishkDsBgpfaaTCqKo1vX655_LFh7egx-7hE88ndksqH-aZ39xdEaLbvx_ZI0xCz53Zw13jhUneZ7v_GpYWJaFyGPU5IWQ6un3HDJowep6vgU7cjSyML0FrXtQVdGinhv-IThusuZYmsC10Z0ADWCE2qXGNMl-E37iicMEV5KvYzpoTTOeEcIQzUxYmddbkbXOHRhRRYdjtQvtjiv0iEaihUZSRUwlZSgLLPRGsqUoEuqtroUgHIWoIRiqrIJWzj8Qwjpfj2xaUwsqYslJWgnh1qd0Kq3F-F1J2-BxOFLEqKtNz41-F680N-wxLDHR_mRmgrra5AaluppwVvaB7_AA";
        var metadata = FormulaMetadataCodec.Decode(encoded);
        Assert.NotNull(metadata);
        Assert.Equal("E=mc^2", metadata!.Latex);
        Assert.Equal(14.0, metadata.RenderFontSizePt!.Value, 2);
        Assert.Equal(18f, FormulaFontSize.InferOleFontSize(80.3f, 35.45f, metadata), 1);
    }

    [Fact]
    public void TempGuardAcceptsOnlyPngInsideDedicatedDirectory()
    {
        var root = Path.Combine(Path.GetTempPath(), "VisualTeX.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var png = Path.Combine(root, $"{Guid.NewGuid()}.png");
        File.WriteAllBytes(png, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII="));
        var guard = new TempPathGuard(root);
        Assert.Equal(Path.GetFullPath(png), guard.ValidatePng(png));
        var outside = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.png");
        File.WriteAllBytes(outside, File.ReadAllBytes(png));
        Assert.Throws<UnauthorizedAccessException>(() => guard.ValidatePng(outside));
    }
}
