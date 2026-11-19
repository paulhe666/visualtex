using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class FormulaMetadataTests
{
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
    public void NumberingIsRejectedForInlineMetadata()
    {
        var metadata = new FormulaMetadata
        {
            FormulaId = Guid.NewGuid().ToString(),
            DisplayMode = "inline",
            Numbered = true,
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString(), Latex = "x" },
            },
        };

        Assert.Throws<InvalidOperationException>(metadata.Validate);
    }
}
