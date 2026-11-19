using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

// A batch font-size change records the new size in Word's cached metadata only
// (FormulaFontSize.WithDeferredFontSize, CacheMatchesOleSize).
public sealed class FormulaDeferredFontSizeTests
{
    private static FormulaMetadata Metadata(double size, string updatedAt, string latex = "x+y") => new()
    {
        FormulaId = "0f8fad5b-d9cb-469f-a165-70867728950e",
        Latex = latex,
        Lines = { new FormulaLine { Id = "line-1", Latex = latex } },
        DisplayMode = "block",
        FontSizePt = size,
        RenderFontSizePt = 14,
        RenderWidthPx = 400,
        RenderHeightPx = 100,
        UpdatedWithVersion = "1.2.7",
        UpdatedAt = updatedAt,
    };

    [Fact]
    public void TheLaterCacheOfTheSameFormulaGivesItsSize()
    {
        var embedded = Metadata(10.5, "2026-11-18T10:00:00.0000000+00:00");
        var cached = Metadata(14, "2026-11-18T10:05:00.0000000+00:00");

        var read = FormulaFontSize.WithDeferredFontSize(embedded, cached);

        Assert.Equal(14, read.FontSizePt);
        Assert.Equal(cached.UpdatedAt, read.UpdatedAt);
        Assert.Equal("x+y", read.Latex);
    }

    [Fact]
    public void TheInlineGeometryFollowsTheCachedSize()
    {
        var embedded = Metadata(10.5, "2026-11-18T10:00:00.0000000+00:00");
        embedded.DisplayMode = "inline";
        embedded.WordInlineOleWidthPt = 30;
        embedded.WordInlineOleHeightPt = 9;
        embedded.WordInlineOlePositionPt = -2;
        var cached = Metadata(14, "2026-11-18T10:05:00.0000000+00:00");
        cached.DisplayMode = "inline";
        cached.WordInlineOleWidthPt = 40;
        cached.WordInlineOleHeightPt = 12;
        cached.WordInlineOlePositionPt = -3;

        var read = FormulaFontSize.WithDeferredFontSize(embedded, cached);

        Assert.Equal(14, read.FontSizePt);
        Assert.Equal(40, read.WordInlineOleWidthPt);
        Assert.Equal(12, read.WordInlineOleHeightPt);
        Assert.Equal(-3, read.WordInlineOlePositionPt);
    }

    [Fact]
    public void AnOlderCacheIsIgnored()
    {
        var embedded = Metadata(10.5, "2026-11-18T10:05:00.0000000+00:00");
        var cached = Metadata(14, "2026-11-18T10:00:00.0000000+00:00");

        Assert.Equal(10.5, FormulaFontSize.WithDeferredFontSize(embedded, cached).FontSizePt);
    }

    [Fact]
    public void ACacheWithTheSameTimeIsIgnored()
    {
        var embedded = Metadata(10.5, "2026-11-18T10:05:00.0000000+00:00");
        var cached = Metadata(14, "2026-11-18T10:05:00.0000000+00:00");

        Assert.Equal(10.5, FormulaFontSize.WithDeferredFontSize(embedded, cached).FontSizePt);
    }

    [Fact]
    public void ACacheOfAnotherFormulaIsIgnored()
    {
        var embedded = Metadata(10.5, "2026-11-18T10:00:00.0000000+00:00");
        var cached = Metadata(14, "2026-11-18T10:05:00.0000000+00:00");
        cached.FormulaId = "7c9e6679-7425-40de-944b-e07fc1f90ae7";

        Assert.Equal(10.5, FormulaFontSize.WithDeferredFontSize(embedded, cached).FontSizePt);
    }

    [Fact]
    public void ACacheWithOtherContentIsIgnored()
    {
        var embedded = Metadata(10.5, "2026-11-18T10:00:00.0000000+00:00");
        var cached = Metadata(14, "2026-11-18T10:05:00.0000000+00:00", latex: "a+b");

        var read = FormulaFontSize.WithDeferredFontSize(embedded, cached);

        Assert.Equal(10.5, read.FontSizePt);
        Assert.Equal("x+y", read.Latex);
    }

    [Fact]
    public void NoCacheOrAnUndatedOneLeavesThePayloadAsItIs()
    {
        var embedded = Metadata(10.5, "2026-11-18T10:00:00.0000000+00:00");
        Assert.Equal(10.5, FormulaFontSize.WithDeferredFontSize(embedded, null).FontSizePt);

        var cached = Metadata(14, string.Empty);
        Assert.Equal(10.5, FormulaFontSize.WithDeferredFontSize(embedded, cached).FontSizePt);
    }

    [Fact]
    public void TheCacheDescribesAnObjectOfItsOwnSize()
    {
        var cached = Metadata(14, "2026-11-18T10:00:00.0000000+00:00");
        var size = FormulaFontSize.OleSizeAt(cached, 14);

        Assert.True(FormulaFontSize.CacheMatchesOleSize(cached, size.Width, size.Height));
        // Word rounds the sizes it stores
        Assert.True(FormulaFontSize.CacheMatchesOleSize(cached, size.Width + 0.4f, size.Height - 0.3f));
        // the object at another font size, or stretched
        Assert.False(FormulaFontSize.CacheMatchesOleSize(cached, size.Width * 1.33f, size.Height * 1.33f));
        Assert.False(FormulaFontSize.CacheMatchesOleSize(cached, size.Width * 1.33f, size.Height));
        Assert.False(FormulaFontSize.CacheMatchesOleSize(null, size.Width, size.Height));
    }

    [Fact]
    public void AnInlineCacheIsCheckedAgainstTheSizeWordWasGiven()
    {
        var cached = Metadata(14, "2026-11-18T10:00:00.0000000+00:00");
        cached.DisplayMode = "inline";
        cached.WordInlineOleWidthPt = 52.25;
        cached.WordInlineOleHeightPt = 13.75;

        Assert.True(FormulaFontSize.CacheMatchesOleSize(cached, 52.3f, 13.7f));
        Assert.False(FormulaFontSize.CacheMatchesOleSize(cached, 80f, 21f));
    }

    [Fact]
    public void ACacheWithoutRenderSizeIsNotUsed()
    {
        var cached = Metadata(14, "2026-11-18T10:00:00.0000000+00:00");
        cached.RenderWidthPx = null;

        Assert.False(FormulaFontSize.CacheMatchesOleSize(cached, 300f, 75f));
    }
}
