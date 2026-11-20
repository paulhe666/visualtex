using System.Reflection;
using System.Runtime.InteropServices;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class FormulaOleContractTests
{
    [Fact]
    public void PublishedComIdentitiesAndStorageNamesAreStable()
    {
        Assert.Equal(2, FormulaOleContract.ProtocolVersion);
        Assert.Equal(1, FormulaOleContract.StorageSchemaVersion);
        Assert.Equal("VisualTeX.Formula.1", FormulaOleContract.ProgId);
        Assert.Equal("VisualTeX.Formula", FormulaOleContract.VersionIndependentProgId);
        Assert.Equal("8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B", FormulaOleContract.ClassId);
        Assert.Equal("6C672AF0-7321-4D21-B325-868CB34592C2", FormulaOleContract.InterfaceId);
        Assert.Equal("3C72FF7F-B04A-4FD0-AA7D-61D110D8B3C1", FormulaOleContract.AppId);
        Assert.Equal("DF66EC66-3B3A-4675-A7BE-30456A04EB96", FormulaOleContract.TypeLibraryId);
        Assert.Equal("VisualTeX.Formula.json", FormulaOleContract.MetadataStream);
        Assert.Equal("VisualTeX.Preview.emf", FormulaOleContract.EmfPreviewStream);
        Assert.Equal("VisualTeX.Preview.png", FormulaOleContract.PngPreviewStream);
    }

    [Fact]
    public void NativeInitializationInterfaceUsesThePublishedDualAutomationAbi()
    {
        var interfaceType = typeof(IVisualTeXFormulaObject);
        Assert.NotNull(interfaceType.GetCustomAttribute<ComImportAttribute>());
        Assert.Equal(
            FormulaOleContract.InterfaceId,
            interfaceType.GetCustomAttribute<GuidAttribute>()?.Value,
            ignoreCase: true);
        Assert.Equal(
            ComInterfaceType.InterfaceIsDual,
            interfaceType.GetCustomAttribute<InterfaceTypeAttribute>()?.Value);

        var methods = interfaceType.GetMethods();
        Assert.Equal(
            new[] { "InitializeFromFiles", "UpdateFromFiles", "GetFormulaJson" },
            methods.Select(method => method.Name).ToArray());
        Assert.Equal(
            new[] { 1, 2, 3 },
            methods.Select(method => method.GetCustomAttribute<DispIdAttribute>()?.Value ?? -1).ToArray());
    }

    [Fact]
    public void BaselineRoundTripsAndMustFitInsideRenderedHeight()
    {
        var metadata = CreateMetadata();
        metadata.RenderWidthPx = 320;
        metadata.RenderHeightPx = 80;
        metadata.Baseline = 62;
        metadata.FontSizePt = 18;
        metadata.RenderFontSizePt = 14;
        metadata.WordInlineOleWidthPt = 42.5;
        metadata.WordInlineOleHeightPt = 10.25;
        metadata.WordInlineOlePositionPt = -2;

        var decoded = FormulaMetadataCodec.Decode(FormulaMetadataCodec.Encode(metadata));
        Assert.NotNull(decoded);
        Assert.Equal(62, decoded!.Baseline);
        Assert.Equal(18, decoded.FontSizePt);
        Assert.Equal(14, decoded.RenderFontSizePt);
        Assert.Equal(42.5, decoded.WordInlineOleWidthPt);
        Assert.Equal(10.25, decoded.WordInlineOleHeightPt);
        Assert.Equal(-2, decoded.WordInlineOlePositionPt);

        var jsonDecoded = FormulaMetadataCodec.DeserializeJson(
            FormulaMetadataCodec.SerializeJson(metadata));
        Assert.NotNull(jsonDecoded);
        Assert.Equal(metadata.FormulaId, jsonDecoded!.FormulaId);
        Assert.Equal(62, jsonDecoded.Baseline);
        Assert.Equal(18, jsonDecoded.FontSizePt);
        Assert.Equal(14, jsonDecoded.RenderFontSizePt);
        Assert.Equal(42.5, jsonDecoded.WordInlineOleWidthPt);
        Assert.Equal(10.25, jsonDecoded.WordInlineOleHeightPt);
        Assert.Equal(-2, jsonDecoded.WordInlineOlePositionPt);

        metadata.Baseline = 81;
        Assert.Throws<InvalidOperationException>(() => metadata.Validate());

        metadata.Baseline = 62;
        metadata.WordInlineOleHeightPt = null;
        Assert.Throws<InvalidOperationException>(() => metadata.Validate());
    }

    private static FormulaMetadata CreateMetadata()
    {
        return new FormulaMetadata
        {
            FormulaId = Guid.NewGuid().ToString(),
            Title = "Formula",
            Latex = "a=b",
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString(), Latex = "a=b" },
            },
            CodeFormat = "raw",
            DisplayMode = "inline",
            CreatedWithVersion = "1.1.0",
            UpdatedWithVersion = "1.1.0",
            CreatedAt = DateTimeOffset.UtcNow.ToString("O"),
            UpdatedAt = DateTimeOffset.UtcNow.ToString("O"),
        };
    }
}
