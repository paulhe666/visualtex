using System.Drawing;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class RibbonIconDataTests
{
    public static IEnumerable<object[]> Icons()
    {
        yield return new object[] { "OLE display", RibbonIconData.OleDisplay };
        yield return new object[] { "OMML display", RibbonIconData.OmmlDisplay };
        yield return new object[] { "OLE inline", RibbonIconData.OleInline };
        yield return new object[] { "OMML inline", RibbonIconData.OmmlInline };
        yield return new object[] { "insert formula", RibbonIconData.InsertFormula };
        yield return new object[] { "update numbers", RibbonIconData.UpdateNumbers };
        yield return new object[] { "edit selected", RibbonIconData.EditSelected };
        yield return new object[] { "convert to OMML", RibbonIconData.ConvertToOmml };
        yield return new object[] { "convert to OLE", RibbonIconData.ConvertToOle };
    }

    [Theory]
    [MemberData(nameof(Icons))]
    public void EmbeddedRibbonIconIsAValidTransparentPng(string name, string encoded)
    {
        var bytes = Convert.FromBase64String(encoded);
        Assert.True(bytes.Length > 500, $"{name} icon is unexpectedly small.");
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4e, 0x47 }, bytes.Take(4).ToArray());

        using var stream = new MemoryStream(bytes, writable: false);
        using var bitmap = new Bitmap(stream);
        Assert.Equal(32, bitmap.Width);
        Assert.Equal(32, bitmap.Height);

        var transparent = 0;
        var visible = 0;
        for (var y = 0; y < bitmap.Height; y++)
        for (var x = 0; x < bitmap.Width; x++)
        {
            var alpha = bitmap.GetPixel(x, y).A;
            if (alpha == 0) transparent++;
            if (alpha > 16) visible++;
        }

        Assert.True(transparent > 0, $"{name} icon lost its transparent background.");
        Assert.True(visible > 40, $"{name} icon contains too little visible artwork.");
    }
}
