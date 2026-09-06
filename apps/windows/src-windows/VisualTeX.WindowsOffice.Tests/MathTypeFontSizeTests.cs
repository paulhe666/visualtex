using System;
using System.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeFontSizeTests
{
    private const string MathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msup><mi>x</mi><mn>2</mn></msup></math>";

    [Theory]
    [InlineData(10.5)]
    [InlineData(13.25)]
    [InlineData(24)]
    public void CreationCarriesTheRequestedNativeFullSize(double size)
    {
        var result = MathTypeMtefCodec.CreateEquationNativeAtFontSize(MathMl, false, size);
        Assert.Equal(size, MathTypeMtefCodec.ReadEquationNativeFullFontSize(result.EquationNative));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(MathMl),
            MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(result.EquationNative)));
    }

    [Fact]
    public void ContentOnlyEditKeepsTheEntireNativePrefix()
    {
        var source = MathTypeMtefCodec.CreateEquationNativeAtFontSize(MathMl, false, 13.25);
        var edited = MathTypeMtefCodec.RewriteEquationNativeAtFontSize(source.EquationNative, MathMl.Replace("<mi>x</mi>", "<mi>y</mi>"), false, 13.25);
        Assert.Equal(source.Mtef.Take(source.StructureOffset), edited.Mtef.Take(edited.StructureOffset));
        Assert.Equal(13.25, MathTypeMtefCodec.ReadEquationNativeFullFontSize(edited.EquationNative));
    }

    [Fact]
    public void FullSizeChangePreservesTheExactEquationStructure()
    {
        var source = MathTypeMtefCodec.CreateEquationNativeAtFontSize(MathMl, false, 10.5);
        var resized = MathTypeMtefCodec.RewriteEquationNativeAtFontSize(source.EquationNative, MathMl, false, 18.75);
        Assert.Equal(source.Mtef.Skip(source.StructureOffset), resized.Mtef.Skip(resized.StructureOffset));
        Assert.Equal(18.75, MathTypeMtefCodec.ReadEquationNativeFullFontSize(resized.EquationNative));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(double.NaN)]
    public void InvalidSizesAreRejected(double size)
        => Assert.Throws<ArgumentOutOfRangeException>(() =>
            MathTypeMtefCodec.CreateEquationNativeAtFontSize(MathMl, false, size));
}
