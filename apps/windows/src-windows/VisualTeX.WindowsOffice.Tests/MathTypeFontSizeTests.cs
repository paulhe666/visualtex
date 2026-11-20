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
        var result = MathTypeMtefFontSize.CreateEquationNativeAtFontSize(MathMl, false, size);
        Assert.Equal(size, MathTypeMtefFontSize.ReadEquationNativeFullFontSize(result.EquationNative));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(MathMl),
            MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(result.EquationNative)));
    }

    [Fact]
    public void ContentOnlyEditKeepsTheEntireNativePrefix()
    {
        var source = MathTypeMtefFontSize.CreateEquationNativeAtFontSize(MathMl, false, 13.25);
        var edited = MathTypeMtefFontSize.RewriteEquationNativeAtFontSize(source.EquationNative, MathMl.Replace("<mi>x</mi>", "<mi>y</mi>"), false, 13.25);
        Assert.Equal(source.Mtef.Take(source.StructureOffset), edited.Mtef.Take(edited.StructureOffset));
        Assert.Equal(13.25, MathTypeMtefFontSize.ReadEquationNativeFullFontSize(edited.EquationNative));
    }

    [Fact]
    public void FullSizeChangePreservesTheExactEquationStructure()
    {
        var source = MathTypeMtefFontSize.CreateEquationNativeAtFontSize(MathMl, false, 10.5);
        var resized = MathTypeMtefFontSize.RewriteEquationNativeAtFontSize(source.EquationNative, MathMl, false, 18.75);
        Assert.Equal(source.Mtef.Skip(source.StructureOffset), resized.Mtef.Skip(resized.StructureOffset));
        Assert.Equal(18.75, MathTypeMtefFontSize.ReadEquationNativeFullFontSize(resized.EquationNative));
    }

    [Fact]
    public void PreviewScalingChangesOnlyTheSizePreferences()
    {
        var source = MathTypeMtefFontSize.CreateEquationNativeAtFontSize(MathMl, false, 32);
        var scaled = MathTypeMtefFontSize.ScaleMtefPreferenceSizes(source.Mtef, 0.25);
        var native = source.EquationNative.Take(28).Concat(scaled).ToArray();
        BitConverter.GetBytes((uint)scaled.Length).CopyTo(native, 8);
        Assert.Equal(8, MathTypeMtefFontSize.ReadEquationNativeFullFontSize(native));
        Assert.Equal(source.Mtef.Skip(source.StructureOffset), scaled.Skip(scaled.Length - (source.Mtef.Length - source.StructureOffset)));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(MathMl),
            MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(native)));
        Assert.Throws<ArgumentOutOfRangeException>(() => MathTypeMtefFontSize.ScaleMtefPreferenceSizes(source.Mtef, 1.5));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(double.NaN)]
    public void InvalidSizesAreRejected(double size)
        => Assert.Throws<ArgumentOutOfRangeException>(() =>
            MathTypeMtefFontSize.CreateEquationNativeAtFontSize(MathMl, false, size));
}
