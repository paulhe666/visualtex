using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathSpaceContractTests
{
    [Theory]
    [InlineData("1em", "1em", @"\quad", 0xEF05)]
    [InlineData("+001.000em", "1em", @"\quad", 0xEF05)]
    [InlineData("2.00em", "2em", @"\qquad", 0xEF06)]
    [InlineData("0.123456789em", "0.123456789em", @"\hspace{0.123456789em}", 0)]
    [InlineData("-0.333333em", "-0.333333em", @"\hspace{-0.333333em}", 0)]
    [InlineData("1.00000000000000000000000000001em", "1.00000000000000000000000000001em", @"\hspace{1.00000000000000000000000000001em}", 0)]
    [InlineData("0.99999999999999999999999999999em", "0.99999999999999999999999999999em", @"\hspace{0.99999999999999999999999999999em}", 0)]
    [InlineData("2.00000000000000000000000000001em", "2.00000000000000000000000000001em", @"\hspace{2.00000000000000000000000000001em}", 0)]
    [InlineData("1.25ex", "1.25ex", @"\hspace{1.25ex}", 0)]
    [InlineData("3pt", "3pt", @"\hspace{3pt}", 0)]
    [InlineData("12px", "12px", @"\hspace{12px}", 0)]
    [InlineData(".5in", "0.5in", @"\hspace{.5in}", 0)]
    [InlineData("2cm", "2cm", @"\hspace{2cm}", 0)]
    [InlineData("-2mm", "-2mm", @"\hspace{-2mm}", 0)]
    [InlineData("+01.20pc", "1.2pc", @"\hspace{+01.20pc}", 0)]
    [InlineData("-0.00em", "0em", @"\hspace{-0.00em}", 0)]
    [InlineData(" .500 pt ", "0.5pt", @"\hspace{.500pt}", 0)]
    [InlineData(null, "0em", @"\hspace{0em}", 0)]
    public void SharedDescriptorRecoversExactDimensions(string? width, string canonical, string latex, int nativeCode)
    {
        var descriptor = MathSpaceContract.ParseWidth(width);
        Assert.True(descriptor.SourceSupported);
        Assert.Equal(canonical, descriptor.CanonicalWidth);
        Assert.Equal(latex, descriptor.LatexSource!.Trim());
        Assert.Equal(nativeCode, descriptor.NativeSpace?.MtCode ?? 0);
        var attribute = width is null ? "" : $" width='{width}'";
        Assert.Equal("x" + latex + (nativeCode != 0 ? " " : "") + "y",
            MathMlToLatexConverter.Convert($"<math><mi>x</mi><mspace{attribute}/><mi>y</mi></math>"));
    }

    [Theory]
    [InlineData("thinmathspace", MathSpaceSourceSupport.SymbolicMathSpace)]
    [InlineData("negativeveryverythickmathspace", MathSpaceSourceSupport.SymbolicMathSpace)]
    [InlineData("12%", MathSpaceSourceSupport.NeedsContext)]
    [InlineData("1mu", MathSpaceSourceSupport.NotRepresented)]
    [InlineData("1e-3em", MathSpaceSourceSupport.NotRepresented)]
    [InlineData("1bp", MathSpaceSourceSupport.NotRepresented)]
    [InlineData("1EM", MathSpaceSourceSupport.NotRepresented)]
    [InlineData("1em plus 1em", MathSpaceSourceSupport.NotRepresented)]
    [InlineData("1em}\\input{payload}", MathSpaceSourceSupport.NotRepresented)]
    [InlineData("NaNem", MathSpaceSourceSupport.NotRepresented)]
    public void UnsupportedSourceFormsStayExplicit(string width, object expected)
    {
        var descriptor = MathSpaceContract.ParseWidth(width);
        Assert.False(descriptor.SourceSupported);
        Assert.Equal((MathSpaceSourceSupport)expected, descriptor.SourceSupport);
        Assert.Null(descriptor.LatexSource);
        Assert.Null(descriptor.NativeSpace);
        Assert.NotEmpty(descriptor.UnsupportedReason!);
        Assert.Contains("uncharacterized", descriptor.NativeSupport);
    }

    [Fact]
    public void CurrentUnsupportedConverterFallbackIsNotClaimedAsRecovered()
    {
        Assert.Equal(@"\,", MathMlToLatexConverter.Convert("<math><mspace width='12%'/></math>"));
        Assert.False(MathSpaceContract.ParseWidth("12%").SourceSupported);
    }

    [Fact]
    public void WidthFidelityComparisonIsSeparateFromTokenSemantics()
    {
        Assert.True(MathSpaceContract.HaveSameWidth("001.000em", "1em"));
        Assert.True(MathSpaceContract.HaveSameWidth("-0.00px", "+0px"));
        Assert.True(MathSpaceContract.HaveSameWidth("0em", "-0.00pt"));
        Assert.True(MathSpaceContract.HaveSameWidth("0cm", "0px"));
        Assert.True(MathSpaceContract.HaveSameWidth("negativeverythinmathspace", "negativeverythinmathspace"));
        Assert.False(MathSpaceContract.HaveSameWidth("1em", "1.00000000000000000000000000001em"));
        Assert.False(MathSpaceContract.HaveSameWidth("1em", "10pt"));
        Assert.False(MathSpaceContract.HaveSameWidth("thinmathspace", "0.167em"));
        Assert.False(MathSpaceContract.HaveSameWidth("12%", "12%"));
    }

    [Fact]
    public void EditedWidthOverridesStaleNativeIdentity()
    {
        Assert.Equal(@"x\hspace{1.25em}y", MathMlToLatexConverter.Convert(
            "<math><mi>x</mi><mspace width='1.25em' data-mtef-mtcode='61189' data-mtef-explicit-typeface='24'/><mi>y</mi></math>"));
        Assert.Null(MathSpaceContract.ParseWidth("1.25em").NativeSpace);
    }

    [Fact]
    public void BothNativeDirectionsShareOneVerifiedIdentity()
    {
        foreach (var mtCode in new[] { 0xEF05, 0xEF06 })
        {
            var native = MathSpaceContract.FindNativeSpace(mtCode)!;
            Assert.Same(native, MathSpaceContract.ParseWidth(native.MathMlWidth).NativeSpace);
            Assert.Contains("verified", native.SupportEvidence);
        }
        // TeX's spacing commands carry the code MathType's own importers store.
        foreach (var mtCode in new[] { 0xEF07, 0xEF03, 0xEF04, 0xEF01 })
        {
            var native = MathSpaceContract.FindNativeSpace(mtCode)!;
            Assert.Same(native, MathSpaceContract.ParseWidth(native.MathMlWidth).NativeSpace);
            Assert.Contains("MathType 7.9", native.SupportEvidence);
        }
        Assert.Null(MathSpaceContract.FindNativeSpace(0xEF02));
    }
}
