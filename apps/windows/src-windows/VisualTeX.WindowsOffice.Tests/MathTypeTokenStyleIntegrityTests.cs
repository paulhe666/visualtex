using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;
namespace VisualTeX.WindowsOffice.Tests;
public sealed class MathTypeTokenStyleIntegrityTests
{
    [Theory]
    [InlineData("bold",@"\mathbf{123}")]
    [InlineData("italic",@"\mathit{123}")]
    [InlineData("bold-italic",@"\mathbfit{123}")]
    [InlineData("monospace",@"\mathtt{123}")]
    [InlineData("sans-serif",@"\mathsf{123}")]
    [InlineData("bold-sans-serif",@"\boldsymbol{\mathsf{123}}")]
    [InlineData("double-struck",@"\mathbb{123}")]
    public void NumberStyleIsValidatedWrittenAndRecovered(string style,string latex)
    {
        var source="<math><mn mathvariant='"+style+"'>123</mn></math>";
        var expected=MathTypeMtefCodec.SemanticSignature(source);
        Assert.NotEqual(expected,MathTypeMtefCodec.SemanticSignature("<math><mn>123</mn></math>"));
        var native=MathTypeMtefCodec.CreateEquationNativeAtFontSize(source,false,12);
        var readback=MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var number=Assert.Single(XElement.Parse(readback).Descendants().Where(n=>n.Name.LocalName=="mn"));
        Assert.Equal(style,(string?)number.Attribute("mathvariant"));
        Assert.Equal(expected,MathTypeMtefCodec.SemanticSignature(readback));
        Assert.Equal(latex,MathMlToLatexConverter.Convert(readback));
        var rewritten=MathTypeMtefCodec.RewriteEquationNative(native.EquationNative,source,false);
        Assert.Equal(expected,MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative)));
    }
}
