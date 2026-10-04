using VisualTeX.WordVsto;
namespace VisualTeX.WindowsOffice.Tests;
public sealed class MathMlFunctionTokenRecoveryTests
{
 [Theory]
 [InlineData("<mi>sin</mi><mi>x</mi>","\\operatorname{sin}x")]
 [InlineData("<mi>sin</mi><mo>⁡</mo><mi>x</mi>","\\operatorname{sin}⁡x")]
 [InlineData("<mi>foo</mi><mi>x</mi>","\\operatorname{foo}x")]
 [InlineData("<mi data-mjx-auto-op='false'>foo</mi><mi>x</mi>","\\mathrm{foo}x")]
 [InlineData("<mi data-mjx-texclass='OP' data-mjx-auto-op='false'>sin</mi><mi>x</mi>","\\operatorname{sin}x")]
 [InlineData("<mi data-mjx-texclass='ORD'>sin</mi><mi>x</mi>","\\mathrm{sin}x")]
 [InlineData("<msup><mi>foo</mi><mn>2</mn></msup><mo>⁡</mo><mi>x</mi>","\\operatorname{foo}^{2}⁡x")]
 [InlineData("<mi>foo</mi><mo>⁡</mo><mo>+</mo><mi>x</mi>","\\operatorname{foo}⁡+x")]
 [InlineData("<mi>x</mi><mi>y</mi>","xy")]
 [InlineData("<mo>lim</mo>","\\lim")]
 [InlineData("<munder><mo>lim</mo><mi>x</mi></munder>","\\lim_{x}")]
 [InlineData("<munder><mo>max</mo><mi>x</mi></munder>","\\max_{x}")]
 public void PreservesTokenBoundaryAndActualApplicationSlot(string body,string expected)
  => Assert.Equal(expected,MathMlToLatexConverter.Convert("<math>"+body+"</math>"));
}
