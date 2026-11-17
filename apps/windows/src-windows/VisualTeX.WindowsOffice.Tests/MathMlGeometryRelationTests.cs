using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// The congruence and similarity signs of Chinese geometry read from Word's math as the names the editor stores them
// under: ≌ is \backcong, ∽ is \backsim (exam-zh's \cong and \sim are the editor's spelling of them, not the stored names).
public class MathMlGeometryRelationTests
{
    [Theory]
    [InlineData("≌", @"\backcong")]
    [InlineData("∽", @"\backsim")]
    [InlineData("≅", @"\cong")]
    [InlineData("∼", @"\sim")]
    public void RelationReadsAsItsStoredName(string sign, string command)
    {
        var latex = MathMlToLatexConverter.Convert($"<math><mo>△</mo><mi>A</mi><mo>{sign}</mo><mo>△</mo><mi>B</mi></math>");
        Assert.Contains(command + " ", latex + " ");
        Assert.DoesNotContain(sign, latex);
    }

    [Theory]
    [InlineData("≌", @"\backcong")]
    [InlineData("∽", @"\backsim")]
    public void LoneSignStoredAsTextIsTheMathSign(string sign, string command)
    {
        var latex = MathMlToLatexConverter.Convert($"<math><mi>A</mi><mtext>{sign}</mtext><mi>B</mi></math>");
        Assert.Contains(command, latex);
        Assert.DoesNotContain(@"\text", latex);
    }
}
