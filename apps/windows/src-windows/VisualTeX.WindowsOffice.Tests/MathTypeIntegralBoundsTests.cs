using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// Integral bounds sit beside the sign, as TeX draws \int_0^1, in display as
// well as inline. MathType's TeX import draws a display \int_0^1 like
// \int\limits_0^1; VisualTeX keeps TeX's placement and puts the bounds above
// and below only for \limits.
public sealed class MathTypeIntegralBoundsTests
{
    [Theory]
    [InlineData("block", "msubsup", false)]
    [InlineData("inline", "msubsup", false)]
    [InlineData("block", "munderover", true)] // \int\limits_0^1
    [InlineData("inline", "munderover", true)]
    public void BoundsSitBesideTheSignUnlessLimitsIsWritten(string display, string script, bool above)
    {
        var source = $"<math display=\"{display}\"><{script}><mo data-mjx-texclass=\"OP\">∫</mo><mn>0</mn><mn>1</mn></{script}>"
            + "<msup><mi>x</mi><mn>2</mn></msup></math>";
        var native = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, false, 12);

        // The integral template (selector 15): a single integral with a lower
        // and an upper bound (0x31), above and below the sign only with 0x40.
        var root = native.Mtef.Skip(native.StructureOffset).ToArray();
        var header = Enumerable.Range(0, root.Length - 3).First(i => root[i] == 3 && root[i + 2] == 15);
        Assert.Equal(0x31, root[header + 3] & 0x3F);
        Assert.Equal(above, (root[header + 3] & 0x40) != 0);

        var read = XElement.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative));
        var bounds = read.Descendants().First(e => e.Name.LocalName is "msubsup" or "munderover");
        Assert.Equal(above ? "munderover" : "msubsup", bounds.Name.LocalName);
    }
}
