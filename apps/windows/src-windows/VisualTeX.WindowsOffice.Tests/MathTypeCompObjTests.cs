using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeCompObjTests
{
    private static readonly Guid MathTypeClsid = new("0002CE03-0000-0000-C000-000000000046");

    // \x01CompObj of a genuine MathType 6 Equation.DSMT4 storage (jure/mathtype
    // fixture mathtype5/299.bin).
    private const string GenuineMathType6CompObj =
        "0100feff030a0000ffffffff03ce020000000000c000000000000046"
        + "160000004d6174685479706520362e30204571756174696f6e00"
        + "0c0000004d6174685479706520454600"
        + "0f0000004571756174696f6e2e44534d543400"
        + "f439b271000000000000000000000000";

    [Fact]
    public void MatchesTheStreamMathTypeWrites()
    {
        var bytes = MathTypeCompObj.Build(MathTypeClsid, "MathType 6.0 Equation", "Equation.DSMT4");
        Assert.Equal(GenuineMathType6CompObj, Convert.ToHexString(bytes).ToLowerInvariant());
    }

    [Fact]
    public void ProgIdNeverDependsOnTheRegistry()
    {
        // A machine without MathType used to get an empty ProgID here.
        var bytes = MathTypeCompObj.Build(MathTypeClsid, "MathType 7.0 Equation", null);
        Assert.Contains("0f0000004571756174696f6e2e44534d543400", Convert.ToHexString(bytes).ToLowerInvariant());
    }
}
