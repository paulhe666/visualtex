using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeNativeScriptEmissionTests
{
    [Theory]
    [InlineData("msub", "<mn>1</mn>", 27)]
    [InlineData("msup", "<mn>1</mn>", 28)]
    [InlineData("msubsup", "<mn>1</mn><mn>2</mn>", 29)]
    public void EmptyBaseScriptIsFirstInItsOwnLineNotAttachedToANullSlot(string tag, string slots, byte selector)
    {
        // Keep the empty base inside an outer script slot: there is no following
        // atom that could instead make it a legitimate preceding-script pair.
        var mtef = Generate($"<msub><mi>R</mi><{tag}><mrow />{slots}</{tag}></msub><mi>z</mi>");
        Assert.True(Contains(mtef, new byte[] { 1, 0, 3, 0, selector, 0, 0 }),
            "A standalone empty-base script must be the first object in its own LINE; an empty preceding object changes its native baseline.");
        Assert.False(Contains(mtef, new byte[] { 1, 1, 3, 0, selector, 0, 0 }));
    }

    [Theory]
    [InlineData("msub", "<mn>1</mn>", 27)]
    [InlineData("msup", "<mn>1</mn>", 28)]
    [InlineData("msubsup", "<mn>1</mn><mn>2</mn>", 29)]
    public void NestedScriptUsesSub2RatherThanRestartingSub(string tag, string slots, byte selector)
    {
        var mtef = Generate($"<msub><mi>R</mi><{tag}><mi>x</mi>{slots}</{tag}></msub><mi>z</mi>");
        Assert.True(Contains(mtef, new byte[] { 3, 0, selector, 0, 0, 12 }),
            "The nested template must switch from SUB to SUB2, not restart SUB.");
    }

    [Fact]
    public void NestedTemplateRestoresTheEnclosingSubSizeBeforeSiblingText()
    {
        var mtef = Generate("<msub><mi>R</mi><mrow><msub><mi>x</mi><mn>1</mn></msub><mi>y</mi></mrow></msub><mi>z</mi>");
        Assert.True(Contains(mtef, new byte[] { 0, 11, 2, 0, 0x83, (byte)'y', 0 }),
            "A nested script must return to its parent's SUB size before y.");
        Assert.True(Contains(mtef, new byte[] { 0, 10, 2, 0, 0x83, (byte)'z', 0 }),
            "The outer script must return to FULL before z.");
    }

    [Fact]
    public void LiteralPrimeOperatorDoesNotUseTheSymbolFontsContainsGlyph()
    {
        var mtef = Generate("<mi>R</mi><msub><mo>'</mo><mi>N</mi></msub>");
        Assert.True(Contains(mtef, new byte[] { 2, 0, 0x82, 0x27, 0 }),
            "A literal legacy prime must retain Function/Times apostrophe, not Symbol font position 0x27 (contains).");
    }

    private static byte[] Generate(string body)
    {
        var result = MathTypeMtefCodec.CreateEquationNative("<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" + body + "</math>", false);
        return result.Mtef.Skip(result.StructureOffset).ToArray();
    }

    private static bool Contains(byte[] data, byte[] pattern)
    {
        for (var i = 0; i <= data.Length - pattern.Length; i++)
        {
            var match = true;
            for (var j = 0; j < pattern.Length && match; j++) match = data[i + j] == pattern[j];
            if (match) return true;
        }
        return false;
    }
}
