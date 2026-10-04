using System.Text;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeUnsupportedFenceBindingTests
{
    private static readonly XNamespace P = "http://www.w3.org/1998/Math/MathML";
    private static string Source(string body) => $"<math xmlns=\"{P}\">{body}</math>";
    private static string Row(string open, string body, string close) =>
        "<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\" fence=\"true\" stretchy=\"true\" symmetric=\"true\">"
        + open + "</mo>" + body + "<mo data-mjx-texclass=\"CLOSE\" fence=\"true\" stretchy=\"true\" symmetric=\"true\">"
        + close + "</mo></mrow>";

    [Fact]
    public void F07UnsupportedCornerFencesKeepTheExactPreRecognitionNativeBytesAndGlyphBindings()
    {
        // CP218 original representative20 F07 stream, before shared source
        // recognition exposed the native template gap. This is a CHAR-binding
        // baseline, not evidence that a growing fence template was supported.
        const string expectedBase64 = "HAAAAAIAQsInAQAAAAAAAPzeVgot39QADACFCQUBAAcIRFNNVDcAABNXaW5BbGxCYXNpY0NvZGVQYWdlcwARBVRpbWVzIE5ldyBSb21hbgARA1N5bWJvbAARBUNvdXJpZXIgTmV3ABEETVQgRXh0cmEAE1dpbkFsbENvZGVQYWdlcwARBsvOzOUAEgAIIS9Fj0QvQVD0EA9HX0FQ8h8eQVD0FQ9BAPRF9CX0j0JfQQD0EA9DX0EA9I9F9CpfSPSPQQD0EA9A9I9Bf0j0EA9BKl9EX0X0X0X0X0EPDAEAAQABAgICAgACAAEBAQADAAEABAAFAAoBABNFdWNsaWRNYXRoMQARB0V1Y2xpZCBNYXRoIE9uZQAIBgACBH8cI4QCAIN4AAIEfx0jhgIEhisAKwIEfx4jhQIAg3kAAgR/HyOHAAA=";
        var source = Source(Row("⌜", "<mi>x</mi>", "⌝") + "<mo>+</mo>" + Row("⌞", "<mi>y</mi>", "⌟"));
        var native = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, false, 12.0);
        Assert.Equal(Convert.FromBase64String(expectedBase64), native.EquationNative);
        var recovered = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var corners = XDocument.Parse(recovered).Descendants(P + "mo")
            .Where(token => token.Value != "+").ToArray();
        Assert.Equal(new int?[] { 0x231C, 0x231D, 0x231E, 0x231F },
            corners.Select(token => (int?)token.Attribute("data-mtef-mtcode")));
        Assert.Equal(new int?[] { 0x84, 0x86, 0x85, 0x87 },
            corners.Select(token => (int?)token.Attribute("data-mtef-encoded8")));
        Assert.All(corners, token => Assert.Equal(-1, (int?)token.Attribute("data-mtef-explicit-typeface")));
        Assert.Contains("EuclidMath1\0", Encoding.ASCII.GetString(native.Mtef));
        Assert.Contains("Euclid Math One\0", Encoding.ASCII.GetString(native.Mtef));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(recovered));
    }

    [Theory]
    [InlineData("(", ")")]
    [InlineData("[", "]")]
    [InlineData("{", "}")]
    [InlineData("{", "")]
    [InlineData("", "}")]
    [InlineData("⟨", "⟩")]
    public void SupportedNativeTemplatesStillMaterializeFences(string open, string close)
    {
        var source = Source(Row(open, "<mi>x</mi>", close));
        var native = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, false, 12.0);
        var recovered = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var fence = Assert.Single(XDocument.Parse(recovered).Descendants(P + "mfenced"));
        Assert.Equal(open, (string?)fence.Attribute("open"));
        Assert.Equal(close, (string?)fence.Attribute("close"));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(recovered));
    }
}
