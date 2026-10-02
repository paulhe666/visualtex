using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeEmptyBaseScriptTests
{
    private static readonly XNamespace Ns = "http://www.w3.org/1998/Math/MathML";

    [Theory]
    [InlineData(27, "msub")]
    [InlineData(28, "msup")]
    [InlineData(29, "msubsup")]
    public void LeadingNativeScriptHasAnEmptyBaseAndDoesNotConsumeTheFollowingCharacter(byte selector, string element)
    {
        // ScrBoxClass stores sub/sup slots, not a base slot. Genuine native
        // objects can put that template first in a LINE, with no base CHAR/LINE.
        var template = new List<byte> { 3, 0, selector, 0, 0, 11 };
        var filled = new byte[] { 1, 0, 2, 0, 0x88, (byte)'1', 0, 0 };
        var absent = new byte[] { 1, 1 };
        template.AddRange(selector == 28 ? absent : filled);
        template.AddRange(selector == 27 ? absent : filled);
        template.Add(0);
        var body = new byte[] { 1, 0 }.Concat(template)
            .Concat(new byte[] { 10, 2, 0, 0x83, (byte)'z', 0, 0, 0 }).ToArray();
        var native = WithNativeRoot(body);
        var mathml = MathTypeMtefCodec.ReadEquationNativeMathMl(native);
        var root = XDocument.Parse(mathml).Root!;
        var script = root.Elements().First();
        Assert.Equal(Ns + element, script.Name);
        Assert.Equal(Ns + "mrow", script.Elements().First().Name);
        Assert.Empty(script.Elements().First().Nodes());
        Assert.Equal("z", root.Elements().Last().Value);
        Assert.Equal(2, root.Elements().Count());
        AssertRoundTrip(native, mathml);
    }

    [Fact]
    public void CopiedEquation255PreservesItsNestedEmptyBaseSubscript()
    {
        // Exact Equation Native bytes captured from the user's (2-55), before
        // any edit. A subscript slot itself starts with another tmSUB record.
        var native = Convert.FromBase64String(
            "HAAAAAIAScOAAgAAAAAAAFTa7wgt334ApAFvCAUBAAcIRFNNVDcAARNXaW5BbGxCYXNpY0NvZGVQYWdlcwARBVRpbWVzIE5ldyBSb21hbgARA1N5bWJvbAARBUNvdXJpZXIgTmV3ABEETVQgRXh0cmEAE1dpbkFsbENvZGVQYWdlcwARBsvOzOUAEgAIIS9Fj0QvQVD0EA9HX0FQ8h8eQVD0FQ9BAPRF9CX0j0JfQQD0EA9DX0EA9I9F9CpfSPSPQQD0EA9A9I9Bf0j0EA9BKl9EX0X0X0X0X0EPDAEAAQABAgICAgACAAEBAQADAAEABAAFAAoEAAEBAQACAINSAAMAHQAACwEAAgCDeAAAAQACAIgyAAAACgIEhj0APQIAg1IAAwAbAAALAQACAIgxAAABAQAKAgCCLwACAINSAAMAGwAACwEAAgCIMgAAAQEACgIEhsUi1wIAg1IAAwAbAAALAQACAINOAAABAQAKAgSGxSLXAgCDUgADABsAAAsBAAIAiDIAAAEBAAoCAIIvAAIAg1IAAwAbAAALAQADABsAAAwBAAIAiDEAAAEBAAALAQEACgIEhsUi1wIAg1IAAgCCJwADABsAAAsBAAIAg04AAAEBAAAKAQACAINSAAMAGwAACwEAAgCDeAAAAQEACgIEhj0APQMACgAAAQACAINSAAMAGwAACwEAAgCDTgAAAQEACgIEhsUi1wIAg1IAAgCCJwADABsAAAsBAAIAg04AAAEBAAABAQAKAgSGSCK7AwALAAABAAIAiDEAAAEAAgCIMgAAAAIAgigAAgCDUgADABsAAAsBAAIAg04AAAEBAAoCBIYrACsCAINSAAIAgicAAwAbAAALAQACAINOAAABAQAKAgCCKQAAAAA=");
        var mathml = MathTypeMtefCodec.ReadEquationNativeMathMl(native);
        var root = XDocument.Parse(mathml).Root!;
        var nested = Assert.Single(root.Descendants(Ns + "msub").Where(e =>
            e.Elements().First().Name == Ns + "mrow" && !e.Elements().First().Nodes().Any()));
        Assert.Equal("1", nested.Elements().Last().Value);
        Assert.Equal(Ns + "msub", nested.Parent!.Name);
        Assert.Equal("R", nested.Parent.Elements().First().Value);
        AssertRoundTrip(native, mathml);
        var latex = MathMlToLatexConverter.Convert(mathml);
        Assert.Contains("{}", latex, StringComparison.Ordinal);
    }

    private static void AssertRoundTrip(byte[] native, string mathml)
    {
        var rewritten = MathTypeMtefCodec.RewriteEquationNative(native, mathml, inline: true);
        var reread = MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(mathml), MathTypeMtefCodec.SemanticSignature(reread));
    }

    private static byte[] WithNativeRoot(byte[] body)
    {
        var seed = MathTypeMtefCodec.CreateEquationNative($"<math xmlns=\"{Ns}\"><mi>x</mi></math>", inline: true);
        var mtef = seed.Mtef.Take(seed.StructureOffset).Concat(body).ToArray();
        var native = seed.EquationNative.Take(28).Concat(mtef).ToArray();
        BitConverter.GetBytes(mtef.Length).CopyTo(native, 8);
        return native;
    }
}
