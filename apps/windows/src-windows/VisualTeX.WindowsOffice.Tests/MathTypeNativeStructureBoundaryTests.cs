using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeNativeStructureBoundaryTests
{
    // Only the vendor-format prefix is borrowed from a valid seed. Structural
    // bytes below are assembled independently of the production writer.
    private static byte[] Native(IEnumerable<byte> structure)
    {
        var seed = MathTypeMtefCodec.CreateEquationNative("<math><mi>x</mi></math>", false);
        var data = seed.EquationNative.Take(28 + seed.StructureOffset).Concat(structure).ToArray();
        Buffer.BlockCopy(BitConverter.GetBytes((uint)(data.Length - 28)), 0, data, 8, 4);
        return data;
    }

    private static readonly byte[] X = { 2, 0, 131, (byte)'x', 0 };

    [Theory]
    [InlineData("missing-end")]
    [InlineData("extra-character")]
    [InlineData("second-root")]
    public void ReaderRequiresTheEntireNativeEquation(string corruption)
    {
        var root = new byte[] { 1, 0 }.Concat(X).Concat(new byte[] { 0 });
        var tail = corruption switch
        {
            "missing-end" => Array.Empty<byte>(),
            "extra-character" => X.Concat(new byte[] { 0 }).ToArray(),
            _ => new byte[] { 1, 0 }.Concat(X).Concat(new byte[] { 0, 0 }).ToArray(),
        };
        Assert.Throws<InvalidDataException>(() => MathTypeMtefCodec.ReadEquationNativeMathMl(Native(root.Concat(tail))));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RecursiveNativeContainersHaveABoundary(bool template)
    {
        IEnumerable<byte> body = X;
        for (var i = 0; i < 270; i++)
            body = template
                // Underbar TMPL, main LINE, ENDs.
                ? new byte[] { 3, 0, 12, 0, 0, 1, 0 }.Concat(body).Concat(new byte[] { 0, 0 }).ToArray()
                : new byte[] { 1, 0 }.Concat(body).Concat(new byte[] { 0 }).ToArray();
        var data = Native(new byte[] { 1, 0 }.Concat(body).Concat(new byte[] { 0, 0 }));
        Assert.Throws<InvalidDataException>(() => MathTypeMtefCodec.ReadEquationNativeMathMl(data));
    }

    [Fact]
    public void FlatTrailingScriptsCannotBuildAnUnboundedMathMlTree()
    {
        // Native TMPLs are siblings, but each attaches to the previous result.
        // A recursion counter alone therefore cannot protect the resulting XML.
        var sup = new byte[] { 3, 0, 28, 0, 0, 1, 1, 1, 0 }.Concat(X).Concat(new byte[] { 0, 0 }).ToArray();
        var body = X.Concat(Enumerable.Range(0, 270).SelectMany(_ => sup));
        Assert.Throws<InvalidDataException>(() => MathTypeMtefCodec.ReadEquationNativeMathMl(
            Native(new byte[] { 1, 0 }.Concat(body).Concat(new byte[] { 0, 0 }))));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BoundedNativeInputAndTerminalFormattingRemainReadable(bool terminalFormatting)
    {
        IEnumerable<byte> body = X;
        for (var i = 0; i < 24; i++)
            body = new byte[] { 1, 0 }.Concat(body).Concat(new byte[] { 0 }).ToArray();
        var tail = terminalFormatting ? new byte[] { 10, 0 } : new byte[] { 0 };
        var actual = XElement.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(Native(body.Concat(tail))));
        Assert.Equal("x", actual.Value);
        Assert.All(actual.DescendantsAndSelf(), node => Assert.Equal("http://www.w3.org/1998/Math/MathML", node.Name.NamespaceName));
    }
}
