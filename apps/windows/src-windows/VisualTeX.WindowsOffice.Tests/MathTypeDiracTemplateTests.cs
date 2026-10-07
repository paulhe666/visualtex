using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeDiracTemplateTests
{
    [Fact]
    public void FrontendExplicitKetRowUsesTheSameTemplateAsMfenced()
    {
        var row = "<math><mrow data-mjx-texclass='INNER'><mo data-mjx-texclass='OPEN'>|</mo><msub><mi>ψ</mi><mn>19</mn></msub><mo data-mjx-texclass='CLOSE'>⟩</mo></mrow></math>";
        var fenced = "<math><mfenced open='|' close='⟩' separators=''><msub><mi>ψ</mi><mn>19</mn></msub></mfenced></math>";
        var actual = MathTypeMtefCodec.CreateEquationNative(row, false);
        Assert.Equal(MathTypeMtefCodec.CreateEquationNative(fenced, false).Mtef, actual.Mtef);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(row), MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(actual.EquationNative)));
    }

    [Theory]
    [InlineData(1, "⟨", "|", "x")]
    [InlineData(2, "|", "⟩", "y")]
    [InlineData(3, "⟨", "⟩", "x|y")]
    public void DocumentedDiracRecordReadsBothSlotsAndOptionalAngles(int variation, string open, string close, string text)
    {
        var seed = MathTypeMtefCodec.CreateEquationNative("<math><mi>x</mi></math>", false);
        // MTEF v5 tmDIRAC: left LINE, right LINE, optional left angle,
        // mandatory bar, optional right angle. Independent of the writer.
        var records = new List<byte> { 1, 0, 3, 0, 30, (byte)variation, 0 };
        records.AddRange((variation & 1) != 0 ? new byte[] { 1, 0, 2, 0, 131, 120, 0, 0 } : new byte[] { 1, 1 });
        records.AddRange((variation & 2) != 0 ? new byte[] { 1, 0, 2, 0, 131, 121, 0, 0 } : new byte[] { 1, 1 });
        if ((variation & 1) != 0) records.AddRange(new byte[] { 2, 0, 150, 0x29, 0x23 });
        records.AddRange(new byte[] { 2, 0, 150, 0x07, 0xEC });
        if ((variation & 2) != 0) records.AddRange(new byte[] { 2, 0, 150, 0x2A, 0x23 });
        records.AddRange(new byte[] { 0, 0, 0 });
        var mtef = seed.Mtef.Take(seed.StructureOffset).Concat(records).ToArray();
        var native = seed.EquationNative.Take(28).Concat(mtef).ToArray();
        BitConverter.GetBytes(mtef.Length).CopyTo(native, 8);
        var xml = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native));
        var fence = Assert.Single(xml.Descendants(), e => e.Name.LocalName == "mfenced");
        Assert.Equal(open, (string?)fence.Attribute("open"));
        Assert.Equal(close, (string?)fence.Attribute("close"));
        Assert.Equal(text, fence.Value);
    }

    [Theory]
    [InlineData("⟨", "|", 1)]
    [InlineData("|", "⟩", 2)]
    public void BraAndKetUseNativeDiracWithNestedBodyAndScripts(string open, string close, int variation)
    {
        var source = $"<math><mfenced open='{open}' close='{close}' separators=''><mfrac><msub><mi>ψ</mi><mn>2</mn></msub><msup><mi>x</mi><mn>3</mn></msup></mfrac></mfenced></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var records = native.Mtef.Skip(native.StructureOffset).ToArray();
        Assert.Equal(new byte[] { 1, 0, 3, 0, 30, (byte)variation, 0 }, records.Take(7));
        var actual = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(actual));
        var rewritten = MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, actual, false);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative)));
        var xml = XDocument.Parse(actual);
        Assert.Single(xml.Descendants(), e => e.Name.LocalName == "mfrac");
        Assert.Single(xml.Descendants(), e => e.Name.LocalName == "msub");
        Assert.Single(xml.Descendants(), e => e.Name.LocalName == "msup");
    }
}
