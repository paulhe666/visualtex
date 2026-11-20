using VisualTeX.WordVsto;
using System.Xml.Linq;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeColorDefinitionTests
{
    private static readonly byte[] Black = { 16, 0, 0, 0, 0, 0, 0, 0 };
    private static readonly byte[] Red = { 16, 0, 0xE8, 3, 0, 0, 0, 0 };
    private const string Cases = "<math><mfenced open='{' close=''><mtable>"
        + "<mtr><mtd><mi>x</mi><mo>=</mo><mn>0</mn></mtd></mtr>"
        + "<mtr><mtd><msqrt><mfrac><mi>m</mi><mn>2</mn></mfrac></msqrt></mtd></mtr>"
        + "</mtable></mfenced></math>";

    [Theory]
    [InlineData(Cases)]
    [InlineData("<math><mtable><mtr><mtd><msqrt><mi>x</mi></msqrt></mtd></mtr></mtable></math>")]
    public void StandalonePaletteIsDefinedAndSurvivesRepeatedNativeEdits(string source)
    {
        var created = MathTypeMtefCodec.CreateEquationNative(source, inline: false);
        var palette = created.Mtef.AsSpan(0, created.StructureOffset).IndexOf(Black);
        Assert.True(palette >= 0, "Native template colors must be defined in the equation itself.");
        var edited = MathTypeMtefCodec.RewriteEquationNative(created.EquationNative, source, inline: false);
        Assert.Equal(created.Mtef, edited.Mtef);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(edited.EquationNative)));
    }

    [Fact]
    public void RootWithoutColorReferencesDoesNotInventAPalette()
    {
        // tmROOT has no COLOR record. The blanket palette assertion predates
        // the native-prefix fidelity fix and incorrectly demanded unused data.
        const string source="<math><mroot><mi>x</mi><mn>3</mn></mroot></math>";
        var native=MathTypeMtefCodec.CreateEquationNative(source,false);
        Assert.True(native.Mtef.AsSpan(0,native.StructureOffset).IndexOf(Black)<0);
        var rewritten=MathTypeMtefCodec.RewriteEquationNative(native.EquationNative,source,false);
        Assert.Equal(native.Mtef,rewritten.Mtef);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source),MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative)));
    }

    [Fact]
    public void EditingLegacyEquationWithoutPaletteDefinesItsTemplateColor()
    {
        var created = MathTypeMtefCodec.CreateEquationNative(Cases, inline: false);
        var at = created.Mtef.AsSpan(0, created.StructureOffset).IndexOf(Black);
        Assert.True(at >= 0);
        var legacy = created.Mtef.Take(at).Concat(created.Mtef.Skip(at + Black.Length)).ToArray();
        var edited = MathTypeMtefCodec.RewriteEquationNative(Native(created.EquationNative, legacy), Cases, inline: false);
        Assert.Equal(created.Mtef, edited.Mtef);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EditingPreservesExistingPaletteAndItsIndices(bool rootLeading)
    {
        var created = MathTypeMtefCodec.CreateEquationNative(Cases, inline: false);
        var at = created.Mtef.AsSpan(0, created.StructureOffset).IndexOf(Black);
        Assert.True(at >= 0);
        var source = created.Mtef.ToList();
        source.RemoveRange(at, Black.Length);
        var insertion = rootLeading ? created.StructureOffset - Black.Length + 2 : at;
        source.InsertRange(insertion, Red);
        var bytes = source.ToArray();
        var edited = MathTypeMtefCodec.RewriteEquationNative(Native(created.EquationNative, bytes), Cases, inline: false);
        Assert.Equal(bytes, edited.Mtef);
    }

    [Theory]
    [InlineData("<math><mstyle mathcolor='#DC2626'><mi>x</mi><mo>+</mo><mi>y</mi></mstyle></math>")]
    [InlineData("<math mathcolor='#2563EB'><mfrac><mi>x</mi><mn>2</mn></mfrac></math>")]
    [InlineData("<math><mstyle mathcolor='#16A34A'><mtable><mtr><mtd><mi>x</mi></mtd><mtd><mi>y</mi></mtd></mtr></mtable></mstyle></math>")]
    [InlineData("<math><mstyle mathcolor='#DC2626'><mover accent='true'><mi>x</mi><mo>~</mo></mover></mstyle></math>")]
    public void ExplicitMathMlColorSurvivesStandaloneCreation(string source)
    {
        var native = MathTypeMtefCodec.CreateEquationNative(source, inline: false);
        var readback = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var expected = XDocument.Parse(source).Root!
            .DescendantsAndSelf()
            .Where(element => element.Name.LocalName is "mi" or "mn" or "mo")
            .Select(MathMlColorContract.Effective)
            .ToArray();
        var actual = XDocument.Parse(readback).Root!
            .DescendantsAndSelf()
            .Where(element => element.Name.LocalName is "mi" or "mn" or "mo")
            .Select(MathMlColorContract.Effective)
            .ToArray();

        Assert.Equal(expected, actual);
        Assert.Equal(
            MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(readback));
    }

    [Fact]
    public void EditingVisualTexMathTypeOleAppliesNestedSelectionColors()
    {
        const string plain = "<math><mi>x</mi><mi>y</mi><mi>z</mi><mi>q</mi></math>";
        const string colored = "<math><mstyle mathcolor='#DC2626'><mi>x</mi>"
            + "<mstyle mathcolor='#2563EB'><mi>y</mi></mstyle><mi>z</mi></mstyle><mi>q</mi></math>";
        var seed = MathTypeMtefCodec.CreateEquationNative(plain, inline: true);
        var edited = MathTypeMtefCodec.RewriteEquationNative(seed.EquationNative, colored, inline: true);
        var readback = MathTypeMtefCodec.ReadEquationNativeMathMl(edited.EquationNative);
        var tokens = XDocument.Parse(readback).Root!
            .DescendantsAndSelf()
            .Where(element => element.Name.LocalName == "mi")
            .ToDictionary(element => element.Value, MathMlColorContract.Effective);

        Assert.Equal("#DC2626", tokens["x"]);
        Assert.Equal("#2563EB", tokens["y"]);
        Assert.Equal("#DC2626", tokens["z"]);
        Assert.Null(tokens["q"]);
        Assert.NotEqual(
            MathTypeMtefCodec.SemanticSignature(plain),
            MathTypeMtefCodec.SemanticSignature(colored));
        Assert.Equal(
            MathTypeMtefCodec.SemanticSignature(colored),
            MathTypeMtefCodec.SemanticSignature(readback));

        var repeated = MathTypeMtefCodec.RewriteEquationNative(
            edited.EquationNative,
            readback,
            inline: true);
        var repeatedReadback = MathTypeMtefCodec.ReadEquationNativeMathMl(
            repeated.EquationNative);
        Assert.Equal(
            MathTypeMtefCodec.SemanticSignature(readback),
            MathTypeMtefCodec.SemanticSignature(repeatedReadback));
        Assert.Null(XDocument.Parse(repeatedReadback).Root!
            .DescendantsAndSelf()
            .Single(element => element.Name.LocalName == "mi" && element.Value == "q")
            .Attribute("mathcolor"));
    }

    private static byte[] Native(byte[] headerSource, byte[] mtef)
    {
        var result = headerSource.Take(28).Concat(mtef).ToArray();
        BitConverter.GetBytes((uint)mtef.Length).CopyTo(result, 8);
        return result;
    }
}
