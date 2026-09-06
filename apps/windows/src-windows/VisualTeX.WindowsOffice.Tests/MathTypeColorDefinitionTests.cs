using VisualTeX.WordVsto;

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
    [InlineData("<math><mroot><mi>x</mi><mn>3</mn></mroot></math>")]
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

    private static byte[] Native(byte[] headerSource, byte[] mtef)
    {
        var result = headerSource.Take(28).Concat(mtef).ToArray();
        BitConverter.GetBytes((uint)mtef.Length).CopyTo(result, 8);
        return result;
    }
}
