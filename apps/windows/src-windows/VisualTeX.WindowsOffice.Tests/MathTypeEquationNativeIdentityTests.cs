using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// MathType writes EQNOLEFILEHDR.cf and reserved2..reserved4 from uninitialized
// memory. Re-saving an unchanged equation (which Office requests while the
// equation is open in MathType's own window) therefore changes those bytes.
// The source-unchanged checks must compare the MTEF payload only.
public sealed class MathTypeEquationNativeIdentityTests
{
    private const string Fraction =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>";
    private const string Root =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msqrt><mi>a</mi></msqrt></math>";

    private static byte[] Native(string mathMl) =>
        MathTypeMtefCodec.CreateEquationNative(mathMl, false).EquationNative;

    // Header bytes 6..7 (cf) and 16..27 (reserved2..4) as MathType 7 wrote them
    // for one unchanged PowerPoint equation on two consecutive saves.
    private static byte[] WithHeader(byte[] native, ushort clipboardFormat, uint reserved2, uint reserved3, uint reserved4)
    {
        var copy = (byte[])native.Clone();
        BitConverter.GetBytes(clipboardFormat).CopyTo(copy, 6);
        BitConverter.GetBytes(reserved2).CopyTo(copy, 16);
        BitConverter.GetBytes(reserved3).CopyTo(copy, 20);
        BitConverter.GetBytes(reserved4).CopyTo(copy, 24);
        return copy;
    }

    [Fact]
    public void ResavedHeaderWithSamePayloadIsTheSameEquation()
    {
        var native = Native(Fraction);
        var firstSave = WithHeader(native, 0xC310, 0x0A1B43A4, 0x00DB376D, 0x099D000C);
        var secondSave = WithHeader(native, 0xC310, 0x0A1B3D74, 0x00DB376D, 0x099D000C);
        Assert.False(firstSave.AsSpan().SequenceEqual(secondSave));
        Assert.True(MathTypeMtefCodec.EquationNativeContentEquals(firstSave, secondSave));
    }

    [Fact]
    public void ClipboardFormatFromAnotherSessionIsNotContent()
    {
        var native = Native(Root);
        Assert.True(MathTypeMtefCodec.EquationNativeContentEquals(
            WithHeader(native, 0xC310, 0, 0, 0), WithHeader(native, 0xC242, 0, 0, 0)));
    }

    [Fact]
    public void DifferentEquationsAreDifferent()
    {
        Assert.False(MathTypeMtefCodec.EquationNativeContentEquals(Native(Fraction), Native(Root)));
    }

    [Fact]
    public void OneChangedPayloadByteIsAChange()
    {
        var native = Native(Fraction);
        var edited = (byte[])native.Clone();
        edited[edited.Length - 3] ^= 0x01;
        Assert.False(MathTypeMtefCodec.EquationNativeContentEquals(native, edited));
    }

    [Fact]
    public void DeclaredPayloadLengthIsPartOfTheContent()
    {
        var native = Native(Fraction);
        var shorter = (byte[])native.Clone();
        var declared = BitConverter.ToUInt32(shorter, 8);
        BitConverter.GetBytes(declared - 1).CopyTo(shorter, 8);
        Assert.False(MathTypeMtefCodec.EquationNativeContentEquals(native, shorter));
    }

    [Fact]
    public void BytesAfterTheDeclaredPayloadAreIgnored()
    {
        var native = Native(Fraction);
        var padded = new byte[native.Length + 4];
        native.CopyTo(padded, 0);
        padded[native.Length] = 0xCD;
        Assert.True(MathTypeMtefCodec.EquationNativeContentEquals(native, padded));
    }

    [Fact]
    public void UnreadableHeadersFallBackToExactBytes()
    {
        var a = new byte[] { 1, 2, 3 };
        Assert.True(MathTypeMtefCodec.EquationNativeContentEquals(a, new byte[] { 1, 2, 3 }));
        Assert.False(MathTypeMtefCodec.EquationNativeContentEquals(a, new byte[] { 1, 2, 4 }));
        Assert.False(MathTypeMtefCodec.EquationNativeContentEquals(Native(Fraction), a));
    }
}
