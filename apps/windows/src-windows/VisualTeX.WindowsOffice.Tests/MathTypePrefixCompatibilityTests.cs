using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypePrefixCompatibilityTests
{
    private const string Simple = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mn>1</mn></math>";
    private const string Styled = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi mathvariant=\"monospace\">x</mi><mo>+</mo><mi mathvariant=\"bold-italic\">a</mi></math>";

    private static byte[] Insert(byte[] native, int mtefOffset, byte[] record)
    {
        var result = new byte[native.Length + record.Length];
        Buffer.BlockCopy(native, 0, result, 0, 28 + mtefOffset);
        Buffer.BlockCopy(record, 0, result, 28 + mtefOffset, record.Length);
        Buffer.BlockCopy(native, 28 + mtefOffset, result, 28 + mtefOffset + record.Length, native.Length - 28 - mtefOffset);
        Buffer.BlockCopy(BitConverter.GetBytes((uint)(result.Length - 28)), 0, result, 8, 4);
        return result;
    }

    private static byte[] Extension(byte type, int length)
    {
        var bytes = new List<byte> { type };
        if (length < 255) bytes.Add((byte)length);
        else { bytes.Add(255); bytes.Add((byte)length); bytes.Add((byte)(length >> 8)); }
        // Payload deliberately resembles valid records. A parser must skip its
        // declared extent, never search the opaque payload for a plausible root.
        bytes.AddRange(Enumerable.Range(0, length).Select(i => (byte)(i % 20)));
        return bytes.ToArray();
    }

    [Theory]
    [InlineData(100, 0, false)]
    [InlineData(102, 3, false)]
    [InlineData(102, 260, false)]
    [InlineData(255, 1, false)]
    [InlineData(102, 3, true)]
    public void EveryReaderAndWriterTraversesTheSameExtensionAwarePrefix(int type, int length, bool afterInitialSize)
    {
        var original = MathTypeMtefCodec.CreateEquationNativeAtFontSize(Simple, true, 12);
        var offset = afterInitialSize ? original.StructureOffset : Array.IndexOf(original.Mtef, (byte)0, 5) + 2;
        var extended = Insert(original.EquationNative, offset, Extension((byte)type, length));
        Assert.Equal(12, MathTypeMtefCodec.ReadEquationNativeFullFontSize(extended));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(Simple),
            MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(extended)));
        var rewritten = MathTypeMtefCodec.RewriteEquationNativeAtFontSize(extended, Styled, true, 15);
        Assert.Equal(15, MathTypeMtefCodec.ReadEquationNativeFullFontSize(rewritten.EquationNative));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(Styled),
            MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative)));
        Assert.True(Contains(rewritten.Mtef, Extension((byte)type, length)));
    }

    [Fact]
    public void TruncatedExtensionDoesNotRecoverARootFromItsOpaquePayload()
    {
        var original = MathTypeMtefCodec.CreateEquationNative(Simple, true);
        var extended = Insert(original.EquationNative, 12, new byte[] { 102, 255, 255, 127, 1, 0, 0, 0 });
        Assert.ThrowsAny<IOException>(() => MathTypeMtefCodec.ReadEquationNativeMathMl(extended));
    }

    private static bool Contains(byte[] data, byte[] value)
        => Enumerable.Range(0, Math.Max(0, data.Length - value.Length + 1)).Any(i => data.Skip(i).Take(value.Length).SequenceEqual(value));
}
