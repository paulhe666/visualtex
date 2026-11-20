using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class OleCompoundFileDigestTests
{
    private static byte[] Fixture(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "ole-container", name));

    [Fact]
    public void WpsRewrittenContainerOfTheSameObjectHasTheSameContent()
    {
        // A user's MathType object before and after WPS undid a failed
        // conversion: the root entry's timestamp differs, every stream is equal.
        var before = Fixture("wps-undo-before.bin");
        var after = Fixture("wps-undo-after.bin");
        Assert.NotEqual(before, after);
        var digest = OleCompoundFileDigest.TryCompute(before);
        Assert.NotNull(digest);
        Assert.Equal(digest, OleCompoundFileDigest.TryCompute(after));
    }

    [Fact]
    public void AChangedStreamChangesTheContent()
    {
        var before = Fixture("wps-undo-before.bin");
        var changed = (byte[])before.Clone();
        // Equation Native (297 bytes) lives in the mini stream at file offset
        // 2304 in this fixture; flip one byte inside it.
        changed[2400] ^= 0x5A;
        Assert.NotEqual(OleCompoundFileDigest.TryCompute(before), OleCompoundFileDigest.TryCompute(changed));
    }

    [Fact]
    public void BytesThatAreNotACompoundFileHaveNoDigest()
    {
        Assert.Null(OleCompoundFileDigest.TryCompute(new byte[] { 1, 2, 3 }));
        Assert.Null(OleCompoundFileDigest.TryCompute(new byte[1024]));
    }
}
