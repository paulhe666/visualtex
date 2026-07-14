using System.Runtime.InteropServices;
using VisualTeX.WindowsOleBridge;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class ComReleaseAndDoubleClickTests
{
    [Fact]
    public void FinalReleaseInvalidatesARealComObject()
    {
        var type = Type.GetTypeFromProgID("Scripting.Dictionary", throwOnError: true);
        Assert.NotNull(type);
        var instance = Activator.CreateInstance(type!);
        Assert.NotNull(instance);
        Assert.True(Marshal.IsComObject(instance));

        ComRelease.Final(instance);

        Assert.Throws<InvalidComObjectException>(() =>
        {
            _ = ((dynamic)instance!).Count;
        });
    }

    [Fact]
    public void FormulaDoubleClickDeduplicatesOnlyTheSamePersistentTarget()
    {
        var deduplicator = new FormulaDoubleClickDeduplicator(TimeSpan.FromSeconds(1));
        var now = DateTimeOffset.UtcNow;
        var formulaId = Guid.NewGuid().ToString();

        Assert.True(deduplicator.ShouldAccept(
            "powerpoint",
            "deck-a",
            $"VisualTeX_{formulaId}",
            formulaId,
            now));
        Assert.False(deduplicator.ShouldAccept(
            "powerpoint",
            "deck-a",
            $"VisualTeX_{formulaId}",
            formulaId,
            now.AddMilliseconds(250)));
        Assert.True(deduplicator.ShouldAccept(
            "powerpoint",
            "deck-b",
            $"VisualTeX_{formulaId}",
            formulaId,
            now.AddMilliseconds(300)));
        Assert.True(deduplicator.ShouldAccept(
            "powerpoint",
            "deck-a",
            $"VisualTeX_{formulaId}",
            formulaId,
            now.AddSeconds(2)));
        Assert.False(deduplicator.ShouldAccept(
            "word",
            "document-a",
            null,
            null,
            now));
    }
}
