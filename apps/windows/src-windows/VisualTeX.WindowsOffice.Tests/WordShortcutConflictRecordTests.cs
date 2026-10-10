using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordShortcutConflictRecordTests
{
    [Theory]
    [InlineData("Ctrl+D")]
    [InlineData("Ctrl+Shift+7")]
    [InlineData("Ctrl+Alt+Shift+F16")]
    [InlineData("Alt+Shift+F24")]
    [InlineData("Ctrl+Alt+Q")]
    public void FormatMatchesTheSettingsText(string text)
    {
        Assert.True(OfficeFormulaShortcut.TryParse(text, out var shortcut));
        Assert.Equal(text, shortcut.Format());
    }

    [Fact]
    public void MergeReplacesAndRemovesOneChord()
    {
        var value = WordShortcutConflictRecord.Merge(null, "Ctrl+D", "1:FormatFont");
        Assert.Equal("Ctrl+D=1:FormatFont", value);
        value = WordShortcutConflictRecord.Merge(value, "Ctrl+Alt+C", "6:");
        value = WordShortcutConflictRecord.Merge(value, "Ctrl+D", "2:MTCommand_InsertInlineEqn");
        Assert.Equal("Ctrl+Alt+C=6:;Ctrl+D=2:MTCommand_InsertInlineEqn", value);
        value = WordShortcutConflictRecord.Merge(value, "Ctrl+Alt+C", null);
        Assert.Equal("Ctrl+D=2:MTCommand_InsertInlineEqn", value);
        Assert.Equal(string.Empty, WordShortcutConflictRecord.Merge(value, "Ctrl+D", null));
    }

    [Fact]
    public void MergeDoesNotTreatAPrefixChordAsTheSameChord()
    {
        var value = WordShortcutConflictRecord.Merge("Ctrl+D=1:FormatFont;Ctrl+D9=1:X", "Ctrl+D", null);
        Assert.Equal("Ctrl+D9=1:X", value);
    }

    [Theory]
    [InlineData("FormatFont", "FormatFont")]
    [InlineData("Normal.NewMacros.Insert_Eq", "Normal.NewMacros.Insert_Eq")]
    [InlineData("©", "")]
    [InlineData("标题 1", "")]
    [InlineData("Macro公式", "")]
    [InlineData(null, "")]
    public void OnlyWholeAsciiNamesAreRecorded(string? name, string expected) =>
        Assert.Equal(expected, WordShortcutConflictRecord.AsciiName(name));
}
