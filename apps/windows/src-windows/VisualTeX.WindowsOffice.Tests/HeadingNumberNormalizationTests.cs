using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class HeadingNumberNormalizationTests
{
    [Theory]
    [InlineData("3", 1, "3")]
    [InlineData("3.", 1, "3")]
    [InlineData("3、", 1, "3")]
    [InlineData("（3）", 1, "3")]
    [InlineData("第3章", 1, "3")]
    [InlineData("第 3 章", 1, "3")]
    [InlineData("第３章", 1, "3")]
    [InlineData("Chapter 3", 1, "3")]
    [InlineData("chapter 3:", 1, "3")]
    [InlineData("3-2", 2, "3.2")]
    [InlineData("3．2", 2, "3.2")]
    [InlineData("3—2", 2, "3.2")]
    [InlineData("第3章第2节", 2, "3.2")]
    [InlineData("第3章 第2节", 2, "3.2")]
    [InlineData("第３章　第２节", 2, "3.2")]
    [InlineData("Chapter 3 Section 2", 2, "3.2")]
    public void WordListDisplayTextIsNormalizedToNumericHierarchy(
        string value,
        int outlineLevel,
        string expected)
    {
        var actual = WordEquationNumbering.NormalizeHeadingNumberText(
            value,
            outlineLevel);

        Assert.Equal(expected, actual);
        Assert.DoesNotContain("第", actual, StringComparison.Ordinal);
        Assert.DoesNotContain("章", actual, StringComparison.Ordinal);
        Assert.DoesNotContain("节", actual, StringComparison.Ordinal);
        Assert.DoesNotContain("Chapter", actual, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("", 1)]
    [InlineData("第X章", 1)]
    [InlineData("Chapter Three", 1)]
    [InlineData("ISO 9001", 1)]
    [InlineData("第3章 第2节", 1)]
    [InlineData("3.2.1", 2)]
    [InlineData("第3章（上）", 1)]
    public void AmbiguousOrOverdeepListTextFailsClosed(
        string value,
        int outlineLevel)
    {
        Assert.Equal(
            string.Empty,
            WordEquationNumbering.NormalizeHeadingNumberText(value, outlineLevel));
    }
}
