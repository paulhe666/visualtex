using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class OfficePluginLanguageTests
{
    private static readonly object LanguageLock = new();

    private static T InEnglish<T>(Func<T> callback)
    {
        lock (LanguageLock)
        {
            var previous = Environment.GetEnvironmentVariable("VISUALTEX_OFFICE_LANGUAGE");
            try
            {
                Environment.SetEnvironmentVariable("VISUALTEX_OFFICE_LANGUAGE", "en");
                return callback();
            }
            finally
            {
                Environment.SetEnvironmentVariable("VISUALTEX_OFFICE_LANGUAGE", previous);
            }
        }
    }

    [Fact]
    public void SafeErrorMessage_TranslatesKnownOfficeWriteErrorWithoutDroppingDetail()
    {
        var message = InEnglish(() =>
            OfficePluginLanguage.SafeErrorMessage(
                "VisualTeX Word 写入失败：InsertXML rejected the OMML payload at m:oMathPara."));

        Assert.Equal(
            "VisualTeX Word formula write failed: InsertXML rejected the OMML payload at m:oMathPara.",
            message);
        Assert.DoesNotContain(
            "See the application log for technical details",
            message,
            StringComparison.Ordinal);
    }

    [Fact]
    public void SafeErrorMessage_PreservesUnknownLocalizedHostDetail()
    {
        var message = InEnglish(() =>
            OfficePluginLanguage.SafeErrorMessage(
                "调用 Word COM 接口失败，参数无效。 HRESULT=0x800A16DC"));

        Assert.StartsWith(
            "The VisualTeX Office operation failed. Localized host detail:",
            message,
            StringComparison.Ordinal);
        Assert.Contains("HRESULT=0x800A16DC", message, StringComparison.Ordinal);
        Assert.Contains("参数无效", message, StringComparison.Ordinal);
    }

    [Fact]
    public void FormatOperationError_ProvidesStageTypeAndHResultInEnglish()
    {
        var error = new InvalidOperationException("活动 Word 文档已切换，未写入公式。");
        var message = InEnglish(() =>
            OfficePluginLanguage.FormatOperationError(
                error,
                "VisualTeX Word 写入失败",
                "VisualTeX Word formula write failed"));

        Assert.Contains(
            "VisualTeX Word formula write failed:",
            message,
            StringComparison.Ordinal);
        Assert.Contains(
            "The active Word document changed, so the formula was not written.",
            message,
            StringComparison.Ordinal);
        Assert.Contains("[InvalidOperationException, HRESULT=0x", message, StringComparison.Ordinal);
        Assert.DoesNotContain("VisualTeX Office operation failed. See", message, StringComparison.Ordinal);
    }

    [Fact]
    public void SafeStatusMessage_UsesSpecificTranslatedFailureInsteadOfGenericStatus()
    {
        var message = InEnglish(() =>
            OfficePluginLanguage.SafeStatusMessage(
                "VisualTeX PowerPoint 写入失败：The shape collection rejected the replacement."));

        Assert.Equal(
            "VisualTeX PowerPoint formula write failed: The shape collection rejected the replacement.",
            message);
        Assert.NotEqual("VisualTeX Office operation failed.", message);
    }
}
