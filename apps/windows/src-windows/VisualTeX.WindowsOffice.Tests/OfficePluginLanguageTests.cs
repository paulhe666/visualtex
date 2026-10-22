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

    [Theory]
    [InlineData(null, "en", "zh-CN", true)]
    [InlineData(null, "cn", "en-US", false)]
    [InlineData("en", "cn", "zh-CN", true)]
    [InlineData("cn", "en", "en-US", false)]
    [InlineData(null, null, "zh-TW", false)]
    [InlineData(null, null, "en-US", true)]
    public void ResolveIsEnglish_PrefersVisualTeXLanguageBeforeOfficeCulture(
        string? forcedLanguage,
        string? applicationLanguage,
        string currentUiCultureName,
        bool expected)
    {
        Assert.Equal(
            expected,
            OfficePluginLanguage.ResolveIsEnglish(
                forcedLanguage,
                applicationLanguage,
                currentUiCultureName));
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
            "The VisualTeX Office operation failed. Untranslated detail:",
            message,
            StringComparison.Ordinal);
        Assert.Contains("HRESULT=0x800A16DC", message, StringComparison.Ordinal);
        Assert.Contains("调用 Word COM 接口失败", message, StringComparison.Ordinal);
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
        // the reason, not the exception type, is what the user reads
        Assert.DoesNotContain("[InvalidOperationException", message, StringComparison.Ordinal);
        Assert.DoesNotContain("VisualTeX Office operation failed. See", message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(
        "当前 Word 文档中没有找到可转换的公式。",
        "No convertible equations were found in the current Word document.")]
    [InlineData(
        "所选公式仍使用旧版编号宿主，必须先迁移为标准编号结构后再转为 LaTeX。",
        "The selected equation still uses the legacy numbering host. Migrate it to the standard numbering structure before converting it to LaTeX.")]
    [InlineData(
        "PowerPoint 没有返回有效的文本插入位置。",
        "PowerPoint did not return a valid text insertion position.")]
    public void SafeErrorMessage_MapsConcreteOfficeErrorsToEnglish(
        string localized,
        string expected)
    {
        var message = InEnglish(() =>
            OfficePluginLanguage.SafeErrorMessage(localized));

        Assert.Equal(expected, message);
        Assert.False(
            OfficePluginLanguage.ContainsCjk(message),
            $"Known Office error still contains CJK text: {message}");
        Assert.DoesNotContain(
            "See the application log",
            message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FormatOperationError_EquationToLatexKeepsConcreteReasonAndDiagnostics()
    {
        var error = new InvalidDataException(
            "当前 Word 文档中没有找到可转换的公式。");
        var message = InEnglish(() =>
            OfficePluginLanguage.FormatOperationError(
                error,
                "VisualTeX 公式转为 LaTeX 失败",
                "VisualTeX equation-to-LaTeX conversion failed"));

        Assert.Contains(
            "VisualTeX equation-to-LaTeX conversion failed:",
            message,
            StringComparison.Ordinal);
        Assert.Contains(
            "No convertible equations were found in the current Word document.",
            message,
            StringComparison.Ordinal);
        Assert.False(
            OfficePluginLanguage.ContainsCjk(message),
            $"Known equation-to-LaTeX error still contains CJK text: {message}");
    }

    [Fact]
    public void FormatOperationError_UsesKnownComHResultWhenHostTextIsLocalized()
    {
        var error = new System.Runtime.InteropServices.COMException(
            "对象已被删除。",
            unchecked((int)0x800A16C1));
        var message = InEnglish(() =>
            OfficePluginLanguage.FormatOperationError(
                error,
                "VisualTeX 公式转为 LaTeX 失败",
                "VisualTeX equation-to-LaTeX conversion failed"));

        Assert.Contains(
            "The Word object has been deleted or invalidated.",
            message,
            StringComparison.Ordinal);
        Assert.Contains(
            "HRESULT=0x800A16C1",
            message,
            StringComparison.Ordinal);
        Assert.False(
            OfficePluginLanguage.ContainsCjk(message),
            $"Known COM error still contains CJK text: {message}");
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

    private static T InLanguage<T>(string language, Func<T> callback)
    {
        lock (LanguageLock)
        {
            var previous = Environment.GetEnvironmentVariable("VISUALTEX_OFFICE_LANGUAGE");
            try
            {
                Environment.SetEnvironmentVariable("VISUALTEX_OFFICE_LANGUAGE", language);
                return callback();
            }
            finally
            {
                Environment.SetEnvironmentVariable("VISUALTEX_OFFICE_LANGUAGE", previous);
            }
        }
    }

    private static T Tagged<T>(T error, string failure, string operation) where T : Exception
    {
        error.Data[VisualTeX.WindowsOffice.Contracts.VisualTeXServiceErrors.FailureKey] = failure;
        error.Data[VisualTeX.WindowsOffice.Contracts.VisualTeXServiceErrors.OperationKey] = operation;
        return error;
    }

    [Fact]
    public void OfficeErrorText_RequestTimeoutNamesRequestAndLimitInBothLanguages()
    {
        var error = Tagged(new TimeoutException("x"),
            VisualTeX.WindowsOffice.Contracts.VisualTeXServiceErrors.FailureRequestTimedOut,
            "POST /api/v1/app/converter/render-batches");
        error.Data[VisualTeX.WindowsOffice.Contracts.VisualTeXServiceErrors.LimitSecondsKey] = 30.0;
        var english = InLanguage("en", () => OfficeErrorText.Describe(error));
        var chinese = InLanguage("cn", () => OfficeErrorText.Describe(error));
        Assert.Contains("within 30 s", english, StringComparison.Ordinal);
        Assert.Contains("POST /api/v1/app/converter/render-batches", english, StringComparison.Ordinal);
        Assert.False(OfficePluginLanguage.ContainsCjk(english), english);
        Assert.Contains("30 秒", chinese, StringComparison.Ordinal);
        Assert.DoesNotContain("cancel", english, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OfficeErrorText_BatchFailureCodesHaveTheirOwnSentences()
    {
        string Describe(string code)
        {
            var error = Tagged(new InvalidOperationException("VisualTeX render batch failed: detail"),
                VisualTeX.WindowsOffice.Contracts.VisualTeXServiceErrors.FailureBatchFailed, "render batch");
            error.Data[VisualTeX.WindowsOffice.Contracts.VisualTeXServiceErrors.BatchErrorCodeKey] = code;
            error.Data[VisualTeX.WindowsOffice.Contracts.VisualTeXServiceErrors.ServerErrorKey] = "detail";
            error.Data[VisualTeX.WindowsOffice.Contracts.VisualTeXServiceErrors.FormulaCountKey] = 3;
            return InLanguage("en", () => OfficeErrorText.Describe(error));
        }
        var texts = new[] { Describe("converter-not-started"), Describe("converter-failed"), Describe("cancelled") };
        Assert.Equal(3, texts.Distinct().Count());
        Assert.Contains("did not start", texts[0], StringComparison.Ordinal);
        Assert.Contains("detail", texts[1], StringComparison.Ordinal);
        Assert.Contains("released", texts[2], StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeErrorText_AggregateKeepsEveryCause()
    {
        var error = new AggregateException("The formula mutation failed and could not be rolled back.",
            new InvalidOperationException("InsertXML rejected the payload."),
            new System.Runtime.InteropServices.COMException("x", unchecked((int)0x800A16C1)));
        var message = InLanguage("en", () => OfficeErrorText.Describe(error));
        Assert.Contains("InsertXML rejected the payload.", message, StringComparison.Ordinal);
        Assert.Contains("HRESULT=0x800A16C1", message, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeErrorText_ItemResultChecksAreTranslated()
    {
        var chinese = InLanguage("cn", () => OfficeErrorText.DescribeItemError("VisualTeX's result for this formula has no MathML."));
        Assert.True(OfficePluginLanguage.ContainsCjk(chinese));
        var empty = InLanguage("en", () => OfficeErrorText.DescribeItemError("[object Object]"));
        Assert.Contains("gave no reason", empty, StringComparison.Ordinal);
    }
}
