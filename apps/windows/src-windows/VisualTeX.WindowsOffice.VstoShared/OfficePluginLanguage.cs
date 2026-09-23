using System;
using System.Globalization;

namespace VisualTeX.WindowsOffice.VstoShared;

internal static class OfficePluginLanguage
{
    internal static bool IsEnglish
    {
        get
        {
            var forced = Environment.GetEnvironmentVariable("VISUALTEX_OFFICE_LANGUAGE");
            if (string.Equals(forced, "en", StringComparison.OrdinalIgnoreCase)) return true;
            if (string.Equals(forced, "cn", StringComparison.OrdinalIgnoreCase)
                || string.Equals(forced, "zh", StringComparison.OrdinalIgnoreCase)
                || string.Equals(forced, "zh-cn", StringComparison.OrdinalIgnoreCase)) return false;
            var name = CultureInfo.CurrentUICulture.Name;
            return !name.StartsWith("zh", StringComparison.OrdinalIgnoreCase);
        }
    }

    internal static string Text(string chinese, string english) =>
        IsEnglish ? english : chinese;

    internal static bool ContainsCjk(string? value)
    {
        if (string.IsNullOrEmpty(value)) return false;
        foreach (var character in value!)
        {
            if ((character >= '\u2E80' && character <= '\u9FFF')
                || (character >= '\uF900' && character <= '\uFAFF')
                || (character >= '\uFF00' && character <= '\uFFEF'))
                return true;
        }
        return false;
    }

    private static readonly (string Chinese, string English)[] ErrorTranslations =
    {
        ("VisualTeX Word 写入失败：", "VisualTeX Word formula write failed: "),
        ("VisualTeX PowerPoint 写入失败：", "VisualTeX PowerPoint formula write failed: "),
        ("插入公式引用失败：", "Failed to insert the equation reference: "),
        ("删除 Word 公式失败：", "Failed to delete the Word formula: "),
        ("删除 PowerPoint 公式失败：", "Failed to delete the PowerPoint formula: "),
        ("无法设置公式字号：", "Unable to set the formula font size: "),
        ("当前 Word 文档为只读状态。", "The current Word document is read-only."),
        ("当前 PowerPoint 演示文稿为只读状态。", "The current PowerPoint presentation is read-only."),
        ("请先选择一个 VisualTeX 公式。", "Select a VisualTeX formula first."),
        ("活动 Word 文档已切换，未写入公式。", "The active Word document changed, so the formula was not written."),
        ("活动演示文稿已切换，未写入公式。", "The active PowerPoint presentation changed, so the formula was not written."),
        ("VisualTeX Session 失败。", "The VisualTeX Session failed."),
        ("VisualTeX 操作已取消。", "The VisualTeX Office operation was cancelled."),
        ("无法为 MathType 公式编号创建唯一引用书签。", "Unable to create a unique reference bookmark for the MathType equation number."),
        ("Word 无法建立公式转 LaTeX 的撤销事务。为避免公式丢失，本次转换未开始。",
            "Word could not create the undo transaction for equation-to-LaTeX conversion. The conversion was not started to avoid losing the equation."),
    };

    private static string TranslateKnownErrorText(string message)
    {
        var translated = message;
        foreach (var (chinese, english) in ErrorTranslations)
            translated = translated.Replace(chinese, english);
        return translated;
    }

    internal static string SafeUserMessage(string? value, string englishFallback)
    {
        var message = value ?? string.Empty;
        if (!IsEnglish || !ContainsCjk(message)) return message;
        if (message.IndexOf("Localized host detail:", StringComparison.Ordinal) >= 0)
            return message;

        var translated = TranslateKnownErrorText(message);
        if (!ContainsCjk(translated)) return translated;

        // Never collapse a specific Office/COM failure to a generic sentence.
        // Unknown localized host text is retained after an English explanation so
        // the exact diagnostic remains available to the user and support logs.
        return $"{englishFallback} Localized host detail: {translated}";
    }

    internal static string SafeStatusMessage(string? value)
    {
        var message = value ?? string.Empty;
        if (!IsEnglish || !ContainsCjk(message)) return message;
        if (message.Contains("失败") || message.Contains("错误") || message.Contains("无法"))
            return SafeErrorMessage(message);
        if (message.Contains("取消"))
            return "VisualTeX Office operation was cancelled.";
        if (message.Contains("正在") || message.Contains("请稍候"))
            return "VisualTeX is processing the current Office operation…";
        if (message.Contains("已") || message.Contains("完成"))
            return "VisualTeX Office operation completed.";
        return SafeUserMessage(message, "VisualTeX Office status updated.");
    }

    internal static string SafeErrorMessage(string? value) =>
        SafeUserMessage(
            value,
            "The VisualTeX Office operation failed.");

    internal static string FormatOperationError(
        Exception error,
        string chineseOperation,
        string englishOperation)
    {
        if (!IsEnglish)
            return $"{chineseOperation}：{error.Message}";

        var detail = SafeUserMessage(
            error.Message,
            "The Office host returned a localized error.");
        var exceptionType = error.GetType().Name;
        var hresult = $"0x{error.HResult:X8}";
        return $"{englishOperation}: {detail} [{exceptionType}, HRESULT={hresult}]";
    }

    internal static string UiFontFamily => IsEnglish ? "Segoe UI" : "Microsoft YaHei UI";
}
