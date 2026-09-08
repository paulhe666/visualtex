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
        foreach (var character in value)
        {
            if ((character >= '\u2E80' && character <= '\u9FFF')
                || (character >= '\uF900' && character <= '\uFAFF')
                || (character >= '\uFF00' && character <= '\uFFEF'))
                return true;
        }
        return false;
    }

    internal static string SafeUserMessage(string? value, string englishFallback)
    {
        var message = value ?? string.Empty;
        if (!IsEnglish || !ContainsCjk(message)) return message;
        return englishFallback;
    }

    internal static string SafeStatusMessage(string? value)
    {
        var message = value ?? string.Empty;
        if (!IsEnglish || !ContainsCjk(message)) return message;
        if (message.Contains("失败") || message.Contains("错误") || message.Contains("无法"))
            return "VisualTeX Office operation failed.";
        if (message.Contains("取消"))
            return "VisualTeX Office operation was cancelled.";
        if (message.Contains("正在") || message.Contains("请稍候"))
            return "VisualTeX is processing the current Office operation…";
        if (message.Contains("已") || message.Contains("完成"))
            return "VisualTeX Office operation completed.";
        return "VisualTeX Office status updated.";
    }

    internal static string SafeErrorMessage(string? value) =>
        SafeUserMessage(
            value,
            "The VisualTeX Office operation failed. See the application log for technical details.");

    internal static string UiFontFamily => IsEnglish ? "Segoe UI" : "Microsoft YaHei UI";
}
