using System;
using VisualTeX.WindowsOffice.VstoShared;

internal static class OfficePluginLanguageSmoke
{
    private static void Require(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    public static int Main()
    {
        Require(
            OfficePluginLanguage.ResolveIsEnglish(null, "en", "zh-CN"),
            "VisualTeX English mode did not override Chinese Office UI culture.");
        Require(
            !OfficePluginLanguage.ResolveIsEnglish(null, "cn", "en-US"),
            "VisualTeX Chinese mode did not override English Office UI culture.");
        Require(
            OfficePluginLanguage.ResolveIsEnglish("en", "cn", "zh-CN"),
            "The explicit Office language test override lost precedence.");

        var previous = Environment.GetEnvironmentVariable("VISUALTEX_OFFICE_LANGUAGE");
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_OFFICE_LANGUAGE", "en");

            var translated = OfficePluginLanguage.SafeErrorMessage(
                "VisualTeX Word 写入失败：InsertXML rejected the OMML payload at m:oMathPara.");
            Require(
                translated == "VisualTeX Word formula write failed: InsertXML rejected the OMML payload at m:oMathPara.",
                "Known Word write error was not translated while preserving technical detail: " + translated);

            var localized = OfficePluginLanguage.SafeErrorMessage(
                "调用 Word COM 接口失败，参数无效。 HRESULT=0x800A16DC");
            Require(
                localized.StartsWith("The VisualTeX Office operation failed. Localized host detail:", StringComparison.Ordinal),
                "Unknown localized Office error lost the English diagnostic prefix: " + localized);
            Require(localized.Contains("HRESULT=0x800A16DC"), "HRESULT detail was lost: " + localized);
            Require(localized.Contains("调用 Word COM 接口失败"), "Original localized COM detail was lost: " + localized);
            Require(
                OfficePluginLanguage.SafeStatusMessage(localized) == localized,
                "Already-formatted localized diagnostics must not be wrapped again: " + localized);

            var operation = OfficePluginLanguage.FormatOperationError(
                new InvalidOperationException("活动 Word 文档已切换，未写入公式。"),
                "VisualTeX Word 写入失败",
                "VisualTeX Word formula write failed");
            Require(
                operation.Contains("The active Word document changed, so the formula was not written."),
                "Operation reason was not translated: " + operation);
            Require(
                operation.Contains("[InvalidOperationException, HRESULT=0x"),
                "Exception type/HRESULT were not preserved: " + operation);

            var status = OfficePluginLanguage.SafeStatusMessage(
                "VisualTeX PowerPoint 写入失败：The shape collection rejected the replacement.");
            Require(
                status == "VisualTeX PowerPoint formula write failed: The shape collection rejected the replacement.",
                "Specific PowerPoint error collapsed to a generic status: " + status);

            Require(
                translated.IndexOf("See the application log for technical details", StringComparison.Ordinal) < 0,
                "Legacy generic Office error fallback is still exposed.");

            Console.WriteLine("VisualTeX Office English error compatibility smoke passed.");
            return 0;
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_OFFICE_LANGUAGE", previous);
        }
    }
}
