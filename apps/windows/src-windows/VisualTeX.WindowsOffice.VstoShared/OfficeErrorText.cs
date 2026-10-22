using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.VstoShared;

/// <summary>
/// The specific reason for a failure, worded for the user in the VisualTeX
/// language: every cause its own sentence (a timeout names the request and the
/// limit, a rejected request VisualTeX's reason, a batch the converter never
/// started says so, a COM error keeps its code, an aggregated failure all its
/// causes). Callers add the operation ("Equation format conversion failed: …").
/// </summary>
internal static class OfficeErrorText
{
    private static string T(string chinese, string english) => OfficePluginLanguage.Text(chinese, english);

    private static string Seconds(object? value) =>
        value is double seconds ? seconds.ToString("0.#", CultureInfo.InvariantCulture) : "?";

    /// <summary>The reason for <paramref name="error"/>, without an operation prefix.</summary>
    internal static string Describe(Exception? error)
    {
        if (error is null) return T("未知错误（没有异常信息）。", "Unknown error (no exception information).");
        if (error is AggregateException aggregate)
        {
            // every distinct cause, the primary one first (a failed rollback or
            // release must not hide what actually went wrong)
            var causes = aggregate.Flatten().InnerExceptions
                .Select(Describe)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .Distinct(StringComparer.Ordinal)
                .ToList();
            if (causes.Count == 0) return Plain(aggregate);
            var headline = OwnMessage(aggregate);
            var list = string.Join(T("；", "; "), causes);
            return string.IsNullOrWhiteSpace(headline) || causes.Any(c => c.Contains(headline))
                ? list
                : headline + T("（原因：", " (causes: ") + list + T("）", ")");
        }
        var service = DescribeService(error);
        if (service != null) return service;
        if (error is OperationCanceledException)
            return T("操作已被取消。", "The operation was cancelled.");
        if (error is COMException com) return DescribeCom(com);
        return Plain(error);
    }

    /// <summary>"{operation}：{reason}" (or "{operation}: {reason}").</summary>
    internal static string Operation(Exception error, string chineseOperation, string englishOperation) =>
        OfficePluginLanguage.IsEnglish
            ? $"{englishOperation}: {Describe(error)}"
            : $"{chineseOperation}：{Describe(error)}";

    /// <summary>
    /// One formula's failure as a render batch item reports it (the converter
    /// words its own reasons in the VisualTeX language; VisualTeX's result checks
    /// are fixed sentences translated here).
    /// </summary>
    internal static string DescribeItemError(string? error)
    {
        var text = (error ?? string.Empty).Trim();
        if (text.Length == 0 || text == "[object Object]")
            return T("VisualTeX 没有说明该公式失败的原因（转换窗口返回了空的错误信息）。",
                "VisualTeX gave no reason for this formula's failure (the conversion window returned an empty error).");
        switch (text)
        {
            case "VisualTeX's result for this formula has no MathML.":
                return T("VisualTeX 为该公式生成的结果缺少 MathML。", text);
            case "VisualTeX's result for this formula has an invalid size (width, height or baseline).":
                return T("VisualTeX 为该公式生成的结果尺寸无效（宽度、高度或基线）。", text);
            case "VisualTeX's result for this formula has no SVG or PNG image.":
                return T("VisualTeX 为该公式生成的结果缺少 SVG 或 PNG 图片。", text);
        }
        return OfficePluginLanguage.SafeErrorMessage(text);
    }

    private static string? DescribeService(Exception error)
    {
        if (VisualTeXServiceErrors.Get(error, VisualTeXServiceErrors.FailureKey) is not string failure)
            return null;
        var operation = VisualTeXServiceErrors.Get(error, VisualTeXServiceErrors.OperationKey) as string ?? "?";
        var serverError = VisualTeXServiceErrors.Get(error, VisualTeXServiceErrors.ServerErrorKey) as string;
        var seconds = Seconds(VisualTeXServiceErrors.Get(error, VisualTeXServiceErrors.LimitSecondsKey));
        var count = VisualTeXServiceErrors.Get(error, VisualTeXServiceErrors.FormulaCountKey) is int n ? n : 0;
        switch (failure)
        {
            case VisualTeXServiceErrors.FailureRequestTimedOut:
                return T(
                    $"VisualTeX 没有在 {seconds} 秒内响应请求（{operation}）。请确认 VisualTeX 正在运行；如反复出现，请重启 VisualTeX 后重试。",
                    $"VisualTeX did not answer within {seconds} s ({operation}). Make sure VisualTeX is running; if this keeps happening, restart VisualTeX and try again.");
            case VisualTeXServiceErrors.FailureUnreachable:
                return T(
                    $"无法连接 VisualTeX（{operation}）：{VisualTeXServiceErrors.InnermostMessage(error)}。请确认 VisualTeX 正在运行。",
                    $"VisualTeX could not be reached ({operation}): {VisualTeXServiceErrors.InnermostMessage(error)}. Make sure VisualTeX is running.");
            case VisualTeXServiceErrors.FailureHttpStatus:
            {
                var status = VisualTeXServiceErrors.Get(error, VisualTeXServiceErrors.StatusCodeKey) is int code ? code : 0;
                var reason = serverError ?? T("没有给出原因", "no reason was given");
                return T(
                    $"VisualTeX 拒绝了请求（{operation}，HTTP {status}）：{reason}",
                    $"VisualTeX refused the request ({operation}, HTTP {status}): {reason}");
            }
            case VisualTeXServiceErrors.FailureBatchTimedOut:
                return T(
                    $"VisualTeX 在 {seconds} 秒内没有完成 {count} 个公式的渲染。转换窗口可能已停止响应；请重启 VisualTeX 后重试。",
                    $"VisualTeX did not finish rendering {count} formula(s) within {seconds} s. Its conversion window may have stopped responding; restart VisualTeX and try again.");
            case VisualTeXServiceErrors.FailureBatchFailed:
            {
                var code = VisualTeXServiceErrors.Get(error, VisualTeXServiceErrors.BatchErrorCodeKey) as string;
                var detail = serverError ?? error.Message;
                return code switch
                {
                    "converter-not-started" => T(
                        $"VisualTeX 的转换窗口没有开始渲染这 {count} 个公式（{detail}）。请重启 VisualTeX 后重试。",
                        $"VisualTeX's conversion window did not start rendering these {count} formula(s) ({detail}). Restart VisualTeX and try again."),
                    "converter-failed" => T(
                        $"VisualTeX 的转换窗口无法完成这 {count} 个公式：{detail}",
                        $"VisualTeX's conversion window could not finish these {count} formula(s): {detail}"),
                    "cancelled" => T(
                        "VisualTeX 在渲染完成前释放了这批公式（转换任务被取消）。",
                        "VisualTeX released this batch of formulas before rendering finished (the conversion job was cancelled)."),
                    _ => T($"VisualTeX 渲染这批公式失败：{detail}", $"VisualTeX failed to render this batch of formulas: {detail}"),
                };
            }
        }
        return null;
    }

    private static string DescribeCom(COMException error)
    {
        var code = $"0x{error.HResult:X8}";
        var message = OwnMessage(error);
        var known = unchecked((uint)error.HResult) switch
        {
            0x800A16C1u => T("该 Word 对象已被删除或失效。", "The Word object has been deleted or invalidated."),
            0x80010001u => T("Office 正忙（拒绝了调用），请关闭 Office 中打开的对话框后重试。", "Office is busy and rejected the call; close any open Office dialog and try again."),
            0x8001010Au => T("Office 正忙（稍后重试）。", "Office is busy (retry later)."),
            0x800706BAu => T("与 Office 的连接已中断（RPC 服务器不可用）。", "The connection to Office was lost (the RPC server is unavailable)."),
            0x80004005u => null,
            _ => null,
        };
        var text = known ?? (string.IsNullOrWhiteSpace(message) ? T("Office 返回了错误。", "Office returned an error.") : OfficePluginLanguage.SafeErrorMessage(message));
        return T($"{text}（Office 错误代码 HRESULT={code}）", $"{text} (Office error HRESULT={code})");
    }

    /// <summary>The message, translated where known, with an inner cause it does not already contain.</summary>
    private static string Plain(Exception error)
    {
        var message = OwnMessage(error);
        var text = string.IsNullOrWhiteSpace(message)
            ? T($"发生了 {error.GetType().Name} 错误，但没有错误信息。", $"A {error.GetType().Name} occurred without a message.")
            : OfficePluginLanguage.SafeErrorMessage(message);
        var inner = error.InnerException;
        if (inner != null && !(inner is OperationCanceledException && error is TimeoutException))
        {
            var cause = Describe(inner);
            if (!string.IsNullOrWhiteSpace(cause) && !text.Contains(cause) && !(message ?? "").Contains(inner.Message))
                text += T("（原因：", " (cause: ") + cause + T("）", ")");
        }
        return text;
    }

    private static string OwnMessage(Exception error) => (error.Message ?? string.Empty).Trim();
}
