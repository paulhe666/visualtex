using System;
using System.Globalization;
using Microsoft.Win32;

namespace VisualTeX.WindowsOffice.VstoShared;

internal static class OfficePluginLanguage
{
    private const string SharedSettingsRegistryPath = @"Software\VisualTeX\OfficeIntegration";
    private const string SharedLanguageValue = "Language";

    internal static bool IsEnglish
    {
        get
        {
            var forced = Environment.GetEnvironmentVariable("VISUALTEX_OFFICE_LANGUAGE");
            return ResolveIsEnglish(
                forced,
                ReadSharedApplicationLanguage(),
                CultureInfo.CurrentUICulture.Name);
        }
    }

    internal static bool ResolveIsEnglish(
        string? forcedLanguage,
        string? applicationLanguage,
        string? currentUiCultureName)
    {
        var forced = ParseEnglishPreference(forcedLanguage);
        if (forced.HasValue) return forced.Value;

        var application = ParseEnglishPreference(applicationLanguage);
        if (application.HasValue) return application.Value;

        return !(currentUiCultureName ?? string.Empty)
            .StartsWith("zh", StringComparison.OrdinalIgnoreCase);
    }

    private static bool? ParseEnglishPreference(string? value)
    {
        if (string.Equals(value, "en", StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(value, "cn", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "zh", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "zh-cn", StringComparison.OrdinalIgnoreCase)) return false;
        return null;
    }

    private static string? ReadSharedApplicationLanguage()
    {
        var views = Environment.Is64BitOperatingSystem
            ? new[] { RegistryView.Registry64, RegistryView.Registry32 }
            : new[] { RegistryView.Default };
        foreach (var view in views)
        {
            try
            {
                using var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
                using var key = root.OpenSubKey(SharedSettingsRegistryPath, writable: false);
                var value = key?.GetValue(
                    SharedLanguageValue,
                    null,
                    RegistryValueOptions.DoNotExpandEnvironmentNames) as string;
                if (ParseEnglishPreference(value).HasValue) return value;
            }
            catch
            {
                // A missing or inaccessible shared setting falls back to Office UI culture.
            }
        }
        return null;
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
        ("没有选中公式。请先选中一个 VisualTeX 公式、Word 原生公式或 MathType 公式。",
            "No formula is selected. Select a VisualTeX formula, a native Word equation or a MathType equation first."),
        ("没有选中公式。请先选中一个 VisualTeX 公式、PowerPoint 原生公式或 MathType 公式。",
            "No formula is selected. Select a VisualTeX formula, a native PowerPoint equation or a MathType equation first."),
        ("活动 Word 文档已切换，未写入公式。", "The active Word document changed, so the formula was not written."),
        ("活动演示文稿已切换，未写入公式。", "The active PowerPoint presentation changed, so the formula was not written."),
        ("VisualTeX Session 失败。", "The VisualTeX Session failed."),
        ("VisualTeX 操作已取消。", "The VisualTeX Office operation was cancelled."),
        ("无法为 MathType 公式编号创建唯一引用书签。", "Unable to create a unique reference bookmark for the MathType equation number."),
        ("Word 无法建立公式转 LaTeX 的撤销事务。为避免公式丢失，本次转换未开始。",
            "Word could not create the undo transaction for equation-to-LaTeX conversion. The conversion was not started to avoid losing the equation."),
        ("VisualTeX 公式转为 LaTeX 失败：", "VisualTeX equation-to-LaTeX conversion failed: "),
        ("公式格式转换失败：", "Equation format conversion failed: "),
        ("VisualTeX LaTeX 重绘失败：", "VisualTeX LaTeX redraw failed: "),
        ("VisualTeX 批量导入失败：", "VisualTeX bulk import failed: "),
        ("VisualTeX Word 插件操作失败：", "VisualTeX Word add-in operation failed: "),
        ("VisualTeX PowerPoint 插件操作失败：", "VisualTeX PowerPoint add-in operation failed: "),
        ("VisualTeX Word 双击检测失败：", "VisualTeX Word double-click detection failed: "),
        ("VisualTeX 双击检测失败：", "VisualTeX double-click detection failed: "),
        ("已有 VisualTeX 编辑任务，但状态检查失败：",
            "A VisualTeX editor task is already active, but its status check failed: "),
        ("已有编辑任务，但无法置前窗口：",
            "An editor task is already active, but VisualTeX could not bring its window to the foreground: "),
        ("无法打开 VisualTeX：", "Unable to open VisualTeX: "),
        ("无法读取当前插入位置：", "Unable to read the current insertion position: "),
        ("无法识别公式格式转换方向。",
            "Unable to determine the equation format-conversion direction."),
        ("更新 Word 公式编号失败：", "Failed to update Word equation numbers: "),
        ("设置公式编号格式失败：", "Failed to set the equation-number format: "),
        ("当前 Word 文档中没有找到可转换的公式。",
            "No convertible equations were found in the current Word document."),
        ("所选内容中没有找到可转换的公式。",
            "No convertible equations were found in the current selection."),
        ("当前 Word 文档中没有找到可转换的 Word OMML 公式。",
            "No convertible Word OMML equations were found in the current Word document."),
        ("所选内容中没有找到可转换的 Word OMML 公式。",
            "No convertible Word OMML equations were found in the current selection."),
        ("所选公式仍使用旧版编号宿主，必须先迁移为标准编号结构后再转为 LaTeX。",
            "The selected equation still uses the legacy numbering host. Migrate it to the standard numbering structure before converting it to LaTeX."),
        ("VisualTeX OLE 没有可读取的内嵌公式元数据。",
            "The VisualTeX OLE object contains no readable embedded equation metadata."),
        ("公式转 LaTeX 失败，而且 Word 无法完整恢复转换前的文档结构。请保留当前文档以便排查。",
            "Equation-to-LaTeX conversion failed, and Word could not fully restore the pre-conversion document structure. Keep the current document for diagnosis."),
        ("请只在一个文本框中插入 OMML 公式。",
            "Insert the OMML equation into only one text box at a time."),
        ("当前选择不是可编辑文本框。",
            "The current selection is not an editable text box."),
        ("PowerPoint 没有返回有效的文本插入位置。",
            "PowerPoint did not return a valid text insertion position."),
        ("请先选择一个 VisualTeX 公式或 PowerPoint 原生公式。",
            "Select a VisualTeX formula or a native PowerPoint equation first."),
        ("请先把文本光标或文本选区放到一个现有文本框中。",
            "Place the text caret or a text selection inside an existing text box first."),
        ("插入行内 OMML 公式前，请先把文本光标或文本选区放到一个现有文本框中。",
            "Before inserting an inline OMML equation, place the text caret or a text selection inside an existing text box."),
        ("VisualTeX 文档导入窗口返回失败状态。",
            "The VisualTeX document-import window returned a failed state."),
        ("批量导入没有返回可用的文档结构。",
            "Bulk import did not return a usable document structure."),
        ("批量导入文档结构不能超过 5 MB。",
            "The bulk-import document structure cannot exceed 5 MB."),
        ("批量导入返回了无效的文档结构。",
            "Bulk import returned an invalid document structure."),
        ("批量导入没有找到可以插入 Word 的内容。",
            "Bulk import found no content that can be inserted into Word."),
        ("没有找到可以插入 Word 的文字或公式。",
            "No text or equations that can be inserted into Word were found."),
        ("批量导入包含过多段落（上限 10000）。",
            "Bulk import contains too many paragraphs (maximum 10,000)."),
        ("批量导入包含过多公式（上限 1000）。",
            "Bulk import contains too many equations (maximum 1,000)."),
        ("批量导入内容不能超过 5 MB。",
            "Bulk-import source content cannot exceed 5 MB."),
        ("LaTeX 重绘范围不能超过 5 MB。",
            "The LaTeX redraw range cannot exceed 5 MB."),
        ("LaTeX 重绘包含过多公式（上限 1000）。",
            "LaTeX redraw contains too many equations (maximum 1,000)."),
        ("LaTeX 重绘缺少一个或多个已渲染目标。",
            "LaTeX redraw is missing one or more rendered targets."),
        ("批量导入的行内公式占位符发生了变化。",
            "The inline-equation placeholder changed during bulk import."),
        ("批量导入的行间公式插入点不属于唯一段落。",
            "The display-equation insertion point does not belong to exactly one paragraph during bulk import."),
        ("文档中仍存在旧版 VisualTeX 编号结构；请先执行编号结构迁移，再刷新编号。",
            "The document still contains a legacy VisualTeX numbering structure. Migrate the numbering structure before refreshing equation numbers."),
        ("公式引用未能保留一个完整的嵌套 REF 字段。",
            "The equation reference could not preserve a complete nested REF field."),
        ("公式引用在刷新后丢失了嵌套 REF 字段。",
            "The equation reference lost its nested REF field after refresh."),
        ("对象已被删除。", "The Word object has been deleted or invalidated."),
        ("此对象已被删除。", "The Word object has been deleted or invalidated."),
        ("参数无效。", "The Office host reported an invalid argument."),
        ("参数不正确。", "The Office host reported an invalid argument."),
        ("此命令不可用。", "This Office command is not available in the current context."),
        ("命令失败。", "The Office command failed."),
        ("该命令失败。", "The Office command failed."),
        ("集合所要求的成员不存在。", "The requested member of the Office collection does not exist."),
        ("请求的成员不存在。", "The requested Office member does not exist."),
        ("无法编辑范围。", "Word cannot edit the current range."),
        ("无法编辑该范围。", "Word cannot edit the current range."),
        ("无法编辑区域。", "Word cannot edit the current range."),
        ("无法编辑该区域。", "Word cannot edit the current range."),
        ("当前 Word 文档中没有找到", "The current Word document contains no "),
        ("所选内容中没有找到", "The current selection contains no "),
        ("缺少公式 ", "Equation "),
        (" 的批量渲染结果。", " is missing its batch render result."),
        (" 的渲染结果。", " is missing its render result."),
        ("公式 ", "Equation "),
        (" 的源文本范围无效。", " has an invalid source-text range."),
        ("Word 未能定位行间公式所在段落。",
            "Word could not locate the paragraph containing the display equation."),
        ("公式渲染失败：", "Formula rendering failed: "),
        ("原因：", "Reason: "),
        ("批量公式渲染已取消。", "Batch formula rendering was cancelled."),
        ("没有生成导出结果。", "did not produce an export result."),
        ("没有生成 MathML。", "did not produce MathML."),
        ("OLE 公式预览生成失败：", "OLE equation preview generation failed: "),
        ("VisualTeX 排版服务找不到 MathType 字体，为避免公式显示错误，文档尚未修改。请在 VisualTeX 设置中点击“修复 Office 集成”后重试；如仍失败，请重启电脑。",
            "VisualTeX's renderer cannot find MathType's fonts, so the document was not changed (its symbols would display wrongly). In VisualTeX settings, click “Repair integration” and try again; if it still fails, restart Windows."),
        ("缺少 MathML。", "is missing MathML."),
        ("缺少公式“", "Missing the redraw result for equation “"),
        ("”的重绘结果。", "”."),
        ("”的 MathType MathML。", "” MathType MathML."),
        ("公式引用目标书签“", "The equation-reference target bookmark “"),
        ("”已不存在。文档内容可能已发生变化。",
            "” no longer exists. The document content may have changed."),
        (" 的 LaTeX 元数据为空。为避免删除原公式，转换已中止。",
            " has empty LaTeX metadata. Conversion was stopped to avoid deleting the original equation."),
        (" 的 LaTeX 源码仍位于隐藏编号框架中。为避免源码不可见，转换已回滚。",
            " still has LaTeX source inside a hidden numbering frame. The conversion was rolled back to avoid invisible source."),
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
        if (message.IndexOf("Untranslated detail:", StringComparison.Ordinal) >= 0
            || message.IndexOf("Localized host detail:", StringComparison.Ordinal) >= 0)
            return message;

        var translated = TranslateKnownErrorText(message);
        if (!ContainsCjk(translated)) return translated;

        // Never collapse a specific failure to a generic sentence: text without a
        // translation is kept (it may come from Office or from VisualTeX itself).
        return $"{englishFallback} Untranslated detail: {translated}";
    }

    internal static string SafeStatusMessage(string? value)
    {
        var message = value ?? string.Empty;
        if (!IsEnglish || !ContainsCjk(message)) return message;

        var knownTranslation = TranslateKnownErrorText(message);
        if (!string.Equals(
                knownTranslation,
                message,
                StringComparison.Ordinal)
            && !ContainsCjk(knownTranslation))
            return knownTranslation;

        if (message.Contains("失败")
            || message.Contains("错误")
            || message.Contains("无法")
            || message.Contains("未找到")
            || message.Contains("没有找到")
            || message.Contains("无效")
            || message.Contains("只读")
            || message.Contains("缺少")
            || message.Contains("中止")
            || message.Contains("回滚")
            || message.Contains("必须先")
            || message.Contains("不存在"))
            return SafeErrorMessage(message);
        // last resort for a status without an English text (every VisualTeX
        // status has one; this keeps an unknown one from being misreported)
        if (message.Contains("取消"))
            return SafeUserMessage(message, "VisualTeX Office operation was cancelled.");
        if (message.Contains("停止"))
            return SafeUserMessage(message, "VisualTeX Office operation stopped.");
        if (message.Contains("正在") || message.Contains("请稍候"))
            return SafeUserMessage(message, "VisualTeX is processing the current Office operation…");
        if (message.Contains("完成"))
            return SafeUserMessage(message, "VisualTeX Office operation completed.");
        return SafeUserMessage(message, "VisualTeX Office status:");
    }

    internal static string SafeErrorMessage(string? value) =>
        SafeUserMessage(
            value,
            "The VisualTeX Office operation failed.");

    /// <summary>
    /// "{operation}: {specific reason}" in the VisualTeX language: timeouts,
    /// unreachable VisualTeX, refused requests, failed render batches, COM codes
    /// and every cause of an aggregated failure are worded by OfficeErrorText.
    /// </summary>
    internal static string FormatOperationError(
        Exception error,
        string chineseOperation,
        string englishOperation) =>
        OfficeErrorText.Operation(error, chineseOperation, englishOperation);

    internal static string UiFontFamily => IsEnglish ? "Segoe UI" : "Microsoft YaHei UI";
}
