using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Office = Microsoft.Office.Core;
using Task = System.Threading.Tasks.Task;

namespace VisualTeX.WordVsto;

// Ribbon 行内 / 行间 (as on the Mac): one size for every inline or every display formula of the document,
// kept for the formulas LaTeX redraw, bulk import and format conversion write later (WordFormulaFontSizeSetting).
public sealed partial class ThisAddIn
{
    private static readonly string[] FormulaKindFontSizeControlIds =
    {
        "VisualTeX.WordVsto.InlineFontSize",
        "VisualTeX.WordVsto.DisplayFontSize",
    };

    private static bool IsDisplayFontSizeControl(Office.IRibbonControl? control) =>
        string.Equals(control?.Id, "VisualTeX.WordVsto.DisplayFontSize", StringComparison.Ordinal);

    public string GetFormulaKindFontSizeText(Office.IRibbonControl control)
    {
        try
        {
            var size = WordFormulaFontSizeSetting.Get(IsDisplayFontSizeControl(control));
            return size.HasValue
                ? FormulaFontSize.FormatDisplay(size.Value, OfficePluginLanguage.IsEnglish)
                : T("跟随正文", "Follow Text");
        }
        catch { return string.Empty; }
    }

    public bool GetFormulaKindFontSizeEnabled(Office.IRibbonControl control) =>
        Volatile.Read(ref _formulaFormatMutationDepth) == 0;

    public void OnFormulaKindFontSizeChanged(Office.IRibbonControl control, string value)
    {
        var display = IsDisplayFontSizeControl(control);
        var text = (value ?? string.Empty).Replace("(CN preset)", string.Empty).Trim();
        if (text.Length == 0
            || text == "跟随正文"
            || string.Equals(text, "Follow Text", StringComparison.OrdinalIgnoreCase))
        {
            // every formula of the kind takes the text size at its place, and new ones follow the text
            WordFormulaFontSizeSetting.Save(display, null);
            InvalidateFormulaKindFontSizeControls();
            _ = ApplyFormulaKindFontSizeAsync(display, null);
            return;
        }
        if (!FormulaFontSize.TryParse(text, out var size))
        {
            InvalidateFormulaKindFontSizeControls();
            SetStatus(T(
                "请输入 5–200 之间的公式字号，或初号、小初等中文字号。",
                "Enter a formula font size between 5 and 200 pt."));
            return;
        }
        WordFormulaFontSizeSetting.Save(display, size);
        InvalidateFormulaKindFontSizeControls();
        _ = ApplyFormulaKindFontSizeAsync(display, size);
    }

    /// <paramref name="size"/> null: every formula of the kind follows the text at its place.
    private async Task ApplyFormulaKindFontSizeAsync(bool display, float? size)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        var lifetime = _lifetime;
        if (dispatcher is null
            || service is null
            || lifetime is null
            || lifetime.IsCancellationRequested)
            return;
        if (!await _operationGate.WaitAsync(TimeSpan.FromSeconds(2), lifetime.Token).ConfigureAwait(false))
        {
            SetStatus(T(
                "VisualTeX 正在执行其他 Word 操作，请稍候再试。",
                "VisualTeX is busy with another Word operation. Please try again shortly."));
            return;
        }
        try
        {
            BeginFormulaFormatMutation();
            var result = await dispatcher.InvokeAsync(
                    () => service.ApplyFontSizeToDocumentFormulas(display, size))
                .ConfigureAwait(false);
            var message = DescribeKindFontSizeResult(display, size, result.Resized);
            if (result.Failures.Count > 0)
                message += T(
                    $"{result.Failures.Count} 个未能更改：{result.Failures[0]}",
                    $" {result.Failures.Count} could not be changed: {result.Failures[0]}");
            WriteRedrawAcceptanceLog(
                $"formula-kind-font-size display={display} size={(size.HasValue ? size.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "follow")} resized={result.Resized} failed={result.Failures.Count} {result.MetricDetail}");
            SetStatus(message);
        }
        catch (Exception error)
        {
            SetStatus(OfficePluginLanguage.FormatOperationError(
                error,
                "无法设置公式字号",
                "Unable to set formula font size"));
            await ShowWordOperationErrorAsync(dispatcher, T("设置公式字号", "Set formula font size"), error)
                .ConfigureAwait(false);
        }
        finally
        {
            EndFormulaFormatMutation();
            _operationGate.Release();
            InvalidateFormulaFontControls();
        }
    }

    private static string DescribeKindFontSizeResult(bool display, float? size, int resized)
    {
        var kind = display ? T("行间公式", "display formula(s)") : T("行内公式", "inline formula(s)");
        return size.HasValue
            ? T($"已把 {resized} 个{kind}改为 {FormulaFontSize.Describe(size.Value)}。",
                $"{resized} {kind} set to {FormulaFontSize.Describe(size.Value, english: true)}.")
            : T($"已让 {resized} 个{kind}跟随正文字号。",
                $"{resized} {kind} now follow the text size.");
    }

    private void InvalidateFormulaKindFontSizeControls()
    {
        var ribbon = _ribbonUi;
        if (ribbon is null) return;
        try
        {
            dynamic ui = ribbon;
            foreach (var id in FormulaKindFontSizeControlIds) ui.InvalidateControl(id);
        }
        catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
    }
}
