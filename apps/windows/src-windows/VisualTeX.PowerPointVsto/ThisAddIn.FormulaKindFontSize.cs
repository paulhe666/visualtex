using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Office = Microsoft.Office.Core;
using Task = System.Threading.Tasks.Task;

namespace VisualTeX.PowerPointVsto;

// Ribbon 行内 / 行间 (as in Word and on the Mac): one size for every inline or every display formula of the
// presentation, kept for the formulas LaTeX redraw and format conversion write later (PowerPointFormulaFontSizeSetting).
public sealed partial class ThisAddIn
{
    private static readonly string[] FormulaKindFontSizeControlIds =
    {
        "VisualTeX.PowerPointVsto.InlineFontSize",
        "VisualTeX.PowerPointVsto.DisplayFontSize",
    };

    private int _formulaKindFontSizeRunning;

    private static bool IsDisplayFontSizeControl(Office.IRibbonControl? control) =>
        string.Equals(control?.Id, "VisualTeX.PowerPointVsto.DisplayFontSize", StringComparison.Ordinal);

    public string GetFormulaKindFontSizeText(Office.IRibbonControl control)
    {
        try
        {
            var size = PowerPointFormulaFontSizeSetting.Get(IsDisplayFontSizeControl(control));
            return size.HasValue
                ? FormulaFontSize.FormatDisplay(size.Value, OfficePluginLanguage.IsEnglish)
                : T("跟随正文", "Follow Text");
        }
        catch { return string.Empty; }
    }

    public bool GetFormulaKindFontSizeEnabled(Office.IRibbonControl control) =>
        Volatile.Read(ref _formulaKindFontSizeRunning) == 0;

    public void OnFormulaKindFontSizeChanged(Office.IRibbonControl control, string value)
    {
        var display = IsDisplayFontSizeControl(control);
        var text = (value ?? string.Empty).Replace("(CN preset)", string.Empty).Trim();
        if (text.Length == 0
            || text == "跟随正文"
            || string.Equals(text, "Follow Text", StringComparison.OrdinalIgnoreCase))
        {
            PowerPointFormulaFontSizeSetting.Save(display, null);
            InvalidateFormulaKindFontSizeControls();
            SetStatus(display
                ? T("行间公式字号：跟随正文。", "Display formulas: follow the text.")
                : T("行内公式字号：跟随正文。", "Inline formulas: follow the text."));
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
        PowerPointFormulaFontSizeSetting.Save(display, size);
        InvalidateFormulaKindFontSizeControls();
        _ = ApplyFormulaKindFontSizeAsync(display, size);
    }

    private async Task ApplyFormulaKindFontSizeAsync(bool display, float size)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        if (!await _operationGate.WaitAsync(TimeSpan.FromSeconds(2)).ConfigureAwait(false))
        {
            SetStatus(T(
                "VisualTeX 正在执行其他操作，请稍候再试。",
                "VisualTeX is busy with another operation. Please try again shortly."));
            return;
        }
        Interlocked.Exchange(ref _formulaKindFontSizeRunning, 1);
        InvalidateFormulaKindFontSizeControls();
        var kind = display ? T("行间公式", "display formula(s)") : T("行内公式", "inline formula(s)");
        try
        {
            var result = await dispatcher.InvokeAsync(
                    () => service.ApplyFontSizeToPresentationFormulas(display, size))
                .ConfigureAwait(false);
            var message = T(
                $"已把 {result.Resized} 个{kind}改为 {FormulaFontSize.Describe(size)}。",
                $"{result.Resized} {kind} set to {FormulaFontSize.Describe(size, english: true)}.");
            if (result.MathTypeLeft > 0)
                message += T(
                    $"{result.MathTypeLeft} 个 MathType 公式保持原字号。",
                    $" {result.MathTypeLeft} MathType equation(s) keep their size.");
            if (result.Failures.Count > 0)
                message += T(
                    $"{result.Failures.Count} 个未能更改：{result.Failures[0]}",
                    $" {result.Failures.Count} could not be changed: {result.Failures[0]}");
            SetStatus(message);
        }
        catch (Exception error)
        {
            ReportError(OfficePluginLanguage.FormatOperationError(
                error,
                "无法设置公式字号",
                "Unable to set formula font size"));
        }
        finally
        {
            Interlocked.Exchange(ref _formulaKindFontSizeRunning, 0);
            _operationGate.Release();
            InvalidateFormulaKindFontSizeControls();
            InvalidateFormulaFontControls();
        }
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
        catch { }
    }
}
