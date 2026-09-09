using System.Drawing;
using System.Windows.Forms;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed class LatexRedrawDialog : Form
{
    private static string T(string chinese, string english) =>
        OfficePluginLanguage.Text(chinese, english);

    private readonly CheckBox _numberDisplayFormulas = new();

    internal LatexRedrawDialog(
        bool wholeDocument,
        int formulaCount,
        int displayFormulaCount,
        string objectModeLabel,
        string equationNumberFormatDisplayName,
        bool allowNumbering)
    {
        Text = T("VisualTeX LaTeX 重绘", "VisualTeX LaTeX Redraw");
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(540, 235);
        Font = new Font(
            OfficePluginLanguage.UiFontFamily,
            9f,
            FontStyle.Regular,
            GraphicsUnit.Point);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowIcon = false;
        ShowInTaskbar = false;

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(18),
        };
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        var contentHost = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            Margin = Padding.Empty,
        };
        var content = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 1,
            RowCount = 3,
            Margin = Padding.Empty,
        };
        content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        contentHost.Controls.Add(content);
        root.Controls.Add(contentHost, 0, 0);

        var description = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(500, 0),
            Text = OfficePluginLanguage.IsEnglish
                ? wholeDocument
                    ? $"Redraw {formulaCount} LaTeX equations in place across the entire document as {objectModeLabel}."
                    : $"Redraw {formulaCount} LaTeX equations in place within the selection as {objectModeLabel}."
                : wholeDocument
                    ? $"将在整个文档中原位重绘 {formulaCount} 个 LaTeX 公式为 {objectModeLabel}。"
                    : $"将在所选内容中原位重绘 {formulaCount} 个 LaTeX 公式为 {objectModeLabel}。",
            Margin = new Padding(0, 0, 0, 12),
        };
        content.Controls.Add(description, 0, 0);

        if (allowNumbering)
        {
            _numberDisplayFormulas.AutoSize = true;
            _numberDisplayFormulas.Enabled = displayFormulaCount > 0;
            _numberDisplayFormulas.Text = OfficePluginLanguage.IsEnglish
                ? displayFormulaCount > 0
                    ? $"Number all {displayFormulaCount} display equations"
                    : "Number all display equations (none detected in this operation)"
                : displayFormulaCount > 0
                    ? $"为全部 {displayFormulaCount} 个行间公式添加编号"
                    : "为所有行间公式添加编号（本次未检测到行间公式）";
            _numberDisplayFormulas.Margin = new Padding(0, 0, 0, 5);
            content.Controls.Add(_numberDisplayFormulas, 0, 1);

            var detail = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(500, 0),
                Text = OfficePluginLanguage.IsEnglish
                    ? $"Number format: {equationNumberFormatDisplayName}. Content outside equation delimiters is unchanged. One Ctrl+Z undoes the entire operation."
                    : $"编号格式：{equationNumberFormatDisplayName}。正文和公式定界符以外的内容不会改变；本次操作可通过一次 Ctrl+Z 整体撤销。",
                ForeColor = Color.FromArgb(88, 88, 88),
                Margin = new Padding(22, 0, 0, 12),
            };
            content.Controls.Add(detail, 0, 2);
        }

        var actions = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
            Margin = new Padding(0, 12, 0, 0),
        };
        var cancel = new Button
        {
            AutoSize = true,
            Text = T("取消", "Cancel"),
            DialogResult = DialogResult.Cancel,
            Padding = new Padding(10, 3, 10, 3),
        };
        var redraw = new Button
        {
            AutoSize = true,
            Text = T("开始重绘", "Start Redraw"),
            DialogResult = DialogResult.OK,
            Padding = new Padding(10, 3, 10, 3),
        };
        actions.Controls.Add(cancel);
        actions.Controls.Add(redraw);
        root.Controls.Add(actions, 0, 1);

        AcceptButton = redraw;
        CancelButton = cancel;
        Shown += (_, _) => FitToContent(root, content, actions);
    }

    internal bool NumberDisplayFormulas =>
        _numberDisplayFormulas.Enabled && _numberDisplayFormulas.Checked;

    private void FitToContent(
        TableLayoutPanel root,
        TableLayoutPanel content,
        FlowLayoutPanel actions)
    {
        SuspendLayout();
        try
        {
            root.PerformLayout();
            content.PerformLayout();
            actions.PerformLayout();

            var availableContentWidth = Math.Max(
                1,
                ClientSize.Width - root.Padding.Horizontal);
            var contentHeight = content.GetPreferredSize(
                new Size(availableContentWidth, 0)).Height;
            var actionsHeight = actions.GetPreferredSize(
                new Size(availableContentWidth, 0)).Height + actions.Margin.Vertical;
            var desiredClientHeight = root.Padding.Vertical
                + contentHeight
                + actionsHeight;

            var workingArea = Screen.FromControl(this).WorkingArea;
            var nonClientHeight = Math.Max(0, Height - ClientSize.Height);
            const int screenMargin = 48;
            var maxClientHeight = Math.Max(
                235,
                workingArea.Height - nonClientHeight - screenMargin);
            var targetClientHeight = Math.Min(
                Math.Max(235, desiredClientHeight),
                maxClientHeight);

            ClientSize = new Size(ClientSize.Width, targetClientHeight);
            root.PerformLayout();
        }
        finally
        {
            ResumeLayout(performLayout: true);
        }
    }
}
