using System.Drawing;
using System.Windows.Forms;

namespace VisualTeX.WordVsto;

internal sealed class LatexRedrawDialog : Form
{
    private readonly CheckBox _numberDisplayFormulas = new();

    internal LatexRedrawDialog(
        bool wholeDocument,
        int formulaCount,
        int displayFormulaCount,
        string objectModeLabel,
        string equationNumberFormatDisplayName)
    {
        Text = "VisualTeX LaTeX 重绘";
        StartPosition = FormStartPosition.CenterParent;
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(540, 235);
        MinimumSize = new Size(540, 235);
        MaximumSize = new Size(780, 360);
        Font = new Font(
            "Microsoft YaHei UI",
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
            RowCount = 5,
            Padding = new Padding(18),
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        var description = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(500, 0),
            Text = wholeDocument
                ? $"将在整个文档中原位重绘 {formulaCount} 个 LaTeX 公式为 {objectModeLabel}。"
                : $"将在所选内容中原位重绘 {formulaCount} 个 LaTeX 公式为 {objectModeLabel}。",
            Margin = new Padding(0, 0, 0, 12),
        };
        root.Controls.Add(description, 0, 0);

        _numberDisplayFormulas.AutoSize = true;
        _numberDisplayFormulas.Enabled = displayFormulaCount > 0;
        _numberDisplayFormulas.Text = displayFormulaCount > 0
            ? $"为全部 {displayFormulaCount} 个行间公式添加编号"
            : "为所有行间公式添加编号（本次未检测到行间公式）";
        _numberDisplayFormulas.Margin = new Padding(0, 0, 0, 5);
        root.Controls.Add(_numberDisplayFormulas, 0, 1);

        var detail = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(500, 0),
            Text = $"编号格式：{equationNumberFormatDisplayName}。正文和公式定界符以外的内容不会改变；本次操作可通过一次 Ctrl+Z 整体撤销。",
            ForeColor = Color.FromArgb(88, 88, 88),
            Margin = new Padding(22, 0, 0, 12),
        };
        root.Controls.Add(detail, 0, 2);

        var actions = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
        };
        var cancel = new Button
        {
            AutoSize = true,
            Text = "取消",
            DialogResult = DialogResult.Cancel,
            Padding = new Padding(10, 3, 10, 3),
        };
        var redraw = new Button
        {
            AutoSize = true,
            Text = "开始重绘",
            DialogResult = DialogResult.OK,
            Padding = new Padding(10, 3, 10, 3),
        };
        actions.Controls.Add(cancel);
        actions.Controls.Add(redraw);
        root.Controls.Add(actions, 0, 4);

        AcceptButton = redraw;
        CancelButton = cancel;
    }

    internal bool NumberDisplayFormulas =>
        _numberDisplayFormulas.Enabled && _numberDisplayFormulas.Checked;
}
