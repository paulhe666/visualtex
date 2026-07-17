using System.Drawing;
using System.Windows.Forms;

namespace VisualTeX.WordVsto;

internal sealed class EquationReferenceDialog : Form
{
    private readonly IReadOnlyList<EquationReferenceTarget> _targets;
    private readonly TextBox _searchBox = new();
    private readonly ListBox _listBox = new();
    private readonly ComboBox _styleBox = new();

    public EquationReferenceTarget? SelectedTarget =>
        _listBox.SelectedItem as EquationReferenceTarget;

    public EquationReferenceStyle SelectedStyle => _styleBox.SelectedIndex switch
    {
        1 => EquationReferenceStyle.EquationPrefix,
        2 => EquationReferenceStyle.NumberOnly,
        _ => EquationReferenceStyle.Parenthesized,
    };

    public EquationReferenceDialog(IReadOnlyList<EquationReferenceTarget> targets)
    {
        _targets = targets;
        Text = "插入 VisualTeX 公式引用";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(620, 420);
        Font = new Font("Microsoft YaHei UI", 9f);

        var searchLabel = new Label
        {
            Text = "搜索公式：",
            AutoSize = true,
            Location = new Point(18, 19),
        };
        _searchBox.SetBounds(92, 15, 508, 26);
        _searchBox.TextChanged += (_, _) => RefreshTargets();

        _listBox.SetBounds(18, 54, 582, 260);
        _listBox.IntegralHeight = false;
        _listBox.DoubleClick += (_, _) => ConfirmSelection();

        var styleLabel = new Label
        {
            Text = "引用格式：",
            AutoSize = true,
            Location = new Point(18, 333),
        };
        _styleBox.DropDownStyle = ComboBoxStyle.DropDownList;
        _styleBox.SetBounds(92, 328, 240, 28);
        _styleBox.Items.AddRange(new object[]
        {
            "(1)",
            "式（1）",
            "1",
        });
        _styleBox.SelectedIndex = 0;

        var insertButton = new Button
        {
            Text = "插入引用",
            DialogResult = DialogResult.OK,
            Location = new Point(414, 370),
            Size = new Size(90, 32),
        };
        insertButton.Click += (_, _) =>
        {
            if (SelectedTarget is not null) return;
            DialogResult = DialogResult.None;
            MessageBox.Show(
                this,
                "请先选择一个带编号公式。",
                "VisualTeX",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        };

        var cancelButton = new Button
        {
            Text = "取消",
            DialogResult = DialogResult.Cancel,
            Location = new Point(510, 370),
            Size = new Size(90, 32),
        };

        Controls.AddRange(new Control[]
        {
            searchLabel,
            _searchBox,
            _listBox,
            styleLabel,
            _styleBox,
            insertButton,
            cancelButton,
        });
        AcceptButton = insertButton;
        CancelButton = cancelButton;
        RefreshTargets();
    }

    private void RefreshTargets()
    {
        var query = _searchBox.Text.Trim();
        var filtered = string.IsNullOrWhiteSpace(query)
            ? _targets
            : _targets.Where(target =>
                    target.NumberText.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0
                    || target.LatexPreview.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToArray();

        _listBox.BeginUpdate();
        try
        {
            _listBox.Items.Clear();
            foreach (var target in filtered) _listBox.Items.Add(target);
            if (_listBox.Items.Count > 0) _listBox.SelectedIndex = 0;
        }
        finally { _listBox.EndUpdate(); }
    }

    private void ConfirmSelection()
    {
        if (SelectedTarget is null) return;
        DialogResult = DialogResult.OK;
        Close();
    }
}

internal sealed class NativeWindowOwner : IWin32Window
{
    public NativeWindowOwner(IntPtr handle) => Handle = handle;
    public IntPtr Handle { get; }
}
