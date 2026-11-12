using System.Drawing;
using System.Windows.Forms;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed class EquationReferenceDialog : Form
{
    private static string T(string chinese, string english) =>
        OfficePluginLanguage.Text(chinese, english);

    private readonly IReadOnlyList<EquationReferenceTarget> _visualTexTargets;
    private readonly IReadOnlyList<EquationReferenceTarget> _wordOmmlTargets;
    private readonly IReadOnlyList<EquationReferenceTarget> _mathTypeTargets;
    private readonly ComboBox _sourceBox = new();
    private readonly TextBox _searchBox = new();
    private readonly TargetListBox _listBox = new();
    // Fills the targets' LaTeX previews in a few at a time while the dialog shows (each is a Word read).
    private readonly System.Windows.Forms.Timer _previewTimer = new() { Interval = 1 };

    private sealed class TargetListBox : ListBox
    {
        public void RefreshTarget(int index) => RefreshItem(index);
    }
    private readonly Label _styleLabel = new();
    private readonly ComboBox _styleBox = new();

    private sealed class SourceOption
    {
        public SourceOption(EquationReferenceSource source, string label, int count)
        {
            Source = source;
            Label = label;
            Count = count;
        }

        public EquationReferenceSource Source { get; }
        public string Label { get; }
        public int Count { get; }
        public override string ToString() => OfficePluginLanguage.IsEnglish
            ? $"{Label} ({Count})"
            : $"{Label}（{Count}）";
    }

    public EquationReferenceTarget? SelectedTarget =>
        _listBox.SelectedItem as EquationReferenceTarget;

    public EquationReferenceStyle SelectedStyle =>
        CurrentSource == EquationReferenceSource.MathType
            ? EquationReferenceStyle.NumberOnly
            : _styleBox.SelectedIndex switch
            {
                1 => EquationReferenceStyle.EquationPrefix,
                2 => EquationReferenceStyle.NumberOnly,
                _ => EquationReferenceStyle.Parenthesized,
            };

    public EquationReferenceSource CurrentSource =>
        (_sourceBox.SelectedItem as SourceOption)?.Source
        ?? EquationReferenceSource.VisualTeX;

    public EquationReferenceDialog(
        IReadOnlyList<EquationReferenceTarget> visualTexTargets,
        IReadOnlyList<EquationReferenceTarget> wordOmmlTargets,
        IReadOnlyList<EquationReferenceTarget> mathTypeTargets)
    {
        _visualTexTargets = visualTexTargets ?? Array.Empty<EquationReferenceTarget>();
        _wordOmmlTargets = wordOmmlTargets ?? Array.Empty<EquationReferenceTarget>();
        _mathTypeTargets = mathTypeTargets ?? Array.Empty<EquationReferenceTarget>();

        Text = T("插入公式引用", "Insert Equation Reference");
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(640, 474);
        Font = new Font(OfficePluginLanguage.UiFontFamily, 9f);

        var sourceLabel = new Label
        {
            Text = T("引用来源：", "Reference source:"),
            AutoSize = true,
            Location = new Point(18, 19),
        };
        _sourceBox.DropDownStyle = ComboBoxStyle.DropDownList;
        _sourceBox.SetBounds(92, 15, 300, 28);
        if (_visualTexTargets.Count > 0)
            _sourceBox.Items.Add(new SourceOption(
                EquationReferenceSource.VisualTeX,
                T("VisualTeX 编号公式", "Numbered VisualTeX equations"),
                _visualTexTargets.Count));
        if (_wordOmmlTargets.Count > 0)
            _sourceBox.Items.Add(new SourceOption(
                EquationReferenceSource.WordOmml,
                T("Word OMML 编号公式", "Numbered Word OMML equations"),
                _wordOmmlTargets.Count));
        if (_mathTypeTargets.Count > 0)
            _sourceBox.Items.Add(new SourceOption(
                EquationReferenceSource.MathType,
                T("MathType 编号公式", "Numbered MathType equations"),
                _mathTypeTargets.Count));
        if (_sourceBox.Items.Count > 0) _sourceBox.SelectedIndex = 0;
        _sourceBox.SelectedIndexChanged += (_, _) =>
        {
            ConfigureStyleForSource();
            RefreshTargets();
        };

        var searchLabel = new Label
        {
            Text = T("搜索公式：", "Search equations:"),
            AutoSize = true,
            Location = new Point(18, 59),
        };
        _searchBox.SetBounds(92, 55, 528, 26);
        _searchBox.TextChanged += (_, _) => RefreshTargets();

        _listBox.SetBounds(18, 94, 602, 270);
        _listBox.IntegralHeight = false;
        _listBox.DoubleClick += (_, _) => ConfirmSelection();

        _styleLabel.Text = T("引用格式：", "Reference format:");
        _styleLabel.AutoSize = true;
        _styleLabel.Location = new Point(18, 385);
        _styleBox.DropDownStyle = ComboBoxStyle.DropDownList;
        _styleBox.SetBounds(92, 380, 260, 28);

        var insertButton = new Button
        {
            Text = T("插入引用", "Insert Reference"),
            DialogResult = DialogResult.OK,
            Location = new Point(434, 424),
            Size = new Size(90, 32),
        };
        insertButton.Click += (_, _) =>
        {
            if (SelectedTarget is not null) return;
            DialogResult = DialogResult.None;
            MessageBox.Show(
                this,
                T("请先选择一个带编号公式。", "Select a numbered equation first."),
                "VisualTeX",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        };

        var cancelButton = new Button
        {
            Text = T("取消", "Cancel"),
            DialogResult = DialogResult.Cancel,
            Location = new Point(530, 424),
            Size = new Size(90, 32),
        };

        Controls.AddRange(new Control[]
        {
            sourceLabel,
            _sourceBox,
            searchLabel,
            _searchBox,
            _listBox,
            _styleLabel,
            _styleBox,
            insertButton,
            cancelButton,
        });
        AcceptButton = insertButton;
        CancelButton = cancelButton;
        ConfigureStyleForSource();
        RefreshTargets();
        _previewTimer.Tick += (_, _) => FillPreviews();
        Shown += (_, _) => _previewTimer.Start();
        FormClosed += (_, _) => _previewTimer.Stop();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) _previewTimer.Dispose();
        base.Dispose(disposing);
    }

    /// Reads previews for about 30 ms: the listed targets first (in list order), then the other
    /// sources'. Each one read is drawn at once; a search that matched on LaTeX is run again when all
    /// are in.
    private void FillPreviews()
    {
        var started = Environment.TickCount;
        for (var index = 0; index < _listBox.Items.Count; index++)
        {
            if (_listBox.Items[index] is not EquationReferenceTarget target || target.LatexPreviewReady) continue;
            _ = target.LatexPreview;
            _listBox.RefreshTarget(index);
            if (Environment.TickCount - started > 30) return;
        }
        foreach (var target in _visualTexTargets.Concat(_wordOmmlTargets).Concat(_mathTypeTargets))
        {
            if (target.LatexPreviewReady) continue;
            _ = target.LatexPreview;
            if (Environment.TickCount - started > 30) return;
        }
        _previewTimer.Stop();
        if (_searchBox.Text.Trim().Length > 0) RefreshTargets();
    }

    private IReadOnlyList<EquationReferenceTarget> CurrentTargets =>
        CurrentSource switch
        {
            EquationReferenceSource.WordOmml => _wordOmmlTargets,
            EquationReferenceSource.MathType => _mathTypeTargets,
            _ => _visualTexTargets,
        };

    private void ConfigureStyleForSource()
    {
        _styleBox.BeginUpdate();
        try
        {
            _styleBox.Items.Clear();
            if (CurrentSource == EquationReferenceSource.MathType)
            {
                _styleBox.Items.Add(T("沿用 MathType 原编号格式", "Keep the native MathType number format"));
                _styleBox.SelectedIndex = 0;
                _styleBox.Enabled = false;
                _styleLabel.Enabled = false;
                return;
            }

            _styleBox.Items.AddRange(new object[]
            {
                "(1)",
                T("式(1)", "Eq. (1)"),
                "1",
            });
            _styleBox.SelectedIndex = 0;
            _styleBox.Enabled = true;
            _styleLabel.Enabled = true;
        }
        finally { _styleBox.EndUpdate(); }
    }

    private void RefreshTargets()
    {
        var query = _searchBox.Text.Trim();
        var targets = CurrentTargets;
        var filtered = string.IsNullOrWhiteSpace(query)
            ? targets
            // (a preview not read yet matches once it is: FillPreviews runs the search again)
            : targets.Where(target =>
                    target.NumberText.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0
                    || target.LatexPreviewReady
                        && target.LatexPreview.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToArray();

        var selected = SelectedTarget;
        _listBox.BeginUpdate();
        try
        {
            _listBox.Items.Clear();
            foreach (var target in filtered) _listBox.Items.Add(target);
            var keep = selected is null ? -1 : _listBox.Items.IndexOf(selected);
            if (keep >= 0) _listBox.SelectedIndex = keep;
            else if (_listBox.Items.Count > 0) _listBox.SelectedIndex = 0;
        }
        finally { _listBox.EndUpdate(); }
        if (filtered.Any(target => !target.LatexPreviewReady) || targets.Any(target => !target.LatexPreviewReady))
            _previewTimer.Start();
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
