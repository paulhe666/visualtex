using System.Drawing;
using System.Text;
using System.Windows.Forms;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed class BulkImportDialog : Form
{
    private readonly ComboBox _sourceFormat = new();
    private readonly ComboBox _objectMode = new();
    private readonly TextBox _source = new();
    private readonly Label _summary = new();
    private readonly TextBox _warnings = new();
    private readonly Button _insert = new();
    private WordBulkImportDocument? _parsed;

    internal BulkImportDialog()
    {
        Text = "VisualTeX 批量导入 LaTeX / Markdown";
        StartPosition = FormStartPosition.CenterParent;
        MinimumSize = new Size(820, 600);
        Size = new Size(980, 720);
        Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimizeBox = false;
        ShowIcon = false;

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5,
            Padding = new Padding(14),
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 82));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        var title = new Label
        {
            AutoSize = true,
            Text = "批量导入为 Word 原生文字和独立公式",
            Font = new Font(Font, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 5),
        };
        var description = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(900, 0),
            Text = "普通文字会成为 Word 原生段落；每个行内公式和行间公式都会成为独立的 VisualTeX 公式，可分别编辑和调整字号。",
            ForeColor = Color.FromArgb(80, 80, 80),
            Margin = new Padding(0, 0, 0, 10),
        };
        var heading = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
        };
        heading.Controls.Add(title);
        heading.Controls.Add(description);
        root.Controls.Add(heading, 0, 0);

        var options = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            Margin = new Padding(0, 0, 0, 8),
        };
        options.Controls.Add(new Label
        {
            AutoSize = true,
            Text = "源格式：",
            Padding = new Padding(0, 7, 0, 0),
        });
        _sourceFormat.DropDownStyle = ComboBoxStyle.DropDownList;
        _sourceFormat.Width = 145;
        _sourceFormat.Items.AddRange(new object[] { "自动识别", "Markdown", "LaTeX" });
        _sourceFormat.SelectedIndex = 0;
        options.Controls.Add(_sourceFormat);
        options.Controls.Add(new Label
        {
            AutoSize = true,
            Text = "公式格式：",
            Padding = new Padding(18, 7, 0, 0),
        });
        _objectMode.DropDownStyle = ComboBoxStyle.DropDownList;
        _objectMode.Width = 210;
        _objectMode.Items.AddRange(new object[]
        {
            "Word 原生 OMML（推荐）",
            "VisualTeX OLE",
            "MathType OLE",
        });
        _objectMode.SelectedIndex = 0;
        options.Controls.Add(_objectMode);
        var open = new Button
        {
            AutoSize = true,
            Text = "打开文件…",
            Margin = new Padding(18, 0, 0, 0),
        };
        open.Click += (_, _) => OpenFile();
        options.Controls.Add(open);
        var preview = new Button
        {
            AutoSize = true,
            Text = "解析预览",
            Margin = new Padding(8, 0, 0, 0),
        };
        preview.Click += (_, _) => ParseAndPreview(showError: true);
        options.Controls.Add(preview);
        root.Controls.Add(options, 0, 1);

        _source.Dock = DockStyle.Fill;
        _source.Multiline = true;
        _source.AcceptsReturn = true;
        _source.AcceptsTab = true;
        _source.ScrollBars = ScrollBars.Both;
        _source.WordWrap = false;
        _source.Font = new Font("Consolas", 10.5f, FontStyle.Regular, GraphicsUnit.Point);
        _source.Text = "# 示例\r\n\r\n这是正文和行内公式 $E=mc^2$。\r\n\r\n$$\r\n\\int_0^1 x^2\\,\\mathrm{d}x=\\frac13\r\n$$";
        _source.TextChanged += (_, _) =>
        {
            _parsed = null;
            _summary.Text = "内容已更改，请解析预览。";
            _warnings.Clear();
        };
        root.Controls.Add(_source, 0, 2);

        var status = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Margin = new Padding(0, 8, 0, 8),
        };
        status.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
        status.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
        _summary.Dock = DockStyle.Fill;
        _summary.Padding = new Padding(8);
        _summary.Text = "点击“解析预览”查看导入结构。";
        _summary.BorderStyle = BorderStyle.FixedSingle;
        status.Controls.Add(_summary, 0, 0);
        _warnings.Dock = DockStyle.Fill;
        _warnings.Multiline = true;
        _warnings.ReadOnly = true;
        _warnings.ScrollBars = ScrollBars.Vertical;
        _warnings.BackColor = SystemColors.Window;
        status.Controls.Add(_warnings, 1, 0);
        root.Controls.Add(status, 0, 3);

        var actions = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
        };
        var cancel = new Button
        {
            Text = "取消",
            DialogResult = DialogResult.Cancel,
            AutoSize = true,
            Padding = new Padding(10, 3, 10, 3),
        };
        _insert.Text = "插入到 Word";
        _insert.AutoSize = true;
        _insert.Padding = new Padding(10, 3, 10, 3);
        _insert.Click += (_, _) =>
        {
            if (!ParseAndPreview(showError: true)) return;
            DialogResult = DialogResult.OK;
            Close();
        };
        actions.Controls.Add(cancel);
        actions.Controls.Add(_insert);
        root.Controls.Add(actions, 0, 4);
        AcceptButton = _insert;
        CancelButton = cancel;
    }

    internal WordBulkImportDocument ParsedDocument =>
        _parsed ?? throw new InvalidOperationException("批量导入内容尚未解析。");

    internal string SourceText
    {
        get => _source.Text;
        set => _source.Text = value ?? string.Empty;
    }

    internal WordBulkSourceFormat SelectedSourceFormat
    {
        get => _sourceFormat.SelectedIndex switch
        {
            1 => WordBulkSourceFormat.Markdown,
            2 => WordBulkSourceFormat.Latex,
            _ => WordBulkSourceFormat.Auto,
        };
        set => _sourceFormat.SelectedIndex = value switch
        {
            WordBulkSourceFormat.Markdown => 1,
            WordBulkSourceFormat.Latex => 2,
            _ => 0,
        };
    }

    internal WordBulkFormulaObjectMode SelectedObjectMode
    {
        get => _objectMode.SelectedIndex switch
        {
            1 => WordBulkFormulaObjectMode.Ole,
            2 => WordBulkFormulaObjectMode.MathType,
            _ => WordBulkFormulaObjectMode.Omml,
        };
        set => _objectMode.SelectedIndex = value switch
        {
            WordBulkFormulaObjectMode.Ole => 1,
            WordBulkFormulaObjectMode.MathType => 2,
            _ => 0,
        };
    }

    private bool ParseAndPreview(bool showError)
    {
        try
        {
            _parsed = WordBulkImportParser.Parse(
                _source.Text,
                SelectedSourceFormat,
                SelectedObjectMode);
            _summary.Text =
                $"识别为 {_parsed.SourceFormat}；共 {_parsed.Blocks.Count} 个块，" +
                $"{_parsed.TextCharacterCount} 个文字字符，" +
                $"{_parsed.InlineFormulaCount} 个行内公式，" +
                $"{_parsed.DisplayFormulaCount} 个行间公式。";
            _warnings.Text = _parsed.Warnings.Count == 0
                ? "没有解析警告。"
                : string.Join(Environment.NewLine, _parsed.Warnings.Select((warning, index) => $"{index + 1}. {warning}"));
            return true;
        }
        catch (Exception error)
        {
            _parsed = null;
            _summary.Text = "无法解析当前内容。";
            _warnings.Text = error.Message;
            if (showError)
            {
                MessageBox.Show(
                    this,
                    error.Message,
                    "VisualTeX 批量导入",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            return false;
        }
    }

    private void OpenFile()
    {
        using var dialog = new OpenFileDialog
        {
            Title = "打开 Markdown 或 LaTeX 文件",
            Filter = "Markdown / LaTeX (*.md;*.markdown;*.tex;*.txt)|*.md;*.markdown;*.tex;*.txt|所有文件 (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false,
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var file = new FileInfo(dialog.FileName);
        if (file.Length > 5_000_000)
        {
            MessageBox.Show(
                this,
                "文件超过 5 MB，无法批量导入。",
                "VisualTeX 批量导入",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }
        _source.Text = File.ReadAllText(file.FullName, DetectEncoding(file.FullName));
        if (file.Extension.Equals(".tex", StringComparison.OrdinalIgnoreCase))
            SelectedSourceFormat = WordBulkSourceFormat.Latex;
        else if (file.Extension.Equals(".md", StringComparison.OrdinalIgnoreCase)
                 || file.Extension.Equals(".markdown", StringComparison.OrdinalIgnoreCase))
            SelectedSourceFormat = WordBulkSourceFormat.Markdown;
        ParseAndPreview(showError: false);
    }

    private static Encoding DetectEncoding(string path)
    {
        using var stream = File.OpenRead(path);
        var prefix = new byte[Math.Min(4, (int)stream.Length)];
        _ = stream.Read(prefix, 0, prefix.Length);
        if (prefix.Length >= 3 && prefix[0] == 0xEF && prefix[1] == 0xBB && prefix[2] == 0xBF)
            return new UTF8Encoding(true, true);
        if (prefix.Length >= 2 && prefix[0] == 0xFF && prefix[1] == 0xFE)
            return Encoding.Unicode;
        if (prefix.Length >= 2 && prefix[0] == 0xFE && prefix[1] == 0xFF)
            return Encoding.BigEndianUnicode;
        return new UTF8Encoding(false, true);
    }
}
