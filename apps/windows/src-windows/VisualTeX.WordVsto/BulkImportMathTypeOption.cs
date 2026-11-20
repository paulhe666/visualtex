using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

/// <summary>
/// Extends the existing bulk-import dialog with the native MathType OLE target and
/// keeps its mixed text/formula syntax user-facing as $...$. The established
/// parser still receives its compatible $$...$$ representation immediately before
/// submission, so existing Markdown/block parsing behavior remains unchanged.
/// </summary>
internal static class BulkImportMathTypeOption
{
    private const string MathTypeLabel = "MathType 公式（原生 OLE）";
    private static readonly object Gate = new();
    private static readonly Dictionary<BulkImportDialog, ComboBox> Selectors = new();

    internal static BulkImportDialog CreateDialog()
    {
        var dialog = new BulkImportDialog();
        Attach(dialog);
        return dialog;
    }

    internal static string ResolveObjectMode(
        BulkImportDialog dialog,
        Func<string> fallbackFactory)
    {
        if (fallbackFactory is null)
            throw new ArgumentNullException(nameof(fallbackFactory));
        if (dialog.SelectedObjectMode == WordBulkFormulaObjectMode.MathType)
            return FormulaOleContract.MathTypeOleMode;
        var selector = ResolveObjectSelector(dialog, attachIfMissing: false);
        var selected = selector?.SelectedItem?.ToString() ?? string.Empty;
        return selected.IndexOf("MathType", StringComparison.OrdinalIgnoreCase) >= 0
            ? FormulaOleContract.MathTypeOleMode
            : fallbackFactory();
    }

    internal static string ConvertSingleDollarMixedSyntaxForParser(string source)
    {
        if (string.IsNullOrEmpty(source)) return source ?? string.Empty;
        var output = new StringBuilder(source.Length + 16);
        var inSingleDollarMath = false;
        for (var index = 0; index < source.Length; index++)
        {
            var character = source[index];
            if (character != '$' || IsEscaped(source, index))
            {
                output.Append(character);
                continue;
            }

            // Existing $$...$$ input remains supported and is copied verbatim.
            if (index + 1 < source.Length && source[index + 1] == '$')
            {
                output.Append("$$");
                index += 1;
                continue;
            }

            output.Append("$$");
            inSingleDollarMath = !inSingleDollarMath;
        }

        // Do not turn malformed input into a different malformed grammar. The
        // existing parser/dialog validation should report the unmatched delimiter.
        return inSingleDollarMath ? source : output.ToString();
    }

    private static bool IsEscaped(string source, int offset)
    {
        var slashes = 0;
        for (var index = offset - 1; index >= 0 && source[index] == '\\'; index--)
            slashes += 1;
        return slashes % 2 == 1;
    }

    private static void Attach(BulkImportDialog dialog)
    {
        if (dialog is null) throw new ArgumentNullException(nameof(dialog));
        var objectSelector = ResolveObjectSelector(dialog, attachIfMissing: true)
            ?? throw new InvalidOperationException(
                "VisualTeX could not find the bulk-import formula-format selector.");
        if (!objectSelector.Items.Cast<object>().Any(item =>
                (item?.ToString() ?? string.Empty).IndexOf(
                    "MathType",
                    StringComparison.OrdinalIgnoreCase) >= 0))
        {
            objectSelector.Items.Add(MathTypeLabel);
        }

        RewriteMixedModeHelp(dialog);
        var formatSelector = ResolveMixedFormatSelector(dialog);
        var sourceEditor = ResolveSourceEditor(dialog);
        void PrepareMixedModeForSubmit()
        {
            if (formatSelector is null || sourceEditor is null) return;
            var selectedFormat = formatSelector.SelectedItem?.ToString() ?? string.Empty;
            if (!IsMixedFormatLabel(selectedFormat)) return;
            sourceEditor.Text = ConvertSingleDollarMixedSyntaxForParser(sourceEditor.Text);
        }

        // WinForms runs validation before a focused editor yields to the OK button,
        // which places the compatible delimiters in the source before the dialog's
        // original Click handler parses it. Mouse/key hooks cover controls with
        // validation disabled; FormClosing is a final defensive fallback.
        if (sourceEditor is not null)
        {
            sourceEditor.Validating += (_, _) => PrepareMixedModeForSubmit();
        }
        foreach (var button in Descendants(dialog).OfType<Button>())
        {
            if (button.DialogResult != DialogResult.OK
                && button.Text.IndexOf("确定", StringComparison.OrdinalIgnoreCase) < 0
                && button.Text.IndexOf("导入", StringComparison.OrdinalIgnoreCase) < 0
                && button.Text.IndexOf("插入", StringComparison.OrdinalIgnoreCase) < 0
                && !string.Equals(button.Text.Trim(), "OK", StringComparison.OrdinalIgnoreCase))
                continue;
            button.MouseDown += (_, _) => PrepareMixedModeForSubmit();
            button.PreviewKeyDown += (_, eventArgs) =>
            {
                if (eventArgs.KeyCode is Keys.Enter or Keys.Space)
                    PrepareMixedModeForSubmit();
            };
        }
        dialog.KeyPreview = true;
        dialog.KeyDown += (_, eventArgs) =>
        {
            if (eventArgs.KeyCode == Keys.Enter)
                PrepareMixedModeForSubmit();
        };
        dialog.FormClosing += (_, eventArgs) =>
        {
            if (!eventArgs.Cancel && dialog.DialogResult == DialogResult.OK)
                PrepareMixedModeForSubmit();
        };
        dialog.FormClosed += (_, _) =>
        {
            lock (Gate) Selectors.Remove(dialog);
        };
    }

    private static void RewriteMixedModeHelp(Control root)
    {
        foreach (var control in DescendantsAndSelf(root))
        {
            if (control is TextBoxBase { ReadOnly: false }) continue;
            var text = control.Text ?? string.Empty;
            var updated = text
                .Replace("$$...$$", "$...$")
                .Replace("$$…$$", "$…$")
                .Replace("$$ 公式 $$", "$ 公式 $");
            if (!string.Equals(text, updated, StringComparison.Ordinal))
                control.Text = updated;
        }
    }

    private static bool IsMixedFormatLabel(string label) =>
        label.IndexOf("文字公式混排", StringComparison.OrdinalIgnoreCase) >= 0
        || label.IndexOf("mixed", StringComparison.OrdinalIgnoreCase) >= 0;

    private static ComboBox? ResolveMixedFormatSelector(BulkImportDialog dialog) =>
        Descendants(dialog)
            .OfType<ComboBox>()
            .FirstOrDefault(combo => combo.Items.Cast<object>().Any(item =>
                IsMixedFormatLabel(item?.ToString() ?? string.Empty)));

    private static TextBoxBase? ResolveSourceEditor(BulkImportDialog dialog) =>
        Descendants(dialog)
            .OfType<TextBoxBase>()
            .Where(textBox => !textBox.ReadOnly && textBox.Multiline)
            .OrderByDescending(textBox =>
                (long)Math.Max(1, textBox.Width) * Math.Max(1, textBox.Height))
            .FirstOrDefault();

    private static ComboBox? ResolveObjectSelector(
        BulkImportDialog dialog,
        bool attachIfMissing)
    {
        lock (Gate)
        {
            if (Selectors.TryGetValue(dialog, out var cached))
                return cached;
        }

        var candidates = Descendants(dialog)
            .OfType<ComboBox>()
            .Where(combo => combo.Items.Cast<object>().Any(item =>
            {
                var label = item?.ToString() ?? string.Empty;
                return label.IndexOf("OMML", StringComparison.OrdinalIgnoreCase) >= 0
                    || label.IndexOf("VisualTeX", StringComparison.OrdinalIgnoreCase) >= 0;
            }))
            .ToArray();
        if (candidates.Length != 1)
        {
            if (!attachIfMissing) return null;
            throw new InvalidOperationException(
                $"VisualTeX found {candidates.Length} possible bulk-import formula-format selectors; expected exactly one.");
        }

        lock (Gate) Selectors[dialog] = candidates[0];
        return candidates[0];
    }

    private static IEnumerable<Control> DescendantsAndSelf(Control root)
    {
        yield return root;
        foreach (var descendant in Descendants(root))
            yield return descendant;
    }

    private static IEnumerable<Control> Descendants(Control root)
    {
        foreach (Control child in root.Controls)
        {
            yield return child;
            foreach (var descendant in Descendants(child))
                yield return descendant;
        }
    }
}
