using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordBulkImportDialogOptionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        AssertBulkImportMathTypeDialogSubmission();

        Console.WriteLine(
            "Word bulk-import dialog acceptance passed: MathType native OLE is selectable through a real insert submission. The separate LaTeX code-format regression owns $...$ mixed text/formula syntax.");
    }

    private static void AssertBulkImportMathTypeDialogSubmission()
    {
        using var dialog = BulkImportMathTypeOption.CreateDialog();
        dialog.NumberDisplayFormulas = true;
        var selectors = DescendantControls(dialog)
            .OfType<ComboBox>()
            .Where(combo => combo.Items.Cast<object>().Any(item =>
                (item?.ToString() ?? string.Empty).IndexOf(
                    "MathType",
                    StringComparison.OrdinalIgnoreCase) >= 0))
            .ToArray();
        AssertEqual(1, selectors.Length,
            "Bulk import does not expose exactly one MathType formula-format option.");
        var selector = selectors[0];
        var mathTypeIndex = Enumerable.Range(0, selector.Items.Count)
            .First(index =>
                (selector.Items[index]?.ToString() ?? string.Empty).IndexOf(
                    "MathType",
                    StringComparison.OrdinalIgnoreCase) >= 0);
        selector.SelectedIndex = mathTypeIndex;

        var sourceEditor = FindBulkImportSourceEditor(dialog);
        sourceEditor.Text = "x+1";
        SubmitBulkImportDialog(dialog, sourceEditor);
        AssertEqual(
            FormulaOleContract.MathTypeOleMode,
            BulkImportMathTypeOption.ResolveObjectMode(
                dialog,
                () => FormulaOleContract.WordOmmlMode),
            "Bulk import MathType selection did not resolve to the native MathType OLE object mode.");
        AssertTrue(
            dialog.ParsedDocument.NumberDisplayFormulas,
            "Bulk import did not preserve the global display-formula numbering option through submission.");

        using var redrawDialog = new LatexRedrawDialog(
            wholeDocument: false,
            formulaCount: 3,
            displayFormulaCount: 2,
            objectModeLabel: "MathType",
            equationNumberFormatDisplayName: "按章编号（1.1）");
        var redrawNumberingOption = DescendantControls(redrawDialog)
            .OfType<CheckBox>()
            .Single();
        redrawNumberingOption.Checked = true;
        AssertTrue(
            redrawDialog.NumberDisplayFormulas,
            "LaTeX redraw did not expose a selectable global display-formula numbering option.");
    }

    private static void AssertBulkImportMixedSingleDollarSubmission()
    {
        using var dialog = BulkImportMathTypeOption.CreateDialog();
        var formatSelector = DescendantControls(dialog)
            .OfType<ComboBox>()
            .FirstOrDefault(combo => combo.Items.Cast<object>().Any(item =>
                (item?.ToString() ?? string.Empty).IndexOf(
                    "文字公式混排",
                    StringComparison.OrdinalIgnoreCase) >= 0));
        AssertTrue(formatSelector is not null,
            "Bulk import dialog no longer exposes the mixed text/formula input mode.");
        var mixedIndex = Enumerable.Range(0, formatSelector!.Items.Count)
            .First(index =>
                (formatSelector.Items[index]?.ToString() ?? string.Empty).IndexOf(
                    "文字公式混排",
                    StringComparison.OrdinalIgnoreCase) >= 0);
        formatSelector.SelectedIndex = mixedIndex;

        var sourceEditor = FindBulkImportSourceEditor(dialog);
        sourceEditor.Text = "正文 $x+1$。";
        SubmitBulkImportDialog(dialog, sourceEditor);
        AssertTrue(
            sourceEditor.Text.IndexOf("$$x+1$$", StringComparison.Ordinal) >= 0,
            "The confirmed mixed text/formula dialog did not translate user-facing $...$ before parsing.");

        var visibleHelpStillUsesDoubleDollar = DescendantControls(dialog)
            .Where(control => control is not TextBoxBase { ReadOnly: false })
            .Select(control => control.Text ?? string.Empty)
            .Any(text => text.IndexOf("$$...$$", StringComparison.Ordinal) >= 0
                || text.IndexOf("$$…$$", StringComparison.Ordinal) >= 0);
        AssertTrue(!visibleHelpStillUsesDoubleDollar,
            "Bulk import still tells users to wrap mixed inline formulas in $$...$$.");
    }

    private static TextBoxBase FindBulkImportSourceEditor(BulkImportDialog dialog) =>
        DescendantControls(dialog)
            .OfType<TextBoxBase>()
            .Where(textBox => !textBox.ReadOnly && textBox.Multiline)
            .OrderByDescending(textBox =>
                (long)Math.Max(1, textBox.Width) * Math.Max(1, textBox.Height))
            .FirstOrDefault()
        ?? throw new InvalidOperationException(
            "Bulk import dialog no longer exposes its editable source text box.");

    private static void SubmitBulkImportDialog(
        BulkImportDialog dialog,
        TextBoxBase sourceEditor)
    {
        var okButton = DescendantControls(dialog)
            .OfType<Button>()
            .FirstOrDefault(button => button.DialogResult == DialogResult.OK
                || button.Text.IndexOf("确定", StringComparison.OrdinalIgnoreCase) >= 0
                || button.Text.IndexOf("导入", StringComparison.OrdinalIgnoreCase) >= 0
                || button.Text.IndexOf("插入", StringComparison.OrdinalIgnoreCase) >= 0
                || string.Equals(button.Text.Trim(), "OK", StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException(
                "Bulk import dialog no longer exposes an OK/import button.");

        dialog.Show();
        sourceEditor.Focus();
        Application.DoEvents();
        okButton.PerformClick();
        Application.DoEvents();
        AssertEqual(
            DialogResult.OK,
            dialog.DialogResult,
            "Bulk import dialog did not complete its real OK submission.");
        if (dialog.Visible) dialog.Close();
    }

    private static IEnumerable<Control> DescendantControls(Control root)
    {
        foreach (Control child in root.Controls)
        {
            yield return child;
            foreach (var descendant in DescendantControls(child))
                yield return descendant;
        }
    }
}
