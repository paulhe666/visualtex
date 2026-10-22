using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace VisualTeX.PowerPointVsto;

// A synchronous UI-thread text mutation. Source proof is transient, never stored
// in the presentation. Undo may only remove the prefix above the captured native
// history; original authored shape XML must then verify native formatting recovery.
internal static class PowerPointTextEditTransaction
{
    internal static string Fingerprint(string source)
    {
        var document = XElement.Parse(source, LoadOptions.PreserveWhitespace);
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        foreach (var properties in document.DescendantsAndSelf().Where(element =>
                     element.Name.Namespace == drawing && element.Name.LocalName is "rPr" or "defRPr" or "endParaRPr"))
        {
            // ECMA DrawingML CT_TextCharacterProperties: proofing-result caches
            // change during idle spelling/smart-tag checks without an author edit.
            // Keep ALL authoring data (including noProof, language, font, color,
            // spacing, text, geometry and links). Do not normalize math/content.
            properties.Attribute("dirty")?.Remove();
            properties.Attribute("err")?.Remove();
            properties.Attribute("smtClean")?.Remove();
        }
        using var hash = SHA256.Create();
        return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting)))).Replace("-", "");
    }

    internal static void Execute(Application application, Presentation presentation,
        int slideId, int shapeId, string expectedSource, Action mutate)
    {
        RequireActive(application, presentation);
        var originalXml = PowerPointNativeOmmlPackage.ReadTextShapeXml(presentation, slideId, shapeId);
        if (!string.Equals(Fingerprint(originalXml), expectedSource, StringComparison.Ordinal))
            throw new InvalidOperationException(VisualTeX.WindowsOffice.VstoShared.OfficePluginLanguage.Text(
                "文本框内容或格式在编辑器打开后已改变，请重新选择插入位置；未写入公式。",
                "The text box's content or formatting changed after the editor opened. Choose the insertion point again; the formula was not written."));
        application.StartNewUndoEntry(); // Failure here must precede any mutation.
        var originalHistory = ReadHistory(application);
        try
        {
            RequireActive(application, presentation);
            mutate();
            application.StartNewUndoEntry();
        }
        catch (Exception primary)
        {
            try
            {
                RequireActive(application, presentation);
                application.StartNewUndoEntry();
                var current = ReadHistory(application);
                var owned = CountOwnedPrefix(originalHistory, current);
                for (var index = 0; index < owned; index++)
                {
                    RequireActive(application, presentation);
                    if (!ReadHistory(application).SequenceEqual(current.Skip(index), StringComparer.Ordinal))
                        throw new InvalidDataException("PowerPoint undo history changed during recovery; no further Undo was attempted.");
                    application.CommandBars.ExecuteMso("Undo");
                }
                if (!ReadHistory(application).SequenceEqual(originalHistory, StringComparer.Ordinal)
                    || !string.Equals(Fingerprint(PowerPointNativeOmmlPackage.ReadTextShapeXml(presentation, slideId, shapeId)),
                        expectedSource, StringComparison.Ordinal))
                    throw new InvalidDataException("PowerPoint could not verify exact text, formatting and native Undo recovery.");
            }
            catch (Exception recovery)
            {
                throw new AggregateException(VisualTeX.WindowsOffice.VstoShared.OfficePluginLanguage.Text(
                    "PowerPoint 公式写入失败，且无法确认原内容已完整恢复；请保留文档并检查。",
                    "Writing the PowerPoint formula failed, and VisualTeX could not confirm that the original content was fully restored. Keep the presentation and check it."),
                    primary, recovery);
            }
            throw;
        }
    }

    internal static int CountOwnedPrefix(IReadOnlyList<string> original, IReadOnlyList<string> current)
    {
        var count = current.Count - original.Count;
        if (count < 0 || !current.Skip(Math.Max(0, count)).SequenceEqual(original, StringComparer.Ordinal))
            throw new InvalidDataException("PowerPoint's previous undo boundary changed or was truncated; recovery cannot cross it.");
        return count;
    }

    private static string[] ReadHistory(Application application)
    {
        CommandBars? bars = null;
        CommandBar? standard = null;
        CommandBarControl? control = null;
        try
        {
            bars = application.CommandBars;
            standard = bars["Standard"];
            control = standard.FindControl(Id: 128);
            var history = control as CommandBarComboBox
                ?? throw new InvalidOperationException("PowerPoint exposes no readable native Undo history; text has not been changed.");
            var enabled = bars.GetEnabledMso("Undo");
            if (history.Enabled != enabled) throw new InvalidDataException("PowerPoint Undo availability is inconsistent.");
            if (!enabled) return Array.Empty<string>();
            var count = history.ListCount;
            if (count <= 0) throw new InvalidDataException("PowerPoint enables Undo without a readable history.");
            var entries = new string[count];
            for (var index = 0; index < count; index++) entries[index] = history.get_List(index + 1);
            if (history.ListCount != count) throw new InvalidDataException("PowerPoint Undo history changed while being captured.");
            return entries;
        }
        finally { Release(control); Release(standard); Release(bars); }
    }

    private static void RequireActive(Application application, Presentation presentation)
    {
        Presentation? active = null;
        IntPtr expected = IntPtr.Zero, actual = IntPtr.Zero;
        try
        {
            active = application.ActivePresentation;
            expected = Marshal.GetIUnknownForObject(presentation);
            actual = Marshal.GetIUnknownForObject(active);
            if (expected != actual) throw new InvalidOperationException("The active PowerPoint presentation changed; no edit or Undo is authorized.");
        }
        finally
        {
            if (actual != IntPtr.Zero) Marshal.Release(actual);
            if (expected != IntPtr.Zero) Marshal.Release(expected);
            Release(active);
        }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
