using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;

namespace VisualTeX.PowerPointVsto;

internal static class PowerPointOmmlBridge
{
    internal static Shape AddNativeEquation(
        Application application,
        Slide slide,
        string mathMl,
        FormulaMetadata metadata,
        float left,
        float top,
        float width,
        float height)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("PowerPoint OMML conversion requires MathML export data.");

        Shape? shape = null;
        try
        {
            shape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                left,
                top,
                Math.Max(1f, width),
                Math.Max(1f, height));
            ConfigureEquationTextBox(shape, metadata);
            PasteMathMlDirectly(shape, mathMl);
            ApplyFontSize(shape, FormulaFontSize.ResolveSemanticFontSize(metadata));
            shape.Left = left;
            shape.Top = top;
            shape.Width = Math.Max(1f, width);
            shape.Height = Math.Max(1f, height);
            var result = shape;
            shape = null;
            return result;
        }
        catch
        {
            if (shape is not null)
            {
                try { shape.Delete(); } catch { }
            }
            throw;
        }
        finally
        {
            Release(shape);
        }
    }

    internal static bool IsNativeEquation(Shape shape)
    {
        object? range = null;
        object? mathZones = null;
        object? textFrame = null;
        try
        {
            if (shape.HasTextFrame != MsoTriState.msoTrue) return false;
            textFrame = shape.TextFrame2;
            range = ((dynamic)textFrame).TextRange;
            mathZones = GetMathZones(range);
            // MathZones returns a fallback text range even when Count is zero.
            return Convert.ToInt32(((dynamic)mathZones).Count) > 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(mathZones);
            Release(range);
            Release(textFrame);
        }
    }

    internal static float? ReadFontSize(Shape shape)
    {
        object? range = null;
        object? mathZones = null;
        object? font = null;
        object? textFrame = null;
        try
        {
            if (shape.HasTextFrame != MsoTriState.msoTrue) return null;
            textFrame = shape.TextFrame2;
            range = ((dynamic)textFrame).TextRange;
            mathZones = GetMathZones(range);
            if (Convert.ToInt32(((dynamic)mathZones).Count) <= 0) return null;
            font = ((dynamic)mathZones).Font;
            return FormulaFontSize.Normalize(Convert.ToDouble(((dynamic)font).Size), 20f);
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(font);
            Release(mathZones);
            Release(range);
            Release(textFrame);
        }
    }

    internal static string? TryReadCurrentLatex(Shape shape)
    {
        var mathMl = TryReadMathMl(shape);
        if (string.IsNullOrWhiteSpace(mathMl)) return null;
        try { return MathMlToLatexConverter.ConvertWordEquation(mathMl!); }
        catch { return null; }
    }

    internal static string? TryReadMathMl(Shape shape)
    {
        var clipboard = CaptureClipboard();
        object? range = null;
        object? mathZones = null;
        object? textFrame = null;
        try
        {
            if (shape.HasTextFrame != MsoTriState.msoTrue) return null;
            textFrame = shape.TextFrame2;
            range = ((dynamic)textFrame).TextRange;
            mathZones = GetMathZones(range);
            if (Convert.ToInt32(((dynamic)mathZones).Count) <= 0) return null;
            ((dynamic)mathZones).Copy();
            return ReadMathMlFromClipboard();
        }
        catch
        {
            return null;
        }
        finally
        {
            RestoreClipboard(clipboard);
            Release(mathZones);
            Release(range);
            Release(textFrame);
        }
    }

    private static void PasteMathMlDirectly(Shape shape, string mathMl)
    {
        var clipboard = CaptureClipboard();
        Exception? lastError = null;
        try
        {
            // PowerPoint can import native Office Math directly from the standard
            // MathML clipboard format. The payload Office itself publishes is
            // UTF-8 with a trailing NUL. Avoiding a transient Word.Application
            // removes both the multi-second startup cost and intermittent Word
            // automation failures from PowerPoint OMML conversion.
            for (var attempt = 0; attempt < 3; attempt++)
            {
                object? textFrame = null;
                object? targetRange = null;
                MemoryStream? mathMlStream = null;
                try
                {
                    var data = new System.Windows.Forms.DataObject();
                    var bytes = Encoding.UTF8.GetBytes(mathMl + "\0");
                    mathMlStream = new MemoryStream(bytes, writable: false);
                    data.SetData("MathML", autoConvert: false, mathMlStream);
                    data.SetData(System.Windows.Forms.DataFormats.UnicodeText, mathMl);
                    SetClipboardDataObject(data);

                    textFrame = shape.TextFrame2;
                    targetRange = ((dynamic)textFrame).TextRange;
                    ((dynamic)targetRange).Text = string.Empty;
                    ((dynamic)targetRange).Paste();
                    if (WaitForNativeEquation(shape, attempt == 0 ? 80 : 180))
                        return;
                    lastError = new InvalidOperationException(
                        "PowerPoint did not materialize the pasted MathML as native Office Math.");
                }
                catch (COMException error)
                {
                    lastError = error;
                }
                finally
                {
                    mathMlStream?.Dispose();
                    Release(targetRange);
                    Release(textFrame);
                }

                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(20 * (attempt + 1));
            }

            if (lastError is COMException comError)
            {
                throw new InvalidOperationException(
                    $"PowerPoint native Office Math import failed after retry (0x{comError.HResult:X8}): {comError.Message}",
                    comError);
            }
            throw new InvalidOperationException(
                "PowerPoint did not import the MathML as native Office Math after retry.",
                lastError);
        }
        finally
        {
            RestoreClipboard(clipboard);
        }
    }

    private static bool WaitForNativeEquation(Shape shape, int timeoutMilliseconds)
    {
        if (IsNativeEquation(shape)) return true;
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        while (stopwatch.ElapsedMilliseconds < timeoutMilliseconds)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(10);
            if (IsNativeEquation(shape)) return true;
        }
        return false;
    }

    internal static object? ReadSingleSelectedMathZone(object range)
    {
        object? zones = null;
        try
        {
            zones = GetMathZones(range);
            if (Convert.ToInt32(((dynamic)zones).Count) != 1) return null;
            return ((dynamic)range).MathZones(1, 1);
        }
        finally { Release(zones); }
    }

    private static object GetMathZones(object range)
    {
        try { return ((dynamic)range).MathZones(); }
        catch { return ((dynamic)range).MathZones(Type.Missing, Type.Missing); }
    }

    private static void ConfigureEquationTextBox(Shape shape, FormulaMetadata metadata)
    {
        try { shape.Fill.Visible = MsoTriState.msoFalse; } catch { }
        try { shape.Line.Visible = MsoTriState.msoFalse; } catch { }
        try
        {
            var frame = shape.TextFrame2;
            frame.MarginLeft = 0;
            frame.MarginRight = 0;
            frame.MarginTop = 0;
            frame.MarginBottom = 0;
            frame.WordWrap = MsoTriState.msoFalse;
            frame.AutoSize = MsoAutoSize.msoAutoSizeNone;
            frame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
            frame.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
            Release(frame);
        }
        catch { }
        ApplyFontSize(shape, FormulaFontSize.ResolveSemanticFontSize(metadata));
    }

    internal static void SetFontSize(Shape shape, float size) =>
        ApplyFontSize(shape, FormulaFontSize.Normalize(size, 20f));

    private static void ApplyFontSize(Shape shape, float size)
    {
        object? range = null;
        object? mathZones = null;
        object? font = null;
        object? textFrame = null;
        try
        {
            if (shape.HasTextFrame != MsoTriState.msoTrue) return;
            textFrame = shape.TextFrame2;
            range = ((dynamic)textFrame).TextRange;
            mathZones = GetMathZones(range);
            if (Convert.ToInt32(((dynamic)mathZones).Count) <= 0) return;
            font = ((dynamic)mathZones).Font;
            ((dynamic)font).Size = size;
        }
        finally
        {
            Release(font);
            Release(mathZones);
            Release(range);
            Release(textFrame);
        }
    }

    private static System.Windows.Forms.IDataObject? CaptureClipboard()
    {
        try { return System.Windows.Forms.Clipboard.GetDataObject(); }
        catch { return null; }
    }

    private static void SetClipboardDataObject(System.Windows.Forms.IDataObject data)
    {
        Exception? lastError = null;
        for (var attempt = 0; attempt < 4; attempt++)
        {
            try
            {
                System.Windows.Forms.Clipboard.SetDataObject(data, true);
                return;
            }
            catch (ExternalException error)
            {
                lastError = error;
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(15 * (attempt + 1));
            }
        }
        throw new InvalidOperationException(
            "PowerPoint OMML conversion could not access the Windows clipboard.",
            lastError);
    }

    private static void RestoreClipboard(System.Windows.Forms.IDataObject? data)
    {
        if (data is null) return;
        // Restoring with copy=true forces every delayed clipboard format to be
        // rendered synchronously. If the user's clipboard currently contains a
        // PowerPoint/Word OLE object that can take seconds. We only need to hand
        // the original IDataObject back to OLE; keep its delayed formats lazy.
        for (var attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                System.Windows.Forms.Clipboard.SetDataObject(data, false);
                return;
            }
            catch (ExternalException)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(10 * (attempt + 1));
            }
        }
    }

    private static string? ReadMathMlFromClipboard()
    {
        foreach (var format in new[] { "MathML", "MathML Presentation" })
        {
            try
            {
                var data = System.Windows.Forms.Clipboard.GetData(format);
                var decoded = DecodeClipboardMathMl(data);
                if (!string.IsNullOrWhiteSpace(decoded)) return decoded;
            }
            catch { }
        }
        try
        {
            var text = System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.UnicodeText);
            return LooksLikeMathMl(text) ? text : null;
        }
        catch { return null; }
    }

    private static string? DecodeClipboardMathMl(object? data)
    {
        switch (data)
        {
            case null:
                return null;
            case string text:
                return text.TrimEnd('\0');
            case byte[] bytes:
                return DecodeMathMlBytes(bytes);
            case MemoryStream stream:
            {
                var position = stream.CanSeek ? stream.Position : 0;
                try
                {
                    if (stream.CanSeek) stream.Position = 0;
                    using var copy = new MemoryStream();
                    stream.CopyTo(copy);
                    return DecodeMathMlBytes(copy.ToArray());
                }
                finally
                {
                    if (stream.CanSeek) stream.Position = position;
                }
            }
            default:
            {
                var text = data.ToString();
                return LooksLikeMathMl(text) ? text : null;
            }
        }
    }

    private static string? DecodeMathMlBytes(byte[] bytes)
    {
        if (bytes.Length == 0) return null;
        foreach (var encoding in new[]
                 {
                     new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: false),
                     Encoding.Unicode,
                     Encoding.BigEndianUnicode,
                 })
        {
            var text = encoding.GetString(bytes).TrimStart('\uFEFF').TrimEnd('\0');
            if (LooksLikeMathMl(text)) return text;
        }
        return null;
    }

    private static bool LooksLikeMathMl(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        var value = text!.TrimStart();
        return value.StartsWith("<math", StringComparison.OrdinalIgnoreCase)
            || value.StartsWith("<mml:math", StringComparison.OrdinalIgnoreCase);
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
