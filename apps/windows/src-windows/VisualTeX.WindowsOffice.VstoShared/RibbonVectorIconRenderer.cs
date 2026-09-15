using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;

namespace VisualTeX.WindowsOffice.VstoShared;

internal static class RibbonVectorIconRenderer
{
    internal const int PixelSize = 64;
    internal const float Dpi = 192f;

    private static readonly Color Ink = Color.FromArgb(255, 48, 55, 65);
    private static readonly Color Accent = Color.FromArgb(255, 38, 111, 150);
    private static readonly Color Warm = Color.FromArgb(255, 151, 103, 53);
    private static readonly Color Paper = Color.FromArgb(255, 250, 248, 243);

    internal static readonly HashSet<string> Keys = new(
        new[]
        {
            "oleDisplay",
            "ommlDisplay",
            "oleInline",
            "ommlInline",
            "insertFormula",
            "updateNumbers",
            "editSelected",
            "convertToOmml",
            "convertToOle",
            "batchImport",
            "repairBaseline",
        },
        StringComparer.OrdinalIgnoreCase);

    internal static Bitmap Create(string key)
    {
        if (!Keys.Contains(key))
            throw new ArgumentOutOfRangeException(nameof(key), key, "Unknown Ribbon icon key.");

        var bitmap = new Bitmap(PixelSize, PixelSize, PixelFormat.Format32bppPArgb);
        bitmap.SetResolution(Dpi, Dpi);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.Transparent);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        graphics.CompositingQuality = CompositingQuality.HighQuality;
        graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

        switch (key.ToLowerInvariant())
        {
            case "oledisplay":
                DrawDisplayFormula(graphics, native: false);
                break;
            case "ommldisplay":
                DrawDisplayFormula(graphics, native: true);
                break;
            case "oleinline":
                DrawInlineFormula(graphics, native: false);
                break;
            case "ommlinline":
                DrawInlineFormula(graphics, native: true);
                break;
            case "insertformula":
                DrawInsertFormula(graphics);
                break;
            case "updatenumbers":
                DrawUpdateNumbers(graphics);
                break;
            case "editselected":
                DrawEditSelected(graphics);
                break;
            case "converttoomml":
                DrawConversion(graphics, toNative: true);
                break;
            case "converttoole":
                DrawConversion(graphics, toNative: false);
                break;
            case "batchimport":
                DrawBatchImport(graphics);
                break;
            case "repairbaseline":
                DrawRepairBaseline(graphics);
                break;
        }

        return bitmap;
    }

    private static Pen Pen(Color color, float width, LineCap cap = LineCap.Round)
    {
        return new Pen(color, width)
        {
            StartCap = cap,
            EndCap = cap,
            LineJoin = LineJoin.Round,
        };
    }

    private static void DrawDocument(Graphics graphics, RectangleF bounds, Color outline)
    {
        using var fill = new SolidBrush(Paper);
        using var pen = Pen(outline, 3.2f);
        graphics.FillRoundedRectangle(fill, bounds, 4.5f);
        graphics.DrawRoundedRectangle(pen, bounds, 4.5f);
        using var fold = Pen(outline, 2.4f);
        graphics.DrawLine(fold, bounds.Right - 12, bounds.Top, bounds.Right - 12, bounds.Top + 11);
        graphics.DrawLine(fold, bounds.Right - 12, bounds.Top + 11, bounds.Right, bounds.Top + 11);
    }

    private static void DrawFormulaText(Graphics graphics, RectangleF bounds, Color color, float points)
    {
        using var font = ResolveMathFont(points, FontStyle.Italic);
        using var brush = new SolidBrush(color);
        using var format = new StringFormat
        {
            Alignment = StringAlignment.Center,
            LineAlignment = StringAlignment.Center,
            FormatFlags = StringFormatFlags.NoWrap,
        };
        graphics.DrawString("fx", font, brush, bounds, format);
    }

    private static Font ResolveMathFont(float points, FontStyle style)
    {
        foreach (var family in new[] { "Cambria Math", "Cambria", "Times New Roman", "Segoe UI" })
        {
            try { return new Font(family, points, style, GraphicsUnit.Point); }
            catch { }
        }
        return new Font(FontFamily.GenericSerif, points, style, GraphicsUnit.Point);
    }

    private static void DrawDisplayFormula(Graphics graphics, bool native)
    {
        var color = native ? Accent : Warm;
        DrawDocument(graphics, new RectangleF(8, 6, 48, 52), Ink);
        using var line = Pen(Ink, 2.4f);
        graphics.DrawLine(line, 15, 18, 49, 18);
        graphics.DrawLine(line, 15, 47, 49, 47);
        using var boxPen = Pen(color, 3.2f);
        graphics.DrawRoundedRectangle(boxPen, new RectangleF(14, 24, 36, 18), 4f);
        DrawFormulaText(graphics, new RectangleF(17, 22, 30, 22), color, 10.5f);
        DrawModeBadge(graphics, native, new PointF(48, 49));
    }

    private static void DrawInlineFormula(Graphics graphics, bool native)
    {
        var color = native ? Accent : Warm;
        using var textPen = Pen(Ink, 2.7f);
        graphics.DrawLine(textPen, 7, 20, 22, 20);
        graphics.DrawLine(textPen, 43, 20, 57, 20);
        graphics.DrawLine(textPen, 7, 42, 57, 42);
        using var formulaPen = Pen(color, 3f);
        graphics.DrawRoundedRectangle(formulaPen, new RectangleF(23, 10, 19, 20), 4f);
        DrawFormulaText(graphics, new RectangleF(23, 9, 19, 22), color, 8f);
        DrawModeBadge(graphics, native, new PointF(50, 48));
    }

    private static void DrawModeBadge(Graphics graphics, bool native, PointF center)
    {
        var color = native ? Accent : Warm;
        using var fill = new SolidBrush(color);
        graphics.FillEllipse(fill, center.X - 6.5f, center.Y - 6.5f, 13, 13);
        using var pen = Pen(Color.White, 2.1f);
        if (native)
        {
            graphics.DrawLine(pen, center.X - 3.5f, center.Y, center.X - 0.5f, center.Y + 3);
            graphics.DrawLine(pen, center.X - 0.5f, center.Y + 3, center.X + 4, center.Y - 3.5f);
        }
        else
        {
            graphics.DrawEllipse(pen, center.X - 3.5f, center.Y - 3.5f, 7, 7);
        }
    }

    private static void DrawInsertFormula(Graphics graphics)
    {
        DrawDocument(graphics, new RectangleF(7, 8, 40, 48), Ink);
        DrawFormulaText(graphics, new RectangleF(12, 21, 30, 23), Accent, 11.5f);
        using var circle = new SolidBrush(Warm);
        graphics.FillEllipse(circle, 39, 34, 20, 20);
        using var plus = Pen(Color.White, 3.5f);
        graphics.DrawLine(plus, 49, 39, 49, 49);
        graphics.DrawLine(plus, 44, 44, 54, 44);
    }

    private static void DrawUpdateNumbers(Graphics graphics)
    {
        using var arc = Pen(Accent, 4f);
        graphics.DrawArc(arc, 8, 8, 44, 44, 35, 285);
        using var arrow = new SolidBrush(Accent);
        graphics.FillPolygon(arrow, new[]
        {
            new PointF(49, 8),
            new PointF(59, 11),
            new PointF(52, 19),
        });
        using var font = new Font("Segoe UI", 13f, FontStyle.Bold, GraphicsUnit.Point);
        using var brush = new SolidBrush(Ink);
        using var format = new StringFormat
        {
            Alignment = StringAlignment.Center,
            LineAlignment = StringAlignment.Center,
        };
        graphics.DrawString("1", font, brush, new RectangleF(17, 17, 30, 30), format);
    }

    private static void DrawEditSelected(Graphics graphics)
    {
        using var selection = Pen(Accent, 3f);
        selection.DashStyle = DashStyle.Dash;
        graphics.DrawRoundedRectangle(selection, new RectangleF(6, 10, 42, 38), 5f);
        DrawFormulaText(graphics, new RectangleF(9, 16, 35, 24), Ink, 11.5f);
        using var pencil = Pen(Warm, 6f, LineCap.Square);
        graphics.DrawLine(pencil, 34, 49, 55, 28);
        using var tip = new SolidBrush(Ink);
        graphics.FillPolygon(tip, new[]
        {
            new PointF(30, 53),
            new PointF(34, 44),
            new PointF(39, 49),
        });
    }

    private static void DrawConversion(Graphics graphics, bool toNative)
    {
        var left = toNative ? Warm : Accent;
        var right = toNative ? Accent : Warm;
        using var leftPen = Pen(left, 3f);
        graphics.DrawRoundedRectangle(leftPen, new RectangleF(5, 13, 20, 32), 4f);
        DrawFormulaText(graphics, new RectangleF(6, 18, 18, 20), left, 7.3f);
        using var arrow = Pen(Ink, 3.5f);
        graphics.DrawLine(arrow, 27, 29, 39, 29);
        graphics.DrawLine(arrow, 35, 24, 40, 29);
        graphics.DrawLine(arrow, 35, 34, 40, 29);
        using var rightPen = Pen(right, 3f);
        graphics.DrawRoundedRectangle(rightPen, new RectangleF(40, 13, 19, 32), 4f);
        DrawFormulaText(graphics, new RectangleF(41, 18, 17, 20), right, 7.3f);
    }

    private static void DrawRepairBaseline(Graphics graphics)
    {
        using var baseline = Pen(Ink, 3f);
        graphics.DrawLine(baseline, 7, 46, 57, 46);
        DrawFormulaText(graphics, new RectangleF(8, 23, 35, 25), Warm, 11.5f);
        using var arrow = Pen(Accent, 4f);
        graphics.DrawLine(arrow, 49, 43, 49, 16);
        graphics.DrawLine(arrow, 49, 16, 42, 24);
        graphics.DrawLine(arrow, 49, 16, 56, 24);
        using var reference = Pen(Accent, 2.2f);
        reference.DashStyle = DashStyle.Dash;
        graphics.DrawLine(reference, 8, 31, 38, 31);
    }

    private static void DrawBatchImport(Graphics graphics)
    {
        DrawDocument(graphics, new RectangleF(6, 6, 43, 52), Ink);
        using var line = Pen(Ink, 2.5f);
        graphics.DrawLine(line, 13, 17, 40, 17);
        graphics.DrawLine(line, 13, 25, 31, 25);
        DrawFormulaText(graphics, new RectangleF(27, 20, 16, 18), Accent, 6.8f);
        graphics.DrawLine(line, 13, 39, 40, 39);
        graphics.DrawLine(line, 13, 47, 31, 47);
        using var disk = new SolidBrush(Warm);
        graphics.FillEllipse(disk, 39, 34, 21, 21);
        using var arrow = Pen(Color.White, 3.2f);
        graphics.DrawLine(arrow, 49.5f, 38, 49.5f, 49);
        graphics.DrawLine(arrow, 44.5f, 45, 49.5f, 50);
        graphics.DrawLine(arrow, 54.5f, 45, 49.5f, 50);
    }

    private static void FillRoundedRectangle(
        this Graphics graphics,
        Brush brush,
        RectangleF rectangle,
        float radius)
    {
        using var path = RoundedRectangle(rectangle, radius);
        graphics.FillPath(brush, path);
    }

    private static void DrawRoundedRectangle(
        this Graphics graphics,
        Pen pen,
        RectangleF rectangle,
        float radius)
    {
        using var path = RoundedRectangle(rectangle, radius);
        graphics.DrawPath(pen, path);
    }

    private static GraphicsPath RoundedRectangle(RectangleF rectangle, float radius)
    {
        var diameter = Math.Min(radius * 2, Math.Min(rectangle.Width, rectangle.Height));
        var path = new GraphicsPath();
        path.AddArc(rectangle.Left, rectangle.Top, diameter, diameter, 180, 90);
        path.AddArc(rectangle.Right - diameter, rectangle.Top, diameter, diameter, 270, 90);
        path.AddArc(rectangle.Right - diameter, rectangle.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(rectangle.Left, rectangle.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        return path;
    }
}
