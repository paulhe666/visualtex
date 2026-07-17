using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace VisualTeX.WindowsOffice.VstoShared;

internal static class OfficeOlePreview
{
    private const long MaximumSvgBytes = 16L * 1024L * 1024L;
    private const int HorzRes = 8;
    private const int VertRes = 10;
    private const int DesktopVertRes = 117;
    private const int DesktopHorzRes = 118;
    internal static string LastRecordingDiagnostics { get; private set; } = string.Empty;

    [DllImport("gdi32.dll")]
    private static extern int GetDeviceCaps(IntPtr hdc, int index);
    private static readonly HashSet<uint> BitmapEmfRecords = new()
    {
        76,  // EMR_STRETCHBLT
        77,  // EMR_MASKBLT
        78,  // EMR_PLGBLT
        80,  // EMR_SETDIBITSTODEVICE
        81,  // EMR_STRETCHDIBITS
        114, // EMR_ALPHABLEND
        116, // EMR_TRANSPARENTBLT
    };

    public static string CreateVectorEmfFromSvg(
        string svgPath,
        float widthPixels,
        float heightPixels)
    {
        if (string.IsNullOrWhiteSpace(svgPath))
            throw new ArgumentException("SVG preview path is required.", nameof(svgPath));
        if (!File.Exists(svgPath))
            throw new FileNotFoundException("SVG preview does not exist.", svgPath);
        var information = new FileInfo(svgPath);
        if (information.Length <= 0 || information.Length > MaximumSvgBytes)
            throw new InvalidDataException("SVG preview size is invalid.");
        if (!IsPositiveFinite(widthPixels) || !IsPositiveFinite(heightPixels))
            throw new InvalidDataException("SVG preview dimensions are invalid.");

        var directory = Path.GetDirectoryName(svgPath)
            ?? throw new InvalidOperationException("SVG preview has no parent directory.");
        var emfPath = Path.Combine(directory, $"{Guid.NewGuid():N}.emf");
        try
        {
            var renderer = SvgVectorRenderer.Load(svgPath);
            renderer.Render(emfPath, widthPixels, heightPixels);
            ValidateVectorEmf(emfPath);
            return emfPath;
        }
        catch
        {
            try { File.Delete(emfPath); } catch { }
            throw;
        }
    }

    internal static void ValidateVectorEmf(string emfPath)
    {
        var bytes = File.ReadAllBytes(emfPath);
        if (bytes.Length < 88)
            throw new InvalidDataException("EMF preview is truncated.");

        var offset = 0;
        var sawVectorRecord = false;
        while (offset + 8 <= bytes.Length)
        {
            var recordType = ReadUInt32(bytes, offset);
            var recordSize = ReadUInt32(bytes, offset + 4);
            if (recordSize < 8 || recordSize > bytes.Length - offset)
                throw new InvalidDataException("EMF preview contains an invalid record size.");
            if (BitmapEmfRecords.Contains(recordType))
                throw new InvalidDataException(
                    $"EMF preview contains forbidden raster record type {recordType}.");

            if (recordType == 70) // EMR_GDICOMMENT; GDI+ records are carried here.
            {
                ValidateEmfPlusComment(bytes, offset, checked((int)recordSize), ref sawVectorRecord);
            }
            else if (recordType is 3 or 4 or 5 or 6 or 7 or 8 or 9 or 10 or 11 or 12
                     or 13 or 27 or 30 or 31 or 32 or 33 or 34 or 35 or 36 or 37
                     or 38 or 39 or 40 or 41 or 42 or 43 or 54 or 55 or 56 or 57
                     or 58 or 59 or 60 or 61 or 62 or 63 or 64 or 65 or 66 or 67
                     or 68 or 69 or 82 or 83 or 84 or 85 or 86 or 87 or 88 or 89
                     or 90 or 91 or 92 or 93 or 94 or 95 or 96 or 97 or 98 or 99)
            {
                sawVectorRecord = true;
            }

            offset += checked((int)recordSize);
            if (recordType == 14) break; // EMR_EOF
        }

        if (!sawVectorRecord)
            throw new InvalidDataException("EMF preview contains no vector drawing records.");
    }

    private static void ValidateEmfPlusComment(
        byte[] bytes,
        int recordOffset,
        int recordSize,
        ref bool sawVectorRecord)
    {
        if (recordSize < 16) return;
        var dataSize = checked((int)ReadUInt32(bytes, recordOffset + 8));
        if (dataSize < 4 || 12 + dataSize > recordSize) return;
        var dataOffset = recordOffset + 12;
        if (bytes[dataOffset] != (byte)'E'
            || bytes[dataOffset + 1] != (byte)'M'
            || bytes[dataOffset + 2] != (byte)'F'
            || bytes[dataOffset + 3] != (byte)'+')
            return;

        var cursor = dataOffset + 4;
        var end = dataOffset + dataSize;
        while (cursor + 12 <= end)
        {
            var type = ReadUInt16(bytes, cursor);
            var flags = ReadUInt16(bytes, cursor + 2);
            var size = ReadUInt32(bytes, cursor + 4);
            if (size < 12 || size > end - cursor)
                throw new InvalidDataException("EMF+ preview contains an invalid record size.");

            if (type is 0x401A or 0x401B) // DrawImage / DrawImagePoints
                throw new InvalidDataException("EMF+ preview contains a raster image draw record.");
            if (type == 0x4008 && ((flags >> 8) & 0x7F) == 5) // ObjectTypeImage
                throw new InvalidDataException("EMF+ preview embeds a raster image object.");
            if (type is 0x4014 or 0x4015 or 0x4016 or 0x4017 or 0x4018 or 0x4019
                        or 0x401C or 0x401D or 0x401E or 0x401F or 0x4020 or 0x4021)
                sawVectorRecord = true;

            cursor += checked((int)size);
        }
    }

    private static ushort ReadUInt16(byte[] bytes, int offset) =>
        (ushort)(bytes[offset] | (bytes[offset + 1] << 8));

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        (uint)(bytes[offset]
               | (bytes[offset + 1] << 8)
               | (bytes[offset + 2] << 16)
               | (bytes[offset + 3] << 24));

    private static bool IsPositiveFinite(float value) =>
        !float.IsNaN(value) && !float.IsInfinity(value) && value > 0;

    private sealed class SvgVectorRenderer
    {
        private static readonly XNamespace XLinkNamespace = "http://www.w3.org/1999/xlink";
        private static readonly Regex TransformPattern = new(
            @"([A-Za-z]+)\s*\(([^)]*)\)",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private readonly XElement _root;
        private readonly Dictionary<string, XElement> _definitions;
        private readonly SvgViewBox _viewBox;

        private SvgVectorRenderer(XElement root)
        {
            _root = root;
            _viewBox = ParseViewBox(root.Attribute("viewBox")?.Value);
            _definitions = root
                .Descendants()
                .Select(element => new { Element = element, Id = element.Attribute("id")?.Value })
                .Where(item => !string.IsNullOrWhiteSpace(item.Id))
                .ToDictionary(item => item.Id!, item => item.Element, StringComparer.Ordinal);
        }

        public static SvgVectorRenderer Load(string path)
        {
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                MaxCharactersInDocument = MaximumSvgBytes,
            };
            using var stream = File.OpenRead(path);
            using var reader = XmlReader.Create(stream, settings);
            var document = XDocument.Load(reader, LoadOptions.None);
            var root = document.Root
                ?? throw new InvalidDataException("SVG preview has no root element.");
            if (!string.Equals(root.Name.LocalName, "svg", StringComparison.Ordinal))
                throw new InvalidDataException("SVG preview root must be an svg element.");
            RejectUnsafeContent(root);
            return new SvgVectorRenderer(root);
        }

        public void Render(string emfPath, float widthPixels, float heightPixels)
        {
            // Use a real display DC as the EMF reference device. A memory DC
            // obtained from a 1x1 Bitmap records a mismatched physical-device
            // mapping on Windows: a 100x50 SVG frame replays as roughly 67x34
            // pixels in Office. The display HDC keeps CSS pixels, EMF frame
            // bounds, and Office's 96-DPI point conversion on the same scale.
            using var referenceGraphics = Graphics.FromHwnd(IntPtr.Zero);
            var referenceHdc = referenceGraphics.GetHdc();
            try
            {
                var logicalWidth = GetDeviceCaps(referenceHdc, HorzRes);
                var logicalHeight = GetDeviceCaps(referenceHdc, VertRes);
                var desktopWidth = GetDeviceCaps(referenceHdc, DesktopHorzRes);
                var desktopHeight = GetDeviceCaps(referenceHdc, DesktopVertRes);
                var deviceScaleX = logicalWidth > 0 && desktopWidth > 0
                    ? desktopWidth / (double)logicalWidth
                    : 1d;
                var deviceScaleY = logicalHeight > 0 && desktopHeight > 0
                    ? desktopHeight / (double)logicalHeight
                    : 1d;
                if (!IsPositiveFinite(deviceScaleX)) deviceScaleX = 1d;
                if (!IsPositiveFinite(deviceScaleY)) deviceScaleY = 1d;
                using var metafile = new Metafile(
                    emfPath,
                    referenceHdc,
                    new RectangleF(0, 0, widthPixels, heightPixels),
                    MetafileFrameUnit.Pixel,
                    EmfType.EmfOnly,
                    "VisualTeX vector formula preview");
                using var graphics = Graphics.FromImage(metafile);
                // A Metafile recording Graphics defaults to GraphicsUnit.Display
                // on some Office/DPI configurations. The SVG transform below is
                // expressed in CSS pixels, so leaving that default records every
                // glyph smaller than the EMF pixel frame (and PowerPoint then
                // shows a shrunken formula inside a correctly sized OLE box).
                // Lock both the EMF frame and all vector coordinates to pixels.
                graphics.PageUnit = GraphicsUnit.Pixel;
                graphics.PageScale = 1f;
                graphics.ResetTransform();
                graphics.SmoothingMode = SmoothingMode.None;
                graphics.PixelOffsetMode = PixelOffsetMode.None;
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.Default;

                // On a scaled Windows desktop, Graphics.FromImage(metafile)
                // can expose a logical recording canvas larger than the pixel
                // frame passed to the Metafile constructor (for example 150x75
                // logical units for a requested 100x50 frame). Rendering against
                // the requested pixel dimensions then occupies only two thirds
                // of the EMF. Map the SVG viewBox to the actual recording canvas;
                // Office will replay that canvas into the requested physical frame.
                LastRecordingDiagnostics = string.Format(
                    CultureInfo.InvariantCulture,
                    "DPI={0:0.###}x{1:0.###}; PageUnit={2}; PageScale={3:0.###}; logical={4}x{5}; desktop={6}x{7}; deviceScale={8:0.###}x{9:0.###}; requested={10:0.###}x{11:0.###}",
                    graphics.DpiX,
                    graphics.DpiY,
                    graphics.PageUnit,
                    graphics.PageScale,
                    logicalWidth,
                    logicalHeight,
                    desktopWidth,
                    desktopHeight,
                    deviceScaleX,
                    deviceScaleY,
                    widthPixels,
                    heightPixels);
                var recordingWidth = widthPixels * deviceScaleX;
                var recordingHeight = heightPixels * deviceScaleY;
                var rootTransform = new SvgMatrix(
                    recordingWidth / _viewBox.Width,
                    0,
                    0,
                    recordingHeight / _viewBox.Height,
                    -_viewBox.X * recordingWidth / _viewBox.Width,
                    -_viewBox.Y * recordingHeight / _viewBox.Height);
                var style = SvgStyle.Default;
                foreach (var child in _root.Elements())
                    RenderElement(graphics, child, rootTransform, style, false, new HashSet<string>());
            }
            finally
            {
                referenceGraphics.ReleaseHdc(referenceHdc);
            }

            if (!File.Exists(emfPath) || new FileInfo(emfPath).Length == 0)
                throw new InvalidDataException("Vector EMF generation produced an empty file.");
        }

        private void RenderElement(
            Graphics graphics,
            XElement element,
            SvgMatrix parentTransform,
            SvgStyle parentStyle,
            bool definitionReference,
            HashSet<string> referenceStack)
        {
            var name = element.Name.LocalName;
            if (name == "defs" && !definitionReference) return;
            if (name is "title" or "desc" or "metadata") return;

            var transform = parentTransform.Multiply(ParseTransform(element.Attribute("transform")?.Value));
            var style = SvgStyle.Resolve(element, parentStyle);
            switch (name)
            {
                case "svg":
                case "g":
                    foreach (var child in element.Elements())
                        RenderElement(graphics, child, transform, style, definitionReference, referenceStack);
                    return;
                case "use":
                    RenderUse(graphics, element, transform, style, referenceStack);
                    return;
                case "path":
                    RenderPath(graphics, element, transform, style);
                    return;
                case "rect":
                    RenderRectangle(graphics, element, transform, style);
                    return;
                case "line":
                    RenderLine(graphics, element, transform, style);
                    return;
                case "polyline":
                case "polygon":
                    RenderPoly(graphics, element, transform, style, name == "polygon");
                    return;
                case "circle":
                case "ellipse":
                    RenderEllipse(graphics, element, transform, style, name == "circle");
                    return;
                default:
                    throw new InvalidDataException(
                        $"SVG element <{name}> is not supported by the native vector EMF renderer.");
            }
        }

        private void RenderUse(
            Graphics graphics,
            XElement use,
            SvgMatrix transform,
            SvgStyle style,
            HashSet<string> referenceStack)
        {
            var href = use.Attribute("href")?.Value
                       ?? use.Attribute(XLinkNamespace + "href")?.Value;
            if (string.IsNullOrWhiteSpace(href))
                throw new InvalidDataException("SVG use element must reference a local definition.");
            var reference = href!;
            if (!reference.StartsWith("#", StringComparison.Ordinal))
                throw new InvalidDataException("SVG use element must reference a local definition.");
            var id = reference.Substring(1);
            if (!_definitions.TryGetValue(id, out var referenced))
                throw new InvalidDataException($"SVG definition '{id}' does not exist.");
            if (!referenceStack.Add(id))
                throw new InvalidDataException("SVG contains a cyclic definition reference.");
            try
            {
                var x = ReadNumber(use.Attribute("x")?.Value, 0);
                var y = ReadNumber(use.Attribute("y")?.Value, 0);
                var useTransform = transform.Multiply(SvgMatrix.Translate(x, y));
                RenderElement(graphics, referenced, useTransform, style, true, referenceStack);
            }
            finally
            {
                referenceStack.Remove(id);
            }
        }

        private static void RenderPath(
            Graphics graphics,
            XElement element,
            SvgMatrix transform,
            SvgStyle style)
        {
            var data = element.Attribute("d")?.Value;
            if (string.IsNullOrWhiteSpace(data)) return;
            using var path = SvgPathParser.Parse(data!);
            path.FillMode = string.Equals(
                element.Attribute("fill-rule")?.Value,
                "evenodd",
                StringComparison.OrdinalIgnoreCase)
                ? FillMode.Alternate
                : FillMode.Winding;
            ApplyTransform(path, transform);
            PaintPath(graphics, path, style, transform);
        }

        private static void RenderRectangle(
            Graphics graphics,
            XElement element,
            SvgMatrix transform,
            SvgStyle style)
        {
            var x = ReadNumber(element.Attribute("x")?.Value, 0);
            var y = ReadNumber(element.Attribute("y")?.Value, 0);
            var width = ReadNumber(element.Attribute("width")?.Value, double.NaN);
            var height = ReadNumber(element.Attribute("height")?.Value, double.NaN);
            if (!IsPositiveFinite(width) || !IsPositiveFinite(height))
                throw new InvalidDataException("SVG rectangle dimensions are invalid.");
            using var path = new GraphicsPath();
            path.AddRectangle(new RectangleF((float)x, (float)y, (float)width, (float)height));
            ApplyTransform(path, transform);
            PaintPath(graphics, path, style, transform);
        }

        private static void RenderLine(
            Graphics graphics,
            XElement element,
            SvgMatrix transform,
            SvgStyle style)
        {
            using var path = new GraphicsPath();
            path.AddLine(
                (float)ReadNumber(element.Attribute("x1")?.Value, 0),
                (float)ReadNumber(element.Attribute("y1")?.Value, 0),
                (float)ReadNumber(element.Attribute("x2")?.Value, 0),
                (float)ReadNumber(element.Attribute("y2")?.Value, 0));
            ApplyTransform(path, transform);
            PaintPath(graphics, path, style.WithoutFill(), transform);
        }

        private static void RenderPoly(
            Graphics graphics,
            XElement element,
            SvgMatrix transform,
            SvgStyle style,
            bool close)
        {
            var values = ParseNumberList(element.Attribute("points")?.Value ?? string.Empty);
            if (values.Count < 4 || values.Count % 2 != 0)
                throw new InvalidDataException("SVG poly points are invalid.");
            var points = new PointF[values.Count / 2];
            for (var index = 0; index < points.Length; index++)
                points[index] = new PointF((float)values[index * 2], (float)values[index * 2 + 1]);
            using var path = new GraphicsPath();
            if (close) path.AddPolygon(points);
            else path.AddLines(points);
            ApplyTransform(path, transform);
            PaintPath(graphics, path, close ? style : style.WithoutFill(), transform);
        }

        private static void RenderEllipse(
            Graphics graphics,
            XElement element,
            SvgMatrix transform,
            SvgStyle style,
            bool circle)
        {
            var cx = ReadNumber(element.Attribute("cx")?.Value, 0);
            var cy = ReadNumber(element.Attribute("cy")?.Value, 0);
            var rx = circle
                ? ReadNumber(element.Attribute("r")?.Value, double.NaN)
                : ReadNumber(element.Attribute("rx")?.Value, double.NaN);
            var ry = circle
                ? rx
                : ReadNumber(element.Attribute("ry")?.Value, double.NaN);
            if (!IsPositiveFinite(rx) || !IsPositiveFinite(ry))
                throw new InvalidDataException("SVG ellipse dimensions are invalid.");
            using var path = new GraphicsPath();
            path.AddEllipse(
                (float)(cx - rx),
                (float)(cy - ry),
                (float)(rx * 2),
                (float)(ry * 2));
            ApplyTransform(path, transform);
            PaintPath(graphics, path, style, transform);
        }

        private static void PaintPath(
            Graphics graphics,
            GraphicsPath path,
            SvgStyle style,
            SvgMatrix transform)
        {
            var fill = style.ResolveFillColor();
            if (fill.HasValue && fill.Value.A > 0)
            {
                using var brush = new SolidBrush(fill.Value);
                graphics.FillPath(brush, path);
            }

            var stroke = style.ResolveStrokeColor();
            if (stroke.HasValue && stroke.Value.A > 0 && style.StrokeWidth > 0)
            {
                var scaleX = Math.Sqrt(transform.A * transform.A + transform.B * transform.B);
                var scaleY = Math.Sqrt(transform.C * transform.C + transform.D * transform.D);
                var scale = Math.Max(0.0001, (scaleX + scaleY) / 2);
                using var pen = new Pen(stroke.Value, (float)(style.StrokeWidth * scale))
                {
                    LineJoin = LineJoin.Miter,
                    StartCap = LineCap.Flat,
                    EndCap = LineCap.Flat,
                };
                graphics.DrawPath(pen, path);
            }
        }

        private static void ApplyTransform(GraphicsPath path, SvgMatrix transform)
        {
            using var matrix = new Matrix(
                (float)transform.A,
                (float)transform.B,
                (float)transform.C,
                (float)transform.D,
                (float)transform.E,
                (float)transform.F);
            path.Transform(matrix);
        }

        private static SvgMatrix ParseTransform(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return SvgMatrix.Identity;
            var transformText = value!;
            var result = SvgMatrix.Identity;
            var cursor = 0;
            foreach (Match match in TransformPattern.Matches(transformText))
            {
                if (!OnlySeparators(transformText, cursor, match.Index - cursor))
                    throw new InvalidDataException("SVG transform syntax is invalid.");
                cursor = match.Index + match.Length;
                var name = match.Groups[1].Value;
                var values = ParseNumberList(match.Groups[2].Value);
                SvgMatrix operation;
                switch (name)
                {
                    case "matrix" when values.Count == 6:
                        operation = new SvgMatrix(
                            values[0], values[1], values[2], values[3], values[4], values[5]);
                        break;
                    case "translate" when values.Count is 1 or 2:
                        operation = SvgMatrix.Translate(values[0], values.Count == 2 ? values[1] : 0);
                        break;
                    case "scale" when values.Count is 1 or 2:
                        operation = SvgMatrix.Scale(values[0], values.Count == 2 ? values[1] : values[0]);
                        break;
                    case "rotate" when values.Count == 1:
                        operation = SvgMatrix.Rotate(values[0]);
                        break;
                    case "rotate" when values.Count == 3:
                        operation = SvgMatrix.Translate(values[1], values[2])
                            .Multiply(SvgMatrix.Rotate(values[0]))
                            .Multiply(SvgMatrix.Translate(-values[1], -values[2]));
                        break;
                    case "skewX" when values.Count == 1:
                        operation = SvgMatrix.SkewX(values[0]);
                        break;
                    case "skewY" when values.Count == 1:
                        operation = SvgMatrix.SkewY(values[0]);
                        break;
                    default:
                        throw new InvalidDataException($"Unsupported SVG transform '{name}'.");
                }
                result = result.Multiply(operation);
            }
            if (!OnlySeparators(transformText, cursor, transformText.Length - cursor))
                throw new InvalidDataException("SVG transform syntax is invalid.");
            return result;
        }

        private static bool OnlySeparators(string value, int start, int length)
        {
            for (var index = start; index < start + length; index++)
            {
                var character = value[index];
                if (!char.IsWhiteSpace(character) && character != ',') return false;
            }
            return true;
        }

        private static List<double> ParseNumberList(string value)
        {
            var reader = new SvgNumberReader(value);
            var result = new List<double>();
            while (reader.TryReadNumber(out var number)) result.Add(number);
            if (!reader.AtEnd)
                throw new InvalidDataException("SVG numeric list is invalid.");
            return result;
        }

        private static double ReadNumber(string? value, double fallback)
        {
            if (string.IsNullOrWhiteSpace(value)) return fallback;
            var normalized = value!.Trim();
            if (normalized.EndsWith("px", StringComparison.OrdinalIgnoreCase))
                normalized = normalized.Substring(0, normalized.Length - 2);
            return double.TryParse(
                normalized,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out var number)
                && IsFinite(number)
                ? number
                : throw new InvalidDataException($"Invalid SVG number '{value}'.");
        }

        private static SvgViewBox ParseViewBox(string? value)
        {
            var values = ParseNumberList(value ?? string.Empty);
            if (values.Count != 4
                || !IsFinite(values[0])
                || !IsFinite(values[1])
                || !IsPositiveFinite(values[2])
                || !IsPositiveFinite(values[3]))
                throw new InvalidDataException("SVG preview has an invalid viewBox.");
            return new SvgViewBox(values[0], values[1], values[2], values[3]);
        }

        private static void RejectUnsafeContent(XElement root)
        {
            foreach (var element in root.DescendantsAndSelf())
            {
                var name = element.Name.LocalName;
                if (name is "image" or "foreignObject" or "script" or "style" or "link")
                    throw new InvalidDataException($"SVG element <{name}> is forbidden.");
                foreach (var attribute in element.Attributes())
                {
                    if (attribute.IsNamespaceDeclaration) continue;
                    var attributeName = attribute.Name.LocalName;
                    var attributeValue = attribute.Value;
                    if (attributeName.StartsWith("on", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidDataException("SVG event handler attributes are forbidden.");
                    if (attributeName is "href" && !attributeValue.StartsWith("#", StringComparison.Ordinal))
                        throw new InvalidDataException("SVG external references are forbidden.");
                    if (attributeValue.IndexOf("url(", StringComparison.OrdinalIgnoreCase) >= 0)
                        throw new InvalidDataException("SVG URL paint servers are not supported.");
                }
            }
        }

        private static bool IsFinite(double value) =>
            !double.IsNaN(value) && !double.IsInfinity(value);

        private static bool IsPositiveFinite(double value) => IsFinite(value) && value > 0;
    }

    private readonly struct SvgViewBox
    {
        public SvgViewBox(double x, double y, double width, double height)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        public double X { get; }
        public double Y { get; }
        public double Width { get; }
        public double Height { get; }
    }

    private readonly struct SvgMatrix
    {
        public static readonly SvgMatrix Identity = new(1, 0, 0, 1, 0, 0);

        public SvgMatrix(double a, double b, double c, double d, double e, double f)
        {
            A = a;
            B = b;
            C = c;
            D = d;
            E = e;
            F = f;
        }

        public double A { get; }
        public double B { get; }
        public double C { get; }
        public double D { get; }
        public double E { get; }
        public double F { get; }

        public SvgMatrix Multiply(SvgMatrix other) => new(
            A * other.A + C * other.B,
            B * other.A + D * other.B,
            A * other.C + C * other.D,
            B * other.C + D * other.D,
            A * other.E + C * other.F + E,
            B * other.E + D * other.F + F);

        public static SvgMatrix Translate(double x, double y) => new(1, 0, 0, 1, x, y);
        public static SvgMatrix Scale(double x, double y) => new(x, 0, 0, y, 0, 0);

        public static SvgMatrix Rotate(double degrees)
        {
            var radians = degrees * Math.PI / 180;
            var cosine = Math.Cos(radians);
            var sine = Math.Sin(radians);
            return new SvgMatrix(cosine, sine, -sine, cosine, 0, 0);
        }

        public static SvgMatrix SkewX(double degrees) =>
            new(1, 0, Math.Tan(degrees * Math.PI / 180), 1, 0, 0);

        public static SvgMatrix SkewY(double degrees) =>
            new(1, Math.Tan(degrees * Math.PI / 180), 0, 1, 0, 0);
    }

    private readonly struct SvgStyle
    {
        public static readonly SvgStyle Default = new(
            "#000000",
            "none",
            "#000000",
            1,
            1,
            1,
            1);

        private SvgStyle(
            string fill,
            string stroke,
            string currentColor,
            double strokeWidth,
            double fillOpacity,
            double strokeOpacity,
            double opacity)
        {
            Fill = fill;
            Stroke = stroke;
            CurrentColor = currentColor;
            StrokeWidth = strokeWidth;
            FillOpacity = fillOpacity;
            StrokeOpacity = strokeOpacity;
            Opacity = opacity;
        }

        public string Fill { get; }
        public string Stroke { get; }
        public string CurrentColor { get; }
        public double StrokeWidth { get; }
        public double FillOpacity { get; }
        public double StrokeOpacity { get; }
        public double Opacity { get; }

        public static SvgStyle Resolve(XElement element, SvgStyle parent)
        {
            var declarations = ParseStyleDeclarations(element.Attribute("style")?.Value);
            string Attribute(string name, string fallback) =>
                element.Attribute(name)?.Value
                ?? (declarations.TryGetValue(name, out var value) ? value : fallback);
            double Numeric(string name, double fallback) =>
                ParseStyleNumber(Attribute(name, fallback.ToString(CultureInfo.InvariantCulture)), fallback);

            return new SvgStyle(
                Attribute("fill", parent.Fill),
                Attribute("stroke", parent.Stroke),
                Attribute("color", parent.CurrentColor),
                Math.Max(0, Numeric("stroke-width", parent.StrokeWidth)),
                Clamp01(Numeric("fill-opacity", parent.FillOpacity)),
                Clamp01(Numeric("stroke-opacity", parent.StrokeOpacity)),
                Clamp01(parent.Opacity * Numeric("opacity", 1)));
        }

        public SvgStyle WithoutFill() =>
            new("none", Stroke, CurrentColor, StrokeWidth, FillOpacity, StrokeOpacity, Opacity);

        public Color? ResolveFillColor() =>
            ResolveColor(ResolveCurrentColor(Fill), FillOpacity * Opacity);

        public Color? ResolveStrokeColor() =>
            ResolveColor(ResolveCurrentColor(Stroke), StrokeOpacity * Opacity);

        private string ResolveCurrentColor(string value) =>
            string.Equals(value, "currentColor", StringComparison.OrdinalIgnoreCase)
                ? CurrentColor
                : value;

        private static Dictionary<string, string> ParseStyleDeclarations(string? value)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(value)) return result;
            foreach (var declaration in value!.Split(';'))
            {
                var separator = declaration.IndexOf(':');
                if (separator <= 0) continue;
                result[declaration.Substring(0, separator).Trim()] =
                    declaration.Substring(separator + 1).Trim();
            }
            return result;
        }

        private static double ParseStyleNumber(string value, double fallback)
        {
            var normalized = value.Trim();
            if (normalized.EndsWith("px", StringComparison.OrdinalIgnoreCase))
                normalized = normalized.Substring(0, normalized.Length - 2);
            return double.TryParse(
                normalized,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out var number)
                && !double.IsNaN(number)
                && !double.IsInfinity(number)
                ? number
                : fallback;
        }

        private static Color? ResolveColor(string value, double opacity)
        {
            if (string.IsNullOrWhiteSpace(value)
                || string.Equals(value, "none", StringComparison.OrdinalIgnoreCase))
                return null;
            var normalized = value.Trim();
            Color color;
            if (normalized.StartsWith("#", StringComparison.Ordinal))
            {
                var hex = normalized.Substring(1);
                if (hex.Length == 3)
                {
                    color = Color.FromArgb(
                        Convert.ToInt32(new string(hex[0], 2), 16),
                        Convert.ToInt32(new string(hex[1], 2), 16),
                        Convert.ToInt32(new string(hex[2], 2), 16));
                }
                else if (hex.Length is 6 or 8)
                {
                    var red = Convert.ToInt32(hex.Substring(0, 2), 16);
                    var green = Convert.ToInt32(hex.Substring(2, 2), 16);
                    var blue = Convert.ToInt32(hex.Substring(4, 2), 16);
                    var embeddedAlpha = hex.Length == 8
                        ? Convert.ToInt32(hex.Substring(6, 2), 16) / 255.0
                        : 1;
                    color = Color.FromArgb(red, green, blue);
                    opacity *= embeddedAlpha;
                }
                else
                {
                    throw new InvalidDataException($"Unsupported SVG color '{value}'.");
                }
            }
            else if (normalized.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase)
                     && normalized.EndsWith(")", StringComparison.Ordinal))
            {
                var components = normalized.Substring(4, normalized.Length - 5)
                    .Split(',')
                    .Select(component => int.Parse(component.Trim(), CultureInfo.InvariantCulture))
                    .ToArray();
                if (components.Length != 3 || components.Any(component => component < 0 || component > 255))
                    throw new InvalidDataException($"Unsupported SVG color '{value}'.");
                color = Color.FromArgb(components[0], components[1], components[2]);
            }
            else
            {
                color = Color.FromName(normalized);
                if (!color.IsKnownColor && !color.IsNamedColor)
                    throw new InvalidDataException($"Unsupported SVG color '{value}'.");
            }

            if (opacity <= 0.01) return Color.FromArgb(0, color.R, color.G, color.B);
            if (opacity < 0.999)
                throw new InvalidDataException(
                    "Semi-transparent SVG paint cannot be represented as a true vector EMF.");
            return Color.FromArgb(255, color.R, color.G, color.B);
        }

        private static double Clamp01(double value) => Math.Max(0, Math.Min(1, value));
    }

    private sealed class SvgNumberReader
    {
        private readonly string _value;
        private int _index;

        public SvgNumberReader(string value) => _value = value;

        public bool AtEnd
        {
            get
            {
                SkipSeparators();
                return _index >= _value.Length;
            }
        }

        public bool TryReadNumber(out double number)
        {
            SkipSeparators();
            number = 0;
            if (_index >= _value.Length) return false;
            var start = _index;
            if (_value[_index] is '+' or '-') _index++;
            var digits = 0;
            while (_index < _value.Length && char.IsDigit(_value[_index]))
            {
                _index++;
                digits++;
            }
            if (_index < _value.Length && _value[_index] == '.')
            {
                _index++;
                while (_index < _value.Length && char.IsDigit(_value[_index]))
                {
                    _index++;
                    digits++;
                }
            }
            if (digits == 0)
            {
                _index = start;
                return false;
            }
            if (_index < _value.Length && _value[_index] is 'e' or 'E')
            {
                var exponentStart = _index;
                _index++;
                if (_index < _value.Length && _value[_index] is '+' or '-') _index++;
                var exponentDigits = 0;
                while (_index < _value.Length && char.IsDigit(_value[_index]))
                {
                    _index++;
                    exponentDigits++;
                }
                if (exponentDigits == 0) _index = exponentStart;
            }
            var token = _value.Substring(start, _index - start);
            if (!double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out number)
                || double.IsNaN(number)
                || double.IsInfinity(number))
                throw new InvalidDataException($"Invalid SVG number '{token}'.");
            return true;
        }

        private void SkipSeparators()
        {
            while (_index < _value.Length
                   && (char.IsWhiteSpace(_value[_index]) || _value[_index] == ','))
                _index++;
        }
    }

    private sealed class SvgPathParser
    {
        private readonly string _data;
        private int _index;
        private PointF _current;
        private PointF _figureStart;
        private PointF? _lastCubicControl;
        private PointF? _lastQuadraticControl;
        private char _previousCommand;
        private readonly GraphicsPath _path = new();

        private SvgPathParser(string data) => _data = data;

        public static GraphicsPath Parse(string data)
        {
            var parser = new SvgPathParser(data);
            try
            {
                parser.ParseAll();
                return parser._path;
            }
            catch
            {
                parser._path.Dispose();
                throw;
            }
        }

        private void ParseAll()
        {
            char command = '\0';
            while (true)
            {
                SkipSeparators();
                if (_index >= _data.Length) break;
                if (char.IsLetter(_data[_index])) command = _data[_index++];
                else if (command == '\0') throw InvalidPath();

                var relative = char.IsLower(command);
                var normalized = char.ToUpperInvariant(command);
                switch (normalized)
                {
                    case 'M': ParseMove(relative, ref command); break;
                    case 'L': ParseLine(relative); break;
                    case 'H': ParseHorizontal(relative); break;
                    case 'V': ParseVertical(relative); break;
                    case 'C': ParseCubic(relative); break;
                    case 'S': ParseSmoothCubic(relative); break;
                    case 'Q': ParseQuadratic(relative); break;
                    case 'T': ParseSmoothQuadratic(relative); break;
                    case 'A': ParseArc(relative); break;
                    case 'Z':
                        _path.CloseFigure();
                        _current = _figureStart;
                        ResetControls();
                        _previousCommand = 'Z';
                        command = '\0';
                        break;
                    default: throw InvalidPath();
                }
            }
        }

        private void ParseMove(bool relative, ref char command)
        {
            var first = ReadPoint(relative);
            _path.StartFigure();
            _current = first;
            _figureStart = first;
            ResetControls();
            _previousCommand = 'M';
            var lineCommand = relative ? 'l' : 'L';
            while (HasNumber())
            {
                var point = ReadPoint(relative);
                _path.AddLine(_current, point);
                _current = point;
                _previousCommand = 'L';
            }
            command = lineCommand;
        }

        private void ParseLine(bool relative)
        {
            RequireNumber();
            do
            {
                var point = ReadPoint(relative);
                _path.AddLine(_current, point);
                _current = point;
            } while (HasNumber());
            ResetControls();
            _previousCommand = 'L';
        }

        private void ParseHorizontal(bool relative)
        {
            RequireNumber();
            do
            {
                var x = ReadNumber();
                var point = new PointF((float)(relative ? _current.X + x : x), _current.Y);
                _path.AddLine(_current, point);
                _current = point;
            } while (HasNumber());
            ResetControls();
            _previousCommand = 'H';
        }

        private void ParseVertical(bool relative)
        {
            RequireNumber();
            do
            {
                var y = ReadNumber();
                var point = new PointF(_current.X, (float)(relative ? _current.Y + y : y));
                _path.AddLine(_current, point);
                _current = point;
            } while (HasNumber());
            ResetControls();
            _previousCommand = 'V';
        }

        private void ParseCubic(bool relative)
        {
            RequireNumber();
            do
            {
                var control1 = ReadPoint(relative);
                var control2 = ReadPoint(relative);
                var end = ReadPoint(relative);
                _path.AddBezier(_current, control1, control2, end);
                _current = end;
                _lastCubicControl = control2;
                _lastQuadraticControl = null;
                _previousCommand = 'C';
            } while (HasNumber());
        }

        private void ParseSmoothCubic(bool relative)
        {
            RequireNumber();
            do
            {
                var control1 = _previousCommand is 'C' or 'S' && _lastCubicControl.HasValue
                    ? Reflect(_lastCubicControl.Value, _current)
                    : _current;
                var control2 = ReadPoint(relative);
                var end = ReadPoint(relative);
                _path.AddBezier(_current, control1, control2, end);
                _current = end;
                _lastCubicControl = control2;
                _lastQuadraticControl = null;
                _previousCommand = 'S';
            } while (HasNumber());
        }

        private void ParseQuadratic(bool relative)
        {
            RequireNumber();
            do
            {
                var control = ReadPoint(relative);
                var end = ReadPoint(relative);
                AddQuadratic(_current, control, end);
                _current = end;
                _lastQuadraticControl = control;
                _lastCubicControl = null;
                _previousCommand = 'Q';
            } while (HasNumber());
        }

        private void ParseSmoothQuadratic(bool relative)
        {
            RequireNumber();
            do
            {
                var control = _previousCommand is 'Q' or 'T' && _lastQuadraticControl.HasValue
                    ? Reflect(_lastQuadraticControl.Value, _current)
                    : _current;
                var end = ReadPoint(relative);
                AddQuadratic(_current, control, end);
                _current = end;
                _lastQuadraticControl = control;
                _lastCubicControl = null;
                _previousCommand = 'T';
            } while (HasNumber());
        }

        private void ParseArc(bool relative)
        {
            RequireNumber();
            do
            {
                var rx = Math.Abs(ReadNumber());
                var ry = Math.Abs(ReadNumber());
                var rotation = ReadNumber();
                var largeArc = ReadFlag();
                var sweep = ReadFlag();
                var end = ReadPoint(relative);
                AddArc(_current, end, rx, ry, rotation, largeArc, sweep);
                _current = end;
                ResetControls();
                _previousCommand = 'A';
            } while (HasNumber());
        }

        private void AddQuadratic(PointF start, PointF control, PointF end)
        {
            var control1 = new PointF(
                start.X + (control.X - start.X) * 2f / 3f,
                start.Y + (control.Y - start.Y) * 2f / 3f);
            var control2 = new PointF(
                end.X + (control.X - end.X) * 2f / 3f,
                end.Y + (control.Y - end.Y) * 2f / 3f);
            _path.AddBezier(start, control1, control2, end);
        }

        private void AddArc(
            PointF start,
            PointF end,
            double rx,
            double ry,
            double rotationDegrees,
            bool largeArc,
            bool sweep)
        {
            if (rx == 0 || ry == 0 || (start.X == end.X && start.Y == end.Y))
            {
                _path.AddLine(start, end);
                return;
            }

            var phi = rotationDegrees * Math.PI / 180;
            var cosPhi = Math.Cos(phi);
            var sinPhi = Math.Sin(phi);
            var dx = (start.X - end.X) / 2.0;
            var dy = (start.Y - end.Y) / 2.0;
            var xPrime = cosPhi * dx + sinPhi * dy;
            var yPrime = -sinPhi * dx + cosPhi * dy;
            var lambda = xPrime * xPrime / (rx * rx) + yPrime * yPrime / (ry * ry);
            if (lambda > 1)
            {
                var scale = Math.Sqrt(lambda);
                rx *= scale;
                ry *= scale;
            }

            var numerator = rx * rx * ry * ry
                            - rx * rx * yPrime * yPrime
                            - ry * ry * xPrime * xPrime;
            var denominator = rx * rx * yPrime * yPrime
                              + ry * ry * xPrime * xPrime;
            var coefficient = denominator == 0
                ? 0
                : (largeArc == sweep ? -1 : 1)
                  * Math.Sqrt(Math.Max(0, numerator / denominator));
            var centerPrimeX = coefficient * (rx * yPrime / ry);
            var centerPrimeY = coefficient * (-ry * xPrime / rx);
            var centerX = cosPhi * centerPrimeX - sinPhi * centerPrimeY
                          + (start.X + end.X) / 2.0;
            var centerY = sinPhi * centerPrimeX + cosPhi * centerPrimeY
                          + (start.Y + end.Y) / 2.0;

            var startVectorX = (xPrime - centerPrimeX) / rx;
            var startVectorY = (yPrime - centerPrimeY) / ry;
            var endVectorX = (-xPrime - centerPrimeX) / rx;
            var endVectorY = (-yPrime - centerPrimeY) / ry;
            var startAngle = Math.Atan2(startVectorY, startVectorX);
            var delta = VectorAngle(startVectorX, startVectorY, endVectorX, endVectorY);
            if (!sweep && delta > 0) delta -= Math.PI * 2;
            if (sweep && delta < 0) delta += Math.PI * 2;

            var segments = Math.Max(1, (int)Math.Ceiling(Math.Abs(delta) / (Math.PI / 2)));
            var segmentDelta = delta / segments;
            var current = start;
            for (var segment = 0; segment < segments; segment++)
            {
                var angle1 = startAngle + segment * segmentDelta;
                var angle2 = angle1 + segmentDelta;
                var alpha = 4.0 / 3.0 * Math.Tan((angle2 - angle1) / 4.0);
                var point1 = EllipsePoint(centerX, centerY, rx, ry, cosPhi, sinPhi, angle1);
                var point2 = EllipsePoint(centerX, centerY, rx, ry, cosPhi, sinPhi, angle2);
                var derivative1 = EllipseDerivative(rx, ry, cosPhi, sinPhi, angle1);
                var derivative2 = EllipseDerivative(rx, ry, cosPhi, sinPhi, angle2);
                var control1 = new PointF(
                    (float)(point1.X + alpha * derivative1.X),
                    (float)(point1.Y + alpha * derivative1.Y));
                var control2 = new PointF(
                    (float)(point2.X - alpha * derivative2.X),
                    (float)(point2.Y - alpha * derivative2.Y));
                var finalPoint = segment == segments - 1 ? end : point2;
                _path.AddBezier(current, control1, control2, finalPoint);
                current = finalPoint;
            }
        }

        private static PointF EllipsePoint(
            double centerX,
            double centerY,
            double rx,
            double ry,
            double cosPhi,
            double sinPhi,
            double angle) =>
            new(
                (float)(centerX + rx * cosPhi * Math.Cos(angle) - ry * sinPhi * Math.Sin(angle)),
                (float)(centerY + rx * sinPhi * Math.Cos(angle) + ry * cosPhi * Math.Sin(angle)));

        private static PointF EllipseDerivative(
            double rx,
            double ry,
            double cosPhi,
            double sinPhi,
            double angle) =>
            new(
                (float)(-rx * cosPhi * Math.Sin(angle) - ry * sinPhi * Math.Cos(angle)),
                (float)(-rx * sinPhi * Math.Sin(angle) + ry * cosPhi * Math.Cos(angle)));

        private static double VectorAngle(double ux, double uy, double vx, double vy) =>
            Math.Atan2(ux * vy - uy * vx, ux * vx + uy * vy);

        private PointF ReadPoint(bool relative)
        {
            var x = ReadNumber();
            var y = ReadNumber();
            return new PointF(
                (float)(relative ? _current.X + x : x),
                (float)(relative ? _current.Y + y : y));
        }

        private double ReadNumber()
        {
            SkipSeparators();
            var start = _index;
            if (_index < _data.Length && _data[_index] is '+' or '-') _index++;
            var digits = 0;
            while (_index < _data.Length && char.IsDigit(_data[_index]))
            {
                _index++;
                digits++;
            }
            if (_index < _data.Length && _data[_index] == '.')
            {
                _index++;
                while (_index < _data.Length && char.IsDigit(_data[_index]))
                {
                    _index++;
                    digits++;
                }
            }
            if (digits == 0) throw InvalidPath();
            if (_index < _data.Length && _data[_index] is 'e' or 'E')
            {
                _index++;
                if (_index < _data.Length && _data[_index] is '+' or '-') _index++;
                var exponentDigits = 0;
                while (_index < _data.Length && char.IsDigit(_data[_index]))
                {
                    _index++;
                    exponentDigits++;
                }
                if (exponentDigits == 0) throw InvalidPath();
            }
            var token = _data.Substring(start, _index - start);
            return double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
                   && !double.IsNaN(value)
                   && !double.IsInfinity(value)
                ? value
                : throw InvalidPath();
        }

        private bool ReadFlag()
        {
            var value = ReadNumber();
            if (value == 0) return false;
            if (value == 1) return true;
            throw InvalidPath();
        }

        private bool HasNumber()
        {
            SkipSeparators();
            return _index < _data.Length
                   && (_data[_index] is '+' or '-' or '.' || char.IsDigit(_data[_index]));
        }

        private void RequireNumber()
        {
            if (!HasNumber()) throw InvalidPath();
        }

        private void SkipSeparators()
        {
            while (_index < _data.Length
                   && (char.IsWhiteSpace(_data[_index]) || _data[_index] == ','))
                _index++;
        }

        private void ResetControls()
        {
            _lastCubicControl = null;
            _lastQuadraticControl = null;
        }

        private static PointF Reflect(PointF point, PointF around) =>
            new(around.X * 2 - point.X, around.Y * 2 - point.Y);

        private static InvalidDataException InvalidPath() =>
            new("SVG path data is invalid or unsupported.");
    }
}
