namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    internal readonly struct MathTypeWordPresentationGeometry
    {
        internal MathTypeWordPresentationGeometry(
            string shapeStyle,
            int originalWidthTwips,
            int originalHeightTwips)
        {
            ShapeStyle = shapeStyle;
            OriginalWidthTwips = originalWidthTwips;
            OriginalHeightTwips = originalHeightTwips;
        }

        internal string ShapeStyle { get; }
        internal int OriginalWidthTwips { get; }
        internal int OriginalHeightTwips { get; }
    }

    internal static float CalculateMathTypeNativePresentationScale(
        float sourceWordExtent,
        float? sourceNativeExtent)
    {
        if (!(sourceWordExtent > 0) || sourceNativeExtent is not > 0) return 1f;
        var scale = sourceWordExtent / sourceNativeExtent.Value;
        if (!(scale > 0) || float.IsNaN(scale) || float.IsInfinity(scale)) return 1f;
        return Math.Max(0.25f, Math.Min(4f, scale));
    }

    internal static MathTypeWordPresentationGeometry CalculateMathTypeWordPresentationGeometry(
        string shapeStyle,
        float widthPt,
        float heightPt,
        float originalWidthPt,
        float originalHeightPt)
    {
        var style = ReplaceMathTypeStylePoints(shapeStyle, "width", widthPt);
        style = ReplaceMathTypeStylePoints(style, "height", heightPt);
        return new MathTypeWordPresentationGeometry(
            style,
            MathTypeOriginalTwips(originalWidthPt),
            MathTypeOriginalTwips(originalHeightPt));
    }

    internal static byte[] EnsureMathTypePlaceableWmfWindowRecords(byte[] previewWmf)
    {
        const uint aldusPlaceableKey = 0x9ac6cdd7;
        const ushort metaSetWindowOrg = 0x020b;
        const ushort metaSetWindowExt = 0x020c;
        const int firstRecordOffset = 40;
        if (previewWmf is null
            || previewWmf.Length < firstRecordOffset + 6
            || BitConverter.ToUInt32(previewWmf, 0) != aldusPlaceableKey)
            throw new InvalidDataException("MathType preview is not a valid placeable WMF.");

        var declaredWords = BitConverter.ToUInt32(previewWmf, 28);
        if (declaredWords > int.MaxValue / 2
            || checked((int)declaredWords * 2) != previewWmf.Length - 22)
            throw new InvalidDataException("MathType WMF header length does not match its payload.");

        var hasWindowOrigin = false;
        var hasWindowExtent = false;
        var insertionOffset = firstRecordOffset;
        var leadingSetup = true;
        for (var offset = firstRecordOffset; offset + 6 <= previewWmf.Length;)
        {
            var words = BitConverter.ToUInt32(previewWmf, offset);
            if (words < 3 || words > int.MaxValue / 2)
                throw new InvalidDataException("MathType WMF contains an invalid record length.");
            var size = checked((int)words * 2);
            if (offset + size > previewWmf.Length)
                throw new InvalidDataException("MathType WMF contains a truncated record.");
            var function = BitConverter.ToUInt16(previewWmf, offset + 4);
            hasWindowOrigin |= function == metaSetWindowOrg;
            hasWindowExtent |= function == metaSetWindowExt;
            if (leadingSetup
                && function is 0x0102 or 0x0201 or 0x012e)
                insertionOffset = offset + size;
            else
                leadingSetup = false;
            offset += size;
            if (function == 0) break;
        }

        if (hasWindowOrigin && hasWindowExtent) return previewWmf;
        if (hasWindowOrigin || hasWindowExtent)
            throw new InvalidDataException(
                "MathType WMF contains only one of SETWINDOWORG/SETWINDOWEXT.");

        var left = BitConverter.ToInt16(previewWmf, 6);
        var top = BitConverter.ToInt16(previewWmf, 8);
        var right = BitConverter.ToInt16(previewWmf, 10);
        var bottom = BitConverter.ToInt16(previewWmf, 12);
        var width = checked((short)(right - left));
        var height = checked((short)(bottom - top));
        if (width <= 0 || height <= 0)
            throw new InvalidDataException("MathType placeable WMF has invalid logical bounds.");

        var result = new byte[checked(previewWmf.Length + 20)];
        Buffer.BlockCopy(previewWmf, 0, result, 0, insertionOffset);
        WriteMathTypeWmfRecord(result, insertionOffset, metaSetWindowOrg, top, left);
        WriteMathTypeWmfRecord(result, insertionOffset + 10, metaSetWindowExt, height, width);
        Buffer.BlockCopy(
            previewWmf,
            insertionOffset,
            result,
            insertionOffset + 20,
            previewWmf.Length - insertionOffset);
        WriteMathTypeUInt32(result, 28, checked(declaredWords + 10));
        return result;
    }

    private static void WriteMathTypeWmfRecord(
        byte[] target,
        int offset,
        ushort function,
        short firstParameter,
        short secondParameter)
    {
        WriteMathTypeUInt32(target, offset, 5);
        WriteMathTypeUInt16(target, offset + 4, function);
        WriteMathTypeUInt16(target, offset + 6, unchecked((ushort)firstParameter));
        WriteMathTypeUInt16(target, offset + 8, unchecked((ushort)secondParameter));
    }

    private static void WriteMathTypeUInt16(byte[] target, int offset, ushort value)
    {
        target[offset] = (byte)value;
        target[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteMathTypeUInt32(byte[] target, int offset, uint value)
    {
        target[offset] = (byte)value;
        target[offset + 1] = (byte)(value >> 8);
        target[offset + 2] = (byte)(value >> 16);
        target[offset + 3] = (byte)(value >> 24);
    }

    private static int MathTypeOriginalTwips(float valuePt) =>
        checked(Math.Max(
            1,
            (int)Math.Round(valuePt * 20f, MidpointRounding.AwayFromZero)));

    private static string ReplaceMathTypeStylePoints(
        string style,
        string property,
        float valuePt)
    {
        var replacement = valuePt.ToString(
            "0.###",
            System.Globalization.CultureInfo.InvariantCulture) + "pt";
        var segments = style.Split(';').ToList();
        var replaced = false;
        for (var index = 0; index < segments.Count; index++)
        {
            var parts = segments[index].Split(new[] { ':' }, 2);
            if (parts.Length != 2
                || !string.Equals(
                    parts[0].Trim(),
                    property,
                    StringComparison.OrdinalIgnoreCase))
                continue;
            segments[index] = parts[0] + ":" + replacement;
            replaced = true;
            break;
        }
        if (!replaced) segments.Add(property + ":" + replacement);
        return string.Join(";", segments);
    }
}
