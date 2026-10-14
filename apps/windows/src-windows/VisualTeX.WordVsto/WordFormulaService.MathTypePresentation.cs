using System.Security.Cryptography;

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

    internal static float CalculateMathTypePreservedOriginalExtent(
        float sourceOriginalExtent,
        float sourceDisplayedExtent,
        float targetDisplayedExtent)
    {
        // Word stores the OLE object's intrinsic extent (dxaOrig/dyaOrig)
        // independently from the VML extent it displays. Existing MathType
        // documents legitimately use a non-1:1 relationship between them.
        // Preserve that relationship across an edit instead of replacing it
        // with frontend-export or raw MathPage dimensions.
        if (!(sourceOriginalExtent > 0)
            || !(sourceDisplayedExtent > 0)
            || !(targetDisplayedExtent > 0))
            return targetDisplayedExtent;
        var result = targetDisplayedExtent * sourceOriginalExtent / sourceDisplayedExtent;
        return result > 0 && !float.IsNaN(result) && !float.IsInfinity(result)
            ? result
            : targetDisplayedExtent;
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

    // What a placeable WMF draws: its frame and every record except MathType's
    // comments (escapes) and the window records Word's copy adds.
    internal static string MathTypePictureDrawing(byte[] picture)
    {
        const uint placeableKey = 0x9ac6cdd7;
        using var hash = SHA256.Create();
        if (picture is null || picture.Length < 22 + 18 || BitConverter.ToUInt32(picture, 0) != placeableKey)
            return "not-a-placeable-wmf:"
                + Convert.ToBase64String(hash.ComputeHash(picture ?? Array.Empty<byte>()));
        hash.TransformBlock(picture, 6, 10, null, 0); // bounding box and units per inch
        var headerWords = BitConverter.ToUInt16(picture, 22 + 2);
        var position = 22 + headerWords * 2;
        while (position + 6 <= picture.Length)
        {
            var words = BitConverter.ToUInt32(picture, position);
            var function = BitConverter.ToUInt16(picture, position + 4);
            if (words < 3 || position + (long)words * 2 > picture.Length) break;
            if (function != 0x0626 && function != 0x020B && function != 0x020C)
            {
                var record = new byte[(int)words * 2 - 4];
                Buffer.BlockCopy(picture, position + 4, record, 0, record.Length);
                ClearMathTypeRecordPadding(function, record);
                hash.TransformBlock(record, 0, record.Length, null, 0);
            }
            if (function == 0) break;
            position += (int)words * 2;
        }
        hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        return Convert.ToBase64String(hash.Hash);
    }

    // MathType leaves uninitialized bytes where a record is padded, so two
    // renders of one equation differ there: the pad after an odd-length string,
    // the rest of a font's face name, a pen's unused width.y, a solid brush's
    // hatch and the reserved word of a mode record. None of them draws.
    // The record starts at its function word; parameters follow at offset 2.
    private static void ClearMathTypeRecordPadding(ushort function, byte[] record)
    {
        const int parameters = 2;
        switch (function)
        {
            case 0x0521: // TEXTOUT: length, string padded to a word, y, x
            {
                if (record.Length < parameters + 2) break;
                var length = BitConverter.ToInt16(record, parameters);
                var pad = parameters + 2 + length;
                if (length > 0 && (length & 1) == 1 && pad < record.Length) record[pad] = 0;
                break;
            }
            case 0x0A32: // EXTTEXTOUT: y, x, length, options, [rectangle], string padded to a word, [dx]
            {
                if (record.Length < parameters + 8) break;
                var length = BitConverter.ToInt16(record, parameters + 4);
                var options = BitConverter.ToUInt16(record, parameters + 6);
                var pad = parameters + 8 + ((options & 0x0006) != 0 ? 8 : 0) + length;
                if (length > 0 && (length & 1) == 1 && pad < record.Length) record[pad] = 0;
                break;
            }
            case 0x02FB: // CREATEFONTINDIRECT: 18-byte LOGFONT, then a NUL-terminated face name
            {
                if (record.Length <= parameters + 18) break;
                var end = Array.IndexOf(record, (byte)0, parameters + 18);
                if (end >= 0) Array.Clear(record, end, record.Length - end);
                break;
            }
            case 0x02FA: // CREATEPENINDIRECT: style, width (x, unused y), color
                if (record.Length >= parameters + 6) record[parameters + 4] = record[parameters + 5] = 0;
                break;
            case 0x02FC: // CREATEBRUSHINDIRECT: style, color, hatch (used only by BS_HATCHED)
                if (record.Length >= parameters + 8 && BitConverter.ToUInt16(record, parameters) != 2)
                    record[parameters + 6] = record[parameters + 7] = 0;
                break;
            case 0x0102: // SETBKMODE
            case 0x0104: // SETROP2
            case 0x0106: // SETPOLYFILLMODE
            case 0x0107: // SETSTRETCHBLTMODE
            case 0x012E: // SETTEXTALIGN: one mode word, then an optional reserved word
                if (record.Length > parameters + 2)
                    Array.Clear(record, parameters + 2, record.Length - parameters - 2);
                break;
        }
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
