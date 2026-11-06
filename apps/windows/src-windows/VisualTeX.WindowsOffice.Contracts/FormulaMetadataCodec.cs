using System;
using System.IO;
using System.IO.Compression;
using System.Text.Json;

namespace VisualTeX.WindowsOffice.Contracts;

public static class FormulaMetadataCodec
{
    public const string Prefix = "visualtex:v1:deflate:";

    public static string Encode(FormulaMetadata metadata)
    {
        metadata.Validate();
        var json = JsonSerializer.SerializeToUtf8Bytes(metadata, JsonOptions.Default);
        using var output = new MemoryStream();
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true))
            deflate.Write(json, 0, json.Length);
        return Prefix + Base64UrlEncode(output.ToArray());
    }

    public static string SerializeJson(FormulaMetadata metadata)
    {
        metadata.Validate();
        return JsonSerializer.Serialize(metadata, JsonOptions.Default);
    }

    public static FormulaMetadata? DeserializeJson(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var json = value!;
        try
        {
            var metadata = JsonSerializer.Deserialize<FormulaMetadata>(json, JsonOptions.Default);
            metadata?.Validate();
            return metadata;
        }
        catch
        {
            return null;
        }
    }

    public static FormulaMetadata? Decode(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var encodedValue = value!;
        var index = encodedValue.IndexOf(Prefix, StringComparison.Ordinal);
        if (index < 0) return null;
        var encoded = encodedValue.Substring(index + Prefix.Length).Trim();
        var end = encoded.IndexOfAny(new[] { '\r', '\n', ' ', '\t' });
        if (end >= 0) encoded = encoded.Substring(0, end);
        try
        {
            using var input = new MemoryStream(Base64UrlDecode(encoded));
            using var deflate = new DeflateStream(input, CompressionMode.Decompress);
            using var json = new MemoryStream();
            deflate.CopyTo(json);
            var bytes = json.ToArray();
            var metadata = JsonSerializer.Deserialize<FormulaMetadata>(bytes, JsonOptions.Default);
            if (metadata is not null) AdoptReferenceRenderSize(metadata, bytes);
            metadata?.Validate();
            return metadata;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>The 14 pt size at which VisualTeX for Mac measures a formula's picture.</summary>
    private const double MacReferenceFontSizePt = 14;

    /// <summary>
    /// VisualTeX for Mac stores a formula's render size at its 14 pt reference size (with that size in points as
    /// referenceWidthPt) and no renderFontSizePt; here render pixels are at renderFontSizePt (else fontSizePt). Its
    /// formulas get the font size their render pixels were measured at, so they keep their size when edited here.
    /// </summary>
    private static void AdoptReferenceRenderSize(FormulaMetadata metadata, byte[] json)
    {
        if (metadata.RenderFontSizePt.HasValue || metadata.RenderWidthPx is not > 0) return;
        try
        {
            using var document = JsonDocument.Parse(json);
            if (!document.RootElement.TryGetProperty("referenceWidthPt", out var reference)
                || reference.ValueKind != JsonValueKind.Number
                || !reference.TryGetDouble(out var referenceWidthPt)
                || referenceWidthPt <= 0)
                return;
            var renderFontSizePt = MacReferenceFontSizePt * metadata.RenderWidthPx.Value * 0.75 / referenceWidthPt;
            if (double.IsNaN(renderFontSizePt) || double.IsInfinity(renderFontSizePt)
                || renderFontSizePt < FormulaFontSize.MinimumPt || renderFontSizePt > FormulaFontSize.MaximumPt)
                return;
            metadata.RenderFontSizePt = Math.Round(renderFontSizePt, 3);
        }
        catch (JsonException) { }
    }

    private static string Base64UrlEncode(byte[] bytes) =>
        Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');

    private static byte[] Base64UrlDecode(string value)
    {
        var normalized = value.Replace('-', '+').Replace('_', '/');
        normalized = normalized.PadRight((normalized.Length + 3) / 4 * 4, '=');
        return Convert.FromBase64String(normalized);
    }
}
