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
            var metadata = JsonSerializer.Deserialize<FormulaMetadata>(deflate, JsonOptions.Default);
            metadata?.Validate();
            return metadata;
        }
        catch
        {
            return null;
        }
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
