using System.IO.Compression;
using System.Text;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOleBridge;

internal static class MetadataCodec
{
    public const string Prefix = "visualtex:v1:deflate:";

    public static string Encode(FormulaMetadata metadata)
    {
        metadata.Validate();
        var json = JsonSerializer.SerializeToUtf8Bytes(metadata, JsonOptions.Default);
        using var output = new MemoryStream();
        using (var deflate = new DeflateStream(output, CompressionLevel.SmallestSize, leaveOpen: true))
            deflate.Write(json, 0, json.Length);
        return Prefix + Base64UrlEncode(output.ToArray());
    }

    public static FormulaMetadata? Decode(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var index = value.IndexOf(Prefix, StringComparison.Ordinal);
        if (index < 0) return null;
        var encoded = value[(index + Prefix.Length)..].Trim();
        var end = encoded.IndexOfAny(new[] { '\r', '\n', ' ', '\t' });
        if (end >= 0) encoded = encoded[..end];
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

internal static class JsonOptions
{
    public static readonly JsonSerializerOptions Default = new()
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };
}

internal sealed class TempPathGuard
{
    private readonly string _root;

    public TempPathGuard(string root)
    {
        Directory.CreateDirectory(root);
        _root = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
    }

    public string ValidatePng(string path) => Validate(path, ".png", stream =>
    {
        Span<byte> signature = stackalloc byte[8];
        if (stream.Read(signature) != 8
            || !signature.SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }))
            throw new InvalidDataException("Formula image has an invalid PNG signature.");
    });

    public string ValidateSvg(string path) => Validate(path, ".svg", stream =>
    {
        if (stream.Length <= 0 || stream.Length > 16 * 1024 * 1024)
            throw new InvalidDataException("Formula SVG has an invalid size.");
        using var reader = new StreamReader(
            stream,
            new UTF8Encoding(false, true),
            detectEncodingFromByteOrderMarks: true,
            bufferSize: 1024,
            leaveOpen: true);
        var svg = reader.ReadToEnd().Trim();
        if (!svg.StartsWith("<svg", StringComparison.OrdinalIgnoreCase)
            || svg.IndexOf("</svg>", StringComparison.OrdinalIgnoreCase) < 0)
            throw new InvalidDataException("Formula image has an invalid SVG document.");
        if (svg.IndexOf("<foreignObject", StringComparison.OrdinalIgnoreCase) >= 0
            || svg.IndexOf("<image", StringComparison.OrdinalIgnoreCase) >= 0
            || svg.IndexOf("<script", StringComparison.OrdinalIgnoreCase) >= 0
            || svg.IndexOf("<iframe", StringComparison.OrdinalIgnoreCase) >= 0
            || svg.IndexOf("javascript:", StringComparison.OrdinalIgnoreCase) >= 0)
            throw new InvalidDataException("Formula SVG contains forbidden content.");
    });

    private string Validate(string path, string extension, Action<FileStream> validateContent)
    {
        var full = Path.GetFullPath(path);
        if (!full.StartsWith(_root, StringComparison.OrdinalIgnoreCase))
            throw new UnauthorizedAccessException("Formula image is outside the VisualTeX Office temp directory.");
        if (!string.Equals(Path.GetExtension(full), extension, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException($"Windows Office formula image must be {extension}.");
        if (!File.Exists(full)) throw new FileNotFoundException("Formula image does not exist.", full);
        using var stream = File.OpenRead(full);
        validateContent(stream);
        return full;
    }
}
