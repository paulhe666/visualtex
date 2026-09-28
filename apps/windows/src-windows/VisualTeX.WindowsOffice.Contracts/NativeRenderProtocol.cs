using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace VisualTeX.WindowsOffice.Contracts;

// This boundary carries equation data, never paths, executable names or commands.
// A renderer can be replaced without changing Office insertion or the MTEF codec.
public sealed class NativeRenderRequest
{
    public int Version { get; set; } = NativeRenderProtocol.Version;
    public List<NativeRenderInput> Items { get; set; } = new();
}

public sealed class NativeRenderInput
{
    public string Id { get; set; } = string.Empty;
    public byte[] Mtef { get; set; } = Array.Empty<byte>();
}

public sealed class NativeRenderResponse
{
    public int Version { get; set; } = NativeRenderProtocol.Version;
    public string Error { get; set; } = string.Empty;
    public List<NativeRenderOutput> Items { get; set; } = new();
}

public sealed class NativeRenderOutput
{
    public string Id { get; set; } = string.Empty;
    public byte[] Wmf { get; set; } = Array.Empty<byte>();
    public float WidthPt { get; set; }
    public float HeightPt { get; set; }
    public int WordPosition { get; set; }
}

public static class NativeRenderProtocol
{
    public const int Version = 1;
    public const string ServiceName = "VisualTeX.NativeRender";
    public const string PipeName = "VisualTeX.NativeRender.v1";
    public const int MaxItems = 128;
    public const int MaxMtefBytes = 1024 * 1024;
    public const int MaxRequestBytes = 8 * 1024 * 1024;
    public const int MaxResponseBytes = 32 * 1024 * 1024;
    public const int MaxWmfBytes = 4 * 1024 * 1024;
    public const int RequestTimeoutMs = 180_000;

    public static void Validate(NativeRenderRequest request)
    {
        if (request == null || request.Version != Version || request.Items == null
            || request.Items.Count == 0 || request.Items.Count > MaxItems)
            throw new InvalidDataException("Unsupported render request.");
        var ids = new HashSet<string>(StringComparer.Ordinal);
        long total = 0;
        foreach (var item in request.Items)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.Id) || item.Id.Length > 256
                || !ids.Add(item.Id) || item.Mtef == null || item.Mtef.Length == 0
                || item.Mtef.Length > MaxMtefBytes)
                throw new InvalidDataException("Invalid render item.");
            total += item.Mtef.Length;
        }
        if (total > MaxRequestBytes / 2)
            throw new InvalidDataException("Render input exceeds the batch byte limit.");
    }

    public static void Validate(NativeRenderResponse response, NativeRenderRequest request)
    {
        if (response == null || response.Version != Version || response.Items == null
            || response.Items.Count > request.Items.Count)
            throw new InvalidDataException("Unsupported render response.");
        var ids = new HashSet<string>(request.Items.Select(item => item.Id), StringComparer.Ordinal);
        foreach (var item in response.Items)
        {
            if (item == null || !ids.Remove(item.Id) || item.Wmf == null
                || item.Wmf.Length <= 22 || item.Wmf.Length > MaxWmfBytes
                || !FinitePositive(item.WidthPt) || !FinitePositive(item.HeightPt)
                || item.WordPosition < -1_000_000 || item.WordPosition > 0)
                throw new InvalidDataException("Invalid render output.");
        }
    }

    private static bool FinitePositive(float value) =>
        value > 0 && value <= 1_000_000 && !float.IsInfinity(value) && !float.IsNaN(value);

    public static async Task WriteAsync<T>(Stream stream, T value, int maxBytes, CancellationToken token)
    {
        var bytes = JsonSerializer.SerializeToUtf8Bytes(value);
        if (bytes.Length == 0 || bytes.Length > maxBytes)
            throw new InvalidDataException("Render frame exceeds the byte limit.");
        // Explicit little endian; no BinaryFormatter or unbounded ReadToEnd.
        var size = bytes.Length;
        var header = new[] { (byte)size, (byte)(size >> 8), (byte)(size >> 16), (byte)(size >> 24) };
        await stream.WriteAsync(header, 0, header.Length, token).ConfigureAwait(false);
        await stream.WriteAsync(bytes, 0, bytes.Length, token).ConfigureAwait(false);
        await stream.FlushAsync(token).ConfigureAwait(false);
    }

    public static async Task<T> ReadAsync<T>(Stream stream, int maxBytes, CancellationToken token)
    {
        var header = new byte[4];
        await ReadExactlyAsync(stream, header, token).ConfigureAwait(false);
        int size = header[0] | header[1] << 8 | header[2] << 16 | header[3] << 24;
        if (size <= 0 || size > maxBytes) throw new InvalidDataException("Invalid render frame size.");
        var bytes = new byte[size];
        await ReadExactlyAsync(stream, bytes, token).ConfigureAwait(false);
        return JsonSerializer.Deserialize<T>(bytes, new JsonSerializerOptions { MaxDepth = 16 })
            ?? throw new InvalidDataException("Empty render frame.");
    }

    private static async Task ReadExactlyAsync(Stream stream, byte[] bytes, CancellationToken token)
    {
        for (int offset = 0; offset < bytes.Length;)
        {
            int count = await stream.ReadAsync(bytes, offset, bytes.Length - offset, token).ConfigureAwait(false);
            if (count == 0) throw new EndOfStreamException("Incomplete render frame.");
            offset += count;
        }
    }
}
