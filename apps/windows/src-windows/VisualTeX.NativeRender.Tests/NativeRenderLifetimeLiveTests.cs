using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using Xunit;

namespace VisualTeX.NativeRender.Tests;

public sealed partial class NativeRenderLiveTests
{
    [NativeRenderLiveFact]
    public async Task RepeatedRequestsKeepExactNativeLayoutAndSerializeIndependentClients()
    {
        var samples = LifetimeSamples();
        var baseline = await NativeRenderClient.RenderAsync(samples, default);
        Assert.Equal(string.Empty, baseline.Error);
        Assert.Equal(samples.Items.Count, baseline.Items.Count);
        var expected = baseline.Items.ToDictionary(item => item.Id);
        foreach (var item in baseline.Items) SaveWmfEvidence("baseline-" + item.Id + ".wmf", item.Wmf);
        var measurements = new List<object>();
        for (int round = 0; round < 4; round++)
        {
            foreach (var input in samples.Items.AsEnumerable().Reverse())
            {
                var request = new NativeRenderRequest { Items = new() { input } };
                var watch = Stopwatch.StartNew();
                var response = await NativeRenderClient.RenderAsync(request, default);
                watch.Stop();
                Assert.Equal(string.Empty, response.Error);
                var actual = Assert.Single(response.Items);
                SaveWmfEvidence("round-" + round + "-" + input.Id + ".wmf", actual.Wmf);
                AssertSameNativeOutput(expected[input.Id], actual);
                measurements.Add(new { Round = round, input.Id, Milliseconds = watch.Elapsed.TotalMilliseconds,
                    actual.WidthPt, actual.HeightPt, actual.WordPosition,
                    WmfSha256 = Convert.ToHexString(SHA256.HashData(actual.Wmf)) });
            }
        }
        // Distinct concurrent clients must never cross-associate their output.
        var reversed = new NativeRenderRequest { Items = samples.Items.AsEnumerable().Reverse().ToList() };
        var responses = await Task.WhenAll(NativeRenderClient.RenderAsync(samples, default),
            NativeRenderClient.RenderAsync(reversed, default));
        foreach (var response in responses)
        {
            Assert.Equal(string.Empty, response.Error);
            Assert.Equal(expected.Count, response.Items.Count);
            foreach (var actual in response.Items) AssertSameNativeOutput(expected[actual.Id], actual);
        }
        SaveLifetimeEvidence("repeated-native-layout.json", new { Passed = true, Measurements = measurements });
    }

    [NativeRenderLiveFact]
    public async Task SeparateClientAfterIdleRetainsTheSamePrivateApiLease()
    {
        var request = LifetimeSamples();
        request.Items = new() { request.Items[0] };
        var first = await NativeRenderClient.RenderAsync(request, default);
        var expected = Assert.Single(first.Items);
        var workers = PrivateProcessIds("VisualTeX.NativeRenderWorker");
        var engines = PrivateProcessIds("MathType");
        Assert.Single(workers);
        Assert.Single(engines);
        // The earlier job-only candidate passed immediate calls but released
        // MTTermAPI on broker exit; its engine died before the next UI Apply.
        await Task.Delay(TimeSpan.FromSeconds(25));
        Assert.Equal(workers, PrivateProcessIds("VisualTeX.NativeRenderWorker"));
        Assert.Equal(engines, PrivateProcessIds("MathType"));
        var watch = Stopwatch.StartNew();
        var next = await NativeRenderClient.RenderAsync(request, default);
        watch.Stop();
        AssertSameNativeOutput(expected, Assert.Single(next.Items));
        Assert.True(watch.ElapsedMilliseconds < 1000,
            "Idle request paid cold-start cost: " + watch.Elapsed.TotalMilliseconds + " ms");
        SaveLifetimeEvidence("idle-api-lease.json", new { Passed = true, IdleSeconds = 25,
            WorkerIds = workers, EngineIds = engines, Milliseconds = watch.Elapsed.TotalMilliseconds });
    }

    private static int[] PrivateProcessIds(string name)
    {
        var ids = new List<int>();
        foreach (var process in Process.GetProcessesByName(name))
            using (process)
                if (process.SessionId == 0) ids.Add(process.Id);
        return ids.OrderBy(id => id).ToArray();
    }

    private static NativeRenderRequest LifetimeSamples()
    {
        string repo = Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_RENDER_TEST_REPO")
            ?? throw new InvalidOperationException("Missing test repository path.");
        var request = new NativeRenderRequest();
        foreach (int number in new[] { 1, 15, 27, 219 })
            request.Items.Add(new NativeRenderInput { Id = "source-" + number,
                Mtef = File.ReadAllBytes(Path.Combine(repo, "docs", "remediation-3d207d7", "evidence",
                    "mtef-redraw-20260923", "native-core", "formula-" + number + ".mtef")) });
        request.Items.Add(new NativeRenderInput { Id = "nuclear",
            Mtef = File.ReadAllBytes(Path.Combine(repo, "docs", "remediation-3d207d7", "evidence",
                "nuclear-prescript-native-r3-20260926", "formula-1.mtef")) });
        return request;
    }

    private static void AssertSameNativeOutput(NativeRenderOutput expected, NativeRenderOutput actual)
    {
        Assert.Equal(expected.Id, actual.Id);
        Assert.Equal(expected.WidthPt, actual.WidthPt);
        Assert.Equal(expected.HeightPt, actual.HeightPt);
        Assert.Equal(expected.WordPosition, actual.WordPosition);
        // WMF contains undefined bytes after a font-face NUL and alignment bytes
        // after odd-length text. Raw MathPage output varies there even when every
        // drawing operation is identical. Compare ALL defined bytes, then pixels.
        if (!DefinedWmfBytes(expected.Wmf).SequenceEqual(DefinedWmfBytes(actual.Wmf)))
        {
            SaveWmfEvidence("baseline-failed.wmf", expected.Wmf);
            SaveWmfEvidence("round-0-failed.wmf", actual.Wmf);
        }
        Assert.Equal(RasterHash(expected.Wmf), RasterHash(actual.Wmf));
        Assert.Equal(DefinedWmfBytes(expected.Wmf), DefinedWmfBytes(actual.Wmf));
    }

    private static byte[] DefinedWmfBytes(byte[] source)
    {
        var bytes = (byte[])source.Clone();
        // MS-WMF 2.2.1.2 Font Object: bytes after the Facename terminating NUL
        // are ignored. 2.3.3.5 META_EXTTEXTOUT: odd text has a WORD-alignment byte.
        // https://learn.microsoft.com/openspecs/windows_protocols/ms-wmf/dabb1ed6-e5e8-4243-80ed-e63443e5484f
        // https://learn.microsoft.com/openspecs/windows_protocols/ms-wmf/7d07c44a-a828-4b82-9af0-e0a81cced5a8
        Assert.True(bytes.Length >= 40);
        for (int offset = 40; offset < bytes.Length;)
        {
            Assert.True(offset + 6 <= bytes.Length);
            int size = checked((int)BitConverter.ToUInt32(bytes, offset) * 2);
            Assert.True(size >= 6 && offset + size <= bytes.Length);
            int end = offset + size;
            switch (BitConverter.ToUInt16(bytes, offset + 4))
            {
                case 0x02fb:
                    Assert.True(size >= 56);
                    int faceStart = offset + 24, faceEnd = faceStart + 32;
                    int nul = Array.IndexOf(bytes, (byte)0, faceStart, 32);
                    if (nul >= 0) Array.Clear(bytes, nul + 1, faceEnd - nul - 1);
                    break;
                case 0x0a32:
                    Assert.True(size >= 14);
                    int count = BitConverter.ToUInt16(bytes, offset + 10);
                    int options = BitConverter.ToUInt16(bytes, offset + 12);
                    int text = offset + 14 + ((options & 6) != 0 ? 8 : 0);
                    Assert.True(text + count + (count & 1) <= end);
                    if ((count & 1) != 0) bytes[text + count] = 0;
                    break;
                case 0x0626:
                    // META_ESCAPE ByteCount excludes the trailing WORD padding.
                    // Keep the entire MathType comment/MTEF payload unchanged.
                    Assert.True(size >= 10);
                    int dataBytes = BitConverter.ToUInt16(bytes, offset + 8);
                    Assert.Equal(size, 10 + dataBytes + (dataBytes & 1));
                    if ((dataBytes & 1) != 0) bytes[offset + 10 + dataBytes] = 0;
                    break;
                case 0x0521:
                    Assert.True(size >= 8);
                    int length = BitConverter.ToUInt16(bytes, offset + 6);
                    Assert.True(offset + 8 + length + (length & 1) <= end);
                    if ((length & 1) != 0) bytes[offset + 8 + length] = 0;
                    break;
            }
            offset = end;
        }
        return bytes;
    }

    private static string RasterHash(byte[] wmf)
    {
        using var stream = new MemoryStream(wmf, writable: false);
        using var metafile = new Metafile(stream);
        using var bitmap = new Bitmap(1600, 800, PixelFormat.Format32bppArgb);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            graphics.DrawImage(metafile, new Rectangle(20, 20, 1520, 720));
        }
        var bits = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height),
            ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        try
        {
            var pixels = new byte[Math.Abs(bits.Stride) * bitmap.Height];
            Marshal.Copy(bits.Scan0, pixels, 0, pixels.Length);
            Assert.Contains(pixels, value => value != 255); // not two blank renders
            return Convert.ToHexString(SHA256.HashData(pixels));
        }
        finally { bitmap.UnlockBits(bits); }
    }

    private static void SaveWmfEvidence(string file, byte[] bytes)
    {
        string? directory = Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_RENDER_LIFETIME_EVIDENCE");
        if (string.IsNullOrEmpty(directory)) return;
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, file), bytes);
    }

    private static void SaveLifetimeEvidence(string file, object data)
    {
        string? directory = Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_RENDER_LIFETIME_EVIDENCE");
        if (string.IsNullOrEmpty(directory)) return;
        Directory.CreateDirectory(directory);
        File.WriteAllText(Path.Combine(directory, file), JsonSerializer.Serialize(data,
            new JsonSerializerOptions { WriteIndented = true }));
    }
}
