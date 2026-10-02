using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeGroupAccentNativeProbeTests
{
    // A renderer-only experiment using documented MTEF structural records.
    // It does not alter the product writer or attach to Word/interactive MathType.
    [Fact]
    public async Task ProbeNativeCenteredOverscriptWhenEnabled()
    {
        if (Environment.GetEnvironmentVariable("VISUALTEX_RUN_GROUP_ACCENT_PROBE") != "1") return;
        var root = Environment.GetEnvironmentVariable("VISUALTEX_CHARACTER_RENDER_ARTIFACT_DIR")
            ?? throw new InvalidOperationException("An owned evidence directory is required.");
        Directory.CreateDirectory(root);
        var seed = MathTypeMtefCodec.CreateEquationNativeAtFontSize("<math><mi>a</mi><mi>b</mi></math>", false, 12);
        var mark = MathTypeMtefCodec.CreateEquationNativeAtFontSize("<math><mo>¨</mo></math>", false, 12);
        var basisLine = seed.Mtef.Skip(seed.StructureOffset).Take(seed.Mtef.Length - seed.StructureOffset - 1).ToArray();
        var markLine = mark.Mtef.Skip(mark.StructureOffset).Take(mark.Mtef.Length - mark.StructureOffset - 1).ToArray();
        var candidates = new Dictionary<string, byte[]> {
            ["plain-control"] = MathTypeMtefCodec.CreateEquationNativeAtFontSize("<math><mi>x</mi><mo>+</mo><mn>1</mn></math>", false, 12).Mtef
        };
        byte[] Native(IEnumerable<byte> objects) => seed.Mtef.Take(seed.StructureOffset)
            .Concat(new byte[] { 1, 0 }).Concat(objects).Concat(new byte[] { 0, 0 }).ToArray();
        // LimBoxClass: main, lower, upper, END; tvBO_UPPER|tvBO_SUM.
        candidates.Add("centered-upper-limit", Native(new byte[] { 3, 0, 23, 0x42, 0 }
            .Concat(basisLine).Concat(new byte[] { 1, 1 }).Concat(markLine).Concat(new byte[] { 0, 10 })));
        candidates.Add("normal-hat-control", MathTypeMtefCodec.CreateEquationNativeAtFontSize(
            "<math><mover accent='true'><mrow><mi>a</mi><mi>b</mi></mrow><mo>^</mo></mover></math>", false, 12).Mtef);
        // HatBoxClass may have a serialized mark glyph after the main slot.
        candidates.Add("hat-with-dot-glyph", Native(new byte[] { 3, 0, 33, 0, 0 }
            .Concat(basisLine).Concat(markLine.Skip(2).Take(markLine.Length - 3)).Concat(new byte[] { 0 })));
        candidates.Add("script-negative-control", MathTypeMtefCodec.CreateEquationNativeAtFontSize(
            "<math><msup><mrow><mi>a</mi><mi>b</mi></mrow><mo>¨</mo></msup></math>", false, 12).Mtef);
        // Prior batches returned worker-disconnected with no per-item result.
        // That does not identify the failing input. Run precisely one explicitly
        // selected request, with a plain arithmetic control as the safe default.
        var selected = Environment.GetEnvironmentVariable("VISUALTEX_GROUP_PROBE_CASE") ?? "plain-control";
        var request = new NativeRenderRequest();
        foreach (var item in candidates.Where(item => item.Key == selected))
        {
            File.WriteAllBytes(Path.Combine(root, item.Key + ".mtef"), item.Value);
            request.Items.Add(new NativeRenderInput { Id = item.Key, Mtef = item.Value });
        }
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(70));
        var response = await NativeRenderClient.RenderAsync(request, timeout.Token);
        File.WriteAllText(Path.Combine(root, "response.json"), JsonSerializer.Serialize(response, new JsonSerializerOptions { WriteIndented = true }));
        foreach (var item in response.Items)
        {
            File.WriteAllBytes(Path.Combine(root, item.Id + ".wmf"), item.Wmf);
            File.WriteAllText(Path.Combine(root, item.Id + ".records.json"), JsonSerializer.Serialize(DescribeWmf(item.Wmf), new JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine($"{item.Id}: {item.WidthPt} x {item.HeightPt}, wmf={item.Wmf.Length}");
        }
        Assert.Equal(string.Empty, response.Error);
        Assert.Equal(request.Items.Count, response.Items.Count);
    }

    private static object[] DescribeWmf(byte[] bytes)
    {
        var records = new List<object>();
        var start = bytes.Length >= 22 && BitConverter.ToUInt32(bytes, 0) == 0x9AC6CDD7 ? 22 : 0;
        var cursor = start + BitConverter.ToUInt16(bytes, start + 2) * 2;
        while (cursor + 6 <= bytes.Length)
        {
            var size = checked((int)BitConverter.ToUInt32(bytes, cursor) * 2);
            if (size < 6 || cursor + size > bytes.Length) throw new InvalidDataException("Bad WMF record.");
            var kind = BitConverter.ToUInt16(bytes, cursor + 4);
            if (kind == 0x02FB)
            {
                var end = cursor + 24;
                while (end < cursor + size && bytes[end] != 0) end++;
                records.Add(new { Kind = "Font", Name = System.Text.Encoding.ASCII.GetString(bytes, cursor + 24, end - cursor - 24), Weight = BitConverter.ToInt16(bytes, cursor + 14), Italic = bytes[cursor + 16] != 0 });
            }
            else if (kind == 0x0A32)
            {
                var count = BitConverter.ToUInt16(bytes, cursor + 10);
                var flags = BitConverter.ToUInt16(bytes, cursor + 12);
                var textStart = cursor + 14 + ((flags & 6) != 0 ? 8 : 0);
                records.Add(new { Kind = "Text", X = BitConverter.ToInt16(bytes, cursor + 8), Y = BitConverter.ToInt16(bytes, cursor + 6), Flags = flags, Codes = bytes.Skip(textStart).Take(count).Select(x => (int)x).ToArray() });
            }
            else if (kind == 0x0521)
            {
                var count = BitConverter.ToUInt16(bytes, cursor + 6);
                records.Add(new { Kind = "TextOut", Codes = bytes.Skip(cursor + 8).Take(count).Select(x => (int)x).ToArray() });
            }
            cursor += size;
            if (kind == 0) break;
        }
        return records.ToArray();
    }
}
