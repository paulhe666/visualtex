using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using Xunit;

namespace VisualTeX.NativeRender.Tests;

public sealed class NativeRenderLiveFactAttribute : FactAttribute
{
    public NativeRenderLiveFactAttribute()
    {
        if (Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_RENDER_LIVE_TEST") != "1")
            Skip = "Requires explicit opt-in and the temporary native-render service.";
    }
}

public sealed partial class NativeRenderLiveTests
{
    [NativeRenderLiveFact]
    public async Task IndependentClientsRenderFiveNativeSamples()
    {
        string repo = Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_RENDER_TEST_REPO")
            ?? throw new InvalidOperationException("Missing test repository path.");
        var request = new NativeRenderRequest();
        foreach (int number in new[] { 1, 15, 27, 219 })
        {
            string file = Path.Combine(repo, "docs", "remediation-3d207d7", "evidence",
                "mtef-redraw-20260923", "native-core", "formula-" + number + ".mtef");
            request.Items.Add(new NativeRenderInput { Id = "source-" + number, Mtef = File.ReadAllBytes(file) });
        }
        request.Items.Add(new NativeRenderInput
        {
            Id = "nuclear",
            Mtef = File.ReadAllBytes(Path.Combine(repo, "docs", "remediation-3d207d7", "evidence",
                "nuclear-prescript-native-r3-20260926", "formula-1.mtef")),
        });
        Assert.True(NativeRenderClient.IsAvailable);
        // Separate pipe connections simulate Word and PowerPoint requesting at
        // once; the shared service must serialize its native engine, not share it
        // with MathType running in the interactive user session.
        var responses = await Task.WhenAll(NativeRenderClient.RenderAsync(request, default),
            NativeRenderClient.RenderAsync(request, default));
        string output = Path.Combine(repo, "artifacts", "mathtype-install-20260927", "service-candidate", "client-results");
        string? stage = Environment.GetEnvironmentVariable("VISUALTEX_NATIVE_RENDER_TEST_STAGE");
        if (!string.IsNullOrEmpty(stage))
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(stage, "^[a-z0-9-]{1,64}$"))
                throw new InvalidOperationException("Invalid evidence stage.");
            output = Path.Combine(output, stage);
        }
        Directory.CreateDirectory(output);
        for (int index = 0; index < responses.Length; index++)
        {
            var result = responses[index];
            File.WriteAllText(Path.Combine(output, "response-" + index + ".json"), JsonSerializer.Serialize(result));
            Assert.Equal(string.Empty, result.Error);
            Assert.Equal(request.Items.Count, result.Items.Count);
            foreach (var item in result.Items)
                File.WriteAllBytes(Path.Combine(output, index + "-" + item.Id + ".wmf"), item.Wmf);
        }
    }
}
