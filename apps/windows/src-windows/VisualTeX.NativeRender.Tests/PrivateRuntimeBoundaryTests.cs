using System.Text.Json;
using VisualTeX.NativeRenderWorker;
using Xunit;

namespace VisualTeX.NativeRender.Tests;

public sealed class PrivateRuntimeBoundaryTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MissingOrInvalidPrivateLibraryNeverFallsBackToInstalledMathType(bool invalidLibrary)
    {
        string root = Path.Combine(Path.GetTempPath(), "VisualTeX-PrivateBoundary-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            string runtime = Path.Combine(root, "runtime");
            if (invalidLibrary)
            {
                string libraryDirectory = Path.Combine(runtime, "MathPage", "64");
                Directory.CreateDirectory(libraryDirectory);
                File.WriteAllBytes(Path.Combine(libraryDirectory, "MathPage.wll"), new byte[32]);
            }
            string response = Path.Combine(root, "response.json");
            string manifest = Path.Combine(root, "request.json");
            File.WriteAllText(manifest, JsonSerializer.Serialize(new
            {
                ResultPath = response,
                // Even a real installed-server hint must be ignored in strict mode.
                MathTypeServerPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "MathType", "MathType.exe"),
                Items = new[] { new { Id = "one", MtefPath = Path.Combine(root, "not-read.mtef"), WmfPath = Path.Combine(root, "not-created.wmf") } },
            }));
            Assert.Equal(3, MathTypeNativePreviewCommand.Run(manifest, runtime));
            using var result = JsonDocument.Parse(File.ReadAllText(response));
            Assert.NotEqual(string.Empty, result.RootElement.GetProperty("Error").GetString());
            Assert.Equal(0, result.RootElement.GetProperty("Items").GetArrayLength());
        }
        finally { Directory.Delete(root, recursive: true); }
    }
}
