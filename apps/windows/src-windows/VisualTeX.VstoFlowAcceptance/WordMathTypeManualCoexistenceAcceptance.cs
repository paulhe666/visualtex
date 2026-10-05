using System.Diagnostics;
using System.Text.Json;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // Physical Ribbon clicks previously differed from automated macro calls.
    // Keep one owned blank Word document open for that specific manual check.
    private static void RunWordMathTypeManualCoexistenceAcceptance(string artifactRoot)
    {
        RequireRegistryFreeAcceptance("word-mathtype-manual-coexistence");
        if (Process.GetProcessesByName("WINWORD").Any())
            throw new InvalidOperationException("Close existing Word sessions before this owned-document test.");
        var statePath = Path.Combine(artifactRoot, "manual-state.json");
        var completionSignal = Path.Combine(artifactRoot, "insertion-returned.signal");
        if (File.Exists(completionSignal))
            throw new InvalidOperationException("Use a fresh artifact directory without an old completion signal.");
        try
        {
        using var filter = OfficeComMessageFilter.Register();
        RunMathTypeNativePreviewPrivateBoundaryAcceptance(artifactRoot);
        using var host = new WordPerformanceHost(documentPath: null);
        host.Application.Caption = "VisualTeX MathType 共存验收";
        host.Application.Visible = true;
        host.Document.Activate();
        using var word = Process.GetProcessesByName("WINWORD").Single();
        File.WriteAllText(statePath, JsonSerializer.Serialize(new
        {
            Status = "WaitingForPhysicalRibbonClick", WordProcessId = word.Id,
            StartedUtc = word.StartTime.ToUniversalTime(),
            DeadlineUtc = DateTime.UtcNow.AddMinutes(10),
            Instruction = "Click MathType Display and return an equation. Leave Word open; the operator must then create insertion-returned.signal to request validation.",
        }));
        Console.WriteLine("READY: Click MathType Display and return an equation. Leave Word open until the operator signals validation.");
        var deadline = DateTime.UtcNow.AddMinutes(10);
        while (DateTime.UtcNow < deadline)
        {
            System.Windows.Forms.Application.DoEvents();
            var exited = word.HasExited;
            if (exited)
                throw new InvalidOperationException($"The test Word process exited before validation, code 0x{word.ExitCode:X8}.");
            // Do not poll the COM document model while MathType is inserting.
            // Only the explicit post-insertion operator signal permits reads.
            if (File.Exists(completionSignal))
            {
                var observedMathTypeObjects = CountMathTypeOleShapes(host.Document);
                if (observedMathTypeObjects == 0)
                    throw new InvalidOperationException("No returned MathType object; manual insertion is unverified.");
                host.Save(Path.Combine(artifactRoot, "manual-native-insertion.docx"));
                File.WriteAllText(statePath, JsonSerializer.Serialize(new
                {
                    Status = "Passed", WordProcessId = word.Id,
                    MathTypeObjects = observedMathTypeObjects, CompletedUtc = DateTime.UtcNow,
                }));
                Console.WriteLine("PASS: Physical native MathType insertion returned an OLE object and the owned document saved.");
                return;
            }
            Thread.Sleep(100);
        }
        throw new TimeoutException("Manual coexistence check was not completed in ten minutes.");
        }
        catch (Exception error)
        {
            File.WriteAllText(statePath, JsonSerializer.Serialize(new
            {
                Status = "Failed", Error = error.ToString(), CompletedUtc = DateTime.UtcNow,
            }));
            throw;
        }
    }
}
