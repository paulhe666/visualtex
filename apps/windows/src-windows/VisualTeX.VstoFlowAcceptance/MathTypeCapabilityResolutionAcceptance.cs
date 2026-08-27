using System.Diagnostics;
using VisualTeX.WordVsto;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunMathTypeCapabilityResolutionAcceptance()
    {
        var preferred = MathTypeOleInterop.ResolvePreferredStorageIdentity();
        AssertTrue(!string.IsNullOrWhiteSpace(preferred.ProgId),
            "MathType preferred storage identity has no ProgID.");
        AssertTrue(preferred.Clsid != Guid.Empty,
            "MathType preferred storage identity has no CLSID.");
        AssertTrue(!string.IsNullOrWhiteSpace(preferred.UserType),
            "MathType preferred storage identity has no user type.");

        var canonicalAvailable = MathTypeOleInterop.TryResolveCapabilities(
            MathTypeOleInterop.CanonicalProgId,
            out var canonical);
        var genericAvailable = MathTypeOleInterop.TryResolveCapabilities(
            "Equation",
            out var generic);

        if (canonicalAvailable)
        {
            AssertTrue(canonical.ResolvedClsid != Guid.Empty,
                "Canonical MathType capability resolution returned an empty CLSID.");
            AssertTrue(MathTypeOleInterop.IsRegisteredMathTypeClass(canonical.ResolvedClsid),
                "Resolved canonical MathType CLSID is not recognized as a MathType class.");
            if (!string.IsNullOrWhiteSpace(canonical.ServerPath))
                AssertTrue(File.Exists(canonical.ServerPath),
                    $"Resolved MathType LocalServer does not exist: {canonical.ServerPath}");
        }
        if (genericAvailable)
        {
            AssertTrue(generic.ResolvedClsid != Guid.Empty,
                "Generic Equation capability resolution returned an empty CLSID.");
            AssertTrue(MathTypeOleInterop.IsRegisteredMathTypeClass(generic.ResolvedClsid),
                "Resolved generic Equation CLSID is not recognized as a MathType class.");
            if (!string.IsNullOrWhiteSpace(generic.ServerPath))
                AssertTrue(File.Exists(generic.ServerPath),
                    $"Resolved generic Equation LocalServer does not exist: {generic.ServerPath}");
        }
        AssertTrue(canonicalAvailable || genericAvailable || preferred.Clsid == MathTypeOleInterop.CanonicalClsid,
            "MathType storage identity has no registered server and did not fall back to the canonical offline CLSID.");

        // Capability lookup runs in formula detection and creation paths. Repeated
        // calls must use the bounded TTL cache rather than reopening registry views
        // and traversing OleGetAutoConvert for every formula in a large document.
        var watch = Stopwatch.StartNew();
        for (var index = 0; index < 2000; index++)
        {
            _ = MathTypeOleInterop.TryResolveCapabilities(
                MathTypeOleInterop.CanonicalProgId,
                out _);
            _ = MathTypeOleInterop.ResolvePreferredStorageIdentity();
        }
        watch.Stop();
        AssertTrue(watch.Elapsed < TimeSpan.FromSeconds(3),
            $"MathType capability lookup is not effectively cached: 2000 probes took {watch.Elapsed.TotalSeconds:F2}s.");

        Console.WriteLine(
            $"[MATHTYPE CAPABILITY RESOLUTION] preferred={preferred.ProgId}/{preferred.Clsid:D}/'{preferred.UserType}', "
            + $"canonicalAvailable={canonicalAvailable}, genericAvailable={genericAvailable}, "
            + $"2000 cached probes={watch.ElapsedMilliseconds}ms.");
    }
}
