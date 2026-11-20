using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// The add-ins also ship as x86 for 32-bit Office, where user32.dll has no
// Get/SetWindowLongPtrW or Get/SetClassLongPtrW exports (they are macros for the
// 32-bit functions). Calling one there throws EntryPointNotFoundException, as
// Word format conversion did on 32-bit Office. The test host runs x64 only, so
// this checks the sources: every such import must sit behind an IntPtr.Size switch.
public sealed class Win32PointerSizeImportTests
{
    private static string SourcesRoot([CallerFilePath] string path = "") =>
        Path.GetFullPath(Path.Combine(Path.GetDirectoryName(path)!, ".."));

    private static readonly Regex X64OnlyImport = new(
        @"\[(?:System\.Runtime\.InteropServices\.)?DllImport\(""user32(?:\.dll)?""[^\]]*\]\s*(?:private|internal|public)?\s*static\s+extern\s+\S+\s+(?<name>\w+)\s*\(",
        RegexOptions.Singleline);

    [Fact]
    public void X64OnlyUser32ImportsAreSwitchedOnPointerSize()
    {
        var offenders = new List<string>();
        foreach (var file in Directory.EnumerateFiles(SourcesRoot(), "*.cs", SearchOption.AllDirectories))
        {
            var relative = Path.GetRelativePath(SourcesRoot(), file);
            if (relative.Split(Path.DirectorySeparatorChar).Any(part =>
                    part.StartsWith("bin", StringComparison.OrdinalIgnoreCase) ||
                    part.StartsWith("obj", StringComparison.OrdinalIgnoreCase) ||
                    part.EndsWith(".Tests", StringComparison.Ordinal)))
                continue;
            var source = File.ReadAllText(file);
            foreach (Match match in X64OnlyImport.Matches(source))
            {
                var entry = Regex.Match(match.Value, @"EntryPoint\s*=\s*""(\w+)""");
                var name = entry.Success ? entry.Groups[1].Value : match.Groups["name"].Value;
                if (!Regex.IsMatch(name, @"^(Get|Set)(Window|Class)LongPtr[AW]?$")) continue;
                if (!source.Contains("IntPtr.Size == 8", StringComparison.Ordinal))
                    offenders.Add($"{relative}: {name}");
            }
        }
        Assert.Empty(offenders);
    }
}
