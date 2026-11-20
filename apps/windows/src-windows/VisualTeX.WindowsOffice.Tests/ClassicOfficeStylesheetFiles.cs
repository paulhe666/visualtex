using System.Xml.Linq;
using System.Xml.Xsl;
using Microsoft.Win32;

namespace VisualTeX.WindowsOffice.Tests;

// Locating and loading the installed Office Math stylesheets. Only the test
// oracles use them; product code never reads Office's XSL files.
internal static class ClassicOfficeStylesheetFiles
{
    internal static string ResolveTransformPath(string fileName) =>
        ExistingTransformCandidates(fileName).FirstOrDefault()
        ?? throw new FileNotFoundException(
            $"Unable to locate Office {fileName}. Searched the registered Word installation and Office 2007-365 roots. Repair Microsoft Word or reinstall the Office integration.");

    private static IEnumerable<string> ExistingTransformCandidates(string fileName)
    {
        var candidates = new List<string>();
        var overrideRoot = Environment.GetEnvironmentVariable(
            "VISUALTEX_OFFICE_MATH_XSL_ROOT");
        AddCandidateRoot(candidates, overrideRoot, fileName);

        // The stylesheet shipped beside the Office application that is running
        // this add-in (WINWORD.EXE or POWERPNT.EXE share one install folder) is
        // the one that matches this Office. Registry App Paths can name another
        // or stale installation, and the conventional roots below are guesses.
        try
        {
            using var process = System.Diagnostics.Process.GetCurrentProcess();
            var hostDirectory = Path.GetDirectoryName(process.MainModule?.FileName ?? string.Empty);
            if (!string.IsNullOrWhiteSpace(hostDirectory))
                candidates.Add(Path.Combine(hostDirectory!, fileName));
        }
        catch { }

        // App Paths is the authoritative Office location for MSI, Click-to-Run,
        // per-user and alternate-bit installations. Derive the stylesheet folder
        // from every visible 32/64-bit registry view before trying conventional
        // Program Files roots.
        foreach (var wordPath in ReadRegisteredWordPaths())
        {
            var directory = Path.GetDirectoryName(wordPath);
            if (!string.IsNullOrWhiteSpace(directory))
                candidates.Add(Path.Combine(directory!, fileName));
        }

        AddCandidateRoot(candidates, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), fileName);
        AddCandidateRoot(candidates, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), fileName);
        AddCandidateRoot(candidates, Environment.GetEnvironmentVariable("ProgramW6432"), fileName);
        AddCandidateRoot(candidates, Environment.GetEnvironmentVariable("CommonProgramFiles"), fileName);
        AddCandidateRoot(candidates, Environment.GetEnvironmentVariable("CommonProgramFiles(x86)"), fileName);
        AddCandidateRoot(candidates, AppContext.BaseDirectory, fileName);
        return candidates.Distinct(StringComparer.OrdinalIgnoreCase).Where(File.Exists).ToArray();
    }

    // VisualTeX adapts Microsoft's stylesheet at named templates. Any file with
    // the same name (another Office build, a third-party copy found first on
    // the search path) is only a candidate until that adaptation succeeds; the
    // first one every adaptation fits is used. Office builds differ, so when no
    // copy fits them all, the first Microsoft stylesheet that fits some of them
    // is used with stock behaviour for the rest. Report every rejected file
    // when none fits at all.
    internal static XslCompiledTransform LoadFirstUsableTransform(
        string fileName,
        IEnumerable<string> candidates,
        Func<string, bool, XslCompiledTransform> load)
    {
        var paths = candidates.ToArray();
        var rejected = new List<string>();
        foreach (var candidate in paths)
        {
            try { return load(candidate, false); }
            catch (Exception error)
            {
                rejected.Add($"{candidate} ({error.GetType().Name}: {error.Message})");
            }
        }
        foreach (var candidate in paths)
        {
            try { return load(candidate, true); }
            catch { }
        }
        if (rejected.Count == 0)
            throw new FileNotFoundException(
                $"Unable to locate Office {fileName}. Searched the running Office, the registered Word installation and Office 2007-365 roots. Repair Microsoft Office or reinstall the Office integration.");
        throw new InvalidDataException(
            $"No usable Office {fileName}: none of the installed copies has the Microsoft templates VisualTeX adapts. Repair Microsoft Office. Rejected: "
            + string.Join("; ", rejected));
    }

    // Each step edits Microsoft templates on a copy, so a step whose template
    // is absent is never half-applied. A strict load requires every step; a
    // tolerant load skips absent ones but requires at least one, which also
    // tells a Microsoft stylesheet from a foreign file of the same name.
    internal static XDocument AdaptStylesheet(
        XDocument stylesheet,
        bool tolerant,
        params (string Name, Action<XDocument> Apply)[] steps)
    {
        var skipped = new List<string>();
        foreach (var (name, apply) in steps)
        {
            var copy = new XDocument(stylesheet);
            try
            {
                apply(copy);
                stylesheet = copy;
            }
            catch (Exception error) when (
                error is InvalidOperationException or InvalidDataException or ArgumentException)
            {
                if (!tolerant)
                    throw new InvalidDataException(
                        $"Office stylesheet does not have the {name} template VisualTeX adapts ({error.Message})",
                        error);
                skipped.Add(name);
            }
        }
        if (skipped.Count == steps.Length)
            throw new InvalidDataException(
                "The stylesheet has none of the Microsoft templates VisualTeX adapts.");
        if (skipped.Count > 0)
            System.Diagnostics.Trace.TraceWarning(
                "VisualTeX adapted the Office stylesheet without: " + string.Join(", ", skipped));
        return stylesheet;
    }

    private static void AddCandidateRoot(List<string> candidates, string? root, string fileName)
    {
        if (string.IsNullOrWhiteSpace(root)) return;
        var expandedRoot = Environment.ExpandEnvironmentVariables(
            root!.Trim().Trim('"'));
        if (File.Exists(expandedRoot))
        {
            candidates.Add(expandedRoot);
            return;
        }

        candidates.Add(Path.Combine(expandedRoot, fileName));
        foreach (var version in new[] { "Office16", "Office15", "Office14", "Office12" })
        {
            candidates.Add(Path.Combine(
                expandedRoot,
                "Microsoft Office",
                "root",
                version,
                fileName));
            candidates.Add(Path.Combine(
                expandedRoot,
                "Microsoft Office",
                version,
                fileName));
            candidates.Add(Path.Combine(expandedRoot, version, fileName));
            candidates.Add(Path.Combine(
                expandedRoot,
                "Microsoft Shared",
                version.ToUpperInvariant(),
                fileName));
        }
    }

    private static IEnumerable<string> ReadRegisteredWordPaths()
    {
        var views = Environment.Is64BitOperatingSystem
            ? new[] { RegistryView.Registry64, RegistryView.Registry32 }
            : new[] { RegistryView.Registry32 };
        foreach (var view in views)
        {
            foreach (var hive in new[]
                     {
                         RegistryHive.CurrentUser,
                         RegistryHive.LocalMachine,
                     })
            {
                RegistryKey? baseKey = null;
                RegistryKey? appPathKey = null;
                try
                {
                    baseKey = RegistryKey.OpenBaseKey(hive, view);
                    appPathKey = baseKey.OpenSubKey(
                        "Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE");
                    var value = appPathKey?.GetValue(null) as string;
                    if (string.IsNullOrWhiteSpace(value)) continue;
                    var path = NormalizeRegisteredExecutablePath(value!);
                    if (!string.IsNullOrWhiteSpace(path)) yield return path;
                }
                finally
                {
                    appPathKey?.Dispose();
                    baseKey?.Dispose();
                }
            }
        }
    }

    private static string NormalizeRegisteredExecutablePath(string value)
    {
        var source = Environment.ExpandEnvironmentVariables(value.Trim());
        if (source.StartsWith("\"", StringComparison.Ordinal))
        {
            var closing = source.IndexOf('"', 1);
            return closing > 1
                ? source.Substring(1, closing - 1)
                : source.Trim('"');
        }
        var exeEnd = source.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
        return (exeEnd >= 0
                ? source.Substring(0, exeEnd + 4)
                : source.Split(new[] { ' ', '\t' }, 2)[0])
            .Trim()
            .Trim('"');
    }
}
