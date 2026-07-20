using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;

namespace VisualTeX.WindowsOffice.VstoShared;

internal static class VstoDependencyResolver
{
    private static readonly HashSet<string> AllowedAssemblyNames = new(
        StringComparer.OrdinalIgnoreCase)
    {
        "VisualTeX.WindowsOffice.Contracts",
        "System.Text.Json",
        "System.Text.Encodings.Web",
        "Microsoft.Bcl.AsyncInterfaces",
        "System.Memory",
        "System.Buffers",
        "System.Numerics.Vectors",
        "System.ValueTuple",
        "System.Runtime.CompilerServices.Unsafe",
        "System.Threading.Tasks.Extensions",
        "Microsoft.Office.Interop.Word",
        "Microsoft.Office.Interop.PowerPoint",
    };

    private static int _installed;

    public static void Install()
    {
        if (Interlocked.Exchange(ref _installed, 1) != 0)
            return;

        AppDomain.CurrentDomain.AssemblyResolve += ResolveFromAddInDirectory;
    }

    private static Assembly? ResolveFromAddInDirectory(object? sender, ResolveEventArgs args)
    {
        AssemblyName requested;
        try
        {
            requested = new AssemblyName(args.Name);
        }
        catch
        {
            return null;
        }

        if (string.IsNullOrWhiteSpace(requested.Name)
            || !AllowedAssemblyNames.Contains(requested.Name))
            return null;

        foreach (var loaded in AppDomain.CurrentDomain.GetAssemblies())
        {
            AssemblyName loadedName;
            try
            {
                loadedName = loaded.GetName();
            }
            catch
            {
                continue;
            }

            if (MatchesIdentityIgnoringVersion(requested, loadedName))
                return loaded;
        }

        var location = typeof(VstoDependencyResolver).Assembly.Location;
        var directory = Path.GetDirectoryName(location);
        if (string.IsNullOrWhiteSpace(directory))
            return null;

        var candidatePath = Path.Combine(directory, requested.Name + ".dll");
        if (!File.Exists(candidatePath))
            return null;

        AssemblyName candidateName;
        try
        {
            candidateName = AssemblyName.GetAssemblyName(candidatePath);
        }
        catch
        {
            return null;
        }

        if (!MatchesIdentityIgnoringVersion(requested, candidateName))
            return null;

        try
        {
            return Assembly.LoadFrom(candidatePath);
        }
        catch
        {
            return null;
        }
    }

    private static bool MatchesIdentityIgnoringVersion(
        AssemblyName requested,
        AssemblyName candidate)
    {
        if (!string.Equals(requested.Name, candidate.Name, StringComparison.OrdinalIgnoreCase))
            return false;

        var requestedCulture = NormalizeCulture(requested.CultureName);
        var candidateCulture = NormalizeCulture(candidate.CultureName);
        if (!string.Equals(requestedCulture, candidateCulture, StringComparison.OrdinalIgnoreCase))
            return false;

        return TokensEqual(requested.GetPublicKeyToken(), candidate.GetPublicKeyToken());
    }

    private static string NormalizeCulture(string? culture) =>
        string.IsNullOrWhiteSpace(culture) ? "neutral" : culture!;

    private static bool TokensEqual(byte[]? left, byte[]? right)
    {
        left ??= Array.Empty<byte>();
        right ??= Array.Empty<byte>();
        if (left.Length != right.Length)
            return false;

        var difference = 0;
        for (var index = 0; index < left.Length; index++)
            difference |= left[index] ^ right[index];
        return difference == 0;
    }
}
