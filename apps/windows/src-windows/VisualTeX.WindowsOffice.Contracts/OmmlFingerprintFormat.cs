using System;

namespace VisualTeX.WindowsOffice.Contracts;

/// <summary>
/// Self-describing native equation fingerprints. Unprefixed 64-hex values are
/// the legacy XML spelling hash. New values use the versioned content model.
/// Never infer mathematical content from a legacy digest or guess a version.
/// </summary>
public static class OmmlFingerprintFormat
{
    public const string CanonicalPrefix = "omml2:";

    public static bool IsSupported(string? value) => GetVersionOrZero(value) != 0;

    public static int GetVersion(string value)
    {
        var version = GetVersionOrZero(value);
        if (version == 0)
            throw new InvalidOperationException("Unsupported or malformed native OMML fingerprint.");
        return version;
    }

    private static int GetVersionOrZero(string? value)
    {
        if (value is null) return 0;
        var start = 0;
        var version = 1;
        if (value.StartsWith(CanonicalPrefix, StringComparison.OrdinalIgnoreCase))
        {
            start = CanonicalPrefix.Length;
            version = 2;
        }
        if (value.Length - start != 64) return 0;
        for (var i = start; i < value.Length; i++)
        {
            var c = value[i];
            if (!((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F')))
                return 0;
        }
        return version;
    }
}
