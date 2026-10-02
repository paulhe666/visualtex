using System;
using System.IO;
using VisualTeX.WindowsOffice.Contracts;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class VisualTeXAcceptanceIdentityConfigurationTests
{
    [Fact]
    public void DefaultConfigurationKeepsExistingInstallPathAndPort()
    {
        var configuration = new CompanionConfiguration
        {
            AppDataRoot = Path.GetFullPath("normal-runtime-data"),
            CompanionPort = 43127,
        };

        VisualTeXSessionClient.ValidateAcceptanceIdentityConfiguration(configuration);

        Assert.Equal(Path.Combine(configuration.AppDataRoot, "office", "install.json"),
            configuration.InstallJsonPath);
        Assert.Equal(43127, configuration.InstallJsonPort);
    }

    [Fact]
    public void AcceptanceConfigurationReadsSourceIdentityWhileKeepingRuntimeSeparate()
    {
        var configuration = IsolatedConfiguration();

        VisualTeXSessionClient.ValidateAcceptanceIdentityConfiguration(configuration);

        Assert.Equal(Path.Combine(configuration.IdentityRoot, "office", "install.json"),
            configuration.InstallJsonPath);
        Assert.Equal(43127, configuration.InstallJsonPort);
        Assert.Equal(43129, configuration.CompanionPort);
        Assert.EndsWith("isolated-runtime-data", configuration.AppDataRoot);
    }

    [Theory]
    [InlineData("missing-root")]
    [InlineData("missing-port")]
    [InlineData("invalid-port")]
    [InlineData("same-port")]
    [InlineData("same-root")]
    [InlineData("nested-data")]
    [InlineData("nested-identity")]
    [InlineData("relative-root")]
    [InlineData("drive-relative-root")]
    [InlineData("current-drive-root")]
    public void AcceptanceConfigurationRejectsIncompleteOrOverlappingIdentity(string defect)
    {
        var configuration = IsolatedConfiguration();
        switch (defect)
        {
            case "missing-root": configuration.IdentityRoot = string.Empty; break;
            case "missing-port": configuration.IdentityPort = null; break;
            case "invalid-port": configuration.IdentityPort = 0; break;
            case "same-port": configuration.IdentityPort = configuration.CompanionPort; break;
            case "same-root": configuration.IdentityRoot = configuration.AppDataRoot; break;
            case "nested-data": configuration.AppDataRoot = Path.Combine(configuration.IdentityRoot, "child"); break;
            case "nested-identity": configuration.IdentityRoot = Path.Combine(configuration.AppDataRoot, "child"); break;
            case "relative-root": configuration.IdentityRoot = "relative"; break;
            case "drive-relative-root": configuration.IdentityRoot = "C:relative"; break;
            case "current-drive-root": configuration.IdentityRoot = @"\relative"; break;
        }

        var error = Assert.Throws<VisualTeXCompanionException>(() =>
            VisualTeXSessionClient.ValidateAcceptanceIdentityConfiguration(configuration));

        Assert.Equal("AcceptanceIdentityInvalid", error.Diagnostic.FailureType);
    }

    [Fact]
    public void AcceptanceDiagnosticsUseRuntimeArtifactRoot()
    {
        var configuration = IsolatedConfiguration();

        Assert.Equal(Path.Combine(configuration.AppDataRoot, "office", "logs"),
            VisualTeXSessionClient.ResolveCompanionLogRoot("43129", configuration.AppDataRoot));
    }

    [Fact]
    public void NormalDiagnosticsIgnoreInactiveAcceptanceRoot()
    {
        Assert.Equal(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX", "office", "logs"),
            VisualTeXSessionClient.ResolveCompanionLogRoot(null, Path.GetFullPath("unused-artifacts")));
    }

    [Fact]
    public void AcceptanceDiagnosticsDoNotFallBackWhenRuntimeRootIsMissing()
    {
        Assert.Throws<InvalidDataException>(() =>
            VisualTeXSessionClient.ResolveCompanionLogRoot("43129", null));
    }

    private static CompanionConfiguration IsolatedConfiguration() => new()
    {
        AppDataRoot = Path.GetFullPath("isolated-runtime-data"),
        IdentityRoot = Path.GetFullPath("source-identity"),
        CompanionPort = 43129,
        IdentityPort = 43127,
    };
}
