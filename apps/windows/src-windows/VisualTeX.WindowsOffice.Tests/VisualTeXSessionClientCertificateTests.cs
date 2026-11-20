using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class VisualTeXSessionClientCertificateTests
{
    [Fact]
    public void ExactPerUserThumbprintIsAcceptedAcrossFormattingDifferences()
    {
        var now = DateTime.UtcNow;
        using var certificate = CreateCertificate(
            now.AddDays(-1),
            now.AddDays(30),
            "CN=VisualTeX Local Companion");
        var formattedThumbprint = string.Join(
            " ",
            certificate.Thumbprint!.ToLowerInvariant().Chunk(2).Select(chars => new string(chars)));

        var accepted = VisualTeXSessionClient.IsPinnedCertificateAccepted(
            certificate,
            formattedThumbprint,
            now,
            out var timeValid,
            out var thumbprintMatches);

        Assert.True(accepted);
        Assert.True(timeValid);
        Assert.True(thumbprintMatches);
    }

    [Fact]
    public void DifferentMachineOrInterceptingCertificateIsRejected()
    {
        var now = DateTime.UtcNow;
        using var expectedCertificate = CreateCertificate(
            now.AddDays(-1),
            now.AddDays(30),
            "CN=VisualTeX Expected Companion");
        using var differentCertificate = CreateCertificate(
            now.AddDays(-1),
            now.AddDays(30),
            "CN=Unexpected Local TLS Interceptor");

        var accepted = VisualTeXSessionClient.IsPinnedCertificateAccepted(
            differentCertificate,
            expectedCertificate.Thumbprint,
            now,
            out var timeValid,
            out var thumbprintMatches);

        Assert.False(accepted);
        Assert.True(timeValid);
        Assert.False(thumbprintMatches);
    }

    [Theory]
    [InlineData(-30, -1)]
    [InlineData(1, 30)]
    public void ExpiredOrNotYetValidCertificateIsRejected(
        int notBeforeOffsetDays,
        int notAfterOffsetDays)
    {
        var now = DateTime.UtcNow;
        using var certificate = CreateCertificate(
            now.AddDays(notBeforeOffsetDays),
            now.AddDays(notAfterOffsetDays),
            "CN=VisualTeX Invalid Time Window");

        var accepted = VisualTeXSessionClient.IsPinnedCertificateAccepted(
            certificate,
            certificate.Thumbprint,
            now,
            out var timeValid,
            out var thumbprintMatches);

        Assert.False(accepted);
        Assert.False(timeValid);
        Assert.True(thumbprintMatches);
    }

    [Fact]
    public void MissingRegistryThumbprintIsRejected()
    {
        var now = DateTime.UtcNow;
        using var certificate = CreateCertificate(
            now.AddDays(-1),
            now.AddDays(30),
            "CN=VisualTeX Missing Registry State");

        var accepted = VisualTeXSessionClient.IsPinnedCertificateAccepted(
            certificate,
            string.Empty,
            now,
            out var timeValid,
            out var thumbprintMatches);

        Assert.False(accepted);
        Assert.True(timeValid);
        Assert.False(thumbprintMatches);
    }

    private static X509Certificate2 CreateCertificate(
        DateTime notBeforeUtc,
        DateTime notAfterUtc,
        string subject)
    {
        using var rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            subject,
            rsa,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(
            new X509BasicConstraintsExtension(false, false, 0, false));
        request.CertificateExtensions.Add(
            new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, false));
        return request.CreateSelfSigned(
            new DateTimeOffset(DateTime.SpecifyKind(notBeforeUtc, DateTimeKind.Utc)),
            new DateTimeOffset(DateTime.SpecifyKind(notAfterUtc, DateTimeKind.Utc)));
    }
}
