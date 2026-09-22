using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlFingerprintVersionTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static string Math(string body) => "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">" + body + "</m:oMath>";
    private static string Run(string text) => "<m:r><m:t>" + text + "</m:t></m:r>";
    private static string Equation(string text) => Math(Run(text));
    private const string First = "11111111-1111-4111-8111-111111111111";
    private const string Second = "22222222-2222-4222-8222-222222222222";
    private static XElement Paragraph(string id, string bookmarkId, string xml) => new(W + "p",
        new XElement(W + "bookmarkStart", new XAttribute(W + "id", bookmarkId),
            new XAttribute(W + "name", WordOmmlFormulaStore.BookmarkName(id))),
        new XElement(W + "bookmarkEnd", new XAttribute(W + "id", bookmarkId)), XElement.Parse(xml));

    // Values captured from the production test assembly BEFORE changing the
    // fingerprint implementation. These are not calculated by the new algorithm.
    [Theory]
    [InlineData("<m:r><m:t>x</m:t></m:r>", "282569ce1a97d2ead931100049aa0e4acb8dc3686d0668ce60dfa7d3f48bcccb")]
    [InlineData("<m:r><m:t>xy+1</m:t></m:r>", "cc915a593b6643f17f6ff4cc4a61ba3aa9075e816baef6da38ff5efde8492b67")]
    [InlineData("<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>", "a88766462530c621bbab0f85262fc0361311c5beda95b3baa5986ea64e601d36")]
    [InlineData("<m:rad><m:radPr><m:degHide m:val=\"on\"/></m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>", "718a3c0045082860751c66445491946a5fee6f9dee692a5ab75007f4ffc76bb8")]
    public void LegacyDigestsRemainByteCompatible(string body, string golden)
    {
        var xml = Math(body);
        Assert.Equal(golden, WordOmmlConverter.ComputeLegacyOmmlFingerprint(xml));
        Assert.Equal(golden, WordOmmlConverter.ComputeOmmlFingerprintForExpectedVersion(xml, golden));
        Assert.True(WordOmmlConverter.MatchesStoredOmmlFingerprint(xml, golden));
        Assert.Equal(WordOmmlConverter.ComputeOmmlFingerprint(xml), WordOmmlConverter.UpgradeVerifiedOmmlFingerprint(xml, golden));
    }

    [Theory]
    [InlineData("", 1)]
    [InlineData("omml2:", 2)]
    [InlineData("OMML2:", 2)]
    public void SupportedFingerprintFormatsHaveAnExplicitVersion(string prefix, int version)
    {
        var value = prefix + new string('A', 64);
        Assert.True(OmmlFingerprintFormat.IsSupported(value));
        Assert.Equal(version, OmmlFingerprintFormat.GetVersion(value));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("omml3:")]
    [InlineData("v2:")]
    [InlineData("omml2: xyz")]
    [InlineData("   ")]
    public void UnknownOrMalformedVersionsAreNotSilentlyGuessed(string? value)
    {
        Assert.False(OmmlFingerprintFormat.IsSupported(value));
        Assert.Throws<InvalidOperationException>(() => OmmlFingerprintFormat.GetVersion(value!));
    }

    [Fact]
    public void WrongDigestLengthsAndNonHexCharactersAreRejected()
    {
        foreach (var value in new[] { new string('a', 63), new string('a', 65), new string('g', 64),
                     "omml2:" + new string('a', 63), "omml3:" + new string('a', 64), " " + new string('a', 64) })
        {
            Assert.False(OmmlFingerprintFormat.IsSupported(value));
            Assert.Throws<InvalidOperationException>(() => WordOmmlConverter.ComputeOmmlFingerprintForExpectedVersion(Equation("x"), value));
        }
    }

    [Fact]
    public void NewPersistentIdentityUsesTheImportContentModel()
    {
        var xml = Equation("x");
        var alias = xml.Replace("xmlns:m=", "xmlns:q=").Replace("m:", "q:");
        var expected = WordOmmlConverter.ComputeOmmlFingerprint(xml);
        Assert.Equal("omml2:" + WordOmmlConverter.ComputeImportedOmmlContentSignature(xml), expected);
        Assert.True(WordOmmlConverter.MatchesStoredOmmlFingerprint(alias, expected));
        Assert.Equal(expected, WordOmmlConverter.ComputeVerifiedMaterializedOmmlFingerprint(xml, alias));
        Assert.NotEqual(WordOmmlConverter.ComputeLegacyOmmlFingerprint(xml), WordOmmlConverter.ComputeLegacyOmmlFingerprint(alias));
    }

    [Theory]
    [InlineData("1")]
    [InlineData("true")]
    [InlineData("on")]
    [InlineData(null)]
    public void EquivalentHiddenSlotsRemainTheSameAfterPersistence(string? value)
    {
        string Root(string? flag) => Math("<m:rad><m:radPr><m:degHide"
            + (flag is null ? "" : " m:val=\"" + flag + "\"")
            + "/></m:radPr><m:deg/><m:e>" + Run("x") + "</m:e></m:rad>");
        var expected = WordOmmlConverter.ComputeOmmlFingerprint(Root("1"));
        Assert.True(WordOmmlConverter.MatchesStoredOmmlFingerprint(Root(value), expected));
        Assert.False(WordOmmlConverter.MatchesStoredOmmlFingerprint(Root("0"), expected));
    }

    [Fact]
    public void AChangedEquationCannotUpgradeALegacyIdentity()
    {
        var legacy = WordOmmlConverter.ComputeLegacyOmmlFingerprint(Equation("x"));
        Assert.False(WordOmmlConverter.MatchesStoredOmmlFingerprint(Equation("y"), legacy));
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.UpgradeVerifiedOmmlFingerprint(Equation("y"), legacy));
        Assert.False(WordOmmlConverter.MatchesStoredOmmlFingerprint(Equation("x"), null));
    }

    [Theory]
    [InlineData("x", "y")]
    [InlineData("x", "x ")]
    [InlineData("x", "xx")]
    public void NewVersionStillRejectsRealContentChanges(string from, string to)
    {
        Assert.False(WordOmmlConverter.MatchesStoredOmmlFingerprint(Equation(to), WordOmmlConverter.ComputeOmmlFingerprint(Equation(from))));
    }

    [Fact]
    public void LegacyAndNewCopiesShareOneVerifiedLogicalCountKey()
    {
        var xml = Equation("x");
        var old = WordOmmlConverter.ComputeLegacyOmmlFingerprint(xml);
        var current = WordOmmlConverter.ComputeOmmlFingerprint(xml);
        var identities = new[] { old, current, old }.Select(value => WordOmmlConverter.UpgradeVerifiedOmmlFingerprint(xml, value)).ToArray();
        Assert.Single(identities.Distinct());
        Assert.Equal(3, identities.Count(value => value == current));
    }

    [Fact]
    public void MixedVersionsKeepIdenticalFormulaOwnersSeparate()
    {
        var xml = Equation("x");
        var body = new XElement(W + "body", Paragraph(First, "1", xml), Paragraph(Second, "2", xml));
        var expected = new Dictionary<string, string>
        {
            [First] = WordOmmlConverter.ComputeLegacyOmmlFingerprint(xml),
            [Second] = WordOmmlConverter.ComputeOmmlFingerprint(xml),
        };
        var result = WordOmmlNativeSource.IndexConversionEquationIdentities(body, expected);
        Assert.Equal((0, true), result[First]);
        Assert.Equal((1, true), result[Second]);
    }

    [Fact]
    public void MixedVersionsCannotClaimOnePhysicalEquationTwice()
    {
        var xml = Equation("x");
        var body = new XElement(W + "body", Paragraph(First, "1", xml), Paragraph(Second, "2", Equation("y")));
        Assert.Throws<InvalidDataException>(() => WordOmmlNativeSource.IndexConversionEquationIdentities(body,
            new Dictionary<string, string>
            {
                [First] = WordOmmlConverter.ComputeLegacyOmmlFingerprint(xml),
                [Second] = WordOmmlConverter.ComputeOmmlFingerprint(xml),
            }, liveInsertedOwners: new Dictionary<string, int> { [First] = 0, [Second] = 0 }));
    }

    [Fact]
    public void MixedVersionDriftCannotUseAnUnrelatedMatchingEquation()
    {
        var body = new XElement(W + "body", Paragraph(First, "1", Equation("y")), Paragraph(Second, "2", Equation("x")));
        Assert.Throws<InvalidDataException>(() => WordOmmlNativeSource.IndexConversionEquationIdentities(body,
            new Dictionary<string, string> { [First] = WordOmmlConverter.ComputeLegacyOmmlFingerprint(Equation("x")) },
            exactAnchorOwners: new Dictionary<string, int> { [First] = 0 }));
    }

    [Fact]
    public void ThousandMixedVersionIdentitiesUseTheCorrectPhysicalOwners()
    {
        var body = new XElement(W + "body");
        var expected = new Dictionary<string, string>();
        for (var i = 0; i < 1000; i++)
        {
            var id = Guid.NewGuid().ToString();
            var xml = Equation("x" + (i % 7)); // Many identical contents, distinct anchors.
            body.Add(Paragraph(id, i.ToString(), xml));
            expected[id] = i % 2 == 0 ? WordOmmlConverter.ComputeLegacyOmmlFingerprint(xml) : WordOmmlConverter.ComputeOmmlFingerprint(xml);
        }
        var result = WordOmmlNativeSource.IndexConversionEquationIdentities(body, expected);
        Assert.Equal(1000, result.Count);
        Assert.Equal(Enumerable.Range(0, 1000), expected.Keys.Select(key => result[key].Index));
    }

    [Fact]
    public void MerelyReadingALegacyEquationPreservesOriginalLatexAndStoredIdentity()
    {
        var xml = Equation("x");
        var metadata = Metadata(WordOmmlConverter.ComputeLegacyOmmlFingerprint(xml));
        var before = FormulaMetadataCodec.Encode(metadata);
        var refreshed = WordOmmlNativeSource.RefreshForVisualTeXFromCapturedWordOpenXml(metadata, xml);
        Assert.NotSame(metadata, refreshed);
        Assert.Equal(metadata.Latex, refreshed.Latex);
        Assert.Equal(metadata.NativeOmmlFingerprint, refreshed.NativeOmmlFingerprint);
        Assert.Equal(before, FormulaMetadataCodec.Encode(metadata));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MetadataSerializationPreservesBothVersions(bool legacy)
    {
        var xml = Equation("x");
        var expected = legacy ? WordOmmlConverter.ComputeLegacyOmmlFingerprint(xml) : WordOmmlConverter.ComputeOmmlFingerprint(xml);
        var metadata = Metadata(expected);
        var roundTrip = FormulaMetadataCodec.Decode(FormulaMetadataCodec.Encode(metadata));
        Assert.NotNull(roundTrip);
        Assert.Equal(expected, roundTrip.NativeOmmlFingerprint);
        Assert.True(WordOmmlConverter.MatchesStoredOmmlFingerprint(xml, roundTrip.NativeOmmlFingerprint));
    }

    [Fact]
    public void EquivalentRunPropertyAliasesDoNotPreventContentGrouping()
    {
        const string p = "<m:rPr><m:sty m:val=\"b\"/></m:rPr>";
        const string q = "<q:rPr xmlns:q=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><q:sty q:val=\"b\"/></q:rPr>";
        var split = Math("<m:r>" + p + "<m:t>x</m:t></m:r><m:r>" + q + "<m:t>y</m:t></m:r>");
        var joined = Math("<m:r>" + p + "<m:t>xy</m:t></m:r>");
        Assert.Equal(WordOmmlConverter.ComputeOmmlFingerprint(joined), WordOmmlConverter.ComputeOmmlFingerprint(split));
    }

    [Fact]
    public void EmptyNamespaceOnlyPropertiesAreNotMathematicalContent()
    {
        var properties = Equation("x").Replace("<m:r>", "<m:r><q:rPr xmlns:q=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"/>");
        Assert.Equal(WordOmmlConverter.ComputeOmmlFingerprint(Equation("x")), WordOmmlConverter.ComputeOmmlFingerprint(properties));
    }

    [Fact]
    public void RunGroupingCannotSilentlyDiscardWordControlChildren()
    {
        var withControl = Math(Run("x") + "<m:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:br/><m:t>y</m:t></m:r>");
        Assert.NotEqual(WordOmmlConverter.ComputeOmmlFingerprint(Equation("xy")), WordOmmlConverter.ComputeOmmlFingerprint(withControl));
    }

    [Fact]
    public void RunGroupingCannotSilentlyDiscardUnknownTextAttributes()
    {
        var annotated = Math(Run("x") + "<m:r><m:t xmlns:test=\"urn:visualtex:audit\" test:token=\"retained\">y</m:t></m:r>");
        Assert.NotEqual(WordOmmlConverter.ComputeOmmlFingerprint(Equation("xy")), WordOmmlConverter.ComputeOmmlFingerprint(annotated));
    }

    [Theory]
    [InlineData("<m:r><m:t>x</m:t></m:r>", "46c08e0421e16bb81dcccc94326d1d5e6312fa09058009b086dc1f0218c87ba9")]
    [InlineData("<m:r><m:t>xy+1</m:t></m:r>", "c3a67f0c40bdfb4f9f709982ca7ea3358aec1f522285c89457d17edf05b57a84")]
    [InlineData("<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>", "07ba310dfd9aa321bca9ff1a245ac24653ce24122ada67bc800d9ff432a5e7fb")]
    [InlineData("<m:rad><m:radPr><m:degHide m:val=\"on\"/></m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>", "d6c1b3110cfbdcd5f1afa06558624004a31eb29afec21ba77ed2332f86de8f99")]
    public void VersionTwoDigestValuesCannotChangeWithoutVersionReview(string body, string digest)
        => Assert.Equal("omml2:" + digest, WordOmmlConverter.ComputeOmmlFingerprint(Math(body)));

    private static FormulaMetadata Metadata(string fingerprint) => new()
    {
        FormulaId = First, Latex = "{x}", CodeFormat = "raw", DisplayMode = "inline",
        Lines = new() { new FormulaLine { Id = Second, Latex = "{x}" } },
        NativeOmmlFingerprint = fingerprint,
        CreatedAt = "2026-09-19T00:00:00Z", UpdatedAt = "2026-09-19T00:00:00Z",
        CreatedWithVersion = "1.2.8", UpdatedWithVersion = "1.2.8",
    };
}
