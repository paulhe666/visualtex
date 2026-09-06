using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlConversionIdentityTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string First = "11111111-1111-4111-8111-111111111111";
    private const string Second = "22222222-2222-4222-8222-222222222222";
    private static XElement Math(string text) => new(M + "oMath", new XElement(M + "r", new XElement(M + "t", text)));
    private static string Fingerprint(string text) => WordOmmlConverter.ComputeOmmlFingerprint(Math(text).ToString());
    private static XElement Start(string id, string name) => new(W + "bookmarkStart", new XAttribute(W + "id", id), new XAttribute(W + "name", name));
    private static XElement End(string id) => new(W + "bookmarkEnd", new XAttribute(W + "id", id));
    private static object[] Anchor(string formula, string id) => new object[] { Start(id, WordOmmlFormulaStore.BookmarkName(formula)), End(id) };
    private static XElement Row(string formula, string text, string id) => new(W + "tr",
        new XElement(W + "tc", new XElement(W + "p")),
        new XElement(W + "tc", new XElement(W + "p", Math(text))),
        new XElement(W + "tc", new XElement(W + "p",
            Start(id, WordEquationNumbering.NativeNumberBookmarkName(formula)),
            new XElement(W + "r", new XElement(W + "t", "1")), End(id))));

    [Fact]
    public void CollapsedAnchorInSameParagraphOwnsItsFormula()
    {
        var body = new XElement(W + "body", new XElement(W + "p", Anchor(First, "1"), Math("a")));
        var result = WordOmmlNativeSource.IndexConversionEquationIdentities(body, new Dictionary<string, string> { [First] = Fingerprint("a") });
        Assert.Equal((0, true), result[First]);
    }

    [Fact]
    public void DisplacedAnchorCannotOverwriteFollowingFormulaFingerprint()
    {
        var body = new XElement(W + "body", Anchor(First, "1"),
            new XElement(W + "p", Math("b")), new XElement(W + "p", Math("a")));
        var result = WordOmmlNativeSource.IndexConversionEquationIdentities(body, new Dictionary<string, string> { [First] = Fingerprint("a") });
        Assert.Equal((1, false), result[First]);
    }

    [Fact]
    public void IdenticalContentsRemainDistinctThroughNumberedRowIdentity()
    {
        var body = new XElement(W + "body", new XElement(W + "tbl", Row(First, "a", "1"), Row(Second, "a", "2")));
        var result = WordOmmlNativeSource.IndexConversionEquationIdentities(body, new Dictionary<string, string>
        { [First] = Fingerprint("a"), [Second] = Fingerprint("a") });
        Assert.Equal(0, result[First].Index);
        Assert.Equal(1, result[Second].Index);
    }

    [Fact]
    public void NumberedRowContentMismatchCannotRecoverFromAnotherRow()
    {
        var body = new XElement(W + "body", new XElement(W + "tbl", Row(First, "b", "1"), Row(Second, "a", "2")));
        Assert.Throws<InvalidDataException>(() => WordOmmlNativeSource.IndexConversionEquationIdentities(body,
            new Dictionary<string, string> { [First] = Fingerprint("a") }));
    }

    [Fact]
    public void DriftRecoveryRequiresGlobalUniqueness()
    {
        var body = new XElement(W + "body", Anchor(First, "1"),
            new XElement(W + "p", Math("a")), new XElement(W + "p", Math("a")));
        Assert.Throws<InvalidDataException>(() => WordOmmlNativeSource.IndexConversionEquationIdentities(body,
            new Dictionary<string, string> { [First] = Fingerprint("a") }));
    }

    [Fact]
    public void TwoAnchorsCannotClaimOneEquation()
    {
        var body = new XElement(W + "body", new XElement(W + "p", Anchor(First, "1"), Anchor(Second, "2"), Math("a")));
        Assert.Throws<InvalidDataException>(() => WordOmmlNativeSource.IndexConversionEquationIdentities(body,
            new Dictionary<string, string> { [First] = Fingerprint("a"), [Second] = Fingerprint("a") }));
    }

    [Fact]
    public void ExactLiveOwnersDisambiguateIdenticalEquationsWithBodyLevelBookmarks()
    {
        var body = new XElement(W + "body", Anchor(First, "1"), new XElement(W + "p", Math("a")),
            Anchor(Second, "2"), new XElement(W + "p", Math("a")));
        var result = WordOmmlNativeSource.IndexConversionEquationIdentities(body,
            new Dictionary<string, string> { [First] = Fingerprint("a"), [Second] = Fingerprint("a") },
            exactAnchorOwners: new Dictionary<string, int> { [First] = 0, [Second] = 1 });
        Assert.Equal((0, true), result[First]);
        Assert.Equal((1, true), result[Second]);
    }

    [Fact]
    public void ExactAnchorContentMismatchCannotRecoverFromAnotherEquation()
    {
        var body = new XElement(W + "body", Anchor(First, "1"),
            new XElement(W + "p", Math("b")), new XElement(W + "p", Math("a")));
        Assert.Throws<InvalidDataException>(() => WordOmmlNativeSource.IndexConversionEquationIdentities(body,
            new Dictionary<string, string> { [First] = Fingerprint("a") },
            exactAnchorOwners: new Dictionary<string, int> { [First] = 0 }));
    }

    [Fact]
    public void ExactAnchorMustAgreeWithNumberedRowEvenWhenContentsMatch()
    {
        var body = new XElement(W + "body", Anchor(First, "3"),
            new XElement(W + "tbl", Row(First, "a", "1"), Row(Second, "a", "2")));
        Assert.Throws<InvalidDataException>(() => WordOmmlNativeSource.IndexConversionEquationIdentities(body,
            new Dictionary<string, string> { [First] = Fingerprint("a") },
            exactAnchorOwners: new Dictionary<string, int> { [First] = 1 }));
    }

    [Fact]
    public void ExactAnchorsCannotClaimTheSamePhysicalEquation()
    {
        var body = new XElement(W + "body", Anchor(First, "1"), Anchor(Second, "2"), new XElement(W + "p", Math("a")));
        Assert.Throws<InvalidDataException>(() => WordOmmlNativeSource.IndexConversionEquationIdentities(body,
            new Dictionary<string, string> { [First] = Fingerprint("a"), [Second] = Fingerprint("a") },
            exactAnchorOwners: new Dictionary<string, int> { [First] = 0, [Second] = 0 }));
    }

    [Fact]
    public void RetainedFreshObjectRepairsInsertionGravityWithoutChangingContent()
    {
        var body = new XElement(W + "body", Anchor(First, "1"),
            new XElement(W + "p", Math("b")), new XElement(W + "p", Math("a")));
        var result = WordOmmlNativeSource.IndexConversionEquationIdentities(body,
            new Dictionary<string, string> { [First] = Fingerprint("a") },
            exactAnchorOwners: new Dictionary<string, int> { [First] = 0 },
            liveInsertedOwners: new Dictionary<string, int> { [First] = 1 });
        Assert.Equal((1, false), result[First]);
    }

    [Fact]
    public void RetainedFreshObjectStillRequiresItsPreparedContent()
    {
        var body = new XElement(W + "body", Anchor(First, "1"),
            new XElement(W + "p", Math("b")), new XElement(W + "p", Math("a")));
        Assert.Throws<InvalidDataException>(() => WordOmmlNativeSource.IndexConversionEquationIdentities(body,
            new Dictionary<string, string> { [First] = Fingerprint("a") },
            liveInsertedOwners: new Dictionary<string, int> { [First] = 0 }));
    }

    [Fact]
    public void RetainedFreshObjectCannotClaimAnotherNumberedRow()
    {
        var body = new XElement(W + "body", Anchor(First, "3"),
            new XElement(W + "tbl", Row(First, "a", "1"), Row(Second, "a", "2")));
        Assert.Throws<InvalidDataException>(() => WordOmmlNativeSource.IndexConversionEquationIdentities(body,
            new Dictionary<string, string> { [First] = Fingerprint("a") },
            liveInsertedOwners: new Dictionary<string, int> { [First] = 1 }));
    }

    [Fact]
    public void LiveAnchorMissingFromXmlIsRejected()
    {
        var body = new XElement(W + "body", new XElement(W + "p", Math("a")));
        Assert.Throws<InvalidDataException>(() => WordOmmlNativeSource.IndexConversionEquationIdentities(body,
            new Dictionary<string, string> { [First] = Fingerprint("a") },
            exactAnchorOwners: new Dictionary<string, int> { [First] = 0 }));
    }
}
