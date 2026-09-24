using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordCellNumberingClassificationTests
{
    private const string Id = "01234567-89ab-cdef-0123-456789abcdef";
    private const string Prefix = "<w:p><w:r><w:object><o:OLEObject ProgID=\"VisualTeX.Formula.1\" /></w:object></w:r>";
    private const string Ref = "<w:r><w:instrText> REF VTEqNum_0123456789abcdef0123456789abcdef \\h </w:instrText></w:r>";

    [Fact]
    public void UserCellOleWithOwnReferenceIsNotAManagedOmmlTable()
        => Assert.True(WordEquationNumbering.IsOleNumberingParagraphXml(Prefix + Ref + "</w:p>", Id));

    [Fact]
    public void AnotherEquationsNumberCannotClaimTheOle()
        => Assert.False(WordEquationNumbering.IsOleNumberingParagraphXml(Prefix + Ref + "</w:p>", "11111111-2222-3333-4444-555555555555"));

    [Fact]
    public void NativeMathNumberCellDoesNotTakeTheOlePath()
        => Assert.False(WordEquationNumbering.IsOleNumberingParagraphXml("<w:p><m:oMath/>" + Ref + "</w:p>", Id));

    [Fact]
    public void TwoOleObjectsCannotShareOneNumberParagraph()
        => Assert.False(WordEquationNumbering.IsOleNumberingParagraphXml(Prefix + Prefix + Ref + "</w:p>", Id));

    [Fact]
    public void SimilarProgIdIsNotNativeVisualTeX()
        => Assert.False(WordEquationNumbering.IsOleNumberingParagraphXml((Prefix + Ref + "</w:p>").Replace("VisualTeX.Formula.1", "VisualTeXxFormulax1"), Id));

    [Fact]
    public void MalformedFormulaIdentityIsRejected()
        => Assert.False(WordEquationNumbering.IsOleNumberingParagraphXml(Prefix + Ref + "</w:p>", "not-an-id"));
}
