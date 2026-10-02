using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Prepares literal native-math text before a target Word undo transaction.
/// Range.Text is math input and Functions.Add(NormalText) is not a persistent
/// literal-run transport. All prefix writers copy a proven m:nor run instead.
/// One operation owns one hidden document; only the target is written in Undo.
/// </summary>
internal sealed class WordNativeOmmlLiteralTextStaging : IDisposable
{
    private readonly Document _target;
    private Document? _staging;
    private readonly List<OMath> _owners = new();

    internal WordNativeOmmlLiteralTextStaging(Document target)
    {
        _target = target ?? throw new ArgumentNullException(nameof(target));
    }

    internal Range Prepare(string value, WordCharacterFormatting formatting)
    {
        if (WordFormulaMutationTransaction.IsActiveOnCurrentThread)
            throw new InvalidOperationException("Literal native text must be prepared before target Undo.");
        if (string.IsNullOrEmpty(value) || value.Any(char.IsControl))
            throw new ArgumentException("A native literal prefix must be nonempty plain text.", nameof(value));
        Range? content = null;
        Range? slot = null;
        Range? inserted = null;
        Range? exact = null;
        Range? number = null;
        Range? code = null;
        Fields? fields = null;
        Field? field = null;
        OMaths? maths = null;
        OMath? math = null;
        try
        {
            if (_staging is null)
            {
                _staging = _target.Application.Documents.Add(Visible: false);
                _staging.TrackRevisions = false;
                _staging.OMathFontName = _target.OMathFontName;
            }
            content = _staging.Content;
            var start = content.End - 1;
            slot = _staging.Range(start, start);
            slot.InsertAfter("\uE101\r");
            slot.SetRange(start, start + 1);
            // Use the same canonical m:nor construction as native numbered
            // creation. The private SEQ only gives the literal an exact right
            // boundary; neither that field nor the dummy body is transferred.
            const string body = "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>";
            var prepared = WordOmmlConverter.BuildWordNativeNumberedOmml(
                body, "Equation", null, 1, string.Empty, "1", value);
            inserted = WordOmmlConverter.ReplaceWithPreparedOmmlDirect(
                _staging, slot, prepared, display: true, mathFontName: _staging.OMathFontName);
            maths = inserted.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException("A prepared literal must have one private native owner.");
            math = maths[1]; exact = math.Range;
            number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
            fields = number.Fields;
            if (fields.Count != 1)
                throw new InvalidDataException("A prepared literal lost its private field boundary.");
            field = fields[1]; code = field.Code;
            if (field.Type != WdFieldType.wdFieldSequence || !code.InStory(number)
                || code.Start <= number.Start || code.End >= number.End)
                throw new InvalidDataException("The private literal field escaped its native argument.");
            number.End = code.Start - 1;
            if (number.Text != value)
                throw new InvalidDataException("Word changed a canonical literal during preparation.");
            formatting.ApplyToNativeMath(number);
            var result = number.Duplicate;
            _owners.Add(math); math = null;
            return result;
        }
        finally
        {
            Release(code); Release(field); Release(fields); Release(number);
            Release(exact); Release(math); Release(maths); Release(inserted); Release(slot); Release(content);
            _target.Activate();
        }
    }

    public void Dispose()
    {
        foreach (var owner in _owners) Release(owner);
        _owners.Clear();
        if (_staging is null) return;
        try { _staging.Close(WdSaveOptions.wdDoNotSaveChanges); }
        finally { Release(_staging); _staging = null; }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}
