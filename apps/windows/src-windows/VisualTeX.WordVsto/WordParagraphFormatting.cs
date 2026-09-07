using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>Preserves the identified paragraph's style and direct paragraph format
/// while its generated formula host is replaced. Character formatting is separate.</summary>
internal sealed class WordParagraphFormatting : IDisposable
{
    private readonly string styleName;
    private ParagraphFormat? format;

    private WordParagraphFormatting(string styleName, ParagraphFormat format)
    {
        this.styleName = styleName;
        this.format = format;
    }

    internal static WordParagraphFormatting Capture(Range owner)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        Style? style = null;
        ParagraphFormat? paragraphFormat = null;
        try
        {
            paragraphs = owner.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException("Paragraph formatting requires one identified owner.");
            paragraph = paragraphs[1];
            range = paragraph.Range;
            style = range.get_Style() as Style
                ?? throw new InvalidDataException("The formula host has no identifiable paragraph style.");
            paragraphFormat = range.ParagraphFormat;
            return new WordParagraphFormatting(style.NameLocal, paragraphFormat.Duplicate);
        }
        finally
        {
            Release(paragraphFormat);
            Release(style);
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal void Apply(Range owner)
    {
        if (format is null) throw new ObjectDisposedException(nameof(WordParagraphFormatting));
        Paragraphs? paragraphs = null;
        try
        {
            paragraphs = owner.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException("The replacement no longer owns one paragraph.");
            object style = styleName;
            owner.set_Style(ref style);
            owner.ParagraphFormat = format;
        }
        finally { Release(paragraphs); }
    }

    public void Dispose()
    {
        Release(format);
        format = null;
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
            Marshal.ReleaseComObject(value);
    }
}
