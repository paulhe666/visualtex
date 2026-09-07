using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>Character appearance shared by body references and generated hosts.
/// Undefined values mean mixed user formatting, not a failed COM read.</summary>
internal sealed class WordCharacterFormatting
{
    private readonly List<Action<Microsoft.Office.Interop.Word.Font>> setters = new();
    internal WdColor? Color { get; set; }

    internal static WordCharacterFormatting CaptureParagraphMarkAtPosition(Document document, int position)
    {
        var paragraph = ResolveMainStoryParagraphAtPosition(document, position);
        try { return CaptureParagraphMark(paragraph); }
        finally { Marshal.ReleaseComObject(paragraph); }
    }

    internal static Range ResolveMainStoryParagraphAtPosition(Document document, int position)
    {
        var content = document.Content;
        try { return ResolveParagraphAtPosition(content, position); }
        finally { Marshal.ReleaseComObject(content); }
    }

    internal static Range ResolveParagraphAtPosition(Range storyReference, int position)
    {
        Range? content = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? owner = null;
        try
        {
            content = storyReference.Duplicate;
            content.Expand(WdUnits.wdStory);
            if (position < content.Start || position >= content.End)
                throw new InvalidDataException("Body format insertion position is outside the document.");
            // A collapsed boundary can report the previous paragraph in Word.
            // The forward character identifies the paragraph that owns insertion.
            probe = content.Duplicate;
            probe.SetRange(position, position + 1);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException("The insertion character has no unique paragraph owner.");
            paragraph = paragraphs[1];
            owner = paragraph.Range.Duplicate;
            if (owner.StoryType != content.StoryType || position < owner.Start || position >= owner.End)
                throw new InvalidDataException("Word returned a different paragraph for the insertion character.");
            var result = owner;
            owner = null;
            return result;
        }
        finally
        {
            if (owner is not null) Marshal.ReleaseComObject(owner);
            if (paragraph is not null) Marshal.ReleaseComObject(paragraph);
            if (paragraphs is not null) Marshal.ReleaseComObject(paragraphs);
            if (probe is not null) Marshal.ReleaseComObject(probe);
            if (content is not null) Marshal.ReleaseComObject(content);
        }
    }

    internal static WordCharacterFormatting CaptureParagraphMark(Range range)
    {
        var mark = GetParagraphMark(range);
        try { return Capture(mark); }
        finally { Marshal.ReleaseComObject(mark); }
    }

    internal void ApplyToParagraphMark(Range range)
    {
        var mark = GetParagraphMark(range);
        try { Apply(mark); }
        finally { Marshal.ReleaseComObject(mark); }
    }

    private static Range GetParagraphMark(Range range)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? mark = null;
        try
        {
            paragraphs = range.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException("Body character formatting requires one identified paragraph.");
            paragraph = paragraphs[1];
            mark = paragraph.Range;
            if (mark.End <= mark.Start)
                throw new InvalidDataException("The body paragraph has no terminal mark.");
            mark.SetRange(mark.End - 1, mark.End);
            var result = mark;
            mark = null;
            return result;
        }
        finally
        {
            if (mark is not null) Marshal.ReleaseComObject(mark);
            if (paragraph is not null) Marshal.ReleaseComObject(paragraph);
            if (paragraphs is not null) Marshal.ReleaseComObject(paragraphs);
        }
    }

    internal static WordCharacterFormatting Capture(Range range)
    {
        var result = new WordCharacterFormatting();
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            void Add<T>(T value, Action<Microsoft.Office.Interop.Word.Font, T> set)
            {
                if (value is null) return;
                if (value is string name ? string.IsNullOrWhiteSpace(name)
                    : Convert.ToDouble(value) == (int)WdConstants.wdUndefined) return;
                result.setters.Add(target => set(target, value));
            }
            Add(font.Name, (target, value) => target.Name = value);
            Add(font.NameAscii, (target, value) => target.NameAscii = value);
            Add(font.NameOther, (target, value) => target.NameOther = value);
            Add(font.NameFarEast, (target, value) => target.NameFarEast = value);
            Add(font.NameBi, (target, value) => target.NameBi = value);
            Add(font.Size, (target, value) => target.Size = value);
            Add(font.SizeBi, (target, value) => target.SizeBi = value);
            Add(font.Bold, (target, value) => target.Bold = value);
            Add(font.BoldBi, (target, value) => target.BoldBi = value);
            Add(font.Italic, (target, value) => target.Italic = value);
            Add(font.ItalicBi, (target, value) => target.ItalicBi = value);
            Add(font.Underline, (target, value) => target.Underline = value);
            Add(font.Position, (target, value) => target.Position = value);
            Add(font.Spacing, (target, value) => target.Spacing = value);
            Add(font.Scaling, (target, value) => target.Scaling = value);
            Add(font.Kerning, (target, value) => target.Kerning = value);
            Add(font.Subscript, (target, value) => target.Subscript = value);
            Add(font.Superscript, (target, value) => target.Superscript = value);
            Add(font.StrikeThrough, (target, value) => target.StrikeThrough = value);
            Add(font.DoubleStrikeThrough, (target, value) => target.DoubleStrikeThrough = value);
            Add(font.AllCaps, (target, value) => target.AllCaps = value);
            Add(font.SmallCaps, (target, value) => target.SmallCaps = value);
            var color = font.Color;
            if (color == WdColor.wdColorAutomatic || (int)color >= 0)
                result.Color = color;
            return result;
        }
        finally { if (font is not null) Marshal.ReleaseComObject(font); }
    }

    internal void Apply(Range range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            Apply(font);
        }
        finally { if (font is not null) Marshal.ReleaseComObject(font); }
    }

    internal void Apply(Microsoft.Office.Interop.Word.Font font)
    {
        // Use the PIA property setters so failures propagate to the owning
        // edit transaction. A partial format application is not success.
        foreach (var set in setters) set(font);
        if (Color.HasValue) font.Color = Color.Value;
        font.Hidden = 0;
    }
}
