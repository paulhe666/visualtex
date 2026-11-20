using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordSourceFontSize;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    /// <summary>
    /// Maps Range.Text indices of a scope to Word story positions one paragraph at
    /// a time. Used when the scope's Characters count disagrees with its text: WPS
    /// counts some native equations (e.g. an n-ary with \underset) one character
    /// short, and that one object must not make every LaTeX formula in a whole
    /// document unlocatable. A paragraph whose own counts disagree stays
    /// unmappable; only a formula inside it fails.
    /// </summary>
    private sealed class ParagraphCoordinateMap
    {
        private sealed class Segment
        {
            internal int TextStart;
            internal int TextEnd;
            internal int StoryStart;
            internal int[] Offsets = Array.Empty<int>();
            internal bool Direct;
            internal int[]? CharacterStarts;
            internal int[]? CharacterEnds;
        }

        private readonly List<Segment> _segments = new();

        internal static ParagraphCoordinateMap? Build(Document document, Range scope, string sourceText)
        {
            var map = new ParagraphCoordinateMap();
            var textOffset = 0;
            Paragraphs? paragraphs = null;
            try
            {
                paragraphs = scope.Paragraphs;
                for (var index = 1; index <= paragraphs.Count; index++)
                {
                    Paragraph? paragraph = null;
                    Range? paragraphRange = null;
                    Range? segmentRange = null;
                    Characters? characters = null;
                    try
                    {
                        paragraph = paragraphs[index];
                        paragraphRange = paragraph.Range;
                        var start = Math.Max(paragraphRange.Start, scope.Start);
                        var end = Math.Min(paragraphRange.End, scope.End);
                        if (end <= start) continue;
                        segmentRange = document.Range(start, end);
                        var text = segmentRange.Text ?? string.Empty;
                        if (textOffset + text.Length > sourceText.Length
                            || string.CompareOrdinal(sourceText, textOffset, text, 0, text.Length) != 0)
                            return null;
                        var offsets = BuildWordStoryOffsetIndex(text);
                        var logical = offsets[text.Length];
                        var segment = new Segment
                        {
                            TextStart = textOffset,
                            TextEnd = textOffset + text.Length,
                            StoryStart = start,
                            Offsets = offsets,
                            Direct = end - start == logical,
                        };
                        if (!segment.Direct)
                        {
                            characters = segmentRange.Characters;
                            if (characters.Count == logical)
                            {
                                segment.CharacterStarts = new int[logical];
                                segment.CharacterEnds = new int[logical];
                                for (var position = 0; position < logical; position++)
                                {
                                    Range? character = null;
                                    try
                                    {
                                        character = characters[position + 1];
                                        segment.CharacterStarts[position] = character.Start;
                                        segment.CharacterEnds[position] = character.End;
                                    }
                                    finally { Release(character); }
                                }
                            }
                        }
                        map._segments.Add(segment);
                        textOffset += text.Length;
                    }
                    finally
                    {
                        Release(characters);
                        Release(segmentRange);
                        Release(paragraphRange);
                        Release(paragraph);
                    }
                }
            }
            finally { Release(paragraphs); }
            return textOffset == sourceText.Length ? map : null;
        }

        /// <summary>Story position of a text boundary, or -1 when its paragraph is unmappable.</summary>
        internal int Map(int textIndex, bool isEnd)
        {
            foreach (var segment in _segments)
            {
                var inside = isEnd
                    ? textIndex > segment.TextStart && textIndex <= segment.TextEnd
                    : textIndex >= segment.TextStart && textIndex < segment.TextEnd;
                if (!inside) continue;
                var local = segment.Offsets[textIndex - segment.TextStart];
                if (segment.Direct) return segment.StoryStart + local;
                if (segment.CharacterStarts is null || segment.CharacterEnds is null) return -1;
                if (isEnd)
                    return local == 0 ? segment.StoryStart : segment.CharacterEnds[local - 1];
                return local < segment.CharacterStarts.Length ? segment.CharacterStarts[local] : -1;
            }
            return -1;
        }
    }
}
