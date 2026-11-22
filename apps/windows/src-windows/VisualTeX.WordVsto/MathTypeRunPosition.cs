using System;

namespace VisualTeX.WordVsto;

/// <summary>
/// The Word run position (the vertical offset Word applies to an inline object, in points, negative = lowered) of
/// a rewritten MathType object: the rewritten equation's own position, as MathPage gives it, in the document's
/// drawing scale, plus whatever the document had added to the source equation's own position (a user's nudge).
/// A new size or a deeper equation moves the object with its baseline; keeping the old position put a resized
/// display formula above or below the line of its tabs and number and an edited inline formula off the baseline.
/// </summary>
internal static class MathTypeRunPosition
{
    internal const int WordLimitPt = 1584;

    internal static int Rewritten(
        int originalWordPosition,
        int sourceNativeWordPosition,
        int rewrittenNativeWordPosition,
        float heightScale)
    {
        var scale = float.IsNaN(heightScale) || float.IsInfinity(heightScale) || heightScale <= 0f ? 1f : heightScale;
        var sourceOwn = (int)Math.Round(sourceNativeWordPosition * scale, MidpointRounding.AwayFromZero);
        var rewrittenOwn = (int)Math.Round(rewrittenNativeWordPosition * scale, MidpointRounding.AwayFromZero);
        var position = rewrittenOwn + (originalWordPosition - sourceOwn);
        return Math.Max(-WordLimitPt, Math.Min(WordLimitPt, position));
    }
}
