using System;
using System.Xml.Linq;

namespace VisualTeX.WordVsto;

internal static class MathMlAccentGeometry
{
    // MathML operator dictionary defaults for the accent vocabulary supported
    // by the existing source/native adapters. Explicit stretchy always wins.
    // https://www.w3.org/TR/MathML3/appendixc.html
    internal static bool IsStretchy(XElement mark)
    {
        var declared = ((string?)mark.Attribute("stretchy"))?.Trim();
        if (string.Equals(declared, "false", StringComparison.OrdinalIgnoreCase) || declared == "0") return false;
        if (string.Equals(declared, "true", StringComparison.OrdinalIgnoreCase) || declared == "1") return true;
        return mark.Value.Trim() is "^" or "ˆ" or "\u0302" or "~" or "˜" or "\u0303"
            or "→" or "←" or "↔" or "\u20D7" or "\u20D6" or "\u20E1"
            or "¯" or "‾" or "―" or "ˉ" or "⌢" or "⏜";
    }
}
