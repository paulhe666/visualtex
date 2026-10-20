using System.Globalization;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.PowerPointVsto;

// Pure planning rules for PowerPoint batch format conversion. Nothing in this
// file touches COM, so it is shared with the unit tests.

internal enum PowerPointFormulaKind
{
    VisualTeXOle,
    VisualTeXPicture,
    MathType,
    Omml,
}

internal enum PowerPointConversionScope
{
    Selection,
    Slide,
    Presentation,
}

internal enum PowerPointConversionSkipReason
{
    // The equation shares its text box with ordinary text; an object cannot be
    // placed inline in PowerPoint text, so it stays native.
    OmmlMixedWithText,
    // An equation-only text box with several equations would be split into
    // several objects, which an animation on the box cannot follow.
    OmmlSplitWithAnimation,
    InGroup,
    InTable,
    RotatedOrFlipped,
    Linked,
    Unreadable,
    Changed,
    Failed,
}

internal sealed class PowerPointConversionSkip
{
    internal PowerPointConversionSkip(PowerPointConversionSkipReason reason, int slideIndex, int count = 1, string? detail = null)
    {
        Reason = reason;
        SlideIndex = slideIndex;
        Count = Math.Max(1, count);
        Detail = detail;
    }

    internal PowerPointConversionSkipReason Reason { get; }
    internal int SlideIndex { get; }
    internal int Count { get; }
    internal string? Detail { get; }
}

internal sealed class PowerPointConversionRequest
{
    internal const string VisualTeXFamily = "visualtex";
    internal const string AnySource = "any";

    private PowerPointConversionRequest(string sourceFamily, string targetFamily, string targetMode, PowerPointConversionScope scope)
    {
        SourceFamily = sourceFamily;
        TargetFamily = targetFamily;
        TargetMode = targetMode;
        Scope = scope;
    }

    // visualtex | mathTypeOle | wordOmml | any
    internal string SourceFamily { get; }
    // visualtex | mathTypeOle | wordOmml
    internal string TargetFamily { get; }
    // nativeOle | crossPlatformPicture | mathTypeOle | wordOmml
    internal string TargetMode { get; }
    internal PowerPointConversionScope Scope { get; }
    internal bool Unify => SourceFamily == AnySource;

    internal bool Accepts(PowerPointFormulaKind kind)
    {
        var family = PowerPointConversionRules.FamilyOf(kind);
        return !string.Equals(family, TargetFamily, StringComparison.Ordinal)
            && (Unify || string.Equals(family, SourceFamily, StringComparison.Ordinal));
    }

    // Ribbon tag: "<source>|<target>|<scope>". A VisualTeX target resolves to the
    // user's chosen VisualTeX form (OLE object or SVG picture) at click time.
    internal static bool TryParse(string? tag, string visualTeXTargetMode, out PowerPointConversionRequest request)
    {
        request = null!;
        if (string.IsNullOrWhiteSpace(tag)) return false;
        var parts = tag!.Split('|');
        if (parts.Length != 3) return false;
        var source = parts[0];
        var target = parts[1];
        if (!IsFamily(source) && source != AnySource) return false;
        if (!IsFamily(target) || source == target) return false;
        PowerPointConversionScope scope;
        switch (parts[2])
        {
            case "selection": scope = PowerPointConversionScope.Selection; break;
            case "slide": scope = PowerPointConversionScope.Slide; break;
            case "presentation": scope = PowerPointConversionScope.Presentation; break;
            default: return false;
        }
        var targetMode = target == VisualTeXFamily
            ? (visualTeXTargetMode == FormulaOleContract.CrossPlatformPictureMode
                ? FormulaOleContract.CrossPlatformPictureMode
                : FormulaOleContract.NativeOleMode)
            : target;
        request = new PowerPointConversionRequest(source, target, targetMode, scope);
        return true;
    }

    private static bool IsFamily(string value) =>
        value == VisualTeXFamily
        || value == FormulaOleContract.MathTypeOleMode
        || value == FormulaOleContract.WordOmmlMode;
}

internal static class PowerPointConversionRules
{
    internal static string FamilyOf(PowerPointFormulaKind kind) => kind switch
    {
        PowerPointFormulaKind.VisualTeXOle => PowerPointConversionRequest.VisualTeXFamily,
        PowerPointFormulaKind.VisualTeXPicture => PowerPointConversionRequest.VisualTeXFamily,
        PowerPointFormulaKind.MathType => FormulaOleContract.MathTypeOleMode,
        _ => FormulaOleContract.WordOmmlMode,
    };

    // PowerPoint cannot rotate or mirror an OLE object; text boxes rotate but
    // must keep their text readable; only pictures accept both.
    internal static bool TargetSupportsRotation(string targetMode) =>
        targetMode == FormulaOleContract.CrossPlatformPictureMode
        || targetMode == FormulaOleContract.WordOmmlMode;

    internal static bool TargetSupportsFlip(string targetMode) =>
        targetMode == FormulaOleContract.CrossPlatformPictureMode;

    // MathType draws an equation at its own size; PowerPoint then scales the
    // object. The size the audience sees is that MTEF size times the scale.
    internal static double ScaledFontSize(
        double nativeFontSizePt,
        double shapeWidth,
        double shapeHeight,
        double naturalWidth,
        double naturalHeight)
    {
        static bool Valid(double current, double natural) =>
            current > 0 && natural > 0 && !double.IsNaN(current) && !double.IsInfinity(current)
            && !double.IsNaN(natural) && !double.IsInfinity(natural);
        double scale;
        if (Valid(shapeWidth, naturalWidth) && Valid(shapeHeight, naturalHeight))
            scale = Math.Sqrt(shapeWidth / naturalWidth * (shapeHeight / naturalHeight));
        else if (Valid(shapeHeight, naturalHeight))
            scale = shapeHeight / naturalHeight;
        else if (Valid(shapeWidth, naturalWidth))
            scale = shapeWidth / naturalWidth;
        else
            scale = 1;
        return nativeFontSizePt * scale;
    }

    // Renderers accept 5-200 pt. Render inside that range and scale the placed
    // object by the remainder so the visible size still matches the source.
    internal static (float RenderFontSizePt, float PlacementScale) ResolveRenderFontSize(double effectiveFontSizePt)
    {
        if (!(effectiveFontSizePt > 0) || double.IsInfinity(effectiveFontSizePt))
            return (FormulaFontSize.DefaultPt, 1f);
        var render = FormulaFontSize.Normalize(effectiveFontSizePt);
        return (render, (float)(effectiveFontSizePt / render));
    }

    internal static (float Left, float Top) CenteredAt(float centerX, float centerY, float width, float height) =>
        (centerX - width / 2f, centerY - height / 2f);

    internal static string RenderKey(string targetMode, string displayMode, float fontSizePt, string? letterFont, string? chineseFont, string latex) =>
        string.Join(
            "\u001F",
            targetMode,
            displayMode,
            fontSizePt.ToString("0.##", CultureInfo.InvariantCulture),
            letterFont ?? string.Empty,
            chineseFont ?? string.Empty,
            latex);

    internal static string FamilyLabel(string familyOrMode) => familyOrMode switch
    {
        FormulaOleContract.MathTypeOleMode => "MathType",
        FormulaOleContract.WordOmmlMode => "OMML",
        PowerPointConversionRequest.AnySource => "",
        _ => "VisualTeX",
    };

    internal static string TargetLabel(string targetMode, bool english) => targetMode switch
    {
        FormulaOleContract.MathTypeOleMode => "MathType",
        FormulaOleContract.WordOmmlMode => english ? "OMML (native PowerPoint equations)" : "OMML（PowerPoint 原生公式）",
        FormulaOleContract.CrossPlatformPictureMode => english ? "VisualTeX SVG pictures" : "VisualTeX SVG 图片",
        _ => english ? "VisualTeX OLE objects" : "VisualTeX OLE 对象",
    };

    internal static string SkipReasonText(PowerPointConversionSkipReason reason, int count, bool english) => reason switch
    {
        PowerPointConversionSkipReason.OmmlMixedWithText => english
            ? $"{count} OMML equation(s) share a text box with other text and stay native (an object cannot sit inside a line of text)"
            : $"{count} 个 OMML 公式与文字在同一文本框中，保持原样（对象无法嵌入文字行内）",
        PowerPointConversionSkipReason.OmmlSplitWithAnimation => english
            ? $"{count} OMML equation(s) are in an animated text box that holds several equations; splitting it would lose the animation"
            : $"{count} 个 OMML 公式所在文本框含多个公式且带动画，拆分会丢失动画",
        PowerPointConversionSkipReason.InGroup => english
            ? $"{count} equation(s) are inside groups (ungroup them first)"
            : $"{count} 个公式位于组合中（请先取消组合）",
        PowerPointConversionSkipReason.InTable => english
            ? $"{count} equation(s) are inside tables"
            : $"{count} 个公式位于表格中",
        PowerPointConversionSkipReason.RotatedOrFlipped => english
            ? $"{count} equation(s) are rotated or flipped, which the target format cannot keep"
            : $"{count} 个公式已旋转或翻转，目标格式无法保持",
        PowerPointConversionSkipReason.Linked => english
            ? $"{count} MathType object(s) are linked files, not embedded equations"
            : $"{count} 个 MathType 对象是链接文件，而非嵌入公式",
        PowerPointConversionSkipReason.Unreadable => english
            ? $"{count} equation(s) could not be read"
            : $"{count} 个公式无法读取",
        PowerPointConversionSkipReason.Changed => english
            ? $"{count} equation(s) changed or disappeared during conversion"
            : $"{count} 个公式在转换过程中被修改或删除",
        _ => english
            ? $"{count} equation(s) failed to convert"
            : $"{count} 个公式转换失败",
    };

    // One line per reason, most common first; the first failure detail is kept
    // because it usually explains the rest.
    internal static string DescribeSkips(IReadOnlyCollection<PowerPointConversionSkip> skips, bool english)
    {
        if (skips.Count == 0) return string.Empty;
        var lines = skips
            .GroupBy(skip => skip.Reason)
            .Select(group => (Reason: group.Key, Count: group.Sum(skip => skip.Count),
                Detail: group.Select(skip => skip.Detail).FirstOrDefault(detail => !string.IsNullOrWhiteSpace(detail))))
            .OrderByDescending(entry => entry.Count)
            .Select(entry =>
            {
                var text = "• " + SkipReasonText(entry.Reason, entry.Count, english);
                if (!string.IsNullOrWhiteSpace(entry.Detail))
                    text += (english ? ": " : "：") + Truncate(entry.Detail!, 300);
                return text;
            });
        return string.Join("\r\n", lines);
    }

    internal static int CountSkips(IEnumerable<PowerPointConversionSkip> skips) => skips.Sum(skip => skip.Count);

    private static string Truncate(string value, int maximum)
    {
        var single = value.Replace("\r", " ").Replace("\n", " ").Trim();
        return single.Length <= maximum ? single : single.Substring(0, maximum - 1) + "…";
    }
}

// The colors the audience sees must survive a conversion. Explicit colors travel
// inside MathML (mathcolor). What a format shows for the uncolored parts differs:
// a native equation inherits the slide's text color (white on a dark theme),
// while MathType and VisualTeX draw them black. The source's displayed color for
// its uncolored parts is therefore written explicitly, but only where the target
// would otherwise show a different color.
internal static class PowerPointConversionColors
{
    internal const string Black = "#000000";

    // Office reports RGB as 0x00BBGGRR.
    internal static string FromOfficeRgb(int rgb) =>
        "#" + (rgb & 0xFF).ToString("X2", CultureInfo.InvariantCulture)
        + ((rgb >> 8) & 0xFF).ToString("X2", CultureInfo.InvariantCulture)
        + ((rgb >> 16) & 0xFF).ToString("X2", CultureInfo.InvariantCulture);

    // The plain RGB run colors of a native equation: exactly the colors its
    // MathML carries as mathcolor.
    internal static HashSet<string> ExplicitOmmlColors(string omml)
    {
        var colors = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        XElement root;
        try { root = XElement.Parse(omml); }
        catch { return colors; }
        foreach (var rgb in root.Descendants(PowerPointSlideXml.A + "srgbClr"))
        {
            if (rgb.Parent?.Name != PowerPointSlideXml.A + "solidFill") continue;
            var value = (string?)rgb.Attribute("val");
            if (value?.Length == 6 && int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _))
                colors.Add("#" + value.ToUpperInvariant());
        }
        return colors;
    }

    // The displayed color of a native equation's uncolored runs: the most used
    // run color that is not one of its explicit colors. Null when every run is
    // explicitly colored or no run color could be read.
    internal static string? InheritedColor(IEnumerable<(string Color, int Length)> runs, ISet<string> explicitColors) =>
        runs.Where(run => run.Length > 0 && !string.IsNullOrEmpty(run.Color) && !explicitColors.Contains(run.Color))
            .GroupBy(run => run.Color.ToUpperInvariant(), StringComparer.Ordinal)
            .OrderByDescending(group => group.Sum(run => run.Length))
            .Select(group => group.Key)
            .FirstOrDefault();

    // True when a target that shows uncolored parts in targetDefault would not
    // show the source's color for them.
    internal static bool NeedsExplicit(string? sourceColor, string? targetDefault) =>
        !string.IsNullOrEmpty(sourceColor)
        && !string.IsNullOrEmpty(targetDefault)
        && !string.Equals(sourceColor, targetDefault, StringComparison.OrdinalIgnoreCase);

    // Colors the uncolored parts of an equation; explicitly colored parts keep
    // their own color because an inner mathcolor wins over the root's.
    internal static string ApplyBaseColor(string mathMl, string? color)
    {
        if (string.IsNullOrWhiteSpace(color) || string.IsNullOrWhiteSpace(mathMl)) return mathMl;
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var root = document.Root;
        if (root is null || root.Attributes().Any(attribute => attribute.Name.LocalName is "mathcolor" or "color"))
            return mathMl;
        root.SetAttributeValue("mathcolor", color!.ToUpperInvariant());
        return document.Root!.ToString(SaveOptions.DisableFormatting);
    }
}

// One native equation inside a PowerPoint text shape.
internal sealed class PowerPointOmmlEquation
{
    internal PowerPointOmmlEquation(string omml, bool display, int paragraph)
    {
        Omml = omml;
        Display = display;
        Paragraph = paragraph;
    }

    internal string Omml { get; }
    internal bool Display { get; }
    internal int Paragraph { get; }
}

internal sealed class PowerPointOmmlShapeAnalysis
{
    internal PowerPointOmmlShapeAnalysis(IReadOnlyList<PowerPointOmmlEquation> equations, bool hasOtherContent)
    {
        Equations = equations;
        HasOtherContent = hasOtherContent;
    }

    internal IReadOnlyList<PowerPointOmmlEquation> Equations { get; }
    // Visible text or fields outside the equations.
    internal bool HasOtherContent { get; }
    internal bool EquationOnly => Equations.Count > 0 && !HasOtherContent;
}

internal enum PowerPointSlideShapeKind
{
    Other,
    Ole,
    Picture,
    TextWithMath,
    Group,
    TableWithMath,
}

internal sealed class PowerPointSlideShapeInfo
{
    internal int ShapeId { get; set; }
    internal PowerPointSlideShapeKind Kind { get; set; }
    internal string? ProgId { get; set; }
    internal bool Linked { get; set; }
    internal PowerPointOmmlShapeAnalysis? Math { get; set; }
    // Formulas the XML can see inside a group or table (exact kinds are
    // confirmed through PowerPoint for groups).
    internal int NestedFormulaHints { get; set; }
    internal XElement Element { get; set; } = null!;
}

// Reads a saved slide part. PowerPoint writes a text shape that contains math
// twice under mc:AlternateContent (the a14 choice and a picture fallback) with
// the same shape id; the choice branch is the live shape.
internal static class PowerPointSlideXml
{
    internal static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    internal static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    internal static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    internal static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    internal static Dictionary<int, PowerPointSlideShapeInfo> IndexTopLevelShapes(XDocument slide)
    {
        var result = new Dictionary<int, PowerPointSlideShapeInfo>();
        var tree = slide.Root?.Element(P + "cSld")?.Element(P + "spTree");
        if (tree is null) return result;
        foreach (var child in tree.Elements())
        {
            var shape = child.Name == Mc + "AlternateContent" ? LiveBranch(child) : child;
            if (shape is null) continue;
            var info = Describe(shape);
            if (info is null || result.ContainsKey(info.ShapeId)) continue;
            result.Add(info.ShapeId, info);
        }
        return result;
    }

    internal static XElement? LiveBranch(XElement alternateContent) =>
        alternateContent.Elements(Mc + "Choice").Select(choice => choice.Elements().FirstOrDefault()).FirstOrDefault(element => element is not null)
        ?? alternateContent.Element(Mc + "Fallback")?.Elements().FirstOrDefault();

    internal static PowerPointOmmlShapeAnalysis AnalyzeText(XElement shape)
    {
        var equations = new List<PowerPointOmmlEquation>();
        var otherContent = false;
        var body = shape.Element(P + "txBody");
        var paragraphs = body?.Elements(A + "p").ToArray() ?? Array.Empty<XElement>();
        for (var index = 0; index < paragraphs.Length; index++)
        {
            foreach (var math in LiveDescendants(paragraphs[index]).Where(element => element.Name == M + "oMath"))
            {
                if (math.Ancestors(M + "oMath").Any()) continue;
                equations.Add(new PowerPointOmmlEquation(
                    math.ToString(SaveOptions.DisableFormatting),
                    math.Parent?.Name == M + "oMathPara",
                    index));
            }
            foreach (var element in LiveDescendants(paragraphs[index]))
            {
                if (element.Name == A + "fld") otherContent = true;
                if (element.Name == A + "t"
                    && !element.Ancestors(M + "oMath").Any()
                    && !string.IsNullOrWhiteSpace(element.Value))
                    otherContent = true;
            }
        }
        return new PowerPointOmmlShapeAnalysis(equations, otherContent);
    }

    private static PowerPointSlideShapeInfo? Describe(XElement shape)
    {
        var id = ShapeId(shape);
        if (id is null) return null;
        var info = new PowerPointSlideShapeInfo { ShapeId = id.Value, Element = shape, Kind = PowerPointSlideShapeKind.Other };
        if (shape.Name == P + "sp")
        {
            if (LiveDescendants(shape).Any(element => element.Name == M + "oMath"))
            {
                info.Kind = PowerPointSlideShapeKind.TextWithMath;
                info.Math = AnalyzeText(shape);
            }
        }
        else if (shape.Name == P + "pic")
        {
            info.Kind = PowerPointSlideShapeKind.Picture;
        }
        else if (shape.Name == P + "graphicFrame")
        {
            var ole = LiveDescendants(shape).FirstOrDefault(element => element.Name == P + "oleObj");
            if (ole is not null)
            {
                info.Kind = PowerPointSlideShapeKind.Ole;
                info.ProgId = (string?)ole.Attribute("progId");
                info.Linked = ole.Element(P + "link") is not null;
            }
            else if (LiveDescendants(shape).Any(element => element.Name == A + "tbl"))
            {
                var math = LiveDescendants(shape).Count(element => element.Name == M + "oMath"
                    && !element.Ancestors(M + "oMath").Any());
                if (math > 0)
                {
                    info.Kind = PowerPointSlideShapeKind.TableWithMath;
                    info.NestedFormulaHints = math;
                }
            }
        }
        else if (shape.Name == P + "grpSp")
        {
            info.Kind = PowerPointSlideShapeKind.Group;
            info.NestedFormulaHints = CountGroupFormulaHints(shape);
        }
        return info;
    }

    private static int CountGroupFormulaHints(XElement group)
    {
        var count = 0;
        foreach (var element in LiveDescendants(group))
        {
            if (element.Name == P + "oleObj") count++;
            else if (element.Name == P + "pic") count++;
            else if (element.Name == P + "sp")
                count += LiveDescendants(element).Count(child => child.Name == M + "oMath" && !child.Ancestors(M + "oMath").Any());
        }
        return count;
    }

    // Descendants that belong to the live representation: whenever an
    // mc:AlternateContent is met, only its first choice is followed.
    internal static IEnumerable<XElement> LiveDescendants(XElement root)
    {
        var stack = new Stack<XElement>();
        for (var child = root.LastNode; child is not null; child = child.PreviousNode)
            if (child is XElement element) stack.Push(element);
        while (stack.Count > 0)
        {
            var element = stack.Pop();
            if (element.Name == Mc + "AlternateContent")
            {
                var live = element.Elements(Mc + "Choice").FirstOrDefault() ?? element.Element(Mc + "Fallback");
                if (live is null) continue;
                for (var child = live.LastNode; child is not null; child = child.PreviousNode)
                    if (child is XElement nested) stack.Push(nested);
                continue;
            }
            yield return element;
            for (var child = element.LastNode; child is not null; child = child.PreviousNode)
                if (child is XElement nested) stack.Push(nested);
        }
    }

    private static int? ShapeId(XElement shape)
    {
        var properties = shape.Elements().FirstOrDefault(element => element.Name.LocalName.StartsWith("nv", StringComparison.Ordinal)
            && element.Name.Namespace == P);
        var value = (string?)properties?.Element(P + "cNvPr")?.Attribute("id");
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) && id > 0 ? id : null;
    }
}
