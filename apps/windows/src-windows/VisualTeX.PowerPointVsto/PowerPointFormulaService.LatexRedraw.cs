using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using ColorFormat = Microsoft.Office.Interop.PowerPoint.ColorFormat;
using FillFormat = Microsoft.Office.Interop.PowerPoint.FillFormat;
using LineFormat = Microsoft.Office.Interop.PowerPoint.LineFormat;
using PptFont = Microsoft.Office.Interop.PowerPoint.Font;
using PptTextFrame = Microsoft.Office.Interop.PowerPoint.TextFrame;
using PptTextRange = Microsoft.Office.Interop.PowerPoint.TextRange;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace VisualTeX.PowerPointVsto;

// LaTeX redraw and "formulas to LaTeX code", as part of the batch format
// conversion (same gate, progress, snapshot capture, source proof, undo entries
// and report). LaTeX code in a text box becomes native equations where it
// stands; a box that holds just one formula can instead become a VisualTeX or
// MathType object. Formulas become LaTeX code again: a native equation in text
// where it stands, an object as a text box in its place.
public sealed partial class PowerPointFormulaService
{
    private const string MathFontName = "Cambria Math";

    // Commands whose source or result is text: a text selection limits them.
    private static bool ConvertsText(PowerPointConversionRequest request) =>
        request.SourceFamily == PowerPointConversionRequest.LatexFamily
        || request.TargetFamily == PowerPointConversionRequest.LatexFamily;

    // The selected text (length 0: the caret) of a one-shape text selection. A
    // selection that cannot be read stops the command: converting the whole box
    // instead could change far more than the user selected.
    private static (int ShapeId, int Start, int Length) ReadSelectedTextRange(Selection selection, int shapeId)
    {
        object? text = null;
        Exception? failure = null;
        try
        {
            text = SelectedText(selection);
            var start = Convert.ToInt32(((dynamic)text).Start);
            var length = Convert.ToInt32(((dynamic)text).Length);
            if (start >= 1 && length >= 0) return (shapeId, start, length);
        }
        catch (Exception error) { failure = error; }
        finally { Release(text); }
        throw new InvalidOperationException(Text(
            "无法读取文本框中所选文字的位置，未转换任何公式。请重新选择后再试。",
            "The position of the selected text in the text box could not be read, so nothing was converted. Select it again and retry."),
            failure);
    }

    private static (int Start, int Length)? SelectedTextIn(PowerPointFormulaConversionPlan plan, int shapeId) =>
        plan.SelectedText is { } selected && selected.ShapeId == shapeId ? (selected.Start, selected.Length) : null;

    // ---- Capture ------------------------------------------------------------

    private void CaptureLatexShape(
        PowerPointFormulaConversionPlan plan,
        XDocument slideXml,
        Slide slide,
        Shape shape,
        int slideIndex,
        PowerPointSlideShapeInfo info)
    {
        var element = info.Element;
        if (info.Kind == PowerPointSlideShapeKind.Group || element.Name == PowerPointSlideXml.P + "graphicFrame")
        {
            // Text inside a group or a table is not redrawn; what is there is reported.
            int found;
            try { found = PowerPointLatexText.CountFormulasInTextBodies(element); }
            catch { found = 0; }
            if (found > 0)
                plan.Skips.Add(new PowerPointConversionSkip(
                    info.Kind == PowerPointSlideShapeKind.Group
                        ? PowerPointConversionSkipReason.InGroup
                        : PowerPointConversionSkipReason.InTable,
                    slideIndex,
                    found));
            return;
        }
        if (element.Name != PowerPointSlideXml.P + "sp" || element.Element(PowerPointSlideXml.P + "txBody") is null) return;
        CaptureLatexUnit(plan, slideXml, slide, shape, slideIndex, info);
    }

    private void CaptureLatexUnit(
        PowerPointFormulaConversionPlan plan,
        XDocument slideXml,
        Slide slide,
        Shape shape,
        int slideIndex,
        PowerPointSlideShapeInfo info)
    {
        string text;
        try { text = ReadConversionText(shape); }
        catch (Exception error)
        {
            plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.Unreadable, slideIndex, 1,
                Text("无法读取文本框的文字：", "The text box's text could not be read: ") + Describe(error)));
            return;
        }
        if (!PowerPointLatexText.MayContainLatex(text)) return;
        try
        {
            var saved = PowerPointNativeOmmlPackage.SelectTextShape(slideXml, shape.Id);
            // Native equations in the box are masked: what they spell is never read as LaTeX.
            var equations = info.Kind == PowerPointSlideShapeKind.TextWithMath
                ? ReadEquationRanges(shape, saved)
                : new List<(int Start, int Length)>();
            var spans = PowerPointLatexText.FindFormulas(text, equations);
            if (spans.Count == 0) return;
            // Selected text redraws the formulas it holds; a caret, the whole box.
            var selected = SelectedTextIn(plan, shape.Id) is { Length: > 0 } range
                ? range
                : ((int Start, int Length)?)null;
            var chosen = spans
                .Where(span => selected is null || PowerPointLatexText.Inside(span, selected.Value.Start, selected.Value.Length))
                .ToList();
            if (chosen.Count == 0) return;

            var unit = NewConversionUnit(slide, shape, slideIndex, PowerPointFormulaKind.Latex);
            unit.TextFingerprint = PowerPointTextEditTransaction.Fingerprint(saved.ToString(SaveOptions.DisableFormatting));
            unit.SourceText = text;
            if (plan.Request.TargetMode == FormulaOleContract.WordOmmlMode)
            {
                // Each equation goes where its code is; the rest of the box stays.
                unit.InPlaceText = true;
                unit.EquationRanges.AddRange(equations);
                var ranges = PowerPointLatexText.ReplacedRanges(text, equations, spans);
                for (var index = 0; index < spans.Count; index++)
                {
                    var span = spans[index];
                    if (!chosen.Contains(span)) continue;
                    unit.Items.Add(new PowerPointConversionItem
                    {
                        Latex = span.Latex,
                        DisplayMode = span.DisplayMode,
                        // The writer reads the size again from the code's text.
                        EffectiveFontSizePt = ReadCharacterStyle(shape, span.Start).Size ?? 20f,
                        TextStart = ranges[index].Start,
                        TextLength = ranges[index].Length,
                    });
                }
                plan.Units.Add(unit);
                return;
            }

            // A VisualTeX or MathType object replaces the whole box: only a box
            // holding nothing but this one formula.
            if (!PowerPointLatexText.IsSingleFormula(text, equations, spans))
            {
                plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.LatexMixedWithText, slideIndex, chosen.Count));
                return;
            }
            // Only a text box: a drawn shape (a filled circle labelled $v_1$) is
            // never deleted for its label.
            if (!CanBecomeFormulaObject(shape))
            {
                plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.LatexInShape, slideIndex, chosen.Count));
                return;
            }
            var formula = chosen[0];
            // The object shows the formula as the box showed its text: never
            // mirrored, so the box's flips are not carried over.
            unit.Rotation = PowerPointLatexText.TextRotation(unit.Rotation, unit.FlipVertical);
            unit.FlipHorizontal = false;
            unit.FlipVertical = false;
            if (!TargetKeepsTransform(plan, unit))
            {
                plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.RotatedOrFlipped, slideIndex, 1,
                    latex: formula.Latex));
                return;
            }
            var (size, color) = ReadCharacterStyle(shape, formula.Start);
            var width = shape.Width;
            var height = shape.Height;
            unit.Items.Add(new PowerPointConversionItem
            {
                Latex = PowerPointLatexText.WithTextColor(formula.Latex, color),
                DisplayMode = formula.DisplayMode,
                // The size the audience sees: the code's size, shrunk as PowerPoint
                // shrinks the box's text when it does not fit.
                EffectiveFontSizePt = (size ?? 20f) * PowerPointLatexText.AutofitFontScale(saved),
                BaseColor = color,
                CenterX = shape.Left + width / 2f,
                CenterY = shape.Top + height / 2f,
                VisualWidth = width,
                VisualHeight = height,
                TextStart = formula.Start,
                TextLength = formula.Length,
            });
            plan.Units.Add(unit);
        }
        catch (Exception error)
        {
            plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.Unreadable, slideIndex, 1,
                Text("读取文本框中的 LaTeX 公式失败：", "Reading the LaTeX formulas of the text box failed: ") + Describe(error)));
        }
    }

    // Formulas to LaTeX code, native equations in text (equation-only boxes
    // included): each equation's characters become its code where they are.
    private void CaptureEquationsAsLatexUnit(
        PowerPointFormulaConversionPlan plan,
        XDocument slideXml,
        Slide slide,
        Shape shape,
        int slideIndex,
        PowerPointOmmlShapeAnalysis analysis)
    {
        var count = analysis.Equations.Count;
        if (count == 0) return;
        PowerPointConversionUnit unit;
        string text;
        List<(int Start, int Length)> zones;
        try
        {
            unit = NewConversionUnit(slide, shape, slideIndex, PowerPointFormulaKind.Omml);
            unit.InPlaceText = true;
            var saved = PowerPointNativeOmmlPackage.SelectTextShape(slideXml, shape.Id);
            unit.TextFingerprint = PowerPointTextEditTransaction.Fingerprint(saved.ToString(SaveOptions.DisableFormatting));
            text = ReadConversionText(shape);
            unit.SourceText = text;
            zones = ReadEquationRanges(shape, saved);
            if (zones.Count != count)
                throw new InvalidDataException(Text(
                    "文本框中的公式数量与已保存内容不一致。",
                    "The text box's equations do not match its saved content."));
            unit.EquationRanges.AddRange(zones);
        }
        catch (Exception error)
        {
            plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.Unreadable, slideIndex, count,
                Text("读取原生公式失败：", "Reading the native equation failed: ") + Describe(error)));
            return;
        }
        // Selected text converts the equations it touches. A caret converts the
        // equation it is in (as the Mac add-ins do), or the whole box when it is
        // in none.
        var selected = SelectedTextIn(plan, shape.Id);
        var caretEquation = selected is { Length: 0 } caret ? WpsMathText.ZoneAt(zones, caret.Start, 0) : -1;
        for (var ordinal = 0; ordinal < count; ordinal++)
        {
            var (start, length) = zones[ordinal];
            if (caretEquation >= 0 ? ordinal != caretEquation
                : selected is { Length: > 0 } range && !PowerPointLatexText.Overlaps(start, length, range.Start, range.Length))
                continue;
            try
            {
                var mathMl = PowerPointNativeOmmlKernel.ReadNativeMathMl(analysis.Equations[ordinal].Omml);
                var latex = MathMlToLatexConverter.ConvertWordEquation(mathMl).Trim();
                if (string.IsNullOrWhiteSpace(latex))
                    throw new InvalidDataException(Text("原生公式为空或无法识别。", "The native equation is empty or unsupported."));
                unit.Items.Add(new PowerPointConversionItem
                {
                    Latex = latex,
                    SourceMathMl = mathMl,
                    DisplayMode = PowerPointLatexText.IsAloneInParagraph(text, start, length) ? "block" : "inline",
                    TextStart = start,
                    TextLength = length,
                });
            }
            catch (Exception error)
            {
                // That equation stays native; the others are still converted.
                plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.Unreadable, slideIndex, 1,
                    Text("读取原生公式失败：", "Reading the native equation failed: ") + Describe(error)));
            }
        }
        if (unit.Items.Count > 0) plan.Units.Add(unit);
    }

    // The text a conversion reads and writes, in the coordinates of the native
    // equation writer: PowerPoint's TextFrame2 text; in WPS its TextFrame text,
    // which spells equations as WpsMathText expects.
    private static string ReadConversionText(Shape shape)
    {
        if (shape.HasTextFrame != MsoTriState.msoTrue) return string.Empty;
        if (OfficeHostEnvironment.IsWps) return WpsNativeEquations.LiveText(shape);
        object? frame = null, range = null;
        try
        {
            frame = shape.TextFrame2;
            range = ((dynamic)frame).TextRange;
            return Convert.ToString(((dynamic)range).Text) ?? string.Empty;
        }
        finally { Release(range); Release(frame); }
    }

    // Only a text box is replaced by a formula object: a text box, a placeholder,
    // or an autoshape that shows nothing but its text.
    private static bool CanBecomeFormulaObject(Shape shape)
    {
        int type;
        try { type = (int)shape.Type; }
        catch { return false; }
        var host = PowerPointLatexText.HostOf(type);
        if (host != PowerPointLatexText.TextHost.AutoShape)
            return PowerPointLatexText.CanBecomeObject(host, fillVisible: false, lineVisible: false);
        FillFormat? fill = null;
        LineFormat? line = null;
        try
        {
            fill = shape.Fill;
            line = shape.Line;
            return PowerPointLatexText.CanBecomeObject(host,
                fillVisible: fill.Visible != MsoTriState.msoFalse,
                lineVisible: line.Visible != MsoTriState.msoFalse);
        }
        catch { return false; }
        finally { Release(line); Release(fill); }
    }

    // A text's native equations (1-based start and length, text order): through
    // MathZones in PowerPoint, through the saved XML in WPS, which has none.
    private static List<(int Start, int Length)> ReadEquationRanges(Shape shape, XElement savedShape) =>
        OfficeHostEnvironment.IsWps
            ? WpsNativeEquations.Locate(savedShape, shape).ToList()
            : ReadMathZoneRanges(shape);

    private static List<(int Start, int Length)> ReadMathZoneRanges(Shape shape)
    {
        var result = new List<(int Start, int Length)>();
        object? frame = null, range = null, zones = null;
        try
        {
            frame = shape.TextFrame2;
            range = ((dynamic)frame).TextRange;
            try { zones = ((dynamic)range).MathZones(); }
            catch { zones = ((dynamic)range).MathZones(Type.Missing, Type.Missing); }
            var count = Convert.ToInt32(((dynamic)zones).Count);
            for (var index = 1; index <= count; index++)
            {
                object? zone = null;
                try
                {
                    zone = ((dynamic)range).MathZones(index, 1);
                    result.Add((Convert.ToInt32(((dynamic)zone).Start), Convert.ToInt32(((dynamic)zone).Length)));
                }
                finally { Release(zone); }
            }
            return result;
        }
        finally { Release(zones); Release(range); Release(frame); }
    }

    // One character's size and drawn color (#RRGGBB, theme colors resolved).
    private static (float? Size, string? Color) ReadCharacterStyle(Shape shape, int position)
    {
        static float? ValidSize(double value) =>
            value > 0 && value <= 4000 && !double.IsNaN(value) ? (float)value : null;

        if (OfficeHostEnvironment.IsWps)
        {
            PptTextRange? range = null;
            PptFont? font = null;
            ColorFormat? color = null;
            try
            {
                range = WpsNativeEquations.Characters(shape, position, 1);
                font = range.Font;
                float? size = null;
                string? rgb = null;
                try { size = ValidSize(font.Size); } catch { }
                try
                {
                    color = font.Color;
                    rgb = PowerPointConversionColors.FromOfficeRgb(color.RGB);
                }
                catch { }
                return (size, rgb);
            }
            catch { return (null, null); }
            finally { Release(color); Release(font); Release(range); }
        }
        object? frame = null, text = null, character = null, characterFont = null, fill = null, fore = null;
        try
        {
            frame = shape.TextFrame2;
            text = ((dynamic)frame).TextRange;
            character = ((dynamic)text).Characters(position, 1);
            characterFont = ((dynamic)character).Font;
            float? size = null;
            string? rgb = null;
            try { size = ValidSize(Convert.ToDouble(((dynamic)characterFont).Size)); } catch { }
            try
            {
                fill = ((dynamic)characterFont).Fill;
                fore = ((dynamic)fill).ForeColor;
                rgb = PowerPointConversionColors.FromOfficeRgb(Convert.ToInt32(((dynamic)fore).RGB));
            }
            catch { }
            return (size, rgb);
        }
        catch { return (null, null); }
        finally { Release(fore); Release(fill); Release(characterFont); Release(character); Release(text); Release(frame); }
    }

    // ---- Writing: rewriting text in place ------------------------------------

    private PowerPointConversionUnitResult ApplyInPlaceTextUnit(
        PowerPointFormulaConversionPlan plan,
        PowerPointConversionUnit unit,
        IReadOnlyDictionary<string, PreparedPowerPointConversion> prepared)
    {
        var result = new PowerPointConversionUnitResult { SlideId = unit.SlideId, FormulaCount = unit.Items.Count };
        Presentation? presentation = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            EnsureNotSlideShow();
            presentation = _application.ActivePresentation;
            EnsureSourceDocument(presentation, plan.DocumentId);
            string? missing;
            (slide, shape, missing) = FindConversionSource(presentation, unit);
            if (slide is null || shape is null)
            {
                result.Reason = PowerPointConversionSkipReason.Changed;
                result.Detail = missing;
                return result;
            }
            var target = shape;
            // Last to first: a write moves only the text after it, so every
            // other formula keeps the position it was scanned at.
            var items = unit.Items.OrderByDescending(item => item.TextStart).ToList();
            Action write = plan.Request.TargetMode == FormulaOleContract.WordOmmlMode
                ? () => WriteLatexAsEquations(target, items, prepared)
                : () => WriteEquationsAsLatex(target, unit, items, prepared);
            // One transaction per text box: PowerPoint undoes all of its writes if
            // one fails (WPS, without an undo history, reports what changed). The
            // source was proven right before the batch started writing.
            PowerPointTextEditTransaction.Execute(_application, presentation, unit.SlideId,
                new PowerPointTextHost(unit.ShapeId), unit.TextFingerprint!, write, sourceProven: true);
            result.Converted = true;
            result.ReplacementShapeIds.Add(unit.ShapeId);
            return result;
        }
        catch (Exception error)
        {
            result.Converted = false;
            result.Reason = PowerPointConversionSkipReason.Failed;
            result.Detail = ApplyFailure(unit, error, Text("改写文本框中的公式", "rewriting the formulas in the text box"));
            return result;
        }
        finally
        {
            StartNewUndoEntry();
            Release(shape);
            Release(slide);
            Release(presentation);
        }
    }

    // LaTeX code to native equations: the editor's text insertion writer (the
    // code's own size and language; a display equation in its own centered
    // paragraph without a bullet), over the code's range. In WPS every paste is
    // proven to have made a native equation with one saved copy of the
    // presentation for the whole box, not one per formula.
    private void WriteLatexAsEquations(
        Shape shape,
        IReadOnlyList<PowerPointConversionItem> items,
        IReadOnlyDictionary<string, PreparedPowerPointConversion> prepared)
    {
        var wps = OfficeHostEnvironment.IsWps;
        var pastes = new List<(int Start, int Length)>();
        foreach (var item in items)
        {
            var mathMl = prepared[item.Id].MathMl
                ?? throw new InvalidDataException(Text("缺少用于原生公式的 MathML。", "Missing MathML for the native equation."));
            if (!wps)
            {
                PowerPointNativeOmmlKernel.PrepareTextInsertion(
                    _application, shape, item.TextStart, item.TextLength, mathMl, item.DisplayMode)();
                continue;
            }
            (int Start, int Length)? pasted = null;
            var before = WpsNativeEquations.LiveText(shape).Length;
            PowerPointNativeOmmlKernel.PrepareTextInsertion(
                _application, shape, item.TextStart, item.TextLength, mathMl, item.DisplayMode,
                wpsPasted: range => pasted = range)();
            // Formulas are written last to first: the ones pasted before lie after
            // this one and move by its change of length.
            var shift = WpsNativeEquations.LiveText(shape).Length - before;
            for (var index = 0; index < pastes.Count; index++)
                pastes[index] = (pastes[index].Start + shift, pastes[index].Length);
            pastes.Add(pasted ?? throw new InvalidOperationException(
                "WPS reported no pasted range for the native equation."));
        }
        if (pastes.Count == 0) return;
        var zones = WpsNativeEquations.Locate(shape);
        if (pastes.Any(range => !zones.Any(zone => PowerPointNativeOmmlKernel.Covers(zone, range))))
            throw new InvalidOperationException(
                "WPS pasted the equation package without creating a native equation.");
    }

    private sealed class CharacterFormat
    {
        internal string? Name { get; set; }
        internal MsoTriState Italic { get; set; }
    }

    // Native equations to LaTeX code. Each equation is removed first and its
    // code typed where it was, so the code takes the format of the text around
    // it, not the equation's; it keeps the equation's size. Then the text must
    // show every code exactly where it belongs, outside any equation.
    private void WriteEquationsAsLatex(
        Shape shape,
        PowerPointConversionUnit unit,
        IReadOnlyList<PowerPointConversionItem> items,
        IReadOnlyDictionary<string, PreparedPowerPointConversion> prepared)
    {
        var text = ReadConversionText(shape);
        var equations = unit.EquationRanges;
        var formats = new Dictionary<string, (CharacterFormat? Donor, float? Size)>(StringComparer.Ordinal);
        foreach (var item in items)
        {
            var donor = PowerPointLatexText.FormatDonor(text, equations, item.TextStart, item.TextLength);
            formats[item.Id] = (donor > 0 ? ReadCharacterFormat(shape, donor) : null, ReadCharacterStyle(shape, item.TextStart).Size);
        }
        foreach (var item in items)
        {
            var source = prepared[item.Id].LatexSource
                ?? throw new InvalidDataException(Text("缺少公式的 LaTeX 代码。", "Missing the equation's LaTeX code."));
            ReplaceTextRange(shape, item.TextStart, item.TextLength, source);
            var (donor, size) = formats[item.Id];
            FormatLatexCode(shape, item.TextStart, source.Length, donor, size);
        }
        VerifyLatexCode(shape, unit, items, prepared);
    }

    // Assign, never Delete: Delete also removes a space next to the text.
    private static void ReplaceTextRange(Shape shape, int start, int length, string replacement)
    {
        PptTextFrame? frame = null;
        PptTextRange? all = null;
        PptTextRange? range = null;
        try
        {
            frame = shape.TextFrame;
            all = frame.TextRange;
            range = all.Characters(start, length);
            range.Text = string.Empty;
            Release(range);
            range = null;
            Release(all);
            all = frame.TextRange;
            // An emptied text has no character range to type into.
            if (all.Length == 0)
            {
                all.Text = replacement;
                return;
            }
            range = all.Characters(start, 0);
            range.Text = replacement;
        }
        finally { Release(range); Release(all); Release(frame); }
    }

    private static CharacterFormat? ReadCharacterFormat(Shape shape, int position)
    {
        PptTextFrame? frame = null;
        PptTextRange? all = null;
        PptTextRange? character = null;
        PptFont? font = null;
        try
        {
            frame = shape.TextFrame;
            all = frame.TextRange;
            character = all.Characters(position, 1);
            font = character.Font;
            return new CharacterFormat { Name = font.Name, Italic = font.Italic };
        }
        catch { return null; }
        finally { Release(font); Release(character); Release(all); Release(frame); }
    }

    // The code reads as text: a math font or italic that PowerPoint carried over
    // from the equation is replaced by the format of the text around it (the
    // theme's body font where there is none), and it keeps the equation's size.
    // Each fix is tried on its own: one that fails does not skip the others.
    private static void FormatLatexCode(Shape shape, int start, int length, CharacterFormat? donor, float? size)
    {
        PptTextFrame? frame = null;
        PptTextRange? all = null;
        PptTextRange? code = null;
        PptFont? font = null;
        try
        {
            frame = shape.TextFrame;
            all = frame.TextRange;
            code = all.Characters(start, length);
            font = code.Font;
        }
        catch
        {
            Release(font); Release(code); Release(all); Release(frame);
            return;
        }
        try
        {
            string? name = null;
            try { name = font.Name; } catch { }
            var mathFont = string.Equals(name, MathFontName, StringComparison.OrdinalIgnoreCase);
            if (mathFont)
            {
                try
                {
                    var replacement = donor?.Name is { Length: > 0 } donorName
                        && !string.Equals(donorName, MathFontName, StringComparison.OrdinalIgnoreCase)
                        ? donorName
                        : ThemeBodyFont(shape);
                    if (replacement is not null) font.Name = replacement;
                }
                catch { }
            }
            try
            {
                if (donor is not null && donor.Italic is MsoTriState.msoTrue or MsoTriState.msoFalse)
                {
                    if (font.Italic != donor.Italic) font.Italic = donor.Italic;
                }
                else if (mathFont && font.Italic != MsoTriState.msoFalse)
                {
                    font.Italic = MsoTriState.msoFalse;
                }
            }
            catch { }
            try
            {
                if (size is > 0 && Math.Abs(font.Size - size.Value) > 0.01f) font.Size = size.Value;
            }
            catch { }
        }
        finally { Release(font); Release(code); Release(all); Release(frame); }
    }

    // The theme's body font: PowerPoint's theme reference; WPS gets the font's
    // name (null when it cannot be read).
    private static string? ThemeBodyFont(Shape shape)
    {
        if (!OfficeHostEnvironment.IsWps) return "+mn-lt";
        object? slide = null, master = null, theme = null, scheme = null, minor = null, latin = null;
        try
        {
            slide = shape.Parent;
            master = ((dynamic)slide).Master;
            theme = ((dynamic)master).Theme;
            scheme = ((dynamic)theme).ThemeFontScheme;
            minor = ((dynamic)scheme).MinorFont;
            latin = ((dynamic)minor).Item(MsoFontLanguageIndex.msoThemeLatin);
            var name = Convert.ToString(((dynamic)latin).Name);
            return string.IsNullOrWhiteSpace(name) ? null : name;
        }
        catch { return null; }
        finally { Release(latin); Release(minor); Release(scheme); Release(theme); Release(master); Release(slide); }
    }

    private static void VerifyLatexCode(
        Shape shape,
        PowerPointConversionUnit unit,
        IReadOnlyList<PowerPointConversionItem> items,
        IReadOnlyDictionary<string, PreparedPowerPointConversion> prepared)
    {
        var ordered = items.OrderBy(item => item.TextStart).ToList();
        var sources = ordered.Select(item => prepared[item.Id].LatexSource!).ToList();
        var final = PowerPointLatexText.FinalRanges(ordered
            .Select((item, index) => (item.TextStart, item.TextLength, sources[index].Length))
            .ToList());
        var text = ReadConversionText(shape);
        for (var index = 0; index < final.Count; index++)
        {
            var (start, length) = final[index];
            if (start < 1 || start - 1 + length > text.Length
                || !string.Equals(text.Substring(start - 1, length), sources[index], StringComparison.Ordinal))
                throw new InvalidOperationException(Text(
                    "写入后文本框中的 LaTeX 代码与预期不一致。",
                    "After writing, the LaTeX code in the text box is not what was written."));
        }
        var remaining = OfficeHostEnvironment.IsWps
            ? WpsNativeEquations.Locate(shape).ToList()
            : ReadMathZoneRanges(shape);
        if (remaining.Count != unit.EquationRanges.Count - items.Count
            || remaining.Any(zone => final.Any(range => PowerPointLatexText.Overlaps(zone.Start, zone.Length, range.Start, range.Length))))
            throw new InvalidOperationException(Text(
                "LaTeX 代码仍位于原生公式之内，未能变为普通文字。",
                "The LaTeX code is still inside a native equation instead of being ordinary text."));
    }

    // ---- Writing: a formula object becomes a text box with its code -----------

    private static Shape CreateLatexTextReplacement(
        Slide slide,
        PowerPointConversionUnit unit,
        PowerPointConversionItem item,
        PreparedPowerPointConversion formula)
    {
        var source = formula.LatexSource
            ?? throw new InvalidDataException(Text("缺少公式的 LaTeX 代码。", "Missing the equation's LaTeX code."));
        var width = Math.Max(1f, item.VisualWidth);
        var height = Math.Max(1f, item.VisualHeight);
        var (left, top) = PowerPointConversionRules.CenteredAt(item.CenterX, item.CenterY, width, height);
        var shape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
        try
        {
            try { shape.Fill.Visible = MsoTriState.msoFalse; } catch { }
            try { shape.Line.Visible = MsoTriState.msoFalse; } catch { }
            PptTextFrame? frame = null;
            PptTextRange? text = null;
            PptFont? font = null;
            ParagraphFormat? paragraph = null;
            try
            {
                frame = shape.TextFrame;
                // One line per line of code, the box hugging it (TextFrame, which
                // WPS honours; it ignores TextFrame2.WordWrap).
                try { frame.WordWrap = MsoTriState.msoFalse; } catch { }
                try { frame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText; } catch { }
                text = frame.TextRange;
                text.Text = source;
                font = text.Font;
                font.Size = PowerPointLatexText.TextFontSize(item.EffectiveFontSizePt);
                try
                {
                    paragraph = text.ParagraphFormat;
                    paragraph.Alignment = PpParagraphAlignment.ppAlignCenter;
                }
                catch { }
            }
            finally { Release(paragraph); Release(font); Release(text); Release(frame); }
            // Centered where the formula was.
            shape.Left = item.CenterX - shape.Width / 2f;
            shape.Top = item.CenterY - shape.Height / 2f;
            TryApplyRotation(shape, unit.Rotation);
            return shape;
        }
        catch
        {
            TryDelete(shape);
            Release(shape);
            throw;
        }
    }
}
