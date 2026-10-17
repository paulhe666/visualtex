using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using PptTextFrame = Microsoft.Office.Interop.PowerPoint.TextFrame;
using PptTextRange = Microsoft.Office.Interop.PowerPoint.TextRange;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.PowerPointVsto;

// Native PowerPoint equations are written the way VisualTeX for macOS writes
// them: the shared writer's OMML (WordOmmlConverter.TransformMathMlToOmml, the
// equation Word receives) as PowerPoint's own mathematical text clipboard
// package (PowerPointMathClipboard), pasted into the target text. No Word
// instance and no linear-format round trip is involved.
internal static class PowerPointNativeOmmlKernel
{
    private static readonly object ConverterSync = new();
    private static MethodInfo? _transformMathMlToOmmlMethod;

    internal static void Prewarm() => _ = ResolveTransform();

    internal static Shape CreateStandalone(
        Application application,
        Slide slide,
        string mathMl,
        FormulaMetadata metadata,
        float left,
        float top,
        float width,
        float height,
        string? baseColor = null)
    {
        Shape? shape = null;
        try
        {
            shape = slide.Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal,
                left,
                top,
                Math.Max(1f, width),
                Math.Max(1f, height));
            try { shape.Fill.Visible = MsoTriState.msoFalse; } catch { }
            try { shape.Line.Visible = MsoTriState.msoFalse; } catch { }
            try
            {
                var frame = shape.TextFrame2;
                frame.MarginLeft = 0;
                frame.MarginRight = 0;
                frame.MarginTop = 0;
                frame.MarginBottom = 0;
                frame.WordWrap = MsoTriState.msoFalse;
                frame.AutoSize = MsoAutoSize.msoAutoSizeNone;
                frame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                frame.TextRange.ParagraphFormat.Alignment =
                    MsoParagraphAlignment.msoAlignCenter;
                Release(frame);
            }
            catch { }

            // A new text box shows uncolored text in the slide's text color
            // (white on a dark theme). A converted equation keeps the color its
            // source showed for those parts (black for MathType/VisualTeX).
            if (baseColor is not null
                && PowerPointConversionColors.NeedsExplicit(baseColor, ReadDefaultTextColor(shape)))
                mathMl = PowerPointConversionColors.ApplyBaseColor(mathMl, baseColor);

            InsertStandalone(
                application,
                shape,
                mathMl,
                string.Equals(
                    metadata.DisplayMode,
                    "block",
                    StringComparison.Ordinal),
                FormulaFontSize.ResolveSemanticFontSize(metadata));

            shape.Left = left;
            shape.Top = top;
            shape.Width = Math.Max(1f, width);
            shape.Height = Math.Max(1f, height);

            var result = shape;
            shape = null;
            return result;
        }
        catch
        {
            if (shape is not null)
            {
                try { shape.Delete(); } catch { }
            }
            throw;
        }
        finally
        {
            Release(shape);
        }
    }

    // The color PowerPoint gives text typed into this (empty) text box.
    private static string? ReadDefaultTextColor(Shape shape)
    {
        object? frame = null, range = null, font = null, fill = null, fore = null;
        try
        {
            frame = shape.TextFrame2;
            range = ((dynamic)frame).TextRange;
            font = ((dynamic)range).Font;
            fill = ((dynamic)font).Fill;
            fore = ((dynamic)fill).ForeColor;
            return PowerPointConversionColors.FromOfficeRgb(Convert.ToInt32(((dynamic)fore).RGB));
        }
        catch { return null; }
        finally { Release(fore); Release(fill); Release(font); Release(range); Release(frame); }
    }

    internal static void InsertStandalone(
        Application application,
        Shape shape,
        string mathMl,
        bool display,
        float fontSize)
    {
        InsertPrepared(
            application,
            shape,
            start: 1,
            replacementLength: 0,
            PowerPointMathClipboard.Build(ConvertMathMlToOmml(mathMl), fontSize, display),
            display);
    }

    // editedZone: the edited equation's native zone (its a14:m as saved). The
    // new equation is pasted back inside it, the way VisualTeX for macOS pastes
    // an edit into the copied original context, so it keeps its justification.
    internal static Action PrepareTextInsertion(
        Application application,
        Shape shape,
        int start,
        int length,
        string mathMl,
        string displayMode,
        string? editedZone = null)
    {
        if (shape.HasTextFrame != MsoTriState.msoTrue)
            throw new InvalidOperationException(
                "The PowerPoint OMML insertion target no longer has an editable text frame.");
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException(
                "PowerPoint text OMML insertion requires MathML export data.");

        var display = string.Equals(
            displayMode,
            "block",
            StringComparison.Ordinal);
        if (!display
            && !string.Equals(
                displayMode,
                "inline",
                StringComparison.Ordinal))
            throw new InvalidDataException(
                "PowerPoint text OMML insertion requires inline or block display mode.");
        if (display && editedZone is not null)
            throw new InvalidDataException(
                "A native equation edit replaces its own zone and adds no paragraphs.");

        // Resolve the native converter and all pure preparation before a single
        // character of the PowerPoint source can be replaced.
        var omml = ConvertMathMlToOmml(mathMl);

        object? frame2 = null;
        object? fullRange2 = null;
        try
        {
            frame2 = shape.TextFrame2;
            fullRange2 = ((dynamic)frame2).TextRange;
            var originalLength =
                Convert.ToInt32(((dynamic)fullRange2).Length);
            if (start < 1
                || start > originalLength + 1
                || length < 0
                || start - 1 + length > originalLength)
                throw new InvalidOperationException(
                    "The PowerPoint text changed while the VisualTeX editor was open; the captured insertion range is no longer valid.");

            // OMML embedded in an existing text box must follow the
            // surrounding PowerPoint body text, not VisualTeX's persistent
            // formula-size preference. Resolve the actual target formatting
            // before any paragraph breaks or formula source text are inserted.
            var contextFontSize = ResolveTextContextFontSize(
                fullRange2,
                start,
                length);
            var package = editedZone is null
                ? PowerPointMathClipboard.Build(omml, contextFontSize, display)
                : PowerPointMathClipboard.BuildEdit(omml, contextFontSize, editedZone);

            var insertionStart = start;
            var replacementLength = length;
            string? paragraphBreaks = null;
            if (display)
            {
                var originalText =
                    Convert.ToString(((dynamic)fullRange2).Text)
                    ?? string.Empty;
                var zeroBasedStart = start - 1;
                var zeroBasedEnd =
                    zeroBasedStart + length;
                var needsLeadingBreak =
                    zeroBasedStart > 0
                    && originalText.Length >= zeroBasedStart
                    && originalText[zeroBasedStart - 1] != '\r';
                var needsTrailingBreak =
                    zeroBasedEnd < originalText.Length
                    && originalText[zeroBasedEnd] != '\r';

                paragraphBreaks =
                    (needsLeadingBreak ? "\r" : string.Empty)
                    + (needsTrailingBreak ? "\r" : string.Empty);
                insertionStart =
                    start + (needsLeadingBreak ? 1 : 0);
                replacementLength = 0;
            }

            return () =>
            {
                if (paragraphBreaks is not null)
                {
                    object? frame = null;
                    object? text = null;
                    object? selected = null;
                    try
                    {
                        frame = shape.TextFrame2;
                        text = ((dynamic)frame).TextRange;
                        selected = GetCharacterRange(text, start, length);
                        ((dynamic)selected).Text = paragraphBreaks;
                    }
                    finally { Release(selected); Release(text); Release(frame); }
                }
                InsertPrepared(application, shape, insertionStart, replacementLength, package, display);
            };
        }
        finally
        {
            Release(fullRange2);
            Release(frame2);
        }
    }

    // PowerPoint's paste of an equation also removes a space right next to it:
    // pasting at the caret in "before |after" gives "before" + equation +
    // "after". An equation pasted into text goes between two of these instead,
    // which are removed again afterwards; text around it stays exactly as it was.
    private const char WordJoiner = '\u2060';

    // Pastes the package over [start, start + replacementLength) of the shape's
    // text as VisualTeX for macOS does (VTApplyPowerPointNativeEquation): the
    // owner slide in view, the user's clipboard kept, a display equation in its
    // own centered paragraph without a bullet, and the caret after the equation.
    private static void InsertPrepared(Application application, Shape shape, int start,
        int replacementLength, byte[] package, bool display)
    {
        PptTextFrame? frame = null;
        PptTextRange? text = null;
        PptTextRange? pasted = null;
        PptTextRange? caret = null;
        try
        {
            frame = shape.TextFrame;
            text = frame.TextRange;
            if (start < 1
                || start > text.Length + 1
                || replacementLength < 0
                || start - 1 + replacementLength
                    > text.Length)
                throw new InvalidOperationException(
                    "The PowerPoint text insertion position is no longer valid.");
            // A display equation lands in its own empty paragraph and a new
            // shape's text is empty: no neighbouring text to protect.
            var joined = !display && text.Length > 0;
            if (joined)
            {
                // Assign, never Delete: Delete also removes a space next to the
                // deleted text, as cutting a word in PowerPoint does.
                caret = text.Characters(start, replacementLength);
                caret.Text = new string(WordJoiner, 2);
                Release(caret);
                caret = null;
                start++;
                replacementLength = 0;
                Release(text);
                text = frame.TextRange;
            }
            var original = text.Text;
            Release(text);
            text = null;

            ActivateShapeView(application, shape);
            using (PowerPointClipboard.Preserve())
                pasted = PasteEquation(frame, start, replacementLength, original, package);
            int pastedStart = pasted.Start, pastedLength = pasted.Length;
            if (joined)
            {
                Release(pasted);
                pasted = null;
                RemoveJoiners(frame, pastedStart, pastedLength);
                pastedStart--;
                text = frame.TextRange;
                pasted = text.Characters(pastedStart, pastedLength);
                Release(text);
                text = null;
            }
            if (!ContainsMath(shape, pastedStart, pastedLength))
                throw new InvalidOperationException(
                    "PowerPoint pasted the equation package without creating a native equation.");
            if (display)
            {
                try
                {
                    var paragraph = pasted.ParagraphFormat;
                    paragraph.Alignment = PpParagraphAlignment.ppAlignCenter;
                    paragraph.Bullet.Visible = MsoTriState.msoFalse;
                    Release(paragraph);
                }
                catch { }
            }

            // A caret failure after a successful paste is not a failed insertion.
            try
            {
                text = frame.TextRange;
                caret = text.Characters(pastedStart + pastedLength, 0);
                caret.Select();
            }
            catch { }
        }
        finally
        {
            Release(caret);
            Release(pasted);
            Release(text);
            Release(frame);
        }
    }

    // TextRange.Paste reads the clipboard at once, and a clipboard listener
    // (input method, clipboard history) still reading the new data makes it
    // paste nothing. Serve those requests first and retry, as CopyShapeInto does
    // for OLE shapes; a paste that inserted nothing must also have removed
    // nothing, so the target range stays exact for the next attempt.
    private static PptTextRange PasteEquation(PptTextFrame frame, int start, int length,
        string original, byte[] package)
    {
        for (var attempt = 1; ; attempt++)
        {
            PptTextRange? text = null;
            PptTextRange? target = null;
            PptTextRange? pasted = null;
            try
            {
                PowerPointClipboard.Place(PowerPointMathClipboard.Format, package);
                PowerPointClipboard.ServeSentMessages(attempt == 1 ? 40 : 150 * attempt);
                text = frame.TextRange;
                // An empty text body has no character range to replace.
                target = text.Length == 0 ? text : text.Characters(start, length);
                if (ReferenceEquals(target, text)) text = null;
                pasted = target.Paste();
                if (pasted is not null && pasted.Length > 0)
                {
                    var result = pasted;
                    pasted = null;
                    return result;
                }
            }
            catch (COMException) when (attempt < 5) { }
            catch (COMException error)
            {
                throw new InvalidOperationException(
                    $"{error.Message} (clipboard held by: {PowerPointClipboard.Holder()})", error);
            }
            finally
            {
                Release(pasted);
                Release(target);
                Release(text);
            }

            if (!string.Equals(ReadText(frame), original, StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "PowerPoint changed the text without pasting the native equation.");
            if (attempt >= 5)
                throw new InvalidOperationException(
                    $"PowerPoint did not paste the native equation (clipboard held by: {PowerPointClipboard.Holder()}).");
        }
    }

    // Removes the joiners on both sides of the pasted equation (trailing first, so
    // the leading one keeps its position).
    private static void RemoveJoiners(PptTextFrame frame, int pastedStart, int pastedLength)
    {
        var text = ReadText(frame);
        var before = pastedStart - 2;
        var after = pastedStart - 1 + pastedLength;
        if (before < 0 || after >= text.Length || text[before] != WordJoiner || text[after] != WordJoiner)
            throw new InvalidOperationException(
                "PowerPoint did not paste the native equation at the expected text position.");
        foreach (var position in new[] { after + 1, before + 1 })
        {
            PptTextRange? all = null;
            PptTextRange? joiner = null;
            try
            {
                all = frame.TextRange;
                joiner = all.Characters(position, 1);
                joiner.Text = string.Empty;
            }
            finally { Release(joiner); Release(all); }
        }
    }

    private static string ReadText(PptTextFrame frame)
    {
        PptTextRange? text = null;
        try
        {
            text = frame.TextRange;
            return text.Text;
        }
        finally { Release(text); }
    }

    private static bool ContainsMath(Shape shape, int start, int length)
    {
        object? frame = null, text = null, range = null, zones = null;
        try
        {
            frame = shape.TextFrame2;
            text = ((dynamic)frame).TextRange;
            range = GetCharacterRange(text, start, length);
            zones = ((dynamic)range).MathZones();
            return Convert.ToInt32(((dynamic)zones).Count) > 0;
        }
        catch (COMException) { return false; }
        finally { Release(zones); Release(range); Release(text); Release(frame); }
    }

    private static string ConvertMathMlToOmml(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException(
                "PowerPoint OMML conversion requires MathML export data.");
        try
        {
            return ResolveTransform().Invoke(null, new object[] { mathMl }) as string
                ?? throw new InvalidOperationException(
                    "VisualTeX Word OMML converter returned no OMML.");
        }
        catch (TargetInvocationException error)
            when (error.InnerException is not null)
        {
            throw new InvalidOperationException(
                "VisualTeX Word MathML-to-OMML conversion failed: "
                    + error.InnerException.Message,
                error.InnerException);
        }
    }

    internal static string ReadNativeMathMl(string omml)
    {
        var converter = ResolveWordOmmlConverterType();
        var method = converter.GetMethod(
            "TransformOmmlToMathMl", BindingFlags.Static | BindingFlags.NonPublic,
            binder: null, types: new[] { typeof(string), typeof(bool) }, modifiers: null)
            ?? throw new MissingMethodException(converter.FullName, "TransformOmmlToMathMl");
        try
        {
            return method.Invoke(null, new object[] { PowerPointNativeOmmlFormatting.ForWordReverseTransform(omml), true }) as string
                ?? throw new InvalidDataException("Native PowerPoint equation returned no MathML.");
        }
        catch (TargetInvocationException error) when (error.InnerException is not null)
        {
            throw new InvalidOperationException("Native PowerPoint source recovery failed: " + error.InnerException.Message, error.InnerException);
        }
    }

    // The shared writer lives in the Word add-in's assembly and is reached by
    // reflection: keep this exact signature in step with WordOmmlConverter
    // (WordOmmlConverterReflectionContractTests).
    private static MethodInfo ResolveTransform()
    {
        lock (ConverterSync)
        {
            if (_transformMathMlToOmmlMethod is null)
            {
                var converter = ResolveWordOmmlConverterType();
                _transformMathMlToOmmlMethod =
                    converter.GetMethod(
                        "TransformMathMlToOmml",
                        BindingFlags.Static | BindingFlags.NonPublic,
                        binder: null,
                        types: new[] { typeof(string) },
                        modifiers: null)
                    ?? throw new MissingMethodException(
                        converter.FullName,
                        "TransformMathMlToOmml");
            }
            return _transformMathMlToOmmlMethod;
        }
    }

    private static Type ResolveWordOmmlConverterType()
    {
        const string assemblyName =
            "VisualTeX.WordVsto";
        var assembly =
            AppDomain.CurrentDomain
                .GetAssemblies()
                .FirstOrDefault(candidate =>
                    string.Equals(
                        candidate.GetName().Name,
                        assemblyName,
                        StringComparison.OrdinalIgnoreCase));

        if (assembly is null)
        {
            var candidates =
                new List<string>();

            var appDomainOverride =
                AppDomain.CurrentDomain.GetData(
                    "VisualTeX.PowerPoint.WordVstoPath")
                as string;
            if (!string.IsNullOrWhiteSpace(
                    appDomainOverride))
                candidates.Add(
                    appDomainOverride!);

            var environmentOverride =
                Environment.GetEnvironmentVariable(
                    "VISUALTEX_POWERPOINT_WORD_VSTO_PATH");
            if (!string.IsNullOrWhiteSpace(
                    environmentOverride))
                candidates.Add(
                    environmentOverride!);

            try
            {
                using var acceptanceKey =
                    Microsoft.Win32.Registry
                        .CurrentUser
                        .OpenSubKey(
                            @"Software\VisualTeX\Acceptance\PowerPointOmml");
                var acceptancePath =
                    acceptanceKey?.GetValue(
                        "CurrentWordVsto")
                    as string;
                if (!string.IsNullOrWhiteSpace(
                        acceptancePath))
                    candidates.Add(
                        acceptancePath!);
            }
            catch { }

            try
            {
                var ownPath =
                    typeof(
                        PowerPointNativeOmmlKernel)
                    .Assembly
                    .Location;
                var directory =
                    Path.GetDirectoryName(
                        ownPath);
                if (!string.IsNullOrWhiteSpace(
                        directory))
                    candidates.Add(
                        Path.Combine(
                            directory!,
                            "VisualTeX.WordVsto.dll"));
            }
            catch { }

            candidates.Add(
                Path.Combine(
                    AppContext.BaseDirectory,
                    "VisualTeX.WordVsto.dll"));

            foreach (var candidate
                     in candidates
                         .Where(path =>
                             !string.IsNullOrWhiteSpace(
                                 path))
                         .Distinct(
                             StringComparer
                                 .OrdinalIgnoreCase))
            {
                if (!File.Exists(candidate))
                    continue;
                try
                {
                    assembly =
                        Assembly.LoadFrom(
                            candidate);
                    break;
                }
                catch (Exception error)
                {
                    throw new InvalidOperationException(
                        $"Unable to load the VisualTeX Word OMML converter from '{candidate}'.",
                        error);
                }
            }

            if (assembly is null)
                throw new FileNotFoundException(
                    "VisualTeX Word OMML converter assembly is unavailable. Tried: "
                        + string.Join(
                            " | ",
                            candidates));
        }

        return assembly.GetType(
            "VisualTeX.WordVsto.WordOmmlConverter",
            throwOnError: true)
            ?? throw new InvalidOperationException(
                "VisualTeX Word OMML converter type is unavailable.");
    }

    private static void ActivateShapeView(
        Application application,
        Shape shape)
    {
        Slide? slide = null;
        Presentation? presentation = null;
        Presentation? activePresentation = null;
        Slide? activeSlide = null;
        DocumentWindows? windows = null;
        DocumentWindow? activeWindow = null;
        DocumentWindow? window = null;
        Microsoft.Office.Interop.PowerPoint.View? activeView = null;
        Microsoft.Office.Interop.PowerPoint.View? view = null;
        try
        {
            slide = shape.Parent as Slide
                ?? throw new InvalidOperationException(
                    "The PowerPoint equation shape is no longer attached to a slide.");
            presentation = slide.Parent as Presentation
                ?? throw new InvalidOperationException(
                    "The PowerPoint equation slide is no longer attached to a presentation.");

            // The normal insertion path already runs in the target slide. Avoid
            // activating the window and calling GotoSlide on every formula; both
            // operations force PowerPoint to refresh its view and were a fixed
            // cost even when nothing had changed.
            try
            {
                activePresentation = application.ActivePresentation;
                activeWindow = application.ActiveWindow;
                activeView = activeWindow?.View;
                activeSlide = activeView?.Slide as Slide;
                if (activePresentation is not null
                    && activeSlide is not null
                    && string.Equals(
                        activePresentation.Name,
                        presentation.Name,
                        StringComparison.OrdinalIgnoreCase)
                    && activeSlide.SlideID == slide.SlideID)
                {
                    activeWindow?.Activate();
                    return;
                }
            }
            catch { }
            finally
            {
                Release(activeSlide);
                activeSlide = null;
                Release(activeView);
                activeView = null;
                Release(activeWindow);
                activeWindow = null;
                Release(activePresentation);
                activePresentation = null;
            }

            windows = presentation.Windows;
            if (windows.Count > 0)
            {
                window = windows[1];
                window.Activate();
                view = window.View;
                view.GotoSlide(slide.SlideIndex);
                System.Windows.Forms.Application.DoEvents();
                return;
            }

            // Fallback for unusual window-less automation states.
            window = application.ActiveWindow;
            if (window is null)
                throw new InvalidOperationException(
                    "PowerPoint has no active document window for native equation insertion.");
            view = window.View;
            view.GotoSlide(slide.SlideIndex);
            System.Windows.Forms.Application.DoEvents();
        }
        finally
        {
            Release(activeSlide);
            Release(activeView);
            Release(activeWindow);
            Release(activePresentation);
            Release(view);
            Release(window);
            Release(windows);
            Release(presentation);
            Release(slide);
        }
    }

    private static float ResolveTextContextFontSize(
        object fullRange,
        int start,
        int length)
    {
        object? candidate = null;
        object? font = null;
        try
        {
            // A zero-length TextRange2 at the captured caret normally carries
            // PowerPoint's exact typing format. For a non-empty replacement,
            // the selected text range is the most faithful context.
            candidate = GetCharacterRange(
                fullRange,
                start,
                length);
            var direct = TryReadPositiveFontSize(
                candidate,
                ref font);
            if (direct.HasValue)
                return direct.Value;
        }
        catch { }
        finally
        {
            Release(font);
            font = null;
            Release(candidate);
            candidate = null;
        }

        var totalLength = 0;
        try
        {
            totalLength = Convert.ToInt32(
                ((dynamic)fullRange).Length);
        }
        catch { }

        // If PowerPoint reports a mixed/undefined caret format, use nearby
        // prose without crossing a paragraph mark. Prefer the character on the
        // right, which represents the run the caret is about to insert into;
        // then fall back to the preceding run.
        foreach (var index in new[] { start, start - 1 })
        {
            if (index < 1 || index > totalLength)
                continue;
            try
            {
                candidate = GetCharacterRange(
                    fullRange,
                    index,
                    1);
                var text = Convert.ToString(
                    ((dynamic)candidate).Text)
                    ?? string.Empty;
                if (text == "\r" || text == "\n")
                    continue;
                var nearby = TryReadPositiveFontSize(
                    candidate,
                    ref font);
                if (nearby.HasValue)
                    return nearby.Value;
            }
            catch { }
            finally
            {
                Release(font);
                font = null;
                Release(candidate);
                candidate = null;
            }
        }

        try
        {
            var whole = TryReadPositiveFontSize(
                fullRange,
                ref font);
            if (whole.HasValue)
                return whole.Value;
        }
        catch { }
        finally
        {
            Release(font);
        }

        return 20f;
    }

    private static float? TryReadPositiveFontSize(
        object range,
        ref object? font)
    {
        Release(font);
        font = null;
        font = ((dynamic)range).Font;
        var raw = Convert.ToDouble(
            ((dynamic)font).Size);
        if (double.IsNaN(raw)
            || double.IsInfinity(raw)
            || raw <= 0
            || raw > 1000)
            return null;
        return (float)raw;
    }

    private static object GetCharacterRange(
        object range,
        int start,
        int length)
    {
        try
        {
            return ((dynamic)range)
                .Characters(
                    start,
                    length);
        }
        catch (COMException error)
        {
            throw new InvalidOperationException(
                $"PowerPoint could not resolve the captured text range ({start}, {length}).",
                error);
        }
    }

    private static void Release(
        object? value)
    {
        if (value is null
            || !Marshal.IsComObject(value))
            return;
        try
        {
            Marshal.ReleaseComObject(value);
        }
        catch { }
    }
}
