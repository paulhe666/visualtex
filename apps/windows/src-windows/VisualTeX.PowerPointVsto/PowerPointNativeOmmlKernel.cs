using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using PptTextFrame = Microsoft.Office.Interop.PowerPoint.TextFrame;
using PptTextRange = Microsoft.Office.Interop.PowerPoint.TextRange;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Word = Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.PowerPointVsto;

internal static class PowerPointNativeOmmlKernel
{
    private static readonly object WordHostSync = new();
    private static Word.Application? _wordHost;
    private static Word.Document? _wordScratchDocument;
    private static int _wordHostProcessId;
    private static MethodInfo? _transformMathMlToOmmlMethod;
    private static MethodInfo? _replaceWithPreparedOmmlDirectMethod;

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(
        IntPtr windowHandle,
        out uint processId);

    internal static void Shutdown()
    {
        lock (WordHostSync)
            ResetWordHost();
    }

    internal static void Prewarm()
    {
        lock (WordHostSync)
        {
            ResolveConverterMethods(out _, out _);
            EnsureWordHost();
        }
    }

    internal static Shape CreateStandalone(
        Application application,
        Slide slide,
        string mathMl,
        FormulaMetadata metadata,
        float left,
        float top,
        float width,
        float height)
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

    internal static void InsertStandalone(
        Application application,
        Shape shape,
        string mathMl,
        bool display,
        float fontSize)
    {
        Insert(
            application,
            shape,
            start: 1,
            replacementLength: 0,
            mathMl,
            display,
            fontSize);
    }

    internal static Action PrepareTextInsertion(
        Application application,
        Shape shape,
        int start,
        int length,
        string mathMl,
        string displayMode)
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

        // Resolve the native converter and all pure/Word-scratch preparation
        // before a single character of the PowerPoint source can be replaced.
        var linear = PrepareLinearMath(mathMl, display);

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
                InsertPrepared(application, shape, insertionStart, replacementLength, linear, display, contextFontSize);
            };
        }
        finally
        {
            Release(fullRange2);
            Release(frame2);
        }
    }

    private static void Insert(
        Application application,
        Shape shape,
        int start,
        int replacementLength,
        string mathMl,
        bool display,
        float fontSize)
        => InsertPrepared(application, shape, start, replacementLength, PrepareLinearMath(mathMl, display), display, fontSize);

    private static StyledLinearMath PrepareLinearMath(string mathMl, bool display)
    {
        var linear = ConvertMathMlToUnicodeMath(mathMl, display);
        if (string.IsNullOrWhiteSpace(linear.Text))
            throw new InvalidOperationException(
                "VisualTeX could not derive a linear Office Math expression for PowerPoint.");
        return linear;
    }

    private static void InsertPrepared(Application application, Shape shape, int start,
        int replacementLength, StyledLinearMath linear, bool display, float fontSize)
    {
        var linearText = linear.Text;
        PptTextFrame? legacyFrame = null;
        PptTextRange? legacyRange = null;
        PptTextRange? target = null;
        DocumentWindow? window = null;
        Selection? selection = null;
        object? sourceFrame2 = null;
        object? sourceRange2 = null;
        object? sourceFont2 = null;
        try
        {
            legacyFrame = shape.TextFrame;
            legacyRange = legacyFrame.TextRange;
            if (start < 1
                || start > legacyRange.Length + 1
                || replacementLength < 0
                || start - 1 + replacementLength
                    > legacyRange.Length)
                throw new InvalidOperationException(
                    "The PowerPoint text insertion position is no longer valid.");

            if (legacyRange.Length == 0)
            {
                legacyRange.Text = linearText;
                start = 1;
            }
            else
            {
                target =
                    legacyRange.Characters(
                        start,
                        replacementLength);
                target.Text = linearText;
                Release(target);
                target = null;
                Release(legacyRange);
                legacyRange =
                    legacyFrame.TextRange;
            }

            Release(target);
            target = null;

            // Apply formula-only formatting while the inserted UnicodeMath is
            // still an ordinary, exact text range. After EquationProfessional,
            // PowerPoint may expand Selection.TextRange2 beyond the math zone
            // and include neighboring prose, so mutating that selection can
            // change the font/size of text before or after the formula.
            sourceFrame2 = shape.TextFrame2;
            sourceRange2 = GetCharacterRange(
                ((dynamic)sourceFrame2).TextRange,
                start,
                linearText.Length);
            sourceFont2 = ((dynamic)sourceRange2).Font;
            ((dynamic)sourceFont2).Size =
                Math.Max(5f, fontSize);
            // Apply retained colors to exact inserted UnicodeMath spans before
            // Professional conversion; never format its expanded selection.
            foreach (var color in linear.Colors)
            {
                object? coloredRange = null;
                object? coloredFont = null;
                try
                {
                    coloredRange = GetCharacterRange(((dynamic)sourceFrame2).TextRange, start + color.Start, color.Length);
                    coloredFont = ((dynamic)coloredRange).Font;
                    ((dynamic)coloredFont).Fill.ForeColor.RGB = color.Rgb;
                }
                finally { Release(coloredFont); Release(coloredRange); }
            }
            if (display)
            {
                try
                {
                    ((dynamic)sourceRange2)
                        .ParagraphFormat
                        .Alignment =
                        MsoParagraphAlignment
                            .msoAlignCenter;
                }
                catch { }
            }

            SelectTextRangeForEquation(
                application,
                shape,
                start,
                linearText.Length);

            window =
                application.ActiveWindow
                ?? throw new InvalidOperationException(
                    "PowerPoint has no active window for native equation insertion.");
            selection = GetTextSelectionWithRetry(
                window,
                "before native equation insertion",
                350);
            Release(selection);
            selection = null;
            var canInsertGallery =
                application.CommandBars.GetEnabledMso(
                    "InsertBuildingBlocksEquationsGallery");
            var canInsertNew =
                application.CommandBars.GetEnabledMso(
                    "EquationInsertNew");
            if (!canInsertGallery
                && !canInsertNew)
                throw new InvalidOperationException(
                    "PowerPoint exposes no native equation insertion command at the target text position.");

            application.CommandBars.ExecuteMso(
                canInsertGallery
                    ? "InsertBuildingBlocksEquationsGallery"
                    : "EquationInsertNew");

            selection = GetTextSelectionWithRetry(
                window,
                "after native equation insertion",
                700,
                () =>
                    application.CommandBars.GetEnabledMso(
                        "EquationProfessional")
                    || application.CommandBars.GetEnabledMso(
                        "EquationLinearFormat"));

            ApplyNativeUprightLetters(selection, linear);
            if (application.CommandBars.GetEnabledMso(
                    "EquationProfessional"))
            {
                application.CommandBars.ExecuteMso(
                    "EquationProfessional");
            }
            else if (!application.CommandBars.GetEnabledMso(
                         "EquationLinearFormat"))
            {
                throw new InvalidOperationException(
                    "PowerPoint created a math zone but exposed neither Professional nor Linear equation conversion.");
            }

            Release(selection);
            selection = GetTextSelectionWithRetry(
                window,
                "after professional equation conversion",
                500,
                () => application.CommandBars.GetEnabledMso(
                    "EquationLinearFormat"));
        }
        finally
        {
            Release(sourceFont2);
            Release(sourceRange2);
            Release(sourceFrame2);
            Release(selection);
            Release(window);
            Release(target);
            Release(legacyRange);
            Release(legacyFrame);
        }
    }

    private sealed class LinearColor
    {
        internal int Start;
        internal int Length;
        internal int Rgb;
    }

    private sealed class StyledLinearMath
    {
        internal string Text = string.Empty;
        internal List<LinearColor> Colors = new();
        internal List<int> UprightLetters = new();
    }

    private static void ApplyNativeUprightLetters(Selection selection, StyledLinearMath linear)
    {
        if (linear.UprightLetters.Count == 0) return;
        object? selected = null;
        object? zone = null;
        try
        {
            selected = selection.TextRange2;
            zone = PowerPointOmmlBridge.ReadSingleSelectedMathZone(selected);
            if (zone is null)
                throw new InvalidOperationException("PowerPoint did not expose one exact native zone for upright-math formatting.");
            var actual = Convert.ToString(((dynamic)zone).Text) ?? string.Empty;
            var expectedElements = System.Globalization.StringInfo.ParseCombiningCharacters(linear.Text);
            var actualElements = System.Globalization.StringInfo.ParseCombiningCharacters(actual);
            if (expectedElements.Length != actualElements.Length)
                throw new InvalidOperationException("PowerPoint changed the native linear expression before upright-math formatting; conversion cannot preserve its semantics.");
            var targets = new List<(int Start, int Length)>();
            for (var i = 0; i < expectedElements.Length; i++)
            {
                var sourceStart = expectedElements[i];
                var sourceEnd = i + 1 < expectedElements.Length ? expectedElements[i + 1] : linear.Text.Length;
                var targetStart = actualElements[i];
                var targetEnd = i + 1 < actualElements.Length ? actualElements[i + 1] : actual.Length;
                // Comparison only: do not replace either text with normalized
                // text. Native build-up can encode italic letters as scalars.
                if (linear.Text.Substring(sourceStart, sourceEnd - sourceStart).Normalize(NormalizationForm.FormKC)
                    != actual.Substring(targetStart, targetEnd - targetStart).Normalize(NormalizationForm.FormKC))
                    throw new InvalidOperationException("PowerPoint changed the native linear expression before upright-math formatting; conversion cannot preserve its semantics.");
                if (linear.UprightLetters.Contains(sourceStart)) targets.Add((targetStart, targetEnd - targetStart));
            }
            // Native style assignment may shorten an astral italic scalar.
            // Work backwards so earlier exact offsets remain unchanged.
            foreach (var target in targets.AsEnumerable().Reverse())
            {
                object? character = null;
                object? font = null;
                try
                {
                    character = GetCharacterRange(zone, target.Start + 1, target.Length);
                    font = ((dynamic)character).Font;
                    ((dynamic)font).Italic = MsoTriState.msoFalse;
                }
                finally { Release(font); Release(character); }
            }
            ((dynamic)zone).Select();
        }
        finally { Release(zone); Release(selected); }
    }

    private static StyledLinearMath ConvertMathMlToUnicodeMath(
        string mathMl,
        bool display)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException(
                "PowerPoint OMML conversion requires MathML export data.");

        MethodInfo transform;
        MethodInfo replace;
        lock (WordHostSync)
        {
            ResolveConverterMethods(out transform, out replace);
        }

        string omml;
        try
        {
            omml =
                transform.Invoke(
                    null,
                    new object[] { mathMl })
                as string
                ?? throw new InvalidOperationException(
                    "VisualTeX Word OMML converter returned no OMML.");
        }
        catch (TargetInvocationException error)
            when (error.InnerException
                  is not null)
        {
            throw new InvalidOperationException(
                "VisualTeX Word MathML-to-OMML conversion failed: "
                    + error.InnerException.Message,
                error.InnerException);
        }

        // Word is needed only to ask Office itself for the canonical linear
        // UnicodeMath representation. Starting WINWORD and creating a scratch
        // document for every formula dominated insertion time. Keep one private,
        // hidden Word instance and one scratch document for the lifetime of the
        // PowerPoint add-in instead. The host is never attached to a user's Word
        // process and is shut down when the add-in disconnects.
        lock (WordHostSync)
        {
            EnsureWordHost();
            return LinearizePreparedOmml(
                _wordScratchDocument!,
                replace,
                omml,
                display);
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

    private static void ResolveConverterMethods(
        out MethodInfo transform,
        out MethodInfo replace)
    {
        if (_transformMathMlToOmmlMethod is null
            || _replaceWithPreparedOmmlDirectMethod is null)
        {
            var converter =
                ResolveWordOmmlConverterType();
            _transformMathMlToOmmlMethod =
                converter.GetMethod(
                    "TransformMathMlToOmml",
                    BindingFlags.Static
                        | BindingFlags.NonPublic)
                ?? throw new MissingMethodException(
                    converter.FullName,
                    "TransformMathMlToOmml");
            _replaceWithPreparedOmmlDirectMethod =
                converter.GetMethod(
                    "ReplaceWithPreparedOmmlDirect",
                    BindingFlags.Static | BindingFlags.NonPublic,
                    binder: null,
                    types: new[]
                    {
                        typeof(Word.Document),
                        typeof(Word.Range),
                        typeof(string),
                        typeof(bool),
                        typeof(string),
                        typeof(bool),
                    },
                    modifiers: null)
                ?? throw new MissingMethodException(
                    converter.FullName,
                    "ReplaceWithPreparedOmmlDirect");
        }

        transform = _transformMathMlToOmmlMethod;
        replace = _replaceWithPreparedOmmlDirectMethod;
    }

    private static void EnsureWordHost()
    {
        if (_wordHost is not null
            && _wordScratchDocument is not null)
        {
            try
            {
                _ = _wordHost.Name;
                _ = _wordScratchDocument.Name;
                return;
            }
            catch
            {
                ResetWordHost();
            }
        }

        var existingWordProcessIds =
            Process.GetProcessesByName("WINWORD")
                .Select(process => process.Id)
                .ToHashSet();
        Word.Application? word = null;
        Word.Document? document = null;
        var mayQuitWord = existingWordProcessIds.Count == 0;
        var privateProcessId = 0;
        try
        {
            word = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };

            privateProcessId = TryGetWordProcessId(word);
            if (privateProcessId <= 0)
                privateProcessId = WaitForNewWordProcessId(
                    existingWordProcessIds,
                    900);
            if (privateProcessId <= 0)
                throw new InvalidOperationException(
                    "PowerPoint OMML conversion could not verify ownership of its private Word process.");

            mayQuitWord =
                !existingWordProcessIds.Contains(privateProcessId);
            if (!mayQuitWord)
                throw new InvalidOperationException(
                    "PowerPoint OMML conversion refused to reuse an existing user Word process.");

            document = word.Documents.Add();
            _wordHost = word;
            _wordScratchDocument = document;
            _wordHostProcessId = privateProcessId;
            word = null;
            document = null;
        }
        catch
        {
            if (document is not null)
            {
                try
                {
                    document.Close(
                        Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(document);
            if (word is not null && mayQuitWord)
            {
                try
                {
                    word.Quit(
                        Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(word);
            if (mayQuitWord && privateProcessId > 0)
                EnsurePrivateWordProcessExited(privateProcessId);
            throw;
        }
    }

    private static StyledLinearMath LinearizePreparedOmml(
        Word.Document document,
        MethodInfo replace,
        string omml,
        bool display)
    {
        Word.Range? insertionRange = null;
        Word.Range? formulaRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? linearRange = null;
        try
        {
            ClearWordScratchDocument(document);
            insertionRange =
                document.Range(0, 0);

            try
            {
                formulaRange =
                    replace.Invoke(
                        null,
                        new object?[]
                        {
                            document,
                            insertionRange,
                            omml,
                            display,
                            null,
                            false, // Do not promote the private Word inline scratch range.
                        }) as Word.Range
                    ?? throw new InvalidOperationException(
                        "VisualTeX Word OMML converter returned no materialized formula range.");
            }
            catch (TargetInvocationException error)
                when (error.InnerException
                      is not null)
            {
                throw new InvalidOperationException(
                    "VisualTeX Word native OMML materialization failed: "
                        + error.InnerException.Message,
                    error.InnerException);
            }

            maths = formulaRange.OMaths;
            if (maths.Count <= 0)
                throw new InvalidOperationException(
                    "Microsoft Word materialized no native Office Math object.");

            math = maths[1];
            math.Linearize();
            linearRange = math.Range;
            var raw = Convert.ToString(linearRange.Text) ?? string.Empty;
            var trimChars = new[] { ' ', '\t', '\r', '\n', '\a', '\0' };
            var text = raw.Trim(trimChars);
            if (string.IsNullOrWhiteSpace(text))
                throw new InvalidOperationException(
                    "Microsoft Word produced an empty linear Office Math expression.");
            var prefix = raw.Length - raw.TrimStart(trimChars).Length;
            var result = new StyledLinearMath { Text = text };
            // Keep Office's UnicodeMath alphabet scalars intact. NFKC collapses
            // bold/script/double-struck letters and discards their semantics.
            var elements = System.Globalization.StringInfo.GetTextElementEnumerator(text);
            while (elements.MoveNext())
            {
                var element = elements.GetTextElement();
                Word.Range? character = null;
                Word.Font? font = null;
                try
                {
                    character = document.Range(linearRange.Start + prefix + elements.ElementIndex,
                        linearRange.Start + prefix + elements.ElementIndex + element.Length);
                    font = character.Font;
                    // Only plain Latin/Greek letters are mapped here. Do not
                    // reinterpret letterlike math symbols (for example ℝ) or
                    // claim combining/supplementary upright support.
                    if (element.Length == 1 && font.Italic == 0
                        && ((element[0] >= 'A' && element[0] <= 'Z') || (element[0] >= 'a' && element[0] <= 'z')
                            || (element[0] >= '\u0370' && element[0] <= '\u03FF' && char.IsLetter(element[0]))))
                        result.UprightLetters.Add(elements.ElementIndex);
                    var rgb = (int)font.Color;
                    if (rgb < 0 || rgb > 0xFFFFFF) continue; // Office automatic/mixed.
                    var previous = result.Colors.LastOrDefault();
                    if (previous is not null && previous.Rgb == rgb && previous.Start + previous.Length == elements.ElementIndex)
                        previous.Length += element.Length;
                    else result.Colors.Add(new LinearColor { Start = elements.ElementIndex, Length = element.Length, Rgb = rgb });
                }
                finally { Release(font); Release(character); }
            }
            return result;
        }
        finally
        {
            Release(linearRange);
            Release(math);
            Release(maths);
            Release(formulaRange);
            Release(insertionRange);
            try { ClearWordScratchDocument(document); } catch { }
        }
    }

    private static void ClearWordScratchDocument(
        Word.Document document)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            content.Text = string.Empty;
        }
        finally
        {
            Release(content);
        }
    }

    private static int TryGetWordProcessId(
        Word.Application word)
    {
        try
        {
            var handle = new IntPtr(
                Convert.ToInt32(((dynamic)word).Hwnd));
            if (handle == IntPtr.Zero)
                return 0;
            _ = GetWindowThreadProcessId(
                handle,
                out var processId);
            return unchecked((int)processId);
        }
        catch
        {
            return 0;
        }
    }

    private static void ResetWordHost()
    {
        var document = _wordScratchDocument;
        var word = _wordHost;
        var processId = _wordHostProcessId;
        _wordScratchDocument = null;
        _wordHost = null;
        _wordHostProcessId = 0;

        if (document is not null)
        {
            try
            {
                document.Close(
                    Word.WdSaveOptions.wdDoNotSaveChanges);
            }
            catch { }
        }
        Release(document);
        if (word is not null)
        {
            try
            {
                word.Quit(
                    Word.WdSaveOptions.wdDoNotSaveChanges);
            }
            catch { }
        }
        Release(word);
        if (processId > 0)
            EnsurePrivateWordProcessExited(processId);
    }

    private static void EnsurePrivateWordProcessExited(
        int processId)
    {
        Process? process = null;
        try
        {
            process = Process.GetProcessById(processId);
            if (!string.Equals(
                    process.ProcessName,
                    "WINWORD",
                    StringComparison.OrdinalIgnoreCase))
                return;
            if (process.WaitForExit(750))
                return;

            // This PID is captured from the private Word.Application instance
            // created by this kernel before it is published as the shared host.
            // If Word ignores Quit because an Office COM proxy is still alive,
            // terminate only that owned hidden process; never enumerate/kill a
            // user's unrelated WINWORD process.
            process.Kill();
            process.WaitForExit(750);
        }
        catch (ArgumentException)
        {
            // The private process has already exited.
        }
        catch (InvalidOperationException)
        {
            // The private process has already exited or is no longer available.
        }
        finally
        {
            process?.Dispose();
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

    private static int WaitForNewWordProcessId(
        HashSet<int> existingProcessIds,
        int timeoutMilliseconds)
    {
        var stopwatch =
            Stopwatch.StartNew();
        while (stopwatch.ElapsedMilliseconds
               < timeoutMilliseconds)
        {
            var processId = Process
                .GetProcessesByName("WINWORD")
                .Select(process => process.Id)
                .FirstOrDefault(processId =>
                    !existingProcessIds.Contains(
                        processId));
            if (processId > 0)
                return processId;
            Thread.Sleep(10);
        }
        return 0;
    }

    private static Selection GetTextSelectionWithRetry(
        DocumentWindow window,
        string context,
        int timeoutMilliseconds,
        Func<bool>? ready = null)
    {
        Exception? lastError = null;
        var stopwatch = Stopwatch.StartNew();
        do
        {
            Selection? selection = null;
            try
            {
                selection = window.Selection;
                if (selection is not null
                    && selection.Type
                        == PpSelectionType.ppSelectionText)
                {
                    var isReady = false;
                    try { isReady = ready?.Invoke() ?? true; }
                    catch (COMException error)
                    {
                        lastError = error;
                    }
                    if (isReady)
                        return selection;
                }
            }
            catch (COMException error)
            {
                lastError = error;
                const int RpcCallRejected =
                    unchecked((int)0x80010001);
                const int RpcRetryLater =
                    unchecked((int)0x8001010A);
                if (error.HResult != RpcCallRejected
                    && error.HResult != RpcRetryLater)
                    throw;
            }
            Release(selection);

            if (stopwatch.ElapsedMilliseconds
                >= timeoutMilliseconds)
                break;
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(8);
        }
        while (true);

        throw new InvalidOperationException(
            $"PowerPoint remained busy while preparing the text selection ({context}).",
            lastError);
    }

    private static void SelectTextRangeForEquation(
        Application application,
        Shape shape,
        int start,
        int length)
    {
        Exception? lastError = null;
        for (var attempt = 0; attempt < 8; attempt++)
        {
            PptTextFrame? frame = null;
            PptTextRange? full = null;
            PptTextRange? candidate = null;
            DocumentWindow? window = null;
            try
            {
                // The hidden Word materialization step can leave a
                // different PowerPoint window/view active. Restore the exact
                // presentation/slide that owns this shape before selecting it.
                ActivateShapeView(application, shape);

                // Selecting the shape invalidates some legacy TextRange COM
                // proxies on Office 2021. Reacquire TextFrame/TextRange on every
                // attempt instead of reusing the range used to write the text.
                shape.Select(MsoTriState.msoTrue);
                frame = shape.TextFrame;
                full = frame.TextRange;
                candidate = full.Characters(start, length);
                candidate.Select();

                window = application.ActiveWindow;
                if (window is not null)
                {
                    Selection? verify = null;
                    try
                    {
                        verify = GetTextSelectionWithRetry(
                            window,
                            "selecting the formula source text",
                            180);
                        return;
                    }
                    finally
                    {
                        Release(verify);
                    }
                }

                lastError = new InvalidOperationException(
                    "PowerPoint did not enter text-edit mode for the target formula range.");
            }
            catch (Exception error)
                when (error is COMException
                      || error is InvalidOperationException)
            {
                lastError = error;
            }
            finally
            {
                Release(window);
                Release(candidate);
                Release(full);
                Release(frame);
            }

            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(8 + attempt * 4);
        }

        throw new InvalidOperationException(
            "PowerPoint could not select the target text range for native equation insertion.",
            lastError);
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
