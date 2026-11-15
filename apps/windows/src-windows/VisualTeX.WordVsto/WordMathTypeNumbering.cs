using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using static VisualTeX.WordVsto.WordFormulaService;
using static VisualTeX.WordVsto.WordMathTypeDisplayLayout;

namespace VisualTeX.WordVsto;

/// <summary>
/// MathType's numbered displays: the display scaffold, PlaceRef number templates and fields, and the section breaks
/// and heading scopes MathType numbering depends on.
/// </summary>
internal static class WordMathTypeNumbering
{
    internal static void RelocateCompleteMathTypePlaceRefFromRightToLeft(
        Document document,
        InlineShape shape)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Field? sourcePlaceRef = null;
        Range? sourceCode = null;
        Range? sourceResult = null;
        Range? sourceSpan = null;
        Range? formattedField = null;
        Range? destination = null;
        Field? leftPlaceRef = null;
        Field? rightPlaceRef = null;
        Range? rightCode = null;
        Range? rightResult = null;
        Range? removal = null;
        Range? separator = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display equation does not occupy one paragraph while relocating its number.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            sourcePlaceRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: false)
                ?? throw new InvalidOperationException(
                    "The temporary right-numbered MathType row has no MTPlaceRef field to relocate.");
            if (!TryReadCompleteMathTypePlaceRefTemplate(
                    document,
                    sourcePlaceRef,
                    out _))
                throw new InvalidOperationException(
                    "The temporary right-numbered MathType row has an incomplete MTPlaceRef field tree.");

            sourceCode = sourcePlaceRef.Code;
            sourceResult = sourcePlaceRef.Result;
            var bodyEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            var sourceStart = Math.Max(paragraphRange.Start, sourceCode.Start - 1);
            var sourceEnd = Math.Min(bodyEnd, sourceResult.End + 1);
            if (sourceEnd <= sourceStart)
                throw new InvalidOperationException(
                    "Word did not expose a complete MTPlaceRef field span for relocation.");

            // Copy the already-complete outer field as one formatted Word range.
            // This preserves the nested field delimiters atomically; no SEQ field is
            // ever created while MTPlaceRef is half-built in the target position.
            sourceSpan = document.Range(sourceStart, sourceEnd);
            formattedField = sourceSpan.FormattedText;
            destination = document.Range(paragraphRange.Start, paragraphRange.Start);
            destination.FormattedText = formattedField;
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                && string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE"),
                    "left-relocate-after-copy",
                    StringComparison.Ordinal))
                throw new COMException(
                    "Injected MathType left-number relocation failure for rollback acceptance.",
                    unchecked((int)0x800A1710));

            Release(sourceResult);
            sourceResult = null;
            Release(sourceCode);
            sourceCode = null;
            Release(sourcePlaceRef);
            sourcePlaceRef = null;
            Release(shapeRange);
            shapeRange = shape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;

            leftPlaceRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: true)
                ?? throw new InvalidOperationException(
                    "Word did not preserve the relocated MathType number on the left of the OLE.");
            if (!TryReadCompleteMathTypePlaceRefTemplate(
                    document,
                    leftPlaceRef,
                    out _))
                throw new InvalidOperationException(
                    "Word damaged the MTPlaceRef field tree while copying it to the left of the OLE.");

            rightPlaceRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: false)
                ?? throw new InvalidOperationException(
                    "Word lost the original temporary right-side MTPlaceRef before relocation completed.");
            rightCode = rightPlaceRef.Code;
            rightResult = rightPlaceRef.Result;
            bodyEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            var rightFieldStart = Math.Max(paragraphRange.Start, rightCode.Start - 1);
            var removalStart = Math.Max(shapeRange.End, rightFieldStart - 1);
            var removalEnd = Math.Min(bodyEnd, rightResult.End + 1);
            if (removalEnd <= removalStart)
                throw new InvalidOperationException(
                    "Word did not expose the temporary right-side MathType number for removal.");
            removal = document.Range(removalStart, removalEnd);
            if ((removal.Text ?? string.Empty).IndexOf('\t') < 0)
                throw new InvalidOperationException(
                    "The temporary right-side MathType number is missing its separator tab.");
            removal.Delete();

            Release(rightResult);
            rightResult = null;
            Release(rightCode);
            rightCode = null;
            Release(rightPlaceRef);
            rightPlaceRef = null;
            Release(leftPlaceRef);
            leftPlaceRef = null;
            Release(shapeRange);
            shapeRange = shape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;

            leftPlaceRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: true)
                ?? throw new InvalidOperationException(
                    "The final left-numbered MathType row has no MTPlaceRef field.");
            if (!TryReadCompleteMathTypePlaceRefTemplate(
                    document,
                    leftPlaceRef,
                    out _))
                throw new InvalidOperationException(
                    "The final left-numbered MathType row has an incomplete MTPlaceRef field tree.");
            rightPlaceRef = FindMathTypePlaceRefFieldForShape(
                paragraphRange,
                shapeRange,
                numberOnLeft: false);
            if (rightPlaceRef is not null)
                throw new InvalidOperationException(
                    "The temporary right-side MTPlaceRef survived left-number relocation.");

            Range? finalCode = null;
            try
            {
                finalCode = leftPlaceRef.Code;
                var finalCodeText = finalCode.Text ?? string.Empty;
                if (finalCodeText.Length == 0
                    || char.IsWhiteSpace(finalCodeText[finalCodeText.Length - 1]))
                    throw new InvalidOperationException(
                        "The relocated MathType equation number contains trailing whitespace.");
            }
            finally { Release(finalCode); }

            var separatorStart = Math.Max(paragraphRange.Start, shapeRange.Start - 1);
            separator = document.Range(separatorStart, shapeRange.Start);
            if (!string.Equals(separator.Text, "\t", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "The relocated left MathType number is not separated from its OLE by one tab.");
        }
        finally
        {
            Release(separator);
            Release(removal);
            Release(rightResult);
            Release(rightCode);
            Release(rightPlaceRef);
            Release(leftPlaceRef);
            Release(destination);
            Release(formattedField);
            Release(sourceSpan);
            Release(sourceResult);
            Release(sourceCode);
            Release(sourcePlaceRef);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static void BuildIndependentMathTypeDisplayScaffold(
        Document document,
        InlineShape shape,
        bool numbered,
        string mathTypeNumberPosition,
        MathTypeWordOpenXml.NumberTemplate? numberTemplate,
        Range? placeRefSource = null,
        bool centerTabPresent = false)
    {
        Range? shapeRange = null;
        Range? separator = null;
        Field? placeRef = null;
        try
        {
            var numberOnLeft = numbered && string.Equals(
                mathTypeNumberPosition,
                "left",
                StringComparison.OrdinalIgnoreCase);
            var numberOnRight = numbered && string.Equals(
                mathTypeNumberPosition,
                "right",
                StringComparison.OrdinalIgnoreCase);
            if (numbered && !numberOnLeft && !numberOnRight)
                throw new InvalidDataException(
                    "MathType equation number position must be left or right.");
            if (numbered && numberTemplate is null)
                throw new InvalidDataException(
                    "A numbered MathType display equation requires an MTPlaceRef template.");

            shapeRange = shape.Range;
            if (!numberOnLeft && !centerTabPresent)
            {
                // MathType's display style centers the equation on its center tab.
                separator = document.Range(shapeRange.Start, shapeRange.Start);
                separator.Text = "\t";
                Release(separator);
                separator = null;
                Release(shapeRange);
                shapeRange = shape.Range;
            }

            if (!numbered) return;

            if (numberOnLeft)
            {
                // Do not create MTPlaceRef first and then insert a tab at its end.
                // Word can expand the field's boundary to absorb that character,
                // leaving no ordinary separator between the number and the OLE.
                // Materialize the tab first, then create MTPlaceRef before it so
                // the final structure is unambiguously FIELD + TAB + OLE.
                var numberInsertionPosition = shapeRange.Start;
                separator = document.Range(
                    numberInsertionPosition,
                    numberInsertionPosition);
                separator.Text = "\t";
                Release(separator);
                separator = null;
                using (WordOperationMetrics.Measure("MathType.Display.CreatePlaceRef"))
                    placeRef = placeRefSource is not null
                        ? CopyMathTypePlaceRef(document, numberInsertionPosition, placeRefSource)
                        : CreateIndependentMathTypePlaceRef(
                            document,
                            numberInsertionPosition,
                            numberTemplate!);
            }
            else
            {
                separator = document.Range(shapeRange.End, shapeRange.End);
                separator.Text = "\t";
                var placeRefPosition = separator.End;
                using (WordOperationMetrics.Measure("MathType.Display.CreatePlaceRef"))
                    placeRef = placeRefSource is not null
                        ? CopyMathTypePlaceRef(document, placeRefPosition, placeRefSource)
                        : CreateIndependentMathTypePlaceRef(
                            document,
                            placeRefPosition,
                            numberTemplate!);
            }
            try { placeRef.ShowCodes = false; } catch { }
        }
        finally
        {
            Release(placeRef);
            Release(separator);
            Release(shapeRange);
        }
    }

    internal static MathTypeWordOpenXml.NumberTemplate ReadMathTypePlaceRefTemplateForShape(
        Document document,
        InlineShape shape,
        string numberPosition)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Field? placeRef = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display equation does not occupy one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var numberOnLeft = string.Equals(
                numberPosition,
                "left",
                StringComparison.OrdinalIgnoreCase);
            placeRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft)
                ?? throw new InvalidOperationException(
                    "The numbered MathType equation has no readable MTPlaceRef field.");
            return ReadMathTypePlaceRefTemplate(document, placeRef);
        }
        finally
        {
            Release(placeRef);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    internal static MathTypeWordOpenXml.NumberTemplate ResolveMathTypeEditNumberTemplate(
        Document document,
        InlineShape shape,
        MathTypeWordOpenXml.NumberTemplate? sourceTemplate,
        out int createdSectionBreakCodeStart)
    {
        createdSectionBreakCodeStart = -1;
        if (sourceTemplate is not null)
            return sourceTemplate;

        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Field? nearest = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display equation does not occupy one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            nearest = FindNearestMathTypePlaceRefField(
                document,
                shapeRange.Start,
                paragraphRange.Start,
                paragraphRange.End);
            var documentNumberFormat = EquationNumberFormat.Resolve(
                WordEquationNumberFormatAuthority.Resolve(document));
            var template = nearest is not null
                ? ReadMathTypePlaceRefTemplate(document, nearest)
                : MathTypeWordOpenXml.CreateVisualTeXNumberTemplate(
                    documentNumberFormat.Id);

            if (documentNumberFormat.UsesHeading
                && MathTypeNumberTemplateUsesHeading(template))
            {
                EnsureMathTypeHeadingScopeState(
                    document,
                    shapeRange.Start,
                    documentNumberFormat,
                    out createdSectionBreakCodeStart);
            }
            return template;
        }
        finally
        {
            Release(nearest);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    internal static void RebuildMathTypeDisplayScaffold(
        Document document,
        InlineShape shape,
        bool numbered,
        string numberPosition,
        MathTypeWordOpenXml.NumberTemplate? numberTemplate,
        bool updateNumberFields = true,
        bool nativeStylesAlreadyPrepared = false,
        float? knownDisplayColumnWidth = null,
        Range? placeRefSource = null,
        bool keepExistingCenterTab = false)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? paragraphShapes = null;
        Field? placeRef = null;
        Range? placeRefCode = null;
        Range? placeRefResult = null;
        Range? fieldSpan = null;
        Range? prefix = null;
        Range? suffix = null;
        try
        {
            var locateMetric = WordOperationMetrics.Measure("MathType.Scaffold.Locate");
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "MathType numbering can only be changed for a standalone display equation.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            paragraphShapes = paragraphRange.InlineShapes;
            if (paragraphShapes.Count != 1)
                throw new InvalidOperationException(
                    "MathType numbering cannot be changed in a paragraph containing other inline objects.");
            locateMetric?.Dispose();

            // Remove only the native MathType number field. Nested SEQ fields live
            // inside its code range and disappear with the outer MACROBUTTON.
            var removeMetric = WordOperationMetrics.Measure("MathType.Scaffold.RemoveOldNumber");
            placeRef = FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: true)
                ?? FindMathTypePlaceRefFieldForShape(
                    paragraphRange,
                    shapeRange,
                    numberOnLeft: false);
            if (placeRef is not null)
            {
                placeRefCode = placeRef.Code;
                placeRefResult = placeRef.Result;
                var fieldStart = Math.Max(
                    paragraphRange.Start,
                    placeRefCode.Start - 1);
                var fieldEnd = Math.Min(
                    Math.Max(paragraphRange.Start, paragraphRange.End - 1),
                    placeRefResult.End + 1);
                fieldSpan = document.Range(fieldStart, Math.Max(fieldStart, fieldEnd));
                fieldSpan.Delete();
            }
            removeMetric?.Dispose();
            var clearMetric = WordOperationMetrics.Measure("MathType.Scaffold.ClearWhitespace");

            Release(shapeRange);
            shapeRange = shape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            var bodyEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            prefix = document.Range(paragraphRange.Start, shapeRange.Start);
            suffix = document.Range(shapeRange.End, bodyEnd);
            if (!IsMathTypeDisplayScaffoldWhitespace(prefix.Text)
                || !IsMathTypeDisplayScaffoldWhitespace(suffix.Text))
                throw new InvalidOperationException(
                    "MathType numbering was not changed because the display paragraph contains user text outside the equation.");

            // Delete from the end first so the OLE range remains stable while its
            // surrounding tabs are normalized back to a bare equation object.
            if (suffix.End > suffix.Start) suffix.Delete();
            Release(shapeRange);
            shapeRange = shape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            Release(prefix);
            prefix = document.Range(paragraphRange.Start, shapeRange.Start);
            // Batch numbering keeps an unnumbered display's existing center TAB
            // instead of deleting and re-inserting the same character; next to
            // an OLE object that delete cost ~95 ms in a 500-equation document.
            var centerTabPresent = keepExistingCenterTab
                && !string.Equals(numberPosition, "left", StringComparison.OrdinalIgnoreCase)
                && string.Equals(prefix.Text, "\t", StringComparison.Ordinal);
            if (prefix.End > prefix.Start && !centerTabPresent) prefix.Delete();
            clearMetric?.Dispose();

            using (WordOperationMetrics.Measure("MathType.Scaffold.Build"))
            BuildIndependentMathTypeDisplayScaffold(
                document,
                shape,
                numbered,
                numberPosition,
                numberTemplate,
                placeRefSource,
                centerTabPresent);
            using (WordOperationMetrics.Measure("MathType.Scaffold.Configure"))
            ConfigureNewMathTypeDisplayEquation(
                document,
                shape,
                numbered,
                numberPosition,
                updateNestedNumberFields: updateNumberFields,
                nativeStylesAlreadyPrepared: nativeStylesAlreadyPrepared,
                knownDisplayColumnWidth: knownDisplayColumnWidth);
        }
        finally
        {
            Release(suffix);
            Release(prefix);
            Release(fieldSpan);
            Release(placeRefResult);
            Release(placeRefCode);
            Release(placeRef);
            Release(paragraphShapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static bool IsMathTypeDisplayScaffoldWhitespace(string? text)
    {
        if (string.IsNullOrEmpty(text)) return true;
        foreach (var character in text)
        {
            if (char.IsWhiteSpace(character)
                || character is '\t' or '\r' or '\n' or '\v'
                    or '\u0013' or '\u0014' or '\u0015')
                continue;
            return false;
        }
        return true;
    }

    // Batch numbering: copy one finished, verified MTPlaceRef instead of building
    // its nested fields one by one. Fields.Add evaluates each new SEQ at once,
    // re-counting every earlier number (O(n^2) over a batch); a FormattedText
    // copy keeps the source's cached results, and the batch evaluates every
    // number in one pass at the end. source spans exactly one MTPlaceRef field.
    internal static Field CopyMathTypePlaceRef(Document document, int position, Range source)
    {
        Range? insertion = null;
        Range? inserted = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            var length = source.End - source.Start;
            insertion = document.Range(position, position);
            insertion.FormattedText = source.FormattedText;
            inserted = document.Range(position, position + length);
            fields = inserted.Fields;
            if (fields.Count < 1)
                throw new InvalidOperationException("Word did not copy the MathType equation number.");
            field = fields[1];
            code = field.Code;
            if (field.Type != WdFieldType.wdFieldMacroButton
                || code.Start != position + 1
                || !MathTypeEquationReferences.IsMathTypePlaceRefCode(code.Text))
                throw new InvalidOperationException("The copied MathType equation number is not one MTPlaceRef field.");
            var result = field;
            field = null;
            return result;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(inserted);
            Release(insertion);
        }
    }

    internal static Field CreateIndependentMathTypePlaceRef(
        Document document,
        int position,
        MathTypeWordOpenXml.NumberTemplate template)
    {
        if (template.Segments.Count == 0)
            throw new InvalidDataException("MathType MTPlaceRef numbering template is empty.");

        Range? insertion = null;
        Range? outerCode = null;
        Fields? outerNestedFields = null;
        Field? outer = null;
        Field? nested = null;
        try
        {
            insertion = document.Range(position, position);
            outer = document.Fields.Add(
                insertion,
                WdFieldType.wdFieldMacroButton,
                "MTPlaceRef",
                false);
            // Word only treats Fields.Add inside another field's Code range as a
            // genuinely nested field while the outer field is exposing its code.
            // If MTPlaceRef remains in result view, the exact same numeric Range is
            // silently promoted to a sibling document field. That is the direct
            // cause of the left-number '(.)' + escaped SEQ corruption.
            outer.ShowCodes = true;

            // Word itself creates MACROBUTTON MTPlaceRef with one trailing ASCII
            // space in its native field instruction. Keep that native instruction
            // untouched. The previous implementation appended a temporary marker
            // and deleted it after building the nested fields; in the real in-process
            // VSTO path Word can refuse or silently ignore that final Range.Delete,
            // leaking the synthetic tail token or throwing 0x800A1710.
            //
            // Every supported MathType number template ends with literal closing
            // punctuation (currently ')'). Materialize that real final literal first
            // at Code.End. Its original insertion coordinate immediately becomes a
            // strictly interior point of MTPlaceRef. Build all preceding template
            // segments in reverse at that fixed point. No synthetic character and no
            // post-construction deletion are needed, so the visible number naturally
            // ends at its real closing punctuation with zero trailing whitespace.
            outerCode = outer.Code;
            var nativeCodeText = outerCode.Text ?? string.Empty;
            if (nativeCodeText.Length == 0
                || !char.IsWhiteSpace(nativeCodeText[nativeCodeText.Length - 1]))
                throw new InvalidOperationException(
                    "Word did not create MTPlaceRef with its native instruction separator.");

            var terminalSegmentIndex = template.Segments.Count - 1;
            var terminalSegment = template.Segments[terminalSegmentIndex];
            if (terminalSegment.IsField)
                throw new InvalidDataException(
                    "MathType MTPlaceRef numbering template must end with literal closing punctuation.");
            var terminalText = terminalSegment.Value.TrimEnd();
            if (terminalText.Length == 0)
                throw new InvalidDataException(
                    "MathType MTPlaceRef numbering template has no terminal closing punctuation.");

            var insertionPosition = outerCode.End;
            Release(insertion);
            insertion = document.Range(insertionPosition, insertionPosition);
            insertion.InsertAfter(terminalText);
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                && string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE"),
                    "left-scaffold-after-marker",
                    StringComparison.Ordinal))
                throw new COMException(
                    "Injected MathType left scaffold Range failure for rollback acceptance.",
                    unchecked((int)0x800A1710));

            Release(outerCode);
            outerCode = outer.Code;
            if (insertionPosition <= outerCode.Start || insertionPosition >= outerCode.End)
                throw new InvalidOperationException(
                    "Word did not retain the MathType terminal punctuation inside MTPlaceRef.");

            for (var segmentIndex = terminalSegmentIndex - 1;
                 segmentIndex >= 0;
                 segmentIndex--)
            {
                var segment = template.Segments[segmentIndex];
                Release(insertion);
                insertion = document.Range(insertionPosition, insertionPosition);
                if (!segment.IsField)
                {
                    var text = segment.Value;
                    // Word's native MTPlaceRef already supplies the instruction
                    // separator before the template. Avoid duplicating the leading
                    // whitespace stored in VisualTeX/MathType templates; trailing
                    // whitespace remains meaningful only between template tokens.
                    if (segmentIndex == 0)
                        text = text.TrimStart();
                    if (!string.IsNullOrEmpty(text))
                        insertion.InsertAfter(text);
                    continue;
                }

                Release(nested);
                Release(outerNestedFields);
                Release(outerCode);
                outerCode = outer.Code;
                outerNestedFields = outerCode.Fields;
                nested = outerNestedFields.Add(
                    insertion,
                    WdFieldType.wdFieldEmpty,
                    segment.Value.Trim(),
                    false);
                // Word automatically leaves a hidden SEQ ... \\h field with its
                // code displayed. In that state the parent Code.Text substitutes
                // the (empty) child result and the hidden increment disappears from
                // the outer instruction stream. MathType's native MTPlaceRef keeps
                // every child in result view, which also makes the complete outer
                // code serializable/reusable as one field tree.
                try { nested.ShowCodes = false; } catch { }
            }

            Release(outerCode);
            outerCode = outer.Code;
            var completedCodeText = outerCode.Text ?? string.Empty;
            if (completedCodeText.Length == 0
                || char.IsWhiteSpace(completedCodeText[completedCodeText.Length - 1]))
                throw new InvalidOperationException(
                    "Word left trailing whitespace after the completed MathType equation number.");

            if (!TryReadCompleteMathTypePlaceRefTemplate(document, outer, out _))
                throw new InvalidOperationException(
                    "Word detached one or more nested MathType number fields while constructing MTPlaceRef.");
            try { outer.ShowCodes = false; } catch { }
            var result = outer;
            outer = null;
            return result;
        }
        finally
        {
            Release(nested);
            Release(outerNestedFields);
            Release(outerCode);
            Release(insertion);
            Release(outer);
        }
    }

    internal static void ConfigureNewMathTypeDisplayEquation(
        Document document,
        InlineShape shape,
        bool numbered,
        string mathTypeNumberPosition,
        bool updateNestedNumberFields = true,
        WordCharacterFormatting? bodyFormatting = null,
        bool nativeStylesAlreadyPrepared = false,
        MathTypeDisplayParagraphLayout? preservedDisplayParagraphLayout = null,
        float? knownDisplayColumnWidth = null,
        WordParagraphSnapshot? sourceParagraph = null)
    {
        Range? shapeRange = null;
        try
        {
            shapeRange = shape.Range;
            ConfigureNewMathTypeDisplayEquation(
                document,
                shapeRange,
                numbered,
                mathTypeNumberPosition,
                updateNestedNumberFields,
                bodyFormatting,
                nativeStylesAlreadyPrepared,
                preservedDisplayParagraphLayout,
                knownDisplayColumnWidth,
                sourceParagraph);
        }
        finally { Release(shapeRange); }
    }

    internal static void ConfigureNewMathTypeDisplayEquation(
        Document document,
        Range hostRange,
        bool numbered,
        string mathTypeNumberPosition,
        bool updateNestedNumberFields = true,
        WordCharacterFormatting? bodyFormatting = null,
        bool nativeStylesAlreadyPrepared = false,
        MathTypeDisplayParagraphLayout? preservedDisplayParagraphLayout = null,
        float? knownDisplayColumnWidth = null,
        WordParagraphSnapshot? sourceParagraph = null)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        TabStops? tabs = null;
        TabStop? tab = null;
        Field? placeRef = null;
        Range? shapePrefix = null;
        try
        {
            shapeRange = hostRange.Duplicate;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "A new MathType display equation must occupy one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (!nativeStylesAlreadyPrepared)
            {
                using (WordOperationMetrics.Measure(
                           "MathType.Display.CaptureBodyFormatting"))
                    bodyFormatting ??= WordCharacterFormatting.CaptureParagraphMark(paragraphRange);
                using (WordOperationMetrics.Measure(
                           "MathType.Display.EnsureNativeStyles"))
                    EnsureMathTypeNativeStyles(document, bodyFormatting);
            }
            // A converted formula stays in the user's paragraph. Only a body
            // paragraph in the default style becomes MathType's native
            // MTDisplayEquation; a heading, list or custom style is kept and
            // MathType adds just its center/right tab stops.
            var keepUserParagraph = sourceParagraph is not null
                && !sourceParagraph.UsesDefaultStyle;
            object displayStyle = "MTDisplayEquation";
            // Clear direct formatting inherited from the source/blank paragraph
            // first, then apply MathType's native paragraph style. Paragraph.Reset
            // can itself restore Normal, so applying MTDisplayEquation before Reset
            // silently lost the native style in real Word.
            if (!keepUserParagraph)
            {
                using (WordOperationMetrics.Measure(
                           "MathType.Display.ResetAndApplyStyle"))
                {
                    paragraph.Reset();
                    paragraphRange.set_Style(ref displayStyle);
                }
            }
            using (WordOperationMetrics.Measure(
                       "MathType.Display.NormalizeParagraph"))
            {
                NormalizeMathTypeDisplayParagraphFormat(
                    document,
                    paragraph,
                    paragraphRange,
                    preservedDisplayParagraphLayout,
                    knownDisplayColumnWidth,
                    sourceParagraph,
                    keepUserParagraph);
            }
            format = paragraph.Format;
            tabs = format.TabStops;
            using (WordOperationMetrics.Measure(
                       "MathType.Display.ValidateTabs"))
            {
                var hasCenter = false;
                var hasRight = false;
                for (var index = 1; index <= tabs.Count; index++)
                {
                    Release(tab);
                    tab = tabs[index];
                    hasCenter |= tab.Alignment == WdTabAlignment.wdAlignTabCenter;
                    hasRight |= tab.Alignment == WdTabAlignment.wdAlignTabRight;
                }
                if (!hasCenter || !hasRight)
                    throw new InvalidOperationException(
                        "MTDisplayEquation does not contain MathType's center/right tab stops.");
            }
            if (!keepUserParagraph)
            {
                using (WordOperationMetrics.Measure(
                           "MathType.Display.RemoveListNumbering"))
                {
                    try { paragraphRange.ListFormat.RemoveNumbers(); } catch { }
                }
            }

            // Do not infer the OLE boundary from Paragraph.Range.Text. When Word's
            // field-code view is enabled, an embedded equation is exposed as
            // FIELD-BEGIN + " EMBED Equation..." rather than U+0001 even though
            // InlineShape.Range still identifies the exact OLE story interval.
            // Validate the real characters before that interval instead; this is
            // invariant under Alt+F9 / Field.ShowCodes.
            using (WordOperationMetrics.Measure(
                       "MathType.Display.CreatePrefixRange"))
                shapePrefix = document.Range(paragraphRange.Start, shapeRange.Start);
            if (!numbered)
            {
                if (!string.Equals(shapePrefix.Text, "\t", StringComparison.Ordinal))
                {
                    // Word can discard the Flat OPC leading tab when a VisualTeX
                    // display paragraph is replaced in-place. Rebuild the native
                    // MathType center-tab scaffold from the live OLE boundary
                    // instead of accepting an uncentered paragraph or failing the
                    // whole format conversion.
                    if (!IsMathTypeDisplayScaffoldWhitespace(shapePrefix.Text))
                        throw new InvalidOperationException(
                            "The unnumbered MathType display equation contains user text before its OLE object.");
                    shapePrefix.SetRange(shapeRange.Start, shapeRange.Start);
                    shapePrefix.Text = "\t";
                    Release(shapeRange);
                    shapeRange = hostRange.Duplicate;
                    Release(paragraphRange);
                    paragraphRange = paragraph.Range;
                    Release(shapePrefix);
                    shapePrefix = document.Range(paragraphRange.Start, shapeRange.Start);
                }
                if (!string.Equals(shapePrefix.Text, "\t", StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "The unnumbered MathType display equation does not begin with Word's native center tab before its OLE object.");
                return;
            }

            var numberOnLeft = string.Equals(
                mathTypeNumberPosition,
                "left",
                StringComparison.OrdinalIgnoreCase);
            var numberOnRight = string.Equals(
                mathTypeNumberPosition,
                "right",
                StringComparison.OrdinalIgnoreCase);
            if (!numberOnLeft && !numberOnRight)
                throw new InvalidDataException(
                    "MathType equation number position must be left or right.");

            using (WordOperationMetrics.Measure("MathType.Display.VerifyPlaceRef"))
            {
                placeRef = FindMathTypePlaceRefFieldForShape(
                        paragraphRange,
                        shapeRange,
                        numberOnLeft)
                    ?? throw new InvalidOperationException(
                        "The numbered MathType display equation has no MTPlaceRef field on the requested side of its OLE object.");
                if (!TryReadCompleteMathTypePlaceRefTemplate(
                        document,
                        placeRef,
                        out _))
                    throw new InvalidOperationException(
                        "The numbered MathType display equation has a detached or incomplete MTPlaceRef field tree.");
            }
            Range? placeRefCode = null;
            Range? placeRefResult = null;
            Range? separator = null;
            try
            {
                placeRefCode = placeRef.Code;
                placeRefResult = placeRef.Result;
                var fieldStart = Math.Max(paragraphRange.Start, placeRefCode.Start - 1);
                var fieldEnd = Math.Min(paragraphRange.End, placeRefResult.End + 1);
                if (numberOnLeft)
                {
                    if (fieldEnd > shapeRange.Start)
                        throw new InvalidOperationException(
                            "The MathType left equation number is not positioned before the equation object.");
                    // An empty-result MTPlaceRef can report a field-end boundary
                    // that coincides with the following OLE field. The one story
                    // character immediately before InlineShape.Range is still the
                    // actual separator TAB in both result and field-code views.
                    Release(shapePrefix);
                    shapePrefix = document.Range(
                        Math.Max(paragraphRange.Start, shapeRange.Start - 1),
                        shapeRange.Start);
                    if (!string.Equals(shapePrefix.Text, "\t", StringComparison.Ordinal))
                        throw new InvalidOperationException(
                            "The MathType left-numbered display equation has no tab between its number and equation.");
                }
                else
                {
                    if (fieldStart < shapeRange.End)
                        throw new InvalidOperationException(
                            "The MathType right equation number is not positioned after the equation object.");
                    separator = document.Range(shapeRange.End, fieldStart);
                    if ((separator.Text ?? string.Empty).IndexOf('\t') < 0)
                        throw new InvalidOperationException(
                            "The MathType right-numbered display equation has no tab between its equation and number.");
                    if (!string.Equals(shapePrefix.Text, "\t", StringComparison.Ordinal))
                    {
                        var prefix = shapePrefix.Text ?? string.Empty;
                        var codes = string.Join(
                            ",",
                            prefix.Select(ch => $"U+{(int)ch:X4}"));
                        throw new InvalidOperationException(
                            $"The MathType right-numbered display equation does not begin with Word's native center tab before its OLE object. paragraph={paragraphRange.Start}-{paragraphRange.End}; shape={shapeRange.Start}-{shapeRange.End}; prefix={codes}.");
                    }
                }
                if (updateNestedNumberFields)
                    UpdateNestedMathTypeNumberFields(placeRef);
            }
            finally
            {
                Release(separator);
                Release(placeRefResult);
                Release(placeRefCode);
            }
        }
        finally
        {
            Release(shapePrefix);
            Release(placeRef);
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static string StripWordFieldControlCharacters(string text) =>
        (text ?? string.Empty)
            .Replace("\u0013", string.Empty)
            .Replace("\u0014", string.Empty)
            .Replace("\u0015", string.Empty)
            .Replace("\u0001", string.Empty)
            .Replace("\uFFFC", string.Empty)
            .Replace("\r", string.Empty);

    internal static MathTypeWordOpenXml.NumberTemplate ReadMathTypePlaceRefTemplate(
        Document document,
        Field source)
    {
        _ = document;
        Range? sourceCode = null;
        try
        {
            sourceCode = source.Code;
            var sourceText = sourceCode.Text ?? string.Empty;
            var macroIndex = sourceText.IndexOf(
                "MTPlaceRef",
                StringComparison.OrdinalIgnoreCase);
            if (macroIndex < 0)
                throw new InvalidDataException(
                    "The source MathType numbering field is not MTPlaceRef.");

            // Word exposes nested fields inside Field.Code using the same field
            // control characters stored in the document: U+0013 begin, optional
            // U+0014 separate, and U+0015 end. Parse that stream directly so
            // literal MathType punctuation such as '(', '.', '-' and ')' is kept
            // byte-for-byte instead of inferred from COM Range coordinates.
            var pattern = sourceText.Substring(
                macroIndex + "MTPlaceRef".Length);
            var template = new MathTypeWordOpenXml.NumberTemplate();
            var literal = new StringBuilder();
            for (var index = 0; index < pattern.Length;)
            {
                if (pattern[index] != '\u0013')
                {
                    literal.Append(pattern[index]);
                    index++;
                    continue;
                }

                if (literal.Length > 0)
                {
                    template.Segments.Add(
                        MathTypeWordOpenXml.NumberSegment.Text(literal.ToString()));
                    literal.Clear();
                }
                var end = pattern.IndexOf('\u0015', index + 1);
                if (end < 0)
                    throw new InvalidDataException(
                        "MathType MTPlaceRef contains an unterminated nested Word field.");
                var fieldBody = pattern.Substring(index + 1, end - index - 1);
                var separate = fieldBody.IndexOf('\u0014');
                if (separate >= 0)
                    fieldBody = fieldBody.Substring(0, separate);
                if (string.IsNullOrWhiteSpace(fieldBody))
                    throw new InvalidDataException(
                        "MathType MTPlaceRef contains an empty nested Word field.");
                template.Segments.Add(
                    MathTypeWordOpenXml.NumberSegment.Field(fieldBody));
                index = end + 1;
            }
            if (literal.Length > 0)
            {
                // Trailing whitespace is not part of a MathType equation-number
                // format. Never let it enter a reusable template, otherwise a later
                // equation/reference can inherit visible blanks after the number.
                var finalLiteral = literal.ToString().TrimEnd();
                if (finalLiteral.Length > 0)
                    template.Segments.Add(
                        MathTypeWordOpenXml.NumberSegment.Text(finalLiteral));
            }
            if (template.Segments.Count == 0)
                throw new InvalidDataException(
                    "MathType MTPlaceRef numbering template has no usable segments.");
            return template;
        }
        finally { Release(sourceCode); }
    }

    internal static bool TryReadCompleteMathTypePlaceRefTemplate(
        Document document,
        Field source,
        out MathTypeWordOpenXml.NumberTemplate? template)
    {
        template = null;
        Range? sourceCode = null;
        Fields? nestedFields = null;
        try
        {
            sourceCode = source.Code;
            nestedFields = sourceCode.Fields;
            var owners = WordFieldCode.OwnerEntries(nestedFields, sourceCode);
            // A native MTPlaceRef always owns at least the hidden MTEqn increment
            // plus the visible MTEqn current-value field. Heading-aware templates
            // additionally own MTChap/MTSec fields. A literal-only outer field such
            // as "(.)" is a detached/corrupt scaffold and must never be accepted as
            // a valid numbered MathType row, even if it happens to sit on the
            // requested side of the OLE object.
            if (nestedFields.Count - owners < 2) return false;

            var candidate = ReadMathTypePlaceRefTemplate(document, source);
            var hasHiddenEquationIncrement = candidate.Segments.Any(segment =>
                segment.IsField
                && segment.Value.IndexOf("SEQ MTEqn", StringComparison.OrdinalIgnoreCase) >= 0
                && segment.Value.IndexOf("\\h", StringComparison.OrdinalIgnoreCase) >= 0);
            var hasVisibleEquationValue = candidate.Segments.Any(segment =>
                segment.IsField
                && segment.Value.IndexOf("SEQ MTEqn", StringComparison.OrdinalIgnoreCase) >= 0
                && segment.Value.IndexOf("\\c", StringComparison.OrdinalIgnoreCase) >= 0);
            if (!hasHiddenEquationIncrement || !hasVisibleEquationValue)
                return false;

            template = candidate;
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(nestedFields);
            Release(sourceCode);
        }
    }

    private static bool TryReadReusableMathTypePlaceRefTemplate(
        Document document,
        Field source,
        out MathTypeWordOpenXml.NumberTemplate? template)
    {
        template = null;
        Range? visibleNumberRange = null;
        try
        {
            if (!MathTypeEquationReferences.TryGetVisibleNumberRange(
                    document,
                    source,
                    out visibleNumberRange)
                || visibleNumberRange is null)
                return false;
            if (string.IsNullOrWhiteSpace(
                    MathTypeEquationReferences.ReadVisibleNumberText(source)))
                return false;
            return TryReadCompleteMathTypePlaceRefTemplate(
                document,
                source,
                out template);
        }
        finally
        {
            Release(visibleNumberRange);
        }
    }

    internal static bool MathTypeNumberTemplateUsesHeading(
        MathTypeWordOpenXml.NumberTemplate template) =>
        template.Segments.Any(segment =>
            segment.IsField
            && (segment.Value.IndexOf("SEQ MTChap", StringComparison.OrdinalIgnoreCase) >= 0
                || segment.Value.IndexOf("SEQ MTSec", StringComparison.OrdinalIgnoreCase) >= 0));

    internal static Field? FindMathTypePlaceRefFieldForShape(
        Range range,
        Range shapeRange,
        bool numberOnLeft)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? resultRange = null;
        Field? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            fields = WordFormulaHost.GetLocalFields(range);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(resultRange);
                resultRange = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                resultRange = field.Result;
                var fieldStart = Math.Max(range.Start, code.Start - 1);
                var fieldEnd = Math.Min(range.End, resultRange.End + 1);
                int distance;
                if (numberOnLeft)
                {
                    if (fieldEnd > shapeRange.Start) continue;
                    distance = shapeRange.Start - fieldEnd;
                }
                else
                {
                    if (fieldStart < shapeRange.End) continue;
                    distance = fieldStart - shapeRange.End;
                }
                if (distance >= bestDistance) continue;
                Release(best);
                best = field;
                field = null;
                bestDistance = distance;
            }
            return best;
        }
        finally
        {
            Release(resultRange);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    internal static Field? FindNearestMathTypePlaceRefField(
        Document document,
        int position,
        int excludeStart,
        int excludeEnd)
    {
        using var searchMetric = WordOperationMetrics.Measure("MathType.Numbering.NearestTemplate");
        Range? content = null;
        Range? beforeRange = null;
        Range? afterRange = null;
        Range? beforeCode = null;
        Range? afterCode = null;
        Field? before = null;
        Field? after = null;
        try
        {
            content = document.Content;
            var safePosition = Math.Max(content.Start, Math.Min(position, content.End));
            beforeRange = content.Duplicate;
            beforeRange.SetRange(content.Start, safePosition);
            before = MathTypeFieldLocator.Find(beforeRange, MathTypeFieldLocator.PlaceRef,
                forward: false, excludeStart, excludeEnd);
            afterRange = content.Duplicate;
            afterRange.SetRange(safePosition, content.End);
            after = MathTypeFieldLocator.Find(afterRange, MathTypeFieldLocator.PlaceRef,
                forward: true, excludeStart, excludeEnd);

            // Keep actual nearest-template inheritance, including a distant custom
            // MathType template. Only candidate discovery changed: Word searches
            // field codes natively; we never marshal every unrelated EMBED field.
            if (before is not null) beforeCode = before.Code;
            if (after is not null) afterCode = after.Code;
            if (beforeCode is not null && (afterCode is null
                || Math.Abs(beforeCode.Start - 1 - safePosition)
                    <= Math.Abs(afterCode.Start - 1 - safePosition)))
            {
                var result = before;
                before = null;
                return result;
            }
            var afterResult = after;
            after = null;
            return afterResult;
        }
        finally
        {
            Release(afterCode); Release(beforeCode);
            Release(after); Release(before);
            Release(afterRange); Release(beforeRange); Release(content);
        }
    }

    internal static bool HasMathTypeSectionBreak(
        Document document,
        int beforePosition = int.MaxValue)
    {
        using var searchMetric = WordOperationMetrics.Measure("MathType.Numbering.HasSection");
        Range? content = null;
        Range? searchRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            content = document.Content;
            var searchEnd = beforePosition == int.MaxValue
                ? content.End
                : Math.Max(
                    content.Start,
                    Math.Min(content.End, beforePosition + 1));
            if (searchEnd <= content.Start) return false;
            searchRange = document.Range(content.Start, searchEnd);
            fields = searchRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(searchRange);
            Release(content);
        }
    }

    private static bool HasMathTypeSectionBreakBetween(
        Document document,
        int afterPosition,
        int beforePosition)
    {
        using var searchMetric = WordOperationMetrics.Measure("MathType.Numbering.HasSectionBetween");
        Range? content = null;
        Range? searchRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            content = document.Content;
            var start = Math.Max(
                content.Start,
                Math.Min(Math.Min(afterPosition, beforePosition), content.End));
            var end = Math.Max(
                start,
                Math.Min(Math.Max(afterPosition, beforePosition) + 1, content.End));
            if (end <= start) return false;
            searchRange = document.Range(start, end);
            fields = searchRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(searchRange);
            Release(content);
        }
    }

    internal static int EnsureMathTypeHeadingScopeState(
        Document document,
        int formulaPosition,
        EquationNumberFormat format) =>
        EnsureMathTypeHeadingScopeState(
            document,
            formulaPosition,
            format,
            out _);

    internal static void EnsureExistingMathTypeHeadingScopes(
        Document document, string formatId, IReadOnlyList<int> numberedParagraphStarts,
        Range? preparedSectionParagraph = null)
    {
        var format = EquationNumberFormat.Resolve(formatId);
        if (!format.UsesHeading || numberedParagraphStarts.Count == 0) return;
        var scopes = WordEquationHeadings.CaptureHeadingScopesAtPositions(
            document, formatId, numberedParagraphStarts);
        // Insert once per actual heading, from the end of the document. Lower
        // captured positions then stay valid, including when several equations
        // share a scope. The shared inserter preserves existing native breaks.
        foreach (var group in scopes.Where(pair => pair.Value.ScopeStart != int.MinValue)
                     .GroupBy(pair => pair.Value.ScopeStart)
                     .OrderByDescending(group => group.Key))
        {
            var first = group.OrderBy(pair => pair.Key).First();
            var paragraphsBefore = ReadDocumentParagraphCount(document);
            var inserted = EnsureMathTypeHeadingScopeState(
                document, first.Key, format, out _, first.Value, preparedSectionParagraph);
            if (ReadDocumentParagraphCount(document) != paragraphsBefore + (inserted > 0 ? 1 : 0))
                throw new InvalidDataException(
                    "MathType heading initialization changed paragraphs beyond its owned section state.");
        }
    }

    internal sealed class MathTypeHeadingPlan
    {
        internal int ScopeStart;
        internal int ScopeEnd;
        internal int Chapter;
        internal int Section;
    }

    internal static MathTypeHeadingPlan? PrepareMathTypeHeadingScope(
        Document document, int formulaPosition, EquationNumberFormat format,
        ResolvedEquationHeadingScope? preResolvedHeadingScope = null)
    {
        if (!format.UsesHeading) return null;
        (int ScopeStart, int ScopeEnd, string NumberText) scope;
        using (WordOperationMetrics.Measure("MathType.Numbering.ResolveHeading"))
            scope = preResolvedHeadingScope is null
                ? WordEquationHeadings.ResolveHeadingScopeAtPosition(document, formulaPosition, format.Id)
                : (preResolvedHeadingScope.ScopeStart, preResolvedHeadingScope.ScopeEnd, preResolvedHeadingScope.NumberText);
        // No heading means the established zero prefix, never manufactured chapter1.
        if (scope.ScopeStart == int.MinValue || scope.ScopeEnd == int.MinValue
            || scope.ScopeEnd > formulaPosition) return null;
        var parts = (scope.NumberText ?? string.Empty).Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0 || !int.TryParse(parts[0], System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var chapter)) return null;
        var section = 0;
        if (format.HeadingLevel >= 2 && parts.Length >= 2
            && !int.TryParse(parts[1], System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out section)) return null;
        if (chapter == 0 && section == 0) return null;
        if (HasMathTypeSectionBreakBetween(document, scope.ScopeEnd, formulaPosition)) return null;
        return new MathTypeHeadingPlan
        {
            ScopeStart = scope.ScopeStart, ScopeEnd = scope.ScopeEnd, Chapter = chapter, Section = section,
        };
    }

    internal static int EnsureMathTypeHeadingScopeState(
        Document document,
        int formulaPosition,
        EquationNumberFormat format,
        out int createdSectionBreakCodeStart,
        ResolvedEquationHeadingScope? preResolvedHeadingScope = null,
        Range? preparedSectionParagraph = null)
    {
        using var headingMetric = WordOperationMetrics.Measure("MathType.Numbering.HeadingState");
        var plan = PrepareMathTypeHeadingScope(document, formulaPosition, format, preResolvedHeadingScope);
        return ApplyMathTypeHeadingScope(document, formulaPosition, plan, out createdSectionBreakCodeStart,
            preparedSectionParagraph);
    }

    internal static int ApplyMathTypeHeadingScope(
        Document document, int formulaPosition, MathTypeHeadingPlan? plan,
        out int createdSectionBreakCodeStart, Range? preparedSectionParagraph = null)
    {
        createdSectionBreakCodeStart = -1;
        if (plan is null) return 0;
        WordDoubleClickHook.TraceMessage(
            $"mathtype-heading-scope-state chapter={plan.Chapter} section={plan.Section} heading={plan.ScopeStart}:{plan.ScopeEnd} formula={formulaPosition}");
        var insertedLength = InsertMathTypeSectionBreakState(
            document, plan.ScopeEnd, plan.Chapter, plan.Section, preparedSectionParagraph);
        if (insertedLength <= 0) return insertedLength;
        createdSectionBreakCodeStart = FindMathTypeSectionBreakCodeStartBetween(
            document, plan.ScopeEnd, Math.Min(document.Content.End, formulaPosition + insertedLength));
        if (createdSectionBreakCodeStart < 0)
            throw new InvalidOperationException(
                "VisualTeX inserted MathType heading state but could not identify the new section field for transactional rollback.");
        return insertedLength;
    }

    private static int FindMathTypeSectionBreakCodeStartBetween(
        Document document,
        int afterPosition,
        int beforePosition)
    {
        using var searchMetric = WordOperationMetrics.Measure("MathType.Numbering.ResolveNewSection");
        Range? content = null;
        Range? searchRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        var best = int.MaxValue;
        try
        {
            content = document.Content;
            var start = Math.Max(
                content.Start,
                Math.Min(Math.Min(afterPosition, beforePosition), content.End));
            var end = Math.Max(
                start,
                Math.Min(Math.Max(afterPosition, beforePosition) + 1, content.End));
            if (end <= start) return -1;
            searchRange = document.Range(start, end);
            fields = searchRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                best = Math.Min(best, code.Start);
            }
            return best == int.MaxValue ? -1 : best;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(searchRange);
            Release(content);
        }
    }

    internal static void InsertIsolatedMathTypeParagraphBody(
        Document targetDocument,
        Range targetRange,
        string flatOpc)
    {
        Microsoft.Office.Interop.Word.Application? application = null;
        Documents? documents = null;
        Document? stagingDocument = null;
        Range? stagingInsertion = null;
        InlineShapes? stagingShapes = null;
        InlineShape? stagingShape = null;
        Range? stagingShapeRange = null;
        Paragraphs? stagingParagraphs = null;
        Paragraph? stagingParagraph = null;
        Range? stagingParagraphRange = null;
        Range? stagingBody = null;
        Range? formattedBody = null;
        try
        {
            if (!string.Equals(
                    targetRange.Text,
                    BulkInlineFormulaPlaceholder,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The isolated MathType paragraph transfer lost its body placeholder.");

            application = targetDocument.Application;
            documents = application.Documents;
            stagingDocument = documents.Add(Visible: false);
            stagingInsertion = stagingDocument.Range(
                stagingDocument.Content.Start,
                stagingDocument.Content.Start);
            WpsXmlInsertion.Insert(stagingInsertion, flatOpc);

            stagingShapes = stagingDocument.InlineShapes;
            if (stagingShapes.Count != 1)
                throw new InvalidDataException(
                    $"The isolated MathType staging document materialized {stagingShapes.Count} OLE objects instead of one.");
            stagingShape = stagingShapes[1];
            if (!MathTypeOleInterop.IsMathTypeOle(stagingShape))
                throw new InvalidDataException(
                    "The isolated MathType staging document did not materialize Equation.DSMT4.");
            stagingShapeRange = stagingShape.Range;
            stagingParagraphs = stagingShapeRange.Paragraphs;
            if (stagingParagraphs.Count != 1)
                throw new InvalidDataException(
                    "The isolated MathType staging equation spans multiple paragraphs.");
            stagingParagraph = stagingParagraphs[1];
            stagingParagraphRange = stagingParagraph.Range;
            var bodyEnd = Math.Max(
                stagingParagraphRange.Start,
                stagingParagraphRange.End - 1);
            if (stagingShapeRange.Start < stagingParagraphRange.Start
                || stagingShapeRange.End > bodyEnd)
                throw new InvalidDataException(
                    "The isolated MathType staging equation escaped its paragraph body.");

            // Transfer only the completed paragraph BODY. The embedded OLE, WMF
            // presentation, native MathType field, tabs and optional MTPlaceRef
            // field tree are preserved by FormattedText, while the staging <w:p>
            // itself is deliberately excluded. This prevents Word from adopting a
            // new paragraph into a user table immediately following the target.
            stagingBody = stagingDocument.Range(
                stagingParagraphRange.Start,
                bodyEnd);
            formattedBody = stagingBody.FormattedText;
            targetRange.FormattedText = formattedBody;
        }
        finally
        {
            Release(formattedBody);
            Release(stagingBody);
            Release(stagingParagraphRange);
            Release(stagingParagraph);
            Release(stagingParagraphs);
            Release(stagingShapeRange);
            Release(stagingShape);
            Release(stagingShapes);
            Release(stagingInsertion);
            if (stagingDocument is not null)
            {
                try { stagingDocument.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(stagingDocument);
            Release(documents);
            Release(application);
            try { targetDocument.Activate(); } catch { }
        }
    }

    private static void InsertIsolatedMathTypeSectionParagraph(
        Document targetDocument, int insertionPosition, string sectionFlatOpc,
        Range? preparedSectionParagraph = null)
    {
        using var staging = preparedSectionParagraph is null
            ? new MathTypeSectionStaging(targetDocument, sectionFlatOpc) : null;
        var source = preparedSectionParagraph ?? staging!.Paragraph;
        Range? targetBreak = null;
        Range? targetProbe = null;
        Paragraphs? targetParagraphs = null;
        Paragraph? targetParagraph = null;
        Range? targetParagraphRange = null;
        Range? formatted = null;
        try
        {
            targetBreak = targetDocument.Range(insertionPosition, insertionPosition);
            targetBreak.Text = "\r";
            targetProbe = targetDocument.Range(
                insertionPosition, Math.Min(targetDocument.Content.End, insertionPosition + 1));
            targetParagraphs = targetProbe.Paragraphs;
            if (targetParagraphs.Count != 1)
                throw new InvalidOperationException(
                    "Word did not create a dedicated target paragraph for MathType section state.");
            targetParagraph = targetParagraphs[1];
            targetParagraphRange = targetParagraph.Range;
            formatted = source.FormattedText;
            targetParagraphRange.FormattedText = formatted;
        }
        finally
        {
            Release(formatted); Release(targetParagraphRange); Release(targetParagraph);
            Release(targetParagraphs); Release(targetProbe); Release(targetBreak);
        }
    }

    internal static int EnsureDefaultMathTypeSectionBreak(
        Document document,
        int beforePosition = int.MaxValue)
    {
        if (HasMathTypeSectionBreak(document, beforePosition)) return 0;
        return InsertMathTypeSectionBreakState(
            document,
            beforePosition,
            chapter: 1,
            section: 1);
    }

    internal static string CreateMathTypeSectionFlatOpc(int chapter, int section)
    {
        var label = string.Equals(System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName,
            "zh", StringComparison.OrdinalIgnoreCase)
            ? $"公式章 {chapter} 节 {section}" : $"Equation Chapter {chapter} Section {section}";
        return MathTypeWordOpenXml.CreateSectionBreakFlatOpc(label, chapter, section);
    }

    private static int InsertMathTypeSectionBreakState(
        Document document,
        int beforePosition,
        int chapter,
        int section,
        Range? preparedSectionParagraph = null)
    {
        using var sectionMetric = WordOperationMetrics.Measure("MathType.Numbering.InsertSection");
        EnsureMathTypeNativeStyles(document);

        var contentEndBefore = document.Content.End;
        var breakXml = CreateMathTypeSectionFlatOpc(chapter, section);
        Range? insertion = null;
        Field? sectionBreak = null;
        Range? sectionCode = null;
        Range? sectionResult = null;
        Range? sectionFull = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Range? sectionParagraphSplit = null;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = document.Content.End;
            var insertionPosition = beforePosition == int.MaxValue
                ? contentStart
                : Math.Max(
                    contentStart,
                    Math.Min(beforePosition, Math.Max(contentStart, contentEnd - 1)));
            insertion = document.Range(insertionPosition, insertionPosition);
            // Always anchor the hidden section state at the containing paragraph
            // boundary. A shape/bookmark Start inside an MTDisplayEquation row can
            // otherwise place MTEditEquationSection2 after the row's leading tab.
            paragraphs = insertion.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
                insertionPosition = paragraphRange.Start;
                insertion.SetRange(insertionPosition, insertionPosition);
            }
            Release(paragraphRange);
            paragraphRange = null;
            Release(paragraph);
            paragraph = null;
            Release(paragraphs);
            paragraphs = null;

            // Target-document InsertXML can merge this complete <w:p> field tree
            // back into the following MTDisplayEquation paragraph. Materialize it
            // in an isolated hidden Word document first, then transfer the single
            // finished paragraph with FormattedText into a reserved target row.
            InsertIsolatedMathTypeSectionParagraph(
                document,
                insertionPosition,
                breakXml,
                preparedSectionParagraph);
            sectionBreak = FindFirstMathTypeSectionBreakField(document, insertionPosition)
                ?? throw new InvalidOperationException(
                    "Word did not materialize MathType's MTEditEquationSection2 field.");
            if (preparedSectionParagraph is not null)
                RetargetPreparedMathTypeSection(document, sectionBreak, chapter, section);
            sectionCode = sectionBreak.Code;
            sectionResult = sectionBreak.Result;
            sectionFull = document.Range(
                Math.Max(document.Content.Start, sectionCode.Start - 1),
                Math.Min(document.Content.End, sectionResult.End + 1));

            // Do not infer isolation from the total paragraph count. Word can add
            // a paragraph while still leaving MTEditEquationSection2 and the first
            // Equation.DSMT4 in the same row. Inspect the field's actual paragraph
            // and split immediately after the outer field whenever anything else
            // remains before that paragraph mark.
            paragraphs = sectionFull.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "MathType chapter/section break materialized across multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var paragraphContentEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            if (sectionFull.End < paragraphContentEnd)
            {
                sectionParagraphSplit = document.Range(sectionFull.End, sectionFull.End);
                sectionParagraphSplit.Text = "\r";
            }

            Release(paragraphRange);
            paragraphRange = null;
            Release(paragraph);
            paragraph = null;
            Release(paragraphs);
            paragraphs = null;

            Release(sectionFull);
            sectionFull = null;
            Release(sectionResult);
            sectionResult = null;
            Release(sectionCode);
            sectionCode = null;
            sectionCode = sectionBreak.Code;
            sectionResult = sectionBreak.Result;
            sectionFull = document.Range(
                Math.Max(document.Content.Start, sectionCode.Start - 1),
                Math.Min(document.Content.End, sectionResult.End + 1));
            paragraphs = sectionFull.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "MathType chapter/section break materialized across multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.InlineShapes.Count != 0)
                throw new InvalidOperationException(
                    "MathType chapter/section state still shares a paragraph with an OLE equation.");
            for (var fieldIndex = 1; fieldIndex <= paragraphRange.Fields.Count; fieldIndex++)
            {
                Field? paragraphField = null;
                Range? paragraphFieldCode = null;
                try
                {
                    paragraphField = paragraphRange.Fields[fieldIndex];
                    paragraphFieldCode = paragraphField.Code;
                    var fieldCodeText = paragraphFieldCode.Text ?? string.Empty;
                    if (fieldCodeText.IndexOf("MTPlaceRef", StringComparison.OrdinalIgnoreCase) >= 0
                        || fieldCodeText.IndexOf("EMBED ", StringComparison.OrdinalIgnoreCase) >= 0)
                        throw new InvalidOperationException(
                            "MathType chapter/section state still shares a paragraph with equation-owned fields.");
                }
                finally
                {
                    Release(paragraphFieldCode);
                    Release(paragraphField);
                }
            }
            paragraphMark = document.Range(
                Math.Max(paragraphRange.Start, paragraphRange.End - 1),
                paragraphRange.End);
            if (!string.Equals(paragraphMark.Text, "\r", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "MathType chapter/section break has no isolated paragraph mark.");

            // MTEquationSection is a hidden red *character* style.  The paragraph
            // mark must stay with the hidden field.  Deleting that mark merges the
            // hidden character formatting into the following user paragraph, so
            // the next field/text insertion inherits MathType's red state.
            object sectionStyle = "MTEquationSection";
            paragraphRange.set_Style(ref sectionStyle);
            UpdateNestedMathTypeNumberFields(sectionBreak);
            return document.Content.End - contentEndBefore;
        }
        finally
        {
            Release(sectionParagraphSplit);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(sectionFull);
            Release(sectionResult);
            Release(sectionCode);
            Release(sectionBreak);
            Release(insertion);
        }
    }

    private static Field? FindFirstMathTypeSectionBreakField(Document document, int? insertionPosition = null)
    {
        using var searchMetric = WordOperationMetrics.Measure("MathType.Numbering.FindSection");
        Range? search = null;
        try
        {
            search = document.Content;
            if (insertionPosition.HasValue)
            {
                var position = insertionPosition.Value;
                search.SetRange(position, Math.Min(position + 1, search.End));
                search.Expand(WdUnits.wdParagraph);
                return MathTypeFieldLocator.AtCodeStart(search, MathTypeFieldLocator.Section,
                    position + 1);
            }
            return MathTypeFieldLocator.Find(search, MathTypeFieldLocator.Section, forward: true);
        }
        finally { Release(search); }
    }

    internal static void RemoveAllMathTypeSectionBreakFields(Document document)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        var starts = new List<int>();
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    starts.Add(code.Start);
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }

        foreach (var start in starts.OrderByDescending(value => value))
            RemoveMathTypeSectionBreakFieldAtCodeStart(document, start);
    }

    internal static void RemoveMathTypeSectionBreakFieldAtCodeStart(
        Document document,
        int codeStart)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        Range? full = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (code.Start != codeStart
                    || (code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                paragraphs = code.Paragraphs;
                if (paragraphs.Count == 1)
                {
                    paragraph = paragraphs[1];
                    paragraphRange = paragraph.Range.Duplicate;
                }

                // Delete the exact section field created by this operation. Never
                // delete "the first" MathType section field: later-chapter failures
                // must not destroy an earlier chapter's valid native state.
                try { field.Delete(); }
                catch
                {
                    result = field.Result;
                    var start = Math.Max(document.Content.Start, code.Start - 1);
                    var end = Math.Min(
                        document.Content.End,
                        Math.Max(code.End, result.End) + 1);
                    full = document.Range(start, end);
                    full.Delete();
                }

                if (paragraphRange is not null
                    && paragraphRange.InlineShapes.Count == 0)
                {
                    try
                    {
                        object normalStyle = WdBuiltinStyle.wdStyleNormal;
                        paragraphRange.set_Style(ref normalStyle);
                        paragraphRange.Font.Hidden = 0;
                        paragraphRange.Delete();
                    }
                    catch { }
                }
                return;
            }
        }
        catch { }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(full);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
        }
    }
}
