using System.Globalization;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;
using WordApplication = Microsoft.Office.Interop.Word.Application;

namespace VisualTeX.WordVsto;

/// <summary>
/// The only low-level writer for native Word OMML hosts in the rebuilt core.
/// It materializes one OMath and nothing else.
/// </summary>
internal static class WordOmmlHostWriter
{
    internal static bool HasFieldFreeWordNativeNumberHost(
        Document document,
        WordFormulaHostDescriptor source)
    {
        using var metric = WordOperationMetrics.Measure("FormatConversion.OmmlNumbering.ProbeFieldFree");
        if (document is null || source is null)
            return false;
        if (source.Kind != WordFormulaHostKind.Omml
            || !source.Display)
            return false;

        Range? exact = null;
        Range? number = null;
        Fields? fields = null;
        try
        {
            exact = WordFormulaHostSemanticReader.CreateRange(
                document,
                source.Range);
            number = WordNativeOmmlNumbering.TryResolveNumberContentRange(exact);
            if (number is null)
                return false;

            fields = exact.Fields;
            return fields.Count == 0;
        }
        finally
        {
            Release(number);
            Release(fields);
            Release(exact);
        }
    }

    internal static WordFormulaHostWriteResult ReplacePreservingWordNativeNumberHost(
        Document document,
        WordFormulaHostDescriptor source,
        WordFormulaHostWriteRequest request)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (source is null)
            throw new ArgumentNullException(nameof(source));
        if (request is null)
            throw new ArgumentNullException(nameof(request));
        if (source.Kind != WordFormulaHostKind.Omml
            || !source.Display
            || request.Kind != WordFormulaHostKind.Omml
            || !string.Equals(
                request.DisplayMode,
                "block",
                StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "Preserving a Word-native number host requires display OMML -> display OMML.");
        }
        var sourceNumbering =
            WordFormulaNumberingResolver.ResolveLocal(
                document,
                source);
        var preservingManagedNumber =
            request.Numbered
            && sourceNumbering.ContainerKind ==
                WordFormulaNumberingContainerKind.CanonicalNativeOmml;
        var preservingFieldFreeWordNumber =
            !request.Numbered
            && HasFieldFreeWordNativeNumberHost(
                document,
                source);
        if (!preservingManagedNumber
            && !preservingFieldFreeWordNumber)
        {
            throw new InvalidOperationException(
                "The source is not a preservable Word-native numbered OMML host.");
        }
        if (string.IsNullOrWhiteSpace(request.MathMl)
            || !request.MathMl!.TrimStart().StartsWith(
                "<math",
                StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidDataException(
                "OMML replacement requires valid MathML.");
        }

        Range? exact = null;
        Range? replacement = null;
        try
        {
            exact = WordFormulaHostSemanticReader.CreateRange(
                document,
                source.Range);
            var sourceXml = WordOmmlConverter.ReadLocalEquationXml(exact);
            if (!WordOmmlConverter.HasWordNativeEquationNumberHost(sourceXml))
            {
                throw new InvalidDataException(
                    "The source OMML no longer contains one Word-native equation-number host.");
            }

            var semanticOmml =
                WordOmmlConverter.TransformMathMlToOmml(request.MathMl!);
            semanticOmml =
                ApplySemanticOmmlTypography(
                    semanticOmml,
                    request.FontSizePoints);

            // Keep Word's existing equation-number host verbatim. Only the
            // mathematical prefix of m:eqArr/m:e is replaced. VisualTeX does
            // not regenerate, renumber, or reinterpret the native suffix.
            var preservedHostOmml =
                WordOmmlConverter.ReplaceWordNativeEquationNumberHostBody(
                    sourceXml,
                    semanticOmml);

            replacement =
                WordOmmlConverter.ReplaceWithPreparedOmmlDirect(
                    document,
                    exact,
                    preservedHostOmml,
                    display: true,
                    mathFontName: document.OMathFontName);

            // The kept number is part of the equation: it takes the formula's
            // size, as Word sizes a whole equation, number included.
            WordFormulaHostLayout.ApplyOmmlLocalTypography(
                replacement,
                request.FontSizePoints);
            WordFormulaHostLayout.ConfigureDisplayParagraph(
                replacement);

            var described = DescribeInsertedOmml(
                replacement,
                requestedNumbered: request.Numbered);
            if (request.Numbered)
            {
                described.Numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        described);
                if (described.Numbering.ContainerKind !=
                    WordFormulaNumberingContainerKind.CanonicalNativeOmml)
                    throw new InvalidDataException(
                        "Editing a referenced numbered OMML formula lost its native numbering host.");
            }
            return new WordFormulaHostWriteResult
            {
                Host = described,
            };
        }
        finally
        {
            Release(replacement);
            Release(exact);
        }
    }

    internal static WordFormulaHostWriteResult ReplaceExactUnnumberedOmmlHost(
        Document document,
        WordFormulaHostDescriptor source,
        WordFormulaHostWriteRequest request)
    {
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.OmmlHostWriter.ReplaceExactUnnumbered");
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (source is null)
            throw new ArgumentNullException(nameof(source));
        if (request is null)
            throw new ArgumentNullException(nameof(request));
        if (source.Kind != WordFormulaHostKind.Omml
            || request.Kind != WordFormulaHostKind.Omml
            || source.Numbering.Numbered
            || request.Numbered)
            throw new InvalidOperationException(
                "Exact unnumbered OMML replacement requires an unnumbered OMML source and target.");
        if (!string.Equals(
                source.DisplayMode,
                request.DisplayMode,
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "Exact unnumbered OMML replacement requires an unchanged display mode.");
        if (string.IsNullOrWhiteSpace(request.MathMl)
            || !request.MathMl!.TrimStart().StartsWith(
                "<math",
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                "OMML replacement requires valid MathML.");

        Range? exact = null;
        Range? replacement = null;
        try
        {
            exact = WordFormulaHostSemanticReader.CreateRange(
                document,
                source.Range);
            var preparedOmml =
                WordOmmlConverter.TransformMathMlToOmml(
                    request.MathMl!);
            preparedOmml =
                ApplySemanticOmmlTypography(
                    preparedOmml,
                    request.FontSizePoints);

            replacement =
                WordOmmlConverter.ReplaceWithPreparedOmmlDirect(
                    document,
                    exact,
                    preparedOmml,
                    display: source.Display,
                    mathFontName: document.OMathFontName);

            WordFormulaHostLayout.ApplyOmmlLocalTypography(
                replacement,
                request.FontSizePoints);
            if (source.Display)
                WordFormulaHostLayout.ConfigureDisplayParagraph(
                    replacement);

            return new WordFormulaHostWriteResult
            {
                Host = DescribeInsertedOmml(
                    replacement,
                    requestedNumbered: false),
            };
        }
        finally
        {
            Release(replacement);
            Release(exact);
        }
    }

    internal static string ApplySemanticOmmlTypography(
        string omml,
        double fontSizePoints)
    {
        var equation = XElement.Parse(
            WordOmmlConverter.ExtractSingleOMath(omml),
            LoadOptions.PreserveWhitespace);
        XNamespace math =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        XNamespace word =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var halfPoints = ((int)Math.Round(
                FormulaFontSize.NormalizeWordOmmlSize(fontSizePoints) * 2d,
                MidpointRounding.AwayFromZero))
            .ToString(CultureInfo.InvariantCulture);

        foreach (var run in equation.DescendantsAndSelf(math + "r"))
        {
            var properties = run.Element(word + "rPr");
            if (properties is null)
            {
                properties = new XElement(word + "rPr");
                var mathProperties = run.Element(math + "rPr");
                if (mathProperties is not null)
                    mathProperties.AddAfterSelf(properties);
                else
                    run.AddFirst(properties);
            }
            SetWordRunTypography(
                properties,
                word,
                halfPoints);
        }

        foreach (var control in equation.DescendantsAndSelf(math + "ctrlPr"))
        {
            var properties = control.Element(word + "rPr");
            if (properties is null)
            {
                properties = new XElement(word + "rPr");
                control.Add(properties);
            }
            SetWordRunTypography(
                properties,
                word,
                halfPoints);
        }

        return equation.ToString(
            SaveOptions.DisableFormatting);
    }

    private static void SetWordRunTypography(
        XElement properties,
        XNamespace word,
        string halfPoints)
    {
        var size = properties.Element(word + "sz");
        if (size is null)
        {
            size = new XElement(word + "sz");
            properties.Add(size);
        }
        size.SetAttributeValue(
            word + "val",
            halfPoints);

        var complexSize = properties.Element(word + "szCs");
        if (complexSize is null)
        {
            complexSize = new XElement(word + "szCs");
            properties.Add(complexSize);
        }
        complexSize.SetAttributeValue(
            word + "val",
            halfPoints);

        var position = properties.Element(word + "position");
        if (position is null)
        {
            position = new XElement(word + "position");
            properties.Add(position);
        }
        position.SetAttributeValue(
            word + "val",
            "0");
    }

    // A new native equation is imported as one prepared atom at a captured
    // caret. It does not replace a text marker or serialize/reimport the owning
    // paragraph. Location preparation and any heading reads belong to preflight.
    internal static WordFormulaHostWriteResult InsertPreparedNative(
        Document document, Range insertion, WordFormulaHostWriteRequest request,
        string preparedOmml, string? paragraphMarkRunProperties = null)
    {
        if (request.Kind != WordFormulaHostKind.Omml)
            throw new InvalidOperationException("Native OMML creation requires a prepared native target.");
        Range? inserted = null;
        Tables? tables = null;
        Cells? cells = null;
        Cell? cell = null;
        Range? cellRange = null;
        Range? tail = null;
        Paragraphs? cellParagraphs = null;
        OMaths? maths = null;
        OMath? native = null;
        var originalCellParagraphs = 0;
        var originalCellSuffix = string.Empty;
        try
        {
            // Capture the right-hand cell content using native range bounds,
            // not string-length arithmetic. A cell terminator occupies one Word
            // position but emits two characters; whole-story substring probes
            // therefore cannot prove a cell-local transport paragraph.
            tables = insertion.Tables;
            if (tables.Count > 0)
            {
                cells = insertion.Cells;
                if (cells.Count != 1) throw new InvalidDataException("Native creation crosses cell owners.");
                cell = cells[1]; cellRange = cell.Range;
                tail = insertion.Duplicate; tail.SetRange(insertion.End, cellRange.End);
                originalCellSuffix = tail.Text ?? string.Empty;
                cellParagraphs = cellRange.Paragraphs;
                originalCellParagraphs = cellParagraphs.Count;
            }
            WordFormulaMutationTransaction.TraceBoundary(document.Application, "native-create-before-import");
            using (WordOperationMetrics.Measure("FormatConversion.OmmlNative.Import"))
            inserted = WordOmmlConverter.ReplaceWithPreparedOmmlDirect(
                document, insertion, preparedOmml,
                display: string.Equals(request.DisplayMode, "block", StringComparison.OrdinalIgnoreCase),
                mathFontName: document.OMathFontName, allowWordNativeInlinePromotion: true,
                paragraphMarkRunProperties: paragraphMarkRunProperties);
            WordFormulaMutationTransaction.TraceBoundary(document.Application, "native-create-after-import");
            if (originalCellParagraphs > 0)
            {
                maths = inserted.OMaths;
                if (maths.Count != 1) throw new InvalidDataException("Native cell import lost its exact equation.");
                native = maths[1];
                Release(cellRange); cellRange = cell.Range;
                Release(cellParagraphs); cellParagraphs = cellRange.Paragraphs;
                tail.SetRange(inserted.End, cellRange.End);
                if (cellParagraphs.Count == originalCellParagraphs + 1
                    && string.Equals(tail.Text, "\r" + originalCellSuffix, StringComparison.Ordinal))
                {
                    Range? transport = null;
                    try
                    {
                        transport = inserted.Duplicate;
                        transport.SetRange(inserted.End, inserted.End + 1);
                        if (transport.Text != "\r")
                            throw new InvalidDataException("Native cell transport separator is not one paragraph mark.");
                        transport.Delete();
                    }
                    finally { Release(transport); }
                    Release(inserted); inserted = native.Range.Duplicate;
                    Release(cellRange); cellRange = cell.Range;
                    Release(cellParagraphs); cellParagraphs = cellRange.Paragraphs;
                    tail.SetRange(inserted.End, cellRange.End);
                }
                if (cellParagraphs.Count != originalCellParagraphs
                    || !string.Equals(tail.Text ?? string.Empty, originalCellSuffix, StringComparison.Ordinal))
                    throw new InvalidDataException("Native import changed content outside its exact cell insertion boundary.");
            }
            // A number kept by a native edit is part of the equation and takes
            // the formula's size, as Word sizes a whole equation.
            using (WordOperationMetrics.Measure("FormatConversion.OmmlNative.Typography"))
                WordFormulaHostLayout.ApplyOmmlLocalTypography(inserted, request.FontSizePoints);
            WordFormulaHostDescriptor described;
            using (WordOperationMetrics.Measure("FormatConversion.OmmlNative.Describe"))
                described = DescribeInsertedOmml(inserted, request.Numbered);
            if (described.Display)
                using (WordOperationMetrics.Measure("FormatConversion.OmmlNative.DisplayParagraph"))
                    WordFormulaHostLayout.ConfigureDisplayParagraph(inserted);
            return new WordFormulaHostWriteResult { Host = described };
        }
        finally
        {
            Release(native); Release(maths); Release(cellParagraphs); Release(tail);
            Release(cellRange); Release(cell); Release(cells); Release(tables); Release(inserted);
        }
    }

    internal static WordFormulaHostWriteResult Insert(
        WordApplication application,
        Document document,
        Range insertion,
        WordFormulaHostWriteRequest request)
    {
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.OmmlHostWriter.Insert");
        if (request.Kind != WordFormulaHostKind.Omml)
            throw new ArgumentException("OMML writer received a non-OMML request.");
        if (string.IsNullOrWhiteSpace(request.MathMl)
            || !request.MathMl!.TrimStart().StartsWith(
                "<math",
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                "OMML insertion requires valid MathML.");

        // OMML carries no VisualTeX-owned durable identity. Any
        // FormulaId on the request belongs only to the editor operation/session
        // and must never be serialized into the Word document.
        var display = string.Equals(
            request.DisplayMode,
            "block",
            StringComparison.OrdinalIgnoreCase);

        Range? inserted = null;
        Range? target = null;
        try
        {
            WordFormulaMutationTransaction.TraceBoundary(application, "omml-writer-enter");
            string preparedOmml;
            using (WordOperationMetrics.Measure("FormatConversion.OmmlHostWriter.Transform"))
                preparedOmml = request.PreparedOmml
                    ?? WordOmmlConverter.TransformMathMlToOmml(request.MathMl!);
            WordFormulaMutationTransaction.TraceBoundary(application, "omml-writer-transformed");

            // Numbered OMML is one native Word host, not an unnumbered OMath
            // followed by a second rebuild. Materialize the final #(SEQ) form
            // once at the target position.
            if (display && request.Numbered && request.PreparedOmml is null)
            {
                preparedOmml =
                    WordNativeOmmlNumbering.PrepareNumberedOmml(
                        document,
                        insertion.Start,
                        preparedOmml);
            }

            // A selected external OLE is a physical replacement, not a text
            // marker insertion. Do not serialize/reimport the containing live
            // paragraph: it can include deleted predecessor objects. The exact
            // source envelope also leaves neighbouring native OMaths Word-owned.
            if (!WordOleToOmmlReplacement.TryReplace(
                    document, insertion, preparedOmml, display, out inserted, request.PreparedOleReplacement))
            {
                if (!display)
                {
                    using var adjacentMetric = WordOperationMetrics.Measure(
                        "FormatConversion.OmmlHostWriter.AdjacentMerge");
                    var preparedInline =
                        PrepareAdjacentInlineNativeMerge(
                            document,
                            insertion,
                            preparedOmml);
                    preparedOmml = preparedInline.Omml;
                    target = preparedInline.Target;
                }
                else
                {
                    target = insertion.Duplicate;
                }

                if (request.BatchMutation
                    && target.Start == target.End)
                {
                    // A batch writer (LaTeX redraw) has already cleared the
                    // source text and inserts at that collapsed caret, exactly
                    // like an interactive create. Import the prepared equation
                    // atom directly, as InsertPreparedNative does, instead of
                    // serializing, marking and re-importing the whole owning
                    // paragraph for every formula. A caret left in an empty
                    // paragraph keeps that paragraph mark's own formatting.
                    string? markRunProperties;
                    using (WordOperationMetrics.Measure("FormatConversion.OmmlHostWriter.MarkRunProperties"))
                        markRunProperties = ReadEmptyParagraphMarkRunProperties(target);
                    var written = InsertPreparedNative(
                        document,
                        target,
                        request,
                        preparedOmml,
                        paragraphMarkRunProperties: markRunProperties);
                    WordFormulaMutationTransaction.TraceBoundary(application, "omml-writer-complete");
                    return written;
                }

                // New OMML is materialized in the actual target paragraph. No
                // durable VisualTeX identity, scratch document or layout table
                // participates in the final Word structure.
                inserted = WordOmmlConverter.ReplaceWithPreparedOmmlInParagraph(
                    document,
                    target,
                    preparedOmml,
                    display,
                    mathFontName: document.OMathFontName,
                    allowWordNativeInlinePromotion: !display);
            }

            WordFormulaMutationTransaction.TraceBoundary(application, "omml-writer-before-typography");
            WordFormulaHostLayout.ApplyOmmlLocalTypography(
                inserted,
                request.FontSizePoints);
            WordFormulaMutationTransaction.TraceBoundary(application, "omml-writer-after-typography");
            // OMML has no VisualTeX-owned durable identity. FormulaId is
            // operation/session state only and is never serialized into Word.

            var described = DescribeInsertedOmml(
                inserted,
                request.Numbered);
            WordFormulaMutationTransaction.TraceBoundary(application, "omml-writer-after-describe");
            if (described.Display)
                WordFormulaHostLayout.ConfigureDisplayParagraph(
                    inserted);
            WordFormulaMutationTransaction.TraceBoundary(application, "omml-writer-complete");
            return new WordFormulaHostWriteResult
            {
                Host = described,
            };
        }
        catch
        {
            if (inserted is not null)
            {
                try { inserted.Delete(); } catch { }
            }
            throw;
        }
        finally
        {
            Release(target);
            Release(inserted);
        }
    }

    // Character formatting of the paragraph mark when the caret sits in an
    // otherwise empty paragraph (only then does Word keep the imported
    // package's mark). Null for a caret inside text.
    private static string? ReadEmptyParagraphMarkRunProperties(Range caret)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = caret.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.End - paragraphRange.Start != 1) return null;
            return WordOmmlInsertionXml.ParagraphMarkRunProperties(
                paragraphRange.WordOpenXML);
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static (Range Target, string Omml)
        PrepareAdjacentInlineNativeMerge(
            Document document,
            Range insertion,
            string preparedOmml)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? exact = null;
        Range? leftRange = null;
        Range? rightRange = null;
        try
        {
            paragraphs =
                insertion.Paragraphs;
            if (paragraphs.Count != 1)
                return (
                    insertion.Duplicate,
                    preparedOmml);

            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            maths =
                paragraphRange.OMaths;
            var sourceStart =
                insertion.Start;
            var sourceEnd =
                insertion.End;
            for (var index = 1;
                 index <= maths.Count;
                 index++)
            {
                Release(exact);
                exact = null;
                Release(math);
                math = maths[index];
                if (math.Type !=
                    WdOMathType.wdOMathInline)
                    continue;

                exact =
                    math.Range.Duplicate;
                if (exact.End == sourceStart
                    && leftRange is null)
                {
                    leftRange =
                        exact.Duplicate;
                }
                if (exact.Start == sourceEnd
                    && rightRange is null)
                {
                    rightRange =
                        exact.Duplicate;
                }
            }

            if (leftRange is null
                && rightRange is null)
                return (
                    insertion.Duplicate,
                    preparedOmml);

            var targetStart =
                sourceStart;
            var targetEnd =
                sourceEnd;
            string combinedOmml;

            if (leftRange is not null
                && rightRange is not null)
            {
                combinedOmml =
                    WordOmmlConverter.CombineInlineOmml(
                        leftRange.WordOpenXML,
                        preparedOmml,
                        rightRange.WordOpenXML);
                targetStart =
                    leftRange.Start;
                targetEnd =
                    rightRange.End;
            }
            else if (leftRange is not null)
            {
                combinedOmml =
                    WordOmmlConverter.CombineInlineOmml(
                        leftRange.WordOpenXML,
                        preparedOmml);
                targetStart =
                    leftRange.Start;
                targetEnd =
                    sourceEnd;
            }
            else
            {
                combinedOmml =
                    WordOmmlConverter.CombineInlineOmml(
                        preparedOmml,
                        rightRange!.WordOpenXML);
                targetStart =
                    sourceStart;
                targetEnd =
                    rightRange.End;
            }

            return (
                document.Range(
                    targetStart,
                    targetEnd),
                combinedOmml);
        }
        finally
        {
            Release(rightRange);
            Release(leftRange);
            Release(exact);
            Release(math);
            Release(maths);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static Range DeleteExactToPlainTextBoundary(
        Document document,
        WordFormulaHostDescriptor host)
    {
        using var metric = WordOperationMetrics.Measure("FormatConversion.OmmlDelete.PreparePlainBoundary");
        if (host.Kind != WordFormulaHostKind.Omml)
            throw new ArgumentException(
                "OMML plain-boundary delete received another host kind.");

        Range? deletedAnchor = null;
        Range? guard = null;
        Range? anchor = null;
        try
        {
            // There must be exactly one OMML deletion primitive. DeleteExact()
            // clears the exact OMath text and proves from a fresh document range
            // that Word actually removed the mathematical host. OMath.Remove()
            // is not equivalent for wdOMathDisplay: Word can keep/recreate the
            // surrounding oMathPara, leaving the replacement boundary inside
            // math mode.
            deletedAnchor =
                DeleteExact(
                    document,
                    host);
            var start =
                deletedAnchor.Start;

            // Never reuse the collapsed Range that existed while the OMath was
            // alive. Reacquire an ordinary document boundary, then keep one
            // non-whitespace character alive during OLE insertion so Word cannot
            // rematerialize a display OMath around the new external host.
            Release(deletedAnchor);
            deletedAnchor = null;
            guard =
                document.Range(
                    start,
                    start);
            guard.Text =
                "x";
            Release(guard);
            guard = null;

            anchor =
                document.Range(
                    start + 1,
                    start + 1);

            Range? verify = null;
            OMaths? verifyMaths = null;
            try
            {
                var contentStart =
                    document.Content.Start;
                var contentEnd =
                    document.Content.End;
                var verifyStart =
                    Math.Max(
                        contentStart,
                        start - 1);
                var verifyEnd =
                    Math.Min(
                        contentEnd,
                        start + 1);
                verify =
                    document.Range(
                        verifyStart,
                        verifyEnd);
                verifyMaths =
                    verify.OMaths;
                if (verifyMaths.Count != 0)
                    throw new InvalidDataException(
                        "Cross-format OMML deletion left mathematical affinity at the fresh insertion boundary.");
            }
            finally
            {
                Release(verifyMaths);
                Release(verify);
            }

            var returned =
                anchor;
            anchor = null;
            return returned;
        }
        finally
        {
            Release(anchor);
            Release(guard);
            Release(deletedAnchor);
        }
    }

    internal static Range DeleteExactToExternalDisplayBoundary(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (host.Kind != WordFormulaHostKind.Omml
            || !host.Display)
            throw new ArgumentException(
                "External display-boundary delete requires one display OMML host.");

        Range? range = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? exact = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefix = null;
        Range? suffix = null;
        Range? followingWitness = null;
        Range? refreshedParagraphRange = null;
        Range? body = null;
        Range? refreshedWitness = null;
        Range? anchor = null;
        try
        {
            range = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            maths = range.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    "The display OMML source no longer contains exactly one equation.");
            math = maths[1];
            exact = math.Range.Duplicate;
            if (!WordFormulaHostSemanticReader.SameAddress(
                    exact,
                    host.Range)
                || math.Type != WdOMathType.wdOMathDisplay)
                throw new InvalidDataException(
                    "The external replacement source is no longer the exact display OMath.");

            paragraphs = exact.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "A display OMML external replacement must occupy one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            var bodyEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            if (exact.Start < paragraphRange.Start
                || exact.End > bodyEnd)
                throw new InvalidDataException(
                    "The display OMath escapes its owning paragraph body.");

            prefix = document.Range(
                paragraphRange.Start,
                exact.Start);
            suffix = document.Range(
                exact.End,
                bodyEnd);
            if (!ContainsOnlyExternalDisplayScaffold(prefix.Text)
                || !ContainsOnlyExternalDisplayScaffold(suffix.Text))
                throw new InvalidDataException(
                    "VisualTeX refused to remove a display OMML paragraph that contains non-structural text outside the equation.");

            var witnessStart = paragraphRange.End;
            var witnessSpan = Math.Min(
                32,
                Math.Max(0, document.Content.End - witnessStart));
            var witnessText = string.Empty;
            if (witnessSpan > 0)
            {
                followingWitness = document.Range(
                    witnessStart,
                    witnessStart + witnessSpan);
                witnessText = followingWitness.Text ?? string.Empty;
            }

            // Use the one canonical OMML deletion primitive. DeleteExact()
            // clears the exact OMath and then proves from a freshly acquired
            // document range that no mathematical host/affinity survived. The
            // older display-boundary path duplicated only the text-clearing half
            // of that logic, which could leave a numbered native OMath produced
            // by a previous cross-format conversion mathematically active around
            // the insertion boundary.
            anchor = DeleteExact(
                document,
                host);

            Release(exact);
            exact = null;
            Release(math);
            math = null;
            Release(maths);
            maths = null;
            Release(range);
            range = null;

            Release(paragraphRange);
            paragraphRange = null;
            Release(paragraph);
            paragraph = null;
            Release(paragraphs);
            paragraphs = null;

            var probePosition = Math.Max(
                document.Content.Start,
                Math.Min(anchor.Start, Math.Max(document.Content.Start, document.Content.End - 1)));
            var probe = document.Range(
                probePosition,
                Math.Min(document.Content.End, probePosition + 1));
            try
            {
                paragraphs = probe.Paragraphs;
                if (paragraphs.Count != 1)
                    throw new InvalidDataException(
                        "The display OMML paragraph disappeared while creating the external replacement boundary.");
                paragraph = paragraphs[1];
                refreshedParagraphRange = paragraph.Range.Duplicate;
            }
            finally { Release(probe); }

            var refreshedBodyEnd = Math.Max(
                refreshedParagraphRange.Start,
                refreshedParagraphRange.End - 1);
            body = document.Range(
                refreshedParagraphRange.Start,
                refreshedBodyEnd);
            if (!ContainsOnlyExternalDisplayScaffold(body.Text))
                throw new InvalidDataException(
                    "Clearing the display OMath left non-structural paragraph content at the external replacement boundary.");
            body.Text = string.Empty;

            Release(refreshedParagraphRange);
            refreshedParagraphRange = paragraph.Range.Duplicate;
            if (witnessSpan > 0)
            {
                var refreshedWitnessStart = refreshedParagraphRange.End;
                var refreshedWitnessEnd = Math.Min(
                    document.Content.End,
                    refreshedWitnessStart + witnessSpan);
                refreshedWitness = document.Range(
                    refreshedWitnessStart,
                    refreshedWitnessEnd);
                if (!string.Equals(
                        refreshedWitness.Text ?? string.Empty,
                        witnessText,
                        StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "Normalizing the display OMML external boundary changed following user content.");
            }

            var insertion = document.Range(
                refreshedParagraphRange.Start,
                refreshedParagraphRange.Start);
            return insertion;
        }
        finally
        {
            Release(anchor);
            Release(refreshedWitness);
            Release(body);
            Release(refreshedParagraphRange);
            Release(followingWitness);
            Release(suffix);
            Release(prefix);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(exact);
            Release(math);
            Release(maths);
            Release(range);
        }
    }

    private static bool ContainsOnlyExternalDisplayScaffold(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return true;
        foreach (var value in text!)
        {
            if (value is '\t' or '\v')
                continue;
            return false;
        }
        return true;
    }

    internal static WordFormulaHostDescriptor RemovePlainTextBoundaryGuard(
        Document document,
        WordFormulaHostDescriptor host,
        InlineShape? physicalOwner = null)
    {
        using var metric = WordOperationMetrics.Measure("FormatConversion.OmmlDelete.RemovePlainGuard");
        if (host.Kind == WordFormulaHostKind.Omml)
            throw new ArgumentException(
                "The plain-text boundary guard is only used after OMML was replaced by a non-OMML host.");

        Range? guard = null;
        Range? leading = null;
        Range? probe = null;
        try
        {
            var guardStart =
                host.Range.Start - 1;
            if (guardStart < document.Content.Start)
                throw new InvalidDataException(
                    "The cross-format replacement lost its temporary plain-text boundary.");

            guard = document.Range(
                guardStart,
                host.Range.Start);
            if (!string.Equals(
                    guard.Text,
                    "x",
                    StringComparison.Ordinal))
            {
                // A display VisualTeX OLE canonicalizes itself to the shared
                // center/right-tab host before this OMML boundary is removed.
                // Its generated leading TAB is therefore now immediately before
                // the OLE, while the temporary ordinary-text guard remains one
                // character farther left. Accept exactly that owned structure;
                // do not consume the TAB or any arbitrary user prefix.
                if (host.Kind != WordFormulaHostKind.VisualTeX
                    || !host.Display
                    || !string.Equals(
                        guard.Text,
                        "\t",
                        StringComparison.Ordinal)
                    || guardStart - 1 < document.Content.Start)
                    throw new InvalidDataException(
                        "The cross-format replacement boundary is no longer the expected temporary ordinary text character.");

                leading = guard;
                guard = null;
                guardStart--;
                guard = document.Range(
                    guardStart,
                    guardStart + 1);
                if (!string.Equals(
                        guard.Text,
                        "x",
                        StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "The tab-driven VisualTeX replacement lost its temporary ordinary text boundary.");
            }

            guard.Text =
                string.Empty;

            var expectedStart =
                host.Range.Start - 1;
            var expectedEnd =
                host.Range.End - 1;
            probe = physicalOwner is null
                ? document.Range(expectedStart, expectedEnd)
                : physicalOwner.Range;
            var refreshed = physicalOwner is null
                ? WordFormulaHostLookup.ResolveLocal(document, probe, host.Kind)
                    ?? throw new InvalidDataException(
                        "The replacement host disappeared while removing its temporary plain-text boundary.")
                : host;
            if (physicalOwner is not null)
            {
                refreshed.Range = new WordFormulaRangeAddress
                {
                    StoryType = probe.StoryType,
                    Start = probe.Start,
                    End = probe.End,
                };
            }
            if (refreshed.Range.StoryType != host.Range.StoryType
                || refreshed.Range.Start != expectedStart
                || refreshed.Range.End != expectedEnd)
                throw new InvalidDataException(
                    "Removing the temporary plain-text boundary did not shift the replacement host exactly once.");
            if (refreshed.FormulaId is null
                && !string.IsNullOrWhiteSpace(
                    host.FormulaId))
                refreshed.FormulaId =
                    host.FormulaId;
            refreshed.Numbering =
                host.Numbering;
            return refreshed;
        }
        finally
        {
            Release(probe);
            Release(leading);
            Release(guard);
        }
    }

    internal static Range DeleteExact(
        Document document,
        WordFormulaHostDescriptor host)
    {
        using var metric = WordOperationMetrics.Measure("FormatConversion.OmmlDelete.ExactOwnerAndBoundary");
        if (host.Kind != WordFormulaHostKind.Omml)
            throw new ArgumentException("OMML delete received another host kind.");

        Range? range = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? exact = null;
        Range? anchor = null;
        try
        {
            range = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            maths = range.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    "The OMML source host moved before deletion.");
            math = maths[1];
            exact = math.Range.Duplicate;
            if (!WordFormulaHostSemanticReader.SameAddress(
                    exact,
                    host.Range))
                throw new InvalidDataException(
                    "The OMML source range is no longer the exact OMath.");

            anchor = exact.Duplicate;
            anchor.Collapse(WdCollapseDirection.wdCollapseStart);

            // Range.Delete on a native OMath can make Word normalize away the
            // first whitespace character immediately following the equation.
            // Clearing the exact OMath text removes the same local host without
            // asking Word to collapse the neighboring character boundary.
            using (WordOperationMetrics.Measure("FormatConversion.OmmlDelete.ClearExactText"))
                exact.Text = string.Empty;

            Range? verify = null;
            OMaths? verifyMaths = null;
            try
            {
                var verifyStart = Math.Max(
                    document.Content.Start,
                    anchor.Start - 1);
                var verifyEnd = Math.Min(
                    document.Content.End,
                    anchor.Start + 1);
                verify = document.Range(
                    verifyStart,
                    verifyEnd);
                verifyMaths = verify.OMaths;
                if (verifyMaths.Count != 0
                    && HasEmptyMathZone(verifyMaths)
                    && TryClearMathParagraphMark(document, anchor.Start))
                {
                    Release(verifyMaths);
                    Release(verify);
                    verify = document.Range(verifyStart, verifyEnd);
                    verifyMaths = verify.OMaths;
                }
                if (verifyMaths.Count != 0)
                    throw new InvalidDataException(
                        "The exact OMML source host survived text clearing.");
            }
            finally
            {
                Release(verifyMaths);
                Release(verify);
            }

            var returned = anchor;
            anchor = null;
            return returned;
        }
        finally
        {
            Release(anchor);
            Release(exact);
            Release(math);
            Release(maths);
            Release(range);
        }
    }

    private static bool HasEmptyMathZone(OMaths maths)
    {
        for (var index = 1; index <= maths.Count; index++)
        {
            OMath? math = null;
            Range? range = null;
            try
            {
                math = maths[index];
                range = math.Range;
                if (range.Start == range.End) return true;
            }
            finally
            {
                Release(range);
                Release(math);
            }
        }
        return false;
    }

    // A display paragraph's mark may itself carry the math property
    // (w:pPr/w:rPr/w:oMath). macOS VisualTeX writes it, and so does Word for
    // equations typed in their own paragraph. Clearing the equation text then
    // leaves an empty math zone on that mark, which Word still reports as an
    // OMath. Rewrite only that mark without the property; its paragraph
    // formatting travels in the same XML.
    private static bool TryClearMathParagraphMark(
        Document document,
        int position)
    {
        Range? content = null;
        Range? mark = null;
        try
        {
            content = document.Content;
            if (position < content.Start || position + 1 > content.End) return false;
            mark = document.Range(position, position + 1);
            if (!string.Equals(mark.Text, "\r", StringComparison.Ordinal)) return false;
            var package = XDocument.Parse(mark.WordOpenXML, LoadOptions.PreserveWhitespace);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var flags = package
                .Descendants(w + "pPr")
                .Elements(w + "rPr")
                .Elements(w + "oMath")
                .ToArray();
            if (flags.Length == 0) return false;
            foreach (var flag in flags) flag.Remove();
            mark.InsertXML(
                (package.Declaration?.ToString() ?? string.Empty)
                + package.ToString(SaveOptions.DisableFormatting));
            return true;
        }
        finally
        {
            Release(mark);
            Release(content);
        }
    }

    private static WordFormulaHostDescriptor DescribeInsertedOmml(
        Range inserted,
        bool requestedNumbered)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? exact = null;
        try
        {
            maths = inserted.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    "Word did not materialize exactly one OMML host.");
            math = maths[1];
            exact = math.Range.Duplicate;
            if (exact.End <= exact.Start
                || exact.Start != inserted.Start
                || exact.End != inserted.End
                || exact.StoryType != inserted.StoryType)
                throw new InvalidDataException(
                    "Word returned an invalid or ambiguous OMML insertion range.");

            var address =
                new WordFormulaRangeAddress
                {
                    StoryType = exact.StoryType,
                    Start = exact.Start,
                    End = exact.End,
                };
            return new WordFormulaHostDescriptor
            {
                Kind = WordFormulaHostKind.Omml,
                Range = address,
                DisplayMode =
                    math.Type == WdOMathType.wdOMathDisplay
                        ? "block"
                        : "inline",
                FormulaId = null,
                Metadata = null,
                MetadataAuthoritative = false,
                Numbering = new WordFormulaNumberingDescriptor
                {
                    Numbered = requestedNumbered,
                    FormulaId = null,
                    ContainerKind = requestedNumbered
                        ? WordFormulaNumberingContainerKind
                            .CanonicalNativeOmml
                        : WordFormulaNumberingContainerKind.None,
                    ContainerRange = requestedNumbered
                        ? address
                        : null,
                },
                WithinTable = IsWithinTable(exact),
            };
        }
        finally
        {
            Release(exact);
            Release(math);
            Release(maths);
        }
    }

    private static bool IsWithinTable(Range range) =>
        WordFormulaHostLookup.IsWhollyInsideTable(range);

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
