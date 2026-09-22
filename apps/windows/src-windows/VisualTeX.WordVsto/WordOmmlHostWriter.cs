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
        if (document is null || source is null)
            return false;
        if (source.Kind != WordFormulaHostKind.Omml
            || !source.Display)
            return false;

        Range? exact = null;
        Fields? fields = null;
        try
        {
            exact = WordFormulaHostSemanticReader.CreateRange(
                document,
                source.Range);
            if (!WordOmmlConverter.HasWordNativeEquationNumberHost(
                    exact.WordOpenXML ?? string.Empty))
                return false;

            fields = exact.Fields;
            return fields.Count == 0;
        }
        catch
        {
            return false;
        }
        finally
        {
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
        if (!HasFieldFreeWordNativeNumberHost(
                document,
                source))
        {
            throw new InvalidOperationException(
                "The source is not a field-free Word-native numbered OMML host.");
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
            var sourceXml = exact.WordOpenXML ?? string.Empty;
            if (!WordOmmlConverter.HasWordNativeEquationNumberHost(sourceXml))
            {
                throw new InvalidDataException(
                    "The source OMML no longer contains one Word-native equation-number host.");
            }

            var semanticOmml =
                WordOmmlConverter.TransformMathMlToOmml(request.MathMl!);
            WordOmmlConverter.ValidateOmmlResult(
                semanticOmml,
                request.MathMl!);
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

            // Typography was applied to the replacement mathematical body
            // before it was merged with the source host. Do not set Range.Font
            // on the completed OMath: that would also rewrite the field-free
            // native number payload, which VisualTeX deliberately does not own.
            WordFormulaHostLayout.ConfigureDisplayParagraph(
                replacement);

            var described = DescribeInsertedOmml(
                replacement,
                requestedNumbered: false,
                semanticOmml,
                request.AdditionalValidOmmlContentSignatures);
            return new WordFormulaHostWriteResult
            {
                Host = described.Host,
                SemanticPostconditionValidated =
                    described.SemanticValidated,
            };
        }
        finally
        {
            Release(replacement);
            Release(exact);
        }
    }

    private static string ApplySemanticOmmlTypography(
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

    internal static WordFormulaHostWriteResult Insert(
        WordApplication application,
        Document document,
        Range insertion,
        WordFormulaHostWriteRequest request)
    {
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
            var preparedOmml =
                WordOmmlConverter.TransformMathMlToOmml(request.MathMl!);
            WordOmmlConverter.ValidateOmmlResult(
                preparedOmml,
                request.MathMl!);
            var expectedSemanticOmml =
                preparedOmml;

            // Numbered OMML is one native Word host, not an unnumbered OMath
            // followed by a second rebuild. Materialize the final #(SEQ) form
            // once at the target position.
            if (display && request.Numbered)
            {
                preparedOmml =
                    WordNativeOmmlNumbering.PrepareNumberedOmml(
                        document,
                        insertion.Start,
                        preparedOmml);
            }

            if (!display)
            {
                var preparedInline =
                    PrepareAdjacentInlineNativeMerge(
                        document,
                        insertion,
                        preparedOmml,
                        request);
                preparedOmml =
                    preparedInline.Omml;
                expectedSemanticOmml =
                    preparedOmml;
                target =
                    preparedInline.Target;
            }
            else
            {
                target =
                    insertion.Duplicate;
            }

            // OMML is materialized only in the actual target paragraph. No
            // VisualTeX bookmark, scratch document, layout table or hidden
            // separator participates in the final Word structure.
            inserted = WordOmmlConverter.ReplaceWithPreparedOmmlInParagraph(
                document,
                target,
                preparedOmml,
                display,
                mathFontName:
                    document.OMathFontName);

            WordFormulaHostLayout.ApplyOmmlLocalTypography(
                inserted,
                request.FontSizePoints);
            if (display)
                WordFormulaHostLayout.ConfigureDisplayParagraph(
                    inserted);
            // OMML has no VisualTeX-owned durable identity. FormulaId is
            // operation/session state only and is never serialized into Word.

            var described = DescribeInsertedOmml(
                inserted,
                request.Numbered,
                expectedSemanticOmml,
                request.AdditionalValidOmmlContentSignatures);
            return new WordFormulaHostWriteResult
            {
                Host = described.Host,
                SemanticPostconditionValidated =
                    described.SemanticValidated,
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

    private static (Range Target, string Omml)
        PrepareAdjacentInlineNativeMerge(
            Document document,
            Range insertion,
            string preparedOmml,
            WordFormulaHostWriteRequest request)
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

            var combinedSignature =
                WordOmmlConverter.ComputeImportedOmmlContentSignature(
                    combinedOmml);
            if (!request.AdditionalValidOmmlContentSignatures.Contains(
                    combinedSignature,
                    StringComparer.Ordinal))
            {
                request.AdditionalValidOmmlContentSignatures.Add(
                    combinedSignature);
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

            anchor = exact.Duplicate;
            anchor.Collapse(WdCollapseDirection.wdCollapseStart);
            exact.Text = string.Empty;

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
        WordFormulaHostDescriptor host)
    {
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
            probe = document.Range(
                expectedStart,
                expectedEnd);
            var refreshed =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    probe,
                    host.Kind)
                ?? throw new InvalidDataException(
                    "The replacement host disappeared while removing its temporary plain-text boundary.");
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

    private static (
        WordFormulaHostDescriptor Host,
        bool SemanticValidated)
        DescribeInsertedOmml(
            Range inserted,
            bool requestedNumbered,
            string expectedSemanticOmml,
            IReadOnlyCollection<string> additionalValidSignatures)
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
            if (exact.Start != inserted.Start
                || exact.End != inserted.End
                || exact.StoryType != inserted.StoryType)
                throw new InvalidDataException(
                    "Word returned an ambiguous OMML insertion range.");

            var actualSemanticXml =
                exact.WordOpenXML
                ?? string.Empty;

            var display =
                math.Type ==
                WdOMathType.wdOMathDisplay;
            var comparison =
                WordNativeOmmlSemanticComparer.CompareOmml(
                    expectedSemanticOmml,
                    actualSemanticXml,
                    display,
                    additionalValidSignatures);
            if (!comparison.Equivalent)
            {
                throw new InvalidDataException(
                    "Word materialized an OMML equation whose canonical mathematical semantics differ from the requested content. "
                    + $"display={(display ? "block" : "inline")}; "
                    + $"expectedOmmlSignature=[{comparison.ExpectedOmmlSignature}]; "
                    + $"actualOmmlSignature=[{comparison.ActualOmmlSignature}]; "
                    + $"expectedMathMlSignature=[{comparison.ExpectedMathMlSignature}]; "
                    + $"actualMathMlSignature=[{comparison.ActualMathMlSignature}]; "
                    + $"semanticExtractionError=[{comparison.SemanticExtractionError ?? string.Empty}].");
            }
            var semanticValidated = true;

            var address =
                new WordFormulaRangeAddress
                {
                    StoryType = exact.StoryType,
                    Start = exact.Start,
                    End = exact.End,
                };
            var descriptor =
                new WordFormulaHostDescriptor
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
            return (
                descriptor,
                semanticValidated);
        }
        finally
        {
            Release(exact);
            Release(math);
            Release(maths);
        }
    }

    private static bool IsWithinTable(Range range)
    {
        try
        {
            return Convert.ToBoolean(
                range.get_Information(WdInformation.wdWithInTable));
        }
        catch { return false; }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}
