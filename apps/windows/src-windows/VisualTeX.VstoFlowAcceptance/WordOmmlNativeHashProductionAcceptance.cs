using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlNativeHashProductionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-native-hash-production.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Field? navigableReference = null;
        var originalFormulaIds = new List<string>();
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();

            // Use an explicit outline-number prefix so the acceptance is independent
            // of the machine's localized multilevel-list gallery. VisualTeX's heading
            // resolver treats this as chapter 2 and the mathematical SEQ supplies the
            // local 1/2/3 ordinal.
            document.Content.Text =
                "2 Production chapter\rProduction body references: \r";
            Word.Paragraph? heading = null;
            Word.Paragraph? referenceParagraph = null;
            try
            {
                heading = document.Paragraphs[1];
                heading.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel1;
                referenceParagraph = document.Paragraphs[2];
                referenceParagraph.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
            }
            finally
            {
                Release(referenceParagraph);
                Release(heading);
            }
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);

            var service = new WordFormulaService(application);
            for (var index = 1; index <= 3; index++)
            {
                originalFormulaIds.Add(InsertNativeHashProductionFormula(
                    application,
                    document,
                    service,
                    document.Content.End - 1,
                    @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}+" + index,
                    QuadraticFormulaMathMl()));
            }

            UpdateNativeHashProductionFields(document, originalFormulaIds);
            AssertNativeHashProductionNumbers(
                document,
                originalFormulaIds,
                new[] { "2.1", "2.2", "2.3" },
                "initial chapter numbering");
            var bodyReferenceTargets = InsertNativeHashProductionReferences(
                document,
                originalFormulaIds);
            AssertNativeHashProductionReferences(
                document,
                bodyReferenceTargets,
                originalFormulaIds,
                new[] { "2.1", "2.2", "2.3" },
                "initial body REF");
            navigableReference = InsertNavigableEquationReference(
                application,
                document,
                originalFormulaIds[0]);
            AssertNavigableEquationReference(
                application,
                document,
                navigableReference,
                originalFormulaIds[0],
                "2.1",
                "initial product GOTOBUTTON + nested REF");
            Console.WriteLine("  native-hash stage passed: initial chapter 2.1/2.2/2.3 + body REF + navigable reference");

            // A format switch must atomically rebuild the mathematical wrapper when
            // its prefix/SEQ switches change. It must never mutate Field.Code.Text
            // inside OMath. Existing body REF fields prove that VTEqNum identities
            // survive the replacement.
            AssertEqual(
                3,
                WordEquationNumbering.SetEquationNumberFormat(
                    document,
                    EquationNumberFormat.ContinuousId),
                "Continuous format switch did not update all three numbered OMML formulas.");
            UpdateNativeHashProductionFields(document, originalFormulaIds);
            AssertNativeHashProductionNumbers(
                document,
                originalFormulaIds,
                new[] { "1", "2", "3" },
                "continuous numbering after format switch");
            AssertNativeHashProductionReferences(
                document,
                bodyReferenceTargets,
                originalFormulaIds,
                new[] { "1", "2", "3" },
                "body REF after continuous format switch");
            AssertNavigableEquationReference(
                application,
                document,
                navigableReference!,
                originalFormulaIds[0],
                "1",
                "navigable REF after continuous format switch");
            Console.WriteLine("  native-hash stage passed: atomic switch to continuous 1/2/3");

            // Insert a fourth formula immediately before the old second formula.
            // Word's SEQ engine must reflow 1,2,3,4 in document order without any
            // mathematical field-code rewrite.
            var oldSecondOwner = WordEquationNumbering.FindNumberingOwnerRange(
                    document,
                    originalFormulaIds[1])
                ?? throw new InvalidDataException(
                    "The old second numbered OMML has no owner before middle insertion.");
            var middleInsertionStart = oldSecondOwner.Start;
            Word.Paragraphs? oldSecondParagraphs = null;
            Word.Paragraph? oldSecondParagraph = null;
            Word.Range? oldSecondParagraphRange = null;
            try
            {
                oldSecondParagraphs = oldSecondOwner.Paragraphs;
                oldSecondParagraph = oldSecondParagraphs[1];
                oldSecondParagraphRange = oldSecondParagraph.Range;
                oldSecondParagraphRange.InsertParagraphBefore();
            }
            finally
            {
                Release(oldSecondParagraphRange);
                Release(oldSecondParagraph);
                Release(oldSecondParagraphs);
                Release(oldSecondOwner);
            }
            var insertedFormulaId = InsertNativeHashProductionFormula(
                application,
                document,
                service,
                middleInsertionStart,
                @"y=\sum_{k=1}^{n}k",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
                + "<mrow><mi>y</mi><mo>=</mo><munderover><mo>∑</mo><mrow><mi>k</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></munderover><mi>k</mi></mrow></math>");
            var insertedOrder = new[]
            {
                originalFormulaIds[0],
                insertedFormulaId,
                originalFormulaIds[1],
                originalFormulaIds[2],
            };
            UpdateNativeHashProductionFields(document, insertedOrder);
            AssertNativeHashProductionNumbers(
                document,
                insertedOrder,
                new[] { "1", "2", "3", "4" },
                "middle insertion renumbering");
            AssertNativeHashProductionReferences(
                document,
                bodyReferenceTargets,
                originalFormulaIds,
                new[] { "1", "3", "4" },
                "body REF after middle insertion");
            AssertNavigableEquationReference(
                application,
                document,
                navigableReference!,
                originalFormulaIds[0],
                "1",
                "navigable REF after middle insertion");
            Console.WriteLine("  native-hash stage passed: middle insertion -> 1/2/3/4");

            DeleteNativeHashProductionFormula(service, document, insertedFormulaId);
            UpdateNativeHashProductionFields(document, originalFormulaIds);
            AssertNativeHashProductionNumbers(
                document,
                originalFormulaIds,
                new[] { "1", "2", "3" },
                "middle deletion renumbering");
            AssertNativeHashProductionReferences(
                document,
                bodyReferenceTargets,
                originalFormulaIds,
                new[] { "1", "2", "3" },
                "body REF after middle deletion");
            AssertNavigableEquationReference(
                application,
                document,
                navigableReference!,
                originalFormulaIds[0],
                "1",
                "navigable REF after middle deletion");
            Console.WriteLine("  native-hash stage passed: middle deletion -> 1/2/3");

            AssertEqual(
                3,
                WordEquationNumbering.SetEquationNumberFormat(
                    document,
                    EquationNumberFormat.Heading1DotId),
                "Heading format restoration did not update all three numbered OMML formulas.");
            UpdateNativeHashProductionFields(document, originalFormulaIds);
            AssertNativeHashProductionNumbers(
                document,
                originalFormulaIds,
                new[] { "2.1", "2.2", "2.3" },
                "chapter numbering restored");
            AssertNativeHashProductionReferences(
                document,
                bodyReferenceTargets,
                originalFormulaIds,
                new[] { "2.1", "2.2", "2.3" },
                "body REF after chapter format restoration");
            AssertNavigableEquationReference(
                application,
                document,
                navigableReference!,
                originalFormulaIds[0],
                "2.1",
                "navigable REF after chapter format restoration");
            Console.WriteLine("  native-hash stage passed: atomic restoration to chapter 2.1/2.2/2.3");

            Release(navigableReference);
            navigableReference = null;
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            ForceComCleanup();

            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            UpdateNativeHashProductionFields(document, originalFormulaIds);
            AssertNativeHashProductionNumbers(
                document,
                originalFormulaIds,
                new[] { "2.1", "2.2", "2.3" },
                "save/reopen chapter numbering");
            AssertNativeHashProductionReferences(
                document,
                bodyReferenceTargets,
                originalFormulaIds,
                new[] { "2.1", "2.2", "2.3" },
                "save/reopen body REF");
            navigableReference = FindNavigableEquationReference(
                    document,
                    originalFormulaIds[0])
                ?? throw new InvalidDataException(
                    "Save/reopen lost the product GOTOBUTTON + nested REF reference.");
            AssertNavigableEquationReference(
                application,
                document,
                navigableReference,
                originalFormulaIds[0],
                "2.1",
                "save/reopen navigable REF");
            Console.WriteLine("  native-hash stage passed: save/reopen + F9-equivalent refresh + navigable reference");

            Console.WriteLine(
                "Production OMML native #(SEQ) acceptance passed: chapter 2.1/2.2/2.3, continuous 1/2/3, middle insertion/deletion, external REF, F9-equivalent updates, and save/reopen all remained table/Shape-free wdOMathDisplay.");
        }
        finally
        {
            Release(navigableReference);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static string InsertNativeHashProductionFormula(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        int insertionPosition,
        string latex,
        string mathMl)
    {
        Word.Range? insertion = null;
        try
        {
            var content = document.Content;
            try
            {
                insertionPosition = Math.Max(
                    content.Start,
                    Math.Min(insertionPosition, content.End - 1));
            }
            finally { Release(content); }
            insertion = document.Range(insertionPosition, insertionPosition);
            application.Selection.SetRange(insertion.Start, insertion.End);
            var formulaId = Guid.NewGuid().ToString("D");
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertion.Start,
                insertion.End,
                latex,
                originalMetadata: null);
            service.InsertOmml(session, mathMl);
            return formulaId;
        }
        finally { Release(insertion); }
    }

    private static void UpdateNativeHashProductionFields(
        Word.Document document,
        IReadOnlyList<string> formulaIds)
    {
        foreach (var formulaId in formulaIds
                     .OrderBy(id =>
                     {
                         Word.Range? owner = null;
                         try
                         {
                             owner = WordEquationNumbering.FindNumberingOwnerRange(document, id);
                             return owner?.Start ?? int.MaxValue;
                         }
                         finally { Release(owner); }
                     }))
        {
            Word.Bookmark? bookmark = null;
            Word.Range? range = null;
            Word.Fields? fields = null;
            try
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                    ?? throw new InvalidDataException(
                        $"Numbered OMML {formulaId} lost its VTOMML identity during field update.");
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                    ?? throw new InvalidDataException(
                        $"Numbered OMML {formulaId} lost metadata during field update.");
                range = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    formulaId,
                    metadata);
                fields = range.Fields;
                AssertEqual(1, fields.Count,
                    $"Numbered OMML {formulaId} no longer contains exactly one mathematical SEQ field.");
                fields.Update();
            }
            finally
            {
                Release(fields);
                Release(range);
                Release(bookmark);
            }
        }
        WordEquationNumbering.UpdateEquationNumbers(document);
        Word.Fields? bodyFields = null;
        try
        {
            bodyFields = document.Fields;
            if (bodyFields.Count > 0) bodyFields.Update();
        }
        finally { Release(bodyFields); }
    }

    private static void AssertNativeHashProductionNumbers(
        Word.Document document,
        IReadOnlyList<string> formulaIds,
        IReadOnlyList<string> expectedNumbers,
        string context)
    {
        AssertEqual(formulaIds.Count, expectedNumbers.Count,
            context + ": formula/expectation count mismatch.");
        AssertEqual(0, document.Tables.Count,
            context + ": a Word table exists in the production numbered-OMML document.");
        AssertEqual(0, document.Shapes.Count,
            context + ": a floating Shape/TextBox exists in the production numbered-OMML document.");
        for (var index = 0; index < formulaIds.Count; index++)
        {
            var formulaId = formulaIds[index];
            var expected = expectedNumbers[index];
            Word.Bookmark? formulaBookmark = null;
            Word.Range? formulaRange = null;
            Word.OMaths? maths = null;
            Word.OMath? math = null;
            Word.Range? mathRange = null;
            Word.Fields? fields = null;
            Word.Field? field = null;
            Word.Range? code = null;
            Word.Bookmarks? bookmarks = null;
            Word.Bookmark? numberBookmark = null;
            Word.Range? numberRange = null;
            Word.Range? ownerRange = null;
            try
            {
                formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                    ?? throw new InvalidDataException(context + $": VTOMML identity missing for {formulaId}.");
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                    ?? throw new InvalidDataException(context + $": metadata missing for {formulaId}.");
                formulaRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    formulaId,
                    metadata);
                maths = formulaRange.OMaths;
                AssertEqual(1, maths.Count,
                    context + $": {formulaId} does not contain exactly one OMath.");
                math = maths[1];
                AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                    context + $": {formulaId} degraded from wdOMathDisplay.");
                mathRange = math.Range;
                fields = mathRange.Fields;
                AssertEqual(1, fields.Count,
                    context + $": {formulaId} does not contain exactly one SEQ field.");
                field = fields[1];
                code = field.Code;
                var instruction = code.Text ?? string.Empty;
                AssertTrue(
                    WordEquationNumbering.IsVisualTeXSequenceFieldCode(instruction),
                    context + $": {formulaId} field is not SEQ VisualTeXEquation: '{instruction}'.");
                AssertTrue(
                    instruction.IndexOf("REF ", StringComparison.OrdinalIgnoreCase) < 0,
                    context + $": {formulaId} incorrectly contains REF inside #().");

                bookmarks = document.Bookmarks;
                var numberName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
                AssertTrue(bookmarks.Exists(numberName),
                    context + $": {formulaId} lost {numberName}.");
                numberBookmark = bookmarks[numberName];
                numberRange = numberBookmark.Range;
                AssertEqual(expected, NormalizeNativeHashProductionNumber(numberRange.Text),
                    context + $": {formulaId} SEQ bookmark result mismatch.");
                var numberInsideMath =
                    numberRange.Start >= mathRange.Start
                    && numberRange.End <= mathRange.End;
                if (!numberInsideMath)
                {
                    Console.WriteLine(
                        $"  native-hash-range-diagnostic context='{context}' index={index + 1}/{formulaIds.Count} id={formulaId} expected={expected} formula={formulaRange.Start}:{formulaRange.End} math={mathRange.Start}:{mathRange.End} number={numberRange.Start}:{numberRange.End} numberText='{NormalizeNativeHashProductionNumber(numberRange.Text)}'");
                    Word.Range? diagnosticContent = null;
                    try
                    {
                        diagnosticContent = document.Content;
                        var documentXml = diagnosticContent.WordOpenXML ?? string.Empty;
                        var normalizedId = Guid.Parse(formulaId).ToString("N");
                        foreach (var alias in new[]
                                 {
                                     "VTEq_" + normalizedId,
                                     "VTEqCap_" + normalizedId,
                                     "VTEqNum_" + normalizedId,
                                 })
                        {
                            var starts = System.Text.RegularExpressions.Regex.Matches(
                                documentXml,
                                $@"<w:bookmarkStart\b(?=[^>]*\bw:id=""(?<id>-?\d+)"")(?=[^>]*\bw:name=""{System.Text.RegularExpressions.Regex.Escape(alias)}"")[^>]*/?>",
                                System.Text.RegularExpressions.RegexOptions.IgnoreCase
                                | System.Text.RegularExpressions.RegexOptions.CultureInvariant);
                            var ids = starts.Cast<System.Text.RegularExpressions.Match>()
                                .Select(match => match.Groups["id"].Value)
                                .ToArray();
                            var pairing = ids.Select(bookmarkId =>
                            {
                                var sameStarts = System.Text.RegularExpressions.Regex.Matches(
                                    documentXml,
                                    $@"<w:bookmarkStart\b(?=[^>]*\bw:id=""{System.Text.RegularExpressions.Regex.Escape(bookmarkId)}"")[^>]*/?>",
                                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                                    | System.Text.RegularExpressions.RegexOptions.CultureInvariant).Count;
                                var sameEnds = System.Text.RegularExpressions.Regex.Matches(
                                    documentXml,
                                    $@"<w:bookmarkEnd\b[^>]*\bw:id=""{System.Text.RegularExpressions.Regex.Escape(bookmarkId)}""[^>]*/?>",
                                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                                    | System.Text.RegularExpressions.RegexOptions.CultureInvariant).Count;
                                return $"{bookmarkId}:starts={sameStarts},ends={sameEnds}";
                            });
                            Console.WriteLine(
                                $"    alias-id-diagnostic alias={alias} ids=[{string.Join(",", ids)}] pairing=[{string.Join(";", pairing)}]");
                        }
                    }
                    finally { Release(diagnosticContent); }
                }
                AssertTrue(
                    numberInsideMath,
                    context + $": formula #{index + 1} {formulaId} VTEqNum bookmark is outside OMath.");

                ownerRange = WordEquationNumbering.FindNumberingOwnerRange(document, formulaId)
                    ?? throw new InvalidDataException(context + $": {formulaId} owner is missing.");
                AssertEqual(mathRange.End + 1, ownerRange.End,
                    context + $": {formulaId} does not end immediately before the normal paragraph mark.");
                var xml = XDocument.Parse(
                    ownerRange.WordOpenXML ?? string.Empty,
                    LoadOptions.PreserveWhitespace);
                XNamespace mathNamespace =
                    "http://schemas.openxmlformats.org/officeDocument/2006/math";
                XNamespace wordNamespace =
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                AssertEqual(1, xml.Descendants(mathNamespace + "oMathPara").Count(),
                    context + $": {formulaId} lost m:oMathPara.");
                AssertEqual(1, xml.Descendants(mathNamespace + "eqArr").Count(),
                    context + $": {formulaId} lost Word's native #() m:eqArr.");
                AssertTrue(xml.Descendants(mathNamespace + "t").Any(node => node.Value == "#"),
                    context + $": {formulaId} lost its native # separator.");
                AssertTrue(!xml.Descendants(wordNamespace + "tbl").Any(),
                    context + $": {formulaId} contains a table.");
                AssertTrue(!xml.Descendants(wordNamespace + "drawing").Any(),
                    context + $": {formulaId} contains a drawing/Shape.");
            }
            finally
            {
                Release(ownerRange);
                Release(numberRange);
                Release(numberBookmark);
                Release(bookmarks);
                Release(code);
                Release(field);
                Release(fields);
                Release(mathRange);
                Release(math);
                Release(maths);
                Release(formulaRange);
                Release(formulaBookmark);
            }
        }
    }

    private static IReadOnlyDictionary<string, string> InsertNativeHashProductionReferences(
        Word.Document document,
        IReadOnlyList<string> formulaIds)
    {
        var targets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var formulaId in formulaIds)
        {
            Word.Range? insertion = null;
            Word.Fields? fields = null;
            Word.Field? field = null;
            Word.Range? fieldCodeRange = null;
            Word.OMaths? fieldCodeMaths = null;
            try
            {
                // Use the product's own post-display typing-boundary helper. A
                // collapsed Range derived directly from Content.End or an arbitrary
                // paragraph can still carry OMath affinity after a display equation,
                // especially in long documents. This helper either reuses a proven
                // ordinary paragraph or inserts one immediately after this exact
                // FormulaId without touching the mathematical #(SEQ) host.
                insertion = WordEquationNumbering
                    .EnsureNormalTypingParagraphAfterNumberedDisplay(
                        document,
                        formulaId)
                    ?? throw new InvalidDataException(
                        "The production reference body paragraph could not be created safely.");
                insertion.InsertAfter("  ");
                insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                fields = insertion.Fields;
                object fieldType = Word.WdFieldType.wdFieldRef;
                var target = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
                object fieldCode = target + " \\h";
                object preserveFormatting = true;
                field = fields.Add(insertion, ref fieldType, ref fieldCode, ref preserveFormatting);
                field.Update();
                fieldCodeRange = field.Code;
                fieldCodeMaths = fieldCodeRange.OMaths;
                AssertEqual(0, fieldCodeMaths.Count,
                    "The production body-reference field code is inside OMath.");
                targets[formulaId] = target;
            }
            finally
            {
                Release(fieldCodeMaths);
                Release(fieldCodeRange);
                Release(field);
                Release(fields);
                Release(insertion);
            }
        }
        return targets;
    }

    private static void AssertNativeHashProductionReferences(
        Word.Document document,
        IReadOnlyDictionary<string, string> targets,
        IReadOnlyList<string> formulaIds,
        IReadOnlyList<string> expectedNumbers,
        string context)
    {
        AssertEqual(formulaIds.Count, expectedNumbers.Count,
            context + ": reference/expectation count mismatch.");
        for (var targetIndex = 0; targetIndex < formulaIds.Count; targetIndex++)
        {
            var formulaId = formulaIds[targetIndex];
            var targetName = targets[formulaId];
            var expected = expectedNumbers[targetIndex];
            Word.Fields? fields = null;
            Word.Field? match = null;
            try
            {
                fields = document.Fields;
                for (var index = 1; index <= fields.Count; index++)
                {
                    Word.Field? candidate = null;
                    Word.Range? code = null;
                    try
                    {
                        candidate = fields[index];
                        if (candidate.Type != Word.WdFieldType.wdFieldRef) continue;
                        code = candidate.Code;
                        if ((code.Text ?? string.Empty).IndexOf(
                                "REF " + targetName,
                                StringComparison.OrdinalIgnoreCase) < 0)
                            continue;
                        match = candidate;
                        candidate = null;
                        break;
                    }
                    finally
                    {
                        Release(code);
                        Release(candidate);
                    }
                }
                if (match is null)
                    throw new InvalidDataException(
                        context + $": body REF to {targetName} is missing.");
                match.Update();
                Word.Range? result = null;
                try
                {
                    result = match.Result;
                    AssertEqual(expected, NormalizeNativeHashProductionNumber(result.Text),
                        context + $": body REF to {targetName} is stale.");
                    Word.Range? fieldCode = null;
                    Word.OMaths? codeMaths = null;
                    try
                    {
                        fieldCode = match.Code;
                        codeMaths = fieldCode.OMaths;
                        AssertEqual(0, codeMaths.Count,
                            context + $": body REF field code to {targetName} was inserted inside OMath.");
                        AssertEqual(Word.WdStoryType.wdMainTextStory, result.StoryType,
                            context + $": body REF result to {targetName} left the main Word story.");
                    }
                    finally
                    {
                        Release(codeMaths);
                        Release(fieldCode);
                    }
                }
                finally { Release(result); }
            }
            finally
            {
                Release(match);
                Release(fields);
            }
        }
    }

    private static void AssertNativeHashProductionReferenceHyperlinkJump(
        Word.Application application,
        Word.Document document,
        string formulaId,
        string context)
    {
        Word.Fields? fields = null;
        Word.Field? reference = null;
        Word.Range? code = null;
        Word.Range? result = null;
        Word.Hyperlinks? hyperlinks = null;
        Word.Hyperlink? hyperlink = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? targetBookmark = null;
        Word.Range? targetRange = null;
        Word.Selection? selection = null;
        try
        {
            var targetName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? candidate = null;
                Word.Range? candidateCode = null;
                try
                {
                    candidate = fields[index];
                    if (candidate.Type != Word.WdFieldType.wdFieldRef) continue;
                    candidateCode = candidate.Code;
                    var instruction = candidateCode.Text ?? string.Empty;
                    if (instruction.IndexOf(
                            "REF " + targetName,
                            StringComparison.OrdinalIgnoreCase) < 0
                        || instruction.IndexOf(
                            "\\h",
                            StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    reference = candidate;
                    candidate = null;
                    code = candidateCode;
                    candidateCode = null;
                    break;
                }
                finally
                {
                    Release(candidateCode);
                    Release(candidate);
                }
            }
            if (reference is null)
                throw new InvalidDataException(context + ": hyperlink REF field is missing.");
            result = reference.Result;
            hyperlinks = result.Hyperlinks;
            AssertTrue(hyperlinks.Count > 0,
                context + ": Word did not expose REF \\h as a hyperlink.");
            hyperlink = hyperlinks[1];
            hyperlink.Follow(NewWindow: false, AddHistory: true);

            bookmarks = document.Bookmarks;
            AssertTrue(bookmarks.Exists(targetName),
                context + ": target bookmark disappeared before navigation.");
            targetBookmark = bookmarks[targetName];
            targetRange = targetBookmark.Range;
            selection = application.Selection;
            AssertTrue(
                selection.Start >= targetRange.Start
                && selection.Start <= targetRange.End,
                context + $": Word navigated to {selection.Start}, outside target {targetRange.Start}:{targetRange.End}.");
        }
        finally
        {
            Release(selection);
            Release(targetRange);
            Release(targetBookmark);
            Release(bookmarks);
            Release(hyperlink);
            Release(hyperlinks);
            Release(result);
            Release(code);
            Release(reference);
            Release(fields);
        }
    }

    private static void DeleteNativeHashProductionFormula(
        WordFormulaService service,
        Word.Document document,
        string formulaId)
    {
        var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
            ?? throw new InvalidDataException(
                $"Numbered OMML {formulaId} lost metadata before deletion acceptance.");
        Word.Range? formulaRange = null;
        try
        {
            formulaRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                formulaId,
                metadata);
            formulaRange.Select();
            AssertEqual(
                formulaId,
                service.DeleteSelectedFormula(),
                "DeleteSelectedFormula removed the wrong native #(SEQ) formula.");
        }
        finally { Release(formulaRange); }
    }

    private static string NormalizeNativeHashProductionNumber(string? value) =>
        (value ?? string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace("\a", string.Empty)
            .Replace("(", string.Empty)
            .Replace(")", string.Empty)
            .Trim();
}
