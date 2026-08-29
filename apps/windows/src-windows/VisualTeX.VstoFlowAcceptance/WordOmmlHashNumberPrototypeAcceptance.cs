using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private enum HashNumberFieldKind
    {
        Reference,
        DirectSequence,
    }

    private sealed class HashNumberPrototypeResult
    {
        internal HashNumberPrototypeResult(string name, bool success, string detail)
        {
            Name = name;
            Success = success;
            Detail = detail;
        }

        internal string Name { get; }
        internal bool Success { get; }
        internal string Detail { get; }
    }

    private static void RunWordOmmlHashNumberPrototypeAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            var cases = CreateNumberedOmmlDisplayCases();
            var referenceResult = RunHashNumberPrototypeVariant(
                application,
                artifactRoot,
                HashNumberFieldKind.Reference,
                cases[1]);
            var directSequenceResult = RunHashNumberPrototypeVariant(
                application,
                artifactRoot,
                HashNumberFieldKind.DirectSequence,
                cases[cases.Count - 1]);
            var dynamicSequenceResult = RunDirectSequenceDynamicPrototype(
                application,
                artifactRoot,
                cases[0],
                cases[cases.Count - 1]);

            Console.WriteLine(
                $"Hash-number prototype REF result: success={referenceResult.Success}; {referenceResult.Detail}");
            Console.WriteLine(
                $"Hash-number prototype direct-SEQ result: success={directSequenceResult.Success}; {directSequenceResult.Detail}");
            Console.WriteLine(
                $"Hash-number prototype dynamic direct-SEQ result: success={dynamicSequenceResult.Success}; {dynamicSequenceResult.Detail}");

            if (!directSequenceResult.Success || !dynamicSequenceResult.Success)
                throw new InvalidOperationException(
                    "OMML #(SEQ ...) did not satisfy the independent Word feasibility prototype. "
                    + $"single={directSequenceResult.Detail}; dynamic={dynamicSequenceResult.Detail}; REF={referenceResult.Detail}");

            Console.WriteLine(
                "Word OMML #(...) numbering prototype completed without changing production numbering code.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static HashNumberPrototypeResult RunHashNumberPrototypeVariant(
        Word.Application application,
        string artifactRoot,
        HashNumberFieldKind kind,
        NumberedOmmlDisplayCase testCase)
    {
        var variantName = kind == HashNumberFieldKind.Reference
            ? "ref"
            : "direct-seq";
        var documentPath = Path.Combine(
            artifactRoot,
            $"word-omml-hash-number-{variantName}.docx");
        var formulaBookmarkName = kind == HashNumberFieldKind.Reference
            ? "VTHashRefFormula"
            : "VTHashSeqFormula";
        var sourceSequenceName = kind == HashNumberFieldKind.Reference
            ? "VTHashRefSeq"
            : "VTHashDirectSeq";
        var formulaId = Guid.NewGuid().ToString("D");
        var targetBookmarkName = "VTEqNum_" + Guid.Parse(formulaId).ToString("N");

        Word.Document? document = null;
        Word.Field? sourceSequenceField = null;
        Word.Range? taggedRange = null;
        Word.Bookmark? formulaBookmark = null;
        try
        {
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();

            if (kind == HashNumberFieldKind.Reference)
            {
                sourceSequenceField = AppendHiddenSequenceTarget(
                    document,
                    sourceSequenceName,
                    targetBookmarkName,
                    restartAt: 1);
            }

            taggedRange = InsertHashTaggedDisplayFormula(
                application,
                document,
                testCase.MathMl,
                kind,
                sourceSequenceName,
                targetBookmarkName,
                formulaBookmarkName);
            AppendBodyReference(document, targetBookmarkName);

            AssertHashNumberPrototypeHost(
                document,
                formulaBookmarkName,
                targetBookmarkName,
                kind,
                expectedResult: "1",
                context: $"{variantName} initial");

            if (kind == HashNumberFieldKind.Reference)
            {
                if (sourceSequenceField is null)
                    throw new InvalidOperationException("The REF prototype lost its source SEQ field.");
                SetFieldInstruction(
                    sourceSequenceField,
                    $" SEQ {sourceSequenceName} \\r 7 \\* ARABIC ");
                sourceSequenceField.Update();
                UpdateHashTagAndBodyFields(document, formulaBookmarkName, targetBookmarkName);
                AssertHashNumberPrototypeHost(
                    document,
                    formulaBookmarkName,
                    targetBookmarkName,
                    kind,
                    expectedResult: "7",
                    context: $"{variantName} after F9/source change");
            }
            else
            {
                // Production numbering never needs to rewrite the mathematical
                // field code in place. Word stores field instructions inside OMath
                // as mathematical runs, so changing Field.Code.Text can re-tokenize
                // switches such as \\* and corrupt an otherwise valid SEQ. The
                // relevant operation is F9/Field.Update on the unchanged SEQ code.
                UpdateHashTagAndBodyFields(document, formulaBookmarkName, targetBookmarkName);
                AssertHashNumberPrototypeHost(
                    document,
                    formulaBookmarkName,
                    targetBookmarkName,
                    kind,
                    expectedResult: "1",
                    context: $"{variantName} after F9");
            }

            SetHashTaggedFormulaFontSize(document, formulaBookmarkName, 12f);
            AssertHashNumberPrototypeHost(
                document,
                formulaBookmarkName,
                targetBookmarkName,
                kind,
                expectedResult: kind == HashNumberFieldKind.Reference ? "7" : "1",
                context: $"{variantName} at 12pt");
            SetHashTaggedFormulaFontSize(document, formulaBookmarkName, 16f);
            AssertHashNumberPrototypeHost(
                document,
                formulaBookmarkName,
                targetBookmarkName,
                kind,
                expectedResult: kind == HashNumberFieldKind.Reference ? "7" : "1",
                context: $"{variantName} at 16pt");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            Release(taggedRange);
            taggedRange = null;
            Release(formulaBookmark);
            formulaBookmark = null;
            Release(sourceSequenceField);
            sourceSequenceField = null;

            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            UpdateHashTagAndBodyFields(document, formulaBookmarkName, targetBookmarkName);
            AssertHashNumberPrototypeHost(
                document,
                formulaBookmarkName,
                targetBookmarkName,
                kind,
                expectedResult: kind == HashNumberFieldKind.Reference ? "7" : "1",
                context: $"{variantName} save/reopen");

            return new HashNumberPrototypeResult(
                variantName,
                true,
                "genuine wdOMathDisplay survived; the #() field stayed live; the formula paragraph ended immediately after the OMath; external REF stayed dynamic; 12→16pt and save/reopen were stable");
        }
        catch (Exception error)
        {
            Console.WriteLine(
                $"  Hash-number prototype {variantName} FAILED: {error.GetType().Name}: {error.Message}");
            return new HashNumberPrototypeResult(
                variantName,
                false,
                $"{error.GetType().Name}: {error.Message}");
        }
        finally
        {
            Release(formulaBookmark);
            Release(taggedRange);
            Release(sourceSequenceField);
            if (document is not null)
            {
                try { document.Save(); } catch { }
                try { document.Close(Word.WdSaveOptions.wdSaveChanges); } catch { }
            }
            Release(document);
            ForceComCleanup();
        }
    }

    private static HashNumberPrototypeResult RunDirectSequenceDynamicPrototype(
        Word.Application application,
        string artifactRoot,
        NumberedOmmlDisplayCase firstCase,
        NumberedOmmlDisplayCase secondCase)
    {
        const string sequenceName = "VTHashDynamicSeq";
        const string firstFormulaBookmark = "VTHashDynFormulaA";
        const string secondFormulaBookmark = "VTHashDynFormulaB";
        var firstFormulaId = Guid.NewGuid().ToString("D");
        var secondFormulaId = Guid.NewGuid().ToString("D");
        var firstTarget = "VTEqNum_" + Guid.Parse(firstFormulaId).ToString("N");
        var secondTarget = "VTEqNum_" + Guid.Parse(secondFormulaId).ToString("N");
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-hash-number-direct-seq-dynamic.docx");

        Word.Document? document = null;
        Word.Range? firstRange = null;
        Word.Range? secondRange = null;
        try
        {
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();

            firstRange = InsertHashTaggedDisplayFormula(
                application,
                document,
                firstCase.MathMl,
                HashNumberFieldKind.DirectSequence,
                sequenceName,
                firstTarget,
                firstFormulaBookmark);
            AppendBodyReference(document, firstTarget);
            secondRange = InsertHashTaggedDisplayFormula(
                application,
                document,
                secondCase.MathMl,
                HashNumberFieldKind.DirectSequence,
                sequenceName,
                secondTarget,
                secondFormulaBookmark);
            AppendBodyReference(document, secondTarget);

            // F9 semantics: keep both field codes unchanged and update the two
            // mathematical SEQ fields in document order. The second field must see
            // the first one and advance to 2; ordinary body REF fields must read the
            // bookmarks hosted inside the mathematical number slots.
            UpdateHashTagAndBodyFields(document, firstFormulaBookmark, firstTarget);
            UpdateHashTagAndBodyFields(document, secondFormulaBookmark, secondTarget);
            AssertHashNumberPrototypeHost(
                document,
                firstFormulaBookmark,
                firstTarget,
                HashNumberFieldKind.DirectSequence,
                expectedResult: "1",
                context: "dynamic direct-seq first after F9");
            AssertHashNumberPrototypeHost(
                document,
                secondFormulaBookmark,
                secondTarget,
                HashNumberFieldKind.DirectSequence,
                expectedResult: "2",
                context: "dynamic direct-seq second after F9");

            SetHashTaggedFormulaFontSize(document, firstFormulaBookmark, 12f);
            SetHashTaggedFormulaFontSize(document, secondFormulaBookmark, 16f);
            UpdateHashTagAndBodyFields(document, firstFormulaBookmark, firstTarget);
            UpdateHashTagAndBodyFields(document, secondFormulaBookmark, secondTarget);
            AssertHashNumberPrototypeHost(
                document,
                firstFormulaBookmark,
                firstTarget,
                HashNumberFieldKind.DirectSequence,
                expectedResult: "1",
                context: "dynamic direct-seq first 12pt");
            AssertHashNumberPrototypeHost(
                document,
                secondFormulaBookmark,
                secondTarget,
                HashNumberFieldKind.DirectSequence,
                expectedResult: "2",
                context: "dynamic direct-seq second 16pt");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            Release(firstRange);
            firstRange = null;
            Release(secondRange);
            secondRange = null;

            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            UpdateHashTagAndBodyFields(document, firstFormulaBookmark, firstTarget);
            UpdateHashTagAndBodyFields(document, secondFormulaBookmark, secondTarget);
            AssertHashNumberPrototypeHost(
                document,
                firstFormulaBookmark,
                firstTarget,
                HashNumberFieldKind.DirectSequence,
                expectedResult: "1",
                context: "dynamic direct-seq first save/reopen");
            AssertHashNumberPrototypeHost(
                document,
                secondFormulaBookmark,
                secondTarget,
                HashNumberFieldKind.DirectSequence,
                expectedResult: "2",
                context: "dynamic direct-seq second save/reopen");

            return new HashNumberPrototypeResult(
                "dynamic-direct-seq",
                true,
                "two unchanged #(SEQ) fields renumbered as 1/2 under F9; each internal VTEqNum bookmark remained externally REF-addressable; mixed 12/16pt and save/reopen stayed wdOMathDisplay");
        }
        catch (Exception error)
        {
            Console.WriteLine(
                $"  Hash-number prototype dynamic direct-SEQ FAILED: {error.GetType().Name}: {error.Message}");
            return new HashNumberPrototypeResult(
                "dynamic-direct-seq",
                false,
                $"{error.GetType().Name}: {error.Message}");
        }
        finally
        {
            Release(secondRange);
            Release(firstRange);
            if (document is not null)
            {
                try { document.Save(); } catch { }
                try { document.Close(Word.WdSaveOptions.wdSaveChanges); } catch { }
            }
            Release(document);
            ForceComCleanup();
        }
    }

    private static Word.Field AppendHiddenSequenceTarget(
        Word.Document document,
        string sequenceName,
        string bookmarkName,
        int restartAt)
    {
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Range? insertion = null;
        Word.Field? field = null;
        Word.Range? resultRange = null;
        Word.Bookmark? bookmark = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            paragraph = document.Paragraphs.Add();
            paragraphRange = paragraph.Range;
            insertion = document.Range(paragraphRange.Start, paragraphRange.Start);
            field = document.Fields.Add(
                insertion,
                Word.WdFieldType.wdFieldEmpty,
                $"SEQ {sequenceName} \\r {restartAt} \\* ARABIC",
                PreserveFormatting: true);
            field.Update();
            resultRange = field.Result;
            bookmark = document.Bookmarks.Add(bookmarkName, resultRange);
            font = paragraphRange.Font;
            font.Hidden = -1;
            var returned = field;
            field = null;
            return returned;
        }
        finally
        {
            Release(font);
            Release(bookmark);
            Release(resultRange);
            Release(field);
            Release(insertion);
            Release(paragraphRange);
            Release(paragraph);
        }
    }

    private static Word.Range InsertHashTaggedDisplayFormula(
        Word.Application application,
        Word.Document document,
        string mathMl,
        HashNumberFieldKind kind,
        string sequenceName,
        string targetBookmarkName,
        string formulaBookmarkName)
    {
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Range? insertion = null;
        Word.Range? baseRange = null;
        Word.Range? taggedRange = null;
        Word.Bookmark? bookmark = null;
        try
        {
            paragraph = document.Paragraphs.Add();
            paragraphRange = paragraph.Range;
            insertion = document.Range(paragraphRange.Start, paragraphRange.Start);
            baseRange = WordOmmlConverter.Insert(
                application,
                document,
                insertion,
                mathMl,
                display: true,
                out _,
                includeLeadingTab: false,
                replaceTarget: false,
                mathFontName: document.OMathFontName);
            var semanticOmml = WordOmmlConverter.ExtractSingleOMath(baseRange.WordOpenXML);
            var taggedOmml = BuildHashTaggedOmml(
                semanticOmml,
                kind,
                sequenceName,
                targetBookmarkName);
            taggedRange = WordOmmlConverter.ReplaceWithPreparedOmml(
                application,
                document,
                baseRange,
                taggedOmml,
                display: true,
                mathFontName: document.OMathFontName);
            bookmark = document.Bookmarks.Add(formulaBookmarkName, taggedRange);
            var returned = taggedRange;
            taggedRange = null;
            return returned;
        }
        finally
        {
            Release(bookmark);
            Release(taggedRange);
            Release(baseRange);
            Release(insertion);
            Release(paragraphRange);
            Release(paragraph);
        }
    }

    private static string BuildHashTaggedOmml(
        string semanticOmml,
        HashNumberFieldKind kind,
        string sequenceName,
        string targetBookmarkName)
    {
        const string MathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        const string WordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var math = (XNamespace)MathNamespace;
        var word = (XNamespace)WordNamespace;
        var equation = XElement.Parse(
            WordOmmlConverter.ExtractSingleOMath(semanticOmml),
            LoadOptions.PreserveWhitespace);
        var formulaNodes = equation
            .Elements()
            .Select(element => new XElement(element))
            .Cast<object>()
            .ToList();
        if (formulaNodes.Count == 0)
            throw new InvalidDataException("The #() prototype requires a nonempty OMath body.");

        XElement FieldRun(XElement content) =>
            new(
                math + "r",
                new XElement(math + "rPr", new XElement(math + "nor")),
                new XElement(word + "rPr", new XElement(word + "noProof")),
                content);

        var fieldElements = new List<object>();
        if (kind == HashNumberFieldKind.DirectSequence)
        {
            fieldElements.Add(new XElement(
                word + "bookmarkStart",
                new XAttribute(word + "id", "31001"),
                new XAttribute(word + "name", targetBookmarkName)));
        }
        fieldElements.Add(FieldRun(new XElement(
            word + "fldChar",
            new XAttribute(word + "fldCharType", "begin"),
            new XAttribute(word + "dirty", "true"))));
        fieldElements.Add(FieldRun(new XElement(
            word + "instrText",
            new XAttribute(XNamespace.Xml + "space", "preserve"),
            kind == HashNumberFieldKind.Reference
                ? $" REF {targetBookmarkName} \\h \\* CHARFORMAT "
                : $" SEQ {sequenceName} \\* ARABIC ")));
        fieldElements.Add(FieldRun(new XElement(
            word + "fldChar",
            new XAttribute(word + "fldCharType", "separate"))));
        fieldElements.Add(FieldRun(new XElement(math + "t", "1")));
        fieldElements.Add(FieldRun(new XElement(
            word + "fldChar",
            new XAttribute(word + "fldCharType", "end"))));
        if (kind == HashNumberFieldKind.DirectSequence)
        {
            fieldElements.Add(new XElement(
                word + "bookmarkEnd",
                new XAttribute(word + "id", "31001")));
        }

        var delimiter = new XElement(
            math + "d",
            new XElement(math + "e", fieldElements));
        var equationBody = new XElement(math + "e", formulaNodes);
        equationBody.Add(
            new XElement(math + "r", new XElement(math + "t", "#")),
            delimiter);
        return new XElement(
                math + "oMath",
                new XElement(
                    math + "eqArr",
                    new XElement(
                        math + "eqArrPr",
                        new XElement(
                            math + "maxDist",
                            new XAttribute(math + "val", "1"))),
                    equationBody))
            .ToString(SaveOptions.DisableFormatting);
    }

    private static void AppendBodyReference(
        Word.Document document,
        string targetBookmarkName)
    {
        Word.Paragraph? paragraph = null;
        Word.Range? prefixRange = null;
        Word.Range? fieldInsertion = null;
        Word.Field? field = null;
        try
        {
            paragraph = document.Paragraphs.Add();
            var paragraphStart = paragraph.Range.Start;
            const string prefix = "VisualTeX body REF: ";
            prefixRange = document.Range(paragraphStart, paragraphStart);
            prefixRange.Text = prefix;
            fieldInsertion = document.Range(
                paragraphStart + prefix.Length,
                paragraphStart + prefix.Length);
            field = document.Fields.Add(
                fieldInsertion,
                Word.WdFieldType.wdFieldEmpty,
                $"REF {targetBookmarkName} \\h \\* CHARFORMAT",
                PreserveFormatting: true);
            field.Update();
        }
        finally
        {
            Release(field);
            Release(fieldInsertion);
            Release(prefixRange);
            Release(paragraph);
        }
    }

    private static void SetFieldInstruction(Word.Field field, string instruction)
    {
        Word.Range? code = null;
        try
        {
            code = field.Code;
            code.Text = instruction;
        }
        finally { Release(code); }
    }

    private static void SetHashTagFieldInstruction(
        Word.Document document,
        string formulaBookmarkName,
        string instruction)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        try
        {
            bookmark = document.Bookmarks[formulaBookmarkName];
            range = bookmark.Range;
            maths = range.OMaths;
            if (maths.Count != 1)
                throw new InvalidOperationException("The #() prototype formula bookmark no longer contains one OMath.");
            math = maths[1];
            mathRange = math.Range;
            fields = mathRange.Fields;
            if (fields.Count != 1)
                throw new InvalidOperationException("The #() prototype no longer contains exactly one live field.");
            field = fields[1];
            SetFieldInstruction(field, instruction);
            field.Update();
        }
        finally
        {
            Release(field);
            Release(fields);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(range);
            Release(bookmark);
        }
    }

    private static void UpdateHashTagAndBodyFields(
        Word.Document document,
        string formulaBookmarkName,
        string targetBookmarkName)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? formulaRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Fields? mathFields = null;
        try
        {
            bookmark = document.Bookmarks[formulaBookmarkName];
            formulaRange = bookmark.Range;
            maths = formulaRange.OMaths;
            if (maths.Count == 1)
            {
                math = maths[1];
                mathRange = math.Range;
                mathFields = mathRange.Fields;
                mathFields.Update();
            }
            document.Fields.Update();
            var body = FindBodyReferenceField(document, targetBookmarkName, formulaRange);
            if (body is not null)
            {
                try { body.Update(); }
                finally { Release(body); }
            }
        }
        finally
        {
            Release(mathFields);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(formulaRange);
            Release(bookmark);
        }
    }

    private static void SetHashTaggedFormulaFontSize(
        Word.Document document,
        string formulaBookmarkName,
        float size)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            bookmark = document.Bookmarks[formulaBookmarkName];
            range = bookmark.Range;
            maths = range.OMaths;
            if (maths.Count != 1)
                throw new InvalidOperationException("The #() prototype lost its OMath before font-size verification.");
            math = maths[1];
            mathRange = math.Range;
            font = mathRange.Font;
            font.Size = size;
        }
        finally
        {
            Release(font);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(range);
            Release(bookmark);
        }
    }

    private static void AssertHashNumberPrototypeHost(
        Word.Document document,
        string formulaBookmarkName,
        string targetBookmarkName,
        HashNumberFieldKind kind,
        string expectedResult,
        string context)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? formulaRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Fields? mathFields = null;
        Word.Field? labelField = null;
        Word.Range? labelCode = null;
        Word.Range? labelResult = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? ownerRange = null;
        Word.Field? bodyReference = null;
        try
        {
            AssertEqual(0, document.Tables.Count, context + ": a Word table appeared.");
            AssertEqual(0, document.Shapes.Count, context + ": a floating Shape appeared.");
            AssertTrue(document.Bookmarks.Exists(formulaBookmarkName),
                context + ": the formula bookmark is missing.");
            AssertTrue(document.Bookmarks.Exists(targetBookmarkName),
                context + ": the VisualTeX number target bookmark is missing.");

            bookmark = document.Bookmarks[formulaBookmarkName];
            formulaRange = bookmark.Range;
            maths = formulaRange.OMaths;
            AssertEqual(1, maths.Count, context + ": the formula bookmark does not contain one OMath.");
            math = maths[1];
            AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                context + ": the #() formula degraded from wdOMathDisplay.");
            mathRange = math.Range;
            mathFields = mathRange.Fields;
            AssertEqual(1, mathFields.Count,
                context + ": the #() label is no longer a live Word field.");
            labelField = mathFields[1];
            labelCode = labelField.Code;
            labelResult = labelField.Result;
            var instruction = labelCode.Text ?? string.Empty;
            if (kind == HashNumberFieldKind.Reference)
                AssertTrue(instruction.IndexOf("REF " + targetBookmarkName, StringComparison.OrdinalIgnoreCase) >= 0,
                    context + ": the #() field is not the expected REF field.");
            else
                AssertTrue(instruction.IndexOf("SEQ ", StringComparison.OrdinalIgnoreCase) >= 0,
                    context + ": the #() field is not the expected SEQ field.");
            AssertEqual(expectedResult, (labelResult.Text ?? string.Empty).Trim(),
                context + ": the #() field result is stale.");

            paragraphs = mathRange.Paragraphs;
            AssertEqual(1, paragraphs.Count, context + ": the #() formula spans multiple paragraphs.");
            paragraph = paragraphs[1];
            ownerRange = paragraph.Range;
            AssertEqual(mathRange.End + 1, ownerRange.End,
                context + ": ordinary content exists after the OMath; the paragraph mark is not immediately after the numbered formula.");
            AssertTrue((ownerRange.Text ?? string.Empty).EndsWith("\r", StringComparison.Ordinal),
                context + ": the numbered formula paragraph has no normal paragraph mark.");
            AssertTrue((ownerRange.Text ?? string.Empty).IndexOf('\t') < 0,
                context + ": the #() prototype unexpectedly contains a paragraph TAB.");

            const string WordNamespace =
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            const string MathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var word = (XNamespace)WordNamespace;
            var mathNs = (XNamespace)MathNamespace;
            var ownerXml = XDocument.Parse(
                ownerRange.WordOpenXML ?? string.Empty,
                LoadOptions.PreserveWhitespace);
            AssertEqual(1, ownerXml.Descendants(mathNs + "oMathPara").Count(),
                context + ": Word did not retain one m:oMathPara.");
            AssertEqual(1, ownerXml.Descendants(mathNs + "eqArr").Count(),
                context + ": Word did not retain the #() equation-array host.");
            AssertTrue(ownerXml.Descendants(mathNs + "t").Any(node => node.Value == "#"),
                context + ": the # token disappeared from the math structure.");
            var paragraphXml = ownerXml.Descendants(word + "p").FirstOrDefault()
                ?? throw new InvalidDataException(context + ": owner paragraph XML is missing.");
            AssertTrue(!paragraphXml.Elements(word + "r").Any(),
                context + ": ordinary visible Word runs exist outside m:oMathPara.");

            bodyReference = FindBodyReferenceField(document, targetBookmarkName, formulaRange);
            if (bodyReference is null)
                throw new InvalidOperationException(context + ": the external VisualTeX REF field is missing.");
            Word.Range? bodyResult = null;
            Word.Range? bodyCode = null;
            try
            {
                bodyCode = bodyReference.Code;
                bodyResult = bodyReference.Result;
                AssertTrue((bodyCode.Text ?? string.Empty).IndexOf("\\h", StringComparison.OrdinalIgnoreCase) >= 0,
                    context + ": the external VisualTeX REF lost its hyperlink switch.");
                AssertEqual(expectedResult, (bodyResult.Text ?? string.Empty).Trim(),
                    context + ": the external VisualTeX REF result is stale.");
            }
            finally
            {
                Release(bodyCode);
                Release(bodyResult);
            }

            Console.WriteLine(
                $"  {context}: type={math.Type}, formula={mathRange.Start}:{mathRange.End}, owner={ownerRange.Start}:{ownerRange.End}, field='{instruction.Trim()}', result='{(labelResult.Text ?? string.Empty).Trim()}', tables={document.Tables.Count}, shapes={document.Shapes.Count}.");
        }
        finally
        {
            Release(bodyReference);
            Release(ownerRange);
            Release(paragraph);
            Release(paragraphs);
            Release(labelResult);
            Release(labelCode);
            Release(labelField);
            Release(mathFields);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(formulaRange);
            Release(bookmark);
        }
    }

    private static Word.Field? FindBodyReferenceField(
        Word.Document document,
        string targetBookmarkName,
        Word.Range? formulaRange)
    {
        Word.Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? field = null;
                Word.Range? code = null;
                Word.Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    var instruction = code.Text ?? string.Empty;
                    if (instruction.IndexOf(
                            "REF " + targetBookmarkName,
                            StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    result = field.Result;
                    if (formulaRange is not null
                        && result.Start >= formulaRange.Start
                        && result.End <= formulaRange.End)
                        continue;
                    var returned = field;
                    field = null;
                    return returned;
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
            return null;
        }
        finally { Release(fields); }
    }
}
