using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordHostCorePureOmmlAcceptance(
        string artifactRoot)
    {
        AssertEqual(
            "(ax-1)(x+1)>(2a+1)x-a",
            WordFormulaMutationValidator.NormalizeWordNativeLatexSemantics(
                @"\left(ax-1\right)\left(x+1\right)>\left(2a+1\right)x-a"),
            "Word-native delimiter sizing was not normalized for the reported OMML regression.");
        AssertEqual(
            @"\leftarrowx\rightarrowy",
            WordFormulaMutationValidator.NormalizeWordNativeLatexSemantics(
                @"\leftarrow x\rightarrow y"),
            "Word-native delimiter normalization corrupted arrow commands.");
        AssertEqual(
            @"M=\{x∣a-4\lex\lea+4\}",
            WordFormulaMutationValidator.NormalizeWordNativeLatexSemantics(
                @"M=\left\{x\mid a-4\le x\le a+4\right\}"),
            "Word-native relation normalization did not preserve MathType \\mid semantics.");
        AssertEqual(
            @"\{x",
            WordFormulaMutationValidator.NormalizeWordNativeLatexSemantics(
                @"\left\{x\right."),
            "Word-native invisible right delimiter was not discarded.");
        const string looseOneSidedMatrix =
            "<math><mrow><mo fence='true' stretchy='true'>{</mo>"
            + "<mtable><mtr><mtd><mi>x</mi></mtd><mtd><mi>y</mi></mtd></mtr></mtable>"
            + "<mo fence='true' stretchy='true'></mo></mrow></math>";
        const string fencedOneSidedMatrix =
            "<math><mfenced open='{' close=''><mtable>"
            + "<mtr><mtd><mi>x</mi></mtd><mtd><mi>y</mi></mtd></mtr>"
            + "</mtable></mfenced></math>";
        AssertEqual(
            MathTypeMtefCodec.SemanticSignature(looseOneSidedMatrix),
            MathTypeMtefCodec.SemanticSignature(fencedOneSidedMatrix),
            "One-sided Word mfenced semantics did not match a loose left delimiter plus matrix.");
        const string namedMax =
            "<math><msub><mi>x</mi><mo>max</mo></msub></math>";
        const string uprightMax =
            "<math><msub><mi>x</mi><mrow>"
            + "<mi mathvariant='normal'>m</mi><mi mathvariant='normal'>a</mi><mi mathvariant='normal'>x</mi>"
            + "</mrow></msub></math>";
        AssertEqual(
            MathTypeMtefCodec.SemanticSignature(namedMax),
            MathTypeMtefCodec.SemanticSignature(uprightMax),
            "Named max operator did not match Word's upright-letter normalization.");

        const string editedNativeMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>y</mi><mo>+</mo><mn>2</mn></mrow></math>";
        const string displayMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></mrow></math>";

        var path = Path.Combine(
            artifactRoot,
            "VisualTeX-Host-Core-Pure-OMML.docx");
        TryDeleteAcceptanceFile(path);

        using (var host = new WordPerformanceHost(documentPath: null))
        {
            var application = host.Application;
            var document = host.Document;
            var service = new WordFormulaService(application);

            // 1. A Word-native OMath that was never inserted by VisualTeX must be
            // directly editable. No ownership/adoption metadata is allowed.
            document.Content.Text = "Lx+1R";
            Word.Range? source = null;
            Word.Range? added = null;
            Word.OMaths? nativeMaths = null;
            Word.OMath? nativeMath = null;
            Word.Range? nativeRange = null;
            try
            {
                source = document.Range(1, 4);
                added = document.OMaths.Add(source);
                nativeMaths = added.OMaths;
                AssertEqual(
                    1,
                    nativeMaths.Count,
                    "Word did not create exactly one native OMath.");
                nativeMath = nativeMaths[1];
                nativeMath.Type = Word.WdOMathType.wdOMathInline;
                nativeMath.BuildUp();
                nativeRange = nativeMath.Range.Duplicate;
                nativeRange.Select();

                var selected = service.ReadSelection();
                AssertTrue(
                    selected.Metadata is not null
                    && selected.Metadata.Lines.Count == 1
                    && Guid.TryParseExact(
                        selected.Metadata.Lines[0].Id,
                        "D",
                        out _),
                    "Plain OMML edit metadata did not expose a canonical UUID formula-line id.");
                AssertEqual(
                    FormulaOleContract.WordOmmlMode,
                    selected.ObjectMode ?? string.Empty,
                    "A plain Word OMath was not exposed to VisualTeX as OMML.");
                AssertTrue(
                    Guid.TryParse(selected.FormulaId, out _),
                    "A plain Word OMath did not receive a session-local editor FormulaId.");
                AssertTrue(
                    !string.IsNullOrWhiteSpace(selected.ObjectId),
                    "A plain Word OMath did not expose its exact captured range.");

                var editSession = CreateOmmlMathTypeAcceptanceSession(
                    editedNativeMathMl,
                    "inline",
                    numbered: false,
                    FormulaOleContract.WordOmmlMode);
                editSession.Mode = "edit";
                editSession.FormulaId = selected.FormulaId!;
                editSession.SourceDocumentId = selected.DocumentId;
                editSession.SourceObjectId = selected.ObjectId;
                editSession.OriginalMetadata = selected.Metadata;
                _ = service.ReplaceOmml(
                    editSession,
                    editedNativeMathMl);

                Release(nativeRange);
                nativeRange = null;
                Release(nativeMath);
                nativeMath = null;
                Release(nativeMaths);
                nativeMaths = null;
                Release(added);
                added = null;
                Release(source);
                source = null;

                var maths = document.OMaths;
                try
                {
                    AssertEqual(
                        1,
                        maths.Count,
                        "Editing one plain Word OMath changed the OMath count.");
                    nativeMath = maths[1];
                    nativeRange = nativeMath.Range.Duplicate;
                    var resolved = WordFormulaHostResolver.ResolveLocal(
                            document,
                            nativeRange,
                            WordFormulaHostKind.Omml)
                        ?? throw new InvalidDataException(
                            "Edited native Word OMath cannot be resolved locally.");
                    AssertTrue(
                        string.IsNullOrWhiteSpace(resolved.FormulaId),
                        "Edited pure OMML acquired a durable VisualTeX FormulaId.");
                    var payload = WordFormulaHostSemanticReader.Read(
                        document,
                        resolved);
                    AssertEqual(
                        MathTypeMtefCodec.SemanticSignature(editedNativeMathMl),
                        MathTypeMtefCodec.SemanticSignature(
                            payload.MathMl
                            ?? throw new InvalidDataException(
                                "Edited pure OMML returned no MathML.")),
                        "Editing a native Word OMath changed its mathematical semantics.");
                }
                finally
                {
                    Release(maths);
                }

                AssertNoVisualTeXOmmlArtifacts(
                    document,
                    "after editing a plain Word-native OMath");
            }
            finally
            {
                Release(nativeRange);
                Release(nativeMath);
                Release(nativeMaths);
                Release(added);
                Release(source);
            }

            // 2. Unnumbered display OMML inserted through VisualTeX must still be
            // nothing more than Word's own display OMath.
            document.Content.Text = string.Empty;
            application.Selection.SetRange(0, 0);
            var unnumberedSession = CreateOmmlMathTypeAcceptanceSession(
                displayMathMl,
                "block",
                numbered: false,
                FormulaOleContract.WordOmmlMode);
            _ = service.InsertOmml(
                unnumberedSession,
                displayMathMl);

            var unnumberedIndex =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            AssertEqual(
                1,
                unnumberedIndex.Omml.Count,
                "Unnumbered display insertion did not leave one Word OMath.");
            var unnumbered = unnumberedIndex.Omml.Single();
            AssertEqual(
                "block",
                unnumbered.DisplayMode,
                "Unnumbered display insertion did not remain wdOMathDisplay.");
            unnumbered.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    unnumbered);
            AssertTrue(
                !unnumbered.Numbering.Numbered
                && unnumbered.Numbering.ContainerKind ==
                    WordFormulaNumberingContainerKind.None,
                "Unnumbered display OMML unexpectedly acquired a numbering container.");
            AssertNoVisualTeXOmmlArtifacts(
                document,
                "after unnumbered display OMML insertion");

            // 3. Numbered OMML is exactly Word-native #(SEQ Equation) structure.
            // No VTOMML/VTEqNum/VisualTeXEquation namespace is allowed.
            document.Content.Text = string.Empty;
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            application.Selection.SetRange(0, 0);
            var numberedSession = CreateOmmlMathTypeAcceptanceSession(
                displayMathMl,
                "block",
                numbered: true,
                FormulaOleContract.WordOmmlMode);
            _ = service.InsertOmml(
                numberedSession,
                displayMathMl);

            var numberedIndex =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            AssertEqual(
                1,
                numberedIndex.Omml.Count,
                "Numbered display insertion did not leave one Word OMath.");
            var numbered = numberedIndex.Omml.Single();
            numbered.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    numbered);
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                numbered.Numbering.ContainerKind,
                "Numbered OMML is not Word-native #(SEQ).");

            Word.Range? numberedRange = null;
            Word.Fields? numberFields = null;
            Word.Field? numberField = null;
            Word.Range? numberCode = null;
            try
            {
                numberedRange =
                    WordFormulaHostSemanticReader.CreateRange(
                        document,
                        numbered.Range);
                numberFields = numberedRange.Fields;
                var sequenceCount = 0;
                for (var index = 1;
                     index <= numberFields.Count;
                     index++)
                {
                    Release(numberCode);
                    numberCode = null;
                    Release(numberField);
                    numberField = numberFields[index];
                    numberCode = numberField.Code.Duplicate;
                    var code = numberCode.Text ?? string.Empty;
                    if (code.TrimStart().StartsWith(
                            "SEQ ",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        sequenceCount++;
                        AssertTrue(
                            code.IndexOf(
                                "VisualTeXEquation",
                                StringComparison.OrdinalIgnoreCase) < 0,
                            "Pure OMML numbering still uses the VisualTeXEquation SEQ namespace.");
                    }
                }
                AssertEqual(
                    1,
                    sequenceCount,
                    "Pure numbered OMML does not contain exactly one Word SEQ field.");
            }
            finally
            {
                Release(numberCode);
                Release(numberField);
                Release(numberFields);
                Release(numberedRange);
            }

            AssertEqual(
                0,
                document.Tables.Count,
                "Pure numbered OMML created a layout table.");
            AssertEqual(
                0,
                document.Shapes.Count,
                "Pure numbered OMML created a Shape.");
            AssertNoVisualTeXOmmlArtifacts(
                document,
                "after native numbered OMML insertion");

            // 4. Word itself must recognize the native equation number as an
            // Equation cross-reference target. Insert the reference using Word's
            // own API rather than a VisualTeX bookmark.
            var nativeItems =
                document.GetCrossReferenceItems(
                    Word.WdCaptionLabelID.wdCaptionEquation)
                as Array;
            AssertTrue(
                nativeItems is not null
                && nativeItems.Length >= 1,
                "Word did not expose the native #(SEQ Equation) OMML in its Equation cross-reference list.");

            var serviceTargets =
                service.GetCanonicalEquationReferenceTargets(
                    document)
                    .Where(target =>
                        target.Source ==
                            EquationReferenceSource.WordOmml)
                    .ToArray();
            AssertEqual(
                1,
                serviceTargets.Length,
                "VisualTeX did not expose exactly one pure OMML native-reference target.");
            AssertTrue(
                string.IsNullOrWhiteSpace(
                    serviceTargets[0].FormulaId),
                "Pure OMML reference target unexpectedly depends on FormulaId.");
            AssertTrue(
                serviceTargets[0].NativeReferenceItem > 0,
                "Pure OMML reference target has no Word native reference item.");

            application.Selection.EndKey(
                Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            application.Selection.TypeText("ref ");
            service.InsertEquationReference(
                document,
                application.Selection,
                serviceTargets[0],
                EquationReferenceStyle.Parenthesized,
                Word.WdColor.wdColorAutomatic);

            var foundNativeRef = false;
            var fields = document.Fields;
            Word.Field? field = null;
            Word.Range? codeRange = null;
            try
            {
                for (var index = 1;
                     index <= fields.Count;
                     index++)
                {
                    Release(codeRange);
                    codeRange = null;
                    Release(field);
                    field = fields[index];
                    codeRange = field.Code.Duplicate;
                    var code = codeRange.Text ?? string.Empty;
                    if (code.TrimStart().StartsWith(
                            "REF ",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        foundNativeRef = true;
                        AssertTrue(
                            code.IndexOf(
                                "VTEqNum_",
                                StringComparison.OrdinalIgnoreCase) < 0,
                            "Word native cross-reference unexpectedly targets a VisualTeX bookmark.");
                    }
                }
            }
            finally
            {
                Release(codeRange);
                Release(field);
                Release(fields);
            }
            AssertTrue(
                foundNativeRef,
                "Word did not create a native REF field for the numbered OMML.");

            AssertNoVisualTeXOmmlArtifacts(
                document,
                "after Word-native OMML cross-reference insertion");

            host.Save(path);
            Console.WriteLine(
                "[host-core pure OMML] native edit + inline/display/numbered mapping + Word cross-reference passed.");
        }

        // 5. Persistence must not rely on any VisualTeX ownership metadata.
        using (var reopened = new WordPerformanceHost(path))
        {
            var document = reopened.Document;
            AssertTrue(
                document.OMaths.Count >= 1,
                "Save/reopen lost all Word OMath content.");
            AssertNoVisualTeXOmmlArtifacts(
                document,
                "after pure OMML save/reopen");

            var index =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            var canonicalNumbered =
                index.Omml
                    .Select(host =>
                    {
                        host.Numbering =
                            WordFormulaNumberingResolver.ResolveLocal(
                                document,
                                host);
                        return host;
                    })
                    .Where(host =>
                        host.Numbering.ContainerKind ==
                            WordFormulaNumberingContainerKind.CanonicalNativeOmml)
                    .ToArray();
            AssertEqual(
                1,
                canonicalNumbered.Length,
                "Save/reopen did not retain exactly one canonical numbered source OMath. "
                + "Additional OMath objects are allowed when Word materializes a native REF result.");

            Word.Fields? fields = null;
            Word.Field? field = null;
            Word.Range? code = null;
            Word.Range? result = null;
            Word.Bookmarks? bookmarks = null;
            Word.Bookmark? bookmark = null;
            try
            {
                fields = document.Fields;
                var foundLiveReference = false;
                string? nativeRefTarget = null;
                for (var fieldIndex = 1;
                     fieldIndex <= fields.Count;
                     fieldIndex++)
                {
                    Release(code); code = null;
                    Release(result); result = null;
                    Release(field); field = fields[fieldIndex];
                    code = field.Code.Duplicate;
                    if (!(code.Text ?? string.Empty)
                        .TrimStart()
                        .StartsWith(
                            "REF ",
                            StringComparison.OrdinalIgnoreCase))
                        continue;

                    var instruction =
                        (code.Text ?? string.Empty)
                            .Split(
                                new[] { ' ', '\t', '\r', '\n' },
                                StringSplitOptions.RemoveEmptyEntries);
                    AssertTrue(
                        instruction.Length >= 2
                        && instruction[0].Equals(
                            "REF",
                            StringComparison.OrdinalIgnoreCase)
                        && instruction[1].StartsWith(
                            "_Ref",
                            StringComparison.OrdinalIgnoreCase),
                        "Saved/reopened Word-native REF does not target Word's native _Ref bookmark.");
                    nativeRefTarget = instruction[1];

                    field.Update();
                    result = field.Result.Duplicate;
                    AssertTrue(
                        !IsMissingReferenceResult(
                            result.Text ?? string.Empty),
                        "Saved/reopened Word-native REF no longer resolves.");
                    foundLiveReference = true;
                    break;
                }
                AssertTrue(
                    foundLiveReference,
                    "Save/reopen removed the Word-native REF field.");

                bookmarks = document.Bookmarks;
                bookmarks.ShowHidden = true;
                AssertTrue(
                    !string.IsNullOrWhiteSpace(nativeRefTarget)
                    && bookmarks.Exists(nativeRefTarget),
                    "Save/reopen removed Word's native hidden _Ref cross-reference target.");
            }
            finally
            {
                Release(bookmark);
                Release(bookmarks);
                Release(result);
                Release(code);
                Release(field);
                Release(fields);
            }

            Console.WriteLine(
                "[host-core pure OMML] save/reopen kept the native OMath + _Ref/REF structure live without VisualTeX metadata.");
        }

        RunPureOmmlHeadingNumberingAcceptance();
        RunPureOmmlMissingHeadingFallbackAcceptance();
        RunPureOmmlAdjacentInlineAcceptance();
        RunReportedOmmlParenthesisNormalizationAcceptance();
        RunReportedOmmlOneSidedMatrixAcceptance();
        RunWordNativeSemanticCorpusAcceptance();
        RunPureOmmlVisualTeXComplexAdjacentInlineConversionAcceptance(
            artifactRoot);
        RunPureOmmlCrossFormatAcceptance(
            artifactRoot);
        RunPureOmmlMixedConversionMatrixAcceptance(
            artifactRoot);
        RunPureOmmlMathTypeBoundaryAcceptance(
            artifactRoot);
        RunPureOmmlCopyPasteAcceptance();
        RunPureOmmlDeleteReferenceAcceptance();
        RunPureOmmlNumberToggleAcceptance();
    }

    private static void RunPureOmmlHeadingNumberingAcceptance()
    {
        const string firstMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>a</mi><mo>=</mo><mn>1</mn></mrow></math>";
        const string secondMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>b</mi><mo>=</mo><mn>2</mn></mrow></math>";

        using var host =
            new WordPerformanceHost(
                documentPath: null);
        var application =
            host.Application;
        var document =
            host.Document;
        var service =
            new WordFormulaService(
                application);

        Word.Selection? selection = null;
        Word.ListTemplate? listTemplate = null;
        Word.ListLevel? listLevel = null;
        Word.Range? headingRange = null;
        Word.ListFormat? listFormat = null;
        try
        {
            WordEquationNumbering
                .SetEquationNumberFormatPreference(
                    document,
                    EquationNumberFormat.Heading1DotId);

            listTemplate =
                document.ListTemplates.Add(
                    OutlineNumbered: false,
                    Name:
                        "VisualTeXPureOmmlHeading"
                        + Guid.NewGuid().ToString("N"));
            listLevel =
                listTemplate.ListLevels[1];
            listLevel.NumberStyle =
                Word.WdListNumberStyle.wdListNumberStyleArabic;
            listLevel.NumberFormat =
                "%1";
            listLevel.StartAt =
                1;

            selection =
                application.Selection;
            selection.SetRange(
                0,
                0);

            void AppendHeading(
                string text,
                bool continuePreviousList)
            {
                Release(listFormat);
                listFormat = null;
                Release(headingRange);
                headingRange = null;

                var start =
                    selection.Start;
                selection.TypeText(
                    text);
                selection.TypeParagraph();
                headingRange =
                    document.Range(
                        start,
                        selection.Start);
                object headingStyle =
                    Word.WdBuiltinStyle.wdStyleHeading1;
                headingRange.set_Style(
                    ref headingStyle);
                listFormat =
                    headingRange.ListFormat;
                listFormat.ApplyListTemplateWithLevel(
                    listTemplate,
                    ContinuePreviousList:
                        continuePreviousList,
                    ApplyTo:
                        Word.WdListApplyTo.wdListApplyToWholeList,
                    DefaultListBehavior:
                        Word.WdDefaultListBehavior.wdWord10ListBehavior,
                    ApplyLevel:
                        1);
                object normalStyle =
                    Word.WdBuiltinStyle.wdStyleNormal;
                selection.set_Style(
                    ref normalStyle);
            }

            AppendHeading(
                "Chapter One",
                continuePreviousList:
                    false);
            _ = service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    firstMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode),
                firstMathMl);

            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            AppendHeading(
                "Chapter Two",
                continuePreviousList:
                    true);
            _ = service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    secondMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode),
                secondMathMl);

            var targets =
                service
                    .GetCanonicalEquationReferenceTargets(
                        document)
                    .Where(target =>
                        target.Source ==
                            EquationReferenceSource.WordOmml)
                    .OrderBy(target =>
                        target.Position)
                    .ToArray();
            AssertEqual(
                2,
                targets.Length,
                "Pure OMML heading numbering did not expose two Word-native reference targets.");
            AssertTrue(
                targets.All(target =>
                    string.IsNullOrWhiteSpace(
                        target.FormulaId)),
                "Pure OMML heading targets unexpectedly depend on FormulaId.");
            AssertEqual(
                "1.1",
                targets[0].NumberText,
                "First pure OMML heading equation did not resolve to 1.1.");
            AssertEqual(
                "2.1",
                targets[1].NumberText,
                "Second pure OMML heading equation did not restart at Heading 1.");

            var index =
                WordFormulaHostResolver
                    .CaptureDocumentIndex(
                        document);
            AssertEqual(
                2,
                index.Omml.Count,
                "Pure OMML heading numbering did not leave exactly two source OMath hosts.");

            foreach (var equationHost in
                     index.Omml)
            {
                equationHost.Numbering =
                    WordFormulaNumberingResolver
                        .ResolveLocal(
                            document,
                            equationHost);
                AssertEqual(
                    WordFormulaNumberingContainerKind.CanonicalNativeOmml,
                    equationHost.Numbering.ContainerKind,
                    "Heading-numbered pure OMML is not one native #(SEQ) host.");

                Word.Range? equationRange = null;
                Word.Fields? fields = null;
                Word.Field? field = null;
                Word.Range? code = null;
                try
                {
                    equationRange =
                        WordFormulaHostSemanticReader
                            .CreateRange(
                                document,
                                equationHost.Range);
                    fields =
                        equationRange.Fields;
                    var styleRefCount =
                        0;
                    var sequenceCount =
                        0;
                    for (var fieldIndex = 1;
                         fieldIndex <= fields.Count;
                         fieldIndex++)
                    {
                        Release(code);
                        code = null;
                        Release(field);
                        field =
                            fields[fieldIndex];
                        code =
                            field.Code.Duplicate;
                        var normalized =
                            (code.Text
                             ?? string.Empty)
                            .Trim();
                        if (normalized.StartsWith(
                                "STYLEREF 1",
                                StringComparison.OrdinalIgnoreCase))
                            styleRefCount++;
                        if (normalized.StartsWith(
                                "SEQ ",
                                StringComparison.OrdinalIgnoreCase)
                            && normalized.IndexOf(
                                "\\s 1",
                                StringComparison.OrdinalIgnoreCase)
                                >= 0)
                        {
                            sequenceCount++;
                            AssertTrue(
                                normalized.IndexOf(
                                    "VisualTeXEquation",
                                    StringComparison.OrdinalIgnoreCase)
                                    < 0,
                                "Pure OMML heading sequence still uses the VisualTeXEquation namespace.");
                        }
                    }

                    AssertEqual(
                        1,
                        styleRefCount,
                        "Heading-numbered pure OMML does not contain exactly one STYLEREF 1 field.");
                    AssertEqual(
                        1,
                        sequenceCount,
                        "Heading-numbered pure OMML does not contain exactly one native SEQ \\s 1 field.");
                }
                finally
                {
                    Release(code);
                    Release(field);
                    Release(fields);
                    Release(equationRange);
                }
            }

            AssertEqual(
                0,
                document.Tables.Count,
                "Pure OMML heading numbering created a layout table.");
            AssertEqual(
                0,
                document.Shapes.Count,
                "Pure OMML heading numbering created a Shape.");
            AssertNoVisualTeXOmmlArtifacts(
                document,
                "after pure OMML heading numbering");

            Console.WriteLine(
                "[host-core pure OMML] real Heading 1 semantics produced native 1.1 / 2.1 STYLEREF + SEQ numbering.");
        }
        finally
        {
            Release(listFormat);
            Release(headingRange);
            Release(listLevel);
            Release(listTemplate);
            Release(selection);
        }
    }

    private static void RunPureOmmlMissingHeadingFallbackAcceptance()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>x</mi><mo>=</mo><mn>1</mn></mrow></math>";

        var previousDefault =
            WordEquationNumbering.GetDefaultEquationNumberFormatId();

        using var host =
            new WordPerformanceHost(
                documentPath: null);
        var application =
            host.Application;
        var document =
            host.Document;
        var service =
            new WordFormulaService(
                application);

        Word.Range? equationRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            WordEquationNumbering
                .SetEquationNumberFormatPreference(
                    document,
                    EquationNumberFormat.Heading1DashId);

            AssertEqual(
                previousDefault,
                WordEquationNumbering.GetDefaultEquationNumberFormatId(),
                "Acceptance changed the real user default equation-number format.");

            application.Selection.SetRange(0, 0);
            _ = service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode),
                mathMl);

            var equation =
                WordFormulaHostResolver
                    .CaptureDocumentIndex(document)
                    .Omml
                    .Single();
            equationRange =
                WordFormulaHostSemanticReader
                    .CreateRange(
                        document,
                        equation.Range);
            fields =
                equationRange.Fields;

            var styleRefCount = 0;
            var sequenceCount = 0;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code.Duplicate;
                var normalized =
                    (code.Text ?? string.Empty)
                    .Trim();

                if (normalized.StartsWith(
                        "STYLEREF ",
                        StringComparison.OrdinalIgnoreCase))
                    styleRefCount++;

                if (normalized.StartsWith(
                        "SEQ ",
                        StringComparison.OrdinalIgnoreCase))
                {
                    sequenceCount++;
                    AssertTrue(
                        normalized.IndexOf(
                            "\\s 1",
                            StringComparison.OrdinalIgnoreCase) < 0,
                        "Missing-heading fallback left a heading-reset switch on the native SEQ field.");
                }
            }

            AssertEqual(
                0,
                styleRefCount,
                "Missing-heading numbered OMML emitted a STYLEREF field.");
            AssertEqual(
                1,
                sequenceCount,
                "Missing-heading numbered OMML does not contain exactly one native SEQ field.");

            var target =
                service
                    .GetCanonicalEquationReferenceTargets(
                        document)
                    .Single(item =>
                        item.Source ==
                            EquationReferenceSource.WordOmml);
            AssertEqual(
                "1",
                target.NumberText,
                "Missing-heading numbered OMML did not fall back to a valid continuous number.");

            Console.WriteLine(
                "[host-core pure OMML] missing Heading fallback emitted native continuous numbering without STYLEREF and did not mutate the user default.");
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(equationRange);
        }
    }

    private static void RunPureOmmlAdjacentInlineAcceptance()
    {
        const string firstMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></math>";
        const string secondMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>y</mi><mo>+</mo><mn>2</mn></mrow></math>";

        using var host =
            new WordPerformanceHost(
                documentPath: null);
        var application =
            host.Application;
        var document =
            host.Document;
        var service =
            new WordFormulaService(
                application);

        document.Content.Text =
            "LR";
        application.Selection.SetRange(
            1,
            1);
        _ = service.InsertOmml(
            CreateOmmlMathTypeAcceptanceSession(
                firstMathMl,
                "inline",
                numbered: false,
                FormulaOleContract.WordOmmlMode),
            firstMathMl);

        var firstIndex =
            WordFormulaHostResolver
                .CaptureDocumentIndex(
                    document);
        AssertEqual(
            1,
            firstIndex.Omml.Count,
            "First adjacent-inline insertion did not create one OMath.");
        var firstHost =
            firstIndex.Omml[0];

        application.Selection.SetRange(
            firstHost.Range.End,
            firstHost.Range.End);
        _ = service.InsertOmml(
            CreateOmmlMathTypeAcceptanceSession(
                secondMathMl,
                "inline",
                numbered: false,
                FormulaOleContract.WordOmmlMode),
            secondMathMl);

        var index =
            WordFormulaHostResolver
                .CaptureDocumentIndex(
                    document);
        AssertTrue(
            index.Omml.Count is 1 or 2,
            "Word produced an unexpected adjacent-inline OMath topology.");

        var combinedLatex =
            string.Join(
                " ",
                index.Omml.Select(item =>
                {
                    var payload =
                        WordFormulaHostSemanticReader
                            .Read(
                                document,
                                item);
                    return payload.Latex
                           ?? string.Empty;
                }));
        AssertTrue(
            combinedLatex.IndexOf(
                "x",
                StringComparison.OrdinalIgnoreCase)
                >= 0
            && combinedLatex.IndexOf(
                "1",
                StringComparison.OrdinalIgnoreCase)
                >= 0
            && combinedLatex.IndexOf(
                "y",
                StringComparison.OrdinalIgnoreCase)
                >= 0
            && combinedLatex.IndexOf(
                "2",
                StringComparison.OrdinalIgnoreCase)
                >= 0,
            "Word's adjacent-inline normalization lost formula semantics.");

        var content =
            document.Content.Text
            ?? string.Empty;
        AssertTrue(
            content.StartsWith(
                "L",
                StringComparison.Ordinal)
            && content.TrimEnd('\r')
                .EndsWith(
                    "R",
                    StringComparison.Ordinal),
            "Adjacent inline OMML insertion changed surrounding user text.");

        foreach (var equationHost in
                 index.Omml)
        {
            Word.Range? equationRange = null;
            try
            {
                equationRange =
                    WordFormulaHostSemanticReader
                        .CreateRange(
                            document,
                            equationHost.Range);
                equationRange.Select();
                var selected =
                    service.ReadSelection();
                AssertEqual(
                    FormulaOleContract.WordOmmlMode,
                    selected.ObjectMode
                    ?? string.Empty,
                    "An adjacent native OMath cannot be opened in VisualTeX.");
                AssertTrue(
                    Guid.TryParse(
                        selected.FormulaId,
                        out _),
                    "Adjacent native OMath did not receive a session-local editor ID.");
            }
            finally
            {
                Release(equationRange);
            }
        }

        AssertNoVisualTeXOmmlArtifacts(
            document,
            "after adjacent inline Word normalization");
        AssertNoHostCoreSentinelArtifacts(
            document,
            "after adjacent inline Word normalization");

        Console.WriteLine(
            $"[host-core pure OMML] adjacent inline insertion accepted Word's native topology (OMaths={index.Omml.Count}) without separators or VisualTeX metadata.");
    }

    private static void RunReportedOmmlParenthesisNormalizationAcceptance()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow>"
            + "<mo>(</mo><mi>a</mi><mi>x</mi><mo>-</mo><mn>1</mn><mo>)</mo>"
            + "<mo>(</mo><mi>x</mi><mo>+</mo><mn>1</mn><mo>)</mo>"
            + "<mo>&gt;</mo>"
            + "<mo>(</mo><mn>2</mn><mi>a</mi><mo>+</mo><mn>1</mn><mo>)</mo>"
            + "<mi>x</mi><mo>-</mo><mi>a</mi></mrow></math>";

        using var host =
            new WordPerformanceHost(
                documentPath: null);
        var application =
            host.Application;
        var document =
            host.Document;
        var service =
            new WordFormulaService(
                application);

        document.Content.Text = "L R";
        Word.Range? insertion = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? range = null;
        try
        {
            insertion =
                document.Range(
                    1,
                    1);
            insertion.Select();
            _ = service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "inline",
                    numbered: false,
                    FormulaOleContract.WordOmmlMode),
                mathMl);

            maths = document.OMaths;
            AssertEqual(
                1,
                maths.Count,
                "Reported OMML parenthesis regression did not create one native Word equation.");
            math = maths[1];
            range = math.Range.Duplicate;
            var resolved =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    range,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Reported OMML parenthesis regression could not resolve the inserted equation.");
            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    resolved);
            AssertEqual(
                "(ax-1)(x+1)>(2a+1)x-a",
                WordFormulaMutationValidator.NormalizeWordNativeLatexSemantics(
                    payload.Latex),
                "Reported OMML parenthesis regression changed mathematical content after Word BuildUp.");

            Console.WriteLine(
                "[host-core pure OMML] reported (ax-1)(x+1)>(2a+1)x-a Word BuildUp normalization passed without a false write error.");
        }
        finally
        {
            Release(range);
            Release(math);
            Release(maths);
            Release(insertion);
        }
    }

    private static void RunReportedOmmlOneSidedMatrixAcceptance()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow>"
            + "<mo fence=\"true\" stretchy=\"true\">{</mo>"
            + "<mtable>"
            + "<mtr><mtd><msub><mo>min</mo><mrow><mi mathvariant=\"bold\">x</mi><mo>∈</mo><msub><mi>Ω</mi><mi>s</mi></msub></mrow></msub></mtd>"
            + "<mtd><msub><mfenced><msub><mi mathvariant=\"bold\">σ</mi><mrow><mi>v</mi><mi>m</mi></mrow></msub></mfenced><mo>max</mo></msub></mtd></mtr>"
            + "<mtr><mtd><msub><mo>min</mo><mrow><mi mathvariant=\"bold\">x</mi><mo>∈</mo><msub><mi>Ω</mi><mi>c</mi></msub></mrow></msub></mtd>"
            + "<mtd><msup><mi mathvariant=\"bold\">F</mi><mi>T</mi></msup><mi mathvariant=\"bold\">u</mi></mtd></mtr>"
            + "<mtr><mtd><mtext>s.t.</mtext></mtd><mtd><mi>f</mi><mo>*</mo><msub><mi mathvariant=\"bold\">V</mi><mn>0</mn></msub><mo>-</mo><mi>x</mi><mo>=</mo><mn>0</mn></mtd></mtr>"
            + "<mtr><mtd></mtd><mtd><msub><mi mathvariant=\"bold\">x</mi><mi>i</mi></msub><mo>∈</mo><mfenced open=\"{\" close=\"}\"><mrow><mn>0</mn><mo>,</mo><mn>1</mn></mrow></mfenced></mtd></mtr>"
            + "</mtable><mo fence=\"true\" stretchy=\"true\"></mo></mrow></math>";

        using var host =
            new WordPerformanceHost(
                documentPath: null);
        var application = host.Application;
        var document = host.Document;
        var service = new WordFormulaService(application);
        Word.Range? insertion = null;
        Word.OMath? math = null;
        Word.Range? range = null;
        try
        {
            insertion = document.Range(0, 0);
            insertion.Select();
            _ = service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: false,
                    FormulaOleContract.WordOmmlMode),
                mathMl);
            AssertEqual(
                1,
                document.OMaths.Count,
                "Reported one-sided matrix OMML regression did not create one native Word equation.");
            math = document.OMaths[1];
            range = math.Range.Duplicate;
            var xml = range.WordOpenXML ?? string.Empty;
            AssertTrue(
                xml.IndexOf(
                    "m:begChr m:val=\"{\"",
                    StringComparison.Ordinal) >= 0
                && xml.IndexOf(
                    "m:endChr m:val=\"\"",
                    StringComparison.Ordinal) >= 0,
                "Word did not materialize the reported matrix as a one-sided native delimiter.");
            var resolved =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    range,
                    WordFormulaHostKind.Omml)
                ?? throw new InvalidDataException(
                    "Reported one-sided matrix OMML regression could not resolve the inserted equation.");
            var payload = WordFormulaHostSemanticReader.Read(document, resolved);
            AssertTrue(
                payload.Latex.IndexOf(
                    @"\begin{matrix}",
                    StringComparison.Ordinal) >= 0
                && payload.Latex.IndexOf(
                    "max",
                    StringComparison.OrdinalIgnoreCase) >= 0,
                "Reported one-sided matrix OMML regression lost matrix/max content.");
            Console.WriteLine(
                "[host-core pure OMML] reported one-sided brace + matrix + max Word BuildUp passed semantic postcondition validation.");
        }
        finally
        {
            Release(range);
            Release(math);
            Release(insertion);
        }
    }

    private static void RunWordNativeSemanticCorpusAcceptance()
    {
        var cases = new (string Name, string DisplayMode, string MathMl)[]
        {
            (
                "fraction-root",
                "inline",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><msqrt><mi>c</mi></msqrt></mfrac></math>"),
            (
                "subsup-sum",
                "inline",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msubsup><mo>∑</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mi>N</mi></msubsup><msub><mi>x</mi><mi>i</mi></msub></math>"),
            (
                "integral-limits",
                "block",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msubsup><mo>∫</mo><mn>0</mn><mo>∞</mo></msubsup><mi>f</mi><mfenced><mi>x</mi></mfenced><mi>d</mi><mi>x</mi></math>"),
            (
                "paired-matrix",
                "block",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mfenced open=\"(\" close=\")\"><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable></mfenced></math>"),
            (
                "one-sided-cases",
                "block",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mo fence=\"true\" stretchy=\"true\">{</mo><mtable><mtr><mtd><mi>x</mi></mtd><mtd><mrow><mi>x</mi><mo>&gt;</mo><mn>0</mn></mrow></mtd></mtr><mtr><mtd><mo>−</mo><mi>x</mi></mtd><mtd><mrow><mi>x</mi><mo>≤</mo><mn>0</mn></mrow></mtd></mtr></mtable><mo fence=\"true\" stretchy=\"true\"></mo></mrow></math>"),
            (
                "floor-ceiling",
                "inline",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfenced open=\"⌊\" close=\"⌋\"><mi>x</mi></mfenced><mo>+</mo><mfenced open=\"⌈\" close=\"⌉\"><mi>y</mi></mfenced></math>"),
            (
                "absolute-norm",
                "inline",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfenced open=\"|\" close=\"|\"><mi>x</mi></mfenced><mo>+</mo><mfenced open=\"‖\" close=\"‖\"><mi>v</mi></mfenced></math>"),
            (
                "named-operators",
                "block",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msub><mo>max</mo><mrow><mi>x</mi><mo>∈</mo><mi>Ω</mi></mrow></msub><mi>f</mi><mfenced><mi>x</mi></mfenced><mo>+</mo><msub><mo>min</mo><mi>y</mi></msub><mi>g</mi><mfenced><mi>y</mi></mfenced></math>"),
            (
                "accents",
                "inline",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mover accent=\"true\"><mi>x</mi><mo>→</mo></mover><mo>+</mo><mover accent=\"true\"><mi>y</mi><mo>¯</mo></mover><mo>+</mo><mover accent=\"true\"><mi>z</mi><mo>^</mo></mover></math>"),
            (
                "binomial",
                "inline",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfenced><mfrac linethickness=\"0\"><mi>n</mi><mi>k</mi></mfrac></mfenced></math>"),
            (
                "align-like-table",
                "block",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mtable displaystyle=\"true\" columnalign=\"right left\"><mtr><mtd><mi>a</mi></mtd><mtd><mo>=</mo><mi>b</mi><mo>+</mo><mi>c</mi></mtd></mtr><mtr><mtd><mi>d</mi></mtd><mtd><mo>=</mo><mi>e</mi><mo>−</mo><mi>f</mi></mtd></mtr></mtable></math>"),
            (
                "set-relations",
                "inline",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>∣</mo><mi>x</mi><mo>∈</mo><mi>A</mi><mo>∧</mo><mi>x</mi><mo>≤</mo><mi>b</mi><mo>∧</mo><mi>A</mi><mo>⊂</mo><mi>B</mi></math>"),
            (
                "nested-structures",
                "block",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mfrac><mrow><msqrt><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><mn>1</mn></mrow></msqrt></mrow><mrow><mn>1</mn><mo>+</mo><mfrac><mn>1</mn><mi>n</mi></mfrac></mrow></mfrac></math>"),
        };

        foreach (var item in cases)
        {
            using var host =
                new WordPerformanceHost(
                    documentPath: null);
            var application = host.Application;
            var document = host.Document;
            var service = new WordFormulaService(application);
            Word.Range? insertion = null;
            try
            {
                insertion = document.Range(0, 0);
                insertion.Select();
                _ = service.InsertOmml(
                    CreateOmmlMathTypeAcceptanceSession(
                        item.MathMl,
                        item.DisplayMode,
                        numbered: false,
                        FormulaOleContract.WordOmmlMode),
                    item.MathMl);
                AssertTrue(
                    document.OMaths.Count >= 1,
                    $"Word-native semantic corpus '{item.Name}' produced no OMath.");
            }
            finally
            {
                Release(insertion);
            }
        }

        const string plusOne =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mn>1</mn></math>";
        const string minusOne =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>−</mo><mn>1</mn></math>";
        var plusOmml = WordOmmlConverter.TransformMathMlToOmml(plusOne);
        var minusOmml = WordOmmlConverter.TransformMathMlToOmml(minusOne);
        var negative = WordNativeOmmlSemanticComparer.CompareOmml(
            plusOmml,
            minusOmml,
            display: false);
        AssertTrue(
            !negative.Equivalent,
            "Word-native semantic comparer became too permissive: x+1 matched x-1.");

        Console.WriteLine(
            $"[host-core pure OMML] Word-native semantic corpus passed {cases.Length} complex structures plus a strict negative-control mismatch.");
    }

    private static void RunPureOmmlVisualTeXComplexAdjacentInlineConversionAcceptance(
        string artifactRoot)
    {
        const string firstLatex =
            @"\mathrm{e}^{\mathrm{i}\pi}+1=0";
        const string firstMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<msup><mi mathvariant=\"normal\">e</mi><mrow><mi mathvariant=\"normal\">i</mi><mi>π</mi></mrow></msup>"
            + "<mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>";
        const string secondLatex =
            @"(a+b)^n=\sum_{k=0}^{n}\binom{n}{k}a^{n-k}b^k";
        const string secondMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<msup><mfenced><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow></mfenced><mi>n</mi></msup>"
            + "<mo>=</mo><msubsup><mo>∑</mo><mrow><mi>k</mi><mo>=</mo><mn>0</mn></mrow><mi>n</mi></msubsup>"
            + "<mfenced><mfrac linethickness=\"0\"><mi>n</mi><mi>k</mi></mfrac></mfenced>"
            + "<msup><mi>a</mi><mrow><mi>n</mi><mo>−</mo><mi>k</mi></mrow></msup>"
            + "<msup><mi>b</mi><mi>k</mi></msup></math>";

        var assetRoot =
            Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX",
                "office",
                "temp",
                "complex-adjacent-vt-omml-"
                + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(
            assetRoot);
        var pngPath =
            Path.Combine(
                assetRoot,
                "preview.png");
        var svgPath =
            Path.Combine(
                assetRoot,
                "preview.svg");
        WriteAcceptancePng(
            pngPath,
            "formula",
            260,
            96);
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"96\" viewBox=\"0 0 260 96\"><text x=\"8\" y=\"66\" font-family=\"Cambria Math\" font-size=\"40\">formula</text></svg>");
        var emfPath =
            OfficeOlePreview.CreateVectorEmfFromSvg(
                svgPath,
                260,
                96);

        using var host =
            new WordPerformanceHost(
                documentPath: null);
        var application =
            host.Application;
        var document =
            host.Document;
        var service =
            new WordFormulaService(
                application);

        Word.Range? insertion = null;
        Word.InlineShape? firstShape = null;
        Word.InlineShape? secondShape = null;
        try
        {
            document.Content.Text =
                "风格豆腐干豆腐干";

            void InsertVisualTeXInline(
                string latex)
            {
                Release(insertion);
                insertion =
                    document.Range(
                        document.Content.End - 1,
                        document.Content.End - 1);
                if (document.InlineShapes.Count > 0)
                {
                    Word.InlineShape? last = null;
                    Word.Range? lastRange = null;
                    try
                    {
                        last =
                            document.InlineShapes[
                                document.InlineShapes.Count];
                        lastRange =
                            last.Range.Duplicate;
                        insertion.SetRange(
                            lastRange.End,
                            lastRange.End);
                    }
                    finally
                    {
                        Release(lastRange);
                        Release(last);
                    }
                }

                insertion.Select();
                var session =
                    CreateSimpleVisualTeXSourceSession(
                        latex,
                        numbered: false);
                session.DisplayMode =
                    "inline";
                session.Numbered =
                    false;
                service.InsertOle(
                    session,
                    pngPath,
                    emfPath);
            }

            InsertVisualTeXInline(
                firstLatex);
            InsertVisualTeXInline(
                secondLatex);

            AssertEqual(
                2,
                CountVisualTeXNativeOleShapes(
                    document),
                "Complex adjacent VisualTeX setup did not create two OLE formulas.");
            firstShape =
                document.InlineShapes[1];
            secondShape =
                document.InlineShapes[2];
            AssertEqual(
                firstShape.Range.End,
                secondShape.Range.Start,
                "Complex adjacent VisualTeX setup inserted a separator between the two inline formulas.");

            var plan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.NativeOleMode,
                    FormulaOleContract.WordOmmlMode);
            AssertEqual(
                2,
                plan.Targets.Count,
                "Complex adjacent VisualTeX→OMML conversion did not capture both formulas.");

            var sourceMathMl =
                new Dictionary<string, string>(
                    StringComparer.Ordinal)
                {
                    [firstLatex] =
                        firstMathMl,
                    [secondLatex] =
                        secondMathMl,
                };
            var prepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal);
            foreach (var target in plan.Targets)
            {
                if (!sourceMathMl.TryGetValue(
                        target.Latex,
                        out var mathMl))
                    throw new InvalidDataException(
                        "Complex adjacent conversion captured an unexpected VisualTeX formula: '"
                        + target.Latex
                        + "'.");

                prepared[target.Id] =
                    new PreparedWordBulkFormula
                    {
                        Run =
                            new WordBulkRun
                            {
                                Id =
                                    target.Id,
                                IsFormula =
                                    true,
                                Latex =
                                    target.Latex,
                                DisplayMode =
                                    target.DisplayMode,
                            },
                        Session =
                            CreateSimpleFormatTargetSession(
                                target,
                                FormulaOleContract.WordOmmlMode,
                                mathMl),
                        MathMl =
                            mathMl,
                    };
            }

            var result =
                service.ApplyFormulaFormatConversionPlan(
                    plan,
                    prepared);
            AssertEqual(
                2,
                result.FormulaCount,
                "Complex adjacent VisualTeX→OMML conversion did not convert both formulas.");
            AssertEqual(
                0,
                result.FailedFormulaCount,
                "Complex adjacent VisualTeX→OMML conversion reported a failure: "
                + string.Join(
                    " | ",
                    result.Failures));
            AssertEqual(
                0,
                CountVisualTeXNativeOleShapes(
                    document),
                "Complex adjacent VisualTeX→OMML conversion left a VisualTeX OLE behind.");

            var index =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            AssertTrue(
                index.Omml.Count is 1 or 2,
                "Complex adjacent VisualTeX→OMML conversion produced an unexpected Word OMath topology.");

            var combinedLatex =
                string.Join(
                    string.Empty,
                    index.Omml.Select(item =>
                        WordFormulaHostSemanticReader
                            .Read(
                                document,
                                item)
                            .Latex));
            var normalized =
                string.Concat(
                    combinedLatex.Where(character =>
                        !char.IsWhiteSpace(
                            character)));
            AssertTrue(
                normalized.IndexOf(
                    @"\mathrm{e}^{\mathrm{i}\pi}+1=0",
                    StringComparison.Ordinal) >= 0
                && normalized.IndexOf(
                    @"\sum_{k=0}^{n}",
                    StringComparison.Ordinal) >= 0
                && normalized.IndexOf(
                    @"\binom{n}{k}",
                    StringComparison.Ordinal) >= 0,
                "Complex adjacent VisualTeX→OMML conversion lost or cross-contaminated formula semantics: '"
                + normalized
                + "'.");

            AssertTrue(
                (document.Content.Text
                    ?? string.Empty)
                    .StartsWith(
                        "风格豆腐干豆腐干",
                        StringComparison.Ordinal),
                "Complex adjacent VisualTeX→OMML conversion changed the surrounding user text.");

            Console.WriteLine(
                $"[host-core pure OMML] complex adjacent VisualTeX→OMML batch conversion passed with Word-native merged topology (OMaths={index.Omml.Count}) and retained Euler/binomial semantics.");
        }
        finally
        {
            Release(secondShape);
            Release(firstShape);
            Release(insertion);
            try
            {
                Directory.Delete(
                    assetRoot,
                    recursive: true);
            }
            catch { }
        }
    }

    private static void RunPureOmmlCrossFormatAcceptance(
        string artifactRoot)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>q</mi><mo>=</mo><mi>p</mi><mo>+</mo><mn>3</mn></mrow></math>";

        var assetRoot =
            Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX",
                "office",
                "temp",
                "pure-omml-cross-format-"
                + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(
            assetRoot);
        var pngPath =
            Path.Combine(
                assetRoot,
                "preview.png");
        var svgPath =
            Path.Combine(
                assetRoot,
                "preview.svg");
        WriteAcceptancePng(
            pngPath,
            "q=p+3",
            260,
            96);
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"96\" viewBox=\"0 0 260 96\"><rect width=\"260\" height=\"96\" fill=\"white\"/><text x=\"10\" y=\"66\" font-family=\"Cambria Math\" font-size=\"44\">q=p+3</text></svg>");
        var emfPath =
            OfficeOlePreview.CreateVectorEmfFromSvg(
                svgPath,
                260,
                96);

        using var host =
            new WordPerformanceHost(
                documentPath: null);
        var application =
            host.Application;
        var document =
            host.Document;
        var service =
            new WordFormulaService(
                application);

        document.Content.Text =
            "L R";
        application.Selection.SetRange(
            1,
            1);
        _ = service.InsertOmml(
            CreateOmmlMathTypeAcceptanceSession(
                mathMl,
                "inline",
                numbered: false,
                FormulaOleContract.WordOmmlMode),
            mathMl);

        var index =
            WordFormulaHostResolver
                .CaptureDocumentIndex(
                    document);
        AssertEqual(
            1,
            index.Omml.Count,
            "Pure cross-format fixture did not create one OMML source.");
        var source =
            index.Omml[0];
        Word.Range? sourceRange = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? visualShape = null;
        Word.Range? visualRange = null;
        Word.Range? finalOmmlRange = null;
        try
        {
            sourceRange =
                WordFormulaHostSemanticReader
                    .CreateRange(
                        document,
                        source.Range);
            sourceRange.Select();

            var toVisualPlan =
                service
                    .CaptureFormulaFormatConversionPlan(
                        wholeDocument: false,
                        FormulaOleContract.WordOmmlMode,
                        FormulaOleContract.NativeOleMode);
            AssertEqual(
                1,
                toVisualPlan.Targets.Count,
                "Pure OMML→VisualTeX did not capture exactly one formula.");
            var toVisualTarget =
                toVisualPlan.Targets.Single();
            var sourceMathMl =
                toVisualTarget.SourceMathMl
                ?? mathMl;
            var preparedVisual =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal)
                {
                    [toVisualTarget.Id] =
                        new PreparedWordBulkFormula
                        {
                            Run =
                                new WordBulkRun
                                {
                                    Id =
                                        toVisualTarget.Id,
                                    IsFormula =
                                        true,
                                    Latex =
                                        toVisualTarget.Latex,
                                    DisplayMode =
                                        toVisualTarget.DisplayMode,
                                },
                            Session =
                                CreateSimpleFormatTargetSession(
                                    toVisualTarget,
                                    FormulaOleContract.NativeOleMode,
                                    sourceMathMl),
                            MathMl =
                                sourceMathMl,
                            PngPath =
                                pngPath,
                            EmfPath =
                                emfPath,
                        },
                };

            Release(sourceRange);
            sourceRange = null;
            var toVisualResult =
                service
                    .ApplyFormulaFormatConversionPlan(
                        toVisualPlan,
                        preparedVisual);
            AssertEqual(
                1,
                toVisualResult.FormulaCount,
                "Pure OMML→VisualTeX did not convert exactly one formula.");
            AssertEqual(
                0,
                toVisualResult.FailedFormulaCount,
                "Pure OMML→VisualTeX reported a failure.");
            AssertEqual(
                0,
                document.OMaths.Count,
                "Pure OMML→VisualTeX left the source OMath alive.");
            AssertEqual(
                1,
                CountVisualTeXNativeOleShapes(
                    document),
                "Pure OMML→VisualTeX did not create one VisualTeX OLE.");

            shapes =
                document.InlineShapes;
            for (var shapeIndex = 1;
                 shapeIndex <= shapes.Count;
                 shapeIndex++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate =
                        shapes[shapeIndex];
                    if (!WordFormulaMetadataReader
                        .IsNativeOle(
                            candidate))
                        continue;
                    visualShape =
                        candidate;
                    candidate = null;
                    break;
                }
                finally
                {
                    Release(candidate);
                }
            }
            AssertTrue(
                visualShape is not null,
                "Converted VisualTeX OLE cannot be located.");
            var visualMetadata =
                WordFormulaMetadataReader
                    .TryReadEmbeddedNativeOle(
                        visualShape!)
                ?? throw new InvalidDataException(
                    "Converted VisualTeX OLE has no embedded identity.");
            AssertTrue(
                Guid.TryParse(
                    visualMetadata.FormulaId,
                    out _),
                "Converted VisualTeX OLE did not receive its own FormulaId.");

            visualRange =
                visualShape!.Range.Duplicate;
            visualRange.Select();
            var toOmmlPlan =
                service
                    .CaptureFormulaFormatConversionPlan(
                        wholeDocument: false,
                        FormulaOleContract.NativeOleMode,
                        FormulaOleContract.WordOmmlMode);
            AssertEqual(
                1,
                toOmmlPlan.Targets.Count,
                "VisualTeX→pure OMML did not capture exactly one formula.");
            var toOmmlTarget =
                toOmmlPlan.Targets.Single();
            var preparedOmml =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal)
                {
                    [toOmmlTarget.Id] =
                        new PreparedWordBulkFormula
                        {
                            Run =
                                new WordBulkRun
                                {
                                    Id =
                                        toOmmlTarget.Id,
                                    IsFormula =
                                        true,
                                    Latex =
                                        toOmmlTarget.Latex,
                                    DisplayMode =
                                        toOmmlTarget.DisplayMode,
                                },
                            Session =
                                CreateSimpleFormatTargetSession(
                                    toOmmlTarget,
                                    FormulaOleContract.WordOmmlMode,
                                    mathMl),
                            MathMl =
                                mathMl,
                        },
                };

            Release(visualRange);
            visualRange = null;
            Release(visualShape);
            visualShape = null;
            Release(shapes);
            shapes = null;

            var toOmmlResult =
                service
                    .ApplyFormulaFormatConversionPlan(
                        toOmmlPlan,
                        preparedOmml);
            AssertEqual(
                1,
                toOmmlResult.FormulaCount,
                "VisualTeX→pure OMML did not convert exactly one formula.");
            AssertEqual(
                0,
                toOmmlResult.FailedFormulaCount,
                "VisualTeX→pure OMML reported a failure.");
            AssertEqual(
                0,
                CountVisualTeXNativeOleShapes(
                    document),
                "VisualTeX→pure OMML left its OLE host alive.");

            var finalIndex =
                WordFormulaHostResolver
                    .CaptureDocumentIndex(
                        document);
            AssertEqual(
                1,
                finalIndex.Omml.Count,
                "VisualTeX→pure OMML did not leave one Word OMath.");
            var finalHost =
                finalIndex.Omml[0];
            AssertTrue(
                string.IsNullOrWhiteSpace(
                    finalHost.FormulaId),
                "VisualTeX→pure OMML leaked a durable VisualTeX FormulaId.");
            var finalPayload =
                WordFormulaHostSemanticReader
                    .Read(
                        document,
                        finalHost);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(
                    mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    finalPayload.MathMl
                    ?? throw new InvalidDataException(
                        "Round-tripped pure OMML returned no MathML.")),
                "OMML↔VisualTeX↔OMML changed formula semantics.");
            AssertNoVisualTeXOmmlArtifacts(
                document,
                "after OMML↔VisualTeX→pure OMML round-trip");

            finalOmmlRange =
                WordFormulaHostSemanticReader
                    .CreateRange(
                        document,
                        finalHost.Range);
            finalOmmlRange.Select();
            var latexResult =
                service.ConvertFormulaObjectsToLatex(
                    wholeDocument: false,
                    FormulaOleContract.WordOmmlMode);
            AssertEqual(
                1,
                latexResult.FormulaCount,
                "Pure OMML→LaTeX did not convert exactly one formula.");
            AssertEqual(
                0,
                document.OMaths.Count,
                "Pure OMML→LaTeX left an OMath behind.");
            AssertEqual(
                0,
                CountVisualTeXNativeOleShapes(
                    document),
                "Pure OMML→LaTeX introduced a VisualTeX OLE.");
            AssertNoVisualTeXOmmlArtifacts(
                document,
                "after pure OMML→LaTeX");

            var text =
                document.Content.Text
                ?? string.Empty;
            AssertTrue(
                text.IndexOf(
                    "q",
                    StringComparison.OrdinalIgnoreCase)
                    >= 0
                && text.IndexOf(
                    "p",
                    StringComparison.OrdinalIgnoreCase)
                    >= 0
                && text.IndexOf(
                    "3",
                    StringComparison.OrdinalIgnoreCase)
                    >= 0,
                "Pure OMML→LaTeX did not leave the formula source in ordinary text.");

            Console.WriteLine(
                "[host-core pure OMML] OMML→VisualTeX→OMML returned to marker-free Word state; OMML→LaTeX left plain text.");
        }
        finally
        {
            Release(finalOmmlRange);
            Release(visualRange);
            Release(visualShape);
            Release(shapes);
            Release(sourceRange);
        }
    }

    private static void RunPureOmmlMixedConversionMatrixAcceptance(
        string artifactRoot)
    {
        const string inlineMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></math>";
        const string displayMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></mrow></math>";
        const string numberedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><msup><mi>a</mi><mn>2</mn></msup><mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup></mrow></math>";
        const string secondNumberedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>F</mi><mo>=</mo><mi>m</mi><mi>a</mi></mrow></math>";

        var formulas = new[]
        {
            (MathMl: inlineMathMl, DisplayMode: "inline", Numbered: false),
            (MathMl: displayMathMl, DisplayMode: "block", Numbered: false),
            (MathMl: numberedMathMl, DisplayMode: "block", Numbered: true),
            (MathMl: secondNumberedMathMl, DisplayMode: "block", Numbered: true),
        };
        var expectedFormulaCount =
            formulas.Length;
        var expectedNumberedCount =
            formulas.Count(item => item.Numbered);
        var expectedInlineCount =
            formulas.Count(item =>
                string.Equals(
                    item.DisplayMode,
                    "inline",
                    StringComparison.OrdinalIgnoreCase));
        var expectedDisplayCount =
            formulas.Count(item =>
                string.Equals(
                    item.DisplayMode,
                    "block",
                    StringComparison.OrdinalIgnoreCase));
        var expectedSignatures = formulas
            .Select(item => MathTypeMtefCodec.SemanticSignature(item.MathMl))
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToArray();

        var assetRoot =
            Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX",
                "office",
                "temp",
                "pure-omml-mixed-conversion-"
                + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(assetRoot);
        var pngPath = Path.Combine(assetRoot, "preview.png");
        var svgPath = Path.Combine(assetRoot, "preview.svg");
        WriteAcceptancePng(pngPath, "mixed", 320, 112);
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"320\" height=\"112\" viewBox=\"0 0 320 112\"><rect width=\"320\" height=\"112\" fill=\"white\"/><text x=\"10\" y=\"74\" font-family=\"Cambria Math\" font-size=\"42\">mixed</text></svg>");
        var emfPath =
            OfficeOlePreview.CreateVectorEmfFromSvg(
                svgPath,
                320,
                112);

        var previousNativePreview =
            Environment.GetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW");

        try
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                "1");

            using var host =
                new WordPerformanceHost(
                    documentPath: null);
            var application = host.Application;
            var document = host.Document;
            var service =
                new WordFormulaService(
                    application);

            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            document.Content.Text = "L R\r";
            application.Selection.SetRange(1, 1);
            _ = service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    inlineMathMl,
                    "inline",
                    numbered: false,
                    FormulaOleContract.WordOmmlMode),
                inlineMathMl);

            AppendAcceptanceText(document, "\r");
            SelectDocumentEnd(document);
            _ = service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    displayMathMl,
                    "block",
                    numbered: false,
                    FormulaOleContract.WordOmmlMode),
                displayMathMl);

            AppendAcceptanceText(document, "\r");
            SelectDocumentEnd(document);
            _ = service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    numberedMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode),
                numberedMathMl);

            AppendAcceptanceText(document, "\r");
            SelectDocumentEnd(document);
            _ = service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    secondNumberedMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode),
                secondNumberedMathMl);

            void AssertPureOmmlState(
                string stage)
            {
                var index =
                    WordFormulaHostResolver.CaptureDocumentIndex(
                        document);
                AssertEqual(
                    expectedFormulaCount,
                    index.Omml.Count,
                    stage + " changed the Word OMath host count.");
                AssertEqual(
                    expectedInlineCount,
                    index.Omml.Count(item =>
                        string.Equals(
                            item.DisplayMode,
                            "inline",
                            StringComparison.OrdinalIgnoreCase)),
                    stage + " changed the inline OMath count.");
                AssertEqual(
                    expectedDisplayCount,
                    index.Omml.Count(item =>
                        string.Equals(
                            item.DisplayMode,
                            "block",
                            StringComparison.OrdinalIgnoreCase)),
                    stage + " changed the display OMath count.");

                var canonicalNumbered = 0;
                var actualNativeNumbers =
                    new List<string>();
                var actualSignatures =
                    new List<string>();
                foreach (var item in index.Omml
                             .OrderBy(item => item.Range.Start))
                {
                    item.Numbering =
                        WordFormulaNumberingResolver.ResolveLocal(
                            document,
                            item);
                    if (item.Numbering.ContainerKind ==
                        WordFormulaNumberingContainerKind.CanonicalNativeOmml)
                    {
                        canonicalNumbered++;
                        AssertTrue(
                            item.Numbering.NumberRange is not null,
                            stage + " lost a native equation-number range.");
                        Word.Range? numberRange = null;
                        try
                        {
                            numberRange =
                                WordFormulaHostSemanticReader.CreateRange(
                                    document,
                                    item.Numbering.NumberRange!);
                            actualNativeNumbers.Add(
                                (numberRange.Text ?? string.Empty)
                                    .Trim()
                                    .Trim('(', ')'));
                        }
                        finally
                        {
                            Release(numberRange);
                        }
                    }
                    else
                    {
                        AssertTrue(
                            !item.Numbering.Numbered,
                            stage + " introduced a non-canonical numbered OMML host.");
                    }

                    var payload =
                        WordFormulaHostSemanticReader.Read(
                            document,
                            item);
                    actualSignatures.Add(
                        MathTypeMtefCodec.SemanticSignature(
                            payload.MathMl
                            ?? throw new InvalidDataException(
                                stage + " returned an OMML host without MathML.")));
                    AssertTrue(
                        string.IsNullOrWhiteSpace(
                            item.FormulaId),
                        stage + " leaked a durable FormulaId into pure OMML.");
                }

                AssertEqual(
                    expectedNumberedCount,
                    canonicalNumbered,
                    stage + " changed the native numbered OMath count.");
                AssertEqual(
                    expectedNumberedCount,
                    actualNativeNumbers.Count,
                    stage + " returned the wrong native equation-number count.");
                for (var numberIndex = 0;
                     numberIndex < actualNativeNumbers.Count;
                     numberIndex++)
                {
                    AssertEqual(
                        (numberIndex + 1).ToString(),
                        actualNativeNumbers[numberIndex],
                        stage + $" produced the wrong native equation number at position {numberIndex + 1}.");
                }

                actualSignatures.Sort(
                    StringComparer.Ordinal);
                AssertEqual(
                    expectedSignatures.Length,
                    actualSignatures.Count,
                    stage + " changed the semantic formula count.");
                for (var indexValue = 0;
                     indexValue < expectedSignatures.Length;
                     indexValue++)
                {
                    AssertEqual(
                        expectedSignatures[indexValue],
                        actualSignatures[indexValue],
                        stage + $" changed semantic signature #{indexValue + 1}.");
                }

                AssertEqual(
                    0,
                    document.Tables.Count,
                    stage + " created a layout table.");
                AssertEqual(
                    0,
                    document.Shapes.Count,
                    stage + " created a Shape.");
                AssertNoVisualTeXOmmlArtifacts(
                    document,
                    stage);
            }

            Dictionary<string, PreparedWordBulkFormula>
                PrepareVisualTeXTargets(
                    WordFormulaFormatConversionPlan plan)
            {
                var prepared =
                    new Dictionary<string, PreparedWordBulkFormula>(
                        StringComparer.Ordinal);
                foreach (var target in plan.Targets)
                {
                    var sourceMathMl =
                        target.SourceMathMl
                        ?? throw new InvalidDataException(
                            $"VisualTeX conversion target '{target.Latex}' has no source MathML.");
                    prepared[target.Id] =
                        new PreparedWordBulkFormula
                        {
                            Run =
                                new WordBulkRun
                                {
                                    Id = target.Id,
                                    IsFormula = true,
                                    Latex = target.Latex,
                                    DisplayMode = target.DisplayMode,
                                },
                            Session =
                                CreateSimpleFormatTargetSession(
                                    target,
                                    FormulaOleContract.NativeOleMode,
                                    sourceMathMl),
                            MathMl = sourceMathMl,
                            PngPath = pngPath,
                            EmfPath = emfPath,
                        };
                }
                return prepared;
            }

            AssertPureOmmlState(
                "Initial mixed pure OMML fixture");
            var initialOrdinaryEmptyParagraphCount =
                CountPureConversionOrdinaryEmptyParagraphs(
                    document);

            var toVisualTeXPlan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.NativeOleMode);
            AssertEqual(
                expectedFormulaCount,
                toVisualTeXPlan.Targets.Count,
                "Mixed OMML→VisualTeX did not capture all formulas.");
            AssertEqual(
                expectedNumberedCount,
                toVisualTeXPlan.Targets.Count(target => target.Numbered),
                "Mixed OMML→VisualTeX capture lost numbered state.");
            var sourceMathMlByLatex =
                toVisualTeXPlan.Targets.ToDictionary(
                    target => target.Latex,
                    target => target.SourceMathMl
                        ?? throw new InvalidDataException(
                            $"Mixed OMML source '{target.Latex}' exposed no MathML before conversion."),
                    StringComparer.Ordinal);
            var toVisualTeXResult =
                service.ApplyFormulaFormatConversionPlan(
                    toVisualTeXPlan,
                    PrepareVisualTeXTargets(
                        toVisualTeXPlan));
            AssertEqual(
                expectedFormulaCount,
                toVisualTeXResult.FormulaCount,
                "Mixed OMML→VisualTeX did not convert all formulas.");
            AssertEqual(
                0,
                toVisualTeXResult.FailedFormulaCount,
                "Mixed OMML→VisualTeX reported failures: "
                + string.Join(
                    " | ",
                    toVisualTeXResult.Failures));
            AssertEqual(
                0,
                document.OMaths.Count,
                "Mixed OMML→VisualTeX left OMath sources behind.");
            AssertEqual(
                expectedFormulaCount,
                CountVisualTeXNativeOleShapes(
                    document),
                "Mixed OMML→VisualTeX created the wrong VisualTeX OLE count.");
            AssertEqual(
                expectedNumberedCount,
                CountInstalledVisualTeXNumberedFormulaHosts(
                    document),
                "Mixed OMML→VisualTeX changed the numbered formula count.");
            AssertPureConversionVisualTeXTabStructure(
                document,
                expectedNumberedCount,
                initialOrdinaryEmptyParagraphCount);

            // Real regression: a body REF created while the equation is
            // VisualTeX must remain live when the numbered host becomes pure
            // Word OMML. The old VTEqNum_* target is intentionally retired; the
            // conversion must migrate this existing REF to a Word-native _Ref*
            // bookmark instead of leaving Word's "reference source not found".
            var visualIndexForReference =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            var referencedVisualTeX =
                visualIndexForReference.VisualTeX
                    .OrderBy(item => item.Range.Start)
                    .First(item =>
                        WordFormulaNumberingResolver.ResolveLocal(
                            document,
                            item).Numbered);
            AssertTrue(
                Guid.TryParse(
                    referencedVisualTeX.FormulaId,
                    out var referencedVisualTeXId),
                "Mixed OMML→VisualTeX numbered target has no FormulaId for the reference migration fixture.");
            var oldVisualTeXReferenceTarget =
                "VTEqNum_"
                + referencedVisualTeXId.ToString("N");
            Word.Selection? migrationSelection = null;
            try
            {
                AppendAcceptanceText(
                    document,
                    "\r");
                SelectDocumentEnd(
                    document);
                migrationSelection =
                    application.Selection;
                var referenceTarget =
                    service.GetCanonicalEquationReferenceTargets(
                            document)
                        .SingleOrDefault(candidate =>
                            candidate.Source ==
                                EquationReferenceSource.VisualTeX
                            && string.Equals(
                                candidate.FormulaId,
                                referencedVisualTeXId.ToString("D"),
                                StringComparison.OrdinalIgnoreCase))
                    ?? throw new InvalidDataException(
                        "Mixed OMML→VisualTeX numbered target is missing from the product reference picker.");

                // New self-contained VisualTeX numbering creates VTEqNum lazily.
                // Use the real product reference path so this fixture exercises
                // exactly the transition users perform before converting to OMML.
                service.InsertEquationReferenceCore(
                    document,
                    migrationSelection,
                    referenceTarget,
                    EquationReferenceStyle.Parenthesized,
                    Word.WdColor.wdColorAutomatic);
            }
            finally
            {
                Release(migrationSelection);
            }
            AssertExternalReferenceTarget(
                document,
                oldVisualTeXReferenceTarget,
                expectTargetExists: true,
                context:
                    "before mixed VisualTeX→OMML reference migration");

            var visualTeXBackPlan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.NativeOleMode,
                    FormulaOleContract.WordOmmlMode);
            AssertEqual(
                expectedFormulaCount,
                visualTeXBackPlan.Targets.Count,
                "Mixed VisualTeX→OMML did not recapture all formulas.");
            var visualTeXBackPrepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal);
            foreach (var target in visualTeXBackPlan.Targets)
            {
                if (!sourceMathMlByLatex.TryGetValue(
                        target.Latex,
                        out var sourceMathMl))
                    throw new InvalidDataException(
                        $"Mixed VisualTeX source '{target.Latex}' cannot be matched to its original OMML MathML fixture.");
                visualTeXBackPrepared[target.Id] =
                    new PreparedWordBulkFormula
                    {
                        Run =
                            new WordBulkRun
                            {
                                Id = target.Id,
                                IsFormula = true,
                                Latex = target.Latex,
                                DisplayMode = target.DisplayMode,
                            },
                        Session =
                            CreateSimpleFormatTargetSession(
                                target,
                                FormulaOleContract.WordOmmlMode,
                                sourceMathMl),
                        MathMl =
                            sourceMathMl,
                    };
            }
            var visualTeXBackResult =
                service.ApplyFormulaFormatConversionPlan(
                    visualTeXBackPlan,
                    visualTeXBackPrepared);
            AssertEqual(
                expectedFormulaCount,
                visualTeXBackResult.FormulaCount,
                "Mixed VisualTeX→OMML did not convert all formulas.");
            AssertEqual(
                0,
                visualTeXBackResult.FailedFormulaCount,
                "Mixed VisualTeX→OMML reported failures: "
                + string.Join(
                    " | ",
                    visualTeXBackResult.Failures));
            AssertEqual(
                0,
                CountVisualTeXNativeOleShapes(
                    document),
                "Mixed VisualTeX→OMML left VisualTeX OLE sources behind.");
            AssertPureOmmlState(
                "After mixed OMML↔VisualTeX round-trip");
            AssertPureOmmlDisplayParagraphBoundaries(
                document,
                "After mixed VisualTeX→OMML");
            AssertMigratedPureOmmlReference(
                document,
                oldVisualTeXReferenceTarget,
                "after mixed VisualTeX→OMML reference migration");

            EnsurePureConversionUserBlankParagraph(
                document);
            var mathTypeBaselineEmptyParagraphCount =
                CountPureConversionOrdinaryEmptyParagraphs(
                    document);
            AssertTrue(
                mathTypeBaselineEmptyParagraphCount > 0,
                "OMML→MathType structure fixture did not contain a deliberate user-authored empty paragraph.");

            var toMathTypePlan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.MathTypeOleMode);
            AssertEqual(
                expectedFormulaCount,
                toMathTypePlan.Targets.Count,
                "Mixed OMML→MathType did not capture all formulas.");
            AssertEqual(
                expectedNumberedCount,
                toMathTypePlan.Targets.Count(target => target.Numbered),
                "Mixed OMML→MathType capture lost numbered state.");
            var toMathTypeResult =
                service.ApplyFormulaFormatConversionPlan(
                    toMathTypePlan,
                    PrepareOmmlMathTypeTargets(
                        toMathTypePlan,
                        emfPath));
            AssertEqual(
                expectedFormulaCount,
                toMathTypeResult.FormulaCount,
                "Mixed OMML→MathType did not convert all formulas.");
            AssertEqual(
                0,
                toMathTypeResult.FailedFormulaCount,
                "Mixed OMML→MathType reported failures: "
                + string.Join(
                    " | ",
                    toMathTypeResult.Failures));
            AssertEqual(
                0,
                document.OMaths.Count,
                "Mixed OMML→MathType left OMath sources behind.");
            AssertEqual(
                expectedFormulaCount,
                CountMathTypeOleShapes(
                    document),
                "Mixed OMML→MathType created the wrong MathType OLE count.");
            AssertPureConversionMathTypeTabStructure(
                document,
                expectedNumberedCount,
                mathTypeBaselineEmptyParagraphCount);

            var mathTypeBackPlan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.MathTypeOleMode,
                    FormulaOleContract.WordOmmlMode);
            AssertEqual(
                expectedFormulaCount,
                mathTypeBackPlan.Targets.Count,
                "Mixed MathType→OMML did not recapture all formulas.");
            var mathTypeBackResult =
                service.ApplyFormulaFormatConversionPlan(
                    mathTypeBackPlan,
                    PrepareOmmlMathTypeTargets(
                        mathTypeBackPlan,
                        emfPath));
            AssertEqual(
                expectedFormulaCount,
                mathTypeBackResult.FormulaCount,
                "Mixed MathType→OMML did not convert all formulas.");
            AssertEqual(
                0,
                mathTypeBackResult.FailedFormulaCount,
                "Mixed MathType→OMML reported failures: "
                + string.Join(
                    " | ",
                    mathTypeBackResult.Failures));
            AssertEqual(
                0,
                CountMathTypeOleShapes(
                    document),
                "Mixed MathType→OMML left MathType OLE sources behind.");
            AssertPureOmmlState(
                "After mixed OMML↔MathType round-trip");

            var ommlToLatex =
                service.ConvertFormulaObjectsToLatex(
                    wholeDocument: true,
                    FormulaOleContract.WordOmmlMode);
            AssertEqual(
                expectedFormulaCount,
                ommlToLatex.FormulaCount,
                "Mixed OMML→LaTeX did not convert all formulas.");
            AssertEqual(
                0,
                document.OMaths.Count,
                "Mixed OMML→LaTeX left native equations behind.");
            AssertEqual(
                0,
                document.InlineShapes.Count,
                "Mixed OMML→LaTeX introduced OLE objects.");
            AssertNoVisualTeXOmmlArtifacts(
                document,
                "after mixed OMML→LaTeX");

            document.Content.Text =
                string.Empty;
            application.Selection.SetRange(
                0,
                0);
            service.InsertOle(
                CreateSimpleVisualTeXSourceSession(
                    "a^2+b^2=c^2",
                    numbered: true),
                pngPath,
                emfPath);
            var visualTeXToLatex =
                service.ConvertFormulaObjectsToLatex(
                    wholeDocument: true,
                    FormulaOleContract.NativeOleMode);
            AssertEqual(
                1,
                visualTeXToLatex.FormulaCount,
                "VisualTeX→LaTeX did not convert the numbered source.");
            AssertEqual(
                0,
                CountVisualTeXNativeOleShapes(
                    document),
                "VisualTeX→LaTeX left its OLE source behind.");
            AssertEqual(
                0,
                CountVisualTeXNumberingBookmarkTriples(
                    document),
                "VisualTeX→LaTeX left a VisualTeX numbering host behind.");

            document.Content.Text =
                string.Empty;
            application.Selection.SetRange(
                0,
                0);
            _ = service.InsertMathTypeOle(
                CreateOmmlMathTypeAcceptanceSession(
                    numberedMathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.MathTypeOleMode),
                numberedMathMl,
                emfPath,
                updateCreatedMathTypeNumberFields: true);
            var mathTypeToLatex =
                service.ConvertFormulaObjectsToLatex(
                    wholeDocument: true,
                    FormulaOleContract.MathTypeOleMode);
            AssertEqual(
                1,
                mathTypeToLatex.FormulaCount,
                "MathType→LaTeX did not convert the numbered source.");
            AssertEqual(
                0,
                CountMathTypeOleShapes(
                    document),
                "MathType→LaTeX left its OLE source behind.");
            AssertEqual(
                0,
                CountMathTypePlaceRefFields(
                    document),
                "MathType→LaTeX left MTPlaceRef numbering fields behind.");

            Console.WriteLine(
                "[host-core pure OMML] mixed inline/display/numbered conversion matrix passed: OMML↔VisualTeX, OMML↔MathType, and all three source formats→LaTeX.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                previousNativePreview);
            try { File.Delete(emfPath); } catch { }
            try { File.Delete(pngPath); } catch { }
            try { File.Delete(svgPath); } catch { }
            try { Directory.Delete(assetRoot, recursive: true); } catch { }
        }
    }

    private static void RunPureOmmlMathTypeBoundaryAcceptance(
        string artifactRoot)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>m</mi><mo>+</mo><mn>3</mn></mrow></math>";
        const string prefixText =
            "mt-left ";
        const string suffixText =
            " mt-right";

        var assetRoot =
            Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX",
                "office",
                "temp",
                "pure-omml-mathtype-"
                + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(
            assetRoot);
        var svgPath =
            Path.Combine(
                assetRoot,
                "preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"96\" viewBox=\"0 0 260 96\"><rect width=\"260\" height=\"96\" fill=\"white\"/><text x=\"10\" y=\"66\" font-family=\"Cambria Math\" font-size=\"44\">m+3</text></svg>");
        var emfPath =
            OfficeOlePreview.CreateVectorEmfFromSvg(
                svgPath,
                260,
                96);
        var pngPath =
            Path.Combine(
                assetRoot,
                "preview.png");
        WriteAcceptancePng(
            pngPath,
            "m+3",
            260,
            96);

        var previousNativePreview =
            Environment.GetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW");

        using var host =
            new WordPerformanceHost(
                documentPath: null);
        var application =
            host.Application;
        var document =
            host.Document;
        var service =
            new WordFormulaService(
                application);

        Word.Selection? selection = null;
        Word.Range? insertion = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Word.Range? ommlRange = null;
        try
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                "1");

            selection =
                application.Selection;
            selection.SetRange(
                0,
                0);
            selection.TypeText(
                prefixText);

            insertion =
                selection.Range.Duplicate;
            var sourceSession =
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "inline",
                    numbered: false,
                    FormulaOleContract.MathTypeOleMode);
            sourceSession.SourceDocumentId =
                document.FullName;
            sourceSession.SourceObjectId =
                WordRangeReference(
                    insertion.Start,
                    insertion.End);

            _ = service.InsertMathTypeOle(
                sourceSession,
                mathMl,
                emfPath,
                updateCreatedMathTypeNumberFields: true);

            Release(insertion);
            insertion = null;
            selection.TypeText(
                suffixText);

            AssertEqual(
                1,
                CountMathTypeOleShapes(
                    document),
                "Pure OMML MathType boundary fixture did not create one MathType OLE.");

            shapes =
                document.InlineShapes;
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate =
                        shapes[index];
                    if (!MathTypeOleInterop
                        .IsMathTypeOle(
                            candidate))
                        continue;
                    if (shape is not null)
                        throw new InvalidDataException(
                            "Pure OMML MathType boundary found more than one MathType source.");
                    shape =
                        candidate;
                    candidate = null;
                }
                finally
                {
                    Release(candidate);
                }
            }

            AssertTrue(
                shape is not null,
                "Pure OMML MathType boundary could not resolve the source OLE.");
            shapeRange =
                shape!.Range.Duplicate;
            AssertHostCoreImmediateText(
                document,
                shapeRange,
                prefixText,
                suffixText,
                "before MathType→pure OMML");
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(
                    mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    MathTypeOleStorage.ReadMathMl(
                        shape)),
                "MathType source semantics changed before conversion.");

            shapeRange.Select();
            var toOmmlPlan =
                service
                    .CaptureFormulaFormatConversionPlan(
                        wholeDocument: false,
                        FormulaOleContract.MathTypeOleMode,
                        FormulaOleContract.WordOmmlMode);
            AssertEqual(
                1,
                toOmmlPlan.Targets.Count,
                "MathType→pure OMML did not capture exactly one MathType source.");
            var toOmmlResult =
                service
                    .ApplyFormulaFormatConversionPlan(
                        toOmmlPlan,
                        PrepareOmmlMathTypeTargets(
                            toOmmlPlan,
                            emfPath));
            AssertEqual(
                1,
                toOmmlResult.FormulaCount,
                "MathType→pure OMML did not convert exactly one formula.");
            AssertEqual(
                0,
                toOmmlResult.FailedFormulaCount,
                "MathType→pure OMML reported a failure.");
            AssertEqual(
                0,
                CountMathTypeOleShapes(
                    document),
                "MathType→pure OMML left the source MathType OLE alive.");

            Release(shapeRange);
            shapeRange = null;
            Release(shape);
            shape = null;
            Release(shapes);
            shapes = null;

            var ommlIndex =
                WordFormulaHostResolver
                    .CaptureDocumentIndex(
                        document);
            AssertEqual(
                1,
                ommlIndex.Omml.Count,
                "MathType→pure OMML did not leave one Word OMath.");
            var ommlHost =
                ommlIndex.Omml[0];
            AssertTrue(
                string.IsNullOrWhiteSpace(
                    ommlHost.FormulaId),
                "MathType→pure OMML leaked a durable VisualTeX FormulaId.");
            var ommlPayload =
                WordFormulaHostSemanticReader
                    .Read(
                        document,
                        ommlHost);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(
                    mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    ommlPayload.MathMl
                    ?? throw new InvalidDataException(
                        "MathType→pure OMML target returned no MathML.")),
                "MathType→pure OMML changed formula semantics.");

            ommlRange =
                WordFormulaHostSemanticReader
                    .CreateRange(
                        document,
                        ommlHost.Range);
            AssertHostCoreImmediateText(
                document,
                ommlRange,
                prefixText,
                suffixText,
                "after MathType→pure OMML");
            AssertNoVisualTeXOmmlArtifacts(
                document,
                "after MathType→pure OMML");

            ommlRange.Select();
            var toMathTypePlan =
                service
                    .CaptureFormulaFormatConversionPlan(
                        wholeDocument: false,
                        FormulaOleContract.WordOmmlMode,
                        FormulaOleContract.MathTypeOleMode);
            AssertEqual(
                1,
                toMathTypePlan.Targets.Count,
                "Pure OMML→MathType did not capture exactly one OMath.");
            var toMathTypeResult =
                service
                    .ApplyFormulaFormatConversionPlan(
                        toMathTypePlan,
                        PrepareOmmlMathTypeTargets(
                            toMathTypePlan,
                            emfPath));
            AssertEqual(
                1,
                toMathTypeResult.FormulaCount,
                "Pure OMML→MathType did not convert exactly one formula.");
            AssertEqual(
                0,
                toMathTypeResult.FailedFormulaCount,
                "Pure OMML→MathType reported a failure.");
            AssertEqual(
                0,
                document.OMaths.Count,
                "Pure OMML→MathType left the source OMath alive.");
            AssertEqual(
                1,
                CountMathTypeOleShapes(
                    document),
                "Pure OMML→MathType did not recreate one MathType OLE.");

            Release(ommlRange);
            ommlRange = null;

            shapes =
                document.InlineShapes;
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate =
                        shapes[index];
                    if (!MathTypeOleInterop
                        .IsMathTypeOle(
                            candidate))
                        continue;
                    if (shape is not null)
                        throw new InvalidDataException(
                            "Pure OMML→MathType created more than one MathType OLE.");
                    shape =
                        candidate;
                    candidate = null;
                }
                finally
                {
                    Release(candidate);
                }
            }

            AssertTrue(
                shape is not null,
                "Pure OMML→MathType target OLE is missing.");
            shapeRange =
                shape!.Range.Duplicate;
            AssertHostCoreImmediateText(
                document,
                shapeRange,
                prefixText,
                suffixText,
                "after pure OMML→MathType");
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(
                    mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    MathTypeOleStorage.ReadMathMl(
                        shape)),
                "Pure OMML→MathType changed formula semantics.");

            SetInlineOleObjectCharacterFontSizeForAcceptance(
                document,
                shape,
                10.5f);
            shapeRange.Select();
            var toVisualTeXPlan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: false,
                    FormulaOleContract.MathTypeOleMode,
                    FormulaOleContract.NativeOleMode);
            AssertEqual(
                1,
                toVisualTeXPlan.Targets.Count,
                "MathType→VisualTeX presentation-scale regression did not capture one source.");
            var toVisualTeXTarget =
                toVisualTeXPlan.Targets.Single();
            AssertNear(
                10.5f,
                (float)toVisualTeXTarget.FontSizePt,
                0.01f,
                "MathType→VisualTeX did not persist the Word-visible 10.5 pt size.");

            var toVisualTeXPrepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal)
                {
                    [toVisualTeXTarget.Id] =
                        new PreparedWordBulkFormula
                        {
                            Run = new WordBulkRun
                            {
                                Id = toVisualTeXTarget.Id,
                                IsFormula = true,
                                Latex = toVisualTeXTarget.Latex,
                                DisplayMode = toVisualTeXTarget.DisplayMode,
                            },
                            Session = CreateSimpleFormatTargetSession(
                                toVisualTeXTarget,
                                FormulaOleContract.NativeOleMode,
                                mathMl),
                            MathMl = mathMl,
                            PngPath = pngPath,
                            EmfPath = emfPath,
                        },
                };
            var toVisualTeXResult =
                service.ApplyFormulaFormatConversionPlan(
                    toVisualTeXPlan,
                    toVisualTeXPrepared);
            AssertEqual(
                1,
                toVisualTeXResult.FormulaCount,
                "MathType→VisualTeX presentation-scale regression did not convert one formula.");
            AssertEqual(
                0,
                toVisualTeXResult.FailedFormulaCount,
                "MathType→VisualTeX presentation-scale regression reported a failure.");
            AssertEqual(
                0,
                CountMathTypeOleShapes(document),
                "MathType→VisualTeX presentation-scale regression left the MathType source alive.");

            Release(shapeRange);
            shapeRange = null;
            Release(shape);
            shape = null;
            Release(shapes);
            shapes = document.InlineShapes;
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate))
                        continue;
                    shape = candidate;
                    candidate = null;
                    break;
                }
                finally
                {
                    Release(candidate);
                }
            }
            AssertTrue(
                shape is not null,
                "MathType→VisualTeX presentation-scale regression produced no VisualTeX target.");
            var visualMetadata =
                WordFormulaMetadataReader.TryReadEmbeddedNativeOle(
                    shape!)
                ?? throw new InvalidDataException(
                    "MathType→VisualTeX presentation-scale target has no embedded metadata.");
            AssertNear(
                10.5f,
                (float)(visualMetadata.FontSizePt ?? 0),
                0.01f,
                "MathType→VisualTeX embedded metadata did not retain the Word-visible 10.5 pt size.");
            AssertNear(
                10.5f,
                (float)(visualMetadata.RenderFontSizePt ?? 0),
                0.01f,
                "MathType→VisualTeX render metadata did not retain the Word-visible 10.5 pt size.");
            shape!.Range.Select();
            var reopenedVisualTeX = service.ReadSelection();
            AssertNear(
                10.5f,
                (float)(reopenedVisualTeX.Metadata?.FontSizePt ?? 0),
                0.01f,
                "Reopening the converted VisualTeX formula restored the hidden 12 pt MathType size.");
            AssertNear(
                260f * 0.75f,
                shape.Width,
                0.75f,
                "MathType→VisualTeX unexpectedly applied an extra presentation scale to width.");
            AssertNear(
                96f * 0.75f,
                shape.Height,
                0.75f,
                "MathType→VisualTeX unexpectedly applied an extra presentation scale to height.");

            Console.WriteLine(
                "[host-core pure OMML] MathType↔OMML preserved semantics/text; MathType→VisualTeX stores/reopens the Word-visible 10.5 pt size without a second presentation scale.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                previousNativePreview);
            Release(ommlRange);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(insertion);
            Release(selection);
        }
    }

    private static void SetInlineOleObjectCharacterFontSizeForAcceptance(
        Word.Document document,
        Word.InlineShape shape,
        float fontSizePoints)
    {
        Word.Range? shapeRange = null;
        Word.Range? probe = null;
        Word.Font? font = null;
        try
        {
            shapeRange = shape.Range.Duplicate;
            for (var position = shapeRange.Start;
                 position < shapeRange.End;
                 position++)
            {
                Release(font);
                font = null;
                Release(probe);
                probe = document.Range(
                    position,
                    position + 1);
                if (!string.Equals(
                        probe.Text,
                        "\u0001",
                        StringComparison.Ordinal))
                    continue;

                font = probe.Font;
                font.Size = fontSizePoints;
                return;
            }

            throw new InvalidDataException(
                "Acceptance MathType OLE has no U+0001 object character.");
        }
        finally
        {
            Release(font);
            Release(probe);
            Release(shapeRange);
        }
    }

    private static void RunPureOmmlCopyPasteAcceptance()
    {
        const string firstMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>a</mi><mo>=</mo><mn>1</mn></mrow></math>";
        const string secondMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>b</mi><mo>=</mo><mn>2</mn></mrow></math>";

        using var host =
            new WordPerformanceHost(
                documentPath: null);
        var application =
            host.Application;
        var document =
            host.Document;
        var service =
            new WordFormulaService(
                application);
        var selection =
            application.Selection;

        WordEquationNumbering.SetEquationNumberFormatPreference(
            document,
            EquationNumberFormat.ContinuousId);

        selection.SetRange(0, 0);
        _ = service.InsertOmml(
            CreateOmmlMathTypeAcceptanceSession(
                firstMathMl,
                "block",
                numbered: true,
                FormulaOleContract.WordOmmlMode),
            firstMathMl);

        var sourceIndex =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        var source =
            sourceIndex.Omml.Single();
        source.Numbering =
            WordFormulaNumberingResolver.ResolveLocal(
                document,
                source);
        AssertEqual(
            WordFormulaNumberingContainerKind.CanonicalNativeOmml,
            source.Numbering.ContainerKind,
            "Pure OMML copy source is not native numbered OMath.");

        Word.Range? sourceRange = null;
        try
        {
            sourceRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    source.Range);
            sourceRange.Select();

            var snapshot =
                service.CaptureSelectedFormulaForCopy()
                ?? throw new InvalidDataException(
                    "Pure OMML copy did not create a host-core snapshot.");
            AssertTrue(
                snapshot.UsesHostCore
                && snapshot.Metadata.Numbered,
                "Pure OMML copy snapshot lost host-core/numbered state.");

            selection.Copy();
            selection.EndKey(
                Word.WdUnits.wdStory);
            selection.TypeParagraph();
            service.ArmHostCorePaste(
                snapshot);
            selection.Paste();

            var repair =
                service.RepairPastedFormula(
                    snapshot);
            for (var attempt = 0;
                 repair ==
                    WordFormulaService.PastedFormulaRepairResult.NotReady
                 && attempt < 5;
                 attempt++)
            {
                Thread.Sleep(30);
                repair =
                    service.RepairPastedFormula(
                        snapshot);
            }
            AssertEqual(
                WordFormulaService.PastedFormulaRepairResult.Repaired,
                repair,
                "Pure OMML native paste was not recognized by CopyPasteCore.");

            var pastedIndex =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            var canonical =
                pastedIndex.Omml
                    .Where(item =>
                    {
                        item.Numbering =
                            WordFormulaNumberingResolver.ResolveLocal(
                                document,
                                item);
                        return item.Numbering.ContainerKind ==
                            WordFormulaNumberingContainerKind.CanonicalNativeOmml;
                    })
                    .ToArray();
            AssertEqual(
                2,
                canonical.Length,
                "One native OMML paste did not leave source + one pasted native equation.");
            AssertTrue(
                canonical.All(item =>
                    string.IsNullOrWhiteSpace(
                        item.FormulaId)),
                "Native OMML paste introduced durable VisualTeX identities.");
            AssertNoVisualTeXOmmlArtifacts(
                document,
                "after pure OMML single paste");

            object oneUndo = 1;
            AssertTrue(
                document.Undo(
                    ref oneUndo),
                "Word refused to undo the native OMML paste once.");

            var afterUndo =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            var afterUndoCanonical =
                afterUndo.Omml
                    .Count(item =>
                    {
                        item.Numbering =
                            WordFormulaNumberingResolver.ResolveLocal(
                                document,
                                item);
                        return item.Numbering.ContainerKind ==
                            WordFormulaNumberingContainerKind.CanonicalNativeOmml;
                    });
            AssertEqual(
                1,
                afterUndoCanonical,
                "One Word Undo did not remove the pasted native OMML copy.");
            AssertNoVisualTeXOmmlArtifacts(
                document,
                "after one native-paste Undo");
        }
        finally
        {
            Release(sourceRange);
        }

        // Group copy/paste: two native equations are copied and pasted as one
        // ordinary Word paste. CopyPasteCore only validates the local result.
        document.Content.Text =
            string.Empty;
        selection.SetRange(
            0,
            0);
        _ = service.InsertOmml(
            CreateOmmlMathTypeAcceptanceSession(
                firstMathMl,
                "block",
                numbered: true,
                FormulaOleContract.WordOmmlMode),
            firstMathMl);
        selection.EndKey(
            Word.WdUnits.wdStory);
        selection.TypeParagraph();
        _ = service.InsertOmml(
            CreateOmmlMathTypeAcceptanceSession(
                secondMathMl,
                "block",
                numbered: true,
                FormulaOleContract.WordOmmlMode),
            secondMathMl);

        var groupIndex =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        AssertEqual(
            2,
            groupIndex.Omml.Count,
            "Pure OMML group copy fixture did not contain two source equations.");
        var ordered =
            groupIndex.Omml
                .OrderBy(item =>
                    item.Range.Start)
                .ToArray();
        selection.SetRange(
            ordered[0].Range.Start,
            ordered[1].Range.End);

        var groupSnapshot =
            service.CaptureSelectedFormulaForCopy()
            ?? throw new InvalidDataException(
                "Pure OMML group copy did not create a snapshot.");
        AssertEqual(
            2,
            groupSnapshot.GroupItems.Count,
            "Pure OMML group snapshot did not contain two equations.");

        selection.Copy();
        selection.EndKey(
            Word.WdUnits.wdStory);
        selection.TypeParagraph();
        service.ArmHostCorePaste(
            groupSnapshot);
        selection.Paste();

        var groupRepair =
            service.RepairPastedFormula(
                groupSnapshot);
        for (var attempt = 0;
             groupRepair ==
                WordFormulaService.PastedFormulaRepairResult.NotReady
             && attempt < 5;
             attempt++)
        {
            Thread.Sleep(30);
            groupRepair =
                service.RepairPastedFormula(
                    groupSnapshot);
        }
        AssertEqual(
            WordFormulaService.PastedFormulaRepairResult.Repaired,
            groupRepair,
            "Pure OMML group paste was not recognized by CopyPasteCore.");

        var afterGroupPaste =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        var groupCanonical =
            afterGroupPaste.Omml
                .Count(item =>
                {
                    item.Numbering =
                        WordFormulaNumberingResolver.ResolveLocal(
                            document,
                            item);
                    return item.Numbering.ContainerKind ==
                        WordFormulaNumberingContainerKind.CanonicalNativeOmml;
                });
        AssertEqual(
            4,
            groupCanonical,
            "Pure OMML group paste did not leave two sources + two pasted equations.");
        AssertNoVisualTeXOmmlArtifacts(
            document,
            "after pure OMML group paste");

        object groupUndo = 1;
        AssertTrue(
            document.Undo(
                ref groupUndo),
            "Word refused to undo the native OMML group paste once.");
        var afterGroupUndo =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        var groupUndoCanonical =
            afterGroupUndo.Omml
                .Count(item =>
                {
                    item.Numbering =
                        WordFormulaNumberingResolver.ResolveLocal(
                            document,
                            item);
                    return item.Numbering.ContainerKind ==
                        WordFormulaNumberingContainerKind.CanonicalNativeOmml;
                });
        AssertEqual(
            2,
            groupUndoCanonical,
            "One Word Undo did not remove the pasted native OMML group.");
        AssertNoVisualTeXOmmlArtifacts(
            document,
            "after one native group-paste Undo");

        Console.WriteLine(
            "[host-core pure OMML] single/group native Paste stayed marker-free and one Word Undo removed each pasted copy.");
    }

    private static void RunPureOmmlDeleteReferenceAcceptance()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>F</mi><mo>=</mo><mi>m</mi><mi>a</mi></mrow></math>";

        using var host =
            new WordPerformanceHost(
                documentPath: null);
        var application =
            host.Application;
        var document =
            host.Document;
        var service =
            new WordFormulaService(
                application);
        var selection =
            application.Selection;

        WordEquationNumbering.SetEquationNumberFormatPreference(
            document,
            EquationNumberFormat.ContinuousId);
        selection.SetRange(
            0,
            0);
        _ = service.InsertOmml(
            CreateOmmlMathTypeAcceptanceSession(
                mathMl,
                "block",
                numbered: true,
                FormulaOleContract.WordOmmlMode),
            mathMl);

        var targets =
            service.GetCanonicalEquationReferenceTargets(
                document)
                .Where(item =>
                    item.Source ==
                        EquationReferenceSource.WordOmml)
                .ToArray();
        AssertEqual(
            1,
            targets.Length,
            "Pure delete/reference fixture has no native Word target.");

        selection.EndKey(
            Word.WdUnits.wdStory);
        selection.TypeParagraph();
        selection.TypeText(
            "ref ");
        service.InsertEquationReference(
            document,
            selection,
            targets[0],
            EquationReferenceStyle.Parenthesized,
            Word.WdColor.wdColorAutomatic);

        var source =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document)
                .Omml
                .First(item =>
                {
                    item.Numbering =
                        WordFormulaNumberingResolver.ResolveLocal(
                            document,
                            item);
                    return item.Numbering.ContainerKind ==
                        WordFormulaNumberingContainerKind.CanonicalNativeOmml;
                });
        Word.Range? sourceRange = null;
        try
        {
            sourceRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    source.Range);
            sourceRange.Select();
            _ = service.DeleteSelectedFormula();
        }
        finally
        {
            Release(sourceRange);
        }

        var remaining =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        AssertTrue(
            !remaining.Omml.Any(item =>
            {
                item.Numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        item);
                return item.Numbering.ContainerKind ==
                    WordFormulaNumberingContainerKind.CanonicalNativeOmml;
            }),
            "Deleting pure numbered OMML left the source equation alive.");
        AssertEqual(
            0,
            service.GetCanonicalEquationReferenceTargets(
                    document)
                .Count(item =>
                    item.Source ==
                        EquationReferenceSource.WordOmml),
            "Deleted pure OMML still appears as a Word-native reference target.");

        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? result = null;
        try
        {
            fields =
                document.Fields;
            var foundBrokenRef =
                false;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(result);
                result = null;
                Release(code);
                code = null;
                Release(field);
                field =
                    fields[index];
                code =
                    field.Code.Duplicate;
                if (!(code.Text ?? string.Empty)
                    .TrimStart()
                    .StartsWith(
                        "REF ",
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                field.Update();
                result =
                    field.Result.Duplicate;
                AssertTrue(
                    IsMissingReferenceResult(
                        result.Text
                        ?? string.Empty),
                    "Word native REF did not report a missing source after deleting the OMML equation.");
                foundBrokenRef =
                    true;
                break;
            }
            AssertTrue(
                foundBrokenRef,
                "Deleting pure OMML removed the REF field instead of leaving Word's natural broken reference.");
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
        }

        AssertNoVisualTeXOmmlArtifacts(
            document,
            "after deleting pure OMML with native REF");
        Console.WriteLine(
            "[host-core pure OMML] deleting a numbered equation left Word's native REF in the natural missing-source state.");
    }

    private static void RunPureOmmlNumberToggleAcceptance()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>u</mi><mo>=</mo><mi>v</mi></mrow></math>";

        using var host =
            new WordPerformanceHost(
                documentPath: null);
        var application =
            host.Application;
        var document =
            host.Document;
        var service =
            new WordFormulaService(
                application);
        var selection =
            application.Selection;

        WordEquationNumbering.SetEquationNumberFormatPreference(
            document,
            EquationNumberFormat.ContinuousId);
        selection.SetRange(
            0,
            0);
        _ = service.InsertOmml(
            CreateOmmlMathTypeAcceptanceSession(
                mathMl,
                "block",
                numbered: true,
                FormulaOleContract.WordOmmlMode),
            mathMl);

        void ReplaceNumberedState(
            bool numbered)
        {
            var index =
                WordFormulaHostResolver.CaptureDocumentIndex(
                    document);
            var source =
                index.Omml.Single();
            Word.Range? range = null;
            try
            {
                range =
                    WordFormulaHostSemanticReader.CreateRange(
                        document,
                        source.Range);
                range.Select();
                var opened =
                    service.ReadSelection();
                var session =
                    CreateOmmlMathTypeAcceptanceSession(
                        mathMl,
                        "block",
                        numbered,
                        FormulaOleContract.WordOmmlMode);
                session.Mode =
                    "edit";
                session.FormulaId =
                    opened.FormulaId!;
                session.SourceDocumentId =
                    opened.DocumentId;
                session.SourceObjectId =
                    opened.ObjectId;
                session.OriginalMetadata =
                    opened.Metadata;
                _ = service.ReplaceOmml(
                    session,
                    mathMl);
            }
            finally
            {
                Release(range);
            }
        }

        ReplaceNumberedState(
            numbered: false);
        var offIndex =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        AssertEqual(
            1,
            offIndex.Omml.Count,
            "Numbered→unnumbered pure OMML changed the source OMath count.");
        var off =
            offIndex.Omml.Single();
        off.Numbering =
            WordFormulaNumberingResolver.ResolveLocal(
                document,
                off);
        AssertTrue(
            !off.Numbering.Numbered
            && off.Numbering.ContainerKind ==
                WordFormulaNumberingContainerKind.None,
            "Numbered→unnumbered pure OMML retained native numbering.");
        AssertNoVisualTeXOmmlArtifacts(
            document,
            "after pure OMML numbering off");

        ReplaceNumberedState(
            numbered: true);
        var onIndex =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        AssertEqual(
            1,
            onIndex.Omml.Count,
            "Unnumbered→numbered pure OMML changed the source OMath count.");
        var on =
            onIndex.Omml.Single();
        on.Numbering =
            WordFormulaNumberingResolver.ResolveLocal(
                document,
                on);
        AssertEqual(
            WordFormulaNumberingContainerKind.CanonicalNativeOmml,
            on.Numbering.ContainerKind,
            "Unnumbered→numbered pure OMML did not restore Word-native #(SEQ).");
        AssertTrue(
            string.IsNullOrWhiteSpace(
                on.FormulaId),
            "Numbering off→on created a durable VisualTeX FormulaId.");
        AssertNoVisualTeXOmmlArtifacts(
            document,
            "after pure OMML numbering on");

        _ = service.UpdateEquationNumbers();
        AssertNoVisualTeXOmmlArtifacts(
            document,
            "after pure OMML number refresh");

        Console.WriteLine(
            "[host-core pure OMML] numbering off→on + refresh stayed entirely in Word-native OMath state.");
    }

    private static void EnsurePureConversionUserBlankParagraph(
        Word.Document document)
    {
        if (CountPureConversionOrdinaryEmptyParagraphs(document) > 0)
            return;

        var index =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        var target = index.Omml
            .Where(host => host.Display)
            .OrderBy(host => host.Range.Start)
            .FirstOrDefault()
            ?? throw new InvalidDataException(
                "The OMML→MathType structure fixture has no display equation before which to insert a user blank paragraph.");

        Word.Range? range = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        try
        {
            range =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    target.Range);
            paragraphs = range.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "The OMML→MathType structure fixture display equation does not occupy one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            paragraphRange.InsertParagraphBefore();
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(range);
        }
    }

    private static void AssertExternalReferenceTarget(
        Word.Document document,
        string bookmarkName,
        bool expectTargetExists,
        string context)
    {
        Word.Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            AssertEqual(
                expectTargetExists,
                bookmarks.Exists(
                    bookmarkName),
                context
                + " has the wrong reference-target existence state.");

            var externalReferences =
                WordEquationReferenceFields
                    .CaptureReferenceCounts(
                        document);
            externalReferences.TryGetValue(
                bookmarkName,
                out var count);
            AssertEqual(
                1,
                count,
                context
                + " did not contain exactly one external REF to the expected equation target.");
        }
        finally
        {
            Release(bookmarks);
        }
    }

    private static void AssertMigratedPureOmmlReference(
        Word.Document document,
        string retiredVisualTeXTarget,
        string context)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? result = null;
        Word.Bookmark? nativeTarget = null;
        Word.Range? nativeTargetRange = null;
        string? migratedTargetName = null;
        var matchingReferences = 0;
        try
        {
            bookmarks = document.Bookmarks;
            AssertTrue(
                !bookmarks.Exists(
                    retiredVisualTeXTarget),
                context
                + " retained the retired VisualTeX VTEqNum target.");

            fields = document.Fields;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(result);
                result = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                if (field.Type !=
                    Word.WdFieldType.wdFieldRef)
                    continue;

                code =
                    field.Code.Duplicate;
                var codeText =
                    (code.Text
                        ?? string.Empty)
                    .Trim();
                AssertTrue(
                    codeText.IndexOf(
                        retiredVisualTeXTarget,
                        StringComparison.OrdinalIgnoreCase)
                    < 0,
                    context
                    + " left a REF field pointing at the retired VisualTeX bookmark.");

                if (!codeText.StartsWith(
                        "REF _Ref",
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                var firstSpace =
                    codeText.IndexOf(' ');
                var tail =
                    firstSpace < 0
                        ? string.Empty
                        : codeText.Substring(
                            firstSpace + 1)
                            .TrimStart();
                var tokenEnd =
                    tail.IndexOfAny(
                        new[] { ' ', '\\' });
                var targetName =
                    tokenEnd < 0
                        ? tail
                        : tail.Substring(
                            0,
                            tokenEnd);
                if (!targetName.StartsWith(
                        "_Ref",
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                matchingReferences++;
                migratedTargetName =
                    targetName;
                AssertTrue(
                    bookmarks.Exists(
                        targetName),
                    context
                    + " points to a missing Word-native _Ref bookmark.");

                result =
                    field.Result.Duplicate;
                var resultText =
                    result.Text
                    ?? string.Empty;
                AssertTrue(
                    resultText.IndexOf(
                        "引用源",
                        StringComparison.OrdinalIgnoreCase)
                        < 0
                    && resultText.IndexOf(
                        "reference source",
                        StringComparison.OrdinalIgnoreCase)
                        < 0
                    && resultText.IndexOf(
                        "Error!",
                        StringComparison.OrdinalIgnoreCase)
                        < 0,
                    context
                    + " still renders Word's missing-reference error.");

                nativeTarget =
                    bookmarks[targetName];
                nativeTargetRange =
                    nativeTarget.Range.Duplicate;
                AssertEqual(
                    nativeTargetRange.Text
                        ?? string.Empty,
                    resultText,
                    context
                    + " REF result differs from its native OMML number target.");
            }

            AssertEqual(
                1,
                matchingReferences,
                context
                + " did not migrate exactly one body REF to Word-native _Ref.");
            AssertTrue(
                !string.IsNullOrWhiteSpace(
                    migratedTargetName),
                context
                + " did not expose a native _Ref target name.");
        }
        finally
        {
            Release(nativeTargetRange);
            Release(nativeTarget);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(bookmarks);
        }
    }

    private static void AssertPureOmmlDisplayParagraphBoundaries(
        Word.Document document,
        string context)
    {
        var index =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        foreach (var host in index.Omml
                     .Where(item => item.Display))
        {
            Word.Range? hostRange = null;
            Word.Paragraphs? paragraphs = null;
            Word.Paragraph? paragraph = null;
            Word.Range? paragraphRange = null;
            Word.Range? prefix = null;
            try
            {
                hostRange =
                    WordFormulaHostSemanticReader.CreateRange(
                        document,
                        host.Range);
                paragraphs =
                    hostRange.Paragraphs;
                AssertEqual(
                    1,
                    paragraphs.Count,
                    context
                    + " display OMML spans multiple Word paragraphs.");
                paragraph =
                    paragraphs[1];
                paragraphRange =
                    paragraph.Range.Duplicate;
                prefix =
                    document.Range(
                        paragraphRange.Start,
                        hostRange.Start);
                var prefixText =
                    prefix.Text
                    ?? string.Empty;
                AssertTrue(
                    prefixText.IndexOf('\t') < 0,
                    context
                    + " retained VisualTeX's leading TAB before a pure OMML display equation.");
                AssertTrue(
                    prefixText.IndexOf('\v') < 0,
                    context
                    + " materialized a manual line break before a pure OMML display equation.");
                AssertTrue(
                    prefixText.All(character =>
                        character is '\r' or '\n'
                        || char.IsWhiteSpace(
                            character)),
                    context
                    + " retained non-structural text before a pure OMML display equation.");
            }
            finally
            {
                Release(prefix);
                Release(paragraphRange);
                Release(paragraph);
                Release(paragraphs);
                Release(hostRange);
            }
        }
    }

    private static void AssertPureConversionVisualTeXTabStructure(
        Word.Document document,
        int expectedNumberedCount,
        int expectedOrdinaryEmptyParagraphCount)
    {
        AssertEqual(
            0,
            document.Tables.Count,
            "OMML→VisualTeX introduced a numbering/layout table instead of the canonical center/right-tab paragraph.");
        AssertEqual(
            expectedOrdinaryEmptyParagraphCount,
            CountPureConversionOrdinaryEmptyParagraphs(document),
            "OMML→VisualTeX introduced an extra ordinary empty paragraph around a converted formula.");

        var index =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        var numberedCount = 0;
        foreach (var host in index.VisualTeX)
        {
            var numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            if (!numbering.Numbered)
                continue;
            numberedCount++;
            AssertEqual(
                WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph,
                numbering.ContainerKind,
                "OMML→VisualTeX did not materialize the numbered target as the canonical body tab paragraph.");
            AssertTrue(
                numbering.NumberRange is not null,
                "OMML→VisualTeX tab paragraph lost its visible number range.");

            Word.Range? hostRange = null;
            Word.Paragraphs? paragraphs = null;
            Word.Paragraph? paragraph = null;
            Word.Range? paragraphRange = null;
            Word.ParagraphFormat? format = null;
            Word.TabStops? tabs = null;
            Word.TabStop? tab = null;
            Word.Range? leading = null;
            Word.Range? numberRange = null;
            try
            {
                hostRange =
                    WordFormulaHostSemanticReader.CreateRange(
                        document,
                        host.Range);
                AssertTrue(
                    !Convert.ToBoolean(
                        hostRange.get_Information(
                            Word.WdInformation.wdWithInTable)),
                    "OMML→VisualTeX left a numbered OLE inside a Word table.");
                paragraphs = hostRange.Paragraphs;
                AssertEqual(
                    1,
                    paragraphs.Count,
                    "OMML→VisualTeX numbered OLE spans more than one paragraph.");
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
                format = paragraphRange.ParagraphFormat;
                AssertEqual(
                    Word.WdParagraphAlignment.wdAlignParagraphJustify,
                    format.Alignment,
                    "OMML→VisualTeX numbered paragraph is not using the tab-driven justified layout.");
                tabs = format.TabStops;
                var hasCenter = false;
                var hasRight = false;
                for (var tabIndex = 1;
                     tabIndex <= tabs.Count;
                     tabIndex++)
                {
                    Release(tab);
                    tab = tabs[tabIndex];
                    hasCenter |=
                        tab.Alignment ==
                        Word.WdTabAlignment.wdAlignTabCenter;
                    hasRight |=
                        tab.Alignment ==
                        Word.WdTabAlignment.wdAlignTabRight;
                }
                AssertTrue(
                    hasCenter && hasRight,
                    "OMML→VisualTeX numbered paragraph lost its center/right tab stops.");
                AssertTrue(
                    hostRange.Start > paragraphRange.Start,
                    "OMML→VisualTeX numbered OLE has no leading center-tab boundary.");
                leading = document.Range(
                    hostRange.Start - 1,
                    hostRange.Start);
                AssertEqual(
                    "\t",
                    leading.Text,
                    "OMML→VisualTeX numbered OLE is not positioned immediately after the center tab.");
                numberRange =
                    WordFormulaHostSemanticReader.CreateRange(
                        document,
                        numbering.NumberRange!);
                AssertTrue(
                    numberRange.Start >= hostRange.End
                    && numberRange.Start >= paragraphRange.Start
                    && numberRange.End <= paragraphRange.End,
                    "OMML→VisualTeX visible number is not on the same paragraph after the OLE host.");
            }
            finally
            {
                Release(numberRange);
                Release(leading);
                Release(tab);
                Release(tabs);
                Release(format);
                Release(paragraphRange);
                Release(paragraph);
                Release(paragraphs);
                Release(hostRange);
            }
        }

        AssertEqual(
            expectedNumberedCount,
            numberedCount,
            "OMML→VisualTeX returned the wrong number of canonical numbered tab paragraphs.");
    }

    private static void AssertPureConversionMathTypeTabStructure(
        Word.Document document,
        int expectedNumberedCount,
        int expectedOrdinaryEmptyParagraphCount)
    {
        AssertEqual(
            0,
            document.Tables.Count,
            "OMML→MathType introduced a Word table instead of MathType's native tab paragraph.");
        AssertEqual(
            expectedOrdinaryEmptyParagraphCount,
            CountPureConversionOrdinaryEmptyParagraphs(document),
            "OMML→MathType introduced an extra ordinary empty paragraph around a converted formula.");

        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabs = null;
        Word.TabStop? tab = null;
        var numberedCount = 0;
        try
        {
            shapes = document.InlineShapes;
            for (var shapeIndex = 1;
                 shapeIndex <= shapes.Count;
                 shapeIndex++)
            {
                Release(shape);
                shape = shapes[shapeIndex];
                if (!MathTypeOleInterop.IsMathTypeOle(shape))
                    continue;

                Release(shapeRange);
                shapeRange = shape.Range.Duplicate;
                Release(paragraphs);
                paragraphs = shapeRange.Paragraphs;
                AssertEqual(
                    1,
                    paragraphs.Count,
                    "OMML→MathType produced a display OLE spanning multiple paragraphs.");
                Release(paragraph);
                paragraph = paragraphs[1];
                Release(paragraphRange);
                paragraphRange = paragraph.Range.Duplicate;
                Release(fields);
                fields = paragraphRange.Fields;
                var hasPlaceRef = false;
                for (var fieldIndex = 1;
                     fieldIndex <= fields.Count;
                     fieldIndex++)
                {
                    Release(code);
                    code = null;
                    Release(field);
                    field = fields[fieldIndex];
                    code = field.Code.Duplicate;
                    if ((code.Text ?? string.Empty).IndexOf(
                            "MTPlaceRef",
                            StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        hasPlaceRef = true;
                        break;
                    }
                }
                if (!hasPlaceRef)
                    continue;

                numberedCount++;
                AssertTrue(
                    !Convert.ToBoolean(
                        shapeRange.get_Information(
                            Word.WdInformation.wdWithInTable)),
                    "OMML→MathType left a numbered MathType OLE inside a Word table.");
                Release(format);
                format = paragraphRange.ParagraphFormat;
                Release(tabs);
                tabs = format.TabStops;
                var hasCenter = false;
                var hasRight = false;
                for (var tabIndex = 1;
                     tabIndex <= tabs.Count;
                     tabIndex++)
                {
                    Release(tab);
                    tab = tabs[tabIndex];
                    hasCenter |=
                        tab.Alignment ==
                        Word.WdTabAlignment.wdAlignTabCenter;
                    hasRight |=
                        tab.Alignment ==
                        Word.WdTabAlignment.wdAlignTabRight;
                }
                AssertTrue(
                    hasCenter && hasRight,
                    "OMML→MathType numbered paragraph lost MathType's center/right tab stops.");
            }
        }
        finally
        {
            Release(tab);
            Release(tabs);
            Release(format);
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
        }

        AssertEqual(
            expectedNumberedCount,
            numberedCount,
            "OMML→MathType returned the wrong number of numbered MTPlaceRef tab paragraphs.");
    }

    private static int CountPureConversionOrdinaryEmptyParagraphs(
        Word.Document document)
    {
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? range = null;
        Word.OMaths? maths = null;
        Word.InlineShapes? shapes = null;
        Word.Fields? fields = null;
        Word.Frames? frames = null;
        var count = 0;
        try
        {
            paragraphs = document.Paragraphs;
            for (var index = 1;
                 index <= paragraphs.Count;
                 index++)
            {
                Release(frames);
                frames = null;
                Release(fields);
                fields = null;
                Release(shapes);
                shapes = null;
                Release(maths);
                maths = null;
                Release(range);
                range = null;
                Release(paragraph);
                paragraph = paragraphs[index];
                range = paragraph.Range.Duplicate;
                if (Convert.ToBoolean(
                        range.get_Information(
                            Word.WdInformation.wdWithInTable)))
                    continue;

                maths = range.OMaths;
                shapes = range.InlineShapes;
                fields = range.Fields;
                frames = range.Frames;
                var text =
                    (range.Text ?? string.Empty)
                    .Trim(
                        '\r',
                        '\n',
                        '\t',
                        '\v',
                        '\a',
                        ' ');
                if (text.Length == 0
                    && maths.Count == 0
                    && shapes.Count == 0
                    && fields.Count == 0
                    && frames.Count == 0)
                    count++;
            }
            return count;
        }
        finally
        {
            Release(frames);
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }
}
