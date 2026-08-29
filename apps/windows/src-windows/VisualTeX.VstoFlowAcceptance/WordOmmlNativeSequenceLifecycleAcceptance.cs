using System.Text.RegularExpressions;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlNativeSequenceLifecycleAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            RunContinuousNativeSequenceLifecycle(
                application,
                Path.Combine(
                    artifactRoot,
                    "word-omml-native-seq-middle-insert-delete.docx"));
            RunHeadingNativeSequenceLifecycle(
                application,
                Path.Combine(
                    artifactRoot,
                    "word-omml-native-seq-heading-2-3.docx"));
            RunNumberedNativeSequenceCopyLifecycle(
                application,
                Path.Combine(
                    artifactRoot,
                    "word-omml-native-seq-copy-rekey.docx"));
            Console.WriteLine(
                "Word native #(SEQ) lifecycle acceptance passed: middle insertion/deletion renumbered 1→2→3→4 and back to 1→2→3, ordinary and double-clickable body REF fields followed and navigated to the same mathematical VTEqNum target, Heading 1 numbering reached 2.3, and a pasted numbered OMath received an independent FormulaId/VTEqNum identity without Shape/Table hosts.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunContinuousNativeSequenceLifecycle(
        Word.Application application,
        string documentPath)
    {
        Word.Document? document = null;
        Word.Field? externalReference = null;
        Word.Field? navigableReference = null;
        Word.Range? secondFormulaRange = null;
        Word.Paragraphs? secondParagraphs = null;
        Word.Paragraph? secondParagraph = null;
        Word.Range? secondParagraphRange = null;
        Word.Range? insertedFormulaRange = null;
        try
        {
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);
            var formulaIds = new List<string>();
            for (var index = 1; index <= 3; index++)
            {
                var formulaId = Guid.NewGuid().ToString("D");
                InsertLifecycleOmml(
                    application,
                    document,
                    service,
                    formulaId,
                    $"x_{{{index}}}={index}",
                    LifecycleMathMl(index));
                formulaIds.Add(formulaId);
            }

            UpdateAllDocumentFields(document);
            AssertNativeSequenceNumbers(
                document,
                formulaIds,
                new[] { "1", "2", "3" },
                "initial continuous sequence");
            externalReference = InsertExternalEquationReference(
                document,
                formulaIds[2]);
            AssertExternalReferenceText(
                externalReference,
                "3",
                "initial continuous body REF");
            navigableReference = InsertNavigableEquationReference(
                application,
                document,
                formulaIds[2]);
            AssertNavigableEquationReference(
                application,
                document,
                navigableReference,
                formulaIds[2],
                "3",
                "initial double-clickable body reference");

            var secondMetadata = WordOmmlFormulaStore.TryRead(
                    document,
                    formulaIds[1])
                ?? throw new InvalidDataException(
                    "The second continuous formula lost its metadata before middle insertion.");
            secondFormulaRange = WordOmmlFormulaStore
                .GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    formulaIds[1],
                    secondMetadata);
            secondParagraphs = secondFormulaRange.Paragraphs;
            AssertEqual(1, secondParagraphs.Count,
                "The second formula does not occupy one paragraph before middle insertion.");
            secondParagraph = secondParagraphs[1];
            secondParagraphRange = secondParagraph.Range;
            var insertionPosition = secondParagraphRange.Start;
            secondParagraphRange.InsertParagraphBefore();
            application.Selection.SetRange(insertionPosition, insertionPosition);

            var insertedFormulaId = Guid.NewGuid().ToString("D");
            var insertedSession = CreateOmmlNumberingSession(
                document,
                insertedFormulaId,
                mode: "create",
                sourceObjectId: WordRangeReference(
                    insertionPosition,
                    insertionPosition),
                latex: @"x_{1.5}=15",
                originalMetadata: null);
            service.InsertOmml(insertedSession, LifecycleMathMl(15));

            var insertedOrder = new[]
            {
                formulaIds[0],
                insertedFormulaId,
                formulaIds[1],
                formulaIds[2],
            };
            var fieldCodesBeforeF9 = ReadNativeSequenceCodes(
                document,
                insertedOrder);
            UpdateAllDocumentFields(document);
            document.Save();
            AssertNativeSequenceCodesUnchanged(
                document,
                insertedOrder,
                fieldCodesBeforeF9,
                "middle insertion F9");
            AssertNativeSequenceNumbers(
                document,
                insertedOrder,
                new[] { "1", "2", "3", "4" },
                "continuous sequence after middle insertion");
            externalReference.Update();
            AssertExternalReferenceText(
                externalReference,
                "4",
                "body REF after middle insertion");
            WordEquationNumbering.UpdateNativeCrossReferences(document);
            Release(navigableReference);
            navigableReference = FindNavigableEquationReference(
                    document,
                    formulaIds[2])
                ?? throw new InvalidDataException(
                    "Middle insertion lost the double-clickable body reference.");
            AssertNavigableEquationReference(
                application,
                document,
                navigableReference,
                formulaIds[2],
                "4",
                "double-clickable body reference after middle insertion");

            var insertedMetadata = WordOmmlFormulaStore.TryRead(
                    document,
                    insertedFormulaId)
                ?? throw new InvalidDataException(
                    "The middle formula lost its metadata before deletion.");
            insertedFormulaRange = WordOmmlFormulaStore
                .GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    insertedFormulaId,
                    insertedMetadata);
            insertedFormulaRange.Select();
            AssertEqual(
                insertedFormulaId,
                service.DeleteSelectedFormula(),
                "DeleteSelectedFormula removed the wrong native #(SEQ) formula.");
            Release(insertedFormulaRange);
            insertedFormulaRange = null;

            var remainingCodesBeforeF9 = ReadNativeSequenceCodes(
                document,
                formulaIds);
            UpdateAllDocumentFields(document);
            AssertNativeSequenceCodesUnchanged(
                document,
                formulaIds,
                remainingCodesBeforeF9,
                "middle deletion F9");
            AssertNativeSequenceNumbers(
                document,
                formulaIds,
                new[] { "1", "2", "3" },
                "continuous sequence after middle deletion");
            externalReference.Update();
            AssertExternalReferenceText(
                externalReference,
                "3",
                "body REF after middle deletion");
            WordEquationNumbering.UpdateNativeCrossReferences(document);
            Release(navigableReference);
            navigableReference = FindNavigableEquationReference(
                    document,
                    formulaIds[2])
                ?? throw new InvalidDataException(
                    "Middle deletion lost the double-clickable body reference.");
            AssertNavigableEquationReference(
                application,
                document,
                navigableReference,
                formulaIds[2],
                "3",
                "double-clickable body reference after middle deletion");
            AssertTrue(
                WordOmmlFormulaStore.TryRead(document, insertedFormulaId) is null,
                "Deleted middle formula metadata survived.");
            AssertTrue(
                !document.Bookmarks.Exists(
                    WordEquationNumbering.NativeNumberBookmarkName(
                        insertedFormulaId)),
                "Deleted middle formula VTEqNum bookmark survived.");
            AssertEqual(0, document.Shapes.Count,
                "Continuous native sequence lifecycle created a Shape.");
            AssertEqual(0, document.Tables.Count,
                "Continuous native sequence lifecycle created a table.");

            document.Save();
            Release(navigableReference);
            navigableReference = null;
            Release(externalReference);
            externalReference = null;
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            UpdateAllDocumentFields(document);
            AssertNativeSequenceNumbers(
                document,
                formulaIds,
                new[] { "1", "2", "3" },
                "continuous sequence save/reopen");
            externalReference = FindExternalEquationReference(
                    document,
                    formulaIds[2])
                ?? throw new InvalidDataException(
                    "Continuous sequence save/reopen lost the external body REF.");
            externalReference.Update();
            AssertExternalReferenceText(
                externalReference,
                "3",
                "continuous body REF save/reopen");
            navigableReference = FindNavigableEquationReference(
                    document,
                    formulaIds[2])
                ?? throw new InvalidDataException(
                    "Continuous sequence save/reopen lost the double-clickable body reference.");
            WordEquationNumbering.UpdateNativeCrossReferences(document);
            AssertNavigableEquationReference(
                application,
                document,
                navigableReference,
                formulaIds[2],
                "3",
                "double-clickable body reference save/reopen");
            AssertEqual(0, document.Shapes.Count,
                "Continuous sequence save/reopen recreated a Shape.");
            AssertEqual(0, document.Tables.Count,
                "Continuous sequence save/reopen recreated a table.");
            document.Save();
        }
        finally
        {
            Release(insertedFormulaRange);
            Release(secondParagraphRange);
            Release(secondParagraph);
            Release(secondParagraphs);
            Release(secondFormulaRange);
            Release(navigableReference);
            Release(externalReference);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            ForceComCleanup();
        }
    }

    private static void RunHeadingNativeSequenceLifecycle(
        Word.Application application,
        string documentPath)
    {
        Word.Document? document = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? heading = null;
        Word.Range? headingRange = null;
        try
        {
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();
            document.Content.Text = "2 Native sequence chapter\r";
            paragraphs = document.Paragraphs;
            heading = paragraphs[1];
            headingRange = heading.Range;
            object headingStyle = Word.WdBuiltinStyle.wdStyleHeading1;
            headingRange.set_Style(ref headingStyle);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);

            var service = new WordFormulaService(application);
            var formulaIds = new List<string>();
            for (var index = 1; index <= 3; index++)
            {
                application.Selection.EndKey(Word.WdUnits.wdStory);
                var formulaId = Guid.NewGuid().ToString("D");
                InsertLifecycleOmml(
                    application,
                    document,
                    service,
                    formulaId,
                    $"h_{{{index}}}={index}",
                    LifecycleMathMl(index + 20));
                formulaIds.Add(formulaId);
            }

            var codesBeforeF9 = ReadNativeSequenceCodes(document, formulaIds);
            AssertTrue(codesBeforeF9.Values.All(code => Regex.IsMatch(
                    code,
                    @"\\s\s+1\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)),
                "Heading-numbered native SEQ fields were not created with \\s 1.");
            AssertTrue(codesBeforeF9.Values.All(code => !Regex.IsMatch(
                    code,
                    @"\\r\s+\d+\b",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)),
                "Heading-numbered native SEQ fields were frozen with a \\r ordinal.");

            UpdateAllDocumentFields(document);
            AssertNativeSequenceCodesUnchanged(
                document,
                formulaIds,
                codesBeforeF9,
                "heading F9");
            AssertNativeSequenceNumbers(
                document,
                formulaIds,
                new[] { "2.1", "2.2", "2.3" },
                "Heading 1 native sequence");
            AssertEqual(0, document.Shapes.Count,
                "Heading native sequence created a Shape.");
            AssertEqual(0, document.Tables.Count,
                "Heading native sequence created a table.");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            UpdateAllDocumentFields(document);
            AssertNativeSequenceNumbers(
                document,
                formulaIds,
                new[] { "2.1", "2.2", "2.3" },
                "Heading 1 native sequence save/reopen");
            AssertNativeSequenceCodesUnchanged(
                document,
                formulaIds,
                codesBeforeF9,
                "heading save/reopen F9");
            AssertEqual(0, document.Shapes.Count,
                "Heading sequence save/reopen recreated a Shape.");
            AssertEqual(0, document.Tables.Count,
                "Heading sequence save/reopen recreated a table.");
            document.Save();
        }
        finally
        {
            Release(headingRange);
            Release(heading);
            Release(paragraphs);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            ForceComCleanup();
        }
    }

    private static void RunNumberedNativeSequenceCopyLifecycle(
        Word.Application application,
        string documentPath)
    {
        Word.Document? document = null;
        Word.Range? sourceRange = null;
        Word.Range? pastedRange = null;
        Word.Field? sourceReference = null;
        Word.Field? copyReference = null;
        try
        {
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);
            var sourceFormulaId = Guid.NewGuid().ToString("D");
            var secondFormulaId = Guid.NewGuid().ToString("D");
            InsertLifecycleOmml(
                application,
                document,
                service,
                sourceFormulaId,
                @"c_1=1",
                LifecycleMathMl(31));
            InsertLifecycleOmml(
                application,
                document,
                service,
                secondFormulaId,
                @"c_2=2",
                LifecycleMathMl(32));
            UpdateAllDocumentFields(document);
            AssertNativeSequenceNumbers(
                document,
                new[] { sourceFormulaId, secondFormulaId },
                new[] { "1", "2" },
                "numbered copy source fixture");

            var sourceMetadata = WordOmmlFormulaStore.TryRead(
                    document,
                    sourceFormulaId)
                ?? throw new InvalidDataException(
                    "The numbered copy source lost its metadata.");
            sourceRange = WordOmmlFormulaStore
                .GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    sourceFormulaId,
                    sourceMetadata);
            sourceRange.Copy();
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            application.Selection.Paste();

            var pastedMathIndex = document.OMaths.Count;
            AssertEqual(3, pastedMathIndex,
                "Pasting a numbered OMath did not create exactly one copied equation.");
            pastedRange = GetOmmlRangeByIndex(document, pastedMathIndex);
            pastedRange.Select();
            var copiedSelection = service.ReadSelection();
            AssertTrue(copiedSelection.Metadata is not null,
                "The pasted numbered OMath was not adopted as VisualTeX metadata.");
            AssertTrue(copiedSelection.Metadata!.Numbered,
                "The pasted native #(SEQ) OMath was adopted as unnumbered.");
            AssertEqual("block", copiedSelection.Metadata.DisplayMode,
                "The pasted numbered OMath lost block display mode.");
            AssertEqual(FormulaOleContract.WordOmmlMode, copiedSelection.ObjectMode,
                "The pasted numbered OMath lost Word OMML mode.");
            var copyFormulaId = copiedSelection.FormulaId
                ?? throw new InvalidDataException(
                    "The pasted numbered OMath did not receive a FormulaId.");
            AssertTrue(!string.Equals(
                    sourceFormulaId,
                    copyFormulaId,
                    StringComparison.OrdinalIgnoreCase),
                "The pasted numbered OMath reused the source FormulaId.");
            AssertTrue(!string.Equals(
                    secondFormulaId,
                    copyFormulaId,
                    StringComparison.OrdinalIgnoreCase),
                "The pasted numbered OMath reused another formula's FormulaId.");

            UpdateAllDocumentFields(document);
            var orderedFormulaIds = new[]
            {
                sourceFormulaId,
                secondFormulaId,
                copyFormulaId,
            };
            AssertNativeSequenceNumbers(
                document,
                orderedFormulaIds,
                new[] { "1", "2", "3" },
                "numbered copy after FormulaId rekey");
            AssertTrue(document.Bookmarks.Exists(
                    WordEquationNumbering.NativeNumberBookmarkName(sourceFormulaId)),
                "Rekeying the copy removed the source VTEqNum bookmark.");
            AssertTrue(document.Bookmarks.Exists(
                    WordEquationNumbering.NativeNumberBookmarkName(copyFormulaId)),
                "Rekeying the copy did not create an independent VTEqNum bookmark.");

            sourceReference = InsertExternalEquationReference(
                document,
                sourceFormulaId);
            copyReference = InsertExternalEquationReference(
                document,
                copyFormulaId);
            AssertExternalReferenceText(
                sourceReference,
                "1",
                "numbered copy source body REF");
            AssertExternalReferenceText(
                copyReference,
                "3",
                "numbered copy body REF");

            var copyMetadata = WordOmmlFormulaStore.TryRead(
                    document,
                    copyFormulaId)
                ?? throw new InvalidDataException(
                    "The rekeyed numbered copy lost its metadata before editing.");
            var copyRange = WordOmmlFormulaStore
                .GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    copyFormulaId,
                    copyMetadata);
            try
            {
                var editSession = CreateOmmlNumberingSession(
                    document,
                    copyFormulaId,
                    mode: "edit",
                    sourceObjectId: WordRangeReference(
                        copyRange.Start,
                        copyRange.End),
                    latex: @"c_3=303",
                    originalMetadata: copyMetadata);
                service.ReplaceOmml(editSession, LifecycleMathMl(303));
            }
            finally { Release(copyRange); }

            UpdateAllDocumentFields(document);
            AssertNativeSequenceNumbers(
                document,
                orderedFormulaIds,
                new[] { "1", "2", "3" },
                "numbered copy after independent edit");
            sourceReference.Update();
            copyReference.Update();
            AssertExternalReferenceText(
                sourceReference,
                "1",
                "numbered copy source REF after copy edit");
            AssertExternalReferenceText(
                copyReference,
                "3",
                "numbered copy REF after copy edit");
            var sourceAfterCopyEdit = WordOmmlFormulaStore.TryRead(
                    document,
                    sourceFormulaId)
                ?? throw new InvalidDataException(
                    "Editing the numbered copy removed source metadata.");
            AssertEqual(sourceFormulaId, sourceAfterCopyEdit.FormulaId,
                "Editing the numbered copy changed the source FormulaId.");
            AssertEqual(0, document.Shapes.Count,
                "Numbered copy/rekey lifecycle created a Shape.");
            AssertEqual(0, document.Tables.Count,
                "Numbered copy/rekey lifecycle created a table.");

            document.Save();
            Release(sourceReference);
            sourceReference = null;
            Release(copyReference);
            copyReference = null;
            Release(pastedRange);
            pastedRange = null;
            Release(sourceRange);
            sourceRange = null;
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            UpdateAllDocumentFields(document);
            AssertNativeSequenceNumbers(
                document,
                orderedFormulaIds,
                new[] { "1", "2", "3" },
                "numbered copy save/reopen");
            sourceReference = FindExternalEquationReference(
                    document,
                    sourceFormulaId)
                ?? throw new InvalidDataException(
                    "Numbered copy save/reopen lost the source body REF.");
            copyReference = FindExternalEquationReference(
                    document,
                    copyFormulaId)
                ?? throw new InvalidDataException(
                    "Numbered copy save/reopen lost the copy body REF.");
            sourceReference.Update();
            copyReference.Update();
            AssertExternalReferenceText(
                sourceReference,
                "1",
                "numbered copy source REF save/reopen");
            AssertExternalReferenceText(
                copyReference,
                "3",
                "numbered copy REF save/reopen");
            AssertEqual(0, document.Shapes.Count,
                "Numbered copy save/reopen recreated a Shape.");
            AssertEqual(0, document.Tables.Count,
                "Numbered copy save/reopen recreated a table.");
            document.Save();
        }
        finally
        {
            Release(copyReference);
            Release(sourceReference);
            Release(pastedRange);
            Release(sourceRange);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            ForceComCleanup();
        }
    }

    private static void InsertLifecycleOmml(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        string formulaId,
        string latex,
        string mathMl)
    {
        application.Selection.EndKey(Word.WdUnits.wdStory);
        var insertion = application.Selection.Range;
        try
        {
            var session = CreateOmmlNumberingSession(
                document,
                formulaId,
                mode: "create",
                sourceObjectId: WordRangeReference(
                    insertion.Start,
                    insertion.End),
                latex,
                originalMetadata: null);
            service.InsertOmml(session, mathMl);
        }
        finally { Release(insertion); }
    }

    private static string LifecycleMathMl(int value) =>
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
        + $"<msub><mi>x</mi><mn>{value}</mn></msub><mo>=</mo><mn>{value}</mn>"
        + "</math>";

    private static Word.Field InsertNavigableEquationReference(
        Word.Application application,
        Word.Document document,
        string formulaId)
    {
        Word.Selection? selection = null;
        try
        {
            document.Activate();
            selection = application.Selection;
            var insertion = Math.Max(document.Content.Start, document.Content.End - 1);
            selection.SetRange(insertion, insertion);
            selection.TypeParagraph();
            selection.TypeText("Navigable equation reference: ");
            WordEquationReferenceFields.InsertNavigableReference(
                document,
                selection,
                WordEquationNumbering.NativeNumberBookmarkName(formulaId),
                prefix: "(",
                suffix: ")");
            return FindNavigableEquationReference(document, formulaId)
                ?? throw new InvalidDataException(
                    "Word did not create the expected GOTOBUTTON + nested REF equation reference.");
        }
        finally { Release(selection); }
    }

    private static Word.Field? FindNavigableEquationReference(
        Word.Document document,
        string formulaId)
    {
        var targetBookmark = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
        Word.Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? field = null;
                Word.Range? outerCode = null;
                Word.Fields? nestedFields = null;
                try
                {
                    field = fields[index];
                    if (field.Type != Word.WdFieldType.wdFieldGoToButton)
                        continue;
                    outerCode = field.Code;
                    var outerTargetsBookmark = (outerCode.Text ?? string.Empty).IndexOf(
                        targetBookmark,
                        StringComparison.OrdinalIgnoreCase) >= 0;
                    nestedFields = outerCode.Fields;
                    var nestedTargetsBookmark = false;
                    for (var nestedIndex = 1;
                         nestedIndex <= nestedFields.Count;
                         nestedIndex++)
                    {
                        Word.Field? nested = null;
                        Word.Range? nestedCode = null;
                        try
                        {
                            nested = nestedFields[nestedIndex];
                            if (nested.Type != Word.WdFieldType.wdFieldRef)
                                continue;
                            nestedCode = nested.Code;
                            if ((nestedCode.Text ?? string.Empty).IndexOf(
                                    "REF " + targetBookmark,
                                    StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                nestedTargetsBookmark = true;
                                break;
                            }
                        }
                        finally
                        {
                            Release(nestedCode);
                            Release(nested);
                        }
                    }
                    if (!outerTargetsBookmark || !nestedTargetsBookmark)
                        continue;
                    var result = field;
                    field = null;
                    return result;
                }
                finally
                {
                    Release(nestedFields);
                    Release(outerCode);
                    Release(field);
                }
            }
            return null;
        }
        finally { Release(fields); }
    }

    private static void AssertNavigableEquationReference(
        Word.Application application,
        Word.Document document,
        Word.Field outerField,
        string formulaId,
        string expectedNumber,
        string context)
    {
        Word.Range? outerCode = null;
        Word.Fields? nestedFields = null;
        Word.Field? nestedReference = null;
        Word.Range? nestedCode = null;
        Word.Range? nestedResult = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? targetBookmark = null;
        Word.Range? targetRange = null;
        Word.Range? formulaRange = null;
        Word.Range? fieldResult = null;
        Word.Selection? selection = null;
        Word.Range? selectedRange = null;
        try
        {
            var targetName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            AssertEqual(Word.WdFieldType.wdFieldGoToButton, outerField.Type,
                context + ": outer field is not GOTOBUTTON.");
            outerCode = outerField.Code;
            AssertTrue((outerCode.Text ?? string.Empty).IndexOf(
                    targetName,
                    StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": GOTOBUTTON does not target VTEqNum.");
            nestedFields = outerCode.Fields;
            AssertEqual(1, nestedFields.Count,
                context + ": GOTOBUTTON does not contain exactly one nested REF.");
            nestedReference = nestedFields[1];
            AssertEqual(Word.WdFieldType.wdFieldRef, nestedReference.Type,
                context + ": nested field is not REF.");
            nestedCode = nestedReference.Code;
            AssertTrue((nestedCode.Text ?? string.Empty).IndexOf(
                    "REF " + targetName,
                    StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": nested REF targets the wrong bookmark.");
            nestedReference.Update();
            nestedResult = nestedReference.Result;
            AssertEqual(
                expectedNumber,
                NormalizeEquationNumberText(nestedResult.Text),
                context + ": nested REF result is stale.");

            bookmarks = document.Bookmarks;
            AssertTrue(bookmarks.Exists(targetName),
                context + ": VTEqNum target bookmark is missing.");
            targetBookmark = bookmarks[targetName];
            targetRange = targetBookmark.Range.Duplicate;
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    context + ": target formula metadata is missing.");
            formulaRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                formulaId,
                metadata);

            // Field.DoClick is Word's COM equivalent of activating a GOTOBUTTON by
            // the user's double-click. Begin on the rendered reference, invoke that
            // native action, then require the live Selection to land in this exact
            // formula's VTEqNum/OMath range rather than merely somewhere nearby.
            document.Activate();
            fieldResult = outerField.Result;
            fieldResult.Select();
            outerField.DoClick();
            System.Windows.Forms.Application.DoEvents();
            selection = application.Selection;
            selectedRange = selection.Range;
            var overlapsNumber = selectedRange.Start <= targetRange.End
                && selectedRange.End >= targetRange.Start;
            var insideFormula = selectedRange.Start >= formulaRange.Start
                && selectedRange.Start <= formulaRange.End;
            AssertTrue(overlapsNumber || insideFormula,
                context + $": GOTOBUTTON activation did not navigate to the target OMath; selection={selectedRange.Start}:{selectedRange.End}, number={targetRange.Start}:{targetRange.End}, formula={formulaRange.Start}:{formulaRange.End}.");
            Console.WriteLine(
                $"  {context}: nested REF='{expectedNumber}', DoClick selection={selectedRange.Start}:{selectedRange.End}, target={targetRange.Start}:{targetRange.End}.");
        }
        finally
        {
            Release(selectedRange);
            Release(selection);
            Release(fieldResult);
            Release(formulaRange);
            Release(targetRange);
            Release(targetBookmark);
            Release(bookmarks);
            Release(nestedResult);
            Release(nestedCode);
            Release(nestedReference);
            Release(nestedFields);
            Release(outerCode);
        }
    }

    private static void UpdateAllDocumentFields(Word.Document document)
    {
        Word.Fields? fields = null;
        try
        {
            fields = document.Fields;
            if (fields.Count > 0)
                fields.Update();
        }
        finally { Release(fields); }
    }

    private static void AssertNativeSequenceNumbers(
        Word.Document document,
        IReadOnlyList<string> formulaIds,
        IReadOnlyList<string> expectedNumbers,
        string context)
    {
        AssertEqual(expectedNumbers.Count, formulaIds.Count,
            context + ": expected number count does not match formula count.");
        for (var index = 0; index < formulaIds.Count; index++)
        {
            AssertOmmlTabNumberingHost(
                document,
                formulaIds[index],
                context + $" formula #{index + 1}",
                updateReference: false);
            Word.Range? numberRange = null;
            try
            {
                numberRange = WordEquationNumbering.FindVisibleEquationNumberTextRange(
                        document,
                        formulaIds[index])
                    ?? throw new InvalidDataException(
                        context + $": formula #{index + 1} lost VTEqNum.");
                AssertEqual(
                    expectedNumbers[index],
                    NormalizeEquationNumberText(numberRange.Text),
                    context + $": formula #{index + 1} has a stale number.");
            }
            finally { Release(numberRange); }
        }
    }

    private static IReadOnlyDictionary<string, string> ReadNativeSequenceCodes(
        Word.Document document,
        IReadOnlyList<string> formulaIds)
    {
        var result = new Dictionary<string, string>(
            StringComparer.OrdinalIgnoreCase);
        foreach (var formulaId in formulaIds)
        {
            Word.Bookmark? bookmark = null;
            Word.Range? equationRange = null;
            Word.Fields? fields = null;
            Word.Field? sequenceField = null;
            Word.Range? code = null;
            try
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                    ?? throw new InvalidDataException(
                        $"Formula {formulaId} lost its VTOMML bookmark while reading SEQ code.");
                equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                fields = equationRange.Fields;
                AssertEqual(1, fields.Count,
                    $"Formula {formulaId} does not own exactly one native SEQ field.");
                sequenceField = fields[1];
                code = sequenceField.Code;
                var normalized = Regex.Replace(
                    code.Text ?? string.Empty,
                    @"\s+",
                    " ").Trim();
                AssertTrue(WordEquationNumbering.IsVisualTeXSequenceFieldCode(normalized),
                    $"Formula {formulaId} field is not SEQ VisualTeXEquation.");
                result[formulaId] = normalized;
            }
            finally
            {
                Release(code);
                Release(sequenceField);
                Release(fields);
                Release(equationRange);
                Release(bookmark);
            }
        }
        return result;
    }

    private static void AssertNativeSequenceCodesUnchanged(
        Word.Document document,
        IReadOnlyList<string> formulaIds,
        IReadOnlyDictionary<string, string> expectedCodes,
        string context)
    {
        var actualCodes = ReadNativeSequenceCodes(document, formulaIds);
        foreach (var formulaId in formulaIds)
        {
            AssertTrue(expectedCodes.TryGetValue(formulaId, out var expected),
                context + $": no baseline SEQ code exists for {formulaId}.");
            AssertEqual(expected!, actualCodes[formulaId],
                context + $": F9 rewrote the mathematical SEQ field code for {formulaId}.");
        }
    }

    private static void AssertExternalReferenceText(
        Word.Field field,
        string expected,
        string context)
    {
        Word.Range? result = null;
        try
        {
            result = field.Result;
            AssertEqual(
                expected,
                NormalizeEquationNumberText(result.Text),
                context + ": external REF result is stale.");
        }
        finally { Release(result); }
    }
}
