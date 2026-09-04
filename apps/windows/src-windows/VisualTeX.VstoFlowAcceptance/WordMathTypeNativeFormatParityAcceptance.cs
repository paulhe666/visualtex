using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class NativeMathTypeSectionIdentity
    {
        internal int FieldStart { get; set; }
        internal int CodeStart { get; set; }
        internal int CodeEnd { get; set; }
        internal int ParagraphStart { get; set; }
        internal string Code { get; set; } = string.Empty;
    }

    private sealed class NativeMathTypeFormatParityCase
    {
        internal string FormatId { get; set; } = string.Empty;
        internal string[] ExpectedNumbers { get; set; } = Array.Empty<string>();
        internal string[] ExpectedVisibleSequences { get; set; } = Array.Empty<string>();
        internal string ExpectedSeparator { get; set; } = string.Empty;
    }

    private static void RunWordMathTypeNativeFormatParityAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var configuredFixture = Environment.GetEnvironmentVariable(
            "VISUALTEX_NATIVE_MATHTYPE_FIXTURE");
        if (string.IsNullOrWhiteSpace(configuredFixture))
            throw new InvalidDataException(
                "VISUALTEX_NATIVE_MATHTYPE_FIXTURE must point to a DOCX created by MathType's own left/right numbered-equation and Insert Reference commands.");
        var fixturePath = Path.GetFullPath(
            Environment.ExpandEnvironmentVariables(configuredFixture.Trim().Trim('"')));
        if (!File.Exists(fixturePath))
            throw new FileNotFoundException(
                "The native MathType numbering fixture does not exist.",
                fixturePath);

        var cases = new[]
        {
            new NativeMathTypeFormatParityCase
            {
                FormatId = EquationNumberFormat.ContinuousId,
                ExpectedNumbers = new[] { "(1)", "(2)" },
                ExpectedVisibleSequences = new[] { "MTEqn" },
                ExpectedSeparator = string.Empty,
            },
            new NativeMathTypeFormatParityCase
            {
                FormatId = EquationNumberFormat.Heading1DotId,
                ExpectedNumbers = new[] { "(1.1)", "(1.2)" },
                ExpectedVisibleSequences = new[] { "MTChap", "MTEqn" },
                ExpectedSeparator = ".",
            },
            new NativeMathTypeFormatParityCase
            {
                FormatId = EquationNumberFormat.Heading1DashId,
                ExpectedNumbers = new[] { "(1-1)", "(1-2)" },
                ExpectedVisibleSequences = new[] { "MTChap", "MTEqn" },
                ExpectedSeparator = "-",
            },
            new NativeMathTypeFormatParityCase
            {
                FormatId = EquationNumberFormat.Heading2DotId,
                ExpectedNumbers = new[] { "(1.1.1)", "(1.1.2)" },
                ExpectedVisibleSequences = new[] { "MTChap", "MTSec", "MTEqn" },
                ExpectedSeparator = ".",
            },
            new NativeMathTypeFormatParityCase
            {
                FormatId = EquationNumberFormat.Heading2DashId,
                ExpectedNumbers = new[] { "(1.1-1)", "(1.1-2)" },
                ExpectedVisibleSequences = new[] { "MTChap", "MTSec", "MTEqn" },
                ExpectedSeparator = "-",
            },
        };

        var mathTypeProcessIdsBefore = CaptureMathTypeProcessIds();
        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            AssertNoNewMathTypeProcesses(
                mathTypeProcessIdsBefore,
                "Starting malformed-native-structure guard acceptance");
            try
            {
                RunNativeMathTypeMalformedFormatGuard(
                    application,
                    fixturePath,
                    artifactRoot);
            }
            finally
            {
                try { QuitWordApplicationIfOwned(application); } catch { }
                Release(application);
                application = null;
                ForceComCleanup();
                Thread.Sleep(900);
            }
            AssertNoNewMathTypeProcesses(
                mathTypeProcessIdsBefore,
                "VisualTeX malformed MathType structure guard");

            // Isolate every save/reopen case in its own WINWORD.EXE. Office COM
            // automation can terminate an otherwise healthy hidden Word process
            // after several document lifecycles; per-case isolation keeps that
            // process-lifetime behavior out of the numbering result.
            foreach (var testCase in cases)
            {
                application = CreateWordApplication(visible: false);
                AssertNoNewMathTypeProcesses(
                    mathTypeProcessIdsBefore,
                    $"Starting isolated Word for '{testCase.FormatId}'");
                try
                {
                    RunNativeMathTypeFormatParityCase(
                        application,
                        fixturePath,
                        artifactRoot,
                        testCase);
                }
                finally
                {
                    try { QuitWordApplicationIfOwned(application); } catch { }
                    Release(application);
                    application = null;
                    ForceComCleanup();
                    Thread.Sleep(900);
                }
                AssertNoNewMathTypeProcesses(
                    mathTypeProcessIdsBefore,
                    $"VisualTeX formatting '{testCase.FormatId}'");
            }

            application = CreateWordApplication(visible: false);
            AssertNoNewMathTypeProcesses(
                mathTypeProcessIdsBefore,
                "Starting the repeated-format Word acceptance process");
            RunNativeMathTypeFormatParityCycle(
                application,
                fixturePath,
                artifactRoot,
                cases);
            AssertNoNewMathTypeProcesses(
                mathTypeProcessIdsBefore,
                "VisualTeX repeated MathType number-format cycle");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }

        AssertNoNewMathTypeProcesses(
            mathTypeProcessIdsBefore,
            "Completing the VisualTeX-only MathType numbering acceptance");
        Console.WriteLine(
            "[MathType native format parity] All VisualTeX presets preserved MathType's native section state, left/right MTDisplayEquation layout, nested MTPlaceRef ownership, ZEqnNum bookmark identity and GOTOBUTTON/REF reference across save/reopen without starting MathType.");
    }

    private static HashSet<int> CaptureMathTypeProcessIds()
    {
        var result = new HashSet<int>();
        foreach (var process in Process.GetProcesses())
        {
            try
            {
                if (process.ProcessName.StartsWith(
                        "MathType",
                        StringComparison.OrdinalIgnoreCase))
                    result.Add(process.Id);
            }
            catch
            {
                // A process may exit between enumeration and reading its name.
            }
            finally { process.Dispose(); }
        }
        return result;
    }

    private static void AssertNoNewMathTypeProcesses(
        ISet<int> processIdsBefore,
        string context)
    {
        var current = CaptureMathTypeProcessIds();
        var started = current
            .Where(processId => !processIdsBefore.Contains(processId))
            .OrderBy(processId => processId)
            .ToArray();
        AssertEqual(
            0,
            started.Length,
            context + " unexpectedly started MathType process(es): "
                + string.Join(", ", started) + ".");
    }

    private static void RunNativeMathTypeFormatParityCase(
        Word.Application application,
        string fixturePath,
        string artifactRoot,
        NativeMathTypeFormatParityCase testCase)
    {
        var safeFormat = Regex.Replace(testCase.FormatId, @"[^A-Za-z0-9_.-]", "-");
        var workingPath = Path.Combine(artifactRoot, $"native-{safeFormat}-working.docx");
        var outputPath = Path.Combine(artifactRoot, $"native-{safeFormat}-after.docx");
        File.Copy(fixturePath, workingPath, overwrite: true);

        Word.Document? document = null;
        try
        {
            document = application.Documents.Open(
                workingPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();

            var sectionIdentity = ReadSingleNativeMathTypeSectionIdentity(document);
            var bookmarkName = ReadSingleNativeMathTypeBookmarkName(document);
            var referenceCode = ReadSingleNativeMathTypeReferenceCode(document);
            var nativeTargets = MathTypeEquationReferences.GetTargets(document);
            AssertEqual(2, nativeTargets.Count,
                "The MathType-native fixture must expose exactly two numbered equations.");
            var nativeNumbers = nativeTargets
                .Select(target => target.NumberText.Trim())
                .ToArray();
            var nativeFormat = ReadNativeMathTypePlaceRefFormat(document);
            AssertNativeMathTypeParityStructure(
                document,
                nativeNumbers,
                nativeFormat.VisibleSequences,
                nativeFormat.Separator,
                sectionIdentity,
                bookmarkName,
                referenceCode,
                $"Native fixture before '{testCase.FormatId}'");
            DumpNativeMathTypeParityStructure(
                document,
                Path.Combine(artifactRoot, $"native-{safeFormat}-before.txt"));

            Console.WriteLine(
                $"[MathType native format parity] applying {testCase.FormatId} to MathType-native fixture...");
            var changed = MathTypeEquationNumbering.SetEquationNumberFormat(
                document,
                testCase.FormatId);
            AssertEqual(
                2,
                changed,
                $"Format '{testCase.FormatId}' did not rewrite exactly two native MTPlaceRef fields.");

            DumpNativeMathTypeParityStructure(
                document,
                Path.Combine(artifactRoot, $"native-{safeFormat}-after-live.txt"));
            AssertNativeMathTypeParityStructure(
                document,
                testCase.ExpectedNumbers,
                testCase.ExpectedVisibleSequences,
                testCase.ExpectedSeparator,
                sectionIdentity,
                bookmarkName,
                referenceCode,
                $"Native fixture after '{testCase.FormatId}'");

            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            document = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            AssertNativeMathTypeParityStructure(
                document,
                testCase.ExpectedNumbers,
                testCase.ExpectedVisibleSequences,
                testCase.ExpectedSeparator,
                sectionIdentity,
                bookmarkName,
                referenceCode,
                $"Reopened native fixture after '{testCase.FormatId}'");
            DumpNativeMathTypeParityStructure(
                document,
                Path.Combine(artifactRoot, $"native-{safeFormat}-after-reopen.txt"));
            Console.WriteLine(
                $"[MathType native format parity] {testCase.FormatId} passed live and after reopen.");
        }
        catch
        {
            if (document is not null)
            {
                try
                {
                    DumpNativeMathTypeParityStructure(
                        document,
                        Path.Combine(artifactRoot, $"native-{safeFormat}-failure.txt"));
                    document.SaveAs2(
                        Path.Combine(artifactRoot, $"native-{safeFormat}-failure.docx"),
                        Word.WdSaveFormat.wdFormatXMLDocument);
                }
                catch { }
            }
            throw;
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }

    private static void RunNativeMathTypeFormatParityCycle(
        Word.Application application,
        string fixturePath,
        string artifactRoot,
        IReadOnlyList<NativeMathTypeFormatParityCase> cases)
    {
        var workingPath = Path.Combine(artifactRoot, "native-format-cycle-working.docx");
        var outputPath = Path.Combine(artifactRoot, "native-format-cycle-after.docx");
        File.Copy(fixturePath, workingPath, overwrite: true);

        Word.Document? document = null;
        try
        {
            document = application.Documents.Open(
                workingPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();

            var sectionIdentity = ReadSingleNativeMathTypeSectionIdentity(document);
            var bookmarkName = ReadSingleNativeMathTypeBookmarkName(document);
            var referenceCode = ReadSingleNativeMathTypeReferenceCode(document);
            var sequence = cases
                .Concat(new[] { cases[0], cases[cases.Count - 1], cases[1] })
                .ToArray();

            for (var step = 0; step < sequence.Length; step++)
            {
                var testCase = sequence[step];
                try
                {
                    var changed = MathTypeEquationNumbering.SetEquationNumberFormat(
                        document,
                        testCase.FormatId);
                    AssertEqual(
                        2,
                        changed,
                        $"Repeated format step {step + 1} ('{testCase.FormatId}') did not rewrite exactly two native MTPlaceRef fields.");
                    AssertNativeMathTypeParityStructure(
                        document,
                        testCase.ExpectedNumbers,
                        testCase.ExpectedVisibleSequences,
                        testCase.ExpectedSeparator,
                        sectionIdentity,
                        bookmarkName,
                        referenceCode,
                        $"Repeated native format step {step + 1} ('{testCase.FormatId}')");
                    DumpNativeMathTypeParityStructure(
                        document,
                        Path.Combine(
                            artifactRoot,
                            $"native-format-cycle-{step + 1:D2}-{testCase.FormatId}.txt"));
                }
                catch (Exception error)
                {
                    DumpNativeMathTypeParityStructure(
                        document,
                        Path.Combine(
                            artifactRoot,
                            $"native-format-cycle-{step + 1:D2}-{testCase.FormatId}-failure.txt"));
                    Console.Error.WriteLine(
                        $"[MathType native format parity] repeated step {step + 1} ('{testCase.FormatId}') failed: {error}");
                    if (error is AggregateException aggregate)
                    {
                        var flattened = aggregate.Flatten();
                        for (var index = 0; index < flattened.InnerExceptions.Count; index++)
                            Console.Error.WriteLine(
                                $"[MathType native format parity] inner {index + 1}: {flattened.InnerExceptions[index]}");
                    }
                    throw;
                }
            }

            var finalCase = sequence[sequence.Length - 1];
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            document = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            AssertNativeMathTypeParityStructure(
                document,
                finalCase.ExpectedNumbers,
                finalCase.ExpectedVisibleSequences,
                finalCase.ExpectedSeparator,
                sectionIdentity,
                bookmarkName,
                referenceCode,
                "Reopened repeated native MathType number-format cycle");
            DumpNativeMathTypeParityStructure(
                document,
                Path.Combine(artifactRoot, "native-format-cycle-after-reopen.txt"));
            Console.WriteLine(
                $"[MathType native format parity] repeated cycle passed {sequence.Length} consecutive format changes and save/reopen.");
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }

    private static void RunNativeMathTypeMalformedFormatGuard(
        Word.Application application,
        string fixturePath,
        string artifactRoot)
    {
        var workingPath = Path.Combine(
            artifactRoot,
            "native-malformed-preflight-working.docx");
        File.Copy(fixturePath, workingPath, overwrite: true);

        Word.Document? document = null;
        Word.InlineShape? removedShape = null;
        try
        {
            document = application.Documents.Open(
                workingPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            AssertEqual(2, document.InlineShapes.Count,
                "Malformed-preflight fixture did not start with two MathType OLE objects.");
            AssertEqual(2, MathTypeEquationNumbering.CountPlaceRefFields(document),
                "Malformed-preflight fixture did not start with two MTPlaceRef fields.");

            // Reproduce the dangerous class of existing-document damage without
            // involving MathType: remove the first Equation.DSMT4 object while
            // leaving its native MTPlaceRef tree in place. SetEquationNumberFormat
            // must validate every target before writing any field, so the healthy
            // second equation must remain byte-for-byte unchanged when the orphan
            // owner is discovered later in the descending preflight.
            removedShape = document.InlineShapes[1];
            removedShape.Delete();
            Release(removedShape);
            removedShape = null;

            var paragraphsBefore = document.Paragraphs.Count;
            var shapesBefore = document.InlineShapes.Count;
            var placeRefCodesBefore = ReadNativeMathTypePlaceRefCodes(document);
            var referenceCodeBefore = ReadSingleNativeMathTypeReferenceCode(document);
            AssertEqual(2, placeRefCodesBefore.Length,
                "Malformed-preflight setup unexpectedly lost an MTPlaceRef field.");

            var rejected = false;
            Exception? rejection = null;
            try
            {
                _ = MathTypeEquationNumbering.SetEquationNumberFormat(
                    document,
                    EquationNumberFormat.Heading1DashId);
            }
            catch (Exception error)
            {
                rejected = true;
                rejection = error;
            }

            AssertTrue(rejected,
                "Malformed MathType structure was accepted instead of being rejected before rewrite.");
            AssertTrue(rejection is InvalidDataException,
                "Malformed MathType structure failed with an unexpected exception type: "
                    + rejection?.GetType().FullName + ".");
            AssertEqual(paragraphsBefore, document.Paragraphs.Count,
                "Rejected malformed MathType format operation changed paragraph count.");
            AssertEqual(shapesBefore, document.InlineShapes.Count,
                "Rejected malformed MathType format operation changed OLE count.");
            AssertEqual(2, MathTypeEquationNumbering.CountPlaceRefFields(document),
                "Rejected malformed MathType format operation changed MTPlaceRef count.");
            AssertEqual(
                string.Join("\n", placeRefCodesBefore),
                string.Join("\n", ReadNativeMathTypePlaceRefCodes(document)),
                "Rejected malformed MathType format operation partially rewrote a healthy MTPlaceRef field.");
            AssertEqual(referenceCodeBefore, ReadSingleNativeMathTypeReferenceCode(document),
                "Rejected malformed MathType format operation changed the native GOTOBUTTON/REF reference.");

            DumpNativeMathTypeParityStructure(
                document,
                Path.Combine(artifactRoot, "native-malformed-preflight-after-rejection.txt"));
            Console.WriteLine(
                "[MathType native format parity] malformed preflight guard rejected an orphan MTPlaceRef owner without partially rewriting the remaining native equation.");
        }
        finally
        {
            Release(removedShape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }

    private static string[] ReadNativeMathTypePlaceRefCodes(Word.Document document)
    {
        var result = new List<string>();
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                if (!MathTypeEquationReferences.IsMathTypePlaceRefCode(code.Text))
                    continue;
                result.Add(NormalizeNativeMathTypeFieldCode(code.Text));
            }
            return result.ToArray();
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static (string[] VisibleSequences, string Separator)
        ReadNativeMathTypePlaceRefFormat(Word.Document document)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Fields? nestedFields = null;
        Word.Field? nested = null;
        Word.Range? nestedCode = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                if (!MathTypeEquationReferences.IsMathTypePlaceRefCode(code.Text))
                    continue;

                nestedFields = code.Fields;
                var visibleSequences = new List<string>();
                for (var nestedIndex = 1; nestedIndex <= nestedFields.Count; nestedIndex++)
                {
                    Release(nestedCode); nestedCode = null;
                    Release(nested); nested = nestedFields[nestedIndex];
                    nestedCode = nested.Code;
                    var normalized = NormalizeNativeMathTypeFieldCode(nestedCode.Text);
                    if (normalized.IndexOf("\\h", StringComparison.Ordinal) >= 0)
                        continue;
                    var match = Regex.Match(
                        normalized,
                        @"^SEQ\s+(MTChap|MTSec|MTEqn)\s+\\c\b",
                        RegexOptions.IgnoreCase);
                    if (match.Success)
                        visibleSequences.Add(match.Groups[1].Value);
                }
                AssertTrue(visibleSequences.Count > 0,
                    "The MathType-native fixture has no visible MTPlaceRef sequence.");

                var outerCode = code.Text ?? string.Empty;
                var separator = outerCode.IndexOf('-') >= 0
                    ? "-"
                    : outerCode.IndexOf('.') >= 0
                        ? "."
                        : string.Empty;
                return (visibleSequences.ToArray(), separator);
            }
        }
        finally
        {
            Release(nestedCode);
            Release(nested);
            Release(nestedFields);
            Release(code);
            Release(field);
            Release(fields);
        }
        throw new InvalidDataException(
            "The MathType-native fixture has no MTPlaceRef field.");
    }

    private static NativeMathTypeSectionIdentity ReadSingleNativeMathTypeSectionIdentity(
        Word.Document document)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? result = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        NativeMathTypeSectionIdentity? identity = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(paragraphRange); paragraphRange = null;
                Release(paragraph); paragraph = null;
                Release(paragraphs); paragraphs = null;
                Release(result); result = null;
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                if (!MathTypeEquationNumbering.IsMathTypeSectionBreakCode(code.Text))
                    continue;
                AssertTrue(identity is null,
                    "The MathType-native fixture contains more than one MTEditEquationSection2 state field.");
                result = field.Result;
                paragraphs = code.Paragraphs;
                AssertEqual(1, paragraphs.Count,
                    "The native MTEditEquationSection2 field does not occupy exactly one paragraph.");
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
                identity = new NativeMathTypeSectionIdentity
                {
                    FieldStart = code.Start - 1,
                    CodeStart = code.Start,
                    CodeEnd = code.End,
                    ParagraphStart = paragraphRange.Start,
                    Code = NormalizeNativeMathTypeFieldCode(code.Text),
                };
            }
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
        }
        return identity
            ?? throw new InvalidDataException(
                "The MathType-native fixture has no MTEditEquationSection2 state field.");
    }

    private static string ReadSingleNativeMathTypeBookmarkName(Word.Document document)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        string? found = null;
        try
        {
            bookmarks = document.Bookmarks;
            var showHidden = bookmarks.ShowHidden;
            bookmarks.ShowHidden = true;
            try
            {
                for (var index = 1; index <= bookmarks.Count; index++)
                {
                    Release(bookmark); bookmark = bookmarks[index];
                    if (!bookmark.Name.StartsWith("ZEqnNum", StringComparison.OrdinalIgnoreCase))
                        continue;
                    AssertTrue(found is null,
                        "The MathType-native fixture contains more than one ZEqnNum bookmark.");
                    found = bookmark.Name;
                }
            }
            finally { bookmarks.ShowHidden = showHidden; }
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
        return found
            ?? throw new InvalidDataException(
                "The MathType-native fixture has no ZEqnNum reference bookmark.");
    }

    private static string ReadSingleNativeMathTypeReferenceCode(Word.Document document)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        string? found = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                var normalized = NormalizeNativeMathTypeFieldCode(code.Text);
                if (!normalized.StartsWith("GOTOBUTTON ZEqnNum", StringComparison.OrdinalIgnoreCase))
                    continue;
                AssertTrue(found is null,
                    "The MathType-native fixture contains more than one GOTOBUTTON reference field.");
                found = normalized;
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
        return found
            ?? throw new InvalidDataException(
                "The MathType-native fixture has no GOTOBUTTON ZEqnNum reference field.");
    }

    private static void AssertNativeMathTypeParityStructure(
        Word.Document document,
        IReadOnlyList<string> expectedNumbers,
        IReadOnlyList<string> expectedVisibleSequences,
        string expectedSeparator,
        NativeMathTypeSectionIdentity expectedSection,
        string expectedBookmarkName,
        string expectedReferenceCode,
        string context)
    {
        AssertEqual(4, document.Paragraphs.Count,
            context + " changed the four-paragraph native document layout.");
        AssertEqual(2, CountMathTypeOleShapes(document),
            context + " changed the two native Equation.DSMT4 OLE objects.");
        AssertEqual(2, MathTypeEquationNumbering.CountPlaceRefFields(document),
            context + " changed the two native MTPlaceRef fields.");

        var actualSection = ReadSingleNativeMathTypeSectionIdentity(document);
        AssertEqual(expectedSection.FieldStart, actualSection.FieldStart,
            context + " moved the native MTEditEquationSection2 field begin.");
        AssertEqual(expectedSection.CodeStart, actualSection.CodeStart,
            context + " moved the native MTEditEquationSection2 code start.");
        AssertEqual(expectedSection.CodeEnd, actualSection.CodeEnd,
            context + " changed the native MTEditEquationSection2 field length.");
        AssertEqual(expectedSection.ParagraphStart, actualSection.ParagraphStart,
            context + " moved the native MTEditEquationSection2 to another paragraph.");
        AssertEqual(expectedSection.Code, actualSection.Code,
            context + " rewrote MathType's native MTEditEquationSection2 state.");

        Word.InlineShape? rightShape = null;
        Word.InlineShape? leftShape = null;
        try
        {
            rightShape = document.InlineShapes[1];
            leftShape = document.InlineShapes[2];
            AssertNativeMathTypeDisplayRow(
                rightShape,
                expectedNumberPosition: "right",
                context + " right-numbered equation");
            AssertNativeMathTypeDisplayRow(
                leftShape,
                expectedNumberPosition: "left",
                context + " left-numbered equation");
        }
        finally
        {
            Release(leftShape);
            Release(rightShape);
        }

        AssertMathTypeNumberTexts(document, expectedNumbers.ToArray());
        AssertNativeMathTypeReference(document, expectedNumbers[0]);
        AssertEqual(expectedReferenceCode, ReadSingleNativeMathTypeReferenceCode(document),
            context + " rebuilt or changed the native GOTOBUTTON/REF reference code.");
        AssertNativeMathTypeBookmark(
            document,
            expectedBookmarkName,
            expectedNumbers[0],
            context);
        AssertNativeMathTypePlaceRefComponents(
            document,
            expectedVisibleSequences,
            expectedSeparator,
            context);
        AssertNoOrphanNativeMathTypeSequenceFields(document, context);

        Word.Range? content = null;
        try
        {
            content = document.Content;
            var text = content.Text ?? string.Empty;
            AssertTrue(
                text.IndexOf("MACROBUTTON MTPlaceRef", StringComparison.OrdinalIgnoreCase) < 0,
                context + " exposed a raw MTPlaceRef field instruction as document text.");
            AssertTrue(
                text.IndexOf("SEQ MTEqn", StringComparison.OrdinalIgnoreCase) < 0,
                context + " exposed a raw MTEqn instruction as document text.");
        }
        finally { Release(content); }
    }

    private static void AssertNativeMathTypeDisplayRow(
        Word.InlineShape shape,
        string expectedNumberPosition,
        string context)
    {
        Word.Range? shapeRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabs = null;
        Word.TabStop? tab = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? fieldResult = null;
        Word.Range? separator = null;
        object? paragraphStyleObject = null;
        Word.Style? paragraphStyle = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            AssertEqual(1, paragraphs.Count, context + " spans more than one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            paragraphStyleObject = paragraphRange.get_Style();
            paragraphStyle = paragraphStyleObject as Word.Style;
            AssertTrue(paragraphStyle is not null,
                context + " does not expose a Word paragraph style.");
            AssertEqual("MTDisplayEquation", paragraphStyle!.NameLocal,
                context + " does not use MathType's MTDisplayEquation style.");

            format = paragraph.Format;
            tabs = format.TabStops;
            var sawCenter = false;
            var sawRight = false;
            for (var index = 1; index <= tabs.Count; index++)
            {
                Release(tab);
                tab = tabs[index];
                sawCenter |= tab.Alignment == Word.WdTabAlignment.wdAlignTabCenter;
                sawRight |= tab.Alignment == Word.WdTabAlignment.wdAlignTabRight;
            }
            AssertTrue(sawCenter && sawRight,
                context + " does not retain MathType's center/right tab stops.");

            fields = paragraphRange.Fields;
            var placeRefCount = 0;
            var placeRefStart = -1;
            var placeRefEnd = -1;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(fieldResult); fieldResult = null;
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                if (!MathTypeEquationReferences.IsMathTypePlaceRefCode(code.Text))
                    continue;
                placeRefCount++;
                fieldResult = field.Result;
                placeRefStart = code.Start - 1;
                placeRefEnd = fieldResult.End + 1;
            }
            AssertEqual(1, placeRefCount,
                context + " does not contain exactly one native MTPlaceRef.");

            var visibleParagraphText = (paragraphRange.Text ?? string.Empty)
                .Replace("\u0013", string.Empty)
                .Replace("\u0014", string.Empty)
                .Replace("\u0015", string.Empty);
            if (string.Equals(expectedNumberPosition, "left", StringComparison.Ordinal))
            {
                AssertTrue(placeRefEnd <= shapeRange.Start,
                    context + " moved the native left number after the equation.");
                separator = shapeRange.Document.Range(
                    Math.Max(paragraphRange.Start, shapeRange.Start - 1),
                    shapeRange.Start);
                AssertEqual("\t", separator.Text,
                    context + " lost MathType's number-to-equation center tab.");
                AssertTrue(
                    visibleParagraphText.StartsWith("\t\u0001", StringComparison.Ordinal),
                    context + " no longer has the native left-number + tab + OLE order.");
            }
            else
            {
                AssertEqual("right", expectedNumberPosition,
                    context + " uses an unsupported native number position.");
                AssertTrue(placeRefStart >= shapeRange.End,
                    context + " moved the native right number before the equation.");
                separator = shapeRange.Document.Range(shapeRange.End, placeRefStart);
                AssertTrue((separator.Text ?? string.Empty).IndexOf('\t') >= 0,
                    context + " lost MathType's equation-to-number right tab.");
                AssertTrue(
                    visibleParagraphText.StartsWith("\t\u0001\t", StringComparison.Ordinal),
                    context + " no longer has the native center-tab + OLE + right-tab order.");
            }
        }
        finally
        {
            Release(paragraphStyle);
            paragraphStyleObject = null;
            Release(separator);
            Release(fieldResult);
            Release(code);
            Release(field);
            Release(fields);
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static void AssertNativeMathTypeBookmark(
        Word.Document document,
        string expectedName,
        string expectedText,
        string context)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            var showHidden = bookmarks.ShowHidden;
            bookmarks.ShowHidden = true;
            try
            {
                AssertTrue(bookmarks.Exists(expectedName),
                    context + " lost the original native ZEqnNum bookmark name.");
                bookmark = bookmarks[expectedName];
                range = bookmark.Range;
                AssertEqual(expectedText, (range.Text ?? string.Empty).Trim(),
                    context + " did not move the native ZEqnNum bookmark onto the rebuilt visible number.");
                AssertTrue(range.Start < range.End,
                    context + " collapsed the native ZEqnNum bookmark.");
            }
            finally { bookmarks.ShowHidden = showHidden; }
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void AssertNativeMathTypePlaceRefComponents(
        Word.Document document,
        IReadOnlyList<string> expectedVisibleSequences,
        string expectedSeparator,
        string context)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Fields? nestedFields = null;
        Word.Field? nested = null;
        Word.Range? nestedCode = null;
        try
        {
            fields = document.Fields;
            var placeRefCount = 0;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(nestedCode); nestedCode = null;
                Release(nested); nested = null;
                Release(nestedFields); nestedFields = null;
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                if (!MathTypeEquationReferences.IsMathTypePlaceRefCode(code.Text))
                    continue;
                placeRefCount++;
                nestedFields = code.Fields;
                AssertEqual(
                    expectedVisibleSequences.Count + 1,
                    nestedFields.Count,
                    context + " rebuilt MTPlaceRef with the wrong number of nested SEQ fields.");

                var actualVisibleSequences = new List<string>();
                for (var nestedIndex = 1; nestedIndex <= nestedFields.Count; nestedIndex++)
                {
                    Release(nestedCode); nestedCode = null;
                    Release(nested); nested = nestedFields[nestedIndex];
                    nestedCode = nested.Code;
                    var normalizedNested = NormalizeNativeMathTypeFieldCode(nestedCode.Text);
                    if (nestedIndex == 1)
                    {
                        AssertTrue(
                            normalizedNested.StartsWith("SEQ MTEqn ", StringComparison.OrdinalIgnoreCase)
                            && normalizedNested.IndexOf("\\h", StringComparison.Ordinal) >= 0,
                            context + " lost MathType's leading hidden MTEqn increment.");
                        continue;
                    }
                    var match = Regex.Match(
                        normalizedNested,
                        @"^SEQ\s+(MTChap|MTSec|MTEqn)\s+\\c\b",
                        RegexOptions.IgnoreCase);
                    AssertTrue(match.Success,
                        context + " contains a non-native visible MTPlaceRef child field.");
                    actualVisibleSequences.Add(match.Groups[1].Value);
                }
                AssertEqual(
                    string.Join(",", expectedVisibleSequences),
                    string.Join(",", actualVisibleSequences),
                    context + " uses the wrong MTChap/MTSec/MTEqn component order.");

                var outerCode = NormalizeNativeMathTypeFieldCode(code.Text);
                AssertTrue(outerCode.EndsWith(")", StringComparison.Ordinal),
                    context + " does not end MTPlaceRef at its native closing parenthesis.");
                AssertTrue(!char.IsWhiteSpace(outerCode[outerCode.Length - 1]),
                    context + " leaves trailing whitespace in MTPlaceRef.");
                if (expectedVisibleSequences.Count > 1)
                {
                    AssertTrue(
                        outerCode.IndexOf(expectedSeparator, StringComparison.Ordinal) >= 0,
                        context + " lost the configured literal number separator.");
                }
            }
            AssertEqual(2, placeRefCount,
                context + " did not validate exactly two native MTPlaceRef fields.");
        }
        finally
        {
            Release(nestedCode);
            Release(nested);
            Release(nestedFields);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static void AssertNoOrphanNativeMathTypeSequenceFields(
        Word.Document document,
        string context)
    {
        var ownerRanges = new List<(int Start, int End, string Kind)>();
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                var codeText = code.Text ?? string.Empty;
                if (MathTypeEquationReferences.IsMathTypePlaceRefCode(codeText))
                    ownerRanges.Add((code.Start, code.End, "MTPlaceRef"));
                else if (MathTypeEquationNumbering.IsMathTypeSectionBreakCode(codeText))
                    ownerRanges.Add((code.Start, code.End, "MTEditEquationSection2"));
            }

            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                var normalized = NormalizeNativeMathTypeFieldCode(code.Text);
                if (!Regex.IsMatch(
                        normalized,
                        @"^SEQ\s+(?:MTEqn|MTSec|MTChap)\b",
                        RegexOptions.IgnoreCase))
                    continue;
                var fullStart = code.Start - 1;
                var owners = ownerRanges.Count(owner =>
                    fullStart >= owner.Start && code.End <= owner.End);
                AssertEqual(
                    1,
                    owners,
                    context + $" left native sequence '{normalized}' outside exactly one MTPlaceRef/MTEditEquationSection2 owner.");
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static void DumpNativeMathTypeParityStructure(
        Word.Document document,
        string path)
    {
        var builder = new StringBuilder();
        builder.AppendLine($"Paragraphs={document.Paragraphs.Count}");
        builder.AppendLine($"Fields={document.Fields.Count}");
        builder.AppendLine($"InlineShapes={document.InlineShapes.Count}");
        builder.AppendLine($"Bookmarks={document.Bookmarks.Count}");

        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Style? style = null;
        try
        {
            paragraphs = document.Paragraphs;
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Release(style); style = null;
                Release(paragraphRange); paragraphRange = null;
                Release(paragraph); paragraph = paragraphs[index];
                paragraphRange = paragraph.Range;
                try { style = paragraphRange.get_Style() as Word.Style; } catch { }
                builder.AppendLine(
                    $"P{index}|{paragraphRange.Start}:{paragraphRange.End}|style={style?.NameLocal}|text={EscapeNativeMathTypeDebugText(paragraphRange.Text)}|fields={paragraphRange.Fields.Count}|shapes={paragraphRange.InlineShapes.Count}");
            }
        }
        finally
        {
            Release(style);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }

        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? result = null;
        Word.Fields? nested = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(nested); nested = null;
                Release(result); result = null;
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                result = field.Result;
                nested = code.Fields;
                builder.AppendLine(
                    $"F{index}|type={(int)field.Type}|code={code.Start}:{code.End}|result={result.Start}:{result.End}|nested={nested.Count}|showCodes={field.ShowCodes}|instruction={EscapeNativeMathTypeDebugText(code.Text)}|value={EscapeNativeMathTypeDebugText(result.Text)}");
            }
        }
        finally
        {
            Release(nested);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
        }

        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? bookmarkRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var showHidden = bookmarks.ShowHidden;
            bookmarks.ShowHidden = true;
            try
            {
                for (var index = 1; index <= bookmarks.Count; index++)
                {
                    Release(bookmarkRange); bookmarkRange = null;
                    Release(bookmark); bookmark = bookmarks[index];
                    bookmarkRange = bookmark.Range;
                    builder.AppendLine(
                        $"B{index}|{bookmark.Name}|{bookmarkRange.Start}:{bookmarkRange.End}|text={EscapeNativeMathTypeDebugText(bookmarkRange.Text)}");
                }
            }
            finally { bookmarks.ShowHidden = showHidden; }
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
        }

        File.WriteAllText(path, builder.ToString(), new UTF8Encoding(false));
    }

    private static string NormalizeNativeMathTypeFieldCode(string? code) =>
        Regex.Replace(
            (code ?? string.Empty)
                .Replace('\u0013', ' ')
                .Replace('\u0014', ' ')
                .Replace('\u0015', ' ')
                .Trim(),
            @"\s+",
            " ");

    private static string EscapeNativeMathTypeDebugText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        return text
            .Replace("\u0001", "<OLE>")
            .Replace("\u0009", "<TAB>")
            .Replace("\u000D", "<PARA>")
            .Replace("\u0013", "<FIELD_BEGIN>")
            .Replace("\u0014", "<FIELD_SEPARATOR>")
            .Replace("\u0015", "<FIELD_END>")
            .Replace("\r", "<CR>")
            .Replace("\n", "<LF>");
    }
}
