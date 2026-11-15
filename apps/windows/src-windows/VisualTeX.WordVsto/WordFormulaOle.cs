using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using static VisualTeX.WordVsto.WordFormulaService;

namespace VisualTeX.WordVsto;

/// <summary>
/// VisualTeX's OLE formula objects: adding and initializing them, updating their data, recovering one after a
/// failed Word command, MathType's data inside them and their display preview metrics.
/// </summary>
internal static class WordFormulaOle
{
    // Shared by redraw preflight and the actual insertion. Validate generated
    // bytes before the first source deletion; preview success alone does not
    // establish a lossless MathML -> Equation Native -> MathML round trip.
    internal static (MathTypeMtefCodec.RewriteResult Generated, byte[] CompoundFile,
        string SemanticSignature) PrepareStandaloneMathTypeOleData(
        string mathMl, bool inline, double fontSizePt, string latex)
    {
        var generated = MathTypeMtefFontSize.CreateEquationNativeAtFontSize(mathMl, inline, fontSizePt);
        var compoundFile = MathTypeOleStorage.CreateStandaloneCompoundFile(generated);
        var actualMathMl = MathTypeOleStorage.ReadMathMl(compoundFile);
        var expected = MathTypeMtefCodec.SemanticSignature(mathMl);
        if (!MathTypeMathMlRoundTripMatches(expected, actualMathMl))
        {
            // The user sees what happened; the exact difference goes to the trace.
            WordDoubleClickHook.TraceMessage(
                $"mathtype-roundtrip-mismatch latex={latex} "
                + DescribeSemanticSignatureDifference(expected, MathTypeMtefCodec.SemanticSignature(actualMathMl)));
            throw new InvalidDataException(OfficePluginLanguage.Text(
                "转换会丢失或改变公式中的内容（例如 MathType 不能拉伸的括号）。",
                "Converting would lose or change part of the equation (for example a bracket MathType cannot stretch)."));
        }
        return (generated, compoundFile, expected);
    }

    internal static bool MathTypeMathMlRoundTripMatches(
        string expectedSignature,
        string actualMathMl) =>
        string.Equals(
            expectedSignature,
            MathTypeMtefCodec.SemanticSignature(actualMathMl),
            StringComparison.Ordinal);

    private static string DescribeSemanticSignatureDifference(
        string expected,
        string actual)
    {
        expected ??= string.Empty;
        actual ??= string.Empty;
        var common = Math.Min(expected.Length, actual.Length);
        var index = 0;
        while (index < common && expected[index] == actual[index]) index++;
        if (index == common && expected.Length == actual.Length)
            return $"identical(length={expected.Length})";

        static string DescribeAt(string value, int position)
        {
            if (position < 0 || position >= value.Length) return "<end>";
            var ch = value[position];
            return $"U+{(int)ch:X4}('{(char.IsControl(ch) ? '?' : ch)}')";
        }

        var start = Math.Max(0, index - 6);
        var expectedEnd = Math.Min(expected.Length, index + 7);
        var actualEnd = Math.Min(actual.Length, index + 7);
        static string Escape(string value) => string.Concat(value.Select(ch =>
            ch >= ' ' && ch <= '~' ? ch.ToString() : $"\\u{(int)ch:X4}"));
        return $"index={index}; expectedLen={expected.Length}; actualLen={actual.Length}; "
            + $"expectedChar={DescribeAt(expected, index)}; actualChar={DescribeAt(actual, index)}; "
            + $"expectedContext='{Escape(expected.Substring(start, expectedEnd - start))}'; "
            + $"actualContext='{Escape(actual.Substring(start, actualEnd - start))}'";
    }

    internal static InlineShape? ResolveCapturedFormulaPicture(
        Document document,
        string? sourceObjectId,
        string formulaId)
    {
        Range? captured = null;
        InlineShapes? shapes = null;
        InlineShape? match = null;
        try
        {
            captured =
                WordFormulaOperationLocator.ResolveCapturedRange(
                    document,
                    sourceObjectId);
            shapes = captured.InlineShapes;
            for (var index = 1;
                 index <= shapes.Count;
                 index++)
            {
                InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (WordFormulaMetadataReader.IsNativeOle(
                            candidate)
                        || MathTypeOleInterop.IsMathTypeOle(
                            candidate))
                        continue;

                    var metadata =
                        WordFormulaMetadataReader
                            .TryReadCachedPreview(
                                candidate);
                    if (!string.Equals(
                            metadata?.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (match is not null)
                        throw new InvalidDataException(
                            "The captured Word range contains more than one formula picture with the same identity.");

                    match = candidate;
                    candidate = null;
                }
                finally
                {
                    Release(candidate);
                }
            }

            var result = match;
            match = null;
            return result;
        }
        finally
        {
            Release(match);
            Release(shapes);
            Release(captured);
        }
    }

    private static InlineShape AddOleObject(Document document, Range range)
    {
        var insertionStart = range.Start;
        Range? originalRightCharacter = null;
        try
        {
            // Keep one live boundary, not an inventory of document.InlineShapes.
            // On a collapsed insertion the original right-hand character must
            // remain AFTER anything newly inserted. Its live End distinguishes a
            // newly materialized OLE from an old object at the same numeric offset.
            // Failure to capture the witness disables recovery, not insertion.
            if (range.End == insertionStart)
            {
                try
                {
                    originalRightCharacter = range.Duplicate;
                    originalRightCharacter.SetRange(insertionStart, checked(insertionStart + 1));
                }
                catch (COMException)
                {
                    Release(originalRightCharacter);
                    originalRightCharacter = null;
                }
            }
            // Successful insertion is the overwhelmingly common path. Do not
            // enumerate document.InlineShapes merely to prepare evidence for the
            // rare wdErrorCommandFailed recovery case; that turns a 100-formula
            // batch into repeated whole-document COM work. The recovery path below
            // proves ownership from the exact EMBED field boundary instead.
            return document.InlineShapes.AddOLEObject(
                ClassType: FormulaOleContract.ProgId,
                LinkToFile: false,
                DisplayAsIcon: false,
                Range: range);
        }
        catch (COMException error) when (
            error.HResult == unchecked((int)0x800A1066))
        {
            // Word can finish embedding the OLE storage and then report
            // wdErrorCommandFailed after the server's creation verb. Recover only
            // a VisualTeX EMBED field that starts at this exact insertion point;
            // otherwise preserve the original failure.
            var recovered = TryRecoverMaterializedOleAfterCommandFailure(
                document,
                insertionStart,
                originalRightCharacter,
                out var recoveryEvidence);
            if (recovered is null)
                throw new COMException(
                    $"{error.Message} VisualTeX OLE recovery evidence: {recoveryEvidence}",
                    error.HResult);
            WordDoubleClickHook.TraceMessage(
                $"ole-add-command-failed-but-materialized start={insertionStart}");
            return recovered;
        }
        finally { Release(originalRightCharacter); }
    }

    private static InlineShape? TryRecoverMaterializedOleAfterCommandFailure(
        Document document,
        int insertionStart,
        Range? originalRightCharacter,
        out string evidence)
    {
        Range? probe = null;
        InlineShapes? shapes = null;
        InlineShape? candidate = null;
        Range? candidateRange = null;
        OLEFormat? format = null;
        try
        {
            if (originalRightCharacter is null
                || originalRightCharacter.StoryType != WdStoryType.wdMainTextStory)
            {
                evidence = $"start={insertionStart}; new-insertion-boundary=unproven";
                return null;
            }
            var newInsertionEnd = originalRightCharacter.End - 1;
            if (newInsertionEnd <= insertionStart)
            {
                evidence = $"start={insertionStart}; original-right-character={newInsertionEnd}; no-new-insertion";
                return null;
            }
            var content = document.Content;
            try
            {
                var probeStart = Math.Max(content.Start, insertionStart);
                var probeEnd = Math.Min(Math.Min(content.End, newInsertionEnd), insertionStart + 128);
                if (probeEnd <= probeStart)
                    probeEnd = Math.Min(content.End, probeStart + 1);
                probe = document.Range(probeStart, probeEnd);
            }
            finally { Release(content); }
            shapes = probe.InlineShapes;
            evidence = $"start={insertionStart}; localShapes={shapes.Count}";
            for (var index = shapes.Count; index >= 1; index--)
            {
                candidate = shapes[index];
                candidateRange = candidate.Range;
                // A successful AddOLEObject normally returns the result Range at
                // the insertion coordinate (or after the display TAB). When the
                // OLE server finishes storage creation and Word subsequently
                // raises wdErrorCommandFailed, however, Word can expose only the
                // one-character result of its EMBED field. The result starts after
                // the hidden field code (about 30 characters), not at the field's
                // opening coordinate. Accept that shape only when the surrounding
                // wdFieldEmbed is independently proven to start at this exact
                // insertion point and names the VisualTeX ProgID.
                if (candidateRange.Start < insertionStart
                    || candidateRange.End > newInsertionEnd
                    || (candidateRange.Start != insertionStart
                        && candidateRange.Start != insertionStart + 1
                        && !IsVisualTeXOleEmbedFieldAtInsertion(
                            document,
                            insertionStart,
                            candidateRange)))
                {
                    Release(candidateRange); candidateRange = null;
                    Release(candidate); candidate = null;
                    continue;
                }
                format = candidate.OLEFormat;
                if (!string.Equals(
                        WordOleObjectAccessor.ReadProgId(format),
                        FormulaOleContract.ProgId,
                        StringComparison.OrdinalIgnoreCase))
                {
                    Release(format); format = null;
                    Release(candidateRange); candidateRange = null;
                    Release(candidate); candidate = null;
                    continue;
                }

                var recovered = candidate;
                candidate = null;
                return recovered;
            }
            var coordinates = new List<string>(shapes.Count);
            for (var index = 1; index <= shapes.Count; index++)
            {
                candidate = shapes[index];
                candidateRange = candidate.Range;
                coordinates.Add($"{index}:{candidateRange.Start}-{candidateRange.End}");
                Release(candidateRange); candidateRange = null;
                Release(candidate); candidate = null;
            }
            evidence += "; ranges=" + string.Join(",", coordinates);
            return null;
        }
        catch (COMException)
        {
            evidence = $"start={insertionStart}; localInventory=COM-failed";
            return null;
        }
        finally
        {
            Release(format);
            Release(candidateRange);
            Release(candidate);
            Release(shapes);
            Release(probe);
        }
    }

    private static bool IsVisualTeXOleEmbedFieldAtInsertion(
        Document document,
        int insertionStart,
        Range candidateRange)
    {
        // Bound the recovery to Word's short EMBED field-code displacement. A
        // remotely placed object must never be adopted merely because its ProgID
        // matches.
        if (candidateRange.Start < insertionStart
            || candidateRange.Start - insertionStart > 128)
            return false;

        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        try
        {
            fields = document.Fields;
            for (var index = fields.Count; index >= 1; index--)
            {
                field = fields[index];
                if (field.Type != WdFieldType.wdFieldEmbed)
                {
                    Release(field); field = null;
                    continue;
                }
                code = field.Code;
                result = field.Result;
                var codeStartsAtInsertion = code.Start == insertionStart
                    || code.Start == insertionStart + 1;
                var ownsCandidate = result.Start <= candidateRange.Start
                    && result.End >= candidateRange.End;
                var namesVisualTeX = (code.Text ?? string.Empty).IndexOf(
                        FormulaOleContract.ProgId,
                        StringComparison.OrdinalIgnoreCase) >= 0;
                if (codeStartsAtInsertion && ownsCandidate && namesVisualTeX)
                    return true;

                Release(result); result = null;
                Release(code); code = null;
                Release(field); field = null;
            }
            return false;
        }
        catch (COMException)
        {
            return false;
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static void InitializeOle(
        InlineShape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        IVisualTeXFormulaObject? formula = null;
        var initialized = false;
        try
        {
            StoreWordDisplayPreviewMetrics(metadata, emfPath);
            format = shape.OLEFormat;
            oleObject = WordOleObjectAccessor.GetRunningObject(format);
            formula = oleObject as IVisualTeXFormulaObject;
            if (formula is null)
                throw new InvalidOperationException(
                    "The inserted Word object does not expose the VisualTeX native OLE interface.");
            FormulaOleInterop.Initialize(formula, metadata, emfPath, pngPath);
            WordFormulaMetadataReader.CacheMetadata(shape, metadata);
            initialized = true;
        }
        finally
        {
            if (formula is not null)
            {
                try { FormulaOleInterop.CloseAfterSave(formula); }
                catch when (!initialized) { }
            }
            Release(oleObject);
            Release(format);
        }
    }

    private static bool TryUpdateOle(
        InlineShape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        IVisualTeXFormulaObject? formula = null;
        var updated = false;
        try
        {
            StoreWordDisplayPreviewMetrics(metadata, emfPath);
            try { format = shape.OLEFormat; }
            catch { return false; }
            try { oleObject = WordOleObjectAccessor.GetRunningObject(format); }
            catch { return false; }
            formula = oleObject as IVisualTeXFormulaObject;
            if (formula is null) return false;
            FormulaOleInterop.Update(formula, metadata, emfPath, pngPath);
            WordFormulaMetadataReader.CacheMetadata(shape, metadata);
            updated = true;
            return true;
        }
        finally
        {
            if (formula is not null)
            {
                try { FormulaOleInterop.CloseAfterSave(formula); }
                catch when (!updated) { }
            }
            Release(oleObject);
            Release(format);
        }
    }

    private static void StoreWordDisplayPreviewMetrics(
        FormulaMetadata metadata,
        string emfPath)
    {
        if (!string.Equals(
                metadata.DisplayMode,
                "block",
                StringComparison.OrdinalIgnoreCase))
        {
            metadata.WordDisplayPreviewInkHeightRatio = null;
            metadata.WordDisplayPreviewBottomWhitespaceRatio = null;
            return;
        }

        var metrics = OfficeOlePreview.TryMeasureVerticalInkMetrics(emfPath);
        if (!metrics.HasValue)
        {
            metadata.WordDisplayPreviewInkHeightRatio = null;
            metadata.WordDisplayPreviewBottomWhitespaceRatio = null;
            return;
        }
        metadata.WordDisplayPreviewInkHeightRatio =
            metrics.Value.InkHeightRatio;
        metadata.WordDisplayPreviewBottomWhitespaceRatio =
            metrics.Value.BottomWhitespaceRatio;
        metadata.Validate();
    }

    internal static byte[] ReadMathTypeMtefFromCompoundFile(byte[] compoundFile)
    {
        var equationNative = MathTypeOleStorage.ReadEquationNative(compoundFile);
        if (equationNative.Length < 12)
            throw new InvalidDataException("MathType Equation Native is too short for MTEF extraction.");
        var headerLength = BitConverter.ToUInt16(equationNative, 0);
        var mtefLength = checked((int)BitConverter.ToUInt32(equationNative, 8));
        if (headerLength < 12
            || mtefLength <= 0
            || headerLength + mtefLength > equationNative.Length)
            throw new InvalidDataException("MathType Equation Native contains an invalid MTEF extent.");
        var mtef = new byte[mtefLength];
        Buffer.BlockCopy(equationNative, headerLength, mtef, 0, mtefLength);
        return mtef;
    }

    private static bool TryRenderMathTypeNativePreviewFromCompoundFile(
        byte[] compoundFile,
        string outputDirectory,
        out MathTypeNativePreviewRenderer.Result? result)
    {
        result = null;
        try
        {
            var mtef = ReadMathTypeMtefFromCompoundFile(compoundFile);
            if (!MathTypeNativePreviewRenderer.TryRender(
                    mtef,
                    outputDirectory,
                    out var rendered))
                return false;
            result = rendered;
            return true;
        }
        catch { return false; }
    }
}
