using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordEquationNumbering;
using static VisualTeX.WordVsto.WordEquationSequenceFields;
using static VisualTeX.WordVsto.WordNumberedTrueDisplay;
using static VisualTeX.WordVsto.WordNumberingFields;

namespace VisualTeX.WordVsto;

/// <summary>
/// Reading the numbering state of a document: field codes, number artifacts and caption entries from its Word
/// OpenXML or its bookmarks, and reference counts.
/// </summary>
internal static class WordNumberingState
{
    private sealed class OpenXmlFieldCodeState
    {
        internal System.Text.StringBuilder Code { get; } = new();
        internal bool ReadingInstruction { get; set; } = true;
    }

    private static IReadOnlyList<string> ExtractFieldCodesFromWordOpenXml(
        string wordOpenXml)
    {
        if (string.IsNullOrWhiteSpace(wordOpenXml))
            return Array.Empty<string>();
        const string wordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string mathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        try
        {
            var document = System.Xml.Linq.XDocument.Parse(
                wordOpenXml,
                System.Xml.Linq.LoadOptions.PreserveWhitespace);
            var word = (System.Xml.Linq.XNamespace)wordNamespace;
            var math = (System.Xml.Linq.XNamespace)mathNamespace;
            var active = new Stack<OpenXmlFieldCodeState>();
            var result = new List<string>();
            foreach (var element in document.Descendants())
            {
                if (element.Name == word + "fldSimple")
                {
                    var instruction = (string?)element.Attribute(word + "instr");
                    if (!string.IsNullOrWhiteSpace(instruction))
                        result.Add(instruction!);
                    continue;
                }

                if (element.Name == word + "fldChar")
                {
                    var kind = (string?)element.Attribute(word + "fldCharType")
                        ?? string.Empty;
                    if (string.Equals(kind, "begin", StringComparison.OrdinalIgnoreCase))
                    {
                        active.Push(new OpenXmlFieldCodeState());
                    }
                    else if (string.Equals(
                                 kind,
                                 "separate",
                                 StringComparison.OrdinalIgnoreCase))
                    {
                        if (active.Count > 0)
                            active.Peek().ReadingInstruction = false;
                    }
                    else if (string.Equals(kind, "end", StringComparison.OrdinalIgnoreCase))
                    {
                        if (active.Count > 0)
                        {
                            var completed = active.Pop().Code.ToString();
                            if (!string.IsNullOrWhiteSpace(completed))
                                result.Add(completed);
                        }
                    }
                    continue;
                }

                if (active.Count == 0 || !active.Peek().ReadingInstruction)
                    continue;
                if (element.Name == word + "instrText"
                    || element.Name == word + "t"
                    || element.Name == math + "t")
                    active.Peek().Code.Append(element.Value);
            }
            return result;
        }
        catch
        {
            return Array.Empty<string>();
        }
    }

    internal static bool IsOleNumberingParagraphXml(string paragraphXml, string formulaId)
    {
        // Classification only: the normal OLE health path below still checks tab
        // geometry, unique aliases, sequence state and reference ownership.
        // Never classify a user table by column count or by a neighbouring cell.
        if (!Guid.TryParse(formulaId, out var id)) return false;
        return Regex.Matches(paragraphXml, Regex.Escape("ProgID=\"VisualTeX.Formula.1\""),
                   RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count == 1
            && !Regex.IsMatch(paragraphXml, @"<m:oMath(?:\s|>)",
                   RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
            && Regex.IsMatch(paragraphXml,
                @"\bREF\s+(?:&quot;|"")?VTEqNum_" + id.ToString("N") + @"(?=\s|&quot;|""|\\|<|$)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }

    internal static bool IsUnmanagedEquationDocumentXml(string xml)
    {
        if (string.IsNullOrWhiteSpace(xml)) return false;
        // Absence of number aliases alone is not enough: an unnumbered/copied
        // VisualTeX object may still need adoption. Be conservative for all known
        // current and legacy metadata/reference spellings.
        if (xml.IndexOf("VTEq", StringComparison.OrdinalIgnoreCase) >= 0
            || xml.IndexOf("VTO_", StringComparison.OrdinalIgnoreCase) >= 0
            || xml.IndexOf("VTOMML_", StringComparison.OrdinalIgnoreCase) >= 0
            || xml.IndexOf("visualtex:", StringComparison.OrdinalIgnoreCase) >= 0
            || xml.IndexOf("VisualTeXEquation", StringComparison.OrdinalIgnoreCase) >= 0)
            return false;
        try
        {
            var parsed = XDocument.Parse(xml);
            var elements = parsed.Descendants().ToArray();
            return elements.Any(e => e.Name.LocalName == "body")
                && !elements.Any(e => e.Name.LocalName is "oMath" or "oMathPara")
                && !elements.Attributes().Any(a => a.Name.LocalName == "ProgID"
                    && a.Value.StartsWith("VisualTeX", StringComparison.OrdinalIgnoreCase));
        }
        catch { return false; }
    }

    internal static bool TryReadHealthyEquationNumberArtifactsFromOpenXml(
        Document document,
        out IReadOnlyList<NativeEquationCaptionEntry> entries,
        out IReadOnlyDictionary<string, int> referenceCounts)
    {
        entries = Array.Empty<NativeEquationCaptionEntry>();
        referenceCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var xml = string.Empty;
        bool Fail(string reason)
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                    "1",
                    StringComparison.Ordinal)
                || string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                TraceNumberingPerformance($"[perf] update-numbers.health-fail reason={reason}");
            return false;
        }
        int LastElementStart(string qualifiedName, int beforeIndex)
        {
            var bare = xml.LastIndexOf(
                $"<{qualifiedName}>",
                beforeIndex,
                StringComparison.OrdinalIgnoreCase);
            var attributed = xml.LastIndexOf(
                $"<{qualifiedName} ",
                beforeIndex,
                StringComparison.OrdinalIgnoreCase);
            return Math.Max(bare, attributed);
        }
        int FirstElementStart(string qualifiedName, int afterIndex)
        {
            var bare = xml.IndexOf(
                $"<{qualifiedName}>",
                afterIndex,
                StringComparison.OrdinalIgnoreCase);
            var attributed = xml.IndexOf(
                $"<{qualifiedName} ",
                afterIndex,
                StringComparison.OrdinalIgnoreCase);
            if (bare < 0) return attributed;
            if (attributed < 0) return bare;
            return Math.Min(bare, attributed);
        }
        Range? content = null;
        try
        {
            content = document.Content;
            xml = WordDocumentXml.Read(document) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(xml)) return Fail("empty-openxml");
            var parsedFieldCodes = ExtractFieldCodesFromWordOpenXml(xml);
            if (parsedFieldCodes.Count == 0)
                return Fail("no-parseable-field-codes");

            var bookmarkMatches = Regex.Matches(
                xml,
                @"<w:bookmarkStart\b[^>]*\bw:name=""(?<name>VTEq(?:Cap|Num)?_[^""]+)""[^>]*/?>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (bookmarkMatches.Count == 0)
            {
                if (IsUnmanagedEquationDocumentXml(xml)) return true;
                return Fail("no-numbering-bookmarks");
            }
            if (Regex.IsMatch(
                    xml,
                    @"<w:bookmarkStart\b(?=[^>]*\bw:name=""(?:VTEqAnc_|VTAncR_)[0-9A-F]{32}""?)[^>]*/?>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                || Regex.IsMatch(
                    xml,
                    @"\b(?:name|id|descr|alt)=""(?:VTEqShape_[0-9A-F]{32}|VisualTeX numbered OMML [0-9A-F-]{36})""",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                return Fail("legacy-native-display-anchor-or-shape-remains");

            var visibleIds = new List<string>();
            var captionIds = new List<string>();
            var numberIds = new List<string>();
            foreach (Match match in bookmarkMatches)
            {
                var name = match.Groups["name"].Value;
                if (TryFormulaIdFromBookmark(
                        name,
                        NativeCaptionBookmarkPrefix,
                        out var formulaId))
                {
                    captionIds.Add(formulaId);
                    continue;
                }
                if (TryFormulaIdFromBookmark(
                        name,
                        NativeNumberBookmarkPrefix,
                        out formulaId))
                {
                    numberIds.Add(formulaId);
                    continue;
                }
                if (TryFormulaIdFromBookmark(
                        name,
                        EquationBookmarkPrefix,
                        out formulaId))
                {
                    visibleIds.Add(formulaId);
                    continue;
                }
                return false;
            }

            var visibleSet = visibleIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var captionSet = captionIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var numberSet = numberIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
            // A floating Word text box is commonly serialized twice inside one
            // mc:AlternateContent block: once in the DrawingML Choice and once in
            // the VML Fallback. Both representations contain the same VTEq_ bookmark
            // and generated REF, but they are one logical Shape in Word COM. Allow
            // only that visible-bookmark duplication; caption/number targets remain
            // unique main-story structures.
            if (visibleSet.Count == 0
                || captionSet.Count != captionIds.Count
                || numberSet.Count != numberIds.Count
                || !visibleSet.SetEquals(captionSet)
                || !visibleSet.SetEquals(numberSet))
                return Fail($"bookmark-id-set-mismatch visible={visibleIds.Count}/{visibleSet.Count}[{string.Join(",", visibleIds)}] caption={captionIds.Count}/{captionSet.Count}[{string.Join(",", captionIds)}] number={numberIds.Count}/{numberSet.Count}[{string.Join(",", numberIds)}]");

            var visualTeXSequenceCodes = parsedFieldCodes
                .Where(IsVisualTeXSequenceFieldCode)
                .ToArray();
            if (visualTeXSequenceCodes.Length != visibleSet.Count)
                return Fail(
                    $"visualtex-sequence-count-mismatch fields={visualTeXSequenceCodes.Length} formulas={visibleSet.Count}");
            // Legacy/current OLE captions may intentionally use an ordinal restart
            // while rebuilding a mixed-format document. That is ordinary Word text
            // and remains editable. Only a mathematical native #(SEQ) field is
            // forbidden from carrying \\r N, because its instruction must stay
            // immutable and dynamically ordered by F9. Enforce that per native
            // OMath below rather than rejecting the entire mixed document here.

            // A complete scaffold whose visible number no longer shares an owner
            // with its formula is an orphan (for example after a user manually
            // deletes only the OLE/OMML object). Current OLE and native OMML both
            // use one tab paragraph; legacy managed OMML tables remain accepted
            // long enough to be migrated by the structural reconciliation path.
            // Older OLE documents do not necessarily carry VTO_* identity
            // bookmarks, so validate the actual owner payload rather than requiring
            // an OLE identity bookmark.
            var allVisibleStartMatches = Regex.Matches(
                    xml,
                    @"<w:bookmarkStart\b(?=[^>]*\bw:name=""VTEq_(?<guid>[0-9A-F]{32})"")[^>]*/>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                .Cast<Match>()
                .ToArray();
            var visibleRepresentationCounts = allVisibleStartMatches
                .GroupBy(
                    match => match.Groups["guid"].Value,
                    StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    group => Guid.ParseExact(group.Key, "N").ToString("D"),
                    group => group.Count(),
                    StringComparer.OrdinalIgnoreCase);
            var visibleStartMatches = allVisibleStartMatches
                .GroupBy(
                    match => match.Groups["guid"].Value,
                    StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToArray();
            if (visibleStartMatches.Length != visibleSet.Count)
                return Fail($"visible-bookmark-start-count-mismatch representations={allVisibleStartMatches.Length} unique={visibleStartMatches.Length} expected={visibleSet.Count}");
            var hashSequenceFormulaIds = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            var directTableFormulaIds = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (var visibleStart in visibleStartMatches)
            {
                var normalizedFormulaId = visibleStart.Groups["guid"].Value;
                var formulaId = Guid.ParseExact(
                    normalizedFormulaId,
                    "N").ToString("D");

                // Current numbered OMML stores all three VisualTeX aliases inside
                // Word's own #(SEQ VisualTeXEquation) delimiter. A native m:eqArr
                // is therefore healthy when the VTEq_ alias is inside the one
                // display OMath, the same wrapper owns VTEqNum_<FormulaId>, and the
                // math field is SEQ rather than REF. Do this check before the legacy
                // Shape/TextBox branch so legal Word #() markup is never classified
                // as a malformed artificial eqArr.
                var mathStart = LastElementStart("m:oMath", visibleStart.Index);
                var precedingMathEnd = xml.LastIndexOf(
                    "</m:oMath>",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                var visibleNumberIsInMath = mathStart >= 0
                    && mathStart > precedingMathEnd;
                if (visibleNumberIsInMath)
                {
                    var mathEnd = xml.IndexOf(
                        "</m:oMath>",
                        visibleStart.Index,
                        StringComparison.OrdinalIgnoreCase);
                    if (mathEnd <= visibleStart.Index)
                        return Fail($"hash-seq-math-end-missing formulaId={formulaId}");
                    mathEnd += "</m:oMath>".Length;
                    if (mathEnd - mathStart > 262144)
                        return Fail($"hash-seq-math-too-large formulaId={formulaId}");
                    var mathXml = xml.Substring(mathStart, mathEnd - mathStart);
                    var expectedVisibleBookmark = EquationBookmarkPrefix + normalizedFormulaId;
                    var expectedNumberBookmark = NativeNumberBookmarkPrefix + normalizedFormulaId;
                    var expectedCaptionBookmark = NativeCaptionBookmarkPrefix + normalizedFormulaId;
                    bool BookmarkIsFullyInsideMath(string bookmarkName)
                    {
                        var escapedName = Regex.Escape(bookmarkName);
                        var start = Regex.Match(
                            mathXml,
                            $@"<w:bookmarkStart\b(?=[^>]*\bw:id=""(?<id>-?\d+)"")(?=[^>]*\bw:name=""{escapedName}"")[^>]*/?>",
                            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                        if (!start.Success) return false;
                        var escapedId = Regex.Escape(start.Groups["id"].Value);
                        var end = Regex.Match(
                            mathXml,
                            $@"<w:bookmarkEnd\b[^>]*\bw:id=""{escapedId}""[^>]*/?>",
                            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                        return end.Success && end.Index > start.Index;
                    }
                    var directSequence = Regex.IsMatch(
                        mathXml,
                        @"\bSEQ\s+(?:&quot;|"")?VisualTeXEquation(?:&quot;|"")?\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasVisibleAlias =
                        BookmarkIsFullyInsideMath(expectedVisibleBookmark);
                    var hasNumberAlias =
                        BookmarkIsFullyInsideMath(expectedNumberBookmark);
                    var hasCaptionAlias =
                        BookmarkIsFullyInsideMath(expectedCaptionBookmark);
                    var hasHash = Regex.IsMatch(
                        mathXml,
                        @"<m:t(?:\s[^>]*)?>#</m:t>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasEquationArray = mathXml.IndexOf(
                        "<m:eqArr",
                        StringComparison.OrdinalIgnoreCase) >= 0;
                    var hasMathReference = Regex.IsMatch(
                        mathXml,
                        @"\bREF\s+VTEqNum_",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasOrdinalRestart = Regex.IsMatch(
                        mathXml,
                        @"\\r\s+\d+\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    if (!directSequence || !hasVisibleAlias || !hasNumberAlias
                        || !hasCaptionAlias || !hasHash || !hasEquationArray
                        || hasMathReference || hasOrdinalRestart)
                        return Fail(
                            $"hash-seq-invalid formulaId={formulaId} seq={directSequence} visible={hasVisibleAlias} num={hasNumberAlias} cap={hasCaptionAlias} hash={hasHash} eqArr={hasEquationArray} ref={hasMathReference} restart={hasOrdinalRestart}");

                    var hashParagraphStart = LastElementStart("w:p", mathStart);
                    var hashPrecedingParagraphEnd = xml.LastIndexOf(
                        "</w:p>",
                        mathStart,
                        StringComparison.OrdinalIgnoreCase);
                    var hashParagraphEnd = xml.IndexOf(
                        "</w:p>",
                        mathEnd,
                        StringComparison.OrdinalIgnoreCase);
                    if (hashParagraphStart < 0
                        || hashParagraphStart <= hashPrecedingParagraphEnd
                        || hashParagraphEnd < mathEnd)
                        return Fail($"hash-seq-owner-paragraph-missing formulaId={formulaId}");
                    hashParagraphEnd += "</w:p>".Length;
                    var hashParagraphXml = xml.Substring(
                        hashParagraphStart,
                        hashParagraphEnd - hashParagraphStart);
                    if (hashParagraphXml.IndexOf(
                            "<m:oMathPara",
                            StringComparison.OrdinalIgnoreCase) < 0)
                        return Fail($"hash-seq-not-display formulaId={formulaId}");
                    var formulaBookmarkName =
                        WordOmmlFormulaStore.BookmarkName(formulaId);
                    var formulaBookmarkMatches = Regex.Matches(
                        xml,
                        $@"<w:bookmarkStart\b(?=[^>]*\bw:name=""{Regex.Escape(formulaBookmarkName)}"")[^>]*/?>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    if (formulaBookmarkMatches.Count != 1)
                        return Fail(
                            $"hash-seq-formula-identity-count formulaId={formulaId} count={formulaBookmarkMatches.Count}");
                    var formulaBookmarkIndex = formulaBookmarkMatches[0].Index;
                    var firstMathAfterFormulaBookmark = FirstElementStart(
                        "m:oMath",
                        formulaBookmarkIndex);
                    if (firstMathAfterFormulaBookmark != mathStart)
                        return Fail(
                            $"hash-seq-formula-identity-mismatch formulaId={formulaId} formulaBookmark={formulaBookmarkIndex} expectedMath={mathStart} firstMathAfterIdentity={firstMathAfterFormulaBookmark}");
                    var hashParagraphTableStart = LastElementStart(
                        "w:tbl",
                        hashParagraphStart);
                    var hashPrecedingTableEnd = xml.LastIndexOf(
                        "</w:tbl>",
                        hashParagraphStart,
                        StringComparison.OrdinalIgnoreCase);
                    // The owning display paragraph may be inside a user cell.
                    // The checks above already prove the equation, native number,
                    // and all three aliases belong to one OMath in that paragraph.
                    // A table ancestor is not a different numbering format.
                    if (hashParagraphXml.IndexOf(
                            NativeDisplayNumberShapeName(formulaId),
                            StringComparison.OrdinalIgnoreCase) >= 0
                        || hashParagraphXml.IndexOf(
                            "<w:txbxContent",
                            StringComparison.OrdinalIgnoreCase) >= 0)
                        return Fail($"hash-seq-has-shape formulaId={formulaId}");
                    hashSequenceFormulaIds.Add(formulaId);
                    continue;
                }

                // Shape-era numbered OMML stored the visible VTEq_ bookmark and an
                // ordinary REF inside a floating text-box story. Keep accepting it
                // only as migration input; the current producer never creates it.
                var textBoxStart = xml.LastIndexOf(
                    "<w:txbxContent",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                var precedingTextBoxEnd = xml.LastIndexOf(
                    "</w:txbxContent>",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                var visibleNumberIsInTextBox = textBoxStart >= 0
                    && textBoxStart > precedingTextBoxEnd;
                if (visibleNumberIsInTextBox)
                    return Fail($"legacy-numbering-shape formulaId={formulaId}");

                var tableStart = LastElementStart("w:tbl", visibleStart.Index);
                var precedingTableEnd = xml.LastIndexOf(
                    "</w:tbl>",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                var visibleNumberIsInTable = tableStart >= 0
                    && tableStart > precedingTableEnd;
                var localParagraphStart = LastElementStart("w:p", visibleStart.Index);
                var localParagraphEnd = xml.IndexOf("</w:p>", visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                var cellOleNumber = visibleNumberIsInTable && localParagraphStart >= 0
                    && localParagraphEnd > visibleStart.Index
                    && IsOleNumberingParagraphXml(xml.Substring(localParagraphStart,
                        localParagraphEnd + "</w:p>".Length - localParagraphStart), formulaId);
                if (visibleNumberIsInTable && !cellOleNumber)
                {
                    // Current numbered OMML uses one minimal 1x3 table only as a
                    // layout container. Cell (1,2) owns the one genuine display
                    // OMath and no fields; cell (1,3) owns TAB + direct ordinary
                    // SEQ + paragraph mark. Accept that exact structure as healthy
                    // rather than sending every Update Numbers command through the
                    // legacy-table migration/reconciliation path.
                    var tableEnd = xml.IndexOf(
                        "</w:tbl>",
                        visibleStart.Index,
                        StringComparison.OrdinalIgnoreCase);
                    if (tableEnd <= visibleStart.Index)
                        return Fail($"direct-table-end-missing formulaId={formulaId}");
                    tableEnd += "</w:tbl>".Length;
                    if (tableEnd - tableStart > 524288)
                        return Fail($"direct-table-too-large formulaId={formulaId}");
                    var tableXml = xml.Substring(tableStart, tableEnd - tableStart);
                    var rowCount = Regex.Matches(
                        tableXml,
                        @"<w:tr(?:\s|>)",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                    var tableCellCount = Regex.Matches(
                        tableXml,
                        @"<w:tc(?:\s|>)",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                    if (rowCount < 1 || tableCellCount != rowCount * 3)
                        return Fail(
                            $"direct-table-dimensions formulaId={formulaId} rows={rowCount} cells={tableCellCount}");
                    if (rowCount > 1)
                    {
                        // Current dense numbered-OMML tables may legitimately
                        // share one N x 3 host, but every row then owns exactly one
                        // center OMath. Historical Word 2019 insertions could leave
                        // an additional persisted 3-cell row containing no formula.
                        // Do not call that host healthy: failing the XML-only fast
                        // probe routes Update Numbers through the targeted spacing
                        // repair, which removes only the structurally empty row.
                        var tableRows = Regex.Matches(
                                tableXml,
                                @"<w:tr(?:\s|>).*?</w:tr>",
                                RegexOptions.IgnoreCase
                                | RegexOptions.CultureInvariant
                                | RegexOptions.Singleline)
                            .Cast<Match>()
                            .ToArray();
                        if (tableRows.Length != rowCount
                            || tableRows.Any(row => !Regex.IsMatch(
                                row.Value,
                                @"<m:oMath(?:\s|>)",
                                RegexOptions.IgnoreCase
                                | RegexOptions.CultureInvariant)))
                            return Fail(
                                $"direct-table-empty-or-unowned-row formulaId={formulaId} rows={rowCount}");
                    }

                    // Consecutive numbered OMML formulas share one managed N x 3
                    // table. Validate the row that owns this formula's visible
                    // bookmark, rather than applying the retired 1 x 3 assumption
                    // to the entire table.
                    var rowStart = LastElementStart("w:tr", visibleStart.Index);
                    var precedingRowEnd = xml.LastIndexOf(
                        "</w:tr>",
                        visibleStart.Index,
                        StringComparison.OrdinalIgnoreCase);
                    var rowEnd = xml.IndexOf(
                        "</w:tr>",
                        visibleStart.Index,
                        StringComparison.OrdinalIgnoreCase);
                    if (rowStart < tableStart
                        || rowStart <= precedingRowEnd
                        || rowEnd <= visibleStart.Index
                        || rowEnd >= tableEnd)
                        return Fail(
                            $"direct-table-row-owner-missing formulaId={formulaId}");
                    rowEnd += "</w:tr>".Length;
                    var rowXml = xml.Substring(rowStart, rowEnd - rowStart);
                    var cellStarts = Regex.Matches(
                            rowXml,
                            @"<w:tc(?:\s|>)",
                            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                        .Cast<Match>()
                        .ToArray();
                    if (cellStarts.Length != 3)
                        return Fail(
                            $"direct-table-row-dimensions formulaId={formulaId} cells={cellStarts.Length}");
                    string ReadCellXml(int cellIndex)
                    {
                        var cellStart = cellStarts[cellIndex].Index;
                        var cellEnd = rowXml.IndexOf(
                            "</w:tc>",
                            cellStart,
                            StringComparison.OrdinalIgnoreCase);
                        return cellEnd <= cellStart
                            ? string.Empty
                            : rowXml.Substring(
                                cellStart,
                                cellEnd + "</w:tc>".Length - cellStart);
                    }
                    var leftCellXml = ReadCellXml(0);
                    var centerCellXml = ReadCellXml(1);
                    var rightCellXml = ReadCellXml(2);
                    if (string.IsNullOrEmpty(leftCellXml)
                        || string.IsNullOrEmpty(centerCellXml)
                        || string.IsNullOrEmpty(rightCellXml))
                        return Fail($"direct-table-cell-end-missing formulaId={formulaId}");
                    var centerMathCount = Regex.Matches(
                        centerCellXml,
                        @"<m:oMath(?:\s|>)",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                    var centerDisplayCount = Regex.Matches(
                        centerCellXml,
                        @"<m:oMathPara(?:\s|>)",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                    var centerFieldCount = Regex.Matches(
                        centerCellXml,
                        @"<w:fldChar\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                    if (centerMathCount != 1
                        || centerDisplayCount != 1
                        || centerFieldCount != 0)
                        return Fail(
                            $"direct-table-center-invalid formulaId={formulaId} math={centerMathCount} display={centerDisplayCount} fields={centerFieldCount}");
                    var formulaBookmarkName =
                        WordOmmlFormulaStore.BookmarkName(formulaId);
                    // VTEqNum_<FormulaId> in this row's cell (3) is the durable
                    // physical owner of current numbered OMML. Word can normalize
                    // the zero-length VTOMML_* convenience anchor from the start of
                    // this row's center paragraph to the row boundary immediately
                    // before that cell, especially after MathType→OMML batch
                    // conversion. Accept that serialization only when the anchor
                    // remains uniquely inside this same row; do not make an unstable
                    // collapsed bookmark carry the formula's primary identity.
                    var formulaBookmarkMatches = Regex.Matches(
                        rowXml,
                        $@"<w:bookmarkStart\b(?=[^>]*\bw:name=""{Regex.Escape(formulaBookmarkName)}"")[^>]*/?>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    if (formulaBookmarkMatches.Count != 1)
                        return Fail(
                            $"direct-table-formula-identity-missing formulaId={formulaId} count={formulaBookmarkMatches.Count}");
                    var expectedVisibleBookmark = EquationBookmarkPrefix + normalizedFormulaId;
                    var expectedNumberBookmark = NativeNumberBookmarkPrefix + normalizedFormulaId;
                    var expectedCaptionBookmark = NativeCaptionBookmarkPrefix + normalizedFormulaId;
                    bool RightCellHasBookmark(string bookmarkName) => Regex.IsMatch(
                        rightCellXml,
                        $@"<w:bookmarkStart\b(?=[^>]*\bw:name=""{Regex.Escape(bookmarkName)}"")[^>]*/?>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasDirectSequence = Regex.IsMatch(
                        rightCellXml,
                        @"\bSEQ\s+(?:&quot;|"")?VisualTeXEquation(?:&quot;|"")?\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var directTableHasRightTabStop = Regex.IsMatch(
                        rightCellXml,
                        @"<w:tab\b(?=[^>]*\bw:val=""right"")(?=[^>]*\bw:pos=""(?<pos>\d+)"")[^>]*/>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasLayoutTab = Regex.IsMatch(
                        rightCellXml,
                        @"<w:tab\s*/>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasMathInRightCell = Regex.IsMatch(
                        rightCellXml,
                        @"<m:oMath(?:\s|>)",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var hasRefInRightCell = Regex.IsMatch(
                        rightCellXml,
                        @"\bREF\s+VTEqNum_",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var rightSequenceCount = Regex.Matches(
                        rightCellXml,
                        @"\bSEQ\s+(?:&quot;|"")?VisualTeXEquation(?:&quot;|"")?\b",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                    var gridWidths = Regex.Matches(
                            tableXml,
                            @"<w:gridCol\b(?=[^>]*\bw:w=""(?<width>\d+)"")[^>]*/>",
                            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                        .Cast<Match>()
                        .Select(match => int.TryParse(
                            match.Groups["width"].Value,
                            out var width) ? width : -1)
                        .ToArray();
                    var rightTabMatch = Regex.Match(
                        rightCellXml,
                        @"<w:tab\b(?=[^>]*\bw:val=""right"")(?=[^>]*\bw:pos=""(?<pos>\d+)"")[^>]*/>",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    var rightTabMatchesColumn = gridWidths.Length == 3
                        && gridWidths[0] > 0
                        && gridWidths[1] > 0
                        && gridWidths[2] == gridWidths[0]
                        && rightTabMatch.Success
                        && int.TryParse(
                            rightTabMatch.Groups["pos"].Value,
                            out var rightTabPosition)
                        && rightTabPosition == gridWidths[2];
                    if (!hasDirectSequence
                        || rightSequenceCount != 1
                        || !RightCellHasBookmark(expectedVisibleBookmark)
                        || !RightCellHasBookmark(expectedNumberBookmark)
                        || !RightCellHasBookmark(expectedCaptionBookmark)
                        || !directTableHasRightTabStop
                        || !hasLayoutTab
                        || hasMathInRightCell
                        || hasRefInRightCell
                        || !rightTabMatchesColumn)
                        return Fail(
                            $"direct-table-right-invalid formulaId={formulaId} seq={hasDirectSequence}/{rightSequenceCount} visible={RightCellHasBookmark(expectedVisibleBookmark)} num={RightCellHasBookmark(expectedNumberBookmark)} cap={RightCellHasBookmark(expectedCaptionBookmark)} rightTab={directTableHasRightTabStop} layoutTab={hasLayoutTab} math={hasMathInRightCell} ref={hasRefInRightCell} grid={string.Join("/", gridWidths)} tabPos={(rightTabMatch.Success ? rightTabMatch.Groups["pos"].Value : "-")}");
                    directTableFormulaIds.Add(formulaId);
                    continue;
                }

                var paragraphStart = LastElementStart("w:p", visibleStart.Index);
                var precedingParagraphEnd = xml.LastIndexOf(
                    "</w:p>",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                var paragraphEnd = xml.IndexOf(
                    "</w:p>",
                    visibleStart.Index,
                    StringComparison.OrdinalIgnoreCase);
                if (paragraphStart < 0
                    || paragraphStart <= precedingParagraphEnd
                    || paragraphEnd <= visibleStart.Index)
                    return Fail($"visible-bookmark-has-no-owner-paragraph formulaId={formulaId}");
                paragraphEnd += "</w:p>".Length;
                if (paragraphEnd - paragraphStart > 262144)
                    return Fail($"numbering-paragraph-too-large formulaId={formulaId}");
                var paragraphXml = xml.Substring(
                    paragraphStart,
                    paragraphEnd - paragraphStart);
                var oleFormulaCount = Regex.Matches(
                    paragraphXml,
                    Regex.Escape("ProgID=\"VisualTeX.Formula.1\""),
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                var ommlFormulaCount = Regex.Matches(
                    paragraphXml,
                    @"<m:oMath(?:\s|>)",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                var ownsOle = oleFormulaCount == 1 && ommlFormulaCount == 0;
                if (!ownsOle)
                    return Fail($"legacy-numbering-paragraph formulaId={formulaId} ole={oleFormulaCount} omml={ommlFormulaCount}");
                var hasCenterTabStop = Regex.IsMatch(
                    paragraphXml,
                    @"<w:tab\b(?=[^>]*\bw:val=""center"")[^>]*/>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                var hasRightTabStop = Regex.IsMatch(
                    paragraphXml,
                    @"<w:tab\b(?=[^>]*\bw:val=""right"")[^>]*/>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                var layoutTabCount = Regex.Matches(
                    paragraphXml,
                    @"<w:tab\s*/>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant).Count;
                if (!hasCenterTabStop || !hasRightTabStop || layoutTabCount < 2)
                {
                    var tabEvidence = string.Join(
                        "|",
                        Regex.Matches(
                                paragraphXml,
                                @"<w:tabs(?:\s[^>]*)?>.*?</w:tabs>|<w:tab\b[^>]*/>",
                                RegexOptions.IgnoreCase
                                | RegexOptions.CultureInvariant
                                | RegexOptions.Singleline)
                            .Cast<Match>()
                            .Select(match => match.Value));
                    return Fail(
                        $"numbering-paragraph-tab-geometry-invalid formulaId={formulaId} owner=ole "
                        + $"center={hasCenterTabStop} right={hasRightTabStop} tabs={layoutTabCount} "
                        + $"evidence={tabEvidence}");
                }
            }

            var startMatches = Regex.Matches(
                xml,
                @"<w:bookmarkStart\b(?=[^>]*\bw:id=""(?<id>-?\d+)"")(?=[^>]*\bw:name=""VTEqNum_(?<guid>[0-9A-F]{32})"")[^>]*/>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (startMatches.Count != numberSet.Count) return Fail($"number-bookmark-start-count-mismatch starts={startMatches.Count} expected={numberSet.Count}");

            // Visible-owner bookmarks move with an equation's tab paragraph.
            // Its external SEQ caption can remain behind on Cut: keep the result
            // by FormulaId, but never use that caption's position as formula order.
            var ownerPositions = visibleStartMatches.ToDictionary(
                match => Guid.ParseExact(match.Groups["guid"].Value, "N").ToString("D"),
                match => match.Index,
                StringComparer.OrdinalIgnoreCase);
            var result = new List<NativeEquationCaptionEntry>(startMatches.Count);
            var seenNumbers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var claimedNumberResultTokens = new HashSet<int>();
            foreach (Match startMatch in startMatches)
            {
                if (!Guid.TryParseExact(
                        startMatch.Groups["guid"].Value,
                        "N",
                        out var formulaGuid))
                    return false;
                var formulaId = formulaGuid.ToString("D");
                if (!numberSet.Contains(formulaId) || !seenNumbers.Add(formulaId))
                    return Fail($"duplicate-or-unknown-number-bookmark formulaId={formulaId}");

                var bookmarkId = Regex.Escape(startMatch.Groups["id"].Value);
                var endMatch = Regex.Match(
                    xml,
                    $@"<w:bookmarkEnd\b[^>]*\bw:id=""{bookmarkId}""[^>]*/>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                    TimeSpan.FromSeconds(1));
                if (!endMatch.Success || endMatch.Index <= startMatch.Index)
                    return Fail($"number-bookmark-end-missing formulaId={formulaId}");
                if (endMatch.Index - startMatch.Index > 16384)
                    return Fail($"number-bookmark-range-too-large formulaId={formulaId}");

                var segment = xml.Substring(
                    startMatch.Index + startMatch.Length,
                    endMatch.Index - startMatch.Index - startMatch.Length);
                var numberText = hashSequenceFormulaIds.Contains(formulaId)
                    ? ExtractNativeHashSequenceResultFromBookmarkXml(segment)
                    : NormalizeNativeEquationNumberText(string.Concat(
                        Regex.Matches(
                                segment,
                                @"<(?:w|m):t(?:\s[^>]*)?>(?<text>.*?)</(?:w|m):t>",
                                RegexOptions.IgnoreCase
                                | RegexOptions.CultureInvariant
                                | RegexOptions.Singleline)
                            .Cast<Match>()
                            .Select(match => System.Net.WebUtility.HtmlDecode(
                                match.Groups["text"].Value))));
                if (string.IsNullOrWhiteSpace(numberText)) return Fail($"number-bookmark-empty formulaId={formulaId}");
                var resultTokens = Regex.Matches(segment,
                    @"<(?:w|m):t(?:\s[^>]*)?>(?<text>.*?)</(?:w|m):t>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline);
                if (resultTokens.Count == 0
                    || !claimedNumberResultTokens.Add(startMatch.Index + startMatch.Length
                        + resultTokens[resultTokens.Count - 1].Index))
                    return Fail($"number-bookmarks-share-physical-result formulaId={formulaId}");
                result.Add(new NativeEquationCaptionEntry(
                    formulaId,
                    ownerPositions[formulaId],
                    numberText));
            }

            if (result.Count != visibleSet.Count) return Fail($"number-entry-count-mismatch entries={result.Count} expected={visibleSet.Count}");

            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var fieldCode in parsedFieldCodes)
            {
                var referenceMatch = Regex.Match(
                    fieldCode,
                    @"\bREF\s+(?:""(?<quoted>VTEqNum_[0-9A-F]{32})""|(?<plain>VTEqNum_[0-9A-F]{32}))(?=\s|\\|$)",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                if (!referenceMatch.Success) continue;
                var bookmarkName = referenceMatch.Groups["quoted"].Success
                    ? referenceMatch.Groups["quoted"].Value
                    : referenceMatch.Groups["plain"].Value;
                var normalizedGuid = bookmarkName.Substring(
                    NativeNumberBookmarkPrefix.Length);
                if (!Guid.TryParseExact(
                        normalizedGuid,
                        "N",
                        out var referenceGuid))
                    continue;
                var formulaId = referenceGuid.ToString("D");
                if (!visibleSet.Contains(formulaId)) continue;
                counts.TryGetValue(formulaId, out var currentCount);
                counts[formulaId] = currentCount + 1;
            }
            foreach (var formulaId in hashSequenceFormulaIds.Concat(directTableFormulaIds))
            {
                // Preserve the historical reference-count contract used by format
                // conversion: one means "no external body reference" and values
                // above one prove external dynamic references exist. Native #SEQ
                // and the current direct-SEQ 1x3 OMML host have no generated visible
                // REF, so account for their own visible number slot virtually.
                counts.TryGetValue(formulaId, out var externalCount);
                counts[formulaId] = externalCount + 1;
            }

            // OLE and legacy Shape hosts still physically own one generated REF;
            // native #SEQ hosts contribute the virtual slot above. Either way the
            // compatibility count must be at least one for every numbered formula.
            if (visibleSet.Any(formulaId =>
                    !counts.TryGetValue(formulaId, out var count) || count < 1))
                return Fail("generated-or-native-number-slot-missing");

            entries = result.OrderBy(entry => entry.Position).ToArray();
            referenceCounts = counts;
            return true;
        }
        catch
        {
            entries = Array.Empty<NativeEquationCaptionEntry>();
            referenceCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            return false;
        }
        finally { Release(content); }
    }

    internal static string ExtractVisibleEquationNumberTextFromOpenXmlSegment(
        string segment)
    {
        if (string.IsNullOrEmpty(segment)) return string.Empty;
        const string wordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string mathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        try
        {
            var wrapped =
                $"<root xmlns:w=\"{wordNamespace}\" xmlns:m=\"{mathNamespace}\">"
                + segment
                + "</root>";
            var document = System.Xml.Linq.XDocument.Parse(
                wrapped,
                System.Xml.Linq.LoadOptions.PreserveWhitespace);
            var word = (System.Xml.Linq.XNamespace)wordNamespace;
            var math = (System.Xml.Linq.XNamespace)mathNamespace;
            var fieldCodeState = new Stack<bool>();
            var visible = new System.Text.StringBuilder();
            foreach (var element in document.Root!.Descendants())
            {
                if (element.Name == word + "fldChar")
                {
                    var kind = (string?)element.Attribute(word + "fldCharType")
                        ?? string.Empty;
                    if (string.Equals(kind, "begin", StringComparison.OrdinalIgnoreCase))
                    {
                        fieldCodeState.Push(true);
                    }
                    else if (string.Equals(
                                 kind,
                                 "separate",
                                 StringComparison.OrdinalIgnoreCase))
                    {
                        if (fieldCodeState.Count > 0)
                        {
                            fieldCodeState.Pop();
                            fieldCodeState.Push(false);
                        }
                    }
                    else if (string.Equals(kind, "end", StringComparison.OrdinalIgnoreCase))
                    {
                        if (fieldCodeState.Count > 0)
                            fieldCodeState.Pop();
                    }
                    continue;
                }

                if (element.Name == word + "instrText")
                    continue;
                if (element.Name != word + "t" && element.Name != math + "t")
                    continue;
                if (fieldCodeState.Any(inCode => inCode))
                    continue;
                visible.Append(element.Value);
            }
            return visible.ToString();
        }
        catch
        {
            // The fast health probe is deliberately conservative. A malformed or
            // namespace-incomplete bookmark segment belongs on full reconciliation;
            // never guess its visible number by concatenating mathematical m:t runs,
            // because those runs can contain the hidden SEQ instruction itself.
            return string.Empty;
        }
    }

    internal static IReadOnlyDictionary<string, int> ReadSingleFormulaReferenceCounts(
        Document document, string formulaId)
    {
        // Source-host validation has already established this formula's number
        // identity. Reference discovery must not revalidate every other numbered
        // host or export their OLE payloads just to count references to this one.
        var names = new[] { EquationBookmarkName(formulaId), NativeNumberBookmarkName(formulaId) };
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        var count = 0;
        try
        {
            fields = document.Fields;
            var fieldCount = fields.Count;
            for (var index = 1; index <= fieldCount; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                var type = field.Type;
                if (type != WdFieldType.wdFieldRef && type != WdFieldType.wdFieldGoToButton) continue;
                code = field.Code;
                var text = code.Text ?? string.Empty;
                if (names.Any(name => System.Text.RegularExpressions.Regex.IsMatch(text,
                        @"(?i)\bREF\s+" + System.Text.RegularExpressions.Regex.Escape(name) + @"(?=\s|\\|$)")))
                    count++;
            }
            return new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) { [formulaId] = count };
        }
        finally { Release(code); Release(field); Release(fields); }
    }

    internal static bool TryGetHealthyEquationReferenceCounts(
        Document document,
        out IReadOnlyDictionary<string, int> referenceCounts)
    {
        // Reuse the same one-pass WordOpenXML validation used by the fast
        // numbering updater. A healthy numbered formula has exactly one generated
        // REF in its own number cell; counts > 1 mean genuine external references
        // also exist and must be frozen before that VisualTeX number is removed.
        return TryReadHealthyEquationNumberArtifactsFromOpenXml(
            document,
            out _,
            out referenceCounts);
    }

    internal static IReadOnlyList<NativeEquationCaptionEntry> ResolveNativeEquationCaptionPositions(
        Document document,
        IReadOnlyList<NativeEquationCaptionEntry> openXmlEntries)
    {
        var result = new List<NativeEquationCaptionEntry>(openXmlEntries.Count);
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var entry in openXmlEntries)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    // VTEq_ moves with the visible host in all current layouts;
                    // VTEqNum_ may point to an external caption left behind by Cut.
                    bookmark = bookmarks[EquationBookmarkName(entry.FormulaId)];
                    range = bookmark.Range;
                    result.Add(new NativeEquationCaptionEntry(
                        entry.FormulaId,
                        range.Start,
                        entry.NumberText));
                }
                catch
                {
                    return Array.Empty<NativeEquationCaptionEntry>();
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }
        return result.OrderBy(item => item.Position).ToArray();
    }

    internal static bool HasOnlyGeneratedEquationReferences(
        Document document,
        int numberedFormulaCount)
    {
        if (!TryReadHealthyEquationNumberArtifactsFromOpenXml(
                document,
                out var entries,
                out var referenceCounts))
            return false;
        return entries.Count == numberedFormulaCount
            && referenceCounts.Count == numberedFormulaCount
            && referenceCounts.Values.All(count => count == 1);
    }

    private static bool HasOnlyGeneratedEquationReferencesLegacy(
        Document document,
        int numberedFormulaCount)
    {
        Range? content = null;
        try
        {
            content = document.Content;
            var xmlText = WordDocumentXml.Read(document) ?? string.Empty;
            var referenceCount = Regex.Matches(
                xmlText,
                @"\bREF\s+VTEqNum_[0-9A-F]{32}\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                .Count;

            // Shape/table/OLE-era layouts own one generated REF per formula. The
            // native OMML #() layout owns none: its visible number is the SEQ result
            // itself. Count those direct mathematical hosts and subtract them from
            // the generated-REF expectation. Any remaining REF is therefore a real
            // body reference and must be refreshed.
            var directHashIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var xml = XDocument.Parse(xmlText, LoadOptions.PreserveWhitespace);
                XNamespace math =
                    "http://schemas.openxmlformats.org/officeDocument/2006/math";
                XNamespace word =
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                foreach (var equation in xml.Descendants(math + "oMath"))
                {
                    var fieldText = string.Concat(
                        equation.Descendants(word + "instrText").Select(node => node.Value)
                        .Concat(equation.Descendants(math + "t").Select(node => node.Value)));
                    if (!IsVisualTeXSequenceFieldCode(fieldText)) continue;
                    foreach (var bookmarkStart in equation.Descendants(word + "bookmarkStart"))
                    {
                        var name = (string?)bookmarkStart.Attribute(word + "name");
                        if (TryFormulaIdFromBookmark(
                                name,
                                NativeNumberBookmarkPrefix,
                                out var formulaId))
                            directHashIds.Add(formulaId);
                    }
                }
            }
            catch
            {
                // Ambiguous XML must take the conservative reference-update path.
                return false;
            }

            var generatedReferenceCount = Math.Max(
                0,
                numberedFormulaCount - directHashIds.Count);
            return referenceCount == generatedReferenceCount;
        }
        catch
        {
            return false;
        }
        finally { Release(content); }
    }

    internal sealed class NativeEquationCaptionEntry
    {
        public NativeEquationCaptionEntry(
            string formulaId,
            int position,
            string numberText)
        {
            FormulaId = formulaId;
            Position = position;
            NumberText = numberText;
        }

        public string FormulaId { get; }
        public int Position { get; }
        public string NumberText { get; }
    }

    internal static bool TryReadCompleteEquationNumberArtifactsFromBookmarks(
        Document document,
        string nativeSequenceName,
        out IReadOnlyList<NativeEquationCaptionEntry> captions)
    {
        captions = Array.Empty<NativeEquationCaptionEntry>();
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            var visibleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var captionIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var numberIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmark);
                bookmark = bookmarks[index];
                var name = bookmark.Name ?? string.Empty;
                if (TryFormulaIdFromBookmark(name, NativeCaptionBookmarkPrefix, out var formulaId))
                {
                    captionIds.Add(formulaId);
                    continue;
                }
                if (TryFormulaIdFromBookmark(name, NativeNumberBookmarkPrefix, out formulaId))
                {
                    numberIds.Add(formulaId);
                    continue;
                }
                if (TryFormulaIdFromBookmark(name, EquationBookmarkPrefix, out formulaId))
                {
                    visibleIds.Add(formulaId);
                    continue;
                }
                if (name.StartsWith(NativeDisplayAnchorBookmarkPrefix, StringComparison.OrdinalIgnoreCase)
                    || name.StartsWith(NativeDisplayAnchorCommitBookmarkPrefix, StringComparison.OrdinalIgnoreCase))
                    WordDoubleClickHook.TraceMessage($"equation-number-scaffold-leftover-anchor name={name}");
            }

            // Numbers are computed from complete triples (visible, caption and
            // number bookmark of one formula). A leftover bookmark of a deleted
            // equation elsewhere is reported, not a reason to refuse numbering.
            var complete = new HashSet<string>(visibleIds, StringComparer.OrdinalIgnoreCase);
            complete.IntersectWith(captionIds);
            complete.IntersectWith(numberIds);
            if (complete.Count == 0)
                return false;
            if (complete.Count != visibleIds.Count || complete.Count != captionIds.Count || complete.Count != numberIds.Count)
                WordDoubleClickHook.TraceMessage(
                    $"equation-number-scaffold-incomplete complete={complete.Count} visible={visibleIds.Count} captions={captionIds.Count} numbers={numberIds.Count}");

            var resolved = GetNativeEquationCaptionEntries(document, nativeSequenceName);
            var resolvedIds = resolved
                .Select(item => item.FormulaId)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (!resolvedIds.SetEquals(complete))
                return false;
            captions = resolved;
            return true;
        }
        catch
        {
            captions = Array.Empty<NativeEquationCaptionEntry>();
            return false;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static IReadOnlyList<NativeEquationCaptionEntry> GetNativeEquationCaptionEntries(
        Document document,
        string nativeSequenceName)
    {
        var result = new List<NativeEquationCaptionEntry>();
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Bookmark? visibleBookmark = null;
                Range? numberRange = null;
                Range? visibleRange = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryFormulaIdFromBookmark(
                            bookmark.Name,
                            NativeNumberBookmarkPrefix,
                            out var formulaId))
                        continue;
                    if (!bookmarks.Exists(NativeCaptionBookmarkName(formulaId))
                        || !bookmarks.Exists(EquationBookmarkName(formulaId)))
                        continue;
                    // The fallback must use the same owner order as the XML path.
                    // Resolve named bookmarks only; do not activate embedded OLE.
                    numberRange = bookmark.Range;
                    visibleBookmark = bookmarks[EquationBookmarkName(formulaId)];
                    visibleRange = visibleBookmark.Range;
                    result.Add(new NativeEquationCaptionEntry(
                        formulaId,
                        visibleRange.Start,
                        numberRange.Text ?? string.Empty));
                }
                finally
                {
                    Release(visibleRange);
                    Release(visibleBookmark);
                    Release(numberRange);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }
        return result.OrderBy(item => item.Position).ToArray();
    }
}
