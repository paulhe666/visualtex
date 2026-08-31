using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

internal static partial class WordEquationNumbering
{
    [ThreadStatic]
    private static int _targetedNumberingMutationDepth;

    [ThreadStatic]
    private static string? _targetedNumberingFormulaId;

    [ThreadStatic]
    private static bool _targetedNumberingExpectedNumbered;

    internal static IDisposable BeginTargetedNumberingMutation(
        string formulaId,
        bool expectedNumbered)
    {
        var scope = new TargetedNumberingMutationScope(
            _targetedNumberingMutationDepth,
            _targetedNumberingFormulaId,
            _targetedNumberingExpectedNumbered);
        _targetedNumberingMutationDepth += 1;
        _targetedNumberingFormulaId = formulaId;
        _targetedNumberingExpectedNumbered = expectedNumbered;
        return scope;
    }

    private static bool IsTargetedNumberingMutationActive =>
        _targetedNumberingMutationDepth > 0;

    private sealed class TargetedNumberingMutationScope : IDisposable
    {
        private readonly int _previousDepth;
        private readonly string? _previousFormulaId;
        private readonly bool _previousExpectedNumbered;
        private bool _disposed;

        internal TargetedNumberingMutationScope(
            int previousDepth,
            string? previousFormulaId,
            bool previousExpectedNumbered)
        {
            _previousDepth = previousDepth;
            _previousFormulaId = previousFormulaId;
            _previousExpectedNumbered = previousExpectedNumbered;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _targetedNumberingMutationDepth = _previousDepth;
            _targetedNumberingFormulaId = _previousFormulaId;
            _targetedNumberingExpectedNumbered = _previousExpectedNumbered;
        }
    }

    /// <summary>
    /// Prevents Word from repainting intermediate selections while a numbering
    /// operation reshapes a display equation or updates its fields. Word can
    /// otherwise show a transient jump to the first story range and repeatedly
    /// lay out every unrelated OMML formula in a large document.
    /// </summary>
    private sealed class ScreenUpdatingScope : IDisposable
    {
        private Microsoft.Office.Interop.Word.Application? _application;
        private readonly bool _previous;
        private readonly bool _active;

        private ScreenUpdatingScope(Microsoft.Office.Interop.Word.Application? application)
        {
            _application = application;
            if (application is null) return;
            try
            {
                _previous = application.ScreenUpdating;
                application.ScreenUpdating = false;
                _active = true;
            }
            catch
            {
                _application = null;
            }
        }

        internal static ScreenUpdatingScope Suspend(Document document)
        {
            Microsoft.Office.Interop.Word.Application? application = null;
            try { application = document.Application; }
            catch { }
            return new ScreenUpdatingScope(application);
        }

        public void Dispose()
        {
            var application = _application;
            _application = null;
            if (!_active || application is null) return;
            try { application.ScreenUpdating = _previous; }
            catch { }
        }
    }

    private static bool IsInteractiveNumberingCommandCall()
    {
        try
        {
            foreach (var frame in new StackTrace(false).GetFrames()
                         ?? Array.Empty<StackFrame>())
            {
                var method = frame.GetMethod();
                var name = method?.Name ?? string.Empty;
                if (name.IndexOf("Number", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                if (name.IndexOf("Update", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("Refresh", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("Format", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("Apply", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("Change", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("Set", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
        }
        catch { }
        return false;
    }

    /// <summary>
    /// Fast path for the Ribbon "update numbering" action once every numbered
    /// equation already uses the canonical 1x3 direct-SEQ host. It deliberately
    /// indexes only VTEqNum_ bookmarks and their right cells; unrelated OMML
    /// formulas are never opened or inspected. Any legacy or damaged host falls
    /// back to the full reconciliation/migration path.
    /// </summary>
    private static bool TryFastUpdateCanonicalNumberFields(Document document)
    {
        if (document is null) return false;
        using var screenUpdating = ScreenUpdatingScope.Suspend(document);

        Bookmarks? bookmarks = null;
        Fields? documentFields = null;
        var numberedHosts = new List<PerformanceCanonicalNumberedHost>();
        var numberedBookmarkNames = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase);
        var documentFieldObjects = new List<Field>();
        var referenceFields = new List<Field>();
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? bookmarkRange = null;
                Tables? tables = null;
                Table? table = null;
                Range? tableRange = null;
                Cell? numberCell = null;
                Range? numberRange = null;
                Fields? numberFields = null;
                try
                {
                    bookmark = bookmarks[index];
                    var name = bookmark.Name ?? string.Empty;
                    if (!name.StartsWith("VTEqNum_", StringComparison.OrdinalIgnoreCase))
                        continue;
                    numberedBookmarkNames.Add(name);

                    bookmarkRange = bookmark.Range.Duplicate;
                    tables = bookmarkRange.Tables;
                    if (tables.Count != 1) return false;
                    table = tables[1];
                    if (table.Rows.Count != 1 || table.Columns.Count != 3)
                        return false;

                    numberCell = table.Cell(1, 3);
                    numberRange = numberCell.Range.Duplicate;
                    numberFields = numberRange.Fields;
                    var fields = new List<Field>();
                    var directSequenceCode = string.Empty;
                    for (var fieldIndex = 1; fieldIndex <= numberFields.Count; fieldIndex++)
                    {
                        var field = numberFields[fieldIndex];
                        var code = PerformanceReadFieldCode(field);
                        if (PerformanceIsVisualTeXSequenceField(code))
                            directSequenceCode = code;
                        fields.Add(field);
                    }
                    if (string.IsNullOrWhiteSpace(directSequenceCode))
                    {
                        PerformanceReleaseFields(fields);
                        return false;
                    }

                    tableRange = table.Range;
                    numberedHosts.Add(
                        new PerformanceCanonicalNumberedHost(
                            tableRange.Start,
                            fields,
                            directSequenceCode,
                            numberRange.Text ?? string.Empty));
                }
                finally
                {
                    PerformanceReleaseComObject(numberFields);
                    PerformanceReleaseComObject(numberRange);
                    PerformanceReleaseComObject(numberCell);
                    PerformanceReleaseComObject(tableRange);
                    PerformanceReleaseComObject(table);
                    PerformanceReleaseComObject(tables);
                    PerformanceReleaseComObject(bookmarkRange);
                    PerformanceReleaseComObject(bookmark);
                }
            }

            if (IsTargetedNumberingMutationActive
                && !string.IsNullOrWhiteSpace(_targetedNumberingFormulaId))
            {
                var targetBookmark = NativeNumberBookmarkName(
                    _targetedNumberingFormulaId!);
                var targetIsCanonical = numberedBookmarkNames.Contains(targetBookmark);
                if (targetIsCanonical != _targetedNumberingExpectedNumbered)
                    return false;
            }

            documentFields = document.Fields;
            var documentSequenceCount = 0;
            for (var index = 1; index <= documentFields.Count; index++)
            {
                var field = documentFields[index];
                documentFieldObjects.Add(field);
                var code = PerformanceReadFieldCode(field);
                if (PerformanceIsVisualTeXSequenceField(code))
                {
                    documentSequenceCount += 1;
                    continue;
                }
                if (PerformanceIsVisualTeXReferenceField(code))
                    referenceFields.Add(field);
            }

            // A direct SEQ without a VTEqNum_ bookmark is a legacy/damaged host;
            // let the full reconciliation path repair it instead of hiding it.
            if (documentSequenceCount != numberedHosts.Count)
                return false;
            if (!PerformanceCanonicalHostsMatchDocumentFormat(document, numberedHosts))
                return false;

            foreach (var host in numberedHosts.OrderBy(item => item.Start))
            {
                foreach (var field in host.Fields)
                {
                    try { field.Update(); }
                    catch { return false; }
                }
            }
            foreach (var field in referenceFields.OrderBy(PerformanceFieldStart))
            {
                try { field.Update(); }
                catch { return false; }
            }
            return true;
        }
        finally
        {
            foreach (var host in numberedHosts)
                PerformanceReleaseFields(host.Fields);
            PerformanceReleaseFields(documentFieldObjects);
            PerformanceReleaseComObject(documentFields);
            PerformanceReleaseComObject(bookmarks);
        }
    }

    private static bool PerformanceCanonicalHostsMatchDocumentFormat(
        Document document,
        IReadOnlyList<PerformanceCanonicalNumberedHost> hosts)
    {
        if (hosts.Count == 0) return true;

        Variables? variables = null;
        Variable? formatVariable = null;
        var formatId = string.Empty;
        try
        {
            variables = document.Variables;
            try
            {
                formatVariable = variables[EquationNumberFormatVariableName];
                formatId = formatVariable.Value ?? string.Empty;
            }
            catch
            {
                formatId = string.Empty;
            }
        }
        finally
        {
            PerformanceReleaseComObject(formatVariable);
            PerformanceReleaseComObject(variables);
        }

        var resetLevel =
            string.Equals(formatId, EquationNumberFormat.Heading1DotId, StringComparison.OrdinalIgnoreCase)
            || string.Equals(formatId, EquationNumberFormat.Heading1DashId, StringComparison.OrdinalIgnoreCase)
                ? 1
                : string.Equals(formatId, EquationNumberFormat.Heading2DotId, StringComparison.OrdinalIgnoreCase)
                  || string.Equals(formatId, EquationNumberFormat.Heading2DashId, StringComparison.OrdinalIgnoreCase)
                    ? 2
                    : 0;
        var expectsDot =
            string.Equals(formatId, EquationNumberFormat.Heading1DotId, StringComparison.OrdinalIgnoreCase)
            || string.Equals(formatId, EquationNumberFormat.Heading2DotId, StringComparison.OrdinalIgnoreCase);
        var expectsDash =
            string.Equals(formatId, EquationNumberFormat.Heading1DashId, StringComparison.OrdinalIgnoreCase)
            || string.Equals(formatId, EquationNumberFormat.Heading2DashId, StringComparison.OrdinalIgnoreCase);

        foreach (var host in hosts)
        {
            var resetMatch = Regex.Match(
                host.SequenceCode,
                @"\\s\s+([12])(?:\s|$)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            var actualResetLevel = resetMatch.Success
                && int.TryParse(resetMatch.Groups[1].Value, out var parsedLevel)
                    ? parsedLevel
                    : 0;
            if (actualResetLevel != resetLevel) return false;

            var visibleNumber = host.NumberText
                .Replace("\r", string.Empty)
                .Replace("\a", string.Empty)
                .Trim();
            if (expectsDash && visibleNumber.IndexOf('-') < 0) return false;
            if (expectsDot && visibleNumber.IndexOf('.') < 0) return false;
            if (!expectsDot && !expectsDash
                && Regex.IsMatch(
                    visibleNumber,
                    @"\p{N}\s*[.\-]\s*\p{N}",
                    RegexOptions.CultureInvariant))
                return false;
        }
        return true;
    }

    private sealed class PerformanceCanonicalNumberedHost
    {
        internal PerformanceCanonicalNumberedHost(
            int start,
            IReadOnlyList<Field> fields,
            string sequenceCode,
            string numberText)
        {
            Start = start;
            Fields = fields;
            SequenceCode = sequenceCode;
            NumberText = numberText;
        }

        internal int Start { get; }
        internal IReadOnlyList<Field> Fields { get; }
        internal string SequenceCode { get; }
        internal string NumberText { get; }
    }

    private static int PerformanceFieldStart(Field field)
    {
        Range? result = null;
        try
        {
            result = field.Result;
            return result.Start;
        }
        catch { return int.MaxValue; }
        finally { PerformanceReleaseComObject(result); }
    }

    private static string PerformanceReadFieldCode(Field field)
    {
        Range? code = null;
        try
        {
            code = field.Code;
            return code.Text ?? string.Empty;
        }
        catch { return string.Empty; }
        finally { PerformanceReleaseComObject(code); }
    }

    private static bool PerformanceIsVisualTeXSequenceField(string code) =>
        code.IndexOf("SEQ", StringComparison.OrdinalIgnoreCase) >= 0
        && code.IndexOf("VisualTeXEquation", StringComparison.OrdinalIgnoreCase) >= 0;

    private static bool PerformanceIsVisualTeXReferenceField(string code) =>
        code.IndexOf("REF", StringComparison.OrdinalIgnoreCase) >= 0
        && (code.IndexOf("VTEqCap_", StringComparison.OrdinalIgnoreCase) >= 0
            || code.IndexOf("VTEqNum_", StringComparison.OrdinalIgnoreCase) >= 0
            || code.IndexOf("VTEq_", StringComparison.OrdinalIgnoreCase) >= 0);

    private static void PerformanceReleaseFields(IEnumerable<Field> fields)
    {
        foreach (var field in fields) PerformanceReleaseComObject(field);
    }

    private static void PerformanceReleaseComObject(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.FinalReleaseComObject(value); }
        catch { }
    }
}
