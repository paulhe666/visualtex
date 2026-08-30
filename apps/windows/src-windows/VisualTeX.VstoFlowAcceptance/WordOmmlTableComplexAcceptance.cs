using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class OmmlTableComplexPair
    {
        internal OmmlTableComplexPair(
            NumberedOmmlDisplayCase testCase,
            string numberedId,
            string plainId,
            string semanticSignature)
        {
            TestCase = testCase;
            NumberedId = numberedId;
            PlainId = plainId;
            SemanticSignature = semanticSignature;
        }
        internal NumberedOmmlDisplayCase TestCase { get; }
        internal string NumberedId { get; }
        internal string PlainId { get; }
        internal string SemanticSignature { get; }
    }

    private static void RunWordOmmlTableComplexAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The complex 1x3 acceptance refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(artifactRoot, "word-omml-1x3-complex.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            ConfigureOmmlTableNumberPage(document);
            document.SaveAs2(
                documentPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);
            var pairs = new List<OmmlTableComplexPair>();

            foreach (var testCase in CreateNumberedOmmlDisplayCases())
            {
                var numberedId = Guid.NewGuid().ToString("D");
                InsertComplexDisplayFormula(
                    application,
                    document,
                    service,
                    testCase,
                    numberedId,
                    numbered: true);
                AssertOmmlTableNumberLifecyclePhase(
                    application,
                    document,
                    numberedId,
                    "complex numbered " + testCase.Name);

                var plainId = Guid.NewGuid().ToString("D");
                InsertComplexDisplayFormula(
                    application,
                    document,
                    service,
                    testCase,
                    plainId,
                    numbered: false);
                var signature = AssertComplexTablePair(
                    application,
                    document,
                    testCase,
                    numberedId,
                    plainId,
                    "initial");
                pairs.Add(new OmmlTableComplexPair(
                    testCase,
                    numberedId,
                    plainId,
                    signature));
            }

            AssertEqual(pairs.Count, document.Tables.Count,
                "Complex 1x3 acceptance created a wrong table count.");
            document.Fields.Update();
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();

            AssertEqual(pairs.Count, document.Tables.Count,
                "Complex 1x3 save/reopen changed the table count.");
            foreach (var pair in pairs)
            {
                AssertOmmlTableNumberLifecyclePhase(
                    application,
                    document,
                    pair.NumberedId,
                    "complex reopen " + pair.TestCase.Name);
                var reopenedSignature = AssertComplexTablePair(
                    application,
                    document,
                    pair.TestCase,
                    pair.NumberedId,
                    pair.PlainId,
                    "save/reopen");
                AssertEqual(pair.SemanticSignature, reopenedSignature,
                    pair.TestCase.Name + ": save/reopen changed the semantic OMath structure.");
            }

            AssertDirectTableEditRefreshesMinimumHeight(
                application,
                document,
                service,
                pairs[0],
                pairs[pairs.Count - 1]);
            AssertLegacyAutoHeightTableRepairsOnOpenRefresh(
                document,
                pairs[2]);

            Console.WriteLine(
                "Word OMML complex 1x3 acceptance passed: fraction/radical, n-ary limits, nested fractions and matrices retained the same semantic OMath structure, native font size and horizontal glyph geometry as ordinary wdOMathDisplay references, while numbering stayed outside math across F9/save/reopen; Word's table-specific vertical character Range boxes were recorded separately rather than treated as glyph scaling.");
        }
        finally
        {
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

    private static string AssertComplexTablePair(
        Word.Application application,
        Word.Document document,
        NumberedOmmlDisplayCase testCase,
        string numberedId,
        string plainId,
        string phase)
    {
        Word.Range? numberedRange = null;
        Word.Range? plainRange = null;
        Word.Window? window = null;
        try
        {
            var numberedMetadata = WordOmmlFormulaStore.TryRead(document, numberedId)
                ?? throw new InvalidDataException(testCase.Name + ": numbered metadata missing.");
            var plainMetadata = WordOmmlFormulaStore.TryRead(document, plainId)
                ?? throw new InvalidDataException(testCase.Name + ": plain metadata missing.");
            numberedRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document, numberedId, numberedMetadata);
            plainRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document, plainId, plainMetadata);

            AssertEqual(Word.WdOMathType.wdOMathDisplay, numberedRange.OMaths[1].Type,
                phase + " " + testCase.Name + ": numbered formula is not Display.");
            AssertEqual(Word.WdOMathType.wdOMathDisplay, plainRange.OMaths[1].Type,
                phase + " " + testCase.Name + ": plain formula is not Display.");
            AssertEqual(0, numberedRange.Fields.Count,
                phase + " " + testCase.Name + ": a number field entered the complex OMath.");
            AssertTrue(!WordOmmlConverter.HasVisualTeXNativeEquationNumber(
                    numberedRange.WordOpenXML ?? string.Empty),
                phase + " " + testCase.Name + ": retired #()/eqArr numbering entered the complex OMath.");

            const string MathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var math = (XNamespace)MathNamespace;
            var numberedEquation = XElement.Parse(
                WordOmmlConverter.ExtractSingleOMath(numberedRange.WordOpenXML),
                LoadOptions.PreserveWhitespace);
            var plainEquation = XElement.Parse(
                WordOmmlConverter.ExtractSingleOMath(plainRange.WordOpenXML),
                LoadOptions.PreserveWhitespace);
            var numberedSignature = BuildOmmlSemanticStructureSignature(numberedEquation);
            var plainSignature = BuildOmmlSemanticStructureSignature(plainEquation);
            AssertEqual(plainSignature, numberedSignature,
                phase + " " + testCase.Name + ": 1x3 changed semantic OMath structure.");
            foreach (var requirement in testCase.RequiredElements)
            {
                var count = numberedEquation
                    .DescendantsAndSelf(math + requirement.ElementName)
                    .Count();
                AssertTrue(count >= requirement.MinimumCount,
                    phase + " " + testCase.Name + $": expected m:{requirement.ElementName}>={requirement.MinimumCount}, found {count}.");
            }

            window = document.ActiveWindow;
            object scrollStart = true;
            window.ScrollIntoView(numberedRange, ref scrollStart);
            document.Repaginate();
            Thread.Sleep(70);
            var numberedBox = ReadVisibleMathInkBox(
                document,
                window,
                numberedRange,
                phase + " " + testCase.Name + " numbered ink");
            window.ScrollIntoView(plainRange, ref scrollStart);
            Thread.Sleep(70);
            var plainBox = ReadVisibleMathInkBox(
                document,
                window,
                plainRange,
                phase + " " + testCase.Name + " plain ink");
            AssertMetricRatio(plainBox.Width, numberedBox.Width, 0.96, 1.04,
                phase + " " + testCase.Name + ": 1x3 changed formula width.");
            var numberedFontSize = numberedRange.Font.Size;
            var plainFontSize = plainRange.Font.Size;
            AssertNear(plainFontSize, numberedFontSize, 0.1f,
                phase + " " + testCase.Name + ": 1x3 changed the native OMath font size.");
            var heightRatio = plainBox.Height > 0
                ? numberedBox.Height / (double)plainBox.Height
                : 1.0;
            if (string.Equals(phase, "initial", StringComparison.Ordinal))
                AssertNumberedOmmlTableCoversPlainDisplayAcrossZoom(
                    document,
                    numberedRange,
                    plainRange,
                    testCase.Name);
            Console.WriteLine(
                $"  {phase} complex {testCase.Name}: numbered={numberedBox.Width}x{numberedBox.Height}px, plain={plainBox.Width}x{plainBox.Height}px, heightRangeRatio={heightRatio:0.###}, font={numberedFontSize:0.##}/{plainFontSize:0.##}pt, signature={numberedSignature.Length}.");
            return numberedSignature;
        }
        finally
        {
            Release(window);
            Release(plainRange);
            Release(numberedRange);
        }
    }

    private static void AssertDirectTableEditRefreshesMinimumHeight(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        OmmlTableComplexPair sourcePair,
        OmmlTableComplexPair tallReferencePair)
    {
        var metadata = WordOmmlFormulaStore.TryRead(document, sourcePair.NumberedId)
            ?? throw new InvalidDataException("Direct-table height edit source metadata is missing.");
        Word.Range? sourceRange = null;
        Word.Tables? sourceTables = null;
        Word.Table? sourceTable = null;
        Word.Rows? rows = null;
        Word.Row? row = null;
        Word.Range? editedRange = null;
        Word.Range? plainTallRange = null;
        try
        {
            sourceRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                sourcePair.NumberedId,
                metadata);
            sourceTables = sourceRange.Tables;
            AssertEqual(1, sourceTables.Count,
                "Direct-table height edit source is not in one 1x3 table.");
            sourceTable = sourceTables[1];
            rows = sourceTable.Rows;
            row = rows[1];
            var beforeHeight = row.Height;

            var lineId = metadata.Lines.FirstOrDefault()?.Id;
            if (string.IsNullOrWhiteSpace(lineId)) lineId = Guid.NewGuid().ToString("D");
            var editSession = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Mode = "edit",
                Host = "word",
                FormulaId = sourcePair.NumberedId,
                SourceDocumentId = document.FullName,
                SourceObjectId = WordRangeReference(sourceRange.Start, sourceRange.End),
                Title = "VisualTeX 1x3 dynamic-height edit acceptance",
                CodeFormat = "latex",
                DisplayMode = "block",
                ObjectMode = FormulaOleContract.WordOmmlMode,
                Numbered = true,
                FontSizePt = 14,
                OriginalMetadata = metadata,
                Lines = new List<FormulaLine>
                {
                    new() { Id = lineId!, Latex = tallReferencePair.TestCase.Latex },
                },
                ExportResult = new OfficeExportDocument
                {
                    FormulaLetterFont = "katex",
                    FormulaChineseFont = "system",
                },
            };
            service.ReplaceOmml(editSession, tallReferencePair.TestCase.MathMl);

            Release(row); row = null;
            Release(rows); rows = null;
            Release(sourceTable); sourceTable = null;
            Release(sourceTables); sourceTables = null;
            Release(sourceRange); sourceRange = null;

            var editedMetadata = WordOmmlFormulaStore.TryRead(document, sourcePair.NumberedId)
                ?? throw new InvalidDataException("Direct-table height edit lost metadata.");
            editedRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                sourcePair.NumberedId,
                editedMetadata);
            sourceTables = editedRange.Tables;
            sourceTable = sourceTables[1];
            rows = sourceTable.Rows;
            row = rows[1];
            AssertTrue(row.Height > beforeHeight + 1f,
                $"Direct-table edit did not grow its native-display row minimum: before={beforeHeight:0.###}pt after={row.Height:0.###}pt.");

            var plainMetadata = WordOmmlFormulaStore.TryRead(document, tallReferencePair.PlainId)
                ?? throw new InvalidDataException("Tall plain-display comparison metadata is missing.");
            plainTallRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                tallReferencePair.PlainId,
                plainMetadata);
            AssertNumberedOmmlTableCoversPlainDisplayAcrossZoom(
                document,
                editedRange,
                plainTallRange,
                "direct-edit-to-" + tallReferencePair.TestCase.Name);
            Console.WriteLine(
                $"  direct-table height edit: {sourcePair.TestCase.Name} {beforeHeight:0.###}pt -> {tallReferencePair.TestCase.Name} {row.Height:0.###}pt.");
        }
        finally
        {
            Release(plainTallRange);
            Release(editedRange);
            Release(row);
            Release(rows);
            Release(sourceTable);
            Release(sourceTables);
            Release(sourceRange);
        }
    }

    private static void AssertLegacyAutoHeightTableRepairsOnOpenRefresh(
        Word.Document document,
        OmmlTableComplexPair pair)
    {
        var metadata = WordOmmlFormulaStore.TryRead(document, pair.NumberedId)
            ?? throw new InvalidDataException("Legacy auto-height repair metadata is missing.");
        var plainMetadata = WordOmmlFormulaStore.TryRead(document, pair.PlainId)
            ?? throw new InvalidDataException("Legacy auto-height repair plain metadata is missing.");
        Word.Range? numberedRange = null;
        Word.Range? plainRange = null;
        Word.Tables? tables = null;
        Word.Table? table = null;
        Word.Rows? rows = null;
        Word.Row? row = null;
        try
        {
            numberedRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                pair.NumberedId,
                metadata);
            tables = numberedRange.Tables;
            AssertEqual(1, tables.Count,
                "Legacy auto-height repair target is not inside one 1x3 table.");
            table = tables[1];
            rows = table.Rows;
            row = rows[1];
            row.HeightRule = Word.WdRowHeightRule.wdRowHeightAuto;
            AssertEqual(
                Word.WdRowHeightRule.wdRowHeightAuto,
                row.HeightRule,
                "The legacy auto-height fixture did not reproduce the old row rule.");

            var refreshed = WordEquationNumbering.RefreshNumberedOmmlTabLayouts(document);
            AssertTrue(refreshed > 0,
                "Document-open layout refresh did not visit the legacy auto-height fixture.");

            Release(row); row = null;
            Release(rows); rows = null;
            Release(table); table = null;
            Release(tables); tables = null;
            Release(numberedRange); numberedRange = null;
            numberedRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                pair.NumberedId,
                metadata);
            tables = numberedRange.Tables;
            table = tables[1];
            rows = table.Rows;
            row = rows[1];
            AssertEqual(
                Word.WdRowHeightRule.wdRowHeightAtLeast,
                row.HeightRule,
                "Document-open refresh did not migrate the legacy auto-height 1x3 row to AtLeast.");
            AssertTrue(row.Height > 0f && row.Height < 1000f,
                $"Document-open refresh produced an invalid row minimum {row.Height:0.###}pt.");

            plainRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                pair.PlainId,
                plainMetadata);
            AssertNumberedOmmlTableCoversPlainDisplayAcrossZoom(
                document,
                numberedRange,
                plainRange,
                "legacy-auto-repair-" + pair.TestCase.Name);
            Console.WriteLine(
                $"  legacy auto-height 1x3 repaired on document-open refresh: {pair.TestCase.Name} -> {row.Height:0.###}pt AtLeast.");
        }
        finally
        {
            Release(row);
            Release(rows);
            Release(table);
            Release(tables);
            Release(plainRange);
            Release(numberedRange);
        }
    }

    private static void AssertNumberedOmmlTableCoversPlainDisplayAcrossZoom(
        Word.Document document,
        Word.Range numberedRange,
        Word.Range plainRange,
        string formulaName)
    {
        Word.Tables? tables = null;
        Word.Table? table = null;
        Word.Range? tableRange = null;
        Word.Rows? rows = null;
        Word.Row? row = null;
        Word.Window? window = null;
        try
        {
            tables = numberedRange.Tables;
            AssertEqual(1, tables.Count,
                formulaName + ": numbered OMML is not inside exactly one 1x3 table.");
            table = tables[1];
            AssertNear(0f, table.TopPadding, 0.01f,
                formulaName + ": numbered OMML top padding should remain zero.");
            AssertNear(0f, table.BottomPadding, 0.01f,
                formulaName + ": numbered OMML bottom padding should remain zero.");
            rows = table.Rows;
            AssertEqual(1, rows.Count,
                formulaName + ": numbered OMML table no longer has one row.");
            row = rows[1];
            AssertEqual(Word.WdRowHeightRule.wdRowHeightAtLeast, row.HeightRule,
                formulaName + ": numbered OMML row is not using its native-display minimum height.");
            AssertTrue(row.Height > 0f && row.Height < 1000000f,
                formulaName + ": numbered OMML row minimum height is invalid.");
            tableRange = table.Range;
            window = document.ActiveWindow;
            var originalZoom = window.View.Zoom.Percentage;
            object scrollStart = true;
            try
            {
                foreach (var zoom in new[] { 75, 100, 125, 150, 185, 225 })
                {
                    window.View.Zoom.Percentage = zoom;
                    document.Repaginate();
                    window.ScrollIntoView(tableRange, ref scrollStart);
                    Thread.Sleep(45);
                    window.GetPoint(
                        out _, out _, out _, out var tableHeight,
                        tableRange);
                    window.GetPoint(
                        out _, out _, out _, out var numberedMathHeight,
                        numberedRange);
                    window.ScrollIntoView(plainRange, ref scrollStart);
                    Thread.Sleep(45);
                    window.GetPoint(
                        out _, out _, out _, out var plainHeight,
                        plainRange);
                    AssertTrue(
                        tableHeight >= plainHeight,
                        formulaName
                        + $": at {zoom}% the 1x3 host is only {tableHeight}px high, "
                        + $"below the ordinary Word display height {plainHeight}px; "
                        + "the formula can be clipped at the table boundary.");
                    Console.WriteLine(
                        $"    zoom-height {formulaName} {zoom}%: table={tableHeight}px numberedMath={numberedMathHeight}px plain={plainHeight}px tableMargin={tableHeight - plainHeight}px mathRatio={(plainHeight > 0 ? numberedMathHeight / (double)plainHeight : 0):0.###} rowMin={row.Height:0.###}pt.");
                }
            }
            finally
            {
                try { window.View.Zoom.Percentage = originalZoom; } catch { }
            }
        }
        finally
        {
            Release(window);
            Release(row);
            Release(rows);
            Release(tableRange);
            Release(table);
            Release(tables);
        }
    }
}
