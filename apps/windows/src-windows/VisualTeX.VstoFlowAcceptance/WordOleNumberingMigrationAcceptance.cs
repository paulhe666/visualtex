using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOleNumberingMigration(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-visualtex-ole-tab-numbering-migration.docx");
        var previewPath = Path.Combine(artifactRoot, "word-visualtex-ole-tab-preview.png");
        var previewDataUrl = CreatePngDataUrl(
            "word-ole-numbering-migration",
            180,
            64);
        File.WriteAllBytes(
            previewPath,
            Convert.FromBase64String(
                previewDataUrl.Substring(previewDataUrl.IndexOf(',') + 1)));

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.Table? legacyTable = null;
        Word.Cell? formulaCell = null;
        Word.Range? formulaCellRange = null;
        Word.Row? emptyRow = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();

            const string latex = @"x^2+y^2=z^2";
            var formulaId = Guid.NewGuid().ToString("D");
            var now = DateTimeOffset.UtcNow.ToString("O");
            var metadata = new FormulaMetadata
            {
                FormulaId = formulaId,
                Title = "VisualTeX OLE tab-numbering migration acceptance",
                Latex = latex,
                CodeFormat = "latex",
                DisplayMode = "block",
                Numbered = true,
                RenderWidthPx = 180,
                RenderHeightPx = 64,
                Baseline = 48,
                FontSizePt = 12,
                RenderFontSizePt = 12,
                FormulaLetterFont = "katex",
                FormulaChineseFont = "system",
                CreatedWithVersion = "1.2.5",
                UpdatedWithVersion = "1.2.5",
                CreatedAt = now,
                UpdatedAt = now,
                Lines = new List<FormulaLine>
                {
                    new()
                    {
                        Id = Guid.NewGuid().ToString("D"),
                        Latex = latex,
                    },
                },
            };
            metadata.Validate();

            insertion = document.Range(0, 0);
            legacyTable = document.Tables.Add(insertion, 1, 3);
            formulaCell = legacyTable.Cell(1, 2);
            formulaCellRange = formulaCell.Range.Duplicate;
            formulaCellRange.End = Math.Max(
                formulaCellRange.Start,
                formulaCellRange.End - 1);
            shape = document.InlineShapes.AddPicture(
                FileName: previewPath,
                LinkToFile: false,
                SaveWithDocument: true,
                Range: formulaCellRange);
            shape.Width = 135f;
            shape.Height = 48f;
            WordFormulaMetadataReader.Write(shape, metadata);
            shapeRange = shape.Range;

            // Construct the complete legacy 2x3 input before invoking production
            // code. BuildFormulaNumberingScaffoldForConversion is now deliberately
            // table-destructive: it accepts this old host only as migration input,
            // trims the benign empty row and immediately creates the final tab-only
            // numbering scaffold. The fixture must therefore never try to append a
            // row after that call through an invalidated legacy Table RCW.
            emptyRow = legacyTable.Rows.Add();
            AssertEqual(2, legacyTable.Rows.Count,
                "The legacy-numbering migration fixture did not create a 2x3 table.");
            AssertEqual(3, legacyTable.Columns.Count,
                "The legacy-numbering migration fixture changed the table column count.");

            WordEquationNumbering.BuildFormulaNumberingScaffoldForConversion(
                document,
                shapeRange,
                shape.Height,
                metadata,
                legacyTable,
                plannedOrdinal: 1,
                plannedPrefix: "0.");
            Release(shapeRange); shapeRange = null;
            Release(shape); shape = null;

            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "structural VisualTeX OLE surrogate",
                requireNativeOle: false,
                requireFormulaMetadata: false);

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "saved/reopened structural VisualTeX OLE surrogate",
                requireNativeOle: false,
                requireFormulaMetadata: false);

            Console.WriteLine(
                "Word VisualTeX OLE numbering migration acceptance passed: the first reconcile trimmed a benign empty legacy row, migrated the 1x3 host to one MathType-style center/right-tab paragraph, and save/reopen kept that structure.");
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            Release(emptyRow);
            Release(formulaCellRange);
            Release(formulaCell);
            Release(legacyTable);
            Release(insertion);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            try { File.Delete(previewPath); } catch { }
            ForceComCleanup();
        }
    }
}
