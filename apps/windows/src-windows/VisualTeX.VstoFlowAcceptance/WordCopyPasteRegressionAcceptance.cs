using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordCopyPasteRegressionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var oleAssetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(oleAssetRoot);

        var inlineAssets = CreatePerformanceOleAssets(
            oleAssetRoot,
            9401,
            "x+y",
            PerformanceSvg(9401));
        var blockAssets = CreatePerformanceOleAssets(
            oleAssetRoot,
            9402,
            "u+v",
            PerformanceSvg(9402));
        var mathTypeAssets = CreatePerformanceOleAssets(
            oleAssetRoot,
            9403,
            "a+b",
            PerformanceSvg(9403));

        RunVisualTeXPastePersistenceCase(
            artifactRoot,
            "inline",
            "x+y",
            inlineAssets.PngPath,
            inlineAssets.EmfPath,
            9401);
        RunVisualTeXPastePersistenceCase(
            artifactRoot,
            "block",
            "u+v",
            blockAssets.PngPath,
            blockAssets.EmfPath,
            9402);
        RunOmmlImmediateNumberPasteCase(artifactRoot);
        RunMathTypeImmediateNumberPasteCase(
            artifactRoot,
            mathTypeAssets.EmfPath);

        Console.WriteLine(
            "Copy/paste regression acceptance passed: VisualTeX inline/block OLE copies persisted with fresh identity, and OMML/MathType numbered display copies renumbered immediately.");
    }

    private static void RunVisualTeXPastePersistenceCase(
        string artifactRoot,
        string displayMode,
        string latex,
        string pngPath,
        string emfPath,
        int renderIndex)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? sourceShape = null;
        Word.InlineShape? pastedShape = null;
        Word.Range? sourceRange = null;
        Word.Range? pastedRange = null;
        Word.Selection? selection = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            var service = new WordFormulaService(application);

            selection = application.Selection;
            selection.SetRange(0, 0);
            var sourceFormulaId = Guid.NewGuid().ToString("D");
            service.InsertOle(
                CreateVisualTeXCopySession(
                    document,
                    sourceFormulaId,
                    latex,
                    displayMode,
                    numbered: false,
                    renderIndex),
                pngPath,
                emfPath);

            AssertEqual(
                1,
                document.InlineShapes.Count,
                $"VisualTeX {displayMode} copy fixture did not create one source OLE.");

            sourceShape = document.InlineShapes[1];
            // Force a non-natural Word host size so this regression catches the
            // exact failure mode where an OLE reload falls back to the preview's
            // own natural extent instead of preserving the copied InlineShape.
            sourceShape.LockAspectRatio =
                Microsoft.Office.Core.MsoTriState.msoFalse;
            sourceShape.Width *= 1.5f;
            sourceShape.Height *= 1.5f;
            var sourceWidth = sourceShape.Width;
            var sourceHeight = sourceShape.Height;
            sourceRange = sourceShape.Range.Duplicate;
            sourceRange.Select();
            var snapshot =
                service.CaptureSelectedFormulaForCopy()
                ?? throw new InvalidDataException(
                    $"VisualTeX {displayMode} source was not captured for copy.");
            AssertTrue(
                snapshot.UsesHostCore
                && snapshot.CoreHostKind == WordFormulaHostKind.VisualTeX,
                $"VisualTeX {displayMode} source did not use the host-core copy path.");

            selection.Copy();
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            service.ArmHostCorePaste(snapshot);
            selection.Paste();

            AssertEqual(
                2,
                document.InlineShapes.Count,
                $"Word did not materialize the pasted VisualTeX {displayMode} OLE.");

            var repair = WordFormulaService.PastedFormulaRepairResult.NotReady;
            for (var attempt = 0;
                 attempt < 8
                 && repair == WordFormulaService.PastedFormulaRepairResult.NotReady;
                 attempt++)
            {
                repair = service.RepairPastedFormula(snapshot);
                if (repair == WordFormulaService.PastedFormulaRepairResult.NotReady)
                    System.Threading.Thread.Sleep(100);
            }

            AssertEqual(
                WordFormulaService.PastedFormulaRepairResult.Repaired,
                repair,
                $"VisualTeX {displayMode} pasted OLE did not complete host-core identity repair.");

            System.Threading.Thread.Sleep(1000);
            AssertEqual(
                2,
                document.InlineShapes.Count,
                $"VisualTeX {displayMode} pasted OLE disappeared after delayed repair.");

            pastedShape = document.InlineShapes[2];
            AssertNear(
                sourceWidth,
                pastedShape.Width,
                0.15f,
                $"VisualTeX {displayMode} pasted OLE changed Word host width.");
            AssertNear(
                sourceHeight,
                pastedShape.Height,
                0.15f,
                $"VisualTeX {displayMode} pasted OLE changed Word host height.");
            var pastedMetadata =
                WordFormulaMetadataReader.TryReadEmbeddedNativeOle(pastedShape)
                ?? throw new InvalidDataException(
                    $"VisualTeX {displayMode} pasted OLE has no authoritative metadata.");
            AssertTrue(
                !string.Equals(
                    sourceFormulaId,
                    pastedMetadata.FormulaId,
                    StringComparison.OrdinalIgnoreCase),
                $"VisualTeX {displayMode} pasted OLE retained the source FormulaId.");
            AssertEqual(
                latex,
                pastedMetadata.Latex,
                $"VisualTeX {displayMode} pasted OLE changed LaTeX semantics.");

            pastedRange = pastedShape.Range.Duplicate;
            pastedRange.Select();
            var opened = service.ReadSelection();
            AssertEqual(
                pastedMetadata.FormulaId,
                opened.FormulaId,
                $"VisualTeX {displayMode} pasted OLE was not immediately editable with its fresh identity.");

            var path = Path.Combine(
                artifactRoot,
                $"visualtex-{displayMode}-copy-paste.docx");
            document.SaveAs2(
                path,
                Word.WdSaveFormat.wdFormatXMLDocument);
            var persistedFormulaId = pastedMetadata.FormulaId;
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(
                path,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            AssertEqual(
                2,
                document.InlineShapes.Count,
                $"VisualTeX {displayMode} pasted OLE did not survive save/reopen.");
            Release(pastedShape);
            pastedShape = document.InlineShapes[2];
            AssertNear(
                sourceWidth,
                pastedShape.Width,
                0.15f,
                $"VisualTeX {displayMode} pasted OLE width changed after save/reopen.");
            AssertNear(
                sourceHeight,
                pastedShape.Height,
                0.15f,
                $"VisualTeX {displayMode} pasted OLE height changed after save/reopen.");
            var reopenedMetadata =
                WordFormulaMetadataReader.TryReadEmbeddedNativeOle(pastedShape)
                ?? throw new InvalidDataException(
                    $"VisualTeX {displayMode} pasted OLE lost authoritative metadata after save/reopen.");
            AssertEqual(
                persistedFormulaId,
                reopenedMetadata.FormulaId,
                $"VisualTeX {displayMode} pasted OLE lost its fresh FormulaId after save/reopen.");
        }
        finally
        {
            Release(selection);
            Release(pastedRange);
            Release(sourceRange);
            Release(pastedShape);
            Release(sourceShape);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunOmmlImmediateNumberPasteCase(string artifactRoot)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>a</mi><mo>=</mo><mi>b</mi></mrow></math>";

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? sourceRange = null;
        Word.Selection? selection = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);

            selection = application.Selection;
            selection.SetRange(0, 0);
            service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered: true,
                    FormulaOleContract.WordOmmlMode),
                mathMl);

            var source =
                WordFormulaHostResolver.CaptureDocumentIndex(document)
                    .Omml
                    .Single();
            source.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    source);
            AssertEqual(
                "1",
                WordNativeOmmlNumbering.ReadVisibleNumber(
                    document,
                    source.Numbering),
                "OMML numbered copy fixture did not start at 1.");

            sourceRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    source.Range);
            sourceRange.Select();
            var snapshot =
                service.CaptureSelectedFormulaForCopy()
                ?? throw new InvalidDataException(
                    "OMML numbered source was not captured for copy.");
            selection.Copy();
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            service.ArmHostCorePaste(snapshot);
            selection.Paste();

            var repair = WordFormulaService.PastedFormulaRepairResult.NotReady;
            for (var attempt = 0;
                 attempt < 6
                 && repair == WordFormulaService.PastedFormulaRepairResult.NotReady;
                 attempt++)
            {
                repair = service.RepairPastedFormula(snapshot);
                if (repair == WordFormulaService.PastedFormulaRepairResult.NotReady)
                    System.Threading.Thread.Sleep(80);
            }
            AssertEqual(
                WordFormulaService.PastedFormulaRepairResult.Repaired,
                repair,
                "OMML numbered paste was not repaired.");

            var hosts =
                WordFormulaHostResolver.CaptureDocumentIndex(document)
                    .Omml
                    .OrderBy(host => host.Range.Start)
                    .ToArray();
            AssertEqual(
                2,
                hosts.Length,
                "OMML numbered paste did not leave two equations.");

            foreach (var host in hosts)
                host.Numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        host);

            var sourceNumber =
                WordNativeOmmlNumbering.ReadVisibleNumber(
                    document,
                    hosts[0].Numbering);
            var pastedNumber =
                WordNativeOmmlNumbering.ReadVisibleNumber(
                    document,
                    hosts[1].Numbering);
            AssertEqual(
                "1",
                sourceNumber,
                "OMML source number changed after paste.");
            AssertEqual(
                "2",
                pastedNumber,
                "OMML pasted equation did not renumber immediately to 2.");

            var path = Path.Combine(
                artifactRoot,
                "omml-numbered-copy-paste-immediate.docx");
            document.SaveAs2(
                path,
                Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(
                path,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            var reopenedHosts =
                WordFormulaHostResolver.CaptureDocumentIndex(document)
                    .Omml
                    .OrderBy(host => host.Range.Start)
                    .ToArray();
            AssertEqual(
                2,
                reopenedHosts.Length,
                "OMML numbered copy did not survive save/reopen.");
            foreach (var host in reopenedHosts)
                host.Numbering =
                    WordFormulaNumberingResolver.ResolveLocal(
                        document,
                        host);
            AssertEqual(
                "1",
                WordNativeOmmlNumbering.ReadVisibleNumber(
                    document,
                    reopenedHosts[0].Numbering),
                "OMML source number changed after save/reopen.");
            AssertEqual(
                "2",
                WordNativeOmmlNumbering.ReadVisibleNumber(
                    document,
                    reopenedHosts[1].Numbering),
                "OMML pasted number did not persist after save/reopen.");
        }
        finally
        {
            Release(selection);
            Release(sourceRange);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunMathTypeImmediateNumberPasteCase(
        string artifactRoot,
        string emfPath)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? sourceShape = null;
        Word.Range? sourceRange = null;
        Word.Selection? selection = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            var service = new WordFormulaService(application);

            selection = application.Selection;
            selection.SetRange(0, 0);
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "right"),
                FirstNumberedMathMl,
                emfPath);

            AssertEqual(
                1,
                document.InlineShapes.Count,
                "MathType numbered copy fixture did not create one source OLE.");
            sourceShape = document.InlineShapes[1];
            sourceRange = sourceShape.Range.Duplicate;
            sourceRange.Select();

            var snapshot =
                service.CaptureSelectedFormulaForCopy()
                ?? throw new InvalidDataException(
                    "MathType numbered source was not captured for copy.");
            AssertTrue(
                snapshot.Metadata.Numbered
                && string.Equals(
                    snapshot.ObjectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal),
                "MathType numbered copy snapshot lost its native numbering state.");

            selection.Copy();
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            selection.Paste();

            var repair = WordFormulaService.PastedFormulaRepairResult.NotReady;
            for (var attempt = 0;
                 attempt < 6
                 && repair == WordFormulaService.PastedFormulaRepairResult.NotReady;
                 attempt++)
            {
                repair = service.RepairPastedFormula(snapshot);
                if (repair == WordFormulaService.PastedFormulaRepairResult.NotReady)
                    System.Threading.Thread.Sleep(80);
            }
            AssertEqual(
                WordFormulaService.PastedFormulaRepairResult.Repaired,
                repair,
                "MathType numbered paste was not repaired.");
            AssertEqual(
                2,
                document.InlineShapes.Count,
                "MathType numbered paste did not leave two OLE equations.");

            var targets =
                MathTypeEquationReferences.GetTargets(document)
                    .OrderBy(target => target.Position)
                    .ToArray();
            AssertEqual(
                2,
                targets.Length,
                "MathType numbered paste did not expose two native equation numbers.");
            AssertTrue(
                !string.Equals(
                    targets[0].NumberText,
                    targets[1].NumberText,
                    StringComparison.Ordinal),
                $"MathType pasted equation retained the source number [{targets[0].NumberText}].");

            var path = Path.Combine(
                artifactRoot,
                "mathtype-numbered-copy-paste-immediate.docx");
            document.SaveAs2(
                path,
                Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(
                path,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            var reopenedTargets =
                MathTypeEquationReferences.GetTargets(document)
                    .OrderBy(target => target.Position)
                    .ToArray();
            AssertEqual(
                2,
                reopenedTargets.Length,
                "MathType numbered copy did not survive save/reopen.");
            AssertTrue(
                !string.Equals(
                    reopenedTargets[0].NumberText,
                    reopenedTargets[1].NumberText,
                    StringComparison.Ordinal),
                "MathType pasted equation reverted to the source number after save/reopen.");
        }
        finally
        {
            Release(selection);
            Release(sourceRange);
            Release(sourceShape);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static OfficeSessionDocument CreateVisualTeXCopySession(
        Word.Document document,
        string formulaId,
        string latex,
        string displayMode,
        bool numbered,
        int renderIndex)
    {
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Host = "word",
            Mode = "create",
            FormulaId = formulaId,
            SourceDocumentId = document.FullName,
            SourceObjectId = WordRangeReference(0, 0),
            Title = "VisualTeX copy/paste regression acceptance",
            CodeFormat = "latex",
            DisplayMode = displayMode,
            ObjectMode = FormulaOleContract.NativeOleMode,
            Numbered = numbered,
            FontSizePt = 14,
            Lines = new List<FormulaLine>
            {
                new()
                {
                    Id = Guid.NewGuid().ToString("D"),
                    Latex = latex,
                },
            },
            ExportResult = new OfficeExportDocument
            {
                Width = 160 + (renderIndex % 5),
                Height = 42,
                Baseline = 31,
                FormulaLetterFont = "katex",
                FormulaChineseFont = "system",
            },
        };
    }
}
