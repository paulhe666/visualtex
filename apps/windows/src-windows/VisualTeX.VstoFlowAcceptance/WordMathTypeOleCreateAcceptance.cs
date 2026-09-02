using System.Diagnostics;
using System.Runtime.InteropServices;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeOleCreateAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "mathtype-create-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"96\" viewBox=\"0 0 240 96\"><text x=\"4\" y=\"64\" font-family=\"Times New Roman\" font-size=\"48\">x+1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 96);

        var previousCreateObjectMode = WordEquationNumbering.GetDefaultCreateObjectMode();
        try
        {
            WordEquationNumbering.SetDefaultCreateObjectMode(FormulaOleContract.MathTypeOleMode);
            AssertEqual(
                FormulaOleContract.MathTypeOleMode,
                WordEquationNumbering.GetDefaultCreateObjectMode(),
                "Word did not remember MathType OLE as the create object format.");
            WordEquationNumbering.SetDefaultCreateObjectMode(FormulaOleContract.NativeOleMode);
            AssertEqual(
                FormulaOleContract.NativeOleMode,
                WordEquationNumbering.GetDefaultCreateObjectMode(),
                "Word did not remember VisualTeX OLE as the create object format.");
        }
        finally
        {
            WordEquationNumbering.SetDefaultCreateObjectMode(previousCreateObjectMode);
        }

        RunWordMathTypeInlineCreateAcceptance(artifactRoot, emfPath);
        RunWordMathTypeRightThenLeftCreateAcceptance(artifactRoot, emfPath);
        RunWordMathTypeDisplayCreateAcceptance(artifactRoot, emfPath);
        RunWordMathTypeInsertionRollbackAcceptance(artifactRoot, emfPath);
        RunWordMathTypeCreateEditCreateLifecycleAcceptance(artifactRoot, emfPath);
        RunWordMathTypeSequentialCreateStressAcceptance(artifactRoot, emfPath);
    }

    private static void RunWordMathTypeRightThenLeftLiveAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "mathtype-right-left-live-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"96\" viewBox=\"0 0 240 96\"><text x=\"4\" y=\"64\" font-family=\"Times New Roman\" font-size=\"48\">x+1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 96);

        Word.Application? application = null;
        Word.Document? originalDocument = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        try
        {
            application = CreateWordApplication(visible: false);
            try { originalDocument = application.ActiveDocument; } catch { }
            document = application.Documents.Add();
            document.Content.Text = "VisualTeX MathType live right then left acceptance";
            var service = new WordFormulaService(application);

            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "right"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "Live right-then-left acceptance did not create the first MathType OLE.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Live right-then-left acceptance did not create the first MTPlaceRef.");
            shape = document.InlineShapes[1];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "right",
                "Live right-then-left first MathType row");
            AssertWordMathTypePreviewVisible(
                shape,
                "Live right-then-left first MathType row",
                FirstNumberedMathMl,
                inline: false,
                artifactRoot);

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"c+d",
                    mathTypeNumberPosition: "left"),
                SecondNumberedMathMl,
                emfPath);
            AssertEqual(2, document.InlineShapes.Count,
                "Live right-then-left acceptance failed to create the second MathType OLE.");
            AssertEqual(2, CountMathTypePlaceRefFields(document),
                "Live right-then-left acceptance lost or duplicated MTPlaceRef fields.");
            Release(shape);
            shape = document.InlineShapes[2];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "Live right-then-left second MathType row");
            AssertWordMathTypePreviewVisible(
                shape,
                "Live right-then-left second MathType row",
                SecondNumberedMathMl,
                inline: false,
                artifactRoot);

            var path = Path.Combine(
                artifactRoot,
                "VisualTeX-MathType-Live-Right-Then-Left.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            var embeddings = ReadDocxOleEmbeddings(path);
            AssertEqual(2, embeddings.Count,
                "Live right-then-left acceptance did not persist exactly two OLE package parts.");
            var signatures = embeddings
                .Select(MathTypeOleStorage.ReadMathMl)
                .Select(MathTypeMtefCodec.SemanticSignature)
                .ToList();
            AssertTrue(
                signatures.Contains(MathTypeMtefCodec.SemanticSignature(FirstNumberedMathMl))
                && signatures.Contains(MathTypeMtefCodec.SemanticSignature(SecondNumberedMathMl)),
                "Live right-then-left acceptance persisted the wrong MathType MTEF data.");
            Console.WriteLine(
                "[MathType LIVE] In the already-running user Word process: right-numbered Equation.DSMT4 -> left-numbered Equation.DSMT4 passed; both native previews, MTPlaceRef rows and saved CFB/MTEF package parts are valid without reading the Word clipboard.");
        }
        finally
        {
            Release(shape);
            Release(range);
            try { originalDocument?.Activate(); } catch { }
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            Release(originalDocument);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordMathTypeLeftThenRightStabilityAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "mathtype-left-right-stability-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"360\" height=\"120\" viewBox=\"0 0 360 120\"><text x=\"4\" y=\"76\" font-family=\"Times New Roman\" font-size=\"48\">MathType stability</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 360, 120);
        const string quadraticMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>";
        const string eulerMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msup><mi>e</mi><mrow><mi>i</mi><mi>π</mi></mrow></msup><mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>";

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        Word.Window? hostWindow = null;
        Process? wordProcess = null;
        try
        {
            application = CreateWordApplication(visible: true);
            document = application.Documents.Add();
            document.Content.Text = "VisualTeX MathType real-environment left then right stability acceptance";
            hostWindow = application.ActiveWindow;
            var hwnd = new IntPtr(hostWindow.Hwnd);
            _ = GetWindowThreadProcessId(hwnd, out var processId);
            if (processId == 0)
                throw new InvalidOperationException("Could not resolve the real Word PID for MathType stability acceptance.");
            wordProcess = Process.GetProcessById(unchecked((int)processId));
            Console.WriteLine(
                $"[MathType stability] Word pid={wordProcess.Id} started={wordProcess.StartTime:O} responding={wordProcess.Responding}");

            try
            {
                object addInKey = "VisualTeX.WordVsto";
                var addIn = application.COMAddIns.Item(ref addInKey);
                try
                {
                    Console.WriteLine(
                        $"[MathType stability] installed VisualTeX add-in Connect={addIn.Connect} Description={addIn.Description}");
                }
                finally { Release(addIn); }
            }
            catch (Exception error)
            {
                Console.WriteLine($"[MathType stability] installed VisualTeX add-in inventory unavailable: {error.Message}");
            }

            var service = new WordFormulaService(application);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                    mathTypeNumberPosition: "left"),
                quadraticMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "Real-environment left-right stability acceptance did not create the first MathType OLE.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Real-environment left-right stability acceptance did not create the first MTPlaceRef.");
            shape = document.InlineShapes[1];
            AssertMathTypeDisplayRow(shape, "left", "Real-environment first left-numbered MathType row");

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"e^{i\pi}+1=0",
                    mathTypeNumberPosition: "right"),
                eulerMathMl,
                emfPath);
            AssertEqual(2, document.InlineShapes.Count,
                "Real-environment left-right stability acceptance did not create the second MathType OLE.");
            AssertEqual(2, CountMathTypePlaceRefFields(document),
                "Real-environment left-right stability acceptance lost or duplicated MTPlaceRef fields.");
            Release(shape);
            shape = document.InlineShapes[2];
            AssertMathTypeDisplayRow(shape, "right", "Real-environment second right-numbered MathType row");

            Console.WriteLine(
                "[MathType stability] both inserts returned; waiting 10 seconds for delayed OLE/clipboard callbacks before declaring success...");
            Thread.Sleep(TimeSpan.FromSeconds(10));
            wordProcess.Refresh();
            AssertTrue(wordProcess.Responding,
                "Word became non-responsive after left-numbered -> right-numbered MathType insertion.");

            var cpuBefore = wordProcess.TotalProcessorTime;
            Thread.Sleep(TimeSpan.FromSeconds(5));
            wordProcess.Refresh();
            var cpuDelta = (wordProcess.TotalProcessorTime - cpuBefore).TotalMilliseconds;
            Console.WriteLine(
                $"[MathType stability] delayed window responding={wordProcess.Responding} cpuDelta5s={cpuDelta:0.0}ms");
            AssertTrue(wordProcess.Responding,
                "Word became non-responsive during the delayed MathType stability window.");
            AssertTrue(cpuDelta < 2500,
                $"Word entered a sustained CPU loop after MathType insertion (5s CPU delta={cpuDelta:0.0}ms).");

            AssertEqual(2, document.InlineShapes.Count,
                "Word changed the MathType OLE count during the delayed stability window.");
            AssertEqual(2, CountMathTypePlaceRefFields(document),
                "Word changed MTPlaceRef fields during the delayed stability window.");
            Console.WriteLine(
                "[MathType stability] REAL loaded-addin left-numbered quadratic -> right-numbered Euler insertion remained responsive and CPU-idle for 15 seconds after both inserts.");
        }
        finally
        {
            wordProcess?.Dispose();
            Release(hostWindow);
            Release(shape);
            Release(range);
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

    private static void RunWordMathTypeInlineCreateAcceptance(
        string artifactRoot,
        string emfPath)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "LEFT RIGHT";
            range = document.Range(5, 5);
            range.Select();

            var service = new WordFormulaService(application);
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "inline",
                    numbered: false,
                    latex: @"\frac{x+1}{y}"),
                FractionMathMl,
                emfPath);

            AssertEqual(1, document.InlineShapes.Count,
                "Standalone MathType inline create did not insert exactly one OLE object.");
            shape = document.InlineShapes[1];
            AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                "Standalone MathType inline create did not materialize Equation.DSMT4.");
            AssertWordMathTypePreviewVisible(
                shape,
                "Standalone MathType inline create",
                FractionMathMl,
                inline: true,
                artifactRoot);
            Console.WriteLine($"[MathType create inline probe] paragraphs={document.Paragraphs.Count}, shape={shape.Range.Start}-{shape.Range.End}");
            for (var paragraphIndex = 1; paragraphIndex <= document.Paragraphs.Count; paragraphIndex++)
            {
                var probeParagraph = document.Paragraphs[paragraphIndex];
                try
                {
                    var probeText = probeParagraph.Range.Text ?? string.Empty;
                    Console.WriteLine(
                        $"  P{paragraphIndex}={probeParagraph.Range.Start}-{probeParagraph.Range.End} cp="
                        + string.Join(",", probeText.Select(character => $"U+{(int)character:X4}")));
                }
                finally { Release(probeParagraph); }
            }
            AssertEqual(1, document.Paragraphs.Count,
                "Standalone MathType inline create unexpectedly split the text paragraph.");
            AssertTrue((document.Content.Text ?? string.Empty).Contains("LEFT")
                && (document.Content.Text ?? string.Empty).Contains("RIGHT"),
                "Standalone MathType inline create damaged surrounding prose.");
            if (MathTypeOleStorage.TryCaptureCompoundFileFromWordOpenXml(
                    shape,
                    out var immediateCompoundFile))
            {
                var immediateReadback = MathTypeOleStorage.ReadMathMl(immediateCompoundFile);
                AssertEqual(
                    MathTypeMtefCodec.SemanticSignature(FractionMathMl),
                    MathTypeMtefCodec.SemanticSignature(immediateReadback),
                    "Standalone MathType inline create changed formula semantics.");
            }
            else
            {
                Console.WriteLine(
                    "[MathType create] Live Word deferred the OLE package from Range.WordOpenXML; semantic validation is deferred to the saved DOCX package without using the clipboard.");
            }

            var path = Path.Combine(artifactRoot, "VisualTeX-MathType-Create-Inline.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            var savedEmbeddings = ReadDocxOleEmbeddings(path);
            AssertEqual(1, savedEmbeddings.Count,
                "Standalone MathType inline create did not persist exactly one OLE package part.");
            var savedReadback = MathTypeOleStorage.ReadMathMl(savedEmbeddings[0]);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(FractionMathMl),
                MathTypeMtefCodec.SemanticSignature(savedReadback),
                "Standalone MathType inline create changed in the saved DOCX package.");
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            Release(shape);
            shape = document.InlineShapes[1];
            AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                "Standalone MathType inline create lost its ProgID after Word reopen.");
            AssertWordMathTypePreviewVisible(
                shape,
                "Reopened standalone MathType inline create",
                FractionMathMl,
                inline: true,
                artifactRoot);
            Console.WriteLine("[MathType create] Inline Equation.DSMT4 insert + save/reopen passed.");
        }
        finally
        {
            Release(shape);
            Release(range);
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

    private static void RunWordMathTypeRightThenLeftCreateAcceptance(
        string artifactRoot,
        string emfPath)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "MathType right then left acceptance";
            var service = new WordFormulaService(application);

            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "right"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "Right-then-left acceptance did not create the first MathType OLE.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Right-then-left acceptance did not create the first MTPlaceRef.");
            shape = document.InlineShapes[1];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "right",
                "Right-then-left first MathType row");
            AssertWordMathTypePreviewVisible(
                shape,
                "Right-then-left first MathType row",
                FirstNumberedMathMl,
                inline: false,
                artifactRoot);

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"c+d",
                    mathTypeNumberPosition: "left"),
                SecondNumberedMathMl,
                emfPath);
            AssertEqual(2, document.InlineShapes.Count,
                "Right-then-left acceptance failed to create the second MathType OLE.");
            AssertEqual(2, CountMathTypePlaceRefFields(document),
                "Right-then-left acceptance lost or duplicated MTPlaceRef fields.");
            Release(shape);
            shape = document.InlineShapes[2];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "Right-then-left second MathType row");
            AssertWordMathTypePreviewVisible(
                shape,
                "Right-then-left second MathType row",
                SecondNumberedMathMl,
                inline: false,
                artifactRoot);

            var path = Path.Combine(
                artifactRoot,
                "VisualTeX-MathType-Right-Then-Left.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            var embeddings = ReadDocxOleEmbeddings(path);
            AssertEqual(2, embeddings.Count,
                "Right-then-left acceptance did not persist exactly two OLE package parts.");
            var signatures = embeddings
                .Select(MathTypeOleStorage.ReadMathMl)
                .Select(MathTypeMtefCodec.SemanticSignature)
                .ToList();
            AssertTrue(
                signatures.Contains(MathTypeMtefCodec.SemanticSignature(FirstNumberedMathMl))
                && signatures.Contains(MathTypeMtefCodec.SemanticSignature(SecondNumberedMathMl)),
                "Right-then-left acceptance persisted the wrong MathType MTEF data.");

            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            AssertEqual(2, document.InlineShapes.Count,
                "Right-then-left MathType OLE count changed after reopen.");
            AssertEqual(2, CountMathTypePlaceRefFields(document),
                "Right-then-left MTPlaceRef count changed after reopen.");
            Release(shape);
            shape = document.InlineShapes[1];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "right",
                "Reopened right-then-left first row");
            Release(shape);
            shape = document.InlineShapes[2];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "Reopened right-then-left second row");
            Console.WriteLine(
                "[MathType real sequence] right-numbered create -> left-numbered create passed in the same Word process; both Equation.DSMT4 CFB/MTEF packages and MTPlaceRef rows survived save + reopen.");
        }
        finally
        {
            Release(shape);
            Release(range);
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

    private static void RunWordMathTypeDisplayCreateAcceptance(
        string artifactRoot,
        string emfPath)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        Word.Selection? selection = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabs = null;
        Word.TabStop? tab = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = string.Empty;
            InsertNumberingHeading(
                application,
                document,
                level: 1,
                text: "Display create heading");
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Text = "display create acceptance";
            Release(range);
            range = null;
            // This scenario explicitly validates MathType's native heading-scope
            // state. Pin the format instead of inheriting the machine's current
            // continuous/heading default from HKCU.
            WordEquationNumbering.SetEquationNumberFormat(
                document,
                EquationNumberFormat.Heading1DotId);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            var service = new WordFormulaService(application);

            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: false,
                    latex: @"x+y"),
                SimpleMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "Standalone unnumbered MathType display create did not insert one OLE.");
            shape = document.InlineShapes[1];
            AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                "Standalone unnumbered MathType display create did not materialize Equation.DSMT4.");
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: null,
                "Standalone unnumbered MathType display create");
            AssertWordMathTypePreviewVisible(
                shape,
                "Standalone unnumbered MathType display create",
                SimpleMathMl,
                inline: false,
                artifactRoot);

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "left"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(2, document.InlineShapes.Count,
                "First numbered MathType display create did not add exactly one OLE.");
            Release(shape);
            shape = document.InlineShapes[2];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "First left-numbered MathType display create");
            AssertEqual(
                "(1.1)",
                ReadMathTypeVisibleNumberForShape(shape),
                "First left-numbered MathType display create rendered the wrong visible number.");
            AssertWordMathTypePreviewVisible(
                shape,
                "First numbered MathType display create",
                FirstNumberedMathMl,
                inline: false,
                artifactRoot);
            AssertTrue(CountMathTypePlaceRefFields(document) == 1,
                "First numbered MathType display create did not create exactly one MTPlaceRef field.");
            AssertNativeMathTypeSectionBreak(
                document,
                expectedCount: 1,
                expectedChapter: 1,
                expectedSection: 0);

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"c+d",
                    mathTypeNumberPosition: "right"),
                SecondNumberedMathMl,
                emfPath);
            AssertEqual(3, document.InlineShapes.Count,
                "Second numbered MathType display create did not add exactly one OLE.");
            Release(shape);
            shape = document.InlineShapes[3];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "right",
                "Second right-numbered MathType display create after a left-numbered row");
            AssertEqual(
                "(1.2)",
                ReadMathTypeVisibleNumberForShape(shape),
                "Second right-numbered MathType display create rendered the wrong visible number.");
            AssertWordMathTypePreviewVisible(
                shape,
                "Second numbered MathType display create",
                SecondNumberedMathMl,
                inline: false,
                artifactRoot);
            AssertTrue(CountMathTypePlaceRefFields(document) == 2,
                "Second numbered MathType display create did not preserve/clone MTPlaceRef numbering.");
            AssertNativeMathTypeSectionBreak(
                document,
                expectedCount: 1,
                expectedChapter: 1,
                expectedSection: 0);
            var codesBeforeSave = ReadMathTypePlaceRefCodes(document);
            AssertEqual(2, codesBeforeSave.Count,
                "MathType numbered create did not produce two durable MTPlaceRef codes.");
            AssertEqual(codesBeforeSave[0], codesBeforeSave[1],
                "Second numbered MathType display create did not inherit the existing MathType numbering template.");
            var path = Path.Combine(artifactRoot, "VisualTeX-MathType-Create-Display.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            document.Activate();
            AssertEqual(3, document.InlineShapes.Count,
                "MathType display creates changed object count after Word reopen.");
            AssertTrue(CountMathTypePlaceRefFields(document) == 2,
                "MathType MTPlaceRef fields did not survive Word reopen.");
            AssertNativeMathTypeSectionBreak(
                document,
                expectedCount: 1,
                expectedChapter: 1,
                expectedSection: 0);
            var codesAfterReopen = ReadMathTypePlaceRefCodes(document);
            AssertEqual(codesBeforeSave[0], codesAfterReopen[0],
                "First MathType numbering template changed after Word reopen.");
            AssertEqual(codesBeforeSave[1], codesAfterReopen[1],
                "Second MathType numbering template changed after Word reopen.");
            for (var index = 1; index <= document.InlineShapes.Count; index++)
            {
                Release(shape);
                shape = document.InlineShapes[index];
                AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                    $"MathType display create #{index} lost Equation.DSMT4 after reopen.");
                AssertMathTypeDisplayRow(
                    shape,
                    expectedNumberPosition: index == 1
                        ? null
                        : index == 2
                            ? "left"
                            : "right",
                    $"Reopened MathType display create #{index}");
                var expectedMathMl = index == 1
                    ? SimpleMathMl
                    : index == 2
                        ? FirstNumberedMathMl
                        : SecondNumberedMathMl;
                AssertWordMathTypePreviewVisible(
                    shape,
                    $"Reopened MathType display create #{index}",
                    expectedMathMl,
                    inline: false,
                    artifactRoot);
            }
            Console.WriteLine(
                "[MathType create] Unnumbered + left-then-right numbered Equation.DSMT4 inserts completed without mutating MathType's global number-side document state; native MTPlaceRef numbering, template inheritance and save/reopen passed.");
        }
        finally
        {
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraph);
            Release(selection);
            Release(shape);
            Release(range);
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

    private static void RunWordMathTypeInsertionRollbackAcceptance(
        string artifactRoot,
        string emfPath)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        var previousFailureStage = Environment.GetEnvironmentVariable(
            "VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE");
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "MathType rollback acceptance";
            var service = new WordFormulaService(application);
            var paragraphCountBefore = document.Paragraphs.Count;

            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            Environment.SetEnvironmentVariable(
                "VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE",
                "after-flat-opc");
            var failed = false;
            try
            {
                service.InsertMathTypeOle(
                    CreateMathTypeCreateSession(
                        displayMode: "block",
                        numbered: true,
                        latex: @"a+b",
                        mathTypeNumberPosition: "right"),
                    FirstNumberedMathMl,
                    emfPath);
            }
            catch (InvalidOperationException error)
            {
                failed = true;
                AssertTrue(
                    error.Message.IndexOf("insert-flat-opc", StringComparison.Ordinal) >= 0
                    && error.Message.IndexOf("8007000E", StringComparison.OrdinalIgnoreCase) >= 0,
                    "Injected MathType insertion failure did not preserve stage/HRESULT diagnostics.");
            }
            finally
            {
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE",
                    previousFailureStage);
            }
            AssertTrue(failed,
                "MathType rollback acceptance did not inject the expected failure.");
            AssertEqual(0, document.InlineShapes.Count,
                "Failed MathType create left a partial Equation.DSMT4 object behind.");
            AssertEqual(0, CountMathTypePlaceRefFields(document),
                "Failed MathType create left an orphan MTPlaceRef field behind.");
            AssertEqual(paragraphCountBefore, document.Paragraphs.Count,
                "Failed MathType create left an extra display paragraph behind.");
            AssertNativeMathTypeSectionBreak(document, 0);

            // A retry in the same Word document/process must succeed after the
            // rollback; the exception must not poison OLE clipboard state.
            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "right"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "MathType create could not recover after a rolled-back failure.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "MathType retry after rollback produced the wrong number count.");

            // Model a document saved by an older build after PasteSpecial failed:
            // remove only the OLE and leave its MTPlaceRef/tabs as an orphan row.
            shape = document.InlineShapes[1];
            var orphanStart = shape.Range.Paragraphs[1].Range.Start;
            shape.Delete();
            Release(shape);
            shape = null;
            AssertEqual(0, document.InlineShapes.Count,
                "Legacy-orphan setup still contains a MathType OLE.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Legacy-orphan setup did not leave exactly one MTPlaceRef.");
            Release(range);
            range = document.Range(orphanStart, orphanStart);
            range.Select();
            var captured = service.ReadSelection();
            var retrySession = CreateMathTypeCreateSession(
                displayMode: "block",
                numbered: true,
                latex: @"c+d",
                mathTypeNumberPosition: "left");
            retrySession.SourceDocumentId = captured.DocumentId;
            retrySession.SourceObjectId = captured.ObjectId;
            service.InsertMathTypeOle(retrySession, SecondNumberedMathMl, emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "VisualTeX did not recover an old number-only MathType failure row.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Recovering an old number-only MathType row duplicated MTPlaceRef.");
            shape = document.InlineShapes[1];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "Recovered legacy MathType orphan row");
            Console.WriteLine(
                "[MathType rollback] injected E_OUTOFMEMORY rolled back OLE + MTPlaceRef + paragraph + section state; same-process retry and legacy orphan-row recovery passed.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE",
                previousFailureStage);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shape);
            Release(range);
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

    private static void RunWordMathTypeCreateEditCreateLifecycleAcceptance(
        string artifactRoot,
        string emfPath)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "MathType lifecycle acceptance";
            var service = new WordFormulaService(application);

            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "right"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "MathType lifecycle first create did not insert one OLE.");

            Release(shape);
            shape = document.InlineShapes[1];
            shape.Range.Select();
            var firstSelection = service.ReadSelection();
            AssertTrue(firstSelection.Metadata is not null,
                "MathType lifecycle edit did not read source metadata.");
            AssertTrue(firstSelection.Metadata!.Numbered,
                "A numbered MathType display equation was read back as unnumbered.");
            AssertEqual("block", firstSelection.Metadata.DisplayMode,
                "A numbered MathType display equation was not read back as block display.");
            AssertEqual("right", service.GetMathTypeNumberPositionForRange(firstSelection.ObjectId),
                "A right-numbered MathType display equation was not read back on the right.");

            service.ReplaceMathTypeOle(
                CreateMathTypeEditSession(
                    firstSelection,
                    @"c+d",
                    mathTypeNumberPosition: "right"),
                SecondNumberedMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "MathType lifecycle edit changed the OLE count.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "MathType lifecycle edit lost or duplicated its native number.");

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: @"a+b",
                    mathTypeNumberPosition: "left"),
                FirstNumberedMathMl,
                emfPath);
            AssertEqual(2, document.InlineShapes.Count,
                "MathType lifecycle second create did not insert a second OLE.");

            Release(shape);
            shape = document.InlineShapes[2];
            shape.Range.Select();
            var secondSelection = service.ReadSelection();
            AssertTrue(secondSelection.Metadata is not null && secondSelection.Metadata.Numbered,
                "The second numbered MathType display equation was read back as unnumbered.");
            AssertEqual("left", service.GetMathTypeNumberPositionForRange(secondSelection.ObjectId),
                "A left-numbered MathType display equation was not read back on the left.");
            // Reading the second equation without replacing it models an editor
            // open/cancel cycle. No Word write is allowed on cancellation.

            Release(range);
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
            var capturedCreateSelection = service.ReadSelection();
            var thirdCreate = CreateMathTypeCreateSession(
                displayMode: "block",
                numbered: true,
                latex: @"a+b",
                mathTypeNumberPosition: "right");
            thirdCreate.SourceDocumentId = capturedCreateSelection.DocumentId;
            thirdCreate.SourceObjectId = capturedCreateSelection.ObjectId;

            // Simulate Word moving the live Selection while the external editor is
            // in front. The create must still use the captured empty paragraph and
            // the numbering template nearest to that captured position.
            Release(shape);
            shape = document.InlineShapes[1];
            shape.Range.Select();
            service.InsertMathTypeOle(thirdCreate, FirstNumberedMathMl, emfPath);
            AssertEqual(3, document.InlineShapes.Count,
                "MathType lifecycle third create failed after create/edit/cancel state transitions.");
            AssertEqual(3, CountMathTypePlaceRefFields(document),
                "MathType lifecycle third create lost or duplicated MTPlaceRef fields.");
            Release(shape);
            shape = document.InlineShapes[3];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "right",
                "MathType lifecycle third create");

            var path = Path.Combine(artifactRoot, "VisualTeX-MathType-Create-Edit-Create.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            AssertEqual(3, document.InlineShapes.Count,
                "MathType lifecycle OLE count changed after reopen.");
            AssertEqual(3, CountMathTypePlaceRefFields(document),
                "MathType lifecycle numbering count changed after reopen.");
            Console.WriteLine(
                "[MathType lifecycle] create -> numbered edit/readback -> create -> edit/cancel -> captured-position third create + save/reopen passed.");
        }
        finally
        {
            Release(shape);
            Release(range);
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

    private static OfficeSessionDocument CreateMathTypeEditSession(
        OfficeSelection source,
        string latex,
        string mathTypeNumberPosition)
    {
        var metadata = source.Metadata
            ?? throw new InvalidDataException("MathType lifecycle source metadata is unavailable.");
        var lineId = metadata.Lines.FirstOrDefault()?.Id;
        if (string.IsNullOrWhiteSpace(lineId)) lineId = Guid.NewGuid().ToString("D");
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "edit",
            Host = "word",
            FormulaId = source.FormulaId ?? metadata.FormulaId,
            SourceDocumentId = source.DocumentId,
            SourceObjectId = source.ObjectId,
            Title = "MathType lifecycle edit acceptance",
            CodeFormat = "latex",
            DisplayMode = metadata.DisplayMode,
            ObjectMode = FormulaOleContract.MathTypeOleMode,
            Numbered = metadata.Numbered,
            MathTypeNumberPosition = mathTypeNumberPosition,
            FontSizePt = metadata.FontSizePt ?? 12,
            OriginalMetadata = metadata,
            Lines = new List<FormulaLine>
            {
                new() { Id = lineId!, Latex = latex },
            },
            ExportResult = new OfficeExportDocument
            {
                Width = 240,
                Height = 96,
                Baseline = 72,
            },
        };
    }

    private static void RunWordMathTypeSequentialCreateStressAcceptance(
        string artifactRoot,
        string emfPath)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "MathType sequential create stress";
            var service = new WordFormulaService(application);
            const int targetCount = 24;
            for (var index = 1; index <= targetCount; index++)
            {
                Release(range);
                range = document.Range(document.Content.End - 1, document.Content.End - 1);
                range.Select();
                var side = index % 2 == 0 ? "right" : "left";
                service.InsertMathTypeOle(
                    CreateMathTypeCreateSession(
                        displayMode: "block",
                        numbered: true,
                        latex: @"a+b",
                        mathTypeNumberPosition: side),
                    FirstNumberedMathMl,
                    emfPath);
                AssertEqual(index, document.InlineShapes.Count,
                    $"MathType sequential create #{index} changed the OLE count unexpectedly.");
                AssertEqual(index, CountMathTypePlaceRefFields(document),
                    $"MathType sequential create #{index} lost or duplicated MTPlaceRef fields.");
                Release(shape);
                shape = document.InlineShapes[index];
                AssertMathTypeDisplayRow(
                    shape,
                    expectedNumberPosition: side,
                    $"MathType sequential create #{index}");
            }

            var path = Path.Combine(artifactRoot, "VisualTeX-MathType-Sequential-Stress.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            AssertEqual(targetCount, document.InlineShapes.Count,
                "MathType sequential create stress changed OLE count after reopen.");
            AssertEqual(targetCount, CountMathTypePlaceRefFields(document),
                "MathType sequential create stress changed MTPlaceRef count after reopen.");
            Console.WriteLine(
                $"[MathType create stress] {targetCount} consecutive numbered Equation.DSMT4 inserts with alternating left/right numbering passed in one Word process + save/reopen.");
        }
        finally
        {
            Release(shape);
            Release(range);
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

    private static void AssertWordMathTypePreviewVisible(
        Word.InlineShape shape,
        string context,
        string mathMl,
        bool inline,
        string artifactRoot)
    {
        var preview = ReadInlineShapeEnhancedMetafile(shape);
        var ink = DescribeEmfInkBounds(preview);
        Console.WriteLine($"[MathType create preview] {context}: {ink}");
        AssertTrue(!string.Equals(ink, "empty", StringComparison.Ordinal),
            context + " is a valid Equation.DSMT4 object but Word renders its OLE preview as blank.");

        var generated = MathTypeMtefCodec.CreateEquationNative(mathMl, inline);
        if (!MathTypeNativePreviewRenderer.TryRender(
                generated.Mtef,
                artifactRoot,
                out var nativePreview))
            return;
        using (nativePreview)
        {
            var expectedNativeWmf = File.ReadAllBytes(nativePreview.WmfPath);
            var difference = MeasureEmfPixelDifference(expectedNativeWmf, preview);
            Console.WriteLine(
                $"[MathType create preview] {context}: native diff={difference:0.0000}, "
                + $"size={shape.Width:0.0}x{shape.Height:0.0}pt, "
                + $"native={nativePreview.WidthPt:0.0}x{nativePreview.HeightPt:0.0}pt");
            AssertTrue(
                difference < 0.03,
                context + " is visible but does not visually match MathType's native renderer.");
            AssertNear(
                nativePreview.WidthPt,
                shape.Width,
                0.7f,
                context + " does not use MathType's native width.");
            AssertNear(
                nativePreview.HeightPt,
                shape.Height,
                0.7f,
                context + " does not use MathType's native height.");
            AssertNear(
                nativePreview.WordPosition,
                ReadInlineOlePositionForAcceptance(shape),
                1.0f,
                context + " does not use MathType's native baseline.");
        }
    }

    private static void AssertMathTypeDisplayRow(
        Word.InlineShape shape,
        string? expectedNumberPosition,
        string context)
    {
        Word.Range? shapeRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
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
            paragraphStyleObject = paragraph.Range.get_Style();
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
                context + " does not have MathType center/right tab stops.");

            var paragraphRange = paragraph.Range;
            var paragraphText = paragraphRange.Text ?? string.Empty;
            var expectNumber = expectedNumberPosition is not null;

            fields = paragraphRange.Fields;
            var sawPlaceRef = false;
            var placeRefStart = -1;
            var placeRefEnd = -1;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(fieldResult);
                fieldResult = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                sawPlaceRef = true;
                fieldResult = field.Result;
                placeRefStart = code.Start - 1;
                placeRefEnd = fieldResult.End + 1;
            }
            AssertEqual(expectNumber, sawPlaceRef,
                context + " has the wrong MathType MTPlaceRef numbering ownership.");

            if (!expectNumber)
            {
                AssertTrue(
                    paragraphText.StartsWith("\t\u0001", StringComparison.Ordinal),
                    context + " does not begin with Word's native tab + OLE sequence.");
            }
            else if (string.Equals(expectedNumberPosition, "left", StringComparison.Ordinal))
            {
                AssertTrue(placeRefEnd <= shapeRange.Start,
                    context + " does not place its MathType number before the equation.");
                separator = shapeRange.Document.Range(placeRefEnd, shapeRange.Start);
                AssertTrue((separator.Text ?? string.Empty).IndexOf('\t') >= 0,
                    context + " does not have MathType's number-to-equation center tab.");
            }
            else
            {
                AssertEqual("right", expectedNumberPosition,
                    context + " uses an unsupported MathType number position in the acceptance fixture.");
                AssertTrue(
                    paragraphText.StartsWith("\t\u0001", StringComparison.Ordinal),
                    context + " does not begin with Word's native center tab + OLE sequence.");
                AssertTrue(placeRefStart >= shapeRange.End,
                    context + " does not place its MathType number after the equation.");
                separator = shapeRange.Document.Range(shapeRange.End, placeRefStart);
                AssertTrue((separator.Text ?? string.Empty).IndexOf('\t') >= 0,
                    context + " does not have MathType's equation-to-number right tab.");
            }
            Release(paragraphRange);
        }
        finally
        {
            Release(paragraphStyle);
            paragraphStyle = null;
            paragraphStyleObject = null;
            Release(separator);
            Release(fieldResult);
            Release(code);
            Release(field);
            Release(fields);
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static string ReadMathTypeVisibleNumberForShape(Word.InlineShape shape)
    {
        Word.Range? shapeRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return string.Empty;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            fields = paragraphRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (!MathTypeEquationReferences.IsMathTypePlaceRefCode(code.Text))
                    continue;
                return MathTypeEquationReferences.ReadVisibleNumberText(field);
            }
            return string.Empty;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static void AssertNativeMathTypeReference(
        Word.Document document,
        string expectedNumberText)
    {
        Word.Fields? fields = null;
        Word.Field? outer = null;
        Word.Range? outerCode = null;
        Word.Fields? nestedFields = null;
        Word.Field? nested = null;
        Word.Range? nestedCode = null;
        Word.Range? nestedResult = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? bookmarkRange = null;
        try
        {
            fields = document.Fields;
            string? bookmarkName = null;
            var sawReference = false;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(outerCode);
                outerCode = null;
                Release(outer);
                outer = fields[index];
                outerCode = outer.Code;
                var outerText = outerCode.Text ?? string.Empty;
                var match = System.Text.RegularExpressions.Regex.Match(
                    outerText,
                    @"\bGOTOBUTTON\s+(ZEqnNum\d{6})\b",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (!match.Success) continue;

                bookmarkName = match.Groups[1].Value;
                nestedFields = outerCode.Fields;
                AssertEqual(1, nestedFields.Count,
                    "Native MathType GOTOBUTTON reference must contain exactly one nested REF field.");
                nested = nestedFields[1];
                nestedCode = nested.Code;
                var normalizedNested = NormalizeFieldCodeForMathTypeAcceptance(
                    nestedCode.Text ?? string.Empty);
                AssertTrue(
                    normalizedNested.IndexOf(
                        "REF " + bookmarkName,
                        StringComparison.OrdinalIgnoreCase) >= 0,
                    "Native MathType reference does not target the same ZEqnNum bookmark as its GOTOBUTTON field.");
                AssertTrue(
                    normalizedNested.IndexOf("\\!", StringComparison.Ordinal) >= 0,
                    "Native MathType REF field is missing the \\! non-recursive-update switch.");
                AssertTrue(
                    normalizedNested.IndexOf("\\* Charformat", StringComparison.OrdinalIgnoreCase) >= 0,
                    "Native MathType REF field is missing MathType's Charformat switch.");
                nestedResult = nested.Result;
                AssertEqual(
                    expectedNumberText.Trim(),
                    (nestedResult.Text ?? string.Empty).Trim(),
                    "Native MathType REF field shows the wrong equation number.");
                sawReference = true;
                break;
            }

            AssertTrue(sawReference && !string.IsNullOrWhiteSpace(bookmarkName),
                "Document does not contain a native MathType GOTOBUTTON + nested REF equation reference.");
            bookmarks = document.Bookmarks;
            AssertTrue(bookmarks.Exists(bookmarkName!),
                "Native MathType reference target bookmark does not exist.");
            bookmark = bookmarks[bookmarkName!];
            bookmarkRange = bookmark.Range;
            AssertEqual(
                expectedNumberText.Trim(),
                (bookmarkRange.Text ?? string.Empty).Trim(),
                "Native MathType ZEqnNum bookmark does not cover the visible equation number.");
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(nestedResult);
            Release(nestedCode);
            Release(nested);
            Release(nestedFields);
            Release(outerCode);
            Release(outer);
            Release(fields);
        }
    }

    private static void AssertNativeMathTypeSectionBreak(
        Word.Document document,
        int expectedCount,
        int expectedChapter = 1,
        int expectedSection = 1)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Fields? nestedFields = null;
        Word.Field? nestedField = null;
        Word.Range? nestedCode = null;
        object? styleObject = null;
        Word.Style? style = null;
        var breakCount = 0;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var codeText = code.Text ?? string.Empty;
                if (codeText.IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                breakCount++;
                var normalizedOuter = NormalizeFieldCodeForMathTypeAcceptance(codeText);
                AssertTrue(normalizedOuter.IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0,
                    "MathType section break lost its MTEditEquationSection2 MacroButton.");

                nestedFields = code.Fields;
                AssertEqual(3, nestedFields.Count,
                    "MathType default section break does not contain exactly three nested SEQ fields.");
                var nestedCodes = new List<(int Start, string Code)>();
                for (var nestedIndex = 1; nestedIndex <= nestedFields.Count; nestedIndex++)
                {
                    Release(nestedCode);
                    nestedCode = null;
                    Release(nestedField);
                    nestedField = nestedFields[nestedIndex];
                    nestedCode = nestedField.Code;
                    nestedCodes.Add((
                        nestedCode.Start,
                        NormalizeFieldCodeForMathTypeAcceptance(nestedCode.Text ?? string.Empty)));
                }
                var ordered = nestedCodes.OrderBy(item => item.Start).Select(item => item.Code).ToArray();
                AssertEqual("SEQ MTEqn \\r \\h \\* MERGEFORMAT", ordered[0],
                    "MathType section break does not reset MTEqn using MathType's native field code.");
                AssertEqual(
                    $"SEQ MTSec \\r {expectedSection} \\h \\* MERGEFORMAT",
                    ordered[1],
                    $"MathType section break does not initialize MTSec to {expectedSection} using MathType's native field code.");
                AssertEqual(
                    $"SEQ MTChap \\r {expectedChapter} \\h \\* MERGEFORMAT",
                    ordered[2],
                    $"MathType section break does not initialize MTChap to {expectedChapter} using MathType's native field code.");

                styleObject = code.get_Style();
                style = styleObject as Word.Style;
                AssertTrue(style is not null,
                    "MathType section break does not expose its character style.");
                AssertEqual("MTEquationSection", style!.NameLocal,
                    "MathType section break does not use MTEquationSection.");
                AssertEqual(-1, style.Font.Hidden,
                    "MTEquationSection is not hidden like native MathType.");
                AssertEqual((int)Word.WdColor.wdColorRed, (int)style.Font.Color,
                    "MTEquationSection is not red like native MathType.");

                Release(style);
                style = null;
                styleObject = null;
                Release(nestedCode);
                nestedCode = null;
                Release(nestedField);
                nestedField = null;
                Release(nestedFields);
                nestedFields = null;
            }
            AssertEqual(expectedCount, breakCount,
                "MathType create inserted the wrong number of chapter/section breaks.");
            if (expectedCount > 0)
            {
                var firstPlaceRefStart = FindFirstMathTypePlaceRefStartForAcceptance(document);
                var firstSectionBreakStart = FindFirstMathTypeSectionBreakStartForAcceptance(document);
                AssertTrue(firstPlaceRefStart >= 0 && firstSectionBreakStart >= 0,
                    "MathType create could not resolve the native section break / MTPlaceRef ordering.");
                AssertTrue(firstSectionBreakStart < firstPlaceRefStart,
                    "The default MathType chapter/section break must precede the first numbered equation.");
            }
        }
        finally
        {
            Release(style);
            styleObject = null;
            Release(nestedCode);
            Release(nestedField);
            Release(nestedFields);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static int FindFirstMathTypePlaceRefStartForAcceptance(Word.Document document)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        var best = int.MaxValue;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                best = Math.Min(best, Math.Max(document.Content.Start, code.Start - 1));
            }
            return best == int.MaxValue ? -1 : best;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static int FindFirstMathTypeSectionBreakStartForAcceptance(Word.Document document)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        var best = int.MaxValue;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                best = Math.Min(best, Math.Max(document.Content.Start, code.Start - 1));
            }
            return best == int.MaxValue ? -1 : best;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static int CountMathTypePlaceRefFields(Word.Document document) =>
        ReadMathTypePlaceRefCodes(document).Count;

    private static List<string> ReadMathTypePlaceRefCodes(Word.Document document)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        var result = new List<(int Start, string Code)>();
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var text = code.Text ?? string.Empty;
                if (text.IndexOf("MACROBUTTON MTPlaceRef", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                result.Add((code.Start, NormalizeFieldCodeForMathTypeAcceptance(text)));
            }
            return result.OrderBy(item => item.Start).Select(item => item.Code).ToList();
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static string NormalizeFieldCodeForMathTypeAcceptance(string value) =>
        string.Join(
            " ",
            (value ?? string.Empty)
                .Replace("\u0013", " ")
                .Replace("\u0014", " ")
                .Replace("\u0015", " ")
                .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private static OfficeSessionDocument CreateMathTypeCreateSession(
        string displayMode,
        bool numbered,
        string latex,
        string mathTypeNumberPosition = "right") =>
        new()
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = Guid.NewGuid().ToString("D"),
            SourceDocumentId = null,
            SourceObjectId = null,
            Title = "MathType standalone create acceptance",
            CodeFormat = "latex",
            DisplayMode = displayMode,
            ObjectMode = FormulaOleContract.MathTypeOleMode,
            Numbered = numbered,
            MathTypeNumberPosition = mathTypeNumberPosition,
            FontSizePt = 12,
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
            ExportResult = new OfficeExportDocument
            {
                Width = 240,
                Height = 96,
                Baseline = 72,
            },
        };

    private const string FractionMathMl =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow><mi>y</mi></mfrac></math>";
    private const string SimpleMathMl =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mi>y</mi></math>";
    private const string FirstNumberedMathMl =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>a</mi><mo>+</mo><mi>b</mi></math>";
    private const string SecondNumberedMathMl =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>c</mi><mo>+</mo><mi>d</mi></math>";
}
