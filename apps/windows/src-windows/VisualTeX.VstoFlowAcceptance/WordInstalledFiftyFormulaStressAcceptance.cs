using System.Diagnostics;
using System.Globalization;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class StressFormulaCase
    {
        internal string Name { get; set; } = string.Empty;
        internal string Latex { get; set; } = string.Empty;
        internal string DisplayMode { get; set; } = string.Empty;
        internal bool Numbered { get; set; }
    }

    private sealed class PreparedStressFormula
    {
        internal StressFormulaCase Formula { get; set; } = new();
        internal OfficeExportDocument Export { get; set; } = new();
        internal string Signature { get; set; } = string.Empty;
        internal string CanonicalLatex { get; set; } = string.Empty;
    }

    private sealed class StressEditMetric
    {
        internal int Index { get; set; }
        internal string Name { get; set; } = string.Empty;
        internal bool Numbered { get; set; }
        internal string DisplayMode { get; set; } = string.Empty;
        internal double OpenMilliseconds { get; set; }
        internal double CommitMilliseconds { get; set; }
    }

    private static IReadOnlyList<StressFormulaCase> CreateFiftyFormulaStressCorpus()
    {
        var latex = new (string Name, string Latex)[]
        {
            ("basic-arithmetic", @"a+b=c"),
            ("fraction", @"\frac{a+b}{c-d}"),
            ("square-root", @"\sqrt{x^{2}+y^{2}}"),
            ("nth-root", @"\sqrt[3]{x^{2}+1}"),
            ("subsup", @"x_{i}^{2}+y_{jk}^{3}"),
            ("sum", @"\sum_{i=1}^{n} i^{2}"),
            ("product", @"\prod_{k=1}^{n} x_{k}"),
            ("integral", @"\int_{0}^{1}x^{2}\,\mathrm{d}x"),
            ("double-integral", @"\iint_{D}f(x,y)\,\mathrm{d}A"),
            ("triple-integral", @"\iiint_{V}\rho\,\mathrm{d}V"),
            ("contour-integral", @"\oint_{C}f(z)\,\mathrm{d}z"),
            ("limit", @"\lim_{x\to0}\frac{\sin x}{x}"),
            ("maximum", @"\max_{x\in A}f(x)"),
            ("minimum", @"\min_{x\in A}f(x)"),
            ("supremum", @"\sup_{x\in A}f(x)"),
            ("infimum", @"\inf_{x\in A}f(x)"),
            ("trig", @"\sin^{2}\theta+\cos^{2}\theta=1"),
            ("log-exp", @"\log x+\ln y+\exp(z)"),
            ("euler", @"\mathrm{e}^{\mathrm{i}\pi}+1=0"),
            ("lower-greek", @"\alpha+\beta+\gamma+\delta+\theta+\lambda+\mu+\sigma+\omega"),
            ("upper-greek", @"\Gamma+\Delta+\Theta+\Lambda+\Sigma+\Omega"),
            ("greek-variants", @"\epsilon+\varepsilon+\phi+\varphi"),
            ("physics-symbols", @"\hbar\omega+\partial_{t}\psi+\nabla^{2}\psi"),
            ("arithmetic-symbols", @"a\pm b\mp c\times d\div e\cdot f"),
            ("relations", @"a\le b\ge c\ne d\approx e\equiv f\propto g"),
            ("sets", @"x\in A\notin B\subset C\subseteq D\supset E\supseteq F"),
            ("set-ops", @"(A\cup B)\cap C\setminus D=\varnothing"),
            ("logic", @"\forall x\,\exists y:\neg P(x)\lor Q(y)\land R(x,y)"),
            ("arrows", @"A\to B\leftarrow C\leftrightarrow D\Rightarrow E\Leftrightarrow F"),
            ("accents", @"\vec{v}+\hat{x}+\tilde{y}+\dot{q}+\ddot{q}"),
            ("lines", @"\overline{AB}+\underline{x}+\bar{z}"),
            ("braces", @"\overbrace{a+b+c}^{n}+\underbrace{x+y}_{m}"),
            ("pmatrix", @"\begin{pmatrix}a&b\\c&d\end{pmatrix}"),
            ("bmatrix", @"\begin{bmatrix}1&2\\3&4\end{bmatrix}"),
            ("determinant", @"\left|\begin{matrix}a&b\\c&d\end{matrix}\right|"),
            ("cases", @"f(x)=\begin{cases}x^{2},&x\ge0\\-x,&x<0\end{cases}"),
            ("binomial", @"\binom{n}{k}=\frac{n!}{k!(n-k)!}"),
            ("explicit-column-matrix", @"(a+b)^{n}=\sum_{k=0}^{n}\left(\begin{matrix}n\\k\end{matrix}\right)a^{n-k}b^{k}"),
            ("matrix-3x3", @"A=\begin{matrix}a_{11}&a_{12}&a_{13}\\a_{21}&a_{22}&a_{23}\\a_{31}&a_{32}&a_{33}\end{matrix}"),
            ("text", @"f(x)=x^{2}\quad\text{if }x>0"),
            ("blackboard", @"\mathbb{R}\subset\mathbb{C},\quad\mathbb{Z}\subset\mathbb{Q}"),
            ("calligraphic", @"\mathcal{F}\{f\}(\omega)=\int_{-\infty}^{\infty}f(t)\mathrm{e}^{-\mathrm{i}\omega t}\,\mathrm{d}t"),
            ("fraktur", @"\mathfrak{g}=\mathfrak{h}\oplus\mathfrak{m}"),
            ("braket", @"\langle\psi|H|\psi\rangle=E"),
            ("delimiters", @"|x|+\|v\|+\lfloor x\rfloor+\lceil y\rceil"),
            ("cancel", @"\cancel{x}+\bcancel{y}+\xcancel{z}"),
            ("partial-derivative", @"\frac{\partial^{2}u}{\partial x^{2}}+\frac{\partial^{2}u}{\partial y^{2}}=0"),
            ("tensor", @"T_{ij}^{\;kl}=A_{i}^{k}B_{j}^{l}"),
            ("primes", @"f'(x)+g''(x)+p_{2}^{\prime\prime}(x)"),
            ("wide-arrows", @"\overrightarrow{AB}+\overleftarrow{CD}+\overleftrightarrow{EF}"),
        };

        if (latex.Length != 50)
            throw new InvalidOperationException($"Stress corpus must contain 50 formulas, actual={latex.Length}.");

        var result = new List<StressFormulaCase>(50);
        for (var index = 0; index < latex.Length; index++)
        {
            var display = index < 45;
            var numbered = display && (index < 30 || index >= 40);
            result.Add(new StressFormulaCase
            {
                Name = latex[index].Name,
                Latex = latex[index].Latex,
                DisplayMode = display ? "block" : "inline",
                Numbered = numbered,
            });
        }
        AssertEqual(35, result.Count(item => item.Numbered),
            "Stress corpus must contain 35 numbered display formulas.");
        AssertEqual(5, result.Count(item => item.DisplayMode == "inline"),
            "Stress corpus must contain five inline formulas.");
        return result;
    }

    private static IReadOnlyList<PreparedStressFormula> PrepareFiftyStressCorpus(
        VisualTeXSessionClient client)
    {
        var result = new List<PreparedStressFormula>(50);
        var formulas = CreateFiftyFormulaStressCorpus();
        for (var index = 0; index < formulas.Count; index++)
        {
            var formula = formulas[index];
            var export = CreateInstalledMathTypeProductExport(
                client,
                formula.Latex,
                FormulaOleContract.MathTypeOleMode,
                formula.DisplayMode,
                formula.Numbered);
            if (string.IsNullOrWhiteSpace(export.MathMl))
                throw new InvalidDataException(
                    $"Stress corpus #{index + 1} '{formula.Name}' produced no MathML.");
            var signature = MathTypeMtefCodec.SemanticSignature(export.MathMl!);
            var standalone = MathTypeMtefCodec.CreateEquationNative(
                export.MathMl!,
                formula.DisplayMode == "inline");
            var standaloneCompound = MathTypeOleStorage.CreateStandaloneCompoundFile(standalone);
            var standaloneMathMl = MathTypeOleStorage.ReadMathMl(standaloneCompound);
            var standaloneSignature = MathTypeMtefCodec.SemanticSignature(standaloneMathMl);
            if (!string.Equals(signature, standaloneSignature, StringComparison.Ordinal))
                throw new InvalidDataException(
                    $"Stress corpus #{index + 1} '{formula.Name}' fails standalone MTEF preflight. expected='{signature}' actual='{standaloneSignature}' sourceMathMl='{export.MathMl}' actualMathMl='{standaloneMathMl}'.");
            var canonicalLatex = NormalizeStressLatex(
                MathMlToLatexConverter.Convert(export.MathMl!));
            result.Add(new PreparedStressFormula
            {
                Formula = formula,
                Export = export,
                Signature = signature,
                CanonicalLatex = canonicalLatex,
            });
            Console.WriteLine(
                $"[STRESS CORPUS {index + 1:D2}/50] {formula.Name}: mode={formula.DisplayMode}, numbered={formula.Numbered}, latex={formula.Latex}");
        }
        return result;
    }

    private static string NormalizeStressLatex(string value) =>
        new string((value ?? string.Empty).Where(character => !char.IsWhiteSpace(character)).ToArray())
            .Replace(@"\left.", string.Empty)
            .Replace(@"\right.", string.Empty)
            .Replace(@"\left", string.Empty)
            .Replace(@"\right", string.Empty)
            .Replace(@"\qquad", string.Empty)
            .Replace(@"\quad", string.Empty)
            .Replace(@"\,", string.Empty)
            .Replace(@"\!", string.Empty)
            .Replace(@"\:", string.Empty)
            .Replace(@"\;", string.Empty);

    private static void RunWordInstalledMathTypeFiftyEditStressAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var previousAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        var previousFormatAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousTracePath = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        var previousCreateMode = WordEquationNumbering.GetDefaultCreateObjectMode();
        var previousNumbered = WordEquationNumbering.GetDefaultDisplayEquationNumbered();
        var previousNumberPosition = WordEquationNumbering.GetDefaultMathTypeNumberPosition();
        var tracePath = Path.Combine(artifactRoot, "mathtype-50-edit-stress.trace.log");
        Word.Application? application = null;
        Word.Document? document = null;
        Microsoft.Office.Core.COMAddIns? addIns = null;
        Microsoft.Office.Core.COMAddIn? installedAddIn = null;
        object? callbacksObject = null;
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
            var mathTypeBaseline = SnapshotMathTypeProcessIds();
            if (mathTypeBaseline.Count != 0)
                throw new InvalidOperationException("MathType.exe must not be running before 50-formula stress acceptance.");

            var prepared = PrepareFiftyStressCorpus(client);
            application = CreateWordApplication(visible: true);
            callbacksObject = GetInstalledStressCallbacks(application, out addIns, out installedAddIn);
            dynamic callbacks = callbacksObject;
            WordEquationNumbering.SetDefaultCreateObjectMode(FormulaOleContract.MathTypeOleMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(true);
            WordEquationNumbering.SetDefaultMathTypeNumberPosition("right");

            document = application.Documents.Add();
            document.Activate();
            var creationWatch = Stopwatch.StartNew();
            for (var index = 0; index < prepared.Count; index++)
            {
                EnsureStressIndependentFormulaHost(document, index);
                var item = prepared[index];
                if (item.Formula.DisplayMode == "inline")
                {
                    CommitInstalledMathTypeInlineFromRibbon(
                        client,
                        callbacks,
                        document,
                        item.Formula.Latex,
                        item.Export,
                        index + 1,
                        mathTypeBaseline);
                }
                else
                {
                    CommitInstalledMathTypeDisplayFromRibbon(
                        client,
                        callbacks,
                        document,
                        item.Formula.Latex,
                        item.Export,
                        "right",
                        index + 1,
                        mathTypeBaseline,
                        numbered: item.Formula.Numbered);
                }
                if ((index + 1) % 10 == 0)
                    Console.WriteLine($"[MT50 CREATE] {index + 1}/50 elapsed={creationWatch.Elapsed.TotalSeconds:0.00}s");
            }
            creationWatch.Stop();
            AssertEqual(50, CountMathTypeOleShapes(document), "MT50 creation produced the wrong object count.");
            AssertEqual(35, CountStressMathTypePlaceRefFields(document), "MT50 creation produced the wrong numbered-equation count.");
            AssertStressMathTypeInventory(document, prepared, "after creating 50 MathType formulas");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "creating 50 MathType formulas");

            var fixturePath = Path.Combine(artifactRoot, "Stress-MathType-50.docx");
            document.SaveAs2(fixturePath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(fixturePath, ReadOnly: false, AddToRecentFiles: false, Visible: true);
            document.Activate();
            AssertStressMathTypeInventory(document, prepared, "reopened before MT50 edit stress");

            var metrics = new List<StressEditMetric>(50);
            for (var index = 0; index < prepared.Count; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = document.InlineShapes[index + 1];
                    shape.Range.Select();
                    var existing = SnapshotSessionIds();
                    var openWatch = Stopwatch.StartNew();
                    callbacks.OnEditSelected(null);
                    var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(20));
                    var session = WaitForStressEditorOpened(client, sessionId, TimeSpan.FromSeconds(20));
                    openWatch.Stop();
                    AssertEqual(FormulaOleContract.MathTypeOleMode, session.ObjectMode,
                        $"MT50 edit #{index + 1} opened the wrong object mode.");
                    AssertEqual(prepared[index].Formula.Numbered, session.Numbered,
                        $"MT50 edit #{index + 1} changed numbered state before commit.");
                    var importedLatex = string.Join("\n", session.Lines.Select(line => line.Latex)).Trim();
                    if (string.IsNullOrWhiteSpace(importedLatex))
                        throw new InvalidDataException($"MT50 edit #{index + 1} imported empty LaTeX.");

                    // Validate the actual source recovered from Equation Native.
                    // Equivalent TeX spellings such as \lim_{x\to0} and
                    // \underset{x\to0}{\lim} must be compared by MathML/MTEF
                    // semantics, not by raw LaTeX spelling. Re-render the imported
                    // source through the production converter, then commit exactly
                    // that imported source so this stress test cannot hide a bad
                    // MathType→VisualTeX read by writing the original fixture back.
                    var importedExport = CreateInstalledMathTypeProductExport(
                        client,
                        importedLatex,
                        FormulaOleContract.MathTypeOleMode,
                        prepared[index].Formula.DisplayMode,
                        prepared[index].Formula.Numbered);
                    if (string.IsNullOrWhiteSpace(importedExport.MathMl))
                        throw new InvalidDataException($"MT50 edit #{index + 1} imported source produced no MathML.");
                    var importedSignature = MathTypeMtefCodec.SemanticSignature(importedExport.MathMl!);
                    AssertEqual(prepared[index].Signature, importedSignature,
                        $"MT50 edit #{index + 1} imported different formula semantics. imported='{importedLatex}'.");
                    var importedItem = new PreparedStressFormula
                    {
                        Formula = new StressFormulaCase
                        {
                            Name = prepared[index].Formula.Name,
                            Latex = importedLatex,
                            DisplayMode = prepared[index].Formula.DisplayMode,
                            Numbered = prepared[index].Formula.Numbered,
                        },
                        Export = importedExport,
                        Signature = importedSignature,
                        CanonicalLatex = NormalizeStressLatex(importedLatex),
                    };

                    var commitWatch = Stopwatch.StartNew();
                    PatchStressSessionForCommit(
                        client,
                        session,
                        importedItem,
                        FormulaOleContract.MathTypeOleMode,
                        dirty: true);
                    WaitForStressSessionCompletion(client, sessionId, TimeSpan.FromSeconds(30));
                    WaitForInstalledRibbonSessionRelease(sessionId);
                    commitWatch.Stop();
                    AssertNoNewMathTypeProcess(mathTypeBaseline, $"MT50 edit #{index + 1}");
                    AssertEqual(50, CountMathTypeOleShapes(document),
                        $"MT50 edit #{index + 1} changed the MathType object count.");
                    Release(shape); shape = document.InlineShapes[index + 1];
                    var signature = MathTypeMtefCodec.SemanticSignature(MathTypeOleStorage.ReadMathMl(shape));
                    AssertEqual(prepared[index].Signature, signature,
                        $"MT50 edit #{index + 1} changed Equation Native semantics.");
                    metrics.Add(new StressEditMetric
                    {
                        Index = index + 1,
                        Name = prepared[index].Formula.Name,
                        Numbered = prepared[index].Formula.Numbered,
                        DisplayMode = prepared[index].Formula.DisplayMode,
                        OpenMilliseconds = openWatch.Elapsed.TotalMilliseconds,
                        CommitMilliseconds = commitWatch.Elapsed.TotalMilliseconds,
                    });
                    Console.WriteLine(
                        $"[MT50 EDIT {index + 1:D2}/50] numbered={prepared[index].Formula.Numbered} mode={prepared[index].Formula.DisplayMode} open={openWatch.Elapsed.TotalMilliseconds:0}ms commit={commitWatch.Elapsed.TotalMilliseconds:0}ms");
                }
                finally { Release(shape); }
            }

            AssertEqual(35, CountStressMathTypePlaceRefFields(document),
                "Editing all 50 MathType formulas changed the numbered-equation inventory.");
            AssertStressEditPerformance(metrics);
            WriteStressEditCsv(Path.Combine(artifactRoot, "mathtype-50-edit-times.csv"), metrics);

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(fixturePath, ReadOnly: false, AddToRecentFiles: false, Visible: true);
            AssertStressMathTypeInventory(document, prepared, "after MT50 edit stress save/reopen");
            AssertEqual(35, CountStressMathTypePlaceRefFields(document),
                "Save/reopen after MT50 edit stress changed numbered-equation inventory.");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "MT50 edit stress save/reopen");

            Console.WriteLine(
                $"[MT50 STRESS PASS] created=50, numbered=35, unnumberedDisplay=10, inline=5, createTotal={creationWatch.Elapsed.TotalSeconds:0.00}s; every formula opened through the installed Ribbon edit callback and recommitted; save/reopen semantics remained exact; MathTypeProcessCount=0; fixture={fixturePath}");
        }
        finally
        {
            WordEquationNumbering.SetDefaultCreateObjectMode(previousCreateMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(previousNumbered);
            WordEquationNumbering.SetDefaultMathTypeNumberPosition(previousNumberPosition);
            Release(callbacksObject);
            Release(installedAddIn);
            Release(addIns);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTracePath);
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", previousFormatAcceptance);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", previousAcceptance);
        }
    }

    private static void RunWordInstalledFormatFiftyBatchStressAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var previousAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        var previousFormatAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousTracePath = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        var previousCreateMode = WordEquationNumbering.GetDefaultCreateObjectMode();
        var previousNumbered = WordEquationNumbering.GetDefaultDisplayEquationNumbered();
        var tracePath = Path.Combine(artifactRoot, "format-50-batch-stress.trace.log");
        Word.Application? application = null;
        Word.Document? document = null;
        Microsoft.Office.Core.COMAddIns? addIns = null;
        Microsoft.Office.Core.COMAddIn? installedAddIn = null;
        object? callbacksObject = null;
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
            var mathTypeBaseline = SnapshotMathTypeProcessIds();
            if (mathTypeBaseline.Count != 0)
                throw new InvalidOperationException("MathType.exe must not be running before 50-formula conversion stress.");
            var prepared = PrepareFiftyStressCorpus(client);
            application = CreateWordApplication(visible: true);
            callbacksObject = GetInstalledStressCallbacks(application, out addIns, out installedAddIn);
            dynamic callbacks = callbacksObject;
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(true);

            // Diagnostic fast path for the expensive fourth direction. It still
            // uses the installed Word Ribbon callback, all 50 real MathType OLEs,
            // strict semantic/numbering inventory checks, and save/reopen; it only
            // avoids rebuilding the same MathType fixture through the first three
            // directions while iterating on MT→OMML performance.
            var reusableMathTypeFixture = Environment.GetEnvironmentVariable(
                "VISUALTEX_STRESS_REUSE_MT_FIXTURE");
            if (!string.IsNullOrWhiteSpace(reusableMathTypeFixture)
                && File.Exists(reusableMathTypeFixture))
            {
                var diagnosticPath = Path.Combine(artifactRoot, "Stress-MT-OMML-50.docx");
                File.Copy(reusableMathTypeFixture, diagnosticPath, true);
                document = application.Documents.Open(
                    diagnosticPath,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: true);
                AssertStressMathTypeInventory(document, prepared, "reused MathType-50 source");
                AssertEqual(35, CountStressMathTypePlaceRefFields(document),
                    "Reused MathType-50 fixture has the wrong numbered-equation count.");
                var mtToOmmlOnly = TimeStressDocumentConversion(
                    () => callbacks.OnConvertMathTypeToOmmlDocument(null),
                    document,
                    tracePath,
                    "source=MathType target=OMML",
                    mathTypeBaseline);
                AssertStressOmmlInventory(document, prepared, "diagnostic MT→OMML 50 batch");
                document.Save();
                document.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(document); document = null;
                document = application.Documents.Open(
                    diagnosticPath,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: true);
                AssertStressOmmlInventory(document, prepared, "reopened diagnostic MT→OMML 50 batch");
                AssertNoNewMathTypeProcess(mathTypeBaseline, "diagnostic MT→OMML 50 batch");
                Console.WriteLine(
                    $"[FORMAT50 MT→OMML DIAGNOSTIC PASS] elapsed={mtToOmmlOnly.TotalSeconds:0.00}s; all 50 semantic signatures and 35 numbered formulas survived save/reopen; MathTypeProcessCount=0; source={reusableMathTypeFixture}");
                return;
            }

            var vtFixture = Path.Combine(artifactRoot, "Stress-VisualTeX-50.docx");
            var reusableVtFixture = Environment.GetEnvironmentVariable("VISUALTEX_STRESS_REUSE_VT_FIXTURE");
            if (!string.IsNullOrWhiteSpace(reusableVtFixture) && File.Exists(reusableVtFixture))
            {
                File.Copy(reusableVtFixture, vtFixture, true);
                Console.WriteLine($"[STRESS DEBUG] Reused real VisualTeX-50 fixture: {reusableVtFixture}");
            }
            else
            {
                document = CreateStressSourceDocument(
                    client,
                    application,
                    callbacks,
                    prepared,
                    FormulaOleContract.NativeOleMode,
                    mathTypeBaseline);
                AssertStressVisualTeXInventory(document, prepared, expectCanonical: false, "VisualTeX-50 source");
                document.SaveAs2(vtFixture, Word.WdSaveFormat.wdFormatXMLDocument);
                document.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(document); document = null;
            }

            var ommlFixture = Path.Combine(artifactRoot, "Stress-OMML-50.docx");
            var reusableOmmlFixture = Environment.GetEnvironmentVariable("VISUALTEX_STRESS_REUSE_OMML_FIXTURE");
            if (!string.IsNullOrWhiteSpace(reusableOmmlFixture) && File.Exists(reusableOmmlFixture))
            {
                File.Copy(reusableOmmlFixture, ommlFixture, true);
                document = application.Documents.Open(ommlFixture, ReadOnly: false, AddToRecentFiles: false, Visible: true);
                AssertStressOmmlInventory(document, prepared, "reused OMML-50 source");
                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                Release(document); document = null;
                Console.WriteLine($"[STRESS DEBUG] Reused and revalidated real OMML-50 fixture: {reusableOmmlFixture}");
            }
            else
            {
                document = CreateStressSourceDocument(
                    client,
                    application,
                    callbacks,
                    prepared,
                    FormulaOleContract.WordOmmlMode,
                    mathTypeBaseline);
                // Persist the real Word-materialized OMML fixture before semantic
                // assertions so a failing stress case leaves the exact Office XML
                // available for root-cause inspection.
                document.SaveAs2(ommlFixture, Word.WdSaveFormat.wdFormatXMLDocument);
                AssertStressOmmlInventory(document, prepared, "OMML-50 source");
                document.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(document); document = null;
            }

            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_STRESS_FORMAT_PHASE"),
                    "omml-only",
                    StringComparison.OrdinalIgnoreCase))
            {
                var ommlOnlyRoundTrip = Path.Combine(
                    artifactRoot,
                    "Stress-OMML-MT-OMML-50.docx");
                File.Copy(ommlFixture, ommlOnlyRoundTrip, true);
                document = application.Documents.Open(
                    ommlOnlyRoundTrip,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: true);
                var ommlOnlyToMt = TimeStressDocumentConversion(
                    () => callbacks.OnConvertOmmlToMathTypeDocument(null),
                    document,
                    tracePath,
                    "source=OMML target=MathType",
                    mathTypeBaseline);
                AssertStressMathTypeInventory(
                    document,
                    prepared,
                    "OMML→MT 50 batch (OMML-only phase)");
                AssertEqual(
                    35,
                    CountStressMathTypePlaceRefFields(document),
                    "OMML→MT lost numbered equations in OMML-only phase.");
                var mtOnlyToOmml = TimeStressDocumentConversion(
                    () => callbacks.OnConvertMathTypeToOmmlDocument(null),
                    document,
                    tracePath,
                    "source=MathType target=OMML",
                    mathTypeBaseline);
                AssertStressOmmlInventory(
                    document,
                    prepared,
                    "OMML→MT→OMML 50 batch (OMML-only phase)");
                document.Save();
                document.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(document); document = null;
                document = application.Documents.Open(
                    ommlOnlyRoundTrip,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: true);
                AssertStressOmmlInventory(
                    document,
                    prepared,
                    "reopened OMML→MT→OMML 50 batch (OMML-only phase)");
                AssertNoNewMathTypeProcess(
                    mathTypeBaseline,
                    "50-formula OMML-only batch conversion stress");
                Console.WriteLine(
                    $"[FORMAT50 OMML-ONLY PASS] OMML→MT={ommlOnlyToMt.TotalSeconds:0.00}s, MT→OMML={mtOnlyToOmml.TotalSeconds:0.00}s; all 50 semantic signatures and 35 numbered formulas survived; save/reopen passed; MathTypeProcessCount=0; OMMLfixture={ommlFixture}");
                return;
            }

            var vtRoundTrip = Path.Combine(artifactRoot, "Stress-VT-MT-VT-50.docx");
            File.Copy(vtFixture, vtRoundTrip, true);
            document = application.Documents.Open(vtRoundTrip, ReadOnly: false, AddToRecentFiles: false, Visible: true);
            var vtToMt = TimeStressDocumentConversion(
                () => callbacks.OnConvertVisualTeXToMathTypeDocument(null),
                document,
                tracePath,
                "source=VisualTeX target=MathType",
                mathTypeBaseline);
            AssertStressMathTypeInventory(document, prepared, "VT→MT 50 batch");
            AssertEqual(35, CountStressMathTypePlaceRefFields(document), "VT→MT lost numbered equations.");
            var mtToVt = TimeStressDocumentConversion(
                () => callbacks.OnConvertMathTypeToVisualTeXDocument(null),
                document,
                tracePath,
                "source=MathType target=VisualTeX",
                mathTypeBaseline);
            AssertStressVisualTeXInventory(document, prepared, expectCanonical: true, "VT→MT→VT 50 batch");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            var ommlRoundTrip = Path.Combine(artifactRoot, "Stress-OMML-MT-OMML-50.docx");
            File.Copy(ommlFixture, ommlRoundTrip, true);
            document = application.Documents.Open(ommlRoundTrip, ReadOnly: false, AddToRecentFiles: false, Visible: true);
            var ommlToMt = TimeStressDocumentConversion(
                () => callbacks.OnConvertOmmlToMathTypeDocument(null),
                document,
                tracePath,
                "source=OMML target=MathType",
                mathTypeBaseline);
            AssertStressMathTypeInventory(document, prepared, "OMML→MT 50 batch");
            AssertEqual(35, CountStressMathTypePlaceRefFields(document), "OMML→MT lost numbered equations.");
            var mtToOmml = TimeStressDocumentConversion(
                () => callbacks.OnConvertMathTypeToOmmlDocument(null),
                document,
                tracePath,
                "source=MathType target=OMML",
                mathTypeBaseline);
            AssertStressOmmlInventory(document, prepared, "OMML→MT→OMML 50 batch");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            document = application.Documents.Open(vtRoundTrip, ReadOnly: false, AddToRecentFiles: false, Visible: true);
            AssertStressVisualTeXInventory(document, prepared, expectCanonical: true, "reopened VT→MT→VT 50 batch");
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(ommlRoundTrip, ReadOnly: false, AddToRecentFiles: false, Visible: true);
            AssertStressOmmlInventory(document, prepared, "reopened OMML→MT→OMML 50 batch");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "50-formula four-direction batch conversion stress");

            File.WriteAllLines(
                Path.Combine(artifactRoot, "format-50-batch-times.csv"),
                new[]
                {
                    "direction,milliseconds",
                    $"VT->MT,{vtToMt.TotalMilliseconds.ToString("0", CultureInfo.InvariantCulture)}",
                    $"MT->VT,{mtToVt.TotalMilliseconds.ToString("0", CultureInfo.InvariantCulture)}",
                    $"OMML->MT,{ommlToMt.TotalMilliseconds.ToString("0", CultureInfo.InvariantCulture)}",
                    $"MT->OMML,{mtToOmml.TotalMilliseconds.ToString("0", CultureInfo.InvariantCulture)}",
                });
            Console.WriteLine(
                $"[FORMAT50 BATCH PASS] VT→MT={vtToMt.TotalSeconds:0.00}s, MT→VT={mtToVt.TotalSeconds:0.00}s, OMML→MT={ommlToMt.TotalSeconds:0.00}s, MT→OMML={mtToOmml.TotalSeconds:0.00}s; all 50 semantic signatures and 35 numbered formulas survived; save/reopen passed; MathTypeProcessCount=0; VTfixture={vtFixture}; OMMLfixture={ommlFixture}");
        }
        finally
        {
            WordEquationNumbering.SetDefaultCreateObjectMode(previousCreateMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(previousNumbered);
            Release(callbacksObject);
            Release(installedAddIn);
            Release(addIns);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTracePath);
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", previousFormatAcceptance);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", previousAcceptance);
        }
    }

    private static void RunWordInstalledVisualTeXOmmlFiftyDirectStressAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var previousAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        var previousFormatAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousTracePath = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        var previousCreateMode = WordEquationNumbering.GetDefaultCreateObjectMode();
        var previousNumbered = WordEquationNumbering.GetDefaultDisplayEquationNumbered();
        var tracePath = Path.Combine(artifactRoot, "vt-omml-50-direct-stress.trace.log");
        Word.Application? application = null;
        Word.Document? document = null;
        Microsoft.Office.Core.COMAddIns? addIns = null;
        Microsoft.Office.Core.COMAddIn? installedAddIn = null;
        object? callbacksObject = null;
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
            var mathTypeBaseline = SnapshotMathTypeProcessIds();
            if (mathTypeBaseline.Count != 0)
                throw new InvalidOperationException("MathType.exe must not be running before direct VT↔OMML stress.");
            var prepared = PrepareFiftyStressCorpus(client);
            application = CreateWordApplication(visible: true);
            callbacksObject = GetInstalledStressCallbacks(application, out addIns, out installedAddIn);
            dynamic callbacks = callbacksObject;
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(true);

            var suppliedFixture = Environment.GetEnvironmentVariable("VISUALTEX_STRESS_VT_FIXTURE");
            var sourcePath = Path.Combine(artifactRoot, "Stress-VisualTeX-50-Direct-Source.docx");
            if (!string.IsNullOrWhiteSpace(suppliedFixture) && File.Exists(suppliedFixture))
            {
                File.Copy(suppliedFixture, sourcePath, true);
            }
            else
            {
                document = CreateStressSourceDocument(
                    client,
                    application,
                    callbacks,
                    prepared,
                    FormulaOleContract.NativeOleMode,
                    mathTypeBaseline);
                document.SaveAs2(sourcePath, Word.WdSaveFormat.wdFormatXMLDocument);
                document.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(document); document = null;
            }

            var outputPath = Path.Combine(artifactRoot, "Stress-VT-OMML-VT-50.docx");
            File.Copy(sourcePath, outputPath, true);
            document = application.Documents.Open(outputPath, ReadOnly: false, AddToRecentFiles: false, Visible: true);
            AssertStressVisualTeXInventory(document, prepared, expectCanonical: false, "direct VT→OMML source");

            var vtToOmmlWatch = Stopwatch.StartNew();
            for (var index = prepared.Count - 1; index >= 0; index--)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = document.InlineShapes[index + 1];
                    shape.Range.Select();
                    ConvertSelectedStressFormula(
                        client,
                        () => callbacks.OnConvertSelectedToOmml(null),
                        prepared[index],
                        FormulaOleContract.WordOmmlMode,
                        mathTypeBaseline);
                }
                finally { Release(shape); }
                if ((prepared.Count - index) % 10 == 0)
                    Console.WriteLine($"[VT→OMML DIRECT] {prepared.Count - index}/50 elapsed={vtToOmmlWatch.Elapsed.TotalSeconds:0.00}s");
            }
            vtToOmmlWatch.Stop();
            AssertStressOmmlInventory(document, prepared, "direct VT→OMML 50");
            AssertEqual(0, CountInstalledVisualTeXOleShapes(document), "Direct VT→OMML left VisualTeX OLE objects behind.");

            var ommlToVtWatch = Stopwatch.StartNew();
            for (var index = prepared.Count - 1; index >= 0; index--)
            {
                Word.OMath? math = null;
                Word.Range? range = null;
                try
                {
                    math = document.OMaths[index + 1];
                    range = math.Range;
                    range.Select();
                    ConvertSelectedStressFormula(
                        client,
                        () => callbacks.OnConvertSelected(null),
                        prepared[index],
                        FormulaOleContract.NativeOleMode,
                        mathTypeBaseline);
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
                if ((prepared.Count - index) % 10 == 0)
                    Console.WriteLine($"[OMML→VT DIRECT] {prepared.Count - index}/50 elapsed={ommlToVtWatch.Elapsed.TotalSeconds:0.00}s");
            }
            ommlToVtWatch.Stop();
            AssertStressVisualTeXInventory(document, prepared, expectCanonical: false, "direct VT→OMML→VT 50");
            AssertEqual(0, document.OMaths.Count, "Direct OMML→VT left OMML equations behind.");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "direct VT↔OMML 50-formula stress");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(outputPath, ReadOnly: false, AddToRecentFiles: false, Visible: true);
            AssertStressVisualTeXInventory(document, prepared, expectCanonical: false, "reopened direct VT→OMML→VT 50");
            File.WriteAllLines(
                Path.Combine(artifactRoot, "vt-omml-50-direct-times.csv"),
                new[]
                {
                    "direction,milliseconds",
                    $"VT->OMML,{vtToOmmlWatch.Elapsed.TotalMilliseconds.ToString("0", CultureInfo.InvariantCulture)}",
                    $"OMML->VT,{ommlToVtWatch.Elapsed.TotalMilliseconds.ToString("0", CultureInfo.InvariantCulture)}",
                });
            Console.WriteLine(
                $"[VT↔OMML50 DIRECT PASS] real selected-formula Ribbon conversion ran 50 VT→OMML and 50 OMML→VT operations; VT→OMML={vtToOmmlWatch.Elapsed.TotalSeconds:0.00}s, OMML→VT={ommlToVtWatch.Elapsed.TotalSeconds:0.00}s; all 50 formulas and numbered hosts survived save/reopen; MathTypeProcessCount=0; artifact={outputPath}");
        }
        finally
        {
            WordEquationNumbering.SetDefaultCreateObjectMode(previousCreateMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(previousNumbered);
            Release(callbacksObject);
            Release(installedAddIn);
            Release(addIns);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTracePath);
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", previousFormatAcceptance);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", previousAcceptance);
        }
    }

    private static object GetInstalledStressCallbacks(
        Word.Application application,
        out Microsoft.Office.Core.COMAddIns addIns,
        out Microsoft.Office.Core.COMAddIn installedAddIn)
    {
        addIns = application.COMAddIns;
        object addInKey = "VisualTeX.WordVsto";
        installedAddIn = addIns.Item(ref addInKey);
        if (!installedAddIn.Connect) installedAddIn.Connect = true;
        for (var index = 0; index < 100 && installedAddIn.Object is null; index++)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(100);
        }
        return installedAddIn.Object
            ?? throw new InvalidOperationException(
                "Installed VisualTeX.WordVsto automation object was unavailable for 50-formula stress acceptance.");
    }

    private static Word.Document CreateStressSourceDocument(
        VisualTeXSessionClient client,
        Word.Application application,
        dynamic callbacks,
        IReadOnlyList<PreparedStressFormula> prepared,
        string objectMode,
        IReadOnlyCollection<int> mathTypeBaseline)
    {
        var document = application.Documents.Add();
        document.Activate();
        WordEquationNumbering.SetDefaultCreateObjectMode(objectMode);
        for (var index = 0; index < prepared.Count; index++)
        {
            EnsureStressIndependentFormulaHost(document, index);
            var item = prepared[index];
            if (string.Equals(objectMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal))
            {
                if (item.Formula.DisplayMode == "block")
                    CommitInstalledVisualTeXDisplayFromRibbon(
                        client,
                        callbacks,
                        document,
                        item.Formula.Latex,
                        item.Export,
                        index + 1,
                        item.Formula.Numbered);
                else
                    CommitStressCreateFromRibbon(
                        client,
                        () => callbacks.OnInsertInline(null),
                        document,
                        item,
                        FormulaOleContract.NativeOleMode,
                        index + 1,
                        mathTypeBaseline);
            }
            else if (string.Equals(objectMode, FormulaOleContract.WordOmmlMode, StringComparison.Ordinal))
            {
                CommitStressCreateFromRibbon(
                    client,
                    item.Formula.DisplayMode == "block"
                        ? (Action)(() => callbacks.OnInsertDisplayOmml(null))
                        : () => callbacks.OnInsertInlineOmml(null),
                    document,
                    item,
                    FormulaOleContract.WordOmmlMode,
                    index + 1,
                    mathTypeBaseline);
            }
            else
            {
                throw new ArgumentOutOfRangeException(nameof(objectMode), objectMode, null);
            }
            if ((index + 1) % 10 == 0)
                Console.WriteLine($"[SOURCE50 {objectMode}] {index + 1}/50");
        }
        AssertNoNewMathTypeProcess(mathTypeBaseline, $"creating 50 {objectMode} source formulas");
        return document;
    }

    private static void EnsureStressIndependentFormulaHost(Word.Document document, int index)
    {
        if (index <= 0) return;
        // SelectDocumentEnd points immediately before Word's final paragraph mark.
        // After a numbered MathType display this can still be the MTPlaceRef
        // paragraph itself. Insert an explicit paragraph break before creating the
        // next independent stress formula so an inline equation cannot accidentally
        // share the previous equation-number field host.
        AppendAcceptanceText(document, "\r");
    }

    private static void CommitStressCreateFromRibbon(
        VisualTeXSessionClient client,
        Action callback,
        Word.Document document,
        PreparedStressFormula item,
        string objectMode,
        int expectedCount,
        IReadOnlyCollection<int> mathTypeBaseline)
    {
        SelectDocumentEnd(document);
        var existing = SnapshotSessionIds();
        callback();
        var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
        var session = client.GetSessionAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
        AssertEqual(objectMode, session.ObjectMode, "Stress create callback requested the wrong object mode.");
        PatchStressSessionForCommit(client, session, item, objectMode, dirty: true);
        WaitForStressSessionCompletion(client, sessionId, TimeSpan.FromSeconds(45));
        WaitForInstalledRibbonSessionRelease(sessionId);
        AssertNoNewMathTypeProcess(mathTypeBaseline, $"stress create {expectedCount}/50 {objectMode}");
        if (string.Equals(objectMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal))
            AssertEqual(expectedCount, CountInstalledVisualTeXOleShapes(document), "VisualTeX stress create count mismatch.");
        else
            AssertEqual(expectedCount, document.OMaths.Count, "OMML stress create count mismatch.");
    }

    private static void PatchStressSessionForCommit(
        VisualTeXSessionClient client,
        OfficeSessionDocument session,
        PreparedStressFormula item,
        string objectMode,
        bool dirty)
    {
        var lineId = session.Lines.FirstOrDefault()?.Id ?? Guid.NewGuid().ToString("D");
        client.PatchAsync(
                session.Id,
                new Dictionary<string, object>
                {
                    ["lines"] = new[]
                    {
                        new Dictionary<string, object>
                        {
                            ["id"] = lineId,
                            ["latex"] = item.Formula.Latex,
                        },
                    },
                    ["activeLineId"] = lineId,
                    ["codeFormat"] = "latex",
                    ["displayMode"] = item.Formula.DisplayMode,
                    ["objectMode"] = objectMode,
                    ["numbered"] = item.Formula.Numbered,
                    ["mathTypeNumberPosition"] = "right",
                    ["fontSizePt"] = 12d,
                    ["exportWidth"] = item.Export.Width,
                    ["exportHeight"] = item.Export.Height,
                    ["exportResult"] = item.Export,
                    ["dirty"] = dirty,
                    ["status"] = "committing",
                    ["explicitCancel"] = false,
                },
                CancellationToken.None)
            .GetAwaiter().GetResult();
    }

    private static OfficeSessionDocument WaitForStressEditorOpened(
        VisualTeXSessionClient client,
        string sessionId,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        OfficeSessionDocument? session = null;
        while (DateTime.UtcNow < deadline)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(40);
            session = client.GetSessionAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            if (session.Dirty)
                throw new InvalidOperationException(
                    $"Stress edit Session {sessionId} became dirty before any edit was made.");
            if (session.Status is "failed" or "cancelled" or "completed")
                throw new InvalidOperationException(
                    $"Stress edit Session {sessionId} reached {session.Status} before opening: {session.Error}");
            if ((session.Status is "created" or "editing")
                && session.Lines.Count > 0
                && session.Lines.Any(line => !string.IsNullOrWhiteSpace(line.Latex)))
                return session;
        }
        throw new TimeoutException(
            $"Stress edit Session {sessionId} did not expose its source within {timeout}; last={session?.Status ?? "unknown"}.");
    }

    private static void WaitForStressSessionCompletion(
        VisualTeXSessionClient client,
        string sessionId,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(75);
            var current = client.GetSessionAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            if (string.Equals(current.Status, "failed", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(current.Error ?? $"Stress Session {sessionId} failed.");
            if (string.Equals(current.Status, "completed", StringComparison.OrdinalIgnoreCase))
                return;
        }
        throw new TimeoutException($"Stress Session {sessionId} did not complete within {timeout}.");
    }

    private static TimeSpan TimeStressDocumentConversion(
        Action callback,
        Word.Document document,
        string tracePath,
        string marker,
        IReadOnlyCollection<int> mathTypeBaseline)
    {
        document.Activate();
        ResetInstalledFormatConversionTrace(tracePath);
        var watch = Stopwatch.StartNew();
        callback();
        WaitForInstalledOmmlMathTypeConversion(tracePath, marker, mathTypeBaseline);
        watch.Stop();
        AssertNoNewMathTypeProcess(mathTypeBaseline, marker);
        Console.WriteLine($"[FORMAT50] {marker} elapsed={watch.Elapsed.TotalSeconds:0.00}s");
        return watch.Elapsed;
    }

    private static void ConvertSelectedStressFormula(
        VisualTeXSessionClient client,
        Action callback,
        PreparedStressFormula item,
        string objectMode,
        IReadOnlyCollection<int> mathTypeBaseline)
    {
        var existing = SnapshotSessionIds();
        callback();
        var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(20));
        var session = client.GetSessionAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
        AssertEqual(objectMode, session.ObjectMode, "Selected stress conversion requested the wrong target mode.");
        PatchStressSessionForCommit(client, session, item, objectMode, dirty: false);
        WaitForStressSessionCompletion(client, sessionId, TimeSpan.FromSeconds(30));
        WaitForInstalledRibbonSessionRelease(sessionId);
        AssertNoNewMathTypeProcess(mathTypeBaseline, $"selected conversion to {objectMode}");
    }

    private static void AssertStressMathTypeInventory(
        Word.Document document,
        IReadOnlyList<PreparedStressFormula> prepared,
        string context)
    {
        AssertEqual(50, CountMathTypeOleShapes(document), context + ": MathType count mismatch.");
        var seen = 0;
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                if (seen >= prepared.Count)
                    throw new InvalidDataException(context + ": too many MathType objects.");
                var actual = MathTypeMtefCodec.SemanticSignature(MathTypeOleStorage.ReadMathMl(shape));
                AssertEqual(prepared[seen].Signature, actual,
                    context + $": formula #{seen + 1} '{prepared[seen].Formula.Name}' changed semantics.");
                AssertTrue(shape.Width > 1f && shape.Height > 1f,
                    context + $": formula #{seen + 1} has invalid visible size {shape.Width}x{shape.Height}.");
                seen++;
            }
            finally { Release(shape); }
        }
        AssertEqual(50, seen, context + ": did not inspect all MathType formulas.");
    }

    private static void AssertStressVisualTeXInventory(
        Word.Document document,
        IReadOnlyList<PreparedStressFormula> prepared,
        bool expectCanonical,
        string context)
    {
        AssertEqual(50, CountInstalledVisualTeXOleShapes(document), context + ": VisualTeX OLE count mismatch.");
        var seen = 0;
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                var metadata = WordFormulaMetadataReader.TryRead(shape)
                    ?? throw new InvalidDataException(context + $": VisualTeX formula #{seen + 1} has no metadata.");
                var expected = expectCanonical
                    ? prepared[seen].CanonicalLatex
                    : NormalizeStressLatex(prepared[seen].Formula.Latex);
                var actual = NormalizeStressLatex(metadata.Latex ?? string.Empty);
                AssertEqual(expected, actual,
                    context + $": VisualTeX formula #{seen + 1} '{prepared[seen].Formula.Name}' changed LaTeX semantics.");
                AssertEqual(prepared[seen].Formula.Numbered, metadata.Numbered,
                    context + $": VisualTeX formula #{seen + 1} changed numbered state.");
                seen++;
            }
            finally { Release(shape); }
        }
        AssertEqual(50, seen, context + ": did not inspect all VisualTeX formulas.");
    }

    private static void AssertStressOmmlInventory(
        Word.Document document,
        IReadOnlyList<PreparedStressFormula> prepared,
        string context)
    {
        AssertEqual(50, document.OMaths.Count, context + ": OMML count mismatch.");
        for (var index = 0; index < prepared.Count; index++)
        {
            Word.OMath? math = null;
            Word.Range? range = null;
            try
            {
                math = document.OMaths[index + 1];
                range = math.Range;
                var mathMl = WordOmmlConverter.TransformOmmlToMathMl(
                    range.WordOpenXML,
                    display: prepared[index].Formula.DisplayMode == "block");
                var actual = MathTypeMtefCodec.SemanticSignature(mathMl);
                if (!string.Equals(prepared[index].Signature, actual, StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        context + $": OMML formula #{index + 1} '{prepared[index].Formula.Name}' changed semantics. Expected {prepared[index].Signature}, actual {actual}. MathML={mathMl} WordOpenXML={range.WordOpenXML}");
            }
            finally
            {
                Release(range);
                Release(math);
            }
        }
    }

    private static int CountStressMathTypePlaceRefFields(Word.Document document)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        var count = 0;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf("MACROBUTTON MTPlaceRef", StringComparison.OrdinalIgnoreCase) >= 0)
                    count++;
            }
            return count;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static void AssertStressEditPerformance(IReadOnlyList<StressEditMetric> metrics)
    {
        AssertEqual(50, metrics.Count, "MT50 edit stress did not record 50 measurements.");
        var opens = metrics.Select(item => item.OpenMilliseconds).OrderBy(value => value).ToArray();
        var commits = metrics.Select(item => item.CommitMilliseconds).OrderBy(value => value).ToArray();
        var numberedOpens = metrics.Where(item => item.Numbered).Select(item => item.OpenMilliseconds).OrderBy(value => value).ToArray();
        var firstTen = metrics.Take(10).Average(item => item.OpenMilliseconds);
        var lastTen = metrics.Skip(40).Take(10).Average(item => item.OpenMilliseconds);
        var firstNumbered = metrics.Where(item => item.Numbered).Take(5).Average(item => item.OpenMilliseconds);
        var lastNumbered = metrics.Where(item => item.Numbered).Reverse().Take(5).Average(item => item.OpenMilliseconds);
        var p50Open = StressPercentile(opens, 0.50);
        var p95Open = StressPercentile(opens, 0.95);
        var p95Commit = StressPercentile(commits, 0.95);
        var p95Numbered = StressPercentile(numberedOpens, 0.95);
        var maxOpen = opens[opens.Length - 1];
        var maxCommit = commits[commits.Length - 1];
        Console.WriteLine(
            $"[MT50 PERF] open p50={p50Open:0}ms p95={p95Open:0}ms max={maxOpen:0}ms; commit p95={p95Commit:0}ms max={maxCommit:0}ms; numbered-open p95={p95Numbered:0}ms; first10-avg={firstTen:0}ms last10-avg={lastTen:0}ms; first5-numbered={firstNumbered:0}ms last5-numbered={lastNumbered:0}ms.");

        AssertTrue(p95Open <= 5000,
            $"MT50 editor opening is slow: P95={p95Open:0}ms exceeds 5000ms.");
        AssertTrue(maxOpen <= 8000,
            $"MT50 editor opening has an unacceptable outlier: max={maxOpen:0}ms.");
        AssertTrue(p95Numbered <= 5000,
            $"MT50 numbered editor opening is slow: P95={p95Numbered:0}ms exceeds 5000ms.");
        AssertTrue(p95Commit <= 7000 && maxCommit <= 10000,
            $"MT50 edit commit is slow: P95={p95Commit:0}ms max={maxCommit:0}ms.");
        AssertTrue(lastTen <= firstTen * 2.0 + 500,
            $"MT50 editor opening degrades with document size: first10={firstTen:0}ms last10={lastTen:0}ms.");
        AssertTrue(lastNumbered <= firstNumbered * 2.0 + 500,
            $"MT50 numbered editor opening degrades with document size: first5={firstNumbered:0}ms last5={lastNumbered:0}ms.");
    }

    private static double StressPercentile(IReadOnlyList<double> sorted, double percentile)
    {
        if (sorted.Count == 0) return 0;
        var boundedPercentile = Math.Max(0d, Math.Min(1d, percentile));
        var position = boundedPercentile * (sorted.Count - 1);
        var lower = (int)Math.Floor(position);
        var upper = (int)Math.Ceiling(position);
        if (lower == upper) return sorted[lower];
        var weight = position - lower;
        return sorted[lower] * (1d - weight) + sorted[upper] * weight;
    }

    private static void WriteStressEditCsv(string path, IReadOnlyList<StressEditMetric> metrics)
    {
        var lines = new List<string>
        {
            "index,name,numbered,displayMode,openMilliseconds,commitMilliseconds",
        };
        lines.AddRange(metrics.Select(item => string.Join(",",
            item.Index.ToString(CultureInfo.InvariantCulture),
            item.Name,
            item.Numbered ? "true" : "false",
            item.DisplayMode,
            item.OpenMilliseconds.ToString("0", CultureInfo.InvariantCulture),
            item.CommitMilliseconds.ToString("0", CultureInfo.InvariantCulture))));
        File.WriteAllLines(path, lines);
    }
}
