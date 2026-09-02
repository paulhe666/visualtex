using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordInstalledMathTypeNativeRegressionAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var previousAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        var previousFormatAcceptance = Environment.GetEnvironmentVariable(
            "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousTracePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_WORD_HOOK_TRACE_PATH");
        var previousCreateMode = WordEquationNumbering.GetDefaultCreateObjectMode();
        var previousNumbered = WordEquationNumbering.GetDefaultDisplayEquationNumbered();
        var previousNumberPosition = WordEquationNumbering.GetDefaultMathTypeNumberPosition();
        var previousNumberFormat = WordEquationNumbering.GetDefaultEquationNumberFormatId();
        var tracePath = Path.Combine(artifactRoot, "installed-mathtype-native-regression.trace.log");

        Word.Application? application = null;
        Microsoft.Office.Core.COMAddIns? addIns = null;
        Microsoft.Office.Core.COMAddIn? installedAddIn = null;
        object? callbacksObject = null;
        IReadOnlyCollection<int>? mathTypeBaseline = null;
        var acceptanceCompleted = false;
        try
        {
            // This mode must exercise the installed add-in exactly as Word loads it.
            // The acceptance executable is only the external driver/oracle.
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);

            var reverseOnly = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_INSTALLED_MATHTYPE_REVERSE_ONLY"),
                "1",
                StringComparison.Ordinal);
            var firstLeftOnly = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_INSTALLED_MATHTYPE_FIRST_LEFT_ONLY"),
                "1",
                StringComparison.Ordinal);

            mathTypeBaseline = SnapshotMathTypeProcessIds();
            if (firstLeftOnly)
            {
                foreach (var processId in mathTypeBaseline)
                {
                    if (!MathTypeNativePreviewRenderer.IsControlledMathTypeRpcHelperProcess(processId))
                        throw new InvalidOperationException(
                            $"Blank-document first-left acceptance found a non-RPC MathType process before Word starts: pid={processId}.");
                }
            }
            else if (mathTypeBaseline.Count != 0)
            {
                throw new InvalidOperationException(
                    "Installed MathType regression acceptance requires MathType.exe process count to be zero before Word starts.");
            }
            // The reverse-only gate validates the installed VSTO/Word conversion
            // path and does not need screenshot/UI interaction. Keep Word hidden so
            // the low-level double-click hook cannot compete with the COM driver.
            // Full installed-native regression remains visible as before.
            application = CreateWordApplication(visible: !reverseOnly);
            addIns = application.COMAddIns;
            object addInKey = "VisualTeX.WordVsto";
            installedAddIn = addIns.Item(ref addInKey);
            if (!installedAddIn.Connect)
                installedAddIn.Connect = true;
            for (var index = 0; index < 80 && installedAddIn.Object is null; index++)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
            }
            callbacksObject = installedAddIn.Object
                ?? throw new InvalidOperationException(
                    "Installed VisualTeX.WordVsto automation object was unavailable. This acceptance refuses to construct ThisAddIn locally.");
            dynamic callbacks = callbacksObject;

            WordEquationNumbering.SetDefaultCreateObjectMode(FormulaOleContract.MathTypeOleMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(true);
            WordEquationNumbering.SetDefaultEquationNumberFormatPreference(
                EquationNumberFormat.Heading1DotId);

            if (firstLeftOnly)
            {
                RunInstalledFirstLeftMathTypePath(
                    client,
                    application,
                    callbacks,
                    artifactRoot,
                    mathTypeBaseline);
                AssertNoUnexpectedMathTypeProcessDuringInstalledSession(
                    mathTypeBaseline,
                    "installed blank-document first-left MathType acceptance");
                acceptanceCompleted = true;
                Console.WriteLine(
                    "[MATHTYPE INSTALLED FIRST LEFT] Actual installed VisualTeX.WordVsto inserted the first formula in a blank Word document through Ribbon -> Session -> in-process VSTO commit with left numbering.");
                return;
            }

            if (reverseOnly)
            {
                // The full installed regression naturally spends several seconds
                // in earlier paths while the add-in prewarms its windowless native
                // MathType preview session. A reverse-only diagnostic starts at once,
                // so pump Word briefly before the first ribbon action to avoid an
                // acceptance-only RPC_E_CALL_REJECTED startup race.
                for (var settle = 0; settle < 40; settle++)
                {
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(100);
                }
            }
            if (!reverseOnly)
            {
                RunInstalledDirectMathTypePath(
                    client,
                    application,
                    callbacks,
                    artifactRoot,
                    mathTypeBaseline);
                RunInstalledMathTypeFunctionSpacingPath(
                    client,
                    application,
                    callbacks,
                    artifactRoot,
                    mathTypeBaseline);
                RunInstalledOmmlToMathTypeContextPath(
                    application,
                    callbacks,
                    artifactRoot,
                    tracePath,
                    mathTypeBaseline);
                RunInstalledVisualTeXToMathTypeContextPath(
                    client,
                    application,
                    callbacks,
                    artifactRoot,
                    tracePath,
                    mathTypeBaseline);
            }
            RunInstalledMathTypeToVisualTeXPath(
                client,
                application,
                callbacks,
                artifactRoot,
                tracePath,
                mathTypeBaseline);

            AssertNoUnexpectedMathTypeProcessDuringInstalledSession(
                mathTypeBaseline,
                "installed direct/OMML/VisualTeX MathType regression acceptance");
            acceptanceCompleted = true;
            Console.WriteLine(
                "[MATHTYPE INSTALLED NATIVE REGRESSION] Actual installed VisualTeX.WordVsto passed direct MathType display insertion plus OMML→MathType and VisualTeX→MathType document conversion. Word-native MTDisplayEquation/MTPlaceRef layout, left/right number placement, heading-aware MTChap/MTSec state, physical centering, production Times MathType preview selection and save/reopen persistence were verified; any live MathType.exe was restricted to the single controlled windowless -mtrpc preview helper.");
        }
        finally
        {
            WordEquationNumbering.SetDefaultCreateObjectMode(previousCreateMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(previousNumbered);
            WordEquationNumbering.SetDefaultMathTypeNumberPosition(previousNumberPosition);
            WordEquationNumbering.SetDefaultEquationNumberFormatPreference(previousNumberFormat);
            Release(callbacksObject);
            Release(installedAddIn);
            Release(addIns);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            if (acceptanceCompleted && mathTypeBaseline is not null)
                AssertInstalledMathTypeRpcHelpersEventuallyCleaned(
                    mathTypeBaseline,
                    "installed native regression teardown");
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTracePath);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                previousFormatAcceptance);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", previousAcceptance);
        }
    }

    private static void RunInstalledFirstLeftMathTypePath(
        VisualTeXSessionClient client,
        Word.Application application,
        dynamic callbacks,
        string artifactRoot,
        IReadOnlyCollection<int> mathTypeBaseline)
    {
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);
            WordEquationNumbering.SetDefaultMathTypeNumberPosition("left");

            var export = CreateInstalledMathTypeProductExport(
                client,
                @"\frac{a}{b}",
                FormulaOleContract.MathTypeOleMode);
            CommitInstalledMathTypeDisplayFromRibbon(
                client,
                callbacks,
                document,
                @"\frac{a}{b}",
                export,
                "left",
                1,
                mathTypeBaseline);

            AssertEqual(1, CountMathTypeOleShapes(document),
                "Blank-document installed first-left insertion did not create exactly one Equation.DSMT4 object.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Blank-document installed first-left insertion did not create exactly one MTPlaceRef field.");
            AssertMathTypeNumberTexts(document, "(0.1)");
            shape = document.InlineShapes[1];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "installed blank-document first-left equation");
            AssertInstalledMathTypeDisplayGeometry(
                document,
                shape,
                "left",
                "installed blank-document first-left equation");

            var targetPath = Path.Combine(
                artifactRoot,
                "Installed-MathType-Blank-First-Left.docx");
            document.SaveAs2(targetPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(targetPath, ReadOnly: false, Visible: false);
            AssertEqual(1, CountMathTypeOleShapes(document),
                "Blank-document installed first-left equation did not survive save/reopen.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Blank-document installed first-left MTPlaceRef did not survive save/reopen.");
            AssertMathTypeNumberTexts(document, "(0.1)");
            Release(shape);
            shape = document.InlineShapes[1];
            AssertMathTypeDisplayRow(
                shape,
                expectedNumberPosition: "left",
                "reopened installed blank-document first-left equation");
            Console.WriteLine(
                "[installed first-left] Blank Word document -> Ribbon -> real Session -> installed in-process VSTO commit -> left-numbered Equation.DSMT4 passed and survived save/reopen.");
        }
        finally
        {
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }

    private static void RunInstalledDirectMathTypePath(
        VisualTeXSessionClient client,
        Word.Application application,
        dynamic callbacks,
        string artifactRoot,
        IReadOnlyCollection<int> mathTypeBaseline)
    {
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);

            AppendInstalledAcceptanceHeading(document, "Installed Chapter One");
            WordEquationNumbering.SetDefaultMathTypeNumberPosition("right");
            var firstExport = CreateInstalledMathTypeProductExport(
                client,
                @"x^2+1",
                FormulaOleContract.MathTypeOleMode);
            AssertEqual("times", firstExport.FormulaLetterFont ?? string.Empty,
                "The production MathType export did not select its native Times-based preview typography.");
            CommitInstalledMathTypeDisplayFromRibbon(
                client,
                callbacks,
                document,
                @"x^2+1",
                firstExport,
                "right",
                1,
                mathTypeBaseline);

            AppendInstalledAcceptanceHeading(document, "Installed Chapter Two");
            WordEquationNumbering.SetDefaultMathTypeNumberPosition("left");
            var secondExport = CreateInstalledMathTypeProductExport(
                client,
                @"\frac{a}{b}",
                FormulaOleContract.MathTypeOleMode);
            AssertEqual("times", secondExport.FormulaLetterFont ?? string.Empty,
                "The second production MathType export did not retain native Times preview typography.");
            CommitInstalledMathTypeDisplayFromRibbon(
                client,
                callbacks,
                document,
                @"\frac{a}{b}",
                secondExport,
                "left",
                2,
                mathTypeBaseline);

            AssertMathTypeNumberTexts(document, "(1.1)", "(2.1)");
            AssertEqual(2, CountMathTypeSectionBreakFieldsForInstalledRegression(document),
                "Direct installed MathType insertion did not create one native heading-state field per heading scope.");
            // Genuine MathType 7 reference from the user's Word document: a 12 pt
            // Full Size a/b Equation.DSMT4 is 12 x 30.85 pt and its U+0001 object
            // character sits at Word Font.Position=-12. This reference is stored in
            // the acceptance so the test never has to start MathType.exe.
            const float nativeFractionWidthPt = 12f;
            const float nativeFractionHeightPt = 30.85f;
            const int nativeFractionWordPosition = -12;

            shape = document.InlineShapes[1];
            AssertInstalledMathTypeDisplayGeometry(
                document,
                shape,
                "right",
                "installed direct first equation");
            AssertInstalledMathTypeNumberVerticalAlignment(
                application,
                document,
                shape,
                "right",
                Path.Combine(artifactRoot, "installed-numbered-mathtype-right-simple.png"),
                "installed direct right-numbered equation");
            Release(shape); shape = null;
            shape = document.InlineShapes[2];
            AssertInstalledMathTypeDisplayGeometry(
                document,
                shape,
                "left",
                "installed direct second equation",
                nativeFractionWordPosition,
                nativeFractionWidthPt,
                nativeFractionHeightPt);
            var displayFractionWidth = shape.Width;
            var displayFractionHeight = shape.Height;
            AssertInstalledMathTypeNumberVerticalAlignment(
                application,
                document,
                shape,
                "left",
                Path.Combine(artifactRoot, "installed-numbered-mathtype-left-fraction.png"),
                "installed direct left-numbered fraction");

            document.Content.InsertAfter("\rInstalled inline MathType: ");
            var inlineExport = CreateInstalledMathTypeProductExport(
                client,
                @"\frac{a}{b}",
                FormulaOleContract.MathTypeOleMode,
                displayMode: "inline",
                numbered: false);
            CommitInstalledMathTypeInlineFromRibbon(
                client,
                callbacks,
                document,
                @"\frac{a}{b}",
                inlineExport,
                3,
                mathTypeBaseline);
            Release(shape); shape = document.InlineShapes[3];
            AssertInstalledMathTypeInlineGeometry(
                document,
                shape,
                "installed direct inline equation",
                (nativeFractionWidthPt, nativeFractionHeightPt, nativeFractionWordPosition));
            // Word quantizes OLE extents at roughly one tenth of a point on this
            // Office build. Both placements are already checked against the same
            // genuine MathType reference geometry above; allow only that one-step
            // display/inline rounding difference here rather than treating it as
            // formula scaling.
            AssertNear(displayFractionWidth, shape.Width, 0.15f,
                "The same MathType a/b formula changed OLE width between Word display and inline placement.");
            AssertNear(displayFractionHeight, shape.Height, 0.15f,
                "The same MathType a/b formula changed OLE height between Word display and inline placement.");

            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "installed direct MathType insertion");
            AssertNoUnexpectedMathTypeProcessDuringInstalledSession(mathTypeBaseline, "installed direct MathType insertion");

            var path = Path.Combine(artifactRoot, "Installed-Direct-MathType-Native.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false);
            AssertMathTypeNumberTexts(document, "(1.1)", "(2.1)");
            AssertEqual(2, CountMathTypeSectionBreakFieldsForInstalledRegression(document),
                "Direct MathType native heading state changed after save/reopen.");
            Release(shape); shape = document.InlineShapes[1];
            AssertInstalledMathTypeDisplayGeometry(
                document,
                shape,
                "right",
                "reopened installed direct first equation");
            Release(shape); shape = document.InlineShapes[2];
            AssertInstalledMathTypeDisplayGeometry(
                document,
                shape,
                "left",
                "reopened installed direct second equation",
                nativeFractionWordPosition,
                nativeFractionWidthPt,
                nativeFractionHeightPt);
            Release(shape); shape = document.InlineShapes[3];
            AssertInstalledMathTypeInlineGeometry(
                document,
                shape,
                "reopened installed direct inline equation",
                (nativeFractionWidthPt, nativeFractionHeightPt, nativeFractionWordPosition));
        }
        finally
        {
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }

    private static void RunInstalledOmmlToMathTypeContextPath(
        Word.Application application,
        dynamic callbacks,
        string artifactRoot,
        string tracePath,
        IReadOnlyCollection<int> mathTypeBaseline)
    {
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);
            SetMathTypeDocumentNumberPosition(document, numberOnRight: true);
            var service = new WordFormulaService(application);

            AppendInstalledAcceptanceHeading(document, "OMML Chapter One");
            SelectDocumentEnd(document);
            service.InsertOmml(
                CreateInstalledRegressionSourceSession(
                    FormulaOleContract.WordOmmlMode,
                    @"a+b",
                    numbered: true),
                FirstNumberedMathMl);
            AppendInstalledAcceptanceHeading(document, "OMML Chapter Two");
            SelectDocumentEnd(document);
            service.InsertOmml(
                CreateInstalledRegressionSourceSession(
                    FormulaOleContract.WordOmmlMode,
                    @"c+d",
                    numbered: true),
                SecondNumberedMathMl);
            AssertEqual(2, document.OMaths.Count,
                "Installed OMML context fixture did not create two numbered OMath sources.");

            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertOmmlToMathTypeDocument(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=OMML target=MathType",
                mathTypeBaseline);
            AssertEqual(0, document.OMaths.Count,
                "Installed OMML→MathType context conversion left OMath sources behind.");
            AssertEqual(2, CountMathTypeOleShapes(document),
                "Installed OMML→MathType context conversion produced the wrong object count.");
            AssertMathTypeNumberTexts(document, "(1.1)", "(2.1)");
            AssertEqual(2, CountMathTypeSectionBreakFieldsForInstalledRegression(document),
                "Installed OMML→MathType did not establish native heading state in both scopes.");
            shape = document.InlineShapes[1];
            AssertInstalledMathTypeDisplayGeometry(
                document,
                shape,
                "right",
                "installed OMML→MathType first equation");
            Release(shape); shape = document.InlineShapes[2];
            AssertInstalledMathTypeDisplayGeometry(
                document,
                shape,
                "right",
                "installed OMML→MathType second equation");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "installed OMML→MathType context conversion");

            var path = Path.Combine(artifactRoot, "Installed-OMML-To-MathType-Context.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false);
            AssertMathTypeNumberTexts(document, "(1.1)", "(2.1)");
            AssertEqual(2, CountMathTypeSectionBreakFieldsForInstalledRegression(document),
                "Installed OMML→MathType heading state changed after save/reopen.");
        }
        finally
        {
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }

    private static void RunInstalledVisualTeXToMathTypeContextPath(
        VisualTeXSessionClient client,
        Word.Application application,
        dynamic callbacks,
        string artifactRoot,
        string tracePath,
        IReadOnlyCollection<int> mathTypeBaseline)
    {
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);
            SetMathTypeDocumentNumberPosition(document, numberOnRight: false);
            WordEquationNumbering.SetDefaultCreateObjectMode(FormulaOleContract.NativeOleMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(true);

            AppendInstalledAcceptanceHeading(document, "VisualTeX Chapter One");
            var firstExport = CreateInstalledMathTypeProductExport(
                client,
                @"a+b",
                FormulaOleContract.NativeOleMode);
            CommitInstalledVisualTeXDisplayFromRibbon(
                client,
                callbacks,
                document,
                @"a+b",
                firstExport,
                expectedVisualTeXCount: 1);

            AppendInstalledAcceptanceHeading(document, "VisualTeX Chapter Two");
            var secondExport = CreateInstalledMathTypeProductExport(
                client,
                @"c+d",
                FormulaOleContract.NativeOleMode);
            CommitInstalledVisualTeXDisplayFromRibbon(
                client,
                callbacks,
                document,
                @"c+d",
                secondExport,
                expectedVisualTeXCount: 2);
            AssertEqual(2, CountInstalledVisualTeXOleShapes(document),
                "Installed Ribbon path did not create two real VisualTeX source OLE formulas.");

            // The two sources above were created by the same installed Ribbon and
            // companion Session path a user actually invokes. Switch only the
            // document's MathType number-side preference before converting them.
            WordEquationNumbering.SetDefaultCreateObjectMode(FormulaOleContract.MathTypeOleMode);
            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertVisualTeXToMathTypeDocument(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=VisualTeX target=MathType",
                mathTypeBaseline);
            AssertEqual(2, CountMathTypeOleShapes(document),
                "Installed VisualTeX→MathType context conversion produced the wrong MathType count.");
            AssertMathTypeNumberTexts(document, "(1.1)", "(2.1)");
            AssertEqual(2, CountMathTypeSectionBreakFieldsForInstalledRegression(document),
                "Installed VisualTeX→MathType did not establish native heading state in both scopes.");
            shape = document.InlineShapes[1];
            AssertInstalledMathTypeDisplayGeometry(
                document,
                shape,
                "left",
                "installed VisualTeX→MathType first equation");
            Release(shape); shape = document.InlineShapes[2];
            AssertInstalledMathTypeDisplayGeometry(
                document,
                shape,
                "left",
                "installed VisualTeX→MathType second equation");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "installed VisualTeX→MathType context conversion");

            var path = Path.Combine(artifactRoot, "Installed-VisualTeX-To-MathType-Context.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false);
            AssertMathTypeNumberTexts(document, "(1.1)", "(2.1)");
            Release(shape); shape = document.InlineShapes[1];
            AssertInstalledMathTypeDisplayGeometry(
                document,
                shape,
                "left",
                "reopened installed VisualTeX→MathType first equation");
            Release(shape); shape = document.InlineShapes[2];
            AssertInstalledMathTypeDisplayGeometry(
                document,
                shape,
                "left",
                "reopened installed VisualTeX→MathType second equation");
        }
        finally
        {
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }

    private static void RunInstalledMathTypeToVisualTeXPath(
        VisualTeXSessionClient client,
        Word.Application application,
        dynamic callbacks,
        string artifactRoot,
        string tracePath,
        IReadOnlyCollection<int> mathTypeBaseline)
    {
        Word.Document? document = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);
            WordEquationNumbering.SetDefaultCreateObjectMode(FormulaOleContract.MathTypeOleMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(true);
            WordEquationNumbering.SetDefaultMathTypeNumberPosition("right");

            AppendInstalledAcceptanceHeading(document, "Reverse Chapter One");
            var firstExport = CreateInstalledMathTypeProductExport(
                client,
                @"\hbar\omega+1",
                FormulaOleContract.MathTypeOleMode);
            CommitInstalledMathTypeDisplayFromRibbon(
                client,
                callbacks,
                document,
                @"\hbar\omega+1",
                firstExport,
                "right",
                1,
                mathTypeBaseline);

            AppendInstalledAcceptanceHeading(document, "Reverse Chapter Two");
            var secondExport = CreateInstalledMathTypeProductExport(
                client,
                @"\int_0^1 x^2\,dx",
                FormulaOleContract.MathTypeOleMode);
            CommitInstalledMathTypeDisplayFromRibbon(
                client,
                callbacks,
                document,
                @"\int_0^1 x^2\,dx",
                secondExport,
                "right",
                2,
                mathTypeBaseline);
            AssertMathTypeNumberTexts(document, "(1.1)", "(2.1)");

            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertMathTypeToVisualTeXDocument(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=MathType target=VisualTeX",
                mathTypeBaseline);
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Installed MathType→VisualTeX conversion left MathType source objects behind.");
            AssertEqual(2, CountInstalledVisualTeXOleShapes(document),
                "Installed MathType→VisualTeX conversion did not create two VisualTeX OLE formulas.");
            AssertEqual(2, CountInstalledVisualTeXNumberedFormulaHosts(document),
                "Installed MathType→VisualTeX conversion did not preserve the two numbered display formulas.");
            AssertNoUnexpectedMathTypeProcessDuringInstalledSession(mathTypeBaseline, "installed MathType→VisualTeX conversion");

            var path = Path.Combine(artifactRoot, "Installed-MathType-To-VisualTeX-Context.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false);
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Reopened installed MathType→VisualTeX document restored a MathType source object.");
            AssertEqual(2, CountInstalledVisualTeXOleShapes(document),
                "Reopened installed MathType→VisualTeX document lost a VisualTeX target formula.");
            AssertEqual(2, CountInstalledVisualTeXNumberedFormulaHosts(document),
                "Reopened installed MathType→VisualTeX document changed numbered state.");
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

    private static void RunInstalledMathTypeFunctionSpacingPath(
        VisualTeXSessionClient client,
        Word.Application application,
        dynamic callbacks,
        string artifactRoot,
        IReadOnlyCollection<int> mathTypeBaseline)
    {
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(false);

            const string rankLatex = @"\operatorname{rank}";
            var displayExport = CreateInstalledMathTypeProductExport(
                client,
                rankLatex,
                FormulaOleContract.MathTypeOleMode,
                displayMode: "block",
                numbered: false);
            CommitInstalledMathTypeDisplayFromRibbon(
                client,
                callbacks,
                document,
                rankLatex,
                displayExport,
                "right",
                1,
                mathTypeBaseline,
                numbered: false);
            shape = document.InlineShapes[1];
            AssertInstalledMathTypeFunctionRunExpanded(
                shape,
                "installed display MathType operatorname rank");

            document.Content.InsertAfter("\rInstalled inline function: ");
            var inlineExport = CreateInstalledMathTypeProductExport(
                client,
                rankLatex,
                FormulaOleContract.MathTypeOleMode,
                displayMode: "inline",
                numbered: false);
            CommitInstalledMathTypeInlineFromRibbon(
                client,
                callbacks,
                document,
                rankLatex,
                inlineExport,
                2,
                mathTypeBaseline);
            Release(shape); shape = document.InlineShapes[2];
            AssertInstalledMathTypeFunctionRunExpanded(
                shape,
                "installed inline MathType operatorname rank");

            document.Content.InsertAfter("\rInstalled display lim: ");
            const string limLatex = @"\lim";
            var limDisplayExport = CreateInstalledMathTypeProductExport(
                client,
                limLatex,
                FormulaOleContract.MathTypeOleMode,
                displayMode: "block",
                numbered: false);
            CommitInstalledMathTypeDisplayFromRibbon(
                client,
                callbacks,
                document,
                limLatex,
                limDisplayExport,
                "right",
                3,
                mathTypeBaseline,
                numbered: false);
            Release(shape); shape = document.InlineShapes[3];
            AssertInstalledMathTypeFunctionRunExpanded(
                shape,
                "installed display MathType lim",
                minimumCoverage: 0.50);

            document.Content.InsertAfter("\rInstalled inline lim: ");
            var limInlineExport = CreateInstalledMathTypeProductExport(
                client,
                limLatex,
                FormulaOleContract.MathTypeOleMode,
                displayMode: "inline",
                numbered: false);
            CommitInstalledMathTypeInlineFromRibbon(
                client,
                callbacks,
                document,
                limLatex,
                limInlineExport,
                4,
                mathTypeBaseline);
            Release(shape); shape = document.InlineShapes[4];
            AssertInstalledMathTypeFunctionRunExpanded(
                shape,
                "installed inline MathType lim",
                minimumCoverage: 0.50);

            AssertNoUnexpectedMathTypeProcessDuringInstalledSession(mathTypeBaseline, "installed MathType function spacing");
            var path = Path.Combine(artifactRoot, "Installed-MathType-Function-Spacing.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false);
            Release(shape); shape = document.InlineShapes[1];
            AssertInstalledMathTypeFunctionRunExpanded(
                shape,
                "reopened display MathType operatorname rank");
            Release(shape); shape = document.InlineShapes[2];
            AssertInstalledMathTypeFunctionRunExpanded(
                shape,
                "reopened inline MathType operatorname rank");
            Release(shape); shape = document.InlineShapes[3];
            AssertInstalledMathTypeFunctionRunExpanded(
                shape,
                "reopened display MathType lim",
                minimumCoverage: 0.50);
            Release(shape); shape = document.InlineShapes[4];
            AssertInstalledMathTypeFunctionRunExpanded(
                shape,
                "reopened inline MathType lim",
                minimumCoverage: 0.50);
        }
        finally
        {
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(true);
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }

    private static void AssertInstalledMathTypeFunctionRunExpanded(
        Word.InlineShape shape,
        string context,
        double minimumCoverage = 0.70)
    {
        var preview = ReadInlineShapeEnhancedMetafile(shape);
        using var bitmap = RenderEmf(preview, 600, 180);
        var minX = bitmap.Width;
        var maxX = -1;
        var ink = 0;
        for (var y = 0; y < bitmap.Height; y++)
        {
            for (var x = 0; x < bitmap.Width; x++)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R > 245 && pixel.G > 245 && pixel.B > 245) continue;
                ink++;
                minX = Math.Min(minX, x);
                maxX = Math.Max(maxX, x);
            }
        }
        if (ink == 0 || maxX < minX)
            throw new InvalidDataException(context + ": MathType function preview is empty.");
        var coverage = (maxX - minX + 1d) / bitmap.Width;
        Console.WriteLine(
            $"[installed MathType function spacing] {context}: shape={shape.Width:0.###}x{shape.Height:0.###}pt, inkX={minX}-{maxX}, coverage={coverage:0.000}, ink={ink}, "
            + $"left/right={MeasureLeftWhiteMargin(bitmap)}/{MeasureRightWhiteMargin(bitmap)}, edge={DescribeEdgeInk(bitmap)}.");
        AssertTrue(
            coverage >= minimumCoverage,
            context + $": multi-letter function glyphs collapsed/overlapped horizontally; ink coverage={coverage:0.000}, minimum={minimumCoverage:0.000}.");
    }

    private static OfficeExportDocument CreateInstalledMathTypeProductExport(
        VisualTeXSessionClient client,
        string latex,
        string objectMode,
        string displayMode = "block",
        bool numbered = true)
    {
        var line = new FormulaLine
        {
            Id = Guid.NewGuid().ToString("D"),
            Latex = latex,
        };
        var session = client.CreateSessionAsync(
                new CreateVstoSessionRequest
                {
                    Mode = "create",
                    Host = "word",
                    Title = "Installed MathType native regression renderer",
                    Lines = new List<FormulaLine> { line },
                    ActiveLineId = line.Id,
                    CodeFormat = "latex",
                    DisplayMode = displayMode,
                    ObjectMode = objectMode,
                    Numbered = numbered,
                    MathTypeNumberPosition = "right",
                    FontSizePt = 12d,
                    AutoCommitOnClose = false,
                },
                CancellationToken.None)
            .GetAwaiter().GetResult();
        client.OpenConverterAsync(session.Id, CancellationToken.None)
            .GetAwaiter().GetResult();
        session = client.WaitForCommitAsync(
                session.Id,
                TimeSpan.FromMinutes(3),
                CancellationToken.None)
            .GetAwaiter().GetResult();
        if (string.Equals(session.Status, "failed", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                session.Error ?? "VisualTeX product converter failed while preparing MathType export.");
        var export = session.ExportResult
            ?? throw new InvalidOperationException(
                "VisualTeX product converter returned no MathType export payload.");
        client.CompleteAsync(session.Id, CancellationToken.None).GetAwaiter().GetResult();
        return export;
    }

    private static void CommitInstalledMathTypeInlineFromRibbon(
        VisualTeXSessionClient client,
        dynamic callbacks,
        Word.Document document,
        string latex,
        OfficeExportDocument export,
        int expectedMathTypeCount,
        IReadOnlyCollection<int> mathTypeBaseline,
        bool allowTransientMathType = false)
    {
        SelectDocumentEnd(document);
        var existing = SnapshotSessionIds();
        callbacks.OnInsertInline(null);
        var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
        var session = client.GetSessionAsync(sessionId, CancellationToken.None)
            .GetAwaiter().GetResult();
        AssertEqual(FormulaOleContract.MathTypeOleMode, session.ObjectMode,
            "Installed Ribbon inline insertion did not honor MathType as the active create mode.");
        var lineId = session.Lines.FirstOrDefault()?.Id ?? Guid.NewGuid().ToString("D");
        client.PatchAsync(
                sessionId,
                new Dictionary<string, object>
                {
                    ["lines"] = new[]
                    {
                        new Dictionary<string, object>
                        {
                            ["id"] = lineId,
                            ["latex"] = latex,
                        },
                    },
                    ["activeLineId"] = lineId,
                    ["codeFormat"] = "latex",
                    ["displayMode"] = "inline",
                    ["objectMode"] = FormulaOleContract.MathTypeOleMode,
                    ["numbered"] = false,
                    ["fontSizePt"] = 12d,
                    ["exportWidth"] = export.Width,
                    ["exportHeight"] = export.Height,
                    ["exportResult"] = export,
                    ["dirty"] = true,
                    ["status"] = "committing",
                    ["explicitCancel"] = false,
                },
                CancellationToken.None)
            .GetAwaiter().GetResult();

        var deadline = DateTime.UtcNow.AddSeconds(90);
        while (DateTime.UtcNow < deadline)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(100);
            if (!allowTransientMathType)
                AssertNoUnexpectedMathTypeProcessDuringInstalledSession(mathTypeBaseline, "installed Ribbon direct inline MathType commit");
            var current = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            if (string.Equals(current.Status, "failed", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    current.Error ?? "Installed Ribbon inline MathType insertion failed.");
            if (string.Equals(current.Status, "completed", StringComparison.OrdinalIgnoreCase)
                && CountMathTypeOleShapes(document) == expectedMathTypeCount)
            {
                WaitForInstalledRibbonSessionRelease(sessionId);
                return;
            }
        }
        throw new TimeoutException(
            $"Installed Ribbon inline MathType insertion did not complete with {expectedMathTypeCount} Equation.DSMT4 objects.");
    }

    private static void CommitInstalledMathTypeDisplayFromRibbon(
        VisualTeXSessionClient client,
        dynamic callbacks,
        Word.Document document,
        string latex,
        OfficeExportDocument export,
        string numberPosition,
        int expectedMathTypeCount,
        IReadOnlyCollection<int> mathTypeBaseline,
        bool numbered = true)
    {
        SelectDocumentEnd(document);
        var existing = SnapshotSessionIds();
        callbacks.OnInsertDisplay(null);
        var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
        var session = client.GetSessionAsync(sessionId, CancellationToken.None)
            .GetAwaiter().GetResult();
        AssertEqual(FormulaOleContract.MathTypeOleMode, session.ObjectMode,
            "Installed Ribbon display insertion did not honor MathType as the active create mode.");
        var lineId = session.Lines.FirstOrDefault()?.Id ?? Guid.NewGuid().ToString("D");
        client.PatchAsync(
                sessionId,
                new Dictionary<string, object>
                {
                    ["lines"] = new[]
                    {
                        new Dictionary<string, object>
                        {
                            ["id"] = lineId,
                            ["latex"] = latex,
                        },
                    },
                    ["activeLineId"] = lineId,
                    ["codeFormat"] = "latex",
                    ["displayMode"] = "block",
                    ["objectMode"] = FormulaOleContract.MathTypeOleMode,
                    ["numbered"] = numbered,
                    ["mathTypeNumberPosition"] = numberPosition,
                    ["fontSizePt"] = 12d,
                    ["exportWidth"] = export.Width,
                    ["exportHeight"] = export.Height,
                    ["exportResult"] = export,
                    ["dirty"] = true,
                    ["status"] = "committing",
                    ["explicitCancel"] = false,
                },
                CancellationToken.None)
            .GetAwaiter().GetResult();

        var deadline = DateTime.UtcNow.AddSeconds(90);
        while (DateTime.UtcNow < deadline)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(100);
            AssertNoUnexpectedMathTypeProcessDuringInstalledSession(mathTypeBaseline, "installed Ribbon direct MathType commit");
            var current = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            if (string.Equals(current.Status, "failed", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    current.Error ?? "Installed Ribbon MathType insertion failed.");
            if (string.Equals(current.Status, "completed", StringComparison.OrdinalIgnoreCase)
                && CountMathTypeOleShapes(document) == expectedMathTypeCount)
            {
                WaitForInstalledRibbonSessionRelease(sessionId);
                return;
            }
        }
        throw new TimeoutException(
            $"Installed Ribbon MathType insertion did not complete with {expectedMathTypeCount} Equation.DSMT4 objects.");
    }

    private static void CommitInstalledVisualTeXDisplayFromRibbon(
        VisualTeXSessionClient client,
        dynamic callbacks,
        Word.Document document,
        string latex,
        OfficeExportDocument export,
        int expectedVisualTeXCount,
        bool numbered = true)
    {
        SelectDocumentEnd(document);
        var existing = SnapshotSessionIds();
        callbacks.OnInsertDisplay(null);
        var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
        var session = client.GetSessionAsync(sessionId, CancellationToken.None)
            .GetAwaiter().GetResult();
        AssertEqual(FormulaOleContract.NativeOleMode, session.ObjectMode,
            "Installed Ribbon display insertion did not honor VisualTeX OLE as the active create mode.");
        var lineId = session.Lines.FirstOrDefault()?.Id ?? Guid.NewGuid().ToString("D");
        client.PatchAsync(
                sessionId,
                new Dictionary<string, object>
                {
                    ["lines"] = new[]
                    {
                        new Dictionary<string, object>
                        {
                            ["id"] = lineId,
                            ["latex"] = latex,
                        },
                    },
                    ["activeLineId"] = lineId,
                    ["codeFormat"] = "latex",
                    ["displayMode"] = "block",
                    ["objectMode"] = FormulaOleContract.NativeOleMode,
                    ["numbered"] = numbered,
                    ["fontSizePt"] = 12d,
                    ["exportWidth"] = export.Width,
                    ["exportHeight"] = export.Height,
                    ["exportResult"] = export,
                    ["dirty"] = true,
                    ["status"] = "committing",
                    ["explicitCancel"] = false,
                },
                CancellationToken.None)
            .GetAwaiter().GetResult();

        var deadline = DateTime.UtcNow.AddSeconds(90);
        while (DateTime.UtcNow < deadline)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(100);
            var current = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            if (string.Equals(current.Status, "failed", StringComparison.OrdinalIgnoreCase))
            {
                var detail = current.Error ?? "Installed Ribbon VisualTeX OLE insertion failed.";
                Console.WriteLine(
                    "[INSTALLED VISUALTEX INSERT ERROR B64] "
                    + Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(detail)));
                throw new InvalidOperationException(detail);
            }
            if (string.Equals(current.Status, "completed", StringComparison.OrdinalIgnoreCase)
                && CountInstalledVisualTeXOleShapes(document) == expectedVisualTeXCount)
            {
                WaitForInstalledRibbonSessionRelease(sessionId);
                return;
            }
        }
        throw new TimeoutException(
            $"Installed Ribbon VisualTeX OLE insertion did not complete with {expectedVisualTeXCount} source objects.");
    }

    private static void WaitForInstalledRibbonSessionRelease(string sessionId)
    {
        var tracePath = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        if (string.IsNullOrWhiteSpace(tracePath))
            throw new InvalidOperationException(
                "Installed Ribbon session-release acceptance requires VISUALTEX_WORD_HOOK_TRACE_PATH.");
        var marker = $"ribbon-session-operation-released sessionId={sessionId}";
        var deadline = DateTime.UtcNow.AddSeconds(15);
        while (DateTime.UtcNow < deadline)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(50);
            try
            {
                if (File.Exists(tracePath)
                    && File.ReadAllText(tracePath).IndexOf(marker, StringComparison.Ordinal) >= 0)
                    return;
            }
            catch { }
        }
        throw new TimeoutException(
            $"Installed Ribbon session {sessionId} reached completed state but did not release the Word operation gate.");
    }

    private static OfficeSessionDocument CreateInstalledRegressionSourceSession(
        string objectMode,
        string latex,
        bool numbered) =>
        new()
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = Guid.NewGuid().ToString("D"),
            Title = "Installed MathType regression source",
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = objectMode,
            Numbered = numbered,
            MathTypeNumberPosition = "right",
            FontSizePt = 12d,
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
            ExportResult = new OfficeExportDocument
            {
                Width = 96,
                Height = 40,
                Baseline = 30,
            },
        };

    private static void AppendInstalledAcceptanceHeading(
        Word.Document document,
        string text)
    {
        Word.Range? insertion = null;
        Word.Range? headingRange = null;
        try
        {
            var start = Math.Max(document.Content.Start, document.Content.End - 1);
            insertion = document.Range(start, start);
            var documentIsEmpty = document.Content.End - document.Content.Start <= 1;
            // A collapsed range immediately before Word's final paragraph mark is
            // still part of the preceding MathType display paragraph.  Appending a
            // heading there would restyle the equation row itself as Heading 1.
            // Split first whenever content already exists, then style only the new
            // heading text in its own paragraph.
            var prefix = documentIsEmpty ? string.Empty : "\r";
            insertion.InsertAfter(prefix + text + "\r");
            var headingStart = start + prefix.Length;
            headingRange = document.Range(headingStart, headingStart + text.Length);
            object headingStyle = Word.WdBuiltinStyle.wdStyleHeading1;
            headingRange.set_Style(ref headingStyle);
        }
        finally
        {
            Release(headingRange);
            Release(insertion);
        }
    }

    private static void SetMathTypeDocumentNumberPosition(
        Word.Document document,
        bool numberOnRight)
    {
        object? propertiesObject = null;
        object? propertyObject = null;
        try
        {
            propertiesObject = document.CustomDocumentProperties;
            dynamic properties = propertiesObject;
            try
            {
                propertyObject = properties["MTEqnNumsOnRight"];
                dynamic property = propertyObject;
                property.Value = numberOnRight;
            }
            catch
            {
                propertyObject = properties.Add(
                    "MTEqnNumsOnRight",
                    false,
                    Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeBoolean,
                    numberOnRight);
            }
        }
        finally
        {
            Release(propertyObject);
            Release(propertiesObject);
        }
    }

    private static int CountMathTypeSectionBreakFieldsForInstalledRegression(
        Word.Document document)
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
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
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

    private static int ReadInstalledMathTypeObjectCharacterPosition(
        Word.Document document,
        Word.Range shapeRange,
        string context)
    {
        Word.Range? probe = null;
        Word.Font? font = null;
        try
        {
            for (var position = shapeRange.Start; position < shapeRange.End; position++)
            {
                Release(probe); probe = document.Range(position, position + 1);
                if (!string.Equals(probe.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                Release(font); font = probe.Font;
                var value = font.Position;
                if (value == (int)Word.WdConstants.wdUndefined)
                    throw new InvalidOperationException(
                        context + ": MathType OLE object character has undefined Word Font.Position.");
                return value;
            }
            throw new InvalidOperationException(
                context + ": MathType OLE range contains no U+0001 object character.");
        }
        finally
        {
            Release(font);
            Release(probe);
        }
    }

    private static void AssertInstalledMathTypeInlineGeometry(
        Word.Document document,
        Word.InlineShape shape,
        string context,
        (float WidthPt, float HeightPt, int WordPosition) expected)
    {
        Word.Range? shapeRange = null;
        try
        {
            shapeRange = shape.Range;
            AssertNear(expected.WidthPt, shape.Width, 0.6f,
                context + ": inline MathType OLE width differs from the genuine MathType reference geometry.");
            AssertNear(expected.HeightPt, shape.Height, 0.6f,
                context + ": inline MathType OLE height differs from the genuine MathType reference geometry.");
            var actualWordPosition = ReadInstalledMathTypeObjectCharacterPosition(
                document,
                shapeRange,
                context);
            AssertEqual(expected.WordPosition, actualWordPosition,
                context + ": inline MathType OLE baseline differs from the genuine MathType reference geometry.");
        }
        finally
        {
            Release(shapeRange);
        }
    }

    private static void AssertInstalledMathTypeNumberVerticalAlignment(
        Word.Application application,
        Word.Document document,
        Word.InlineShape shape,
        string numberPosition,
        string imagePath,
        string context)
    {
        Word.Range? shapeRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Selection? selection = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? fieldCode = null;
        Word.Range? numberRange = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        var originalNumberColor = Word.WdColor.wdColorAutomatic;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            paragraphRange.End = Math.Max(paragraphRange.Start, paragraphRange.End - 1);

            fields = paragraphRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(fieldCode); fieldCode = null;
                Release(field); field = fields[index];
                fieldCode = field.Code;
                if ((fieldCode.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                if (!MathTypeEquationReferences.TryGetVisibleNumberRange(
                        document,
                        field,
                        out numberRange)
                    || numberRange is null)
                    throw new InvalidOperationException(context + ": MTPlaceRef has no visible number range.");
                break;
            }
            AssertTrue(numberRange is not null, context + ": native MTPlaceRef field is missing.");
            numberFont = numberRange!.Font;
            originalNumberColor = numberFont.Color;
            numberFont.Color = Word.WdColor.wdColorRed;

            selection = application.Selection;
            selection.SetRange(paragraphRange.Start, paragraphRange.End);
            document.Activate();
            application.ActiveWindow.Activate();
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(200);
            System.Windows.Forms.Clipboard.Clear();
            selection.CopyAsPicture();
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(350);
            ExportClipboardPictureThroughPowerPoint(imagePath);

            using var bitmap = new System.Drawing.Bitmap(imagePath);
            var number = FindPixelBounds(bitmap, static color =>
                color.A > 0
                && color.R >= 120
                && color.R >= color.G + 45
                && color.R >= color.B + 45);
            var formula = FindPixelBounds(bitmap, static color =>
                color.A > 0
                && color.R <= 95
                && color.G <= 95
                && color.B <= 95);
            if (number.Count < 10 || formula.Count < 20)
                throw new InvalidDataException(
                    context + $": numbered MathType pixel classification was insufficient; number={number.Count}, formula={formula.Count}.");
            var numberCenterY = (number.MinY + number.MaxY) / 2.0;
            var formulaCenterY = (formula.MinY + formula.MaxY) / 2.0;
            var centerDelta = numberCenterY - formulaCenterY;
            Console.WriteLine(
                $"[installed MathType number vertical] {context}: "
                + $"formula=({formula.MinX},{formula.MinY})-({formula.MaxX},{formula.MaxY}), "
                + $"number=({number.MinX},{number.MinY})-({number.MaxX},{number.MaxY}), "
                + $"centerDelta={centerDelta:F1}px, image={imagePath}");
            if (Math.Abs(centerDelta) > 6.0)
                throw new InvalidDataException(
                    context + $": MathType equation number is vertically misaligned by {centerDelta:F1}px. Screenshot: {imagePath}");
        }
        finally
        {
            if (numberFont is not null)
            {
                try { numberFont.Color = originalNumberColor; } catch { }
            }
            Release(numberFont);
            Release(numberRange);
            Release(fieldCode);
            Release(field);
            Release(fields);
            Release(selection);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static void AssertInstalledMathTypeDisplayGeometry(
        Word.Document document,
        Word.InlineShape shape,
        string expectedNumberPosition,
        string context,
        int? expectedWordPosition = null,
        float? expectedWidthPt = null,
        float? expectedHeightPt = null)
    {
        Word.Range? shapeRange = null;
        Word.Range? shapeStart = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabs = null;
        Word.TabStop? tab = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? fieldCode = null;
        Word.Range? numberRange = null;
        Word.Range? numberStart = null;
        Word.Range? numberEnd = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
        Word.Style? paragraphStyle = null;
        try
        {
            document.Repaginate();
            Thread.Sleep(80);
            shapeRange = shape.Range;
            if (expectedWidthPt.HasValue)
                AssertNear(expectedWidthPt.Value, shape.Width, 0.6f,
                    context + ": MathType OLE width differs from the genuine MathType reference geometry.");
            if (expectedHeightPt.HasValue)
                AssertNear(expectedHeightPt.Value, shape.Height, 0.6f,
                    context + ": MathType OLE height differs from the genuine MathType reference geometry.");
            if (expectedWordPosition.HasValue)
            {
                var actualWordPosition = ReadInstalledMathTypeObjectCharacterPosition(
                    document,
                    shapeRange,
                    context);
                AssertEqual(
                    expectedWordPosition.Value,
                    actualWordPosition,
                    context + ": MathType OLE object character is not using the genuine MathType reference baseline.");
                AssertTrue(
                    expectedWordPosition.Value < 0,
                    context + ": regression fixture did not exercise a display formula that needs baseline compensation.");
            }
            paragraphs = shapeRange.Paragraphs;
            AssertEqual(1, paragraphs.Count, context + ": MathType OLE spans multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            object styleObject = paragraphRange.get_Style();
            paragraphStyle = styleObject as Word.Style;
            var styleName = paragraphStyle?.NameLocal ?? Convert.ToString(styleObject) ?? string.Empty;
            Console.WriteLine(
                $"[installed MathType geometry] {context}: style='{styleName}' styleType='{styleObject?.GetType().FullName ?? "<null>"}' paragraph={paragraphRange.Start}-{paragraphRange.End} shape={shapeRange.Start}-{shapeRange.End} textCodes={string.Join(",", (paragraphRange.Text ?? string.Empty).Take(24).Select(ch => $"U+{(int)ch:X4}"))}");
            AssertTrue(styleName.IndexOf("MTDisplayEquation", StringComparison.OrdinalIgnoreCase) >= 0,
                context + $": paragraph is not using MathType's MTDisplayEquation style; actual='{styleName}'.");
            format = paragraph.Format;
            tabs = format.TabStops;
            var sawCenter = false;
            var sawRight = false;
            for (var index = 1; index <= tabs.Count; index++)
            {
                Release(tab); tab = tabs[index];
                sawCenter |= tab.Alignment == Word.WdTabAlignment.wdAlignTabCenter;
                sawRight |= tab.Alignment == Word.WdTabAlignment.wdAlignTabRight;
            }
            AssertTrue(sawCenter && sawRight,
                context + ": MathType center/right tab stops are missing.");

            fields = paragraphRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(fieldCode); fieldCode = null;
                Release(field); field = fields[index];
                fieldCode = field.Code;
                if ((fieldCode.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                if (!MathTypeEquationReferences.TryGetVisibleNumberRange(
                        document,
                        field,
                        out numberRange)
                    || numberRange is null)
                    throw new InvalidOperationException(context + ": MTPlaceRef has no visible number range.");
                break;
            }
            AssertTrue(numberRange is not null, context + ": native MTPlaceRef field is missing.");

            shapeStart = shapeRange.Duplicate;
            shapeStart.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            numberStart = numberRange!.Duplicate;
            numberStart.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            numberEnd = numberRange.Duplicate;
            numberEnd.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            var shapeStartX = Convert.ToSingle(
                shapeStart.Information[Word.WdInformation.wdHorizontalPositionRelativeToPage]);
            var shapeCenterX = shapeStartX + shape.Width / 2f;
            var numberStartX = Convert.ToSingle(
                numberStart.Information[Word.WdInformation.wdHorizontalPositionRelativeToPage]);
            var numberEndX = Convert.ToSingle(
                numberEnd.Information[Word.WdInformation.wdHorizontalPositionRelativeToPage]);

            sections = shapeRange.Sections;
            section = sections[1];
            pageSetup = section.PageSetup;
            var textLeft = pageSetup.LeftMargin;
            var textRight = pageSetup.PageWidth - pageSetup.RightMargin;
            var textCenter = (textLeft + textRight) / 2f;
            AssertNear(textCenter, shapeCenterX, 20f,
                context + ": MathType equation is not physically centered in the Word text area.");
            if (string.Equals(expectedNumberPosition, "left", StringComparison.OrdinalIgnoreCase))
            {
                AssertNear(textLeft, numberStartX, 24f,
                    context + ": native MathType number is not left-aligned.");
                AssertTrue(numberEndX < shapeStartX - 4f,
                    context + ": left MathType number overlaps the centered equation.");
            }
            else
            {
                AssertNear(textRight, numberEndX, 24f,
                    context + ": native MathType number is not right-aligned.");
                AssertTrue(numberStartX > shapeStartX + shape.Width + 4f,
                    context + ": right MathType number overlaps the centered equation.");
            }
        }
        finally
        {
            Release(paragraphStyle);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(numberEnd);
            Release(numberStart);
            Release(numberRange);
            Release(fieldCode);
            Release(field);
            Release(fields);
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeStart);
            Release(shapeRange);
        }
    }

    private static void AssertNoUnexpectedMathTypeProcessDuringInstalledSession(
        IReadOnlyCollection<int> baseline,
        string stage)
    {
        var started = SnapshotMathTypeProcessIds()
            .Except(baseline)
            .ToArray();
        AssertTrue(started.Length <= 1,
            $"More than one MathType.exe process started during {stage}: {string.Join(", ", started)}.");
        foreach (var processId in started)
        {
            AssertTrue(
                MathTypeNativePreviewRenderer.IsControlledMathTypeRpcHelperProcess(processId),
                $"Unexpected MathType.exe process started during {stage}: pid={processId}. Only one windowless -mtrpc preview helper is allowed.");
        }
    }

    private static void AssertInstalledMathTypeRpcHelpersEventuallyCleaned(
        IReadOnlyCollection<int> baseline,
        string stage)
    {
        var deadline = DateTime.UtcNow.AddSeconds(8);
        int[] started;
        do
        {
            started = SnapshotMathTypeProcessIds()
                .Except(baseline)
                .ToArray();
            if (started.Length == 0) return;
            Thread.Sleep(100);
        }
        while (DateTime.UtcNow < deadline);

        throw new InvalidOperationException(
            $"MathType preview helper did not clean up after {stage}: {string.Join(", ", started)}.");
    }
}
