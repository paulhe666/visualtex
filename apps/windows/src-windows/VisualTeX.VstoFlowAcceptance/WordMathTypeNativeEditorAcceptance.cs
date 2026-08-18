using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeNativeEditorAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Path.Combine(
            artifactRoot,
            $"VisualTeX-MathType7-NativeEditor-{Guid.NewGuid():N}.docx");
        var fallback = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", "..", "..",
            "artifacts", "mathtype-native-editor",
            "VisualTeX-MathType7-NativeEditor-5f04f8b3545e444a824705446e314ba1.docx"));
        // Always start from the untouched genuine MathType fixture. A successful
        // prior acceptance intentionally saves edits into sourcePath, so reusing
        // it would make the next run look like the read path returned stale data.
        if (File.Exists(fallback))
            File.Copy(fallback, sourcePath, overwrite: true);
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException(
                "A genuine MathType Equation.DSMT4 Word fixture is required.", sourcePath);

        var nativeCaseFilter = Environment.GetEnvironmentVariable(
            "VISUALTEX_MATHTYPE_NATIVE_CASE");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.OLEFormat? format = null;
        var clipboardBackup = TryCaptureClipboard();
        try
        {
            application = CreateWordApplication(visible: true);
            document = application.Documents.Open(sourcePath, ReadOnly: false, Visible: true);
            document.Activate();
            AssertEqual(1, document.InlineShapes.Count,
                "Native MathType editor acceptance expected one inline OLE equation.");
            shape = document.InlineShapes[1];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                "Native MathType editor acceptance source is not MathType OLE.");
            format = shape.OLEFormat;

            Console.WriteLine("[MathType native editor 1/4] Opening the actual Word-owned OLE editor and copying its MathML without saving...");
            var sourceMathMl = InvokeWordOwnedMathTypeEditor(
                application,
                format,
                replacementLatex: null,
                saveChanges: false);
            var sourceLatex = MathMlToLatexConverter.Convert(sourceMathMl).Replace(" ", string.Empty);
            Console.WriteLine("  source LaTeX=" + sourceLatex);
            if (string.IsNullOrWhiteSpace(nativeCaseFilter))
            {
                AssertTrue(
                    sourceLatex.IndexOf("sqrt", StringComparison.OrdinalIgnoreCase) >= 0
                    && (sourceLatex.IndexOf("p^2", StringComparison.Ordinal) >= 0
                        || sourceLatex.IndexOf("p^{2}", StringComparison.Ordinal) >= 0)
                    && (sourceLatex.IndexOf("q^2", StringComparison.Ordinal) >= 0
                        || sourceLatex.IndexOf("q^{2}", StringComparison.Ordinal) >= 0),
                    $"Word-owned MathType editor exposed the wrong source. LaTeX='{sourceLatex}'.");
            }

            Console.WriteLine("[MathType native editor 2/4] Reopening the same OLE, replacing its contents, and saving back into Word...");
            const string editedLatex = @"\frac{x+1}{y}";
            var editedMathMl = InvokeWordOwnedMathTypeEditor(
                application,
                format,
                replacementLatex: editedLatex,
                saveChanges: true);
            var editorLatex = MathMlToLatexConverter.Convert(editedMathMl).Replace(" ", string.Empty);
            Console.WriteLine("  editor LaTeX before save=" + editorLatex);
            AssertTrue(
                editorLatex.IndexOf(@"\frac{x+1}{y}", StringComparison.Ordinal) >= 0,
                $"MathType did not contain the requested edit before saving. LaTeX='{editorLatex}'.");

            Release(format);
            format = null;
            format = shape.OLEFormat;
            AssertTrue(MathTypeOleInterop.TryResolveCapabilities(format.ProgID, out _),
                $"Saving through MathType changed the Word OLE class to '{format.ProgID}'.");

            Console.WriteLine("[MathType native editor 3/4] Saving and reopening Word, then asking MathType to expose the same object again...");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(sourcePath, ReadOnly: false, Visible: true);
            document.Activate();
            Release(shape);
            shape = document.InlineShapes[1];
            Release(format);
            format = shape.OLEFormat;
            var reopenedMathMl = InvokeWordOwnedMathTypeEditor(
                application,
                format,
                replacementLatex: null,
                saveChanges: false);
            var reopenedLatex = MathMlToLatexConverter.Convert(reopenedMathMl).Replace(" ", string.Empty);
            Console.WriteLine("  reopened LaTeX=" + reopenedLatex);
            AssertTrue(
                reopenedLatex.IndexOf(@"\frac{x+1}{y}", StringComparison.Ordinal) >= 0,
                $"Word save/reopen lost the MathType native-editor update. LaTeX='{reopenedLatex}'.");

            Console.WriteLine("[MathType native editor 4/5] Saving several genuine MathType 7 complex structures, then reading their persisted MTEF directly through VisualTeX...");
            const string symbolMatrixMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow>"
                + "<mi>α</mi><mi>β</mi><mi>γ</mi><mi>δ</mi><mi>θ</mi><mi>λ</mi><mi>μ</mi><mi>π</mi><mi>ρ</mi><mi>σ</mi><mi>φ</mi><mi>ω</mi>"
                + "<mi>Γ</mi><mi>Δ</mi><mi>Θ</mi><mi>Λ</mi><mi>Π</mi><mi>Σ</mi><mi>Φ</mi><mi>Ω</mi>"
                + "<mo>−</mo><mo>±</mo><mo>∓</mo><mo>×</mo><mo>·</mo><mo>÷</mo><mo>∞</mo><mo>∂</mo><mo>∇</mo>"
                + "<mo>†</mo><mo>‡</mo><mo>′</mo><mo>″</mo><mo>∀</mo><mo>∃</mo><mo>∈</mo><mo>∉</mo>"
                + "<mo>⊂</mo><mo>⊆</mo><mo>⊃</mo><mo>⊇</mo><mo>∪</mo><mo>∩</mo>"
                + "<mo>→</mo><mo>←</mo><mo>↔</mo><mo>⇒</mo><mo>⇔</mo><mo>↦</mo>"
                + "<mo>⟨</mo><mo>⟩</mo><mo>∗</mo><mo>|</mo><mo>∥</mo>"
                + "<mo>≠</mo><mo>≤</mo><mo>≥</mo><mo>≈</mo><mo>≡</mo><mi>ℏ</mi>"
                + "</mrow></math>";
            var genuineCases = new[]
            {
                (Name: "symbol-matrix", Latex: string.Empty),
                (Name: "sum", Latex: @"\sum_{i=1}^{n} a_i"),
                (Name: "integral", Latex: @"\int_{0}^{1} x^2\,dx"),
                (Name: "vector", Latex: @"\vec{v}"),
                (Name: "hbar", Latex: @"\hbar"),
                (Name: "rho", Latex: @"\rho"),
                (Name: "dagger", Latex: @"\dagger"),
                (Name: "prime", Latex: @"x'"),
                (Name: "double-prime", Latex: @"x''"),
                (Name: "minus", Latex: @"a-b"),
                (Name: "langle", Latex: @"\langle f\rangle"),
                (Name: "forall", Latex: @"\forall"),
                (Name: "ast", Latex: @"f^*"),
                (Name: "mid", Latex: @"a\mid b"),
                (Name: "bra-ket", Latex: @"\langle f|L|g\rangle"),
                (Name: "bigbar", Latex: @"Q\big|_a^b"),
                (Name: "overline", Latex: @"\overline{AB}"),
                (Name: "matrix", Latex: @"\left|\begin{matrix}a&b\\c&d\end{matrix}\right|"),
                (Name: "bmatrix", Latex: @"\begin{bmatrix}a&b\\c&d\end{bmatrix}"),
                (Name: "mathbb", Latex: @"\mathbb{R}"),
                (Name: "mathcal", Latex: @"\mathcal{F}"),
                (Name: "mathfrak", Latex: @"\mathfrak{g}"),
                (Name: "mathbf", Latex: @"\mathbf{v}"),
                (Name: "underbrace", Latex: @"\underbrace{a+b}_{n}"),
                (Name: "overbrace", Latex: @"\overbrace{a+b}^{n}"),
                (Name: "bigcup", Latex: @"\bigcup_{i=1}^{n} A_i"),
                (Name: "bigcap", Latex: @"\bigcap_{i=1}^{n} A_i"),
                (Name: "product", Latex: @"\prod_{i=1}^{n} a_i"),
                (Name: "coproduct", Latex: @"\coprod_{i=1}^{n} a_i"),
                (Name: "max", Latex: @"\max_{x\in A} f(x)"),
                (Name: "iiint", Latex: @"\iiint_{V} f\,dV"),
                (Name: "inline-mixed", Latex: @"\frac{a}{b}e^{i\pi}+1=0"),
            };
            var nativeCases = genuineCases.AsEnumerable();
            if (!string.IsNullOrWhiteSpace(nativeCaseFilter))
            {
                nativeCases = nativeCases.Where(item => string.Equals(
                    item.Name,
                    nativeCaseFilter,
                    StringComparison.OrdinalIgnoreCase));
                Console.WriteLine($"  MathType native-case filter={nativeCaseFilter}.");
            }
            foreach (var genuineCase in nativeCases)
            {
                var mathTypeMathMl = string.Equals(genuineCase.Name, "mathbb", StringComparison.Ordinal)
                    ? InvokeWordOwnedMathTypeEditor(
                        application,
                        format,
                        replacementLatex: null,
                        saveChanges: true,
                        replacementMathMl: null,
                        replacementKeySequence: "blackboard-R")
                    : string.Equals(genuineCase.Name, "symbol-matrix", StringComparison.Ordinal)
                        ? InvokeWordOwnedMathTypeEditor(
                            application,
                            format,
                            replacementLatex: null,
                            saveChanges: true,
                            replacementMathMl: symbolMatrixMathMl)
                        : InvokeWordOwnedMathTypeEditor(
                            application,
                            format,
                            replacementLatex: genuineCase.Latex,
                            saveChanges: true);
                Release(format);
                format = null;
                document.Save();
                Release(shape);
                shape = document.InlineShapes[1];
                format = shape.OLEFormat;
                var nativeObjectPosition = ReadMathTypeObjectCharacterPosition(shape);
                Console.WriteLine(
                    $"  {genuineCase.Name}: genuine MathType geometry={shape.Width:0.00}x{shape.Height:0.00}pt, objectPosition={nativeObjectPosition}.");
                var compoundFile = MathTypeOleStorage.CaptureCompoundFile(shape);
                var equationNative = MathTypeOleStorage.ReadEquationNative(compoundFile);
                File.WriteAllBytes(
                    Path.Combine(artifactRoot, $"genuine-{genuineCase.Name}-compound.cfb"),
                    compoundFile);
                File.WriteAllBytes(
                    Path.Combine(artifactRoot, $"genuine-{genuineCase.Name}-equation-native.bin"),
                    equationNative);
                var nativeMtefLength = checked((int)BitConverter.ToUInt32(equationNative, 8));
                var nativeMtef = new byte[nativeMtefLength];
                Buffer.BlockCopy(equationNative, 28, nativeMtef, 0, nativeMtef.Length);
                Console.WriteLine(
                    $"  {genuineCase.Name}: genuine MathType root offset={MathTypeMtefCodec.FindRootStructureOffset(nativeMtef)}.");
                var directMathMl = MathTypeMtefCodec.ReadEquationNativeMathMl(equationNative);
                if (!string.Equals(genuineCase.Name, "symbol-matrix", StringComparison.Ordinal))
                {
                    AssertEqual(
                        MathTypeMtefCodec.SemanticSignature(mathTypeMathMl),
                        MathTypeMtefCodec.SemanticSignature(directMathMl),
                        $"VisualTeX direct MTEF parser disagreed with MathType 7 for genuine '{genuineCase.Name}' equation. MathType='{mathTypeMathMl}', direct='{directMathMl}'.");
                }
                if (string.Equals(genuineCase.Name, "mathbb", StringComparison.Ordinal)
                    || string.Equals(genuineCase.Name, "hbar", StringComparison.Ordinal))
                {
                    var rewrittenNative = MathTypeMtefCodec.RewriteEquationNative(
                        equationNative,
                        mathTypeMathMl,
                        inline: true).EquationNative;
                    if (!equationNative.SequenceEqual(rewrittenNative))
                    {
                        File.WriteAllBytes(
                            Path.Combine(artifactRoot, $"genuine-{genuineCase.Name}-rewritten-by-visualtex.bin"),
                            rewrittenNative);
                        var compareLength = Math.Min(equationNative.Length, rewrittenNative.Length);
                        var firstDifference = Enumerable.Range(0, compareLength)
                            .FirstOrDefault(index => equationNative[index] != rewrittenNative[index]);
                        throw new InvalidDataException(
                            $"VisualTeX rewrote genuine MathType 7 '{genuineCase.Name}' to different Equation Native bytes. "
                            + $"originalLength={equationNative.Length}, rewrittenLength={rewrittenNative.Length}, "
                            + $"firstDifference={firstDifference}, original=0x{equationNative[firstDifference]:X2}, rewritten=0x{rewrittenNative[firstDifference]:X2}.");
                    }
                    Console.WriteLine(
                        $"  {genuineCase.Name}: VisualTeX rewrite is byte-for-byte identical to genuine MathType 7 Equation Native.");
                }
                Console.WriteLine(
                    $"  {genuineCase.Name}: direct VisualTeX MTEF source={MathMlToLatexConverter.Convert(directMathMl)}.");
                var typographySensitiveCases = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "inline-mixed",
                    "integral",
                    "iiint",
                    "sum",
                    "product",
                    "coproduct",
                    "bigcup",
                    "bigcap",
                    "max",
                };
                if (typographySensitiveCases.Contains(genuineCase.Name))
                {
                    var originalLength = checked((int)BitConverter.ToUInt32(equationNative, 8));
                    var originalMtef = new byte[originalLength];
                    Buffer.BlockCopy(equationNative, 28, originalMtef, 0, originalLength);
                    var visualTexRewrite = MathTypeMtefCodec.RewriteEquationNative(
                        equationNative,
                        mathTypeMathMl,
                        inline: true);
                    if (!MathTypeNativePreviewRenderer.TryRender(
                            originalMtef,
                            artifactRoot,
                            out var genuinePreview)
                        || !MathTypeNativePreviewRenderer.TryRender(
                            visualTexRewrite.Mtef,
                            artifactRoot,
                            out var visualTexPreview))
                        throw new InvalidDataException(
                            "MathType native renderer was unavailable for the typography regression.");
                    using (genuinePreview)
                    using (visualTexPreview)
                    {
                        var genuineWmf = File.ReadAllBytes(genuinePreview.WmfPath);
                        var visualTexWmf = File.ReadAllBytes(visualTexPreview.WmfPath);
                        var pixelDifference = MeasureEmfPixelDifference(genuineWmf, visualTexWmf);
                        Console.WriteLine(
                            $"  {genuineCase.Name} native typography: genuine={genuinePreview.WidthPt:0.00}x{genuinePreview.HeightPt:0.00}pt pos={genuinePreview.WordPosition}, "
                            + $"VisualTeX={visualTexPreview.WidthPt:0.00}x{visualTexPreview.HeightPt:0.00}pt pos={visualTexPreview.WordPosition}, diff={pixelDifference:0.0000}.");
                        AssertNear(
                            genuinePreview.HeightPt,
                            visualTexPreview.HeightPt,
                            0.5f,
                            $"VisualTeX changed MathType's native full-size equation height for {genuineCase.Name}.");
                        AssertNear(
                            genuinePreview.WordPosition,
                            visualTexPreview.WordPosition,
                            1.0f,
                            $"VisualTeX changed MathType's native inline baseline for {genuineCase.Name}.");
                        var requiresExactNativeTypography = genuineCase.Name is
                            "inline-mixed" or "sum" or "product" or "coproduct" or "bigcup" or "bigcap";
                        AssertTrue(
                            pixelDifference < (requiresExactNativeTypography ? 0.03 : 0.08),
                            $"VisualTeX changed MathType's internal font or relative limit/script sizing for {genuineCase.Name}.");
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(nativeCaseFilter))
            {
                Console.WriteLine("[MathType native editor 5/5] Filtered native MathType case completed.");
                return;
            }

            const string blackboardMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi mathvariant=\"double-struck\">R</mi></math>";
            var genuineBlackboardMathMl = InvokeWordOwnedMathTypeEditor(
                application,
                format,
                replacementLatex: null,
                saveChanges: true,
                replacementMathMl: blackboardMathMl);
            Release(format);
            format = null;
            document.Save();
            Release(shape);
            shape = document.InlineShapes[1];
            format = shape.OLEFormat;
            var directBlackboardMathMl = MathTypeOleStorage.ReadMathMl(shape);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(genuineBlackboardMathMl),
                MathTypeMtefCodec.SemanticSignature(directBlackboardMathMl),
                $"VisualTeX direct MTEF parser disagreed with MathType 7 for genuine MathML blackboard R. MathType='{genuineBlackboardMathMl}', direct='{directBlackboardMathMl}'.");
            Console.WriteLine(
                $"  mathbb-mathml: direct VisualTeX MTEF source={MathMlToLatexConverter.Convert(directBlackboardMathMl)}.");

            Console.WriteLine("[MathType native editor 5/5] Word-owned MathType edit and genuine complex-MTEF direct-read coverage passed without any standalone Untitled MathType document.");
        }
        finally
        {
            TryRestoreClipboard(clipboardBackup);
            Release(format);
            Release(shape);
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

    private static int ReadMathTypeObjectCharacterPosition(Word.InlineShape shape)
    {
        Word.Range? shapeRange = null;
        Word.Range? probe = null;
        Word.Font? font = null;
        Word.Document? document = null;
        try
        {
            shapeRange = shape.Range;
            document = shapeRange.Document;
            for (var position = shapeRange.Start; position < shapeRange.End; position++)
            {
                Release(font);
                font = null;
                Release(probe);
                probe = document.Range(position, position + 1);
                if (!string.Equals(probe.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                font = probe.Font;
                return font.Position;
            }
            return (int)Word.WdConstants.wdUndefined;
        }
        finally
        {
            Release(font);
            Release(probe);
            Release(shapeRange);
            Release(document);
        }
    }

    private static void HandleKnownMathTypeModalDialogs(
        int processId,
        IntPtr editorWindow,
        bool saveChanges)
    {
        string? unknownDialog = null;
        EnumWindowsForMathType((dialog, _) =>
        {
            GetWindowThreadProcessId(dialog, out var ownerPid);
            if (ownerPid != (uint)processId) return true;
            var className = new System.Text.StringBuilder(64);
            GetClassNameForMathType(dialog, className, className.Capacity);
            if (!string.Equals(className.ToString(), "#32770", StringComparison.Ordinal))
                return true;

            var owner = GetWindowOwnerForMathType(dialog, 4); // GW_OWNER
            if (owner != IntPtr.Zero && owner != editorWindow) return true;
            var title = GetWindowTitleForMathType(dialog);
            var text = GetDialogStaticTextForMathType(dialog);

            // Never confirm MathType's warning about being forcibly exited while
            // servicing Word. Seeing this dialog means the acceptance used the
            // wrong lifecycle command and must fail instead of dismissing it.
            if ((text.IndexOf("另一个应用", StringComparison.OrdinalIgnoreCase) >= 0
                    || text.IndexOf("another application", StringComparison.OrdinalIgnoreCase) >= 0)
                && (text.IndexOf("退出", StringComparison.OrdinalIgnoreCase) >= 0
                    || text.IndexOf("exit", StringComparison.OrdinalIgnoreCase) >= 0))
            {
                unknownDialog = $"MathType reported a forced-exit-while-servicing-Word warning. Title='{title}', Text='{text}'.";
                return false;
            }

            var isTeachingTip =
                title.IndexOf("教学说明", StringComparison.OrdinalIgnoreCase) >= 0
                || (text.IndexOf("MathType符号工具栏", StringComparison.OrdinalIgnoreCase) >= 0
                    && text.IndexOf("预置", StringComparison.OrdinalIgnoreCase) >= 0);
            if (isTeachingTip)
            {
                var okButton = GetDlgItemForMathType(dialog, 1); // IDOK
                if (okButton != IntPtr.Zero)
                {
                    SendMessageForMathType(okButton, 0x00F5, IntPtr.Zero, IntPtr.Zero); // BM_CLICK
                    Console.WriteLine("  dismissed MathType teaching-tip dialog.");
                    return true;
                }
            }

            var isSavePrompt =
                text.IndexOf("保存", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("save", StringComparison.OrdinalIgnoreCase) >= 0;
            if (isSavePrompt)
            {
                var buttonId = saveChanges ? 6 : 7; // IDYES / IDNO
                var button = GetDlgItemForMathType(dialog, buttonId);
                if (button != IntPtr.Zero)
                {
                    SendMessageForMathType(button, 0x00F5, IntPtr.Zero, IntPtr.Zero); // BM_CLICK
                    Console.WriteLine(saveChanges
                        ? "  confirmed MathType save-back into the Word OLE object."
                        : "  closed MathType read-only session without saving.");
                    return true;
                }
                // Some MathType warnings mention saving in explanatory text but
                // are not save-confirmation dialogs. Fall through to the unknown
                // diagnostic instead of clicking an unrelated button.
            }

            unknownDialog = $"Unknown MathType modal dialog '{title}'. Text='{text}'. "
                + $"TitleBase64={Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(title))}; "
                + $"TextBase64={Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(text))}.";
            return false;
        }, IntPtr.Zero);

        if (!string.IsNullOrWhiteSpace(unknownDialog))
            throw new InvalidOperationException(unknownDialog);
    }

    private static string CaptureMathMlFromActiveMathTypeEditor(
        IntPtr editorWindow,
        int processId,
        bool saveChanges,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        Exception? lastError = null;
        while (DateTime.UtcNow < deadline)
        {
            HandleKnownMathTypeModalDialogs(processId, editorWindow, saveChanges);
            if (!IsWindowForMathType(editorWindow))
                throw new InvalidOperationException(
                    "The Word-owned MathType editor closed before its equation became readable.");
            SetForegroundWindow(editorWindow);
            System.Windows.Forms.SendKeys.SendWait("^a");
            Thread.Sleep(80);
            System.Windows.Forms.SendKeys.SendWait("^c");
            Thread.Sleep(220);
            try
            {
                var data = System.Windows.Forms.Clipboard.GetDataObject();
                if (data is not null)
                    return ReadMathMlFromClipboardDataObject(data);
            }
            catch (Exception error) when (
                error is InvalidDataException
                or System.Runtime.InteropServices.ExternalException)
            {
                lastError = error;
            }
            Thread.Sleep(120);
        }
        throw new InvalidDataException(
            "MathType OLE editor did not expose MathML before the readiness timeout.",
            lastError);
    }

    private static void PutTeXOnMathTypeClipboard(string latex)
    {
        var customFormat = RegisterClipboardFormatForMathType("TeX Input Language");
        if (customFormat == 0)
            throw new InvalidOperationException(
                "Windows could not register MathType's TeX Input Language clipboard format.");
        var payloads = new[]
        {
            (Format: customFormat, Bytes: System.Text.Encoding.ASCII.GetBytes(latex + "\0")),
            (Format: 1u, Bytes: System.Text.Encoding.Default.GetBytes(latex + "\0")), // CF_TEXT
            (Format: 13u, Bytes: System.Text.Encoding.Unicode.GetBytes(latex + "\0")), // CF_UNICODETEXT
        };
        var allocated = new List<(uint Format, IntPtr Memory, bool Transferred)>();
        try
        {
            foreach (var payload in payloads)
                allocated.Add((
                    payload.Format,
                    AllocateMathTypeClipboardPayload(payload.Bytes),
                    false));

            var opened = false;
            for (var attempt = 0; attempt < 20 && !opened; attempt++)
            {
                opened = OpenClipboardForMathType(IntPtr.Zero);
                if (!opened) Thread.Sleep(25);
            }
            if (!opened)
                throw new InvalidOperationException(
                    "Unable to open the Windows clipboard for MathType TeX input.");
            try
            {
                if (!EmptyClipboardForMathType())
                    throw new InvalidOperationException(
                        "Unable to clear the Windows clipboard for MathType TeX input.");
                for (var index = 0; index < allocated.Count; index++)
                {
                    var item = allocated[index];
                    if (SetClipboardDataForMathType(item.Format, item.Memory) == IntPtr.Zero)
                        throw new InvalidOperationException(
                            $"Unable to place MathType TeX clipboard format {item.Format}.");
                    allocated[index] = (item.Format, item.Memory, true);
                }
            }
            finally { CloseClipboardForMathType(); }
        }
        finally
        {
            foreach (var item in allocated)
            {
                if (!item.Transferred && item.Memory != IntPtr.Zero)
                    GlobalFreeForMathType(item.Memory);
            }
        }
    }

    private static string InvokeWordOwnedMathTypeEditor(
        Word.Application application,
        Word.OLEFormat format,
        string? replacementLatex,
        bool saveChanges,
        string? replacementMathMl = null,
        string? replacementKeySequence = null)
    {
        Word.Window? wordWindow = null;
        Exception? driverError = null;
        string? copiedMathMl = null;
        using var driverFinished = new ManualResetEventSlim(false);
        var existingMathTypeWindows = GetMathTypeTopLevelWindows();
        try
        {
            wordWindow = application.ActiveWindow;
            var wordWindowHandle = new IntPtr(wordWindow.Hwnd);
            SetForegroundWindow(wordWindowHandle);
            Thread.Sleep(180);

            var driver = new Thread(() =>
            {
                try
                {
                    var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(15);
                    IntPtr editorWindow = IntPtr.Zero;
                    uint processId = 0;
                    while (DateTime.UtcNow < deadline)
                    {
                        var foreground = GetForegroundWindowForAcceptance();
                        if (foreground != IntPtr.Zero && foreground != wordWindowHandle)
                        {
                            GetWindowThreadProcessId(foreground, out var candidatePid);
                            if (candidatePid != 0)
                            {
                                try
                                {
                                    using var process = System.Diagnostics.Process.GetProcessById((int)candidatePid);
                                    if (string.Equals(
                                            process.ProcessName,
                                            "MathType",
                                            StringComparison.OrdinalIgnoreCase))
                                    {
                                        editorWindow = foreground;
                                        processId = candidatePid;
                                        break;
                                    }
                                }
                                catch { }
                            }
                        }

                        foreach (var candidate in GetMathTypeTopLevelWindows())
                        {
                            if (existingMathTypeWindows.Contains(candidate)) continue;
                            GetWindowThreadProcessId(candidate, out var candidatePid);
                            if (candidatePid == 0) continue;
                            editorWindow = candidate;
                            processId = candidatePid;
                            break;
                        }
                        if (editorWindow != IntPtr.Zero) break;
                        Thread.Sleep(60);
                    }
                    if (editorWindow == IntPtr.Zero || processId == 0)
                        throw new TimeoutException(
                            "MathType did not expose the Word-owned OLE editor window.");

                    SetForegroundWindow(editorWindow);
                    Thread.Sleep(300);
                    if (!string.IsNullOrWhiteSpace(replacementLatex)
                        || !string.IsNullOrWhiteSpace(replacementMathMl)
                        || !string.IsNullOrWhiteSpace(replacementKeySequence))
                    {
                        System.Windows.Forms.SendKeys.SendWait("^a");
                        Thread.Sleep(80);
                        System.Windows.Forms.SendKeys.SendWait("{DELETE}");
                        Thread.Sleep(100);
                        if (string.Equals(replacementKeySequence, "blackboard-R", StringComparison.Ordinal))
                        {
                            // MathType 7's own shortcut for blackboard-bold capital R.
                            // Using the editor's native command here gives us a canonical
                            // Equation Native/MTEF sample instead of guessing Euclid Math Two codes.
                            System.Windows.Forms.SendKeys.SendWait("^d");
                            Thread.Sleep(120);
                            System.Windows.Forms.SendKeys.SendWait("+r");
                            Thread.Sleep(500);
                        }
                        else
                        {
                            if (!string.IsNullOrWhiteSpace(replacementMathMl))
                                PutMathMlOnNativeClipboard(replacementMathMl!);
                            else
                                PutTeXOnMathTypeClipboard(replacementLatex!);
                            System.Windows.Forms.SendKeys.SendWait("^v");
                            Thread.Sleep(500);
                        }
                    }

                    copiedMathMl = CaptureMathMlFromActiveMathTypeEditor(
                        editorWindow,
                        (int)processId,
                        saveChanges,
                        TimeSpan.FromSeconds(4));

                    // Ctrl+F4 closes the current equation/document. Alt+F4 exits
                    // the MathType application and is invalid while it is acting
                    // as Word's OLE server.
                    System.Windows.Forms.SendKeys.SendWait("^{F4}");
                    var closeDeadline = DateTime.UtcNow + TimeSpan.FromSeconds(8);
                    while (DateTime.UtcNow < closeDeadline)
                    {
                        HandleKnownMathTypeModalDialogs(
                            (int)processId,
                            editorWindow,
                            saveChanges);
                        if (!IsWindowForMathType(editorWindow)) break;
                        Thread.Sleep(80);
                    }
                }
                catch (Exception error)
                {
                    driverError = error;
                }
                finally { driverFinished.Set(); }
            })
            {
                IsBackground = true,
                Name = "VisualTeX Word-owned MathType OLE acceptance driver",
            };
            driver.SetApartmentState(ApartmentState.STA);
            driver.Start();

            object openVerb = (int)Word.WdOLEVerb.wdOLEVerbOpen;
            format.DoVerb(ref openVerb);
            if (!driverFinished.Wait(TimeSpan.FromSeconds(15)))
                throw new TimeoutException(
                    "MathType OLE editor returned to Word before the driver completed within the safety timeout.");
            if (driverError is not null)
                throw new InvalidOperationException(
                    "MathType Word-owned OLE editor automation failed.", driverError);
            if (string.IsNullOrWhiteSpace(copiedMathMl))
                throw new InvalidDataException(
                    "MathType Word-owned OLE editor returned no readable MathML.");

            SetForegroundWindow(wordWindowHandle);
            Thread.Sleep(300);
            return copiedMathMl!;
        }
        finally { Release(wordWindow); }
    }
}
