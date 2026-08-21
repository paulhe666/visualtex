using System.Diagnostics;
using System.IO.Compression;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeStandaloneCodecAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        string M(string body) =>
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" + body + "</math>";
        var mathMl = M(
            "<mi>\u03C1</mi><mo>=</mo><mn>1</mn><mo>;</mo>"
            + "<msup><mi>L</mi><mo>\u2020</mo></msup><mo>;</mo>"
            + "<msubsup><mi>p</mi><mn>2</mn><mo>\u2032</mo></msubsup><mo>;</mo>"
            + "<msubsup><mi>p</mi><mn>2</mn><mo>\u2033</mo></msubsup><mo>;</mo>"
            + "<mo>\u27E8</mo><mi>f</mi><mo>\u27E9</mo><mo>;</mo>"
            + "<mo>\u2200</mo><mo>\u27E8</mo><mi>f</mi><mo>|</mo><mi>L</mi><mo>|</mo><mi>g</mi><mo>\u27E9</mo>"
            + "<mo>\u2212</mo><mo>\u27E8</mo><mi>g</mi><mo>|</mo><msup><mi>L</mi><mo>\u2020</mo></msup><mo>|</mo><mi>f</mi><mo>\u27E9</mo>"
            + "<mo>=</mo><mi>Q</mi><mo>[</mo><msup><mi>f</mi><mo>\u2217</mo></msup><mo>,</mo><mi>g</mi><mo>]</mo>"
            + "<msubsup><mo>|</mo><mi>a</mi><mi>b</mi></msubsup><mo>;</mo>"
            + "<mi>a</mi><mo>\u2223</mo><mi>b</mi><mo>;</mo><mi>\u210F</mi><mo>;</mo><mi>L</mi><mo>\u2261</mo><mi>R</mi>");

        var before = SnapshotMathTypeProcessIds();
        var generated = MathTypeMtefCodec.CreateEquationNative(mathMl, inline: false);
        var compound = MathTypeOleStorage.CreateStandaloneCompoundFile(generated);
        var readBack = MathTypeOleStorage.ReadMathMl(compound);
        var after = SnapshotMathTypeProcessIds();
        var started = after.Except(before).ToArray();
        if (started.Length > 0)
            throw new InvalidDataException(
                "Standalone VisualTeX MTEF generation unexpectedly started MathType process(es): "
                + string.Join(", ", started));
        AssertEqual(
            MathTypeMtefCodec.SemanticSignature(mathMl),
            MathTypeMtefCodec.SemanticSignature(readBack),
            $"Standalone VisualTeX MTEF semantic mismatch. actual='{readBack}'");
        AssertTrue(MathTypeOleStorage.LooksLikeMathTypeCompoundFile(compound),
            "Standalone VisualTeX CFB does not look like Equation.DSMT4 storage.");

        static bool Contains(byte[] data, params byte[] pattern)
        {
            if (pattern.Length == 0 || data.Length < pattern.Length) return false;
            for (var start = 0; start <= data.Length - pattern.Length; start++)
            {
                var match = true;
                for (var offset = 0; offset < pattern.Length; offset++)
                {
                    if (data[start + offset] == pattern[offset]) continue;
                    match = false;
                    break;
                }
                if (match) return true;
            }
            return false;
        }

        var mtef = generated.Mtef;
        AssertTrue(Contains(mtef, 0x02, 0x04, 0x84, 0xC1, 0x03, 0x72),
            "rho was not emitted as LowerGreek/U+03C1/Adobe-Symbol 0x72.");
        AssertTrue(Contains(mtef, 0x02, 0x04, 0x86, 0x32, 0x20, 0xA2),
            "prime was not emitted as Symbol/U+2032/Adobe-Symbol 0xA2.");
        AssertTrue(Contains(mtef, 0x02, 0x04, 0x86, 0x33, 0x20, 0xB2),
            "double-prime was not emitted as Symbol/U+2033/Adobe-Symbol 0xB2.");
        AssertTrue(Contains(mtef, 0x02, 0x04, 0x86, 0x12, 0x22, 0x2D),
            "minus was not emitted as Symbol/U+2212/Adobe-Symbol 0x2D.");
        AssertTrue(Contains(mtef, 0x02, 0x04, 0x86, 0x00, 0x22, 0x22),
            "forall was not emitted as Symbol/U+2200/Adobe-Symbol 0x22.");
        AssertTrue(Contains(mtef, 0x02, 0x04, 0x86, 0x61, 0x22, 0xBA),
            "equivalence was not emitted as Symbol/U+2261/Adobe-Symbol 0xBA.");
        AssertTrue(Contains(mtef, 0x02, 0x04, 0x86, 0x29, 0x23, 0xE1)
                   && Contains(mtef, 0x02, 0x04, 0x86, 0x2A, 0x23, 0xF1),
            "angle brackets were not emitted with MathType's historical MTCode and Adobe-Symbol positions.");
        AssertTrue(Contains(mtef, 0x02, 0x04, 0x8B, 0x0F, 0x21, 0x68),
            "hbar was not emitted as MT Extra/U+210F/0x68.");
        AssertTrue(Contains(mtef, 0x02, 0x00, 0x81, 0x20, 0x20),
            "dagger was not emitted as a Unicode-capable non-Symbol CHAR.");

        File.WriteAllBytes(Path.Combine(artifactRoot, "visualtex-standalone-doc4-equation-native.bin"), generated.EquationNative);
        File.WriteAllBytes(Path.Combine(artifactRoot, "visualtex-standalone-doc4.cfb"), compound);
        File.WriteAllText(Path.Combine(artifactRoot, "visualtex-standalone-doc4-readback.xml"), readBack);

        var workspaceArtifacts = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", "..", "..",
            "artifacts"));
        var genuineInlineMixedPath = Path.Combine(
            workspaceArtifacts,
            "mathtype-inline-mixed-native",
            "genuine-inline-mixed-equation-native.bin");
        if (File.Exists(genuineInlineMixedPath))
        {
            var genuineNative = File.ReadAllBytes(genuineInlineMixedPath);
            var genuineMathMl = MathTypeMtefCodec.ReadEquationNativeMathMl(genuineNative);
            var standaloneInline = MathTypeMtefCodec.CreateEquationNative(
                genuineMathMl,
                inline: true);
            var genuineMtefLength = checked((int)BitConverter.ToUInt32(genuineNative, 8));
            var genuineMtef = new byte[genuineMtefLength];
            Buffer.BlockCopy(genuineNative, 28, genuineMtef, 0, genuineMtefLength);
            if (MathTypeNativePreviewRenderer.TryRender(
                    genuineMtef,
                    artifactRoot,
                    out var genuinePreview)
                && MathTypeNativePreviewRenderer.TryRender(
                    standaloneInline.Mtef,
                    artifactRoot,
                    out var standalonePreview))
            {
                using (genuinePreview)
                using (standalonePreview)
                {
                    var genuineWmf = File.ReadAllBytes(genuinePreview.WmfPath);
                    var standaloneWmf = File.ReadAllBytes(standalonePreview.WmfPath);
                    var pixelDifference = MeasureEmfPixelDifference(genuineWmf, standaloneWmf);
                    Console.WriteLine(
                        $"[STANDALONE VS GENUINE] inline-mixed genuine={genuinePreview.WidthPt:0.##}x{genuinePreview.HeightPt:0.##} pos={genuinePreview.WordPosition}; "
                        + $"standalone={standalonePreview.WidthPt:0.##}x{standalonePreview.HeightPt:0.##} pos={standalonePreview.WordPosition}; diff={pixelDifference:0.0000}; "
                        + $"genuineRoot={MathTypeMtefCodec.FindRootStructureOffset(genuineMtef)} standaloneRoot={MathTypeMtefCodec.FindRootStructureOffset(standaloneInline.Mtef)}.");
                    File.WriteAllBytes(
                        Path.Combine(artifactRoot, "visualtex-standalone-inline-mixed-equation-native.bin"),
                        standaloneInline.EquationNative);
                }
            }
        }

        Console.WriteLine(
            "[STANDALONE MTEF] VisualTeX created Equation Native + CFB from scratch, "
            + "round-tripped the doc4 symbol family, matched all representative CHAR encodings, "
            + "and started no MathType process.");
    }

    private sealed class MathTypeOfflineCase
    {
        public string Name { get; set; } = string.Empty;
        public string MathMl { get; set; } = string.Empty;
        public bool ValidateWithMathType { get; set; }
        public bool Inline { get; set; } = true;
    }

    private static void RunWordMathTypeOfflineStorageAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var workspaceArtifacts = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", "..", "..",
            "artifacts"));
        var source = Path.Combine(
            workspaceArtifacts,
            "mathtype-native-editor",
            "VisualTeX-MathType7-NativeEditor-5f04f8b3545e444a824705446e314ba1.docx");
        if (!File.Exists(source))
            throw new FileNotFoundException(
                "The synchronized MathType 7 source fixture is missing.", source);

        string M(string body) =>
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" + body + "</math>";
        var cases = new[]
        {
            new MathTypeOfflineCase { Name = "a-plus-b", MathMl = M("<mi>a</mi><mo>+</mo><mi>b</mi>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "hbar", MathMl = M("<mi>ℏ</mi>"), ValidateWithMathType = true },
            new MathTypeOfflineCase
            {
                Name = "doc4-operator-symbol-family",
                MathMl = M(
                    "<mi>\u03C1</mi><mo>=</mo><mn>1</mn><mo>;</mo>"
                    + "<msup><mi>L</mi><mo>\u2020</mo></msup><mo>;</mo>"
                    + "<msubsup><mi>p</mi><mn>2</mn><mo>\u2032</mo></msubsup><mo>;</mo>"
                    + "<msubsup><mi>p</mi><mn>2</mn><mo>\u2033</mo></msubsup><mo>;</mo>"
                    + "<mo>\u27E8</mo><mi>f</mi><mo>\u27E9</mo><mo>;</mo>"
                    + "<mo>\u2200</mo><mo>\u27E8</mo><mi>f</mi><mo>|</mo><mi>L</mi><mo>|</mo><mi>g</mi><mo>\u27E9</mo>"
                    + "<mo>\u2212</mo><mo>\u27E8</mo><mi>g</mi><mo>|</mo><msup><mi>L</mi><mo>\u2020</mo></msup><mo>|</mo><mi>f</mi><mo>\u27E9</mo>"
                    + "<mo>=</mo><mi>Q</mi><mo>[</mo><msup><mi>f</mi><mo>\u2217</mo></msup><mo>,</mo><mi>g</mi><mo>]</mo>"
                    + "<msubsup><mo>|</mo><mi>a</mi><mi>b</mi></msubsup><mo>;</mo>"
                    + "<mi>a</mi><mo>\u2223</mo><mi>b</mi><mo>;</mo><mi>\u210F</mi><mo>;</mo><mi>L</mi><mo>\u2261</mo><mi>R</mi>"),
            },
            new MathTypeOfflineCase
            {
                Name = "upright-euler",
                MathMl = M(
                    "<mi>π</mi><mi>θ</mi><msup><mi mathvariant=\"normal\">e</mi>"
                    + "<mrow><mi mathvariant=\"normal\">i</mi><mi>π</mi></mrow></msup>"
                    + "<mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn>"),
                ValidateWithMathType = true,
            },
            new MathTypeOfflineCase
            {
                Name = "fraction",
                MathMl = M("<mfrac><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow><mi>y</mi></mfrac>"),
                ValidateWithMathType = true,
            },
            new MathTypeOfflineCase
            {
                Name = "nested-fraction",
                MathMl = M("<mfrac><mfrac><mi>a</mi><mi>b</mi></mfrac><mfrac><mi>c</mi><mi>d</mi></mfrac></mfrac>"),
            },
            new MathTypeOfflineCase
            {
                Name = "square-root-superscript",
                MathMl = M("<msqrt><mrow><msup><mi>r</mi><mn>2</mn></msup><mo>+</mo><msup><mi>s</mi><mn>2</mn></msup></mrow></msqrt>"),
                ValidateWithMathType = true,
            },
            new MathTypeOfflineCase { Name = "nth-root", MathMl = M("<mroot><mi>x</mi><mn>3</mn></mroot>") },
            new MathTypeOfflineCase { Name = "subsup", MathMl = M("<msubsup><mi>x</mi><mi>i</mi><mn>2</mn></msubsup>") },
            new MathTypeOfflineCase { Name = "paren", MathMl = M("<mfenced open=\"(\" close=\")\"><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow></mfenced>") },
            new MathTypeOfflineCase { Name = "one-sided-brace", MathMl = M("<mfenced open=\"{\" close=\".\"><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></mfenced>") },
            new MathTypeOfflineCase { Name = "floor-ceiling", MathMl = M("<mfenced open=\"⌊\" close=\"⌋\"><mi>x</mi></mfenced><mo>+</mo><mfenced open=\"⌈\" close=\"⌉\"><mi>y</mi></mfenced>") },
            new MathTypeOfflineCase { Name = "absolute-norm", MathMl = M("<mfenced open=\"|\" close=\"|\"><mi>x</mi></mfenced><mo>+</mo><mfenced open=\"‖\" close=\"‖\"><mi>v</mi></mfenced>") },
            new MathTypeOfflineCase
            {
                Name = "matrix-2x2",
                MathMl = M("<mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable>"),
                ValidateWithMathType = true,
            },
            new MathTypeOfflineCase
            {
                Name = "determinant",
                MathMl = M("<mfenced open=\"|\" close=\"|\"><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable></mfenced>"),
                ValidateWithMathType = true,
            },
            new MathTypeOfflineCase
            {
                Name = "mathjax-det-characteristic",
                MathMl = M("<mo data-mjx-texclass=\"OP\" movablelimits=\"true\">det</mo><mo stretchy=\"false\">(</mo><mi>A</mi><mo>−</mo><mi>λ</mi><mi>I</mi><mo stretchy=\"false\">)</mo><mo>=</mo><mn>0</mn>"),
                ValidateWithMathType = true,
            },
            new MathTypeOfflineCase
            {
                Name = "cases",
                MathMl = M("<mfenced open=\"{\" close=\".\"><mtable><mtr><mtd><mi>x</mi></mtd><mtd><mrow><mi>x</mi><mo>&gt;</mo><mn>0</mn></mrow></mtd></mtr><mtr><mtd><mo>−</mo><mi>x</mi></mtd><mtd><mrow><mi>x</mi><mo>≤</mo><mn>0</mn></mrow></mtd></mtr></mtable></mfenced>"),
                ValidateWithMathType = true,
            },
            new MathTypeOfflineCase { Name = "sum-limits", MathMl = M("<msubsup><mo>∑</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></msubsup><msub><mi>a</mi><mi>i</mi></msub>") , ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "product-limits", MathMl = M("<msubsup><mo>∏</mo><mrow><mi>k</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></msubsup><msub><mi>x</mi><mi>k</mi></msub>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "coproduct-limits", MathMl = M("<msubsup><mo>∐</mo><mrow><mi>k</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></msubsup><msub><mi>X</mi><mi>k</mi></msub>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "integral-bounds", MathMl = M("<msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup><msup><mi>x</mi><mn>2</mn></msup><mi mathvariant=\"normal\">d</mi><mi>x</mi>"), ValidateWithMathType = true },
            new MathTypeOfflineCase
            {
                Name = "gaussian-integral-normal-symbol",
                MathMl = M(
                    "<msubsup><mo>∫</mo><mrow><mo>−</mo><mo mathvariant=\"normal\">∞</mo></mrow><mo mathvariant=\"normal\">∞</mo></msubsup>"
                    + "<msup><mi mathvariant=\"normal\">e</mi><mrow><mo>−</mo><msup><mi>x</mi><mn>2</mn></msup></mrow></msup>"
                    + "<mi mathvariant=\"normal\">d</mi><mi>x</mi><mo>=</mo><msqrt><mi>π</mi></msqrt>"),
                ValidateWithMathType = true,
            },
            new MathTypeOfflineCase
            {
                Name = "mixed-two-integrals-infinity",
                MathMl = M(
                    "<mi>θ</mi>"
                    + "<msubsup><mo>∫</mo><mi>b</mi><mi>a</mi></msubsup><mi>d</mi><mi mathvariant=\"normal\">d</mi><mi>e</mi>"
                    + "<msubsup><mo>∫</mo><mrow><mo>−</mo><mo mathvariant=\"normal\">∞</mo></mrow><mo mathvariant=\"normal\">∞</mo></msubsup>"
                    + "<msup><mi mathvariant=\"normal\">e</mi><mrow><mo>−</mo><msup><mi>x</mi><mn>2</mn></msup></mrow></msup>"
                    + "<mi mathvariant=\"normal\">d</mi><mi>x</mi><mo>=</mo><msqrt><mi>π</mi></msqrt>"),
                ValidateWithMathType = true,
            },
            new MathTypeOfflineCase
            {
                Name = "normal-variant-symbol-family",
                MathMl = M(
                    "<mo mathvariant=\"normal\">∞</mo><mo mathvariant=\"normal\">≤</mo><mo mathvariant=\"normal\">→</mo>"
                    + "<mo mathvariant=\"normal\">∈</mo><mo mathvariant=\"normal\">∪</mo><mo mathvariant=\"normal\">∅</mo>"),
                ValidateWithMathType = true,
            },
            new MathTypeOfflineCase
            {
                Name = "mixed-bigop-sizing",
                MathMl = M(
                    "<mfrac><mi>a</mi><mi>b</mi></mfrac><mo>+</mo>"
                    + "<mo>∫</mo><mi>c</mi><mo>.</mo><mi>f</mi><mi>g</mi><msup><mi>a</mi><mn>2</mn></msup>"
                    + "<mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo>"
                    + "<msup><mi>c</mi><mn>2</mn></msup><msup><mfenced><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow></mfenced><mi>n</mi></msup><mo>=</mo>"
                    + "<msubsup><mo>∑</mo><mrow><mi>k</mi><mo>=</mo><mn>0</mn></mrow><mi>n</mi></msubsup>"
                    + "<mfrac><mi>n</mi><mi>k</mi></mfrac><msup><mi>a</mi><mrow><mi>n</mi><mo>−</mo><mi>k</mi></mrow></msup><msup><mi>b</mi><mi>k</mi></msup>"
                    + "<msubsup><mo>∑</mo><mi>c</mi><mi>a</mi></msubsup><mi>b</mi>"),
                ValidateWithMathType = true,
            },
            new MathTypeOfflineCase
            {
                Name = "document1-quadratic",
                MathMl = M(
                    "<mi>x</mi><mo>=</mo><mfrac>"
                    + "<mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt>"
                    + "<msup><mi>b</mi><mrow><mn>2</mn></mrow></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi>"
                    + "</msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac>"),
            },
            new MathTypeOfflineCase
            {
                Name = "document1-binomial-matrix",
                MathMl = M(
                    "<mo stretchy=\"false\">(</mo><mi>a</mi><mo>+</mo><mi>b</mi>"
                    + "<msup><mo stretchy=\"false\">)</mo><mrow><mi>n</mi></mrow></msup><mo>=</mo>"
                    + "<munderover><mo data-mjx-texclass=\"OP\">∑</mo>"
                    + "<mrow><mi>k</mi><mo>=</mo><mn>0</mn></mrow><mrow><mi>n</mi></mrow></munderover>"
                    + "<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">(</mo>"
                    + "<mtable><mtr><mtd><mi>n</mi></mtd></mtr><mtr><mtd><mi>k</mi></mtd></mtr></mtable>"
                    + "<mo data-mjx-texclass=\"CLOSE\">)</mo></mrow>"
                    + "<msup><mi>a</mi><mrow><mi>n</mi><mo>−</mo><mi>k</mi></mrow></msup>"
                    + "<msup><mi>b</mi><mrow><mi>k</mi></mrow></msup>"),
            },
            new MathTypeOfflineCase
            {
                Name = "document1-limit-chain",
                MathMl = M(
                    "<msub><mi mathvariant=\"normal\">lim</mi><mrow><mi>a</mi><mo>→</mo><mi>b</mi></mrow></msub>"
                    + "<mi>c</mi><mo>=</mo><mi>d</mi><msup><mi>a</mi><mn>2</mn></msup>"
                    + "<mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup>"),
            },
            new MathTypeOfflineCase
            {
                Name = "document1-cos-theta",
                MathMl = M("<mi mathvariant=\"normal\">cos</mi><mi>θ</mi><mo>=</mo><mn>1</mn>"),
            },
            new MathTypeOfflineCase { Name = "double-integral", MathMl = M("<msubsup><mo>∬</mo><mi>D</mi><mrow></mrow></msubsup><mi>f</mi><mi mathvariant=\"normal\">d</mi><mi>A</mi>") },
            new MathTypeOfflineCase { Name = "limit", MathMl = M("<msub><mi mathvariant=\"normal\">lim</mi><mrow><mi>x</mi><mo>→</mo><mn>0</mn></mrow></msub><mfrac><mrow><mi>sin</mi><mi>x</mi></mrow><mi>x</mi></mfrac>") },
            new MathTypeOfflineCase { Name = "bar-macron", MathMl = M("<mover accent=\"true\"><mi>a</mi><mo>¯</mo></mover>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "overline", MathMl = M("<mover accent=\"true\"><mrow><mi>A</mi><mi>B</mi></mrow><mo>¯</mo></mover>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "overline-horizontal-bar", MathMl = M("<mover accent=\"true\"><mrow><mi>A</mi><mi>B</mi></mrow><mo>&#x2015;</mo></mover>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "underline", MathMl = M("<munder accentunder=\"true\"><mi>x</mi><mo>¯</mo></munder>") },
            new MathTypeOfflineCase { Name = "vector", MathMl = M("<mover accent=\"true\"><mi>v</mi><mo>→</mo></mover>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "overleftrightarrow", MathMl = M("<mover accent=\"true\"><mrow><mi>A</mi><mi>B</mi></mrow><mo>↔</mo></mover>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "hat", MathMl = M("<mover accent=\"true\"><mi>x</mi><mo>^</mo></mover>") },
            new MathTypeOfflineCase { Name = "tilde", MathMl = M("<mover accent=\"true\"><mi>x</mi><mo>~</mo></mover>") },
            new MathTypeOfflineCase { Name = "dot", MathMl = M("<mover accent=\"true\"><mi>x</mi><mo>.</mo></mover>") },
            new MathTypeOfflineCase { Name = "ddot", MathMl = M("<mover accent=\"true\"><mi>x</mi><mo>¨</mo></mover>") },
            new MathTypeOfflineCase { Name = "boxed", MathMl = M("<menclose notation=\"box\"><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></menclose>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "greek", MathMl = M("<mi>α</mi><mo>+</mo><mi>β</mi><mo>=</mo><mi>Γ</mi><mo>+</mo><mi>Ω</mi>") },
            new MathTypeOfflineCase { Name = "greek-epsilon-phi-variants", MathMl = M("<mi>ε</mi><mo>+</mo><mi>ϵ</mi><mo>+</mo><mi>φ</mi><mo>+</mo><mi>ϕ</mi>") },
            new MathTypeOfflineCase { Name = "sets-relations", MathMl = M("<mi>x</mi><mo>∈</mo><mi>A</mi><mo>⊆</mo><mi>B</mi><mo>∪</mo><mi>C</mi><mo>∩</mo><mi>D</mi>") },
            new MathTypeOfflineCase { Name = "relations", MathMl = M("<mi>a</mi><mo>≤</mo><mi>b</mi><mo>≠</mo><mi>c</mi><mo>≈</mo><mi>d</mi><mo>≥</mo><mi>e</mi>") },
            new MathTypeOfflineCase { Name = "arrows", MathMl = M("<mi>A</mi><mo>→</mo><mi>B</mi><mo>↔</mo><mi>C</mi><mo>⇒</mo><mi>D</mi>") },
            new MathTypeOfflineCase { Name = "upright-bold", MathMl = M("<mi mathvariant=\"normal\">e</mi><mo>+</mo><mi mathvariant=\"bold\">v</mi><mo>+</mo><mi>x</mi>") },
            new MathTypeOfflineCase { Name = "mathcal", MathMl = M("<mi mathvariant=\"script\">F</mi>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "mathbb", MathMl = M("<mi mathvariant=\"double-struck\">R</mi>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "mathfrak", MathMl = M("<mi mathvariant=\"fraktur\">g</mi>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "multiscripts", MathMl = M("<mmultiscripts><mi>T</mi><mi>i</mi><mi>j</mi></mmultiscripts>") },
            new MathTypeOfflineCase { Name = "function-text", MathMl = M("<mi>sin</mi><mfenced><mi>x</mi></mfenced><mo>+</mo><mi>cos</mi><mfenced><mi>y</mi></mfenced>") },
            new MathTypeOfflineCase
            {
                Name = "mathjax-limit",
                MathMl = M("<munder><mo data-mjx-texclass=\"OP\" movablelimits=\"true\">lim</mo><mrow><mi>x</mi><mo>→</mo><mn>0</mn></mrow></munder><mfrac><mrow><mi>sin</mi><mo>⁡</mo><mi>x</mi></mrow><mi>x</mi></mfrac>"),
            },
            new MathTypeOfflineCase
            {
                Name = "mathjax-binom",
                MathMl = M("<mrow data-mjx-texclass=\"ORD\"><mrow data-mjx-texclass=\"OPEN\"><mo minsize=\"2.047em\" maxsize=\"2.047em\">(</mo></mrow><mfrac linethickness=\"0\"><mi>n</mi><mi>k</mi></mfrac><mrow data-mjx-texclass=\"CLOSE\"><mo minsize=\"2.047em\" maxsize=\"2.047em\">)</mo></mrow></mrow>"),
            },
            new MathTypeOfflineCase
            {
                // Word's OMML -> MathML XSLT emits a no-bar fraction between
                // ordinary parenthesis tokens, without MathJax OPEN/CLOSE tags,
                // and keeps numerator/denominator content inside mrow wrappers.
                // The MTEF writer turns the no-bar fraction into a native PILE;
                // the reader must still recognize that loose PILE fence as binom.
                Name = "omml-loose-binom",
                MathMl = M("<mo>(</mo><mfrac linethickness=\"0\"><mrow><mi>n</mi></mrow><mrow><mi>k</mi></mrow></mfrac><mo>)</mo><mo>=</mo><mn>1</mn>"),
            },
            new MathTypeOfflineCase
            {
                Name = "mathjax-pmatrix",
                MathMl = M("<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">(</mo><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable><mo data-mjx-texclass=\"CLOSE\">)</mo></mrow>"),
            },
            new MathTypeOfflineCase
            {
                Name = "mathjax-cases",
                MathMl = M("<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">{</mo><mtable><mtr><mtd><mi>x</mi></mtd><mtd><mi>x</mi><mo>&gt;</mo><mn>0</mn></mtd></mtr><mtr><mtd><mo>−</mo><mi>x</mi></mtd><mtd><mi>x</mi><mo>≤</mo><mn>0</mn></mtd></mtr></mtable><mo data-mjx-texclass=\"CLOSE\" fence=\"true\" stretchy=\"true\"></mo></mrow>"),
            },
            new MathTypeOfflineCase
            {
                Name = "mathjax-overbrace",
                MathMl = M("<mover><mrow data-mjx-texclass=\"OP\"><mover><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mo>⏞</mo></mover></mrow><mrow><mi>n</mi></mrow></mover>"),
                ValidateWithMathType = true,
            },
            new MathTypeOfflineCase
            {
                Name = "mathjax-cancel",
                MathMl = M("<menclose notation=\"updiagonalstrike\"><mi>x</mi><mo>+</mo><mn>1</mn></menclose>"),
            },
            new MathTypeOfflineCase
            {
                Name = "mathjax-oint",
                MathMl = M("<msub><mo data-mjx-texclass=\"OP\">∮</mo><mi>C</mi></msub><mi>f</mi><mo stretchy=\"false\">(</mo><mi>z</mi><mo stretchy=\"false\">)</mo><mi>d</mi><mi>z</mi>"),
            },
            new MathTypeOfflineCase
            {
                Name = "mathjax-iint",
                MathMl = M("<msub><mo data-mjx-texclass=\"OP\">∬</mo><mi>D</mi></msub><mi>f</mi><mstyle scriptlevel=\"0\"><mspace width=\"0.167em\"></mspace></mstyle><mi>d</mi><mi>A</mi>"),
            },
            new MathTypeOfflineCase
            {
                Name = "mathjax-mathbb",
                MathMl = M("<mrow data-mjx-texclass=\"ORD\"><mi mathvariant=\"double-struck\">R</mi></mrow>"),
            },
            new MathTypeOfflineCase
            {
                Name = "mathjax-text",
                MathMl = M("<mtext>if </mtext><mi>x</mi><mo>&gt;</mo><mn>0</mn>"),
            },
            new MathTypeOfflineCase
            {
                Name = "mathjax-underbrace",
                MathMl = M("<munder><mrow data-mjx-texclass=\"OP\"><munder><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mo>⏟</mo></munder></mrow><mrow><mi>n</mi></mrow></munder>"),
                ValidateWithMathType = true,
            },
            new MathTypeOfflineCase { Name = "mathjax-overset", MathMl = M("<mover><mi>x</mi><mo>∗</mo></mover>") },
            new MathTypeOfflineCase { Name = "mathjax-underset", MathMl = M("<munder><mi>x</mi><mi>n</mi></munder>") },
            new MathTypeOfflineCase { Name = "mathjax-overrightarrow", MathMl = M("<mover><mrow><mi>A</mi><mi>B</mi></mrow><mo>→</mo></mover>") },
            new MathTypeOfflineCase
            {
                Name = "doc5-angular-momentum",
                MathMl = M(
                    "<mi>L</mi><mo>=</mo><mo>−</mo><mtext> </mtext><mi>i</mi><mi>ℏ</mi>"
                    + "<mrow><mo>(</mo>"
                    + "<mover><msub><mi>e</mi><mrow><mi>φ</mi></mrow></msub><mo>→</mo></mover>"
                    + "<mfrac><mi>∂</mi><mrow><mi>∂</mi><mi>θ</mi></mrow></mfrac><mo>−</mo>"
                    + "<mover><msub><mi>e</mi><mrow><mi>θ</mi></mrow></msub><mo>→</mo></mover>"
                    + "<mfrac><mn>1</mn><mrow><mi>s</mi><mi>i</mi><mi>n</mi><mi>θ</mi></mrow></mfrac>"
                    + "<mfrac><mi>∂</mi><mrow><mi>∂</mi><mi>φ</mi></mrow></mfrac><mo>)</mo></mrow>"),
            },
            new MathTypeOfflineCase { Name = "mathjax-overleftarrow", MathMl = M("<mover><mrow><mi>A</mi><mi>B</mi></mrow><mo>←</mo></mover>") },
            new MathTypeOfflineCase { Name = "mathjax-widehat", MathMl = M("<mrow><mover><mrow><mi>A</mi><mi>B</mi><mi>C</mi></mrow><mo>^</mo></mover></mrow>") },
            new MathTypeOfflineCase { Name = "mathjax-widetilde", MathMl = M("<mrow><mover><mrow><mi>A</mi><mi>B</mi><mi>C</mi></mrow><mo>~</mo></mover></mrow>") },
            new MathTypeOfflineCase
            {
                Name = "mathjax-bmatrix",
                MathMl = M("<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">[</mo><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable><mo data-mjx-texclass=\"CLOSE\">]</mo></mrow>"),
                ValidateWithMathType = true,
                Inline = false,
            },
            new MathTypeOfflineCase { Name = "mathjax-curly-Bmatrix", MathMl = M("<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">{</mo><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable><mo data-mjx-texclass=\"CLOSE\">}</mo></mrow>") },
            new MathTypeOfflineCase { Name = "mathjax-vmatrix", MathMl = M("<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">|</mo><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable><mo data-mjx-texclass=\"CLOSE\">|</mo></mrow>") },
            new MathTypeOfflineCase { Name = "mathjax-double-vmatrix", MathMl = M("<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">‖</mo><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable><mo data-mjx-texclass=\"CLOSE\">‖</mo></mrow>") },
            new MathTypeOfflineCase { Name = "mathjax-bigcup", MathMl = M("<munderover><mo data-mjx-texclass=\"OP\">⋃</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mrow><mi>n</mi></mrow></munderover><msub><mi>A</mi><mi>i</mi></msub>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "mathjax-bigcap", MathMl = M("<munderover><mo data-mjx-texclass=\"OP\">⋂</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mrow><mi>n</mi></mrow></munderover><msub><mi>A</mi><mi>i</mi></msub>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "mathjax-max", MathMl = M("<munder><mo data-mjx-texclass=\"OP\" movablelimits=\"true\">max</mo><mrow><mi>x</mi><mo>∈</mo><mi>A</mi></mrow></munder><mi>f</mi><mo>(</mo><mi>x</mi><mo>)</mo>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "mathjax-min", MathMl = M("<munder><mo data-mjx-texclass=\"OP\" movablelimits=\"true\">min</mo><mrow><mi>x</mi><mo>∈</mo><mi>A</mi></mrow></munder><mi>f</mi><mo>(</mo><mi>x</mi><mo>)</mo>") },
            new MathTypeOfflineCase { Name = "mathjax-sup", MathMl = M("<munder><mo data-mjx-texclass=\"OP\" movablelimits=\"true\">sup</mo><mrow><mi>x</mi><mo>∈</mo><mi>A</mi></mrow></munder><mi>f</mi><mo>(</mo><mi>x</mi><mo>)</mo>") },
            new MathTypeOfflineCase { Name = "mathjax-inf", MathMl = M("<munder><mo data-mjx-texclass=\"OP\" movablelimits=\"true\">inf</mo><mrow><mi>x</mi><mo>∈</mo><mi>A</mi></mrow></munder><mi>f</mi><mo>(</mo><mi>x</mi><mo>)</mo>") },
            new MathTypeOfflineCase { Name = "mathjax-iiint", MathMl = M("<msub><mo data-mjx-texclass=\"OP\">∭</mo><mi>V</mi></msub><mi>f</mi><mspace width=\"0.167em\"></mspace><mi>d</mi><mi>V</mi>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "mathjax-bcancel", MathMl = M("<menclose notation=\"downdiagonalstrike\"><mi>x</mi><mo>+</mo><mn>1</mn></menclose>") },
            new MathTypeOfflineCase { Name = "mathjax-xcancel", MathMl = M("<menclose notation=\"updiagonalstrike downdiagonalstrike\"><mi>x</mi><mo>+</mo><mn>1</mn></menclose>"), ValidateWithMathType = true },
        };

        var offlineCaseFilter = Environment.GetEnvironmentVariable(
            "VISUALTEX_MATHTYPE_OFFLINE_CASE");
        if (!string.IsNullOrWhiteSpace(offlineCaseFilter))
        {
            cases = cases.Where(testCase => string.Equals(
                    testCase.Name,
                    offlineCaseFilter,
                    StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (cases.Length == 0)
                throw new InvalidDataException(
                    $"Unknown MathType offline case filter '{offlineCaseFilter}'.");
            Console.WriteLine($"MathType offline case filter={offlineCaseFilter}.");
        }

        Console.WriteLine(
            "[MathType offline 1/4] Reading the synchronized real Equation.DSMT4 directly from CFB/MTEF with no MathType server...");
        var readBefore = SnapshotMathTypeProcessIds();
        var sourceCompound = ReadSingleDocxOleEmbedding(source);
        var sourceMathMl = MathTypeOleStorage.ReadMathMl(sourceCompound);
        var sourceLatex = MathMlToLatexConverter.Convert(sourceMathMl).Trim();
        AssertMathTypeLatexEquivalent(
            "\\sqrt{p^2+q^2}",
            sourceLatex,
            "Offline MTEF reader recovered the wrong source from the synchronized MathType 7 fixture.");
        Thread.Sleep(200);
        var readAfter = SnapshotMathTypeProcessIds();
        var readStarted = readAfter.Except(readBefore).ToArray();
        if (readStarted.Length > 0)
            throw new InvalidDataException(
                "Offline MathType read unexpectedly started MathType process(es): "
                + string.Join(", ", readStarted));
        Console.WriteLine($"  Offline reader source={sourceLatex}; no MathType process started.");

        Console.WriteLine(
            $"[MathType offline 2/4] Rewriting {cases.Length} real Equation.DSMT4 documents without MathType...");
        var targets = new List<(MathTypeOfflineCase TestCase, string Path)>();
        var recoveredLatex = new List<(MathTypeOfflineCase TestCase, string Latex)>();
        foreach (var testCase in cases)
        {
            var target = Path.Combine(
                artifactRoot,
                $"MathType7-Offline-Rewrite-{testCase.Name}.docx");
            File.Copy(source, target, overwrite: true);
            var before = SnapshotMathTypeProcessIds();
            var rewritten = RewriteSingleDocxMathTypeEmbedding(
                target,
                testCase.MathMl,
                inline: testCase.Inline);
            Thread.Sleep(200);
            var after = SnapshotMathTypeProcessIds();
            var newProcesses = after.Except(before).ToArray();
            if (newProcesses.Length > 0)
                throw new InvalidDataException(
                    $"Offline rewrite '{testCase.Name}' unexpectedly started MathType process(es): "
                    + string.Join(", ", newProcesses));
            AssertEqual(217, rewritten.StructureOffset,
                $"Offline rewrite '{testCase.Name}' did not preserve the source MathType header/preferences prefix.");
            AssertTrue(
                MathTypeOleStorage.LooksLikeMathTypeCompoundFile(rewritten.CompoundFile),
                $"Offline rewrite '{testCase.Name}' stopped looking like Equation.DSMT4 CFB.");
            if (!string.IsNullOrWhiteSpace(offlineCaseFilter)
                || string.Equals(testCase.Name, "mathjax-bmatrix", StringComparison.Ordinal)
                || string.Equals(testCase.Name, "mathbb", StringComparison.Ordinal))
            {
                File.WriteAllBytes(
                    Path.Combine(artifactRoot, $"visualtex-{testCase.Name}-compound.cfb"),
                    rewritten.CompoundFile);
                File.WriteAllBytes(
                    Path.Combine(artifactRoot, $"visualtex-{testCase.Name}-equation-native.bin"),
                    MathTypeOleStorage.ReadEquationNative(rewritten.CompoundFile));
            }
            var offlineReadMathMl = MathTypeOleStorage.ReadMathMl(rewritten.CompoundFile);
            var offlineReadLatex = MathMlToLatexConverter.Convert(offlineReadMathMl).Trim();
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(testCase.MathMl),
                MathTypeMtefCodec.SemanticSignature(offlineReadMathMl),
                $"Offline MTEF semantic mismatch after rewrite '{testCase.Name}'. expected MathML='{testCase.MathMl}', actual MathML='{offlineReadMathMl}'");
            Console.WriteLine(
                $"  {testCase.Name}: rewritten root offset={rewritten.StructureOffset}; offline read={offlineReadLatex}; no MathType process started.");
            targets.Add((testCase, target));
            recoveredLatex.Add((testCase, offlineReadLatex));
        }

        if (string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_MATHTYPE_VALIDATE_VISUALTEX_ROUNDTRIP"),
                "1",
                StringComparison.Ordinal))
        {
            Console.WriteLine(
                $"[MathType offline 2b/4] Re-parsing {recoveredLatex.Count} MTEF-recovered LaTeX formulas through the production VisualTeX MathJax converter...");
            client.PrewarmConverterAsync(CancellationToken.None).GetAwaiter().GetResult();
            var sessions = new List<(MathTypeOfflineCase TestCase, string Latex, string SessionId)>();
            try
            {
                foreach (var item in recoveredLatex)
                {
                    var line = new FormulaLine
                    {
                        Id = Guid.NewGuid().ToString("D"),
                        Latex = item.Latex,
                    };
                    var session = client.CreateSessionAsync(
                            new CreateVstoSessionRequest
                            {
                                Mode = "create",
                                Host = "word",
                                Title = "MathType semantic round-trip acceptance",
                                Lines = new List<FormulaLine> { line },
                                ActiveLineId = line.Id,
                                CodeFormat = "latex",
                                DisplayMode = item.TestCase.Inline ? "inline" : "block",
                                ObjectMode = FormulaOleContract.MathTypeOleMode,
                                Numbered = false,
                                FontSizePt = 12d,
                                AutoCommitOnClose = false,
                            },
                            CancellationToken.None)
                        .GetAwaiter().GetResult();
                    sessions.Add((item.TestCase, item.Latex, session.Id));
                }

                client.OpenConverterBatchAsync(
                        sessions.Select(item => item.SessionId).ToArray(),
                        CancellationToken.None)
                    .GetAwaiter().GetResult();

                foreach (var item in sessions)
                {
                    var session = client.WaitForCommitAsync(
                            item.SessionId,
                            TimeSpan.FromMinutes(3),
                            CancellationToken.None)
                        .GetAwaiter().GetResult();
                    if (string.Equals(session.Status, "failed", StringComparison.Ordinal))
                        throw new InvalidDataException(
                            $"VisualTeX MathJax could not re-parse MathType-recovered LaTeX for '{item.TestCase.Name}': {item.Latex}. {session.Error}");
                    var reparsedMathMl = session.ExportResult?.MathMl;
                    if (string.IsNullOrWhiteSpace(reparsedMathMl))
                        throw new InvalidDataException(
                            $"VisualTeX MathJax returned no MathML while re-parsing '{item.TestCase.Name}': {item.Latex}");
                    var expectedSignature = MathTypeMtefCodec.SemanticSignature(item.TestCase.MathMl);
                    var reparsedSignature = MathTypeMtefCodec.SemanticSignature(reparsedMathMl!);
                    AssertEqual(
                        expectedSignature,
                        reparsedSignature,
                        $"MathType -> VisualTeX LaTeX -> MathJax semantic mismatch for '{item.TestCase.Name}'. recoveredLatex='{item.Latex}', reparsedMathMl='{reparsedMathMl}'");

                    var regenerated = MathTypeMtefCodec.CreateEquationNative(
                        reparsedMathMl!,
                        item.TestCase.Inline);
                    var regeneratedCompound = MathTypeOleStorage.CreateStandaloneCompoundFile(regenerated);
                    var secondReadMathMl = MathTypeOleStorage.ReadMathMl(regeneratedCompound);
                    AssertEqual(
                        expectedSignature,
                        MathTypeMtefCodec.SemanticSignature(secondReadMathMl),
                        $"Second VisualTeX -> MathType MTEF semantic mismatch for '{item.TestCase.Name}'. recoveredLatex='{item.Latex}', secondReadMathMl='{secondReadMathMl}'");
                    Console.WriteLine(
                        $"  {item.TestCase.Name}: MTEF -> LaTeX -> MathJax -> MTEF semantic round-trip passed.");
                }
            }
            finally
            {
                foreach (var item in sessions)
                {
                    try
                    {
                        client.CompleteAsync(item.SessionId, CancellationToken.None)
                            .GetAwaiter().GetResult();
                        client.CloseEditorAsync(item.SessionId, CancellationToken.None)
                            .GetAwaiter().GetResult();
                    }
                    catch { }
                }
            }
        }

        Console.WriteLine(
            "[MathType offline 3/4] Letting installed MathType 7 validate each VisualTeX-generated MTEF...");
        var validationCaseFilter = Environment.GetEnvironmentVariable(
            "VISUALTEX_MATHTYPE_VALIDATION_CASE");
        var skipInstalledValidation = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_MATHTYPE_SKIP_INSTALLED_VALIDATION"),
            "1",
            StringComparison.Ordinal);
        var validationTargets = skipInstalledValidation
            ? Enumerable.Empty<(MathTypeOfflineCase TestCase, string Path)>()
            : targets.Where(item => item.TestCase.ValidateWithMathType);
        if (skipInstalledValidation)
            Console.WriteLine("  Installed MathType validation skipped; offline semantic corpus remains authoritative for this run.");
        if (!string.IsNullOrWhiteSpace(validationCaseFilter))
        {
            validationTargets = validationTargets.Where(item => string.Equals(
                item.TestCase.Name,
                validationCaseFilter,
                StringComparison.OrdinalIgnoreCase));
            Console.WriteLine($"  MathType validation filter={validationCaseFilter}.");
        }
        foreach (var (testCase, target) in validationTargets)
        {
            Word.Application? application = null;
            Word.Document? document = null;
            Word.InlineShape? shape = null;
            try
            {
                var interactiveValidation = !string.IsNullOrWhiteSpace(validationCaseFilter);
                application = CreateWordApplication(visible: interactiveValidation);
                document = application.Documents.Open(
                    target,
                    ReadOnly: true,
                    Visible: interactiveValidation);
                if (interactiveValidation)
                    document.Activate();
                AssertEqual(1, document.InlineShapes.Count,
                    $"Offline rewrite '{testCase.Name}' changed the Word OLE object inventory.");
                shape = document.InlineShapes[1];
                AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                    $"Offline rewrite '{testCase.Name}' is no longer recognized as MathType OLE.");

                string validationMathMl;
                if (!string.IsNullOrWhiteSpace(validationCaseFilter))
                {
                    Word.OLEFormat? validationFormat = null;
                    try
                    {
                        validationFormat = shape.OLEFormat;
                        validationMathMl = InvokeWordOwnedMathTypeEditor(
                            application,
                            validationFormat,
                            replacementLatex: null,
                            saveChanges: false);
                    }
                    finally { Release(validationFormat); }
                }
                else
                {
                    validationMathMl = MathTypeOleInterop.ReadMathMl(shape);
                }
                if (!string.IsNullOrWhiteSpace(validationCaseFilter))
                {
                    File.WriteAllText(
                        Path.Combine(artifactRoot, $"mathtype-{testCase.Name}-readback.xml"),
                        validationMathMl);
                    var markerCodePoints = string.Join(
                        ",",
                        validationMathMl.Where(character => character == '\uFFFD')
                            .Select(character => $"U+{(int)character:X4}"));
                    Console.WriteLine($"  {testCase.Name}: replacement markers={markerCodePoints}.");
                }
                var validationLatex = MathMlToLatexConverter.Convert(validationMathMl).Trim();
                AssertEqual(
                    MathTypeMtefCodec.SemanticSignature(testCase.MathMl),
                    MathTypeMtefCodec.SemanticSignature(validationMathMl),
                    $"MathType 7 semantic read-back mismatch for '{testCase.Name}'. expected MathML='{testCase.MathMl}', actual MathML='{validationMathMl}'");
                Console.WriteLine(
                    $"  {testCase.Name}: MathType 7 read-back source={validationLatex}.");
            }
            finally
            {
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

        Console.WriteLine("[MathType offline 4/4] Rechecking persisted Compound File identity...");
        foreach (var (_, target) in targets)
        {
            var persisted = ReadSingleDocxOleEmbedding(target);
            AssertTrue(
                MathTypeOleStorage.LooksLikeMathTypeCompoundFile(persisted),
                $"Rewritten DOCX '{Path.GetFileName(target)}' no longer contains valid MathType CFB.");
        }
        Console.WriteLine(
            $"MathType offline-storage acceptance passed: VisualTeX rewrote and semantically re-read {cases.Length} MTEF v5 formula structures without starting MathType; {cases.Count(testCase => testCase.ValidateWithMathType)} representative structures were also read back by installed MathType 7.");
    }

    private static MathTypeOleStorage.RewriteResult RewriteSingleDocxMathTypeEmbedding(
        string target,
        string mathMl,
        bool inline)
    {
        using var archive = ZipFile.Open(target, ZipArchiveMode.Update);
        var embedded = archive.Entries
            .Where(entry => entry.FullName.StartsWith(
                "word/embeddings/oleObject",
                StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (embedded.Length != 1)
            throw new InvalidDataException(
                $"Expected exactly one embedded OLE object, found {embedded.Length}.");

        byte[] sourceCompound;
        using (var sourceStream = embedded[0].Open())
        using (var memory = new MemoryStream())
        {
            sourceStream.CopyTo(memory);
            sourceCompound = memory.ToArray();
        }
        AssertTrue(
            MathTypeOleStorage.LooksLikeMathTypeCompoundFile(sourceCompound),
            "The synchronized source embedding is not recognized as MathType CFB.");
        var sourceNative = MathTypeOleStorage.ReadEquationNative(sourceCompound);
        var sourceMtefLength = checked((int)BitConverter.ToUInt32(sourceNative, 8));
        var sourceMtef = new byte[sourceMtefLength];
        Buffer.BlockCopy(sourceNative, 28, sourceMtef, 0, sourceMtef.Length);
        AssertEqual(217, MathTypeMtefCodec.FindRootStructureOffset(sourceMtef),
            "The synchronized MathType 7 source root structure moved unexpectedly.");

        var rewritten = MathTypeOleStorage.RewriteMathTypeCompoundFile(
            sourceCompound,
            mathMl,
            inline);
        embedded[0].Delete();
        var replacement = archive.CreateEntry(
            "word/embeddings/oleObject1.bin",
            CompressionLevel.Optimal);
        using var destination = replacement.Open();
        destination.Write(rewritten.CompoundFile, 0, rewritten.CompoundFile.Length);
        return rewritten;
    }

    private static void AssertMathTypeLatexEquivalent(
        string expected,
        string actual,
        string message)
    {
        static string Normalize(string value) =>
            value.Replace(" ", string.Empty)
                .Replace("{", string.Empty)
                .Replace("}", string.Empty);
        AssertEqual(Normalize(expected), Normalize(actual),
            $"{message} expected='{expected}', actual='{actual}'");
    }

    private static void RunWordCaptionFrameRepairAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? numberRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? captionRange = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? numberBookmark = null;
        Word.Bookmark? captionBookmark = null;
        Word.Frames? frames = null;
        Word.Frame? frame = null;
        Word.Range? bodyProbe = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "BODY_AFTER_CAPTION\r";

            insertion = document.Range(0, 0);
            fields = insertion.Fields;
            object fieldType = Word.WdFieldType.wdFieldEmpty;
            object fieldCode = "SEQ Equation \\r 10 \\* ARABIC";
            object preserveFormatting = true;
            field = fields.Add(
                insertion,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            field.Update();
            numberRange = field.Result.Duplicate;
            paragraphs = numberRange.Paragraphs;
            paragraph = paragraphs[1];
            captionRange = paragraph.Range.Duplicate;

            var formulaId = Guid.NewGuid().ToString("D");
            var numberName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            var captionName = WordEquationNumbering.NativeCaptionBookmarkName(formulaId);
            bookmarks = document.Bookmarks;
            numberBookmark = bookmarks.Add(numberName, numberRange);
            captionBookmark = bookmarks.Add(captionName, captionRange);

            frames = captionRange.Frames;
            frame = frames.Add(captionRange);
            frame.WidthRule = Word.WdFrameSizeRule.wdFrameExact;
            frame.HeightRule = Word.WdFrameSizeRule.wdFrameExact;
            frame.Width = 0.1f;
            frame.Height = 0.1f;
            frame.RelativeHorizontalPosition =
                Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            frame.RelativeVerticalPosition =
                Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            frame.HorizontalPosition = (float)Word.WdFramePosition.wdFrameRight;
            frame.VerticalPosition = (float)Word.WdFramePosition.wdFrameBottom;
            frame.TextWrap = false;
            frame.LockAnchor = true;

            var damagedCaptionText = captionRange.Text ?? string.Empty;
            Console.WriteLine(
                $"[CAPTION FIXTURE] caption={captionRange.Start}:{captionRange.End} number={numberRange.Start}:{numberRange.End} "
                + $"frame={frame.Range.Start}:{frame.Range.End} width={frame.Width} height={frame.Height} "
                + $"widthRule={frame.WidthRule} heightRule={frame.HeightRule} wrap={frame.TextWrap} lock={frame.LockAnchor} "
                + $"text='{damagedCaptionText.Replace("\r", "<CR>")}'");
            AssertTrue(
                damagedCaptionText.IndexOf("BODY_AFTER_CAPTION", StringComparison.Ordinal) >= 0,
                "The damaged caption fixture did not reproduce a caption Frame containing the following body paragraph.");

            Release(frame);
            frame = null;
            Release(frames);
            frames = null;
            Release(captionBookmark);
            captionBookmark = null;
            Release(numberBookmark);
            numberBookmark = null;
            Release(bookmarks);
            bookmarks = null;
            Release(captionRange);
            captionRange = null;
            Release(paragraph);
            paragraph = null;
            Release(paragraphs);
            paragraphs = null;
            Release(numberRange);
            numberRange = null;
            Release(field);
            field = null;
            Release(fields);
            fields = null;
            Release(insertion);
            insertion = null;

            var repaired = WordEquationNumbering.RepairLeakedNativeCaptionFrames(document);
            AssertEqual(1, repaired,
                "VisualTeX did not repair exactly one leaked native caption Frame.");

            bookmarks = document.Bookmarks;
            AssertTrue(bookmarks.Exists(numberName),
                "Caption repair deleted the durable VTEqNum bookmark.");
            AssertTrue(bookmarks.Exists(captionName),
                "Caption repair deleted the durable VTEqCap bookmark.");
            numberBookmark = bookmarks[numberName];
            captionBookmark = bookmarks[captionName];
            numberRange = numberBookmark.Range;
            captionRange = captionBookmark.Range;
            var repairedCaptionText = captionRange.Text ?? string.Empty;
            AssertTrue(
                repairedCaptionText.IndexOf("BODY_AFTER_CAPTION", StringComparison.Ordinal) < 0,
                "Caption repair still leaves the following body paragraph inside VTEqCap.");
            AssertTrue(captionRange.Fields.Count == 1,
                "Caption repair lost or duplicated the native SEQ field.");
            frames = captionRange.Frames;
            AssertEqual(1, frames.Count,
                "Caption repair did not leave exactly one clipped native caption Frame.");

            Release(paragraph);
            paragraph = null;
            Release(paragraphs);
            paragraphs = captionRange.Paragraphs;
            paragraph = paragraphs[1];
            var captionParagraphEnd = paragraph.Range.End;
            var contentEnd = document.Content.End;
            AssertTrue(captionParagraphEnd < contentEnd,
                "Caption repair did not leave any following body paragraph.");
            bodyProbe = document.Range(
                captionParagraphEnd,
                Math.Min(contentEnd, captionParagraphEnd + 1));
            AssertEqual(0, bodyProbe.Frames.Count,
                "Caption repair still leaves the following body text inside the hidden Frame.");
            AssertTrue(bodyProbe.Paragraphs.Count > 0,
                "Caption repair did not preserve a standalone body paragraph.");
            var followingParagraphText = bodyProbe.Paragraphs[1].Range.Text ?? string.Empty;
            AssertTrue(
                followingParagraphText.IndexOf("BODY_AFTER_CAPTION", StringComparison.Ordinal) >= 0,
                "Caption repair deleted or displaced the following body text.");

            var outputPath = Path.Combine(artifactRoot, "word-caption-frame-repaired.docx");
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                "[CAPTION FRAME REPAIR PASS] legacy VTEqCap/VTEqNum Frame was split without losing the SEQ field or following body content.");
            Console.WriteLine($"Artifact: {outputPath}");
        }
        finally
        {
            Release(bodyProbe);
            Release(frame);
            Release(frames);
            Release(captionBookmark);
            Release(numberBookmark);
            Release(bookmarks);
            Release(captionRange);
            Release(paragraph);
            Release(paragraphs);
            Release(numberRange);
            Release(field);
            Release(fields);
            Release(insertion);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordVisualTeXMathTypeAdjacentFrameAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var genuineFlatOpcPath = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", "..", "..",
            "artifacts", "mathtype-openxml-unregistered", "source-wordopenxml.xml"));
        if (File.Exists(genuineFlatOpcPath))
        {
            var genuine = MathTypeWordOpenXml.Read(File.ReadAllText(genuineFlatOpcPath));
            var genuineWmfPath = Path.Combine(artifactRoot, "genuine-mathtype-preview.wmf");
            File.WriteAllBytes(genuineWmfPath, genuine.PreviewWmf);
            var genuineEmfPath = MathTypeWordOpenXml.ConvertPlaceableWmfToEnhancedMetafile(
                genuineWmfPath,
                genuine.WidthPt,
                genuine.HeightPt,
                artifactRoot);
            using var genuineBitmap = RenderEmf(
                File.ReadAllBytes(genuineEmfPath),
                Math.Max(240, (int)Math.Ceiling(genuine.WidthPt * 6d)),
                Math.Max(96, (int)Math.Ceiling(genuine.HeightPt * 6d)));
            Console.WriteLine(
                $"[GENUINE MATHTYPE PREVIEW] size={genuine.WidthPt:0.###}x{genuine.HeightPt:0.###}pt; "
                + $"left/right={MeasureLeftWhiteMargin(genuineBitmap)}/{MeasureRightWhiteMargin(genuineBitmap)}; "
                + $"edge-ink={DescribeEdgeInk(genuineBitmap)}");
        }
        var sourcePath = Path.Combine(artifactRoot, "adjacent-numbered-inline.tex");
        var bulkLogPath = Path.Combine(artifactRoot, "adjacent-numbered-inline.bulk.log");
        var tracePath = Path.Combine(artifactRoot, "adjacent-numbered-inline.conversion.trace.log");
        var outputPath = Path.Combine(artifactRoot, "adjacent-numbered-inline-mathtype.docx");
        const string firstLatex = @"\vec{a}\cdot\vec{b}=|\vec a||\vec b|\cos\theta";
        const string secondLatex = @"\vec{a}\times\vec{b}";
        var source = "$$" + firstLatex + "$$\r\n正文，$" + secondLatex + "$。\r\n";
        File.WriteAllText(sourcePath, source, new System.Text.UTF8Encoding(false));
        DeleteBulkPerformanceArtifact(bulkLogPath);
        DeleteBulkPerformanceArtifact(tracePath);
        DeleteBulkPerformanceArtifact(outputPath);

        var previousBulkSource = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH");
        var previousBulkFormat = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT");
        var previousBulkMode = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE");
        var previousBulkLog = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG");
        var previousFormatAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousTracePath = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        Word.Application? application = null;
        Word.Document? document = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Word.InlineShape? firstShape = null;
        Word.InlineShape? secondShape = null;
        Word.Range? firstRange = null;
        Word.Range? secondRange = null;
        Word.Frames? secondFrames = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? captionBookmark = null;
        Word.Range? captionRange = null;
        Array custom = Array.Empty<object>();
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", sourcePath);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", "latex");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", "ole");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", bulkLogPath);
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);

            var mathTypeBaseline = SnapshotMathTypeProcessIds();
            if (mathTypeBaseline.Count != 0)
                throw new InvalidOperationException(
                    "Adjacent numbered-inline acceptance requires MathType.exe to be closed before Word starts.");

            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Activate();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(
                application,
                Extensibility.ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);

            addIn.OnBulkImport(new object());
            WaitForBulkImportCompletion(bulkLogPath, TimeSpan.FromMinutes(3));
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(20));
            AssertEqual(2, document.InlineShapes.Count,
                "Adjacent fixture bulk import did not create exactly two VisualTeX OLE formulas.");

            firstShape = document.InlineShapes[1];
            secondShape = document.InlineShapes[2];
            var firstMetadata = WordFormulaMetadataReader.TryRead(firstShape)
                ?? throw new InvalidDataException("First adjacent VisualTeX OLE lost metadata after bulk import.");
            var secondMetadata = WordFormulaMetadataReader.TryRead(secondShape)
                ?? throw new InvalidDataException("Second adjacent VisualTeX OLE lost metadata after bulk import.");
            AssertEqual(firstLatex, firstMetadata.Latex,
                "First bulk-imported OLE changed the source formula before numbering.");
            AssertEqual(secondLatex, secondMetadata.Latex,
                "Second bulk-imported OLE already contains the previous formula tail before numbering.");

            // Reproduce the user's structure exactly: a numbered display formula
            // immediately followed by an ordinary inline VisualTeX formula.
            firstMetadata.Numbered = true;
            firstMetadata.DisplayMode = "block";
            firstRange = firstShape.Range.Duplicate;
            WordEquationNumbering.ReconcileFormula(
                document,
                firstRange,
                firstShape.Height,
                firstMetadata);

            Release(firstRange);
            firstRange = null;
            Release(firstShape);
            firstShape = null;
            Release(secondShape);
            secondShape = null;

            AssertEqual(2, document.InlineShapes.Count,
                "Numbering the first formula lost or duplicated an adjacent formula.");
            secondShape = document.InlineShapes[2];
            secondMetadata = WordFormulaMetadataReader.TryRead(secondShape)
                ?? throw new InvalidDataException("Second VisualTeX OLE lost metadata after numbering its predecessor.");
            AssertEqual(secondLatex, secondMetadata.Latex,
                "Numbering the preceding display formula changed the following inline formula metadata.");
            secondRange = secondShape.Range.Duplicate;
            secondFrames = secondRange.Frames;
            AssertEqual(0, secondFrames.Count,
                "New numbered-caption creation still leaks its hidden Frame into the following inline formula.");

            bookmarks = document.Bookmarks;
            var captionName = WordEquationNumbering.NativeCaptionBookmarkName(firstMetadata.FormulaId);
            AssertTrue(bookmarks.Exists(captionName),
                "Numbered display fixture did not create its native caption bookmark.");
            captionBookmark = bookmarks[captionName];
            captionRange = captionBookmark.Range;
            AssertTrue(captionRange.End <= secondRange.Start,
                $"Native caption overlaps the following inline formula: caption={captionRange.Start}:{captionRange.End}, inline={secondRange.Start}:{secondRange.End}.");

            Release(captionRange);
            captionRange = null;
            Release(captionBookmark);
            captionBookmark = null;
            Release(bookmarks);
            bookmarks = null;
            Release(secondFrames);
            secondFrames = null;
            Release(secondRange);
            secondRange = null;
            Release(secondShape);
            secondShape = null;

            ResetInstalledFormatConversionTrace(tracePath);
            addIn.OnConvertVisualTeXToMathTypeDocument(new object());
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=VisualTeX target=MathType",
                mathTypeBaseline);
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(20));
            AssertEqual(2, CountMathTypeOleShapes(document),
                "VisualTeX→MathType adjacent conversion did not create exactly two MathType OLE formulas.");
            AssertEqual(2, document.InlineShapes.Count,
                "VisualTeX→MathType adjacent conversion changed the formula object count.");

            firstShape = document.InlineShapes[1];
            secondShape = document.InlineShapes[2];
            var firstMathMl = MathTypeOleStorage.ReadMathMl(firstShape);
            var secondMathMl = MathTypeOleStorage.ReadMathMl(secondShape);
            var firstAfter = MathMlToLatexConverter.Convert(firstMathMl).Trim();
            var secondAfter = MathMlToLatexConverter.Convert(secondMathMl).Trim();
            AssertTrue(firstAfter.IndexOf("\\theta", StringComparison.Ordinal) >= 0,
                $"Converted first MathType formula lost theta: '{firstAfter}'.");
            AssertTrue(secondAfter.IndexOf("\\theta", StringComparison.Ordinal) < 0,
                $"Converted following inline MathType formula inherited theta from its predecessor: '{secondAfter}'.");
            AssertTrue(secondAfter.IndexOf("\\times", StringComparison.Ordinal) >= 0,
                $"Converted following inline MathType formula lost its multiplication operator: '{secondAfter}'.");

            var firstFragment = MathTypeWordOpenXml.Read(firstShape);
            var secondFragment = MathTypeWordOpenXml.Read(secondShape);
            var firstWmfPath = Path.Combine(artifactRoot, "adjacent-first-preview.wmf");
            var secondWmfPath = Path.Combine(artifactRoot, "adjacent-second-preview.wmf");
            File.WriteAllBytes(firstWmfPath, firstFragment.PreviewWmf);
            File.WriteAllBytes(secondWmfPath, secondFragment.PreviewWmf);
            var firstWmfEmfPath = MathTypeWordOpenXml.ConvertPlaceableWmfToEnhancedMetafile(
                firstWmfPath,
                firstShape.Width,
                firstShape.Height,
                artifactRoot);
            var secondWmfEmfPath = MathTypeWordOpenXml.ConvertPlaceableWmfToEnhancedMetafile(
                secondWmfPath,
                secondShape.Width,
                secondShape.Height,
                artifactRoot);
            var firstWmfEmf = File.ReadAllBytes(firstWmfEmfPath);
            var secondWmfEmf = File.ReadAllBytes(secondWmfEmfPath);
            var firstPreview = ReadInlineShapeEnhancedMetafile(firstShape);
            var secondPreview = ReadInlineShapeEnhancedMetafile(secondShape);
            var firstRenderWidth = Math.Max(240, (int)Math.Ceiling(firstShape.Width * 6d));
            var firstRenderHeight = Math.Max(96, (int)Math.Ceiling(firstShape.Height * 6d));
            var secondRenderWidth = Math.Max(180, (int)Math.Ceiling(secondShape.Width * 6d));
            var secondRenderHeight = Math.Max(72, (int)Math.Ceiling(secondShape.Height * 6d));
            using (var firstWmfBitmap = RenderEmf(firstWmfEmf, firstRenderWidth, firstRenderHeight))
            using (var secondWmfBitmap = RenderEmf(secondWmfEmf, secondRenderWidth, secondRenderHeight))
            using (var firstBitmap = RenderEmf(firstPreview, firstRenderWidth, firstRenderHeight))
            using (var secondBitmap = RenderEmf(secondPreview, secondRenderWidth, secondRenderHeight))
            {
                Console.WriteLine(
                    $"[ADJACENT WMF] first-left/right={MeasureLeftWhiteMargin(firstWmfBitmap)}/{MeasureRightWhiteMargin(firstWmfBitmap)}; "
                    + $"second-left/right={MeasureLeftWhiteMargin(secondWmfBitmap)}/{MeasureRightWhiteMargin(secondWmfBitmap)}; "
                    + $"first-edge-ink={DescribeEdgeInk(firstWmfBitmap)}; second-edge-ink={DescribeEdgeInk(secondWmfBitmap)}");
                var firstLeftMargin = MeasureLeftWhiteMargin(firstBitmap);
                var firstRightMargin = MeasureRightWhiteMargin(firstBitmap);
                var secondLeftMargin = MeasureLeftWhiteMargin(secondBitmap);
                var secondRightMargin = MeasureRightWhiteMargin(secondBitmap);
                AssertTrue(firstLeftMargin >= 4 && firstRightMargin >= 4,
                    $"Converted theta formula still touches the MathType preview edge: left/right={firstLeftMargin}/{firstRightMargin}px.");
                AssertTrue(secondLeftMargin >= 4 && secondRightMargin >= 4,
                    $"Converted adjacent inline formula still touches the MathType preview edge: left/right={secondLeftMargin}/{secondRightMargin}px.");
                Console.WriteLine(
                    $"[ADJACENT PREVIEW] first-left/right={firstLeftMargin}/{firstRightMargin}px/{firstBitmap.Width}px; "
                    + $"second-left/right={secondLeftMargin}/{secondRightMargin}px/{secondBitmap.Width}px; "
                    + $"first-edge-ink={DescribeEdgeInk(firstBitmap)}; second-edge-ink={DescribeEdgeInk(secondBitmap)}");
            }
            AssertNoNewMathTypeProcess(mathTypeBaseline,
                "adjacent numbered VisualTeX→MathType conversion");

            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                $"[ADJACENT WORD FLOW PASS] caption Frame stayed local and MathType semantics stayed isolated: first='{firstAfter}', second='{secondAfter}'.");
            Console.WriteLine($"Artifact: {outputPath}");
        }
        finally
        {
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
            Release(secondFrames);
            Release(secondRange);
            Release(firstRange);
            Release(secondShape);
            Release(firstShape);
            if (addIn is not null)
            {
                try
                {
                    addIn.OnDisconnection(
                        Extensibility.ext_DisconnectMode.ext_dm_UserClosed,
                        ref custom);
                }
                catch { }
            }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", previousBulkSource);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", previousBulkFormat);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", previousBulkMode);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", previousBulkLog);
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", previousFormatAcceptance);
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTracePath);
        }
    }

    private static string DescribeEdgeInk(System.Drawing.Bitmap bitmap)
    {
        var left = new List<int>();
        var right = new List<int>();
        var count = Math.Min(8, bitmap.Width);
        for (var offset = 0; offset < count; offset++)
        {
            left.Add(CountDarkPixelsInColumn(bitmap, offset));
            right.Add(CountDarkPixelsInColumn(bitmap, bitmap.Width - 1 - offset));
        }
        return "L[" + string.Join(",", left) + "] R[" + string.Join(",", right) + "]";
    }

    private static int CountDarkPixelsInColumn(System.Drawing.Bitmap bitmap, int x)
    {
        var count = 0;
        for (var y = 0; y < bitmap.Height; y++)
        {
            var pixel = bitmap.GetPixel(x, y);
            if (pixel.R <= 245 || pixel.G <= 245 || pixel.B <= 245)
                count++;
        }
        return count;
    }

    private static int MeasureLeftWhiteMargin(System.Drawing.Bitmap bitmap)
    {
        var minInkX = bitmap.Width;
        for (var y = 0; y < bitmap.Height; y++)
        {
            for (var x = 0; x < bitmap.Width; x++)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R > 245 && pixel.G > 245 && pixel.B > 245) continue;
                minInkX = Math.Min(minInkX, x);
                break;
            }
        }
        return minInkX == bitmap.Width ? bitmap.Width : minInkX;
    }

    private static int MeasureRightWhiteMargin(System.Drawing.Bitmap bitmap)
    {
        var maxInkX = -1;
        for (var y = 0; y < bitmap.Height; y++)
        {
            for (var x = bitmap.Width - 1; x >= 0; x--)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R > 245 && pixel.G > 245 && pixel.B > 245) continue;
                maxInkX = Math.Max(maxInkX, x);
                break;
            }
        }
        return maxInkX < 0 ? bitmap.Width : bitmap.Width - 1 - maxInkX;
    }

    private static void RunMathTypeAdjacentBatchIsolationAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var formulas = new[]
        {
            (Latex: @"\vec{a}\cdot\vec{b}=|\vec a||\vec b|\cos\theta", Display: "block"),
            (Latex: @"\vec{a}\times\vec{b}", Display: "inline"),
        };
        var parsed = WordBulkImportParser.Parse(
            "$$" + formulas[0].Latex + "$$\n正文，$" + formulas[1].Latex + "$。",
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Ole);
        var parsedFormulas = parsed.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => run.IsFormula)
            .ToArray();
        AssertEqual(2, parsedFormulas.Length,
            "Adjacent LaTeX parser fixture did not produce exactly two formulas.");
        AssertEqual(formulas[0].Latex, parsedFormulas[0].Latex,
            "Bulk-import parser changed the first adjacent formula.");
        AssertEqual(formulas[1].Latex, parsedFormulas[1].Latex,
            "Bulk-import parser leaked the first formula tail into the second formula.");
        var sessions = new List<(string Latex, string SessionId)>();
        try
        {
            client.PrewarmConverterAsync(CancellationToken.None).GetAwaiter().GetResult();
            foreach (var formula in formulas)
            {
                var line = new FormulaLine
                {
                    Id = Guid.NewGuid().ToString("D"),
                    Latex = formula.Latex,
                };
                var session = client.CreateSessionAsync(
                        new CreateVstoSessionRequest
                        {
                            Mode = "create",
                            Host = "word",
                            Title = "Adjacent batch isolation acceptance",
                            Lines = new List<FormulaLine> { line },
                            ActiveLineId = line.Id,
                            CodeFormat = "latex",
                            DisplayMode = formula.Display,
                            ObjectMode = FormulaOleContract.MathTypeOleMode,
                            Numbered = false,
                            FontSizePt = 12d,
                            AutoCommitOnClose = false,
                        },
                        CancellationToken.None)
                    .GetAwaiter().GetResult();
                sessions.Add((formula.Latex, session.Id));
            }

            client.OpenConverterBatchAsync(
                    sessions.Select(item => item.SessionId).ToArray(),
                    CancellationToken.None)
                .GetAwaiter().GetResult();

            var exported = new List<string>();
            foreach (var item in sessions)
            {
                var session = client.WaitForCommitAsync(
                        item.SessionId,
                        TimeSpan.FromMinutes(3),
                        CancellationToken.None)
                    .GetAwaiter().GetResult();
                if (string.Equals(session.Status, "failed", StringComparison.Ordinal))
                    throw new InvalidDataException(
                        $"Adjacent batch converter failed for '{item.Latex}': {session.Error}");
                var mathMl = session.ExportResult?.MathMl;
                if (string.IsNullOrWhiteSpace(mathMl))
                    throw new InvalidDataException(
                        $"Adjacent batch converter returned no MathML for '{item.Latex}'.");
                exported.Add(mathMl!);
            }

            var firstLatex = MathMlToLatexConverter.Convert(exported[0]).Trim();
            var secondLatex = MathMlToLatexConverter.Convert(exported[1]).Trim();
            AssertTrue(firstLatex.IndexOf("\\theta", StringComparison.Ordinal) >= 0,
                $"First adjacent formula lost its theta tail: '{firstLatex}'.");
            AssertTrue(secondLatex.IndexOf("\\theta", StringComparison.Ordinal) < 0,
                $"Second adjacent formula inherited theta from the previous converter session: '{secondLatex}'.");

            var equationNative = MathTypeMtefCodec.CreateEquationNative(exported[1], inline: true);
            var compound = MathTypeOleStorage.CreateStandaloneCompoundFile(equationNative);
            var rereadMathMl = MathTypeOleStorage.ReadMathMl(compound);
            var rereadLatex = MathMlToLatexConverter.Convert(rereadMathMl).Trim();
            AssertTrue(rereadLatex.IndexOf("\\theta", StringComparison.Ordinal) < 0,
                $"MathType MTEF encoder appended theta to the second adjacent formula: '{rereadLatex}'.");
            Console.WriteLine(
                $"[ADJACENT BATCH PASS] first='{firstLatex}' second='{secondLatex}' MTEF-second='{rereadLatex}'.");

            var previewSvg = Path.Combine(artifactRoot, "adjacent-word-preview.svg");
            File.WriteAllText(
                previewSvg,
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"96\" viewBox=\"0 0 240 96\"><text x=\"4\" y=\"64\" font-family=\"Times New Roman\" font-size=\"48\">x+1</text></svg>");
            var previewEmf = OfficeOlePreview.CreateVectorEmfFromSvg(previewSvg, 240, 96);
            Word.Application? application = null;
            Word.Document? document = null;
            Word.Range? range = null;
            try
            {
                application = CreateWordApplication(visible: false);
                document = application.Documents.Add();
                var service = new WordFormulaService(application);

                range = document.Range(document.Content.End - 1, document.Content.End - 1);
                range.Select();
                service.InsertMathTypeOle(
                    CreateMathTypeCreateSession(
                        displayMode: "block",
                        numbered: false,
                        latex: formulas[0].Latex),
                    exported[0],
                    previewEmf,
                    createdObjectBookmarkName: "VTMT_ADJACENT_FIRST");

                Release(range);
                range = document.Range(document.Content.End - 1, document.Content.End - 1);
                range.Text = "正文，";
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.Select();
                service.InsertMathTypeOle(
                    CreateMathTypeCreateSession(
                        displayMode: "inline",
                        numbered: false,
                        latex: formulas[1].Latex),
                    exported[1],
                    previewEmf,
                    createdObjectBookmarkName: "VTMT_ADJACENT_SECOND");

                AssertEqual(2, document.InlineShapes.Count,
                    "Adjacent Word insertion did not retain exactly two MathType OLE objects.");
                var immediateMathMl = new List<string>();
                for (var index = 1; index <= 2; index++)
                {
                    Word.InlineShape? shape = null;
                    try
                    {
                        shape = document.InlineShapes[index];
                        immediateMathMl.Add(MathTypeOleStorage.ReadMathMl(shape));
                    }
                    finally { Release(shape); }
                }
                var immediateSecondLatex = MathMlToLatexConverter.Convert(immediateMathMl[1]).Trim();
                AssertTrue(immediateSecondLatex.IndexOf("\\theta", StringComparison.Ordinal) < 0,
                    $"Word InsertXML leaked theta into the second adjacent MathType OLE before save: '{immediateSecondLatex}'.");

                var wordPath = Path.Combine(artifactRoot, "adjacent-word-insert.docx");
                document.SaveAs2(wordPath, Word.WdSaveFormat.wdFormatXMLDocument);
                document.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = application.Documents.Open(
                    wordPath,
                    ReadOnly: false,
                    AddToRecentFiles: false);
                AssertEqual(2, document.InlineShapes.Count,
                    "Adjacent Word insertion lost a MathType OLE after save/reopen.");
                Word.InlineShape? reopenedSecond = null;
                try
                {
                    reopenedSecond = document.InlineShapes[2];
                    var reopenedSecondMathMl = MathTypeOleStorage.ReadMathMl(reopenedSecond);
                    var reopenedSecondLatex = MathMlToLatexConverter.Convert(reopenedSecondMathMl).Trim();
                    AssertTrue(reopenedSecondLatex.IndexOf("\\theta", StringComparison.Ordinal) < 0,
                        $"Word save/reopen leaked theta into the second adjacent MathType OLE: '{reopenedSecondLatex}'.");
                    Console.WriteLine(
                        $"[ADJACENT WORD PASS] immediate-second='{immediateSecondLatex}' reopened-second='{reopenedSecondLatex}'.");
                }
                finally { Release(reopenedSecond); }
            }
            finally
            {
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
        finally
        {
            foreach (var item in sessions)
            {
                try
                {
                    client.CompleteAsync(item.SessionId, CancellationToken.None)
                        .GetAwaiter().GetResult();
                    client.CloseEditorAsync(item.SessionId, CancellationToken.None)
                        .GetAwaiter().GetResult();
                }
                catch { }
            }
        }
    }

    private static HashSet<int> SnapshotMathTypeProcessIds() =>
        Process.GetProcessesByName("MathType")
            .Select(process =>
            {
                try { return process.Id; }
                finally { process.Dispose(); }
            })
            .ToHashSet();

    private static byte[] ReadSingleDocxOleEmbedding(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var embedded = archive.Entries
            .Where(entry => entry.FullName.StartsWith(
                "word/embeddings/oleObject",
                StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (embedded.Length != 1)
            throw new InvalidDataException(
                $"Expected one DOCX OLE embedding, found {embedded.Length}.");
        using var stream = embedded[0].Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }
}
