using System.Diagnostics;
using System.IO.Compression;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class MathTypeOfflineCase
    {
        public string Name { get; set; } = string.Empty;
        public string MathMl { get; set; } = string.Empty;
        public bool ValidateWithMathType { get; set; }
        public bool Inline { get; set; } = true;
    }

    private static void RunWordMathTypeOfflineStorageAcceptance(string artifactRoot)
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
            new MathTypeOfflineCase { Name = "double-integral", MathMl = M("<msubsup><mo>∬</mo><mi>D</mi><mrow></mrow></msubsup><mi>f</mi><mi mathvariant=\"normal\">d</mi><mi>A</mi>") },
            new MathTypeOfflineCase { Name = "limit", MathMl = M("<msub><mi mathvariant=\"normal\">lim</mi><mrow><mi>x</mi><mo>→</mo><mn>0</mn></mrow></msub><mfrac><mrow><mi>sin</mi><mi>x</mi></mrow><mi>x</mi></mfrac>") },
            new MathTypeOfflineCase { Name = "overline", MathMl = M("<mover accent=\"true\"><mrow><mi>A</mi><mi>B</mi></mrow><mo>¯</mo></mover>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "underline", MathMl = M("<munder accentunder=\"true\"><mi>x</mi><mo>¯</mo></munder>") },
            new MathTypeOfflineCase { Name = "vector", MathMl = M("<mover accent=\"true\"><mi>v</mi><mo>→</mo></mover>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "hat", MathMl = M("<mover accent=\"true\"><mi>x</mi><mo>^</mo></mover>") },
            new MathTypeOfflineCase { Name = "tilde", MathMl = M("<mover accent=\"true\"><mi>x</mi><mo>~</mo></mover>") },
            new MathTypeOfflineCase { Name = "dot", MathMl = M("<mover accent=\"true\"><mi>x</mi><mo>.</mo></mover>") },
            new MathTypeOfflineCase { Name = "ddot", MathMl = M("<mover accent=\"true\"><mi>x</mi><mo>¨</mo></mover>") },
            new MathTypeOfflineCase { Name = "boxed", MathMl = M("<menclose notation=\"box\"><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></menclose>"), ValidateWithMathType = true },
            new MathTypeOfflineCase { Name = "greek", MathMl = M("<mi>α</mi><mo>+</mo><mi>β</mi><mo>=</mo><mi>Γ</mi><mo>+</mo><mi>Ω</mi>") },
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
        }

        Console.WriteLine(
            "[MathType offline 3/4] Letting installed MathType 7 validate each VisualTeX-generated MTEF...");
        var validationCaseFilter = Environment.GetEnvironmentVariable(
            "VISUALTEX_MATHTYPE_VALIDATION_CASE");
        var validationTargets = targets.Where(item => item.TestCase.ValidateWithMathType);
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
