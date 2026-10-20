using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace VisualTeX.WordVsto;

internal static class MathMlToLatexConverter
{
    // MathJax AMS alternate forms have the same semantic scalar as the
    // ordinary glyph. Only explicit per-token alternate evidence selects these
    // commands; bare Unicode and container attributes never imply an alternate.
    private static readonly IReadOnlyDictionary<string, string> AlternateTokenMap =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["ℏ"] = @"\hbar ",
            ["∥"] = @"\shortparallel ",
            ["∦"] = @"\nshortparallel ",
            ["∣"] = @"\shortmid ",
            ["∤"] = @"\nshortmid ",
            ["∖"] = @"\smallsetminus ",
            ["⊊"] = @"\varsubsetneq ",
            ["⊋"] = @"\varsupsetneq ",
            ["⫋"] = @"\varsubsetneqq ",
            ["⫌"] = @"\varsupsetneqq ",
            ["⊈"] = @"\nsubseteqq ",
            ["⊉"] = @"\nsupseteqq ",
            ["≰"] = @"\nleqq ",
            ["≱"] = @"\ngeqq ",
        };

    // Defaults of the canonical named token commands below, as authored by
    // MathJax's AMS mi mapping. Preserve a normal imported token explicitly
    // rather than letting command expansion silently restore implicit italic.
    private static readonly IReadOnlyDictionary<string, string> TokenDefaultMathVariants =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["∡"] = "italic",
            ["∢"] = "italic",
        };

    private static readonly IReadOnlyDictionary<string, string> TokenMap =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["−"] = "-",
            ["±"] = @"\pm ",
            ["∓"] = @"\mp ",
            ["×"] = @"\times ",
            ["·"] = @"\cdot ",
            ["⋅"] = @"\cdot ",
            ["°"] = @"\circ ",
            ["∘"] = @"\circ ",
            ["•"] = @"\bullet ",
            ["⋄"] = @"\diamond ",
            ["◃"] = @"\triangleleft ",
            ["⊲"] = @"\triangleleft ",
            ["▹"] = @"\triangleright ",
            ["⊳"] = @"\triangleright ",
            ["⊕"] = @"\oplus ",
            ["⊎"] = @"\uplus ",
            ["⊓"] = @"\sqcap ",
            ["⊔"] = @"\sqcup ",
            ["≀"] = @"\wr ",
            ["⨿"] = @"\amalg ",
            ["⊖"] = @"\ominus ",
            ["⊗"] = @"\otimes ",
            ["⊘"] = @"\oslash ",
            ["⊙"] = @"\odot ",
            ["◯"] = @"\bigcirc ",
            ["÷"] = @"\div ",
            ["∞"] = @"\infty ",
            ["ℏ"] = @"\hslash ",
            ["ℵ"] = @"\aleph ",
            ["℘"] = @"\wp ",
            ["ℜ"] = @"\Re ",
            ["ℑ"] = @"\Im ",
            ["ℓ"] = @"\ell ",
            ["ı"] = @"\imath ",
            ["ȷ"] = @"\jmath ",
            ["∫"] = @"\int ",
            ["∬"] = @"\iint ",
            ["∭"] = @"\iiint ",
            ["∮"] = @"\oint ",
            ["∑"] = @"\sum ",
            ["∏"] = @"\prod ",
            ["∐"] = @"\coprod ",
            ["⋃"] = @"\bigcup ",
            ["⋂"] = @"\bigcap ",
            ["∗"] = @"\ast ",
            ["⋆"] = @"\star ",
            ["†"] = @"\dagger ",
            ["‡"] = @"\ddagger ",
            ["'"] = @"\prime ",
            ["′"] = @"\prime ",
            ["″"] = @"\prime\prime ",
            ["‴"] = @"\prime\prime\prime ",
            ["⁗"] = @"\prime\prime\prime\prime ",
            ["‵"] = @"\backprime ",
            ["∀"] = @"\forall ",
            ["∃"] = @"\exists ",
            ["∣"] = @"\mid ",
            ["〈"] = @"\langle ",
            ["〈"] = @"\langle ",
            ["⟨"] = @"\langle ",
            ["〉"] = @"\rangle ",
            ["〉"] = @"\rangle ",
            ["⟩"] = @"\rangle ",
            ["∂"] = @"\partial ",
            ["∇"] = @"\nabla ",
            ["⁡"] = "⁡", // Preserve the actual application marker in its source slot.
            // A standalone radical glyph is TeX \\surd. Structured square
            // roots arrive as <msqrt> and are handled separately below.
            ["√"] = @"\surd ",
            ["≠"] = @"\ne ",
            ["≈"] = @"\approx ",
            ["≡"] = @"\equiv ",
            ["≤"] = @"\le ",
            ["≥"] = @"\ge ",
            ["≪"] = @"\ll ",
            ["≫"] = @"\gg ",
            ["≃"] = @"\simeq ",
            ["≍"] = @"\asymp ",
            ["≐"] = @"\doteq ",
            ["≺"] = @"\prec ",
            ["≻"] = @"\succ ",
            ["⪯"] = @"\preceq ",
            ["≼"] = @"\preccurlyeq ",
            ["⪰"] = @"\succeq ",
            ["≽"] = @"\succcurlyeq ",
            ["⊢"] = @"\vdash ",
            ["⊣"] = @"\dashv ",
            ["⊨"] = @"\models ",
            ["∠"] = @"\angle ",
            ["∡"] = @"\measuredangle ",
            ["∢"] = @"\sphericalangle ",
            ["∴"] = @"\therefore ",
            ["∵"] = @"\because ",
            // Geometry shapes, including the small/medium forms MathJax emits.
            ["△"] = @"\triangle ",
            ["▵"] = @"\triangle ",
            ["▲"] = @"\blacktriangle ",
            ["▴"] = @"\blacktriangle ",
            ["▽"] = @"\triangledown ",
            ["▿"] = @"\triangledown ",
            ["▼"] = @"\blacktriangledown ",
            ["▾"] = @"\blacktriangledown ",
            ["□"] = @"\square ",
            ["◻"] = @"\square ",
            ["■"] = @"\blacksquare ",
            ["◼"] = @"\blacksquare ",
            ["★"] = @"\bigstar ",
            ["∤"] = @"\nmid ",
            ["∦"] = @"\nparallel ",
            ["⊈"] = @"\nsubseteq ",
            ["⊉"] = @"\nsupseteq ",
            ["⊊"] = @"\subsetneq ",
            ["⊋"] = @"\supsetneq ",
            ["⫋"] = @"\subsetneqq ",
            ["⫌"] = @"\supsetneqq ",
            // Generated from MathJax's TeX symbol tables (base, AMS), keeping the
            // commands the VisualTeX editor kernel also defines, so a symbol read
            // back from MathType or Word comes back as its command, not raw Unicode.
            ["¥"] = @"\yen ",
            ["§"] = @"\S ",
            ["®"] = @"\circledR ",
            ["℧"] = @"\mho ",
            ["Ⅎ"] = @"\Finv ",
            ["ℶ"] = @"\beth ",
            ["ℷ"] = @"\gimel ",
            ["ℸ"] = @"\daleth ",
            ["⅁"] = @"\Game ",
            ["↚"] = @"\nleftarrow ",
            ["↛"] = @"\nrightarrow ",
            ["↞"] = @"\twoheadleftarrow ",
            ["↠"] = @"\twoheadrightarrow ",
            ["↢"] = @"\leftarrowtail ",
            ["↣"] = @"\rightarrowtail ",
            ["↫"] = @"\looparrowleft ",
            ["↬"] = @"\looparrowright ",
            ["↭"] = @"\leftrightsquigarrow ",
            ["↮"] = @"\nleftrightarrow ",
            ["↰"] = @"\Lsh ",
            ["↱"] = @"\Rsh ",
            ["↶"] = @"\curvearrowleft ",
            ["↷"] = @"\curvearrowright ",
            ["↺"] = @"\circlearrowleft ",
            ["↻"] = @"\circlearrowright ",
            ["↾"] = @"\upharpoonright ",
            ["↿"] = @"\upharpoonleft ",
            ["⇂"] = @"\downharpoonright ",
            ["⇃"] = @"\downharpoonleft ",
            ["⇄"] = @"\rightleftarrows ",
            ["⇆"] = @"\leftrightarrows ",
            ["⇇"] = @"\leftleftarrows ",
            ["⇈"] = @"\upuparrows ",
            ["⇉"] = @"\rightrightarrows ",
            ["⇊"] = @"\downdownarrows ",
            ["⇋"] = @"\leftrightharpoons ",
            ["⇌"] = @"\rightleftharpoons ",
            ["⇍"] = @"\nLeftarrow ",
            ["⇎"] = @"\nLeftrightarrow ",
            ["⇏"] = @"\nRightarrow ",
            ["⇚"] = @"\Lleftarrow ",
            ["⇛"] = @"\Rrightarrow ",
            ["⇝"] = @"\rightsquigarrow ",
            ["⇠"] = @"\dashleftarrow ",
            ["⇢"] = @"\dashrightarrow ",
            ["∁"] = @"\complement ",
            ["∄"] = @"\nexists ",
            ["∍"] = @"\backepsilon ",
            ["∔"] = @"\dotplus ",
            ["∙"] = @"\bullet ",
            ["∽"] = @"\backsim ",
            ["≁"] = @"\nsim ",
            ["≂"] = @"\eqsim ",
            ["≇"] = @"\ncong ",
            ["≊"] = @"\approxeq ",
            ["≎"] = @"\Bumpeq ",
            ["≏"] = @"\bumpeq ",
            ["≑"] = @"\doteqdot ",
            ["≒"] = @"\fallingdotseq ",
            ["≓"] = @"\risingdotseq ",
            ["≖"] = @"\eqcirc ",
            ["≗"] = @"\circeq ",
            ["≜"] = @"\triangleq ",
            ["≦"] = @"\leqq ",
            ["≧"] = @"\geqq ",
            ["≨"] = @"\lneqq ",
            ["≩"] = @"\gneqq ",
            ["≬"] = @"\between ",
            ["≮"] = @"\nless ",
            ["≯"] = @"\ngtr ",
            ["≰"] = @"\nleq ",
            ["≱"] = @"\ngeq ",
            ["≲"] = @"\lesssim ",
            ["≳"] = @"\gtrsim ",
            ["≶"] = @"\lessgtr ",
            ["≷"] = @"\gtrless ",
            ["≾"] = @"\precsim ",
            ["≿"] = @"\succsim ",
            ["⊀"] = @"\nprec ",
            ["⊁"] = @"\nsucc ",
            ["⊚"] = @"\circledcirc ",
            ["⊛"] = @"\circledast ",
            ["⊝"] = @"\circleddash ",
            ["⊞"] = @"\boxplus ",
            ["⊟"] = @"\boxminus ",
            ["⊠"] = @"\boxtimes ",
            ["⊡"] = @"\boxdot ",
            ["⊩"] = @"\Vdash ",
            ["⊪"] = @"\Vvdash ",
            ["⊬"] = @"\nvdash ",
            ["⊭"] = @"\nvDash ",
            ["⊮"] = @"\nVdash ",
            ["⊯"] = @"\nVDash ",
            ["⊴"] = @"\trianglelefteq ",
            ["⊵"] = @"\trianglerighteq ",
            ["⊸"] = @"\multimap ",
            ["⊺"] = @"\intercal ",
            ["⊻"] = @"\veebar ",
            ["⊼"] = @"\barwedge ",
            ["⋀"] = @"\bigwedge ",
            ["⋁"] = @"\bigvee ",
            ["⋇"] = @"\divideontimes ",
            ["⋈"] = @"\bowtie ",
            ["⋉"] = @"\ltimes ",
            ["⋊"] = @"\rtimes ",
            ["⋋"] = @"\leftthreetimes ",
            ["⋌"] = @"\rightthreetimes ",
            ["⋍"] = @"\backsimeq ",
            ["⋎"] = @"\curlyvee ",
            ["⋏"] = @"\curlywedge ",
            ["⋐"] = @"\Subset ",
            ["⋑"] = @"\Supset ",
            ["⋒"] = @"\Cap ",
            ["⋓"] = @"\Cup ",
            ["⋔"] = @"\pitchfork ",
            ["⋖"] = @"\lessdot ",
            ["⋗"] = @"\gtrdot ",
            ["⋘"] = @"\lll ",
            ["⋙"] = @"\ggg ",
            ["⋚"] = @"\lesseqgtr ",
            ["⋛"] = @"\gtreqless ",
            ["⋞"] = @"\curlyeqprec ",
            ["⋟"] = @"\curlyeqsucc ",
            ["⋠"] = @"\npreceq ",
            ["⋡"] = @"\nsucceq ",
            ["⋦"] = @"\lnsim ",
            ["⋧"] = @"\gnsim ",
            ["⋨"] = @"\precnsim ",
            ["⋩"] = @"\succnsim ",
            ["⋪"] = @"\ntriangleleft ",
            ["⋫"] = @"\ntriangleright ",
            ["⋬"] = @"\ntrianglelefteq ",
            ["⋭"] = @"\ntrianglerighteq ",
            ["⌢"] = @"\frown ",
            ["⌣"] = @"\smile ",
            ["Ⓢ"] = @"\circledS ",
            ["╱"] = @"\diagup ",
            ["▸"] = @"\blacktriangleright ",
            ["◂"] = @"\blacktriangleleft ",
            ["◊"] = @"\lozenge ",
            ["♭"] = @"\flat ",
            ["♮"] = @"\natural ",
            ["♯"] = @"\sharp ",
            ["✓"] = @"\checkmark ",
            ["✠"] = @"\maltese ",
            ["⟵"] = @"\longleftarrow ",
            ["⟶"] = @"\longrightarrow ",
            ["⟷"] = @"\longleftrightarrow ",
            ["⟸"] = @"\Longleftarrow ",
            ["⟹"] = @"\Longrightarrow ",
            ["⟺"] = @"\Longleftrightarrow ",
            ["⟼"] = @"\longmapsto ",
            ["⧫"] = @"\blacklozenge ",
            ["⨀"] = @"\bigodot ",
            ["⨁"] = @"\bigoplus ",
            ["⨂"] = @"\bigotimes ",
            ["⨄"] = @"\biguplus ",
            ["⨆"] = @"\bigsqcup ",
            ["⩞"] = @"\doublebarwedge ",
            ["⩽"] = @"\leqslant ",
            ["⩾"] = @"\geqslant ",
            ["⪅"] = @"\lessapprox ",
            ["⪆"] = @"\gtrapprox ",
            ["⪇"] = @"\nleqslant ",
            ["⪈"] = @"\ngeqslant ",
            ["⪉"] = @"\lnapprox ",
            ["⪊"] = @"\gnapprox ",
            ["⪋"] = @"\lesseqqgtr ",
            ["⪌"] = @"\gtreqqless ",
            ["⪕"] = @"\eqslantless ",
            ["⪖"] = @"\eqslantgtr ",
            // Characters MathType stores for these commands (the codec's
            // MathTypeStoredCharacters); read back as the same commands.
            ["⋜"] = @"\eqslantless ",
            ["⋝"] = @"\eqslantgtr ",
            ["⌆"] = @"\doublebarwedge ",
            ["◀"] = @"\blacktriangleleft ",
            ["▶"] = @"\blacktriangleright ",
            ["♦"] = @"\blacklozenge ",
            ["Ϝ"] = @"\digamma ",
            ["ϝ"] = @"\digamma ",
            ["⪵"] = @"\precneqq ",
            ["⪶"] = @"\succneqq ",
            ["⪷"] = @"\precapprox ",
            ["⪸"] = @"\succapprox ",
            ["⪹"] = @"\precnapprox ",
            ["⪺"] = @"\succnapprox ",
            ["⫅"] = @"\subseteqq ",
            ["⫆"] = @"\supseteqq ",
            ["∖"] = @"\setminus ",
            ["∝"] = @"\propto ",
            ["∼"] = @"\sim ",
            ["~"] = @"\sim ",
            ["≅"] = @"\cong ",
            ["∈"] = @"\in ",
            ["∉"] = @"\notin ",
            ["∋"] = @"\ni ",
            ["⊏"] = @"\sqsubset ",
            ["⊑"] = @"\sqsubseteq ",
            ["⊐"] = @"\sqsupset ",
            ["⊒"] = @"\sqsupseteq ",
            ["⊂"] = @"\subset ",
            ["⊆"] = @"\subseteq ",
            ["⊃"] = @"\supset ",
            ["⊇"] = @"\supseteq ",
            ["∪"] = @"\cup ",
            ["∩"] = @"\cap ",
            ["∅"] = @"\varnothing ",
            ["∧"] = @"\land ",
            ["∨"] = @"\lor ",
            ["¬"] = @"\neg ",
            ["⇐"] = @"\Leftarrow ",
            ["⇒"] = @"\Rightarrow ",
            ["⇔"] = @"\Leftrightarrow ",
            ["→"] = @"\to ",
            ["←"] = @"\leftarrow ",
            ["↔"] = @"\leftrightarrow ",
            ["↑"] = @"\uparrow ",
            ["↓"] = @"\downarrow ",
            ["↕"] = @"\updownarrow ",
            ["⇑"] = @"\Uparrow ",
            ["⇓"] = @"\Downarrow ",
            ["⇕"] = @"\Updownarrow ",
            ["↼"] = @"\leftharpoonup ",
            ["↽"] = @"\leftharpoondown ",
            ["⇀"] = @"\rightharpoonup ",
            ["⇁"] = @"\rightharpoondown ",
            ["↩"] = @"\hookleftarrow ",
            ["↪"] = @"\hookrightarrow ",
            ["↦"] = @"\mapsto ",
            ["↗"] = @"\nearrow ",
            ["↘"] = @"\searrow ",
            ["↙"] = @"\swarrow ",
            ["↖"] = @"\nwarrow ",
            ["⟂"] = @"\perp ",
            ["⊤"] = @"\top ",
            ["⊥"] = @"\perp ",
            ["∥"] = @"\parallel ",
            ["…"] = @"\dots ",
            ["⋯"] = @"\cdots ",
            ["⋮"] = @"\vdots ",
            ["⋱"] = @"\ddots ",
            ["♠"] = @"\spadesuit ",
            ["♣"] = @"\clubsuit ",
            ["♡"] = @"\heartsuit ",
            ["♢"] = @"\diamondsuit ",
            ["α"] = @"\alpha ",
            ["β"] = @"\beta ",
            ["γ"] = @"\gamma ",
            ["δ"] = @"\delta ",
            // MathJax follows TeX glyph naming rather than Unicode's character
            // names: \epsilon is U+03F5 (ϵ) while \varepsilon is U+03B5 (ε).
            // Keep this mapping inverse-exact so MathType -> VisualTeX ->
            // MathType cannot silently swap the two epsilon forms.
            ["ε"] = @"\varepsilon ",
            ["ϵ"] = @"\epsilon ",
            ["ζ"] = @"\zeta ",
            ["η"] = @"\eta ",
            ["θ"] = @"\theta ",
            ["ϑ"] = @"\vartheta ",
            ["ι"] = @"\iota ",
            ["κ"] = @"\kappa ",
            ["ϰ"] = @"\varkappa ",
            ["λ"] = @"\lambda ",
            ["μ"] = @"\mu ",
            ["ν"] = @"\nu ",
            ["ξ"] = @"\xi ",
            ["π"] = @"\pi ",
            ["ϖ"] = @"\varpi ",
            ["ρ"] = @"\rho ",
            ["ϱ"] = @"\varrho ",
            ["σ"] = @"\sigma ",
            ["ς"] = @"\varsigma ",
            ["τ"] = @"\tau ",
            ["υ"] = @"\upsilon ",
            // Likewise TeX \phi is U+03D5 (ϕ) and \varphi is U+03C6 (φ).
            ["φ"] = @"\varphi ",
            ["ϕ"] = @"\phi ",
            ["χ"] = @"\chi ",
            ["ψ"] = @"\psi ",
            ["ω"] = @"\omega ",
            ["Γ"] = @"\Gamma ",
            ["Δ"] = @"\Delta ",
            ["Θ"] = @"\Theta ",
            ["Λ"] = @"\Lambda ",
            ["Ξ"] = @"\Xi ",
            ["Π"] = @"\Pi ",
            ["Σ"] = @"\Sigma ",
            ["Υ"] = @"\Upsilon ",
            ["Φ"] = @"\Phi ",
            ["Ψ"] = @"\Psi ",
            ["Ω"] = @"\Omega ",
        };

    internal static string Convert(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("Word OMML conversion produced empty MathML.");
        var document = FormulaXml.Parse(mathMl, "MathML", LoadOptions.PreserveWhitespace);
        var root = document.Root
            ?? throw new InvalidDataException("Word MathML has no root element.");
        return Normalize(ConvertElement(root));
    }

    // Word and PowerPoint equations stay pure Office objects with no VisualTeX
    // source stored in them, so the LaTeX read back from one is all the user
    // gets. Write it the way people type it, without changing its meaning.
    internal static string ConvertWordEquation(string mathMl) =>
        WordEquationSourceStyle(Convert(mathMl));

    private static readonly HashSet<string> SourceStyleTextCommands = new(StringComparer.Ordinal)
    {
        @"\text", @"\textrm", @"\textit", @"\textbf", @"\textsf", @"\texttt",
        @"\textup", @"\textnormal", @"\mbox", @"\operatorname",
    };

    private static readonly Dictionary<string, (string Environment, string Close)> SourceStyleMatrixFences =
        new(StringComparer.Ordinal)
        {
            ["("] = ("pmatrix", ")"),
            ["["] = ("bmatrix", "]"),
            ["|"] = ("vmatrix", "|"),
            [@"\|"] = ("Vmatrix", @"\|"),
            [@"\{"] = ("Bmatrix", @"\}"),
        };

    // 𝑎 -> a (the mathematical alphanumerics PowerPoint stores), x^{2} -> x^2,
    // {x}^2 -> x^2, f^{\prime} -> f', \left(\begin{matrix}..\end{matrix}\right)
    // -> pmatrix (bmatrix, vmatrix, Vmatrix, Bmatrix), \left\{ matrix \right.
    // -> cases, and drops the spaces TeX ignores. Text groups are left alone.
    // The macOS port (omml_latex/source_style.rs) must match this exactly.
    internal static string WordEquationSourceStyle(string latex)
    {
        if (string.IsNullOrEmpty(latex)) return latex;
        var tokens = SourceStyleTokens(latex);
        tokens = SourceStyleAlphanumerics(tokens);
        tokens = SourceStyleMatrices(tokens);
        tokens = SourceStyleScripts(tokens);
        return SourceStyleSpacing(tokens);
    }

    private static bool IsAsciiLetter(char value) =>
        value is >= 'a' and <= 'z' or >= 'A' and <= 'Z';

    private static bool IsAsciiLetterOrDigit(char value) =>
        IsAsciiLetter(value) || value is >= '0' and <= '9';

    private static bool IsSourceSpace(char value) => value is ' ' or '\t' or '\r' or '\n';

    // Control words, control symbols, runs of ASCII letters, runs of spaces
    // (as one " "), and every other character (a whole surrogate pair) on its own.
    private static List<string> SourceStyleTokens(string latex)
    {
        var tokens = new List<string>();
        var index = 0;
        while (index < latex.Length)
        {
            var start = index;
            var current = latex[index];
            if (current == '\\' && index + 1 < latex.Length)
            {
                index++;
                if (IsAsciiLetter(latex[index]))
                    while (index < latex.Length && IsAsciiLetter(latex[index])) index++;
                else
                    index += SourceCharLength(latex, index);
                tokens.Add(latex.Substring(start, index - start));
            }
            else if (IsSourceSpace(current))
            {
                while (index < latex.Length && IsSourceSpace(latex[index])) index++;
                tokens.Add(" ");
            }
            else if (IsAsciiLetter(current))
            {
                while (index < latex.Length && IsAsciiLetter(latex[index])) index++;
                tokens.Add(latex.Substring(start, index - start));
            }
            else
            {
                index += SourceCharLength(latex, index);
                tokens.Add(latex.Substring(start, index - start));
            }
        }
        return tokens;
    }

    private static int SourceCharLength(string text, int index) =>
        index + 1 < text.Length && char.IsSurrogatePair(text[index], text[index + 1]) ? 2 : 1;

    // Unicode mathematical alphanumerics (U+1D400-U+1D7FF), which PowerPoint
    // stores for math letters (𝑎 for an italic a), as LaTeX tokens.
    private static readonly (int Start, string Style)[] SourceStyleLatinBlocks =
    {
        (0x1D400, @"\mathbf"), (0x1D434, ""), (0x1D468, @"\boldsymbol"), (0x1D49C, @"\mathcal"),
        (0x1D4D0, @"\boldsymbol\mathcal"), (0x1D504, @"\mathfrak"), (0x1D538, @"\mathbb"),
        (0x1D56C, @"\boldsymbol\mathfrak"), (0x1D5A0, @"\mathsf"), (0x1D5D4, @"\boldsymbol\mathsf"),
        (0x1D608, @"\mathsf"), (0x1D63C, @"\boldsymbol\mathsf"), (0x1D670, @"\mathtt"),
    };
    private static readonly (int Start, string Style)[] SourceStyleGreekBlocks =
    {
        (0x1D6A8, @"\boldsymbol"), (0x1D6E2, ""), (0x1D71C, @"\boldsymbol"),
        (0x1D756, @"\boldsymbol"), (0x1D790, @"\boldsymbol"),
    };
    private const string SourceStyleGreek =
        "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡϴΣΤΥΦΧΨΩ∇αβγδεζηθικλμνξοπρςστυφχψω∂ϵϑϰϕϱϖ";
    private static readonly (int Start, string Style)[] SourceStyleDigitBlocks =
    {
        (0x1D7CE, @"\mathbf"), (0x1D7D8, @"\mathbb"), (0x1D7E2, @"\mathsf"),
        (0x1D7EC, @"\boldsymbol\mathsf"), (0x1D7F6, @"\mathtt"),
    };

    private static List<string>? SourceStyleMathAlphanumeric(int code)
    {
        static List<string> Styled(string style, List<string> inner)
        {
            var commands = style.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            var output = new List<string>();
            foreach (var command in commands) { output.Add("\\" + command); output.Add("{"); }
            output.AddRange(inner);
            foreach (var _ in commands) output.Add("}");
            return output;
        }
        if (code == 0x210E) return new List<string> { "h" };
        if (code is 0x1D6A4 or 0x1D6A5) return new List<string> { code == 0x1D6A4 ? @"\imath" : @"\jmath" };
        foreach (var (start, style) in SourceStyleLatinBlocks)
            if (code >= start && code < start + 52)
            {
                var offset = code - start;
                var letter = offset < 26 ? (char)('A' + offset) : (char)('a' + offset - 26);
                return Styled(style, new List<string> { letter.ToString() });
            }
        foreach (var (start, style) in SourceStyleGreekBlocks)
            if (code >= start && code < start + 58)
            {
                var inner = SourceStyleTokens(ConvertToken(SourceStyleGreek[code - start].ToString()).Trim())
                    .Where(token => token != " ").ToList();
                return Styled(style, inner);
            }
        foreach (var (start, style) in SourceStyleDigitBlocks)
            if (code >= start && code < start + 10)
                return Styled(style, new List<string> { ((char)('0' + code - start)).ToString() });
        return null;
    }

    private static List<string> SourceStyleAlphanumerics(List<string> tokens)
    {
        var output = new List<string>();
        for (var index = 0; index < tokens.Count; index++)
        {
            var textEnd = SourceTextGroupEnd(tokens, index);
            if (textEnd >= 0)
            {
                output.AddRange(tokens.GetRange(index, textEnd - index + 1));
                index = textEnd;
                continue;
            }
            var token = tokens[index];
            var single = token.Length == 1 && !char.IsSurrogate(token[0])
                || token.Length == 2 && char.IsSurrogatePair(token[0], token[1]);
            var mapped = single ? SourceStyleMathAlphanumeric(char.ConvertToUtf32(token, 0)) : null;
            if (mapped is not null) output.AddRange(mapped);
            else output.Add(token);
        }
        return output;
    }

    private static bool IsControlWord(string token) =>
        token.Length > 1 && token[0] == '\\' && IsAsciiLetter(token[1]);

    private static int SkipSourceSpaces(List<string> tokens, int index)
    {
        while (index < tokens.Count && tokens[index] == " ") index++;
        return index;
    }

    private static int MatchingSourceBrace(List<string> tokens, int open)
    {
        var depth = 0;
        for (var index = open; index < tokens.Count; index++)
        {
            if (tokens[index] == "{") depth++;
            else if (tokens[index] == "}" && --depth == 0) return index;
        }
        return -1;
    }

    // The last token of a text-mode group (\text{..}) starting at index, or -1.
    private static int SourceTextGroupEnd(List<string> tokens, int index)
    {
        if (!SourceStyleTextCommands.Contains(tokens[index])) return -1;
        var open = SkipSourceSpaces(tokens, index + 1);
        if (open < tokens.Count && tokens[open] == "*") open = SkipSourceSpaces(tokens, open + 1);
        return open < tokens.Count && tokens[open] == "{" ? MatchingSourceBrace(tokens, open) : -1;
    }

    private static bool IsSourceEnvironment(List<string> tokens, int index, string command, string name) =>
        index + 3 < tokens.Count && tokens[index] == command && tokens[index + 1] == "{"
        && tokens[index + 2] == name && tokens[index + 3] == "}";

    private static List<string> SourceStyleMatrices(List<string> tokens)
    {
        var output = new List<string>();
        for (var index = 0; index < tokens.Count; index++)
        {
            var textEnd = SourceTextGroupEnd(tokens, index);
            if (textEnd >= 0)
            {
                output.AddRange(tokens.GetRange(index, textEnd - index + 1));
                index = textEnd;
                continue;
            }
            if (tokens[index] == @"\left" && TrySourceNamedMatrix(tokens, index, out var named, out var last))
            {
                output.AddRange(named);
                index = last;
                continue;
            }
            output.Add(tokens[index]);
        }
        return output;
    }

    private static bool TrySourceNamedMatrix(
        List<string> tokens, int left, out List<string> named, out int last)
    {
        named = new List<string>();
        last = left;
        var fence = SkipSourceSpaces(tokens, left + 1);
        if (fence >= tokens.Count || !SourceStyleMatrixFences.TryGetValue(tokens[fence], out var kind))
            return false;
        var begin = SkipSourceSpaces(tokens, fence + 1);
        if (!IsSourceEnvironment(tokens, begin, @"\begin", "matrix")) return false;
        var depth = 0;
        var end = -1;
        for (var index = begin; index < tokens.Count; index++)
        {
            if (IsSourceEnvironment(tokens, index, @"\begin", "matrix")) depth++;
            else if (IsSourceEnvironment(tokens, index, @"\end", "matrix") && --depth == 0)
            {
                end = index;
                break;
            }
        }
        if (end < 0) return false;
        var right = SkipSourceSpaces(tokens, end + 4);
        if (right >= tokens.Count || tokens[right] != @"\right") return false;
        var close = SkipSourceSpaces(tokens, right + 1);
        if (close >= tokens.Count) return false;
        string environment;
        if (tokens[close] == kind.Close) environment = kind.Environment;
        else if (tokens[fence] == @"\{" && tokens[close] == ".") environment = "cases";
        else return false;
        named.AddRange(new[] { @"\begin", "{", environment, "}" });
        named.AddRange(SourceStyleMatrices(tokens.GetRange(begin + 4, end - begin - 4)));
        named.AddRange(new[] { @"\end", "{", environment, "}" });
        last = close;
        return true;
    }

    private static List<string> SourceStyleScripts(List<string> tokens)
    {
        var output = new List<string>();
        for (var index = 0; index < tokens.Count; index++)
        {
            var textEnd = SourceTextGroupEnd(tokens, index);
            if (textEnd >= 0)
            {
                output.AddRange(tokens.GetRange(index, textEnd - index + 1));
                index = textEnd;
                continue;
            }
            var token = tokens[index];
            // {x}^2 -> x^2 (PowerPoint letters arrive braced); a command's own
            // argument (\mathrm{x}^2, \frac{a}{x}^2) keeps its braces.
            if (token == "{" && index + 3 < tokens.Count && tokens[index + 2] == "}"
                && tokens[index + 3] is "^" or "_"
                && tokens[index + 1].Length == 1 && IsAsciiLetterOrDigit(tokens[index + 1][0]))
            {
                var before = output.LastOrDefault(item => item != " ");
                if (before is null || !IsControlWord(before) && before is not ("}" or "]" or "^" or "_"))
                {
                    output.Add(tokens[index + 1]);
                    index += 2;
                    continue;
                }
            }
            if (token is "^" or "_" && index + 1 < tokens.Count && tokens[index + 1] == "{")
            {
                var close = MatchingSourceBrace(tokens, index + 1);
                var inner = close < 0
                    ? new List<string>()
                    : tokens.GetRange(index + 2, close - index - 2).Where(item => item != " ").ToList();
                // x^{2} -> x^2, keeping a space before a following letter or digit.
                if (inner.Count == 1 && inner[0].Length == 1 && IsAsciiLetterOrDigit(inner[0][0]))
                {
                    output.Add(token);
                    output.Add(inner[0]);
                    if (close + 1 < tokens.Count && IsAsciiLetterOrDigit(tokens[close + 1][0]))
                        output.Add(" ");
                    index = close;
                    continue;
                }
                // f^{\prime} -> f', the same thing in TeX.
                var previous = output.LastOrDefault(item => item != " ");
                if (token == "^" && inner.Count > 0 && inner.All(item => item == @"\prime")
                    && previous is not null and not ("{" or "&" or "^" or "_" or @"\\"))
                {
                    output.Add(new string('\'', inner.Count));
                    index = close;
                    continue;
                }
            }
            output.Add(token);
        }
        return output;
    }

    private static string SourceStyleSpacing(List<string> tokens)
    {
        var builder = new StringBuilder();
        string? previous = null;
        string? beforePrevious = null;
        for (var index = 0; index < tokens.Count; index++)
        {
            var textEnd = SourceTextGroupEnd(tokens, index);
            if (textEnd >= 0)
            {
                for (var item = index; item <= textEnd; item++) builder.Append(tokens[item]);
                beforePrevious = previous;
                previous = "}";
                index = textEnd;
                continue;
            }
            var token = tokens[index];
            if (token != " ")
            {
                // A control word or a bare script needs a space before a letter
                // that a mapped PowerPoint letter put right after it (\sin x).
                if (previous is not null && (builder.Length == 0 || builder[builder.Length - 1] != ' ')
                    && SourceNeedsSpace(previous, beforePrevious, token))
                    builder.Append(' ');
                builder.Append(token);
                beforePrevious = previous;
                previous = token;
                continue;
            }
            var next = index + 1 < tokens.Count ? tokens[index + 1] : null;
            if (previous is null || next is null) continue;
            // A space ends a control word before a letter (\sin x) and a bare
            // script before one (\int_0^1 x); around & and \\ it keeps
            // matrices and aligned rows readable.
            var keep = IsControlWord(previous) && (IsAsciiLetterOrDigit(next[0]) || next == "*")
                || beforePrevious is "^" or "_" && previous.Length == 1
                    && IsAsciiLetterOrDigit(previous[0]) && IsAsciiLetterOrDigit(next[0])
                || previous is "&" or @"\\" || next is "&" or @"\\";
            if (keep) builder.Append(' ');
        }
        return builder.ToString();
    }

    private static bool SourceNeedsSpace(string previous, string? beforePrevious, string next) =>
        IsControlWord(previous) && IsAsciiLetter(next[0])
        || beforePrevious is "^" or "_" && previous.Length == 1
            && IsAsciiLetterOrDigit(previous[0]) && IsAsciiLetterOrDigit(next[0]);

    private static string ConvertElement(XElement element) =>
        MathMlColorContract.Wrap(ConvertElementContent(element), MathMlColorContract.Explicit(element));

    private static string ConvertElementContent(XElement element)
    {
        var name = element.Name.LocalName;
        return name switch
        {
            "math" => ConvertChildren(element),
            "mrow" or "mstyle" or "mpadded" => ConvertClassedContainer(element),
            "semantics" or "maction" => element.Elements().Select(ConvertElement).FirstOrDefault() ?? string.Empty,
            "annotation" or "annotation-xml" => string.Empty,
            "mi" or "mn" or "mo" => ConvertTokenElement(element),
            "mtext" => ConvertTextElement(element),
            "mspace" => ConvertSpace(element),
            "mfrac" => ConvertFraction(element),
            "msqrt" => @"\sqrt{" + ConvertChildren(element) + "}",
            "mroot" => ConvertRoot(element),
            "msup" => ConvertScript(element, subscript: false, superscript: true),
            "msub" => ConvertScript(element, subscript: true, superscript: false),
            "msubsup" => ConvertSubSup(element),
            "mover" => ConvertOver(element),
            "munder" => ConvertUnder(element),
            "munderover" => ConvertUnderOver(element),
            "mfenced" => ConvertFenced(element),
            "mtable" => ConvertTable(element),
            "mtr" or "mlabeledtr" => string.Join(" & ", element.Elements().Select(ConvertElement)),
            "mtd" => ConvertChildren(element),
            "menclose" => ConvertEnclose(element),
            "mphantom" => @"\phantom{" + ConvertChildren(element) + "}",
            "mmultiscripts" => ConvertMultiScripts(element),
            "malignmark" or "maligngroup" => alignmentMarkDepth > 0 ? "&" : string.Empty,
            "none" => string.Empty,
            _ => ConvertChildren(element),
        };
    }

    private static string ConvertSpace(XElement element)
    {
        var space = MathSpaceContract.ParseWidth((string?)element.Attribute("width"));
        // Recover every exactly supported source dimension using the shared
        // contract. Unrepresented/context-dependent widths retain the legacy
        // behavior for now; that fallback is not a claim of width fidelity.
        return space.LatexSource ?? @"\,";
    }

    private static string? ExplicitAtomClassCommand(XElement element) =>
        ((string?)element.Attribute("data-mjx-texclass"))?.Trim().ToUpperInvariant() switch
        {
            "REL" => @"\mathrel",
            "BIN" => @"\mathbin",
            "OP" => @"\mathop",
            _ => null,
        };

    private static string ConvertClassedContainer(XElement element)
    {
        var content = ConvertChildren(element);
        var command = ExplicitAtomClassCommand(element);
        // Preserve explicit atom semantics, never infer a package command from
        // the glyph. ORD and delimiter wrappers keep their established paths.
        return command is null || string.IsNullOrWhiteSpace(content)
            ? content
            : command + "{" + content.Trim() + "}";
    }

    private static string ConvertChildren(XElement element)
    {
        var children = element.Elements().ToArray();
        var builder = new StringBuilder();
        for (var index = 0; index < children.Length;)
        {
            if (index + 2 < children.Length
                && IsSimpleOperator(children[index], "(")
                && IsSimpleOperator(children[index + 2], ")")
                && TryConvertLooseBinomialBody(children[index + 1], out var binomial))
            {
                builder.Append(binomial);
                index += 3;
                continue;
            }
            // A known function name in a real application slot (U+2061, Word's
            // m:func) is TeX's \sin: MathJax emits the U+2061 again for the
            // command, so the slot survives without an invisible character in
            // the source. Unknown names keep U+2061 as their only carrier.
            if (index + 1 < children.Length
                && IsFunctionApplication(children[index + 1])
                && TryConvertAppliedNamedFunction(children[index], out var applied))
            {
                builder.Append(applied);
                index += 2;
                continue;
            }
            // Word keeps an upright word such as "Tr" in one run, which the
            // projection splits into letters; write it as one \mathrm group.
            var run = index;
            while (run < children.Length && IsPlainUprightLatinLetter(children[run])) run++;
            if (run - index > 1 && !(run < children.Length && IsFunctionApplication(children[run])))
            {
                builder.Append(@"\mathrm{")
                    .Append(string.Concat(children.Skip(index).Take(run - index).Select(child => child.Value.Trim())))
                    .Append('}');
                index = run;
                continue;
            }
            builder.Append(ConvertElement(children[index]));
            index++;
        }
        return builder.ToString();
    }

    private static bool IsPlainUprightLatinLetter(XElement element)
    {
        if (element.Name.LocalName != "mi") return false;
        var token = element.Value.Trim();
        return token.Length == 1 && token[0] is >= 'A' and <= 'Z' or >= 'a' and <= 'z'
            && element.Attributes().All(attribute => attribute.Name.LocalName == "mathvariant")
            && MathMlTokenStyleContract.ResolveEffectiveVariant(
                element, MathMlTokenStyleContract.ResolveDeclaredVariant(element)) is "normal" or "upright";
    }

    private static bool IsFunctionApplication(XElement element)
    {
        element = UnwrapSingleTransparentMathMlContainer(element);
        return element.Name.LocalName == "mo"
            && element.Value.Trim() == "⁡"
            && MathMlColorContract.Explicit(element) is null;
    }

    private static bool TryConvertAppliedNamedFunction(XElement head, out string latex)
    {
        latex = string.Empty;
        var unwrapped = UnwrapSingleTransparentMathMlContainer(head);
        if (head.DescendantsAndSelf().Any(item => MathMlColorContract.Explicit(item) is not null))
            return false;
        if (TryGetNamedFunctionIdentifier(unwrapped, out var command))
        {
            latex = command;
            return true;
        }
        var kind = unwrapped.Name.LocalName;
        // Word writes a limit function as m:func around m:limLow.
        if (kind is "munder" or "munderover"
            && unwrapped.Elements().FirstOrDefault() is { } limitBase
            && TryGetNamedFunctionIdentifier(UnwrapSingleTransparentMathMlContainer(limitBase), out command)
            && IsLimitLikeOperatorLatex(command))
        {
            latex = ConvertElement(unwrapped) + " ";
            return true;
        }
        // A function name slot may hold more than the name (e.g. a sum before
        // it); the name applied to the argument is the slot's last item.
        if (kind == "mrow" && ExplicitAtomClassCommand(unwrapped) is null)
        {
            var items = unwrapped.Elements().ToArray();
            if (items.Length >= 2 && TryConvertAppliedNamedFunction(items[items.Length - 1], out var tail))
            {
                latex = string.Concat(items.Take(items.Length - 1).Select(ConvertElement)) + tail;
                return true;
            }
            return false;
        }
        if (kind is not ("msup" or "msub" or "msubsup")) return false;
        var parts = unwrapped.Elements().ToArray();
        if (parts.Length < (kind == "msubsup" ? 3 : 2)
            || !TryGetNamedFunctionIdentifier(UnwrapSingleTransparentMathMlContainer(parts[0]), out command))
            return false;
        latex = command.TrimEnd() + kind switch
        {
            "msub" => "_{" + ScriptArgument(parts[1]) + "}",
            "msup" => "^{" + ScriptArgument(parts[1]) + "}",
            _ => "_{" + ScriptArgument(parts[1]) + "}^{" + ScriptArgument(parts[2]) + "}",
        } + " ";
        return true;
    }

    // An upright multi-letter operator identifier whose name TeX has a command
    // for (sin, log, lim, …). Explicit ORD/non-operator roles never qualify.
    private static bool TryGetNamedFunctionIdentifier(XElement element, out string command)
    {
        command = string.Empty;
        if (element.Name.LocalName != "mi") return false;
        var token = element.Value.Trim();
        if (token.Length < 2 || !token.All(c => c is >= 'a' and <= 'z')) return false;
        var declared = MathMlTokenStyleContract.ResolveDeclaredVariant(element);
        if (MathMlTokenStyleContract.ResolveEffectiveVariant(element, declared) is not ("normal" or "upright"))
            return false;
        var atomClass = ((string?)element.Attribute("data-mjx-texclass"))?.Trim().ToUpperInvariant();
        var isOperator = atomClass == "OP" || (atomClass is null
            && (string?)element.Attribute("data-mjx-auto-op") != "false");
        return isOperator && TryConvertNamedOperator(token, out command);
    }

    private static bool TryConvertLooseBinomialBody(XElement body, out string latex)
    {
        latex = string.Empty;
        body = UnwrapSingleTransparentMathMlContainer(body);
        if (body.Name.LocalName == "mfrac")
        {
            var lineThickness = ((string?)body.Attribute("linethickness") ?? string.Empty).Trim();
            var fractionChildren = body.Elements().ToArray();
            if (!lineThickness.StartsWith("0", StringComparison.Ordinal)
                || fractionChildren.Length < 2)
                return false;
            latex = @"\binom{" + ConvertElement(fractionChildren[0]) + "}{"
                + ConvertElement(fractionChildren[1]) + "}";
            return true;
        }

        if (body.Name.LocalName != "mtable"
            || !string.Equals(
                (string?)body.Attribute("data-mtef-pile"),
                "true",
                StringComparison.OrdinalIgnoreCase))
            return false;
        var rows = body.Elements()
            .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
            .ToArray();
        if (rows.Length != 2) return false;
        var cells = rows.Select(row => row.Elements()
                .Where(cell => cell.Name.LocalName == "mtd")
                .ToArray())
            .ToArray();
        if (cells.Any(row => row.Length != 1)) return false;
        latex = @"\binom{" + ConvertElement(cells[0][0]) + "}{"
            + ConvertElement(cells[1][0]) + "}";
        return true;
    }

    private static bool IsSimpleOperator(XElement element, string value)
    {
        element = UnwrapSingleTransparentMathMlContainer(element);
        return element.Name.LocalName == "mo"
            && string.Equals(element.Value.Trim(), value, StringComparison.Ordinal);
    }

    private static string ConvertFraction(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return ConvertChildren(element);
        // A zero-width rule is a stacked pair without a bar (\atop, an m:f of
        // type noBar). \binom is recognized with its fence elsewhere; the bare
        // form must not come back with a rule, or an edit would draw one.
        var thickness = ((string?)element.Attribute("linethickness") ?? string.Empty).Trim();
        if (thickness.StartsWith("0", StringComparison.Ordinal))
            return @"\genfrac{}{}{0pt}{}{" + ConvertElement(children[0]) + "}{" + ConvertElement(children[1]) + "}";
        // \tfrac and \dfrac: the fraction alone in an mstyle setting displaystyle.
        var command = @"\frac";
        if (element.Parent is { } style && style.Name.LocalName == "mstyle" && style.Elements().Count() == 1)
        {
            var displayStyle = ((string?)style.Attribute("displaystyle"))?.Trim();
            if (displayStyle == "false") command = @"\tfrac";
            else if (displayStyle == "true") command = @"\dfrac";
        }
        return command + "{" + ConvertElement(children[0]) + "}{" + ConvertElement(children[1]) + "}";
    }

    private static string ConvertRoot(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return @"\sqrt{" + ConvertChildren(element) + "}";
        return @"\sqrt[" + ConvertElement(children[1]) + "]{" + ConvertElement(children[0]) + "}";
    }

    private static string ConvertScript(XElement element, bool subscript, bool superscript)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return ConvertChildren(element);
        var result = GroupBase(ConvertElement(children[0]));
        // A bare prime superscript is TeX's f' (MathJax reads it back as the
        // same bare prime, MathType's prime embellishment); f^{\prime} is an mi.
        if (superscript && children[1].Name.LocalName == "mo"
            && children[1].Value.Trim() is { Length: > 0 } primes
            && primes.All(prime => prime is '′' or '″' or '‴'))
            return result + new string('\'', primes.Sum(prime => prime - 0x2031)); // ′ is one prime
        if (subscript) result += "_{" + ScriptArgument(children[1]) + "}";
        if (superscript) result += "^{" + ScriptArgument(children[1]) + "}";
        return result;
    }

    private static string ConvertSubSup(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 3) return ConvertChildren(element);
        return GroupBase(ConvertElement(children[0]))
            + "_{" + ScriptArgument(children[1]) + "}"
            + "^{" + ScriptArgument(children[2]) + "}";
    }

    private static string ScriptArgument(XElement element) =>
        ConvertElement(element).TrimEnd();

    private static string ConvertOver(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return ConvertChildren(element);
        if (TryConvertAnnotatedHorizontalBrace(children[0], children[1], top: true, out var brace))
            return brace;
        var body = ConvertElement(children[0]);
        var over = children[1].Value.Trim();
        var stretchy = MathMlAccentGeometry.IsStretchy(children[1]);
        var accentStyle = over is "¯" or "‾" or "―" or "→" or "←" or "↔"
            or "^" or "ˆ" or "~" or "˜" or "." or "˙" or "¨"
            ? MathMlTokenStyleContract.ResolveEffectiveVariant(children[1]) : null;
        var styleCommand = accentStyle switch
        {
            "bold" => @"\mathbf",
            "italic" => @"\mathit",
            "bold-italic" => @"\mathbfit",
            _ => null,
        };
        // The mark owns this scope; reset its body before emitting the body's
        // own styles so the accent font cannot restyle its operands.
        if (styleCommand is not null) body = @"\mathnormal{" + body + "}";
        var markColor = MathMlColorContract.Effective(children[1]);
        var parentColor = MathMlColorContract.Effective(element);
        var independentColor = markColor is not null && markColor != parentColor;
        if (independentColor)
            body = MathMlColorContract.Wrap(body, MathMlColorContract.Effective(children[0]) ?? "#000000");
        var fixedAccent = IsFixedAccentOverOneCharacter(element, children[0], children[1]);
        var result = over switch
        {
            "¯" or "‾" or "―" => (fixedAccent ? @"\bar{" : @"\overline{") + body + "}",
            "→" => (stretchy && !fixedAccent ? @"\overrightarrow{" : @"\vec{") + body + "}",
            "←" => @"\overleftarrow{" + body + "}",
            "↔" => @"\overleftrightarrow{" + body + "}",
            "^" or "ˆ" => (stretchy && !fixedAccent ? @"\widehat{" : @"\hat{") + body + "}",
            "~" or "˜" => (stretchy && !fixedAccent ? @"\widetilde{" : @"\tilde{") + body + "}",
            "." or "˙" => @"\dot{" + body + "}",
            "¨" => @"\ddot{" + body + "}",
            "⃛" => @"\dddot{" + body + "}",
            "⃜" => @"\ddddot{" + body + "}",
            "ˇ" => @"\check{" + body + "}",
            "´" => @"\acute{" + body + "}",
            "`" => @"\grave{" + body + "}",
            "˘" => @"\breve{" + body + "}",
            "˚" => @"\mathring{" + body + "}",
            "⏞" or "\uFE37" => @"\overbrace{" + body + "}",
            _ => @"\overset{" + ConvertElement(children[1]) + "}{" + body + "}",
        };
        if (styleCommand is not null) result = styleCommand + "{" + result + "}";
        return MathMlColorContract.Wrap(result, independentColor ? markColor : null);
    }

    // Word's accent object (m:acc, MathML accent="true") over one character is
    // the fixed accent (\vec, \hat, \tilde, \bar). A declared stretchy mark,
    // or a wider base, keeps the stretched form (\overrightarrow, \widehat, …).
    private static bool IsFixedAccentOverOneCharacter(XElement over, XElement baseElement, XElement mark)
    {
        if (!string.Equals(((string?)over.Attribute("accent"))?.Trim(), "true", StringComparison.OrdinalIgnoreCase))
            return false;
        if (mark.Attribute("stretchy") is not null) return false;
        var token = UnwrapSingleTransparentMathMlContainer(baseElement);
        return token.Name.LocalName is "mi" or "mn" or "mo"
            && new System.Globalization.StringInfo(token.Value.Trim()).LengthInTextElements == 1;
    }

    private static string ConvertUnder(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return ConvertChildren(element);
        if (TryConvertAnnotatedHorizontalBrace(children[0], children[1], top: false, out var brace))
            return brace;
        // Word's limit (m:limLow over an upright lim, max, …) keeps its limit
        // underneath even inline; \limits states that while staying a real
        // \lim. MathJax's own \lim_{} (an mo) keeps the movable form below.
        var limitBase = UnwrapSingleTransparentMathMlContainer(children[0]);
        if (MathMlColorContract.Explicit(limitBase) is null
            && (TryGetNamedFunctionIdentifier(limitBase, out var limitCommand)
                || (limitBase.Name.LocalName == "mo"
                    && string.Equals(((string?)limitBase.Attribute("movablelimits"))?.Trim(), "false", StringComparison.OrdinalIgnoreCase)
                    && TryConvertNamedOperator(limitBase.Value.Trim(), out limitCommand)))
            && IsLimitLikeOperatorLatex(limitCommand))
            return limitCommand.TrimEnd() + @"\limits_{" + ScriptArgument(children[1]) + "}";
        // Word's n-ary with only a lower limit is TeX's \sum_{k}, the same
        // spelling as the two-limit form (munderover) below.
        if (limitBase.Name.LocalName == "mo" && MathMlColorContract.Explicit(limitBase) is null
            && IsLargeOperatorGlyph(limitBase.Value.Trim()))
            return GroupBase(LargeOperatorBaseLatex(children[0])) + "_{" + ScriptArgument(children[1]) + "}";
        var body = ConvertElement(children[0]);
        var under = children[1].Value.Trim();
        if (IsLimitLikeOperatorLatex(body))
            return GroupBase(body) + "_{" + ScriptArgument(children[1]) + "}";
        return under switch
        {
            "_" or "¯" or "‾" or "―" => @"\underline{" + body + "}",
            "⏟" or "\uFE38" => @"\underbrace{" + body + "}",
            _ => @"\underset{" + ConvertElement(children[1]) + "}{" + body + "}",
        };
    }

    private static bool TryConvertAnnotatedHorizontalBrace(
        XElement baseElement,
        XElement annotation,
        bool top,
        out string latex)
    {
        latex = string.Empty;
        baseElement = UnwrapSingleTransparentMathMlContainer(baseElement);
        if (baseElement.Name.LocalName != (top ? "mover" : "munder")) return false;
        var inner = baseElement.Elements().ToList();
        if (inner.Count < 2) return false;
        var expectedMark = top ? "⏞" : "⏟";
        var presentationMark = top ? "\uFE37" : "\uFE38";
        var marker = inner[1].Value.Trim();
        var replacementOnly = marker.Length > 0
            && marker.All(character => character == '\uFFFD');
        var stretchy = string.Equals(
            (string?)inner[1].Attribute("stretchy"),
            "true",
            StringComparison.OrdinalIgnoreCase);
        // MathType 7 may export horizontal braces as the Unicode vertical
        // presentation forms U+FE37/U+FE38, or as replacement glyphs on some
        // machines. The nested mover/munder shape plus stretchy operator remains
        // the stable signal for recovering an annotated brace.
        if (marker != expectedMark
            && marker != presentationMark
            && !(replacementOnly && stretchy)) return false;
        var body = ConvertElement(inner[0]);
        var note = ConvertElement(annotation).TrimEnd();
        latex = top
            ? @"\overbrace{" + body + "}^{" + note + "}"
            : @"\underbrace{" + body + "}_{" + note + "}";
        return true;
    }

    private static XElement UnwrapSingleTransparentMathMlContainer(XElement element)
    {
        var current = element;
        while (current.Name.LocalName is "mrow" or "mstyle" or "mpadded")
        {
            if (ExplicitAtomClassCommand(current) is not null) break;
            var children = current.Elements().ToArray();
            if (children.Length != 1) break;
            current = children[0];
        }
        return current;
    }

    private static string ConvertUnderOver(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 3) return ConvertChildren(element);
        var baseLatex = LargeOperatorBaseLatex(children[0]);
        return GroupBase(baseLatex)
            + "_{" + ScriptArgument(children[1]) + "}"
            + "^{" + ScriptArgument(children[2]) + "}";
    }

    // Word writes an n-ary union/intersection with the binary glyph.
    private static string LargeOperatorBaseLatex(XElement element) =>
        element.Name.LocalName == "mo"
            ? element.Value.Trim() switch
            {
                "∪" => @"\bigcup ",
                "∩" => @"\bigcap ",
                _ => ConvertElement(element),
            }
            : ConvertElement(element);

    private static bool IsLargeOperatorGlyph(string glyph) =>
        glyph is "∑" or "∏" or "∐" or "⋃" or "⋂" or "∪" or "∩" or "⋁" or "⋀"
            or "⨀" or "⨁" or "⨂" or "⨄" or "⨆" or "∫" or "∬" or "∭" or "∮" or "∯" or "∰";

    private static string ConvertFenced(XElement element)
    {
        var open = (string?)element.Attribute("open") ?? "(";
        var close = (string?)element.Attribute("close") ?? ")";
        if (TryConvertFencedNoBarFractionToBinomial(element, open, close, out var noBarBinomial))
            return noBarBinomial;
        if (TryConvertMathTypeBinomialPile(element, open, close, out var binomial))
            return binomial;
        var separators = (string?)element.Attribute("separators") ?? ",";
        var parts = element.Elements().Select(ConvertElement).ToList();
        var separator = separators.Length > 0 ? separators[0].ToString() : string.Empty;
        return Delimiter(open, left: true)
            + string.Join(separator, parts)
            + Delimiter(close, left: false);
    }

    private static bool TryConvertFencedNoBarFractionToBinomial(
        XElement fenced,
        string open,
        string close,
        out string latex)
    {
        latex = string.Empty;
        if (open != "(" || close != ")") return false;
        var children = fenced.Elements().Take(2).ToArray();
        if (children.Length != 1) return false;
        var child = UnwrapSingleTransparentMathMlContainer(children[0]);
        if (child.Name.LocalName != "mfrac") return false;
        var lineThickness = ((string?)child.Attribute("linethickness") ?? string.Empty).Trim();
        if (!lineThickness.StartsWith("0", StringComparison.Ordinal)) return false;
        var fractionChildren = child.Elements().ToArray();
        if (fractionChildren.Length < 2) return false;
        latex = @"\binom{" + ConvertElement(fractionChildren[0]) + "}{"
            + ConvertElement(fractionChildren[1]) + "}";
        return true;
    }

    private static bool TryConvertMathTypeBinomialPile(
        XElement fenced,
        string open,
        string close,
        out string latex)
    {
        latex = string.Empty;
        if (open != "(" || close != ")") return false;
        var fencedChildren = fenced.Elements().Take(2).ToArray();
        if (fencedChildren.Length != 1) return false;
        var table = fencedChildren[0];
        if (table.Name.LocalName == "mrow")
        {
            var nested = table.Elements().ToArray();
            if (nested.Length == 1 && nested[0].Name.LocalName == "mtable")
                table = nested[0];
        }
        if (table is null || table.Name.LocalName != "mtable") return false;
        if (!string.Equals(
                (string?)table.Attribute("data-mtef-pile"),
                "true",
                StringComparison.OrdinalIgnoreCase))
            return false;
        var rows = table.Elements()
            .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
            .ToArray();
        if (rows.Length != 2) return false;
        var cells = rows.Select(row => row.Elements()
                .Where(cell => cell.Name.LocalName == "mtd")
                .ToArray())
            .ToArray();
        if (cells.Any(row => row.Length != 1)) return false;
        latex = @"\binom{" + ConvertElement(cells[0][0]) + "}{"
            + ConvertElement(cells[1][0]) + "}";
        return true;
    }

    [ThreadStatic] private static int alignmentMarkDepth;

    private static IEnumerable<XElement> TableRows(XElement table) =>
        table.Elements().Where(row => row.Name.LocalName is "mtr" or "mlabeledtr");

    // `aligned` reaches Word as a matrix whose columns alternate right/left
    // (mcJc); OmmlMathMlProjection records every column in data-vt-columnalign.
    private static bool IsAlignedColumnPattern(XElement table)
    {
        var columns = ((string?)table.Attribute("data-vt-columnalign") ?? string.Empty)
            .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (columns.Length < 2) return false;
        for (var index = 0; index < columns.Length; index++)
            if (columns[index] != (index % 2 == 0 ? "right" : "left")) return false;
        return true;
    }

    private static string ConvertTable(XElement element)
    {
        if (IsAlignedColumnPattern(element))
        {
            var alignedRows = TableRows(element)
                .Select(row => string.Join(" & ", row.Elements().Select(ConvertElement)).Trim());
            return @"\begin{aligned}" + string.Join(@" \\ ", alignedRows) + @"\end{aligned}";
        }
        // Word's equation array (m:eqArr) marks its alignment points with '&'
        // in the text, which arrives as malignmark/maligngroup. Each row starts
        // with a group of its own; the following marks are the `&` of `aligned`.
        if (element.Descendants().Any(item => item.Name.LocalName == "malignmark"))
        {
            alignmentMarkDepth++;
            try
            {
                var markedRows = TableRows(element)
                    .Select(row => string.Join(" & ", row.Elements().Select(ConvertElement)).Trim())
                    .Select(row => row.StartsWith("&", StringComparison.Ordinal) ? row.Substring(1).TrimStart() : row);
                return @"\begin{aligned}" + string.Join(@" \\ ", markedRows) + @"\end{aligned}";
            }
            finally
            {
                alignmentMarkDepth--;
            }
        }
        var rows = element.Elements()
            .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
            .Select(row => string.Join(" & ", row.Elements().Select(ConvertElement)))
            .ToList();
        if (string.Equals(
                (string?)element.Attribute("data-mtef-tabs"),
                "true",
                StringComparison.OrdinalIgnoreCase))
        {
            // MathType stores native multi-point equation alignment as PILE rows
            // separated by fnMARKER tab characters. The MTEF reader reconstructs
            // those tabs as table cells and marks the private provenance here so
            // editing/conversion returns to `aligned` instead of degrading to a
            // plain matrix and losing the user's `&`/`&&` alignment semantics.
            return @"\begin{aligned}" + string.Join(@" \\ ", rows) + @"\end{aligned}";
        }
        return @"\begin{matrix}" + string.Join(@" \\ ", rows) + @"\end{matrix}";
    }

    private static string ConvertEnclose(XElement element)
    {
        var notation = ((string?)element.Attribute("notation") ?? string.Empty).ToLowerInvariant();
        var body = ConvertChildren(element);
        if (notation.Contains("box")) return @"\boxed{" + body + "}";
        if (notation.Contains("radical")) return @"\sqrt{" + body + "}";
        var up = notation.Contains("updiagonalstrike");
        var down = notation.Contains("downdiagonalstrike");
        if (up && down) return @"\xcancel{" + body + "}";
        if (up) return @"\cancel{" + body + "}";
        if (down) return @"\bcancel{" + body + "}";
        // MathType's mid bar and horizontal strike template.
        if (notation.Contains("horizontalstrike")) return @"\enclose{horizontalstrike}{" + body + "}";
        return body;
    }

    private static string ConvertMultiScripts(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count == 0) return string.Empty;
        var builder = new StringBuilder(GroupBase(ConvertElement(children[0])));
        var index = 1;
        var postPair = 0;
        while (index < children.Count && children[index].Name.LocalName != "mprescripts")
        {
            var sub = ScriptArgument(children[index]);
            var sup = index + 1 < children.Count && children[index + 1].Name.LocalName != "mprescripts"
                ? ScriptArgument(children[index + 1]) : string.Empty;
            // Each additional pair is an adjacent script column, not a second
            // superscript on the same TeX atom (which would be invalid TeX).
            if (postPair++ > 0 && (sub.Length > 0 || sup.Length > 0)) builder.Append("{}");
            if (sub.Length > 0) builder.Append("_{").Append(sub).Append('}');
            if (sup.Length > 0) builder.Append("^{").Append(sup).Append('}');
            index += 2;
        }
        if (index < children.Count && children[index].Name.LocalName == "mprescripts")
        {
            var prescripts = new List<string>();
            for (index++; index < children.Count; index += 2)
            {
                var sub = ScriptArgument(children[index]);
                var sup = index + 1 < children.Count ? ScriptArgument(children[index + 1]) : string.Empty;
                if (sub.Length == 0 && sup.Length == 0) continue;
                prescripts.Add("{}"
                    + (sub.Length > 0 ? "_{" + sub + "}" : string.Empty)
                    + (sup.Length > 0 ? "^{" + sup + "}" : string.Empty));
            }
            // MathML lists prescript columns from the base outward. TeX is
            // emitted from left to right, so those columns precede it in reverse.
            prescripts.Reverse();
            builder.Insert(0, string.Concat(prescripts));
        }
        return builder.ToString();
    }

    private static string ConvertTextElement(XElement element)
    {
        var value = element.Value.Trim();
        // Word exports one-character upright mathematical identifiers such as
        // \mathrm{e}, \mathrm{i} and \mathrm{d} as mtext rather than mi.
        // Preserve their mathematical upright semantics while keeping genuine
        // prose and multi-character annotations as \text{...}.
        if (value.Length == 1 && value[0] <= '\u024F' && char.IsLetter(value[0]))
            return @"\mathrm{" + EscapeMathIdentifier(value) + "}";
        // A lone geometry symbol stored as text (older VisualTeX equations put
        // \triangle in MathType's Text face) is the math symbol; inside \text{}
        // the editor's text font has no glyph for it.
        if (value.Length == 1 && GeometryTextSymbols.Contains(value[0]))
            return ConvertToken(value);
        return @"\text{" + EscapeText(element.Value) + "}";
    }

    private static readonly HashSet<char> GeometryTextSymbols = new()
    {
        '△', '▵', '▲', '▴', '▽', '▿', '▼', '▾', '□', '◻', '■', '◼', '★',
        '∠', '∡', '∵', '∴', '⊥', '∥', '≅',
    };

    private static string ConvertTokenElement(XElement element)
    {
        var token = element.Value.Trim();
        var declaredVariant = MathMlTokenStyleContract.ResolveDeclaredVariant(element);
        var variant = declaredVariant ?? string.Empty;
        var effectiveVariant = MathMlTokenStyleContract.ResolveEffectiveVariant(element, declaredVariant);
        // Source spelling is an adapter decision: some named TeX symbols have
        // upright intrinsic defaults, so materialize their effective mi italic.
        // Ordinary Latin and Greek retain their established command spelling.
        if (declaredVariant is null && effectiveVariant == "italic"
            && char.GetUnicodeCategory(token, 0) == System.Globalization.UnicodeCategory.MathSymbol)
            variant = effectiveVariant;
        var nativeCharacterStyle = int.TryParse(
            (string?)element.Attribute("data-mtef-character-style"),
            out var parsedNativeCharacterStyle)
            ? parsedNativeCharacterStyle
            : 0;
        if (element.Name.LocalName == "mi" && token.Length > 1
            && token.All(c => c is >= 'A' and <= 'Z' or >= 'a' and <= 'z')
            && effectiveVariant is "normal" or "upright")
        {
            var atomClass = ((string?)element.Attribute("data-mjx-texclass"))?.Trim().ToUpperInvariant();
            var isOperator = atomClass == "OP" || (atomClass is null
                && (string?)element.Attribute("data-mjx-auto-op") != "false");
            // Preserve one upright identifier token without inventing application.
            // Explicit atom classes override MathJax normal-mi autoOP inference.
            return (isOperator ? @"\operatorname{" : @"\mathrm{") + EscapeMathIdentifier(token) + "}";
        }
        if (element.Name.LocalName == "mo"
            && TryConvertNamedOperator(token, out var namedOperator))
        {
            // A genuine MathType explicit FONT_STYLE_DEF can style an operator
            // even though its server-exported MathML omits that presentation.
            // The direct MTEF reader carries the style as mathvariant; do not
            // return the named command before applying it.
            if (variant.IndexOf("bold-italic", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathbfit{" + namedOperator.Trim() + "}";
            if (variant.IndexOf("bold", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\boldsymbol{" + namedOperator.Trim() + "}";
            if (variant.IndexOf("italic", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathit{" + namedOperator.Trim() + "}";
            return namedOperator;
        }
        var explicitlyUpright =
            variant.IndexOf("normal", StringComparison.OrdinalIgnoreCase) >= 0
            || variant.IndexOf("upright", StringComparison.OrdinalIgnoreCase) >= 0;
        if (element.Name.LocalName == "mi"
            && TryConvertLetterlikeIdentifier(token, out var letterlikeLatex))
            return ApplyNativeMathTypeCharacterStyle(
                letterlikeLatex,
                nativeCharacterStyle,
                variant,
                greekIdentifier: false);
        if (variant == "italic" && nativeCharacterStyle is 0 or 2
            && TryGetIntrinsicItalicIdentifierSource(element, token, out var italicSource))
            return italicSource;
        if (element.Name.LocalName == "mi" && IsLatinIdentifier(token)
            || element.Name.LocalName == "mn")
        {
            var escaped = EscapeMathIdentifier(token);
            string? converted = null;
            if (variant.IndexOf("double-struck", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathbb{" + escaped + "}";
            else if (variant.IndexOf("bold-fraktur", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\boldsymbol{\mathfrak{" + escaped + "}}";
            else if (variant.IndexOf("fraktur", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathfrak{" + escaped + "}";
            else if (variant.IndexOf("bold-script", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\boldsymbol{\mathcal{" + escaped + "}}";
            else if (variant.IndexOf("script", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathcal{" + escaped + "}";
            else if (variant.IndexOf("monospace", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathtt{" + escaped + "}";
            else if (variant.IndexOf("bold-sans-serif", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\boldsymbol{\mathsf{" + escaped + "}}";
            else if (variant.IndexOf("sans-serif", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathsf{" + escaped + "}";
            else if (variant.IndexOf("bold-italic", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathbfit{" + escaped + "}";
            else if (variant.IndexOf("bold", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathbf{" + escaped + "}";
            else if (element.Name.LocalName == "mn"
                && string.Equals(variant, "italic", StringComparison.OrdinalIgnoreCase))
                converted = @"\mathit{" + escaped + "}";
            else if (explicitlyUpright)
                converted = @"\mathrm{" + escaped + "}";
            if (converted is not null)
                return ApplyNativeMathTypeCharacterStyle(
                    converted,
                    nativeCharacterStyle,
                    variant,
                    greekIdentifier: false);
        }
        if (element.Name.LocalName == "mi" && IsGreekIdentifier(token))
            return ConvertGreekIdentifier(token, nativeCharacterStyle, declaredVariant, effectiveVariant);
        var plain = element.Name.LocalName is "mi" or "mo"
            && string.Equals(((string?)element.Attribute("data-mjx-alternate"))?.Trim(), "1", StringComparison.Ordinal)
            && AlternateTokenMap.TryGetValue(token, out var alternateCommand)
            ? alternateCommand
            : ConvertToken(token);
        if (element.Name.LocalName is "mi" or "mo")
        {
            if (variant.IndexOf("bold-italic", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathbfit{" + plain.Trim() + "}";
            if (variant.IndexOf("bold", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathbf{" + plain.Trim() + "}";
            if (variant.IndexOf("italic", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathit{" + plain.Trim() + "}";
            // A prime has no upright form; VisualTeX for Mac writes f' as an
            // upright apostrophe, which is f^{\prime}, not f^{\mathrm{\prime}}.
            if (token.Length > 0 && token.All(character => character is '\'' or '′' or '″' or '‴' or '⁗'))
                return plain;
            if ((explicitlyUpright && element.Name.LocalName == "mi")
                || ((explicitlyUpright || variant.Length == 0)
                    && TokenDefaultMathVariants.TryGetValue(token, out var intrinsicVariant)
                    && string.Equals(intrinsicVariant, "italic", StringComparison.Ordinal)))
                return @"\mathrm{" + plain.Trim() + "}";
        }
        return plain;
    }

    private static bool TryGetIntrinsicItalicIdentifierSource(
        XElement element, string token, out string source)
    {
        source = string.Empty;
        // MathML italic is the mathematical italic alphabet. TeX \mathit is a
        // different alphabet for Latin letters (MathJax TEX-MI vs TEX-I), and
        // only its actual token-level provenance can authorize that spelling.
        if (element.Name.LocalName != "mi"
            || !MathMlTokenStyleContract.IsSingleMathMlCharacter(token)
            || string.Equals((string?)element.Attribute("data-mjx-variant"),
                "-tex-mathit", StringComparison.Ordinal))
            return false;
        if (IsLatinIdentifier(token))
            source = EscapeMathIdentifier(token);
        else if (IsGreekIdentifier(token) && token.All(char.IsLower))
            source = ConvertToken(token);
        return source.Length > 0;
    }

    private static string ConvertGreekIdentifier(
        string token, int nativeCharacterStyle, string? declaredVariant, string? effectiveVariant)
    {
        // Keep the established Unicode spelling for explicitly styled Greek.
        // It participates directly in the requested mathematical alphabet,
        // including frontends with TeX fixed-family control-sequence defaults.
        // Preserve the requested style before choosing source spelling.
        var variant = declaredVariant ?? effectiveVariant ?? string.Empty;
        // Lowercase Greek commands already have an italic intrinsic default.
        // An undeclared uppercase mi does not: e.g. TeX \\Gamma is normal while
        // MathML <mi>Γ</mi> is italic. Source spelling must preserve that default.
        if (declaredVariant is null && nativeCharacterStyle == 0
            && variant == "italic" && token.All(char.IsLower))
            return ConvertToken(token);
        variant = nativeCharacterStyle switch
        {
            1 => "bold",
            2 => "italic",
            3 => "bold-italic",
            _ => variant,
        };
        if (variant is "normal" or "upright")
            return @"\mathrm{" + token + "}";
        if (variant == "italic")
            return @"\mathit{" + token + "}";
        if (variant == "bold")
            return token.All(char.IsUpper) && TokenMap.TryGetValue(token, out var upperCommand)
                ? @"\boldsymbol{" + upperCommand.Trim() + "}"
                : @"\mathbf{" + token + "}";
        if (variant == "bold-italic")
            return token.All(char.IsLower)
                ? @"\boldsymbol{" + ConvertToken(token).Trim() + "}"
                : @"\boldsymbol{\mathit{" + token + "}}";
        return ApplyNativeMathTypeCharacterStyle(
            ConvertToken(token), nativeCharacterStyle, variant, greekIdentifier: true);
    }

    private static string ApplyNativeMathTypeCharacterStyle(
        string latex,
        int characterStyle,
        string mathVariant,
        bool greekIdentifier)
    {
        if (characterStyle is < 1 or > 3)
            return latex;

        var hasBold = mathVariant.IndexOf(
            "bold",
            StringComparison.OrdinalIgnoreCase) >= 0;
        var hasItalic = mathVariant.IndexOf(
            "italic",
            StringComparison.OrdinalIgnoreCase) >= 0;

        return characterStyle switch
        {
            1 when hasBold => latex,
            1 when greekIdentifier => @"\boldsymbol{" + latex.Trim() + "}",
            1 => @"\mathbf{" + latex + "}",
            2 when hasItalic => latex,
            2 => @"\mathit{" + latex + "}",
            3 when hasBold && hasItalic => latex,
            3 when greekIdentifier => @"\boldsymbol{" + latex.Trim() + "}",
            3 => @"\mathbfit{" + latex + "}",
            _ => latex,
        };
    }

    private static bool TryConvertNamedOperator(string token, out string latex)
    {
        latex = token switch
        {
            "lim" => @"\lim ",
            "max" => @"\max ",
            "min" => @"\min ",
            "sup" => @"\sup ",
            "inf" => @"\inf ",
            "limsup" => @"\limsup ",
            "liminf" => @"\liminf ",
            "sin" => @"\sin ",
            "cos" => @"\cos ",
            "tan" => @"\tan ",
            "cot" => @"\cot ",
            "sec" => @"\sec ",
            "csc" => @"\csc ",
            "sinh" => @"\sinh ",
            "cosh" => @"\cosh ",
            "tanh" => @"\tanh ",
            "log" => @"\log ",
            "ln" => @"\ln ",
            "exp" => @"\exp ",
            "det" => @"\det ",
            "gcd" => @"\gcd ",
            // The rest of TeX's predefined function names: each is an upright
            // operator MathJax applies with U+2061, like Word's m:func.
            "arccos" => @"\arccos ",
            "arcsin" => @"\arcsin ",
            "arctan" => @"\arctan ",
            "arg" => @"\arg ",
            "coth" => @"\coth ",
            "deg" => @"\deg ",
            "dim" => @"\dim ",
            "hom" => @"\hom ",
            "ker" => @"\ker ",
            "lg" => @"\lg ",
            "Pr" => @"\Pr ",
            _ => string.Empty,
        };
        return latex.Length > 0;
    }

    private static bool IsLimitLikeOperatorLatex(string value)
    {
        return value.Trim() switch
        {
            @"\lim" or @"\max" or @"\min" or @"\sup" or @"\inf"
                or @"\limsup" or @"\liminf" => true,
            _ => false,
        };
    }

    private static bool TryConvertLetterlikeIdentifier(
        string token,
        out string latex)
    {
        latex = token switch
        {
            "ℂ" => @"\mathbb{C}",
            "ℍ" => @"\mathbb{H}",
            "ℕ" => @"\mathbb{N}",
            "ℙ" => @"\mathbb{P}",
            "ℚ" => @"\mathbb{Q}",
            "ℝ" => @"\mathbb{R}",
            "ℤ" => @"\mathbb{Z}",
            "ℋ" => @"\mathcal{H}",
            "ℐ" => @"\mathcal{I}",
            "ℒ" => @"\mathcal{L}",
            "ℛ" => @"\mathcal{R}",
            "ℬ" => @"\mathcal{B}",
            "ℯ" => @"\mathcal{e}",
            "ℰ" => @"\mathcal{E}",
            "ℱ" => @"\mathcal{F}",
            "ℳ" => @"\mathcal{M}",
            "ℴ" => @"\mathcal{o}",
            "ℌ" => @"\mathfrak{H}",
            "ℑ" => @"\Im ",
            "ℜ" => @"\Re ",
            "ℭ" => @"\mathfrak{C}",
            _ => string.Empty,
        };
        return latex.Length > 0;
    }

    private static bool IsLatinIdentifier(string token) =>
        token.Length > 0
        && token.Any(char.IsLetter)
        && token.All(character =>
            (character <= '\u024F' && char.IsLetterOrDigit(character))
            || character is '.' or ',');

    private static bool IsGreekIdentifier(string token) =>
        token.Length > 0
        && token.All(character =>
            character is >= '\u0370' and <= '\u03FF'
            || character is >= '\u1F00' and <= '\u1FFF');

    private static string EscapeMathIdentifier(string value) =>
        value.Replace("\\", @"\backslash ")
            .Replace("{", @"\{")
            .Replace("}", @"\}")
            .Replace("#", @"\#")
            .Replace("%", @"\%")
            .Replace("&", @"\&")
            .Replace("_", @"\_");

    private static string ConvertToken(string value)
    {
        var token = value.Trim();
        if (TokenMap.TryGetValue(token, out var mapped)) return mapped;
        if (token.Length == 0) return string.Empty;
        // MathML token text is literal data, not a TeX fragment. MTEF can
        // legitimately group a percent with digits (one <mn>100%</mn>).
        // Escaping only when the entire token equalled "%" turned its suffix
        // into a TeX comment on the next edit. Serialize every literal character
        // before combining tokens; structural commands/alignments are emitted
        // by their own nodes and TokenMap above, never escaped a second time.
        return EscapeMathIdentifier(token);
    }

    // Canonical scalar spellings from the production TeX delimiter handlers:
    // base, AMS and mathtools (including their active priority/shadowing order).
    // The production roundtrip probe checks the complete runtime inventory.
    // GetDelimiter accepts registered spellings, not arbitrary Unicode scalars:
    // e.g. U+27EE and U+23D0 require \\lgroup and \\arrowvert. Keep this role
    // separate from TokenMap so ordinary operators/prose retain their semantics.
    private static readonly IReadOnlyDictionary<string, string> DelimiterMap =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["("] = @"(",
            [")"] = @")",
            ["["] = @"[",
            ["]"] = @"]",
            ["/"] = @"/",
            ["\\"] = @"\backslash",
            ["|"] = @"|",
            ["‖"] = @"\|",
            ["{"] = @"\{",
            ["}"] = @"\}",
            ["⟨"] = @"\langle",
            ["⟩"] = @"\rangle",
            ["⌈"] = @"\lceil",
            ["⌉"] = @"\rceil",
            ["⌊"] = @"\lfloor",
            ["⌋"] = @"\rfloor",
            ["⌜"] = @"\ulcorner",
            ["⌝"] = @"\urcorner",
            ["⌞"] = @"\llcorner",
            ["⌟"] = @"\lrcorner",
            ["⎰"] = @"\lmoustache",
            ["⎱"] = @"\rmoustache",
            ["⟮"] = @"\lgroup",
            ["⟯"] = @"\rgroup",
            ["⏐"] = @"\arrowvert",
            ["⎪"] = @"\bracevert",
            ["↑"] = @"\uparrow",
            ["↓"] = @"\downarrow",
            ["↕"] = @"\updownarrow",
            ["⇑"] = @"\Uparrow",
            ["⇓"] = @"\Downarrow",
            ["⇕"] = @"\Updownarrow",
            // Preserve existing invisible and legacy angle spellings.
            [""] = ".",
            ["."] = ".",
            ["<"] = "<",
            [">"] = ">",
            ["〈"] = @"\langle",
            ["〉"] = @"\rangle",
        };

    private static string Delimiter(string value, bool left)
    {
        // Unknown values retain the legacy source rather than being silently
        // substituted with a different glyph or claimed as supported.
        var escaped = DelimiterMap.TryGetValue(value, out var mapped) ? mapped : value;
        return (left ? @"\left" : @"\right") + escaped + " ";
    }

    private static string GroupBase(string value)
    {
        var trimmed = value.Trim();
        if (trimmed.Length == 1 || trimmed.StartsWith("\\", StringComparison.Ordinal)) return trimmed;
        return "{" + trimmed + "}";
    }

    private static string EscapeText(string value) =>
        value.Replace("\\", @"\textbackslash{}")
            .Replace("{", @"\{")
            .Replace("}", @"\}")
            .Replace("#", @"\#")
            .Replace("%", @"\%")
            .Replace("&", @"\&")
            .Replace("_", @"\_");

    private static string Normalize(string value)
    {
        var result = value.Trim();
        while (result.IndexOf("  ", StringComparison.Ordinal) >= 0)
            result = result.Replace("  ", " ");
        return result;
    }
}
