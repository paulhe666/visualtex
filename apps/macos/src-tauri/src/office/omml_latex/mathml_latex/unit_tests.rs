//! Converter-only assertions ported from VisualTeX.WindowsOffice.Tests: every
//! assertion whose only subject is MathMlToLatexConverter.Convert(mathMl).
//! Each check prints "[cases] <test> <n>" so the total can be summed.

use super::mathml_to_latex;

const NS: &str = "http://www.w3.org/1998/Math/MathML";

struct Check {
    name: &'static str,
    count: usize,
    failures: Vec<String>,
}

impl Check {
    fn new(name: &'static str) -> Self {
        Self { name, count: 0, failures: Vec::new() }
    }

    fn convert(&mut self, input: &str) -> Option<String> {
        match mathml_to_latex(input) {
            Ok(latex) => Some(latex),
            Err(error) => {
                self.failures.push(format!("Convert failed: {error}\n  input: {input}"));
                None
            }
        }
    }

    fn eq(&mut self, input: &str, expected: &str) {
        self.count += 1;
        if let Some(actual) = self.convert(input) {
            if actual != expected {
                self.failures.push(format!("input:    {input}\n  expected: {expected}\n  actual:   {actual}"));
            }
        }
    }

    fn holds(&mut self, input: &str, what: &str, predicate: impl Fn(&str) -> bool) {
        self.count += 1;
        if let Some(actual) = self.convert(input) {
            if !predicate(&actual) {
                self.failures.push(format!("input: {input}\n  expected {what}\n  actual: {actual}"));
            }
        }
    }

    fn fails(&mut self, input: &str) {
        self.count += 1;
        if let Ok(actual) = mathml_to_latex(input) {
            let shown: String = input.chars().take(200).collect();
            self.failures.push(format!("expected an error\n  input: {shown}\n  actual: {actual}"));
        }
    }

    fn finish(self) {
        println!("[cases] {} {}", self.name, self.count);
        assert!(self.failures.is_empty(), "{} failure(s):\n{}", self.failures.len(), self.failures.join("\n"));
    }
}

fn math_ns(body: &str) -> String {
    format!("<math xmlns=\"{NS}\">{body}</math>")
}

fn math(body: &str) -> String {
    format!("<math>{body}</math>")
}

// ---- MathMlToLatexConverterTests ------------------------------------------------

#[test]
fn converts_common_word_mathml_to_editable_latex() {
    let mut c = Check::new("ConvertsCommonWordMathMlToEditableLatex");
    c.eq(&math_ns("<mfrac><mi>a</mi><mi>b</mi></mfrac>"), r"\frac{a}{b}");
    c.eq(
        &math_ns("<msqrt><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><msup><mi>y</mi><mn>2</mn></msup></mrow></msqrt>"),
        r"\sqrt{x^{2}+y^{2}}",
    );
    c.eq(
        &math_ns("<msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup><msup><mi>t</mi><mn>2</mn></msup><mi>d</mi><mi>t</mi>"),
        r"\int_{0}^{1}t^{2}dt",
    );
    c.finish();
}

#[test]
fn keeps_a_barless_fraction_barless() {
    let mut c = Check::new("KeepsABarlessFractionBarless");
    c.eq(&math_ns("<mfrac linethickness=\"0\"><mi>n</mi><mi>k</mi></mfrac>"), r"\genfrac{}{}{0pt}{}{n}{k}");
    c.eq(&math_ns("<mfrac linethickness=\"0pt\"><mi>n</mi><mi>k</mi></mfrac>"), r"\genfrac{}{}{0pt}{}{n}{k}");
    c.eq(&math_ns("<mfrac linethickness=\"1\"><mi>n</mi><mi>k</mi></mfrac>"), r"\frac{n}{k}");
    c.eq(
        &math_ns("<mo>[</mo><mfrac linethickness=\"0\"><mi>n</mi><mi>k</mi></mfrac><mo>]</mo>"),
        r"[\genfrac{}{}{0pt}{}{n}{k}]",
    );
    c.finish();
}

#[test]
fn geometry_symbols_become_their_commands() {
    let mut c = Check::new("GeometrySymbolsBecomeTheirCommands");
    c.eq(&math_ns("<mo>△</mo><mi>A</mi>"), r"\triangle A");
    c.eq(&math_ns("<mo>▲</mo><mi>A</mi>"), r"\blacktriangle A");
    c.eq(&math_ns("<mo>∵</mo>"), r"\because");
    c.eq(&math_ns("<mtext>△</mtext><mi>A</mi>"), r"\triangle A");
    c.eq(&math_ns("<mtext>∠</mtext><mi>A</mi>"), r"\angle A");
    c.finish();
}

#[test]
fn symbols_read_back_as_their_commands() {
    let mut c = Check::new("SymbolsReadBackAsTheirCommands");
    for (body, expected) in [
        ("<mo>⟶</mo>", r"\longrightarrow"),
        ("<mo>⩽</mo>", r"\leqslant"),
        ("<mo>⇌</mo>", r"\rightleftharpoons"),
        ("<mo>⨁</mo>", r"\bigoplus"),
        ("<mo>≲</mo>", r"\lesssim"),
        ("<mo>⊊</mo>", r"\subsetneq"),
        ("<mo>∤</mo>", r"\nmid"),
        ("<mo>↠</mo>", r"\twoheadrightarrow"),
    ] {
        c.eq(&math_ns(body), expected);
    }
    c.finish();
}

#[test]
fn symbol_spelling_facts() {
    let mut c = Check::new("SymbolSpellingFacts");
    c.eq(
        &math_ns("<mi>ε</mi><mi>ϵ</mi><mi>ϑ</mi><mi>ϖ</mi><mi>ϱ</mi><mi>ς</mi><mi>φ</mi><mi>ϕ</mi><mi>ϰ</mi><mo>⋅</mo>"),
        r"\varepsilon \epsilon \vartheta \varpi \varrho \varsigma \varphi \phi \varkappa \cdot",
    );
    c.eq(&math_ns("<mo>⊎</mo><mo>⊓</mo><mo>⊔</mo><mo>≀</mo><mo>⨿</mo>"), r"\uplus \sqcap \sqcup \wr \amalg");
    c.eq(
        &math_ns(
            "<mo>⋆</mo><mo>⊲</mo><mo>⊳</mo><mo>⊏</mo><mo>⊑</mo><mo>⊐</mo><mo>⊒</mo>\
             <mo>⇐</mo><mo>↗</mo><mo>↘</mo><mo>↙</mo><mo>↖</mo>\
             <mo>♠</mo><mo>♣</mo><mo>♡</mo><mo>♢</mo>",
        ),
        r"\star \triangleleft \triangleright \sqsubset \sqsubseteq \sqsupset \sqsupseteq \Leftarrow \nearrow \searrow \swarrow \nwarrow \spadesuit \clubsuit \heartsuit \diamondsuit",
    );
    c.eq(
        &math_ns("<mi>ℵ</mi><mi mathvariant=\"normal\">℘</mi><mi>ℜ</mi><mi>ℑ</mi><mi>ℓ</mi>"),
        r"\aleph \mathrm{\wp}\Re \Im \ell",
    );
    c.eq(&math_ns("<mo>⪯</mo><mo>⪰</mo><mo>≼</mo><mo>≽</mo>"), r"\preceq \succeq \preccurlyeq \succcurlyeq");
    c.eq(&math_ns("<mo>∥</mo><mo>⊥</mo><mo>∝</mo><mo>~</mo><mo>≅</mo>"), r"\parallel \perp \propto \sim \cong");
    c.eq(
        &math_ns(
            "<mo>⊕</mo><mo>⊗</mo><mo>°</mo><mo>•</mo><mo>⋄</mo>\
             <mo>≪</mo><mo>≫</mo><mo>∋</mo><mo>∠</mo><mo>∴</mo><mo>∖</mo>\
             <mo>…</mo><mo>⋯</mo><mo>⋮</mo><mo>⋱</mo>",
        ),
        r"\oplus \otimes \circ \bullet \diamond \ll \gg \ni \angle \therefore \setminus \dots \cdots \vdots \ddots",
    );
    c.finish();
}

#[test]
fn writes_upright_limit_identifiers_as_limit_operators() {
    let mut c = Check::new("WritesUprightLimitIdentifiersAsLimitOperatorsWithLimitsUnderneath");
    for (name, expected) in [
        ("lim", r"\lim\limits_{x\to 0}"),
        ("max", r"\max\limits_{x\in A}"),
        ("min", r"\min\limits_{x\in A}"),
        ("sup", r"\sup\limits_{x\in A}"),
        ("inf", r"\inf\limits_{x\in A}"),
        ("limsup", r"\limsup\limits_{n\to \mathrm{\infty}}"),
        ("liminf", r"\liminf\limits_{n\to \mathrm{\infty}}"),
    ] {
        let mut script = if name.starts_with("lim") {
            "<mrow><mi>n</mi><mo>→</mo><mi mathvariant=\"normal\">∞</mi></mrow>"
        } else {
            "<mrow><mi>x</mi><mo>∈</mo><mi>A</mi></mrow>"
        };
        if name == "lim" {
            script = "<mrow><mi>x</mi><mo>→</mo><mn>0</mn></mrow>";
        }
        c.eq(&math_ns(&format!("<munder><mi mathvariant=\"normal\">{name}</mi>{script}</munder>")), expected);
    }
    c.finish();
}

#[test]
fn binomial_and_fence_recovery() {
    let mut c = Check::new("BinomialAndFenceRecovery");
    c.eq(&math_ns("<munder><mi>x</mi><mi>i</mi></munder>"), r"\underset{i}{x}");
    c.eq(
        &math_ns("<mo>(</mo><mfrac linethickness=\"0\"><mi>n</mi><mi>k</mi></mfrac><mo>)</mo><mo>=</mo><mn>1</mn>"),
        r"\binom{n}{k}=1",
    );
    c.eq(
        &math_ns(
            "<mrow data-mjx-texclass=\"ORD\">\
             <mrow data-mjx-texclass=\"OPEN\"><mo minsize=\"2.047em\" maxsize=\"2.047em\">(</mo></mrow>\
             <mfrac linethickness=\"0\"><mi>n</mi><mi>k</mi></mfrac>\
             <mrow data-mjx-texclass=\"CLOSE\"><mo minsize=\"2.047em\" maxsize=\"2.047em\">)</mo></mrow>\
             </mrow>",
        ),
        r"\binom{n}{k}",
    );
    c.eq(
        &math_ns("<mfenced open=\"(\" close=\")\"><mrow><mfrac linethickness=\"0\"><mi>n</mi><mi>k</mi></mfrac></mrow></mfenced>"),
        r"\binom{n}{k}",
    );
    c.eq(
        &math_ns(
            "<mo>(</mo><mtable data-mtef-pile=\"true\">\
             <mtr><mtd><mi>n</mi></mtd></mtr>\
             <mtr><mtd><mi>k</mi></mtd></mtr>\
             </mtable><mo>)</mo><mo>=</mo><mn>1</mn>",
        ),
        r"\binom{n}{k}=1",
    );
    c.eq(&math_ns("<mo>(</mo><mfrac><mi>n</mi><mi>k</mi></mfrac><mo>)</mo>"), r"(\frac{n}{k})");
    c.eq(
        &math_ns("<mfenced open=\"(\" close=\")\" separators=\",\"><msub><mi>A</mi><mn>9</mn></msub><mi>B</mi></mfenced>"),
        r"\left( A_{9},B\right)",
    );
    c.finish();
}

#[test]
fn converts_word_prime_spellings_to_latex_prime() {
    let mut c = Check::new("ConvertsWordPrimeSpellingsToLatexPrime");
    for (prime, expected) in [("'", r"f^{\prime}"), ("′", r"f^{\prime}"), ("″", r"f^{\prime\prime}")] {
        c.eq(&math_ns(&format!("<msup><mi>f</mi><mi>{prime}</mi></msup>")), expected);
    }
    c.finish();
}

#[test]
fn styles_braces_bars_and_delimiters() {
    let mut c = Check::new("StylesBracesBarsAndDelimiters");
    c.eq(
        &math_ns("<mi mathvariant=\"normal\">e</mi><msup><mi mathvariant=\"normal\">i</mi><mi>π</mi></msup><mi>x</mi>"),
        r"\mathrm{e}\mathrm{i}^{\pi}x",
    );
    for (inner, marker, expected) in
        [("mover", "⏞", r"\overbrace{a+b}^{n}"), ("munder", "⏟", r"\underbrace{a+b}_{n}")]
    {
        c.eq(
            &math_ns(&format!(
                "<{inner}><mrow><{inner}><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mo stretchy=\"true\">{marker}</mo></{inner}></mrow><mi>n</mi></{inner}>"
            )),
            expected,
        );
    }
    for (name, expected) in [("mover", r"\overline{AB}"), ("munder", r"\underline{AB}")] {
        c.eq(&math_ns(&format!("<{name}><mrow><mi>A</mi><mi>B</mi></mrow><mo>―</mo></{name}>")), expected);
    }
    for (fence, expected) in [("|", r"\left| x\right|"), ("‖", r"\left\| x\right\|")] {
        c.eq(&math_ns(&format!("<mfenced open=\"{fence}\" close=\"{fence}\"><mi>x</mi></mfenced>")), expected);
    }
    for name in ["mi", "mo"] {
        c.eq(&math_ns(&format!("<{name} mathvariant=\"bold\">∞</{name}>")), r"\mathbf{\infty}");
    }
    c.finish();
}

#[test]
fn converts_mathtype_tab_pile_table_back_to_aligned() {
    let mut c = Check::new("ConvertsMathTypeTabPileTableBackToAligned");
    let input = math_ns(
        "<mtable data-mtef-tabs=\"true\" columnalign=\"right left right left\">\
         <mtr><mtd><msub><mi>p</mi><mn>0</mn></msub></mtd>\
         <mtd><mo>=</mo><mi>x</mi></mtd><mtd></mtd><mtd><mtext>Low word</mtext></mtd></mtr>\
         </mtable>",
    );
    c.holds(&input, "StartsWith \\begin{aligned}", |l| l.starts_with(r"\begin{aligned}"));
    c.holds(&input, "EndsWith \\end{aligned}", |l| l.ends_with(r"\end{aligned}"));
    c.holds(&input, "three &", |l| l.chars().filter(|&ch| ch == '&').count() == 3);
    c.holds(&input, "Contains & & \\text{Low word}", |l| l.contains(r"& & \text{Low word}"));
    c.finish();
}

#[test]
fn converts_matrices_and_greek_symbols() {
    let mut c = Check::new("ConvertsMatricesAndGreekSymbols");
    let input = math_ns(
        "<mfenced open=\"(\" close=\")\"><mtable>\
         <mtr><mtd><mi>α</mi></mtd><mtd><mi>β</mi></mtd></mtr>\
         <mtr><mtd><mi>γ</mi></mtd><mtd><mi>δ</mi></mtd></mtr>\
         </mtable></mfenced>",
    );
    c.holds(&input, "Contains \\begin{matrix}", |l| l.contains(r"\begin{matrix}"));
    c.holds(&input, "Contains \\alpha", |l| l.contains(r"\alpha"));
    c.holds(&input, "Contains \\delta", |l| l.contains(r"\delta"));
    c.holds(&input, "StartsWith \\left(", |l| l.starts_with(r"\left("));
    c.holds(&input, "EndsWith \\right)", |l| l.ends_with(r"\right)"));
    c.finish();
}

// ---- MathMlFunctionTokenRecoveryTests --------------------------------------------

#[test]
fn preserves_token_boundary_and_actual_application_slot() {
    let mut c = Check::new("PreservesTokenBoundaryAndActualApplicationSlot");
    for (body, expected) in [
        ("<mi>sin</mi><mi>x</mi>", "\\operatorname{sin}x"),
        ("<mi>sin</mi><mo>\u{2061}</mo><mi>x</mi>", "\\sin x"),
        ("<mi>foo</mi><mi>x</mi>", "\\operatorname{foo}x"),
        ("<mi data-mjx-auto-op='false'>foo</mi><mi>x</mi>", "\\mathrm{foo}x"),
        ("<mi data-mjx-texclass='OP' data-mjx-auto-op='false'>sin</mi><mi>x</mi>", "\\operatorname{sin}x"),
        ("<mi data-mjx-texclass='ORD'>sin</mi><mi>x</mi>", "\\mathrm{sin}x"),
        ("<msup><mi>foo</mi><mn>2</mn></msup><mo>\u{2061}</mo><mi>x</mi>", "\\operatorname{foo}^{2}\u{2061}x"),
        ("<mi>foo</mi><mo>\u{2061}</mo><mo>+</mo><mi>x</mi>", "\\operatorname{foo}\u{2061}+x"),
        ("<mi>x</mi><mi>y</mi>", "xy"),
        ("<mo>lim</mo>", "\\lim"),
        ("<munder><mo>lim</mo><mi>x</mi></munder>", "\\lim_{x}"),
        ("<munder><mo>max</mo><mi>x</mi></munder>", "\\max_{x}"),
    ] {
        c.eq(&math(body), expected);
    }
    c.finish();
}

// ---- MathMlAccentWidthContractTests -----------------------------------------------

#[test]
fn recovers_declared_or_dictionary_accent_width() {
    let mut c = Check::new("RecoversDeclaredOrDictionaryAccentWidth");
    for (mark, stretchy, command) in [
        ("^", Some("true"), "widehat"),
        ("^", Some("false"), "hat"),
        ("^", None, "widehat"),
        ("~", Some("true"), "widetilde"),
        ("~", Some("false"), "tilde"),
        ("~", None, "widetilde"),
        ("→", Some("true"), "overrightarrow"),
        ("→", Some("false"), "vec"),
        ("→", None, "overrightarrow"),
    ] {
        let attribute = stretchy.map(|s| format!(" stretchy='{s}'")).unwrap_or_default();
        c.eq(
            &format!("<math><mover><mrow><mi>x</mi><mi>y</mi></mrow><mo{attribute}>{mark}</mo></mover></math>"),
            &format!("\\{command}{{xy}}"),
        );
    }
    c.finish();
}

// ---- MathMlAlternateTokenRecoveryTests ----------------------------------------------

#[test]
fn alternate_token_recovery() {
    let mut c = Check::new("MathMlAlternateTokenRecoveryTests");
    // ExplicitAlternateRecoversItsSupportedCommand
    for (token, command) in [("∥", r"\shortparallel"), ("∦", r"\nshortparallel")] {
        c.eq(&format!("<math><mo data-mjx-alternate='1'>{token}</mo></math>"), command);
        c.eq(
            &format!("<math><mi>a</mi><mo data-mjx-alternate='1'>{token}</mo><mi>b</mi></math>"),
            &format!("a{command} b"),
        );
    }
    // AlternateSelectionPrecedesStyleAndScriptRecovery
    for (token, command, variant, wrapper) in [
        ("∥", r"\shortparallel", "italic", r"\mathit"),
        ("∥", r"\shortparallel", "bold", r"\mathbf"),
        ("∥", r"\shortparallel", "bold-italic", r"\mathbfit"),
        ("∦", r"\nshortparallel", "italic", r"\mathit"),
        ("∦", r"\nshortparallel", "bold", r"\mathbf"),
        ("∦", r"\nshortparallel", "bold-italic", r"\mathbfit"),
    ] {
        c.eq(
            &format!(
                "<math><msubsup><mstyle mathvariant='{variant}'><mo data-mjx-alternate='1'>{token}</mo></mstyle><mi>i</mi><mn>2</mn></msubsup></math>"
            ),
            &format!("{wrapper}{{{command}}}_{{i}}^{{2}}"),
        );
    }
    // OnlyKnownExplicitAlternateEvidenceChangesTheCommand
    for alternate in ["", "0", "2", "true"] {
        c.eq(&format!("<math><mo data-mjx-alternate='{alternate}'>∥</mo></math>"), r"\parallel");
    }
    // ContainerAlternateDoesNotLeakToChildren
    c.eq("<math><mrow data-mjx-alternate='1'><mo>∥</mo></mrow></math>", r"\parallel");
    // PlanckGlyphSelectionPreservesDeclaredInheritedStyleAndScriptSlots
    for role in ["mi", "mo"] {
        for style in ["normal", "italic", "bold", "bold-italic"] {
            for alternate in [false, true] {
                let command = if alternate { r"\hbar" } else { r"\hslash" };
                let expected = match style {
                    "italic" => format!("\\mathit{{{command}}}"),
                    "bold" => format!("\\mathbf{{{command}}}"),
                    "bold-italic" => format!("\\mathbfit{{{command}}}"),
                    "normal" if role == "mi" => format!("\\mathrm{{{command}}}"),
                    _ => command.to_string(),
                };
                let flag = if alternate { " data-mjx-alternate='1'" } else { "" };
                let token = format!("<{role}{flag}>ℏ</{role}>");
                c.eq(&format!("<math><{role} mathvariant='{style}'{flag}>ℏ</{role}></math>"), &expected);
                c.eq(&format!("<math><mstyle mathvariant='{style}'>{token}</mstyle></math>"), &expected);
                c.eq(
                    &format!("<math><msubsup><mstyle mathvariant='{style}'>{token}</mstyle><mi>i</mi><mn>2</mn></msubsup></math>"),
                    &format!("{expected}_{{i}}^{{2}}"),
                );
            }
        }
    }
    // PlanckGlyphWithoutDeclaredStyleUsesOnlyItsOwnAlternate
    for (role, alternate) in [("mi", false), ("mo", false), ("mi", true), ("mo", true)] {
        let flag = if alternate { " data-mjx-alternate='1'" } else { "" };
        c.eq(&format!("<math><{role}{flag}>ℏ</{role}></math>"), if alternate { r"\hbar" } else { r"\hslash" });
    }
    // PlanckGlyphInvalidOrContainerAlternateCannotInventHbar
    for flag in ["", "0", "2", "true"] {
        for role in ["mi", "mo"] {
            c.eq(&format!("<math><{role} data-mjx-alternate='{flag}'>ℏ</{role}></math>"), r"\hslash");
            c.eq(&format!("<math><mrow data-mjx-alternate='1'><{role}>ℏ</{role}></mrow></math>"), r"\hslash");
        }
    }
    // PlanckGlyphAlternateDoesNotChangeProseRecovery
    c.eq("<math><mtext data-mjx-alternate='1'>ℏ</mtext></math>", r"\text{ℏ}");
    // OtherScalarsRemainOnTheirExistingRecoveryPath
    c.eq("<math><mo data-mjx-alternate='1'>⫽</mo></math>", "⫽");
    c.finish();
}

// ---- MathMlAngleTokenRecoveryTests ----------------------------------------------------

#[test]
fn angle_token_recovery() {
    let mut c = Check::new("MathMlAngleTokenRecoveryTests");
    let mut cases = Vec::new();
    for (token, command) in [("∡", r"\measuredangle"), ("∢", r"\sphericalangle")] {
        for (tag, variant, wrapper) in [
            ("mo", "", r"\mathrm"),
            ("mi", "", r"\mathit"),
            ("mi", "normal", r"\mathrm"),
            ("mo", "normal", r"\mathrm"),
            ("mi", "italic", r"\mathit"),
            ("mo", "italic", r"\mathit"),
            ("mi", "bold", r"\mathbf"),
            ("mo", "bold-italic", r"\mathbfit"),
        ] {
            let attribute = if variant.is_empty() { String::new() } else { format!(" mathvariant='{variant}'") };
            cases.push((format!("<{tag}{attribute}>{token}</{tag}>"), format!("{wrapper}{{{command}}}")));
        }
    }
    for (token, expected) in &cases {
        c.eq(&format!("<math>{token}</math>"), expected);
    }
    for (token, expected) in &cases {
        c.eq(&format!("<math><msubsup>{token}<mi>i</mi><mn>2</mn></msubsup></math>"), &format!("{expected}_{{i}}^{{2}}"));
    }
    for (token, command) in [("∡", r"\measuredangle"), ("∢", r"\sphericalangle")] {
        c.eq(
            &format!("<math><mstyle mathvariant='italic'><mstyle mathvariant='normal'><mi>{token}</mi></mstyle></mstyle></math>"),
            &format!("\\mathrm{{{command}}}"),
        );
        c.eq(
            &format!("<math><mstyle mathvariant='normal'><mi mathvariant='italic'>{token}</mi></mstyle></math>"),
            &format!("\\mathit{{{command}}}"),
        );
        c.eq(
            &format!("<math><mstyle mathvariant='bold-italic'><mo mathvariant='normal'>{token}</mo></mstyle></math>"),
            &format!("\\mathrm{{{command}}}"),
        );
    }
    c.eq("<math><mi mathvariant='normal'>∠</mi></math>", r"\mathrm{\angle}");
    c.eq("<math><mo>∠</mo></math>", r"\angle");
    c.eq("<math><mi>∠</mi></math>", r"\mathit{\angle}");
    c.finish();
}

// ---- MathMlAtomClassRecoveryTests -------------------------------------------------------

#[test]
fn atom_class_recovery() {
    let mut c = Check::new("MathMlAtomClassRecoveryTests");
    for tag in ["mrow", "mstyle", "mpadded"] {
        for (class, command) in [("REL", r"\mathrel"), ("BIN", r"\mathbin"), ("OP", r"\mathop")] {
            c.eq(
                &format!("<math><{tag} data-mjx-texclass='{class}'><mo>&#x2AFD;</mo></{tag}></math>"),
                &format!("{command}{{⫽}}"),
            );
        }
    }
    for (class, expected) in [("REL", r"\mathrel{x+y}"), ("BIN", r"\mathbin{x+y}"), ("OP", r"\mathop{x+y}")] {
        c.eq(&format!("<math><mrow data-mjx-texclass='{class}'><mi>x</mi><mo>+</mo><mi>y</mi></mrow></math>"), expected);
    }
    c.eq(
        "<math><msubsup><mrow data-mjx-texclass='BIN'><mstyle mathvariant='italic' data-mjx-texclass='REL'><mo>∥</mo></mstyle></mrow><mi>i</mi><mn>2</mn></msubsup></math>",
        r"\mathbin{\mathrel{\mathit{\parallel}}}_{i}^{2}",
    );
    for class in ["ORD", "OPEN", "CLOSE", "INNER", "unknown"] {
        c.eq(&format!("<math><mrow data-mjx-texclass='{class}'><mo>⫽</mo></mrow></math>"), "⫽");
    }
    c.eq("<math><mo>⫽</mo></math>", "⫽");
    c.eq(
        "<math><mrow data-mjx-texclass='REL'><mo>(</mo></mrow><mfrac linethickness='0'><mi>n</mi><mi>k</mi></mfrac><mo>)</mo></math>",
        r"\mathrel{(}\genfrac{}{}{0pt}{}{n}{k})",
    );
    c.eq("<math><mrow data-mjx-texclass='REL'/></math>", "");
    c.finish();
}

// ---- MathMlDelimiterSourceRecoveryTests ----------------------------------------------------

/// XAttribute serialization (XElement.ToString(SaveOptions.DisableFormatting)).
fn xml_attr(value: &str) -> String {
    value.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('"', "&quot;")
}

/// `Fence(open, close, children...)` serialized as XElement does.
fn fence(open: &str, close: &str, children: &[&str]) -> String {
    let head = format!("<mfenced open=\"{}\" close=\"{}\" separators=\"\"", xml_attr(open), xml_attr(close));
    if children.is_empty() { format!("{head} />") } else { format!("{head}>{}</mfenced>", children.concat()) }
}

fn recover(children: &[&str]) -> String {
    format!("<math>{}</math>", children.concat())
}

#[test]
fn delimiter_source_recovery() {
    let mut c = Check::new("MathMlDelimiterSourceRecoveryTests");
    let x = "<mi>x</mi>";
    for (scalar, source) in [
        ("(", "("),
        (")", ")"),
        ("[", "["),
        ("]", "]"),
        ("/", "/"),
        ("\\", r"\backslash"),
        ("|", "|"),
        ("‖", r"\|"),
        ("{", r"\{"),
        ("}", r"\}"),
        ("⟨", r"\langle"),
        ("⟩", r"\rangle"),
        ("⌈", r"\lceil"),
        ("⌉", r"\rceil"),
        ("⌊", r"\lfloor"),
        ("⌋", r"\rfloor"),
        ("⌜", r"\ulcorner"),
        ("⌝", r"\urcorner"),
        ("⌞", r"\llcorner"),
        ("⌟", r"\lrcorner"),
        ("⎰", r"\lmoustache"),
        ("⎱", r"\rmoustache"),
        ("⟮", r"\lgroup"),
        ("⟯", r"\rgroup"),
        ("⏐", r"\arrowvert"),
        ("⎪", r"\bracevert"),
        ("↑", r"\uparrow"),
        ("↓", r"\downarrow"),
        ("↕", r"\updownarrow"),
        ("⇑", r"\Uparrow"),
        ("⇓", r"\Downarrow"),
        ("⇕", r"\Updownarrow"),
        ("", "."),
        (".", "."),
        ("<", "<"),
        (">", ">"),
        ("\u{2329}", r"\langle"),
        ("\u{232A}", r"\rangle"),
    ] {
        c.eq(&recover(&[&fence(scalar, "", &[x])]), &format!("\\left{source} x\\right."));
        c.eq(&recover(&[&fence("", scalar, &[x])]), &format!("\\left. x\\right{source}"));
    }
    for scalar in ["⚑", "∥"] {
        c.eq(&recover(&[&fence(scalar, "", &[x])]), &format!("\\left{scalar} x\\right."));
    }
    let xyz = ["<mi>x</mi>", "<mi>y</mi>", "<mi>z</mi>"];
    for (open, close, expected) in [
        ("(", ")", r"\left( xyz\right)"),
        ("⟮", "⟯", r"\left\lgroup xyz\right\rgroup"),
        ("⏐", "", r"\left\arrowvert xyz\right."),
        ("", "", r"\left. xyz\right."),
    ] {
        c.eq(&recover(&[&fence(open, close, &xyz)]), expected);
    }
    c.eq("<math><mfenced><mi>x</mi><mi>y</mi><mi>z</mi></mfenced></math>", r"\left( x,y,z\right)");
    c.eq("<math><mfenced separators=','><mi>x</mi><mi>y</mi><mi>z</mi></mfenced></math>", r"\left( x,y,z\right)");
    c.eq(
        &recover(&[&fence("⌜", "⌝", &[x]), "<mo>+</mo>", &fence("⌞", "⌟", &["<mi>y</mi>"])]),
        r"\left\ulcorner x\right\urcorner +\left\llcorner y\right\lrcorner",
    );
    c.eq(
        &recover(&[&fence("⟮", "⟯", &[x]), "<mo>+</mo>", &fence("⏐", "", &["<mi>y</mi>"])]),
        r"\left\lgroup x\right\rgroup +\left\arrowvert y\right.",
    );
    c.eq(
        &recover(&[&fence("⟮", "⟯", &[x, &fence("⏐", "", &["<mi>y</mi>"]), "<mi>z</mi>"])]),
        r"\left\lgroup x\left\arrowvert y\right. z\right\rgroup",
    );
    c.eq(&recover(&[&fence("⎰", "⎱", &[])]), r"\left\lmoustache \right\rmoustache");
    for (body, expected) in [
        ("<mo>↑</mo>", r"\uparrow"),
        ("<mo>∥</mo>", r"\parallel"),
        ("<mo>⟮</mo><mo>⏐</mo>", "⟮⏐"),
        ("<mtext>⟮|⏐</mtext>", r"\text{⟮|⏐}"),
    ] {
        c.eq(&math(body), expected);
    }
    c.finish();
}

// ---- MathMlGreekSourceRecoveryTests ---------------------------------------------------------

#[test]
fn greek_source_recovery() {
    let mut c = Check::new("MathMlGreekSourceRecoveryTests");
    for (text, variant, expected) in [
        ("α", "normal", r"\mathrm{α}"),
        ("Γ", "normal", r"\mathrm{Γ}"),
        ("α", "upright", r"\mathrm{α}"),
        ("α", "italic", r"\alpha"),
        ("Γ", "italic", r"\mathit{Γ}"),
        ("α", "bold", r"\mathbf{α}"),
        ("Γ", "bold", r"\boldsymbol{\Gamma}"),
        ("α", "bold-italic", r"\boldsymbol{\alpha}"),
        ("Γ", "bold-italic", r"\boldsymbol{\mathit{Γ}}"),
        ("αβ", "normal", r"\mathrm{αβ}"),
        ("Γα", "italic", r"\mathit{Γα}"),
        ("ΓΔ", "bold", r"\mathbf{ΓΔ}"),
        ("\u{0391}", "bold", "\\mathbf{\u{0391}}"),
    ] {
        c.eq(&format!("<math><mi mathvariant='{variant}'>{text}</mi></math>"), expected);
    }
    for (text, expected) in
        [("α", r"\alpha"), ("Γ", r"\mathit{Γ}"), ("ΓΔ", r"\mathrm{ΓΔ}"), ("αβ", r"\mathrm{αβ}")]
    {
        c.eq(&format!("<math><mi>{text}</mi></math>"), expected);
    }
    c.eq(
        "<math><mstyle mathvariant='normal'><mi>α</mi><mstyle mathvariant='italic'><mi>β</mi></mstyle><mi mathvariant='bold'>γ</mi></mstyle></math>",
        r"\mathrm{α}\beta \mathbf{γ}",
    );
    for (bits, expected) in [(0, r"\mathrm{α}"), (1, r"\mathbf{α}"), (2, r"\mathit{α}"), (3, r"\boldsymbol{\alpha}")] {
        c.eq(
            &format!("<math><mi mathvariant='normal' data-mtef-character-style='{bits}'>α</mi></math>"),
            expected,
        );
    }
    c.finish();
}

// ---- MathMlItalicSourceAlphabetTests ------------------------------------------------------

#[test]
fn italic_source_alphabet() {
    let mut c = Check::new("MathMlItalicSourceAlphabetTests");
    for (token, expected) in [("A", "A"), ("x", "x"), ("α", r"\alpha"), ("π", r"\pi")] {
        c.eq(&format!("<math><mi mathvariant='italic'>{token}</mi></math>"), expected);
        // element.SetAttributeValue("data-mtef-character-style", 2); Convert(element.Parent.ToString())
        c.eq(
            &format!("<math>\n  <mi mathvariant=\"italic\" data-mtef-character-style=\"2\">{token}</mi>\n</math>"),
            expected,
        );
    }
    for token in ["A", "x", "α", "π"] {
        c.eq(
            &format!("<math><mi mathvariant='italic' data-mjx-variant='-tex-mathit'>{token}</mi></math>"),
            &format!("\\mathit{{{token}}}"),
        );
    }
    c.eq("<math><mstyle mathvariant='italic' data-mjx-variant='-tex-mathit'><mi>A</mi></mstyle></math>", "A");
    c.eq("<math><mi mathvariant='normal'>A</mi></math>", r"\mathrm{A}");
    c.eq("<math><mi mathvariant='bold-italic'>A</mi></math>", r"\mathbfit{A}");
    c.finish();
}

// ---- MathMlSymbolStyleRecoveryTests -------------------------------------------------------

#[test]
fn symbol_style_recovery() {
    let mut c = Check::new("MathMlSymbolStyleRecoveryTests");
    for (symbol, expected) in
        [("∠", r"\mathit{\angle}"), ("∇", r"\mathit{\nabla}"), ("∞", r"\mathit{\infty}"), ("℘", r"\mathit{\wp}")]
    {
        c.eq(&format!("<math><mi>{symbol}</mi></math>"), expected);
    }
    for (variant, expected) in
        [("italic", r"\mathit{\parallel}"), ("bold", r"\mathbf{\parallel}"), ("bold-italic", r"\mathbfit{\parallel}")]
    {
        c.eq(&format!("<math><mstyle mathvariant='{variant}'><mo>∥</mo></mstyle></math>"), expected);
    }
    c.eq("<math><mstyle mathvariant='italic'><mo mathvariant='normal'>∥</mo></mstyle></math>", r"\parallel");
    for (symbol, attributes, expected) in [
        ("ℏ", "", r"\mathrm{\hslash}"),
        ("ℏ", " data-mjx-alternate='1'", r"\mathrm{\hbar}"),
        ("∠", "", r"\mathrm{\angle}"),
    ] {
        c.eq(&format!("<math><mi mathvariant='normal'{attributes}>{symbol}</mi></math>"), expected);
        c.eq(&format!("<math><mstyle mathvariant='normal'><mi{attributes}>{symbol}</mi></mstyle></math>"), expected);
    }
    c.eq("<math><mi>x</mi><mi>α</mi></math>", r"x\alpha");
    c.finish();
}

// ---- MathMlVerifiedAlternateSourceTests (Convert-only assertion) ---------------------------

#[test]
fn container_alternate_does_not_acquire_alternate_source() {
    let mut c = Check::new("OrdinaryNativeCounterpartDoesNotAcquireAlternateSource(container)");
    for (token, command) in [
        ("∣", r"\shortmid"),
        ("∤", r"\nshortmid"),
        ("⊊", r"\varsubsetneq"),
        ("⊋", r"\varsupsetneq"),
        ("⫋", r"\varsubsetneqq"),
        ("⫌", r"\varsupsetneqq"),
        ("⊈", r"\nsubseteqq"),
        ("⊉", r"\nsupseteqq"),
    ] {
        for style in ["normal", "italic", "bold", "bold-italic"] {
            c.holds(
                &format!("<math><mrow data-mjx-alternate='1'><mo mathvariant='{style}'>{token}</mo></mrow></math>"),
                &format!("DoesNotContain {command}"),
                |latex| !latex.contains(command),
            );
        }
    }
    c.finish();
}

// ---- MathSpaceContractTests -------------------------------------------------------------------

#[test]
fn math_space_contract() {
    let mut c = Check::new("MathSpaceContractTests");
    for (width, latex, native) in [
        (Some("1em"), r"\quad", true),
        (Some("+001.000em"), r"\quad", true),
        (Some("2.00em"), r"\qquad", true),
        (Some("0.123456789em"), r"\hspace{0.123456789em}", false),
        (Some("-0.333333em"), r"\hspace{-0.333333em}", false),
        (Some("1.00000000000000000000000000001em"), r"\hspace{1.00000000000000000000000000001em}", false),
        (Some("0.99999999999999999999999999999em"), r"\hspace{0.99999999999999999999999999999em}", false),
        (Some("2.00000000000000000000000000001em"), r"\hspace{2.00000000000000000000000000001em}", false),
        (Some("1.25ex"), r"\hspace{1.25ex}", false),
        (Some("3pt"), r"\hspace{3pt}", false),
        (Some("12px"), r"\hspace{12px}", false),
        (Some(".5in"), r"\hspace{.5in}", false),
        (Some("2cm"), r"\hspace{2cm}", false),
        (Some("-2mm"), r"\hspace{-2mm}", false),
        (Some("+01.20pc"), r"\hspace{+01.20pc}", false),
        (Some("-0.00em"), r"\hspace{-0.00em}", false),
        (Some(" .500 pt "), r"\hspace{.500pt}", false),
        (None, r"\hspace{0em}", false),
    ] {
        let attribute = width.map(|w| format!(" width='{w}'")).unwrap_or_default();
        c.eq(
            &format!("<math><mi>x</mi><mspace{attribute}/><mi>y</mi></math>"),
            &format!("x{latex}{}y", if native { " " } else { "" }),
        );
    }
    c.eq("<math><mspace width='12%'/></math>", r"\,");
    c.eq(
        "<math><mi>x</mi><mspace width='1.25em' data-mtef-mtcode='61189' data-mtef-explicit-typeface='24'/><mi>y</mi></math>",
        r"x\hspace{1.25em}y",
    );
    c.finish();
}

// ---- FormulaXmlBoundaryTests (latex route) -------------------------------------------------

#[test]
fn formula_xml_boundary() {
    let mut c = Check::new("FormulaXmlBoundaryTests(latex)");
    let deep = "<mrow>".repeat(260) + "<mi>x</mi>" + &"</mrow>".repeat(260);
    c.fails(&format!("<math xmlns='{NS}'>{deep}</math>"));
    c.fails(&format!(
        "<!DOCTYPE math [<!ENTITY external SYSTEM 'file:///does-not-exist'>]><math xmlns='{NS}'><mtext>&external;</mtext></math>"
    ));
    c.fails(&format!("<math xmlns='{NS}'><mtext>{}</mtext></math>", "x".repeat(4_000_000)));
    c.finish();
}

// ---- MathType*/WordOmml* Convert-only assertions ---------------------------------------------

#[test]
fn mathtype_and_word_convert_only_assertions() {
    let mut c = Check::new("MathType/WordOmml Convert-only");
    // MathTypeEmbellishmentTests
    c.eq(
        "<math><mmultiscripts><mi>T</mi><mi>i</mi><mi>j</mi><mi>k</mi><mi>l</mi><mprescripts/><mi>a</mi><mi>b</mi><mi>c</mi><mi>d</mi></mmultiscripts></math>",
        r"{}_{c}^{d}{}_{a}^{b}T_{i}^{j}{}_{k}^{l}",
    );
    c.eq(
        "<math><mmultiscripts><mi>x</mi><none/><mo>&#x2033;</mo><mprescripts/><none/><mo>&#x2035;</mo></mmultiscripts></math>",
        r"{}^{\backprime}x^{\prime\prime}",
    );
    // MathTypeLiteralTokenTests
    for (tag, text, expected) in [
        ("mn", "100%", r"100\%"),
        ("mi", "rate%", r"rate\%"),
        ("mo", "%%", r"\%\%"),
        ("mi", "a_b", r"a\_b"),
        ("mo", "a&amp;b", r"a\&b"),
        ("mi", "a#b", r"a\#b"),
        ("mi", "a{b}", r"a\{b\}"),
    ] {
        c.eq(&format!("<math><{tag}>{text}</{tag}><mo>+</mo><mn>2</mn></math>"), &format!("{expected}+2"));
    }
    c.eq("<math><mi>α</mi><mo>+</mo><mn>100</mn><mo>%</mo></math>", r"\alpha +100\%");
    c.eq(
        "<math><mtable><mtr><mtd><mi>x</mi></mtd><mtd><mn>100%</mn></mtd></mtr></mtable></math>",
        r"\begin{matrix}x & 100\%\end{matrix}",
    );
    // WordOmmlAccentControlStyleTests.RecoversAccentScopeWithoutRestylingBase
    for (variant, expected) in [
        ("bold", r"\mathbf{\hat{\mathnormal{x}}}"),
        ("italic", r"\mathit{\hat{\mathnormal{x}}}"),
        ("bold-italic", r"\mathbfit{\hat{\mathnormal{x}}}"),
    ] {
        c.eq(
            &format!("<math xmlns='{NS}'><mover><mi>x</mi><mo mathvariant='{variant}' stretchy='false'>^</mo></mover></math>"),
            expected,
        );
    }
    // WordOmmlColorTests.MathMlSourceKeepsNestedColorAndDefaultScope
    c.eq(
        &math_ns("<mstyle mathcolor=\"#2563eb\"><mi>x</mi><mstyle mathcolor=\"#dc2626\"><mi>y</mi></mstyle><mi>z</mi></mstyle><mi>q</mi>"),
        r"\textcolor{#2563EB}{x\textcolor{#DC2626}{y}z}q",
    );
    c.finish();
}
