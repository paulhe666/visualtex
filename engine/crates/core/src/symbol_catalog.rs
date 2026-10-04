//! Bounded authored P07 symbol vocabulary. A command names an atom, not a
//! template, a CSS glyph or a macro expansion. Source spellings remain on nodes.
//! Unicode reference: latex3/unicode-math/unicode-math-table.tex; existing
//! VisualTeX commandRegistry/additionalCommands determine the product names.
use crate::model::AtomClass;

#[derive(Debug, Clone, Copy)]
pub struct SymbolSpec {
    pub command: &'static str,
    pub value: &'static str,
    pub class: AtomClass,
}

/// A command-specific glyph alternate, separate from its semantic Unicode value.
/// These are real bundled AMS cmap slots, as used by MathLive's AMS definitions;
/// they are never substituted Unicode values or scaled ordinary symbol outlines.
#[derive(Debug, Clone, Copy)]
pub struct AlternateSymbolSpec {
    pub command: &'static str,
    pub value: &'static str,
    pub tex_font_role: &'static str,
    pub font_character: char,
}
pub const ALTERNATE_SYMBOLS: &[AlternateSymbolSpec] = &[
    AlternateSymbolSpec { command: "\\shortparallel", value: "∥", tex_font_role: "ams", font_character: '\u{2225}' },
    AlternateSymbolSpec { command: "\\nshortparallel", value: "∦", tex_font_role: "ams", font_character: '\u{E007}' },
];
pub fn alternate(command: &str) -> Option<&'static AlternateSymbolSpec> {
    ALTERNATE_SYMBOLS.iter().find(|spec| spec.command == command)
}

// The same declaration drives the public inventory and compiled lookup. This
// avoids a second hand-maintained list that can incorrectly report support.
macro_rules! symbols {
    ($( $($command:literal)|+ => ($value:literal, $class:ident) ),* $(,)?) => {
        pub const SYMBOLS: &[SymbolSpec] = &[
            $($(SymbolSpec { command: $command, value: $value, class: AtomClass::$class },)+)*
        ];
        pub fn lookup(command: &str) -> Option<(&'static str, AtomClass)> {
            match command {
                $($($command)|+ => Some(($value, AtomClass::$class)),)*
                _ => None,
            }
        }
    };
}

symbols! {
    "\\digamma" => ("ϝ", Ordinary),
    "\\Digamma" => ("Ϝ", Ordinary),
    "\\varkappa" => ("ϰ", Ordinary),
    "\\varTheta" => ("ϴ", Ordinary),
    "\\omicron" => ("ο", Ordinary),
    "\\aleph" => ("ℵ", Ordinary),
    "\\beth" => ("ℶ", Ordinary),
    "\\gimel" => ("ℷ", Ordinary),
    "\\daleth" => ("ℸ", Ordinary),
    "\\ell" => ("ℓ", Ordinary),
    "\\hbar" | "\\hslash" => ("ℏ", Ordinary),
    "\\Re" => ("ℜ", Ordinary),
    "\\Im" => ("ℑ", Ordinary),
    "\\wp" => ("℘", Ordinary),
    "\\imath" => ("ı", Ordinary),
    "\\jmath" => ("ȷ", Ordinary),
    "\\forall" => ("∀", Ordinary),
    "\\exists" => ("∃", Ordinary),
    "\\nexists" => ("∄", Ordinary),
    "\\emptyset" => ("∅", Ordinary),
    "\\neg" | "\\lnot" => ("¬", Ordinary),
    "\\top" => ("⊤", Ordinary),
    "\\bot" => ("⊥", Ordinary),
    "\\angle" => ("∠", Ordinary),
    "\\measuredangle" => ("∡", Ordinary),
    "\\sphericalangle" => ("∢", Ordinary),
    "\\Box" | "\\square" => ("□", Ordinary),
    "\\Diamond" => ("◇", Ordinary),
    "\\triangle" => ("△", Ordinary),
    "\\prime" => ("′", Ordinary),
    "\\backprime" => ("‵", Ordinary),
    "\\ldots" => ("…", Inner),
    "\\cdots" => ("⋯", Inner),
    "\\vdots" => ("⋮", Ordinary),
    "\\ddots" => ("⋱", Ordinary),

    "\\sim" => ("∼", Relation),
    "\\simeq" => ("≃", Relation),
    "\\cong" => ("≅", Relation),
    "\\propto" => ("∝", Relation),
    "\\ll" => ("≪", Relation),
    "\\gg" => ("≫", Relation),
    "\\parallel" => ("∥", Relation),
    "\\nparallel" => ("∦", Relation),
    "\\shortparallel" => ("∥", Relation),
    "\\nshortparallel" => ("∦", Relation),
    // txfonts/Fourier slanted relations use U+2AFD, not italic U+2225.
    // lwarp source manual: fourier and common-mathjax-newpxtxmath mappings.
    "\\varparallel" | "\\parallelslant" => ("⫽", Relation),
    // unicode-math assigns the same character a binary-operator spelling.
    "\\sslash" => ("⫽", Binary),
    "\\perp" => ("⊥", Relation),
    "\\mid" => ("∣", Relation),
    "\\nmid" => ("∤", Relation),
    "\\prec" => ("≺", Relation),
    "\\succ" => ("≻", Relation),
    "\\leqq" => ("≦", Relation),
    "\\geqq" => ("≧", Relation),
    "\\lneq" => ("⪇", Relation),
    "\\gneq" => ("⪈", Relation),
    "\\lneqq" => ("≨", Relation),
    "\\gneqq" => ("≩", Relation),
    "\\leqslant" => ("⩽", Relation),
    "\\geqslant" => ("⩾", Relation),
    "\\lesssim" => ("≲", Relation),
    "\\gtrsim" => ("≳", Relation),
    "\\lessapprox" => ("⪅", Relation),
    "\\gtrapprox" => ("⪆", Relation),
    "\\preceq" => ("⪯", Relation),
    "\\succeq" => ("⪰", Relation),
    "\\preccurlyeq" => ("≼", Relation),
    "\\succcurlyeq" => ("≽", Relation),
    "\\asymp" => ("≍", Relation),
    "\\doteq" => ("≐", Relation),
    "\\triangleq" => ("≜", Relation),
    "\\models" | "\\vDash" => ("⊨", Relation),
    "\\vdash" => ("⊢", Relation),
    "\\dashv" => ("⊣", Relation),
    "\\Vdash" => ("⊩", Relation),
    "\\bowtie" => ("⋈", Relation),
    "\\smile" => ("⌣", Relation),
    "\\frown" => ("⌢", Relation),
    "\\ni" | "\\owns" => ("∋", Relation),
    "\\subset" => ("⊂", Relation),
    "\\supset" => ("⊃", Relation),
    "\\subseteq" => ("⊆", Relation),
    "\\supseteq" => ("⊇", Relation),
    "\\subsetneq" => ("⊊", Relation),
    "\\supsetneq" => ("⊋", Relation),
    "\\nsubseteq" => ("⊈", Relation),
    "\\nsupseteq" => ("⊉", Relation),
    "\\sqsubset" => ("⊏", Relation),
    "\\sqsupset" => ("⊐", Relation),
    "\\sqsubseteq" => ("⊑", Relation),
    "\\sqsupseteq" => ("⊒", Relation),
    "\\nless" => ("≮", Relation),
    "\\ngtr" => ("≯", Relation),
    "\\nleq" => ("≰", Relation),
    "\\ngeq" => ("≱", Relation),

    "\\cup" => ("∪", Binary),
    "\\cap" => ("∩", Binary),
    "\\setminus" => ("∖", Binary),
    "\\land" | "\\wedge" => ("∧", Binary),
    "\\lor" | "\\vee" => ("∨", Binary),
    "\\uplus" => ("⊎", Binary),
    "\\sqcup" => ("⊔", Binary),
    "\\sqcap" => ("⊓", Binary),
    "\\oplus" => ("⊕", Binary),
    "\\ominus" => ("⊖", Binary),
    "\\otimes" => ("⊗", Binary),
    "\\oslash" => ("⊘", Binary),
    "\\odot" => ("⊙", Binary),
    "\\circ" => ("∘", Binary),
    "\\bullet" => ("∙", Binary),
    "\\ast" => ("∗", Binary),
    "\\star" => ("⋆", Binary),
    "\\diamond" => ("⋄", Binary),
    "\\dagger" => ("†", Binary),
    "\\ddagger" => ("‡", Binary),
    "\\wr" => ("≀", Binary),
    "\\amalg" => ("⨿", Binary),

    "\\leftrightarrow" => ("↔", Relation),
    "\\uparrow" => ("↑", Relation),
    "\\downarrow" => ("↓", Relation),
    "\\updownarrow" => ("↕", Relation),
    "\\Leftarrow" => ("⇐", Relation),
    "\\Rightarrow" => ("⇒", Relation),
    "\\Leftrightarrow" => ("⇔", Relation),
    "\\Uparrow" => ("⇑", Relation),
    "\\Downarrow" => ("⇓", Relation),
    "\\Updownarrow" => ("⇕", Relation),
    "\\mapsto" => ("↦", Relation),
    "\\longrightarrow" => ("⟶", Relation),
    "\\longleftarrow" => ("⟵", Relation),
    "\\longleftrightarrow" => ("⟷", Relation),
    "\\Longrightarrow" => ("⟹", Relation),
    "\\Longleftarrow" => ("⟸", Relation),
    "\\Longleftrightarrow" => ("⟺", Relation),
    "\\longmapsto" => ("⟼", Relation),
    "\\hookrightarrow" => ("↪", Relation),
    "\\hookleftarrow" => ("↩", Relation),
    "\\rightharpoonup" => ("⇀", Relation),
    "\\rightharpoondown" => ("⇁", Relation),
    "\\leftharpoonup" => ("↼", Relation),
    "\\leftharpoondown" => ("↽", Relation),
    "\\rightleftarrows" => ("⇄", Relation),
    "\\leftrightarrows" => ("⇆", Relation),
    "\\rightleftharpoons" => ("⇌", Relation),
    "\\leftrightharpoons" => ("⇋", Relation),
    "\\nearrow" => ("↗", Relation),
    "\\nwarrow" => ("↖", Relation),
    "\\searrow" => ("↘", Relation),
    "\\swarrow" => ("↙", Relation),
    "\\circlearrowright" => ("↻", Relation),
    "\\circlearrowleft" => ("↺", Relation),
    "\\rightsquigarrow" | "\\leadsto" => ("⇝", Relation),
    "\\twoheadrightarrow" => ("↠", Relation),
    "\\twoheadleftarrow" => ("↞", Relation),

    "\\langle" => ("⟨", Open),
    "\\rangle" => ("⟩", Close),
    "\\lfloor" => ("⌊", Open),
    "\\rfloor" => ("⌋", Close),
    "\\lceil" => ("⌈", Open),
    "\\rceil" => ("⌉", Close),
    "\\lbrace" => ("{", Open),
    "\\rbrace" => ("}", Close),
    "\\lbrack" => ("[", Open),
    "\\rbrack" => ("]", Close),
    "\\lvert" => ("|", Open),
    "\\rvert" => ("|", Close),
    "\\vert" => ("|", Ordinary),
    "\\lVert" => ("‖", Open),
    "\\rVert" => ("‖", Close),
    "\\Vert" => ("‖", Ordinary),
    "\\backslash" => ("\\", Ordinary),
}

/// Bare Unicode has no left/right delimiter spelling. Keep ambiguous bars and
/// truth/perpendicular glyphs ordinary; named commands retain their own class.
/// ASCII typing takes the existing constant-time path and never scans this list.
pub fn literal_class(value: &str) -> Option<AtomClass> {
    if matches!(value, "|" | "‖" | "⊥") { return Some(AtomClass::Ordinary); }
    // Raw U+2AFD follows its Unicode operator class, independent of whether a
    // relation alias appears first in the command inventory above.
    if value == "⫽" { return Some(AtomClass::Binary); }
    if value.is_ascii() || value.chars().count() != 1 { return None; }
    SYMBOLS.iter().find(|entry| entry.value == value).map(|entry| entry.class)
}
