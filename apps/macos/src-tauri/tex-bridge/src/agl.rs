//! Glyph names to Unicode (the Adobe Glyph List's rules and its common names): what a Type 1 font's character
//! stands for, so that the caret's places in a live paragraph can be found from its glyphs.

/// The characters glyph `name` stands for ("" when unknown): uniXXXX, uXXXX[XX], the common Latin, Greek and
/// punctuation names, ligatures as their letters.
pub fn chars(name: &str) -> String {
    let base = name.split('.').next().unwrap_or(name);
    // ligature names joined with "_" (f_f_i)
    if base.contains('_') {
        return base.split('_').map(chars).collect();
    }
    if let Some(hex) = base.strip_prefix("uni") {
        if hex.len() >= 4 && hex.len() % 4 == 0 && hex.chars().all(|c| c.is_ascii_hexdigit()) {
            return (0..hex.len() / 4)
                .filter_map(|i| u32::from_str_radix(&hex[i * 4..i * 4 + 4], 16).ok().and_then(char::from_u32))
                .collect();
        }
    }
    if let Some(hex) = base.strip_prefix('u') {
        if (4..=6).contains(&hex.len()) && hex.chars().all(|c| c.is_ascii_hexdigit()) {
            return u32::from_str_radix(hex, 16).ok().and_then(char::from_u32).map(String::from).unwrap_or_default();
        }
    }
    if base.len() == 1 && base.chars().all(|c| c.is_ascii_alphabetic()) {
        return base.to_string();
    }
    let known = match base {
        "space" => " ",
        "exclam" => "!",
        "quotedbl" => "\"",
        "numbersign" => "#",
        "dollar" => "$",
        "percent" => "%",
        "ampersand" => "&",
        "quotesingle" => "'",
        "quoteright" => "\u{2019}",
        "quoteleft" => "\u{2018}",
        "quotedblright" => "\u{201D}",
        "quotedblleft" => "\u{201C}",
        "quotesinglbase" => "\u{201A}",
        "quotedblbase" => "\u{201E}",
        "parenleft" => "(",
        "parenright" => ")",
        "asterisk" | "asteriskmath" => "*",
        "plus" => "+",
        "comma" => ",",
        "hyphen" | "sfthyphen" => "-",
        "period" => ".",
        "slash" => "/",
        "zero" => "0",
        "one" => "1",
        "two" => "2",
        "three" => "3",
        "four" => "4",
        "five" => "5",
        "six" => "6",
        "seven" => "7",
        "eight" => "8",
        "nine" => "9",
        "colon" => ":",
        "semicolon" => ";",
        "less" => "<",
        "equal" => "=",
        "greater" => ">",
        "question" => "?",
        "at" => "@",
        "bracketleft" => "[",
        "backslash" => "\\",
        "bracketright" => "]",
        "asciicircum" | "circumflex" => "^",
        "underscore" => "_",
        "grave" => "`",
        "braceleft" => "{",
        "bar" => "|",
        "braceright" => "}",
        "asciitilde" | "tilde" => "~",
        "endash" => "\u{2013}",
        "emdash" => "\u{2014}",
        "ellipsis" => "\u{2026}",
        "bullet" => "\u{2022}",
        "dagger" => "\u{2020}",
        "daggerdbl" => "\u{2021}",
        "section" => "\u{A7}",
        "paragraph" => "\u{B6}",
        "exclamdown" => "\u{A1}",
        "questiondown" => "\u{BF}",
        "guillemotleft" | "guillemetleft" => "\u{AB}",
        "guillemotright" | "guillemetright" => "\u{BB}",
        "guilsinglleft" => "\u{2039}",
        "guilsinglright" => "\u{203A}",
        "ff" => "ff",
        "fi" => "fi",
        "fl" => "fl",
        "ffi" => "ffi",
        "ffl" => "ffl",
        "dotlessi" => "\u{131}",
        "dotlessj" => "\u{237}",
        "germandbls" => "\u{DF}",
        "ae" => "\u{E6}",
        "AE" => "\u{C6}",
        "oe" => "\u{153}",
        "OE" => "\u{152}",
        "oslash" => "\u{F8}",
        "Oslash" => "\u{D8}",
        "lslash" => "\u{142}",
        "Lslash" => "\u{141}",
        "eth" => "\u{F0}",
        "Eth" => "\u{D0}",
        "thorn" => "\u{FE}",
        "Thorn" => "\u{DE}",
        "aring" => "\u{E5}",
        "Aring" => "\u{C5}",
        "copyright" => "\u{A9}",
        "registered" => "\u{AE}",
        "trademark" => "\u{2122}",
        "degree" => "\u{B0}",
        "sterling" => "\u{A3}",
        "yen" => "\u{A5}",
        "Euro" | "euro" => "\u{20AC}",
        "cent" => "\u{A2}",
        "periodcentered" => "\u{B7}",
        "multiply" => "\u{D7}",
        "divide" => "\u{F7}",
        "minus" => "\u{2212}",
        "plusminus" => "\u{B1}",
        "minusplus" => "\u{2213}",
        "infinity" => "\u{221E}",
        "partialdiff" => "\u{2202}",
        "element" => "\u{2208}",
        "owner" => "\u{220B}",
        "universal" => "\u{2200}",
        "existential" => "\u{2203}",
        "emptyset" => "\u{2205}",
        "prime" => "\u{2032}",
        "union" => "\u{222A}",
        "intersection" => "\u{2229}",
        "integral" => "\u{222B}",
        "summation" => "\u{2211}",
        "product" => "\u{220F}",
        "radical" => "\u{221A}",
        "arrowright" => "\u{2192}",
        "arrowleft" => "\u{2190}",
        "arrowup" => "\u{2191}",
        "arrowdown" => "\u{2193}",
        "arrowboth" => "\u{2194}",
        "arrowdblright" => "\u{21D2}",
        "arrowdblleft" => "\u{21D0}",
        "arrowdblboth" => "\u{21D4}",
        "lessequal" => "\u{2264}",
        "greaterequal" => "\u{2265}",
        "notequal" => "\u{2260}",
        "approxequal" => "\u{2248}",
        "equivalence" => "\u{2261}",
        "propersubset" => "\u{2282}",
        "propersuperset" => "\u{2283}",
        "reflexsubset" => "\u{2286}",
        "reflexsuperset" => "\u{2287}",
        "logicaland" => "\u{2227}",
        "logicalor" => "\u{2228}",
        "logicalnot" => "\u{AC}",
        "nabla" | "gradient" => "\u{2207}",
        "angleleft" => "\u{27E8}",
        "angleright" => "\u{27E9}",
        "Gamma" => "\u{393}",
        "Delta" => "\u{394}",
        "Theta" => "\u{398}",
        "Lambda" => "\u{39B}",
        "Xi" => "\u{39E}",
        "Pi" => "\u{3A0}",
        "Sigma" => "\u{3A3}",
        "Upsilon" | "Upsilon1" => "\u{3A5}",
        "Phi" => "\u{3A6}",
        "Psi" => "\u{3A8}",
        "Omega" => "\u{3A9}",
        "alpha" => "\u{3B1}",
        "beta" => "\u{3B2}",
        "gamma" => "\u{3B3}",
        "delta" => "\u{3B4}",
        "epsilon" | "epsilon1" => "\u{3B5}",
        "zeta" => "\u{3B6}",
        "eta" => "\u{3B7}",
        "theta" | "theta1" => "\u{3B8}",
        "iota" => "\u{3B9}",
        "kappa" => "\u{3BA}",
        "lambda" => "\u{3BB}",
        "mu" => "\u{3BC}",
        "nu" => "\u{3BD}",
        "xi" => "\u{3BE}",
        "omicron" => "\u{3BF}",
        "pi" | "pi1" => "\u{3C0}",
        "rho" | "rho1" => "\u{3C1}",
        "sigma" | "sigma1" => "\u{3C3}",
        "tau" => "\u{3C4}",
        "upsilon" => "\u{3C5}",
        "phi" | "phi1" => "\u{3C6}",
        "chi" => "\u{3C7}",
        "psi" => "\u{3C8}",
        "omega" => "\u{3C9}",
        _ => "",
    };
    if !known.is_empty() {
        return known.to_string();
    }
    // accented Latin letters: base letter and accent name (eacute, Udieresis, ccedilla)
    const ACCENTS: [(&str, [char; 6]); 8] = [
        ("acute", ['\u{301}'; 6]),
        ("grave", ['\u{300}'; 6]),
        ("circumflex", ['\u{302}'; 6]),
        ("dieresis", ['\u{308}'; 6]),
        ("tilde", ['\u{303}'; 6]),
        ("cedilla", ['\u{327}'; 6]),
        ("caron", ['\u{30C}'; 6]),
        ("ring", ['\u{30A}'; 6]),
    ];
    for (suffix, mark) in ACCENTS {
        if let Some(letter) = base.strip_suffix(suffix) {
            if letter.len() == 1 && letter.chars().all(|c| c.is_ascii_alphabetic()) {
                return format!("{letter}{}", mark[0]);
            }
        }
    }
    String::new()
}

#[cfg(test)]
mod tests {
    use super::chars;

    #[test]
    fn names() {
        assert_eq!(chars("a"), "a");
        assert_eq!(chars("fi"), "fi");
        assert_eq!(chars("uni4E00"), "\u{4E00}");
        assert_eq!(chars("u1D400"), "\u{1D400}");
        assert_eq!(chars("f_f_i"), "ffi");
        assert_eq!(chars("endash"), "\u{2013}");
        assert_eq!(chars("eacute"), "e\u{301}");
        assert_eq!(chars("alpha"), "\u{3B1}");
        assert_eq!(chars("zero.oldstyle"), "0");
        assert_eq!(chars(".notdef"), "");
    }
}
