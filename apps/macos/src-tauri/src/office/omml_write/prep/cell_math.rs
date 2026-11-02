//! MathML for an equation in a PowerPoint table cell (PowerPointCellMath in the
//! Windows add-in). A cell takes an equation only through PowerPoint's linear
//! format, which keeps the characters of an equation but not their formatting:
//! a letter's style must be in the character itself. Upright letters become
//! normal text (quoted in linear format) and bold, script, fraktur,
//! double-struck, sans-serif and monospace letters become Unicode's
//! mathematical alphanumerics. Colors are not kept.

use super::dom::{Doc, Name, NodeId, MATHML_NS};
use super::xlinq::{self, XDOCUMENT_PARSE_PRESERVE};

fn is_mathml(doc: &Doc, node: NodeId, local: &str) -> bool {
    doc.name(node).is_some_and(|name| (name.ns == MATHML_NS || name.ns.is_empty()) && name.local == local)
}

pub fn prepare(mathml: &str) -> Result<String, String> {
    let mut document = xlinq::parse(mathml, &XDOCUMENT_PARSE_PRESERVE)
        .map_err(|error| format!("The formula's MathML is not valid XML: {error}"))?;
    let tokens: Vec<NodeId> = document
        .descendants()
        .into_iter()
        .filter(|&node| ["mi", "mn", "mtext", "ms"].iter().any(|local| is_mathml(&document.doc, node, local)))
        .collect();
    for node in tokens {
        let doc = &mut document.doc;
        let variant = inherited_variant(doc, node);
        let text = doc.value(node);
        let local = doc.local(node).to_string();
        if local == "mi" && is_upright(variant.as_deref(), &text) && text.chars().any(char::is_alphabetic) {
            // A function name before U+2061 stays an identifier: PowerPoint keeps
            // functions upright itself.
            if is_function_name(doc, node) {
                continue;
            }
            let ns = doc.name(node).map(|name| name.ns.clone()).unwrap_or_default();
            doc.set_name(node, Name::new(&ns, "mtext"));
            doc.remove_attr(node, "mathvariant");
            continue;
        }
        let Some(variant) = variant else { continue };
        if matches!(variant.as_str(), "normal" | "italic") || local == "mtext" || local == "ms" {
            continue;
        }
        let Some(styled) = styled(&text, &variant) else { continue };
        doc.set_value(node, &styled);
        doc.remove_attr(node, "mathvariant");
        // An identifier spelled in mathematical alphanumerics is never italicized again.
        if local == "mi" && text.chars().count() == 1 {
            doc.set_attr(node, "mathvariant", "normal");
        }
    }
    document.to_string()
}

/// A multi-letter identifier is upright unless styled otherwise; a single letter
/// only when marked normal.
fn is_upright(variant: Option<&str>, text: &str) -> bool {
    variant == Some("normal") || variant.is_none() && text.trim().chars().count() > 1
}

fn is_function_name(doc: &Doc, identifier: NodeId) -> bool {
    doc.next_element(identifier)
        .is_some_and(|next| is_mathml(doc, next, "mo") && doc.value(next).trim() == "\u{2061}")
}

fn inherited_variant(doc: &Doc, node: NodeId) -> Option<String> {
    let mut current = Some(node);
    while let Some(element) = current {
        if let Some(variant) = doc.attr(element, "mathvariant") {
            return Some(variant.to_string());
        }
        if is_mathml(doc, element, "math") {
            break;
        }
        current = doc.parent(element);
    }
    None
}

/// Unicode's Mathematical Alphanumeric Symbols for a MathML mathvariant, None when
/// the variant has no letters there. Letters with an earlier code point (the
/// Letterlike Symbols holes) use it.
pub fn styled(text: &str, variant: &str) -> Option<String> {
    let (upper, lower, digit): (Option<u32>, Option<u32>, Option<u32>) = match variant {
        "bold" => (Some(0x1D400), Some(0x1D41A), Some(0x1D7CE)),
        "bold-italic" => (Some(0x1D468), Some(0x1D482), Some(0x1D7CE)),
        "script" => (Some(0x1D49C), Some(0x1D4B6), None),
        "bold-script" => (Some(0x1D4D0), Some(0x1D4EA), Some(0x1D7CE)),
        "fraktur" => (Some(0x1D504), Some(0x1D51E), None),
        "double-struck" => (Some(0x1D538), Some(0x1D552), Some(0x1D7D8)),
        "bold-fraktur" => (Some(0x1D56C), Some(0x1D586), Some(0x1D7CE)),
        "sans-serif" => (Some(0x1D5A0), Some(0x1D5BA), Some(0x1D7E2)),
        "bold-sans-serif" => (Some(0x1D5D4), Some(0x1D5EE), Some(0x1D7EC)),
        "sans-serif-italic" => (Some(0x1D608), Some(0x1D622), Some(0x1D7E2)),
        "sans-serif-bold-italic" => (Some(0x1D63C), Some(0x1D656), Some(0x1D7EC)),
        "monospace" => (Some(0x1D670), Some(0x1D68A), Some(0x1D7F6)),
        _ => return None,
    };
    let mut result = String::new();
    for character in text.chars() {
        let code = match character {
            'A'..='Z' => upper.map(|base| base + (character as u32 - 'A' as u32)),
            'a'..='z' => lower.map(|base| base + (character as u32 - 'a' as u32)),
            '0'..='9' => digit.map(|base| base + (character as u32 - '0' as u32)),
            _ => None,
        };
        match code {
            None => result.push(character),
            Some(code) => match hole(code) {
                Some(letter) => result.push(letter),
                None => result.push(char::from_u32(code).unwrap_or(character)),
            },
        }
    }
    Some(result)
}

fn hole(code: u32) -> Option<char> {
    Some(match code {
        0x1D49D => 'ℬ', 0x1D4A0 => 'ℰ', 0x1D4A1 => 'ℱ', 0x1D4A3 => 'ℋ', 0x1D4A4 => 'ℐ',
        0x1D4A7 => 'ℒ', 0x1D4A8 => 'ℳ', 0x1D4AD => 'ℛ', 0x1D4BA => 'ℯ', 0x1D4BC => 'ℊ',
        0x1D4C4 => 'ℴ', 0x1D506 => 'ℭ', 0x1D50B => 'ℌ', 0x1D50C => 'ℑ', 0x1D515 => 'ℜ',
        0x1D51D => 'ℨ', 0x1D53A => 'ℂ', 0x1D53F => 'ℍ', 0x1D545 => 'ℕ', 0x1D547 => 'ℙ',
        0x1D548 => 'ℚ', 0x1D549 => 'ℝ', 0x1D551 => 'ℤ',
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn tokens(body: &str) -> Vec<(String, String, Option<String>)> {
        let prepared = prepare(&format!("<math xmlns=\"{MATHML_NS}\">{body}</math>")).unwrap();
        let document = xlinq::parse(&prepared, &XDOCUMENT_PARSE_PRESERVE).unwrap();
        document
            .descendants()
            .into_iter()
            .filter(|&node| ["mi", "mn", "mtext"].contains(&document.doc.local(node)))
            .map(|node| {
                let doc = &document.doc;
                (doc.local(node).to_string(), doc.value(node), doc.attr(node, "mathvariant").map(str::to_string))
            })
            .collect()
    }

    #[test]
    fn upright_letters_become_normal_text() {
        let found = tokens("<msub><mi mathvariant=\"normal\">M</mi><mrow><mi mathvariant=\"normal\">Na</mi></mrow></msub><mi>abc</mi>");
        assert_eq!(
            found,
            vec![
                ("mtext".into(), "M".into(), None),
                ("mtext".into(), "Na".into(), None),
                ("mtext".into(), "abc".into(), None),
            ]
        );
    }

    #[test]
    fn function_names_and_italic_letters_stay_identifiers() {
        let found = tokens("<mi>sin</mi><mo>\u{2061}</mo><mi>x</mi><mi mathvariant=\"normal\">cos</mi><mo>\u{2061}</mo><mi>y</mi>");
        let names: Vec<_> = found.iter().map(|token| (token.0.as_str(), token.1.as_str())).collect();
        assert_eq!(names, vec![("mi", "sin"), ("mi", "x"), ("mi", "cos"), ("mi", "y")]);
    }

    #[test]
    fn styled_letters_are_spelled_in_mathematical_alphanumerics() {
        let found = tokens(
            "<mi mathvariant=\"bold\">v</mi><mi mathvariant=\"double-struck\">R</mi><mi mathvariant=\"script\">B</mi>\
             <mi mathvariant=\"fraktur\">g</mi><mn mathvariant=\"bold\">12</mn><mstyle mathvariant=\"bold\"><mi>A</mi></mstyle>",
        );
        let values: Vec<_> = found.iter().map(|token| token.1.as_str()).collect();
        assert_eq!(values, vec!["\u{1D42F}", "ℝ", "ℬ", "\u{1D524}", "\u{1D7CF}\u{1D7D0}", "\u{1D400}"]);
        // A single styled letter is not italicized again.
        assert_eq!(found[0].2.as_deref(), Some("normal"));
    }

    #[test]
    fn non_letters_and_unknown_variants_are_kept() {
        assert_eq!(styled("α+", "bold").as_deref(), Some("α+"));
        assert_eq!(styled("x", "stretched"), None);
        assert!(prepare("<math").is_err());
    }
}
