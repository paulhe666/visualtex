use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum PhantomAxis { Full, Horizontal, Vertical }

/// TeX-like explicit spacing in math units (18mu = 1em). The command spelling
/// is retained by the node; this table is the engine's layout semantic.
pub fn space_mu(command: &str) -> Option<i16> {
    Some(match command {
        "\\," | "\\thinspace" => 3,
        "\\:" | "\\medspace" => 4,
        "\\;" | "\\thickspace" => 5,
        "\\!" | "\\negthinspace" => -3,
        "\\negmedspace" => -4,
        "\\negthickspace" => -5,
        "\\enspace" | "\\enskip" => 9,
        "\\quad" => 18,
        "\\qquad" => 36,
        _ => return None,
    })
}

/// Mu glue follows the mathematical script style. Explicit em skips retain
/// the surrounding semantic font size (the TeX hskip distinction).
pub fn space_scale(command: &str, math_scale: f64) -> f64 {
    if matches!(command, "\\quad" | "\\qquad" | "\\enspace" | "\\enskip") { 1.0 } else { math_scale }
}

pub fn phantom_axis(command: &str) -> Option<PhantomAxis> {
    Some(match command {
        "\\phantom" => PhantomAxis::Full,
        "\\hphantom" => PhantomAxis::Horizontal,
        "\\vphantom" => PhantomAxis::Vertical,
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn spacing_table_uses_math_units_and_keeps_negative_space_explicit() {
        assert_eq!(space_mu("\\,"), Some(3));
        assert_eq!(space_mu("\\quad"), Some(18));
        assert_eq!(space_mu("\\qquad"), Some(36));
        assert_eq!(space_mu("\\!"), Some(-3));
        assert_eq!(space_mu("\\unknown"), None);
    }
    #[test]
    fn phantom_commands_have_independent_axes() {
        assert_eq!(phantom_axis("\\phantom"), Some(PhantomAxis::Full));
        assert_eq!(phantom_axis("\\hphantom"), Some(PhantomAxis::Horizontal));
        assert_eq!(phantom_axis("\\vphantom"), Some(PhantomAxis::Vertical));
    }
}
