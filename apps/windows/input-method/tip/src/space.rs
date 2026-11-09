//! The space a field holds in its text for a formula that VisualTeX draws over it (the formula's
//! composition), as the macOS input method holds it (input-method/Sources/FormulaSpace.swift).
//! VisualTeX cuts the formula where TeX may break an in-line formula (after a relation or a binary
//! operator) and tells the width of each piece; each piece is held by blank characters as wide as
//! it and the pieces are joined by word spaces, where the field may break the line. So the field
//! wraps the formula as it wraps its own text, in its own box, and pushes the text after the caret
//! along; the TIP reads back where each piece went and VisualTeX draws it there.
//!
//! The composition starts with one hair space, selected: being selected (not a caret) it hides the
//! field's caret, so the formula's caret is the only one, and it lets the field break the line
//! before the formula.

/// The selected anchor before the formula.
pub(crate) const ANCHOR: u16 = 0x200A;
/// Braille pattern blank: drawn as nothing, and a letter to line breaking (no break inside a run
/// of them, a break at the word space after one). A no-break space glues to its neighbours in
/// some text systems (no break before or after it).
pub(crate) const FILLER: u16 = 0x2800;
/// A word space: where the field may break the line between pieces.
pub(crate) const JOINT: u16 = 0x0020;

#[derive(Clone, Debug, Default, PartialEq)]
pub(crate) struct Layout {
    pub(crate) text: Vec<u16>,
    /// Each piece's first character (UTF-16 index in `text`).
    pub(crate) starts: Vec<usize>,
    /// How many fillers hold each piece (a piece before another also ends with a joint).
    pub(crate) counts: Vec<usize>,
}

/// The composition holding pieces `widths` pixels wide with fillers `unit` pixels wide and joints
/// `joint` pixels wide, at least one filler per piece. The fillers follow the formula as a whole:
/// where each piece starts in the held space is never more than half a filler from where it starts
/// in the formula, so on a line of several pieces the space the field holds is the formula's
/// (VisualTeX draws the pieces of one line as the formula spaces them); the last piece is held
/// whole and a little more, so the formula never covers the text after it.
pub(crate) fn layout(widths: &[f64], unit: f64, joint: f64) -> Layout {
    let unit = if unit.is_finite() && unit > 0.5 { unit } else { 4.0 };
    let joint = if joint.is_finite() && (0.0..60.0).contains(&joint) { joint } else { unit };
    let mut out = Layout { text: vec![ANCHOR], ..Default::default() };
    let pieces: &[f64] = if widths.is_empty() { &[0.0] } else { widths };
    let (mut formula, mut held) = (0.0f64, 0.0f64);
    for (index, &width) in pieces.iter().enumerate() {
        let last = index + 1 == pieces.len();
        out.starts.push(out.text.len());
        formula += if width.is_finite() && width > 0.0 { width } else { 0.0 };
        let wanted = (formula - held - if last { 0.0 } else { joint }) / unit;
        // the last piece with a third of a filler to spare: a field may draw the fillers a little
        // narrower than they were measured
        let count = if last { (wanted + 0.3).ceil() } else { wanted.round() };
        let count = if count.is_finite() && count >= 1.0 { (count as usize).min(4096) } else { 1 };
        out.counts.push(count);
        held += count as f64 * unit + if last { 0.0 } else { joint };
        out.text.extend(std::iter::repeat(FILLER).take(count));
        if !last {
            out.text.push(JOINT);
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn an_empty_formula_holds_one_filler_after_its_anchor() {
        let held = layout(&[], 5.0, 4.0);
        assert_eq!(held.text, vec![ANCHOR, FILLER]);
        assert_eq!(held.starts, vec![1]);
    }

    #[test]
    fn pieces_start_within_half_a_filler_of_the_formula() {
        let widths = [33.8, 34.02, 31.95, 33.64, 32.58, 37.27, 33.49, 30.43];
        let (unit, joint) = (7.5, 4.25);
        let held = layout(&widths, unit, joint);
        assert_eq!(held.starts.len(), widths.len());
        let mut formula = 0.0f64;
        let mut place = 0.0f64;
        for (index, width) in widths.iter().enumerate() {
            assert!((place - formula).abs() <= unit / 2.0 + 1e-9, "piece {index}: {place} vs {formula}");
            formula += width;
            place += held.counts[index] as f64 * unit + if index + 1 < widths.len() { joint } else { 0.0 };
        }
        // the last piece is held whole: the formula never covers what follows
        assert!(place >= formula + 0.3 * unit - 1e-9);
        // joints between pieces only
        assert_eq!(held.text.iter().filter(|&&c| c == JOINT).count(), widths.len() - 1);
    }

    #[test]
    fn every_piece_has_a_filler() {
        let held = layout(&[0.5, 0.5, 0.5], 8.0, 4.0);
        assert!(held.counts.iter().all(|&count| count >= 1));
        for (index, &start) in held.starts.iter().enumerate() {
            assert_eq!(held.text[start], FILLER, "piece {index}");
        }
    }
}
