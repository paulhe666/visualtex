//! The space a field holds in its text for a formula that VisualTeX draws over it (the formula's
//! composition), as the macOS input method holds it (input-method/Sources/FormulaSpace.swift).
//! VisualTeX cuts the formula where TeX may break an in-line formula (after a relation or a binary
//! operator) and tells the width of each piece; each piece is held by blank characters as wide as
//! it and the pieces are joined by word spaces, where the field may break the line. So the field
//! wraps the formula as it wraps its own text, in its own box, and pushes the text after the caret
//! along; the engine reads back where each piece went and VisualTeX draws it there.
//! (The Windows input method's tip/src/space.rs; its characters are all in the BMP, so a UTF-16 index
//! here is also the character index IBus counts the preedit text's cursor in.)
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
/// `margin`: fillers the last piece is held short of its width besides, for widths not measured yet.
/// `line_start`: the piece the field began the last line with (both the field and VisualTeX set it at
/// the line's start, so the pieces from it are held from there). The last piece is held to its width
/// at most, in whole fillers, so the held space never ends past the formula (the field's underline
/// and caret stood past it, macOS 2026-11-10); a last piece narrower than a filler is held with the
/// one before (not the piece a line begins with).
pub(crate) fn layout(widths: &[f64], unit: f64, joint: f64, margin: f64, line_start: usize) -> Layout {
    let unit = if unit.is_finite() && unit > 0.5 { unit } else { 4.0 };
    let joint = if joint.is_finite() && (0.0..60.0).contains(&joint) { joint } else { unit };
    let margin = if margin.is_finite() { margin.max(0.0) } else { 0.0 };
    let mut pieces: Vec<f64> = if widths.is_empty() { vec![0.0] } else { widths.iter().map(|&w| if w.is_finite() && w > 0.0 { w } else { 0.0 }).collect() };
    loop {
        let mut out = Layout { text: vec![ANCHOR], ..Default::default() };
        let (mut formula, mut held) = (0.0f64, 0.0f64);
        let mut short = false;
        for (index, &width) in pieces.iter().enumerate() {
            let last = index + 1 == pieces.len();
            out.starts.push(out.text.len());
            if index == line_start && index > 0 {
                formula = 0.0;
                held = 0.0;
            }
            formula += width;
            let wanted = (formula - held - if last { 0.0 } else { joint }) / unit;
            // (a hair under a whole filler is a whole one: rounding of the widths)
            let whole = if last { (wanted - margin + 0.02).floor() } else { wanted.round() };
            if last && whole < 1.0 && pieces.len() > 1 && index != line_start {
                short = true;
            }
            let count = if whole.is_finite() && whole >= 1.0 { (whole as usize).min(4096) } else { 1 };
            out.counts.push(count);
            held += count as f64 * unit + if last { 0.0 } else { joint };
            out.text.extend(std::iter::repeat(FILLER).take(count));
            if !last {
                out.text.push(JOINT);
            }
        }
        if !short {
            return out;
        }
        let tail = pieces.pop().unwrap_or(0.0);
        if let Some(before) = pieces.last_mut() {
            *before += tail;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn an_empty_formula_holds_one_filler_after_its_anchor() {
        let held = layout(&[], 5.0, 4.0, 0.0, 0);
        assert_eq!(held.text, vec![ANCHOR, FILLER]);
        assert_eq!(held.starts, vec![1]);
    }

    #[test]
    fn pieces_start_within_half_a_filler_of_the_formula() {
        let widths = [33.8, 34.02, 31.95, 33.64, 32.58, 37.27, 33.49, 30.43];
        let (unit, joint) = (7.5, 4.25);
        let held = layout(&widths, unit, joint, 0.0, 0);
        assert_eq!(held.starts.len(), widths.len());
        let mut formula = 0.0f64;
        let mut place = 0.0f64;
        for (index, width) in widths.iter().enumerate() {
            assert!((place - formula).abs() <= unit / 2.0 + 1e-9, "piece {index}: {place} vs {formula}");
            formula += width;
            place += held.counts[index] as f64 * unit + if index + 1 < widths.len() { joint } else { 0.0 };
        }
        // the last piece is held whole: the formula never covers what follows
        // the held space ends within the formula, less than a filler short of it
        assert!(place <= formula + 1e-9 && place > formula - unit, "{place} vs {formula}");
        // joints between pieces only
        assert_eq!(held.text.iter().filter(|&&c| c == JOINT).count(), widths.len() - 1);
    }

    #[test]
    fn a_last_piece_narrower_than_a_filler_is_held_with_the_one_before() {
        let held = layout(&[30.0, 30.0, 4.0], 10.0, 4.0, 0.0, 0);
        assert_eq!(held.starts.len(), 2);
        let total: f64 = held.counts.iter().sum::<usize>() as f64 * 10.0 + 4.0;
        assert!(total <= 64.0 + 1e-9, "{total}");
    }

    #[test]
    fn the_last_line_is_held_from_where_it_begins() {
        // the field began a line with piece 2: the pieces from it are held from 0 again
        let held = layout(&[33.0, 33.0, 25.0, 26.0], 10.0, 4.0, 0.0, 2);
        let line: f64 = held.counts[2] as f64 * 10.0 + 4.0 + held.counts[3] as f64 * 10.0;
        assert!(line <= 51.0 + 1e-9, "{line}");
        // a piece a line begins with is not held with the one before it, however narrow
        let held = layout(&[33.0, 33.0, 4.0], 10.0, 4.0, 0.0, 2);
        assert_eq!(held.starts.len(), 3);
    }

    #[test]
    fn widths_not_measured_hold_a_margin() {
        let held = layout(&[95.0], 10.0, 4.0, 0.6, 0);
        assert_eq!(held.counts, vec![8]);
    }

    #[test]
    fn every_piece_has_a_filler() {
        let held = layout(&[0.5, 0.5, 0.5], 8.0, 4.0, 0.0, 0);
        assert!(held.counts.iter().all(|&count| count >= 1));
        for (index, &start) in held.starts.iter().enumerate() {
            assert_eq!(held.text[start], FILLER, "piece {index}");
        }
    }
}
