import Cocoa

/// The space a field holds in its text for a formula that VisualTeX draws
/// over it: the formula's marked text. VisualTeX cuts the formula where TeX
/// may break an in-line formula (after a relation or a binary operator) and
/// tells the width of each piece; each piece is held by blank characters as
/// wide as it, the pieces are joined by word spaces, where the field may
/// break the line (every text system breaks there). A joint counts in the
/// piece before it. So the field wraps the formula as it wraps its own text,
/// in its own box, and pushes the text after the caret along; the input
/// method reads back where each piece went and VisualTeX draws it there,
/// over the field's background (what the field draws for a composition, its
/// underline, the selected anchor and its caret, is covered).
///
/// The marked text starts with one hair space, selected: being selected (not
/// a caret) it hides the field's own caret, so the formula's caret is the
/// only one, and it lets the field break the line before the formula.
enum FormulaSpace {
    static let anchor = "\u{200A}"
    /// Braille pattern blank: drawn as nothing, and a letter to line breaking
    /// (no break inside a run of them, a break at the word space after one).
    /// (A no-break space would do in browsers, but TextKit, in TextEdit and
    /// Notes, then breaks the line nowhere in the held space.)
    static let filler = "\u{2800}"
    /// A word space: where the field may break the line between pieces.
    static let joint = " "

    struct Layout: Equatable {
        var text: String
        /// Each piece's first character, in UTF-16 units of `text`.
        var starts: [Int]
        /// How many fillers hold each piece (a piece before another also ends
        /// with a joint).
        var counts: [Int]
    }

    /// The marked text holding pieces `widths` points wide with fillers
    /// `unit` points wide each, at least one per piece. The fillers follow
    /// the formula as a whole: where each piece starts in the held space is
    /// never more than half a filler from where it starts in the formula, so
    /// on a line of several pieces the space the field holds is the formula's
    /// (VisualTeX draws the pieces of one line as the formula spaces them);
    /// the last piece is held whole and a little more, so the formula never
    /// covers the text after it.
    static func layout(widths: [CGFloat], unit: CGFloat, joint jointWidth: CGFloat) -> Layout {
        let unit = unit.isFinite && unit > 0.5 ? unit : 4
        let jointWidth = jointWidth.isFinite && jointWidth >= 0 && jointWidth < 40 ? jointWidth : unit
        var text = anchor
        var starts: [Int] = []
        var counts: [Int] = []
        let pieces = widths.isEmpty ? [0] : widths
        var formula: CGFloat = 0
        var held: CGFloat = 0
        for (index, width) in pieces.enumerated() {
            let last = index == pieces.count - 1
            starts.append((text as NSString).length)
            formula += width.isFinite && width > 0 ? width : 0
            // a piece before another ends with the joint
            let wanted = (formula - held - (last ? 0 : jointWidth)) / unit
            // the last piece with a third of a filler to spare: a field may draw the fillers a
            // little narrower than they were measured
            let count = max(1, Int(last ? (wanted + 0.3).rounded(.up) : wanted.rounded()))
            counts.append(count)
            held += CGFloat(count) * unit + (last ? 0 : jointWidth)
            text += String(repeating: filler, count: count)
            if !last { text += joint }
        }
        return Layout(text: text, starts: starts, counts: counts)
    }

    /// A filler's (or with `of`, another string's) width in `font` (a field
    /// that tells its font; the blank comes from a fallback font).
    static func unit(in font: NSFont, of text: String = filler) -> CGFloat {
        (text as NSString).size(withAttributes: [.font: font]).width
    }

    /// A filler's width in the system font at `size` points (a field that
    /// tells only its line's height).
    static func unit(atSize size: CGFloat, of text: String = filler) -> CGFloat {
        unit(in: NSFont.systemFont(ofSize: max(6, min(size, 96))), of: text)
    }
}
