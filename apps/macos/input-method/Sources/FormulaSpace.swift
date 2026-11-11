import Cocoa

/// The space a field holds in its text for a formula that VisualTeX draws
/// over it: the formula's marked text. VisualTeX cuts the formula where TeX
/// may break an in-line formula (after a relation or a binary operator) and
/// tells the width of each piece; each piece is held by blank characters as
/// wide as it, the pieces are joined by word spaces, where the field may
/// break the line (every text system breaks there). A joint counts in the
/// piece before it. So the field wraps the formula as it wraps its own text,
/// in its own box, and pushes the text after the caret along; the input
/// method works out where each piece went (SpaceModel) and VisualTeX draws
/// it there, over the field's background (what the field draws for a
/// composition, its underline and its caret, is covered).
///
/// The marked text starts with one hair space: it lets the field break the
/// line before the formula, and where the field has it is where the space
/// starts. (It is set selected. Fields show their caret at the space's end
/// all the same, measured 2026-11-11: VisualTeX covers it there.)
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
    /// (VisualTeX draws the pieces of one line as the formula spaces them).
    /// The last piece is held to its width at most (in whole fillers): the
    /// held space's end, where the field draws its caret and the end of its
    /// underline, stays under the formula, so VisualTeX's cover hides them as
    /// the formula grows, without waiting for the field to tell where its
    /// space ends (held a little wider, the field's caret and underline stood
    /// past the formula all along a held key, 2026-11-10). The formula may
    /// stand over less than a filler of text after it.
    /// `margin`: fillers the last piece is held short of its width besides, for widths not known well
    /// (the field may draw the fillers a little wider than they were measured). `lineStart`: the
    /// piece the field began the last line with: the field and VisualTeX both set it at the line's
    /// start, so the pieces from it are held from there (held from the formula's start, the space on a
    /// later line ended past the formula by what the lines before had been rounded, 2026-11-10).
    static func layout(widths: [CGFloat], unit: CGFloat, joint jointWidth: CGFloat, margin: CGFloat = 0, lineStart: Int = 0) -> Layout {
        let unit = unit.isFinite && unit > 0.5 ? unit : 4
        let jointWidth = jointWidth.isFinite && jointWidth >= 0 && jointWidth < 40 ? jointWidth : unit
        var pieces = (widths.isEmpty ? [0] : widths).map { $0.isFinite && $0 > 0 ? $0 : 0 }
        while true {
            var text = anchor
            var starts: [Int] = []
            var counts: [Int] = []
            var formula: CGFloat = 0
            var held: CGFloat = 0
            var short = false
            for (index, width) in pieces.enumerated() {
                let last = index == pieces.count - 1
                starts.append((text as NSString).length)
                if index == lineStart, index > 0 {
                    formula = 0
                    held = 0
                }
                formula += width
                // a piece before another ends with the joint
                let wanted = (formula - held - (last ? 0 : jointWidth)) / unit
                // (a hair under a whole filler is a whole one: rounding of the widths)
                let whole = Int(last ? (wanted - max(0, margin) + 0.02).rounded(.down) : wanted.rounded())
                // A last piece narrower than a filler (a piece just begun) would end the space past the
                // formula by its filler and the joint before it: it is held with the piece before it
                // (VisualTeX draws it after that one, the field has no break there until it is wider).
                // (not the piece a line begins with: held with the one before, it would go back to that line)
                if last, whole < 1, pieces.count > 1, index != lineStart { short = true }
                let count = max(1, whole)
                counts.append(count)
                held += CGFloat(count) * unit + (last ? 0 : jointWidth)
                text += String(repeating: filler, count: count)
                if !last { text += joint }
            }
            if !short { return Layout(text: text, starts: starts, counts: counts) }
            let tail = pieces.removeLast()
            pieces[pieces.count - 1] += tail
        }
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
