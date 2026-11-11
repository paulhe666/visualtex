import Foundation

/// A place in a field, in points from where the field's caret stood before the formula's space
/// went in: x to the right, y down (the bottoms of lines).
struct Place: Equatable {
    var x: CGFloat
    var y: CGFloat

    func near(_ other: Place, _ tolerance: CGFloat) -> Bool {
        abs(x - other.x) <= tolerance && abs(y - other.y) <= tolerance
    }
}

/// What a field answered for the characters of the marked text holding a formula's space.
struct FieldReads: Equatable {
    /// The marked text's length the field was asked about.
    var length: Int
    /// The places answered, by character index (`length`: after the last character).
    var at: [Int: Place]
}

/// Where a formula's space lies in a field: each piece's start, where the space starts (the
/// anchor before the first piece) and where it ends (where the field has its caret).
struct Arrangement: Equatable {
    var hair: Place
    var pieces: [Place]
    var end: Place

    /// The first piece of the last line.
    func lastLineStart(_ tolerance: CGFloat) -> Int {
        guard var index = pieces.indices.last else { return 0 }
        while index > 0, abs(pieces[index - 1].y - pieces[index].y) <= tolerance { index -= 1 }
        return index
    }
}

/// The hair's, a filler's and a joint's widths as a field draws them: from what was measured there
/// (spans of so many fillers, joints and hairs so many points wide), starting from an estimate.
/// A least-squares fit held towards the estimate: one span corrects the filler, spans of different
/// make-up tell the joint apart, and nothing measured leaves the estimate.
struct WidthFit {
    struct Widths: Equatable {
        var hair: CGFloat
        var filler: CGFloat
        var joint: CGFloat
    }

    private struct Span {
        var fillers: CGFloat
        var joints: CGFloat
        var hairs: CGFloat
        var width: CGFloat
    }

    let estimate: Widths
    /// How firmly the estimate is held (widths learnt in this application before: firmly).
    let trust: CGFloat
    private var spans: [Span] = []
    private(set) var widths: Widths
    /// The most fillers a span was measured over (0: nothing measured).
    private(set) var longest = 0
    /// Every span measured so far was a whole number of points wide (a field that rounds its places).
    private var wholePoints = true
    /// A span with a joint in it was measured.
    var jointMeasured: Bool { spans.contains { $0.joints > 0 } }

    init(estimate: Widths, trust: CGFloat = 1) {
        self.estimate = estimate
        self.trust = trust
        widths = estimate
    }

    /// A span of the held space measured in the field. False when it changes nothing.
    @discardableResult
    mutating func measured(fillers: Int, joints: Int, hairs: Int, width: CGFloat) -> Bool {
        guard fillers + joints + hairs > 0, fillers >= 0, width.isFinite, width >= 0 else { return false }
        // (a width far from what such a span can be is not this span's; a hair alone may be anything
        // up to a filler: a field that sets its text in cells draws it a cell wide)
        let expected = CGFloat(fillers) * estimate.filler + CGFloat(joints) * estimate.joint + CGFloat(hairs) * estimate.hair
        if fillers > 0 { guard width > expected * 0.5, width < expected * 2 else { return false } }
        let span = Span(fillers: CGFloat(fillers), joints: CGFloat(joints), hairs: CGFloat(hairs), width: width)
        if let index = spans.firstIndex(where: { $0.fillers == span.fillers && $0.joints == span.joints && $0.hairs == span.hairs }) {
            if abs(spans[index].width - width) < 0.05 { return false }
            spans[index] = span
        } else {
            spans.append(span)
            if spans.count > 48 { spans.removeFirst() }
        }
        longest = max(longest, fillers)
        let before = widths
        // The widths stand while they account for every span to within what a field's whole points
        // allow (a point): fitted anew at every measure, they moved by hundredths, the space was held
        // a filler longer or shorter for it at every key, and a long formula's lines went back and
        // forth (Chrome, 2026-11-11). The hair follows its own measure.
        let hair = spans.last(where: { $0.fillers == 0 && $0.joints == 0 })?.width ?? widths.hair
        // (a field that gives places to fractions of a point is taken at its word: a third of a point)
        if abs(width - width.rounded()) > 0.01 { wholePoints = false }
        let allowed: CGFloat = wholePoints ? 1.05 : 0.3
        let contradicted = spans.contains {
            ($0.fillers > 0 || $0.joints > 0) && abs($0.fillers * widths.filler + $0.joints * widths.joint + $0.hairs * hair - $0.width) > allowed
        }
        if contradicted {
            solve()
        } else if abs(hair - widths.hair) > 0.05 {
            widths.hair = min(max(hair, 0), widths.filler * 1.3)
        }
        return before != widths
    }

    private mutating func solve() {
        // The hair: as measured by itself (from its place to the next character's), else the
        // estimate. It is not fitted with the others: a fraction of a point going between it and the
        // fillers at every measure moved the formula's start at every key.
        let hair = spans.last(where: { $0.fillers == 0 && $0.joints == 0 })?.width ?? estimate.hair
        // The filler f and the joint j: minimise sum (a f + b j - (d - c hair))^2 + lf (f - f0)^2 + lj (j - j0)^2
        let lf = 0.02 * trust, lj = 0.6 * trust
        var m00 = lf, m01: CGFloat = 0, m11 = lj
        var v0 = lf * estimate.filler, v1 = lj * estimate.joint
        for span in spans where span.fillers > 0 || span.joints > 0 {
            let width = span.width - span.hairs * hair
            m00 += span.fillers * span.fillers
            m01 += span.fillers * span.joints
            m11 += span.joints * span.joints
            v0 += span.fillers * width
            v1 += span.joints * width
        }
        let d = m00 * m11 - m01 * m01
        guard abs(d) > 1e-9 else { return }
        let filler = (v0 * m11 - v1 * m01) / d
        let joint = (m00 * v1 - m01 * v0) / d
        guard filler.isFinite, joint.isFinite, filler > estimate.filler * 0.5, filler < estimate.filler * 2 else { return }
        widths = Widths(
            hair: min(max(hair, 0), filler * 1.3),
            filler: filler,
            joint: min(max(joint, estimate.joint * 0.4), max(estimate.joint * 2.5, filler * 1.5))
        )
    }
}

/// Where the formula's space lies in the field. The space is VisualTeX's own text, so its shape is
/// known: so many fillers and joints of known widths from where the field's caret stood, broken
/// into lines where a line is full. The model says where each piece is at once when a space is set
/// (`hold`); what the field answers afterwards (`absorb`) confirms that or corrects what the model
/// did not know: the widths the field draws the blanks with, where its lines begin and end, that it
/// moved its text. The formula is never drawn from an answer the model cannot account for.
///
/// What fields answer for a character of the marked text, measured with HostProbe (2026-11-11):
/// - TextKit, WebKit, Chrome: each character's own place, within a few milliseconds.
/// - WeChat: where each character ENDS (its right edge, where the others answer the left one), for
///   the characters it had before; where the last character ends for all the others (for all of
///   them once the text was shortened).
/// - VS Code: the places of the text set before (and where that ended, for characters added since).
/// - Chrome right after a composition was typed in: places of that composition, for a long while.
/// - Asked for a place at or past the marked text's end, TextKit and WeChat answer nothing and
///   then answer wrongly for other characters: nothing past the last character is ever asked.
/// So two things are taken from the answers: pieces whose characters stand a filler apart (they
/// are where the field has them), and the last character's place (the space's tail: true in every
/// field above but VS Code, where it is the old end and tells nothing).
///
/// (Before this the formula was drawn where the field answered for the first piece's character,
/// and the end of the marked text was asked for at every poll: in WeChat and QQ the formula stood
/// after its own space, with the field's underline showing before it.)
struct SpaceModel {
    typealias Widths = WidthFit.Widths

    /// The height of the field's line (its caret's rectangle).
    let lineHeight: CGFloat
    private(set) var fit: WidthFit
    var widths: Widths { fit.widths }
    private(set) var layout = FormulaSpace.Layout(text: "", starts: [], counts: [])
    private(set) var arrangement = Arrangement(hair: Place(x: 0, y: 0), pieces: [], end: Place(x: 0, y: 0))
    /// The field's answers agree with the arrangement of the space held now.
    private(set) var confirmed = false
    /// Where the space starts (0, 0: where the caret stood).
    private(set) var origin = Place(x: 0, y: 0)
    /// What placed the origin last: the caret before the space, the field's answers for the
    /// characters, or its caret after the space.
    private(set) var originFrom = OriginSource.caret
    /// Where the lines below the first begin, once known, and the distance between lines.
    private(set) var left: CGFloat?
    private(set) var pitch: CGFloat?
    /// How far below the first line each further line is, as the field told.
    private var downs: [CGFloat] = []
    /// The lines' right edge: at least `rightLow` (a line's pieces reached it), less than
    /// `rightHigh` (a piece that would have ended there went to the next line).
    private(set) var rightLow = -CGFloat.infinity
    private(set) var rightHigh = CGFloat.infinity
    /// Where the space ended and where its last character was before this space was set (answered
    /// again, they tell nothing of this space), and the last character's place as last answered.
    private var endAtHold = Place(x: 0, y: 0)
    private var tailAtHold: Place?
    private var appendedOne = false
    /// What the field answered last, and last before this space was set.
    private var lastReads: FieldReads?
    private var readsAtHold: FieldReads?
    private var textAtHold = ""
    /// How many characters this space's text and the one before it begin alike.
    private var unchanged = 0
    /// The field answers the places of the text set before (it answered where that ended for
    /// characters added since): its answers are true only as far as the two texts begin alike.
    private(set) var late = false
    private var lastTail: Place?
    /// What the field's answers gave last: pieces' places, or only the last character's (for the log).
    private(set) var answers = Answers.unknown
    /// The last space whose tail the field told on one line: its fillers and joints, and the tail.
    private var lastMeasured: (fillers: Int, joints: Int, tail: Place)?

    /// Which edge of a character the field answers: its left one, or (WeChat) its right one. Told
    /// by the first answer for a space's last character: where the space ends, or a filler before.
    private(set) var edges: Edges?
    enum Edges { case left, right }
    /// (0 for a field that answers right edges, 1 for left: how many fillers the last character's
    /// answer is before the space's end)
    private var back: CGFloat { edges == .right ? 0 : 1 }

    enum OriginSource { case caret, characters, end }
    enum Answers: String { case unknown, characters, tail }

    init(lineHeight: CGFloat, estimate: Widths, trust: CGFloat = 1) {
        self.lineHeight = max(lineHeight, 6)
        fit = WidthFit(estimate: estimate, trust: trust)
    }

    private var hairCount: Int { layout.starts.first ?? 0 }
    private var lineTolerance: CGFloat { (pitch ?? lineHeight) * 0.4 }

    // MARK: A space is set

    /// The field is given `layout` to hold: where its pieces are, as far as the model knows.
    mutating func hold(_ layout: FormulaSpace.Layout) {
        // (where the field had the space end: a filler after the last character as it answered, else
        // as the model had it; where the caret stood before the first space)
        endAtHold = lastTail.map { Place(x: $0.x + widths.filler * back, y: $0.y) } ?? (self.layout.counts.isEmpty ? Place(x: 0, y: 0) : arrangement.end)
        tailAtHold = lastTail
        lastTail = nil
        readsAtHold = lastReads
        textAtHold = self.layout.text
        appendedOne = !self.layout.text.isEmpty && layout.text == self.layout.text + FormulaSpace.filler
        unchanged = zip(Array(self.layout.text.utf16), Array(layout.text.utf16)).prefix(while: { $0 == $1 }).count
        self.layout = layout
        confirmed = false
        arrangement = arranged()
    }

    /// The arrangement for the lines as far as they are known: a piece that fits for sure stays, one
    /// that does not for sure goes down, one in between does what it did (as near as that tells).
    private func arranged() -> Arrangement {
        let right: CGFloat
        if rightHigh.isFinite {
            right = rightLow.isFinite ? (max(rightLow, rightHigh - widths.filler) + rightHigh) / 2 : rightHigh - widths.filler / 2
        } else {
            right = .infinity
        }
        return Self.arrange(layout, widths, origin: origin, left: left ?? min(origin.x, 0), right: right, pitch: pitch ?? lineHeight * 1.25, downs: downs)
    }

    /// The space set out as a field sets its text: the pieces one after another, a piece that does
    /// not fit before `right` on the line below, from `left` (the joint before it stays behind).
    /// `downs`: how far below the first line each further line is, as far as the field told (a field
    /// that gives places in whole points puts its lines a point off even steps).
    static func arrange(_ layout: FormulaSpace.Layout, _ widths: Widths, origin: Place, left: CGFloat, right: CGFloat, pitch: CGFloat, downs: [CGFloat] = []) -> Arrangement {
        var x = origin.x + ((layout.starts.first ?? 0) > 0 ? widths.hair : 0)
        var y = origin.y
        var line = 0
        var pieces: [Place] = []
        for (index, count) in layout.counts.enumerated() {
            let width = CGFloat(count) * widths.filler
            // (a piece at a line's start stays there however wide)
            if x + width > right + 0.01, x > left + 0.5 {
                line += 1
                y = origin.y + (line <= downs.count ? downs[line - 1] : (downs.last ?? 0) + pitch * CGFloat(line - downs.count))
                x = left
            }
            pieces.append(Place(x: x, y: y))
            x += width + (index < layout.counts.count - 1 ? widths.joint : 0)
        }
        return Arrangement(hair: origin, pieces: pieces, end: Place(x: x, y: y))
    }

    // MARK: What the field answers

    /// What the field's answers say, read as answers about the space held now.
    struct Evidence {
        /// The last character's place.
        var tail: Place?
        /// The pieces whose characters stand a filler apart: their starts.
        var pieces: [Int: Place] = [:]
        /// Runs of fillers measured inside pieces: (fillers, points).
        var runs: [(Int, CGFloat)] = []
        /// The hair's width: from its place to the first piece's.
        var hair: CGFloat?
    }

    func evidence(_ reads: FieldReads) -> Evidence {
        var out = Evidence()
        let length = (layout.text as NSString).length
        guard reads.length == length, length > 0 else { return out }
        let tolerance = lineTolerance
        // (a late field's answers are of the text before: true as far as the two begin alike, and for
        // one filler added at the end)
        let trusted = late ? (appendedOne ? length : unchanged) : length
        if length - 1 < trusted { out.tail = reads.at[length - 1] }
        // Each piece's start, from its first few characters and its last: the first's own place when
        // the others go on from it a filler apart each; without the first, back from the others by
        // their own advance. Characters that are not a filler apart are not where the field has
        // this space's characters (places of an earlier text, or one place for several characters).
        for (piece, start) in layout.starts.enumerated() {
            let count = layout.counts[piece]
            var found: [(offset: Int, place: Place)] = []
            for offset in Set([0, 1, 2, count - 1]).sorted() where offset >= 0 && offset < count {
                guard let place = reads.at[start + offset] else { continue }
                found.append((offset, place))
            }
            // (a piece one filler long has nothing to show that its character's place is its own: it
            // goes by the pieces beside it, and the last one by the space's tail)
            guard count >= 2, start + count <= trusted, found.count >= 2, let first = found.first, let far = found.last else { continue }
            var sound = found.allSatisfy { abs($0.place.y - first.place.y) <= tolerance }
            var advances: [CGFloat] = []
            for index in 1..<found.count where sound {
                let advance = (found[index].place.x - found[index - 1].place.x) / CGFloat(found[index].offset - found[index - 1].offset)
                if advance < widths.filler * 0.5 || advance > widths.filler * 1.7 { sound = false }
                advances.append(advance)
            }
            if sound, let least = advances.min(), let most = advances.max(), most > least * 1.3 + 0.5 { sound = false }
            guard sound else { continue }
            let span = far.offset - first.offset
            let advance = (far.place.x - first.place.x) / CGFloat(span)
            // (a right edge is a character further along than the left one)
            out.pieces[piece] = Place(x: first.place.x - (CGFloat(first.offset) + 1 - back) * advance, y: first.place.y)
            out.runs.append((span, far.place.x - first.place.x))
        }
        // The pieces in order: each after the one before on its line, as far as its fillers and the
        // joint go (closely, once the widths are measured), or on a lower line. Answers that are not
        // are not about this space.
        let known = out.pieces.keys.sorted()
        let slack: CGFloat = fit.longest >= 4 ? 0.06 : 0.3
        for (order, piece) in known.enumerated() where order > 0 {
            let before = known[order - 1]
            let a = out.pieces[before]!, b = out.pieces[piece]!
            var wrong = false
            if abs(a.y - b.y) <= tolerance {
                var expected: CGFloat = 0
                for index in before..<piece { expected += CGFloat(layout.counts[index]) * widths.filler + widths.joint }
                wrong = abs(b.x - a.x - expected) > max(fit.longest >= 4 ? 2 : 3, expected * slack)
            } else {
                wrong = b.y < a.y
            }
            if wrong {
                out.pieces = [:]
                break
            }
        }
        if edges != .right, hairCount > 0, let hair = reads.at[0], let first = out.pieces[0], abs(hair.y - first.y) <= tolerance {
            let width = first.x - hair.x
            if width > -0.3, width < widths.filler * 1.3 { out.hair = max(0, width) }
        }
        return out
    }

    /// Takes in what the field answers (answers that came twice the same), `sinceHold` seconds after
    /// the space was set. True when the arrangement, the origin or the widths changed.
    @discardableResult
    mutating func absorb(_ reads: FieldReads, sinceHold: TimeInterval = 1) -> Bool {
        let beforeArrangement = arrangement
        let beforeWidths = widths
        noteLate(reads)
        lastReads = reads
        let tolerance = lineTolerance
        // Which edge of a character the field answers. The hair is the space's first character: a
        // field that answers left edges has it where the caret stood, and the character after it the
        // hair's width on (measured there: a field that sets its text in cells draws the hair a cell
        // wide, VS Code). One that answers right edges (WeChat) has neither there, and its answer for
        // the last character of a space on one line is where the space ends.
        if reads.length == (layout.text as NSString).length, hairCount > 0 {
            let first = layout.starts[0]
            if edges != .right, let hair = reads.at[0], abs(hair.y - origin.y) <= tolerance, abs(hair.x - origin.x) <= 1.5 {
                edges = .left
                if let next = reads.at[first], abs(next.y - hair.y) <= tolerance, next.x - hair.x > -0.3, next.x - hair.x < widths.filler * 1.6 {
                    fit.measured(fillers: 0, joints: 0, hairs: 1, width: max(0, next.x - hair.x))
                }
            } else if edges == nil, let hair = reads.at[0], let mark = reads.at[reads.length - 1],
                      abs(mark.y - origin.y) <= tolerance, arrangement.pieces.allSatisfy({ abs($0.y - origin.y) <= tolerance }),
                      hair.x - origin.x >= widths.filler * 0.45, abs(mark.x - arrangement.end.x) <= widths.filler * 0.45 {
                edges = .right
            }
        }
        if beforeWidths != widths { arrangement = arranged() }
        let evidence = self.evidence(reads)
        var tail = evidence.tail
        if let place = tail {
            // (where the model has what the field answered for: the last character's edge)
            let predicted = Place(x: arrangement.end.x - widths.filler * back, y: arrangement.end.y)
            let fillers = layout.counts.reduce(0, +)
            let least = edges == .right ? widths.filler * 0.5 : (fillers >= 2 ? widths.filler * 0.5 : -widths.filler * 0.5 - 1)
            if abs(place.y - origin.y) <= tolerance, place.x < origin.x + least {
                // On the line the space starts on, left of where its first filler ends (of where it
                // starts, for the left edge of a space's only filler): not the last character of this
                // space (a field that answers where the marked text starts for every character).
                tail = nil
            } else if edges != .right, place.near(endAtHold, 0.5), !predicted.near(endAtHold, 1),
                      reads.length - (textAtHold as NSString).length >= 2 || textAtHold.isEmpty,
                      reads.length < 2 || reads.at[reads.length - 2].map({ $0.near(place, 0.3) }) ?? true {
                // Where the space ended before this one was set, where this one's last character is
                // not, and the character before it answered there too: the field answers the old end
                // for the characters added since (VS Code), or the caret where it stood before any
                // space. No news of this space. (One character more than the text had stands where
                // that ended, or as near as a joint and a filler differ.)
                tail = nil
                if !textAtHold.isEmpty, unchanged >= 2 { late = true }
            } else if let before = tailAtHold, place.near(before, 0.3), !predicted.near(before, 1), late || sinceHold < 0.12 {
                // The last character where the one before this space was, where this one's is not: the
                // field has not laid this space out yet (believed once it has kept to it for a moment:
                // the lines may be broken otherwise than the model has them), or it answers for the
                // text before and keeps to that (a late field). No news of this space.
                tail = nil
            }
        }
        if let place = evidence.tail { lastTail = place }

        if !evidence.pieces.isEmpty {
            answers = .characters
            adopt(evidence)
        } else if tail != nil {
            answers = .tail
        }
        if let tail {
            settle(tail, piecesTold: evidence.pieces[layout.counts.count - 1] != nil)
        }
        return arrangement != beforeArrangement || widths != beforeWidths
    }

    /// A field that answers for the text set before: characters after the first one that changed,
    /// which the change put further along or less far in the text by most of a filler or more,
    /// answered exactly as before this space was set (two of them at least, at places of their own:
    /// in a field that answers for the text as it is, they have moved).
    private mutating func noteLate(_ reads: FieldReads) {
        let new = Array(layout.text.utf16), old = Array(textAtHold.utf16)
        guard !late, let before = readsAtHold, reads.length == new.count, before.length == old.count else { return }
        let filler = FormulaSpace.filler.utf16.first, joint = FormulaSpace.joint.utf16.first
        func width(_ unit: UInt16) -> CGFloat { unit == filler ? widths.filler : (unit == joint ? widths.joint : widths.hair) }
        var along: CGFloat = 0
        var places: [Place] = []
        for index in unchanged..<min(new.count, old.count) {
            if abs(along) >= widths.filler * 0.8, let now = reads.at[index], let was = before.at[index], now.near(was, 0.05),
               !places.contains(where: { $0.near(now, 0.3) }) {
                places.append(now)
            }
            along += width(new[index]) - width(old[index])
        }
        if places.count >= 2 { late = true }
    }

    private mutating func move(to start: Place, from source: OriginSource) {
        let dx = start.x - origin.x
        origin = start
        originFrom = source
        if let known = left { left = known + dx }
        if rightLow.isFinite { rightLow += dx }
        if rightHigh.isFinite { rightHigh += dx }
        arrangement = arranged()
    }

    /// The pieces where the field's answers put them: its lines are the formula's lines.
    private mutating func adopt(_ evidence: Evidence) {
        let tolerance = lineTolerance
        let count = layout.counts.count
        // the widths, from pieces after one another on a line and from runs inside pieces
        for (fillers, width) in evidence.runs { fit.measured(fillers: fillers, joints: 0, hairs: 0, width: width) }
        if let hair = evidence.hair { fit.measured(fillers: 0, joints: 0, hairs: 1, width: hair) }
        // (and from a line's first piece the field told of to its last: many fillers and joints, a
        // point's rounding at most between them)
        var from: Int?
        var previous: Int?
        for piece in evidence.pieces.keys.sorted() + [-1] {
            if let start = from, let last = previous, piece < 0 || abs(evidence.pieces[piece]!.y - evidence.pieces[start]!.y) > tolerance {
                if last - start >= 2 {
                    fit.measured(fillers: layout.counts[start..<last].reduce(0, +), joints: last - start, hairs: 0, width: evidence.pieces[last]!.x - evidence.pieces[start]!.x)
                }
                from = piece < 0 ? nil : piece
            } else if from == nil, piece >= 0 {
                from = piece
            }
            previous = piece
        }
        let known = evidence.pieces.keys.sorted()
        for (order, piece) in known.enumerated() where order > 0 && known[order - 1] == piece - 1 {
            let a = evidence.pieces[piece - 1]!, b = evidence.pieces[piece]!
            guard abs(a.y - b.y) <= tolerance else { continue }
            fit.measured(fillers: layout.counts[piece - 1], joints: 1, hairs: 0, width: b.x - a.x)
        }
        // where the space starts: the first piece's start less the hair, when it is on the line the
        // space starts on, or straight below or above where it was (the field scrolled)
        // (the first piece the field told of, less what lies before it on the line)
        if let earliest = known.first, let told = evidence.pieces[earliest] {
            var before = hairCount > 0 ? widths.hair : 0
            for index in 0..<earliest { before += CGFloat(layout.counts[index]) * widths.filler + widths.joint }
            // on the first line: no piece between it and the start was told on another line, and it is
            // where the pieces before it leave room for it
            let start = Place(x: told.x - before, y: told.y)
            if abs(told.y - origin.y) <= tolerance || abs(start.x - origin.x) <= 1.5 {
                // (a field's places for characters and for its caret differ by a point or two: the caret's
                // stands unless the characters are clearly elsewhere)
                if !start.near(origin, earliest == 0 ? 2 : 3) { move(to: start, from: .characters) }
            }
        }
        // the lines: a piece the field told of is where it told; the others follow the one before
        var pieces: [Place] = []
        var x = origin.x + (hairCount > 0 ? widths.hair : 0)
        var y = origin.y
        let step = pitch ?? lineHeight * 1.25
        let leftEdge = left ?? min(origin.x, 0)
        for index in 0..<count {
            let width = CGFloat(layout.counts[index]) * widths.filler
            if let told = evidence.pieces[index] {
                // where the field told; on the line the piece before is on, and within a point or so of
                // where the widths put it, where the widths put it (a field's whole points would move
                // every piece a point back and forth between a key and the field's answer to it)
                if !(abs(told.y - y) <= 0.01 && abs(told.x - x) <= 1.5) {
                    x = told.x
                    y = told.y
                }
            } else if x + width > (rightHigh.isFinite ? rightHigh - 0.01 : .infinity), x > leftEdge + 0.5 {
                // (the next line as the field told it before, else a line further down)
                y = downs.map { origin.y + $0 }.first(where: { $0 > y + tolerance }) ?? y + step
                x = leftEdge
            }
            pieces.append(Place(x: x, y: y))
            x += width + (index < count - 1 ? widths.joint : 0)
        }
        var end = Place(x: x, y: y)
        // A piece the field did not tell of, before one it placed with room before it on its line:
        // the pieces are one text, so it stands right before that one.
        if count >= 2 {
            for index in stride(from: count - 2, through: 0, by: -1) where evidence.pieces[index] == nil {
                let next = pieces[index + 1]
                let width = CGFloat(layout.counts[index]) * widths.filler + widths.joint
                let onLowerLine = next.y - pieces[index].y > tolerance
                if onLowerLine, next.x - width >= leftEdge - widths.filler * 0.5 {
                    pieces[index] = Place(x: next.x - width, y: next.y)
                }
            }
            // (and the end after the last piece where that now is)
            if let lastPlace = pieces.last, evidence.pieces[count - 1] == nil, abs(lastPlace.y - end.y) > tolerance || abs(lastPlace.x + CGFloat(layout.counts[count - 1]) * widths.filler - end.x) > 0.01 {
                end = Place(x: lastPlace.x + CGFloat(layout.counts[count - 1]) * widths.filler, y: lastPlace.y)
            }
        }
        arrangement = Arrangement(hair: origin, pieces: pieces, end: end)
        learnLines(told: Set(known))
        confirmed = true
    }

    /// What the arrangement tells of the field's lines: where the lines below the first begin, how
    /// far apart they are, and between what the right edge lies. `told`: the pieces the field
    /// placed itself (nil: all of them, an arrangement its caret confirmed).
    private mutating func note(down: CGFloat, line: Int) {
        guard line >= 1, line <= downs.count + 1 else { return }
        if line == downs.count + 1 { downs.append(down) } else { downs[line - 1] = down }
    }

    private mutating func learnLines(told: Set<Int>?) {
        let tolerance = lineTolerance
        let pieces = arrangement.pieces
        guard !pieces.isEmpty else { return }
        var low = -CGFloat.infinity
        var high = CGFloat.infinity
        var lefts: [CGFloat] = []
        var steps: [CGFloat] = []
        // (How far a line's pieces reach is reckoned as arranged() reckons it: from the line's start by
        // the widths, not from each piece's own place as the field told it. A field's whole points
        // put a piece a point off that, and an edge learnt on one footing did not hold on the other:
        // the last piece of a full line went up and down at every key.)
        var along: CGFloat = 0
        for index in pieces.indices {
            let width = CGFloat(layout.counts[index]) * widths.filler
            let known = told?.contains(index) ?? true
            let begins = index == 0 || pieces[index].y - pieces[index - 1].y > tolerance
            let before = along
            if begins {
                along = index == 0 && pieces[0].y - origin.y <= tolerance ? origin.x + (hairCount > 0 ? widths.hair : 0) : (left ?? pieces[index].x)
            }
            let reached = along + width
            if index > 0, begins {
                // a line begins: the piece did not fit after the one before
                if known {
                    lefts.append(pieces[index].x)
                    steps.append(pieces[index].y - pieces[index - 1].y)
                }
                if known, told?.contains(index - 1) ?? true {
                    // (`before`: where this piece would have begun on the line above)
                    high = min(high, before + width)
                }
            } else if index == 0, pieces[0].y - origin.y > tolerance, known {
                // the formula itself went to the line below
                lefts.append(pieces[0].x)
                steps.append(pieces[0].y - origin.y)
                high = min(high, origin.x + (hairCount > 0 ? widths.hair : 0) + width)
            }
            if known { low = max(low, reached) }
            along = reached + widths.joint
        }
        if let first = lefts.min() { left = first }
        if let step = steps.min(), step > lineHeight * 0.5 { pitch = step }
        // how far down each line is, where the field placed the pieces itself
        if let told {
            var line = pieces[0].y - origin.y > tolerance ? 1 : 0
            if line == 1, told.contains(0) { note(down: pieces[0].y - origin.y, line: 1) }
            for index in pieces.indices.dropFirst() where pieces[index].y - pieces[index - 1].y > tolerance {
                line += 1
                if told.contains(index) { note(down: pieces[index].y - origin.y, line: line) }
            }
        }
        // (edges that contradict what was known: the field's box changed; the new ones stand)
        if low > rightHigh || high < rightLow {
            rightLow = low
            rightHigh = high
        } else {
            rightLow = max(rightLow, low)
            rightHigh = min(rightHigh, high)
        }
        if rightLow >= rightHigh { rightLow = rightHigh - 0.5 }
    }

    /// The field has the space's last character at `tail` (the space ends a filler after it): the
    /// arrangement ends there, or is put right. `piecesTold`: the field placed the last piece itself.
    private mutating func settle(_ tail: Place, piecesTold: Bool) {
        let end = Place(x: tail.x + widths.filler * back, y: tail.y)
        let tolerance = lineTolerance
        let tight = max(1.5, widths.filler * 0.2)
        let predicted = arrangement.end
        if !piecesTold { measureGrowth(tail) }
        if abs(end.y - predicted.y) <= tolerance, abs(end.x - predicted.x) <= tight {
            confirmed = true
            if !piecesTold { measureWhole(tail) }
            if !piecesTold {
                // a start worked out from the field's caret stays the space's length before it
                if originFrom == .end {
                    let length = arranged().end.x - origin.x
                    let single = arrangement.pieces.allSatisfy { abs($0.y - origin.y) <= tolerance }
                    if single { move(to: Place(x: end.x - length, y: origin.y), from: .end) }
                }
                learnLines(told: nil)
            }
            return
        }
        // the field placed the last piece itself: its last character adds nothing to that
        if piecesTold { return }
        // Lines broken otherwise; or that, with the field's text scrolled up a line as the space grew
        // onto a new one (a box as tall as its lines are many keeps the caret's line in view).
        // (What is known of the lines' right edge first, either way; then any edge: the box changed.)
        let step = pitch ?? lineHeight * 1.25
        for bounded in [true, false] {
            for scrolled in [CGFloat(0), step] {
                let kept = origin
                origin.y -= scrolled
                if let other = rearranged(endingAt: end, bounded: bounded) {
                    arrangement = other.arrangement
                    if let known = other.left { left = known }
                    if let known = other.pitch { pitch = known }
                    learnLines(told: nil)
                    confirmed = true
                    return
                }
                origin = kept
            }
            if !rightHigh.isFinite, !rightLow.isFinite { break }
        }
        let single = arrangement.pieces.allSatisfy { abs($0.y - origin.y) <= tolerance }
        if single, abs(end.y - origin.y) <= tolerance, originFrom != .end {
            // one line, and wider or narrower than expected by what the widths may be off: the widths
            let expected = predicted.x - origin.x
            let measured = end.x - origin.x
            if expected > 0, measured > expected * 0.6, measured < expected * 1.6, measureWhole(tail) {
                arrangement = arranged()
                confirmed = true
                return
            }
        }
        // The field's text moved, or the space does not start where the caret stood: the space ends
        // at the field's caret, so it starts its own length before it. (An end further down and
        // elsewhere on its line is lines broken in a way the model cannot account for: the formula
        // stays where it starts.)
        let dx = end.x - predicted.x, dy = end.y - predicted.y
        guard abs(dy) <= tolerance || abs(dx) <= tight || dy < 0 else {
            confirmed = false
            return
        }
        move(to: Place(x: origin.x + dx, y: origin.y + dy), from: .end)
        confirmed = true
    }

    /// What the space grew (or shrank) by on its line since the field last told its last character's
    /// place: so many fillers and joints, so many points. (It does not depend on where the space
    /// starts.)
    private mutating func measureGrowth(_ tail: Place) {
        let tolerance = lineTolerance
        let single = arrangement.pieces.allSatisfy { abs($0.y - origin.y) <= tolerance } && abs(tail.y - origin.y) <= tolerance
        let now = (fillers: layout.counts.reduce(0, +), joints: max(0, layout.counts.count - 1), tail: tail)
        defer { lastMeasured = single ? now : nil }
        guard single, let last = lastMeasured, abs(last.tail.y - tail.y) <= tolerance else { return }
        let fillers = now.fillers - last.fillers, joints = now.joints - last.joints
        if fillers > 0, joints >= 0 {
            fit.measured(fillers: fillers, joints: joints, hairs: 0, width: tail.x - last.tail.x)
        } else if fillers < 0, joints <= 0 {
            fit.measured(fillers: -fillers, joints: -joints, hairs: 0, width: last.tail.x - tail.x)
        }
    }

    /// The whole space on one line, from where it starts (as the caret before it or the field's
    /// answers told, not as worked out from this tail) to its last character: a span for the widths.
    @discardableResult
    private mutating func measureWhole(_ tail: Place) -> Bool {
        let tolerance = lineTolerance
        guard originFrom != .end, arrangement.pieces.allSatisfy({ abs($0.y - origin.y) <= tolerance }), abs(tail.y - origin.y) <= tolerance else { return false }
        return fit.measured(fillers: layout.counts.reduce(0, +) - Int(back), joints: max(0, layout.counts.count - 1), hairs: hairCount, width: tail.x - origin.x)
    }

    /// The arrangement, with the lines broken as some right edge would break them, that ends at
    /// `end` (the field's caret); with where the lower lines begin and how far apart the lines are
    /// when that is what tells them.
    private func rearranged(endingAt end: Place, bounded: Bool) -> (arrangement: Arrangement, left: CGFloat?, pitch: CGFloat?)? {
        let tolerance = lineTolerance
        let tight = max(1.5, widths.filler * 0.2)
        let step = pitch ?? lineHeight * 1.25
        let down = end.y - origin.y
        // how many lines below the space's first its end is
        let linesDown: Int
        if let pitch { linesDown = Int((down / pitch).rounded()) } else { linesDown = down <= lineHeight * 0.5 ? 0 : max(1, Int((down / (lineHeight * 1.25)).rounded())) }
        guard linesDown >= 0 else { return nil }
        let stepNow = linesDown > 0 && pitch == nil ? down / CGFloat(linesDown) : step
        guard linesDown == 0 || (stepNow > lineHeight * 0.6 && stepNow < lineHeight * 3) else { return nil }
        var best: (Arrangement, CGFloat?, CGFloat?, CGFloat)?
        do {
            let lefts: [CGFloat?] = left != nil ? [left] : [nil]
            for knownLeft in lefts {
                var right = bounded ? rightHigh - 0.01 : CGFloat.infinity
                let floor = bounded ? rightLow : -CGFloat.infinity
                var rounds = 0
                while rounds < 400 {
                    rounds += 1
                    var leftEdge = knownLeft ?? min(origin.x, 0)
                    var candidate = Self.arrange(layout, widths, origin: origin, left: leftEdge, right: right, pitch: stepNow)
                    if knownLeft == nil, linesDown > 0 {
                        // where the lower lines begin is not known: where the last line must begin for the
                        // space to end at the field's caret (a few rounds: the lines between depend on it)
                        for _ in 0..<4 {
                            let lastStart = candidate.lastLineStart(tolerance)
                            guard lastStart > 0 || candidate.pieces[0].y - origin.y > tolerance else { break }
                            let wanted = leftEdge + end.x - candidate.end.x
                            if abs(wanted - leftEdge) < 0.05 { break }
                            leftEdge = wanted
                            candidate = Self.arrange(layout, widths, origin: origin, left: leftEdge, right: right, pitch: stepNow)
                        }
                        // (the lower lines begin at or left of where the space starts)
                        if leftEdge > origin.x + 1 { candidate.end.x = .infinity }
                    }
                    let lines = Int(((candidate.end.y - origin.y) / stepNow).rounded())
                    let off = abs(candidate.end.x - end.x)
                    if lines == linesDown, off <= tight, off < (best?.3 ?? .infinity) {
                        let learntLeft: CGFloat? = knownLeft == nil && linesDown > 0 ? leftEdge : nil
                        best = (candidate, learntLeft, linesDown > 0 && pitch == nil ? stepNow : nil, off)
                    }
                    // the next arrangement as the edge moves left: the piece reaching furthest goes down
                    var furthest = -CGFloat.infinity
                    for index in candidate.pieces.indices {
                        // (a piece a line begins with stays whatever the edge)
                        let movable = index == 0
                            ? abs(candidate.pieces[0].y - origin.y) <= tolerance && origin.x > leftEdge + 0.5
                            : abs(candidate.pieces[index].y - candidate.pieces[index - 1].y) <= tolerance
                        if !movable { continue }
                        furthest = max(furthest, candidate.pieces[index].x + CGFloat(layout.counts[index]) * widths.filler)
                    }
                    guard furthest.isFinite, furthest - 0.02 >= floor, lines <= linesDown else { break }
                    right = furthest - 0.02
                }
            }
        }
        guard let best else { return nil }
        return (best.0, best.1, best.2)
    }
}
