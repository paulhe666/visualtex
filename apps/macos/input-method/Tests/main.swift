// Tests of SpaceModel against fields that answer as real applications do (measured with HostProbe):
//   swiftc -O Sources/FormulaSpace.swift Sources/SpaceModel.swift Tests/main.swift -o /tmp/spacetests && /tmp/spacetests
// A simulated field sets the held space as a text system sets text (words of fillers, breaks at spaces) and
// answers in one of the ways applications answer; after each change the model's pieces must be where the field
// really has them.
import Foundation

/// A text box: where its lines begin and end, the widths it draws the blanks with, where its caret stood.
struct SimField {
    var left: CGFloat = 0
    var right: CGFloat = 400
    var pitch: CGFloat = 20
    var hair: CGFloat = 1.1
    var filler: CGFloat = 9.6
    var joint: CGFloat = 4.1
    /// The caret before the space, from the line's start.
    var caretX: CGFloat = 0
    /// Lines the box shows before it scrolls its text up (nil: it grows).
    var visibleLines: Int?
    var caretLine = 0

    /// Each character's place and the place after the last (index 0...length), from where the caret stood.
    func places(_ text: String) -> [Place] {
        var out: [Place] = []
        var x = left + caretX
        var line = 0
        let characters = Array(text.unicodeScalars)
        var index = 0
        while index < characters.count {
            let character = characters[index]
            if character == "\u{2800}" {
                var run = 0
                while index + run < characters.count, characters[index + run] == "\u{2800}" { run += 1 }
                if x + CGFloat(run) * filler > right + 0.001, x > left + 0.001 {
                    line += 1
                    x = left
                }
                for _ in 0..<run {
                    out.append(Place(x: x, y: CGFloat(line)))
                    x += filler
                }
                index += run
            } else {
                out.append(Place(x: x, y: CGFloat(line)))
                x += character == "\u{200A}" ? hair : joint
                index += 1
            }
        }
        out.append(Place(x: x, y: CGFloat(line)))
        // a box that scrolls keeps the caret's line in view
        var scrolled = 0
        if let visibleLines { scrolled = max(0, caretLine + line - (visibleLines - 1)) }
        return out.map { Place(x: $0.x - left - caretX, y: (CGFloat($0.y) - CGFloat(scrolled)) * pitch) }
    }
}

/// How a field answers for a character of the marked text (measured in the applications named, HostProbe).
enum Answering {
    /// each character's own place (TextKit, WebKit, Chrome)
    case characters
    /// where the character ends for the characters the text before had (while the text only grew), where the last
    /// character ends for all the others; for the first character, where the first space of all ended (WeChat)
    case wechat
    /// each character's place as of the text set before this one, and where that text ended for the characters
    /// added since (VS Code)
    case late
    /// each character's own place, but the place after the space for one of them
    case poisoned(Int)
    /// where the marked text starts, for every character
    case start
}

final class Session {
    var field: SimField
    let answering: Answering
    var model: SpaceModel
    var previousText = ""
    var firstEnd: Place?
    var failures = 0
    let name: String
    var layout = FormulaSpace.Layout(text: "", starts: [], counts: [])
    /// The field's places are this far right of where the model's caret stood (the caret the field told before
    /// the space was not where the space went).
    var startOff: CGFloat = 0

    init(_ name: String, field: SimField, answering: Answering, estimate: SpaceModel.Widths = .init(hair: 1.2, filler: 9.2, joint: 3.7)) {
        self.name = name
        self.field = field
        self.answering = answering
        model = SpaceModel(lineHeight: field.pitch * 0.85, estimate: estimate)
    }

    func reads(_ text: String) -> FieldReads {
        let truth = field.places(text).map { Place(x: $0.x + startOff, y: $0.y) }
        let length = truth.count - 1
        var at: [Int: Place] = [:]
        // (the place after the last character is never asked)
        switch answering {
        case .characters:
            for index in 0..<length { at[index] = truth[index] }
        case .wechat:
            // (right edges: where each character ends)
            let previous = (previousText as NSString).length
            let grew = !previousText.isEmpty && text.hasPrefix(previousText)
            for index in 0..<length { at[index] = grew && index < previous ? truth[index + 1] : truth[length] }
            if firstEnd == nil { firstEnd = truth[length] }
            at[0] = firstEnd
        case .late:
            let old = previousText.isEmpty ? truth : field.places(previousText).map { Place(x: $0.x + startOff, y: $0.y) }
            for index in 0..<length { at[index] = index < old.count - 1 ? old[index] : old[old.count - 1] }
        case .poisoned(let index):
            for item in 0..<length { at[item] = truth[item] }
            // (never the last character: no application measured answers another place for that one)
            if index < length - 1 { at[index] = truth[length] }
        case .start:
            for index in 0..<length { at[index] = truth[0] }
        }
        return FieldReads(length: length, at: at)
    }

    /// The formula's pieces are now `widths` points wide: the space is held and the field answers.
    /// `settles`: how many times the field's answers are taken in (a late field needs the text set again).
    func type(_ widths: [CGFloat], expectAtOnce: Bool = false, tolerance: CGFloat = 1.0, check: Bool = true, note: String = "") {
        var rounds = 0
        while rounds < 4 {
            rounds += 1
            let lineStart = model.arrangement.lastLineStart(model.lineHeight * 0.4)
            let next = FormulaSpace.layout(widths: widths, unit: model.widths.filler, joint: model.widths.joint, margin: 0.15, lineStart: lineStart < widths.count ? lineStart : 0)
            if next == layout, rounds > 1 { break }
            layout = next
            model.hold(layout)
            if expectAtOnce, rounds == 1 { compare("\(note) at once", tolerance) }
            let answered = reads(layout.text)
            previousText = layout.text
            let before = model.widths
            model.absorb(answered)
            if model.widths == before { break }
        }
        if check { compare(note, tolerance) }
    }

    /// The field answers again for the space as it is (a late field has caught up).
    func settle(tolerance: CGFloat = 1.0, note: String = "") {
        model.absorb(reads(layout.text))
        compare(note + " settled", tolerance)
    }

    func compare(_ note: String, _ tolerance: CGFloat) {
        let truth = field.places(layout.text).map { Place(x: $0.x + startOff, y: $0.y) }
        var worst: CGFloat = 0
        var where_ = ""
        for (piece, start) in layout.starts.enumerated() {
            let have = model.arrangement.pieces[piece], want = truth[start]
            let off = max(abs(have.x - want.x), abs(have.y - want.y))
            if off > worst { worst = off; where_ = "piece \(piece) at \(fmt(have)) wanted \(fmt(want))" }
        }
        let end = model.arrangement.end, wantEnd = truth[truth.count - 1]
        let endOff = max(abs(end.x - wantEnd.x), abs(end.y - wantEnd.y))
        // (the end may be off by what the widths are still off over the last piece)
        if endOff > worst, endOff > tolerance * 2 { worst = endOff; where_ = "end at \(fmt(end)) wanted \(fmt(wantEnd))" }
        if worst > tolerance {
            failures += 1
            print("  FAIL \(name) [\(note)] pieces=\(layout.counts) \(where_) answers=\(model.answers.rawValue) widths=\(fmt(model.widths.filler))/\(fmt(model.widths.joint)) left=\(model.left.map(fmt) ?? "-") pitch=\(model.pitch.map(fmt) ?? "-") right=\(fmt(model.rightLow))..\(fmt(model.rightHigh))")
        }
    }

    func fmt(_ value: CGFloat) -> String { value.isFinite ? String(format: "%.1f", Double(value)) : "inf" }
    func fmt(_ place: Place) -> String { "(\(fmt(place.x)),\(fmt(place.y)))" }
}

/// A formula typed at its end: pieces as VisualTeX cuts them (a piece grows, then another begins).
func typing(_ pieces: Int, piece: CGFloat = 26, steps: Int = 3) -> [[CGFloat]] {
    var out: [[CGFloat]] = []
    var widths: [CGFloat] = []
    for _ in 0..<pieces {
        widths.append(0)
        for step in 1...steps {
            widths[widths.count - 1] = piece * CGFloat(step) / CGFloat(steps)
            out.append(widths)
        }
    }
    return out
}

var total = 0
func run(_ session: Session, _ body: (Session) -> Void) {
    body(session)
    total += session.failures
    print("\(session.failures == 0 ? "ok  " : "FAIL") \(session.name)")
}

let answerings: [(String, Answering)] = [("characters", .characters), ("wechat", .wechat), ("poisoned", .poisoned(1)), ("start", .start)]

// 1. One line, typed at the end: right from the first key, in every kind of field.
for (label, answering) in answerings {
    run(Session("one line, \(label)", field: SimField(right: 2000, caretX: 60), answering: answering)) { s in
        for widths in typing(6) { s.type(widths, tolerance: label == "start" ? 12 : 1.5) }
    }
}

// 2. The estimate a fifth off (Safari draws the blank narrower than its font says): no jump, the widths are learnt.
for (label, answering) in answerings.prefix(3) {
    run(Session("estimate 20% off, \(label)", field: SimField(right: 2000, filler: 10.3, caretX: 60), answering: answering, estimate: .init(hair: 1.2, filler: 12.6, joint: 4))) { s in
        for widths in typing(6) { s.type(widths, tolerance: 2.5) }
        s.type([26, 26, 26, 26, 26, 26, 12], tolerance: 1.2, note: "after learning")
    }
}

// 3. Wrapping while typing at the end, the formula after some text on the line.
for (label, answering) in answerings.prefix(3) {
    run(Session("wraps, \(label)", field: SimField(right: 300, caretX: 120), answering: answering)) { s in
        for widths in typing(14) { s.type(widths, tolerance: 1.5) }
    }
}

// 4. Once the field's lines are known, the next lines are right at once (before the field answers).
run(Session("later lines at once, wechat", field: SimField(right: 300, caretX: 120), answering: .wechat)) { s in
    for (index, widths) in typing(16).enumerated() { s.type(widths, expectAtOnce: index > 30, tolerance: 2.0) }
}

// 5. A field that answers one set late (Chromium): one line is right all along; a wrap is learnt a key late.
run(Session("late, one line", field: SimField(right: 2000, caretX: 60), answering: .late)) { s in
    for widths in typing(6) { s.type(widths, tolerance: 2.0) }
}
run(Session("late, wraps", field: SimField(right: 300, caretX: 120), answering: .late)) { s in
    for widths in typing(12) { s.type(widths, check: false) }
    s.settle(tolerance: 1.5)
}

// 6. Deleting from the end, back onto one line.
for (label, answering) in answerings.prefix(3) {
    run(Session("delete, \(label)", field: SimField(right: 300, caretX: 120), answering: answering)) { s in
        let grown = typing(10)
        for widths in grown { s.type(widths, tolerance: 1.5) }
        for widths in grown.reversed() { s.type(widths, tolerance: 1.5, note: "deleting") }
    }
}

// 7. The caret was not where the field said before the space (a place of the text before): the field's caret after
// the space puts the formula right.
// (not WeChat's kind: which edge of a character it answers is told from where the caret stood)
for (label, answering) in answerings.prefix(3) where label != "wechat" {
    run(Session("stale start, \(label)", field: SimField(right: 2000, caretX: 60), answering: answering)) { s in
        s.startOff = 45
        for (index, widths) in typing(4).enumerated() { s.type(widths, tolerance: 2.5, check: index >= 3) }
    }
}

// 8. The formula at the start of a message in a chat box that wraps (WeChat).
run(Session("chat box from the line's start, wechat", field: SimField(right: 260, caretX: 0), answering: .wechat)) { s in
    for widths in typing(18) { s.type(widths, tolerance: 2.0) }
}

// 9. A box two lines tall that scrolls as the space grows (the caret's line stays in view).
run(Session("scrolling box, wechat", field: SimField(right: 260, caretX: 40, visibleLines: 2), answering: .wechat)) { s in
    for widths in typing(16) { s.type(widths, tolerance: 2.0) }
}
run(Session("scrolling box, characters", field: SimField(right: 260, caretX: 40, visibleLines: 2), answering: .characters)) { s in
    for widths in typing(16) { s.type(widths, tolerance: 1.5) }
}

// 10. Edited in the middle (an early piece grows, the later ones move along and over the edge).
for (label, answering) in answerings.prefix(3) {
    run(Session("edit in the middle, \(label)", field: SimField(right: 300, caretX: 100), answering: answering)) { s in
        for widths in typing(5) { s.type(widths, tolerance: 1.5) }
        var widths: [CGFloat] = [26, 26, 26, 26, 26]
        for _ in 0..<8 {
            widths[1] += 9
            s.type(widths, tolerance: 2.0, note: "growing piece 1")
        }
    }
}

// 10b. A field that sets its text in cells (VS Code): the hair is a cell wide, a joint a filler wide; its answers are
// of the text before.
for (label, answering) in [("late", Answering.late), ("characters", Answering.characters)] {
    run(Session("cells, \(label)", field: SimField(right: 100000, hair: 7.2, filler: 8, joint: 8, caretX: 40), answering: answering, estimate: .init(hair: 1.1, filler: 9.2, joint: 3.7))) { s in
        for (index, widths) in typing(8).enumerated() {
            s.type(widths, tolerance: 2.5, check: index >= 2)
            if CommandLine.arguments.count > 2, CommandLine.arguments[2] == "cells", index < 6 {
                print("  cells \(label) step \(index) counts=\(s.layout.counts) widths=\(s.fmt(s.model.widths.hair))/\(s.fmt(s.model.widths.filler))/\(s.fmt(s.model.widths.joint)) edges=\(String(describing: s.model.edges)) late=\(s.model.late) answers=\(s.model.answers.rawValue) model=\(s.model.arrangement.pieces.map(s.fmt).joined(separator: " ")) end=\(s.fmt(s.model.arrangement.end))")
            }
        }
    }
}

// 11. Random boxes, widths and edits (typing at the end, deleting, an earlier piece growing or shrinking), each kind
// of field: after every change the pieces are where the field has them.
struct Random {
    var state: UInt64
    mutating func next() -> UInt64 {
        state = state &* 6364136223846793005 &+ 1442695040888963407
        return state >> 33
    }
    mutating func real(_ low: CGFloat, _ high: CGFloat) -> CGFloat { low + (high - low) * CGFloat(next() % 10000) / 10000 }
    mutating func int(_ count: Int) -> Int { Int(next() % UInt64(max(1, count))) }
}

var fuzzFailures: [String: Int] = [:]
var fuzzRuns: [String: Int] = [:]
for seed in 1...(CommandLine.arguments.count > 1 ? Int(CommandLine.arguments[1]) ?? 300 : 300) {
    var random = Random(state: UInt64(seed) &* 7919)
    let kinds: [(String, Answering)] = [("characters", .characters), ("wechat", .wechat), ("poisoned", .poisoned(1 + random.int(3))), ("late", .late)]
    let (label, answering) = kinds[seed % kinds.count]
    var field = SimField()
    field.right = random.real(170, 520)
    field.caretX = random.int(3) == 0 ? 0 : random.real(0, field.right * 0.6)
    field.filler = random.real(7, 14)
    field.joint = field.filler * random.real(0.3, 0.5)
    field.hair = field.filler * random.real(0.08, 0.16)
    field.pitch = random.real(16, 26)
    if random.int(3) == 0 { field.visibleLines = 2 + random.int(3) }
    let off = random.real(0.86, 1.16)
    // (the one field known to answer for the text before, VS Code, sets its text in cells on one line: a joint is
    // as wide as a filler, the hair up to as wide. A late field with other widths would have its joint mismeasured
    // by the one character a restructured text adds: none is known.)
    if label == "late" { field.right = 100000; field.joint = field.filler; field.hair = field.filler * random.real(0.1, 1) }
    let session = Session("fuzz \(seed) \(label)", field: field, answering: answering, estimate: .init(hair: field.hair * off, filler: field.filler * off, joint: field.joint * off))
    var widths: [CGFloat] = []
    for step in 0..<70 {
        let action = widths.isEmpty ? 0 : random.int(10)
        if action < 5 {
            // typing at the end: the last piece grows, or another begins
            if widths.isEmpty || widths[widths.count - 1] > random.real(14, 34) { widths.append(random.real(5, 12)) } else { widths[widths.count - 1] += random.real(5, 11) }
        } else if action < 7 {
            // deleting at the end
            widths[widths.count - 1] -= random.real(5, 11)
            if widths[widths.count - 1] < 3 { widths.removeLast() }
            if widths.isEmpty { widths = [random.real(5, 12)] }
        } else if action < 9 {
            // an earlier piece grows
            widths[random.int(widths.count)] += random.real(5, 11)
        } else {
            let index = random.int(widths.count)
            widths[index] = max(5, widths[index] - random.real(5, 11))
        }
        session.type(widths, tolerance: 3, note: "step \(step) action \(action)")
        if CommandLine.arguments.count > 2, Int(CommandLine.arguments[2]) == seed {
            let truth = session.field.places(session.layout.text)
            let m = session.model
            print("  step \(step) action \(action) counts=\(session.layout.counts) origin=\(session.fmt(m.origin)) from=\(m.originFrom) left=\(m.left.map(session.fmt) ?? "-") pitch=\(m.pitch.map(session.fmt) ?? "-") right=\(session.fmt(m.rightLow))..\(session.fmt(m.rightHigh)) confirmed=\(m.confirmed)")
            print("      model: \(m.arrangement.pieces.map(session.fmt).joined(separator: " ")) end \(session.fmt(m.arrangement.end))")
            let r = session.reads(session.layout.text)
            print("      reads: \((0..<r.length).map { r.at[$0].map(session.fmt) ?? "-" }.joined(separator: " "))")
            print("      truth: \(session.layout.starts.map { session.fmt(truth[$0]) }.joined(separator: " ")) end \(session.fmt(truth[truth.count - 1]))  box right=\(session.fmt(session.field.right - session.field.left - session.field.caretX)) left=\(session.fmt(-session.field.caretX)) visible=\(session.field.visibleLines ?? 0)")
        }
        if session.failures > 3 { break }
    }
    fuzzRuns[label, default: 0] += 1
    if session.failures > 0 { fuzzFailures[label, default: 0] += 1 }
    total += session.failures
}
for label in fuzzRuns.keys.sorted() { print("fuzz \(label): \(fuzzFailures[label] ?? 0) of \(fuzzRuns[label]!) runs failed") }

print(total == 0 ? "ALL PASSED" : "\(total) FAILURES")
exit(total == 0 ? 0 : 1)
