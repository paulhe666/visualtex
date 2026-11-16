import Carbon
import Cocoa
import InputMethodKit

/// One per host text field. VisualTeX's input method types formulas and nothing else, and it is no
/// keyboard input source: a palette input method, selected beside whatever the user types with (as
/// the system's press-and-hold accents and its character viewer are), so the user's input sources
/// and the switching between them are never touched. It gets no keys of its own.
///
/// VisualTeX's shortcut opens a formula in the field that has the keyboard: VisualTeX asks where
/// the caret is ("caret"), opens the formula ("session"), edits it and draws it in the line itself,
/// from where the field's caret was, with its own caret. The field holds the formula's space in its
/// text as marked text (FormulaSpace), so it wraps the formula as it wraps its own text. Where that
/// space lies is worked out by SpaceModel from the field's caret before and after it; the field's
/// other answers only confirm or correct it. While the formula is open the keys are taken before
/// the field and its input source see them (KeyGrab) and go to VisualTeX. Return inserts the
/// formula's text with insertText.
@objc(VTFormulaController)
final class VTFormulaController: IMKInputController {
    static let noRange = NSRange(location: NSNotFound, length: 0)

    /// The application of the field the shortcut was pressed in, as VisualTeX tells it (a field
    /// whose application does not tell the input method its name).
    static var shortcutApplication: String?

    /// The controller whose field has the keyboard: VisualTeX's messages go to it.
    private static weak var active: VTFormulaController?
    /// (for HostProbe)
    static var current: VTFormulaController? { active }
    /// The height of a line of text by application, measured where text
    /// stands before the caret. Chrome and Electron report
    /// an empty line's whole line box instead (taller than its text), and no
    /// font: VisualTeX sizes the formula from this so that it does not depend
    /// on whether the line has text yet.
    private static var textHeights: [String: CGFloat] = [:]

    /// The marked text this field shows (the formula's space).
    private var marked = ""
    /// What the formula inserts, for a field that ends the composition itself.
    private var formulaText = ""
    private var lastGeometry: [String: Any] = [:]
    /// Where the formula starts: its first piece's start on its line (AppKit
    /// screen coordinates), zero until the field's caret is known.
    private var heldRect = NSRect.zero
    /// Each space held; a poll for an older one stops.
    private var holdGeneration = 0
    /// A formula is open: its keys belong to VisualTeX.
    private var composing = false
    /// Where the field's caret stood when the formula opened, before its
    /// space went in: SpaceModel's places are from there.
    private var formulaStart = NSRect.zero
    /// The formula's pieces (points wide, from VisualTeX), the marked text
    /// holding them (FormulaSpace) and where that lies in the field.
    private var pieceWidths: [CGFloat] = []
    private var space = FormulaSpace.Layout(text: "", starts: [], counts: [])
    private var model: SpaceModel?
    /// A filler's and a joint's widths by application and line height, as
    /// measured there (points), with how many fillers the measure spanned and
    /// the hair's width; kept across starts.
    private static var learntWidths: [String: [Double]] = UserDefaults.standard.dictionary(forKey: "VisualTeXLearntWidths3") as? [String: [Double]] ?? [:]
    /// The line the formula starts on as the field tells it for a character
    /// there: its bottom and height (the caret's before the field has told).
    private var anchorLine: (minY: CGFloat, height: CGFloat)?
    /// The space's last character's line as the field tells it (its bottom and
    /// height on the screen): the field's own caret stands after it, as tall.
    private var lastCharacterLine: (minY: CGFloat, height: CGFloat)?
    /// Spaces held again in this formula because the widths were measured
    /// better: a few at most (a field rounding places to pixels must not make
    /// the space flip between two widths).
    private var reholds = 0
    /// When the space was last held.
    private var lastHold = Date.distantPast
    /// The formula's place and its pieces' places last reported to VisualTeX.
    private var lastPlaces: [[Double]] = []
    private var lastEnd: [Double]?
    private var lastCaretLine: [Double]?
    private var placed = false
    /// Applications that answer nothing after the marked text's last character
    /// (Chrome, Electron): they keep a composition's text when they end it
    /// themselves, and tell nothing while they do (endedByField).
    private static var lateClients: Set<String> = []
    /// Where the field had the held space in its text when it last told, and
    /// the space then; polls since it told of no marked text at all.
    private var heldRange = NSRange(location: NSNotFound, length: 0)
    private var heldText = ""
    private var lostPolls = 0
    /// Held spaces a field may have kept as its own text when it ended a
    /// composition (blank characters nobody sees), by application: taken out
    /// when found before the caret (takeOutLeftovers).
    private static var leftovers: [String: [Left]] = [:]
    /// A held space is being taken out (its outcome is read in a moment).
    private var takingOut = false

    /// A held space a field may have kept as its text: where it was in the
    /// field's text and on the screen, when, and what goes in its place (the formula's text when the field
    /// ended the composition; nothing when the input method left the field:
    /// VisualTeX keeps the formula).
    private struct Left {
        var text: String
        var location: Int
        var at: Date
        var first: NSRect
        var end: [Double]?
        var replacement: String
    }




    override func activateServer(_ sender: Any!) {
        super.activateServer(sender)
        Self.active = self
        // (the field's application as it tells it: not the shortcut's application, which may be another)
        let application = (sender as? IMKTextInput)?.bundleIdentifier() ?? "?"
        InputLog.write("activate client=\(application) linked=\(AppLink.shared.connect())")
        AppLink.shared.send(["t": "activate", "client": application])
        // a held space this field kept as its text when it lost the keyboard
        if Self.leftovers[clientID(sender)]?.isEmpty == false {
            DispatchQueue.main.asyncAfter(deadline: .now() + 0.3) { [weak self] in
                guard let self, Self.active === self, !self.composing, self.marked.isEmpty, let client = self.client() else { return }
                self.takeOutLeftovers(client)
            }
        }
    }

    /// The field lost the keyboard: an open formula goes (VisualTeX keeps its draft).
    override func deactivateServer(_ sender: Any!) {
        InputLog.write("deactivate client=\(clientID(sender)) composing=\(composing)")
        let client = sender as? IMKTextInput
        // (Chrome has kept the space as its text by now: taken out where it is; else cleared as
        // marked text is)
        if !marked.isEmpty, let client, !takeOutKept(client, with: "") { clearMarked(client) }
        marked = ""
        holdGeneration += 1
        if composing { KeyGrab.shared.stop() }
        composing = false
        formulaText = ""
        if Self.active === self { Self.active = nil }
        AppLink.shared.send(["t": "deactivate", "client": clientID(sender)])
        super.deactivateServer(sender)
    }

    /// The field ends the composition itself (a click elsewhere in it): it
    /// gets the formula's text, not the held space.
    override func commitComposition(_ sender: Any!) {
        InputLog.write("commitComposition composing=\(composing) text=\(formulaText.count)")
        if composing, let client = sender as? IMKTextInput { endedByField(client) }
    }

    /// The field ended the composition: the formula's text goes in where the
    /// space was. Chrome has kept the space as its own text by then (blank
    /// characters nobody sees, sent with the next message): the text replaces
    /// them where they are found; a field that still has the space marked
    /// has it replaced as marked text is.
    private func endedByField(_ client: IMKTextInput) {
        let text = formulaText
        if !takeOutKept(client, with: text) {
            if Self.lateClients.contains(clientID(client)) {
                // Chrome tells nothing of its text or its caret while it ends the composition: the
                // text goes in where the space was once it does (takeOutLeftovers)
                for delay in [0.15, 0.6] {
                    DispatchQueue.main.asyncAfter(deadline: .now() + delay) { [weak self] in
                        guard let self, !self.composing, self.marked.isEmpty, let client = self.client() else { return }
                        self.takeOutLeftovers(client)
                    }
                }
            } else if !text.isEmpty {
                insertReplacingMarked(text, client: client)
            } else if !marked.isEmpty {
                clearMarked(client)
            }
        }
        marked = ""
        formulaText = ""
        holdGeneration += 1
        space = FormulaSpace.Layout(text: "", starts: [], counts: [])
        model = nil
        composing = false
        KeyGrab.shared.stop()
        AppLink.shared.send(["t": "forced"])
    }

    private func isOnScreen(_ rect: NSRect) -> Bool {
        guard rect.height > 0, rect.minX.isFinite, rect.minY.isFinite else { return false }
        let point = NSPoint(x: rect.minX, y: rect.midY)
        return NSScreen.screens.contains { $0.frame.insetBy(dx: -40, dy: -40).contains(point) }
    }

    /// The held space, where the field last had it, replaced by `text` when
    /// the field's text there is the space (as much of it as the field tells:
    /// Chrome tells a hundred characters around its caret). The space is
    /// remembered as possibly left behind all the same (a field that has lost
    /// the keyboard drops what it is given).
    private func takeOutKept(_ client: IMKTextInput, with text: String) -> Bool {
        guard heldRange.location != NSNotFound else { return false }
        // (the space as it was set last, or as the field last told of it)
        var candidates: [String] = []
        for candidate in [marked, space.text, heldText] where !candidate.isEmpty && !candidates.contains(candidate) {
            candidates.append(candidate)
        }
        guard !candidates.isEmpty else { return false }
        let id = clientID(client)
        let location = heldRange.location
        Self.leftovers[id] = Array(((Self.leftovers[id] ?? []).filter { !candidates.contains($0.text) }
            + candidates.map { Left(text: $0, location: location, at: Date(), first: heldRect, end: lastEnd, replacement: text) }).suffix(6))
        heldRange = NSRange(location: NSNotFound, length: 0)
        heldText = ""
        for candidate in candidates {
            let range = NSRange(location: location, length: (candidate as NSString).length)
            guard keptThere(client, range, candidate) else { continue }
            replace(client, range, with: text)
            InputLog.write("the held space (\(range.length) characters at \(location)) replaced by \((text as NSString).length) characters")
            Self.leftovers[id]?.removeAll { candidates.contains($0.text) }
            return true
        }
        let selected = client.selectedRange()
        InputLog.write("the held space (\(candidates.map { String(($0 as NSString).length) }.joined(separator: "/")) characters at \(location)) is not found in the field's text: selected=\(selected.location == NSNotFound ? -1 : selected.location)+\(selected.length)")
        return false
    }

    /// `range` of the field's text replaced by `text`. (Chrome does nothing
    /// with an empty text: a marked character in the range's place, then no
    /// marked text.)
    private func replace(_ client: IMKTextInput, _ range: NSRange, with text: String) {
        if !text.isEmpty {
            client.insertText(text, replacementRange: range)
            return
        }
        client.setMarkedText("\u{200B}", selectionRange: NSRange(location: 0, length: 0), replacementRange: range)
        client.setMarkedText("", selectionRange: NSRange(location: 0, length: 0), replacementRange: Self.noRange)
    }

    private enum Told { case same, different, nothing }

    /// Whether the field's text in `range` is `text`, as far as the field
    /// tells it: its end at least (Chrome tells the text around its caret
    /// only, of a range reaching past that the part it has, and often
    /// nothing at all).
    private func textThere(_ client: IMKTextInput, _ range: NSRange, _ text: String) -> Told {
        let whole = text as NSString
        guard whole.length == range.length, range.length > 0 else { return .different }
        var checked = 0
        var end = range.length
        while end > 0 {
            let count = min(32, end)
            let part = NSRange(location: range.location + end - count, length: count)
            let expected = whole.substring(with: NSRange(location: end - count, length: count))
            guard let told = client.attributedSubstring(from: part)?.string, !told.isEmpty else { break }
            if told == expected {
                checked += count
                end -= count
                continue
            }
            // a part of what was asked: its end, as far as the field's text goes back
            let length = (told as NSString).length
            if length < count, expected.hasSuffix(told) {
                checked += length
                break
            }
            return .different
        }
        return checked >= min(range.length, 12) ? .same : (checked == 0 ? .nothing : .different)
    }

    /// The field kept the space in `range` as its text: its text there is the
    /// space; or, when it tells nothing of its text, it is Chrome (which
    /// keeps what it was composing) or its caret stands right after the
    /// space (a composition dropped would have left it where the space
    /// began).
    private func keptThere(_ client: IMKTextInput, _ range: NSRange, _ text: String) -> Bool {
        switch textThere(client, range, text) {
        case .same: return true
        case .different: return false
        case .nothing:
            let selected = client.selectedRange()
            guard selected.location != NSNotFound else { return false }
            // (Chrome keeps the text of a composition it ends itself, wherever the click that
            // ended it put the caret)
            return Self.lateClients.contains(clientID(client)) || (selected.location == NSMaxRange(range) && selected.length == 0)
        }
    }

    /// A held space this application's field kept as its text a moment ago
    /// (it lost the keyboard and dropped what it was given then), where it
    /// was, with the caret right after it: taken out. (Only what the field's
    /// text or, within a minute, its caret's place vouches for.)
    private func takeOutLeftovers(_ client: IMKTextInput) {
        let id = clientID(client)
        // (a field that lost the keyboard has the input method deactivated: it drops what it is
        // given then, and answers as if it had not)
        guard Self.active === self else { return }
        guard !takingOut, let lefts = Self.leftovers[id]?.filter({ Date().timeIntervalSince($0.at) < 60 }), !lefts.isEmpty else {
            if !takingOut { Self.leftovers[id] = [] }
            return
        }
        let selected = client.selectedRange()
        let markedRange = client.markedRange()
        guard selected.location != NSNotFound, selected.length == 0,
              markedRange.location == NSNotFound || markedRange.length == 0
        else { return }
        for left in lefts.sorted(by: { $0.text.count > $1.text.count }) {
            let length = (left.text as NSString).length
            guard length > 0 else { continue }
            let range = NSRange(location: left.location, length: length)
            let told = textThere(client, range, left.text)
            var vouched = told == .same
            if told == .nothing {
                // the caret right after the space, or (Chrome: a click ended the composition) in it,
                // and on the screen where the space was
                let there = selected.location == NSMaxRange(range)
                    || (Self.lateClients.contains(id) && selected.location >= range.location && selected.location <= NSMaxRange(range))
                vouched = there && caretWhere(left, client)
            }
            guard vouched else { continue }
            // The space out first, then what goes in its place. A field that has not the keyboard
            // (the click that ended the composition was outside it) drops what it is given and
            // still answers as if it had it: the space is gone when the caret has moved (it stood
            // after the space's start), and until then it is looked for again.
            let before = caretRect(client)
            replace(client, range, with: "")
            let sure = told == .same || selected.location <= range.location
            takingOut = true
            DispatchQueue.main.asyncAfter(deadline: .now() + 0.2) { [weak self] in
                guard let self else { return }
                self.takingOut = false
                guard Self.active === self, let client = self.client() else { return }
                let after = self.caretRect(client)
                let moved = abs(after.minX - before.minX) >= 2 || abs(after.minY - before.minY) >= 2
                guard sure || moved else {
                    InputLog.write("a held space left in \(id) (\(length) characters) is still there: the field has not the keyboard")
                    return
                }
                Self.leftovers[id] = []
                InputLog.write("took out \(length) characters a held space left in \(id)")
                // (a formula opened meanwhile takes the place)
                if !left.replacement.isEmpty, !self.composing, self.marked.isEmpty {
                    client.insertText(left.replacement, replacementRange: NSRange(location: left.location, length: 0))
                    InputLog.write("the formula's text (\((left.replacement as NSString).length) characters) in its place")
                }
            }
            return
        }
    }

    /// The field's caret is on the screen where that space was.
    private func caretWhere(_ left: Left, _ client: IMKTextInput) -> Bool {
        let caret = caretRect(client)
        guard isOnScreen(caret), isOnScreen(left.first) else { return false }
        let height = max(left.first.height, caret.height)
        let down = left.end.flatMap { $0.count > 1 ? CGFloat($0[1]) : nil } ?? 0
        guard caret.midY <= left.first.maxY + height * 0.5, caret.midY >= left.first.minY - down - height * 0.5 else { return false }
        if down < height * 0.5, let across = left.end?.first {
            return caret.minX >= left.first.minX - 4 && caret.minX <= left.first.minX + CGFloat(across) + 4
        }
        return true
    }

    // MARK: VisualTeX's messages (main thread)

    /// Counts VisualTeX's "grab" requests (the newest one's safety stop is the one that counts).
    private static var grabs = 0

    static func receive(_ message: [String: Any]) {
        guard let type = message["t"] as? String else { return }
        if let application = message["app"] as? String, !application.isEmpty { shortcutApplication = application }
        if type == "grab" {
            // VisualTeX's pasting strip, for a field this input method does not reach (an app that
            // never gets it: Notes, WeChat, QQ): until the strip is up (its caret is looked for on the
            // screen meanwhile) the keys typed are taken here and sent to VisualTeX, never the field's.
            grabs += 1
            let grab = grabs
            if message["on"] as? Bool == true {
                KeyGrab.shared.start(skipping: message["skip"] as? [[Int]] ?? [])
                // whatever happens to VisualTeX, the keys are given back
                DispatchQueue.main.asyncAfter(deadline: .now() + 3) {
                    if grab == Self.grabs, Self.active?.composing != true, KeyGrab.shared.isOn {
                        InputLog.write("grab timed out")
                        KeyGrab.shared.stop()
                    }
                }
            } else if active?.composing != true {
                KeyGrab.shared.stop()
            }
            InputLog.write("grab on=\(KeyGrab.shared.isOn)")
            return
        }
        guard let controller = active, let client = controller.client() else {
            InputLog.write("message \(type) without an active field")
            // (the shortcut asks where the caret is: no field of this input method has the keyboard)
            if type == "caret" { AppLink.shared.send(["t": "nofield"]) }
            KeyGrab.shared.stop()
            return
        }
        switch type {
        case "disconnected":
            if controller.composing {
                controller.clearMarked(client)
                controller.composing = false
            }
            KeyGrab.shared.stop()
        case "caret":
            // VisualTeX's shortcut: where the formula would go
            var reply = controller.geometry(client)
            reply["t"] = "geometry"
            reply["client"] = controller.clientID(client)
            AppLink.shared.send(reply)
        case "session":
            // VisualTeX opened (the shortcut) or closed a formula for this field
            let on = message["on"] as? Bool ?? false
            controller.composing = on
            controller.formulaStart = on ? controller.caretRect(client) : .zero
            controller.heldRect = .zero
            controller.holdGeneration += 1
            controller.pieceWidths = []
            controller.lastPlaces = []
            controller.lastEnd = nil
            controller.lastCaretLine = nil
            controller.placed = false
            controller.anchorLine = nil
            controller.lastCharacterLine = nil
            controller.reholds = 0
            controller.heldRange = NSRange(location: NSNotFound, length: 0)
            controller.heldText = ""
            controller.lostPolls = 0
            if on { controller.takeOutLeftovers(client) }
            controller.space = FormulaSpace.Layout(text: "", starts: [], counts: [])
            controller.model = nil
            let before = on ? " \(controller.fieldState(client))" : ""
            if on {
                // the formula's keys from now on (but the chords VisualTeX has for itself)
                KeyGrab.shared.start(skipping: message["skip"] as? [[Int]] ?? [])
                controller.openSpace(client)
            } else {
                KeyGrab.shared.stop()
                if !controller.marked.isEmpty { controller.clearMarked(client) }
            }
            InputLog.write("session \(on ? "on" : "off")\(before)")
        case "inline":
            controller.formulaText = message["text"] as? String ?? ""
            if let pieces = message["pieces"] as? [Double], controller.composing {
                let widths = pieces.map { CGFloat(max(0, min($0, 20000))) }
                if widths != controller.pieceWidths {
                    controller.pieceWidths = widths
                    controller.hold(client)
                }
            }
        case "commit":
            controller.commit(message["text"] as? String ?? "", client: client)
        case "cancel":
            controller.cancel(client: client)
        default:
            break
        }
    }

    // MARK: Where the formula goes in the line

    /// A formula opens where the field's caret stands (formulaStart): the model of its space starts
    /// from the widths this application drew the blanks with before, or from an estimate.
    private func openSpace(_ client: IMKTextInput) {
        let estimated = estimatedWidths(client)
        var estimate = SpaceModel.Widths(hair: estimated.hair, filler: estimated.filler, joint: estimated.joint)
        var trust: CGFloat = 1
        if let learnt = Self.learntWidths[widthKey(client)], learnt.count == 4, learnt[0] > 0.5 {
            estimate.filler = CGFloat(learnt[0])
            estimate.joint = CGFloat(learnt[1])
            estimate.hair = CGFloat(learnt[3])
            trust = 6
        }
        let height = isOnScreen(formulaStart) ? formulaStart.height : (Self.textHeights[clientID(client)] ?? 17)
        model = SpaceModel(lineHeight: height, estimate: estimate, trust: trust)
        hold(client)
    }

    /// The widths' key: the application and the height of the line where the formula opened.
    private func widthKey(_ client: IMKTextInput) -> String {
        "\(clientID(client))@\(Int(formulaStart.height.rounded()))"
    }

    /// The blanks' widths before the field has shown one: in the field's font
    /// when it tells it, else from its line's height (Chrome and Electron
    /// report a default font, not the page's).
    private func estimatedWidths(_ client: IMKTextInput) -> (hair: CGFloat, filler: CGFloat, joint: CGFloat) {
        var line = NSRect.zero
        let attributes = client.attributes(forCharacterIndex: 0, lineHeightRectangle: &line) ?? [:]
        let font = attributes[NSAttributedString.Key.font] as? NSFont
            ?? attributes[NSAttributedString.Key.font.rawValue] as? NSFont
        if let font, line.height > 0, line.height < (font.ascender - font.descender) * 1.25 + 1 {
            return (FormulaSpace.unit(in: font, of: FormulaSpace.anchor), FormulaSpace.unit(in: font), FormulaSpace.unit(in: font, of: FormulaSpace.joint))
        }
        let height = Self.textHeights[clientID(client)] ?? (line.height > 6 && line.height < 120 ? line.height : 17)
        let size = height / 1.19
        return (FormulaSpace.unit(atSize: size, of: FormulaSpace.anchor), FormulaSpace.unit(atSize: size), FormulaSpace.unit(atSize: size, of: FormulaSpace.joint))
    }

    /// The field holds the formula's pieces (FormulaSpace). Where they lie is told to VisualTeX at
    /// once, as the model has it; what the field answers afterwards is taken in (poll).
    private func hold(_ client: IMKTextInput) {
        guard var model else { return }
        // held short of the formula by as much as the widths may be off: a filler measured over a run
        // of n fillers is off by up to about 0.7/n point a filler (the field gives places in whole
        // points), an estimated one by 5%, a joint not measured by a quarter of its width
        let widths = model.widths
        let unit = max(widths.filler, 0.5)
        let fillers = pieceWidths.reduce(0, +) / unit
        let joints = CGFloat(max(0, pieceWidths.count - 1))
        let span = model.fit.longest
        let margin = 0.15 + (span > 0 ? fillers * 0.7 / CGFloat(span) / unit : fillers * 0.05)
            + (model.fit.jointMeasured ? 0 : joints * 0.25 * widths.joint / unit)
        let lineStart = model.arrangement.lastLineStart(model.lineHeight * 0.4)
        let layout = FormulaSpace.layout(widths: pieceWidths, unit: widths.filler, joint: widths.joint, margin: margin,
                                         lineStart: lineStart < pieceWidths.count ? lineStart : 0)
        guard layout != space || marked != layout.text else { return }
        space = layout
        InputLog.write("hold pieces=\(layout.starts.count) fillers=\(layout.counts.reduce(0, +)) filler=\(String(format: "%.2f", Double(widths.filler))) joint=\(String(format: "%.2f", Double(widths.joint)))")
        lastHold = Date()
        holdGeneration += 1
        marked = layout.text
        // One marked text a key (set twice a key, the field's underline and caret moved twice a key
        // and each key flashed, 2026-11-10).
        client.setMarkedText(Self.heldString(layout.text), selectionRange: Self.heldSelection, replacementRange: Self.noRange)
        model.hold(layout)
        self.model = model
        // (the first time, once the field has told its line, or in a moment: poll)
        if placed { report(client) }
        poll(generation: holdGeneration, last: nil, polls: 0)
    }

    /// What the field answers for the held space, asked a few times after it is set (a field's first
    /// answer is often still of the text before) and then now and again while the formula is open
    /// (the field may move its text: scrolling, a window dragged). Answers that come twice the same
    /// are taken in by the model; VisualTeX is told when that changes where the formula lies.
    private func poll(generation: Int, last: FieldReads?, polls: Int) {
        guard generation == holdGeneration, composing, let client = client(), var model else { return }
        if polls == 0, !placed { InputLog.write("held \(fieldState(client))") }
        // Where the field has the space in its text; a field that says it has no marked text any
        // more (it had) ended the composition on its own and kept the space as its text (a click in
        // it; Chrome when its window or its page's focus changes) without telling: as when it tells
        // (commitComposition).
        let markedRange = client.markedRange()
        if markedRange.location != NSNotFound, markedRange.length == (marked as NSString).length {
            heldRange = markedRange
            heldText = marked
            lostPolls = 0
        } else if markedRange.location == NSNotFound, heldRange.location != NSNotFound {
            lostPolls += 1
            if lostPolls >= 4 {
                InputLog.write("the field has no marked text any more: the composition ended")
                endedByField(client)
                return
            }
        }
        let started = Date()
        let (reads, lines) = read(client)
        let elapsed = Date().timeIntervalSince(started)
        if reads == last, !reads.at.isEmpty {
            let before = (model.answers, model.origin, model.arrangement.pieces.last?.y ?? 0)
            let changed = model.absorb(reads, sinceHold: Date().timeIntervalSince(lastHold))
            self.model = model
            noteLine(lines, model)
            lastCharacterLine = lines[reads.length - 1]
            if before.0 != model.answers || before.1 != model.origin || before.2 != (model.arrangement.pieces.last?.y ?? 0) {
                InputLog.write("space answers=\(model.answers.rawValue) origin=\(Int(model.origin.x)),\(Int(model.origin.y)) from=\(model.originFrom) lines=\(Set(model.arrangement.pieces.map { Int($0.y.rounded()) }).count) confirmed=\(model.confirmed) filler=\(String(format: "%.2f", Double(model.widths.filler))) joint=\(String(format: "%.2f", Double(model.widths.joint))) \(fieldState(client))")
            }
            if !placed || changed {
                placed = true
                report(client)
            }
            // widths measured better: remembered for this application, and the space held again when
            // it would be held otherwise with them
            if model.fit.longest >= 5 {
                let learnt = [Double(model.widths.filler), Double(model.widths.joint), Double(model.fit.longest), Double(model.widths.hair)]
                if Self.learntWidths[widthKey(client)] != learnt {
                    Self.learntWidths[widthKey(client)] = learnt
                    UserDefaults.standard.set(Self.learntWidths, forKey: "VisualTeXLearntWidths3")
                }
            }
            if changed, reholds < 8 {
                let generation = holdGeneration
                hold(client)
                if holdGeneration != generation {
                    reholds += 1
                    return
                }
            }
        } else if !placed, polls >= 4 {
            // a field that tells nothing (or never the same twice): the formula where its caret stood
            placed = true
            report(client)
        }
        // often at first (the field lays the space out within a frame or two), then now and again; a
        // field slow to answer is asked less often
        let delay: TimeInterval = polls < 6 && elapsed < 0.02 ? 0.016 : (polls < 12 && elapsed < 0.05 ? 0.05 : 0.25)
        DispatchQueue.main.asyncAfter(deadline: .now() + delay) { [weak self] in
            self?.poll(generation: generation, last: reads, polls: polls + 1)
        }
    }

    /// What the field answers for the held space: the place of the hair, of each piece's first three
    /// characters and its last (the space's last character among them), from where the caret stood
    /// (SpaceModel's places); and each answered character's line (its bottom and height on the
    /// screen). Nothing at or past the marked text's end is asked: TextKit and WeChat answer nothing
    /// for it and then answer wrongly for other characters.
    private func read(_ client: IMKTextInput) -> (FieldReads, [Int: (minY: CGFloat, height: CGFloat)]) {
        let length = (space.text as NSString).length
        var wanted = Set([0])
        for (piece, start) in space.starts.enumerated() {
            let count = space.counts[piece]
            for offset in [0, 1, 2, count - 1] where offset >= 0 && offset < count { wanted.insert(start + offset) }
        }
        var at: [Int: Place] = [:]
        var lines: [Int: (minY: CGFloat, height: CGFloat)] = [:]
        for index in wanted.sorted() where index < length {
            guard let rect = characterRect(client, index) else { continue }
            // (a field that told no caret before the space: its first answer is where the formula starts)
            if !isOnScreen(formulaStart) { formulaStart = rect }
            at[index] = Place(x: rect.minX - formulaStart.minX, y: formulaStart.minY - rect.minY)
            lines[index] = (rect.minY, rect.height)
        }
        return (FieldReads(length: length, at: at), lines)
    }

    /// The line the formula starts on, as the field tells it for a character on it.
    private func noteLine(_ lines: [Int: (minY: CGFloat, height: CGFloat)], _ model: SpaceModel) {
        guard let first = model.arrangement.pieces.first else { return }
        let wanted = formulaStart.minY - first.y
        let tolerance = model.lineHeight * 0.4
        for index in lines.keys.sorted() {
            guard let line = lines[index], abs(line.minY - wanted) <= tolerance else { continue }
            anchorLine = line
            return
        }
        anchorLine = nil
    }

    /// Tells VisualTeX where the formula lies: where its first piece starts on the screen (with how
    /// the field's text looks there), and each piece from there.
    private func report(_ client: IMKTextInput) {
        guard let model, isOnScreen(formulaStart), let first = model.arrangement.pieces.first else { return }
        let arrangement = model.arrangement
        // The line the formula stands on: as the field tells it for a character of the space, or, on
        // the line where the caret stood, as it told for the caret, whichever is the less tall (the
        // tighter about the text: WebKit's blanks come from another font and their line is a point
        // or two lower than the text's, Chrome's empty line is its whole line box).
        var line = anchorLine ?? (formulaStart.minY - first.y, formulaStart.height)
        if anchorLine != nil, abs(first.y - model.origin.y) < model.lineHeight * 0.4, abs(model.origin.y) < model.lineHeight * 0.4,
           formulaStart.height >= 6, formulaStart.height < line.height - 0.5 {
            line = (formulaStart.minY, formulaStart.height)
        }
        let rect = NSRect(x: formulaStart.minX + first.x, y: line.minY, width: 0, height: line.height)
        if abs(rect.minX - heldRect.minX) > 0.05 || abs(rect.minY - heldRect.minY) > 0.05 || abs(rect.height - heldRect.height) > 0.05 {
            noteTextHeight(client, rect)
            heldRect = rect
            var message = geometry(client)
            message["t"] = "geometry"
            message["caret"] = payload(rect)
            InputLog.write("anchor \(NSStringFromRect(rect)) answers=\(model.answers.rawValue) \(fieldState(client))")
            if !NSDictionary(dictionary: message).isEqual(to: lastGeometry) {
                lastGeometry = message
                AppLink.shared.send(message)
            }
        }
        let places = arrangement.pieces.map { [Double(($0.x - first.x).rounded(toPlaces: 2)), Double(($0.y - first.y).rounded(toPlaces: 2))] }
        // (an end worked out with widths not measured yet is marked as estimated: VisualTeX covers no
        // further than its formula's end)
        var end = [Double((arrangement.end.x - first.x).rounded(toPlaces: 2)), Double((arrangement.end.y - first.y).rounded(toPlaces: 2))]
        if model.fit.longest == 0 { end.append(1) }
        // the field's caret after the space: how far its line's bottom is below the last piece's line
        // as told here, and how tall it is (VisualTeX covers it)
        var caretLine: [Double]?
        if let last = arrangement.pieces.last, let character = lastCharacterLine, character.height >= 6, character.height < 120 {
            let bottom = rect.minY - (last.y - first.y)
            // (a caret as tall as the field's whole line, where the lines are further apart than a
            // character's line is tall: WebKit's, with the character's line in the middle of it)
            let spare = max(0, min((model.pitch ?? character.height) - character.height, character.height)) / 2
            let below = bottom - character.minY + spare
            let height = character.height + spare * 2
            if abs(below) < height { caretLine = [Double(below.rounded(toPlaces: 1)), Double(height.rounded(toPlaces: 1))] }
        }
        guard places != lastPlaces || end != lastEnd || caretLine != lastCaretLine else { return }
        lastPlaces = places
        lastEnd = end
        lastCaretLine = caretLine
        let lines = Set(places.map { $0[1] }).count
        InputLog.write("pieces \(places.count) lines=\(lines) \(places.prefix(6).map { "\(Int($0[0])),\(Int($0[1]))" }.joined(separator: " ")) end=\(Int(end[0])),\(Int(end[1]))\(end.count > 2 ? "~" : "")")
        var message: [String: Any] = ["t": "pieces", "places": places, "count": space.starts.count, "end": end]
        if let caretLine { message["caretLine"] = caretLine }
        // the selected anchor before the formula (the field may highlight it)
        if abs(arrangement.hair.y - first.y) < model.lineHeight * 0.4 {
            message["start"] = [Double((arrangement.hair.x - first.x).rounded(toPlaces: 2)), 0]
        }
        AppLink.shared.send(message)
    }

    /// The selection in the held space: none, right after its anchor. The field's own caret then
    /// stands at the formula's start, inside the space (VisualTeX covers it there), and stays there
    /// whatever is typed (set after the space, it moved with every key and showed beside the
    /// formula's caret until VisualTeX had covered it again).
    static var heldSelection = NSRange(location: (FormulaSpace.anchor as NSString).length, length: 0)

    /// The held space as marked text: with the thin underline. (Given no
    /// attribute the field draws a thick one. Its colour is the system's,
    /// whatever colour is given: an underline nobody sees cannot be asked
    /// for, 2026-11-10.)
    private static func heldString(_ text: String) -> NSAttributedString {
        NSAttributedString(string: text, attributes: [.underlineStyle: NSUnderlineStyle.single.rawValue])
    }

    /// What the field tells of its marked text and selection, for the log:
    /// their ranges, the places of the marked text's first two characters and
    /// its last one, and where the formula opened.
    private func fieldState(_ client: IMKTextInput) -> String {
        let markedRange = client.markedRange()
        let selected = client.selectedRange()
        let length = (marked as NSString).length
        let range = { (range: NSRange) in range.location == NSNotFound ? "-" : "\(range.location)+\(range.length)" }
        let place = { (index: Int) -> String in
            // (only characters of the marked text; the caret when there is none)
            guard index >= 0, index < max(length, 1) else { return "-" }
            var rect = NSRect.zero
            _ = client.attributes(forCharacterIndex: index, lineHeightRectangle: &rect)
            return "\(Int(rect.minX)),\(Int(rect.minY))"
        }
        return "marked=\(range(markedRange)) selected=\(range(selected)) at0=\(place(0)) at1=\(place(1)) last=\(place(length - 1)) start=\(Int(formulaStart.minX)),\(Int(formulaStart.minY))"
    }

    /// Whether this application answers nothing for the place after the marked text's last
    /// character (lateClients), asked once an application, as the formula closes: asked while the
    /// formula is open, TextKit and WeChat answer wrongly for other characters afterwards.
    private static var askedPastEnd: Set<String> = []
    private func notePastEnd(_ client: IMKTextInput) {
        let id = clientID(client)
        let length = (marked as NSString).length
        guard length >= 2, !Self.askedPastEnd.contains(id) else { return }
        Self.askedPastEnd.insert(id)
        if characterRect(client, length) == nil { Self.lateClients.insert(id) }
    }

    /// `text` where the held space is: marked text, replaced as such. Text with line breaks (a display
    /// formula) goes in once the space is cleared: a rich editor (Gemini, ChatGPT, Claude: Quill,
    /// ProseMirror) that has a composition replaced by it splits its paragraphs first and puts its caret
    /// one character short, so the formula's last character (its closing brace) went after the closing
    /// $$ (2026-11-16: reproduced with both editors in Chromium at the field's start, end and middle;
    /// cleared first the text is exact in them and in a plain text field too).
    private func insertReplacingMarked(_ text: String, client: IMKTextInput) {
        if text.contains("\n"), !marked.isEmpty { clearMarked(client) }
        client.insertText(text, replacementRange: Self.noRange)
    }

    private func commit(_ text: String, client: IMKTextInput) {
        KeyGrab.shared.stop()
        notePastEnd(client)
        insertReplacingMarked(text, client: client)
        InputLog.write("commit length=\((text as NSString).length)")
        marked = ""
        formulaText = ""
        composing = false
        space = FormulaSpace.Layout(text: "", starts: [], counts: [])
        model = nil
    }

    private func cancel(client: IMKTextInput) {
        KeyGrab.shared.stop()
        notePastEnd(client)
        clearMarked(client)
        formulaText = ""
        composing = false
        space = FormulaSpace.Layout(text: "", starts: [], counts: [])
        model = nil
        InputLog.write("cancel")
    }

    private func clearMarked(_ client: IMKTextInput) {
        marked = ""
        client.setMarkedText("", selectionRange: NSRange(location: 0, length: 0), replacementRange: Self.noRange)
    }

    /// A character of the marked text (by its index in it) as the field
    /// places it: a caret-like rectangle at its left edge, the line's height
    /// (AppKit screen coordinates). nil when the field answers with nothing
    /// on a screen. Every application measured answers this for the marked
    /// text's characters; firstRect is garbage in Chrome, WebKit and WeChat
    /// and a line off in TextKit.
    private func characterRect(_ client: IMKTextInput, _ index: Int) -> NSRect? {
        var rect = NSRect.zero
        _ = client.attributes(forCharacterIndex: index, lineHeightRectangle: &rect)
        return isOnScreen(rect) ? rect : nil
    }

    /// The marked text's rectangle: from its first character's left edge to
    /// the edge after its last character (its place and one character's
    /// advance: the place after it is never asked), on the first character's line.
    private func markedRect(_ client: IMKTextInput) -> NSRect? {
        let length = (marked as NSString).length
        guard length > 0, let first = characterRect(client, 0) else { return nil }
        var right = first.minX
        if length > 1, let last = characterRect(client, length - 1), abs(last.minY - first.minY) < first.height * 0.5, last.minX >= first.minX {
            right = max(last.maxX, last.minX + (last.minX - first.minX) / CGFloat(length - 1))
        }
        return NSRect(x: first.minX, y: first.minY, width: max(right - first.minX, 0), height: first.height)
    }

    /// The caret line's rectangle in screen coordinates (AppKit, y up) as the
    /// field's text system reports it: the marked text's start while composing.
    func caretRect(_ sender: Any?) -> NSRect {
        guard let client = sender as? IMKTextInput else { return .zero }
        var rect = NSRect.zero
        _ = client.attributes(forCharacterIndex: 0, lineHeightRectangle: &rect)
        return rect
    }

    /// Where the formula goes and how the field's text looks there: the
    /// marked text's rectangle (or the caret's line), the font size and
    /// descent, and the text colour when the field tells it.
    private func geometry(_ client: IMKTextInput?) -> [String: Any] {
        guard let client else { return [:] }
        var line = NSRect.zero
        let attributes = client.attributes(forCharacterIndex: 0, lineHeightRectangle: &line) ?? [:]
        let rect = markedRect(client) ?? line
        let markedRange = client.markedRange()
        var message: [String: Any] = ["caret": payload(rect), "line": payload(line)]
        let font = attributes[NSAttributedString.Key.font] as? NSFont
            ?? attributes[NSAttributedString.Key.font.rawValue] as? NSFont
        if let font {
            message["fontSize"] = Double(font.pointSize)
            message["ascent"] = Double(font.ascender)
            message["descent"] = Double(-font.descender)
            if font.pointSize > 0 { message["xHeight"] = Double(font.xHeight / font.pointSize) }
        }
        if let color = textColor(client, before: markedRange) { message["color"] = color }
        message["textBefore"] = hasTextBefore(client)
        if let height = Self.textHeights[clientID(client)] { message["textHeight"] = Double(height) }
        return message
    }

    /// Text stands before the caret (or the composition) on its line.
    private func hasTextBefore(_ client: IMKTextInput) -> Bool {
        var start = client.markedRange().location
        if start == NSNotFound { start = client.selectedRange().location }
        guard start != NSNotFound, start > 0,
              let previous = client.attributedSubstring(from: NSRange(location: start - 1, length: 1))?.string,
              let character = previous.unicodeScalars.first
        else { return false }
        return !CharacterSet.newlines.contains(character)
    }

    private func noteTextHeight(_ client: IMKTextInput, _ rect: NSRect) {
        guard rect.height >= 6, rect.height <= 80, hasTextBefore(client) else { return }
        Self.textHeights[clientID(client)] = rect.height
    }

    /// The colour of the character before the composition, as #rrggbbaa.
    private func textColor(_ client: IMKTextInput, before markedRange: NSRange) -> String? {
        var start = markedRange.location
        if start == NSNotFound { start = client.selectedRange().location }
        guard start != NSNotFound, start > 0,
              let sample = client.attributedSubstring(from: NSRange(location: start - 1, length: 1)),
              sample.length > 0,
              let color = sample.attribute(.foregroundColor, at: 0, effectiveRange: nil) as? NSColor,
              let rgb = color.usingColorSpace(.sRGB)
        else { return nil }
        let channel = { (value: CGFloat) in Int((max(0, min(1, value)) * 255).rounded()) }
        return String(format: "#%02x%02x%02x%02x", channel(rgb.redComponent), channel(rgb.greenComponent), channel(rgb.blueComponent), channel(rgb.alphaComponent))
    }

    private func payload(_ rect: NSRect) -> [Double] {
        [Double(rect.minX), Double(rect.minY), Double(rect.width), Double(rect.height)]
    }

    /// The field's application (its bundle identifier): as the field tells it, else the one
    /// VisualTeX's shortcut was pressed in.
    private func clientID(_ sender: Any?) -> String {
        (sender as? IMKTextInput)?.bundleIdentifier() ?? Self.shortcutApplication ?? "?"
    }
}

private extension CGFloat {
    func rounded(toPlaces places: Int) -> CGFloat {
        let scale = pow(10, CGFloat(places))
        return (self * scale).rounded() / scale
    }
}
