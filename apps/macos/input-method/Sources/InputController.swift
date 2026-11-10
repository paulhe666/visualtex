import Carbon
import Cocoa
import InputMethodKit

/// One per host text field. VisualTeX is a whole input method: Chinese
/// through pinyin (RIME) and English (Shift switches), with formulas as one
/// more mode. While a formula is open (the shortcut) every key but a Command
/// shortcut goes to VisualTeX, which edits the formula and draws it in the
/// line itself, from where the field's caret was, with its own caret: the
/// field holds the formula's space in its text (FormulaSpace), so it wraps
/// the formula as it wraps its own text and shows nothing of its own. Return inserts the formula's
/// text with insertText and, when the shortcut opened the formula from
/// another input source, hands the field back to it.
@objc(VTInputController)
final class VTInputController: IMKInputController {
    static let noRange = NSRange(location: NSNotFound, length: 0)

    /// Return, keypad Enter, Z and V by virtual key code: with Command they
    /// act on the formula while it is open (⌘V pastes into it: an image is
    /// recognized, text goes in).
    private static let formulaCommandKeys: Set<UInt16> = [36, 76, 6, 9]

    /// The controller whose field has the keyboard: VisualTeX's messages go to it.
    private static weak var active: VTInputController?
    /// The height of a line of text by application, measured where text
    /// stands before the caret (pinyin, formulas). Chrome and Electron report
    /// an empty line's whole line box instead (taller than its text), and no
    /// font: VisualTeX sizes the formula from this so that it does not depend
    /// on whether the line has text yet.
    private static var textHeights: [String: CGFloat] = [:]

    /// The marked text this field shows (pinyin's spelling, or the formula's anchor).
    private var marked = ""
    /// What the formula inserts, for a field that ends the composition itself.
    private var formulaText = ""
    private var lastGeometry: [String: Any] = [:]
    /// Where the formula starts: the anchor's rectangle as the field placed
    /// it (AppKit screen coordinates), zero until the field reports it.
    private var heldRect = NSRect.zero
    /// Each formula's anchor; a poll for an older one stops.
    private var holdGeneration = 0
    /// Where the pinyin being spelled started (the caret before it).
    private var pinyinStart = NSRect.zero
    /// A formula is open: its keys belong to VisualTeX.
    private var composing = false
    /// This field's pinyin session (0 until RIME is ready).
    private var rimeSession: RimeSessionId = 0
    /// The formula's caret is in text (\text{}): its keys spell pinyin too.
    private var formulaTextMode = false
    /// The formula's caret from the formula's start, in points.
    private var formulaCaretOffset: CGFloat = 0
    /// Pinyin is being spelled inside the formula (as its composition).
    private var formulaPinyin = false
    /// Where the formula opened (the caret's line), to tell a real marked
    /// text rectangle from a stale one.
    private var formulaStart = NSRect.zero
    /// The formula's pieces (points wide, from VisualTeX) and the marked text
    /// holding them (FormulaSpace).
    private var pieceWidths: [CGFloat] = []
    private var space = FormulaSpace.Layout(text: "", starts: [], counts: [])
    /// A filler's width by application, learnt from where the field put the
    /// pieces (points).
    private static var fillerWidths: [String: CGFloat] = [:]
    private var fillerWidth: CGFloat = 0
    /// A joint's (word space's) width, likewise.
    private static var jointWidths: [String: CGFloat] = [:]
    private var jointWidth: CGFloat = 0
    /// Calibrations in this formula: a few at most (a field rounding places
    /// to pixels must not make the space flip between two widths).
    private var calibrations = 0
    /// When the space was last held: the field draws it a moment later (calibrate).
    private var lastHold = Date.distantPast
    /// The pieces' places last reported to VisualTeX.
    private var lastPlaces: [[Double]] = []
    private var lastEnd: [Double]?
    /// The space was set again for a field that kept old places (once per change of the pieces).
    private var rekicked = false
    /// The formula's caret below the formula's first baseline, points (it may
    /// be on a later line).
    private var formulaCaretY: CGFloat = 0
    /// When pinyin last went into the field: the field takes a moment to
    /// place its caret after it (see holdLater).
    private var lastPinyinCommit = Date.distantPast
    /// The space is held a moment later (holdLater): nothing holds it meanwhile.
    private var holdDeferred = false
    /// The space was held afresh for a field that placed it where an earlier
    /// composition was (staleAnchor): twice a formula at most.
    private var staleHolds = 0
    /// The field placed this formula's space where its caret was (once it
    /// did, the field may move it: a window dragged, a line scrolled).
    private var anchorTrusted = false
    /// The field keeps answering a place that is not the space's: the
    /// formula stays where the caret was.
    private var anchorDistrusted = false
    /// Applications that learn a marked text's places one set late (hold).
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
    /// How often the space held now was set again for places of an earlier
    /// one (stalePlaces), and a field where that changed nothing.
    private var staleKicks = 0
    private var kickedText = ""
    /// The hold whose places were checked (once each, when typing pauses),
    /// and the one whose places passed.
    private var placesChecked = -1
    private var placesSound = -1
    private var kicksUseless = false
    /// The place that looked stale and had the space held afresh: answered
    /// again for the fresh space, it is the field's own.
    private var suspectAnchor: NSRect?
    /// How far from the caret an application placed a formula's start when
    /// that was its own doing (staleAnchor), points.
    private static var anchorOffsets: [String: CGFloat] = [:]

    deinit {
        RimeEngine.shared.destroySession(rimeSession)
    }

    override func activateServer(_ sender: Any!) {
        super.activateServer(sender)
        Self.active = self
        let geometry = self.geometry(sender as? IMKTextInput)
        InputLog.write("activate client=\(clientID(sender)) geometry=\(geometry) linked=\(AppLink.shared.connect())")
        var message = geometry
        message["t"] = "activate"
        message["client"] = clientID(sender)
        AppLink.shared.send(message)
        // a held space this field kept as its text when the input method left it
        if Self.leftovers[clientID(sender)]?.isEmpty == false {
            DispatchQueue.main.asyncAfter(deadline: .now() + 0.3) { [weak self] in
                guard let self, Self.active === self, !self.composing, self.marked.isEmpty, let client = self.client() else { return }
                self.takeOutLeftovers(client)
            }
        }
    }

    override func deactivateServer(_ sender: Any!) {
        InputLog.write("deactivate client=\(clientID(sender)) composing=\(composing)")
        let client = sender as? IMKTextInput
        // (Chrome has kept the space as its text by now: taken out where it is; else cleared as
        // marked text is)
        if !marked.isEmpty, let client, !takeOutKept(client, with: "") { clearMarked(client) }
        marked = ""
        holdGeneration += 1
        if composing { setFormulaInputMode(on: false) }
        composing = false
        formulaText = ""
        if let client { finishPinyin(client) }
        CandidateWindow.shared.hide()
        if Self.active === self { Self.active = nil }
        AppLink.shared.send(["t": "deactivate", "client": clientID(sender)])
        super.deactivateServer(sender)
    }

    override func recognizedEvents(_ sender: Any!) -> Int {
        Int(NSEvent.EventTypeMask.keyDown.rawValue | NSEvent.EventTypeMask.flagsChanged.rawValue)
    }

    override func handle(_ event: NSEvent!, client sender: Any!) -> Bool {
        guard let event, let client = sender as? IMKTextInput else { return false }
        Self.active = self
        if event.type == .flagsChanged {
            // Shift switches Chinese and English in a formula too.
            return pinyinModifiers(event, client: client)
        }
        guard event.type == .keyDown else { return false }
        // a held space this field kept as its text (it has the keyboard now)
        if !composing, marked.isEmpty, Self.leftovers[clientID(client)]?.isEmpty == false { takeOutLeftovers(client) }
        if composing, [36, 53, 76].contains(event.keyCode) {
            InputLog.write("formula key \(event.keyCode == 53 ? "escape" : "return")")
        }
        if composing {
            if event.modifierFlags.contains(.command) {
                // The formula's own Command keys (a row's inline / display
                // switch ⌘⇧↩ and other Return chords, undo ⌘Z, redo ⌘⇧Z,
                // paste ⌘V);
                // the others are the application's.
                guard Self.formulaCommandKeys.contains(event.keyCode) else { return false }
                AppLink.shared.send(KeyMapping.message(for: event))
                return true
            }
            // Control and Option chords are the formula's shortcuts (output
            // format, OCR, the editor's own): never pinyin's.
            if !event.modifierFlags.intersection([.control, .option]).isEmpty {
                AppLink.shared.send(KeyMapping.message(for: event))
                return true
            }
            // In a formula, as in the main window: in Chinese the keys spell
            // pinyin and the chosen characters become the formula's text; in
            // English (Shift) and for what RIME does not take (digits,
            // operators, editing keys) they are the formula's.
            if pinyinKey(event, client: client) { return true }
            AppLink.shared.send(KeyMapping.message(for: event))
            return true
        }
        return pinyinKey(event, client: client)
    }

    // MARK: Pinyin and English (RIME)

    private func ensureRimeSession() -> RimeSessionId {
        if rimeSession == 0 {
            rimeSession = RimeEngine.shared.createSession()
            // A formula that opened before RIME was ready (the first one after the input method
            // started, 2026-11-09) starts in English as every formula does: it typed pinyin.
            if rimeSession != 0, composing { setFormulaInputMode(on: true) }
        }
        return rimeSession
    }

    private func pinyinKey(_ event: NSEvent, client: IMKTextInput) -> Bool {
        let session = ensureRimeSession()
        guard session != 0, let keysym = RimeKeys.keysym(of: event) else { return false }
        // Command shortcuts belong to the application (a composition is
        // finished first, as with the system's Pinyin).
        if event.modifierFlags.contains(.command) {
            finishPinyin(client)
            return false
        }
        let handled = RimeEngine.shared.process(session, keysym: keysym, mask: RimeKeys.modifiers(event.modifierFlags))
        if composing {
            if handled { showFormulaPinyin(client) }
        } else {
            showPinyin(client)
        }
        return handled
    }

    /// Shift alone switches between Chinese and English (RIME's
    /// ascii_composer sees the press and the release).
    private func pinyinModifiers(_ event: NSEvent, client: IMKTextInput) -> Bool {
        let keysym: Int32
        let isDown: Bool
        switch event.keyCode {
        case 56: keysym = RimeKeys.shiftLeft; isDown = event.modifierFlags.contains(.shift)
        case 60: keysym = RimeKeys.shiftRight; isDown = event.modifierFlags.contains(.shift)
        default: return false
        }
        let session = ensureRimeSession()
        guard session != 0 else { return false }
        let wasAscii = RimeEngine.shared.isAsciiMode(session)
        var mask = RimeKeys.modifiers(event.modifierFlags)
        if !isDown { mask |= RimeKeys.releaseMask }
        let handled = RimeEngine.shared.process(session, keysym: keysym, mask: mask)
        if composing { showFormulaPinyin(client) } else { showPinyin(client) }
        let isAscii = RimeEngine.shared.isAsciiMode(session)
        if isAscii != wasAscii {
            ModeIndicator.shared.show(isAscii ? "英" : "中", under: composing ? formulaCaret(client) : caretRect(client))
        }
        return handled
    }

    /// RIME's commit into the field, its spelling as marked text, its
    /// candidates in the candidate window.
    private func showPinyin(_ client: IMKTextInput) {
        let engine = RimeEngine.shared
        if let commit = engine.takeCommit(rimeSession), !commit.isEmpty {
            client.insertText(commit, replacementRange: Self.noRange)
            marked = ""
            lastPinyinCommit = Date()
        }
        let state = engine.state(rimeSession)
        if state.isComposing {
            if marked.isEmpty { pinyinStart = caretRect(client) }
            marked = state.preedit
            let length = (state.preedit as NSString).length
            client.setMarkedText(
                NSAttributedString(string: state.preedit, attributes: markAttributes(length)),
                selectionRange: NSRange(location: min(state.cursor, length), length: 0),
                replacementRange: Self.noRange
            )
            placeCandidates(state, client: client)
        } else {
            if !marked.isEmpty { clearMarked(client) }
            CandidateWindow.shared.hide()
        }
    }

    /// The candidates under the start of what is being spelled, as the
    /// system's input methods place them. Chrome and Electron report the
    /// marked text's place a moment late (meanwhile a stale or a bogus one),
    /// so the window starts under the caret as it was before the spelling
    /// and follows the marked text once the field reports it on screen.
    private func placeCandidates(_ state: RimeEngine.State, client: IMKTextInput) {
        let show = { (line: NSRect) in
            CandidateWindow.shared.show(state.candidates, labels: state.labels, highlighted: state.highlighted, under: line)
        }
        let fallback = isOnScreen(pinyinStart) ? pinyinStart : caretRect(client)
        if let rect = markedRect(client), isOnScreen(rect), abs(rect.minY - fallback.minY) < max(rect.height * 3, 60) {
            show(NSRect(x: rect.minX, y: rect.minY, width: 1, height: rect.height))
        } else {
            show(fallback)
        }
        let preedit = state.preedit
        for delay in [0.05, 0.15] {
            DispatchQueue.main.asyncAfter(deadline: .now() + delay) { [weak self] in
                guard let self, !self.composing, self.marked == preedit, let client = self.client(),
                      let rect = self.markedRect(client), self.isOnScreen(rect),
                      abs(rect.minY - fallback.minY) < max(rect.height * 3, 60)
                else { return }
                if delay > 0.1 { self.noteTextHeight(client, rect) }
                show(NSRect(x: rect.minX, y: rect.minY, width: 1, height: rect.height))
            }
        }
    }

    private func isOnScreen(_ rect: NSRect) -> Bool {
        rect.height > 0 && rect.height < 400
            && NSScreen.screens.contains(where: { $0.frame.insetBy(dx: -2, dy: -2).intersects(rect) })
    }

    /// Pinyin in the formula's text: RIME's spelling is the editor's
    /// composition (shown in the formula), its commit ends the composition
    /// with the chosen characters, its candidates go under the formula's caret.
    private func showFormulaPinyin(_ client: IMKTextInput) {
        let engine = RimeEngine.shared
        if let commit = engine.takeCommit(rimeSession), !commit.isEmpty {
            AppLink.shared.send(["t": "compose", "phase": "end", "text": commit])
            formulaPinyin = false
        }
        let state = engine.state(rimeSession)
        if state.isComposing {
            AppLink.shared.send(["t": "compose", "phase": "update", "text": state.preedit])
            formulaPinyin = true
            CandidateWindow.shared.show(state.candidates, labels: state.labels, highlighted: state.highlighted, under: formulaCaret(client))
        } else {
            if formulaPinyin { AppLink.shared.send(["t": "compose", "phase": "end", "text": ""]) }
            formulaPinyin = false
            CandidateWindow.shared.hide()
        }
    }

    /// The formula's caret's line on the screen (AppKit coordinates): the
    /// line of the formula's start moved to the caret (which may be on a
    /// later line of the formula).
    private func formulaCaret(_ client: IMKTextInput) -> NSRect {
        var line = heldRect.height > 0 ? heldRect : caretRect(client)
        line.origin.x += formulaCaretOffset
        line.origin.y -= formulaCaretY
        line.size.width = 1
        return line
    }

    /// Chinese or English before the formula opened, to come back to.
    private var asciiBeforeFormula: Bool?

    private func setFormulaInputMode(on: Bool) {
        let session = ensureRimeSession()
        guard session != 0 else { return }
        if on {
            if asciiBeforeFormula == nil { asciiBeforeFormula = RimeEngine.shared.isAsciiMode(session) }
            RimeEngine.shared.setAsciiMode(session, true)
        } else if let ascii = asciiBeforeFormula {
            RimeEngine.shared.setAsciiMode(session, ascii)
            asciiBeforeFormula = nil
        }
    }

    /// Drops pinyin being spelled in the formula (the formula closed).
    private func dropFormulaPinyin() {
        guard formulaPinyin || RimeEngine.shared.state(rimeSession).isComposing else { return }
        RimeEngine.shared.clear(rimeSession)
        if formulaPinyin { AppLink.shared.send(["t": "compose", "phase": "end", "text": ""]) }
        formulaPinyin = false
        CandidateWindow.shared.hide()
    }

    /// What is being spelled goes into the field as it is (focus moved, the
    /// input source changed, a formula opens).
    private func finishPinyin(_ client: IMKTextInput) {
        guard rimeSession != 0 else { return }
        let state = RimeEngine.shared.state(rimeSession)
        guard state.isComposing else { return }
        if let text = RimeEngine.shared.commitComposition(rimeSession), !text.isEmpty {
            client.insertText(text, replacementRange: Self.noRange)
            lastPinyinCommit = Date()
            InputLog.write("pinyin finished length=\((text as NSString).length)")
        } else {
            clearMarked(client)
        }
        marked = ""
        CandidateWindow.shared.hide()
    }

    /// The field ends the composition itself (a click elsewhere in it): it
    /// gets the formula's text, not the anchor.
    override func commitComposition(_ sender: Any!) {
        let client = sender as? IMKTextInput
        InputLog.write("commitComposition composing=\(composing) text=\(formulaText.count)")
        if !composing, let client {
            finishPinyin(client)
            return
        }
        if composing {
            if let client { endedByField(client) }
        }
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
                client.insertText(text, replacementRange: Self.noRange)
            } else if !marked.isEmpty {
                clearMarked(client)
            }
        }
        marked = ""
        formulaText = ""
        holdGeneration += 1
        space = FormulaSpace.Layout(text: "", starts: [], counts: [])
        setFormulaInputMode(on: false)
        composing = false
        AppLink.shared.send(["t": "forced"])
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

    static func receive(_ message: [String: Any]) {
        guard let type = message["t"] as? String else { return }
        guard let controller = active, let client = controller.client() else {
            InputLog.write("message \(type) without an active field")
            return
        }
        switch type {
        case "disconnected":
            if controller.composing {
                controller.clearMarked(client)
                controller.setFormulaInputMode(on: false)
                controller.composing = false
            }
        case "session":
            // VisualTeX opened (the shortcut) or closed a formula for this field;
            // what was being spelled is typed first.
            let on = message["on"] as? Bool ?? false
            if on { controller.finishPinyin(client) } else { controller.dropFormulaPinyin() }
            // A formula starts in English (its letters are variables); Shift
            // switches to Chinese for text. The input mode comes back after.
            controller.setFormulaInputMode(on: on)
            controller.composing = on
            controller.formulaTextMode = false
            controller.formulaStart = controller.composing ? controller.caretRect(client) : .zero
            controller.heldRect = .zero
            controller.holdGeneration += 1
            controller.pieceWidths = []
            controller.lastPlaces = []
            controller.lastEnd = nil
            controller.calibrations = 0
            controller.rekicked = false
            controller.holdDeferred = false
            controller.staleHolds = 0
            controller.suspectAnchor = nil
            controller.heldRange = NSRange(location: NSNotFound, length: 0)
            controller.heldText = ""
            controller.lostPolls = 0
            if on { controller.takeOutLeftovers(client) }
            controller.staleKicks = 0
            controller.kickedText = ""
            controller.kicksUseless = false
            controller.anchorTrusted = false
            controller.anchorDistrusted = false
            controller.formulaCaretOffset = 0
            controller.formulaCaretY = 0
            controller.space = FormulaSpace.Layout(text: "", starts: [], counts: [])
            // Pinyin that went into the field a moment ago (finished just above, or chosen right
            // before the shortcut): Chrome and Electron place a composition set at once where the
            // pinyin was spelled, not at the caret after it (the formula stood over the text before
            // it and its space showed after it, 2026-11-09). The space is held once the field has
            // its caret after the pinyin.
            let sincePinyin = Date().timeIntervalSince(controller.lastPinyinCommit)
            let before = controller.composing ? " \(controller.fieldState(client, brief: true))" : ""
            if controller.composing, sincePinyin < 0.06 {
                InputLog.write("pinyin \(Int(max(0, sincePinyin) * 1000)) ms ago: the space is held in a moment")
                controller.holdLater(0.06 - max(0, sincePinyin)) { controller, client in
                    controller.formulaStart = controller.caretRect(client)
                    controller.placeAnchor(client)
                }
            } else if controller.composing {
                controller.placeAnchor(client)
            } else if !controller.marked.isEmpty {
                controller.clearMarked(client)
            }
            InputLog.write("session \(controller.composing ? "on" : "off")\(before)")
        case "mode":
            controller.formulaTextMode = message["text"] as? Bool ?? false
        case "inline":
            controller.formulaText = message["text"] as? String ?? ""
            controller.formulaCaretOffset = CGFloat(message["caretX"] as? Double ?? 0)
            controller.formulaCaretY = CGFloat(message["caretY"] as? Double ?? 0)
            if let pieces = message["pieces"] as? [Double], controller.composing {
                let widths = pieces.map { CGFloat(max(0, min($0, 20000))) }
                if widths != controller.pieceWidths {
                    controller.pieceWidths = widths
                    controller.rekicked = false
                    controller.hold(client)
                }
            }
        case "commit":
            controller.commit(message["text"] as? String ?? "", returnTo: message["returnTo"] as? String, client: client)
        case "cancel":
            controller.cancel(returnTo: message["returnTo"] as? String, client: client)
        case "caret":
            var reply = controller.geometry(client)
            reply["t"] = "geometry"
            AppLink.shared.send(reply)
        default:
            break
        }
    }

    // MARK: Where the formula goes in the line

    /// Puts the formula's space in the field where its caret is, then
    /// reports where the field placed it (see pollAnchor).
    private func placeAnchor(_ client: IMKTextInput) {
        let estimated = estimatedWidths(client)
        fillerWidth = Self.fillerWidths[clientID(client)] ?? estimated.filler
        jointWidth = Self.jointWidths[clientID(client)] ?? estimated.joint
        hold(client)
    }

    /// A filler's width before the field has shown one: in the field's font
    /// when it tells it, else from its line's height (Chrome and Electron
    /// report a default font, not the page's).
    private func estimatedWidths(_ client: IMKTextInput) -> (filler: CGFloat, joint: CGFloat) {
        var line = NSRect.zero
        let attributes = client.attributes(forCharacterIndex: 0, lineHeightRectangle: &line) ?? [:]
        let font = attributes[NSAttributedString.Key.font] as? NSFont
            ?? attributes[NSAttributedString.Key.font.rawValue] as? NSFont
        if let font, line.height > 0, line.height < (font.ascender - font.descender) * 1.25 + 1 {
            return (FormulaSpace.unit(in: font), FormulaSpace.unit(in: font, of: FormulaSpace.joint))
        }
        let height = Self.textHeights[clientID(client)] ?? (line.height > 6 && line.height < 120 ? line.height : 17)
        return (FormulaSpace.unit(atSize: height / 1.19), FormulaSpace.unit(atSize: height / 1.19, of: FormulaSpace.joint))
    }

    /// The field holds the formula's pieces (FormulaSpace) and the place of
    /// each is followed.
    private func hold(_ client: IMKTextInput) {
        guard !holdDeferred else { return }
        let layout = FormulaSpace.layout(widths: pieceWidths, unit: fillerWidth, joint: jointWidth)
        guard layout != space || marked != layout.text else { return }
        if layout.text != kickedText {
            staleKicks = 0
            kickedText = layout.text
        }
        space = layout
        InputLog.write("hold pieces=\(layout.starts.count) fillers=\(layout.counts.reduce(0, +)) filler=\(String(format: "%.2f", Double(fillerWidth))) joint=\(String(format: "%.2f", Double(jointWidth)))")
        lastHold = Date()
        holdGeneration += 1
        let generation = holdGeneration
        let selection = NSRange(location: 0, length: (FormulaSpace.anchor as NSString).length)
        guard Self.lateClients.contains(clientID(client)) else {
            marked = layout.text
            client.setMarkedText(Self.heldString(layout.text), selectionRange: selection, replacementRange: Self.noRange)
            pollAnchor(generation: generation, last: nil, polls: 0)
            return
        }
        // Chrome and Electron learn where a marked text's characters are one set late (and keep
        // what they learnt until the next): the space with a character of no width after it first,
        // the space itself a moment later, whose places are then the ones they answer. (Typing
        // goes on meanwhile: a space held before the moment is up takes its place.)
        let other = layout.text + "\u{200B}"
        marked = other
        client.setMarkedText(Self.heldString(other), selectionRange: selection, replacementRange: Self.noRange)
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.04) { [weak self] in
            guard let self, generation == self.holdGeneration, self.composing, self.space == layout, let client = self.client() else { return }
            self.marked = layout.text
            client.setMarkedText(Self.heldString(layout.text), selectionRange: selection, replacementRange: Self.noRange)
            self.lastHold = Date()
            self.pollAnchor(generation: generation, last: nil, polls: 0)
        }
    }

    /// Chrome and Electron place marked text a frame or so after it is set
    /// and answer with a stale (or bogus) rectangle meanwhile: the pieces are
    /// reported once two reads agree, and followed (more slowly) while the
    /// formula is open, in case the field moves them (scrolling, a late
    /// layout). The first piece's place is the formula's anchor.
    private func pollAnchor(generation: Int, last: [NSRect]?, polls: Int) {
        guard generation == holdGeneration, composing, let client = client() else { return }
        let starts = space.starts.isEmpty ? [0] : space.starts
        if polls == 0, !anchorTrusted { InputLog.write("held \(fieldState(client, brief: true))") }
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
        var stop: String?
        var rects: [NSRect] = []
        if let first = characterRect(client, starts[0]).flatMap({ isPlausible($0) ? $0 : nil }) {
            rects.append(first)
            for start in starts.dropFirst() {
                guard let rect = pieceRect(client, start, line: first) else {
                    stop = "piece at \(start) has no place"
                    break
                }
                // each piece after the one before on its line, or on a lower line: a field that
                // answers one rectangle for every part of the marked text places nothing
                if let before = rects.last {
                    let sameLine = abs(rect.minY - before.minY) < before.height * 0.5
                    if sameLine ? rect.minX <= before.minX : rect.minY >= before.minY {
                        stop = "piece at \(start) at \(NSStringFromRect(rect)) does not follow \(NSStringFromRect(before))"
                        break
                    }
                }
                rects.append(rect)
            }
        }
        let elapsed = Date().timeIntervalSince(started)
        // Chrome sometimes keeps the places of a marked text replaced at once by another (a kept
        // formula held anew as it opens) and answers the old ones: the space is set again, once
        if polls == 12, rects.count < starts.count, !rekicked {
            rekicked = true
            InputLog.write("pieces \(rects.count)/\(starts.count) placed after \(polls) polls: set again")
            setAgain(client)
            return
        }
        if polls == 2 || elapsed > 0.1, let stop { InputLog.write("poll \(polls): \(stop) ms=\(Int(elapsed * 1000))") }
        // the places are told once two reads agree; a first one that is not the space's (staleAnchor)
        // is not
        let settled = !rects.isEmpty && last == rects
        var stale = settled ? staleAnchor(client, rects[0]) : nil
        if let rect = rects.first, let reason = stale {
            if let suspect = suspectAnchor, abs(suspect.minX - rect.minX) <= 1.5, abs(suspect.minY - rect.minY) <= 1.5 {
                // held afresh and placed there again: the field's own doing (a centred or
                // right-aligned line moves left as the space goes in)
                InputLog.write("anchor \(NSStringFromRect(rect)) polls=\(polls) \(reason), placed there twice: taken \(fieldState(client))")
                if abs(rect.minY - formulaStart.minY) < rect.height * 0.5 {
                    Self.anchorOffsets[clientID(client)] = rect.minX - formulaStart.minX
                }
                stale = nil
            } else if staleHolds < 2 {
                staleHolds += 1
                suspectAnchor = rect
                InputLog.write("anchor \(NSStringFromRect(rect)) polls=\(polls) stale (\(reason)): held afresh \(fieldState(client))")
                holdAfresh(client)
                return
            } else if !anchorDistrusted {
                anchorDistrusted = true
                InputLog.write("anchor \(NSStringFromRect(rect)) polls=\(polls) stale (\(reason)): the formula stays at the caret \(fieldState(client))")
                var message = geometry(client)
                message["t"] = "geometry"
                message["caret"] = payload(NSRect(x: formulaStart.minX, y: formulaStart.minY, width: 0, height: formulaStart.height))
                if !NSDictionary(dictionary: message).isEqual(to: lastGeometry) {
                    lastGeometry = message
                    AppLink.shared.send(message)
                }
            }
        }
        if settled, stale == nil, let rect = rects.first {
            if anchorDistrusted { InputLog.write("anchor \(NSStringFromRect(rect)) polls=\(polls): at the caret again") }
            anchorTrusted = true
            anchorDistrusted = false
            if !rect.equalTo(heldRect) {
                noteTextHeight(client, rect)
                heldRect = rect
                var message = geometry(client)
                message["t"] = "geometry"
                message["caret"] = payload(NSRect(x: rect.minX, y: rect.minY, width: 0, height: rect.height))
                InputLog.write("anchor \(NSStringFromRect(rect)) polls=\(polls) \(fieldState(client))")
                if !NSDictionary(dictionary: message).isEqual(to: lastGeometry) {
                    lastGeometry = message
                    AppLink.shared.send(message)
                }
            }
            reportPlaces(rects, end: rects.count == space.starts.count ? endPlace(client, rects) : nil, anchor: characterRect(client, 0))
            // (a field that answers nothing after the last character: Chrome, Electron)
            if (marked as NSString).length >= 2, characterRect(client, (marked as NSString).length) == nil {
                Self.lateClients.insert(clientID(client))
            }
            // places of a space held before, still answered once typing has paused: the space is
            // set again (twice a space at most; a field that answers so all the same is left alone).
            // The widths are read off places that passed (read off an earlier space's, the joint
            // came out a filler wide and the formula ran past its space, 2026-11-09).
            if !kicksUseless, placesChecked != generation, Date().timeIntervalSince(lastHold) > 0.12 {
                placesChecked = generation
                if let sign = stalePlaces(client, rects) {
                    if staleKicks < 2 {
                        staleKicks += 1
                        // set again first (nothing moves in the field); if that changed nothing, held
                        // afresh (the text after the formula moves for a moment)
                        InputLog.write("pieces stale (\(sign)): \(staleKicks == 1 ? "set again" : "held afresh")")
                        if staleKicks == 1 { setAgain(client) } else { holdAfresh(client) }
                        return
                    }
                    kicksUseless = true
                    InputLog.write("pieces stale (\(sign)) after being set again: left as they are")
                } else {
                    placesSound = generation
                }
            }
            if placesSound == generation, calibrate(client, rects) { return }
        }
        // a field slow to answer (its layout not ready) is asked less often
        let delay = polls < 40 && elapsed < 0.05 ? 0.016 : 0.25
        DispatchQueue.main.asyncAfter(deadline: .now() + delay) { [weak self] in
            self?.pollAnchor(generation: generation, last: rects, polls: polls + 1)
        }
    }

    /// The held space as marked text: with the thin underline. (Given no
    /// attribute the field draws a thick one. Its colour is the system's,
    /// whatever colour is given: an underline nobody sees cannot be asked
    /// for, 2026-11-10.)
    private static func heldString(_ text: String) -> NSAttributedString {
        NSAttributedString(string: text, attributes: [.underlineStyle: NSUnderlineStyle.single.rawValue])
    }

    /// The space set again so that Chrome places it anew: as hold does for
    /// a field known to learn places late (another text first, the space and
    /// a character of no width, so the text after the formula stays where it
    /// is meanwhile; the same text again changes nothing for Chrome).
    private func setAgain(_ client: IMKTextInput) {
        Self.lateClients.insert(clientID(client))
        marked = ""
        hold(client)
    }

    /// Chrome answers a character's place from what its page told it of the
    /// marked text, except the character where it believes the caret is: for
    /// that one it answers the caret's place, and what it knows of the caret
    /// can be several spaces old (typing holds one after another). So one
    /// character of the held space, often a piece's first (where the space
    /// held before ended), gets a place that is not its own, for as long as
    /// the formula stays as it is: a piece drawn off its place, the cover
    /// short of the space's end (2026-11-09). The sign: a filler is as wide
    /// here as there; a piece whose first filler is not as wide as the
    /// others', or that starts where the joint before it does, has a place
    /// that is not its own.
    private func stalePlaces(_ client: IMKTextInput, _ rects: [NSRect]) -> String? {
        guard rects.count == space.starts.count else {
            return "piece \(rects.count) of \(space.starts.count) has no place after the one before it"
        }
        let length = (space.text as NSString).length
        var widths: [(piece: Int, width: CGFloat)] = []
        // (Chrome answers nothing after the last character; a field that does is not asked why it
        // answers nothing for another)
        let chromeLike = characterRect(client, length) == nil
        // Chrome knows the places of an earlier, shorter space only
        if chromeLike, length >= 3, characterRect(client, length - 1) == nil { return "no place for the last character" }
        for (index, start) in space.starts.enumerated() where start + 1 < length {
            // (the character after a piece's first filler, on its line: another filler, or the joint)
            guard let next = characterRect(client, start + 1) else {
                if chromeLike { return "no place for the character after piece \(index)'s first" }
                continue
            }
            guard abs(next.minY - rects[index].minY) < 0.5 else { continue }
            widths.append((index, next.minX - rects[index].minX))
        }
        // a piece that starts where the joint before it does (a word space has a width): the last
        // piece when one filler holds it, which has no character after its first filler to tell
        for (index, start) in space.starts.enumerated() where index > 0 && start >= 2 {
            guard let joint = characterRect(client, start - 1) else {
                if chromeLike { return "no place for the joint before piece \(index)" }
                continue
            }
            guard abs(joint.minY - rects[index].minY) < 0.5 else { continue }
            // (the caret's place Chrome answers may be a point off the character's)
            if rects[index].minX - joint.minX <= 1.5 { return "piece \(index) starts at the joint before it" }
        }
        guard widths.count >= 2 else { return nil }
        let middle = widths.map(\.width).sorted()[widths.count / 2]
        guard let odd = widths.first(where: { abs($0.width - middle) > max(2, abs(middle) * 0.25) }) else { return nil }
        return "piece \(odd.piece)'s first filler \(Int(odd.width.rounded())) pt wide, the others \(Int(middle.rounded()))"
    }

    /// Why the first piece's place is not the held space's, nil when it is.
    /// Chrome and Electron answer the places of marked text from what their
    /// page told them last, and after a composition that went in right before
    /// the formula opened (pinyin) that is still the pinyin's: a character of
    /// the held space gets the place of one in the spelling, or the
    /// spelling's end, for as long as the space is only set again. The
    /// formula stood over the text before it, its cover with it, and the held
    /// space showed (underlined) after it (2026-11-09). The sign, before the
    /// field has placed this formula once: the first piece is not right after
    /// where the caret was (the anchor's width past it), on its line, or is
    /// on a line above. (Held afresh, a place that is the field's own comes
    /// back the same and is taken, pollAnchor: a centred or right-aligned
    /// line moves as the space goes in; that application's next formulas
    /// start with it.)
    private func staleAnchor(_ client: IMKTextInput, _ rect: NSRect) -> String? {
        guard isOnScreen(formulaStart) else { return nil }
        if anchorTrusted, !anchorDistrusted { return nil }
        let height = max(rect.height, formulaStart.height)
        // (AppKit coordinates: up is more)
        if rect.minY - formulaStart.minY > height * 0.5 { return "above the caret's line" }
        // a line below: the space did not fit on the caret's
        guard abs(rect.minY - formulaStart.minY) < height * 0.5 else { return nil }
        let offset = rect.minX - formulaStart.minX
        if offset >= -0.5, offset <= max(4, height * 0.2) { return nil }
        if let known = Self.anchorOffsets[clientID(client)], abs(known - offset) <= 1.5 { return nil }
        return "\(Int(offset.rounded())) pt from the caret"
    }

    /// The field drops the composition, the space is held anew once it has
    /// its caret back (a composition only set again keeps a stale place).
    private func holdAfresh(_ client: IMKTextInput) {
        clearMarked(client)
        space = FormulaSpace.Layout(text: "", starts: [], counts: [])
        holdGeneration += 1
        holdLater(0.06) { controller, client in
            // where the field has its caret now is where the space goes
            let caret = controller.caretRect(client)
            if controller.isOnScreen(caret) { controller.formulaStart = caret }
            controller.hold(client)
        }
    }

    /// Holds the space in a moment; nothing else holds it meanwhile (the
    /// formula's pieces that come are held then).
    private func holdLater(_ delay: TimeInterval, _ hold: @escaping (VTInputController, IMKTextInput) -> Void) {
        holdDeferred = true
        let generation = holdGeneration
        DispatchQueue.main.asyncAfter(deadline: .now() + delay) { [weak self] in
            guard let self, self.holdDeferred, generation == self.holdGeneration else { return }
            self.holdDeferred = false
            guard self.composing, let client = self.client() else { return }
            hold(self, client)
        }
    }

    /// What the field tells of its marked text and selection, for the log:
    /// their ranges, the places of the marked text's first two characters,
    /// its last one and its end, the first character's rectangle as asked by
    /// its place in the document, and where the formula opened.
    private func fieldState(_ client: IMKTextInput, brief: Bool = false) -> String {
        let markedRange = client.markedRange()
        let selected = client.selectedRange()
        let length = (marked as NSString).length
        let range = { (range: NSRange) in range.location == NSNotFound ? "-" : "\(range.location)+\(range.length)" }
        let place = { (index: Int) -> String in
            guard index >= 0 else { return "-" }
            var rect = NSRect.zero
            _ = client.attributes(forCharacterIndex: index, lineHeightRectangle: &rect)
            return "\(Int(rect.minX)),\(Int(rect.minY))"
        }
        if brief {
            return "marked=\(range(markedRange)) selected=\(range(selected)) at0=\(place(0)) start=\(Int(formulaStart.minX)),\(Int(formulaStart.minY))"
        }
        var first = "-"
        if markedRange.location != NSNotFound, markedRange.length > 0 {
            var actual = Self.noRange
            let rect = client.firstRect(forCharacterRange: NSRange(location: markedRange.location, length: 1), actualRange: &actual)
            first = "\(Int(rect.minX)),\(Int(rect.minY)),\(Int(rect.width))x\(Int(rect.height))@\(range(actual))"
        }
        return "marked=\(range(markedRange)) selected=\(range(selected)) at0=\(place(0)) at1=\(place(1)) last=\(place(length - 1)) end=\(place(length)) first=\(first) start=\(Int(formulaStart.minX)),\(Int(formulaStart.minY))"
    }

    /// Where the field put each piece, from the first one's start on its
    /// line (points, y down; the lines' bottoms, so a line of another height
    /// keeps the formula on its baseline as nearly as the field tells).
    private func reportPlaces(_ rects: [NSRect], end: [Double]?, anchor: NSRect?) {
        guard let first = rects.first else { return }
        let places = rects.map { [Double(($0.minX - first.minX).rounded(toPlaces: 2)), Double((first.minY - $0.minY).rounded(toPlaces: 2))] }
        guard places != lastPlaces || end != lastEnd else { return }
        lastPlaces = places
        lastEnd = end
        let lines = Set(places.map { $0[1] }).count
        InputLog.write("pieces \(places.count)/\(space.starts.count) lines=\(lines) \(places.prefix(6).map { "\(Int($0[0])),\(Int($0[1]))" }.joined(separator: " ")) end=\(end.map { "\(Int($0[0])),\(Int($0[1]))" } ?? "-")")
        var message: [String: Any] = ["t": "pieces", "places": places, "count": space.starts.count]
        if let end { message["end"] = end }
        // the selected anchor before the formula (the field may highlight it)
        if let anchor, abs(anchor.minY - first.minY) < first.height * 0.5 {
            message["start"] = [Double((anchor.minX - first.minX).rounded(toPlaces: 2)), 0]
        }
        AppLink.shared.send(message)
    }

    /// The filler's and the joint's widths as the field draws them: the
    /// filler over the longest piece on one line (the field may round each
    /// character's place to a pixel, so a long run reads it best), the joint
    /// between a piece's fillers and the next piece on its line. Widths other
    /// than the ones the pieces were held with hold them again. True when they
    /// did.
    private func calibrate(_ client: IMKTextInput, _ rects: [NSRect]) -> Bool {
        // A space held a moment ago is not drawn everywhere yet: typing holds a new one every few
        // milliseconds and Chrome answers some places from the one it replaced. Widths read from
        // those flipped the joint between 0, 4 and 10 points and left the formula wider than its
        // space (2026-11-09); the polls that follow read them once the field has drawn the space.
        guard Date().timeIntervalSince(lastHold) > 0.12 else { return false }
        var filler: (count: Int, width: CGFloat)?
        // the fillers of all the pieces read whole on their lines: together a run long enough that a
        // field's pixel rounding is a small part of it (a formula of short pieces holds two or three
        // fillers each, so a single piece's never was; Safari draws the filler 10.3 points wide where
        // its font says 12.6, and the first line ran past the field's edge, 2026-11-09)
        var fillerCount = 0
        var fillerRun: CGFloat = 0
        var joints: [CGFloat] = []
        for (index, start) in space.starts.enumerated() where index < rects.count {
            let count = space.counts[index]
            guard let end = characterRect(client, start + count),
                  abs(end.minY - rects[index].minY) < rects[index].height * 0.5,
                  end.minX > rects[index].minX
            else { continue }
            fillerCount += count
            fillerRun += end.minX - rects[index].minX
            if fillerCount >= 6 { filler = (fillerCount, fillerRun / CGFloat(fillerCount)) }
            if index + 1 < rects.count, abs(rects[index + 1].minY - end.minY) < end.height * 0.5,
               rects[index + 1].minX > end.minX {
                joints.append(rects[index + 1].minX - end.minX)
            }
        }
        // the joint: the middle one of those read on the lines, once two of them agree (a single
        // reading flipped between 4 and 10 points in Chrome)
        var joint: CGFloat?
        let middle = joints.isEmpty ? nil : joints.sorted()[joints.count / 2]
        if let middle, middle >= 1, middle < 40, joints.filter({ abs($0 - middle) <= 1.5 }).count >= 2 {
            joint = middle
        }
        if let width = filler?.width, !(width > 0.5 && width < 40) { filler = nil }
        let id = clientID(client)
        // a pixel over the run is the field's rounding, not another width
        let fillerChanged = filler.map { abs($0.width - fillerWidth) * CGFloat($0.count) > 1.5 } ?? false
        let jointChanged = joint.map { abs($0 - jointWidth) > 1.5 } ?? false
        guard fillerChanged || jointChanged else {
            // the widths the space was held with: this application's next formulas start with them
            if let filler { Self.fillerWidths[id] = filler.width }
            if let joint { Self.jointWidths[id] = joint }
            return false
        }
        // past three, the widths stay those the space is held with (never a space held with other
        // widths than the ones the pieces are placed with)
        guard calibrations < 3 else { return false }
        calibrations += 1
        if fillerChanged, let filler {
            InputLog.write("filler \(String(format: "%.2f", Double(fillerWidth))) -> \(String(format: "%.2f", Double(filler.width))) from \(filler.count)")
            fillerWidth = filler.width
            Self.fillerWidths[id] = filler.width
        }
        if jointChanged, let joint {
            InputLog.write("joint \(String(format: "%.2f", Double(jointWidth))) -> \(String(format: "%.2f", Double(joint))) from \(joints.count)")
            jointWidth = joint
            Self.jointWidths[id] = joint
        }
        // (the same space with the new widths is not held again: its places go on being followed)
        let generation = holdGeneration
        hold(client)
        return holdGeneration != generation
    }

    /// Where the field ends the held space, from the first piece's start
    /// (the end of the formula's last line), for VisualTeX's cover: the
    /// place after the last character, else the last one's and its width,
    /// else the last piece's start and its fillers, marked as estimated.
    private func endPlace(_ client: IMKTextInput, _ rects: [NSRect]) -> [Double]? {
        guard let first = rects.first, let lastStart = rects.last, let count = space.counts.last else { return nil }
        let length = (space.text as NSString).length
        var end = characterRect(client, length)
        // Chrome answers nothing after the last character: the last one's place and its width (the
        // advance from the character before it, a filler too)
        if end == nil, let last = characterRect(client, length - 1) {
            var advance = fillerWidth
            if count >= 2, let before = characterRect(client, length - 2),
               abs(before.minY - last.minY) < last.height * 0.5, last.minX > before.minX {
                advance = last.minX - before.minX
            }
            end = NSRect(x: last.minX + advance, y: last.minY, width: 0, height: last.height)
        }
        if let end {
            return [Double((end.minX - first.minX).rounded(toPlaces: 2)), Double((first.minY - end.minY).rounded(toPlaces: 2))]
        }
        // estimated (a third component): VisualTeX covers no further than its formula's end
        let x = lastStart.minX + CGFloat(count) * fillerWidth
        return [Double((x - first.minX).rounded(toPlaces: 2)), Double((first.minY - lastStart.minY).rounded(toPlaces: 2)), 1]
    }

    /// On a screen and on (or near) the line where the formula opened.
    private func isPlausible(_ rect: NSRect) -> Bool {
        guard isOnScreen(rect) else { return false }
        guard isOnScreen(formulaStart) else { return true }
        return abs(rect.minY - formulaStart.minY) <= max(rect.height * 4, 80)
    }

    private func commit(_ text: String, returnTo: String?, client: IMKTextInput) {
        dropFormulaPinyin()
        setFormulaInputMode(on: false)
        formulaTextMode = false
        client.insertText(text, replacementRange: Self.noRange)
        InputLog.write("commit length=\((text as NSString).length) return=\(returnTo ?? "-")")
        marked = ""
        formulaText = ""
        composing = false
        space = FormulaSpace.Layout(text: "", starts: [], counts: [])
        hand(client, to: returnTo)
    }

    private func cancel(returnTo: String?, client: IMKTextInput) {
        dropFormulaPinyin()
        setFormulaInputMode(on: false)
        formulaTextMode = false
        clearMarked(client)
        formulaText = ""
        composing = false
        space = FormulaSpace.Layout(text: "", starts: [], counts: [])
        InputLog.write("cancel return=\(returnTo ?? "-")")
        hand(client, to: returnTo)
    }

    /// The field's own text input system switches input source: at once,
    /// unlike a selection made from another process.
    private func hand(_ client: IMKTextInput, to source: String?) {
        guard let source, !source.isEmpty, source != Bundle.main.bundleIdentifier else { return }
        client.selectMode(source)
    }

    private func clearMarked(_ client: IMKTextInput) {
        marked = ""
        client.setMarkedText("", selectionRange: NSRange(location: 0, length: 0), replacementRange: Self.noRange)
    }

    func markAttributes(_ length: Int) -> [NSAttributedString.Key: Any] {
        var attributes: [NSAttributedString.Key: Any] = [:]
        if let marks = mark(forStyle: kTSMHiliteSelectedRawText, at: NSRange(location: 0, length: length)) {
            for (key, value) in marks {
                if let key = key as? NSAttributedString.Key {
                    attributes[key] = value
                } else if let key = key as? String {
                    attributes[NSAttributedString.Key(key)] = value
                }
            }
        }
        if attributes.isEmpty { attributes[.underlineStyle] = NSUnderlineStyle.single.rawValue }
        return attributes
    }

    /// A character of the marked text (by its index in it) as the field
    /// places it: a caret-like rectangle at its left edge, the line's height
    /// (AppKit screen coordinates). nil when the field answers with nothing
    /// on a screen. Chrome answers this, but not firstRect (garbage there).
    private func characterRect(_ client: IMKTextInput, _ index: Int) -> NSRect? {
        var rect = NSRect.zero
        _ = client.attributes(forCharacterIndex: index, lineHeightRectangle: &rect)
        return isOnScreen(rect) ? rect : nil
    }

    /// Where a piece starts (its first character, by its index in the marked
    /// text). TextKit 2 fields (TextEdit, Notes) answer nothing for a place at
    /// the start of a line they wrapped: there the character's own rectangle
    /// is asked for, on a line as tall as the first piece's.
    private func pieceRect(_ client: IMKTextInput, _ index: Int, line: NSRect) -> NSRect? {
        if let rect = characterRect(client, index) { return rect }
        let marked = client.markedRange()
        guard marked.location != NSNotFound, index < marked.length else { return nil }
        let glyph = client.firstRect(forCharacterRange: NSRange(location: marked.location + index, length: 1), actualRange: nil)
        guard isOnScreen(glyph) else { return nil }
        // the line's bottom, as the line rectangles are placed
        return NSRect(x: glyph.minX, y: glyph.minY - (line.height - glyph.height) / 2, width: 1, height: line.height)
    }

    /// The marked text's rectangle: from its first character's left edge to
    /// the edge after its last character, on the first character's line.
    private func markedRect(_ client: IMKTextInput) -> NSRect? {
        let length = (marked as NSString).length
        guard length > 0, let first = characterRect(client, 0) else { return nil }
        var right = first.minX
        if let end = characterRect(client, length), abs(end.minY - first.minY) < first.height * 0.5, end.minX > first.minX {
            right = end.minX
        } else if let last = characterRect(client, length - 1), abs(last.minY - first.minY) < first.height * 0.5 {
            right = max(last.maxX, last.minX + (length > 1 ? (last.minX - first.minX) / CGFloat(length - 1) : 0))
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

    private func clientID(_ sender: Any?) -> String {
        (sender as? IMKTextInput)?.bundleIdentifier() ?? "?"
    }
}

private extension CGFloat {
    func rounded(toPlaces places: Int) -> CGFloat {
        let scale = pow(10, CGFloat(places))
        return (self * scale).rounded() / scale
    }
}
