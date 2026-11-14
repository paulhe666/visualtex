import Carbon
import Cocoa
import InputMethodKit

/// One per host text field. VisualTeX is a whole input method: Chinese
/// through pinyin (RIME) and English (Shift switches), with formulas as one
/// more mode. While a formula is open (the shortcut) every key but a Command
/// shortcut goes to VisualTeX, which edits the formula and draws it in the
/// line itself, from where the field's caret was, with its own caret: the
/// field holds the formula's space in its text (FormulaSpace), so it wraps
/// the formula as it wraps its own text and shows nothing of its own. Where
/// that space lies is worked out by SpaceModel from the field's caret before
/// and after it; the field's other answers only confirm or correct it.
/// Return inserts the formula's text with insertText and, when the shortcut
/// opened the formula from another input source, hands the field back to it.
@objc(VTInputController)
final class VTInputController: IMKInputController {
    static let noRange = NSRange(location: NSNotFound, length: 0)

    /// Return, keypad Enter, Z and V by virtual key code: with Command they
    /// act on the formula while it is open (⌘V pastes into it: an image is
    /// recognized, text goes in).
    private static let formulaCommandKeys: Set<UInt16> = [36, 76, 6, 9]

    /// The controller whose field has the keyboard: VisualTeX's messages go to it.
    private static weak var active: VTInputController?
    /// (for HostProbe)
    static var current: VTInputController? { active }
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
    /// Where the formula starts: its first piece's start on its line (AppKit
    /// screen coordinates), zero until the field's caret is known.
    private var heldRect = NSRect.zero
    /// Each space held; a poll for an older one stops.
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
    /// The formula's caret below the formula's first baseline, points (it may
    /// be on a later line).
    private var formulaCaretY: CGFloat = 0
    /// When pinyin last went into the field: the field takes a moment to
    /// place its caret after it (see holdLater).
    private var lastPinyinCommit = Date.distantPast
    /// The space is held a moment later (holdLater): nothing holds it meanwhile.
    private var holdDeferred = false
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
    /// A Shift press that becomes a tap when it is let go soon with nothing pressed or clicked meanwhile: its key,
    /// when, and the input events counted then (shiftChanged).
    private var shiftTap: (keyCode: UInt16, at: TimeInterval, events: UInt32)?
    /// Shift's changes and taps while this field had the keyboard (logged when it ends: a field whose Shift never
    /// arrives shows there).
    private var shiftChanges = 0
    private var shiftTaps = 0

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
        InputLog.write("deactivate client=\(clientID(sender)) composing=\(composing) shift=\(shiftChanges)/\(shiftTaps)")
        shiftChanges = 0
        shiftTaps = 0
        shiftTap = nil
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
            return shiftChanged(event, client: client)
        }
        guard event.type == .keyDown else { return false }
        // a key with Shift down: Shift was part of a chord
        shiftTap = nil
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
        if rimeSession != 0, !RimeEngine.shared.isAlive(rimeSession) {
            InputLog.write("rime session lost: a new one")
            rimeSession = 0
        }
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

    /// The input events so far (keys, modifiers, clicks): a Shift press and its release with nothing else between
    /// are two apart (Squirrel's test).
    private static func inputEventCount() -> UInt32 {
        let types: [CGEventType] = [.flagsChanged, .keyDown, .leftMouseDown, .rightMouseDown, .otherMouseDown]
        return types.reduce(UInt32(0)) { $0 &+ CGEventSource.counterForEventType(.combinedSessionState, eventType: $1) }
    }

    /// Shift alone, let go within half a second with nothing else pressed or clicked meanwhile (no chord, no
    /// Shift-click), switches between Chinese and English, as macOS Pinyin's does. VisualTeX tells the tap itself:
    /// RIME's ascii_composer is left out of it (its own Shift state could keep Shift from switching after a while,
    /// 2026-11-10). With Caps Lock on, letters are English capitals and Shift switches nothing.
    private func shiftChanged(_ event: NSEvent, client: IMKTextInput) -> Bool {
        guard event.keyCode == 56 || event.keyCode == 60 else {
            // another modifier: a chord, not a tap
            if event.keyCode != 57 { shiftTap = nil }
            return false
        }
        shiftChanges += 1
        // this side's Shift (the device's left and right bits; .shift when the device sets neither)
        let raw = event.modifierFlags.rawValue
        let down = raw & 0x6 != 0 ? raw & (event.keyCode == 56 ? 0x2 : 0x4) != 0 : event.modifierFlags.contains(.shift)
        if down {
            // (with Command, Control, Option or Fn, or both Shifts: a chord)
            let chord = !event.modifierFlags.intersection([.command, .control, .option, .function]).isEmpty || raw & 0x6 == 0x6
            if chord {
                shiftTap = nil
            } else {
                shiftTap = (keyCode: event.keyCode, at: event.timestamp, events: Self.inputEventCount())
            }
            return false
        }
        guard let tap = shiftTap, tap.keyCode == event.keyCode else {
            shiftTap = nil
            return false
        }
        shiftTap = nil
        guard event.timestamp - tap.at < 0.5, Self.inputEventCount() &- tap.events <= 1 else { return false }
        if event.modifierFlags.contains(.capsLock) {
            InputLog.write("shift tap with Caps Lock on: no switch")
            return false
        }
        switchChineseEnglish(client)
        return false
    }

    /// Chinese ⇄ English: what is being spelled goes in as the letters typed (as macOS Pinyin's Shift does), the
    /// mode changes and its indicator shows.
    private func switchChineseEnglish(_ client: IMKTextInput) {
        let session = ensureRimeSession()
        guard session != 0 else {
            InputLog.write("shift tap before RIME is ready")
            return
        }
        let english = !RimeEngine.shared.isAsciiMode(session)
        if RimeEngine.shared.state(session).isComposing, let typed = RimeEngine.shared.takeInput(session) {
            if composing {
                AppLink.shared.send(["t": "compose", "phase": "end", "text": typed])
                formulaPinyin = false
            } else if !typed.isEmpty {
                client.insertText(typed, replacementRange: Self.noRange)
                marked = ""
                lastPinyinCommit = Date()
            } else if !marked.isEmpty {
                clearMarked(client)
            }
            CandidateWindow.shared.hide()
        }
        RimeEngine.shared.setAsciiMode(session, english)
        shiftTaps += 1
        ModeIndicator.shared.show(english ? "英" : "中", under: composing ? formulaCaret(client) : caretRect(client))
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
        model = nil
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
            controller.lastCaretLine = nil
            controller.placed = false
            controller.anchorLine = nil
            controller.lastCharacterLine = nil
            controller.reholds = 0
            controller.holdDeferred = false
            controller.heldRange = NSRange(location: NSNotFound, length: 0)
            controller.heldText = ""
            controller.lostPolls = 0
            if on { controller.takeOutLeftovers(client) }
            controller.formulaCaretOffset = 0
            controller.formulaCaretY = 0
            controller.space = FormulaSpace.Layout(text: "", starts: [], counts: [])
            controller.model = nil
            // Pinyin that went into the field a moment ago (finished just above, or chosen right
            // before the shortcut): Chrome and Electron still answer the caret where the pinyin was
            // spelled, not after it. The space is held once the field has its caret after the pinyin.
            let sincePinyin = Date().timeIntervalSince(controller.lastPinyinCommit)
            let before = controller.composing ? " \(controller.fieldState(client))" : ""
            if controller.composing, sincePinyin < 0.06 {
                InputLog.write("pinyin \(Int(max(0, sincePinyin) * 1000)) ms ago: the space is held in a moment")
                controller.holdLater(0.06 - max(0, sincePinyin)) { controller, client in
                    controller.formulaStart = controller.caretRect(client)
                    controller.openSpace(client)
                }
            } else if controller.composing {
                controller.openSpace(client)
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
        guard !holdDeferred, var model else { return }
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
        client.setMarkedText(
            Self.heldString(layout.text),
            selectionRange: NSRange(location: 0, length: (FormulaSpace.anchor as NSString).length),
            replacementRange: Self.noRange
        )
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

    /// The held space as marked text: with the thin underline. (Given no
    /// attribute the field draws a thick one. Its colour is the system's,
    /// whatever colour is given: an underline nobody sees cannot be asked
    /// for, 2026-11-10.)
    private static func heldString(_ text: String) -> NSAttributedString {
        NSAttributedString(string: text, attributes: [.underlineStyle: NSUnderlineStyle.single.rawValue])
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

    private func commit(_ text: String, returnTo: String?, client: IMKTextInput) {
        dropFormulaPinyin()
        setFormulaInputMode(on: false)
        formulaTextMode = false
        notePastEnd(client)
        client.insertText(text, replacementRange: Self.noRange)
        InputLog.write("commit length=\((text as NSString).length) return=\(returnTo ?? "-")")
        marked = ""
        formulaText = ""
        composing = false
        space = FormulaSpace.Layout(text: "", starts: [], counts: [])
        model = nil
        hand(client, to: returnTo)
    }

    private func cancel(returnTo: String?, client: IMKTextInput) {
        dropFormulaPinyin()
        setFormulaInputMode(on: false)
        formulaTextMode = false
        notePastEnd(client)
        clearMarked(client)
        formulaText = ""
        composing = false
        space = FormulaSpace.Layout(text: "", starts: [], counts: [])
        model = nil
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
