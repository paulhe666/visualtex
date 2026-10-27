import Carbon
import Cocoa
import InputMethodKit

/// One per host text field. VisualTeX is a whole input method: Chinese
/// through pinyin (RIME) and English (Shift switches), with formulas as one
/// more mode. While a formula is open (the shortcut) every key but a Command
/// shortcut goes to VisualTeX, which edits the formula and draws it in the
/// line itself, from where the field's caret was, with its own caret: the
/// field shows nothing of its own meanwhile. Return inserts the formula's
/// text with insertText and, when the shortcut opened the formula from
/// another input source, hands the field back to it.
@objc(VTInputController)
final class VTInputController: IMKInputController {
    static let noRange = NSRange(location: NSNotFound, length: 0)
    /// The formula's marked text: one hair space, selected. Being selected
    /// (not a caret) it hides the field's caret, so the formula's caret is
    /// the only one; a hair space is about a point wide, so the field's
    /// composition underline and selection under it are next to nothing, and
    /// the formula is drawn over them. (Chrome draws the underline in the
    /// accent colour, whatever colour the input method gives it; a selected
    /// character of no width, a word joiner, still shows Chrome's caret.)
    private static let anchorText = "\u{200A}"

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
    }

    override func deactivateServer(_ sender: Any!) {
        InputLog.write("deactivate client=\(clientID(sender)) composing=\(composing)")
        let client = sender as? IMKTextInput
        if !marked.isEmpty, let client { clearMarked(client) }
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
        if rimeSession == 0 { rimeSession = RimeEngine.shared.createSession() }
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
            ModeIndicator.shared.show(isAscii ? "英" : "中", under: caretRect(client))
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
            var line = heldRect.height > 0 ? heldRect : caretRect(client)
            line.origin.x += formulaCaretOffset
            line.size.width = 1
            CandidateWindow.shared.show(state.candidates, labels: state.labels, highlighted: state.highlighted, under: line)
        } else {
            if formulaPinyin { AppLink.shared.send(["t": "compose", "phase": "end", "text": ""]) }
            formulaPinyin = false
            CandidateWindow.shared.hide()
        }
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
            if !formulaText.isEmpty {
                client?.insertText(formulaText, replacementRange: Self.noRange)
            } else if !marked.isEmpty, let client {
                clearMarked(client)
            }
            marked = ""
            formulaText = ""
            setFormulaInputMode(on: false)
            composing = false
            AppLink.shared.send(["t": "forced"])
        }
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
            if controller.composing {
                controller.placeAnchor(client)
            } else if !controller.marked.isEmpty {
                controller.clearMarked(client)
            }
            InputLog.write("session \(controller.composing ? "on" : "off")")
        case "mode":
            controller.formulaTextMode = message["text"] as? Bool ?? false
        case "inline":
            controller.formulaText = message["text"] as? String ?? ""
            controller.formulaCaretOffset = CGFloat(message["caretX"] as? Double ?? 0)
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

    /// Puts the formula's anchor in the field where its caret is, then
    /// reports where the field placed it (see pollAnchor).
    private func placeAnchor(_ client: IMKTextInput) {
        marked = Self.anchorText
        let length = (Self.anchorText as NSString).length
        client.setMarkedText(
            NSAttributedString(string: Self.anchorText),
            selectionRange: NSRange(location: 0, length: length),
            replacementRange: Self.noRange
        )
        pollAnchor(generation: holdGeneration, last: nil, polls: 0)
    }

    /// Chrome and Electron place marked text a frame or so after it is set
    /// and answer with a stale (or bogus) rectangle meanwhile: the anchor is
    /// reported once two reads agree, and followed (more slowly) while the
    /// formula is open, in case the field moves it (scrolling, a late layout).
    private func pollAnchor(generation: Int, last: NSRect?, polls: Int) {
        guard generation == holdGeneration, composing, let client = client() else { return }
        let rect = characterRect(client, 0).flatMap { isPlausible($0) ? $0 : nil }
        if let rect, let last, rect.equalTo(last), !rect.equalTo(heldRect) {
            noteTextHeight(client, rect)
            heldRect = rect
            var message = geometry(client)
            message["t"] = "geometry"
            message["caret"] = payload(NSRect(x: rect.minX, y: rect.minY, width: 0, height: rect.height))
            InputLog.write("anchor \(NSStringFromRect(rect)) polls=\(polls)")
            if !NSDictionary(dictionary: message).isEqual(to: lastGeometry) {
                lastGeometry = message
                AppLink.shared.send(message)
            }
        }
        let delay = polls < 40 ? 0.016 : 0.25
        DispatchQueue.main.asyncAfter(deadline: .now() + delay) { [weak self] in
            self?.pollAnchor(generation: generation, last: rect, polls: polls + 1)
        }
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
        hand(client, to: returnTo)
    }

    private func cancel(returnTo: String?, client: IMKTextInput) {
        dropFormulaPinyin()
        setFormulaInputMode(on: false)
        formulaTextMode = false
        clearMarked(client)
        formulaText = ""
        composing = false
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
