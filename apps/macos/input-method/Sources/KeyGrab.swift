import Carbon
import Cocoa

/// The formula's keys. A palette input method is given no keys, and the field's own input source
/// (Pinyin, say) must not get them while a formula is open: for that time every key the formula
/// uses is registered as a hot key, which macOS hands over before the frontmost application (no
/// permission is needed, as for VisualTeX's own shortcut). Each press goes to VisualTeX as the key
/// event its editor expects (KeyMapping); a key held repeats as the system's keyboard settings say.
/// Command shortcuts stay the application's, but the formula's own (Return chords, undo, redo,
/// paste); Control with Space or an arrow stays the system's (input sources, Mission Control).
final class KeyGrab {
    static let shared = KeyGrab()

    private static let signature: OSType = 0x5654_4B47 // "VTKG"
    private static let none = 0
    private static let shift = Int(shiftKey)
    private static let option = Int(optionKey)
    private static let control = Int(controlKey)
    private static let command = Int(cmdKey)

    /// Keys that type a character (ANSI positions), the keys that edit, and the keypad.
    private static let typing: [Int] = Array(0...9) + Array(11...35) + Array(37...47) + [50]
    private static let editing: [Int] = [36, 48, 49, 51, 53, 76, 115, 116, 117, 119, 121, 123, 124, 125, 126]
    private static let keypad: [Int] = [65, 67, 69, 71, 75, 78, 81, 82, 83, 84, 85, 86, 87, 88, 89, 91, 92]
    /// Return, keypad Enter, Z and V: with Command they act on the formula.
    private static let formulaCommandKeys: [Int] = [36, 76, 6, 9]
    private static let arrows: Set<Int> = [123, 124, 125, 126]

    private var handler: EventHandlerRef?
    /// The registered keys: (key code, modifiers) by hot key id.
    private var keys: [UInt32: (code: Int, modifiers: Int)] = [:]
    private var references: [EventHotKeyRef] = []
    /// Chords still to register (the plain and the Shift ones are registered at once, the others
    /// right after: a formula's first key is never waited for).
    private var pending: [(code: Int, modifiers: Int)] = []
    private var generation = 0
    private var repeating: (id: UInt32, timer: Timer)?

    var isOn: Bool { !references.isEmpty || !pending.isEmpty }

    /// Takes the formula's keys; `skipping`: chords ([key code, Carbon modifiers]) VisualTeX has for
    /// itself (its shortcuts).
    func start(skipping: [[Int]]) {
        stop()
        installHandler()
        generation += 1
        let skipped = Set(skipping.filter { $0.count == 2 }.map { "\($0[0])/\($0[1])" })
        var first: [(Int, Int)] = []
        var later: [(Int, Int)] = []
        let chords = [Self.option, Self.option | Self.shift, Self.control, Self.control | Self.shift,
                      Self.control | Self.option, Self.control | Self.option | Self.shift]
        for code in Self.typing + Self.editing {
            first.append((code, Self.none))
            first.append((code, Self.shift))
            for modifiers in chords {
                // Control-Space and Control-arrows are the system's
                if modifiers & Self.control != 0, code == 49 || Self.arrows.contains(code) { continue }
                later.append((code, modifiers))
            }
        }
        for code in Self.keypad {
            first.append((code, Self.none))
            first.append((code, Self.shift))
        }
        for code in Self.formulaCommandKeys {
            first.append((code, Self.command))
            first.append((code, Self.command | Self.shift))
        }
        let started = Date()
        var refused = 0
        for (code, modifiers) in first where !skipped.contains("\(code)/\(modifiers)") {
            if !register(code, modifiers) { refused += 1 }
        }
        pending = later.filter { !skipped.contains("\($0.0)/\($0.1)") }.map { (code: $0.0, modifiers: $0.1) }
        InputLog.write("keys taken: \(references.count) in \(Int(Date().timeIntervalSince(started) * 1000)) ms, refused \(refused), \(pending.count) chords to come")
        registerPending(generation)
    }

    /// Gives the keys back.
    func stop() {
        generation += 1
        pending = []
        repeating?.timer.invalidate()
        repeating = nil
        guard !references.isEmpty else { return }
        for reference in references { UnregisterEventHotKey(reference) }
        references = []
        keys = [:]
        InputLog.write("keys given back")
    }

    private func registerPending(_ generation: Int) {
        guard generation == self.generation, !pending.isEmpty else { return }
        for chord in pending.prefix(60) { _ = register(chord.code, chord.modifiers) }
        pending.removeFirst(min(60, pending.count))
        DispatchQueue.main.async { [weak self] in self?.registerPending(generation) }
    }

    private func register(_ code: Int, _ modifiers: Int) -> Bool {
        let id = UInt32(keys.count + 1)
        var reference: EventHotKeyRef?
        let status = RegisterEventHotKey(UInt32(code), UInt32(modifiers), EventHotKeyID(signature: Self.signature, id: id),
                                         GetApplicationEventTarget(), 0, &reference)
        guard status == noErr, let reference else { return false }
        references.append(reference)
        keys[id] = (code, modifiers)
        return true
    }

    private func installHandler() {
        guard handler == nil else { return }
        var spec = [EventTypeSpec(eventClass: OSType(kEventClassKeyboard), eventKind: UInt32(kEventHotKeyPressed)),
                    EventTypeSpec(eventClass: OSType(kEventClassKeyboard), eventKind: UInt32(kEventHotKeyReleased))]
        InstallEventHandler(GetApplicationEventTarget(), { _, event, _ in
            guard let event else { return OSStatus(eventNotHandledErr) }
            var id = EventHotKeyID()
            let status = GetEventParameter(event, EventParamName(kEventParamDirectObject), EventParamType(typeEventHotKeyID), nil,
                                           MemoryLayout<EventHotKeyID>.size, nil, &id)
            guard status == noErr, id.signature == KeyGrab.signature else { return OSStatus(eventNotHandledErr) }
            KeyGrab.shared.hotKey(id.id, pressed: GetEventKind(event) == UInt32(kEventHotKeyPressed))
            return noErr
        }, 2, &spec, nil, &handler)
    }

    private func hotKey(_ id: UInt32, pressed: Bool) {
        guard let key = keys[id] else { return }
        if !pressed {
            if repeating?.id == id {
                repeating?.timer.invalidate()
                repeating = nil
            }
            return
        }
        repeating?.timer.invalidate()
        repeating = nil
        send(key, isRepeat: false)
        // A key held repeats (not Return, Escape or a Command chord: those end or change the formula once).
        guard key.modifiers & Self.command == 0, ![36, 53, 76].contains(key.code) else { return }
        let delay = max(0.15, NSEvent.keyRepeatDelay)
        let interval = max(0.02, NSEvent.keyRepeatInterval)
        let generation = self.generation
        let timer = Timer(fire: Date().addingTimeInterval(delay), interval: interval, repeats: true) { [weak self] timer in
            guard let self, generation == self.generation, self.repeating?.id == id else {
                timer.invalidate()
                return
            }
            self.send(key, isRepeat: true)
        }
        RunLoop.main.add(timer, forMode: .common)
        repeating = (id, timer)
    }

    private func send(_ key: (code: Int, modifiers: Int), isRepeat: Bool) {
        // VisualTeX is gone (the link closed between two keys): the keys are nobody's to take.
        guard AppLink.shared.isConnected else {
            stop()
            return
        }
        AppLink.shared.send(KeyMapping.message(
            keyCode: UInt16(key.code),
            shift: key.modifiers & Self.shift != 0,
            option: key.modifiers & Self.option != 0,
            control: key.modifiers & Self.control != 0,
            command: key.modifiers & Self.command != 0,
            isRepeat: isRepeat
        ))
    }
}
