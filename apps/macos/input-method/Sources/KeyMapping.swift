import Cocoa

/// A key event as VisualTeX's editor page receives it: the DOM KeyboardEvent
/// fields (`key`, `code`, modifiers) and, for a character key, the text it
/// types.
enum KeyMapping {
    /// macOS virtual key codes (ANSI positions) to DOM `code` values.
    private static let codes: [UInt16: String] = [
        0: "KeyA", 1: "KeyS", 2: "KeyD", 3: "KeyF", 4: "KeyH", 5: "KeyG", 6: "KeyZ", 7: "KeyX",
        8: "KeyC", 9: "KeyV", 10: "IntlBackslash", 11: "KeyB", 12: "KeyQ", 13: "KeyW", 14: "KeyE",
        15: "KeyR", 16: "KeyY", 17: "KeyT", 18: "Digit1", 19: "Digit2", 20: "Digit3", 21: "Digit4",
        22: "Digit6", 23: "Digit5", 24: "Equal", 25: "Digit9", 26: "Digit7", 27: "Minus",
        28: "Digit8", 29: "Digit0", 30: "BracketRight", 31: "KeyO", 32: "KeyU", 33: "BracketLeft",
        34: "KeyI", 35: "KeyP", 36: "Enter", 37: "KeyL", 38: "KeyJ", 39: "Quote", 40: "KeyK",
        41: "Semicolon", 42: "Backslash", 43: "Comma", 44: "Slash", 45: "KeyN", 46: "KeyM",
        47: "Period", 48: "Tab", 49: "Space", 50: "Backquote", 51: "Backspace", 53: "Escape",
        65: "NumpadDecimal", 67: "NumpadMultiply", 69: "NumpadAdd", 71: "NumLock",
        75: "NumpadDivide", 76: "NumpadEnter", 78: "NumpadSubtract", 81: "NumpadEqual",
        82: "Numpad0", 83: "Numpad1", 84: "Numpad2", 85: "Numpad3", 86: "Numpad4", 87: "Numpad5",
        88: "Numpad6", 89: "Numpad7", 91: "Numpad8", 92: "Numpad9", 96: "F5", 97: "F6", 98: "F7",
        99: "F3", 100: "F8", 101: "F9", 103: "F11", 109: "F10", 111: "F12", 115: "Home",
        116: "PageUp", 117: "Delete", 118: "F4", 119: "End", 120: "F2", 121: "PageDown", 122: "F1",
        123: "ArrowLeft", 124: "ArrowRight", 125: "ArrowDown", 126: "ArrowUp",
    ]

    /// Keys whose DOM `key` is a name, not the character they produce.
    private static let namedKeys: [UInt16: String] = [
        36: "Enter", 76: "Enter", 48: "Tab", 51: "Backspace", 53: "Escape", 117: "Delete",
        115: "Home", 119: "End", 116: "PageUp", 121: "PageDown", 123: "ArrowLeft",
        124: "ArrowRight", 125: "ArrowDown", 126: "ArrowUp", 71: "Clear", 122: "F1", 120: "F2",
        99: "F3", 118: "F4", 96: "F5", 97: "F6", 98: "F7", 100: "F8", 101: "F9", 109: "F10",
        103: "F11", 111: "F12",
    ]

    static func isReturn(_ event: NSEvent) -> Bool { event.keyCode == 36 || event.keyCode == 76 }
    static func isEscape(_ event: NSEvent) -> Bool { event.keyCode == 53 }

    /// The text a character key types (nil for a function key or a
    /// Control-modified key).
    static func text(of event: NSEvent) -> String? {
        if namedKeys[event.keyCode] != nil { return nil }
        if event.modifierFlags.contains(.control) || event.modifierFlags.contains(.command) { return nil }
        guard let characters = event.characters, !characters.isEmpty else { return nil }
        let printable = characters.unicodeScalars.allSatisfy { scalar in
            scalar.value >= 0x20 && scalar.value != 0x7f && !(0xF700...0xF8FF).contains(scalar.value)
        }
        return printable ? characters : nil
    }

    static func message(for event: NSEvent) -> [String: Any] {
        let flags = event.modifierFlags.intersection(.deviceIndependentFlagsMask)
        let text = text(of: event)
        let key = namedKeys[event.keyCode]
            ?? text
            ?? event.charactersIgnoringModifiers.flatMap { $0.isEmpty ? nil : $0 }
            ?? "Unidentified"
        var message: [String: Any] = [
            "t": "key",
            "key": key,
            "code": codes[event.keyCode] ?? "",
            "shift": flags.contains(.shift),
            "alt": flags.contains(.option),
            "ctrl": flags.contains(.control),
            "meta": flags.contains(.command),
            "repeat": event.isARepeat,
        ]
        if let text { message["text"] = text }
        return message
    }
}
