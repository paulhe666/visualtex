import Cocoa

/// macOS key events as librime's keys: X11 keysyms and modifier masks.
enum RimeKeys {
    static let shiftMask: Int32 = 1 << 0
    static let lockMask: Int32 = 1 << 1
    static let controlMask: Int32 = 1 << 2
    static let altMask: Int32 = 1 << 3
    static let superMask: Int32 = 1 << 26
    static let releaseMask: Int32 = 1 << 30

    static let shiftLeft: Int32 = 0xffe1
    static let shiftRight: Int32 = 0xffe2
    static let capsLock: Int32 = 0xffe5

    /// Keys with a keysym of their own, by macOS virtual key code.
    private static let named: [UInt16: Int32] = [
        36: 0xff0d, // Return
        76: 0xff8d, // keypad Enter
        48: 0xff09, // Tab
        49: 0x0020, // space
        51: 0xff08, // BackSpace
        53: 0xff1b, // Escape
        117: 0xffff, // Delete
        115: 0xff50, // Home
        119: 0xff57, // End
        116: 0xff55, // Page_Up
        121: 0xff56, // Page_Down
        123: 0xff51, // Left
        124: 0xff53, // Right
        125: 0xff54, // Down
        126: 0xff52, // Up
        122: 0xffbe, 120: 0xffbf, 99: 0xffc0, 118: 0xffc1, 96: 0xffc2, 97: 0xffc3,
        98: 0xffc4, 100: 0xffc5, 101: 0xffc6, 109: 0xffc7, 103: 0xffc8, 111: 0xffc9, // F1-F12
    ]

    static func modifiers(_ flags: NSEvent.ModifierFlags) -> Int32 {
        var mask: Int32 = 0
        if flags.contains(.shift) { mask |= shiftMask }
        if flags.contains(.capsLock) { mask |= lockMask }
        if flags.contains(.control) { mask |= controlMask }
        if flags.contains(.option) { mask |= altMask }
        if flags.contains(.command) { mask |= superMask }
        return mask
    }

    /// The keysym of a key-down event, nil for keys librime does not know.
    static func keysym(of event: NSEvent) -> Int32? {
        if let keysym = named[event.keyCode] { return keysym }
        let flags = event.modifierFlags
        // With Control or Option the character of the key itself.
        let source = flags.contains(.control) || flags.contains(.option) ? event.charactersIgnoringModifiers : event.characters
        guard let scalar = source?.unicodeScalars.first else { return nil }
        let value = scalar.value
        if value >= 0x20 && value < 0x7f { return Int32(value) }
        // Other characters: keysym 0x01000000 + code point (X11 Unicode keysyms).
        if value >= 0xa0 && !(0xF700...0xF8FF).contains(value) { return Int32(0x0100_0000 + value) }
        return nil
    }
}
