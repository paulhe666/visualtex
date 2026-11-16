import Carbon
import Foundation

/// This input method among the system's input sources. VisualTeX enables it (it writes the list
/// macOS keeps, as System Settings does) and starts this process, which then selects itself: a
/// palette input method is selected beside the keyboard input source, never in its place, and stays
/// selected. VisualTeX is told ("hello", "selected"): only a selected input method has the field
/// that has the keyboard.
enum SourceState {
    private(set) static var selected = false
    private static var tries = 0

    private static func source() -> TISInputSource? {
        guard let id = Bundle.main.bundleIdentifier else { return nil }
        let filter = [kTISPropertyInputSourceID as String: id] as CFDictionary
        return (TISCreateInputSourceList(filter, true)?.takeRetainedValue() as? [TISInputSource])?.first
    }

    private static func flag(_ source: TISInputSource, _ key: CFString) -> Bool {
        guard let raw = TISGetInputSourceProperty(source, key) else { return false }
        return CFBooleanGetValue(Unmanaged<CFBoolean>.fromOpaque(raw).takeUnretainedValue())
    }

    /// Selects this input method once it is enabled (the list is read again a moment after VisualTeX
    /// wrote it: asked a few times).
    static func ensureSelected() {
        tries += 1
        var enabled = false
        if let source = source() {
            enabled = flag(source, kTISPropertyInputSourceIsEnabled)
            if enabled, !flag(source, kTISPropertyInputSourceIsSelected) {
                let status = TISSelectInputSource(source)
                InputLog.write("selected as an input source: status=\(status)")
            }
            let now = flag(source, kTISPropertyInputSourceIsSelected)
            if now != selected {
                selected = now
                AppLink.shared.send(["t": "selected", "selected": now])
            }
        }
        if tries == 1 || (!selected && tries % 10 == 0) {
            InputLog.write("input source enabled=\(enabled) selected=\(selected)")
        }
        // (not yet enabled, or deselected by something else later: looked at again)
        DispatchQueue.main.asyncAfter(deadline: .now() + (selected ? 30 : 0.5)) { ensureSelected() }
    }
}
