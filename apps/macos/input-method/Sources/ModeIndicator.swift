import Cocoa

/// A small "中" / "英" under the caret for a moment after Shift switches
/// between Chinese and English.
final class ModeIndicator {
    static let shared = ModeIndicator()

    private let panel: NSPanel
    private let label = NSTextField(labelWithString: "")
    private var hideWork: DispatchWorkItem?

    private init() {
        panel = NSPanel(
            contentRect: NSRect(x: 0, y: 0, width: 30, height: 26),
            styleMask: [.borderless, .nonactivatingPanel],
            backing: .buffered,
            defer: true
        )
        panel.level = NSWindow.Level(rawValue: NSWindow.Level.popUpMenu.rawValue + 1)
        panel.isOpaque = false
        panel.backgroundColor = .clear
        panel.hasShadow = true
        panel.ignoresMouseEvents = true
        panel.collectionBehavior = [.canJoinAllSpaces, .fullScreenAuxiliary, .transient, .ignoresCycle]
        let effect = NSVisualEffectView(frame: NSRect(x: 0, y: 0, width: 30, height: 26))
        effect.material = .popover
        effect.state = .active
        effect.wantsLayer = true
        effect.layer?.cornerRadius = 6
        effect.layer?.masksToBounds = true
        label.font = .systemFont(ofSize: 14, weight: .medium)
        label.alignment = .center
        label.frame = NSRect(x: 0, y: 4, width: 30, height: 18)
        effect.addSubview(label)
        panel.contentView = effect
    }

    func show(_ text: String, under caret: NSRect) {
        guard caret != .zero else { return }
        label.stringValue = text
        panel.setFrameOrigin(NSPoint(x: caret.minX - 4, y: caret.minY - 30))
        panel.orderFrontRegardless()
        hideWork?.cancel()
        let work = DispatchWorkItem { [weak self] in self?.panel.orderOut(nil) }
        hideWork = work
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.8, execute: work)
    }
}
