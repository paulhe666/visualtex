import Cocoa

/// The candidate window under the line: one row of numbered candidates,
/// the highlighted one in the accent colour, like the system's Pinyin. It
/// never takes the keyboard or the mouse.
final class CandidateWindow {
    static let shared = CandidateWindow()

    private let panel: NSPanel
    private let effect: NSVisualEffectView
    private let row = CandidateRow()

    private init() {
        panel = NSPanel(
            contentRect: NSRect(x: 0, y: 0, width: 200, height: 34),
            styleMask: [.borderless, .nonactivatingPanel],
            backing: .buffered,
            defer: true
        )
        panel.level = NSWindow.Level(rawValue: NSWindow.Level.popUpMenu.rawValue + 1)
        panel.isOpaque = false
        panel.backgroundColor = .clear
        panel.hasShadow = true
        panel.ignoresMouseEvents = true
        panel.hidesOnDeactivate = false
        panel.collectionBehavior = [.canJoinAllSpaces, .fullScreenAuxiliary, .transient, .ignoresCycle]
        effect = NSVisualEffectView(frame: NSRect(x: 0, y: 0, width: 200, height: 34))
        effect.material = .popover
        effect.state = .active
        effect.blendingMode = .behindWindow
        effect.wantsLayer = true
        effect.layer?.cornerRadius = 8
        effect.layer?.masksToBounds = true
        effect.autoresizingMask = [.width, .height]
        row.frame = effect.bounds
        row.autoresizingMask = [.width, .height]
        effect.addSubview(row)
        panel.contentView = effect
    }

    /// Shows the candidates under `line` (AppKit screen coordinates), or
    /// hides the window when there are none.
    func show(_ candidates: [RimeEngine.Candidate], labels: [String], highlighted: Int, under line: NSRect) {
        guard !candidates.isEmpty, line != .zero else {
            hide()
            return
        }
        row.update(candidates: candidates, labels: labels, highlighted: highlighted)
        let size = row.fittingSize
        var frame = NSRect(x: line.minX - 6, y: line.minY - size.height - 3, width: size.width, height: size.height)
        if let screen = NSScreen.screens.first(where: { $0.frame.contains(NSPoint(x: line.minX, y: line.midY)) }) ?? NSScreen.main {
            let visible = screen.visibleFrame
            frame.origin.x = min(max(frame.origin.x, visible.minX), visible.maxX - frame.width)
            if frame.minY < visible.minY { frame.origin.y = line.maxY + 3 }
        }
        panel.setFrame(frame, display: true)
        row.needsDisplay = true
        panel.orderFrontRegardless()
    }

    func hide() {
        panel.orderOut(nil)
    }
}

/// The candidates drawn in one row.
private final class CandidateRow: NSView {
    private var items: [(label: NSAttributedString, text: NSAttributedString, comment: NSAttributedString?)] = []
    private var highlighted = 0
    private var rects: [NSRect] = []
    private let inset = NSEdgeInsets(top: 5, left: 6, bottom: 5, right: 6)
    private let gap: CGFloat = 4
    private let padding: CGFloat = 7

    override var isFlipped: Bool { true }

    func update(candidates: [RimeEngine.Candidate], labels: [String], highlighted: Int) {
        self.highlighted = highlighted
        let textFont = NSFont.systemFont(ofSize: 17)
        let labelFont = NSFont.monospacedDigitSystemFont(ofSize: 12, weight: .regular)
        let commentFont = NSFont.systemFont(ofSize: 12)
        items = candidates.enumerated().map { index, candidate in
            let isHighlighted = index == highlighted
            let label = index < labels.count ? labels[index] : "\(index + 1)"
            return (
                NSAttributedString(string: label, attributes: [
                    .font: labelFont,
                    .foregroundColor: isHighlighted ? NSColor.white.withAlphaComponent(0.85) : NSColor.secondaryLabelColor,
                ]),
                NSAttributedString(string: candidate.text, attributes: [
                    .font: textFont,
                    .foregroundColor: isHighlighted ? NSColor.white : NSColor.labelColor,
                ]),
                candidate.comment.isEmpty ? nil : NSAttributedString(string: candidate.comment, attributes: [
                    .font: commentFont,
                    .foregroundColor: isHighlighted ? NSColor.white.withAlphaComponent(0.8) : NSColor.tertiaryLabelColor,
                ])
            )
        }
        layoutItems()
    }

    private func layoutItems() {
        rects = []
        var x = inset.left
        let height: CGFloat = 26
        for item in items {
            var width = padding + item.label.size().width + 4 + item.text.size().width + padding
            if let comment = item.comment { width += 4 + comment.size().width }
            rects.append(NSRect(x: x, y: inset.top, width: ceil(width), height: height))
            x += ceil(width) + gap
        }
    }

    override var fittingSize: NSSize {
        let width = (rects.last?.maxX ?? 0) + inset.right
        return NSSize(width: max(40, width), height: inset.top + 26 + inset.bottom)
    }

    override func draw(_ dirtyRect: NSRect) {
        for (index, item) in items.enumerated() where index < rects.count {
            let rect = rects[index]
            if index == highlighted {
                NSColor.controlAccentColor.setFill()
                NSBezierPath(roundedRect: rect, xRadius: 6, yRadius: 6).fill()
            }
            let labelSize = item.label.size()
            let textSize = item.text.size()
            var x = rect.minX + padding
            item.label.draw(at: NSPoint(x: x, y: rect.midY - labelSize.height / 2 + 0.5))
            x += labelSize.width + 4
            item.text.draw(at: NSPoint(x: x, y: rect.midY - textSize.height / 2))
            x += textSize.width + 4
            if let comment = item.comment {
                let size = comment.size()
                comment.draw(at: NSPoint(x: x, y: rect.midY - size.height / 2 + 0.5))
            }
        }
    }
}
