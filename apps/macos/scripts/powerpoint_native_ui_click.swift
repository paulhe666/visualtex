// UI acceptance helper only; never included in production builds.
import AppKit
import CoreGraphics
import Foundation
let args = CommandLine.arguments
guard args.count == 3, let x = Double(args[1]), let y = Double(args[2]), x.isFinite, y.isFinite else { exit(2) }
guard let app = NSRunningApplication.runningApplications(withBundleIdentifier: "com.microsoft.Powerpoint").first else { exit(3) }
app.activate(options: [.activateAllWindows])
usleep(300_000)
let point = CGPoint(x: x, y: y)
let source = CGEventSource(stateID: .hidSystemState)
for type in [CGEventType.mouseMoved, .leftMouseDown, .leftMouseUp] {
    guard let event = CGEvent(mouseEventSource: source, mouseType: type, mouseCursorPosition: point, mouseButton: .left) else { exit(4) }
    event.setIntegerValueField(.mouseEventClickState, value: type == .mouseMoved ? 0 : 1)
    event.post(tap: .cghidEventTap)
    usleep(50_000)
}
