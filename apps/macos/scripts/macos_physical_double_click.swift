import CoreGraphics
import Foundation
import AppKit

func usage() -> Never {
    FileHandle.standardError.write(
        Data("Usage: swift macos_physical_double_click.swift <screen-x> <screen-y> [--appkit-y] [--activate-powerpoint]\n".utf8)
    )
    exit(2)
}

guard CommandLine.arguments.count >= 3,
      let x = Double(CommandLine.arguments[1]),
      let sourceY = Double(CommandLine.arguments[2]),
      x.isFinite,
      sourceY.isFinite else {
    usage()
}
let options = Set(CommandLine.arguments.dropFirst(3))
if options.count != CommandLine.arguments.count - 3 ||
    !options.isSubset(of: ["--appkit-y", "--activate-powerpoint"]) {
    usage()
}
let appKitY = options.contains("--appkit-y")
if options.contains("--activate-powerpoint") {
    guard let powerpoint = NSRunningApplication.runningApplications(
        withBundleIdentifier: "com.microsoft.Powerpoint"
    ).first else {
        FileHandle.standardError.write(Data("PowerPoint is not running\n".utf8))
        exit(1)
    }
    guard powerpoint.activate() else {
        FileHandle.standardError.write(Data("Unable to activate PowerPoint\n".utf8))
        exit(1)
    }
    var activationError: NSDictionary?
    NSAppleScript(source: "tell application \"Microsoft PowerPoint\" to activate")?
        .executeAndReturnError(&activationError)
    for _ in 0..<20 {
        if NSWorkspace.shared.frontmostApplication?.bundleIdentifier == "com.microsoft.Powerpoint" {
            break
        }
        usleep(50_000)
    }
    guard NSWorkspace.shared.frontmostApplication?.bundleIdentifier == "com.microsoft.Powerpoint" else {
        FileHandle.standardError.write(Data("PowerPoint did not become frontmost\n".utf8))
        exit(1)
    }
}
let mainDisplayBounds = CGDisplayBounds(CGMainDisplayID())
let y = appKitY
    ? mainDisplayBounds.maxY - sourceY
    : sourceY
let point = CGPoint(x: x, y: y)
let source = CGEventSource(stateID: .hidSystemState)

func post(_ type: CGEventType, clickState: Int64) {
    guard let event = CGEvent(
        mouseEventSource: source,
        mouseType: type,
        mouseCursorPosition: point,
        mouseButton: .left
    ) else {
        FileHandle.standardError.write(Data("Unable to create Quartz mouse event\n".utf8))
        exit(1)
    }
    event.setIntegerValueField(.mouseEventClickState, value: clickState)
    event.post(tap: .cghidEventTap)
}

post(.mouseMoved, clickState: 0)
usleep(80_000)
post(.leftMouseDown, clickState: 1)
post(.leftMouseUp, clickState: 1)
usleep(90_000)
post(.leftMouseDown, clickState: 2)
post(.leftMouseUp, clickState: 2)
FileHandle.standardOutput.write(
    Data("CLICK|\(x)|\(y)|appKitY=\(appKitY)\n".utf8)
)
