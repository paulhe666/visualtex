import AppKit
import ApplicationServices
import Carbon
import CoreGraphics
import Foundation

func attribute(_ element: AXUIElement, _ name: CFString) -> CFTypeRef? {
    var value: CFTypeRef?
    guard AXUIElementCopyAttributeValue(element, name, &value) == .success else { return nil }
    return value
}

func stringAttribute(_ element: AXUIElement, _ name: CFString) -> String? {
    attribute(element, name) as? String
}

func boolAttribute(_ element: AXUIElement, _ name: CFString) -> Bool? {
    if let value = attribute(element, name) as? Bool { return value }
    if let value = attribute(element, name) as? NSNumber { return value.boolValue }
    return nil
}

func pointAttribute(_ element: AXUIElement, _ name: CFString) -> [Double]? {
    guard let raw = attribute(element, name), CFGetTypeID(raw) == AXValueGetTypeID() else { return nil }
    let value = unsafeBitCast(raw, to: AXValue.self)
    var point = CGPoint.zero
    guard AXValueGetType(value) == .cgPoint, AXValueGetValue(value, .cgPoint, &point) else { return nil }
    return [point.x, point.y]
}

func sizeAttribute(_ element: AXUIElement, _ name: CFString) -> [Double]? {
    guard let raw = attribute(element, name), CFGetTypeID(raw) == AXValueGetTypeID() else { return nil }
    let value = unsafeBitCast(raw, to: AXValue.self)
    var size = CGSize.zero
    guard AXValueGetType(value) == .cgSize, AXValueGetValue(value, .cgSize, &size) else { return nil }
    return [size.width, size.height]
}

func children(_ element: AXUIElement) -> [AXUIElement] {
    attribute(element, kAXChildrenAttribute as CFString) as? [AXUIElement] ?? []
}

func mainWindow(_ pid: pid_t) -> AXUIElement? {
    let app = AXUIElementCreateApplication(pid)
    guard let windows = attribute(app, kAXWindowsAttribute as CFString) as? [AXUIElement] else { return nil }
    return windows.first { stringAttribute($0, kAXTitleAttribute as CFString) == "VisualTeX" } ?? windows.first
}

struct AxNode {
    let element: AXUIElement
    let role: String
    let title: String
    let description: String
    let value: String
    let focused: Bool
    let position: [Double]?
    let size: [Double]?
}

func collect(_ root: AXUIElement, maxDepth: Int = 24) -> [AxNode] {
    var result: [AxNode] = []
    func walk(_ element: AXUIElement, _ depth: Int) {
        if depth > maxDepth { return }
        let role = stringAttribute(element, kAXRoleAttribute as CFString) ?? ""
        let title = stringAttribute(element, kAXTitleAttribute as CFString) ?? ""
        let description = stringAttribute(element, kAXDescriptionAttribute as CFString) ?? ""
        let value = stringAttribute(element, kAXValueAttribute as CFString) ?? ""
        let focused = boolAttribute(element, kAXFocusedAttribute as CFString) ?? false
        if role == kAXTextFieldRole as String || role == kAXTextAreaRole as String || role == kAXButtonRole as String || role == kAXStaticTextRole as String {
            result.append(AxNode(
                element: element,
                role: role,
                title: title,
                description: description,
                value: value,
                focused: focused,
                position: pointAttribute(element, kAXPositionAttribute as CFString),
                size: sizeAttribute(element, kAXSizeAttribute as CFString)
            ))
        }
        for child in children(element) { walk(child, depth + 1) }
    }
    walk(root, 0)
    return result
}

func json(_ value: Any) {
    let data = try! JSONSerialization.data(withJSONObject: value, options: [.prettyPrinted, .sortedKeys])
    print(String(data: data, encoding: .utf8)!)
}

func currentInputSourceID() -> String {
    let source = TISCopyCurrentKeyboardInputSource().takeRetainedValue()
    guard let rawID = TISGetInputSourceProperty(source, kTISPropertyInputSourceID) else {
        fatalError("current input source has no ID")
    }
    return Unmanaged<CFString>.fromOpaque(rawID).takeUnretainedValue() as String
}

func snapshot(_ pid: pid_t) {
    guard let window = mainWindow(pid) else { fatalError("main window missing") }
    let nodes = collect(window)
    let fields = nodes.filter { $0.role == kAXTextFieldRole as String }.map { node -> [String: Any] in
        var item: [String: Any] = [
            "title": node.title,
            "description": node.description,
            "value": node.value,
            "focused": node.focused,
        ]
        if let position = node.position { item["position"] = position }
        if let size = node.size { item["size"] = size }
        return item
    }
    let areas = nodes.filter { $0.role == kAXTextAreaRole as String }.map { node -> [String: Any] in
        var item: [String: Any] = [
            "title": node.title,
            "description": node.description,
            "value": node.value,
            "focused": node.focused,
        ]
        if let position = node.position { item["position"] = position }
        if let size = node.size { item["size"] = size }
        return item
    }
    let named = nodes.filter {
        let text = [$0.title, $0.description, $0.value].joined(separator: " ")
        return text.contains("\\alpha") || text.contains("\\aleph") || text.contains("\\int") || text.contains("VisualTeX IME Diagnostic Trace")
    }.map { node in
        ["role": node.role, "title": node.title, "description": node.description, "value": node.value]
    }
    json(["pid": Int(pid), "fields": fields, "areas": areas, "named": named])
}

func formulaFields(_ pid: pid_t) -> [AxNode] {
    guard let window = mainWindow(pid) else { return [] }
    return collect(window).filter { node in
        guard node.role == kAXTextFieldRole as String else { return false }
        let label = [node.title, node.description].joined(separator: " ")
        return !label.contains("公式文档标题") && !label.contains("Formula document title")
    }
}

func focusLastFormula(_ pid: pid_t) {
    let fields = formulaFields(pid)
    guard let target = fields.last else { fatalError("formula field missing") }
    let result = AXUIElementSetAttributeValue(target.element, kAXFocusedAttribute as CFString, kCFBooleanTrue)
    json(["result": result.rawValue, "fieldCount": fields.count])
}

func postKey(_ pid: pid_t, keyCode: CGKeyCode, flags: CGEventFlags = []) {
    guard let down = CGEvent(keyboardEventSource: nil, virtualKey: keyCode, keyDown: true),
          let up = CGEvent(keyboardEventSource: nil, virtualKey: keyCode, keyDown: false) else {
        fatalError("unable to create keyboard event")
    }
    down.flags = flags
    up.flags = flags
    down.postToPid(pid)
    usleep(45_000)
    up.postToPid(pid)
    usleep(75_000)
}

func postGlobalKey(_ pid: pid_t, keyCode: CGKeyCode, flags: CGEventFlags = []) {
    guard let app = NSRunningApplication(processIdentifier: pid) else { fatalError("application missing") }
    _ = app.activate(options: [])
    usleep(120_000)
    guard let down = CGEvent(keyboardEventSource: nil, virtualKey: keyCode, keyDown: true),
          let up = CGEvent(keyboardEventSource: nil, virtualKey: keyCode, keyDown: false) else {
        fatalError("unable to create global keyboard event")
    }
    down.flags = flags
    up.flags = flags
    down.post(tap: .cghidEventTap)
    usleep(45_000)
    up.post(tap: .cghidEventTap)
    usleep(100_000)
}

func syntheticControlSpace(_ pid: pid_t) {
    guard let app = NSRunningApplication(processIdentifier: pid) else { fatalError("application missing") }
    let inputSourceBefore = currentInputSourceID()
    _ = app.activate(options: [])
    usleep(140_000)
    let controlDown = CGEvent(keyboardEventSource: nil, virtualKey: 59, keyDown: true)!
    controlDown.flags = [.maskControl]
    controlDown.post(tap: .cghidEventTap)
    usleep(55_000)
    let spaceDown = CGEvent(keyboardEventSource: nil, virtualKey: 49, keyDown: true)!
    spaceDown.flags = [.maskControl]
    spaceDown.post(tap: .cghidEventTap)
    usleep(70_000)
    let spaceUp = CGEvent(keyboardEventSource: nil, virtualKey: 49, keyDown: false)!
    spaceUp.flags = [.maskControl]
    spaceUp.post(tap: .cghidEventTap)
    usleep(55_000)
    let controlUp = CGEvent(keyboardEventSource: nil, virtualKey: 59, keyDown: false)!
    controlUp.flags = []
    controlUp.post(tap: .cghidEventTap)
    usleep(260_000)
    let inputSourceAfter = currentInputSourceID()
    let verified = inputSourceBefore == "com.apple.inputmethod.SCIM.ITABC" &&
        inputSourceAfter == "com.apple.keylayout.ABC"
    json([
        "pid": Int(pid),
        "mode": "cg-event-control-space",
        "inputSourceBefore": inputSourceBefore,
        "inputSourceAfter": inputSourceAfter,
        "verified": verified,
    ])
    if !verified { exit(30) }
}

func physicalCapsLockFromHIDState(_ pid: pid_t) {
    guard let app = NSRunningApplication(processIdentifier: pid),
          let source = CGEventSource(stateID: .hidSystemState) else {
        fatalError("application or HID event source missing")
    }
    let inputSourceBefore = currentInputSourceID()
    _ = app.activate(options: [])
    usleep(140_000)
    let down = CGEvent(keyboardEventSource: source, virtualKey: 57, keyDown: true)!
    down.type = .flagsChanged
    down.flags = [.maskAlphaShift, .maskNonCoalesced]
    down.post(tap: .cghidEventTap)
    usleep(22_000)
    // The successful physical trace reports the release-side transition as
    // flagsChanged/keyCode 255 rather than a CapsLock keyUp.
    let up = CGEvent(keyboardEventSource: source, virtualKey: 255, keyDown: false)!
    up.type = .flagsChanged
    up.flags = [.maskNonCoalesced]
    up.post(tap: .cghidEventTap)
    usleep(320_000)
    let inputSourceAfter = currentInputSourceID()
    let verified = inputSourceBefore == "com.apple.inputmethod.SCIM.ITABC" &&
        inputSourceAfter == "com.apple.keylayout.ABC"
    json([
        "pid": Int(pid),
        "mode": "hid-state-caps-lock",
        "inputSourceBefore": inputSourceBefore,
        "inputSourceAfter": inputSourceAfter,
        "verified": verified,
    ])
    if !verified { exit(31) }
}

func parseFlags(_ values: ArraySlice<String>) -> CGEventFlags {
    let flagArgs = Set(values)
    var flags: CGEventFlags = []
    if flagArgs.contains("ctrl") { flags.insert(.maskControl) }
    if flagArgs.contains("meta") { flags.insert(.maskCommand) }
    if flagArgs.contains("shift") { flags.insert(.maskShift) }
    if flagArgs.contains("alt") { flags.insert(.maskAlternate) }
    return flags
}

func selectInputSource(_ id: String) {
    let filter = [kTISPropertyInputSourceID as String: id] as CFDictionary
    let list = TISCreateInputSourceList(filter, false).takeRetainedValue() as! [TISInputSource]
    guard let source = list.first else { fatalError("input source not found: \(id)") }
    let status = TISSelectInputSource(source)
    guard status == noErr else { fatalError("input source selection failed: \(status)") }
    for _ in 0..<25 {
        let currentID = currentInputSourceID()
        if currentID == id {
            json(["status": status, "id": id, "currentId": currentID, "verified": true])
            return
        }
        usleep(40_000)
    }
    fatalError("input source selection was not applied: expected \(id), actual \(currentInputSourceID())")
}

let args = CommandLine.arguments
if args.count < 2 {
    fatalError("usage: swift macos_ax_ime_probe.swift <snapshot|focus-last|key|current-source|source> ...")
}

switch args[1] {
case "snapshot":
    guard args.count >= 3, let pid = pid_t(args[2]) else { fatalError("snapshot <pid>") }
    snapshot(pid)
case "focus-last":
    guard args.count >= 3, let pid = pid_t(args[2]) else { fatalError("focus-last <pid>") }
    focusLastFormula(pid)
case "key":
    guard args.count >= 4, let pid = pid_t(args[2]), let code = UInt16(args[3]) else { fatalError("key <pid> <keyCode> [flags]") }
    let flags = parseFlags(args.dropFirst(4))
    postKey(pid, keyCode: CGKeyCode(code), flags: flags)
    json(["pid": Int(pid), "keyCode": Int(code), "mode": "pid"])
case "key-global":
    guard args.count >= 4, let pid = pid_t(args[2]), let code = UInt16(args[3]) else { fatalError("key-global <pid> <keyCode> [flags]") }
    let flags = parseFlags(args.dropFirst(4))
    postGlobalKey(pid, keyCode: CGKeyCode(code), flags: flags)
    json(["pid": Int(pid), "keyCode": Int(code), "mode": "global"])
case "ctrl-space":
    guard args.count >= 3, let pid = pid_t(args[2]) else { fatalError("ctrl-space <pid>") }
    syntheticControlSpace(pid)
case "caps-lock-hid":
    guard args.count >= 3, let pid = pid_t(args[2]) else { fatalError("caps-lock-hid <pid>") }
    physicalCapsLockFromHIDState(pid)
case "current-source":
    json([
        "id": currentInputSourceID(),
        "capsLockKeyState": CGEventSource.keyState(.hidSystemState, key: 57),
        "alphaShiftFlag": CGEventSource.flagsState(.hidSystemState).contains(.maskAlphaShift),
    ])
case "source":
    guard args.count >= 3 else { fatalError("source <inputSourceId>") }
    selectInputSource(args[2])
default:
    fatalError("unknown command: \(args[1])")
}
