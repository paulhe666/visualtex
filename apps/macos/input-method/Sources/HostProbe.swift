import Cocoa
import InputMethodKit

/// Development and support: what a field answers about marked text, written down as it is. The
/// placement of formulas (FormulaHold, SpaceModel) rests on what applications really answer, and
/// they differ: this records it for one application, so that a new one is measured, not guessed.
///
/// `kill -USR1 <the input method's pid>` runs the steps in /tmp/visualtex-ime-probe.json on the
/// field that has the keyboard and appends what the field answered to
/// /tmp/visualtex-ime-probe.out (one JSON object a line). Nothing happens without that file.
///
/// A step: `text` (set as marked text; `sel` = [location, length] of the selection in it, the end
/// when left out), or `insert` (typed into the field), or `clear` (no marked text); `reads`: the
/// milliseconds after it at which the field is asked (default: once, at once); `beyond`: how far
/// past the last character it is asked (TextKit answers nothing for a character it was once asked
/// about while the marked text did not reach it, for as long as the marked text only grows); `label`.
enum HostProbe {
    private static var source: DispatchSourceSignal?
    private static var running = false
    private static let script = "/tmp/visualtex-ime-probe.json"
    private static let output = "/tmp/visualtex-ime-probe.out"

    private struct Step: Decodable {
        var label: String?
        var text: String?
        var sel: [Int]?
        var insert: String?
        var clear: Bool?
        var reads: [Int]?
        /// How far past the last character the field is asked (default 1: the place after the last
        /// character and the one after that; -1: characters of the marked text only).
        var beyond: Int?
    }

    static func install() {
        signal(SIGUSR1, SIG_IGN)
        let source = DispatchSource.makeSignalSource(signal: SIGUSR1, queue: .main)
        source.setEventHandler { run() }
        source.resume()
        self.source = source
    }

    private static func run() {
        guard !running, let data = FileManager.default.contents(atPath: script) else { return }
        guard let steps = try? JSONDecoder().decode([Step].self, from: data) else {
            InputLog.write("probe: the steps cannot be read")
            return
        }
        guard let controller = VTFormulaController.current, let client = controller.client() else {
            InputLog.write("probe: no field has the keyboard")
            write(["error": "no field has the keyboard"])
            return
        }
        running = true
        InputLog.write("probe: \(steps.count) steps in \(client.bundleIdentifier() ?? "?")")
        step(0, steps, controller)
    }

    private static func step(_ index: Int, _ steps: [Step], _ controller: VTFormulaController) {
        guard index < steps.count, VTFormulaController.current === controller, let client = controller.client() else {
            // (cut short: the field is not left with the probe's marked text)
            if index < steps.count { controller.client()?.setMarkedText("", selectionRange: NSRange(location: 0, length: 0), replacementRange: VTFormulaController.noRange) }
            running = false
            write(["done": index, "of": steps.count])
            InputLog.write("probe: done \(index)/\(steps.count)")
            return
        }
        let step = steps[index]
        var length = 0
        if step.clear == true {
            client.setMarkedText("", selectionRange: NSRange(location: 0, length: 0), replacementRange: VTFormulaController.noRange)
        } else if let insert = step.insert {
            client.insertText(insert, replacementRange: VTFormulaController.noRange)
        } else if let text = step.text {
            length = (text as NSString).length
            let selection = step.sel.flatMap { $0.count == 2 ? NSRange(location: $0[0], length: $0[1]) : nil } ?? NSRange(location: length, length: 0)
            client.setMarkedText(
                NSAttributedString(string: text, attributes: [.underlineStyle: NSUnderlineStyle.single.rawValue]),
                selectionRange: selection,
                replacementRange: VTFormulaController.noRange
            )
        }
        let started = Date()
        let reads = (step.reads ?? [0]).sorted()
        for (order, delay) in reads.enumerated() {
            DispatchQueue.main.asyncAfter(deadline: .now() + Double(delay) / 1000) {
                guard VTFormulaController.current === controller, let client = controller.client() else { return }
                var record = read(client, length, beyond: step.beyond ?? 1)
                record["step"] = index
                record["label"] = step.label ?? ""
                record["t"] = Int(Date().timeIntervalSince(started) * 1000)
                write(record)
                if order == reads.count - 1 {
                    DispatchQueue.main.asyncAfter(deadline: .now() + 0.03) { self.step(index + 1, steps, controller) }
                }
            }
        }
    }

    /// Everything the field can be asked about a marked text `length` characters long.
    private static func read(_ client: IMKTextInput, _ length: Int, beyond: Int) -> [String: Any] {
        let began = Date()
        let marked = client.markedRange()
        let selected = client.selectedRange()
        var record: [String: Any] = ["client": client.bundleIdentifier() ?? "?", "len": length]
        record["marked"] = pair(marked)
        record["selected"] = pair(selected)
        // each character's place by its index in the marked text, and the two after the last
        var at: [Any] = []
        for index in 0...min(length + 1, 90) {
            var rect = NSRect.zero
            if index <= length + beyond { _ = client.attributes(forCharacterIndex: index, lineHeightRectangle: &rect) }
            at.append(quad(rect))
        }
        record["at"] = at
        // by ranges of the field's text: the selection, the marked text, some of its characters
        var actual = VTFormulaController.noRange
        if selected.location != NSNotFound {
            record["firstSel"] = quad(client.firstRect(forCharacterRange: selected, actualRange: &actual)) + pair(actual)
        }
        if marked.location != NSNotFound, marked.length > 0 {
            record["firstMarked"] = quad(client.firstRect(forCharacterRange: marked, actualRange: &actual)) + pair(actual)
            var chars: [Any] = []
            for index in [0, 1, 2, marked.length - 1] where index >= 0 && index < marked.length {
                chars.append([index] + quad(client.firstRect(forCharacterRange: NSRange(location: marked.location + index, length: 1), actualRange: &actual)))
            }
            record["firstChar"] = chars
            record["firstEnd"] = quad(client.firstRect(forCharacterRange: NSRange(location: NSMaxRange(marked), length: 0), actualRange: &actual))
        }
        record["ms"] = Int(Date().timeIntervalSince(began) * 1000)
        return record
    }

    private static func pair(_ range: NSRange) -> [Int] {
        range.location == NSNotFound ? [-1, 0] : [range.location, range.length]
    }

    private static func quad(_ rect: NSRect) -> [Int] {
        // tenths of a point
        [rect.minX, rect.minY, rect.width, rect.height].map { $0.isFinite ? Int(($0 * 10).rounded()) : -1 }
    }

    private static func write(_ record: [String: Any]) {
        guard let data = try? JSONSerialization.data(withJSONObject: record, options: [.sortedKeys]) else { return }
        let line = data + Data([0x0A])
        if let handle = FileHandle(forWritingAtPath: output) {
            handle.seekToEndOfFile()
            handle.write(line)
            try? handle.close()
        } else {
            try? line.write(to: URL(fileURLWithPath: output))
        }
    }
}
