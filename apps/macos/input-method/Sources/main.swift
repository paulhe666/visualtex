import Cocoa
import InputMethodKit

// VisualTeX's macOS input method: formulas, which VisualTeX edits and draws in the field's line
// while the field holds their space as marked text. A palette input method, selected beside the
// user's own input source and never in its place (FormulaController).
let connectionName = Bundle.main.infoDictionary?["InputMethodConnectionName"] as? String
    ?? "com.visualtex.inputmethod.Formula_Connection"
let server = IMKServer(name: connectionName, bundleIdentifier: Bundle.main.bundleIdentifier)
InputLog.write("server started \(Bundle.main.infoDictionary?["CFBundleShortVersionString"] ?? "?") server=\(server != nil)")
SourceState.ensureSelected()
AppLink.shared.onMessage = { message in VTFormulaController.receive(message) }
AppLink.shared.keepConnected()
HostProbe.install()
NSApplication.shared.run()
