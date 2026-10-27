import Cocoa
import InputMethodKit

// VisualTeX as a macOS input method: Chinese through pinyin (RIME), English
// (Shift switches), and formulas, which VisualTeX edits and draws in the
// field's line while the field holds their space as marked text.
let connectionName = Bundle.main.infoDictionary?["InputMethodConnectionName"] as? String
    ?? "com.visualtex.inputmethod.VisualTeX_Connection"
let server = IMKServer(name: connectionName, bundleIdentifier: Bundle.main.bundleIdentifier)
InputLog.write("server started \(Bundle.main.infoDictionary?["CFBundleShortVersionString"] ?? "?") server=\(server != nil)")
RimeEngine.shared.start()
AppLink.shared.onMessage = { message in VTInputController.receive(message) }
AppLink.shared.keepConnected()
NSApplication.shared.run()
