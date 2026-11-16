import CryptoKit
import Darwin
import Foundation
import Security

/// The link to the VisualTeX app: a Unix socket in its application support
/// folder carrying newline-delimited JSON both ways. VisualTeX edits the
/// formula and decides what the field shows; this side forwards the keys and
/// types what it is told. The socket's other end must be VisualTeX itself
/// (its code signature), since what it sends is typed into the user's field.
final class AppLink {
    static let shared = AppLink()

    static let socketPath = FileManager.default.homeDirectoryForCurrentUser
        .appendingPathComponent("Library/Application Support/com.visualtex.studio/anywhere-ime.sock").path

    /// Called on the main thread for every message from VisualTeX.
    var onMessage: (([String: Any]) -> Void)?

    private let lock = NSLock()
    private var socket: Int32 = -1
    private var lastAttempt = Date.distantPast
    private var refusalLogged = false
    private var keeper: Timer?
    private let writer = DispatchQueue(label: "com.visualtex.inputmethod.link.write")

    var isConnected: Bool {
        lock.lock()
        defer { lock.unlock() }
        return socket >= 0
    }

    /// Connects now and again within a second whenever the link is down
    /// (VisualTeX started later, or restarted), so the shortcut finds the
    /// input method linked without the field having to change first.
    func keepConnected() {
        connect()
        keeper?.invalidate()
        keeper = Timer.scheduledTimer(withTimeInterval: 1, repeats: true) { [weak self] _ in
            guard let self, !self.isConnected else { return }
            self.connect()
        }
    }

    /// Connects unless connected (at most twice a second when VisualTeX is
    /// not running).
    @discardableResult
    func connect() -> Bool {
        lock.lock()
        defer { lock.unlock() }
        if socket >= 0 { return true }
        if Date().timeIntervalSince(lastAttempt) < 0.5 { return false }
        lastAttempt = Date()
        let candidate = Darwin.socket(AF_UNIX, SOCK_STREAM, 0)
        guard candidate >= 0 else { return false }
        var address = sockaddr_un()
        address.sun_family = sa_family_t(AF_UNIX)
        let capacity = MemoryLayout.size(ofValue: address.sun_path)
        guard Self.socketPath.utf8.count < capacity else {
            close(candidate)
            return false
        }
        withUnsafeMutablePointer(to: &address.sun_path) { pointer in
            pointer.withMemoryRebound(to: CChar.self, capacity: capacity) { buffer in
                _ = strncpy(buffer, Self.socketPath, capacity - 1)
            }
        }
        let connected = withUnsafePointer(to: &address) { pointer in
            pointer.withMemoryRebound(to: sockaddr.self, capacity: 1) {
                Darwin.connect(candidate, $0, socklen_t(MemoryLayout<sockaddr_un>.size))
            }
        }
        guard connected == 0 else {
            close(candidate)
            return false
        }
        guard PeerCheck.isVisualTeX(candidate) else {
            if !refusalLogged { InputLog.write("link refused: the socket's owner is not VisualTeX") }
            refusalLogged = true
            close(candidate)
            return false
        }
        refusalLogged = false
        var on: Int32 = 1
        setsockopt(candidate, SOL_SOCKET, SO_NOSIGPIPE, &on, socklen_t(MemoryLayout<Int32>.size))
        socket = candidate
        InputLog.write("link connected")
        let reader = Thread { [weak self] in self?.read(candidate) }
        reader.name = "VisualTeX link"
        reader.start()
        writeLine(["t": "hello", "version": Bundle.main.infoDictionary?["CFBundleShortVersionString"] ?? "?", "selected": SourceState.selected], to: candidate)
        return true
    }

    /// Sends a message when connected (connecting first if needed).
    func send(_ message: [String: Any]) {
        guard connect() else { return }
        lock.lock()
        let target = socket
        lock.unlock()
        guard target >= 0 else { return }
        writeLine(message, to: target)
    }

    private func writeLine(_ message: [String: Any], to target: Int32) {
        guard var data = try? JSONSerialization.data(withJSONObject: message) else { return }
        data.append(0x0a)
        writer.async { [weak self] in
            let ok = data.withUnsafeBytes { buffer -> Bool in
                var offset = 0
                while offset < buffer.count {
                    let written = Darwin.write(target, buffer.baseAddress! + offset, buffer.count - offset)
                    if written <= 0 { return false }
                    offset += written
                }
                return true
            }
            if !ok { self?.drop(target) }
        }
    }

    private func read(_ source: Int32) {
        var pending = Data()
        var buffer = [UInt8](repeating: 0, count: 64 * 1024)
        while true {
            let count = Darwin.read(source, &buffer, buffer.count)
            if count <= 0 { break }
            pending.append(buffer, count: count)
            while let newline = pending.firstIndex(of: 0x0a) {
                let line = pending.subdata(in: pending.startIndex..<newline)
                pending.removeSubrange(pending.startIndex...newline)
                guard let message = (try? JSONSerialization.jsonObject(with: line)) as? [String: Any] else { continue }
                DispatchQueue.main.async { [weak self] in self?.onMessage?(message) }
            }
        }
        drop(source)
        DispatchQueue.main.async { [weak self] in self?.onMessage?(["t": "disconnected"]) }
    }

    private func drop(_ target: Int32) {
        lock.lock()
        defer { lock.unlock() }
        guard socket == target else { return }
        close(target)
        socket = -1
        InputLog.write("link closed")
    }
}

/// The peer of a Unix socket is VisualTeX: signed as com.visualtex.studio and,
/// when this input method carries a certificate, by the same certificate.
enum PeerCheck {
    private static let appIdentifier = "com.visualtex.studio"

    static func isVisualTeX(_ socket: Int32) -> Bool {
        var token = audit_token_t()
        var length = socklen_t(MemoryLayout<audit_token_t>.size)
        guard getsockopt(socket, SOL_LOCAL, LOCAL_PEERTOKEN, &token, &length) == 0 else { return false }
        let tokenData = withUnsafeBytes(of: &token) { Data($0) }
        var code: SecCode?
        let attributes = [kSecGuestAttributeAudit: tokenData] as CFDictionary
        guard SecCodeCopyGuestWithAttributes(nil, attributes, [], &code) == errSecSuccess, let code else { return false }
        var requirementText = "identifier \"\(appIdentifier)\""
        if let leaf = ownLeafCertificateHash() { requirementText += " and certificate leaf = H\"\(leaf)\"" }
        var requirement: SecRequirement?
        guard SecRequirementCreateWithString(requirementText as CFString, [], &requirement) == errSecSuccess,
              let requirement else { return false }
        return SecCodeCheckValidity(code, [], requirement) == errSecSuccess
    }

    /// The SHA-1 of this process's signing certificate, nil when ad-hoc signed.
    private static func ownLeafCertificateHash() -> String? {
        var me: SecCode?
        guard SecCodeCopySelf([], &me) == errSecSuccess, let me else { return nil }
        var staticCode: SecStaticCode?
        guard SecCodeCopyStaticCode(me, [], &staticCode) == errSecSuccess, let staticCode else { return nil }
        var information: CFDictionary?
        guard SecCodeCopySigningInformation(staticCode, SecCSFlags(rawValue: kSecCSSigningInformation), &information) == errSecSuccess,
              let dictionary = information as? [String: Any],
              let certificates = dictionary[kSecCodeInfoCertificates as String] as? [SecCertificate],
              let leaf = certificates.first
        else { return nil }
        let der = SecCertificateCopyData(leaf) as Data
        return Self.sha1Hex(der)
    }

    private static func sha1Hex(_ data: Data) -> String {
        Insecure.SHA1.hash(data: data).map { String(format: "%02X", $0) }.joined()
    }
}
