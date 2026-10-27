import Foundation

/// The pinyin engine: librime with VisualTeX's pinyin schema. Its shared
/// data (schema, dictionaries, prebuilt tables) is in the input method's
/// bundle; what it learns (the user dictionary) goes to
/// ~/Library/Application Support/VisualTeX/Rime.
final class RimeEngine {
    static let shared = RimeEngine()

    struct Candidate {
        let text: String
        let comment: String
    }

    /// What one session shows: the spelling being composed, its caret and
    /// selection (UTF-16 offsets), and the candidate page.
    struct State {
        var preedit = ""
        var cursor = 0
        var selection = NSRange(location: 0, length: 0)
        var candidates: [Candidate] = []
        var labels: [String] = []
        var highlighted = 0
        var pageNumber = 0
        var isLastPage = true
        var isComposing: Bool { !preedit.isEmpty }
    }

    private var api: UnsafeMutablePointer<RimeApi>?
    private(set) var isReady = false
    /// C strings handed to librime (it keeps the pointers).
    private var retained: [UnsafeMutablePointer<CChar>] = []
    private let queue = DispatchQueue(label: "com.visualtex.inputmethod.rime")

    private func keep(_ text: String) -> UnsafePointer<CChar> {
        let copy = strdup(text)!
        retained.append(copy)
        return UnsafePointer(copy)
    }

    /// Sets librime up and installs the prebuilt data (in the background:
    /// keys typed before it is ready are not taken).
    func start() {
        guard api == nil, let rime = rime_get_api() else { return }
        api = rime
        let resources = Bundle.main.sharedSupportURL?.appendingPathComponent("rime").path
            ?? Bundle.main.bundlePath + "/Contents/SharedSupport/rime"
        let user = FileManager.default.homeDirectoryForCurrentUser
            .appendingPathComponent("Library/Application Support/VisualTeX/Rime")
        try? FileManager.default.createDirectory(at: user, withIntermediateDirectories: true)
        var traits = RimeTraits()
        traits.data_size = Int32(MemoryLayout<RimeTraits>.size - MemoryLayout<Int32>.size)
        traits.shared_data_dir = keep(resources)
        traits.user_data_dir = keep(user.path)
        traits.prebuilt_data_dir = keep(resources + "/build")
        traits.staging_dir = keep(user.appendingPathComponent("build").path)
        traits.distribution_name = keep("VisualTeX")
        traits.distribution_code_name = keep("visualtex")
        traits.distribution_version = keep(Bundle.main.infoDictionary?["CFBundleShortVersionString"] as? String ?? "1")
        traits.app_name = keep("rime.visualtex")
        traits.min_log_level = 2
        traits.log_dir = keep("")
        rime.pointee.setup?(&traits)
        rime.pointee.initialize?(&traits)
        let started = Date()
        queue.async { [weak self] in
            if rime.pointee.start_maintenance?(0) != 0 {
                rime.pointee.join_maintenance_thread?()
            }
            DispatchQueue.main.async {
                self?.isReady = true
                InputLog.write(String(format: "rime ready in %.0f ms", Date().timeIntervalSince(started) * 1000))
            }
        }
    }

    func createSession() -> RimeSessionId {
        guard isReady, let api else { return 0 }
        return api.pointee.create_session?() ?? 0
    }

    func destroySession(_ session: RimeSessionId) {
        guard session != 0, let api else { return }
        _ = api.pointee.destroy_session?(session)
    }

    /// True when librime took the key.
    func process(_ session: RimeSessionId, keysym: Int32, mask: Int32) -> Bool {
        guard session != 0, let api else { return false }
        return api.pointee.process_key?(session, keysym, mask) != 0
    }

    /// The text librime committed with the last key, if any.
    func takeCommit(_ session: RimeSessionId) -> String? {
        guard session != 0, let api else { return nil }
        var commit = RimeCommit()
        commit.data_size = Int32(MemoryLayout<RimeCommit>.size - MemoryLayout<Int32>.size)
        guard api.pointee.get_commit?(session, &commit) != 0 else { return nil }
        defer { _ = api.pointee.free_commit?(&commit) }
        return commit.text.map { String(cString: $0) }
    }

    func state(_ session: RimeSessionId) -> State {
        var state = State()
        guard session != 0, let api else { return state }
        var context = RimeContext()
        context.data_size = Int32(MemoryLayout<RimeContext>.size - MemoryLayout<Int32>.size)
        guard api.pointee.get_context?(session, &context) != 0 else { return state }
        defer { _ = api.pointee.free_context?(&context) }
        if let preedit = context.composition.preedit {
            let bytes = Array(UnsafeBufferPointer(start: UnsafeRawPointer(preedit).assumingMemoryBound(to: UInt8.self), count: Int(context.composition.length)))
            state.preedit = String(decoding: bytes, as: UTF8.self)
            let utf16 = { (byteOffset: Int32) -> Int in
                let offset = max(0, min(Int(byteOffset), bytes.count))
                return String(decoding: bytes[0..<offset], as: UTF8.self).utf16.count
            }
            state.cursor = utf16(context.composition.cursor_pos)
            let start = utf16(context.composition.sel_start)
            let end = utf16(context.composition.sel_end)
            state.selection = NSRange(location: start, length: max(0, end - start))
        }
        let menu = context.menu
        if let candidates = menu.candidates, menu.num_candidates > 0 {
            for index in 0..<Int(menu.num_candidates) {
                let candidate = candidates[index]
                state.candidates.append(Candidate(
                    text: candidate.text.map { String(cString: $0) } ?? "",
                    comment: candidate.comment.map { String(cString: $0) } ?? ""
                ))
            }
        }
        if let keys = menu.select_keys {
            state.labels = String(cString: keys).map { String($0) }
        }
        state.highlighted = Int(menu.highlighted_candidate_index)
        state.pageNumber = Int(menu.page_no)
        state.isLastPage = menu.is_last_page != 0
        return state
    }

    func isAsciiMode(_ session: RimeSessionId) -> Bool {
        guard session != 0, let api else { return false }
        return api.pointee.get_option?(session, "ascii_mode") != 0
    }

    func setAsciiMode(_ session: RimeSessionId, _ on: Bool) {
        guard session != 0, let api else { return }
        api.pointee.set_option?(session, "ascii_mode", on ? 1 : 0)
    }

    func clear(_ session: RimeSessionId) {
        guard session != 0, let api else { return }
        api.pointee.clear_composition?(session)
    }

    /// Commits what is being composed (the field lost the composition).
    func commitComposition(_ session: RimeSessionId) -> String? {
        guard session != 0, let api else { return nil }
        guard api.pointee.commit_composition?(session) != 0 else { return nil }
        return takeCommit(session)
    }
}
