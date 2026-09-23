# P05 active implementation coordination

2026-09-13. Product baseline remains 00b48e1; no engine commit/push authorized.

P05 fonts/layout/scene and protocol/CLI/WASM resource injection were added in this lane. Initial 4 font tests and 12 layout tests passed. Review lane tests have exposed additional empty-formula and vertical-navigation issues. Concurrent current files now already include scene_navigation/scene_validation and has_formula_content fixes; attempted exact-text patches that raced these edits failed without changing files. Do not replace these files from older reads.

This lane is now concentrating on `engine/tools/layout-parity.mjs`, the TS geometry API and independent Native/WASM layout validation. Please avoid writing that parity runner concurrently. Core review fixes may proceed in their small named modules; preserve all existing assertions. A fresh full build will run after those corrections stabilize. We will archive verified combined source, with explicit test counts and resource hashes; no false attribution of concurrent fixes.

Fonts are loaded read-only. Latin Modern Math uses the existing repo font artifact (SHA256 6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f). Chinese tests use the installed NotoSansSC-VF.ttf default coordinates and record its hash. No system font is copied into source or archives. No production application, Office document, installer, Git commit or remote is changed.
