# CP04 — resumed structural editor checkpoint

Date: 2026-09-13. Local product baseline remains `00b48e17b736c7e57a0a72e19e271df82906f89d`. This checkpoint is an uncommitted source archive, NOT a Git commit or release.

## Executed verification

- `cargo test --locked -p visualtex-engine-core --test transactions_contract_review -- --nocapture`: 14/14 passed, including the previously reported transient text-split PositionMap regression and the seeded command sequence test.
- `node engine/tools/build.mjs` (engine working directory): workspace tests passed (157 core tests + 2 CLI tests); native CLI and wasm32 build/bindings passed. Native/WASM comparison: 2,151 protocol requests, 222 source-session cases, 440 source-edit requests, 473 editor requests including 100 history states.
- Native binary SHA256: `7df370f26df7585f3253642812f98f1e367ecbb3bef59b8cba3b6f095be800d8`.
- WASM SHA256: `e0239423801f0f612ffb5efa5670d69e551f258414212d661b15a5767e63c374`.
- Protection check with the recorded P00 Git transition: PASS (12 groups, 12 artifacts, 65 original dirty paths).
- Office contract check: PASS (11 sources, 3 interfaces, 10 routes, 3 negative mutation tests; no writes).

## Scope and limits

Structure navigation, selection, grapheme-safe edits, fresh identities on copy, bounded shared source/visual history, event replay protection and composition state are implemented in the isolated core. Source projection and transaction blocker regressions currently pass. The earlier blocker reports are preserved as historical evidence, not deleted or rewritten.

This does NOT establish real OS IME, rendered caret hit testing, font layout, SVG, Office UI, or complete replacement of MathLive. K0 is the supported language boundary. SetStyle outside K0 is explicitly rejected. Vertical movement returns a request for the forthcoming geometry layer; it does not guess from source offsets.

Next: implement P05 fixed-resource font service, K0 mathematical layout and SVG scene generation within `engine/` only. Production app, Word/PowerPoint, installed packages, macOS and Web platform folders remain untouched. The user has forbidden Git commit/push for this rebuild.
