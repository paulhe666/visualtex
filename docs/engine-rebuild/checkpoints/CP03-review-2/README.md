# CP03 review checkpoint 2 — parser/session parity and host foundations

Date: 2026-09-13. Local baseline: 00b48e17b736c7e57a0a72e19e271df82906f89d. Engine changes remain uncommitted. This checkpoint does not replace CP02 or CP03-review-1.

## Verified execution

`node tools/build.mjs` in engine: 61 Rust tests passed (13 core, 15 implementation parser, 8 model-review, 6 parser-review, 8 session-review, 9 source-review, 2 CLI). Actual Native/WASM protocol replay: 2151 cases. Actual Native/WASM parsed-source session replay: 222 cases.

`node tools/syntax-parity-review.mjs`: 305 additional actual Native/WASM full-JSON comparisons passed, including 41 independent K0 cases, 256 seeded generated formulas, resource/schema failures, and recovery. The stateless analysis route is explicitly separate from EngineSession ownership; analysis-local IDs cannot become editing anchors.

Built Native SHA256: b9e8eaaf70929eb7af35beb71d8d40f215d40c2455890399f9097e6098ae274e.
Built WASM SHA256: ace83bf16a4be3d86e787e7086e11c21320311f590d42488024876d59a64c10c.

The latest build has a dead-code warning for the not-yet-connected parse_braced_fragment helper; no tests were suppressed. This is work in progress toward local source edits, not a claim that P04 is complete.

## Host units, not product acceptance

New coordinates.ts validates CSS client/layout mappings, excludes devicePixelRatio from semantic geometry, and guards document/revision/layout-generation identity. New resources.ts verifies fixed manifest build/protocol, resource URL confinement, streamed length, MIME and SHA256; bounds cache/concurrency; supports cancellation, timeout and disposal without publishing stale assets.

Strict TypeScript compile and `node --test tools/host-infrastructure-review.test.mjs` passed 21 host-unit tests. Test byte assets are synthetic and do not demonstrate real font parsing, mathematical layout, UI input or system IME. A real empty WASM module was compiled only to validate the loader's byte handoff.

## Protection and limits

Protection check at 2026-09-13T04:29:26.758Z passed all 12 groups, 12 artifacts and 65 original dirty files. Production application, Office source/contracts/installation/documents and default MathLive renderer remain unchanged.

Not accepted yet: local source/structural transactions with stable identities and history; MATH font layout; SVG edit view; real Chinese IME; any Office UI or final installer. P03 local-identity completion and all later phase gates remain explicit, not inferred from syntax tests.

The archive generated in this directory is an immutable source snapshot, not a Git commit. Its hash and source-stability result are recorded in audit.json. No fonts, private credentials, user document data, build cache or installed binaries are included.
