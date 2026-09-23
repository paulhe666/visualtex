# CP06 browser input checkpoint 1

Date: 2026-09-13. Product baseline remains 00b48e17b736c7e57a0a72e19e271df82906f89d. This directory is a non-Git source checkpoint.

Actual execution: editor-browser-review.mjs passed using an isolated Edge browser. It starts with an empty core document, clicks real toolbar controls, inserts text and scripts through browser input, commits Chinese through CDP composition, navigates fraction arguments, clicks into a Chinese grapheme, tests candidate Enter, undo/redo and grapheme-safe shift selection/deletion, then exports vector SVG and verifies pointer queries remain valid after export.

Final source: \\frac{x_{1}+\\text{速度}}{y^{2}}.
44 captured browser events: 41 marked trusted; 3 compositionend events marked untrusted by this Edge/CDP path. These flags are recorded, not rewritten. This is NOT physical Windows IME acceptance.

Vector export SHA256: b9cbb2cfceeac292666dcd5f144cda62ce9a26dd1726050423196374f16ec561.
Web WASM SHA256 during input test: e3f4d26c1ddf651b56562285299778d1f7a421f1d591d6e05a810be348c6e28b.
Artifacts: engine/target/editor-browser-review/report.json and editor-input.png. Earlier failures are retained in failure.json.

The runtime, SVG view and BrowserInputBridge are real modules. editor-host.ts is the independently tested host integration; editor-controller.ts is being developed concurrently and the two must be consolidated before a product entry point is selected. Production app, Office, macOS, and Web platform entry points have not been switched.
