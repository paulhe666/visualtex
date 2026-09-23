# CP09-language-clipboard-1

Stable non-Git checkpoint for the P09 language-extension and structural-clipboard slices. This checkpoint does not enable the new engine in the production Windows application or Office.

## Stable source

- Verification receipt: `engine/target/p09-verification/2026-09-14T01-59-14-762Z/report.json`
- Stable source SHA256: `31bb923e50dd5bd6727229a7c6199a2c887286135d896d63b5c8525d2b742a0e`
- Stable verification: PASS, 12 runner steps, source unchanged before/after.
- Git branch/HEAD were not modified by this checkpoint. No commit or push is authorized for engine rebuild work.

## Capability covered at this checkpoint

The existing static Windows command registry is machine-audited as 411/411 covered: 406 standalone materialized templates plus 5 contextual script templates that require an existing base. Forty trigger collisions remain explicit and are not collapsed. `\partial` remains a literal symbol action and the derivative template remains a separate command action.

Bounded built-in compatibility macros remain real `MacroCall` nodes with editable arguments and authored-source preservation. Unknown/executable TeX remains fail-closed.

Additional language structures now include:

- `pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix` as one Grid owner with real font-driven wrappers;
- `\middle` as a real Fence child using the owning Fence's stretch target;
- one- and two-column `cases` without invented condition cells;
- bounded `array` l/c/r column specifications and real vector `|` rules, including `\left[\begin{array}{cc|c}...` augmented matrices;
- exact rare-integral aliases `\dashint`, `\ddashint`, `\varointctrclockwise` using the established rare-operator glyph path.

The internal structural clipboard now publishes a versioned, identity-free semantic payload bound to the exact LaTeX fallback by SHA-256. It is validated by Rust core before a normal paste transaction; it is never hydrated directly into a live tree. Browser copy also exposes `application/x-latex` and `text/plain`; legacy custom v1 remains readable. HTML is ignored. Asset-bearing payloads are explicitly refused until the custom-symbol slice defines an asset snapshot contract.

## Actual acceptance evidence

The stable P09 runner executed these steps on one unchanged source set:

1. P08 full regression, including the lower-phase Rust/Native/WASM/browser/IME gates.
2. Command-registry inventory: 411/411 capability coverage.
3. Command input state contract.
4. Command input owned-Edge keyboard acceptance.
5. Macro Native/WASM parity.
6. Macro owned-Edge acceptance.
7. Matrix-family Native/WASM parity.
8. Matrix-family owned-Edge acceptance.
9. `\middle` Native/WASM parity.
10. `\middle` owned-Edge acceptance.
11. Input-bridge contract, including custom-v2/v1/LaTeX/plain fallback and inert HTML.
12. Owned-Edge real clipboard acceptance using trusted Ctrl+C/Ctrl+X/Ctrl+V.

Additional Rust tests executed in the lower-phase build include `array_environment_review` 6/6, `cases_single_column_review` 4/4, `structural_clipboard_review` 5/5, plus all previous workspace tests. Rare-integral parity reports 25 registered rare integral commands, 2654 requests, 1939 scenes, zero numeric difference and exact SVG parity. The structural clipboard browser receipt records custom schema v2, real `application/x-latex`/plain fallbacks, trusted clipboard events, destination Grid/cell identity retention and independent frame B.

## Boundaries

P09 is not complete at this checkpoint. Remaining P09 work includes custom-symbol asset snapshots and self-owned rendering, OCR stale-target insertion, and persistence/history/font/shortcut compatibility closure. CodeMirror remains untouched here; its final production bridge is still outside the isolated sandbox until the appropriate integration stage.

P00 protection remains red because of parallel Office/VSTO/evidence changes that this engine task must not modify. No Word document or Office process was used by this checkpoint.