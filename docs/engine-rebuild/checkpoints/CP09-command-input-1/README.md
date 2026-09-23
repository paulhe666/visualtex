# CP09 — self-owned transient command input, checkpoint 1

Captured 2026-09-14 +08:00. Checkout `reference-3d207d7` remains at local HEAD
`00b48e17b736c7e57a0a72e19e271df82906f89d`. This is a non-Git engine/sandbox
checkpoint. No production main-app switch, Office action, commit or push occurred.

## Status

The first P09 slice is complete: command draft state is owned by the new host and
is independent of MathLive's command popover/offset state. Command display text and
confirmed structural action are separate values.

Stable qualifying receipt:
`engine/target/p09-verification/2026-09-14T00-40-23-857Z/report.json`

Source SHA256:
`63b9bd5313c178bdd54d31665745896e5468d0c814ee1daeec4a3d011e855093`

The qualifying run first executed the complete P08 regression gate, including
physical Windows IME and 1000-formula virtualization, then command inventory, pure
command-state contract and real Edge keyboard acceptance. All four steps PASS and
`sourceUnchanged=true`.

## Command action separation

`engine/web/host/command-input.ts` introduces CommandDescriptor and
CommandInputState. A descriptor owns an explicit core EditCommand action rather
than deriving behavior from a LaTeX trigger at confirmation time.

The existing Windows registry has a concrete ambiguity: its item `partial` exposes
`command="\\partial"` while `insertTemplate` is a complete partial-derivative
fraction. The self-owned seed deliberately separates them:

- exact `\\partial` -> `insertSymbol("\\partial")`;
- `partialderivative` -> explicit one-transaction fraction template.

Exact matching outranks prefix/personalization, so usage frequency cannot turn an
exact `\\partial` into the derivative template. A user can explicitly move to the
other candidate and confirm it.

## Transient state contract

Beginning a command captures documentId, revision, selectionEpoch and Selection.
Typing command characters, candidate movement, Backspace and Escape do not write
source/revision/history. Space, Enter or Tab confirms one candidate and produces
one normal core transaction. A stale target is cancelled and cannot apply at a new
caret/formula. Composition start and focus loss cancel the transient command state.

The popup is a projection only. It is not a mathematical state owner and does not
clone or observe a third-party suggestion DOM.

## Inventory evidence

`engine/tools/p09-command-inventory.mjs` reads the static Windows built-in registry
without reading user localStorage/custom-symbol state and analyzes every materialized
template through the current WASM parser.

Qualifying inventory:
`engine/target/p09-command-inventory/2026-09-14T00-42-08-971Z/report.json`

- built-in registry entries: 411;
- materialized templates currently parser-supported: 358;
- direct supported command/template atoms: 179;
- normalized trigger collision groups: 40;
- explicit unsupported entries are retained as later P09 work, not silently accepted.

The remaining unsupported set includes compatibility macros such as `\\pdv`,
`\\dv`, `\\abs`, `\\norm`, Dirac helpers, delimiter variants and matrix environment
variants. They are not claimed complete by this checkpoint.

## Actual acceptance

Pure state receipt:
`engine/target/command-input-contract-review/2026-09-14T00-42-09-892Z/report.json`

Seven contract groups PASS: transient candidate movement, exact partial action,
separate derivative template, Backspace/cancel, stale target refusal, bounded usage
ranking and malformed/duplicate descriptor rejection. Document writes = 0.

Real Edge receipt:
`engine/target/command-input-browser-review/2026-09-14T00-42-10-716Z/report.json`

- Edge 152.0.4191.62;
- trusted keyboard path exercised;
- `\\partial` draft leaves complete EditorSnapshot unchanged;
- Space confirms one `\\partial` symbol transaction;
- Enter confirms explicit derivative template without outer Enter;
- Escape/Backspace produce no history;
- Arrow selection + Tab explicitly selects another candidate;
- source owner replacement cancels a stale draft;
- browser composition takes priority and cancels the command draft;
- independent frame B remains unchanged.

The screenshot was read and visually inspected after the run; no candidate popup or
transient command state remained after cancellation/commit.

## Preserved failures

Earlier runner failures remain in their unique target directories. One was caused
by loading generated ESM as CommonJS; another was an incorrect test expectation
that usage should beat an exact alias. Neither was deleted or reported as engine
success. The final state runner loads the exact compiled module as ESM and preserves
exact-match precedence.

## Phase boundary

This checkpoint does not claim all 411 commands are mapped, user-defined macros,
custom-symbol asset integration, rich structured clipboard v2, OCR target revision
arbitration, production CodeMirror/main-app integration, P10 default switching or
Office output. Those remain subsequent P09/P10 phases.

P00 remains red because of unrelated parallel Office/evidence changes. Those files
were not modified, reset, cleaned, staged or absorbed by this work.
