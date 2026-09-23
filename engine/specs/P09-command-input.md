# P09 command input — transient draft and explicit structural actions

Status: first isolated P09 slice. This specification defines the self-owned command
input state used by the new engine host. It does not switch the production main app
and does not replace the existing VisualTeX command registry yet.

## Separation of display trigger and edit action

A command candidate is not a LaTeX string with implicit behavior. Each descriptor
has a stable id, trigger/aliases/ranking metadata and one explicit core EditCommand
action. The displayed trigger cannot decide at confirmation time whether a symbol
or template is inserted.

This is required by the existing Windows registry: its `partial` item currently has
`command="\\partial"` but an `insertTemplate` containing a full partial-derivative
fraction. The self-owned seed therefore defines:

- exact `\\partial` -> `InsertSymbol("\\partial")`;
- explicit `partialderivative` -> one template/paste transaction.

Typing the exact symbol name cannot silently become the fraction template because
of usage frequency or legacy registry shape. A user may explicitly move to another
candidate and confirm it.

## Transient ownership

CommandInputState is host transient state, not another mathematical document store.
Beginning a command captures documentId, revision, selectionEpoch and Selection.
Typing ASCII command characters, changing the selected candidate, Backspace or
Escape never writes source, revision or history.

Space, Enter or Tab confirms the selected candidate and produces exactly one normal
core edit transaction. The state is cleared before the edit is dispatched. A stale
target (revision, selection epoch or selection changed) is cancelled and cannot be
confirmed into another caret/formula.

IME composition has priority. Composition start cancels the command draft before
core composition begins. Focus loss and explicit source-owner replacement also
cancel the draft. No unfinished command can be parked as committed formula content.

## Ranking

The isolated state supports exact, prefix and alias matching plus injected usage
metadata. Exact text matches outrank prefix/personalization so history cannot hijack
a precise command spelling. Prefix matches may be reordered by bounded frequency,
accepted-prefix, recency and pin scores. Descriptor inputs are immutable.

Production P10/P09 adapters will map the existing registry's aliases, usage and
custom-symbol entries into this schema. The new host must not depend on MathLive's
popover DOM, candidate node classes or command-offset representation.

## Current inventory evidence

`tools/p09-command-inventory.mjs` reads the static built-in Windows registry without
reading localStorage/custom-symbol user state and analyzes every materialized
insert template through the current WASM parser.

At the first inventory point the registry contains 411 entries, 358 materialized
templates are currently supported by the self-owned parser, 179 entries are direct
supported atoms/templates where the command itself equals the template, and 40
normalized trigger collisions exist. The remaining entries are explicit P09 work,
not silently accepted.

## Acceptance

Pure state acceptance must prove: exact partial action, separate derivative action,
Backspace/cancel, stale target refusal, prefix usage ranking, immutable descriptors
and invalid-descriptor rejection.

Real Edge acceptance must use trusted keyboard input and prove:

- the complete command draft leaves source/revision/history unchanged;
- Space confirms exact partial as one symbol transaction;
- Enter confirms the explicit derivative template without reaching outer Enter;
- Escape and Backspace leave history unchanged;
- Arrow selection + Tab can explicitly select a different candidate;
- source owner replacement cancels a stale draft;
- browser composition cancels the draft and takes input priority;
- another formula frame remains byte-for-byte unchanged.

This slice does not yet claim complete mapping of all 411 commands, user macros,
custom-symbol assets, P09 versioned rich clipboard, OCR target arbitration or
production command popup styling. Those are subsequent P09 slices.
