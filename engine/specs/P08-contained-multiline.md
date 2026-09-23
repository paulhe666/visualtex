# P08 — contained multiline formula decision and acceptance contract

User decision recorded 2026-09-13, before P08 implementation. This supplements
`docs/engine-rebuild/VISUALTEX_ENGINE_PROJECT_PLAN.md`, particularly P08 items 1,
4 and 5. It is a requirement, not a report of completed functionality.

## Product decision

A multiline mathematical environment belongs to exactly ONE formula editor frame,
one engine document/session and one host formula identity. Adjacent outer formula
frames are independent formulas. Never create an aligned group by borrowing,
joining or rearranging adjacent formula frames.

The default dedicated shortcut is **Shift+Enter**. Do not repurpose an Office
Apply shortcut or send the event to the outer editor's ordinary Enter callback.
Keep the existing ordinary Enter behavior separate; inside an existing cases or
matrix environment its established environment-row behavior remains applicable.
Keyboard configuration may change the binding later without changing the command.

For an ordinary single formula, Shift+Enter promotes its root into a contained
multiline environment and inserts a row at the active mathematical boundary as ONE
atomic transaction. Use `gathered` as the initial general-purpose environment (no
invented alignment markers); an explicit aligned/cases/matrix choice retains its
own schema and geometry. In an existing supported multiline environment, insert
within the nearest applicable environment instead of nesting another wrapper.

At an end-of-formula caret the old content is the first row and the second row is
empty and editable. At a boundary within a root row split its entries into left
and right rows, preserving their source spelling and node identities. A caret
inside a fraction, script, root, accent or text run must not slice raw LaTeX or
silently remove that enclosing structure: resolve an explicit structural split
policy at P08 implementation time and test it. Until that policy is implemented,
reject unsupported interior splits atomically and keep source/selection unchanged.

A completed example at the end of the first line is:

```latex
\begin{gathered}
x+y=1 \\
x-y=0
\end{gathered}
```

All of this is inside one frame. This is NOT two outer document rows synchronized
by CSS or two Mathfield/EngineEditorHost instances pretending to be one formula.

## Ownership, history and source invariants

- Preserve outer documentId/session handle/formulaId/frame identity. Only internal
  environment/row/cell identities are created as required; copied cells are fresh.
- Promotion and the initial break are one undo item. Undo restores original source,
  structure, selection and identities; redo restores the same internal structure.
- Source and visual views share the same owner and history. A whole begin/end
  environment round-trips in one outer formula record, even when its source contains
  physical newlines. Never split the document by a raw newline regex.
- Backspace/Delete and row merge operate within the active environment. They must
  not join the preceding/following formula frame. An empty environment row stays
  editable and cannot cause an automatic Office commit or identity reassignment.
- IME composition and unfinished command confirmation have priority. Shift+Enter
  during unconfirmed preedit cannot insert an extra environment/document row.
- Selection spanning different frames is not an environment operation. Do not
  resolve it by reading or editing another EngineSession.
- Geometry, caret navigation, SVG export and later MathML output use the same
  environment tree and layout anchors. One contained multiline formula exports as
  one object. Existing outer numbering/Office contracts are not changed here.

## Required actual acceptance (all initially NOT_RUN)

M01: From blank, type `x+y=1`, Shift+Enter, type `x-y=0`; one frame/one session,
     two internal rows, matching source/geometry. No outer Enter callback.
M02: Repeat Shift+Enter to build at least three rows; single-owner history.
M03: Promote at a middle root-row boundary; source and identities preserved.
M04: Nested fraction/root/script/interior text boundary; documented structural
     behavior or atomic refusal, never damaged LaTeX or another frame.
M05: Existing gathered/aligned/cases/matrix: nearest-environment insertion,
     correct cell navigation/alignment and no redundant nested environment.
M06: Two independently mounted formula frames A/B. Shift+Enter in A leaves B's
     source, revision, selection, history and identity byte-for-byte unchanged.
M07: IME candidate confirmation/cancellation and pending command: no duplicate row.
M08: Undo first promotion and redo it; formula identity and internal IDs restored.
M09: Copy/paste/save/reopen complete multiline source; one frame, independent copy.
M10: Empty row deletion/merge, first/last row boundaries: never cross frames.
M11: Ordinary Enter and configured host Apply shortcuts retain separate behavior.
M12: Native/WASM parity, trusted browser keyboard/pointer, zoom/DPR and natural
     vector export bounds; later real Office acceptance remains a separate gate.

Implementation must follow the P07 -> P08 phase dependency. Writing this decision
is not evidence that environment nodes, keyboard routing, rendering or any M-test
already pass. Production UI remains unchanged during isolated engine development.
