# P04 K0 transaction contract

Implementation contract, not a completed acceptance report. The engine owns mathematical positions; no MathLive offset or LaTeX string index is used as a visual caret.

## Positions and selections

A Row position is (NodeId, boundary offset, affinity), with boundaries 0..childCount. A Text position is (NodeId, extended-grapheme offset, affinity). Node selection selects a complete subtree. Continuous selection records anchor and focus independently, including reversed selections. Cell rectangles will be added with the P08 matrix model, not emulated with row integers.

Horizontal movement traverses structural children: numerator then denominator, base then script arguments, text graphemes. Enter/exit boundaries are explicit. Vertical movement requests a direction and desired horizontal coordinate for the P05 layout resolver; it must not guess a row from adjacent source characters. Movement and selection updates do not add content history or parse source.

Range deletion is exact inside one Row or TextRun. A range crossing child structures is promoted to whole affected children of the lowest common editable Row; it never cuts half a fraction. Normalization is returned so a host can show the actual selected structure. Backspace/Delete at a structure boundary selects that structure first; another explicit key deletes it. An empty argument navigates out rather than silently flattening nonempty sibling arguments. There is no retry-until-changed delete loop.

## Content commands and state

InsertText edits a TextRun by grapheme, or creates ordinary math atoms at a Row boundary. InsertStructure creates a fraction or script with fresh IDs and selects its first empty argument; a selected range can become the numerator/base. ReplaceRange/paste has one transaction. Text split/merge preserve textual grapheme semantics; multiline Row split and matrix cell operations wait for P08.

Full source replacement and source patches use the same owner/history. Visual transactions update the mathematical tree directly, then generate a lossless source projection with global Unicode-aware origins. They must not parse the complete serialized LaTeX again to discover what they just did. Unaffected tree IDs and source trivia survive. Unsupported or incomplete source requires explicit source resolution before visual mutation; no old renderer takes over.

Each successful content transaction checks expectedRevision, commits atomically, increments revision once, returns a PositionMap, and stores the before/after structure, source, draft, retained tree and selection. Undo restores those values with a fresh monotonically increasing revision; NodeIds/source-buffer identities are never recycled for unrelated objects. Redo is discarded after a new edit. History is bounded by entry and byte budgets. Consecutive typing may share a caller-provided input group only while target, selection progression and history adjacency still match. Paste and composition each have one history entry.

Replayed event IDs must not execute twice. Reusing an event ID with different payload or after an incompatible revision is rejected; no event silently applies to a new active document. Asynchronous command targets include document, revision and selection identity.

Composition captures the original selection and revision. Updating preedit changes composition state only. Commit applies one structural insertion; duplicate commit is detected, cancel creates no content history, stale target is rejected. Pure composition tests do not constitute a real system IME test; P06 must exercise the installed input method.

## Position maps and acceptance

Splice maps include old/new row intervals, text grapheme intervals, retained and removed node IDs, and explicit fallback positions for deleted subtrees. Undo/redo provide maps too; every resulting anchor must validate against the result tree. Source edits map old mathematical positions through retained identities or an explicit enclosing-row fallback, never by assuming UTF-16 equals a visual position.

Tests must cover nested fractions/scripts, reversed and cross-structure selections, duplicate formula copies, emoji/combining deletion, incomplete source history, event replay, composition cancel/commit/staleness, and seeded action/undo/redo sequences. A history round trip checks source, structure, identities and selection—not source text alone.

SetStyle will become a supported content command when the P07 style node and serialization/layout rules are implemented. It must return an explicit unsupported-capability error until then, not silently succeed. The K0 checkpoint must disclose this boundary.
