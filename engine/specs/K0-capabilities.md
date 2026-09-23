# K0 language support at P03

Status: independent source/model/parser and source-edit protocol implemented. This is not a complete LaTeX language, interactive editor, layout engine, or Office acceptance report.

| Feature | Parse / preserve | Source edits / identities | Visual edit | Layout / SVG / MathML | Office |
|---|---|---|---|---|---|
| Row, braced groups, ordinary symbols/digits | PASS | PASS | P04 pending | P05/P11 pending | NOT_RUN |
| Fraction: frac, dfrac, tfrac | PASS, command spelling retained | Closed argument rows parsed locally | P04 pending | P05/P11 pending | NOT_RUN |
| Superscript and subscript | PASS, argument order retained | PASS; duplicate scripts diagnosed | P04 pending | P05/P11 pending | NOT_RUN |
| Text, Chinese, emoji, combining clusters | PASS, explicit Unicode offsets | PASS, no split UTF-8/UTF-16 boundaries | P04/P06 pending | Font shaping pending | NOT_RUN |
| Named Greek/basic operator/relation symbols | Bounded vocabulary in symbols.rs | PASS | P04 pending | P05 pending | NOT_RUN |
| Empty arguments and labeled placeholder | Preserved, not export-ready | PASS | P04 pending | UI only, never normal export | NOT_RUN |
| Unknown commands with star/optional/braced args | Opaque original source + unsupported diagnostic | PASS, not reinterpreted | Opaque-node behavior pending | NO silent fallback | NOT_RUN |
| Missing arguments/braces/unfinished command | Exact draft + diagnostic | PASS, explicit retained older tree | Draft arbitration pending | No repaired export | NOT_RUN |
| Radicals, stretchy fences, accents, large operators, math styles | Not claimed | Original source retained when unsupported | P07 pending | P07 pending | NOT_RUN |
| Matrix/cases/align/multiline environments | Not claimed | Opaque source preservation only | P08 pending | P08 pending | NOT_RUN |
| User macros, custom symbols, rich clipboard | Not claimed | User source must remain intact | P09 pending | P09 pending | NOT_RUN |

Source-edit position uses UTF-16 code-unit offsets explicitly converted to valid scalar UTF-8 boundaries. Mathematical caret positions are a different type and will be implemented in P04; source offset is never a MathLive position.

## Source editing strategy

The source is fully lexed to prove that comments or control words did not cross a proposed parse boundary. A smallest existing closed braced Row is reparsed only when its braces remain token-aligned, the parser consumes that entire group, and it has no incomplete syntax. Otherwise the whole formula is reparsed. The returned SourceChangeMap reports the actual scope, parsed node count and lexed token count. Global tree validation remains in place; no claim of wholly local time complexity is made.

Identity reconciliation matches transformed source ranges, node kinds and structural roles. Each old identity can be reused at most once. Exact replaced atoms receive fresh identities; unaffected repeated siblings keep their own distinct identities. New identities are monotonic. Explicit replaceSource is a full replacement and intentionally assigns fresh inner identities. No-op source writes preserve all identities and revision.

## Validation limits

Source bytes: 262144. Tokens: 32768. Nodes: 8192. Tree/parser depth: 64. Diagnostics: 128. Wire request: 1048576 bytes. Active sessions: 64. Exceeding budgets is a structured error and leaves committed session state unchanged. No macro expansion or external TeX commands execute.

## Distinct outputs

Lossless serialization preserves unchanged source slices and their trivia. Canonical serialization is a separate explicit output and is not offered for unsupported/incomplete input. Valid but empty slots are still not export-ready. RetainedTree in an incomplete source snapshot carries its older source identity and must never be mistaken for the parse of the current draft.
