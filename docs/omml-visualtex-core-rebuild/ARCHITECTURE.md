# OMML / VisualTeX Word Host Core Rebuild

## Scope

This rebuild covers every Windows Word path that reads, inserts, edits, deletes,
copies, pastes, redraws, numbers, references, converts, validates, rolls back,
or batch-processes an OMML or VisualTeX host.

MathType's MTEF/native OLE codec and renderer remain outside the rebuild.
Any route from or to MathType that touches an OMML or VisualTeX host must use
the new host core on the OMML/VisualTeX side.

## Non-negotiable invariants

1. A local operation never performs a document-wide scan.
2. A batch may build one immutable document index before mutation. It must not
   re-enumerate Bookmarks, OMaths, InlineShapes, CustomXMLParts, or WordOpenXML
   once per formula.
3. OMML/VisualTeX host identity is location only. Identity data must never be
   treated as proof that semantic content is correct.
4. Semantic validation is local to the formula being mutated.
5. Numbering is a container concern applied after the formula host succeeds.
   Numbering does not decide whether the formula host conversion succeeded.
6. No invisible or zero-width document character may be inserted merely to
   preserve a caret or typing state. Selection/caret state is transient state,
   not document content.
7. Insert, edit, redraw, format conversion, copy/paste repair, and bulk import
   must all use the same host Insert/Delete/Replace primitives.
8. A replacement has one explicit mutation boundary and one rollback boundary.
9. A failed local mutation may restore only the local owned region and owned
   identity/numbering metadata. It must not "repair" unrelated formulas.
10. User tables, paragraph formatting, run formatting, body REF fields, section
    fields, and unrelated bookmarks are foreign state and are immutable unless
    explicitly part of the requested operation.
11. Save/reopen is part of correctness. In-memory COM state alone is not proof.
12. Performance optimizations may alter scheduling only; they may not create a
    second implementation of host semantics.

## New architecture

### WordFormulaHostDescriptor

Immutable snapshot of one formula host:
- host kind: OMML / VisualTeX / MathType
- story/range coordinates
- inline/display mode
- FormulaId when owned by VisualTeX
- semantic metadata
- numbering descriptor
- paragraph/table ownership
- source MathML/LaTeX when already available

It contains no live COM object.

### WordFormulaHostResolver

Two modes only:
- local: resolve one host from a supplied Selection/Range with a bounded probe.
- batch: enumerate the requested source kind once and build an immutable index.

There is no local-to-global fallback.

### WordFormulaDocumentIndex

One-operation immutable index of:
- OMML formula ranges
- VisualTeX OLE ranges and FormulaIds
- VisualTeX-owned identity bookmarks
- VisualTeX metadata parts
- numbering owners
- user-table ownership

The index is captured once before a batch and invalid after the first mutation.

### WordFormulaHostMutationKernel

The only layer allowed to structurally mutate an OMML or VisualTeX formula.

Primitives:
- InsertHost
- DeleteHost
- ReplaceHost

ReplaceHost is always:
capture local boundary -> delete exact source -> insert exact target at retained
anchor -> apply target-local formatting -> apply numbering container -> validate
target locally -> commit.

### WordFormulaIdentityStore

Owns only VisualTeX-created identity:
- VisualTeX OLE identity
- managed OMML identity
- semantic metadata part

It must never enumerate the full document during a local operation.

### WordFormulaNumberingKernel

Owns numbering topology and equation references. Host mutation sees numbering as
one descriptor rather than VTEq/VTEqCap/VTEqNum/VTEqAnc/VTEqShape implementation
details.

The new canonical display topology must have one implementation per supported
host kind. Legacy topologies are read-only migration inputs, not alternative
production writers.

### WordFormulaMutationValidator

Local postconditions only:
- target host exists at the retained anchor
- target kind matches request
- FormulaId/metadata match when VisualTeX-owned
- semantic payload matches requested payload
- paragraph/table boundary invariants remain unchanged
- numbering descriptor is satisfied when requested

Forbidden success criteria:
- whole-document source/target object counts
- nearest-formula heuristics
- unrelated bookmark cardinality
- unrelated OMML/OLE inventory equality

### WordFormulaMutationTransaction

Captures only the owned mutation region plus owned metadata/numbering state.
Rollback restores that region and verifies foreign state was not touched.

## Migration order

1. host contracts + local resolver + document index
2. VisualTeX insert/update/delete primitive
3. OMML insert/update/delete primitive
4. shared ReplaceHost transaction
5. single insert/edit paths
6. LaTeX redraw
7. OMML/VisualTeX format conversion
8. MathType -> OMML/VisualTeX and OMML/VisualTeX -> MathType host side
9. numbering/un-numbering/references
10. copy/paste
11. bulk import and document-wide conversion
12. delete legacy writers/scanners and compatibility fallbacks

## Acceptance gates

Every stage must pass real Word UI flows before later performance work:
- empty document, one inline formula
- caret inside formula
- exact formula selection
- formula followed by prose
- formula at paragraph end
- display unnumbered
- display numbered
- user table
- adjacent formulas
- copy/paste
- edit/apply
- redraw
- each format-conversion direction touching OMML or VisualTeX
- undo
- save/reopen

Only after correctness gates pass:
- 100 unnumbered
- 100 numbered
- 1000-formula large document

