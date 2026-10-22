# VisualTeX editing core — design and change process

Date: 2026-10-21. Status: design accepted for implementation; phases below are executed in order. Progress: Phases
0–5 done (2026-10-21; Phase 5 split into 5a–5d from the inventory in `audit/phase5-inventory.md`, VPS clone).
Phase 6 step 1 (6a) done: the build-time patch chain is folded into the kernel source (`vendor/mathlive/kernel-parts`);
generated tables fill named data slots at load (`vite.visualtexKernelData.ts`); the built app is byte-identical.
Phase 6 is split into 6a–6f (table in §10). 6b: the kernel is TypeScript source in
`apps/shared/visualtex-core/src/kernel/`; the fragments, the loader and the text slots are gone (verification
below). 6c: the parts VisualTeX never runs are deleted from that source (about 30% of it; list in §10). 6d: the
kernel is strict-clean and part of the app's own type check; the npm `mathlive` package is gone from the Windows app
(Node scripts run the kernel through `scripts/kernel-node.mts`).
Audit this design answers: `~/claude/visualtex-render-profiles/audit/caret-movement-placeholder-audit.md`
(every caret, movement and placeholder behaviour is decided in two to six places today).

## 1. Goal

One module of our own owns everything about where the caret can be, how it moves, what an empty slot is and how
the caret and empty slots are drawn. Every key, click and command goes through it. Nothing else decides these
things: no build-time text patches, no app key interception, no DOM measurement.

The module grows phase by phase until it is the whole editing core (model, parsing, serialising, input); the forked
kernel shrinks to nothing. The saved LaTeX format does not change.

## 2. Ownership after the refactor

| Concern | Owner | Today (removed) |
|---|---|---|
| Caret stops (where the caret may be) | core `stops` | kernel offsets + `isValidPosition` macro rules + click paths that ignore them |
| Order of a structure's parts | core `structures` (one table) | ≥6 orders: NAMED_BRANCHES, VISUALTEX_NAVIGATION_BRANCHES, per-class overrides, hard-coded lists |
| Arrows, word moves, Home/End, Tab, Shift+Tab, Up/Down, Space-exit | core `navigation` | MathLive move/leap/moveToNextGroup/moveUp/Down, visualTexTab, visualTexMoveVertical, BF/SC patches, 6 app layers |
| Empty slots (placeholders): show, enter, type, delete, restore, first slot after insert | core `slots` | 5× "placeholder selected", 5× "replace on typing", restore rules, BF/SC/EK patches, app reselection |
| Backspace/Delete at structure edges | core `editing` | kernel onDelete + VT rules, EK/BF patches, app takeover + second-delete heuristic |
| Selection (keyboard and mouse, one rule) | core `selection` | setSelection vs extendSelectionTo vs SEL patches vs app |
| Click → caret, drag, double/triple click | core `pointer` | 4 point→offset versions, SEL/BF/RS patches, app host handler |
| Caret, empty-slot marker, selection and "inside" cue geometry | engine snapshot + core `geometry` | engine region guesses + kernel DOM measuring (getAtomBounds, getCaretPoint, slot stretch) + app CSS marker |
| Text vs math mode at a text run's edge | core `mode` (one rule) | 4 kernel rules + paste/Space using another |
| Key dispatch | core `input` (one listener) | 7 listener stages, capture/bubble races |
| Focus + caret restore after an app rewrite | core `restore` (one function) | 8 focus sequences, 4 restore heuristics |
| Scrolling the caret into view | core (one scroller) | 2 kernel scrollers |
| Colours of caret, marker, selection | CSS variables only | 4 theme sheets + injected overrides |

The app keeps document-level work (formula rows, toolbar, Office) and talks to the core through commands and events.

## 3. Concepts

**Slot** — an editable list inside a structure (numerator, superscript, radicand, matrix cell, macro argument) or the
formula's top level. Items in a slot are atoms.

**Stop** — a caret position: `(slot, i)`, between item `i-1` and item `i`, `0 ≤ i ≤ length`. During migration a stop
is stored as the kernel offset it maps to (offset of the slot's `first` atom for `i = 0`, otherwise the offset of item
`i-1`); every stop maps to exactly one offset. A stop is **allowed** unless its slot lies inside an atomic structure's
internals (see `atomic`). Arrows, Tab, clicks, insertions and restores only ever produce allowed stops.

**Structure spec** — one entry per structure kind, the only place its behaviour is described:
- `slots`: slot names in navigation order (used identically by arrows, Tab, Shift+Tab, Backspace order and hit-testing);
- `vertical`: which slot is above/below which, for Up/Down;
- `atomic`: the structure's own internals have no stops; only its declared argument slots do (macros such as
  `\ket`, `\bra`, `\braket`);
- `empty`: what an empty slot is — `marker` (keeps the slot, shows a marker) or `drop` (an empty optional script or
  index disappears when the caret leaves it);
- `backspaceAtStart` / `deleteAtEnd` per slot: `exit`, `dissolve` (replace the structure by its other content),
  `remove-if-empty`, `select-structure`;
- `exit`: where the caret goes after the last slot (after the structure, or the next empty slot of the enclosing
  structure for Tab).

**Empty slot** — a slot with no items, or whose only item is a `\placeholder{}` atom. While the kernel model is still
used, the placeholder atom is the persistence form of an empty slot; the core treats "caret in an empty slot" as one
state, whether the kernel holds it as a collapsed caret or as the placeholder atom selected. Saved LaTeX keeps
`\placeholder{}` exactly as today.

**Caret state** — `{stop, anchor, mode}`; mode is math, text or latex (command being typed).

## 4. Structure table (initial)

Order is reading order, the same for every key. Decisions marked ★ change today's behaviour (§9).

| Kind | Slots in order | Vertical | Empty slot | Notes |
|---|---|---|---|---|
| fraction, binom, genfrac, cfrac | numerator, denominator (or the user's fraction-order setting) | num above den | marker | setting applies to arrows, Tab and Backspace alike ★ |
| scripts on a nucleus / subsup | subscript, superscript | sup above sub | drop | ArrowDown in a lone superscript returns to the baseline; it never creates a subscript ★ |
| big operator with limits (∑ ∫ lim) | lower, upper | upper above lower | marker for inserted templates, else drop | |
| root `\sqrt[n]{x}` | index, radicand | — | index: drop; radicand: marker | one order for arrows and Tab ★ |
| `\left…\right` | body | — | marker | |
| accents, over/under (`\hat`, `\overline`, `\overset`, `\underbrace`) | body, then above/below | above/below body | marker | placeholder inside an accent reachable by keys and clicks alike ★ |
| array, matrix, cases, aligned | cells row by row | rows | marker | last-cell Tab adds a row and keeps the column for align/split ★ |
| macro (`\ket`, `\bra`, `\braket`, …) | its arguments | — | marker | atomic: one Tab from an argument leaves the macro; no stops inside its expansion ★ |
| text run (`\text{…}`) | — (items are characters) | — | — | caret inside the run is text mode; at either edge math (§6.6) |
| top level | — | — | — | |

New kinds are added to this table, not to handlers.

## 5. Geometry and drawing

The engine already places every glyph in TeX's layout. It additionally returns a **layout snapshot** for the editor:

- `atomBox(atom)` — the atom's own box (its glyphs and the italic correction TeX added after it), never a wrapper
  shared with other atoms (fixes text, colour and font runs whose first character was given the whole run);
- `slotBox(slot)` — union of the slot's item boxes, or the empty-slot marker box;
- `caret(stop)` — `{x, top, bottom}` in formula units (sp, y down from the formula top):
  - `x`: start of slot → left edge of the first item's box; after item → right edge of its box (includes italic
    correction, so the bar never cuts a slanted letter); after a binary operator or relation → left edge of the next
    item (past the operator's space); empty slot → left edge of the marker; neighbouring stops that would share one
    place at a structure boundary (the end of a radicand and after the radical) are set 0.1em apart, the outer one
    outside the structure, so every key press shows the caret moving;
  - `top/bottom`: one bar everywhere: 0.76em tall, foot 0.05em below the baseline, scaled with the size TeX sets the
    atom beside the caret in (1 text, 0.7 script, 0.5 script's script; the typesetter records it), never stretched to
    the slot's content; the empty-slot marker spans the same band, so the caret keeps its size and place when the
    first character is typed (user, 2026-10-21: the slot-height caret was too tall in scripts and nested structures);
  - the baseline used is the item's own line (a script's caret on the script, the caret after `\int_0^1` on the
    main baseline);
- `marker(slot)` — empty-slot marker rectangle (0.40em of the slot's size wide, the caret bar's band high, centred on
  the slot's x-box), drawn by the engine itself; the caret in that slot touches its left edge and spans its height;
- `selectionRects(range)`, `insideCue(stop)` — rectangles for the selection highlight; the delimiters/sign of the
  structure holding the caret are tinted with the caret colour (as the old build did);
- `hitTest(point)` — point (formula units) → allowed stop. Beside the formula (left of its first caret, right of its
  last): its first/last stop, as beside a line of text. On an empty-slot marker (or within 0.15em of it): that slot's
  stop. Otherwise the innermost atom whose box holds the point (an atom's box ends where its scripts begin, also when
  TeX boxed them with it, as for an accented letter): on a glyph, the stop on the side of the glyph the point is on;
  on a structure's own ink (fraction bar, radical sign, delimiter) or in a gap inside it, the structure's nearest stop
  — its inner stops and the stops before and after it — by the distance to the caret's centre with vertical
  distances counted half; elsewhere the nearest stop of the formula. A structure without inner stops (an accent, a
  macro's expansion) passes the point to the structure holding it.

The painter draws the caret, markers, selection and tint from these numbers. The kernel and the app never measure the
page: anything that needs a position (popovers, IME candidate window, scrolling, the app's row layout) asks the core,
which converts formula units to page pixels with the field's origin, zoom and scroll.

CSS keeps only colours (`--formula-caret`, `--formula-placeholder`, `--formula-placeholder-selected`, selection) and
the blink animation.

## 6. Behaviour rules

### 6.1 Left/Right
Move to the next/previous allowed stop in document order (slot order from §4, depth-first). A range selection
collapses to its edge first. Word moves (Ctrl+Left/Right) skip runs of letters/digits within a slot.

### 6.2 Home/End
Start/end of the current slot; pressed again, of the top level.

### 6.3 Tab / Shift+Tab
Within the structure holding the caret: next empty slot ahead in the current slot's subtree, else the next slot of the
structure, else out of the structure (after it). Atomic structures are left in one press. In the last cell of an
array, Tab adds a row. While a command is being typed, Tab accepts the suggestion (handled by the completion, which
reports back to the core). Tab never moves focus out of the field; Esc-then-Tab leaves the formula as today.

### 6.4 Up/Down
Inside a structure with a vertical relation (§4): move to the slot above/below, keeping x as close as possible
(geometry from the snapshot). Lone superscript/subscript: back to the baseline. Otherwise the request goes to the app
(previous/next formula row). Never creates branches.

### 6.5 Space
Input-conversion hook first (upright/escapes). Then, inside a structure whose spec says so, exit after it; at the top
level insert a visible space `\ `. One place, one judgment.

### 6.6 Text/math mode at text runs
Text mode iff the items on both sides of the caret are text (between two characters of one run); at either edge of a
run the caret is in math. The mode after a caret move, `setValue`, paste, IME composition and mode display all use
this rule (core `text.ts`). Space beside a run (at its start or end) adds one word space to that run. An explicit mode
switch (typing `\text`, the menu) holds until the caret moves.

### 6.7 Empty slots
- Entering an empty slot (any way: keys, click, insertion, restore) gives the single "caret in empty slot" state.
- Typing, IME commit, paste or a command in that state fills the slot; the slot's style (bold, colour) carries over.
- Backspace in an empty slot follows the slot's `backspaceAtStart`; deleting the last item of a `marker` slot leaves
  the empty slot; of a `drop` slot removes it when the caret leaves.
- After inserting a template, the caret goes to its first empty slot in slot order (one function for typed commands,
  toolbar tiles, suggestion acceptance and input conversion).

### 6.8 Backspace/Delete
At a slot's start/end, per the structure spec (enter the script first, dissolve an emptied fraction, remove an empty
script, select a filled big operator before deleting it). Inside a slot, delete one item. Multi-row deletion and row
merging stay app commands invoked by the core when the caret is at the top level's edge.

### 6.9 Selection and pointer
A selection is a run of siblings in one slot: between an anchor stop and a focus stop it is the run of the innermost
slot holding both, each end widened to the item that holds it; two slots of one structure (numerator and denominator)
select the structure; matrix cells select a run of cells; a macro's expansion selects the macro; a base takes its
scripts along, while the scripts alone can be selected (to format an exponent). The selection remembers the stop it started from, so a selection widened to a whole structure shrinks back
into the slot it started in. Shift+Left/Right move the focus to the next stop that changes the selection (no dead
presses); Shift+Up/Down, Shift+Home/End, Ctrl+Shift+Left/Right, Alt+Shift+Left/Right, Shift+click and mouse drags
extend from the anchor with the same rule. Click = `hitTest`; double-click selects the item under the pointer (a run
of text letters, a number, a math letter, or a structure when its own ink is hit); triple-click selects the slot.
One pointer handler (the kernel's, deciding nothing itself); the app only forwards clicks in a row's space beside the
field.

### 6.10 Restore after an app rewrite
When the formula is replaced (an input conversion, its Backspace revert, contextual upright normalisation), the
caret is carried through an alignment of the old and new atoms (core `restore.ts`: longest common subsequence of
atom descriptions — type, value, command, branch, depth): if the atom before the caret survived, the caret follows
it; if it was rewritten, the caret goes after its replacement; the result snaps to an allowed stop. Not by distance
from the end, which put the caret in the wrong place as soon as something after it changed too (typing before an
unconverted `cos` converted it and the caret jumped back two atoms; old build too). A reparse of the same LaTeX
with new macro definitions (custom symbols) keeps anchoring by the LaTeX text before each end, which is exact there.
Moving to another row (Up/Down past the formula's edge) focuses it through the editor's one guarded focus routine
(`focusLine`): its delayed repairs never replay over input that arrived meanwhile (the old unguarded re-applications
at 0 ms, next frame and 80 ms reversed fast typing — ArrowUp, x, y gave "yx" — in the old build too).

### 6.11 Enter and Escape
Enter: Shift+Enter (no other modifier, not composing, no LaTeX typed in the box) breaks the line inside the formula
(it becomes aligned or gathered); in the LaTeX box Enter completes the command; inside `cases` it adds a case row;
elsewhere, with or without modifiers (after the editor's row hotkeys), it splits the formula row at the caret — a new
row after it when the caret is at its end or inside a structure. Escape with a selection collapses it to its end
(the kernel would turn it into a raw LaTeX group whose text the saved value drops); otherwise it opens the LaTeX box,
or completes it when open. Row-level states take Escape first (a multi-row selection, a pending Greek-letter prefix,
the context menu). The core (`keys.ts`) decides; the editor performs row commands and the kernel the LaTeX box.

## 7. Input and app boundary

One keydown/beforeinput/composition listener on the field, owned by the core. A keymap maps keys to core commands.
The app registers its own commands (rows, hotkeys, Office) with the core instead of listening to keys; the core calls
them for keys it does not consume (Up/Down at the top level, Enter outside arrays, Ctrl+Z/Y → app history).
The app may: send commands (`insert`, `select`, `focus`, `setValue(latex, {restore})`), read state and geometry, set
colours. The app may not: intercept movement or editing keys, set positions by offset arithmetic, measure the field's
DOM, style geometry.

## 8. Removed at the end

- Build-time patches touching these concerns (BF, SEL, SC, EK, RS pointer/caret/placeholder/navigation patches; AR is
  already dead) — eventually the whole patch chain (done in Phase 6 step 1: no build-time patch remains).
- Kernel-direct VisualTeX code for these concerns and the MathLive functions they replace (move, leap,
  moveToNextGroup, moveUp/Down, visualTexTab, visualTexMoveVertical, offsetFromPoint, getAtomBounds, getCaretPoint,
  visualTexRenderCaretSlot, getSelectionBounds, placeholder replacement paths).
- App: window/field key handlers for movement and editing, pointer gap correction, placeholder reselection and CSS
  marker, the dead raw-command anchor subsystem, unimported files.

## 9. Decisions taken (defaults; say if you want otherwise)

1. `\ket`-type macros are one structure: one Tab from an argument leaves the macro; arrows, Tab and clicks never stop
   inside its expansion (today typing there is silently lost on save).
2. Caret height: the standard bar scaled with the script size at every position, the empty-slot marker on the same
   band (user, 2026-10-21; replaces "slot-height bar inside structures").
3. `\sqrt[n]{x}`: index then radicand for arrows and Tab (today Tab goes radicand → index and skips the radicand from
   the index).
4. ArrowDown in a lone superscript goes back to the baseline instead of creating an empty subscript (`_` still
   creates one).
5. The fraction-order setting applies to Tab as well.
6. A placeholder inside an accent is reachable by keys exactly when it is reachable by clicks.
7. Last-cell Tab in align/split keeps the column of the new row.
8. Space inside a structure leaves it (a macro as a whole) without inserting anything; at the top level it inserts a
   visible space. (The old build inserted a stray `\ ` inside the structure and also moved out.)
9. The start of a macro argument (before ψ in `\ket{ψ}`) is a stop; the engine draws it apart from "before the macro".
10. Backspace right after a script always enters it first (caret after its last character); the next Backspace
    deletes. The old build sometimes deleted in one press, depending on whether the script was written with braces.
11. `\overset{A}{B}`, `\underset`, `\stackrel`, `\overunderset`: slots in argument order (A first) for arrows, Tab,
    Backspace and the first slot after insertion (before, insertion selected A but Tab then skipped B).

## 10. Change process

Every phase follows the same steps, and nothing is removed before its replacement passes:

1. **Tests first** — extend the harness (Phase 0) with the phase's behaviours: expected results are today's behaviour
   except the §9 decisions, which get explicit tests.
2. **Build in the core** — new code only in `apps/shared/visualtex-core` (and the engine snapshot in
   `apps/shared/tex-layout`).
3. **Switch callers** — kernel, plugins and app call the core.
4. **Delete** the old paths and patches for that concern in the same phase; the kernel build's exact-anchor checks
   and the patch tracer (`audit/patch-trace`) confirm the remaining patches still apply.
5. **Verify** — harness old-vs-new on the dev servers (old build 78bf2884, new build), invariants (§11), the existing
   regression suites against the 78bf2884 baseline, `build-windows`, `build-windows-office`.
6. **Windows** — port to the Windows workspace with hash checks (stop and merge if anything changed there), build
   the NSIS installer (`npm run tauri:build`), and with your OK install it and check the real main window and Office
   editor with screenshots.
7. **Checkpoint** — patch snapshot in `~/claude/visualtex-render-profiles/checkpoints/`; I ask whether to commit.

### Phases

| Phase | Scope | Done when |
|---|---|---|
| 0 Harness | Keystroke/click harness and invariants; baseline from the old and current builds | harness runs on both builds; current violations listed |
| 1 Geometry | Core skeleton (structures, stops, geometry API); engine snapshot (own atom boxes, caret rects, markers, hit regions); engine draws caret/marker/selection/tint; kernel geometry functions and slot stretch answer from the core; app marker CSS and placeholder sync removed | caret invariants pass; placeholder caret touches the marker; no backward jumps in runs; old-look caret heights |
| 2 Navigation | Arrows, words, Home/End, Tab, Up/Down, Space-exit from the core; macro atomicity; one order table | navigation tests and §9 decisions pass; old movement code, patches and app key layers gone |
| 3 Empty slots and editing | Empty-slot rules, first slot after insert, Backspace/Delete at edges | slot tests pass; 5× duplicates, related patches and app Backspace takeover gone |
| 4 Selection and pointer | One hit-test, one extension rule, double/triple click, drag | pointer tests pass; pointer patches and app pointer duplicates gone |
| 5a Text runs and dead code | One "inside a text run" rule in the core, replacing the kernel's five (mode after a selection change, setValue, paste, composition, Space beside text); app code left dead by Phases 1–4 removed (placeholder and vertical-structure helpers, the raw-command anchor subsystem, no-op state) and unimported editing files | one text-run rule; text-run tests pass; dead code gone |
| 5b Enter and Escape | One decision for Enter (array row, cases row, formula-row split, aligned row break) and Escape (collapse a selection, multi-row selection, Greek prefix, LaTeX box); the app registers its row commands with the core instead of listening for the keys | Enter/Escape tests pass; app and kernel duplicates gone |
| 5c Restore | `setValue(latex, {restore})` maps the caret by structure path (§6.10) for conversions, symbol re-rendering, reparse and Office reload; one focus routine replaces the timer and retry sequences | restore tests pass; distance-from-end and length heuristics gone |
| 5d Scrolling and anchors | One scroller (caret visible in the row horizontally and in the editor pane vertically; `preventScroll` honoured); the suggestion list positioned once, from the core; the IME candidate window anchored at the caret | caret stays visible in long rows; IME window beside the caret; duplicates gone |
| 6 Own core | Model, parser, serialiser, mode editors, IME and undo move into the core's source; patch chain deleted; macOS adopts the shared core | kernel fragments and patches removed |
| 6a Patch chain | The build-time patches are written into the kernel source; generated tables fill named data slots | built app byte-identical (done) |
| 6b Kernel as core source | The kernel becomes TypeScript in `apps/shared/visualtex-core/src/kernel/`: MathLive 0.109.2's TypeScript (its own build of it reproduces the published bundle byte for byte) with every VisualTeX change of the fragments carried over; the generated tables and the input judgment come from a host module each app provides (`virtual:visualtex-kernel-host`) instead of text slots; the fragments and the loader go | the TypeScript compiled like the published bundle equals the current kernel module for module (data sites aside); it type-checks; harness, suites and targeted scenarios unchanged (done) |
| 6c Unused parts | Parts VisualTeX never runs are deleted from the core source (as found by the usage inventory: virtual keyboard, MathLive's menus and toggles, scrim, environment popover, keystroke caption, speech output, sounds and haptics, static rendering and `<math-span>`/`<math-div>`, the Typst, MathJSON and extra SSR conversions, the Compute Engine hooks, the UI string tables, deprecated-option shims, the unused font lookup and render hooks, and every exported helper nothing references) | type check passes (also with unused-locals checking); module evaluation order unchanged; tests unchanged (done) |
| 6d One type check | The kernel source compiles under the app's strict settings; the app's types come from the core; the copy service validates with the kernel's parser instead of the npm package's (`mathlive/ssr`), the Node tests use the core, and the npm `mathlive` package is dropped | `tsc` of the app covers the core; no `mathlive` dependency (done: 525 strict errors fixed with the compiled kernel byte-identical; Node scripts that reach the kernel run with `tsx --import ./scripts/kernel-node.mts`) |
| 6e One input listener | Key, beforeinput and composition dispatch and the keymap move into the core (§7); the app registers its commands | key, IME and shortcut tests unchanged (in progress: the field has one key entry, owned by the kernel, offering each keydown to the app's key commands kept by the editing core's host; & is decided by the core like Enter and Escape; next: the IME/Backslash guards and the auto-exit after composed input, owned twice today, move into the kernel's composition handling) |
| 6f macOS | macOS uses the shared core instead of its own kernel copy | macOS suites pass on the Mac |

## 11. Invariants checked on every formula of the corpus

1. Right-walk and left-walk visit the same allowed stops, in exactly opposite order; every stop is visited once.
2. Within one slot, caret x never decreases on Right and never increases on Left.
3. Each press moves the caret visibly (no two consecutive stops with the same caret rectangle), except where the
   spec allows it.
4. Tab visits slots in the structure table's order.
5. Typing a character at every allowed stop changes the saved LaTeX (no input lost).
6. Clicking the centre of each stop's caret bar returns that stop; where two carets coincide (the end of a radicand and
   the end of the radical), a click just beside the bar, on the item next to it, does.
7. The caret bar never overlaps a glyph's ink box of the item before it.
8. The empty-slot caret touches the marker's left edge and spans its height.
