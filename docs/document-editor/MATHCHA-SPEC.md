# Document mode = Mathcha (measured spec)

The user's requirement (2026-10-30): the document mode follows Mathcha strictly — interface, every trigger,
typesetting. Second round, same day: follow Mathcha's *behaviour*, but not its look or its names — the `\` menu
is VisualTeX's own (its look and LaTeX command names), and the writing surface fills the window instead of
Mathcha's 700 px page. These and the other changes from that round are listed under "Deliberate differences". Everything below was measured in the real editor (mathcha.io/editor, headless Chromium, 2026-10-30;
screens in ~/claude/visualtex-document/mathcha on the VPS) or read from its code (bundle-cc720325.js).

## Interface

| Part | Mathcha |
|---|---|
| Canvas | `#eeeeee`; the page starts right under the toolbar (VisualTeX differs: see below) |
| Page | white, 700 px wide, 1 px light border, padding 15 px top, 20 px left/right/bottom (content 660 px) (VisualTeX differs) |
| Toolbar (text) | one light-grey row: Align L / C / R / Justify · Section ▾ · Ordered list ▾ · Unordered list ▾ · Indent · Outdent · Font size (tiny…Huge, default normalsize) · B I U S · Text colour · Highlight · Undo · Redo (tooltips name the shortcuts) |
| Toolbar (math) | alignment disabled · Σ→Σ (display style of the inline formula) · colour · highlight · Export: Latex · Undo · Redo |
| Left items bar | a white floating vertical bar on the canvas, left of the page: Suggestion Box, Table (+ drawing, image, checkbox: not in VisualTeX). VisualTeX: none — the table picker is in the toolbar (user, 2026-10-30) |
| Status | a box at the canvas's bottom right: zoom `100%` ▾ · `Words: n, Math: n` (VisualTeX: a status line under the canvas) |
| Context menu | right click: Copy · Copy Plain Text · Copy Latex · Cut · Delete; in a table also Insert Row Above / Below, Insert Column on Left / Right, Delete Row / Column / Table (and Merge Cells with a cell selection) |

## Triggers (text mode)

| Input | Result |
|---|---|
| `$` | an inline formula at once (opened) |
| `\` | (VisualTeX differs: see below) the Suggestion Box at the caret: a search field (focused) over a list; typing filters, ↑/↓ choose, Enter inserts, Esc closes without typing anything. Text-mode list: `\inline-math ($)`, `\math-container`, `\table`, `\align`, `\gather`, `\multiline`, `\horizontal-line` (Mathcha also: diagrams, images, checkbox, page break, TOC, groups, Z schema) |
| `/` | the same box (Mathcha's alternative trigger), at the start of a line or after a space |
| Ctrl+Alt+1 | toggle a section (heading) on the line; Tab on a section demotes it (1 → 1.1 → 1.1.1), Shift+Tab promotes |
| Ctrl+Shift+7 / 8 | ordered / unordered list |
| Ctrl+Shift+L / E / R / J | align left / centre / right / justify |
| Ctrl+B / I / U | bold / italic / underline |
| Tab (plain text) | a tab character; in a table: next cell; in a list: indent |
| Enter | new paragraph with the same style (a section continues; Ctrl+Alt+1 turns it off); Enter on an empty list item leaves the list |

## In a formula

| Key | Mathcha |
|---|---|
| Tab | next empty slot / next slot of the structure (numerator → denominator) / out of the structure; at the top level: leave the formula, caret after it |
| Shift+Tab | the same backwards; at the top level it stays |
| → at the end / ← at the start | leave after / before |
| → / ← from the text next to a formula | enter it at its start / end |
| ↑ / ↓ in an inline formula | stay (VisualTeX differs: leave to the line above / below) |
| ↑ / ↓ in a display formula | leave to the line above / below |
| Enter in a display formula | a new line inside it (each line centred) |
| Esc, Space, `$` | stay (Space and `$` are typed) |
| `\` | the math Suggestion Box (VisualTeX: its own LaTeX command completion) |
| Backspace right after a formula | deletes the whole formula |
| Click | enters at the clicked place |

## Typesetting

| | Mathcha default |
|---|---|
| Text | Arial (Helvetica, sans-serif) 15 px, line height 1.4 (21 px lines), paragraph spacing 0 / 0 |
| Math | 17 px, line height 1.2 (font Asana; VisualTeX keeps its TeX-exact engine and fonts) |
| Formula being edited | inline: background rgba(76,175,80,0.05) + 1 px outline rgb(193,212,193); display: the same background, no outline, full width; caret 2 px green rgb(0,128,0) |
| Formula not being edited | no background, no outline |
| Display formula | its own line, centred, line padding 8 px top/bottom, margin 0.3em; the edited line shows a grey `(...)` at the right: click to number it `(n)` |
| Section | level 1: 1.2em bold, padding 0.5em top/bottom, number prefix `1.` + 0.5em; level 2: 1.1em, `1.1.` + 0.25em; level 3: 1em, `1.1.1.` |
| Lists | prefix `•` / `1.` (min width 1em) indented ~30 px, no extra spacing |
| Table | full width (660 px), cells 1 px solid black, padding 7.5 px, empty row 37 px; no header styling |

## Deliberate differences

- Formulas are VisualTeX's kernel and TeX-exact engine (Computer Modern / Times profile), not Asana.
- **Writing surface** (user, 2026-10-30): no page and no grey canvas — white from side to side, at least as tall
  as the window, with 45vh of room below the last line; a click below the text goes on writing there (on a new
  line after a table, list or display formula), a click in a margin puts the caret on that line. Zoom and the
  word count sit in a status line under the canvas, where they cannot cover text.
- **`\` and `/` type themselves** (user, 2026-10-30): the command is ordinary text and a menu under it offers
  LaTeX commands in VisualTeX's look (the formula editor's suggestion list: numbered rows, the command, a
  picture). Complete commands run as typed: `\section` + Space, `\section{Title}`, `\textbf{…}`,
  `\begin{itemize}`, `\item`, `\begin{equation}`, `\align*` + Space, `\(`, `\[`, `\\`; `\$ \% \& \# \_ \{ \}`
  are those characters; formula commands (`\alpha`, `\frac`, …) go into an inline formula. Enter/Tab/1-9 pick,
  Esc keeps the text. Because the command is read from the text, a `\` typed through an input method (a "Process"
  key, then text) works like any other — Mathcha's key-based trigger missed those. DESIGN.md §6.1.
- **↑/↓ in an inline formula** leave it for the line above/below at the caret's x when there is nowhere to go
  inside (user: "the arrow keys sometimes can't take me out"); in a fraction they still move between its parts.
- A selected (not open) formula is stepped over by ←/→ instead of being entered; an inline formula left empty
  is removed (Backspace in it already did that).
- Third round (user, 2026-10-30):
  - Backspace at the start of a styled line removes the style first, as in Word: a section becomes body text (the
    new line after a section keeps the section style, as in Mathcha), a list item leaves its list, a quote ends,
    an indent steps back; then the line joins the one before.
  - No items bar: the table button (with the 8 x 8 size picker) is in the toolbar after Indent / Outdent; the
    Suggestion Box button is gone. Toolbar menus open over the page (nothing clips them); below 720 px the toolbar
    tightens its spacing and stays one row down to the app's 640 px minimum.
  - Formula colours: an edited or selected formula, inline or display, has a light blue background
    (rgba(33,150,243,0.08)); the caret is VisualTeX's. No green.
  - Selections look like the formula rows': selected text has their selection colour (the accent at 22%), and a
    formula inside a selection is highlighted whole in that colour, not glyph box by glyph box.
  - A drag that begins in a formula (closed or open) and leaves it becomes a text selection from the formula,
    which is selected whole, across the line and other lines (scrolling past the edge).
  - The `\` menu numbers its rows from the first `\`; 1-9 pick.
- Fourth round (user, 2026-10-31): Space picks the highlighted command (as in the formula editor's suggestions);
  the document takes the formula mode's layout (its bottom bar, formula tools and tiles, LaTeX tab, status bar;
  DESIGN.md §6.4); code blocks in VS Code's Dark+ (§6.5); dragging a selection, in the text and inside formulas, as
  AxMath does (§6.6). Markdown input/output is planned after the LaTeX side.
- Not in scope: diagrams, images, checkbox, page break, TOC, theorem boxes, bibliography, printing.
