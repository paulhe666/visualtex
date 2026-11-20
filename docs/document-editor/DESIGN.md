# VisualTeX document editor (block editing in the main app)

Status: in development (started 2026-10-29). Since 2026-10-30 the interface, every trigger and the
typesetting follow Mathcha as measured in **MATHCHA-SPEC.md** (the user's requirement), with the deliberate
differences listed there (VisualTeX's own `\` menu and names, a full-window writing surface); where this file
disagrees on keys, menus or layout, MATHCHA-SPEC.md wins. The model, the formula islands and the export below stand. Scope of the first version: body text, inline formulas, display
formulas, headings, lists, quotes and tables. Images, drawings/TikZ and the Office editing window are out of scope.
Background research: `~/claude/block-editor-research/reports/Notion Mathcha block editor study.md` (VPS).

## 1. Decisions

1. **Document-first, not per-block editors.** One ProseMirror editor holds the whole document. Prose lives in
   one contenteditable; formulas and tables are nodes inside it. Block affordances (slash menu, turn into,
   handle, block selection) are plugins on top. Notion abandoned per-block contenteditable; WordPress's block
   editor scored SUS 46. Custom-rendered prose (EditContext) is impossible on macOS WKWebView.
2. **The document stores LaTeX, never kernel state.** A formula node is `{ id, latex }` (+ `numbered` for
   display). The kernel and the engine are views of that string.
3. **Same engine in both states.** A formula that is not being edited is the engine's static markup (the
   MathPreview path: `convertVisualTexLatexToMarkup`, cached by font revision). The one being edited is the
   row editor's own `FormulaField` (exported from `MathEditor.tsx`), so typing behaves exactly as in rows:
   autocomplete, auto-exit, input judgment, fonts, IME guards, raw-LaTeX box.
4. **One live formula at a time.** Opening a formula mounts a `FormulaField` into that node's DOM; leaving
   unmounts it and shows the static markup again. A document with hundreds of formulas has one kernel field.
5. **One undo history.** Every `FormulaField` edit (`onEdit`: before/after LaTeX + selection snapshots) becomes a
   ProseMirror transaction on the node's `latex`. Ctrl+Z/Ctrl+Y inside a formula run the document's history;
   the open field takes the restored LaTeX and the matching selection snapshot. The kernel's own undo is unused
   (as in the row editor, which also calls `resetUndo`).
6. **Edge rules live in the editing core.** `EditingHost.horizontalExit(direction)` (new) is asked when
   Left/Right have no stop left inside the formula, like `verticalExit` for Up/Down. The document's host
   answers by closing the formula and placing the text caret beside it.
7. **Esc keeps its app meaning** (open the raw LaTeX box). A formula is left with the arrows, Enter or a click.

## 2. Document model (ProseMirror schema, `apps/shared/visualtex-document/src/schema.ts`)

Blocks (every block has a stable `id`, assigned by the ids plugin):
- `paragraph` (inline*), `heading` (level 1-4, inline*), `blockquote` (block+),
- `bullet_list` / `ordered_list` (`start`) of `list_item` (paragraph block*),
- `math_display` (atom; `latex`, `numbered`), `horizontal_rule`,
- `table` / `table_row` / `table_cell` / `table_header` (prosemirror-tables; cells hold `paragraph | math_display`
  blocks, with `colspan`, `rowspan`, `colwidth`, `align`).

Inline: text with marks `strong`, `em`, `underline`, `strike`, `code`; `math_inline` (atom; `latex`);
`hard_break`; `eqref` (atom; `target` = a display formula id).

File: a Markdown file and the document's memory (§13). `.vtxdoc` — JSON `{ format: "visualtex-document", version,
doc: <ProseMirror JSON>, meta, images }` — was the file until 2026-11-18 and stays readable (`fileFormat.ts`).

## 3. Formula islands (`formulas.ts` + app `DocumentEditor.tsx`, `DocumentFormulaField.tsx`)

States: **static** (engine markup, caret beside it in the text), **selected** (ProseMirror NodeSelection, the
whole formula highlighted), **open** (live FormulaField inside the node).

| Action | Result |
|---|---|
| Right with a formula after the caret / Left with one before | open it, caret at its start / end |
| Click a formula | open it, caret at the clicked stop (`placeCaretAtPoint`) |
| Ctrl+M (no selection) | insert an empty inline formula and open it |
| Ctrl+M with text selected | the selected text becomes the formula's LaTeX (`\text{}` for CJK) and opens |
| Ctrl+Shift+M | insert / convert to a display formula |
| `$$` typed in text | an empty inline formula in their place, open (2026-11-09; a single `$` stays text) |
| `$$$$` typed in text | an empty display formula in their place, open (the third and fourth `$` reach the new inline formula, which gives way to it) |
| Backspace after a formula / Delete before it | first press selects it, second deletes it |
| Backspace in an empty line after a display formula / Delete in one before it | the line goes, the formula opens with the caret at its end / start (the formula is never deleted with the empty line, 2026-11-01) |
| Left/Right at a formula's edge (horizontalExit) | close, caret before / after the formula |
| Up/Down with nothing above/below inside (verticalExit) | close, caret to the line above / below at the same x |
| Enter in an inline formula | close, caret after it |
| Enter in a display formula | close, new paragraph after it |
| Tab (→) at the end of a display formula | the block after it: a text line's start, the next display formula (opened at its start), a list's or table's first line; a new empty line only when nothing follows (2026-11-01: before, any block that was not a text line got an empty line inserted before it) |
| Click outside | close (the formula is kept, never dropped) |
| Backspace in an empty formula | remove it, caret where it was |
| Shift+arrows in text | the formula is one unit; it is never opened |
| Mouse selection | §6.2: the formula rows' rules — a drag in a formula selects in it from the first press, an end of a selection through the text inside a formula takes the formula's part |
| Bold (Ctrl+B, toolbar) | §6.2: the text, the formulas in the selection and the formulas' parts — every math symbol in its bold font (`\bm`) |

Paste: text containing `$...$`, `\(...\)`, `\[...\]`, `$$...$$` or `\begin{equation|align|gather...}` becomes
text with formulas (Pandoc dollar rule); plain text stays text.

## 4. Numbering

Display formulas are numbered by default (`numbered: true`), counted in document order by a decoration plugin;
the number is drawn by the document (round brackets, right aligned), never by the kernel. `\label{eq:<id>}`
comes from the block id on export.

## 5. LaTeX export (`latex.ts`)

`article` (ctexart when the text has CJK), only the packages used. Paragraph → text + blank line; headings →
`\section`/`\subsection`/`\subsubsection`/`\paragraph`; marks → `\textbf`, `\emph`, `\underline`, `\sout`
(ulem), `\texttt`; lists → itemize/enumerate (`start` via `\setcounter`); quote → `quote`; inline math →
`\( ... \)`; display → `equation` + `\label{eq:id}` / `equation*`; multi-row formulas → `align`/`align*`;
tables → `tabular` (+ booktabs rules, `\multicolumn`/`\multirow` for spans, `p{}` columns when a cell holds
display math). Text is escaped (`# $ % & ~ _ ^ \ { }`).

## 6. Keys and menus (as built)

| Key / action | Does |
|---|---|
| `/` at a block start or after a space | block menu: Text, Heading 1-3, Inline / Display formula, Equation reference, lists, Quote, Table, Divider (filter by typing; English, Chinese and pinyin initials) |
| `#`..`####`, `-`, `1.`, `>` + space | heading, bulleted / numbered list, quote |
| `**b**`, `*i*`, `` `c` ``, `~~s~~` | bold, italic, code, strikethrough |
| Ctrl+B / I / U / E | bold / italic / underline / code |
| Ctrl+Alt+0..4 | text, heading 1-4 |
| Alt+Shift+Up/Down, Ctrl+D | move block, duplicate block |
| Tab / Shift+Tab | next / previous table cell (Tab in the last cell adds a row); list indent |
| Ctrl+M, Ctrl+Shift+M, Ctrl+Shift+N | inline formula, display formula, toggle equation number |
| Block handle (left margin) | click: Turn into / Move / Duplicate / Delete; drag: move; `+`: new block with the `/` menu |
| Table toolbar (above a table with the caret) | rows / columns before-after, delete row / column, merge / split, header row, cell alignment, delete table |
| Toolbar `(1)` or `/ref` | reference to a numbered display formula: shows "(n)", follows renumbering, Ctrl+click jumps, exports `\eqref{eq:<id>}` |
| Formula toolbar and tiles | insert into the open formula (or a new inline formula at the caret), through MathEditor's own `prepare/attemptLatexCommandInsertion` |

Main window: the header switch Formulas | Document; the document autosaves into its file (§13), exports Markdown
and PDF; the source panel shows the
document's LaTeX read-only; Ctrl+S / Ctrl+O go to the document; the header zoom scales the text (17px x zoom).

Not yet: images, drawings, Word export of documents, LaTeX import, `\eqref` from pasted text, numbered tables/figures
with captions, the formula colour/style strip in documents, the command-suggestion popup of the row editor (the kernel's
own completion works).

### 6.1 LaTeX commands in the text (`suggest.ts`, app `CommandMenu.tsx`, `documentMathCommands.ts`)

- `\` (and `/` at a line start or after a space) is typed into the text. The plugin state notices a trigger
  character that the transaction itself inserted right before the caret (not undo, paste or drop; not in code), so
  every input path counts — a key, an input method's commit, the items bar button. While the caret stays at the end
  of `\name`, `\name{argument` or `\begin{environment`, the command is "being typed" and the menu shows matches:
  document commands first (`DOCUMENT_COMMANDS`; `\sec` also finds `\subsection`), then the formula editor's commands
  (the app's `searchCommands`, the same ranking as in formulas), then looser matches.
- No decoration marks the typed command: changing decorations next to a composition breaks Chrome's IME.
- Running an entry deletes the typed command and runs the entry's command where it was. `}` closing an argument or
  environment, `(` / `[` / `\` / an escaped character right after `\` run at once. A formula symbol becomes a
  closed inline formula and the text goes on after it; anything with slots opens an inline formula and the formula
  editor inserts the command (`formulaInsert.ts`, as the toolbar does).
- Keys: ↑/↓, Enter/Tab/Space (Space as in the formula editor's suggestions: once a name is typed, not inside an
  argument), 1-9 (the numbered page of nine, always shown; not while an argument is typed), Esc (dismiss: the text
  stays). A click on a row runs it.

### 6.2 Selections and drags around formulas

Selections with the mouse follow the formula rows (docs/editing-core/DESIGN.md §6.9; user, 2026-11-01: "a formula
is one whole, but selecting must not treat it as one"). `SelectGesture` (`formulas.ts`) owns every drag: a press in
the text (the plugin's `mousedown`: not Shift, not double/triple clicks, not in tables) puts the caret there; a press
on a closed formula opens it at the press point. While the pointer stays in the formula the drag began in, the
formula selects in itself (its structural selection) from the first press — the editor that opened with the press
never saw it, so the gesture drives it (`selectStops`); an open formula's own editor follows its drag itself until
the pointer leaves it. Beyond that formula the selection runs through the text, and each end of it inside a formula
— the one the drag began in, or the one under the pointer, inline or display — takes that formula's part at the
formula's top level (an end inside a structure takes the whole structure), as a selection across formula rows does:
the formula before the text from the end's stop to its end, the formula after it from its start to the end's stop.
Back in the formula it began in, the drag is that formula's own selection again. Only moves with the left button held
count (AxMath); past the top or bottom of the document it scrolls.

A **mixed selection** is the text selection (ProseMirror's, between the formulas) plus the formulas' parts
(`FORMULA_PARTS` meta, plugin state `parts`, cleared by any other selection or document change). Each partly selected
formula shows its part in a live field mounted for it, passive (`FormulaSession.showPart`; not focused, its presses go
to the document: a click opens it, a drag moves the selection). `mixedSelection.ts`: Backspace/Delete/Enter and typing
remove the text and the parts in one step (a formula left empty goes; begun in a formula, the caret returns into it);
Ctrl+C/Ctrl+X copy the text with each formula's part as a formula of its own; a drag of the selection
(`documentDrag.ts`) moves it the same way; Bold sets the text's mark, wraps whole formulas in `\bm{…}` and sets the
parts bold through the kernel (all of it bold already: all of it back to normal). Formulas wholly inside a text
selection are still marked whole (`is-in-selection`).

### 6.4 The document in the formula workspace (2026-10-31)

Document mode renders inside `EditorWorkspace`'s own frame instead of replacing it: the document editor takes the
formula rows' card (`.document-editor-surface`), and everything around it is formula mode's — the header tools
(quick OCR, zoom, undo/redo bound to the document, the LaTeX | Markdown input switch; with the formula tiles
collapsed, their expand button sits at the right end of the header instead of over the page). The document's own
buttons are a toolbar across the top of the card (`DocumentTextToolbar`: alignment, section, lists, indent, table,
code block, link, size, B/I/U/S, text colour and background with the formula editor's palette — the same fixed and
custom colours, shared through the same storage, `ColorPalette.tsx`); the bottom dock's strip shows the open
formula's controls (`DocumentControls`: \displaystyle, the palette, copy LaTeX), the 公式工具 tiles (they insert into
the open formula or a new inline one), the source tab (`DocumentSourceView`: the document as LaTeX or Markdown, read-only, §7.5), the right tile panel
and the status bar (words, formulas). The editor publishes its view, state and zoom in `stores/documentEditorStore.ts`. Quick OCR
inserts at the caret captured when the screenshot starts (adapter `captureInsertionTarget` / `insertLatexAt`).

### 6.5 Code blocks

`code_block` (text only, a `language` attribute) is drawn by `codeBlocks.ts` (a header with a language select and
Copy over a VS Code Dark+ box) and highlighted by the app's `codeHighlighter.ts`: Shiki with VS Code's own TextMate
grammars and its Dark+ theme, the JavaScript regex engine, 33 grammars each loaded on first use. Keys in code: Tab /
Shift+Tab indent, Enter keeps the indentation (one step more after `{ [ ( :`), ↓/↑ at the ends step out (a new line
when nothing follows), Backspace in an empty block makes it body text; `$`, `\` and pastes are plain text there.
Highlighting decorations stay as they are while an input method composes. Export: `lstlisting` (listings' language
names where it knows the language).

### 6.6 Dragging a selection (AxMath's drag)

- Document: pressing on a selection (text, formulas, a selected formula) and moving is the browser's drag, which
  ProseMirror moves (Ctrl: copies) to the drop point; a click on a formula in the selection still opens it.
- Formula: no drag of a formula's selection (2026-10-31, the user's decision after trying it in formula mode: the
  kernel's selection drag, its row and document drop targets and the one-step move between rows were removed; a
  press on a formula's selection works as it did before, in formula rows, document formulas, Office and Anywhere).
  Superseded 2026-11-01: AxMath's selection drag is back (editing-core DESIGN §6.12) and in the document a formula's
  selection can be dragged out of it (§7.7).

### 6.7 IME next to inline formulas

The static markup of a formula holds no real text (its glyphs are in absolutely placed boxes). With the caret just
after (or before) such an inline atom, Chrome loses an IME composition after its first letter: the rest is dropped
and the selection jumps onto the formula. Each inline formula's node view therefore carries a zero-width text
character at both ends (inside its non-editable span); pinyin then works in every position (tools/pw/ime-matrix.mjs).

## 7. Markdown (2026-10-31)

### 7.1 Markdown only

Documents are typed and written as Markdown (Typora and the note apps). The LaTeX input mode, LaTeX output and
import, and the TeX-typeset LaTeX documents of the personal edition were removed on 2026-11-09, in every build; the
shared package still has its LaTeX input code (`inputMode`), which the apps no longer select.
`stores/documentFormatStore.ts` keeps the Markdown settings (references, input rule groups).

### 7.2 Typing Markdown (`markdownInput.ts`, `markdownLive.ts`, `inlineMarkdown.ts`) — Typora's input (2026-11-07)

The user's request (2026-11-06): Markdown input follows Typora's input logic exactly, the formula environment apart.
Measured on Typora 1.14.15 (and Obsidian) with scripted keyboard runs: ~/claude/block-editor-research (report, `handson/*.jsonl`).

- Inline Markdown is live, as in Typora: `**` / `__` bold, `*` / `_` italic, `***` both, `~~` strike, `==` highlight,
  `` ` `` code, `[text](url "title")`, `<url>`, `https://…` / `www.…` (linked when a space follows), `<u>` `<sup>` `<sub>`
  `<kbd>` `<mark>` `<s>` `<b>` `<i>`, `\` escapes. A span is formatted once the delimiter that closes (or opens) it is
  typed. While the caret is in a span or at its edges its Markdown is shown as grey characters (a link's address as an
  address); the caret steps through them, Backspace deletes them, and the line is read again after every change
  (deleting one `*` of `**bold**` makes it italic, the other plain text; typing `**` round text makes it bold); when the
  caret leaves, the characters go and the formatting stays. The shown characters are text with the `md_syntax` mark,
  one region at a time, outside the undo history (Ctrl+Z hides them first); saving, exports and the clipboard leave them
  out (`withoutMarkdownSyntax`). An input method's text and a mouse press are read when they are done.
- Reading is CommonMark's with Typora's leniency: a closer may follow a space (`**x **`), punctuation and CJK next to a
  delimiter do not stop it (`前**“引号”**后`), `__b__c` closes, an opener never has a space after it (`** x**`, `a * b`
  stay text), `_` never opens inside a word (`snake_case`), `**a *b** c*` is bold "a *b", `*it**` stays text. A code
  span closed round formatted text keeps that Markdown as code.
- Typora's format keys (Ctrl+B/I/U, Ctrl+Shift+`, Alt+Shift+5; Mac ⌃` / ⌃⇧`): a selection in one line takes the
  formatting, or loses it when all of it has it, and the caret goes to its end inside the span; with no selection, in a
  span of that kind its markers go, in a word (the system's word cutting, Chinese too) the word takes it, elsewhere an
  empty pair `****` with the caret inside. Ctrl+K: `[selection]()` with the caret where the address goes (`[]()` with
  none); in a link, the link dialog. Ctrl+\: the formatting goes.
- Line start + Space: `-` `*` `+`, `1.` / `1)` (the list starts at that number), `>` (`>>` nests), `[ ]` / `[x]` in a
  list item (a task); a list marker at the start of a later item nests the item (in that kind of list); `[!NOTE]` as a
  quote's first line makes it an alert (on `]`). A Chinese input method's `1。`, `》` and full-width space count too.
- Enter, or the caret leaving a line just typed: `# Title` … `###### Title` (closing `#`s dropped) → section, the text
  after the caret going on as body text; ```` ```lang ```` / `~~~` → code block; `---` `***` `___` → rule. Enter only:
  `| a | b |` → table (header row and an empty row, the caret in it); `[toc]`; a line of HTML that opens a CommonMark HTML
  block (`<div …>`, `<details>`, a comment) → HTML block; `[^1]: text` gives an empty footnote its text.
- `[^1]` makes a footnote (its text is asked for at once), `![alt](picture)` the picture when it is found, `:smile:` the
  emoji (`:` and letters open Typora's emoji list; Enter takes one), `</` completes the open tag. `( [ { " '` come in
  pairs (not before a word; `'` not after one); a closing character typed over its pair steps over it, Backspace
  between them deletes both, a selection is wrapped.
- Enter in a section goes on with body text; at a section's start an empty line goes before it. Backspace at the start
  of a later list item: the item joins the one before as a paragraph of its own (again: the two lines join); right
  after a list or a quote: the line joins its last line; at a cell's start: the end of the cell before; in an empty first
  line of the document: the line goes. Enter on an empty list item outdents it / leaves the list; on an empty quote line
  leaves the quote; in a table goes to the cell below and leaves the table from its last row; a new task is unticked.
- Formulas are VisualTeX's (not Typora's), opened by doubled dollars so that a single `$` is never taken for one
  (user, 2026-11-09): `$$` opens an inline formula, `$$$$` a display formula — the third `$` reaches the new empty
  inline formula and is held there (nothing shown), the fourth turns it into a display formula; anything else typed
  in between lets the held `$` go. A single `$` stays text (`$5`, `\$`). `$` typed in a formula closes it, and the rest
  of its closing delimiter typed right after does nothing (one more `$` after an inline formula, up to three after a
  display formula), so `$$x$$` and `$$$$x$$$$` type as written; Enter in an empty display formula keeps the caret
  there; Ctrl+Enter leaves any formula; `￥` from a Chinese input method counts as `$`. Files still read and write
  `$…$` and `$$…$$` (§7.5).
- Typora's keys — Windows: Ctrl+1…6, Ctrl+0 (body text), Ctrl+= / Ctrl+- (a level up / down), Ctrl+T, Ctrl+Shift+K
  (code block), Ctrl+Shift+M (display formula), Ctrl+Shift+Q, Ctrl+Shift+[ / ], Ctrl+] / Ctrl+[ (list level),
  Ctrl+Shift+I (picture), Ctrl+L / E / D (select the line, the formatted span or cell, the word). Mac: ⌘1…6, ⌘0, ⌘= / ⌘-,
  ⌥⌘T, ⌥⌘C, ⌥⌘B, ⌥⌘Q, ⌥⌘O / ⌥⌘U, ⌥⌘X (task), ⌘] / ⌘[, ⌃⌘I, ⌘L / E / D (the ⌘⇧ keys of before stay). In Markdown input
  these are the document's (the app's zoom and keypad keys act elsewhere).
- In the "/" menu Space takes the highlighted entry once a name is typed, once the highlight was moved with ↑/↓, or
  after a lone "/" at the start of a line (mid-line "km / h" stays text).
- `\` + an ASCII punctuation character types the character, which then never acts as Markdown (Markdown's escape).
- A line-start or Enter conversion is an undo step of its own after the characters typed; the change after it starts a
  new step (another plugin's follow-up — ids, the shown Markdown — does not).
- Both inputs: Backspace / Delete next to a display formula or a rule selects it first (a second press deletes it); next
  to a code block the caret goes into it (Typora deletes such blocks silently, a reported data loss).

### 7.3 The "/" menu

In Markdown input `/` (at a line start, after a space or after Chinese text; also `／`, and `、` at a line start, which
a Chinese input method types for `/` or `\`) opens the command menu with Markdown's
blocks and marks, written as Markdown writes them (`# …`, `1. …`, `$$ … $$`, ```` ``` ````, `| … | … |`, `**…**`,
…), and references to numbered formulas as `(n)`. Entries are found by English name, pinyin and its initials
(`/biaoge`, `/bg`) or Chinese (`/表格`); the keys are §6.1's (Enter, Tab, Space, 1-9, Esc). `\` opens no menu in
Markdown input.

### 7.4 Paste

In Markdown input a plain-text paste that reads as Markdown (`looksLikeMarkdown`) is parsed into blocks, marks,
formulas and tables; in LaTeX input only an unmistakable Markdown document (headings, fences, a table rule, a list of
several lines: an AI answer, a README; never LaTeX source) — Typora parses Markdown pastes in every mode. HTML pastes
go the usual way.

Pasted text whose HTML only shows it as source (VS Code, a `<pre>`, a chat message that still shows the `#` or
`**`) is read as that source; HTML that renders it (an `<h1>` without the `#`) pastes as formatted HTML.

LaTeX source pasted in either input, or a `.tex` file opened, is read into the document's structure
(`latexImport.ts` turns it into Markdown the document reads): \section… into headings, \textbf / \emph /
\underline / \texttt / \href / \textcolor into marks, itemize / enumerate into lists, quote, tabular into a table,
lstlisting / verbatim into code blocks, math as it is with equation / align / gather as numbered display formulas
and \ref / \eqref as references; preamble commands and comments are dropped.

### 7.5 Output (`markdown.ts`)

The source tab, Copy and Export follow the output switch in the source tab's header (LaTeX or Markdown). Markdown
is CommonMark with GFM tables and `$…$` / `$$…$$` math (as Typora, Obsidian and Pandoc read it), `==highlight==`,
`<u>underline</u>`, `~~strike~~`, colour, size and other highlight colours as `<span style>` (Typora shows and keeps
them); bold or italic that CommonMark could not read back next to Chinese punctuation is written `<strong>` / `<em>`
(CommonMark #650); literal `==` and tags are escaped. Numbered formulas are `$$ … \tag{n} $$`; a formula the text
refers to also carries `\label{…}` and the references are `$\eqref{…}$`, links in Typora and MathJax (the source
tab can switch them to plain `(n)`, which KaTeX viewers such as VS Code need). The first table row is the header and
a cell's lines are joined with `<br>`. A `.md` file opens with the same rules (plus `\(…\)`, `\[…\]`, `equation` /
`align` environments; `\label`s become the formulas' ids and `$\eqref{…}$` references, YAML front matter a yaml code
block).

### 7.6 Keeping the file

A `.md` file is kept with the document (`readMarkdown`, the `markdownSourceKey` plugin state): writing Markdown
gives back every top-level block that is still as it was read character for character, with the space between such
blocks, the front matter, link definitions and the line ends (CRLF stays CRLF) — an unchanged file comes back byte
for byte, as Typora keeps files. Changed and new blocks are written anew in the file's own habits (its bullet, its
ordered-list delimiter, underlined headings); a block edited in place keeps the spacing around it; a numbered formula
whose number changed is rewritten with its new `\tag`.

### 7.7 Dragging (DESIGN §6.6, revised)

Selections in the document are dragged with the pointer, no longer with the browser's drag and drop (the Windows
app keeps Tauri's drag-drop handler for its OCR model files, and Tauri documents that it blocks HTML5 drag and drop
on Windows): `documentDrag.ts` (press on the selection, 3 px, drop caret or a line between blocks, autoscroll, Ctrl
copies, Esc / right button cancel, a click places the caret or opens a formula). Inside a formula nothing is
dragged (§6.6).

2026-11-01, display formulas (the user: "I can't select and drag display formulas"):
- A selection whose ends sit at line edges (the end of the line above a display formula, the start of the line
  below) drags the blocks between them, not the line breaks: the lines around it are no longer joined and no empty
  paragraph is left (`documentDrag.ts` `wholeBlocks`). A formula part that is all of the formula drags as the formula
  itself (`mixedSelection.ts` `wholeFormulasRange`).
- A passive field (a formula showing its part) is `inert`: a press on it is the document's, so a selection that
  takes a whole formula can be dragged from the formula too.
- An open formula's selection dragged out of the formula (the kernel's selection drag, `SelectionDragHost.outside` /
  `dropOutside`) lands in the document (`formulas.ts` `FormulaDragOut`, drop caret from `dropPlace.ts`): all of the
  formula moves the formula itself (between blocks for a display formula); a part becomes a formula of its own
  (inline in a line, display between blocks) and the formula keeps the rest (left empty: it goes). One undo step;
  Ctrl copies; Esc cancels. Inside the formula the drag moves within it as before. Dropping into another (closed)
  formula is not supported: it lands beside it.

## 8. Pages, body text, footnotes and citations (2026-11-01)

The user: "the document itself should be A4 (the earlier infinite canvas was a mistake); by default as wide as the
window, but adjustable; font sizes selectable; Chinese and Latin body fonts from the computer's fonts; citations and
footnotes". The bottom bar's source tab is now "源码 / Source" in document mode (it shows LaTeX or Markdown).

### 8.1 Settings (`settings.ts`)

Paper (A4 default; A5, JIS B5, Letter), margins in mm (Word's 25.4 / 31.8), body size in pt (五号 10.5 default;
Chinese named sizes and points), CJK font (SimSun) and Latin font (Times New Roman), line spacing (multiples of
ctex's single spacing, 1.3 x the size). They are the doc node's attributes: saved in .vtxdoc and autosave, undoable
(`setDocumentSettings` = `setDocAttribute` steps); null = the default, so older documents keep working. The CSS
body font is "Latin, CJK, fallbacks" (Latin text from the Latin font, Chinese falls through to the CJK font).
The font menus list the system's families: Tauri `list_system_font_families` (system_math_glyphs.rs: the existing
font catalog, English typographic family + Chinese name, "has 中" flag, read off the main thread; on macOS the same
command lists CoreText's families with their names in the user's language); without Tauri a list of common Windows /
macOS fonts (app `systemFonts.ts`).

### 8.2 Page view (`pagination.ts`, app `DocumentEditor.tsx`)

The editor lies on a sheet the paper's width at the zoom (96 px per inch x zoom), padded by the margins, its body
size/fonts/spacing from the settings. Zoom: "fit width" (default: the page fills the window's width, recomputed on
resize) or the +/- steps (turn fitting off; the status bar's ↔ button turns it back on). The sheet's background
repeats white pages and 16 px gaps. `paginationPlugin` measures after each change (debounced, one frame): units are
the lines of textblocks (paragraphs, headings, reference entries; also inside lists and quotes) and whole blocks
(display formulas, tables, code, rules); positions are taken without the spacers already placed; a unit that would
cross the text area's foot (or the room its page's footnotes take) is pushed to the next page by a spacer widget —
between blocks, or inside a paragraph at the start of the line (an inline `display:block` span; the lines before it
do not reflow). A heading's last line goes with what follows it. Overlong blocks run over. The page layer (an overlay
on the sheet) prints each page's footnotes at its foot (above them a short rule) and "第 n 页，共 m 页" in the bottom
margin; the sheet's height is the pages'. Editing across a break: arrows, typing at the continued line's start,
clicks in the gap (tools/pw/page4.mjs).

Margins (round 2, the user: "left/right and top/bottom margins adjustable"): besides the page-setup panel, the text
area's four edges can be dragged on every page (app `MarginGuides.tsx`, as on Word's ruler): the guide shows the new
edge and "左边距 25.4 mm" while dragging (a preview of the settings, the sheet relays out live); the release sets the
margin in one undo step. 0.1 mm steps; the text area keeps at least 30 mm (`clampMargin`). Exact values: the page
setup panel's four mm fields (they show the current margins, also after a drag; Enter or leaving the field sets the
value, Esc puts it back, ↑/↓ step 0.1 mm, Shift 1 mm) and Word's presets (常规 / 窄 / 适中 / 宽); a double-click on a
grip opens a mm field at the pointer (Enter sets, Esc cancels).

Font size: one toolbar control (字号). Its menu sets the body size (the document's, §8.1: Chinese named sizes or
points) and, for the selected text, LaTeX's relative sizes (\tiny … \Huge, each with the size it gives at the
current body size; a `fontSize` mark). The button shows the body size, or the mark (e.g. \large) inside sized text.

### 8.3 Footnotes and citations (`notes.ts`)

- `footnote`: an inline atom with its note's text (`$…$` for formulas); its number is its place among the footnotes;
  the page view prints the note at the foot of the page its mark is on. Toolbar ⁽x²⁾ button, Ctrl+Alt+F, `\footnote{…}`
  (with the argument: that text), `/footnote`; the app's FootnoteDialog edits it (a click on its number or its note).
- `bibliography` (the reference list, `bib_entry+`, titled 参考文献/References by CSS) and `cite` (inline atom with
  the entries' ids, drawn [1] / [1, 2] / [1–3] by list order; a missing entry [?]; Ctrl+click goes to the entry).
  The toolbar's citation picker (and `\cite`, `/cite` at the caret): tick entries and cite them, or type a new
  reference and add it (and cite it). `\begin{thebibliography}` / `/references` adds an empty entry, caret in it.
- LaTeX: `\footnote{…}` (formulas as `\(…\)`), `\cite{key,…}`, `\begin{thebibliography}{99} \bibitem{key} …`
  (key: the entry's own, else `ref-<id>`). Markdown: `[^n]` with `[^n]: text` after the document, `[n]` citations,
  the list as `[n] text` lines under a bold title. Plain text: `[^text]`, `[n]`.

### 8.4 Numbering formats

The toolbar's 编号格式 (#) panel; doc attributes like the rest of §8.1. Equations: `1` (through the document),
`1.1` or `1-1` (by section: the section's number, then the equation's within it; equations before the first
section are 0.n); brackets `()` or `[]`. Headings: `1.` / `1` (numbered, with or without the stop) or none.
`equationLabels(doc)` gives every numbered formula its label (numbers, `\eqref`s, the Markdown / plain-text
exports use it). LaTeX: `\numberwithin{equation}{section}` (`1.1`), `\renewcommand{\theequation}{\thesection-\arabic{equation}}`
(`1-1`), a `\tagform@` override for `[]`, `\setcounter{secnumdepth}{0}` for no heading numbers. Markdown: `\tag*{[n]}`.

### 8.5 Alignment in multi-line formulas

As in the formula mode: `&` asks the kernel (`ampersandAction(model, keys, latex, alignRows)`, alignRows = the
formula is a display). In a plain display formula it leaves an alignment marker; Enter then splits the row there
into `aligned` rows (`gathered` only if the formula already is one). Inside aligned rows `&` places the row's
alignment point, or moves to the next column when the row has one. Inline formulas get the formula mode's inline-row
answer. `\begin{align}` / `/align` insert an aligned template with the caret in its first slot (tools/pw/align6.mjs).

### 8.6 Export preamble

Always the paper and margins (`geometry` with paperwidth/paperheight); the body size (ctex `zihao=` for named
sizes, class `10/11/12pt`, else `scrextend`), spacing (ctex `linespread=`, article `\linespread`) and fonts
(`\setmainfont`, `\setCJKmainfont`; fontspec for article) only when the document chose them; a font or ctex
writes `% !TEX program = xelatex` first.

## 9. Milestones

1. Islands: schema, static/open formulas with FormulaField, edge rules (core `horizontalExit`), undo, input
   rules, paste; dev page `document-lab.html`; headless walk tests.
2. Blocks: headings, lists, quote, slash menu, Markdown line-start rules, turn into, block handle.
3. Tables: insert, Tab navigation, add/remove rows/columns, merge/split, formulas in cells.
4. LaTeX export + numbering; `.vtxdoc` open/save.
5. App integration: Formula / Document switch in the main window, toolbar inserts into the open formula (or a
   new inline one), Windows build + NSIS.

## 10. macOS (2026-11-01)

The macOS app has the same document mode: `apps/macos/src/document` and the stores are the Windows files, the
document package and the editing core are shared (`apps/shared`). The macOS formula editor (`editor/MathEditor.tsx`)
follows the Windows one (AxMath rows: row hit rule, a drag across rows, Shift+press, dragging a rows selection or a
formula's selection to another row, edge scrolling; the exports the document needs) and keeps the macOS parts: the
session history (`useHistoryManager`, `captureDocumentSnapshot` / `restoreDocumentSelection` with the selection
across rows), typing, pasting or deleting over a selection across rows replaces all of it, one toolbar-insertion
attempt (`visualTexInsertToolbarTemplate`), copy in the row's code format, the WebKit focus fix, the opt-in IME
trace. `\bm` is the kernel's own command on both. The body-font menus read the system's families on macOS through
CoreText (`list_system_font_families`).

## 11. Images (2026-11-04)

Research behind every rule here (measured with TeX Live 2023, GitHub's renderer, Obsidian, Pandoc; Mathcha probed
live): VPS `~/claude/document-images/research/` (`latex-image-layout.md`, `markdown-image-layout.md`,
`interaction-survey.md`). The requirement: what the user arranges with the mouse is where the exported LaTeX (and,
as far as Markdown allows, the Markdown) puts it when compiled elsewhere. So every arrangement is one deterministic
LaTeX construct, and the editor lays it out by the same rule.

### 11.1 Model: an image is a box in a line, as `\includegraphics` is

`image` is an inline atom (`images.ts`): `src` (a key into the document's images, `vtx-image:<sha-256>.png|jpg`),
`alt`, `width` (a fraction of the line width — LaTeX's `\linewidth` — to 0.001), `wrap` (null, `left`, `right`).
TeX's `\includegraphics` is a box whose bottom sits on the baseline; the editor's `<img>` is an inline box with
`vertical-align: baseline`. The arrangements follow from where the box is:

| Arrangement | In the document | LaTeX | Markdown |
|---|---|---|---|
| One image on its own line | a paragraph holding only the image (an **image line**); its alignment is the paragraph's (centre when inserted) | centre `{\centering\includegraphics[width=f\linewidth]{figures/x.png}\par}`; left `\noindent\includegraphics[..]{..}\par`; right `{\raggedleft\includegraphics[..]{..}\par}` | `<p align="center"><img src alt width="f%"></p>` (left/right: `align="left|right"`) |
| 2-3 (or more) images on one line | an image line holding several images, bottoms on one baseline, `0.02\linewidth` between neighbours | the same, neighbours joined by `\hspace{0.02\linewidth}%` | the same `<p>` with the `<img>`s separated by a space |
| An image in a sentence | the image among the text: one character | `\includegraphics[width=f\linewidth]{..}` in the text | `<img src alt width="f%">` in the text |
| Text wrapped around an image | the image (`wrap` = side) first in a body paragraph that has text; the paragraph's lines (and the next paragraphs') run beside it | a blank line, `\begin{wrapfigure}[N]{r|l}{f\linewidth}\centering\includegraphics[width=\linewidth]{..}\end{wrapfigure}`, the paragraph | `<img src alt width="f%" align="right|left">` at the start of the paragraph's line |

- `\centering` / `\noindent` (never `center`): they add no vertical space, so the image line is spaced like any line
  (previous descender + `\lineskip` above, `\baselineskip` to the next baseline), as the editor's line box spaces it.
- Widths are always written (TeX's natural size depends on DPI metadata and engine). A line's images plus gaps are
  kept within the line (≤ 0.99): an image dropped beside others scales the line's images to one height to fit; a
  resize stops at the room left. (Inside `\centering` TeX would silently break an overfull line onto two — so would the
  browser — but the editor never makes one.)
- Wrapped images: the box is placed as wrapfig places it mid-page (its top `\intextsep − 1.4ex` ≈ 0.55em below the
  first line's baseline, `\columnsep` = 10pt beside the text, `\intextsep` = 12pt below); wrapfig counts the
  narrowed lines itself (the export's fonts set the lines, so the image is never overrun). Wrapping goes on into the
  next paragraphs as wrapfig's does; a heading, list, quote, table, display formula, code, rule, image line or the
  next wrapped paragraph starts below the image in the editor (`clear`), and the export puts `\vtxwrapclear` before
  it (defined in the preamble: wrapfig's remaining lines, `\c@WF@wrappedlines − 1`, as `\vspace*`, then
  `\WF@finale`). Wrap is offered only for body paragraphs (wrapfig fails in lists, quotes, tables). A wrapped image
  that would cross the page's foot takes its paragraph to the next page (wrapfig's forced break).
- Images are not allowed in table cells, headings, the title, footnotes or code.
- Packages: `graphicx` (always with images), `wrapfig` (with a wrap). Files: `figures/fig-1.png`, numbered by first
  appearance (one file per distinct image); the export with images is a `.zip` (name.tex + figures/, Overleaf's
  upload), Markdown `.zip` (name.md + assets/). Copy LaTeX/Markdown writes the same paths.
- Markdown (GitHub keeps `width` in %, `align`, `alt`; strips `style`; `align` on `<img>` does not centre, so lines are
  `<p align>`; one line per construct, blank lines around).

### 11.2 Images and their files

The bytes live beside the document, not in its nodes: `.vtxdoc` gets `images: { key: base64 }` (only the keys the
document uses; `version` 2 when it has any, so an older VisualTeX refuses it instead of losing them). In the editor
a registry holds key → bytes, type, pixel size, object URL. Inserted images are stored as PNG or JPEG (what every TeX
engine reads): JPEG stays JPEG, PNG stays PNG, anything else (GIF, WebP, BMP, SVG, HEIC as the webview decodes it)
is drawn into a PNG. The key is the content's hash, so the same image twice is one file. Copy within the app carries
the image as a `data:` URL in the HTML, so a paste into another document finds it.

### 11.3 Gestures

- Insert: paste or drop an image file into the text, the toolbar's image button (file dialog) or `\includegraphics`
  in the `\` menu (Markdown: `/image`). In an empty paragraph the image takes it; at a line's end or in its middle the
  image gets an image line after that paragraph (Mathcha; the sentence is not split); in an image line it joins the
  line. Width: its natural size (96 px per inch at zoom 1) up to half the line. Pasting an image in an **open formula**
  still recognises it (OCR); in the text it is always inserted (the user, 2026-11-04).
- Select: a click selects the image (outline, two side handles, a bar above: left / centre / right, wrap left / wrap
  right). A selected image: Backspace/Delete removes it, Enter opens a paragraph after its line, typing does not
  replace it.
- Resize: drag a side handle; the ratio stays. A centred line grows to both sides (the pointer keeps the edge), a
  left/right one from its free edge. A label shows "45%"; the width snaps to 25/33/50/67/75/100% and other images'
  widths within 1%; Alt turns snapping off; a double-click on a handle gives the inserted width back. One undo step.
- Move (`documentDrag.ts`): press on the selected image and move. The drop caret decides: between two blocks (a line
  across) → an image line of its own; beside an image on an image line (a caret the line's height) → joins that line
  (fitted); in text → a character there (a block-sized image dropped into text is made one line high: its height the
  text size). A wrapped image dragged over a body paragraph shows the paragraph's start on the nearer side: it wraps
  that paragraph on that side; elsewhere it stops wrapping. Ctrl copies, Esc cancels, one undo step.
- Wrap: the bar's wrap buttons. The image goes to the start of its paragraph (an image line wraps the next body
  paragraph); "centre / left / right" in the bar makes it an image line again before that paragraph.

## 13. Documents are Markdown files (2026-11-18)

Owner's decisions: documents are ordinary `.md` files in a visible folder (default `Documents/VisualTeX`, changeable);
a Markdown file from elsewhere is edited in place; nothing only VisualTeX reads is written into the file — what Markdown
has no syntax for is kept beside it ("the document's memory") and restored when VisualTeX opens the file again; other
editors see plain Markdown. VisualTeX never makes itself the default program for `.md`.

### 13.1 The file (`markdownFile.ts`, `markdown.ts` with `file: true`)

- The file is the document's lines, one per line. An object (display formula, table with its caption, row of
  pictures, and any block that is not a line) stands with one blank line on each side; that blank line is the file's
  spacing, not a line of the document. Read: a blank line beside an object, between it and another block, is taken
  as that spacing (never the file's first or last line).
- Inline formula `$…$`, display formula `$$` / LaTeX / `$$`, GFM table with Pandoc's `Table: …` caption line, picture
  `![alt](path)` (several on one line: an image line or board), footnote `[^label]` with the definitions as the
  file's last lines, title `# …`, author / date / abstract / caption lines as text, `[toc]`, a citation its
  number, a reference list `[n] entry` lines. A page command made in VisualTeX has no line in the file (decision 2
  lists page commands among what the memory keeps): the memory keeps it with the block after it (before it at the
  end; a document of nothing else has an empty line to keep it by). A file's own `\newpage`, `\pagenumbering{…}`…
  lines are read as page commands and written back as the file has them.
- A numbered formula is written with `\tag{n}` (`\tag*{[n]}` with square brackets). A file's own `\tag` stays the
  formula's own (no renumbering of what another program wrote); `\label{name}` names the formula (its id) and
  `$\eqref{name}$` refers to it; formulas made here are referred to by number, "(n)". The names a file gave are
  kept with its source (`DocumentFileSource.labels`), so a name shaped like the editor's ids (ten lowercase letters
  and digits, "definition") is written back as a name too, also by Save as.
- Formulas the dollars cannot say plainly: a word joiner (U+2060) between a formula and a digit or dollar next to
  it (the reader removes it), LaTeX of several lines on one, a closing backslash followed by ` {}`, `\$` inside a
  formula read as part of it. An empty inline formula writes nothing (the memory keeps it).
- Read from any Markdown: also `\(…\)`, `\[…\]` (on one line, or opened by a row that is only `\[`), tables
  without outer pipes, pictures in a line of text when their file is found. Footnotes are read only when every
  definition is among the file's last lines and has exactly one reference; otherwise they stay text.
- A block read from a file and not changed is written back as the file had it (its lines, the blank lines around an
  object); the document unchanged is the file's text itself. Line ends (CRLF / LF / CR), a missing last newline, the
  encoding (UTF-8 with or without a mark, UTF-16, GB 18030 with its two- and four-byte characters, and any other
  bytes read as Windows-1252, which gives every byte back) are kept; a character the file's encoding cannot hold
  makes the file UTF-8. In a file whose lines end in different ways every row keeps its own line end (an edited
  row the one of the row it replaces); an added row takes the line end most rows have.

### 13.2 The memory

JSON, `{ format: "visualtex-document-memory", version: 1, settings, blocks, images }`:

- `settings`: the doc node's attributes that are set (page, fonts, spacing, numbering formats).
- `blocks`: `{ hash, occurrence, lines, nodes }` — for a run of `lines` lines whose text alone does not read as the
  document's nodes, those nodes as ProseMirror JSON. `hash` is a 64-bit hash of the lines' text joined by `\n`;
  `occurrence` says which of the places in the file with that text (from 0). Ids made by the editor are left out
  unless a reference names them.
- `images`: image key → path, for the pictures the blocks show.

How entries are made: the writer reads its own text back as the text alone would read and compares, run by run,
with the document; every run that differs gets an entry. So the file and the memory together are the document by
construction, and a document that needs no memory has none. Read: an entry is applied where its lines still stand
as whole blocks; a line another program rewrote no longer has its text, so its entry finds no place and is dropped
(the rest stays). Two equal lines are told apart by `occurrence` only.

Where it is kept (`document_files.rs`): for a document of the documents folder, `.visualtex/<relative path>.json`
in that folder (hidden); for a file elsewhere, the app's data folder (`documents/memory/<uuid>.json`, indexed in
`documents/outside.json` by the file's path and the SHA-256 of its bytes, by which a moved or copied file is found
again while its text is unchanged). Rename, move, duplicate, trash and restore carry the record along.

### 13.3 Pictures

Files, named by the paths in the text. A file made elsewhere keeps every path; a picture without a file gets
`assets/<name of the file it came from>` or `assets/image-<date>-<time>.png` (pasted) beside the document, `-2`, `-3`
when the name is taken; a path once given stays. A picture file is never replaced by a save, nor by a move,
restore or folder deletion (another file of a picture's name refuses those; a save whose new picture's name another
program took meanwhile names the picture again; an imported archive whose pictures' names are taken goes into a
folder of its own). The image registry still holds the bytes by content hash; `DocumentFile.imagePaths` maps key →
path for the open document.

### 13.4 The library (`document_files.rs`, `document_library.rs`, `documentLibraryApi.ts`)

The documents folder's `.md` / `.markdown` (and `.vtxdoc`) files; its folders are its subfolders (`assets` and hidden
folders aside); a document's id is its relative path. `.trash/<uuid>/` holds a deleted document with its record,
pictures and where it was, for 30 days. Files elsewhere are listed as recent files (`file:<path>` ids); taken out
of that list, a file keeps its record (`unlisted` in `outside.json`) and is listed again when opened.

Saving (`DocumentWorkspace.tsx`, `documentFiles.ts`): 400 ms after a change and on leaving, one save after another.
Each save carries the fingerprint the editor read; the backend refuses it when the file's bytes are others now.
Coming back to the window the file is checked: changed on the disk only → read again; changed on both sides → the
user chooses Reload, Overwrite or Save a copy. The documents folder is listed again then too (files other programs
added, moved or deleted). Writes go through a temporary file beside the target; a link is written through to the
file it points at, a file keeps its permissions, and a read-only file is not written.

### 13.5 Earlier documents (`documentMigration.ts`)

The first time document mode opens, each `.vtxdoc` of the library in the app's data folder is written as Markdown
and memory, read back and compared (`sameDocument`); the same → a `.md` in the documents folder (its folder a
subfolder, its time kept); otherwise the `.vtxdoc` itself is put there and reported (until the report is answered;
one kept in a later run, after a run cut short, is reported then). The earlier library is left where it is; its
documents in its own trash are not converted.
Section numbers: off unless the document says so; a `.vtxdoc` written before file version 7 gets `"1"`.

### 13.6 Files from the system

A queue in the backend (`queue_files`, `document_pending_files`, event `document-open-files`); the page takes the
paths at its start and when told, switches to document mode and opens them in place. Windows: start arguments and
the single-instance notification; "Open with" through a ProgID in `OpenWithProgids` and the `Applications` key
(`hooks.nsh`). macOS: `CFBundleDocumentTypes` with `LSHandlerRank` Alternate, file URLs of `RunEvent::Opened`.
