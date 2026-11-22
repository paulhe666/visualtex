# Formula font size in Word and WPS Writer (Windows): expected behaviour

Status: 2026-11-22. Each numbered expectation below was checked in Word (Windows) by the runs in §6, the sizes and
the layout in WPS Writer too; a deviation is a bug unless this document is changed. The Mac add-ins are not checked
(§8).

The Ribbon group 「公式字号 / Formula size」 has two kinds of controls:

| Control | Acts on | Kept |
|---|---|---|
| 「行内 / Inline」 box | every inline formula of the active document | the choice (per user), for new inline formulas |
| 「行间 / Display」 box | every display formula of the active document, numbered ones included | the choice (per user), for new display formulas |
| 「字号 / Size」 box, 「增大 / Larger」, 「减小 / Smaller」 | the selected formula(s) only | nothing |

Both kind boxes offer 「跟随正文 / Follow text」, the Chinese sizes (初号 … 八号) and point sizes; a size can also be
typed (5–200 pt). A narrow Word window folds the group into one 「公式字号」 drop-down holding the same controls.

## 1. Words

- **Formula**: a VisualTeX OLE object, a Word equation (OMML) or a MathType object, inline or display.
- **Display formula**: a formula alone on its line — a VisualTeX/MathType display paragraph (tab, formula, tab,
  number), a display Word equation (oMathPara), any numbered equation. Everything else is **inline**. (A MathType
  object carries no display flag: one alone in its paragraph counts as display.)
- **Text size at a formula** ("follow text"):
  - inline formula: the size of the nearest visible text beside it in its paragraph, the character before it first;
    a formula alone in its paragraph is treated as a display formula;
  - display formula: the size of the nearest visible body text before its paragraph (after it when there is none);
  - other formulas, the whole paragraph of a display formula (its tabs and its number) and headings (paragraphs with
    an outline level) are stepped over: a formula never takes the size of a neighbouring formula, of a number a
    batch has just resized, or of a heading; nothing found: the size of the document's Normal style.

  LaTeX redraw (and bulk import, which redraws) gives the formulas it writes the same size when the kind box follows
  the text (`WordSourceFontSize` for the text beside the source, `WordFormulaTextSize.TrySizeBelowHeading` for a
  display formula right below a heading); 「跟随正文」 on an existing formula is `WordFormulaTextSize.FollowSizes`.

## 2. The kind boxes (行内 / 行间)

E1. Choosing a size in 「行间」 gives that size to every display formula of the document — VisualTeX, Word
    equation and MathType alike. Inline formulas do not change.

E2. Choosing a size in 「行内」 gives that size to every inline formula; display formulas do not change.

E3. Choosing 「跟随正文」 in either box gives every formula of that kind the text size at the formula (§1): after
    「行内：跟随正文」, an inline formula in a 小四 paragraph is 12 pt, one in a 五号 paragraph 10.5 pt, one in a
    小五 paragraph 9 pt. The box then shows 「跟随正文」.

E4. **Office limitation.** Choosing the value the box already shows does nothing: Office reports a choice in a
    Ribbon combo box only when its text changes. To bring formulas pasted or resized since then back in line,
    choose another value and then the wanted one again.

E5. The number of a numbered display formula has the formula's size, whatever the format.

E6. One Undo restores every VisualTeX formula and Word equation (with its number) the choice changed. MathType
    objects are rewritten one by one (a MathType rewrite reads the object's XML, which ends an open Word Undo
    record): one Undo per MathType formula, its number included.

E7. The status bar says how many formulas changed; a formula that could not be changed is named with the reason,
    and the others are still changed.

E8. The boxes are disabled while VisualTeX is writing to the document, and show the current choice otherwise.

## 3. New formulas

E9. While a kind box holds a size, every new formula of that kind is written at that size: inserted from the
    VisualTeX editor (Insert), by LaTeX redraw, bulk import, format conversion, or an agent (MCP) — unless the
    request names a size.

E10. While a kind box holds 「跟随正文」, a new formula takes the text size where it is written (§1).

E11. Editing an existing formula keeps its size.

## 4. The selected formula (字号 / 增大 / 减小)

E12. The 「字号」 box shows the selected formula's size, MathType included; it is empty and disabled when no
     formula is selected. A MathType object's size is read from its own data (exporting it makes Word save the
     object, which would re-enter MathType's program while that builds or edits the object): the box reads it only
     while no MathType program runs in the user's session (VisualTeX's own isolated MathType renderer does not
     count) and shows nothing otherwise; the size commands always read it.

E13. A size chosen there, or 「增大」/「减小」 (next/previous preset), changes only the selected formula(s) and its
     number; one Undo restores it (a MathType formula with its number too).

## 5. What must not change

E14. Layout: an inline formula stays on the text baseline with the spaces around it; a display formula stays
     centred on its centre tab (the column's centre, the cell's centre in a table) and on the baseline of its tabs
     and number, its number stays right-aligned on the right tab, formula and number stay on one line when they fit,
     and the paragraph keeps its alignment, indents, tab stops and spacing. This holds after any sequence of size
     changes (16 → 9 → 10.5 pt gives the original layout back).
     A display formula's line has no first-line indent, whatever the body text's style (「首行缩进 2 字符」): on a
     line placed by tab stops Word moves the formula and the number by that indent.
     Line spacing as Word and MathType do it: a display formula's paragraph keeps the body text's line spacing
     and space before/after (the paragraph it is written into, or Normal for a new one), in every format and for
     inserts, redraws and conversions alike. VisualTeX changes neither the paragraph above nor the one typed after
     it. The MTDisplayEquation style VisualTeX creates, like MathType's own, adds only the centre and right tabs to
     the body text's formatting; an existing MTDisplayEquation style is left as it is. Exact line spacing smaller
     than a formula therefore clips it, as it does for Word's and MathType's own equations.
     Not covered: a formula too wide to leave room for its number pushes the number to the next line (Word's tab
     rule; it comes back at a smaller size), and exact line spacing clips a formula taller than the line (Word).

E15. Content: the formula's LaTeX, identity, numbering, references, colour and fonts stay as they were; a VisualTeX
     formula opened in the editor after a size change opens at the new size.

E16. Other text, other formulas of the other kind, tables and section layout are untouched.

## 6. Test runs (Word, Windows)

Setup: Word 16 (Microsoft 365, 64-bit) on the Windows dev PC; the add-in built from this change loaded into a
separate Word process (temporary HKCU class override), an isolated companion app, steps sent through MCP
(`set_formula_font_size` with `kind` / `follow_text` / `formula_id`, `insert_formula`, `latex_text_to_formulas`,
`convert_formulas`), the Ribbon boxes driven through UI Automation (the folded group's drop-down, then the box's
list item), Undo through `Document.Undo`. After each step: every formula's kind, number and size (MCP
`list_formulas`), and per paragraph the tab stops, object positions and a window capture.

Fixtures (one per format: VisualTeX OLE, Word equation, MathType; 13 inline and 8 display formulas, 7 numbered):
body text 五号 10.5 pt with four inline formulas, display formulas after body text, a 小四 12 pt paragraph with two
inline formulas and a display formula after it, a heading directly followed by a display formula, a 小五 9 pt
paragraph, a table with an inline and a display formula, a two-column section with two display formulas. Three more
source variants carry a document grid with a 2-character first-line indent, plain / exactly 20 pt / 1.5 lines.

| Exp. | What was done | Result |
|---|---|---|
| E1, E2 | 行间 16 → 9 → 10.5, 行内 14, in all three formats | pass: only the chosen kind changes |
| E3 | 行间 16, then 跟随正文; 行内 14, then 跟随正文 (MCP and Ribbon) | pass: display 10.5 after body text, 12 after the 小四 paragraph and below the heading that follows it, 10.5 in the table; inline 10.5 / 12 / 9 by paragraph. Before the fix 跟随正文 changed no formula |
| E3, E8 (Ribbon) | 行间 shows 三号; pick 跟随正文 in its list | pass: all 8 display formulas follow the text, the box shows 跟随正文 and stays enabled |
| E4 | 行间 小四, one formula set to 18, pick 小四 again; same with 跟随正文 in 行内 | the formula stays at 18: Office limitation (E4) |
| E5 | display batches in all formats | pass. Before the fix MathType numbers stayed 10.5 |
| E6 | Undo after each batch | pass for VisualTeX and Word equations (one Undo); MathType one Undo per formula (before the fix two: formula, then number) |
| E7 | status texts | pass: 「已把 8 个行间公式改为 三号（16 磅）。」, 「已让 13 个行内公式跟随正文字号。」 |
| E9 | 行间 12, then `insert_formula` (display, numbered) without a size; LaTeX redraw with 行间 10.5; OMML → VisualTeX conversion with 行间 14 / 行内 12 | pass: 12; 10.5; 14 / 12. Before the fix an editor/agent insertion took the typing size |
| E10 | both boxes 跟随正文, LaTeX redraw of the source fixture, then 跟随正文 on the result | pass: both give 10.5 / 12 / 9 and 12 below the heading. Before the fix redraw gave that formula the heading's 16 pt |
| E11 | `update_formula` with new LaTeX for a 16 pt display and a 10.5 pt inline formula (VisualTeX, MathType) | pass: both keep their size |
| E12 | select a 16 pt display formula, read the 「字号」 box; put the caret in body text | VisualTeX: 三号, then empty and disabled: pass |
| E12 | click (posted to Word's document window) and COM selection of a 10.5 pt MathType and VisualTeX formula | pass: 五号 for both; the MathType read takes 22 ms once per selection. Before the fix a MathType formula left the box empty |
| E13 (Ribbon) | 增大 once, 减小 twice on the selected formula; Undo | pass: 16 → 18 → 16 → 15 (VisualTeX and MathType), one Undo per press |
| E12, E13 | one display (18 pt) and one inline (16 pt) formula through `formula_id`, VisualTeX and MathType; Undo | pass: only that formula and its number; one Undo restores both (MathType too) |
| E14 | geometry after every step above | pass: display formulas within ±0.2 pt of the centre tab, numbers on the formula's baseline at 9, 10.5, 16 and 18 pt in all formats; 16 → 9 → 10.5 gives back the original positions. Before the fix a resized MathType display formula moved above (larger) or below (smaller) its tabs and number. At 16 pt the widest two-column formula (191 pt in a 215 pt column) leaves no room for its number, which goes to the next line |
| E14 spacing | body text 1.5 lines + 3 pt before / 6 pt after, and exactly 20 pt + 3/6 pt (grid, 2-character indent): LaTeX redraw (numbered) and new display formulas (numbered and not) in each format | pass: every display paragraph has the body's line spacing and space before/after; MathType's through MTDisplayEquation (tabs only); centres within ±0.2 pt. Before: MathType lines single spacing with 0/0 (VisualTeX created MTDisplayEquation that way), new Word-equation and VisualTeX display formulas 0/0, and inserting an OLE display formula set the paragraph above to 0 pt after |
| — | older VisualTeX objects without Word's cached data (AlternativeText cleared), 行间 14, 行内 12, 跟随正文 | pass: 8 display / 13 inline. Before the fix their display formulas were counted as inline |
| E14 | grid + 2-character first-line indent (plain, exactly 20 pt, 1.5 lines): LaTeX redraw to VisualTeX, MathType and Word equations, OMML → VisualTeX conversion, then 16 pt and 跟随正文 | pass: display formulas within ±0.2 pt of the centre tab, numbers at the right margin. Before the fix every VisualTeX display formula written there (redraw, conversion, renumbering) kept the paragraph's 21 pt first-line indent, and formula and number both sat 21 pt to the left, before any size change. With exactly 20 pt line spacing (kept from the source paragraph) Word clips formulas taller than the line |
| E14 | a document written by the previous build in that layout: 更新公式格式 | pass: first-line indents cleared, every display formula back on the centre tab |
| E5, E9 | 行间 12, `insert_formula` of a numbered MathType display formula; LaTeX redraw to MathType with 行间 16 | pass: the numbers have the formula's size (12; 16). Before the fix a new MathType number took the paragraph's 10.5 |

WPS Writer 12.1 (same fixture, MCP steps): 行间 16, 跟随正文, 行内 14, 跟随正文, one formula at 18, 行间 12 and a new
numbered display formula give the same sizes as Word, and the display formulas stay centred (±0.2 pt) with their
numbers on their baseline.

## 7. Formula fonts and how big they look

The fonts differ in design: at one point size their letters do not have one height or width. Measured in Word at
10.5 pt (400 % zoom, ink height in pixels, `fszG` fixture; Times New Roman italic text on the same line as the
reference):

| At 10.5 pt | x | α | θ | α width |
|---|---|---|---|---|
| Times New Roman italic (text) | 39 | — | — | — |
| OMML, Cambria Math | 41 | 41 | 59 | 44 |
| OMML, TeX Gyre Termes Math | 38 | 38 | 55 | 40 |
| OMML, STIX Two Math | 41 | 41 | 62 | 44 |
| OMML, Latin Modern Math | 38 | 38 | 60 | 46 |
| MathType (Latin in Times New Roman, Greek in Symbol) | 39 | 40 | 60 | 50 |
| VisualTeX OLE | 40 | 40 | 63 | 48 |

- Switching 「OMML 字体」 changes no size (the equations stay at 10.5 pt); Cambria Math and STIX Two Math are drawn
  about 7 % taller and up to 20 % wider than TeX Gyre Termes Math and Latin Modern Math (font metrics: math-italic x
  0.483 / 0.489 em against 0.452 / 0.453 em).
- MathType at a size is that size: its Latin letters are Times New Roman's, the same as Termes. Its Greek letters are
  Symbol's, which are taller and wider than Termes' Greek (α 0.476 em against 0.457 em, θ 0.709 against 0.660) and
  about as large as Cambria's: next to MathType, Termes looks a size smaller in formulas with Greek letters, and
  Termes at 11 pt matches MathType's Greek at 10.5 pt.

## 8. Not checked yet / open

- The Mac add-ins (Word, WPS) have their own size code; not compared (the Mac was not reachable).
- The kind choice is kept per user (HKCU), not per document: opening another document keeps the choice, but does not
  apply it to that document's formulas until a box is used.
- VisualTeX's table number format (formula and number in a borderless 1×3 table) keeps its compact cells; it has
  no Word or MathType counterpart.
