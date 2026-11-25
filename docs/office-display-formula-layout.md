# Display formulas in Word and WPS Writer (Windows): the line layout and its repair

Status: 2026-11-23. Measured in Word 16 (Windows) with the runs in §4 and checked in WPS Writer 12.1.0.29681; the Mac
add-ins (Word for Mac, WPS for Mac) repair their own formulas the same way (§5).

## 1. The line of a display formula

| Formula | Line |
|---|---|
| VisualTeX OLE | TAB, object, TAB, number (MACROBUTTON VisualTeXPlaceRef); centre stop at the column's centre, right stop at its right edge; justified; no indents; the object lowered by its own render baseline (Font.Position of its U+0001 character) |
| MathType | the same tab layout (MathType's own), number MACROBUTTON MTPlaceRef; the object lowered by MathType's baseline |
| Word equation (OMML) | oMathPara centred by Word; a numbered one is an equation array "formula # (SEQ)"; no indents |

The stops are those of the column (newspaper columns) or table cell the formula is in (`WordDisplayColumns`).

## 2. What users' whole-document changes do (Ctrl+A, then …)

Measured on a document with 9 VisualTeX, 2 MathType and 3 Word display equations (P3/P20 VisualTeX, P25 OMML, P27
MathType in the logs):

- First-line or left indent, in points or characters (首行缩进 2 字符): a tab-laid formula and its number move by the
  indent (dx −21 pt). A Word equation stays centred (Word ignores the indent) but keeps the indent.
- Exact line spacing (固定值): every formula taller than the line is cut off (h 27–44 pt > 20 pt), Word equations too.
  "At least", multiple and 1.5 lines change nothing.
- Ctrl+Q (paragraph reset), the Normal style, the format painter from a body paragraph, Clear All Formatting: the
  line's tab stops go; the formula moves to the left margin, the number right behind it or onto the next line.
- Clear All Formatting and Ctrl+Space also set every object's Font.Position to 0: VisualTeX and MathType formulas rise
  off the baseline (MathType −10/−14 pt → 0).
- Bullets or numbering: a list indent on every line; Word turns its display equations into inline ones (the repair
  leaves those, and an unnumbered MathType object in a list paragraph counts as inline — `MathTypeOleInterop`).
- Centred/left/distributed alignment does not move a tab-laid line (the tabs place it); page margins move the column,
  so the old stops are wrong.

## 3. The repair (修复行间公式 / Repair Display Formulas)

`WordFormulaService.RepairDisplayFormulaLayout`, Ribbon `OnRepairDisplayFormulas`, MCP `repair_display_formulas`.
Every display formula of the main text is read first; a document without a fault gets no Undo record. Faulty lines
are then put back from the end of the document to its start in one Undo record:

1. list numbering removed; 2. indents (points and character units) zeroed; 3. alignment: tab-laid lines justified
(left is accepted), Word equations centred; 4. tab-laid lines: the column's centre and right stops alone (no other
custom stop before the right one); 5. exact spacing becomes "at least" the same value; 6. the TAB before a tab-laid
formula; 7. a VisualTeX object's Font.Position from its render baseline for the height VisualTeX gave it (Word keeps
the size to the twip, which moves the offset by a point for some heights: `WrittenHeight`); 8. a MathType object
that is not lowered (position ≥ 0) gets MathType's own offset from VisualTeX's native MathType renderer
(`MathTypeNativePreviewRenderer`, the WordPosition of the unchanged MTEF scaled to the shown height); an object that
is lowered keeps its offset; no renderer, no check.

Line spacing otherwise stays the paragraph's own; fonts and sizes of numbers are not touched (a number formatted by
hand stays).

## 4. Runs and kits (PC `.cache\dispfix`)

- `repairtest.ps1` (modes from `modify.ps1`, measurements `measure.ps1`), Word via `.cache\xref\run-e2e.ps1 -Script
  dispfix.mjs -Mode rt1.json|rt2.json`: every mode returns to the base layout, Undo restores the change, a second run
  reports nothing.
- `wps1.json` with `e2e-wps-dispfix.ps1` (WPS through MCP); `ribbonshot.ps1` / `wpsribbon.ps1` for the Ribbon.
- WPS writes malformed XML for some equations it read from a Word file (a numbered OMML whose SEQ became
  "错误！未定义书签"); OMML insertion and 更新公式编号 already failed there before this work.

## 5. The Mac add-ins

Word for Mac (`VTRepairDisplayFormulaLayout` in VTWordNumbering.bas, macro `VisualTeX_RepairDisplayFormulas`, Ribbon
group 兼容修复, agent op `repair_display_formulas`) and WPS for Mac (`repairDisplayFormulas` in office/wps/module.js,
button `vtxRepairDisplayFormulas` with WPS's AlignCenter icon: CenterAlignTabs exists only in small sizes there) apply
the rules of §3 to the formulas the Mac add-ins write; the Mac add-ins do not read OLE objects (MathType, the Windows
VisualTeX OLE), so those are left alone. Their lines differ from Windows in two ways:

- An unnumbered picture formula is a centred paragraph without tab stops (VTNormalizeUnnumberedDisplayParagraph), so
  it is put back centred; a numbered one has the TAB, picture, TAB, number line of §1 with its stops at half the
  column width and 1 pt inside its right edge (VTConfigureNumberedEquationParagraph). A numbered picture without its
  number in the line (an older numbered table) is left to its table.
- A picture is lowered by its ink centre against its number (or the body text): VTExpectedImageFormulaPosition in Word,
  displayPosition in WPS (for the height VisualTeX gave it, from VT_ImageScale_: WPS reports sizes in half points).
  A picture within 1 pt of that keeps its offset.

Messages: English in Word for Mac (its add-in's language), Chinese on the WPS ribbon, English in agent answers.
