import { Atom, isCellBranch } from '../core/atom';

import { register as registerCommand } from '../editor/commands';

import type { _Model } from './model-private';
import { ArrayAtom, type ColumnFormat } from '../atoms/array';
import { Environment, TabularEnvironment } from '../public/core-types';
import { makeEnvironment } from '../latex-commands/environments';
import { PlaceholderAtom } from '../atoms/placeholder';
import { LeftRightAtom } from '../atoms/leftright';
import { range } from './selection-utils';
import { LETTER_SHAPE_RANGES, LETTER_SHAPE_MODIFIER } from '../core/modes-math';
import { ungroup } from '../editor-mathfield/styling';
import type { _Mathfield } from '../editor-mathfield/mathfield-private';
import type { PrivateStyle } from '../core/types';
import type { Selection } from '../public/core-types';

/** The bold/italic/color state the app asks new glyphs to be typed in. */
export type VisualTexPersistentTypingStyle = {
  bold: boolean;
  italic: boolean | null;
  color?: string;
  backgroundColor?: string;
};

/** A glyph's resolved weight and slant. */
type VisualTexFont = { bold: boolean; italic: boolean };

/** Style properties by name (the persistent-typing bookkeeping reads them by key). */
type VisualTexStyleRecord = Record<string, unknown>;

/**
 * Return the first and last siblings of the current cell
 */

function cellSiblings(model: _Model): [Atom, Atom] {
  let atom: Atom | undefined = model.at(model.position);

  while (atom && !isCellBranch(atom.parentBranch)) atom = atom.parent!;

  return [atom.firstSibling, atom.lastSibling];
}

/**
 * If we're inside an array, return that parent array.
 *
 * Otherwise, consider creating an array based on the context (sibling atoms).
 *
 * If at root, create a `lines` environment.
 * If left sibling is `(` (or `\left( with matching \right))`) create a `pmatrix`
 * If left sibling is `[` create a `bmatrix`
 * If left sibling is { create a `cases`
 * If left sibling is \vert `vmatrix`
 * If left sibling is \Vert `Vmatrix`
 * If other \left\right create a matrix (inside the left-right)
 *
 * If inside a leftright, use the content of the leftright as the initial cell
 */
function parentArray(
  model: _Model,
  where: 'after row' | 'before row' | 'after column' | 'before column'
): [ArrayAtom | undefined, [row: number, col: number]] {
  let atom: Atom | undefined = model.at(model.position);

  while (atom && !(atom.parent instanceof ArrayAtom)) atom = atom.parent;

  //
  // Conversion:
  // if we are in an array, but the addition of cell is not compatible,
  // change the array type
  //
  //
  if (atom?.type === 'array') {
    const array = atom as ArrayAtom;
    if (array.environmentName === 'lines') {
      // Convert to `split` if adding columns
      // @todo
    }
  }

  //
  // If no array was found, try to create one
  //
  if (!atom || !(atom.parent instanceof ArrayAtom)) {
    const cursor = model.at(model.position);
    atom = cursor;

    //
    // 1/ Handle insertion at the root (when the root is not already an array)
    //
    if (!atom.parent!.parent) {
      let secondCell = model.extractAtoms([model.position, model.lastOffset]);
      let firstCell = model.extractAtoms([0, model.position]);

      let array: ArrayAtom;
      if (where.endsWith('column')) {
        if (firstCell.length === 0) firstCell = placeholderCell();
        if (secondCell.length === 0) secondCell = placeholderCell();
        array = makeEnvironment('split', [[firstCell, secondCell]]);
        model.root = array;
        if (isPlaceholderCell(array, 0, 0)) selectCell(model, array, 0, 0);
        else if (isPlaceholderCell(array, 0, 1)) selectCell(model, array, 0, 1);
        else model.position = model.offsetOf(cursor);
      } else {
        // if (firstCell.length === 0) firstCell = emptyCell();
        // if (secondCell.length === 0) secondCell = emptyCell();
        array = makeEnvironment('lines', [[firstCell], [secondCell]]);
        model.root = array;
        // if (isPlaceholderCell(array, 0, 0)) selectCell(model, array, 0, 0);
        // else if (isPlaceholderCell(array, 1, 0)) selectCell(model, array, 1, 0);
        // else
        // Select the second line
        selectCell(model, array, 1, 0);
      }

      // We've created the environment and the cells, no need to add a row/column, so return undefined
      return [undefined, [0, 0]];
    }

    //
    // 2/ Are we inside a \left...\right...?
    //
    if (atom.parent instanceof LeftRightAtom) {
      const parent = atom.parent;
      let secondCell = model.extractAtoms([
        model.position,
        model.offsetOf(parent.lastChild),
      ]);
      let firstCell = model.extractAtoms([
        model.offsetOf(parent.firstChild),
        model.position,
      ]);
      if (firstCell.length === 0) firstCell = placeholderCell();
      if (secondCell.length === 0) secondCell = placeholderCell();

      let envName: Environment = 'pmatrix';
      const lDelim = parent.leftDelim;
      const rDelim = parent.rightDelim;
      if (lDelim === '(' && (rDelim === ')' || rDelim === '?'))
        envName = 'pmatrix';
      else if (
        (lDelim === '[' || lDelim === '\\lbrack') &&
        (rDelim === ']' || rDelim === '\\rbrack' || rDelim === '?')
      )
        envName = 'bmatrix';
      else if (lDelim === '\\vert' && rDelim === '\\vert') envName = 'vmatrix';
      else if (lDelim === '\\Vert' && rDelim === '\\Vert') envName = 'Vmatrix';
      else if (
        (lDelim === '{' || lDelim === '\\lbrace') &&
        (rDelim === '.' || rDelim === '?')
      )
        envName = 'cases';

      const array = makeEnvironment(
        envName,
        where.endsWith('column')
          ? [[firstCell, secondCell]]
          : [[firstCell], [secondCell]]
      );

      parent.parent!.addChildBefore(array, parent);
      parent.parent!.removeChild(parent);
      if (isPlaceholderCell(array, 0, 0)) selectCell(model, array, 0, 0);
      else if (where.endsWith('column')) {
        if (isPlaceholderCell(array, 0, 1)) selectCell(model, array, 0, 1);
        else model.position = model.offsetOf(atom);
      } else {
        if (isPlaceholderCell(array, 1, 0)) selectCell(model, array, 1, 0);
        else model.position = model.offsetOf(atom);
      }

      return [undefined, [0, 0]];
    }
  }

  return atom && atom.parent instanceof ArrayAtom
    ? [atom.parent, atom.parentBranch! as [number, number]]
    : [undefined, [0, 0]];
}

function isPlaceholderCell(
  array: ArrayAtom,
  row: number,
  column: number
): boolean {
  // const pos = model.offsetOf(array.getCell(row, column)![1]);
  // return pos >= 0 && model.at(pos).type === 'placeholder';
  const cell = array.getCell(row, column);
  if (cell?.length !== 2) return false;
  return cell[1].type === 'placeholder';
}

function cellRange(
  model: _Model,
  array: ArrayAtom,
  row: number,
  column: number
): [number, number] | number {
  const cell = array.getCell(row, column);
  if (!cell) return -1;
  return [model.offsetOf(cell[0]), model.offsetOf(cell[cell.length - 1])];
}

function selectCell(
  model: _Model,
  array: ArrayAtom,
  row: number,
  column: number
) {
  const range = cellRange(model, array, row, column);
  if (typeof range !== 'number') model.setSelection(range);
}

function setPositionInCell(
  model: _Model,
  array: ArrayAtom,
  row: number,
  column: number,
  pos: 'start' | 'end'
) {
  const cell = array.getCell(row, column);
  if (!cell) return;
  model.setPositionHandlingPlaceholder(
    model.offsetOf(cell[pos === 'start' ? 0 : cell.length - 1])
  );
}

/**
 * Internal primitive to add a column/row in an array.
 * Insert an array if necessary.
 */
function addCell(
  model: _Model,
  where: 'after row' | 'before row' | 'after column' | 'before column'
): void {
  const [arrayAtom, [row, column]] = parentArray(model, where);
  if (!arrayAtom) return;

  switch (where) {
    case 'after row':
      arrayAtom.addRowAfter(row);
      setPositionInCell(model, arrayAtom, row + 1, 0, 'end');
      break;

    case 'after column':
      if (arrayAtom.maxColumns <= arrayAtom.colCount) {
        model.announce('plonk');
        return;
      }
      arrayAtom.addColumnAfter(column);
      setPositionInCell(model, arrayAtom, row, column + 1, 'end');
      break;

    case 'before row':
      arrayAtom.addRowBefore(row);
      setPositionInCell(model, arrayAtom, row, 0, 'start');
      break;

    case 'before column':
      if (arrayAtom.maxColumns <= arrayAtom.colCount) {
        model.announce('plonk');
        return;
      }
      arrayAtom.addColumnBefore(column);
      setPositionInCell(model, arrayAtom, row, column, 'start');
      break;
  }
}

export function addRowAfter(model: _Model): boolean {
  // Only add a row if the current position is in the top level of a cell
  // or at the top level of the root (in which case we'll convert to a `lines` environment)
  const cursor = model.at(model.position);
  if (
    !isCellBranch(cursor.parentBranch) &&
    cursor.parent !== model.root &&
    model.root.type !== 'root'
  ) {
    model.announce('plonk');
    return false;
  }

  if (!model.contentWillChange({ inputType: 'insertText' })) return false;

  // If in multiline mode, split the current line
  if (model.parentEnvironment?.isMultiline) {
    // If there's a selection, delete it
    if (!model.selectionIsCollapsed) model.deleteAtoms(range(model.selection));

    // Get content before/after the cursor
    const [first, last] = cellSiblings(model)!;
    const after = model.extractAtoms([model.position, model.offsetOf(last)]);
    const before = model.extractAtoms([model.offsetOf(first), model.position]);

    const array = first.parent as ArrayAtom;
    const [row, col] = first.parentBranch as [number, number];

    array.setCell(row, col, before);

    // Add row after the current row
    addCell(model, 'after row');

    // Set row following cell
    array.setCell(row + 1, col, after);

    model.position = model.offsetOf(array.getCell(row + 1, col)![0]);

    model.contentDidChange({ inputType: 'insertText' });
    return true;
  }

  addCell(model, 'after row');
  model.contentDidChange({ inputType: 'insertText' });
  return true;
}

export function addRowBefore(model: _Model): boolean {
  if (!model.contentWillChange({ inputType: 'insertText' })) return false;
  addCell(model, 'before row');
  model.contentDidChange({ inputType: 'insertText' });
  return true;
}

export function addColumnAfter(model: _Model): boolean {
  if (!model.contentWillChange({ inputType: 'insertText' })) return false;
  addCell(model, 'after column');
  model.contentDidChange({ inputType: 'insertText' });
  return true;
}

export function addColumnBefore(model: _Model): boolean {
  if (!model.contentWillChange({ inputType: 'insertText' })) return false;
  addCell(model, 'before column');
  model.contentDidChange({ inputType: 'insertText' });
  return true;
}

export function setEnvironment(
  model: _Model,
  environment: TabularEnvironment
): boolean {
  if (!model.contentWillChange({})) return false;
  model.mathfield.snapshot();
  let leftDelim = '.';
  let rightDelim = '.';
  switch (environment) {
    case 'pmatrix':
    case 'pmatrix*':
      leftDelim = '(';
      rightDelim = ')';
      break;

    case 'bmatrix':
    case 'bmatrix*':
      leftDelim = '[';
      rightDelim = ']';
      break;

    case 'Bmatrix':
    case 'Bmatrix*':
      leftDelim = '\\lbrace';
      rightDelim = '\\rbrace';
      break;

    case 'vmatrix':
    case 'vmatrix*':
      leftDelim = '\\vert';
      rightDelim = '\\vert';
      break;

    case 'Vmatrix':
    case 'Vmatrix*':
      leftDelim = '\\Vert';
      rightDelim = '\\Vert';
      break;

    case 'matrix':
    case 'matrix*':
      // Specifying a fence, even a null fence,
      // will prevent the insertion of an initial and final gap
      leftDelim = '.';
      rightDelim = '.';
      break;
    case 'cases':
    case 'dcases':
      leftDelim = '\\lbrace';
      break;
    case 'rcases':
      rightDelim = '\\rbrace';
      break;
  }

  const atom = model.at(model.position);
  const arrayAtom =
    atom.type === 'array' ? (atom as ArrayAtom) : model.parentEnvironment!;
  arrayAtom.environmentName = environment;
  arrayAtom.leftDelim = leftDelim;
  arrayAtom.rightDelim = rightDelim;
  model.contentDidChange({});
  return true;
}

/**
 * Internal primitive to remove a column/row in a matrix
 */
function removeCell(model: _Model, where: 'row' | 'column'): void {
  // This command is only applicable if we're in an ArrayAtom
  let atom = model.at(model.position);

  while (
    atom &&
    !(Array.isArray(atom.parentBranch) && atom.parent instanceof ArrayAtom)
  )
    atom = atom.parent!;

  if (Array.isArray(atom?.parentBranch) && atom?.parent instanceof ArrayAtom) {
    const arrayAtom = atom.parent;
    const treeBranch = atom.parentBranch;
    let pos: number | undefined;
    switch (where) {
      case 'row':
        if (arrayAtom.rowCount > 1) {
          arrayAtom.removeRow(treeBranch[0]);
          const cell = arrayAtom.getCell(
            Math.max(0, treeBranch[0] - 1),
            treeBranch[1]
          )!;
          pos = model.offsetOf(cell[cell.length - 1]);
        }
        break;

      case 'column':
        if (arrayAtom.colCount > arrayAtom.minColumns) {
          arrayAtom.removeColumn(treeBranch[1]);
          const cell = arrayAtom.getCell(
            treeBranch[0],
            Math.max(0, treeBranch[1] - 1)
          )!;
          pos = model.offsetOf(cell[cell.length - 1]);
        }
        break;
    }
    if (pos !== undefined) model.setPositionHandlingPlaceholder(pos);
  }
}

export function removeRow(model: _Model): boolean {
  if (!model.contentWillChange({ inputType: 'deleteContent' })) return false;
  removeCell(model, 'row');
  model.contentDidChange({ inputType: 'deleteContent' });
  return true;
}

export function removeColumn(model: _Model): boolean {
  if (!model.contentWillChange({ inputType: 'deleteContent' })) return false;
  removeCell(model, 'column');
  model.contentDidChange({ inputType: 'deleteContent' });
  return true;
}

registerCommand(
  {
    addRowAfter,
    addColumnAfter,
    addRowBefore,
    addColumnBefore,
    removeRow,
    removeColumn,
    setEnvironment,
  },
  {
    target: 'model',
    canUndo: true,
    changeContent: true,
    changeSelection: true,
  }
);

function placeholderCell() {
  return [new PlaceholderAtom()];
}

// VisualTeX: edit the existing atom tree instead of serializing and reparsing it.
export function visualTexDefaultItalic(atom: Atom, mathfield: _Mathfield): boolean {
  const shape = ((atom.style as VisualTexStyleRecord).letterShapeStyle ||
    mathfield.letterShapeStyle ||
    "tex") as keyof typeof LETTER_SHAPE_MODIFIER;
  return LETTER_SHAPE_RANGES.some((range, i) =>
    range.test(atom.value || "") && LETTER_SHAPE_MODIFIER[shape][i] === "it");
}
export function visualTexAtomFont(atom: Atom, mathfield: _Mathfield): VisualTexFont {
  const style = atom.style;
  if (atom.mode === "text") return {
    bold: style.fontSeries === "b" || (style.fontSeries as string) === "bx",
    italic: style.fontShape === "it" || style.fontShape === "sl"
  };
  const variant = style.variant || "normal";
  return {
    bold: (style.variantStyle || "").includes("bold"),
    italic: (style.variantStyle || "").includes("italic") ||
      (!style.variantStyle && variant === "normal" && visualTexDefaultItalic(atom, mathfield))
  };
}
export const VISUALTEX_PERSISTENT_STYLE_KEYS = [
  "variant",
  "variantStyle",
  "fontSeries",
  "fontShape",
  "color",
  "backgroundColor",
  "verbatimColor",
  "verbatimBackgroundColor"
];
export function visualTexPersistentTypingGlyphs(atoms: readonly Atom[] | null | undefined): Atom[] {
  const result: Atom[] = [];
  const seen = new Set<Atom>();
  const visit = (atom: any): void => {
    if (!atom || typeof atom !== "object" || seen.has(atom)) return;
    seen.add(atom);
    const children = Array.isArray(atom.children) ? atom.children : [];
    if (
      children.length === 0 &&
      atom.mode !== "latex" &&
      atom.type !== "first" &&
      atom.type !== "placeholder" &&
      atom.type !== "prompt" &&
      typeof atom.value === "string" &&
      atom.value.length > 0
    ) {
      result.push(atom);
    }
    for (const child of children) visit(child);
  };
  for (const atom of atoms || []) visit(atom);
  return result;
}
export function visualTexPersistentStyleSnapshot(atom: Atom): VisualTexStyleRecord {
  const result: VisualTexStyleRecord = {};
  const style = (atom.style || {}) as VisualTexStyleRecord;
  for (const key of VISUALTEX_PERSISTENT_STYLE_KEYS) {
    if (Object.prototype.hasOwnProperty.call(style, key))
      result[key] = style[key];
  }
  return result;
}
export function visualTexPersistentStyleValue(snapshot: VisualTexStyleRecord, key: string): unknown {
  return Object.prototype.hasOwnProperty.call(snapshot, key)
    ? snapshot[key]
    : undefined;
}
export function visualTexPersistentStyleSource(mathfield: _Mathfield): Atom | null | undefined {
  const model = mathfield.model;
  const bias = mathfield.styleBias;
  if (!model || bias === "none") return null;
  const atom = model.at(model.position);
  if (!atom) return null;
  const before = ungroup(model, atom, bias);
  const after = ungroup(model, atom.rightSibling, bias);
  const offset = bias === "right" ? after : before;
  if (!Number.isFinite(offset) || offset < 0) return null;
  return model.at(offset);
}
export function visualTexRestoreInheritedPersistentStyle(mathfield: _Mathfield, atoms: readonly Atom[]): void {
  const source = visualTexPersistentStyleSource(mathfield);
  const marker = source && source.visualTexPersistentTypingApplied;
  if (!marker || !Array.isArray(marker.keys) || marker.keys.length === 0) return;
  for (const atom of visualTexPersistentTypingGlyphs(atoms)) {
    const style: VisualTexStyleRecord = { ...atom.style };
    let changed = false;
    for (const key of marker.keys) {
      const inherited = visualTexPersistentStyleValue(marker.after, key);
      const current = Object.prototype.hasOwnProperty.call(style, key)
        ? style[key]
        : undefined;
      if (current !== inherited) continue;
      const baseline = visualTexPersistentStyleValue(marker.before, key);
      if (baseline === undefined) delete style[key];
      else style[key] = baseline;
      changed = true;
    }
    if (changed) {
      atom.style = style as PrivateStyle;
      atom.isDirty = true;
    }
  }
}
export function visualTexApplyPersistentFontStyle(
  atom: Atom,
  mathfield: _Mathfield,
  persistent: VisualTexPersistentTypingStyle
): void {
  if (!persistent.bold && persistent.italic === null) return;
  const font = visualTexAtomFont(atom, mathfield);
  if (persistent.bold) font.bold = true;
  if (persistent.italic !== null) font.italic = persistent.italic;
  const style = { ...atom.style };
  if (atom.mode === "text") {
    if (persistent.bold) style.fontSeries = font.bold ? "b" : "m";
    if (persistent.italic !== null) style.fontShape = font.italic ? "it" : "n";
  } else if (atom.mode === "math") {
    const variant = style.variant || "normal";
    if (!["normal", "main", "ams"].includes(variant)) {
      style.variantStyle = font.bold ? (font.italic ? "bolditalic" : "bold") :
        (font.italic ? "italic" : undefined);
    } else {
      style.variant = font.italic ? "main" : "normal";
      style.variantStyle = font.bold ? (font.italic ? "bolditalic" : "bold") :
        (font.italic ? "italic" : "up");
      if (!font.bold && font.italic === visualTexDefaultItalic(atom, mathfield)) {
        style.variant = "normal";
        style.variantStyle = undefined;
      }
    }
  }
  atom.style = style;
  atom.isDirty = true;
}
export function visualTexApplyPersistentTypingStyle(mathfield: _Mathfield, atoms: readonly Atom[]): void {
  // Native MathLive inherits style from the glyph next to the caret. First
  // remove only the style delta that VisualTeX itself added to that source
  // glyph; this prevents a disabled persistent style from leaking forward
  // without breaking normal \mathbf/\mathrm/color inheritance.
  visualTexRestoreInheritedPersistentStyle(mathfield, atoms);

  const persistent = mathfield.visualTexPersistentTypingStyle;
  if (!persistent || (
    !persistent.bold &&
    persistent.italic === null &&
    !persistent.color &&
    !persistent.backgroundColor
  )) return;

  for (const atom of visualTexPersistentTypingGlyphs(atoms)) {
    const before = visualTexPersistentStyleSnapshot(atom);
    visualTexApplyPersistentFontStyle(atom, mathfield, persistent);
    const style = { ...atom.style };
    if (persistent.color) style.color = persistent.color;
    if (persistent.backgroundColor) style.backgroundColor = persistent.backgroundColor;
    atom.style = style;
    atom.isDirty = true;

    const after = visualTexPersistentStyleSnapshot(atom);
    const changedKeys = VISUALTEX_PERSISTENT_STYLE_KEYS.filter(
      (key) =>
        visualTexPersistentStyleValue(before, key) !==
        visualTexPersistentStyleValue(after, key),
    );
    if (changedKeys.length > 0) {
      atom.visualTexPersistentTypingApplied = {
        before,
        after,
        keys: changedKeys
      };
    } else {
      delete atom.visualTexPersistentTypingApplied;
    }
  }
}
export function visualTexToggleSelectionStyle(
  model: _Model,
  kind: string,
  requestedSelection?: Selection
): boolean {
  const selection = requestedSelection ? model.normalizeSelection(requestedSelection) : model.selection;
  if (kind !== "bold" && kind !== "italic" || !selection ||
      selection.ranges.every(([start, end]) => start === end)) return false;
  // Public setSelection() collapses disjoint ranges to their enclosing range.
  // Accept the caller's explicit ranges without formatting the gaps between them.
  const atoms: Atom[] = [...new Set<Atom>(model.getAtoms(selection, { includeChildren: true }))];
  // Containers (accent, fraction, root...) are structural, not font-bearing glyphs.
  const glyphs = atoms.filter(atom => atom.type !== "first" && atom.value &&
    atom.type !== "placeholder" && atom.type !== "prompt");
  if (!glyphs.length) return false;
  const inputType = kind === "bold" ? "formatBold" : "formatItalic";
  if (!model.contentWillChange({ inputType })) return false;
  const mathfield = model.mathfield;
  mathfield.flushInlineShortcutBuffer();
  mathfield.stopCoalescingUndo();
  if (mathfield.undoManager.index < 0) mathfield.snapshot();
  const fonts = glyphs.map(atom => visualTexAtomFont(atom, mathfield));
  const enabled = !fonts.every(font => font[kind]);
  model.deferNotifications({ content: true, type: inputType }, () => {
    // Parsed font commands copy style onto their descendants. Clear selected
    // container styles too, otherwise their old context overrides a glyph reset.
    for (const atom of atoms) {
      if (glyphs.includes(atom) || atom.type === "first") continue;
      atom.style = { ...atom.style, variant: undefined, variantStyle: undefined };
      atom.isDirty = true;
    }
    glyphs.forEach((atom, i) => {
      const font = { ...fonts[i], [kind]: enabled };
      const style = { ...atom.style };
      if (atom.mode === "text") {
        style.fontSeries = font.bold ? "b" : "m";
        style.fontShape = font.italic ? "it" : "n";
      } else {
        const variant = style.variant || "normal";
        if (!["normal", "main", "ams"].includes(variant)) {
          // Preserve a selected alphabet (mathbb, mathfrak, etc.).
          style.variantStyle = font.bold ? (font.italic ? "bolditalic" : "bold") :
            (font.italic ? "italic" : undefined);
        } else {
          style.variant = font.italic ? "main" : "normal";
          style.variantStyle = font.bold ? (font.italic ? "bolditalic" : "bold") :
            (font.italic ? "italic" : "up");
          if (!font.bold && font.italic === visualTexDefaultItalic(atom, mathfield)) {
            style.variant = "normal";
            style.variantStyle = undefined;
          }
        }
      }
      atom.style = style;
      atom.isDirty = true;
    });
    mathfield.snapshot();
  });
  return true;
}
export function visualTexPaintDetachedScripts(model: _Model, requestedSelection?: Selection): boolean {
  const selection = requestedSelection ? model.normalizeSelection(requestedSelection) : model.selection;
  if (!selection || selection.ranges.every(([start, end]) => start === end)) return false;
  const atoms: Atom[] = [...new Set<Atom>(model.getAtoms(selection, { includeChildren: true }))];
  // A keyboard selection can cover only the ^/_ attachment of c^2. MathLive
  // then paints that glyph-less container, which serializes as
  // \textcolor{..}{^2}: invisible, and dropped by the next LaTeX round trip.
  // Paint the script content instead, giving c^{\textcolor{..}{2}}.
  const detached = atoms.filter(atom => atom.type === "subsup" &&
    !atoms.includes(atom.leftSibling) &&
    (atom.style?.color !== undefined || atom.style?.backgroundColor !== undefined));
  if (!detached.length) return false;
  model.deferNotifications({ content: true, type: "formatFontColor" }, () => {
    for (const script of detached) {
      const paint: VisualTexStyleRecord = {};
      for (const key of ["color", "backgroundColor"] as const) {
        if (script.style[key] !== undefined) paint[key] = script.style[key];
      }
      for (const child of script.children) {
        if (child.type === "first") continue;
        child.style = { ...child.style, ...paint };
        child.isDirty = true;
      }
      script.style = { ...script.style, color: undefined, backgroundColor: undefined };
      script.isDirty = true;
    }
  });
  return true;
}
export function visualTexAlignedCellIsEmpty(cell: readonly Atom[] | undefined): boolean {
  return !cell || cell.every(atom => atom.type === "first");
}
export function visualTexEnvironmentCellIsEmpty(cell: readonly Atom[] | undefined): boolean {
  return !cell || cell.every(atom =>
    atom.type === "first" || atom.type === "placeholder" || atom.type === "prompt");
}
export function visualTexEnvironmentCellBranch(model: _Model, array: Atom | undefined): [number, number] | null {
  if (!(array instanceof ArrayAtom)) return null;
  const cursor = model.at(model.position);
  const candidates = [cursor, cursor?.rightSibling, cursor?.leftSibling];
  for (const candidate of candidates) {
    if (candidate?.parent === array && Array.isArray(candidate.parentBranch)) {
      return candidate.parentBranch;
    }
  }
  return null;
}
export function visualTexEmptyEnvironmentRowInfo(
  model: _Model
): { array: ArrayAtom; row: number; column: number } | null {
  const array = model.parentEnvironment;
  const branch = visualTexEnvironmentCellBranch(model, array);
  if (!Array.isArray(branch) ||
      !visualTexCanDeleteEmptyEnvironmentRow(array, branch[0])) return null;
  return { array, row: branch[0], column: branch[1] };
}
export function visualTexCanDeleteEmptyEnvironmentRow(array: Atom | undefined, row: number): array is ArrayAtom {
  if (!(array instanceof ArrayAtom) ||
      !["cases", "dcases", "rcases", "aligned", "gathered"].includes(array.environmentName) ||
      array.rowCount <= 1 ||
      row < 0 ||
      row >= array.rowCount) return false;
  for (let column = 0; column < array.colCount; column += 1) {
    if (!visualTexEnvironmentCellIsEmpty(array.getCell(row, column))) return false;
  }
  return true;
}
export function visualTexCollapseSingleVisualTexMultilineRow(model: _Model, array: Atom): boolean {
  if (!(array instanceof ArrayAtom) ||
      !["aligned", "gathered"].includes(array.environmentName) ||
      array.rowCount !== 1) return false;

  const parent = array.parent;
  const parentBranch = array.parentBranch;
  if (!parent || parentBranch == null) return false;

  const flattened: Atom[] = [];
  const columnCount = array.colCount;
  for (let column = 0; column < columnCount; column += 1) {
    const cell = array.getCell(0, column);
    if (!cell) continue;
    flattened.push(...cell.filter(atom => atom.type !== "first"));
  }

  if (Array.isArray(parentBranch)) {
    if (!(parent instanceof ArrayAtom)) return false;
    parent.setCell(parentBranch[0], parentBranch[1], flattened);
  } else {
    const siblings = parent.branch(parentBranch);
    if (!siblings) return false;
    const arrayIndex = siblings.indexOf(array);
    if (arrayIndex < 0) return false;
    const replacement = [
      ...siblings.slice(0, arrayIndex),
      ...flattened,
      ...siblings.slice(arrayIndex + 1),
    ];
    parent.setChildren(replacement, parentBranch);
  }

  array.parent = undefined;
  array.parentBranch = undefined;
  if (flattened.length > 0) {
    model.position = model.offsetOf(flattened[flattened.length - 1]);
  } else {
    model.position = Math.max(0, Math.min(model.position, model.lastOffset));
  }
  return true;
}
export function visualTexDeleteEmptyEnvironmentRowAt(
  model: _Model,
  array: Atom | undefined,
  row: number,
  column: number,
  direction: 'forward' | 'backward'
): boolean {
  if (!visualTexCanDeleteEmptyEnvironmentRow(array, row)) return false;
  const shouldCollapse =
    ["aligned", "gathered"].includes(array.environmentName) &&
    array.rowCount === 2;
  array.removeRow(row);
  if (shouldCollapse && visualTexCollapseSingleVisualTexMultilineRow(model, array)) {
    return true;
  }
  const targetRow = direction === "backward"
    ? Math.max(0, row - 1)
    : Math.min(row, array.rowCount - 1);
  const targetColumn = Math.max(0, Math.min(column || 0, array.colCount - 1));
  setPositionInCell(
    model,
    array,
    targetRow,
    targetColumn,
    direction === "backward" ? "end" : "start",
  );
  return true;
}
export function visualTexProvisionalAlignRows(array: ArrayAtom): Set<number> {
  if (!(array.visualTexProvisionalAlignRows instanceof Set)) {
    array.visualTexProvisionalAlignRows = new Set();
  }
  return array.visualTexProvisionalAlignRows;
}
export function visualTexEnsureAlignedColumnCapacity(array: ArrayAtom, requiredColumns: number): void {
  while (array.maxColumns < requiredColumns) {
    const trailing = array.colFormat[array.colFormat.length - 1];
    const hasTrailingGap = trailing && Object.prototype.hasOwnProperty.call(trailing, "gap");
    if (hasTrailingGap) (array.colFormat as ColumnFormat[]).pop();
    (array.colFormat as ColumnFormat[]).push(
      { gap: 1 },
      { align: "r" },
      { gap: 0.25 },
      { align: "l" },
      hasTrailingGap ? trailing : { gap: 0 },
    );
    array.isDirty = true;
  }
}
export function visualTexInsertAlignmentPoint(model: _Model): boolean {
  const array = model.parentEnvironment;
  if (!array || array.type !== "array" ||
      !["aligned", "align", "align*"].includes(array.environmentName)) return false;
  if (!model.contentWillChange({ inputType: "insertText" })) return false;

  const mathfield = model.mathfield;
  mathfield.flushInlineShortcutBuffer();
  mathfield.stopCoalescingUndo();
  if (mathfield.undoManager.index < 0) mathfield.snapshot();

  model.deferNotifications({ content: true, selection: true, type: "insertText" }, () => {
    if (!model.selectionIsCollapsed) model.deleteAtoms(range(model.selection));
    const cursor = model.at(model.position);
    if (cursor.parent !== array || !Array.isArray(cursor.parentBranch)) return;
    const [row, col] = cursor.parentBranch;
    const siblings = array.getCell(row, col);
    if (!siblings) return;
    const split = siblings.indexOf(cursor) + 1;
    const before = siblings.slice(1, split);
    const after = siblings.slice(split);
    const provisionalRows = visualTexProvisionalAlignRows(array);

    // A freshly-created VisualTeX align row starts in MathLive's left-aligned
    // second cell so typing grows naturally from left to right. The first '&'
    // commits that provisional row to standard TeX align semantics by moving
    // the content before the caret into the right-aligned first cell.
    if (
      col === 1 &&
      visualTexAlignedCellIsEmpty(array.getCell(row, 0)) &&
      (
        provisionalRows.has(row) ||
        (array.environmentName === "aligned" && array.colCount === 2)
      )
    ) {
      array.setCell(row, 0, before);
      array.setCell(row, 1, after);
      provisionalRows.delete(row);
      model.position = model.offsetOf(array.getCell(row, 1)![0]);
      mathfield.snapshot();
      return;
    }

    const targetCol = col + 1;
    const targetCell = array.getCell(row, targetCol);

    array.setCell(row, col, before);
    if (!targetCell || !visualTexAlignedCellIsEmpty(targetCell)) {
      visualTexEnsureAlignedColumnCapacity(array, array.colCount + 1);
      array.addColumnAfter(col);
    }
    array.setCell(row, targetCol, after);
    model.position = model.offsetOf(array.getCell(row, targetCol)![0]);
    mathfield.snapshot();
  });
  return true;
}
// Ctrl+J creates base_{#?}^{#?}. Auto-exit after the first script must enter
// the still-empty sibling script instead of leaving the whole structure.
export function visualTexMoveToEmptySiblingScript(model: _Model): boolean {
  let atom: Atom | undefined = model.at(model.position);
  while (atom && atom.parentBranch !== "subscript" && atom.parentBranch !== "superscript") {
    atom = atom.parent;
  }
  if (!atom?.parent) return false;
  const sibling = atom.parentBranch === "subscript" ? "superscript" : "subscript";
  const content = (atom.parent.branch(sibling) ?? []).filter(item => item.type !== "first");
  if (content.length !== 1 || content[0].type !== "placeholder") return false;
  const offset = model.offsetOf(content[0]);
  model.setSelection(offset - 1, offset);
  return true;
}
export function visualTexIsAlignmentMarker(atom: Atom | undefined): boolean {
  return atom?.command === "\\class" && atom.args?.[0] === "visualtex-align-marker";
}
// A top-level '&' typed before the first Shift+Enter is a zero-width marker.
// When the row becomes an aligned row, that marker is its column split.
export function visualTexSplitAtAlignmentMarker(atoms: readonly Atom[]): [Atom[], Atom[]] | null {
  const index = atoms.findIndex(visualTexIsAlignmentMarker);
  return index < 0 ? null : [atoms.slice(0, index), atoms.slice(index + 1)];
}
export function visualTexInsertRowBreak(model: _Model, environment = "gathered"): boolean {
  if (environment !== "aligned" && environment !== "gathered") return false;
  if (!model.contentWillChange({ inputType: "insertLineBreak" })) return false;
  const mathfield = model.mathfield;
  mathfield.flushInlineShortcutBuffer();
  mathfield.stopCoalescingUndo();
  if (mathfield.undoManager.index < 0) mathfield.snapshot();
  model.deferNotifications({ content: true, selection: true, type: "insertLineBreak" }, () => {
    if (!model.selectionIsCollapsed) model.deleteAtoms(range(model.selection));
    let cursor = model.at(model.position);
    const body = model.root.body?.filter(atom => atom.type !== "first");
    // A click outside a standalone array is a root position, not a cell position.
    if (body?.length === 1 && body[0].type === "array" &&
        (model.position === 0 || model.position === model.lastOffset)) {
      const array = body[0] as ArrayAtom;
      const atStart = model.position === 0;
      setPositionInCell(model, array, atStart ? 0 : array.rowCount - 1,
        atStart ? 0 : array.colCount - 1, atStart ? "start" : "end");
      cursor = model.at(model.position);
    }
    // Split the current branch. At a cell boundary this extends the array;
    // inside a fraction/accent it creates local lines without flattening it.
    // (The parent is an array when the branch is a cell.)
    const parent = cursor.parent as ArrayAtom | undefined;
    const branch = cursor.parentBranch;
    if (!parent || branch == null) return;
    if (Array.isArray(branch)) {
      const [row, col] = branch;
      const siblings = parent.getCell(row, col)!;
      const split = siblings.indexOf(cursor) + 1;
      const before = siblings.slice(1, split);
      const after = siblings.slice(split);
      parent.setCell(row, col, before);
      parent.addRowAfter(row);
      if (parent.environmentName === "aligned") {
        for (let column = 0; column < parent.colCount; column += 1) {
          parent.setCell(row + 1, column, []);
        }
        const afterCells = visualTexSplitAtAlignmentMarker(after);
        if (afterCells) {
          parent.setCell(row + 1, 0, afterCells[0]);
          parent.setCell(row + 1, 1, afterCells[1]);
          model.position = model.offsetOf(parent.getCell(row + 1, 0)![0]);
        } else {
          // Put a newly-created align row into the left-aligned second cell.
          // This keeps normal typing direction intuitive until the user enters
          // an explicit '&', which visualTexInsertAlignmentPoint then commits
          // into the standard right/left TeX column pair.
          visualTexProvisionalAlignRows(parent).add(row + 1);
          parent.setCell(row + 1, 1, after);
          model.position = model.offsetOf(parent.getCell(row + 1, 1)![0]);
        }
      } else {
        parent.setCell(row + 1, col, after);
        model.position = model.offsetOf(parent.getCell(row + 1, col)![0]);
      }
    } else {
      const siblings = parent.branch(branch)!;
      const split = siblings.indexOf(cursor) + 1;
      const before = siblings.slice(1, split);
      const after = siblings.slice(split);
      parent.removeBranch(branch);
      const beforeCells = environment === "aligned" ? visualTexSplitAtAlignmentMarker(before) : null;
      const afterCells = environment === "aligned" ? visualTexSplitAtAlignmentMarker(after) : null;
      const array =
        environment === "aligned"
          ? makeEnvironment(environment, [beforeCells ?? [[], before], afterCells ?? [[], after]])
          : makeEnvironment(environment, [[before], [after]]);
      if (environment === "aligned") {
        const provisionalRows = visualTexProvisionalAlignRows(array);
        if (!beforeCells) provisionalRows.add(0);
        if (!afterCells) provisionalRows.add(1);
      }
      parent.setChildren([array], branch);
      model.position = model.offsetOf(
        array.getCell(1, environment === "aligned" && !afterCells ? 1 : 0)![0],
      );
    }
    mathfield.snapshot();
  });
  return true;
}

registerCommand(
  {
    visualTexToggleSelectionStyle,
    visualTexPaintDetachedScripts,
    visualTexInsertAlignmentPoint,
    visualTexInsertRowBreak,
  },
  { target: 'model', canUndo: true, changeContent: true, changeSelection: true }
);

registerCommand(
  {
    visualTexMoveToEmptySiblingScript,
  },
  { target: 'model', changeSelection: true }
);
