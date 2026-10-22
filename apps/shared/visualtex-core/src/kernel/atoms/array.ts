import type {
  MathstyleName,
  Dimension,
  Environment,
  ParseMode,
} from '../public/core-types';

import { Atom, isCellBranch, isNamedBranch } from '../core/atom-class';
import { joinLatex } from '../core/tokenizer';

import { PlaceholderAtom } from './placeholder';
import type { AtomJson, Branch, ToLatexOptions } from '../core/types';

export type ColumnFormat =
  | {
      // A regular content column, with the specified alignment.
      // 'm' is a special alignment for multline: left on first row, right on last
      // row, centered otherwise
      align: 'l' | 'c' | 'r' | 'm';
    }
  | {
      // The width of a gap between columns, or a LaTeX expression between columns
      gap: number | readonly Atom[];
    }
  | {
      // A rule (line) separating columns
      separator: 'solid' | 'dashed';
    };

export type ColSeparationType =
  // | 'align'
  // | 'alignat'
  // | 'gather'
  | 'small'
  // | 'CD'
  | undefined;

export type ArrayAtomConstructorOptions = {
  isRoot?: boolean;

  mathstyleName?: MathstyleName;

  colSeparationType?: ColSeparationType;
  leftDelim?: string;
  rightDelim?: string;
  // A multiplication factor applied to the spacing between rows and columns
  arraystretch?: number;
  // The spacing between columns
  arraycolsep?: number;

  columns?: ColumnFormat[];
  minColumns?: number;
  maxColumns?: number;
  minRows?: number;
  maxRows?: number;

  displayEquationNumber?: boolean;

  classes?: string[];
};

/**
 * Normalize cells:
 * - ensure the array of cells is dense (not sparse)
 * - fold rows that overflow (longer than maximum number of columns)
 * - ensure each cell begins with a `first` atom
 * - ensure the minimum/maximum number of cells is respected
 * - remove last row if empty (TeX behavior)
 */

function normalizeCells(
  atom: ArrayAtom,
  cells: (readonly Atom[])[][],
  options: {
    columns: readonly ColumnFormat[];
    minColumns: number;
    minRows: number;
    maxRows: number;
  }
): (readonly Atom[])[][] {
  //
  // 1/
  // - Fold the array so that there are no more columns of content than
  // there are columns prescribed by the column format.
  // - Fill rows that have fewer cells than expected with empty cells
  // - Ensure that all the cells have a `first` atom.
  //

  // The number of column is determined by the colFormat
  let maxColCount = 0;
  for (const colSpec of options.columns)
    if ('align' in colSpec) maxColCount += 1;
  maxColCount = Math.max(maxColCount, options.minColumns);

  // Actual number of columns (at most `maxColCount`)
  let colCount = 0;
  const rows: (readonly Atom[])[][] = [];

  for (const row of cells) {
    colCount = Math.max(colCount, Math.min(row.length, maxColCount));
    let colIndex = 0;
    while (colIndex < row.length) {
      const newRow: (readonly Atom[])[] = [];
      const lastCol = Math.min(row.length, colIndex + maxColCount);
      while (colIndex < lastCol) {
        newRow.push(normalizeCell(row[colIndex], atom.mode));
        colIndex += 1;
      }

      if (newRow.length < options.minColumns) {
        while (newRow.length < options.minColumns)
          newRow.push([new Atom({ type: 'first', mode: atom.mode })]);
      }

      rows.push(newRow);
    }
  }

  //
  // 2/ If the last row is empty, ignore it (TeX behavior)
  // (unless there's only one row)
  //
  if (
    !atom.isMultiline &&
    rows.length > 0 &&
    rows[rows.length - 1].length === 1 &&
    isEmptyCell(rows[rows.length - 1][0])
  )
    rows.pop();

  //
  // 3/ Fill out any missing cells
  //
  const result: (readonly Atom[])[][] = [];
  for (const row of rows) {
    if (row.length !== colCount) {
      for (let i = row.length; i < colCount; i++) {
        if (atom.isMultiline)
          row.push([new Atom({ type: 'first', mode: atom.mode })]);
        else {
          row.push([
            new Atom({ type: 'first', mode: atom.mode }),
            new PlaceholderAtom(),
          ]);
        }
      }
    }
    result.push(row);
  }

  //
  // 4/ Set the `parent` and `treeBranch` for each cell
  //
  let rowIndex = 0;
  let colIndex = 0;
  for (const row of result) {
    colIndex = 0;
    for (const cell of row) {
      for (const element of cell) {
        element.parent = atom;
        element.parentBranch = [rowIndex, colIndex];
      }
      colIndex += 1;
    }
    rowIndex += 1;
  }

  atom.isDirty = true;

  return result;
}

function normalizeCell(
  cell: readonly Atom[],
  mode: ParseMode
): readonly Atom[] {
  if (cell.length === 0) return [new Atom({ type: 'first', mode })];
  if (cell[0].type !== 'first')
    return [new Atom({ type: 'first', mode }), ...cell];
  console.assert(!cell.slice(1).some((x) => x.type === 'first'));
  return cell;
}

function isEmptyCell(cell: readonly Atom[]): boolean {
  return cell.length === 1 && cell[0].type === 'first';
}

// See http://ctan.math.utah.edu/ctan/tex-archive/macros/latex/base/lttab.dtx
export class ArrayAtom extends Atom {
  environmentName: Environment;

  /** True if the environment is multiline. In a multiline environment,
   * pressing return automatically creates a new row. In non-multiline
   * environments, pressing alt+return creates a new row.
   */
  get isMultiline(): boolean {
    const env = this.environmentName;
    return [
      'lines',
      'multline',
      'multline*',
      'align',
      'split',
      'gather',
      'gathered',
    ].includes(env);
  }

  // The array is a 2D array of cells, each cell being an array of atoms
  private _rows: (undefined | readonly Atom[])[][];

  rowGaps: readonly Dimension[];
  arraystretch?: number;
  arraycolsep?: number;
  colSeparationType?: ColSeparationType;
  leftDelim?: string;
  rightDelim?: string;
  mathstyleName?: MathstyleName;

  colFormat: readonly ColumnFormat[];
  /** VisualTeX: aligned rows still typing into their second cell (see visualTexInsertAlignmentPoint) */
  visualTexProvisionalAlignRows?: Set<number>;
  minColumns: number;
  minRows: number;
  maxRows: number;

  // Additional classes when rendering the array
  classes: string[];

  constructor(
    envName: Environment,
    array: (readonly Atom[])[][],
    rowGaps: readonly Dimension[],
    options: Readonly<ArrayAtomConstructorOptions> = {}
  ) {
    super({ type: 'array', isRoot: options.isRoot });
    this.environmentName = envName;

    if (options.columns) {
      if (options.columns.length === 0) this.colFormat = [{ align: 'l' }];
      else this.colFormat = [...options.columns];
    } else {
      if (options.minColumns) {
        // Repeat {align: 'l'} as many times as minColumns
        const columns: ColumnFormat[] = [];
        for (let i = 0; i < options.minColumns; i++)
          columns.push({ align: 'l' });
        this.colFormat = columns;
      } else {
        // In TeX definition, if no columns are specified, a maximum of
        // 10 left-aligned columns is assumed

        this.colFormat = [
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
        ];
      }
    }

    this.minColumns = options.minColumns ?? 1;
    this.minRows = options.minRows ?? 1;
    this.maxRows = options.maxRows ?? Infinity;

    this._rows = normalizeCells(this, array, {
      columns: this.colFormat,
      minColumns: this.minColumns,
      minRows: this.minRows,
      maxRows: this.maxRows,
    });

    this.rowGaps = rowGaps;

    if (options.arraycolsep !== undefined)
      this.arraycolsep = options.arraycolsep;
    this.colSeparationType = options.colSeparationType;
    // Default \arraystretch from lttab.dtx
    if (options.arraystretch !== undefined)
      this.arraystretch = options.arraystretch;

    if (options.mathstyleName) this.mathstyleName = options.mathstyleName;
    if (options.leftDelim) this.leftDelim = options.leftDelim;
    if (options.rightDelim) this.rightDelim = options.rightDelim;

    this.classes = options.classes ?? [];
  }

  static fromJson(json: AtomJson): ArrayAtom {
    return new ArrayAtom(
      json.environmentName,
      json.array,
      json.rowGaps,
      json as any
    );
  }

  toJson(): AtomJson {
    const result: AtomJson = {
      ...super.toJson(),
      environmentName: this.environmentName,
      array: this._rows.map((row) =>
        row.map((col) => col!.map((x) => x.toJson()))
      ),
      rowGaps: this.rowGaps,
      columns: this.colFormat,
      colSeparationType: this.colSeparationType,
      minColumns: this.minColumns,
      minRows: this.minRows,
      maxRows: this.maxRows,
    };

    if (this.arraystretch !== undefined)
      result.arraystretch = this.arraystretch;
    if (this.arraycolsep !== undefined) result.arraycolsep = this.arraycolsep;
    if (this.leftDelim) result.leftDelim = this.leftDelim;
    if (this.rightDelim) result.rightDelim = this.rightDelim;
    if (this.isRoot) result.isRoot = true;
    result.minColumns = this.minColumns;
    result.minRows = this.minRows;
    result.maxRows = this.maxRows;
    if (this.mathstyleName) result.mathstyleName = this.mathstyleName;
    if (this.classes.length > 0) result.classes = this.classes;

    return result;
  }

  branch(cell: Branch): readonly Atom[] | undefined {
    if (!isCellBranch(cell)) return undefined;
    return this._rows[cell[0]][cell[1]] ?? undefined;
  }

  createBranch(cell: Branch): Atom[] {
    if (!isCellBranch(cell)) return [];
    this.isDirty = true;
    return (this.branch(cell) as Atom[]) ?? [];
  }

  get rowCount(): number {
    return this._rows.length;
  }

  get colCount(): number {
    return this._rows[0].length;
  }

  get maxColumns(): number {
    return this.colFormat.filter((col) => Boolean((col as { align?: string })['align'])).length;
  }

  removeBranch(name: Branch): readonly Atom[] {
    if (isNamedBranch(name)) return super.removeBranch(name);

    const [_first, ...children] = this.branch(name)!;
    // Drop the 'first' element
    console.assert(_first.type === 'first');

    this._rows[name[0]][name[1]] = undefined;
    children.forEach((x) => {
      x.parent = undefined;
      x.parentBranch = undefined;
    });

    this.isDirty = true;
    return children;
  }

  get hasChildren(): boolean {
    return this.children.length > 0;
  }

  _serialize(options: ToLatexOptions): string {
    const result: string[] = [];

    // Only wrap in \displaylines{} if there are multiple rows
    if (this.environmentName === 'lines' && this._rows.length > 1)
      result.push(`\\displaylines{`);
    else if (this.environmentName !== 'lines')
      result.push(`\\begin{${this.environmentName}}`);

    if (this.environmentName === 'array') {
      result.push('{');
      if (this.colFormat !== undefined) {
        for (const format of this.colFormat) {
          if ('align' in format && typeof format.align === 'string')
            result.push(format.align);
          else if ('separator' in format && format.separator === 'solid')
            result.push('|');
          else if ('separator' in format && format.separator === 'dashed')
            result.push(':');
        }
      }

      result.push('}');
    }

    for (let row = 0; row < this._rows.length; row++) {
      for (let col = 0; col < this._rows[row].length; col++) {
        if (col > 0) result.push(' & ');
        result.push(Atom.serialize(this._rows[row][col], options));
      }

      // Adds a separator between rows (but not after the last row)
      if (row < this._rows.length - 1) {
        const gap = this.rowGaps[row];
        if (gap?.dimension)
          result.push(`\\\\[${gap.dimension} ${gap.unit ?? 'pt'}] `);
        else result.push('\\\\ ');
      }
    }

    // Only close \displaylines{} if there are multiple rows
    if (this.environmentName === 'lines' && this._rows.length > 1)
      result.push(`}`);
    else if (this.environmentName !== 'lines')
      result.push(`\\end{${this.environmentName}}`);

    return joinLatex(result);
  }

  forEachCell(
    callback: (cell: readonly Atom[], row: number, col: number) => void
  ): void {
    for (let i = 0; i < this.rowCount; i++)
      for (let j = 0; j < this.colCount; j++) callback(this._rows[i][j]!, i, j);
  }

  getCell(row: number, col: number): readonly Atom[] | undefined {
    return this._rows[row]?.[col];
  }

  setCell(row: number, column: number, value: readonly Atom[]): void {
    console.assert(
      this.type === 'array' &&
        Array.isArray(this._rows) &&
        this._rows[row][column] !== undefined
    );
    for (const atom of this._rows[row][column]!) {
      atom.parent = undefined;
      atom.parentBranch = undefined;
    }

    const atoms = [
      new Atom({ type: 'first', mode: this.mode }),
      ...value.filter((x) => x.type !== 'first'),
    ];

    this._rows[row][column] = atoms;

    adjustBranches(this);

    this.isDirty = true;
  }

  addRowBefore(row: number): void {
    console.assert(this.type === 'array' && Array.isArray(this._rows));

    this._rows.splice(
      row,
      0,
      Array.from({ length: this.colCount }, () =>
        makeEmptyCell(this, !this.isMultiline)
      )
    );
    adjustBranches(this);
    this.isDirty = true;
  }

  addRowAfter(row: number): void {
    console.assert(this.type === 'array' && Array.isArray(this._rows));

    this._rows.splice(
      row + 1,
      0,
      Array.from({ length: this.colCount }, () =>
        makeEmptyCell(this, !this.isMultiline)
      )
    );

    adjustBranches(this);
    this.isDirty = true;
  }

  removeRow(row: number): void {
    console.assert(
      this.type === 'array' && Array.isArray(this._rows) && this.rowCount > row
    );

    const deleted = this._rows.splice(row, 1);
    for (const column of deleted) {
      for (const cell of column) {
        if (cell) {
          for (const child of cell) {
            child.parent = undefined;
            child.parentBranch = undefined;
          }
        }
      }
    }

    adjustBranches(this);
    this.isDirty = true;
  }

  addColumnBefore(col: number): void {
    console.assert(this.type === 'array' && Array.isArray(this._rows));
    for (const row of this._rows) row.splice(col, 0, makeEmptyCell(this));

    adjustBranches(this);
    this.isDirty = true;
  }

  addColumnAfter(col: number): void {
    console.assert(this.type === 'array' && Array.isArray(this._rows));
    for (const row of this._rows) row.splice(col + 1, 0, makeEmptyCell(this));

    adjustBranches(this);
    this.isDirty = true;
  }

  addColumn(): void {
    this.addColumnAfter(this.colCount - 1);
  }

  removeColumn(col: number): void {
    console.assert(
      this.type === 'array' && Array.isArray(this._rows) && this.colCount > col
    );
    for (const row of this._rows) {
      const deleted = row.splice(col, 1);
      for (const cell of deleted) {
        if (cell) {
          for (const child of cell) {
            child.parent = undefined;
            child.parentBranch = undefined;
          }
        }
      }
    }
    adjustBranches(this);
    this.isDirty = true;
  }

  get cells(): (readonly Atom[])[] {
    const result: (readonly Atom[])[] = [];
    for (const row of this._rows) {
      for (const cell of row)
        if (cell) result.push(cell.filter((x) => x.type !== 'first'));
    }
    return result;
  }

  get rows(): (undefined | readonly Atom[])[][] {
    return this._rows!;
  }
}

/**
 * Create an empty cell
 */
function makeEmptyCell(
  parent: ArrayAtom,
  withPlaceholder = false
): readonly Atom[] {
  const first = new Atom({ type: 'first', mode: parent.mode });
  first.parent = parent;
  const result = [first];
  if (withPlaceholder) {
    const placeholder = new PlaceholderAtom();
    placeholder.parent = parent;
    result.push(placeholder);
  }
  return result;
}

function adjustBranches(array: ArrayAtom): void {
  for (let i = 0; i < array.rowCount; i++) {
    for (let j = 0; j < array.colCount; j++) {
      const atoms = array.getCell(i, j);
      if (atoms) {
        for (const atom of atoms) {
          if (atom) {
            atom.parent = array;
            atom.parentBranch = [i, j];
          }
        }
      }
    }
  }
}
