import { editText, formatValue, GENERAL, parseInput } from '../format';
import { columnName, evaluate, parse, type Ast } from '../formula';
import { VALUE_ERROR, type Value } from '../value';

export type { Value } from '../value';
export { columnName } from '../formula';

/** An edit to the workbook. Rows and columns count from zero. */
export type Command =
  | {
      readonly type: 'setCell';
      readonly sheet: string;
      readonly row: number;
      readonly col: number;
      /**
       * What the user typed: a number, a date, a time, a percentage, a
       * boolean, text, or a formula starting with `=`.
       */
      readonly input: string;
      /**
       * The year of a date typed without one. When absent, the current
       * year, which the workbook adds to the command its listeners get.
       */
      readonly year?: number;
    }
  | {
      /** Empties the cells of ranges, keeping their formats. */
      readonly type: 'clear';
      readonly sheet: string;
      readonly ranges: readonly Range[];
    }
  | {
      /**
       * Makes ranges hold exactly the cells listed, with their contents and
       * styles, and nothing else: how the workbook undoes and redoes.
       */
      readonly type: 'restore';
      readonly sheet: string;
      readonly ranges: readonly Range[];
      readonly cells: readonly CellContents[];
    };

/** A cell's contents and style, as a `restore` command puts them back. */
export interface CellContents {
  readonly row: number;
  readonly col: number;
  readonly value: Constant;
  /** The formula with its leading `=`, in place of the value. */
  readonly formula?: string;
  /** Absent for the default style. */
  readonly style?: Style;
}

/** The cells that an undo or a redo changed. */
export interface Change {
  readonly sheet: string;
  readonly ranges: readonly Range[];
}

/** How many changes can be undone (R5). */
const UNDO_LIMIT = 100;

/** The number of rows and columns of every sheet (R1). */
export const SHEET_ROWS = 100_000;
export const SHEET_COLUMNS = 100;

/** A rectangle of cells, its bounds included. */
export interface Range {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

/** What a set of cells adds up to, as the status bar shows it. */
export interface Aggregate {
  /** The cells that are not empty. */
  readonly count: number;
  /** The cells that hold a number, and their total. */
  readonly numbers: number;
  readonly sum: number;
}

/** What a cell holds and shows. */
export interface Cell {
  readonly value: Value;
  /** The formula with its leading `=`, or `null` for a plain value. */
  readonly formula: string | null;
  /** The value shown in the cell's format. */
  readonly display: string;
  /** The text that edits the cell, as the formula bar shows it. */
  readonly input: string;
}

/** How cells look. Cells that look the same share one. */
export interface Style {
  /** An Excel format code; `General` when absent. */
  readonly numberFormat?: string;
}

export interface SheetInfo {
  readonly id: string;
  readonly name: string;
}

/** A stored value: formulas keep their text, not their result. */
type Constant = number | string | boolean | null;

export interface SheetSnapshot extends SheetInfo {
  /** Values by column, then by row; `null` where a cell is empty. */
  readonly columns: Constant[][];
  /** Formulas as `[row, column, text]`, the text with its leading `=`. */
  readonly formulas: [number, number, string][];
  /**
   * Indexes in the workbook's styles by column, then by row; 0 where a cell
   * has the default style. Absent in snapshots saved before cells had
   * styles.
   */
  readonly styleIds?: number[][];
}

/**
 * The versioned, JSON-compatible form in which a workbook leaves the
 * engine, for storage and files.
 */
export interface Snapshot {
  readonly version: 1;
  readonly name: string;
  readonly sheets: SheetSnapshot[];
  /**
   * The distinct styles of the cells, the default one first. Absent in
   * snapshots saved before cells had styles.
   */
  readonly styles?: Style[];
}

interface Formula {
  readonly row: number;
  readonly col: number;
  readonly text: string;
  readonly ast: Ast;
}

interface Sheet extends SheetInfo {
  /** Values by column, then by row. Both levels are sparse. */
  readonly columns: Constant[][];
  readonly formulas: Map<string, Formula>;
  /** Indexes in the workbook's styles by column, then by row. Both sparse. */
  readonly styleIds: number[][];
  /** Results of the formulas read since the last change. */
  readonly results: Map<string, Value>;
}

/** A command that undoes or redoes another. */
type Restore = Extract<Command, { type: 'restore' }>;

const key = (row: number, col: number) => `${row},${col}`;

/** The name of a cell, as in `A1`. */
export function cellName(row: number, col: number): string {
  return columnName(col) + (row + 1);
}

/** A new workbook with one empty sheet. */
export function createWorkbook(): Workbook {
  return Workbook.fromSnapshot({
    version: 1,
    name: 'Book1',
    sheets: [{ id: 'sheet1', name: 'Sheet1', columns: [], formulas: [] }],
  });
}

export class Workbook {
  private readonly listeners = new Set<(command: Command) => void>();
  /** Formulas being calculated now, to detect ones that read each other. */
  private readonly calculating = new Set<Formula>();
  /** The index of each style in `styles`, by its JSON text. */
  private readonly styleIndex: Map<string, number>;
  /** The commands that undo the last changes, the latest last. */
  private readonly undoList: Restore[] = [];
  /** The commands that redo the changes undone, the latest undone last. */
  private readonly redoList: Restore[] = [];

  private constructor(
    private readonly name: string,
    private readonly sheetList: Sheet[],
    private readonly styles: Style[],
  ) {
    this.styleIndex = new Map(
      styles.map((style, index) => [JSON.stringify(style), index]),
    );
  }

  /**
   * The workbook of a snapshot, with `changes` made after it replayed.
   * There is nothing to undo in it.
   */
  static fromSnapshot(
    snapshot: Snapshot,
    changes: readonly Command[] = [],
  ): Workbook {
    if (snapshot.version !== 1) {
      throw new Error(`Unknown snapshot version: ${String(snapshot.version)}`);
    }
    const sheets = snapshot.sheets.map((sheet): Sheet => {
      const formulas = new Map<string, Formula>();
      for (const [row, col, text] of sheet.formulas) {
        const ast = parse(text.slice(1));
        if (!ast) throw new Error(`Invalid formula in snapshot: ${text}`);
        formulas.set(key(row, col), { row, col, text, ast });
      }
      return {
        id: sheet.id,
        name: sheet.name,
        columns: sheet.columns.map((column) => [...column]),
        formulas,
        styleIds: (sheet.styleIds ?? []).map((column) => [...column]),
        results: new Map(),
      };
    });
    const workbook = new Workbook(
      snapshot.name,
      sheets,
      snapshot.styles ?? [{}],
    );
    for (const command of changes) workbook.apply(command);
    return workbook;
  }

  toSnapshot(): Snapshot {
    return {
      version: 1,
      name: this.name,
      sheets: this.sheetList.map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        columns: Array.from(sheet.columns, (column) =>
          Array.from(column ?? [], (value) => value ?? null),
        ),
        formulas: Array.from(
          sheet.formulas.values(),
          ({ row, col, text }): [number, number, string] => [row, col, text],
        ),
        styleIds: Array.from(sheet.styleIds, (column) =>
          Array.from(column ?? [], (id) => id ?? 0),
        ),
      })),
      styles: [...this.styles],
    };
  }

  /** The sheets in tab order. */
  get sheets(): readonly SheetInfo[] {
    return this.sheetList.map(({ id, name }) => ({ id, name }));
  }

  /**
   * Makes a change, which can then be undone; what could be redone is
   * forgotten.
   */
  execute(command: Command): void {
    this.redoList.length = 0;
    this.apply(command, this.undoList);
  }

  /** Undoes the last change, and gives the cells it changed. */
  undo(): Change | null {
    return this.reverse(this.undoList, this.redoList);
  }

  /** Redoes the last change undone, and gives the cells it changed. */
  redo(): Change | null {
    return this.reverse(this.redoList, this.undoList);
  }

  get canUndo(): boolean {
    return this.undoList.length > 0;
  }

  get canRedo(): boolean {
    return this.redoList.length > 0;
  }

  cell(sheetId: string, row: number, col: number): Cell {
    const sheet = this.sheet(sheetId);
    const value = this.value(sheet, row, col);
    const formula =
      sheet.formulas.size === 0
        ? null
        : (sheet.formulas.get(key(row, col))?.text ?? null);
    const style = this.styles[sheet.styleIds[col]?.[row] ?? 0];
    const code = style?.numberFormat ?? GENERAL;
    return {
      value,
      formula,
      display: formatValue(value, code),
      // Only the formula bar and the editor need it, not every drawn cell.
      get input() {
        return formula ?? editText(value, code);
      },
    };
  }

  /**
   * The count, and the total of the numbers, of the cells of each range. A
   * cell in two ranges counts twice, as in `SUM`.
   */
  aggregate(sheetId: string, ranges: readonly Range[]): Aggregate {
    const sheet = this.sheet(sheetId);
    let count = 0;
    let numbers = 0;
    let sum = 0;
    const add = (value: Value) => {
      if (value === null) return;
      count++;
      if (typeof value === 'number') {
        numbers++;
        sum += value;
      }
    };
    for (const { top, left, bottom, right } of ranges) {
      // Formula cells hold null among the values, so they are added apart.
      const last = Math.min(right, sheet.columns.length - 1);
      for (let col = left; col <= last; col++) {
        const column = sheet.columns[col];
        if (!column) continue;
        const end = Math.min(bottom, column.length - 1);
        // The hot loop of a whole sheet, so `add` is written out here.
        for (let row = top; row <= end; row++) {
          const value = column[row];
          if (typeof value === 'number') {
            count++;
            numbers++;
            sum += value;
          } else if (value != null) {
            count++;
          }
        }
      }
      for (const { row, col } of sheet.formulas.values()) {
        if (row >= top && row <= bottom && col >= left && col <= right) {
          add(this.value(sheet, row, col));
        }
      }
    }
    return { count, numbers, sum };
  }

  /** How many rows and columns, from A1, the cells in use extend over. */
  extent(sheetId: string): { rows: number; cols: number } {
    const sheet = this.sheet(sheetId);
    let rows = 0;
    let cols = 0;
    sheet.columns.forEach((column, col) => {
      let row = column.length - 1;
      while (row >= 0 && column[row] == null) row--;
      if (row < 0) return;
      rows = Math.max(rows, row + 1);
      cols = col + 1;
    });
    for (const { row, col } of sheet.formulas.values()) {
      rows = Math.max(rows, row + 1);
      cols = Math.max(cols, col + 1);
    }
    return { rows, cols };
  }

  /**
   * Calls `listener` with each command after it is executed, until the
   * returned function is called.
   */
  subscribe(listener: (command: Command) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  /** Runs the last command of `from`, and keeps its reverse in `to`. */
  private reverse(from: Restore[], to: Restore[]): Change | null {
    const command = from.pop();
    if (!command) return null;
    this.apply(command, to);
    return { sheet: command.sheet, ranges: command.ranges };
  }

  /**
   * Runs a command, keeps the command that reverses it in `history`, and
   * then tells the listeners.
   */
  private apply(command: Command, history?: Restore[]): void {
    const sheet = this.sheet(command.sheet);
    const ranges =
      command.type === 'setCell'
        ? [
            {
              top: command.row,
              left: command.col,
              bottom: command.row,
              right: command.col,
            },
          ]
        : command.ranges;
    const reverse: Restore = {
      type: 'restore',
      sheet: sheet.id,
      ranges,
      cells: this.contents(sheet, ranges),
    };
    let executed = command;
    if (command.type === 'setCell') {
      const year = this.setCell(sheet, command);
      // Listeners that record the command replay it with the same date.
      if (year !== undefined) executed = { ...command, year };
    } else if (command.type === 'clear') {
      for (const range of command.ranges) clear(sheet, range);
    } else {
      this.restore(sheet, command);
    }
    // Any formula may read the changed cells, so every result is stale.
    for (const each of this.sheetList) each.results.clear();
    if (history) {
      history.push(reverse);
      if (history.length > UNDO_LIMIT) history.shift();
    }
    for (const listener of this.listeners) listener(executed);
  }

  /** The cells of ranges that have contents or a style other than the default. */
  private contents(sheet: Sheet, ranges: readonly Range[]): CellContents[] {
    const cells = new Map<string, CellContents>();
    for (const { top, left, bottom, right } of ranges) {
      const last = Math.min(
        right,
        Math.max(sheet.columns.length, sheet.styleIds.length) - 1,
      );
      for (let col = left; col <= last; col++) {
        const values = sheet.columns[col] ?? [];
        const ids = sheet.styleIds[col] ?? [];
        const end = Math.min(bottom, Math.max(values.length, ids.length) - 1);
        for (let row = top; row <= end; row++) {
          const value = values[row] ?? null;
          const id = ids[row] ?? 0;
          if (value === null && id === 0) continue;
          cells.set(key(row, col), {
            row,
            col,
            value,
            ...(id !== 0 && { style: this.styles[id] }),
          });
        }
      }
      for (const [cellKey, { row, col, text }] of sheet.formulas) {
        if (row < top || row > bottom || col < left || col > right) continue;
        const id = sheet.styleIds[col]?.[row] ?? 0;
        cells.set(cellKey, {
          row,
          col,
          value: null,
          formula: text,
          ...(id !== 0 && { style: this.styles[id] }),
        });
      }
    }
    return [...cells.values()];
  }

  /** Empties the ranges of a `restore`, styles too, and fills in its cells. */
  private restore(sheet: Sheet, { ranges, cells }: Restore): void {
    for (const range of ranges) {
      clear(sheet, range);
      const { top, left, bottom, right } = range;
      const last = Math.min(right, sheet.styleIds.length - 1);
      for (let col = left; col <= last; col++) {
        const ids = sheet.styleIds[col];
        if (!ids) continue;
        const end = Math.min(bottom, ids.length - 1);
        for (let row = top; row <= end; row++) ids[row] = 0;
      }
    }
    for (const { row, col, value, formula, style } of cells) {
      if (formula !== undefined) {
        const ast = parse(formula.slice(1));
        // A command read back from storage may be damaged.
        if (!ast) throw new Error(`Invalid formula in restore: ${formula}`);
        sheet.formulas.set(key(row, col), { row, col, text: formula, ast });
      }
      (sheet.columns[col] ??= [])[row] = value;
      if (style) (sheet.styleIds[col] ??= [])[row] = this.styleId(style);
    }
  }

  /** Sets a cell, and gives the year it gave a date typed without one. */
  private setCell(
    sheet: Sheet,
    command: Extract<Command, { type: 'setCell' }>,
  ): number | undefined {
    const { row, col, input } = command;
    const ast = input.startsWith('=') ? parse(input.slice(1)) : null;
    const column = (sheet.columns[col] ??= []);
    if (ast) {
      sheet.formulas.set(key(row, col), { row, col, text: input, ast });
      column[row] = null;
      return undefined;
    }
    sheet.formulas.delete(key(row, col));
    const { value, format, year } = parseInput(input, command.year);
    column[row] = value;
    if (format === undefined) return year;
    // The way the value was typed sets the cell's format, as in Excel.
    const ids = (sheet.styleIds[col] ??= []);
    const style = { ...this.styles[ids[row] ?? 0], numberFormat: format };
    ids[row] = this.styleId(style);
    return year;
  }

  /** The index of a style in the table, added if it is new. */
  private styleId(style: Style): number {
    const text = JSON.stringify(style);
    let id = this.styleIndex.get(text);
    if (id === undefined) {
      id = this.styles.push(style) - 1;
      this.styleIndex.set(text, id);
    }
    return id;
  }

  private sheet(id: string): Sheet {
    const sheet = this.sheetList.find((each) => each.id === id);
    if (!sheet) throw new Error(`No such sheet: ${id}`);
    return sheet;
  }

  private value(sheet: Sheet, row: number, col: number): Value {
    // Most sheets read are plain values; their key is never built.
    if (sheet.formulas.size === 0) return sheet.columns[col]?.[row] ?? null;
    const cellKey = key(row, col);
    const formula = sheet.formulas.get(cellKey);
    if (!formula) return sheet.columns[col]?.[row] ?? null;
    const known = sheet.results.get(cellKey);
    if (known !== undefined) return known;
    if (this.calculating.has(formula)) return VALUE_ERROR;
    this.calculating.add(formula);
    const result = evaluate(formula.ast, {
      cell: (r, c) => this.value(sheet, r, c),
    });
    this.calculating.delete(formula);
    sheet.results.set(cellKey, result);
    return result;
  }
}

/** Empties the cells of a range of a sheet, leaving their styles. */
function clear(sheet: Sheet, { top, left, bottom, right }: Range): void {
  const last = Math.min(right, sheet.columns.length - 1);
  for (let col = left; col <= last; col++) {
    const column = sheet.columns[col];
    if (!column) continue;
    const end = Math.min(bottom, column.length - 1);
    for (let row = top; row <= end; row++) column[row] = null;
  }
  for (const [cellKey, { row, col }] of sheet.formulas) {
    if (row >= top && row <= bottom && col >= left && col <= right) {
      sheet.formulas.delete(cellKey);
    }
  }
}
