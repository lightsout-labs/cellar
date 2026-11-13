import { formatValue, parseInput } from '../format';
import { columnName, evaluate, parse, type Ast } from '../formula';
import { VALUE_ERROR, type Value } from '../value';

export type { Value } from '../value';
export { columnName } from '../formula';

/** An edit to the workbook. Rows and columns count from zero. */
export type Command = {
  readonly type: 'setCell';
  readonly sheet: string;
  readonly row: number;
  readonly col: number;
  /** What the user typed: a number, text, or a formula starting with `=`. */
  readonly input: string;
};

/** What a cell holds and shows. */
export interface Cell {
  readonly value: Value;
  /** The formula with its leading `=`, or `null` for a plain value. */
  readonly formula: string | null;
  readonly display: string;
}

export interface SheetInfo {
  readonly id: string;
  readonly name: string;
}

/** A stored value: formulas keep their text, not their result. */
type Constant = number | string | null;

export interface SheetSnapshot extends SheetInfo {
  /** Values by column, then by row; `null` where a cell is empty. */
  readonly columns: Constant[][];
  /** Formulas as `[row, column, text]`, the text with its leading `=`. */
  readonly formulas: [number, number, string][];
}

/**
 * The versioned, JSON-compatible form in which a workbook leaves the
 * engine, for storage and files.
 */
export interface Snapshot {
  readonly version: 1;
  readonly name: string;
  readonly sheets: SheetSnapshot[];
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
  /** Results of the formulas read since the last change. */
  readonly results: Map<string, Value>;
}

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
  private readonly listeners = new Set<() => void>();
  /** Formulas being calculated now, to detect ones that read each other. */
  private readonly calculating = new Set<Formula>();

  private constructor(
    private readonly name: string,
    private readonly sheetList: Sheet[],
  ) {}

  static fromSnapshot(snapshot: Snapshot): Workbook {
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
        results: new Map(),
      };
    });
    return new Workbook(snapshot.name, sheets);
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
      })),
    };
  }

  /** The sheets in tab order. */
  get sheets(): readonly SheetInfo[] {
    return this.sheetList.map(({ id, name }) => ({ id, name }));
  }

  execute(command: Command): void {
    const sheet = this.sheet(command.sheet);
    const { row, col, input } = command;
    const ast = input.startsWith('=') ? parse(input.slice(1)) : null;
    const column = (sheet.columns[col] ??= []);
    if (ast) {
      sheet.formulas.set(key(row, col), { row, col, text: input, ast });
      column[row] = null;
    } else {
      sheet.formulas.delete(key(row, col));
      column[row] = parseInput(input);
    }
    // Any formula may read the changed cell, so every result is stale.
    for (const each of this.sheetList) each.results.clear();
    for (const listener of this.listeners) listener();
  }

  cell(sheetId: string, row: number, col: number): Cell {
    const sheet = this.sheet(sheetId);
    const value = this.value(sheet, row, col);
    const formula = sheet.formulas.get(key(row, col))?.text ?? null;
    return { value, formula, display: formatValue(value) };
  }

  /** Calls `listener` after every change, until the returned function is called. */
  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private sheet(id: string): Sheet {
    const sheet = this.sheetList.find((each) => each.id === id);
    if (!sheet) throw new Error(`No such sheet: ${id}`);
    return sheet;
  }

  private value(sheet: Sheet, row: number, col: number): Value {
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
