import { evaluate, references, type Area, type Ast } from '../formula';
import type { ErrorValue, Value } from '../value';

/** A stored value: formulas keep their text, not their result. */
export type Constant = number | string | boolean | null;

export interface Formula {
  readonly sheet: Sheet;
  readonly row: number;
  readonly col: number;
  /** The text with its leading `=`. */
  readonly text: string;
  readonly ast: Ast;
  /** The result, or `undefined` while the formula needs calculating. */
  result?: Value;
  /** What the formula reads, as entered in the dependents of the sheets. */
  reads: Dependency[];
}

/** That a formula reads an area of a sheet. */
interface Dependency {
  readonly formula: Formula;
  readonly sheet: Sheet;
  readonly area: Area;
}

export interface Sheet {
  readonly id: string;
  name: string;
  color?: string;
  /** Values by column, then by row. Both levels are sparse. */
  readonly columns: Constant[][];
  readonly formulas: Map<string, Formula>;
  /** The same formulas by column, then by row. */
  readonly formulaColumns: Map<number, Map<number, Formula>>;
  /** Indexes in the workbook's styles by column, then by row. Both sparse. */
  readonly styleIds: number[][];
  /** The formulas that read one cell, by the cell's key. */
  readonly cellReaders: Map<string, Set<Dependency>>;
  /** The formulas that read an area of several cells, by each column of it. */
  readonly areaReaders: Map<number, Set<Dependency>>;
}

export const key = (row: number, col: number) => `${row},${col}`;

const REF_ERROR: ErrorValue = { error: '#REF!' };

/** An empty sheet with the name and contents given. */
export function emptySheet(
  id: string,
  name: string,
  columns: Constant[][],
  styleIds: number[][],
): Sheet {
  return {
    id,
    name,
    columns,
    formulas: new Map(),
    formulaColumns: new Map(),
    styleIds,
    cellReaders: new Map(),
    areaReaders: new Map(),
  };
}

/**
 * Puts a formula in a cell that has none, without entering what it reads:
 * for a new sheet, which `Calculation.rebuild` then enters.
 */
export function place(
  sheet: Sheet,
  row: number,
  col: number,
  text: string,
  ast: Ast,
): Formula {
  const formula: Formula = { sheet, row, col, text, ast, reads: [] };
  sheet.formulas.set(key(row, col), formula);
  let column = sheet.formulaColumns.get(col);
  if (!column) sheet.formulaColumns.set(col, (column = new Map()));
  column.set(row, formula);
  return formula;
}

/**
 * The formulas of a workbook and what they depend on. A change marks the
 * formulas that depend on it as needing calculation, and a formula is
 * calculated when it is read, after the formulas it reads, with an
 * explicit stack: a chain of 100,000 formulas does not overflow the call
 * stack (ADR 0004).
 */
export class Calculation {
  /** How many formula results have been calculated. */
  calculations = 0;
  /** Formulas found in circular references, in the order found. */
  private readonly circular = new Set<Formula>();
  /** Formulas set since the last `settle`. */
  private readonly fresh: Formula[] = [];

  constructor(private readonly sheets: () => readonly Sheet[]) {}

  /** Sets a formula in a cell, in place of what the cell had. */
  put(sheet: Sheet, row: number, col: number, text: string, ast: Ast): void {
    this.remove(sheet, row, col);
    const formula = place(sheet, row, col, text, ast);
    this.enter(formula);
    this.fresh.push(formula);
  }

  /** Removes the formula of a cell, if it has one. */
  remove(sheet: Sheet, row: number, col: number): void {
    const formula = sheet.formulas.get(key(row, col));
    if (!formula) return;
    sheet.formulas.delete(key(row, col));
    sheet.formulaColumns.get(col)!.delete(row);
    this.leave(formula);
    this.circular.delete(formula);
  }

  /**
   * Marks the formulas that read cells of the ranges, directly or through
   * other formulas, as needing calculation.
   */
  changed(sheet: Sheet, areas: readonly Area[]): void {
    const stack: Formula[] = [];
    for (const area of areas) this.readers(sheet, area, stack);
    while (stack.length > 0) {
      const formula = stack.pop()!;
      // A formula that needs calculating has readers that need it too.
      if (formula.result === undefined) continue;
      formula.result = undefined;
      const { sheet, row, col } = formula;
      this.readers(
        sheet,
        { top: row, left: col, bottom: row, right: col },
        stack,
      );
    }
  }

  /**
   * Finds again what every formula reads, after a change to the sheets,
   * and calculates them all.
   */
  rebuild(): void {
    const sheets = this.sheets();
    for (const sheet of sheets) {
      sheet.cellReaders.clear();
      sheet.areaReaders.clear();
    }
    this.circular.clear();
    for (const sheet of sheets) {
      for (const formula of sheet.formulas.values()) {
        formula.result = undefined;
        formula.reads = [];
        this.enter(formula);
        this.fresh.push(formula);
      }
    }
    this.settle();
  }

  /**
   * Calculates the formulas set since the last call and those that were
   * in circular references, so that every circular reference is known.
   */
  settle(): void {
    const pending = [...this.fresh];
    this.fresh.length = 0;
    for (const formula of this.circular) {
      if (formula.result !== undefined) continue;
      this.circular.delete(formula);
      pending.push(formula);
    }
    for (const formula of pending) {
      if (
        formula.sheet.formulas.get(key(formula.row, formula.col)) === formula
      ) {
        this.value(formula);
      }
    }
  }

  /** A formula in a circular reference, or `null` when there is none. */
  circularReference(): Formula | null {
    for (const formula of this.circular) return formula;
    return null;
  }

  /** The result of a formula, calculated if it needs to be. */
  value(target: Formula): Value {
    if (target.result !== undefined) return target.result;
    const path = new Set<Formula>([target]);
    const stack = [{ formula: target, pending: this.precedents(target) }];
    while (stack.length > 0) {
      const frame = stack[stack.length - 1]!;
      const { formula } = frame;
      if (formula.result !== undefined) {
        stack.pop();
        path.delete(formula);
        continue;
      }
      const next = frame.pending.pop();
      if (next === undefined) {
        formula.result = evaluate(formula.ast, this.context(formula));
        this.calculations++;
      } else if (next.result !== undefined) {
        continue;
      } else if (path.has(next)) {
        // The formulas from `next` up the stack read each other. Excel
        // gives them 0 and reports the reference.
        const first = stack.findIndex((each) => each.formula === next);
        for (const each of stack.slice(first)) {
          each.formula.result = 0;
          this.circular.add(each.formula);
        }
      } else {
        path.add(next);
        stack.push({ formula: next, pending: this.precedents(next) });
      }
    }
    return target.result!;
  }

  /** The value of a cell, calculated if it is a formula. */
  cell(sheet: Sheet, row: number, col: number): Value {
    const formula =
      sheet.formulas.size === 0 ? undefined : sheet.formulas.get(key(row, col));
    return formula ? this.value(formula) : (sheet.columns[col]?.[row] ?? null);
  }

  private context(formula: Formula) {
    const resolve = (name: string | undefined) =>
      name === undefined ? formula.sheet : this.find(name);
    return {
      row: formula.row,
      col: formula.col,
      cell: (name: string | undefined, row: number, col: number): Value => {
        const sheet = resolve(name);
        return sheet ? this.cell(sheet, row, col) : REF_ERROR;
      },
      cells: (
        name: string | undefined,
        area: Area,
        visit: (value: Value, row: number, col: number) => void,
      ) => {
        const sheet = resolve(name);
        if (!sheet) return REF_ERROR;
        this.cells(sheet, area, visit);
        return undefined;
      },
    };
  }

  /** Visits the cells of an area that are not empty. */
  private cells(
    sheet: Sheet,
    area: Area,
    visit: (value: Value, row: number, col: number) => void,
  ): void {
    const { top, left, bottom, right } = area;
    const last = Math.min(right, sheet.columns.length - 1);
    for (let col = left; col <= last; col++) {
      const column = sheet.columns[col];
      if (!column) continue;
      const end = Math.min(bottom, column.length - 1);
      for (let row = top; row <= end; row++) {
        const value = column[row];
        if (value != null) visit(value, row, col);
      }
    }
    this.formulasIn(sheet, area, (formula) =>
      visit(this.value(formula), formula.row, formula.col),
    );
  }

  private find(name: string): Sheet | undefined {
    const folded = name.toLowerCase();
    return this.sheets().find((sheet) => sheet.name.toLowerCase() === folded);
  }

  /** The formulas a formula reads that need calculating. */
  private precedents(formula: Formula): Formula[] {
    const found: Formula[] = [];
    for (const { sheet, area } of formula.reads) {
      this.formulasIn(sheet, area, (each) => {
        if (each.result === undefined) found.push(each);
      });
    }
    return found;
  }

  private formulasIn(
    sheet: Sheet,
    { top, left, bottom, right }: Area,
    visit: (formula: Formula) => void,
  ): void {
    if (top === bottom && left === right) {
      const formula = sheet.formulas.get(key(top, left));
      if (formula) visit(formula);
      return;
    }
    for (const [col, column] of sheet.formulaColumns) {
      if (col < left || col > right) continue;
      for (const [row, formula] of column) {
        if (row >= top && row <= bottom) visit(formula);
      }
    }
  }

  /** Adds the formulas that read cells of an area to `found`. */
  private readers(sheet: Sheet, area: Area, found: Formula[]): void {
    const { top, left, bottom, right } = area;
    const cells = (bottom - top + 1) * (right - left + 1);
    if (cells <= sheet.cellReaders.size) {
      for (let row = top; row <= bottom; row++) {
        for (let col = left; col <= right; col++) {
          const readers = sheet.cellReaders.get(key(row, col));
          if (readers) for (const each of readers) found.push(each.formula);
        }
      }
    } else {
      for (const readers of sheet.cellReaders.values()) {
        for (const each of readers) {
          if (overlap(each.area, area)) found.push(each.formula);
        }
      }
    }
    for (const [col, readers] of sheet.areaReaders) {
      if (col < left || col > right) continue;
      for (const each of readers) {
        if (each.area.top <= bottom && each.area.bottom >= top) {
          found.push(each.formula);
        }
      }
    }
  }

  /** Enters what a formula reads in the readers of the sheets. */
  private enter(formula: Formula): void {
    for (const { sheet: name, area } of references(formula.ast)) {
      const sheet = name === undefined ? formula.sheet : this.find(name);
      // A sheet that does not exist gives #REF!; a later one rebuilds.
      if (!sheet) continue;
      const dependency = { formula, sheet, area };
      formula.reads.push(dependency);
      if (area.top === area.bottom && area.left === area.right) {
        const cellKey = key(area.top, area.left);
        let readers = sheet.cellReaders.get(cellKey);
        if (!readers) sheet.cellReaders.set(cellKey, (readers = new Set()));
        readers.add(dependency);
      } else {
        for (let col = area.left; col <= area.right; col++) {
          let readers = sheet.areaReaders.get(col);
          if (!readers) sheet.areaReaders.set(col, (readers = new Set()));
          readers.add(dependency);
        }
      }
    }
  }

  private leave(formula: Formula): void {
    for (const dependency of formula.reads) {
      const { sheet, area } = dependency;
      if (area.top === area.bottom && area.left === area.right) {
        const cellKey = key(area.top, area.left);
        const readers = sheet.cellReaders.get(cellKey)!;
        readers.delete(dependency);
        if (readers.size === 0) sheet.cellReaders.delete(cellKey);
      } else {
        for (let col = area.left; col <= area.right; col++) {
          sheet.areaReaders.get(col)!.delete(dependency);
        }
      }
    }
    formula.reads = [];
  }
}

function overlap(a: Area, b: Area): boolean {
  return (
    a.top <= b.bottom &&
    a.bottom >= b.top &&
    a.left <= b.right &&
    a.right >= b.left
  );
}
