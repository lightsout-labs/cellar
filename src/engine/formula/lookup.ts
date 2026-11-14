import { isError, type Value } from '../value';
import { wildcards } from './conditional';
import {
  columnName,
  compare,
  error,
  fail,
  NA,
  quoteSheetName,
  Reference,
  typed,
  VALUE,
  type Args,
  type Spec,
  type Typed,
} from './values';

const REF = error('#REF!');

/** A reference argument, or the error that it gives instead. */
function reference(read: Typed, index: number): Reference {
  const found = read.args.reference(index);
  return isError(found) ? fail(found) : found;
}

/** The value looked up: the first argument, which must not be an error. */
function sought(read: Typed): Value {
  const value = read.args.value(0);
  return isError(value) ? fail(value) : value;
}

/** A whole number argument, `fallback` when not given or left out. */
const option = (read: Typed, index: number, fallback: number) =>
  read.args.given(index) ? read.integer(index) : fallback;

/** Whether a reference is one row or one column. */
const isVector = (ref: Reference) => ref.rows === 1 || ref.cols === 1;

/**
 * The values of one row (`across`) or column of a reference, counted from
 * its first, with empty cells as `null`, up to the last that is not empty,
 * so that a whole column costs only its cells.
 */
function line(
  args: Args,
  ref: Reference,
  index: number,
  across: boolean,
): Value[] {
  const { area } = ref;
  const part = across
    ? { ...area, top: area.top + index, bottom: area.top + index }
    : { ...area, left: area.left + index, right: area.left + index };
  const found: [at: number, value: Value][] = [];
  const missing = args.cells(new Reference(ref.sheet, part), (value, r, c) => {
    found.push([across ? c - area.left : r - area.top, value]);
  });
  if (missing) return fail(missing);
  let length = 0;
  for (const [at] of found) length = Math.max(length, at + 1);
  const values = new Array<Value>(length).fill(null);
  for (const [at, value] of found) values[at] = value;
  return values;
}

/** The order of a cell against a value, or `null` when it is of another type. */
function ordering(value: Value): (cell: Value) => number | null {
  return (cell) =>
    cell === null || isError(cell) || typeof cell !== typeof value
      ? null
      : compare(cell, value);
}

/** A test of a cell for being the value, with wildcards in text if `wild`. */
function equals(value: Value, wild: boolean): (cell: Value) => boolean {
  if (wild && typeof value === 'string' && /[*?~]/.test(value)) {
    const matches = wildcards(value);
    return (cell) => typeof cell === 'string' && matches(cell);
  }
  const order = ordering(value);
  return (cell) => order(cell) === 0;
}

/**
 * The last place of a value or the largest below it in sorted cells, as
 * Excel's approximate lookups find it; or, `descending`, the smallest
 * above it. Cells of other types are passed over. -1 when none.
 */
function sorted(value: Value, cells: Value[], descending = false): number {
  const order = ordering(value);
  let found = -1;
  for (let at = 0; at < cells.length; at++) {
    const o = order(cells[at]!);
    if (o === null) continue;
    if (descending ? o < 0 : o > 0) break;
    found = at;
  }
  return found;
}

/**
 * The place of a value as `XMATCH` finds it: `mode` 0 exactly, -1 or the
 * next smaller, 1 or the next larger, 2 with wildcards; `search` from the
 * last when negative. Sorted searches find what a scan finds. -1 when none.
 */
function nearest(
  value: Value,
  cells: Value[],
  mode: number,
  search: number,
): number {
  if (![-1, 0, 1, 2].includes(mode) || ![-2, -1, 1, 2].includes(search)) {
    return fail(VALUE);
  }
  const same = equals(value, mode === 2);
  const order = ordering(value);
  let best = -1;
  for (let step = 0; step < cells.length; step++) {
    const at = search < 0 ? cells.length - 1 - step : step;
    const cell = cells[at]!;
    if (same(cell)) return at;
    const o = order(cell);
    if (o === null || mode === 0 || mode === 2 || o * mode < 0) continue;
    if (best < 0 || compare(cell, cells[best]!) * mode < 0) best = at;
  }
  return best;
}

/** `VLOOKUP`, or `HLOOKUP` when `across`. */
function tableLookup(across: boolean): Spec {
  return typed(3, 4, (read) => {
    const value = sought(read);
    const table = reference(read, 1);
    const index = read.integer(2);
    if (index < 1) return VALUE;
    if (index > (across ? table.rows : table.cols)) return REF;
    const keys = line(read.args, table, 0, across);
    const at = read.boolean(3, true)
      ? sorted(value, keys)
      : keys.findIndex(equals(value, true));
    if (at < 0) return NA;
    const { top, left } = table.area;
    const found = across
      ? read.args.cell(table, top + index - 1, left + at)
      : read.args.cell(table, top + at, left + index - 1);
    return found ?? 0;
  });
}

/** One row or column of a reference, a cell of a vector, by its place. */
function part(ref: Reference, at: number, across: boolean): Reference {
  const { area } = ref;
  return new Reference(
    ref.sheet,
    across
      ? { ...area, left: area.left + at, right: area.left + at }
      : { ...area, top: area.top + at, bottom: area.top + at },
  );
}

/** `ROW` or `COLUMN`: the first of a reference, or the formula's own. */
function place(row: boolean): Spec {
  return {
    min: 0,
    max: 1,
    range: (index) => index === 0,
    call(args) {
      if (args.length === 0) return (row ? args.row : args.col) + 1;
      const ref = args.reference(0);
      if (isError(ref)) return ref;
      return (row ? ref.area.top : ref.area.left) + 1;
    },
  };
}

/** `ROWS` or `COLUMNS`: of a reference, or 1 for one value. */
function size(rows: boolean): Spec {
  return {
    min: 1,
    max: 1,
    call(args) {
      const result = args.result(0);
      if (!(result instanceof Reference)) return isError(result) ? result : 1;
      return rows ? result.rows : result.cols;
    },
  };
}

export const LOOKUP_FUNCTIONS: Record<string, Spec> = {
  ADDRESS: typed(2, 5, (read) => {
    const row = read.integer(0);
    const col = read.integer(1);
    const absolute = option(read, 2, 1);
    const a1 = read.boolean(3, true);
    if (row < 1 || col < 1 || row > 1_048_576 || col > 16_384) return VALUE;
    if (absolute < 1 || absolute > 4) return VALUE;
    const fixedRow = absolute <= 2;
    const fixedCol = absolute % 2 === 1;
    const sheet =
      read.args.length > 4 ? `${quoteSheetName(read.text(4))}!` : '';
    if (a1) {
      const letters = columnName(col - 1);
      return `${sheet}${fixedCol ? '$' : ''}${letters}${fixedRow ? '$' : ''}${row}`;
    }
    const r = fixedRow ? `R${row}` : `R[${row}]`;
    const c = fixedCol ? `C${col}` : `C[${col}]`;
    return sheet + r + c;
  }),
  CHOOSE: typed(2, 255, (read) => {
    const index = read.integer(0);
    if (index < 1 || index >= read.args.length) return VALUE;
    return read.args.result(index);
  }),
  COLUMN: place(false),
  COLUMNS: size(false),
  HLOOKUP: tableLookup(true),
  INDEX: typed(2, 4, (read) => {
    const result = read.args.result(0);
    if (isError(result)) return result;
    let row = read.integer(1);
    let col = read.integer(2, 0);
    if (read.integer(3, 1) !== 1) return REF;
    if (row < 0 || col < 0) return VALUE;
    if (!(result instanceof Reference))
      return row > 1 || col > 1 ? REF : result;
    if (read.args.length === 2 && result.rows === 1) [row, col] = [0, row];
    if (row > result.rows || col > result.cols) return REF;
    const { area } = result;
    return new Reference(result.sheet, {
      top: row ? area.top + row - 1 : area.top,
      bottom: row ? area.top + row - 1 : area.bottom,
      left: col ? area.left + col - 1 : area.left,
      right: col ? area.left + col - 1 : area.right,
    });
  }),
  LOOKUP: typed(2, 3, (read) => {
    const value = sought(read);
    const ref = reference(read, 1);
    if (read.args.length === 2) {
      // The array form: along the longer side, from the last row or column.
      const across = ref.cols > ref.rows;
      const at = sorted(value, line(read.args, ref, 0, across));
      if (at < 0) return NA;
      const last = across ? ref.rows - 1 : ref.cols - 1;
      return line(read.args, ref, last, across)[at] ?? 0;
    }
    const results = reference(read, 2);
    const at = sorted(value, line(read.args, ref, 0, ref.rows === 1));
    if (at < 0) return NA;
    const across = results.rows === 1;
    const { top, left } = results.area;
    const found = read.args.cell(
      results,
      across ? top : top + at,
      across ? left + at : left,
    );
    return found ?? 0;
  }),
  MATCH: typed(2, 3, (read) => {
    const value = sought(read);
    const ref = reference(read, 1);
    const type = read.number(2, 1);
    if (!isVector(ref)) return NA;
    const cells = line(read.args, ref, 0, ref.rows === 1);
    const at =
      type === 0
        ? cells.findIndex(equals(value, true))
        : sorted(value, cells, type < 0);
    return at < 0 ? NA : at + 1;
  }),
  ROW: place(true),
  ROWS: size(true),
  VLOOKUP: tableLookup(false),
  XLOOKUP: typed(
    3,
    6,
    (read) => {
      const value = sought(read);
      const keys = reference(read, 1);
      const results = reference(read, 2);
      if (!isVector(keys)) return VALUE;
      const across = keys.rows === 1 && keys.cols > 1;
      if (across ? results.cols !== keys.cols : results.rows !== keys.rows) {
        return VALUE;
      }
      const cells = line(read.args, keys, 0, across);
      const at = nearest(value, cells, option(read, 4, 0), option(read, 5, 1));
      if (at >= 0) return part(results, at, across);
      return read.args.given(3) ? read.args.result(3) : NA;
    },
    true,
  ),
  XMATCH: typed(
    2,
    4,
    (read) => {
      const value = sought(read);
      const ref = reference(read, 1);
      if (!isVector(ref)) return VALUE;
      const cells = line(read.args, ref, 0, ref.rows === 1);
      const at = nearest(value, cells, option(read, 2, 0), option(read, 3, 1));
      return at < 0 ? NA : at + 1;
    },
    true,
  ),
};
