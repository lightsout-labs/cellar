import { expect } from 'vitest';
import {
  evaluate,
  parse,
  type Area,
  type Context,
} from '../src/engine/formula';
import type { Value } from '../src/engine/value';

/** The sheets of the test workbook, besides the formula's own. */
const SHEETS = ['Sheet2', 'My Sheet', "It's"];

/**
 * A context over cells named as in `A1` or `Sheet2!A1`, for a formula in
 * the cell `at`.
 */
export function context(cells: Record<string, Value>, at = 'Z100'): Context {
  const position = locate(at);
  const sheetOf = (name: string | undefined) =>
    name === undefined
      ? ''
      : SHEETS.find((each) => each.toLowerCase() === name.toLowerCase());
  const located = Object.entries(cells).map(([name, value]) => {
    const [sheet, cell] = name.includes('!') ? name.split('!') : ['', name];
    return { sheet: sheet!, ...locate(cell!), value };
  });
  return {
    ...position,
    cell(sheet, row, col) {
      const found = sheetOf(sheet);
      if (found === undefined) return { error: '#REF!' };
      const cell = located.find(
        (each) => each.sheet === found && each.row === row && each.col === col,
      );
      return cell?.value ?? null;
    },
    cells(sheet, area: Area, visit) {
      const found = sheetOf(sheet);
      if (found === undefined) return { error: '#REF!' };
      // In reverse, as the order of the cells is not set.
      for (const { sheet, row, col, value } of [...located].reverse()) {
        if (
          sheet === found &&
          value !== null &&
          row >= area.top &&
          row <= area.bottom &&
          col >= area.left &&
          col <= area.right
        ) {
          visit(value, row, col);
        }
      }
      return undefined;
    },
  };
}

function locate(name: string): { row: number; col: number } {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(name)!;
  let col = 0;
  for (const letter of letters!) col = col * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(digits) - 1, col: col - 1 };
}

/** The result of a formula, without its `=`, over the cells given. */
export function calculate(
  text: string,
  cells: Record<string, Value> = {},
  at?: string,
): Value {
  const ast = parse(text);
  if (!ast) throw new Error(`not a formula: ${text}`);
  return evaluate(ast, context(cells, at));
}

export const error = (code: string) => ({ error: code });

/**
 * Cells given as columns of values from a first cell down, as Excel's
 * documentation lays out its examples: `column('A2', [1, 2])`.
 */
export function column(
  first: string,
  values: readonly Value[],
): Record<string, Value> {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(first)!;
  return Object.fromEntries(
    values.map((value, index) => [
      `${letters}${Number(digits) + index}`,
      value,
    ]),
  );
}

/**
 * Checks a result against Excel's: to 9 decimals, or, when Excel's
 * documentation shows numbers rounded, a fraction to as many decimals as
 * it shows.
 */
export function expectExcel(
  actual: Value,
  expected: Value,
  rounded = false,
): void {
  if (typeof expected === 'number') {
    const shown = String(expected).split('.')[1]?.length ?? 0;
    const exact = !rounded || Number.isInteger(expected);
    expect(actual).toBeCloseTo(expected, exact ? 9 : shown);
  } else {
    expect(actual).toEqual(expected);
  }
}
