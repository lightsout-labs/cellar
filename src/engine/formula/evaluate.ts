import { isError, type ErrorValue, type Value } from '../value';
import type { Ast, BinaryOperator, Context } from './index';
import { LIBRARY } from './library';
import {
  compare,
  DIV0,
  error,
  NUM,
  power,
  Reference,
  toNumber,
  toText,
  VALUE,
  type Area,
  type Args,
  type Result,
} from './values';

/** The number of rows and columns of every sheet (R1). */
export const SHEET_ROWS = 100_000;
export const SHEET_COLUMNS = 100;

/**
 * Calculates the value of a parsed formula. An empty result, as that of a
 * reference to an empty cell, is 0, as Excel shows it.
 */
export function evaluate(ast: Ast, context: Context): Value {
  return scalar(run(ast, context), context) ?? 0;
}

/** The area of the sheet that a range covers. */
export function areaOf(range: Extract<Ast, { type: 'range' }>): Area {
  const { start, end = range.start } = range;
  // Both ends of a range have rows, or neither has; columns alike.
  const [top, bottom] =
    start.row === null ? [0, SHEET_ROWS - 1] : order(start.row, end.row!);
  const [left, right] =
    start.col === null ? [0, SHEET_COLUMNS - 1] : order(start.col, end.col!);
  return { top, left, bottom, right };
}

const order = (a: number, b: number): [number, number] =>
  a <= b ? [a, b] : [b, a];

function run(ast: Ast, context: Context): Result {
  switch (ast.type) {
    case 'number':
    case 'string':
    case 'boolean':
    case 'error':
      return ast.value;
    case 'missing':
      return 0;
    case 'name':
      return error('#NAME?');
    case 'range':
      return new Reference(ast.sheet, areaOf(ast));
    case 'paren':
      return run(ast.inner, context);
    case 'unary': {
      const operand = scalar(run(ast.operand, context), context);
      if (ast.operator === '+') return operand;
      const number = toNumber(operand);
      return isError(number) ? number : -number;
    }
    case 'percent': {
      const number = toNumber(scalar(run(ast.operand, context), context));
      return isError(number) ? number : number / 100;
    }
    case 'binary':
      return ast.operator === ' '
        ? intersect(run(ast.left, context), run(ast.right, context))
        : binary(
            ast.operator,
            scalar(run(ast.left, context), context),
            scalar(run(ast.right, context), context),
          );
    case 'call': {
      const spec = LIBRARY[ast.name];
      if (!spec) return error('#NAME?');
      const result = spec.call(new Arguments(ast.args, context));
      if (typeof result !== 'number') return result;
      // Out of range or undefined, as `SQRT(-1)`; and never a negative 0.
      return Number.isFinite(result) ? result + 0 : NUM;
    }
  }
}

/**
 * The one value of a result. A range where one value is expected gives
 * its cell in the formula's row or column, as Excel does.
 */
function scalar(result: Result, context: Context): Value {
  if (!(result instanceof Reference)) return result;
  const { sheet, area } = result;
  if (area.top === area.bottom && area.left === area.right) {
    return context.cell(sheet, area.top, area.left);
  }
  if (area.left === area.right && within(context.row, area.top, area.bottom)) {
    return context.cell(sheet, context.row, area.left);
  }
  if (area.top === area.bottom && within(context.col, area.left, area.right)) {
    return context.cell(sheet, area.top, context.col);
  }
  return VALUE;
}

const within = (n: number, low: number, high: number) => n >= low && n <= high;

function intersect(left: Result, right: Result): Result {
  if (!(left instanceof Reference)) return isError(left) ? left : VALUE;
  if (!(right instanceof Reference)) return isError(right) ? right : VALUE;
  if (left.sheet?.toLowerCase() !== right.sheet?.toLowerCase()) return VALUE;
  const a = left.area;
  const b = right.area;
  const area = {
    top: Math.max(a.top, b.top),
    left: Math.max(a.left, b.left),
    bottom: Math.min(a.bottom, b.bottom),
    right: Math.min(a.right, b.right),
  };
  if (area.top > area.bottom || area.left > area.right) {
    return error('#NULL!');
  }
  return new Reference(left.sheet, area);
}

function binary(operator: BinaryOperator, left: Value, right: Value): Value {
  if (isError(left)) return left;
  if (isError(right)) return right;
  if (operator === '&') return toText(left) + toText(right);
  if (
    operator === '=' ||
    operator === '<>' ||
    operator === '<' ||
    operator === '>' ||
    operator === '<=' ||
    operator === '>='
  ) {
    const order = compare(left, right);
    switch (operator) {
      case '=':
        return order === 0;
      case '<>':
        return order !== 0;
      case '<':
        return order < 0;
      case '>':
        return order > 0;
      case '<=':
        return order <= 0;
      default:
        return order >= 0;
    }
  }
  const a = toNumber(left);
  if (isError(a)) return a;
  const b = toNumber(right);
  if (isError(b)) return b;
  let result: number | ErrorValue;
  switch (operator) {
    case '+':
      result = a + b;
      break;
    case '-':
      result = a - b;
      break;
    case '*':
      result = a * b;
      break;
    case '/':
      if (b === 0) return DIV0;
      result = a / b;
      break;
    default:
      result = power(a, b);
  }
  return typeof result !== 'number' || Number.isFinite(result) ? result : NUM;
}

class Arguments implements Args {
  constructor(
    private readonly asts: readonly Ast[],
    private readonly context: Context,
  ) {}

  get length(): number {
    return this.asts.length;
  }

  given(index: number): boolean {
    const ast = this.asts[index];
    return ast !== undefined && ast.type !== 'missing';
  }

  result(index: number): Result {
    const ast = this.asts[index];
    return ast ? run(ast, this.context) : null;
  }

  value(index: number): Value {
    return scalar(this.result(index), this.context);
  }

  number(index: number): number | ErrorValue {
    return toNumber(this.value(index));
  }

  reference(index: number): Reference | ErrorValue {
    const result = this.result(index);
    if (result instanceof Reference || isError(result)) return result;
    return VALUE;
  }

  cells(
    reference: Reference,
    visit: (value: Value, row: number, col: number) => void,
  ): ErrorValue | undefined {
    return this.context.cells(reference.sheet, reference.area, visit);
  }

  cell(reference: Reference, row: number, col: number): Value {
    return this.context.cell(reference.sheet, row, col);
  }

  numbers({
    from = 0,
    to = this.asts.length,
    all = false,
    ordered = false,
  } = {}): number[] | ErrorValue {
    const found: number[] = [];
    for (let index = from; index < to; index++) {
      const result = this.result(index);
      if (!(result instanceof Reference)) {
        const number = toNumber(result);
        if (isError(number)) return number;
        found.push(number);
        continue;
      }
      const cells: { n: number; row: number; col: number }[] = [];
      let failure: ErrorValue | undefined;
      const missing = this.cells(result, (value, row, col) => {
        let n: number;
        if (typeof value === 'number') n = value;
        else if (isError(value)) return void (failure ??= value);
        else if (!all) return;
        else n = value === true ? 1 : 0;
        if (ordered) cells.push({ n, row, col });
        else found.push(n);
      });
      if (missing ?? failure) return (missing ?? failure)!;
      cells.sort((a, b) => a.row - b.row || a.col - b.col);
      for (const { n } of cells) found.push(n);
    }
    return found;
  }
}
