import { parseInput } from '../format';
import { isError, type ErrorValue, type Value } from '../value';

/** A rectangle of cells, its bounds included. Rows and columns count from zero. */
export interface Area {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

/** The seven error values of Excel (R13). */
export type ErrorCode =
  '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A';

export const error = (code: ErrorCode): ErrorValue => ({ error: code });
export const VALUE = error('#VALUE!');
export const DIV0 = error('#DIV/0!');
export const NUM = error('#NUM!');
export const NA = error('#N/A');

/** A reference as an operand evaluates to, before its cells are read. */
export class Reference {
  constructor(
    readonly sheet: string | undefined,
    readonly area: Area,
  ) {}

  get rows(): number {
    return this.area.bottom - this.area.top + 1;
  }

  get cols(): number {
    return this.area.right - this.area.left + 1;
  }
}

export type Result = Value | Reference;

/**
 * The arguments of a function call. Each is evaluated only when asked
 * for, so that `IF` and its kind evaluate only the branch they choose.
 */
export interface Args {
  readonly length: number;
  /** Where the formula is, counted from zero. */
  readonly row: number;
  readonly col: number;
  /** Whether an argument is there, and not left out as in `ROUND(1,)`. */
  given(index: number): boolean;
  /** An argument evaluated, a reference left unread. */
  result(index: number): Result;
  /** The one value of an argument, as an operator reads it. */
  value(index: number): Value;
  /** The number an argument stands for, as arithmetic reads it. */
  number(index: number): number | ErrorValue;
  /** An argument that must be a reference, or the error it gives. */
  reference(index: number): Reference | ErrorValue;
  /**
   * Visits the cells of a reference that are not empty, in no set order;
   * gives `#REF!` when its sheet does not exist.
   */
  cells(
    reference: Reference,
    visit: (value: Value, row: number, col: number) => void,
  ): ErrorValue | undefined;
  /** The value of one cell of the sheet of a reference. */
  cell(reference: Reference, row: number, col: number): Value;
  /**
   * The numbers of the arguments from `from` up to `to`, or all of them,
   * as `SUM` reads them: in references only the numbers, typed directly
   * anything that reads as a number. With `all`, as `AVERAGEA` reads them:
   * text in references is 0 and booleans 1 or 0. With `ordered`, in the
   * order of the arguments and of the cells, row by row. The first error
   * found is given instead.
   */
  numbers(options?: {
    from?: number;
    to?: number;
    all?: boolean;
    ordered?: boolean;
  }): number[] | ErrorValue;
}

/** A function of the library: the arguments it takes, and what it does. */
export interface Spec {
  readonly min: number;
  readonly max: number;
  /** Whether the arguments after the first `min` come in pairs. */
  readonly pairs?: boolean;
  /**
   * Whether an argument must be a reference, as the ranges of `SUMIF`; a
   * formula that gives a constant there is refused, as Excel refuses it.
   */
  readonly range?: (index: number) => boolean;
  /** Whether files write it with `_xlfn.`, as a function newer than 2007. */
  readonly prefixed?: boolean;
  call(args: Args): Result;
}

/**
 * The arguments of a function as one type each. A read that fails throws
 * the error it gives, and `typed` makes that error the result.
 */
export interface Typed {
  readonly args: Args;
  /** The text an argument stands for, as `&` reads it; `fallback` when not there. */
  text(index: number, fallback?: string): string;
  /** The number an argument stands for, as arithmetic reads it. */
  number(index: number, fallback?: number): number;
  /** The number of an argument without its fraction. */
  integer(index: number, fallback?: number): number;
  boolean(index: number, fallback?: boolean): boolean;
}

/** A function of `min` to `max` arguments, each read as one type. */
export function typed(
  min: number,
  max: number,
  fn: (read: Typed) => Result,
  prefixed?: boolean,
): Spec {
  return {
    min,
    max,
    ...(prefixed && { prefixed }),
    call(args) {
      try {
        return fn(new Reader(args));
      } catch (thrown) {
        if (thrown instanceof Failure) return thrown.error;
        throw thrown;
      }
    },
  };
}

class Failure {
  constructor(readonly error: ErrorValue) {}
}

/** Gives the result of a typed function: `error` instead of anything else. */
export function fail(error: ErrorValue): never {
  throw new Failure(error);
}

class Reader implements Typed {
  constructor(readonly args: Args) {}

  private read<T>(
    index: number,
    fallback: T | undefined,
    convert: (value: Exclude<Value, ErrorValue>) => T | ErrorValue,
  ): T {
    if (index >= this.args.length && fallback !== undefined) return fallback;
    const value = this.args.value(index);
    const result = isError(value) ? value : convert(value);
    return isError(result) ? fail(result) : result;
  }

  text(index: number, fallback?: string): string {
    return this.read(index, fallback, toText);
  }

  number(index: number, fallback?: number): number {
    return this.read(index, fallback, toNumber);
  }

  integer(index: number, fallback?: number): number {
    return Math.trunc(this.number(index, fallback));
  }

  boolean(index: number, fallback?: boolean): boolean {
    return this.read(index, fallback, toBoolean);
  }
}

/**
 * The number an operand stands for: an empty cell is 0, a boolean 1 or 0,
 * and text a number only if it reads as one.
 */
export function toNumber(value: Value): number | ErrorValue {
  if (value === null) return 0;
  if (typeof value === 'boolean') return Number(value);
  if (typeof value === 'string') return numberIn(value) ?? VALUE;
  return value;
}

/** The number a text reads as, as typed in a cell, if any. */
export function numberIn(text: string): number | undefined {
  const read = parseInput(text.trim()).value;
  return typeof read === 'number' ? read : undefined;
}

/** The text an operand of `&` stands for: numbers to 15 digits. */
export function toText(value: Exclude<Value, ErrorValue>): string {
  if (value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') {
    return String(digits15(value)).toUpperCase();
  }
  return value;
}

export function toBoolean(value: Value): boolean | ErrorValue {
  if (isError(value)) return value;
  if (typeof value === 'string') {
    const upper = value.toUpperCase();
    if (upper === 'TRUE' || upper === 'FALSE') return upper === 'TRUE';
    return VALUE;
  }
  return Boolean(value);
}

/** A number to the 15 significant digits that Excel keeps. */
export const digits15 = (n: number) => Number(n.toPrecision(15));

/**
 * The order of two values as Excel compares them: numbers before text
 * before booleans, text without regard to case, and an empty cell as the
 * empty value of the other's type.
 */
export function compare(left: Value, right: Value): number {
  const a = left ?? emptyLike(right);
  const b = right ?? emptyLike(left);
  const rank = (value: Value) =>
    typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2;
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === 'string' && typeof b === 'string') {
    return compareText(a, b);
  }
  // Numbers are equal when they agree to the 15 digits Excel keeps.
  return digits15(Number(a)) - digits15(Number(b));
}

export function compareText(a: string, b: string): number {
  const x = a.toLowerCase();
  const y = b.toLowerCase();
  return x < y ? -1 : x > y ? 1 : 0;
}

function emptyLike(value: Value): Value {
  if (typeof value === 'string') return '';
  if (typeof value === 'boolean') return false;
  return 0;
}

/** `a ^ b`, with Excel's errors. */
export function power(a: number, b: number): number | ErrorValue {
  if (a === 0 && b === 0) return NUM;
  if (a === 0 && b < 0) return DIV0;
  return a ** b;
}

/** A sheet name as a reference writes it: quoted unless it is a plain name. */
export function quoteSheetName(name: string): string {
  const plain =
    /^[A-Za-z_][\w.]*$/.test(name) &&
    !/^[A-Za-z]{1,3}\d+$/.test(name) &&
    !/^(TRUE|FALSE)$/i.test(name);
  return plain ? name : `'${name.replaceAll("'", "''")}'`;
}

/** The letters of a column, counted from zero: `A`, ..., `Z`, `AA`, ... */
export function columnName(col: number): string {
  let name = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return name;
}
