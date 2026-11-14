import { isError, type ErrorValue, type Value } from '../value';
import {
  DIV0,
  digits15,
  NUM,
  power,
  Reference,
  VALUE,
  type Args,
  type Spec,
} from './values';

type Outcome = number | ErrorValue;

/** A function of one number. */
function unary(fn: (n: number) => Outcome): Spec {
  return {
    min: 1,
    max: 1,
    call(args) {
      const n = args.number(0);
      return isError(n) ? n : fn(n);
    },
  };
}

/**
 * A function of numbers, the ones after `min` optional: one not given is
 * `undefined`, and one left out, as in `LOG(8,)`, is 0.
 */
function numeric(
  min: number,
  max: number,
  fn: (...n: (number | undefined)[]) => Outcome,
): Spec {
  return {
    min,
    max,
    call(args) {
      const numbers: number[] = [];
      for (let index = 0; index < args.length; index++) {
        const n = args.number(index);
        if (isError(n)) return n;
        numbers.push(n);
      }
      return fn(...numbers);
    },
  };
}

/** A function of all the numbers of its arguments, as `SUM` reads them. */
export function aggregate(fn: (values: number[]) => Outcome): Spec {
  return {
    min: 1,
    max: 255,
    call(args) {
      const values = args.numbers();
      return isError(values) ? values : fn(values);
    },
  };
}

export const sum = (values: number[]) =>
  values.reduce((total, n) => total + n, 0);

type Rounding = 'near' | 'up' | 'down';

/**
 * A number rounded to some digits after the point (before it when
 * negative), away from zero for `near` halves and `up`. The number is
 * first taken to 15 digits, so that 1.005 rounds as Excel shows it.
 */
function round(n: number, digits: number, rounding: Rounding): number {
  const places = Math.trunc(digits);
  const scale = 10 ** Math.abs(places);
  const scaled = digits15(
    places >= 0 ? Math.abs(n) * scale : Math.abs(n) / scale,
  );
  const whole =
    rounding === 'near'
      ? Math.floor(scaled + 0.5)
      : rounding === 'up'
        ? Math.ceil(scaled)
        : Math.floor(scaled);
  return Math.sign(n) * (places >= 0 ? whole / scale : whole * scale);
}

/** `n` taken to a multiple of `step` by `to`, as `CEILING` and its kind do. */
function multiple(n: number, step: number, to: (q: number) => number): number {
  return digits15(to(digits15(n / step)) * step);
}

function factorial(n: number, step: number): number {
  let product = 1;
  for (let k = Math.trunc(n); k > 1; k -= step) product *= k;
  return product;
}

function gcd(a: number, b: number): number {
  return b === 0 ? a : gcd(b, a % b);
}

/** The whole numbers of the arguments of `GCD` and `LCM`. */
function integers(args: Args): number[] | ErrorValue {
  const values = args.numbers();
  if (isError(values)) return values;
  if (values.some((n) => n < 0)) return NUM;
  return values.map(Math.trunc);
}

/** The values of an argument laid out as its rows and columns. */
export function grid(
  args: Args,
  index: number,
): { rows: number; cols: number; values: Value[] } | ErrorValue {
  const result = args.result(index);
  if (!(result instanceof Reference)) {
    return { rows: 1, cols: 1, values: [result] };
  }
  const { rows, cols, area } = result;
  const values: Value[] = new Array<Value>(rows * cols).fill(null);
  const missing = args.cells(result, (value, row, col) => {
    values[(row - area.top) * cols + (col - area.left)] = value;
  });
  return missing ?? { rows, cols, values };
}

export const MATH: Record<string, Spec> = {
  ABS: unary(Math.abs),
  ACOS: unary(Math.acos),
  ACOSH: unary(Math.acosh),
  ASIN: unary(Math.asin),
  ASINH: unary(Math.asinh),
  ATAN: unary(Math.atan),
  ATAN2: numeric(2, 2, (x, y) =>
    x === 0 && y === 0 ? DIV0 : Math.atan2(y!, x!),
  ),
  ATANH: unary(Math.atanh),
  CEILING: numeric(2, 2, (n, step) => {
    if (step === 0) return 0;
    if (n! > 0 && step! < 0) return NUM;
    return multiple(n!, step!, Math.ceil);
  }),
  'CEILING.MATH': {
    ...numeric(1, 3, (n, step = 1, mode = 0) => {
      if (step === 0) return 0;
      const away = n! < 0 && mode !== 0;
      return multiple(n!, Math.abs(step), away ? Math.floor : Math.ceil);
    }),
    prefixed: true,
  },
  COMBIN: numeric(2, 2, (n, k) => {
    const [total, chosen] = [Math.trunc(n!), Math.trunc(k!)];
    if (total < 0 || chosen < 0 || total < chosen) return NUM;
    let result = 1;
    for (let i = 1; i <= chosen; i++) {
      result = (result * (total - chosen + i)) / i;
    }
    return Math.round(result);
  }),
  COS: unary(Math.cos),
  COSH: unary(Math.cosh),
  DEGREES: unary((n) => (n * 180) / Math.PI),
  EVEN: unary((n) => Math.sign(n) * Math.ceil(Math.abs(n) / 2) * 2),
  EXP: unary(Math.exp),
  FACT: unary((n) => (n < 0 ? NUM : factorial(n, 1))),
  FACTDOUBLE: unary((n) => (n < -1 ? NUM : factorial(n, 2))),
  FLOOR: numeric(2, 2, (n, step) => {
    if (step === 0) return n === 0 ? 0 : DIV0;
    if (n! > 0 && step! < 0) return NUM;
    return multiple(n!, step!, Math.floor);
  }),
  'FLOOR.MATH': {
    ...numeric(1, 3, (n, step = 1, mode = 0) => {
      if (step === 0) return 0;
      const towardZero = n! < 0 && mode !== 0;
      return multiple(n!, Math.abs(step), towardZero ? Math.ceil : Math.floor);
    }),
    prefixed: true,
  },
  GCD: {
    min: 1,
    max: 255,
    call(args) {
      const values = integers(args);
      return isError(values) ? values : values.reduce(gcd, 0);
    },
  },
  INT: unary(Math.floor),
  LCM: {
    min: 1,
    max: 255,
    call(args) {
      const values = integers(args);
      if (isError(values)) return values;
      return values.reduce(
        (a, b) => (a === 0 || b === 0 ? 0 : (a * b) / gcd(a, b)),
        1,
      );
    },
  },
  LN: unary(Math.log),
  LOG: numeric(1, 2, (n, base = 10) => {
    if (n! <= 0 || base <= 0) return NUM;
    return base === 1 ? DIV0 : Math.log(n!) / Math.log(base);
  }),
  LOG10: unary(Math.log10),
  MOD: numeric(2, 2, (n, d) =>
    d === 0 ? DIV0 : n! - d! * Math.floor(n! / d!),
  ),
  MROUND: numeric(2, 2, (n, step) => {
    if (step === 0) return 0;
    if (n! * step! < 0) return NUM;
    return multiple(n!, step!, (q) => Math.floor(q + 0.5));
  }),
  ODD: unary((n) => {
    const up = Math.ceil(Math.abs(n));
    return (n < 0 ? -1 : 1) * (up % 2 === 0 ? up + 1 : up);
  }),
  PI: numeric(0, 0, () => Math.PI),
  POWER: numeric(2, 2, (a, b) => power(a!, b!)),
  PRODUCT: aggregate((values) =>
    values.length === 0 ? 0 : values.reduce((a, b) => a * b),
  ),
  QUOTIENT: numeric(2, 2, (n, d) => (d === 0 ? DIV0 : Math.trunc(n! / d!))),
  RADIANS: unary((n) => (n * Math.PI) / 180),
  ROUND: numeric(2, 2, (n, digits) => round(n!, digits!, 'near')),
  ROUNDDOWN: numeric(2, 2, (n, digits) => round(n!, digits!, 'down')),
  ROUNDUP: numeric(2, 2, (n, digits) => round(n!, digits!, 'up')),
  SIGN: unary(Math.sign),
  SIN: unary(Math.sin),
  SINH: unary(Math.sinh),
  SQRT: unary(Math.sqrt),
  SQRTPI: unary((n) => Math.sqrt(n * Math.PI)),
  SUM: aggregate(sum),
  SUMPRODUCT: {
    min: 1,
    max: 255,
    call(args) {
      const grids = [];
      for (let index = 0; index < args.length; index++) {
        const each = grid(args, index);
        if (isError(each)) return each;
        grids.push(each);
      }
      const [first] = grids;
      if (grids.some((g) => g.rows !== first!.rows || g.cols !== first!.cols)) {
        return VALUE;
      }
      let total = 0;
      for (let cell = 0; cell < first!.values.length; cell++) {
        let product = 1;
        for (const { values } of grids) {
          const value = values[cell]!;
          if (isError(value)) return value;
          product *= typeof value === 'number' ? value : 0;
        }
        total += product;
      }
      return total;
    },
  },
  SUMSQ: aggregate((values) => sum(values.map((n) => n * n))),
  TAN: unary(Math.tan),
  TANH: unary(Math.tanh),
  TRUNC: numeric(1, 2, (n, digits = 0) => round(n!, digits, 'down')),
};
