import { isError, type ErrorValue } from '../value';
import { aggregate, grid, sum } from './math';
import {
  DIV0,
  NA,
  NUM,
  Reference,
  toNumber,
  type Args,
  type Spec,
} from './values';

type Outcome = number | ErrorValue;

/** A function of all the numbers of its arguments, as `AVERAGEA` reads them. */
function aggregateAll(fn: (values: number[]) => Outcome): Spec {
  return {
    min: 1,
    max: 255,
    call(args) {
      const values = args.numbers({ all: true });
      return isError(values) ? values : fn(values);
    },
  };
}

/** A function of the numbers of its first argument and a number. */
function ofArray(fn: (values: number[], k: number) => Outcome): Spec {
  return {
    min: 2,
    max: 2,
    call(args) {
      const values = args.numbers({ to: 1 });
      if (isError(values)) return values;
      const k = args.number(1);
      return isError(k) ? k : fn(values, k);
    },
  };
}

const mean = (values: number[]) =>
  values.length === 0 ? DIV0 : sum(values) / values.length;

/** The extreme of numbers, without spreading a whole column as arguments. */
const extreme =
  (pick: (a: number, b: number) => number) => (values: number[]) =>
    values.length === 0 ? 0 : values.reduce((a, b) => pick(a, b));

/** The sum of the squares of the deviations from the mean. */
function squares(values: number[]): number {
  const average = sum(values) / values.length;
  return sum(values.map((n) => (n - average) ** 2));
}

/** The variance of a sample, or of a whole population. */
const variance = (sample: boolean) => (values: number[]) =>
  values.length < (sample ? 2 : 1)
    ? DIV0
    : squares(values) / (values.length - (sample ? 1 : 0));

const deviation = (sample: boolean) => (values: number[]) => {
  const result = variance(sample)(values);
  return isError(result) ? result : Math.sqrt(result);
};

const ascending = (values: number[]) => [...values].sort((a, b) => a - b);

/** The `k`th largest number, or with `smallest` the `k`th smallest. */
const nth = (smallest: boolean) => (values: number[], k: number) => {
  const rank = Math.ceil(k);
  if (rank < 1 || rank > values.length) return NUM;
  const sorted = ascending(values);
  return sorted[smallest ? rank - 1 : values.length - rank]!;
};

/** The `k`th percentile, interpolated between the numbers either side. */
function percentile(values: number[], k: number): Outcome {
  if (values.length === 0 || k < 0 || k > 1) return NUM;
  const sorted = ascending(values);
  const position = (sorted.length - 1) * k;
  const below = Math.floor(position);
  const low = sorted[below]!;
  const high = sorted[Math.min(below + 1, sorted.length - 1)]!;
  return low + (position - below) * (high - low);
}

function quartile(values: number[], quart: number): Outcome {
  const q = Math.trunc(quart);
  return q < 0 || q > 4 ? NUM : percentile(values, q / 4);
}

/** The most frequent number, the first of those that are; `#N/A` if none repeats. */
function mode(values: number[]): Outcome {
  const counts = new Map<number, number>();
  for (const n of values) counts.set(n, (counts.get(n) ?? 0) + 1);
  let most = 1;
  for (const count of counts.values()) most = Math.max(most, count);
  return most > 1 ? values.find((n) => counts.get(n) === most)! : NA;
}

const modeSpec: Spec = {
  min: 1,
  max: 255,
  call(args) {
    const values = args.numbers({ ordered: true });
    return isError(values) ? values : mode(values);
  },
};

/** The rank of a number among those of a reference, or their average rank when `average`. */
function rank(average: boolean, prefixed: boolean): Spec {
  return {
    min: 2,
    max: 3,
    range: (index) => index === 1,
    prefixed,
    call(args) {
      const n = args.number(0);
      if (isError(n)) return n;
      const reference = args.reference(1);
      if (isError(reference)) return reference;
      const values = args.numbers({ from: 1, to: 2 });
      if (isError(values)) return values;
      const order = args.given(2) ? args.number(2) : 0;
      if (isError(order)) return order;
      let before = 0;
      let ties = 0;
      for (const value of values) {
        if (value === n) ties++;
        else if (order === 0 ? value > n : value < n) before++;
      }
      if (ties === 0) return NA;
      return before + 1 + (average ? (ties - 1) / 2 : 0);
    },
  };
}

/** Counts the cells of references that hold anything, and typed values. */
function countAll(args: Args): number | ErrorValue {
  let count = 0;
  for (let index = 0; index < args.length; index++) {
    const result = args.result(index);
    if (result instanceof Reference) {
      const missing = args.cells(result, () => count++);
      if (missing) return missing;
    } else {
      count++;
    }
  }
  return count;
}

const newer = (spec: Spec): Spec => ({ ...spec, prefixed: true });

export const STATISTICS: Record<string, Spec> = {
  AVEDEV: aggregate((values) => {
    if (values.length === 0) return NUM;
    const average = sum(values) / values.length;
    return sum(values.map((n) => Math.abs(n - average))) / values.length;
  }),
  AVERAGE: aggregate(mean),
  AVERAGEA: aggregateAll(mean),
  CORREL: {
    min: 2,
    max: 2,
    call(args) {
      const xs = grid(args, 0);
      if (isError(xs)) return xs;
      const ys = grid(args, 1);
      if (isError(ys)) return ys;
      if (xs.values.length !== ys.values.length) return NA;
      const pairs: [number, number][] = [];
      for (let index = 0; index < xs.values.length; index++) {
        const [x, y] = [xs.values[index]!, ys.values[index]!];
        if (isError(x)) return x;
        if (isError(y)) return y;
        if (typeof x === 'number' && typeof y === 'number') pairs.push([x, y]);
      }
      const mx = sum(pairs.map(([x]) => x)) / pairs.length;
      const my = sum(pairs.map(([, y]) => y)) / pairs.length;
      let [sxy, sxx, syy] = [0, 0, 0];
      for (const [x, y] of pairs) {
        sxy += (x - mx) * (y - my);
        sxx += (x - mx) ** 2;
        syy += (y - my) ** 2;
      }
      return sxx === 0 || syy === 0 ? DIV0 : sxy / Math.sqrt(sxx * syy);
    },
  },
  COUNT: {
    min: 1,
    max: 255,
    call(args) {
      let count = 0;
      for (let index = 0; index < args.length; index++) {
        const result = args.result(index);
        if (result instanceof Reference) {
          const missing = args.cells(result, (value) => {
            if (typeof value === 'number') count++;
          });
          if (missing) return missing;
        } else if (!isError(result) && !isError(toNumber(result))) {
          count++;
        }
      }
      return count;
    },
  },
  COUNTA: { min: 1, max: 255, call: countAll },
  COUNTBLANK: {
    min: 1,
    max: 1,
    range: () => true,
    call(args) {
      const reference = args.reference(0);
      if (isError(reference)) return reference;
      let blank = reference.rows * reference.cols;
      const missing = args.cells(reference, (value) => {
        if (value !== '') blank--;
      });
      return missing ?? blank;
    },
  },
  DEVSQ: aggregate((values) => (values.length === 0 ? NUM : squares(values))),
  GEOMEAN: aggregate((values) =>
    values.length === 0 || values.some((n) => n <= 0)
      ? NUM
      : Math.exp(sum(values.map(Math.log)) / values.length),
  ),
  HARMEAN: aggregate((values) =>
    values.length === 0 || values.some((n) => n <= 0)
      ? NUM
      : values.length / sum(values.map((n) => 1 / n)),
  ),
  LARGE: ofArray(nth(false)),
  MAX: aggregate(extreme(Math.max)),
  MAXA: aggregateAll(extreme(Math.max)),
  MEDIAN: aggregate((values) =>
    values.length === 0 ? NUM : percentile(values, 0.5),
  ),
  MIN: aggregate(extreme(Math.min)),
  MINA: aggregateAll(extreme(Math.min)),
  MODE: modeSpec,
  'MODE.SNGL': newer(modeSpec),
  PERCENTILE: ofArray(percentile),
  'PERCENTILE.INC': newer(ofArray(percentile)),
  QUARTILE: ofArray(quartile),
  'QUARTILE.INC': newer(ofArray(quartile)),
  RANK: rank(false, false),
  'RANK.AVG': rank(true, true),
  'RANK.EQ': rank(false, true),
  SMALL: ofArray(nth(true)),
  STDEV: aggregate(deviation(true)),
  'STDEV.P': newer(aggregate(deviation(false))),
  'STDEV.S': newer(aggregate(deviation(true))),
  STDEVP: aggregate(deviation(false)),
  VAR: aggregate(variance(true)),
  'VAR.P': newer(aggregate(variance(false))),
  'VAR.S': newer(aggregate(variance(true))),
  VARP: aggregate(variance(false)),
};
