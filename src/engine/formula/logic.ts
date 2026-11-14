import { isError, type ErrorValue } from '../value';
import {
  compare,
  NA,
  Reference,
  toBoolean,
  VALUE,
  type Args,
  type Spec,
} from './values';

/**
 * The logical values of the arguments of `AND` and its kind: in references
 * numbers and booleans, ignoring text; typed directly anything that reads
 * as one. `#VALUE!` when there is none.
 */
function logicals(args: Args): boolean[] | ErrorValue {
  const found: boolean[] = [];
  for (let index = 0; index < args.length; index++) {
    const result = args.result(index);
    if (result instanceof Reference) {
      let failure: ErrorValue | undefined;
      const missing = args.cells(result, (value) => {
        if (isError(value)) failure ??= value;
        else if (typeof value === 'number') found.push(value !== 0);
        else if (typeof value === 'boolean') found.push(value);
      });
      if (missing ?? failure) return (missing ?? failure)!;
    } else {
      const logical = toBoolean(result);
      if (isError(logical)) return logical;
      found.push(logical);
    }
  }
  return found.length === 0 ? VALUE : found;
}

function logical(combine: (values: boolean[]) => boolean): Spec {
  return {
    min: 1,
    max: 255,
    call(args) {
      const values = logicals(args);
      return isError(values) ? values : combine(values);
    },
  };
}

/** The first argument, or the second when the first is an error that `catches`. */
function fallback(catches: (error: ErrorValue) => boolean): Spec {
  return {
    min: 2,
    max: 2,
    call(args) {
      const value = args.value(0);
      return isError(value) && catches(value) ? args.result(1) : value;
    },
  };
}

const constant = (value: boolean): Spec => ({
  min: 0,
  max: 0,
  call: () => value,
});

export const LOGIC: Record<string, Spec> = {
  AND: logical((values) => values.every(Boolean)),
  FALSE: constant(false),
  IF: {
    min: 2,
    max: 3,
    call(args) {
      const condition = toBoolean(args.value(0));
      if (isError(condition)) return condition;
      if (condition) return args.result(1);
      return args.length > 2 ? args.result(2) : false;
    },
  },
  IFERROR: fallback(() => true),
  IFNA: { ...fallback((error) => error.error === NA.error), prefixed: true },
  IFS: {
    min: 2,
    max: 254,
    pairs: true,
    prefixed: true,
    call(args) {
      for (let index = 0; index < args.length; index += 2) {
        const condition = toBoolean(args.value(index));
        if (isError(condition)) return condition;
        if (condition) return args.result(index + 1);
      }
      return NA;
    },
  },
  NOT: {
    min: 1,
    max: 1,
    call(args) {
      const value = toBoolean(args.value(0));
      return isError(value) ? value : !value;
    },
  },
  OR: logical((values) => values.some(Boolean)),
  SWITCH: {
    min: 3,
    max: 254,
    prefixed: true,
    call(args) {
      const value = args.value(0);
      if (isError(value)) return value;
      let index = 1;
      for (; index + 1 < args.length; index += 2) {
        const candidate = args.value(index);
        if (isError(candidate)) return candidate;
        if (compare(value, candidate) === 0) return args.result(index + 1);
      }
      // What is left is the default, when the count is even.
      return index < args.length ? args.result(index) : NA;
    },
  },
  TRUE: constant(true),
  XOR: {
    ...logical((values) => values.filter(Boolean).length % 2 === 1),
    prefixed: true,
  },
};
