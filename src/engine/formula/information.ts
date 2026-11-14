import { isError, type Value } from '../value';
import { NA, Reference, typed, type Spec } from './values';

/** A function of the one value of an argument, errors included. */
const of = (fn: (value: Value) => Value): Spec => ({
  min: 1,
  max: 1,
  call: (args) => fn(args.value(0)),
});

/** `ISEVEN` or `ISODD`, of the number without its fraction. */
const parity = (even: boolean): Spec =>
  typed(1, 1, (read) => (Math.trunc(read.number(0)) % 2 === 0) === even);

/** The numbers that `ERROR.TYPE` gives the error values. */
const ERROR_TYPES: Record<string, number> = {
  '#NULL!': 1,
  '#DIV/0!': 2,
  '#VALUE!': 3,
  '#REF!': 4,
  '#NAME?': 5,
  '#NUM!': 6,
  '#N/A': 7,
};

export const INFORMATION: Record<string, Spec> = {
  'ERROR.TYPE': of((value) =>
    isError(value) ? (ERROR_TYPES[value.error] ?? NA) : NA,
  ),
  ISBLANK: of((value) => value === null),
  ISERR: of((value) => isError(value) && value.error !== NA.error),
  ISERROR: of(isError),
  ISEVEN: parity(true),
  ISLOGICAL: of((value) => typeof value === 'boolean'),
  ISNA: of((value) => isError(value) && value.error === NA.error),
  ISNONTEXT: of((value) => typeof value !== 'string'),
  ISNUMBER: of((value) => typeof value === 'number'),
  ISODD: parity(false),
  ISREF: {
    min: 1,
    max: 1,
    call: (args) => args.result(0) instanceof Reference,
  },
  ISTEXT: of((value) => typeof value === 'string'),
  N: of((value) => {
    if (isError(value) || typeof value === 'number') return value;
    return value === true ? 1 : 0;
  }),
  NA: { min: 0, max: 0, call: () => NA },
  TYPE: of((value) => {
    if (isError(value)) return 16;
    if (typeof value === 'string') return 2;
    return typeof value === 'boolean' ? 4 : 1;
  }),
};
