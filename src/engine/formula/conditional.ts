import { isError, type ErrorValue, type Value } from '../value';
import { sum } from './math';
import {
  compareText,
  DIV0,
  digits15,
  numberIn,
  Reference,
  VALUE,
  type Args,
  type Spec,
} from './values';

/** Whether a cell meets a criterion. */
type Test = (cell: Value) => boolean;

/**
 * The test of a criterion in Excel's syntax: a value to equal, or text
 * that starts with a comparison operator, where `*` and `?` are wildcards
 * when it compares text for equality, and `~` escapes them.
 */
function criterion(value: Value): Test {
  if (value === null) return equalsNumber(0, '=');
  if (typeof value === 'number') return equalsNumber(value, '=');
  if (typeof value === 'boolean') return (cell) => cell === value;
  if (isError(value)) {
    return (cell) => isError(cell) && cell.error === value.error;
  }
  const [, operator = '', operand = ''] = /^(<=|>=|<>|=|<|>)?(.*)$/s.exec(
    value,
  )!;
  if (operand === '') {
    if (operator === '') return (cell) => cell === null || cell === '';
    if (operator === '=') return (cell) => cell === null;
    if (operator === '<>') return (cell) => cell !== null;
  }
  const number = numberIn(operand);
  if (number !== undefined) return equalsNumber(number, operator);
  const upper = operand.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') {
    const wanted = upper === 'TRUE';
    return comparing(operator, (cell) =>
      typeof cell === 'boolean' ? Number(cell) - Number(wanted) : null,
    );
  }
  if (operator === '' || operator === '=' || operator === '<>') {
    const matches = wildcards(operand);
    return comparing(operator, (cell) =>
      typeof cell === 'string' && matches(cell) ? 0 : 1,
    );
  }
  return comparing(operator, (cell) =>
    typeof cell === 'string' ? compareText(cell, operand) : null,
  );
}

/**
 * A test of a cell by its order against the operand: `null` when it
 * cannot be compared, which only `<>` accepts.
 */
function comparing(
  operator: string,
  order: (cell: Value) => number | null,
): Test {
  return (cell) => {
    const found = order(cell);
    if (operator === '<>') return found !== 0;
    if (found === null) return false;
    switch (operator) {
      case '<':
        return found < 0;
      case '<=':
        return found <= 0;
      case '>':
        return found > 0;
      case '>=':
        return found >= 0;
      default:
        return found === 0;
    }
  };
}

/**
 * Numbers compare with numbers; text that reads as the number equals it,
 * but is neither less nor more.
 */
function equalsNumber(number: number, operator: string): Test {
  const wanted = digits15(number);
  const equality = operator === '' || operator === '=' || operator === '<>';
  return comparing(operator, (cell) => {
    if (typeof cell === 'number') return digits15(cell) - wanted;
    if (typeof cell !== 'string' || !equality) return null;
    return numberIn(cell) === number ? 0 : null;
  });
}

/**
 * A test of text against a pattern with Excel's wildcards, without regard
 * to case. It never backtracks further than the last `*`, so that no
 * pattern takes more than the product of the two lengths.
 */
export function wildcards(text: string): (cell: string) => boolean {
  const match = matcher(text);
  return (cell) => match(cell.toLowerCase(), 0, false);
}

/**
 * A test of whether a pattern with Excel's wildcards matches the start of
 * text in lower case from a position, as `SEARCH` finds it.
 */
export function wildcardsAt(
  text: string,
): (subject: string, from: number) => boolean {
  const match = matcher(text);
  return (subject, from) => match(subject, from, true);
}

function matcher(
  text: string,
): (subject: string, start: number, prefix: boolean) => boolean {
  const kinds: ('*' | '?' | 'char')[] = [];
  const chars: string[] = [];
  for (let index = 0; index < text.length; index++) {
    let char = text[index]!;
    if (char === '~' && index + 1 < text.length) {
      char = text[++index]!;
      kinds.push('char');
    } else {
      kinds.push(char === '*' || char === '?' ? char : 'char');
    }
    chars.push(char.toLowerCase());
  }
  return (subject, start, prefix) => {
    let [at, from, star, resume] = [0, start, -1, start];
    while (from < subject.length) {
      if (prefix && at === kinds.length) return true;
      const kind = kinds[at];
      if (kind === '?' || (kind === 'char' && chars[at] === subject[from])) {
        at++;
        from++;
      } else if (kind === '*') {
        star = at++;
        resume = from;
      } else if (star >= 0) {
        at = star + 1;
        from = ++resume;
      } else {
        return false;
      }
    }
    while (kinds[at] === '*') at++;
    return at === kinds.length;
  };
}

interface Criterion {
  readonly reference: Reference;
  readonly test: Test;
}

/**
 * The criteria of the pairs of arguments from `from` up to `to`: each a
 * range and its criterion, all of the same shape.
 */
function criteria(
  args: Args,
  from: number,
  to = args.length,
): Criterion[] | ErrorValue {
  const found: Criterion[] = [];
  for (let index = from; index < to; index += 2) {
    const reference = args.reference(index);
    if (isError(reference)) return reference;
    found.push({ reference, test: criterion(args.value(index + 1)) });
  }
  const [first] = found;
  return found.every((each) => sameShape(each.reference, first!.reference))
    ? found
    : VALUE;
}

const sameShape = (a: Reference, b: Reference) =>
  a.rows === b.rows && a.cols === b.cols;

/** Whether the cells at an offset in the ranges of all the criteria meet them. */
function meets(
  args: Args,
  all: readonly Criterion[],
  down: number,
  across: number,
): boolean {
  return all.every(({ reference, test }) =>
    test(
      args.cell(
        reference,
        reference.area.top + down,
        reference.area.left + across,
      ),
    ),
  );
}

/**
 * The numbers of the cells of a range whose cells at the same place in
 * the ranges of the criteria meet them all, or the first error among them.
 */
function selected(
  args: Args,
  target: Reference,
  all: readonly Criterion[],
): number[] | ErrorValue {
  const found: number[] = [];
  let failure: ErrorValue | undefined;
  const { top, left } = target.area;
  const missing = args.cells(target, (value, row, col) => {
    if (typeof value !== 'number' && !isError(value)) return;
    if (!meets(args, all, row - top, col - left)) return;
    if (isError(value)) failure ??= value;
    else found.push(value);
  });
  return missing ?? failure ?? found;
}

/** How many places of the ranges of the criteria meet them all. */
function count(args: Args, all: readonly Criterion[]): number | ErrorValue {
  const shape = all[0]!.reference;
  // A criterion that empty cells fail limits the places to its cells.
  const strict = all.find(({ test }) => !test(null));
  let found = 0;
  if (strict) {
    const { top, left } = strict.reference.area;
    const missing = args.cells(strict.reference, (_, row, col) => {
      if (meets(args, all, row - top, col - left)) found++;
    });
    return missing ?? found;
  }
  // Every criterion accepts empty cells: count the places that fail.
  const seen = new Set<number>();
  let failing = 0;
  for (const { reference } of all) {
    const { top, left } = reference.area;
    const missing = args.cells(reference, (_, row, col) => {
      const place = (row - top) * shape.cols + (col - left);
      if (seen.has(place)) return;
      seen.add(place);
      if (!meets(args, all, row - top, col - left)) failing++;
    });
    if (missing) return missing;
  }
  return shape.rows * shape.cols - failing;
}

type Combine = (values: number[]) => number | ErrorValue;

const average: Combine = (values) =>
  values.length === 0 ? DIV0 : sum(values) / values.length;
const extreme =
  (pick: (a: number, b: number) => number): Combine =>
  (values) =>
    values.length === 0 ? 0 : values.reduce((a, b) => pick(a, b));

/**
 * `SUMIF` and `AVERAGEIF`: a range, a criterion, and the range to combine,
 * which takes the shape of the first from its top left cell.
 */
function single(combine: Combine): Spec {
  return {
    min: 2,
    max: 3,
    range: (index) => index !== 1,
    call(args) {
      const all = criteria(args, 0, 2);
      if (isError(all)) return all;
      const range = all[0]!.reference;
      let target = range;
      if (args.given(2)) {
        const given = args.reference(2);
        if (isError(given)) return given;
        const { top, left } = given.area;
        target = new Reference(given.sheet, {
          top,
          left,
          bottom: top + range.rows - 1,
          right: left + range.cols - 1,
        });
      }
      const values = selected(args, target, all);
      return isError(values) ? values : combine(values);
    },
  };
}

/** `SUMIFS` and its kind: the range to combine, then pairs of criteria. */
function multiple(combine: Combine, prefixed: boolean): Spec {
  return {
    min: 3,
    max: 255,
    pairs: true,
    range: (index) => index === 0 || index % 2 === 1,
    prefixed,
    call(args) {
      const target = args.reference(0);
      if (isError(target)) return target;
      const all = criteria(args, 1);
      if (isError(all)) return all;
      if (!sameShape(target, all[0]!.reference)) return VALUE;
      const values = selected(args, target, all);
      return isError(values) ? values : combine(values);
    },
  };
}

export const CONDITIONAL: Record<string, Spec> = {
  AVERAGEIF: single(average),
  AVERAGEIFS: multiple(average, false),
  COUNTIF: {
    min: 2,
    max: 2,
    range: (index) => index === 0,
    call(args) {
      const all = criteria(args, 0);
      return isError(all) ? all : count(args, all);
    },
  },
  COUNTIFS: {
    min: 2,
    max: 254,
    pairs: true,
    range: (index) => index % 2 === 0,
    call(args) {
      const all = criteria(args, 0);
      return isError(all) ? all : count(args, all);
    },
  },
  MAXIFS: multiple(extreme(Math.max), true),
  MINIFS: multiple(extreme(Math.min), true),
  SUMIF: single(sum),
  SUMIFS: multiple(sum, false),
};
