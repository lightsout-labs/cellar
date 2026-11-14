import { formatValue } from '../format';
import { isError, type ErrorValue, type Value } from '../value';
import { wildcardsAt } from './conditional';
import { grid, round } from './math';
import {
  fail,
  numberIn,
  toText,
  typed,
  VALUE,
  type Args,
  type Spec,
  type Typed,
} from './values';

/** The longest text a cell holds, as in Excel. */
const MAX_LENGTH = 32_767;

/** A function of one text. */
const unary = (fn: (text: string) => Value, prefixed?: boolean): Spec =>
  typed(1, 1, (read) => fn(read.text(0)), prefixed);

/** A count of characters, which must not be negative. */
function count(read: Typed, index: number, fallback?: number): number {
  const n = read.integer(index, fallback);
  return n < 0 ? fail(VALUE) : n;
}

/** A position of a character, counted from 1. */
function position(read: Typed, index: number, fallback?: number): number {
  const n = read.integer(index, fallback);
  return n < 1 ? fail(VALUE) : n;
}

/** Text no longer than a cell holds. */
const bounded = (text: string) => (text.length > MAX_LENGTH ? VALUE : text);

/**
 * The texts of the arguments from `from` on, with the cells of references
 * row by row, empty ones as empty text.
 */
function texts(args: Args, from: number): string[] | ErrorValue {
  const found: string[] = [];
  for (let index = from; index < args.length; index++) {
    const each = grid(args, index);
    if (isError(each)) return each;
    for (const value of each.values) {
      if (isError(value)) return value;
      found.push(toText(value));
    }
  }
  return found;
}

/** The position of `find` in `within` from `start`, or `#VALUE!`. */
function locate(
  find: string,
  within: string,
  start: number,
  matches: (at: number) => boolean,
): number {
  if (start > Math.max(within.length, 1)) return fail(VALUE);
  for (let at = start - 1; at <= within.length - find.length; at++) {
    if (matches(at)) return at + 1;
  }
  if (find === '') return start;
  return fail(VALUE);
}

/** A number in a format code, rounded first to the decimals it shows. */
function formatted(read: Typed, code: (decimals: string) => string): string {
  const n = read.number(0);
  const decimals = Math.min(read.integer(1, 2), 127);
  const fraction = decimals > 0 ? `.${'0'.repeat(decimals)}` : '';
  return formatValue(round(n, decimals, 'near'), code(fraction));
}

const LETTER = /\p{L}/u;

/**
 * The characters of codes 128 to 159 in Windows-1252, the character set of
 * `CHAR` and `CODE` in Excel; the other codes up to 255 are Unicode's.
 */
const WINDOWS_1252 = '€\x81‚ƒ„…†‡ˆ‰Š‹Œ\x8dŽ\x8f\x90‘’“”•–—˜™š›œ\x9džŸ';

/** The character of a code in Windows-1252. */
const character = (code: number) =>
  code >= 128 && code < 160
    ? WINDOWS_1252[code - 128]!
    : String.fromCharCode(code);

/** The code of a character in Windows-1252: 63, a `?`, if it has none. */
function codeOf(char: string): number {
  const code = char.charCodeAt(0);
  if (code < 128 || (code >= 160 && code < 256)) return code;
  const at = WINDOWS_1252.indexOf(char);
  return at >= 0 ? at + 128 : 63;
}

export const TEXT_FUNCTIONS: Record<string, Spec> = {
  CHAR: typed(1, 1, (read) => {
    const code = read.integer(0);
    return code < 1 || code > 255 ? VALUE : character(code);
  }),
  CLEAN: unary((text) =>
    [...text].filter((char) => char.charCodeAt(0) >= 32).join(''),
  ),
  CODE: unary((text) => (text === '' ? VALUE : codeOf(text[0]!))),
  CONCAT: {
    min: 1,
    max: 253,
    prefixed: true,
    call(args) {
      const found = texts(args, 0);
      return isError(found) ? found : bounded(found.join(''));
    },
  },
  CONCATENATE: typed(1, 255, (read) => {
    const parts = [];
    for (let index = 0; index < read.args.length; index++) {
      parts.push(read.text(index));
    }
    return bounded(parts.join(''));
  }),
  DOLLAR: typed(1, 2, (read) =>
    formatted(read, (fraction) => `$#,##0${fraction};($#,##0${fraction})`),
  ),
  EXACT: typed(2, 2, (read) => read.text(0) === read.text(1)),
  FIND: typed(2, 3, (read) => {
    const [find, within] = [read.text(0), read.text(1)];
    return locate(find, within, position(read, 2, 1), (at) =>
      within.startsWith(find, at),
    );
  }),
  FIXED: typed(1, 3, (read) => {
    const plain = read.boolean(2, false);
    return formatted(read, (fraction) => (plain ? '0' : '#,##0') + fraction);
  }),
  LEFT: typed(1, 2, (read) => read.text(0).slice(0, count(read, 1, 1))),
  LEN: unary((text) => text.length),
  LOWER: unary((text) => text.toLowerCase()),
  MID: typed(3, 3, (read) => {
    const text = read.text(0);
    const start = position(read, 1);
    return text.slice(start - 1, start - 1 + count(read, 2));
  }),
  NUMBERVALUE: typed(
    1,
    3,
    (read) => {
      const decimal = read.text(1, '.')[0] ?? '.';
      const group = read.text(2, ',')[0] ?? ',';
      let text = read.text(0).replace(/\s/g, '');
      let scale = 1;
      while (text.endsWith('%')) {
        text = text.slice(0, -1);
        scale /= 100;
      }
      text = text.split(group).join('').replace(decimal, '.');
      if (text === '') return 0;
      const n = /^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(text)
        ? Number(text)
        : NaN;
      return Number.isNaN(n) ? VALUE : n * scale;
    },
    true,
  ),
  PROPER: unary((text) => {
    let result = '';
    let after = false;
    for (const char of text) {
      result += after ? char.toLowerCase() : char.toUpperCase();
      after = LETTER.test(char);
    }
    return result;
  }),
  REPLACE: typed(4, 4, (read) => {
    const text = read.text(0);
    const start = position(read, 1);
    const length = count(read, 2);
    return bounded(
      text.slice(0, start - 1) + read.text(3) + text.slice(start - 1 + length),
    );
  }),
  REPT: typed(2, 2, (read) => {
    const text = read.text(0);
    const times = count(read, 1);
    return text.length * times > MAX_LENGTH ? VALUE : text.repeat(times);
  }),
  RIGHT: typed(1, 2, (read) => {
    const text = read.text(0);
    const n = count(read, 1, 1);
    return n === 0 ? '' : text.slice(-n);
  }),
  SEARCH: typed(2, 3, (read) => {
    const find = read.text(0).toLowerCase();
    const within = read.text(1).toLowerCase();
    const start = position(read, 2, 1);
    if (!/[*?~]/.test(find)) {
      return locate(find, within, start, (at) => within.startsWith(find, at));
    }
    const matches = wildcardsAt(find);
    return locate('', within, start, (at) => matches(within, at));
  }),
  SUBSTITUTE: typed(3, 4, (read) => {
    const [text, old, replacement] = [read.text(0), read.text(1), read.text(2)];
    if (read.args.length < 4) {
      return old === '' ? text : bounded(text.split(old).join(replacement));
    }
    const instance = position(read, 3);
    if (old === '') return text;
    let at = -old.length;
    for (let found = 0; found < instance; found++) {
      at = text.indexOf(old, at + old.length);
      if (at < 0) return text;
    }
    return bounded(
      text.slice(0, at) + replacement + text.slice(at + old.length),
    );
  }),
  T: {
    min: 1,
    max: 1,
    call(args) {
      const value = args.value(0);
      return typeof value === 'string' || isError(value) ? value : '';
    },
  },
  TEXT: typed(2, 2, (read) => {
    const value = read.args.value(0);
    const code = read.text(1);
    if (isError(value)) return value;
    const number = typeof value === 'string' ? numberIn(value) : value;
    return formatValue(number ?? value, code);
  }),
  TEXTJOIN: typed(
    3,
    254,
    (read) => {
      const delimiter = read.text(0);
      const skipEmpty = read.boolean(1);
      const found = texts(read.args, 2);
      if (isError(found)) return found;
      const kept = skipEmpty ? found.filter((text) => text !== '') : found;
      return bounded(kept.join(delimiter));
    },
    true,
  ),
  TRIM: unary((text) => text.replace(/^ +| +$/g, '').replace(/ {2,}/g, ' ')),
  UNICHAR: typed(
    1,
    1,
    (read) => {
      const code = read.integer(0);
      const valid =
        code >= 1 && code <= 0x10ffff && (code < 0xd800 || code > 0xdfff);
      return valid ? String.fromCodePoint(code) : VALUE;
    },
    true,
  ),
  UNICODE: unary((text) => text.codePointAt(0) ?? VALUE, true),
  UPPER: unary((text) => text.toUpperCase()),
  VALUE: {
    min: 1,
    max: 1,
    call(args) {
      const value = args.value(0);
      if (value === null || typeof value === 'number' || isError(value)) {
        return value ?? 0;
      }
      return typeof value === 'string' ? (numberIn(value) ?? VALUE) : VALUE;
    },
  },
};
