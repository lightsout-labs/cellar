import { parseDate } from '../format';
import { isError, type ErrorValue } from '../value';
import {
  fail,
  NUM,
  typed,
  VALUE,
  type Args,
  type Spec,
  type Typed,
} from './values';

// Dates are Excel's serial numbers: 1 is 1 January 1900, and 60 the
// 29 February 1900 that Excel keeps from Lotus 1-2-3, so that every date
// from 1 March 1900 is one day later than a count of real days.

const DAY = 86_400_000;
const START = Date.UTC(1899, 11, 31);
/** The serial number of 31 December 9999, the last date of Excel. */
const LAST = 2_958_465;

type Parts = [year: number, month: number, day: number];

/** The serial number of a date, with months and days past their ends carried. */
function serial(year: number, month: number, day: number): number {
  const months = year * 12 + month - 1;
  const first =
    (Date.UTC(Math.floor(months / 12), months % 12, 1) - START) / DAY;
  return (first >= 60 ? first + 1 : first) + day - 1;
}

/** The year, month and day of a serial number, without its time. */
function parts(date: number): Parts {
  if (date === 0) return [1900, 1, 0];
  if (date === 60) return [1900, 2, 29];
  const real = new Date(START + (date < 60 ? date : date - 1) * DAY);
  return [real.getUTCFullYear(), real.getUTCMonth() + 1, real.getUTCDate()];
}

/** The day of the week of a date, from 0 for Sunday. */
const weekday = (date: number) => (date + 6) % 7;

const lastDay = (year: number, month: number) =>
  serial(year, month + 1, 0) - serial(year, month, 0);

/** A date argument, without its time; `#NUM!` outside Excel's dates. */
function date(read: Typed, index: number, fallback?: number): number {
  const n = read.number(index, fallback);
  return n < 0 || n >= LAST + 1 ? fail(NUM) : Math.floor(n);
}

/** A result that must be one of Excel's dates. */
const dated = (n: number) => (n < 0 || n > LAST ? NUM : n);

/** The seconds of the time of day of a number, to the nearest second. */
function seconds(read: Typed): number {
  const n = read.number(0);
  if (n < 0) return fail(NUM);
  return Math.round((n - Math.floor(n)) * 86_400) % 86_400;
}

/** The serial number of a text that reads as a date or time. */
function textDate(read: Typed): number {
  const value = read.args.value(0);
  if (isError(value)) return fail(value);
  const found = typeof value === 'string' ? parseDate(value) : undefined;
  return found ?? fail(VALUE);
}

/** The local date and time of the clock, as a serial number. */
function now(): number {
  const t = new Date();
  const local = Date.UTC(
    t.getFullYear(),
    t.getMonth(),
    t.getDate(),
    t.getHours(),
    t.getMinutes(),
    t.getSeconds(),
    t.getMilliseconds(),
  );
  return (local - START) / DAY + 1;
}

/** The day of the week that starts the week, from 0 for Sunday, by type. */
const WEEK_STARTS: Record<number, number> = {
  1: 0,
  2: 1,
  11: 1,
  12: 2,
  13: 3,
  14: 4,
  15: 5,
  16: 6,
  17: 0,
};

/** The ISO 8601 week of a date: weeks start on Monday, week 1 has a Thursday. */
function isoWeek(day: number): number {
  const thursday = day - ((weekday(day) + 6) % 7) + 3;
  const [year] = parts(thursday);
  return Math.floor((thursday - serial(year, 1, 1)) / 7) + 1;
}

/**
 * The days of the week that are not working days, by day from Sunday, as
 * the weekend argument of `NETWORKDAYS.INTL` gives them; Saturday and
 * Sunday for the functions without `.INTL`.
 */
function weekend(args: Args, intl: boolean): boolean[] | ErrorValue {
  const days = new Array<boolean>(7).fill(false);
  const value = intl && args.given(2) ? args.value(2) : 1;
  if (isError(value)) return value;
  if (typeof value === 'string') {
    if (!/^[01]{7}$/.test(value) || value === '1111111') return VALUE;
    // From Monday.
    [...value].forEach((flag, at) => (days[(at + 1) % 7] = flag === '1'));
    return days;
  }
  const n = typeof value === 'number' ? value : NaN;
  if (n >= 1 && n <= 7 && Number.isInteger(n)) {
    days[(n + 5) % 7] = days[(n + 6) % 7] = true;
  } else if (n >= 11 && n <= 17 && Number.isInteger(n)) {
    days[n - 11] = true;
  } else {
    return NUM;
  }
  return days;
}

/** The holidays of an argument, as a set of dates. */
function holidays(args: Args, index: number): Set<number> | ErrorValue {
  if (index >= args.length) return new Set();
  const found = args.numbers({ from: index, to: index + 1 });
  return isError(found) ? found : new Set(found.map(Math.floor));
}

/** The working days from `start` to `end`, both counted; negative backwards. */
function working(read: Typed, intl: boolean): number {
  const start = date(read, 0);
  const end = date(read, 1);
  const off = weekend(read.args, intl);
  const free = holidays(read.args, intl ? 3 : 2);
  if (isError(off)) return fail(off);
  if (isError(free)) return fail(free);
  const step = start <= end ? 1 : -1;
  let total = 0;
  for (let day = start; day !== end + step; day += step) {
    if (!off[weekday(day)] && !free.has(day)) total += step;
  }
  return total;
}

/** The working day `days` working days from a date. */
function workday(read: Typed, intl: boolean): number {
  let day = date(read, 0);
  let left = read.integer(1);
  const off = weekend(read.args, intl);
  const free = holidays(read.args, intl ? 3 : 2);
  if (isError(off)) return fail(off);
  if (isError(free)) return fail(free);
  const step = Math.sign(left);
  while (left !== 0) {
    day += step;
    if (day < 0 || day > LAST) return fail(NUM);
    if (!off[weekday(day)] && !free.has(day)) left -= step;
  }
  return day;
}

/** Days between two dates counted as twelve months of 30 days. */
function days360(from: Parts, to: Parts, european: boolean): number {
  const [y1, m1] = from;
  const [y2, m2] = to;
  let d1 = from[2];
  let d2 = to[2];
  if (european) {
    d1 = Math.min(d1, 30);
    d2 = Math.min(d2, 30);
  } else {
    const endOfFeb = ([y, m, d]: Parts) => m === 2 && d === lastDay(y, m);
    if (endOfFeb(from) && endOfFeb(to)) d2 = 30;
    if (endOfFeb(from) || d1 === 31) d1 = 30;
    if (d2 === 31 && d1 === 30) d2 = 30;
  }
  return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
}

const isLeap = (year: number) =>
  (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;

/** The fraction of a year between two dates by actual days (basis 1). */
function actualYears(start: number, end: number): number {
  const [y1, m1, d1] = parts(start);
  const [y2, m2, d2] = parts(end);
  const days = end - start;
  const withinYear =
    y1 === y2 || (y2 === y1 + 1 && (m1 > m2 || (m1 === m2 && d1 >= d2)));
  if (!withinYear) {
    const span = serial(y2 + 1, 1, 1) - serial(y1, 1, 1);
    return days / (span / (y2 - y1 + 1));
  }
  let leap = y1 === y2 && isLeap(y1);
  for (let year = y1; year <= y2 && !leap; year++) {
    const february29 = serial(year, 2, 29);
    leap = isLeap(year) && february29 >= start && february29 <= end;
  }
  return days / (leap ? 366 : 365);
}

export const DATE_FUNCTIONS: Record<string, Spec> = {
  DATE: typed(3, 3, (read) => {
    let year = read.integer(0);
    if (year >= 0 && year < 1900) year += 1900;
    if (year < 1900 || year > 9999) return NUM;
    return dated(serial(year, read.integer(1), read.integer(2)));
  }),
  DATEDIF: typed(3, 3, (read) => {
    const start = date(read, 0);
    const end = date(read, 1);
    if (start > end) return NUM;
    const [y1, m1, d1] = parts(start);
    const [y2, m2, d2] = parts(end);
    const months = (y2 - y1) * 12 + m2 - m1 - (d2 < d1 ? 1 : 0);
    switch (read.text(2).toUpperCase()) {
      case 'Y':
        return Math.floor(months / 12);
      case 'M':
        return months;
      case 'D':
        return end - start;
      case 'YM':
        return months % 12;
      case 'MD':
        return d2 >= d1 ? d2 - d1 : end - serial(y2, m2 - 1, d1);
      case 'YD': {
        const since = serial(y2, m1, d1);
        return end - (since <= end ? since : serial(y2 - 1, m1, d1));
      }
      default:
        return NUM;
    }
  }),
  DATEVALUE: typed(1, 1, (read) => Math.floor(textDate(read))),
  DAY: typed(1, 1, (read) => parts(date(read, 0))[2]),
  DAYS: typed(2, 2, (read) => date(read, 0) - date(read, 1), true),
  DAYS360: typed(2, 3, (read) => {
    const start = parts(date(read, 0));
    const end = parts(date(read, 1));
    return days360(start, end, read.boolean(2, false));
  }),
  EDATE: typed(2, 2, (read) => {
    const [year, month, day] = parts(date(read, 0));
    const target = month + read.integer(1);
    return dated(serial(year, target, Math.min(day, lastDay(year, target))));
  }),
  EOMONTH: typed(2, 2, (read) => {
    const [year, month] = parts(date(read, 0));
    return dated(serial(year, month + read.integer(1) + 1, 0));
  }),
  HOUR: typed(1, 1, (read) => Math.floor(seconds(read) / 3600)),
  ISOWEEKNUM: typed(1, 1, (read) => isoWeek(date(read, 0)), true),
  MINUTE: typed(1, 1, (read) => Math.floor(seconds(read) / 60) % 60),
  MONTH: typed(1, 1, (read) => parts(date(read, 0))[1]),
  NETWORKDAYS: typed(2, 3, (read) => working(read, false)),
  'NETWORKDAYS.INTL': typed(2, 4, (read) => working(read, true), true),
  NOW: typed(0, 0, now),
  SECOND: typed(1, 1, (read) => seconds(read) % 60),
  TIME: typed(3, 3, (read) => {
    const total =
      read.integer(0) * 3600 + read.integer(1) * 60 + read.integer(2);
    return total < 0 ? NUM : (total % 86_400) / 86_400;
  }),
  TIMEVALUE: typed(1, 1, (read) => {
    const n = textDate(read);
    return n - Math.floor(n);
  }),
  TODAY: typed(0, 0, () => Math.floor(now())),
  WEEKDAY: typed(1, 2, (read) => {
    const day = weekday(date(read, 0));
    const type = read.integer(1, 1);
    if (type === 3) return (day + 6) % 7;
    const first = WEEK_STARTS[type];
    return first === undefined ? NUM : ((day - first + 7) % 7) + 1;
  }),
  WEEKNUM: typed(1, 2, (read) => {
    const day = date(read, 0);
    const type = read.integer(1, 1);
    if (type === 21) return isoWeek(day);
    const first = WEEK_STARTS[type];
    if (first === undefined) return NUM;
    const january1 = serial(parts(day)[0], 1, 1);
    const offset = (weekday(january1) - first + 7) % 7;
    return Math.floor((day - january1 + offset) / 7) + 1;
  }),
  WORKDAY: typed(2, 3, (read) => workday(read, false)),
  'WORKDAY.INTL': typed(2, 4, (read) => workday(read, true), true),
  YEAR: typed(1, 1, (read) => parts(date(read, 0))[0]),
  YEARFRAC: typed(2, 3, (read) => {
    let start = date(read, 0);
    let end = date(read, 1);
    if (start > end) [start, end] = [end, start];
    switch (read.integer(2, 0)) {
      case 0:
        return days360(parts(start), parts(end), false) / 360;
      case 1:
        return actualYears(start, end);
      case 2:
        return (end - start) / 360;
      case 3:
        return (end - start) / 365;
      case 4:
        return days360(parts(start), parts(end), true) / 360;
      default:
        return NUM;
    }
  }),
};
