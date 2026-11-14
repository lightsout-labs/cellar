import {
  dateFromSerial,
  dateToSerial,
  format,
  getFormatDateInfo,
  getFormatFromId,
  isDateFormat,
  isPercentFormat,
  parseValue,
} from 'numfmt';
import { isError, type Value } from '../value';

/** The format of a cell that was given none. */
export const GENERAL = 'General';

/** What a typed entry stands for. */
export interface Input {
  readonly value: number | string | boolean | null;
  /** The format code that the way it was typed implies, as `0%` for `12%`. */
  readonly format?: string;
  /** The year given to a date typed without one. */
  readonly year?: number;
}

/**
 * Recognizes a typed entry as Excel does: a number, a date, a time, a
 * percentage or a boolean, with the format it implies; an empty cell when
 * nothing was typed; and otherwise the text itself. Formulas are not
 * recognized here. A date typed without a year is in `year`, by default
 * the current one.
 */
export function parseInput(
  text: string,
  year: number = new Date().getFullYear(),
): Input {
  if (text === '') return { value: null };
  const parsed = parseValue(text);
  if (!parsed) return { value: text };
  if (typeof parsed.value === 'boolean') return { value: parsed.value };
  if (parsed.formatId === undefined) return { value: parsed.value };
  const code = getFormatFromId(parsed.formatId);
  const info = getFormatDateInfo(code);
  if (!isDateFormat(code) || info.type === 'time' || info.year) {
    return { value: parsed.value, format: code };
  }
  // The parser gave it the year of the clock; it gets the one asked for.
  const [, month, day] = dateFromSerial(parsed.value);
  const time = parsed.value - Math.floor(parsed.value);
  const value = dateToSerial([year, month!, day!])! + time;
  return { value, format: code, year };
}

/** The text that a cell shows for a value, in a format code. */
export function formatValue(value: Value, code: string = GENERAL): string {
  if (value === null) return '';
  if (isError(value)) return value.error;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'string') return value;
  // The common case, which the grid meets in every cell it draws.
  if (code === GENERAL && Number.isInteger(value) && Math.abs(value) < 1e11) {
    return String(value);
  }
  return format(code, value);
}

/**
 * The text that edits a value, as the formula bar shows it: numbers in
 * full, percentages with their sign, and dates and times in a form that is
 * read back as the same value.
 */
export function editText(value: Value, code: string = GENERAL): string {
  if (typeof value !== 'number') return formatValue(value);
  if (isPercentFormat(code)) return `${plain(value * 100)}%`;
  if (isDateFormat(code)) {
    const { type } = getFormatDateInfo(code);
    const parts = [];
    if (type !== 'time') parts.push('m/d/yyyy');
    if (type !== 'date') parts.push('h:mm:ss AM/PM');
    return format(parts.join(' '), value);
  }
  return plain(value);
}

/** A number with the 15 significant digits that Excel keeps. */
function plain(value: number): string {
  return String(Number(value.toPrecision(15)));
}
