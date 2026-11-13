import { isError, type Value } from '../value';

const NUMBER = /^\s*[+-]?(?:\d+\.?\d*|\.\d+)\s*$/;

/**
 * The value that a typed entry stands for: a number when it reads as one,
 * an empty cell when nothing was typed, and otherwise the text itself.
 */
export function parseInput(text: string): number | string | null {
  if (text === '') return null;
  return NUMBER.test(text) ? Number(text) : text;
}

/** The text that a cell shows for a value. */
export function formatValue(value: Value): string {
  if (value === null) return '';
  if (isError(value)) return value.error;
  if (typeof value === 'number') return String(Number(value.toPrecision(15)));
  return value;
}
