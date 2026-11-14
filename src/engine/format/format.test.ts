import { describe, expect, it } from 'vitest';
import { editText, formatValue, parseInput } from './index';

describe('parseInput', () => {
  it.each([
    ['2', 2],
    [' 2.5 ', 2.5],
    ['-3', -3],
    ['.5', 0.5],
    ['(5)', -5],
  ])('recognizes %j as a plain number', (text, value) => {
    expect(parseInput(text)).toEqual({ value });
  });

  it.each([
    ['1,234.5', 1234.5, '#,##0.00'],
    ['12%', 0.12, '0%'],
    ['1e3', 1000, '0.00E+00'],
    ['1/2/2026', 46024, 'm/d/yyyy'],
    ['2026-01-02', 46024, 'm/d/yyyy'],
    ['2-Jan', 46024, 'd-mmm'],
    ['13:45', 0.5729166666666666, 'h:mm'],
    ['1:30 PM', 0.5625, 'h:mm AM/PM'],
  ])('recognizes %j with its format', (text, value, format) => {
    const parsed = parseInput(text);
    expect(parsed.format).toBe(format);
    // Dates take the current year when none is typed.
    if (text !== '2-Jan') expect(parsed.value).toBeCloseTo(value, 10);
  });

  it('gives a date typed without a year the year it is told, and says which', () => {
    expect(parseInput('2-Jan', 2030)).toEqual({
      value: 47485, // 1/2/2030
      format: 'd-mmm',
      year: 2030,
    });
    expect(parseInput('2-Jan').year).toBe(new Date().getFullYear());
    expect(parseInput('1/2/2026', 2030)).toEqual({
      value: 46024,
      format: 'm/d/yyyy',
    });
    expect(parseInput('13:45', 2030).year).toBeUndefined();
  });

  it.each([
    ['TRUE', true],
    ['false', false],
  ])('recognizes %j as a boolean', (text, value) => {
    expect(parseInput(text)).toEqual({ value });
  });

  it.each(['abc', '2 apples', '1.2.3', '-', 'TRUE!'])(
    'keeps %j as text',
    (text) => {
      expect(parseInput(text)).toEqual({ value: text });
    },
  );

  it('reads an empty entry as an empty cell', () => {
    expect(parseInput('')).toEqual({ value: null });
  });
});

describe('formatValue', () => {
  it('shows each kind of value in the General format', () => {
    expect(formatValue(3)).toBe('3');
    expect(formatValue(0.1 + 0.2)).toBe('0.3');
    expect(formatValue('abc')).toBe('abc');
    expect(formatValue(true)).toBe('TRUE');
    expect(formatValue(false)).toBe('FALSE');
    expect(formatValue(null)).toBe('');
    expect(formatValue({ error: '#VALUE!' })).toBe('#VALUE!');
  });

  it('shows numbers in a format code', () => {
    expect(formatValue(46024, 'm/d/yyyy')).toBe('1/2/2026');
    expect(formatValue(0.12, '0%')).toBe('12%');
    expect(formatValue(1234.5, '#,##0.00')).toBe('1,234.50');
    expect(formatValue(0.5625, 'h:mm AM/PM')).toBe('1:30 PM');
  });

  it('shows text and booleans as they are, whatever the format', () => {
    expect(formatValue('abc', '0%')).toBe('abc');
    expect(formatValue(true, 'm/d/yyyy')).toBe('TRUE');
  });
});

describe('editText', () => {
  it('gives the text that edits a value back to itself, as Excel does', () => {
    expect(editText(1234.5, '#,##0.00')).toBe('1234.5');
    expect(editText(0.1 + 0.2)).toBe('0.3');
    expect(editText(0.12, '0%')).toBe('12%');
    expect(editText(46024, 'd-mmm')).toBe('1/2/2026');
    expect(editText(0.5625, 'h:mm')).toBe('1:30:00 PM');
    expect(editText(46024.5625, 'm/d/yyyy h:mm')).toBe('1/2/2026 1:30:00 PM');
    expect(editText(true)).toBe('TRUE');
    expect(editText('abc', '0%')).toBe('abc');
    expect(editText(null)).toBe('');
  });

  it.each(['1,234.5', '12%', '1/2/2026', '13:45', '1/2/2026 13:45', 'TRUE'])(
    'reads %j back as the same value',
    (text) => {
      const { value, format } = parseInput(text);
      expect(parseInput(editText(value, format)).value).toBe(value);
    },
  );
});
