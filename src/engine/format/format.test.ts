import { describe, expect, it } from 'vitest';
import { formatValue, parseInput } from './index';

describe('parseInput', () => {
  it.each([
    ['2', 2],
    [' 2.5 ', 2.5],
    ['-3', -3],
    ['.5', 0.5],
  ])('recognizes %j as a number', (text, value) => {
    expect(parseInput(text)).toBe(value);
  });

  it.each(['abc', '2 apples', '1.2.3', '-'])('keeps %j as text', (text) => {
    expect(parseInput(text)).toBe(text);
  });

  it('reads an empty entry as an empty cell', () => {
    expect(parseInput('')).toBeNull();
  });
});

describe('formatValue', () => {
  it('shows each kind of value', () => {
    expect(formatValue(3)).toBe('3');
    expect(formatValue(0.1 + 0.2)).toBe('0.3');
    expect(formatValue('abc')).toBe('abc');
    expect(formatValue(null)).toBe('');
    expect(formatValue({ error: '#VALUE!' })).toBe('#VALUE!');
  });
});
