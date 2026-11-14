import { describe, expect, it } from 'vitest';
import type { Value } from '../../engine/workbook';
import { COLUMN_WIDTH } from './geometry';
import { rowTexts } from './text';

/** A row of cells from values, each shown as its text. */
function row(values: Record<number, Value>) {
  return (col: number) => {
    const value = values[col] ?? null;
    const display =
      value === null
        ? ''
        : typeof value === 'object'
          ? value.error
          : String(value).toUpperCase();
    return { value, display };
  };
}

// Each character is a tenth of a column wide.
const measure = (text: string) => (text.length * COLUMN_WIDTH) / 10;

describe('rowTexts', () => {
  it('aligns numbers right, text left, and booleans and errors in the middle', () => {
    const texts = rowTexts(
      row({ 0: 1, 1: 'a', 2: true, 3: { error: '#VALUE!' } }),
      0,
      5,
      measure,
    );
    expect(texts).toEqual([
      { col: 0, text: '1', align: 'right', from: 0, to: 0 },
      { col: 1, text: 'A', align: 'left', from: 1, to: 1 },
      { col: 2, text: 'TRUE', align: 'center', from: 2, to: 2 },
      { col: 3, text: '#VALUE!', align: 'center', from: 3, to: 3 },
    ]);
  });

  it('lets long text run over the empty cells to its right, as far as it needs', () => {
    const long = 'x'.repeat(15); // One and a half columns.
    expect(rowTexts(row({ 1: long }), 0, 9, measure)).toEqual([
      { col: 1, text: long.toUpperCase(), align: 'left', from: 1, to: 2 },
    ]);
  });

  it('stops long text at the next cell that is not empty', () => {
    const long = 'x'.repeat(35);
    const texts = rowTexts(row({ 0: long, 2: 5 }), 0, 9, measure);
    expect(texts[0]).toMatchObject({ col: 0, from: 0, to: 1 });
  });

  it('does not let numbers run over', () => {
    const texts = rowTexts(row({ 0: 1234567890123456 }), 0, 9, measure);
    expect(texts[0]).toMatchObject({ from: 0, to: 0 });
  });

  it('includes long text from the left of the visible columns that reaches them', () => {
    const long = 'x'.repeat(35); // Three and a half columns.
    expect(rowTexts(row({ 1: long }), 3, 9, measure)).toEqual([
      { col: 1, text: long.toUpperCase(), align: 'left', from: 1, to: 4 },
    ]);
    expect(rowTexts(row({ 1: long }), 5, 9, measure)).toEqual([]);
    expect(rowTexts(row({ 0: long, 1: 'b' }), 3, 9, measure)).toEqual([]);
  });
});
