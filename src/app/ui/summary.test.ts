import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../../engine/workbook';
import { pick, single } from '../grid/selection';
import { summarize } from './summary';

function setup(...inputs: [number, number, string][]) {
  const workbook = createWorkbook();
  const sheet = workbook.sheets[0]!.id;
  for (const [row, col, input] of inputs) {
    workbook.execute({ type: 'setCell', sheet, row, col, input });
  }
  const range = (top: number, left: number, bottom: number, right: number) =>
    pick(
      single({ row: top, col: left }),
      { kind: 'cell', row: bottom, col: right },
      'extend',
      {
        row: 0,
        col: 0,
      },
    );
  return {
    summary: (selection: ReturnType<typeof single>) =>
      summarize(workbook, sheet, selection),
    range,
  };
}

describe('the summary of the selection', () => {
  it('gives the average, count and sum of a range', () => {
    const { summary, range } = setup(
      [0, 0, '1'],
      [1, 0, '2'],
      [2, 0, 'x'],
      [3, 0, '=A1+A2'],
    );
    expect(summary(range(0, 0, 4, 0))).toEqual({
      average: '2',
      count: '4',
      sum: '6',
    });
    // Numbers are shown as the cells show them.
    const thirds = setup([0, 0, '1'], [1, 0, '1'], [2, 0, '0']);
    expect(thirds.summary(thirds.range(0, 0, 2, 0))?.average).toBe(
      '0.666666667',
    );
  });

  it('gives only the count when no cell holds a number', () => {
    const { summary, range } = setup([0, 0, 'a'], [0, 1, 'b']);
    expect(summary(range(0, 0, 0, 1))).toEqual({
      average: null,
      count: '2',
      sum: null,
    });
  });

  it('gives nothing for one cell or for empty cells', () => {
    const { summary, range } = setup([0, 0, '5']);
    expect(summary(single({ row: 0, col: 0 }))).toBeNull();
    expect(summary(range(3, 3, 9, 9))).toBeNull();
    expect(summary(range(0, 0, 1, 0))).toEqual({
      average: '5',
      count: '1',
      sum: '5',
    });
  });

  it('counts every range of the selection', () => {
    const { summary } = setup([0, 0, '1'], [5, 5, '2']);
    const both = pick(
      single({ row: 0, col: 0 }),
      { kind: 'cell', row: 5, col: 5 },
      'add',
      {
        row: 0,
        col: 0,
      },
    );
    expect(summary(both)).toEqual({ average: '1.5', count: '2', sum: '3' });
  });
});
