import { describe, expect, it } from 'vitest';
import { SHEET_COLUMNS, SHEET_ROWS } from '../../engine/workbook';
import { pick, single } from './selection';

const LAST_ROW = SHEET_ROWS - 1;
const LAST_COL = SHEET_COLUMNS - 1;
// The top-left cell on screen, where a whole row or column starts.
const origin = { row: 10, col: 3 };

describe('selection', () => {
  it('starts as one cell, which is active', () => {
    expect(single({ row: 2, col: 1 })).toEqual({
      ranges: [{ top: 2, left: 1, bottom: 2, right: 1 }],
      active: { row: 2, col: 1 },
      anchor: { row: 2, col: 1 },
      focus: { row: 2, col: 1 },
    });
  });

  it('picks a cell, a row, a column or the whole sheet', () => {
    const start = single({ row: 0, col: 0 });
    expect(
      pick(start, { kind: 'cell', row: 4, col: 5 }, 'replace', origin),
    ).toEqual(single({ row: 4, col: 5 }));

    const row = pick(start, { kind: 'row', row: 12 }, 'replace', origin);
    expect(row.ranges).toEqual([
      { top: 12, left: 0, bottom: 12, right: LAST_COL },
    ]);
    expect(row.active).toEqual({ row: 12, col: 3 });

    const column = pick(start, { kind: 'column', col: 7 }, 'replace', origin);
    expect(column.ranges).toEqual([
      { top: 0, left: 7, bottom: LAST_ROW, right: 7 },
    ]);
    expect(column.active).toEqual({ row: 10, col: 7 });

    const all = pick(start, { kind: 'all' }, 'replace', origin);
    expect(all.ranges).toEqual([
      { top: 0, left: 0, bottom: LAST_ROW, right: LAST_COL },
    ]);
    expect(all.active).toEqual(origin);
  });

  it('extends from the anchor, which stays active, in any direction', () => {
    const start = single({ row: 5, col: 5 });
    const down = pick(
      start,
      { kind: 'cell', row: 8, col: 6 },
      'extend',
      origin,
    );
    expect(down.ranges).toEqual([{ top: 5, left: 5, bottom: 8, right: 6 }]);
    expect(down.active).toEqual({ row: 5, col: 5 });

    const up = pick(down, { kind: 'cell', row: 1, col: 2 }, 'extend', origin);
    expect(up.ranges).toEqual([{ top: 1, left: 2, bottom: 5, right: 5 }]);
    expect(up.focus).toEqual({ row: 1, col: 2 });
  });

  it('extends a row selection over whole rows', () => {
    const row = pick(
      single(origin),
      { kind: 'row', row: 4 },
      'replace',
      origin,
    );
    expect(pick(row, { kind: 'row', row: 2 }, 'extend', origin).ranges).toEqual(
      [{ top: 2, left: 0, bottom: 4, right: LAST_COL }],
    );

    // From a cell, a row header extends over the rows between them.
    const fromCell = pick(
      single({ row: 1, col: 1 }),
      { kind: 'row', row: 3 },
      'extend',
      origin,
    );
    expect(fromCell.ranges).toEqual([
      { top: 1, left: 0, bottom: 3, right: LAST_COL },
    ]);
    expect(fromCell.active).toEqual({ row: 1, col: 1 });

    const columns = pick(
      single({ row: 1, col: 1 }),
      { kind: 'column', col: 4 },
      'extend',
      origin,
    );
    expect(columns.ranges).toEqual([
      { top: 0, left: 1, bottom: LAST_ROW, right: 4 },
    ]);
  });

  it('adds further ranges, the new one active and extended', () => {
    const first = pick(
      single({ row: 0, col: 0 }),
      { kind: 'cell', row: 1, col: 1 },
      'extend',
      origin,
    );
    const added = pick(first, { kind: 'cell', row: 5, col: 5 }, 'add', origin);
    expect(added.ranges).toEqual([
      { top: 0, left: 0, bottom: 1, right: 1 },
      { top: 5, left: 5, bottom: 5, right: 5 },
    ]);
    expect(added.active).toEqual({ row: 5, col: 5 });

    const extended = pick(
      added,
      { kind: 'cell', row: 6, col: 7 },
      'extend',
      origin,
    );
    expect(extended.ranges).toEqual([
      { top: 0, left: 0, bottom: 1, right: 1 },
      { top: 5, left: 5, bottom: 6, right: 7 },
    ]);

    const column = pick(extended, { kind: 'column', col: 9 }, 'add', origin);
    expect(column.ranges).toHaveLength(3);
    expect(column.ranges[2]).toEqual({
      top: 0,
      left: 9,
      bottom: LAST_ROW,
      right: 9,
    });
  });
});
