import { describe, expect, it } from 'vitest';
import {
  cellAt,
  cellRect,
  COLUMN_HEADER_HEIGHT,
  COLUMN_WIDTH,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
  visibleCells,
} from './geometry';

describe('grid geometry', () => {
  it('places A1 after the headers and other cells by their size', () => {
    expect(cellRect({ row: 0, col: 0 })).toEqual({
      x: ROW_HEADER_WIDTH,
      y: COLUMN_HEADER_HEIGHT,
      width: COLUMN_WIDTH,
      height: ROW_HEIGHT,
    });
    expect(cellRect({ row: 3, col: 2 })).toMatchObject({
      x: ROW_HEADER_WIDTH + 2 * COLUMN_WIDTH,
      y: COLUMN_HEADER_HEIGHT + 3 * ROW_HEIGHT,
    });
  });

  it('finds the cell under a point', () => {
    const rect = cellRect({ row: 3, col: 2 });
    expect(cellAt(rect.x, rect.y)).toEqual({ row: 3, col: 2 });
    expect(cellAt(rect.x + rect.width - 1, rect.y + rect.height - 1)).toEqual({
      row: 3,
      col: 2,
    });
    expect(cellAt(rect.x + rect.width, rect.y)).toEqual({ row: 3, col: 3 });
  });

  it('finds no cell under the headers', () => {
    expect(cellAt(ROW_HEADER_WIDTH - 1, 100)).toBeNull();
    expect(cellAt(100, COLUMN_HEADER_HEIGHT - 1)).toBeNull();
  });

  it('counts the cells a canvas shows, including partial ones', () => {
    const width = ROW_HEADER_WIDTH + 2.5 * COLUMN_WIDTH;
    const height = COLUMN_HEADER_HEIGHT + 4 * ROW_HEIGHT;
    expect(visibleCells(width, height)).toEqual({ rows: 4, cols: 3 });
    expect(visibleCells(10, 10)).toEqual({ rows: 0, cols: 0 });
  });
});
