import { describe, expect, it } from 'vitest';
import { SHEET_COLUMNS, SHEET_ROWS } from '../../engine/workbook';
import {
  cellRect,
  COLUMN_HEADER_HEIGHT,
  COLUMN_WIDTH,
  hitTest,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
  scrollToShow,
  sheetSize,
  visibleRange,
  type Viewport,
} from './geometry';

const top: Viewport = { scrollX: 0, scrollY: 0, zoom: 1 };

describe('grid geometry', () => {
  it('places A1 after the headers and other cells by their size', () => {
    expect(cellRect({ row: 0, col: 0 }, top)).toEqual({
      x: ROW_HEADER_WIDTH,
      y: COLUMN_HEADER_HEIGHT,
      width: COLUMN_WIDTH,
      height: ROW_HEIGHT,
    });
    expect(cellRect({ row: 3, col: 2 }, top)).toMatchObject({
      x: ROW_HEADER_WIDTH + 2 * COLUMN_WIDTH,
      y: COLUMN_HEADER_HEIGHT + 3 * ROW_HEIGHT,
    });
  });

  it('moves the cells, not the headers, by the scroll position', () => {
    const scrolled = { scrollX: 10, scrollY: 2 * ROW_HEIGHT, zoom: 1 };
    expect(cellRect({ row: 2, col: 0 }, scrolled)).toMatchObject({
      x: ROW_HEADER_WIDTH - 10,
      y: COLUMN_HEADER_HEIGHT,
    });
    expect(
      hitTest(ROW_HEADER_WIDTH + 1, COLUMN_HEADER_HEIGHT + 1, scrolled),
    ).toEqual({ kind: 'cell', row: 2, col: 0 });
  });

  it('scales everything, headers included, by the zoom', () => {
    const zoomed = { scrollX: 0, scrollY: 0, zoom: 2 };
    expect(cellRect({ row: 1, col: 1 }, zoomed)).toEqual({
      x: 2 * (ROW_HEADER_WIDTH + COLUMN_WIDTH),
      y: 2 * (COLUMN_HEADER_HEIGHT + ROW_HEIGHT),
      width: 2 * COLUMN_WIDTH,
      height: 2 * ROW_HEIGHT,
    });
    expect(sheetSize(0.5)).toEqual({
      width: (ROW_HEADER_WIDTH + SHEET_COLUMNS * COLUMN_WIDTH) / 2,
      height: (COLUMN_HEADER_HEIGHT + SHEET_ROWS * ROW_HEIGHT) / 2,
    });
  });

  it('finds the cell or header under a point', () => {
    const rect = cellRect({ row: 3, col: 2 }, top);
    expect(hitTest(rect.x, rect.y, top)).toEqual({
      kind: 'cell',
      row: 3,
      col: 2,
    });
    expect(
      hitTest(rect.x + rect.width - 1, rect.y + rect.height - 1, top),
    ).toEqual({ kind: 'cell', row: 3, col: 2 });
    expect(hitTest(ROW_HEADER_WIDTH - 1, rect.y, top)).toEqual({
      kind: 'row',
      row: 3,
    });
    expect(hitTest(rect.x, COLUMN_HEADER_HEIGHT - 1, top)).toEqual({
      kind: 'column',
      col: 2,
    });
    expect(hitTest(1, 1, top)).toEqual({ kind: 'all' });
  });

  it('finds the nearest cell for a point beyond the sheet', () => {
    const end = { scrollX: 1e9, scrollY: 1e9, zoom: 1 };
    expect(hitTest(5000, 5000, end)).toEqual({
      kind: 'cell',
      row: SHEET_ROWS - 1,
      col: SHEET_COLUMNS - 1,
    });
    expect(hitTest(-5, -5, top)).toEqual({ kind: 'all' });
  });

  it('gives the cells a canvas shows, including partial ones', () => {
    const width = ROW_HEADER_WIDTH + 2.5 * COLUMN_WIDTH;
    const height = COLUMN_HEADER_HEIGHT + 4 * ROW_HEIGHT;
    expect(visibleRange(width, height, top)).toEqual({
      top: 0,
      left: 0,
      bottom: 3,
      right: 2,
    });
    const scrolled = { scrollX: COLUMN_WIDTH * 0.75, scrollY: 0, zoom: 1 };
    expect(visibleRange(width, height, scrolled)).toMatchObject({
      left: 0,
      right: 3,
    });
    const end = { scrollX: 0, scrollY: SHEET_ROWS * ROW_HEIGHT, zoom: 1 };
    expect(visibleRange(width, height, end)).toMatchObject({
      top: SHEET_ROWS - 1,
      bottom: SHEET_ROWS - 1,
    });
  });

  it('scrolls as little as it can to show a cell', () => {
    const width = ROW_HEADER_WIDTH + 4 * COLUMN_WIDTH;
    const height = COLUMN_HEADER_HEIGHT + 10 * ROW_HEIGHT;
    // Already shown: no change.
    expect(scrollToShow({ row: 5, col: 2 }, top, width, height)).toEqual({
      scrollX: 0,
      scrollY: 0,
    });
    // Below and to the right: the cell ends at the far edges.
    expect(scrollToShow({ row: 20, col: 6 }, top, width, height)).toEqual({
      scrollX: 3 * COLUMN_WIDTH,
      scrollY: 11 * ROW_HEIGHT,
    });
    // Above and to the left: the cell starts just after the headers.
    const scrolled = { scrollX: 1000, scrollY: 1000, zoom: 1 };
    expect(scrollToShow({ row: 2, col: 1 }, scrolled, width, height)).toEqual({
      scrollX: COLUMN_WIDTH,
      scrollY: 2 * ROW_HEIGHT,
    });
    // At twice the size, the distances double.
    const zoomed = { scrollX: 0, scrollY: 0, zoom: 2 };
    expect(
      scrollToShow({ row: 20, col: 0 }, zoomed, width, height).scrollY,
    ).toBe(2 * (COLUMN_HEADER_HEIGHT + 21 * ROW_HEIGHT) - height);
  });
});
