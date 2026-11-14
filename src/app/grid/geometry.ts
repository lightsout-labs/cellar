import { SHEET_COLUMNS, SHEET_ROWS, type Range } from '../../engine/workbook';

/** Sizes of the grid at 100%, in CSS pixels. */
export const ROW_HEADER_WIDTH = 46;
export const COLUMN_HEADER_HEIGHT = 24;
export const COLUMN_WIDTH = 96;
export const ROW_HEIGHT = 24;

/** A cell by its row and column, counted from zero. */
export interface CellPosition {
  readonly row: number;
  readonly col: number;
}

export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/**
 * Which part of the sheet the canvas shows: how far it is scrolled, in
 * CSS pixels at the zoom, and the zoom, where 1 is 100%.
 */
export interface Viewport {
  readonly scrollX: number;
  readonly scrollY: number;
  readonly zoom: number;
}

/** What is under a point of the canvas. */
export type Target =
  | { readonly kind: 'cell'; readonly row: number; readonly col: number }
  | { readonly kind: 'row'; readonly row: number }
  | { readonly kind: 'column'; readonly col: number }
  | { readonly kind: 'all' };

const clamp = (value: number, min: number, max: number) =>
  Math.min(max, Math.max(min, value));

/** The size of the whole sheet, headers included, at a zoom. */
export function sheetSize(zoom: number): { width: number; height: number } {
  return {
    width: (ROW_HEADER_WIDTH + SHEET_COLUMNS * COLUMN_WIDTH) * zoom,
    height: (COLUMN_HEADER_HEIGHT + SHEET_ROWS * ROW_HEIGHT) * zoom,
  };
}

/** Where a cell is drawn on the canvas. The headers stay in place. */
export function cellRect(
  { row, col }: CellPosition,
  { scrollX, scrollY, zoom }: Viewport,
): Rect {
  return {
    x: (ROW_HEADER_WIDTH + col * COLUMN_WIDTH) * zoom - scrollX,
    y: (COLUMN_HEADER_HEIGHT + row * ROW_HEIGHT) * zoom - scrollY,
    width: COLUMN_WIDTH * zoom,
    height: ROW_HEIGHT * zoom,
  };
}

/**
 * The cell or header under a point of the canvas. A point beyond the last
 * row or column is taken to be on it.
 */
export function hitTest(x: number, y: number, viewport: Viewport): Target {
  const { zoom } = viewport;
  const inRowHeader = x < ROW_HEADER_WIDTH * zoom;
  const inColumnHeader = y < COLUMN_HEADER_HEIGHT * zoom;
  const { row, col } = cellAt(x, y, viewport);
  if (inRowHeader && inColumnHeader) return { kind: 'all' };
  if (inRowHeader) return { kind: 'row', row };
  if (inColumnHeader) return { kind: 'column', col };
  return { kind: 'cell', row, col };
}

/**
 * The cell at a point as if the headers were not drawn over the cells: over
 * a header, the cell scrolled behind it. A drag past the headers or the
 * edges of the canvas reaches the cells beyond them this way.
 */
export function cellAt(x: number, y: number, viewport: Viewport): CellPosition {
  return { row: rowAt(y, viewport), col: colAt(x, viewport) };
}

function rowAt(y: number, { scrollY, zoom }: Viewport): number {
  const row = Math.floor(
    ((y + scrollY) / zoom - COLUMN_HEADER_HEIGHT) / ROW_HEIGHT,
  );
  return clamp(row, 0, SHEET_ROWS - 1);
}

function colAt(x: number, { scrollX, zoom }: Viewport): number {
  const col = Math.floor(
    ((x + scrollX) / zoom - ROW_HEADER_WIDTH) / COLUMN_WIDTH,
  );
  return clamp(col, 0, SHEET_COLUMNS - 1);
}

/** The cells that a canvas of this size shows, partial ones included. */
export function visibleRange(
  width: number,
  height: number,
  viewport: Viewport,
): Range {
  const { zoom } = viewport;
  return {
    top: rowAt(COLUMN_HEADER_HEIGHT * zoom, viewport),
    left: colAt(ROW_HEADER_WIDTH * zoom, viewport),
    bottom: rowAt(height - 1, viewport),
    right: colAt(width - 1, viewport),
  };
}

/**
 * The scroll position closest to the current one at which a canvas of
 * this size shows the whole of a cell, or as much of it as fits.
 */
export function scrollToShow(
  cell: CellPosition,
  viewport: Viewport,
  width: number,
  height: number,
): { scrollX: number; scrollY: number } {
  const { scrollX, scrollY, zoom } = viewport;
  const rect = cellRect(cell, viewport);
  const show = (start: number, size: number, from: number, to: number) => {
    if (start < from) return start - from;
    if (start + size > to) return Math.min(start + size - to, start - from);
    return 0;
  };
  return {
    scrollX: scrollX + show(rect.x, rect.width, ROW_HEADER_WIDTH * zoom, width),
    scrollY:
      scrollY + show(rect.y, rect.height, COLUMN_HEADER_HEIGHT * zoom, height),
  };
}
