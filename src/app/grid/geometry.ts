/** Sizes of the grid, in CSS pixels. */
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

/** Where a cell is drawn on the canvas. */
export function cellRect({ row, col }: CellPosition): Rect {
  return {
    x: ROW_HEADER_WIDTH + col * COLUMN_WIDTH,
    y: COLUMN_HEADER_HEIGHT + row * ROW_HEIGHT,
    width: COLUMN_WIDTH,
    height: ROW_HEIGHT,
  };
}

/** The cell under a point of the canvas, or `null` under a header. */
export function cellAt(x: number, y: number): CellPosition | null {
  if (x < ROW_HEADER_WIDTH || y < COLUMN_HEADER_HEIGHT) return null;
  return {
    row: Math.floor((y - COLUMN_HEADER_HEIGHT) / ROW_HEIGHT),
    col: Math.floor((x - ROW_HEADER_WIDTH) / COLUMN_WIDTH),
  };
}

/** How many rows and columns a canvas of this size shows, partial ones included. */
export function visibleCells(
  width: number,
  height: number,
): { rows: number; cols: number } {
  return {
    rows: Math.max(0, Math.ceil((height - COLUMN_HEADER_HEIGHT) / ROW_HEIGHT)),
    cols: Math.max(0, Math.ceil((width - ROW_HEADER_WIDTH) / COLUMN_WIDTH)),
  };
}
