import { SHEET_COLUMNS, SHEET_ROWS, type Range } from '../../engine/workbook';
import type { CellPosition } from './geometry';
import { single, withLast, type Selection } from './selection';

const LAST_ROW = SHEET_ROWS - 1;
const LAST_COL = SHEET_COLUMNS - 1;

/**
 * A move of the selection by the keyboard. With `extend`, the focus of the
 * last range moves and the active cell stays; without it, the active cell
 * moves and is selected alone.
 */
export type Navigation =
  /** By a number of rows and columns: the arrows and the page keys. */
  | { type: 'step'; rows: number; cols: number; extend: boolean }
  /** To the edge of the data in a direction: Ctrl or Cmd with an arrow. */
  | { type: 'edge'; rows: number; cols: number; extend: boolean }
  /** To column A: Home. */
  | { type: 'rowStart'; extend: boolean }
  /** To A1: Ctrl or Cmd with Home. */
  | { type: 'sheetStart'; extend: boolean }
  /** To the last row and column in use: Ctrl or Cmd with End. */
  | { type: 'lastCell'; extend: boolean }
  /** The data around the active cell, then the whole sheet: Ctrl or Cmd with A. */
  | { type: 'all' }
  /** Turns End mode on or off: End. */
  | { type: 'endMode' };

/** A key pressed, with its modifiers. */
export interface KeyPress {
  readonly key: string;
  readonly shift: boolean;
  readonly ctrl: boolean;
  readonly meta: boolean;
  readonly alt: boolean;
}

/** What navigation needs to know of the sheet. */
export interface SheetContents {
  isEmpty(row: number, col: number): boolean;
  /** How many rows and columns, from A1, the cells in use extend over. */
  extent(): { rows: number; cols: number };
}

const ARROWS: Record<string, [number, number]> = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
};

/**
 * The navigation that a key stands for, as in Excel, or `null`. The
 * shortcut key is Cmd on macOS and Ctrl elsewhere. In End mode, an arrow
 * goes to the edge of the data and Home to the last cell.
 */
export function navigationFor(
  press: KeyPress,
  {
    mac,
    endMode,
    page,
  }: { mac: boolean; endMode: boolean; page: { rows: number; cols: number } },
): Navigation | null {
  const { key, shift: extend, alt } = press;
  const shortcut = mac ? press.meta : press.ctrl;
  const arrow = ARROWS[key];
  if (arrow) {
    const [rows, cols] = arrow;
    return { type: shortcut || endMode ? 'edge' : 'step', rows, cols, extend };
  }
  switch (key) {
    case 'PageDown':
    case 'PageUp': {
      const sign = key === 'PageDown' ? 1 : -1;
      return alt
        ? { type: 'step', rows: 0, cols: sign * page.cols, extend }
        : { type: 'step', rows: sign * page.rows, cols: 0, extend };
    }
    case 'Home':
      if (shortcut) return { type: 'sheetStart', extend };
      return { type: endMode ? 'lastCell' : 'rowStart', extend };
    case 'End':
      return shortcut ? { type: 'lastCell', extend } : { type: 'endMode' };
    case 'a':
    case 'A':
      return shortcut ? { type: 'all' } : null;
  }
  return null;
}

const clamp = ({ row, col }: CellPosition): CellPosition => ({
  row: Math.min(LAST_ROW, Math.max(0, row)),
  col: Math.min(LAST_COL, Math.max(0, col)),
});

/** The selection after a navigation over the sheet. */
export function navigate(
  selection: Selection,
  navigation: Navigation,
  sheet: SheetContents,
): Selection {
  if (navigation.type === 'endMode') return selection;
  if (navigation.type === 'all') return selectAll(selection, sheet);
  const { extend } = navigation;
  const from = extend ? selection.focus : selection.active;
  const to = clamp(target(from, navigation, sheet));
  if (!extend) return single(to);
  const { active, anchor } = selection;
  return withLast(selection, { active, anchor, focus: to });
}

function target(
  from: CellPosition,
  navigation: Exclude<Navigation, { type: 'all' | 'endMode' }>,
  sheet: SheetContents,
): CellPosition {
  switch (navigation.type) {
    case 'step':
      return {
        row: from.row + navigation.rows,
        col: from.col + navigation.cols,
      };
    case 'edge':
      return edge(from, navigation.rows, navigation.cols, sheet);
    case 'rowStart':
      return { row: from.row, col: 0 };
    case 'sheetStart':
      return { row: 0, col: 0 };
    case 'lastCell': {
      const { rows, cols } = sheet.extent();
      return { row: rows - 1, col: cols - 1 };
    }
  }
}

/**
 * Where Ctrl+arrow goes: from inside a block of filled cells, to its last
 * cell; otherwise, to the next filled cell, or the edge of the sheet.
 */
function edge(
  from: CellPosition,
  rows: number,
  cols: number,
  sheet: SheetContents,
): CellPosition {
  const inSheet = (row: number, col: number) =>
    row >= 0 && row <= LAST_ROW && col >= 0 && col <= LAST_COL;
  const filled = (row: number, col: number) => !sheet.isEmpty(row, col);
  let { row, col } = from;
  if (!inSheet(row + rows, col + cols)) return from;
  if (filled(row, col) && filled(row + rows, col + cols)) {
    while (inSheet(row + rows, col + cols) && filled(row + rows, col + cols)) {
      row += rows;
      col += cols;
    }
    return { row, col };
  }
  do {
    row += rows;
    col += cols;
  } while (inSheet(row + rows, col + cols) && !filled(row, col));
  return { row, col };
}

/**
 * Ctrl+A: the block of filled cells around the active cell, found by
 * growing a range while a cell touching it, diagonals included, is
 * filled; once that is selected, or with no data around, the whole sheet.
 */
function selectAll(selection: Selection, sheet: SheetContents): Selection {
  const { active } = selection;
  const region = regionAround(active, sheet);
  const last = selection.ranges.at(-1)!;
  const single = region.top === region.bottom && region.left === region.right;
  const already =
    selection.ranges.length === 1 &&
    last.top === region.top &&
    last.left === region.left &&
    last.bottom === region.bottom &&
    last.right === region.right;
  const range: Range =
    (single && sheet.isEmpty(active.row, active.col)) || already
      ? { top: 0, left: 0, bottom: LAST_ROW, right: LAST_COL }
      : region;
  return {
    ranges: [range],
    active,
    anchor: { row: range.top, col: range.left },
    focus: { row: range.bottom, col: range.right },
  };
}

function regionAround(cell: CellPosition, sheet: SheetContents): Range {
  let { row: top, col: left } = cell;
  let bottom = top;
  let right = left;
  const anyFilled = (r1: number, c1: number, r2: number, c2: number) => {
    for (let row = Math.max(0, r1); row <= Math.min(LAST_ROW, r2); row++) {
      for (let col = Math.max(0, c1); col <= Math.min(LAST_COL, c2); col++) {
        if (!sheet.isEmpty(row, col)) return true;
      }
    }
    return false;
  };
  // Each side grows as far as it can before the next, so that a block of
  // many rows is scanned a few times rather than once for each row.
  for (let grown = true; grown;) {
    grown = false;
    while (top > 0 && anyFilled(top - 1, left - 1, top - 1, right + 1)) {
      top--;
      grown = true;
    }
    while (
      bottom < LAST_ROW &&
      anyFilled(bottom + 1, left - 1, bottom + 1, right + 1)
    ) {
      bottom++;
      grown = true;
    }
    while (left > 0 && anyFilled(top - 1, left - 1, bottom + 1, left - 1)) {
      left--;
      grown = true;
    }
    while (
      right < LAST_COL &&
      anyFilled(top - 1, right + 1, bottom + 1, right + 1)
    ) {
      right++;
      grown = true;
    }
  }
  return { top, left, bottom, right };
}
