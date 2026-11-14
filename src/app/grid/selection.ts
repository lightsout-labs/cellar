import { SHEET_COLUMNS, SHEET_ROWS, type Range } from '../../engine/workbook';
import type { CellPosition, Target } from './geometry';

const LAST_ROW = SHEET_ROWS - 1;
const LAST_COL = SHEET_COLUMNS - 1;

/**
 * The selected cells: one or more ranges, and the active cell, which is in
 * the last. The last range spans from its anchor to its focus; extending
 * the selection moves the focus.
 */
export interface Selection {
  readonly ranges: readonly Range[];
  readonly active: CellPosition;
  readonly anchor: CellPosition;
  readonly focus: CellPosition;
}

/** The range between two corners. */
function span(a: CellPosition, b: CellPosition): Range {
  return {
    top: Math.min(a.row, b.row),
    left: Math.min(a.col, b.col),
    bottom: Math.max(a.row, b.row),
    right: Math.max(a.col, b.col),
  };
}

/** One cell, selected and active. */
export function single(cell: CellPosition): Selection {
  return {
    ranges: [span(cell, cell)],
    active: cell,
    anchor: cell,
    focus: cell,
  };
}

/**
 * The selection with its last range stretched from its anchor to a new
 * focus, the active cell unchanged: Shift with a key or a click.
 */
export function extendTo(selection: Selection, focus: CellPosition): Selection {
  const { active, anchor } = selection;
  return withLast(selection, { active, anchor, focus });
}

/**
 * A selection whose last range is replaced by the one from `anchor` to
 * `focus`, or added after the others.
 */
function withLast(
  selection: Selection,
  last: Pick<Selection, 'active' | 'anchor' | 'focus'>,
  add = false,
): Selection {
  const kept = add ? selection.ranges : selection.ranges.slice(0, -1);
  return { ...last, ranges: [...kept, span(last.anchor, last.focus)] };
}

/**
 * The selection after the user picks a target with the mouse: `replace`
 * selects it alone, `extend` stretches the last range to it (Shift), and
 * `add` adds it as a further range (Ctrl or Cmd). A whole row or column
 * becomes active at `origin`, the top-left cell on screen.
 */
export function pick(
  selection: Selection,
  target: Target,
  mode: 'replace' | 'extend' | 'add',
  origin: CellPosition,
): Selection {
  if (mode === 'extend') {
    const { anchor, active } = selection;
    switch (target.kind) {
      case 'cell':
        return extendTo(selection, { row: target.row, col: target.col });
      case 'row':
        return withLast(selection, {
          active,
          anchor: { row: anchor.row, col: 0 },
          focus: { row: target.row, col: LAST_COL },
        });
      case 'column':
        return withLast(selection, {
          active,
          anchor: { row: 0, col: anchor.col },
          focus: { row: LAST_ROW, col: target.col },
        });
    }
  }
  return withLast(selection, corners(target, origin), mode === 'add');
}

function corners(
  target: Target,
  origin: CellPosition,
): Pick<Selection, 'active' | 'anchor' | 'focus'> {
  switch (target.kind) {
    case 'cell': {
      const cell = { row: target.row, col: target.col };
      return { active: cell, anchor: cell, focus: cell };
    }
    case 'row':
      return {
        active: { row: target.row, col: origin.col },
        anchor: { row: target.row, col: 0 },
        focus: { row: target.row, col: LAST_COL },
      };
    case 'column':
      return {
        active: { row: origin.row, col: target.col },
        anchor: { row: 0, col: target.col },
        focus: { row: LAST_ROW, col: target.col },
      };
    case 'all':
      return {
        active: origin,
        anchor: { row: 0, col: 0 },
        focus: { row: LAST_ROW, col: LAST_COL },
      };
  }
}
