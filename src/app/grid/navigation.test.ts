import { describe, expect, it } from 'vitest';
import { SHEET_COLUMNS, SHEET_ROWS } from '../../engine/workbook';
import {
  navigate,
  navigationFor,
  type KeyPress,
  type Navigation,
  type SheetContents,
} from './navigation';
import { single, type Selection } from './selection';

const LAST_ROW = SHEET_ROWS - 1;
const LAST_COL = SHEET_COLUMNS - 1;

/** A sheet whose filled cells are given as `row,col`. */
function contents(...filled: string[]): SheetContents {
  const cells = new Set(filled);
  return {
    isEmpty: (row, col) => !cells.has(`${row},${col}`),
    extent: () => {
      let rows = 0;
      let cols = 0;
      for (const each of cells) {
        const [row, col] = each.split(',').map(Number);
        rows = Math.max(rows, row! + 1);
        cols = Math.max(cols, col! + 1);
      }
      return { rows, cols };
    },
  };
}

const empty = contents();
const at = (row: number, col: number) => single({ row, col });
const step = (rows: number, cols: number, extend = false): Navigation => ({
  type: 'step',
  rows,
  cols,
  extend,
});
const edge = (rows: number, cols: number, extend = false): Navigation => ({
  type: 'edge',
  rows,
  cols,
  extend,
});
const activeAfter = (
  selection: Selection,
  navigation: Navigation,
  sheet = empty,
) => navigate(selection, navigation, sheet).active;

describe('keyboard navigation', () => {
  it('steps one cell, and stops at the edges of the sheet', () => {
    expect(activeAfter(at(5, 5), step(1, 0))).toEqual({ row: 6, col: 5 });
    expect(activeAfter(at(5, 5), step(0, -1))).toEqual({ row: 5, col: 4 });
    expect(activeAfter(at(0, 0), step(-1, -1))).toEqual({ row: 0, col: 0 });
    expect(activeAfter(at(LAST_ROW, LAST_COL), step(1, 1))).toEqual({
      row: LAST_ROW,
      col: LAST_COL,
    });
    // A page step stops at the edge too.
    expect(activeAfter(at(3, 0), step(-20, 0))).toEqual({ row: 0, col: 0 });
  });

  it('collapses a range to one cell when it moves', () => {
    const range = navigate(at(1, 1), step(2, 2, true), empty);
    expect(navigate(range, step(1, 0), empty)).toEqual(at(2, 1));
  });

  it('extends from the moving corner, keeping the active cell', () => {
    const once = navigate(at(5, 5), step(1, 0, true), empty);
    const twice = navigate(once, step(0, -2, true), empty);
    expect(twice.ranges).toEqual([{ top: 5, left: 3, bottom: 6, right: 5 }]);
    expect(twice.active).toEqual({ row: 5, col: 5 });
    const back = navigate(twice, step(-1, 2, true), empty);
    expect(back.ranges).toEqual([{ top: 5, left: 5, bottom: 5, right: 5 }]);
  });

  it('jumps to the edges of the data as Excel does', () => {
    // Column A holds rows 3 to 6, and row 10.
    const sheet = contents('2,0', '3,0', '4,0', '5,0', '9,0');
    // From an empty cell, to the next filled one.
    expect(activeAfter(at(0, 0), edge(1, 0), sheet)).toEqual({
      row: 2,
      col: 0,
    });
    // Inside a block, to its last cell.
    expect(activeAfter(at(2, 0), edge(1, 0), sheet)).toEqual({
      row: 5,
      col: 0,
    });
    // At the end of a block, to the start of the next one.
    expect(activeAfter(at(5, 0), edge(1, 0), sheet)).toEqual({
      row: 9,
      col: 0,
    });
    // After the last one, to the edge of the sheet.
    expect(activeAfter(at(9, 0), edge(1, 0), sheet)).toEqual({
      row: LAST_ROW,
      col: 0,
    });
    expect(activeAfter(at(LAST_ROW, 0), edge(-1, 0), sheet)).toEqual({
      row: 9,
      col: 0,
    });
    expect(activeAfter(at(4, 0), edge(0, 1), sheet)).toEqual({
      row: 4,
      col: LAST_COL,
    });
    expect(activeAfter(at(4, 7), edge(0, -1), sheet)).toEqual({
      row: 4,
      col: 0,
    });
    expect(activeAfter(at(0, 0), edge(-1, 0), sheet)).toEqual({
      row: 0,
      col: 0,
    });
  });

  it('extends to the edge of the data from the moving corner', () => {
    const sheet = contents('0,0', '1,0', '2,0', '0,1', '0,2');
    const down = navigate(at(0, 0), edge(1, 0, true), sheet);
    const right = navigate(down, edge(0, 1, true), sheet);
    expect(right.ranges).toEqual([
      { top: 0, left: 0, bottom: 2, right: LAST_COL },
    ]);
    expect(right.active).toEqual({ row: 0, col: 0 });
  });

  it('goes to the start of the row, the start of the sheet and the last cell', () => {
    const sheet = contents('3,8', '12,2');
    const at47 = at(4, 7);
    expect(activeAfter(at47, { type: 'rowStart', extend: false })).toEqual({
      row: 4,
      col: 0,
    });
    expect(activeAfter(at47, { type: 'sheetStart', extend: false })).toEqual({
      row: 0,
      col: 0,
    });
    expect(
      activeAfter(at47, { type: 'lastCell', extend: false }, sheet),
    ).toEqual({
      row: 12,
      col: 8,
    });
    expect(activeAfter(at47, { type: 'lastCell', extend: false })).toEqual({
      row: 0,
      col: 0,
    });
    expect(
      navigate(at47, { type: 'rowStart', extend: true }, sheet).ranges,
    ).toEqual([{ top: 4, left: 0, bottom: 4, right: 7 }]);
    expect(
      navigate(at47, { type: 'sheetStart', extend: true }, sheet).ranges,
    ).toEqual([{ top: 0, left: 0, bottom: 4, right: 7 }]);
  });

  it('selects the data around the active cell, then the whole sheet', () => {
    // A block from B2 to D4 with a gap, touching E5 at a corner.
    const sheet = contents('1,1', '1,2', '1,3', '2,1', '3,2', '3,3', '4,4');
    const region = navigate(at(2, 1), { type: 'all' }, sheet);
    expect(region.ranges).toEqual([{ top: 1, left: 1, bottom: 4, right: 4 }]);
    expect(region.active).toEqual({ row: 2, col: 1 });

    const all = navigate(region, { type: 'all' }, sheet);
    expect(all.ranges).toEqual([
      { top: 0, left: 0, bottom: LAST_ROW, right: LAST_COL },
    ]);
    expect(all.active).toEqual({ row: 2, col: 1 });

    // Away from any data, the first press selects the whole sheet.
    expect(navigate(at(20, 20), { type: 'all' }, sheet).ranges).toEqual(
      all.ranges,
    );
  });
});

describe('navigation keys', () => {
  const press = (key: string, modifiers: Partial<KeyPress> = {}): KeyPress => ({
    key,
    shift: false,
    ctrl: false,
    meta: false,
    alt: false,
    ...modifiers,
  });
  const windows = { mac: false, endMode: false, page: { rows: 20, cols: 6 } };
  const mac = { ...windows, mac: true };

  it('maps arrows, Shift and the page keys to steps', () => {
    expect(navigationFor(press('ArrowDown'), windows)).toEqual(step(1, 0));
    expect(navigationFor(press('ArrowLeft', { shift: true }), windows)).toEqual(
      step(0, -1, true),
    );
    expect(navigationFor(press('PageDown'), windows)).toEqual(step(20, 0));
    expect(navigationFor(press('PageUp', { shift: true }), windows)).toEqual(
      step(-20, 0, true),
    );
    expect(navigationFor(press('PageDown', { alt: true }), windows)).toEqual(
      step(0, 6),
    );
    expect(navigationFor(press('PageUp', { alt: true }), windows)).toEqual(
      step(0, -6),
    );
  });

  it('uses Ctrl on Windows and Linux and Cmd on macOS', () => {
    expect(navigationFor(press('ArrowUp', { ctrl: true }), windows)).toEqual(
      edge(-1, 0),
    );
    expect(navigationFor(press('ArrowUp', { meta: true }), mac)).toEqual(
      edge(-1, 0),
    );
    expect(
      navigationFor(press('ArrowRight', { meta: true, shift: true }), mac),
    ).toEqual(edge(0, 1, true));
    expect(navigationFor(press('a', { ctrl: true }), windows)).toEqual({
      type: 'all',
    });
    expect(navigationFor(press('A', { meta: true, shift: true }), mac)).toEqual(
      { type: 'all' },
    );
    expect(navigationFor(press('a', { meta: true }), mac)).toEqual({
      type: 'all',
    });
    expect(navigationFor(press('a', { ctrl: true }), mac)).toBeNull();
    expect(navigationFor(press('Home', { ctrl: true }), windows)).toEqual({
      type: 'sheetStart',
      extend: false,
    });
    expect(
      navigationFor(press('End', { ctrl: true, shift: true }), windows),
    ).toEqual({
      type: 'lastCell',
      extend: true,
    });
  });

  it('maps Home to the start of the row, and End to End mode', () => {
    expect(navigationFor(press('Home'), windows)).toEqual({
      type: 'rowStart',
      extend: false,
    });
    expect(navigationFor(press('End'), windows)).toEqual({ type: 'endMode' });
    // In End mode, an arrow jumps to the edge of the data, and Home goes to the last cell.
    const endMode = { ...windows, endMode: true };
    expect(navigationFor(press('ArrowDown'), endMode)).toEqual(edge(1, 0));
    expect(navigationFor(press('ArrowDown', { shift: true }), endMode)).toEqual(
      edge(1, 0, true),
    );
    expect(navigationFor(press('Home'), endMode)).toEqual({
      type: 'lastCell',
      extend: false,
    });
    expect(navigationFor(press('End'), endMode)).toEqual({ type: 'endMode' });
  });

  it('leaves every other key alone', () => {
    expect(navigationFor(press('x'), windows)).toBeNull();
    expect(navigationFor(press('Enter'), windows)).toBeNull();
    expect(navigationFor(press('z', { ctrl: true }), windows)).toBeNull();
  });
});
