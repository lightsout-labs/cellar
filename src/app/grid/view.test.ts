import { describe, expect, it, vi } from 'vitest';
import { createWorkbook } from '../../engine/workbook';
import { single } from './selection';
import { View } from './view';

function setup() {
  const workbook = createWorkbook();
  const view = new View(workbook);
  const sheet = workbook.sheets[0]!.id;
  return { workbook, view, sheet };
}

const down = { type: 'step', rows: 1, cols: 0, extend: false } as const;

describe('View', () => {
  it('starts on A1 of the first sheet at 100%, with no entry', () => {
    const { view, sheet } = setup();
    expect(view.state).toEqual({
      sheet,
      selection: single({ row: 0, col: 0 }),
      entry: null,
      zoom: 1,
      endMode: false,
    });
  });

  it('commits the entry to the active cell', () => {
    const { workbook, view, sheet } = setup();
    view.select(single({ row: 2, col: 1 }));
    view.type('4');
    view.type('42');
    expect(view.state.entry).toEqual({
      text: '42',
      mode: 'enter',
      place: 'cell',
    });
    expect(workbook.cell(sheet, 2, 1).display).toBe('');

    view.commit();
    expect(view.state.entry).toBeNull();
    expect(workbook.cell(sheet, 2, 1).display).toBe('42');
  });

  it('commits a waiting entry to its own cell when the selection changes', () => {
    const { workbook, view, sheet } = setup();
    view.type('7');
    view.select(single({ row: 5, col: 5 }));
    expect(view.state).toMatchObject({
      selection: single({ row: 5, col: 5 }),
      entry: null,
    });
    expect(workbook.cell(sheet, 0, 0).display).toBe('7');

    view.type('8');
    view.navigate(down);
    expect(view.state.selection.active).toEqual({ row: 6, col: 5 });
    expect(workbook.cell(sheet, 5, 5).display).toBe('8');
  });

  it('navigates over the cells of the workbook', () => {
    const { workbook, view, sheet } = setup();
    for (const row of [0, 1, 2]) {
      workbook.execute({ type: 'setCell', sheet, row, col: 0, input: '1' });
    }
    view.navigate({ type: 'edge', rows: 1, cols: 0, extend: false });
    expect(view.state.selection.active).toEqual({ row: 2, col: 0 });
    view.navigate({ type: 'lastCell', extend: false });
    expect(view.state.selection.active).toEqual({ row: 2, col: 0 });
  });

  it('turns End mode on and off, and ends it with the next move', () => {
    const { view } = setup();
    view.navigate({ type: 'endMode' });
    expect(view.state.endMode).toBe(true);
    view.navigate({ type: 'endMode' });
    expect(view.state.endMode).toBe(false);
    view.navigate({ type: 'endMode' });
    view.navigate(down);
    expect(view.state.endMode).toBe(false);
    view.navigate({ type: 'endMode' });
    view.select(single({ row: 1, col: 1 }));
    expect(view.state.endMode).toBe(false);
  });

  it('edits the contents of the active cell, in the cell or the formula bar', () => {
    const { workbook, view, sheet } = setup();
    workbook.execute({ type: 'setCell', sheet, row: 0, col: 0, input: '12%' });
    view.edit('cell');
    expect(view.state.entry).toEqual({
      text: '12%',
      mode: 'edit',
      place: 'cell',
    });
    view.type('15%');
    view.commit();
    expect(workbook.cell(sheet, 0, 0).display).toBe('15%');

    view.edit('formulaBar');
    expect(view.state.entry).toEqual({
      text: '15%',
      mode: 'edit',
      place: 'formulaBar',
    });
  });

  it('keeps the text typed so far when an entry switches to Edit mode', () => {
    const { view } = setup();
    view.type('ab');
    view.edit('cell');
    expect(view.state.entry).toEqual({
      text: 'ab',
      mode: 'edit',
      place: 'cell',
    });
    view.edit('formulaBar');
    expect(view.state.entry?.place).toBe('formulaBar');
    view.type('abc');
    expect(view.state.entry).toEqual({
      text: 'abc',
      mode: 'edit',
      place: 'formulaBar',
    });
  });

  it('clears the selected cells', () => {
    const { workbook, view, sheet } = setup();
    for (const [row, col] of [
      [0, 0],
      [1, 1],
      [3, 3],
    ] as const) {
      workbook.execute({ type: 'setCell', sheet, row, col, input: 'x' });
    }
    view.select({
      ranges: [
        { top: 0, left: 0, bottom: 1, right: 1 },
        { top: 3, left: 3, bottom: 3, right: 3 },
      ],
      active: { row: 0, col: 0 },
      anchor: { row: 3, col: 3 },
      focus: { row: 3, col: 3 },
    });
    view.clear();
    expect(workbook.extent(sheet)).toEqual({ rows: 0, cols: 0 });
  });

  it('abandons the entry on cancel', () => {
    const { workbook, view, sheet } = setup();
    view.type('7');
    view.cancel();
    view.commit(); // Nothing is waiting any more.
    expect(view.state.entry).toBeNull();
    expect(workbook.cell(sheet, 0, 0).display).toBe('');
  });

  it('zooms between 50% and 200%', () => {
    const { view } = setup();
    view.zoomTo(1.5);
    expect(view.state.zoom).toBe(1.5);
    view.zoomTo(3);
    expect(view.state.zoom).toBe(2);
    view.zoomTo(0.1);
    expect(view.state.zoom).toBe(0.5);
    // Steps of a tenth add up without rounding errors.
    view.zoomTo(0.5 + 0.1 + 0.1);
    expect(view.state.zoom).toBe(0.7);
  });

  it('tells its listeners of each change, with a new state', () => {
    const { view } = setup();
    const before = view.state;
    const listener = vi.fn();
    const unsubscribe = view.subscribe(listener);
    view.select(single({ row: 1, col: 0 }));
    expect(listener).toHaveBeenCalledTimes(1);
    expect(view.state).not.toBe(before);
    expect(view.state).toBe(view.state);

    unsubscribe();
    view.type('1');
    expect(listener).toHaveBeenCalledTimes(1);
  });
});
