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
      sheet,
      caret: 2,
      allFunctions: false,
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
      sheet,
      caret: 3,
      allFunctions: false,
    });
    view.type('15%');
    view.commit();
    expect(workbook.cell(sheet, 0, 0).display).toBe('15%');

    view.edit('formulaBar');
    expect(view.state.entry).toEqual({
      text: '15%',
      mode: 'edit',
      place: 'formulaBar',
      sheet,
      caret: 3,
      allFunctions: false,
    });
  });

  it('keeps the text typed so far when an entry switches to Edit mode', () => {
    const { view, sheet } = setup();
    view.type('ab');
    view.edit('cell');
    expect(view.state.entry).toEqual({
      text: 'ab',
      mode: 'edit',
      place: 'cell',
      sheet,
      caret: 2,
      allFunctions: false,
    });
    view.edit('formulaBar');
    expect(view.state.entry?.place).toBe('formulaBar');
    view.type('abc');
    expect(view.state.entry).toEqual({
      text: 'abc',
      mode: 'edit',
      place: 'formulaBar',
      sheet,
      caret: 3,
      allFunctions: false,
    });
  });

  it('keeps the caret of the entry, at the end of what is typed unless moved', () => {
    const { view } = setup();
    view.type('=SUM()');
    expect(view.state.entry?.caret).toBe(6);
    view.moveCaret(5);
    expect(view.state.entry?.caret).toBe(5);
    view.type('=SUM(1)', 6);
    expect(view.state.entry?.caret).toBe(6);
  });

  it('inserts references picked on the grid into a formula', () => {
    const { workbook, view, sheet } = setup();
    view.select(single({ row: 4, col: 0 }));
    expect(view.canPick).toBe(false);
    view.type('=SUM(');
    expect(view.canPick).toBe(true);
    view.pick({ top: 0, left: 0, bottom: 0, right: 0 });
    // Dragging replaces the reference picked.
    view.pick({ top: 0, left: 0, bottom: 2, right: 1 });
    expect(view.state.entry).toMatchObject({ text: '=SUM(A1:B3', caret: 10 });
    expect(view.state.selection).toEqual(single({ row: 4, col: 0 }));
    view.type('=SUM(A1:B3)');
    expect(view.canPick).toBe(false);
    view.commit();
    expect(workbook.cell(sheet, 4, 0).input).toBe('=SUM(A1:B3)');
  });

  it('goes on with a formula on another sheet, and commits it to its own cell', () => {
    const { workbook, view, sheet } = setup();
    workbook.execute({ type: 'addSheet', index: 1 });
    const second = workbook.sheets[1]!.id;
    workbook.execute({ type: 'setSheet', sheet: second, name: 'My data' });
    view.select(single({ row: 1, col: 1 }));
    view.type('=1+');
    view.showSheet(second);
    expect(view.state.sheet).toBe(second);
    // The cell is not shown, so the entry goes on in the formula bar.
    expect(view.state.entry).toMatchObject({ place: 'formulaBar', sheet });
    expect(view.canPick).toBe(true);
    view.pick({ top: 2, left: 2, bottom: 2, right: 2 });
    expect(view.state.entry?.text).toBe("=1+'My data'!C3");
    view.navigate(down);
    expect(view.state.sheet).toBe(sheet);
    expect(view.state.selection).toEqual(single({ row: 2, col: 1 }));
    expect(workbook.cell(sheet, 1, 1).input).toBe("=1+'My data'!C3");

    // Back on its own sheet, a reference has no sheet name; cancelling
    // also goes back.
    view.type('=');
    view.showSheet(second);
    view.showSheet(sheet);
    view.pick({ top: 0, left: 0, bottom: 0, right: 0 });
    expect(view.state.entry?.text).toBe('=A1');
    view.showSheet(second);
    view.cancel();
    expect(view.state).toMatchObject({ sheet, entry: null });

    // Where no reference can go, a click on the other sheet commits the
    // formula to its own cell and selects there.
    view.select(single({ row: 1, col: 1 }));
    view.type('=SUM(');
    view.showSheet(second);
    view.pick({ top: 0, left: 0, bottom: 0, right: 0 });
    view.type("=SUM('My data'!A1)");
    expect(view.canPick).toBe(false);
    view.select(single({ row: 3, col: 3 }));
    expect(view.state).toMatchObject({
      sheet: second,
      entry: null,
      selection: single({ row: 3, col: 3 }),
    });
    expect(workbook.cell(sheet, 1, 1).input).toBe("=SUM('My data'!A1)");
  });

  it('cycles the reference at the caret between absolute and relative', () => {
    const { view } = setup();
    view.cycleReference(); // Nothing to cycle.
    expect(view.state.entry).toBeNull();
    view.type('=A1+2', 3);
    view.cycleReference();
    expect(view.state.entry).toMatchObject({ text: '=$A$1+2', caret: 5 });
  });

  it('completes the function name typed', () => {
    const { view } = setup();
    view.type('=1+ro');
    view.complete('ROUND');
    expect(view.state.entry).toMatchObject({ text: '=1+ROUND(', caret: 9 });
  });

  it('opens the list of every function, in a formula or a new one', () => {
    const { workbook, view, sheet } = setup();
    workbook.execute({ type: 'setCell', sheet, row: 0, col: 0, input: 'x' });
    view.insertFunction();
    expect(view.state.entry).toMatchObject({
      text: '=',
      caret: 1,
      mode: 'edit',
      place: 'formulaBar',
      allFunctions: true,
    });
    view.complete('ABS');
    expect(view.state.entry).toMatchObject({
      text: '=ABS(',
      allFunctions: false,
    });
    view.type('=ABS(1)+', 8);
    view.insertFunction();
    expect(view.state.entry).toMatchObject({ text: '=ABS(1)+', caret: 8 });
    // Moving the caret or typing closes the list.
    view.moveCaret(2);
    expect(view.state.entry?.allFunctions).toBe(false);
    view.insertFunction();
    view.type('=ABS(1)+S');
    expect(view.state.entry?.allFunctions).toBe(false);
    view.cancel();
    view.type('abc');
    view.insertFunction();
    expect(view.state.entry).toMatchObject({ text: '=', mode: 'enter' });
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

  it('undoes and redoes a change, selecting the cells it changed', () => {
    const { workbook, view, sheet } = setup();
    view.select(single({ row: 2, col: 1 }));
    view.type('4');
    view.commit();
    view.select(single({ row: 9, col: 9 }));
    view.undo();
    expect(workbook.cell(sheet, 2, 1).value).toBeNull();
    expect(view.state.selection).toEqual(single({ row: 2, col: 1 }));
    view.select(single({ row: 9, col: 9 }));
    view.redo();
    expect(workbook.cell(sheet, 2, 1).value).toBe(4);
    expect(view.state.selection).toEqual(single({ row: 2, col: 1 }));
  });

  it('selects every range an undone change cleared, the first cell of the last active', () => {
    const { view } = setup();
    const ranges = [
      { top: 0, left: 0, bottom: 1, right: 1 },
      { top: 3, left: 3, bottom: 5, right: 4 },
    ];
    view.select({
      ranges,
      active: { row: 0, col: 0 },
      anchor: { row: 3, col: 3 },
      focus: { row: 5, col: 4 },
    });
    view.clear();
    view.select(single({ row: 9, col: 9 }));
    view.undo();
    expect(view.state.selection).toEqual({
      ranges,
      active: { row: 3, col: 3 },
      anchor: { row: 3, col: 3 },
      focus: { row: 5, col: 4 },
    });
  });

  it('abandons an entry to undo, and keeps the selection when there is nothing to undo or redo', () => {
    const { workbook, view, sheet } = setup();
    view.select(single({ row: 1, col: 1 }));
    view.type('7');
    view.undo();
    view.redo();
    expect(view.state.entry).toBeNull();
    expect(view.state.selection).toEqual(single({ row: 1, col: 1 }));
    expect(workbook.cell(sheet, 1, 1).value).toBeNull();
  });

  it('shows another sheet at A1, committing the entry first', () => {
    const { workbook, view, sheet } = setup();
    workbook.execute({ type: 'addSheet', index: 1 });
    const second = workbook.sheets[1]!.id;
    view.select(single({ row: 3, col: 3 }));
    view.type('5');
    view.showSheet(second);
    expect(workbook.cell(sheet, 3, 3).value).toBe(5);
    expect(view.state.sheet).toBe(second);
    expect(view.state.selection).toEqual(single({ row: 0, col: 0 }));
  });

  it('shows a sheet it adds, and the neighbor of the shown sheet it deletes', () => {
    const { workbook, view, sheet } = setup();
    view.execute({ type: 'addSheet', index: 1 });
    const second = workbook.sheets[1]!.id;
    expect(view.state.sheet).toBe(second);
    // A change to another sheet leaves the shown one.
    view.select(single({ row: 2, col: 2 }));
    view.execute({ type: 'setSheet', sheet, name: 'First' });
    view.execute({ type: 'setSheet', sheet: second, index: 0 });
    expect(view.state.sheet).toBe(second);
    expect(view.state.selection).toEqual(single({ row: 2, col: 2 }));
    view.execute({ type: 'deleteSheet', sheet: second });
    expect(view.state.sheet).toBe(sheet);
  });

  it('shows the sheet whose change it undoes or redoes', () => {
    const { workbook, view, sheet } = setup();
    // Whatever reads the shown sheet on a change, as the grid does, finds it.
    workbook.subscribe(() => workbook.cell(view.state.sheet, 0, 0));
    view.execute({ type: 'addSheet', index: 1 });
    const second = workbook.sheets[1]!.id;
    view.select(single({ row: 4, col: 4 }));
    view.undo();
    expect(view.state.sheet).toBe(sheet);
    expect(view.state.selection).toEqual(single({ row: 0, col: 0 }));
    view.redo();
    expect(view.state.sheet).toBe(second);
    // Undoing the addition of a sheet that is not shown stays where it is.
    view.showSheet(sheet);
    view.undo();
    expect(view.state.sheet).toBe(sheet);
    view.redo();
    view.showSheet(second);
    view.select(single({ row: 4, col: 4 }));
    view.execute({ type: 'setSheet', sheet: second, color: '#30a46c' });
    view.undo();
    // The same sheet: the selection stays.
    expect(view.state.selection).toEqual(single({ row: 4, col: 4 }));
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
