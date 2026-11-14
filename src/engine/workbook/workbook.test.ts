import { describe, expect, it, vi } from 'vitest';
import {
  cellName,
  createWorkbook,
  SHEET_COLUMNS,
  SHEET_ROWS,
  Workbook,
} from './index';

function setup() {
  const workbook = createWorkbook();
  const sheet = workbook.sheets[0]!.id;
  const set = (row: number, col: number, input: string) =>
    workbook.execute({ type: 'setCell', sheet, row, col, input });
  const show = (row: number, col: number) =>
    workbook.cell(sheet, row, col).display;
  return { workbook, sheet, set, show };
}

describe('workbook', () => {
  it('runs without a DOM', () => {
    expect(typeof document).toBe('undefined');
    expect(typeof window).toBe('undefined');
  });

  it('calculates a formula from another cell', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '2');
    set(1, 0, '=A1+1');
    expect(workbook.cell(sheet, 1, 0)).toEqual({
      value: 3,
      formula: '=A1+1',
      display: '3',
      input: '=A1+1',
    });
    expect(workbook.cell(sheet, 0, 0)).toEqual({
      value: 2,
      formula: null,
      display: '2',
      input: '2',
    });
  });

  it('recalculates a formula when a cell it reads changes', () => {
    const { set, show } = setup();
    set(1, 0, '=A1+1');
    set(2, 0, '=A2+1');
    expect(show(2, 0)).toBe('2');
    set(0, 0, '40');
    expect(show(2, 0)).toBe('42');
  });

  it('replaces a formula with a value, and a value with nothing', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '=B1+1');
    set(0, 0, 'text');
    expect(workbook.cell(sheet, 0, 0)).toEqual({
      value: 'text',
      formula: null,
      display: 'text',
      input: 'text',
    });
    set(0, 0, '');
    expect(workbook.cell(sheet, 0, 0).value).toBeNull();
  });

  it('keeps an entry that is not a valid formula as text', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '=A1+');
    expect(workbook.cell(sheet, 0, 0)).toEqual({
      value: '=A1+',
      formula: null,
      display: '=A1+',
      input: '=A1+',
    });
  });

  it('gives #VALUE! for formulas that read each other, without hanging', () => {
    const { set, show } = setup();
    set(0, 0, '=B1+1');
    set(0, 1, '=A1+1');
    expect(show(0, 0)).toBe('#VALUE!');
    set(0, 1, '1');
    expect(show(0, 0)).toBe('2');
  });

  it('reads a cell that was never set as empty', () => {
    const { workbook, sheet } = setup();
    expect(workbook.cell(sheet, 99_999, 99)).toEqual({
      value: null,
      formula: null,
      display: '',
      input: '',
    });
  });

  it('notifies subscribers of each change until they unsubscribe', () => {
    const { workbook, sheet, set } = setup();
    const listener = vi.fn();
    const unsubscribe = workbook.subscribe(listener);
    set(0, 0, '1');
    set(0, 0, '2');
    expect(listener).toHaveBeenCalledTimes(2);
    expect(listener).toHaveBeenLastCalledWith({
      type: 'setCell',
      sheet,
      row: 0,
      col: 0,
      input: '2',
    });
    unsubscribe();
    set(0, 0, '3');
    expect(listener).toHaveBeenCalledTimes(2);
  });

  it('restores values and formulas from a snapshot', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '2');
    set(1, 0, '=A1+1');
    set(4, 2, 'note');
    set(5, 0, '1/2/2026');
    set(6, 0, 'TRUE');
    const snapshot = JSON.parse(JSON.stringify(workbook.toSnapshot()));
    const restored = Workbook.fromSnapshot(snapshot);
    expect(restored.sheets).toEqual(workbook.sheets);
    expect(restored.cell(sheet, 0, 0).display).toBe('2');
    expect(restored.cell(sheet, 1, 0)).toEqual(workbook.cell(sheet, 1, 0));
    expect(restored.cell(sheet, 4, 2).display).toBe('note');
    expect(restored.cell(sheet, 5, 0)).toEqual(workbook.cell(sheet, 5, 0));
    expect(restored.cell(sheet, 6, 0).value).toBe(true);
    expect(restored.toSnapshot()).toEqual(snapshot);
  });

  it('refuses a snapshot of an unknown version', () => {
    const snapshot = { ...createWorkbook().toSnapshot(), version: 2 };
    expect(() => Workbook.fromSnapshot(snapshot as never)).toThrow(/version/);
  });

  it('refuses to restore a formula that does not parse', () => {
    const { workbook, sheet } = setup();
    const cell = { top: 0, left: 0, bottom: 0, right: 0 };
    expect(() =>
      workbook.execute({
        type: 'restore',
        sheet,
        ranges: [cell],
        cells: [{ row: 0, col: 0, value: null, formula: '=1+' }],
      }),
    ).toThrow(/=1\+/);
  });

  it('refuses a command for a sheet that does not exist', () => {
    const { workbook } = setup();
    expect(() =>
      workbook.execute({
        type: 'setCell',
        sheet: 'missing',
        row: 0,
        col: 0,
        input: '1',
      }),
    ).toThrow(/missing/);
  });

  it('names cells as Excel does', () => {
    expect(cellName(0, 0)).toBe('A1');
    expect(cellName(1, 25)).toBe('Z2');
    expect(cellName(9, 26)).toBe('AA10');
    expect(cellName(0, 701)).toBe('ZZ1');
    expect(cellName(0, 702)).toBe('AAA1');
  });

  it('totals the numbers of ranges and counts the cells that are not empty', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '1');
    set(1, 0, '2');
    set(2, 0, 'text');
    set(0, 1, '=A1+A2');
    set(5, 5, '100'); // Outside both ranges.
    const ranges = [
      { top: 0, left: 0, bottom: 3, right: 0 },
      { top: 0, left: 1, bottom: 0, right: 1 },
    ];
    expect(workbook.aggregate(sheet, ranges)).toEqual({
      count: 4,
      numbers: 3,
      sum: 6,
    });
    expect(
      workbook.aggregate(sheet, [{ top: 10, left: 10, bottom: 20, right: 20 }]),
    ).toEqual({ count: 0, numbers: 0, sum: 0 });
  });

  it('totals a whole sheet of 100,000 by 100 cells quickly', () => {
    const columns = Array.from({ length: SHEET_COLUMNS }, () =>
      Array.from({ length: SHEET_ROWS }, () => 1),
    );
    const workbook = Workbook.fromSnapshot({
      version: 1,
      name: 'Big',
      sheets: [{ id: 's', name: 'S', columns, formulas: [] }],
    });
    const whole = {
      top: 0,
      left: 0,
      bottom: SHEET_ROWS - 1,
      right: SHEET_COLUMNS - 1,
    };
    const start = performance.now();
    expect(workbook.aggregate('s', [whole]).sum).toBe(
      SHEET_ROWS * SHEET_COLUMNS,
    );
    expect(performance.now() - start).toBeLessThan(500);
  });

  it('knows how far the cells in use extend', () => {
    const { workbook, sheet, set } = setup();
    expect(workbook.extent(sheet)).toEqual({ rows: 0, cols: 0 });
    set(4, 1, '1');
    set(2, 6, '=B5+1');
    expect(workbook.extent(sheet)).toEqual({ rows: 5, cols: 7 });
    set(4, 1, ''); // An emptied cell is no longer in use.
    expect(workbook.extent(sheet)).toEqual({ rows: 3, cols: 7 });
  });

  it('stores typed numbers, dates, times, percentages and booleans as values with their format', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '1/2/2026');
    set(1, 0, '12%');
    set(2, 0, 'true');
    set(3, 0, '1,234.5');
    set(4, 0, '13:45');
    expect(workbook.cell(sheet, 0, 0)).toEqual({
      value: 46024,
      formula: null,
      display: '1/2/2026',
      input: '1/2/2026',
    });
    expect(workbook.cell(sheet, 1, 0)).toMatchObject({
      value: 0.12,
      display: '12%',
      input: '12%',
    });
    expect(workbook.cell(sheet, 2, 0)).toMatchObject({
      value: true,
      display: 'TRUE',
    });
    expect(workbook.cell(sheet, 3, 0)).toMatchObject({
      value: 1234.5,
      display: '1,234.50',
      input: '1234.5',
    });
    expect(workbook.cell(sheet, 4, 0).display).toBe('13:45');
  });

  it("keeps a cell's format when a plain number replaces its value, as Excel does", () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '1/2/2026');
    set(0, 0, '46025');
    expect(workbook.cell(sheet, 0, 0).display).toBe('1/3/2026');
    set(0, 0, '50%');
    expect(workbook.cell(sheet, 0, 0).display).toBe('50%');
  });

  it('shows the result of a formula in the format of its cell', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '1/2/2026');
    set(0, 0, '=1+46024');
    expect(workbook.cell(sheet, 0, 0)).toMatchObject({
      display: '1/3/2026',
      input: '=1+46024',
    });
  });

  it('adds booleans as numbers in formulas', () => {
    const { set, show } = setup();
    set(0, 0, 'TRUE');
    set(0, 1, '=A1+1');
    expect(show(0, 1)).toBe('2');
  });

  it('clears the values and formulas of ranges, and keeps their formats', () => {
    const { workbook, sheet, set, show } = setup();
    set(0, 0, '1');
    set(1, 1, '=A1+1');
    set(2, 0, '12%');
    set(5, 5, 'outside');
    const listener = vi.fn();
    workbook.subscribe(listener);
    const command = {
      type: 'clear',
      sheet,
      ranges: [
        { top: 0, left: 0, bottom: 1, right: 1 },
        { top: 2, left: 0, bottom: SHEET_ROWS - 1, right: 0 },
      ],
    } as const;
    workbook.execute(command);
    expect(listener).toHaveBeenCalledWith(command);
    expect(workbook.cell(sheet, 0, 0).value).toBeNull();
    expect(workbook.cell(sheet, 1, 1)).toMatchObject({
      value: null,
      formula: null,
    });
    expect(workbook.cell(sheet, 2, 0).value).toBeNull();
    expect(show(5, 5)).toBe('outside');
    expect(workbook.extent(sheet)).toEqual({ rows: 6, cols: 6 });
    set(2, 0, '0.5');
    expect(show(2, 0)).toBe('50%');
  });

  it('tells listeners the year it gave a date typed without one, which a replay keeps', () => {
    const { workbook, sheet } = setup();
    const listener = vi.fn();
    workbook.subscribe(listener);
    workbook.execute({
      type: 'setCell',
      sheet,
      row: 0,
      col: 0,
      input: '2-Jan',
    });
    const year = new Date().getFullYear();
    expect(listener).toHaveBeenCalledWith({
      type: 'setCell',
      sheet,
      row: 0,
      col: 0,
      input: '2-Jan',
      year,
    });

    const replay = createWorkbook();
    replay.execute({ ...listener.mock.calls[0]![0], year: 2030 });
    expect(replay.cell(sheet, 0, 0).input).toBe('1/2/2030');
  });
});

describe('undo and redo', () => {
  it('undoes a change to a value, a formula or a format, and redo it', () => {
    const { workbook, sheet, set, show } = setup();
    set(0, 0, '12%');
    set(0, 0, '=1+1');
    set(0, 0, '1/2/2026');
    workbook.undo();
    expect(workbook.cell(sheet, 0, 0)).toMatchObject({
      value: 2,
      formula: '=1+1',
      display: '200%',
    });
    workbook.undo();
    expect(show(0, 0)).toBe('12%');
    workbook.undo();
    expect(workbook.cell(sheet, 0, 0)).toMatchObject({
      value: null,
      formula: null,
    });
    set(0, 0, '0.5');
    expect(show(0, 0)).toBe('0.5'); // The format is undone too.
    workbook.undo();
    workbook.redo();
    workbook.redo(); // Nothing more to redo after a new change.
    expect(show(0, 0)).toBe('0.5');
  });

  it('undoes clearing ranges, and redoes it', () => {
    const { workbook, sheet, set, show } = setup();
    set(0, 0, '1');
    set(1, 1, '=A1+1');
    set(3, 0, '12%');
    set(5, 5, 'outside');
    workbook.execute({
      type: 'clear',
      sheet,
      ranges: [
        { top: 0, left: 0, bottom: 1, right: 1 },
        { top: 3, left: 0, bottom: SHEET_ROWS - 1, right: 0 },
      ],
    });
    workbook.undo();
    expect(show(0, 0)).toBe('1');
    expect(workbook.cell(sheet, 1, 1)).toMatchObject({
      formula: '=A1+1',
      display: '2',
    });
    expect(show(3, 0)).toBe('12%');
    workbook.redo();
    expect(workbook.extent(sheet)).toEqual({ rows: 6, cols: 6 });
    expect(workbook.cell(sheet, 1, 1).formula).toBeNull();
  });

  it('gives the cells that each undo and redo changed', () => {
    const { workbook, sheet, set } = setup();
    expect(workbook.undo()).toBeNull();
    set(2, 3, 'x');
    const ranges = [
      { top: 0, left: 0, bottom: 1, right: 1 },
      { top: 4, left: 4, bottom: 4, right: 4 },
    ];
    workbook.execute({ type: 'clear', sheet, ranges });
    expect(workbook.undo()).toEqual({ sheet, ranges });
    const cell = { sheet, ranges: [{ top: 2, left: 3, bottom: 2, right: 3 }] };
    expect(workbook.undo()).toEqual(cell);
    expect(workbook.redo()).toEqual(cell);
    expect(workbook.redo()).toEqual({ sheet, ranges });
    expect(workbook.redo()).toBeNull();
  });

  it('undoes the last 100 changes in order, and no more', () => {
    const { workbook, sheet, set, show } = setup();
    for (let step = 1; step <= 101; step++) set(0, 0, String(step));
    for (let step = 100; step >= 1; step--) {
      expect(workbook.canUndo).toBe(true);
      workbook.undo();
      expect(show(0, 0)).toBe(String(step));
    }
    expect(workbook.canUndo).toBe(false);
    expect(workbook.undo()).toBeNull();
    expect(workbook.cell(sheet, 0, 0).value).toBe(1);
  });

  it('tells whether there is something to undo and to redo, already to its listeners', () => {
    const { workbook, set } = setup();
    const seen: boolean[][] = [];
    workbook.subscribe(() => seen.push([workbook.canUndo, workbook.canRedo]));
    expect([workbook.canUndo, workbook.canRedo]).toEqual([false, false]);
    set(0, 0, '1');
    workbook.undo();
    workbook.redo();
    expect(seen).toEqual([
      [true, false],
      [false, true],
      [true, false],
    ]);
  });

  it('discards what could be redone when a new change follows an undo', () => {
    const { workbook, set, show } = setup();
    set(0, 0, '1');
    set(0, 0, '2');
    workbook.undo();
    set(1, 0, '3');
    expect(workbook.canRedo).toBe(false);
    expect(workbook.redo()).toBeNull();
    expect(show(0, 0)).toBe('1');
  });

  it('tells listeners of each undo and redo as a command that replays it', () => {
    const { workbook, sheet, set } = setup();
    const replay = createWorkbook();
    workbook.subscribe((command) => replay.execute(command));
    set(0, 0, '1/2/2026');
    set(1, 0, '=A1+1');
    workbook.execute({
      type: 'clear',
      sheet,
      ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }],
    });
    workbook.undo();
    workbook.undo();
    workbook.redo();
    workbook.undo();
    workbook.undo();
    workbook.redo();
    const stored = JSON.parse(JSON.stringify(replay.toSnapshot()));
    expect(Workbook.fromSnapshot(stored).toSnapshot()).toEqual(
      workbook.toSnapshot(),
    );
    expect(replay.cell(sheet, 0, 0).display).toBe('1/2/2026');
  });

  it('opens a stored workbook with its recorded changes, which cannot be undone', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '1');
    const opened = Workbook.fromSnapshot(workbook.toSnapshot(), [
      { type: 'setCell', sheet, row: 1, col: 0, input: '=A1+1' },
    ]);
    expect(opened.cell(sheet, 1, 0).display).toBe('2');
    expect(opened.canUndo).toBe(false);
  });
});

describe('sheets', () => {
  const names = (workbook: Workbook) =>
    workbook.sheets.map((sheet) => sheet.name);

  it('adds empty sheets with the next free names, where asked', () => {
    const { workbook, sheet } = setup();
    const added = workbook.execute({ type: 'addSheet', index: 1 });
    expect(names(workbook)).toEqual(['Sheet1', 'Sheet2']);
    expect(added).toEqual({ sheet: workbook.sheets[1]!.id, ranges: [] });
    workbook.execute({ type: 'addSheet', index: 0 });
    expect(names(workbook)).toEqual(['Sheet3', 'Sheet1', 'Sheet2']);
    expect(new Set(workbook.sheets.map((each) => each.id)).size).toBe(3);
    expect(workbook.cell(added.sheet, 0, 0).value).toBeNull();
    expect(workbook.sheets[1]!.id).toBe(sheet);
  });

  it('duplicates a sheet beside it, with its contents and formats', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '50%');
    set(1, 0, '=A1+1');
    const { sheet: copy } = workbook.execute({
      type: 'addSheet',
      index: 1,
      copyOf: sheet,
    });
    expect(names(workbook)).toEqual(['Sheet1', 'Sheet1 (2)']);
    expect(workbook.cell(copy, 0, 0).display).toBe('50%');
    expect(workbook.cell(copy, 1, 0).formula).toBe('=A1+1');
    // The copy is a sheet of its own.
    set(0, 0, '1');
    expect(workbook.cell(copy, 1, 0).value).toBe(1.5);
    workbook.execute({ type: 'addSheet', index: 2, copyOf: sheet });
    expect(names(workbook)).toEqual(['Sheet1', 'Sheet1 (2)', 'Sheet1 (3)']);
  });

  it('renames, recolors and moves a sheet', () => {
    const { workbook, sheet } = setup();
    workbook.execute({ type: 'addSheet', index: 1 });
    workbook.execute({ type: 'setSheet', sheet, name: 'Sales' });
    workbook.execute({ type: 'setSheet', sheet, color: '#e5484d' });
    expect(workbook.sheets[0]).toEqual({
      id: sheet,
      name: 'Sales',
      color: '#e5484d',
    });
    workbook.execute({ type: 'setSheet', sheet, index: 1 });
    expect(names(workbook)).toEqual(['Sheet2', 'Sales']);
    workbook.execute({ type: 'setSheet', sheet, color: null });
    expect(workbook.sheets[1]).toEqual({ id: sheet, name: 'Sales' });
  });

  it('deletes a sheet, and shows the one that takes its place', () => {
    const { workbook, sheet } = setup();
    const { sheet: second } = workbook.execute({ type: 'addSheet', index: 1 });
    const { sheet: third } = workbook.execute({ type: 'addSheet', index: 2 });
    expect(workbook.execute({ type: 'deleteSheet', sheet: second })).toEqual({
      sheet: third,
      ranges: [],
    });
    expect(workbook.execute({ type: 'deleteSheet', sheet: third })).toEqual({
      sheet,
      ranges: [],
    });
    expect(names(workbook)).toEqual(['Sheet1']);
  });

  it('refuses to delete the last sheet', () => {
    const { workbook, sheet } = setup();
    expect(() => workbook.execute({ type: 'deleteSheet', sheet })).toThrow();
    expect(names(workbook)).toEqual(['Sheet1']);
    expect(workbook.canUndo).toBe(false);
  });

  it('tells what is wrong with a sheet name, as Excel does', () => {
    const { workbook, sheet } = setup();
    const { sheet: second } = workbook.execute({ type: 'addSheet', index: 1 });
    expect(workbook.sheetNameProblem('Budget 2026', sheet)).toBeNull();
    expect(workbook.sheetNameProblem('sheet1', sheet)).toBeNull();
    expect(workbook.sheetNameProblem('SHEET1', second)).toBe('taken');
    expect(workbook.sheetNameProblem('  ', sheet)).toBe('empty');
    expect(workbook.sheetNameProblem('x'.repeat(32), sheet)).toBe('tooLong');
    for (const name of ['a:b', 'a/b', 'a\\b', 'a?', 'a*', '[a]', "'a"]) {
      expect(workbook.sheetNameProblem(name, sheet)).toBe('character');
    }
    expect(() =>
      workbook.execute({ type: 'setSheet', sheet, name: 'Sheet2' }),
    ).toThrow();
    expect(names(workbook)).toEqual(['Sheet1', 'Sheet2']);
  });

  it('undoes and redoes each change to the sheets, showing the sheet it changed', () => {
    const { workbook, sheet, set } = setup();
    set(0, 0, '7');
    const before = workbook.toSnapshot();
    const { sheet: copy } = workbook.execute({
      type: 'addSheet',
      index: 1,
      copyOf: sheet,
    });
    workbook.execute({
      type: 'setSheet',
      sheet: copy,
      name: 'Copy',
      color: '#30a46c',
    });
    workbook.execute({ type: 'setSheet', sheet: copy, index: 0 });
    workbook.execute({ type: 'deleteSheet', sheet });
    const after = workbook.toSnapshot();

    expect(workbook.undo()).toEqual({ sheet, ranges: [] });
    expect(names(workbook)).toEqual(['Copy', 'Sheet1']);
    expect(workbook.cell(sheet, 0, 0).value).toBe(7);
    expect(workbook.undo()).toEqual({ sheet: copy, ranges: [] });
    expect(names(workbook)).toEqual(['Sheet1', 'Copy']);
    workbook.undo();
    expect(workbook.sheets[1]).toEqual({ id: copy, name: 'Sheet1 (2)' });
    expect(workbook.undo()).toEqual({ sheet, ranges: [] });
    expect(workbook.toSnapshot()).toEqual(before);

    for (let step = 0; step < 4; step++) workbook.redo();
    expect(workbook.toSnapshot()).toEqual(after);
  });

  it('tells listeners of sheet changes as commands that replay them', () => {
    const { workbook, sheet, set } = setup();
    const replay = createWorkbook();
    workbook.subscribe((command) => replay.execute(command));
    set(0, 0, '1');
    workbook.execute({ type: 'addSheet', index: 1 });
    const { sheet: copy } = workbook.execute({
      type: 'addSheet',
      index: 0,
      copyOf: sheet,
    });
    workbook.execute({ type: 'setSheet', sheet: copy, name: 'Copy' });
    workbook.execute({ type: 'deleteSheet', sheet });
    workbook.undo();
    workbook.undo();
    workbook.redo();
    const stored = JSON.parse(JSON.stringify(replay.toSnapshot()));
    expect(Workbook.fromSnapshot(stored).toSnapshot()).toEqual(
      workbook.toSnapshot(),
    );
  });
});
