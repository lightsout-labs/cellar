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
