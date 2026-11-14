import { describe, expect, it, vi } from 'vitest';
import { createWorkbook } from '../../engine/workbook';
import { View } from './view';

function setup() {
  const workbook = createWorkbook();
  const view = new View(workbook);
  const sheet = workbook.sheets[0]!.id;
  return { workbook, view, sheet };
}

describe('View', () => {
  it('starts on A1 of the first sheet with no entry', () => {
    const { view, sheet } = setup();
    expect(view.state).toEqual({
      sheet,
      active: { row: 0, col: 0 },
      entry: null,
    });
  });

  it('commits the entry to the active cell', () => {
    const { workbook, view, sheet } = setup();
    view.select({ row: 2, col: 1 });
    view.type('4');
    view.type('42');
    expect(view.state.entry).toBe('42');
    expect(workbook.cell(sheet, 2, 1).display).toBe('');

    view.commit();
    expect(view.state.entry).toBeNull();
    expect(workbook.cell(sheet, 2, 1).display).toBe('42');
  });

  it('commits a waiting entry to its own cell when another is selected', () => {
    const { workbook, view, sheet } = setup();
    view.type('7');
    view.select({ row: 5, col: 5 });
    expect(view.state).toMatchObject({
      active: { row: 5, col: 5 },
      entry: null,
    });
    expect(workbook.cell(sheet, 0, 0).display).toBe('7');
    expect(workbook.cell(sheet, 5, 5).display).toBe('');
  });

  it('abandons the entry on cancel', () => {
    const { workbook, view, sheet } = setup();
    view.type('7');
    view.cancel();
    view.commit(); // Nothing is waiting any more.
    expect(view.state.entry).toBeNull();
    expect(workbook.cell(sheet, 0, 0).display).toBe('');
  });

  it('tells its listeners of each change, with a new state', () => {
    const { view } = setup();
    const before = view.state;
    const listener = vi.fn();
    const unsubscribe = view.subscribe(listener);
    view.select({ row: 1, col: 0 });
    expect(listener).toHaveBeenCalledTimes(1);
    expect(view.state).not.toBe(before);
    expect(view.state).toBe(view.state);

    unsubscribe();
    view.type('1');
    expect(listener).toHaveBeenCalledTimes(1);
  });
});
