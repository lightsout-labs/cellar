import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook, type Command } from '../../engine/workbook';
import { restoreWorkbook, startAutosave } from './autosave';

function setup() {
  const workbook = createWorkbook();
  const sheet = workbook.sheets[0]!.id;
  const store = {
    record: vi.fn<(command: Command) => Promise<void>>(() => Promise.resolve()),
    save: vi.fn(() => Promise.resolve()),
  };
  const autosave = startAutosave(workbook, store);
  const edit = (input: string) =>
    workbook.execute({ type: 'setCell', sheet, row: 0, col: 0, input });
  return { workbook, sheet, store, autosave, edit };
}

describe('autosave', () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });
  afterEach(() => {
    vi.useRealTimers();
  });

  it('records each change at once', () => {
    const { sheet, store, edit } = setup();
    edit('1');
    expect(store.record).toHaveBeenCalledWith({
      type: 'setCell',
      sheet,
      row: 0,
      col: 0,
      input: '1',
    });
    edit('2');
    expect(store.record).toHaveBeenCalledTimes(2);
  });

  it('saves the whole workbook once, half a second after the last change', () => {
    const { workbook, store, edit } = setup();
    edit('1');
    vi.advanceTimersByTime(400);
    edit('2');
    vi.advanceTimersByTime(499);
    expect(store.save).not.toHaveBeenCalled();
    vi.advanceTimersByTime(1);
    expect(store.save).toHaveBeenCalledTimes(1);
    expect(store.save).toHaveBeenCalledWith(workbook.toSnapshot());
  });

  it('saves at once when flushed, and not again later', () => {
    const { store, autosave, edit } = setup();
    edit('1');
    autosave.flush();
    expect(store.save).toHaveBeenCalledTimes(1);
    vi.advanceTimersByTime(1000);
    expect(store.save).toHaveBeenCalledTimes(1);
  });

  it('does not save when nothing changed', () => {
    const { store, autosave } = setup();
    autosave.flush();
    vi.advanceTimersByTime(1000);
    expect(store.save).not.toHaveBeenCalled();
  });

  it('saves pending changes when stopped, then ignores later ones', () => {
    const { store, autosave, edit } = setup();
    edit('1');
    autosave.stop();
    expect(store.save).toHaveBeenCalledTimes(1);
    edit('2');
    vi.advanceTimersByTime(1000);
    expect(store.save).toHaveBeenCalledTimes(1);
    expect(store.record).toHaveBeenCalledTimes(1);
  });
});

describe('restoreWorkbook', () => {
  it('replays the recorded changes over the saved workbook', () => {
    const { workbook, sheet } = setup();
    workbook.execute({ type: 'setCell', sheet, row: 0, col: 0, input: '1' });
    const restored = restoreWorkbook({
      snapshot: workbook.toSnapshot(),
      changes: [{ type: 'setCell', sheet, row: 1, col: 0, input: '=A1+1' }],
    });
    expect(restored.cell(sheet, 1, 0).display).toBe('2');
    // The history of undo starts with the workbook opened.
    expect(restored.canUndo).toBe(false);
  });

  it('replays changes over a new workbook when none was saved yet', () => {
    const sheet = createWorkbook().sheets[0]!.id;
    const restored = restoreWorkbook({
      snapshot: null,
      changes: [{ type: 'setCell', sheet, row: 0, col: 0, input: '7' }],
    });
    expect(restored.cell(sheet, 0, 0).display).toBe('7');
  });

  it('starts a new workbook when nothing is stored', () => {
    expect(restoreWorkbook(null).extent('sheet1')).toEqual({
      rows: 0,
      cols: 0,
    });
  });
});
