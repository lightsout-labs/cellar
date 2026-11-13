import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook } from '../../engine/workbook';
import { startAutosave } from './autosave';

function setup() {
  const workbook = createWorkbook();
  const sheet = workbook.sheets[0]!.id;
  const save = vi.fn(() => Promise.resolve());
  const autosave = startAutosave(workbook, save);
  const edit = (input: string) =>
    workbook.execute({ type: 'setCell', sheet, row: 0, col: 0, input });
  return { workbook, save, autosave, edit };
}

describe('autosave', () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });
  afterEach(() => {
    vi.useRealTimers();
  });

  it('saves once, half a second after the last change', () => {
    const { workbook, save, edit } = setup();
    edit('1');
    vi.advanceTimersByTime(400);
    edit('2');
    vi.advanceTimersByTime(499);
    expect(save).not.toHaveBeenCalled();
    vi.advanceTimersByTime(1);
    expect(save).toHaveBeenCalledTimes(1);
    expect(save).toHaveBeenCalledWith(workbook.toSnapshot());
  });

  it('saves at once when flushed, and not again later', () => {
    const { save, autosave, edit } = setup();
    edit('1');
    autosave.flush();
    expect(save).toHaveBeenCalledTimes(1);
    vi.advanceTimersByTime(1000);
    expect(save).toHaveBeenCalledTimes(1);
  });

  it('does not save when nothing changed', () => {
    const { save, autosave } = setup();
    autosave.flush();
    vi.advanceTimersByTime(1000);
    expect(save).not.toHaveBeenCalled();
  });

  it('saves pending changes when stopped, then ignores later ones', () => {
    const { save, autosave, edit } = setup();
    edit('1');
    autosave.stop();
    expect(save).toHaveBeenCalledTimes(1);
    edit('2');
    vi.advanceTimersByTime(1000);
    expect(save).toHaveBeenCalledTimes(1);
  });
});
