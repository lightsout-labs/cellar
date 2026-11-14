import type { Workbook } from '../../engine/workbook';
import type { CellPosition } from './geometry';

export interface ViewState {
  /** The identifier of the sheet that is shown. */
  readonly sheet: string;
  readonly active: CellPosition;
  /** The text of the entry being typed, or `null` when there is none. */
  readonly entry: string | null;
}

/**
 * What the user is looking at and doing in a workbook: the sheet, the
 * active cell, and the entry being typed. The grid edits it, and the parts
 * of the window around the grid read it.
 */
export class View {
  private current: ViewState;
  private readonly listeners = new Set<() => void>();

  constructor(private readonly workbook: Workbook) {
    this.current = {
      sheet: workbook.sheets[0]!.id,
      active: { row: 0, col: 0 },
      entry: null,
    };
  }

  /** The same object until the next change. */
  get state(): ViewState {
    return this.current;
  }

  /** Calls `listener` after every change, until the returned function is called. */
  readonly subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  /** Makes a cell the active one, after committing an entry that is waiting. */
  select(cell: CellPosition): void {
    this.commit();
    this.change({ active: cell });
  }

  /** Starts an entry in the active cell, or replaces the text of the one there. */
  type(text: string): void {
    this.change({ entry: text });
  }

  /** Writes the entry to the active cell. Does nothing without an entry. */
  commit(): void {
    const { sheet, active, entry } = this.current;
    if (entry === null) return;
    this.change({ entry: null });
    this.workbook.execute({ type: 'setCell', sheet, ...active, input: entry });
  }

  /** Abandons the entry. */
  cancel(): void {
    this.change({ entry: null });
  }

  private change(change: Partial<ViewState>): void {
    this.current = { ...this.current, ...change };
    for (const listener of this.listeners) listener();
  }
}
