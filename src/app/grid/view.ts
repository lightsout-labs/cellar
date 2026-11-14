import type { Workbook } from '../../engine/workbook';
import type { EntryMode } from './keys';
import { navigate, type Navigation, type SheetContents } from './navigation';
import { single, type Selection } from './selection';

export const ZOOM_MIN = 0.5;
export const ZOOM_MAX = 2;

/** An entry being typed into the active cell. */
export interface Entry {
  readonly text: string;
  readonly mode: EntryMode;
  /** Where it is typed: in the cell or in the formula bar. */
  readonly place: 'cell' | 'formulaBar';
}

export interface ViewState {
  /** The identifier of the sheet that is shown. */
  readonly sheet: string;
  readonly selection: Selection;
  /** The entry being typed, or `null` when there is none. */
  readonly entry: Entry | null;
  /** The size of the grid, where 1 is 100%. */
  readonly zoom: number;
  /** Whether End was pressed, so that the next arrow jumps as Ctrl+arrow. */
  readonly endMode: boolean;
}

/**
 * What the user is looking at and doing in a workbook: the sheet, the
 * selection, the entry being typed and the zoom. The grid edits it, and
 * the parts of the window around the grid read it.
 */
export class View {
  private current: ViewState;
  private readonly listeners = new Set<() => void>();

  constructor(private readonly workbook: Workbook) {
    this.current = {
      sheet: workbook.sheets[0]!.id,
      selection: single({ row: 0, col: 0 }),
      entry: null,
      zoom: 1,
      endMode: false,
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

  /** Changes the selection, after committing an entry that is waiting. */
  select(selection: Selection): void {
    this.commit();
    this.change({ selection, endMode: false });
  }

  /** Moves the selection by the keyboard, after committing a waiting entry. */
  navigate(navigation: Navigation): void {
    if (navigation.type === 'endMode') {
      this.change({ endMode: !this.current.endMode });
      return;
    }
    this.commit();
    const { workbook } = this;
    const { sheet, selection } = this.current;
    const contents: SheetContents = {
      isEmpty: (row, col) => workbook.cell(sheet, row, col).value === null,
      extent: () => workbook.extent(sheet),
    };
    this.select(navigate(selection, navigation, contents));
  }

  /**
   * Starts an entry in Enter mode that replaces the contents of the active
   * cell, or replaces the text of the entry there.
   */
  type(text: string): void {
    const { entry } = this.current;
    this.change({
      entry: entry
        ? { ...entry, text }
        : { text, mode: 'enter', place: 'cell' },
    });
  }

  /**
   * Starts an entry in Edit mode that holds the contents of the active cell,
   * or moves the entry there to Edit mode, in the cell or the formula bar.
   */
  edit(place: Entry['place']): void {
    const { sheet, selection, entry } = this.current;
    const { row, col } = selection.active;
    const text = entry?.text ?? this.workbook.cell(sheet, row, col).input;
    this.change({ entry: { text, mode: 'edit', place } });
  }

  /** Writes the entry to the active cell. Does nothing without an entry. */
  commit(): void {
    const { sheet, selection, entry } = this.current;
    if (entry === null) return;
    this.change({ entry: null });
    const { row, col } = selection.active;
    const input = entry.text;
    this.workbook.execute({ type: 'setCell', sheet, row, col, input });
  }

  /** Empties the selected cells. */
  clear(): void {
    const { sheet, selection } = this.current;
    this.workbook.execute({ type: 'clear', sheet, ranges: selection.ranges });
  }

  /** Abandons the entry. */
  cancel(): void {
    this.change({ entry: null });
  }

  /** Sets the zoom, kept between 50% and 200% and in whole percents. */
  zoomTo(zoom: number): void {
    const percent = Math.round(
      Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, zoom)) * 100,
    );
    this.change({ zoom: percent / 100 });
  }

  private change(change: Partial<ViewState>): void {
    this.current = { ...this.current, ...change };
    for (const listener of this.listeners) listener();
  }
}
