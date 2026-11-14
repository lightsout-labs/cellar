import { navigationFor, type KeyPress, type Navigation } from './navigation';

/**
 * The two modes of an entry, as in Excel: Enter mode, started by typing,
 * where an arrow commits the entry and moves; and Edit mode, started by F2,
 * a double-click or the formula bar, where the arrows move the caret.
 */
export type EntryMode = 'enter' | 'edit';

/** What a key does to the grid. */
export type KeyAction =
  /** Moves the selection, committing the entry first. */
  | { type: 'navigate'; navigation: Navigation }
  /** Abandons the entry. */
  | { type: 'cancel' }
  /** Edits the contents of the active cell, or switches to Edit mode. */
  | { type: 'edit' }
  /** Starts an entry that replaces the contents of the active cell. */
  | { type: 'type'; text: string }
  /** Empties the selected cells. */
  | { type: 'clear' }
  /** Cycles the reference at the caret between absolute and relative. */
  | { type: 'cycleReference' };

/**
 * What a key does, as in Excel, with no entry or during one in a mode, or
 * `null` when it is left to the browser, such as a character typed into
 * the editor.
 */
export function keyAction(
  press: KeyPress,
  context: {
    mac: boolean;
    mode: EntryMode | null;
    endMode: boolean;
    page: { rows: number; cols: number };
  },
): KeyAction | null {
  const { key, shift, ctrl, meta, alt } = press;
  const { mac, mode } = context;
  const plain = !ctrl && !meta && !alt;
  const back = shift ? -1 : 1;
  if (plain && key === 'Enter') return step(back, 0);
  if (plain && key === 'Tab') return step(0, back);
  if (key === 'Escape') return mode ? { type: 'cancel' } : null;
  if (key === 'F2') return mode === 'edit' ? null : { type: 'edit' };
  if (key === 'F4') return mode ? { type: 'cycleReference' } : null;
  if (mode === 'edit') return null;
  if (mode === 'enter') {
    if (!key.startsWith('Arrow')) return null;
    const navigation = navigationFor(press, context);
    return navigation && { type: 'navigate', navigation };
  }
  // The key labeled delete on a Mac is Backspace.
  if (key === 'Delete' || (mac && key === 'Backspace'))
    return { type: 'clear' };
  if (key === 'Backspace') return { type: 'type', text: '' };
  const navigation = navigationFor(press, context);
  if (navigation) return { type: 'navigate', navigation };
  if (key.length === 1 && plain) return { type: 'type', text: key };
  return null;
}

const step = (rows: number, cols: number): KeyAction => ({
  type: 'navigate',
  navigation: { type: 'step', rows, cols, extend: false },
});
