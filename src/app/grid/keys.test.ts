import { describe, expect, it } from 'vitest';
import { keyAction, type EntryMode } from './keys';
import type { KeyPress } from './navigation';

const press = (key: string, modifiers: Partial<KeyPress> = {}): KeyPress => ({
  key,
  shift: false,
  ctrl: false,
  meta: false,
  alt: false,
  ...modifiers,
});

const action = (
  key: KeyPress,
  mode: EntryMode | null,
  mac = false,
): ReturnType<typeof keyAction> =>
  keyAction(key, { mac, mode, endMode: false, page: { rows: 20, cols: 8 } });

const step = (rows: number, cols: number, extend = false) => ({
  type: 'navigate',
  navigation: { type: 'step', rows, cols, extend },
});

describe('keyAction', () => {
  it.each([null, 'enter', 'edit'] as const)(
    'commits and moves with Enter and Tab, backwards with Shift, in mode %s',
    (mode) => {
      expect(action(press('Enter'), mode)).toEqual(step(1, 0));
      expect(action(press('Enter', { shift: true }), mode)).toEqual(
        step(-1, 0),
      );
      expect(action(press('Tab'), mode)).toEqual(step(0, 1));
      expect(action(press('Tab', { shift: true }), mode)).toEqual(step(0, -1));
    },
  );

  it('cancels an entry with Escape', () => {
    expect(action(press('Escape'), 'enter')).toEqual({ type: 'cancel' });
    expect(action(press('Escape'), 'edit')).toEqual({ type: 'cancel' });
    expect(action(press('Escape'), null)).toBeNull();
  });

  it('starts editing, or switches an entry to Edit mode, with F2', () => {
    expect(action(press('F2'), null)).toEqual({ type: 'edit' });
    expect(action(press('F2'), 'enter')).toEqual({ type: 'edit' });
    expect(action(press('F2'), 'edit')).toBeNull();
  });

  it('starts an entry with a typed character, but not with a shortcut', () => {
    expect(action(press('a'), null)).toEqual({ type: 'type', text: 'a' });
    expect(action(press('='), null)).toEqual({ type: 'type', text: '=' });
    expect(action(press('b', { ctrl: true }), null)).toBeNull();
    expect(action(press('b', { meta: true }), null)).toBeNull();
    expect(action(press('Shift', { shift: true }), null)).toBeNull();
  });

  it('leaves typed characters to the editor during an entry', () => {
    expect(action(press('a'), 'enter')).toBeNull();
    expect(action(press('a'), 'edit')).toBeNull();
  });

  it('clears the selection with Delete, and with the delete key of a Mac', () => {
    expect(action(press('Delete'), null)).toEqual({ type: 'clear' });
    expect(action(press('Backspace'), null, true)).toEqual({ type: 'clear' });
    expect(action(press('Delete'), 'enter')).toBeNull();
  });

  it('starts an empty entry with Backspace elsewhere, as Excel does', () => {
    expect(action(press('Backspace'), null)).toEqual({
      type: 'type',
      text: '',
    });
  });

  it('commits and moves with the arrows in Enter mode, and leaves them to the caret in Edit mode', () => {
    expect(action(press('ArrowRight'), 'enter')).toEqual(step(0, 1));
    expect(action(press('ArrowUp', { shift: true }), 'enter')).toEqual(
      step(-1, 0, true),
    );
    expect(action(press('ArrowRight'), 'edit')).toBeNull();
    expect(action(press('Home'), 'enter')).toBeNull();
  });

  it('navigates without an entry', () => {
    expect(action(press('ArrowDown'), null)).toEqual(step(1, 0));
    expect(action(press('Home', { ctrl: true }), null)).toEqual({
      type: 'navigate',
      navigation: { type: 'sheetStart', extend: false },
    });
  });
});
