import type { Command, Snapshot } from '../../engine/workbook';

/** A workbook as it is stored. */
export interface Stored {
  /** The workbook as it was last saved, or `null` if it never was. */
  readonly snapshot: Snapshot | null;
  /** The changes recorded after that save, in order. */
  readonly changes: Command[];
}

/**
 * Everything the app needs from the system it runs on. Only an
 * implementation of this interface may touch files, storage, the clipboard
 * or printing; `web` is the implementation for browsers.
 */
export interface Platform {
  /** The workbooks kept in the system's storage. */
  readonly workbooks: {
    /** The stored workbook, or `null` when there is none with this identifier. */
    load(id: string): Promise<Stored | null>;
    /**
     * Keeps a change to a workbook. It is small and quick to store, so that
     * a change is kept even when the browser closes right after it.
     */
    record(id: string, command: Command): Promise<void>;
    /**
     * Keeps a whole workbook, in place of the changes recorded before this
     * call.
     */
    save(id: string, snapshot: Snapshot): Promise<void>;
  };
  /** The user's choices, such as the language, that outlive the page. */
  readonly preferences: {
    /** The remembered value, or `null` when there is none. */
    get(name: string): string | null;
    set(name: string, value: string): void;
  };
}
