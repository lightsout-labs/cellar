import type { Snapshot } from '../../engine/workbook';

/**
 * Everything the app needs from the system it runs on. Only an
 * implementation of this interface may touch files, storage, the clipboard
 * or printing; `web` is the implementation for browsers.
 */
export interface Platform {
  /** The workbooks kept in the system's storage. */
  readonly workbooks: {
    /** The stored workbook, or `null` when there is none with this identifier. */
    load(id: string): Promise<Snapshot | null>;
    save(id: string, snapshot: Snapshot): Promise<void>;
  };
}
