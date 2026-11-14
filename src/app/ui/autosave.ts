import {
  createWorkbook,
  Workbook,
  type Command,
  type Snapshot,
} from '../../engine/workbook';
import type { Stored } from '../platform';

const DELAY_MS = 500;

/** Where a workbook is kept. */
export interface Store {
  /** Keeps one change, until the next save includes it. */
  record(command: Command): Promise<void>;
  /** Keeps the whole workbook, in place of the changes recorded so far. */
  save(snapshot: Snapshot): Promise<void>;
}

export interface Autosave {
  /** Saves now if a change is waiting. */
  flush(): void;
  /** Saves a waiting change and stops watching the workbook. */
  stop(): void;
}

/**
 * Keeps every change to the workbook as it is made, so that none is lost
 * when the browser closes, and saves the whole workbook half a second after
 * the last change.
 */
export function startAutosave(workbook: Workbook, store: Store): Autosave {
  let timer: ReturnType<typeof setTimeout> | undefined;

  const flush = () => {
    if (timer === undefined) return;
    clearTimeout(timer);
    timer = undefined;
    void store.save(workbook.toSnapshot());
  };
  const unsubscribe = workbook.subscribe((command) => {
    void store.record(command);
    clearTimeout(timer);
    timer = setTimeout(flush, DELAY_MS);
  });

  return {
    flush,
    stop() {
      unsubscribe();
      flush();
    },
  };
}

/**
 * The workbook as it was stored: the saved workbook, or a new one, with
 * the changes recorded after it replayed.
 */
export function restoreWorkbook(stored: Stored | null): Workbook {
  const workbook = stored?.snapshot
    ? Workbook.fromSnapshot(stored.snapshot)
    : createWorkbook();
  for (const command of stored?.changes ?? []) workbook.execute(command);
  return workbook;
}
