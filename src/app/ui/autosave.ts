import type { Snapshot, Workbook } from '../../engine/workbook';

const DELAY_MS = 500;

export interface Autosave {
  /** Saves now if a change is waiting. */
  flush(): void;
  /** Saves a waiting change and stops watching the workbook. */
  stop(): void;
}

/** Saves the workbook half a second after the last change to it. */
export function startAutosave(
  workbook: Workbook,
  save: (snapshot: Snapshot) => Promise<void>,
): Autosave {
  let timer: ReturnType<typeof setTimeout> | undefined;

  const flush = () => {
    if (timer === undefined) return;
    clearTimeout(timer);
    timer = undefined;
    void save(workbook.toSnapshot());
  };
  const unsubscribe = workbook.subscribe(() => {
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
