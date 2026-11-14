import type {
  Command,
  SheetSnapshot,
  Snapshot,
  Style,
} from '../../../engine/workbook';
import type { Platform, Stored } from '..';

// The IndexedDB database `cellar`. `workbooks` has one record for each
// workbook; `sheets` has one for each sheet, keyed by workbook and sheet;
// `changes` has the changes made since a workbook was last saved, keyed by
// workbook and a sequence number.
const WORKBOOKS = 'workbooks';
const SHEETS = 'sheets';
const CHANGES = 'changes';

interface WorkbookRecord {
  id: string;
  name: string;
  modified: number;
  version: Snapshot['version'];
  sheetIds: string[];
  /** Absent in records saved before cells had styles. */
  styles?: Style[];
}

type SheetRecord = SheetSnapshot & { workbook: string };

interface ChangeRecord {
  workbook: string;
  sequence: number;
  command: Command;
}

/** The sequence number of the last change recorded or loaded. */
let sequence = 0;

let opening: Promise<IDBDatabase> | undefined;

function database(): Promise<IDBDatabase> {
  opening ??= new Promise((resolve, reject) => {
    const request = indexedDB.open('cellar', 2);
    request.onupgradeneeded = ({ oldVersion }) => {
      const db = request.result;
      if (oldVersion < 1) {
        db.createObjectStore(WORKBOOKS, { keyPath: 'id' });
        db.createObjectStore(SHEETS, { keyPath: ['workbook', 'id'] });
      }
      db.createObjectStore(CHANGES, { keyPath: ['workbook', 'sequence'] });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => {
      opening = undefined; // The next call tries again.
      reject(request.error!);
    };
  });
  return opening;
}

function result<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error!);
  });
}

/** The changes of a workbook up to a sequence number. */
const changesOf = (id: string, last = Infinity) =>
  IDBKeyRange.bound([id, -Infinity], [id, last]);

function completion(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error!);
    transaction.onabort = () => reject(transaction.error!);
  });
}

async function load(id: string): Promise<Stored | null> {
  const transaction = (await database()).transaction([
    WORKBOOKS,
    SHEETS,
    CHANGES,
  ]);
  const sheetStore = transaction.objectStore(SHEETS);
  const [workbook, changes] = await Promise.all([
    result<WorkbookRecord | undefined>(
      transaction.objectStore(WORKBOOKS).get(id),
    ),
    result<ChangeRecord[]>(
      transaction.objectStore(CHANGES).getAll(changesOf(id)),
    ),
  ]);
  sequence = Math.max(sequence, ...changes.map((change) => change.sequence));
  if (!workbook && changes.length === 0) return null;
  const snapshot: Snapshot | null = workbook
    ? {
        version: workbook.version,
        name: workbook.name,
        styles: workbook.styles,
        sheets: await Promise.all(
          workbook.sheetIds.map(async (sheetId) => {
            const record = await result<SheetRecord>(
              sheetStore.get([id, sheetId]),
            );
            return {
              id: record.id,
              name: record.name,
              columns: record.columns,
              formulas: record.formulas,
              styleIds: record.styleIds,
            };
          }),
        ),
      }
    : null;
  return { snapshot, changes: changes.map((change) => change.command) };
}

async function record(id: string, command: Command): Promise<void> {
  // Numbered at once, so that a save called next counts it as saved; its
  // transaction is created first, so it is stored before that save.
  const change: ChangeRecord = { workbook: id, sequence: ++sequence, command };
  const transaction = (await database()).transaction(CHANGES, 'readwrite');
  transaction.objectStore(CHANGES).put(change);
  await completion(transaction);
}

async function save(id: string, snapshot: Snapshot): Promise<void> {
  // The snapshot holds the changes recorded so far, and no later one.
  const last = sequence;
  const transaction = (await database()).transaction(
    [WORKBOOKS, SHEETS, CHANGES],
    'readwrite',
  );
  const workbook: WorkbookRecord = {
    id,
    name: snapshot.name,
    modified: Date.now(),
    version: snapshot.version,
    sheetIds: snapshot.sheets.map((sheet) => sheet.id),
    styles: snapshot.styles,
  };
  transaction.objectStore(WORKBOOKS).put(workbook);
  for (const sheet of snapshot.sheets) {
    transaction
      .objectStore(SHEETS)
      .put({ ...sheet, workbook: id } satisfies SheetRecord);
  }
  transaction.objectStore(CHANGES).delete(changesOf(id, last));
  await completion(transaction);
}

/** The platform for browsers. */
export const webPlatform: Platform = {
  workbooks: { load, record, save },
  // A browser that blocks storage throws; the choice is then not remembered.
  preferences: {
    get(name) {
      try {
        return localStorage.getItem(`cellar.${name}`);
      } catch {
        return null;
      }
    },
    set(name, value) {
      try {
        localStorage.setItem(`cellar.${name}`, value);
      } catch {
        // Not remembered.
      }
    },
  },
};
