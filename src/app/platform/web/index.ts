import type { SheetSnapshot, Snapshot } from '../../../engine/workbook';
import type { Platform } from '..';

// The IndexedDB database `cellar`. `workbooks` has one record for each
// workbook; `sheets` has one for each sheet, keyed by workbook and sheet.
const WORKBOOKS = 'workbooks';
const SHEETS = 'sheets';

interface WorkbookRecord {
  id: string;
  name: string;
  modified: number;
  version: Snapshot['version'];
  sheetIds: string[];
}

type SheetRecord = SheetSnapshot & { workbook: string };

let opening: Promise<IDBDatabase> | undefined;

function database(): Promise<IDBDatabase> {
  opening ??= new Promise((resolve, reject) => {
    const request = indexedDB.open('cellar', 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore(WORKBOOKS, { keyPath: 'id' });
      request.result.createObjectStore(SHEETS, { keyPath: ['workbook', 'id'] });
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

async function load(id: string): Promise<Snapshot | null> {
  const transaction = (await database()).transaction([WORKBOOKS, SHEETS]);
  const sheetStore = transaction.objectStore(SHEETS);
  const workbook = await result<WorkbookRecord | undefined>(
    transaction.objectStore(WORKBOOKS).get(id),
  );
  if (!workbook) return null;
  const sheets = await Promise.all(
    workbook.sheetIds.map(async (sheetId) => {
      const record = await result<SheetRecord>(sheetStore.get([id, sheetId]));
      return {
        id: record.id,
        name: record.name,
        columns: record.columns,
        formulas: record.formulas,
      };
    }),
  );
  return { version: workbook.version, name: workbook.name, sheets };
}

async function save(id: string, snapshot: Snapshot): Promise<void> {
  const transaction = (await database()).transaction(
    [WORKBOOKS, SHEETS],
    'readwrite',
  );
  const record: WorkbookRecord = {
    id,
    name: snapshot.name,
    modified: Date.now(),
    version: snapshot.version,
    sheetIds: snapshot.sheets.map((sheet) => sheet.id),
  };
  transaction.objectStore(WORKBOOKS).put(record);
  for (const sheet of snapshot.sheets) {
    transaction
      .objectStore(SHEETS)
      .put({ ...sheet, workbook: id } satisfies SheetRecord);
  }
  await new Promise<void>((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error!);
    transaction.onabort = () => reject(transaction.error!);
  });
}

/** The platform for browsers. */
export const webPlatform: Platform = { workbooks: { load, save } };
