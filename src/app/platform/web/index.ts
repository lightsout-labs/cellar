import type {
  Command,
  SheetSnapshot,
  Snapshot,
  Style,
} from '../../../engine/workbook';
import type { ChosenFile, Platform, Stored } from '..';

// The IndexedDB database `gridworks`. `workbooks` has one record for each
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
    const request = indexedDB.open('gridworks', 2);
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
              ...(record.color !== undefined && { color: record.color }),
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
  // Deleted sheets go too: every sheet of the workbook is written again.
  // Arrays sort after strings, so `[id, []]` follows every sheet key.
  transaction.objectStore(SHEETS).delete(IDBKeyRange.bound([id], [id, []]));
  for (const sheet of snapshot.sheets) {
    transaction
      .objectStore(SHEETS)
      .put({ ...sheet, workbook: id } satisfies SheetRecord);
  }
  transaction.objectStore(CHANGES).delete(changesOf(id, last));
  await completion(transaction);
}

// A file input works in every browser; the File System Access API does not.
function openFile(): Promise<ChosenFile | null> {
  return new Promise((resolve, reject) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx';
    input.hidden = true;
    let changed = false;
    input.addEventListener('change', () => {
      changed = true;
      input.remove();
      const file = input.files?.[0];
      if (!file) return resolve(null);
      file
        .arrayBuffer()
        .then((buffer) =>
          resolve({ name: file.name, bytes: new Uint8Array(buffer) }),
        )
        .catch(reject);
    });
    // Firefox may fire `cancel` just before the `change` of a file chosen,
    // so a cancellation counts only when no change follows it; a late
    // answer costs nothing, as no file chosen means nothing to do.
    input.addEventListener('cancel', () => {
      setTimeout(() => {
        if (changed) return;
        input.remove();
        resolve(null);
      }, 1000);
    });
    document.body.append(input);
    input.click();
  });
}

// A download: browsers keep it in the downloads folder or ask where.
function saveFile(name: string, bytes: Uint8Array): Promise<void> {
  const url = URL.createObjectURL(new Blob([bytes as Uint8Array<ArrayBuffer>]));
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.hidden = true;
  document.body.append(link);
  link.click();
  link.remove();
  // Revoked later, as some browsers read the address after the click.
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
  return Promise.resolve();
}

/** The platform for browsers. */
export const webPlatform: Platform = {
  openFile,
  saveFile,
  workbooks: { load, record, save },
  // A browser that blocks storage throws; the choice is then not remembered.
  preferences: {
    get(name) {
      try {
        return localStorage.getItem(`gridworks.${name}`);
      } catch {
        return null;
      }
    },
    set(name, value) {
      try {
        localStorage.setItem(`gridworks.${name}`, value);
      } catch {
        // Not remembered.
      }
    },
  },
};
