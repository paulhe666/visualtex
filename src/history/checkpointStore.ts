import type { DocumentCheckpoint } from "./historyTypes";

const DATABASE_NAME = "visualtex-history";
const DATABASE_VERSION = 1;
const STORE_NAME = "checkpoints";
const MAX_PERSISTED_CHECKPOINTS = 10;

function hasIndexedDb() {
  return typeof indexedDB !== "undefined";
}

function openDatabase(): Promise<IDBDatabase | null> {
  if (!hasIndexedDb()) return Promise.resolve(null);

  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
    request.onupgradeneeded = () => {
      const database = request.result;
      if (!database.objectStoreNames.contains(STORE_NAME)) {
        const store = database.createObjectStore(STORE_NAME, {
          keyPath: "id",
        });
        store.createIndex("createdAt", "createdAt", { unique: false });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function requestToPromise<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function transactionToPromise(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
    transaction.onabort = () => reject(transaction.error);
  });
}

async function readAllCheckpoints(
  database: IDBDatabase,
): Promise<DocumentCheckpoint[]> {
  const transaction = database.transaction(STORE_NAME, "readonly");
  const completed = transactionToPromise(transaction);
  const request = transaction
    .objectStore(STORE_NAME)
    .getAll() as IDBRequest<DocumentCheckpoint[]>;
  const checkpoints = await requestToPromise(request);
  await completed;
  return checkpoints;
}

export async function persistCheckpoint(
  checkpoint: DocumentCheckpoint,
): Promise<void> {
  const database = await openDatabase();
  if (!database) return;

  try {
    const writeTransaction = database.transaction(STORE_NAME, "readwrite");
    writeTransaction.objectStore(STORE_NAME).put(checkpoint);
    await transactionToPromise(writeTransaction);

    const all = await readAllCheckpoints(database);
    const stale = all
      .sort((left, right) => right.createdAt - left.createdAt)
      .slice(MAX_PERSISTED_CHECKPOINTS);
    if (!stale.length) return;

    const cleanupTransaction = database.transaction(STORE_NAME, "readwrite");
    const cleanupStore = cleanupTransaction.objectStore(STORE_NAME);
    stale.forEach((item) => cleanupStore.delete(item.id));
    await transactionToPromise(cleanupTransaction);
  } finally {
    database.close();
  }
}

export async function loadRecentCheckpoints(): Promise<DocumentCheckpoint[]> {
  const database = await openDatabase();
  if (!database) return [];

  try {
    const checkpoints = await readAllCheckpoints(database);
    return checkpoints
      .sort((left, right) => right.createdAt - left.createdAt)
      .slice(0, MAX_PERSISTED_CHECKPOINTS);
  } finally {
    database.close();
  }
}
