import type { LogoAsset, LogoPlacement } from "./types";

const DB_NAME = "brand-logo-stamp";
const DB_VERSION = 1;
const ASSET_STORE = "logoAssets";
const PLACEMENT_STORE = "logoPlacements";

let dbPromise: Promise<IDBDatabase> | null = null;

type StoreName = typeof ASSET_STORE | typeof PLACEMENT_STORE;

function openDatabase(): Promise<IDBDatabase> {
  if (dbPromise) {
    return dbPromise;
  }

  dbPromise = new Promise((resolve, reject) => {
    if (typeof indexedDB === "undefined") {
      reject(new Error("当前环境不支持本地 Logo 库。"));
      return;
    }

    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(ASSET_STORE)) {
        db.createObjectStore(ASSET_STORE, { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains(PLACEMENT_STORE)) {
        db.createObjectStore(PLACEMENT_STORE, { keyPath: "id" });
      }
    };

    request.onerror = () => {
      resetDatabase();
      reject(request.error ?? new Error("无法打开本地 Logo 库。"));
    };
    request.onsuccess = () => {
      const db = request.result;
      const closeAwareDb = db as IDBDatabase & { onclose?: (() => void) | null };
      closeAwareDb.onclose = resetDatabase;
      db.onversionchange = () => {
        db.close();
        resetDatabase();
      };
      resolve(db);
    };
  });

  return dbPromise;
}

function resetDatabase(): void {
  dbPromise = null;
}

async function runStore<T>(
  storeName: StoreName,
  mode: IDBTransactionMode,
  operation: (store: IDBObjectStore) => IDBRequest<T>,
): Promise<T> {
  try {
    return await runStoreOnce(storeName, mode, operation);
  } catch (error) {
    if (!isDatabaseConnectionClosingError(error)) {
      throw error;
    }

    resetDatabase();
    return runStoreOnce(storeName, mode, operation);
  }
}

function runStoreOnce<T>(
  storeName: StoreName,
  mode: IDBTransactionMode,
  operation: (store: IDBObjectStore) => IDBRequest<T>,
): Promise<T> {
  return openDatabase().then(
    (db) =>
      new Promise<T>((resolve, reject) => {
        let transaction: IDBTransaction;
        let request: IDBRequest<T>;

        try {
          transaction = db.transaction(storeName, mode);
          const store = transaction.objectStore(storeName);
          request = operation(store);
        } catch (error) {
          reject(error);
          return;
        }

        request.onerror = () => reject(request.error ?? new Error("本地 Logo 库操作失败。"));
        request.onsuccess = () => resolve(request.result);
        transaction.onerror = () => reject(transaction.error ?? new Error("本地 Logo 库事务失败。"));
        transaction.onabort = () => reject(transaction.error ?? new Error("本地 Logo 库事务已中止。"));
      }),
  );
}

export function isDatabaseConnectionClosingError(error: unknown): boolean {
  if (!(error instanceof Error || error instanceof DOMException)) {
    return false;
  }

  const message = error.message.toLowerCase();
  return message.includes("database connection is closing") || message.includes("connection is closing");
}

export function __resetLogoRepositoryForTests(): void {
  resetDatabase();
}

export const logoRepository = {
  listAssets(): Promise<LogoAsset[]> {
    return runStore<LogoAsset[]>(ASSET_STORE, "readonly", (store) => store.getAll());
  },

  saveAsset(asset: LogoAsset): Promise<IDBValidKey> {
    return runStore<IDBValidKey>(ASSET_STORE, "readwrite", (store) => store.put(asset));
  },

  deleteAsset(assetId: string): Promise<undefined> {
    return runStore<undefined>(ASSET_STORE, "readwrite", (store) => store.delete(assetId));
  },

  listPlacements(): Promise<LogoPlacement[]> {
    return runStore<LogoPlacement[]>(PLACEMENT_STORE, "readonly", (store) => store.getAll());
  },

  savePlacement(placement: LogoPlacement): Promise<IDBValidKey> {
    return runStore<IDBValidKey>(PLACEMENT_STORE, "readwrite", (store) => store.put(placement));
  },

  deletePlacement(placementId: string): Promise<undefined> {
    return runStore<undefined>(PLACEMENT_STORE, "readwrite", (store) => store.delete(placementId));
  },
};
