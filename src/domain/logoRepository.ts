import type { LogoAsset, LogoPlacement } from "./types";

const DB_NAME = "brand-logo-stamp";
const DB_VERSION = 1;
const ASSET_STORE = "logoAssets";
const PLACEMENT_STORE = "logoPlacements";

let dbPromise: Promise<IDBDatabase> | null = null;

function openDatabase(): Promise<IDBDatabase> {
  if (dbPromise) {
    return dbPromise;
  }

  dbPromise = new Promise((resolve, reject) => {
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

    request.onerror = () => reject(request.error ?? new Error("无法打开本地 Logo 库。"));
    request.onsuccess = () => resolve(request.result);
  });

  return dbPromise;
}

function runStore<T>(
  storeName: typeof ASSET_STORE | typeof PLACEMENT_STORE,
  mode: IDBTransactionMode,
  operation: (store: IDBObjectStore) => IDBRequest<T>,
): Promise<T> {
  return openDatabase().then(
    (db) =>
      new Promise((resolve, reject) => {
        const transaction = db.transaction(storeName, mode);
        const store = transaction.objectStore(storeName);
        const request = operation(store);

        request.onerror = () => reject(request.error ?? new Error("本地 Logo 库操作失败。"));
        request.onsuccess = () => resolve(request.result);
        transaction.onerror = () => reject(transaction.error ?? new Error("本地 Logo 库事务失败。"));
      }),
  );
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
