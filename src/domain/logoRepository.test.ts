import { afterEach, describe, expect, it, vi } from "vitest";
import {
  __resetLogoRepositoryForTests,
  isDatabaseConnectionClosingError,
  logoRepository,
} from "./logoRepository";

describe("logoRepository", () => {
  afterEach(() => {
    __resetLogoRepositoryForTests();
    vi.unstubAllGlobals();
  });

  it("detects closing IndexedDB connections", () => {
    expect(
      isDatabaseConnectionClosingError(
        new Error("Failed to execute 'transaction' on 'IDBDatabase': The database connection is closing."),
      ),
    ).toBe(true);
    expect(isDatabaseConnectionClosingError(new Error("QuotaExceededError"))).toBe(false);
  });

  it("reopens IndexedDB once when a cached connection is closing", async () => {
    let didThrowClosingError = false;
    const open = vi.fn(() => {
      const db = createFakeDatabase(() => {
        if (!didThrowClosingError) {
          didThrowClosingError = true;
          throw new Error("Failed to execute 'transaction' on 'IDBDatabase': The database connection is closing.");
        }
      });
      return createAsyncRequest(db);
    });

    vi.stubGlobal("indexedDB", { open });

    await expect(logoRepository.listAssets()).resolves.toEqual([]);
    expect(open).toHaveBeenCalledTimes(2);
  });
});

function createFakeDatabase(beforeTransaction: () => void): IDBDatabase {
  const database = {
    objectStoreNames: {
      contains: () => true,
    },
    createObjectStore: vi.fn(),
    close: vi.fn(),
    onversionchange: null,
    transaction: vi.fn(() => {
      beforeTransaction();
      return {
        error: null,
        onabort: null,
        onerror: null,
        objectStore: () => ({
          getAll: () => createAsyncRequest([]),
        }),
      };
    }),
  };

  return database as unknown as IDBDatabase;
}

function createAsyncRequest<T>(result: T): IDBRequest<T> {
  const request = {
    result,
    error: null,
    onerror: null,
    onsuccess: null,
  } as unknown as IDBRequest<T>;
  queueMicrotask(() => {
    request.onsuccess?.call(request, { target: request } as unknown as Event);
  });
  return request;
}
