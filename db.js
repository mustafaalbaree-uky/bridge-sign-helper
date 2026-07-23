// Tiny IndexedDB wrapper for on-device photo captures.
// One record per sign per photo slot. Key = `${signId}__${slot}` so a retake
// replaces the same slot and a sign never exceeds two photos.
//
// Record shape:
//   { key, signId, slot (1|2), blob, capturedAt (ISO string),
//     captureLat, captureLng }
window.DB = (() => {
  const DB_NAME = "bridge-sign-helper";
  const DB_VERSION = 1;
  const STORE = "captures";
  let dbp = null;

  function open() {
    if (dbp) return dbp;
    dbp = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(STORE)) {
          db.createObjectStore(STORE, { keyPath: "key" });
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return dbp;
  }

  function tx(mode, fn) {
    return open().then(
      (db) =>
        new Promise((resolve, reject) => {
          const t = db.transaction(STORE, mode);
          const store = t.objectStore(STORE);
          let result;
          Promise.resolve(fn(store))
            .then((r) => (result = r))
            .catch(reject);
          t.oncomplete = () => resolve(result);
          t.onerror = () => reject(t.error);
          t.onabort = () => reject(t.error);
        })
    );
  }

  const reqToPromise = (req) =>
    new Promise((resolve, reject) => {
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });

  return {
    put(record) {
      return tx("readwrite", (s) => reqToPromise(s.put(record)));
    },
    all() {
      return tx("readonly", (s) => reqToPromise(s.getAll()));
    },
    remove(key) {
      return tx("readwrite", (s) => reqToPromise(s.delete(key)));
    },
    clear() {
      return tx("readwrite", (s) => reqToPromise(s.clear()));
    },
  };
})();
