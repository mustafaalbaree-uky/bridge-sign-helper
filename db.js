// On-device storage (IndexedDB). Stores:
//   captures — offline-first photo queue; blobs live here until a flush clears them.
//   signs    — local cache of the sign list so the field app works without signal.
//   meta     — small cached payloads (e.g. the server capture list) for instant paint.
//   thumbs   — downscaled thumbnails keyed by storage path, so Review doesn't
//              re-download full-size photos every time it opens.
//
// Capture record:
//   { key, signId, slot (1|2), batchDate 'YYYY-MM-DD', blob, capturedAt (ISO),
//     captureLat, captureLng, storagePath, status 'pending'|'synced'|'error', error }
window.DB = (() => {
  const DB_NAME = "bridge-sign-helper";
  const DB_VERSION = 3;
  const CAPS = "captures";
  const SIGNS = "signs";
  const META = "meta";
  const THUMBS = "thumbs";
  let dbp = null;

  function open() {
    if (dbp) return dbp;
    dbp = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(CAPS)) db.createObjectStore(CAPS, { keyPath: "key" });
        if (!db.objectStoreNames.contains(SIGNS)) db.createObjectStore(SIGNS, { keyPath: "id" });
        if (!db.objectStoreNames.contains(META)) db.createObjectStore(META);
        if (!db.objectStoreNames.contains(THUMBS)) db.createObjectStore(THUMBS);
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return dbp;
  }

  const reqP = (req) =>
    new Promise((resolve, reject) => {
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });

  function tx(store, mode, fn) {
    return open().then(
      (db) =>
        new Promise((resolve, reject) => {
          const t = db.transaction(store, mode);
          let result;
          Promise.resolve(fn(t.objectStore(store)))
            .then((r) => (result = r))
            .catch(reject);
          t.oncomplete = () => resolve(result);
          t.onerror = () => reject(t.error);
          t.onabort = () => reject(t.error);
        })
    );
  }

  return {
    // captures
    putCapture: (rec) => tx(CAPS, "readwrite", (s) => reqP(s.put(rec))),
    allCaptures: () => tx(CAPS, "readonly", (s) => reqP(s.getAll())),
    removeCapture: (key) => tx(CAPS, "readwrite", (s) => reqP(s.delete(key))),
    clearCaptures: () => tx(CAPS, "readwrite", (s) => reqP(s.clear())),
    // signs cache
    putSigns: (rows) =>
      tx(SIGNS, "readwrite", (s) => {
        s.clear();
        rows.forEach((r) => s.put(r));
      }),
    allSigns: () => tx(SIGNS, "readonly", (s) => reqP(s.getAll())),
    // small cached payloads
    getMeta: (k) => tx(META, "readonly", (s) => reqP(s.get(k))),
    setMeta: (k, v) => tx(META, "readwrite", (s) => reqP(s.put(v, k))),
    // thumbnails
    getThumb: (k) => tx(THUMBS, "readonly", (s) => reqP(s.get(k))),
    setThumb: (k, blob) => tx(THUMBS, "readwrite", (s) => reqP(s.put(blob, k))),
    dropThumbs: (keys) =>
      tx(THUMBS, "readwrite", (s) => keys.forEach((k) => s.delete(k))),
  };
})();
