// Minimal, dependency-free Supabase client (PostgREST + Storage over fetch).
// Enough for this app; avoids shipping the full supabase-js SDK.
window.SB = (() => {
  const { SUPABASE_URL, SUPABASE_KEY, BUCKET } = window.CONFIG;
  const REST = `${SUPABASE_URL}/rest/v1/`;
  const OBJ = `${SUPABASE_URL}/storage/v1/object/`;

  const headers = (extra = {}) =>
    Object.assign({ apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` }, extra);

  async function select(table, query = "") {
    const res = await fetch(REST + table + (query ? `?${query}` : ""), { headers: headers() });
    if (!res.ok) throw new Error(`select ${table}: ${res.status} ${await res.text()}`);
    return res.json();
  }

  // Upsert rows. `onConflict` = comma-separated columns of the unique constraint.
  async function upsert(table, rows, onConflict) {
    const q = onConflict ? `?on_conflict=${onConflict}` : "";
    const res = await fetch(REST + table + q, {
      method: "POST",
      headers: headers({
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates,return=minimal",
      }),
      body: JSON.stringify(Array.isArray(rows) ? rows : [rows]),
    });
    if (!res.ok) throw new Error(`upsert ${table}: ${res.status} ${await res.text()}`);
  }

  async function remove(table, query) {
    const res = await fetch(REST + table + `?${query}`, {
      method: "DELETE",
      headers: headers({ Prefer: "return=minimal" }),
    });
    if (!res.ok) throw new Error(`delete ${table}: ${res.status} ${await res.text()}`);
  }

  async function uploadPhoto(path, blob) {
    const res = await fetch(OBJ + BUCKET + "/" + path, {
      method: "POST",
      headers: headers({ "Content-Type": blob.type || "image/jpeg", "x-upsert": "true" }),
      body: blob,
    });
    if (!res.ok) throw new Error(`upload ${path}: ${res.status} ${await res.text()}`);
  }

  const publicUrl = (path) => `${SUPABASE_URL}/storage/v1/object/public/${BUCKET}/${path}`;

  async function downloadPhoto(path) {
    const res = await fetch(publicUrl(path));
    if (!res.ok) throw new Error(`download ${path}: ${res.status}`);
    return res.blob();
  }

  // Lightweight connectivity probe.
  async function ping() {
    try {
      const res = await fetch(REST + "signs?select=id&limit=1", { headers: headers() });
      return res.ok;
    } catch {
      return false;
    }
  }

  return { select, upsert, remove, uploadPhoto, downloadPhoto, publicUrl, ping };
})();
