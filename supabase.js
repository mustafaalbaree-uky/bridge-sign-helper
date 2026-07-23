// Minimal, dependency-free Supabase client (Auth + PostgREST + Storage over fetch).
window.SB = (() => {
  const { SUPABASE_URL, SUPABASE_KEY, BUCKET } = window.CONFIG;
  const REST = `${SUPABASE_URL}/rest/v1/`;
  const OBJ = `${SUPABASE_URL}/storage/v1/object/`;
  const AUTH = `${SUPABASE_URL}/auth/v1/`;
  const SESSION_KEY = "bsh_session";

  // ---- session -------------------------------------------------------------
  let session = null;
  try { session = JSON.parse(localStorage.getItem(SESSION_KEY) || "null"); } catch {}

  function saveSession(s) {
    session = s;
    if (s) localStorage.setItem(SESSION_KEY, JSON.stringify(s));
    else localStorage.removeItem(SESSION_KEY);
  }

  const isLoggedIn = () => !!(session && session.access_token);
  const currentUser = () => (session ? session.username || session.email : null);

  async function login(email, password, username) {
    const res = await fetch(`${AUTH}token?grant_type=password`, {
      method: "POST",
      headers: { apikey: SUPABASE_KEY, "Content-Type": "application/json" },
      body: JSON.stringify({ email, password }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok)
      throw new Error(data.error_description || data.msg || data.message || `Login failed (${res.status})`);
    saveSession({
      access_token: data.access_token,
      refresh_token: data.refresh_token,
      expires_at: Date.now() + (data.expires_in || 3600) * 1000,
      email,
      username: username || email,
    });
  }

  async function refresh() {
    if (!session || !session.refresh_token) throw new Error("no session");
    const res = await fetch(`${AUTH}token?grant_type=refresh_token`, {
      method: "POST",
      headers: { apikey: SUPABASE_KEY, "Content-Type": "application/json" },
      body: JSON.stringify({ refresh_token: session.refresh_token }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) { saveSession(null); throw new Error("session expired"); }
    saveSession(Object.assign({}, session, {
      access_token: data.access_token,
      refresh_token: data.refresh_token,
      expires_at: Date.now() + (data.expires_in || 3600) * 1000,
    }));
  }

  function logout() { saveSession(null); }

  async function token() {
    if (!session) return null;
    if (Date.now() > session.expires_at - 60000) {
      try { await refresh(); } catch { return null; }
    }
    return session.access_token;
  }

  async function headers(extra = {}) {
    const t = await token();
    return Object.assign({ apikey: SUPABASE_KEY, Authorization: `Bearer ${t || SUPABASE_KEY}` }, extra);
  }

  // ---- data ----------------------------------------------------------------
  async function select(table, query = "") {
    const res = await fetch(REST + table + (query ? `?${query}` : ""), { headers: await headers() });
    if (!res.ok) throw new Error(`select ${table}: ${res.status} ${await res.text()}`);
    return res.json();
  }

  async function upsert(table, rows, onConflict) {
    const q = onConflict ? `?on_conflict=${onConflict}` : "";
    const res = await fetch(REST + table + q, {
      method: "POST",
      headers: await headers({ "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" }),
      body: JSON.stringify(Array.isArray(rows) ? rows : [rows]),
    });
    if (!res.ok) throw new Error(`upsert ${table}: ${res.status} ${await res.text()}`);
  }

  async function remove(table, query) {
    const res = await fetch(REST + table + `?${query}`, {
      method: "DELETE",
      headers: await headers({ Prefer: "return=minimal" }),
    });
    if (!res.ok) throw new Error(`delete ${table}: ${res.status} ${await res.text()}`);
  }

  // ---- storage (private bucket; all access uses the login token) -----------
  async function uploadPhoto(path, blob) {
    const res = await fetch(OBJ + BUCKET + "/" + path, {
      method: "POST",
      headers: await headers({ "Content-Type": blob.type || "image/jpeg", "x-upsert": "true" }),
      body: blob,
    });
    if (!res.ok) throw new Error(`upload: ${res.status} ${await res.text()}`);
  }

  async function downloadPhoto(path) {
    const res = await fetch(OBJ + BUCKET + "/" + path, { headers: await headers() });
    if (!res.ok) throw new Error(`download ${path}: ${res.status}`);
    return res.blob();
  }

  async function deletePhotos(paths) {
    if (!paths.length) return;
    const res = await fetch(`${SUPABASE_URL}/storage/v1/object/${BUCKET}`, {
      method: "DELETE",
      headers: await headers({ "Content-Type": "application/json" }),
      body: JSON.stringify({ prefixes: paths }),
    });
    if (!res.ok) throw new Error(`delete photos: ${res.status} ${await res.text()}`);
  }

  // Invoke an Edge Function with the current session.
  async function invoke(name, payload) {
    const res = await fetch(`${SUPABASE_URL}/functions/v1/${name}`, {
      method: "POST",
      headers: await headers({ "Content-Type": "application/json" }),
      body: JSON.stringify(payload || {}),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || `Function ${name}: ${res.status}`);
    return data;
  }

  // True if we can produce a usable access token (refreshing if needed).
  const ensureSession = async () => !!(await token());

  return { login, logout, refresh, isLoggedIn, ensureSession, currentUser, select, upsert, remove, uploadPhoto, downloadPhoto, deletePhotos, invoke };
})();
