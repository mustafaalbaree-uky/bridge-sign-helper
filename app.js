// Bridge Sign Helper — full workflow.
//   Setup  (computer): import the Excel sheet -> Supabase `signs`.
//   Signs  (phone):    pick a sign, shoot up to two photos -> local queue -> Supabase.
//   Review (computer): pull the batch from Supabase, export named files, notify.

(() => {
  "use strict";

  const App = {
    signs: [],
    online: false,
    position: null,
    search: "",
    currentSignId: null,
    screen: "signs", // signs | capture | review | setup
    localCaptures: [],
    remoteCaptures: [],
    recipients: [],
    settings: {}, // key/value app settings (e.g. email webhook)
    captures: [], // cached server captures (for counts + photos-on-file)
    photoCounts: {}, // signId -> number of photos on the server
    signDone: {}, // signId -> true if it has photos and all are saved+emailed
    selectMode: false, // signs-list multi-select for deletion
    selectedSigns: new Set(),
    isMobile: false,
    parsed: null, // staged Excel import { type, rows, sheetNames, sheet }
  };

  // ---- helpers ---------------------------------------------------------------
  const el = (id) => document.getElementById(id);
  const esc = (s) =>
    String(s == null ? "" : s).replace(/[&<>"']/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
    );

  const toRad = (d) => (d * Math.PI) / 180;
  function distanceMeters(a, b) {
    const R = 6371000;
    const dLat = toRad(b.lat - a.lat), dLng = toRad(b.lng - a.lng);
    const la1 = toRad(a.lat), la2 = toRad(b.lat);
    const h = Math.sin(dLat / 2) ** 2 + Math.cos(la1) * Math.cos(la2) * Math.sin(dLng / 2) ** 2;
    return 2 * R * Math.asin(Math.sqrt(h));
  }
  const fmtDistance = (m) => {
    const mi = m / 1609.344;
    return mi < 0.19 ? `${Math.round(m * 3.28084)} ft` : `${mi.toFixed(mi < 10 ? 2 : 1)} mi`;
  };

  function todayStr() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(
      d.getDate()
    ).padStart(2, "0")}`;
  }
  function yymmddFromDate(dateStr) {
    const [y, m, d] = dateStr.split("-");
    return `${y.slice(2)}${m}${d}`;
  }
  const extFromPath = (p) => {
    const m = /\.([a-z0-9]+)$/i.exec(p || "");
    return m ? m[1] : "jpg";
  };

  // Object URL for a blob. Each <img data-revoke> frees its own URL on load,
  // so we never revoke a URL an on-screen image is still using (that was the
  // "image shows its alt text" bug).
  function preview(blob) {
    return URL.createObjectURL(blob);
  }
  function wireThumbs(root) {
    (root || document).querySelectorAll("img[data-revoke]").forEach((img) => {
      if (img.dataset.wired) return;
      img.dataset.wired = "1";
      const free = () => URL.revokeObjectURL(img.src);
      img.addEventListener("load", free);
      img.addEventListener("error", free);
    });
  }
  let statusTimer = null;
  // Transient toast. Auto-hides after `ms` (pass 0 to keep it until replaced).
  function setStatus(msg, kind, ms = 5000) {
    const s = el("statusBar");
    if (!s) return;
    clearTimeout(statusTimer);
    s.textContent = msg || "";
    s.hidden = !msg;
    s.className = "status" + (kind ? " " + kind : "");
    if (msg && ms) statusTimer = setTimeout(() => { s.hidden = true; s.textContent = ""; }, ms);
  }

  // ---- data loading ----------------------------------------------------------
  async function loadSigns() {
    try {
      const rows = await SB.select("signs", "select=*&order=county.asc,route.asc,mile_point.asc");
      App.signs = rows;
      App.online = true;
      await DB.putSigns(rows);
    } catch {
      App.online = false;
      const cached = await DB.allSigns();
      App.signs = cached.length ? cached : window.SEED_SIGNS || [];
    }
  }

  async function loadLocalCaptures() {
    App.localCaptures = await DB.allCaptures();
  }

  // Cache every capture once. Feeds both the signs-list counts and the
  // "Photos on file" section (so both are instant, no per-click network wait).
  // Returns true if the per-sign counts changed.
  // Rebuild per-sign counts and completion from the capture cache.
  function recomputeDerived() {
    const counts = {}, done = {}, any = {};
    for (const r of App.captures) {
      counts[r.sign_id] = (counts[r.sign_id] || 0) + 1;
      any[r.sign_id] = true;
      const complete = !!(r.exported_at && r.emailed_at);
      done[r.sign_id] = done[r.sign_id] === undefined ? complete : done[r.sign_id] && complete;
    }
    App.photoCounts = counts;
    App.signDone = {};
    for (const id in any) App.signDone[id] = !!done[id];
  }

  async function refreshCaptures() {
    try {
      const rows = await SB.select(
        "captures",
        "select=sign_id,slot,batch_date,storage_path,captured_at,exported_at,emailed_at&order=batch_date.desc,slot.asc"
      );
      // Guard against a transient empty result wiping the badges.
      if (!rows.length && App.captures.length) return false;
      App.captures = rows;
      const before = JSON.stringify(App.photoCounts);
      recomputeDerived();
      return JSON.stringify(App.photoCounts) !== before;
    } catch {
      return false; // keep whatever we had
    }
  }

  async function loadRecipients() {
    try {
      const rows = await SB.select("recipients", "select=email&order=added_at.desc");
      App.recipients = rows.map((r) => r.email);
    } catch {
      App.recipients = [];
    }
  }

  async function loadSettings() {
    try {
      const rows = await SB.select("settings", "select=key,value");
      App.settings = Object.fromEntries(rows.map((r) => [r.key, r.value]));
    } catch {
      App.settings = {};
    }
  }

  // Is automatic sending turned on? (The actual method — Gmail SMTP or Apps
  // Script — lives in server-side secrets, invisible to the browser.)
  function emailConfigured() {
    const s = App.settings || {};
    if (s.email_enabled === "true") return true;
    if (s.email_enabled === "false") return false;
    return !!s.email_webhook_url; // legacy default
  }

  // The "Compose email" fallback button can be turned off in Developer settings.
  function composeEnabled() {
    return App.settings.compose_enabled !== "false"; // default on
  }

  // ---- geolocation -----------------------------------------------------------
  function requestLocation() {
    const btn = el("locateBtn");
    if (!navigator.geolocation) return setStatus("This device can't share location.");
    if (btn) { btn.disabled = true; btn.textContent = "Locating…"; }
    navigator.geolocation.getCurrentPosition(
      (pos) => {
        App.position = { lat: pos.coords.latitude, lng: pos.coords.longitude, accuracy: pos.coords.accuracy };
        render();
      },
      (err) => {
        setStatus(`Location unavailable (${err.message}). Search still works.`);
        if (btn) { btn.disabled = false; btn.textContent = "📍 Sort by nearest"; }
      },
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 30000 }
    );
  }

  // ---- capture + sync --------------------------------------------------------
  function localToday() {
    const t = todayStr();
    return App.localCaptures.filter((c) => c.batchDate === t);
  }

  async function onPhoto(signId, slot, file) {
    if (!file) return;
    const batchDate = todayStr();
    const ext = (file.type === "image/png" && "png") || "jpg";
    const storagePath = `${signId}/${batchDate}__${slot}.${ext}`;
    const rec = {
      key: `${signId}__${slot}__${batchDate}`,
      signId, slot, batchDate,
      blob: file,
      capturedAt: new Date().toISOString(),
      captureLat: App.position ? App.position.lat : null,
      captureLng: App.position ? App.position.lng : null,
      storagePath,
      status: "pending",
      error: null,
    };
    await DB.putCapture(rec);
    await loadLocalCaptures();
    render();
    syncCapture(rec); // fire-and-forget; UI reflects status on completion
  }

  async function syncCapture(rec) {
    try {
      await SB.uploadPhoto(rec.storagePath, rec.blob);
      await SB.upsert(
        "captures",
        {
          sign_id: rec.signId,
          slot: rec.slot,
          batch_date: rec.batchDate,
          storage_path: rec.storagePath,
          captured_at: rec.capturedAt,
          capture_lat: rec.captureLat,
          capture_lng: rec.captureLng,
        },
        "sign_id,slot,batch_date"
      );
      rec.status = "synced";
      rec.error = null;
      refreshCaptures(); // keep the server cache / counts current
    } catch (e) {
      rec.status = "error";
      rec.error = String(e.message || e);
    }
    await DB.putCapture(rec);
    await loadLocalCaptures();
    render();
  }

  async function syncAllPending() {
    const pending = App.localCaptures.filter((c) => c.status !== "synced");
    setStatus(`Uploading ${pending.length} photo(s)…`);
    for (const rec of pending) await syncCapture(rec);
    const stillBad = App.localCaptures.filter((c) => c.status !== "synced").length;
    setStatus(stillBad ? `${stillBad} still not uploaded. Try again when you have signal.` : "All photos uploaded.", stillBad ? "warn" : "ok");
  }

  async function removePhoto(signId, slot) {
    const batchDate = todayStr();
    await DB.removeCapture(`${signId}__${slot}__${batchDate}`);
    try {
      await SB.remove("captures", `sign_id=eq.${signId}&slot=eq.${slot}&batch_date=eq.${batchDate}`);
      await refreshCaptures();
    } catch { /* offline: server copy (if any) cleared on next flush review */ }
    await loadLocalCaptures();
    render();
  }

  // ---- rendering: chrome -----------------------------------------------------
  function render() {
    for (const [id, scr] of [["navSetup", "setup"], ["navSigns", "signs"], ["navReview", "review"]]) {
      const active = scr === "signs" ? App.screen === "signs" || App.screen === "capture" : App.screen === scr;
      el(id).classList.toggle("active", active);
    }
    const pending = App.localCaptures.filter((c) => c.status !== "synced").length;
    const dot = el("syncDot");
    dot.hidden = pending === 0;
    dot.textContent = pending || "";

    if (App.screen === "setup") renderSetup();
    else if (App.screen === "capture") renderCapture();
    else if (App.screen === "review") renderReview();
    else renderSigns();
  }

  // Go to the signs list and refresh photo counts from the server in the
  // background (badge updates once they arrive).
  function goSigns() {
    editing = false;
    App.screen = "signs";
    render();
    // Refresh in the background; only re-render if counts actually changed
    // (avoids the badge flicker when nothing has changed).
    refreshCaptures().then((changed) => { if (changed && App.screen === "signs") render(); });
  }

  // ---- rendering: signs list -------------------------------------------------
  // Build the shell (search box + list container) ONCE. Typing updates only the
  // list, so the input element is never recreated — which is what was resetting
  // the iOS number keyboard back to letters on every keystroke.
  function renderSigns() {
    const conn = App.online
      ? ""
      : `<div class="banner warn">Offline. Showing the last synced sign list; photos upload when you're back on signal.</div>`;
    const note = App.position
      ? `<div class="hint">Nearest first · GPS ±${Math.round(App.position.accuracy)} m. <strong>Confirm the ID by eye.</strong></div>`
      : "";

    el("view").innerHTML = `
      ${conn}
      <div class="toolbar">
        <input id="searchInput" class="search" type="text" inputmode="text" autocapitalize="characters"
          autocomplete="off" placeholder="Search ID, route, county…" value="${esc(App.search)}" />
        <button id="locateBtn" class="btn secondary">📍 Sort by nearest</button>
      </div>${note}
      ${App.selectMode
        ? `<div class="sel-bar">
             <button id="selCompleted" class="btn small secondary">Select completed</button>
             <span id="signSelSummary" class="sel-summary"></span>
             <button id="signDelete" class="btn small danger">Delete</button>
             <button id="signCancel" class="btn small secondary">Cancel</button>
           </div>`
        : `<div class="sub-actions"><button id="enterSelect" class="btn link">Select signs to delete…</button></div>`}
      <ul id="signList" class="sign-list"></ul>`;

    renderSignList();
    updateSignSelSummary();
    el("searchInput").addEventListener("input", (e) => { App.search = e.target.value; renderSignList(); });
    el("locateBtn").addEventListener("click", requestLocation);
    if (App.selectMode) {
      el("selCompleted").addEventListener("click", () => {
        App.signs.forEach((s) => { if (App.signDone[s.id]) App.selectedSigns.add(s.id); });
        renderSignList();
        updateSignSelSummary();
      });
      el("signCancel").addEventListener("click", () => { App.selectMode = false; App.selectedSigns = new Set(); renderSigns(); });
      el("signDelete").addEventListener("click", () => {
        const ids = [...App.selectedSigns];
        if (!ids.length) return setStatus("Select at least one sign.", "warn");
        showSignDeleteConfirm(ids);
      });
    } else {
      el("enterSelect").addEventListener("click", () => { App.selectMode = true; App.selectedSigns = new Set(); renderSigns(); });
    }
  }

  function updateSignSelSummary() {
    const s = el("signSelSummary");
    if (s) s.textContent = `${App.selectedSigns.size} selected`;
  }

  function renderSignList() {
    const ul = el("signList");
    if (!ul) return;
    const q = App.search.trim().toLowerCase();
    let list = App.signs.filter((s) => (s.active_status || "Active") !== "Inactive");
    if (q)
      list = list.filter((s) =>
        [s.id, s.county, s.route, s.direction, s.side_of_road].join(" ").toLowerCase().includes(q)
      );
    if (App.position)
      list = list
        .filter((s) => s.lat != null && s.lng != null)
        .map((s) => Object.assign({}, s, { _d: distanceMeters(App.position, s) }))
        .sort((a, b) => a._d - b._d)
        .concat(list.filter((s) => s.lat == null || s.lng == null));

    ul.innerHTML =
      list
        .map((s) => {
          const localN = App.localCaptures.filter((c) => c.signId === s.id).length;
          const caps = Math.max(App.photoCounts[s.id] || 0, localN);
          const badge = caps ? `<span class="badge done">${caps} Photo${caps > 1 ? "s" : ""}</span>` : "";
          const doneChip = App.signDone[s.id] ? `<span class="chip ok">✓ Done</span>` : "";
          const dist = s._d != null ? `<span class="dist">${fmtDistance(s._d)}</span>` : "";
          const cb = App.selectMode ? `<input type="checkbox" class="sign-cb" ${App.selectedSigns.has(s.id) ? "checked" : ""} />` : "";
          return `<li class="sign-item ${App.selectMode ? "selecting" : ""}" data-id="${esc(s.id)}">
              ${cb}
              <div class="sign-main">
                <div class="sign-id">${esc(s.id)} ${badge}${doneChip}</div>
                <div class="sign-sub">${esc(s.county)} · ${esc(s.route)} · MP ${esc(s.mile_point)} · ${esc(s.direction)} · ${esc(s.side_of_road)}</div>
              </div>${dist}
            </li>`;
        })
        .join("") || `<li class="empty">No signs. Import a sheet under <strong>Setup</strong>.</li>`;

    ul.querySelectorAll(".sign-item").forEach((li) =>
      li.addEventListener("click", () => {
        const id = li.dataset.id;
        if (App.selectMode) {
          if (App.selectedSigns.has(id)) App.selectedSigns.delete(id);
          else App.selectedSigns.add(id);
          const cb = li.querySelector(".sign-cb");
          if (cb) cb.checked = App.selectedSigns.has(id);
          updateSignSelSummary();
        } else {
          App.currentSignId = id;
          App.screen = "capture";
          render();
        }
      })
    );
  }

  // ---- rendering: capture ----------------------------------------------------
  function renderCapture() {
    const sign = App.signs.find((s) => s.id === App.currentSignId);
    if (!sign) { App.screen = "signs"; return render(); }
    const today = todayStr();
    const slotRec = (n) => App.localCaptures.find((c) => c.signId === sign.id && c.slot === n && c.batchDate === today);

    const statusChip = (r) =>
      r.status === "synced" ? `<span class="chip ok">Uploaded</span>`
      : r.status === "error" ? `<span class="chip err" title="${esc(r.error || "")}">Not uploaded</span>`
      : `<span class="chip">Uploading…</span>`;

    const slotHtml = (n) => {
      const r = slotRec(n);
      if (r)
        return `<div class="slot filled">
            <img src="${preview(r.blob)}" alt="Photo ${n}" data-revoke />
            <div class="slot-bar">${statusChip(r)}
              <label class="btn small block">Retake<input type="file" accept="image/*" capture="environment" data-slot="${n}" hidden /></label>
              <button class="btn small danger block" data-remove="${n}">Remove</button>
            </div></div>`;
      return `<label class="slot empty"><span class="slot-plus">＋</span><span>Photo ${n}</span>
          <input type="file" accept="image/*" capture="environment" data-slot="${n}" hidden /></label>`;
    };

    const hasCoords = sign.lat != null && sign.lng != null;
    const mapsBtn = hasCoords
      ? `<a class="btn secondary block" target="_blank" rel="noopener"
           href="https://www.google.com/maps/search/?api=1&query=${sign.lat},${sign.lng}">Open in Google Maps</a>`
      : "";

    el("view").innerHTML = `
      <button id="backBtn" class="btn link">‹ All signs</button>
      <div class="detail-card">
        <div class="detail-id">${esc(sign.id)}</div>
        <div class="detail-grid">
          <div><span>County</span>${esc(sign.county)}</div>
          <div><span>Route</span>${esc(sign.route)}</div>
          <div><span>Mile point</span>${esc(sign.mile_point)}</div>
          <div><span>Direction</span>${esc(sign.direction)}</div>
          <div><span>Side</span>${esc(sign.side_of_road)}</div>
          <div><span>Coordinates</span>${esc(sign.lat)}, ${esc(sign.lng)}</div>
        </div>
        ${mapsBtn}
      </div>
      <div class="confirm-note">Make sure the ID above matches the sign in front of you.</div>
      <div class="slots">${slotHtml(1)}${slotHtml(2)}</div>
      <button id="doneBtn" class="btn primary block">Done, back to list</button>
      <div id="onFile" class="on-file"></div>
      <button id="editToggle" class="btn link">${editing ? "Cancel editing" : "Edit sign details"}</button>
      ${editing ? editForm(sign) : ""}`;

    el("backBtn").addEventListener("click", goSigns);
    el("doneBtn").addEventListener("click", () => {
      goSigns();
      if (App.localCaptures.some((c) => c.status !== "synced")) syncAllPending();
    });
    el("view").querySelectorAll('input[type="file"]').forEach((inp) =>
      inp.addEventListener("change", (e) => onPhoto(sign.id, Number(inp.dataset.slot), e.target.files[0]))
    );
    el("view").querySelectorAll("[data-remove]").forEach((b) =>
      b.addEventListener("click", () => removePhoto(sign.id, Number(b.dataset.remove)))
    );
    wireThumbs(el("view"));
    loadOnFile(sign.id);
    el("editToggle").addEventListener("click", () => { editing = !editing; render(); });
    if (editing) wireEditForm(sign);
  }

  let editing = false;

  // Show photos already in the database for this sign. The frames and count
  // render instantly from the cached capture list; the images themselves
  // stream in afterward, so it's obvious right away that photos exist.
  function loadOnFile(signId) {
    const box = el("onFile");
    if (!box) return;
    const rows = App.captures.filter((c) => c.sign_id === signId);
    if (!rows.length) { box.innerHTML = ""; return; }

    const byDate = {};
    for (const r of rows) (byDate[r.batch_date] = byDate[r.batch_date] || []).push(r);
    const dates = Object.keys(byDate).sort().reverse();

    box.innerHTML =
      `<h3 class="on-file-h">Photos on file (${rows.length})</h3>` +
      dates
        .map(
          (d) =>
            `<div class="on-file-day"><div class="on-file-date">${esc(d)}</div>
               <div class="on-file-thumbs">${byDate[d]
                 .map((r) => `<a class="of-thumb" data-path="${esc(r.storage_path)}"><span class="of-spin"></span></a>`)
                 .join("")}</div></div>`
        )
        .join("");

    // Download each image via the login token and swap it in as it arrives.
    box.querySelectorAll(".of-thumb").forEach(async (a) => {
      try {
        const blob = await SB.downloadPhoto(a.dataset.path);
        if (App.currentSignId !== signId || App.screen !== "capture") return;
        const url = preview(blob); // kept alive: the anchor opens the full image
        a.href = url;
        a.target = "_blank";
        a.rel = "noopener";
        a.innerHTML = `<img src="${url}" alt="" />`;
      } catch {
        a.innerHTML = `<span class="of-fail">?</span>`;
      }
    });
  }

  function editForm(sign) {
    const f = (label, name, val, type) =>
      `<label class="field"><span>${label}</span>
         <input class="ef" data-name="${name}" type="${type || "text"}" value="${esc(val == null ? "" : val)}" /></label>`;
    return `<div class="setup-card edit-card">
        <div class="banner warn">Editing changes the app's copy only. The next Excel import can overwrite it. Make sure that's what you want.</div>
        ${f("County", "county", sign.county)}
        ${f("Route", "route", sign.route)}
        ${f("Section", "section", sign.section)}
        ${f("Direction", "direction", sign.direction)}
        ${f("Mile point", "mile_point", sign.mile_point, "number")}
        ${f("Side of road", "side_of_road", sign.side_of_road)}
        ${f("Latitude", "lat", sign.lat, "number")}
        ${f("Longitude", "lng", sign.lng, "number")}
        <button id="saveEdit" class="btn primary block">Save changes</button>
        <button id="deleteSign" class="btn danger block">Delete this sign</button>
      </div>`;
  }

  function wireEditForm(sign) {
    el("saveEdit").addEventListener("click", async () => {
      const patch = { id: sign.id, updated_at: new Date().toISOString() };
      el("view").querySelectorAll(".ef").forEach((inp) => {
        const name = inp.dataset.name;
        let v = inp.value.trim();
        if (name === "mile_point" || name === "lat" || name === "lng") v = v === "" ? null : parseFloat(v);
        patch[name] = v;
      });
      try {
        await SB.upsert("signs", patch, "id");
        editing = false;
        await loadSigns();
        setStatus("Sign updated.", "ok");
        render();
      } catch (e) { setStatus(`Update failed: ${e.message}`, "warn"); }
    });
    el("deleteSign").addEventListener("click", async () => {
      if (!confirm(`Delete ${sign.id}? This removes it from the app (not from the Excel sheet).`)) return;
      try {
        await SB.remove("signs", `id=eq.${encodeURIComponent(sign.id)}`);
        editing = false;
        App.screen = "signs";
        App.currentSignId = null;
        await loadSigns();
        setStatus("Sign deleted.", "ok");
        render();
      } catch (e) { setStatus(`Delete failed: ${e.message}`, "warn"); }
    });
  }

  // ---- rendering: review / flush --------------------------------------------
  // Returns file objects tagged with their capture date, correctly named.
  function buildExportList(remoteRows) {
    const groups = {};
    for (const c of remoteRows) {
      const g = `${c.sign_id}|${c.batch_date}`;
      (groups[g] = groups[g] || []).push(c);
    }
    const out = [];
    for (const g of Object.keys(groups)) {
      const [signId, batch] = g.split("|");
      const items = groups[g].sort((a, b) => a.slot - b.slot);
      const two = items.length > 1;
      items.forEach((c, i) => {
        out.push({
          filename: `${signId}-${yymmddFromDate(batch)}${two ? i + 1 : ""}.${extFromPath(c.storage_path)}`,
          storage_path: c.storage_path,
          captured_at: c.captured_at,
          batch_date: batch,
          exported_at: c.exported_at,
          emailed_at: c.emailed_at,
        });
      });
    }
    return out.sort((a, b) => a.filename.localeCompare(b.filename));
  }

  function prettyDate(dateStr) {
    const [y, m, d] = dateStr.split("-").map(Number);
    return new Date(y, m - 1, d).toLocaleDateString(undefined, {
      weekday: "short", month: "short", day: "numeric", year: "numeric",
    });
  }

  // Which day-groups are ticked in Review. null = "select all" on next open.
  let reviewSelected = null;

  async function renderReview() {
    el("view").innerHTML = `<p class="hint">Loading…</p>`;
    let remote = [];
    try {
      remote = await SB.select("captures", `select=*&order=batch_date.desc,sign_id.asc,slot.asc`);
    } catch {
      el("view").innerHTML = `<div class="banner warn">Can't reach the server. Connect to load photos for export.</div>`;
      return;
    }
    App.remoteCaptures = remote;
    const files = buildExportList(remote);
    const pending = App.localCaptures.filter((c) => c.status !== "synced").length;
    const fsa = typeof window.showDirectoryPicker === "function";
    const emailReady = emailConfigured();
    const today = todayStr();

    const thumbs = {};
    await Promise.all(
      files.map(async (f) => {
        try { thumbs[f.storage_path] = preview(await SB.downloadPhoto(f.storage_path)); } catch {}
      })
    );

    // Group files by capture date, newest day first.
    const byDate = new Map();
    for (const f of files) {
      if (!byDate.has(f.batch_date)) byDate.set(f.batch_date, []);
      byDate.get(f.batch_date).push(f);
    }
    const dates = [...byDate.keys()].sort().reverse();
    const allPaths = new Set(files.map((f) => f.storage_path));
    if (reviewSelected === null) reviewSelected = new Set(allPaths); // default: all
    else for (const p of [...reviewSelected]) if (!allPaths.has(p)) reviewSelected.delete(p);
    const selectedFiles = () => files.filter((f) => reviewSelected.has(f.storage_path));

    // How many photos of a day are selected: none / some / all.
    const daySelState = (date) => {
      const ps = byDate.get(date).map((f) => f.storage_path);
      const n = ps.filter((p) => reviewSelected.has(p)).length;
      return n === 0 ? "none" : n === ps.length ? "all" : "some";
    };

    const isDone = (f) => f.exported_at && f.emailed_at;
    const statusPills = (f) =>
      (f.exported_at ? `<span class="chip ok">Saved</span>` : `<span class="chip">Not yet saved</span>`) +
      (f.emailed_at ? `<span class="chip ok">Emailed</span>` : `<span class="chip">Not yet emailed</span>`);

    // perPhoto=true shows each photo's own status + a Delete pill (only when done).
    const fileRow = (f, perPhoto) => `<li class="file-row">
        <input type="checkbox" class="photo-cb" data-path="${esc(f.storage_path)}" data-date="${f.batch_date}" ${reviewSelected.has(f.storage_path) ? "checked" : ""} />
        ${thumbs[f.storage_path] ? `<img src="${thumbs[f.storage_path]}" alt="" data-revoke />` : `<div class="thumb-missing">?</div>`}
        <div class="file-meta">
          <div class="file-name">${esc(f.filename)}</div>
          <div class="file-sub">${new Date(f.captured_at).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}</div>
          ${perPhoto ? `<div class="pill-row">${statusPills(f)}${isDone(f) ? `<button class="chip del" data-del-photo="${esc(f.storage_path)}">Delete</button>` : ""}</div>` : ""}
        </div>
      </li>`;

    const dayBlock = (date) => {
      const df = byDate.get(date);
      const label = date === today ? "Today" : prettyDate(date);
      const nSaved = df.filter((f) => f.exported_at).length;
      const nEmailed = df.filter((f) => f.emailed_at).length;
      const dayDone = nSaved === df.length && nEmailed === df.length;
      const dayFresh = nSaved === 0 && nEmailed === 0;
      const perPhoto = !dayDone && !dayFresh; // mixed → per-photo detail
      const flags = dayDone
        ? `<span class="chip ok">Saved</span><span class="chip ok">Emailed</span><button class="chip del" data-del-day="${date}">Delete</button>`
        : dayFresh
          ? `<span class="chip">Not yet saved</span><span class="chip">Not yet emailed</span>`
          : "";
      return `<section class="day-group">
          <div class="day-head">
            <label class="day-check"><input type="checkbox" class="day-cb" data-date="${date}" ${daySelState(date) === "all" ? "checked" : ""} />
              <span class="day-label">${label}</span></label>
            <span class="day-flags">${flags}</span>
            <span class="day-count">${df.length} photo${df.length > 1 ? "s" : ""}</span></div>
          <ul class="file-list">${df.map((f) => fileRow(f, perPhoto)).join("")}</ul>
        </section>`;
    };

    const chips = App.recipients
      .map((e) => `<button type="button" class="recip-chip" data-email="${esc(e)}">${esc(e)}</button>`)
      .join("");

    el("view").innerHTML = `
      <h2 class="screen-title">Review &amp; export</h2>
      <p class="hint">Every inspection photo on the server, newest day first. Photos stay here until you delete them; exporting or emailing does not remove them. Files are named <code>ID-YYMMDD</code>.</p>
      ${pending ? `<div class="banner warn">${pending} photo(s) on this device haven't uploaded. <button id="syncNow" class="btn small">Sync now</button></div>` : ""}
      ${files.length ? `
        <div class="notify">
          <label class="fieldlabel" for="emailInput">Email recipient</label>
          <input id="emailInput" class="search" type="email"
            name="bsh-notify-recipient" autocomplete="off" autocapitalize="off"
            autocorrect="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true"
            data-form-type="other" placeholder="engineer@example.com" />
          ${chips ? `<div class="recip-chips"><span class="recip-label">Saved:</span>${chips}</div>` : ""}
          <div class="sel-bar">
            <span id="selSummary" class="sel-summary"></span>
            <button id="selAll" class="btn small secondary">Select all</button>
            <button id="selNone" class="btn small secondary">Clear</button>
          </div>
          <div class="day-actions">
            ${fsa ? `<button id="actFolder" class="btn small primary">Save to folder</button>` : ""}
            <button id="actDownload" class="btn small secondary">Download</button>
            ${emailReady ? `<button id="actSend" class="btn small secondary">Send email</button>` : ""}
            ${composeEnabled() ? `<button id="actCompose" class="btn small secondary">Compose email</button>` : ""}
          </div>
          <p class="hint">Tick individual photos or a whole day, then these buttons act on everything selected.</p>
        </div>
        ${dates.map(dayBlock).join("")}
        ${App.localCaptures.length ? `<button id="clearLocal" class="btn danger block">Clear photos saved on this device (${App.localCaptures.length})</button>` : ""}
      ` : `<p class="empty">No photos on the server yet.</p>`}`;

    wireThumbs(el("view"));
    if (pending) el("syncNow").addEventListener("click", syncAllPending);
    if (!files.length) return;

    const updateSummary = () => {
      const sel = selectedFiles();
      const days = new Set(sel.map((f) => f.batch_date)).size;
      el("selSummary").textContent =
        `${sel.length} photo${sel.length !== 1 ? "s" : ""} selected` + (days > 1 ? ` (${days} days)` : "");
    };
    const syncDayCb = (date) => {
      const cb = el("view").querySelector(`.day-cb[data-date="${date}"]`);
      if (!cb) return;
      const st = daySelState(date);
      cb.checked = st === "all";
      cb.indeterminate = st === "some";
    };
    dates.forEach(syncDayCb); // set initial indeterminate states
    updateSummary();

    el("view").querySelectorAll(".photo-cb").forEach((cb) =>
      cb.addEventListener("change", () => {
        if (cb.checked) reviewSelected.add(cb.dataset.path);
        else reviewSelected.delete(cb.dataset.path);
        syncDayCb(cb.dataset.date);
        updateSummary();
      })
    );
    el("view").querySelectorAll(".day-cb").forEach((cb) =>
      cb.addEventListener("change", () => {
        const ps = byDate.get(cb.dataset.date).map((f) => f.storage_path);
        ps.forEach((p) => (cb.checked ? reviewSelected.add(p) : reviewSelected.delete(p)));
        cb.indeterminate = false;
        el("view").querySelectorAll(`.photo-cb[data-date="${cb.dataset.date}"]`).forEach((pc) => (pc.checked = cb.checked));
        updateSummary();
      })
    );
    el("selAll").addEventListener("click", () => {
      reviewSelected = new Set(allPaths);
      el("view").querySelectorAll(".photo-cb").forEach((c) => (c.checked = true));
      dates.forEach(syncDayCb);
      updateSummary();
    });
    el("selNone").addEventListener("click", () => {
      reviewSelected.clear();
      el("view").querySelectorAll(".photo-cb, .day-cb").forEach((c) => { c.checked = false; c.indeterminate = false; });
      updateSummary();
    });

    const act = async (fn, kind, btn) => {
      const f = selectedFiles();
      if (!f.length) return setStatus("Select at least one photo first.", "warn");
      const ok = await fn(f, btn);
      if (ok !== false) {
        await markPhotos(f.map((x) => x.storage_path), kind);
        renderReview(); // refresh the Saved/Emailed badges
      }
    };
    if (fsa) el("actFolder").addEventListener("click", () => act(exportToFolder, "exported"));
    el("actDownload").addEventListener("click", () => act(downloadAll, "exported"));
    if (emailReady) el("actSend").addEventListener("click", (e) => act(sendEmail, "emailed", e.currentTarget));
    if (composeEnabled()) el("actCompose").addEventListener("click", () => act(composeEmail, "emailed"));

    el("view").querySelectorAll(".recip-chip").forEach((c) =>
      c.addEventListener("click", () => { el("emailInput").value = c.dataset.email; el("emailInput").focus(); })
    );
    el("view").querySelectorAll("[data-del-day]").forEach((b) =>
      b.addEventListener("click", () => showDeleteConfirm(byDate.get(b.dataset.delDay) || []))
    );
    el("view").querySelectorAll("[data-del-photo]").forEach((b) =>
      b.addEventListener("click", () => {
        const f = files.find((x) => x.storage_path === b.dataset.delPhoto);
        if (f) showDeleteConfirm([f]);
      })
    );
    if (App.localCaptures.length) el("clearLocal").addEventListener("click", clearLocalCopies);
  }

  // Confirmation dialog that lists exactly what will be deleted, with status.
  function showDeleteConfirm(files) {
    if (!files.length) return;
    const rows = files
      .map(
        (f) => `<li>
          <span class="mono">${esc(f.filename)}</span>
          ${f.exported_at ? `<span class="chip ok">Saved</span>` : `<span class="chip">Not yet saved</span>`}
          ${f.emailed_at ? `<span class="chip ok">Emailed</span>` : `<span class="chip">Not yet emailed</span>`}
        </li>`
      )
      .join("");
    const incomplete = files.filter((f) => !(f.exported_at && f.emailed_at)).length;
    const overlay = document.createElement("div");
    overlay.className = "modal-overlay";
    overlay.innerHTML = `<div class="modal">
        <h3>Delete ${files.length} photo${files.length > 1 ? "s" : ""} from the server?</h3>
        <p class="hint">${incomplete
          ? `<strong>Warning: ${incomplete} of these ${incomplete > 1 ? "are" : "is"} not fully saved and emailed.</strong> `
          : "All of these are saved and emailed. "}This permanently removes them and cannot be undone.</p>
        <ul class="del-list">${rows}</ul>
        <div class="modal-actions">
          <button class="btn secondary" data-mc="cancel">Cancel</button>
          <button class="btn danger" data-mc="ok">Delete ${files.length}</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    const close = () => overlay.remove();
    overlay.addEventListener("click", (e) => { if (e.target === overlay) close(); });
    overlay.querySelector('[data-mc="cancel"]').addEventListener("click", close);
    overlay.querySelector('[data-mc="ok"]').addEventListener("click", () => { close(); deleteFromServer(files); });
  }

  async function deleteFromServer(files) {
    const paths = files.map((f) => f.storage_path);
    setStatus(`Deleting ${paths.length} photo(s)…`, null, 0);
    try {
      try { await SB.deletePhotos(paths); } catch { /* storage object may already be gone */ }
      const list = paths.map((p) => encodeURIComponent(`"${p}"`)).join(",");
      await SB.remove("captures", `storage_path=in.(${list})`);
      // Update the cache immediately so counts/photos-on-file are right now,
      // not only after a page refresh.
      const del = new Set(paths);
      App.captures = App.captures.filter((c) => !del.has(c.storage_path));
      recomputeDerived();
      refreshCaptures(); // reconcile in the background
      setStatus(`Deleted ${paths.length} photo(s).`, "ok");
      renderReview();
    } catch (e) {
      setStatus(`Delete failed: ${e.message}`, "warn", 9000);
    }
  }

  // Confirm + delete whole signs (and their photos).
  function showSignDeleteConfirm(ids) {
    if (!ids.length) return;
    const rowFor = (id) => {
      const count = App.photoCounts[id] || 0;
      const status = count === 0
        ? `<span class="chip">no photos</span>`
        : App.signDone[id]
          ? `<span class="chip ok">completed</span>`
          : `<span class="chip">${count} photo${count > 1 ? "s" : ""}, not all saved &amp; emailed</span>`;
      return `<li><span class="mono">${esc(id)}</span>${status}</li>`;
    };
    const notDone = ids.filter((id) => (App.photoCounts[id] || 0) > 0 && !App.signDone[id]).length;
    const overlay = document.createElement("div");
    overlay.className = "modal-overlay";
    overlay.innerHTML = `<div class="modal">
        <h3>Delete ${ids.length} sign${ids.length > 1 ? "s" : ""}?</h3>
        <p class="hint">${notDone
          ? `<strong>Warning: ${notDone} of these still have photos that aren't saved &amp; emailed — those photos will be lost.</strong> `
          : ""}This removes the sign${ids.length > 1 ? "s" : ""} and any photos still on the server. The next Excel import can add them back. This can't be undone.</p>
        <ul class="del-list">${ids.map(rowFor).join("")}</ul>
        <div class="modal-actions">
          <button class="btn secondary" data-mc="cancel">Cancel</button>
          <button class="btn danger" data-mc="ok">Delete ${ids.length}</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    const close = () => overlay.remove();
    overlay.addEventListener("click", (e) => { if (e.target === overlay) close(); });
    overlay.querySelector('[data-mc="cancel"]').addEventListener("click", close);
    overlay.querySelector('[data-mc="ok"]').addEventListener("click", () => { close(); deleteSigns(ids); });
  }

  async function deleteSigns(ids) {
    setStatus(`Deleting ${ids.length} sign(s)…`, null, 0);
    try {
      const idSet = new Set(ids);
      const paths = App.captures.filter((c) => idSet.has(c.sign_id)).map((c) => c.storage_path);
      if (paths.length) { try { await SB.deletePhotos(paths); } catch { /* may be gone */ } }
      const idList = ids.map((i) => encodeURIComponent(`"${i}"`)).join(",");
      await SB.remove("captures", `sign_id=in.(${idList})`);
      await SB.remove("signs", `id=in.(${idList})`);
      App.selectMode = false;
      App.selectedSigns = new Set();
      await loadSigns();
      await refreshCaptures();
      setStatus(`Deleted ${ids.length} sign(s).`, "ok");
      render();
    } catch (e) {
      setStatus(`Delete failed: ${e.message}`, "warn", 9000);
    }
  }

  // Mark the given photos (by storage path) as exported or emailed.
  async function markPhotos(paths, kind) {
    if (!paths.length) return;
    const col = kind === "exported" ? "exported_at" : "emailed_at";
    const list = paths.map((p) => encodeURIComponent(`"${p}"`)).join(",");
    try {
      await SB.update("captures", `storage_path=in.(${list})`, { [col]: new Date().toISOString() });
    } catch { /* non-fatal: badge just won't update */ }
  }

  // Subject + body for a set of files that may span multiple days.
  function emailSubjectBody(files) {
    const dset = [...new Set(files.map((f) => f.batch_date))].sort();
    const subject = dset.length <= 1
      ? `Bridge sign photos, ${dset[0] || todayStr()}`
      : `Bridge sign photos, ${dset[0]} to ${dset[dset.length - 1]} (${dset.length} days)`;
    const sorted = [...files].sort((a, b) =>
      a.batch_date.localeCompare(b.batch_date) || a.filename.localeCompare(b.filename));
    const body = "The following sign inspection photos have been added:\n\n" +
      sorted.map((f) => f.filename).join("\n") + "\n";
    return { subject, body };
  }

  async function sendEmail(files, btn) {
    const to = (el("emailInput").value || "").trim();
    if (!to) { setStatus("Enter an email address first.", "warn"); return false; }
    const label = btn ? btn.textContent : "";
    if (btn) { btn.disabled = true; btn.textContent = "Sending…"; }
    const { subject, body } = emailSubjectBody(files);
    let ok = false;
    try {
      try { await SB.upsert("recipients", { email: to }, "email"); await loadRecipients(); } catch {}
      await SB.invoke("notify", { to, subject, body });
      setStatus(`Email sent to ${to}.`, "ok");
      ok = true;
    } catch (e) {
      setStatus(`Send failed: ${e.message}. You can use Compose instead.`, "warn", 9000);
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = label || "Send email"; }
    }
    return ok;
  }

  async function fetchBlob(path) {
    return SB.downloadPhoto(path);
  }

  async function exportToFolder(files) {
    try {
      const dir = await window.showDirectoryPicker({ mode: "readwrite" });
      let ok = 0;
      for (const f of files) {
        const blob = await fetchBlob(f.storage_path);
        const fh = await dir.getFileHandle(f.filename, { create: true });
        const w = await fh.createWritable();
        await w.write(blob);
        await w.close();
        ok++;
      }
      setStatus(`Saved ${ok} photo(s) to the chosen folder.`, "ok");
      return true;
    } catch (e) {
      if (e && e.name === "AbortError") return false;
      setStatus(`Could not save: ${e.message}`, "warn");
      return false;
    }
  }

  async function downloadAll(files) {
    for (let i = 0; i < files.length; i++) {
      try {
        const blob = await fetchBlob(files[i].storage_path);
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url; a.download = files[i].filename;
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1500);
      } catch {}
      await new Promise((r) => setTimeout(r, 250));
    }
    setStatus(`Downloaded ${files.length} photo(s).`, "ok");
    return true;
  }

  async function copyNames(files) {
    const text = files.map((f) => f.filename).join("\n");
    try { await navigator.clipboard.writeText(text); setStatus("Filename list copied.", "ok"); }
    catch { setStatus(text); }
  }

  async function composeEmail(files) {
    const to = (el("emailInput").value || "").trim();
    if (to) { try { await SB.upsert("recipients", { email: to }, "email"); await loadRecipients(); } catch {} }
    const { subject, body } = emailSubjectBody(files);
    window.location.href = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    return true;
  }

  async function clearLocalCopies() {
    if (!confirm("Remove photos stored on THIS device? They stay on the server. Make sure everything is uploaded first.")) return;
    await DB.clearCaptures();
    await loadLocalCaptures();
    setStatus("Local copies cleared.", "ok");
    render();
  }

  // ---- rendering: setup / Excel import --------------------------------------
  let pendingToken = "";
  const randToken = () => Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);

  function renderSetup() {
    const p = App.parsed;
    const emailReady = emailConfigured();
    const tokenVal = App.settings.email_webhook_token || pendingToken || (pendingToken = randToken());
    el("view").innerHTML = `
      <h2 class="screen-title">Setup: import signs</h2>
      <p class="hint">Load the Excel sheet on the computer. Existing IDs are updated; new ones are added. Do this again whenever the sheet changes.</p>
      <div class="setup-card">
        <label class="field"><span>Sign type</span>
          <select id="typeSel">
            <option value="advance">Advance weight-limit signs</option>
            <option value="bridge">Bridge weight-limit signs</option>
          </select>
        </label>
        <label class="btn secondary block">Choose Excel / CSV file…
          <input id="fileInput" type="file" accept=".xlsx,.xls,.csv" hidden /></label>
        ${workbookCache && workbookCache.SheetNames.length > 1 ? `
          <label class="field"><span>Sheet</span>
            <select id="sheetSel">${workbookCache.SheetNames.map((n) => `<option ${n === currentSheet ? "selected" : ""}>${esc(n)}</option>`).join("")}</select>
          </label>` : ""}
      </div>
      ${p ? renderParsedPreview(p) : ""}
      <div class="setup-card">
        <h3>Current database</h3>
        <p class="hint">${App.online ? `${App.signs.length} sign(s) loaded from the server.` : "Offline. Can't reach the server."}</p>
      </div>
      <div class="setup-card">
        <p class="hint">Signed in as <strong>${esc(SB.currentUser() || "")}</strong></p>
        <button id="logoutBtn" class="btn secondary block">Log out</button>
      </div>
      <details class="setup-card collapsible">
        <summary>Email notifications ${emailReady ? '<span class="chip ok">on</span>' : '<span class="chip">off</span>'}</summary>
        <label class="check-row"><input type="checkbox" id="emailEnabled" ${emailConfigured() ? "checked" : ""} />
          <span>Enable automatic sending (shows the “Send email” button in Review)</span></label>
        <p class="hint"><strong>Gmail (recommended):</strong> in the Supabase dashboard → Edge Functions → Secrets, add
          <code>GMAIL_USER</code> (the sending address) and <code>GMAIL_APP_PASSWORD</code> (a Google app password).
          Then tick the box above, Save, and send a test. See docs/EMAIL_SETUP.md.</p>
        <details class="sub">
          <summary>Alternative: Google Apps Script webhook</summary>
          <label class="field"><span>Apps Script Web App URL</span>
            <input id="whUrl" type="url" autocomplete="off" data-lpignore="true" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(App.settings.email_webhook_url || "")}" /></label>
          <label class="field"><span>Shared token (paste this same value into the script)</span>
            <input id="whToken" type="text" autocomplete="off" data-lpignore="true" value="${esc(tokenVal)}" /></label>
        </details>
        <button id="saveEmail" class="btn primary block">Save email settings</button>
        <button id="testEmail" class="btn secondary block">Send a test email</button>
      </details>
      <details class="setup-card collapsible">
        <summary>Developer settings</summary>
        <p class="hint">For testing only.</p>
        <label class="check-row"><input type="checkbox" id="composeToggle" ${composeEnabled() ? "checked" : ""} />
          <span>Show the “Compose email” button in Review</span></label>
        <p class="hint">Turn off if automatic “Send email” is working and the mail-app fallback isn't needed.</p>
        <button id="dlMock" class="btn secondary block">Download mock Excel sheet</button>
        <p class="hint">A sample sheet in the real R12-6 layout (split header, a duplicate ID, an inactive row, and a base-ID grouping row) to test importing.</p>
        <button id="clearAll" class="btn danger block">Clear signs &amp; photos</button>
        <p class="hint">Deletes all signs and photos from the database and this device. Kept: your login, the email-sending setup, and saved recipient emails.</p>
      </details>`;

    el("fileInput").addEventListener("change", (e) => handleFile(e.target.files[0]));
    el("typeSel").value = p ? p.type : "advance";
    el("typeSel").addEventListener("change", (e) => { if (App.parsed) App.parsed.type = e.target.value; });
    const sheetSel = el("sheetSel");
    if (sheetSel) sheetSel.addEventListener("change", (e) => reparseSheet(e.target.value));
    const imp = el("importBtn");
    if (imp) imp.addEventListener("click", doImport);

    async function saveEmailSettings() {
      await SB.upsert("settings", [
        { key: "email_enabled", value: el("emailEnabled").checked ? "true" : "false" },
        { key: "email_webhook_url", value: el("whUrl").value.trim() },
        { key: "email_webhook_token", value: el("whToken").value.trim() },
      ], "key");
      await loadSettings();
    }
    el("saveEmail").addEventListener("click", async () => {
      try { await saveEmailSettings(); setStatus("Email settings saved.", "ok"); render(); }
      catch (e) { setStatus(`Save failed: ${e.message}`, "warn"); }
    });
    el("testEmail").addEventListener("click", async () => {
      const to = prompt("Send a test email to which address?");
      if (!to) return;
      try {
        await saveEmailSettings();
        await SB.invoke("notify", { to, subject: "Bridge Sign Helper test", body: "Test email from Bridge Sign Helper. If you got this, notifications work." });
        setStatus(`Test email sent to ${to}.`, "ok");
      } catch (e) { setStatus(`Test failed: ${e.message}`, "warn", 9000); }
    });
    el("logoutBtn").addEventListener("click", doLogout);
    el("composeToggle").addEventListener("change", async (e) => {
      try {
        await SB.upsert("settings", { key: "compose_enabled", value: e.target.checked ? "true" : "false" }, "key");
        await loadSettings();
        setStatus(e.target.checked ? "Compose button shown." : "Compose button hidden.", "ok");
      } catch (err) { setStatus(`Save failed: ${err.message}`, "warn"); }
    });
    el("dlMock").addEventListener("click", downloadMockSheet);
    el("clearAll").addEventListener("click", clearAllData);
  }

  // Build a sample workbook in the real R12-6 layout and download it.
  function downloadMockSheet() {
    const aoa = [
      ["Active Status", "ID", "Assembly Location Information", "", "", "", "", "", ""],
      [],
      [],
      ["", "", "County", "Route", "Section", "Direction", "Mile Point", "Side of Road", "Lat, Long"],
      ["Active", "AW1052508121223", "Scott", "US-0025", "", "Decreasing", 20.656, "Right", "38.433738,-84.5661614"],
      ["Active", "AW1052508121227", "Scott", "KY-0620", "", "Increasing", 16.537, "Right", " 38.3322507,-84.51439005"],
      ["Active", "AW1052508121228", "Scott", "US-0025", "", "Increasing", 14.971, "Right", "38.3591820,-84.5639286"],
      ["Inactive", "AW1052508121229", "Scott", "US-0025", "", "Increasing", 12.0, "Right", "38.3000000,-84.5600000"],
      ["Active", "AW1202510140833", "Woodford", "US-0062", "", "Decreasing", 5.293, "Right", "38.043665,-84.758290"],
      ["Active", "AW1202510140833", "Woodford", "US-0062", "", "Decreasing", 5.293, "Right", "38.043665,-84.758290"],
      ["Active", "AW2604171416", "", "", "", "", "", "", ""],
      ["Active", "AW2604171416A01", "Fayette", "US-0027", "", "Increasing", 1.2, "Right", "38.100000,-84.500000"],
      ["Active", "AW2604171416B01", "Fayette", "US-0027", "", "Decreasing", 1.3, "Right", "38.200000,-84.600000"],
      ["Active", "AW260414155A02", "Madison", "KY-2877", "", "Increasing", 0.75, "Right", "37.6406962,-84.3159329"],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Master List R12-6");
    XLSX.writeFile(wb, "mock-signs-R12-6.xlsx");
  }

  async function clearAllData() {
    if (!confirm("Delete ALL signs and photos from the database and this device? This cannot be undone.\n\nYour login, email-sending setup, and saved recipient emails are kept.")) return;
    setStatus("Clearing signs and photos…", null, 0);
    try {
      let paths = [];
      try {
        paths = (await SB.select("captures", "select=storage_path")).map((r) => r.storage_path).filter(Boolean);
      } catch {}
      if (paths.length) { try { await SB.deletePhotos(paths); } catch {} }
      // Only signs and their photos. Left untouched: auth users (logins),
      // settings (email webhook), and recipients (saved colleague emails).
      await SB.remove("captures", "id=not.is.null");
      await SB.remove("signs", "id=not.is.null");
      await DB.clearCaptures();
      App.localCaptures = [];
      App.captures = [];
      App.photoCounts = {};
      await loadSigns();
      setStatus("Signs and photos cleared.", "ok");
      render();
    } catch (e) {
      setStatus(`Clear failed: ${e.message}`, "warn", 9000);
    }
  }

  function renderParsedPreview(p) {
    if (!p.rows.length)
      return `<div class="banner warn">No rows with an ID found on this sheet. Try another sheet.</div>`;
    const head = p.rows.slice(0, 5).map(
      (r) => `<tr><td>${esc(r.id)}</td><td>${esc(r.county)}</td><td>${esc(r.route)}</td><td>${esc(r.lat)}, ${esc(r.lng)}</td></tr>`
    ).join("");
    const bad = p.rows.filter((r) => r.lat == null || r.lng == null).length;
    return `<div class="setup-card">
        <h3>${p.rows.length} sign(s) found</h3>
        ${bad ? `<div class="banner warn">${bad} row(s) have no usable coordinates; they'll import but won't sort by distance.</div>` : ""}
        <div class="table-scroll"><table class="preview">
          <thead><tr><th>ID</th><th>County</th><th>Route</th><th>Lat, Long</th></tr></thead>
          <tbody>${head}</tbody></table></div>
        ${p.rows.length > 5 ? `<p class="hint">…and ${p.rows.length - 5} more.</p>` : ""}
        <button id="importBtn" class="btn primary block">Import ${p.rows.length} sign(s)</button>
      </div>`;
  }

  let workbookCache = null;
  let currentSheet = null;

  async function handleFile(file) {
    if (!file) return;
    setStatus("Reading file…");
    try {
      const buf = await file.arrayBuffer();
      workbookCache = XLSX.read(buf, { type: "array" });
      currentSheet = pickBestSheet(workbookCache);
      App.parsed = Object.assign(parseSheet(workbookCache, currentSheet), {
        type: el("typeSel") ? el("typeSel").value : "advance",
        sheet: currentSheet,
      });
      setStatus("");
      render();
    } catch (e) {
      setStatus(`Couldn't read that file: ${e.message}`, "warn");
    }
  }

  function reparseSheet(sheet) {
    if (!workbookCache) return;
    currentSheet = sheet;
    const type = App.parsed ? App.parsed.type : el("typeSel") ? el("typeSel").value : "advance";
    App.parsed = Object.assign(parseSheet(workbookCache, sheet), { type, sheet });
    render();
  }

  function pickBestSheet(wb) {
    for (const name of wb.SheetNames) if (parseSheet(wb, name).rows.length > 0) return name;
    return wb.SheetNames[0];
  }

  const norm = (v) => String(v == null ? "" : v).trim().toLowerCase();

  // The location columns (County/Route/.../Lat, Long) share a header row, which
  // may sit BELOW the row holding Active Status / ID. Find that row.
  function findGeoRow(aoa) {
    for (let i = 0; i < Math.min(aoa.length, 25); i++) {
      const hits = (aoa[i] || [])
        .map(norm)
        .filter((c) => /county|route|mile|side of road|lat.*long|long.*lat/.test(c)).length;
      if (hits >= 2) return i;
    }
    return -1;
  }

  // Map a header label to our field name.
  function classify(label) {
    const c = norm(label);
    if (!c) return null;
    if (c === "id" || c.endsWith(" id")) return "id";
    if (/active/.test(c)) return "active_status";
    if (/county/.test(c)) return "county";
    if (/route/.test(c)) return "route";
    if (/section/.test(c)) return "section";
    if (/direction/.test(c)) return "direction";
    if (/mile/.test(c)) return "mile_point";
    if (/side/.test(c)) return "side_of_road";
    if (/lat.*long|long.*lat|coordinate|lat, ?long/.test(c)) return "latlng";
    if (/^lat/.test(c)) return "lat";
    if (/^long|^lng/.test(c)) return "lng";
    return null;
  }

  function parseSheet(wb, sheetName) {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, blankrows: false });
    const geoRow = findGeoRow(aoa);
    if (geoRow < 0) return { rows: [] };
    // Header labels can be split across several rows (ID up top, the location
    // columns lower). Scan every header row (0..geoRow) and map each column.
    const map = {};
    for (let r = 0; r <= geoRow; r++) {
      (aoa[r] || []).forEach((label, idx) => {
        const f = classify(label);
        if (f && map[f] === undefined) map[f] = idx;
      });
    }
    if (map.id == null) return { rows: [] };
    const rows = [];
    for (let i = geoRow + 1; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const id = map.id != null ? String(row[map.id] == null ? "" : row[map.id]).trim() : "";
      if (!id) continue;
      let lat = null, lng = null;
      if (map.latlng != null && row[map.latlng] != null) {
        const parts = String(row[map.latlng]).split(",");
        lat = parseFloat(parts[0]);
        lng = parseFloat(parts[1]);
      }
      if (map.lat != null && row[map.lat] != null) lat = parseFloat(row[map.lat]);
      if (map.lng != null && row[map.lng] != null) lng = parseFloat(row[map.lng]);
      const rec = {
        id,
        active_status: map.active_status != null ? String(row[map.active_status] || "").trim() || "Active" : "Active",
        county: map.county != null ? String(row[map.county] || "").trim() : "",
        route: map.route != null ? String(row[map.route] || "").trim() : "",
        section: map.section != null ? String(row[map.section] || "").trim() : "",
        direction: map.direction != null ? String(row[map.direction] || "").trim() : "",
        mile_point: map.mile_point != null && row[map.mile_point] !== "" ? parseFloat(row[map.mile_point]) : null,
        side_of_road: map.side_of_road != null ? String(row[map.side_of_road] || "").trim() : "",
        lat: Number.isFinite(lat) ? lat : null,
        lng: Number.isFinite(lng) ? lng : null,
      };
      // Skip grouping/header rows that carry an ID but no actual sign data
      // (e.g. an assembly's base ID sitting above its A01/B01/... sub-signs).
      const hasData = rec.county || rec.route || rec.section || rec.direction ||
        rec.side_of_road || rec.mile_point != null || rec.lat != null || rec.lng != null;
      if (!hasData) continue;
      rows.push(rec);
    }
    return { rows };
  }

  async function doImport() {
    const p = App.parsed;
    if (!p || !p.rows.length) return;
    const btn = el("importBtn");
    btn.disabled = true;
    btn.textContent = "Importing…";
    try {
      let existing = new Set();
      try {
        const cur = await SB.select("signs", "select=id");
        existing = new Set(cur.map((r) => r.id));
      } catch {}
      // Collapse duplicate IDs within the sheet (last occurrence wins) — Postgres
      // upsert rejects the same key twice in one command.
      const byId = new Map();
      for (const r of p.rows) byId.set(r.id, r);
      const deduped = [...byId.values()];
      const dupes = p.rows.length - deduped.length;
      const payload = deduped.map((r) => Object.assign({ type: p.type, updated_at: new Date().toISOString() }, r));
      // Chunk to keep requests small.
      for (let i = 0; i < payload.length; i += 200)
        await SB.upsert("signs", payload.slice(i, i + 200), "id");
      const added = deduped.filter((r) => !existing.has(r.id)).length;
      const updated = deduped.length - added;
      App.parsed = null; // keep workbookCache so another sheet can be imported
      await loadSigns();
      setStatus(
        `Imported ${deduped.length} sign(s): ${added} new, ${updated} updated${dupes ? ` · ${dupes} duplicate ID(s) collapsed` : ""}.`,
        "ok"
      );
      render();
    } catch (e) {
      setStatus(`Import failed: ${e.message}`, "warn");
      btn.disabled = false;
      btn.textContent = `Import ${p.rows.length} sign(s)`;
    }
  }

  // ---- auth / boot -----------------------------------------------------------
  const USERNAME_DOMAIN = "bridgesign.app";
  const usernameToEmail = (u) => (u.includes("@") ? u : `${u}@${USERNAME_DOMAIN}`);

  function showChrome(show) {
    const tabs = document.querySelector(".tabs");
    if (tabs) tabs.style.display = show ? "" : "none";
  }

  function renderLogin(prefillUser, err) {
    showChrome(false);
    setStatus("");
    el("view").innerHTML = `
      <div class="login-card">
        <h2>Sign in</h2>
        <p class="hint">Access is limited to authorized users.</p>
        ${err ? `<div class="banner warn">${esc(err)}</div>` : ""}
        <label class="field"><span>Username</span>
          <input id="loginUser" type="text" autocomplete="username" value="${esc(prefillUser || "")}" /></label>
        <label class="field"><span>Password</span>
          <input id="loginPass" type="password" autocomplete="current-password" /></label>
        <button id="loginBtn" class="btn primary block">Sign in</button>
        <p class="hint">You'll stay signed in on this device until you log out.</p>
      </div>`;
    const submit = async () => {
      const u = el("loginUser").value.trim();
      const pw = el("loginPass").value;
      if (!u || !pw) return;
      const btn = el("loginBtn");
      btn.disabled = true;
      btn.textContent = "Signing in…";
      try {
        await SB.login(usernameToEmail(u), pw, u);
        await startApp();
      } catch (e) {
        renderLogin(u, e.message || "Login failed");
      }
    };
    el("loginBtn").addEventListener("click", submit);
    el("loginPass").addEventListener("keydown", (e) => { if (e.key === "Enter") submit(); });
    el("loginUser").focus();
  }

  function doLogout() {
    SB.logout();
    App.signs = [];
    App.remoteCaptures = [];
    App.recipients = [];
    editing = false;
    renderLogin();
  }

  async function startApp() {
    // Confirm the stored session is still good; otherwise back to login.
    if (!(await SB.ensureSession())) { doLogout(); return; }
    showChrome(true);
    App.screen = "signs";
    el("view").innerHTML = `<p class="hint">Loading…</p>`;
    await loadLocalCaptures();
    await loadSigns();
    loadRecipients();
    await loadSettings();
    await refreshCaptures();
    render();
    if (App.localCaptures.some((c) => c.status !== "synced")) syncAllPending();
  }

  // Accent presets. `gold` is the built-in default (uses the CSS defaults, which
  // include per-theme tuning); the others override the accent CSS variables.
  const ACCENTS = {
    gold:   { label: "Gold",   accent: "#e6bc3a" },
    blue:   { label: "Blue",   accent: "#3b82f6", strong: "#2563eb", ink: "#ffffff" },
    green:  { label: "Green",  accent: "#22c55e", strong: "#16884a", ink: "#04220f" },
    red:    { label: "Red",    accent: "#ef5350", strong: "#c0392b", ink: "#ffffff" },
    purple: { label: "Purple", accent: "#a78bfa", strong: "#7c3aed", ink: "#ffffff" },
    teal:   { label: "Teal",   accent: "#2dd4bf", strong: "#0f9488", ink: "#04241f" },
    orange: { label: "Orange", accent: "#fb923c", strong: "#d1701a", ink: "#241300" },
  };

  function applyTheme(t) {
    document.documentElement.setAttribute("data-theme", t);
    document.querySelectorAll(".tp-mode").forEach((b) => b.classList.toggle("active", b.dataset.mode === t));
  }
  function setTheme(t) {
    localStorage.setItem("bsh_theme", t);
    applyTheme(t);
  }
  function applyAccent(key) {
    const a = ACCENTS[key] || ACCENTS.gold;
    const root = document.documentElement.style;
    if (key === "gold" || !ACCENTS[key]) {
      root.removeProperty("--accent");
      root.removeProperty("--accent-strong");
      root.removeProperty("--accent-ink");
    } else {
      root.setProperty("--accent", a.accent);
      root.setProperty("--accent-strong", a.strong);
      root.setProperty("--accent-ink", a.ink);
    }
    document.querySelectorAll(".tp-swatch").forEach((b) => b.classList.toggle("active", b.dataset.accent === key));
  }
  function setAccent(key) {
    localStorage.setItem("bsh_accent", key);
    applyAccent(key);
  }
  function initTheme() {
    let t = localStorage.getItem("bsh_theme");
    if (!t) t = matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
    applyTheme(t);
    applyAccent(localStorage.getItem("bsh_accent") || "gold");
  }
  function buildThemePanel() {
    const box = el("tpSwatches");
    if (box)
      box.innerHTML = Object.entries(ACCENTS)
        .map(([k, a]) => `<button class="tp-swatch" data-accent="${k}" title="${a.label}" style="background:${a.accent}"></button>`)
        .join("");
    document.querySelectorAll(".tp-mode").forEach((b) => b.addEventListener("click", () => setTheme(b.dataset.mode)));
    document.querySelectorAll(".tp-swatch").forEach((b) => b.addEventListener("click", () => setAccent(b.dataset.accent)));
    applyTheme(document.documentElement.getAttribute("data-theme") || "light");
    applyAccent(localStorage.getItem("bsh_accent") || "gold");
    const btn = el("themeBtn");
    const panel = el("themePanel");
    btn.addEventListener("click", (e) => { e.stopPropagation(); panel.hidden = !panel.hidden; });
    document.addEventListener("click", (e) => {
      if (!panel.hidden && !panel.contains(e.target) && e.target !== btn) panel.hidden = true;
    });
  }

  function detectDevice() {
    const mobile =
      (matchMedia("(pointer: coarse)").matches && matchMedia("(max-width: 900px)").matches) ||
      /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent);
    App.isMobile = mobile;
    document.body.classList.toggle("is-mobile", mobile);
    document.body.classList.toggle("is-desktop", !mobile);
  }

  async function init() {
    initTheme();
    buildThemePanel();
    detectDevice();
    el("navSetup").addEventListener("click", () => { App.selectMode = false; App.screen = "setup"; render(); });
    el("navSigns").addEventListener("click", goSigns);
    el("navReview").addEventListener("click", () => { App.selectMode = false; reviewSelected = null; App.screen = "review"; render(); });
    if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});

    if (SB.isLoggedIn()) await startApp();
    else renderLogin();
  }

  // expose logout for the Setup screen button
  window.__bshLogout = doLogout;

  document.addEventListener("DOMContentLoaded", init);
})();
