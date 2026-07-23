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
    parsed: null, // staged Excel import { type, rows, sheetNames, sheet }
  };

  let previewUrls = [];

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

  function clearPreviews() {
    previewUrls.forEach((u) => URL.revokeObjectURL(u));
    previewUrls = [];
  }
  function preview(blob) {
    const u = URL.createObjectURL(blob);
    previewUrls.push(u);
    return u;
  }
  function setStatus(msg, kind) {
    const s = el("statusBar");
    if (!s) return;
    s.textContent = msg || "";
    s.hidden = !msg;
    s.className = "status" + (kind ? " " + kind : "");
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

  async function loadRecipients() {
    try {
      const rows = await SB.select("recipients", "select=email&order=added_at.desc");
      App.recipients = rows.map((r) => r.email);
    } catch {
      App.recipients = [];
    }
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
    setStatus(stillBad ? `${stillBad} still not uploaded — check signal.` : "All photos uploaded.", stillBad ? "warn" : "ok");
  }

  async function removePhoto(signId, slot) {
    const batchDate = todayStr();
    await DB.removeCapture(`${signId}__${slot}__${batchDate}`);
    try {
      await SB.remove("captures", `sign_id=eq.${signId}&slot=eq.${slot}&batch_date=eq.${batchDate}`);
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

  // ---- rendering: signs list -------------------------------------------------
  function renderSigns() {
    clearPreviews();
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

    const today = todayStr();
    const rows = list
      .map((s) => {
        const caps = App.localCaptures.filter((c) => c.signId === s.id && c.batchDate === today).length;
        const badge = caps ? `<span class="badge done">✓ ${caps}</span>` : "";
        const dist = s._d != null ? `<span class="dist">${fmtDistance(s._d)}</span>` : "";
        return `<li class="sign-item" data-id="${esc(s.id)}">
            <div class="sign-main">
              <div class="sign-id">${esc(s.id)} ${badge}</div>
              <div class="sign-sub">${esc(s.county)} · ${esc(s.route)} · MP ${esc(s.mile_point)} · ${esc(s.direction)} · ${esc(s.side_of_road)}</div>
            </div>${dist}
          </li>`;
      })
      .join("");

    const conn = App.online
      ? ""
      : `<div class="banner warn">Offline — showing the last synced sign list. Photos will upload when you're back on signal.</div>`;
    const note = App.position
      ? `<div class="hint">Nearest first · GPS ±${Math.round(App.position.accuracy)} m. <strong>Confirm the ID by eye.</strong></div>`
      : "";

    el("view").innerHTML = `
      ${conn}
      <div class="toolbar">
        <input id="searchInput" class="search" type="search" placeholder="Search ID, route, county…" value="${esc(App.search)}" />
        <button id="locateBtn" class="btn secondary">📍 Sort by nearest</button>
      </div>${note}
      <ul class="sign-list">${rows || `<li class="empty">No signs. Import a sheet under <strong>Setup</strong>.</li>`}</ul>`;

    const search = el("searchInput");
    search.addEventListener("input", (e) => {
      App.search = e.target.value;
      renderSigns();
      const s2 = el("searchInput");
      s2.focus();
      s2.setSelectionRange(s2.value.length, s2.value.length);
    });
    el("locateBtn").addEventListener("click", requestLocation);
    el("view").querySelectorAll(".sign-item").forEach((li) =>
      li.addEventListener("click", () => { App.currentSignId = li.dataset.id; App.screen = "capture"; render(); })
    );
  }

  // ---- rendering: capture ----------------------------------------------------
  function renderCapture() {
    clearPreviews();
    const sign = App.signs.find((s) => s.id === App.currentSignId);
    if (!sign) { App.screen = "signs"; return render(); }
    const today = todayStr();
    const slotRec = (n) => App.localCaptures.find((c) => c.signId === sign.id && c.slot === n && c.batchDate === today);

    const statusChip = (r) =>
      r.status === "synced" ? `<span class="chip ok">Uploaded</span>`
      : r.status === "error" ? `<span class="chip err">Not uploaded</span>`
      : `<span class="chip">Saving…</span>`;

    const slotHtml = (n) => {
      const r = slotRec(n);
      if (r)
        return `<div class="slot filled">
            <img src="${preview(r.blob)}" alt="Photo ${n}" />
            <div class="slot-bar">${statusChip(r)}
              <label class="btn small">Retake<input type="file" accept="image/*" capture="environment" data-slot="${n}" hidden /></label>
              <button class="btn small danger" data-remove="${n}">✕</button>
            </div></div>`;
      return `<label class="slot empty"><span class="slot-plus">＋</span><span>Photo ${n}</span>
          <input type="file" accept="image/*" capture="environment" data-slot="${n}" hidden /></label>`;
    };

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
      </div>
      <div class="confirm-note">Confirm this is the correct sign before shooting.</div>
      <div class="slots">${slotHtml(1)}${slotHtml(2)}</div>
      <button id="doneBtn" class="btn primary block">Done — back to list</button>`;

    el("backBtn").addEventListener("click", () => { App.screen = "signs"; render(); });
    el("doneBtn").addEventListener("click", () => { App.screen = "signs"; render(); });
    el("view").querySelectorAll('input[type="file"]').forEach((inp) =>
      inp.addEventListener("change", (e) => onPhoto(sign.id, Number(inp.dataset.slot), e.target.files[0]))
    );
    el("view").querySelectorAll("[data-remove]").forEach((b) =>
      b.addEventListener("click", () => removePhoto(sign.id, Number(b.dataset.remove)))
    );
  }

  // ---- rendering: review / flush --------------------------------------------
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
        });
      });
    }
    return out.sort((a, b) => a.filename.localeCompare(b.filename));
  }

  async function renderReview() {
    clearPreviews();
    el("view").innerHTML = `<p class="hint">Loading today's batch…</p>`;
    let remote = [];
    try {
      remote = await SB.select("captures", `select=*&order=batch_date.desc,sign_id.asc,slot.asc`);
    } catch {
      el("view").innerHTML = `<div class="banner warn">Can't reach the server. Connect to load the batch for export.</div>`;
      return;
    }
    App.remoteCaptures = remote;
    const files = buildExportList(remote);
    const pending = App.localCaptures.filter((c) => c.status !== "synced").length;
    const fsa = typeof window.showDirectoryPicker === "function";

    // Fetch thumbnails (small day batches — fine to pull).
    const thumbs = {};
    await Promise.all(
      files.map(async (f) => {
        try { thumbs[f.storage_path] = preview(await SB.downloadPhoto(f.storage_path)); } catch {}
      })
    );

    const rows = files
      .map(
        (f) => `<li class="file-row">
          ${thumbs[f.storage_path] ? `<img src="${thumbs[f.storage_path]}" alt="" />` : `<div class="thumb-missing">?</div>`}
          <div class="file-meta"><div class="file-name">${esc(f.filename)}</div>
            <div class="file-sub">${new Date(f.captured_at).toLocaleString()}</div></div>
        </li>`
      )
      .join("");

    const datalist = App.recipients.map((e) => `<option value="${esc(e)}">`).join("");

    el("view").innerHTML = `
      <h2 class="screen-title">Review &amp; export</h2>
      ${pending ? `<div class="banner warn">${pending} photo(s) on this device haven't uploaded. <button id="syncNow" class="btn small">Sync now</button></div>` : ""}
      ${files.length ? `<p class="hint">${files.length} photo(s) ready · names follow <code>ID-YYMMDD</code>.</p>` : `<p class="empty">No photos on the server yet.</p>`}
      <ul class="file-list">${rows}</ul>
      ${files.length ? `
      <div class="export-actions">
        ${fsa ? `<button id="exportFolder" class="btn primary block">Save all to a folder…</button>` : ""}
        <button id="downloadAll" class="btn ${fsa ? "secondary" : "primary"} block">Download all</button>
        <button id="copyNames" class="btn secondary block">Copy filename list</button>
      </div>
      <div class="notify">
        <h3>Notify the engineer</h3>
        <input id="emailInput" class="search" list="recips" type="email" placeholder="engineer@example.com" />
        <datalist id="recips">${datalist}</datalist>
        <button id="composeBtn" class="btn secondary block">Compose email with file list</button>
      </div>
      <button id="clearLocal" class="btn danger block">Clear this phone's local copies</button>` : ""}`;

    if (pending) el("syncNow").addEventListener("click", syncAllPending);
    if (!files.length) return;
    if (fsa) el("exportFolder").addEventListener("click", () => exportToFolder(files));
    el("downloadAll").addEventListener("click", () => downloadAll(files));
    el("copyNames").addEventListener("click", () => copyNames(files));
    el("composeBtn").addEventListener("click", () => composeEmail(files));
    el("clearLocal").addEventListener("click", clearLocalCopies);
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
    } catch (e) {
      if (e && e.name === "AbortError") return;
      setStatus(`Could not save: ${e.message}`, "warn");
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
  }

  async function copyNames(files) {
    const text = files.map((f) => f.filename).join("\n");
    try { await navigator.clipboard.writeText(text); setStatus("Filename list copied.", "ok"); }
    catch { setStatus(text); }
  }

  async function composeEmail(files) {
    const to = (el("emailInput").value || "").trim();
    if (!to) return setStatus("Enter an email address first.", "warn");
    try { await SB.upsert("recipients", { email: to }, "email"); await loadRecipients(); } catch {}
    const subject = `Bridge sign photos — ${todayStr()}`;
    const body =
      "The following sign inspection photos have been added:\n\n" +
      files.map((f) => f.filename).join("\n") +
      "\n";
    window.location.href = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  }

  async function clearLocalCopies() {
    if (!confirm("Remove photos stored on THIS device? They stay on the server. Make sure everything is uploaded first.")) return;
    await DB.clearCaptures();
    await loadLocalCaptures();
    setStatus("Local copies cleared.", "ok");
    render();
  }

  // ---- rendering: setup / Excel import --------------------------------------
  function renderSetup() {
    clearPreviews();
    const p = App.parsed;
    el("view").innerHTML = `
      <h2 class="screen-title">Setup — import signs</h2>
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
        ${p && p.sheetNames && p.sheetNames.length > 1 ? `
          <label class="field"><span>Sheet</span>
            <select id="sheetSel">${p.sheetNames.map((n) => `<option ${n === p.sheet ? "selected" : ""}>${esc(n)}</option>`).join("")}</select>
          </label>` : ""}
      </div>
      ${p ? renderParsedPreview(p) : ""}
      <div class="setup-card">
        <h3>Current database</h3>
        <p class="hint">${App.online ? `${App.signs.length} sign(s) loaded from the server.` : "Offline — can't reach the server."}</p>
      </div>`;

    el("fileInput").addEventListener("change", (e) => handleFile(e.target.files[0]));
    el("typeSel").value = p ? p.type : "advance";
    el("typeSel").addEventListener("change", (e) => { if (App.parsed) App.parsed.type = e.target.value; });
    const sheetSel = el("sheetSel");
    if (sheetSel) sheetSel.addEventListener("change", (e) => reparseSheet(e.target.value));
    const imp = el("importBtn");
    if (imp) imp.addEventListener("click", doImport);
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
        ${bad ? `<div class="banner warn">${bad} row(s) have no usable coordinates — they'll import but won't sort by distance.</div>` : ""}
        <div class="table-scroll"><table class="preview">
          <thead><tr><th>ID</th><th>County</th><th>Route</th><th>Lat, Long</th></tr></thead>
          <tbody>${head}</tbody></table></div>
        ${p.rows.length > 5 ? `<p class="hint">…and ${p.rows.length - 5} more.</p>` : ""}
        <button id="importBtn" class="btn primary block">Import ${p.rows.length} sign(s)</button>
      </div>`;
  }

  let workbookCache = null;

  async function handleFile(file) {
    if (!file) return;
    setStatus("Reading file…");
    try {
      const buf = await file.arrayBuffer();
      workbookCache = XLSX.read(buf, { type: "array" });
      const sheetNames = workbookCache.SheetNames;
      const sheet = pickBestSheet(workbookCache);
      App.parsed = Object.assign(parseSheet(workbookCache, sheet), {
        type: el("typeSel") ? el("typeSel").value : "advance",
        sheetNames,
        sheet,
      });
      setStatus("");
      render();
    } catch (e) {
      setStatus(`Couldn't read that file: ${e.message}`, "warn");
    }
  }

  function reparseSheet(sheet) {
    if (!workbookCache) return;
    const type = App.parsed ? App.parsed.type : "advance";
    App.parsed = Object.assign(parseSheet(workbookCache, sheet), {
      type, sheetNames: workbookCache.SheetNames, sheet,
    });
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
      rows.push({
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
      });
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
      const payload = p.rows.map((r) => Object.assign({ type: p.type, updated_at: new Date().toISOString() }, r));
      // Chunk to keep requests small.
      for (let i = 0; i < payload.length; i += 200)
        await SB.upsert("signs", payload.slice(i, i + 200), "id");
      const added = p.rows.filter((r) => !existing.has(r.id)).length;
      const updated = p.rows.length - added;
      App.parsed = null;
      workbookCache = null;
      await loadSigns();
      setStatus(`Imported ${p.rows.length} sign(s): ${added} new, ${updated} updated.`, "ok");
      render();
    } catch (e) {
      setStatus(`Import failed: ${e.message}`, "warn");
      btn.disabled = false;
      btn.textContent = `Import ${p.rows.length} sign(s)`;
    }
  }

  // ---- boot ------------------------------------------------------------------
  async function init() {
    el("navSetup").addEventListener("click", () => { App.screen = "setup"; render(); });
    el("navSigns").addEventListener("click", () => { App.screen = "signs"; render(); });
    el("navReview").addEventListener("click", () => { App.screen = "review"; render(); });

    el("view").innerHTML = `<p class="hint">Loading…</p>`;
    await loadLocalCaptures();
    await loadSigns();
    loadRecipients();
    render();

    // Retry any photos that didn't upload last session.
    if (App.localCaptures.some((c) => c.status !== "synced")) syncAllPending();

    if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
  }

  document.addEventListener("DOMContentLoaded", init);
})();
