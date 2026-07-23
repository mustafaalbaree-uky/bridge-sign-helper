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
    photoCounts: {}, // signId -> number of photos on the server
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

  async function loadPhotoCounts() {
    try {
      const rows = await SB.select("captures", "select=sign_id");
      const m = {};
      for (const r of rows) m[r.sign_id] = (m[r.sign_id] || 0) + 1;
      App.photoCounts = m;
    } catch { /* keep whatever we had */ }
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
    setStatus(stillBad ? `${stillBad} still not uploaded. Try again when you have signal.` : "All photos uploaded.", stillBad ? "warn" : "ok");
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

  // Go to the signs list and refresh photo counts from the server in the
  // background (badge updates once they arrive).
  function goSigns() {
    editing = false;
    App.screen = "signs";
    render();
    loadPhotoCounts().then(() => { if (App.screen === "signs") render(); });
  }

  // ---- rendering: signs list -------------------------------------------------
  function renderSigns() {
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

    const rows = list
      .map((s) => {
        const localN = App.localCaptures.filter((c) => c.signId === s.id).length;
        const caps = Math.max(App.photoCounts[s.id] || 0, localN);
        const badge = caps ? `<span class="badge done">${caps} Photo${caps > 1 ? "s" : ""}</span>` : "";
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
      : `<div class="banner warn">Offline. Showing the last synced sign list; photos upload when you're back on signal.</div>`;
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

  // Show photos already in the database for this sign (works on any device,
  // e.g. reviewing on the computer what was shot in the field).
  async function loadOnFile(signId) {
    const box = el("onFile");
    if (!box) return;
    let rows = [];
    try {
      rows = await SB.select(
        "captures",
        `select=storage_path,batch_date,slot,captured_at&sign_id=eq.${encodeURIComponent(signId)}&order=batch_date.desc,slot.asc`
      );
    } catch { return; }
    if (App.currentSignId !== signId || App.screen !== "capture") return; // navigated away
    if (!rows.length) return;

    // Group by date, most recent first.
    const byDate = {};
    for (const r of rows) (byDate[r.batch_date] = byDate[r.batch_date] || []).push(r);
    const dates = Object.keys(byDate).sort().reverse();

    box.innerHTML =
      `<h3 class="on-file-h">Photos on file</h3>` +
      dates
        .map(
          (d) =>
            `<div class="on-file-day"><div class="on-file-date">${esc(d)}</div>
               <div class="on-file-thumbs">${byDate[d]
                 .map((r) => `<a class="of-thumb" data-path="${esc(r.storage_path)}" title="${esc(r.storage_path)}"><span>loading…</span></a>`)
                 .join("")}</div></div>`
        )
        .join("");

    // Fetch each thumbnail via the login token and swap it in.
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
    const emailReady = !!(App.settings && App.settings.email_webhook_url);
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

    const fileRow = (f) => `<li class="file-row">
        ${thumbs[f.storage_path] ? `<img src="${thumbs[f.storage_path]}" alt="" data-revoke />` : `<div class="thumb-missing">?</div>`}
        <div class="file-meta"><div class="file-name">${esc(f.filename)}</div>
          <div class="file-sub">${new Date(f.captured_at).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}</div></div>
      </li>`;

    const dayBlock = (date) => {
      const df = byDate.get(date);
      const label = date === today ? "Today" : prettyDate(date);
      return `<section class="day-group">
          <div class="day-head"><span class="day-label">${label}</span>
            <span class="day-count">${df.length} photo${df.length > 1 ? "s" : ""}</span></div>
          <ul class="file-list">${df.map(fileRow).join("")}</ul>
          <div class="day-actions">
            ${fsa ? `<button class="btn small primary" data-act="folder" data-date="${date}">Save to folder</button>` : ""}
            <button class="btn small secondary" data-act="download" data-date="${date}">Download</button>
            <button class="btn small secondary" data-act="email" data-date="${date}">Email this day</button>
          </div>
        </section>`;
    };

    const datalist = App.recipients.map((e) => `<option value="${esc(e)}">`).join("");

    el("view").innerHTML = `
      <h2 class="screen-title">Review &amp; export</h2>
      <p class="hint">Every inspection photo on the server, newest day first. Photos stay here until you delete them; exporting or emailing does not remove them. Files are named <code>ID-YYMMDD</code>.</p>
      ${pending ? `<div class="banner warn">${pending} photo(s) on this device haven't uploaded. <button id="syncNow" class="btn small">Sync now</button></div>` : ""}
      ${files.length ? `
        <div class="notify">
          <label class="fieldlabel" for="emailInput">Email recipient (used by “Email this day”)</label>
          <input id="emailInput" class="search" type="email" list="recips"
            name="bsh-notify-recipient" autocomplete="off" autocapitalize="off"
            autocorrect="off" spellcheck="false" data-lpignore="true" data-1p-ignore="true"
            data-form-type="other" placeholder="engineer@example.com" />
          <datalist id="recips">${datalist}</datalist>
          <p class="hint">${emailReady ? "Emailing sends automatically." : "Emailing opens your mail app. Set up automatic sending under Setup."}</p>
        </div>
        ${dates.map(dayBlock).join("")}
        ${App.localCaptures.length ? `<button id="clearLocal" class="btn danger block">Clear photos saved on this device (${App.localCaptures.length})</button>` : ""}
      ` : `<p class="empty">No photos on the server yet.</p>`}`;

    wireThumbs(el("view"));
    if (pending) el("syncNow").addEventListener("click", syncAllPending);
    if (!files.length) return;
    el("view").querySelectorAll("[data-act]").forEach((b) =>
      b.addEventListener("click", () => {
        const df = byDate.get(b.dataset.date) || [];
        if (b.dataset.act === "folder") exportToFolder(df);
        else if (b.dataset.act === "download") downloadAll(df);
        else if (b.dataset.act === "email") (emailReady ? sendEmail(df, b) : composeEmail(df));
      })
    );
    if (App.localCaptures.length) el("clearLocal").addEventListener("click", clearLocalCopies);
  }

  async function sendEmail(files, btn) {
    const to = (el("emailInput").value || "").trim();
    if (!to) return setStatus("Enter an email address first.", "warn");
    const label = btn ? btn.textContent : "";
    if (btn) { btn.disabled = true; btn.textContent = "Sending…"; }
    const subject = `Bridge sign photos, ${files[0] ? files[0].batch_date : todayStr()}`;
    const body =
      "The following sign inspection photos have been added:\n\n" +
      files.map((f) => f.filename).join("\n") + "\n";
    try {
      try { await SB.upsert("recipients", { email: to }, "email"); await loadRecipients(); } catch {}
      await SB.invoke("notify", { to, subject, body });
      setStatus(`Email sent to ${to}.`, "ok");
    } catch (e) {
      setStatus(`Send failed: ${e.message}. You can use Compose instead.`, "warn", 9000);
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = label || "Email this day"; }
    }
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
    const subject = `Bridge sign photos, ${todayStr()}`;
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
  let pendingToken = "";
  const randToken = () => Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);

  function renderSetup() {
    const p = App.parsed;
    const emailReady = !!(App.settings && App.settings.email_webhook_url);
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
      <details class="setup-card collapsible"${emailReady ? "" : ""}>
        <summary>Email notifications ${emailReady ? '<span class="chip ok">on</span>' : '<span class="chip">off</span>'}</summary>
        <p class="hint">${emailReady
          ? "Automatic sending is set up. Review can send email directly."
          : "Not set up yet. Until then, Review composes an email in your mail app. Steps: docs/EMAIL_SETUP.md."}</p>
        <label class="field"><span>Apps Script Web App URL</span>
          <input id="whUrl" type="url" autocomplete="off" data-lpignore="true" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(App.settings.email_webhook_url || "")}" /></label>
        <label class="field"><span>Shared token (paste this same value into the script)</span>
          <input id="whToken" type="text" autocomplete="off" data-lpignore="true" value="${esc(tokenVal)}" /></label>
        <button id="saveEmail" class="btn primary block">Save email settings</button>
        <button id="testEmail" class="btn secondary block">Send a test email</button>
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
    await loadPhotoCounts();
    render();
    if (App.localCaptures.some((c) => c.status !== "synced")) syncAllPending();
  }

  function applyTheme(t) {
    document.documentElement.setAttribute("data-theme", t);
    const b = el("themeToggle");
    if (b) b.textContent = t === "dark" ? "☀️" : "🌙";
  }
  function initTheme() {
    let t = localStorage.getItem("bsh_theme");
    if (!t) t = matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
    applyTheme(t);
  }
  function toggleTheme() {
    const next = document.documentElement.getAttribute("data-theme") === "dark" ? "light" : "dark";
    localStorage.setItem("bsh_theme", next);
    applyTheme(next);
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
    detectDevice();
    el("themeToggle").addEventListener("click", toggleTheme);
    el("navSetup").addEventListener("click", () => { App.screen = "setup"; render(); });
    el("navSigns").addEventListener("click", goSigns);
    el("navReview").addEventListener("click", () => { App.screen = "review"; render(); });
    if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});

    if (SB.isLoggedIn()) await startApp();
    else renderLogin();
  }

  // expose logout for the Setup screen button
  window.__bshLogout = doLogout;

  document.addEventListener("DOMContentLoaded", init);
})();
