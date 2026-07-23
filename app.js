// Bridge Sign Helper — client-side field app (Phase 1).
// No backend: sign data comes from data/signs.js, captures live in IndexedDB.

(() => {
  "use strict";

  const App = {
    signs: [],
    captures: [], // [{ key, signId, slot, blob, capturedAt, captureLat, captureLng }]
    position: null, // { lat, lng, accuracy }
    search: "",
    currentSignId: null,
    screen: "signs", // 'signs' | 'capture' | 'review'
  };

  let previewUrls = [];

  // ---- helpers ---------------------------------------------------------------

  const $ = (sel) => document.querySelector(sel);
  const el = (id) => document.getElementById(id);

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])
    );
  }

  function toRad(d) {
    return (d * Math.PI) / 180;
  }

  // Great-circle distance in meters.
  function distanceMeters(a, b) {
    const R = 6371000;
    const dLat = toRad(b.lat - a.lat);
    const dLng = toRad(b.lng - a.lng);
    const la1 = toRad(a.lat);
    const la2 = toRad(b.lat);
    const h =
      Math.sin(dLat / 2) ** 2 +
      Math.cos(la1) * Math.cos(la2) * Math.sin(dLng / 2) ** 2;
    return 2 * R * Math.asin(Math.sqrt(h));
  }

  function fmtDistance(m) {
    const miles = m / 1609.344;
    if (miles < 0.19) return `${Math.round(m * 3.28084)} ft`;
    return `${miles.toFixed(miles < 10 ? 2 : 1)} mi`;
  }

  function yymmdd(d) {
    const yy = String(d.getFullYear()).slice(-2);
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    return `${yy}${mm}${dd}`;
  }

  function extFor(blob) {
    if (blob && blob.type === "image/png") return "png";
    if (blob && blob.type === "image/heic") return "heic";
    return "jpg";
  }

  function capturesForSign(signId) {
    return App.captures.filter((c) => c.signId === signId);
  }

  // Build the correctly-named file list for export/review.
  // Groups by sign + capture date; single photo has no suffix, two get 1 / 2.
  function buildExportList() {
    const groups = {};
    for (const c of App.captures) {
      const date = yymmdd(new Date(c.capturedAt));
      const gkey = `${c.signId}|${date}`;
      (groups[gkey] = groups[gkey] || []).push(c);
    }
    const out = [];
    for (const gkey of Object.keys(groups)) {
      const [signId, date] = gkey.split("|");
      const items = groups[gkey].sort((a, b) => a.slot - b.slot);
      const two = items.length > 1;
      items.forEach((c, i) => {
        const suffix = two ? String(i + 1) : "";
        out.push({
          key: c.key,
          signId,
          filename: `${signId}-${date}${suffix}.${extFor(c.blob)}`,
          blob: c.blob,
          capturedAt: c.capturedAt,
        });
      });
    }
    return out.sort((a, b) => a.filename.localeCompare(b.filename));
  }

  function clearPreviews() {
    previewUrls.forEach((u) => URL.revokeObjectURL(u));
    previewUrls = [];
  }

  function previewFor(blob) {
    const u = URL.createObjectURL(blob);
    previewUrls.push(u);
    return u;
  }

  // ---- data / geolocation ----------------------------------------------------

  async function loadCaptures() {
    App.captures = await DB.all();
  }

  function requestLocation() {
    const btn = el("locateBtn");
    if (!navigator.geolocation) {
      setStatus("This device can't share location.");
      return;
    }
    if (btn) {
      btn.disabled = true;
      btn.textContent = "Locating…";
    }
    navigator.geolocation.getCurrentPosition(
      (pos) => {
        App.position = {
          lat: pos.coords.latitude,
          lng: pos.coords.longitude,
          accuracy: pos.coords.accuracy,
        };
        render();
      },
      (err) => {
        setStatus(`Location unavailable (${err.message}). You can still search.`);
        if (btn) {
          btn.disabled = false;
          btn.textContent = "📍 Sort by nearest";
        }
      },
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 30000 }
    );
  }

  function setStatus(msg) {
    const s = el("statusBar");
    if (!s) return;
    s.textContent = msg || "";
    s.hidden = !msg;
  }

  // ---- rendering -------------------------------------------------------------

  function render() {
    setNav();
    if (App.screen === "signs") renderSigns();
    else if (App.screen === "capture") renderCapture();
    else if (App.screen === "review") renderReview();
  }

  function setNav() {
    el("navSigns").classList.toggle("active", App.screen !== "review");
    el("navReview").classList.toggle("active", App.screen === "review");
    const n = buildExportList().length;
    el("reviewCount").textContent = n ? String(n) : "";
    el("reviewCount").hidden = !n;
  }

  function renderSigns() {
    clearPreviews();
    const q = App.search.trim().toLowerCase();
    let list = App.signs.filter((s) => s.activeStatus !== "Inactive");
    if (q) {
      list = list.filter((s) =>
        [s.id, s.county, s.route, s.direction, s.sideOfRoad]
          .join(" ")
          .toLowerCase()
          .includes(q)
      );
    }
    if (App.position) {
      list = list
        .map((s) => ({ s, d: distanceMeters(App.position, s) }))
        .sort((a, b) => a.d - b.d)
        .map((x) => Object.assign({}, x.s, { _dist: x.d }));
    }

    const accuracyNote = App.position
      ? `<div class="hint">Sorted by distance · GPS accuracy ±${Math.round(
          App.position.accuracy
        )} m. <strong>Confirm the ID by eye</strong> — nearby L/R signs can look identical.</div>`
      : "";

    const rows = list
      .map((s) => {
        const caps = capturesForSign(s.id).length;
        const badge = caps
          ? `<span class="badge done">✓ ${caps} photo${caps > 1 ? "s" : ""}</span>`
          : "";
        const dist =
          s._dist != null ? `<span class="dist">${fmtDistance(s._dist)}</span>` : "";
        return `
          <li class="sign-item" data-id="${esc(s.id)}">
            <div class="sign-main">
              <div class="sign-id">${esc(s.id)} ${badge}</div>
              <div class="sign-sub">${esc(s.county)} · ${esc(s.route)} · MP ${esc(
          s.milePoint
        )} · ${esc(s.direction)} · ${esc(s.sideOfRoad)}</div>
            </div>
            ${dist}
          </li>`;
      })
      .join("");

    el("view").innerHTML = `
      <div class="banner">Demo data from R12-6. Real sheet import comes next.</div>
      <div class="toolbar">
        <input id="searchInput" class="search" type="search"
          placeholder="Search ID, route, county…" value="${esc(App.search)}" />
        <button id="locateBtn" class="btn secondary">📍 Sort by nearest</button>
      </div>
      ${accuracyNote}
      <ul class="sign-list">${rows || `<li class="empty">No matching signs.</li>`}</ul>
    `;

    const search = el("searchInput");
    search.addEventListener("input", (e) => {
      App.search = e.target.value;
      // Re-render list only, keep focus.
      renderSigns();
      const s2 = el("searchInput");
      s2.focus();
      s2.setSelectionRange(s2.value.length, s2.value.length);
    });
    el("locateBtn").addEventListener("click", requestLocation);
    el("view")
      .querySelectorAll(".sign-item")
      .forEach((li) =>
        li.addEventListener("click", () => openCapture(li.dataset.id))
      );
  }

  function openCapture(signId) {
    App.currentSignId = signId;
    App.screen = "capture";
    render();
  }

  function renderCapture() {
    clearPreviews();
    const sign = App.signs.find((s) => s.id === App.currentSignId);
    if (!sign) {
      App.screen = "signs";
      return render();
    }
    const caps = capturesForSign(sign.id).sort((a, b) => a.slot - b.slot);
    const slot = (n) => caps.find((c) => c.slot === n);

    const slotHtml = (n) => {
      const c = slot(n);
      if (c) {
        return `
          <div class="slot filled">
            <img src="${previewFor(c.blob)}" alt="Photo ${n}" />
            <div class="slot-actions">
              <label class="btn small">Retake
                <input type="file" accept="image/*" capture="environment"
                  data-slot="${n}" hidden />
              </label>
              <button class="btn small danger" data-remove="${n}">Remove</button>
            </div>
          </div>`;
      }
      return `
        <label class="slot empty">
          <span class="slot-plus">＋</span>
          <span>Photo ${n}</span>
          <input type="file" accept="image/*" capture="environment"
            data-slot="${n}" hidden />
        </label>`;
    };

    el("view").innerHTML = `
      <button id="backBtn" class="btn link">‹ All signs</button>
      <div class="detail-card">
        <div class="detail-id">${esc(sign.id)}</div>
        <div class="detail-grid">
          <div><span>County</span>${esc(sign.county)}</div>
          <div><span>Route</span>${esc(sign.route)}</div>
          <div><span>Mile point</span>${esc(sign.milePoint)}</div>
          <div><span>Direction</span>${esc(sign.direction)}</div>
          <div><span>Side</span>${esc(sign.sideOfRoad)}</div>
          <div><span>Coordinates</span>${esc(sign.lat)}, ${esc(sign.lng)}</div>
        </div>
      </div>
      <div class="confirm-note">Confirm this is the correct sign before shooting.</div>
      <div class="slots">${slotHtml(1)}${slotHtml(2)}</div>
      <button id="doneBtn" class="btn primary block">Save &amp; back to list</button>
    `;

    el("backBtn").addEventListener("click", () => {
      App.screen = "signs";
      render();
    });
    el("doneBtn").addEventListener("click", () => {
      App.screen = "signs";
      render();
    });
    el("view")
      .querySelectorAll('input[type="file"]')
      .forEach((inp) =>
        inp.addEventListener("change", (e) =>
          onPhoto(sign.id, Number(inp.dataset.slot), e.target.files[0])
        )
      );
    el("view")
      .querySelectorAll("[data-remove]")
      .forEach((b) =>
        b.addEventListener("click", () =>
          removePhoto(sign.id, Number(b.dataset.remove))
        )
      );
  }

  async function onPhoto(signId, slot, file) {
    if (!file) return;
    const record = {
      key: `${signId}__${slot}`,
      signId,
      slot,
      blob: file,
      capturedAt: new Date().toISOString(),
      captureLat: App.position ? App.position.lat : null,
      captureLng: App.position ? App.position.lng : null,
    };
    await DB.put(record);
    await loadCaptures();
    render();
  }

  async function removePhoto(signId, slot) {
    await DB.remove(`${signId}__${slot}`);
    await loadCaptures();
    render();
  }

  function renderReview() {
    clearPreviews();
    const files = buildExportList();
    const fsaSupported = typeof window.showDirectoryPicker === "function";

    const rows = files
      .map(
        (f) => `
        <li class="file-row">
          <img src="${previewFor(f.blob)}" alt="" />
          <div class="file-meta">
            <div class="file-name">${esc(f.filename)}</div>
            <div class="file-sub">${new Date(f.capturedAt).toLocaleString()}</div>
          </div>
        </li>`
      )
      .join("");

    el("view").innerHTML = `
      <h2 class="screen-title">Review &amp; export</h2>
      ${
        files.length
          ? `<p class="hint">${files.length} photo${
              files.length > 1 ? "s" : ""
            } ready. Names follow <code>ID-YYMMDD</code>.</p>`
          : `<p class="empty">No photos captured yet.</p>`
      }
      <ul class="file-list">${rows}</ul>
      ${
        files.length
          ? `<div class="export-actions">
              ${
                fsaSupported
                  ? `<button id="exportFolder" class="btn primary block">Save all to a folder…</button>`
                  : ""
              }
              <button id="downloadAll" class="btn ${
                fsaSupported ? "secondary" : "primary"
              } block">Download all</button>
              <button id="copyNames" class="btn secondary block">Copy filename list</button>
              <button id="flushBtn" class="btn danger block">Clear all captures…</button>
            </div>`
          : ""
      }
    `;

    if (!files.length) return;
    if (fsaSupported)
      el("exportFolder").addEventListener("click", () => exportToFolder(files));
    el("downloadAll").addEventListener("click", () => downloadAll(files));
    el("copyNames").addEventListener("click", () => copyNames(files));
    el("flushBtn").addEventListener("click", flushAll);
  }

  // ---- export ----------------------------------------------------------------

  async function exportToFolder(files) {
    try {
      const dir = await window.showDirectoryPicker({ mode: "readwrite" });
      for (const f of files) {
        const fh = await dir.getFileHandle(f.filename, { create: true });
        const w = await fh.createWritable();
        await w.write(f.blob);
        await w.close();
      }
      setStatus(`Saved ${files.length} photo(s) to the chosen folder.`);
    } catch (err) {
      if (err && err.name === "AbortError") return;
      setStatus(`Could not save to folder: ${err.message}`);
    }
  }

  function downloadAll(files) {
    files.forEach((f, i) => {
      setTimeout(() => {
        const url = URL.createObjectURL(f.blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = f.filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
      }, i * 300);
    });
  }

  async function copyNames(files) {
    const text = files.map((f) => f.filename).join("\n");
    try {
      await navigator.clipboard.writeText(text);
      setStatus("Filename list copied.");
    } catch {
      setStatus(text);
    }
  }

  async function flushAll() {
    if (!confirm("Clear all captured photos from this device? Export them first.")) return;
    await DB.clear();
    await loadCaptures();
    setStatus("Captures cleared.");
    render();
  }

  // ---- boot ------------------------------------------------------------------

  async function init() {
    App.signs = window.SEED_SIGNS || [];
    await loadCaptures();

    el("navSigns").addEventListener("click", () => {
      App.screen = "signs";
      render();
    });
    el("navReview").addEventListener("click", () => {
      App.screen = "review";
      render();
    });

    render();

    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("sw.js").catch(() => {});
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})();
