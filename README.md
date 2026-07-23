# Bridge Sign Helper

A field tool for documenting bridge weight-limit sign inspections. Pick the sign,
snap up to two photos, and the app names the files correctly and hands them off to
the bridge engineer.

See [`docs/SPEC.md`](docs/SPEC.md) for the full requirements, data model, and roadmap.

## Status — Phase 1 (client-side experiment)

Runs entirely in the browser, no backend yet. Uses placeholder sign data
transcribed from the `Master List R12-6` sheet until the real Excel export is
available.

Working now:

- **Sign picker** — search by ID / route / county, or tap **Sort by nearest** to
  order signs by GPS distance (always confirm the ID by eye).
- **Capture** — up to two photos per sign, straight from the camera.
- **Naming** — files follow `ID-YYMMDD`, and `ID-YYMMDD1` / `ID-YYMMDD2` for two.
- **Review & export** — save all photos into a chosen folder (Chrome/Edge desktop)
  or download them; copy the filename list.
- **Offline** — installable PWA; works without signal at the bridge.

Coming next: Excel import, a shared Supabase database + photo storage so the phone
and computer stay in sync, and the email notification to the engineer.

## Run it

It's a static site — any web server works:

```bash
python3 -m http.server 8000
# open http://localhost:8000
```

## Deploy (GitHub Pages)

Served from the `main` branch. Enable it under
**Settings → Pages → Source: Deploy from a branch → `main` / root**.

> Camera and GPS require HTTPS. GitHub Pages is HTTPS, so both work there;
> `localhost` also counts as secure for local testing.
