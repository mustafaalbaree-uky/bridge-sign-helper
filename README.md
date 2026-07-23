# Bridge Sign Helper

A field tool for documenting bridge weight-limit sign inspections. Pick the sign,
snap up to two photos, and the app names the files correctly and hands them off to
the bridge engineer.

See [`docs/SPEC.md`](docs/SPEC.md) for the full requirements, data model, and roadmap.

## Status — Phase 2 (full workflow, connected)

The three processes are wired together through a shared **Supabase** backend
(Postgres + Storage), so photos taken on the phone show up on the computer.

- **Setup (computer)** — import the Excel sheet (`.xlsx`/`.csv`). It finds the
  header row, maps the columns, splits `Lat, Long`, and upserts into the database.
  Existing IDs update; new ones add. Pick "advance" or "bridge" per sheet.
- **Signs (phone)** — search or **Sort by nearest** (GPS); confirm the ID by eye;
  shoot up to two photos. Each photo saves locally first (offline-safe) and uploads
  to the backend, with a per-photo upload status.
- **Review (computer)** — pulls the batch from the backend, names each file
  `ID-YYMMDD` (`…1`/`…2` for two), exports into a chosen folder or downloads, and
  composes a notification email listing the files. Recipients are remembered.
- **Offline** — installable PWA; the sign list is cached and captures queue until
  signal returns.

Seeded with the 15 `Master List R12-6` signs so it works immediately; the real
import replaces them.

Coming next: the bridge-signs dataset (`…N`/`…L`/`…R`), automated email sending,
and access control (see Security below).

## Configuration

`config.js` holds the Supabase URL and **publishable** key. That key is meant to be
public — it only grants what the database's Row Level Security allows. The secret
service key is never in this repo.

## Security

Access requires a **login** (Supabase Auth). Row Level Security and the photo bucket
are locked to authenticated users, so the publishable key alone grants nothing without
a valid session. Sessions persist on the device until logout. User accounts are managed
in the Supabase dashboard (Authentication). Usernames map to an internal
`<username>@bridgesign.app` email for the auth server.

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
