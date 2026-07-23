# Bridge Sign Helper — Specification

Field tool for documenting bridge weight-limit signage inspections and getting the
photos to the bridge engineer with zero manual filing.

## The problem

An inspector (Brian) drives a route inspecting installed weight-limit signs. For each
sign he must photograph it and get that photo to the bridge engineer, who drops it into
his report. Today that means fiddling with SharePoint, renaming files by hand, and
emailing. We want: pick the sign → snap the photo → done. The photo lands in a folder,
correctly named, and the engineer is notified.

## Two kinds of signs (two datasets)

1. **Advance Weight Limit signs** ("advance signs") — one ID per sign.
   Source sheet: `Master List R12-6`. ID looks like `AW1052508121223`.
2. **Weight Limit signs at the bridge** ("bridge signs") — tied to a bridge, and a
   single crossing can have more than one physical sign.
   ID looks like `105B00015N`, decoded as:
   - `105` = county code (105 = Scott)
   - `B` = bridge
   - `00015` = bridge number 15
   - `N` / `L` / `R` = configuration. `N` = one two-way structure; `L` / `R` = separate
     left/right structures (e.g. `105B00015L` and `105B00015R`).

   So one bridge may appear as `…N`, or as a `…L` + `…R` pair. **These are close
   together and GPS margin can confuse them — the inspector must always confirm the ID
   by eye; the app must never silently pick L when it's R.**

Both datasets fit one record shape, so the app treats them uniformly.

## Data model (per sign)

Columns we pull from the sheet (from the R12-6 screenshot):

| Field         | Example              | Notes                                   |
|---------------|----------------------|-----------------------------------------|
| `activeStatus`| `Active`             | Only `Active` rows are shown            |
| `id`          | `AW1052508121223`    | The filename stem                       |
| `county`      | `Scott`              |                                         |
| `route`       | `US-0025`            |                                         |
| `section`     | (often blank)        |                                         |
| `direction`   | `Increasing`         | Increasing / Decreasing mile point      |
| `milePoint`   | `20.656`             |                                         |
| `sideOfRoad`  | `Right`              |                                         |
| `lat, lng`    | `38.433738,-84.566…` | One "Lat, Long" cell, split on import   |

The two relevant fields for the core job are **`id`** and **`lat,lng`**; the rest is
context to help the inspector confirm he's at the right sign.

## Filename standard

```
{ID}-{YYMMDD}        ← single photo
{ID}-{YYMMDD}1       ← first of two
{ID}-{YYMMDD}2       ← second of two
```

- `YYMMDD` = 2-digit year, month, day of the capture date (e.g. 2026-07-23 → `260723`).
- The number suffix appears **only when there are two photos** for that sign that day.
- Max two photos per sign. Extension preserved (`.jpg`).
- The engineer searches a folder by bridge/sign ID and reads the date to find the most
  recent photo — no email required, though we send one as a courtesy.

*Known edge case:* re-inspecting the same sign twice in one day would collide on the
name. For now the later capture replaces the earlier (matches "latest photo wins"). If
that becomes a problem we append a time suffix.

## Workflow

### 1. Setup (computer, occasional)
Load the Excel export. Parse the useful columns, split `Lat, Long`, store the active
signs. Re-runnable when the sheet changes — detect what's already stored and add the
new rows rather than duplicating.

### 2. Field (phone, daily)
- Start a job → choose the sign.
- Sign picker: search by ID/route/county, **and** "sort by nearest" using device GPS
  against each sign's stored coordinates (nearest at top, with distance).
- Up to two photo slots. Tap → camera → photo. Preview, retake, or clear.
- Each photo is saved with its capture timestamp and the device GPS at capture time.
- Save and move on. Works offline; captures are queued on-device.

### 3. Flush (computer, end of day)
- Review the day's captures with their computed filenames.
- Export straight into a chosen folder (correctly named), or download.
- Notify the engineer by email: a plain list of the filenames/IDs added. The recipient
  list is remembered (type once, it's in the dropdown next time; only saved when a send
  actually happens).

## Architecture

- **Frontend:** static, dependency-free web app on **GitHub Pages**. Installable PWA so
  it runs full-screen on the phone and **works offline** — essential at rural bridges.
- **Backend (Phase 2):** **Supabase** — Postgres for sign data + captures, Storage for
  the photo blobs, Edge Function for the notification email. Free tier is enough to
  start.
- **Email:** Supabase Edge Function (transactional email) is the default. A Google Apps
  Script sender remains a fallback option.

## Backend (implemented)

Supabase project `bridge-sign-helper` (org `mammer55-cloud's Org`), free tier.

- Tables: `signs`, `captures` (unique on `sign_id, slot, batch_date`), `recipients`.
- Storage bucket `sign-photos` (public read via object URL).
- Row Level Security is **on**; the publishable/anon key may read+write (no login yet).
- The front end talks to PostgREST + Storage over `fetch` (no SDK) — see `supabase.js`.

## Roadmap

- **Phase 1 — client-side experiment.** ✅ Phone flow end to end with placeholder data.
- **Phase 2 — full connected workflow.** ✅ Excel import → shared DB + photo storage →
  offline-first capture with upload status → review/export → mailto notification with
  remembered recipients.
- **Phase 3 — real data + polish.** Real R12-6 import (Thursday) and the bridge-sign
  dataset (`…N`/`…L`/`…R`); **authentication / access control** (top item, see below);
  automated email send (Edge Function or Apps Script) instead of `mailto`; optional
  "mark exported / purge server copies" after a flush.

## Open questions (for later)

- Which mailbox sends the notification (Supabase transactional vs. a Google account)?
- What, if anything, do we retain in the database after a successful flush?
- Do we want geolocation to *auto-suggest* the nearest sign, or only offer it on tap?
  (Current choice: offer on tap, never auto-select — avoids the L/R mix-up.)
