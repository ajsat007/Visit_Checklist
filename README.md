# Smart Services — Tour Visit Checklist System (v2.0.0)

Enterprise Tour Visit Checklist for Project Managers, District Managers, and
Associate District Managers. Static frontend on GitHub Pages, Google Apps
Script REST API backend, Google Sheets database, Google Drive file storage.

---

## Architecture

```
GitHub Pages (frontend)          Google Apps Script (backend)         Google Workspace
┌─────────────────────┐          ┌──────────────────────┐            ┌──────────────┐
│ index.html   (login) │  HTTPS   │ TourVisit_Backend.gs  │            │ Google Sheets│
│ visit.html   (form)  │ ───────► │ PDF_Engine.gs          │ ─────────► │ (database)   │
│ dashboard.html(admin)│  fetch   │ Setup.gs / Assets.gs   │            │ Google Drive │
└─────────────────────┘          └──────────────────────┘            │ (PDFs/files) │
                                                                        └──────────────┘
```

---

## 1. Backend setup (Google Apps Script)

1. Go to [script.google.com](https://script.google.com) → New project.
2. Create four script files matching the names in `backend/`: `Setup.gs`,
   `Assets.gs`, `PDF_Engine.gs`, `TourVisit_Backend.gs` — paste each file's
   content in. Also open **Project Settings → paste `appsscript.json`'s
   content** into the manifest (enable "Show appsscript.json" first).
3. **Enable the Drive API advanced service** (required for PDF generation):
   Editor → Services (+) → find **Drive API** → Add.
4. In the function dropdown, select `runInitialSetup`, click **Run**. Grant
   the requested permissions. Check **Execution log** — it will confirm:
   - Connected to your spreadsheet (`1SGqVc3O...`)
   - Created any missing tabs (`Base`, `Users`, `Checklist_Master`,
     `Tour_Visit`, `Checklist_Response`, `Audit_Log`) with headers —
     **existing data in already-present tabs is never touched.**
   - Created the root "Tour Visit" Drive folder.
5. **Deploy → New deployment → type: Web app.**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Copy the `/exec` URL — you'll need it in step 2 below.

### Populate your sheets
- **Base**: District | Bus Station — one row per station.
- **Users**: Employee ID | Name | Designation | Role | District | Password
  | Status. Role is one of `Admin` / `PM` / `DM` / `ADM`. Status must be
  exactly `Active` for the account to log in. **Passwords are entered here
  manually by an Admin** — there's no self-registration or reset-password
  flow in this build.
- **Checklist_Master**: Question ID | Category | Question | Display Order
  | Active (`TRUE`/`FALSE`).

---

## 2. Frontend setup (GitHub Pages)

1. Open `assets/js/config.js` and replace `BASE_URL` with the `/exec` URL
   from step 1.5 above.
2. Push the `frontend/` contents to a GitHub repo (paths in this build are
   all **relative**, so it works whether the repo is served from a project
   subpath like `username.github.io/repo-name/` or a root/custom domain —
   no hardcoded absolute paths to fix).
3. Repo → Settings → Pages → Source: your default branch, root folder.
4. Visit the published URL — you should see the login page.

---

## 3. First login

Since passwords are entered manually (see step 1), add at least one Admin
row to the `Users` sheet before testing:

| Employee ID | Name | Designation | Role | District | Password | Status |
|---|---|---|---|---|---|---|
| EMP001 | (your name) | Administrator | Admin | | (choose a password) | Active |

Admin logs in → redirected to `dashboard.html`. PM/DM/ADM → redirected to
`visit.html`.

---

## Folder structure created automatically in Drive

```
Tour Visit/
  2026/
    July/
      <District>/
        <Bus Station>/
          Generated PDF/
          Signed Checklist/
```

---

## Key design decisions & known trade-offs

- **PDF generation**: HTML → temporary Google Doc (via Drive API v2's
  `convert: true`) → exported as PDF → temp Doc deleted. This is why the
  Drive API advanced service is required (step 1.3). If PDF generation
  fails on submit (e.g. a transient issue), the visit data is **not**
  lost — it's saved first — and the officer gets a "Regenerate PDF" retry
  button instead of losing their checklist.
- **QR codes**: via QuickChart.io (no API key). Google's old
  `chart.googleapis.com` QR endpoint was retired in 2024.
- **GPS Address**: best-effort reverse geocoding via OpenStreetMap
  Nominatim (free, no key, but rate-limited — fine for normal field use).
  If you later get a billed Google Maps API key, swap it into
  `reverseGeocode()` in `assets/js/utils.js` for higher reliability at
  scale.
- **Passwords**: stored in plain text in the Users sheet, per explicit
  requirement for this build. This is materially weaker than hashing —
  restrict edit access to the Users sheet to trusted Admins only, and
  treat the spreadsheet itself as sensitive.
- **Sessions**: capped at 6 hours — this is a hard ceiling in Apps
  Script's `CacheService`, not a configurable preference.
- **Role scoping**: PM/DM/ADM dashboards (if you build out a view for
  them later) are automatically restricted server-side to their own
  district; only Admin sees org-wide data. This is enforced in
  `TourVisit_Backend.gs`, not just hidden in the UI.

## Support / troubleshooting

- **"Drive is not defined" error on submit** → Drive API advanced service
  isn't enabled (step 1.3).
- **PDF has no logo** → check `Assets.gs` wasn't truncated when pasted.
- **Login always fails** → confirm the Users row's `Status` column is
  exactly `Active` (case-sensitive) and Password matches exactly.
- **CORS / blank response in browser console** → confirm the Web app
  deployment's access is set to "Anyone" and you're using the `/exec` URL
  (not `/dev`).
