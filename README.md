# Masters-System

One master to rule them all.

Central master-data registry for 3 Star Projects — the single source of truth
for **Projects, Customers, Parties (unified suppliers/jobworkers/transporters/
service providers), Employees and Sites**. Consumed by the Finance, IMS and
Project-Management Apps Script systems via a published Apps Script library.

---

## Why this exists

Each of the three operational systems used to maintain its own copy of this
data as free text or local lists. The result: project names drifted, vendors
appeared twice, employee spellings forked. Masters is the one place every
entity exists, with a stable ID; the three other systems **consume** these IDs
and show the current name on every refresh.

Full architectural rationale: see
`scratchpad/combined_analysis.md` in the parent working set.

---

## Status — v0.1 (scaffold)

What works:

| | |
|---|---|
| ✅ | Spreadsheet setup (`firstTimeInit`) — builds all seven sheets with headers, dropdowns, protection, seed data |
| ✅ | Permission tiers — Admin (`modular@`), Editors (5 named accounts), Viewers (everyone else) |
| ✅ | ID generation (`PRJ-###`, `CUST-####`, `VEND-####`, `EMP-####`, `SITE-##`) with per-entity counters |
| ✅ | Audit log — append-only history of all writes |
| ✅ | Admin UI shell (web app) — tabs, user role indicator, entity counts |
| ✅ | Public library API (read-only) — `getProjects`, `getCustomer`, `getParties({isSupplier:true})`, `resolveFreeformName`, etc. |

What's still TODO (planned for v0.2+):

- Per-entity CRUD (Create / Update / Soft-delete) in the Admin UI
- Write-side library API (`createProject`, `updateCustomer`, …)
- Reconciliation tool for mapping existing Finance/IMS/PM freeform names to Masters IDs
- `_Interlink_*` modules in the three consumer repos

---

## Permissions

| Email | Role |
|---|---|
| `modular@3starprojects.com` | Admin |
| `gsp@3starprojects.com` | Editor |
| `info@3starprojects.com` | Editor |
| `projects@3starprojects.com` | Editor |
| `production@3starprojects.com` | Editor |
| `sales@3starprojects.com` | Editor |
| Everyone else with sheet access | Viewer (read-only) |

Change these in the `Config` sheet (admin-only) once live. The seed values
here are only used on `firstTimeInit`.

---

## First-time setup (one-off)

1. Create a new, blank Google Sheet. Name it e.g. "**3SP Masters**".
2. Share it with the admin + all editor accounts (Editor permission on
   the sheet; the script enforces finer-grained protection on top).
3. In the sheet: **Extensions → Apps Script**.
4. Connect the Apps Script project to this repo:
   - In the editor sidebar, open the GitHub integration pane and link
     `ThreeStarProjects/Masters-System`, branch `main`.
   - Pull.
5. Set the script's timezone to **Asia/Kolkata** (should already be set from
   `appsscript.json`).
6. Run the function **`firstTimeInit`** once (editor top bar → select function →
   Run). It will:
   - Build the seven sheets with correct headers and formatting
   - Seed `Config` and `Sites` (Thane Office + Dombivli Factory)
   - Apply dropdowns and protection
7. Deploy as a **Web App**:
   - Deploy → New deployment → Web app
   - Execute as: **Me (modular@3starprojects.com)**
   - Who has access: **Anyone within 3starprojects.com**
   - Deploy, copy the URL — this is the Admin UI.
8. Publish as an **Apps Script Library**:
   - Project settings → copy the Script ID
   - Deploy → New deployment → Library → Description "v0.1 scaffold"
   - Note the deployment version number
   - This is what the three consumer repos will add in their `appsscript.json`
     once we start the interlink work.

---

## File layout

```
appsscript.json       Manifest — timezone, V8, OAuth scopes, webapp config
Config.gs             SCHEMA + seed data for every sheet. Edit here, not elsewhere.
Setup.gs              firstTimeInit, rebuildHeaders, rebuildDropdowns, diagnose
Menu.gs               ⚙ Masters menu on the bound spreadsheet
Lib.gs                Shared helpers: IDs, permissions, audit, name matching
MastersAPI.gs         Public library API — what Finance/IMS/PM call
WebApp.gs             doGet — serves the Admin UI
Admin.html            Admin SPA (tabbed, role-aware)
_styles.html          CSS (light + dark)
_helpers.html         Client-side JS helpers
README.md             This file
```

---

## Consumer repos

- [Finance-System](https://github.com/ThreeStarProjects/Finance-System)
- [IMS](https://github.com/ThreeStarProjects/IMS)
- [Project-Management](https://github.com/ThreeStarProjects/Project-Management)

Interlink work happens on `feat/masters-integration` branches on each of
those repos. New files only; existing files touched minimally (one line in
each `onOpen` to add the "🔗 Interlink" menu).

---

## Development workflow

- Edits land directly on `main` of this repo during build (nothing to break;
  no live consumers yet).
- Each new version increments `MS_VERSION` in `Config.gs`.
- The deployed Apps Script Library pins to a specific version — bump only
  after a release checklist is clean.

---

## License & attribution

Internal 3 Star Projects tooling. Not open source.
