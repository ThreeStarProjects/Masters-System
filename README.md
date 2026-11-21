# Masters-System

The master-data registry for 3 Star Projects and the Apps Script **library** the other systems read it through. Single source of truth for **Projects, Customers, Parties (suppliers, job workers, transporters, service providers), Employees and Sites**, each with a stable ID.

> Part of the 3SP family: Finance-System, IMS, Project-Management, Ticketing, 3SP_DocGen. See "How the systems fit together" at the bottom.

## Why it exists

Each system used to keep its own free-text lists: project names drifted, vendors appeared twice, employee spellings forked. Masters is where each thing exists once; the others consume the ID and show the current name.

## Who owns what

| Entity | Mastered in | Notes |
|---|---|---|
| Projects (`PRJ-###`) | **Project-Management** creates them; Masters keeps the record | PM sends every project automatically; Masters is what IMS, Finance, DocGen and Ticketing read. |
| Employees (`EMP-####`) | **Masters** | The only employee list: name, role (designation), department, email, phone, salary terms, status. Finance payroll, PM, IMS and Ticketing read it. |
| Customers (`CUST-####`) | **Finance** | Masters keeps a copy (Finance_ID stored). |
| Suppliers / job workers | **IMS** (SUPP- / JW- codes) | Masters keeps a copy (codes stored); PM reads suppliers from it. Finance keeps its own separate vendor list. |
| Sites | Masters | Thane Office, Dombivli Factory seeded. |

## What is here

- **Sheets:** Projects, Customers, Parties, Employees, Sites, Config, Audit_Log (append-only), Project_Map, Project_Deletes, Rec_* (reconciliation staging).
- **Library API** (`MastersAPI.gs`): `getProjects / getCustomers / getParties({isSupplier...}) / getEmployees / getSites` and their create / update / soft-delete forms, `resolveFreeformName(entity, text)`, `getProjectLabels()`, `addEmployeeChecked` (duplicate-safe), `fillEmployeeBlanks`, `syncCustomersFromFinance`, `syncPartiesFromIms`, `getProjectMap / saveProjectMap`, `previewProjectDelete / executeProjectDelete`, `registerSystemSheet`.
- **Project delete everywhere** (`ProjectDelete.gs`): previews what would go from PM, IMS and Finance, removes it in order (Finance, IMS, PM), copies every removed row to a hidden sheet first, then removes the Masters record. Matches the exact `PRJ-###`.
- **Admin web app** (`Admin.html`, `WebApp.gs`) and a bound-sheet menu (`Menu.gs`); `Reconcile.gs` holds the one-time import tools (their menu is retired).
- **Permissions:** Admin and Editor accounts are listed in the Config sheet (`admin_emails`, `editor_emails`); everyone else with sheet access is read-only. Writes through the library need an editor.
- **Audit:** every write is logged (who, when, what changed). Deletes are soft (`Active? = No`) except the explicit project delete.

## Library rules (important)

- **Library code must never use `SpreadsheetApp.getActive()`**: inside a library call it returns the *caller's* sheet. Masters always uses `openById` (`ss_()` in `Lib.gs`).
- Every change to Masters needs a **new library version** (Deploy > Manage deployments), and each consumer must bump the version in its `appsscript.json`. After changing the column **order** of a schema, re-pin all consumers in the same session.
- Always `clasp pull` and compare before pushing Masters: an editor-only edit once got overwritten by a plain push.
- `appsscript.json` carries the OAuth scopes (including `userinfo.email`) callers need so `getUserEmail()` works through the library.

## First-time setup (one-off)

1. Create a blank Google Sheet "3SP Masters", share with the admin and editor accounts.
2. Extensions > Apps Script, push this repo with `clasp`, run **`firstTimeInit`** (builds the sheets, headers, dropdowns, protection and seed data).
3. Deploy as a Web App (Execute as *Me*, access within the company domain) for the admin UI, and as a **Library** for the other systems.

## Files

`Config.gs` schema and seed data (edit here) · `Setup.gs` first-time init and rebuilds · `Menu.gs` · `Lib.gs` helpers, ids, permissions, audit · `CRUD.gs` entity read / save / soft delete, column ordering by header name · `MastersAPI.gs` the public library surface · `ProjectMap.gs`, `ProjectDelete.gs`, `Reconcile.gs` · `WebApp.gs` + `Admin.html`, `_styles.html`, `_helpers.html` the admin UI.

## Deploy / development

- `clasp push --force`, commit, push to `main`. The GitHub Action (`.github/workflows/deploy.yml`) repeats the clasp push and needs the repo secret `CLASPRC_JSON` (if it fails with `invalid_rapt`, run `clasp login` and refresh the secret). Pushing to `main` does **not** publish a library version: do that in the editor.
- Offline tests (Node `vm`) are kept in `_tests` next to the repos, not in git.

## How the systems fit together

```
              Masters-System  (library: projects, customers, parties, employees, sites)
             /     |      |      \
     Finance   IMS   Project-Mgmt  Ticketing         3SP_DocGen (Next.js) --web app--> IMS
        ^        ^        |                              (documents, numbers, PDFs)
        |        |--reservations / requests-----|
        |--- 10-minute import of project material costs from IMS
```
