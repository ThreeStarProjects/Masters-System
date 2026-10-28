/**
 * Masters-System · Lib
 * -----------------------------------------------------------------------------
 * Shared helpers: IDs, permissions, audit, locking, name matching.
 * Everything in here is pure utility — no sheet writes outside of audit().
 */

// -----------------------------------------------------------------------------
// Spreadsheet access
// -----------------------------------------------------------------------------
function ss_() { return SpreadsheetApp.getActive(); }
function sh_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Sheet missing: ' + name + '. Run First-time setup.');
  return s;
}

// -----------------------------------------------------------------------------
// User & permissions
// -----------------------------------------------------------------------------
function getUserEmail() {
  return (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').toLowerCase();
}

function getConfigList_(key) {
  const raw = getConfigValue(key) || '';
  return raw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
}

function isAdmin(email) {
  const who = (email || getUserEmail()).toLowerCase();
  return getConfigList_('admin_emails').indexOf(who) >= 0;
}

function isEditor(email) {
  const who = (email || getUserEmail()).toLowerCase();
  if (isAdmin(who)) return true;
  return getConfigList_('editor_emails').indexOf(who) >= 0;
}

function requireAdmin_()  { if (!isAdmin())  throw new Error('Admin only.'); }
function requireEditor_() { if (!isEditor()) throw new Error('Editors only.'); }

// -----------------------------------------------------------------------------
// Config read / write (Config sheet key-value)
// -----------------------------------------------------------------------------
function getConfigValue(key) {
  const sh = ss_().getSheetByName(SH.CONFIG);
  if (!sh) return '';
  const last = sh.getLastRow();
  if (last < 2) return '';
  const vals = sh.getRange(2, 1, last - 1, 2).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === key) return String(vals[i][1]);
  }
  return '';
}

function setConfigValue(key, value) {
  requireAdmin_();
  const sh = sh_(SH.CONFIG);
  const last = sh.getLastRow();
  const vals = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === key) {
      sh.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value, '']);
}

// -----------------------------------------------------------------------------
// ID generation — atomic per-entity counter in Config sheet.
// Call under withLock_() when minting IDs in a loop.
// -----------------------------------------------------------------------------
function nextId_(entitySheetName) {
  const schema = SCHEMA[entitySheetName];
  if (!schema || !schema.idPrefix) throw new Error('No ID scheme for ' + entitySheetName);
  const key = schema.counterKey;
  const n = parseInt(getConfigValue(key) || '1', 10);
  const padded = String(n).padStart(schema.idPad, '0');
  const id = schema.idPrefix + '-' + padded;
  setConfigValue_internal_(key, String(n + 1));
  return id;
}

// Internal setter — bypasses admin check, used by nextId_ which runs as
// whatever user is creating a row. Writes still require editor-level above.
function setConfigValue_internal_(key, value) {
  const sh = sh_(SH.CONFIG);
  const last = sh.getLastRow();
  const vals = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === key) {
      sh.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value, '']);
}

// -----------------------------------------------------------------------------
// Locking — serialize writes that mint IDs or touch counters.
// -----------------------------------------------------------------------------
function withLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// -----------------------------------------------------------------------------
// Audit log — one row per create/update/delete/rename. Never edited later.
// -----------------------------------------------------------------------------
function audit_(entityType, entityId, action, field, oldVal, newVal, note) {
  const sh = ss_().getSheetByName(SH.AUDIT);
  if (!sh) return;    // audit sheet missing = we're mid-setup; skip silently
  sh.appendRow([
    new Date(),
    getUserEmail(),
    entityType,
    entityId || '',
    action,
    field || '',
    oldVal == null ? '' : String(oldVal),
    newVal == null ? '' : String(newVal),
    note || ''
  ]);
}

// -----------------------------------------------------------------------------
// Name matching — ported from Finance's _normName/_normMatch so migration-time
// name resolution behaves consistently with Finance's existing dedup logic.
// -----------------------------------------------------------------------------
const LEGAL_SUFFIX_RE = /\b(pvt|private|ltd|limited|llp|inc|co|company|corp|corporation|and|&|\.)\b/gi;

function normName(s) {
  if (!s) return '';
  return String(s)
    .toLowerCase()
    .replace(LEGAL_SUFFIX_RE, ' ')
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normMatch(a, b) {
  const na = normName(a), nb = normName(b);
  if (!na || !nb) return false;
  if (na === nb) return true;
  if (na.length >= 6 && nb.length >= 6) {
    if (na.startsWith(nb) || nb.startsWith(na)) return true;
  }
  return false;
}

// -----------------------------------------------------------------------------
// Row helpers — fetch all data rows of a sheet as array of header-keyed objects
// -----------------------------------------------------------------------------
function readAll_(sheetName) {
  const sh = sh_(sheetName);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  return rows.map(r => {
    const o = {};
    headers.forEach((h, i) => { o[h] = r[i]; });
    o._row = 0;  // populated below
    return o;
  }).map((o, i) => { o._row = i + 2; return o; });
}

function headerIndex_(sheetName) {
  const sh = sh_(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });
  return idx;
}

// -----------------------------------------------------------------------------
// Timestamp helpers for Created_At / Updated_At
// -----------------------------------------------------------------------------
function nowStr_() { return new Date(); }
