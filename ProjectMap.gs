/**
 * Masters-System · ProjectMap
 * -----------------------------------------------------------------------------
 * Text → Finance project pairings that Finance's "IMS material costs" import remembers.
 * Key = a PRJ-### id, or the exact upper-cased project text of an old IMS requisition
 * ("MAROL SITE"); value = the Finance project label it was paired with by hand.
 * Lives here so every system reads one answer. The sheet is created on first save.
 *
 * Library calls (from Finance):
 *   Masters.getProjectMap()            → { 'PRJ-001': 'PRJ-001 — Test', 'MAROL SITE': '...' }
 *   Masters.saveProjectMap({key: v})   → upsert; returns how many pairs were new / changed
 */

const PMAP_SHEET = 'Project_Map';
const PMAP_HEADERS = ['Key', 'Finance_Project', 'Updated_At', 'Updated_By'];

function getProjectMap() {
  const out = {};
  const sh = ss_().getSheetByName(PMAP_SHEET);
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0] !== '' && r[1] !== '') out[String(r[0]).trim().toUpperCase()] = String(r[1]);
  });
  return out;
}

function saveProjectMap(pairs) {
  requireEditor_();
  const keys = Object.keys(pairs || {}).filter(function (k) { return String(k).trim() && String(pairs[k]).trim(); });
  if (!keys.length) return { added: 0, changed: 0 };
  return withLock_(function () {
    const ss = ss_();
    const sh = ss.getSheetByName(PMAP_SHEET) || ss.insertSheet(PMAP_SHEET);
    if (sh.getLastRow() < 1) {
      sh.getRange(1, 1, 1, PMAP_HEADERS.length).setValues([PMAP_HEADERS]).setFontWeight('bold');
      sh.setFrozenRows(1); sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 260);
    }
    const rows = sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
    const at = {};  // upper-cased key → sheet row
    rows.forEach(function (r, i) { if (r[0] !== '') at[String(r[0]).trim().toUpperCase()] = i + 2; });
    const now = new Date(), me = getUserEmail();
    let added = 0, changed = 0, next = rows.length + 2;
    keys.forEach(function (k) {
      const key = String(k).trim().toUpperCase(), val = String(pairs[k]).trim();
      if (at[key]) {
        if (String(sh.getRange(at[key], 2).getValue()) === val) return;
        sh.getRange(at[key], 2, 1, 3).setValues([[val, now, me]]);
        audit_('Project_Map', key, 'update', 'Finance_Project', '', val, '');
        changed++;
      } else {
        sh.getRange(next, 1, 1, 4).setValues([[key, val, now, me]]);
        audit_('Project_Map', key, 'create', 'Finance_Project', '', val, '');
        at[key] = next++; added++;
      }
    });
    return { added: added, changed: changed };
  });
}
