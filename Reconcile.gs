/**
 * Masters-System · Reconcile (v0.3)
 * -----------------------------------------------------------------------------
 * One-time migration tool: pulls existing master data from Finance / IMS / PM
 * into Rec_<Entity> sheets, where the admin reviews, dedup-merges, and
 * approves rows. reconcile_import() then writes approved rows into the real
 * Masters entity sheets.
 *
 * Needs 3 config keys (set via the Config sheet before first pull):
 *   finance_sheet_id   — Finance spreadsheet ID
 *   ims_sheet_id       — IMS spreadsheet ID
 *   pm_sheet_id        — PM spreadsheet ID
 *
 * The modular@ account must have at least read access to all three — the
 * script runs as the user clicking the menu item.
 *
 * Rec_<Entity> columns (same shape across all four entities):
 *   Status | Target_Master_ID | Source | Source_ID | Source_Name |
 *   Normalized | Payload (JSON) | Notes
 *
 * Status values:
 *   pending   — not yet reviewed (default)
 *   approved  — on import, create a Masters row from this one
 *   skipped   — ignore
 *
 * Target_Master_ID is informational for now (alias map is v0.4).
 */

const REC_SHEETS = {
  Projects:  'Rec_Projects',
  Customers: 'Rec_Customers',
  Parties:   'Rec_Parties',
  Employees: 'Rec_Employees'
};

const REC_HEADERS = [
  'Status', 'Target_Master_ID', 'Source', 'Source_ID', 'Source_Name',
  'Normalized', 'Payload', 'Notes'
];

const REC_STATUS_VALUES = ['pending', 'approved', 'skipped'];

// -----------------------------------------------------------------------------
// Pull — public entry points (menu wires them up)
// -----------------------------------------------------------------------------

function reconcile_pullProjects() {
  requireAdmin_();
  const rows = [];
  const fin = _openSrcSS_('finance_sheet_id', 'Finance');

  // Finance.M_Project — Project_Code | Project_Name | Active
  _readSheet_(fin, 'M_Project').forEach(function (r) {
    const code = String(r['Project_Code'] || '').trim();
    const name = String(r['Project_Name'] || '').trim();
    if (!name) return;
    rows.push({
      source: 'Finance',
      sourceId: code,
      sourceName: name,
      payload: { Name: name, Project_Code: code, Active: r['Active'] }
    });
  });

  // PM.Projects — headers are 'Project ID' and 'Project Name' (spaces,
  // not underscores). Soft: skip quietly if sheet missing or empty.
  try {
    const pm = _openSrcSS_('pm_sheet_id', 'PM');
    _readSheet_(pm, 'Projects').forEach(function (r) {
      const id   = String(r['Project ID']   || r['Project_ID']   || '').trim();
      const name = String(r['Project Name'] || r['Project_Name'] || '').trim();
      if (!id && !name) return;
      rows.push({
        source: 'PM',
        sourceId: id,
        sourceName: name,
        payload: r
      });
    });
  } catch (e) {
    Logger.log('PM Projects pull skipped: ' + e.message);
  }

  _writeRecSheet_(REC_SHEETS.Projects, rows);
  SpreadsheetApp.getUi().alert('Rec_Projects built: ' + rows.length + ' rows.');
}

function reconcile_pullCustomers() {
  requireAdmin_();
  const rows = [];
  const fin = _openSrcSS_('finance_sheet_id', 'Finance');

  // Finance.M_Party, filtered to Party_Type=Customer
  _readSheet_(fin, 'M_Party').forEach(function (r) {
    if (String(r['Party_Type'] || '').trim().toLowerCase() !== 'customer') return;
    const name = String(r['Name'] || '').trim();
    if (!name) return;
    rows.push({
      source: 'Finance',
      sourceId: String(r['Party_ID'] || '').trim(),
      sourceName: name,
      payload: r
    });
  });

  // IMS has no customer master — Customer Name is freeform on FG Delivery.
  // Pull distinct customer names from FG Delivery Register as low-confidence hints.
  try {
    const ims = _openSrcSS_('ims_sheet_id', 'IMS');
    const seen = {};
    _readSheet_(ims, 'FG Delivery Register').forEach(function (r) {
      const name = String(r['Customer Name'] || '').trim();
      if (!name) return;
      const key = normName(name);
      if (seen[key]) return;
      seen[key] = true;
      rows.push({
        source: 'IMS',
        sourceId: '',
        sourceName: name,
        payload: { Name: name, _origin: 'FG Delivery Register (freeform)' },
        notes: 'freeform from IMS deliveries — match to Finance customer if same'
      });
    });
  } catch (e) {
    Logger.log('IMS customer pull skipped: ' + e.message);
  }

  _writeRecSheet_(REC_SHEETS.Customers, rows);
  SpreadsheetApp.getUi().alert('Rec_Customers built: ' + rows.length + ' rows.');
}

function reconcile_pullParties() {
  requireAdmin_();
  const rows = [];
  const fin = _openSrcSS_('finance_sheet_id', 'Finance');
  const ims = _openSrcSS_('ims_sheet_id', 'IMS');

  // Finance.M_Party where Party_Type=Vendor
  _readSheet_(fin, 'M_Party').forEach(function (r) {
    if (String(r['Party_Type'] || '').trim().toLowerCase() !== 'vendor') return;
    const name = String(r['Name'] || '').trim();
    if (!name) return;
    rows.push({
      source: 'Finance',
      sourceId: String(r['Party_ID'] || '').trim(),
      sourceName: name,
      payload: Object.assign({ _flag: 'Is_Supplier' }, r)
    });
  });

  // IMS.Supplier Master — Supplier Code | Firm Name | Category | ...
  _readSheet_(ims, 'Supplier Master').forEach(function (r) {
    const name = String(r['Firm Name'] || '').trim();
    if (!name) return;
    rows.push({
      source: 'IMS-Supplier',
      sourceId: String(r['Supplier Code'] || '').trim(),
      sourceName: name,
      payload: Object.assign({ _flag: 'Is_Supplier' }, r)
    });
  });

  // IMS.Job Worker Master — Job Worker Code | Firm Name | Process Type | ...
  _readSheet_(ims, 'Job Worker Master').forEach(function (r) {
    const name = String(r['Firm Name'] || '').trim();
    if (!name) return;
    rows.push({
      source: 'IMS-JobWorker',
      sourceId: String(r['Job Worker Code'] || '').trim(),
      sourceName: name,
      payload: Object.assign({ _flag: 'Is_JobWorker' }, r)
    });
  });

  _writeRecSheet_(REC_SHEETS.Parties, rows);
  SpreadsheetApp.getUi().alert('Rec_Parties built: ' + rows.length + ' rows.');
}

function reconcile_pullEmployees() {
  requireAdmin_();
  const rows = [];
  const fin = _openSrcSS_('finance_sheet_id', 'Finance');

  // Finance.M_Employee — Name | Email | Phone | Designation | Gross_Monthly_Salary | Active | Daily_Rate | Holds_Float
  _readSheet_(fin, 'M_Employee').forEach(function (r) {
    const name = String(r['Name'] || '').trim();
    if (!name) return;
    rows.push({
      source: 'Finance',
      sourceId: '',   // Finance employees have no stable ID
      sourceName: name,
      payload: r
    });
  });

  // PM.People — depends on columns, pull generically
  try {
    const pm = _openSrcSS_('pm_sheet_id', 'PM');
    _readSheet_(pm, 'People').forEach(function (r) {
      const name = String(r['Name'] || r['Person'] || r['People'] || '').trim();
      if (!name) return;
      rows.push({
        source: 'PM',
        sourceId: '',
        sourceName: name,
        payload: r
      });
    });
  } catch (e) {
    Logger.log('PM People pull skipped: ' + e.message);
  }

  // IMS.Dropdowns — Employees column (ninth col by SHEET_SPECS)
  try {
    const ims = _openSrcSS_('ims_sheet_id', 'IMS');
    const names = _readDropdownsColumn_(ims, 'Employees');
    names.forEach(function (name) {
      rows.push({
        source: 'IMS',
        sourceId: '',
        sourceName: name,
        payload: { Name: name, _origin: 'IMS Dropdowns.Employees' }
      });
    });
  } catch (e) {
    Logger.log('IMS Dropdowns pull skipped: ' + e.message);
  }

  _writeRecSheet_(REC_SHEETS.Employees, rows);
  SpreadsheetApp.getUi().alert('Rec_Employees built: ' + rows.length + ' rows.');
}

// -----------------------------------------------------------------------------
// Import — read approved rows from Rec_<entity>, create Masters rows
// -----------------------------------------------------------------------------

function reconcile_importProjects()  { return _importFromRec_('Projects');  }
function reconcile_importCustomers() { return _importFromRec_('Customers'); }
function reconcile_importParties()   { return _importFromRec_('Parties');   }
function reconcile_importEmployees() { return _importFromRec_('Employees'); }

function _importFromRec_(entity) {
  requireAdmin_();
  const recName = REC_SHEETS[entity];
  if (!recName) throw new Error('No reconciliation sheet for ' + entity);
  const sh = sh_(recName);
  const last = sh.getLastRow();
  if (last < 2) {
    SpreadsheetApp.getUi().alert('No rows in ' + recName + '.');
    return;
  }
  const headers = sh.getRange(1, 1, 1, REC_HEADERS.length).getValues()[0];
  const vals = sh.getRange(2, 1, last - 1, REC_HEADERS.length).getValues();
  const idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  let created = 0, skippedStatus = 0, skippedAlready = 0, errors = [];
  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];
    const status = String(row[idx['Status']] || '').trim().toLowerCase();
    if (status !== 'approved') { skippedStatus++; continue; }
    // Idempotency: if Target_Master_ID is already filled, this row was imported.
    const existingId = String(row[idx['Target_Master_ID']] || '').trim();
    if (existingId) { skippedAlready++; continue; }
    let payload = {};
    try { payload = JSON.parse(row[idx['Payload']] || '{}'); } catch (_) {}
    const sourceName = String(row[idx['Source_Name']] || '').trim();
    const data = _buildMasterPayload_(entity, sourceName, payload);
    try {
      const saved = entity_save(entity, data);
      sh.getRange(i + 2, idx['Target_Master_ID'] + 1).setValue(saved[SCHEMA[entity].idCol]);
      created++;
    } catch (e) {
      errors.push('Row ' + (i + 2) + ': ' + e.message);
    }
  }

  const msg = 'Imported into ' + entity + ':\n' +
    '  created: ' + created + '\n' +
    '  already-imported (skipped): ' + skippedAlready + '\n' +
    '  not approved (skipped): ' + skippedStatus +
    (errors.length ? '\n  errors (' + errors.length + '):\n    ' + errors.slice(0, 10).join('\n    ') : '');
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Map a Rec_<Entity> row (source name + payload) to a Masters entity row.
 * Only the fields the Masters schema has — anything else is dropped.
 * Finance's data is the canonical source for every entity.
 */
function _buildMasterPayload_(entity, sourceName, payload) {
  if (entity === 'Projects') {
    // Finance.Active=No usually means the project is finished (or cancelled —
    // user can refine in Masters). Default completed ones to Delivered,
    // active ones to Enquiry.
    const active = String(payload['Active'] || payload['Active?'] || 'Yes').toLowerCase() !== 'no';
    return {
      Name: sourceName,
      Status: active ? 'Enquiry' : 'Delivered',
      'Active?': active ? 'Yes' : 'No'
    };
  }
  if (entity === 'Customers') {
    return {
      Name: sourceName,
      GSTIN: payload['GSTIN'] || '',
      Email: payload['Email'] || '',
      Phone: payload['Phone'] || '',
      Billing_Address: payload['Address'] || '',
      'Active?': 'Yes'
    };
  }
  if (entity === 'Parties') {
    const flag = payload['_flag'] || 'Is_Supplier';
    const out = {
      Name: sourceName,
      GSTIN: payload['GSTIN'] || payload['GST No'] || '',
      Email: payload['Email'] || '',
      Phone: payload['Phone'] || payload['Mobile No'] || '',
      Address: payload['Address'] || '',
      'Is_Supplier?':   'No',
      'Is_JobWorker?':  'No',
      'Is_Transporter?':'No',
      'Is_ServiceProv?':'No',
      'Active?': 'Yes'
    };
    if (flag === 'Is_JobWorker') out['Is_JobWorker?'] = 'Yes';
    else out['Is_Supplier?'] = 'Yes';
    return out;
  }
  if (entity === 'Employees') {
    return {
      Name: sourceName,
      Phone: payload['Phone'] || '',
      Role: payload['Designation'] || payload['Role'] || '',
      Daily_Rate: payload['Daily_Rate'] || '',
      'Holds_Float?': String(payload['Holds_Float'] || 'No').toLowerCase() === 'yes' ? 'Yes' : 'No',
      Status: String(payload['Active'] || 'Yes').toLowerCase() === 'no' ? 'Left' : 'Active'
    };
  }
  throw new Error('Unknown entity: ' + entity);
}

// -----------------------------------------------------------------------------
// Internals
// -----------------------------------------------------------------------------

function _openSrcSS_(configKey, label) {
  const id = (getConfigValue(configKey) || '').trim();
  if (!id) throw new Error(label + ' spreadsheet ID missing. Set "' + configKey + '" in the Config sheet.');
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('Cannot open ' + label + ' (' + configKey + '=' + id + '): ' + e.message);
  }
}

function _readSheet_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet "' + sheetName + '" not found in ' + ss.getName());
  const last = sh.getLastRow();
  const cols = sh.getLastColumn();
  if (last < 2 || cols < 1) return [];
  const headers = sh.getRange(1, 1, 1, cols).getValues()[0].map(function (h) { return String(h).trim(); });
  const vals = sh.getRange(2, 1, last - 1, cols).getValues();
  return vals.map(function (r) {
    const o = {};
    headers.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
}

/**
 * IMS keeps per-dropdown lists in a wide 'Dropdowns' sheet with named columns.
 * Pull the column whose header matches `columnName`, return distinct non-empty.
 */
function _readDropdownsColumn_(ss, columnName) {
  const sh = ss.getSheetByName('Dropdowns');
  if (!sh) throw new Error('Dropdowns sheet not found');
  const last = sh.getLastRow();
  const cols = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, cols).getValues()[0].map(function (h) { return String(h).trim(); });
  const colIdx = headers.indexOf(columnName);
  if (colIdx < 0) throw new Error('Column "' + columnName + '" not in Dropdowns');
  if (last < 2) return [];
  const vals = sh.getRange(2, colIdx + 1, last - 1, 1).getValues();
  const seen = {}, out = [];
  vals.forEach(function (r) {
    const v = String(r[0] || '').trim();
    if (!v || seen[v]) return;
    seen[v] = true;
    out.push(v);
  });
  return out;
}

/**
 * Build / overwrite a Rec_<Entity> sheet.
 * Sorts by Normalized so cross-source duplicates sit next to each other, and
 * auto-dedups within each normalized-name group: ONE row (highest-priority
 * source) stays `pending`, siblings become `skipped`. Admin un-skips to
 * override. Source priority: Finance > IMS-Supplier > IMS-JobWorker > PM > IMS.
 */
const _REC_SOURCE_PRIORITY = {
  'Finance':       1,
  'IMS-Supplier':  2,
  'IMS-JobWorker': 2,
  'PM':            3,
  'IMS':           4
};

/**
 * Clean-case a human name as seen in source systems. Rules:
 *   - underscores → spaces, collapse whitespace
 *   - each word's first char uppercased
 *   - long all-caps words (>4 chars) title-cased (FASTECH → Fastech)
 *   - short all-caps left alone (BKC, L&T, G35, AC)
 * Idempotent on already-clean strings.
 */
function _cleanName_(s) {
  if (!s) return '';
  const spaced = String(s).replace(/_+/g, ' ').replace(/\s+/g, ' ').trim();
  return spaced.split(' ').map(function (w) {
    if (!w) return w;
    const isAllUpper = w === w.toUpperCase();
    if (isAllUpper && w.length > 4) {
      return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
    }
    return w.charAt(0).toUpperCase() + w.slice(1);
  }).join(' ');
}

/** Key used to match a row across pulls so we preserve Status + Target_Master_ID. */
function _recRowKey_(source, sourceId, normalized) {
  return source + '|' + (sourceId || normalized);
}

/** Read Status + Target_Master_ID from a Rec_* sheet, keyed by _recRowKey_. */
function _readPriorRecState_(ss, sheetName) {
  const out = {};
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return out;
  const last = sh.getLastRow();
  if (last < 2) return out;
  const headers = sh.getRange(1, 1, 1, REC_HEADERS.length).getValues()[0];
  const idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });
  const vals = sh.getRange(2, 1, last - 1, REC_HEADERS.length).getValues();
  vals.forEach(function (r) {
    const source = String(r[idx['Source']] || '').trim();
    const sourceId = String(r[idx['Source_ID']] || '').trim();
    const normalized = String(r[idx['Normalized']] || '').trim();
    if (!source) return;
    out[_recRowKey_(source, sourceId, normalized)] = {
      status:   String(r[idx['Status']] || '').trim().toLowerCase(),
      targetId: String(r[idx['Target_Master_ID']] || '').trim()
    };
  });
  return out;
}

/** True if `a` is a better dedup rep than `b`: longer name wins, tie → source. */
function _recRowBetter_(a, b) {
  if (a.normalized.length !== b.normalized.length) {
    return a.normalized.length > b.normalized.length;
  }
  const pa = _REC_SOURCE_PRIORITY[a.source] || 99;
  const pb = _REC_SOURCE_PRIORITY[b.source] || 99;
  return pa < pb;
}

function _writeRecSheet_(sheetName, rows) {
  const ss = ss_();
  // Preserve prior Status + Target_Master_ID across re-pulls — keyed by
  // Source|Source_ID when both present, else Source|Normalized.
  const prior = _readPriorRecState_(ss, sheetName);

  let sh = ss.getSheetByName(sheetName);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(sheetName);

  sh.getRange(1, 1, 1, REC_HEADERS.length)
    .setValues([REC_HEADERS])
    .setFontWeight('bold')
    .setBackground('#0F172A')
    .setFontColor('#FFFFFF');
  sh.setFrozenRows(1);

  // Clean-case source names (BKC_project → BKC Project), then normalize + sort.
  // User can still hand-edit Source_Name in the sheet before approving.
  rows.forEach(function (r) {
    r.sourceName = _cleanName_(r.sourceName);
    r.normalized = normName(r.sourceName);
  });
  rows.sort(function (a, b) {
    if (a.normalized !== b.normalized) return a.normalized < b.normalized ? -1 : 1;
    const pa = _REC_SOURCE_PRIORITY[a.source] || 99;
    const pb = _REC_SOURCE_PRIORITY[b.source] || 99;
    return pa - pb;
  });

  // One `pending` per name group. normMatch catches prefix dupes
  // ("ASHA PLASTIC" ≡ "ASHA PLASTIC PRODUCT", "Omprakash" ≡ "Omprakash Ram").
  // Within a group, keep the "better" row: longer normalized name wins
  // (fuller record), then source priority. Swap rep if a later row is better.
  let groupRep = null;
  rows.forEach(function (r) {
    if (!r.normalized) { r.defaultStatus = 'pending'; groupRep = null; return; }
    if (groupRep && normMatch(groupRep.normalized, r.normalized)) {
      if (_recRowBetter_(r, groupRep)) {
        // demote current rep, promote this row
        groupRep.defaultStatus = 'skipped';
        groupRep.notes = (groupRep.notes ? groupRep.notes + ' · ' : '') +
          'auto-skipped as dup of ' + r.source + (r.sourceId ? ':' + r.sourceId : '');
        r.defaultStatus = 'pending';
        groupRep = r;
      } else {
        r.defaultStatus = 'skipped';
        r.notes = (r.notes ? r.notes + ' · ' : '') + 'auto-skipped as dup of ' +
          groupRep.source + (groupRep.sourceId ? ':' + groupRep.sourceId : '');
      }
    } else {
      r.defaultStatus = 'pending';
      groupRep = r;
    }
  });

  if (rows.length) {
    const data = rows.map(function (r) {
      const key = _recRowKey_(r.source, r.sourceId, r.normalized);
      const kept = prior[key];
      return [
        kept ? kept.status : r.defaultStatus,
        kept ? kept.targetId : '',
        r.source,
        r.sourceId || '',
        r.sourceName,
        r.normalized,
        JSON.stringify(r.payload || {}),
        r.notes || ''
      ];
    });
    sh.getRange(2, 1, data.length, REC_HEADERS.length).setValues(data);

    // Status dropdown on column A
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(REC_STATUS_VALUES, true)
      .setAllowInvalid(false)
      .build();
    sh.getRange(2, 1, data.length, 1).setDataValidation(rule);
  }

  sh.setColumnWidth(1, 90);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 260);
  sh.setColumnWidth(6, 220);
  sh.setColumnWidth(7, 350);
  sh.setColumnWidth(8, 240);
}

// -----------------------------------------------------------------------------
// Menu wiring — called from Menu.gs
// -----------------------------------------------------------------------------

function menu_reconcilePullAll_() {
  reconcile_pullProjects();
  reconcile_pullCustomers();
  reconcile_pullParties();
  reconcile_pullEmployees();
}

// -----------------------------------------------------------------------------
// Reset — wipe Projects/Customers/Parties/Employees + Rec_* + reset counters.
// Keeps Sites (seeded), Config, Audit_Log headers.
// -----------------------------------------------------------------------------
function reconcile_resetMigration() {
  requireAdmin_();
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Reset migration data?',
    'This will DELETE all rows from:\n' +
    '  • Projects / Customers / Parties / Employees\n' +
    '  • Rec_Projects / Rec_Customers / Rec_Parties / Rec_Employees\n' +
    '  • Audit_Log\n\n' +
    'And reset ID counters (next_prj / next_cust / next_vend / next_emp) to 1.\n\n' +
    'Sites and Config are kept.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const ss = ss_();
  const entitySheets = ['Projects', 'Customers', 'Parties', 'Employees'];
  const recSheets    = ['Rec_Projects', 'Rec_Customers', 'Rec_Parties', 'Rec_Employees'];
  const extras       = ['Audit_Log'];

  function clearData(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const last = sh.getLastRow();
    if (last >= 2) sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
  }
  entitySheets.forEach(clearData);
  extras.forEach(clearData);
  recSheets.forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });

  // Reset counters
  ['next_prj', 'next_cust', 'next_vend', 'next_emp'].forEach(function (k) {
    setConfigValue_internal_(k, '1');
  });

  ui.alert('Reset complete. Pull again for a fresh migration.');
}
