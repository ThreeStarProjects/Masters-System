/**
 * Masters-System · Setup
 * -----------------------------------------------------------------------------
 * First-time init + idempotent rebuild of sheets, headers, dropdowns and
 * seed data. Safe to re-run — never destroys existing data rows.
 *
 * Run order on a brand-new spreadsheet:
 *   1. User creates a new Google Sheet
 *   2. User opens Extensions > Apps Script, pulls this repo from GitHub
 *   3. User runs firstTimeInit() from the editor OR via the ⚙ Masters menu
 */

/**
 * The one function the owner runs the first time on a fresh spreadsheet.
 * Builds all sheets, writes headers, seeds Config + Sites, applies
 * formatting + protection + data validation. Idempotent.
 */
function firstTimeInit() {
  const ss = ss_();
  SpreadsheetApp.flush();

  // 1. Create every sheet declared in SCHEMA, in order.
  const sheetOrder = [SH.CONFIG, SH.PROJECTS, SH.CUSTOMERS, SH.PARTIES, SH.EMPLOYEES, SH.SITES, SH.AUDIT];
  sheetOrder.forEach((name, i) => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name, i);
    } else {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(i + 1);
    }
  });

  // 2. Remove the default "Sheet1" if it's still lying around and empty.
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && sheet1.getLastRow() <= 1 && sheet1.getLastColumn() <= 1) {
    ss.deleteSheet(sheet1);
  }

  // 3. Write headers for every sheet (idempotent — overwrites row 1 only).
  sheetOrder.forEach(name => applyHeaders_(name));

  // 4. Seed Config if empty.
  seedConfigIfEmpty_();

  // 5. Seed Sites if empty.
  seedSitesIfEmpty_();

  // 6. Apply data validations (dropdowns).
  applyAllValidations_();

  // 7. Apply protections (editor/admin ranges).
  applyAllProtections_();

  // 8. Freeze header rows + first-col banding.
  applyAllFormatting_();

  SpreadsheetApp.getActive().toast('Masters setup complete.', 'Masters', 5);
}

/**
 * Re-apply headers only (useful when SCHEMA changes). Data rows untouched.
 */
function rebuildHeaders() {
  Object.keys(SCHEMA).forEach(name => applyHeaders_(name));
  SpreadsheetApp.getActive().toast('Headers refreshed.', 'Masters', 3);
}

/**
 * Re-apply all dropdowns from current SCHEMA + referential data (Customers,
 * Sites, Employees). Safe to re-run.
 */
function rebuildDropdowns() {
  applyAllValidations_();
  SpreadsheetApp.getActive().toast('Dropdowns refreshed.', 'Masters', 3);
}

// -----------------------------------------------------------------------------
// Internals
// -----------------------------------------------------------------------------

function applyHeaders_(sheetName) {
  const schema = SCHEMA[sheetName];
  if (!schema) return;
  const sh = sh_(sheetName);
  const headers = schema.headers;
  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1f3a5f')
    .setFontColor('#ffffff');
  sh.setFrozenRows(1);
  if (schema.freezeCols) sh.setFrozenColumns(schema.freezeCols);
  // Make sure we have at least `headers.length` columns; trim any extras
  // beyond that (so Admin UI doesn't accidentally read phantom columns).
  const haveCols = sh.getMaxColumns();
  if (haveCols < headers.length) {
    sh.insertColumnsAfter(haveCols, headers.length - haveCols);
  } else if (haveCols > headers.length) {
    sh.deleteColumns(headers.length + 1, haveCols - headers.length);
  }
}

function seedConfigIfEmpty_() {
  const sh = sh_(SH.CONFIG);
  if (sh.getLastRow() < 2) {
    const rows = CONFIG_SEED.map(r => [r[0], r[1], r[2]]);
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
    return;
  }
  // Additive: append any CONFIG_SEED keys not already in the sheet.
  // Lets us ship new config keys to an existing deployment without wiping state.
  const existing = {};
  const last = sh.getLastRow();
  sh.getRange(2, 1, last - 1, 1).getValues().forEach(r => {
    existing[String(r[0]).trim()] = true;
  });
  CONFIG_SEED.forEach(r => { if (!existing[r[0]]) sh.appendRow([r[0], r[1], r[2]]); });
}

function seedSitesIfEmpty_() {
  const sh = sh_(SH.SITES);
  if (sh.getLastRow() >= 2) return;
  const now = nowStr_();
  const by = SEED_ADMIN_EMAILS[0];
  const rows = SITES_SEED.map(s => [
    s.id, s.name, s.type, s.address, 'Yes', '',
    now, by, now, by
  ]);
  sh.getRange(2, 1, rows.length, SCHEMA[SH.SITES].headers.length).setValues(rows);
}

function applyAllValidations_() {
  // Projects: Status, Customer_ID lookup, Site_ID lookup, Active?
  applyEnumValidation_(SH.PROJECTS, 'Status', SCHEMA[SH.PROJECTS].statusValues);
  applyLookupValidation_(SH.PROJECTS, 'Customer_ID',  SH.CUSTOMERS, 'Customer_ID');
  applyLookupValidation_(SH.PROJECTS, 'Site_ID',      SH.SITES,     'Site_ID');
  applyYesNoValidation_(SH.PROJECTS, 'Active?');

  // Customers: Active?
  applyYesNoValidation_(SH.CUSTOMERS, 'Active?');

  // Parties: Is_* flags + Active?
  ['Is_Supplier?','Is_JobWorker?','Is_Transporter?','Is_ServiceProv?','Active?'].forEach(col => {
    applyYesNoValidation_(SH.PARTIES, col);
  });

  // Employees: Status, Holds_Float?, Site_ID lookup
  applyEnumValidation_(SH.EMPLOYEES, 'Status', SCHEMA[SH.EMPLOYEES].statusValues);
  applyYesNoValidation_(SH.EMPLOYEES, 'Holds_Float?');
  applyLookupValidation_(SH.EMPLOYEES, 'Site_ID', SH.SITES, 'Site_ID');

  // Sites: Type, Active?
  applyEnumValidation_(SH.SITES, 'Type', SCHEMA[SH.SITES].typeValues);
  applyYesNoValidation_(SH.SITES, 'Active?');
}

function applyEnumValidation_(sheetName, colName, values) {
  const sh = sh_(sheetName);
  const col = SCHEMA[sheetName].headers.indexOf(colName) + 1;
  if (!col) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(rule);
}

function applyYesNoValidation_(sheetName, colName) {
  applyEnumValidation_(sheetName, colName, ['Yes', 'No']);
}

function applyLookupValidation_(sheetName, colName, srcSheet, srcCol) {
  const sh = sh_(sheetName);
  const col = SCHEMA[sheetName].headers.indexOf(colName) + 1;
  if (!col) return;
  const src = ss_().getSheetByName(srcSheet);
  if (!src) return;
  const srcColIdx = SCHEMA[srcSheet].headers.indexOf(srcCol) + 1;
  if (!srcColIdx) return;
  // Dropdown reads live from the source column (A2:A).
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      src.getRange(2, srcColIdx, src.getMaxRows() - 1, 1),
      true
    )
    .setAllowInvalid(true)    // allow blank; migration backfills later
    .build();
  sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(rule);
}

function applyAllProtections_() {
  const ss = ss_();
  Object.keys(SCHEMA).forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    const schema = SCHEMA[sheetName];
    // Clear existing protections first
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

    // Always protect the header row — nobody edits headers except via rebuild.
    const headerProt = sh.getRange(1, 1, 1, schema.headers.length).protect()
      .setDescription('Headers · do not edit');
    headerProt.setDomainEdit(false);
    headerProt.removeEditors(headerProt.getEditors());
    headerProt.addEditor(SEED_ADMIN_EMAILS[0]);

    // Sheet-level protection based on tier
    if (schema.protect === 'admin') {
      const p = sh.protect().setDescription(sheetName + ' · admin only');
      p.removeEditors(p.getEditors());
      SEED_ADMIN_EMAILS.forEach(e => p.addEditor(e));
    } else if (schema.protect === 'editor') {
      const p = sh.protect().setDescription(sheetName + ' · editors');
      p.removeEditors(p.getEditors());
      [].concat(SEED_ADMIN_EMAILS, SEED_EDITOR_EMAILS).forEach(e => p.addEditor(e));
    } else if (schema.protect === 'nobody') {
      const p = sh.protect().setDescription(sheetName + ' · append-only (admin may edit)');
      p.removeEditors(p.getEditors());
      SEED_ADMIN_EMAILS.forEach(e => p.addEditor(e));
    }
  });
}

function applyAllFormatting_() {
  Object.keys(SCHEMA).forEach(sheetName => {
    const sh = ss_().getSheetByName(sheetName);
    if (!sh) return;
    const nCols = Math.min(SCHEMA[sheetName].headers.length, sh.getMaxColumns());
    // Autosize first, then cap wide columns + enable wrap so text doesn't
    // overflow visually. 300px is a reasonable readable max for long names.
    sh.autoResizeColumns(1, nCols);
    for (let c = 1; c <= nCols; c++) {
      if (sh.getColumnWidth(c) > 300) sh.setColumnWidth(c, 300);
    }
    const lastRow = Math.max(sh.getLastRow(), 1);
    sh.getRange(1, 1, lastRow, nCols).setWrap(true).setVerticalAlignment('top');
  });
}

// -----------------------------------------------------------------------------
// Diagnostics
// -----------------------------------------------------------------------------
function diagnose() {
  const report = [];
  Object.keys(SCHEMA).forEach(name => {
    const sh = ss_().getSheetByName(name);
    if (!sh) {
      report.push('[MISSING] ' + name);
      return;
    }
    const headers = sh.getRange(1, 1, 1, SCHEMA[name].headers.length).getValues()[0];
    const expected = SCHEMA[name].headers;
    const mismatch = [];
    for (let i = 0; i < expected.length; i++) {
      if (String(headers[i]).trim() !== expected[i]) {
        mismatch.push('col ' + (i + 1) + ': expected "' + expected[i] + '", got "' + headers[i] + '"');
      }
    }
    report.push('[' + (mismatch.length ? 'DRIFT' : 'OK') + '] ' + name +
                ' (' + (sh.getLastRow() - 1) + ' rows)' +
                (mismatch.length ? '\n    ' + mismatch.join('\n    ') : ''));
  });
  const admin = (getConfigValue('admin_emails') || '').split(',').length;
  const editor = (getConfigValue('editor_emails') || '').split(',').length;
  report.push('[CONFIG] admin accounts: ' + admin + ', editor accounts: ' + editor);
  report.push('[VERSION] ' + MS_VERSION);
  const msg = report.join('\n');
  SpreadsheetApp.getUi().alert('Masters diagnostics', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  return msg;
}
