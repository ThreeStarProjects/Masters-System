/**
 * Masters-System · CRUD
 * -----------------------------------------------------------------------------
 * Generic entity CRUD, driven entirely by SCHEMA metadata.
 * One module, five entities. Add a sixth → edit SCHEMA; this file doesn't change.
 *
 * All functions are server-side; the Admin UI calls them via google.script.run,
 * and consumer systems call them through MastersAPI (which wraps these).
 */

// -----------------------------------------------------------------------------
// Read
// -----------------------------------------------------------------------------

/**
 * List rows of an entity. Returns every row.
 * ponytail: Date and other non-primitive values are coerced to strings before
 * returning — google.script.run silently drops rows whose values it can't
 * serialise, so we flatten them here.
 */
function entity_list(entityName, opts) {
  if (!SCHEMA[entityName]) throw new Error('Unknown entity: ' + entityName);
  var sh = ss_().getSheetByName(entityName);
  if (!sh) throw new Error('Sheet missing: ' + entityName);
  var last = sh.getLastRow();
  var cols = sh.getLastColumn();
  if (last < 2 || cols < 1) return [];
  var headers = sh.getRange(1, 1, 1, cols).getValues()[0];
  var values  = sh.getRange(2, 1, last - 1, cols).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var o = {};
    for (var j = 0; j < headers.length; j++) {
      var v = values[i][j];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      else if (v === undefined || v === null) v = '';
      else if (typeof v === 'object') v = String(v);
      o[String(headers[j])] = v;
    }
    out.push(o);
  }
  if (opts && opts.filter) {
    out = out.filter(function (r) { return matchesFilter_(r, opts.filter); });
  }
  return out;
}

/** Fetch one row by ID. */
function entity_get(entityName, id) {
  const schema = SCHEMA[entityName];
  if (!schema || !schema.idCol) throw new Error('Entity has no ID column: ' + entityName);
  if (!id) return null;
  const row = readAll_(entityName).find(function (r) {
    return String(r[schema.idCol]).trim() === String(id).trim();
  });
  return row ? stripInternal_(row) : null;
}

// -----------------------------------------------------------------------------
// Write
// -----------------------------------------------------------------------------

/**
 * Create or update a row. If the posted row has no ID (or an unknown ID), it's
 * a create — the ID is minted server-side. If the ID exists, it's an update.
 * Returns the saved row (with ID, timestamps).
 */
function entity_save(entityName, row) {
  requireEditor_();
  const schema = SCHEMA[entityName];
  if (!schema) throw new Error('Unknown entity: ' + entityName);
  const idCol = schema.idCol;
  const data = Object.assign({}, row || {});
  const email = getUserEmail();
  const now = nowStr_();

  return withLock_(function () {
    const sh = sh_(entityName);
    const headers = schema.headers;
    ensureColumns_(sh, headers);   // a column added to the schema later appears in the live sheet by itself
    const idx = {};
    headers.forEach(function (h, i) { idx[h] = i; });

    const incomingId = data[idCol] ? String(data[idCol]).trim() : '';
    const existing = incomingId ? _findRowNum(sh, idx[idCol], incomingId) : 0;

    if (existing) {
      // UPDATE
      const oldRow = sh.getRange(existing, 1, 1, headers.length).getValues()[0];
      const newRow = headers.map(function (h, i) {
        if (isAuditField_(h)) {
          if (h === 'Updated_At') return now;
          if (h === 'Updated_By') return email;
          return oldRow[i];       // keep Created_At / Created_By
        }
        if (h === idCol) return oldRow[i];   // ID is immutable
        if (data.hasOwnProperty(h)) return _coerce_(data[h], _fieldType_(schema, h));
        return oldRow[i];
      });
      sh.getRange(existing, 1, 1, headers.length).setValues([newRow]);
      _auditDiff_(entityName, incomingId, 'Update', headers, oldRow, newRow);
      return _rowToObj_(headers, newRow);
    }

    // CREATE
    const newId = incomingId || nextId_(entityName);
    const newRow = headers.map(function (h) {
      if (h === idCol) return newId;
      if (h === 'Created_At' || h === 'Updated_At') return now;
      if (h === 'Created_By' || h === 'Updated_By') return email;
      if (data.hasOwnProperty(h)) return _coerce_(data[h], _fieldType_(schema, h));
      // honour defaults from fieldTypes
      const ft = schema.fieldTypes && schema.fieldTypes[h];
      if (ft && ft.default !== undefined) return ft.default;
      return '';
    });
    sh.appendRow(newRow);
    audit_(entityName, newId, 'Create', '', '', JSON.stringify(_rowToObj_(headers, newRow)), '');
    return _rowToObj_(headers, newRow);
  });
}

/**
 * Soft-delete — set Active? to 'No'. Hard delete is not exposed.
 */
function entity_softDelete(entityName, id) {
  requireEditor_();
  const schema = SCHEMA[entityName];
  if (!schema) throw new Error('Unknown entity: ' + entityName);
  if (!id) throw new Error('ID required');
  if (schema.headers.indexOf('Active?') < 0) {
    throw new Error('Entity does not support soft delete: ' + entityName);
  }
  const current = entity_get(entityName, id);
  if (!current) throw new Error('Not found: ' + id);
  return entity_save(entityName, Object.assign({}, current, { 'Active?': 'No' }));
}

/** Hard delete — admin only. Avoid unless migrating.*/
function entity_hardDelete(entityName, id) {
  requireAdmin_();
  const schema = SCHEMA[entityName];
  if (!schema) throw new Error('Unknown entity: ' + entityName);
  const sh = sh_(entityName);
  const idx = headerIndex_(entityName)[schema.idCol];
  const rowNum = _findRowNum(sh, idx, id);
  if (!rowNum) throw new Error('Not found: ' + id);
  const old = sh.getRange(rowNum, 1, 1, schema.headers.length).getValues()[0];
  sh.deleteRow(rowNum);
  audit_(entityName, id, 'HardDelete', '', JSON.stringify(_rowToObj_(schema.headers, old)), '', 'Hard delete by admin');
  return true;
}

// -----------------------------------------------------------------------------
// Form metadata for the UI — single endpoint the client calls to render a form.
// Returns ordered field descriptors + lookup options for FK fields.
// -----------------------------------------------------------------------------
function entity_formMeta(entityName) {
  const schema = SCHEMA[entityName];
  if (!schema) throw new Error('Unknown entity: ' + entityName);
  const fields = schema.headers.map(function (name) {
    const explicit = schema.fieldTypes && schema.fieldTypes[name];
    const type = explicit && explicit.type ? explicit.type
      : (isAuditField_(name) ? 'audit'
        : (name.slice(-1) === '?' ? 'yesno' : 'text'));
    const descriptor = {
      name: name,
      type: type,
      required: !!(explicit && explicit.required),
      readonly: !!(explicit && explicit.readonly) || type === 'id' || type === 'audit',
      hint: (explicit && explicit.hint) || '',
      values: (explicit && explicit.values) || null,
      default: (explicit && explicit.default !== undefined) ? explicit.default : null,
      lookup: null
    };
    if (type === 'lookup' && explicit && explicit.from) {
      descriptor.lookup = {
        from: explicit.from,
        options: _lookupOptions_(explicit.from)
      };
    }
    return descriptor;
  });
  return {
    entity: entityName,
    idCol: schema.idCol,
    labelCol: schema.labelCol || schema.headers[1] || schema.headers[0],
    fields: fields
  };
}

function _lookupOptions_(sheetName) {
  const schema = SCHEMA[sheetName];
  if (!schema || !schema.idCol) return [];
  return readAll_(sheetName)
    .filter(function (r) { return String(r['Active?'] || 'Yes').toLowerCase() !== 'no'; })
    .map(function (r) {
      const labelCol = schema.labelCol || schema.headers[1] || schema.idCol;
      return {
        id: r[schema.idCol],
        label: String(r[schema.idCol]) + ' — ' + String(r[labelCol] || '')
      };
    });
}

// -----------------------------------------------------------------------------
// Internals
// -----------------------------------------------------------------------------

function _findRowNum(sh, idColIndex, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, idColIndex + 1, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(id).trim()) return i + 2;
  }
  return 0;
}

function _fieldType_(schema, header) {
  const ft = schema.fieldTypes && schema.fieldTypes[header];
  if (ft && ft.type) return ft.type;
  if (isAuditField_(header)) return 'audit';
  if (header.slice(-1) === '?') return 'yesno';
  return 'text';
}

function _coerce_(v, type) {
  if (v === undefined || v === null) return '';
  if (type === 'number') {
    if (v === '' || v === null) return '';
    const n = Number(v);
    return isNaN(n) ? '' : n;
  }
  if (type === 'date') {
    if (v === '' || v === null) return '';
    if (v instanceof Date) return v;
    const d = new Date(v);
    return isNaN(d.getTime()) ? v : d;
  }
  if (type === 'yesno') {
    const s = String(v).toLowerCase();
    if (s === 'yes' || s === 'true' || s === '1') return 'Yes';
    if (s === 'no'  || s === 'false'|| s === '0' || s === '') return 'No';
    return v;
  }
  return v;
}

function _rowToObj_(headers, row) {
  const o = {};
  headers.forEach(function (h, i) { o[h] = row[i]; });
  return o;
}

function _auditDiff_(entityName, id, action, headers, oldRow, newRow) {
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (isAuditField_(h)) continue;
    const o = oldRow[i], n = newRow[i];
    const oS = o instanceof Date ? o.toISOString() : String(o == null ? '' : o);
    const nS = n instanceof Date ? n.toISOString() : String(n == null ? '' : n);
    if (oS !== nS) audit_(entityName, id, action, h, oS, nS, '');
  }
}

function matchesFilter_(row, filter) {
  if (!filter) return true;
  return Object.keys(filter).every(function (k) {
    const want = String(filter[k]).toLowerCase();
    const have = String(row[k] == null ? '' : row[k]).toLowerCase();
    return have === want;
  });
}

function stripInternal_(row) {
  const o = Object.assign({}, row);
  delete o._row;
  return o;
}

/**
 * Make the sheet match its schema: room and a header for every schema column, in the schema's ORDER.
 *  - Headers that are blank are filled (nothing else is overwritten).
 *  - If the schema columns are present but in a different order (a column was moved in the schema), the sheet is REORDERED BY
 *    HEADER NAME with all data and number formats following their columns. A hidden backup copy of the sheet is made first.
 *    A column that is not in the schema is kept, moved to the right of the schema columns. Idempotent: a sheet already in order is untouched.
 * Safe to call on every save. Never runs on a sheet whose first header is not the schema's first header.
 */
function ensureColumns_(sh, headers) {
  var n = headers.length;
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());
  var width = Math.max(n, sh.getLastColumn());
  var cur = sh.getRange(1, 1, 1, width).getValues()[0].map(function (v) { return String(v).trim(); });
  var misplaced = cur[0] === headers[0] && headers.some(function (h, i) { return cur[i] !== h && cur.indexOf(h) >= 0; });
  if (misplaced) reorderToSchema_(sh, headers, cur, width);
  cur = sh.getRange(1, 1, 1, n).getValues()[0];
  for (var i = 0; i < n; i++) {
    if (String(cur[i]).trim() === "") {
      sh.getRange(1, i + 1).setValue(headers[i]).setFontWeight("bold").setBackground("#1f3a5f").setFontColor("#ffffff");
    }
  }
}

function reorderToSchema_(sh, headers, cur, width) {
  var last = sh.getLastRow();
  // 1. hidden backup first. If it cannot be made, nothing is moved.
  var bak = sh.copyTo(sh.getParent());
  bak.setName('_bak_' + sh.getName() + '_' + Utilities.formatDate(new Date(), 'GMT+5:30', 'yyyyMMdd_HHmm'));
  bak.hideSheet();
  // 2. new column order: schema columns in schema order (by name; -1 = not in the sheet yet), then any other named column
  var order = headers.map(function (h) { return cur.indexOf(h); });
  cur.forEach(function (h, i) { if (h !== '' && headers.indexOf(h) < 0) order.push(i); });
  var data = sh.getRange(1, 1, last, width).getValues();
  var fmts = sh.getRange(1, 1, last, width).getNumberFormats();
  var outV = data.map(function (row, r) {
    return order.map(function (i, k) { return r === 0 ? (k < headers.length ? headers[k] : cur[i]) : (i < 0 ? '' : row[i]); });
  });
  var outF = fmts.map(function (row) { return order.map(function (i) { return i < 0 ? 'General' : row[i]; }); });
  sh.getRange(1, 1, last, order.length).setValues(outV).setNumberFormats(outF);
  if (width > order.length) sh.getRange(1, order.length + 1, last, width - order.length).clearContent();
  sh.getRange(1, 1, 1, order.length).setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
}
