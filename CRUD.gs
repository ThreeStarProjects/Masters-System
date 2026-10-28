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
 * Returns a per-row _diag field so we can see server-side state inline.
 */
function entity_list(entityName, opts) {
  if (!SCHEMA[entityName]) {
    return [{ _diag: 'UNKNOWN_ENTITY', entityName: entityName, Site_ID: 'DIAG-A', Name: 'Entity ' + entityName + ' not in SCHEMA' }];
  }
  var ss = SpreadsheetApp.getActive();
  if (!ss) {
    return [{ _diag: 'NO_ACTIVE_SPREADSHEET', Site_ID: 'DIAG-B', Name: 'SpreadsheetApp.getActive() returned null' }];
  }
  var sh = ss.getSheetByName(entityName);
  if (!sh) {
    return [{ _diag: 'SHEET_MISSING', entityName: entityName, Site_ID: 'DIAG-C', Name: 'No sheet named ' + entityName }];
  }
  var last = sh.getLastRow();
  var cols = sh.getLastColumn();
  if (last < 2) {
    return [{ _diag: 'EMPTY_SHEET', Site_ID: 'DIAG-D', Name: 'Sheet is empty (lastRow=' + last + ')' }];
  }
  var headers = sh.getRange(1, 1, 1, cols).getValues()[0];
  var values  = sh.getRange(2, 1, last - 1, cols).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var o = {};
    for (var j = 0; j < headers.length; j++) {
      o[String(headers[j])] = values[i][j];
    }
    out.push(o);
  }
  if (out.length === 0) {
    return [{ _diag: 'LOOP_PRODUCED_ZERO', Site_ID: 'DIAG-E', Name: 'Loop ran but produced no rows. values.length=' + values.length }];
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
