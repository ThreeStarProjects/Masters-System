/**
 * Masters-System · Public Library API
 * -----------------------------------------------------------------------------
 * The surface Finance / IMS / Project-Management import via Apps Script library.
 * Keep names stable — every breaking change here is a coordinated release
 * across four repos.
 *
 * Internally this is a thin wrapper over CRUD.gs. The Admin UI calls the
 * CRUD module directly; consumers call these library-friendly wrappers.
 */

// -----------------------------------------------------------------------------
// Projects
// -----------------------------------------------------------------------------
function getProjects(filter)        { return entity_list('Projects', { filter: filter }); }
function getProject(projectId)      { return entity_get('Projects', projectId); }
function createProject(data)        { return entity_save('Projects', data); }
function updateProject(id, data)    { return entity_save('Projects', Object.assign({}, data, { 'Project_ID': id })); }
function deleteProject(id)          { return entity_softDelete('Projects', id); }

// -----------------------------------------------------------------------------
// Customers
// -----------------------------------------------------------------------------
function getCustomers(filter)       { return entity_list('Customers', { filter: filter }); }
function getCustomer(customerId)    { return entity_get('Customers', customerId); }
function createCustomer(data)       { return entity_save('Customers', data); }
function updateCustomer(id, data)   { return entity_save('Customers', Object.assign({}, data, { 'Customer_ID': id })); }
function deleteCustomer(id)         { return entity_softDelete('Customers', id); }

/**
 * FINANCE IS THE MASTER FOR CUSTOMERS AND VENDORS; Masters holds a copy for IMS / PM. Finance calls this with every customer (or vendor) it has:
 *   list = [{ financeId:'CUST-0015', name, gstin?, email?, phone?, address?, active? (false = inactive) }]
 * Each Finance row is matched to a Masters row by (1) Finance_ID, (2) the approved Rec_ row (Source_ID -> Target_Master_ID),
 * (3) the same name (case / punctuation / spaces ignored, only if that Masters row is not yet linked). A match gets the Finance_ID
 * stored (bulk, no audit spam) and any CHANGED Finance value copied over (name, GSTIN, email, phone, address, active); a blank
 * Finance value never blanks Masters. No match -> a new Masters row. Writes are capped at `max` (default 20) per call; { more:true } = call again.
 * Editors only. -> { total, linked, created, updated, unchanged, more, mastersOnly:[names] }   (mastersOnly: customers only)
 */
function syncFromFinance_(cfg, list, max) {
  requireEditor_();
  const key = function (s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); };
  const t = function (v) { return String(v == null ? '' : v).trim(); };
  const cap = max > 0 ? max : 20;
  const sh = sh_(cfg.entity), idCol = SCHEMA[cfg.entity].idCol;
  ensureColumns_(sh, SCHEMA[cfg.entity].headers);
  const rows = entity_list(cfg.entity, {});
  const byFin = {}, byId = {}, byName = {};
  rows.forEach(function (r) {
    byId[r[idCol]] = r;
    if (t(r.Finance_ID)) byFin[t(r.Finance_ID)] = r;
  });
  rows.forEach(function (r) { const k = key(r.Name); if (k && !t(r.Finance_ID) && !byName[k]) byName[k] = r; });
  const rec = {};
  try {
    const rs = ss_().getSheetByName(cfg.rec);
    if (rs && rs.getLastRow() >= 2) {
      const h = rs.getRange(1, 1, 1, rs.getLastColumn()).getValues()[0], iS = h.indexOf('Source'), iI = h.indexOf('Source_ID'), iT = h.indexOf('Target_Master_ID');
      if (iS >= 0 && iI >= 0 && iT >= 0) rs.getRange(2, 1, rs.getLastRow() - 1, h.length).getValues().forEach(function (r) {
        if (t(r[iS]) === 'Finance' && t(r[iI]) && t(r[iT])) rec[t(r[iI])] = t(r[iT]);
      });
    }
  } catch (e) {}
  const res = { total: 0, linked: 0, created: 0, updated: 0, unchanged: 0, more: false, mastersOnly: [] };
  const link = {}, writes = [], claimed = {};
  (list || []).forEach(function (f) {
    const fid = t(f && f.financeId), name = t(f && f.name).replace(/\s+/g, ' ');
    if (!fid || !name) return;
    res.total++;
    let m = byFin[fid];
    if (!m && rec[fid]) { const c = byId[rec[fid]]; if (c && (!t(c.Finance_ID) || t(c.Finance_ID) === fid) && !claimed[c[idCol]]) m = c; }
    if (!m) { const c = byName[key(name)]; if (c && !t(c.Finance_ID) && !claimed[c[idCol]]) m = c; }
    if (!m) {
      const nd = Object.assign({ Name: name, GSTIN: t(f.gstin), Email: t(f.email), Phone: t(f.phone), 'Active?': f.active === false ? 'No' : 'Yes', Finance_ID: fid }, cfg.extra);
      nd[cfg.addr] = t(f.address);
      writes.push({ fid: fid, data: nd }); return;
    }
    claimed[m[idCol]] = true;
    if (t(m.Finance_ID) !== fid) { link[m[idCol]] = fid; res.linked++; }
    const d = {}; d[idCol] = m[idCol]; let any = false;
    const set = function (field, val) { if (val !== '' && t(m[field]) !== val) { d[field] = val; any = true; } };
    set('Name', name); set('GSTIN', t(f.gstin)); set('Email', t(f.email)); set('Phone', t(f.phone)); set(cfg.addr, t(f.address));
    const act = f.active === false ? 'No' : 'Yes';
    if (String(m['Active?'] || 'Yes').toLowerCase() !== act.toLowerCase()) { d['Active?'] = act; any = true; }
    if (any) writes.push({ fid: fid, data: d }); else if (!link[m[idCol]]) res.unchanged++;
  });
  // 1. Finance_ID links in one bulk write
  const lids = Object.keys(link);
  if (lids.length) {
    withLock_(function () {
      const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], cId = hdr.indexOf(idCol), cF = hdr.indexOf('Finance_ID');
      const n = sh.getLastRow() - 1;
      const ids = sh.getRange(2, cId + 1, n, 1).getValues(), fin = sh.getRange(2, cF + 1, n, 1).getValues();
      for (let i = 0; i < n; i++) if (link[ids[i][0]]) fin[i][0] = link[ids[i][0]];
      sh.getRange(2, cF + 1, n, 1).setValues(fin);
      audit_(cfg.entity, '*', 'Link', 'Finance_ID', '', lids.length + ' row(s) linked to Finance codes', 'Finance sync');
    });
  }
  // 2. changed / new rows, a few per call
  writes.slice(0, cap).forEach(function (w) {
    entity_save(cfg.entity, w.data);
    if (w.data[idCol]) res.updated++; else res.created++;
  });
  res.more = writes.length > cap;
  if (cfg.mastersOnly) {
    rows.forEach(function (r) { if (!t(r.Finance_ID) && !claimed[r[idCol]] && String(r['Active?'] || 'Yes').toLowerCase() !== 'no') res.mastersOnly.push(r.Name); });
    res.mastersOnly = res.mastersOnly.filter(function (n) { return !(list || []).some(function (f) { return key(f && f.name) === key(n); }); });
  }
  return res;
}
function syncCustomersFromFinance(list, max) { return syncFromFinance_({ entity: 'Customers', rec: 'Rec_Customers', addr: 'Billing_Address', extra: {}, mastersOnly: true }, list, max); }
/** Finance vendors -> Masters Parties (new ones are suppliers). Same rules as customers; a vendor only in Masters is normal (IMS suppliers, job workers). */
function syncVendorsFromFinance(list, max) { return syncFromFinance_({ entity: 'Parties', rec: 'Rec_Parties', addr: 'Address', extra: { 'Is_Supplier?': 'Yes' }, mastersOnly: false }, list, max); }

/**
 * IMS SUPPLIERS / JOB WORKERS: Masters is the master. IMS sends the codes it already uses once, so every Masters party knows its IMS code
 * (stored in IMS_Supplier_Code / IMS_JW_Code) and IMS's old SUPP-xxxx / JW-xxxx codes stay valid in POs, GRNs and job work challans.
 *   kind = 'supplier' | 'jobworker';  list = [{ code, name, category?, contact?, phone?, email?, gstin?, address?, active? }]
 * Each IMS row is matched to a Masters party by (1) the stored code, (2) the approved Rec_Parties row (Source IMS-Supplier / IMS-JobWorker),
 * (3) the same name (case / punctuation / spaces ignored, only a party without that code yet). A match gets the code stored (and the role
 * flag set to Yes if missing), and any detail Masters has BLANK (category, contact, phone, email, GSTIN, address) is filled from IMS; a value
 * Masters already has is never changed. No match -> a new Masters party (at most `max` per call, default 20;
 * { more:true } = call again). Editors only. -> { total, linked, flagged, created, more }
 */
function syncPartiesFromIms(kind, list, max) {
  requireEditor_();
  const key = function (s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); };
  const t = function (v) { return String(v == null ? '' : v).trim(); };
  const jw = kind === 'jobworker', codeCol = jw ? 'IMS_JW_Code' : 'IMS_Supplier_Code', flag = jw ? 'Is_JobWorker?' : 'Is_Supplier?', recSrc = jw ? 'IMS-JobWorker' : 'IMS-Supplier';
  const cap = max > 0 ? max : 20;
  const sh = sh_('Parties');
  ensureColumns_(sh, SCHEMA['Parties'].headers);
  const rows = entity_list('Parties', {});
  const byCode = {}, byId = {}, byName = {};
  rows.forEach(function (r) { byId[r.Party_ID] = r; if (t(r[codeCol])) byCode[t(r[codeCol])] = r; });
  rows.forEach(function (r) { const k = key(r.Name); if (k && !t(r[codeCol]) && !byName[k]) byName[k] = r; });
  const rec = {};
  try {
    const rs = ss_().getSheetByName('Rec_Parties');
    if (rs && rs.getLastRow() >= 2) {
      const h = rs.getRange(1, 1, 1, rs.getLastColumn()).getValues()[0], iS = h.indexOf('Source'), iI = h.indexOf('Source_ID'), iT = h.indexOf('Target_Master_ID');
      if (iS >= 0 && iI >= 0 && iT >= 0) rs.getRange(2, 1, rs.getLastRow() - 1, h.length).getValues().forEach(function (r) {
        if (t(r[iS]) === recSrc && t(r[iI]) && t(r[iT])) rec[t(r[iI])] = t(r[iT]);
      });
    }
  } catch (e) {}
  const res = { total: 0, linked: 0, flagged: 0, created: 0, more: false };
  const link = {}, setFlag = {}, fill = {}, creates = [], claimed = {};
  (list || []).forEach(function (f) {
    const code = t(f && f.code), name = t(f && f.name).replace(/\s+/g, ' ');
    if (!code || !name) return;
    res.total++;
    let m = byCode[code];
    if (!m && rec[code]) { const c = byId[rec[code]]; if (c && (!t(c[codeCol]) || t(c[codeCol]) === code) && !claimed[c.Party_ID]) m = c; }
    if (!m) { const c = byName[key(name)]; if (c && !t(c[codeCol]) && !claimed[c.Party_ID]) m = c; }
    if (!m) {
      const nd = { Name: name, Category: t(f.category), Contact_Person: t(f.contact), Phone: t(f.phone), Email: t(f.email), GSTIN: t(f.gstin), Address: t(f.address), 'Active?': f.active === false ? 'No' : 'Yes' };
      nd[flag] = 'Yes'; nd[codeCol] = code;
      creates.push(nd); return;
    }
    claimed[m.Party_ID] = true;
    if (t(m[codeCol]) !== code) { link[m.Party_ID] = code; res.linked++; }
    if (String(m[flag]).toLowerCase() !== 'yes') { setFlag[m.Party_ID] = true; res.flagged++; }
    // details Masters lacks are taken from IMS (blank fields only; a value already in Masters is never changed)
    [['Category', f.category], ['Contact_Person', f.contact], ['Phone', f.phone], ['Email', f.email], ['GSTIN', f.gstin], ['Address', f.address]].forEach(function (p) {
      if (t(p[1]) && !t(m[p[0]])) { (fill[m.Party_ID] = fill[m.Party_ID] || {})[p[0]] = t(p[1]); }
    });
  });
  const patch = {};   // Party_ID -> { column: value }
  Object.keys(link).forEach(function (i) { (patch[i] = patch[i] || {})[codeCol] = link[i]; });
  Object.keys(setFlag).forEach(function (i) { (patch[i] = patch[i] || {})[flag] = 'Yes'; });
  Object.keys(fill).forEach(function (i) { patch[i] = Object.assign(patch[i] || {}, fill[i]); });
  const pids = Object.keys(patch);
  if (pids.length) {
    withLock_(function () {
      const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], n = sh.getLastRow() - 1;
      const idv = sh.getRange(2, hdr.indexOf('Party_ID') + 1, n, 1).getValues();
      const cols = {}; pids.forEach(function (i) { Object.keys(patch[i]).forEach(function (c) { cols[c] = true; }); });
      Object.keys(cols).forEach(function (c) {   // one read + one write per touched column
        const ci = hdr.indexOf(c) + 1, vals = sh.getRange(2, ci, n, 1).getValues();
        for (let i = 0; i < n; i++) { const pp = patch[idv[i][0]]; if (pp && pp[c] !== undefined) vals[i][0] = pp[c]; }
        sh.getRange(2, ci, n, 1).setValues(vals);
      });
      audit_('Parties', '*', 'Link', codeCol, '', Object.keys(link).length + ' party(ies) linked to IMS codes, ' + Object.keys(setFlag).length + ' role flag(s) set, ' + Object.keys(fill).length + ' filled from IMS', 'IMS sync');
    });
  }
  creates.slice(0, cap).forEach(function (d) { entity_save('Parties', d); res.created++; });
  res.more = creates.length > cap;
  return res;
}

/**
 * Add a supplier / job worker from IMS (or anywhere). data = { Name, role:'supplier'|'jobworker', code? (IMS code), Category?, Contact_Person?, Phone?, Email?, GSTIN?, Address? }.
 * The same name already in Masters is never added twice: it either already has the role -> { ok:false, duplicate:{id,name} }, or it is given the role (and the
 * IMS code) -> { ok:true, id, name, existing:true }. Otherwise a new party is created. Editors only.
 */
function addPartyChecked(data) {
  requireEditor_();
  const t = function (k) { return String((data && data[k]) || '').trim(); };
  const name = t('Name').replace(/\s+/g, ' ');
  if (!name) throw new Error('Name is required.');
  if (name.length > 120) throw new Error('Name is too long (max 120 characters).');
  const k = function (s) { return String(s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); };
  if (!k(name)) throw new Error('Name must contain letters or numbers.');
  const jw = data && data.role === 'jobworker', flag = jw ? 'Is_JobWorker?' : 'Is_Supplier?', codeCol = jw ? 'IMS_JW_Code' : 'IMS_Supplier_Code';
  const same = entity_list('Parties', {}).filter(function (r) { return k(r.Name) === k(name); })[0];
  if (same) {
    if (String(same[flag]).toLowerCase() === 'yes') return { ok: false, duplicate: { id: same.Party_ID, name: same.Name } };
    const u = { Party_ID: same.Party_ID }; u[flag] = 'Yes'; if (t('code') && !String(same[codeCol] || '').trim()) u[codeCol] = t('code');
    entity_save('Parties', u);
    return { ok: true, id: same.Party_ID, name: same.Name, existing: true };
  }
  const row = { Name: name, Category: t('Category'), Contact_Person: t('Contact_Person'), Phone: t('Phone'), Email: t('Email'), GSTIN: t('GSTIN'), Address: t('Address'), 'Active?': 'Yes' };
  row[flag] = 'Yes'; if (t('code')) row[codeCol] = t('code');
  const saved = entity_save('Parties', row);
  return { ok: true, id: saved.Party_ID, name: saved.Name };
}

// -----------------------------------------------------------------------------
// Parties (unified Supplier / JobWorker / Transporter / ServiceProv)
// -----------------------------------------------------------------------------
function getParties(filter) {
  const f = Object.assign({}, filter || {});
  if (f.isSupplier)    { f['Is_Supplier?']    = 'Yes'; delete f.isSupplier; }
  if (f.isJobWorker)   { f['Is_JobWorker?']   = 'Yes'; delete f.isJobWorker; }
  if (f.isTransporter) { f['Is_Transporter?'] = 'Yes'; delete f.isTransporter; }
  if (f.isServiceProv) { f['Is_ServiceProv?'] = 'Yes'; delete f.isServiceProv; }
  return entity_list('Parties', { filter: f });
}
function getParty(partyId)          { return entity_get('Parties', partyId); }
function createParty(data)          { return entity_save('Parties', data); }
function updateParty(id, data)      { return entity_save('Parties', Object.assign({}, data, { 'Party_ID': id })); }
function deleteParty(id)            { return entity_softDelete('Parties', id); }

// -----------------------------------------------------------------------------
// Employees
// -----------------------------------------------------------------------------
function getEmployees(filter)       { return entity_list('Employees', { filter: filter }); }
function getEmployee(empId)         { return entity_get('Employees', empId); }
function createEmployee(data)       { return entity_save('Employees', data); }
function updateEmployee(id, data)   { return entity_save('Employees', Object.assign({}, data, { 'Emp_ID': id })); }
function deleteEmployee(id)         { return entity_softDelete('Employees', id); }

/**
 * Fill ONLY the blank fields of existing employees (matched by name; case / punctuation / spaces ignored). Used to move details that
 * were typed in PM or Finance into Masters once. A value that is already in Masters is NEVER overwritten. Editors only.
 * updates: [{ Name, Role?, Department?, Phone?, Email?, Gross_Monthly_Salary?, Daily_Rate? }]  →  { filled: n, notFound: [names] }
 */
function fillEmployeeBlanks(updates) {
  requireEditor_();
  const norm = function (s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); };
  const byKey = {};
  entity_list('Employees', {}).forEach(function (e) { const k = norm(e.Name); if (k && !byKey[k]) byKey[k] = e; });
  let filled = 0; const notFound = [];
  (updates || []).forEach(function (u) {
    const e = byKey[norm(u && u.Name)];
    if (!e) { notFound.push(u && u.Name); return; }
    const data = { Emp_ID: e.Emp_ID }; let any = false;
    ['Role', 'Department', 'Phone', 'Email', 'Gross_Monthly_Salary', 'Daily_Rate'].forEach(function (f) {
      if (u[f] === undefined || u[f] === null || String(u[f]).trim() === '') return;
      if (String(e[f] === undefined || e[f] === null ? '' : e[f]).trim() !== '') return;   // Masters already has it
      data[f] = u[f]; any = true;
    });
    if (any) { entity_save('Employees', data); filled++; }
  });
  return { filled: filled, notFound: notFound };
}

/** Put the live Employees sheet in the schema's column order (by header name, data follows, hidden backup first). Idempotent. Editors only. */
function ensureEmployeeLayout() {
  requireEditor_();
  ensureColumns_(sh_('Employees'), SCHEMA['Employees'].headers);
  return true;
}

/**
 * Safe "add employee" for the Add-employee buttons in PM / IMS / Finance. Editors only (like every Masters write).
 * data: { Name, Role? (= designation), Phone?, Department?, Email?, Gross_Monthly_Salary?, Daily_Rate?, 'Holds_Float?'?, confirmSimilar? }
 *   exact same name already there (case / punctuation / spaces ignored) → { ok:false, duplicate:{id,name,status} }, nothing created
 *   a similar name (all words of one are inside the other: "Arjun" / "Arjun Singh") → { ok:false, similar:[{id,name,status}] }
 *     unless confirmSimilar === true (the person looked and wants a new one anyway)
 *   otherwise creates an Active employee → { ok:true, id, name }
 */
function addEmployeeChecked(data) {
  requireEditor_();
  const name = String((data && data.Name) || '').replace(/\s+/g, ' ').trim();
  if (!name) throw new Error('Name is required.');
  if (name.length > 80) throw new Error('Name is too long (max 80 characters).');
  const words = function (s) { return String(s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim().split(' ').filter(String); };
  const mine = words(name), key = mine.join(' ');
  if (!key) throw new Error('Name must contain letters or numbers.');
  const similar = [];
  for (const e of entity_list('Employees', {})) {
    const theirs = words(e.Name), tk = theirs.join(' ');
    if (!tk) continue;
    const info = { id: e.Emp_ID, name: e.Name, status: e.Status };
    if (tk === key) return { ok: false, duplicate: info };
    const short = mine.length <= theirs.length ? mine : theirs, long = mine.length <= theirs.length ? theirs : mine;
    if (short.every(function (w) { return long.indexOf(w) >= 0; })) similar.push(info);
  }
  if (similar.length && !(data && data.confirmSimilar === true)) return { ok: false, similar: similar.slice(0, 5) };
  const t = function (k) { return String((data && data[k]) || '').trim(); };
  const num = function (k) { const v = data && data[k]; return v === undefined || v === null || v === '' ? '' : Number(v); };
  const row = { Name: name, Role: t('Role'), Phone: t('Phone'), Department: t('Department'), Email: t('Email'), Status: 'Active' };
  const gross = num('Gross_Monthly_Salary'), daily = num('Daily_Rate');
  if ((gross !== '' && !(gross >= 0)) || (daily !== '' && !(daily >= 0))) throw new Error('Salary and daily rate must be numbers, 0 or more.');
  if (gross !== '') row.Gross_Monthly_Salary = gross;
  if (daily !== '') row.Daily_Rate = daily;
  if (data && data['Holds_Float?'] === 'Yes') row['Holds_Float?'] = 'Yes';
  const saved = entity_save('Employees', row);
  return { ok: true, id: saved.Emp_ID, name: saved.Name };
}

// -----------------------------------------------------------------------------
// Sites
// -----------------------------------------------------------------------------
function getSites(filter)           { return entity_list('Sites', { filter: filter }); }
function getSite(siteId)            { return entity_get('Sites', siteId); }
function createSite(data)           { return entity_save('Sites', data); }
function updateSite(id, data)       { return entity_save('Sites', Object.assign({}, data, { 'Site_ID': id })); }
function deleteSite(id)             { return entity_softDelete('Sites', id); }

// -----------------------------------------------------------------------------
// Freeform-name resolver — maps legacy project/customer/party/employee text
// to Masters IDs. Used by consumer-side migration jobs.
// -----------------------------------------------------------------------------
function resolveFreeformName(entity, text) {
  if (!text) return null;
  const map = {
    'project':  { sheet: 'Projects',  idCol: 'Project_ID',  nameCol: 'Name' },
    'customer': { sheet: 'Customers', idCol: 'Customer_ID', nameCol: 'Name' },
    'party':    { sheet: 'Parties',   idCol: 'Party_ID',    nameCol: 'Name' },
    'employee': { sheet: 'Employees', idCol: 'Emp_ID',      nameCol: 'Name' }
  };
  const spec = map[String(entity).toLowerCase()];
  if (!spec) return null;
  const all = readAll_(spec.sheet);
  const exact = all.find(function (r) { return normName(r[spec.nameCol]) === normName(text); });
  if (exact) return { id: exact[spec.idCol], name: exact[spec.nameCol], confidence: 'exact' };
  const prefix = all.find(function (r) { return normMatch(r[spec.nameCol], text); });
  if (prefix) return { id: prefix[spec.idCol], name: prefix[spec.nameCol], confidence: 'fuzzy' };
  return null;
}

// -----------------------------------------------------------------------------
// Meta
// -----------------------------------------------------------------------------
function getVersion() { return MS_VERSION; }

function getProjectLabels() {
  return readAll_('Projects')
    .filter(function (r) { return String(r['Active?']).toLowerCase() !== 'no'; })
    .map(function (r) {
      return { id: r.Project_ID, label: r.Project_ID + ' — ' + r.Name };
    });
}
