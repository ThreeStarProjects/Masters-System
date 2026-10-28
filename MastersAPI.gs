/**
 * Masters-System · Public API
 * -----------------------------------------------------------------------------
 * These functions are what the three consumer systems (Finance, IMS, PM)
 * call via the Masters Apps Script library. Keep signatures stable — every
 * breaking change here is a release-coordinated change across four repos.
 *
 * v0.1 — read-only stubs. Writers come in v0.2 (per-entity CRUD).
 * -----------------------------------------------------------------------------
 */

// -----------------------------------------------------------------------------
// Projects
// -----------------------------------------------------------------------------
function getProjects(filter) {
  const all = readAll_(SH.PROJECTS);
  return all.filter(r => matchesFilter_(r, filter)).map(stripInternal_);
}

function getProject(projectId) {
  if (!projectId) return null;
  const all = readAll_(SH.PROJECTS);
  const r = all.find(x => String(x.Project_ID).trim() === String(projectId).trim());
  return r ? stripInternal_(r) : null;
}

// -----------------------------------------------------------------------------
// Customers
// -----------------------------------------------------------------------------
function getCustomers(filter) {
  return readAll_(SH.CUSTOMERS).filter(r => matchesFilter_(r, filter)).map(stripInternal_);
}

function getCustomer(customerId) {
  if (!customerId) return null;
  const r = readAll_(SH.CUSTOMERS).find(x => String(x.Customer_ID).trim() === String(customerId).trim());
  return r ? stripInternal_(r) : null;
}

// -----------------------------------------------------------------------------
// Parties (Suppliers / Job Workers / Transporters / Service Providers)
// -----------------------------------------------------------------------------
function getParties(filter) {
  // filter may include {isSupplier:true}, {isJobWorker:true}, etc.
  // Those get converted to column-name filters below.
  const f = Object.assign({}, filter || {});
  if (f.isSupplier)    { f['Is_Supplier?']    = 'Yes'; delete f.isSupplier; }
  if (f.isJobWorker)   { f['Is_JobWorker?']   = 'Yes'; delete f.isJobWorker; }
  if (f.isTransporter) { f['Is_Transporter?'] = 'Yes'; delete f.isTransporter; }
  if (f.isServiceProv) { f['Is_ServiceProv?'] = 'Yes'; delete f.isServiceProv; }
  return readAll_(SH.PARTIES).filter(r => matchesFilter_(r, f)).map(stripInternal_);
}

function getParty(partyId) {
  if (!partyId) return null;
  const r = readAll_(SH.PARTIES).find(x => String(x.Party_ID).trim() === String(partyId).trim());
  return r ? stripInternal_(r) : null;
}

// -----------------------------------------------------------------------------
// Employees
// -----------------------------------------------------------------------------
function getEmployees(filter) {
  return readAll_(SH.EMPLOYEES).filter(r => matchesFilter_(r, filter)).map(stripInternal_);
}

function getEmployee(empId) {
  if (!empId) return null;
  const r = readAll_(SH.EMPLOYEES).find(x => String(x.Emp_ID).trim() === String(empId).trim());
  return r ? stripInternal_(r) : null;
}

// -----------------------------------------------------------------------------
// Sites
// -----------------------------------------------------------------------------
function getSites(filter) {
  return readAll_(SH.SITES).filter(r => matchesFilter_(r, filter)).map(stripInternal_);
}

function getSite(siteId) {
  if (!siteId) return null;
  const r = readAll_(SH.SITES).find(x => String(x.Site_ID).trim() === String(siteId).trim());
  return r ? stripInternal_(r) : null;
}

// -----------------------------------------------------------------------------
// Freeform-name resolver — maps legacy project/customer/party/employee strings
// in Finance/IMS/PM to Masters IDs during migration. Returns best match + score.
// -----------------------------------------------------------------------------
function resolveFreeformName(entity, text) {
  if (!text) return null;
  const map = {
    'project':  { sheet: SH.PROJECTS,  idCol: 'Project_ID',  nameCol: 'Name' },
    'customer': { sheet: SH.CUSTOMERS, idCol: 'Customer_ID', nameCol: 'Name' },
    'party':    { sheet: SH.PARTIES,   idCol: 'Party_ID',    nameCol: 'Name' },
    'employee': { sheet: SH.EMPLOYEES, idCol: 'Emp_ID',      nameCol: 'Name' }
  };
  const spec = map[String(entity).toLowerCase()];
  if (!spec) return null;
  const all = readAll_(spec.sheet);
  const exact = all.find(r => normName(r[spec.nameCol]) === normName(text));
  if (exact) return { id: exact[spec.idCol], name: exact[spec.nameCol], confidence: 'exact' };
  const prefix = all.find(r => normMatch(r[spec.nameCol], text));
  if (prefix) return { id: prefix[spec.idCol], name: prefix[spec.nameCol], confidence: 'fuzzy' };
  return null;
}

// -----------------------------------------------------------------------------
// Meta — version + bulk project list for dropdowns
// -----------------------------------------------------------------------------
function getVersion() { return MS_VERSION; }

function getProjectLabels() {
  // Shape expected by data-validation list builders:
  //   [{ id: 'PRJ-012', label: 'PRJ-012 — Thane Residence' }, ...]
  return readAll_(SH.PROJECTS)
    .filter(r => String(r['Active?']).toLowerCase() !== 'no')
    .map(r => ({
      id: r.Project_ID,
      label: r.Project_ID + ' — ' + r.Name
    }));
}

// -----------------------------------------------------------------------------
// Internals
// -----------------------------------------------------------------------------

function matchesFilter_(row, filter) {
  if (!filter) return true;
  return Object.keys(filter).every(k => {
    const want = String(filter[k]).toLowerCase();
    const have = String(row[k] == null ? '' : row[k]).toLowerCase();
    return have === want;
  });
}

function stripInternal_(row) {
  // Drop _row (internal) before handing back to a consumer.
  const o = Object.assign({}, row);
  delete o._row;
  return o;
}
