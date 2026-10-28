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
