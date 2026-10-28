/**
 * Masters-System · Config
 * -----------------------------------------------------------------------------
 * Single source of truth for sheet schemas, ID schemes, permissions and
 * seed data. Everything downstream reads from here.
 */

const MS_VERSION = '0.2.2';

// -----------------------------------------------------------------------------
// Permissions — seed values written on first-time setup. After setup the live
// values live in the Config sheet and admin can change them from Admin UI.
// -----------------------------------------------------------------------------
const SEED_ADMIN_EMAILS  = ['modular@3starprojects.com'];
const SEED_EDITOR_EMAILS = [
  'gsp@3starprojects.com',
  'info@3starprojects.com',
  'projects@3starprojects.com',
  'production@3starprojects.com',
  'sales@3starprojects.com'
];
const SEED_DOMAIN = '3starprojects.com';

// -----------------------------------------------------------------------------
// Sheet names
// -----------------------------------------------------------------------------
const SH = {
  CONFIG:    'Config',
  PROJECTS:  'Projects',
  CUSTOMERS: 'Customers',
  PARTIES:   'Parties',
  EMPLOYEES: 'Employees',
  SITES:     'Sites',
  AUDIT:     'Audit_Log'
};

// -----------------------------------------------------------------------------
// Field-type conventions used by the Admin UI's form renderer.
// Types: 'id' (readonly, server-generated), 'text', 'textarea', 'number',
// 'date', 'yesno', 'enum' (needs .values), 'lookup' (needs .from = sheet
// name; shows label "<ID> — <Name>"), 'audit' (readonly, auto-stamped).
// Fields NOT listed in fieldTypes default to plain text, except names ending
// in '_At' (audit datetime), '_By' (audit text) and '?' (yesno).
// -----------------------------------------------------------------------------
const SCHEMA = {

  'Config': {
    headers: ['Key', 'Value', 'Description'],
    freezeCols: 1,
    protect: 'admin',
    hiddenFromUI: true,          // not shown as a CRUD tab
    fieldTypes: {
      'Key':         { type: 'text', required: true, readonly: true },
      'Value':       { type: 'text' },
      'Description': { type: 'text' }
    }
  },

  'Projects': {
    headers: [
      'Project_ID', 'Name', 'Short_Code',
      'Customer_ID', 'Site_ID',
      'Start_Date', 'Target_Handover',
      'Status', 'Owner_Emp_ID',
      'Notes',
      'Active?',
      'Created_At', 'Created_By', 'Updated_At', 'Updated_By'
    ],
    idCol: 'Project_ID',
    idPrefix: 'PRJ',
    idPad: 3,
    counterKey: 'next_prj',
    freezeCols: 2,
    statusValues: ['Enquiry','Confirmed','In Progress','Delivered','Handover','On Hold','Cancelled'],
    protect: 'editor',
    labelCol: 'Name',
    fieldTypes: {
      'Project_ID':      { type: 'id' },
      'Name':            { type: 'text',     required: true },
      'Short_Code':      { type: 'text',     hint: 'optional short handle, e.g. PATEL-01' },
      'Customer_ID':     { type: 'lookup',   from: 'Customers' },
      'Site_ID':         { type: 'lookup',   from: 'Sites' },
      'Start_Date':      { type: 'date' },
      'Target_Handover': { type: 'date' },
      'Status':          { type: 'enum',     values: ['Enquiry','Confirmed','In Progress','Delivered','Handover','On Hold','Cancelled'], default: 'Enquiry' },
      'Owner_Emp_ID':    { type: 'lookup',   from: 'Employees' },
      'Notes':           { type: 'textarea' },
      'Active?':         { type: 'yesno',    default: 'Yes' }
    }
  },

  'Customers': {
    headers: [
      'Customer_ID', 'Name', 'Legal_Name',
      'GSTIN', 'PAN',
      'Billing_Address', 'Shipping_Address', 'State',
      'Contact_Person', 'Phone', 'Email',
      'Payment_Terms_Days', 'Credit_Limit',
      'Active?',
      'Created_At', 'Created_By', 'Updated_At', 'Updated_By'
    ],
    idCol: 'Customer_ID',
    idPrefix: 'CUST',
    idPad: 4,
    counterKey: 'next_cust',
    freezeCols: 2,
    protect: 'editor',
    labelCol: 'Name',
    fieldTypes: {
      'Customer_ID':         { type: 'id' },
      'Name':                { type: 'text',     required: true },
      'Legal_Name':          { type: 'text',     hint: 'legal name for invoicing (if different)' },
      'GSTIN':               { type: 'text' },
      'PAN':                 { type: 'text' },
      'Billing_Address':     { type: 'textarea' },
      'Shipping_Address':    { type: 'textarea' },
      'State':               { type: 'text',     hint: 'for CGST+SGST vs IGST' },
      'Contact_Person':      { type: 'text' },
      'Phone':               { type: 'text' },
      'Email':               { type: 'text' },
      'Payment_Terms_Days':  { type: 'number' },
      'Credit_Limit':        { type: 'number' },
      'Active?':             { type: 'yesno',    default: 'Yes' }
    }
  },

  'Parties': {
    headers: [
      'Party_ID', 'Name', 'Legal_Name',
      'GSTIN', 'PAN',
      'Address', 'State',
      'Contact_Person', 'Phone', 'Email',
      'Payment_Terms_Days',
      'Is_Supplier?', 'Is_JobWorker?', 'Is_Transporter?', 'Is_ServiceProv?',
      'Active?',
      'Created_At', 'Created_By', 'Updated_At', 'Updated_By'
    ],
    idCol: 'Party_ID',
    idPrefix: 'VEND',
    idPad: 4,
    counterKey: 'next_vend',
    freezeCols: 2,
    protect: 'editor',
    labelCol: 'Name',
    fieldTypes: {
      'Party_ID':            { type: 'id' },
      'Name':                { type: 'text',     required: true },
      'Legal_Name':          { type: 'text' },
      'GSTIN':               { type: 'text' },
      'PAN':                 { type: 'text' },
      'Address':             { type: 'textarea' },
      'State':               { type: 'text' },
      'Contact_Person':      { type: 'text' },
      'Phone':               { type: 'text' },
      'Email':               { type: 'text' },
      'Payment_Terms_Days':  { type: 'number' },
      'Is_Supplier?':        { type: 'yesno',   default: 'No' },
      'Is_JobWorker?':       { type: 'yesno',   default: 'No' },
      'Is_Transporter?':     { type: 'yesno',   default: 'No' },
      'Is_ServiceProv?':     { type: 'yesno',   default: 'No' },
      'Active?':             { type: 'yesno',   default: 'Yes' }
    }
  },

  'Employees': {
    headers: [
      'Emp_ID', 'Name', 'Role',
      'Site_ID',
      'Join_Date',
      'Daily_Rate', 'Holds_Float?',
      'Phone',
      'Status', 'Left_Date',
      'Notes',
      'Created_At', 'Created_By', 'Updated_At', 'Updated_By'
    ],
    idCol: 'Emp_ID',
    idPrefix: 'EMP',
    idPad: 4,
    counterKey: 'next_emp',
    freezeCols: 2,
    statusValues: ['Active', 'Left'],
    protect: 'editor',
    labelCol: 'Name',
    fieldTypes: {
      'Emp_ID':         { type: 'id' },
      'Name':           { type: 'text',     required: true },
      'Role':           { type: 'text',     hint: 'e.g. Supervisor, Carpenter, Polisher' },
      'Site_ID':        { type: 'lookup',   from: 'Sites' },
      'Join_Date':      { type: 'date' },
      'Daily_Rate':     { type: 'number' },
      'Holds_Float?':   { type: 'yesno',    default: 'No' },
      'Phone':          { type: 'text' },
      'Status':         { type: 'enum',     values: ['Active', 'Left'], default: 'Active' },
      'Left_Date':      { type: 'date',     hint: 'only if Status=Left' },
      'Notes':          { type: 'textarea' }
    }
  },

  'Sites': {
    headers: [
      'Site_ID', 'Name', 'Type',
      'Address',
      'Active?',
      'Notes',
      'Created_At', 'Created_By', 'Updated_At', 'Updated_By'
    ],
    idCol: 'Site_ID',
    idPrefix: 'SITE',
    idPad: 2,
    counterKey: 'next_site',
    freezeCols: 2,
    typeValues: ['Office', 'Factory', 'Warehouse', 'Showroom'],
    protect: 'editor',
    labelCol: 'Name',
    fieldTypes: {
      'Site_ID':  { type: 'id' },
      'Name':     { type: 'text',     required: true },
      'Type':     { type: 'enum',     values: ['Office','Factory','Warehouse','Showroom'], default: 'Office' },
      'Address':  { type: 'textarea' },
      'Active?':  { type: 'yesno',    default: 'Yes' },
      'Notes':    { type: 'textarea' }
    }
  },

  'Audit_Log': {
    headers: [
      'Timestamp', 'User_Email', 'Entity_Type', 'Entity_ID',
      'Action', 'Field', 'Old_Value', 'New_Value', 'Note'
    ],
    freezeCols: 2,
    protect: 'nobody',
    hiddenFromUI: true           // viewed via the sheet, not CRUD
  }
};

// Audit-field detection — anything matching these endings is auto-filled
// server-side and rendered readonly in the UI.
const AUDIT_FIELDS = ['Created_At', 'Created_By', 'Updated_At', 'Updated_By'];
function isAuditField_(name) { return AUDIT_FIELDS.indexOf(name) >= 0; }

// Entities the Admin UI shows as tabs, in display order.
const UI_ENTITIES = ['Projects', 'Customers', 'Parties', 'Employees', 'Sites'];

// -----------------------------------------------------------------------------
// Config seed — written to Config sheet on first-time init only.
// -----------------------------------------------------------------------------
const CONFIG_SEED = [
  ['library_version',    MS_VERSION,                             'Masters library version'],
  ['domain',             SEED_DOMAIN,                            'Workspace domain for access control'],
  ['admin_emails',       SEED_ADMIN_EMAILS.join(','),            'Full-access accounts (comma-separated)'],
  ['editor_emails',      SEED_EDITOR_EMAILS.join(','),           'Add/edit-registers accounts (comma-separated)'],
  ['next_prj',           '1',                                    'Next Project_ID counter'],
  ['next_cust',          '1',                                    'Next Customer_ID counter'],
  ['next_vend',          '1',                                    'Next Party_ID counter (VEND-####)'],
  ['next_emp',           '1',                                    'Next Emp_ID counter'],
  ['next_site',          '3',                                    'Next Site_ID counter (2 seeded)'],
  ['cache_ttl_seconds',  '300',                                  'Library cache TTL in consumer scripts'],
  ['soft_delete_only',   'true',                                 'Never hard-delete; set Active?=No instead']
];

const SITES_SEED = [
  { id: 'SITE-01', name: 'Thane Office',     type: 'Office',  address: '' },
  { id: 'SITE-02', name: 'Dombivli Factory', type: 'Factory', address: '' }
];
