/**
 * Masters-System · Config
 * -----------------------------------------------------------------------------
 * Single source of truth for sheet schemas, ID schemes, permissions and
 * seed data. Everything downstream reads from here.
 *
 * Design notes (ponytail: edit this file, not scattered constants).
 */

const MS_VERSION = '0.1.0';

// -----------------------------------------------------------------------------
// Permissions — mirrors what the owner authorized on 2026-10-27.
// Changing these in-sheet via Admin UI persists to the Config tab;
// values here are only used on first-time init.
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
// Sheet names — if you rename any of these, the whole system breaks. Don't.
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
// Schemas — the single declaration of every column in every sheet.
// Order IS the column order. The Setup script writes exactly these headers.
// -----------------------------------------------------------------------------
const SCHEMA = {

  'Config': {
    headers: ['Key', 'Value', 'Description'],
    freezeCols: 1,
    protect: 'admin'            // only admins can edit
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
    protect: 'editor'
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
    protect: 'editor'
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
    idPrefix: 'VEND',           // continuity with Finance
    idPad: 4,
    counterKey: 'next_vend',
    freezeCols: 2,
    protect: 'editor'
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
    protect: 'editor'
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
    protect: 'editor'
  },

  'Audit_Log': {
    headers: [
      'Timestamp', 'User_Email', 'Entity_Type', 'Entity_ID',
      'Action', 'Field', 'Old_Value', 'New_Value', 'Note'
    ],
    freezeCols: 2,
    protect: 'nobody'           // append-only, no one can edit after the fact
  }
};

// -----------------------------------------------------------------------------
// Config seed — what Setup writes to the Config sheet on first-time init.
// Values are strings; booleans and numbers get stringified. Admin can edit
// most of these from the Admin UI once the system is up.
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

// -----------------------------------------------------------------------------
// Site seed — the two permanent sites you confirmed. Addresses blank for
// admin to fill via Admin UI.
// -----------------------------------------------------------------------------
const SITES_SEED = [
  { id: 'SITE-01', name: 'Thane Office',     type: 'Office',  address: '' },
  { id: 'SITE-02', name: 'Dombivli Factory', type: 'Factory', address: '' }
];
