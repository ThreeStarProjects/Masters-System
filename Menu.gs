/**
 * Masters-System · Menu
 * -----------------------------------------------------------------------------
 * Adds the ⚙ Masters menu to the bound spreadsheet. Admin-only menu items
 * are visible to everyone but check permissions on execution.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙ Masters')
    .addItem('Open Admin UI (web app)', 'menu_openAdminUI_')
    .addSeparator()
    .addSubMenu(ui.createMenu('Setup')
      .addItem('First-time setup',  'firstTimeInit')
      .addItem('Rebuild headers',   'rebuildHeaders')
      .addItem('Rebuild dropdowns', 'rebuildDropdowns')
      .addItem('Diagnose',          'diagnose'))
    .addSubMenu(ui.createMenu('Reconcile (v0.3)')
      .addItem('Pull · Projects',   'reconcile_pullProjects')
      .addItem('Pull · Customers',  'reconcile_pullCustomers')
      .addItem('Pull · Parties',    'reconcile_pullParties')
      .addItem('Pull · Employees',  'reconcile_pullEmployees')
      .addSeparator()
      .addItem('Pull · ALL',        'menu_reconcilePullAll_')
      .addSeparator()
      .addItem('Import approved · Projects',  'reconcile_importProjects')
      .addItem('Import approved · Customers', 'reconcile_importCustomers')
      .addItem('Import approved · Parties',   'reconcile_importParties')
      .addItem('Import approved · Employees', 'reconcile_importEmployees')
      .addSeparator()
      .addItem('⚠ Reset ALL migration data', 'reconcile_resetMigration'))
    .addSeparator()
    .addItem('About', 'menu_about_')
    .addToUi();
}

function menu_openAdminUI_() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert(
      'Admin UI not deployed yet',
      'Deploy this script as a Web App first:\n' +
      '  Deploy > New deployment > Web app\n' +
      '  Execute as: Me\n' +
      '  Who has access: Anyone within ' + SEED_DOMAIN,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  const html = HtmlService.createHtmlOutput(
    '<p>Admin UI deployed at:</p>' +
    '<p><a href="' + url + '" target="_blank">' + url + '</a></p>' +
    '<p><small>Opens in a new tab.</small></p>'
  ).setWidth(500).setHeight(140);
  SpreadsheetApp.getUi().showModalDialog(html, 'Masters · Admin UI');
}

function menu_about_() {
  SpreadsheetApp.getUi().alert(
    'Masters-System v' + MS_VERSION,
    'Central master data for 3 Star Projects.\n\n' +
    'Repo: github.com/ThreeStarProjects/Masters-System\n' +
    'Admin: ' + SEED_ADMIN_EMAILS[0] + '\n' +
    'Editors: ' + SEED_EDITOR_EMAILS.length + ' accounts',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
