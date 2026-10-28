/**
 * Masters-System · WebApp
 * -----------------------------------------------------------------------------
 * doGet is the single entry point for the Admin UI. The UI is a tabbed SPA;
 * all data access goes through google.script.run → MastersAPI.*
 */

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'home';
  const t = HtmlService.createTemplateFromFile('Admin');
  t.initialPage = page;
  t.user = {
    email:  getUserEmail(),
    isAdmin:  isAdmin(),
    isEditor: isEditor(),
    version: MS_VERSION
  };
  return t.evaluate()
    .setTitle('3SP · Masters')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Lightweight endpoint the UI calls on first load to confirm permissions +
// get the config block it needs for client-side gating.
function ui_getBootstrap() {
  return {
    user: {
      email: getUserEmail(),
      isAdmin: isAdmin(),
      isEditor: isEditor()
    },
    version: MS_VERSION,
    domain: getConfigValue('domain'),
    entities: {
      projects:  { count: safeCount_(SH.PROJECTS) },
      customers: { count: safeCount_(SH.CUSTOMERS) },
      parties:   { count: safeCount_(SH.PARTIES) },
      employees: { count: safeCount_(SH.EMPLOYEES) },
      sites:     { count: safeCount_(SH.SITES) }
    }
  };
}

function safeCount_(sheetName) {
  const sh = ss_().getSheetByName(sheetName);
  if (!sh) return 0;
  return Math.max(0, sh.getLastRow() - 1);
}
