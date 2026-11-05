/**
 * Masters-System · ProjectDelete
 * -----------------------------------------------------------------------------
 * Delete one project everywhere, in one confirmed step, from PM's menu.
 *
 *   Masters.previewProjectDelete(projectId, pmSheetId, pmLabel) → counts per system, nothing changes
 *   Masters.executeProjectDelete(projectId, pmSheetId, pmLabel, typedId) → backs up, then deletes
 *   Masters.registerSystemSheet('finance'|'ims', spreadsheetId)  (each system calls it when opened)
 *
 * Runs here but works on the OTHER spreadsheets through openById (never getActive: in a library
 * that is the caller's sheet). What counts as "the project's rows":
 *   PM       exact PM rows (Projects, Milestones, Hardware List, Purchases, Production, Snags, Tickets)
 *   IMS      Reservation Register, Material Requisition Register, FG Delivery Register (+ its lines)
 *            whose project text contains the PRJ-### id OR is a text paired to this project in Project_Map
 *   Finance  Project_Costs, Bills, Bill_Payments, Bill_Allocations, Invoices, Payments and every
 *            Transactions_<FY> sheet whose Project is the PRJ-### label OR a Finance project this
 *            project is paired with in Project_Map
 * Every removed row is first copied (as JSON) to a hidden "_Deleted_Project_Rows" sheet in the
 * system it came from. PM goes LAST and the Masters record only when all systems succeeded, so a
 * partial failure can simply be run again. The IMS stock snapshots are refreshed by IMS itself
 * (marker sheet "_PD_Refresh", picked up by its 1-minute trigger).
 */

const PD_BACKUP = '_Deleted_Project_Rows';
const PD_LOG = 'Project_Deletes';
const PD_REFRESH = '_PD_Refresh';
const PD_IMS_DEFAULT_ID = '1CglvPxUWZnIHNeToGEDz8-Xj3eNfvaF351j_ogpRTgk';

// c = fixed 1-based project column, h = header name to look up. lines = child rows keyed by a column of the parent.
const PD_TARGETS = {
  pm: [
    { s: 'Projects', c: 1 }, { s: 'Milestones', c: 1 }, { s: 'Hardware List', c: 2 }, { s: 'Purchases', c: 2 },
    { s: 'Production', c: 2 }, { s: 'Snags', c: 2 }, { s: 'Tickets', c: 11 }
  ],
  ims: [
    { s: 'Reservation Register', h: 'For (Reference)', codeH: 'Item/SFG Code' },
    { s: 'Material Requisition Register', h: 'Project', codeH: 'Item Code' },
    { s: 'FG Delivery Register', h: 'Project', keyH: 'DC No', lines: { s: 'FG Delivery Lines', h: 'DC No' } }
  ],
  fin: [
    { s: 'Project_Costs', h: 'Project' }, { s: 'Bills', h: 'Project' }, { s: 'Bill_Payments', h: 'Project' },
    { s: 'Bill_Allocations', h: 'Project' }, { s: 'Invoices', h: 'Project' }, { s: 'Payments', h: 'Project' },
    { prefix: 'Transactions_', h: 'Project' }
  ]
};
const PD_SYSTEM_NAME = { pm: 'PM', ims: 'IMS', fin: 'Finance' };

function registerSystemSheet(system, id) {
  requireEditor_();
  system = String(system || '').toLowerCase();
  if (system !== 'ims' && system !== 'finance') throw new Error('Unknown system: ' + system);
  if (!/^[A-Za-z0-9_-]{20,}$/.test(String(id || ''))) throw new Error('Bad spreadsheet id');
  if (getConfigValue('sheet_id_' + system) !== String(id)) setConfigValue_internal_('sheet_id_' + system, String(id));
  return true;
}

function _pdSheetId_(sys) {
  if (sys === 'ims') return getConfigValue('sheet_id_ims') || PD_IMS_DEFAULT_ID;
  if (sys === 'fin') return getConfigValue('sheet_id_finance') || '';
  return '';
}

/** Matching rules for one project. */
function _pdContext_(projectId, pmLabel) {
  const id = String(projectId || '').trim().toUpperCase();
  if (!/^PRJ-\d+$/.test(id)) throw new Error('Bad project id: ' + projectId);
  const rec = readAll_('Projects').filter(function (r) { return String(r.Project_ID).toUpperCase() === id; })[0];
  const label = String(pmLabel || (rec ? rec.Project_ID + ' — ' + rec.Name : id));
  const map = getProjectMap();                       // KEY(upper) → Finance project label
  const finLabels = {}, aliases = {};
  Object.keys(map).forEach(function (k) { if (k === id) finLabels[String(map[k]).trim().toUpperCase()] = map[k]; });
  Object.keys(map).forEach(function (k) { if (finLabels[String(map[k]).trim().toUpperCase()]) aliases[k] = true; });
  const re = new RegExp('(^|[^A-Z0-9])' + id.replace(/[-]/g, '\\-') + '(?![0-9])');
  const hasId = function (t) { return re.test(String(t).toUpperCase()); };
  return {
    id: id, label: label, inMasters: !!rec, aliasKeys: Object.keys(aliases),
    test: {
      pm:  function (t) { return hasId(t) || String(t).trim().toUpperCase() === label.toUpperCase(); },
      ims: function (t) { return hasId(t) || !!aliases[String(t).trim().toUpperCase()]; },
      fin: function (t) { return hasId(t) || !!finLabels[String(t).trim().toUpperCase()]; }
    }
  };
}

/** One target → { name, sh, rows:[{n, vals}], head, note }. Nothing is changed. */
function _pdPlanSheet_(ss, name, spec, test) {
  const sh = ss.getSheetByName(name);
  if (!sh) return { name: name, rows: [], note: 'sheet not found' };
  const last = sh.getLastRow(), width = sh.getLastColumn();
  if (last < 2) return { name: name, sh: sh, rows: [] };
  const head = sh.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim().toLowerCase(); });
  const col = spec.c || (head.indexOf(String(spec.h).toLowerCase()) + 1);
  if (!col) return { name: name, rows: [], note: 'no "' + spec.h + '" column' };
  const rows = [];
  sh.getRange(2, 1, last - 1, width).getValues().forEach(function (r, i) {
    if (r[col - 1] !== '' && test(r[col - 1])) rows.push({ n: i + 2, vals: r });
  });
  return { name: name, sh: sh, rows: rows, head: head };
}

/** All plan items for one system. */
function _pdPlan_(ss, sys, ctx) {
  const test = ctx.test[sys], out = [];
  PD_TARGETS[sys].forEach(function (spec) {
    const names = spec.prefix
      ? ss.getSheets().map(function (s) { return s.getName(); }).filter(function (n) { return n.indexOf(spec.prefix) === 0; })
      : [spec.s];
    names.forEach(function (name) {
      const item = _pdPlanSheet_(ss, name, spec, test);
      item.codeH = spec.codeH;
      out.push(item);
      if (spec.lines && item.rows.length) {  // child rows of the deleted parents (FG Delivery Lines by DC No)
        const k = item.head.indexOf(String(spec.keyH).toLowerCase());
        const keys = {};
        item.rows.forEach(function (r) { keys[String(r.vals[k])] = true; });
        out.push(_pdPlanSheet_(ss, spec.lines.s, spec.lines, function (t) { return !!keys[String(t)]; }));
      }
    });
  });
  return out;
}

function _pdCounts_(plan) {
  return plan.filter(function (p) { return p.rows.length || p.note; })
    .map(function (p) { return { sheet: p.name, n: p.rows.length, note: p.note || '' }; });
}

function _pdOpen_(sys, pmSheetId) {
  const id = sys === 'pm' ? pmSheetId : _pdSheetId_(sys);
  if (!id) return { err: PD_SYSTEM_NAME[sys] + ' spreadsheet is not registered yet: open ' + PD_SYSTEM_NAME[sys] + ' once, then try again.' };
  try { return { ss: SpreadsheetApp.openById(id) }; }
  catch (e) { return { err: 'Cannot open ' + PD_SYSTEM_NAME[sys] + ' (' + e.message + '). You need edit access to it.' }; }
}

function previewProjectDelete(projectId, pmSheetId, pmLabel) {
  requireEditor_();
  const ctx = _pdContext_(projectId, pmLabel);
  const out = { projectId: ctx.id, label: ctx.label, inMasters: ctx.inMasters, aliases: ctx.aliasKeys, systems: {}, problems: [] };
  ['fin', 'ims', 'pm'].forEach(function (sys) {
    const o = _pdOpen_(sys, pmSheetId);
    if (o.err) { out.problems.push(o.err); out.systems[sys] = []; return; }
    out.systems[sys] = _pdCounts_(_pdPlan_(o.ss, sys, ctx));
  });
  return out;
}

/** Copy the rows to the system's backup sheet, then delete them bottom-up (contiguous runs in one call). */
function _pdApply_(ss, plan, reqId, projectId) {
  const now = new Date(), me = getUserEmail();
  const bk = ss.getSheetByName(PD_BACKUP) || ss.insertSheet(PD_BACKUP);
  if (bk.getLastRow() < 1) {
    bk.getRange(1, 1, 1, 6).setValues([['Deleted_At', 'Request', 'Project', 'Sheet', 'Row_Was', 'Row_JSON']]).setFontWeight('bold');
    bk.hideSheet();
  }
  const log = [];
  plan.forEach(function (p) {
    if (!p.rows.length) return;
    const lines = p.rows.map(function (r) { return [now, reqId, projectId, p.name, r.n, JSON.stringify(r.vals)]; });
    bk.getRange(bk.getLastRow() + 1, 1, lines.length, 6).setValues(lines);   // backup first, always
    const nums = p.rows.map(function (r) { return r.n; }).sort(function (a, b) { return b - a; });
    let i = 0;
    while (i < nums.length) {                                                 // runs of consecutive rows
      let j = i; while (j + 1 < nums.length && nums[j + 1] === nums[j] - 1) j++;
      p.sh.deleteRows(nums[j], j - i + 1);
      i = j + 1;
    }
    log.push({ sheet: p.name, n: p.rows.length });
  });
  return log;
}

/** IMS: hand the affected item codes to IMS so it refreshes their stock snapshots in its own context. */
function _pdQueueImsRefresh_(ss, plan) {
  const codes = {};
  plan.forEach(function (p) {
    if (!p.codeH || !p.rows.length) return;
    const k = p.head.indexOf(String(p.codeH).toLowerCase());
    if (k >= 0) p.rows.forEach(function (r) { if (r.vals[k]) codes[String(r.vals[k])] = true; });
  });
  const list = Object.keys(codes);
  if (!list.length) return 0;
  const sh = ss.getSheetByName(PD_REFRESH) || ss.insertSheet(PD_REFRESH);
  sh.hideSheet();
  sh.getRange(sh.getLastRow() + 1, 1, list.length, 1).setValues(list.map(function (c) { return [c]; }));
  return list.length;
}

function executeProjectDelete(projectId, pmSheetId, pmLabel, typedId) {
  requireEditor_();
  const ctx = _pdContext_(projectId, pmLabel);
  if (String(typedId || '').trim().toUpperCase() !== ctx.id) throw new Error('Confirmation does not match ' + ctx.id + '. Nothing was deleted.');
  return withLock_(function () {
    const reqId = 'PDEL-' + Utilities.getUuid().slice(0, 8);
    const res = { requestId: reqId, projectId: ctx.id, label: ctx.label, deleted: {}, problems: [], complete: false };

    // open everything first: if a system cannot be reached, nothing at all is deleted
    const open = {};
    ['fin', 'ims', 'pm'].forEach(function (sys) {
      const o = _pdOpen_(sys, pmSheetId);
      if (o.err) res.problems.push(o.err); else open[sys] = o.ss;
    });
    if (res.problems.length) return res;

    ['fin', 'ims', 'pm'].forEach(function (sys) {      // PM last: it is where a re-run starts from
      if (res.problems.length) return;                 // an earlier system failed: leave the rest untouched
      try {
        const plan = _pdPlan_(open[sys], sys, ctx);
        res.deleted[sys] = _pdApply_(open[sys], plan, reqId, ctx.id);
        if (sys === 'ims') res.imsCodesQueued = _pdQueueImsRefresh_(open[sys], plan);
      } catch (e) { res.problems.push(PD_SYSTEM_NAME[sys] + ': ' + e.message); }
    });
    if (res.problems.length) return res;

    // Masters itself: the project record and its pairings
    const sh = sh_('Projects');
    const rowNum = _findRowNum(sh, headerIndex_('Projects')['Project_ID'], ctx.id);
    if (rowNum) {
      const headers = SCHEMA['Projects'].headers;
      const old = sh.getRange(rowNum, 1, 1, headers.length).getValues()[0];
      sh.deleteRow(rowNum);
      audit_('Projects', ctx.id, 'HardDelete', '', JSON.stringify(_rowToObj_(headers, old)), '', 'Project delete ' + reqId);
    }
    const keys = {}; ctx.aliasKeys.concat([ctx.id]).forEach(function (k) { keys[k] = true; });
    const ms = ss_().getSheetByName(PMAP_SHEET);
    if (ms && ms.getLastRow() >= 2) {
      const ids = ms.getRange(2, 1, ms.getLastRow() - 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) if (keys[String(ids[i][0]).trim().toUpperCase()]) ms.deleteRow(i + 2);
    }
    const lg = ss_().getSheetByName(PD_LOG) || ss_().insertSheet(PD_LOG);
    if (lg.getLastRow() < 1) lg.getRange(1, 1, 1, 6).setValues([['Request_ID', 'Project_ID', 'Project', 'At', 'By', 'Deleted']]).setFontWeight('bold');
    lg.appendRow([reqId, ctx.id, ctx.label, new Date(), getUserEmail(), JSON.stringify(res.deleted)]);
    res.complete = true;
    return res;
  });
}
