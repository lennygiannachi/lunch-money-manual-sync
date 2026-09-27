/**
 * Sync the "Tracking" tab of this spreadsheet into Lunch Money.
 *
 * Apps Script port of manual_tracking_sync.py. Paste into Extensions -> Apps
 * Script on the "Lunch Money - Manual Tracking" sheet, and it adds a
 * "Lunch Money" menu:
 *
 *   Preview sync...   dry run in a dialog, with an Apply button
 *   Sync now          apply straight away, then show what happened
 *   Set API token...  stores your Lunch Money token in your user properties
 *
 * Column mapping (header names, case-insensitive, any order):
 *
 *   Account       -> manual account name (the match key)
 *   Platform      -> institution_name
 *   Category      -> account type (Cash, Investment, Cryptocurrency, Loan, ...)
 *   Subtype       -> subtype (free text, e.g. "Retirement", "Student Loan")
 *   Balance       -> balance
 *   Last Updated  -> balance_as_of
 *   Notes         -> never sent to Lunch Money
 *
 * Rows sharing an Account name are summed into one account, carrying the
 * newest Last Updated of the group. Liability balances (loan, credit, other
 * liability) are passed through as entered: positive means "amount owed".
 *
 * Lunch Money's account list falls back to "{institution_name} {name}" when
 * display_name is blank. For Cryptocurrency rows with a Platform, display_name
 * is set to the plain Account name so "ADA" doesn't turn into "Coinbase ADA".
 *
 * Editing a Balance cell stamps the current time into Last Updated on that row
 * (onEdit, below), which is what the sync sends as balance_as_of.
 *
 * If this project already has an onOpen() in another file, delete the onOpen
 * below and call lmAddMenu_() from yours.
 *
 * For unattended syncing, add a time-driven trigger on lmScheduledSync
 * (Triggers -> Add Trigger). Failures throw, so the trigger emails you.
 */

// ---------------------------------------------------------------------------
// CONFIG
// ---------------------------------------------------------------------------
const LM_TAB_NAME = 'Tracking';
const LM_BASE_URL = 'https://api.lunchmoney.dev/v2';
const LM_DEFAULT_CURRENCY = 'usd';
const LM_PAUSE_MS = 350; // between calls, to stay clear of rate limits
const LM_TOKEN_PROPERTY = 'LUNCHMONEY_TOKEN';

// Column positions on the Tracking tab (A = 1) used by the onEdit timestamp.
const LM_BALANCE_COL = 5;       // E: Balance
const LM_LAST_UPDATED_COL = 6;  // F: Last Updated

// Lunch Money field limits, enforced before we send anything.
const LM_MAX_NAME = 45;
const LM_MAX_INSTITUTION = 50;
const LM_MAX_SUBTYPE = 75;

// Balances are handled as integer ten-thousandths to avoid float drift when
// summing, matching the API's 4-decimal precision.
const LM_SCALE = 10000;

// The values on the right are the only ones the v2 API accepts. Add aliases on
// the left freely -- lookup is case- and whitespace-insensitive.
const LM_TYPE_MAP = {
  'cash': 'cash',
  'checking': 'cash',
  'savings': 'cash',
  'bank': 'cash',
  'credit': 'credit',
  'credit card': 'credit',
  'cryptocurrency': 'cryptocurrency',
  'crypto': 'cryptocurrency',
  'employee compensation': 'employee compensation',
  'equity': 'employee compensation',
  'rsu': 'employee compensation',
  'investment': 'investment',
  'investments': 'investment',
  'brokerage': 'investment',
  'retirement': 'investment',
  'loan': 'loan',
  'loans': 'loan',
  'mortgage': 'loan',
  'other liability': 'other liability',
  'liability': 'other liability',
  'other asset': 'other asset',
  'asset': 'other asset',
  'real estate': 'real estate',
  'property': 'real estate',
  'vehicle': 'vehicle',
  'car': 'vehicle',
};

// ---------------------------------------------------------------------------
// MENU + ENTRY POINTS
// ---------------------------------------------------------------------------
function onOpen() {
  lmAddMenu_();
}

function lmAddMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('Lunch Money')
    .addItem('Preview sync...', 'lmPreviewSync')
    .addItem('Sync now', 'lmSyncNow')
    .addSeparator()
    .addItem('Set API token...', 'lmSetToken')
    .addToUi();
}

function lmSetToken() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Lunch Money API token',
    'Paste your token from https://my.lunchmoney.app/developers\n' +
      '(stored in your user properties, not in the sheet or the code)',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const token = res.getResponseText().trim();
  if (!token) return;
  PropertiesService.getUserProperties().setProperty(LM_TOKEN_PROPERTY, token);
  SpreadsheetApp.getActiveSpreadsheet().toast('Token saved.', 'Lunch Money');
}

function lmPreviewSync() {
  let plan;
  try {
    plan = lmBuildPlan_();
  } catch (e) {
    SpreadsheetApp.getUi().alert('Lunch Money sync', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const canApply = plan.toCreate.length > 0 || plan.toUpdate.length > 0;
  lmShowDialog_(lmFormatPlan_(plan, true), canApply);
}

function lmSyncNow() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Syncing...', 'Lunch Money', -1);
  let text;
  try {
    text = lmRunSync_().text;
  } catch (e) {
    text = 'ERROR: ' + e.message;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Done.', 'Lunch Money', 3);
  lmShowDialog_(text, false);
}

/** Called from the preview dialog's Apply button. Re-reads everything first. */
function lmApplyFromDialog() {
  return lmRunSync_().text;
}

/** For a time-driven trigger. No UI; throws on failure so the trigger reports it. */
function lmScheduledSync() {
  const result = lmRunSync_();
  console.log(result.text);
  if (result.failures.length) {
    throw new Error(result.failures.length + ' Lunch Money call(s) failed:\n' + result.failures.join('\n'));
  }
}

function lmRunSync_() {
  const plan = lmBuildPlan_();
  const text = [lmFormatPlan_(plan, false)];
  if (!plan.toCreate.length && !plan.toUpdate.length) {
    return { text: text.join('\n'), failures: [] };
  }
  const applied = lmApply_(plan);
  text.push('', applied.log.join('\n'));
  return { text: text.join('\n'), failures: applied.failures };
}

// ---------------------------------------------------------------------------
// LAST UPDATED TIMESTAMPS
// ---------------------------------------------------------------------------
/**
 * Simple trigger: runs on every manual edit. When a Balance cell on the
 * Tracking tab changes, stamp the current time into Last Updated on that row.
 * Edits made by scripts (including the sync) don't fire this.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== LM_TAB_NAME) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  // Only Balance edits below the header row
  if (col === LM_BALANCE_COL && row > 1) {
    sheet.getRange(row, LM_LAST_UPDATED_COL).setValue(new Date());
  }
}

// ---------------------------------------------------------------------------
// SHEET
// ---------------------------------------------------------------------------
function lmReadTracking_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LM_TAB_NAME);
  if (!sheet) throw new Error('No tab named "' + LM_TAB_NAME + '" in this spreadsheet.');

  const values = sheet.getDataRange().getValues();
  if (!values.length || values.every(r => r.every(c => c === ''))) {
    throw new Error('The "' + LM_TAB_NAME + '" tab is empty.');
  }

  const header = values[0].map(c => String(c).trim().toLowerCase());
  const missing = ['account', 'category', 'balance'].filter(c => header.indexOf(c) === -1);
  if (missing.length) {
    throw new Error(
      'The "' + LM_TAB_NAME + '" tab is missing column(s): ' + missing.join(', ') +
        '\nSaw: ' + header.join(', ')
    );
  }

  const idx = {};
  header.forEach((h, i) => {
    if (!(h in idx)) idx[h] = i;
  });
  const cell = (row, name) => {
    const i = idx[name];
    if (i === undefined || i >= row.length || row[i] === null) return '';
    return row[i];
  };

  const rows = values.slice(1).map((raw, i) => ({
    row: i + 2,
    account: String(cell(raw, 'account')).trim(),
    platform: String(cell(raw, 'platform')).trim(),
    category: String(cell(raw, 'category')).trim(),
    subtype: String(cell(raw, 'subtype')).trim(),
    balance: cell(raw, 'balance'),
    lastUpdated: cell(raw, 'last updated'),
    notes: String(cell(raw, 'notes')).trim(),
  }));

  return { rows: rows, tz: ss.getSpreadsheetTimeZone() || 'UTC' };
}

/** Number or decorated string -> integer ten-thousandths, or null. */
function lmParseBalance_(value) {
  if (typeof value === 'number') return Math.round(value * LM_SCALE);
  let text = String(value).trim();
  if (!text) return null;
  const negative = text.charAt(0) === '(' && text.charAt(text.length - 1) === ')';
  text = text.replace(/[()$,£€\s]/g, '');
  if (!text) return null;
  const amount = Number(text);
  if (!isFinite(amount)) return null;
  const units = Math.round(amount * LM_SCALE);
  return negative ? -units : units;
}

/** Date cell, serial number, or text -> Date, or null. */
function lmParseTimestamp_(value, tz) {
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;

  if (typeof value === 'number') {
    // A plain-number cell holding a Sheets serial: wall-clock time in tz.
    const naive = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
    const wall = Utilities.formatDate(naive, 'UTC', 'yyyy-MM-dd HH:mm:ss');
    return Utilities.parseDate(wall, tz, 'yyyy-MM-dd HH:mm:ss');
  }

  const text = String(value).trim();
  if (!text) return null;
  const formats = [
    "yyyy-MM-dd HH:mm:ss",
    "yyyy-MM-dd'T'HH:mm:ss",
    'yyyy-MM-dd',
    'MM/dd/yyyy HH:mm:ss',
    'MM/dd/yyyy',
  ];
  for (const fmt of formats) {
    try {
      return Utilities.parseDate(text, tz, fmt);
    } catch (e) {
      // try the next one
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// AGGREGATION
// ---------------------------------------------------------------------------
function lmKey_(s) {
  return String(s || '').toLowerCase().split(/\s+/).filter(Boolean).join(' ');
}

/** Collapse rows to one entry per account name. */
function lmAggregate_(rows, tz) {
  const groups = {};
  const order = [];
  const problems = [];

  for (const r of rows) {
    if (!r.account) {
      if (r.platform || r.category || String(r.balance).trim()) {
        problems.push('row ' + r.row + ': has data but no Account name, skipped');
      }
      continue;
    }

    const type = LM_TYPE_MAP[lmKey_(r.category)];
    if (!type) {
      problems.push('row ' + r.row + ' (' + r.account + '): unrecognized Category "' +
        (r.category || '<blank>') + '", skipped');
      continue;
    }

    const balance = lmParseBalance_(r.balance);
    if (balance === null) {
      problems.push('row ' + r.row + ' (' + r.account + '): unreadable Balance "' + r.balance + '", skipped');
      continue;
    }

    const key = lmKey_(r.account);
    let g = groups[key];
    if (!g) {
      g = groups[key] = {
        name: r.account, type: type, balance: 0, asOf: null,
        platform: '', subtype: '', rows: [], typeConflict: false,
      };
      order.push(key);
    }
    if (g.type !== type) g.typeConflict = true;
    g.balance += balance;
    g.rows.push(r.row);

    const stamp = lmParseTimestamp_(r.lastUpdated, tz);
    if (stamp && (!g.asOf || stamp.getTime() > g.asOf.getTime())) {
      g.asOf = stamp;
      if (r.platform) g.platform = r.platform; // newest row wins
      if (r.subtype) g.subtype = r.subtype;
    } else {
      if (!g.platform && r.platform) g.platform = r.platform;
      if (!g.subtype && r.subtype) g.subtype = r.subtype;
    }
  }

  const accounts = [];
  for (const key of order) {
    const g = groups[key];
    if (g.typeConflict) {
      problems.push('"' + g.name + '": rows ' + g.rows.join(', ') + ' disagree on Category, skipped');
      continue;
    }
    if (g.name.length > LM_MAX_NAME) {
      problems.push('"' + g.name + '": name is ' + g.name.length + ' characters, Lunch Money allows ' +
        LM_MAX_NAME + ', skipped');
      continue;
    }
    if (g.platform.length > LM_MAX_INSTITUTION) {
      problems.push('"' + g.name + '": Platform truncated to ' + LM_MAX_INSTITUTION + ' characters');
      g.platform = g.platform.slice(0, LM_MAX_INSTITUTION);
    }
    if (g.subtype.length > LM_MAX_SUBTYPE) {
      problems.push('"' + g.name + '": Subtype truncated to ' + LM_MAX_SUBTYPE + ' characters');
      g.subtype = g.subtype.slice(0, LM_MAX_SUBTYPE);
    }
    accounts.push(g);
  }

  accounts.sort((a, b) => {
    const x = a.name.toLowerCase(), y = b.name.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  });
  return { accounts: accounts, problems: problems };
}

/** Integer ten-thousandths -> the string shape the API's balance pattern expects. */
function lmMoney_(units) {
  const abs = Math.abs(units);
  const whole = Math.floor(abs / LM_SCALE);
  const frac = abs % LM_SCALE;
  let text = String(whole);
  if (frac) text += '.' + String(frac).padStart(4, '0').replace(/0+$/, '');
  return (units < 0 ? '-' : '') + text;
}

function lmUtcIso_(date) {
  return Utilities.formatDate(date, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'+00:00'");
}

// ---------------------------------------------------------------------------
// LUNCH MONEY API
// ---------------------------------------------------------------------------
function lmToken_() {
  const token = PropertiesService.getUserProperties().getProperty(LM_TOKEN_PROPERTY);
  if (!token) {
    throw new Error('No Lunch Money token. Use Lunch Money -> Set API token... first.\n' +
      'Get it at https://my.lunchmoney.app/developers');
  }
  return token;
}

function lmRequest_(method, path, payload) {
  const options = {
    method: method.toLowerCase(),
    headers: { Authorization: 'Bearer ' + lmToken_(), Accept: 'application/json' },
    muteHttpExceptions: true,
  };
  if (payload !== undefined) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }

  let resp;
  try {
    resp = UrlFetchApp.fetch(LM_BASE_URL + path, options);
  } catch (e) {
    return { ok: false, error: 'network error: ' + e.message };
  }

  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code >= 200 && code < 300) {
    return { ok: true, data: body.trim() ? JSON.parse(body) : {} };
  }

  let detail;
  try {
    const parsed = JSON.parse(body);
    const errs = parsed.errors || [];
    detail = errs.map(x => x.errMsg || JSON.stringify(x)).join('; ') || parsed.message || body;
  } catch (e) {
    detail = body.slice(0, 300);
  }
  return { ok: false, error: 'HTTP ' + code + ': ' + detail };
}

function lmFetchManualAccounts_() {
  const res = lmRequest_('GET', '/manual_accounts');
  if (!res.ok) throw new Error('Could not read Lunch Money accounts -- ' + res.error);
  return res.data.manual_accounts || [];
}

/** Name -> account. Falls back to display_name; real name wins on a collision. */
function lmIndexByName_(accounts) {
  const index = {};
  for (const a of accounts) {
    const key = lmKey_(a.display_name);
    if (key && !(key in index)) index[key] = a;
  }
  for (const a of accounts) index[lmKey_(a.name)] = a;
  return index;
}

/** What would change on an update, as human-readable strings. */
function lmDiff_(existing, target, tz) {
  const changes = [];

  const oldBalance = Math.round(Number(existing.balance || 0) * LM_SCALE);
  if (oldBalance !== target.balance) {
    changes.push('balance ' + lmMoney_(oldBalance) + ' -> ' + lmMoney_(target.balance));
  }

  if ((existing.type || '') !== target.type) {
    changes.push('type ' + (existing.type || 'none') + ' -> ' + target.type);
  }

  const oldInstitution = existing.institution_name || '';
  if (target.platform && oldInstitution !== target.platform) {
    changes.push('institution ' + (oldInstitution || 'none') + ' -> ' + target.platform);
  }

  const oldSubtype = existing.subtype || '';
  if (target.subtype && oldSubtype !== target.subtype) {
    changes.push('subtype ' + (oldSubtype || 'none') + ' -> ' + target.subtype);
  }

  if (target.type === 'cryptocurrency' && target.platform) {
    const oldDisplay = existing.display_name || '';
    if (oldDisplay !== target.name) {
      changes.push('display_name ' + (oldDisplay || 'none') + ' -> ' + target.name);
    }
  }

  if (target.asOf) {
    const oldAsOf = existing.balance_as_of || '';
    if (oldAsOf.slice(0, 19) !== lmUtcIso_(target.asOf).slice(0, 19)) {
      changes.push('balance_as_of -> ' + Utilities.formatDate(target.asOf, tz, 'yyyy-MM-dd HH:mm:ss z'));
    }
  }

  return changes;
}

function lmBuildPayload_(target, creating) {
  const payload = {
    name: target.name,
    type: target.type,
    balance: lmMoney_(target.balance),
  };
  if (target.asOf) payload.balance_as_of = lmUtcIso_(target.asOf);
  if (target.platform) {
    payload.institution_name = target.platform;
    if (target.type === 'cryptocurrency') payload.display_name = target.name;
  }
  if (target.subtype) payload.subtype = target.subtype;
  if (creating) payload.currency = LM_DEFAULT_CURRENCY;
  return payload;
}

// ---------------------------------------------------------------------------
// PLAN + APPLY
// ---------------------------------------------------------------------------
function lmBuildPlan_() {
  lmToken_(); // fail fast before reading anything

  const sheet = lmReadTracking_();
  const agg = lmAggregate_(sheet.rows, sheet.tz);
  const existingAccounts = lmFetchManualAccounts_();
  const index = lmIndexByName_(existingAccounts);

  const toCreate = [], toUpdate = [], unchanged = [];
  for (const target of agg.accounts) {
    const existing = index[lmKey_(target.name)];
    if (!existing) {
      toCreate.push(target);
      continue;
    }
    const changes = lmDiff_(existing, target, sheet.tz);
    if (changes.length) toUpdate.push({ existing: existing, target: target, changes: changes });
    else unchanged.push(target);
  }

  return {
    rowCount: sheet.rows.length,
    tz: sheet.tz,
    accounts: agg.accounts,
    problems: agg.problems,
    existingCount: existingAccounts.length,
    toCreate: toCreate,
    toUpdate: toUpdate,
    unchanged: unchanged,
  };
}

function lmFormatPlan_(plan, dryRun) {
  const out = [];
  const rule = '='.repeat(70);
  const pad = (s, n) => String(s).padEnd(n);
  const lpad = (s, n) => String(s).padStart(n);

  out.push(plan.rowCount + ' data rows from ' + LM_TAB_NAME + ' (timezone ' + plan.tz + ')');
  out.push(plan.accounts.length + ' accounts after summing duplicates');
  out.push(plan.existingCount + ' manual accounts in Lunch Money');
  out.push('', rule, dryRun ? 'PLAN (dry run -- nothing will change)' : 'PLAN', rule);

  if (plan.toCreate.length) {
    out.push('', 'CREATE (' + plan.toCreate.length + '):');
    for (const t of plan.toCreate) {
      out.push('  ' + pad(t.name, 42) + ' ' + pad(t.type, 14) + ' ' + pad(t.subtype || '-', 18) + ' ' +
        lpad(lmMoney_(t.balance), 14) + '  ' + (t.platform || '-') +
        (t.rows.length > 1 ? ' [rows ' + t.rows.join(',') + ']' : ''));
    }
  }

  if (plan.toUpdate.length) {
    out.push('', 'UPDATE (' + plan.toUpdate.length + '):');
    for (const u of plan.toUpdate) {
      out.push('  ' + u.target.name + ' (id ' + u.existing.id + ')' +
        (u.target.rows.length > 1 ? '  [summed from rows ' + u.target.rows.join(',') + ']' : ''));
      for (const c of u.changes) out.push('      ' + c);
    }
  }

  if (plan.unchanged.length) {
    out.push('', 'UNCHANGED (' + plan.unchanged.length + '): ' + plan.unchanged.map(t => t.name).join(', '));
  }

  if (plan.problems.length) {
    out.push('', 'PROBLEMS (' + plan.problems.length + '):');
    for (const p of plan.problems) out.push('  ' + p);
  }

  if (!plan.toCreate.length && !plan.toUpdate.length) {
    out.push('', 'Nothing to do -- Lunch Money already matches the sheet.');
  }

  return out.join('\n');
}

function lmApply_(plan) {
  const log = ['='.repeat(70), 'APPLYING', '='.repeat(70)];
  const failures = [];
  let created = 0, updated = 0;

  for (const t of plan.toCreate) {
    const res = lmRequest_('POST', '/manual_accounts', lmBuildPayload_(t, true));
    if (res.ok) {
      created++;
      log.push('  created  ' + t.name);
    } else {
      failures.push('create ' + t.name + ' -- ' + res.error);
      log.push('  FAILED   ' + t.name + ' -- ' + res.error);
    }
    Utilities.sleep(LM_PAUSE_MS);
  }

  for (const u of plan.toUpdate) {
    const res = lmRequest_('PUT', '/manual_accounts/' + u.existing.id, lmBuildPayload_(u.target, false));
    if (res.ok) {
      updated++;
      log.push('  updated  ' + u.target.name);
    } else {
      failures.push('update ' + u.target.name + ' -- ' + res.error);
      log.push('  FAILED   ' + u.target.name + ' -- ' + res.error);
    }
    Utilities.sleep(LM_PAUSE_MS);
  }

  log.push('', 'Created ' + created + ', updated ' + updated + '.');
  if (failures.length) {
    log.push(failures.length + ' call(s) failed:');
    for (const f of failures) log.push('  ' + f);
  }
  return { log: log, failures: failures };
}

// ---------------------------------------------------------------------------
// DIALOG
// ---------------------------------------------------------------------------
function lmShowDialog_(text, canApply) {
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const html =
    '<style>' +
    'body{font-family:Roboto,Arial,sans-serif;margin:0;padding:4px}' +
    'pre{font:12px/1.45 "Roboto Mono",Menlo,Consolas,monospace;white-space:pre;overflow:auto;' +
    'background:#f8f9fa;border:1px solid #dadce0;border-radius:6px;padding:12px;height:470px;margin:0 0 12px}' +
    'button{font:500 14px Roboto,Arial,sans-serif;padding:8px 16px;border-radius:4px;cursor:pointer;margin-right:8px}' +
    '#apply{background:#1a73e8;color:#fff;border:0}#apply:disabled{background:#8ab4f8;cursor:default}' +
    '.close{background:#fff;color:#1a73e8;border:1px solid #dadce0}' +
    '#status{color:#5f6368;font-size:13px}' +
    '</style>' +
    '<pre id="out">' + esc(text) + '</pre>' +
    '<div>' +
    (canApply ? '<button id="apply" onclick="applySync()">Apply to Lunch Money</button>' : '') +
    '<button class="close" onclick="google.script.host.close()">Close</button>' +
    '<span id="status"></span>' +
    '</div>' +
    '<script>' +
    'function applySync(){' +
    ' var btn=document.getElementById("apply"),st=document.getElementById("status");' +
    ' btn.disabled=true; st.textContent="Applying...";' +
    ' google.script.run' +
    '  .withSuccessHandler(function(t){document.getElementById("out").textContent=t;btn.remove();st.textContent="";})' +
    '  .withFailureHandler(function(e){st.textContent="Error: "+e.message;btn.disabled=false;})' +
    '  .lmApplyFromDialog();' +
    '}' +
    '</script>';

  const output = HtmlService.createHtmlOutput(html).setWidth(900).setHeight(570);
  SpreadsheetApp.getUi().showModalDialog(output, 'Lunch Money sync');
}
