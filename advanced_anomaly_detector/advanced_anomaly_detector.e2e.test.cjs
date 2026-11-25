'use strict';
// End-to-end tests for advanced_anomaly_detector.js: main() runs against a mocked account and sheets.
// The script sends legacy AWQL reports (CAMPAIGN_PERFORMANCE_REPORT ... DURING a,b); the shared mock
// only parses GAQL, so this file answers those queries itself.
// Run: node --test advanced_anomaly_detector
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { MockGoogleAds, MockSpreadsheet, loadScript } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = process.env.AAD_SCRIPT || path.join(__dirname, 'advanced_anomaly_detector.js');
const CONTROL_URL = 'https://docs.google.com/spreadsheets/d/demo-sheet/edit';
const SETTINGS_URL = 'https://docs.google.com/spreadsheets/d/settings-sheet/edit';
const OUTPUT_URL = 'https://docs.google.com/spreadsheets/d/output-sheet/edit';

const HEADER = ['ALERT_NAME', 'ID', 'NAME', 'EMAIL', 'FLAG', 'SUBJECT_PREFIX', 'SEND_EMAIL', 'EMAIL_LAST_SENT_TIME',
  'SEND_SLACK', 'SLACK_LAST_SENT_TIME', 'NOTIFICATION_FREQUENCY', 'WEEKEND_NOTIFICATIONS', 'REPORT_TYPE',
  'PREVIOUS_PERIOD_START', 'PREVIOUS_PERIOD_END', 'CURRENT_PERIOD_START', 'CURRENT_PERIOD_END', 'SETTINGS_URL', 'OUTPUT_SHEET_URL'];
const LOGS_COL = HEADER.length + 1; // the output URL is written just left of Logs

function yyyymmdd(n) {
  const d = new Date(Date.now() - n * 86400000);
  return d.toISOString().slice(0, 10).replace(/-/g, '');
}

class AwqlAds extends MockGoogleAds {
  constructor(options) { super(options); this.awql = []; this.sentQueries = []; this.awqlFail = null; }
  /** One day of campaign stats; daysAgo is relative to now. */
  addCampaignDay({ name, id, daysAgo, clicks = 0, impressions = 0, cost = 0 }) {
    this.awql.push({ CampaignName: name, CampaignId: String(id), date: yyyymmdd(daysAgo), Clicks: clicks, Impressions: impressions, Cost: cost });
  }
  report(query) {
    this.sentQueries.push(query);
    const m = String(query).match(/^SELECT (.+?) FROM\s+(\w+)\s*(?:WHERE (.+?))?\s*DURING (\d{8}),(\d{8})$/s);
    if (!m) return super.report(query);
    if (this.awqlFail) throw new Error(this.awqlFail);
    const fields = m[1].split(',').map(s => s.trim());
    const [from, to] = [m[4], m[5]];
    const groups = new Map();
    for (const r of this.awql) {
      if (r.date < from || r.date > to) continue;
      // AWQL leaves out rows with no impressions
      const g = groups.get(r.CampaignId) || { CampaignName: r.CampaignName, CampaignId: r.CampaignId, Clicks: 0, Impressions: 0, Cost: 0 };
      g.Clicks += r.Clicks; g.Impressions += r.Impressions; g.Cost += r.Cost;
      groups.set(r.CampaignId, g);
    }
    const rows = [...groups.values()].filter(g => g.Impressions > 0).map(g => {
      const out = {};
      for (const f of fields) out[f] = typeof g[f] === 'number' ? g[f].toLocaleString('en-US') : (g[f] ?? '');
      return out;
    });
    let i = 0;
    return { rows: () => ({ hasNext: () => i < rows.length, next: () => rows[i++], totalNumEntities: () => rows.length }) };
  }
}

function spreadsheetSupport(ss) {
  const sheet = ss.insertSheet('__probe');
  const rangeProto = Object.getPrototypeOf(sheet.getRange(1, 1));
  const sheetProto = Object.getPrototypeOf(sheet);
  ss.sheets = ss.sheets.filter(s => s.name !== '__probe');
  rangeProto.getNotes = function () { return this.getValues().map(line => line.map(() => '')); };
  sheetProto.clearContents = function () { this.cells.clear(); return this; };
  sheetProto.sort = function (column, ascending) {
    const rows = this.rowsFrom(1);
    rows.sort((a, b) => (ascending ? 1 : -1) * String(a[column - 1]).localeCompare(String(b[column - 1])));
    this.cells.clear();
    rows.forEach((line, r) => line.forEach((v, c) => this._set(r + 1, c + 1, v)));
    return this;
  };
  const ssProto = Object.getPrototypeOf(ss);
  ssProto.rename = function (name) { this.renamed = name; };
  ssProto.getEditors = function () { return []; };
  ssProto.addEditor = function () {};
}

function makeAccount({ rows, alerts = [['Clicks', 'Decreases by', 0.5]], customerId = '123-456-7890' } = {}) {
  const ads = new AwqlAds({ customerId, accountName: 'Shop' });
  spreadsheetSupport(ads.spreadsheet);
  const control = ads.spreadsheet.insertSheet('Anomaly Detector');
  HEADER.forEach((key, i) => control.getRange(3, i + 1).setValue(key));
  control.getRange(3, LOGS_COL).setValue('Logs');
  control.getRange(3, LOGS_COL + 1).setValue('Last run');
  (rows || [rule()]).forEach((r, n) => HEADER.forEach((key, i) => { if (key in r) control.getRange(4 + n, i + 1).setValue(r[key]); }));

  const settingsSs = new MockSpreadsheet('settings-sheet');
  const settingsSheet = settingsSs.insertSheet('Anomaly Detector');
  settingsSs.setActiveSheet(settingsSheet);
  alerts.forEach(([metric, op, param], n) => {
    settingsSheet.getRange(4 + n, 13).setValue(metric);
    settingsSheet.getRange(4 + n, 14).setValue(op);
    settingsSheet.getRange(4 + n, 15).setValue(param);
  });
  const outputSs = new MockSpreadsheet('output-sheet');
  outputSs.setActiveSheet(outputSs.insertSheet('Sheet1'));
  const books = [ads.spreadsheet, settingsSs, outputSs];
  return { ads, control, settingsSs, outputSs, books };
}

function rule(overrides = {}) {
  return {
    ALERT_NAME: 'Clicks drop', ID: '123-456-7890', NAME: 'Shop', EMAIL: 'a@example.com', FLAG: 'Yes',
    SUBJECT_PREFIX: 'Urgent', SEND_EMAIL: 'Yes', SEND_SLACK: 'No', NOTIFICATION_FREQUENCY: 1,
    WEEKEND_NOTIFICATIONS: 'Yes', REPORT_TYPE: 'Campaign', PREVIOUS_PERIOD_START: 14, PREVIOUS_PERIOD_END: 8,
    CURRENT_PERIOD_START: 7, CURRENT_PERIOD_END: 1, SETTINGS_URL, OUTPUT_SHEET_URL: OUTPUT_URL, ...overrides,
  };
}

function run(ctx) {
  const { ads, books } = ctx;
  const globals = ads.globals({
    SpreadsheetApp: {
      openByUrl: url => {
        const book = books.find(b => String(url).includes(b.id));
        if (!book) throw new Error(`No spreadsheet ${url}`);
        return book;
      },
    },
    DriveApp: {
      getRootFolder: () => ({ addFile() {} }),
      getFolderById: () => ({ addFile() {} }),
      getFileById: id => ({ id }),
      getFilesByName: name => ({ hasNext: () => true, next: () => ({ name }) }),
    },
  });
  const loaded = loadScript(SCRIPT, globals, ['main'], { replace: [['YOUR_SPREADSHEET_URL_HERE', CONTROL_URL]] });
  loaded.exports.main();
  return loaded;
}

function seed(ads) {
  // A: 140 -> 35 clicks (-75%), alert.  B: 140 -> 126 (-10%), no alert.  C: stops completely, alert.
  for (let d = 8; d <= 14; d++) {
    ads.addCampaignDay({ name: 'A', id: 1, daysAgo: d, clicks: 20, impressions: 200 });
    ads.addCampaignDay({ name: 'B', id: 2, daysAgo: d, clicks: 20, impressions: 200 });
    ads.addCampaignDay({ name: 'C', id: 3, daysAgo: d, clicks: 20, impressions: 200 });
  }
  for (let d = 1; d <= 7; d++) {
    ads.addCampaignDay({ name: 'A', id: 1, daysAgo: d, clicks: 5, impressions: 100 });
    ads.addCampaignDay({ name: 'B', id: 2, daysAgo: d, clicks: 18, impressions: 200 });
  }
}

const output = ctx => ctx.outputSs.getActiveSheet().rowsFrom(1);

test('normal run: falls in clicks, including a campaign that stopped, are written and emailed', () => {
  const ctx = makeAccount();
  seed(ctx.ads);
  const { errors } = run(ctx);
  assert.deepEqual(errors, []);
  const rows = output(ctx);
  assert.deepEqual(rows[0], ['Campaign Name', 'Clicks (before)', 'Clicks (after)', 'Clicks % change']);
  const byName = Object.fromEntries(rows.slice(1).map(r => [r[0], r]));
  assert.deepEqual(Object.keys(byName).sort(), ['A', 'C']);
  assert.deepEqual(byName.A.slice(1), [140, 35, '-75%']);
  assert.deepEqual(byName.C.slice(1), [140, 0, '-100%']);
  for (const line of rows) for (const v of line) assert.ok(!/NaN|Infinity|undefined/.test(String(v)), String(v));
  assert.equal(ctx.ads.emails.length, 1);
  assert.equal(ctx.ads.emails[0][0].to, 'a@example.com');
  assert.match(ctx.ads.emails[0][0].subject, /^Urgent Anomaly Alert: Shop - Clicks drop/);
  assert.equal(Object.prototype.toString.call(ctx.control.getRange(4, HEADER.indexOf('EMAIL_LAST_SENT_TIME') + 1).getValue()), '[object Date]');
  assert.match(String(ctx.control.getRange(4, LOGS_COL).getValue()), /ran sucessfully/);
  // Periods: previous 14..8 days ago, current 7..1 days ago
  assert.ok(ctx.ads.sentQueries.some(q => q.endsWith(`DURING ${yyyymmdd(14)},${yyyymmdd(8)}`)));
  assert.ok(ctx.ads.sentQueries.some(q => q.endsWith(`DURING ${yyyymmdd(7)},${yyyymmdd(1)}`)));
});

test('the whole account going quiet (no rows in the current period) raises alerts', () => {
  const ctx = makeAccount();
  for (let d = 8; d <= 14; d++) ctx.ads.addCampaignDay({ name: 'A', id: 1, daysAgo: d, clicks: 20, impressions: 200 });
  run(ctx);
  assert.deepEqual(output(ctx).slice(1), [['A', 140, 0, '-100%']]);
  assert.equal(ctx.ads.emails.length, 1);
});

test('no alerts: nothing emailed and the log says so', () => {
  const ctx = makeAccount();
  for (let d = 1; d <= 14; d++) ctx.ads.addCampaignDay({ name: 'B', id: 2, daysAgo: d, clicks: 20, impressions: 200 });
  run(ctx);
  assert.equal(ctx.ads.emails.length, 0);
  assert.match(String(ctx.control.getRange(4, LOGS_COL).getValue()), /No alerts/);
});

test('an email sent within the notification frequency is not repeated', () => {
  const ctx = makeAccount({ rows: [rule({ EMAIL_LAST_SENT_TIME: new Date(Date.now() - 3600 * 1000), NOTIFICATION_FREQUENCY: 1 })] });
  seed(ctx.ads);
  run(ctx);
  assert.equal(ctx.ads.emails.length, 0);
  assert.equal(output(ctx).length, 3); // the sheet is still updated
});

test('empty account: no alerts, no errors', () => {
  const ctx = makeAccount();
  const { errors } = run(ctx);
  assert.deepEqual(errors, []);
  assert.equal(ctx.ads.emails.length, 0);
});

test("rows for this account run even when another account's row comes first", () => {
  const ctx = makeAccount({ rows: [rule({ ID: '999-999-9999', ALERT_NAME: 'Other account' }), rule()] });
  seed(ctx.ads);
  run(ctx);
  assert.match(String(ctx.control.getRange(5, LOGS_COL).getValue()), /ran sucessfully/);
  assert.equal(ctx.ads.emails.length, 1);
  assert.match(ctx.ads.emails[0][0].subject, /Clicks drop/);
});

test('a "yes" flag typed in lower case, or a checkbox, does not stop the script', () => {
  const ctx = makeAccount({ rows: [rule({ FLAG: 'yes' }), rule({ FLAG: true, ALERT_NAME: 'Checkbox' })] });
  seed(ctx.ads);
  const { errors } = run(ctx);
  assert.deepEqual(errors, []);
  assert.equal(ctx.ads.emails.length, 1);
});

test('a failing report is logged and leaves the row marked as not finished', () => {
  const ctx = makeAccount();
  seed(ctx.ads);
  ctx.ads.awqlFail = 'AWQL is not supported';
  const { errors } = run(ctx);
  assert.equal(errors.length, 1); // console.error(e) of the Error
  assert.match(String(ctx.control.getRange(4, LOGS_COL).getValue()), /did not run sucessfully/);
  assert.equal(ctx.ads.emails.length, 0);
});

test('zero clicks in both periods is not an alert and writes no NaN', () => {
  const ctx = makeAccount({ alerts: [['Clicks', 'Increases by', -100]] }); // report everything
  for (let d = 1; d <= 14; d++) ctx.ads.addCampaignDay({ name: 'Z', id: 9, daysAgo: d, clicks: 0, impressions: 10 });
  run(ctx);
  const rows = output(ctx);
  assert.deepEqual(rows.slice(1), [['Z', 0, 0, '0%']]);
});
