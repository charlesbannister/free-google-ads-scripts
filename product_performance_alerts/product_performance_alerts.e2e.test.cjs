'use strict';
// End-to-end tests for product_performance_alerts.js: main() runs against a mocked account and sheet.
// Run: node --test product_performance_alerts
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { MockGoogleAds, loadScript } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = process.env.PPA_SCRIPT || path.join(__dirname, 'product_performance_alerts.js');
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/demo-sheet/edit';

// Settings columns, in the script's header order (single account).
const HEADER = ['NAME', 'EMAILS', 'FLAG', 'SEND_EMAIL', 'EMAIL_PREFIX', 'TAB_NAME', 'N', 'INCLUDE_CAMPAIGN',
  'INCLUDE_AD_GROUP', 'CAMPAIGN_NAME_CONTAINS', 'CAMPAIGN_NAME_NOT_CONTAINS', 'ITEM_ID_CONTAINS',
  'ITEM_ID_NOT_CONTAINS', 'PRODUCT_TITLE_CONTAINS', 'PRODUCT_TITLE_NOT_CONTAINS'];
for (let i = 1; i <= 6; i++) HEADER.push(`FILTER_${i}_METRIC`, `FILTER_${i}_OPERATOR`, `FILTER_${i}_VALUE`);
HEADER.push('LOG_SHEET_URL', 'LOGS_COLUMN');
const LOGS_COL = HEADER.length; // "Logs" heading sits in the LOGS_COLUMN position

function daysAgo(n) {
  const d = new Date();
  d.setUTCDate(d.getUTCDate() - n);
  return d.toISOString().slice(0, 10);
}

// The shared mock has no shopping_performance_view: map it onto keyword_view (same prefixes).
class ShoppingAds extends MockGoogleAds {
  constructor(options) { super(options); this.sentQueries = []; }
  addProduct({ itemId, title, campaign = 'Shopping UK', campaignId = '1', adGroup = 'All', adGroupId = '11', date, clicks = 0, impressions = 1, cost = 0, conversions = 0, value = 0 }) {
    return this.addRows('keyword_view', [{
      'segments.product_item_id': itemId, 'segments.product_title': title,
      'campaign.name': campaign, 'campaign.id': campaignId, 'campaign.status': 'ENABLED',
      'ad_group.name': adGroup, 'ad_group.id': adGroupId, 'ad_group.status': 'ENABLED',
      'segments.date': date, 'metrics.clicks': String(clicks), 'metrics.impressions': String(impressions),
      'metrics.cost_micros': String(Math.round(cost * 1e6)), 'metrics.conversions': String(conversions),
      'metrics.conversions_value': String(value),
    }]);
  }
  report(query) {
    this.sentQueries.push(String(query).replace(/\s+/g, ' '));
    if (/\bmetrics\.cost\b(?!_)/.test(query.replace(/select.*?from/is, ''))) {
      throw new Error("Unrecognized field in the query: 'metrics.cost'.");
    }
    return super.report(String(query).replace(/from shopping_performance_view/i, 'from keyword_view'));
  }
}

function addSheetSupport(ads) {
  const sheet = ads.spreadsheet.insertSheet('__probe');
  const rangeProto = Object.getPrototypeOf(sheet.getRange(1, 1));
  const sheetProto = Object.getPrototypeOf(sheet);
  ads.spreadsheet.sheets = ads.spreadsheet.sheets.filter(s => s.name !== '__probe');
  rangeProto.getFilter = function () { return this.sheet._filter; };
  rangeProto.sort = function ({ column, ascending }) {
    const values = this.getValues();
    const key = column - this.col;
    values.sort((a, b) => (ascending ? 1 : -1) * ((Number(a[key]) || 0) - (Number(b[key]) || 0)));
    this.sheet.sortedFrom = this.row;
    values.forEach((line, r) => line.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v)));
    return this;
  };
  sheetProto.setName = function (name) { this.name = name; return this; };
  // Real Sheets refuses ranges outside the grid.
  const getRange = sheetProto.getRange;
  sheetProto.getRange = function (...args) {
    const range = getRange.apply(this, args);
    if (range.rows < 1 || range.cols < 1 || range.row + range.rows - 1 > this.maxRows || range.col + range.cols - 1 > this.maxColumns) {
      throw new Error(`The coordinates of the range are outside the dimensions of the sheet (${range.getA1Notation()}).`);
    }
    return range;
  };
}

function ruleRow(overrides = {}) {
  const base = {
    NAME: 'Low ROAS', EMAILS: 'a@example.com, b@example.com', FLAG: 'Yes', SEND_EMAIL: 'Yes', EMAIL_PREFIX: '[Alert]',
    TAB_NAME: 'Alerts', N: 7, INCLUDE_CAMPAIGN: 'Yes', INCLUDE_AD_GROUP: 'No',
    CAMPAIGN_NAME_CONTAINS: '', CAMPAIGN_NAME_NOT_CONTAINS: '', ITEM_ID_CONTAINS: '', ITEM_ID_NOT_CONTAINS: '',
    PRODUCT_TITLE_CONTAINS: '', PRODUCT_TITLE_NOT_CONTAINS: '',
    FILTER_1_METRIC: 'Clicks', FILTER_1_OPERATOR: '>', FILTER_1_VALUE: 0,
    FILTER_2_METRIC: 'Roas', FILTER_2_OPERATOR: '<', FILTER_2_VALUE: 2,
  };
  return { ...base, ...overrides };
}

function makeAccount(rules = [ruleRow()]) {
  const ads = new ShoppingAds({ accountName: 'Shop' });
  addSheetSupport(ads);
  const control = ads.spreadsheet.insertSheet('Settings');
  control.maxColumns = 40;
  HEADER.forEach((key, i) => control.getRange(3, i + 1).setValue(i + 1 === LOGS_COL ? 'Logs' : key));
  control.getRange(3, LOGS_COL + 1).setValue('Last run');
  rules.forEach((rule, r) => HEADER.forEach((key, i) => {
    if (key in rule) control.getRange(4 + r, i + 1).setValue(rule[key]);
  }));
  return { ads, control };
}

function run(ads) {
  const logs = [];
  const globals = ads.globals({ Logger: { log: (m) => logs.push(String(m)) } });
  const loaded = loadScript(SCRIPT, globals, ['main'], { replace: [['YOUR_SPREADSHEET_URL_HERE', SHEET_URL]] });
  loaded.exports.main();
  return { ...loaded, loggerLogs: logs };
}

function outputRows(ads, tab = 'Alerts') {
  return ads.spreadsheet.getSheetByName(tab).rowsFrom(4).filter(r => r[0] !== '');
}

function noBadNumbers(rows) {
  for (const line of rows) for (const v of line) {
    assert.ok(!/NaN|Infinity/.test(String(v)), `bad value ${v} in ${JSON.stringify(line)}`);
  }
}

function seedProducts(ads) {
  ads.addProduct({ itemId: 'low1', title: 'Poor seller', date: daysAgo(1), clicks: 20, cost: 30, value: 20 });   // ROAS 0.67, yesterday
  ads.addProduct({ itemId: 'low2', title: 'Weak seller', date: daysAgo(7), clicks: 10, cost: 10, value: 5 });    // ROAS 0.5, first day
  ads.addProduct({ itemId: 'good', title: 'Star seller', date: daysAgo(3), clicks: 50, cost: 10, value: 100 });  // ROAS 10
  ads.addProduct({ itemId: 'old', title: 'Old seller', date: daysAgo(20), clicks: 5, cost: 5, value: 0 });       // outside 7 days
  ads.addProduct({ itemId: 'today', title: 'Today only', date: daysAgo(0), clicks: 5, cost: 5, value: 0 });      // partial day
}

test('normal run: low-ROAS products from the full lookback (incl. yesterday) written, sorted by cost, emailed', () => {
  const { ads, control } = makeAccount();
  seedProducts(ads);
  run(ads);
  const rows = outputRows(ads);
  assert.deepEqual(rows.map(r => r[3]), ['low1', 'low2']);       // sorted by cost, desc
  assert.equal(ads.spreadsheet.getSheetByName('Alerts').getRange(3, 9).getValue(), 'Cost'); // header kept in place
  assert.equal(rows[0][8], 30);
  assert.equal(rows[0][5], 'Shopping UK');
  noBadNumbers(rows);
  assert.equal(ads.emails.length, 2);
  assert.equal(ads.emails[0][0].to, 'a@example.com');
  assert.match(ads.emails[0][0].subject, /^\[Alert\] Shop - Product Performance Alerts/);
  assert.match(String(control.getRange(4, LOGS_COL).getValue()), /ran successfully/);
  const q = ads.sentQueries[0];
  assert.ok(q.includes(`segments.date BETWEEN '${daysAgo(7)}' AND '${daysAgo(1)}'`), q);
});

test('a Cost filter is sent as cost_micros, not the non-existent metrics.cost', () => {
  const { ads } = makeAccount([ruleRow({ FILTER_2_METRIC: 'Cost', FILTER_2_OPERATOR: '>', FILTER_2_VALUE: 15 })]);
  seedProducts(ads);
  run(ads);
  assert.ok(ads.sentQueries[0].includes('AND metrics.cost_micros > 15000000'), ads.sentQueries[0]);
  assert.deepEqual(outputRows(ads).map(r => r[3]), ['low1']);
});

test('when nothing matches, old results are cleared and no email is sent', () => {
  const { ads } = makeAccount();
  seedProducts(ads);
  run(ads);
  assert.equal(outputRows(ads).length, 2);
  ads.emails.length = 0;
  ads.rows.keyword_view = ads.rows.keyword_view.filter(r => r['segments.product_item_id'] === 'good');
  run(ads);
  assert.equal(outputRows(ads).length, 0);
  assert.equal(ads.spreadsheet.getSheetByName('Alerts').getRange('A2').getValue(), 'No data found');
  assert.equal(ads.emails.length, 0);
});

test('the output sheet does not grow on every run', () => {
  const { ads } = makeAccount();
  seedProducts(ads);
  run(ads);
  const sheet = ads.spreadsheet.getSheetByName('Alerts');
  const rowsAfterFirst = sheet.getMaxRows();
  run(ads); run(ads);
  assert.equal(sheet.getMaxRows(), rowsAfterFirst);
});

test('a short output sheet gets the rows it needs', () => {
  const { ads } = makeAccount();
  for (let i = 0; i < 8; i++) ads.addProduct({ itemId: `p${i}`, title: 'Poor', date: daysAgo(2), clicks: 1, cost: i + 1, value: 0 });
  ads.spreadsheet.insertSheet('Alerts').maxRows = 5;
  run(ads);
  assert.equal(outputRows(ads).length, 8);
});

test("filters with an apostrophe (Men's) build a valid query", () => {
  const { ads } = makeAccount([ruleRow({ PRODUCT_TITLE_CONTAINS: "Men's", CAMPAIGN_NAME_NOT_CONTAINS: "Kid's" })]);
  ads.addProduct({ itemId: 'm1', title: "Men's jacket", date: daysAgo(2), clicks: 3, cost: 9, value: 1 });
  ads.addProduct({ itemId: 'w1', title: 'Mens jacket', date: daysAgo(2), clicks: 3, cost: 9, value: 1 });
  ads.addProduct({ itemId: 'k1', title: "Men's jacket", campaign: "Kid's", date: daysAgo(2), clicks: 3, cost: 9, value: 1 });
  run(ads);
  assert.deepEqual(outputRows(ads).map(r => r[3]), ['m1']);
});

test('a number in a list cell (item ID 12345) and a lower-case "yes" flag do not stop the script', () => {
  const { ads } = makeAccount([ruleRow({ ITEM_ID_CONTAINS: 12345, FLAG: 'yes' })]);
  ads.addProduct({ itemId: 'sku-12345', title: 'Numbered', date: daysAgo(2), clicks: 3, cost: 9, value: 1 });
  ads.addProduct({ itemId: 'sku-999', title: 'Other', date: daysAgo(2), clicks: 3, cost: 9, value: 1 });
  run(ads);
  assert.deepEqual(outputRows(ads).map(r => r[3]), ['sku-12345']);
});

test('a blank lookback window is reported on the settings row', () => {
  const { ads, control } = makeAccount([ruleRow({ N: '' })]);
  seedProducts(ads);
  run(ads);
  assert.match(String(control.getRange(4, LOGS_COL).getValue()), /Please set a lookback window/);
  assert.equal(ads.sentQueries.length, 0);
});

test('a failing report stops only its own rule', () => {
  const { ads, control } = makeAccount([ruleRow({ NAME: 'Broken', ITEM_ID_CONTAINS: 'boom' }), ruleRow({ NAME: 'Fine', TAB_NAME: 'Fine' })]);
  seedProducts(ads);
  ads.failQuery('boom', 'Internal error');
  const { errors } = run(ads);
  assert.ok(errors.some(e => e.includes('Internal error')));
  assert.equal(outputRows(ads, 'Fine').length, 2);
  assert.match(String(control.getRange(5, LOGS_COL).getValue()), /ran successfully/);
  assert.match(String(control.getRange(4, LOGS_COL).getValue()), /still running or didn't finish/);
});

test('zero cost and zero conversions give 0 for ROAS, CPA and COS, never NaN', () => {
  const { ads } = makeAccount([ruleRow({ FILTER_2_METRIC: '', FILTER_2_OPERATOR: '', FILTER_2_VALUE: '' })]);
  ads.addProduct({ itemId: 'free', title: 'Free', date: daysAgo(2), clicks: 4, cost: 0, value: 0 });
  run(ads);
  const rows = outputRows(ads);
  assert.equal(rows.length, 1);
  assert.deepEqual([rows[0][10], rows[0][11], rows[0][12]], [0, 0, 0]);
  noBadNumbers(rows);
});

test('empty account: no rows, no email, run still marked successful', () => {
  const { ads, control } = makeAccount();
  const { errors } = run(ads);
  assert.deepEqual(errors, []);
  assert.equal(ads.emails.length, 0);
  assert.match(String(control.getRange(4, LOGS_COL).getValue()), /ran successfully/);
});
