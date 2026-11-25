'use strict';
// End-to-end tests for google_shopping_trends.js: main() runs against a mocked account and sheet.
// Run: node --test google_shopping_trends
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const { MockGoogleAds, loadScript } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = process.env.TRENDS_SCRIPT || path.join(__dirname, 'google_shopping_trends.js');
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/demo-sheet/edit';
const START_ROW = 17;

function daysAgo(n) {
  const d = new Date();
  d.setUTCDate(d.getUTCDate() - n);
  return d.toISOString().slice(0, 10);
}

// The shared mock has no shopping_performance_view and compares dates as numbers only,
// so this account maps the resource onto keyword_view (same allowed prefixes) and turns
// `segments.date >= 'a' and segments.date <= 'b'` into BETWEEN.
class ShoppingAds extends MockGoogleAds {
  constructor(options) { super(options); this.sentQueries = []; }
  addProduct({ itemId, title, campaign = 'Shopping UK', date, clicks = 0, impressions = 1, cost = 0, conversions = 0, value = 0 }) {
    return this.addRows('keyword_view', [{
      'segments.product_item_id': itemId, 'segments.product_title': title, 'campaign.name': campaign,
      'segments.date': date, 'metrics.clicks': String(clicks), 'metrics.impressions': String(impressions),
      'metrics.cost_micros': String(Math.round(cost * 1e6)), 'metrics.conversions': String(conversions),
      'metrics.conversions_value': String(value),
    }]);
  }
  report(query) {
    this.sentQueries.push(query);
    const q = String(query)
      .replace(/from shopping_performance_view/i, 'from keyword_view')
      .replace(/segments\.date >= ('[\d-]+') and segments\.date <= ('[\d-]+')/i, 'segments.date BETWEEN $1 AND $2');
    return super.report(q);
  }
}

function addSheetSupport(ads) {
  const proto = Object.getPrototypeOf(ads.spreadsheet.insertSheet('__probe').getRange(1, 1));
  ads.spreadsheet.sheets = ads.spreadsheet.sheets.filter(s => s.name !== '__probe');
  proto.getFilter = function () { return this.sheet._filter; };
  proto.clear = proto.clearContent;
  proto.sort = function ({ column, ascending }) {
    if (this.row + this.rows - 1 > this.sheet.maxRows) throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
    const values = this.getValues();
    const key = column - this.col;
    values.sort((a, b) => {
      const x = a[key]; const y = b[key];
      if (x === '' && y !== '') return 1;
      if (y === '' && x !== '') return -1;
      const d = Number(x) - Number(y);
      return ascending ? d : -d;
    });
    this.sheet.sorted = { row: this.row, column };
    values.forEach((line, r) => line.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v)));
    return this;
  };
}

function makeAccount({ windows, filters = {}, sortColumn = 4, maxRows = 1000 } = {}) {
  const ads = new ShoppingAds();
  const notes = ads.spreadsheet.insertSheet('Notes');
  const report = ads.spreadsheet.insertSheet('UK');
  report.maxRows = maxRows;
  report.maxColumns = 40; // the template has the 39 report columns
  addSheetSupport(ads);
  const outside = (sheet) => {
    const original = sheet.getRange.bind(sheet);
    sheet.getRange = (...args) => {
      const range = original(...args);
      if (range.row + range.rows - 1 > sheet.maxRows || range.col + range.cols - 1 > sheet.maxColumns || range.rows < 1 || range.cols < 1) {
        throw new Error(`The coordinates of the range are outside the dimensions of the sheet (${range.getA1Notation()}).`);
      }
      return range;
    };
  };
  outside(report);
  notes.getRange('A1').setValue('Notes');
  report.getRange('A1').setValue('Last updated');
  const cells = {
    A7: filters.itemContains ?? '', B7: filters.itemNotContains ?? '',
    A10: filters.titleContains ?? '', B10: filters.titleNotContains ?? '',
    A13: filters.campaignContains ?? '', B13: filters.campaignNotContains ?? '',
    G7: sortColumn,
  };
  const rows = ['D7', 'D8', 'D9', 'D10', 'D11', 'D12'];
  (windows || [[7, 'Week'], [0, 'Today'], [30, 'Month'], [1, 'Yesterday'], [14, 'Fortnight'], [60, 'Two months']])
    .forEach(([days, name], i) => { cells[rows[i]] = days; cells[rows[i].replace('D', 'E')] = name; });
  for (const [a1, v] of Object.entries(cells)) report.getRange(a1).setValue(v);
  return { ads, report, notes };
}

function run(ads) {
  const loaded = loadScript(SCRIPT, ads.globals(), ['main'], { replace: [['YOUR_SPREADSHEET_URL_HERE', SHEET_URL]] });
  loaded.exports.main();
  return loaded;
}

function dataRows(sheet) {
  return sheet.rowsFrom(START_ROW + 2).filter(r => r[0] !== '');
}

function noBadNumbers(sheet) {
  for (const line of sheet.rowsFrom(1)) {
    for (const v of line) {
      assert.ok(!(typeof v === 'number' && !Number.isFinite(v)), `bad number ${v} in ${JSON.stringify(line)}`);
      assert.ok(!/NaN|Infinity/.test(String(v)), `bad text ${v}`);
    }
  }
}

test('normal run: per-day averages per window, Total row, sorted, timestamp and version', () => {
  const { ads, report, notes } = makeAccount();
  ads.addProduct({ itemId: 'sku1', title: 'Red jacket', date: daysAgo(1), clicks: 14, impressions: 140, cost: 7, value: 70 });
  ads.addProduct({ itemId: 'sku1', title: 'Red jacket', date: daysAgo(0), clicks: 100, impressions: 1000, cost: 50 });
  ads.addProduct({ itemId: 'sku1', title: 'Red jacket', date: daysAgo(20), clicks: 16, impressions: 160, cost: 8 });
  ads.addProduct({ itemId: 'sku2', title: 'Blue socks', date: daysAgo(3), clicks: 70, impressions: 700, cost: 14, value: 28 });
  run(ads);

  const header = report.rowsFrom(START_ROW + 1)[0];
  assert.deepEqual(header.slice(0, 9), ['Item ID', 'Title', 'Campaigns', 'Week', 'Today', 'Month', 'Yesterday', 'Fortnight', 'Two months']);
  const rows = dataRows(report);
  const byId = Object.fromEntries(rows.map(r => [r[0], r]));
  // Clicks per day: Week (7d to yesterday), Today, Month (30d)
  assert.equal(byId.sku1[3], 2);              // 14 / 7
  assert.equal(byId.sku1[4], 100);            // today, not averaged
  assert.equal(byId.sku1[5], 1);              // (14 + 16) / 30
  assert.equal(byId.sku2[3], 10);             // 70 / 7
  assert.equal(byId.total[3], 12);            // (14 + 70) / 7
  // Sorted descending on column 4 (the first window's clicks)
  assert.deepEqual(rows.map(r => r[0]), ['total', 'sku2', 'sku1']);
  assert.equal(Object.prototype.toString.call(report.getRange('B1').getValue()), '[object Date]');
  assert.equal(notes.getRange('B2').getValue(), '1.1.1');
  // Yesterday is the end date, today excluded, for N > 0
  const week = ads.sentQueries.find(q => q.includes(`>= '${daysAgo(7)}'`));
  assert.ok(week && week.includes(`<= '${daysAgo(1)}'`), 'week window ends yesterday');
  noBadNumbers(report);
});

test('blank lookback windows do not write NaN or Infinity', () => {
  const { ads, report } = makeAccount({ windows: [[7, 'Week'], [0, 'Today'], ['', ''], ['', ''], ['', ''], ['', '']] });
  ads.addProduct({ itemId: 'sku1', title: 'Red jacket', date: daysAgo(2), clicks: 7, cost: 7 });
  ads.addProduct({ itemId: 'sku1', title: 'Red jacket', date: daysAgo(0), clicks: 3, cost: 3 }); // today's data
  run(ads);
  const rows = dataRows(report);
  assert.equal(rows.find(r => r[0] === 'sku1')[3], 1);
  noBadNumbers(report);
  // No query is sent for the blank windows
  assert.equal(ads.sentQueries.length, 2);
});

test('empty account: headers only, no errors', () => {
  const { ads, report } = makeAccount();
  const { errors } = run(ads);
  assert.equal(dataRows(report).length, 0);
  assert.equal(report.rowsFrom(START_ROW + 1)[0][0], 'Item ID');
  assert.ok(errors.some(e => e.includes('No data found')));
});

test('zero cost and zero value: COS is 0, no divide by zero', () => {
  const { ads, report } = makeAccount();
  ads.addProduct({ itemId: 'free', title: 'Free sample', date: daysAgo(2), clicks: 5, cost: 0, value: 0 });
  run(ads);
  noBadNumbers(report);
});

test('second run clears the old rows and fits a sheet with no spare rows', () => {
  const { ads, report } = makeAccount({ maxRows: START_ROW + 1 + 3 });
  for (let i = 0; i < 6; i++) ads.addProduct({ itemId: `sku${i}`, title: `T${i}`, date: daysAgo(2), clicks: i + 1 });
  run(ads);
  assert.equal(dataRows(report).length, 7); // 6 products + total
  ads.rows.keyword_view = ads.rows.keyword_view.slice(0, 2);
  run(ads);
  assert.deepEqual(dataRows(report).map(r => r[0]).sort(), ['sku0', 'sku1', 'total']);
});

test('filters with quotes produce a valid query', () => {
  const { ads, report } = makeAccount({ filters: { titleContains: '12" pizza', campaignNotContains: "Men's" } });
  ads.addProduct({ itemId: 'p1', title: '12" pizza tray', date: daysAgo(2), clicks: 7 });
  ads.addProduct({ itemId: 'p2', title: '10 inch tray', date: daysAgo(2), clicks: 7 });
  ads.addProduct({ itemId: 'p3', title: '12" pizza stone', campaign: "Men's stuff", date: daysAgo(2), clicks: 7 });
  run(ads);
  assert.deepEqual(dataRows(report).map(r => r[0]).sort(), ['p1', 'total']);
});

test('blank sort column does not stop the run', () => {
  const { ads, report } = makeAccount({ sortColumn: '' });
  ads.addProduct({ itemId: 'sku1', title: 'Red jacket', date: daysAgo(2), clicks: 7 });
  run(ads);
  assert.equal(Object.prototype.toString.call(report.getRange('B1').getValue()), '[object Date]');
});

test('report failure stops the run with the error and sends no email by default', () => {
  const { ads } = makeAccount();
  ads.failQuery('keyword_view', 'Internal error');
  assert.throws(() => run(ads), /Internal error/);
  assert.equal(ads.emails.length, 0);
});
