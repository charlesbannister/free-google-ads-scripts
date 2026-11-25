'use strict';
// End-to-end runs of the whole script (main) against the mocked Google Ads Scripts runtime
// in ../test-support. Run with: node --test search_term_trends/
const assert = require('node:assert/strict');
const { join } = require('node:path');
const { test } = require('node:test');
const { MockGoogleAds, MockSheet, loadScript } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = join(__dirname, 'search_term_trends.js');
const SHEET = 'demo-sheet';
const START_ROW = 17;

// The shared mock has no named ranges, Range.sort, Range.clear or Range.getFilter, and does
// not refuse ranges outside the sheet the way Google does. Added here for this file only.
const NAMES = ['search_term_contains', 'search_term_not_contains', 'campaign_contains', 'campaign_not_contains',
  'impressions_filter', 'clicks_filter', 'date_1_days', 'date_1_name', 'date_2_days', 'date_2_name',
  'date_3_days', 'date_3_name', 'date_4_days', 'date_4_name', 'date_5_days', 'date_5_name',
  'date_6_days', 'date_6_name', 'sort_by_column_number', 'last_updated_timestamp'];
(function extendMock() {
  const proto = MockSheet.prototype;
  const getRange = proto.getRange;
  proto.getRange = function (a, b, c, d) {
    if (typeof a === 'string' && NAMES.includes(a)) { const i = NAMES.indexOf(a); return getRange.call(this, (i % 10) + 1, 2 + 2 * Math.floor(i / 10), 1, 1); } // rows 1-10, above START_ROW
    const range = getRange.call(this, a, b, c, d);
    if (range.rows < 1 || range.cols < 1 || range.row + range.rows - 1 > this.maxRows || range.col + range.cols - 1 > this.maxColumns) {
      throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
    }
    const rp = Object.getPrototypeOf(range);
    if (!rp.sort) {
      rp.getFilter = function () { return this.sheet._filter; };
      rp.clear = function () { return this.clearContent(); };
      rp.sort = function ({ column, ascending }) {
        const values = this.getValues();
        values.sort((x, y) => {
          const p = x[column - this.col]; const q = y[column - this.col];
          const pn = typeof p === 'number'; const qn = typeof q === 'number';
          let cmp = pn && qn ? p - q : pn ? -1 : qn ? 1 : String(p).localeCompare(String(q));
          return ascending === false ? -cmp : cmp;
        });
        this.clearContent();
        values.forEach((line, r) => line.forEach((v, i) => this.sheet._set(this.row + r, this.col + i, v)));
        return this;
      };
    }
    return range;
  };
})();

function settingsTab(ads, name, overrides = {}) {
  const settings = {
    search_term_contains: '', search_term_not_contains: '', campaign_contains: '', campaign_not_contains: '',
    impressions_filter: 0, clicks_filter: '',
    date_1_days: 30, date_1_name: '30 days', date_2_days: 7, date_2_name: '7 days', date_3_days: 1, date_3_name: 'Yesterday',
    date_4_days: 14, date_4_name: '14 days', date_5_days: 60, date_5_name: '60 days', date_6_days: 0, date_6_name: 'Today',
    sort_by_column_number: 3, ...overrides,
  };
  const sheet = ads.spreadsheet.insertSheet(name);
  for (const [key, value] of Object.entries(settings)) sheet.getRange(key).setValue(value);
  return sheet;
}

function account() {
  const ads = new MockGoogleAds({ spreadsheetId: SHEET });
  ads.spreadsheet.insertSheet('Notes');
  ads.addCampaign({ id: '100', name: 'Shoes', channel: 'SEARCH' });
  ads.addCampaign({ id: '101', name: 'Brand "Main"', channel: 'SEARCH' });
  ads.addAdGroup({ id: '200', name: 'A', campaignId: '100' });
  ads.addAdGroup({ id: '201', name: 'B', campaignId: '100' });
  ads.addAdGroup({ id: '202', name: 'C', campaignId: '101' });
  // One term in two ad groups of one campaign and in a second campaign: three rows.
  ads.addSearchTerm({ term: 'red shoes', campaignId: '100', adGroupId: '200', impressions: 300, clicks: 30, cost: 60, conversions: 1.5, value: 90 });
  ads.addSearchTerm({ term: 'red shoes', campaignId: '100', adGroupId: '201', impressions: 30, clicks: 3, cost: 6, conversions: 2, value: 30 });
  ads.addSearchTerm({ term: 'red shoes', campaignId: '101', adGroupId: '202', impressions: 5, clicks: 0, cost: 0, conversions: 0, value: 0 });
  ads.addSearchTerm({ term: 'zero shoes', campaignId: '100', adGroupId: '200', impressions: 0 });
  ads.addSearchTerm({ term: 'blue shoes', campaignId: '100', adGroupId: '200', impressions: 60, clicks: 6, cost: 3, conversions: 0, value: 0 });
  return ads;
}

function run(ads) {
  const script = loadScript(SCRIPT, ads.globals(), ['main'], {
    replace: [['YOUR_SPREADSHEET_URL_HERE', `https://docs.google.com/spreadsheets/d/${SHEET}/edit`]],
  });
  script.exports.main();
  return script;
}

const rows = sheet => Object.fromEntries(sheet.rowsFrom(START_ROW + 2).map(r => [r[0], r]));

test('per-day averages add up every ad group row and list each campaign once', () => {
  const ads = account();
  const sheet = settingsTab(ads, 'UK');
  run(ads);
  const header = sheet.rowsFrom(START_ROW + 1)[0];
  assert.deepEqual(header.slice(0, 4), ['Search Term', 'Campaigns', '30 days', '7 days']);
  const red = rows(sheet)['red shoes'];
  assert.equal(red[1], 'Shoes, Brand "Main"');         // Shoes once, though it has two ad group rows
  assert.ok(!rows(sheet)['zero shoes']);                // impressions > 0
  assert.equal(red[2], Math.round(33 / 30 * 100) / 100); // clicks per day, 30 days
  const conv = red[2 + 2 * 6];                          // conversions per day, first window
  assert.equal(conv, Math.round(3.5 / 30 * 100) / 100);
  const value = red[2 + 3 * 6];
  assert.equal(value, 4);                               // 120 / 30
  assert.equal(red[2 + 5 * 6], '55%');                  // COS: 66 / 120
  assert.equal(red[2 + 5 * 6 + 5], '55%');              // today window has the same rows in the mock
  for (const line of sheet.rowsFrom(START_ROW + 2)) for (const v of line) if (typeof v === 'number') assert.ok(Number.isFinite(v));
  assert.equal(Object.prototype.toString.call(sheet.getRange('last_updated_timestamp').getValue()), '[object Date]');
  assert.equal(ads.spreadsheet.getSheetByName('Notes').getLastRow(), 0);
  // Sorted by column 3 (clicks per day), highest first.
  assert.deepEqual(sheet.rowsFrom(START_ROW + 2).map(r => r[0]), ['red shoes', 'blue shoes']);
});

test('date windows: N days ends yesterday, 0 is today', () => {
  const ads = account();
  settingsTab(ads, 'UK');
  run(ads);
  const dates = ads.queries.map(q => q.match(/segments\.date >= '([\d-]+)' and segments\.date <= '([\d-]+)'/).slice(1));
  const span = ([from, to]) => (Date.parse(to) - Date.parse(from)) / 864e5 + 1;
  assert.deepEqual(dates.map(span), [30, 7, 1, 14, 60, 1]);
  assert.equal(dates[5][0], dates[5][1]);
});

test('a second run over a full sheet does not go outside the sheet', () => {
  const ads = account();
  const sheet = settingsTab(ads, 'UK');
  for (let i = 0; i < 1100; i++) ads.addSearchTerm({ term: `term ${i}`, campaignId: '100', adGroupId: '200', impressions: 5, clicks: 1, cost: 1 });
  run(ads);
  assert.equal(sheet.rowsFrom(START_ROW + 2).length, 1102);
  assert.ok(sheet.getMaxRows() >= START_ROW + 1103);
  run(ads);
  assert.equal(sheet.rowsFrom(START_ROW + 2).length, 1102);
});

test('bad sheet input: quotes in filters, blank sort column, text filters', () => {
  const ads = account();
  const sheet = settingsTab(ads, 'UK', { campaign_contains: 'Brand "Main', clicks_filter: 'abc', sort_by_column_number: '', impressions_filter: '' });
  const { warnings } = run(ads);
  assert.ok(ads.queries.every(q => q.includes('campaign.name like "%Brand \\"Main%"')));
  assert.ok(ads.queries.every(q => !q.includes('abc')));
  assert.ok(warnings.some(w => w.includes('clicks_filter')));
  assert.deepEqual(Object.keys(rows(sheet)), ['red shoes']);
});

test('a blank lookback window gives zeros, not Infinity', () => {
  const ads = account();
  const sheet = settingsTab(ads, 'UK', { date_4_days: '', date_4_name: '' });
  run(ads);
  for (const line of sheet.rowsFrom(START_ROW + 2)) for (const v of line) if (typeof v === 'number') assert.ok(Number.isFinite(v), String(line));
});

test('empty account writes the two header rows only', () => {
  const ads = new MockGoogleAds({ spreadsheetId: SHEET });
  const sheet = settingsTab(ads, 'UK');
  run(ads);
  assert.equal(sheet.getLastRow(), START_ROW + 1);
});

test('a report failure stops the run with the error', () => {
  const ads = account();
  settingsTab(ads, 'UK');
  ads.failQuery('search_term_view', 'Internal error');
  assert.throws(() => run(ads), /Internal error/);
});
