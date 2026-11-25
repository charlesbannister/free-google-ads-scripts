'use strict';
// End-to-end runs of the whole script (main) against the mocked Google Ads Scripts runtime in
// ../test-support. Run with: node --test keyword_discovery/
const assert = require('node:assert/strict');
const { join } = require('node:path');
const { test } = require('node:test');
const { MockGoogleAds, MockSheet, loadScript, iterator } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = process.env.SCRIPT_UNDER_TEST || join(__dirname, 'keyword_discovery.js');
const SHEET = 'demo-sheet';
const URL = `https://docs.google.com/spreadsheets/d/${SHEET}/edit`;

// ---- Local additions to the shared mock (see the report for what the shared mock should gain).
// Google Sheets refuses ranges outside the sheet's rows and columns, for reads too.
const baseGetRange = MockSheet.prototype.getRange;
MockSheet.prototype.getRange = function (...args) {
  const range = baseGetRange.apply(this, args);
  if (range.row < 1 || range.row + range.rows - 1 > this.maxRows || range.col + range.cols - 1 > this.maxColumns) {
    throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
  }
  range.setDataValidation = rule => {
    if (!rule || rule.__built !== true) throw new Error('Cannot find method setDataValidation(DataValidationBuilder).');
    this.validations = (this.validations || []).concat([[range.row, range.col, range.rows, rule.kind]]);
    return range;
  };
  range.sort = ({ column, ascending }) => {
    const rows = range.getValues();
    rows.sort((a, b) => (Number(a[column - range.col]) - Number(b[column - range.col])) * (ascending ? 1 : -1));
    range.setValues(rows);
    return range;
  };
  return range;
};
// Inserting rows moves the rows below down, as in Google Sheets.
MockSheet.prototype.insertRowsAfter = function (after, n) {
  const moved = new Map();
  for (const [k, v] of this.cells) {
    const [r, c] = k.split(':').map(Number);
    moved.set(r > after ? this._key(r + n, c) : k, v);
  }
  this.cells = moved; this.maxRows += n; return this;
};
const dataValidationBuilder = () => {
  const rule = { kind: null };
  const builder = {
    requireCheckbox() { rule.kind = 'checkbox'; return builder; },
    requireValueInList() { rule.kind = 'list'; return builder; },
    setAllowInvalid() { return builder; },
    build() { return { ...rule, __built: true }; },
  };
  return builder;
};

// AWQL reports (SEARCH_QUERY_PERFORMANCE_REPORT, KEYWORDS_PERFORMANCE_REPORT), which the shared mock does not parse.
class AwqlAccount extends MockGoogleAds {
  report(query) {
    const q = String(query).replace(/\s+/g, ' ').trim();
    const m = q.match(/^select (.+?) from (\w+)(?: where (.+?))? during (\S+)$/i);
    if (!/^select .+ from (SEARCH_QUERY_PERFORMANCE_REPORT|KEYWORDS_PERFORMANCE_REPORT)\b/i.test(q)) return super.report(query);
    this.queries.push(q);
    for (const f of this.failures) if (q.includes(f.match)) throw new Error(f.message);
    if (!m) throw new Error(`Invalid AWQL query: ${q}`);
    const fields = m[1].split(',').map(f => f.trim());
    const conditions = (m[3] || '').split(/\s+and\s+/i).filter(Boolean).map(c => {
      const p = c.trim().match(/^(\w+) (=|!=|>=|<=|>|<|CONTAINS_IGNORE_CASE|DOES_NOT_CONTAIN_IGNORE_CASE|IN|NOT_IN) (.+)$/);
      if (!p) throw new Error(`Invalid AWQL condition: ${c}`);
      return { field: p[1], op: p[2], value: p[3] };
    });
    const rows = (this.rows[m[2]] || []).filter(row => conditions.every(({ field, op, value }) => {
      const n = Number(value);
      if (!['>', '<', '>=', '<='].includes(op) || isNaN(n)) return true;
      const v = Number(String(row[field]).replace(/,/g, ''));
      return { '>': v > n, '<': v < n, '>=': v >= n, '<=': v <= n }[op];
    }));
    return { rows: () => iterator(rows.map(r => Object.fromEntries(fields.map(f => [f, r[f] ?? ''])))) };
  }
  addQuery({ query, clicks, impressions = 100, cost, conversions = 0, value = 0 }) {
    return this.addRows('SEARCH_QUERY_PERFORMANCE_REPORT', [{
      AdGroupId: '10', CampaignId: '100', CampaignName: 'Search', AdGroupName: 'Widgets', Query: query,
      ConversionValue: value.toLocaleString('en-US'), Impressions: String(impressions), Clicks: String(clicks),
      Cost: cost.toLocaleString('en-US'), Conversions: String(conversions), TopImpressionPercentage: '0.5',
      AbsoluteTopImpressionPercentage: '0.2', QueryTargetingStatus: 'NONE',
    }]);
  }
  addKeyword(text) {
    return this.addRows('KEYWORDS_PERFORMANCE_REPORT', [{ CampaignId: '100', AdGroupId: '10', Criteria: text, KeywordMatchType: 'Exact', Impressions: '50' }]);
  }
  globals() {
    const g = super.globals();
    g.AdWordsApp = g.AdsApp;
    g.SpreadsheetApp.newDataValidation = dataValidationBuilder;
    g.Logger = { log: () => {} };
    return g;
  }
}

const HEADER = ['NAME', 'EMAILS', 'FLAG', 'EMAIL_ALERT', 'TAB_NAME', 'N', 'SEARCH_TERM_CONTAINS', 'SEARCH_TERM_NOT_CONTAINS', 'FUZZY_THRESHOLD'];
for (let i = 1; i <= 6; i++) HEADER.push(`FILTER_${i}_METRIC`, `FILTER_${i}_OPERATOR`, `FILTER_${i}_VALUE`);
HEADER.push('LOG_SHEET_URL', 'Logs', 'Last Run');
const OUTPUT_HEADER = ['Search Term', 'Status', 'Ignore', 'Impressions', 'Clicks', 'Ctr', 'Cost', 'Conversions', 'Cpa', 'Conv. Rate', 'Conv. Value', 'Roas'];

function account({ rule = {}, filters = [['Clicks', '>', 5]], preview = false, outputRows = 6, previous = [], ignored = [] } = {}) {
  const ads = new AwqlAccount({ spreadsheetId: SHEET, preview });
  const settings = new MockSheet('Settings', 1, { maxColumns: 40 });
  const output = new MockSheet('Opportunities', 2, { maxRows: outputRows });
  const ignoredSheet = new MockSheet('Ignored Terms', 3);
  ads.spreadsheet.sheets.push(settings, output, ignoredSheet);
  settings.getRange(3, 1, 1, HEADER.length).setValues([HEADER]);
  const values = {
    NAME: 'Winners', EMAILS: 'me@example.com', FLAG: 'Yes', EMAIL_ALERT: 'No', TAB_NAME: 'Opportunities', N: 30,
    SEARCH_TERM_CONTAINS: '', SEARCH_TERM_NOT_CONTAINS: '', FUZZY_THRESHOLD: 0.8, LOG_SHEET_URL: URL, ...rule,
  };
  filters.forEach(([m, o, v], i) => Object.assign(values, { [`FILTER_${i + 1}_METRIC`]: m, [`FILTER_${i + 1}_OPERATOR`]: o, [`FILTER_${i + 1}_VALUE`]: v }));
  settings.getRange(4, 1, 1, HEADER.length - 2).setValues([HEADER.slice(0, -2).map(h => values[h] ?? '')]);
  output.getRange(1, 1).setValue('Keyword Discovery');
  output.getRange(3, 1, 1, OUTPUT_HEADER.length).setValues([OUTPUT_HEADER]);
  previous.forEach(([term, ignore], i) => output.getRange(4 + i, 1, 1, 3).setValues([[term, 'NONE', ignore]]));
  ignoredSheet.getRange(4, 1, 1, 2).setValues([['Search Term', 'Date Added']]);
  ignored.forEach((term, i) => ignoredSheet.getRange(5 + i, 1, 1, 2).setValues([[term, 'Nov 01, 2026']]));

  ads.addKeyword('blue widgets');
  ads.addQuery({ query: 'buy red widgets', clicks: 30, cost: 12.5, conversions: 3, value: 90 });
  ads.addQuery({ query: 'blue widget', clicks: 40, cost: 20 }); // close to the keyword "blue widgets"
  ads.addQuery({ query: 'green gadgets', clicks: 9, cost: 1234.5 });
  ads.addQuery({ query: 'yellow gizmo', clicks: 7, cost: 3 });
  ads.addQuery({ query: 'purple thing', clicks: 6, cost: 8 });
  ads.addQuery({ query: 'orange doohickey', clicks: 8, cost: 0.5 });
  ads.addQuery({ query: 'free widgets', clicks: 12, cost: 4 });
  ads.addQuery({ query: 'few clicks term', clicks: 2, cost: 9 });
  return ads;
}

function run(ads) {
  const script = loadScript(SCRIPT, ads.globals(), ['main'], { replace: [['YOUR_SPREADSHEET_URL_HERE', URL]] });
  script.exports.main();
  return script;
}

const outputRows = ads => ads.spreadsheet.getSheetByName('Opportunities').rowsFrom(4).filter(r => r[0] !== '');
const controlLog = ads => ads.spreadsheet.getSheetByName('Settings').getRange(4, HEADER.indexOf('Logs') + 1).getValue();

test('normal run: new opportunities written below the header, sorted by cost, with Ignore checkboxes', () => {
  const ads = account({ outputRows: 6, previous: [['old term', false]] });
  run(ads);
  const rows = outputRows(ads);
  assert.deepEqual(rows.map(r => [r[0], r[6]]), [
    ['green gadgets', 1234.5], ['buy red widgets', 12.5], ['purple thing', 8], ['free widgets', 4],
    ['yellow gizmo', 3], ['orange doohickey', 0.5],
  ]);
  assert.equal(rows[0].length, 12);
  assert.equal(rows[1][2], false);
  const out = ads.spreadsheet.getSheetByName('Opportunities');
  assert.deepEqual(out.validations, [[4, 3, 6, 'checkbox']]);
  assert.equal(out.getRange(3, 1).getValue(), 'Search Term'); // header kept
  // Both AWQL queries were valid (the keyword query used to read "...> 0DURING LAST_30_DAYS").
  assert.ok(ads.queries.some(q => /KEYWORDS_PERFORMANCE_REPORT .* DURING LAST_30_DAYS$/.test(q)));
  assert.match(controlLog(ads), /The script ran successfully/);
});

test('ticked Ignore terms go to Ignored Terms and are not reported again, nor are terms already ignored', () => {
  const ads = account({ outputRows: 20, previous: [['free widgets', true], ['old term', false]], ignored: ['Purple Thing'] });
  run(ads);
  const terms = outputRows(ads).map(r => r[0]);
  assert.ok(!terms.includes('free widgets'));
  assert.ok(!terms.includes('purple thing'));
  assert.ok(terms.includes('buy red widgets'));
  const ignoredSheet = ads.spreadsheet.getSheetByName('Ignored Terms');
  assert.ok(ignoredSheet.rowsFrom(5).some(r => r[0] === 'free widgets'));
  assert.match(controlLog(ads), /1 search terms added to the ignore list/);
});

test('email alert is sent, but not in preview mode', () => {
  const live = account({ rule: { EMAIL_ALERT: 'Yes' } });
  run(live);
  assert.equal(live.emails.length, 1);
  assert.equal(live.emails[0][0].to, 'me@example.com');

  const preview = account({ rule: { EMAIL_ALERT: 'Yes' }, preview: true });
  run(preview);
  assert.equal(preview.emails.length, 0);
  assert.match(controlLog(preview), /preview mode/);
});

test('empty account: old results cleared, nothing written, run succeeds', () => {
  const ads = account({ previous: [['old term', false], ['older term', false]] });
  ads.rows.SEARCH_QUERY_PERFORMANCE_REPORT = [];
  run(ads);
  assert.deepEqual(outputRows(ads), []);
  assert.match(controlLog(ads), /The script ran successfully/);
});

test('bad sheet input: an invalid fuzzy threshold is written to Logs and stops the row', () => {
  const ads = account({ rule: { FUZZY_THRESHOLD: 2 } });
  assert.throws(() => run(ads));
  assert.equal(controlLog(ads), 'Please check the fuzzy threshold');
});

test('bad sheet input: a missing output tab stops the run with a clear message', () => {
  const ads = account({ rule: { TAB_NAME: 'Nope' } });
  assert.throws(() => run(ads), /The sheet Nope does not exist/);
});

test('report failure stops the run', () => {
  const ads = account();
  ads.failQuery('SEARCH_QUERY_PERFORMANCE_REPORT', 'Internal error');
  assert.throws(() => run(ads), /Internal error/);
});
