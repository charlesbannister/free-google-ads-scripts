'use strict';
// End-to-end runs of the whole script (main) against the mocked Google Ads Scripts runtime
// in ../test-support. Run with: node --test search_term_n_grams_standard/
//
// This script still uses the legacy AWQL report (SEARCH_QUERY_PERFORMANCE_REPORT) and the
// AdWordsApp / Logger names, which the shared mock does not imitate, so a small AWQL fake
// and the extra Sheets calls it needs are added here for this file only.
const assert = require('node:assert/strict');
const { join } = require('node:path');
const { test } = require('node:test');
const { MockGoogleAds, MockSheet, MockSpreadsheet, loadScript } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = join(__dirname, 'search_term_n_grams.js');
const CONTROL = 'control-sheet';
const OUTPUT = 'output-sheet';
const HEADER = ['ID', 'NAME', 'EMAILS', 'NOTIFY', 'FLAG', 'NEGATIVE_KEYWORD_LIST', 'NEGATIVE_KEYWORD_DEFAULT_MATCH_TYPE',
  'DONT_LOG_ADDED_NEGATIVES', 'N', 'MIN_QUERY_IMPRESSIONS', 'MIN_QUERY_CLICKS', 'CAMPAIGN_NAME_CONTAINS',
  'CAMPAIGN_NAME_NOT_CONTAINS', 'CAMPAIGN_NAME_EQUALS', 'QUERY_NOT_CONTAINS', 'MIN_CLICKS', 'MAX_CONVERSIONS', 'MIN_CPA',
  'MAX_ROAS', 'PLACEHOLDER_1', 'PLACEHOLDER_2', 'PLACEHOLDER_3', 'PLACEHOLDER_4', 'PLACEHOLDER_5', 'LOG_SHEET_URL', 'Logs'];

(function extendMock() {
  const proto = MockSheet.prototype;
  const getRange = proto.getRange;
  proto.setName = function (name) { this.name = String(name); return this; };
  proto.sort = function (column, ascending = true) { // Sheet.sort: every row (no frozen rows here)
    getRange.call(this, 1, 1, this.getLastRow(), this.getLastColumn()).sortRows(column, ascending);
    return this;
  };
  proto.getRange = function (a, b, c, d) {
    const range = getRange.call(this, a, b, c, d);
    if (range.rows < 1 || range.cols < 1 || range.row + range.rows - 1 > this.maxRows || range.col + range.cols - 1 > this.maxColumns) {
      throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
    }
    const rp = Object.getPrototypeOf(range);
    if (!rp.sortRows) {
      rp.clearDataValidations = function () { return this; };
      rp.setDataValidation = function () { return this; };
      rp.sortRows = function (column, ascending) {
        const values = this.getValues();
        values.sort((x, y) => {
          const p = x[column - this.col]; const q = y[column - this.col];
          const pn = typeof p === 'number'; const qn = typeof q === 'number';
          // Google Sheets: numbers before text ascending, text before numbers descending.
          const cmp = pn && qn ? p - q : pn ? -1 : qn ? 1 : String(p).localeCompare(String(q));
          return ascending === false ? -cmp : cmp;
        });
        values.forEach((line, r) => line.forEach((v, i) => this.sheet._set(this.row + r, this.col + i, v)));
        return this;
      };
    }
    return range;
  };
})();

function spreadsheet(id, name) {
  const ss = new MockSpreadsheet(id);
  ss.name = name; ss.editors = [];
  ss.getName = () => ss.name;
  ss.rename = n => { ss.name = n; };
  ss.getEditors = () => ss.editors.map(e => ({ getEmail: () => e }));
  ss.addEditor = e => ss.editors.push(e);
  return ss;
}

function setup({ settings = {}, queries = [], outputUrl } = {}) {
  const ads = new MockGoogleAds({ customerId: '123-456-7890', accountName: 'Shoe Shop' });
  const books = {};
  const control = books[CONTROL] = spreadsheet(CONTROL, 'Settings book');
  const output = books[OUTPUT] = spreadsheet(OUTPUT, 'Shoe Shop - nGrams Output Sheet');
  const sheet = control.insertSheet('Settings');
  sheet.maxColumns = 30; // the template has Log Sheet URL, Logs and Last Run columns (25-27)
  sheet.getRange(3, 1, 1, HEADER.length).setValues([HEADER]);
  const row = {
    ID: '123-456-7890', NAME: 'Shoe Shop', EMAILS: 'me@example.com', NOTIFY: false, FLAG: true,
    NEGATIVE_KEYWORD_LIST: 'nGram negatives', NEGATIVE_KEYWORD_DEFAULT_MATCH_TYPE: 'Phrase', DONT_LOG_ADDED_NEGATIVES: true,
    N: 30, MIN_QUERY_IMPRESSIONS: 0, MIN_QUERY_CLICKS: '', CAMPAIGN_NAME_CONTAINS: '', CAMPAIGN_NAME_NOT_CONTAINS: '',
    CAMPAIGN_NAME_EQUALS: '', QUERY_NOT_CONTAINS: '', MIN_CLICKS: '', MAX_CONVERSIONS: '', MIN_CPA: '', MAX_ROAS: '',
    LOG_SHEET_URL: outputUrl === undefined ? `https://docs.google.com/spreadsheets/d/${OUTPUT}/edit` : outputUrl,
    Logs: '', ...settings,
  };
  sheet.getRange(4, 1, 1, HEADER.length).setValues([HEADER.map(k => (k.startsWith('PLACEHOLDER') ? '' : row[k]))]);

  const awql = [];
  const lists = { 'nGram negatives': [] };
  const logs = [];
  let created = 0;
  const AdWordsApp = {
    ...ads.globals().AdsApp,
    report: (query) => {
      awql.push(query);
      if (ads.failAwql) throw new Error(ads.failAwql);
      const where = query.split(' where ')[1] || '';
      const num = re => { const m = where.match(re); return m ? Number(m[1]) : -Infinity; };
      const minImpr = num(/Impressions > (\S+)/); const minClicks = num(/Clicks > (\S+)/);
      const contains = [...where.matchAll(/CampaignName CONTAINS_IGNORE_CASE '([^']*)'/g)].map(m => m[1].toLowerCase());
      const rows = queries.filter(q => q.Impressions > minImpr && q.Clicks > minClicks &&
        contains.every(c => q.CampaignName.toLowerCase().includes(c)));
      return { rows: () => {
        let i = 0;
        // AWQL rows come back as strings, cost with thousands separators.
        const out = rows.map(q => ({ Query: q.Query, Impressions: String(q.Impressions), Clicks: String(q.Clicks),
          Cost: Number(q.Cost).toLocaleString('en-US', { minimumFractionDigits: 2 }), Conversions: String(q.Conversions),
          ConversionValue: Number(q.ConversionValue).toLocaleString('en-US', { minimumFractionDigits: 2 }) }));
        return { hasNext: () => i < out.length, next: () => out[i++] };
      } };
    },
    negativeKeywordLists: () => {
      let name = null;
      const sel = {
        withCondition: c => { name = c.match(/Name = '(.*)'/)[1]; return sel; },
        get: () => {
          const items = name in lists ? [{ addNegativeKeywords: k => lists[name].push(...k) }] : [];
          let i = 0;
          return { hasNext: () => i < items.length, next: () => items[i++] };
        },
      };
      return sel;
    },
  };
  const globals = ads.globals({
    AdWordsApp,
    Logger: { log: m => logs.push(String(m)) },
    SpreadsheetApp: {
      openByUrl: url => { const id = Object.keys(books).find(k => String(url).includes(k)); if (!id) throw new Error(`No spreadsheet ${url}`); return books[id]; },
      create: name => { const id = `created-${++created}`; books[id] = spreadsheet(id, name); return books[id]; },
      newDataValidation: () => { const b = { requireCheckbox: () => b, requireValueInList: () => b, setAllowInvalid: () => b, build: () => b }; return b; },
    },
  });
  return { ads, books, control, output, awql, lists, logs, globals };
}

function run(env) {
  const script = loadScript(SCRIPT, env.globals, ['main'], {
    replace: [['YOUR_SPREADSHEET_URL_HERE', `https://docs.google.com/spreadsheets/d/${CONTROL}/edit`]],
  });
  script.exports.main();
  return script;
}

const QUERIES = [
  { Query: 'red shoes', CampaignName: 'Shopping 2024', Impressions: 100, Clicks: 10, Cost: 1200.5, Conversions: 0, ConversionValue: 0 },
  { Query: 'shoes for shoes', CampaignName: 'Shopping 2024', Impressions: 40, Clicks: 4, Cost: 4, Conversions: 1, ConversionValue: 0 },
  { Query: 'free  shoes ', CampaignName: 'Shopping 2025', Impressions: 10, Clicks: 2, Cost: 0, Conversions: 1, ConversionValue: 5 },
];
const tab = (book, name) => Object.fromEntries(book.getSheetByName(name).rowsFrom(2).map(r => [r[1], r]));

test('nGrams add up each query once, with no blank nGrams and no Infinity', () => {
  const env = setup({ queries: QUERIES });
  run(env);
  const one = tab(env.output, '1-grams');
  assert.deepEqual(env.output.getSheetByName('1-grams').rowsFrom(1)[0], ['Date', 'nGram', 'Add as negative?', 'Match type',
    'Clicks', 'Impressions', 'Ctr', 'Cost', 'Conversions', 'Cpa', 'Conversion Value', 'ROAS']);
  // shoes: 10 + 4 (once, not twice) + 2 clicks; cost 1,200.50 + 4 + 0 with the separator removed.
  assert.deepEqual(one.shoes.slice(3, 12), ['Phrase', 16, 150, 0.1067, 1204.5, 2, 602.25, 5, 0]);
  assert.ok(!('' in one), 'double spaces give no blank nGram');
  assert.equal(one.free[11], 0, 'value with no cost: ROAS 0, not Infinity');
  assert.deepEqual(Object.keys(tab(env.output, '2-grams')).sort(), ['for shoes', 'free shoes', 'red shoes', 'shoes for']);
  for (const name of ['1-grams', '2-grams', '3-grams', '4-grams']) {
    for (const r of env.output.getSheetByName(name).rowsFrom(2)) for (const v of r) if (typeof v === 'number') assert.ok(Number.isFinite(v));
  }
  // Sorted by cost, header kept on row 1.
  assert.deepEqual(env.output.getSheetByName('1-grams').rowsFrom(2).map(r => r[1]), ['shoes', 'red', 'for', 'free']);
  // The control sheet logs the run.
  const logRow = env.control.getSheetByName('Settings').rowsFrom(4)[0];
  assert.match(String(logRow[25]), /ran successfully/);
  assert.match(env.awql[0], /during \d{8},\d{8}$/);
});

test('a new output sheet and more than 1,000 nGrams are written in full', () => {
  const many = [];
  for (let i = 0; i < 1100; i++) many.push({ Query: `word${i}`, CampaignName: 'Shopping', Impressions: 5, Clicks: 1, Cost: 1, Conversions: 0, ConversionValue: 0 });
  const env = setup({ queries: many, outputUrl: '' });
  run(env);
  const created = env.books['created-1'];
  assert.equal(created.getName(), 'Shoe Shop - nGrams Output Sheet');
  assert.equal(created.getSheetByName('1-grams').rowsFrom(2).length, 1100);
  assert.match(String(env.control.getSheetByName('Settings').rowsFrom(4)[0][24]), /created-1/);
});

test('ticked nGrams are added to the negative list and the Already Added tab', () => {
  const env = setup({ queries: QUERIES });
  run(env);
  const sheet = env.output.getSheetByName('1-grams');
  const rowNum = sheet.rowsFrom(1).findIndex(r => r[1] === 'red') + 1;
  sheet.getRange(rowNum, 3).setValue(true);
  sheet.getRange(rowNum, 4).setValue('Exact');
  run(env);
  assert.deepEqual(env.lists['nGram negatives'], ['[red]']);
  assert.deepEqual(env.output.getSheetByName('Already Added').rowsFrom(3), [['[red]', 'nGram negatives']]);
  assert.ok(!('red' in tab(env.output, '1-grams')), 'added negatives are no longer listed');
});

test('a number in a comma-separated setting does not stop the run', () => {
  const env = setup({ queries: QUERIES, settings: { CAMPAIGN_NAME_CONTAINS: 2024 } });
  run(env);
  assert.match(env.awql[0], /CampaignName CONTAINS_IGNORE_CASE '2024'/);
  assert.ok(!('free' in tab(env.output, '1-grams')));
});

test('empty account: header only, no email', () => {
  const env = setup({ queries: [], settings: { NOTIFY: true } });
  run(env);
  assert.equal(env.output.getSheetByName('1-grams').getLastRow(), 0);
  assert.equal(env.ads.emails.length, 0);
});

test('NOTIFY sends an email when nGrams are found', () => {
  const env = setup({ queries: QUERIES, settings: { NOTIFY: true } });
  run(env);
  assert.equal(env.ads.emails.length, 1);
  assert.equal(env.ads.emails[0][0].to, 'me@example.com');
});

test('a report failure stops the run with the error', () => {
  const env = setup({ queries: QUERIES });
  env.ads.failAwql = 'AWQL report failed';
  assert.throws(() => run(env), /AWQL report failed/);
});
