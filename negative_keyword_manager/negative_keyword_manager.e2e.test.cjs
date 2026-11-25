'use strict';
// End-to-end runs of the whole script (main) against the mocked Google Ads Scripts runtime
// in ../test-support. Run with: node --test negative_keyword_manager/
//
// The script reads search terms with an AWQL report (SEARCH_QUERY_PERFORMANCE_REPORT) and
// uses Sheets features the shared mock lacks (data validation, sort, clear, bounds checks).
// This file adds them locally; this test runs in its own process, so nothing leaks.
const assert = require('node:assert/strict');
const { join } = require('node:path');
const { test } = require('node:test');
const { MockGoogleAds, MockSheet, loadScript, iterator } = require('../test-support/google-ads-scripts-mock.cjs');
// MockRange is not exported; take it from a range.
const MockRange = new MockSheet('probe', 0).getRange(1, 1).constructor;

const SCRIPT = join(__dirname, 'negative_keyword_manager.js');
const SHEET_ID = 'demo-sheet';

// ------------------------------------------------------------------ Sheets extensions

// Real Sheets refuses a range that runs past the end of the sheet.
const baseGetRange = MockSheet.prototype.getRange;
MockSheet.prototype.getRange = function (a, b, c, d) {
  const range = baseGetRange.call(this, a, b, c, d);
  if (range.rows < 1 || range.cols < 1) throw new Error('The number of rows in the range must be at least 1.');
  if (range.row + range.rows - 1 > this.maxRows || range.col + range.cols - 1 > this.maxColumns) {
    throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
  }
  return range;
};
MockRange.prototype.clear = function () { return this.clearContent(); };
MockRange.prototype.setDataValidation = function (rule) { (this.sheet.validations ||= []).push({ col: this.col, rule }); return this; };
MockRange.prototype.sort = function ({ column, ascending }) {
  const values = this.getValues();
  const i = column - this.col;
  values.sort((x, y) => (ascending ? 1 : -1) * (Number(x[i]) - Number(y[i])));
  values.forEach((line, r) => line.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v)));
  return this;
};

// ------------------------------------------------------------------ AWQL + ad group extensions

class ManagerAds extends MockGoogleAds {
  constructor(options) {
    super(options);
    this.keywordsAdded = [];
    this.awqlRows = [];
  }

  /** A search term as the AWQL search query report returns it. */
  addAwqlTerm({ term, campaignId, adGroupId, impressions = 10, clicks = 1, cost = 1, conversions = 0, value = 0, status = 'NONE' }) {
    const campaign = this.campaigns.find(c => c.id === String(campaignId));
    const adGroup = this.adGroups.find(g => g.id === String(adGroupId));
    this.awqlRows.push({
      AdGroupId: String(adGroupId), CampaignId: String(campaignId), CampaignName: campaign.name, AdGroupName: adGroup.name,
      Query: term, ConversionValue: String(value), Impressions: String(impressions), Clicks: String(clicks),
      Cost: cost.toLocaleString('en-GB'), Conversions: String(conversions), TopImpressionPercentage: '0',
      AbsoluteTopImpressionPercentage: '0', QueryTargetingStatus: status,
    });
  }

  report(query) {
    if (!/SEARCH_QUERY_PERFORMANCE_REPORT/.test(query)) return super.report(query);
    this.queries.push(String(query).replace(/\s+/g, ' ').trim());
    for (const f of this.failures) if (query.includes(f.match)) throw new Error(f.message);
    const min = (field) => { const m = query.match(new RegExp(`and ${field} > ([\\d.]+)`)); return m ? Number(m[1]) : null; };
    const clicks = min('Clicks');
    const rows = this.awqlRows.filter(r => clicks === null || Number(r.Clicks) > clicks).map(r => ({ ...r }));
    return { rows: () => iterator(rows) };
  }

  globals(extra) {
    const g = super.globals(extra);
    const ads = this;
    const adGroups = g.AdsApp.adGroups;
    g.AdsApp.adGroups = () => {
      const selector = adGroups();
      const get = selector.get;
      selector.get = () => {
        const it = get(); const items = [];
        while (it.hasNext()) items.push(it.next());
        return iterator(items.map(adGroup => ({
          ...adGroup,
          // Google Ads refuses keywords over 10 words or 80 characters.
          createNegativeKeyword(text) {
            const plain = String(text).replace(/^\[|\]$|^"|"$/g, '');
            if (plain.length > 80 || plain.split(/\s+/).length > 10) throw new Error('Keyword text is too long or has too many words.');
            adGroup.createNegativeKeyword(text);
          },
          newKeywordBuilder() {
            const kw = { adGroupId: String(adGroup.getId()) };
            const builder = {
              withText(t) { kw.text = t; return builder; },
              withCpc(c) { kw.cpc = c; return builder; },
              build() {
                const ok = kw.cpc === undefined || (isFinite(kw.cpc) && kw.cpc > 0);
                if (ok) ads.keywordsAdded.push(kw);
                return { isSuccessful: () => ok, getErrors: () => (ok ? [] : ['Invalid CPC']) };
              },
            };
            return builder;
          },
        })));
      };
      return selector;
    };
    const validation = () => {
      const b = { requireValueInList: () => b, requireCheckbox: () => b, setAllowInvalid: () => b, build: () => ({}) };
      return b;
    };
    g.SpreadsheetApp.newDataValidation = validation;
    g.Logger = { log: () => {} };
    return g;
  }
}

// ------------------------------------------------------------------ the spreadsheet

const HEADER = ['Campaign Name', 'Adgroup', 'Ad Group Id', 'Search Term', 'Query Targeting Status', 'Add as Keyword',
  'Match Type', 'Ignore', 'Add as Ad Group Negative', 'Impr.', 'Clicks', 'Ctr', 'Cost', 'Conversions', 'Cost / Conv.',
  'Conv. Rate', 'Conv. Value', 'Roas'];

/** Settings tab in the template's layout (rules from row 4) plus the report and log tabs. */
function spreadsheet(ads, { rules = [{}], maxRows = 1000, termsSheet = true } = {}) {
  const settings = ads.spreadsheet.insertSheet('Settings');
  settings.insertColumnsAfter(26, 10);
  for (let col = 1; col <= 33; col++) settings.getRange(3, col).setValue(col === 32 ? 'Logs' : col === 33 ? 'Last Run' : `H${col}`);
  rules.forEach((rule, i) => {
    const row = [
      rule.name || 'Rule 1', rule.emails || 'me@example.com', rule.flag ?? 'Yes', rule.emailAlert || 'No', rule.changesEmail || 'No',
      rule.tab || 'Search Terms', rule.n ?? 30, rule.contains ?? '', '', 'ALL', 'Exact', rule.defaultAction || '',
      ...(rule.filters || ['Clicks', '>', '0']), ...Array(18 - (rule.filters || [0, 0, 0]).length).fill(''),
      '', '',
    ];
    row.forEach((v, c) => settings.getRange(4 + i, c + 1).setValue(v));
  });
  if (termsSheet) {
    const terms = ads.spreadsheet.insertSheet('Search Terms');
    terms.maxRows = maxRows;
    terms.getRange('A1').setValue('Search Terms');
    HEADER.forEach((h, c) => terms.getRange(3, c + 1).setValue(h));
  }
  for (const name of ['Negative Keywords', 'Ignored Terms']) {
    const s = ads.spreadsheet.insertSheet(name);
    ['Campaign', 'Ad Group', 'Search Term'].forEach((h, c) => s.getRange(4, c + 1).setValue(h));
  }
  return settings;
}

/** Ticks a row of the Search Terms tab the way a user would. */
function tick(ads, row, { negative, keyword, ignore, matchType = 'Exact' }) {
  const s = ads.spreadsheet.getSheetByName('Search Terms');
  const col = name => HEADER.indexOf(name) + 1;
  s.getRange(row, col('Add as Ad Group Negative')).setValue(negative ? true : 'FALSE');
  s.getRange(row, col('Add as Keyword')).setValue(keyword ? true : 'FALSE');
  s.getRange(row, col('Ignore')).setValue(!!ignore);
  s.getRange(row, col('Match Type')).setValue(matchType);
}

function account(options = {}) {
  const ads = new ManagerAds({ preview: options.preview, spreadsheetId: SHEET_ID });
  ads.addCampaign({ id: '100', name: 'Search - Generic' });
  ads.addAdGroup({ id: '200', name: 'Generic', campaignId: '100' });
  ads.addAwqlTerm({ term: 'cheap socks', campaignId: '100', adGroupId: '200', clicks: 3, cost: 4.5 });
  ads.addAwqlTerm({ term: 'buy shoes', campaignId: '100', adGroupId: '200', clicks: 9, cost: 12 });
  ads.addAwqlTerm({ term: 'free shoes', campaignId: '100', adGroupId: '200', clicks: 1, cost: 0 });
  ads.addAwqlTerm({ term: 'one two three four five six seven eight nine ten eleven', campaignId: '100', adGroupId: '200', clicks: 2, cost: 30 });
  return ads;
}

function run(ads) {
  const script = loadScript(SCRIPT, ads.globals(), ['main', 'VERSION'], {
    replace: [['YOUR_SPREADSHEET_URL_HERE', `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit`]],
  });
  let error = null;
  try { script.exports.main(); } catch (e) { error = e; }
  return { ...script, error };
}

const termsRows = ads => ads.spreadsheet.getSheetByName('Search Terms').rowsFrom(4).filter(r => r[0] !== '');
const termsOrder = ads => termsRows(ads).map(r => r[3]);
const logs = (ads, row = 4) => ads.spreadsheet.getSheetByName('Settings').getRange(row, 32).getValue();

// ------------------------------------------------------------------ tests

test('first run writes every search term, sorted by cost (highest first)', () => {
  const ads = account();
  spreadsheet(ads);
  const { error } = run(ads);
  assert.equal(error, null);
  assert.deepEqual(termsOrder(ads), ['one two three four five six seven eight nine ten eleven', 'buy shoes', 'cheap socks', 'free shoes']);
  assert.match(String(logs(ads)), /The script ran successfully/);
});

test('ticked terms: negative added and logged, ignored term logged, both hidden from the new report', () => {
  const ads = account();
  spreadsheet(ads);
  run(ads);
  const rows = termsRows(ads);
  tick(ads, 4 + rows.findIndex(r => r[3] === 'cheap socks'), { negative: true });
  tick(ads, 4 + rows.findIndex(r => r[3] === 'buy shoes'), { ignore: true });
  const { error } = run(ads);
  assert.equal(error, null);
  assert.deepEqual(ads.entityNegatives, [{ level: 'adgroup', id: '200', text: '[cheap socks]' }]);
  assert.deepEqual(ads.spreadsheet.getSheetByName('Negative Keywords').rowsFrom(5).map(r => r.slice(0, 3)), [['Search - Generic', 'Generic', 'cheap socks']]);
  assert.deepEqual(ads.spreadsheet.getSheetByName('Ignored Terms').rowsFrom(5).map(r => r.slice(0, 3)), [['Search - Generic', 'Generic', 'buy shoes']]);
  assert.deepEqual(termsOrder(ads).sort(), ['free shoes', 'one two three four five six seven eight nine ten eleven']);
});

test('unticked rows written as the text "FALSE" are not acted on', () => {
  const ads = account();
  spreadsheet(ads);
  run(ads);
  for (let row = 4; row < 8; row++) tick(ads, row, {});
  run(ads);
  assert.equal(ads.entityNegatives.length + ads.keywordsAdded.length, 0);
});

test('a refused negative (over 10 words) does not stop the others, and is not logged as added', () => {
  const ads = account();
  spreadsheet(ads);
  run(ads);
  const rows = termsRows(ads);
  tick(ads, 4 + rows.findIndex(r => r[3].startsWith('one two')), { negative: true });
  tick(ads, 4 + rows.findIndex(r => r[3] === 'cheap socks'), { negative: true });
  const { error } = run(ads);
  assert.equal(error, null);
  assert.deepEqual(ads.entityNegatives.map(n => n.text), ['[cheap socks]']);
  assert.equal(ads.spreadsheet.getSheetByName('Negative Keywords').rowsFrom(5).length, 1);
});

test('add as keyword: CPC from cost / clicks, and no CPC (default bid) when cost is 0', () => {
  const ads = account();
  spreadsheet(ads);
  run(ads);
  const rows = termsRows(ads);
  tick(ads, 4 + rows.findIndex(r => r[3] === 'cheap socks'), { keyword: true });
  tick(ads, 4 + rows.findIndex(r => r[3] === 'free shoes'), { keyword: true, matchType: 'Phrase' });
  run(ads);
  assert.deepEqual(ads.keywordsAdded.sort((a, b) => a.text.localeCompare(b.text)), [
    { adGroupId: '200', text: '"free shoes"' },
    { adGroupId: '200', text: '[cheap socks]', cpc: 1.5 },
  ]);
});

test('preview: the negative is not logged, so the term is not hidden from the next real run', () => {
  const ads = account();
  spreadsheet(ads);
  run(ads);
  const row = 4 + termsRows(ads).findIndex(r => r[3] === 'cheap socks');
  tick(ads, row, { negative: true });
  ads.preview = true;
  run(ads);
  assert.equal(ads.spreadsheet.getSheetByName('Negative Keywords').rowsFrom(5).length, 0);
  ads.preview = false;
  run(ads);
  assert.ok(termsOrder(ads).includes('cheap socks'));
});

test('a full Search Terms tab grows to fit the report', () => {
  const ads = account();
  for (let i = 0; i < 40; i++) ads.addAwqlTerm({ term: `term ${i}`, campaignId: '100', adGroupId: '200', clicks: 1, cost: i });
  spreadsheet(ads, { maxRows: 10 });
  const { error } = run(ads);
  assert.equal(error, null);
  assert.equal(termsRows(ads).length, 44);
});

test('empty account: no rows, no error', () => {
  const ads = new ManagerAds({ spreadsheetId: SHEET_ID });
  ads.addCampaign({ id: '100', name: 'Search - Generic' });
  spreadsheet(ads);
  const { error } = run(ads);
  assert.equal(error, null);
  assert.equal(termsRows(ads).length, 0);
});

test('report failure on one rule does not stop the next rule; the run still fails at the end', () => {
  const ads = account();
  spreadsheet(ads, { rules: [{ name: 'Broken', contains: 'boom' }, { name: 'Fine' }] });
  ads.failQuery("CONTAINS_IGNORE_CASE 'boom'", 'Report failed');
  const { error } = run(ads);
  assert.match(String(error && error.message), /1 rule\(s\) failed\. Broken\/4: Report failed/);
  assert.match(String(logs(ads, 5)), /The script ran successfully/);
  assert.match(String(logs(ads, 4)), /didn't finish successfully/);
});

test('bad sheet input: a checkbox in "Run Script", a numeric campaign filter, a missing tab', () => {
  const ads = account();
  spreadsheet(ads, { rules: [{ name: 'Checkbox', flag: true }, { name: 'Number', contains: 2024 }, { name: 'No tab', tab: 'Missing' }] });
  const { error } = run(ads);
  assert.match(String(error && error.message), /1 rule\(s\) failed\. No tab\/6: The sheet Missing does not exist/);
  assert.ok(ads.queries.some(q => q.includes("CONTAINS_IGNORE_CASE '2024'")));
});

test('no rules switched on: nothing runs, no error', () => {
  const ads = account();
  spreadsheet(ads, { rules: [{ flag: 'No' }] });
  const { error, warnings } = run(ads);
  assert.equal(error, null);
  assert.ok(warnings.some(w => w.includes('No rules were specified')));
});

test('the version is bumped', () => {
  const ads = account();
  spreadsheet(ads);
  assert.equal(run(ads).exports.VERSION, '1.4');
});
