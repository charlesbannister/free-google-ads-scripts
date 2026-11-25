'use strict';
// End-to-end runs of the whole script (main) against the mocked Google Ads Scripts runtime
// in ../test-support. Run with: node --test pmax_search_term_n_grams/
const assert = require('node:assert/strict');
const { join } = require('node:path');
const { test } = require('node:test');
const { MockGoogleAds, MockSheet, loadScript } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = join(__dirname, 'pmax_search_term_n_grams.js');
const SHEET = 'demo-sheet';

// The shared mock has no Range.sort or Sheet.setFrozenColumns, and does not refuse ranges
// outside the sheet the way Google does. Added here for this file only.
function strictSheets() {
  const proto = MockSheet.prototype;
  if (!proto.setFrozenColumns) proto.setFrozenColumns = function () { return this; };
  const getRange = proto.getRange;
  proto.getRange = function (a, b, c, d) {
    const range = getRange.call(this, a, b, c, d);
    if (typeof a !== 'string' && (range.row + range.rows - 1 > this.maxRows || range.col + range.cols - 1 > this.maxColumns || range.rows < 1 || range.cols < 1)) {
      throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
    }
    if (!range.sort) {
      Object.getPrototypeOf(range).sort = function (spec) {
        const specs = Array.isArray(spec) ? spec : [typeof spec === 'number' ? { column: spec, ascending: true } : spec];
        const values = this.getValues();
        values.sort((x, y) => {
          for (const s of specs) {
            const i = s.column - this.col; const p = x[i]; const q = y[i];
            const pn = typeof p === 'number'; const qn = typeof q === 'number';
            // Google Sheets: numbers before text ascending, text before numbers descending.
            let cmp = pn && qn ? p - q : pn ? -1 : qn ? 1 : String(p).localeCompare(String(q));
            if (s.ascending === false) cmp = -cmp;
            if (cmp) return cmp;
          }
          return 0;
        });
        values.forEach((line, r) => line.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v)));
        return this;
      };
    }
    return range;
  };
}
strictSheets();

function account() {
  const ads = new MockGoogleAds({ spreadsheetId: SHEET });
  ads.addCampaign({ id: '100', name: 'Search - Generic', channel: 'SEARCH' });
  ads.addCampaign({ id: '500', name: 'PMax - Shoes', channel: 'PERFORMANCE_MAX' });
  ads.addCampaign({ id: '501', name: 'PMax - Brand', channel: 'PERFORMANCE_MAX' });
  ads.addCampaign({ id: '502', name: 'PMax - Old', channel: 'PERFORMANCE_MAX', status: 'PAUSED' });
  // Two days of one term: Google adds them up when no date segment is selected.
  ads.addPmaxSearchTerm({ term: 'Running Shoes', campaignId: '500', impressions: 100, clicks: 10, conversions: 2, value: 50, date: '2026-11-20' });
  ads.addPmaxSearchTerm({ term: 'Running Shoes', campaignId: '500', impressions: 50, clicks: 5, conversions: 1, value: 25, date: '2026-11-21' });
  ads.addPmaxSearchTerm({ term: 'shoes for shoes', campaignId: '500', impressions: 40, clicks: 4, conversions: 1, value: 10 });
  ads.addPmaxSearchTerm({ term: "kids' shoes, size 3", campaignId: '501', impressions: 20, clicks: 2, conversions: 1, value: 5 });
  ads.addPmaxSearchTerm({ term: 'no sale shoes', campaignId: '501', impressions: 30, clicks: 3, conversions: 0, value: 0 });
  ads.addPmaxSearchTerm({ term: 'old shoes', campaignId: '502', impressions: 30, clicks: 3, conversions: 1, value: 1 });
  // campaign_search_term_view also holds Search terms: they must not appear.
  ads.addRows('campaign_search_term_view', [{ 'campaign_search_term_view.search_term': 'search only term', 'campaign.id': '100',
    'metrics.impressions': '9', 'metrics.clicks': '9', 'metrics.conversions': '9', 'metrics.conversions_value': '9' }]);
  ads.addPmaxCategory({ label: 'category label', campaignId: '500', impressions: 99, clicks: 9, conversions: 3 });
  return ads;
}

function run(ads, replace = []) {
  const script = loadScript(SCRIPT, ads.globals(), ['main', 'VERSION'], {
    replace: [['YOUR_SPREADSHEET_URL_HERE', `https://docs.google.com/spreadsheets/d/${SHEET}/edit`], ...replace],
  });
  script.exports.main();
  return script;
}

const tab = (ads, name) => ads.spreadsheet.getSheetByName(name);
const byNgram = (ads, name) => Object.fromEntries(tab(ads, name).rowsFrom(2).map(r => [r[0], r]));

test('real PMax search terms become nGrams with correct totals', () => {
  const ads = account();
  const { exports } = run(ads);
  assert.equal(exports.VERSION, 3);
  const one = byNgram(ads, '1-grams');
  assert.deepEqual(tab(ads, '1-grams').rowsFrom(1)[0], ['nGram', 'Count', 'Campaign ID', 'Campaign Name', 'Impressions', 'Clicks', 'Conversions', 'Conversion Value', 'CTR']);
  // shoes: Running Shoes (two days, 150 imp) + shoes for shoes (counted once, not twice).
  // "kids' shoes, size 3" gives "shoes," (punctuation stays part of the word, as before).
  assert.deepEqual(one.shoes.slice(0, 8), ['shoes', 2, 'All', 'All', 190, 19, 4, 85]);
  assert.equal(one.shoes[8], 19 / 190);
  assert.equal(one.running[4], 150);
  assert.ok(!one['category'], 'categories are only a fallback');
  assert.ok(!one['search'], 'Search campaign terms are filtered out');
  assert.ok(!one['no'], 'MIN_CONVERSIONS > 0 keeps its meaning');
  assert.ok(!one['old'], 'paused campaigns are excluded by default');
  assert.ok(one["kids'"] && one['shoes,'], 'punctuation is kept as part of the word');
  const four = byNgram(ads, '4-grams');
  assert.deepEqual(Object.keys(four), ["kids' shoes, size 3"]);
  // Header stays on top and rows are sorted by conversions, highest first.
  const conv = tab(ads, '1-grams').rowsFrom(2).map(r => r[6]);
  assert.deepEqual(conv, conv.slice().sort((a, b) => b - a));
  // Every cell is finite: no NaN or Infinity reaches the sheet.
  for (const name of ['1-grams', '2-grams', '3-grams', '4-grams']) {
    for (const row of tab(ads, name).rowsFrom(2)) for (const v of row) if (typeof v === 'number') assert.ok(Number.isFinite(v), `${name} ${row}`);
  }
});

test('the date range covers LOOKBACK_DAYS including today', () => {
  const ads = account();
  run(ads);
  const q = ads.queries.find(x => x.includes('FROM campaign_search_term_view'));
  const [, from, to] = q.match(/BETWEEN '(\d{4}-\d{2}-\d{2})' AND '(\d{4}-\d{2}-\d{2})'/);
  assert.equal((Date.parse(to) - Date.parse(from)) / 864e5, 29);
  assert.match(q, /campaign\.advertising_channel_type = 'PERFORMANCE_MAX'/);
});

test('ENABLED_CAMPAIGNS_ONLY = false includes paused campaigns', () => {
  const ads = account();
  run(ads, [['const ENABLED_CAMPAIGNS_ONLY = true;', 'const ENABLED_CAMPAIGNS_ONLY = false;']]);
  assert.ok(byNgram(ads, '1-grams').old);
});

test('SEGMENT_BY_CAMPAIGN splits rows by campaign', () => {
  const ads = account();
  run(ads, [['const SEGMENT_BY_CAMPAIGN = false;', 'const SEGMENT_BY_CAMPAIGN = true;']]);
  const size = tab(ads, '1-grams').rowsFrom(2).filter(r => r[0] === 'shoes' || r[0] === 'size').map(r => [r[0], r[2], r[3], r[1]]);
  assert.deepEqual(size.sort(), [['shoes', '500', 'PMax - Shoes', 2], ['size', '501', 'PMax - Brand', 1]]);
});

test('falls back to search term categories when campaign_search_term_view fails', () => {
  const ads = account();
  ads.failQuery('campaign_search_term_view', 'Unrecognized resource');
  run(ads);
  const one = byNgram(ads, '1-grams');
  assert.equal(one.category[4], 99);
  assert.ok(!one.running);
});

test('empty account writes nothing and does not throw', () => {
  const ads = new MockGoogleAds({ spreadsheetId: SHEET });
  const { logs } = run(ads);
  assert.equal(ads.spreadsheet.getSheets().length, 0);
  assert.ok(logs.some(l => l.includes('No data to write')));
});

test('campaign report failure is logged, not thrown', () => {
  const ads = account();
  ads.failQuery('FROM campaign WHERE', 'Internal error');
  const { logs } = run(ads);
  assert.ok(logs.some(l => l.includes('Error finding matching PMAX campaigns')));
});

test('placeholder spreadsheet URL stops with a message', () => {
  const ads = account();
  const script = loadScript(SCRIPT, ads.globals(), ['main']);
  script.exports.main();
  assert.ok(script.logs.some(l => l.includes("replace 'YOUR_SPREADSHEET_URL_HERE'")));
  assert.equal(ads.queries.length, 0);
});

test('more nGrams than the sheet has rows still writes them all', () => {
  const ads = new MockGoogleAds({ spreadsheetId: SHEET });
  ads.addCampaign({ id: '500', name: 'PMax', channel: 'PERFORMANCE_MAX' });
  for (let i = 0; i < 1200; i++) ads.addPmaxSearchTerm({ term: `term${i}`, campaignId: '500', impressions: 2, clicks: 1, conversions: 1 });
  run(ads);
  assert.equal(tab(ads, '1-grams').rowsFrom(2).length, 1200);
});

test('at MCC level an empty ACCOUNT_ID logs a clear message', () => {
  const ads = account();
  let called = false;
  const MccApp = { accounts: () => ({ withIds: () => ({ withLimit: () => ({ executeInParallel: () => { called = true; } }) }) }) };
  const script = loadScript(SCRIPT, ads.globals({ MccApp }), ['main'], {
    replace: [['YOUR_SPREADSHEET_URL_HERE', `https://docs.google.com/spreadsheets/d/${SHEET}/edit`]],
  });
  script.exports.main();
  assert.equal(called, false);
  assert.ok(script.logs.some(l => l.includes('Set ACCOUNT_ID')));
});
