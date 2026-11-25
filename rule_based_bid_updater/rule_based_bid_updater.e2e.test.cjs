'use strict';
// End-to-end runs of the whole script (main) against the mocked Google Ads Scripts runtime in
// ../test-support. Run with: node --test rule_based_bid_updater/
const assert = require('node:assert/strict');
const { join } = require('node:path');
const { test } = require('node:test');
const { MockGoogleAds, MockSheet, loadScript, parseGaql, iterator } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = process.env.SCRIPT_UNDER_TEST || join(__dirname, 'rule_based_bid_updater.js');
const SHEET = 'demo-sheet';
const URL = `https://docs.google.com/spreadsheets/d/${SHEET}/edit`;

// Metric fields these queries may use (a typo such as metrics.conversion_value is refused, as Google does).
const KNOWN_METRICS = new Set(['metrics.clicks', 'metrics.conversions', 'metrics.conversions_value', 'metrics.cost_micros',
  'metrics.impressions', 'metrics.top_impression_percentage', 'metrics.search_impression_share',
  'metrics.search_absolute_top_impression_share']);

// The shared mock has no keyword/product-group entities or product_group_view, so they are added here.
class BidAccount extends MockGoogleAds {
  constructor(options) {
    super(options);
    this.keywordEntities = []; this.productEntities = []; this.changes = [];
  }
  report(query) {
    const parsed = parseGaql(query);
    for (const f of [...parsed.fields, ...parsed.where.map(c => c.field)]) {
      if (f.startsWith('metrics.') && !KNOWN_METRICS.has(f)) throw new Error(`Unrecognized field in the query: '${f}'.`);
    }
    if (parsed.resource !== 'product_group_view') return super.report(query);
    // Same field rules as keyword_view: run it there against the product rows.
    const saved = this.rows.keyword_view;
    this.rows.keyword_view = this.rows.product_group_view || [];
    try { return super.report(String(query).replace(/FROM product_group_view/, 'FROM keyword_view')); } finally { this.rows.keyword_view = saved; }
  }
  addKeyword({ adGroupId, id, text, bid, effectiveBid = bid, campaignId, strategy = 'MANUAL_CPC', clicks = 0, cost = 0, conversions = 0, value = 0, impressions = 100, failSetCpc = false }) {
    this.addRows('keyword_view', [{
      'ad_group.id': adGroupId, 'ad_group.name': `AG ${adGroupId}`, 'ad_group.status': 'ENABLED', 'campaign.id': campaignId,
      'campaign.bidding_strategy_type': strategy, 'ad_group_criterion.criterion_id': id, 'ad_group_criterion.status': 'ENABLED',
      'ad_group_criterion.keyword.text': text, 'ad_group_criterion.keyword.match_type': 'EXACT',
      'ad_group_criterion.cpc_bid_micros': String(Math.round(bid * 1e6)),
      'ad_group_criterion.effective_cpc_bid_micros': String(Math.round(effectiveBid * 1e6)),
      'metrics.impressions': String(impressions), 'metrics.clicks': String(clicks), 'metrics.cost_micros': String(Math.round(cost * 1e6)),
      'metrics.conversions': String(conversions), 'metrics.conversions_value': String(value),
    }]);
    this.keywordEntities.push({ adGroupId: String(adGroupId), id: String(id), failSetCpc });
  }
  addProductGroup({ adGroupId, id, name, bid, campaignId, negative = false, clicks = 0, cost = 0, failExclude = false }) {
    this.addRows('product_group_view', [{
      'ad_group.id': adGroupId, 'ad_group.name': `AG ${adGroupId}`, 'ad_group.status': 'ENABLED', 'campaign.id': campaignId,
      'campaign.bidding_strategy_type': 'MANUAL_CPC', 'ad_group_criterion.criterion_id': id, 'ad_group_criterion.display_name': name,
      'ad_group_criterion.negative': negative ? 'TRUE' : 'FALSE',
      'ad_group_criterion.cpc_bid_micros': String(Math.round(bid * 1e6)), 'ad_group_criterion.effective_cpc_bid_micros': String(Math.round(bid * 1e6)),
      'metrics.impressions': '100', 'metrics.clicks': String(clicks), 'metrics.cost_micros': String(Math.round(cost * 1e6)),
      'metrics.conversions': '0', 'metrics.conversions_value': '0',
    }]);
    this.productEntities.push({ adGroupId: String(adGroupId), id: String(id), failExclude });
  }
  _entitySelector(list, kind) {
    const ads = this;
    const state = { ids: [] };
    const selector = {
      withIds(ids) { state.ids = ids.map(pair => pair.map(String).join('~')); return selector; },
      get() {
        return iterator(list.filter(e => state.ids.includes(`${e.adGroupId}~${e.id}`)).map(e => ({
          getId: () => Number(e.id),
          getAdGroup: () => ({ getId: () => Number(e.adGroupId) }),
          bidding: () => ({ setCpc: cpc => { if (e.failSetCpc) throw new Error('Bid strategy refuses manual bids'); ads.changes.push([kind, e.id, 'setCpc', cpc]); } }),
          pause: () => ads.changes.push([kind, e.id, 'pause']),
          setMaxCpc: cpc => ads.changes.push([kind, e.id, 'setMaxCpc', cpc]),
          exclude: () => { if (e.failExclude) throw new Error('Cannot exclude'); ads.changes.push([kind, e.id, 'exclude']); },
        })));
      },
    };
    return selector;
  }
  globals() {
    const g = super.globals();
    g.AdsApp.keywords = () => this._entitySelector(this.keywordEntities, 'keyword');
    g.AdsApp.productGroups = () => this._entitySelector(this.productEntities, 'product');
    g.Logger = { log: () => {} };
    return g;
  }
}

// The Settings tab: three header rows, one rule per row from row 4. Column order is the script's header order.
const HEADER = ['NAME', 'EMAILS', 'FLAG', 'N', 'ACTION', 'CHANGE', 'CAMPAIGN_NAME_CONTAINS', 'CAMPAIGN_NAME_NOT_CONTAINS'];
for (let i = 1; i <= 6; i++) HEADER.push(`FILTER_${i}_METRIC`, `FILTER_${i}_OPERATOR`, `FILTER_${i}_VALUE`);
HEADER.push('MIN_BID', 'MAX_BID', 'INCLUDE_SHOPPING', 'INCLUDE_TEXT', 'LOG_SHEET_URL', 'Logs', 'Last Run');

function account({ rule = {}, filters = [], preview = false } = {}) {
  const ads = new BidAccount({ spreadsheetId: SHEET, preview });
  const settings = new MockSheet('Settings', 1, { maxColumns: 40 });
  ads.spreadsheet.sheets.push(settings, new MockSheet('Keywords', 2), new MockSheet('Products', 3));
  settings.getRange(3, 1, 1, HEADER.length).setValues([HEADER]);
  const values = {
    NAME: 'Rule 1', EMAILS: '', FLAG: 'Yes', N: 30, ACTION: 'Increase by %', CHANGE: 0.1,
    CAMPAIGN_NAME_CONTAINS: '', CAMPAIGN_NAME_NOT_CONTAINS: '', MIN_BID: 0.1, MAX_BID: 2,
    INCLUDE_SHOPPING: 'No', INCLUDE_TEXT: 'Yes', LOG_SHEET_URL: URL, ...rule,
  };
  filters.forEach(([m, o, v], i) => Object.assign(values, { [`FILTER_${i + 1}_METRIC`]: m, [`FILTER_${i + 1}_OPERATOR`]: o, [`FILTER_${i + 1}_VALUE`]: v }));
  settings.getRange(4, 1, 1, HEADER.length - 2).setValues([HEADER.slice(0, -2).map(h => values[h] ?? '')]);

  ads.addCampaign({ id: '100', name: 'Search - Brand Terms' });
  ads.addCampaign({ id: '101', name: 'Search - Generic' });
  ads.addCampaign({ id: '102', name: 'Search - Brand tCPA' });
  ads.addKeyword({ campaignId: '100', adGroupId: 10, id: 1, text: 'brand shoes', bid: 1.23, clicks: 20, cost: 10, value: 50 });
  ads.addKeyword({ campaignId: '100', adGroupId: 10, id: 2, text: 'brand boots', bid: 0, effectiveBid: 0.5, clicks: 20, cost: 10 });
  ads.addKeyword({ campaignId: '100', adGroupId: 10, id: 3, text: 'brand capped', bid: 2, clicks: 20, cost: 10 });
  ads.addKeyword({ campaignId: '101', adGroupId: 11, id: 4, text: 'generic shoes', bid: 1, clicks: 20, cost: 10 });
  ads.addKeyword({ campaignId: '102', adGroupId: 12, id: 5, text: 'brand smart', bid: 1, clicks: 20, cost: 10, strategy: 'TARGET_CPA' });
  ads.addKeyword({ campaignId: '100', adGroupId: 10, id: 6, text: 'brand few clicks', bid: 1, clicks: 2, cost: 1 });
  return ads;
}

function run(ads) {
  const script = loadScript(SCRIPT, ads.globals(), ['main'], { replace: [['YOUR_SPREADSHEET_URL_HERE', URL]] });
  script.exports.main();
  return script;
}

const keywordLog = ads => ads.spreadsheet.getSheetByName('Keywords').rowsFrom(4);
const controlLog = ads => ads.spreadsheet.getSheetByName('Settings').getRange(4, HEADER.indexOf('Logs') + 1).getValue();
const ymd = daysAgo => { const d = new Date(); d.setDate(d.getDate() - daysAgo); return d.toISOString().slice(0, 10); };

test('normal run: manual CPC keywords matching the rule get new bids, rounded and capped', () => {
  const ads = account({ rule: { CAMPAIGN_NAME_CONTAINS: 'brand' }, filters: [['Clicks', '>', 10], ['ConversionValue', '>=', 0]] });
  run(ads);
  assert.deepEqual(ads.changes, [
    ['keyword', '1', 'setCpc', 1.35],  // 1.23 * 1.1 = 1.353, rounded to 0.01
    ['keyword', '2', 'setCpc', 0.55],  // no own bid: starts from the ad group bid it uses (0.50)
  ]);
  // 3 is already at the max bid (no change re-sent), 4 is outside "brand", 5 uses Target CPA, 6 has too few clicks.
  const rows = keywordLog(ads);
  assert.deepEqual(rows.map(r => [r[2], r[17], r[18]]), [['brand shoes', 1.23, 1.35], ['brand boots', 0.5, 0.55]]);
  assert.equal(rows[0].length, 21);
  assert.match(controlLog(ads), /The script ran successfully/);
});

test('query covers the last N days up to yesterday with quoted dates', () => {
  const ads = account();
  run(ads);
  const q = ads.queries.find(x => /FROM keyword_view/.test(x));
  assert.ok(q.includes(`segments.date BETWEEN '${ymd(30)}' AND '${ymd(1)}'`), q);
});

test('campaign name does not contain filter excludes matching campaigns', () => {
  const ads = account({ rule: { CAMPAIGN_NAME_NOT_CONTAINS: 'BRAND' }, filters: [['Clicks', '>', 10]] });
  run(ads);
  assert.deepEqual(ads.changes, [['keyword', '4', 'setCpc', 1.1]]);
});

test('CpcBid > avg filter: only keywords bidding above their average CPC change', () => {
  // Average CPC is 0.50 (10 / 20) for every keyword; the 0.50-bid "boots" is not above it ("capped" is already at the max bid).
  const ads = account({ filters: [['Clicks', '>', 10], ['CpcBid', '>', 'avg']] });
  run(ads);
  assert.deepEqual(ads.changes.map(c => c[1]), ['1', '4']);
});

test('pause action in preview mode: pause called, logged as preview', () => {
  const ads = account({ preview: true, rule: { ACTION: 'Pause' }, filters: [['Clicks', '>', 10]] });
  run(ads);
  assert.deepEqual(ads.changes.map(c => [c[1], c[2]]), [['1', 'pause'], ['2', 'pause'], ['3', 'pause'], ['4', 'pause']]);
  assert.equal(keywordLog(ads)[0][19], '(Preview Mode) Pause');
  assert.match(controlLog(ads), /preview mode/);
});

test('a bid refused by Google is logged and skipped, the run continues', () => {
  const ads = account({ filters: [['Clicks', '>', 10]] });
  ads.keywordEntities.find(e => e.id === '1').failSetCpc = true;
  run(ads);
  assert.deepEqual(ads.changes.map(c => c[1]), ['2', '4']);
  assert.match(controlLog(ads), /ran successfully/);
});

test('shopping: excluded product groups are not queried; a failed exclude does not stop the run', () => {
  const ads = account({ rule: { INCLUDE_SHOPPING: 'Yes', INCLUDE_TEXT: 'No', ACTION: 'Pause' }, filters: [['Clicks', '>', 10]] });
  ads.addCampaign({ id: '300', name: 'Shopping', channel: 'SHOPPING' });
  ads.addProductGroup({ campaignId: '300', adGroupId: 30, id: 31, name: 'red shoes', bid: 1, clicks: 20, cost: 5, failExclude: true });
  ads.addProductGroup({ campaignId: '300', adGroupId: 30, id: 32, name: 'blue shoes', bid: 1, clicks: 20, cost: 5 });
  ads.addProductGroup({ campaignId: '300', adGroupId: 30, id: 33, name: 'already excluded', bid: 1, clicks: 20, cost: 5, negative: true });
  run(ads);
  assert.deepEqual(ads.changes, [['product', '32', 'exclude']]);
  assert.match(controlLog(ads), /ran successfully/);
});

test('empty account: no changes and a successful log', () => {
  const ads = account();
  ads.rows.keyword_view = [];
  run(ads);
  assert.deepEqual(ads.changes, []);
  assert.match(controlLog(ads), /ran successfully/);
});

test('bad sheet input: a missing max bid is written to the Logs column and stops the row', () => {
  const ads = account({ rule: { MAX_BID: '' } });
  assert.throws(() => run(ads), /max bid/);
  assert.equal(controlLog(ads), 'Please set a max bid');
  assert.deepEqual(ads.changes, []);
});

test('report failure stops the run before any change', () => {
  const ads = account();
  ads.failQuery('keyword_view', 'Internal error');
  assert.throws(() => run(ads), /Internal error/);
  assert.deepEqual(ads.changes, []);
});
