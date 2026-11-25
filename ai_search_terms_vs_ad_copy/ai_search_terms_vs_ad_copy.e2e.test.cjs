'use strict';
// End-to-end runs of the whole script (main) against the mocked Google Ads Scripts runtime in
// ../test-support. Run with: node --test ai_search_terms_vs_ad_copy/
const assert = require('node:assert/strict');
const { join } = require('node:path');
const { test } = require('node:test');
const { MockGoogleAds, MockSheet, loadScript, parseGaql, iterator } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = join(__dirname, 'ai_search_terms_vs_ad_copy.js');
const SHEET = 'demo-sheet';
const OPENAI = 'https://api.openai.com/v1/chat/completions';

// The shared mock has no ad_group_ad_asset_view resource and no autoResizeColumns, so add them here.
MockSheet.prototype.autoResizeColumns = function () { return this; };

class AdsWithAssets extends MockGoogleAds {
  report(query) {
    const parsed = parseGaql(query);
    if (parsed.resource !== 'ad_group_ad_asset_view') return super.report(query);
    this.queries.push(String(query).replace(/\s+/g, ' ').trim());
    const rows = (this.rows.ad_group_ad_asset_view || []).filter(row => parsed.where.every(c => {
      const v = row[c.field];
      if (c.kind === 'compare' && c.op === '=') return String(v).toUpperCase() === String(c.value).toUpperCase();
      throw new Error(`Test mock: unsupported condition on ${c.field}`);
    }));
    return { rows: () => iterator(rows.map(r => Object.fromEntries(parsed.fields.map(f => [f, r[f] ?? ''])))) };
  }
  addAsset({ adGroupId, adId, type, text, adStatus = 'ENABLED', adType = 'RESPONSIVE_SEARCH_AD', adGroupName = 'Shoes' }) {
    return this.addRows('ad_group_ad_asset_view', [{
      'ad_group.id': adGroupId, 'ad_group.name': adGroupName, 'ad_group_ad.ad.id': adId,
      'ad_group_ad_asset_view.field_type': type, 'asset.text_asset.text': text,
      'ad_group_ad_asset_view.enabled': 'true', 'asset.type': 'TEXT',
      'ad_group_ad.status': adStatus, 'ad_group_ad.ad.type': adType,
    }]);
  }
}

function chat(content) { return { choices: [{ message: { content } }] }; }

function account({ openai = () => ({ status: 200, body: chat('8/10 - relevant.') }), apiKey = 'sk-test' } = {}) {
  const ads = new AdsWithAssets({ spreadsheetId: SHEET, fetch: { [OPENAI]: openai } });
  ads.spreadsheet.insertSheet('settings').getRange('B1').setValue(apiKey);
  ads.addCampaign({ id: '100', name: 'Search - Shoes' });
  ads.addAdGroup({ id: '200', name: 'Shoes', campaignId: '100' });
  ads.addSearchTerm({ term: 'buy running shoes', campaignId: '100', adGroupId: '200', clicks: 5, impressions: 50 });
  ads.addSearchTerm({ term: 'shoe shop', campaignId: '100', adGroupId: '200', clicks: 1, impressions: 10 });
  ads.addSearchTerm({ term: 'never clicked', campaignId: '100', adGroupId: '200', clicks: 0, impressions: 10 });
  ads.addAsset({ adGroupId: '200', adId: '900', type: 'HEADLINE', text: 'Running Shoes $$ Off' });
  ads.addAsset({ adGroupId: '200', adId: '900', type: 'DESCRIPTION', text: 'Free delivery.' });
  ads.addAsset({ adGroupId: '200', adId: '901', type: 'HEADLINE', text: 'Old removed ad', adStatus: 'REMOVED' });
  return ads;
}

function run(ads) {
  const script = loadScript(SCRIPT, ads.globals(), ['main'], {
    replace: [['YOUR_SPREADSHEET_URL_HERE', `https://docs.google.com/spreadsheets/d/${SHEET}/edit`]],
  });
  script.exports.main();
  return script;
}

const results = ads => ads.spreadsheet.getSheetByName('results').rowsFrom(1);

test('normal run: one row per clicked term and live ad, with score and explanation', () => {
  const ads = account();
  run(ads);
  const rows = results(ads);
  assert.equal(rows[0][0], 'Campaign Name');
  assert.equal(rows[0].length, 10);
  // A term with exactly MINIMUM_CLICKS (1) clicks is included; 0 clicks is not; the removed ad is not.
  assert.deepEqual(rows.slice(1).map(r => [r[2], r[3], r[5], r[8]]), [
    ['buy running shoes', 5, '900', 8],
    ['shoe shop', 1, '900', 8],
  ]);
  const sent = JSON.parse(ads.fetches[0].params.payload);
  // Ad text is put into the prompt literally ("$$" must not collapse to "$").
  assert.match(sent.messages[1].content, /Running Shoes \$\$ Off/);
  assert.equal(ads.fetches[0].params.muteHttpExceptions, true);
});

test('score is read from "N/10", not from the "0-10" scale the model repeats', () => {
  const ads = account({ openai: () => ({ status: 200, body: chat('On a 0-10 scale this is a 7/10: close match.') }) });
  run(ads);
  assert.deepEqual(results(ads).slice(1).map(r => r[8]), [7, 7]);
});

test('API error responses are logged and skipped, the run still writes the sheet', () => {
  let n = 0;
  const ads = account({ openai: () => (n++ === 0 ? { status: 429, body: { error: { message: 'rate limit' } } } : { status: 200, body: { choices: [] } }) });
  const { errors } = run(ads);
  assert.ok(errors.some(e => /HTTP 429/.test(e)));
  assert.deepEqual(results(ads), [results(ads)[0]]); // header only
});

test('empty account: no prompts sent, header written', () => {
  const ads = new AdsWithAssets({ spreadsheetId: SHEET, fetch: {} });
  ads.spreadsheet.insertSheet('settings').getRange('B1').setValue('sk-test');
  run(ads);
  assert.equal(ads.fetches.length, 0);
  assert.equal(results(ads).length, 1);
});

test('missing API key stops the run before any OpenAI call', () => {
  const ads = account({ apiKey: 'PASTE_YOUR_OPENAI_API_KEY_HERE' });
  assert.throws(() => run(ads), /API key is missing/);
  assert.equal(ads.fetches.length, 0);
});

test('report failure is reported and stops the run', () => {
  const ads = account();
  ads.failQuery('search_term_view', 'Internal error');
  assert.throws(() => run(ads), /Internal error/);
});
