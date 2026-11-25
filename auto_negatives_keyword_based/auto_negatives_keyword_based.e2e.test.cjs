'use strict';
// End-to-end runs of the whole script (main) against the mocked Google Ads Scripts runtime
// in ../test-support. Run with: node --test auto_negatives_keyword_based/
const assert = require('node:assert/strict');
const { join } = require('node:path');
const { test } = require('node:test');
const { MockGoogleAds, MockSheet, loadScript } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = join(__dirname, 'auto_negatives_keyword_based.js');
const SHEET = 'demo-sheet';

// Sheet calls this script uses that the shared mock lacks (patched in this test process only).
MockSheet.prototype.autoResizeColumn = function () { return this; };
MockSheet.prototype.appendRow = function (row) {
  const r = this.getLastRow() + 1;
  row.forEach((v, c) => this._set(r, c + 1, v));
  return this;
};
const RangeProto = Object.getPrototypeOf(new MockSheet("x", 1).getRange(1, 1));
RangeProto.insertCheckboxes = function () { return this; };
RangeProto.clear = function () { return this.clearContent(); };

// Ad group entities that refuse what Google refuses (long or duplicate keywords).
class Account extends MockGoogleAds {
  addAdGroup(opts) {
    const g = super.addAdGroup(opts);
    this.addRows('ad_group', [{ 'campaign.id': g.campaignId, 'ad_group.id': g.id, 'ad_group.name': g.name, 'ad_group.status': g.status }]);
    return g;
  }
  addKeyword({ adGroupId, text, matchType = 'PHRASE', clicks = 0 }) {
    const g = this.adGroups.find(x => x.id === String(adGroupId));
    return this.addRows('keyword_view', [{
      'campaign.id': g.campaignId, 'ad_group.id': g.id, 'ad_group.status': g.status,
      'ad_group_criterion.status': 'ENABLED', 'ad_group_criterion.negative': 'FALSE',
      'ad_group_criterion.keyword.text': text, 'ad_group_criterion.keyword.match_type': matchType,
      'metrics.clicks': String(clicks), 'segments.date': '2026-11-20',
    }]);
  }
  _adGroupSelector(channels) {
    const selector = super._adGroupSelector(channels);
    const ads = this;
    const get = selector.get;
    selector.get = function () {
      const it = get.call(this);
      return {
        hasNext: it.hasNext,
        next() {
          const e = it.next();
          return {
            ...e,
            createNegativeKeyword(text) {
              const plain = String(text).replace(/^\[|\]$|^"|"$/g, '');
              if (plain.length > 80 || plain.split(/\s+/).length > 10) throw new Error('KEYWORD_TEXT_TOO_LONG');
              const dup = ads.entityNegatives.some(n => n.id === String(e.getId()) && n.text === text);
              e.createNegativeKeyword(text);
              if (dup) ads.duplicateSends = (ads.duplicateSends || 0) + 1;
            },
          };
        },
      };
    };
    return selector;
  }
}

// One Search campaign (shoe keywords), one PMax campaign, and a paused campaign.
function account({ preview = false } = {}) {
  const ads = new Account({ preview, spreadsheetId: SHEET });
  ads.addCampaign({ id: '100', name: 'Search - Shoes', channel: 'SEARCH' });
  ads.addAdGroup({ id: '200', name: 'Running Shoes', campaignId: '100' });
  ads.addKeyword({ adGroupId: '200', text: 'running shoes', clicks: 50 });
  ads.addKeyword({ adGroupId: '200', text: 'trainers', clicks: 10 });
  ads.addSearchTerm({ term: 'best running shoes', campaignId: '100', adGroupId: '200', impressions: 100, clicks: 10, cost: 5 });
  ads.addSearchTerm({ term: 'chocolate cake recipe', campaignId: '100', adGroupId: '200', impressions: 40, clicks: 3, cost: 2.5 });
  ads.addSearchTerm({ term: 'free knitting pattern', campaignId: '100', adGroupId: '200', impressions: 20, clicks: 1, cost: 0.4 });
  ads.addSearchTerm({ term: 'old excluded junk', campaignId: '100', adGroupId: '200', status: 'EXCLUDED', impressions: 9, clicks: 1, cost: 1 });
  ads.addSearchTerm({ term: 'converting gardening gloves', campaignId: '100', adGroupId: '200', impressions: 9, clicks: 2, cost: 1, conversions: 1 });
  ads.addCampaign({ id: '500', name: 'PMax - All', channel: 'PERFORMANCE_MAX' });
  ads.addPmaxSearchTerm({ term: 'pmax term', campaignId: '500', impressions: 30, clicks: 1, cost: 1 });
  return ads;
}

function run(ads, replace = []) {
  const script = loadScript(SCRIPT, ads.globals(), ['main'], {
    replace: [['PASTE_YOUR_SPREADSHEET_URL_HERE\';', `https://docs.google.com/spreadsheets/d/${SHEET}/edit';`], ...replace],
  });
  script.exports.main();
  return script;
}

const output = ads => ads.spreadsheet.getSheetByName('Output');
const logs = ads => ads.spreadsheet.getSheetByName('Logs');
const termsIn = sheet => sheet.rowsFrom(2).map(r => r[1]);

test('normal run lists unmatched, non-converting terms in the fixed 18 columns', () => {
  const ads = account();
  run(ads);
  const sheet = output(ads);
  assert.equal(sheet.rowsFrom(1)[0].length, 18);
  assert.deepEqual(sheet.rowsFrom(1)[0].slice(0, 3), ['Negate', 'Search Term', 'Keywords']);
  assert.deepEqual(termsIn(sheet), ['chocolate cake recipe', 'free knitting pattern']);
  const row = sheet.rowsFrom(2)[0];
  assert.equal(row[0], false);
  assert.equal(row[6], '100');
  assert.equal(row[7], '200');
  assert.equal(row[10], 2.5);
  assert.equal(ads.entityNegatives.length, 0);
  assert.deepEqual(logs(ads).rowsFrom(1)[0].slice(0, 2), ['Timestamp', 'Search Term']);
});

test('already-excluded search terms are not suggested again', () => {
  const ads = account();
  run(ads);
  assert.ok(!termsIn(output(ads)).includes('old excluded junk'));
  assert.ok(ads.queries.some(q => q.includes("search_term_view.status NOT IN ('EXCLUDED', 'ADDED_EXCLUDED')")));
});

test('Performance Max terms never reach the ad-group output', () => {
  const ads = account();
  run(ads);
  assert.ok(!termsIn(output(ads)).includes('pmax term'));
});

test('ticked rows are added as exact negatives, logged, and not listed again', () => {
  const ads = account();
  run(ads);
  output(ads).getRange(2, 1).setValue(true); // tick "chocolate cake recipe"
  run(ads);
  assert.deepEqual(ads.entityNegatives, [{ level: 'adgroup', id: '200', text: '[chocolate cake recipe]' }]);
  const logRows = logs(ads).rowsFrom(2);
  assert.equal(logRows.length, 1);
  assert.equal(logRows[0][1], 'chocolate cake recipe');
  assert.equal(logRows[0][6], 'Live');
  assert.equal(logRows[0][7], 'Added');
  // Google's search_term_view status lags, so the term is still in the mock data: it must be filtered.
  assert.deepEqual(termsIn(output(ads)), ['free knitting pattern']);
});

test('full automate does not re-send a term ticked and applied earlier in the same run', () => {
  const ads = account();
  run(ads);
  output(ads).getRange(2, 1).setValue(true);
  run(ads, [['const FULL_AUTOMATE_MODE = false;', 'const FULL_AUTOMATE_MODE = true;']]);
  const texts = ads.entityNegatives.map(n => n.text).sort();
  assert.deepEqual(texts, ['[chocolate cake recipe]', '[free knitting pattern]']);
  assert.equal(ads.duplicateSends || 0, 0);
});

test('preview mode keeps ticked boxes so the next live run still applies them', () => {
  const ads = account({ preview: true });
  run(ads);
  output(ads).getRange(3, 1).setValue(true); // tick "free knitting pattern"
  run(ads);
  assert.equal(logs(ads).rowsFrom(2)[0][6], 'Preview');
  const ticked = output(ads).rowsFrom(2).filter(r => r[0] === true).map(r => r[1]);
  assert.deepEqual(ticked, ['free knitting pattern']);
  ads.preview = false;
  run(ads);
  assert.ok(logs(ads).rowsFrom(2).some(r => r[1] === 'free knitting pattern' && r[6] === 'Live' && r[7] === 'Added'));
});

test('empty account writes the no-results message and sends nothing', () => {
  const ads = new Account({ spreadsheetId: SHEET });
  run(ads);
  assert.equal(output(ads).getRange('A1').getValue(), 'No potential negatives found matching the criteria.');
  assert.equal(ads.entityNegatives.length, 0);
});

test('a failing search term report for one ad group does not abort the run', () => {
  const ads = account();
  ads.addAdGroup({ id: '201', name: 'Trail Shoes', campaignId: '100' });
  ads.addKeyword({ adGroupId: '201', text: 'trail shoes' });
  ads.failQuery("ad_group.id = '201' AND campaign.id", 'Internal error');
  const { errors } = run(ads);
  assert.ok(errors.some(e => e.includes('201')));
  assert.deepEqual(termsIn(output(ads)), ['chocolate cake recipe', 'free knitting pattern']);
});

test('bad sheet input: long terms, duplicates, message rows and blank IDs are handled per row', () => {
  const ads = account();
  run(ads);
  const sheet = output(ads);
  const long = 'one two three four five six seven eight nine ten eleven';
  sheet.getRange(4, 1, 1, 8).setValues([[true, long, '', '', 'Search - Shoes', 'Running Shoes', '100', '200']]);
  sheet.getRange(5, 1, 1, 8).setValues([[true, 'chocolate cake recipe', '', '', 'Search - Shoes', 'Running Shoes', 100, 200]]);
  sheet.getRange(2, 1).setValue(true);
  sheet.getRange(6, 1, 1, 8).setValues([[true, 'ghost term', '', '', 'x', 'y', '', '']]);
  sheet.getRange(7, 1).setValue(true); // ticked row with no search term
  run(ads);
  assert.deepEqual(ads.entityNegatives.map(n => n.text), ['[chocolate cake recipe]']);
  const status = Object.fromEntries(logs(ads).rowsFrom(2).map(r => [r[1], r[7]]));
  assert.equal(status[long], 'Skipped: over 80 characters or 10 words');
  assert.match(status['ghost term'], /not found/);
});
