'use strict';
// End-to-end runs of the whole script (main) against the mocked Google Ads Scripts runtime
// in ../test-support. Run with: node --test auto_negative_keywords_sheets/
//
// The script still reads search terms and existing negatives with AWQL reports
// (SEARCH_QUERY_PERFORMANCE_REPORT etc.) and legacy selector conditions (CampaignId = ...).
// The shared mock only speaks GAQL, so this file adds a small AWQL layer on top of it.
const assert = require('node:assert/strict');
const { join } = require('node:path');
const { test } = require('node:test');
const { MockGoogleAds, loadScript, iterator } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = join(__dirname, 'auto_negative_keywords_sheets.js');
const SHEET_ID = 'demo-sheet';

// ------------------------------------------------------------------ AWQL + entity extensions

function awqlValue(query, field) {
  const m = query.match(new RegExp(`${field}\\s*=\\s*["']([^"']*)["']`, 'i'));
  return m ? m[1] : null;
}

function awqlNumber(query, field, op) {
  const m = query.match(new RegExp(`${field}\\s*${op}\\s*([\\d.]+)`, 'i'));
  return m ? Number(m[1]) : null;
}

class SheetsAds extends MockGoogleAds {
  constructor(options) {
    super(options);
    this.listCampaigns = {};   // list id -> [campaign ids]
    this.createdLists = [];
  }

  report(query) {
    if (!/_REPORT\b/.test(query)) return super.report(query);
    this.queries.push(String(query).replace(/\s+/g, ' ').trim());
    for (const f of this.failures) {
      if (f.match instanceof RegExp ? f.match.test(query) : query.includes(f.match)) throw new Error(f.message);
    }
    const campaignId = awqlValue(query, 'CampaignId');
    const adGroupId = awqlValue(query, 'AdGroupId');
    let rows = [];
    if (/SEARCH_QUERY_PERFORMANCE_REPORT/.test(query)) {
      const clicks = awqlNumber(query, 'Clicks', '>');
      const impressions = awqlNumber(query, 'Impressions', '>');
      rows = (this.rows.search_term_view || [])
        .filter(r => !campaignId || r['campaign.id'] === campaignId)
        .filter(r => !adGroupId || r['ad_group.id'] === adGroupId)
        .filter(r => clicks === null || Number(r['metrics.clicks']) > clicks)
        .filter(r => impressions === null || Number(r['metrics.impressions']) > impressions)
        .map(r => ({ Query: r['search_term_view.search_term'] }));
    } else if (/CAMPAIGN_NEGATIVE_KEYWORDS_PERFORMANCE_REPORT/.test(query)) {
      rows = (this.rows.campaign_criterion || []).filter(r => r['campaign.id'] === campaignId)
        .map(r => ({ Criteria: r['campaign_criterion.keyword.text'], KeywordMatchType: r['campaign_criterion.keyword.match_type'] }));
    } else if (/KEYWORDS_PERFORMANCE_REPORT/.test(query)) {
      rows = this.entityNegatives.filter(n => n.level === 'adgroup' && n.id === adGroupId)
        .map(n => ({ Criteria: n.text.replace(/^\[|\]$/g, ''), KeywordMatchType: 'Exact' }));
    } else {
      throw new Error(`AWQL report not mocked: ${query}`);
    }
    return { rows: () => iterator(rows) };
  }

  globals(extra) {
    const g = super.globals(extra);
    const ads = this;
    const AdsApp = g.AdsApp;
    // Legacy selector names -> GAQL names.
    const legacy = selectorFn => () => {
      const selector = selectorFn();
      const withCondition = selector.withCondition;
      selector.withCondition = c => withCondition(String(c).replace(/\bCampaignId\b/g, 'campaign.id'));
      const get = selector.get;
      selector.get = () => {
        const it = get();
        const items = [];
        while (it.hasNext()) items.push(it.next());
        return iterator(items.map(entity => ({
          ...entity,
          negativeKeywordLists: () => ({
            get: () => iterator(ads.lists.filter(l => (ads.listCampaigns[l.id] || []).includes(String(entity.getId())))
              .map(l => ads._listEntity(l))),
          }),
          addNegativeKeywordList: list => {
            (ads.listCampaigns[String(list.getId())] ||= []).push(String(entity.getId()));
          },
        })));
      };
      return selector;
    };
    for (const name of ['campaigns', 'shoppingCampaigns', 'adGroups', 'shoppingAdGroups']) AdsApp[name] = legacy(AdsApp[name]);
    AdsApp.negativeKeywordLists = () => ({
      get: () => iterator(ads.lists.filter(l => l.type === 'NEGATIVE_KEYWORDS').map(l => ads._listEntity(l))),
    });
    AdsApp.newNegativeKeywordListBuilder = () => {
      let name = '';
      const builder = {
        withName(n) { name = n; return builder; },
        build() {
          ads.createdLists.push(name);
          if (ads.preview) return { getResult: () => null, isSuccessful: () => true };
          const list = ads.addNegativeList({ id: String(800 + ads.lists.length), name });
          return { getResult: () => ads._listEntity(list), isSuccessful: () => true };
        },
      };
      return builder;
    };
    g.SpreadsheetApp.create = () => { throw new Error('No output spreadsheet in these tests'); };
    return g;
  }

  _listEntity(list) {
    const ads = this;
    return {
      getId: () => Number(list.id),
      getName: () => list.name,
      addNegativeKeywords: texts => ads._addToList(list, texts),
      negativeKeywords: () => ({
        get: () => iterator(list.keywords.map(k => ({
          getText: () => k.replace(/^\[|\]$|^"|"$/g, ''),
          getMatchType: () => (k.startsWith('[') ? 'EXACT' : k.startsWith('"') ? 'PHRASE' : 'BROAD'),
        }))),
      }),
      campaigns: () => ({
        get: () => iterator((ads.listCampaigns[list.id] || []).map(id => ({
          addNegativeKeywordList: l => (ads.listCampaigns[String(l.getId())] ||= []).push(id),
        }))),
      }),
    };
  }
}

// ------------------------------------------------------------------ the rules sheet

const ROW = { SKIP: 11, LOG: 12, OUTPUT: 13, MATCHES: 14, CAMPAIGN: 15, CAMPAIGN_ID: 16, AD_GROUP: 17,
  AD_GROUP_ID: 18, LIST: 19, ADD_TO: 20, APPROX: 21, APPROX_T: 22, KEYWORD: 23 };

/** Writes a rules sheet in the template's layout. rules: [{ campaign, adGroup, addTo, list, keywords }] */
function rulesSheet(ads, { name = 'Rules', settings = {}, rules }) {
  const sheet = ads.spreadsheet.insertSheet(name);
  const set = (a1, v) => sheet.getRange(a1).setValue(v);
  set('A3', 'Account ID (MCC Only)');
  const s = {
    A4: '', B4: '', C4: '', D4: '', E4: 30, F4: 'AdGroup/Campaign', G4: '', H4: '', I4: '',
    J4: '', K4: '', L4: '', M4: false,
    A7: 'me@example.com', B7: true, C7: false, D7: true, E7: false, F7: '', G7: false, H7: 'Exact', I7: '', J7: '',
    ...settings,
  };
  for (const [a1, v] of Object.entries(s)) set(a1, v);
  rules.forEach((rule, i) => {
    const col = 4 + i;
    const put = (row, v) => sheet.getRange(row, col).setValue(v);
    put(ROW.MATCHES, rule.matches ?? 1);
    put(ROW.CAMPAIGN, rule.campaign || '');
    put(ROW.CAMPAIGN_ID, rule.campaignId || '');
    put(ROW.AD_GROUP, rule.adGroup || '');
    put(ROW.LIST, rule.list || '');
    put(ROW.ADD_TO, rule.addTo || 'Campaign');
    (rule.keywords || []).forEach((k, j) => put(ROW.KEYWORD + j, k));
  });
  return sheet;
}

/** A Search campaign, a Shopping-free account, and a Performance Max campaign. */
function account(options = {}) {
  const ads = new SheetsAds({ preview: options.preview, spreadsheetId: SHEET_ID });
  ads.addCampaign({ id: '100', name: 'Search - Generic', channel: 'SEARCH' });
  ads.addAdGroup({ id: '200', name: 'Generic', campaignId: '100' });
  ads.addRows('ad_group', [{ 'ad_group.id': '200', 'ad_group.name': 'Generic', 'ad_group.status': 'ENABLED', 'campaign.id': '100' }]);
  ads.addCampaign({ id: '500', name: 'PMax - Shoes', channel: 'PERFORMANCE_MAX' });
  ads.addSearchTerm({ term: 'cheap socks', campaignId: '100', adGroupId: '200', impressions: 40, clicks: 3 });
  ads.addSearchTerm({ term: 'buy shoes', campaignId: '100', adGroupId: '200', impressions: 90, clicks: 9 });
  ads.addSearchTerm({ term: 'one two three four five six seven eight nine ten eleven', campaignId: '100', adGroupId: '200', clicks: 5 });
  ads.addPmaxSearchTerm({ term: 'running socks', campaignId: '500', impressions: 30, clicks: 1, date: '2026-11-20' });
  ads.addPmaxSearchTerm({ term: 'running socks', campaignId: '500', impressions: 25, clicks: 2, date: '2026-11-21' });
  ads.addPmaxSearchTerm({ term: 'shoes sale', campaignId: '500', impressions: 300, clicks: 20 });
  return ads;
}

function run(ads, replace = []) {
  const script = loadScript(SCRIPT, ads.globals(), ['main', 'VERSION_NUMBER', 'PMAX_SUPPORT'], {
    replace: [['YOUR_SPREADSHEET_URL_HERE', `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit`], ...replace],
  });
  let error = null;
  try { script.exports.main(); } catch (e) { error = e; }
  return { ...script, error };
}

const pmaxTexts = ads => ads.mutations.filter(m => m.result.isSuccessful())
  .map(m => [m.operation.campaignCriterionOperation.create.campaign, m.operation.campaignCriterionOperation.create.keyword.text,
    m.operation.campaignCriterionOperation.create.keyword.matchType]);
const logCell = (ads, col = 4, name = 'Rules') => ads.spreadsheet.getSheetByName(name).getRange(9, col).getValue();

// ------------------------------------------------------------------ tests

test('campaign rule on a Search campaign adds campaign negatives and emails once', () => {
  const ads = account();
  rulesSheet(ads, { rules: [{ campaign: 'Search - Generic', keywords: ['shoes'] }] });
  const { error } = run(ads);
  assert.equal(error, null);
  assert.deepEqual(ads.entityNegatives, [{ level: 'campaign', id: '100', text: '[cheap socks]' }]);
  assert.equal(logCell(ads), '1 new negatives added.');
  assert.equal(ads.emails.length, 1);
  assert.match(ads.emails[0][0].subject, /^1 negatives were added - Rules sheet/);
});

test('search terms over 10 words are not sent (Google refuses them)', () => {
  const ads = account();
  rulesSheet(ads, { settings: { J7: 20 }, rules: [{ campaign: 'Search - Generic', keywords: ['shoes'] }] });
  run(ads);
  assert.ok(!ads.entityNegatives.some(n => n.text.includes('eleven')));
});

test('a second run does not resend campaign negatives (New negatives only)', () => {
  const ads = account();
  rulesSheet(ads, { rules: [{ campaign: 'Search - Generic', keywords: ['shoes'] }] });
  run(ads);
  run(ads);
  assert.equal(ads.entityNegatives.length, 1);
});

test('ad group rule adds ad group negatives', () => {
  const ads = account();
  rulesSheet(ads, { rules: [{ campaign: 'Search - Generic', adGroup: 'Generic', addTo: 'Ad Group', keywords: ['shoes'] }] });
  const { error } = run(ads);
  assert.equal(error, null);
  assert.deepEqual(ads.entityNegatives, [{ level: 'adgroup', id: '200', text: '[cheap socks]' }]);
});

test('Performance Max campaign rule: reads campaign_search_term_view and adds negatives by mutate', () => {
  const ads = account();
  rulesSheet(ads, { rules: [{ campaign: 'PMax - Shoes', keywords: ['shoes'] }] });
  const { error } = run(ads);
  assert.equal(error, null);
  assert.deepEqual(pmaxTexts(ads), [['customers/1234567890/campaigns/500', 'running socks', 'EXACT']]);
  assert.equal(ads.entityNegatives.length, 0);
  assert.equal(logCell(ads), '1 new negatives added.');
  assert.ok(ads.queries.some(q => q.includes('FROM campaign_search_term_view') && q.includes('campaign.id = 500')));
});

test('Performance Max: settings filters apply to the date-range totals', () => {
  const ads = account();
  // running socks has 3 clicks over two days (1 + 2): more than 2 only when added up.
  rulesSheet(ads, { settings: { B4: 2 }, rules: [{ campaign: 'PMax - Shoes', keywords: ['shoes'] }] });
  run(ads);
  assert.deepEqual(pmaxTexts(ads).map(t => t[1]), ['running socks']);
});

test('Performance Max: a second run does not resend existing negatives, even with New negatives only off', () => {
  const ads = account();
  rulesSheet(ads, { settings: { D7: false }, rules: [{ campaign: 'PMax - Shoes', keywords: ['shoes'] }] });
  run(ads);
  const sent = ads.mutations.length;
  run(ads);
  assert.equal(ads.mutations.length, sent);
});

test('Performance Max support can be turned off with PMAX_SUPPORT', () => {
  const ads = account();
  rulesSheet(ads, { rules: [{ campaign: 'PMax - Shoes', keywords: ['shoes'] }] });
  const { error } = run(ads, [['const PMAX_SUPPORT = true;', 'const PMAX_SUPPORT = false;']]);
  assert.equal(error, null);
  assert.equal(ads.mutations.length, 0);
  assert.ok(!ads.queries.some(q => q.includes('campaign_search_term_view')));
});

test('Preview Mode in the sheet: nothing is added, no list is created, but the sheet shows what was found', () => {
  const ads = account();
  ads.addNegativeList({ id: '77', name: 'Junk' });
  rulesSheet(ads, {
    settings: { C7: true, G7: true },
    rules: [
      { campaign: 'Search - Generic', keywords: ['shoes'] },
      { campaign: 'PMax - Shoes', keywords: ['shoes'] },
    ],
  });
  // Make the Junk list full so the script would create 'Junk 1'.
  ads.lists[0].keywords = Array.from({ length: 5000 }, (_, i) => `[t${i}]`);
  ads.spreadsheet.getSheetByName('Rules').getRange(ROW.LIST, 4).setValue('Junk');
  ads.spreadsheet.getSheetByName('Rules').getRange(ROW.ADD_TO, 4).setValue('List (Add Name)');
  const { error } = run(ads);
  assert.equal(error, null);
  assert.equal(ads.entityNegatives.length + ads.mutations.length + ads.listAdds.length, 0);
  assert.deepEqual(ads.createdLists, []);
  assert.equal(logCell(ads, 4), '1 new negatives added.');
  assert.equal(logCell(ads, 5), '1 new negatives added.');
  assert.match(ads.emails[0][0].subject, /were found/);
});

test('Google Ads preview: negatives go through the normal calls (Google ignores them)', () => {
  const ads = account({ preview: true });
  rulesSheet(ads, { rules: [{ campaign: 'Search - Generic', keywords: ['shoes'] }, { campaign: 'PMax - Shoes', keywords: ['shoes'] }] });
  const { error } = run(ads);
  assert.equal(error, null);
  assert.equal(ads.entityNegatives.length, 1);
  assert.equal(ads.mutations.length, 1);
});

test('list rule: a list that fills up overflows into "<name> 1" instead of being overfilled', () => {
  const ads = account();
  ads.addSearchTerm({ term: 'red socks', campaignId: '100', adGroupId: '200', clicks: 2 });
  const junk = ads.addNegativeList({ id: '77', name: 'Junk' });
  junk.keywords = Array.from({ length: 4999 }, (_, i) => `[t${i}]`);
  ads.listCampaigns['77'] = ['100'];
  rulesSheet(ads, { settings: { G7: true, D7: false }, rules: [{ campaign: 'Search - Generic', addTo: 'List (Add Name)', list: 'Junk', keywords: ['shoes'] }] });
  const { error } = run(ads);
  assert.equal(error, null);
  assert.equal(junk.keywords.length, 5000);
  assert.deepEqual(ads.createdLists, ['Junk 1']);
  const junk1 = ads.lists.find(l => l.name === 'Junk 1');
  assert.equal(junk1.keywords.length, 1);
  assert.deepEqual(ads.listCampaigns[junk1.id], ['100'], 'the new list is applied to the same campaigns');
});

test('empty account: no search terms, no changes, no email, no error', () => {
  const ads = new SheetsAds({ spreadsheetId: SHEET_ID });
  ads.addCampaign({ id: '100', name: 'Search - Generic', channel: 'SEARCH' });
  rulesSheet(ads, { rules: [{ campaign: 'Search - Generic', keywords: ['shoes'] }] });
  const { error } = run(ads);
  assert.equal(error, null);
  assert.equal(ads.emails.length, 0);
  assert.equal(logCell(ads), '0 new negatives added.');
});

test('report failure on one rule: logged in its column, other rules still run', () => {
  const ads = account();
  ads.failQuery('CampaignId = "100"', 'AWQL report failed');
  rulesSheet(ads, { rules: [{ campaign: 'Search - Generic', keywords: ['shoes'] }, { campaign: 'PMax - Shoes', keywords: ['shoes'] }] });
  const { error } = run(ads);
  assert.equal(error, null);
  assert.equal(logCell(ads, 4), 'Error. Rule skipped.');
  assert.equal(pmaxTexts(ads).length, 1);
});

test('one broken sheet does not stop the other sheets, and the run still reports the failure', () => {
  const ads = account();
  rulesSheet(ads, { name: 'Broken', rules: [{ campaign: 'Search - Generic', addTo: 'List (Add Name)', list: 'Missing list', keywords: ['shoes'] }] });
  rulesSheet(ads, { name: 'Good', rules: [{ campaign: 'PMax - Shoes', keywords: ['shoes'] }] });
  const { error } = run(ads);
  assert.match(String(error && error.message), /1 sheet\(s\) failed\. Broken: Negative list name 'Missing list' not found/);
  assert.equal(pmaxTexts(ads).length, 1);
});

test('bad sheet input: a numeric "contains" value and no keywords do not crash the run', () => {
  const ads = account();
  rulesSheet(ads, { settings: { G4: 2024 }, rules: [{ campaign: 'Search - Generic', keywords: [] }] });
  const { error } = run(ads);
  assert.equal(error, null);
  assert.match(logCell(ads), /Please add/);
});

test('the version is bumped', () => {
  const ads = account();
  rulesSheet(ads, { rules: [{ campaign: 'Search - Generic', keywords: ['shoes'] }] });
  const { exports } = run(ads);
  assert.equal(exports.VERSION_NUMBER, '4.3.0');
  assert.equal(ads.spreadsheet.getSheetByName('Rules').getRange('B1').getValue(), 'Version: 4.3.0');
});
