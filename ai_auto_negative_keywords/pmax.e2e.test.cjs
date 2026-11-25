'use strict';
// End-to-end runs of the whole script (runAccount) against the mocked Google Ads Scripts
// runtime in ../test-support. Run with: node --test ai_auto_negative_keywords/
const assert = require('node:assert/strict');
const { join } = require('node:path');
const { test } = require('node:test');
const { MockGoogleAds, loadScript } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = join(__dirname, 'ai_auto_negative_keywords.js');
const SHEET = 'demo-sheet';
const API = 'https://autoneg.shabba.io/api/';

// Text conditions describe the RELEVANT terms: a term that fails them becomes a negative.
// Here relevant terms contain "shoes", so anything else is junk.
// A rule shaped like a live /api/rules response (v17), with overrides.
function rule(overrides = {}) {
  return {
    id: 101, rule_number: 1, name: 'Junk terms', level: 'campaign', lookback_days: 30,
    entity_conditions: [], performance_conditions: [],
    text_conditions: { sections: [{ operator: 'OR', conditions: [{ matchType: 'contains', text: 'shoes' }] }] },
    generate_report: true, auto_apply: true, send_email: false, negative_list_name: '', ai_prompt: '',
    max_search_terms: null, ai_list_all_search_terms: false, enabled: true, include_examples: false,
    search_term_status_added: true, search_term_status_excluded: true, search_term_status_added_excluded: true,
    automated_targeting_keywords: { enabled: false, threshold: 70 }, automated_targeting_ad_group_names: {},
    ...overrides,
  };
}

// An account with one Search campaign and two Performance Max campaigns.
function account(options = {}) {
  const ads = new MockGoogleAds({
    preview: options.preview,
    mutateAll: options.mutateAll,
    spreadsheetId: SHEET,
    fetch: {
      [`${API}rules?sheetId=${SHEET}`]: options.rules || [rule()],
      [`${API}script-info/latest-script-version`]: { version: 17 },
      [`${API}script-info/message`]: { message: '' },
      [`${API}example-terms`]: { exampleTerms: [] },
    },
  });
  ads.spreadsheet.insertSheet('Settings');
  ads.addCampaign({ id: '100', name: 'Search - Generic', channel: 'SEARCH' });
  ads.addAdGroup({ id: '200', name: 'Generic', campaignId: '100' });
  ads.addCampaign({ id: '500', name: 'PMax - Shoes', channel: 'PERFORMANCE_MAX' });
  ads.addCampaign({ id: '501', name: 'PMax - Brand', channel: 'PERFORMANCE_MAX' });
  ads.addSearchTerm({ term: 'cheap socks', campaignId: '100', adGroupId: '200', impressions: 40, clicks: 3, cost: 4.5 });
  ads.addSearchTerm({ term: 'buy shoes', campaignId: '100', adGroupId: '200', impressions: 90, clicks: 9, cost: 12 });
  // Two days of the same PMax term: Google adds them up when no date is selected.
  ads.addPmaxSearchTerm({ term: 'running socks', campaignId: '500', impressions: 30, clicks: 1, cost: 0.8, date: '2026-11-20' });
  ads.addPmaxSearchTerm({ term: 'running socks', campaignId: '500', impressions: 25, clicks: 2, cost: 1.1, date: '2026-11-21' });
  ads.addPmaxSearchTerm({ term: 'shoes sale', campaignId: '500', impressions: 300, clicks: 20, cost: 30 });
  ads.addPmaxSearchTerm({ term: 'brand socks', campaignId: '501', impressions: 5, clicks: 0, cost: 0 });
  ads.addPmaxCategory({ label: 'socks', campaignId: '500', impressions: 99 });
  return ads;
}

function run(ads, replace = []) {
  const script = loadScript(SCRIPT, ads.globals(), ['main', 'runAccount', 'SCRIPT_VERSION'], {
    replace: [['YOUR_SPREADSHEET_URL_HERE', `https://docs.google.com/spreadsheets/d/${SHEET}/edit`], ...replace],
  });
  script.exports.main();
  return script;
}

const pmaxCreates = ads => ads.mutations
  .map(m => ({ create: m.operation.campaignCriterionOperation.create, ok: m.result.isSuccessful() }));
const reportRows = ads => ads.spreadsheet.getSheetByName('1').rowsFrom(6);
const statusOf = (ads, term) => {
  const row = reportRows(ads).find(r => r[0] === term);
  return row && row.find(cell => /^(Yes|No|PMAX)/.test(String(cell)));
};

test('campaign rule: Search negatives on the Search campaign, PMax negatives on the PMax campaign', () => {
  const ads = account();
  run(ads);

  assert.deepEqual(ads.entityNegatives, [{ level: 'campaign', id: '100', text: '[cheap socks]' }]);
  const creates = pmaxCreates(ads);
  assert.deepEqual(creates.map(c => [c.create.campaign, c.create.keyword.text, c.create.keyword.matchType, c.create.negative, c.ok]), [
    ['customers/1234567890/campaigns/500', 'running socks', 'EXACT', true, true],
    ['customers/1234567890/campaigns/501', 'brand socks', 'EXACT', true, true],
  ]);
  assert.equal(statusOf(ads, 'running socks'), 'Yes');
  assert.equal(statusOf(ads, 'shoes sale'), undefined, 'a relevant term is not reported');
  // The aggregated cost of both days reaches the report.
  const row = reportRows(ads).find(r => r[0] === 'running socks');
  assert.ok(row.includes('1.90'), `cost 0.80 + 1.10 in ${row}`);
});

test('a second run does not resend PMax negatives that already exist', () => {
  const ads = account();
  run(ads);
  const first = ads.mutations.length;
  run(ads);
  assert.equal(ads.mutations.length, first, 'no new mutations on the second run');
  assert.equal(statusOf(ads, 'running socks'), 'No (already a negative)');
});

test('negatives added by hand in another match type do not block the exact negative', () => {
  const ads = account();
  ads.addCampaignNegative({ campaignId: '500', text: 'running socks', matchType: 'PHRASE' });
  run(ads);
  assert.ok(pmaxCreates(ads).some(c => c.create.keyword.text === 'running socks' && c.ok));
});

test('campaign rule with a negative list: PMax terms go to the list', () => {
  const ads = account({ rules: [rule({ negative_list_name: 'Junk' })] });
  ads.addNegativeList({ id: '77', name: 'Junk' });
  run(ads);
  assert.equal(ads.mutations.length, 0);
  assert.deepEqual(ads.lists[0].keywords.slice().sort(), ['[brand socks]', '[cheap socks]', '[running socks]']);
});

test('account rule with a negative list includes PMax terms', () => {
  const ads = account({ rules: [rule({ level: 'account', negative_list_name: 'Account junk' })] });
  ads.addNegativeList({ id: '78', name: 'Account junk' });
  run(ads);
  assert.deepEqual(ads.lists[0].keywords.slice().sort(), ['[brand socks]', '[cheap socks]', '[running socks]']);
});

test('account rule without a list adds nothing and does not crash', () => {
  const ads = account({ rules: [rule({ level: 'account' })] });
  const { errors } = run(ads);
  assert.equal(ads.mutations.length + ads.entityNegatives.length + ads.listAdds.length, 0);
  assert.ok(errors.some(e => e.includes('No negative list name')));
});

test('campaign name conditions limit which PMax campaigns are read', () => {
  const ads = account({ rules: [rule({ entity_conditions: [{ field: 'campaign.name', operator: 'contains', value: 'shoes' }] })] });
  run(ads);
  // LIKE ignores case, as in Google Ads: "shoes" matches "PMax - Shoes".
  assert.deepEqual(pmaxCreates(ads).map(c => c.create.campaign), ['customers/1234567890/campaigns/500']);
});

test('not_contains and regex campaign conditions apply to PMax too', () => {
  const ads = account({ rules: [rule({ entity_conditions: [
    { field: 'campaign.name', operator: 'not_contains', value: 'brand' },
    { field: 'campaign.name', operator: 'regex_contains', value: 'PMax' },
  ] })] });
  run(ads);
  assert.deepEqual(pmaxCreates(ads).map(c => c.create.campaign), ['customers/1234567890/campaigns/500']);
  assert.equal(ads.entityNegatives.length, 0, 'the Search campaign does not match the regex');
});

test('performance conditions apply to PMax terms', () => {
  const ads = account({ rules: [rule({ performance_conditions: [{ metric: 'impressions', operator: '>', value: '10' }] })] });
  run(ads);
  assert.deepEqual(pmaxCreates(ads).map(c => c.create.keyword.text), ['running socks']);
});

test('derived-metric conditions (cpc) apply to PMax terms', () => {
  const ads = account({ rules: [rule({ performance_conditions: [{ metric: 'cpc', operator: '>', value: '0.5' }] })] });
  run(ads);
  // running socks: 1.90 / 3 clicks = 0.63; brand socks has no clicks.
  assert.deepEqual(pmaxCreates(ads).map(c => c.create.keyword.text), ['running socks']);
});

test('auto-apply off: PMax terms are reported, nothing is changed', () => {
  const ads = account({ rules: [rule({ auto_apply: false })] });
  run(ads);
  assert.equal(ads.mutations.length + ads.entityNegatives.length, 0);
  assert.equal(statusOf(ads, 'running socks'), 'No (Auto Apply Disabled)');
});

test('preview mode sends the operations but Google keeps nothing', () => {
  const ads = account({ preview: true });
  run(ads);
  assert.equal(pmaxCreates(ads).length, 2);
  assert.equal((ads.rows.campaign_criterion || []).length, 0);
  assert.equal(statusOf(ads, 'running socks'), 'Yes (Previewed)');
});

test('PMAX_AUTO_APPLY = false keeps v16 behaviour (report only)', () => {
  const ads = account();
  run(ads, [['const PMAX_AUTO_APPLY = true;', 'const PMAX_AUTO_APPLY = false;']]);
  assert.equal(ads.mutations.length, 0);
  assert.equal(ads.entityNegatives.length, 1, 'Search negatives still apply');
  assert.equal(statusOf(ads, 'running socks'), 'No (PMAX_AUTO_APPLY is false)');
});

test('when the PMax search terms view fails, categories are reported and never negated', () => {
  const ads = account();
  ads.failQuery('campaign_search_term_view', 'Unrecognized resource');
  run(ads);
  assert.equal(ads.mutations.length, 0);
  assert.equal(statusOf(ads, 'socks'), 'No (PMax category, not a search term)');
  assert.ok(ads.queries.some(q => q.includes('FROM campaign_search_term_insight')));
});

test('a failed existing-negatives lookup still runs; Google refuses duplicates one by one', () => {
  const ads = account();
  ads.addCampaignNegative({ campaignId: '500', text: 'running socks' });
  ads.failQuery('FROM campaign_criterion', 'Quota');
  const { errors } = run(ads);
  const creates = pmaxCreates(ads);
  assert.equal(creates.length, 2);
  assert.deepEqual(creates.map(c => c.ok), [false, true]);
  assert.ok(errors.some(e => e.includes('DUPLICATE_CRITERION')));
});

test('terms over 10 words or 80 characters are reported but not sent', () => {
  const ads = account();
  ads.addPmaxSearchTerm({ term: 'socks for running in the rain on a very cold winter day', campaignId: '500', impressions: 3 });
  run(ads);
  assert.ok(pmaxCreates(ads).every(c => c.create.keyword.text.split(' ').length <= 10));
  assert.equal(statusOf(ads, 'socks for running in the rain on a very cold winter day'), 'No (over 10 words or 80 characters)');
});

test('a refused operation does not stop the others (partial failure)', () => {
  const ads = account();
  ads.addPmaxSearchTerm({ term: 'socks?', campaignId: '500', impressions: 3 });
  const { errors } = run(ads);
  const creates = pmaxCreates(ads);
  assert.ok(creates.some(c => c.create.keyword.text === 'socks?' && !c.ok));
  assert.ok(creates.some(c => c.create.keyword.text === 'running socks' && c.ok));
  assert.ok(errors.some(e => e.includes('INVALID_CHARACTERS')));
});

test('runtimes without mutateAll fall back to mutate', () => {
  const ads = account({ mutateAll: false });
  run(ads);
  assert.equal(pmaxCreates(ads).filter(c => c.ok).length, 2);
});

test('ad group rules never read PMax', () => {
  const ads = account({ rules: [rule({ level: 'adgroup' })] });
  run(ads);
  assert.ok(ads.queries.every(q => !q.includes('campaign_search_term')));
  assert.equal(ads.mutations.length, 0);
});

test('keyword protection rules never read PMax', () => {
  const ads = account({ rules: [rule({ automated_targeting_keywords: { enabled: true, scope: 'campaign', threshold: 70 } })] });
  ads.addRows('keyword_view', [{ 'campaign.id': '100', 'ad_group.id': '200', 'ad_group_criterion.keyword.text': 'shoes',
    'campaign.status': 'ENABLED', 'ad_group.status': 'ENABLED', 'ad_group_criterion.status': 'ENABLED', 'ad_group_criterion.negative': 'FALSE' }]);
  run(ads);
  assert.ok(ads.queries.every(q => !q.includes('campaign_search_term')));
  assert.equal(ads.mutations.length, 0);
});

test('a rule limited to Search campaigns by channel type skips PMax', () => {
  const ads = account({ rules: [rule({ entity_conditions: [{ field: 'campaign.advertising_channel_type', operator: 'equals', value: 'SEARCH' }] })] });
  run(ads);
  assert.equal(ads.mutations.length, 0);
});

test('every query the script sends parses as GAQL and uses compatible fields', () => {
  // The mock throws on fields from another resource, as Google does; a full run proves the shapes.
  const ads = account({ rules: [rule(), rule({ id: 102, rule_number: 2, level: 'account', negative_list_name: 'L' }), rule({ id: 103, rule_number: 3, level: 'adgroup' })] });
  ads.addNegativeList({ id: '79', name: 'L' });
  const { errors } = run(ads);
  assert.deepEqual(errors.filter(e => /Unrecognized field|parsing GAQL/.test(e)), []);
  assert.ok(ads.queries.some(q => q.includes('FROM campaign_search_term_view') && q.includes('campaign.id = 500')));
});
