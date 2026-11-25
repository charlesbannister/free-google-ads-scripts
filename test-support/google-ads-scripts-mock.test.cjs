'use strict';
// Checks that the mock behaves like Google Ads where the scripts depend on it.
const assert = require('node:assert/strict');
const { test } = require('node:test');
const { MockGoogleAds, parseGaql, likeToRegExp } = require('./google-ads-scripts-mock.cjs');

test('GAQL parser reads every clause the scripts use', () => {
  const q = parseGaql(`SELECT campaign.id, metrics.clicks FROM campaign
    WHERE segments.date BETWEEN '2026-01-01' AND '2026-01-31' AND campaign.status = ENABLED
    AND campaign.name LIKE '%a[_]b%' AND campaign.id IN (1, 2) AND NOT campaign.name REGEXP_MATCH 'x'
    AND metrics.clicks > 3 ORDER BY metrics.clicks DESC LIMIT 5`);
  assert.equal(q.resource, 'campaign');
  assert.deepEqual(q.where.map(c => c.kind), ['between', 'compare', 'like', 'in', 'not_regex', 'compare']);
  assert.equal(q.limit, 5);
  assert.throws(() => parseGaql('SELECT x FROM y WHERE a = 1 OR b = 2'), /parsing GAQL/);
});

test('LIKE ignores case and treats [x] as a literal', () => {
  assert.ok(likeToRegExp('%brand%').test('PMax - Brand'));
  assert.ok(likeToRegExp('%a[_]b%').test('xa_by'));
  assert.ok(!likeToRegExp('%a[_]b%').test('xacby'));
});

test('fields from another resource are refused, like Google', () => {
  const ads = new MockGoogleAds();
  assert.throws(() => ads.report('SELECT ad_group.id FROM campaign_search_term_view'), /Unrecognized field/);
  assert.throws(() => ads.report('SELECT x.y FROM nope'), /Invalid resource/);
});

test('rows are added up when no segment is selected', () => {
  const ads = new MockGoogleAds();
  ads.addCampaign({ id: 5, name: 'P', channel: 'PERFORMANCE_MAX' });
  ads.addPmaxSearchTerm({ term: 't', campaignId: 5, clicks: 1, date: '2026-01-01' });
  ads.addPmaxSearchTerm({ term: 't', campaignId: 5, clicks: 2, date: '2026-01-02' });
  const it = ads.report("SELECT campaign_search_term_view.search_term, metrics.clicks FROM campaign_search_term_view WHERE segments.date BETWEEN '2026-01-01' AND '2026-01-31'").rows();
  assert.equal(it.totalNumEntities(), 1);
  assert.equal(it.next()['metrics.clicks'], '3');
  const byDay = ads.report("SELECT segments.date, metrics.clicks FROM campaign_search_term_view WHERE campaign.id = 5").rows();
  assert.equal(byDay.totalNumEntities(), 2);
});

test('mutateAll without partial failure is all or nothing', () => {
  const ads = new MockGoogleAds({ customerId: '111-222-3333' });
  ads.addCampaign({ id: 5, name: 'P', channel: 'PERFORMANCE_MAX' });
  const op = text => ({ campaignCriterionOperation: { create: { campaign: 'customers/1112223333/campaigns/5', negative: true, keyword: { text, matchType: 'EXACT' } } } });
  const results = ads.mutateAll([op('good'), op('bad?')]);
  assert.deepEqual(results.map(r => r.isSuccessful()), [false, false]);
  assert.equal((ads.rows.campaign_criterion || []).length, 0);
  const partial = ads.mutateAll([op('good'), op('bad?')], { partialFailure: true });
  assert.deepEqual(partial.map(r => r.isSuccessful()), [true, false]);
  assert.equal(ads.mutate(op('good')).getErrorMessages()[0].startsWith('DUPLICATE_CRITERION'), true);
});

test('AdsApp.campaigns() leaves out Shopping and Performance Max', () => {
  const ads = new MockGoogleAds();
  ads.addCampaign({ id: 1, name: 'S', channel: 'SEARCH' });
  ads.addCampaign({ id: 2, name: 'P', channel: 'PERFORMANCE_MAX' });
  const { AdsApp } = ads.globals();
  assert.equal(AdsApp.campaigns().withIds([2]).get().hasNext(), false);
  assert.equal(AdsApp.performanceMaxCampaigns().withIds([2]).get().next().getName(), 'P');
});

test('sheet setValues must match the range size, like Sheets', () => {
  const sheet = new MockGoogleAds().spreadsheet.insertSheet('1');
  assert.throws(() => sheet.getRange(1, 1, 2, 2).setValues([[1, 2]]), /does not match/);
  sheet.getRange(5, 1, 1, 2).setValues([['a', 'b']]);
  assert.deepEqual(sheet.rowsFrom(5), [['a', 'b']]);
});
