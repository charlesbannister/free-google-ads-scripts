'use strict';
// End-to-end runs of the whole script (main) against the mocked Google Ads Scripts runtime
// in ../test-support. Run with: node --test placement_exclusions/
const assert = require('node:assert/strict');
const { join } = require('node:path');
const { test } = require('node:test');
const { MockGoogleAds, MockSheet, loadScript, parseGaql, iterator, likeToRegExp } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = join(__dirname, 'pmax_placement_exclusions.js');
const SHEET = 'demo-sheet';
const LIST = 'Bad placements';

// Sheet calls this script makes that the shared mock does not have yet (formatting only).
MockSheet.prototype.setTabColor ??= function () { return this; };
MockSheet.prototype.setFrozenColumns ??= function () { return this; };
Object.getPrototypeOf(new MockSheet('x', 1).getRange(1, 1)).insertCheckboxes ??= function () { return this; };

// Fields each placement view accepts. Performance Max only has impressions.
const VIEW_FIELDS = {
  performance_max_placement_view: f => f.startsWith('campaign.') || f.startsWith('performance_max_placement_view.') ||
    f === 'metrics.impressions' || f === 'segments.date',
  detail_placement_view: f => f.startsWith('campaign.') || f.startsWith('detail_placement_view.') ||
    f.startsWith('metrics.') || f === 'segments.date',
};

function condMatches(row, c) {
  const v = row[c.field];
  if (c.field === 'segments.date') return true;
  if (c.kind === 'in') { const hit = c.values.map(String).includes(String(v)); return c.not ? !hit : hit; }
  if (c.kind === 'like') { const hit = likeToRegExp(c.pattern).test(String(v ?? '')); return c.not ? !hit : hit; }
  if (c.kind === 'compare') {
    if (c.op === '=') return String(v) === String(c.value);
    if (c.op === '!=') return String(v) !== String(c.value);
    return { '>': Number(v) > Number(c.value), '<': Number(v) < Number(c.value), '>=': Number(v) >= Number(c.value), '<=': Number(v) <= Number(c.value) }[c.op];
  }
  return true;
}

/** MockGoogleAds plus placement views, placement exclusion lists, Video campaigns and bulk uploads. */
class PlacementAds extends MockGoogleAds {
  constructor(options = {}) {
    super(options);
    this.placementRows = { performance_max_placement_view: [], detail_placement_view: [] };
    this.exclusionLists = [];   // { id, name, placements: [] }
    this.links = [];            // { campaignId, list }
    this.bulkUploads = [];
    this.videoExclusions = [];
  }

  addPlacement(resource, row) {
    const campaign = this.campaigns.find(c => c.id === String(row['campaign.id'])) || {};
    this.placementRows[resource].push({
      'campaign.name': campaign.name, 'campaign.status': campaign.status || 'ENABLED',
      'campaign.advertising_channel_type': campaign.channel, ...row,
    });
  }

  report(query) {
    const parsed = parseGaql(query);
    if (!this.placementRows[parsed.resource]) return super.report(query);
    this.queries.push(String(query).replace(/\s+/g, ' ').trim());
    for (const f of this.failures) if (parsed.resource === f.match) throw new Error(f.message);
    for (const field of [...parsed.fields, ...parsed.where.map(c => c.field)]) {
      if (!VIEW_FIELDS[parsed.resource](field)) throw new Error(`Unrecognized field '${field}' for ${parsed.resource}`);
    }
    if (/cost_micros\s*>\s*\d+\.\d/.test(query)) throw new Error('Invalid int64 value for metrics.cost_micros');
    let rows = this.placementRows[parsed.resource].filter(r => parsed.where.every(c => condMatches(r, c)));
    if (parsed.limit !== null) rows = rows.slice(0, parsed.limit);
    const out = rows.map(r => Object.fromEntries(parsed.fields.map(f => [f, r[f] === undefined ? (f.startsWith('metrics.') ? '0' : '') : r[f]])));
    return { rows: () => iterator(out) };
  }

  addExclusionList(name) {
    const list = { id: String(800 + this.exclusionLists.length), name, placements: [] };
    this.exclusionLists.push(list);
    return list;
  }

  _exclusionListEntity(list) {
    const ads = this;
    return {
      getId: () => Number(list.id),
      getName: () => list.name,
      addExcludedPlacements(urls) { if (!ads.preview) list.placements.push(...urls); },
      addExcludedPlacement(url) { if (!ads.preview) list.placements.push(url); },
      excludedPlacements: () => ({ get: () => iterator(list.placements.map(u => ({ getUrl: () => u }))) }),
    };
  }

  globals() {
    const g = super.globals();
    const ads = this;
    const listSelector = () => {
      let name = null;
      const sel = {
        withCondition(c) { const m = String(c).match(/=\s*'(.*)'/); if (m) name = m[1]; return sel; },
        get: () => iterator(ads.exclusionLists.filter(l => name === null || l.name === name).map(l => ads._exclusionListEntity(l))),
      };
      return sel;
    };
    // Campaign entities: real Google Ads exposes excludedPlacementLists() as a selector.
    const withLists = selectorFactory => () => {
      const sel = selectorFactory();
      const get = sel.get;
      sel.get = () => {
        const it = get();
        const items = [];
        while (it.hasNext()) {
          const c = it.next();
          const id = String(c.getId());
          c.excludedPlacementLists = () => ({
            get: () => iterator(ads.links.filter(l => l.campaignId === id).map(l => ads._exclusionListEntity(l.list))),
          });
          c.addExcludedPlacementList = entity => {
            const list = ads.exclusionLists.find(l => l.id === String(entity.getId()));
            if (!ads.preview) ads.links.push({ campaignId: id, list });
          };
          c.videoTargeting = () => ({
            newYouTubeVideoBuilder: () => ({
              withVideoId: videoId => ({
                exclude: () => { ads.videoExclusions.push({ campaignId: id, videoId }); return { isSuccessful: () => true, getErrors: () => [] }; },
              }),
            }),
          });
          items.push(c);
        }
        return iterator(items);
      };
      return sel;
    };
    g.AdsApp.excludedPlacementLists = listSelector;
    g.AdsApp.performanceMaxCampaigns = withLists(g.AdsApp.performanceMaxCampaigns);
    g.AdsApp.campaigns = withLists(g.AdsApp.campaigns);
    g.AdsApp.videoCampaigns = withLists(() => ads._campaignSelector(['VIDEO']));
    g.AdsApp.bulkUploads = () => ({
      newCsvUpload: columns => {
        const upload = { columns, rows: [], mode: null };
        ads.bulkUploads.push(upload);
        return {
          forCampaignManagement() {},
          append: row => upload.rows.push(row),
          preview: () => { upload.mode = 'preview'; },
          apply: () => { upload.mode = 'apply'; },
        };
      },
    });
    return g;
  }
}

// A realistic account: PMax, Display and Video campaigns, with placements of every type.
function account(options = {}) {
  const ads = new PlacementAds({ preview: options.preview, spreadsheetId: SHEET });
  ads.addCampaign({ id: '500', name: 'PMax - Shoes', channel: 'PERFORMANCE_MAX' });
  ads.addCampaign({ id: '600', name: 'Display - Remarketing', channel: 'DISPLAY' });
  ads.addCampaign({ id: '700', name: 'YouTube - Brand', channel: 'VIDEO' });
  if (options.list !== false) ads.addExclusionList(LIST);

  const settings = ads.spreadsheet.insertSheet('Settings');
  const s = (r, c, v) => settings.getRange(r, c).setValue(v);
  s(4, 1, 'Shared Exclusion List Name'); s(4, 2, LIST);
  s(5, 1, 'Lookback Window (Days)'); s(5, 2, 30);
  s(6, 1, 'Minimum Impressions'); s(6, 2, options.minImpressions ?? 0);
  s(7, 1, 'Minimum Clicks'); s(7, 2, 0);
  s(8, 1, 'Minimum Cost'); s(8, 2, options.minCost ?? 0);
  s(9, 1, 'Maximum Conversions'); s(9, 2, 0);
  s(10, 1, 'Max Results'); s(10, 2, options.maxResults ?? 0);
  s(11, 1, 'Campaign Name Contains'); s(11, 2, '');
  s(12, 1, 'Campaign Name Not Contains'); s(12, 2, '');
  s(13, 1, 'Enabled campaigns only'); s(13, 2, true);
  s(14, 1, 'Active placements only'); s(14, 2, options.activeOnly ?? false);
  for (const row of [18, 19, 20, 21]) { s(row, 2, true); s(row, 3, false); }
  ads.spreadsheet.insertSheet('Settings: Placement Filters');
  ads.spreadsheet.insertSheet('Settings: ChatGPT');
  const log = ads.spreadsheet.insertSheet('Exclusion Log');
  log.getRange(1, 1, 1, 9).setValues([['Status', 'Campaign Name', 'Placement', 'Display Name', 'Placement Type', 'Target URL', 'Notes', 'Preview Mode', 'Timestamp']]);

  if (options.empty) return ads;
  ads.addPlacement('performance_max_placement_view', {
    'campaign.id': '500', 'performance_max_placement_view.placement': 'spammy-site.xyz',
    'performance_max_placement_view.display_name': 'Spammy', 'performance_max_placement_view.placement_type': 'WEBSITE',
    'performance_max_placement_view.target_url': 'spammy-site.xyz', 'metrics.impressions': '900',
  });
  ads.addPlacement('performance_max_placement_view', {
    'campaign.id': '500', 'performance_max_placement_view.placement': 'mobileapp::2-com.bad.game',
    'performance_max_placement_view.display_name': 'Bad Game', 'performance_max_placement_view.placement_type': 'MOBILE_APPLICATION',
    'performance_max_placement_view.target_url': '', 'metrics.impressions': '400',
  });
  ads.addPlacement('detail_placement_view', {
    'campaign.id': '600', 'detail_placement_view.placement': 'news-site.com', 'detail_placement_view.display_name': 'News',
    'detail_placement_view.placement_type': 'WEBSITE', 'detail_placement_view.target_url': 'news-site.com',
    // Conversion value with no cost: ROAS must not be Infinity.
    'metrics.impressions': '300', 'metrics.clicks': '3', 'metrics.cost_micros': '0', 'metrics.conversions': '1', 'metrics.conversions_value': '50',
  });
  ads.addPlacement('detail_placement_view', {
    'campaign.id': '700', 'detail_placement_view.placement': 'K2zZwB3NY04', 'detail_placement_view.display_name': 'Some video',
    'detail_placement_view.placement_type': 'YOUTUBE_VIDEO', 'detail_placement_view.target_url': 'youtube.com/video/K2zZwB3NY04',
    'metrics.impressions': '200', 'metrics.clicks': '1', 'metrics.cost_micros': '2500000', 'metrics.conversions': '0', 'metrics.conversions_value': '0',
  });
  return ads;
}

function run(ads, replace = []) {
  const script = loadScript(SCRIPT, ads.globals(), ['main'], {
    replace: [['YOUR_SPREADSHEET_URL_HERE', `https://docs.google.com/spreadsheets/d/${SHEET}/edit`], ...replace],
  });
  script.exports.main();
  return script;
}

const out = (ads, name) => ads.spreadsheet.getSheetByName(name).rowsFrom(4);
const header = (ads, name) => ads.spreadsheet.getSheetByName(name).rowsFrom(3)[0];

/** Ticks the Exclude box on the output-sheet row whose Placement matches. */
function tick(ads, sheetName, placement) {
  const sheet = ads.spreadsheet.getSheetByName(sheetName);
  const cols = header(ads, sheetName);
  const rowIndex = out(ads, sheetName).findIndex(r => r[cols.indexOf('Placement')] === placement);
  assert.ok(rowIndex >= 0, `${placement} in ${sheetName}`);
  sheet.getRange(4 + rowIndex, cols.indexOf('Exclude') + 1).setValue(true);
}

test('normal run: report rows per placement type, PMax and Display linked to the list', () => {
  const ads = account();
  const { errors } = run(ads);
  assert.deepEqual(errors, []);
  assert.deepEqual(out(ads, 'Output: Website').map(r => r[4]).sort(), ['news-site.com', 'spammy-site.xyz']);
  assert.deepEqual(out(ads, 'Output: YouTube').map(r => r[4]), ['K2zZwB3NY04']);
  assert.deepEqual(out(ads, 'Output: Mobile Application').map(r => r[4]), ['mobileapp::2-com.bad.game']);
  assert.equal(out(ads, 'Output: Google Products')[0][0], 'No results found');
  // PMax query only asks for impressions; the report survives that.
  assert.ok(ads.queries.some(q => q.includes('FROM performance_max_placement_view')));
  // Both the PMax and the Display campaign end up linked to the shared list.
  assert.deepEqual(ads.links.map(l => l.campaignId).sort(), ['500', '600']);
});

test('a second run does not link the list again', () => {
  const ads = account();
  run(ads);
  const first = ads.links.length;
  run(ads);
  assert.equal(ads.links.length, first);
});

test('ROAS is 0, not Infinity, when a placement has conversion value but no cost', () => {
  const ads = account();
  run(ads);
  const cols = header(ads, 'Output: Website');
  const row = out(ads, 'Output: Website').find(r => r[4] === 'news-site.com');
  assert.equal(row[cols.indexOf('ROAS')], 0);
  assert.ok(out(ads, 'Output: Website').flat().every(v => v !== Infinity));
});

test('a decimal minimum cost (2.01 * 1e6 = 2009999.9999999998) gives a whole-number cost_micros filter', () => {
  const ads = account({ minCost: 2.01 });
  const { errors } = run(ads);
  const displayQuery = ads.queries.find(q => q.includes("advertising_channel_type = 'DISPLAY'"));
  assert.match(displayQuery, /metrics\.cost_micros > 2010000 /);
  assert.ok(!errors.some(e => /Display placements/.test(e)), errors.join('\n'));
});

test('ticked placements are excluded: websites to the list, video on its campaign, apps by bulk upload, all logged', () => {
  const ads = account();
  run(ads);
  tick(ads, 'Output: Website', 'spammy-site.xyz');
  tick(ads, 'Output: YouTube', 'K2zZwB3NY04');
  tick(ads, 'Output: Mobile Application', 'mobileapp::2-com.bad.game');
  run(ads);
  assert.deepEqual(ads.exclusionLists[0].placements, ['spammy-site.xyz']);
  assert.deepEqual(ads.videoExclusions, [{ campaignId: '700', videoId: 'K2zZwB3NY04' }]);
  assert.equal(ads.bulkUploads.length, 1);
  assert.equal(ads.bulkUploads[0].mode, 'apply');
  const log = ads.spreadsheet.getSheetByName('Exclusion Log').rowsFrom(2);
  assert.equal(log.length, 3);
  assert.ok(log.every(r => r[7] === 'No'));
});

test('active placements only hides placements already in the exclusion list', () => {
  const ads = account({ activeOnly: true });
  ads.exclusionLists[0].placements.push('spammy-site.xyz');
  run(ads);
  assert.deepEqual(out(ads, 'Output: Website').map(r => r[4]), ['news-site.com']);
});

test('without active placements only, excluded placements stay and show as Excluded', () => {
  const ads = account({ activeOnly: false });
  ads.exclusionLists[0].placements.push('spammy-site.xyz');
  run(ads);
  const row = out(ads, 'Output: Website').find(r => r[4] === 'spammy-site.xyz');
  assert.equal(row[1], 'Excluded');
});

test('preview mode: bulk upload is previewed, nothing stored, log says Preview', () => {
  const ads = account({ preview: true });
  run(ads);
  tick(ads, 'Output: Website', 'spammy-site.xyz');
  tick(ads, 'Output: Mobile Application', 'mobileapp::2-com.bad.game');
  run(ads);
  assert.deepEqual(ads.exclusionLists[0].placements, []);
  assert.equal(ads.links.length, 0);
  assert.equal(ads.bulkUploads[0].mode, 'preview');
  assert.ok(ads.spreadsheet.getSheetByName('Exclusion Log').rowsFrom(2).every(r => r[7] === 'Yes'));
});

test('empty account: every output tab says No results found', () => {
  const ads = account({ empty: true });
  const { errors } = run(ads);
  assert.deepEqual(errors, []);
  for (const name of ['Output: Website', 'Output: YouTube', 'Output: Mobile Application', 'Output: Google Products']) {
    assert.equal(out(ads, name)[0][0], 'No results found');
  }
});

test('one failing report does not stop the others', () => {
  const ads = account();
  ads.failQuery('performance_max_placement_view', 'Internal error');
  const { errors } = run(ads);
  assert.ok(errors.some(e => /Performance Max placements: Internal error/.test(e)));
  assert.deepEqual(out(ads, 'Output: Website').map(r => r[4]), ['news-site.com']);
});

test('all reports failing stops the run with a clear error', () => {
  const ads = account();
  ads.failQuery('performance_max_placement_view', 'boom');
  ads.failQuery('detail_placement_view', 'boom');
  assert.throws(() => run(ads), /Failed to get placement data from all campaign types/);
});

test('missing exclusion list: report still written, ticked rows give a clear error', () => {
  const ads = account({ list: false });
  run(ads);
  assert.equal(out(ads, 'Output: Website').length, 2);
  tick(ads, 'Output: Website', 'spammy-site.xyz');
  assert.throws(() => run(ads), /Shared Placement Exclusion List named 'Bad placements' not found/);
});

test('bad sheet input: text numbers and a blank max results still run', () => {
  const ads = account();
  const settings = ads.spreadsheet.getSheetByName('Settings');
  settings.getRange(5, 2).setValue('thirty');
  settings.getRange(6, 2).setValue('100');
  settings.getRange(10, 2).setValue('');
  const { errors } = run(ads);
  assert.deepEqual(errors, []);
  assert.ok(ads.queries.some(q => /metrics\.impressions > 100/.test(q)));
  assert.deepEqual(out(ads, 'Output: Website').map(r => r[4]).sort(), ['news-site.com', 'spammy-site.xyz']);
});

test('placeholder spreadsheet URL gives a clear error', () => {
  const ads = account();
  const script = loadScript(SCRIPT, ads.globals(), ['main']);
  assert.throws(() => script.exports.main(), /Please update SPREADSHEET_URL/);
});
