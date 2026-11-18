const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const { test } = require('node:test');
const { runInNewContext } = require('node:vm');

const rulesUrl = 'https://autoneg.shabba.io/api/rules?sheetId=demo-sheet&scriptVersion=16';
const examplesUrl = 'https://autoneg.shabba.io/api/example-terms?sheetId=demo-sheet&ruleId=101';
const script = readFileSync(join(__dirname, 'ai_auto_negative_keywords.js'), 'utf8')
  .replace('YOUR_SPREADSHEET_URL_HERE', 'https://docs.google.com/spreadsheets/d/demo-sheet/edit');

function makeRuntime(dependencies) {
  const logs = [];
  return {
    logs,
    script: runInNewContext(script + '\n({ runAccount, processRule_, _processReportRows_, SCRIPT_VERSION });', {
      console: {
        log: message => logs.push(message),
        warn: message => logs.push(message),
        error: message => logs.push(message),
      },
      ...dependencies,
    }),
  };
}

function response(status, body) {
  return {
    getResponseCode: () => status,
    getContentText: () => body,
  };
}

function account() {
  return { getTimeZone: () => 'UTC' };
}

function rule(overrides = {}) {
  return {
    id: 101,
    rule_number: 1,
    name: 'Search terms',
    level: 'account',
    enabled: true,
    generate_report: true,
    lookback_days: 30,
    entity_conditions: [],
    performance_conditions: [],
    text_conditions: null,
    ai_prompt: '',
    ...overrides,
  };
}

function rows(items) {
  let index = 0;
  return Object.freeze({
    hasNext: () => index < items.length,
    next: () => items[index++],
    totalNumEntities: () => items.length,
  });
}

function sampleSearchRow() {
  return {
    'search_term_view.search_term': 'live search term',
    'search_term_view.status': 'NONE',
    'campaign.name': 'Campaign',
    'campaign.id': '100',
    'ad_group.name': 'Ad group',
    'ad_group.id': '200',
    'metrics.impressions': '80',
    'metrics.clicks': '10',
    'metrics.cost_micros': '45000000',
    'metrics.conversions': '5',
    'metrics.conversions_value': '150',
  };
}

function reportSheet(tables) {
  const cells = {
    getValue: () => '',
    setValue: () => cells,
    setFontWeight: () => cells,
    setFontSize: () => cells,
    setNote: () => cells,
    setFontColor: () => cells,
    setFontLine: () => cells,
    setFormula: () => cells,
    setValues: data => tables.push(data),
    createFilter: () => {},
    getA1Notation: () => 'A5:N6',
  };
  return {
    getId: () => 'demo-sheet',
    getSheetByName: () => ({
      getName: () => '1',
      getSheetId: () => '123',
      getRange: () => cells,
      getLastRow: () => 4,
      getMaxRows: () => 100,
      getMaxColumns: () => 20,
      setFrozenRows: () => {},
      getFilter: () => null,
    }),
  };
}

function reportRuntime(items, fetch) {
  const tables = [];
  const { script: runtime } = makeRuntime({
    UrlFetchApp: { fetch },
    AdsApp: {
      currentAccount: account,
      getExecutionInfo: () => ({ isPreview: () => false }),
      report: () => ({ rows: () => rows(items) }),
    },
    Utilities: { formatDate: () => '2026-11-18' },
  });
  return { tables, runtime, spreadsheet: reportSheet(tables) };
}

test('v16 advertises an update to v15 installs', () => {
  assert.equal(makeRuntime({}).script.SCRIPT_VERSION, 16);
});

test('failed rule fetch exits before opening the spreadsheet', () => {
  const urls = [];
  let spreadsheetOpens = 0;
  const { script: runtime, logs } = makeRuntime({
    UrlFetchApp: { fetch(url) {
      urls.push(url);
      return response(404, '{"error":"Account not found"}');
    } },
    SpreadsheetApp: { openById() { spreadsheetOpens += 1; } },
  });

  assert.doesNotThrow(() => runtime.runAccount());
  assert.deepEqual(urls, [rulesUrl]);
  assert.equal(spreadsheetOpens, 0);
  assert.ok(logs.some(message => message.includes('HTTP Status: 404')));
});

test('failed live Ads report skips this rule before reading examples or writing', () => {
  let exampleRequests = 0;
  let sheetLookups = 0;
  let negativeOperations = 0;
  const { script: runtime } = makeRuntime({
    UrlFetchApp: { fetch() { exampleRequests += 1; return response(200, '{"exampleTerms":[{"term":"new keyword"}]}'); } },
    AdsApp: {
      currentAccount: account,
      report: () => { throw Error('Google Ads report unavailable'); },
      getExecutionInfo: () => ({ isPreview: () => false }),
      negativeKeywordLists: () => { negativeOperations += 1; },
    },
    Utilities: { formatDate: () => '2026-11-18' },
  });

  runtime.processRule_(rule({
    auto_apply: true,
    include_examples: true,
    negative_list_name: 'List',
  }), {
    getId: () => 'demo-sheet',
    getSheetByName() { sheetLookups += 1; },
  });

  assert.equal(exampleRequests, 0);
  assert.equal(sheetLookups, 0);
  assert.equal(negativeOperations, 0);
});

test('malformed examples are skipped while live search terms still reach the report', () => {
  const { runtime, tables, spreadsheet } = reportRuntime([sampleSearchRow()], url => {
    assert.equal(url, examplesUrl);
    return response(200, 'upstream returned a non-JSON response');
  });

  runtime.processRule_(rule({ include_examples: true }), spreadsheet);
  assert.equal(tables.length, 1);
  assert.deepEqual(Array.from(tables[0].slice(1), row => row[0]), ['live search term']);
});

test('example terms expand to ad groups without overriding or losing live rows', () => {
  const { runtime, tables, spreadsheet } = reportRuntime([sampleSearchRow()], url => {
    assert.equal(url, examplesUrl);
    return response(200, '{"exampleTerms":[{"term":"example search term"}]}');
  });

  runtime.processRule_(rule({ include_examples: true, level: 'adgroup' }), spreadsheet);
  assert.equal(tables.length, 1);
  assert.deepEqual(Array.from(tables[0].slice(1), row => Array.from(row.slice(0, 3))), [
    ['live search term', 'Campaign', 'Ad group'],
    ['example search term', 'Campaign', 'Ad group'],
  ]);
});

test('a legacy null AI prompt is treated as an empty prompt', () => {
  const { runtime, tables, spreadsheet } = reportRuntime([sampleSearchRow()]);
  runtime.processRule_(rule({ ai_prompt: null }), spreadsheet);
  assert.equal(tables.length, 1);
  assert.equal(tables[0][1][0], 'live search term');
});

test('Performance Max terms still appear in reports with auto-apply enabled', () => {
  const { script: runtime } = makeRuntime({
    AdsApp: { getExecutionInfo: () => ({ isPreview: () => false }) },
  });

  const result = runtime._processReportRows_({
    regularRows: rows([]),
    ruleLevel: 'campaign',
    autoApply: true,
    generateReport: true,
    performanceConditions: [],
    textConditions: null,
    pmaxRows: [{ searchTerm: 'Performance Max term', campaignId: '500', campaignName: 'PMax' }],
    rule: rule({ level: 'campaign' }),
    exampleTerms: [],
  });

  assert.equal(result.outputDataRows.length, 1);
  assert.equal(result.outputDataRows[0][0], 'Performance Max term');
  assert.ok(result.outputDataRows[0].includes('PMAX not supported'));
  assert.equal(Object.keys(result.negativesByCampaign).length, 0);
});

test('explicitly included example terms remain eligible for auto-apply', () => {
  const { script: runtime } = makeRuntime({
    AdsApp: { getExecutionInfo: () => ({ isPreview: () => false }) },
  });

  const result = runtime._processReportRows_({
    regularRows: rows([]),
    ruleLevel: 'account',
    autoApply: true,
    generateReport: true,
    performanceConditions: [],
    textConditions: null,
    pmaxRows: [],
    rule: rule({ include_examples: true }),
    exampleTerms: [{ term: 'Example term' }],
  });

  assert.deepEqual(Array.from(result.negativesByAccount.terms), ['Example term']);
});

function keywordRule(scope, overrides = {}) {
  return rule({
    level: scope,
    automated_targeting_keywords: { enabled: true, scope, threshold: 70 },
    ...overrides,
  });
}

function keywordRuntime({ rules, searchRows, keywordRows, failKeywordQuery = false, examples }) {
  const tables = [];
  const queries = [];
  const { script: runtime, logs } = makeRuntime({
    UrlFetchApp: { fetch(url) {
      if (url === rulesUrl) return response(200, JSON.stringify(rules));
      if (url === examplesUrl && examples) return response(200, JSON.stringify({ exampleTerms: examples }));
      throw Error(`Unexpected request ${url}`);
    } },
    SpreadsheetApp: { openById: () => reportSheet(tables) },
    AdsApp: {
      currentAccount: account,
      getExecutionInfo: () => ({ isPreview: () => false }),
      report(query) {
        queries.push(query);
        if (query.includes('FROM keyword_view')) {
          if (failKeywordQuery) throw Error('Keyword query unavailable');
          return { rows: () => rows(keywordRows) };
        }
        if (query.includes('FROM search_term_view')) return { rows: () => rows(searchRows) };
        throw Error(`Unexpected query ${query}`);
      },
    },
    Utilities: { formatDate: () => '2026-11-18' },
  });
  return { runtime, logs, tables, queries, spreadsheet: reportSheet(tables) };
}

function searchTerm(term, adGroupId = '200', campaignId = '100') {
  return { ...sampleSearchRow(), 'search_term_view.search_term': term, 'campaign.id': campaignId, 'ad_group.id': adGroupId };
}

const acmeKeyword = {
  'campaign.id': '100', 'ad_group.id': '201', 'ad_group_criterion.keyword.text': 'Acme Plumbing',
};

test('campaign keyword protection runs once per account and skips Performance Max', () => {
  const { runtime, tables, queries } = keywordRuntime({
    rules: [keywordRule('campaign', { text_conditions: { sections: [{ conditions: [] }] } }), keywordRule('adgroup', { id: 102 })],
    searchRows: [searchTerm('ACME plumbing services'), searchTerm('plumbing acme'), searchTerm('acmee plumbing'), searchTerm('unrelated job')],
    keywordRows: [acmeKeyword],
  });

  runtime.runAccount();
  assert.equal(queries.filter(query => query.includes('FROM keyword_view')).length, 1);
  assert.equal(queries.filter(query => query.includes('FROM search_term_view')).length, 2);
  assert.ok(queries.every(query => !query.includes('campaign_search_term_insight')));
  assert.deepEqual(Array.from(tables[0].slice(1), row => row[0]), ['unrelated job']);
});

test('ad-group protection scopes keywords and keeps custom words alongside them', () => {
  const { runtime, tables, spreadsheet } = keywordRuntime({
    searchRows: [searchTerm('ACME plumbing', '200'), searchTerm('ACME plumbing', '201'), searchTerm('trusted brand', '200'), searchTerm('jobs', '999')],
    keywordRows: [acmeKeyword, { ...acmeKeyword, 'ad_group.id': '200', 'ad_group_criterion.keyword.text': 'plumbing repairs' }],
  });
  runtime.processRule_(keywordRule('adgroup', {
    text_conditions: { sections: [{ conditions: [{ text: 'trusted', matchType: 'contains' }] }] },
  }), spreadsheet);
  assert.deepEqual(Array.from(tables[0].slice(1), row => row[0]), ['ACME plumbing']);
});

test('a failed or empty keyword query leaves the previous report and skips mutations', () => {
  for (const failKeywordQuery of [true, false]) {
    const { runtime, queries, tables, logs, spreadsheet } = keywordRuntime({
      failKeywordQuery, searchRows: [searchTerm('jobs')], keywordRows: [],
    });
    runtime.processRule_(keywordRule('campaign', { auto_apply: true }), spreadsheet);
    assert.equal(queries.length, 1);
    assert.equal(tables.length, 0);
    assert.ok(logs.some(message => message.includes('keyword protection')));
  }
});

test('legacy automated settings without scope preserve the existing custom-only rule', () => {
  const { runtime, queries, tables, spreadsheet } = keywordRuntime({
    searchRows: [searchTerm('unrelated job')], keywordRows: [acmeKeyword],
  });
  runtime.processRule_(rule({ automated_targeting_keywords: { enabled: true, threshold: 70 } }), spreadsheet);
  assert.ok(queries.every(query => !query.includes('FROM keyword_view')));
  assert.deepEqual(Array.from(tables[0].slice(1), row => row[0]), ['unrelated job']);
});

test('keyword protection reports without writing to a shared negative list', () => {
  const { runtime, tables, logs, spreadsheet } = keywordRuntime({
    searchRows: [searchTerm('unrelated job')], keywordRows: [acmeKeyword],
  });
  runtime.processRule_(keywordRule('campaign', { auto_apply: true, negative_list_name: 'Shared list' }), spreadsheet);
  assert.ok(logs.some(message => message.includes('shared negative list')));
  assert.ok(tables[0][1].includes('No (Auto Apply Disabled)'));
});

test('explicit but invalid keyword scopes skip the rule', () => {
  for (const scope of ['account', null, 'adgroup']) {
    const { runtime, queries, tables, spreadsheet } = keywordRuntime({
      searchRows: [searchTerm('unrelated job')], keywordRows: [acmeKeyword],
    });
    runtime.processRule_(keywordRule(scope, { level: 'campaign' }), spreadsheet);
    assert.equal(queries.length, 0);
    assert.equal(tables.length, 0);
  }
});

test('keyword protection applies to example terms too', () => {
  const { runtime, tables, spreadsheet } = keywordRuntime({
    searchRows: [searchTerm('unrelated job', '201')], keywordRows: [acmeKeyword],
    examples: [{ term: 'acme plumbing' }, { term: 'another example' }],
  });
  runtime.processRule_(keywordRule('adgroup', { include_examples: true }), spreadsheet);
  assert.deepEqual(Array.from(tables[0].slice(1), row => row[0]), ['unrelated job', 'another example']);
});
