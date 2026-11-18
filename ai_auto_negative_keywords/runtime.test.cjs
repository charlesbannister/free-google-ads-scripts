const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { join } = require('node:path');
const { test } = require('node:test');
const { runInNewContext } = require('node:vm');

const rulesUrl = 'https://autoneg.shabba.io/api/rules?sheetId=demo-sheet';
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

test('v15 advertises an update to existing v14 beta installs', () => {
  assert.equal(makeRuntime({}).script.SCRIPT_VERSION, 15);
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
