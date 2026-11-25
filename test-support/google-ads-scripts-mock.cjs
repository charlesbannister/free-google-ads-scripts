'use strict';
/**
 * Reusable mocks of the Google Ads Scripts runtime, for Node's built-in test runner.
 *
 *   const { loadScript, MockGoogleAds } = require('../test-support/google-ads-scripts-mock.cjs');
 *   const ads = new MockGoogleAds({ customerId: '123-456-7890' });
 *   ads.addCampaign({ id: '500', name: 'PMax', channel: 'PERFORMANCE_MAX' });
 *   ads.addRows('campaign_search_term_view', [{ 'campaign.id': '500', 'campaign_search_term_view.search_term': 'x', ... }]);
 *   const { exports, logs } = loadScript('path/to/script.js', ads.globals(), ['main']);
 *
 * What it imitates (and what it does not):
 * - AdsApp.report(GAQL): a small GAQL engine over rows you add per resource. It supports
 *   SELECT / FROM / WHERE (=, !=, <, >, <=, >=, IN, NOT IN, LIKE, NOT LIKE, REGEXP_MATCH,
 *   NOT ... REGEXP_MATCH, BETWEEN), ORDER BY and LIMIT. LIKE ignores case and treats [x] as a
 *   literal x, as Google's does. Fields from another resource than the ones a view allows throw,
 *   like Google's "Unrecognized field". When no segments.* field is selected, rows are added up
 *   per selected attribute (Google aggregates over the date range). REGEXP_MATCH is a partial
 *   JS match here: Google uses RE2, so keep test patterns simple.
 * - AdsApp.mutate / mutateAll: campaignCriterionOperation (negative keywords) and
 *   sharedCriterionOperation creates, with Google's keyword limits (80 characters, 10 words),
 *   duplicate refusal and partialFailure. Successful creates become campaign_criterion /
 *   shared_criterion rows, so a second run sees them. In preview mode nothing is stored.
 * - AdsApp.campaigns / shoppingCampaigns / performanceMaxCampaigns / adGroups /
 *   shoppingAdGroups / negativeKeywordLists selectors. AdsApp.campaigns() does not return
 *   Shopping or Performance Max campaigns, as in Google Ads.
 * - SpreadsheetApp, UrlFetchApp, MailApp, Utilities: enough for report-writing scripts.
 */
const { readFileSync } = require('node:fs');
const { runInNewContext } = require('node:vm');

// ---------------------------------------------------------------- iterators

function iterator(items) {
  let index = 0;
  return {
    hasNext: () => index < items.length,
    next: () => {
      if (index >= items.length) throw new Error('No more entities');
      return items[index++];
    },
    totalNumEntities: () => items.length,
  };
}

// ---------------------------------------------------------------- GAQL engine

// Which field prefixes each resource accepts besides its own (a rough copy of Google's rules).
const RESOURCE_FIELDS = {
  search_term_view: ['campaign', 'ad_group', 'customer', 'metrics', 'segments'],
  campaign_search_term_view: ['campaign', 'customer', 'metrics', 'segments'],
  campaign_search_term_insight: ['customer', 'metrics', 'segments'],
  campaign: ['customer', 'metrics', 'segments', 'campaign_budget', 'bidding_strategy'],
  ad_group: ['campaign', 'customer', 'metrics', 'segments'],
  keyword_view: ['campaign', 'ad_group', 'ad_group_criterion', 'customer', 'metrics', 'segments'],
  campaign_criterion: ['campaign', 'customer'],
  shared_set: ['customer'],
  shared_criterion: ['shared_set', 'customer'],
  campaign_shared_set: ['campaign', 'shared_set', 'customer'],
  customer: ['metrics', 'segments'],
};

function normaliseQuery(query) {
  return String(query).replace(/\s+/g, ' ').trim();
}

function parseGaql(query) {
  const q = normaliseQuery(query);
  const m = q.match(/^SELECT (.+?) FROM (\w+)(?: WHERE (.+?))?(?: ORDER BY (.+?))?(?: LIMIT (\d+))?$/i);
  if (!m) throw new Error(`Error parsing GAQL query: ${q}`);
  return {
    fields: m[1].split(',').map(f => f.trim()).filter(Boolean),
    resource: m[2],
    where: m[3] ? parseWhere(m[3]) : [],
    orderBy: m[4] ? m[4].split(',').map(part => {
      const [field, dir] = part.trim().split(/\s+/);
      return { field, desc: /^desc$/i.test(dir || '') };
    }) : [],
    limit: m[5] ? parseInt(m[5], 10) : null,
  };
}

const STRING = `'(?:[^'\\\\]|\\\\.)*'|"(?:[^"\\\\]|\\\\.)*"`;
const VALUE = `${STRING}|[\\w.\\-]+`;
const FIELD = `[a-z_]+(?:\\.[a-z_]+)+|[a-z_]+`;
const CONDITION_PATTERNS = [
  { kind: 'not_regex', re: new RegExp(`^NOT (${FIELD}) REGEXP_MATCH (${STRING})`, 'i') },
  { kind: 'between', re: new RegExp(`^(${FIELD}) BETWEEN (${VALUE}) AND (${VALUE})`, 'i') },
  { kind: 'in', re: new RegExp(`^(${FIELD}) (NOT IN|IN) \\(((?:\\s*(?:${VALUE})\\s*,?)*)\\)`, 'i') },
  { kind: 'like', re: new RegExp(`^(${FIELD}) (NOT LIKE|LIKE) (${STRING})`, 'i') },
  { kind: 'regex', re: new RegExp(`^(${FIELD}) REGEXP_MATCH (${STRING})`, 'i') },
  { kind: 'compare', re: new RegExp(`^(${FIELD}) (!=|>=|<=|=|>|<) (${VALUE})`, 'i') },
];

function unquote(value) {
  const v = String(value).trim();
  if ((v.startsWith("'") && v.endsWith("'")) || (v.startsWith('"') && v.endsWith('"'))) {
    return v.slice(1, -1).replace(/\\(.)/g, '$1');
  }
  return v;
}

function splitList(text) {
  const out = [];
  const re = new RegExp(VALUE, 'g');
  let m;
  while ((m = re.exec(text))) out.push(unquote(m[0]));
  return out;
}

function parseWhere(text) {
  const conditions = [];
  let rest = text.trim();
  while (rest.length) {
    let matched = null;
    for (const pattern of CONDITION_PATTERNS) {
      const m = rest.match(pattern.re);
      if (m) { matched = { pattern, m }; break; }
    }
    if (!matched) throw new Error(`Error parsing GAQL WHERE near: ${rest}`);
    const { pattern, m } = matched;
    const c = { kind: pattern.kind, field: m[1] };
    if (pattern.kind === 'between') { c.from = unquote(m[2]); c.to = unquote(m[3]); }
    if (pattern.kind === 'in') { c.not = /NOT/i.test(m[2]); c.values = splitList(m[3]); }
    if (pattern.kind === 'like') { c.not = /NOT/i.test(m[2]); c.pattern = unquote(m[3]); }
    if (pattern.kind === 'regex' || pattern.kind === 'not_regex') c.pattern = unquote(m[2]);
    if (pattern.kind === 'compare') { c.op = m[2]; c.value = unquote(m[3]); }
    conditions.push(c);
    rest = rest.slice(m[0].length).trim();
    if (!rest.length) break;
    const and = rest.match(/^AND\s+/i);
    if (!and) throw new Error(`Error parsing GAQL WHERE: expected AND near: ${rest}`);
    rest = rest.slice(and[0].length);
  }
  return conditions;
}

function likeToRegExp(pattern) {
  let out = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '[' && pattern[i + 2] === ']') { out += escapeRe(pattern[i + 1]); i += 2; continue; }
    if (ch === '%') out += '.*';
    else if (ch === '_') out += '.';
    else out += escapeRe(ch);
  }
  return new RegExp(`^${out}$`, 'is');
}

function escapeRe(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function same(a, b) {
  if (a === undefined || a === null) return false;
  const na = Number(a); const nb = Number(b);
  if (String(a).trim() !== '' && String(b).trim() !== '' && !isNaN(na) && !isNaN(nb)) return na === nb;
  return String(a).toUpperCase() === String(b).toUpperCase();
}

function compare(a, op, b) {
  if (op === '=') return same(a, b);
  if (op === '!=') return !same(a, b);
  const na = Number(a); const nb = Number(b);
  if (isNaN(na) || isNaN(nb)) return false;
  return { '>': na > nb, '<': na < nb, '>=': na >= nb, '<=': na <= nb }[op];
}

function matches(row, c) {
  const value = row[c.field];
  // Rows without a date are treated as inside every date range.
  if (c.field === 'segments.date' && value === undefined) return true;
  switch (c.kind) {
    case 'compare': return compare(value, c.op, c.value);
    case 'between': return value >= c.from && value <= c.to;
    case 'in': { const hit = c.values.some(v => same(value, v)); return c.not ? !hit : hit; }
    case 'like': { const hit = likeToRegExp(c.pattern).test(String(value ?? '')); return c.not ? !hit : hit; }
    case 'regex': return new RegExp(c.pattern).test(String(value ?? ''));
    case 'not_regex': return !new RegExp(c.pattern).test(String(value ?? ''));
    default: return false;
  }
}

function checkField(resource, field) {
  const own = field.split('.')[0];
  const allowed = RESOURCE_FIELDS[resource];
  if (!allowed) return;
  if (own === resource || allowed.includes(own)) return;
  throw new Error(`Unrecognized field in the query: '${field}' is not compatible with '${resource}'.`);
}

// ---------------------------------------------------------------- mutate results

class MockMutateResult {
  constructor(ok, resourceName, errors = []) {
    this._ok = ok; this._name = resourceName; this._errors = errors;
  }
  isSuccessful() { return this._ok; }
  getResourceName() { return this._name; }
  getErrorMessages() { return this._errors.slice(); }
  getEntityType() { return 'CampaignCriterion'; }
}

// ---------------------------------------------------------------- spreadsheet

function a1ToRowCol(a1) {
  const m = String(a1).match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
  if (!m) throw new Error(`Bad A1 notation: ${a1}`);
  const col = s => s.toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const r1 = parseInt(m[2], 10); const c1 = col(m[1]);
  const r2 = m[4] ? parseInt(m[4], 10) : r1; const c2 = m[3] ? col(m[3]) : c1;
  return { row: r1, col: c1, rows: r2 - r1 + 1, cols: c2 - c1 + 1 };
}

function colName(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

class MockRange {
  constructor(sheet, row, col, rows = 1, cols = 1) {
    Object.assign(this, { sheet, row, col, rows, cols });
  }
  getValue() { return this.sheet._get(this.row, this.col); }
  getValues() {
    const out = [];
    for (let r = 0; r < this.rows; r++) {
      const line = [];
      for (let c = 0; c < this.cols; c++) line.push(this.sheet._get(this.row + r, this.col + c));
      out.push(line);
    }
    return out;
  }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
  setValues(values) {
    if (values.length !== this.rows || values.some(line => line.length !== this.cols)) {
      throw new Error(`The number of rows or columns in the data does not match the range (${values.length}x${values[0] && values[0].length} vs ${this.rows}x${this.cols}).`);
    }
    values.forEach((line, r) => line.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v)));
    this.sheet.writes.push(values.map(line => line.slice()));
    return this;
  }
  setFormula(f) { return this.setValue(f); }
  clearContent() {
    for (let r = 0; r < this.rows; r++) for (let c = 0; c < this.cols; c++) this.sheet._set(this.row + r, this.col + c, '');
    return this;
  }
  createFilter() { this.sheet._filter = { remove() {} }; return this.sheet._filter; }
  getA1Notation() { return `${colName(this.col)}${this.row}:${colName(this.col + this.cols - 1)}${this.row + this.rows - 1}`; }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getNumRows() { return this.rows; }
  getNumColumns() { return this.cols; }
}
// Formatting calls are accepted and ignored.
for (const name of ['setFontWeight', 'setFontSize', 'setNote', 'setFontColor', 'setFontLine', 'setBackground',
  'setNumberFormat', 'setHorizontalAlignment', 'setWrap', 'setBorder']) {
  MockRange.prototype[name] = function () { return this; };
}

class MockSheet {
  constructor(name, id, { maxRows = 1000, maxColumns = 26 } = {}) {
    this.name = name; this.id = id; this.cells = new Map(); this.writes = [];
    this.maxRows = maxRows; this.maxColumns = maxColumns; this._filter = null;
  }
  _key(r, c) { return `${r}:${c}`; }
  _get(r, c) { return this.cells.has(this._key(r, c)) ? this.cells.get(this._key(r, c)) : ''; }
  _set(r, c, v) {
    if (r > this.maxRows || c > this.maxColumns) throw new Error(`Range ${colName(c)}${r} is outside the sheet (${this.maxRows} rows, ${this.maxColumns} columns).`);
    if (v === '' || v === null || v === undefined) this.cells.delete(this._key(r, c)); else this.cells.set(this._key(r, c), v);
  }
  getName() { return this.name; }
  getSheetId() { return this.id; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') { const p = a1ToRowCol(a); return new MockRange(this, p.row, p.col, p.rows, p.cols); }
    return new MockRange(this, a, b, c || 1, d || 1);
  }
  getDataRange() { return new MockRange(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  getLastRow() { let n = 0; for (const k of this.cells.keys()) n = Math.max(n, +k.split(':')[0]); return n; }
  getLastColumn() { let n = 0; for (const k of this.cells.keys()) n = Math.max(n, +k.split(':')[1]); return n; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxColumns; }
  insertRowsAfter(_after, n) { this.maxRows += n; return this; }
  insertColumnsAfter(_after, n) { this.maxColumns += n; return this; }
  deleteRows(start, n) {
    const moved = new Map();
    for (const [k, v] of this.cells) {
      const [r, c] = k.split(':').map(Number);
      if (r < start) moved.set(k, v); else if (r >= start + n) moved.set(this._key(r - n, c), v);
    }
    this.cells = moved; this.maxRows -= n; return this;
  }
  deleteColumns(_start, n) { this.maxColumns -= n; return this; }
  setFrozenRows() { return this; }
  getFilter() { return this._filter; }
  clear() { this.cells.clear(); return this; }
  /** Rows of the sheet as arrays, from row `from` (1-based). */
  rowsFrom(from = 1) {
    const last = this.getLastRow(); const cols = this.getLastColumn(); const out = [];
    for (let r = from; r <= last; r++) { const line = []; for (let c = 1; c <= cols; c++) line.push(this._get(r, c)); out.push(line); }
    return out;
  }
}

class MockSpreadsheet {
  constructor(id = 'demo-sheet') { this.id = id; this.sheets = []; this.nextId = 100; this.active = null; }
  getId() { return this.id; }
  getUrl() { return `https://docs.google.com/spreadsheets/d/${this.id}/edit`; }
  getName() { return 'Mock spreadsheet'; }
  getSheetByName(name) { return this.sheets.find(s => s.name === String(name)) || null; }
  getSheets() { return this.sheets.slice(); }
  getNumSheets() { return this.sheets.length; }
  insertSheet(name) {
    if (this.getSheetByName(name)) throw new Error(`A sheet with the name "${name}" already exists.`);
    const sheet = new MockSheet(String(name), this.nextId++); this.sheets.push(sheet); return sheet;
  }
  setActiveSheet(sheet) { this.active = sheet; return sheet; }
  getActiveSheet() { return this.active; }
  moveActiveSheet() {}
}

// ---------------------------------------------------------------- the account

const CRITERION_MATCH_TYPES = ['EXACT', 'PHRASE', 'BROAD'];

class MockGoogleAds {
  /**
   * @param {Object} [options]
   * @param {string} [options.customerId] e.g. '123-456-7890'
   * @param {string} [options.accountName]
   * @param {string} [options.timeZone]
   * @param {boolean} [options.preview] AdsApp.getExecutionInfo().isPreview()
   * @param {Object} [options.fetch] { urlOrPrefix: body | (url, params) => {status, body} | Error }
   * @param {boolean} [options.mutateAll] set false to remove AdsApp.mutateAll (older runtimes)
   */
  constructor(options = {}) {
    this.customerId = options.customerId || '123-456-7890';
    this.accountName = options.accountName || 'Mock account';
    this.timeZone = options.timeZone || 'Europe/London';
    this.preview = options.preview === true;
    this.withMutateAll = options.mutateAll !== false;
    this.rows = {};                  // resource -> rows
    this.campaigns = [];             // { id, name, channel, status }
    this.adGroups = [];              // { id, name, campaignId, status }
    this.lists = [];                 // { id, name, type, keywords: [] }
    this.queries = [];               // every GAQL query sent
    this.mutations = [];             // every mutate operation sent, with its result
    this.entityNegatives = [];       // createNegativeKeyword calls: { level, id, text }
    this.listAdds = [];              // addNegativeKeywords calls: { list, keywords }
    this.emails = [];
    this.fetches = [];
    this.failures = [];              // { resource|match, error }
    this.fetchRoutes = options.fetch || {};
    this.spreadsheet = new MockSpreadsheet(options.spreadsheetId || 'demo-sheet');
    this.nextCriterionId = 9000;
  }

  // ------------------------------------------------ setup helpers

  addCampaign({ id, name, channel = 'SEARCH', status = 'ENABLED' }) {
    const campaign = { id: String(id), name, channel, status };
    this.campaigns.push(campaign);
    this.addRows('campaign', [{ 'campaign.id': campaign.id, 'campaign.name': name, 'campaign.status': status, 'campaign.advertising_channel_type': channel }]);
    return campaign;
  }

  addAdGroup({ id, name, campaignId, status = 'ENABLED' }) {
    const adGroup = { id: String(id), name, campaignId: String(campaignId), status };
    this.adGroups.push(adGroup);
    return adGroup;
  }

  addNegativeList({ id, name, type = 'NEGATIVE_KEYWORDS' }) {
    const list = { id: String(id), name, type, keywords: [] };
    this.lists.push(list);
    return list;
  }

  /** Adds rows to a resource. Campaign fields are filled in from addCampaign when missing. */
  addRows(resource, rows) {
    const list = this.rows[resource] || (this.rows[resource] = []);
    for (const row of rows) {
      const full = { ...row };
      const campaign = this.campaigns.find(c => c.id === String(full['campaign.id']));
      if (campaign && resource !== 'campaign') {
        full['campaign.name'] ??= campaign.name;
        full['campaign.status'] ??= campaign.status;
        full['campaign.advertising_channel_type'] ??= campaign.channel;
      }
      list.push(full);
    }
    return this;
  }

  /** Search term rows for a Search/Shopping campaign (search_term_view). */
  addSearchTerm({ term, campaignId, adGroupId, status = 'NONE', impressions = 0, clicks = 0, cost = 0, conversions = 0, value = 0, date }) {
    const adGroup = this.adGroups.find(g => g.id === String(adGroupId)) || {};
    return this.addRows('search_term_view', [{
      'search_term_view.search_term': term, 'search_term_view.status': status,
      'campaign.id': String(campaignId), 'ad_group.id': String(adGroupId), 'ad_group.name': adGroup.name,
      'ad_group.status': adGroup.status || 'ENABLED',
      ...metrics({ impressions, clicks, cost, conversions, value }), ...(date ? { 'segments.date': date } : {}),
    }]);
  }

  /** Performance Max search term rows (campaign_search_term_view). Several rows per term
   * (for example one per date) are added up when the query selects no segment. */
  addPmaxSearchTerm({ term, campaignId, impressions = 0, clicks = 0, cost = 0, conversions = 0, value = 0, date }) {
    return this.addRows('campaign_search_term_view', [{
      'campaign_search_term_view.search_term': term, 'campaign.id': String(campaignId),
      ...metrics({ impressions, clicks, cost, conversions, value }), ...(date ? { 'segments.date': date } : {}),
    }]);
  }

  /** Performance Max search term categories (campaign_search_term_insight). */
  addPmaxCategory({ label, campaignId, impressions = 0, clicks = 0, conversions = 0, value = 0 }) {
    return this.addRows('campaign_search_term_insight', [{
      'campaign_search_term_insight.category_label': label,
      'campaign_search_term_insight.campaign_id': String(campaignId),
      'metrics.impressions': String(impressions), 'metrics.clicks': String(clicks),
      'metrics.conversions': String(conversions), 'metrics.conversions_value': String(value),
    }]);
  }

  /** An existing campaign negative keyword (campaign_criterion). */
  addCampaignNegative({ campaignId, text, matchType = 'EXACT' }) {
    return this.addRows('campaign_criterion', [{
      'campaign.id': String(campaignId), 'campaign_criterion.negative': 'TRUE', 'campaign_criterion.type': 'KEYWORD',
      'campaign_criterion.keyword.text': text, 'campaign_criterion.keyword.match_type': matchType,
      'campaign_criterion.criterion_id': String(this.nextCriterionId++),
    }]);
  }

  /** Makes queries fail. `match` is a resource name, a substring of the query, or a RegExp. */
  failQuery(match, message = 'Mock Google Ads error') {
    this.failures.push({ match, message });
    return this;
  }

  // ------------------------------------------------ AdsApp.report

  report(query) {
    this.queries.push(normaliseQuery(query));
    const parsed = parseGaql(query);
    for (const f of this.failures) {
      const hit = f.match instanceof RegExp ? f.match.test(query) : (parsed.resource === f.match || normaliseQuery(query).includes(f.match));
      if (hit) throw new Error(f.message);
    }
    if (!RESOURCE_FIELDS[parsed.resource]) throw new Error(`Invalid resource name: ${parsed.resource}`);
    for (const field of parsed.fields) checkField(parsed.resource, field);
    for (const c of parsed.where) checkField(parsed.resource, c.field);
    let rows = (this.rows[parsed.resource] || []).filter(row => parsed.where.every(c => matches(row, c)));
    rows = aggregate(rows, parsed.fields);
    for (const { field, desc } of parsed.orderBy.slice().reverse()) {
      rows.sort((a, b) => {
        const x = Number(a[field]); const y = Number(b[field]);
        const d = !isNaN(x) && !isNaN(y) ? x - y : String(a[field]).localeCompare(String(b[field]));
        return desc ? -d : d;
      });
    }
    if (parsed.limit !== null) rows = rows.slice(0, parsed.limit);
    const projected = rows.map(row => {
      const out = {};
      for (const field of parsed.fields) out[field] = row[field] === undefined ? (field.startsWith('metrics.') ? '0' : '') : row[field];
      return out;
    });
    return { rows: () => iterator(projected) };
  }

  // ------------------------------------------------ AdsApp.mutate / mutateAll

  _applyOperation(operation) {
    const [key] = Object.keys(operation);
    const body = operation[key] || {};
    if (key === 'campaignCriterionOperation' && body.create) return this._createCampaignCriterion(body.create);
    if (key === 'sharedCriterionOperation' && body.create) return this._createSharedCriterion(body.create);
    return new MockMutateResult(false, null, [`Mock does not support ${key}`]);
  }

  _keywordErrors(keyword) {
    if (!keyword || typeof keyword.text !== 'string' || !keyword.text.trim()) return ['Keyword text is required.'];
    if (!CRITERION_MATCH_TYPES.includes(keyword.matchType)) return [`Invalid match type: ${keyword.matchType}`];
    if (keyword.text.length > 80) return ['KEYWORD_TEXT_TOO_LONG: The keyword text is too long.'];
    if (keyword.text.trim().split(/\s+/).length > 10) return ['TOO_MANY_WORDS: The keyword has too many words.'];
    if (/[!@%^*()=\\{};~`<>?|,]/.test(keyword.text)) return ['INVALID_CHARACTERS: The keyword has invalid characters.'];
    return [];
  }

  _createCampaignCriterion(create) {
    const m = String(create.campaign || '').match(/^customers\/(\d+)\/campaigns\/(\d+)$/);
    if (!m) return new MockMutateResult(false, null, [`Invalid resource name: ${create.campaign}`]);
    if (m[1] !== this.customerId.replace(/-/g, '')) return new MockMutateResult(false, null, ['USER_PERMISSION_DENIED: the customer does not match.']);
    const campaign = this.campaigns.find(c => c.id === m[2]);
    if (!campaign) return new MockMutateResult(false, null, ['NOT_FOUND: campaign not found.']);
    if (create.negative !== true) return new MockMutateResult(false, null, ['Only negative keyword criteria are mocked.']);
    const errors = this._keywordErrors(create.keyword);
    if (errors.length) return new MockMutateResult(false, null, errors);
    const existing = (this.rows.campaign_criterion || []).find(r => r['campaign.id'] === campaign.id &&
      String(r['campaign_criterion.keyword.text']).toLowerCase() === create.keyword.text.toLowerCase() &&
      r['campaign_criterion.keyword.match_type'] === create.keyword.matchType);
    if (existing) return new MockMutateResult(false, null, ['DUPLICATE_CRITERION: The criterion already exists.']);
    const id = String(this.nextCriterionId++);
    const name = `customers/${m[1]}/campaignCriteria/${campaign.id}~${id}`;
    if (!this.preview) {
      this.addRows('campaign_criterion', [{
        'campaign.id': campaign.id, 'campaign_criterion.negative': 'TRUE', 'campaign_criterion.type': 'KEYWORD',
        'campaign_criterion.keyword.text': create.keyword.text, 'campaign_criterion.keyword.match_type': create.keyword.matchType,
        'campaign_criterion.criterion_id': id,
      }]);
    }
    return new MockMutateResult(true, name);
  }

  _createSharedCriterion(create) {
    const m = String(create.sharedSet || '').match(/^customers\/(\d+)\/sharedSets\/(\d+)$/);
    const list = m && this.lists.find(l => l.id === m[2]);
    if (!list) return new MockMutateResult(false, null, ['NOT_FOUND: shared set not found.']);
    const errors = this._keywordErrors(create.keyword);
    if (errors.length) return new MockMutateResult(false, null, errors);
    if (!this.preview) list.keywords.push(`${create.keyword.matchType}:${create.keyword.text}`);
    return new MockMutateResult(true, `customers/${m[1]}/sharedCriteria/${list.id}~${this.nextCriterionId++}`);
  }

  mutate(operation) {
    const result = this._applyOperation(operation);
    this.mutations.push({ operation, result });
    return result;
  }

  mutateAll(operations, options = {}) {
    if (!Array.isArray(operations)) throw new Error('mutateAll takes an array of operations.');
    if (operations.length > 10000) throw new Error('Too many operations in one request (10,000 at most).');
    if (options.partialFailure === true) return operations.map(op => this.mutate(op));
    // Without partial failure Google applies all or nothing.
    const saved = JSON.stringify({ rows: this.rows, lists: this.lists, next: this.nextCriterionId });
    const results = operations.map(op => this.mutate(op));
    if (results.some(r => !r.isSuccessful())) {
      const restored = JSON.parse(saved);
      this.rows = restored.rows; this.lists.forEach((l, i) => { l.keywords = restored.lists[i].keywords; }); this.nextCriterionId = restored.next;
      return results.map(r => r.isSuccessful() ? new MockMutateResult(false, null, ['Not applied: another operation in the request failed.']) : r);
    }
    return results;
  }

  // ------------------------------------------------ entity selectors

  _selector(source, makeEntity) {
    const state = { ids: null, conditions: [], limit: null };
    const selector = {
      withIds(ids) { state.ids = ids.map(String); return selector; },
      withCondition(condition) { state.conditions.push(...parseWhere(condition)); return selector; },
      withLimit(n) { state.limit = n; return selector; },
      orderBy() { return selector; },
      forDateRange() { return selector; },
      get: () => {
        let items = source().filter(item => (!state.ids || state.ids.includes(item.row['__id'])) &&
          state.conditions.every(c => matches(item.row, c)));
        if (state.limit) items = items.slice(0, state.limit);
        return iterator(items.map(item => makeEntity(item.value)));
      },
    };
    return selector;
  }

  _campaignSelector(channels) {
    return this._selector(
      () => this.campaigns.filter(c => channels.includes(c.channel)).map(c => ({ value: c, row: campaignRow(c) })),
      campaign => this._campaignEntity(campaign),
    );
  }

  _campaignEntity(campaign) {
    const ads = this;
    return {
      getId: () => Number(campaign.id),
      getName: () => campaign.name,
      isEnabled: () => campaign.status === 'ENABLED',
      createNegativeKeyword(text) {
        ads.entityNegatives.push({ level: 'campaign', id: campaign.id, text });
        if (!ads.preview) {
          const { text: plain, matchType } = parseKeywordSyntax(text);
          ads.addCampaignNegative({ campaignId: campaign.id, text: plain, matchType });
        }
      },
      negativeKeywords: () => ads._selector(
        () => (ads.rows.campaign_criterion || []).filter(r => r['campaign.id'] === campaign.id)
          .map(r => ({ value: r, row: r })),
        r => ({ getText: () => formatKeyword(r['campaign_criterion.keyword.text'], r['campaign_criterion.keyword.match_type']), getMatchType: () => r['campaign_criterion.keyword.match_type'] }),
      ),
    };
  }

  _adGroupSelector(channels) {
    return this._selector(
      () => this.adGroups.filter(g => {
        const c = this.campaigns.find(x => x.id === g.campaignId);
        return c && channels.includes(c.channel);
      }).map(g => ({ value: g, row: { __id: g.id, 'ad_group.id': g.id, 'ad_group.name': g.name, 'campaign.id': g.campaignId } })),
      adGroup => ({
        getId: () => Number(adGroup.id),
        getName: () => adGroup.name,
        createNegativeKeyword: text => this.entityNegatives.push({ level: 'adgroup', id: adGroup.id, text }),
      }),
    );
  }

  _listSelector() {
    return this._selector(
      () => this.lists.filter(l => l.type === 'NEGATIVE_KEYWORDS').map(l => ({ value: l, row: { __id: l.id, 'shared_set.name': l.name, 'shared_set.id': l.id } })),
      list => ({
        getId: () => Number(list.id),
        getName: () => list.name,
        addNegativeKeyword: text => this._addToList(list, [text]),
        addNegativeKeywords: texts => this._addToList(list, texts),
      }),
    );
  }

  _addToList(list, texts) {
    this.listAdds.push({ list: list.name, keywords: texts.slice() });
    if (!this.preview) list.keywords.push(...texts);
  }

  // ------------------------------------------------ globals for the script

  /** The Google Ads Scripts globals, ready for loadScript(). */
  globals(extra = {}) {
    const ads = this;
    const AdsApp = {
      report: query => ads.report(query),
      mutate: op => ads.mutate(op),
      campaigns: () => ads._campaignSelector(['SEARCH', 'DISPLAY', 'VIDEO', 'DEMAND_GEN', 'MULTI_CHANNEL']),
      shoppingCampaigns: () => ads._campaignSelector(['SHOPPING']),
      performanceMaxCampaigns: () => ads._campaignSelector(['PERFORMANCE_MAX']),
      adGroups: () => ads._adGroupSelector(['SEARCH', 'DISPLAY']),
      shoppingAdGroups: () => ads._adGroupSelector(['SHOPPING']),
      negativeKeywordLists: () => ads._listSelector(),
      currentAccount: () => ({
        getCustomerId: () => ads.customerId,
        getName: () => ads.accountName,
        getTimeZone: () => ads.timeZone,
        getCurrencyCode: () => 'GBP',
      }),
      getExecutionInfo: () => ({ isPreview: () => ads.preview }),
    };
    if (this.withMutateAll) AdsApp.mutateAll = (ops, options) => ads.mutateAll(ops, options);
    return {
      AdsApp,
      UrlFetchApp: { fetch: (url, params) => ads._fetch(url, params) },
      SpreadsheetApp: {
        openById: id => { if (id !== ads.spreadsheet.id) throw new Error(`No spreadsheet ${id}`); return ads.spreadsheet; },
        openByUrl: url => { if (!String(url).includes(ads.spreadsheet.id)) throw new Error(`No spreadsheet ${url}`); return ads.spreadsheet; },
      },
      MailApp: { sendEmail: (...args) => ads.emails.push(args), getRemainingDailyQuota: () => 100 },
      Utilities: {
        sleep: () => {},
        formatDate: (date, _tz, format) => formatDate(date, format),
      },
      ...extra,
    };
  }

  _fetch(url, params) {
    this.fetches.push({ url, params });
    const keys = Object.keys(this.fetchRoutes).sort((a, b) => b.length - a.length);
    const key = keys.find(k => url === k || url.startsWith(k));
    if (!key) throw new Error(`Unexpected request ${url}`);
    let route = this.fetchRoutes[key];
    if (typeof route === 'function') route = route(url, params);
    if (route instanceof Error) throw route;
    const { status, body } = (route && typeof route === 'object' && 'status' in route) ? route : { status: 200, body: route };
    const text = typeof body === 'string' ? body : JSON.stringify(body);
    return { getResponseCode: () => status, getContentText: () => text };
  }
}

function metrics({ impressions, clicks, cost, conversions, value }) {
  return {
    'metrics.impressions': String(impressions), 'metrics.clicks': String(clicks),
    'metrics.cost_micros': String(Math.round(cost * 1e6)), 'metrics.conversions': String(conversions),
    'metrics.conversions_value': String(value),
  };
}

function campaignRow(c) {
  return { __id: c.id, 'campaign.id': c.id, 'campaign.name': c.name, 'campaign.status': c.status, 'campaign.advertising_channel_type': c.channel, Name: c.name, Status: c.status };
}

function parseKeywordSyntax(text) {
  const t = String(text).trim();
  if (t.startsWith('[') && t.endsWith(']')) return { text: t.slice(1, -1), matchType: 'EXACT' };
  if (t.startsWith('"') && t.endsWith('"')) return { text: t.slice(1, -1), matchType: 'PHRASE' };
  return { text: t, matchType: 'BROAD' };
}

function formatKeyword(text, matchType) {
  return matchType === 'EXACT' ? `[${text}]` : matchType === 'PHRASE' ? `"${text}"` : text;
}

/** Without a segment in SELECT, Google returns one row per selected attribute set, metrics added up. */
function aggregate(rows, fields) {
  if (fields.some(f => f.startsWith('segments.'))) return rows;
  if (!rows.some(r => r['segments.date'] !== undefined)) return rows;
  const keyFields = fields.filter(f => !f.startsWith('metrics.'));
  const groups = new Map();
  for (const row of rows) {
    const key = JSON.stringify(keyFields.map(f => row[f]));
    const group = groups.get(key);
    if (!group) { groups.set(key, { ...row }); continue; }
    for (const f of Object.keys(row)) if (f.startsWith('metrics.')) group[f] = String(Number(group[f] || 0) + Number(row[f] || 0));
  }
  return [...groups.values()];
}

function formatDate(date, format = 'yyyy-MM-dd') {
  const d = date instanceof Date ? date : new Date(date);
  const pad = n => String(n).padStart(2, '0');
  return String(format)
    .replace('yyyy', d.getUTCFullYear()).replace('MM', pad(d.getUTCMonth() + 1)).replace('dd', pad(d.getUTCDate()))
    .replace('HH', pad(d.getUTCHours())).replace('mm', pad(d.getUTCMinutes())).replace('ss', pad(d.getUTCSeconds()));
}

// ---------------------------------------------------------------- loading a script

/**
 * Runs a Google Ads script file in a fresh VM context with the given globals.
 * @param {string} path Script file.
 * @param {Object} globals From MockGoogleAds#globals().
 * @param {Array<string>} names Top-level names to return (functions, consts, classes).
 * @param {Object} [options] { replace: [[from, to], ...] } text replacements before loading.
 * @return {{ exports: Object, logs: Array<string>, warnings: Array<string>, errors: Array<string> }}
 */
function loadScript(path, globals, names = [], options = {}) {
  let source = readFileSync(path, 'utf8');
  for (const [from, to] of options.replace || []) {
    if (!source.includes(from)) throw new Error(`loadScript: "${from}" not found in ${path}`);
    source = source.replace(from, to);
  }
  const logs = []; const warnings = []; const errors = [];
  const text = args => args.map(a => (typeof a === 'string' ? a : JSON.stringify(a))).join(' ');
  const consoleMock = {
    log: (...a) => logs.push(text(a)),
    info: (...a) => logs.push(text(a)),
    warn: (...a) => { warnings.push(text(a)); logs.push(text(a)); },
    error: (...a) => { errors.push(text(a)); logs.push(text(a)); },
  };
  const exportsExpr = `({ ${names.map(n => `${n}: typeof ${n} === 'undefined' ? undefined : ${n}`).join(', ')} })`;
  const exports = runInNewContext(`${source}\n;${exportsExpr}`, { console: consoleMock, ...globals }, { filename: path });
  return { exports, logs, warnings, errors };
}

module.exports = {
  MockGoogleAds, MockSpreadsheet, MockSheet, MockMutateResult,
  loadScript, parseGaql, iterator, likeToRegExp,
};
