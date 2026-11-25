'use strict';
// End-to-end runs of the whole script (main) against the mocked Google Ads Scripts runtime
// in ../test-support. Run with: node --test full_negative_keyword_list_alert/
const assert = require('node:assert/strict');
const { join } = require('node:path');
const { test } = require('node:test');
const { MockGoogleAds, MockSheet, loadScript } = require('../test-support/google-ads-scripts-mock.cjs');

const SCRIPT = join(__dirname, 'full_negative_keyword_list_alert.js');
const SHEET = 'demo-sheet';
const CUSTOMER = '1234567890';

// Sheet calls this script makes that the shared mock does not have yet.
MockSheet.prototype.appendRow ??= function (values) {
  this.getRange(this.getLastRow() + 1, 1, 1, values.length).setValues([values]);
  return this;
};
const baseGetRange = MockSheet.prototype.getRange;
MockSheet.prototype.getRange = function (a, ...rest) {
  const rowOnly = typeof a === 'string' && a.match(/^(\d+):(\d+)$/); // whole-row notation such as '1:1'
  if (rowOnly) return baseGetRange.call(this, +rowOnly[1], 1, +rowOnly[2] - +rowOnly[1] + 1, this.getMaxColumns());
  return baseGetRange.call(this, a, ...rest);
};

/** Adds a shared negative keyword list with `count` keywords as shared_criterion rows. */
function addList(ads, { id, name, count, status = 'ENABLED', type = 'NEGATIVE_KEYWORDS' }) {
  const rows = [];
  for (let i = 0; i < count; i++) {
    rows.push({
      'shared_criterion.shared_set': `customers/${CUSTOMER}/sharedSets/${id}`,
      'shared_criterion.keyword.text': `junk ${i}`,
      'shared_set.name': name, 'shared_set.type': type, 'shared_set.status': status, 'shared_set.id': String(id),
    });
  }
  ads.addRows('shared_criterion', rows);
}

function account() {
  const ads = new MockGoogleAds({ customerId: '123-456-7890', spreadsheetId: SHEET });
  addList(ads, { id: 11, name: 'Full list', count: 5000 });
  addList(ads, { id: 12, name: 'Half list', count: 2500 });
  return ads;
}

function run(ads, replace = [['YOUR_SPREADSHEET_URL_HERE', `https://docs.google.com/spreadsheets/d/${SHEET}/edit`], ['YOUR_EMAIL_HERE', 'me@example.com']]) {
  const script = loadScript(SCRIPT, ads.globals(), ['main'], { replace });
  script.exports.main();
  return script;
}

const tracked = ads => ads.spreadsheet.getSheetByName('Alerted Lists').rowsFrom(2);

test('a full list is alerted once by email and recorded in the sheet', () => {
  const ads = account();
  run(ads);
  assert.equal(ads.emails.length, 1);
  const [email] = ads.emails[0];
  assert.equal(email.to, 'me@example.com');
  assert.match(email.htmlBody, /Full list/);
  assert.doesNotMatch(email.htmlBody, /Half list/);
  assert.deepEqual(tracked(ads).map(r => r.slice(0, 2)), [[`customers/${CUSTOMER}/sharedSets/11`, 'Full list']]);
  assert.deepEqual(ads.spreadsheet.getSheetByName('Alerted Lists').rowsFrom(1)[0], ['List ID', 'List Name', 'Alert Date']);
});

test('a second run does not alert the same list again', () => {
  const ads = account();
  run(ads);
  run(ads);
  assert.equal(ads.emails.length, 1);
  assert.equal(tracked(ads).length, 1);
});

test('a list that fills later is alerted on its own', () => {
  const ads = account();
  run(ads);
  addList(ads, { id: 13, name: 'Newly full', count: 5000 });
  run(ads);
  assert.equal(ads.emails.length, 2);
  assert.match(ads.emails[1][0].htmlBody, /Newly full/);
  assert.doesNotMatch(ads.emails[1][0].htmlBody, /Full list/);
});

test('a removed list is not counted or alerted', () => {
  const ads = new MockGoogleAds({ spreadsheetId: SHEET });
  addList(ads, { id: 21, name: 'Old removed list', count: 5000, status: 'REMOVED' });
  run(ads);
  assert.equal(ads.emails.length, 0);
  assert.match(ads.queries[0], /shared_set\.status = 'ENABLED'/);
});

test('placement exclusion lists are ignored', () => {
  const ads = new MockGoogleAds({ spreadsheetId: SHEET });
  addList(ads, { id: 31, name: 'Placements', count: 6000, type: 'NEGATIVE_PLACEMENTS' });
  run(ads);
  assert.equal(ads.emails.length, 0);
});

test('empty account: no email, tracking sheet created with headers only', () => {
  const ads = new MockGoogleAds({ spreadsheetId: SHEET });
  const { logs } = run(ads);
  assert.equal(ads.emails.length, 0);
  assert.ok(logs.includes('No negative keyword lists found.'));
  assert.equal(tracked(ads).length, 0);
});

test('a tracking sheet with extra user rows and numbers still works', () => {
  const ads = account();
  const sheet = ads.spreadsheet.insertSheet('Alerted Lists');
  sheet.appendRow(['List ID', 'List Name', 'Alert Date']);
  sheet.appendRow([12345, 'something typed by hand', '']);
  sheet.appendRow(['', '', '']);
  run(ads);
  assert.equal(ads.emails.length, 1);
});

test('report failure stops the run before any email or sheet write', () => {
  const ads = account();
  ads.failQuery('shared_criterion', 'Internal error');
  assert.throws(() => run(ads), /Internal error/);
  assert.equal(ads.emails.length, 0);
  assert.equal(ads.spreadsheet.getSheetByName('Alerted Lists'), null);
});

test('placeholder spreadsheet URL gives a clear error', () => {
  const ads = account();
  assert.throws(() => run(ads, []), /Please set the SPREADSHEET_URL/);
  assert.equal(ads.emails.length, 0);
});
