# Test support: Google Ads Scripts mocks

`google-ads-scripts-mock.cjs` runs a Google Ads script in Node against a mocked account, so whole runs (`main()`) can be tested without Google. Use it for any script in this repository.

```js
const { MockGoogleAds, loadScript } = require('../test-support/google-ads-scripts-mock.cjs');
const ads = new MockGoogleAds({ fetch: { 'https://api.example/rules': [/* JSON body */] } });
ads.addCampaign({ id: '500', name: 'PMax', channel: 'PERFORMANCE_MAX' });
ads.addPmaxSearchTerm({ term: 'cheap socks', campaignId: '500', impressions: 30 });
const { exports, logs, errors } = loadScript('my_script.js', ads.globals(), ['main'], {
  replace: [['YOUR_SPREADSHEET_URL_HERE', 'https://docs.google.com/spreadsheets/d/demo-sheet/edit']],
});
exports.main();
// Then check ads.queries, ads.mutations, ads.entityNegatives, ads.listAdds, ads.lists,
// ads.spreadsheet.getSheetByName('1').rowsFrom(1), ads.emails, ads.fetches.
```

The file's header lists what is imitated and where it differs from Google (for example, `REGEXP_MATCH` is a partial JavaScript match, not RE2). It is built from what Shabba's probes of the real API found (`shabba-tools-data/tools-search-term-trends/docs/contract.md`): LIKE ignores case, `[x]` is a literal, rows are added up when no segment is selected, and keyword negatives are limited to 80 characters and 10 words.

When the real Google Ads behaves differently from the mock, fix the mock first and add a test to `google-ads-scripts-mock.test.cjs`.

Run: `node --test ai_auto_negative_keywords test-support` from the repository root (some older folders use other test runners).
