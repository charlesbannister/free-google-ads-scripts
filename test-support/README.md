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

## Wanted additions (from the 2026-11-25 review of all live scripts)

Each script's test file currently fills these gaps locally. Moving them here would let new tests reuse them:

- Sheets: `appendRow`, `autoResizeColumn(s)`, `setName`, `setTabColor`, `setFrozenColumns`, `sort`; ranges `clear`, `sort`, `getFilter`, `insertCheckboxes`, `setDataValidation`, `clearDataValidations`, whole-row ranges (`'1:1'`), named ranges; `SpreadsheetApp.create`/`newDataValidation`; throwing on any range outside the grid (reads too); `insertRowsAfter` shifting rows.
- Ads: `shopping_performance_view`, `performance_max_placement_view`, `detail_placement_view`, `product_group_view`, `ad_group_ad_asset_view`; date comparisons on quoted dates; an AWQL (`DURING`) stub; `AdsApp.keywords()`, `productGroups()`, `labels()`, `videoCampaigns()`, `excludedPlacementLists()`, `bulkUploads()`, `newNegativeKeywordListBuilder`; Google's keyword limits on entity `createNegativeKeyword`; `ad_group` report rows from `addAdGroup`.
- Globals: `Logger`, `AdWordsApp`, `MccApp`, `DriveApp`; `Utilities.formatDate` honouring the time zone; exports of `MockRange`, `matches` and `RESOURCE_FIELDS`.
