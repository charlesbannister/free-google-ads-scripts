/**
 * Keyword-Based Negatives Script
 * Automatically identifies potential negative keywords by comparing search terms
 * against the account's actual keywords within the same ad group.
 * 
 * @authors Charles Bannister & Gabriele Benedetti
 * @version 1.4.0
 * 
 * Google Ads API Query Builder Links:
 * - Search Term View: https://developers.google.com/google-ads/api/fields/v20/search_term_view_query_builder
 * - Keyword View: https://developers.google.com/google-ads/api/fields/v20/keyword_view_query_builder
 * - Ad Group: https://developers.google.com/google-ads/api/fields/v20/ad_group_query_builder
 */

// ============================================================================
// CONFIGURATION - Update these values
// ============================================================================

const SPREADSHEET_URL = 'PASTE_YOUR_SPREADSHEET_URL_HERE';
// The Google Sheet URL where results will be written
// Create a copy of the template: https://docs.google.com/spreadsheets/d/1x3XcN3EwJzo4RzDXphtSINjC5yKIhCd81JJ00CLwQT0/copy
// Then paste the URL of your copy here

const OUTPUT_SHEET_NAME = 'Output';
// Sheet name for potential negatives (with checkboxes)

const LOGS_SHEET_NAME = 'Logs';
// Sheet name for applied negatives log (append only)

const LOOKBACK_DAYS = 30;
// Number of days to look back for search term data

const FULL_AUTOMATE_MODE = false;
// When true: automatically check all boxes and apply negatives immediately
// When false: write to sheet with unchecked boxes for manual review

const DEBUG_MODE = false;
// Set to true to see detailed logs for debugging

const RUN_TEST_SUITE = false;
// Set to true to run the matching logic test suite before the main script
// Tests will show ✓ (pass) or ✗ (fail) for each case
// Set to false to skip tests and run the script normally

// --- Performance Filters ---
const MIN_CLICKS = 0;
// Only include search terms with clicks greater than this value

const MIN_IMPRESSIONS = 1;
// Only include search terms with impressions greater than this value

const MAX_CONVERSIONS = 0;
// Only include search terms with conversions less than or equal to this value
// Set to 0 to only include non-converting search terms

// --- Matching Settings ---
const FUZZY_MATCH_THRESHOLD = 70;
// Similarity threshold for fuzzy matching (0-100)
// Higher = stricter matching, Lower = more lenient

// --- Campaign Filters ---
const CAMPAIGN_NAME_CONTAINS = [];
// Only include campaigns containing ANY of these strings
// Example: ['Brand', 'Search'] - includes campaigns with "Brand" OR "Search"
// Leave empty [] to include all campaigns

const CAMPAIGN_NAME_NOT_CONTAINS = [];
// Exclude campaigns containing ANY of these strings
// Example: ['Test', 'Old'] - excludes campaigns with "Test" OR "Old"
// Leave empty [] to not exclude any campaigns

// --- Ad Group Filters ---
const AD_GROUP_NAME_CONTAINS = [];
// Only include ad groups containing ANY of these strings
// Example: ['Exact', 'Phrase'] - includes ad groups with "Exact" OR "Phrase"
// Leave empty [] to include all ad groups

const AD_GROUP_NAME_NOT_CONTAINS = [];
// Exclude ad groups containing ANY of these strings
// Example: ['DSA', 'Dynamic'] - excludes ad groups with "DSA" OR "Dynamic"
// Leave empty [] to not exclude any ad groups

// --- Email Settings ---
const EMAIL_RECIPIENTS = '';
// Comma-separated email addresses for notifications
// Example: 'email1@example.com, email2@example.com'
// Leave empty '' to disable email notifications

// --- Negative Keyword Settings ---
const NEGATIVE_MATCH_TYPE = 'EXACT';
// Match type for applied negatives: 'EXACT', 'PHRASE', or 'BROAD'

// ============================================================================
// TEST SUITE
// ============================================================================

/**
 * Test cases for verifying the fuzzy matching logic.
 * Each test has: searchTerm, keywords array, and expected shouldExclude result.
 * shouldExclude: true = search term should be excluded (added as negative)
 * shouldExclude: false = search term should be kept (matches keywords)
 */
const TEST_CASES = [
  // Exact and near-exact matches
  { searchTerm: 'running shoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'nike running shoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // No space variations
  { searchTerm: 'runningshoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'runningshoe', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'r u n n i n g s h o e', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // Singular vs plural
  { searchTerm: 'running shoe', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'shoe running', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // Misspellings - running
  { searchTerm: 'runing shoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'runnign shoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'runnnig shoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'runnin shoes', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // Misspellings - shoes
  { searchTerm: 'running sheos', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'running shoess', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'running shose', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // Combined misspellings - now matches via shared word check (shoes ~ sheos)
  { searchTerm: 'runing sheos', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'runnig shoe', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // Extra words (should still match via sliding window)
  { searchTerm: 'best running shoes 2024', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },
  { searchTerm: 'cheap running shoes online', keywords: ['running shoes', 'nike running shoes'], shouldExclude: false },

  // Subset of keyword words (search term words exist in longer keyword)
  { searchTerm: 'osmosis water filter', keywords: ['reverse osmosis water filter'], shouldExclude: false },
  { searchTerm: 'running shoes', keywords: ['best running shoes', 'nike running shoes'], shouldExclude: false },

  // Shared word match (any keyword word exists in search term)
  { searchTerm: 'reverse osmosis ro', keywords: ['ro system'], shouldExclude: false },
  { searchTerm: 'dental autoclave repair', keywords: ['autoclave servicing'], shouldExclude: false },

  // Completely different (should exclude)
  { searchTerm: 'hiking boots', keywords: ['running shoes', 'nike running shoes'], shouldExclude: true },
  { searchTerm: 'basketball sneakers', keywords: ['running shoes', 'nike running shoes'], shouldExclude: true },
];

/**
 * Runs the test suite to verify the fuzzy matching logic.
 * Returns true if all tests pass, false otherwise.
 * @returns {boolean} True if all tests pass
 */
function runTestSuite() {
  console.log('');
  console.log('🧪 ============================================');
  console.log('🧪 RUNNING MATCHING LOGIC TEST SUITE');
  console.log('🧪 ============================================');
  console.log(`Threshold: ${FUZZY_MATCH_THRESHOLD}%`);
  console.log('');

  const matcher = new SearchTermMatcher(false);
  let passed = 0;
  let failed = 0;

  for (const testCase of TEST_CASES) {
    const result = testSearchTermMatch(testCase.searchTerm, testCase.keywords, matcher);
    const actual = !result.isMatch; // shouldExclude = NOT isMatch
    const expected = testCase.shouldExclude;
    const testPassed = actual === expected;

    if (testPassed) {
      passed++;
      console.log(`✓ ${result.isMatch ? 'KEEP' : 'EXCLUDE'} | "${testCase.searchTerm}"`);
    } else {
      failed++;
      console.log(`✗ ${result.isMatch ? 'KEEP' : 'EXCLUDE'} | "${testCase.searchTerm}"`);
      console.log(`   ⚠️  EXPECTED: ${expected ? 'EXCLUDE' : 'KEEP'}, GOT: ${actual ? 'EXCLUDE' : 'KEEP'}`);
    }
    console.log(`   Keywords: ${testCase.keywords.join(', ')}`);
    console.log(`   Score: ${result.bestScore.toFixed(1)}%`);
    console.log('');
  }

  console.log('🧪 ============================================');
  console.log(`🧪 TEST RESULTS: ${passed} passed, ${failed} failed (${TEST_CASES.length} total)`);
  console.log('🧪 ============================================');
  console.log('');

  if (failed > 0) {
    console.log('❌ TESTS FAILED - Fix matching logic before running script');
    return false;
  }

  console.log('✅ ALL TESTS PASSED - Matching logic verified');
  return true;
}

/**
 * Tests a single search term against keywords (used by test suite).
 * Similar to findBestKeywordMatch but accepts string keywords array.
 * @param {string} searchTerm - The search term to test
 * @param {Array<string>} keywords - Array of keyword strings
 * @param {SearchTermMatcher} matcher - The matcher instance
 * @returns {Object} Object with isMatch boolean and bestScore number
 */
function testSearchTermMatch(searchTerm, keywords, matcher) {
  let bestScore = 0;

  for (const keyword of keywords) {
    const condition = {
      text: keyword,
      matchType: 'approx-contains',
      threshold: FUZZY_MATCH_THRESHOLD
    };

    const isMatch = matcher.matchesCondition(searchTerm, condition);
    const score = calculateBestScore(keyword, searchTerm);
    bestScore = Math.max(bestScore, score);

    if (isMatch) {
      return { isMatch: true, bestScore: bestScore };
    }
  }

  return { isMatch: false, bestScore: bestScore };
}

// ============================================================================
// MAIN FUNCTION
// ============================================================================

/**
 * Main entry point for the script.
 * Orchestrates the entire workflow of finding and applying negative keywords.
 */
function main() {
  // Run test suite first if enabled
  if (RUN_TEST_SUITE) {
    const testsPass = runTestSuite();
    if (!testsPass) {
      console.log('Script stopped due to failing tests.');
      return;
    }
    console.log('Continuing with main script...');
    console.log('');
  }
  console.log('=== Keyword-Based Negatives Script Started ===');
  console.log(`Version: 1.4.0`);
  console.log(`Full Automate Mode: ${FULL_AUTOMATE_MODE}`);
  console.log(`Debug Mode: ${DEBUG_MODE}`);

  validateConfiguration();

  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const outputSheet = getOrCreateSheet(spreadsheet, OUTPUT_SHEET_NAME);
  const logsSheet = getOrCreateSheet(spreadsheet, LOGS_SHEET_NAME);

  ensureLogsSheetHeaders(logsSheet);

  const executionMode = getExecutionMode();
  console.log(`Execution Mode: ${executionMode}`);

  // Step 1: Process any checked items from previous run
  const appliedNegatives = processCheckedItems(outputSheet, logsSheet, executionMode);

  // Step 2: Clear the output sheet after processing
  clearOutputSheetData(outputSheet);

  // Step 3: Find potential negatives
  const potentialNegatives = findAllPotentialNegatives();

  if (potentialNegatives.length === 0) {
    console.log('No potential negatives found.');
    writeNoResultsMessage(outputSheet);
    console.log('=== Script Finished ===');
    return;
  }

  console.log(`Found ${potentialNegatives.length} potential negatives`);

  // Step 4: Write to output sheet with checkboxes
  writeToOutputSheet(outputSheet, potentialNegatives);

  // Step 5: Send "Found" email notification
  if (EMAIL_RECIPIENTS && potentialNegatives.length > 0) {
    sendFoundEmail(potentialNegatives, spreadsheet.getUrl());
  }

  // Step 6: If FULL_AUTOMATE_MODE, check all boxes and apply
  if (FULL_AUTOMATE_MODE) {
    console.log('Full Automate Mode enabled - applying all negatives');
    checkAllBoxes(outputSheet, potentialNegatives.length);
    const autoAppliedNegatives = applyNegativesFromData(potentialNegatives, logsSheet, executionMode);

    // Send "Added" email (only in live mode)
    if (EMAIL_RECIPIENTS && autoAppliedNegatives.length > 0 && executionMode === 'Live') {
      sendAddedEmail(autoAppliedNegatives, spreadsheet.getUrl());
    }
  }

  // Send "Added" email for manually checked items (only in live mode)
  if (EMAIL_RECIPIENTS && appliedNegatives.length > 0 && executionMode === 'Live') {
    sendAddedEmail(appliedNegatives, spreadsheet.getUrl());
  }

  console.log('=== Script Finished ===');
}

// ============================================================================
// VALIDATION
// ============================================================================

/**
 * Validates the script configuration.
 * Throws an error if required settings are missing or invalid.
 * @throws {Error} If configuration is invalid
 */
function validateConfiguration() {
  if (!SPREADSHEET_URL || SPREADSHEET_URL === 'PASTE_YOUR_SPREADSHEET_URL_HERE') {
    throw new Error('Please update SPREADSHEET_URL with your actual spreadsheet URL. Create a copy from the template: https://docs.google.com/spreadsheets/d/1x3XcN3EwJzo4RzDXphtSINjC5yKIhCd81JJ00CLwQT0/copy');
  }

  if (LOOKBACK_DAYS < 1 || LOOKBACK_DAYS > 540) {
    throw new Error('LOOKBACK_DAYS must be between 1 and 540');
  }

  if (FUZZY_MATCH_THRESHOLD < 0 || FUZZY_MATCH_THRESHOLD > 100) {
    throw new Error('FUZZY_MATCH_THRESHOLD must be between 0 and 100');
  }

  const validMatchTypes = ['EXACT', 'PHRASE', 'BROAD'];
  if (!validMatchTypes.includes(NEGATIVE_MATCH_TYPE.toUpperCase())) {
    throw new Error(`NEGATIVE_MATCH_TYPE must be one of: ${validMatchTypes.join(', ')}`);
  }

  console.log('Configuration validated successfully');
}

// ============================================================================
// EXECUTION MODE
// ============================================================================

/**
 * Determines the current execution mode (Live or Preview).
 * @returns {string} 'Live' or 'Preview'
 */
function getExecutionMode() {
  const isPreview = AdsApp.getExecutionInfo().isPreview();
  return isPreview ? 'Preview' : 'Live';
}

// ============================================================================
// FIND POTENTIAL NEGATIVES
// ============================================================================

/**
 * Main function to find all potential negatives across filtered ad groups.
 * @returns {Array<Object>} Array of potential negative objects
 */
function findAllPotentialNegatives() {
  console.log('\n--- Finding Potential Negatives ---');

  const adGroups = getFilteredAdGroups();
  console.log(`Found ${adGroups.length} ad groups matching filters`);

  if (adGroups.length === 0) {
    console.log('No ad groups match the specified filters');
    return [];
  }

  const allPotentialNegatives = [];

  for (const adGroup of adGroups) {
    debugLog(`Processing ad group: ${adGroup.adGroupName} (${adGroup.adGroupId})`);

    const adGroupNegatives = findPotentialNegativesForAdGroup(adGroup);
    allPotentialNegatives.push(...adGroupNegatives);

    debugLog(`Found ${adGroupNegatives.length} potential negatives in ad group ${adGroup.adGroupName}`);
  }

  return allPotentialNegatives;
}

/**
 * Finds potential negatives for a single ad group.
 * @param {Object} adGroup - Ad group object with id, name, campaignId, campaignName
 * @returns {Array<Object>} Array of potential negative objects for this ad group
 */
function findPotentialNegativesForAdGroup(adGroup) {
  const keywords = getKeywordsForAdGroup(adGroup.adGroupId);

  if (keywords.length === 0) {
    debugLog(`No keywords found in ad group ${adGroup.adGroupName} - skipping`);
    return [];
  }

  debugLog(`Found ${keywords.length} keywords in ad group ${adGroup.adGroupName}`);

  const searchTerms = getSearchTermsForAdGroup(adGroup.adGroupId, adGroup.campaignId);

  if (searchTerms.length === 0) {
    debugLog(`No search terms found for ad group ${adGroup.adGroupName}`);
    return [];
  }

  debugLog(`Found ${searchTerms.length} search terms for ad group ${adGroup.adGroupName}`);

  const potentialNegatives = compareSearchTermsToKeywords(searchTerms, keywords, adGroup);

  const filteredNegatives = filterByPerformance(potentialNegatives);

  return filteredNegatives;
}

// ============================================================================
// GET AD GROUPS
// ============================================================================

/**
 * Gets all ad groups matching the campaign and ad group name filters.
 * @returns {Array<Object>} Array of ad group objects
 */
function getFilteredAdGroups() {
  const query = buildAdGroupQuery();
  debugLog(`Ad Group Query: ${query}`);

  const adGroups = [];

  try {
    const report = AdsApp.report(query);
    const rows = report.rows();

    while (rows.hasNext()) {
      const row = rows.next();

      const campaignName = row['campaign.name'];
      const adGroupName = row['ad_group.name'];

      if (!passesNameFilters(campaignName, adGroupName)) {
        continue;
      }

      adGroups.push({
        campaignId: row['campaign.id'],
        campaignName: campaignName,
        adGroupId: row['ad_group.id'],
        adGroupName: adGroupName
      });
    }
  } catch (error) {
    console.error(`Error fetching ad groups: ${error.message}`);
    throw error;
  }

  return adGroups;
}

/**
 * Builds the GAQL query to fetch ad groups.
 * @returns {string} GAQL query string
 */
function buildAdGroupQuery() {
  const query = `
    SELECT
      campaign.id,
      campaign.name,
      ad_group.id,
      ad_group.name
    FROM ad_group
    WHERE campaign.status = 'ENABLED'
    AND ad_group.status = 'ENABLED'
  `;

  return query.trim();
}

/**
 * Checks if campaign and ad group names pass the configured filters.
 * @param {string} campaignName - Campaign name to check
 * @param {string} adGroupName - Ad group name to check
 * @returns {boolean} True if passes all filters
 */
function passesNameFilters(campaignName, adGroupName) {
  const campaignNameLower = campaignName.toLowerCase();
  const adGroupNameLower = adGroupName.toLowerCase();

  // Check CAMPAIGN_NAME_CONTAINS (if any match, pass)
  if (CAMPAIGN_NAME_CONTAINS.length > 0) {
    const containsMatch = CAMPAIGN_NAME_CONTAINS.some(
      filter => campaignNameLower.includes(filter.toLowerCase())
    );
    if (!containsMatch) return false;
  }

  // Check CAMPAIGN_NAME_NOT_CONTAINS (if any match, fail)
  if (CAMPAIGN_NAME_NOT_CONTAINS.length > 0) {
    const notContainsMatch = CAMPAIGN_NAME_NOT_CONTAINS.some(
      filter => campaignNameLower.includes(filter.toLowerCase())
    );
    if (notContainsMatch) return false;
  }

  // Check AD_GROUP_NAME_CONTAINS (if any match, pass)
  if (AD_GROUP_NAME_CONTAINS.length > 0) {
    const containsMatch = AD_GROUP_NAME_CONTAINS.some(
      filter => adGroupNameLower.includes(filter.toLowerCase())
    );
    if (!containsMatch) return false;
  }

  // Check AD_GROUP_NAME_NOT_CONTAINS (if any match, fail)
  if (AD_GROUP_NAME_NOT_CONTAINS.length > 0) {
    const notContainsMatch = AD_GROUP_NAME_NOT_CONTAINS.some(
      filter => adGroupNameLower.includes(filter.toLowerCase())
    );
    if (notContainsMatch) return false;
  }

  return true;
}

// ============================================================================
// GET KEYWORDS
// ============================================================================

/**
 * Gets all enabled keywords for a specific ad group, sorted by clicks descending.
 * @param {string} adGroupId - The ad group ID
 * @returns {Array<Object>} Array of keyword objects with text, matchType, and clicks
 */
function getKeywordsForAdGroup(adGroupId) {
  const dateRange = getDateRangeCondition(LOOKBACK_DAYS);

  const query = `
    SELECT
      ad_group_criterion.keyword.text,
      ad_group_criterion.keyword.match_type,
      metrics.clicks
    FROM keyword_view
    WHERE campaign.status = 'ENABLED'
    AND ad_group.status = 'ENABLED'
    AND ad_group_criterion.status = 'ENABLED'
    AND ad_group_criterion.negative = FALSE
    AND ad_group.id = '${adGroupId}'
    AND ${dateRange}
    ORDER BY metrics.clicks DESC
  `;

  const keywords = [];

  try {
    const report = AdsApp.report(query);
    const rows = report.rows();

    while (rows.hasNext()) {
      const row = rows.next();
      keywords.push({
        text: row['ad_group_criterion.keyword.text'],
        matchType: row['ad_group_criterion.keyword.match_type'],
        clicks: parseInt(row['metrics.clicks'] || 0)
      });
    }
  } catch (error) {
    console.error(`Error fetching keywords for ad group ${adGroupId}: ${error.message}`);
  }

  // Sort by clicks descending (in case ORDER BY didn't work)
  keywords.sort((a, b) => b.clicks - a.clicks);

  return keywords;
}

/**
 * Formats keywords array into a display string.
 * Shows first 10 keywords ordered by clicks, comma-separated.
 * Adds ellipsis if there are more than 10.
 * @param {Array<Object>} keywords - Array of keyword objects with text and clicks
 * @returns {string} Formatted keywords string
 */
function formatKeywordsString(keywords) {
  if (!keywords || keywords.length === 0) {
    return '';
  }

  const maxKeywords = 10;
  const keywordTexts = keywords.slice(0, maxKeywords).map(kw => kw.text);

  let result = keywordTexts.join(', ');

  if (keywords.length > maxKeywords) {
    result += '...';
  }

  return result;
}

// ============================================================================
// GET SEARCH TERMS
// ============================================================================

/**
 * Gets search terms for a specific ad group within the lookback period.
 * @param {string} adGroupId - The ad group ID
 * @param {string} campaignId - The campaign ID
 * @returns {Array<Object>} Array of search term objects with metrics
 */
function getSearchTermsForAdGroup(adGroupId, campaignId) {
  const dateRange = getDateRangeCondition(LOOKBACK_DAYS);

  const query = `
    SELECT
      search_term_view.search_term,
      metrics.impressions,
      metrics.clicks,
      metrics.cost_micros,
      metrics.conversions,
      metrics.conversions_value
    FROM search_term_view
    WHERE ${dateRange}
    AND campaign.status = 'ENABLED'
    AND ad_group.status = 'ENABLED'
    AND ad_group.id = '${adGroupId}'
    AND campaign.id = '${campaignId}'
    AND metrics.impressions > 0
  `;

  const searchTerms = [];

  try {
    const report = AdsApp.report(query);
    const rows = report.rows();

    while (rows.hasNext()) {
      const row = rows.next();

      const impressions = parseInt(row['metrics.impressions'] || 0);
      const clicks = parseInt(row['metrics.clicks'] || 0);
      const costMicros = parseInt(row['metrics.cost_micros'] || 0);
      const conversions = parseFloat(row['metrics.conversions'] || 0);
      const conversionsValue = parseFloat(row['metrics.conversions_value'] || 0);

      const cost = costMicros / 1000000;
      const ctr = impressions > 0 ? clicks / impressions : 0;
      const conversionRate = clicks > 0 ? conversions / clicks : 0;
      const cpa = conversions > 0 ? cost / conversions : (cost > 0 ? Infinity : 0);
      const roas = cost > 0 ? conversionsValue / cost : 0;

      searchTerms.push({
        searchTerm: row['search_term_view.search_term'],
        impressions: impressions,
        clicks: clicks,
        cost: cost,
        conversions: conversions,
        conversionsValue: conversionsValue,
        ctr: ctr,
        conversionRate: conversionRate,
        cpa: cpa,
        roas: roas
      });
    }
  } catch (error) {
    console.error(`Error fetching search terms for ad group ${adGroupId}: ${error.message}`);
  }

  return searchTerms;
}

// ============================================================================
// COMPARE SEARCH TERMS TO KEYWORDS
// ============================================================================

/**
 * Compares search terms against keywords using fuzzy matching.
 * Returns search terms that don't match any keyword.
 * @param {Array<Object>} searchTerms - Array of search term objects
 * @param {Array<Object>} keywords - Array of keyword objects
 * @param {Object} adGroup - Ad group info object
 * @returns {Array<Object>} Array of potential negative objects
 */
function compareSearchTermsToKeywords(searchTerms, keywords, adGroup) {
  const matcher = new SearchTermMatcher(DEBUG_MODE);
  const potentialNegatives = [];
  const currentDate = new Date().toISOString().split('T')[0];
  const keywordsString = formatKeywordsString(keywords);

  for (const searchTermData of searchTerms) {
    const searchTerm = searchTermData.searchTerm;

    const matchResult = findBestKeywordMatch(searchTerm, keywords, matcher);

    // If no keyword matches above threshold, it's a potential negative
    if (!matchResult.isMatch) {
      potentialNegatives.push({
        searchTerm: searchTerm,
        keywords: keywordsString,
        campaignName: adGroup.campaignName,
        campaignId: adGroup.campaignId,
        adGroupName: adGroup.adGroupName,
        adGroupId: adGroup.adGroupId,
        impressions: searchTermData.impressions,
        clicks: searchTermData.clicks,
        cost: searchTermData.cost,
        conversions: searchTermData.conversions,
        conversionsValue: searchTermData.conversionsValue,
        ctr: searchTermData.ctr,
        conversionRate: searchTermData.conversionRate,
        cpa: searchTermData.cpa,
        roas: searchTermData.roas,
        bestMatchScore: matchResult.bestScore,
        dateFound: currentDate
      });
    }
  }

  return potentialNegatives;
}

/**
 * Calculates the best fuzzy match score using a sliding window approach.
 * Checks if the keyword exists as a substring within the search term.
 * @param {string} keyword - The keyword to search for
 * @param {string} searchTerm - The search term to search within
 * @returns {number} Best match score from 0 to 100
 */
function getSlidingWindowScore(keyword, searchTerm) {
  let bestScore = 0;
  let startPosition = 0;
  let endPosition = keyword.length;

  while (endPosition <= searchTerm.length) {
    const windowScore = fuzzyMatchScore(
      keyword,
      searchTerm.substring(startPosition, endPosition)
    );
    if (windowScore > bestScore) {
      bestScore = windowScore;
    }
    startPosition++;
    endPosition++;
  }

  return bestScore;
}

/**
 * Calculates the best score across all matching variations.
 * Must match all checks in approxMatchContains for consistency.
 * @param {string} keyword - The keyword text
 * @param {string} searchTerm - The search term text
 * @returns {number} Best match score from 0 to 100
 */
function calculateBestScore(keyword, searchTerm) {
  const normalScore = fuzzyMatchScore(keyword, searchTerm);
  const spacelessScore = fuzzyMatchScore(
    keyword.replace(/\s+/g, ''),
    searchTerm.replace(/\s+/g, '')
  );
  const sortedKeyword = keyword.toLowerCase().split(/\s+/).sort().join(' ');
  const sortedSearchTerm = searchTerm.toLowerCase().split(/\s+/).sort().join(' ');
  const sortedScore = fuzzyMatchScore(sortedKeyword, sortedSearchTerm);
  const slidingWindowScore = getSlidingWindowScore(keyword, searchTerm);

  // Subset words check - if all search term words exist in keyword
  const keywordWords = keyword.toLowerCase().split(/\s+/);
  const searchTermWords = searchTerm.toLowerCase().split(/\s+/);
  let subsetWordsScore = 0;
  const allWordsMatch = searchTermWords.every(searchWord =>
    keywordWords.some(keywordWord => fuzzyMatchScore(searchWord, keywordWord) >= 80)
  );
  if (allWordsMatch) {
    const wordScores = searchTermWords.map(searchWord => {
      const scores = keywordWords.map(keywordWord => fuzzyMatchScore(searchWord, keywordWord));
      return Math.max(...scores);
    });
    subsetWordsScore = wordScores.reduce((a, b) => a + b, 0) / wordScores.length;
  }

  // Shared word check - if any keyword word matches any search term word
  let sharedWordScore = 0;
  for (const keywordWord of keywordWords) {
    for (const searchWord of searchTermWords) {
      const wordScore = fuzzyMatchScore(keywordWord, searchWord);
      sharedWordScore = Math.max(sharedWordScore, wordScore);
    }
  }

  // Individual words check - if any search term word matches the whole keyword
  let individualWordScore = 0;
  for (const searchWord of searchTermWords) {
    const wordScore = fuzzyMatchScore(keyword, searchWord);
    individualWordScore = Math.max(individualWordScore, wordScore);
  }

  return Math.max(normalScore, spacelessScore, sortedScore, slidingWindowScore, subsetWordsScore, sharedWordScore, individualWordScore);
}

/**
 * Finds the best matching keyword for a search term.
 * @param {string} searchTerm - The search term to match
 * @param {Array<Object>} keywords - Array of keyword objects
 * @param {SearchTermMatcher} matcher - The matcher instance
 * @returns {Object} Object with isMatch boolean and bestScore number
 */
function findBestKeywordMatch(searchTerm, keywords, matcher) {
  let bestScore = 0;

  for (const keyword of keywords) {
    const condition = {
      text: keyword.text,
      matchType: 'approx-contains',
      threshold: FUZZY_MATCH_THRESHOLD
    };

    const isMatch = matcher.matchesCondition(searchTerm, condition);
    const score = calculateBestScore(keyword.text, searchTerm);
    bestScore = Math.max(bestScore, score);

    if (isMatch) {
      return { isMatch: true, bestScore: bestScore };
    }
  }

  return { isMatch: false, bestScore: bestScore };
}

// ============================================================================
// PERFORMANCE FILTERS
// ============================================================================

/**
 * Filters potential negatives by performance thresholds.
 * @param {Array<Object>} potentialNegatives - Array of potential negative objects
 * @returns {Array<Object>} Filtered array
 */
function filterByPerformance(potentialNegatives) {
  return potentialNegatives.filter(item => {
    if (item.clicks < MIN_CLICKS) return false;
    if (item.impressions < MIN_IMPRESSIONS) return false;
    if (item.conversions > MAX_CONVERSIONS) return false;
    return true;
  });
}

// ============================================================================
// PROCESS CHECKED ITEMS
// ============================================================================

/**
 * Processes items that were checked in the output sheet from a previous run.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} outputSheet - The output sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} logsSheet - The logs sheet
 * @param {string} executionMode - 'Live' or 'Preview'
 * @returns {Array<Object>} Array of applied negative objects
 */
function processCheckedItems(outputSheet, logsSheet, executionMode) {
  console.log('\n--- Processing Checked Items ---');

  const checkedItems = getCheckedItems(outputSheet);

  if (checkedItems.length === 0) {
    console.log('No checked items to process');
    return [];
  }

  console.log(`Found ${checkedItems.length} checked items to apply`);

  const appliedNegatives = applyNegativesFromData(checkedItems, logsSheet, executionMode);

  return appliedNegatives;
}

/**
 * Gets all checked items from the output sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The output sheet
 * @returns {Array<Object>} Array of checked item objects
 */
function getCheckedItems(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const dataRange = sheet.getRange(2, 1, lastRow - 1, 18);
  const data = dataRange.getValues();

  const checkedItems = [];

  for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
    const row = data[rowIndex];
    const isChecked = row[0] === true;

    if (!isChecked) continue;

    // Skip if search term is empty (might be a message row)
    if (!row[1]) continue;

    checkedItems.push({
      searchTerm: row[1],
      keywords: row[2],
      bestMatchScore: row[3],
      campaignName: row[4],
      adGroupName: row[5],
      campaignId: row[6],
      adGroupId: row[7],
      impressions: row[8],
      clicks: row[9],
      cost: row[10],
      conversions: row[11],
      conversionsValue: row[12],
      ctr: row[13],
      conversionRate: row[14],
      cpa: row[15],
      roas: row[16],
      dateFound: row[17]
    });
  }

  return checkedItems;
}

// ============================================================================
// APPLY NEGATIVES
// ============================================================================

/**
 * Applies negative keywords from the provided data.
 * @param {Array<Object>} items - Array of items to apply as negatives
 * @param {GoogleAppsScript.Spreadsheet.Sheet} logsSheet - The logs sheet
 * @param {string} executionMode - 'Live' or 'Preview'
 * @returns {Array<Object>} Array of successfully applied items
 */
function applyNegativesFromData(items, logsSheet, executionMode) {
  const appliedNegatives = [];
  const timestamp = new Date().toISOString();

  // Group items by ad group for efficiency
  const itemsByAdGroup = groupItemsByAdGroup(items);

  for (const adGroupId in itemsByAdGroup) {
    const adGroupItems = itemsByAdGroup[adGroupId];
    const firstItem = adGroupItems[0];

    for (const item of adGroupItems) {
      const result = applyNegativeKeyword(
        item.campaignId,
        item.adGroupId,
        item.searchTerm
      );

      // Log to sheet
      logAppliedNegative(logsSheet, {
        timestamp: timestamp,
        searchTerm: item.searchTerm,
        campaignName: item.campaignName,
        adGroupName: item.adGroupName,
        campaignId: item.campaignId,
        adGroupId: item.adGroupId,
        executionMode: executionMode,
        status: result.success ? 'Added' : result.error
      });

      if (result.success) {
        appliedNegatives.push(item);
      }
    }
  }

  console.log(`Applied ${appliedNegatives.length} of ${items.length} negatives`);

  return appliedNegatives;
}

/**
 * Groups items by ad group ID.
 * @param {Array<Object>} items - Array of items
 * @returns {Object} Object keyed by ad group ID
 */
function groupItemsByAdGroup(items) {
  const grouped = {};

  for (const item of items) {
    if (!grouped[item.adGroupId]) {
      grouped[item.adGroupId] = [];
    }
    grouped[item.adGroupId].push(item);
  }

  return grouped;
}

/**
 * Applies a single negative keyword to an ad group.
 * @param {string} campaignId - The campaign ID
 * @param {string} adGroupId - The ad group ID
 * @param {string} searchTerm - The search term to add as negative
 * @returns {Object} Result object with success boolean and optional error message
 */
function applyNegativeKeyword(campaignId, adGroupId, searchTerm) {
  try {
    const adGroupIterator = AdsApp.adGroups()
      .withIds([adGroupId])
      .withCondition(`campaign.id = '${campaignId}'`)
      .get();

    if (!adGroupIterator.hasNext()) {
      // Try shopping ad groups
      const shoppingAdGroupIterator = AdsApp.shoppingAdGroups()
        .withIds([adGroupId])
        .withCondition(`campaign.id = '${campaignId}'`)
        .get();

      if (!shoppingAdGroupIterator.hasNext()) {
        return { success: false, error: `Ad group ${adGroupId} not found` };
      }

      const shoppingAdGroup = shoppingAdGroupIterator.next();
      const formattedKeyword = formatNegativeKeyword(searchTerm, NEGATIVE_MATCH_TYPE);
      shoppingAdGroup.createNegativeKeyword(formattedKeyword);

      console.log(`Added negative "${formattedKeyword}" to shopping ad group ${adGroupId}`);
      return { success: true };
    }

    const adGroup = adGroupIterator.next();
    const formattedKeyword = formatNegativeKeyword(searchTerm, NEGATIVE_MATCH_TYPE);
    adGroup.createNegativeKeyword(formattedKeyword);

    console.log(`Added negative "${formattedKeyword}" to ad group ${adGroupId}`);
    return { success: true };

  } catch (error) {
    console.error(`Error adding negative "${searchTerm}" to ad group ${adGroupId}: ${error.message}`);
    return { success: false, error: error.message };
  }
}

/**
 * Formats a keyword with the specified match type.
 * @param {string} keyword - The keyword text
 * @param {string} matchType - The match type (EXACT, PHRASE, BROAD)
 * @returns {string} Formatted keyword
 */
function formatNegativeKeyword(keyword, matchType) {
  const trimmedKeyword = String(keyword).trim();

  if (!trimmedKeyword) {
    return null;
  }

  const matchTypeLower = matchType.toLowerCase();

  if (matchTypeLower === 'exact') {
    return `[${trimmedKeyword}]`;
  }

  if (matchTypeLower === 'phrase') {
    return `"${trimmedKeyword}"`;
  }

  // Broad match - no modification
  return trimmedKeyword;
}

// ============================================================================
// SHEET OPERATIONS
// ============================================================================

/**
 * Gets or creates a sheet by name.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet
 * @param {string} sheetName - The sheet name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    console.log(`Creating sheet: ${sheetName}`);
    sheet = spreadsheet.insertSheet(sheetName);
  }

  return sheet;
}

/**
 * Clears the output sheet data (preserves headers).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to clear
 */
function clearOutputSheetData(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }

  console.log('Output sheet cleared');
}

/**
 * Writes a "no results" message to the output sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The output sheet
 */
function writeNoResultsMessage(sheet) {
  sheet.clear();
  sheet.getRange('A1').setValue('No potential negatives found matching the criteria.');
}

/**
 * Writes potential negatives to the output sheet with checkboxes.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The output sheet
 * @param {Array<Object>} data - Array of potential negative objects
 */
function writeToOutputSheet(sheet, data) {
  console.log(`Writing ${data.length} rows to output sheet`);

  // Clear existing content
  sheet.clear();

  // Sort by clicks (highest first)
  const sortedData = data.sort((a, b) => b.clicks - a.clicks);

  // Define headers
  const headers = [
    'Negate', 'Search Term', 'Keywords', 'Match Score', 'Campaign Name', 'Ad Group Name',
    'Campaign ID', 'Ad Group ID', 'Impressions', 'Clicks',
    'Cost', 'Conversions', 'Conv. Value', 'CTR', 'Conv. Rate',
    'CPA', 'ROAS', 'Date Found'
  ];

  // Prepare data rows
  const rows = sortedData.map(item => [
    false, // Checkbox (unchecked) - when checked, adds as exact match negative
    item.searchTerm,
    item.keywords,
    item.bestMatchScore / 100, // Convert to decimal for percentage format
    item.campaignName,
    item.adGroupName,
    item.campaignId,
    item.adGroupId,
    item.impressions,
    item.clicks,
    item.cost,
    item.conversions,
    item.conversionsValue,
    item.ctr,
    item.conversionRate,
    item.cpa === Infinity ? 'N/A' : item.cpa,
    item.roas,
    item.dateFound
  ]);

  // Write headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e6f3ff');

  // Add note to "Negate" header explaining the behavior
  sheet.getRange(1, 1).setNote(
    'Check the box to add this search term as an EXACT MATCH negative keyword to its parent ad group.\n\n' +
    'On the next script run, all checked items will be negated.'
  );

  // Write data
  if (rows.length > 0) {
    const dataRange = sheet.getRange(2, 1, rows.length, headers.length);
    dataRange.setValues(rows);

    // Add checkboxes to column A
    const checkboxRange = sheet.getRange(2, 1, rows.length, 1);
    checkboxRange.insertCheckboxes();

    // Format numeric columns
    formatOutputSheetColumns(sheet, rows.length);
  }

  // Freeze header row
  sheet.setFrozenRows(1);

  // Auto-resize columns
  for (let col = 1; col <= headers.length; col++) {
    sheet.autoResizeColumn(col);
  }

  console.log('Output sheet updated successfully');
}

/**
 * Formats numeric columns in the output sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The output sheet
 * @param {number} numRows - Number of data rows
 */
function formatOutputSheetColumns(sheet, numRows) {
  // Match Score column (D = 4) - percentage
  sheet.getRange(2, 4, numRows, 1).setNumberFormat('0.00%');

  // Impressions column (I = 9) - integer
  sheet.getRange(2, 9, numRows, 1).setNumberFormat('#,##0');

  // Clicks column (J = 10) - integer
  sheet.getRange(2, 10, numRows, 1).setNumberFormat('#,##0');

  // Cost column (K = 11) - decimal, no currency
  sheet.getRange(2, 11, numRows, 1).setNumberFormat('#,##0.00');

  // Conversions column (L = 12) - decimal
  sheet.getRange(2, 12, numRows, 1).setNumberFormat('#,##0.00');

  // Conv. Value column (M = 13) - decimal
  sheet.getRange(2, 13, numRows, 1).setNumberFormat('#,##0.00');

  // CTR column (N = 14) - percentage
  sheet.getRange(2, 14, numRows, 1).setNumberFormat('0.00%');

  // Conv. Rate column (O = 15) - percentage
  sheet.getRange(2, 15, numRows, 1).setNumberFormat('0.00%');

  // CPA column (P = 16) - decimal, no currency
  sheet.getRange(2, 16, numRows, 1).setNumberFormat('#,##0.00');

  // ROAS column (Q = 17) - decimal
  sheet.getRange(2, 17, numRows, 1).setNumberFormat('#,##0.00');
}

/**
 * Checks all checkboxes in the output sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The output sheet
 * @param {number} numRows - Number of data rows
 */
function checkAllBoxes(sheet, numRows) {
  if (numRows > 0) {
    const checkboxRange = sheet.getRange(2, 1, numRows, 1);
    checkboxRange.setValue(true);
    console.log(`Checked all ${numRows} checkboxes`);
  }
}

/**
 * Ensures the logs sheet has headers.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The logs sheet
 */
function ensureLogsSheetHeaders(sheet) {
  const headers = [
    'Timestamp', 'Search Term', 'Campaign Name', 'Ad Group Name',
    'Campaign ID', 'Ad Group ID', 'Execution Mode', 'Status'
  ];

  const firstCell = sheet.getRange('A1').getValue();

  if (firstCell !== 'Timestamp') {
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f0f0f0');
    sheet.setFrozenRows(1);
  }
}

/**
 * Logs an applied negative to the logs sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The logs sheet
 * @param {Object} data - Log data object
 */
function logAppliedNegative(sheet, data) {
  const row = [
    data.timestamp,
    data.searchTerm,
    data.campaignName,
    data.adGroupName,
    data.campaignId,
    data.adGroupId,
    data.executionMode,
    data.status
  ];

  sheet.appendRow(row);
}

// ============================================================================
// DATE UTILITIES
// ============================================================================

/**
 * Gets a formatted date string for GAQL queries.
 * @param {number} daysAgo - Number of days ago (0 = today)
 * @returns {string} Date in YYYY-MM-DD format
 */
function getFormattedDate(daysAgo) {
  const date = new Date();
  date.setDate(date.getDate() - daysAgo);

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}

/**
 * Gets the GAQL date range condition.
 * @param {number} lookbackDays - Number of days to look back
 * @returns {string} GAQL date condition
 */
function getDateRangeCondition(lookbackDays) {
  const endDate = getFormattedDate(0);
  const startDate = getFormattedDate(lookbackDays);
  return `segments.date BETWEEN '${startDate}' AND '${endDate}'`;
}

// ============================================================================
// EMAIL NOTIFICATIONS
// ============================================================================

/**
 * Sends email notification for found potential negatives.
 * @param {Array<Object>} potentialNegatives - Array of potential negatives
 * @param {string} spreadsheetUrl - URL to the spreadsheet
 */
function sendFoundEmail(potentialNegatives, spreadsheetUrl) {
  const recipients = EMAIL_RECIPIENTS.split(',').map(e => e.trim()).filter(e => e);

  if (recipients.length === 0) return;

  const accountName = AdsApp.currentAccount().getName();
  const accountId = AdsApp.currentAccount().getCustomerId();
  const timestamp = new Date().toLocaleString();

  const subject = `🔍 Keyword-Based Negatives: ${potentialNegatives.length} Potential Negatives Found - ${accountName}`;

  const previewRows = potentialNegatives.slice(0, 10);

  const htmlBody = generateFoundEmailHtml({
    accountName: accountName,
    accountId: accountId,
    timestamp: timestamp,
    totalFound: potentialNegatives.length,
    previewRows: previewRows,
    spreadsheetUrl: spreadsheetUrl
  });

  try {
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      htmlBody: htmlBody
    });
    console.log(`Found email sent to ${recipients.join(', ')}`);
  } catch (error) {
    console.error(`Error sending found email: ${error.message}`);
  }
}

/**
 * Sends email notification for applied negatives.
 * @param {Array<Object>} appliedNegatives - Array of applied negatives
 * @param {string} spreadsheetUrl - URL to the spreadsheet
 */
function sendAddedEmail(appliedNegatives, spreadsheetUrl) {
  const recipients = EMAIL_RECIPIENTS.split(',').map(e => e.trim()).filter(e => e);

  if (recipients.length === 0) return;

  const accountName = AdsApp.currentAccount().getName();
  const accountId = AdsApp.currentAccount().getCustomerId();
  const timestamp = new Date().toLocaleString();

  const subject = `✅ Keyword-Based Negatives: ${appliedNegatives.length} Negatives Applied - ${accountName}`;

  const previewRows = appliedNegatives.slice(0, 10);

  const htmlBody = generateAddedEmailHtml({
    accountName: accountName,
    accountId: accountId,
    timestamp: timestamp,
    totalApplied: appliedNegatives.length,
    previewRows: previewRows,
    spreadsheetUrl: spreadsheetUrl
  });

  try {
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      htmlBody: htmlBody
    });
    console.log(`Added email sent to ${recipients.join(', ')}`);
  } catch (error) {
    console.error(`Error sending added email: ${error.message}`);
  }
}

/**
 * Generates HTML for the "Found" email.
 * @param {Object} data - Email data object
 * @returns {string} HTML email body
 */
function generateFoundEmailHtml(data) {
  const previewTableRows = data.previewRows.map(row => `
    <tr>
      <td style="padding: 8px 12px; border-bottom: 1px solid #eee;">${row.searchTerm}</td>
      <td style="padding: 8px 12px; border-bottom: 1px solid #eee;">${row.campaignName}</td>
      <td style="padding: 8px 12px; border-bottom: 1px solid #eee;">${row.adGroupName}</td>
      <td style="padding: 8px 12px; border-bottom: 1px solid #eee; text-align: right;">${row.impressions}</td>
      <td style="padding: 8px 12px; border-bottom: 1px solid #eee; text-align: right;">${row.clicks}</td>
      <td style="padding: 8px 12px; border-bottom: 1px solid #eee; text-align: right;">${row.bestMatchScore.toFixed(1)}%</td>
    </tr>
  `).join('');

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
    </head>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 0 auto;">
      <div style="background: linear-gradient(135deg, #667eea, #764ba2); color: white; padding: 30px; text-align: center; border-radius: 8px 8px 0 0;">
        <h1 style="margin: 0; font-size: 24px;">🔍 Potential Negatives Found</h1>
        <p style="margin: 10px 0 0 0; opacity: 0.9;">${data.accountName} (${data.accountId})</p>
      </div>
      
      <div style="padding: 30px; background: #f9f9f9; border: 1px solid #ddd; border-top: none; border-radius: 0 0 8px 8px;">
        <div style="background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; text-align: center;">
          <div style="font-size: 48px; font-weight: bold; color: #667eea;">${data.totalFound}</div>
          <div style="color: #666; font-size: 14px;">Potential Negatives Found</div>
        </div>
        
        <h2 style="color: #333; font-size: 18px; margin-bottom: 15px;">Preview (First 10)</h2>
        
        <table style="width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden;">
          <thead>
            <tr style="background: #667eea; color: white;">
              <th style="padding: 12px; text-align: left;">Search Term</th>
              <th style="padding: 12px; text-align: left;">Campaign</th>
              <th style="padding: 12px; text-align: left;">Ad Group</th>
              <th style="padding: 12px; text-align: right;">Imps</th>
              <th style="padding: 12px; text-align: right;">Clicks</th>
              <th style="padding: 12px; text-align: right;">Match %</th>
            </tr>
          </thead>
          <tbody>
            ${previewTableRows}
          </tbody>
        </table>
        
        ${data.totalFound > 10 ? `<p style="color: #666; font-size: 14px; margin-top: 15px;">... and ${data.totalFound - 10} more</p>` : ''}
        
        <div style="margin-top: 30px; text-align: center;">
          <a href="${data.spreadsheetUrl}" style="display: inline-block; background: #667eea; color: white; padding: 12px 30px; text-decoration: none; border-radius: 5px; font-weight: bold;">View Full Report</a>
        </div>
        
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; color: #999; font-size: 12px; text-align: center;">
          Generated by Keyword-Based Negatives Script • ${data.timestamp}
        </div>
      </div>
    </body>
    </html>
  `;
}

/**
 * Generates HTML for the "Added" email.
 * @param {Object} data - Email data object
 * @returns {string} HTML email body
 */
function generateAddedEmailHtml(data) {
  const previewTableRows = data.previewRows.map(row => `
    <tr>
      <td style="padding: 8px 12px; border-bottom: 1px solid #eee;">${row.searchTerm}</td>
      <td style="padding: 8px 12px; border-bottom: 1px solid #eee;">${row.campaignName}</td>
      <td style="padding: 8px 12px; border-bottom: 1px solid #eee;">${row.adGroupName}</td>
    </tr>
  `).join('');

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
    </head>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 0 auto;">
      <div style="background: linear-gradient(135deg, #34a853, #0f9d58); color: white; padding: 30px; text-align: center; border-radius: 8px 8px 0 0;">
        <h1 style="margin: 0; font-size: 24px;">✅ Negatives Applied</h1>
        <p style="margin: 10px 0 0 0; opacity: 0.9;">${data.accountName} (${data.accountId})</p>
      </div>
      
      <div style="padding: 30px; background: #f9f9f9; border: 1px solid #ddd; border-top: none; border-radius: 0 0 8px 8px;">
        <div style="background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; text-align: center;">
          <div style="font-size: 48px; font-weight: bold; color: #34a853;">${data.totalApplied}</div>
          <div style="color: #666; font-size: 14px;">Negative Keywords Applied</div>
        </div>
        
        <h2 style="color: #333; font-size: 18px; margin-bottom: 15px;">Applied Negatives (First 10)</h2>
        
        <table style="width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden;">
          <thead>
            <tr style="background: #34a853; color: white;">
              <th style="padding: 12px; text-align: left;">Search Term</th>
              <th style="padding: 12px; text-align: left;">Campaign</th>
              <th style="padding: 12px; text-align: left;">Ad Group</th>
            </tr>
          </thead>
          <tbody>
            ${previewTableRows}
          </tbody>
        </table>
        
        ${data.totalApplied > 10 ? `<p style="color: #666; font-size: 14px; margin-top: 15px;">... and ${data.totalApplied - 10} more</p>` : ''}
        
        <div style="margin-top: 30px; text-align: center;">
          <a href="${data.spreadsheetUrl}" style="display: inline-block; background: #34a853; color: white; padding: 12px 30px; text-decoration: none; border-radius: 5px; font-weight: bold;">View Logs</a>
        </div>
        
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; color: #999; font-size: 12px; text-align: center;">
          Generated by Keyword-Based Negatives Script • ${data.timestamp}
        </div>
      </div>
    </body>
    </html>
  `;
}

// ============================================================================
// DEBUG LOGGING
// ============================================================================

/**
 * Logs a debug message if DEBUG_MODE is enabled.
 * @param {string} message - The message to log
 */
function debugLog(message) {
  if (DEBUG_MODE) {
    console.log(`[DEBUG] ${message}`);
  }
}

// ============================================================================
// SEARCH TERM MATCHER CLASS
// ============================================================================

/**
 * Class for matching search terms against keywords using various matching methods.
 */
class SearchTermMatcher {
  /**
   * Creates a new SearchTermMatcher instance.
   * @param {boolean} debug - Enable debug logging
   */
  constructor(debug = false) {
    this.debug = debug;
  }

  /**
   * Checks if a search term matches a condition.
   * @param {string} term - The search term to check
   * @param {Object} condition - Condition object with text, matchType, and optional threshold
   * @returns {boolean} True if the term matches the condition
   */
  matchesCondition(term, condition) {
    if (!condition || !condition.text) {
      if (this.debug) console.warn('[SearchTermMatcher] Invalid condition:', condition);
      return false;
    }

    const termLower = term.toLowerCase();
    const textLower = condition.text.toLowerCase();

    switch (condition.matchType) {
      case 'contains':
        return termLower.includes(textLower);

      case 'not-contains':
        return !termLower.includes(textLower);

      case 'regex-contains':
        try {
          const regex = new RegExp(condition.text, 'i');
          return regex.test(term);
        } catch (error) {
          if (this.debug) console.error('[SearchTermMatcher] Invalid regex:', error);
          return false;
        }

      case 'approx-contains':
        return this.approxMatchContains(term, condition);

      default:
        return false;
    }
  }

  /**
   * Performs approximate (fuzzy) matching.
   * @param {string} term - The search term
   * @param {Object} condition - Condition with text and threshold
   * @returns {boolean} True if fuzzy match score exceeds threshold
   */
  approxMatchContains(term, condition) {
    const threshold = condition.threshold || FUZZY_MATCH_THRESHOLD;

    // Check whole term score
    const wholeTermScore = fuzzyMatchScore(condition.text, term);
    if (wholeTermScore >= threshold) {
      return true;
    }

    // Check without spaces
    const spacelessScore = fuzzyMatchScore(
      condition.text.replace(/\s+/g, ''),
      term.replace(/\s+/g, '')
    );
    if (spacelessScore >= threshold) {
      return true;
    }

    // Check with words sorted (ignore word order)
    const sortedKeyword = condition.text.toLowerCase().split(/\s+/).sort().join(' ');
    const sortedTerm = term.toLowerCase().split(/\s+/).sort().join(' ');
    const sortedScore = fuzzyMatchScore(sortedKeyword, sortedTerm);
    if (sortedScore >= threshold) {
      return true;
    }

    // Check if all search term words exist in keyword (subset match)
    // e.g. "osmosis water filter" should match "reverse osmosis water filter"
    const keywordWords = condition.text.toLowerCase().split(/\s+/);
    const searchTermWords = term.toLowerCase().split(/\s+/);
    const allWordsInKeyword = searchTermWords.every(searchWord =>
      keywordWords.some(keywordWord => fuzzyMatchScore(searchWord, keywordWord) >= threshold)
    );
    if (allWordsInKeyword) {
      return true;
    }

    // Check if any keyword word exists in search term words (shared word match)
    // e.g. "reverse osmosis ro" should match "ro system" because "ro" is shared
    const anySharedWord = keywordWords.some(keywordWord =>
      searchTermWords.some(searchWord => fuzzyMatchScore(keywordWord, searchWord) >= threshold)
    );
    if (anySharedWord) {
      return true;
    }

    // Check individual words
    const termWords = term.toLowerCase().split(' ');
    const anyWordMatches = termWords.some(
      word => fuzzyMatchScore(condition.text, word) >= threshold
    );
    if (anyWordMatches) {
      return true;
    }

    // Sliding window match
    let startPosition = 0;
    let endPosition = condition.text.length;
    while (endPosition <= term.length) {
      const windowScore = fuzzyMatchScore(
        condition.text,
        term.substring(startPosition, endPosition)
      );
      if (windowScore >= threshold) {
        return true;
      }
      startPosition++;
      endPosition++;
    }

    return false;
  }
}

// ============================================================================
// FUZZY SET IMPLEMENTATION
// ============================================================================

/**
 * FuzzySet implementation for fuzzy string matching.
 * Based on the FuzzySet.js library.
 */
var FuzzySet = (function () {
  "use strict";

  const FuzzySet = function (arr, useLevenshtein, gramSizeLower, gramSizeUpper) {
    var fuzzyset = {};

    arr = arr || [];
    fuzzyset.gramSizeLower = gramSizeLower || 2;
    fuzzyset.gramSizeUpper = gramSizeUpper || 3;
    fuzzyset.useLevenshtein = typeof useLevenshtein !== "boolean" ? true : useLevenshtein;

    fuzzyset.exactSet = {};
    fuzzyset.matchDict = {};
    fuzzyset.items = {};

    var levenshtein = function (str1, str2) {
      var current = [], prev, value;

      for (var i = 0; i <= str2.length; i++) {
        for (var j = 0; j <= str1.length; j++) {
          if (i && j) {
            if (str1.charAt(j - 1) === str2.charAt(i - 1)) {
              value = prev;
            } else {
              value = Math.min(current[j], current[j - 1], prev) + 1;
            }
          } else {
            value = i + j;
          }
          prev = current[j];
          current[j] = value;
        }
      }
      return current.pop();
    };

    var _distance = function (str1, str2) {
      if (str1 === null && str2 === null) throw "Trying to compare two null values";
      if (str1 === null || str2 === null) return 0;
      str1 = String(str1);
      str2 = String(str2);

      var distance = levenshtein(str1, str2);
      if (str1.length > str2.length) {
        return 1 - distance / str1.length;
      } else {
        return 1 - distance / str2.length;
      }
    };

    var _nonWordRe = /[^a-zA-Z0-9\u00C0-\u00FF\u0621-\u064A\u0660-\u0669, ]+/g;

    var _iterateGrams = function (value, gramSize) {
      gramSize = gramSize || 2;
      var simplified = "-" + value.toLowerCase().replace(_nonWordRe, "") + "-",
        lenDiff = gramSize - simplified.length,
        results = [];
      if (lenDiff > 0) {
        for (var i = 0; i < lenDiff; ++i) {
          simplified += "-";
        }
      }
      for (var i = 0; i < simplified.length - gramSize + 1; ++i) {
        results.push(simplified.slice(i, i + gramSize));
      }
      return results;
    };

    var _gramCounter = function (value, gramSize) {
      gramSize = gramSize || 2;
      var result = {},
        grams = _iterateGrams(value, gramSize),
        i = 0;
      for (i; i < grams.length; ++i) {
        if (grams[i] in result) {
          result[grams[i]] += 1;
        } else {
          result[grams[i]] = 1;
        }
      }
      return result;
    };

    fuzzyset.get = function (value, defaultValue, minMatchScore) {
      if (minMatchScore === undefined) {
        minMatchScore = 0.33;
      }
      var result = this._get(value, minMatchScore);
      if (!result && typeof defaultValue !== "undefined") {
        return defaultValue;
      }
      return result;
    };

    fuzzyset._get = function (value, minMatchScore) {
      var results = [];
      for (var gramSize = this.gramSizeUpper; gramSize >= this.gramSizeLower; --gramSize) {
        results = this.__get(value, gramSize, minMatchScore);
        if (results && results.length > 0) {
          return results;
        }
      }
      return null;
    };

    fuzzyset.__get = function (value, gramSize, minMatchScore) {
      var normalizedValue = this._normalizeStr(value),
        matches = {},
        gramCounts = _gramCounter(normalizedValue, gramSize),
        items = this.items[gramSize],
        sumOfSquareGramCounts = 0,
        gram, gramCount, i, index, otherGramCount;

      for (gram in gramCounts) {
        gramCount = gramCounts[gram];
        sumOfSquareGramCounts += Math.pow(gramCount, 2);
        if (gram in this.matchDict) {
          for (i = 0; i < this.matchDict[gram].length; ++i) {
            index = this.matchDict[gram][i][0];
            otherGramCount = this.matchDict[gram][i][1];
            if (index in matches) {
              matches[index] += gramCount * otherGramCount;
            } else {
              matches[index] = gramCount * otherGramCount;
            }
          }
        }
      }

      function isEmptyObject(obj) {
        for (var prop in obj) {
          if (obj.hasOwnProperty(prop)) return false;
        }
        return true;
      }

      if (isEmptyObject(matches)) {
        return null;
      }

      var vectorNormal = Math.sqrt(sumOfSquareGramCounts),
        results = [],
        matchScore;
      for (var matchIndex in matches) {
        matchScore = matches[matchIndex];
        results.push([
          matchScore / (vectorNormal * items[matchIndex][0]),
          items[matchIndex][1]
        ]);
      }
      var sortDescending = function (a, b) {
        if (a[0] < b[0]) return 1;
        if (a[0] > b[0]) return -1;
        return 0;
      };
      results.sort(sortDescending);
      if (this.useLevenshtein) {
        var newResults = [],
          endIndex = Math.min(50, results.length);
        for (var i = 0; i < endIndex; ++i) {
          newResults.push([_distance(results[i][1], normalizedValue), results[i][1]]);
        }
        results = newResults;
        results.sort(sortDescending);
      }
      newResults = [];
      results.forEach(function (scoreWordPair) {
        if (scoreWordPair[0] >= minMatchScore) {
          newResults.push([scoreWordPair[0], this.exactSet[scoreWordPair[1]]]);
        }
      }.bind(this));
      return newResults;
    };

    fuzzyset.add = function (value) {
      var normalizedValue = this._normalizeStr(value);
      if (normalizedValue in this.exactSet) {
        return false;
      }

      var i = this.gramSizeLower;
      for (i; i < this.gramSizeUpper + 1; ++i) {
        this._add(value, i);
      }
    };

    fuzzyset._add = function (value, gramSize) {
      var normalizedValue = this._normalizeStr(value),
        items = this.items[gramSize] || [],
        index = items.length;

      items.push(0);
      var gramCounts = _gramCounter(normalizedValue, gramSize),
        sumOfSquareGramCounts = 0,
        gram, gramCount;
      for (gram in gramCounts) {
        gramCount = gramCounts[gram];
        sumOfSquareGramCounts += Math.pow(gramCount, 2);
        if (gram in this.matchDict) {
          this.matchDict[gram].push([index, gramCount]);
        } else {
          this.matchDict[gram] = [[index, gramCount]];
        }
      }
      var vectorNormal = Math.sqrt(sumOfSquareGramCounts);
      items[index] = [vectorNormal, normalizedValue];
      this.items[gramSize] = items;
      this.exactSet[normalizedValue] = value;
    };

    fuzzyset._normalizeStr = function (str) {
      if (Object.prototype.toString.call(str) !== "[object String]") {
        throw "Must use a string as argument to FuzzySet functions";
      }
      return str.toLowerCase();
    };

    fuzzyset.length = function () {
      var count = 0, prop;
      for (prop in this.exactSet) {
        if (this.exactSet.hasOwnProperty(prop)) {
          count += 1;
        }
      }
      return count;
    };

    fuzzyset.isEmpty = function () {
      for (var prop in this.exactSet) {
        if (this.exactSet.hasOwnProperty(prop)) {
          return false;
        }
      }
      return true;
    };

    fuzzyset.values = function () {
      var values = [], prop;
      for (prop in this.exactSet) {
        if (this.exactSet.hasOwnProperty(prop)) {
          values.push(this.exactSet[prop]);
        }
      }
      return values;
    };

    var i = fuzzyset.gramSizeLower;
    for (i; i < fuzzyset.gramSizeUpper + 1; ++i) {
      fuzzyset.items[i] = [];
    }
    for (i = 0; i < arr.length; ++i) {
      fuzzyset.add(arr[i]);
    }

    return fuzzyset;
  };

  return FuzzySet;
})();

/**
 * Calculates the fuzzy match score between two strings.
 * @param {string} needle - The string to search for
 * @param {string} haystack - The string to search in
 * @returns {number} Match score from 0 to 100
 */
function fuzzyMatchScore(needle, haystack) {
  let fuzzySet = FuzzySet();
  fuzzySet.add(haystack);
  let result = fuzzySet.get(needle);
  if (!result) return 0;
  return result[0][0] * 100;
}
