/**
 * Placement Exclusions - Semi-automated (PMax, Display, YouTube)
 * @author Charles Bannister (https://www.linkedin.com/in/charles-bannister/)
 * @author Erik Arter (https://www.linkedin.com/in/erikarter/)
 * This script was Erik's idea via the God Tier Ads community.
 * You can find the discussion here: https://go.godtierads.com/c/questions-answers/website-placement-exclusions
 * 
 * More scripts at:
 * https://shabba.io
 * https://github.com/charlesbannister/free-google-ads-scripts/
 * 
 * This script grabs placement data from Performance Max, Display, and Video/YouTube campaigns,
 * writes to sheet with checkboxes for user selection, then excludes the selected placements.
 * YouTube placements (YOUTUBE_CHANNEL, YOUTUBE_VIDEO) are written to the YouTube output tab.
 * Includes optional ChatGPT integration for website content analysis.
 * Version: 2.9.2
 */

// Google Ads API Query Builder Links:
// Performance Max Placement View: https://developers.google.com/google-ads/api/fields/v20/performance_max_placement_view_query_builder
// Display/Video Placement View: https://developers.google.com/google-ads/api/fields/v20/detail_placement_view_query_builder

//Installation Instructions:
// 1. Open the script editor by going to Bulk actions > Scripts
// 2. Click on the big plus button and name the script "PMax Placement Exclusions"
// 3. Create your Google Sheet by typing "sheets.new" in the URL bar and name the sheet "PMax Placement Exclusions"
// 4. Paste the URL of your new sheet into the SPREADSHEET_URL variable below
// 5. Paste the entire script into Google Ads (you can delete what's there)
// 6. Preview the script! (You'll be prompted to authorise the first time it runs)

// --- Configuration ---

// Template: https://docs.google.com/spreadsheets/d/1jG_igH1QGdyBSbeqj2ELxZEg3uOFYd3wcWaQ_eDfY9o
// File > Make a copy or visit https://docs.google.com/spreadsheets/d/1jG_igH1QGdyBSbeqj2ELxZEg3uOFYd3wcWaQ_eDfY9o/copy
const SPREADSHEET_URL = 'YOUR_SPREADSHEET_URL_HERE';


const DEBUG_MODE = false;
// Set to true to see detailed logs for debugging
// Core logs will always appear regardless of this setting

const SETTINGS_SHEET_NAME = 'Settings';
// The name of the sheet tab that contains configuration settings


const LLM_RESPONSES_SHEET_NAME = 'LLM Responses Cache';
// The name of the sheet tab that caches ChatGPT responses by URL

const WEBSITE_OUTPUT_SHEET_NAME = 'Output: Website';
// The name of the sheet tab that contains the output of the website placement analysis

const YOUTUBE_OUTPUT_SHEET_NAME = 'Output: YouTube';
// The name of the sheet tab that contains the output of the YouTube placement analysis

const MOBILE_APPLICATION_OUTPUT_SHEET_NAME = 'Output: Mobile Application';
// The name of the sheet tab that contains the output of the mobile application placement analysis

const GOOGLE_PRODUCTS_OUTPUT_SHEET_NAME = 'Output: Google Products';
// The name of the sheet tab that contains the output of the Google products placement analysis

const LISTS_SHEET_NAME = 'Settings: Placement Filters';
// The name of the sheet tab that contains placement filter lists

const CHATGPT_SHEET_NAME = 'Settings: ChatGPT';
// The name of the sheet tab that contains ChatGPT configuration settings

const EXCLUSION_LOG_SHEET_NAME = 'Exclusion Log';
// The name of the sheet tab that logs all exclusions made by the script
// Assumes sheet already exists with headers: Status, Campaign Name, Placement, Display Name, Placement Type, Target URL, Notes, Preview Mode, Timestamp

const DISABLE_REPORTING = false;
// If true: Skip fetching new placement data and generating reports. Only process checked placements from existing sheets.
// If false: Normal operation - fetch placement data, generate reports, then process checked placements.
// NOTE: Set to true AFTER you've created the shared exclusion list and want to just add placements without regenerating reports.



/**
 * Cell reference configuration for sheets
 * All row/column numbers are 1-based (as used in Google Sheets)
 */
const SETTINGS_CELL_REFERENCES = {
  // Main settings section
  sharedExclusionListName: { row: 4, column: 2 },
  lookbackWindowDays: { row: 5, column: 2 },
  minimumImpressions: { row: 6, column: 2 },
  minimumClicks: { row: 7, column: 2 },
  minimumCost: { row: 8, column: 2 },
  maximumConversions: { row: 9, column: 2 },
  maxResults: { row: 10, column: 2 },
  campaignNameContains: { row: 11, column: 2 },
  campaignNameNotContains: { row: 12, column: 2 },
  enabledCampaignsOnly: { row: 13, column: 2 }, // checkbox
  activePlacementsOnly: { row: 14, column: 2 }, // checkbox

  // Placement Type Filters section
  placementTypes: {
    youtubeVideo: { row: 18, enabledColumn: 2, automatedColumn: 3 },
    website: { row: 19, enabledColumn: 2, automatedColumn: 3 },
    mobileApplication: { row: 20, enabledColumn: 2, automatedColumn: 3 },
    googleProducts: { row: 21, enabledColumn: 2, automatedColumn: 3 }
  }
};

const CHATGPT_CELL_REFERENCES = {
  enableChatGpt: { row: 5, column: 2 }, // checkbox
  chatGptApiKey: { row: 6, column: 2 },
  useCachedChatGpt: { row: 7, column: 2 }, // checkbox
  chatGptPrompt: { row: 8, column: 2 },
  responseContains: { row: 13, column: 2 },
  responseNotContains: { row: 14, column: 2 }
};

/**
 * Cell reference configuration for the Placement Filters sheet
 * 
 * Sheet Layout:
 * Row 7: Checkboxes
 * Row 8+: Data
 * 
 * Columns (Unwanted A-E, Wanted F-J):
 * A: Unwanted Placement (placementNotContains)
 * B: Unwanted Display Name (displayNameNotContains)
 * C: Unwanted Target URL (targetUrlNotContains)
 * D: Spammy Top Level Domains (targetUrlNotEndsWith)
 * E: Out-of-region Top Level Domains (targetUrlNotEndsWith)
 * F: Wanted Placement (placementContains)
 * G: Wanted Display Name (displayNameContains)
 * H: Wanted Target URL (targetUrlContains)
 * I: Target Top Level Domains (targetUrlEndsWith)
 * J: Target Region Top Level Domains (targetUrlEndsWith)
 */
const PLACEMENT_FILTERS_CELL_REFERENCES = {
  placementNotContains: { row: 7, column: 1 },
  displayNameNotContains: { row: 7, column: 2 },
  targetUrlNotContains: { row: 7, column: 3 },
  spammyTlds: { row: 7, column: 4 },
  outOfRegionTlds: { row: 7, column: 5 },
  placementContains: { row: 7, column: 6 },
  displayNameContains: { row: 7, column: 7 },
  targetUrlContains: { row: 7, column: 8 },
  targetTlds: { row: 7, column: 9 },
  targetRegionTlds: { row: 7, column: 10 }
};

const PLACEMENT_FILTERS_LIST_START_ROWS = {
  placementNotContains: 8,
  displayNameNotContains: 8,
  targetUrlNotContains: 8,
  spammyTlds: 8,
  outOfRegionTlds: 8,
  placementContains: 8,
  displayNameContains: 8,
  targetUrlContains: 8,
  targetTlds: 8,
  targetRegionTlds: 8
};


/**
 * Logs a message only when DEBUG_MODE is enabled
 * @param {string} message - The message to log
 */
function debugLog(message) {
  if (DEBUG_MODE) {
    console.log(message);
  }
}

// --- Filter Helper Functions ---

/**
 * Checks if a string contains any of the filter strings as whole words (case-insensitive)
 * Splits on non-alphanumeric characters to extract words, then compares exactly
 * This prevents false positives like "Essex" matching "sex"
 * @param {string} value - The value to check
 * @param {Array<string>} filterList - List of words to match against
 * @returns {boolean} true if value contains any filter word
 */
function containsAny(value, filterList) {
  if (!filterList || filterList.length === 0) return false;
  const valueWords = String(value || '').toLowerCase().split(/[^a-z0-9]+/).filter(word => word.length > 0);
  return filterList.some(filter => {
    const filterLower = filter.toLowerCase();
    return valueWords.some(word => word === filterLower);
  });
}

/**
 * Checks if a URL ends with any of the TLD strings (case-insensitive)
 * @param {string} url - The URL to check
 * @param {Array<string>} tldList - List of TLDs to match against
 * @returns {boolean} true if URL ends with any TLD
 */
function endsWithAny(url, tldList) {
  if (!tldList || tldList.length === 0) return false;
  const urlLower = String(url || '').toLowerCase();
  return tldList.some(tld => urlLower.endsWith(tld.toLowerCase()));
}

/**
 * Logs all settings values when DEBUG_MODE is enabled
 * @param {Object} settings - The settings object
 */
function logSettingsDebug(settings) {
  if (!DEBUG_MODE) return;

  debugLog(`Settings loaded:`);
  debugLog(`  Lookback Window: ${settings.lookbackWindowDays} days`);
  debugLog(`  Impressions >: ${settings.minimumImpressions}`);
  debugLog(`  Clicks >: ${settings.minimumClicks}`);
  debugLog(`  Cost >: ${settings.minimumCost}`);
  debugLog(`  Conversions <: ${settings.maximumConversions === 0 ? 'No limit' : settings.maximumConversions}`);
  debugLog(`  Max Results: ${settings.maxResults === 0 ? 'No limit' : settings.maxResults}`);
  debugLog(`  Campaign Name Contains: "${settings.campaignNameContains}"`);
  debugLog(`  Campaign Name Not Contains: "${settings.campaignNameNotContains}"`);
  debugLog(`  Enabled campaigns only: ${settings.enabledCampaignsOnly}`);
  debugLog(`  Active placements only: ${settings.activePlacementsOnly}`);
  debugLog(`  Placement Type Filters:`);
  debugLog(`    YouTube Video: enabled=${settings.placementTypeFilters.youtubeVideo.enabled}, automated=${settings.placementTypeFilters.youtubeVideo.automated}`);
  debugLog(`    Website: enabled=${settings.placementTypeFilters.website.enabled}, automated=${settings.placementTypeFilters.website.automated}`);
  debugLog(`    Mobile Application: enabled=${settings.placementTypeFilters.mobileApplication.enabled}, automated=${settings.placementTypeFilters.mobileApplication.automated}`);
  debugLog(`    Google Products: enabled=${settings.placementTypeFilters.googleProducts.enabled}, automated=${settings.placementTypeFilters.googleProducts.automated}`);
  debugLog(`  Enable ChatGPT: ${settings.enableChatGpt}`);
  if (settings.enableChatGpt) {
    debugLog(`  ChatGPT API Key: ${settings.chatGptApiKey ? 'Provided' : 'Missing'}`);
    debugLog(`  Use Cached ChatGPT Responses: ${settings.useCachedChatGpt}`);
  }
}


// --- Main Function ---
function main() {
  console.log(`Script started`);

  if (DISABLE_REPORTING) {
    console.log(`\n⚠️ ========================================== ⚠️`);
    console.log(`⚠️  WARNING: DISABLE_REPORTING IS TRUE     ⚠️`);
    console.log(`⚠️  New placement data will NOT be fetched    ⚠️`);
    console.log(`⚠️  Output sheets will NOT be updated         ⚠️`);
    console.log(`⚠️  Only checked boxes will be processed      ⚠️`);
    console.log(`⚠️ ========================================== ⚠️\n`);
  }

  validateConfig();

  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const settingsSheet = getOrCreateSettingsSheet(spreadsheet);
  const listsSheet = getOrCreateListsSheet(spreadsheet);
  const chatGptSheet = getOrCreateChatGptSheet(spreadsheet);
  const chatGptSettings = getChatGptSettingsFromSheet(chatGptSheet);
  const settings = getSettingsFromSheet(settingsSheet, listsSheet);


  // Merge ChatGPT settings into main settings object
  settings.enableChatGpt = chatGptSettings.enableChatGpt;
  settings.chatGptApiKey = chatGptSettings.chatGptApiKey;
  settings.useCachedChatGpt = chatGptSettings.useCachedChatGpt;
  settings.chatGptPrompt = chatGptSettings.chatGptPrompt;
  settings.responseContainsList = chatGptSettings.responseContainsList;
  settings.responseNotContainsList = chatGptSettings.responseNotContainsList;

  validateSharedExclusionList(settings.sharedExclusionListName);

  // Read checked placements from sheets
  const checkedPlacements = readCheckedPlacementsFromSheet(spreadsheet);

  // If DISABLE_REPORTING is true, skip fetching new data and only process checked placements
  if (DISABLE_REPORTING) {
    console.log(`DISABLE_REPORTING mode enabled - skipping data fetch and sheet updates`);

    const hasWebsitePlacements = checkedPlacements.websitePlacements && checkedPlacements.websitePlacements.length > 0;
    const hasYouTubeVideos = checkedPlacements.youtubeVideos && checkedPlacements.youtubeVideos.length > 0;
    const hasMobileAppPlacements = checkedPlacements.mobileAppPlacements && checkedPlacements.mobileAppPlacements.length > 0;

    if (hasWebsitePlacements || hasYouTubeVideos || hasMobileAppPlacements) {
      // Validate the exclusion list exists before attempting to add exclusions
      requireSharedExclusionList(settings.sharedExclusionListName);

      // Add website placements using the standard method
      if (hasWebsitePlacements) {
        addPlacementsToExclusionList(checkedPlacements.websitePlacements, settings.sharedExclusionListName);
      }

      // Add YouTube placements (channels and videos) individually
      if (hasYouTubeVideos) {
        addYouTubeVideosToExclusionList(checkedPlacements.youtubeVideos, settings.sharedExclusionListName);
      }

      // Add mobile app placements using Bulk Upload
      if (hasMobileAppPlacements) {
        addAppExclusionsViaBulkUpload(settings.sharedExclusionListName, checkedPlacements.mobileAppPlacements);
      }

      // Log exclusions to the Exclusion Log sheet
      logExclusionsToSheet(spreadsheet, checkedPlacements.fullPlacementData);

      linkSharedListToPMaxCampaigns(settings.sharedExclusionListName);
    } else {
      console.log(`No placements selected for exclusion. Check boxes in column A of the placement type sheets to exclude placements, then run the script again.`);
      // Still ensure the list is linked to campaigns even if no new placements added
      linkSharedListToPMaxCampaigns(settings.sharedExclusionListName);
    }
  } else {
    // Normal flow: fetch new data, process, and update sheets
    const placementData = getPlacementData(settings);
    console.log(`Found ${placementData.length} placements`);

    if (placementData.length === 0) {
      console.log(`No placement data found. Check your filters and date range.`);
      // Clear sheets and write "No results found" message
      try {
        writePlacementDataToSheet(spreadsheet, [], settings);
      } catch (writeError) {
        console.error(`Error writing to sheets: ${writeError.message}`);
      }
      console.log(`Script finished`);
      return;
    }

    const transformedData = calculatePlacementMetrics(placementData);

    // Filter data first based on thresholds
    const filteredData = filterPlacementData(transformedData, settings);

    // Max results limit is applied in the GAQL query (LIMIT clause)
    // Note: LIMIT applies per query type, so total results may be up to 2x this value (PMax + Display)
    const limitedData = filteredData;

    // Get ChatGPT responses if enabled (only for website placements that passed filters)
    if (settings.enableChatGpt) {
      if (!settings.chatGptApiKey) {
        console.warn(`⚠ ChatGPT is enabled but API key is missing. Skipping ChatGPT analysis.`);
        for (const placement of limitedData) {
          placement.chatGptResponse = '';
        }
      } else {
        console.log(`ChatGPT enabled. Getting responses for placements...`);

        // Load LLM cache once before processing
        const llmSheet = getOrCreateLlmResponsesSheet(spreadsheet);
        const llmCache = getCachedLlmResponses(llmSheet);

        let processedCount = 0;
        let cachedCount = 0;
        let failedCount = 0;
        let filteredCount = 0;

        for (const placement of limitedData) {
          // Skip ChatGPT for mobile applications and Google Products
          const placementType = String(placement.placementType || '').toUpperCase();
          if (placementType.includes('MOBILE_APPLI') || placementType.includes('GOOGLE_PRODUCTS')) {
            placement.chatGptResponse = '';
            continue;
          }

          const url = placement.targetUrl || placement.placement;
          debugLog(`  Processing ChatGPT for: ${url} (targetUrl: "${placement.targetUrl}", placement: "${placement.placement}")`)

          try {
            const chatGptResult = getChatGptResponseForUrl(url, placement, settings, spreadsheet, llmCache);
            placement.chatGptResponse = chatGptResult.response;

            if (chatGptResult.status === 'success') {
              processedCount++;
              if (chatGptResult.wasCached) {
                cachedCount++;
                const responsePreview = chatGptResult.response ? chatGptResult.response.substring(0, 50) + '...' : '(empty)';
                debugLog(`  ✓ Got ChatGPT response for: ${url} (from cache, length: ${chatGptResult.response.length}, preview: ${responsePreview})`)
              } else {
                debugLog(`  ✓ Got ChatGPT response for: ${url} (new)`)
              }
            } else if (chatGptResult.status === 'failed') {
              failedCount++;
              debugLog(`  ✗ ChatGPT failed for: ${url} (status: ${chatGptResult.status})`)
            }
          } catch (chatGptError) {
            console.error(`  ✗ ChatGPT error for ${url}: ${chatGptError.message}`);
            placement.chatGptResponse = '';
            failedCount++;
          }
        }

        // Filter out placements based on Response Contains filter (excludes entire row from output)
        if (settings.responseContainsList && settings.responseContainsList.length > 0) {
          const beforeFilterCount = limitedData.length;
          const filteredLimitedData = limitedData.filter(placement => {
            // Only apply filter to placements that have a ChatGPT response
            // (mobile apps and Google Products have empty responses and should pass)
            if (!placement.chatGptResponse) {
              return true; // No response = pass (don't filter out rows without responses)
            }
            return responsePassesFilter(placement.chatGptResponse, settings);
          });
          filteredCount = beforeFilterCount - filteredLimitedData.length;

          // Replace limitedData with filtered version
          limitedData.length = 0;
          limitedData.push(...filteredLimitedData);
        }

        console.log(`ChatGPT processing complete: ${processedCount} responses (${cachedCount} from cache), ${filteredCount} filtered, ${failedCount} failed`);
      }
    } else {
      // Add empty chatGptResponse field if ChatGPT is disabled
      for (const placement of limitedData) {
        placement.chatGptResponse = '';
      }
    }

    writePlacementDataToSheet(spreadsheet, limitedData, settings);

    const hasWebsitePlacements = checkedPlacements.websitePlacements && checkedPlacements.websitePlacements.length > 0;
    const hasYouTubeVideos = checkedPlacements.youtubeVideos && checkedPlacements.youtubeVideos.length > 0;
    const hasMobileAppPlacements = checkedPlacements.mobileAppPlacements && checkedPlacements.mobileAppPlacements.length > 0;

    if (hasWebsitePlacements || hasYouTubeVideos || hasMobileAppPlacements) {
      // Validate the exclusion list exists before attempting to add exclusions
      requireSharedExclusionList(settings.sharedExclusionListName);

      // Add website placements using the standard method
      if (hasWebsitePlacements) {
        addPlacementsToExclusionList(checkedPlacements.websitePlacements, settings.sharedExclusionListName);
      }

      // Add YouTube placements (channels and videos) individually
      if (hasYouTubeVideos) {
        addYouTubeVideosToExclusionList(checkedPlacements.youtubeVideos, settings.sharedExclusionListName);
      }

      // Add mobile app placements using Bulk Upload
      if (hasMobileAppPlacements) {
        addAppExclusionsViaBulkUpload(settings.sharedExclusionListName, checkedPlacements.mobileAppPlacements);
      }

      // Log exclusions to the Exclusion Log sheet
      logExclusionsToSheet(spreadsheet, checkedPlacements.fullPlacementData);

      linkSharedListToPMaxCampaigns(settings.sharedExclusionListName);
    } else {
      console.log(`No placements selected for exclusion. Check boxes in column A of the placement type sheets to exclude placements, then run the script again.`);
      // Still ensure the list is linked to campaigns even if no new placements added
      linkSharedListToPMaxCampaigns(settings.sharedExclusionListName);
    }
  }

  if (DISABLE_REPORTING) {
    console.log(`\n⚠️ ========================================== ⚠️`);
    console.log(`⚠️  REMINDER: DISABLE_REPORTING WAS TRUE   ⚠️`);
    console.log(`⚠️  No new data was fetched or written        ⚠️`);
    console.log(`⚠️  Set to FALSE to refresh placement data    ⚠️`);
    console.log(`⚠️ ========================================== ⚠️\n`);
  }

  console.log(`Script finished`);
}

// --- Validation Functions ---

/**
 * Validates that the spreadsheet URL is configured
 */
function validateConfig() {
  if (!SPREADSHEET_URL || SPREADSHEET_URL.includes('PASTE_YOUR_SPREADSHEET_URL_HERE')) {
    throw new Error('Please update SPREADSHEET_URL with your actual spreadsheet URL');
  }
}

/**
 * Validates that the shared placement exclusion list exists
 * @param {string} listName - The name of the shared exclusion list
 * @returns {boolean} True if the list exists, false otherwise
 */
function validateSharedExclusionList(listName) {
  const listIterator = AdsApp.excludedPlacementLists()
    .withCondition(`Name = '${listName}'`)
    .get();

  if (!listIterator.hasNext()) {
    console.error(`⚠ Shared Placement Exclusion List named '${listName}' not found.`);
    console.error(`  The script will continue to generate reports, but cannot add exclusions.`);
    console.error(`  To enable exclusions, create the list in Google Ads:`);
    console.error(`  1. Go to Tools & Settings > Shared Library > Placement exclusions`);
    console.error(`  2. Create a new list with the exact name: "${listName}"`);
    return false;
  }

  console.log(`✓ Found shared exclusion list: ${listName}`);
  return true;
}

/**
 * Throws an error if the shared exclusion list doesn't exist
 * Call this only when attempting to make changes that require the list
 * @param {string} listName - The name of the shared exclusion list
 */
function requireSharedExclusionList(listName) {
  const listIterator = AdsApp.excludedPlacementLists()
    .withCondition(`Name = '${listName}'`)
    .get();

  if (!listIterator.hasNext()) {
    const errorMessage = `ERROR: Cannot add exclusions - Shared Placement Exclusion List named '${listName}' not found.\n\n` +
      `Please create this shared list in your Google Ads account first:\n` +
      `1. Go to Tools & Settings > Shared Library > Placement exclusions\n` +
      `2. Create a new list with the exact name: "${listName}"\n` +
      `3. Ensure the list is enabled\n` +
      `4. Link this list to your Performance Max campaigns if not already linked`;
    throw new Error(errorMessage);
  }
}

// --- Settings Management ---

/**
 * Gets or creates the settings sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The settings sheet
 */
function getOrCreateSettingsSheet(spreadsheet) {
  let settingsSheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    settingsSheet = spreadsheet.insertSheet(SETTINGS_SHEET_NAME);
    console.log(`Created settings sheet: ${SETTINGS_SHEET_NAME}`);
  }
  return settingsSheet;
}

/**
 * Gets or creates the Lists sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The Lists sheet
 */
function getOrCreateListsSheet(spreadsheet) {
  let listsSheet = spreadsheet.getSheetByName(LISTS_SHEET_NAME);
  if (!listsSheet) {
    listsSheet = spreadsheet.insertSheet(LISTS_SHEET_NAME);
    console.log(`Created lists sheet: ${LISTS_SHEET_NAME}`);
  }
  return listsSheet;
}

/**
 * Gets or creates the ChatGPT sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The ChatGPT sheet
 */
function getOrCreateChatGptSheet(spreadsheet) {
  let chatGptSheet = spreadsheet.getSheetByName(CHATGPT_SHEET_NAME);
  if (!chatGptSheet) {
    chatGptSheet = spreadsheet.insertSheet(CHATGPT_SHEET_NAME);
  }

  return chatGptSheet;
}

/**
 * Gets settings from the settings sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet - The settings sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The lists sheet
 * @returns {Object} Settings object with all configuration values
 */
function getSettingsFromSheet(settingsSheet, listsSheet) {
  const refs = SETTINGS_CELL_REFERENCES;

  // Helper function to read a cell value
  const getCellValue = (cellRef) => {
    debugLog(`cellRef: ${JSON.stringify(cellRef)}`);
    const row = cellRef.row;
    const col = cellRef.column;
    const value = settingsSheet.getRange(row, col).getValue();
    return value;
  };

  // Read all settings from sheet using cell references
  const sharedExclusionListNameRaw = getCellValue(refs.sharedExclusionListName);
  const sharedExclusionListName = sharedExclusionListNameRaw !== null && sharedExclusionListNameRaw !== undefined
    ? String(sharedExclusionListNameRaw).trim()
    : '';
  const lookbackWindowDaysRaw = getCellValue(refs.lookbackWindowDays);
  const minimumImpressionsRaw = getCellValue(refs.minimumImpressions);
  const minimumClicksRaw = getCellValue(refs.minimumClicks);
  const minimumCostRaw = getCellValue(refs.minimumCost);
  const maximumConversionsRaw = getCellValue(refs.maximumConversions);
  const maxResultsRaw = getCellValue(refs.maxResults);
  const enabledCampaignsOnlyRaw = getCellValue(refs.enabledCampaignsOnly);
  const activePlacementsOnlyRaw = getCellValue(refs.activePlacementsOnly);

  // Handle checkbox value - Google Sheets checkboxes return boolean true/false
  // Use what's in the sheet, default to false only if setting row not found (malformed sheet)
  let enabledCampaignsOnly = false;
  if (enabledCampaignsOnlyRaw === null || enabledCampaignsOnlyRaw === undefined) {
    // Setting row not found (malformed sheet) - use default
    enabledCampaignsOnly = false;
  } else if (typeof enabledCampaignsOnlyRaw === 'boolean') {
    enabledCampaignsOnly = enabledCampaignsOnlyRaw;
  } else if (typeof enabledCampaignsOnlyRaw === 'string') {
    const lowerValue = enabledCampaignsOnlyRaw.toLowerCase().trim();
    enabledCampaignsOnly = lowerValue === 'true' || lowerValue === '1';
  } else {
    // Empty string or other value - treat as false
    enabledCampaignsOnly = false;
  }

  // Handle activePlacementsOnly checkbox
  let activePlacementsOnly = false;
  if (activePlacementsOnlyRaw === null || activePlacementsOnlyRaw === undefined) {
    activePlacementsOnly = false;
  } else if (typeof activePlacementsOnlyRaw === 'boolean') {
    activePlacementsOnly = activePlacementsOnlyRaw;
  } else if (typeof activePlacementsOnlyRaw === 'string') {
    const lowerValue = activePlacementsOnlyRaw.toLowerCase().trim();
    activePlacementsOnly = lowerValue === 'true' || lowerValue === '1';
  } else {
    activePlacementsOnly = false;
  }

  // ChatGPT settings are now read from ChatGPT sheet (removed from here)

  // Get placement type filters
  const placementTypeFilters = getPlacementTypeFilters(settingsSheet);

  // Get all filter lists from Lists sheet (returns objects with enabled and list)
  const placementContainsData = getPlacementContainsList(listsSheet);
  const placementNotContainsData = getPlacementNotContainsList(listsSheet);
  const displayNameContainsData = getDisplayNameContainsList(listsSheet);
  const displayNameNotContainsData = getDisplayNameNotContainsList(listsSheet);
  const targetUrlContainsData = getTargetUrlContainsList(listsSheet);
  const targetUrlNotContainsData = getTargetUrlNotContainsList(listsSheet);
  const targetUrlEndsWithData = getTargetUrlEndsWithList(listsSheet);
  const targetUrlNotEndsWithData = getTargetUrlNotEndsWithList(listsSheet);

  // Handle Max Results: if 0, ignore (no limit), if empty/null use 0 (no limit), otherwise use the value
  // Note: Default value is only used when initially populating the sheet, not when reading from it
  let maxResults = 0; // Default to 0 (no limit) when reading from sheet

  if (maxResultsRaw !== null && maxResultsRaw !== undefined && maxResultsRaw !== '') {
    // Handle both string and number types
    let maxResultsNum;
    if (typeof maxResultsRaw === 'number') {
      maxResultsNum = maxResultsRaw;
    } else {
      const trimmed = String(maxResultsRaw).trim();
      if (trimmed === '') {
        maxResultsNum = 0; // Empty string means no limit
      } else {
        maxResultsNum = parseInt(trimmed, 10);
      }
    }

    debugLog(`  Max Results parsed value: ${maxResultsNum} (isNaN: ${isNaN(maxResultsNum)})`)

    if (!isNaN(maxResultsNum)) {
      maxResults = maxResultsNum; // Use the parsed value (0 means no limit, any other number is the limit)
    } else {
      // If parsing failed, use 0 (no limit)
      maxResults = 0;
    }
  }

  debugLog(`  Max Results final value: ${maxResults}`)

  // Parse numeric values - always use what's in the sheet
  // Empty strings are treated as 0, null/undefined only used if setting row not found (malformed sheet)
  const lookbackWindowDays = (lookbackWindowDaysRaw !== null && lookbackWindowDaysRaw !== undefined)
    ? (typeof lookbackWindowDaysRaw === 'number' ? lookbackWindowDaysRaw : (lookbackWindowDaysRaw === '' ? 0 : parseInt(lookbackWindowDaysRaw, 10) || 0))
    : 0; // Use 0 if setting row not found (malformed sheet)
  const minimumImpressions = (minimumImpressionsRaw !== null && minimumImpressionsRaw !== undefined)
    ? (typeof minimumImpressionsRaw === 'number' ? minimumImpressionsRaw : (minimumImpressionsRaw === '' ? 0 : parseInt(minimumImpressionsRaw, 10) || 0))
    : 0; // Use 0 if setting row not found (malformed sheet)
  const minimumClicks = (minimumClicksRaw !== null && minimumClicksRaw !== undefined)
    ? (typeof minimumClicksRaw === 'number' ? minimumClicksRaw : (minimumClicksRaw === '' ? 0 : parseInt(minimumClicksRaw, 10) || 0))
    : 0; // Use 0 if setting row not found (malformed sheet)
  const minimumCost = (minimumCostRaw !== null && minimumCostRaw !== undefined)
    ? (typeof minimumCostRaw === 'number' ? minimumCostRaw : (minimumCostRaw === '' ? 0 : parseFloat(minimumCostRaw) || 0))
    : 0; // Use 0 if setting row not found (malformed sheet)
  const maximumConversions = (maximumConversionsRaw !== null && maximumConversionsRaw !== undefined)
    ? (typeof maximumConversionsRaw === 'number' ? maximumConversionsRaw : (maximumConversionsRaw === '' ? 0 : parseFloat(maximumConversionsRaw) || 0))
    : 0; // Use 0 if setting row not found (malformed sheet)

  const settings = {
    sharedExclusionListName: sharedExclusionListName,
    lookbackWindowDays: lookbackWindowDays,
    minimumImpressions: minimumImpressions,
    minimumClicks: minimumClicks,
    minimumCost: minimumCost,
    maximumConversions: maximumConversions,
    maxResults: maxResults,
    campaignNameContains: String(getSettingValue(settingsSheet, 'Campaign Name Contains', null) || '').trim(),
    campaignNameNotContains: String(getSettingValue(settingsSheet, 'Campaign Name Not Contains', null) || '').trim(),
    enabledCampaignsOnly: enabledCampaignsOnly,
    activePlacementsOnly: activePlacementsOnly,
    placementTypeFilters: placementTypeFilters,
    // Placement field filters
    placementContainsEnabled: placementContainsData.enabled,
    placementContainsList: placementContainsData.list,
    placementNotContainsEnabled: placementNotContainsData.enabled,
    placementNotContainsList: placementNotContainsData.list,
    // Display Name field filters
    displayNameContainsEnabled: displayNameContainsData.enabled,
    displayNameContainsList: displayNameContainsData.list,
    displayNameNotContainsEnabled: displayNameNotContainsData.enabled,
    displayNameNotContainsList: displayNameNotContainsData.list,
    // Target URL field filters
    targetUrlContainsEnabled: targetUrlContainsData.enabled,
    targetUrlContainsList: targetUrlContainsData.list,
    targetUrlNotContainsEnabled: targetUrlNotContainsData.enabled,
    targetUrlNotContainsList: targetUrlNotContainsData.list,
    targetUrlEndsWithEnabled: targetUrlEndsWithData.enabled,
    targetUrlEndsWithList: targetUrlEndsWithData.list,
    targetUrlNotEndsWithEnabled: targetUrlNotEndsWithData.enabled,
    targetUrlNotEndsWithList: targetUrlNotEndsWithData.list,
    sheet: settingsSheet
  };

  logSettingsDebug(settings);

  return settings;
}

/**
 * Gets a setting value from the settings sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet - The settings sheet
 * @param {string} settingName - The name of the setting to find
 * @param {*} defaultValue - Default value if setting row not found (only used when row doesn't exist)
 * @param {number} startRowIndex - Optional starting row index (0-based). Defaults to 3 (row 4) for Settings sheet
 * @returns {*} The setting value from the sheet, or defaultValue if the setting row doesn't exist
 * Note: If the setting row exists but the value is empty, returns the empty value (not the default)
 */
function getSettingValue(settingsSheet, settingName, defaultValue, startRowIndex = 3) {
  const dataRange = settingsSheet.getDataRange();
  const values = dataRange.getValues();

  debugLog(`  Searching for "${settingName}" starting from row ${startRowIndex + 1}`)

  // Start from the specified row index (defaults to 3 for row 4)
  for (let rowIndex = startRowIndex; rowIndex < values.length; rowIndex++) {
    const cellValue = String(values[rowIndex][0] || '').trim();
    if (cellValue === settingName) {
      const foundValue = values[rowIndex][1];
      debugLog(`  Found "${settingName}" at row ${rowIndex + 1}, value: ${foundValue} (type: ${typeof foundValue})`)
      // Setting row found - always return the value from the sheet (even if empty)
      // This ensures defaults are only used when populating empty sheet, not when reading
      return foundValue;
    }
  }

  debugLog(`  "${settingName}" not found, using default: ${defaultValue}`)

  // Setting row not found - return default (only happens if sheet is malformed)
  return defaultValue;
}

/**
 * Gets the Placement Contains list (Wanted Placement - Column F)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getPlacementContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'placementContains', 6);
}

/**
 * Gets the Placement Not Contains list (Unwanted Placement - Column A)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getPlacementNotContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'placementNotContains', 1);
}

/**
 * Gets the Display Name Contains list (Wanted Display Name - Column G)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getDisplayNameContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'displayNameContains', 7);
}

/**
 * Gets the Display Name Not Contains list (Unwanted Display Name - Column B)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getDisplayNameNotContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'displayNameNotContains', 2);
}

/**
 * Gets the Target URL Contains list (Wanted Target URL - Column H)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getTargetUrlContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'targetUrlContains', 8);
}

/**
 * Gets the Target URL Not Contains list (Unwanted Target URL - Column C)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getTargetUrlNotContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'targetUrlNotContains', 3);
}

/**
 * Gets the Target URL Ends With list by combining Target TLDs (Column I) and Target Region TLDs (Column J)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and combined list array
 */
function getTargetUrlEndsWithList(listsSheet) {
  const targetTldsData = getListFromSheet(listsSheet, 'targetTlds', 9);
  const targetRegionTldsData = getListFromSheet(listsSheet, 'targetRegionTlds', 10);

  const isEnabled = targetTldsData.enabled || targetRegionTldsData.enabled;
  const combinedList = [...targetTldsData.list, ...targetRegionTldsData.list];

  return { enabled: isEnabled, list: combinedList };
}

/**
 * Gets the Target URL Not Ends With list by combining Spammy TLDs (Column D) and Out-of-region TLDs (Column E)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and combined list array
 */
function getTargetUrlNotEndsWithList(listsSheet) {
  const spammyTldsData = getListFromSheet(listsSheet, 'spammyTlds', 4);
  const outOfRegionTldsData = getListFromSheet(listsSheet, 'outOfRegionTlds', 5);

  const isEnabled = spammyTldsData.enabled || outOfRegionTldsData.enabled;
  const combinedList = [...spammyTldsData.list, ...outOfRegionTldsData.list];

  return { enabled: isEnabled, list: combinedList };
}

/**
 * Generic function to get a list and enabled status from the Placement Filters sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @param {string} listKey - The key in PLACEMENT_FILTERS_CELL_REFERENCES (e.g., 'placementContains')
 * @param {number} columnIndex - The column index (1 = A, 2 = B, etc.)
 * @returns {Object} Object with enabled status and list array {enabled: boolean, list: Array<string>}
 */
function getListFromSheet(listsSheet, listKey, columnIndex) {
  const list = [];

  const checkboxRef = PLACEMENT_FILTERS_CELL_REFERENCES[listKey];
  if (!checkboxRef) {
    debugLog(`  ${listKey} not found in PLACEMENT_FILTERS_CELL_REFERENCES`);
    return { enabled: false, list: [] };
  }

  const checkboxValue = listsSheet.getRange(checkboxRef.row, checkboxRef.column).getValue();
  let enabled = false;
  if (typeof checkboxValue === 'boolean') {
    enabled = checkboxValue;
  } else if (typeof checkboxValue === 'string') {
    const lowerValue = checkboxValue.toLowerCase().trim();
    enabled = lowerValue === 'true' || lowerValue === '1';
  }

  const listStartRow = PLACEMENT_FILTERS_LIST_START_ROWS[listKey];
  if (!listStartRow) {
    debugLog(`  ${listKey} not found in PLACEMENT_FILTERS_LIST_START_ROWS`);
    return { enabled: enabled, list: [] };
  }

  const dataRange = listsSheet.getDataRange();
  const values = dataRange.getValues();
  const columnArrayIndex = columnIndex - 1;
  const listStartRowIndex = listStartRow - 1;

  for (let rowIndex = listStartRowIndex; rowIndex < values.length; rowIndex++) {
    const cellValue = values[rowIndex][columnArrayIndex];
    const item = String(cellValue || '').trim();

    if (item) {
      list.push(item);
    }
  }

  debugLog(`  ${listKey}: ${enabled ? 'ENABLED' : 'DISABLED'}, ${list.length} items`);
  if (DEBUG_MODE && list.length > 0 && list.length <= 5) {
    debugLog(`    Items: ${list.join(', ')}`);
  } else if (DEBUG_MODE && list.length > 5) {
    debugLog(`    First 5: ${list.slice(0, 5).join(', ')} ... and ${list.length - 5} more`);
  }

  return { enabled: enabled, list: list };
}

/**
 * Gets placement type filters from the settings sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet - The settings sheet
 * @returns {Object} Object with placement type filter flags
 */
function getPlacementTypeFilters(settingsSheet) {
  const refs = SETTINGS_CELL_REFERENCES.placementTypes;

  // Helper function to read checkbox value
  const getCheckboxValue = (row, column) => {
    const value = settingsSheet.getRange(row, column).getValue();
    if (typeof value === 'boolean') {
      return value;
    } else if (typeof value === 'string') {
      const lowerValue = value.toLowerCase().trim();
      return lowerValue === 'true' || lowerValue === '1';
    }
    return false;
  };

  const filters = {
    youtubeVideo: {
      enabled: getCheckboxValue(refs.youtubeVideo.row, refs.youtubeVideo.enabledColumn),
      automated: getCheckboxValue(refs.youtubeVideo.row, refs.youtubeVideo.automatedColumn)
    },
    website: {
      enabled: getCheckboxValue(refs.website.row, refs.website.enabledColumn),
      automated: getCheckboxValue(refs.website.row, refs.website.automatedColumn)
    },
    mobileApplication: {
      enabled: getCheckboxValue(refs.mobileApplication.row, refs.mobileApplication.enabledColumn),
      automated: getCheckboxValue(refs.mobileApplication.row, refs.mobileApplication.automatedColumn)
    },
    googleProducts: {
      enabled: getCheckboxValue(refs.googleProducts.row, refs.googleProducts.enabledColumn),
      automated: getCheckboxValue(refs.googleProducts.row, refs.googleProducts.automatedColumn)
    }
  };

  return filters;
}


/**
 * Gets ChatGPT settings from the ChatGPT sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} chatGptSheet - The ChatGPT sheet
 * @returns {Object} ChatGPT settings object
 */
function getChatGptSettingsFromSheet(chatGptSheet) {
  const refs = CHATGPT_CELL_REFERENCES;

  // Helper function to read a cell value
  const getCellValue = (cellRef, settingName = '') => {
    const row = cellRef.row;
    const col = cellRef.column;
    const value = chatGptSheet.getRange(row, col).getValue();
    if (DEBUG_MODE && settingName) {
      console.log(`  Reading ${settingName} from row ${row}, column ${col}: ${value} (type: ${typeof value})`);
    }
    return value;
  };

  // Helper function to read checkbox value
  const getCheckboxValue = (cellRef, settingName = '') => {
    const value = getCellValue(cellRef, settingName);
    if (typeof value === 'boolean') {
      return value;
    } else if (typeof value === 'string') {
      const lowerValue = value.toLowerCase().trim();
      return lowerValue === 'true' || lowerValue === '1';
    }
    return false;
  };

  // Read ChatGPT settings using cell references
  const enableChatGptRaw = getCheckboxValue(refs.enableChatGpt, 'Enable ChatGPT');
  const enableChatGpt = enableChatGptRaw;

  const chatGptApiKeyRaw = getCellValue(refs.chatGptApiKey, 'ChatGPT API Key');
  const chatGptApiKey = String(chatGptApiKeyRaw || '').trim();

  const useCachedChatGptRaw = getCheckboxValue(refs.useCachedChatGpt, 'Use Cached ChatGPT Responses');
  const useCachedChatGpt = useCachedChatGptRaw;

  const chatGptPromptRaw = getCellValue(refs.chatGptPrompt, 'ChatGPT Prompt');
  const chatGptPrompt = String(chatGptPromptRaw || '').trim();

  // Read Response Filters using cell references
  const responseContainsData = getResponseContainsList(chatGptSheet);
  const responseNotContainsData = getResponseNotContainsList(chatGptSheet);

  return {
    enableChatGpt: enableChatGpt,
    chatGptApiKey: chatGptApiKey,
    useCachedChatGpt: useCachedChatGpt,
    chatGptPrompt: chatGptPrompt,
    responseContainsList: responseContainsData.list,
    responseNotContainsList: responseNotContainsData.list
  };
}

/**
 * Gets the Response Contains list and enabled status from the ChatGPT sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} chatGptSheet - The ChatGPT sheet
 * @returns {Object} Object with enabled status and list array
 */
function getResponseContainsList(chatGptSheet) {
  const refs = CHATGPT_CELL_REFERENCES;
  const cellRef = refs.responseContains;

  // Read value from cell B13
  const value = chatGptSheet.getRange(cellRef.row, cellRef.column).getValue();
  const list = [];

  if (value !== null && value !== undefined && value !== '') {
    const trimmedValue = String(value).trim();
    if (trimmedValue) {
      // Check if it's comma-separated values and split them
      const valuesArray = trimmedValue.split(',').map(v => v.trim()).filter(v => v);
      list.push(...valuesArray);
    }
  }

  // Also read from rows below if there are multiple values
  const dataRange = chatGptSheet.getDataRange();
  const values = dataRange.getValues();
  const headerRowIndex = cellRef.row - 1; // Convert to 0-based array index

  for (let rowIndex = headerRowIndex + 1; rowIndex < values.length; rowIndex++) {
    // Stop if we hit the next header (Response Not Contains)
    if (values[rowIndex][0] === 'Response Not Contains') {
      break;
    }

    const rowValue = values[rowIndex][cellRef.column - 1];
    if (rowValue !== null && rowValue !== undefined && rowValue !== '') {
      const trimmedValue = String(rowValue).trim();
      if (trimmedValue) {
        const valuesArray = trimmedValue.split(',').map(v => v.trim()).filter(v => v);
        list.push(...valuesArray);
      }
    }
  }

  return { enabled: list.length > 0, list: list };
}

/**
 * Gets the Response Not Contains list and enabled status from the ChatGPT sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} chatGptSheet - The ChatGPT sheet
 * @returns {Object} Object with enabled status and list array
 */
function getResponseNotContainsList(chatGptSheet) {
  const refs = CHATGPT_CELL_REFERENCES;
  const cellRef = refs.responseNotContains;

  // Read value from cell B14
  const value = chatGptSheet.getRange(cellRef.row, cellRef.column).getValue();
  const list = [];

  if (value !== null && value !== undefined && value !== '') {
    const trimmedValue = String(value).trim();
    if (trimmedValue) {
      // Check if it's comma-separated values and split them
      const valuesArray = trimmedValue.split(',').map(v => v.trim()).filter(v => v);
      list.push(...valuesArray);
    }
  }

  // Also read from rows below if there are multiple values
  const dataRange = chatGptSheet.getDataRange();
  const values = dataRange.getValues();
  for (let rowIndex = cellRef.row; rowIndex < values.length; rowIndex++) {
    const rowValue = values[rowIndex][cellRef.column - 1];
    if (rowValue !== null && rowValue !== undefined && rowValue !== '') {
      const trimmedValue = String(rowValue).trim();
      if (trimmedValue) {
        const valuesArray = trimmedValue.split(',').map(v => v.trim()).filter(v => v);
        list.push(...valuesArray);
      }
    }
  }

  return { enabled: list.length > 0, list: list };
}

// --- Data Collection Functions ---

/**
 * Gets placement data from Performance Max and Display campaigns
 * @param {Object} settings - Settings object with filters and date range
 * @returns {Array<Object>} Array of flat placement objects
 */
function getPlacementData(settings) {
  const dateRange = getDateRange(settings.lookbackWindowDays);
  const allPlacementData = [];
  let pMaxError = null;
  let displayError = null;
  let videoError = null;

  // Get Performance Max placements
  console.log(`Getting placements from Performance Max campaigns...`);
  const pMaxQuery = getPlacementGaqlQuery(dateRange, settings, 'PERFORMANCE_MAX');
  console.log(`GAQL Query (Performance Max):`);
  console.log(pMaxQuery);
  console.log(``);

  try {
    const pMaxReport = executePlacementReport(pMaxQuery);
    const pMaxPlacements = extractPlacementDataFromReport(pMaxReport, 'PERFORMANCE_MAX');
    console.log(`Found ${pMaxPlacements.length} placements from Performance Max campaigns`);
    allPlacementData.push(...pMaxPlacements);
  } catch (error) {
    pMaxError = error;
    console.error(`Error getting Performance Max placements: ${error.message}`);
  }

  // Get Display campaign placements
  console.log(`Getting placements from Display campaigns...`);
  const displayQuery = getPlacementGaqlQuery(dateRange, settings, 'DISPLAY');
  console.log(`GAQL Query (Display):`);
  console.log(displayQuery);
  console.log(``);

  try {
    const displayReport = executePlacementReport(displayQuery);
    const displayPlacements = extractPlacementDataFromReport(displayReport, 'DISPLAY');
    console.log(`Found ${displayPlacements.length} placements from Display campaigns`);
    allPlacementData.push(...displayPlacements);
  } catch (error) {
    displayError = error;
    console.error(`Error getting Display placements: ${error.message}`);
  }

  // Get Video/YouTube campaign placements
  console.log(`Getting placements from Video/YouTube campaigns...`);
  const videoQuery = getPlacementGaqlQuery(dateRange, settings, 'VIDEO');
  console.log(`GAQL Query (Video):`);
  console.log(videoQuery);
  console.log(``);

  try {
    const videoReport = executePlacementReport(videoQuery);
    const videoPlacements = extractPlacementDataFromReport(videoReport, 'VIDEO');
    console.log(`Found ${videoPlacements.length} placements from Video/YouTube campaigns`);
    allPlacementData.push(...videoPlacements);
  } catch (error) {
    videoError = error;
    console.error(`Error getting Video placements: ${error.message}`);
  }

  // If ALL queries failed, throw an error - the script can't proceed without any data source working
  if (pMaxError && displayError && videoError) {
    throw new Error(`Failed to get placement data from all campaign types. PMax error: ${pMaxError.message}. Display error: ${displayError.message}. Video error: ${videoError.message}`);
  }

  if (allPlacementData.length > 0) {
    console.warn(`⚠ Note: performance_max_placement_view only supports impressions metric.`);
    console.warn(`   Display and Video campaign placements have additional metrics available.`);
  }

  // Apply campaign name filters in JavaScript (case-insensitive)
  const filteredData = filterPlacementsByCampaignName(allPlacementData, settings);

  if (filteredData.length !== allPlacementData.length) {
    console.log(`Filtered from ${allPlacementData.length} to ${filteredData.length} placements based on campaign name filters`);
  }

  return filteredData;
}

/**
 * Filters placements by campaign name contains/not contains (case-insensitive)
 * @param {Array<Object>} placements - Array of placement objects
 * @param {Object} settings - Settings object with campaignNameContains and campaignNameNotContains
 * @returns {Array<Object>} Filtered array of placement objects
 */
function filterPlacementsByCampaignName(placements, settings) {
  const containsFilter = settings.campaignNameContains ? settings.campaignNameContains.trim().toLowerCase() : '';
  const notContainsFilter = settings.campaignNameNotContains ? settings.campaignNameNotContains.trim().toLowerCase() : '';

  if (!containsFilter && !notContainsFilter) {
    return placements;
  }

  return placements.filter(placement => {
    const campaignNameLower = (placement.campaignName || '').toLowerCase();

    // Check "contains" filter
    if (containsFilter && !campaignNameLower.includes(containsFilter)) {
      return false;
    }

    // Check "not contains" filter
    if (notContainsFilter && campaignNameLower.includes(notContainsFilter)) {
      return false;
    }

    return true;
  });
}

/**
 * Gets the date range for the query
 * @param {number} lookbackDays - Number of days to look back
 * @returns {Object} Object with startDate and endDate strings
 */
function getDateRange(lookbackDays) {
  const endDate = getGoogleAdsApiFormattedDate(0);
  const startDate = getGoogleAdsApiFormattedDate(lookbackDays);
  return { startDate, endDate };
}

/**
 * Formats a date for Google Ads API (YYYY-MM-DD)
 * @param {number} daysAgo - Number of days ago (0 = today)
 * @returns {string} Formatted date string
 */
function getGoogleAdsApiFormattedDate(daysAgo = 0) {
  const date = new Date();
  date.setDate(date.getDate() - daysAgo);

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}

/**
 * Builds the GAQL query for placement data
 * @param {Object} dateRange - Object with startDate and endDate
 * @param {Object} settings - Settings object with campaign filters
 * @param {string} campaignType - 'PERFORMANCE_MAX', 'DISPLAY', or 'VIDEO'
 * @returns {string} GAQL query string
 */
function getPlacementGaqlQuery(dateRange, settings, campaignType) {
  const conditions = [
    `segments.date BETWEEN '${dateRange.startDate}' AND '${dateRange.endDate}'`,
    `campaign.advertising_channel_type = '${campaignType}'`
  ];

  // Campaign status filter
  if (settings.enabledCampaignsOnly) {
    conditions.push(`campaign.status = 'ENABLED'`);
  } else {
    conditions.push(`campaign.status IN ('ENABLED', 'PAUSED')`);
  }

  // Campaign name filters (case-sensitive LIKE)
  if (settings.campaignNameContains && settings.campaignNameContains.trim() !== '') {
    conditions.push(`campaign.name LIKE '%${settings.campaignNameContains}%'`);
  }
  if (settings.campaignNameNotContains && settings.campaignNameNotContains.trim() !== '') {
    conditions.push(`campaign.name NOT LIKE '%${settings.campaignNameNotContains}%'`);
  }

  // Impressions threshold (available for all campaign types)
  if (settings.minimumImpressions > 0) {
    conditions.push(`metrics.impressions > ${settings.minimumImpressions}`);
  }

  // Add LIMIT clause if maxResults is set (0 means no limit)
  const limitClause = settings.maxResults > 0 ? ` LIMIT ${settings.maxResults}` : '';

  if (campaignType === 'PERFORMANCE_MAX') {
    // PMax only has impressions metric available
    const whereClause = conditions.join(' AND ');
    const query = `
      SELECT
        campaign.id,
        campaign.name,
        campaign.advertising_channel_type,
        campaign.status,
        performance_max_placement_view.display_name,
        performance_max_placement_view.placement,
        performance_max_placement_view.placement_type,
        performance_max_placement_view.target_url,
        metrics.impressions
      FROM performance_max_placement_view
      WHERE ${whereClause}
      ORDER BY metrics.impressions DESC${limitClause}
    `;
    return query;
  } else if (campaignType === 'DISPLAY' || campaignType === 'VIDEO') {
    // Display/Video has full metrics - add clicks, cost, conversions filters
    if (settings.minimumClicks > 0) {
      conditions.push(`metrics.clicks > ${settings.minimumClicks}`);
    }
    if (settings.minimumCost > 0) {
      const costMicros = settings.minimumCost * 1000000;
      conditions.push(`metrics.cost_micros > ${costMicros}`);
    }
    if (settings.maximumConversions > 0) {
      conditions.push(`metrics.conversions < ${settings.maximumConversions}`);
    }

    const whereClause = conditions.join(' AND ');
    const query = `
      SELECT
        campaign.id,
        campaign.name,
        campaign.advertising_channel_type,
        campaign.status,
        detail_placement_view.display_name,
        detail_placement_view.placement_type,
        detail_placement_view.resource_name,
        detail_placement_view.placement,
        detail_placement_view.group_placement_target_url,
        detail_placement_view.target_url,
        metrics.impressions,
        metrics.clicks,
        metrics.cost_micros,
        metrics.conversions,
        metrics.conversions_value
      FROM detail_placement_view
      WHERE ${whereClause}
      ORDER BY metrics.clicks DESC${limitClause}
    `;
    return query;
  }

  throw new Error(`Unsupported campaign type: ${campaignType}`);
}

/**
 * Executes the placement report query
 * @param {string} query - GAQL query string
 * @returns {GoogleAppsScript.AdsApp.Report} The report object
 */
function executePlacementReport(query) {
  try {
    const report = AdsApp.report(query);
    const rows = report.rows();
    const totalRows = rows.totalNumEntities();

    console.log(`Number of rows from query: ${totalRows}`);

    if (totalRows === 0) {
      console.warn(`No placement data found for the specified filters and date range`);
    }

    return report;
  } catch (error) {
    console.error(`Error executing placement report query:`);
    console.error(error.message);
    console.error(`\nPlease validate your GAQL query at:`);
    console.error(`https://developers.google.com/google-ads/api/fields/v20/query_validator`);
    console.error(`\nOr use the query builder at:`);
    console.error(`For Performance Max: https://developers.google.com/google-ads/api/fields/v20/performance_max_placement_view_query_builder`);
    console.error(`For Display/Video: https://developers.google.com/google-ads/api/fields/v20/detail_placement_view_query_builder`);
    throw error;
  }
}

/**
 * Extracts placement data from the report and returns flat objects
 * @param {GoogleAppsScript.AdsApp.Report} report - The report object
 * @param {string} campaignType - 'PERFORMANCE_MAX', 'DISPLAY', or 'VIDEO'
 * @returns {Array<Object>} Array of flat placement objects
 */
function extractPlacementDataFromReport(report, campaignType) {
  const placementData = [];
  const rows = report.rows();
  let rowCount = 0;

  // Both DISPLAY and VIDEO use detail_placement_view
  const usesDetailPlacementView = campaignType === 'DISPLAY' || campaignType === 'VIDEO';
  const viewPrefix = campaignType === 'PERFORMANCE_MAX' ? 'performance_max_placement_view' : 'detail_placement_view';

  while (rows.hasNext()) {
    const row = rows.next();
    rowCount++;

    const placement = {
      campaignId: row['campaign.id'],
      campaignName: row['campaign.name'],
      displayName: row[`${viewPrefix}.display_name`],
      placement: row[`${viewPrefix}.placement`],
      placementType: row[`${viewPrefix}.placement_type`],
      targetUrl: usesDetailPlacementView
        ? (row['detail_placement_view.group_placement_target_url'] || row['detail_placement_view.target_url'] || '')
        : row[`${viewPrefix}.target_url`],
      resourceName: usesDetailPlacementView ? (row['detail_placement_view.resource_name'] || '') : '',
      impressions: parseInt(row['metrics.impressions']) || 0,
      clicks: usesDetailPlacementView ? parseInt(row['metrics.clicks']) || 0 : 0,
      costMicros: usesDetailPlacementView ? parseInt(row['metrics.cost_micros']) || 0 : 0,
      conversions: usesDetailPlacementView ? parseFloat(row['metrics.conversions']) || 0 : 0,
      conversionsValue: usesDetailPlacementView ? parseFloat(row['metrics.conversions_value']) || 0 : 0
    };

    placementData.push(placement);

    if (DEBUG_MODE && rowCount <= 3) {
      console.log(`Placement ${rowCount} (${campaignType}):`);
      console.log(`  Campaign: ${placement.campaignName}`);
      console.log(`  Placement: ${placement.placement}`);
      console.log(`  Impressions: ${placement.impressions}`);
      console.log(`  Clicks: ${placement.clicks}`);
      console.log(``);
    }
  }

  if (DEBUG_MODE && rowCount > 3) {
    console.log(`... and ${rowCount - 3} more placements from ${campaignType} campaigns`);
    console.log(``);
  }

  return placementData;
}

// --- Data Transformation Functions ---

/**
 * Calculates additional metrics for placements and applies filters
 * @param {Array<Object>} placementData - Raw placement data
 * @returns {Array<Object>} Transformed placement data with calculated metrics
 */
function calculatePlacementMetrics(placementData) {
  const transformedData = [];

  for (const placement of placementData) {
    const cost = placement.costMicros / 1000000;
    const ctr = placement.impressions > 0 ? placement.clicks / placement.impressions : 0;
    const avgCpc = placement.clicks > 0 ? cost / placement.clicks : 0;
    const convRate = placement.clicks > 0 ? placement.conversions / placement.clicks : 0;
    const cpa = placement.conversions > 0 ? cost / placement.conversions : 0;
    const roas = placement.conversionsValue > 0 ? placement.conversionsValue / cost : 0;

    const transformed = {
      campaignId: placement.campaignId,
      campaignName: placement.campaignName,
      displayName: placement.displayName,
      placement: placement.placement,
      placementType: placement.placementType,
      targetUrl: placement.targetUrl,
      impressions: placement.impressions,
      clicks: placement.clicks,
      cost: cost,
      conversions: placement.conversions,
      conversionsValue: placement.conversionsValue,
      ctr: ctr,
      avgCpc: avgCpc,
      convRate: convRate,
      cpa: cpa,
      roas: roas
    };

    transformedData.push(transformed);
  }

  return transformedData;
}

/**
 * Filters placement data based on minimum thresholds
 * @param {Array<Object>} placementData - Transformed placement data
 * @param {Object} settings - Settings object with thresholds
 * @returns {Array<Object>} Filtered placement data
 */
function filterPlacementData(placementData, settings) {
  // Note: Metric thresholds (impressions, clicks, cost, conversions) and campaign name filters
  // are now applied in GAQL for better performance. This function handles placement-specific filters.
  //
  // WANTED filters (Columns F-J) = Allowlist - if matches, HIDE from report (it's good)
  // UNWANTED filters (Columns A-E) = Flagging only, not used for filtering

  return placementData.filter(placement => {
    // Check placement type filter
    let meetsPlacementTypeFilter = true;
    if (settings.placementTypeFilters) {
      const placementType = String(placement.placementType || '').toUpperCase();
      if (placementType.includes('YOUTUBE') || placementType.includes('YOUTUBE_VIDEO')) {
        meetsPlacementTypeFilter = settings.placementTypeFilters.youtubeVideo.enabled;
      } else if (placementType.includes('WEBSITE')) {
        meetsPlacementTypeFilter = settings.placementTypeFilters.website.enabled;
      } else if (placementType.includes('MOBILE_APPLI')) {
        meetsPlacementTypeFilter = settings.placementTypeFilters.mobileApplication.enabled;
      } else if (placementType.includes('GOOGLE_PRODUCTS')) {
        meetsPlacementTypeFilter = settings.placementTypeFilters.googleProducts.enabled;
      } else {
        meetsPlacementTypeFilter = true;
      }

      if (DEBUG_MODE && !meetsPlacementTypeFilter) {
        console.log(`  Filtered out (placement type not selected): ${placement.placement} (${placement.placementType})`);
      }
    }

    // Column F: Wanted Placement (allowlist) - "The Placement field should include these words/phrases"
    // If matches → GOOD → HIDE from report
    let meetsWantedPlacementFilter = true;
    if (settings.placementContainsEnabled && settings.placementContainsList?.length > 0) {
      const matchesAllowlist = containsAny(placement.placement, settings.placementContainsList);
      meetsWantedPlacementFilter = !matchesAllowlist;
      if (!meetsWantedPlacementFilter) {
        debugLog(`  Hidden (Placement matches allowlist): ${placement.placement}`);
      }
    }

    // Column G: Wanted Display Name (allowlist) - "The Display Name field should include these words/phrases"
    // If matches → GOOD → HIDE from report
    let meetsWantedDisplayNameFilter = true;
    if (settings.displayNameContainsEnabled && settings.displayNameContainsList?.length > 0) {
      const matchesAllowlist = containsAny(placement.displayName, settings.displayNameContainsList);
      meetsWantedDisplayNameFilter = !matchesAllowlist;
      if (!meetsWantedDisplayNameFilter) {
        debugLog(`  Hidden (Display Name matches allowlist): ${placement.displayName}`);
      }
    }

    // Column H: Wanted Target URL (allowlist) - "The Target URL should include these words/phrases"
    // If matches → GOOD → HIDE from report
    let meetsWantedTargetUrlFilter = true;
    if (settings.targetUrlContainsEnabled && settings.targetUrlContainsList?.length > 0) {
      const matchesAllowlist = containsAny(placement.targetUrl, settings.targetUrlContainsList);
      meetsWantedTargetUrlFilter = !matchesAllowlist;
      if (!meetsWantedTargetUrlFilter) {
        debugLog(`  Hidden (Target URL matches allowlist): ${placement.targetUrl}`);
      }
    }

    // Columns I+J: Target TLDs (allowlist) - "The Target URL should end with this TLD"
    // If matches → GOOD → HIDE from report
    let meetsWantedTldsFilter = true;
    if (settings.targetUrlEndsWithEnabled && settings.targetUrlEndsWithList?.length > 0) {
      const matchesAllowlist = endsWithAny(placement.targetUrl, settings.targetUrlEndsWithList);
      meetsWantedTldsFilter = !matchesAllowlist;
      if (!meetsWantedTldsFilter) {
        debugLog(`  Hidden (Target URL ends with allowed TLD): ${placement.targetUrl}`);
      }
    }

    // Note: UNWANTED filters (Columns A-E) are for flagging only, not filtering
    // Placements matching unwanted patterns should APPEAR in the report so users can exclude them

    return meetsPlacementTypeFilter &&
      meetsWantedPlacementFilter && meetsWantedDisplayNameFilter &&
      meetsWantedTargetUrlFilter && meetsWantedTldsFilter;
  });
}

// --- Sheet Writing Functions ---

/**
 * Gets the list of excluded placements from the shared exclusion list
 * @param {string} listName - The name of the shared exclusion list
 * @returns {Set<string>} Set of excluded placement URLs
 */
function getExcludedPlacements(listName) {
  const excludedPlacements = new Set();

  console.log(`\n=== Getting Excluded Placements ===`);
  console.log(`Looking for exclusion list: ${listName}`);

  try {
    const listIterator = AdsApp.excludedPlacementLists()
      .withCondition(`shared_set.name = '${listName}'`)
      .get();

    if (!listIterator.hasNext()) {
      console.log(`⚠ Exclusion list not found`);
      return excludedPlacements; // List doesn't exist, return empty set
    }

    const excludedPlacementList = listIterator.next();
    const retrievedListName = excludedPlacementList.getName();
    console.log(`✓ Found exclusion list: ${retrievedListName}`);

    // Use iterator method to get excluded placements
    // Note: GAQL queries to shared_set_criterion are not supported in Google Ads Scripts
    const placementIterator = excludedPlacementList.excludedPlacements().get();

    let placementCount = 0;
    while (placementIterator.hasNext()) {
      try {
        const sharedPlacement = placementIterator.next();
        placementCount++;

        // Use getUrl() method (standard method per docs)
        const placementUrl = sharedPlacement.getUrl();

        if (placementUrl) {
          // Normalize website URL
          const normalized = normalizeUrl(placementUrl);
          if (normalized) {
            excludedPlacements.add(normalized);
            debugLog(`  Added: ${normalized}`)
          }
        }
      } catch (placementError) {
        console.error(`  ✗ Error processing placement: ${placementError.message}`);
      }
    }

    console.log(`Found ${excludedPlacements.size} unique placements in exclusion list`);
    if (placementCount === 0) {
      console.log(`  Note: Placements added via Bulk Upload are not readable via iterator.`);
    }
    console.log(`===================================\n`);

  } catch (error) {
    console.error(`✗ Error getting excluded placements: ${error.message}`);
  }

  return excludedPlacements;
}

/**
 * Checks if a placement is already excluded
 * Only checks website placements - returns "Unknown" for mobile apps, YouTube, and other types
 * @param {Object} placement - Placement object with placement, placementType, and targetUrl
 * @param {Set<string>} excludedPlacements - Set of excluded placements (normalized website URLs only)
 * @returns {string} "Excluded" if excluded, "Unknown" if can't determine, "Active" if not excluded
 */
function getPlacementStatus(placement, excludedPlacements) {
  if (!placement) {
    debugLog(`  Placement check: placement is null/undefined`)
    return 'Unknown';
  }

  const placementValue = placement.placement || '';
  const placementType = placement.placementType || '';
  const targetUrl = placement.targetUrl || '';
  const placementTypeUpper = String(placementType).toUpperCase();

  debugLog(`\n  Checking placement: value="${placementValue}", type="${placementType}", targetUrl="${targetUrl}"`)

  // Mobile apps - can't check status
  if (placementTypeUpper.includes('MOBILE_APPLI')) {
    debugLog(`    Mobile app placement - returning "Unknown"`)
    return 'Unknown';
  }

  // YouTube placements (channels and videos) - can't easily check status
  if (placementTypeUpper.includes('YOUTUBE')) {
    debugLog(`    YouTube placement - returning "Unknown"`)
    return 'Unknown';
  }

  // For website placements, check if excluded
  if (excludedPlacements.size === 0) {
    debugLog(`  Placement check: excluded placements set is empty - returning "Unknown"`)
    return 'Unknown';
  }

  // Check website placements only
  const url = targetUrl || placementValue;
  if (url) {
    debugLog(`    Website placement detected. URL: "${url}"`)

    const normalized = normalizeUrl(url);
    if (normalized) {
      debugLog(`    Normalized URL: "${normalized}"`)
      debugLog(`    Checking if "${normalized}" is in excluded set...`)

      if (excludedPlacements.has(normalized)) {
        debugLog(`    ✓ MATCH: Normalized URL "${normalized}" found in excluded set`)
        return 'Excluded';
      } else {
        debugLog(`    ✗ NO MATCH: Normalized URL not found in excluded set`)
        return 'Active';
      }
    } else {
      debugLog(`    ⚠ Could not normalize URL: "${url}" - returning "Unknown"`)
      return 'Unknown';
    }
  }

  // If we can't determine the type or URL, return "Unknown"
  debugLog(`    ⚠ Could not determine placement type or URL - returning "Unknown"`)
  return 'Unknown';
}

/**
 * Writes placement data to the sheet with checkboxes
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @param {Array<Object>} placementData - Transformed placement data
 * @param {Object} settings - Settings object with filters
 */
function writePlacementDataToSheet(spreadsheet, placementData, settings) {
  // Data is already filtered before this function is called
  console.log(`Writing ${placementData.length} placements to sheet (after applying filters)`);

  // Define all output sheet names
  const allOutputSheetNames = [
    WEBSITE_OUTPUT_SHEET_NAME,
    YOUTUBE_OUTPUT_SHEET_NAME,
    MOBILE_APPLICATION_OUTPUT_SHEET_NAME,
    GOOGLE_PRODUCTS_OUTPUT_SHEET_NAME
  ];

  // Clear ALL output sheets first (from row 3 onwards), regardless of whether there's data
  // Rows 1-2 are preserved for user content
  for (const sheetName of allOutputSheetNames) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn() || 1;
      if (lastRow >= 3) {
        sheet.getRange(3, 1, lastRow - 2, lastCol).clearContent();
      }
      debugLog(`Cleared sheet: ${sheetName} (from row 3 onwards)`)
    }
  }

  // If no data, write "No results found" message to all output sheets
  if (placementData.length === 0) {
    console.log(`No placement data found. Writing 'No results found' to all output sheets.`);
    const timestamp = new Date().toLocaleString();
    const noResultsData = [['No results found', `Last updated: ${timestamp}`]];

    for (const sheetName of allOutputSheetNames) {
      const sheet = getOrCreateOutputSheet(spreadsheet, sheetName);
      // Clear from row 3 onwards (keep rows 1-2 intact), preserving formatting
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn() || 2;
      if (lastRow >= 3) {
        sheet.getRange(3, 1, lastRow - 2, lastCol).clearContent();
      }
      // Write "No results found" to row 4 (below header row 3)
      sheet.getRange(4, 1, 1, 2).setValues(noResultsData);
      sheet.setTabColor('#ea4335'); // Red tab for no results
    }
    return;
  }

  // Get currently excluded placements
  const excludedPlacements = getExcludedPlacements(settings.sharedExclusionListName);
  debugLog(`Found ${excludedPlacements.size} placements already in exclusion list`)
  if (DEBUG_MODE && excludedPlacements.size > 0 && excludedPlacements.size <= 10) {
    const excludedArray = Array.from(excludedPlacements);
    debugLog(`Excluded placements: ${excludedArray.join(', ')}`)
  }

  // Filter out already-excluded placements if activePlacementsOnly is enabled
  let filteredPlacementData = placementData;
  if (settings.activePlacementsOnly && excludedPlacements.size > 0) {
    const originalCount = placementData.length;
    filteredPlacementData = placementData.filter(placement => {
      const status = getPlacementStatus(placement, excludedPlacements);
      return status !== 'Excluded';
    });
    const filteredOutCount = originalCount - filteredPlacementData.length;
    if (filteredOutCount > 0) {
      console.log(`Active Placements Only: Filtered out ${filteredOutCount} already-excluded placements`);
    }
  }

  // Headers for sheets WITH ChatGPT Response (Website, YouTube)
  const headersWithChatGpt = [
    'Exclude',
    'Status',
    'Campaign ID',
    'Campaign Name',
    'Placement',
    'Display Name',
    'Placement Type',
    'Target URL',
    'Notes',
    'ChatGPT Response',
    'Impressions',
    'Clicks',
    'Cost',
    'Conversions',
    'Conv. Value',
    'CTR (%)',
    'Avg. CPC',
    'Conv. Rate (%)',
    'CPA',
    'ROAS'
  ];

  // Headers for Mobile Application (no ChatGPT Response)
  const headersForMobileApp = [
    'Exclude',
    'Status',
    'Campaign ID',
    'Campaign Name',
    'Placement',
    'Display Name',
    'Placement Type',
    'Target URL',
    'Notes',
    'Impressions',
    'Clicks',
    'Cost',
    'Conversions',
    'Conv. Value',
    'CTR (%)',
    'Avg. CPC',
    'Conv. Rate (%)',
    'CPA',
    'ROAS'
  ];

  // Headers for Google Products (no Exclude column, no ChatGPT Response - cannot be excluded)
  const headersForGoogleProducts = [
    'Status',
    'Campaign ID',
    'Campaign Name',
    'Placement',
    'Display Name',
    'Placement Type',
    'Target URL',
    'Notes',
    'Impressions',
    'Clicks',
    'Cost',
    'Conversions',
    'Conv. Value',
    'CTR (%)',
    'Avg. CPC',
    'Conv. Rate (%)',
    'CPA',
    'ROAS'
  ];

  // Group placements by type
  const placementsByType = {};
  for (const placement of placementData) {
    const placementTypeName = getPlacementTypeName(placement.placementType);
    if (!placementTypeName) {
      // Unknown placement type - skip it
      debugLog(`Skipping placement with unknown type: ${placement.placement} (type: ${placement.placementType})`)
      continue;
    }
    if (!placementsByType[placementTypeName]) {
      placementsByType[placementTypeName] = [];
    }
    placementsByType[placementTypeName].push(placement);
  }

  // Track which sheets received data
  const sheetsWithData = new Set();

  // Write to placement type sheets that have data
  for (const placementTypeName in placementsByType) {
    const placements = placementsByType[placementTypeName];
    const placementTypeSheet = getOrCreatePlacementTypeSheet(spreadsheet, placementTypeName);
    sheetsWithData.add(placementTypeName);

    // Determine which headers to use based on placement type
    const isGoogleProducts = placementTypeName === GOOGLE_PRODUCTS_OUTPUT_SHEET_NAME;
    const isMobileApp = placementTypeName === MOBILE_APPLICATION_OUTPUT_SHEET_NAME;
    let headers;
    if (isGoogleProducts) {
      headers = headersForGoogleProducts;
    } else if (isMobileApp) {
      headers = headersForMobileApp;
    } else {
      headers = headersWithChatGpt;
    }

    // Clear existing data from row 3 onwards (keep rows 1-2 intact for user content)
    // Use clearContent to preserve formatting
    const lastRow = placementTypeSheet.getLastRow();
    const lastCol = placementTypeSheet.getLastColumn() || headers.length;
    if (lastRow >= 3) {
      placementTypeSheet.getRange(3, 1, lastRow - 2, lastCol).clearContent();
    }

    // Write headers to row 3
    const headerRange = placementTypeSheet.getRange(3, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');

    // Get automation setting for this placement type
    const isAutomated = getAutomationForPlacementType(settings, placementTypeName);

    // Write data rows
    const dataRows = placements.map(placement => {
      // Add notes based on placement type
      let notes = '';
      if (isGoogleProducts) {
        notes = 'Google Products placements cannot be excluded';
      } else if (isMobileApp) {
        notes = 'Will be excluded via bulk upload if selected';
      }

      // Check if placement is already excluded (only for website placements)
      const status = getPlacementStatus(placement, excludedPlacements);

      // Use automation setting for this placement type
      const shouldCheck = isAutomated || false;

      // Metrics columns (same for all types)
      const metricsColumns = [
        placement.impressions,
        placement.clicks,
        placement.cost,
        placement.conversions,
        placement.conversionsValue,
        placement.ctr,
        placement.avgCpc,
        placement.convRate,
        placement.cpa,
        placement.roas
      ];

      // Google Products: no Exclude column, no ChatGPT Response
      if (isGoogleProducts) {
        return [
          status,
          placement.campaignId,
          placement.campaignName,
          placement.placement,
          placement.displayName,
          placement.placementType,
          placement.targetUrl,
          notes,
          ...metricsColumns
        ];
      }

      // Mobile App: has Exclude column, no ChatGPT Response
      if (isMobileApp) {
        return [
          shouldCheck,
          status,
          placement.campaignId,
          placement.campaignName,
          placement.placement,
          placement.displayName,
          placement.placementType,
          placement.targetUrl,
          notes,
          ...metricsColumns
        ];
      }

      // Website and YouTube: has Exclude column and ChatGPT Response
      return [
        shouldCheck,
        status,
        placement.campaignId,
        placement.campaignName,
        placement.placement,
        placement.displayName,
        placement.placementType,
        placement.targetUrl,
        notes,
        placement.chatGptResponse || '',
        ...metricsColumns
      ];
    });

    if (dataRows.length > 0) {
      const dataRange = placementTypeSheet.getRange(4, 1, dataRows.length, headers.length);
      dataRange.setValues(dataRows);

      // Add checkboxes to first column (only for sheets that have Exclude column)
      if (!isGoogleProducts) {
        const checkboxRange = placementTypeSheet.getRange(4, 1, dataRows.length, 1);
        checkboxRange.insertCheckboxes();
      }

      // Format number columns based on sheet type
      // Google Products: no Exclude column, no ChatGPT column (columns shift left by 1)
      // Mobile App: has Exclude column, no ChatGPT column
      // Website/YouTube: has Exclude column, has ChatGPT column
      const hasChatGptColumn = !isGoogleProducts && !isMobileApp;
      const hasExcludeColumn = !isGoogleProducts;
      formatPlacementSheet(placementTypeSheet, placements.length, hasChatGptColumn, hasExcludeColumn);
    }

    // Freeze rows 1-3 (rows 1-2 for user content, row 3 for headers) and first column (or no columns for Google Products since no checkbox)
    placementTypeSheet.setFrozenRows(3);
    if (!isGoogleProducts) {
      placementTypeSheet.setFrozenColumns(1);
    }

    // Set tab color to GREEN for sheets with data
    placementTypeSheet.setTabColor('#34a853');

    console.log(`✓ Written ${placements.length} placements to ${placementTypeName} sheet`);
  }

  // For sheets that didn't receive data, write "No results found" with timestamp and set RED tab
  const timestamp = new Date().toLocaleString();
  const noResultsData = [['No results found', `Last updated: ${timestamp}`]];

  for (const sheetName of allOutputSheetNames) {
    if (!sheetsWithData.has(sheetName)) {
      const sheet = getOrCreateOutputSheet(spreadsheet, sheetName);
      // Clear from row 3 onwards (keep rows 1-2 intact), preserving formatting
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn() || 2;
      if (lastRow >= 3) {
        sheet.getRange(3, 1, lastRow - 2, lastCol).clearContent();
      }
      // Write "No results found" to row 4 (below header row 3)
      sheet.getRange(4, 1, 1, 2).setValues(noResultsData);
      sheet.setTabColor('#ea4335'); // Red tab for no results
      console.log(`✓ No data for ${sheetName} - marked as no results`);
    }
  }
}

/**
 * Gets or creates a placement type sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @param {string} placementTypeName - The name of the placement type (e.g., "YouTube Video", "Website")
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The placement type sheet
 */
function getOrCreatePlacementTypeSheet(spreadsheet, placementTypeName) {
  let placementTypeSheet = spreadsheet.getSheetByName(placementTypeName);
  if (!placementTypeSheet) {
    placementTypeSheet = spreadsheet.insertSheet(placementTypeName);
    // Color the sheet green to differentiate it from other sheets
    placementTypeSheet.setTabColor('#34a853'); // Google green color
    debugLog(`Created placement type sheet: ${placementTypeName}`)
  }
  return placementTypeSheet;
}

/**
 * Gets or creates an output sheet by name
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @param {string} sheetName - The name of the output sheet (e.g., "Output: Website")
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The output sheet
 */
function getOrCreateOutputSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.setTabColor('#34a853');
    debugLog(`Created output sheet: ${sheetName}`)
  }
  return sheet;
}

/**
 * Gets the output sheet name from a placement type value
 * @param {string} placementType - The placement type value (e.g., "YOUTUBE_VIDEO", "WEBSITE")
 * @returns {string | null} The output sheet name (e.g., "Output: YouTube", "Output: Website") or null if unknown
 */
function getPlacementTypeName(placementType) {
  if (!placementType) {
    return null;
  }
  const placementTypeUpper = String(placementType).toUpperCase();
  if (placementTypeUpper.includes('YOUTUBE') || placementTypeUpper.includes('YOUTUBE_VIDEO')) {
    return YOUTUBE_OUTPUT_SHEET_NAME;
  } else if (placementTypeUpper.includes('WEBSITE')) {
    return WEBSITE_OUTPUT_SHEET_NAME;
  } else if (placementTypeUpper.includes('MOBILE_APPLI')) {
    return MOBILE_APPLICATION_OUTPUT_SHEET_NAME;
  } else if (placementTypeUpper.includes('GOOGLE_PRODUCTS')) {
    return GOOGLE_PRODUCTS_OUTPUT_SHEET_NAME;
  }
  return null;
}

/**
 * Gets the automation setting for a placement type
 * @param {Object} settings - Settings object with placement type filters
 * @param {string} placementTypeName - The output sheet name (e.g., "Output: YouTube", "Output: Website")
 * @returns {boolean} True if automation is enabled for this placement type
 */
function getAutomationForPlacementType(settings, placementTypeName) {
  if (!settings.placementTypeFilters || !placementTypeName) {
    return false;
  }
  if (placementTypeName === YOUTUBE_OUTPUT_SHEET_NAME) {
    return settings.placementTypeFilters.youtubeVideo.automated || false;
  } else if (placementTypeName === WEBSITE_OUTPUT_SHEET_NAME) {
    return settings.placementTypeFilters.website.automated || false;
  } else if (placementTypeName === MOBILE_APPLICATION_OUTPUT_SHEET_NAME) {
    return settings.placementTypeFilters.mobileApplication.automated || false;
  } else if (placementTypeName === GOOGLE_PRODUCTS_OUTPUT_SHEET_NAME) {
    return settings.placementTypeFilters.googleProducts.automated || false;
  }
  return false;
}

/**
 * Formats the placement sheet with proper number formatting
 * Data starts at row 4 (rows 1-2 are user content, row 3 is headers)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to format
 * @param {number} numRows - Number of data rows
 */
function formatPlacementSheet(sheet, numRows, hasChatGptColumn = true, hasExcludeColumn = true) {
  // Calculate column offset based on which columns are present
  // Base columns (with Exclude + ChatGPT + Campaign ID): metrics start at column 11
  // Without ChatGPT: metrics start at column 10 (offset -1)
  // Without Exclude: metrics start 1 column earlier (additional offset -1)
  let offset = 0;
  if (!hasChatGptColumn) {
    offset -= 1;
  }
  if (!hasExcludeColumn) {
    offset -= 1;
  }

  // Data starts at row 4
  const dataStartRow = 4;

  // Format impressions, clicks, conversions (integers)
  sheet.getRange(dataStartRow, 11 + offset, numRows, 1).setNumberFormat('#,##0'); // Impressions
  sheet.getRange(dataStartRow, 12 + offset, numRows, 1).setNumberFormat('#,##0'); // Clicks
  sheet.getRange(dataStartRow, 14 + offset, numRows, 1).setNumberFormat('#,##0'); // Conversions

  // Format cost, CPA, Avg CPC (currency)
  sheet.getRange(dataStartRow, 13 + offset, numRows, 1).setNumberFormat('#,##0.00'); // Cost
  sheet.getRange(dataStartRow, 17 + offset, numRows, 1).setNumberFormat('#,##0.00'); // Avg CPC
  sheet.getRange(dataStartRow, 19 + offset, numRows, 1).setNumberFormat('#,##0.00'); // CPA

  // Format conversions value, ROAS (currency)
  sheet.getRange(dataStartRow, 15 + offset, numRows, 1).setNumberFormat('#,##0.00'); // Conv. Value
  sheet.getRange(dataStartRow, 20 + offset, numRows, 1).setNumberFormat('#,##0.00'); // ROAS

  // Format percentages
  sheet.getRange(dataStartRow, 16 + offset, numRows, 1).setNumberFormat('0.00%'); // CTR
  sheet.getRange(dataStartRow, 18 + offset, numRows, 1).setNumberFormat('0.00%'); // Conv. Rate

  // Format Notes column (clip text - no wrap)
  // Notes column position depends on whether Exclude column exists
  const notesColumn = hasExcludeColumn ? 9 : 8;
  sheet.getRange(dataStartRow, notesColumn, numRows, 1).setWrap(false);

  // Format ChatGPT Response column if it exists
  if (hasChatGptColumn) {
    const chatGptColumn = hasExcludeColumn ? 10 : 9;
    sheet.getRange(dataStartRow, chatGptColumn, numRows, 1).setWrap(false);
  }
}

// --- Exclusion List Functions ---

/**
 * Reads checked placements from all placement type sheets and normalizes them
 * Also collects full placement data for logging purposes
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @returns {Object} Object with websitePlacements, youtubeVideos (array of {placement, placementType}), mobileAppPlacements arrays, and fullPlacementData for logging
 */
function readCheckedPlacementsFromSheet(spreadsheet) {
  const checkedWebsitePlacementsSet = new Set();
  const checkedYouTubePlacements = []; // Store objects with { placement, placementType }
  const checkedMobileAppPlacementsSet = new Set();
  const fullPlacementData = []; // Array to store full placement data for logging
  let skippedInvalid = 0;
  let totalCheckedCount = 0;
  let totalUncheckedCount = 0;

  const placementTypeNames = [
    YOUTUBE_OUTPUT_SHEET_NAME,
    WEBSITE_OUTPUT_SHEET_NAME,
    MOBILE_APPLICATION_OUTPUT_SHEET_NAME,
    GOOGLE_PRODUCTS_OUTPUT_SHEET_NAME
  ];

  for (const placementTypeName of placementTypeNames) {
    const sheet = spreadsheet.getSheetByName(placementTypeName);
    if (!sheet) {
      debugLog(`  Skipping ${placementTypeName} sheet as it does not exist.`)
      continue;
    }

    const values = sheet.getDataRange().getValues();
    // Headers are at row 3 (index 2), data starts at row 4 (index 3)
    if (values.length <= 3) {
      debugLog(`  No data in ${placementTypeName} sheet. Nothing to read.`)
      continue;
    }

    // Find column indices dynamically - headers are at row 3 (index 2)
    const headers = values[2];
    const checkboxIndex = headers.indexOf('Exclude');
    const statusIndex = headers.indexOf('Status');
    const campaignIdIndex = headers.indexOf('Campaign ID');
    const campaignNameIndex = headers.indexOf('Campaign Name');
    const placementIndex = headers.indexOf('Placement');
    const displayNameIndex = headers.indexOf('Display Name');
    const placementTypeIndex = headers.indexOf('Placement Type');
    const targetUrlIndex = headers.indexOf('Target URL');
    const notesIndex = headers.indexOf('Notes');

    if (checkboxIndex === -1 || placementIndex === -1 || placementTypeIndex === -1 || targetUrlIndex === -1) {
      console.warn(`  Skipping ${placementTypeName} sheet: Missing one or more required columns (Exclude, Placement, Placement Type, Target URL).`);
      continue;
    }

    let checkedCount = 0;
    let uncheckedCount = 0;

    // Data starts at row 4 (index 3)
    for (let rowIndex = 3; rowIndex < values.length; rowIndex++) {
      const row = values[rowIndex];
      const checkboxValue = row[checkboxIndex];
      const isChecked = checkboxValue === true || String(checkboxValue).toUpperCase() === 'TRUE';
      const placementValue = String(row[placementIndex] || '').trim();
      const placementType = String(row[placementTypeIndex] || '').toUpperCase();
      const targetUrl = String(row[targetUrlIndex] || '').trim();

      // Read additional columns for logging
      const status = statusIndex !== -1 ? String(row[statusIndex] || '').trim() : '';
      const campaignId = campaignIdIndex !== -1 ? String(row[campaignIdIndex] || '').trim() : '';
      const campaignName = campaignNameIndex !== -1 ? String(row[campaignNameIndex] || '').trim() : '';
      const displayName = displayNameIndex !== -1 ? String(row[displayNameIndex] || '').trim() : '';
      const notes = notesIndex !== -1 ? String(row[notesIndex] || '').trim() : '';

      if (DEBUG_MODE && rowIndex <= 3) {
        console.log(`  Reading from ${placementTypeName} - Row ${rowIndex + 1}: checkbox=${checkboxValue} (type: ${typeof checkboxValue}), placement=${placementValue}, type=${placementType}, targetUrl=${targetUrl}`);
      }

      if (isChecked) {
        checkedCount++;
      } else {
        uncheckedCount++;
      }

      if (isChecked && placementValue) {
        const placementTypeUpper = String(placementType || '').toUpperCase();

        // Skip GOOGLE_PRODUCTS placements - they cannot be excluded (not real URLs)
        if (placementTypeUpper.includes('GOOGLE_PRODUCTS')) {
          console.warn(`⚠ Skipping GOOGLE_PRODUCTS placement (cannot be excluded via placement lists): ${placementValue}`);
          skippedInvalid++;
          continue;
        }

        const mobileAppPlacement = formatMobileAppPlacement(placementValue, placementType);

        // Collect full placement data for logging
        const placementLogData = {
          status: status,
          campaignId: campaignId,
          campaignName: campaignName,
          placement: placementValue,
          displayName: displayName,
          placementType: row[placementTypeIndex] || '', // Original value (not uppercased)
          targetUrl: targetUrl,
          notes: mobileAppPlacement ? 'Excluded via bulk upload' : notes,
          isMobileApp: !!mobileAppPlacement
        };

        if (mobileAppPlacement) {
          checkedMobileAppPlacementsSet.add(mobileAppPlacement);
          fullPlacementData.push(placementLogData);
        } else {
          const isYouTubePlacement = placementTypeUpper.includes('YOUTUBE');

          if (isYouTubePlacement) {
            // Store YouTube placements as objects with placement type, placement value, and campaign ID
            // The formatYouTubePlacementForExclusion function will handle proper URL formatting
            checkedYouTubePlacements.push({
              placement: placementValue,
              placementType: placementType,
              campaignId: campaignId,
              campaignName: campaignName
            });
            fullPlacementData.push(placementLogData);
          } else {
            // It's a website URL, normalize it
            const normalizedUrl = normalizeUrl(placementValue);

            if (normalizedUrl === null) {
              skippedInvalid++;
            } else {
              checkedWebsitePlacementsSet.add(normalizedUrl);
              fullPlacementData.push(placementLogData);
            }
          }
        }
      }
    }
    debugLog(`  ${placementTypeName} sheet: Checked=${checkedCount}, Unchecked=${uncheckedCount}`)
    totalCheckedCount += checkedCount;
    totalUncheckedCount += uncheckedCount;
  }

  const websitePlacements = Array.from(checkedWebsitePlacementsSet);
  const youtubePlacements = checkedYouTubePlacements;
  const mobileAppPlacements = Array.from(checkedMobileAppPlacementsSet);
  const checkedPlacements = [...websitePlacements, ...youtubePlacements.map(ytp => ytp.placement), ...mobileAppPlacements];

  debugLog(`Total checked boxes found across all sheets: ${totalCheckedCount}, Total unchecked: ${totalUncheckedCount}`)
  debugLog(`Valid placements to exclude: ${checkedPlacements.length}`)

  if (checkedPlacements.length > 0) {
    console.log(`Found ${checkedPlacements.length} unique checked placements to exclude`);
    if (websitePlacements.length > 0) {
      console.log(`  - Website placements: ${websitePlacements.length}`);
    }
    if (youtubePlacements.length > 0) {
      console.log(`  - YouTube placements: ${youtubePlacements.length}`);
    }
    if (mobileAppPlacements.length > 0) {
      console.log(`  - Mobile app placements: ${mobileAppPlacements.length}`);
    }
    if (DEBUG_MODE && websitePlacements.length > 0) {
      debugLog(`Website placements:`)
      websitePlacements.slice(0, 3).forEach(p => debugLog(`  - ${p}`))
      if (websitePlacements.length > 3) debugLog(`  ... and ${websitePlacements.length - 3} more`)
    }
    if (DEBUG_MODE && youtubePlacements.length > 0) {
      debugLog(`YouTube placements:`)
      youtubePlacements.slice(0, 3).forEach(ytp => debugLog(`  - ${ytp.placement} (${ytp.placementType})`))
      if (youtubePlacements.length > 3) debugLog(`  ... and ${youtubePlacements.length - 3} more`)
    }
    if (DEBUG_MODE && mobileAppPlacements.length > 0) {
      debugLog(`Mobile app placements:`)
      mobileAppPlacements.slice(0, 3).forEach(p => debugLog(`  - ${p}`))
      if (mobileAppPlacements.length > 3) debugLog(`  ... and ${mobileAppPlacements.length - 3} more`)
    }
    if (skippedInvalid > 0) {
      console.log(`⚠ Skipped ${skippedInvalid} invalid placement URL(s)`);
    }
    console.log(``);
  } else {
    if (skippedInvalid > 0) {
      console.log(`No valid placements to exclude.`);
      if (skippedInvalid > 0) {
        console.log(`⚠ Skipped ${skippedInvalid} invalid placement URL(s)`);
      }
    }
  }

  return {
    websitePlacements: websitePlacements,
    youtubeVideos: youtubePlacements, // Array of objects with { placement, placementType }
    mobileAppPlacements: mobileAppPlacements,
    fullPlacementData: fullPlacementData
  };
}

/**
 * Checks if the script is running in preview mode
 * @returns {boolean} True if in preview mode, false otherwise
 */
function isPreviewMode() {
  try {
    return AdsApp.getExecutionInfo().isPreview();
  } catch (error) {
    console.warn(`Could not determine preview mode: ${error.message}`);
    return false;
  }
}

/**
 * Logs exclusions to the Exclusion Log sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @param {Array<Object>} placementData - Array of placement objects with full data for logging
 */
function logExclusionsToSheet(spreadsheet, placementData) {
  if (!placementData || placementData.length === 0) {
    debugLog(`No exclusions to log.`)
    return;
  }

  const exclusionLogSheet = spreadsheet.getSheetByName(EXCLUSION_LOG_SHEET_NAME);
  if (!exclusionLogSheet) {
    console.warn(`⚠ Exclusion Log sheet not found. Skipping exclusion logging.`);
    console.warn(`  Create a sheet named "${EXCLUSION_LOG_SHEET_NAME}" with headers: Status, Campaign Name, Placement, Display Name, Placement Type, Target URL, Notes, Preview Mode, Timestamp`);
    return;
  }

  const previewMode = isPreviewMode();
  const timestamp = new Date();

  // Build rows to append
  // Headers expected: Status, Campaign Name, Placement, Display Name, Placement Type, Target URL, Notes, Preview Mode, Timestamp
  const rowsToAppend = placementData.map(placement => {
    return [
      placement.status || '',
      placement.campaignName || '',
      placement.placement || '',
      placement.displayName || '',
      placement.placementType || '',
      placement.targetUrl || '',
      placement.notes || '',
      previewMode ? 'Yes' : 'No',
      timestamp
    ];
  });

  // Find the last row with data in column A (more reliable than getLastRow which checks entire sheet)
  const lastRowInSheet = exclusionLogSheet.getLastRow();
  let lastRowWithDataInColumnA = 0;

  if (lastRowInSheet > 0) {
    const columnAValues = exclusionLogSheet.getRange(1, 1, lastRowInSheet, 1).getValues();
    for (let rowIndex = columnAValues.length - 1; rowIndex >= 0; rowIndex--) {
      if (columnAValues[rowIndex][0] !== '') {
        lastRowWithDataInColumnA = rowIndex + 1; // Convert to 1-based row number
        break;
      }
    }
  }

  // If no data found (empty sheet), start after assumed header row
  const startRow = lastRowWithDataInColumnA > 0 ? lastRowWithDataInColumnA + 1 : 2;

  // Append all rows at once
  if (rowsToAppend.length > 0) {
    exclusionLogSheet.getRange(startRow, 1, rowsToAppend.length, 9).setValues(rowsToAppend);
    console.log(`✓ Logged ${rowsToAppend.length} exclusions to ${EXCLUSION_LOG_SHEET_NAME} sheet (starting at row ${startRow})`);
    if (previewMode) {
      console.log(`  (Script was in Preview Mode)`);
    }
  }
}

/**
 * Adds mobile app exclusions to a Shared Placement Exclusion List using Bulk Upload service
 * Utilizes the Ads Scripts Bulk Upload service for criterion-aware processing
 * @param {string} listName - The name of the target Shared Placement Exclusion List
 * @param {Array<string>} appIds - Array of canonical App IDs (e.g., ["1-1510189987", "2-com.bad.game"])
 */
function addAppExclusionsViaBulkUpload(listName, appIds) {
  if (!appIds || appIds.length === 0) {
    console.log('No mobile app IDs provided for bulk upload.');
    return;
  }

  // Define the necessary CSV column headers based on the Google Ads Bulk Schema
  const columns = [
    'Type',
    'Status',
    'Placement Exclusion List Name',
    'Shared Set Type',
    'Placement Type',
    'Placement url'
  ];

  console.log(`Preparing to bulk upload ${appIds.length} mobile app exclusions to list: ${listName}`);

  try {
    // Initialize the CSV upload object, which acts as an incremental payload builder
    const bulkUpload = AdsApp.bulkUploads().newCsvUpload(columns);
    bulkUpload.forCampaignManagement(); // Specify that this is for entity changes

    // Loop through the App IDs and construct the payload rows
    appIds.forEach(appId => {
      // Construct the row object, ensuring explicit type definition
      const row = {
        'Type': 'Placement Criterion',
        'Status': 'Active',
        'Placement Exclusion List Name': listName,
        'Shared Set Type': 'PLACEMENT_EXCLUSION', // Required for proper criterion linkage
        'Placement Type': 'Mobile Application', // Declares the criterion type
        'Placement url': appId // The canonical App ID (e.g., 1-1510189987)
      };

      bulkUpload.append(row);

      debugLog(`Row appended for App ID: ${appId}`)
    });

    // Execution Phase: Check if we're in preview mode or execution mode
    if (isPreviewMode()) {
      bulkUpload.preview();
      console.log('✓ Bulk upload job submitted for PREVIEW. Check Tools & Settings > Bulk actions > Uploads for results.');
    } else {
      bulkUpload.apply();
      console.log('✓ Bulk upload job APPLIED. Check Tools & Settings > Bulk actions > Uploads for results.');
    }
  } catch (error) {
    console.error(`✗ Error during bulk upload of mobile app exclusions: ${error.message}`);
    throw error;
  }
}

/**
 * Adds YouTube placements (channels and videos) as exclusions to Video campaigns
 * Note: Shared placement exclusion lists do NOT support YouTube URLs.
 * YouTube exclusions must be added at the campaign level using videoTargeting().
 * @param {Array<Object>} youtubePlacements - Array of YouTube placement objects with { placement, placementType }
 * @param {string} listName - The name of the shared exclusion list (not used for YouTube, kept for API compatibility)
 */
function addYouTubeVideosToExclusionList(youtubePlacements, listName) {
  if (youtubePlacements.length === 0) {
    return;
  }

  // Group video IDs by campaign ID (each placement should be excluded from its parent campaign only)
  const videosByCampaignId = new Map();
  const skippedChannels = [];
  const missingCampaignIdPlacements = [];

  for (const ytPlacement of youtubePlacements) {
    const placementType = String(ytPlacement.placementType || '').toUpperCase();
    const placement = String(ytPlacement.placement || '').trim();
    const campaignId = String(ytPlacement.campaignId || '').trim();
    const campaignName = String(ytPlacement.campaignName || '').trim();

    if (!placement) {
      continue;
    }

    // Skip placements without campaign ID - can't target specific campaign
    if (!campaignId) {
      missingCampaignIdPlacements.push({ placement, campaignName });
      continue;
    }

    // Only process YOUTUBE_VIDEO placements - channels need different handling
    if (placementType.includes('VIDEO')) {
      // Extract video ID if it's a URL, otherwise use as-is
      let videoId = placement;
      if (placement.includes('youtube.com/watch?v=')) {
        const match = placement.match(/[?&]v=([a-zA-Z0-9_-]+)/);
        if (match) {
          videoId = match[1];
        }
      } else if (placement.includes('youtu.be/')) {
        const match = placement.match(/youtu\.be\/([a-zA-Z0-9_-]+)/);
        if (match) {
          videoId = match[1];
        }
      }

      // Add to campaign-specific set
      if (!videosByCampaignId.has(campaignId)) {
        videosByCampaignId.set(campaignId, { videoIds: new Set(), campaignName: campaignName });
      }
      videosByCampaignId.get(campaignId).videoIds.add(videoId);
    } else if (placementType.includes('CHANNEL')) {
      skippedChannels.push(placement);
    } else {
      // Generic YOUTUBE type - assume video if it looks like a video ID
      if (!placement.startsWith('UC') && !placement.startsWith('@')) {
        if (!videosByCampaignId.has(campaignId)) {
          videosByCampaignId.set(campaignId, { videoIds: new Set(), campaignName: campaignName });
        }
        videosByCampaignId.get(campaignId).videoIds.add(placement);
      } else {
        skippedChannels.push(placement);
      }
    }
  }

  if (missingCampaignIdPlacements.length > 0) {
    console.warn(`⚠ Skipping ${missingCampaignIdPlacements.length} YouTube placement(s) - missing Campaign ID in sheet`);
    if (DEBUG_MODE) {
      missingCampaignIdPlacements.slice(0, 3).forEach(p => console.warn(`  - ${p.placement} (campaign: ${p.campaignName || 'unknown'})`));
      if (missingCampaignIdPlacements.length > 3) {
        console.warn(`  ... and ${missingCampaignIdPlacements.length - 3} more`);
      }
    }
  }

  if (skippedChannels.length > 0) {
    console.warn(`⚠ Skipping ${skippedChannels.length} YouTube channel(s) - channel exclusions via scripts not fully supported`);
    if (DEBUG_MODE) {
      skippedChannels.slice(0, 3).forEach(ch => console.warn(`  - ${ch}`));
      if (skippedChannels.length > 3) {
        console.warn(`  ... and ${skippedChannels.length - 3} more`);
      }
    }
  }

  if (videosByCampaignId.size === 0) {
    console.log(`No YouTube videos to exclude (only channels found or missing campaign IDs).`);
    return;
  }

  // Count total unique videos
  let totalVideoCount = 0;
  for (const [campaignId, data] of videosByCampaignId) {
    totalVideoCount += data.videoIds.size;
  }

  console.log(`Attempting to exclude ${totalVideoCount} YouTube video(s) across ${videosByCampaignId.size} campaign(s)...`);

  let totalSuccessCount = 0;
  let totalFailureCount = 0;
  let campaignNotFoundCount = 0;
  const failedPlacements = [];

  // Process each campaign's exclusions
  for (const [campaignId, data] of videosByCampaignId) {
    const { videoIds, campaignName } = data;

    // Find the specific Video campaign by ID
    const campaign = findVideoCampaignById(campaignId);

    if (!campaign) {
      console.warn(`⚠ Video campaign not found for ID ${campaignId} (${campaignName}). Skipping ${videoIds.size} video(s).`);
      campaignNotFoundCount++;
      continue;
    }

    let campaignSuccessCount = 0;
    let campaignFailureCount = 0;

    for (const videoId of videoIds) {
      try {
        const operation = campaign.videoTargeting()
          .newYouTubeVideoBuilder()
          .withVideoId(videoId)
          .exclude();

        if (operation.isSuccessful()) {
          campaignSuccessCount++;
          totalSuccessCount++;
          if (DEBUG_MODE && campaignSuccessCount <= 3) {
            console.log(`✓ Excluded video ${videoId} from campaign: ${campaignName} (${campaignId})`);
          }
        } else {
          campaignFailureCount++;
          totalFailureCount++;
          const errors = operation.getErrors();
          failedPlacements.push({ videoId: videoId, campaign: campaignName, campaignId: campaignId, error: errors.join(', ') });
          if (DEBUG_MODE || failedPlacements.length <= 3) {
            console.error(`✗ Failed to exclude video ${videoId} from ${campaignName}: ${errors.join(', ')}`);
          }
        }
      } catch (error) {
        campaignFailureCount++;
        totalFailureCount++;
        failedPlacements.push({ videoId: videoId, campaign: campaignName, campaignId: campaignId, error: error.message });
        if (DEBUG_MODE || failedPlacements.length <= 3) {
          console.error(`✗ Error excluding video ${videoId} from ${campaignName}: ${error.message}`);
        }
      }
    }

    debugLog(`  Campaign "${campaignName}" (${campaignId}): ${campaignSuccessCount} succeeded, ${campaignFailureCount} failed`);
  }

  console.log(`\n=== YouTube Video Exclusion Summary ===`);
  console.log(`Video campaigns processed: ${videosByCampaignId.size - campaignNotFoundCount}`);
  if (campaignNotFoundCount > 0) {
    console.warn(`Video campaigns not found: ${campaignNotFoundCount}`);
  }
  console.log(`Total exclusions added: ${totalSuccessCount}`);
  if (totalFailureCount > 0) {
    console.warn(`Total failures: ${totalFailureCount}`);
    if (failedPlacements.length <= 10) {
      console.warn(`\nFailed exclusions:`);
      for (const failed of failedPlacements) {
        console.warn(`  - Video ${failed.videoId} on ${failed.campaign} (${failed.campaignId}): ${failed.error}`);
      }
    } else {
      console.warn(`\nFirst 10 failed exclusions:`);
      for (let failedIndex = 0; failedIndex < 10; failedIndex++) {
        const failed = failedPlacements[failedIndex];
        console.warn(`  - Video ${failed.videoId} on ${failed.campaign} (${failed.campaignId}): ${failed.error}`);
      }
      console.warn(`  ... and ${failedPlacements.length - 10} more`);
    }
  }
  console.log(`========================================\n`);
}

/**
 * Finds a Video campaign by its ID
 * @param {string} campaignId - The campaign ID to find
 * @returns {Object|null} The campaign object, or null if not found
 */
function findVideoCampaignById(campaignId) {
  if (!campaignId) {
    return null;
  }

  // Search in Video campaigns
  const videoCampaignIterator = AdsApp.videoCampaigns()
    .withIds([campaignId])
    .get();

  if (videoCampaignIterator.hasNext()) {
    return videoCampaignIterator.next();
  }

  return null;
}

/**
 * Formats a YouTube placement for the exclusion list API
 * Channels: youtube.com/channel/UCxxxxxxx
 * Videos: youtube.com/watch?v=xxxxxxx
 * @param {Object} ytPlacement - Object with { placement, placementType }
 * @returns {string|null} Formatted URL for the exclusion list, or null if invalid
 */
function formatYouTubePlacementForExclusion(ytPlacement) {
  const placement = String(ytPlacement.placement || '').trim();
  const placementType = String(ytPlacement.placementType || '').toUpperCase();

  if (!placement) {
    return null;
  }

  // If it's already a full YouTube URL, use it as-is
  if (placement.includes('youtube.com/') || placement.includes('youtu.be/')) {
    return placement;
  }

  // Determine if this is a channel or video based on placement type
  const isChannel = placementType.includes('CHANNEL');
  const isVideo = placementType.includes('VIDEO');

  // If placement type is YOUTUBE_CHANNEL, format as channel URL
  if (isChannel) {
    // Placement might be a channel ID like "UCxxxxxxx" or a handle like "@channelname"
    if (placement.startsWith('UC') || placement.startsWith('@')) {
      return `youtube.com/channel/${placement}`;
    }
    // Otherwise assume it's already in the right format
    return `youtube.com/channel/${placement}`;
  }

  // If placement type is YOUTUBE_VIDEO, format as video URL
  if (isVideo) {
    // Placement is likely a video ID like "K2zZwB3NY04"
    return `youtube.com/watch?v=${placement}`;
  }

  // Generic YOUTUBE type - try to determine from placement format
  if (placement.startsWith('UC')) {
    // Looks like a channel ID
    return `youtube.com/channel/${placement}`;
  }

  // Assume it's a video ID
  return `youtube.com/watch?v=${placement}`;
}

/**
 * Adds website placements to the shared exclusion list
 * @param {Array<string>} placementUrls - Array of placement URLs to exclude
 * @param {string} listName - The name of the shared exclusion list
 */
function addPlacementsToExclusionList(placementUrls, listName) {
  if (placementUrls.length === 0) {
    return;
  }

  const listIterator = AdsApp.excludedPlacementLists()
    .withCondition(`Name = '${listName}'`)
    .get();

  if (!listIterator.hasNext()) {
    throw new Error(`Shared exclusion list '${listName}' not found`);
  }

  const excludedPlacementList = listIterator.next();

  // Use Set to ensure uniqueness
  const uniquePlacements = Array.from(new Set(placementUrls));

  console.log(`Attempting to add ${uniquePlacements.length} unique placements to the shared list (batch processing)...`);

  try {
    // Batch API call for efficiency (supports up to 20,000 placements at a time)
    excludedPlacementList.addExcludedPlacements(uniquePlacements);
    console.log(`✓ Successfully added ${uniquePlacements.length} placements to the list using batch processing`);
  } catch (error) {
    console.error(`✗ Error during batch placement upload: ${error.message}`);
    console.log(`\nAttempting to add placements individually as fallback...`);

    // Fallback to individual additions if batch fails
    let successCount = 0;
    let failureCount = 0;
    const failedPlacements = [];

    for (const placementUrl of uniquePlacements) {
      try {
        excludedPlacementList.addExcludedPlacement(placementUrl);
        successCount++;
        if (DEBUG_MODE && successCount <= 3) {
          console.log(`✓ Added placement: ${placementUrl}`);
        }
      } catch (individualError) {
        failureCount++;
        failedPlacements.push({ placement: placementUrl, error: individualError.message });
        if (DEBUG_MODE || failureCount <= 3) {
          console.error(`✗ Failed to add placement ${placementUrl}: ${individualError.message}`);
        }
      }
    }

    console.log(`\n=== Exclusion List Update Summary ===`);
    console.log(`Successfully added: ${successCount} placements`);
    if (failureCount > 0) {
      console.warn(`Failed to add: ${failureCount} placements`);
      if (failedPlacements.length <= 10) {
        console.warn(`\nFailed placements:`);
        for (const failed of failedPlacements) {
          console.warn(`  - ${failed.placement}: ${failed.error}`);
        }
      } else {
        console.warn(`\nFirst 10 failed placements:`);
        for (let i = 0; i < 10; i++) {
          const failed = failedPlacements[i];
          console.warn(`  - ${failed.placement}: ${failed.error}`);
        }
        console.warn(`  ... and ${failedPlacements.length - 10} more`);
      }
    }
    console.log(`=====================================\n`);
  }
}

/**
 * Links the shared exclusion list to Performance Max and Display campaigns
 * This is mandatory for campaigns to honor the negative placements
 * @param {string} listName - The name of the shared exclusion list
 */
function linkSharedListToPMaxCampaigns(listName) {
  const listIterator = AdsApp.excludedPlacementLists()
    .withCondition(`Name = '${listName}'`)
    .get();

  if (!listIterator.hasNext()) {
    console.error(`Shared exclusion list '${listName}' not found. Cannot link to campaigns.`);
    return;
  }

  const excludedPlacementList = listIterator.next();

  // Link to Performance Max campaigns
  const pMaxCampaignIterator = AdsApp.performanceMaxCampaigns()
    .withCondition('campaign.status = ENABLED')
    .get();

  let campaignsUpdated = 0;
  let campaignsAlreadyLinked = 0;
  let campaignsFailed = 0;

  while (pMaxCampaignIterator.hasNext()) {
    const campaign = pMaxCampaignIterator.next();
    const campaignId = campaign.getId();
    const campaignName = campaign.getName();

    try {
      // Retrieve currently applied lists to avoid redundant mutation calls
      const appliedLists = campaign.getExcludedPlacementLists();
      let isListApplied = false;

      if (appliedLists) {
        const appliedListsIterator = appliedLists.get();
        while (appliedListsIterator.hasNext()) {
          const appliedList = appliedListsIterator.next();
          if (appliedList.getName() === listName) {
            isListApplied = true;
            break;
          }
        }
      }

      // Apply the list if it is not already attached
      if (!isListApplied) {
        campaign.addExcludedPlacementList(excludedPlacementList);
        campaignsUpdated++;
        debugLog(`✓ Linked exclusion list to campaign: ${campaignName} (ID: ${campaignId})`)
      } else {
        campaignsAlreadyLinked++;
        debugLog(`- Campaign already linked: ${campaignName} (ID: ${campaignId})`)
      }
    } catch (error) {
      campaignsFailed++;
      console.error(`✗ Failed to link exclusion list to Campaign ID ${campaignId} (${campaignName}): ${error.message}`);
    }
  }

  console.log(`\n=== Campaign Linking Summary (Performance Max) ===`);
  console.log(`Newly linked: ${campaignsUpdated} campaigns`);
  console.log(`Already linked: ${campaignsAlreadyLinked} campaigns`);
  if (campaignsFailed > 0) {
    console.warn(`Failed to link: ${campaignsFailed} campaigns`);
  }
  console.log(`=================================\n`);

  // Link to Display campaigns
  const displayCampaignIterator = AdsApp.campaigns()
    .withCondition('AdvertisingChannelType = DISPLAY')
    .withCondition('Status = ENABLED')
    .get();

  let displayCampaignsUpdated = 0;
  let displayCampaignsAlreadyLinked = 0;
  let displayCampaignsFailed = 0;

  while (displayCampaignIterator.hasNext()) {
    const campaign = displayCampaignIterator.next();
    const campaignId = campaign.getId();
    const campaignName = campaign.getName();

    try {
      // Retrieve currently applied lists to avoid redundant mutation calls
      const appliedLists = campaign.getExcludedPlacementLists();
      let isListApplied = false;

      while (appliedLists.hasNext()) {
        const appliedList = appliedLists.next();
        if (appliedList.getId() === excludedPlacementList.getId()) {
          isListApplied = true;
          displayCampaignsAlreadyLinked++;
          break;
        }
      }

      if (!isListApplied) {
        campaign.addExcludedPlacementList(excludedPlacementList);
        displayCampaignsUpdated++;
        if (DEBUG_MODE && displayCampaignsUpdated <= 3) {
          console.log(`Linked exclusion list to Display campaign: ${campaignName} (ID: ${campaignId})`);
        }
      }
    } catch (error) {
      displayCampaignsFailed++;
      if (DEBUG_MODE || displayCampaignsFailed <= 3) {
        console.error(`Failed to link exclusion list to Display campaign ${campaignName} (ID: ${campaignId}): ${error.message}`);
      }
    }
  }

  console.log(`\n=== Campaign Linking Summary (Display) ===`);
  console.log(`Newly linked: ${displayCampaignsUpdated} campaigns`);
  console.log(`Already linked: ${displayCampaignsAlreadyLinked} campaigns`);
  if (displayCampaignsFailed > 0) {
    console.warn(`Failed to link: ${displayCampaignsFailed} campaigns`);
  }
  console.log(`=================================\n`);
}

/**
 * Formats a mobile app placement for exclusion
 * @param {string} placement - The placement ID (e.g., "1-1510189987" for iOS or "2-com.example.app" for Android)
 * @param {string} placementType - The placement type from the report
 * @returns {string | null} Formatted mobile app placement or null if invalid/not a mobile app
 */
function formatMobileAppPlacement(placement, placementType) {
  if (!placement || typeof placement !== 'string') {
    return null;
  }

  const placementStr = String(placement).trim();
  const typeStr = placementType ? String(placementType).toUpperCase() : '';

  // Check if it's a mobile app based on placement type or format
  const isMobileApp = typeStr.includes('MOBILE_APPLI') ||
    /^[12]-\d+$/.test(placementStr) || // iOS format: 1-XXXXX or Android format: 2-XXXXX
    /^[12]-[a-z0-9.]+$/i.test(placementStr); // Android package format: 2-com.example.app

  if (!isMobileApp) {
    return null; // Not a mobile app, return null to indicate it should be processed as URL
  }

  // Detect iOS vs Android
  // iOS: typically starts with "1-" followed by numbers (iTunes ID)
  // Android: typically starts with "2-" followed by package name (e.g., com.example.app) or numbers
  if (/^1-/.test(placementStr)) {
    // iOS format: mobileapp::1-[iTunes_ID]
    const itunesId = placementStr.replace(/^1-/, '');
    return `mobileapp::1-${itunesId}`;
  } else if (/^2-/.test(placementStr)) {
    // Android format: mobileapp::2-[package_name]
    const packageName = placementStr.replace(/^2-/, '');
    return `mobileapp::2-${packageName}`;
  }

  return null;
}

/**
 * Normalizes and validates a placement URL for use in exclusion lists
 * Ensures the URL is lowercase, stripped of protocols and trailing slashes,
 * and not a known Google-owned domain (which cannot be excluded)
 * @param {string} url - The raw placement URL from the report
 * @returns {string | null} The normalized URL, or null if it fails validation
 */
function normalizeUrl(url) {
  if (!url || typeof url !== 'string') {
    return null;
  }

  // Convert to lowercase
  let formattedUrl = url.toLowerCase();

  // Remove protocol prefix (http:// or https://)
  formattedUrl = formattedUrl.replace(/^https?:\/\//i, '');

  // Remove 'www.' for subdomain wildcard exclusion
  formattedUrl = formattedUrl.replace(/^www\./, '');

  // Remove trailing slash
  if (formattedUrl.endsWith('/')) {
    formattedUrl = formattedUrl.slice(0, -1);
  }

  // Note: Google domains are no longer filtered out here - they will appear in reports
  // but cannot be excluded. The exclusion logic will skip them.

  return formattedUrl;
}

// --- ChatGPT Functions ---

/**
 * Sleeps for a specified number of milliseconds
 * @param {number} ms - Milliseconds to sleep
 */
function sleep(ms) {
  Utilities.sleep(ms);
}

/**
 * Validates if a string is a valid URL
 * @param {string} url - The URL string to validate
 * @returns {boolean} True if valid URL, false otherwise
 */
function isValidUrl(url) {
  if (!url || typeof url !== 'string') {
    return false;
  }

  const trimmedUrl = url.trim();
  if (!trimmedUrl) {
    return false;
  }

  // For placement URLs, we accept:
  // 1. Full URLs (http:// or https://)
  // 2. Domain names (e.g., express.co.uk, mirror.co.uk)
  // 3. Domains must have at least one dot and valid characters

  // Skip mobile app identifiers
  if (trimmedUrl.startsWith('mobileapp::')) {
    return false;
  }

  // If it has a protocol, try to validate with URL constructor
  if (trimmedUrl.startsWith('http://') || trimmedUrl.startsWith('https://')) {
    try {
      const urlObj = new URL(trimmedUrl);
      return urlObj.hostname && urlObj.hostname.length > 0;
    } catch (e) {
      return false;
    }
  }

  // For domain names without protocol, do simple validation
  // Must contain at least one dot and only valid domain characters
  const domainPattern = /^[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9-]*[a-zA-Z0-9])?)+$/;
  return domainPattern.test(trimmedUrl);
}

/**
 * Fetches website content with exponential backoff retry logic
 * @param {string} url - The URL to fetch
 * @param {number} maxRetries - Maximum number of retry attempts (default: 3)
 * @returns {string | null} The website content or null if all attempts failed
 */
function fetchWebsiteContentWithRetry(url, maxRetries = 3) {
  let attempt = 0;
  let baseDelay = 1000; // Start with 1 second

  while (attempt < maxRetries) {
    try {
      const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true,
        maxRedirects: 5
      });

      if (response.getResponseCode() === 200) {
        const html = response.getContentText();
        // Extract text content from HTML (basic extraction)
        const textContent = html.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
          .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
          .replace(/<[^>]+>/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();

        // Limit content length to avoid token limits
        const maxLength = 10000;
        return textContent.length > maxLength ? textContent.substring(0, maxLength) + '...' : textContent;
      } else {
        debugLog(`Attempt ${attempt + 1}: HTTP ${response.getResponseCode()} for ${url}`)
      }
    } catch (error) {
      debugLog(`Attempt ${attempt + 1} failed for ${url}: ${error.message}`)
    }

    attempt++;
    if (attempt < maxRetries) {
      const delay = baseDelay * Math.pow(2, attempt - 1); // Exponential backoff
      debugLog(`Retrying in ${delay}ms...`)
      sleep(delay);
    }
  }

  console.error(`Failed to fetch ${url} after ${maxRetries} attempts`);
  return null;
}

/**
 * Calls ChatGPT API with exponential backoff retry logic
 * @param {string} apiKey - OpenAI API key
 * @param {string} prompt - The prompt to send (includes user prompt + website content)
 * @param {number} maxRetries - Maximum number of retry attempts (default: 3)
 * @returns {string | null} The ChatGPT response or null if all attempts failed
 */
function callChatGptApiWithRetry(apiKey, prompt, maxRetries = 3) {
  // Always sleep a small amount before LLM request
  sleep(500);

  let attempt = 0;
  let baseDelay = 2000; // Start with 2 seconds

  while (attempt < maxRetries) {
    try {
      const payload = {
        model: 'gpt-3.5-turbo',
        messages: [
          {
            role: 'user',
            content: prompt
          }
        ],
        max_tokens: 500,
        temperature: 0.7
      };

      const options = {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'Authorization': `Bearer ${apiKey}`
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode === 200) {
        const responseJson = JSON.parse(responseText);
        if (responseJson.choices && responseJson.choices.length > 0) {
          return responseJson.choices[0].message.content.trim();
        }
      } else {
        const errorData = JSON.parse(responseText);
        debugLog(`Attempt ${attempt + 1}: API error ${responseCode}: ${errorData.error?.message || responseText}`)

        // Don't retry on authentication errors
        if (responseCode === 401) {
          console.error(`Authentication failed. Check your API key.`);
          return null;
        }
      }
    } catch (error) {
      debugLog(`Attempt ${attempt + 1} failed: ${error.message}`)
    }

    attempt++;
    if (attempt < maxRetries) {
      const delay = baseDelay * Math.pow(2, attempt - 1); // Exponential backoff
      debugLog(`Retrying ChatGPT API call in ${delay}ms...`)
      sleep(delay);
    }
  }

  console.error(`Failed to get ChatGPT response after ${maxRetries} attempts`);
  return null;
}

/**
 * Gets or creates the LLM responses cache sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The LLM responses sheet
 */
function getOrCreateLlmResponsesSheet(spreadsheet) {
  let llmSheet = spreadsheet.getSheetByName(LLM_RESPONSES_SHEET_NAME);
  if (!llmSheet) {
    llmSheet = spreadsheet.insertSheet(LLM_RESPONSES_SHEET_NAME);
    // Add headers
    llmSheet.getRange(1, 1, 1, 3).setValues([['URL', 'Response', 'Timestamp']]);
    llmSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    llmSheet.setFrozenRows(1);
    console.log(`Created LLM responses cache sheet: ${LLM_RESPONSES_SHEET_NAME}`);
  }
  return llmSheet;
}

/**
 * Gets all cached LLM responses as an object keyed by normalized URL
 * @param {GoogleAppsScript.Spreadsheet.Sheet} llmSheet - The LLM responses sheet
 * @returns {Object} Object where keys are normalized URLs (lowercase, trimmed) and values are responses
 */
function getCachedLlmResponses(llmSheet) {
  const cache = {};
  const dataRange = llmSheet.getDataRange();
  const values = dataRange.getValues();

  // Headers are in row 1 (index 0), so data starts from row 2 (index 1)
  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const cachedUrl = String(values[rowIndex][0] || '').trim().toLowerCase();
    const cachedResponse = values[rowIndex][1];
    if (cachedUrl) {
      cache[cachedUrl] = cachedResponse;
    }
  }

  console.log(`Loaded ${Object.keys(cache).length} cached LLM responses from sheet`)
  return cache;
}

/**
 * Caches LLM response for a URL
 * @param {GoogleAppsScript.Spreadsheet.Sheet} llmSheet - The LLM responses sheet
 * @param {string} url - The URL
 * @param {string} response - The LLM response
 */
function cacheLlmResponse(llmSheet, url, response) {
  const lastRow = llmSheet.getLastRow();
  // Headers are in row 1, so if sheet only has header, start from row 2
  const newRow = lastRow < 1 ? 2 : lastRow + 1;

  const timestamp = new Date();
  llmSheet.getRange(newRow, 1, 1, 3).setValues([[url, response, timestamp]]);
}

/**
 * Gets ChatGPT response for a URL, including placement information in JSONL format
 * @param {string} url - The URL to analyze
 * @param {Object} placement - The placement object containing campaign and placement info
 * @param {Object} settings - Settings object with ChatGPT configuration
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet for caching new responses
 * @param {Object} llmCache - Pre-loaded cache object (URL -> response)
 * @returns {Object} Object with { response: string, wasCached: boolean, status: string }
 *          status: 'success' | 'failed' | 'disabled'
 */
function getChatGptResponseForUrl(url, placement, settings, spreadsheet, llmCache) {
  if (!settings.enableChatGpt || !settings.chatGptApiKey) {
    return { response: '', wasCached: false, status: 'disabled' };
  }

  // Check cache first if enabled
  if (settings.useCachedChatGpt && llmCache) {
    const urlNormalized = String(url || '').trim().toLowerCase();
    const cachedResponse = llmCache[urlNormalized];

    // Check if the normalized URL exists as a key in the cache (even if value is empty)
    const keyExists = urlNormalized in llmCache;

    if (cachedResponse) {
      debugLog(`  Cache HIT for: ${urlNormalized} (response length: ${cachedResponse.length})`)
      return {
        response: cachedResponse,
        wasCached: true,
        status: 'success'
      };
    } else if (keyExists) {
      // Key exists but response is empty/falsy - treat as cache hit with empty response
      debugLog(`  Cache HIT for: ${urlNormalized} but response is empty/falsy`)
      return {
        response: '',
        wasCached: true,
        status: 'success'
      };
    } else {
      debugLog(`  Cache MISS for: ${urlNormalized} (cache has ${Object.keys(llmCache).length} keys)`)
    }
  } else {
    if (!settings.useCachedChatGpt) {
      debugLog(`  Cache disabled (useCachedChatGpt=false)`)
    } else if (!llmCache) {
      debugLog(`  Cache not loaded (llmCache is null/undefined)`)
    }
  }

  // Validate URL before fetching
  if (!isValidUrl(url)) {
    console.warn(`  ✗ ChatGPT: Invalid URL, skipping: ${url}`);
    return { response: '', wasCached: false, status: 'failed' };
  }

  // Fetch website content
  const fullUrl = url.startsWith('http') ? url : `https://${url}`;
  debugLog(`Fetching website content for ${fullUrl}...`)

  const websiteContent = fetchWebsiteContentWithRetry(fullUrl);

  if (!websiteContent) {
    console.warn(`  ✗ ChatGPT: Failed to fetch content for ${url}`);
    return { response: '', wasCached: false, status: 'failed' };
  }

  // Build placement information in JSONL format
  const placementInfo = {
    'Campaign Name': placement.campaignName || '',
    'Placement': placement.placement || '',
    'Display Name': placement.displayName || '',
    'Placement Type': placement.placementType || '',
    'Target URL': placement.targetUrl || url
  };
  const placementJsonl = JSON.stringify(placementInfo);

  // Build full prompt with placement info and website content
  const fullPrompt = `${settings.chatGptPrompt}\n\nPlacement Information (JSONL format):\n${placementJsonl}\n\nWebsite content:\n${websiteContent}`;

  // Call ChatGPT API
  debugLog(`Calling ChatGPT API for ${url}...`)

  const response = callChatGptApiWithRetry(settings.chatGptApiKey, fullPrompt);

  if (!response) {
    console.warn(`  ✗ ChatGPT: API call failed for ${url}`);
    return { response: '', wasCached: false, status: 'failed' };
  }

  // Cache the response
  const llmSheet = getOrCreateLlmResponsesSheet(spreadsheet);
  cacheLlmResponse(llmSheet, url, response);

  // Return full response - filtering happens later to exclude entire rows
  return {
    response: response,
    wasCached: false,
    status: 'success'
  };
}

/**
 * Checks if a ChatGPT response passes the Response Contains filter
 * Used to determine if the entire row should appear in output
 * @param {string} response - The ChatGPT response to check
 * @param {Object} settings - Settings object with response filter configuration
 * @returns {boolean} True if response passes filter (or no filter enabled), false if should be excluded
 */
function responsePassesFilter(response, settings) {
  if (!response || typeof response !== 'string') {
    return true; // No response = pass (don't filter out rows without responses)
  }

  const responseLower = response.toLowerCase();

  // Check Response Contains filter (only if list has items)
  if (settings.responseContainsList && settings.responseContainsList.length > 0) {
    const containsMatch = settings.responseContainsList.some(filterString => {
      return responseLower.includes(filterString.toLowerCase());
    });

    if (!containsMatch) {
      debugLog(`  Row excluded: response does not contain any of [${settings.responseContainsList.join(', ')}]`)
      return false; // Row should be excluded from output
    }
  }

  return true; // Passes filter
}


