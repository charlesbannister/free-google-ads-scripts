/**
 * Placement Exclusions - Semi-automated (Shared List)
 * For PMax and Display campaigns with optional ChatGPT integration
 * @author Charles Bannister (https://www.linkedin.com/in/charles-bannister/)
 * More scripts at shabba.io
 * 
 * This script grabs placement data from Performance Max and Display campaigns, writes to sheet with checkboxes for user selection,
 * then adds selected placements to a shared exclusion list using batch processing
 * Includes optional ChatGPT integration for website content analysis
 * Version: 2.4.1
 */

// Google Ads API Query Builder Links:
// Performance Max Placement View: https://developers.google.com/google-ads/api/fields/v20/performance_max_placement_view_query_builder
// Display Placement View: https://developers.google.com/google-ads/api/fields/v20/detail_placement_view_query_builder

//Installation Instructions:
// 1. Open the script editor by going to Bulk actions > Scripts
// 2. Click on the big plus button and name the script "PMax Placement Exclusions"
// 3. Create your Google Sheet by typing "sheets.new" in the URL bar and name the sheet "PMax Placement Exclusions"
// 4. Paste the URL of your new sheet into the SPREADSHEET_URL variable below
// 5. Paste the entire script into Google Ads (you can delete what's there)
// 6. Preview the script! (You'll be prompted to authorise the first time it runs)

// TODO: include a link to the template sheet

// --- Configuration ---

const SPREADSHEET_URL = 'YOUR_SPREADSHEET_URL_HERE';
// The Google Sheet URL where placement data will be written
// Click the link to make a copy:
// https://docs.google.com/spreadsheets/d/18vdejatcc7b3cWLtNGdmpVFoSDQ4JXXvDiVDLRpAUb4/copy
// Then paste the URL into the SPREADSHEET_URL variable above

// Create this at: Tools & Settings > Shared Library > Placement exclusions

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

const PLACEMENT_FILTERS_CELL_REFERENCES = {
  placementContains: { row: 7, column: 1 }, // checkbox
  placementNotContains: { row: 7, column: 2 }, // checkbox
  displayNameContains: { row: 7, column: 3 }, // checkbox
  displayNameNotContains: { row: 7, column: 4 }, // checkbox
  targetUrlContains: { row: 7, column: 5 }, // checkbox
  targetUrlNotContains: { row: 7, column: 6 }, // checkbox
  targetUrlEndsWith: { row: 7, column: 7 }, // checkbox
  targetUrlNotEndsWith: { row: 7, column: 8 } // checkbox
};

const PLACEMENT_FILTERS_LIST_START_ROWS = {
  placementContains: 9, // Column A
  placementNotContains: 9, // Column B
  displayNameContains: 9, // Column C
  displayNameNotContains: 9, // Column D
  targetUrlContains: 9, // Column E
  targetUrlNotContains: 9, // Column F
  targetUrlEndsWith: 9, // Column G
  targetUrlNotEndsWith: 9 // Column H
};

const DEBUG_MODE = false;
// Set to true to see detailed logs for debugging
// Core logs will always appear regardless of this setting

const ONLY_PROCESS_CHANGES = false;
// If true, skip fetching new placement data and only process checked placements from the sheets
// If false, fetch new placement data and update the sheets as normal


// --- Main Function ---
function main() {
  console.log(`Script started`);

  if (ONLY_PROCESS_CHANGES) {
    console.log(`\n⚠️ ========================================== ⚠️`);
    console.log(`⚠️  WARNING: ONLY_PROCESS_CHANGES IS TRUE     ⚠️`);
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

  // If ONLY_PROCESS_CHANGES is true, skip fetching new data and only process checked placements
  if (ONLY_PROCESS_CHANGES) {
    console.log(`Only processing changes mode enabled - skipping data fetch and sheet updates`);

    const hasWebsitePlacements = checkedPlacements.websitePlacements && checkedPlacements.websitePlacements.length > 0;
    const hasYouTubeVideos = checkedPlacements.youtubeVideos && checkedPlacements.youtubeVideos.length > 0;
    const hasMobileAppPlacements = checkedPlacements.mobileAppPlacements && checkedPlacements.mobileAppPlacements.length > 0;

    if (hasWebsitePlacements || hasYouTubeVideos || hasMobileAppPlacements) {
      // Add website placements using the standard method
      if (hasWebsitePlacements) {
        addPlacementsToExclusionList(checkedPlacements.websitePlacements, settings.sharedExclusionListName);
      }

      // Add YouTube videos individually (required for YouTube videos)
      if (hasYouTubeVideos) {
        addYouTubeVideosToExclusionList(checkedPlacements.youtubeVideos, settings.sharedExclusionListName);
      }

      // Add mobile app placements using Bulk Upload
      if (hasMobileAppPlacements) {
        addAppExclusionsViaBulkUpload(settings.sharedExclusionListName, checkedPlacements.mobileAppPlacements);
      }

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
        let processedCount = 0;
        let cachedCount = 0;
        let failedCount = 0;

        const llmSheet = getOrCreateLlmResponsesSheet(spreadsheet);

        for (const placement of limitedData) {
          // Skip ChatGPT for mobile applications and Google Products
          const placementType = String(placement.placementType || '').toUpperCase();
          if (placementType.includes('MOBILE_APPLI') || placementType.includes('GOOGLE_PRODUCTS')) {
            placement.chatGptResponse = '';
            continue;
          }

          const url = placement.targetUrl || placement.placement;
          console.log(`  Processing ChatGPT for: ${url}`);

          // Validate URL before attempting to fetch
          if (!isValidUrl(url)) {
            console.log(`  ⚠ Skipping ChatGPT for invalid URL: ${url}`);
            placement.chatGptResponse = '';
            failedCount++;
            continue;
          }

          // Check if response is already cached
          const wasCachedBefore = getCachedLlmResponse(llmSheet, url);

          try {
            const chatGptResponse = getChatGptResponseForUrl(url, placement, settings, spreadsheet);

            if (chatGptResponse) {
              placement.chatGptResponse = chatGptResponse;
              processedCount++;
              console.log(`  ✓ Got ChatGPT response for: ${url}`);

              if (wasCachedBefore) {
                cachedCount++;
              }
            } else {
              placement.chatGptResponse = ''; // Failed to get response
              failedCount++;
              console.log(`  ✗ No ChatGPT response returned for: ${url}`);
            }
          } catch (chatGptError) {
            console.error(`  ✗ ChatGPT error for ${url}: ${chatGptError.message}`);
            placement.chatGptResponse = '';
            failedCount++;
          }
        }

        console.log(`ChatGPT processing complete: ${processedCount} responses (${cachedCount} from cache, ${processedCount - cachedCount} new), ${failedCount} failed`);
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
      // Add website placements using the standard method
      if (hasWebsitePlacements) {
        addPlacementsToExclusionList(checkedPlacements.websitePlacements, settings.sharedExclusionListName);
      }

      // Add YouTube videos individually (required for YouTube videos)
      if (hasYouTubeVideos) {
        addYouTubeVideosToExclusionList(checkedPlacements.youtubeVideos, settings.sharedExclusionListName);
      }

      // Add mobile app placements using Bulk Upload
      if (hasMobileAppPlacements) {
        addAppExclusionsViaBulkUpload(settings.sharedExclusionListName, checkedPlacements.mobileAppPlacements);
      }

      linkSharedListToPMaxCampaigns(settings.sharedExclusionListName);
    } else {
      console.log(`No placements selected for exclusion. Check boxes in column A of the placement type sheets to exclude placements, then run the script again.`);
      // Still ensure the list is linked to campaigns even if no new placements added
      linkSharedListToPMaxCampaigns(settings.sharedExclusionListName);
    }
  }

  if (ONLY_PROCESS_CHANGES) {
    console.log(`\n⚠️ ========================================== ⚠️`);
    console.log(`⚠️  REMINDER: ONLY_PROCESS_CHANGES WAS TRUE   ⚠️`);
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
 * Validates that the shared exclusion list exists
 * @param {string} listName - The name of the shared exclusion list
 */
function validateSharedExclusionList(listName) {
  const listIterator = AdsApp.excludedPlacementLists()
    .withCondition(`Name = '${listName}'`)
    .get();

  if (!listIterator.hasNext()) {
    const errorMessage = `ERROR: Shared Placement Exclusion List named '${listName}' not found.\n\n` +
      `Please create this shared list in your Google Ads account first:\n` +
      `1. Go to Tools & Settings > Shared Library > Placement exclusions\n` +
      `2. Create a new list with the exact name: "${listName}"\n` +
      `3. Ensure the list is enabled\n` +
      `4. Link this list to your Performance Max campaigns if not already linked`;
    throw new Error(errorMessage);
  }

  const list = listIterator.next();
  console.log(`✓ Found shared exclusion list: ${listName}`);
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
    console.log(`cellRef: ${JSON.stringify(cellRef)}`);
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

    if (DEBUG_MODE) {
      console.log(`  Max Results parsed value: ${maxResultsNum} (isNaN: ${isNaN(maxResultsNum)})`);
    }

    if (!isNaN(maxResultsNum)) {
      maxResults = maxResultsNum; // Use the parsed value (0 means no limit, any other number is the limit)
    } else {
      // If parsing failed, use 0 (no limit)
      maxResults = 0;
    }
  }

  if (DEBUG_MODE) {
    console.log(`  Max Results final value: ${maxResults}`);
  }

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

  if (DEBUG_MODE) {
    console.log(`Settings loaded:`);
    console.log(`  Lookback Window: ${settings.lookbackWindowDays} days`);
    console.log(`  Impressions >: ${settings.minimumImpressions}`);
    console.log(`  Clicks >: ${settings.minimumClicks}`);
    console.log(`  Cost >: ${settings.minimumCost}`);
    console.log(`  Conversions <: ${settings.maximumConversions === 0 ? 'No limit' : settings.maximumConversions}`);
    console.log(`  Max Results: ${settings.maxResults === 0 ? 'No limit' : settings.maxResults}`);
    console.log(`  Campaign Name Contains: "${settings.campaignNameContains}"`);
    console.log(`  Campaign Name Not Contains: "${settings.campaignNameNotContains}"`);
    console.log(`  Enabled campaigns only: ${settings.enabledCampaignsOnly}`);
    console.log(`  Placement Type Filters:`);
    console.log(`    YouTube Video: enabled=${settings.placementTypeFilters.youtubeVideo.enabled}, automated=${settings.placementTypeFilters.youtubeVideo.automated}`);
    console.log(`    Website: enabled=${settings.placementTypeFilters.website.enabled}, automated=${settings.placementTypeFilters.website.automated}`);
    console.log(`    Mobile Application: enabled=${settings.placementTypeFilters.mobileApplication.enabled}, automated=${settings.placementTypeFilters.mobileApplication.automated}`);
    console.log(`    Google Products: enabled=${settings.placementTypeFilters.googleProducts.enabled}, automated=${settings.placementTypeFilters.googleProducts.automated}`);
    console.log(`  Enable ChatGPT: ${settings.enableChatGpt}`);
    if (settings.enableChatGpt) {
      console.log(`  ChatGPT API Key: ${settings.chatGptApiKey ? 'Provided' : 'Missing'}`);
      console.log(`  Use Cached ChatGPT Responses: ${settings.useCachedChatGpt}`);
    }
    console.log(`  Placement Contains Filter: ${settings.placementContainsEnabled ? 'Enabled' : 'Disabled'}`);
    if (settings.placementContainsEnabled && settings.placementContainsList.length > 0) {
      console.log(`    List: ${settings.placementContainsList.length} strings`);
      if (settings.placementContainsList.length <= 10) {
        console.log(`      ${settings.placementContainsList.join(', ')}`);
      } else {
        console.log(`      ${settings.placementContainsList.slice(0, 10).join(', ')} ... and ${settings.placementContainsList.length - 10} more`);
      }
    }
    console.log(`  Placement Not Contains Filter: ${settings.placementNotContainsEnabled ? 'Enabled' : 'Disabled'}`);
    if (settings.placementNotContainsEnabled && settings.placementNotContainsList.length > 0) {
      console.log(`    List: ${settings.placementNotContainsList.length} strings`);
      if (settings.placementNotContainsList.length <= 10) {
        console.log(`      ${settings.placementNotContainsList.join(', ')}`);
      } else {
        console.log(`      ${settings.placementNotContainsList.slice(0, 10).join(', ')} ... and ${settings.placementNotContainsList.length - 10} more`);
      }
    }
    console.log(`  Display Name Contains Filter: ${settings.displayNameContainsEnabled ? 'Enabled' : 'Disabled'}`);
    if (settings.displayNameContainsEnabled && settings.displayNameContainsList.length > 0) {
      console.log(`    List: ${settings.displayNameContainsList.length} strings`);
      if (settings.displayNameContainsList.length <= 10) {
        console.log(`      ${settings.displayNameContainsList.join(', ')}`);
      } else {
        console.log(`      ${settings.displayNameContainsList.slice(0, 10).join(', ')} ... and ${settings.displayNameContainsList.length - 10} more`);
      }
    }
    console.log(`  Display Name Not Contains Filter: ${settings.displayNameNotContainsEnabled ? 'Enabled' : 'Disabled'}`);
    if (settings.displayNameNotContainsEnabled && settings.displayNameNotContainsList.length > 0) {
      console.log(`    List: ${settings.displayNameNotContainsList.length} strings`);
      if (settings.displayNameNotContainsList.length <= 10) {
        console.log(`      ${settings.displayNameNotContainsList.join(', ')}`);
      } else {
        console.log(`      ${settings.displayNameNotContainsList.slice(0, 10).join(', ')} ... and ${settings.displayNameNotContainsList.length - 10} more`);
      }
    }
    console.log(`  Target URL Contains Filter: ${settings.targetUrlContainsEnabled ? 'Enabled' : 'Disabled'}`);
    if (settings.targetUrlContainsEnabled && settings.targetUrlContainsList.length > 0) {
      console.log(`    List: ${settings.targetUrlContainsList.length} strings`);
      if (settings.targetUrlContainsList.length <= 10) {
        console.log(`      ${settings.targetUrlContainsList.join(', ')}`);
      } else {
        console.log(`      ${settings.targetUrlContainsList.slice(0, 10).join(', ')} ... and ${settings.targetUrlContainsList.length - 10} more`);
      }
    }
    console.log(`  Target URL Not Contains Filter: ${settings.targetUrlNotContainsEnabled ? 'Enabled' : 'Disabled'}`);
    if (settings.targetUrlNotContainsEnabled && settings.targetUrlNotContainsList.length > 0) {
      console.log(`    List: ${settings.targetUrlNotContainsList.length} strings`);
      if (settings.targetUrlNotContainsList.length <= 10) {
        console.log(`      ${settings.targetUrlNotContainsList.join(', ')}`);
      } else {
        console.log(`      ${settings.targetUrlNotContainsList.slice(0, 10).join(', ')} ... and ${settings.targetUrlNotContainsList.length - 10} more`);
      }
    }
    console.log(`  Target URL Ends With Filter: ${settings.targetUrlEndsWithEnabled ? 'Enabled' : 'Disabled'}`);
    if (settings.targetUrlEndsWithEnabled && settings.targetUrlEndsWithList.length > 0) {
      console.log(`    List: ${settings.targetUrlEndsWithList.length} strings`);
      if (settings.targetUrlEndsWithList.length <= 10) {
        console.log(`      ${settings.targetUrlEndsWithList.join(', ')}`);
      } else {
        console.log(`      ${settings.targetUrlEndsWithList.slice(0, 10).join(', ')} ... and ${settings.targetUrlEndsWithList.length - 10} more`);
      }
    }
    console.log(`  Target URL Not Ends With Filter: ${settings.targetUrlNotEndsWithEnabled ? 'Enabled' : 'Disabled'}`);
    if (settings.targetUrlNotEndsWithEnabled && settings.targetUrlNotEndsWithList.length > 0) {
      console.log(`    List: ${settings.targetUrlNotEndsWithList.length} strings`);
      if (settings.targetUrlNotEndsWithList.length <= 10) {
        console.log(`      ${settings.targetUrlNotEndsWithList.join(', ')}`);
      } else {
        console.log(`      ${settings.targetUrlNotEndsWithList.slice(0, 10).join(', ')} ... and ${settings.targetUrlNotEndsWithList.length - 10} more`);
      }
    }
    console.log(`  Placement Type Filters:`);
    console.log(`    YouTube Video: ${settings.placementTypeFilters.youtubeVideo}`);
    console.log(`    Website: ${settings.placementTypeFilters.website}`);
    console.log(`    Mobile Application: ${settings.placementTypeFilters.mobileApplication}`);
    console.log(`    Google Products: ${settings.placementTypeFilters.googleProducts}`);
  }

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

  if (DEBUG_MODE) {
    console.log(`  Searching for "${settingName}" starting from row ${startRowIndex + 1}`);
  }

  // Start from the specified row index (defaults to 3 for row 4)
  for (let rowIndex = startRowIndex; rowIndex < values.length; rowIndex++) {
    const cellValue = String(values[rowIndex][0] || '').trim();
    if (cellValue === settingName) {
      const foundValue = values[rowIndex][1];
      if (DEBUG_MODE) {
        console.log(`  Found "${settingName}" at row ${rowIndex + 1}, value: ${foundValue} (type: ${typeof foundValue})`);
      }
      // Setting row found - always return the value from the sheet (even if empty)
      // This ensures defaults are only used when populating empty sheet, not when reading
      return foundValue;
    }
  }

  if (DEBUG_MODE) {
    console.log(`  "${settingName}" not found, using default: ${defaultValue}`);
  }

  // Setting row not found - return default (only happens if sheet is malformed)
  return defaultValue;
}

/**
 * Gets the Placement Contains list and enabled status from the Lists sheet (checks placement field)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getPlacementContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'placementContains', 1);
}

/**
 * Gets the Placement Not Contains list and enabled status from the Lists sheet (checks placement field)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getPlacementNotContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'placementNotContains', 2);
}

/**
 * Gets the Display Name Contains list and enabled status from the Lists sheet (checks displayName field)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getDisplayNameContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'displayNameContains', 3);
}

/**
 * Gets the Display Name Not Contains list and enabled status from the Lists sheet (checks displayName field)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getDisplayNameNotContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'displayNameNotContains', 4);
}

/**
 * Gets the Target URL Contains list and enabled status from the Lists sheet (checks targetUrl field)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getTargetUrlContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'targetUrlContains', 5);
}

/**
 * Gets the Target URL Not Contains list and enabled status from the Lists sheet (checks targetUrl field)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getTargetUrlNotContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'targetUrlNotContains', 6);
}

/**
 * Gets the Target URL Ends With list and enabled status from the Lists sheet (checks targetUrl field)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getTargetUrlEndsWithList(listsSheet) {
  return getListFromSheet(listsSheet, 'targetUrlEndsWith', 7);
}

/**
 * Gets the Target URL Not Ends With list and enabled status from the Lists sheet (checks targetUrl field)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getTargetUrlNotEndsWithList(listsSheet) {
  return getListFromSheet(listsSheet, 'targetUrlNotEndsWith', 8);
}

/**
 * Generic function to get a list and enabled status from the Lists sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @param {string} listKey - The key in PLACEMENT_FILTERS_CELL_REFERENCES (e.g., 'placementContains')
 * @param {number} columnIndex - The column index (1 = A, 2 = B, etc.)
 * @returns {Object} Object with enabled status and list array {enabled: boolean, list: Array<string>}
 */
function getListFromSheet(listsSheet, listKey, columnIndex) {
  const list = [];

  // Get checkbox cell reference
  const checkboxRef = PLACEMENT_FILTERS_CELL_REFERENCES[listKey];
  if (!checkboxRef) {
    if (DEBUG_MODE) {
      console.log(`  ${listKey} not found in PLACEMENT_FILTERS_CELL_REFERENCES`);
    }
    return { enabled: false, list: [] };
  }

  // Read the enable checkbox value
  const checkboxValue = listsSheet.getRange(checkboxRef.row, checkboxRef.column).getValue();
  let enabled = false;
  if (typeof checkboxValue === 'boolean') {
    enabled = checkboxValue;
  } else if (typeof checkboxValue === 'string') {
    const lowerValue = checkboxValue.toLowerCase().trim();
    enabled = lowerValue === 'true' || lowerValue === '1';
  }

  // Get list start row
  const listStartRow = PLACEMENT_FILTERS_LIST_START_ROWS[listKey];
  if (!listStartRow) {
    if (DEBUG_MODE) {
      console.log(`  ${listKey} not found in PLACEMENT_FILTERS_LIST_START_ROWS`);
    }
    return { enabled: enabled, list: [] };
  }

  if (DEBUG_MODE && listKey.indexOf('Key') === -1) {
    console.log(`  Reading ${listKey} - checkbox at row ${checkboxRef.row}, column ${checkboxRef.column}: ${enabled}`);
    console.log(`  Starting to read list from row ${listStartRow}, column ${columnIndex}`);
  }

  const dataRange = listsSheet.getDataRange();
  const values = dataRange.getValues();
  const columnArrayIndex = columnIndex - 1; // Convert to 0-based array index
  const listStartRowIndex = listStartRow - 1; // Convert to 0-based array index

  // Read from the specified column starting from listStartRowIndex
  // Continue reading until we hit the next header (in any column) or end of sheet
  // Don't stop at empty rows - continue until we hit a header or end
  for (let rowIndex = listStartRowIndex; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];

    // Check if this row contains a header (any of the known list headers in any column)
    // If so, stop reading (we've hit the next list section)
    const knownHeaders = [
      'Placement Contains', 'Placement Not Contains',
      'Display Name Contains', 'Display Name Not Contains',
      'Target URL Contains', 'Target URL Not Contains',
      'Target URL Ends With', 'Target URL Not Ends With'
    ];

    let isHeaderRow = false;
    for (let colIndex = 0; colIndex < row.length; colIndex++) {
      const cellValue = String(row[colIndex] || '').trim();
      if (knownHeaders.includes(cellValue)) {
        isHeaderRow = true;
        break;
      }
    }

    if (isHeaderRow && rowIndex > listStartRowIndex) {
      // We've hit the next header row, stop reading
      break;
    }

    // Read the value from the correct column
    const cellValue = values[rowIndex][columnArrayIndex];
    const item = String(cellValue || '').trim();

    // Skip empty cells (but continue reading - don't break)
    if (!item) {
      continue;
    }

    // Skip description rows (they contain "Description" or "Only include" or "Exclude")
    if (item.toLowerCase().includes('description') ||
      item.toLowerCase().includes('only include') ||
      item.toLowerCase().includes('exclude placements') ||
      item.toLowerCase().includes('will still appear')) {
      continue;
    }

    // Skip enable checkbox row
    if (item.toLowerCase().includes('enable')) {
      continue;
    }

    // Add the item to the list
    if (item) {
      list.push(item);
    }
  }

  if (DEBUG_MODE) {
    console.log(`  ${listKey} read: ${list.length} items found, enabled: ${enabled}`);
    if (list.length > 0 && list.length <= 10) {
      console.log(`    Items: ${list.join(', ')}`);
    } else if (list.length > 10) {
      console.log(`    First 10 items: ${list.slice(0, 10).join(', ')}`);
      console.log(`    ... and ${list.length - 10} more`);
    }
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
    console.error(`Error getting Display placements: ${error.message}`);
  }

  if (allPlacementData.length > 0) {
    console.warn(`⚠ Note: performance_max_placement_view only supports impressions metric.`);
    console.warn(`   Display campaign placements may have additional metrics available.`);
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
 * @param {string} campaignType - 'PERFORMANCE_MAX' or 'DISPLAY'
 * @returns {string} GAQL query string
 */
function getPlacementGaqlQuery(dateRange, settings, campaignType) {
  const conditions = [
    `segments.date BETWEEN '${dateRange.startDate}' AND '${dateRange.endDate}'`,
    `campaign.advertising_channel_type = '${campaignType}'`
  ];

  // Campaign status filter: always exclude REMOVED, include PAUSED based on setting
  if (settings.enabledCampaignsOnly) {
    conditions.push(`campaign.status = 'ENABLED'`);
  } else {
    conditions.push(`campaign.status IN ('ENABLED', 'PAUSED')`);
  }

  // Note: Campaign name filters are applied in JavaScript after fetching data
  // because Google Ads Scripts GAQL doesn't support REGEXP_MATCH or case-insensitive matching

  const whereClause = conditions.join(' AND ');

  // Add LIMIT clause if maxResults is set (0 means no limit)
  // Note: LIMIT applies per query type, so total results may be up to 2x this value (PMax + Display)
  const limitClause = settings.maxResults > 0 ? ` LIMIT ${settings.maxResults}` : '';

  if (campaignType === 'PERFORMANCE_MAX') {
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
  } else if (campaignType === 'DISPLAY') {
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
    console.error(`For Display: https://developers.google.com/google-ads/api/fields/v20/detail_placement_view_query_builder`);
    throw error;
  }
}

/**
 * Extracts placement data from the report and returns flat objects
 * @param {GoogleAppsScript.AdsApp.Report} report - The report object
 * @param {string} campaignType - 'PERFORMANCE_MAX' or 'DISPLAY'
 * @returns {Array<Object>} Array of flat placement objects
 */
function extractPlacementDataFromReport(report, campaignType) {
  const placementData = [];
  const rows = report.rows();
  let rowCount = 0;

  // Determine the view prefix based on campaign type
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
      targetUrl: campaignType === 'DISPLAY'
        ? (row['detail_placement_view.group_placement_target_url'] || row['detail_placement_view.target_url'] || '')
        : row[`${viewPrefix}.target_url`],
      resourceName: campaignType === 'DISPLAY' ? (row['detail_placement_view.resource_name'] || '') : '',
      impressions: parseInt(row['metrics.impressions']) || 0,
      clicks: campaignType === 'DISPLAY' ? parseInt(row['metrics.clicks']) || 0 : 0,
      costMicros: campaignType === 'DISPLAY' ? parseInt(row['metrics.cost_micros']) || 0 : 0,
      conversions: campaignType === 'DISPLAY' ? parseFloat(row['metrics.conversions']) || 0 : 0,
      conversionsValue: campaignType === 'DISPLAY' ? parseFloat(row['metrics.conversions_value']) || 0 : 0
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
  return placementData.filter(placement => {
    const meetsImpressionsThreshold = placement.impressions > settings.minimumImpressions;
    // Note: clicks are not available in performance_max_placement_view, so this filter will always pass
    // if minimumClicks is 0, otherwise it will filter out all placements
    const meetsClicksThreshold = settings.minimumClicks === 0 || placement.clicks > settings.minimumClicks;
    const meetsCostThreshold = settings.minimumCost === 0 || placement.cost > settings.minimumCost;
    const meetsConversionsThreshold = settings.maximumConversions === 0 || placement.conversions < settings.maximumConversions;

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
        // Unknown placement type - include it by default
        meetsPlacementTypeFilter = true;
      }

      if (DEBUG_MODE && !meetsPlacementTypeFilter) {
        console.log(`  Filtered out placement (placement type not selected): ${placement.placement} (type: ${placement.placementType})`);
      }
    }

    // Check Placement Contains filter (checks placement field only, only if enabled and list has items)
    let meetsPlacementContainsFilter = true;
    if (settings.placementContainsEnabled) {
      if (settings.placementContainsList && settings.placementContainsList.length > 0) {
        const placementString = String(placement.placement || '').toLowerCase();
        let containsMatch = false;
        for (const containsString of settings.placementContainsList) {
          const lowerContainsString = containsString.toLowerCase();
          if (placementString.includes(lowerContainsString)) {
            containsMatch = true;
            break;
          }
        }
        meetsPlacementContainsFilter = containsMatch;
        if (DEBUG_MODE && !meetsPlacementContainsFilter) {
          console.log(`  Filtered out placement (Placement field does not contain required strings): ${placement.placement}`);
        }
      } else {
        // Filter is enabled but list is empty - skip filter (allow all)
        if (DEBUG_MODE) {
          console.log(`  Placement Contains filter enabled but list is empty - skipping filter`);
        }
      }
    }

    // Check Placement Not Contains filter (checks placement field only, only if enabled and list has items)
    let meetsPlacementNotContainsFilter = true;
    if (settings.placementNotContainsEnabled) {
      if (settings.placementNotContainsList && settings.placementNotContainsList.length > 0) {
        const placementString = String(placement.placement || '').toLowerCase();
        let containsMatch = false;
        for (const notContainsString of settings.placementNotContainsList) {
          const lowerNotContainsString = notContainsString.toLowerCase();
          if (placementString.includes(lowerNotContainsString)) {
            containsMatch = true;
            break;
          }
        }
        // If it contains any "not contains" item, filter it out
        meetsPlacementNotContainsFilter = !containsMatch;
        if (DEBUG_MODE && !meetsPlacementNotContainsFilter) {
          console.log(`  Filtered out placement (Placement field contains excluded strings): ${placement.placement}`);
        }
      } else {
        // Filter is enabled but list is empty - skip filter (allow all)
        if (DEBUG_MODE) {
          console.log(`  Placement Not Contains filter enabled but list is empty - skipping filter`);
        }
      }
    }

    // Check Display Name Contains filter (checks displayName field only, only if enabled and list has items)
    let meetsDisplayNameContainsFilter = true;
    if (settings.displayNameContainsEnabled) {
      if (settings.displayNameContainsList && settings.displayNameContainsList.length > 0) {
        const displayNameString = String(placement.displayName || '').toLowerCase();
        let containsMatch = false;
        for (const containsString of settings.displayNameContainsList) {
          const lowerContainsString = containsString.toLowerCase();
          if (displayNameString.includes(lowerContainsString)) {
            containsMatch = true;
            break;
          }
        }
        meetsDisplayNameContainsFilter = containsMatch;
        if (DEBUG_MODE && !meetsDisplayNameContainsFilter) {
          console.log(`  Filtered out placement (Display Name field does not contain required strings): ${placement.displayName}`);
        }
      } else {
        // Filter is enabled but list is empty - skip filter (allow all)
        if (DEBUG_MODE) {
          console.log(`  Display Name Contains filter enabled but list is empty - skipping filter`);
        }
      }
    }

    // Check Display Name Not Contains filter (checks displayName field only, only if enabled and list has items)
    let meetsDisplayNameNotContainsFilter = true;
    if (settings.displayNameNotContainsEnabled) {
      if (settings.displayNameNotContainsList && settings.displayNameNotContainsList.length > 0) {
        const displayNameString = String(placement.displayName || '').toLowerCase();
        let containsMatch = false;
        for (const notContainsString of settings.displayNameNotContainsList) {
          const lowerNotContainsString = notContainsString.toLowerCase();
          if (displayNameString.includes(lowerNotContainsString)) {
            containsMatch = true;
            break;
          }
        }
        // If it contains any "not contains" item, filter it out
        meetsDisplayNameNotContainsFilter = !containsMatch;
        if (DEBUG_MODE && !meetsDisplayNameNotContainsFilter) {
          console.log(`  Filtered out placement (Display Name field contains excluded strings): ${placement.displayName}`);
        }
      } else {
        // Filter is enabled but list is empty - skip filter (allow all)
        if (DEBUG_MODE) {
          console.log(`  Display Name Not Contains filter enabled but list is empty - skipping filter`);
        }
      }
    }

    // Check Target URL Contains filter (checks targetUrl field only, only if enabled and list has items)
    let meetsTargetUrlContainsFilter = true;
    if (settings.targetUrlContainsEnabled) {
      if (settings.targetUrlContainsList && settings.targetUrlContainsList.length > 0) {
        const targetUrlString = String(placement.targetUrl || '').toLowerCase();
        let containsMatch = false;
        for (const containsString of settings.targetUrlContainsList) {
          const lowerContainsString = containsString.toLowerCase();
          if (targetUrlString.includes(lowerContainsString)) {
            containsMatch = true;
            break;
          }
        }
        meetsTargetUrlContainsFilter = containsMatch;
        if (DEBUG_MODE && !meetsTargetUrlContainsFilter) {
          console.log(`  Filtered out placement (Target URL field does not contain required strings): ${placement.targetUrl}`);
        }
      } else {
        // Filter is enabled but list is empty - skip filter (allow all)
        if (DEBUG_MODE) {
          console.log(`  Target URL Contains filter enabled but list is empty - skipping filter`);
        }
      }
    }

    // Check Target URL Ends With filter (checks targetUrl field only, only if enabled and list has items)
    let meetsTargetUrlEndsWithFilter = true;
    if (settings.targetUrlEndsWithEnabled) {
      if (settings.targetUrlEndsWithList && settings.targetUrlEndsWithList.length > 0) {
        const targetUrlString = String(placement.targetUrl || '').toLowerCase();
        let endsWithMatch = false;
        for (const endsWithString of settings.targetUrlEndsWithList) {
          const lowerEndsWithString = endsWithString.toLowerCase();
          if (targetUrlString.endsWith(lowerEndsWithString)) {
            endsWithMatch = true;
            break;
          }
        }
        meetsTargetUrlEndsWithFilter = endsWithMatch;
        if (DEBUG_MODE && !meetsTargetUrlEndsWithFilter) {
          console.log(`  Filtered out placement (Target URL field does not end with required strings): ${placement.targetUrl}`);
        }
      } else {
        // Filter is enabled but list is empty - skip filter (allow all)
        if (DEBUG_MODE) {
          console.log(`  Target URL Ends With filter enabled but list is empty - skipping filter`);
        }
      }
    }

    // Check Target URL Not Contains filter (checks targetUrl field only, only if enabled and list has items)
    let meetsTargetUrlNotContainsFilter = true;
    if (settings.targetUrlNotContainsEnabled) {
      if (settings.targetUrlNotContainsList && settings.targetUrlNotContainsList.length > 0) {
        const targetUrlString = String(placement.targetUrl || '').toLowerCase();
        let containsMatch = false;
        for (const notContainsString of settings.targetUrlNotContainsList) {
          const lowerNotContainsString = notContainsString.toLowerCase();
          if (targetUrlString.includes(lowerNotContainsString)) {
            containsMatch = true;
            break;
          }
        }
        // If it contains any "not contains" item, filter it out
        meetsTargetUrlNotContainsFilter = !containsMatch;
        if (DEBUG_MODE && !meetsTargetUrlNotContainsFilter) {
          console.log(`  Filtered out placement (Target URL field contains excluded strings): ${placement.targetUrl}`);
        }
      } else {
        // Filter is enabled but list is empty - skip filter (allow all)
        if (DEBUG_MODE) {
          console.log(`  Target URL Not Contains filter enabled but list is empty - skipping filter`);
        }
      }
    }

    // Check Target URL Not Ends With filter (checks targetUrl field only, only if enabled and list has items)
    let meetsTargetUrlNotEndsWithFilter = true;
    if (settings.targetUrlNotEndsWithEnabled) {
      if (settings.targetUrlNotEndsWithList && settings.targetUrlNotEndsWithList.length > 0) {
        const targetUrlString = String(placement.targetUrl || '').toLowerCase();
        let endsWithMatch = false;
        for (const notEndsWithString of settings.targetUrlNotEndsWithList) {
          const lowerNotEndsWithString = notEndsWithString.toLowerCase();
          if (targetUrlString.endsWith(lowerNotEndsWithString)) {
            endsWithMatch = true;
            break;
          }
        }
        // If it ends with any "not ends with" item, filter it out
        meetsTargetUrlNotEndsWithFilter = !endsWithMatch;
        if (DEBUG_MODE && !meetsTargetUrlNotEndsWithFilter) {
          console.log(`  Filtered out placement (Target URL field ends with excluded strings): ${placement.targetUrl}`);
        }
      } else {
        // Filter is enabled but list is empty - skip filter (allow all)
        if (DEBUG_MODE) {
          console.log(`  Target URL Not Ends With filter enabled but list is empty - skipping filter`);
        }
      }
    }

    return meetsImpressionsThreshold && meetsClicksThreshold && meetsCostThreshold && meetsConversionsThreshold && meetsPlacementTypeFilter &&
      meetsPlacementContainsFilter && meetsPlacementNotContainsFilter &&
      meetsDisplayNameContainsFilter && meetsDisplayNameNotContainsFilter &&
      meetsTargetUrlContainsFilter && meetsTargetUrlNotContainsFilter &&
      meetsTargetUrlEndsWithFilter && meetsTargetUrlNotEndsWithFilter;
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
    const listId = excludedPlacementList.getId();
    console.log(`✓ Found exclusion list: ${retrievedListName}`);
    console.log(`  List ID: ${listId}`);

    // Try using GAQL query to get excluded placements (may work better for Bulk Upload placements)
    console.log(`Attempting to get placements via GAQL query...`);
    try {
      const gaqlQuery = `
        SELECT
          shared_set_criterion.placement,
          shared_set_criterion.placement_type
        FROM shared_set_criterion
        WHERE shared_set.id = ${listId}
      `;

      console.log(`  GAQL Query: ${gaqlQuery}`);
      const report = AdsApp.report(gaqlQuery);
      const rows = report.rows();
      const totalEntities = rows.totalNumEntities();
      console.log(`  Total entities from GAQL: ${totalEntities}`);

      if (totalEntities > 0) {
        let placementCount = 0;
        while (rows.hasNext()) {
          try {
            const row = rows.next();
            const placementUrl = row['shared_set_criterion.placement'];
            const placementType = row['shared_set_criterion.placement_type'];
            placementCount++;
            console.log(`  Placement ${placementCount}: ${placementUrl} (type: ${placementType})`);

            // Normalize website URLs and add mobile app IDs as-is
            if (placementUrl) {
              if (placementUrl.startsWith('mobileapp::')) {
                // Mobile app format: mobileapp::1-1510189987 or mobileapp::2-com.example.app
                // Extract the canonical ID (remove mobileapp:: prefix)
                const canonicalId = placementUrl.replace(/^mobileapp::/, '');
                excludedPlacements.add(canonicalId);
                excludedPlacements.add(placementUrl); // Also add the full format for matching
                console.log(`    ✓ Added mobile app: canonical=${canonicalId}, full=${placementUrl}`);
              } else {
                // Website URL - normalize it
                const normalized = normalizeUrl(placementUrl);
                if (normalized) {
                  excludedPlacements.add(normalized);
                  console.log(`    ✓ Added website URL: original=${placementUrl}, normalized=${normalized}`);
                } else {
                  console.log(`    ⚠ Could not normalize URL: ${placementUrl}`);
                }
              }
            } else {
              console.log(`    ⚠ Placement ${placementCount} has no URL`);
            }
          } catch (rowError) {
            console.error(`    ✗ Error processing row ${placementCount}: ${rowError.message}`);
          }
        }
        console.log(`Total placements found via GAQL: ${placementCount}`);
      } else {
        console.log(`⚠ GAQL query returned 0 placements`);
      }
    } catch (gaqlError) {
      console.log(`✗ GAQL query failed: ${gaqlError.message}`);
      console.log(`Falling back to iterator method...`);

      // Fallback to iterator method
      const placementSelector = excludedPlacementList.excludedPlacements();
      console.log(`  Placement selector created`);

      const placementIterator = placementSelector.get();
      console.log(`  Placement iterator created`);

      // Check total entities first
      try {
        const totalEntities = placementIterator.totalNumEntities();
        console.log(`  Total entities in iterator: ${totalEntities}`);
      } catch (e) {
        console.log(`  Could not get total entities: ${e.message}`);
      }

      const hasNext = placementIterator.hasNext();
      console.log(`  Iterator hasNext(): ${hasNext}`);

      if (!hasNext) {
        console.log(`⚠ No placements found in iterator. The list may be empty or there may be an issue accessing placements.`);
      }

      let placementCount = 0;
      while (placementIterator.hasNext()) {
        try {
          const sharedPlacement = placementIterator.next();
          placementCount++;
          console.log(`  Processing placement ${placementCount}...`);

          // Use getUrl() method (standard method per docs)
          const placementUrl = sharedPlacement.getUrl();

          if (placementUrl) {
            console.log(`    Placement URL: ${placementUrl}`);

            // Only add website URLs - mobile apps can't be reliably matched via getUrl()
            // Normalize website URL
            const normalized = normalizeUrl(placementUrl);
            if (normalized) {
              excludedPlacements.add(normalized);
              console.log(`    ✓ Added website URL: original=${placementUrl}, normalized=${normalized}`);
            } else {
              console.log(`    ⚠ Could not normalize URL: ${placementUrl}`);
            }
          } else {
            console.log(`    ⚠ Placement ${placementCount} has no URL`);
          }
        } catch (placementError) {
          console.error(`    ✗ Error processing placement ${placementCount}: ${placementError.message}`);
          console.error(`    Error stack: ${placementError.stack}`);
        }
      }
      console.log(`Total placements found via iterator: ${placementCount}`);
    }

    console.log(`Total unique placements in set: ${excludedPlacements.size}`);
    if (DEBUG_MODE && excludedPlacements.size > 0) {
      const placementsArray = Array.from(excludedPlacements);
      console.log(`Excluded placements set contains:`);
      placementsArray.slice(0, 10).forEach((p, idx) => {
        console.log(`  ${idx + 1}. ${p}`);
      });
      if (placementsArray.length > 10) {
        console.log(`  ... and ${placementsArray.length - 10} more`);
      }
    }
    console.log(`===================================\n`);

  } catch (error) {
    console.error(`✗ Error getting excluded placements: ${error.message}`);
    console.error(`Error stack: ${error.stack}`);
    if (DEBUG_MODE) {
      console.log(`Full error details: ${JSON.stringify(error)}`);
    }
  }

  return excludedPlacements;
}

/**
 * Checks if a placement is already excluded
 * Only checks website placements - returns "Unknown" for mobile apps and other types
 * @param {Object} placement - Placement object with placement, placementType, and targetUrl
 * @param {Set<string>} excludedPlacements - Set of excluded placements (normalized website URLs only)
 * @returns {string} "Excluded" if excluded, "Unknown" if can't determine, "" if not excluded
 */
function getPlacementStatus(placement, excludedPlacements) {
  if (!placement) {
    if (DEBUG_MODE) {
      console.log(`  Placement check: placement is null/undefined`);
    }
    return '';
  }

  if (excludedPlacements.size === 0) {
    if (DEBUG_MODE) {
      console.log(`  Placement check: excluded placements set is empty`);
    }
    return '';
  }

  const placementValue = placement.placement || '';
  const placementType = placement.placementType || '';
  const targetUrl = placement.targetUrl || '';

  if (DEBUG_MODE) {
    console.log(`\n  Checking placement: value="${placementValue}", type="${placementType}", targetUrl="${targetUrl}"`);
  }

  // Only check website placements - mobile apps and others return "Unknown"
  if (placementType && String(placementType).toUpperCase().includes('MOBILE_APPLI')) {
    if (DEBUG_MODE) {
      console.log(`    Mobile app detected - cannot reliably check status, returning "Unknown"`);
    }
    return 'Unknown';
  }

  // Check website placements only
  const url = targetUrl || placementValue;
  if (url) {
    if (DEBUG_MODE) {
      console.log(`    Website placement detected. URL: "${url}"`);
    }

    const normalized = normalizeUrl(url);
    if (normalized) {
      if (DEBUG_MODE) {
        console.log(`    Normalized URL: "${normalized}"`);
        console.log(`    Checking if "${normalized}" is in excluded set...`);
      }

      if (excludedPlacements.has(normalized)) {
        if (DEBUG_MODE) {
          console.log(`    ✓ MATCH: Normalized URL "${normalized}" found in excluded set`);
        }
        return 'Excluded';
      } else {
        if (DEBUG_MODE) {
          console.log(`    ✗ NO MATCH: Normalized URL not found in excluded set`);
        }
        return 'Active';
      }
    } else {
      if (DEBUG_MODE) {
        console.log(`    ⚠ Could not normalize URL: "${url}" - returning "Unknown"`);
      }
      return 'Unknown';
    }
  }

  // If we can't determine the type or URL, return "Unknown"
  if (DEBUG_MODE) {
    console.log(`    ⚠ Could not determine placement type or URL - returning "Unknown"`);
  }
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

  // Clear ALL output sheets first, regardless of whether there's data
  for (const sheetName of allOutputSheetNames) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      sheet.clear();
      if (DEBUG_MODE) {
        console.log(`Cleared sheet: ${sheetName}`);
      }
    }
  }

  // If no data, write "No results found" message to all output sheets
  if (placementData.length === 0) {
    console.log(`No placement data found. Writing 'No results found' to all output sheets.`);
    const timestamp = new Date().toLocaleString();
    const noResultsData = [['No results found', `Last updated: ${timestamp}`]];

    for (const sheetName of allOutputSheetNames) {
      const sheet = getOrCreateOutputSheet(spreadsheet, sheetName);
      sheet.getRange(1, 1, 1, 2).setValues(noResultsData);
      sheet.setTabColor('#ea4335'); // Red tab for no results
    }
    return;
  }

  // Get currently excluded placements
  const excludedPlacements = getExcludedPlacements(settings.sharedExclusionListName);
  if (DEBUG_MODE) {
    console.log(`Found ${excludedPlacements.size} placements already in exclusion list`);
    if (excludedPlacements.size > 0 && excludedPlacements.size <= 10) {
      const excludedArray = Array.from(excludedPlacements);
      console.log(`Excluded placements: ${excludedArray.join(', ')}`);
    }
  }

  // Headers for sheets WITH ChatGPT Response (Website, YouTube)
  const headersWithChatGpt = [
    'Exclude',
    'Status',
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

  // Headers for sheets WITHOUT ChatGPT Response (Mobile Application, Google Products)
  const headersWithoutChatGpt = [
    'Exclude',
    'Status',
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
      if (DEBUG_MODE) {
        console.log(`Skipping placement with unknown type: ${placement.placement} (type: ${placement.placementType})`);
      }
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

    // Determine if this placement type should have ChatGPT Response column
    const skipChatGptColumn = placementTypeName === MOBILE_APPLICATION_OUTPUT_SHEET_NAME ||
      placementTypeName === GOOGLE_PRODUCTS_OUTPUT_SHEET_NAME;
    const headers = skipChatGptColumn ? headersWithoutChatGpt : headersWithChatGpt;

    // Clear existing data
    placementTypeSheet.clear();

    // Write headers
    const headerRange = placementTypeSheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');

    // Get automation setting for this placement type
    const isAutomated = getAutomationForPlacementType(settings, placementTypeName);

    // Write data rows
    const dataRows = placements.map(placement => {
      // Add note for mobile apps
      let notes = '';
      if (placement.placementType && String(placement.placementType).toUpperCase().includes('MOBILE_APPLI')) {
        notes = 'Will be excluded via bulk upload if selected';
      }

      // Check if placement is already excluded (only for website placements)
      const status = getPlacementStatus(placement, excludedPlacements);

      // Use automation setting for this placement type
      const shouldCheck = isAutomated || false;

      // Base row without ChatGPT Response
      const baseRow = [
        shouldCheck, // Checkbox column (checked if automation is enabled for this type, otherwise unchecked)
        status, // Status column
        placement.campaignName,
        placement.placement,
        placement.displayName,
        placement.placementType,
        placement.targetUrl,
        notes // Notes column
      ];

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

      // Include ChatGPT Response only for Website and YouTube
      if (skipChatGptColumn) {
        return [...baseRow, ...metricsColumns];
      } else {
        return [...baseRow, placement.chatGptResponse || '', ...metricsColumns];
      }
    });

    if (dataRows.length > 0) {
      const dataRange = placementTypeSheet.getRange(2, 1, dataRows.length, headers.length);
      dataRange.setValues(dataRows);

      // Add checkboxes to first column
      const checkboxRange = placementTypeSheet.getRange(2, 1, dataRows.length, 1);
      checkboxRange.insertCheckboxes();

      // Format number columns (pass whether this sheet has ChatGPT column)
      formatPlacementSheet(placementTypeSheet, placements.length, !skipChatGptColumn);
    }

    // Freeze header row and first column
    placementTypeSheet.setFrozenRows(1);
    placementTypeSheet.setFrozenColumns(1);

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
      sheet.getRange(1, 1, 1, 2).setValues(noResultsData);
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
    if (DEBUG_MODE) {
      console.log(`Created placement type sheet: ${placementTypeName}`);
    }
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
    if (DEBUG_MODE) {
      console.log(`Created output sheet: ${sheetName}`);
    }
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
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to format
 * @param {number} numRows - Number of data rows
 */
function formatPlacementSheet(sheet, numRows, hasChatGptColumn = true) {
  // Column offset: if no ChatGPT column, metrics start 1 column earlier
  const offset = hasChatGptColumn ? 0 : -1;

  // Format impressions, clicks, conversions (integers)
  sheet.getRange(2, 10 + offset, numRows, 1).setNumberFormat('#,##0'); // Impressions
  sheet.getRange(2, 11 + offset, numRows, 1).setNumberFormat('#,##0'); // Clicks
  sheet.getRange(2, 13 + offset, numRows, 1).setNumberFormat('#,##0'); // Conversions

  // Format cost, CPA, Avg CPC (currency)
  sheet.getRange(2, 12 + offset, numRows, 1).setNumberFormat('#,##0.00'); // Cost
  sheet.getRange(2, 16 + offset, numRows, 1).setNumberFormat('#,##0.00'); // Avg CPC
  sheet.getRange(2, 18 + offset, numRows, 1).setNumberFormat('#,##0.00'); // CPA

  // Format conversions value, ROAS (currency)
  sheet.getRange(2, 14 + offset, numRows, 1).setNumberFormat('#,##0.00'); // Conv. Value
  sheet.getRange(2, 19 + offset, numRows, 1).setNumberFormat('#,##0.00'); // ROAS

  // Format percentages
  sheet.getRange(2, 15 + offset, numRows, 1).setNumberFormat('0.00%'); // CTR
  sheet.getRange(2, 17 + offset, numRows, 1).setNumberFormat('0.00%'); // Conv. Rate

  // Format Notes column (clip text - no wrap)
  sheet.getRange(2, 8, numRows, 1).setWrap(false); // Notes (column H - always same position)

  // Format ChatGPT Response column if it exists
  if (hasChatGptColumn) {
    sheet.getRange(2, 9, numRows, 1).setWrap(false); // ChatGPT Response (column I)
  }
}

// --- Exclusion List Functions ---

/**
 * Reads checked placements from all placement type sheets and normalizes them
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @returns {Object} Object with websitePlacements, youtubeVideos, and mobileAppPlacements arrays
 */
function readCheckedPlacementsFromSheet(spreadsheet) {
  const checkedWebsitePlacementsSet = new Set();
  const checkedYouTubeVideosSet = new Set();
  const checkedMobileAppPlacementsSet = new Set();
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
      if (DEBUG_MODE) {
        console.log(`  Skipping ${placementTypeName} sheet as it does not exist.`);
      }
      continue;
    }

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) {
      if (DEBUG_MODE) {
        console.log(`  No data in ${placementTypeName} sheet. Nothing to read.`);
      }
      continue;
    }

    // Find column indices dynamically
    const headers = values[0];
    const checkboxIndex = headers.indexOf('Exclude');
    const placementIndex = headers.indexOf('Placement');
    const placementTypeIndex = headers.indexOf('Placement Type');
    const targetUrlIndex = headers.indexOf('Target URL');

    if (checkboxIndex === -1 || placementIndex === -1 || placementTypeIndex === -1 || targetUrlIndex === -1) {
      console.warn(`  Skipping ${placementTypeName} sheet: Missing one or more required columns (Exclude, Placement, Placement Type, Target URL).`);
      continue;
    }

    let checkedCount = 0;
    let uncheckedCount = 0;

    for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
      const row = values[rowIndex];
      const checkboxValue = row[checkboxIndex];
      const isChecked = checkboxValue === true || String(checkboxValue).toUpperCase() === 'TRUE';
      const placementValue = String(row[placementIndex] || '').trim();
      const placementType = String(row[placementTypeIndex] || '').toUpperCase();
      const targetUrl = String(row[targetUrlIndex] || '').trim();

      if (DEBUG_MODE && rowIndex <= 3) {
        console.log(`  Reading from ${placementTypeName} - Row ${rowIndex + 1}: checkbox=${checkboxValue} (type: ${typeof checkboxValue}), placement=${placementValue}, type=${placementType}, targetUrl=${targetUrl}`);
      }

      if (isChecked) {
        checkedCount++;
      } else {
        uncheckedCount++;
      }

      if (isChecked && placementValue) {
        const mobileAppPlacement = formatMobileAppPlacement(placementValue, placementType);

        if (mobileAppPlacement) {
          checkedMobileAppPlacementsSet.add(mobileAppPlacement);
        } else {
          const placementTypeUpper = String(placementType || '').toUpperCase();
          const isYouTubeVideo = placementTypeUpper.includes('YOUTUBE') || placementTypeUpper.includes('YOUTUBE_VIDEO');

          if (isYouTubeVideo) {
            // For YouTube videos, use the placement value directly (video ID)
            // The placement value might be a video ID like "K2zZwB3NY04" or a URL
            if (placementValue.includes('youtube.com') || placementValue.includes('youtu.be')) {
              // Extract video ID from URL if it's a full URL
              const videoIdMatch = placementValue.match(/(?:youtube\.com\/watch\?v=|youtu\.be\/)([a-zA-Z0-9_-]+)/);
              if (videoIdMatch && videoIdMatch[1]) {
                checkedYouTubeVideosSet.add(videoIdMatch[1]);
              } else {
                // Use the placement value as-is
                checkedYouTubeVideosSet.add(placementValue);
              }
            } else {
              // It's likely already a video ID, use it directly
              checkedYouTubeVideosSet.add(placementValue);
            }
          } else {
            // It's a website URL, normalize it
            const normalizedUrl = normalizeUrl(placementValue);

            if (normalizedUrl === null) {
              skippedInvalid++;
            } else {
              checkedWebsitePlacementsSet.add(normalizedUrl);
            }
          }
        }
      }
    }
    if (DEBUG_MODE) {
      console.log(`  ${placementTypeName} sheet: Checked=${checkedCount}, Unchecked=${uncheckedCount}`);
    }
    totalCheckedCount += checkedCount;
    totalUncheckedCount += uncheckedCount;
  }

  const websitePlacements = Array.from(checkedWebsitePlacementsSet);
  const youtubeVideos = Array.from(checkedYouTubeVideosSet);
  const mobileAppPlacements = Array.from(checkedMobileAppPlacementsSet);
  const checkedPlacements = [...websitePlacements, ...youtubeVideos, ...mobileAppPlacements];

  if (DEBUG_MODE) {
    console.log(`Total checked boxes found across all sheets: ${totalCheckedCount}, Total unchecked: ${totalUncheckedCount}`);
    console.log(`Valid placements to exclude: ${checkedPlacements.length}`);
  }

  if (checkedPlacements.length > 0) {
    console.log(`Found ${checkedPlacements.length} unique checked placements to exclude`);
    if (websitePlacements.length > 0) {
      console.log(`  - Website placements: ${websitePlacements.length}`);
    }
    if (youtubeVideos.length > 0) {
      console.log(`  - YouTube videos: ${youtubeVideos.length}`);
    }
    if (mobileAppPlacements.length > 0) {
      console.log(`  - Mobile app placements: ${mobileAppPlacements.length}`);
    }
    if (DEBUG_MODE) {
      if (websitePlacements.length > 0) {
        console.log(`Website placements:`);
        const firstThree = websitePlacements.slice(0, 3);
        for (const placement of firstThree) {
          console.log(`  - ${placement}`);
        }
        if (websitePlacements.length > 3) {
          console.log(`  ... and ${websitePlacements.length - 3} more`);
        }
      }
      if (youtubeVideos.length > 0) {
        console.log(`YouTube videos:`);
        const firstThree = youtubeVideos.slice(0, 3);
        for (const placement of firstThree) {
          console.log(`  - ${placement}`);
        }
        if (youtubeVideos.length > 3) {
          console.log(`  ... and ${youtubeVideos.length - 3} more`);
        }
      }
      if (mobileAppPlacements.length > 0) {
        console.log(`Mobile app placements:`);
        const firstThree = mobileAppPlacements.slice(0, 3);
        for (const placement of firstThree) {
          console.log(`  - ${placement}`);
        }
        if (mobileAppPlacements.length > 3) {
          console.log(`  ... and ${mobileAppPlacements.length - 3} more`);
        }
      }
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
    youtubeVideos: youtubeVideos,
    mobileAppPlacements: mobileAppPlacements
  };
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

      if (DEBUG_MODE) {
        console.log(`Row appended for App ID: ${appId}`);
      }
    });

    // Execution Phase: Check if we're in preview mode or execution mode
    // Note: ExecutionMode.isPreview() returns true when script is run in preview mode
    try {
      if (typeof ExecutionMode !== 'undefined' && ExecutionMode.isPreview && ExecutionMode.isPreview()) {
        bulkUpload.preview();
        console.log('✓ Bulk upload job submitted for PREVIEW. Check Tools & Settings > Bulk actions > Uploads for results.');
      } else {
        bulkUpload.apply();
        console.log('✓ Bulk upload job APPLIED. Check Tools & Settings > Bulk actions > Uploads for results.');
      }
    } catch (modeError) {
      // Fallback: if ExecutionMode is not available, use preview by default for safety
      bulkUpload.preview();
      console.log('✓ Bulk upload job submitted for PREVIEW. Check Tools & Settings > Bulk actions > Uploads for results.');
      console.log('  Note: ExecutionMode not available, defaulting to preview mode for safety.');
    }
  } catch (error) {
    console.error(`✗ Error during bulk upload of mobile app exclusions: ${error.message}`);
    throw error;
  }
}

/**
 * Adds YouTube videos to the shared exclusion list individually
 * YouTube videos must be added one at a time using addExcludedPlacement() (singular)
 * @param {Array<string>} videoIds - Array of YouTube video IDs to exclude
 * @param {string} listName - The name of the shared exclusion list
 */
function addYouTubeVideosToExclusionList(videoIds, listName) {
  if (videoIds.length === 0) {
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
  const uniqueVideoIds = Array.from(new Set(videoIds));

  console.log(`Attempting to add ${uniqueVideoIds.length} unique YouTube videos to the shared list (individual processing)...`);

  let successCount = 0;
  let failureCount = 0;
  const failedVideos = [];

  for (const videoId of uniqueVideoIds) {
    try {
      excludedPlacementList.addExcludedPlacement(videoId);
      successCount++;
      if (DEBUG_MODE && successCount <= 3) {
        console.log(`✓ Added YouTube video: ${videoId}`);
      }
    } catch (error) {
      failureCount++;
      failedVideos.push({ videoId: videoId, error: error.message });
      if (DEBUG_MODE || failureCount <= 3) {
        console.error(`✗ Failed to add YouTube video ${videoId}: ${error.message}`);
      }
    }
  }

  console.log(`\n=== YouTube Video Exclusion Summary ===`);
  console.log(`Successfully added: ${successCount} videos`);
  if (failureCount > 0) {
    console.log(`Failed to add: ${failureCount} videos`);
    if (failedVideos.length <= 10) {
      console.log(`\nFailed videos:`);
      for (const failed of failedVideos) {
        console.log(`  - ${failed.videoId}: ${failed.error}`);
      }
    } else {
      console.log(`\nFirst 10 failed videos:`);
      for (let i = 0; i < 10; i++) {
        const failed = failedVideos[i];
        console.log(`  - ${failed.videoId}: ${failed.error}`);
      }
      console.log(`  ... and ${failedVideos.length - 10} more`);
    }
  }
  console.log(`=====================================\n`);
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
      console.log(`Failed to add: ${failureCount} placements`);
      if (failedPlacements.length <= 10) {
        console.log(`\nFailed placements:`);
        for (const failed of failedPlacements) {
          console.log(`  - ${failed.placement}: ${failed.error}`);
        }
      } else {
        console.log(`\nFirst 10 failed placements:`);
        for (let i = 0; i < 10; i++) {
          const failed = failedPlacements[i];
          console.log(`  - ${failed.placement}: ${failed.error}`);
        }
        console.log(`  ... and ${failedPlacements.length - 10} more`);
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
        if (DEBUG_MODE) {
          console.log(`✓ Linked exclusion list to campaign: ${campaignName} (ID: ${campaignId})`);
        }
      } else {
        campaignsAlreadyLinked++;
        if (DEBUG_MODE) {
          console.log(`- Campaign already linked: ${campaignName} (ID: ${campaignId})`);
        }
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
    console.log(`Failed to link: ${campaignsFailed} campaigns`);
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
    console.log(`Failed to link: ${displayCampaignsFailed} campaigns`);
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
        if (DEBUG_MODE) {
          console.log(`Attempt ${attempt + 1}: HTTP ${response.getResponseCode()} for ${url}`);
        }
      }
    } catch (error) {
      if (DEBUG_MODE) {
        console.log(`Attempt ${attempt + 1} failed for ${url}: ${error.message}`);
      }
    }

    attempt++;
    if (attempt < maxRetries) {
      const delay = baseDelay * Math.pow(2, attempt - 1); // Exponential backoff
      if (DEBUG_MODE) {
        console.log(`Retrying in ${delay}ms...`);
      }
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
        if (DEBUG_MODE) {
          console.log(`Attempt ${attempt + 1}: API error ${responseCode}: ${errorData.error?.message || responseText}`);
        }

        // Don't retry on authentication errors
        if (responseCode === 401) {
          console.error(`Authentication failed. Check your API key.`);
          return null;
        }
      }
    } catch (error) {
      if (DEBUG_MODE) {
        console.log(`Attempt ${attempt + 1} failed: ${error.message}`);
      }
    }

    attempt++;
    if (attempt < maxRetries) {
      const delay = baseDelay * Math.pow(2, attempt - 1); // Exponential backoff
      if (DEBUG_MODE) {
        console.log(`Retrying ChatGPT API call in ${delay}ms...`);
      }
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
 * Gets cached LLM response for a URL
 * @param {GoogleAppsScript.Spreadsheet.Sheet} llmSheet - The LLM responses sheet
 * @param {string} url - The URL to look up
 * @returns {string | null} The cached response or null if not found
 */
function getCachedLlmResponse(llmSheet, url) {
  const dataRange = llmSheet.getDataRange();
  const values = dataRange.getValues();

  // Headers are in row 1 (index 0), so data starts from row 2 (index 1)
  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    if (values[rowIndex][0] === url) {
      return values[rowIndex][1]; // Return the response
    }
  }

  return null;
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
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet for caching
 * @returns {string | null} The ChatGPT response or null if failed
 */
function getChatGptResponseForUrl(url, placement, settings, spreadsheet) {
  if (!settings.enableChatGpt || !settings.chatGptApiKey) {
    return null;
  }

  // Check cache first if enabled
  if (settings.useCachedChatGpt) {
    const llmSheet = getOrCreateLlmResponsesSheet(spreadsheet);
    const cachedResponse = getCachedLlmResponse(llmSheet, url);

    if (cachedResponse) {
      if (DEBUG_MODE) {
        console.log(`Using cached response for ${url}`);
      }
      // Apply response filters to cached response (cache stores unfiltered response)
      const filteredResponse = filterChatGptResponse(cachedResponse, settings);
      return filteredResponse;
    }
  }

  // Validate URL before fetching
  if (!isValidUrl(url)) {
    console.log(`  ⚠ ChatGPT: Invalid URL, skipping: ${url}`);
    return null;
  }

  // Fetch website content
  const fullUrl = url.startsWith('http') ? url : `https://${url}`;
  if (DEBUG_MODE) {
    console.log(`Fetching website content for ${fullUrl}...`);
  }

  const websiteContent = fetchWebsiteContentWithRetry(fullUrl);

  if (!websiteContent) {
    console.error(`Failed to fetch content for ${url}`);
    return null;
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
  if (DEBUG_MODE) {
    console.log(`Calling ChatGPT API for ${url}...`);
  }

  const response = callChatGptApiWithRetry(settings.chatGptApiKey, fullPrompt);

  if (response) {
    // Cache the response BEFORE filtering (as per requirements)
    const llmSheet = getOrCreateLlmResponsesSheet(spreadsheet);
    cacheLlmResponse(llmSheet, url, response);

    // Apply response filters after caching
    const filteredResponse = filterChatGptResponse(response, settings);
    return filteredResponse;
  }

  return null;
}

/**
 * Filters ChatGPT response based on Response Contains and Response Not Contains filters
 * @param {string} response - The ChatGPT response to filter
 * @param {Object} settings - Settings object with response filter configuration
 * @returns {string} The filtered response (empty string if filtered out, original response if passes filters)
 */
function filterChatGptResponse(response, settings) {
  if (!response || typeof response !== 'string') {
    return response || '';
  }

  const responseLower = response.toLowerCase();

  // Check Response Contains filter (only if list has items)
  if (settings.responseContainsList && settings.responseContainsList.length > 0) {
    const containsMatch = settings.responseContainsList.some(filterString => {
      return responseLower.includes(filterString.toLowerCase());
    });

    if (!containsMatch) {
      if (DEBUG_MODE) {
        console.log(`  Response filtered out (does not contain any required strings)`);
      }
      return ''; // Filtered out
    }
  }

  // Note: Response Not Contains filters are NOT used for filtering responses out
  // They are informational only (responses matching these will still appear in the report)

  return response; // Passes all filters
}


