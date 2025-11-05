/**
 * Placement Exclusions - Semi-automated
 * Grabs placement data from Performance Max and Display campaigns, writes to sheet with checkboxes for user selection,
 * then adds selected placements to a shared exclusion list using batch processing
 * Includes optional ChatGPT integration for website content analysis
 * Version: 1.4.0
 */

// Google Ads API Query Builder Links:
// Performance Max Placement View: https://developers.google.com/google-ads/api/fields/v20/performance_max_placement_view_query_builder
// Display Placement View: https://developers.google.com/google-ads/api/fields/v20/detail_placement_view_query_builder

// --- Configuration ---
const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1OnuQEg-cUHx4cUfcpGsT6zc4TPsNKh5l1IjtCsOaGOw/edit?gid=0#gid=0';
// The Google Sheet URL where placement data will be written
// Create a Sheet (sheets.new) and paste the URL here
// Example: https://docs.google.com/spreadsheets/d/1abc123def456/edit

const SHARED_EXCLUSION_LIST_NAME = 'PMax Placement Semi-automated Exclusions';
// The name of the shared placement exclusion list that must exist in your Google Ads account
// Create this at: Tools & Settings > Shared Library > Placement exclusions

const SETTINGS_SHEET_NAME = 'Settings';
// The name of the sheet tab that contains configuration settings
// This sheet will be auto-populated on first run if it doesn't exist

const DATA_SHEET_NAME = 'Placements';
// The name of the sheet tab where placement data with checkboxes will be written

const LLM_RESPONSES_SHEET_NAME = 'LLM Responses';
// The name of the sheet tab that caches ChatGPT responses by URL

const MAX_RESULTS_DEFAULT_VALUE = 5;
// Default maximum number of results to show when Max Results setting is empty
// Set to 0 to ignore (no limit) - but this is handled in the settings reading logic

const DEBUG_MODE = true;
// Set to true to see detailed logs for debugging
// Core logs will always appear regardless of this setting

const GOOGLE_DOMAIN_EXCLUSIONS = [
  'youtube.com',
  'mail.google.com',
  'google.com',
  'gmail.com',
  'googleusercontent.com'
];
// Domains that Google prohibits excluding due to policy
// These will appear in reports with a note but cannot be excluded even if checked

// --- Main Function ---
function main() {
  console.log(`Script started`);

  validateConfig();

  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const settingsSheet = getOrCreateSettingsSheet(spreadsheet);
  const listsSheet = getOrCreateListsSheet(spreadsheet);

  // Default settings for initial population
  const defaultSettings = {
    lookbackWindowDays: 30,
    minimumImpressions: 0,
    minimumClicks: 0,
    enabledCampaignsOnly: false
  };

  populateSettingsSheetIfEmpty(settingsSheet, defaultSettings);
  populateListsSheetIfEmpty(listsSheet);

  const settings = getSettingsFromSheet(settingsSheet, listsSheet);

  validateSharedExclusionList();

  const placementData = getPlacementData(settings);
  console.log(`Found ${placementData.length} placements`);

  if (placementData.length === 0) {
    console.log(`No placement data found. Check your filters and date range.`);
    return;
  }

  // Read checked placements BEFORE writing new data (so we don't lose user's selections)
  const checkedPlacements = readCheckedPlacementsFromSheet(spreadsheet);

  const transformedData = calculatePlacementMetrics(placementData);

  // Filter data first based on thresholds
  const filteredData = filterPlacementData(transformedData, settings);

  // Max results limit is applied in the GAQL query (LIMIT clause)
  // Note: There may be more results than maxResults because PMax and Display are separate queries
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
        // Only get ChatGPT responses for website placements
        if (placement.placementType && !String(placement.placementType).toUpperCase().includes('MOBILE_APPLI')) {
          const url = placement.targetUrl || placement.placement;

          // Check if response is already cached
          const wasCachedBefore = getCachedLlmResponse(llmSheet, url);
          const chatGptResponse = getChatGptResponseForUrl(url, placement, settings, spreadsheet);

          if (chatGptResponse) {
            placement.chatGptResponse = chatGptResponse;
            processedCount++;

            if (wasCachedBefore) {
              cachedCount++;
            }
          } else {
            placement.chatGptResponse = ''; // Failed to get response
            failedCount++;
          }
        } else {
          placement.chatGptResponse = ''; // Mobile apps don't get ChatGPT responses
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
  const hasMobileAppPlacements = checkedPlacements.mobileAppPlacements && checkedPlacements.mobileAppPlacements.length > 0;

  if (hasWebsitePlacements || hasMobileAppPlacements) {
    // Add website placements using the standard method
    if (hasWebsitePlacements) {
      addPlacementsToExclusionList(checkedPlacements.websitePlacements);
    }

    // Add mobile app placements using Bulk Upload
    if (hasMobileAppPlacements) {
      addAppExclusionsViaBulkUpload(SHARED_EXCLUSION_LIST_NAME, checkedPlacements.mobileAppPlacements);
    }

    linkSharedListToPMaxCampaigns();
  } else {
    console.log(`No placements selected for exclusion. Check boxes in column A of the Placements sheet to exclude placements, then run the script again.`);
    // Still ensure the list is linked to campaigns even if no new placements added
    linkSharedListToPMaxCampaigns();
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
 */
function validateSharedExclusionList() {
  const listIterator = AdsApp.excludedPlacementLists()
    .withCondition(`Name = '${SHARED_EXCLUSION_LIST_NAME}'`)
    .get();

  if (!listIterator.hasNext()) {
    const errorMessage = `ERROR: Shared Placement Exclusion List named '${SHARED_EXCLUSION_LIST_NAME}' not found.\n\n` +
      `Please create this shared list in your Google Ads account first:\n` +
      `1. Go to Tools & Settings > Shared Library > Placement exclusions\n` +
      `2. Create a new list with the exact name: "${SHARED_EXCLUSION_LIST_NAME}"\n` +
      `3. Ensure the list is enabled\n` +
      `4. Link this list to your Performance Max campaigns if not already linked`;
    throw new Error(errorMessage);
  }

  const list = listIterator.next();
  console.log(`✓ Found shared exclusion list: ${SHARED_EXCLUSION_LIST_NAME}`);
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
  const LISTS_SHEET_NAME = 'Lists';
  let listsSheet = spreadsheet.getSheetByName(LISTS_SHEET_NAME);
  if (!listsSheet) {
    listsSheet = spreadsheet.insertSheet(LISTS_SHEET_NAME);
    console.log(`Created lists sheet: ${LISTS_SHEET_NAME}`);
  }
  return listsSheet;
}

/**
 * Gets settings from the settings sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet - The settings sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The lists sheet
 * @returns {Object} Settings object with all configuration values
 */
function getSettingsFromSheet(settingsSheet, listsSheet) {
  const lookbackWindowDaysRaw = getSettingValue(settingsSheet, 'Lookback Window (Days)', 30);
  const minimumImpressionsRaw = getSettingValue(settingsSheet, 'Minimum Impressions', 0);
  const minimumClicksRaw = getSettingValue(settingsSheet, 'Minimum Clicks', 0);
  const maxResultsRaw = getSettingValue(settingsSheet, 'Max Results', MAX_RESULTS_DEFAULT_VALUE);
  const enabledCampaignsOnlyRaw = getSettingValue(settingsSheet, 'Enabled campaigns only', false);

  // Handle checkbox value - Google Sheets checkboxes return boolean true/false
  // Default to false if not found or invalid
  let enabledCampaignsOnly = false;
  if (typeof enabledCampaignsOnlyRaw === 'boolean') {
    enabledCampaignsOnly = enabledCampaignsOnlyRaw;
  } else if (typeof enabledCampaignsOnlyRaw === 'string') {
    const lowerValue = enabledCampaignsOnlyRaw.toLowerCase().trim();
    enabledCampaignsOnly = lowerValue === 'true' || lowerValue === '1';
  } else if (enabledCampaignsOnlyRaw === null || enabledCampaignsOnlyRaw === undefined || enabledCampaignsOnlyRaw === '') {
    enabledCampaignsOnly = false; // Default to false
  }

  // ChatGPT settings
  const enableChatGptRaw = getSettingValue(settingsSheet, 'Enable ChatGPT', false);
  let enableChatGpt = false;
  if (typeof enableChatGptRaw === 'boolean') {
    enableChatGpt = enableChatGptRaw;
  } else if (typeof enableChatGptRaw === 'string') {
    const lowerValue = enableChatGptRaw.toLowerCase().trim();
    enableChatGpt = lowerValue === 'true' || lowerValue === '1';
  }

  const chatGptApiKey = String(getSettingValue(settingsSheet, 'ChatGPT API Key', '') || '').trim();
  const useCachedChatGptRaw = getSettingValue(settingsSheet, 'Use Cached ChatGPT Responses', true);
  let useCachedChatGpt = true;
  if (typeof useCachedChatGptRaw === 'boolean') {
    useCachedChatGpt = useCachedChatGptRaw;
  } else if (typeof useCachedChatGptRaw === 'string') {
    const lowerValue = useCachedChatGptRaw.toLowerCase().trim();
    useCachedChatGpt = lowerValue === 'true' || lowerValue === '1';
  }


  const defaultChatGptPrompt = `
  You are a Google Ads professional. You will be given a display (content network) placement and its website content.
  Your task is to determine if the placement should be excluded from the Google Ads account.
  You should exclude:
  Inappropriate or Offensive Content: Exclude placements (websites, videos, apps) that contain content that is explicit, violent, hateful, illegal, or politically sensitive and does not align with your brand's values.
  Irrelevant or Low-Quality Sites: Exclude domains that are irrelevant to the target audience or business niche, as well as those that are visibly low-quality, poorly designed, or have minimal actual content (e.g., parked domains).
  User-Generated Content (UGC) Risk: Exclude placements on unmoderated forums, comment sections, or UGC-heavy sites where the surrounding content could potentially damage your brand reputation.
  Note the placement doesn't have to be directly related to the type of campaign. The Guardian website should be included for example
  as a respected publisher and likely higher-than-average income readership.
  Your response will be in the following format:
  {
    "action": "exclude" or "include",
    "reason": "reason for the action",
    "notes": "any additional notes"
  }
  Example response:
  {
    "action": "exclude",
    "reason": "The placement is not legitimate and should be excluded.",
    "notes": "The placement is innappropriate for the campaign"
  }
  `;
  const chatGptPrompt = String(getSettingValue(settingsSheet, 'ChatGPT Prompt', defaultChatGptPrompt) || '').trim();

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

  // Handle Max Results: if 0, ignore (no limit), if empty use default, otherwise use the value
  let maxResults = MAX_RESULTS_DEFAULT_VALUE;
  if (maxResultsRaw !== null && maxResultsRaw !== undefined && maxResultsRaw !== '') {
    const maxResultsNum = typeof maxResultsRaw === 'number' ? maxResultsRaw : parseInt(maxResultsRaw, 10);
    if (!isNaN(maxResultsNum)) {
      if (maxResultsNum === 0) {
        maxResults = 0; // 0 means no limit (ignore)
      } else {
        maxResults = maxResultsNum;
      }
    }
  }

  const settings = {
    lookbackWindowDays: typeof lookbackWindowDaysRaw === 'number' ? lookbackWindowDaysRaw : parseInt(lookbackWindowDaysRaw, 10) || 30,
    minimumImpressions: typeof minimumImpressionsRaw === 'number' ? minimumImpressionsRaw : parseInt(minimumImpressionsRaw, 10) || 0,
    minimumClicks: typeof minimumClicksRaw === 'number' ? minimumClicksRaw : parseInt(minimumClicksRaw, 10) || 0,
    maxResults: maxResults,
    campaignNameContains: String(getSettingValue(settingsSheet, 'Campaign Name Contains', '') || '').trim(),
    campaignNameNotContains: String(getSettingValue(settingsSheet, 'Campaign Name Not Contains', '') || '').trim(),
    enabledCampaignsOnly: enabledCampaignsOnly,
    enableChatGpt: enableChatGpt,
    chatGptApiKey: chatGptApiKey,
    useCachedChatGpt: useCachedChatGpt,
    chatGptPrompt: chatGptPrompt,
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
    console.log(`  Minimum Impressions: ${settings.minimumImpressions}`);
    console.log(`  Minimum Clicks: ${settings.minimumClicks}`);
    console.log(`  Max Results: ${settings.maxResults === 0 ? 'No limit' : settings.maxResults}`);
    console.log(`  Campaign Name Contains: "${settings.campaignNameContains}"`);
    console.log(`  Campaign Name Not Contains: "${settings.campaignNameNotContains}"`);
    console.log(`  Enabled campaigns only: ${settings.enabledCampaignsOnly}`);
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
    console.log(`  Placement Not Contains List (informational): ${settings.placementNotContainsList.length} strings`);
    console.log(`  Display Name Contains Filter: ${settings.displayNameContainsEnabled ? 'Enabled' : 'Disabled'}`);
    if (settings.displayNameContainsEnabled && settings.displayNameContainsList.length > 0) {
      console.log(`    List: ${settings.displayNameContainsList.length} strings`);
      if (settings.displayNameContainsList.length <= 10) {
        console.log(`      ${settings.displayNameContainsList.join(', ')}`);
      } else {
        console.log(`      ${settings.displayNameContainsList.slice(0, 10).join(', ')} ... and ${settings.displayNameContainsList.length - 10} more`);
      }
    }
    console.log(`  Display Name Not Contains List (informational): ${settings.displayNameNotContainsList.length} strings`);
    console.log(`  Target URL Contains Filter: ${settings.targetUrlContainsEnabled ? 'Enabled' : 'Disabled'}`);
    if (settings.targetUrlContainsEnabled && settings.targetUrlContainsList.length > 0) {
      console.log(`    List: ${settings.targetUrlContainsList.length} strings`);
      if (settings.targetUrlContainsList.length <= 10) {
        console.log(`      ${settings.targetUrlContainsList.join(', ')}`);
      } else {
        console.log(`      ${settings.targetUrlContainsList.slice(0, 10).join(', ')} ... and ${settings.targetUrlContainsList.length - 10} more`);
      }
    }
    console.log(`  Target URL Not Contains List (informational): ${settings.targetUrlNotContainsList.length} strings`);
    console.log(`  Target URL Ends With Filter: ${settings.targetUrlEndsWithEnabled ? 'Enabled' : 'Disabled'}`);
    if (settings.targetUrlEndsWithEnabled && settings.targetUrlEndsWithList.length > 0) {
      console.log(`    List: ${settings.targetUrlEndsWithList.length} strings`);
      if (settings.targetUrlEndsWithList.length <= 10) {
        console.log(`      ${settings.targetUrlEndsWithList.join(', ')}`);
      } else {
        console.log(`      ${settings.targetUrlEndsWithList.slice(0, 10).join(', ')} ... and ${settings.targetUrlEndsWithList.length - 10} more`);
      }
    }
    console.log(`  Target URL Not Ends With List (informational): ${settings.targetUrlNotEndsWithList.length} strings`);
    console.log(`  Placement Type Filters:`);
    console.log(`    YouTube Video: ${settings.placementTypeFilters.youtubeVideo}`);
    console.log(`    Website: ${settings.placementTypeFilters.website}`);
    console.log(`    Mobile Application: ${settings.placementTypeFilters.mobileApplication}`);
  }

  return settings;
}

/**
 * Gets a setting value from the settings sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet - The settings sheet
 * @param {string} settingName - The name of the setting to find
 * @param {*} defaultValue - Default value if setting not found
 * @returns {*} The setting value or default
 */
function getSettingValue(settingsSheet, settingName, defaultValue) {
  const dataRange = settingsSheet.getDataRange();
  const values = dataRange.getValues();

  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    if (values[rowIndex][0] === settingName) {
      const value = values[rowIndex][1];
      if (value !== null && value !== undefined && value !== '') {
        return value;
      }
    }
  }

  return defaultValue;
}

/**
 * Gets the Placement Contains list and enabled status from the Lists sheet (checks placement field)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getPlacementContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'Placement Contains', 1); // Column A
}

/**
 * Gets the Placement Not Contains list and enabled status from the Lists sheet (checks placement field, informational)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getPlacementNotContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'Placement Not Contains', 2); // Column B
}

/**
 * Gets the Display Name Contains list and enabled status from the Lists sheet (checks displayName field)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getDisplayNameContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'Display Name Contains', 3); // Column C
}

/**
 * Gets the Display Name Not Contains list and enabled status from the Lists sheet (checks displayName field, informational)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getDisplayNameNotContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'Display Name Not Contains', 4); // Column D
}

/**
 * Gets the Target URL Contains list and enabled status from the Lists sheet (checks targetUrl field)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getTargetUrlContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'Target URL Contains', 5); // Column E
}

/**
 * Gets the Target URL Not Contains list and enabled status from the Lists sheet (checks targetUrl field, informational)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getTargetUrlNotContainsList(listsSheet) {
  return getListFromSheet(listsSheet, 'Target URL Not Contains', 6); // Column F
}

/**
 * Gets the Target URL Ends With list and enabled status from the Lists sheet (checks targetUrl field)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getTargetUrlEndsWithList(listsSheet) {
  return getListFromSheet(listsSheet, 'Target URL Ends With', 7); // Column G
}

/**
 * Gets the Target URL Not Ends With list and enabled status from the Lists sheet (checks targetUrl field, informational)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @returns {Object} Object with enabled status and list array
 */
function getTargetUrlNotEndsWithList(listsSheet) {
  return getListFromSheet(listsSheet, 'Target URL Not Ends With', 8); // Column H
}

/**
 * Generic function to get a list and enabled status from the Lists sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @param {string} listName - The name of the list header
 * @param {number} columnIndex - The column index (1 = A, 2 = B, etc.)
 * @returns {Object} Object with enabled status and list array {enabled: boolean, list: Array<string>}
 */
function getListFromSheet(listsSheet, listName, columnIndex) {
  const list = [];
  let enabled = true; // Default to enabled
  const dataRange = listsSheet.getDataRange();
  const values = dataRange.getValues();
  const columnArrayIndex = columnIndex - 1; // Convert to 0-based array index
  // Checkbox is in the same column as the list (Column A list -> checkbox in A, Column B list -> checkbox in B)
  const checkboxColumnArrayIndex = columnArrayIndex; // Same column as the list

  // Find the header row in the specified column
  let headerRowIndex = -1;
  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    if (values[rowIndex][columnArrayIndex] === listName) {
      headerRowIndex = rowIndex;
      break;
    }
  }

  if (headerRowIndex === -1) {
    if (DEBUG_MODE) {
      console.log(`  ${listName} header not found in Lists sheet column ${columnIndex}`);
    }
    return { enabled: false, list: [] }; // Return disabled and empty list if header not found
  }

  // Enable checkbox is 3 rows after header (header + description + enable label row)
  // Checkbox is in row 4 (header=1, description=2, label=3, checkbox=4)
  const enableCheckboxRowIndex = headerRowIndex + 3;
  // List starts 4 rows after the header (header + description + enable label + checkbox)
  const listStartRowIndex = headerRowIndex + 4;

  // Read the enable checkbox value
  if (enableCheckboxRowIndex < values.length) {
    const enableCheckboxValue = values[enableCheckboxRowIndex][checkboxColumnArrayIndex];
    if (typeof enableCheckboxValue === 'boolean') {
      enabled = enableCheckboxValue;
    } else if (typeof enableCheckboxValue === 'string') {
      const lowerValue = enableCheckboxValue.toLowerCase().trim();
      enabled = lowerValue === 'true' || lowerValue === '1';
    }
  }

  if (DEBUG_MODE) {
    console.log(`  Found ${listName} header at row ${headerRowIndex + 1}, column ${columnIndex}`);
    console.log(`  Enable checkbox at row ${enableCheckboxRowIndex + 1}, column ${columnIndex}: ${enabled}`);
    console.log(`  Starting to read list from row ${listStartRowIndex + 1}`);
  }

  // Read from the specified column starting from listStartRowIndex
  for (let rowIndex = listStartRowIndex; rowIndex < values.length; rowIndex++) {
    const cellValue = values[rowIndex][columnArrayIndex];
    const item = String(cellValue || '').trim();

    // Stop if we hit an empty row
    if (!item) {
      break;
    }

    // Skip the header itself if it somehow appears again
    if (item === listName) {
      continue;
    }

    // Skip description rows (they contain "Description" or "Only include" or "Exclude")
    if (item.toLowerCase().includes('description') ||
      item.toLowerCase().includes('only include') ||
      item.toLowerCase().includes('exclude placements')) {
      continue;
    }

    // Skip enable checkbox row
    if (item.toLowerCase().includes('enable')) {
      continue;
    }

    if (item) {
      list.push(item);
    }
  }

  if (DEBUG_MODE) {
    console.log(`  ${listName} read: ${list.length} items found, enabled: ${enabled}`);
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
  const dataRange = settingsSheet.getDataRange();
  const values = dataRange.getValues();

  // Find the "Placement Type Filters" header row
  let headerRowIndex = -1;
  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    if (values[rowIndex][0] === 'Placement Type Filters') {
      headerRowIndex = rowIndex;
      break;
    }
  }

  const filters = {
    youtubeVideo: false,
    website: false,
    mobileApplication: false
  };

  if (headerRowIndex === -1) {
    // Default to all true if section not found
    filters.youtubeVideo = true;
    filters.website = true;
    filters.mobileApplication = true;
    return filters;
  }

  // Read placement type checkboxes (rows after header)
  for (let rowIndex = headerRowIndex + 1; rowIndex < values.length && rowIndex < headerRowIndex + 4; rowIndex++) {
    const row = values[rowIndex];
    const settingName = String(row[0] || '').trim();
    const value = row[1];

    let isChecked = false;
    if (typeof value === 'boolean') {
      isChecked = value;
    } else if (typeof value === 'string') {
      const lowerValue = value.toLowerCase().trim();
      isChecked = lowerValue === 'true' || lowerValue === '1';
    }

    if (settingName === 'YouTube Video') {
      filters.youtubeVideo = isChecked;
    } else if (settingName === 'Website') {
      filters.website = isChecked;
    } else if (settingName === 'Mobile Application') {
      filters.mobileApplication = isChecked;
    }
  }

  // If no filters are checked, default to all true
  if (!filters.youtubeVideo && !filters.website && !filters.mobileApplication) {
    filters.youtubeVideo = true;
    filters.website = true;
    filters.mobileApplication = true;
  }

  return filters;
}

/**
 * Populates the Lists sheet with all filter lists if it's empty
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 */
function populateListsSheetIfEmpty(listsSheet) {
  const dataRange = listsSheet.getDataRange();
  const existingValues = dataRange.getValues();

  // Check if sheet already has data (more than just headers)
  if (existingValues.length > 1) {
    if (DEBUG_MODE) {
      console.log(`Lists sheet already populated, skipping initialization`);
    }
    return;
  }

  console.log(`Populating Lists sheet with all filter lists...`);

  // Clear the sheet first
  listsSheet.clear();

  // Helper function to populate a list column
  function populateListColumn(columnIndex, headerName, description, isNotContains = false) {
    let row = 1;

    // Header
    listsSheet.getRange(row, columnIndex).setValue(headerName);
    listsSheet.getRange(row, columnIndex).setFontWeight('bold');
    listsSheet.getRange(row, columnIndex).setBackground(isNotContains ? '#fce8e6' : '#e8f0fe');
    row++;

    // Description
    listsSheet.getRange(row, columnIndex).setValue(description);
    listsSheet.getRange(row, columnIndex).setFontStyle('italic');
    listsSheet.getRange(row, columnIndex).setFontColor('#666666');
    row++;

    // Enable checkbox label
    listsSheet.getRange(row, columnIndex).setValue(`Enable ${headerName} filter:`);
    row++;

    // Checkbox
    listsSheet.getRange(row, columnIndex).insertCheckboxes();
    listsSheet.getRange(row, columnIndex).setValue(true); // Default to enabled
    row++;

    return row; // Return starting row for list items
  }

  // === Column A: Placement Contains (checks placement field) ===
  const placementContainsRow = populateListColumn(1, 'Placement Contains', 'Only include placements where the Placement field contains at least one of the strings below');
  const placementContainsDefaults = [];
  addDefaultsToList(listsSheet, placementContainsRow, 1, placementContainsDefaults);

  // === Column B: Placement Not Contains (checks placement field, informational) ===
  const placementNotContainsRow = populateListColumn(2, 'Placement Not Contains', 'Placements where the Placement field contains these strings will still appear in the report for your review', true);
  const placementNotContainsDefaults = [];
  addDefaultsToList(listsSheet, placementNotContainsRow, 2, placementNotContainsDefaults);

  // === Column C: Display Name Contains (checks displayName field) ===
  const displayNameContainsRow = populateListColumn(3, 'Display Name Contains', 'Only include placements where the Display Name field contains at least one of the strings below');
  const displayNameContainsDefaults = [];
  addDefaultsToList(listsSheet, displayNameContainsRow, 3, displayNameContainsDefaults);

  // === Column D: Display Name Not Contains (checks displayName field, informational) ===
  const displayNameNotContainsRow = populateListColumn(4, 'Display Name Not Contains', 'Placements where the Display Name field contains these strings will still appear in the report for your review', true);
  const displayNameNotContainsDefaults = [];
  addDefaultsToList(listsSheet, displayNameNotContainsRow, 4, displayNameNotContainsDefaults);

  // === Column E: Target URL Contains (checks targetUrl field) ===
  const targetUrlContainsRow = populateListColumn(5, 'Target URL Contains', 'Only include placements where the Target URL field contains at least one of the strings below');
  const targetUrlContainsDefaults = [];
  addDefaultsToList(listsSheet, targetUrlContainsRow, 5, targetUrlContainsDefaults);

  // === Column F: Target URL Not Contains (checks targetUrl field, informational) ===
  const targetUrlNotContainsRow = populateListColumn(6, 'Target URL Not Contains', 'Placements where the Target URL field contains these strings will still appear in the report for your review', true);
  const targetUrlNotContainsDefaults = [
    'gambling',
    'casino',
    'adult',
    'porn',
    'sex',
    'dating',
    'escorts',
    'kids',
    'child',
    'game',
    'children',
    "free",
    "cash",
    "money",
    "income",
    "earn",
    "profit",
    "billion",
    "dollars",
    "wealth",
    "credit",
    "loan",
    "investment",
    "refinance",
    "guaranteed",
    "risk-free",
    "no fees",
    "no catch",
    "act now",
    "urgent",
    "limited time",
    "expires",
    "deadline",
    "instant",
    "immediate",
    "exclusive",
    "special offer",
    "buy now",
    "order now",
    "click here",
    "unsubscribe",
    "deal",
    "discount",
    "prize",
    "winner",
    "congratulations",
    "miracle",
    "amazing",
    "fantastic",
    "incredible",
    "breakthrough",
    "cure",
    "treatment",
    "weight loss",
    "viagra",
    "pharmacy",
    "herbal",
    "adult",
    "sex",
    "xxx",
    "password",
    "secret",
    "confidential",
    "dear friend",
    "selected",
    "eliminate debt",
    "be your own boss",
    "work from home",
    "as seen on"
  ];
  addDefaultsToList(listsSheet, targetUrlNotContainsRow, 6, targetUrlNotContainsDefaults);

  // === Column G: Target URL Ends With (checks targetUrl field) ===
  const targetUrlEndsWithRow = populateListColumn(7, 'Target URL Ends With', 'Only include placements where the Target URL field ends with one of the strings below');
  const targetUrlEndsWithDefaults = [

  ];
  addDefaultsToList(listsSheet, targetUrlEndsWithRow, 7, targetUrlEndsWithDefaults);

  // === Column H: Target URL Not Ends With (checks targetUrl field, informational) ===
  const targetUrlNotEndsWithRow = populateListColumn(8, 'Target URL Not Ends With', 'Placements where the Target URL field ends with these strings will still appear in the report for your review', true);
  const targetUrlNotEndsWithDefaults = [
    ".ad",
    ".ae",
    ".af",
    ".al",
    ".am",
    ".ar",
    ".at",
    ".az",
    ".ba",
    ".be",
    ".bg",
    ".bh",
    ".bi",
    ".bj",
    ".bn",
    ".bo",
    ".br",
    ".by",
    ".cl",
    ".cn",
    ".co",
    ".cr",
    ".cu",
    ".cz",
    ".de",
    ".dk",
    ".do",
    ".dz",
    ".ec",
    ".ee",
    ".eg",
    ".es",
    ".et",
    ".fi",
    ".fr",
    ".ge",
    ".gr",
    ".gt",
    ".hk",
    ".hr",
    ".hu",
    ".id",
    ".ie",
    ".il",
    ".in",
    ".ir",
    ".is",
    ".it",
    ".jp",
    ".kr",
    ".kz",
    ".la",
    ".lb",
    ".li",
    ".lk",
    ".lt",
    ".lu",
    ".lv",
    ".ly",
    ".ma",
    ".md",
    ".mx",
    ".my",
    ".ng",
    ".nl",
    ".no",
    ".np",
    ".pe",
    ".ph",
    ".pk",
    ".pl",
    ".pt",
    ".ro",
    ".rs",
    ".ru",
    ".sa",
    ".se",
    ".sg",
    ".si",
    ".sk",
    ".sy",
    ".th",
    ".tn",
    ".tr",
    ".tw",
    ".ua",
    ".uy",
    ".uz",
    ".ve",
    ".vn",
    ".za",
    ".zm",
    ".zw",
  ];
  addDefaultsToList(listsSheet, targetUrlNotEndsWithRow, 8, targetUrlNotEndsWithDefaults);

  // Freeze header row
  listsSheet.setFrozenRows(1);

  console.log(`Lists sheet populated successfully`);
}

/**
 * Adds default values to a list column in the Lists sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listsSheet - The Lists sheet
 * @param {number} startRow - The starting row index (1-based) for list items (after checkbox)
 * @param {number} columnIndex - The column index (1 = A, 2 = B, etc.)
 * @param {Array<string>} defaults - Array of default values to add
 */
function addDefaultsToList(listsSheet, startRow, columnIndex, defaults) {
  for (let i = 0; i < defaults.length; i++) {
    listsSheet.getRange(startRow, columnIndex).setValue(defaults[i]);
    listsSheet.getRange(startRow, columnIndex).setFontColor('#999999');
    startRow++;
  }
}

/**
 * Extracts the TLD from a URL
 * @param {string} url - The URL to extract TLD from
 * @returns {string | null} The TLD (with leading dot) or null if can't extract
 */
function extractTld(url) {
  if (!url || typeof url !== 'string') {
    return null;
  }

  // Remove protocol if present
  let cleanUrl = url.replace(/^https?:\/\//i, '');

  // Remove www. if present
  cleanUrl = cleanUrl.replace(/^www\./i, '');

  // Remove path and query parameters
  cleanUrl = cleanUrl.split('/')[0];
  cleanUrl = cleanUrl.split('?')[0];

  // Split by dots and get the last parts (TLD)
  const parts = cleanUrl.split('.');

  if (parts.length < 2) {
    return null; // Not a valid domain
  }

  // Try to get the TLD (last part or last two parts for country codes like .co.uk)
  // Common two-part TLDs: .co.uk, .com.au, .co.nz, etc.
  const twoPartTlds = ['.co.uk', '.com.au', '.co.nz', '.co.za', '.com.br', '.co.jp', '.com.cn', '.co.in', '.com.mx', '.com.ar', '.com.co'];

  if (parts.length >= 2) {
    const lastTwo = '.' + parts[parts.length - 2] + '.' + parts[parts.length - 1];
    if (twoPartTlds.includes(lastTwo.toLowerCase())) {
      return lastTwo.toLowerCase();
    }
  }

  // Single-part TLD
  const tld = '.' + parts[parts.length - 1];
  return tld.toLowerCase();
}

/**
 * Populates the settings sheet with labels and default values if it's empty
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet - The settings sheet
 * @param {Object} settings - Current settings object
 */
function populateSettingsSheetIfEmpty(settingsSheet, settings) {
  const dataRange = settingsSheet.getDataRange();
  const existingValues = dataRange.getValues();

  // Check if sheet already has data (more than just headers)
  if (existingValues.length > 1) {
    if (DEBUG_MODE) {
      console.log(`Settings sheet already populated, skipping initialization`);
    }
    return;
  }

  console.log(`Populating settings sheet with default values...`);

  // Clear the sheet first
  settingsSheet.clear();

  // Define settings structure - leave checkbox value empty for now, will set after inserting checkbox
  const defaultChatGptPrompt = `You are a Google Ads professional. You will be given a display (content network) placement and its website content.
  Your task is to determine if the placement is legitimate and should be excluded from the Google Ads account.
  Your response will be in the following format:
  {
    "action": "exclude" or "include",
    "reason": "reason for the action",
    "notes": "any additional notes"
  }
  Example response:
  {
    "action": "exclude",
    "reason": "The placement is not legitimate and should be excluded.",
    "notes": "The placement is innappropriate for the campaign"
  }`;

  const settingsStructure = [
    ['Setting', 'Value', 'Description'],
    ['Shared Exclusion List Name', SHARED_EXCLUSION_LIST_NAME, 'The name of the shared placement exclusion list that must exist in your Google Ads account'],
    ['Lookback Window (Days)', settings.lookbackWindowDays, 'Number of days to look back from today (e.g., 30)'],
    ['Minimum Impressions', settings.minimumImpressions, 'Only show placements with at least this many impressions'],
    ['Minimum Clicks', settings.minimumClicks, 'Only show placements with at least this many clicks'],
    ['Max Results', MAX_RESULTS_DEFAULT_VALUE, 'Maximum number of placements per query (0 = no limit). Note: Results may exceed this as PMax and Display are separate queries.'],
    ['Campaign Name Contains', settings.campaignNameContains, 'Filter campaigns by name containing this text (leave empty for all)'],
    ['Campaign Name Not Contains', settings.campaignNameNotContains, 'Exclude campaigns with names containing this text (leave empty for none)'],
    ['Enabled campaigns only', '', 'If checked, only include enabled campaigns. If unchecked, include paused campaigns (removed campaigns are always excluded)'],
    ['Enable ChatGPT', '', 'If checked, ChatGPT will analyze placement websites and provide summaries'],
    ['ChatGPT API Key', '', 'Your OpenAI API key (required if ChatGPT is enabled)'],
    ['Use Cached ChatGPT Responses', '', 'If checked, will use cached responses from previous runs to avoid redundant API calls'],
    ['ChatGPT Prompt', defaultChatGptPrompt, 'The prompt to send to ChatGPT. Website content and placement info will be appended automatically.']
  ];

  // Write settings structure
  const numRows = settingsStructure.length;
  const numCols = settingsStructure[0].length;
  settingsSheet.getRange(1, 1, numRows, numCols).setValues(settingsStructure);

  // Format header row
  const headerRange = settingsSheet.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  // Format description column
  settingsSheet.getRange(2, 3, numRows - 1, 1).setFontStyle('italic');
  settingsSheet.getRange(2, 3, numRows - 1, 1).setFontColor('#666666');

  // Add checkboxes to settings that need them
  // Row 9: Enabled campaigns only (index 9 in structure array + 1 for header = row 10, but array index 8 = row 9)
  const enabledCampaignsOnlyCell = settingsSheet.getRange(9, 2);
  enabledCampaignsOnlyCell.insertCheckboxes();
  enabledCampaignsOnlyCell.setValue(settings.enabledCampaignsOnly !== undefined ? settings.enabledCampaignsOnly : false);

  // Row 10: Enable ChatGPT (index 9 in structure array + 1 for header = row 10)
  const enableChatGptCell = settingsSheet.getRange(10, 2);
  enableChatGptCell.insertCheckboxes();
  enableChatGptCell.setValue(false); // Default to disabled

  // Row 11: ChatGPT API Key should NOT have a checkbox - it's a text field

  // Row 12: Use Cached ChatGPT Responses (index 11 in structure array + 1 for header = row 12)
  const useCachedChatGptCell = settingsSheet.getRange(12, 2);
  useCachedChatGptCell.insertCheckboxes();
  useCachedChatGptCell.setValue(true); // Default to enabled

  // Add Placement Type Filters section
  // Row 13: Gap
  // Row 14: Placement Type Filters header
  // Row 15: YouTube Video checkbox
  // Row 16: Website checkbox
  // Row 17: Mobile Application checkbox
  const placementTypeHeaderRow = 14;
  settingsSheet.getRange(placementTypeHeaderRow, 1).setValue('Placement Type Filters');
  settingsSheet.getRange(placementTypeHeaderRow, 1).setFontWeight('bold');
  settingsSheet.getRange(placementTypeHeaderRow, 1).setBackground('#e8f0fe');

  // Add placement type checkboxes (all checked by default)
  settingsSheet.getRange(15, 1).setValue('YouTube Video');
  settingsSheet.getRange(15, 2).insertCheckboxes();
  settingsSheet.getRange(15, 2).setValue(true);

  settingsSheet.getRange(16, 1).setValue('Website');
  settingsSheet.getRange(16, 2).insertCheckboxes();
  settingsSheet.getRange(16, 2).setValue(true);

  settingsSheet.getRange(17, 1).setValue('Mobile Application');
  settingsSheet.getRange(17, 2).insertCheckboxes();
  settingsSheet.getRange(17, 2).setValue(true);

  // Add Important Notes section
  // Row 18: Gap
  // Row 19: Important Notes header
  // Row 20+: Notes
  const notesHeaderRow = 19;
  settingsSheet.getRange(notesHeaderRow, 1).setValue('Important Notes');
  settingsSheet.getRange(notesHeaderRow, 1).setFontWeight('bold');
  settingsSheet.getRange(notesHeaderRow, 1).setBackground('#fff3cd');
  settingsSheet.getRange(notesHeaderRow, 1, 1, 3).merge();

  const notes = [
    '• Only selected (checeked) placements will be excluded',
    '• Mobile application placements will be excluded using Bulk Upload service',
    '• Bulk uploads can be checked in Google Ads under: Tools & Settings > Bulk actions > Uploads',
    '• Bulk uploads will be PREVIEWED when the script is run in preview mode',
    '• Bulk uploads will be APPLIED when the script is run in execution mode',
    '• Website placements are excluded directly via the script'
  ];

  let currentNoteRow = 20;
  for (let i = 0; i < notes.length; i++) {
    settingsSheet.getRange(currentNoteRow, 1).setValue(notes[i]);
    settingsSheet.getRange(currentNoteRow, 1).setFontStyle('italic');
    settingsSheet.getRange(currentNoteRow, 1, 1, 3).merge();
    currentNoteRow++;
  }

  // Freeze header row
  settingsSheet.setFrozenRows(1);

  console.log(`Settings sheet populated successfully`);
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

  return allPlacementData;
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

  if (settings.campaignNameContains && settings.campaignNameContains.trim()) {
    conditions.push(`campaign.name CONTAINS_IGNORE_CASE '${settings.campaignNameContains.trim()}'`);
  }

  if (settings.campaignNameNotContains && settings.campaignNameNotContains.trim()) {
    conditions.push(`campaign.name DOES_NOT_CONTAIN_IGNORE_CASE '${settings.campaignNameNotContains.trim()}'`);
  }

  const whereClause = conditions.join(' AND ');

  // Add LIMIT clause if maxResults is set (0 means no limit)
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
    const meetsImpressionsThreshold = placement.impressions >= settings.minimumImpressions;
    // Note: clicks are not available in performance_max_placement_view, so this filter will always pass
    // if minimumClicks is 0, otherwise it will filter out all placements
    const meetsClicksThreshold = settings.minimumClicks === 0 || placement.clicks >= settings.minimumClicks;

    // Check placement type filter
    let meetsPlacementTypeFilter = true;
    if (settings.placementTypeFilters) {
      const placementType = String(placement.placementType || '').toUpperCase();
      if (placementType.includes('YOUTUBE') || placementType.includes('YOUTUBE_VIDEO')) {
        meetsPlacementTypeFilter = settings.placementTypeFilters.youtubeVideo;
      } else if (placementType.includes('WEBSITE')) {
        meetsPlacementTypeFilter = settings.placementTypeFilters.website;
      } else if (placementType.includes('MOBILE_APPLI')) {
        meetsPlacementTypeFilter = settings.placementTypeFilters.mobileApplication;
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

    // Note: Placement Not Contains filter is NOT used for filtering placements out of the report
    // Placements that match the "Not Contains" list are still included in the report

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

    // Note: Display Name Not Contains filter is NOT used for filtering placements out of the report

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

    // Note: Target URL Not Contains filter is NOT used for filtering placements out of the report

    return meetsImpressionsThreshold && meetsClicksThreshold && meetsPlacementTypeFilter &&
      meetsPlacementContainsFilter && meetsDisplayNameContainsFilter &&
      meetsTargetUrlContainsFilter && meetsTargetUrlEndsWithFilter;
  });
}

// --- Sheet Writing Functions ---

/**
 * Gets all excluded placements from the shared exclusion list
 * @returns {Set<string>} Set of excluded placements (normalized URLs and mobile app IDs)
 */
function getExcludedPlacements() {
  const excludedPlacements = new Set();

  console.log(`\n=== Getting Excluded Placements ===`);
  console.log(`Looking for exclusion list: ${SHARED_EXCLUSION_LIST_NAME}`);

  try {
    const listIterator = AdsApp.excludedPlacementLists()
      .withCondition(`shared_set.name = '${SHARED_EXCLUSION_LIST_NAME}'`)
      .get();

    if (!listIterator.hasNext()) {
      console.log(`⚠ Exclusion list not found`);
      return excludedPlacements; // List doesn't exist, return empty set
    }

    const excludedPlacementList = listIterator.next();
    const listName = excludedPlacementList.getName();
    const listId = excludedPlacementList.getId();
    console.log(`✓ Found exclusion list: ${listName}`);
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

  const dataSheet = getOrCreateDataSheet(spreadsheet);

  // Clear existing data
  dataSheet.clear();

  if (placementData.length === 0) {
    dataSheet.getRange(1, 1, 1, 1).setValue('No placement data found. Check your filters and date range.');
    return;
  }

  // Get currently excluded placements
  const excludedPlacements = getExcludedPlacements();
  if (DEBUG_MODE) {
    console.log(`Found ${excludedPlacements.size} placements already in exclusion list`);
    if (excludedPlacements.size > 0 && excludedPlacements.size <= 10) {
      const excludedArray = Array.from(excludedPlacements);
      console.log(`Excluded placements: ${excludedArray.join(', ')}`);
    }
  }

  // Write headers
  const headers = [
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

  const headerRange = dataSheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  // Write data rows
  const dataRows = placementData.map(placement => {
    // Add note for mobile apps
    let notes = '';
    if (placement.placementType && String(placement.placementType).toUpperCase().includes('MOBILE_APPLI')) {
      notes = 'Will be excluded via bulk upload if selected';
    }

    // Add note for Google domains (cannot be excluded)
    const placementUrl = placement.targetUrl || placement.placement || '';
    if (isGoogleDomain(placementUrl)) {
      if (notes) {
        notes += '; Cannot exclude Google domain';
      } else {
        notes = 'Cannot exclude Google domain';
      }
    }

    // Check if placement is already excluded (only for website placements)
    const status = getPlacementStatus(placement, excludedPlacements);

    return [
      false, // Checkbox column (default unchecked)
      status, // Status column
      placement.campaignName,
      placement.placement,
      placement.displayName,
      placement.placementType,
      placement.targetUrl,
      notes, // Notes column
      placement.chatGptResponse || '', // ChatGPT response
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
  });

  if (dataRows.length > 0) {
    const dataRange = dataSheet.getRange(2, 1, dataRows.length, headers.length);
    dataRange.setValues(dataRows);

    // Add checkboxes to first column
    const checkboxRange = dataSheet.getRange(2, 1, dataRows.length, 1);
    checkboxRange.insertCheckboxes();

    // Format number columns
    formatPlacementSheet(dataSheet, placementData.length);
  }

  // Freeze header row and first column
  dataSheet.setFrozenRows(1);
  dataSheet.setFrozenColumns(1);

  console.log(`✓ Written ${placementData.length} placements to sheet`);
}

/**
 * Checks if a URL is a Google domain that cannot be excluded
 * @param {string} url - The URL to check
 * @returns {boolean} True if the URL is a Google domain
 */
function isGoogleDomain(url) {
  if (!url) {
    return false;
  }
  const lowerUrl = String(url).toLowerCase();
  for (const excludedDomain of GOOGLE_DOMAIN_EXCLUSIONS) {
    if (lowerUrl.includes(excludedDomain)) {
      return true;
    }
  }
  return false;
}

/**
 * Gets or creates the data sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The data sheet
 */
function getOrCreateDataSheet(spreadsheet) {
  let dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
  if (!dataSheet) {
    dataSheet = spreadsheet.insertSheet(DATA_SHEET_NAME);
    console.log(`Created data sheet: ${DATA_SHEET_NAME}`);
  }
  return dataSheet;
}

/**
 * Formats the placement sheet with proper number formatting
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to format
 * @param {number} numRows - Number of data rows
 */
function formatPlacementSheet(sheet, numRows) {
  // Format impressions, clicks, conversions (integers)
  sheet.getRange(2, 10, numRows, 1).setNumberFormat('#,##0'); // Impressions (column J)
  sheet.getRange(2, 11, numRows, 1).setNumberFormat('#,##0'); // Clicks (column K)
  sheet.getRange(2, 13, numRows, 1).setNumberFormat('#,##0'); // Conversions (column M)

  // Format cost, CPA, Avg CPC (currency)
  sheet.getRange(2, 12, numRows, 1).setNumberFormat('#,##0.00'); // Cost (column L)
  sheet.getRange(2, 16, numRows, 1).setNumberFormat('#,##0.00'); // Avg CPC (column P)
  sheet.getRange(2, 18, numRows, 1).setNumberFormat('#,##0.00'); // CPA (column R)

  // Format conversions value, ROAS (currency)
  sheet.getRange(2, 14, numRows, 1).setNumberFormat('#,##0.00'); // Conv. Value (column N)
  sheet.getRange(2, 19, numRows, 1).setNumberFormat('#,##0.00'); // ROAS (column S)

  // Format percentages
  sheet.getRange(2, 15, numRows, 1).setNumberFormat('0.00%'); // CTR (column O)
  sheet.getRange(2, 17, numRows, 1).setNumberFormat('0.00%'); // Conv. Rate (column Q)

  // Format Notes and ChatGPT Response columns (clip text - no wrap)
  sheet.getRange(2, 8, numRows, 1).setWrap(false); // Notes (column H)
  sheet.getRange(2, 9, numRows, 1).setWrap(false); // ChatGPT Response (column I)
}

// --- Exclusion List Functions ---

/**
 * Reads checked placements from the sheet and normalizes them
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @returns {Object} Object with websitePlacements and mobileAppPlacements arrays
 */
function readCheckedPlacementsFromSheet(spreadsheet) {
  const dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
  if (!dataSheet) {
    console.log(`No data sheet found. Nothing to exclude.`);
    return {
      websitePlacements: [],
      mobileAppPlacements: []
    };
  }

  const dataRange = dataSheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length <= 1) {
    console.log(`No data in sheet. Nothing to exclude.`);
    return {
      websitePlacements: [],
      mobileAppPlacements: []
    };
  }

  // Use Set to ensure uniqueness for both types
  const checkedWebsitePlacementsSet = new Set();
  const checkedMobileAppPlacementsSet = new Set();
  let skippedInvalid = 0;
  let skippedGoogleDomains = 0;

  // Start from row 2 (skip header)
  let checkedCount = 0;
  let uncheckedCount = 0;

  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];
    const checkboxValue = row[0];
    // Checkbox values can be true (boolean), "TRUE" (string), or checked state
    const isChecked = checkboxValue === true || String(checkboxValue).toUpperCase() === 'TRUE';
    const placementUrl = row[3]; // Placement is in column D (index 3) - shifted by Status column
    const placementType = row[5]; // Placement Type is in column F (index 5) - shifted by Status column

    if (DEBUG_MODE && rowIndex <= 3) {
      console.log(`Row ${rowIndex + 1}: checkbox=${checkboxValue} (type: ${typeof checkboxValue}), placement=${placementUrl}, type=${placementType}`);
    }

    if (isChecked) {
      checkedCount++;
    } else {
      uncheckedCount++;
    }

    if (isChecked && placementUrl) {
      // First, try to format as mobile app (if applicable)
      const mobileAppPlacement = formatMobileAppPlacement(placementUrl, placementType);

      if (mobileAppPlacement) {
        // It's a mobile app, use the canonical app ID format (e.g., "1-1510189987")
        checkedMobileAppPlacementsSet.add(mobileAppPlacement);
      } else {
        // Skip Google domains - they cannot be excluded
        const url = String(placementUrl).trim();
        if (isGoogleDomain(url)) {
          skippedGoogleDomains++;
          if (DEBUG_MODE) {
            console.log(`Skipping Google domain (cannot exclude): ${url}`);
          }
          continue;
        }

        // It's a website URL, normalize it
        const normalizedUrl = normalizeUrl(placementUrl);

        if (normalizedUrl === null) {
          skippedInvalid++;
        } else {
          checkedWebsitePlacementsSet.add(normalizedUrl);
        }
      }
    }
  }

  const websitePlacements = Array.from(checkedWebsitePlacementsSet);
  const mobileAppPlacements = Array.from(checkedMobileAppPlacementsSet);
  const checkedPlacements = [...websitePlacements, ...mobileAppPlacements];

  if (DEBUG_MODE) {
    console.log(`Checked boxes found: ${checkedCount}, Unchecked: ${uncheckedCount}`);
    console.log(`Valid placements to exclude: ${checkedPlacements.length}`);
  }

  if (checkedPlacements.length > 0) {
    console.log(`Found ${checkedPlacements.length} unique checked placements to exclude`);
    if (websitePlacements.length > 0) {
      console.log(`  - Website placements: ${websitePlacements.length}`);
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
    if (skippedGoogleDomains > 0) {
      console.log(`⚠ Skipped ${skippedGoogleDomains} Google-owned domain(s) (cannot be excluded per policy)`);
    }
    if (skippedInvalid > 0) {
      console.log(`⚠ Skipped ${skippedInvalid} invalid placement URL(s)`);
    }
    console.log(``);
  } else {
    if (skippedGoogleDomains > 0 || skippedInvalid > 0) {
      console.log(`No valid placements to exclude.`);
      if (skippedGoogleDomains > 0) {
        console.log(`⚠ Skipped ${skippedGoogleDomains} Google-owned domain(s) (cannot be excluded per policy)`);
      }
      if (skippedInvalid > 0) {
        console.log(`⚠ Skipped ${skippedInvalid} invalid placement URL(s)`);
      }
    }
  }

  return {
    websitePlacements: websitePlacements,
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
 * Adds website placements to the shared exclusion list using batch processing
 * @param {Array<string>} placementUrls - Array of normalized placement URLs
 */
function addPlacementsToExclusionList(placementUrls) {
  if (placementUrls.length === 0) {
    return;
  }

  const listIterator = AdsApp.excludedPlacementLists()
    .withCondition(`Name = '${SHARED_EXCLUSION_LIST_NAME}'`)
    .get();

  if (!listIterator.hasNext()) {
    throw new Error(`Shared exclusion list '${SHARED_EXCLUSION_LIST_NAME}' not found`);
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
 * Links the shared exclusion list to all enabled Performance Max and Display campaigns
 * This is mandatory for campaigns to honor the negative placements
 */
function linkSharedListToPMaxCampaigns() {
  const listIterator = AdsApp.excludedPlacementLists()
    .withCondition(`Name = '${SHARED_EXCLUSION_LIST_NAME}'`)
    .get();

  if (!listIterator.hasNext()) {
    console.error(`Shared exclusion list '${SHARED_EXCLUSION_LIST_NAME}' not found. Cannot link to campaigns.`);
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
          if (appliedList.getName() === SHARED_EXCLUSION_LIST_NAME) {
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

    // Set up headers
    const headers = [['URL', 'Response', 'Timestamp']];
    llmSheet.getRange(1, 1, 1, 3).setValues(headers);
    llmSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    llmSheet.getRange(1, 1, 1, 3).setBackground('#4285f4');
    llmSheet.getRange(1, 1, 1, 3).setFontColor('#ffffff');
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

  // Start from row 2 (skip header)
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
  const newRow = lastRow + 1;

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
      return cachedResponse;
    }
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
    // Cache the response
    const llmSheet = getOrCreateLlmResponsesSheet(spreadsheet);
    cacheLlmResponse(llmSheet, url, response);
    return response;
  }

  return null;
}


