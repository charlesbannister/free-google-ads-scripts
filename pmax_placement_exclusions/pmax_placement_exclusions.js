/**
 * Performance Max Placement Exclusions - Semi-automated
 * Grabs placement data from Performance Max campaigns, writes to sheet with checkboxes for user selection,
 * then adds selected placements to a shared exclusion list using batch processing
 * Version: 1.2.0
 */

// Google Ads API Query Builder Link:
// Performance Max Placement View: https://developers.google.com/google-ads/api/fields/v20/performance_max_placement_view_query_builder

// --- Configuration ---
const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1OnuQEg-cUHx4cUfcpGsT6zc4TPsNKh5l1IjtCsOaGOw/edit?gid=0#gid=0';
// The Google Sheet URL where placement data will be written
// Example: https://docs.google.com/spreadsheets/d/1abc123def456/edit

const SHARED_EXCLUSION_LIST_NAME = 'PMax Placement Semi-automated Exclusions';
// The name of the shared placement exclusion list that must exist in your Google Ads account
// Create this at: Tools & Settings > Shared Library > Placement exclusions

const SETTINGS_SHEET_NAME = 'Settings';
// The name of the sheet tab that contains configuration settings
// This sheet will be auto-populated on first run if it doesn't exist

const DATA_SHEET_NAME = 'Placements';
// The name of the sheet tab where placement data with checkboxes will be written

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
// These will be automatically filtered out during normalization

// --- Main Function ---
function main() {
  console.log(`Script started`);

  validateConfig();

  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const settingsSheet = getOrCreateSettingsSheet(spreadsheet);

  const settings = getSettingsFromSheet(settingsSheet);
  populateSettingsSheetIfEmpty(settingsSheet, settings);

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
  writePlacementDataToSheet(spreadsheet, transformedData, settings);

  if (checkedPlacements.length > 0) {
    addPlacementsToExclusionList(checkedPlacements);
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
 * Gets settings from the settings sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet - The settings sheet
 * @returns {Object} Settings object with all configuration values
 */
function getSettingsFromSheet(settingsSheet) {
  const lookbackWindowDaysRaw = getSettingValue(settingsSheet, 'Lookback Window (Days)', 30);
  const minimumImpressionsRaw = getSettingValue(settingsSheet, 'Minimum Impressions', 0);
  const minimumClicksRaw = getSettingValue(settingsSheet, 'Minimum Clicks', 0);
  const enabledCampaignsOnlyRaw = getSettingValue(settingsSheet, 'Enabled campaigns only', true);

  // Handle checkbox value - Google Sheets checkboxes return boolean true/false
  // Default to true if not found or invalid
  let enabledCampaignsOnly = true;
  if (typeof enabledCampaignsOnlyRaw === 'boolean') {
    enabledCampaignsOnly = enabledCampaignsOnlyRaw;
  } else if (typeof enabledCampaignsOnlyRaw === 'string') {
    const lowerValue = enabledCampaignsOnlyRaw.toLowerCase().trim();
    enabledCampaignsOnly = lowerValue === 'true' || lowerValue === '1';
  } else if (enabledCampaignsOnlyRaw === null || enabledCampaignsOnlyRaw === undefined || enabledCampaignsOnlyRaw === '') {
    enabledCampaignsOnly = true; // Default to true
  }

  const settings = {
    lookbackWindowDays: typeof lookbackWindowDaysRaw === 'number' ? lookbackWindowDaysRaw : parseInt(lookbackWindowDaysRaw, 10) || 30,
    minimumImpressions: typeof minimumImpressionsRaw === 'number' ? minimumImpressionsRaw : parseInt(minimumImpressionsRaw, 10) || 0,
    minimumClicks: typeof minimumClicksRaw === 'number' ? minimumClicksRaw : parseInt(minimumClicksRaw, 10) || 0,
    campaignNameContains: String(getSettingValue(settingsSheet, 'Campaign Name Contains', '') || '').trim(),
    campaignNameNotContains: String(getSettingValue(settingsSheet, 'Campaign Name Not Contains', '') || '').trim(),
    enabledCampaignsOnly: enabledCampaignsOnly,
    sheet: settingsSheet
  };

  if (DEBUG_MODE) {
    console.log(`Settings loaded:`);
    console.log(`  Lookback Window: ${settings.lookbackWindowDays} days`);
    console.log(`  Minimum Impressions: ${settings.minimumImpressions}`);
    console.log(`  Minimum Clicks: ${settings.minimumClicks}`);
    console.log(`  Campaign Name Contains: "${settings.campaignNameContains}"`);
    console.log(`  Campaign Name Not Contains: "${settings.campaignNameNotContains}"`);
    console.log(`  Enabled campaigns only: ${settings.enabledCampaignsOnly}`);
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
  const settingsStructure = [
    ['Setting', 'Value', 'Description'],
    ['Shared Exclusion List Name', SHARED_EXCLUSION_LIST_NAME, 'The name of the shared placement exclusion list that must exist in your Google Ads account'],
    ['Lookback Window (Days)', settings.lookbackWindowDays, 'Number of days to look back from today (e.g., 30)'],
    ['Minimum Impressions', settings.minimumImpressions, 'Only show placements with at least this many impressions'],
    ['Minimum Clicks', settings.minimumClicks, 'Only show placements with at least this many clicks'],
    ['Campaign Name Contains', settings.campaignNameContains, 'Filter campaigns by name containing this text (leave empty for all)'],
    ['Campaign Name Not Contains', settings.campaignNameNotContains, 'Exclude campaigns with names containing this text (leave empty for none)'],
    ['Enabled campaigns only', '', 'If checked, only include enabled campaigns. If unchecked, include paused campaigns (removed campaigns are always excluded)']
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

  // Add checkbox to "Enabled campaigns only" setting (row 8, column B - after adding Shared Exclusion List Name)
  const enabledCampaignsOnlyCell = settingsSheet.getRange(8, 2);
  enabledCampaignsOnlyCell.insertCheckboxes();
  enabledCampaignsOnlyCell.setValue(settings.enabledCampaignsOnly !== undefined ? settings.enabledCampaignsOnly : true);

  // Freeze header row
  settingsSheet.setFrozenRows(1);

  console.log(`Settings sheet populated successfully`);
}

// --- Data Collection Functions ---

/**
 * Gets placement data from Performance Max campaigns
 * @param {Object} settings - Settings object with filters and date range
 * @returns {Array<Object>} Array of flat placement objects
 */
function getPlacementData(settings) {
  const dateRange = getDateRange(settings.lookbackWindowDays);
  const gaqlQuery = getPlacementGaqlQuery(dateRange, settings);

  console.log(`GAQL Query:`);
  console.log(gaqlQuery);
  console.log(``);

  const report = executePlacementReport(gaqlQuery);
  const placementData = extractPlacementDataFromReport(report);

  if (placementData.length > 0) {
    console.warn(`⚠ Note: performance_max_placement_view only supports impressions metric.`);
    console.warn(`   Clicks, cost, conversions, and other metrics are not available and will show as 0.`);
    console.warn(`   Consider using this view for placement identification only.`);
  }

  return placementData;
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
 * @returns {string} GAQL query string
 */
function getPlacementGaqlQuery(dateRange, settings) {
  const conditions = [
    `segments.date BETWEEN '${dateRange.startDate}' AND '${dateRange.endDate}'`,
    `campaign.advertising_channel_type = 'PERFORMANCE_MAX'`
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
    ORDER BY metrics.impressions DESC
  `;

  return query;
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
    console.error(`https://developers.google.com/google-ads/api/fields/v20/performance_max_placement_view_query_builder`);
    throw error;
  }
}

/**
 * Extracts placement data from the report and returns flat objects
 * @param {GoogleAppsScript.AdsApp.Report} report - The report object
 * @returns {Array<Object>} Array of flat placement objects
 */
function extractPlacementDataFromReport(report) {
  const placementData = [];
  const rows = report.rows();
  let rowCount = 0;

  while (rows.hasNext()) {
    const row = rows.next();
    rowCount++;

    const placement = {
      campaignId: row['campaign.id'],
      campaignName: row['campaign.name'],
      displayName: row['performance_max_placement_view.display_name'],
      placement: row['performance_max_placement_view.placement'],
      placementType: row['performance_max_placement_view.placement_type'],
      targetUrl: row['performance_max_placement_view.target_url'],
      impressions: parseInt(row['metrics.impressions']) || 0,
      clicks: 0, // Not available in performance_max_placement_view
      costMicros: 0, // Not available in performance_max_placement_view
      conversions: 0, // Not available in performance_max_placement_view
      conversionsValue: 0 // Not available in performance_max_placement_view
    };

    placementData.push(placement);

    if (DEBUG_MODE && rowCount <= 3) {
      console.log(`Placement ${rowCount}:`);
      console.log(`  Campaign: ${placement.campaignName}`);
      console.log(`  Placement: ${placement.placement}`);
      console.log(`  Impressions: ${placement.impressions}`);
      console.log(`  Clicks: ${placement.clicks}`);
      console.log(``);
    }
  }

  if (DEBUG_MODE && rowCount > 3) {
    console.log(`... and ${rowCount - 3} more placements`);
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
    return meetsImpressionsThreshold && meetsClicksThreshold;
  });
}

// --- Sheet Writing Functions ---

/**
 * Writes placement data to the sheet with checkboxes
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @param {Array<Object>} placementData - Transformed placement data
 * @param {Object} settings - Settings object with filters
 */
function writePlacementDataToSheet(spreadsheet, placementData, settings) {
  const filteredData = filterPlacementData(placementData, settings);

  console.log(`Writing ${filteredData.length} placements to sheet (after applying filters)`);

  const dataSheet = getOrCreateDataSheet(spreadsheet);

  // Clear existing data
  dataSheet.clear();

  if (filteredData.length === 0) {
    dataSheet.getRange(1, 1, 1, 1).setValue('No placement data found. Check your filters and date range.');
    return;
  }

  // Write headers
  const headers = [
    'Exclude',
    'Campaign Name',
    'Placement',
    'Display Name',
    'Placement Type',
    'Target URL',
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
  const dataRows = filteredData.map(placement => [
    false, // Checkbox column (default unchecked)
    placement.campaignName,
    placement.placement,
    placement.displayName,
    placement.placementType,
    placement.targetUrl,
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
  ]);

  if (dataRows.length > 0) {
    const dataRange = dataSheet.getRange(2, 1, dataRows.length, headers.length);
    dataRange.setValues(dataRows);

    // Add checkboxes to first column
    const checkboxRange = dataSheet.getRange(2, 1, dataRows.length, 1);
    checkboxRange.insertCheckboxes();

    // Format number columns
    formatPlacementSheet(dataSheet, filteredData.length);
  }

  // Freeze header row and first column
  dataSheet.setFrozenRows(1);
  dataSheet.setFrozenColumns(1);

  console.log(`✓ Written ${filteredData.length} placements to sheet`);
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
  sheet.getRange(2, 7, numRows, 1).setNumberFormat('#,##0'); // Impressions
  sheet.getRange(2, 8, numRows, 1).setNumberFormat('#,##0'); // Clicks
  sheet.getRange(2, 10, numRows, 1).setNumberFormat('#,##0'); // Conversions

  // Format cost, CPA, Avg CPC (currency)
  sheet.getRange(2, 9, numRows, 1).setNumberFormat('#,##0.00'); // Cost
  sheet.getRange(2, 13, numRows, 1).setNumberFormat('#,##0.00'); // Avg CPC
  sheet.getRange(2, 15, numRows, 1).setNumberFormat('#,##0.00'); // CPA

  // Format conversions value, ROAS (currency)
  sheet.getRange(2, 11, numRows, 1).setNumberFormat('#,##0.00'); // Conv. Value
  sheet.getRange(2, 16, numRows, 1).setNumberFormat('#,##0.00'); // ROAS

  // Format percentages
  sheet.getRange(2, 12, numRows, 1).setNumberFormat('0.00%'); // CTR
  sheet.getRange(2, 14, numRows, 1).setNumberFormat('0.00%'); // Conv. Rate
}

// --- Exclusion List Functions ---

/**
 * Reads checked placements from the sheet and normalizes them
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @returns {Array<string>} Array of unique, normalized placement URLs to exclude
 */
function readCheckedPlacementsFromSheet(spreadsheet) {
  const dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
  if (!dataSheet) {
    console.log(`No data sheet found. Nothing to exclude.`);
    return [];
  }

  const dataRange = dataSheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length <= 1) {
    console.log(`No data in sheet. Nothing to exclude.`);
    return [];
  }

  // Use Set to ensure uniqueness
  const checkedPlacementsSet = new Set();
  let skippedInvalid = 0;
  let skippedGoogleDomains = 0;
  let skippedMobileApps = 0;

  // Start from row 2 (skip header)
  let checkedCount = 0;
  let uncheckedCount = 0;

  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];
    const checkboxValue = row[0];
    // Checkbox values can be true (boolean), "TRUE" (string), or checked state
    const isChecked = checkboxValue === true || String(checkboxValue).toUpperCase() === 'TRUE';
    const placementUrl = row[2]; // Placement is in column C (index 2)
    const placementType = row[4]; // Placement Type is in column E (index 4)

    if (DEBUG_MODE && rowIndex <= 3) {
      console.log(`Row ${rowIndex + 1}: checkbox=${checkboxValue} (type: ${typeof checkboxValue}), placement=${placementUrl}, type=${placementType}`);
    }

    if (isChecked) {
      checkedCount++;
    } else {
      uncheckedCount++;
    }

    if (isChecked && placementUrl) {
      // Skip mobile app placements - they can't be added to exclusion lists
      if (placementType && String(placementType).toUpperCase().includes('MOBILE_APPLI')) {
        skippedMobileApps++;
        if (DEBUG_MODE) {
          console.log(`Skipping mobile app placement: ${placementUrl} (cannot exclude mobile apps via placement exclusion lists)`);
        }
        continue;
      }

      const normalizedUrl = normalizeUrl(placementUrl);

      if (normalizedUrl === null) {
        // Check if it was a Google domain or invalid
        const rawUrl = String(placementUrl).toLowerCase();
        let isGoogleDomain = false;
        for (const excludedDomain of GOOGLE_DOMAIN_EXCLUSIONS) {
          if (rawUrl.includes(excludedDomain)) {
            isGoogleDomain = true;
            skippedGoogleDomains++;
            break;
          }
        }
        if (!isGoogleDomain) {
          skippedInvalid++;
        }
      } else {
        checkedPlacementsSet.add(normalizedUrl);
      }
    }
  }

  const checkedPlacements = Array.from(checkedPlacementsSet);

  if (DEBUG_MODE) {
    console.log(`Checked boxes found: ${checkedCount}, Unchecked: ${uncheckedCount}`);
    console.log(`Valid placements to exclude: ${checkedPlacements.length}`);
  }

  if (checkedPlacements.length > 0) {
    console.log(`Found ${checkedPlacements.length} unique checked placements to exclude`);
    if (DEBUG_MODE) {
      const firstThree = checkedPlacements.slice(0, 3);
      for (const placement of firstThree) {
        console.log(`  - ${placement}`);
      }
      if (checkedPlacements.length > 3) {
        console.log(`  ... and ${checkedPlacements.length - 3} more`);
      }
    }
    if (skippedMobileApps > 0) {
      console.log(`⚠ Skipped ${skippedMobileApps} mobile app placements (cannot be excluded via placement exclusion lists)`);
    }
    if (skippedGoogleDomains > 0) {
      console.log(`⚠ Skipped ${skippedGoogleDomains} Google-owned domains (cannot be excluded per policy)`);
    }
    if (skippedInvalid > 0) {
      console.log(`⚠ Skipped ${skippedInvalid} invalid placement URLs`);
    }
    console.log(``);
  } else {
    if (skippedMobileApps > 0 || skippedGoogleDomains > 0 || skippedInvalid > 0) {
      console.log(`No valid placements to exclude. Skipped ${skippedMobileApps} mobile apps, ${skippedGoogleDomains} Google domains, and ${skippedInvalid} invalid URLs.`);
    }
  }

  return checkedPlacements;
}

/**
 * Adds placements to the shared exclusion list using batch processing
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
 * Links the shared exclusion list to all enabled Performance Max campaigns
 * This is mandatory for PMax campaigns to honor the negative placements
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

  console.log(`\n=== Campaign Linking Summary ===`);
  console.log(`Newly linked: ${campaignsUpdated} campaigns`);
  console.log(`Already linked: ${campaignsAlreadyLinked} campaigns`);
  if (campaignsFailed > 0) {
    console.log(`Failed to link: ${campaignsFailed} campaigns`);
  }
  console.log(`=================================\n`);
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

  // Check against policy-violating Google domains
  for (const excludedDomain of GOOGLE_DOMAIN_EXCLUSIONS) {
    if (formattedUrl.includes(excludedDomain)) {
      if (DEBUG_MODE) {
        console.log(`Skipping Google-owned domain: ${formattedUrl}`);
      }
      return null; // Discard policy-violating placement
    }
  }

  return formattedUrl;
}
