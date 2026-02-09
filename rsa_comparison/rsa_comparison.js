/**
 * RSA Headlines & Descriptions Comparison Script
 * Compares headlines and descriptions of responsive search ads across campaigns
 * Version 1.0.0
 * @author: Charles Bannister of shabba.io
 */

// Google Ads API Query Builder: https://developers.google.com/google-ads/api/fields/v19/ad_group_ad_query_builder

// ===== CONFIGURATION =====
const DEBUG_MODE = true;
// Set to true to see detailed logs for debugging

const SPREADSHEET_URL = 'YOUR_SPREADSHEET_URL_HERE';
// The Google Sheets URL where results will be written


// Sheet names
const SETTINGS_SHEET_NAME = 'Settings';
const OUTPUT_SHEET_NAME = 'RSA Headlines & Descriptions';

/**
 * Main function - orchestrates the entire script execution
 */
function main() {

  console.log('RSA Comparison Script started');

  if (SPREADSHEET_URL === 'PASTE_YOUR_SPREADSHEET_URL_HERE') {
    throw new Error('Please update SPREADSHEET_URL with your actual Google Sheets URL');
  }

  // Step 1: Set up settings sheet
  const settings = setupSettingsSheet();

  // Step 2: Get RSA data based on settings
  const rsaData = getRsaData(settings);

  // Step 3: Transform and organize the data
  const organizedData = organizeRsaData(rsaData, settings);

  // Step 4: Write results to output sheet
  writeResultsToSheet(organizedData);

  console.log('RSA Comparison Script completed');
}

/**
 * Sets up the settings sheet with default values if it doesn't exist
 * @returns {Object} Settings object with configuration values
 */
function setupSettingsSheet() {
  debugLog('Setting up settings sheet...');

  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  let settingsSheet;

  settingsSheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);

  if (settingsSheet === null) {
    debugLog('Settings sheet does not exist, creating new one');
    settingsSheet = spreadsheet.insertSheet(SETTINGS_SHEET_NAME);
    createDefaultSettings(settingsSheet);
  } else {
    debugLog('Settings sheet already exists, reading current settings');
  }

  return readSettingsFromSheet(settingsSheet);
}

/**
 * Creates default settings in the settings sheet
 * @param {Sheet} settingsSheet - The settings sheet object
 */
function createDefaultSettings(settingsSheet) {
  debugLog('Creating default settings...');

  const defaultSettings = [
    ['Setting', 'Value', 'Description'],
    ['Lookback Window (Days)', 30, 'Number of days to look back for performance data'],
    ['Campaign Name Contains', '', 'Only analyze campaigns containing this text (leave empty for all)'],
    ['Campaign Name Does Not Contain', '', 'Exclude campaigns containing this text (leave empty to exclude none)'],
    ['Minimum Clicks Threshold', 1, 'Only include headlines/descriptions with at least this many clicks']
  ];

  settingsSheet.getRange(1, 1, defaultSettings.length, defaultSettings[0].length).setValues(defaultSettings);
  settingsSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  settingsSheet.setFrozenRows(1);
  settingsSheet.autoResizeColumns(1, 3);

  console.log('Settings sheet created with default values. Please review and update as needed.');
}

/**
 * Reads settings from the settings sheet
 * @param {Sheet} settingsSheet - The settings sheet object
 * @returns {Object} Settings configuration object
 */
function readSettingsFromSheet(settingsSheet) {
  debugLog('Reading settings from sheet...');

  const data = settingsSheet.getDataRange().getValues();
  const settings = {};

  for (let i = 1; i < data.length; i++) {
    const setting = data[i][0];
    const value = data[i][1];

    if (setting && value !== '') {
      switch (setting) {
        case 'Lookback Window (Days)':
          settings.lookbackDays = parseInt(value);
          break;
        case 'Campaign Name Contains':
          settings.campaignNameContains = value.toString();
          break;
        case 'Campaign Name Does Not Contain':
          settings.campaignNameDoesNotContain = value.toString();
          break;
        case 'Minimum Clicks Threshold':
          settings.minimumClicks = parseInt(value);
          break;
      }
    }
  }

  // Set defaults for missing settings
  settings.lookbackDays = settings.lookbackDays || 30;
  settings.campaignNameContains = settings.campaignNameContains || '';
  settings.campaignNameDoesNotContain = settings.campaignNameDoesNotContain || '';
  settings.minimumClicks = settings.minimumClicks || 1;

  debugLog(`Settings loaded: ${JSON.stringify(settings)}`);
  return settings;
}

/**
 * Debug logging function that only logs when DEBUG_MODE is true
 * @param {string} message - The message to log
 */
function debugLog(message) {
  if (DEBUG_MODE) {
    console.log(`[DEBUG] ${message}`);
  }
}

/**
 * Formats a date for Google Ads API queries
 * @param {number} daysAgo - Number of days ago (0 = today)
 * @returns {string} Formatted date string (YYYY-MM-DD)
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
 * Gets RSA headlines and descriptions data with performance metrics
 * @param {Object} settings - Configuration settings
 * @returns {Array} Array of RSA asset data objects
 */
function getRsaData(settings) {
  debugLog('Getting RSA data...');

  const query = getRsaGaqlQuery(settings);
  console.log(`Executing GAQL query: ${query}`);

  const report = getRsaReport(query);
  if (!report) {
    console.log('No RSA data found matching the criteria');
    return [];
  }

  const transformedData = transformRsaReportData(report);
  console.log(`Found ${transformedData.length} RSA assets`);

  if (DEBUG_MODE && transformedData.length > 0) {
    console.log('Sample RSA data (first 3 items):');
    for (let index = 0; index < Math.min(3, transformedData.length); index++) {
      const item = transformedData[index];
      console.log(`\nItem ${index + 1}:`);
      console.log(`  Campaign: ${item.campaignName}`);
      console.log(`  Asset Type: ${item.assetType}`);
      console.log(`  Text: ${item.assetText.substring(0, 50)}${item.assetText.length > 50 ? '...' : ''}`);
      console.log(`  Clicks: ${item.clicks}`);
      console.log(`  Impressions: ${item.impressions}`);
      console.log(`  Conversions: ${item.conversions}`);
    }
  }

  return transformedData;
}

/**
 * Builds the GAQL query for RSA data
 * @param {Object} settings - Configuration settings
 * @returns {string} GAQL query string
 */
function getRsaGaqlQuery(settings) {
  debugLog('Building RSA GAQL query...');

  const startDate = getGoogleAdsApiFormattedDate(settings.lookbackDays);
  const endDate = getGoogleAdsApiFormattedDate(0);

  let whereClause = `WHERE ad_group_ad.ad.type = 'RESPONSIVE_SEARCH_AD'`;
  whereClause += ` AND segments.date BETWEEN '${startDate}' AND '${endDate}'`;

  // Add campaign name filters
  if (settings.campaignNameContains) {
    whereClause += ` AND campaign.name CONTAINS_IGNORE_CASE('${settings.campaignNameContains}')`;
  }

  if (settings.campaignNameDoesNotContain) {
    whereClause += ` AND campaign.name DOES_NOT_CONTAIN_IGNORE_CASE('${settings.campaignNameDoesNotContain}')`;
  }

  // Only look at enabled RSAs in enabled ad groups and campaigns
  whereClause += ` AND ad_group_ad.status = 'ENABLED'`;
  whereClause += ` AND ad_group.status = 'ENABLED'`;
  whereClause += ` AND campaign.status = 'ENABLED'`;

  const query = `
        SELECT 
            campaign.name,
            campaign.id,
            ad_group.name,
            ad_group.id,
            ad_group_ad.ad.id,
            ad_group_ad.ad.responsive_search_ad.headlines,
            ad_group_ad.ad.responsive_search_ad.descriptions,
            metrics.clicks,
            metrics.impressions,
            metrics.conversions,
            metrics.conversions_value
        FROM ad_group_ad 
        ${whereClause}
        ORDER BY metrics.clicks DESC
    `;

  debugLog(`Generated GAQL query: ${query}`);
  return query;
}

/**
 * Executes the GAQL query and returns the report
 * @param {string} query - GAQL query string
 * @returns {Object|null} Report object or null if error
 */
function getRsaReport(query) {
  try {
    const rows = AdsApp.report(query).rows();
    debugLog(`Number of rows from query: ${rows.totalNumEntities()}`);

    if (!rows.hasNext()) {
      console.warn('No RSA data found matching the criteria');
      return null;
    }

    return rows;
  } catch (error) {
    console.error(`Error executing GAQL query: ${error.message}`);
    console.error('Please validate your query at: https://developers.google.com/google-ads/api/fields/v19/query_validator');
    console.error('You can also use the query builder at: https://developers.google.com/google-ads/api/fields/v19/ad_group_ad_query_builder');
    throw error;
  }
}

/**
 * Transforms the raw report data into usable objects
 * @param {Object} report - Report rows from GAQL query
 * @returns {Array} Array of transformed RSA asset objects
 */
function transformRsaReportData(report) {
  debugLog('Transforming RSA report data...');

  const transformedData = [];

  while (report.hasNext()) {
    const row = report.next();

    const campaignName = row['campaign.name'];
    const campaignId = row['campaign.id'];
    const adGroupName = row['ad_group.name'];
    const adGroupId = row['ad_group.id'];
    const adId = row['ad_group_ad.ad.id'];
    const clicks = parseInt(row['metrics.clicks']) || 0;
    const impressions = parseInt(row['metrics.impressions']) || 0;
    const conversions = parseFloat(row['metrics.conversions']) || 0;
    const conversionsValue = parseFloat(row['metrics.conversions_value']) || 0;

    // Process headlines
    const headlines = row['ad_group_ad.ad.responsive_search_ad.headlines'];
    if (headlines) {
      debugLog(`Headlines data type: ${typeof headlines}, value: ${headlines.toString().substring(0, 100)}`);

      let headlinesList;
      if (typeof headlines === 'string') {
        headlinesList = JSON.parse(headlines);
      } else {
        headlinesList = headlines;
      }

      if (Array.isArray(headlinesList)) {
        headlinesList.forEach(headline => {
          transformedData.push({
            campaignName: campaignName,
            campaignId: campaignId,
            adGroupName: adGroupName,
            adGroupId: adGroupId,
            adId: adId,
            assetType: 'Headline',
            assetText: headline.text,
            clicks: clicks,
            impressions: impressions,
            conversions: conversions,
            conversionsValue: conversionsValue
          });
        });
      }
    }

    // Process descriptions
    const descriptions = row['ad_group_ad.ad.responsive_search_ad.descriptions'];
    if (descriptions) {
      let descriptionsList;
      if (typeof descriptions === 'string') {
        descriptionsList = JSON.parse(descriptions);
      } else {
        descriptionsList = descriptions;
      }

      if (Array.isArray(descriptionsList)) {
        descriptionsList.forEach(description => {
          transformedData.push({
            campaignName: campaignName,
            campaignId: campaignId,
            adGroupName: adGroupName,
            adGroupId: adGroupId,
            adId: adId,
            assetType: 'Description',
            assetText: description.text,
            clicks: clicks,
            impressions: impressions,
            conversions: conversions,
            conversionsValue: conversionsValue
          });
        });
      }
    }
  }

  return transformedData;
}

/**
 * Organizes RSA data by asset text with campaigns as columns at asset level
 * @param {Array} rsaData - Array of RSA asset objects
 * @param {Object} settings - Configuration settings
 * @returns {Object} Organized data structure for sheet output
 */
function organizeRsaData(rsaData, settings) {
  debugLog('Organizing RSA data for output...');

  // Group by asset text, type, and campaign to aggregate metrics at asset level
  const groupedData = {};
  const campaignNames = new Set();
  const assetCounts = {}; // Track total count of each asset across all ads

  rsaData.forEach(item => {
    const assetKey = `${item.assetType}|${item.assetText}`;
    const campaignKey = `${assetKey}|${item.campaignName}`;

    campaignNames.add(item.campaignName);

    // Count occurrences of each asset
    if (!assetCounts[assetKey]) {
      assetCounts[assetKey] = 0;
    }
    assetCounts[assetKey]++;

    if (!groupedData[campaignKey]) {
      groupedData[campaignKey] = {
        assetType: item.assetType,
        assetText: item.assetText,
        campaignName: item.campaignName,
        clicks: 0,
        impressions: 0,
        conversions: 0,
        conversionsValue: 0
      };
    }

    // Aggregate metrics for this asset in this campaign
    groupedData[campaignKey].clicks += item.clicks;
    groupedData[campaignKey].impressions += item.impressions;
    groupedData[campaignKey].conversions += item.conversions;
    groupedData[campaignKey].conversionsValue += item.conversionsValue;
  });

  // Calculate derived metrics after aggregation
  const dataWithMetrics = Object.values(groupedData).map(item => {
    const clickThroughRate = item.impressions > 0 ? (item.clicks / item.impressions) : 0;
    const conversionRate = item.clicks > 0 ? (item.conversions / item.clicks) : 0;

    return {
      ...item,
      clickThroughRate: clickThroughRate,
      conversionRate: conversionRate
    };
  });

  // Now organize by asset text across campaigns
  const finalGroupedData = {};

  dataWithMetrics.forEach(item => {
    const key = `${item.assetType}|${item.assetText}`;

    if (!finalGroupedData[key]) {
      finalGroupedData[key] = {
        assetType: item.assetType,
        assetText: item.assetText,
        totalClicks: 0,
        totalImpressions: 0,
        totalConversions: 0,
        totalCount: 0,
        campaigns: {}
      };
    }

    finalGroupedData[key].campaigns[item.campaignName] = {
      clicks: item.clicks,
      impressions: item.impressions,
      clickThroughRate: item.clickThroughRate,
      conversions: item.conversions,
      conversionRate: item.conversionRate
    };

    finalGroupedData[key].totalClicks += item.clicks;
    finalGroupedData[key].totalImpressions += item.impressions;
    finalGroupedData[key].totalConversions += item.conversions;
    finalGroupedData[key].totalCount = assetCounts[key];
  });

  // Calculate total metrics and convert to array
  const organizedArray = Object.values(finalGroupedData).map(item => {
    const totalClickThroughRate = item.totalImpressions > 0 ? (item.totalClicks / item.totalImpressions) : 0;
    const totalConversionRate = item.totalClicks > 0 ? (item.totalConversions / item.totalClicks) : 0;

    return {
      ...item,
      totalClickThroughRate: totalClickThroughRate,
      totalConversionRate: totalConversionRate
    };
  })
    .filter(item => item.totalClicks >= (settings.minimumClicks || 1))
    .sort((a, b) => b.totalClicks - a.totalClicks);

  const sortedCampaignNames = Array.from(campaignNames).sort();

  debugLog(`Organized ${organizedArray.length} unique assets across ${sortedCampaignNames.length} campaigns`);

  if (DEBUG_MODE && organizedArray.length > 0) {
    console.log('Sample organized data (first 3 items):');
    for (let index = 0; index < Math.min(3, organizedArray.length); index++) {
      const item = organizedArray[index];
      console.log(`\nItem ${index + 1}:`);
      console.log(`  Asset Type: ${item.assetType}`);
      console.log(`  Asset Text: ${item.assetText.substring(0, 50)}${item.assetText.length > 50 ? '...' : ''}`);
      console.log(`  Total Clicks: ${item.totalClicks}`);
      console.log(`  Campaigns: ${Object.keys(item.campaigns).join(', ')}`);
    }
  }

  return {
    assets: organizedArray,
    campaignNames: sortedCampaignNames
  };
}

/**
 * Creates or gets the output sheet and writes the comparison data
 * @param {Object} organizedData - Organized data structure with assets and campaign names
 */
function writeResultsToSheet(organizedData) {
  debugLog('Writing results to sheet...');

  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const outputSheet = getOrCreateSheet(spreadsheet, OUTPUT_SHEET_NAME);

  // Clear existing data
  outputSheet.clear();

  if (organizedData.assets.length === 0) {
    outputSheet.getRange(1, 1).setValue('No RSA data found matching the criteria');
    console.log('No data to write to sheet');
    return;
  }

  // Build headers
  const headers = buildSheetHeaders(organizedData.campaignNames);
  const data = buildSheetData(organizedData.assets, organizedData.campaignNames);

  // Write headers
  outputSheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);

  // Write data
  if (data.length > 0) {
    outputSheet.getRange(headers.length + 1, 1, data.length, data[0].length).setValues(data);
  }

  // Format the sheet
  formatOutputSheet(outputSheet, headers.length, organizedData.campaignNames.length);

  console.log(`Results written to sheet: ${outputSheet.getName()}`);
  console.log(`${organizedData.assets.length} unique assets across ${organizedData.campaignNames.length} campaigns`);
}

/**
 * Creates or gets an existing sheet
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {string} sheetName - Name of the sheet to create or get
 * @returns {Sheet} The sheet object
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet === null) {
    debugLog(`Sheet ${sheetName} does not exist, creating new one`);
    sheet = spreadsheet.insertSheet(sheetName);
  }

  return sheet;
}

/**
 * Builds the header rows for the output sheet
 * @param {Array} campaignNames - Array of campaign names
 * @returns {Array} 2D array of header values
 */
function buildSheetHeaders(campaignNames) {
  debugLog('Building sheet headers...');

  // First row: Asset info columns + Count + Account Total + Campaign names
  const firstRow = ['Asset Type', 'Asset Text', 'Count', 'Account Total', '', '', '', ''];
  campaignNames.forEach(campaignName => {
    firstRow.push(campaignName, '', '', '', ''); // 5 columns per campaign
  });

  // Second row: Metric names
  const secondRow = ['', '', '', 'Clicks', 'Impressions', 'CTR %', 'Conversions', 'Conv Rate %']; // Asset info + Account Total metrics
  campaignNames.forEach(() => {
    secondRow.push('Clicks', 'Impressions', 'CTR %', 'Conversions', 'Conv Rate %');
  });

  return [firstRow, secondRow];
}

/**
 * Builds the data rows for the output sheet
 * @param {Array} assets - Array of asset objects
 * @param {Array} campaignNames - Array of campaign names
 * @returns {Array} 2D array of data values
 */
function buildSheetData(assets, campaignNames) {
  debugLog('Building sheet data...');

  const data = [];

  assets.forEach(asset => {
    const row = [
      asset.assetType,
      asset.assetText,
      asset.totalCount,
      asset.totalClicks,
      asset.totalImpressions,
      asset.totalClickThroughRate,
      asset.totalConversions,
      asset.totalConversionRate
    ];

    campaignNames.forEach(campaignName => {
      const campaignData = asset.campaigns[campaignName];
      if (campaignData) {
        row.push(
          campaignData.clicks,
          campaignData.impressions,
          campaignData.clickThroughRate,
          campaignData.conversions,
          campaignData.conversionRate
        );
      } else {
        row.push('', '', '', '', ''); // Empty cells for campaigns without this asset
      }
    });

    data.push(row);
  });

  return data;
}

/**
 * Formats the output sheet with proper styling and frozen rows/columns
 * @param {Sheet} sheet - The sheet to format
 * @param {number} headerRows - Number of header rows
 * @param {number} campaignCount - Number of campaigns
 */
function formatOutputSheet(sheet, headerRows, campaignCount) {
  debugLog('Formatting output sheet...');

  const totalCols = 8 + (campaignCount * 5); // Asset info (3) + Account Total (5) + campaigns
  const totalRows = sheet.getLastRow();

  // Freeze header rows and first eight columns (asset info + account totals)
  sheet.setFrozenRows(headerRows);
  sheet.setFrozenColumns(8);

  // Format header rows
  const headerRange = sheet.getRange(1, 1, headerRows, totalCols);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f0f0f0');

  // Format Account Total header specifically
  const accountTotalRange = sheet.getRange(1, 4, 1, 5);
  accountTotalRange.merge();
  accountTotalRange.setHorizontalAlignment('center');
  accountTotalRange.setBackground('#ffe6cc'); // Light orange background

  // Add border between Account Total and first campaign
  const accountTotalBorderRange = sheet.getRange(1, 9, totalRows, 1);
  accountTotalBorderRange.setBorder(null, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // Format campaign name cells (merge cells in first row for each campaign)
  for (let index = 0; index < campaignCount; index++) {
    const startCol = 9 + (index * 5); // Updated for account total columns
    const campaignRange = sheet.getRange(1, startCol, 1, 5);
    campaignRange.merge();
    campaignRange.setHorizontalAlignment('center');
    campaignRange.setBackground('#e6f2ff');

    // Add thick vertical borders between campaigns
    if (index > 0) {
      const borderRange = sheet.getRange(1, startCol, totalRows, 1);
      borderRange.setBorder(null, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    }
  }

  // Format Account Total percentage columns (no color coding)
  if (totalRows > headerRows) {
    const accountTotalCtrRange = sheet.getRange(headerRows + 1, 6, totalRows - headerRows, 1);
    const accountTotalConvRateRange = sheet.getRange(headerRows + 1, 8, totalRows - headerRows, 1);

    accountTotalCtrRange.setNumberFormat('0.00%');
    accountTotalConvRateRange.setNumberFormat('0.00%');

    // No color coding for Account Total columns
  }

  // Format percentage columns (CTR and Conversion Rate) for campaigns
  for (let index = 0; index < campaignCount; index++) {
    const ctrCol = 9 + (index * 5) + 2; // CTR column (updated for account total columns)
    const convRateCol = 9 + (index * 5) + 4; // Conversion Rate column (updated for account total columns)

    // Format as percentage
    if (totalRows > headerRows) {
      const ctrRange = sheet.getRange(headerRows + 1, ctrCol, totalRows - headerRows, 1);
      const convRateRange = sheet.getRange(headerRows + 1, convRateCol, totalRows - headerRows, 1);

      ctrRange.setNumberFormat('0.00%');
      convRateRange.setNumberFormat('0.00%');

      // Apply color coding for CTR
      applyCtrColorCoding(sheet, ctrRange);

      // Apply color coding for Conversion Rate
      applyConversionRateColorCoding(sheet, convRateRange);
    }
  }

  // Format number columns with appropriate decimal places
  if (totalRows > headerRows) {
    // Format Count column (no decimals)
    const countRange = sheet.getRange(headerRows + 1, 3, totalRows - headerRows, 1);
    countRange.setNumberFormat('0');

    // Format Account Total columns
    const accountTotalClicksRange = sheet.getRange(headerRows + 1, 4, totalRows - headerRows, 1);
    const accountTotalImpressionsRange = sheet.getRange(headerRows + 1, 5, totalRows - headerRows, 1);
    const accountTotalConversionsRange = sheet.getRange(headerRows + 1, 7, totalRows - headerRows, 1);

    accountTotalClicksRange.setNumberFormat('0');
    accountTotalImpressionsRange.setNumberFormat('0');
    accountTotalConversionsRange.setNumberFormat('0.00');

    // Format each campaign's columns
    for (let index = 0; index < campaignCount; index++) {
      const clicksCol = 9 + (index * 5); // Clicks column (updated for account total)
      const impressionsCol = 9 + (index * 5) + 1; // Impressions column
      const conversionsCol = 9 + (index * 5) + 3; // Conversions column

      // Format Clicks (no decimals)
      const clicksRange = sheet.getRange(headerRows + 1, clicksCol, totalRows - headerRows, 1);
      clicksRange.setNumberFormat('0');

      // Format Impressions (no decimals)  
      const impressionsRange = sheet.getRange(headerRows + 1, impressionsCol, totalRows - headerRows, 1);
      impressionsRange.setNumberFormat('0');

      // Format Conversions (2 decimal places)
      const conversionsRange = sheet.getRange(headerRows + 1, conversionsCol, totalRows - headerRows, 1);
      conversionsRange.setNumberFormat('0.00');
    }

    // Format text columns explicitly as text
    const assetTypeRange = sheet.getRange(headerRows + 1, 1, totalRows - headerRows, 1);
    const assetTextRange = sheet.getRange(headerRows + 1, 2, totalRows - headerRows, 1);
    assetTypeRange.setNumberFormat('@'); // @ format means text
    assetTextRange.setNumberFormat('@');
  }

  // Auto-resize columns
  sheet.autoResizeColumns(1, totalCols);

  // Add horizontal border to separate headers from data
  if (totalRows > headerRows) {
    const headerSeparatorRange = sheet.getRange(headerRows, 1, 1, totalCols);
    headerSeparatorRange.setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  }

  console.log('Sheet formatting completed');
}

/**
 * Applies color coding to CTR columns using pink-green scale
 * @param {Sheet} sheet - The sheet object
 * @param {Range} range - The range to apply color coding to
 */
function applyCtrColorCoding(sheet, range) {
  debugLog('Applying CTR color coding...');

  const values = range.getValues();
  const colors = [];

  // Find min and max values for scaling (excluding zeros)
  const flatValues = values.flat().filter(val => val !== '' && val !== null && !isNaN(val) && val > 0);
  if (flatValues.length === 0) return;

  const minVal = Math.min(...flatValues);
  const maxVal = Math.max(...flatValues);
  const range_val = maxVal - minVal;

  values.forEach(row => {
    const rowColors = [];
    row.forEach(value => {
      if (value === '' || value === null || isNaN(value) || value === 0) {
        rowColors.push('#ffffff'); // White for empty cells and zero values
      } else {
        const normalized = range_val > 0 ? (value - minVal) / range_val : 0.5;
        const color = getColorForValue(normalized);
        rowColors.push(color);
      }
    });
    colors.push(rowColors);
  });

  range.setBackgrounds(colors);
}

/**
 * Applies color coding to conversion rate columns using pink-green scale
 * @param {Sheet} sheet - The sheet object
 * @param {Range} range - The range to apply color coding to
 */
function applyConversionRateColorCoding(sheet, range) {
  debugLog('Applying conversion rate color coding...');

  const values = range.getValues();
  const colors = [];

  // Find min and max values for scaling (excluding zeros)
  const flatValues = values.flat().filter(val => val !== '' && val !== null && !isNaN(val) && val > 0);
  if (flatValues.length === 0) return;

  const minVal = Math.min(...flatValues);
  const maxVal = Math.max(...flatValues);
  const range_val = maxVal - minVal;

  values.forEach(row => {
    const rowColors = [];
    row.forEach(value => {
      if (value === '' || value === null || isNaN(value) || value === 0) {
        rowColors.push('#ffffff'); // White for empty cells and zero values
      } else {
        const normalized = range_val > 0 ? (value - minVal) / range_val : 0.5;
        const color = getColorForValue(normalized);
        rowColors.push(color);
      }
    });
    colors.push(rowColors);
  });

  range.setBackgrounds(colors);
}

/**
 * Gets color for normalized value on pink-green scale
 * @param {number} normalized - Value between 0 and 1
 * @returns {string} Hex color code
 */
function getColorForValue(normalized) {
  // Pink (low) to Green (high) color scale
  const red = Math.round(255 * (1 - normalized) + 144 * normalized);
  const green = Math.round(182 * (1 - normalized) + 238 * normalized);
  const blue = Math.round(193 * (1 - normalized) + 144 * normalized);

  return `#${red.toString(16).padStart(2, '0')}${green.toString(16).padStart(2, '0')}${blue.toString(16).padStart(2, '0')}`;
}
