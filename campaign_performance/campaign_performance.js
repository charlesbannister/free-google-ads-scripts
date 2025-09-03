/**
 * Weekly Campaign Performance Report
 * Generates a comprehensive weekly performance report for all campaigns with color-coded metrics
 * Version: 3.9.4
 * @Authors
 *  - Muhammad Saqib (ideation & design)
 *  - Charles Bannister (code)
 */

// Google Ads API Query Builder Links:
// Campaign Report: https://developers.google.com/google-ads/api/fields/v20/campaign_query_builder
// Campaign Budget Report: https://developers.google.com/google-ads/api/fields/v20/campaign_budget_query_builder
// Accessible Bidding Strategy Report: https://developers.google.com/google-ads/api/fields/v20/accessible_bidding_strategy_query_builder

// Configuration
const SPREADSHEET_URL = 'YOUR_SPREADSHEET_URL_HERE';
// To create a new spreadsheet: Type "sheets.new" in your browser address bar
// Then share the URL here. Example: https://docs.google.com/spreadsheets/d/1abc123def456/edit

const ACCOUNT_IDS = [''];
// Array of account IDs to run the script for when running at MCC/Manager level
// Example: ['123-456-7890', '987-654-3210']
// Leave empty to specify in the script configuration

const SETTINGS_SHEET_NAME = 'settings';
// The sheet that contains the settings and metric checkboxes

const DEBUG_MODE = true;
// Set to true to see detailed logs for debugging

const LIMIT_CAMPAIGNS = false;
// Set to true to limit the number of campaigns processed

const MAX_CAMPAIGNS = 1;
// Maximum number of campaigns to process when LIMIT_CAMPAIGNS is true
// This helps with testing or focusing on specific campaigns

function main() {
  if (!isMCC()) {
    runAccount();
    return;
  }

  if (typeof ACCOUNT_IDS === 'undefined' || ACCOUNT_IDS.length === 0) {
    console.error("To run at MCC/Manager level, specify ACCOUNT_IDS in the script.");
    return;
  }

  MccApp.accounts()
    .withIds(ACCOUNT_IDS)
    .withLimit(50)
    .executeInParallel("runAccount");
}

function runAccount() {
  console.log('Script started');

  validateConfig();

  const settings = getSettings();
  const enabledMetrics = getEnabledMetrics(settings);
  const weeksInPeriod = settings.weeksInPeriod;
  const numberOfPeriods = settings.numberOfPeriods;
  const periodType = settings.periodType;

  console.log(`Enabled metrics: ${enabledMetrics.join(', ')}`);
  const periodLengthLabel = periodType === 'months' ? 'months' : 'weeks';
  console.log(`Period length: ${weeksInPeriod} ${periodLengthLabel}`);
  console.log(`Number of periods: ${numberOfPeriods}`);
  console.log(`Period type: ${periodType}`);

  const campaignData = getCampaignData(enabledMetrics, weeksInPeriod, numberOfPeriods, periodType);

  if (campaignData.length === 0) {
    console.warn('No campaign data found. Exiting script.');
    return;
  }

  const accountName = getAccountName();
  writeDataToSheet(campaignData, enabledMetrics, weeksInPeriod, numberOfPeriods, accountName, settings.sheet, periodType);

  // Send email if email addresses are configured
  if (settings.emailAddresses && settings.emailAddresses.trim()) {
    console.log('Sending email report...');
    sendEmailReport(campaignData, enabledMetrics, weeksInPeriod, numberOfPeriods, accountName, settings.emailAddresses, settings.sheet, periodType);
  }

  console.log('Script finished successfully');
}

function validateConfig() {
  if (SPREADSHEET_URL === 'PASTE_YOUR_SPREADSHEET_URL_HERE') {
    throw new Error('Please update SPREADSHEET_URL with your actual spreadsheet URL');
  }
}

function getSettings() {
  console.log('Getting settings from spreadsheet...');

  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const settingsSheet = getOrCreateSettingsSheet(spreadsheet);

  const periodLength = settingsSheet.getRange('B4').getValue();
  let numberOfPeriods = settingsSheet.getRange('B5').getValue();
  const emailAddresses = settingsSheet.getRange('B6').getValue();
  let periodType = settingsSheet.getRange('B7').getValue();
  console.log(`Raw period type from B7: "${periodType}" (type: ${typeof periodType})`);

  // Normalize period type (handle whitespace and case issues)
  if (periodType && typeof periodType === 'string') {
    periodType = periodType.toString().trim().toLowerCase();
  }
  console.log(`Normalized period type: "${periodType}"`);

  if (!periodLength || periodLength <= 0) {
    throw new Error('Period Length (B4) must be a positive number');
  }

  // Handle case where numberOfPeriods might be empty or invalid (for existing sheets)
  console.log(`Raw value from B5: ${numberOfPeriods} (type: ${typeof numberOfPeriods})`);

  if (!numberOfPeriods || typeof numberOfPeriods !== 'number' || numberOfPeriods <= 0 || isNaN(numberOfPeriods)) {
    console.warn(`Number of Periods (B5) is invalid or missing (value: "${numberOfPeriods}", type: ${typeof numberOfPeriods}). Setting default value of 3 and updating sheet.`);
    numberOfPeriods = 3;
    settingsSheet.getRange('B5').setValue(numberOfPeriods);
  } else {
    // Ensure it's a whole number
    numberOfPeriods = Math.floor(numberOfPeriods);
    if (numberOfPeriods < 1) {
      console.warn(`Number of Periods must be at least 1. Setting to 3.`);
      numberOfPeriods = 3;
      settingsSheet.getRange('B5').setValue(numberOfPeriods);
    }
  }

  const periodTypeDisplay = periodType || 'weeks';
  const periodLengthLabel = periodTypeDisplay === 'months' ? 'months' : 'weeks';
  console.log(`Settings loaded - Period Length: ${periodLength} ${periodLengthLabel}, Number of Periods: ${numberOfPeriods}, Period Type: ${periodTypeDisplay}`);
  if (emailAddresses) {
    console.log(`Email addresses configured: ${emailAddresses}`);
  }

  return {
    weeksInPeriod: periodLength, // Keep the same property name for backward compatibility
    numberOfPeriods: numberOfPeriods,
    emailAddresses: emailAddresses || '',
    periodType: periodType || 'weeks',
    sheet: settingsSheet
  };
}

/**
 * Creates the settings sheet with proper structure if it doesn't exist
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet object
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The settings sheet
 */
function getOrCreateSettingsSheet(spreadsheet) {
  let settingsSheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);

  if (!settingsSheet) {
    console.log(`Creating new settings sheet: ${SETTINGS_SHEET_NAME}`);
    settingsSheet = spreadsheet.insertSheet(SETTINGS_SHEET_NAME);
    setupSettingsSheetStructure(settingsSheet);
  } else {
    console.log(`Using existing settings sheet: ${SETTINGS_SHEET_NAME}`);
    // Check if existing sheet has the new structure and update if needed
    ensureSettingsSheetStructure(settingsSheet);
  }

  return settingsSheet;
}

/**
 * Ensures an existing settings sheet has all the required fields
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The existing settings sheet
 */
function ensureSettingsSheetStructure(sheet) {
  try {
    // Check if B5 (Number of Periods) exists and has a value
    const numberOfPeriodsValue = sheet.getRange('B5').getValue();
    const a5Value = sheet.getRange('A5').getValue();

    // If A5 doesn't contain "Number of Periods" or B5 is empty, update the structure
    if (!a5Value || !a5Value.toString().includes('Number of Periods')) {
      console.log('Updating existing settings sheet with new structure...');

      // Update Period Length setting name if needed
      const a4Value = sheet.getRange('A4').getValue();
      if (!a4Value || a4Value.toString().includes('Weeks in Period')) {
        sheet.getRange('A4').setValue('Period Length');
        sheet.getRange('C4').setValue('Length of each time period. If Period Type = "weeks": number of weeks per period (e.g., 4 = 4 weeks). If Period Type = "months": number of months per period (e.g., 3 = 3 months)');
      }

      // Add the missing Number of Periods setting
      sheet.getRange('A5').setValue('Number of Periods');
      sheet.getRange('B5').setValue(3); // Default to 3 periods
      sheet.getRange('C5').setValue('Total number of time periods to compare (e.g., 3 = compare 3 different time periods)');

      // Check and update Email Addresses setting if needed
      const a6Value = sheet.getRange('A6').getValue();
      if (!a6Value || !a6Value.toString().includes('Email Addresses')) {
        sheet.getRange('A6').setValue('Email Addresses');
        sheet.getRange('B6').setValue(''); // Empty by default
        sheet.getRange('C6').setValue('Comma-separated email addresses to send the report to (leave blank to skip email)');
      }

      // Check and update Period Type setting if needed
      const a7Value = sheet.getRange('A7').getValue();
      if (!a7Value || !a7Value.toString().includes('Period Type')) {
        sheet.getRange('A7').setValue('Period Type');
        sheet.getRange('B7').setValue('weeks'); // Default to weeks
        sheet.getRange('C7').setValue('Type of time period: "weeks" or "months" (e.g., weeks = weekly periods, months = monthly periods)');
      }

      // Update headers if needed
      const c1Value = sheet.getRange('C1').getValue();
      if (!c1Value || !c1Value.toString().includes('Notes')) {
        sheet.getRange('C1').setValue('Notes');

        // Update Notes column formatting
        const notesRange = sheet.getRange('C4:C6');
        notesRange.setFontStyle('italic');
        notesRange.setFontSize(10);
        notesRange.setFontColor('#666666');
        notesRange.setWrap(true);

        // Update column widths
        sheet.setColumnWidth(3, 400); // Column C - Notes
      }

      console.log('Settings sheet structure updated successfully');
    }
  } catch (error) {
    console.warn('Error checking/updating settings sheet structure:', error.message);
  }
}

/**
 * Sets up the initial structure of the settings sheet with labels and checkboxes
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The settings sheet to set up
 */
function setupSettingsSheetStructure(sheet) {
  console.log('Setting up settings sheet structure...');

  // Clear any existing content
  sheet.clear();

  // Set up headers and labels
  sheet.getRange('A1').setValue('Setting');
  sheet.getRange('B1').setValue('Value');
  sheet.getRange('C1').setValue('Notes');

  // Report Configuration section
  sheet.getRange('A3').setValue('Report Configuration:');

  // Period Length setting
  sheet.getRange('A4').setValue('Period Length');
  sheet.getRange('B4').setValue(4); // Default to 4 
  sheet.getRange('C4').setValue('Length of each time period. If Period Type = "weeks": number of weeks per period (e.g., 4 = 4 weeks). If Period Type = "months": number of months per period (e.g., 3 = 3 months)');

  // Number of Periods setting
  sheet.getRange('A5').setValue('Number of Periods');
  sheet.getRange('B5').setValue(3); // Default to 3 periods
  sheet.getRange('C5').setValue('Total number of time periods to compare (e.g., 3 = compare 3 different time periods)');

  // Email setting
  sheet.getRange('A6').setValue('Email Addresses');
  sheet.getRange('B6').setValue(''); // Empty by default - comma separated emails
  sheet.getRange('C6').setValue('Comma-separated email addresses to send the report to (leave blank to skip email)');

  // Period Type setting
  sheet.getRange('A7').setValue('Period Type');
  sheet.getRange('B7').setValue('weeks'); // Default to weeks
  sheet.getRange('C7').setValue('Type of time period: "weeks" or "months" (e.g., weeks = weekly periods, months = monthly periods)');

  // Metric checkboxes section
  sheet.getRange('A9').setValue('Enabled Metrics:');
  sheet.getRange('B9').setValue('Check the metrics you want to include in the report (leave emails blank to skip email)');

  // Metric labels and checkboxes
  const metricLabels = [
    { row: 10, label: 'Impressions', cell: 'B10' },
    { row: 11, label: 'Clicks', cell: 'B11' },
    { row: 12, label: 'Cost', cell: 'B12' },
    { row: 13, label: 'CTR (%)', cell: 'B13' },
    { row: 14, label: 'CPC', cell: 'B14' },
    { row: 15, label: 'Conversions', cell: 'B15' },
    { row: 16, label: 'Conv. Rate (%)', cell: 'B16' },
    { row: 17, label: 'Cost/Conv.', cell: 'B17' },
    { row: 18, label: 'Search Impr. Share (%)', cell: 'B18' },
    { row: 19, label: 'Search Top IS (%)', cell: 'B19' },
    { row: 20, label: 'Search Abs. Top IS (%)', cell: 'B20' }
  ];

  metricLabels.forEach(metric => {
    // Set metric label in column A
    sheet.getRange(`A${metric.row}`).setValue(metric.label);

    // Set checkbox in column B (default all to true)
    sheet.getRange(metric.cell).insertCheckboxes();
    sheet.getRange(metric.cell).setValue(true);
  });

  // Format the sheet
  formatSettingsSheet(sheet);

  console.log('Settings sheet structure created successfully');
}

/**
 * Applies formatting to the settings sheet for better readability
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The settings sheet to format
 */
function formatSettingsSheet(sheet) {
  // Format headers (row 1)
  const headerRange = sheet.getRange('A1:C1');
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e6f3ff');

  // Format section headers
  sheet.getRange('A3').setFontWeight('bold');
  sheet.getRange('A8').setFontWeight('bold');

  // Format description text in column B
  sheet.getRange('B8').setFontStyle('italic');

  // Format notes in column C
  const notesRange = sheet.getRange('C4:C6');
  notesRange.setFontStyle('italic');
  notesRange.setFontSize(10);
  notesRange.setFontColor('#666666');
  notesRange.setWrap(true);

  // Set column widths for better readability
  sheet.setColumnWidth(1, 200); // Column A - Setting names
  sheet.setColumnWidth(2, 200); // Column B - Values/checkboxes
  sheet.setColumnWidth(3, 400); // Column C - Notes

  // Auto-resize rows to fit content
  sheet.autoResizeRows(1, 19);

  // Freeze the header row
  sheet.setFrozenRows(1);
}

function getEnabledMetrics(settings) {
  console.log('Checking which metrics are enabled...');

  const metricMap = {
    'B10': 'impressions',
    'B11': 'clicks',
    'B12': 'cost',
    'B13': 'ctr',
    'B14': 'cpc',
    'B15': 'conversions',
    'B16': 'conversionRate',
    'B17': 'costPerConversion',
    'B18': 'searchImpressionShare',
    'B19': 'searchTopImpressionShare',
    'B20': 'searchAbsoluteTopImpressionShare'
  };

  const enabledMetrics = [];

  for (const [cell, metric] of Object.entries(metricMap)) {
    const isEnabled = settings.sheet.getRange(cell).getValue();
    if (isEnabled === true) {
      enabledMetrics.push(metric);
    }
  }

  if (enabledMetrics.length === 0) {
    throw new Error('At least one metric must be enabled in the settings sheet (B9-B19)');
  }

  return enabledMetrics;
}

function getCampaignData(enabledMetrics, weeksInPeriod, numberOfPeriods, periodType) {
  console.log('Getting campaign data...');

  const dateRanges = calculateDateRanges(weeksInPeriod, numberOfPeriods, periodType);

  // Get data for each period separately
  const periodData = [];
  for (let i = 1; i <= numberOfPeriods; i++) {
    const periodKey = `period${i}`;
    const periodLabel = `Period ${i}`;
    const data = getCampaignDataForPeriod(enabledMetrics, dateRanges[periodKey], periodLabel);
    periodData.push(data);
  }

  const campaignData = combinePeriodData(periodData, dateRanges, numberOfPeriods);

  return campaignData;
}

function calculateDateRanges(weeksInPeriod, numberOfPeriods, periodType) {
  const today = new Date();
  const dateRanges = {};

  console.log(`Date ranges calculated for ${periodType || 'weeks'} periods:`);
  console.log(`Period type check: "${periodType}" === "months" is ${periodType === 'months'}`);

  // Calculate periods starting from the most recent (period with highest number) back to oldest (period 1)
  for (let i = numberOfPeriods; i >= 1; i--) {
    const periodIndex = numberOfPeriods - i; // 0 for most recent, 1 for second most recent, etc.

    let startDate, endDate;
    if (periodType === 'months') {
      console.log(`Using months calculation for period ${i}`);
      // For monthly periods, weeksInPeriod actually represents months per period
      const monthsToSubtract = periodIndex * weeksInPeriod; // weeksInPeriod = months when periodType = 'months'

      if (i === numberOfPeriods) {
        // Most recent period: current month from 1st to today
        const currentMonth = new Date(today);
        currentMonth.setMonth(currentMonth.getMonth() - monthsToSubtract);

        // Start date: first day of the month
        startDate = new Date(currentMonth.getFullYear(), currentMonth.getMonth(), 1);
        // End date: today (or last day of month if we're looking at a past month)
        if (monthsToSubtract === 0) {
          endDate = new Date(today); // Current month: end today
        } else {
          // Past month: end on last day of that month
          endDate = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + weeksInPeriod, 0);
        }
      } else {
        // Historical periods: full months
        const targetMonth = new Date(today);
        targetMonth.setMonth(targetMonth.getMonth() - monthsToSubtract);

        // Start date: first day of the target month
        startDate = new Date(targetMonth.getFullYear(), targetMonth.getMonth(), 1);
        // End date: last day of the period (after weeksInPeriod months)
        endDate = new Date(targetMonth.getFullYear(), targetMonth.getMonth() + weeksInPeriod, 0);
      }
    } else {
      console.log(`Using weeks calculation for period ${i}`);

      // Get the Monday of the current week
      const currentMondayDate = getMondayOfWeek(today);

      if (i === numberOfPeriods) {
        // Most recent period: handle partial week ending today
        if (weeksInPeriod === 1) {
          // Single week: Monday of current week to today
          startDate = new Date(currentMondayDate);
          endDate = new Date(today);
        } else {
          // Multiple weeks: go back the required number of weeks, end today
          startDate = new Date(currentMondayDate);
          startDate.setDate(startDate.getDate() - ((weeksInPeriod - 1) * 7));
          endDate = new Date(today);
        }
      } else {
        // Historical periods: full calendar weeks (Monday to Sunday)
        const weeksBack = (numberOfPeriods - i) * weeksInPeriod;

        // Calculate the Monday that starts this period
        startDate = new Date(currentMondayDate);
        startDate.setDate(startDate.getDate() - (weeksBack * 7));

        // End date is the Sunday of the last week in this period
        endDate = new Date(startDate);
        endDate.setDate(endDate.getDate() + (weeksInPeriod * 7) - 1);
      }
    }

    let label;
    if (numberOfPeriods === 1) {
      label = 'Current period';
    } else if (i === numberOfPeriods) {
      label = 'Most recent period';
    } else if (i === 1) {
      label = 'Oldest period';
    } else {
      label = `Period ${i}`;
    }

    dateRanges[`period${i}`] = {
      startDate: formatDateForGoogleAds(startDate),
      endDate: formatDateForGoogleAds(endDate),
      label: label
    };

    console.log(`Period ${i} (${label}): ${dateRanges[`period${i}`].startDate} to ${dateRanges[`period${i}`].endDate}`);
  }

  console.log('');
  return dateRanges;
}

function getCampaignDataForPeriod(enabledMetrics, dateRange, periodLabel) {
  console.log(`Getting ${periodLabel} data (${dateRange.startDate} to ${dateRange.endDate})...`);

  const query = getCampaignGaqlQueryForPeriod(enabledMetrics, dateRange, periodLabel);
  const report = getCampaignReport(query, periodLabel);
  const campaignData = processCampaignReportForPeriod(report, enabledMetrics, periodLabel, dateRange);

  return campaignData;
}

function getCampaignGaqlQueryForPeriod(enabledMetrics, dateRange, periodLabel) {
  // Build metrics selection based on enabled metrics
  const metricFields = [];
  if (enabledMetrics.includes('impressions')) metricFields.push('metrics.impressions');
  if (enabledMetrics.includes('clicks')) metricFields.push('metrics.clicks');
  if (enabledMetrics.includes('cost')) metricFields.push('metrics.cost_micros');
  if (enabledMetrics.includes('conversions')) metricFields.push('metrics.conversions');
  if (enabledMetrics.includes('conversionRate')) metricFields.push('metrics.conversions_value');
  if (enabledMetrics.includes('searchImpressionShare')) metricFields.push('metrics.search_impression_share');
  if (enabledMetrics.includes('searchTopImpressionShare')) metricFields.push('metrics.search_top_impression_share');
  if (enabledMetrics.includes('searchAbsoluteTopImpressionShare')) metricFields.push('metrics.search_absolute_top_impression_share');

  const query = `
        SELECT 
            campaign.name,
            campaign.bidding_strategy_type,
            accessible_bidding_strategy.name,
            campaign_budget.amount_micros,
            campaign_budget.period,
            ${metricFields.join(', ')}
        FROM campaign 
        WHERE segments.date BETWEEN '${dateRange.startDate}' AND '${dateRange.endDate}'
        AND campaign.status = 'ENABLED'
        ORDER BY campaign.name
    `;

  console.log(`${periodLabel} GAQL Query:`);
  console.log(query);
  console.log('');

  return query;
}

function getCampaignReport(query, periodLabel) {
  try {
    console.log(`Executing ${periodLabel} campaign report query...`);
    const rows = AdsApp.report(query).rows();

    if (!rows.hasNext()) {
      console.warn(`No campaign data found for ${periodLabel}`);
      return null;
    }

    debugLog(`${periodLabel} - Number of rows from query: ${rows.totalNumEntities()}`);
    return rows;

  } catch (error) {
    console.error(`Error executing ${periodLabel} campaign query:`, error.message);
    console.log('Please use Google\'s query builder to validate and fix the query:');
    console.log('Campaign Report: https://developers.google.com/google-ads/api/fields/v20/campaign_query_builder');
    console.log('Campaign Budget Report: https://developers.google.com/google-ads/api/fields/v20/campaign_budget_query_builder');
    console.log('Accessible Bidding Strategy Report: https://developers.google.com/google-ads/api/fields/v20/accessible_bidding_strategy_query_builder');
    throw error;
  }
}

function processCampaignReportForPeriod(report, enabledMetrics, periodLabel, dateRange) {
  if (!report) return new Map();

  console.log(`Processing ${periodLabel} campaign report data...`);

  const campaignMap = new Map();

  while (report.hasNext()) {
    const row = report.next();

    const campaignName = row['campaign.name'];
    const biddingStrategyType = row['campaign.bidding_strategy_type'];
    const biddingStrategyName = row['accessible_bidding_strategy.name'];
    const budgetAmountMicros = row['campaign_budget.amount_micros'];
    const budgetPeriod = row['campaign_budget.period'];

    const periodData = {
      impressions: row['metrics.impressions'] || 0,
      clicks: row['metrics.clicks'] || 0,
      cost: (row['metrics.cost_micros'] || 0) / 1000000, // Convert micros to currency
      conversions: row['metrics.conversions'] || 0,
      conversionsValue: row['metrics.conversions_value'] || 0,
      searchImpressionShare: row['metrics.search_impression_share'] || 0,
      searchTopImpressionShare: row['metrics.search_top_impression_share'] || 0,
      searchAbsoluteTopImpressionShare: row['metrics.search_absolute_top_impression_share'] || 0
    };

    // Calculate derived metrics
    periodData.ctr = periodData.impressions > 0 ? periodData.clicks / periodData.impressions : 0;
    periodData.cpc = periodData.clicks > 0 ? periodData.cost / periodData.clicks : 0;
    periodData.conversionRate = periodData.clicks > 0 ? periodData.conversions / periodData.clicks : 0;
    periodData.costPerConversion = periodData.conversions > 0 ? periodData.cost / periodData.conversions : 0;

    // Add date range information
    periodData.dateRange = `${formatWeekDate(dateRange.startDate)} to ${formatWeekDate(dateRange.endDate)}`;
    periodData.startDate = dateRange.startDate;
    periodData.endDate = dateRange.endDate;

    campaignMap.set(campaignName, {
      name: campaignName,
      biddingStrategyType: biddingStrategyType,
      biddingStrategyName: biddingStrategyName,
      budgetAmount: budgetAmountMicros,
      budgetPeriod: budgetPeriod,
      data: periodData
    });
  }

  console.log(`${periodLabel} - Processed ${campaignMap.size} campaigns`);

  // Log sample data for verification
  if (campaignMap.size > 0) {
    const firstCampaign = campaignMap.values().next().value;
    debugLog(`${periodLabel} sample: ${firstCampaign.name} - Cost: $${firstCampaign.data.cost.toFixed(2)}, Clicks: ${firstCampaign.data.clicks}`);
  }

  return campaignMap;
}

function combinePeriodData(periodDataArray, dateRanges, numberOfPeriods) {
  console.log('Combining period data...');

  // Get all unique campaign names from all periods
  const allCampaignNames = new Set();
  periodDataArray.forEach(periodData => {
    if (periodData) {
      periodData.forEach((campaign, campaignName) => {
        allCampaignNames.add(campaignName);
      });
    }
  });

  const combinedCampaigns = [];

  for (const campaignName of allCampaignNames) {
    const campaigns = [];

    // Get campaign data for each period
    for (let i = 0; i < numberOfPeriods; i++) {
      const periodData = periodDataArray[i];
      campaigns.push(periodData ? periodData.get(campaignName) : null);
    }

    // Use the bidding strategy and budget from the most recent period that has data
    let biddingStrategyType, biddingStrategyName, budgetAmount, budgetPeriod;
    for (let i = numberOfPeriods - 1; i >= 0; i--) {
      const campaign = campaigns[i];
      if (campaign) {
        biddingStrategyType = biddingStrategyType || campaign.biddingStrategyType;
        biddingStrategyName = biddingStrategyName || campaign.biddingStrategyName;
        budgetAmount = budgetAmount || campaign.budgetAmount;
        budgetPeriod = budgetPeriod || campaign.budgetPeriod;
        break;
      }
    }

    // Build periods object dynamically
    const periods = {};
    for (let i = 1; i <= numberOfPeriods; i++) {
      const periodKey = `period${i}`;
      const campaign = campaigns[i - 1];
      periods[periodKey] = campaign?.data || getEmptyPeriodData(dateRanges[periodKey]);
    }

    const combinedCampaign = {
      name: campaignName,
      biddingStrategyType: biddingStrategyType,
      biddingStrategyName: biddingStrategyName,
      budgetAmount: budgetAmount,
      budgetPeriod: budgetPeriod,
      periods: periods
    };

    combinedCampaigns.push(combinedCampaign);
  }

  console.log(`Combined data for ${combinedCampaigns.length} campaigns`);

  // Apply campaign limit if enabled
  let finalCampaigns = combinedCampaigns;
  if (LIMIT_CAMPAIGNS && MAX_CAMPAIGNS > 0) {
    finalCampaigns = combinedCampaigns.slice(0, MAX_CAMPAIGNS);
    console.log(`Campaign limit enabled: Processing ${finalCampaigns.length} of ${combinedCampaigns.length} campaigns (max: ${MAX_CAMPAIGNS})`);
  }

  // Log first few campaigns for verification
  if (finalCampaigns.length > 0) {
    console.log('');
    console.log('Sample combined campaign data (first 3 campaigns):');
    finalCampaigns.slice(0, 3).forEach((campaign, index) => {
      console.log(`${index + 1}. ${campaign.name}`);
      console.log(`   Bidding Strategy Type: ${campaign.biddingStrategyType}`);
      console.log(`   Bidding Strategy Name: ${campaign.biddingStrategyName}`);
      console.log('');

      for (let i = 1; i <= numberOfPeriods; i++) {
        const periodKey = `period${i}`;
        const period = campaign.periods[periodKey];
        console.log(`   Period ${i}:`);
        console.log(`     Date Range: ${period.dateRange}`);
        console.log(`     Sample Metrics - Cost: ${period.cost.toFixed(2)}, Clicks: ${period.clicks}, Impressions: ${period.impressions}`);
        console.log('');
      }
    });
  }

  return finalCampaigns;
}

function getEmptyPeriodData(dateRange) {
  return {
    impressions: 0,
    clicks: 0,
    cost: 0,
    conversions: 0,
    conversionsValue: 0,
    ctr: 0,
    cpc: 0,
    conversionRate: 0,
    costPerConversion: 0,
    searchImpressionShare: 0,
    searchTopImpressionShare: 0,
    searchAbsoluteTopImpressionShare: 0,
    dateRange: `${formatWeekDate(dateRange.startDate)} to ${formatWeekDate(dateRange.endDate)}`,
    startDate: dateRange.startDate,
    endDate: dateRange.endDate
  };
}



function formatWeekDate(weekString) {
  // weekString is in format YYYY-MM-DD (start of week)
  // Parse date in local timezone to avoid UTC conversion issues
  const [year, month, day] = weekString.split('-').map(Number);
  const date = new Date(year, month - 1, day);
  const options = { month: 'short', day: 'numeric', year: 'numeric' };
  return date.toLocaleDateString('en-US', options);
}

function getGoogleAdsApiFormattedDate(daysAgo = 0) {
  const date = new Date();
  date.setDate(date.getDate() - daysAgo);

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}

/**
 * Gets the Monday of the week for a given date
 * @param {Date} date - The date to find Monday for
 * @returns {Date} The Monday of that week
 */
function getMondayOfWeek(date) {
  const d = new Date(date);
  const dayOfWeek = d.getDay(); // 0 = Sunday, 1 = Monday, etc.

  // Calculate days to subtract to get to Monday
  let daysToSubtract;
  if (dayOfWeek === 0) {
    // Sunday: go back 6 days to Monday
    daysToSubtract = 6;
  } else if (dayOfWeek === 1) {
    // Monday: don't subtract anything
    daysToSubtract = 0;
  } else {
    // Tuesday-Saturday: subtract (dayOfWeek - 1) days
    daysToSubtract = dayOfWeek - 1;
  }

  d.setDate(d.getDate() - daysToSubtract);
  return d;
}

/**
 * Formats a Date object for Google Ads API (YYYY-MM-DD format)
 * @param {Date} date - The date to format
 * @returns {string} Date in YYYY-MM-DD format
 */
function formatDateForGoogleAds(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}

/**
 * Calculate week number of the year for a given date (Google Ads API standard - Monday-based weeks)
 * @param {string} dateString - Date string in YYYY-MM-DD format
 * @returns {number} Week number (1-53)
 */
function getWeekNumber(dateString) {
  // Parse date in local timezone to avoid UTC conversion issues
  const [year, month, day] = dateString.split('-').map(Number);
  const date = new Date(year, month - 1, day);

  // Get the Monday of this week - this ensures all days in the same week get the same week number
  const mondayOfWeek = getMondayOfWeek(date);

  // Use ISO 8601 week calculation based on the Monday
  // Set to nearest Thursday from Monday: Monday + 3 days
  const thursday = new Date(mondayOfWeek);
  thursday.setDate(thursday.getDate() + 3);

  // Get first day of year
  const yearStart = new Date(thursday.getFullYear(), 0, 1);

  // Calculate full weeks to nearest Thursday
  const weekNumber = Math.ceil((((thursday - yearStart) / 86400000) + 1) / 7);

  return weekNumber;
}

/**
 * Get month number range and formatted date range for a period
 * @param {Object} dateRange - Object with startDate and endDate
 * @returns {Object} Object with monthRange and formattedDateRange
 */
function getPeriodMonthInfo(dateRange) {
  // Parse dates in local timezone to avoid UTC conversion issues
  const [startYear, startMonth, startDay] = dateRange.startDate.split('-').map(Number);
  const [endYear, endMonth, endDay] = dateRange.endDate.split('-').map(Number);

  const startDate = new Date(startYear, startMonth - 1, startDay);
  const endDate = new Date(endYear, endMonth - 1, endDay);

  const formatOptions = { month: 'short', day: 'numeric', year: 'numeric' };
  const formattedStart = startDate.toLocaleDateString('en-US', formatOptions);
  const formattedEnd = endDate.toLocaleDateString('en-US', formatOptions);

  // Handle month range display
  let monthRange;
  if (startMonth === endMonth && startYear === endYear) {
    // Single month: use abbreviated format "Apr 2025"
    const startMonthAbbr = startDate.toLocaleDateString('en-US', { month: 'short' });
    monthRange = `${startMonthAbbr} ${startYear}`;
  } else {
    // Multiple months: use abbreviated format
    const startMonthAbbr = startDate.toLocaleDateString('en-US', { month: 'short' });
    const endMonthAbbr = endDate.toLocaleDateString('en-US', { month: 'short' });

    // Always use the same format: "Month Year - Month Year" or "Month Year" for single months
    if (startMonth === endMonth && startYear === endYear) {
      monthRange = `${startMonthAbbr} ${startYear}`;
    } else {
      monthRange = `${startMonthAbbr} ${startYear} - ${endMonthAbbr} ${endYear}`;
    }
  }

  // Return only the month range for the sheet header
  return {
    monthRange: monthRange,
    formattedDateRange: monthRange
  };
}

/**
 * Get week number range and formatted date range for a period
 * @param {Object} dateRange - Object with startDate and endDate
 * @returns {Object} Object with weekRange and formattedDateRange
 */
function getPeriodWeekInfo(dateRange) {
  const startWeek = getWeekNumber(dateRange.startDate);
  const endWeek = getWeekNumber(dateRange.endDate);

  // Parse dates in local timezone to avoid UTC conversion issues
  const [startYear, startMonth, startDay] = dateRange.startDate.split('-').map(Number);
  const [endYear, endMonth, endDay] = dateRange.endDate.split('-').map(Number);
  const startDate = new Date(startYear, startMonth - 1, startDay);
  const endDate = new Date(endYear, endMonth - 1, endDay);

  const formatOptions = { month: 'short', day: 'numeric', year: 'numeric' };
  const formattedStart = startDate.toLocaleDateString('en-US', formatOptions);
  const formattedEnd = endDate.toLocaleDateString('en-US', formatOptions);

  // Show week range if multiple weeks
  const weekRange = startWeek === endWeek ?
    `Week ${startWeek}` :
    `Week ${startWeek} - ${endWeek}`;

  return {
    weekRange: weekRange,
    formattedDateRange: `${formattedStart} - ${formattedEnd}`
  };
}

/**
 * Format budget amount and period for display
 * @param {number} budgetAmountMicros - Budget amount in micros
 * @param {string} budgetPeriod - Budget period (DAILY, MONTHLY, etc.)
 * @returns {string} Formatted budget string
 */
function formatBudget(budgetAmountMicros, budgetPeriod) {
  if (!budgetAmountMicros) return 'Budget: N/A';

  const budgetAmount = (budgetAmountMicros / 1000000).toFixed(2);

  // Convert daily budget to monthly estimate if needed
  if (budgetPeriod === 'DAILY') {
    const monthlyEstimate = (budgetAmount * 30.44).toFixed(2); // Average days per month
    return `Budget: $${monthlyEstimate}/month (est.)`;
  } else if (budgetPeriod === 'MONTHLY') {
    return `Budget: $${budgetAmount}/month`;
  } else {
    return `Budget: $${budgetAmount}/${budgetPeriod?.toLowerCase() || 'period'}`;
  }
}

function getAccountName() {
  const account = AdsApp.currentAccount();
  return account.getName().replace(/[^a-zA-Z0-9\s]/g, '').trim(); // Remove special characters for sheet name
}

function writeDataToSheet(campaignData, enabledMetrics, weeksInPeriod, numberOfPeriods, accountName, settingsSheet, periodType) {
  console.log(`Writing data to sheet for account: ${accountName}`);

  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const reportSheet = getOrCreateSheet(spreadsheet, accountName);

  // Calculate date ranges for header information
  const dateRanges = calculateDateRanges(weeksInPeriod, numberOfPeriods, periodType);

  // Clear existing data (but preserve targets in column B)
  clearSheetData(reportSheet);

  // Write headers
  writeHeaders(reportSheet, weeksInPeriod, numberOfPeriods, dateRanges, periodType);

  // Write campaign data
  let currentRow = 2; // Start at A2 since we now have only one header row

  campaignData.forEach((campaign, index) => {
    currentRow = writeCampaignSection(reportSheet, campaign, enabledMetrics, weeksInPeriod, numberOfPeriods, currentRow, settingsSheet);

    // Add 3-row gap between campaigns (except after the last one)
    if (index < campaignData.length - 1) {
      currentRow += 3;
    }
  });

  console.log(`Data written successfully to '${accountName}' sheet`);
}

function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    console.log(`Creating new sheet: ${sheetName}`);
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    console.log(`Using existing sheet: ${sheetName}`);
  }

  return sheet;
}

function clearSheetData(sheet) {
  // Get the last row and column with data
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow > 1 && lastColumn > 0) {
    // Clear everything except column B (targets) starting from row 2
    const rangesToClear = [
      `A2:A${lastRow}`, // Column A (metric names)
      `C2:${String.fromCharCode(67 + lastColumn - 3)}${lastRow}` // Columns C onwards (period data)
    ];

    rangesToClear.forEach(range => {
      try {
        sheet.getRange(range).clear();
      } catch (error) {
        // Range might not exist, continue
      }
    });

    // Clear formatting for all data areas
    if (lastColumn > 2) {
      sheet.getRange(`C2:${String.fromCharCode(67 + lastColumn - 3)}${lastRow}`).clearFormat();
    }
  }
}

function writeHeaders(sheet, weeksInPeriod, numberOfPeriods, dateRanges, periodType) {
  // Row 1: Combined headers with metric, target, and period info
  sheet.getRange('A1').setValue('Metric');
  sheet.getRange('B1').setValue(periodType === 'months' ? 'Monthly Target' : 'Weekly Target');

  // Add period info directly to row 1
  for (let i = 1; i <= numberOfPeriods; i++) {
    const column = String.fromCharCode(67 + i - 1); // C, D, E, F, etc.
    const periodKey = `period${i}`;
    if (periodType === 'months') {
      const periodInfo = getPeriodMonthInfo(dateRanges[periodKey]);
      const cell = sheet.getRange(`${column}1`);
      cell.setValue(periodInfo.monthRange);
      cell.setNote(`${formatWeekDate(dateRanges[periodKey].startDate)} - ${formatWeekDate(dateRanges[periodKey].endDate)}`);
    } else {
      const periodInfo = getPeriodWeekInfo(dateRanges[periodKey]);
      const cell = sheet.getRange(`${column}1`);
      cell.setValue(periodInfo.weekRange);
      cell.setNote(`${formatWeekDate(dateRanges[periodKey].startDate)} - ${formatWeekDate(dateRanges[periodKey].endDate)}`);
    }
  }

  // Format all headers (Row 1)
  const lastColumn = String.fromCharCode(67 + numberOfPeriods - 1);
  const headerRange = sheet.getRange(`A1:${lastColumn}1`);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e6f3ff');
  headerRange.setWrap(true);
  headerRange.setVerticalAlignment('middle');

  // Format period columns with smaller font for the week/date info
  const periodRange = sheet.getRange(`C1:${lastColumn}1`);
  periodRange.setFontSize(10);

  // Freeze only the header row (row 1)
  sheet.setFrozenRows(1);
}

function writeCampaignSection(sheet, campaign, enabledMetrics, weeksInPeriod, numberOfPeriods, startRow, settingsSheet) {
  let currentRow = startRow;

  // Bidding strategy header (separate cell above campaign name)
  const biddingStrategyText = campaign.biddingStrategyName ?
    `${campaign.biddingStrategyType} - ${campaign.biddingStrategyName}` :
    campaign.biddingStrategyType;
  sheet.getRange(currentRow, 1).setValue(biddingStrategyText);
  sheet.getRange(currentRow, 1).setFontWeight('bold');
  sheet.getRange(currentRow, 1).setBackground('#e0e0e0');
  currentRow++;

  // Campaign name header
  sheet.getRange(currentRow, 1).setValue(campaign.name);
  sheet.getRange(currentRow, 1).setFontWeight('bold');
  sheet.getRange(currentRow, 1).setBackground('#f0f0f0');

  // Budget information in column B
  const budgetText = formatBudget(campaign.budgetAmount, campaign.budgetPeriod);
  sheet.getRange(currentRow, 2).setValue(budgetText);
  sheet.getRange(currentRow, 2).setFontWeight('bold');
  sheet.getRange(currentRow, 2).setBackground('#f0f0f0');

  currentRow++;

  // Write metrics
  const metricDisplayNames = {
    impressions: 'Impressions',
    clicks: 'Clicks',
    cost: 'Cost',
    ctr: 'CTR (%)',
    cpc: 'CPC',
    conversions: 'Conversions',
    conversionRate: 'Conv. Rate (%)',
    costPerConversion: 'Cost/Conv.',
    searchImpressionShare: 'Search Impr. Share (%)',
    searchTopImpressionShare: 'Search Top IS (%)',
    searchAbsoluteTopImpressionShare: 'Search Abs. Top IS (%)'
  };

  enabledMetrics.forEach(metric => {
    const metricDisplayName = metricDisplayNames[metric];

    // Column A: Metric name
    sheet.getRange(currentRow, 1).setValue(metricDisplayName);

    // Column B: Monthly/Weekly target (preserve existing value if any)
    const existingTarget = sheet.getRange(currentRow, 2).getValue();
    if (!existingTarget) {
      sheet.getRange(currentRow, 2).setValue(''); // Placeholder for manual entry
    }

    // Dynamic period columns: C, D, E, F, etc.
    for (let i = 1; i <= numberOfPeriods; i++) {
      const column = 2 + i; // C=3, D=4, E=5, F=6, etc.
      const periodKey = `period${i}`;
      const periodData = campaign.periods[periodKey];
      const rawValue = periodData[metric];

      const cell = sheet.getRange(currentRow, column);
      cell.setValue(rawValue);

      // Apply number formatting based on metric type
      if (metric === 'cost' || metric === 'cpc' || metric === 'costPerConversion') {
        cell.setNumberFormat('#,##0.00'); // Currency without dollar sign
      } else if (metric === 'ctr' || metric === 'conversionRate' || metric === 'searchImpressionShare' || metric === 'searchTopImpressionShare' || metric === 'searchAbsoluteTopImpressionShare') {
        cell.setNumberFormat('0.00%'); // Percentage format
      } else if (metric === 'impressions' || metric === 'clicks' || metric === 'conversions') {
        cell.setNumberFormat('#,##0'); // Integer with commas
      } else {
        cell.setNumberFormat('#,##0.00'); // Default decimal format
      }

      // Apply color coding based on target comparison
      const target = sheet.getRange(currentRow, 2).getValue();
      if (target && typeof target === 'number') {
        // Only multiply by weeks for volume metrics, not rate metrics
        const volumeMetrics = ['impressions', 'clicks', 'cost', 'conversions'];
        const adjustedTarget = volumeMetrics.includes(metric) ? target * weeksInPeriod : target;
        const color = getColorForMetric(rawValue, adjustedTarget, metric);
        cell.setBackground(color);
      }
    }

    currentRow++;
  });

  return currentRow;
}

function formatMetricValue(value, metric) {
  if (value === 0) return 0;

  const currencyMetrics = ['cost', 'cpc', 'costPerConversion'];
  if (currencyMetrics.includes(metric)) {
    return value.toFixed(2);
  }

  const percentageMetrics = ['ctr', 'conversionRate', 'searchImpressionShare', 'searchTopImpressionShare', 'searchAbsoluteTopImpressionShare'];
  if (percentageMetrics.includes(metric)) {
    return `${(value * 100).toFixed(2)}%`;
  }

  const integerMetrics = ['impressions', 'clicks', 'conversions'];
  if (integerMetrics.includes(metric)) {
    return Math.round(value);
  }

  return value.toFixed(2);
}

function getColorForMetric(actualValue, targetValue, metric) {
  if (!targetValue || targetValue === 0) return '#ffffff'; // White if no target

  const ratio = actualValue / targetValue;

  // Define whether higher is better for each metric
  const higherIsBetter = ['impressions', 'clicks', 'ctr', 'conversions', 'conversionRate', 'searchImpressionShare', 'searchTopImpressionShare', 'searchAbsoluteTopImpressionShare'];
  const lowerIsBetter = ['cost', 'cpc', 'costPerConversion'];

  let performance;

  if (higherIsBetter.includes(metric)) {
    // For metrics where higher is better
    if (ratio >= 0.95) {
      performance = 'good';
    } else if (ratio >= 0.8) {
      performance = 'warning';
    } else {
      performance = 'poor';
    }
  } else if (lowerIsBetter.includes(metric)) {
    // For metrics where lower is better (costs)
    if (ratio <= 1.0) {
      performance = 'good';
    } else if (ratio <= 1.2) {
      performance = 'warning';
    } else {
      performance = 'poor';
    }
  } else {
    // Default case
    performance = 'neutral';
  }

  // Return color codes using guard clauses
  if (performance === 'good') return '#D3D3D3';    // Light grey
  if (performance === 'warning') return '#F4C430'; // Warm amber
  if (performance === 'poor') return '#F4A5A5';    // Warm coral red
  return '#ffffff';                                 // White
}

function debugLog(message) {
  if (DEBUG_MODE) {
    console.log(`[DEBUG] ${message}`);
  }
}

/**
 * Sends a beautiful HTML email report with campaign data
 * @param {Array} campaignData - The campaign data to include in the email
 * @param {Array} enabledMetrics - List of enabled metrics
 * @param {number} weeksInPeriod - Number of weeks in each period
 * @param {number} numberOfPeriods - Number of periods to display
 * @param {string} accountName - The account name
 * @param {string} emailAddresses - Comma-separated email addresses
 * @param {GoogleAppsScript.Spreadsheet.Sheet} settingsSheet - The settings sheet to read targets from
 */
function sendEmailReport(campaignData, enabledMetrics, weeksInPeriod, numberOfPeriods, accountName, emailAddresses, settingsSheet, periodType) {
  try {
    const emails = emailAddresses.split(',').map(email => email.trim()).filter(email => email.length > 0);

    if (emails.length === 0) {
      console.warn('No valid email addresses found');
      return;
    }

    const dateRanges = calculateDateRanges(weeksInPeriod, numberOfPeriods, periodType);

    // Get targets from the report sheet that was just written
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    const reportSheet = spreadsheet.getSheetByName(accountName);
    const targets = getTargetsFromReportSheet(reportSheet, campaignData, enabledMetrics);

    const htmlContent = generateEmailHtml(campaignData, enabledMetrics, numberOfPeriods, accountName, dateRanges, targets, weeksInPeriod, periodType);

    const subject = `📊 Weekly Campaign Performance Report - ${accountName}`;

    console.log(`Sending email to: ${emails.join(', ')}`);

    emails.forEach(email => {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: htmlContent
      });
    });

    console.log('Email report sent successfully');

  } catch (error) {
    console.error('Error sending email report:', error.message);
  }
}

/**
 * Reads targets from the report sheet for email inclusion
 * @param {GoogleAppsScript.Spreadsheet.Sheet} reportSheet - The report sheet to read targets from
 * @param {Array} campaignData - The campaign data
 * @param {Array} enabledMetrics - List of enabled metrics
 * @returns {Object} Object mapping campaign names and metrics to their targets
 */
function getTargetsFromReportSheet(reportSheet, campaignData, enabledMetrics) {
  const targets = {};

  if (!reportSheet) {
    console.warn('Report sheet not found, targets will not be included in email');
    return targets;
  }

  try {
    let currentRow = 3; // Start at row 3 where data begins

    campaignData.forEach(campaign => {
      currentRow++; // Skip bidding strategy row
      currentRow++; // Skip campaign name row

      const campaignTargets = {};

      enabledMetrics.forEach(metric => {
        try {
          const targetValue = reportSheet.getRange(currentRow, 2).getValue(); // Column B
          if (targetValue && typeof targetValue === 'number') {
            campaignTargets[metric] = targetValue;
          }
        } catch (error) {
          // Skip if error reading target
        }
        currentRow++;
      });

      if (Object.keys(campaignTargets).length > 0) {
        targets[campaign.name] = campaignTargets;
      }

      currentRow += 3; // Skip gap between campaigns
    });

    console.log(`Loaded targets for ${Object.keys(targets).length} campaigns`);
  } catch (error) {
    console.warn('Error reading targets from report sheet:', error.message);
  }

  return targets;
}

/**
 * Generates beautiful HTML content for the email report
 * @param {Array} campaignData - The campaign data
 * @param {Array} enabledMetrics - List of enabled metrics
 * @param {number} numberOfPeriods - Number of periods
 * @param {string} accountName - The account name
 * @param {Object} dateRanges - Date ranges for each period
 * @param {Object} targets - Target values for campaigns and metrics
 * @param {number} weeksInPeriod - Number of weeks in each period
 * @returns {string} HTML content for the email
 */
function generateEmailHtml(campaignData, enabledMetrics, numberOfPeriods, accountName, dateRanges, targets, weeksInPeriod, periodType) {
  const currentDate = new Date().toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });

  // Create period headers
  let periodHeaders = '';
  for (let i = 1; i <= numberOfPeriods; i++) {
    const periodKey = `period${i}`;
    if (periodType === 'months') {
      const periodInfo = getPeriodMonthInfo(dateRanges[periodKey]);
      periodHeaders += `
      <th style="background: linear-gradient(135deg, #ff8c42, #ff6b1a); color: white; padding: 15px 10px; text-align: center; font-weight: 600; border-radius: 8px 8px 0 0; box-shadow: 0 2px 4px rgba(255,140,66,0.3); min-width: 140px;">
        <div style="font-size: 14px; margin-bottom: 4px;">Period ${i}</div>
        <div style="font-size: 11px; opacity: 0.9; font-weight: 400;">${periodInfo.monthRange}</div>
        <div style="font-size: 10px; opacity: 0.8; font-weight: 400;">${periodInfo.formattedDateRange}</div>
      </th>`;
    } else {
      const periodInfo = getPeriodWeekInfo(dateRanges[periodKey]);
      periodHeaders += `
      <th style="background: linear-gradient(135deg, #ff8c42, #ff6b1a); color: white; padding: 15px 10px; text-align: center; font-weight: 600; border-radius: 8px 8px 0 0; box-shadow: 0 2px 4px rgba(255,140,66,0.3); min-width: 140px;">
        <div style="font-size: 14px; margin-bottom: 4px;">Period ${i}</div>
        <div style="font-size: 11px; opacity: 0.9; font-weight: 400;">${periodInfo.weekRange}</div>
        <div style="font-size: 10px; opacity: 0.8; font-weight: 400;">${periodInfo.formattedDateRange}</div>
      </th>`;
    }
  }

  // Generate campaign rows
  let campaignRows = '';
  campaignData.forEach((campaign, index) => {
    campaignRows += generateCampaignEmailRows(campaign, enabledMetrics, numberOfPeriods, index, targets, weeksInPeriod, periodType);

    // Add spacing row between campaigns (except after the last one)
    if (index < campaignData.length - 1) {
      campaignRows += `<tr><td colspan="${numberOfPeriods + 2}" style="height: 12px; border: none;"></td></tr>`;
    }
  });

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Campaign Performance Report</title>
      <style>
        @media only screen and (max-width: 600px) {
          .email-container {
            margin-left: 5px !important;
            margin-right: 5px !important;
            border-radius: 8px !important;
          }
          .content-padding {
            padding: 15px !important;
          }
          .header-padding {
            padding: 20px 15px !important;
          }
          .table-scroll {
            overflow-x: auto !important;
            -webkit-overflow-scrolling: touch !important;
          }
          .table-responsive {
            min-width: 700px !important;
          }
        }
      </style>
    </head>
    <body style="margin: 0; padding: 0; font-family: Arial, Helvetica, sans-serif; background: linear-gradient(135deg, #f8f9fa, #e9ecef); min-height: 100vh;">
      
      <!-- Email Container -->
      <div class="email-container" style="max-width: 1200px; margin: 0 auto; background: white; box-shadow: 0 8px 32px rgba(0,0,0,0.1); border-radius: 16px; overflow: hidden; margin-top: 20px; margin-bottom: 20px; margin-left: 10px; margin-right: 10px;">
        
        <!-- Header -->
        <div class="header-padding" style="background: linear-gradient(135deg, #ff8c42, #ff6b1a); padding: 30px; text-align: center; position: relative; overflow: hidden;">
          <div style="position: absolute; top: -50px; right: -50px; width: 100px; height: 100px; background: rgba(255,255,255,0.1); border-radius: 50%; opacity: 0.6;"></div>
          <div style="position: absolute; bottom: -30px; left: -30px; width: 60px; height: 60px; background: rgba(255,255,255,0.1); border-radius: 50%; opacity: 0.4;"></div>
          
          <h1 style="color: white; margin: 0; font-size: 28px; font-weight: 700; text-shadow: 0 2px 4px rgba(0,0,0,0.2); position: relative; z-index: 1;">
            📊 Campaign Performance Report
          </h1>
          <p style="color: rgba(255,255,255,0.9); margin: 8px 0 0 0; font-size: 16px; font-weight: 400; position: relative; z-index: 1;">
            ${accountName} • ${currentDate}
          </p>
        </div>

        <!-- Content -->
        <div class="content-padding" style="padding: 30px;">
          
          <!-- Summary Stats -->
          <div style="background: linear-gradient(135deg, #fff7ed, #fed7aa); border: 1px solid #fb923c; border-radius: 12px; padding: 20px; margin-bottom: 30px; position: relative; overflow: hidden;">
            <div style="position: absolute; top: 0; right: 0; width: 80px; height: 80px; background: linear-gradient(135deg, rgba(255,140,66,0.1), rgba(255,107,26,0.05)); border-radius: 0 12px 0 80px;"></div>
            <h2 style="color: #c2410c; margin: 0 0 15px 0; font-size: 18px; font-weight: 600; position: relative; z-index: 1;">
              📈 Report Summary
            </h2>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; position: relative; z-index: 1;">
              <div style="text-align: center;">
                <div style="font-size: 24px; font-weight: 700; color: #c2410c;">${campaignData.length}</div>
                <div style="font-size: 12px; color: #92400e; font-weight: 500;">Active Campaigns</div>
              </div>
              <div style="text-align: center;">
                <div style="font-size: 24px; font-weight: 700; color: #c2410c;">${enabledMetrics.length}</div>
                <div style="font-size: 12px; color: #92400e; font-weight: 500;">Tracked Metrics</div>
              </div>
              <div style="text-align: center;">
                <div style="font-size: 24px; font-weight: 700; color: #c2410c;">${numberOfPeriods}</div>
                <div style="font-size: 12px; color: #92400e; font-weight: 500;">Time Periods</div>
              </div>
            </div>
          </div>

          <!-- Performance Table -->
          <div style="background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 16px rgba(0,0,0,0.1); border: 1px solid #e5e7eb;">
            <div class="table-scroll" style="overflow-x: auto; -webkit-overflow-scrolling: touch;">
              <table class="table-responsive" style="min-width: 800px; width: 100%; border-collapse: collapse; font-size: 13px;">
                <thead>
                  <tr>
                    <th style="background: linear-gradient(135deg, #374151, #1f2937); color: white; padding: 15px 12px; text-align: left; font-weight: 600; position: sticky; top: 0; z-index: 10; min-width: 200px;">
                      Campaign & Metrics
                    </th>
                    <th style="background: linear-gradient(135deg, #374151, #1f2937); color: white; padding: 15px 12px; text-align: center; font-weight: 600; position: sticky; top: 0; z-index: 10; min-width: 120px;">
                      ${periodType === 'months' ? 'Monthly Target' : 'Weekly Target'}
                    </th>
                    ${periodHeaders}
                  </tr>
                </thead>
                <tbody>
                  ${campaignRows}
                </tbody>
              </table>
            </div>
          </div>

          <!-- Footer -->
          <div style="margin-top: 30px; padding: 20px; background: #f8f9fa; border-radius: 12px; text-align: center; border: 1px solid #e5e7eb;">
            <p style="margin: 0; color: #6b7280; font-size: 12px; font-weight: 500;">
              🚀 Generated by Google Ads Performance Script • ${currentDate}
            </p>
            <p style="margin: 8px 0 0 0; color: #9ca3af; font-size: 11px;">
              Need help? This automated report tracks your campaign performance across multiple time periods.
            </p>
          </div>

        </div>
      </div>
    </body>
    </html>
  `;
}

/**
 * Generates HTML rows for a single campaign in the email
 * @param {Object} campaign - Campaign data
 * @param {Array} enabledMetrics - List of enabled metrics
 * @param {number} numberOfPeriods - Number of periods
 * @param {number} campaignIndex - Index of the campaign (for styling)
 * @param {Object} targets - Target values for campaigns and metrics
 * @param {number} weeksInPeriod - Number of weeks in each period
 * @returns {string} HTML rows for the campaign
 */
function generateCampaignEmailRows(campaign, enabledMetrics, numberOfPeriods, campaignIndex, targets, weeksInPeriod, periodType) {
  const isEven = campaignIndex % 2 === 0;
  const campaignBg = isEven ? '#fafafa' : '#ffffff';

  let rows = '';

  // Campaign header row
  const budgetText = formatBudget(campaign.budgetAmount, campaign.budgetPeriod);
  const biddingStrategyText = campaign.biddingStrategyName ?
    `${campaign.biddingStrategyType} - ${campaign.biddingStrategyName}` :
    campaign.biddingStrategyType;

  rows += `
    <tr style="background: ${campaignBg};">
      <td colspan="${numberOfPeriods + 2}" style="padding: 16px 12px 8px 12px; border-bottom: 1px solid #e5e7eb;">
        <div style="font-weight: 700; font-size: 14px; color: #1f2937; margin-bottom: 4px;">
          ${campaign.name}
        </div>
        <div style="font-size: 11px; color: #6b7280; margin-bottom: 2px;">
          📊 ${biddingStrategyText}
        </div>
        <div style="font-size: 11px; color: #6b7280;">
          💰 ${budgetText}
        </div>
      </td>
    </tr>
  `;

  // Metric rows
  const metricDisplayNames = {
    impressions: 'Impressions',
    clicks: 'Clicks',
    cost: 'Cost',
    ctr: 'CTR (%)',
    cpc: 'CPC',
    conversions: 'Conversions',
    conversionRate: 'Conv. Rate (%)',
    costPerConversion: 'Cost/Conv.',
    searchImpressionShare: 'Search Impr. Share (%)',
    searchTopImpressionShare: 'Search Top IS (%)',
    searchAbsoluteTopImpressionShare: 'Search Abs. Top IS (%)'
  };

  enabledMetrics.forEach(metric => {
    const metricDisplayName = metricDisplayNames[metric];
    const campaignTargets = targets[campaign.name] || {};
    const target = campaignTargets[metric];

    // Format target value
    let targetDisplay = '';
    if (target && typeof target === 'number') {
      targetDisplay = formatMetricValueForEmail(target, metric);
    } else {
      targetDisplay = '-';
    }

    let periodCells = '';
    for (let i = 1; i <= numberOfPeriods; i++) {
      const periodKey = `period${i}`;
      const periodData = campaign.periods[periodKey];
      const rawValue = periodData[metric];
      const formattedValue = formatMetricValueForEmail(rawValue, metric);

      // Apply color coding based on target comparison
      let backgroundColor = campaignBg;
      if (target && typeof target === 'number') {
        // Calculate adjusted target (multiply by weeks for volume metrics)
        const volumeMetrics = ['impressions', 'clicks', 'cost', 'conversions'];
        const adjustedTarget = volumeMetrics.includes(metric) ? target * weeksInPeriod : target;
        const color = getColorForMetric(rawValue, adjustedTarget, metric);
        backgroundColor = color;
      }

      periodCells += `
        <td style="padding: 8px 12px; text-align: center; border-bottom: 1px solid #f3f4f6; background: ${backgroundColor}; font-weight: 500; color: #374151;">
          ${formattedValue}
        </td>
      `;
    }

    rows += `
      <tr style="background: ${campaignBg};">
        <td style="padding: 8px 12px; border-bottom: 1px solid #f3f4f6; font-weight: 500; color: #4b5563; background: ${campaignBg};">
          ${metricDisplayName}
        </td>
        <td style="padding: 8px 12px; text-align: center; border-bottom: 1px solid #f3f4f6; background: ${campaignBg}; font-weight: 500; color: #374151;">
          ${targetDisplay}
        </td>
        ${periodCells}
      </tr>
    `;
  });

  // Add spacing row between campaigns (will be handled by the calling function)

  return rows;
}

/**
 * Formats metric values for email display
 * @param {number} value - The metric value
 * @param {string} metric - The metric type
 * @returns {string} Formatted value
 */
function formatMetricValueForEmail(value, metric) {
  if (value === 0) return '0';

  const currencyMetrics = ['cost', 'cpc', 'costPerConversion'];
  if (currencyMetrics.includes(metric)) {
    return '$' + value.toFixed(2);
  }

  const percentageMetrics = ['ctr', 'conversionRate', 'searchImpressionShare', 'searchTopImpressionShare', 'searchAbsoluteTopImpressionShare'];
  if (percentageMetrics.includes(metric)) {
    return (value * 100).toFixed(2) + '%';
  }

  const integerMetrics = ['impressions', 'clicks', 'conversions'];
  if (integerMetrics.includes(metric)) {
    return Math.round(value).toLocaleString();
  }

  return value.toFixed(2);
}



function isMCC() {
  try {
    MccApp.accounts();
    return true;
  } catch (e) {
    if (String(e).indexOf("not defined") > -1) {
      return false;
    } else {
      return true;
    }
  }
}
