/**
 * GCLID Keyword Report
 * Reads GCLIDs from a spreadsheet and enriches them with keyword, ad group, and campaign information
 * Version 1.0.4
 */

// Query Builder Links:
// Click View: https://developers.google.com/google-ads/api/fields/v20/click_view_query_builder

// ===== CONFIGURATION =====
const SPREADSHEET_URL = 'PASTE_YOUR_SPREADSHEET_URL_HERE';
// The Google Sheets URL containing GCLIDs in column A

const SHEET_NAME = 'GCLID Data';
// Name of the sheet to read from and write to

const DEBUG_MODE = true;
// Set to true to enable detailed logging for debugging

// ===== MAIN FUNCTION =====
function main() {
  console.log(`GCLID Keyword Report script started`);

  validateConfiguration();

  const sheet = getOrCreateSheet();
  const gclids = readGclidsFromSheet(sheet);

  if (!gclids || gclids.length === 0) {
    console.log(`No GCLIDs found in the sheet. Please add GCLIDs to column A.`);
    return;
  }

  console.log(`Found ${gclids.length} GCLIDs to process`);

  const enrichedData = enrichGclidsWithKeywordData(gclids);
  writeResultsToSheet(sheet, enrichedData);

  console.log(`GCLID Keyword Report script finished`);
}

// ===== VALIDATION FUNCTIONS =====
function validateConfiguration() {
  if (SPREADSHEET_URL === 'PASTE_YOUR_SPREADSHEET_URL_HERE') {
    throw new Error('Please update SPREADSHEET_URL with your actual spreadsheet URL');
  }
}

// ===== SHEET FUNCTIONS =====
function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    console.log(`Created new sheet: ${SHEET_NAME}`);
  }

  // Set up headers if sheet is empty or has different headers
  setupSheetHeaders(sheet);

  return sheet;
}

function setupSheetHeaders(sheet) {
  const headers = [
    'GCLID',
    'Date',
    'Campaign ID',
    'Campaign Name',
    'Ad Group ID',
    'Ad Group Name',
    'Keyword ID',
    'Keyword Text',
    'Match Type',
    'Ad ID',
    'Ad Type',
    'Search Partner'
  ];

  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const headersMatch = headers.every((header, index) => currentHeaders[index] === header);

  if (!headersMatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    console.log(`Set up sheet headers`);
  }
}

function readGclidsFromSheet(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    debugLog(`No data found in sheet beyond headers`);
    return [];
  }

  const gclidsRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const gclidsValues = gclidsRange.getValues();

  const gclids = gclidsValues
    .map(row => row[0])
    .filter(gclid => gclid && gclid.toString().trim() !== '');

  debugLog(`Read ${gclids.length} GCLIDs from sheet`);
  if (gclids.length > 0) {
    debugLog(`First 3 GCLIDs: ${gclids.slice(0, 3).join(', ')}`);
  }

  return gclids;
}

function writeResultsToSheet(sheet, enrichedData) {
  if (!enrichedData || enrichedData.length === 0) {
    console.log(`No data to write to sheet`);
    return;
  }

  // Clear existing data (except headers)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 12).clearContent();
  }

  // Convert enriched data to sheet format
  const sheetData = enrichedData.map(data => [
    data.gclid,
    data.date || '',
    data.campaignId || '',
    data.campaignName || '',
    data.adGroupId || '',
    data.adGroupName || '',
    data.keywordId || '',
    data.keywordText || '',
    data.matchType || '',
    data.adId || '',
    data.adType || '',
    data.searchPartner || ''
  ]);

  // Write data to sheet
  if (sheetData.length > 0) {
    sheet.getRange(2, 1, sheetData.length, 12).setValues(sheetData);
    console.log(`Wrote ${sheetData.length} rows of enriched GCLID data to sheet`);
  }
}

// ===== DATA PROCESSING FUNCTIONS =====
function enrichGclidsWithKeywordData(gclids) {
  const enrichedData = [];

  for (let gclid of gclids) {
    debugLog(`Processing GCLID: ${gclid}`);

    try {
      const keywordData = getKeywordDataForGclid(gclid);

      if (keywordData) {
        enrichedData.push({
          gclid: gclid,
          ...keywordData
        });
        debugLog(`Successfully enriched GCLID: ${gclid}`);
      } else {
        // Add GCLID with empty data if no match found
        enrichedData.push({
          gclid: gclid,
          date: '',
          campaignId: '',
          campaignName: 'No data found',
          adGroupId: '',
          adGroupName: '',
          keywordId: '',
          keywordText: '',
          matchType: '',
          adId: '',
          adType: '',
          searchPartner: ''
        });
        console.log(`No keyword data found for GCLID: ${gclid}`);
      }
    } catch (error) {
      console.error(`Error processing GCLID ${gclid}: ${error.message}`);

      // Add GCLID with error indication
      enrichedData.push({
        gclid: gclid,
        date: '',
        campaignId: '',
        campaignName: 'Error occurred',
        adGroupId: '',
        adGroupName: '',
        keywordId: '',
        keywordText: '',
        matchType: '',
        adId: '',
        adType: '',
        searchPartner: ''
      });
    }

    // Add small delay to avoid rate limiting
    Utilities.sleep(100);
  }

  console.log(`Processed ${enrichedData.length} GCLIDs total`);

  // Log first few results for verification
  if (enrichedData.length > 0) {
    debugLog(`First 3 enriched results:`);
    enrichedData.slice(0, 3).forEach((data, index) => {
      debugLog(`${index + 1}. GCLID: ${data.gclid}, Campaign: ${data.campaignName}, Keyword: ${data.keywordText}`);
    });
  }

  return enrichedData;
}

function getKeywordDataForGclid(gclid) {
  // Try each day for the last 90 days since click_view requires single day queries
  for (let daysAgo = 0; daysAgo < 90; daysAgo++) {
    const query = getClickViewGaqlQuery(gclid, daysAgo);
    const report = getClickViewReport(query);

    if (report) {
      const keywordData = extractKeywordDataFromReport(report);
      if (keywordData) {
        debugLog(`Found data for GCLID ${gclid} on day ${daysAgo} days ago`);
        return keywordData;
      }
    }

    // Small delay between requests to avoid rate limiting
    Utilities.sleep(50);
  }

  debugLog(`No data found for GCLID ${gclid} in the last 90 days`);
  return null;
}

// ===== GAQL QUERY FUNCTIONS =====
function getClickViewGaqlQuery(gclid, daysAgo) {
  const targetDate = getGoogleAdsApiFormattedDate(daysAgo);

  const query = `
        SELECT
          click_view.gclid,
          segments.date,
          campaign.id,
          campaign.name,
          ad_group.id,
          ad_group.name,
          click_view.keyword_info.text,
          click_view.keyword_info.match_type
        FROM click_view
        WHERE click_view.gclid = '${gclid}'
          AND segments.date = '${targetDate}'
        LIMIT 1
    `;

  debugLog(`GAQL Query for GCLID ${gclid} on ${targetDate}: ${query.trim()}`);

  return query;
}

function getClickViewReport(query) {
  try {
    const report = AdsApp.report(query);
    const rows = report.rows();

    debugLog(`Number of rows from click view query: ${rows.totalNumEntities()}`);

    if (!rows.hasNext()) {
      debugLog(`No click view data found for this GCLID`);
      return null;
    }

    return report;
  } catch (error) {
    console.error(`Error executing click view query: ${error.message}`);
    console.error(`If you get GAQL errors, validate your query at: https://developers.google.com/google-ads/api/fields/v19/query_validator`);
    console.error(`You can also use the query builder at: https://developers.google.com/google-ads/api/fields/v20/click_view_query_builder`);

    throw error;
  }
}

function extractKeywordDataFromReport(report) {
  const rows = report.rows();

  if (!rows.hasNext()) {
    return null;
  }

  const row = rows.next();

  const keywordData = {
    date: row['segments.date'] || '',
    campaignId: row['campaign.id'] || '',
    campaignName: row['campaign.name'] || '',
    adGroupId: row['ad_group.id'] || '',
    adGroupName: row['ad_group.name'] || '',
    keywordId: '', // Not available in click_view
    keywordText: row['click_view.keyword_info.text'] || '',
    matchType: row['click_view.keyword_info.match_type'] || '',
    adId: '', // Not available in click_view
    adType: '', // Not available in click_view
    searchPartner: 'Not available in click_view' // Network type not available in click_view API
  };

  debugLog(`Extracted keyword data: Campaign=${keywordData.campaignName}, Keyword=${keywordData.keywordText}`);

  return keywordData;
}

// ===== UTILITY FUNCTIONS =====
function getGoogleAdsApiFormattedDate(daysAgo = 0) {
  const date = new Date();
  date.setDate(date.getDate() - daysAgo);

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}
function debugLog(message) {
  if (DEBUG_MODE) {
    console.log(`[DEBUG] ${message}`);
  }
}
